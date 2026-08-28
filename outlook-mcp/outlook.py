"""Wrapper around the Outlook 2016 COM Object Model.

Read-only: list_folders, list_recent_emails, search_emails, get_email.
Writes (create_folder, move_emails) are isolated at the bottom of this file
and move_emails refuses to run without an explicit confirm=True - callers
are expected to run preview_move first.

Must run on Windows, in the same desktop session as Outlook 2016, with
Outlook already open at least once (COM will otherwise fail to attach).
"""

from datetime import datetime

import pythoncom
import win32com.client

OL_MAIL_ITEM_CLASS = 43  # olMail

# Default-folder aliases, resolved via Namespace/Store.GetDefaultFolder so
# they work regardless of the Outlook UI language (e.g. "收件匣" vs "Inbox").
DEFAULT_FOLDER_ALIASES = {
    "inbox": 6,          # olFolderInbox
    "sent items": 5,     # olFolderSentMail
    "sent": 5,
    "deleted items": 3,  # olFolderDeletedItems
    "trash": 3,
    "drafts": 16,        # olFolderDrafts
    "outbox": 4,         # olFolderOutbox
}


class OutlookError(RuntimeError):
    pass


def _connect():
    """Attach to the local Outlook 2016 application via COM.

    Must be called from the same OS thread that uses the returned objects -
    COM apartments are thread-bound (see CoInitialize calls in callers).
    """
    try:
        outlook_app = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook_app.GetNamespace("MAPI")
    except Exception as exc:  # pragma: no cover - environment dependent
        raise OutlookError(
            "Could not connect to Outlook via COM. Make sure Outlook 2016 is "
            "installed, has been opened at least once in this Windows "
            "session, and that this script runs on the same machine/user "
            "session (not a remote/service context)."
        ) from exc
    return namespace


def _walk_folders(folder, path_prefix=""):
    path = f"{path_prefix}/{folder.Name}" if path_prefix else folder.Name
    yield path, folder
    for sub in folder.Folders:
        yield from _walk_folders(sub, path)


def list_folders():
    """Return every folder path across all connected stores (accounts/PSTs)."""
    pythoncom.CoInitialize()
    try:
        namespace = _connect()
        results = []
        for store in namespace.Stores:
            try:
                root = store.GetRootFolder()
            except Exception:
                continue
            for path, folder in _walk_folders(root):
                try:
                    item_count = folder.Items.Count
                except Exception:
                    item_count = None
                results.append(
                    {
                        "store": store.DisplayName,
                        "path": path,
                        "item_count": item_count,
                    }
                )
        return results
    finally:
        pythoncom.CoUninitialize()


def _find_folder(namespace, folder_path, store_name=None):
    """Resolve a '/'-separated folder path (e.g. 'Inbox/Projects') to a Folder object.

    The first path segment may be a language-independent alias (see
    DEFAULT_FOLDER_ALIASES, e.g. "Inbox" resolves correctly even when the
    Outlook UI shows it as "收件匣"). Remaining segments are matched by the
    folder's literal (localized) display name.
    """
    parts = [p for p in folder_path.split("/") if p]
    if not parts:
        return None

    # No specific store requested and the first segment is a default-folder
    # alias: use the namespace-level default (the account Outlook itself
    # treats as primary), rather than guessing which Store to search first -
    # a secondary/archive PST can otherwise shadow the real Inbox.
    alias = parts[0].strip().lower()
    if store_name is None and alias in DEFAULT_FOLDER_ALIASES:
        try:
            node = namespace.GetDefaultFolder(DEFAULT_FOLDER_ALIASES[alias])
        except Exception:
            node = None
        if node is not None:
            matched = True
            for part in parts[1:]:
                found = None
                for sub in node.Folders:
                    if sub.Name == part:
                        found = sub
                        break
                if found is None:
                    matched = False
                    break
                node = found
            if matched:
                return node

    for store in namespace.Stores:
        if store_name and store.DisplayName != store_name:
            continue

        node = None
        remaining = parts
        alias = parts[0].strip().lower()
        if alias in DEFAULT_FOLDER_ALIASES:
            try:
                node = store.GetDefaultFolder(DEFAULT_FOLDER_ALIASES[alias])
            except Exception:
                node = None
            remaining = parts[1:]

        if node is None:
            try:
                node = store.GetRootFolder()
            except Exception:
                continue
            remaining = parts

        matched = True
        for part in remaining:
            found = None
            for sub in node.Folders:
                if sub.Name == part:
                    found = sub
                    break
            if found is None:
                matched = False
                break
            node = found
        if matched:
            return node
    return None


def _safe_sender_email(item):
    try:
        if item.SenderEmailType == "EX":
            return item.Sender.GetExchangeUser().PrimarySmtpAddress
    except Exception:
        pass
    return getattr(item, "SenderEmailAddress", None)


def list_recent_emails(folder_path="Inbox", count=10, store_name=None):
    """Return metadata (NOT body) for the most recent `count` emails in a folder."""
    pythoncom.CoInitialize()
    try:
        namespace = _connect()
        folder = _find_folder(namespace, folder_path, store_name)
        if folder is None:
            raise OutlookError(f"Folder not found: {folder_path!r}")

        items = folder.Items
        items.Sort("[ReceivedTime]", True)  # newest first

        results = []
        for item in items:
            if len(results) >= count:
                break
            if getattr(item, "Class", None) != OL_MAIL_ITEM_CLASS:
                continue
            try:
                results.append(
                    {
                        "entry_id": item.EntryID,
                        "subject": item.Subject,
                        "sender_name": item.SenderName,
                        "sender_email": _safe_sender_email(item),
                        "received_time": str(item.ReceivedTime),
                        "unread": item.UnRead,
                        "has_attachments": item.Attachments.Count > 0,
                        "size_bytes": item.Size,
                    }
                )
            except Exception:
                continue
        return results
    finally:
        pythoncom.CoUninitialize()


def _parse_date(s):
    if not s:
        return None
    return datetime.strptime(s, "%Y-%m-%d").date()


def search_emails(
    folder_path="Inbox",
    date_from=None,
    date_to=None,
    subject_contains=None,
    sender_contains=None,
    max_results=500,
    store_name=None,
):
    """Search a single folder's mail items by date range / subject / sender keywords.

    Read-only, metadata only (no body) - same shape as list_recent_emails.
    Cheap: Items is sorted newest-first, so we stop scanning as soon as we
    pass date_from instead of walking the whole folder.
    """
    pythoncom.CoInitialize()
    try:
        namespace = _connect()
        folder = _find_folder(namespace, folder_path, store_name)
        if folder is None:
            raise OutlookError(f"Folder not found: {folder_path!r}")

        items = folder.Items
        items.Sort("[ReceivedTime]", True)  # newest first

        d_from = _parse_date(date_from)
        d_to = _parse_date(date_to)
        subject_kw = [s.lower() for s in (subject_contains or [])]
        sender_kw = [s.lower() for s in (sender_contains or [])]

        results = []
        for item in items:
            if len(results) >= max_results:
                break
            if getattr(item, "Class", None) != OL_MAIL_ITEM_CLASS:
                continue
            try:
                received_date = item.ReceivedTime.date()
            except Exception:
                continue
            if d_to and received_date > d_to:
                continue
            if d_from and received_date < d_from:
                break  # sorted descending: everything from here on is even older
            try:
                subject = item.Subject or ""
                sender_email = _safe_sender_email(item) or ""
                sender_name = item.SenderName or ""
            except Exception:
                continue
            if subject_kw and not any(k in subject.lower() for k in subject_kw):
                continue
            if sender_kw and not any(
                k in sender_email.lower() or k in sender_name.lower() for k in sender_kw
            ):
                continue
            try:
                results.append(
                    {
                        "entry_id": item.EntryID,
                        "subject": subject,
                        "sender_name": sender_name,
                        "sender_email": sender_email,
                        "received_time": str(item.ReceivedTime),
                        "unread": item.UnRead,
                        "has_attachments": item.Attachments.Count > 0,
                        "size_bytes": item.Size,
                    }
                )
            except Exception:
                continue
        return results
    finally:
        pythoncom.CoUninitialize()


def get_email(entry_id):
    """Return full content (body + attachment names, NOT attachment content) for one email."""
    pythoncom.CoInitialize()
    try:
        namespace = _connect()
        item = namespace.GetItemFromID(entry_id)
        if item is None or getattr(item, "Class", None) != OL_MAIL_ITEM_CLASS:
            raise OutlookError(f"Not a mail item or not found: {entry_id!r}")
        attachments = []
        for att in item.Attachments:
            try:
                attachments.append({"filename": att.FileName, "size_bytes": att.Size})
            except Exception:
                continue
        return {
            "entry_id": item.EntryID,
            "subject": item.Subject,
            "sender_name": item.SenderName,
            "sender_email": _safe_sender_email(item),
            "to": item.To,
            "cc": item.CC,
            "received_time": str(item.ReceivedTime),
            "body": item.Body,
            "attachments": attachments,
            "current_folder": item.Parent.Name,
        }
    finally:
        pythoncom.CoUninitialize()


# ---------------------------------------------------------------------------
# Write operations. move_emails refuses to run without confirm=True - the
# expected flow is: search_emails/outlook_classify_recent (read-only) ->
# preview_move (read-only, shows exactly what would happen) -> user agrees ->
# move_emails(..., confirm=True).
# ---------------------------------------------------------------------------


def _default_store(namespace):
    return namespace.GetDefaultFolder(DEFAULT_FOLDER_ALIASES["inbox"]).Store


def _resolve_store(namespace, store_name=None):
    if store_name is None:
        return _default_store(namespace)
    for store in namespace.Stores:
        if store.DisplayName == store_name:
            return store
    raise OutlookError(f"Store not found: {store_name!r}")


def _resolve_destination_folder(namespace, path, store_name=None, create_if_missing=True):
    store = _resolve_store(namespace, store_name)
    node = store.GetRootFolder()
    parts = [p for p in path.split("/") if p]
    if not parts:
        raise OutlookError("destination_path must not be empty")
    for part in parts:
        found = None
        for sub in node.Folders:
            if sub.Name == part:
                found = sub
                break
        if found is None:
            if not create_if_missing:
                raise OutlookError(f"Folder not found: {path!r} (missing segment {part!r})")
            found = node.Folders.Add(part)
        node = found
    return node


def create_folder(path, store_name=None):
    """Create a folder (and any missing parent folders) at a '/'-separated path.

    Idempotent - safe to call again if the folder already exists.
    """
    pythoncom.CoInitialize()
    try:
        namespace = _connect()
        folder = _resolve_destination_folder(namespace, path, store_name, create_if_missing=True)
        return {"path": path, "store": folder.Store.DisplayName}
    finally:
        pythoncom.CoUninitialize()


def preview_move(entry_ids, destination_path, store_name=None):
    """Read-only: describe exactly what a move would do, without moving anything."""
    pythoncom.CoInitialize()
    try:
        namespace = _connect()
        try:
            _resolve_destination_folder(namespace, destination_path, store_name, create_if_missing=False)
            destination_exists = True
        except OutlookError:
            destination_exists = False

        items = []
        for entry_id in entry_ids:
            try:
                item = namespace.GetItemFromID(entry_id)
                if getattr(item, "Class", None) != OL_MAIL_ITEM_CLASS:
                    items.append({"entry_id": entry_id, "status": "skipped", "reason": "not a mail item"})
                    continue
                items.append(
                    {
                        "entry_id": entry_id,
                        "status": "would_move",
                        "subject": item.Subject,
                        "sender_name": item.SenderName,
                        "received_time": str(item.ReceivedTime),
                        "current_folder": item.Parent.Name,
                    }
                )
            except Exception as exc:
                items.append({"entry_id": entry_id, "status": "error", "reason": str(exc)})

        return {
            "destination_path": destination_path,
            "destination_exists": destination_exists,
            "will_create_destination": not destination_exists,
            "items": items,
        }
    finally:
        pythoncom.CoUninitialize()


def move_emails(entry_ids, destination_path, store_name=None, confirm=False):
    """Move mail items (by EntryID) to destination_path, creating it if missing.

    Refuses to run unless confirm=True. Callers must have already shown the
    user a preview_move() result for these same arguments and gotten explicit
    agreement before setting confirm=True.
    """
    if not confirm:
        raise OutlookError(
            "Refusing to move emails without confirm=True. Call preview_move first, "
            "show the result to the user, and only retry with confirm=True after "
            "they explicitly agree."
        )
    pythoncom.CoInitialize()
    try:
        namespace = _connect()
        dest_folder = _resolve_destination_folder(namespace, destination_path, store_name, create_if_missing=True)
        results = []
        for entry_id in entry_ids:
            try:
                item = namespace.GetItemFromID(entry_id)
                if getattr(item, "Class", None) != OL_MAIL_ITEM_CLASS:
                    results.append({"entry_id": entry_id, "status": "skipped", "reason": "not a mail item"})
                    continue
                subject = item.Subject
                item.Move(dest_folder)
                results.append({"entry_id": entry_id, "status": "moved", "subject": subject})
            except Exception as exc:
                results.append({"entry_id": entry_id, "status": "error", "reason": str(exc)})
        return results
    finally:
        pythoncom.CoUninitialize()
