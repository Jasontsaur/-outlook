"""Read-only wrapper around the Outlook 2016 COM Object Model.

V0.1 scope: attach to a running Outlook instance and read folder / email
metadata only. No writes, no moves, no deletes anywhere in this module.

Must run on Windows, in the same desktop session as Outlook 2016, with
Outlook already open at least once (COM will otherwise fail to attach).
"""

import pythoncom
import win32com.client

OL_MAIL_ITEM_CLASS = 43  # olMail


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
    """Resolve a '/'-separated folder path (e.g. 'Inbox/Projects') to a Folder object."""
    parts = [p for p in folder_path.split("/") if p]
    for store in namespace.Stores:
        if store_name and store.DisplayName != store_name:
            continue
        try:
            node = store.GetRootFolder()
        except Exception:
            continue
        matched = True
        for part in parts:
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
