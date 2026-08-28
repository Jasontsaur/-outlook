"""Outlook MCP Server.

Exposes Outlook 2016 access over stdio to any MCP client (Claude Code /
Claude Desktop / local Cowork):

Read-only:
  - outlook_list_folders
  - outlook_list_recent_emails
  - outlook_search
  - outlook_get_email
  - outlook_classify_recent   (suggests destinations via rules.yaml, moves nothing)

Write (guarded):
  - outlook_create_folder     (idempotent)
  - outlook_preview_move      (read-only - shows what a move WOULD do)
  - outlook_move_emails       (refuses to run unless confirm=True)

Safety rule for any MCP client driving this server: always call
outlook_preview_move and show the result to the user before ever calling
outlook_move_emails with confirm=True.

Run with:
    python server.py
(the MCP client launches this for you over stdio - see README.md)
"""

import csv
import re
import time
from pathlib import Path
from typing import Optional

from mcp.server.fastmcp import FastMCP

import classify
import outlook

mcp = FastMCP("outlook-mcp")

BASE_DIR = Path(__file__).parent
RULES_PATH = BASE_DIR / "rules.yaml"
STAGING_DIR = BASE_DIR / "staging"


def _write_staging_csv(rows, destination_path):
    STAGING_DIR.mkdir(exist_ok=True)
    slug = re.sub(r"[^A-Za-z0-9_\-一-鿿]+", "_", destination_path.strip("/")) or "candidates"
    path = STAGING_DIR / f"{slug}_{int(time.time())}.csv"
    fieldnames = ["entry_id", "status", "received_time", "sender_name", "subject", "current_folder", "reason"]
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for row in rows:
            writer.writerow({k: row.get(k, "") for k in fieldnames})
    return str(path)


@mcp.tool()
def outlook_list_folders() -> list[dict]:
    """List every Outlook folder across all connected stores (accounts/PSTs).

    Returns each folder's store name, full '/'-separated path, and item
    count. Read-only - does not open or modify anything.
    """
    return outlook.list_folders()


@mcp.tool()
def outlook_list_recent_emails(
    folder_path: str = "Inbox",
    count: int = 10,
    store_name: Optional[str] = None,
) -> list[dict]:
    """List metadata (subject/sender/date - no body) for the most recent emails in a folder.

    Args:
        folder_path: '/'-separated folder path, e.g. "Inbox" or "Inbox/Projects".
        count: how many recent emails to return (default 10, keep this small).
        store_name: optional display name of the Outlook store/account to
            search, only needed if multiple accounts/PSTs are connected.
    """
    return outlook.list_recent_emails(folder_path=folder_path, count=count, store_name=store_name)


@mcp.tool()
def outlook_search(
    folder_path: str = "Inbox",
    date_from: Optional[str] = None,
    date_to: Optional[str] = None,
    subject_contains: Optional[list[str]] = None,
    sender_contains: Optional[list[str]] = None,
    max_results: int = 500,
    store_name: Optional[str] = None,
) -> list[dict]:
    """Search a folder's emails by date range / subject / sender keywords. Read-only, metadata only (no body).

    Args:
        folder_path: '/'-separated folder path, e.g. "Inbox" or "Projects".
        date_from / date_to: "YYYY-MM-DD", inclusive. Omit either for open-ended.
        subject_contains: keep only emails whose Subject contains any of these (case-insensitive).
        sender_contains: keep only emails whose sender name/email contains any of these.
        max_results: safety cap on how many emails to return (default 500) - keep this
            small and narrow the search (date range, keywords) rather than raising it,
            to avoid burning tokens on a large result set.
        store_name: optional Outlook store/account display name, for multi-account setups.
    """
    return outlook.search_emails(
        folder_path=folder_path,
        date_from=date_from,
        date_to=date_to,
        subject_contains=subject_contains,
        sender_contains=sender_contains,
        max_results=max_results,
        store_name=store_name,
    )


@mcp.tool()
def outlook_get_email(entry_id: str) -> dict:
    """Get full content (body + attachment names, not attachment content) of one email by EntryID.

    Only call this after outlook_search / outlook_classify_recent has identified
    a specific email worth reading in full - do not bulk-call this over an
    entire search result, that defeats the point of searching metadata first.
    """
    return outlook.get_email(entry_id)


@mcp.tool()
def outlook_classify_recent(
    folder_path: str = "Inbox",
    date_from: Optional[str] = None,
    date_to: Optional[str] = None,
    max_results: int = 500,
    store_name: Optional[str] = None,
) -> dict:
    """Search a folder and classify each email against outlook-mcp/rules.yaml.

    READ-ONLY - only suggests a destination folder + confidence per email,
    moves nothing. Always run this (and review counts_by_destination /
    low-confidence candidates with the user) before ever calling
    outlook_preview_move or outlook_move_emails.
    """
    emails = outlook.search_emails(
        folder_path=folder_path,
        date_from=date_from,
        date_to=date_to,
        max_results=max_results,
        store_name=store_name,
    )
    rules = classify.load_rules(RULES_PATH)

    counts: dict = {}
    candidates = []
    for email in emails:
        destination, confidence, rule_id = classify.classify(email, rules)
        key = destination or "UNMATCHED"
        counts[key] = counts.get(key, 0) + 1
        candidates.append(
            {**email, "suggested_destination": destination, "confidence": confidence, "matched_rule": rule_id}
        )

    return {"total_scanned": len(emails), "counts_by_destination": counts, "candidates": candidates}


@mcp.tool()
def outlook_create_folder(path: str, store_name: Optional[str] = None) -> dict:
    """Create an Outlook folder (and any missing parent folders) at a '/'-separated
    path, e.g. "Projects/XX工程". Idempotent - safe to call if it already exists.
    """
    return outlook.create_folder(path, store_name=store_name)


@mcp.tool()
def outlook_preview_move(
    entry_ids: list[str],
    destination_path: str,
    store_name: Optional[str] = None,
) -> dict:
    """Show exactly what WOULD move if outlook_move_emails were called with these
    same arguments. Does NOT move anything.

    Always call this and show the full result to the user before ever calling
    outlook_move_emails. Also writes a CSV audit record under
    outlook-mcp/staging/ so there is a record of what was proposed.
    """
    result = outlook.preview_move(entry_ids, destination_path, store_name=store_name)
    result["staging_csv"] = _write_staging_csv(result["items"], destination_path)
    return result


@mcp.tool()
def outlook_move_emails(
    entry_ids: list[str],
    destination_path: str,
    store_name: Optional[str] = None,
    confirm: bool = False,
) -> list[dict]:
    """Move the given emails (by EntryID) to destination_path, creating it if missing.

    This performs a real Outlook move (recoverable via Deleted Items / folder
    history, but still a real write). You MUST call outlook_preview_move with
    the same entry_ids/destination_path first, show the user that preview,
    and get their explicit go-ahead. `confirm` must then be set to True, or
    this call is refused.
    """
    return outlook.move_emails(entry_ids, destination_path, store_name=store_name, confirm=confirm)


if __name__ == "__main__":
    mcp.run(transport="stdio")
