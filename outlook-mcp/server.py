"""V0.1 Outlook MCP Server - READ ONLY.

Exposes two tools over stdio to any MCP client (Claude Code / Claude
Desktop / local Cowork):

  - outlook_list_folders
  - outlook_list_recent_emails

By design, no write/move/delete/create tool exists in this version.
Later phases (search, classification, preview-move, move) get added on
top of this once V0.1 is confirmed working end-to-end.

Run with:
    python server.py
(the MCP client launches this for you over stdio - see README.md)
"""

from typing import Optional

from mcp.server.fastmcp import FastMCP

import outlook

mcp = FastMCP("outlook-mcp")


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
    return outlook.list_recent_emails(
        folder_path=folder_path, count=count, store_name=store_name
    )


if __name__ == "__main__":
    mcp.run(transport="stdio")
