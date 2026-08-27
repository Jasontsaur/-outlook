"""Quick standalone sanity check - run this BEFORE wiring up MCP at all.

Confirms Python can talk to Outlook 2016 via COM on this machine:
    python test_connection.py

Expect to see your folder list and the 10 most recent Inbox emails
printed to the console. If this fails, MCP will fail too - fix this
first (see README.md troubleshooting section).
"""

import outlook


def main():
    print("Connecting to Outlook via COM...")
    folders = outlook.list_folders()
    print(f"\nFound {len(folders)} folders:\n")
    for f in folders[:30]:
        print(f"  [{f['store']}] {f['path']}  ({f['item_count']} items)")
    if len(folders) > 30:
        print(f"  ... and {len(folders) - 30} more")

    print("\nMost recent 10 emails in Inbox:\n")
    emails = outlook.list_recent_emails("Inbox", count=10)
    if not emails:
        print("  (no emails found - check the folder path/name)")
    for e in emails:
        flag = "UNREAD" if e["unread"] else "      "
        print(f"  [{flag}] {e['received_time']}  {e['sender_name']!r:30s}  {e['subject']}")

    print("\nOK - Outlook COM connection works.")


if __name__ == "__main__":
    main()
