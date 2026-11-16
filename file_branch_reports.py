"""File the daily branch-cash reports out of info@'s Sent Items.

ERPNext sends "التقرير اليومي لمصاريف وحركة خزائن الفروع" from info@diri-alnassir.iq
every day, and Exchange keeps the copy in Sent Items. Inbox rules never run on
Sent Items, so this job moves every matching message into the subfolder
Sent Items / التقارير اليومية لخزائن الفروع (created on first run).

Same OIDC sign-in as report.py; no secret is stored. The app may touch info@
only (Exchange RBAC for Applications, see setup-tenant.ps1). It reads subject
and sender only and never opens message bodies.
"""
import os
import sys
import urllib.parse

from report import GRAPH, graph, graph_token

MAILBOX = os.environ.get("FILE_MAILBOX", "info@diri-alnassir.iq")
SUBJECT_PREFIX = "التقرير اليومي لمصاريف وحركة خزائن الفروع"
FOLDER_NAME = "التقارير اليومية لخزائن الفروع"
DRY_RUN = os.environ.get("DRY_RUN", "").lower() in ("1", "true", "yes")

BOX = f"{GRAPH}/v1.0/users/{urllib.parse.quote(MAILBOX)}"


def ensure_folder(token):
    url = f"{BOX}/mailFolders/sentitems/childFolders?" + urllib.parse.urlencode(
        {"$filter": f"displayName eq '{FOLDER_NAME}'", "$select": "id,displayName,totalItemCount"})
    found = graph(token, url).get("value", [])
    if found:
        print(f"Folder exists ({found[0]['totalItemCount']} items already in it).")
        return found[0]["id"]
    if DRY_RUN:
        print("Dry run: would create the folder.")
        return None
    made = graph(token, f"{BOX}/mailFolders/sentitems/childFolders", {"displayName": FOLDER_NAME})
    print("Created folder under Sent Items.")
    return made["id"]


def matching_messages(token):
    url = f"{BOX}/mailFolders/sentitems/messages?" + urllib.parse.urlencode(
        {"$select": "id,subject,from,sentDateTime", "$top": "200"})
    scanned = 0
    while url:
        page = graph(token, url)
        for m in page.get("value", []):
            scanned += 1
            sender = ((m.get("from") or {}).get("emailAddress") or {}).get("address", "").lower()
            if DRY_RUN and "خزائن" in (m.get("subject") or "") and scanned <= 400:
                print("  diag:", repr(sender), ascii((m.get("subject") or "")[:12]), ascii(SUBJECT_PREFIX[:12]))
            if sender == MAILBOX.lower() and (m.get("subject") or "").strip().startswith(SUBJECT_PREFIX):
                yield m
        url = page.get("@odata.nextLink")
    if DRY_RUN:
        print(f"Scanned {scanned} messages in Sent Items.")


def main():
    token = graph_token()
    folder_id = ensure_folder(token)
    hits = list(matching_messages(token))
    print(f"Matching reports still in Sent Items: {len(hits)}")
    moved = failed = 0
    for m in sorted(hits, key=lambda m: m.get("sentDateTime") or ""):
        if DRY_RUN:
            print("  would move:", m.get("sentDateTime"), m.get("subject"))
            continue
        try:
            graph(token, f"{BOX}/messages/{m['id']}/move", {"destinationId": folder_id})
            moved += 1
        except Exception as e:  # keep going; the next run retries what is left
            failed += 1
            print("  move failed:", m.get("sentDateTime"), e, file=sys.stderr)
    if hits and not DRY_RUN:
        left = sum(1 for _ in matching_messages(token))
        print(f"Moved {moved}, failed {failed}, left in Sent Items {left}.")
        if left:
            sys.exit(1)


if __name__ == "__main__":
    main()
