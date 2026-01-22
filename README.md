# Email Bot (Outlook Draft Replies)

This agent scans Outlook emails and creates draft replies only when the message body contains language similar to:

> "not expected to improve with rework"

## Requirements

- Windows
- Outlook desktop application
- Python 3.10+

Install dependencies:

```
pip install -r requirements.txt
```

## Run

```
python main.py
```

## Behavior

- Scans Inbox for the last 24 hours
- Only unread messages (default)
- Creates a draft reply with a short acknowledgement
- Keeps a local state file `.email_agent_state.json` so emails are not processed twice

## Customize

Edit `scan_and_draft()` in [main.py](main.py) to adjust:

- `mailbox_name` (shared mailboxes)
- `folder_name`
- `days_back`
- `unread_only`
- `patterns`

## Notes

- Drafts are saved in the Outlook Drafts folder.
- The script does not send emails.
