import json
import os
import re
from datetime import datetime, timedelta
from typing import Iterable, List, Tuple

import win32com.client

STATE_FILE = ".email_agent_state.json"

DEFAULT_PATTERNS = [
    r"not expected to improve with rework",
    r"not expected to improve with (any )?rework",
    r"additional rework is not expected to improve",
]

# Outlook constants for reply verbs
OL_REPLY = 102
OL_REPLY_ALL = 103


def load_state(path: str) -> dict:
    if not os.path.exists(path):
        return {"processed": {}}
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def save_state(path: str, state: dict) -> None:
    with open(path, "w", encoding="utf-8") as f:
        json.dump(state, f, indent=2)


def compile_patterns(patterns: Iterable[str]) -> List[re.Pattern]:
    return [re.compile(p, re.IGNORECASE) for p in patterns]


def body_matches(body: str, patterns: List[re.Pattern]) -> bool:
    if not body:
        return False
    return any(p.search(body) for p in patterns)


def build_draft_reply_body(sender_name: str) -> str:
    return (
        f"Hi {sender_name},\n\n"
        "Please comment and send lot on if rework is not expected to improve.\n\n"
    )


def iter_messages(folder, since: datetime, unread_only: bool) -> Iterable:
    items = folder.Items
    items.Sort("[ReceivedTime]", True)

    restrictions = []
    if unread_only:
        restrictions.append("[UnRead] = True")
    if since:
        since_str = since.strftime("%m/%d/%Y %H:%M %p")
        restrictions.append(f"[ReceivedTime] >= '{since_str}'")

    if restrictions:
        filter_str = " AND ".join(restrictions)
        items = items.Restrict(filter_str)

    for item in items:
        yield item


def get_mail_identifier(mail) -> Tuple[str, str]:
    entry_id = getattr(mail, "EntryID", None)
    store_id = getattr(mail, "StoreID", None)
    if not store_id:
        try:
            store_id = mail.Parent.StoreID
        except Exception:
            store_id = "UNKNOWN_STORE"
    if not entry_id:
        entry_id = f"UNKNOWN_ENTRY_{id(mail)}"
    return entry_id, store_id


def already_processed(state: dict, mail) -> bool:
    entry_id, store_id = get_mail_identifier(mail)
    return state.get("processed", {}).get(store_id, {}).get(entry_id, False)


def mark_processed(state: dict, mail) -> None:
    entry_id, store_id = get_mail_identifier(mail)
    state.setdefault("processed", {}).setdefault(store_id, {})[entry_id] = True


def already_replied(mail) -> bool:
    try:
        last_verb = getattr(mail, "LastVerbExecuted", None)
        if last_verb in (OL_REPLY, OL_REPLY_ALL):
            return True
        replied = getattr(mail, "Replied", None)
        if replied is True:
            return True
    except Exception:
        return False
    return False


def subject_is_reply(mail) -> bool:
    subject = getattr(mail, "Subject", "") or ""
    return "RE:" in subject.upper()


def create_draft_reply(mail) -> None:
    reply = mail.Reply()
    sender_name = mail.SenderName or "there"
    generated = build_draft_reply_body(sender_name)
    reply.Body = f"{generated}\n\n{reply.Body}"
    reply.Save()


def scan_and_draft(
    mailbox_name: str | None = None,
    folder_name: str = "Inbox",
    days_back: int = 1,
    unread_only: bool = True,
    patterns: Iterable[str] = DEFAULT_PATTERNS,
) -> int:
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")

    if mailbox_name:
        mailbox = namespace.Folders.Item(mailbox_name)
        folder = mailbox.Folders.Item(folder_name)
    else:
        folder = namespace.GetDefaultFolder(6)  # 6 = Inbox

    state = load_state(STATE_FILE)
    regexes = compile_patterns(patterns)

    since = datetime.now() - timedelta(days=days_back) if days_back else None

    drafted = 0
    for mail in iter_messages(folder, since, unread_only):
        if mail.Class != 43:  # 43 = MailItem
            continue
        if already_processed(state, mail):
            continue
        if already_replied(mail):
            continue
        if subject_is_reply(mail):
            continue
        if not body_matches(mail.Body, regexes):
            continue

        create_draft_reply(mail)
        mark_processed(state, mail)
        drafted += 1

    save_state(STATE_FILE, state)
    return drafted


if __name__ == "__main__":
    drafted_count = scan_and_draft()
    print(f"Drafted replies: {drafted_count}")
