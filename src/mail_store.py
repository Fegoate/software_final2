import json
import os
import random
import string
from datetime import datetime
from dataclasses import dataclass, asdict
from typing import List, Dict, Optional


DATA_PATH = os.path.join(os.path.dirname(__file__), "data", "mail_store.json")
CONTACT_PATH = os.path.join(os.path.dirname(__file__), "data", "contacts.csv")


@dataclass
class Message:
    id: str
    sender: str
    recipients: List[str]
    subject: str
    body: str
    timestamp: str
    folder: str = "Inbox"
    attachments: Optional[List[str]] = None
    forwarded_from: Optional[str] = None


class MailStore:
    def __init__(self, data_path: str = DATA_PATH):
        self.data_path = data_path
        self.data = {"messages": [], "contacts": []}
        self._ensure_paths()
        self.load()

    def _ensure_paths(self) -> None:
        directory = os.path.dirname(self.data_path)
        os.makedirs(directory, exist_ok=True)
        if not os.path.exists(self.data_path):
            with open(self.data_path, "w", encoding="utf-8") as f:
                json.dump(self.data, f, ensure_ascii=False, indent=2)

    def load(self) -> None:
        if os.path.exists(self.data_path):
            with open(self.data_path, "r", encoding="utf-8") as f:
                try:
                    self.data = json.load(f)
                except json.JSONDecodeError:
                    self.data = {"messages": [], "contacts": []}
        else:
            self.data = {"messages": [], "contacts": []}

    def save(self) -> None:
        with open(self.data_path, "w", encoding="utf-8") as f:
            json.dump(self.data, f, ensure_ascii=False, indent=2)

    def _generate_id(self) -> str:
        return "msg-" + "".join(random.choices(string.ascii_lowercase + string.digits, k=8))

    def add_message(
        self,
        sender: str,
        recipients: List[str],
        subject: str,
        body: str,
        folder: str = "Inbox",
        attachments: Optional[List[str]] = None,
        forwarded_from: Optional[str] = None,
    ) -> Message:
        message = Message(
            id=self._generate_id(),
            sender=sender,
            recipients=[r.strip() for r in recipients if r.strip()],
            subject=subject,
            body=body,
            timestamp=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            folder=folder,
            attachments=[a.strip() for a in attachments] if attachments else [],
            forwarded_from=forwarded_from,
        )
        self.data.setdefault("messages", []).append(asdict(message))
        self.save()
        return message

    def delete_message(self, message_id: str) -> None:
        messages = self.data.get("messages", [])
        self.data["messages"] = [m for m in messages if m.get("id") != message_id]
        self.save()

    def update_folder(self, message_id: str, folder: str) -> None:
        for msg in self.data.get("messages", []):
            if msg.get("id") == message_id:
                msg["folder"] = folder
                break
        self.save()

    def search_messages(self, query: str = "", folder: Optional[str] = None) -> List[Dict]:
        messages = self.data.get("messages", [])
        filtered = []
        query_lower = query.lower()
        for msg in messages:
            if folder and msg.get("folder") != folder:
                continue
            text = " ".join(
                [
                    msg.get("sender", ""),
                    " ".join(msg.get("recipients", [])),
                    msg.get("subject", ""),
                    msg.get("body", ""),
                ]
            ).lower()
            if query_lower in text:
                filtered.append(msg)
        return sorted(filtered, key=lambda m: m.get("timestamp", ""), reverse=True)

    def get_message(self, message_id: str) -> Optional[Dict]:
        for msg in self.data.get("messages", []):
            if msg.get("id") == message_id:
                return msg
        return None

    def receive_demo_message(self) -> Message:
        senders = ["alice@example.com", "bob@example.com", "carol@example.com"]
        subjects = ["会议提醒", "周报", "项目进展", "午餐邀约"]
        snippets = [
            "请注意明天下午3点的会议。",
            "本周主要完成了UI原型设计。",
            "项目按计划推进，暂无风险。",
            "一起去品尝新开的餐厅吗？",
        ]
        sender = random.choice(senders)
        subject = random.choice(subjects)
        body = random.choice(snippets)
        return self.add_message(sender, ["user@example.com"], subject, body, folder="Inbox")

    # Contact management
    def add_contact(self, name: str, email: str) -> None:
        self.data.setdefault("contacts", []).append({"name": name, "email": email})
        self.save()

    def update_contact(self, index: int, name: str, email: str) -> None:
        contacts = self.data.setdefault("contacts", [])
        if 0 <= index < len(contacts):
            contacts[index] = {"name": name, "email": email}
            self.save()

    def delete_contact(self, index: int) -> None:
        contacts = self.data.setdefault("contacts", [])
        if 0 <= index < len(contacts):
            contacts.pop(index)
            self.save()

    def import_contacts(self, path: str) -> None:
        if not os.path.exists(path):
            return
        with open(path, "r", encoding="utf-8") as f:
            for line in f:
                name, email = line.strip().split(",", 1)
                self.data.setdefault("contacts", []).append({"name": name, "email": email})
        self.save()

    def export_contacts(self, path: str) -> None:
        contacts = self.data.get("contacts", [])
        with open(path, "w", encoding="utf-8") as f:
            for contact in contacts:
                f.write(f"{contact.get('name')},{contact.get('email')}\n")

    def contacts(self) -> List[Dict[str, str]]:
        return self.data.get("contacts", [])


