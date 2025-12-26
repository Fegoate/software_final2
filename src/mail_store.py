import email
import imaplib
import json
import os
import random
import smtplib
import string
from dataclasses import dataclass, asdict
from datetime import datetime
from email.header import decode_header, make_header
from email.message import EmailMessage
from email.utils import parsedate_to_datetime
from typing import Dict, List, Optional, Tuple


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
    message_uid: Optional[str] = None


class MailStore:
    def __init__(self, data_path: str = DATA_PATH):
        self.data_path = data_path
        self.data = {"messages": [], "contacts": []}
        self._ensure_paths()
        self.load()
        self.providers = self._default_providers()

    def _default_providers(self) -> Dict[str, Dict[str, str]]:
        return {
            "qq": {"name": "QQ 邮箱", "imap": "imap.qq.com", "smtp": "smtp.qq.com", "imap_port": 993, "smtp_port": 465},
            "163": {"name": "163 邮箱", "imap": "imap.163.com", "smtp": "smtp.163.com", "imap_port": 993, "smtp_port": 465},
            "126": {"name": "126 邮箱", "imap": "imap.126.com", "smtp": "smtp.126.com", "imap_port": 993, "smtp_port": 465},
            "sina": {"name": "新浪邮箱", "imap": "imap.sina.com", "smtp": "smtp.sina.com", "imap_port": 993, "smtp_port": 465},
            "aliyun": {"name": "阿里云邮箱", "imap": "imap.aliyun.com", "smtp": "smtp.aliyun.com", "imap_port": 993, "smtp_port": 465},
            "outlook": {"name": "Outlook", "imap": "outlook.office365.com", "smtp": "smtp.office365.com", "imap_port": 993, "smtp_port": 587},
        }

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

    def list_providers(self) -> List[Tuple[str, str]]:
        return [(key, value.get("name", key)) for key, value in self.providers.items()]

    def provider_settings(self, key: str) -> Optional[Dict[str, str]]:
        return self.providers.get(key)

    def add_message(
        self,
        sender: str,
        recipients: List[str],
        subject: str,
        body: str,
        folder: str = "Inbox",
        attachments: Optional[List[str]] = None,
        forwarded_from: Optional[str] = None,
        message_uid: Optional[str] = None,
        timestamp: Optional[str] = None,
    ) -> Message:
        message = Message(
            id=self._generate_id(),
            sender=sender,
            recipients=[r.strip() for r in recipients if r.strip()],
            subject=subject,
            body=body,
            timestamp=timestamp or datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            folder=folder,
            attachments=[a.strip() for a in attachments] if attachments else [],
            forwarded_from=forwarded_from,
            message_uid=message_uid,
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

    # IMAP/SMTP integration
    def sync_imap(self, email_address: str, password: str, provider_key: str, limit: int = 20) -> int:
        settings = self.provider_settings(provider_key)
        if not settings:
            return 0
        server = imaplib.IMAP4_SSL(settings["imap"], settings.get("imap_port", 993))
        server.login(email_address, password)
        status, _ = server.select("INBOX")
        if status != "OK":
            server.logout()
            raise RuntimeError("无法进入收件箱，请检查邮箱设置或授权码")
        typ, data = server.search(None, "ALL")
        if typ != "OK":
            server.logout()
            return 0
        ids = data[0].split()[-limit:]
        added = 0
        existing_uids = {msg.get("message_uid") for msg in self.data.get("messages", []) if msg.get("message_uid")}
        for msg_id in ids:
            typ, msg_data = server.fetch(msg_id, "(RFC822)")
            if typ != "OK" or not msg_data:
                continue
            raw_email = msg_data[0][1]
            parsed = email.message_from_bytes(raw_email)
            uid = parsed.get("Message-ID") or f"imap-{msg_id.decode()}"
            if uid in existing_uids:
                continue
            sender, recipients, subject, body, timestamp = self._parse_email_message(parsed)
            self.add_message(
                sender,
                recipients,
                subject,
                body,
                folder="Inbox",
                attachments=[],
                message_uid=uid,
                timestamp=timestamp,
            )
            added += 1
        server.logout()
        return added

    def _parse_email_message(self, message: email.message.Message) -> Tuple[str, List[str], str, str, str]:
        def decode(value: Optional[str]) -> str:
            if not value:
                return ""
            try:
                return str(make_header(decode_header(value)))
            except Exception:
                return value

        sender = decode(message.get("From"))
        recipients_raw = message.get_all("To", []) + message.get_all("Cc", [])
        recipients = [decode(addr) for addr in recipients_raw]
        subject = decode(message.get("Subject"))
        timestamp = ""
        if message.get("Date"):
            try:
                timestamp = parsedate_to_datetime(message.get("Date")).strftime("%Y-%m-%d %H:%M:%S")
            except Exception:
                timestamp = message.get("Date", "")
        body_parts: List[str] = []
        if message.is_multipart():
            for part in message.walk():
                if part.get_content_type() == "text/plain" and not part.get_filename():
                    try:
                        body_parts.append(part.get_payload(decode=True).decode(part.get_content_charset() or "utf-8", errors="ignore"))
                    except Exception:
                        continue
        else:
            try:
                body_parts.append(message.get_payload(decode=True).decode(message.get_content_charset() or "utf-8", errors="ignore"))
            except Exception:
                body_parts.append(message.get_payload())
        return sender, recipients, subject, "\n".join(body_parts).strip(), timestamp

    def send_smtp(self, email_address: str, password: str, provider_key: str, recipients: List[str], subject: str, body: str, attachments: Optional[List[str]] = None) -> None:
        settings = self.provider_settings(provider_key)
        if not settings:
            raise ValueError("未知邮箱服务提供商")
        msg = EmailMessage()
        msg["From"] = email_address
        msg["To"] = ", ".join(recipients)
        msg["Subject"] = subject
        msg.set_content(body)
        for path in attachments or []:
            if not os.path.exists(path):
                continue
            with open(path, "rb") as f:
                data = f.read()
            msg.add_attachment(data, maintype="application", subtype="octet-stream", filename=os.path.basename(path))

        if settings.get("smtp_port") == 587:
            server = smtplib.SMTP(settings["smtp"], settings["smtp_port"])
            server.starttls()
        else:
            server = smtplib.SMTP_SSL(settings["smtp"], settings.get("smtp_port", 465))
        server.login(email_address, password)
        server.send_message(msg)
        server.quit()


