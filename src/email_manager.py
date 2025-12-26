import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from typing import Optional, Tuple

from mail_store import MailStore


class EmailManagerApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("邮箱管理系统")
        self.geometry("1000x650")
        self.state("zoomed")
        self.store = MailStore()
        self.selected_message_id: Optional[str] = None
        self.active_account: Optional[Tuple[str, str]] = None  # (email, provider_key)

        self.login_frame: Optional[ttk.Frame] = None
        self.main_frame: Optional[ttk.Frame] = None
        self.main_built = False

        # Shared variables used across screens
        provider_items = [name for _, name in self.store.list_providers()]
        provider_keys = [key for key, _ in self.store.list_providers()]
        self.provider_mapping = dict(zip(provider_items, provider_keys))
        self.provider_var = tk.StringVar(value=provider_items[0] if provider_items else "")
        self.login_email_var = tk.StringVar()
        self.login_pass_var = tk.StringVar()

        self._build_login_screen()

    # Layout
    def _build_login_screen(self) -> None:
        self.login_frame = ttk.Frame(self)
        self.login_frame.place(relx=0.5, rely=0.5, anchor="center")

        ttk.Label(self.login_frame, text="登录邮箱以同步真实内容", font=("Arial", 14, "bold")).grid(row=0, column=0, columnspan=3, pady=(0, 12))

        ttk.Label(self.login_frame, text="邮箱账号").grid(row=1, column=0, sticky="e", padx=6, pady=4)
        ttk.Entry(self.login_frame, textvariable=self.login_email_var, width=32).grid(row=1, column=1, columnspan=2, sticky="w", pady=4)

        ttk.Label(self.login_frame, text="授权码/密码").grid(row=2, column=0, sticky="e", padx=6, pady=4)
        ttk.Entry(self.login_frame, textvariable=self.login_pass_var, show="*", width=32).grid(row=2, column=1, columnspan=2, sticky="w", pady=4)

        ttk.Label(self.login_frame, text="邮箱类型").grid(row=3, column=0, sticky="e", padx=6, pady=4)
        provider_box = ttk.Combobox(
            self.login_frame, textvariable=self.provider_var, values=list(self.provider_mapping.keys()), state="readonly", width=30
        )
        provider_box.grid(row=3, column=1, columnspan=2, sticky="w", pady=4)
        provider_box.bind("<<ComboboxSelected>>", self._update_provider_info)

        self.imap_info_var = tk.StringVar()
        self.smtp_info_var = tk.StringVar()
        ttk.Label(self.login_frame, text="IMAP 服务器").grid(row=4, column=0, sticky="e", padx=6, pady=2)
        ttk.Label(self.login_frame, textvariable=self.imap_info_var).grid(row=4, column=1, columnspan=2, sticky="w")
        ttk.Label(self.login_frame, text="SMTP 服务器").grid(row=5, column=0, sticky="e", padx=6, pady=2)
        ttk.Label(self.login_frame, textvariable=self.smtp_info_var).grid(row=5, column=1, columnspan=2, sticky="w")

        ttk.Button(self.login_frame, text="登录并进入主界面", command=self.login_and_sync).grid(row=6, column=0, columnspan=3, pady=10)

        self._update_provider_info()

    def _build_layout(self) -> None:
        if self.login_frame:
            self.login_frame.destroy()
            self.login_frame = None

        self.main_frame = ttk.Frame(self)
        self.main_frame.grid(sticky="nsew")

        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        self.main_frame.columnconfigure(0, weight=3)
        self.main_frame.columnconfigure(1, weight=2)
        self.main_frame.rowconfigure(1, weight=1)

        # Search bar and controls
        control_frame = ttk.Frame(self.main_frame)
        control_frame.grid(row=0, column=0, sticky="ew", padx=8, pady=4)
        control_frame.columnconfigure(1, weight=1)

        ttk.Label(control_frame, text="搜索：").grid(row=0, column=0, padx=4)
        self.search_var = tk.StringVar()
        search_entry = ttk.Entry(control_frame, textvariable=self.search_var)
        search_entry.grid(row=0, column=1, sticky="ew")
        ttk.Button(control_frame, text="查找", command=self.refresh_messages).grid(row=0, column=2, padx=4)
        ttk.Button(control_frame, text="同步邮箱", command=self.login_and_sync).grid(row=0, column=3, padx=4)
        ttk.Button(control_frame, text="联系人管理", command=self.open_contacts_window).grid(row=0, column=4, padx=4)

        ttk.Label(control_frame, text="当前账号：").grid(row=1, column=0, sticky="e", padx=4, pady=(4, 0))
        self.current_account_var = tk.StringVar()
        ttk.Label(control_frame, textvariable=self.current_account_var).grid(row=1, column=1, columnspan=2, sticky="w", pady=(4, 0))
        ttk.Button(control_frame, text="切换账号", command=self.switch_account).grid(row=1, column=3, padx=4, pady=(4, 0))
        ttk.Button(control_frame, text="退出登录", command=self.logout_app).grid(row=1, column=4, padx=4, pady=(4, 0))

        # Folder filter
        folder_frame = ttk.Frame(self.main_frame)
        folder_frame.grid(row=1, column=0, sticky="nsew", padx=8, pady=4)
        folder_frame.rowconfigure(1, weight=1)
        ttk.Label(folder_frame, text="文件夹：").grid(row=0, column=0, sticky="w")
        self.folder_var = tk.StringVar(value="Inbox")
        folder_menu = ttk.Combobox(folder_frame, textvariable=self.folder_var, values=["Inbox", "Sent", "Archive"], state="readonly")
        folder_menu.grid(row=0, column=1, sticky="w")
        folder_menu.bind("<<ComboboxSelected>>", lambda _: self.refresh_messages())
        ttk.Button(folder_frame, text="归档", command=self.archive_selected).grid(row=0, column=2, padx=4)

        # Message list
        columns = ("sender", "subject", "timestamp", "folder")
        self.tree = ttk.Treeview(folder_frame, columns=columns, show="headings", height=15)
        for col, text in zip(columns, ["发件人", "主题", "时间", "文件夹"]):
            self.tree.heading(col, text=text)
            self.tree.column(col, width=150 if col != "subject" else 260)
        self.tree.grid(row=1, column=0, columnspan=3, sticky="nsew", pady=6)
        self.tree.bind("<<TreeviewSelect>>", lambda _: self.show_selected_message())

        scrollbar = ttk.Scrollbar(folder_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.grid(row=1, column=3, sticky="ns")

        # Compose frame
        compose = ttk.LabelFrame(self.main_frame, text="写邮件")
        compose.grid(row=0, column=1, rowspan=2, sticky="nsew", padx=8, pady=4)
        compose.columnconfigure(1, weight=1)
        compose.rowconfigure(6, weight=1)

        ttk.Label(compose, text="发件人").grid(row=0, column=0, padx=4, pady=2)
        self.from_var = tk.StringVar(value="user@example.com")
        ttk.Entry(compose, textvariable=self.from_var).grid(row=0, column=1, sticky="ew", padx=4, pady=2)

        ttk.Label(compose, text="收件人(逗号分隔)").grid(row=1, column=0, padx=4, pady=2)
        self.to_var = tk.StringVar()
        ttk.Entry(compose, textvariable=self.to_var).grid(row=1, column=1, sticky="ew", padx=4, pady=2)

        ttk.Label(compose, text="主题").grid(row=2, column=0, padx=4, pady=2)
        self.subject_var = tk.StringVar()
        ttk.Entry(compose, textvariable=self.subject_var).grid(row=2, column=1, sticky="ew", padx=4, pady=2)

        ttk.Label(compose, text="附件(路径,逗号分隔)").grid(row=3, column=0, padx=4, pady=2)
        self.attach_var = tk.StringVar()
        attach_entry = ttk.Entry(compose, textvariable=self.attach_var)
        attach_entry.grid(row=3, column=1, sticky="ew", padx=4, pady=2)
        ttk.Button(compose, text="选择附件", command=self.choose_attachment).grid(row=3, column=2, padx=4)

        ttk.Label(compose, text="正文").grid(row=4, column=0, padx=4, pady=2)
        self.body_text = tk.Text(compose, height=15)
        self.body_text.grid(row=4, column=1, columnspan=2, sticky="nsew", padx=4, pady=2)

        action_frame = ttk.Frame(compose)
        action_frame.grid(row=5, column=1, sticky="e", pady=4)
        ttk.Button(action_frame, text="发送", command=self.send_mail).grid(row=0, column=0, padx=4)
        ttk.Button(action_frame, text="转发选中", command=self.forward_selected).grid(row=0, column=1, padx=4)
        ttk.Button(action_frame, text="清空", command=self.clear_compose).grid(row=0, column=2, padx=4)

        # Detail view
        detail_frame = ttk.LabelFrame(folder_frame, text="邮件详情")
        detail_frame.grid(row=2, column=0, columnspan=3, sticky="nsew")
        detail_frame.columnconfigure(1, weight=1)

        self.detail_from = tk.StringVar()
        self.detail_to = tk.StringVar()
        self.detail_subject = tk.StringVar()
        self.detail_time = tk.StringVar()
        self.detail_attachments = tk.StringVar()

        ttk.Label(detail_frame, text="发件人：").grid(row=0, column=0, sticky="w")
        ttk.Label(detail_frame, textvariable=self.detail_from).grid(row=0, column=1, sticky="w")
        ttk.Label(detail_frame, text="收件人：").grid(row=1, column=0, sticky="w")
        ttk.Label(detail_frame, textvariable=self.detail_to).grid(row=1, column=1, sticky="w")
        ttk.Label(detail_frame, text="主题：").grid(row=2, column=0, sticky="w")
        ttk.Label(detail_frame, textvariable=self.detail_subject).grid(row=2, column=1, sticky="w")
        ttk.Label(detail_frame, text="时间：").grid(row=3, column=0, sticky="w")
        ttk.Label(detail_frame, textvariable=self.detail_time).grid(row=3, column=1, sticky="w")
        ttk.Label(detail_frame, text="附件：").grid(row=4, column=0, sticky="w")
        ttk.Label(detail_frame, textvariable=self.detail_attachments).grid(row=4, column=1, sticky="w")

        ttk.Button(detail_frame, text="删除", command=self.delete_selected).grid(row=0, column=2, padx=4)
        ttk.Button(detail_frame, text="刷新", command=self.refresh_messages).grid(row=1, column=2, padx=4)

        self.body_view = tk.Text(detail_frame, height=8, state="disabled")
        self.body_view.grid(row=5, column=0, columnspan=3, sticky="nsew", padx=4, pady=4)

    # Data helpers
    def refresh_messages(self) -> None:
        if not self.main_built:
            return
        query = self.search_var.get()
        folder = self.folder_var.get()
        messages = self.store.search_messages(query, folder if folder != "All" else None)

        for item in self.tree.get_children():
            self.tree.delete(item)
        for msg in messages:
            self.tree.insert("", tk.END, iid=msg.get("id"), values=(msg.get("sender"), msg.get("subject"), msg.get("timestamp"), msg.get("folder")))

        self.clear_details()

    def clear_details(self) -> None:
        self.selected_message_id = None
        for var in [self.detail_from, self.detail_to, self.detail_subject, self.detail_time, self.detail_attachments]:
            var.set("")
        self.body_view.configure(state="normal")
        self.body_view.delete("1.0", tk.END)
        self.body_view.configure(state="disabled")

    def _update_provider_info(self, *_: object) -> None:
        provider_key = self.provider_mapping.get(self.provider_var.get(), "")
        settings = self.store.provider_settings(provider_key) or {}
        self.imap_info_var.set(f"{settings.get('imap', '--')}:{settings.get('imap_port', '')}")
        self.smtp_info_var.set(f"{settings.get('smtp', '--')}:{settings.get('smtp_port', '')}")

    def login_and_sync(self) -> None:
        email_addr = self.login_email_var.get().strip()
        password = self.login_pass_var.get().strip()
        provider_name = self.provider_var.get()
        provider_key = self.provider_mapping.get(provider_name, "")
        if not email_addr or not password or not provider_key:
            messagebox.showwarning("提示", "请填写邮箱、授权码并选择服务商")
            return
        try:
            added = self.store.sync_imap(email_addr, password, provider_key)
            self.active_account = (email_addr, provider_key)
            if not self.main_built:
                self._build_layout()
                self.main_built = True
            self.from_var.set(email_addr)
            self.current_account_var.set(f"{email_addr}（{provider_name}）")
            self.refresh_messages()
            messagebox.showinfo("同步完成", f"成功同步 {added} 封邮件")
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("同步失败", f"无法同步邮件：{exc}")

    def show_selected_message(self) -> None:
        selection = self.tree.selection()
        if not selection:
            return
        message_id = selection[0]
        self.selected_message_id = message_id
        msg = self.store.get_message(message_id)
        if not msg:
            return
        self.detail_from.set(msg.get("sender", ""))
        self.detail_to.set(", ".join(msg.get("recipients", [])))
        self.detail_subject.set(msg.get("subject", ""))
        self.detail_time.set(msg.get("timestamp", ""))
        attachments = msg.get("attachments", [])
        self.detail_attachments.set(", ".join(attachments) if attachments else "无")

        self.body_view.configure(state="normal")
        self.body_view.delete("1.0", tk.END)
        body_text = msg.get("body", "")
        if msg.get("forwarded_from"):
            body_text = f"--- 转发自 {msg['forwarded_from']} ---\n" + body_text
        self.body_view.insert(tk.END, body_text)
        self.body_view.configure(state="disabled")

    def send_mail(self) -> None:
        sender = self.from_var.get().strip()
        recipients = [r.strip() for r in self.to_var.get().split(",") if r.strip()]
        subject = self.subject_var.get().strip() or "(无主题)"
        body = self.body_text.get("1.0", tk.END).strip()
        attachments = [a.strip() for a in self.attach_var.get().split(",") if a.strip()]

        if not sender or not recipients:
            messagebox.showwarning("提示", "请填写发件人和至少一个收件人。")
            return
        try:
            if self.active_account:
                self.store.send_smtp(
                    sender,
                    self.login_pass_var.get().strip(),
                    self.active_account[1],
                    recipients,
                    subject,
                    body,
                    attachments,
                )
                messagebox.showinfo("成功", "邮件已通过SMTP发送")
            else:
                messagebox.showinfo("成功", "未登录，已本地保存邮件")
            self.store.add_message(sender, recipients, subject, body, folder="Sent", attachments=attachments)
            self.refresh_messages()
            self.clear_compose()
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("发送失败", f"请检查账号或网络：{exc}")

    def forward_selected(self) -> None:
        if not self.selected_message_id:
            messagebox.showinfo("提示", "请选择需要转发的邮件")
            return
        msg = self.store.get_message(self.selected_message_id)
        if not msg:
            return
        self.subject_var.set("转发:" + msg.get("subject", ""))
        self.body_text.delete("1.0", tk.END)
        forward_body = f"\n\n--- 转发自 {msg.get('sender', '')} ---\n" + msg.get("body", "")
        self.body_text.insert(tk.END, forward_body)
        self.attach_var.set(", ".join(msg.get("attachments", [])))

    def receive_demo(self) -> None:
        self.store.receive_demo_message()
        self.refresh_messages()
        messagebox.showinfo("新邮件", "收到1封模拟邮件")

    def archive_selected(self) -> None:
        if not self.selected_message_id:
            messagebox.showinfo("提示", "请选择需要归档的邮件")
            return
        self.store.update_folder(self.selected_message_id, "Archive")
        self.refresh_messages()

    def delete_selected(self) -> None:
        if not self.selected_message_id:
            messagebox.showinfo("提示", "请选择需要删除的邮件")
            return
        self.store.delete_message(self.selected_message_id)
        self.refresh_messages()

    def clear_compose(self) -> None:
        for var in [self.to_var, self.subject_var, self.attach_var]:
            var.set("")
        if getattr(self, "body_text", None) and self.body_text.winfo_exists():
            self.body_text.delete("1.0", tk.END)

    def choose_attachment(self) -> None:
        file_paths = filedialog.askopenfilenames(title="选择附件")
        if file_paths:
            current = self.attach_var.get().split(",") if self.attach_var.get() else []
            all_paths = [*current, *file_paths]
            self.attach_var.set(", ".join([p for p in all_paths if p]))

    def switch_account(self) -> None:
        self.clear_compose()
        if self.main_frame:
            self.main_frame.destroy()
            self.main_frame = None
        self.main_built = False
        self.active_account = None
        self.selected_message_id = None
        self.current_account_var.set("")
        self.from_var.set("")
        self.login_email_var.set("")
        self.login_pass_var.set("")
        if self.provider_mapping:
            first_provider = next(iter(self.provider_mapping.keys()))
            self.provider_var.set(first_provider)
        self._build_login_screen()

    def logout_app(self) -> None:
        self.destroy()

    # Contacts window
    def open_contacts_window(self) -> None:
        window = tk.Toplevel(self)
        window.title("联系人管理")
        window.geometry("400x350")
        ContactManager(window, self.store)


class ContactManager:
    def __init__(self, master: tk.Toplevel, store: MailStore):
        self.master = master
        self.store = store
        self.frame = ttk.Frame(master)
        self.frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        self.tree = ttk.Treeview(self.frame, columns=("name", "email"), show="headings")
        self.tree.heading("name", text="姓名")
        self.tree.heading("email", text="邮箱")
        self.tree.column("name", width=120)
        self.tree.column("email", width=180)
        self.tree.pack(fill=tk.BOTH, expand=True)

        btn_frame = ttk.Frame(self.frame)
        btn_frame.pack(fill=tk.X, pady=4)
        ttk.Button(btn_frame, text="新增", command=self.add_contact).pack(side=tk.LEFT, padx=4)
        ttk.Button(btn_frame, text="编辑", command=self.edit_contact).pack(side=tk.LEFT, padx=4)
        ttk.Button(btn_frame, text="删除", command=self.delete_contact).pack(side=tk.LEFT, padx=4)
        ttk.Button(btn_frame, text="导入CSV", command=self.import_contacts).pack(side=tk.LEFT, padx=4)
        ttk.Button(btn_frame, text="导出CSV", command=self.export_contacts).pack(side=tk.LEFT, padx=4)

        self.refresh()

    def refresh(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)
        for idx, contact in enumerate(self.store.contacts()):
            self.tree.insert("", tk.END, iid=str(idx), values=(contact.get("name", ""), contact.get("email", "")))

    def add_contact(self) -> None:
        name, email = self._contact_dialog()
        if not name or not email:
            return
        self.store.add_contact(name, email)
        self.refresh()

    def edit_contact(self) -> None:
        selection = self.tree.selection()
        if not selection:
            return
        idx = int(selection[0])
        contact = self.store.contacts()[idx]
        name, email = self._contact_dialog(contact.get("name"), contact.get("email"))
        if not name or not email:
            return
        self.store.update_contact(idx, name, email)
        self.refresh()

    def delete_contact(self) -> None:
        selection = self.tree.selection()
        if not selection:
            return
        idx = int(selection[0])
        self.store.delete_contact(idx)
        self.refresh()

    def import_contacts(self) -> None:
        path = filedialog.askopenfilename(title="选择CSV文件", filetypes=[("CSV", "*.csv"), ("所有文件", "*.*")])
        if path:
            self.store.import_contacts(path)
            self.refresh()

    def export_contacts(self) -> None:
        path = filedialog.asksaveasfilename(title="导出CSV", defaultextension=".csv", filetypes=[("CSV", "*.csv")])
        if path:
            self.store.export_contacts(path)

    def _contact_dialog(self, default_name: str = "", default_email: str = "") -> tuple[str, str]:
        dialog = tk.Toplevel(self.master)
        dialog.title("联系人")
        ttk.Label(dialog, text="姓名").grid(row=0, column=0, padx=6, pady=6)
        name_var = tk.StringVar(value=default_name)
        ttk.Entry(dialog, textvariable=name_var).grid(row=0, column=1, padx=6, pady=6)
        ttk.Label(dialog, text="邮箱").grid(row=1, column=0, padx=6, pady=6)
        email_var = tk.StringVar(value=default_email)
        ttk.Entry(dialog, textvariable=email_var).grid(row=1, column=1, padx=6, pady=6)

        result: list[str] = []

        def confirm() -> None:
            result.extend([name_var.get().strip(), email_var.get().strip()])
            dialog.destroy()

        ttk.Button(dialog, text="确定", command=confirm).grid(row=2, column=0, columnspan=2, pady=8)
        dialog.wait_window()
        if len(result) == 2:
            return result[0], result[1]
        return "", ""


def main() -> None:
    app = EmailManagerApp()
    app.mainloop()


if __name__ == "__main__":
    main()
