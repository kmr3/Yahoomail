import csv
import json
import mimetypes
import os
import socket
import ssl
import smtplib
import sys
import threading
import traceback
from datetime import datetime
from email.message import EmailMessage
from email.utils import formataddr
from pathlib import Path
from tkinter import filedialog

import customtkinter as ctk
from dotenv import load_dotenv


ALLOWED_EXTENSIONS = {".xlsx", ".xls", ".xlsm"}
LOG_HEADERS = [
    "sent_at",
    "result",
    "recipient_name",
    "recipient_email",
    "subject",
    "attachment_name",
    "message",
]

mimetypes.add_type(
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", ".xlsx"
)
mimetypes.add_type("application/vnd.ms-excel", ".xls")
mimetypes.add_type("application/vnd.ms-excel.sheet.macroEnabled.12", ".xlsm")


class UserFacingError(Exception):
    """Error text that is safe to show directly to the user."""


def app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


BASE_DIR = app_dir()
CONFIG_PATH = BASE_DIR / "config.json"
LOG_PATH = BASE_DIR / "send_log.csv"
ERROR_LOG_PATH = BASE_DIR / "error_log.txt"
ENV_PATH = BASE_DIR / ".env"


def load_config() -> dict:
    if not CONFIG_PATH.exists():
        raise UserFacingError(
            "設定ファイル config.json が見つかりません。アプリと同じ場所に置いてください。"
        )

    try:
        with CONFIG_PATH.open("r", encoding="utf-8") as file:
            config = json.load(file)
    except json.JSONDecodeError as exc:
        raise UserFacingError(
            "config.json の書き方に問題があります。カンマやかっこの数を確認してください。"
        ) from exc

    return config


def recipient_from_config(config: dict) -> tuple[str, str]:
    recipient = config.get("recipient", {})
    name = str(recipient.get("name", "")).strip()
    email = str(recipient.get("email", "")).strip()
    return name, email


def sender_from_config(config: dict) -> tuple[str, str]:
    sender = config.get("sender", {})
    name = str(sender.get("name", "")).strip()
    email = str(sender.get("email", "")).strip()
    return name, email


def smtp_from_config(config: dict) -> tuple[str, int, str]:
    smtp = config.get("smtp", {})
    host = str(smtp.get("host", "")).strip()
    security = str(smtp.get("security", "ssl")).strip().lower()

    try:
        port = int(smtp.get("port", 465))
    except (TypeError, ValueError) as exc:
        raise UserFacingError(
            "SMTPのポート番号が正しくありません。config.json の smtp.port を確認してください。"
        ) from exc

    if not host:
        raise UserFacingError(
            "SMTPサーバー名が空です。config.json の smtp.host を確認してください。"
        )
    if security not in {"ssl", "starttls"}:
        raise UserFacingError(
            "SMTPの暗号化設定が正しくありません。config.json の smtp.security は ssl か starttls にしてください。"
        )

    return host, port, security


def auth_from_config(config: dict, sender_email: str) -> tuple[str, str]:
    auth = config.get("auth", {})
    username = str(auth.get("username", "")).strip() or sender_email

    if not username:
        raise UserFacingError(
            "送信用メールアドレスが空です。config.json の sender.email または auth.username を確認してください。"
        )

    if auth.get("use_windows_credential_manager"):
        service_name = str(
            auth.get("credential_service_name", "yahoo-mail-excel-sender")
        ).strip()
        password = get_password_from_windows_credentials(service_name, username)
    else:
        password_env = str(auth.get("password_env", "YAHOO_MAIL_PASSWORD")).strip()
        password = os.getenv(password_env)

    if not password:
        raise UserFacingError(
            "送信用パスワードが見つかりません。.env または Windows資格情報マネージャーを確認してください。"
        )

    return username, password


def get_password_from_windows_credentials(service_name: str, username: str) -> str | None:
    try:
        import keyring
    except ImportError as exc:
        raise UserFacingError(
            "Windows資格情報マネージャーを使う準備ができていません。requirements.txt を使って必要なものを入れ直してください。"
        ) from exc

    try:
        return keyring.get_password(service_name, username)
    except Exception as exc:
        raise UserFacingError(
            "Windows資格情報マネージャーからパスワードを取り出せませんでした。登録内容を確認してください。"
        ) from exc


def format_address(name: str, email: str) -> str:
    return formataddr((name, email), charset="utf-8") if name else email


def ensure_log_file() -> None:
    if LOG_PATH.exists():
        return

    with LOG_PATH.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file)
        writer.writerow(LOG_HEADERS)


def append_send_log(
    result: str,
    recipient_name: str,
    recipient_email: str,
    subject: str,
    attachment_name: str,
    message: str,
) -> None:
    ensure_log_file()
    with LOG_PATH.open("a", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file)
        writer.writerow(
            [
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                result,
                recipient_name,
                recipient_email,
                subject,
                attachment_name,
                message,
            ]
        )


def safe_append_send_log(
    result: str,
    recipient_name: str,
    recipient_email: str,
    subject: str,
    attachment_name: str,
    message: str,
) -> None:
    try:
        append_send_log(
            result,
            recipient_name,
            recipient_email,
            subject,
            attachment_name,
            message,
        )
    except OSError:
        pass


def write_error_log(exc: Exception, context: str = "") -> None:
    try:
        with ERROR_LOG_PATH.open("a", encoding="utf-8") as file:
            file.write(f"\n[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}]\n")
            if context:
                file.write(f"{context}\n")
            file.write("".join(traceback.format_exception(exc)))
    except OSError:
        pass


def connection_error_message(exc: Exception, host: str, port: int, security: str) -> str:
    base = (
        f"メールサーバーにつながりませんでした。config.json の smtp.host={host}, "
        f"smtp.port={port}, smtp.security={security} を確認してください。"
    )
    if isinstance(exc, socket.gaierror):
        return base + " サーバー名が間違っているか、インターネット接続に問題がある可能性があります。"
    if isinstance(exc, socket.timeout):
        return base + " 接続に時間がかかりすぎています。会社や学校のネットワークでブロックされている可能性があります。"
    if isinstance(exc, ssl.SSLError):
        return base + " SSL設定とポート番号の組み合わせが合っていない可能性があります。Yahoo!メールは通常 ssl と 465 です。"
    if isinstance(exc, ConnectionRefusedError):
        return base + " ポート番号が違うか、ネットワーク側で接続が止められている可能性があります。"
    return base + " インターネット接続、セキュリティソフト、会社や学校のネットワーク制限も確認してください。"


def validate_excel_file(path: Path) -> None:
    if not path.exists():
        raise UserFacingError("選んだExcelファイルが見つかりません。もう一度選び直してください。")
    if path.suffix.lower() not in ALLOWED_EXTENSIONS:
        raise UserFacingError("添付できるのは Excelファイル（.xlsx / .xls / .xlsm）だけです。")


def build_message(config: dict, subject: str, body: str, attachment_path: Path) -> EmailMessage:
    recipient_name, recipient_email = recipient_from_config(config)
    sender_name, sender_email = sender_from_config(config)

    if not recipient_email:
        raise UserFacingError("宛先メールアドレスが空です。config.json を確認してください。")
    if not sender_email:
        raise UserFacingError("差出人メールアドレスが空です。config.json を確認してください。")

    validate_excel_file(attachment_path)

    message = EmailMessage()
    message["Subject"] = subject
    message["From"] = format_address(sender_name, sender_email)
    message["To"] = format_address(recipient_name, recipient_email)
    message.set_content(body or "")

    content_type, _ = mimetypes.guess_type(str(attachment_path))
    if content_type:
        maintype, subtype = content_type.split("/", 1)
    else:
        maintype, subtype = "application", "octet-stream"

    with attachment_path.open("rb") as file:
        message.add_attachment(
            file.read(),
            maintype=maintype,
            subtype=subtype,
            filename=attachment_path.name,
        )

    return message


def send_email(config: dict, subject: str, body: str, attachment_path: Path) -> None:
    load_dotenv(ENV_PATH)

    recipient_name, recipient_email = recipient_from_config(config)
    sender_name, sender_email = sender_from_config(config)
    host, port, security = smtp_from_config(config)
    username, password = auth_from_config(config, sender_email)
    message = build_message(config, subject, body, attachment_path)

    try:
        if security == "ssl":
            context = ssl.create_default_context()
            with smtplib.SMTP_SSL(host, port, timeout=30, context=context) as server:
                server.login(username, password)
                server.send_message(message)
        else:
            context = ssl.create_default_context()
            with smtplib.SMTP(host, port, timeout=30) as server:
                server.starttls(context=context)
                server.login(username, password)
                server.send_message(message)

        safe_append_send_log(
            "success",
            recipient_name,
            recipient_email,
            subject,
            attachment_path.name,
            "送信完了",
        )
    except smtplib.SMTPAuthenticationError as exc:
        write_error_log(exc, "SMTP authentication failed.")
        safe_append_send_log(
            "error",
            recipient_name,
            recipient_email,
            subject,
            attachment_path.name,
            "認証エラー",
        )
        raise UserFacingError(
            "メールアドレスまたはパスワードが違う可能性があります。Yahoo!メールのSMTP利用設定とパスワードを確認してください。"
        ) from exc
    except smtplib.SMTPRecipientsRefused as exc:
        write_error_log(exc, "SMTP recipient refused.")
        safe_append_send_log(
            "error",
            recipient_name,
            recipient_email,
            subject,
            attachment_path.name,
            "宛先エラー",
        )
        raise UserFacingError(
            "宛先に送信できませんでした。宛先メールアドレスが正しいか確認してください。"
        ) from exc
    except smtplib.SMTPSenderRefused as exc:
        write_error_log(exc, "SMTP sender refused.")
        safe_append_send_log(
            "error",
            recipient_name,
            recipient_email,
            subject,
            attachment_path.name,
            "差出人エラー",
        )
        raise UserFacingError(
            "差出人メールアドレスに問題があります。config.json の sender.email を確認してください。"
        ) from exc
    except (
        smtplib.SMTPConnectError,
        smtplib.SMTPServerDisconnected,
        socket.timeout,
        socket.gaierror,
        ConnectionError,
        OSError,
    ) as exc:
        write_error_log(
            exc,
            f"SMTP connection failed. host={host}, port={port}, security={security}",
        )
        safe_append_send_log(
            "error",
            recipient_name,
            recipient_email,
            subject,
            attachment_path.name,
            "接続エラー",
        )
        raise UserFacingError(connection_error_message(exc, host, port, security)) from exc
    except smtplib.SMTPException as exc:
        write_error_log(exc, "SMTP send failed.")
        safe_append_send_log(
            "error",
            recipient_name,
            recipient_email,
            subject,
            attachment_path.name,
            "送信エラー",
        )
        raise UserFacingError(
            "メール送信中に問題が起きました。しばらくしてからもう一度お試しください。"
        ) from exc


class MailSenderApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()

        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        self.config_data = load_config()
        self.selected_file: Path | None = None
        self.is_sending = False

        app_config = self.config_data.get("app", {})
        self.title(str(app_config.get("title", "Excelメール送信")))
        self.geometry("760x680")
        self.minsize(700, 620)

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.create_widgets()

    def create_widgets(self) -> None:
        root = ctk.CTkFrame(self, fg_color="#f6f8fb", corner_radius=0)
        root.grid(row=0, column=0, sticky="nsew")
        root.grid_columnconfigure(0, weight=1)

        container = ctk.CTkFrame(root, fg_color="white", corner_radius=8)
        container.grid(row=0, column=0, padx=28, pady=24, sticky="nsew")
        container.grid_columnconfigure(0, weight=1)
        root.grid_rowconfigure(0, weight=1)

        title = ctk.CTkLabel(
            container,
            text="Excelファイルをメール送信",
            font=ctk.CTkFont(size=28, weight="bold"),
            text_color="#1f2937",
        )
        title.grid(row=0, column=0, padx=28, pady=(26, 8), sticky="w")

        recipient_name, recipient_email = recipient_from_config(self.config_data)
        recipient_text = f"{recipient_name}\n{recipient_email}" if recipient_name else recipient_email
        recipient_box = ctk.CTkFrame(container, fg_color="#eef6ff", corner_radius=8)
        recipient_box.grid(row=1, column=0, padx=28, pady=(10, 18), sticky="ew")
        recipient_box.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            recipient_box,
            text="送信先",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color="#1d4ed8",
        ).grid(row=0, column=0, padx=18, pady=(16, 4), sticky="w")
        ctk.CTkLabel(
            recipient_box,
            text=recipient_text or "宛先メールアドレスが未設定です",
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color="#111827",
            justify="left",
        ).grid(row=1, column=0, padx=18, pady=(0, 16), sticky="w")

        self.file_button = ctk.CTkButton(
            container,
            text="Excelファイルを選ぶ",
            font=ctk.CTkFont(size=20, weight="bold"),
            height=54,
            command=self.select_file,
        )
        self.file_button.grid(row=2, column=0, padx=28, pady=(4, 12), sticky="ew")

        self.file_label = ctk.CTkLabel(
            container,
            text="まだファイルが選ばれていません",
            font=ctk.CTkFont(size=22, weight="bold"),
            text_color="#374151",
            fg_color="#f3f4f6",
            corner_radius=8,
            height=78,
            wraplength=640,
        )
        self.file_label.grid(row=3, column=0, padx=28, pady=(0, 18), sticky="ew")

        ctk.CTkLabel(
            container,
            text="件名",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color="#374151",
        ).grid(row=4, column=0, padx=28, pady=(0, 6), sticky="w")
        self.subject_entry = ctk.CTkEntry(container, font=ctk.CTkFont(size=18), height=44)
        self.subject_entry.grid(row=5, column=0, padx=28, pady=(0, 16), sticky="ew")
        self.subject_entry.insert(
            0, str(self.config_data.get("app", {}).get("default_subject", ""))
        )

        ctk.CTkLabel(
            container,
            text="本文",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color="#374151",
        ).grid(row=6, column=0, padx=28, pady=(0, 6), sticky="w")
        self.body_text = ctk.CTkTextbox(container, font=ctk.CTkFont(size=17), height=150)
        self.body_text.grid(row=7, column=0, padx=28, pady=(0, 18), sticky="ew")
        self.body_text.insert(
            "1.0", str(self.config_data.get("app", {}).get("default_body", ""))
        )

        self.confirm_button = ctk.CTkButton(
            container,
            text="送信前に確認",
            font=ctk.CTkFont(size=22, weight="bold"),
            height=58,
            fg_color="#16a34a",
            hover_color="#15803d",
            command=self.open_confirmation,
        )
        self.confirm_button.grid(row=8, column=0, padx=28, pady=(0, 16), sticky="ew")

        self.status_label = ctk.CTkLabel(
            container,
            text="",
            font=ctk.CTkFont(size=14),
            text_color="#6b7280",
        )
        self.status_label.grid(row=9, column=0, padx=28, pady=(0, 18), sticky="w")

    def select_file(self) -> None:
        path = filedialog.askopenfilename(
            title="添付するExcelファイルを選んでください",
            filetypes=[
                ("Excelファイル", "*.xlsx *.xls *.xlsm"),
                ("すべてのファイル", "*.*"),
            ],
        )
        if not path:
            return

        selected = Path(path)
        try:
            validate_excel_file(selected)
        except UserFacingError as exc:
            self.show_message("ファイルを選べません", str(exc), kind="error")
            return

        self.selected_file = selected
        self.file_label.configure(text=selected.name, text_color="#111827", fg_color="#ecfdf5")
        self.status_label.configure(text="ファイルを確認しました。次に件名と本文を確認してください。")

    def open_confirmation(self) -> None:
        if self.is_sending:
            return

        try:
            self.validate_before_confirmation()
        except UserFacingError as exc:
            self.show_message("確認が必要です", str(exc), kind="warning")
            return

        recipient_name, recipient_email = recipient_from_config(self.config_data)
        subject = self.subject_entry.get().strip()
        body = self.body_text.get("1.0", "end").strip()
        attachment_name = self.selected_file.name if self.selected_file else ""

        dialog = ctk.CTkToplevel(self)
        dialog.title("送信前の確認")
        dialog.geometry("640x620")
        dialog.minsize(580, 560)
        dialog.transient(self)
        dialog.grab_set()
        dialog.grid_columnconfigure(0, weight=1)
        dialog.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(
            dialog,
            text="この内容で送信しますか？",
            font=ctk.CTkFont(size=24, weight="bold"),
            text_color="#111827",
        ).grid(row=0, column=0, padx=24, pady=(24, 12), sticky="w")

        details = ctk.CTkTextbox(dialog, font=ctk.CTkFont(size=17), wrap="word")
        details.grid(row=1, column=0, padx=24, pady=(0, 18), sticky="nsew")
        details.insert(
            "1.0",
            "\n".join(
                [
                    "宛先",
                    f"{recipient_name} <{recipient_email}>"
                    if recipient_name
                    else recipient_email,
                    "",
                    "件名",
                    subject or "（未入力）",
                    "",
                    "本文",
                    body or "（未入力）",
                    "",
                    "添付ファイル",
                    attachment_name,
                ]
            ),
        )
        details.configure(state="disabled")

        buttons = ctk.CTkFrame(dialog, fg_color="transparent")
        buttons.grid(row=2, column=0, padx=24, pady=(0, 24), sticky="ew")
        buttons.grid_columnconfigure((0, 1), weight=1)

        ctk.CTkButton(
            buttons,
            text="戻る",
            font=ctk.CTkFont(size=18, weight="bold"),
            height=48,
            fg_color="#6b7280",
            hover_color="#4b5563",
            command=dialog.destroy,
        ).grid(row=0, column=0, padx=(0, 8), sticky="ew")

        send_button = ctk.CTkButton(
            buttons,
            text="送信する",
            font=ctk.CTkFont(size=18, weight="bold"),
            height=48,
            fg_color="#dc2626",
            hover_color="#b91c1c",
            command=lambda: self.start_send(dialog, send_button),
        )
        send_button.grid(row=0, column=1, padx=(8, 0), sticky="ew")

    def validate_before_confirmation(self) -> None:
        _, recipient_email = recipient_from_config(self.config_data)
        if not recipient_email:
            raise UserFacingError("宛先メールアドレスが空です。config.json を確認してください。")
        if not self.selected_file:
            raise UserFacingError("添付するExcelファイルを選んでください。")
        validate_excel_file(self.selected_file)

    def start_send(self, dialog: ctk.CTkToplevel, send_button: ctk.CTkButton) -> None:
        if self.is_sending:
            return

        self.is_sending = True
        send_button.configure(state="disabled", text="送信中...")
        self.confirm_button.configure(state="disabled")
        self.file_button.configure(state="disabled")
        self.status_label.configure(text="送信しています。画面を閉じずにお待ちください。")

        subject = self.subject_entry.get().strip()
        body = self.body_text.get("1.0", "end").strip()
        attachment_path = self.selected_file

        thread = threading.Thread(
            target=self.send_in_background,
            args=(dialog, subject, body, attachment_path),
            daemon=True,
        )
        thread.start()

    def send_in_background(
        self,
        dialog: ctk.CTkToplevel,
        subject: str,
        body: str,
        attachment_path: Path | None,
    ) -> None:
        try:
            if attachment_path is None:
                raise UserFacingError("添付するExcelファイルを選んでください。")
            send_email(self.config_data, subject, body, attachment_path)
        except UserFacingError as exc:
            error_message = str(exc)
            self.after(0, lambda: self.finish_send(dialog, False, error_message))
        except Exception as exc:
            write_error_log(exc, "Unexpected application error.")
            self.after(
                0,
                lambda: self.finish_send(
                    dialog,
                    False,
                    "予期しない問題が起きました。最新版に更新してもう一度お試しください。解決しない場合は error_log.txt を確認してください。",
                ),
            )
        else:
            self.after(
                0,
                lambda: self.finish_send(
                    dialog,
                    True,
                    "メールを送信しました。送信履歴は send_log.csv に保存されています。",
                ),
            )

    def finish_send(self, dialog: ctk.CTkToplevel, success: bool, message: str) -> None:
        self.is_sending = False
        self.confirm_button.configure(state="normal")
        self.file_button.configure(state="normal")
        self.status_label.configure(text=message)

        if dialog.winfo_exists():
            dialog.destroy()

        if success:
            self.show_message("送信完了", message, kind="success")
        else:
            self.show_message("送信できませんでした", message, kind="error")

    def show_message(self, title: str, message: str, kind: str = "info") -> None:
        color_map = {
            "success": "#16a34a",
            "error": "#dc2626",
            "warning": "#d97706",
            "info": "#2563eb",
        }
        dialog = ctk.CTkToplevel(self)
        dialog.title(title)
        dialog.geometry("480x260")
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()
        dialog.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            dialog,
            text=title,
            font=ctk.CTkFont(size=22, weight="bold"),
            text_color=color_map.get(kind, "#2563eb"),
        ).grid(row=0, column=0, padx=24, pady=(26, 12), sticky="w")

        ctk.CTkLabel(
            dialog,
            text=message,
            font=ctk.CTkFont(size=17),
            text_color="#111827",
            justify="left",
            wraplength=420,
        ).grid(row=1, column=0, padx=24, pady=(0, 22), sticky="ew")

        ctk.CTkButton(
            dialog,
            text="OK",
            font=ctk.CTkFont(size=18, weight="bold"),
            height=44,
            command=dialog.destroy,
        ).grid(row=2, column=0, padx=24, pady=(0, 24), sticky="ew")


def main() -> None:
    try:
        app = MailSenderApp()
    except UserFacingError as exc:
        ctk.set_appearance_mode("light")
        app = ctk.CTk()
        app.withdraw()
        ErrorDialog(app, "起動できません", str(exc))
        app.mainloop()
        return

    app.mainloop()


class ErrorDialog(ctk.CTkToplevel):
    def __init__(self, parent: ctk.CTk, title: str, message: str) -> None:
        super().__init__(parent)
        self.title(title)
        self.geometry("500x260")
        self.resizable(False, False)
        self.grid_columnconfigure(0, weight=1)
        self.protocol("WM_DELETE_WINDOW", parent.destroy)

        ctk.CTkLabel(
            self,
            text=title,
            font=ctk.CTkFont(size=22, weight="bold"),
            text_color="#dc2626",
        ).grid(row=0, column=0, padx=24, pady=(28, 12), sticky="w")
        ctk.CTkLabel(
            self,
            text=message,
            font=ctk.CTkFont(size=17),
            justify="left",
            wraplength=440,
        ).grid(row=1, column=0, padx=24, pady=(0, 22), sticky="ew")
        ctk.CTkButton(
            self,
            text="閉じる",
            height=44,
            font=ctk.CTkFont(size=18, weight="bold"),
            command=parent.destroy,
        ).grid(row=2, column=0, padx=24, pady=(0, 24), sticky="ew")


if __name__ == "__main__":
    main()
