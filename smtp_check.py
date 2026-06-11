import json
import smtplib
import ssl
from pathlib import Path

from dotenv import load_dotenv


BASE_DIR = Path(__file__).resolve().parent
CONFIG_PATH = BASE_DIR / "config.json"
ENV_PATH = BASE_DIR / ".env"


def load_config() -> dict:
    with CONFIG_PATH.open("r", encoding="utf-8") as file:
        return json.load(file)


def main() -> None:
    print("Yahoo!メール SMTP 接続診断を開始します。")

    load_dotenv(ENV_PATH)
    config = load_config()

    smtp = config.get("smtp", {})
    sender = config.get("sender", {})
    auth = config.get("auth", {})

    host = str(smtp.get("host", "smtp.mail.yahoo.co.jp")).strip()
    port = int(smtp.get("port", 465))
    security = str(smtp.get("security", "ssl")).strip().lower()
    sender_email = str(sender.get("email", "")).strip()
    username = str(auth.get("username", "")).strip() or sender_email
    password_env = str(auth.get("password_env", "YAHOO_MAIL_PASSWORD")).strip()

    print(f"SMTPサーバー: {host}")
    print(f"ポート: {port}")
    print(f"暗号化: {security}")
    print(f"ログイン名: {username or '未設定'}")
    print(f"パスワード設定: {password_env}")

    if security != "ssl":
        print("この診断は ssl / 465 の確認用です。config.json は ssl / 465 をおすすめします。")
        return

    try:
        print("1. SSL接続を確認しています...")
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL(host, port, timeout=30, context=context) as server:
            print("   OK: SSL接続できました。")

            print("2. Yahoo!メールサーバーの応答を確認しています...")
            server.ehlo()
            print("   OK: サーバーから応答がありました。")

            password = None
            if auth.get("use_windows_credential_manager"):
                import keyring

                service_name = str(
                    auth.get("credential_service_name", "yahoo-mail-excel-sender")
                ).strip()
                password = keyring.get_password(service_name, username)
            else:
                import os

                password = os.getenv(password_env)

            if not username or not password:
                print("3. ログイン確認はスキップしました。ログイン名またはパスワードが未設定です。")
                return

            print("3. ログインを確認しています...")
            server.login(username, password)
            print("   OK: ログインできました。アプリからも送信できる可能性が高いです。")
    except smtplib.SMTPAuthenticationError as exc:
        print("NG: ログインできませんでした。")
        print("Yahoo!メールのSMTP利用設定、ログイン名、パスワードを確認してください。")
        print(f"詳細: {exc}")
    except ssl.SSLError as exc:
        print("NG: SSL接続で止まりました。")
        print("セキュリティソフト、VPN、プロキシ、SSL検査機能が影響している可能性があります。")
        print(f"詳細: {exc}")
    except TimeoutError as exc:
        print("NG: 接続に時間がかかりすぎました。")
        print("ネットワークやセキュリティソフトの制限を確認してください。")
        print(f"詳細: {exc}")
    except OSError as exc:
        print("NG: 接続中に問題が起きました。")
        print("ネットワーク、セキュリティソフト、VPN、SMTP設定を確認してください。")
        print(f"詳細: {exc}")
    except Exception as exc:
        print("NG: 予期しない問題が起きました。")
        print(f"詳細: {type(exc).__name__}: {exc}")


if __name__ == "__main__":
    main()
