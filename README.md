# Excelメール送信アプリ

WindowsでExcelファイルを選び、指定された宛先へYahoo!メールのSMTPで送信するためのシンプルなデスクトップアプリです。

送信前に必ず確認画面を表示し、送信結果は `send_log.csv` に保存します。

## セットアップ方法

1. Python 3.11以降をインストールします。
2. このフォルダで仮想環境を作成します。

```powershell
python -m venv .venv
.\.venv\Scripts\activate
```

3. 必要なものをインストールします。

```powershell
python -m pip install --upgrade pip
pip install -r requirements.txt
```

4. `config.json` を実際の宛先・差出人に書き換えます。
5. `.env` ファイルを `main.py` と同じ場所に作成し、Yahoo!メールのパスワードを書きます。

```env
YAHOO_MAIL_PASSWORD=ここにYahooメールのパスワード
```

パスワードは `main.py` や `config.json` には書かないでください。

## Yahoo!メールでSMTPを使うための設定

Yahoo! JAPAN公式案内では、メールソフトで送信する場合のSMTP設定は以下です。

- 送信メール（SMTP）サーバー: `smtp.mail.yahoo.co.jp`
- 認証方式: `SMTP_AUTH`
- 通信方法: `SSL（暗号化）`
- ポート番号: `465`
- アカウント名: 使用したいメールアドレス、またはYahoo! JAPAN ID
- パスワード: Yahoo! JAPAN IDのパスワード

公式情報: [Yahoo!メールの利用に必要な環境](https://announcemail.yahoo.co.jp/info/support/)

また、Yahoo!メール側で「Yahoo! JAPAN公式サービス以外からのアクセス」を許可し、「SMTP」を利用する設定が必要です。

公式ヘルプ: [メールソフトで送受信するには](https://support.yahoo-net.jp/PccMail/s/article/H000007321)

## config.jsonの書き方

`recipient` が送信先、`sender` が差出人です。

```json
{
  "recipient": {
    "name": "山田 太郎",
    "email": "recipient@example.com"
  },
  "sender": {
    "name": "佐藤 花子",
    "email": "your-yahoo-id@yahoo.co.jp"
  },
  "smtp": {
    "host": "smtp.mail.yahoo.co.jp",
    "port": 465,
    "security": "ssl"
  },
  "auth": {
    "username": "your-yahoo-id@yahoo.co.jp",
    "password_env": "YAHOO_MAIL_PASSWORD",
    "use_windows_credential_manager": false,
    "credential_service_name": "yahoo-mail-excel-sender"
  }
}
```

`auth.password_env` は、`.env` に書くパスワード名です。通常は `YAHOO_MAIL_PASSWORD` のままで問題ありません。

## Windows資格情報マネージャーを使う場合

`.env` の代わりにWindows資格情報マネージャーを使う場合は、まずパスワードを登録します。

```powershell
python -c "import keyring; keyring.set_password('yahoo-mail-excel-sender', 'your-yahoo-id@yahoo.co.jp', 'ここにYahooメールのパスワード')"
```

次に `config.json` を変更します。

```json
{
  "auth": {
    "username": "your-yahoo-id@yahoo.co.jp",
    "password_env": "YAHOO_MAIL_PASSWORD",
    "use_windows_credential_manager": true,
    "credential_service_name": "yahoo-mail-excel-sender"
  }
}
```

## 起動方法

```powershell
python main.py
```

アプリが開いたら、以下の順に操作します。

1. 送信先を確認します。
2. 「Excelファイルを選ぶ」を押して、`.xlsx` / `.xls` / `.xlsm` のいずれかを選びます。
3. 件名と本文を確認します。
4. 「送信前に確認」を押します。
5. 内容を確認して「送信する」を押します。

送信履歴は、アプリと同じ場所の `send_log.csv` に保存されます。

## exe化する場合の手順

PyInstallerでexe化できます。

```powershell
pyinstaller --onefile --noconsole --name ExcelMailSender main.py
```

作成後、`dist` フォルダに以下を置いてください。

- `ExcelMailSender.exe`
- `config.json`
- `.env`（Windows資格情報マネージャーを使う場合は不要）

exe版でも `send_log.csv` はexeと同じ場所に作成されます。

## よくあるエラー

- 「送信用パスワードが見つかりません」  
  `.env` が `main.py` と同じ場所にあるか、`YAHOO_MAIL_PASSWORD` の名前が合っているか確認してください。

- 「メールアドレスまたはパスワードが違う可能性があります」  
  Yahoo!メールのパスワード、SMTP利用設定、`config.json` の `auth.username` を確認してください。

- 「ログイン確認中に接続が切れました」  
  Yahoo!メール設定で「Yahoo! JAPAN公式サービス以外からのアクセス」を「許可する」にし、SMTPを「利用する」にしてください。シークレットIDを使う設定の場合は、`config.json` の `auth.username` にシークレットIDを入れてください。

- 「メールサーバーにつながりませんでした」  
  インターネット接続、SMTPサーバー名、ポート番号を確認してください。

WindowsのPowerShellで、Yahoo!メールのSMTPサーバーに接続できるか確認できます。

```powershell
Test-NetConnection smtp.mail.yahoo.co.jp -Port 465
```

`TcpTestSucceeded : True` と表示されれば、パソコンからYahoo!メールのSMTPサーバーへ接続できています。

`False` の場合は、以下を確認してください。

- `config.json` の `smtp.host` が `smtp.mail.yahoo.co.jp` になっている
- `config.json` の `smtp.port` が `465` になっている
- `config.json` の `smtp.security` が `ssl` になっている
- 会社、学校、セキュリティソフト、VPNなどでメール送信用ポートが止められていない

詳しい原因は、アプリと同じ場所に作成される `error_log.txt` に記録されます。

さらに詳しく確認する場合は、次の診断を実行してください。

```powershell
python smtp_check.py
```

この診断では、以下を順番に確認します。

1. Yahoo!メールへのSSL接続
2. Yahoo!メールサーバーからの応答
3. Yahoo!メールへのログイン

`Test-NetConnection` が `True` なのにアプリで接続エラーになる場合は、`python smtp_check.py` の結果を確認してください。
