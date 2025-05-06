from flask import Flask, request, jsonify
import pandas as pd
from openpyxl import Workbook
import io
import smtplib
from email.message import EmailMessage
from datetime import datetime

app = Flask(__name__)

# === 設定項目 ===
PRODUCT_MASTER_PATH = "製品マスタ.xlsx"
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
GMAIL_USER = "your_email@gmail.com"  # ← ご自身のGmailに変更
GMAIL_PASS = "your_app_password"     # ← アプリパスワードを使用（推奨）
TO_ADDRESS = "receiver@example.com"  # ← 送信先のメールアドレス

@app.route('/')
def hello():
    return 'Flask server is running'

@app.route('/webhook', methods=['POST'])
def webhook():
    try:
        data = request.get_json()
        print("受信データ:", data)

        if not isinstance(data, list):
            raise ValueError("データ形式はリスト（JSON array）である必要があります")

        df_input = pd.DataFrame(data)
        df_master = pd.read_excel(PRODUCT_MASTER_PATH)

        merged = pd.merge(df_input, df_master, on="製品名", how="left")
        merged["発注要否"] = merged["現在庫数"] < merged["発注点"]
        merged = merged[merged["発注要否"] == True]

        if merged.empty:
            print("発注対象なし")
            return jsonify({"status": "ok", "message": "発注対象なし"})

        output_df = merged[["製品名", "発注数", "発注要否", "発注先"]]
        filename = f"発注リスト_{datetime.now().strftime('%Y-%m-%d')}.xlsx"

        with io.BytesIO() as buffer:
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                output_df.to_excel(writer, index=False, sheet_name="発注一覧")
            buffer.seek(0)

            # メール送信
            msg = EmailMessage()
            msg["Subject"] = "【自動送信】本日の発注リスト"
            msg["From"] = GMAIL_USER
            msg["To"] = TO_ADDRESS
            msg.set_content("以下の製品について、発注が必要です。\n添付のExcelをご確認ください。")
            msg.add_attachment(buffer.read(), maintype="application", subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename=filename)

            with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as smtp:
                smtp.starttls()
                smtp.login(GMAIL_USER, GMAIL_PASS)
                smtp.send_message(msg)

        print("メール送信完了")
        return jsonify({"status": "success", "message": "メール送信完了"})

    except Exception as e:
        print("エラー内容:", e)
        return jsonify({"status": "error", "message": str(e)}), 500

if __name__ == '__main__':
    app.run()
