from flask import Flask, request, jsonify, send_file
import pandas as pd
import io
from openpyxl import Workbook
import smtplib
from email.message import EmailMessage
import os

app = Flask(__name__)

# === 設定（環境変数で管理） ===
GMAIL_USER = os.environ.get("GMAIL_USER")       # Gmailアドレス
GMAIL_PASS = os.environ.get("GMAIL_PASS")       # アプリパスワード
MAIL_TO     = os.environ.get("MAIL_TO")         # 宛先アドレス

# === 製品マスタの読み込み（CSV） ===
PRODUCT_MASTER_PATH = "製品マスタ.csv"

@app.route('/webhook', methods=['POST'])
def webhook():
    try:
        data = request.get_json()
        df_input = pd.DataFrame(data)  # JSON → DataFrame

        # 製品マスタの読み込み
        df_master = pd.read_csv(PRODUCT_MASTER_PATH)

        # マージして在庫判定
        df = pd.merge(df_input, df_master, on="製品名", how="left")
        df["発注要否"] = df["現在庫数"] <= df["発注点"]
        df["発注数出力"] = df.apply(lambda x: x["発注数"] if x["発注要否"] else 0, axis=1)

        # 発注が必要な製品だけ抽出
        df_order = df[df["発注要否"] == True][["製品名", "発注数出力", "発注要否", "発注先"]]
        df_order = df_order.rename(columns={"発注数出力": "発注数"})

        # Excelに出力
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_order.to_excel(writer, index=False, sheet_name="発注リスト")

        output.seek(0)

        # メール送信
        msg = EmailMessage()
        msg["Subject"] = "【自動送信】本日の発注リスト"
        msg["From"] = GMAIL_USER
        msg["To"] = MAIL_TO
        msg.set_content("以下の製品について、発注が必要です。\n添付のExcelをご確認ください。")

        msg.add_attachment(output.read(),
                           maintype="application",
                           subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           filename="発注リスト.xlsx")

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(GMAIL_USER, GMAIL_PASS)
            smtp.send_message(msg)

        return jsonify({"status": "success", "発注件数": len(df_order)})

    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500
