from flask import Flask, request, jsonify
import pandas as pd
import io
import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

app = Flask(__name__)

@app.route("/webhook", methods=["POST"])
def webhook():
    try:
        data = request.get_json()
        df_input = pd.DataFrame(data)
        df_master = pd.read_excel("製品マスタ.xlsx")

        df_merged = pd.merge(df_input, df_master, on="製品名", how="left")
        df_merged["発注要否"] = df_merged["現在庫数"] < df_merged["発注点"]
        df_merged["発注要否"] = df_merged["発注要否"].map({True: "要", False: "不要"})

        df_order = df_merged[df_merged["発注要否"] == "要"]
        output_df = df_order[["製品名", "発注数", "発注要否", "発注先"]]

        # Excel出力
        output = io.BytesIO()
        filename = f"発注リスト_{datetime.date.today()}.xlsx"
        output_df.to_excel(output, index=False)
        output.seek(0)

        # メール送信
        send_email_with_attachment(output, filename)

        return jsonify({"status": "success"}), 200
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

def send_email_with_attachment(file_bytes, filename):
    sender = "your.email@gmail.com"
    receiver = "order.manager@example.com"
    password = "your_app_password"

    msg = MIMEMultipart()
    msg["From"] = sender
    msg["To"] = receiver
    msg["Subject"] = "【自動送信】本日の発注リスト"

    msg.attach(MIMEText("以下の製品について発注が必要です。添付をご確認ください。"))

    part = MIMEApplication(file_bytes.read(), _subtype="xlsx")
    part.add_header("Content-Disposition", "attachment", filename=filename)
    msg.attach(part)

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(sender, password)
        server.send_message(msg)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
