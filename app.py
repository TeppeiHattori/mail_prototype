from flask import Flask, render_template, request, redirect, url_for
import imaplib
import email
import os
from datetime import datetime

app = Flask(__name__)

def download_attachments(email_address, password, start_date, end_date, save_folder):
    try:
        mail = imaplib.IMAP4_SSL('imap.gmail.com')
        mail.login(email_address, password)
        mail.select('inbox')

        # 日付範囲を指定して検索
        date_format = '%d-%b-%Y'
        start_date_str = start_date.strftime(date_format)
        end_date_str = end_date.strftime(date_format)
        search_criteria = f'(SINCE "{start_date_str}" BEFORE "{end_date_str}")'
        
        status, data = mail.search(None, search_criteria)
        if status != 'OK':
            return "No messages found!"

        mail_ids = data[0].split()

        # フォルダが存在しない場合は作成
        if not os.path.exists(save_folder):
            os.makedirs(save_folder)

        downloaded_files = []

        for mail_id in mail_ids:
            status, msg_data = mail.fetch(mail_id, '(RFC822)')
            if status != 'OK':
                continue

            msg = email.message_from_bytes(msg_data[0][1])
            for part in msg.walk():
                if part.get_content_maintype() == 'multipart':
                    continue
                if part.get('Content-Disposition') is None:
                    continue
                file_name = part.get_filename()
                if file_name:
                    file_path = os.path.join(save_folder, file_name)
                    with open(file_path, 'wb') as f:
                        f.write(part.get_payload(decode=True))
                    downloaded_files.append(file_name)

        mail.logout()

        if downloaded_files:
            return f"Downloaded files: {', '.join(downloaded_files)}"
        else:
            return "No attachments found."
    except Exception as e:
        return f"An error occurred: {e}"

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/download', methods=['POST'])
def download():
    email_address = request.form['email']
    password = request.form['password']
    start_date = request.form['start_date']
    end_date = request.form['end_date']
    folder_name = request.form['folder_name']

    try:
        start_date = datetime.strptime(start_date, '%Y-%m-%d')
        end_date = datetime.strptime(end_date, '%Y-%m-%d')
    except ValueError as e:
        return render_template('index.html', message=f"Invalid date format: {e}")

    # デスクトップパスの取得
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    save_folder = os.path.join(desktop_path, folder_name)

    message = download_attachments(email_address, password, start_date, end_date, save_folder)
    return render_template('index.html', message=message)

if __name__ == '__main__':
    app.run(debug=True)
