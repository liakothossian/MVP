import concurrent.futures
from flask import Flask, render_template, request, redirect, flash, send_from_directory
import pandas as pd
import json
import requests
from werkzeug.utils import secure_filename
import os
from fake_useragent import UserAgent
from bs4 import BeautifulSoup
from tqdm import tqdm
import warnings
import uuid
from pathlib import Path

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = str(Path.home())  # Set upload folder to the user's home directory


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ['csv', 'xlsx', 'json']


@app.route('/', methods=['GET', 'POST'])
def myform():
    if request.method == 'POST':
        if 'x' not in request.files:
            flash('No file part')
            return redirect(request.url)
        file = request.files['x']
        if file.filename == '':
            flash('No selected file')
            return redirect(request.url)
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(file_path)

            if filename.lower().endswith('.csv'):
                df = pd.read_csv(file_path, encoding='utf-8', decimal=',')
            elif filename.lower().endswith('.xlsx'):
                df = pd.read_excel(file_path)
            elif filename.lower().endswith('.json'):
                with open(file_path, 'r') as json_file:
                    data = json.load(json_file)
                df = pd.DataFrame(data)

            responsess = []
            mail_validation = {}
            ua = UserAgent()
            with concurrent.futures.ThreadPoolExecutor(max_workers=10) as executor:
                futures = []
                for i, row in df.iterrows():
                    email = str(row.get('DirectEmail', ''))
                    link = row.get('Source', '')
                    if pd.isna(email) or pd.isna(link):
                        mail_validation[i] = 0
                        responsess.append('')
                        continue
                    headers = {'User-Agent': ua.random}
                    futures.append(executor.submit(requests.get, link, headers=headers))
                for i, future in enumerate(tqdm(futures)):
                    try:
                        response = future.result()
                        if isinstance(response, requests.Response):
                            if response.status_code == 404:
                                mail_validation[i] = 0
                            else:
                                with warnings.catch_warnings():
                                    warnings.filterwarnings("ignore", category=Warning)
                                    soup = BeautifulSoup(response.content, 'html5lib')
                                text = soup.get_text()
                                if '@' in text:
                                    mail_validation[i] = 1
                                else:
                                    mail_validation[i] = 0
                            responsess.append(response)
                        elif isinstance(response, str):
                            mail_validation[i] = -1
                            responsess.append("Request Error")
                        else:
                            mail_validation[i] = 0
                            responsess.append('')
                    except requests.exceptions.RequestException:
                        mail_validation[i] = -1
                        responsess.append("Request Error")

            df["valid_email"] = pd.Series(mail_validation)
            df["Response_Type"] = pd.Series(responsess)

            # Create a new CSV file for writing the summary
            summary_filename = str(uuid.uuid4()) + '_summary.csv'
            summary_output_path = os.path.join(app.config['UPLOAD_FOLDER'], summary_filename)

            # Write the original DataFrame to the CSV summary file
            df.to_csv(summary_output_path, index=False)

            summary = {
                "filename": summary_filename,
                "total": len(df),
                "valid_emails": len(df.loc[df['valid_email'] == 1]),
                "invalid_emails": len(df.loc[df['valid_email'] == 0]),
                "request_errors": len(df.loc[df['valid_email'] == -1]),
            }
            return render_template('summary.html', summary=summary, download_link=summary_filename)
    return render_template('index.html')


@app.route('/download/<filename>')
def download_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename, as_attachment=True, mimetype='text/csv')


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
