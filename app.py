from flask import Flask, render_template, request, redirect, url_for, flash, send_from_directory
import pandas as pd
import requests
import requests.exceptions
from werkzeug.utils import secure_filename
import os
import random
from fake_useragent import UserAgent
from tqdm import tqdm
from bs4 import BeautifulSoup
import fitz

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = './'

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ['xlsx', 'xls']

def extract_text_from_pdf(url):
    try:
        headers = {'User-Agent': UserAgent().random}
        r = requests.get(url, headers=headers)
        with fitz.open(stream=r.content, filetype="pdf") as doc:
            text = ""
            for page in doc:
                text += page.get_text()
            return text
    except:
        return ""

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
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            df = pd.read_excel(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            responsess = []
            mail_validation = []
            ua = UserAgent()

            for i in tqdm(range(len(df))):
                email = str(df['DirectEmail'][i])
                link = df['Source'][i]
                if pd.isna(email) or pd.isna(link):
                    mail_validation.append(0)
                    responsess.append('')
                    continue
                try:
                    if link.endswith('.pdf'):
                        text = extract_text_from_pdf(link)
                    else:
                        headers = {'User-Agent': ua.random}
                        response = requests.get(link, headers=headers)
                        soup = BeautifulSoup(response.content, 'html.parser')
                        text = soup.get_text()
                    if '@' in text:
                        mail_validation.append(1)
                    else:
                        mail_validation.append(0)
                    responsess.append(response)
                except requests.exceptions.RequestException:
                    responsess.append("Request Error")
                    mail_validation.append(-1)

            df["valid_email"] = mail_validation
            df["Response_Type"] = responsess
            filename1 = 'Outputfile.xlsx'
            df.to_excel(filename1)
            summary = {
                "total": len(df),
                "valid_emails": len(df.loc[df['valid_email'] == 1]),
                "invalid_emails": len(df.loc[df['valid_email'] == 0]),
                "request_errors": len(df.loc[df['valid_email'] == -1]),
            }
            return render_template('summary.html', summary=summary)
    return render_template('index.html')

@app.route('/download')
def download_file():
    return send_from_directory(app.config['UPLOAD_FOLDER'], 'Outputfile.xlsx', as_attachment=True)

if __name__ == '__main__':
    app.run(port=1234, debug=True)
