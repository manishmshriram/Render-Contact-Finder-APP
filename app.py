import os
import pandas as pd
import re
import time
import requests
from bs4 import BeautifulSoup
from googlesearch import search
from flask import Flask, request, render_template, send_file
from werkzeug.utils import secure_filename

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

@app.route('/')
def home():
    return render_template('index.html')  # A simple upload form

@app.route('/upload', methods=['POST'])
def upload_file():
    file = request.files['file']
    filename = secure_filename(file.filename)
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(filepath)

    df = pd.read_excel(filepath)
    company_column = df.columns[0]
    df['Website'] = ''
    df['Emails'] = ''
    df['Phones'] = ''

    for i, company in enumerate(df[company_column]):
        website = get_company_website(company)
        df.at[i, 'Website'] = website if website else 'Not Found'
        if website:
            emails, phones = extract_contacts(website)
            df.at[i, 'Emails'] = ', '.join(emails)
            df.at[i, 'Phones'] = ', '.join(phones)
        time.sleep(2)

    output_path = os.path.join(app.config['UPLOAD_FOLDER'], "Company_Contacts_Results.xlsx")
    df.to_excel(output_path, index=False)
    return send_file(output_path, as_attachment=True)

def get_company_website(company_name):
    query = f"{company_name} official site"
    for url in search(query, num=1, stop=1, pause=2):
        return url
    return None

def extract_contacts(url):
    emails, phones = set(), set()
    try:
        headers = {'User-Agent': 'Mozilla/5.0'}
        response = requests.get(url, headers=headers, timeout=5)
        if response.status_code == 200:
            soup = BeautifulSoup(response.text, 'html.parser')
            text = soup.get_text()
            emails = set(re.findall(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', text))
            phones = set(re.findall(r'\+?\d{1,4}?[\s.-]?\(?\d{2,4}\)?[\s.-]?\d{3,4}[\s.-]?\d{3,4}', text))
    except Exception as e:
        print(f"Error fetching {url}: {e}")
    return list(emails), list(phones)

if __name__ == "__main__":
    port = int(os.environ.get('PORT', 10000))
    app.run(host="0.0.0.0", port=port)
