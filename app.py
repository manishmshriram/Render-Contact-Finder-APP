import os
import re
import time
import pandas as pd
import requests
from flask import Flask, request, render_template, send_file
from werkzeug.utils import secure_filename
from bs4 import BeautifulSoup
from googlesearch import search

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

def get_company_website(company_name):
    query = f"{company_name} official site"
    try:
        for url in search(query, num=1, stop=1, pause=2):
            return url
    except Exception as e:
        print(f"Search error for {company_name}: {e}")
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
            phones = set(re.findall(r'\+?\d[\d\s\-\(\)]{7,}\d', text))
    except Exception as e:
        print(f"Error fetching {url}: {e}")
    return list(emails), list(phones)

@app.route('/')
def upload_page():
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process_file():
    if 'file' not in request.files:
        return 'No file uploaded', 400
    file = request.files['file']
    filename = secure_filename(file.filename)
    filepath = os.path.join(UPLOAD_FOLDER, filename)
    file.save(filepath)

    df = pd.read_excel(filepath)
    company_column = df.columns[0]

    df['Website'] = ''
    df['Emails'] = ''
    df['Phones'] = ''

    for i, company in enumerate(df[company_column]):
        print(f"Processing {i+1}/{len(df)}: {company}")
        website = get_company_website(company)
        df.at[i, 'Website'] = website if website else 'Not Found'
        if website:
            emails, phones = extract_contacts(website)
            df.at[i, 'Emails'] = ', '.join(emails)
            df.at[i, 'Phones'] = ', '.join(phones)
        time.sleep(2)

    output_path = os.path.join(OUTPUT_FOLDER, 'results.xlsx')
    df.to_excel(output_path, index=False)

    return send_file(output_path, as_attachment=True)

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
