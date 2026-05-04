import os
from datetime import datetime
from flask import Flask, render_template, redirect, url_for ,request,send_file
import sqlite3 
import pandas as pd 
app = Flask(__name__)   

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATABASE_PATH = os.path.join(BASE_DIR, "database", "students.db")

def create_table():
    conn = sqlite3.connect(DATABASE_PATH)
    conn.execute("PRAGMA journal_mode=WAL;")
    cursor = conn.cursor()
    cursor.execute(""" CREATE TABLE IF NOT EXISTS students
                ( id INTEGER PRIMARY KEY AUTOINCREMENT,
                စဉ် TEXT,
                ကျောင်းဝင်အမှတ် TEXT,
                နာမည် TEXT,
                အဖေနာမည် TEXT,
                ကျားမ TEXT,
                မွေးနေ့ TEXT,
                class TEXT,
                UNIQUE(ကျောင်းဝင်အမှတ်) ) """)
    
    conn.commit()
    conn.close()

create_table()

def calculate_age(dob_str):
    dob = datetime.strptime(dob_str, "%d-%m-%Y")
    today = datetime.today()
    age = today.year - dob.year - ((today.month, today.day) < (dob.month, dob.day))
    return age

def apply_filters(data_age, class_html, gender_html, age_html, age_html_1, age_html_2):
    data = data_age
    if class_html != 'all':
        data = [row for row in data if row[6] == class_html]
    if gender_html != 'all':
        data = [row for row in data if row[3] == gender_html]
    if age_html:
        data = [row for row in data if int(row[7]) == int(age_html)]
    elif age_html_1 and age_html_2:
        data = [row for row in data if int(age_html_1) <= int(row[7]) <= int(age_html_2)]
    return data

def eng_to_mm(number):
    eng_digits = ['0','1','2','3','4','5','6','7','8','9']
    mm_digits = ['၀','၁','၂','၃','၄','၅','၆','၇','၈','၉']
    number_list = []
    for num in number:
        for eng,mm in zip(eng_digits,mm_digits):
            if num == eng:
                number_list.append(mm)

    return ''.join(number_list)

@app.route('/')
def home():
    return render_template('try.html')

@app.route('/insert', methods=['POST'])
def insert_data():
    excel_file = request.files['myfile']
    with sqlite3.connect(DATABASE_PATH, timeout=15) as conn:
        cursor = conn.cursor()
    
        cursor.execute("DELETE FROM students")
    
        df = pd.read_excel(excel_file,dtype={'စဉ်':str,'ကျောင်းဝင်အမှတ်':str})
        df['စဉ်'] = df['စဉ်'].apply(eng_to_mm)
        df['ကျောင်းဝင်အမှတ်'] = df['ကျောင်းဝင်အမှတ်'].apply(eng_to_mm)
        df.columns = df.columns.str.strip()
    
        for _, row in df.iterrows():
            cursor.execute(""" INSERT OR IGNORE INTO students 
            (စဉ်,ကျောင်းဝင်အမှတ်,နာမည်,ကျားမ,အဖေနာမည်,မွေးနေ့,class) 
            VALUES (?, ?, ?, ?, ?, ?, ?) """,
           (row['စဉ်'],
            row['ကျောင်းဝင်အမှတ်'],
            row['နာမည်'],
            row['ကျားမ'],
            row['အဖေနာမည်'],
            pd.to_datetime(row['မွေးနေ့']).strftime('%d-%m-%Y'),
            row['class']))
    return redirect(url_for('home'))

@app.route('/view_tb', methods=['POST'])
def view_data():
    age_html = request.form.get('age')
    age_html_1 = request.form.get('age_1')
    age_html_2 = request.form.get('age_2')
    class_html = request.form.get('class')
    gender_html = request.form.get('gender')
    filename   = request.form.get('file_name')
    action_html = request.form.get('on')

    conn = sqlite3.connect(DATABASE_PATH, timeout=10)
    cursor = conn.cursor()
    cursor.execute("SELECT စဉ်, ကျောင်းဝင်အမှတ်, နာမည်, ကျားမ, အဖေနာမည်, မွေးနေ့ ,class FROM students")
    rows = cursor.fetchall()
    conn.close()
    
    data_age = []
    for row in rows:
        no, roll, name, gender, father, dob, class_name = row
        age = calculate_age(dob)
        data_age.append((no, roll, name, gender, father, dob, class_name, age))

    data = apply_filters(data_age, class_html, gender_html, age_html, age_html_1, age_html_2)

    if action_html == 'run':
        return render_template('try.html', data=data,
                               data_male=len([r for r in data if r[3] == 'ကျား']),
                               data_female=len([r for r in data if r[3] == 'မ']),
                               data_all=len(data),
                               age=age_html, age_1=age_html_1, age_2=age_html_2,
                               class_html=class_html, gender=gender_html)
    elif action_html == 'excel_download':
        download_data = [
            {
                'စဉ်': r[0],
                'ကျောင်းဝင်အမှတ်': r[1],
                'နာမည်': r[2],
                'ကျားမ': r[3],
                'အဖေနာမည်': r[4],
                'မွေးနေ့': r[5],
                'class': r[6],
                'အသက်': r[7]
            }for r in data 
        ]
        df = pd.DataFrame(download_data)
        df.to_excel(filename+'.xlsx', index=False)
        return send_file(filename+'.xlsx', as_attachment=True)

    return render_template('try.html')

if __name__ == '__main__':
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
