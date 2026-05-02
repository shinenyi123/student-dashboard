from tarfile import data_filter
import os
from datetime import datetime
from flask import Flask, render_template, redirect, url_for ,request
import sqlite3 
import pandas as pd 
app = Flask(__name__)   

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATABASE_PATH = os.path.join(BASE_DIR, "database", "students.db")

def create_table():
    conn = sqlite3.connect(DATABASE_PATH)
    cursor = conn.cursor()
    cursor.execute(""" CREATE TABLE IF NOT EXISTS students
                ( id INTEGER PRIMARY KEY AUTOINCREMENT,
                စဉ် TEXT,
                ကျောင်းဝင်အမှတ် INTEGER,
                နံမည် TEXT,
                အဖေနံမည် TEXT,
                ကျားမ TEXT,
                မွေးနေ့ TEXT,
                class TEXT,
                UNIQUE(ကျောင်းဝင်အမှတ်) ) """)
    conn.commit()
    conn.close()

create_table()

def calculate_age(dob_str):
    dob = datetime.strptime(dob_str, "%Y-%m-%d")
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

@app.route('/')
def home():
    return render_template('try.html')

@app.route('/insert', methods=['POST'])
def insert_data():
    excel_file = request.files['myfile']
    conn = sqlite3.connect(DATABASE_PATH)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM students")
    df = pd.read_excel(excel_file)
    df.columns = df.columns.str.strip()
    for _, row in df.iterrows():
        no        = row['စဉ်']
        school_no = row['ကျောင်းဝင်အမှတ်']
        name      = row['နံမည်']
        gender    = row['ကျားမ']
        f_name    = row['အဖေနံမည်']
        birthday  = pd.to_datetime(row['မွေးနေ့']).strftime('%Y-%m-%d')
        class_name = row['class']
        cursor.execute(""" INSERT OR IGNORE INTO students 
            (စဉ်,ကျောင်းဝင်အမှတ်,နံမည်,ကျားမ,အဖေနံမည်,မွေးနေ့,class) 
            VALUES (?, ?, ?, ?, ?, ?, ?) """,
            (no, school_no, name, gender, f_name, birthday, class_name))
    conn.commit()
    conn.close()
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

    conn = sqlite3.connect(DATABASE_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT စဉ်, ကျောင်းဝင်အမှတ်, နံမည်, ကျားမ, အဖေနံမည်, မွေးနေ့ ,class FROM students")
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

    return render_template('try.html')

if __name__ == '__main__':
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
