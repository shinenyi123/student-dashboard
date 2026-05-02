import os
from tarfile import data_filter

from flask import Flask, render_template, redirect, url_for ,request
import sqlite3 
import pandas as pd 
app = Flask(__name__)   

EXCEL_PATH_OUTPUT = 'C:/Users/Shine Nyi/Desktop/code/Python/Python_Real_Life_Project/python_UI/python_UI_student/excel_output/' 
DATABASE_PATH     = 'C:/Users/Shine Nyi/Desktop/code/Python/Python_Real_Life_Project/python_UI/python_UI_student/database/students.db' 

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
from datetime import datetime

def calculate_age(dob_str):
    dob = datetime.strptime(dob_str, "%Y-%m-%d")
    today = datetime.today()

    age = today.year - dob.year - (
        (today.month, today.day) < (dob.month, dob.day)
    )

    return age

def apply_filters(data_age, class_html, gender_html, age_html, age_html_1, age_html_2):

    if class_html == 'all' and gender_html == 'all':
        data = data_age

    if class_html != 'all':
        data = [row for row in data_age if row[6] == class_html]

    if gender_html != 'all':
        data = [row for row in data_age if row[3] == gender_html]

    if age_html:
        data = [row for row in data_age if int(row[7]) == int(age_html)]

    elif age_html_1 and age_html_2:
        data = [row for row in data_age if int(age_html_1) <= int(row[7]) <= int(age_html_2)]

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
    for index, row in df.iterrows(): 
        no        = row['စဉ်']
        school_no = row['ကျောင်းဝင်အမှတ်']
        name      = row['နံမည်'] 
        gender    = row['ကျားမ']
        f_name    = row['အဖေနံမည်']
        birthday  = row['မွေးနေ့'] 
        birthday = pd.to_datetime(birthday)
        birthday = birthday.strftime('%Y-%m-%d')
        class_name = row['class']
        cursor.execute(""" INSERT OR IGNORE INTO students (စဉ်,ကျောင်းဝင်အမှတ်,နံမည်,ကျားမ,အဖေနံမည်,မွေးနေ့,class) 
                       VALUES (?, ?, ?, ?, ?, ?, ?) """, 
                       (no, school_no, name,gender, f_name, birthday, class_name)) 
    conn.commit() 
    conn.close() 
        
    return redirect(url_for('home')) 

###########################################################################################################

@app.route('/view_tb', methods=['POST'])
def view_data():
    if request.method == 'POST':
        age_html = request.form.get('age')
        age_html_1 = request.form.get('age_1')
        age_html_2 = request.form.get('age_2')
        class_html = request.form.get('class')
        gender_html = request.form.get('gender')
        filename   = request.form.get('file_name')
        action_html = request.form.get('on')

        if action_html == 'run':
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
            return render_template('try.html', data=data,
                                        data_male = len([r for r in data if r[3] == 'ကျား']),
                                        data_female = len([r for r in data if r[3] == 'မ']),
                                        data_all = len(data),
                                        age=age_html,
                                        age_1=age_html_1,
                                        age_2=age_html_2,
                                        class_html=class_html,
                                        gender=gender_html)

        elif action_html == 'save_file':
            conn = sqlite3.connect(DATABASE_PATH)
            cursor = conn.cursor()
            cursor.execute("SELECT စဉ်, ကျောင်းဝင်အမှတ်, နံမည်, ကျားမ, အဖေနံမည်, မွေးနေ့ ,class FROM students")
            rows = cursor.fetchall()
            conn.close()

            data_age = []
            for row in rows:
                no, roll, name,gender, father, dob, class_name = row
                age = calculate_age(dob)
                data_age.append((no, roll, name, gender, father, dob, class_name, age))

            data = apply_filters(data_age, class_html, gender_html, age_html, age_html_1, age_html_2)

            no, school_no, name, gender, f_name, b_day, age = [], [], [], [], [], [], []

            for row in data:
                no.append(row[0])
                school_no.append(row[1])
                name.append(row[2])
                gender.append(row[3])
                f_name.append(row[4])
                b_day.append(row[5])
                age.append(row[7])

            data = {
                'စဉ်': no,
                'ကျောင်းဝင်အမှတ်': school_no,
                'နံမည်': name,
                'ကျားမ': gender,
                'အဖေနံမည်': f_name,
                'မွေးနေ့': b_day,
                'အသက်': age
            }

            df = pd.DataFrame(data)
            df.to_excel(EXCEL_PATH_OUTPUT + filename + '.xlsx', index=False)       
            
            return render_template('try.html', output_py = 'Suc2')

    return render_template('try.html')
###########################################################################################################
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
