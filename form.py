from flask import Flask, render_template, request, redirect, url_for
import openpyxl
from openpyxl import Workbook

app = Flask(__name__)

# Save data to Excel
def save_to_excel(name, email, age, birthdate, gender, branch, mis, year):
    try:
        workbook = openpyxl.load_workbook("user_data.xlsx")
        sheet = workbook.active
    except FileNotFoundError:
        workbook = Workbook()
        sheet = workbook.active
        # Column headers
        sheet.append(["Name", "Email", "Age", "Birthdate", "Gender", "Branch", "MIS Number", "Year of Study"])

    # Append new data
    sheet.append([name, email, age, birthdate, gender, branch, mis, year])
    workbook.save("user_data.xlsx")

@app.route('/')
def index():
    return render_template('form.html')

@app.route('/submit', methods=['POST'])
def submit():
    name = request.form['name']
    email = request.form['email']
    age = request.form['age']
    birthdate = request.form['birthdate']
    gender = request.form['gender']
    branch = request.form['branch']
    mis = request.form['mis']
    year = request.form['year']

    if name and email and age and birthdate and gender and branch and mis and year:
        save_to_excel(name, email, age, birthdate, gender, branch, mis, year)
        return redirect(url_for('index'))
    
    return 'Please fill all the fields!'

if __name__ == '__main__':
    app.run(debug=True)
