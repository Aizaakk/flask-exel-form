from flask import Flask, render_template, request
from openpyxl import load_workbook, Workbook
import os

app = Flask(__name__)

excel_file = 'data.xlsx'

if not os.path.exists(excel_file):
    wb = Workbook()
    ws = wb.active
    ws.append(['Name', 'Email'])  # Header row
    wb.save(excel_file)

@app.route('/')
def form():
    wb = load_workbook(excel_file)
    ws = wb.active

    people = [
        (row[0].value, row[1].value)
        for row in ws.iter_rows(min_row=2, max_col=2)
        if row[0].value and row[1].value
    ]

    return render_template('form.html', people=people)

@app.route('/submit', methods=['POST'])
def submit():
    selected = request.form.get('person')       # value from dropdown
    new_name = request.form.get('new_name')     # new name input
    new_email = request.form.get('new_email')   # new email input

    if selected:
        name, email = selected.split('|')
    elif new_name and new_email:
        name = new_name
        email = new_email
    else:
        return "Please select a person or enter a new name and email."

    # Save to Excel
    wb = load_workbook(excel_file)
    ws = wb.active
    ws.append([name, email])
    wb.save(excel_file)

    return f"Saved: {name} ({email})"

if __name__ == '__main__':
    app.run(debug=True)