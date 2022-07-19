from flask import Flask, render_template, request, flash
from flask_mail import Mail,Message
import pandas as pd
from werkzeug.utils import secure_filename
import os
import json

app = Flask(__name__)

file_location= open('config.json')
data = json.load(file_location)

app.config['UPLOAD_FOLDER'] = str(data['location'])

app.secret_key = b'Siddharth'

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/bulk', methods = ['GET', 'POST'])
def bulk():
    return render_template('bulk.html')

@app.route('/send_message', methods = ['GET', 'POST'])
def send_message():
    try:
        if request.method == 'POST':

            user_email = request.form['user_email']
            password = request.form['password']

            app.config['MAIL_SERVER'] = 'smtp.gmail.com'
            app.config['MAIL_PORT'] = 465
            app.config['MAIL_USERNAME'] = user_email
            app.config['MAIL_PASSWORD'] = password
            app.config['MAIL_USE_TLS'] = False
            app.config['MAIL_USE_SSL'] = True

            mail = Mail(app)

            email = request.form['email']
            msg = request.form['message']
            subject = request.form['subject']

            message = Message(subject,sender=str(user_email),recipients=[email])
            message.body = msg
            mail.send(message)
            flash("Done! Your message has been sent","success")
            return render_template('index.html')

    except:
        return render_template('index.html')

@app.route('/bulk_send_message', methods = ['GET', 'POST'])
def bulk_send_message():

    if request.method == 'POST':

        user_email = request.form['user_email']
        password = request.form['password']

        app.config['MAIL_SERVER'] = 'smtp.gmail.com'
        app.config['MAIL_PORT'] = 465
        app.config['MAIL_USERNAME'] = user_email
        app.config['MAIL_PASSWORD'] = password
        app.config['MAIL_USE_TLS'] = False
        app.config['MAIL_USE_SSL'] = True

        mail = Mail(app)

        f = request.files["file1"]
        f.save(os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(f.filename)))

        data = pd.read_excel(f)
        emails = list(data['Email'])

        c = []
        for i in emails:
            if pd.isnull(i) == False:
                c.append(i)

        emails = c

        msg = request.form['message']
        subject = request.form['subject']

        for i in emails:
            message = Message(subject,sender=str(user_email),recipients=[str(i)])
            message.body = msg
            mail.send(message)

        file_location= open('config.json')
        data = json.load(file_location)
            
        for j in os.listdir(data['location']):
            os.remove(os.path.join(data['location'],j))

        flash("Done! Your message has been sent","success")
    
    return render_template('bulk.html')

if __name__ == "__main__":
    app.run(host="localhost", debug=True, port=8000)