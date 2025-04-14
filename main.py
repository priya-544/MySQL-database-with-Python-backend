from flask import Flask, render_template, request, redirect, url_for, flash, session,send_file, jsonify
from werkzeug.utils import secure_filename
import os
import pandas as pd
import datetime
import mysql.connector
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import Mail, Attachment, FileContent, FileName, FileType, Disposition
import base64
import re
from apscheduler.schedulers.background import BackgroundScheduler
import random
import smtplib
from sendgrid.helpers.mail import To
import datetime

# Flask App Configuration
app = Flask(__name__)
app.secret_key = 'secret_key'
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# MySQL Database Connection
db = mysql.connector.connect(
    host='localhost',
    user='username',
    password='db_password',
    database='db_name'
)
cursor = db.cursor()

# Initialize APScheduler
scheduler = BackgroundScheduler()
scheduler.start()

def is_valid_email(email):
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return re.match(pattern, email) is not None

@app.route('/', methods=['GET', 'POST'])
def home():
    if 'user_id' not in session:
        flash('Please log in to access the home page.', 'error')
        return render_template('auth.html')

    user_name = session.get('user_name', 'User')  # Default to 'User' if name is not in session
    selected_choice = session.get('selected_choice', None)

    if request.method == 'POST':
        # Handle the dropdown choice
        if 'submit_choice' in request.form:
            choice = request.form.get('choice')
            if choice:
                session['selected_choice'] = choice
                # session['popup_shown'] = True  # Mark popup as shown
                flash(f"You selected: {choice}", 'success')
            else:
                flash('Please select an option before proceeding.', 'error')
            return redirect(url_for('home'))

        # Handle the file upload
        elif 'upload_file' in request.form:
            selected_choice = session.get('selected_choice')
            # if not selected_choice:
            #     flash('Please select an option before uploading a file.', 'error')
            #     return redirect(url_for('home'))

            # file = request.files['file']
            file = request.files.get('file')
            if file and file.filename.endswith('.xlsx'):
                filename = secure_filename(file.filename)
                filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                file.save(filepath)
                try:
                    df = pd.read_excel(filepath)

                    if "Email" not in df.columns:
                        flash("The uploaded file must contain an 'Email' column.")
                        return redirect(url_for('home'))

                    df["valid_email"] = df["Email"].apply(lambda x: "Valid" if is_valid_email(x) else "Invalid")
                    
                    preview_data = df[["Email", "Name", "valid_email"]].to_dict(orient='records')
                    session['preview_data'] = preview_data
                    first_valid_email = next((row["Email"] for row in preview_data if row["valid_email"] == "Valid"), "")

                    invalid_emails = df[df["valid_email"] == "Invalid"]["Email"].tolist()
                    if invalid_emails:
                        with open("invalid_emails.log", "w") as log_file:
                            log_file.write("\n".join(invalid_emails))

                    df.to_excel(filepath, index=False)

                    session['uploaded_file'] = filename  # Store file in session
                    cc = request.form.get('cc', '').strip()
                    bcc = request.form.get('bcc', '').strip()
                    cc_list = [email.strip() for email in cc.split(',') if email.strip()]
                    bcc_list = [email.strip() for email in bcc.split(',') if email.strip()]
                    session['manual_cc'] = cc_list
                    session['manual_bcc'] = bcc_list

                    flash('File uploaded successfully!', 'success')

                    # return redirect(url_for('home'))

                    return redirect(url_for('add_subject_body', filepath=filename, first_email=first_valid_email))

                except Exception as e:
                    flash(f"Error processing file: {e}")
                    return redirect(url_for('home'))
            
            else:
                flash('Invalid file format. Please upload an Excel file.', 'error')

        # Handle manual email input
        elif 'manual_submit' in request.form:
            emails = request.form.get('manual_emails', '').strip()
            names = request.form.get('manual_names', '').strip()
            cc = request.form.get('cc', '').strip()
            bcc = request.form.get('bcc', '').strip()
            if not emails:
                flash('Please enter at least one email address.', 'error')
                return redirect(url_for('home'))

            # Convert comma-separated values into lists
            email_list = [email.strip() for email in emails.split(',') if email.strip()]
            name_list = [name.strip() for name in names.split(',')] if names else ['Unknown'] * len(email_list)
            cc_list = [email.strip() for email in cc.split(',') if email.strip()]
            bcc_list = [email.strip() for email in bcc.split(',') if email.strip()]
            if len(email_list) != len(name_list):
                flash('The number of names should match the number of emails or be left empty.', 'error')
                return redirect(url_for('home'))

            # Store manually entered data in session
            session['manual_emails'] = email_list
            session['manual_names'] = name_list
            session['manual_cc'] = cc_list
            session['manual_bcc'] = bcc_list

            preview_data = [{"Email": email, "Name": name, "valid_email": "Valid" if is_valid_email(email) else "Invalid"} for email, name in zip(email_list, name_list)]
            session['preview_data'] = preview_data
            first_valid_email = next((row["Email"] for row in preview_data if row["valid_email"] == "Valid"), "")


            flash('Emails added successfully!', 'success')
            
            # return redirect(url_for('home'))
            return redirect(url_for('add_subject_body', filepath='manual', first_email=first_valid_email))

        # Handle resetting the choice
        elif 'reset_choice' in request.form:
            session.pop('selected_choice', None)
            session.pop('uploaded_file', None)
            session.pop('manual_emails', None)
            session.pop('manual_names', None)
            session.pop('manual_cc', None)
            session.pop('manual_bcc', None)
            session.clear()
            flash('You can now choose a different option.', 'info')
            return redirect(url_for('home'))
    preview_data = session.get('preview_data', [])
    return render_template('home.html', user_name=user_name, selected_choice=selected_choice, preview_data=preview_data)

@app.route('/download-template')
def download_template():
    # Path to your template file
    file_path = "./template.xlsx"
    
    # Sending the file as a response
    return send_file(file_path, as_attachment=True, download_name="template.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.route("/download-log", methods=["GET"])
def download_log():
    log_file_path = "invalid_emails.log"
    if os.path.exists(log_file_path):
        return send_file(log_file_path, as_attachment=True)
    else:
        flash("No log file found.")
        return redirect("/")

@app.route('/preview', methods=['POST'])
def preview():
    file = request.files.get('file')
    if file and file.filename.endswith('.xlsx'):
        df = pd.read_excel(file)
        if "Email" not in df.columns:
            return jsonify({"error": "Missing 'Email' column"}), 400
        
        # Add status based on validation
        df["Status"] = df["Email"].apply(lambda x: "Valid" if is_valid_email(x) else "Invalid")

        # Include 'Status' in the preview response
        preview_data = df[["Email", "Name", "Status"]].to_dict(orient='records')
        return jsonify({"preview": preview_data})
    return jsonify({"error": "Invalid file"}), 400

@app.route('/add_subject_body/<filepath>', methods=['GET', 'POST'])
def add_subject_body(filepath):
    selected_choice = session.get('selected_choice', None)  # Get the user's selected choice

    if request.method == 'POST':
        subject = request.form['subject']
        body = request.form['body']
        attachment = request.files['attachment']
        schedule_time = request.form['schedule_time']
        sender_email = request.form['sender_email']

        # Check if it's a meeting request and add meeting details
        meeting_date = request.form.get('meeting_date')
        meeting_duration = request.form.get('meeting_duration')

        # Get CC and BCC
        cc_list = session.get('manual_cc', [])
        bcc_list = session.get('manual_bcc', [])

        # If manually entered emails are used
        if filepath == 'manual':
            valid_emails = session.get('manual_emails', [])
            names = session.get('manual_names', [])
        else:
            # Process Excel File
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filepath)
            df = pd.read_excel(file_path)
            valid_emails = df['Email'].dropna().tolist()
            names = df['Name'].fillna('Unknown').tolist()

        total_emails = len(valid_emails)

        # Determine activity name
        if selected_choice=="Meeting Request":
            activity_name = 'Meeting Request Scheduled' if schedule_time else 'Meeting Request Sent'
        else:
            activity_name = 'Email Scheduled' if schedule_time else 'Email Sent'

        # Schedule or Send Emails Immediately
        if schedule_time:
            activity_name = 'Email Scheduled'
            schedule_datetime = datetime.datetime.strptime(schedule_time, '%Y-%m-%dT%H:%M')
            scheduler.add_job(
                send_emails,
                'date',
                run_date=schedule_datetime,
                args=[sender_email, valid_emails, names, subject, body, attachment, meeting_date, meeting_duration,cc_list, bcc_list]
            )
        else:
            # activity_name = 'Email Sent'
            send_emails(sender_email, valid_emails, names, subject, body, attachment, meeting_date, meeting_duration,cc_list, bcc_list)

        # Log Activity
        cursor.execute(
            """
            INSERT INTO Activity (date_of_activity, time_of_activity, activity_name, performed_by, number_of_emails)
            VALUES (%s, %s, %s, %s, %s)
            """,
            (datetime.datetime.now().date(), datetime.datetime.now().time(), activity_name, sender_email, total_emails)
        )
        db.commit()

        flash(f'{activity_name.lower()} successfully!', 'success')
        return redirect(url_for('home'))
    preview_data = session.get('preview_data', [])
    first_email = next((row["Email"] for row in preview_data if row["valid_email"] == "Valid"), "") if preview_data else ""


    return render_template('home.html', filepath=filepath, selected_choice=selected_choice,first_email=first_email)

    # return render_template('add_subject_body.html', filepath=filepath, selected_choice=selected_choice)


# Function to Send Emails
from icalendar import Calendar, Event

def send_emails(sender_email, recipients, names, subject, body, attachment, meeting_date=None, meeting_duration=None,cc_list=None, bcc_list=None):
    sg = SendGridAPIClient('sendgrid-api-key')

    # Log Activity ID
    cursor.execute("SELECT MAX(activity_id) FROM Activity")
    activity_id = cursor.fetchone()[0]
    cc_sent = False  # Flag to send CC/BCC only once

    for recipient, name in zip(recipients, names):
        status = 'Success'
        if not is_valid_email(recipient):
            status = 'Failed'
        else:
            # Create the email body
            full_body = body
            if meeting_date:
                # Create Accept and Decline links
                accept_url = f"http://127.0.0.1:8000/accept_meeting/{recipient}"
                decline_url = f"http://127.0.0.1:8000/decline_meeting/{recipient}"
                cursor.execute(
                    "INSERT INTO MeetingResponses (email) VALUES (%s) ON DUPLICATE KEY UPDATE response=NULL",
                    (recipient,)
                )
                db.commit()

                full_body = f"""
                <p>Hi {name}, </p><br>
                <p>{body}</p>
                <p><strong>Meeting Details:</strong></p>
                <p>Date and Time: {meeting_date}</p>
                <p>Duration: {meeting_duration} minutes</p>
                <p>Please respond:</p>
                <a href="{accept_url}" style="padding:10px 15px; background-color:green; color:white; text-decoration:none; border-radius:5px;">Accept</a>
                &nbsp;&nbsp;
                <a href="{decline_url}" style="padding:10px 15px; background-color:red; color:white; text-decoration:none; border-radius:5px;">Decline</a>
                """

                # full_body += "\n\nMeeting Request:\nDate and Time: {}\n".format(meeting_date)

            # Prepare the email
            message = Mail(
                from_email=sender_email,
                to_emails=recipient,
                subject=subject,
                html_content=full_body
            )
        
            # Add CC and BCC
            if not cc_sent:
                if cc_list:
                    message.cc = cc_list
                if bcc_list:
                    message.bcc = bcc_list
                cc_sent = True  # Only send once

            # Add attachment if provided
            if attachment:
                attachment_path = os.path.join(app.config['UPLOAD_FOLDER'], attachment.filename)
                attachment.save(attachment_path)
                with open(attachment_path, 'rb') as f:
                    data = f.read()
                encoded_file = base64.b64encode(data).decode()

                file_attachment = Attachment(
                    file_content=FileContent(encoded_file),
                    file_type=FileType(attachment.content_type),
                    file_name=FileName(attachment.filename),
                    disposition=Disposition("attachment")
                )
                message.add_attachment(file_attachment)

            # Add ICS file for a meeting request
            if meeting_date and meeting_duration:
                cal = Calendar()
                event = Event()
                event.add('summary', subject)
                event.add('dtstart', datetime.datetime.strptime(meeting_date, "%Y-%m-%dT%H:%M"))
                event.add('dtend', datetime.datetime.strptime(meeting_date, "%Y-%m-%dT%H:%M") + datetime.timedelta(minutes=int(meeting_duration)))
                event.add('description', body)
                event.add('organizer', sender_email)
                event.add('attendee', recipient)
                cal.add_component(event)

                # Convert calendar to string and encode as base64
                ics_data = cal.to_ical()
                encoded_ics = base64.b64encode(ics_data).decode()
                ics_attachment = Attachment()  # Use a separate variable to avoid overwriting
                ics_attachment.file_content = FileContent(encoded_ics)
                ics_attachment.file_type = FileType("text/calendar")
                ics_attachment.file_name = FileName("invite.ics")
                ics_attachment.disposition = Disposition("attachment")
                message.add_attachment(ics_attachment)

            # Send email
            try:
                sg.send(message)
            except Exception as e:
                print(f"Error sending email to {recipient}: {e}")
                status = 'Failed'

        # Log each email
        cursor.execute(
            """
            INSERT INTO Log (datetime, activity_id, email, name, status)
            VALUES (%s, %s, %s, %s, %s)
            """,
            (datetime.datetime.now(), activity_id, recipient, name, status)
        )
        db.commit()
@app.route('/accept_meeting/<email>')
def accept_meeting(email):
    cursor.execute("UPDATE MeetingResponses SET response='Accepted' WHERE email=%s", (email,))
    db.commit()
    return "<h2>Thank you for accepting the meeting request.</h2>"

@app.route('/decline_meeting/<email>')
def decline_meeting(email):
    cursor.execute("UPDATE MeetingResponses SET response='Declined' WHERE email=%s", (email,))
    db.commit()
    return "<h2>We have recorded your response as Declined.</h2>"


# Database Initialization
cursor.execute(
    """
    CREATE TABLE IF NOT EXISTS Activity (
        activity_id INT AUTO_INCREMENT PRIMARY KEY,
        date_of_activity DATE,
        time_of_activity TIME,
        activity_name VARCHAR(255),
        performed_by VARCHAR(255),
        number_of_emails INT
    )
    """
)

cursor.execute(
    """
    CREATE TABLE IF NOT EXISTS Log (
        id INT AUTO_INCREMENT PRIMARY KEY,
        datetime DATETIME,
        activity_id INT,
        email VARCHAR(255),
        name VARCHAR(255),
        status VARCHAR(50),
        FOREIGN KEY (activity_id) REFERENCES Activity(activity_id)
    )
    """
)
db.commit()

# Registration Table Creation
cursor.execute("""
CREATE TABLE IF NOT EXISTS Registration (
    id INT AUTO_INCREMENT PRIMARY KEY,
    name VARCHAR(255) NOT NULL,
    email VARCHAR(255) UNIQUE NOT NULL,
    phone VARCHAR(15),
    password VARCHAR(255) NOT NULL
)
""")
db.commit()

# Meeting request response Table Creation
cursor.execute("""
CREATE TABLE IF NOT EXISTS MeetingResponses (
    id INT AUTO_INCREMENT PRIMARY KEY,
    email VARCHAR(255) UNIQUE NOT NULL,
    response ENUM('Accepted', 'Declined') DEFAULT NULL,
    response_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP
)
""")
db.commit()

def send_otp_email(email, otp):
    sg = SendGridAPIClient('sendgrid_api_key')
    message = Mail(
        from_email='email_verified_on_sendgrid',
        to_emails=email,
        subject='Your OTP Code from EmailPro',
        plain_text_content=f'Your OTP code is {otp}.'
    )
    try:
        sg.send(message)
    except Exception as e:
        print(f"Error sending OTP: {e}")

@app.route('/signup', methods=['POST'])
def signup():
    name = request.form['name']
    email = request.form['email']
    phone = request.form['phone']
    password = request.form['password']

    # Generate OTP
    otp = random.randint(100000, 999999)

    if not is_valid_email(email):
        flash('Invalid email format.', 'error')
        return redirect(url_for('auth', form='signup'))

    try:
        # Check if the email already exists in the database
        cursor.execute("SELECT id FROM Registration WHERE email = %s", (email,))
        if cursor.fetchone():
            flash('Email already exists. Try logging in.', 'error')
            return redirect(url_for('auth', form='login'))

        # Store user data and OTP temporarily in the session
        session['signup_data'] = {
            'name': name,
            'email': email,
            'phone': phone,
            'password': password
        }
        session['otp'] = otp

        # Send OTP email
        send_otp_email(email, otp)
        flash('OTP sent to your email. Please verify.', 'success')
        return redirect(url_for('auth', form='otp'))
    except Exception as e:
        flash(f'Error during signup: {e}', 'error')
        return redirect(url_for('auth', form='signup'))

@app.route('/verify_otp', methods=['POST'])
def verify_otp():
    otp = request.form['otp']

    # Verify the OTP entered by the user
    if 'otp' in session and otp == str(session['otp']):
        # Retrieve user data from session
        signup_data = session.pop('signup_data', None)

        if signup_data:
            # Insert user into the database
            try:
                cursor.execute("""
                INSERT INTO Registration (name, email, phone, password)
                VALUES (%s, %s, %s, %s)
                """, (signup_data['name'], signup_data['email'], signup_data['phone'], signup_data['password']))
                db.commit()

                # Clear OTP from session
                session.pop('otp', None)

                flash('OTP verified successfully! You can now log in.', 'success')
                return redirect(url_for('auth', form='login'))
            except Exception as e:
                flash(f'Error saving user data: {e}', 'error')
                return redirect(url_for('auth', form='signup'))
        else:
            flash('Signup data not found. Please sign up again.', 'error')
            return redirect(url_for('auth', form='signup'))
    else:
        flash('Invalid OTP. Please try again.', 'error')
        return redirect(url_for('auth', form='otp'))

@app.route('/resend_otp', methods=['POST'])
def resend_otp():
    if 'signup_data' in session:
        otp = random.randint(100000, 999999)
        session['otp'] = otp

        email = session['signup_data']['email']
        send_otp_email(email, otp)

        flash('OTP resent to your email.', 'success')
        return redirect(url_for('auth', form='otp'))
    else:
        flash('Session expired. Please sign up again.', 'error')
        return redirect(url_for('auth', form='signup'))

@app.route('/login_form')
def login_form():
    return render_template('auth.html')

@app.route('/login', methods=['POST'])
def login():
    email = request.form['email']
    password = request.form['password']

    # Verify user credentials
    cursor.execute("""
    SELECT id, name FROM Registration WHERE email = %s AND password = %s
    """, (email, password))
    user = cursor.fetchone()

    if user:
        session['user_id'] = user[0]  # Store user ID in session
        session['user_name'] = user[1]  # Optionally store user name
        flash('Login successful!', 'success')
        return redirect(url_for('home'))
    else:
        flash('Invalid email or password.', 'error')
        return redirect(url_for('auth', form='login'))

@app.route('/auth', methods=['GET'])
def auth():
    form = request.args.get('form', 'login')
    return render_template('auth.html', form=form)

@app.route('/logout', methods=['GET'])
def logout():
    session.clear()  # Clear the session
    flash('You have been logged out.', 'success')
    return redirect(url_for('auth'))

if __name__ == '__main__':
    app.run(debug=True,port=8000)

