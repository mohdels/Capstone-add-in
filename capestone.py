from flask import Flask, request, jsonify
from flask_cors import CORS
import smtplib
import os
from email.mime.text import MIMEText
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__)
CORS(app)

EMAIL_USER = os.getenv("EMAIL_USER")
EMAIL_PASS = os.getenv("EMAIL_PASS")
TO_EMAIL = os.getenv("TO_EMAIL")

@app.route('/send-email', methods=['POST'])
def send_email():
    try:
        data = request.json
        name = data.get('name')
        email = data.get('email')
        phone = data.get('phone')
        category = data.get('category')
        submission_context = data.get('submissionContext')
        description = data.get('description')
        date = data.get('date')

        body = f"""New Form Submission:

        Name: {name}
        Email: {email}
        Phone: {phone}
        Date: {date}
        Category: {category}
        In-person/phone: {submission_context}
        Details: {description}
        """

        msg = MIMEText(body)
        msg['Subject'] = "New Form Submission"
        msg['From'] = EMAIL_USER
        msg['To'] = TO_EMAIL

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(EMAIL_USER, EMAIL_PASS)
            smtp.sendmail(EMAIL_USER, TO_EMAIL, msg.as_string())

        return jsonify({"message": "Email sent successfully!"}), 200

    except Exception as e:
        print("Error:", e)
        return jsonify({"message": "Failed to send email.", "error": str(e)}), 500
    
@app.route('/email', methods=['POST'])
def categorize_email():
    data = request.json  # Get the JSON data sent by the client
    body = data.get('body', '').lower()

    categories = []
    print('backend has been called')
    if 'urgent' in body:
        categories.append('Urgent')
    elif 'meeting' in body:
        categories.append('Meeting')
    elif 'invoice' in body:
        categories.append('Invoice')
    else:
        categories.append('other')

    return jsonify(categories)  # Return the categories as a JSON response

@app.route('/category-change', methods=['POST'])
def category_change():
    # Get the JSON data sent by the client
    data = request.json

    body = data.get('body', '').lower()

    print('Received request data for /category-change:')

    # No return value (no response)
    return '', 204  # 204 No Content response

if __name__ == '__main__':
    app.run(port=3002)
