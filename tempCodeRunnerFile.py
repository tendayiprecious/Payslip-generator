

import pandas as pd
from email import encoders
from fpdf import FPDF
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
import smtplib
import os

# === 1. Employee data ===
data = {
    "Employee ID": ["E001", "E002", "E003", "E004"],
    "Name": ["Nicole Zonke", "Precious Tendayi", "Simbarashe", "Millicent Gumbira"],
    "Email": ["nicolezonke393@gmail.com", "precioustendayi36@gmail.com", "simbarashe@uncommon.org", "millicentlisy@gmail.com"],
    "Basic Salary": [1000, 1600, 2000, 1700],
    "Allowances": [300, 200, 100, 50],
    "Taxi Deduction": [50, 60, 70, 30],
    "NSSA Deduction": [30, 40, 50, 60]
}

df = pd.DataFrame(data)

# === 2. Gmail sender credentials ===
sender_email = "starlisy7@gmail.com"
sender_password = "eoji olom dxjn duuc"  # Use Gmail App Password

# === 3. Create output folder ===
os.makedirs("payslips", exist_ok=True)

# === 4. Loop through each employee ===
for index, row in df.iterrows():
    employee_id = row['Employee ID']
    name = row['Name']
    email = row['Email']
    basic_salary = row['Basic Salary']
    allowances = row['Allowances']
    taxi_deduction = row['Taxi Deduction']
    nssa_deduction = row['NSSA Deduction']

    # Calculate net salary
    net_salary = basic_salary + allowances - taxi_deduction - nssa_deduction

    # === Generate the PDF payslip ===
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    pdf.cell(200, 10, txt="Payslip", ln=True, align='C')
    pdf.ln(10)
    pdf.cell(200, 10, txt=f"Employee ID: {employee_id}", ln=True)
    pdf.cell(200, 10, txt=f"Name: {name}", ln=True)
    pdf.cell(200, 10, txt=f"Basic Salary: ${basic_salary:.2f}", ln=True)
    pdf.cell(200, 10, txt=f"Allowances: ${allowances:.2f}", ln=True)
    pdf.cell(200, 10, txt=f"Taxi Deduction: ${taxi_deduction:.2f}", ln=True)
    pdf.cell(200, 10, txt=f"NSSA Deduction: ${nssa_deduction:.2f}", ln=True)
    pdf.cell(200, 10, txt=f"Net Salary: ${net_salary:.2f}", ln=True)
    pdf.ln(10)
    pdf.cell(200, 10, txt="Thank you for your service!", ln=True)

    # Save the PDF
    filename = f"payslips/{employee_id}.pdf"
    pdf.output(filename)

    # === Create the email ===
    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = email
    message["Subject"] = "Your Payslip for This Month"

    body = f"Dear {name},\n\nPlease find your payslip attached.\n\nBest regards,\nHR Department"
    message.attach(MIMEText(body, "plain"))

    # Attach the PDF
    with open(filename, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename= {os.path.basename(filename)}")
        message.attach(part)

    # === Send the email ===
    try:
        with smtplib.SMTP("smtp.gmail.com", 587) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, email, message.as_string())
        print(f"✅ Email sent to {name} ({email})")
    except Exception as e:
        print(f"❌ Failed to send email to {name}: {e}")
