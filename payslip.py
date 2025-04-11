



import pandas as pd
from email import encoders
from fpdf import FPDF
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
import smtplib

# Define the employee data
data = {
    "Employee ID": ["E001", "E002", "E003", "E004"],
    "Name": ["Nicole Zonke", "Precious Tendayi", "Simbarashe", "Millicent Gumbira"],
    "Email": ["nicolezonke393@gmail.com", "precioustendayi36@gmail.com", "simbarashe@uncommon.org", "millicentlisy@gmail.com"],
    "Basic Salary": [1000, 1600, 2000, 1700],
    "Allowances": [300, 200, 100, 50],
    "Taxi Deduction": [50, 60, 70, 30],
    "NSSA Deduction": [30, 40, 50, 60]
}

# Create a DataFrame
df = pd.DataFrame(data)

# Define the email sender credentials
sender_email = "starlisy7@gmail.com"
sender_password = "eoji olom dxjn duuc"  # Use an App Password for Gmail

# Loop through each employee and send an email
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

    # Generate a PDF for the payslip
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    pdf.cell(200, 10, txt="Payslip", ln=True, align='C')
    pdf.ln(10)
    pdf.cell(200, 10, txt=f"Employee ID: {employee_id}", ln=True)
    pdf.cell(200, 10, txt=f"Name: {name}", ln=True)
    pdf.cell(200, 10, txt=f"Email: {email}", ln=True)
    pdf.cell(200, 10, txt=f"Basic Salary: {basic_salary}", ln=True)
    pdf.cell(200, 10, txt=f"Allowances: {allowances}", ln=True)
    pdf.cell(200, 10, txt=f"Taxi Deduction: {taxi_deduction}", ln=True)
    pdf.cell(200, 10, txt=f"NSSA Deduction: {nssa_deduction}", ln=True)
    pdf.cell(200, 10, txt=f"Net Salary: {net_salary}", ln=True)

    # Save the PDF
    pdf_filename = f"Payslip_{employee_id}.pdf"
    pdf.output(pdf_filename)

    # Create the email message
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = email
    msg['Subject'] = f"Payslip for {name}"

    # Add email body
    body = f"""
Dear {name},
Please find attached your payslip for this month.
Best regards,
RARE DIAMOND INVESTMENT
"""
    msg.attach(MIMEText(body, 'plain'))

    # Attach the PDF
    with open(pdf_filename, "rb") as attachment:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header(
            'Content-Disposition',
            f'attachment; filename={pdf_filename}',
        )
        msg.attach(part)

    # Send the email
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, email, msg.as_string())
        server.quit()
        print(f"Email sent to {email} with payslip attached.")
    except Exception as e:
        print(f"Failed to send email to {email}. Error: {e}")

    # Print confirmation
    print(f"Payslip sent to {name} ({email})")
    print("------------------------")












