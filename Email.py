import io
import os
import csv
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from PyPDF2 import PdfReader, PdfWriter
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics

# Define the folder path to store the PDFs
output_folder = 'certificates2.0'

# Create the folder if it doesn't exist
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# Email configuration
smtp_server = 'smtp.gmail.com'
smtp_port = 587
smtp_username = 'username'
smtp_password = 'token'
sender_email = 'email'
email_subject = 'Certificate Attached'

# Read CSV file
csv_file = 'recipient_list.csv'
topic = "python"
with open(csv_file, mode='r') as file:
    csv_reader = csv.DictReader(file)
    for row in csv_reader:
        name = row['Name']
        recipient_email = row['Email']

        # Load the PDF template
        pdf_template = PdfReader('certificate_template.pdf')

        # Create a PDF certificate with the recipient's name
        pdf_output = PdfWriter()

        # Clone the pages from the template and add to the output
        for page_num in range(len(pdf_template.pages)):
            page = pdf_template.pages[page_num]
            pdf_output.add_page(page)


        # Get the first page (you can modify this to fit your template structure)
        first_page = pdf_output.pages[0]

        # Create a canvas to add text fields to the page
        packet = io.BytesIO()
        can = canvas.Canvas(packet, pagesize=letter)
        # Create a canvas with custom font and size
        # Define the x and y coordinates for the name and topic fields
        x_coord = 42
        y_coord = 390
        x_coord2 = 42
        y_coord2 = 250
        can.setFont('Helvetica-Bold', 20)
        can.drawString(x_coord, y_coord, name)  
        can.setFont('Helvetica-Bold', 17)
        can.drawString(x_coord2, y_coord2, topic)  # Replace x_coord2 and y_coord2 with your desired positions
        can.save()
        packet.seek(0)
        new_pdf = PdfReader(packet)
        page = new_pdf.pages[0]
        #first_page.mergePage(page)
        first_page.merge_page(page)

        # Save the filled PDF with a unique name
        pdf_filename = os.path.join(output_folder, f'certificate_{name}.pdf')
        with open(pdf_filename, 'wb') as output_file:
            pdf_output.write(output_file)

        # Create email message
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg['Subject'] = email_subject

        # Email body
        email_body = f"Hello {name},\n\nPlease find your certificate attached.\n\nBest regards,\nYour Organization"
        msg.attach(MIMEText(email_body, 'plain'))

        # Attach the filled certificate PDF
        with open(pdf_filename, 'rb') as pdf_file:
            attachment = MIMEApplication(pdf_file.read(), _subtype="pdf")
        attachment.add_header('Content-Disposition', f'attachment; filename={pdf_filename}')
        msg.attach(attachment)

        # Connect to SMTP server and send email
        try:
            server = smtplib.SMTP(smtp_server, smtp_port)
            server.starttls()
            server.login(smtp_username, smtp_password)
            server.sendmail(sender_email, recipient_email, msg.as_string())
            server.quit()
            print(f"Email sent to {recipient_email}")
        except Exception as e:
            print(f"Failed to send email to {recipient_email}: {str(e)}")
