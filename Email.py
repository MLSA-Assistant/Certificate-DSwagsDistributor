import os
import io
import csv
import win32com.client as win32
from PyPDF2 import PdfReader, PdfWriter
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

# Define the folder path to store the PDFs
output_folder = "certificates2.0"

# Create the folder if it doesn't exist
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# Email configuration
outlook = win32.Dispatch("outlook.application")

# Email body (you can customize this message)
email_body = "Hello,\n\nPlease find your certificate attached.\n\nBest regards,\nYour Organization"
email_subject = "Your Certificate"

# Read CSV file
csv_file = "participants.csv"

with open(csv_file, mode="r") as file:
    csv_reader = csv.DictReader(file)
    for row in csv_reader:
        recipient_email = row["Email"]
        name = row["Name"]
        topic = "python"  # Set your topic here

        # Load the PDF template
        pdf_template = PdfReader("certificate_template.pdf")

        # Create a PDF certificate with the recipient's name
        pdf_output = PdfWriter()

        # Clone the pages from the template and add to the output
        for page in pdf_template.pages:
            pdf_output.add_page(page)

        # Get the first page (you can modify this to fit your template structure)
        first_page = pdf_output.pages[0]

        # Create a canvas to add text fields to the page
        packet = io.BytesIO()
        can = canvas.Canvas(packet, pagesize=letter)

        # Define the x and y coordinates for the name and topic fields
        x_coord = 42
        y_coord = 390
        x_coord2 = 42
        y_coord2 = 250

        can.setFont("Helvetica-Bold", 20)
        can.drawString(x_coord, y_coord, name)
        can.setFont("Helvetica-Bold", 17)
        can.drawString(x_coord2, y_coord2, topic)

        can.save()
        packet.seek(0)
        new_pdf = PdfReader(packet)
        page = new_pdf.pages[0]
        first_page.merge_page(page)

        # Save the filled PDF with a unique name (using the recipient's name)
        pdf_filename = os.path.join(output_folder, f"certificate_{name}.pdf")
        pdf_filename = os.path.abspath(pdf_filename)  # Ensure it's an absolute path

        with open(pdf_filename, "wb") as output_file:
            pdf_output.write(output_file)

        # print(pdf_filename)
        # Create a new email
        mail = outlook.CreateItem(0)
        mail.To = recipient_email
        mail.Subject = email_subject
        mail.Body = email_body

        # Attach the certificate file (using the unique file path)
        mail.Attachments.Add(pdf_filename)

        # Send the emails
        mail.Send()
