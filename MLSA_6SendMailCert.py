import os
import csv
import win32com.client as win32

# Define the folder path where the certificates are located
certificates_folder = "Certificates"

# Email configuration
outlook = win32.Dispatch("outlook.application")

# Signature HTML code
signature_html = """
<html>
<head>
    <style>
        body {
            font-family: 'Segoe UI', Arial, sans-serif;
        }

        .signature-container {
            display: flex;
            align-items: center;
        }

        .right-content {
            flex-grow: 1;
        }

        .name {
            font-size: 9pt;
            font-weight: bold;
        }

        .role {
            font-size: 9pt;
        }

        .email {
            font-size: 9pt;
            color: #0072C6; /* Blue color for clickable link */
        }

        .contact-info {
            font-size: 14px;
        }
    </style>
</head>
<body>
    <div class="signature-container">
        <div class="right-content">
            <a href="https://www.linkedin.com/in/missmeerali/" target="_blank">Meerali Naseet</a>
            <div class="role">Microsoft Learn Student Ambassador</div>
                
            </div>
        </div>
    </div>
</body>
</html>
"""

# Updated link for the digital swag kit
swag_kit_link = "https://stdntpartners-my.sharepoint.com/:u:/g/personal/meerali_naseet_studentambassadors_com/EevIBlw-REBEmSADs-AlQOwBgyd_GwRnLLnNrxfbW_ZrZQ?e=hzF7ZC"

# Email body (you can customize this message)
email_body = f"Hello!<br><br>Kindly find your Completion Certificate for (Event Name) in the attachment.<br> <a href='{swag_kit_link}'> Digital Swag Kit </a> <br> <br>Keep on upskilling! <br><br><br>Regards,<br>{signature_html}"

# Email subject
email_subject = "Completion Certificate for (Event Name)"

# Read CSV file
csv_file = "completed_modules.csv"

with open(csv_file, mode="r") as file:
    csv_reader = csv.DictReader(file)
    for row in csv_reader:
        recipient_email = row["Email"]
        name = row["Name"]

        # Certificate file path
        certificate_filename = os.path.join(certificates_folder, f"{name}_Certificate.pdf")
        certificate_filename = os.path.abspath(certificate_filename)

        # Check if the certificate file exists
        if os.path.exists(certificate_filename):
            # Create a new email
            mail = outlook.CreateItem(0)
            mail.To = recipient_email
            mail.Subject = email_subject

            # Set email body
            mail.HTMLBody = email_body

            # Attach the existing certificate file
            mail.Attachments.Add(certificate_filename)

            # Send the email
            mail.Send()
        else:
            print(f"Certificate not found for {name}.")
