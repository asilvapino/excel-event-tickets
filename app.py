from reportlab.lib.pagesizes import landscape
from reportlab.pdfgen import canvas
import pandas as pd
import os
import qrcode
from PIL import Image
#IMPORTS FOR MAIL
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import logging

# Setup logging
logging.basicConfig(filename='email_errors.log', level=logging.ERROR)

# Function to send email with PDF ticket attached
def send_email(email, pdf_filename):
    # SMTP Mail server config
    smtp_server = 'mail.example.com'  # Your email server provider
    smtp_port = 587
    smtp_username = 'example@email.com'  # Your email address
    smtp_password = 'EXAMPLE PASSWORD'  # Your email password

    # Email content
    sender_email = smtp_username
    receiver_email = email
    subject = 'Ticket for your event'
    # write the HTML Template
    html = """\
    <html>
    <body>
    <p><span lang="es" dir="ltr">Hey!</span></p>
    <p><span lang="es" dir="ltr">We are sending the tickets for the event.</span></p>
    </br>
    <p><span lang="es" dir="ltr">Atte: Example Stuff</span></p>
    </body>
    </html>
    """

    # Create a multipart message
    message = MIMEMultipart()
    message['From'] = sender_email
    message['To'] = receiver_email
    message['Subject'] = subject

    # Attach HTML body
    message.attach(MIMEText(html, 'html'))
  
    # Adjust PDF filename
    #base_filename = os.path.basename(pdf_filename)

    # Attach PDF ticket
    with open(pdf_filename, 'rb') as attachment:
        part = MIMEApplication(attachment.read(), _subtype='pdf', Name=os.path.basename(pdf_filename))
        part.add_header('Content-Disposition', 'attachment', filename=os.path.basename(pdf_filename))

    # Add attachment to message
    message.attach(part)

    # Convert message to string
    text = message.as_string()

    # Connect to SMTP server and send email
    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(smtp_username, smtp_password)
            server.sendmail(sender_email, receiver_email, text)
        print(f"Email sent to {receiver_email}")
    except Exception as e:
        logging.error(f"Failed to send email to {receiver_email}: {str(e)}")

# Function to generate PDF ticket
def generate_ticket(name, number, prefix_url, background_image, output_directory, fontsize=18):
    img = Image.open(background_image)
    width, height = img.size

    output_filename = os.path.join(output_directory, f"{name}.pdf")
    c = canvas.Canvas(output_filename, pagesize=(width, height))

    # Draw background image
    c.drawImage(background_image, 0, 0, width=width, height=height)

    # Generate QR code
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=8,
        border=2,
    )
    data = f"{prefix_url}{number}"
    qr.add_data(data)
    qr.make(fit=True)
    qr_img = qr.make_image(fill_color="black", back_color="white")

    # Calculate position for QR code (left-aligned and vertically centered)
    qr_img_width, qr_img_height = qr_img.size
    qr_x = 90  # Left margin
    qr_y = ((height - qr_img_height) / 2) - 13

    # Draw QR code on the ticket
    c.drawInlineImage(qr_img, qr_x, qr_y)

    # Set font properties
    c.setFont("Helvetica", fontsize)

    # Write name on ticket
    name_x = qr_x + qr_img_width + 40  # Adjust horizontal position
    name_y = height / 2  # Vertically center
    c.drawString(name_x, name_y, name)

    # Write number below name
    number_x = name_x
    number_y = name_y - fontsize - 5  # Below the name with a slight gap
    c.drawString(number_x, number_y, f"Ticket: #{str(number).zfill(3)}")  # Format number with leading zeros

    c.save()

# Read names, numbers, and emails from Excel sheet
def read_data_from_excel(excel_file):
    df = pd.read_excel(excel_file)
    data = [(row['Name'], row['Number'], row['Email']) for _, row in df.iterrows()]
    return data

if __name__ == "__main__":
    excel_file = "names.xlsx"  # Path to your Excel file
    background_image = "background_image.jpg"  # Path to your background image
    output_directory = "output_tickets"  # Output directory for saving tickets
    prefix_url = "https://api.example.domain/validate?ticketID="  # Prefix URL for the QR code

    # Create output directory if it doesn't exist
    if not os.path.exists(output_directory):
        os.makedirs(output_directory)

    data = read_data_from_excel(excel_file)

    # Generate and send ticket for each name (you can configure the font size here)
    for name, number, email in data:
        generate_ticket(name, number, prefix_url, background_image, output_directory, fontsize=40)
        pdf_filename = os.path.join(output_directory, f"{name}.pdf")
        send_email(email, pdf_filename)

    print("PDF tickets generated and sent successfully.")


