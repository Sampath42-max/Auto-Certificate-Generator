import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from PIL import Image, ImageDraw, ImageFont
import pandas as pd
import openpyxl as pl
import os

# Load the file paths
excel_path = r"F:\python\example.xlsx" # Path to your Excel file
output_folder = r"F:\python\output" # Folder to save images

# Read the names and emails from the Excel file
df = pd.read_excel(excel_path) # Use read_excel for Excel files
names = df['Name'] # Replace 'Name' with the actual column name containing names in your Excel file
emails = df['Email'] # Replace 'Email' with the actual column name containing emails in your Excel file

# Create output folder if it doesn't exist
os.makedirs(output_folder, exist_ok=True)

# Email configuration
sender_email = "kumar485685@gmail.com"
sender_password = "syheoipzxoniizqb"
subject = "Your Certificate of Appreciation"
body = "Dear [Name],\n\nPlease find your Certificate of Appreciation attached.\n\nBest regards,\nYour Organization"

# SMTP server configuration
smtp_server = "smtp.gmail.com"
smtp_port = 587

# Iterate over each name and generate the image
for name, email in zip(names, emails):
    # Open the base image
    img = Image.open(r"F:\python\Red and Yellow Minimalist Certificate of Appreciation.png") # Replace with your image path
    d = ImageDraw.Draw(img)

    # Load the font
    font = ImageFont.truetype(r"F:\python\Milk Candy.otf", 70)

    # Define text position, text, and color
    text = name
    text_pos = (850, 655) # Adjust as needed
    text_color = (0, 0, 0)

    # Add the text to the image
    d.text(text_pos, text=text, fill=text_color, font=font)

    # Save the modified image
    output_path = os.path.join(output_folder, f"{text}_img.png")
    img.save(output_path)

    # Send the email with the certificate attached
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = email
    msg['Subject'] = subject

    # Attach the body with the message instance
    msg.attach(MIMEText(body.replace("[Name]", name), 'plain'))

    # Open the file in bynary
    attachment = open(output_path, "rb")

    # Encode to base64
    part = MIMEBase('application', 'octet-stream')
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', "attachment; filename= " + f"{text}_img.png")

    # Attach the instance 'p' to instance 'msg'
    msg.attach(part)

    # Create server object with SSL option
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    server.login(sender_email, sender_password)
    text = msg.as_string()
    server.sendmail(sender_email, email, text)
    server.quit()

print(f"Images have been saved in the folder: {output_folder} and emails have been sent.")
