# Auto-Certificate-Generator
This Python script automates the generation and emailing of personalized certificates of appreciation. It reads names and emails from an Excel file, creates customized certificates with each recipient’s name, saves the certificates as images, and then sends them via email.

1. Import Required Librarie
    i) smtplib: Handles sending emails using an SMTP server.
    ii) email.mime: Helps in formatting email messages with attachments.
    iii) PIL (Pillow): Used to open, edit, and save images.
    iv) pandas: Reads and processes Excel files.
    v) openpyxl: Helps read and write Excel files (required for .xlsx format).
    vi) os: Manages file paths and directories.

2. Define File Paths:
     i) excel_path: Specifies the location of the Excel file that contains names and emails.
     ii) output_folder: Defines the folder where the generated certificate images will be stored.

3. Read Data from Excel:
    Reads the Excel file and extracts Name and Email columns.

4. Create Output Directory:
    Ensures the output directory exists to store certificate images.

5. Configure Email Settings:
   i) Defines sender credentials (not secure to hardcode passwords, consider using environment variables).
   ii) Sets up the email subject and body, where [Name] is a placeholder for personalization.

6. Define SMTP Server Settings:
   i) Opens a predefined certificate template.
   ii) Loads a custom font and writes the recipient’s name at the specified position.
   iii) Saves the personalized certificate as an image in the output folder.

7. Loop Through Names & Generate Certificates:
      i) Opens a predefined certificate template.
      ii) Loads a custom font and writes the recipient’s name at the specified position.
      iii) Saves the personalized certificate as an image in the output folder.

8. Prepare and Send Email with Certificate Attachment:
   i) Creates an email message.
   ii) Personalizes the body by replacing [Name] with the recipient’s name.
   iii) Attaches the generated certificate image.

9. Send Email via SMTP:
      i) Establishes a secure connection with the SMTP server.
      ii) Logs in and sends the email with the certificate as an attachment.
      iii) Closes the connection.
