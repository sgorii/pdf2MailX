Script to automatically extract email addresses from OCR-processed PDF files and send them in grouped emails with attachments.

## 📌 Description

`Pdf2MailX.ps1` is a PowerShell script designed to automate the process of sending one-page .pdf documents to the mail address found in file by:

- Extracting email addresses from OCR-processed PDFs using `pdftotext` utility (can be found at https://www.xpdfreader.com) 
- Correcting common OCR errors in emails
- Grouping documents by recipient
- Sending them via SMTP with optional attachments, html text and Request for read receipt
- Archiving all sent documents in structured subfolders (by date and recipient)
- Generating a session log and visual progress display


## ⚙️ Requirements

- Windows with PowerShell 5.1+
- [xpdf] (https://www.xpdfreader.com) - (`pdftotext.exe` required)
- OCR-processed PDFs (e.g., Umi OCR for batch processing)
- SMTP credentials for sending emails


## 🛠️ Configuration

Before running the script, configure:

- `$subject`:
- `$body`:
- `$sourceFolder`: path to your PDF files
- `$pdftotextPath`: path to `pdftotext.exe`
- `$smtpServer`, `$smtpPort`, `$username`: SMTP settings
- `$attachmentsFixes`: any permanent attachments to all emails (optional)
- `$blockedEmails`: emails addresses to exclude from sending
- `$lotSize`: Number of maximum PDFs to send per email


## ▶️ Usage

-> run with powershell
You will be prompted to enter your SMTP password. The script then:

    1. Extracts text from each PDF
    2. Detects and corrects email addresses
    3. Groups and sends emails in batches (up to 6 PDFs per mail)
    4. Logs the actions and moves sent PDFs to an archive structure
