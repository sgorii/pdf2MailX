![PowerShell](https://img.shields.io/badge/PowerShell-5.1+-blue?logo=powershell)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey?logo=windows)
![License](https://img.shields.io/github/license/sgorii/pdf2MailX)

## 📌 Description

`Pdf2MailX.ps1` is a PowerShell script that automates the batch emailing of Tewt-based PDF documents by detecting the recipient's email address directly from each file.

It offers the following features:
- Extracting email addresses from OCR-processed PDFs using `pdftotext` utility (can be found at https://www.xpdfreader.com) 
- Supports customizable corrections for common OCR errors in email addresses (via editable regex rules)
- Grouping documents by recipient
- Sending them via SMTP with optional attachments, HTML content, and read receipt requests
- Archiving all sent documents in structured subfolders (by date and recipient)
- Generating a session log and visual progress display


## ⚙️ Requirements

- Windows with PowerShell 5.1+
- [xpdf] (https://www.xpdfreader.com) - (`pdftotext.exe` required)
- OCR-processed PDFs (e.g. using [Umi OCR](https://github.com/hiroi-sora/Umi-OCR) for batch processing)
- SMTP credentials for sending emails


## 🛠️ Configuration

Before running the script, configure:

- `$subject` – email subject
- `$body` – HTML body of the email
- `$sourceFolder`: path to your PDF files
- `$pdftotextPath`: path to `pdftotext.exe`
- `$smtpServer`, `$smtpPort`, `$username`: SMTP settings
- `$attachmentsFixes`: any permanent attachments to all emails (optional)
- `$blockedEmails`: emails addresses to exclude from sending
- `$lotSize`: Number of maximum PDFs to send per email


## ▶️ Usage

-> run with powershell

You will be prompted to enter your SMTP password. The script will then:

    1. Extracts text from each PDF
    2. Detects and corrects email addresses
    3. Groups and sends emails in batches (default: 6 PDFs per email, can be modified)
    4. Logs the actions and moves sent PDFs to an archive structure
