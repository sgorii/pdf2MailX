# === GENERAL CONFIGURATION ===

# IMPORTANT:
# This script is designed to process PDF files, extract email addresses, and send them via SMTP.
# It uses the pdftotext utility (https://www.xpdfreader.com) to extract text from PDF files.

# This script assumes that the PDF files are already OCR-processed before execution.
# If the PDF contains only images (no selectable text), pdftotext will return nothing.
# In such cases, use an OCR tool (e.g., Umi OCR for batch processing) before running this script.


# BlackList : addresses to ignore
$blockedEmails = @(
   "example@domain.tld"     # Example of an address to block
)

# Folder Paths
# Change these paths according to your environment
$sourceFolder = ""                          # Path where the PDF files are located
$archiveRoot = Join-Path $sourceFolder "archive"
$pdftotextPath = ""                         # Path to the pdftotext.exe (e.g., C:\Program Files\Xpdf\pdftotext.exe)
$logFile = Join-Path $sourceFolder "log_p2M.txt"

# Check if pdftotext is available
if (!(Test-Path $pdftotextPath)) {
    Write-Host "[X] pdftotext not found!" -ForegroundColor Red
    Read-Host -Prompt "Check the path and press Enter to exit"
    exit
}

# === EMAIL CONFIGURATION ===   
# Subject and body of the email
$subject = "" # Subject of the email
$body = @"    

"@

# === ATTACHMENTS ===
$attachmentsFixes = @(
                            # Add paths to the attachments you want to include in the email
                            # Example: "C:\path\to\attachment1.pdf", "C:\path\to\attachment2.pdf"
                            # Ensure these paths are valid and accessible
)

$cheminsInvalides = $attachmentsFixes | Where-Object { -not (Test-Path $_) }        # Check if attachment paths are valid
if ($cheminsInvalides.Count -gt 0) {
    Write-Host "[X] Invalid attachment paths found :" -ForegroundColor Red
    $cheminsInvalides | ForEach-Object { Write-Host " - $_" -ForegroundColor Yellow }
    Read-Host -Prompt "Please fix the paths and press Enter to exit"
    exit
}

# === CONFIGURATION SMTP ===
$smtpServer = ""                                    # SMTP server
$smtpPort = 587                                     # SMTP port (usually 587 for TLS)
# Note: Ensure that the SMTP server and port are correct for your environment.

$username = ""                                      # Email address used for sending
$password = Read-Host "Enter your password" -AsSecureString                             # Enter your password
$cred = New-Object System.Management.Automation.PSCredential ($username, $password)     # Create the authentication credentials

# === ERROR HANDLING ===
$ErrorActionPreference = "SilentlyContinue"         # Ignore errors
$global:ProgressPreference = "SilentlyContinue"     # Ignore progress bars

# === INIT LOG + DOSSIERS ===
if (!(Test-Path $logFile)) {                       # Create log file if it doesn't exist
    New-Item -ItemType File -Path $logFile -Force | Out-Null
}
if (!(Test-Path $archiveRoot)) {                   # Create archive folder if it doesn't exist
    New-Item -ItemType Directory -Path $archiveRoot | Out-Null
}
"`n--- New session on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ---" | Out-File -Append -FilePath $logFile

# === FUNCTIONS ===
# Function to print messages safely, ensuring the console width is respected
function Safe-Print {
    param(
        [string]$text,
        [ConsoleColor]$color = $null
    )
    $width = $host.UI.RawUI.WindowSize.Width
    [System.Console]::SetCursorPosition(0, [System.Console]::CursorTop)
    if ($color) {
        Write-Host $text.PadRight($width) -ForegroundColor $color
    } else {
        Write-Host $text.PadRight($width) 
    }         
}
# Function to show a fixed progress bar
# This function displays a progress bar that does not change size, ensuring it fits within the console
function Show-FixedProgress {
    param (
        [int]$Current,
        [int]$Total,
        [datetime]$StartTime
    )
    $percent = [math]::Round(($Current / $Total) * 100)
    $bar = ('■' * ($percent / 5)).PadRight(20, '□')
    if ($Current -ge 5) {
        $elapsed = (Get-Date) - $StartTime
        $avgTime = $elapsed.TotalSeconds / $Current
        $eta = [timespan]::FromSeconds($avgTime * ($Total - $Current))
        $etaStr = $eta.ToString("hh\:mm\:ss")
    } else {
        $etaStr = "Calculating..."
    }
    [System.Console]::SetCursorPosition(0, 0)
    Write-Host "[$Current/$Total] $percent% [$bar] ETA: $etaStr".PadRight(80)
    [System.Console]::SetCursorPosition(0, 2)
}

# Function to correct common email errors
# This function corrects common mistakes in email addresses extracted from PDFs
function CorrectMail {
    param ([string]$email)
    if (-not $email) { return $null }

    # Example corrections
    # Corrects common mistakes found on mail addresses made during OCR conversion or manual entry
    $email = $email -replace '\.gou\.fr', '.gouv.fr'                            # .gou.fr → .gouv.fr

    return $email
}

# Function to extract email from PDF using pdftotext
# This function uses pdftotext to extract text from a PDF file and then searches for
function Get-EmailFromPdf {
    param($pdfPath)
    $tempTxt = "$pdfPath.txt"
    try {
        & "$pdftotextPath" -raw "`"$pdfPath`"" "`"$tempTxt`"" 2>$null
    } catch {
        Safe-Print "[!] Failed extraction : $pdfPath" -color white
        return $null
    }

    if (Test-Path $tempTxt) {
        $text = Get-Content -Path $tempTxt -Raw
        Remove-Item "$tempTxt" -Force

        # === FIX BROKEN EMAILS CAUSED BY OCR ===
        # Some OCR engines may split email addresses across multiple lines.
        # We attempt to reconstruct valid addresses based on common patterns.

        # Case 1: break with hyphen and newline
        # e.g. "firstname.lastname-\nservice@domain.com" → "firstname.lastname@domain.com"
        $text = $text -replace "([a-z0-9._%-]+)-\s*[\r\n]+\s*([a-z0-9._%-]+@[a-z0-9.-]+\.[a-z]{2,6})", '$1$2'

        # Case 2: break at the "@" symbol
        # e.g. "firstname.lastname\n@domain.com" → "firstname.lastname@domain.com"
        $text = $text -replace "([a-z0-9._%-]+)[\r\n]+\s*@([a-z0-9.-]+\.[a-z]{2,6})", '$1@$2'

        # Look for the first valid email address in the cleaned text
        if ($text -match '\b[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}\b') {
            return $matches[0]
        }
    }
    return $null
}

# === MAIN SCRIPT ===
# This is the main script that processes PDF files, extracts email addresses, and sends them via SMTP.
# It handles renaming files with special characters, extracting emails, and sending grouped emails.
$emailGroups = @{}

$allPdfs = Get-ChildItem -Path $sourceFolder -Filter *.pdf |        # Extract all PDF files from the source folder
    Where-Object { $_.DirectoryName -notlike "*archive*" }          # Exclude PDFs in the archive folder

# Renaming files with special characters to avoid issues
foreach ($pdf in $allPdfs) {
    if ($pdf.Name -match '[^\w\.-]') {
        $cleanName = ($pdf.BaseName -replace '[^\w\.-]', '_') + ".pdf"
        $cleanPath = Join-Path $pdf.DirectoryName $cleanName
        if (-not (Test-Path $cleanPath)) {
            Rename-Item -Path $pdf.FullName -NewName $cleanName
        }
    }
}

# Re-fetch the list of PDFs after renaming
$allPdfs = Get-ChildItem -Path $sourceFolder -Filter *.pdf |
    Where-Object { $_.DirectoryName -notlike "*archive*" }
    

# Check if there are any PDFs to process
$total = $allPdfs.Count
if ($total -eq 0) {
    Write-Host "`n"
    Safe-Print "No PDF found in folder `"$sourceFolder`"." -color Yellow
    Write-Host ""
    Read-Host -Prompt "Press Enter to exit"
    exit
}
$startTime = Get-Date
$counter = 0
Write-Host "`n"                         # Reserve the first line
foreach ($pdfFile in $allPdfs) {        # Iterate through each PDF file
    $pdf = $pdfFile.FullName
    $email = Get-EmailFromPdf -pdfPath $pdf

    $emailBefore = $email               # Store the original email for logging
    $email = CorrectMail $email
    if ($email -ne $emailBefore) {
        "[$(Get-Date -Format 'HH:mm:ss')] CORRECTED - $($pdfFile.Name) => $email" | Out-File -Append -FilePath $logFile
    }
    if ($email -and ($blockedEmails -contains $email)) {
        Safe-Print "[!] Blocked address ignored: $email" -color DarkYellow
        "[$(Get-Date -Format 'HH:mm:ss')] BLOCKED - $email - $($pdfFile.Name)" | Out-File -Append -FilePath $logFile
        continue
    }
    if ($email) {
        if (-not $emailGroups.ContainsKey($email)) {
            $emailGroups[$email] = @()
        }

    } else {
        [System.Console]::SetCursorPosition(0, [System.Console]::CursorTop)
        Safe-Print "[X] No email address found in : $($pdfFile.Name)"
        "[$(Get-Date -Format 'HH:mm:ss')] NO EMAIL - $($pdfFile.Name)" | Out-File -Append -FilePath $logFile
    }
    $counter++
    Show-FixedProgress -Current $counter -Total $total -StartTime $startTime        # Show progress
}

# === GROUP EMAIL SENDING IN BATCHES ===
foreach ($email in $emailGroups.Keys) {
    $pdfs = $emailGroups[$email]
    $chunks = @()
    $lotSize = 6         # Number of maximum PDFs to send per email ; you can adjust this value as needed
    for ($i = 0; $i -lt $pdfs.Count; $i += $lotSize) {
        $end = [Math]::Min($i + $lotSize - 1, $pdfs.Count - 1)
        $chunk = $pdfs[$i..$end]
        $chunks += ,$chunk
    }
    foreach ($chunk in $chunks) {
        $allAttachments = $attachmentsFixes + $chunk
        [System.Console]::SetCursorPosition(0, [System.Console]::CursorTop)
        Safe-Print "[-] Sending to $email : $($chunk.Count) attachment(s)..." -color white
        try {
            $mail = New-Object System.Net.Mail.MailMessage
            $mail.From = New-Object System.Net.Mail.MailAddress($username) # Email address used for sending
            $mail.To.Add($email)
            $mail.Bcc.Add($username)    # BCC to the sender's email address
            $mail.Subject = $subject 
            $mail.Body = $body
            $mail.BodyEncoding = [System.Text.Encoding]::UTF8
            $mail.IsBodyHtml = $true

            # Request for read receipt
            $mail.Headers.Add("Disposition-Notification-To", $username)
            foreach ($file in $allAttachments) {
                $mail.Attachments.Add($file)
            }

            $smtp = New-Object Net.Mail.SmtpClient($smtpServer, $smtpPort) # SMTP server configuration
            $smtp.EnableSsl = $true    # Use SSL
            $smtp.Credentials = $cred  # Authentication with credentials
            $smtp.Send($mail)          # Send the email
            $mail.Dispose()
            Safe-Print "[✓] Mail sent to $email" -color Green
            "[$(Get-Date -Format 'HH:mm:ss')] OK - $email => $($chunk.Count) files" | Out-File -Append -FilePath $logFile
            $safeEmail = ($email -replace '[^\w\.-]', '_')
            $archiveFolder = Join-Path $archiveRoot $safeEmail
            if (!(Test-Path $archiveFolder)) {
                New-Item -ItemType Directory -Path $archiveFolder | Out-Null
            }
            foreach ($pdf in $chunk) {      # Move PDFs to the archive folder
                $fileDate = (Get-Item $pdf).LastWriteTime
                $year = $fileDate.Year
                $month = '{0:D2}' -f $fileDate.Month
                $day = '{0:D2}' -f $fileDate.Day

                $subArchive = Join-Path $archiveFolder "$year\$month\$day"      # Sub-folder structure based on date
                if (!(Test-Path $subArchive)) {                                 # Create sub-folder if it doesn't exist
                    New-Item -ItemType Directory -Path $subArchive -Force | Out-Null
                }

                $destinationPath = Join-Path $subArchive $cleanName             # Destination path
                Move-Item -Path $pdf -Destination $destinationPath -Force       # Move PDF to archive
            }
        } catch {
            Safe-Print ("[{0}] ERROR - {1} : {2}" -f (Get-Date -Format 'HH:mm:ss'), $email, $_.Exception.Message)
            "[$(Get-Date -Format 'HH:mm:ss')] ERROR - $email : $($_.Exception.Message)" | Out-File -Append -FilePath $logFile
        }
        Start-Sleep -Milliseconds 200
    }
}

# === SUMMARY ===
# This section summarizes the results of the email sending process.
$fullLog = Get-Content $logFile
$lastsessionStart = ($fullLog | Select-String "--- New session on")[-1].Linenumber
$sessionLines = $fullLog | Select-Object -Skip $lastsessionStart
$logLines = $sessionLines | Where-Object { $_ -match "^\[\d{2}:\d{2}:\d{2}\]" }
$totalOK = ($logLines | Where-Object { $_ -like "*OK*" }).Count                 
$totalNoEmail = ($logLines | Where-Object { $_ -like "*NO EMAIL*" }).Count
$totalErrors = ($logLines | Where-Object { $_ -like "*ERROR*" }).Count
[System.Console]::SetCursorPosition(0, [System.Console]::CursorTop)

# Print the summary of the results
Safe-Print "--> Summary :" -color white
if ($totalOK -gt 0) {
    Safe-Print "[✓] $totalOK mail(s) sent" -color green
} else { 
    Safe-Print "[✓] $totalOK mail(s) sent" -color white
}
if ($totalNoEmail -gt 0) {
    Safe-Print "[X] $totalNoEmail without email address" -color red
} else { 
    Safe-Print "[X] $totalNoEmail without email address" -color white
} 
if ($totalErrors -gt 0) {
    Safe-Print "[!] $totalErrors SMTP error(s)" -color red
} else { 
    Safe-Print "[!] $totalErrors SMTP error(s)" -color white
}
$totalBlocked = ($logLines | Where-Object { $_ -like "*BLOCKED*" }).Count
Safe-Print "[-] $totalBlocked blocked address(es) ignored" -color DarkYellow
Write-Host "`n"
Read-Host -Prompt "Press Enter to exit"