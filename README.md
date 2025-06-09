# Roundcube Webmail Email Downloader

A simple Python script to download emails and attachments from a Roundcube Webmail server using IMAP.

## Features

- Connects securely to Roundcube Webmail via IMAP
- Downloads and saves email messages locally
- Parses and decodes email headers and bodies
- **Resume support**: already downloaded emails or attachments will not be downloaded again
- No third-party dependencies – uses only Python standard libraries

## Libraries Used

This script uses only built-in Python libraries:

| Library        | Purpose                                      |
|----------------|----------------------------------------------|
| `imaplib`      | Connects to the email server using IMAP      |
| `email`        | Parses raw email content                     |
| `os`           | Manages directories and files                |
| `getpass`      | Securely accepts user password input         |
| `decode_header`| Decodes MIME-encoded email headers           |
| `re`           | Performs regular expression matching         |
| `datetime`     | Formats dates and timestamps                 |
| `json`         | Handles saving and loading metadata          |

## How to Run

1. Clone or download the script

   ```bash
   git clone https://your-repo-url
   cd your-repo-folder
  
2. Create a virtual environment

   ```bash
    python3 -m venv .venv

3. Activate the virtual environment

  On macOS/Linux:

      ```bash
      source venv/bin/activate

  On Windows:
      
      ```bash
      venv\Scripts\activate

4. Run the script

   ```bash
    python3 run.py

4. You will be prompted to enter your email credentials.

