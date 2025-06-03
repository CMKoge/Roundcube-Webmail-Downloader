#!/usr/bin/env python3
"""
Roundcube Email Downloader
Downloads all emails as .eml files and extracts attachments to local storage
"""

import imaplib
import email
import os
import sys
import getpass
from email.header import decode_header
import re
from datetime import datetime

class RoundcubeDownloader:
    def __init__(self, server, port=993, use_ssl=True):
        self.server = server
        self.port = port
        self.use_ssl = use_ssl
        self.imap = None
        self.base_dir = "downloaded_emails"
        
    def connect(self, username, password):
        """Connect to IMAP server"""
        try:
            if self.use_ssl:
                self.imap = imaplib.IMAP4_SSL(self.server, self.port)
            else:
                self.imap = imaplib.IMAP4(self.server, self.port)
            
            self.imap.login(username, password)
            print(f"Successfully connected to {self.server}")
            return True
        except Exception as e:
            print(f"Connection failed: {e}")
            return False
    
    def decode_mime_words(self, s):
        """Decode MIME encoded-words"""
        if s is None:
            return ""
        decoded_fragments = decode_header(s)
        decoded_string = ''
        for fragment, encoding in decoded_fragments:
            if isinstance(fragment, bytes):
                if encoding:
                    decoded_string += fragment.decode(encoding)
                else:
                    decoded_string += fragment.decode('utf-8', errors='ignore')
            else:
                decoded_string += fragment
        return decoded_string
    
    def sanitize_filename(self, filename):
        """Sanitize filename for safe file system storage"""
        if not filename:
            return "unnamed_file"
        
        # Remove or replace invalid characters
        filename = re.sub(r'[<>:"/\\|?*]', '_', filename)
        filename = filename.strip('. ')
        
        # Limit length
        if len(filename) > 200:
            name, ext = os.path.splitext(filename)
            filename = name[:200-len(ext)] + ext
            
        return filename
    
    def create_directories(self):
        """Create necessary directories"""
        os.makedirs(self.base_dir, exist_ok=True)
        os.makedirs(os.path.join(self.base_dir, "emails"), exist_ok=True)
        os.makedirs(os.path.join(self.base_dir, "attachments"), exist_ok=True)
    
    def save_attachment(self, part, email_id, att_count):
        """Save email attachment"""
        filename = part.get_filename()
        if filename:
            filename = self.decode_mime_words(filename)
            filename = self.sanitize_filename(filename)
        else:
            # Generate filename based on content type
            content_type = part.get_content_type()
            ext = '.bin'
            if 'image' in content_type:
                ext = '.jpg'
            elif 'text' in content_type:
                ext = '.txt'
            elif 'pdf' in content_type:
                ext = '.pdf'
            filename = f"attachment_{email_id}_{att_count}{ext}"
        
        attachment_dir = os.path.join(self.base_dir, "attachments", f"email_{email_id}")
        os.makedirs(attachment_dir, exist_ok=True)
        
        filepath = os.path.join(attachment_dir, filename)
        
        # Handle duplicate names
        counter = 1
        original_filepath = filepath
        while os.path.exists(filepath):
            name, ext = os.path.splitext(original_filepath)
            filepath = f"{name}_{counter}{ext}"
            counter += 1
        
        try:
            with open(filepath, 'wb') as f:
                f.write(part.get_payload(decode=True))
            print(f"  Saved attachment: {filename}")
            return True
        except Exception as e:
            print(f"  Failed to save attachment {filename}: {e}")
            return False
    
    def process_email(self, email_id, msg_data):
        """Process individual email - optimized for Microsoft import"""
        try:
            # Parse email
            email_message = email.message_from_bytes(msg_data)
            
            # Get email metadata
            subject = self.decode_mime_words(email_message.get('Subject', 'No Subject'))
            sender = self.decode_mime_words(email_message.get('From', 'Unknown Sender'))
            date = email_message.get('Date', 'Unknown Date')
            
            print(f"Processing Email {email_id}: {subject[:50]}...")
            
            # Save .eml file (Microsoft Outlook compatible)
            safe_subject = self.sanitize_filename(subject)
            eml_filename = f"{email_id:06d}_{safe_subject[:50]}.eml"
            eml_path = os.path.join(self.base_dir, "emails", eml_filename)
            
            # Save with proper .eml format for Microsoft compatibility
            with open(eml_path, 'wb') as f:
                f.write(msg_data)
            
            # Also save attachments separately for easier access
            attachment_count = 0
            if email_message.is_multipart():
                for part in email_message.walk():
                    content_disposition = str(part.get("Content-Disposition"))
                    
                    # Check if it's an attachment
                    if "attachment" in content_disposition:
                        attachment_count += 1
                        self.save_attachment(part, email_id, attachment_count)
            
            return True
            
        except Exception as e:
            print(f"Failed to process email {email_id}: {e}")
            return False
    
    def download_all_emails(self, folder='INBOX'):
        """Download all emails from specified folder"""
        try:
            # Select folder with proper error handling
            status, count = self.imap.select(folder)
            if status != 'OK':
                print(f"Failed to select folder '{folder}'. Status: {status}")
                # Try common folder name variations
                folder_variations = [
                    folder,
                    f'"{folder}"',  # Quoted folder name
                    folder.upper(),
                    folder.lower(),
                    f"INBOX.{folder}" if folder != 'INBOX' else folder,
                    f"[Gmail]/{folder}" if 'gmail' in self.server.lower() else folder
                ]
                
                for variation in folder_variations:
                    try:
                        status, count = self.imap.select(variation)
                        if status == 'OK':
                            print(f"Successfully selected folder using variation: '{variation}'")
                            folder = variation
                            break
                    except:
                        continue
                else:
                    print(f"Could not select folder '{folder}' with any variation")
                    return False
            
            print(f"Selected folder: {folder}")
            
            # Search for all emails
            status, messages = self.imap.search(None, 'ALL')
            if status != 'OK':
                print(f"Failed to search emails in folder '{folder}'. Status: {status}")
                return False
            
            email_ids = messages[0].split()
            total_emails = len(email_ids)
            
            print(f"Found {total_emails} emails in {folder}")
            
            if total_emails == 0:
                print("No emails to download")
                return True
            
            self.create_directories()
            self.create_import_instructions()
            
            # Process each email
            successful_downloads = 0
            for i, email_id in enumerate(email_ids, 1):
                try:
                    # Fetch email
                    status, msg_data = self.imap.fetch(email_id, '(RFC822)')
                    if status != 'OK':
                        print(f"Failed to fetch email {email_id}")
                        continue
                    
                    # Process email
                    if self.process_email(int(email_id), msg_data[0][1]):
                        successful_downloads += 1
                    
                    # Progress indicator
                    if i % 10 == 0 or i == total_emails:
                        print(f"Progress: {i}/{total_emails} emails processed")
                        
                except Exception as e:
                    print(f"Error processing email {email_id}: {e}")
                    continue
            
            print(f"\nDownload completed!")
            print(f"Successfully downloaded: {successful_downloads}/{total_emails} emails")
            print(f"Files saved to: {os.path.abspath(self.base_dir)}")
            
            return True
            
        except Exception as e:
            print(f"Download failed: {e}")
            return False
    
    def list_folders(self):
        """List available folders with better formatting"""
        try:
            status, folders = self.imap.list()
            if status == 'OK':
                print("Available folders:")
                folder_names = []
                for folder in folders:
                    folder_str = folder.decode()
                    # Extract folder name from IMAP LIST response
                    parts = folder_str.split('"')
                    if len(parts) >= 3:
                        folder_name = parts[-2]
                    else:
                        # Fallback parsing
                        folder_name = folder_str.split()[-1]
                    
                    folder_names.append(folder_name)
                    print(f"  - {folder_name}")
                
                return folder_names
            else:
                print(f"Failed to list folders. Status: {status}")
                return []
        except Exception as e:
            print(f"Failed to list folders: {e}")
            return []
    
    def disconnect(self):
        """Close IMAP connection"""
        if self.imap:
            try:
                self.imap.close()
                self.imap.logout()
                print("Disconnected from server")
            except:
                pass

def create_import_instructions(self):
    """Create instructions file for importing to Microsoft email clients"""
    instructions = """
# Email Import Instructions for Microsoft Email Clients

## Microsoft Outlook (Desktop)

### Method 1: Drag and Drop
1. Open Microsoft Outlook
2. Navigate to the folder where you want to import emails
3. Open Windows Explorer and navigate to the 'emails' folder
4. Select the .eml files you want to import
5. Drag and drop them into the Outlook folder

### Method 2: File Import
1. Open Outlook
2. Go to File > Open & Export > Import/Export
3. Choose "Import from another program or file"
4. Select "Outlook Data File (.pst)"
5. Browse to select your .eml files

## Outlook.com (Web)
1. Log into Outlook.com
2. Go to the folder where you want to import
3. Drag .eml files directly into the email list
4. Or use the "Upload" option if available

## Windows Mail App
1. Open Windows Mail app
2. Navigate to desired folder
3. Drag .eml files into the email list

## Thunderbird (Alternative)
1. Install ImportExportTools NG add-on
2. Right-click folder > ImportExportTools NG > Import messages > Import EML files
3. Select the emails folder

## Notes:
- All .eml files preserve original formatting, headers, and metadata
- Attachments are embedded in .eml files when possible
- Large attachments are saved separately in the 'attachments' folder
- Import may take time for large numbers of emails
- Some email clients may require you to import in batches

## File Locations:
- Email files (.eml): emails/
- Attachments: attachments/
- This instructions file: IMPORT_INSTRUCTIONS.txt
"""
    
    instructions_path = os.path.join(self.base_dir, "IMPORT_INSTRUCTIONS.txt")
    with open(instructions_path, 'w', encoding='utf-8') as f:
        f.write(instructions.strip())
    print(f"Import instructions saved to: {instructions_path}")

def main():
    print("Roundcube Email Downloader - Microsoft Compatible")
    print("=" * 50)
    
    # Get server details
    server = input("IMAP Server (e.g., mail.yourdomain.com): ").strip()
    if not server:
        print("Server address is required")
        return
    
    # Get port (default 993 for SSL)
    port_input = input("IMAP Port (default 993 for SSL, 143 for non-SSL): ").strip()
    port = 993 if not port_input else int(port_input)
    use_ssl = port == 993 or port == 465
    
    # Get credentials
    username = input("Email address: ").strip()
    if not username:
        print("Email address is required")
        return
    
    password = getpass.getpass("Password: ")
    if not password:
        print("Password is required")
        return
    
    # Create downloader instance
    downloader = RoundcubeDownloader(server, port, use_ssl)
    
    try:
        # Connect to server
        if not downloader.connect(username, password):
            return
        
        # List available folders
        print("\nChecking available folders...")
        available_folders = downloader.list_folders()
        
        # Choose folder
        folder = input("\nEnter folder to download (default: INBOX): ").strip()
        if not folder:
            folder = 'INBOX'
        
        # Validate folder exists
        if available_folders and folder not in available_folders:
            print(f"\nWarning: '{folder}' not found in available folders.")
            print("Available folders are:")
            for f in available_folders:
                print(f"  - {f}")
            
            confirm_folder = input(f"Still try to download from '{folder}'? (y/N): ").strip().lower()
            if confirm_folder not in ['y', 'yes']:
                return
        
        # Confirm download
        print(f"\nReady to download all emails from '{folder}'")
        confirm = input("Continue? (y/N): ").strip().lower()
        if confirm not in ['y', 'yes']:
            print("Download cancelled")
            return
        
        # Start download
        print(f"\nStarting download from '{folder}'...")
        downloader.download_all_emails(folder)
        
    except KeyboardInterrupt:
        print("\nDownload interrupted by user")
    except Exception as e:
        print(f"Unexpected error: {e}")
    finally:
        downloader.disconnect()

if __name__ == "__main__":
    main()
