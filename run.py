#!/usr/bin/env python3
"""
Roundcube Email Downloader
Downloads all emails as .eml files and extracts attachments to local storage
"""

import imaplib
import email
import os
import getpass
from email.header import decode_header
import re
from datetime import datetime
import json

class RoundcubeDownloader:
    def __init__(self, server, port=993, use_ssl=True):
        self.server = server
        self.port = port
        self.use_ssl = use_ssl
        self.imap = None
        self.base_dir = "downloaded_emails"
        self.progress_file = "download_progress.json"
        self.resume_mode = False
        self.processed_ids = set()
        
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
    
    def save_progress(self, current_email_id, total_emails, folder, processed_ids):
        """Save download progress to file"""
        progress_data = {
            'server': self.server,
            'folder': folder,
            'last_processed_id': current_email_id,
            'total_emails': total_emails,
            'processed_count': len(processed_ids),
            'processed_ids': list(processed_ids),
            'timestamp': datetime.now().isoformat(),
            'base_dir': self.base_dir
        }
        
        progress_path = os.path.join(self.base_dir, self.progress_file)
        try:
            with open(progress_path, 'w') as f:
                json.dump(progress_data, f, indent=2)
        except Exception as e:
            print(f"Warning: Could not save progress: {e}")
    
    def load_progress(self):
        """Load previous download progress"""
        progress_path = os.path.join(self.base_dir, self.progress_file)
        if not os.path.exists(progress_path):
            return None
        
        try:
            with open(progress_path, 'r') as f:
                return json.load(f)
        except Exception as e:
            print(f"Warning: Could not load progress file: {e}")
            return None
    
    def check_resume_option(self):
        """Check if resume is possible and ask user"""
        progress = self.load_progress()
        if not progress:
            return False, set()
        
        print(f"\nFound previous download session:")
        print(f"  Server: {progress.get('server', 'Unknown')}")
        print(f"  Folder: {progress.get('folder', 'Unknown')}")
        print(f"  Processed: {progress.get('processed_count', 0)}/{progress.get('total_emails', 0)} emails")
        print(f"  Last session: {progress.get('timestamp', 'Unknown')}")
        
        resume = input("\nWould you like to resume from where you left off? (y/N): ").strip().lower()
        if resume in ['y', 'yes']:
            return True, set(progress.get('processed_ids', []))
        else:
            # Clear old progress file
            try:
                os.remove(os.path.join(self.base_dir, self.progress_file))
            except:
                pass
            return False, set()
    
    def is_email_downloaded(self, email_id):
        """Check if email is already downloaded"""
        # Check multiple possible filename formats
        subject_patterns = [
            f"{int(email_id):06d}_*.eml",
            f"{email_id}_*.eml"
        ]
        
        emails_dir = os.path.join(self.base_dir, "emails")
        if not os.path.exists(emails_dir):
            return False
        
        import glob
        for pattern in subject_patterns:
            if glob.glob(os.path.join(emails_dir, pattern)):
                return True
        return False
    
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
            
            # Convert email_ids to integers for proper processing
            email_ids = [int(eid.decode() if isinstance(eid, bytes) else eid) for eid in email_ids]
            
            # Filter out already processed emails if resuming
            if self.resume_mode and self.processed_ids:
                remaining_ids = [eid for eid in email_ids if eid not in self.processed_ids]
                print(f"Resume mode: Skipping {len(self.processed_ids)} already processed emails")
                print(f"Remaining to process: {len(remaining_ids)} emails")
                email_ids = remaining_ids
            
            # Process each email
            successful_downloads = 0
            processed_count = len(self.processed_ids) if self.resume_mode else 0
            
            for i, email_id in enumerate(email_ids, 1):
                try:
                    # Skip if already processed (additional safety check)
                    if email_id in self.processed_ids:
                        continue
                    
                    # Fetch email
                    status, msg_data = self.imap.fetch(str(email_id), '(RFC822)')
                    if status != 'OK':
                        print(f"Failed to fetch email {email_id}")
                        continue
                    
                    # Process email
                    if self.process_email(email_id, msg_data[0][1]):
                        successful_downloads += 1
                        processed_count += 1
                        self.processed_ids.add(email_id)
                    
                    # Save progress every 10 emails
                    if i % 10 == 0 or i == len(email_ids):
                        self.save_progress(email_id, total_emails, folder, self.processed_ids)
                        print(f"Progress: {processed_count}/{total_emails} emails processed ({i}/{len(email_ids)} in current batch)")
                        
                except Exception as e:
                    print(f"Error processing email {email_id}: {e}")
                    continue
            
            # Final progress save
            if email_ids:  # Only save if we processed any emails
                self.save_progress(email_ids[-1] if email_ids else 0, total_emails, folder, self.processed_ids)
            
            print(f"\nDownload completed!")
            print(f"Successfully downloaded: {successful_downloads} new emails")
            print(f"Total processed: {processed_count}/{total_emails} emails")
            print(f"Files saved to: {os.path.abspath(self.base_dir)}")
            
            # Clean up progress file if download is complete
            if processed_count >= total_emails:
                try:
                    progress_path = os.path.join(self.base_dir, self.progress_file)
                    if os.path.exists(progress_path):
                        os.remove(progress_path)
                        print("Download complete - progress file cleaned up")
                except:
                    pass
            
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
    
    def disconnect(self):
        """Close IMAP connection"""
        if self.imap:
            try:
                self.imap.close()
                self.imap.logout()
                print("Disconnected from server")
            except:
                pass

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
        
        # Check for existing progress before asking for folder
        downloader.create_directories()  # Create directories first
        resume_mode, processed_ids = downloader.check_resume_option()
        
        # Set resume state
        downloader.resume_mode = resume_mode
        downloader.processed_ids = processed_ids
        
        if resume_mode:
            # Load progress to get the folder
            progress = downloader.load_progress()
            if progress:
                folder = progress.get('folder', 'INBOX')
                print(f"Resuming download from folder: {folder}")
            else:
                print("Could not load progress details, please select folder manually.")
                resume_mode = False
                downloader.resume_mode = False
        
        if not resume_mode:
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
        
        # Start download with resume capability
        print(f"\nStarting download from '{folder}'...")
        if not resume_mode:
            confirm = input("Continue? (y/N): ").strip().lower()
            if confirm not in ['y', 'yes']:
                print("Download cancelled")
                return
        
        # Start the download
        downloader.download_all_emails(folder)
        
    except KeyboardInterrupt:
        print("\nDownload interrupted by user")
        print("Progress has been saved. You can resume later.")
    except Exception as e:
        print(f"Unexpected error: {e}")
    finally:
        downloader.disconnect()

if __name__ == "__main__":
    main()
