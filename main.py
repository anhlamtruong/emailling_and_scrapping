import time
from sending_email import config
from sending_email import excel_handler
from sending_email import template_handler
from sending_email import email_sender

def main():
    print(f"--- Starting Email Sender ---")
    
    # --- 1. GET PASSWORD SECURELY ---
    try:
        password = config.PASSWORD
    except Exception as e:
        print(f"Error getting password: {e}")
        return

    # --- 2. LOAD EXCEL DATA ---
    df = excel_handler.load_recipients(config.EXCEL_FILE)
    if df is None:
        print("Exiting program.")
        return
        
    # --- 3. CONNECT TO EMAIL SERVER ---
    server = email_sender.login_to_server(
        config.SMTP_SERVER, 
        config.SMTP_PORT, 
        config.SENDER_EMAIL, 
        password
    )
    
    if server is None:
        print("Could not log in. Exiting program.")
        return

    # --- 4. ITERATE, PERSONALIZE, AND SEND ---
    emails_sent_count = 0
    try:
        for index, row in df.iterrows():
            # Check if the email needs to be sent
            if str(row.get('Sent or Not', '')).lower() == 'sent' or str(row.get('Sent or Not', '')).lower() == 'response':
                print(f"  [SKIPPED] Email for {row.get('Name')} already marked as 'Sent or Response'.")
                continue
                
            # Get the template file name from the row
            template_file = row.get('Template File')
            if not template_file:
                print(f"  [SKIPPED] No template file specified for {row.get('Name')}.")
                continue

            # A. Load the correct template
            subject_template, body_template = template_handler.load_template(
                config.TEMPLATE_FOLDER, 
                template_file
            )
            if subject_template is None:
                print(f"  [SKIPPED] Could not load template for {row.get('Name')}.")
                continue
            
            # B. Personalize the template
            subject, body,image_to_embed = template_handler.personalize_template(
                subject_template, 
                body_template, 
                row, 
                config.YOUR_NAME,
                config.YOUR_PHONE_NUMBER, 
                config.SENDER_EMAIL,       
                config.YOUR_STATE_AND_CITY,
                config.IMAGE_FOLDER,
                template_file
            )
            
            recipient_email = row.get('Email')
            # Attach resume to ALL emails
            files_to_send = config.ATTACHMENT_FILES
            
            # C. Send the email
            success = email_sender.send_email(
                server, 
                config.SENDER_EMAIL, 
                recipient_email, 
                subject, 
                body,
                attachment_paths=files_to_send,
                inline_image_path=image_to_embed,
                
            )
            
            if success:
                print(f"  [SUCCESS] Email sent to {row.get('Name')} at {recipient_email}")
                df.at[index, 'Sent or Not'] = 'Sent'
                emails_sent_count += 1
                
                # Be polite to the server, wait a moment
                time.sleep(2)

    except Exception as e:
        print(f"An unexpected error occurred during sending: {e}")
    finally:
        # --- 5. QUIT SERVER ---
        server.quit()
        print(f"\nLogged out of SMTP server.")
        
        # --- 6. SAVE CHANGES TO EXCEL ---
        if emails_sent_count > 0:
            excel_handler.save_recipients(df, config.EXCEL_FILE)
        else:
            print("No new emails were sent, no changes saved to Excel.")

if __name__ == '__main__':
    main()