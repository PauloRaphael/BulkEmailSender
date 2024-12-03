import win32com.client as win32
import json
import os
import functions as funcs


#Email Body on for Brazil
body_brazil = '''
                Mensagem
                '''
                
body_latam = '''
              Mensagem 
              '''
                
body_us = '''
           Mensagem 
           '''
try:
        # Get the root directory
    current_directory = funcs.get_project_root_directory()

    # Load user data from JSON
    json_file_path = os.path.join(current_directory, 'Emails\\emails_example.json')
    json_users = funcs.load_json_data(json_file_path)

    # Initialize Outlook
    outlook = funcs.initialize_outlook()

    # Process each user entry
    for entry in json_users:
        try:
            mail = outlook.CreateItem(0)  # Create email
            mail.Subject = 'Subject'

            # Set email body and user formatting
            user = funcs.set_mail_body_and_user(mail, entry, body_brazil, body_us, body_latam)

            # Add attachments
            funcs.add_attachments(mail, user, "txt")

            # Set recipient and send email
            mail.To = entry['email']
            mail.Send()
            print(f"Email sent to {entry['email']}!")
        except Exception as e:
            print(f"Failed to send email to {entry.get('email', 'unknown')} - {e}")

except FileNotFoundError as e:
    print(f"File not found: {e}")
except json.JSONDecodeError:
    print("Error: JSON file could not be decoded. Please check its format.")
except Exception as e:
    print(f"An error occurred: {e}")
