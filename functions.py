import os
import glob
import win32com.client as win32
import json

def add_attachments(mail, user, extension):
    """
    Searches for files with the specified user name and file extension
    in the 'Files' directory and adds them as attachments to the email.

    Parameters:
    - mail: The Outlook email object.
    - user: The user name to search for in the file names.
    - extension: The file extension to filter files (e.g., '.pdf').

    Example:
    If user='johndoe' and extension='.txt', this function will attach
    all files matching 'johndoe*.txt' in the 'Files' directory.
    """
    # Get the absolute path of the current script
    current_directory = os.path.dirname(os.path.abspath(__file__))

    # Define the folder containing the files and the search pattern
    folder_path = os.path.join(current_directory, 'Files')
    search_pattern = os.path.join(folder_path, f"{user}*{extension}")

    # Find all matching files in the directory
    attachments = glob.glob(search_pattern)

    # Attach each found file to the email
    for attachment in attachments:
        mail.Attachments.Add(attachment)

def get_project_root_directory():
    """
    Traverses upwards from the current script directory to find the
    root directory containing 'email_sender'.

    Returns:
    - The absolute path to the root directory containing 'email_sender'.
    """
    current_directory = os.path.dirname(os.path.abspath(__file__))
    while 'email_sender' not in os.path.basename(current_directory):
        current_directory = os.path.dirname(current_directory)
    return current_directory

def load_json_data(file_path):
    """
    Loads data from a JSON file.

    Parameters:
    - file_path: The absolute path to the JSON file.

    Returns:
    - A Python object (usually a list or dictionary) representing the JSON data.

    Example:
    Use this to load user data from a file like 'emails_test.json'.
    """
    with open(file_path, 'r') as file:
        return json.load(file)

def initialize_outlook():
    """
    Initializes and returns an instance of the Outlook application.

    Returns:
    - An Outlook application object that can be used to create and send emails.
    """
    return win32.Dispatch('outlook.application')

def set_mail_body_and_user(mail, entry, body_brazil, body_us, body_latam):
    """
    Sets the email body content and formats the user name based on the country.

    Parameters:
    - mail: The Outlook email object.
    - entry: A dictionary containing 'country' and 'user' keys.
    - body_brazil: The email body template for Brazil.
    - body_us: The email body template for the USA.
    - body_latam: The email body template for other Latin American countries.

    Returns:
    - user: The formatted user name based on the country.

    Behavior:
    - For Brazil ('BRASIL'), the user's name is capitalized, and the Brazilian template is set.
    - For the USA ('USA'), the user's name is converted to uppercase, and the US template is set.
    - For other countries, the user's name is converted to uppercase, and the Latin American template is set.
    """
    country = entry['country'].upper()
    user = entry['user']
    
    if country == 'BRASIL':
        mail.Body = body_brazil
        user = user.capitalize()
    elif country == 'USA':
        mail.Body = body_us
        user = user.upper()
    else:
        mail.Body = body_latam
        user = user.upper()
    
    return user
