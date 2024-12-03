# Email Sender Automation  

This project is a Python-based script designed to automate the process of sending emails through Microsoft Outlook. It allows for customized email bodies, personalized attachments, and region-specific content for users, all configured via a JSON file.  

## Features  
- **Automated email generation**: Reads user data from a JSON file and sends personalized emails.  
- **Region-specific email content**: Different email bodies for users from Brazil, the USA, and other Latin American countries.  
- **Attachment handling**: Automatically adds attachments based on user-specific file names and extensions.  
- **Outlook integration**: Utilizes Microsoft Outlook installed on the computer to send emails.

---

## How It Works  

The script reads user data from a JSON file (`emails.json`), formats the email content based on the user's region, attaches files matching the user's name, and sends the email using the computer's Outlook application.  

### Required JSON Structure  
The `emails.json` file should be created in the root directory of the project. The file must follow this structure:  

```json
[
  {
    "email": "user@example.com",
    "user": "username",
    "country": "BRASIL"
  },
  {
    "email": "anotheruser@example.com",
    "user": "anotherusername",
    "country": "USA"
  }
]
```

### JSON Field Descriptions  
- **`email`**: The recipient's email address.  
- **`user`**: The recipient's username or identifier.  
- **`country`**: The recipient's country. Accepted values are `"BRASIL"`, `"USA"`, or any other country.  

---

## Functions  

### `add_attachments(mail, user, extension)`  
- **Purpose**: Adds files matching the user's name and the specified extension as attachments to the email.  
- **Parameters**:  
  - `mail`: The Outlook email object.  
  - `user`: The user's name to match in the file names.  
  - `extension`: The file extension (e.g., `".pdf"`, `".txt"`).  

### `get_project_root_directory()`  
- **Purpose**: Locates the root directory containing the project folder `'email_sender'`.  
- **Returns**: The absolute path to the root directory.  

### `load_json_data(file_path)`  
- **Purpose**: Loads and parses JSON data from a specified file.  
- **Parameters**:  
  - `file_path`: The path to the JSON file.  
- **Returns**: The parsed JSON object.  

### `initialize_outlook()`  
- **Purpose**: Initializes and returns the Outlook application instance.  
- **Returns**: A `win32com.client.Dispatch` object for interacting with Outlook.  

### `set_mail_body_and_user(mail, entry, body_brazil, body_us, body_latam)`  
- **Purpose**: Sets the email body content and formats the user name based on the recipient's country.  
- **Parameters**:  
  - `mail`: The Outlook email object.  
  - `entry`: A dictionary containing `country` and `user`.  
  - `body_brazil`: Email template for Brazil.  
  - `body_us`: Email template for the USA.  
  - `body_latam`: Email template for Latin America.  
- **Returns**: The formatted username.  

---

## Prerequisites  

1. **Python**: Ensure Python 3.x is installed.  
2. **Outlook**: A configured Microsoft Outlook application must be installed on your computer.  
3. **Dependencies**: Install the required Python libraries by running:  
   ```bash
   pip install pywin32
   ```  

---

## Setup  

1. **Clone the Repository**  
   ```bash
   git clone https://github.com/PauloRaphael/bulk_email_sender.git
   cd email-sender
   ```  

2. **Create the JSON File**  
   - Add a file named `emails.json` in the project directory.  
   - Populate it using the provided JSON structure.  

3. **Ensure Directory Structure**  
   - Place user-specific files in a subdirectory named `Files` within the project directory.  

4. **Run the Script**  
   ```bash
   python main.py
   ```  

---

## Notes  

- Emails are sent using the **Outlook application** on your computer. Ensure Outlook is installed and configured correctly.  
- Make sure the JSON file follows the correct structure to avoid errors.  
- Add the `emails.json` file and the `Files` directory to your `.gitignore` file to prevent accidental exposure of sensitive data.  

---

## Example Usage  

1. Populate the `emails.json` file:  
   ```json
   [
     {
       "email": "john.doe@example.com",
       "user": "johndoe",
       "country": "BRASIL"
     },
     {
       "email": "jane.smith@example.com",
       "user": "janesmith",
       "country": "USA"
     }
   ]
   ```  

2. Place user-specific files in the `Files` directory:  
   - `Files/johndoe_invoice.pdf`  
   - `Files/janesmith_receipt.txt`  

3. Run the script and watch Outlook handle the rest!  

---

## License  

This project is licensed under the MIT License.  

---

## Contributing  

Contributions are welcome! Feel free to submit a pull request or open an issue for any suggestions or bugs.  

Enjoy your automated emailing experience! 🚀  