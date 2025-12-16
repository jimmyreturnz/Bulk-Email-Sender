# Bulk-Email-Sender
An email sender I used for MU Mental Math Competition 2024-2025 to send certificates directly to participants' gmail account.

# Python Bulk Email Sender with / without Attachment

This Python script is designed to automate the process of sending personalized bulk emails with unique attachments (like certificates or score reports) to recipients listed in an Excel file. \
The process relies on matching a unique identifier (e.g., `student_id`) from the Excel sheet to the corresponding attachment file name. 

It was created due to the fact that I use pdf merger & splitter \
the app will split pdf into pages with file format named as [file_name]-Page[page] 

## What it does

This code allows you to:

* **Read Recipient Data:** Load recipient names, emails, and unique IDs from a specified Excel sheet using `pandas`.
* **Personalize Attachments:** Dynamically locate and attach individual files (e.g., PDF certificates) based on the file & recipient id (must match with file name)
* **Automate Sending:** Utilize the **SMTP protocol** via Gmail to send personalized emails in bulk.
* **Track Progress:** Print confirmation messages and timing for each email sent.

---

## How to Configure and Use the Script (Setup Guide)

To use this code, you must configure five primary sections.

### 1. SMTP Server (Crucial Step)

The script uses Google's SMTP server. You **must** generate an App Password for your Google account, as using your regular password will fail.

Modify the `email_config_base` dictionary inside the `main()` function:

| Parameter | Code Line to Edit | Description |
| :---: | :--- | :--- |
| `sender_email` | `'youremail@email.com'` | **Your Sender Email Address** (Must be a Gmail account for the current setup). |
| `password` | `'google-generated-password'` | **The App Password** generated from your Google Account Security settings. |
| `smtp_server` | `'smtp.gmail.com'` | Gmail's SMTP server (standard value).  It's ok to leave it like that|
| `smtp_port` | `587` | Port for TLS encryption (standard value). It's ok to leave it like that|

> ⚠️ **IMPORTANT:** Always use an **App Password** for automated login to external services like Gmail via code.

### 2. 📁 Change Excel File Path and Attachment Root Folder

Modify the following variables at the start of the `main()` function:

```python
def main():
    ######################################
    file_path = 'some_file.xlsx'      # <-- CHANGE this to your Excel file name
    folder_path = 'some_folder_of_your_choice' # <-- CHANGE this to the ROOT FOLDER containing your attachments
    ######################################
    # ...
```
### 3. Changing Excel Sheet to Read Data From
If your recipient data is on a sheet other than the first one (index 0), modify the sheet_name parameter in the load_data_from_excel() function:

```python
def load_data_from_excel(file_path):
    try:
        # Load data from the specified sheet index (index starts at 0)
        # For the first sheet: sheet_name=0
        # For the second sheet: sheet_name=1
        df_scores = pd.read_excel(file_path, sheet_name=0) <- change here
        return df_scores
    except Exception as e:
        # ...
```

### 4. Adjust Attachment File Path / Names
The script assumes your attachments are named based on the student_id column from your Excel file. If your file naming convention is different, modify the get_certificate_path() function:
```python
def get_certificate_path(student, folder_path):
    try:
        # Edit this line to match your file structure:
        # 'folder' is a sub-directory inside 'folder_path' (modify if needed)
        # The filename expects a column named 'student_id' in your Excel.
        filename = f"{folder_path}/folder/filename_{student['student_id']}.pdf" <- What I used, can be changed to the path you desire.
        
        if os.path.exists(filename):
            return filename
        # ...
```

### 5. Email Subject and Content
Modify the subject and message inside the loop within the main() function:
```python
# ... inside the loop ...
        
        # 🟢 Customize the Email Body/Message
        message = ("Any message") # Replace with your custom message here
        
        # Example of personalization (if you have a 'name' column in Excel):
        # message = (f"Dear {student['name']},\n\n" 
        #            "Here is your certificate.\n\n" 
        #            "Sincerely,\n[Your Name/Organization]")

        # 🟢 Customize the Email Subject
        send_email("Any topic", # Replace with your custom subject here
                   message, 
                   student['email'], # Assumes a column named 'email' in your Excel
                   attachment_path, 
                   email_config_base)
        # ...
```

> ⚠️ **IMPORTANT:** For experimentation, you should change **sheet_name** (in 3.) to where the test sheet is first.\
> If you would like to terminate the process, use terminate button (on colab) or hit ctrl+c on vscode
