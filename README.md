# Task Automation with Python

## Project Description

This project aims to automate repetitive tasks such as sending emails, renaming files, and updating spreadsheets. The script uses various Python libraries like `smtplib` for sending emails, `os` for file management, and `openpyxl` for Excel operations.

## Skills Demonstrated

- Scripting
- Automation
- Working with file systems
- Interacting with APIs

## Features

1. **Send Emails**: Automate the process of sending emails with customizable content.
2. **Rename Files**: Rename multiple files in a directory based on specific patterns.
3. **Update Spreadsheets**: Modify Excel spreadsheets by updating specific cells.

## Installation

### Prerequisites

- Python 3.x
- `openpyxl` library

### Setup

1. Clone the repository:
    ```bash
    git clone https://github.com/yourusername/TaskAutomation.git
    cd TaskAutomation
    ```

2. (Optional) Create and activate a virtual environment:
    ```bash
    python -m venv venv
    source venv/bin/activate  # On Windows use `venv\Scripts\activate`
    ```

3. Install required libraries:
    ```bash
    pip install openpyxl
    ```

## Usage

Run the script:
bash
```python automation_tool.py```

## Task Automation Menu
- Send Email
>> Enter SMTP server, port, sender email, password, recipient email, subject, and body to send an email.

- Rename Files
>> Enter the directory path, old pattern, and new pattern to rename files in a specified directory.

- Update Spreadsheet
>> Enter the file path, sheet name, cell, and new value to update a cell in an Excel spreadsheet.



