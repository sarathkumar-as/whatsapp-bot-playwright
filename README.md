# WhatsApp Automation Using Python and Playwright

A Python automation project that demonstrates how Playwright can be used to open WhatsApp Web, read contact information from an Excel file, search for contacts, send messages, and generate an execution report.

> This project is intended for learning and authorized personal use only. Do not use it to send unsolicited or bulk messages.

## Project Objective

The objective of this project is to automate a basic WhatsApp Web workflow using Python and Playwright.

The program demonstrates:

* Opening WhatsApp Web
* Using a persistent browser profile
* Reading contact information from an Excel file
* Searching for a WhatsApp contact
* Entering and sending a message
* Handling basic automation errors
* Saving execution results locally

## Technologies Used

* Python
* Playwright
* Microsoft Excel
* Visual Studio Code
* Git and GitHub

## Project Structure

```text
whatsapp-bot-playwright/
├── playwright_assign.py
├── requirements.txt
├── README.md
├── PROJECT_NOTES.md
├── .gitignore
├── sample/
│   └── contacts_sample.xlsx
└── docs/
    ├── workflow.png
    └── output_sample.png
```

## Installation

### 1. Clone the repository

```bash
git clone https://github.com/SarathKumar-as/whatsapp-bot-playwright.git
```

### 2. Open the project folder

```bash
cd whatsapp-bot-playwright
```

### 3. Create a virtual environment

```bash
python -m venv venv
```

### 4. Activate the virtual environment

Windows PowerShell:

```powershell
.\venv\Scripts\Activate.ps1
```

Windows Command Prompt:

```cmd
venv\Scripts\activate
```

### 5. Install the required packages

```bash
pip install -r requirements.txt
```

### 6. Install the Playwright browser

```bash
playwright install chromium
```

## Input File

The program uses an Excel file containing contact and message information.

Example structure:

| Contact Name   | Message                 |
| -------------- | ----------------------- |
| Sample Contact | This is a test message. |

Use only fake information in the public sample file. Do not upload personal contact details to GitHub.

## Run the Program

```bash
python playwright_assign.py
```

On the first run, WhatsApp Web may ask you to scan a QR code. The local WhatsApp profile must not be uploaded to GitHub.

## Important Privacy Information

The following files and folders are excluded from GitHub:

* Actual contacts Excel file
* WhatsApp browser profile
* Login and session information
* Virtual environment
* Generated output
* Environment variables
* Temporary Python files

## Screenshots

### Automation workflow

![Automation workflow](docs/workflow.png)

### Sample output

![Sample output](docs/output_sample.png)

## Learning Outcomes

Through this project, I learned:

* How to create and activate a Python virtual environment
* How to install and use Playwright
* How to work with Excel input
* How to automate browser interactions
* How to handle waiting and element-selection problems
* How to protect private files using `.gitignore`
* How to manage a Python project using Git and GitHub

## Limitations

* WhatsApp Web interface changes can affect element selectors.
* A stable internet connection is required.
* The user must log in to WhatsApp Web.
* The program should only be used with permission from message recipients.

## Disclaimer

This project is created for educational purposes. Users are responsible for following WhatsApp’s applicable terms, privacy requirements, and anti-spam rules.
