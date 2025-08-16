# Email Automation App

![Email Automation App Screenshot](https://i.postimg.cc/1RDrJ4rV/Screenshot-2025-08-17-060026.png)

A **desktop application** to efficiently manage and send emails to multiple recipients. Designed with a Gmail-like interface, this app allows you to compose messages, attach files, maintain recipient lists, and configure SMTP settings, all within a sleek and intuitive GUI.

---

## Features

- **Gmail-like Interface:** Three-panel layout with left sidebar for easy navigation.  
- **Send Mail Page:** Compose subject, body, and add multiple attachments.  
- **Email List Management:** Add, import, and save multiple recipients.  
- **SMTP Configuration:** Import or configure SMTP settings securely and persist them locally.  
- **Non-blocking Email Sending:** Send to all recipients with a live progress log using QThread.  
- **Draft Management:** Save your draft including subject, body, and attachments for later use.  
- **About Page:** Developer info, GitHub repository link, and sponsorship support.  

---

## Screenshots

**Send Mail Page**  
[![Screenshot](https://i.postimg.cc/1RDrJ4rV/Screenshot-2025-08-17-060026.png)](https://postimg.cc/NLfX0sVQ)

**Email List Page**  
[![Screenshot](https://i.postimg.cc/C5f9cY4Q/Screenshot-2025-08-17-060045.png)](https://postimg.cc/xNnsdwhv)

**SMTP Config Page**  
[![Screenshot](https://i.postimg.cc/xjwdn1Bq/Screenshot-2025-08-17-060057.png)](https://postimg.cc/V5FwFmfc)

---

## Installation

1. Clone this repository:

```bash
git clone https://github.com/tanvir-ahamed04/Email-Automation-Desktop-App-PyQt5.git
cd Email-Automation-Desktop-App-PyQt5
````

2. Install dependencies:

```bash
pip install PyQt5
```

3. Run the app:

```bash
python app.py
```

---

## Usage

1. **Add recipients:** Go to the Email List page and add email addresses manually or import from a `.txt` or `.csv` file.
2. **Compose your email:** Go to the Send Mail page, write your subject and body, and attach files if needed.
3. **Configure SMTP:** Go to SMTP Config, import or manually set your SMTP server, username, and password.
4. **Send emails:** Click “Send to All” and monitor the progress log.

**Tip:** Drafts are automatically saved locally, so you can continue your work later.

---

## Download

You can download the latest Windows version of the Email Automation App here:

- [Download Windows EXE](https://github.com/tanvir-ahamed04/Email-Automation-Desktop-App-PyQt5-/releases/download/v1.0/Bulk-Email-Sender.exe)


## Security Note

This app stores SMTP credentials and email lists in local JSON files for convenience. **For production use, consider using OS keyring or secure credential storage.**

---

## About the Developer

* **Name:** Tanvir Ahamed
* **GitHub:** [Email Automation Desktop App](https://github.com/tanvir-ahamed04/Email-Automation-Desktop-App-PyQt5)
* **Contact:** [hireme.tanvir@gmail.com](mailto:hireme.tanvir@gmail.com)

---

## License

This project is open source and available under the MIT License.
