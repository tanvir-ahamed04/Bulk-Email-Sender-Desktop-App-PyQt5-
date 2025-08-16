"""
Email Automation Desktop App (PyQt5)

Features
- Gmail-like three-panel layout with left sidebar
- Pages: Send Mail (default), Email List, SMTP Config, About
- Import SMTP settings from an external JSON file; persist config locally
- Persist email list locally (JSON)
- Compose subject/body, add multiple attachments
- Send to all saved recipients with progress log
- Test SMTP connection
- Non-blocking sending with QThread worker

How to run
1) pip install PyQt5
2) python app.py

Optional external import
- From the SMTP Config page, click "Import Config" to load a JSON file, e.g.:
  {
    "smtp_host": "smtp.gmail.com",
    "smtp_port": 587,
    "use_tls": true,
    "use_ssl": false,
    "username": "your@email.com",
    "password": "your_app_password"
  }

Security note
- This demo stores credentials in plain JSON on disk to match the requested behavior.
  For production, use the OS keyring/credential store.
"""

import os
import sys
import json
import ssl
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

from PyQt5.QtCore import Qt, QThread, pyqtSignal, QObject
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QFileDialog,
    QHBoxLayout, QVBoxLayout,
    QPushButton, QLabel, QLineEdit, QTextEdit, QListWidget,
    QStackedWidget, QGroupBox, QFormLayout, QSpinBox, QCheckBox, QMessageBox
)

# ------------------- Helper for PyInstaller ------------------- #
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and PyInstaller exe """
    try:
        base_path = sys._MEIPASS  # PyInstaller temporary folder
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# ------------------- Paths for JSON files ------------------- #
APP_DIR = os.path.abspath(os.path.dirname(__file__))
CONFIG_PATH = resource_path("config.json")
EMAILS_PATH = resource_path("emails.json")
DRAFT_PATH = resource_path("draft.json")

# -------------------------- Utility: Config Manager -------------------------- #
class ConfigManager:
    @staticmethod
    def load_json(path, default):
        try:
            if os.path.exists(path):
                with open(path, 'r', encoding='utf-8') as f:
                    return json.load(f)
        except Exception:
            pass
        return default

    @staticmethod
    def save_json(path, data):
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2)

    @staticmethod
    def load_smtp():
        return ConfigManager.load_json(CONFIG_PATH, {
            "smtp_host": "",
            "smtp_port": 587,
            "use_tls": True,
            "use_ssl": False,
            "username": "",
            "password": ""
        })

    @staticmethod
    def save_smtp(cfg):
        ConfigManager.save_json(CONFIG_PATH, cfg)

    @staticmethod
    def load_emails():
        data = ConfigManager.load_json(EMAILS_PATH, {"recipients": []})
        uniq, seen = [], set()
        for e in data.get("recipients", []):
            e2 = e.strip()
            if e2 and e2 not in seen:
                uniq.append(e2)
                seen.add(e2)
        return {"recipients": uniq}

    @staticmethod
    def save_emails(lst):
        uniq, seen = [], set()
        for e in lst:
            e2 = e.strip()
            if e2 and e2 not in seen:
                uniq.append(e2)
                seen.add(e2)
        ConfigManager.save_json(EMAILS_PATH, {"recipients": uniq})

    @staticmethod
    def load_draft():
        return ConfigManager.load_json(DRAFT_PATH, {
            "subject": "",
            "body": "",
            "attachments": []
        })

    @staticmethod
    def save_draft(draft):
        atts = [p for p in draft.get("attachments", []) if os.path.exists(p)]
        draft = {
            "subject": draft.get("subject", ""),
            "body": draft.get("body", ""),
            "attachments": atts
        }
        ConfigManager.save_json(DRAFT_PATH, draft)


# ----------------------------- Email Sender Worker --------------------------- #
class SendWorker(QObject):
    progress = pyqtSignal(str)
    finished = pyqtSignal()
    error = pyqtSignal(str)

    def __init__(self, smtp_cfg, recipients, subject, body, attachments):
        super().__init__()
        self.smtp_cfg = smtp_cfg
        self.recipients = recipients
        self.subject = subject
        self.body = body
        self.attachments = attachments
        self._stop = False

    def stop(self):
        self._stop = True

    def _build_message(self, sender, recipient):
        msg = MIMEMultipart()
        msg["From"] = sender
        msg["To"] = recipient
        msg["Subject"] = self.subject
        msg.attach(MIMEText(self.body, 'plain'))
        for file in self.attachments:
            try:
                with open(file, 'rb') as f:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(f.read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(file)}"')
                msg.attach(part)
            except Exception as e:
                self.progress.emit(f"Attachment error: {file} ({e}) — skipping this file.")
        return msg

    def run(self):
        host = self.smtp_cfg.get("smtp_host", "")
        port = int(self.smtp_cfg.get("smtp_port", 587))
        use_tls = bool(self.smtp_cfg.get("use_tls", True))
        use_ssl = bool(self.smtp_cfg.get("use_ssl", False))
        username = self.smtp_cfg.get("username", "")
        password = self.smtp_cfg.get("password", "")

        if not host or not username or not password:
            self.error.emit("SMTP settings incomplete. Please configure SMTP.")
            self.finished.emit()
            return

        context = ssl.create_default_context()

        try:
            if use_ssl:
                server = smtplib.SMTP_SSL(host, port, context=context)
            else:
                server = smtplib.SMTP(host, port)
            with server:
                server.ehlo()
                if use_tls and not use_ssl:
                    server.starttls(context=context)
                    server.ehlo()
                if username:
                    server.login(username, password)

                total = len(self.recipients)
                for idx, rcpt in enumerate(self.recipients, start=1):
                    if self._stop:
                        self.progress.emit("Sending cancelled by user.")
                        break
                    try:
                        msg = self._build_message(username, rcpt)
                        server.sendmail(username, [rcpt], msg.as_string())
                        self.progress.emit(f"[{idx}/{total}] Sent to {rcpt}")
                    except Exception as e:
                        self.progress.emit(f"[{idx}/{total}] Failed to {rcpt}: {e}")
        except Exception as e:
            self.error.emit(f"SMTP error: {e}")
        finally:
            self.finished.emit()


# --------------------------------- UI Pages --------------------------------- #
class EmailListPage(QWidget):
    changed = pyqtSignal()

    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)

        title = QLabel("Email List")
        title.setStyleSheet("font-size: 18px; font-weight: 600;")
        layout.addWidget(title)

        hint = QLabel("Enter one email per line. Example:\n1. user@example.com\n2. other@example.com")
        hint.setStyleSheet("color: gray;")
        layout.addWidget(hint)

        self.text = QTextEdit()
        self.text.setPlaceholderText("1. user@example.com\n2. other@example.com\n...")
        self.text.setStyleSheet("QTextEdit { padding: 10px; }")
        layout.addWidget(self.text)

        btns = QHBoxLayout()
        self.btn_load = QPushButton("Load from File (.txt/.csv)")
        self.btn_save = QPushButton("Save List")
        btns.addWidget(self.btn_load)
        btns.addWidget(self.btn_save)
        layout.addLayout(btns)

        self.btn_load.clicked.connect(self.load_from_file)
        self.btn_save.clicked.connect(self.save_list)

    def set_emails(self, emails):
        lines = [f"{i}. {e}" for i, e in enumerate(emails, start=1)]
        self.text.setPlainText("\n".join(lines))

    def get_emails(self):
        raw = self.text.toPlainText().splitlines()
        emails = []
        for line in raw:
            line = line.strip()
            if not line:
                continue
            if "." in line:
                try:
                    left, right = line.split(".", 1)
                    if left.strip().isdigit():
                        line = right.strip()
                except ValueError:
                    pass
            emails.append(line)
        uniq, seen = [], set()
        for e in emails:
            e2 = e.strip()
            if e2 and e2 not in seen:
                uniq.append(e2)
                seen.add(e2)
        return uniq

    def load_from_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Open emails file", APP_DIR, "Text/CSV (*.txt *.csv)")
        if not path:
            return
        try:
            with open(path, 'r', encoding='utf-8') as f:
                lines = [ln.strip() for ln in f if ln.strip()]
            if len(lines) == 1 and "," in lines[0]:
                parts = [p.strip() for p in lines[0].split(",") if p.strip()]
            else:
                parts = []
                for ln in lines:
                    parts.extend([p.strip() for p in ln.split(",") if p.strip()])
            self.set_emails(parts)
            self.changed.emit()
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to read file:\n{e}")

    def save_list(self):
        emails = self.get_emails()
        ConfigManager.save_emails(emails)
        QMessageBox.information(self, "Saved", f"Saved {len(emails)} recipients.")
        self.changed.emit()


class SendMailPage(QWidget):
    request_send = pyqtSignal(dict)  # payload: {subject, body, attachments}

    def __init__(self):
        super().__init__()

        layout = QVBoxLayout(self)

        title = QLabel("Compose & Send")
        title.setStyleSheet("font-size: 18px; font-weight: 600;")
        layout.addWidget(title)

        form = QFormLayout()
        self.subject = QLineEdit()
        self.body = QTextEdit()
        self.body.setPlaceholderText("Write your message here...")
        self.body.setStyleSheet("QTextEdit { padding: 10px; }")

        form.addRow("Subject:", self.subject)
        form.addRow("Body:", self.body)

        att_group = QGroupBox("Attachments")
        att_layout = QVBoxLayout()
        self.att_list = QListWidget()
        btns = QHBoxLayout()
        self.btn_add_att = QPushButton("Add Attachment")
        self.btn_del_att = QPushButton("Remove Selected")
        btns.addWidget(self.btn_add_att)
        btns.addWidget(self.btn_del_att)

        att_layout.addWidget(self.att_list)
        att_layout.addLayout(btns)
        att_group.setLayout(att_layout)

        layout.addLayout(form)
        layout.addWidget(att_group)

        bottom = QHBoxLayout()
        self.lbl_recipients = QLabel("Recipients: 0")
        bottom.addWidget(self.lbl_recipients)
        bottom.addStretch(1)
        self.btn_save_draft = QPushButton("Save Draft")
        self.btn_send_all = QPushButton("Send to All")
        bottom.addWidget(self.btn_save_draft)
        bottom.addWidget(self.btn_send_all)
        layout.addLayout(bottom)

        self.btn_add_att.clicked.connect(self.add_attachment)
        self.btn_del_att.clicked.connect(self.remove_attachment)
        self.btn_save_draft.clicked.connect(self.save_draft)
        self.btn_send_all.clicked.connect(self.emit_send)

        self.load_draft()

        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setPlaceholderText("Sending log will appear here...")
        layout.addWidget(self.log)

    def set_recipient_count(self, n):
        self.lbl_recipients.setText(f"Recipients: {n}")

    def load_draft(self):
        draft = ConfigManager.load_draft()
        self.subject.setText(draft.get("subject", ""))
        self.body.setPlainText(draft.get("body", ""))
        self.att_list.clear()
        for p in draft.get("attachments", []):
            if os.path.exists(p):
                self.att_list.addItem(p)

    def save_draft(self):
        draft = {
            "subject": self.subject.text().strip(),
            "body": self.body.toPlainText(),
            "attachments": [self.att_list.item(i).text() for i in range(self.att_list.count())]
        }
        ConfigManager.save_draft(draft)
        QMessageBox.information(self, "Saved", "Draft saved.")

    def add_attachment(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Select attachments", APP_DIR, "All Files (*.*)")
        for f in files:
            if f:
                self.att_list.addItem(f)

    def remove_attachment(self):
        for item in self.att_list.selectedItems():
            self.att_list.takeItem(self.att_list.row(item))

    def append_log(self, text):
        self.log.append(text)

    def clear_log(self):
        self.log.clear()

    def emit_send(self):
        payload = {
            "subject": self.subject.text().strip(),
            "body": self.body.toPlainText(),
            "attachments": [self.att_list.item(i).text() for i in range(self.att_list.count())]
        }
        self.request_send.emit(payload)


class SmtpConfigPage(QWidget):
    changed = pyqtSignal()

    def __init__(self):
        super().__init__()

        layout = QVBoxLayout(self)
        title = QLabel("SMTP Configuration")
        title.setStyleSheet("font-size: 18px; font-weight: 600;")
        layout.addWidget(title)

        form = QFormLayout()
        self.host = QLineEdit()
        self.port = QSpinBox()
        self.port.setRange(1, 65535)
        self.username = QLineEdit()
        self.password = QLineEdit()
        self.password.setEchoMode(QLineEdit.Password)
        self.use_tls = QCheckBox("Use STARTTLS")
        self.use_ssl = QCheckBox("Use SSL")

        form.addRow("SMTP Host:", self.host)
        form.addRow("SMTP Port:", self.port)
        form.addRow("Username:", self.username)
        form.addRow("Password:", self.password)
        form.addRow(self.use_tls)
        form.addRow(self.use_ssl)

        layout.addLayout(form)

        btns = QHBoxLayout()
        self.btn_import = QPushButton("Import Config (JSON)")
        self.btn_save = QPushButton("Save Settings")
        self.btn_test = QPushButton("Test Connection")
        btns.addWidget(self.btn_import)
        btns.addStretch(1)
        btns.addWidget(self.btn_test)
        btns.addWidget(self.btn_save)
        layout.addLayout(btns)

        self.load_config()

        self.btn_save.clicked.connect(self.save_config)
        self.btn_test.clicked.connect(self.test_connection)
        self.btn_import.clicked.connect(self.import_config)
        self.use_ssl.stateChanged.connect(self.on_ssl_changed)

    def on_ssl_changed(self):
        if self.use_ssl.isChecked():
            if self.port.value() == 587:
                self.port.setValue(465)

    def load_config(self):
        cfg = ConfigManager.load_smtp()
        self.host.setText(cfg.get("smtp_host", ""))
        self.port.setValue(int(cfg.get("smtp_port", 587)))
        self.username.setText(cfg.get("username", ""))
        self.password.setText(cfg.get("password", ""))
        self.use_tls.setChecked(bool(cfg.get("use_tls", True)))
        self.use_ssl.setChecked(bool(cfg.get("use_ssl", False)))

    def save_config(self):
        cfg = {
            "smtp_host": self.host.text().strip(),
            "smtp_port": int(self.port.value()),
            "username": self.username.text().strip(),
            "password": self.password.text(),
            "use_tls": self.use_tls.isChecked(),
            "use_ssl": self.use_ssl.isChecked()
        }
        ConfigManager.save_smtp(cfg)
        QMessageBox.information(self, "Saved", "SMTP settings saved.")
        self.changed.emit()

    def import_config(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select SMTP config JSON", APP_DIR, "JSON (*.json)")
        if not path:
            return
        try:
            with open(path, 'r', encoding='utf-8') as f:
                cfg = json.load(f)
            merged = ConfigManager.load_smtp()
            merged.update(cfg)
            ConfigManager.save_smtp(merged)
            self.load_config()
            QMessageBox.information(self, "Imported", "SMTP configuration imported and saved.")
            self.changed.emit()
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to import config:\n{e}")

    def test_connection(self):
        cfg = {
            "smtp_host": self.host.text().strip(),
            "smtp_port": int(self.port.value()),
            "username": self.username.text().strip(),
            "password": self.password.text(),
            "use_tls": self.use_tls.isChecked(),
            "use_ssl": self.use_ssl.isChecked()
        }
        try:
            context = ssl.create_default_context()
            if cfg["use_ssl"]:
                server = smtplib.SMTP_SSL(cfg["smtp_host"], cfg["smtp_port"], context=context)
            else:
                server = smtplib.SMTP(cfg["smtp_host"], cfg["smtp_port"])
            with server:
                server.ehlo()
                if cfg["use_tls"] and not cfg["use_ssl"]:
                    server.starttls(context=context)
                    server.ehlo()
                if cfg["username"]:
                    server.login(cfg["username"], cfg["password"])
            QMessageBox.information(self, "Success", "SMTP connection successful.")
        except Exception as e:
            QMessageBox.warning(self, "Failed", f"Connection failed:\n{e}")


# ------------------------------- About Page ---------------------------------- #
class AboutPage(QWidget):
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)
        about_page = QWidget()
        about_layout = QVBoxLayout(about_page)
        layout.addWidget(about_page)

        info = QLabel("""
            <h1>Bulk Email Sender</h1>
            <p>Developed by: <b>Tanvir Ahamed</b></p>
            <p>GitHub: <a href='https://github.com/tanvir-ahamed04/Email-Automation-Desktop-App-PyQt5-'>View Repository</a></p>
            <p>Contact: <a href='mailto:hireme.tanvir@gmail.com'>hireme.tanvir@gmail.com</a></p>
            <p><i>Support this project:</i></p>
            <p><a href='https://github.com/sponsors/tanvir-ahamed04'>Become a Sponsor</a></p>
        """)
        info.setOpenExternalLinks(True)  # so links are clickable
        info.setAlignment(Qt.AlignCenter)  # center text inside QLabel

        about_layout.addStretch()               # push down from top
        about_layout.addWidget(info, alignment=Qt.AlignCenter)  # center label
        about_layout.addStretch()               # push up from bottom


# -------------------------------- Main Window -------------------------------- #
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Bulk Email Sender")
        self.resize(1100, 720)

        # ==== Right stack (create FIRST, so we can add pages safely) ====
        self.stack = QStackedWidget()
        self.page_send = SendMailPage()
        self.page_list = EmailListPage()
        self.page_config = SmtpConfigPage()
        self.page_about = AboutPage()

        self.stack.addWidget(self.page_send)   # index 0
        self.stack.addWidget(self.page_list)   # index 1
        self.stack.addWidget(self.page_config) # index 2
        self.stack.addWidget(self.page_about)  # index 3
        self.stack.setCurrentIndex(0)

        # ==== Left sidebar (Gmail-like) ====
        self.sidebar = QWidget()
        self.sidebar.setFixedWidth(220)
        sb_layout = QVBoxLayout(self.sidebar)
        sb_layout.setContentsMargins(12, 12, 12, 12)
        sb_layout.setSpacing(8)

        self.btn_send = QPushButton("📧 Send Mail")
        self.btn_list = QPushButton("📋 Email List")
        self.btn_config = QPushButton("⚙ SMTP Config")
        self.btn_about = QPushButton("👩🏻‍💻 ℹ About")

        for b in (self.btn_send, self.btn_list, self.btn_config, self.btn_about):
            b.setCursor(Qt.PointingHandCursor)
            b.setStyleSheet(
                "QPushButton { text-align: left; padding: 10px 12px; border-radius: 8px; }"
                "QPushButton:hover { background:#8C92AC; }"
            )
            sb_layout.addWidget(b)

        sb_layout.addStretch(1)

        # Sidebar navigation
        self.btn_send.clicked.connect(lambda: self.stack.setCurrentIndex(0))
        self.btn_list.clicked.connect(lambda: self.stack.setCurrentIndex(1))
        self.btn_config.clicked.connect(lambda: self.stack.setCurrentIndex(2))
        self.btn_about.clicked.connect(lambda: self.stack.setCurrentIndex(3))

        # ==== Main layout ====
        container = QWidget()
        layout = QHBoxLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.sidebar)
        layout.addWidget(self.stack, 1)
        self.setCentralWidget(container)

        # ==== Data wiring ====
        self.refresh_recipients_count()
        self.page_list.changed.connect(self.refresh_recipients_count)
        self.page_config.changed.connect(self.on_config_changed)
        self.page_send.request_send.connect(self.on_send_requested)

        # Thread handle
        self.thread = None
        self.worker = None

        # Minimal theme
        self.setStyleSheet(
            "QMainWindow { background: white; }"
            "QGroupBox { border: 1px solid #e6e6e6; border-radius: 10px; margin-top: 10px; }"
            "QGroupBox::title { subcontrol-origin: margin; left: 14px; padding: 0 4px; }"
            "QLineEdit, QTextEdit, QSpinBox { border: 1px solid #e6e6e6; border-radius: 8px; padding: 8px; }"
            "QListWidget { border: 1px solid #e6e6e6; border-radius: 8px; }"
            "QPushButton { background: #eeee00; color: black; border: none; padding: 10px 14px; border-radius: 10px; }"
            "QPushButton:hover { background: #8C92AC; }"
        )

    def refresh_recipients_count(self):
        emails = ConfigManager.load_emails().get("recipients", [])
        self.page_send.set_recipient_count(len(emails))

    def on_config_changed(self):
        # Reserved for future updates
        pass

    def on_send_requested(self, payload):
        emails = ConfigManager.load_emails().get("recipients", [])
        if not emails:
            QMessageBox.warning(self, "No recipients", "Please add recipients in the Email List page.")
            return
        if not payload.get("subject"):
            if QMessageBox.question(self, "No subject", "Send without a subject?",
                                    QMessageBox.Yes | QMessageBox.No) == QMessageBox.No:
                return

        smtp_cfg = ConfigManager.load_smtp()

        self.page_send.clear_log()
        self.worker = SendWorker(
            smtp_cfg=smtp_cfg,
            recipients=emails,
            subject=payload.get("subject", ""),
            body=payload.get("body", ""),
            attachments=payload.get("attachments", [])
        )
        self.thread = QThread()
        self.worker.moveToThread(self.thread)

        self.worker.progress.connect(self.page_send.append_log)
        self.worker.error.connect(lambda msg: self.page_send.append_log("ERROR: " + msg))
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(lambda: self.page_send.append_log("\nDone."))
        self.thread.finished.connect(self.cleanup_thread)

        self.thread.started.connect(self.worker.run)
        self.thread.start()

    def cleanup_thread(self):
        self.worker = None
        self.thread = None


# ---------------------------------- Entry ----------------------------------- #
def main():
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()

