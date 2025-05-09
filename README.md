# Web_scrapying-2.0
# 🛠️ Amazon Automation Tool - Timetec

Automate customer messaging for Amazon orders through Seller Central using a graphical interface built with **PyQt5** and **Selenium**.

<img src="https://github.com/bojunz/Web_scrapying-2.0/blob/main/Pyqt_original.png" alt="Demo Login Page GIF" style="border: 2px solid black; max-width: 200%; height: 400px;">
---

## 📌 Table of Contents

- [📦 Features](#-features)
- [🧠 Background](#-background)
- [🖥️ User Interface](#-user-interface)
- [⚙️ How It Works](#️-how-it-works)
- [🛠 Installation Guide](#-installation-guide)
- [🚀 Usage Instructions](#-usage-instructions)
- [📝 Message Templates](#-message-templates)
- [📤 Export & Logs](#-export--logs)
- [🔐 Notes & Precautions](#-notes--precautions)
- [🤝 Support](#-support)
- [📄 License](#-license)

---

## 📦 Features

- GUI-based tool for easy use.
- Login panel with:
  - Email and password
  - Country selection
- Automatically detects the customer's language (English or Spanish).
- Optional custom message input.
- Processes multiple Amazon Order IDs (comma-separated).
- Interacts with Shadow DOM using JavaScript injection for robust performance.
- Order and system logs displayed in real-time.
- Export communication results to CSV file.
- Shortcut support (e.g., press `Enter` to start automation).
- Visual feedback and error handling with log output.

---

## 🧠 Background

Customer satisfaction and proactive engagement are critical for e-commerce success. This tool helps automate follow-up communication with Amazon customers, saving time and improving service quality for the Timetec team.

---

## 🖥️ User Interface

### Left Panel

- **Login Form**: Enter your Timetec Amazon credentials and country.
- **Order IDs**: Paste one or more order IDs (comma-separated).
- **Custom Message**: Optional text area to override the default message.
- **Start Automation**: Initiates the process.
- **Support Info**: Displays contact for technical support.

### Right Panel

- **Order Information Log**: Logs per order (success/failure).
- **General Log**: Logs full system operations and debug info.
- **Buttons**:
  - `Clear Content`: Clears logs.
  - `Export to CSV`: Exports results to CSV format.

---

## ⚙️ How It Works

1. **Login to Amazon Seller Central** via Selenium automation.
2. Navigate to each provided **Order ID**.
3. Detect the customer's language preference (English/Spanish).
4. Compose and send a tailored message.
5. Handle Shadow DOM inputs via JavaScript.
6. Log results and errors.
7. Export messages and statuses to a `.csv` file.

---

## 🛠 Installation Guide

### Requirements

- Python 3.7+
- Google Chrome installed
- Required Python packages (see below)

### Installation Steps

```bash
git clone https://github.com/yourusername/amazon-automation-tool.git
cd amazon-automation-tool
pip install -r requirements.txt
