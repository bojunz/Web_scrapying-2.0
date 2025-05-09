# 🧾 Timetec Amazon Automation Tool

A user-friendly GUI tool using **PyQt5** and **Selenium** to help **Timetec** sellers automatically send messages to Amazon customers via Seller Central.

Supports **manual input** and **Excel-based batch messaging** with full logs, timers, and export features.

## Original Version
<img src="https://github.com/bojunz/Web_scrapying-2.0/blob/main/Pyqt_original.png" alt="Demo Login Page GIF" style="border: 2px solid black; max-width: 200%; height: 400px;"> <br>
## Excel Version
<img src="https://github.com/bojunz/Web_scrapying-2.0/blob/main/Pyqt_Excel.png" alt="Demo Login Page GIF" style="border: 2px solid black; max-width: 200%; height: 400px;">
---

## 📌 Table of Contents

- [🚀 Features](#-features)
- [🧠 Overview](#-overview)
- [🖥️ Interface Layout](#️-interface-layout)
- [📄 Excel Format (for Batch Mode)](#-excel-format-for-batch-mode)
- [⚙️ Installation](#-installation)
- [📂 How to Use](#-how-to-use)
  - [Manual Mode](#manual-mode)
  - [Excel Mode](#excel-mode)
- [📤 Export & Save Results](#-export--save-results)
- [🛟 Support](#-support)
- [📄 License](#-license)

---

## 🚀 Features

✅ Manual input for one-time message sending  
✅ Batch automation using Excel files  
✅ Amazon login with email, password, and marketplace selector  
✅ Real-time order and system logs  
✅ Shadow DOM interaction with Selenium  
✅ Auto-detect and confirm buyer language (e.g., Spanish)  
✅ Elapsed time display & success/failure tracking  
✅ One-click CSV export  
✅ Press `Enter` to trigger automation

---

## 🧠 Overview

This tool helps sellers streamline communication with customers on Amazon by:

- Supporting **two automation modes**:
  - 📬 **Manual Mode** – Send message to a single Order ID with a custom message
  - 📊 **Excel Mode** – Load and send messages to many Order IDs via spreadsheet
- Logging into Seller Central, navigating to each order, and messaging the buyer
- Detecting language preferences and using shadow DOM to inject messages
- Logging results and enabling CSV export

---

## 🖥️ Interface Layout

### 🔒 Left Panel

- **Login Section**
  - Email, password
  - Country dropdown (e.g., `United States`, `France`, etc.)
- **Manual Mode Section**
  - Input Order ID and message
  - (Optional) Send one message at a time
- **Excel Mode Section**
  - Load `.xlsx` file with multiple orders
  - View progress bar
- **Summary**
  - Time elapsed
  - Success/Failure count
  - Real-time clock
- **Start Automation Button**
  - Triggers automation
  - Also works with `Enter` key

### 📋 Right Panel

- **Order Log**
  - Displays success/failure per order
- **System Log**
  - Logs application/system-level events and errors
- **Action Buttons**
  - `Clear Content`: Clears the order log
  - `Export to CSV`: Saves results to file

---

## 📄 Excel Format (for Batch Mode)

To use Excel automation, your `.xlsx` file must include these columns:

| Column Name | Description                                     |
|-------------|-------------------------------------------------|
| ID          | Amazon Order ID (e.g. `112-1234567-8901234`)    |
| Content     | The message to send                             |
| Date        | Optional — for tracking                         |
| Market      | Optional — e.g., "US", "EU"                     |
| Status      | Leave blank — tool fills in "Success"/"Failed"  |

### ✅ Example:

| ID                 | Content                              | Date       | Market | Status |
|--------------------|---------------------------------------|------------|--------|--------|
| 112-1234567-8901234| Thanks for your order! Let us know…   | 2025-05-05 | USA    |        |

---

## ⚙️ Installation

### 1. Requirements

- Python 3.7 or newer
- Google Chrome (latest version)
- Pip dependencies:

```bash
pip install -r requirements.txt
