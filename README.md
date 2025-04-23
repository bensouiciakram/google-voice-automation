# Google Voice Message Automation 📩

This Python script automates the process of sending messages via [Google Voice](https://voice.google.com/) using data stored in an Excel file. It utilizes the **Playwright** browser automation library along with **Win32 COM** to interact with Microsoft Excel.

---

## 📌 Features

- Logs in to Google Voice using credentials from Excel.
- Sends messages to specified phone numbers.
- Includes customizable delays between messages.
- Uses stealth mode to avoid bot detection.

---

## 📂 Folder Structure


---

## ⚙️ Requirements

- Windows OS (due to `win32com.client`)
- Excel application must be open with the correct file
- Python 3.7+
- Google account with Google Voice access

---

## 📦 Installation

1. **Clone the repository**
   ```bash
   git clone https://github.com/bensouiciakram/google-voice-automation.git
   cd google-voice-automation

## 📄 Excel Format
- Sheet 2 (Text Messages)
  - Column M: Phone Number
  - Column N: Message Text
  - Column O: Delay in seconds
- Sheet 3 (Authentication)
   - Cell A2: Email
   - Cell B2: Password
  
 Ensure the Excel file is open and loaded in the correct order before running the script.

