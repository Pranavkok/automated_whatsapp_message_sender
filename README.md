# 📲 Automated WhatsApp Message Sender Using Python

This project allows you to **automatically send WhatsApp messages** to a list of contacts from an Excel file using Python. It's perfect for quick outreach, broadcasting updates, or sending event reminders — all through WhatsApp Web.

---

## ✅ Features

- Reads contact numbers from an Excel file.
- Automatically formats phone numbers (adds `+91` if missing).
- Sends custom messages using WhatsApp Web.
- Gracefully handles missing or invalid contact numbers.
- Minimal setup and user-friendly logic.

---

## 🛠️ Requirements

Make sure you have Python installed, then install the required libraries:

```bash
pip install openpyxl pywhatkit
```

Also ensure:
- You're **logged into WhatsApp Web** on your default browser.
- Your Excel file has a valid column name like `"contact"`, `"phone"`, `"mobile"`, etc.

---

## 📂 Excel File Format

- The contact numbers should be in the **first row (header)**.
- Accepted column names (**case-insensitive**):
  ```
  contact, phone, mobile, number, phonenumber, mobilenumber, contactnumber, tel
  ```

### ✅ Example:

| Name       | Contact        |
|------------|----------------|
| Alice      | 9876543210     |
| Bob        | +919876543210  |

---

## 🚀 How to Use

1. Edit the `path` variable in the script to your Excel file path:
   ```python
   path = "/your/path/to/testContacts.xlsx"
   ```

2. Run the script:
   ```bash
   python whatsapp_sender.py
   ```

3. The script will:
   - Identify the correct contact column.
   - Format the numbers correctly.
   - Open WhatsApp Web and send the message instantly.
   - Close the tab after sending.

---

## 📌 Script Overview

```python
def send_whatsapp_message(contact, message):
    pywhatkit.sendwhatmsg_instantly(contact, message, wait_time=10, tab_close=True)
```

- Adds `+91` prefix if not present.
- Filters valid 10-digit numbers.
- Sends a test message with `pywhatkit`.

---

## 💡 Open to Contribute 

You can enhance this project with the following features:

1. **Message Scheduling**
   - Schedule messages at specific times (use `sendwhatmsg()` with hour/minute).

2. **Multiple Contact Columns Support**
   - Detect and merge numbers from multiple columns (e.g., `work`, `home`, `mobile`).

3. **Delivery Logging**
   - Save a success/failure log in a CSV/Excel file with timestamps.

4. **GUI Interface**
   - Build a simple GUI using `tkinter` or `PyQt` to select the Excel file and preview messages.

5. **Country Code Flexibility**
   - Detect and format numbers from other countries besides India.

6. **Attachment Support**
   - Automate sending images, PDFs, or media (using Selenium instead of pywhatkit).

7. **Error Retry Logic**
   - Retry sending messages if the first attempt fails.


