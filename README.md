# 📲 WhatsApp Bulk Sender (Desktop GUI)

A modern desktop application for sending **bulk WhatsApp messages** using an Excel file.
Built with **Python**, **Tkinter / CustomTkinter**, and automation tools, this app allows you to run messaging campaigns with progress tracking, delay control, and reporting.

---

## 🚀 Features

### 📁 Excel-Based Campaigns

* Upload `.xlsx` files containing:

  * Column A → Phone Numbers
  * Column B → Messages
* Automatically detects number of contacts

### 📤 Bulk WhatsApp Messaging

* Sends messages via WhatsApp Web using `pywhatkit`
* Works with your **already logged-in browser session**
* Supports **instant messaging (no scheduling required)**

### 📊 Real-Time Progress Dashboard

* Live progress bar
* Metrics:

  * ✅ Successful messages
  * ❌ Failed messages
  * ⏳ Pending messages
* Percentage completion tracking

### ⏱ Smart Delay System

* Custom delay between messages (in seconds)
* Adds **random delay (2–6 sec)** to reduce spam detection risk

### 📄 Campaign Report Generation

* Automatically generates `report.xlsx` with:

  * Phone number
  * Status (Sent / Failed / Invalid)
  * Timestamp

### 📞 Phone Number Validation

* Uses `phonenumbers` library to:

  * Normalize numbers into international format
  * Filter invalid numbers before sending

### 🖥 Modern UI

* Built with `customtkinter`
* Clean, responsive interface
* Drag-and-drop inspired upload area
* Interactive file preview card

---

## 🧠 How It Works

1. **Upload Excel File**

   * Select `.xlsx` file with contacts & messages

2. **Configure Delay**

   * Set delay between messages to avoid WhatsApp restrictions

3. **Launch Campaign**

   * App processes each row:

     * Validates phone number
     * Sends message
     * Logs result

4. **Track Progress**

   * UI updates in real-time

5. **Get Report**

   * Final report saved as `report.xlsx`

---

## 📂 Project Structure

```
.
├── autoWhatsapp.py        # Main application
├── images/                # UI assets (icons, images)
├── whatsapp_profile/      # Optional folder for login state detection
├── report.xlsx            # Generated after sending
└── icon.ico               # App icon
```

---

## 🛠 Requirements

Install dependencies:

```bash
pip install customtkinter pillow pandas openpyxl pywhatkit phonenumbers
```

---

## ▶️ Running the App

```bash
python autoWhatsapp.py
```

---

## 📦 Build Executable (Optional)

Using PyInstaller:

```bash
pyinstaller --onefile --windowed --icon=icon.ico --add-data "images/*.png:images" autoWhatsapp.py
```

---

## ⚠️ Important Notes

* ⚠️ **WhatsApp must already be logged in** via your browser
* ⚠️ Too many messages without delay may lead to **temporary bans**
* ⚠️ Use responsibly — this tool is intended for:

  * Notifications
  * Business communication
  * Educational purposes

---

## 📊 Excel File Format

| Phone Number | Message            |
| ------------ | ------------------ |
| 201XXXXXXXXX | Hello from app!    |
| 966XXXXXXXXX | Your order shipped |

* No `+` required (handled automatically)
* Must be valid international numbers

---

## 🔧 Core Components

### `Auto_Whatsapp` Class

Handles:

* Phone normalization
* Message sending via `pywhatkit`

### `WhatsAppGUI` Class

Handles:

* UI rendering
* File upload
* Progress tracking
* Threaded message sending

---

## ⚡ Performance & Design Choices

* Uses **threading** to keep UI responsive
* Uses **Tkinter `after()`** for safe UI updates
* Implements **random delays** to mimic human behavior
* Separates backend logic from UI

---

## 📈 Future Improvements (Ideas)

* Add CSV support
* Message templates (variables like `{name}`)
* Contact grouping
* Scheduling campaigns
* Selenium-based WhatsApp control
* Retry failed messages

---

## 🧑‍💻 Author Notes

This project demonstrates:

* Desktop GUI development in Python
* Automation with external services
* File handling and reporting
* Multithreading in UI apps

---

## 📜 License

This project is for educational and personal use.
Use responsibly and comply with WhatsApp policies.

---

## 💡 Tip

Start with small batches (5–10 messages) to test before launching large campaigns.

---

**Enjoy automating your WhatsApp workflows 🚀**
