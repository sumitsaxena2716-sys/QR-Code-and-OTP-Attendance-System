# 📌 QR Code Attendance System

A **Smart Attendance System** built using **Python (Flask), HTML, CSS, JavaScript, and Excel automation**.

This system uses **QR Code scanning** to mark attendance efficiently and securely.

---

## 🚀 Features

✅ QR Code based attendance  
✅ Time-based status (Present / Late / Absent)  
✅ Teacher override (password protected)  
✅ Automatic Excel report generation  
✅ Monthly attendance sheet auto-create  

### 🎨 Excel Highlighting
🟢 Present (Green)  
🔵 Late (Blue)  
🔴 Absent (Red)  
🔴 Sundays highlighted  

✅ Dashboard (Daily + Monthly analytics)  
✅ Student image display on success  

---

## 🧠 Working Logic

| Time | Status |
|------|--------|
| Before 9:05 AM | Present |
| 9:05 – 9:15 AM | Late |
| After 9:15 AM | Teacher Permission Required |
| No Scan | Absent |

---

## 📂 Folder Structure

QR-Code-and-OTP-Attendance-System/
│
├── app.py                    # Main Flask Backend
├── students.xlsx             # Student Database
├── April_Attendance.xlsx     # Auto-generated Monthly File
│
├── templates/                # Frontend HTML Pages
│   ├── index.html
│   ├── scanner.html
│   ├── success.html
│   ├── login.html
│   ├── dashboard.html
│
├── static/                   # JS Files
│   └── login.js
│
├── images/                   # Student Images
│   ├── 101.jpg
│   ├── 102.jpg
│
└── README.md

---

## 🛠️ Technologies Used

- Python (Flask)  
- HTML, CSS, JavaScript  
- OpenCV / HTML5 QR Scanner  
- Pandas  
- OpenPyXL (Excel Automation)  
- SMTP (Email Sending)  

---

## ⚙️ Installation & Setup

### 1️⃣ Clone Repository
```bash
git clone https://github.com/sumitsaxena2716-sys/QR-Code-and-OTP-Attendance-System.git
cd QR-Code-and-OTP-Attendance-System
2️⃣ Install Dependencies
pip install flask pandas openpyxl qrcode
3️⃣ Run Project
python app.py
4️⃣ Open Browser
http://127.0.0.1:5000
🔐 Teacher Login

Username: Admin
Password: admin@123

📧 Email Setup

In app.py:

SENDER_EMAIL = "your_email@gmail.com"
APP_PASSWORD = "your_app_password"

👉 Use Gmail App Password (not normal password)
