# Automatic_Email_Sender_GUI

A desktop Python application to send personalized nomination reminder emails using Outlook. Built for internal use at the company to automate the Value Awards communication process.

---

## ✨ Features

- Connects to SQL Server to fetch user data
- Displays users in a searchable/sortable table
- Filter users by `Title` using a dropdown
- Click-to-sort by FullName, Email, or Title
- Sends personalized HTML emails via Outlook
- Embeds the Ipsos logo in the email and the app window
- Clean GUI built using `tkinter`

---

## 📦 Requirements

- Windows with Microsoft Outlook installed
- Python 3.10+ (or compatible version)

### 📋 Python Packages

Install dependencies:

```bash
pip install pywin32 pyodbc Pillow
