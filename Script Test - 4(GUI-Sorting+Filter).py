import tkinter as tk
from tkinter import ttk, messagebox
import pyodbc
import win32com.client

# === SQL FETCH FUNCTION ===
def get_users():
    conn_str = (
        "DRIVER={ODBC Driver 18 for SQL Server};"
        "SERVER=10.170.11.250;"
        "DATABASE=IpsosRewardsAndRecognitions;"
        "UID=Reward_Admin;"
        "PWD=Reward_Mena@2025;"
        "TrustServerCertificate=yes;"
        "Encrypt=yes;"
    )

    try:
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()
        cursor.execute("SELECT FullName, Email, Title FROM dbo.ADUsers WHERE Country = 'Lebanon'")
        rows = cursor.fetchall()
        conn.close()
        return [{'FullName': row[0], 'Email': row[1], 'Title': row[2]} for row in rows]

    except Exception as e:
        messagebox.showerror("SQL Error", str(e))
        return []

# === OUTLOOK EMAIL FUNCTION ===
def send_email(to_email, full_name):
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.Subject = "Ipsos Value Awards Nomination Reminder"
        mail.To = to_email

        html = f"""
        <html>
        <body>
            <p><img src=\"cid:header001\"></p>
            <p>Dear {full_name},</p>
            <p>Just a reminder to nominate people for the Ipsos Value Awards.</p>
            <p><a href=\"http://lbsd-recog-prog/RewardsRecognitionApp/Home/Nominate\">
            Click here to nominate</a></p>
        </body>
        </html>
        """

        mail.HTMLBody = html
        attachment = mail.Attachments.Add("C:/script/header.jpg")
        attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "header001")
        mail.Send()

        messagebox.showinfo("Email Sent", f"Email sent to {full_name}")
    except Exception as e:
        messagebox.showerror("Email Error", str(e))

# === DATA CONTROL ===
all_users = []
sort_order = {'FullName': True, 'Email': True, 'Title': True}

# === LOAD USERS ===
def load_users():
    global all_users
    all_users = get_users()
    update_table(all_users)
    update_title_filter()

# === TABLE UPDATE ===
def update_table(user_list):
    tree.delete(*tree.get_children())
    for user in user_list:
        tree.insert('', 'end', values=(user['FullName'], user['Email'], user['Title']))

# === COMBOBOX FILTER ===
def update_title_filter():
    titles = sorted(set(user['Title'] for user in all_users if user['Title']))
    title_filter['values'] = ['All'] + titles
    title_filter.set('All')

# === FILTER HANDLER ===
def filter_by_title(event=None):
    selected = title_filter.get()
    if selected == "All":
        update_table(all_users)
    else:
        filtered = [u for u in all_users if u['Title'] == selected]
        update_table(filtered)

# === SORTING ===
def sort_by_column(col):
    global all_users
    order = sort_order[col]
    all_users.sort(key=lambda x: x[col] or "", reverse=not order)
    sort_order[col] = not order
    update_table(all_users)

# === SEND SELECTED ===
def send_selected():
    selected = tree.selection()
    if not selected:
        messagebox.showwarning("No Selection", "Please select a user to send email.")
        return

    for item in selected:
        values = tree.item(item, 'values')
        send_email(values[1], values[0])  # Email, FullName

# === TKINTER GUI ===
root = tk.Tk()
root.title("Ipsos Email Sender")
# Set Ipsos icon in the toolbar
root.iconphoto(False, tk.PhotoImage(file="C:/script/logo.png"))

# Filter Bar
filter_frame = tk.Frame(root)
filter_frame.pack(fill='x', pady=5)

tk.Label(filter_frame, text="Filter by Title:").pack(side='left', padx=5)
title_filter = ttk.Combobox(filter_frame, state='readonly')
title_filter.pack(side='left', padx=5)
title_filter.bind("<<ComboboxSelected>>", filter_by_title)

# Treeview Table
columns = ('FullName', 'Email', 'Title')
tree = ttk.Treeview(root, columns=columns, show='headings')
for col in columns:
    tree.heading(col, text=col, command=lambda _col=col: sort_by_column(_col))
    tree.column(col, width=200)
tree.pack(fill='both', expand=True)

# Button Panel
btn_frame = tk.Frame(root)
btn_frame.pack(fill='x', pady=5)

tk.Button(btn_frame, text="Refresh Users", command=load_users).pack(side='left', padx=5)
tk.Button(btn_frame, text="Send Email to Selected", command=send_selected).pack(side='right', padx=5)

# Initialize
load_users()
root.geometry("750x500")
root.mainloop()