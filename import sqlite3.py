import sqlite3
import tkinter as tk
from tkinter import messagebox, ttk, filedialog
from tkcalendar import DateEntry
from openpyxl import Workbook
import csv

# Initialize the database
def init_db():
    conn = sqlite3.connect("attendance.db", timeout=10)
    cursor = conn.cursor()

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT NOT NULL UNIQUE,
        password TEXT NOT NULL
    )""")

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS students (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        roll_no TEXT NOT NULL UNIQUE,
        name TEXT NOT NULL
    )""")

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS employees (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        emp_id TEXT NOT NULL UNIQUE,
        name TEXT NOT NULL
    )""")

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS attendance (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        student_id INTEGER,
        employee_id INTEGER,
        date TEXT NOT NULL,
        status TEXT NOT NULL,
        user_type TEXT NOT NULL,
        FOREIGN KEY(student_id) REFERENCES students(id) ON DELETE CASCADE,
        FOREIGN KEY(employee_id) REFERENCES employees(id) ON DELETE CASCADE
    )""")

    conn.commit()
    conn.close()

# Auth Functions
def signup():
    username = su_username.get()
    password = su_password.get()
    if username and password:
        try:
            conn = sqlite3.connect("attendance.db")
            cursor = conn.cursor()
            cursor.execute("INSERT INTO users (username, password) VALUES (?, ?)", (username, password))
            conn.commit()
            messagebox.showinfo("Success", "Signup successful! Please login.")
            login_frame.tkraise()
        except sqlite3.IntegrityError:
            messagebox.showerror("Error", "Username already exists.")
        finally:
            conn.close()
    else:
        messagebox.showwarning("Input Error", "All fields are required!")

def login():
    username = username_entry.get().strip()
    password = password_entry.get().strip()
    conn = sqlite3.connect("attendance.db")
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM users WHERE username=? AND password=?", (username, password))
    user = cursor.fetchone()
    conn.close()
    if user:
        main_frame.tkraise()
    else:
        messagebox.showerror("Error", "Incorrect credentials!")

# Student Functions
def add_student():
    roll_no = roll_no_entry.get()
    name = name_entry.get()
    if roll_no and name:
        try:
            conn = sqlite3.connect("attendance.db")
            cursor = conn.cursor()
            cursor.execute("INSERT INTO students (roll_no, name) VALUES (?, ?)", (roll_no, name))
            conn.commit()
            messagebox.showinfo("Success", "Student added successfully")
            fetch_students()
            roll_no_entry.delete(0, tk.END)
            name_entry.delete(0, tk.END)
        except sqlite3.IntegrityError:
            messagebox.showerror("Error", "Roll number must be unique")
        finally:
            conn.close()
    else:
        messagebox.showwarning("Input Error", "All fields required!")

def remove_student():
    selected = student_tree.selection()
    if selected:
        student_id = student_tree.item(selected[0])["values"][0]
        conn = sqlite3.connect("attendance.db")
        cursor = conn.cursor()
        cursor.execute("DELETE FROM students WHERE id=?", (student_id,))
        conn.commit()
        conn.close()
        fetch_students()
        messagebox.showinfo("Deleted", "Student removed successfully")

def fetch_students():
    conn = sqlite3.connect("attendance.db")
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM students")
    records = cursor.fetchall()
    conn.close()
    student_tree.delete(*student_tree.get_children())
    attendance_tree.delete(*attendance_tree.get_children())
    for r in records:
        student_tree.insert("", "end", values=(r[0], r[1], r[2]))
        attendance_tree.insert("", "end", values=(r[0], r[1], r[2], ""))

# Attendance
status_options = ["Present", "Absent", "Late"]

def update_status(event):
    selected_item = attendance_tree.selection()[0]
    new_status = status_var.get()
    vals = attendance_tree.item(selected_item, "values")
    attendance_tree.item(selected_item, values=(vals[0], vals[1], vals[2], new_status))

def mark_attendance():
    date = date_entry.get()
    conn = sqlite3.connect("attendance.db")
    cursor = conn.cursor()
    for child in attendance_tree.get_children():
        student_id, roll_no, name, status = attendance_tree.item(child, "values")
        if status:
            cursor.execute("INSERT INTO attendance (student_id, employee_id, date, status, user_type) VALUES (?, NULL, ?, ?, 'student')", (student_id, date, status))
    conn.commit()
    conn.close()
    messagebox.showinfo("Success", "Attendance marked successfully")

def export_attendance():
    filename = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
    if filename:
        conn = sqlite3.connect("attendance.db")
        cursor = conn.cursor()
        cursor.execute("""
            SELECT students.roll_no, students.name, attendance.date, attendance.status
            FROM attendance
            JOIN students ON attendance.student_id = students.id
            WHERE attendance.user_type = 'student'
        """)
        rows = cursor.fetchall()
        conn.close()
        with open(filename, "w", newline="") as file:
            writer = csv.writer(file)
            writer.writerow(["Roll No", "Name", "Date", "Status"])
            writer.writerows(rows)
        messagebox.showinfo("Exported", "Attendance exported successfully")

def export_attendance_excel():
    filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if filename:
        conn = sqlite3.connect("attendance.db")
        cursor = conn.cursor()
        cursor.execute("""
            SELECT students.roll_no, students.name, attendance.date, attendance.status
            FROM attendance
            JOIN students ON attendance.student_id = students.id
            WHERE attendance.user_type = 'student'
        """)
        rows = cursor.fetchall()
        conn.close()

        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Attendance Report"
        sheet.append(["Roll No", "Name", "Date", "Status"])

        for row in rows:
            sheet.append(row)

        try:
            workbook.save(filename)
            messagebox.showinfo("Exported", "Attendance exported successfully to Excel")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save file: {e}")

# UI Setup
root = tk.Tk()
root.title("Attendance Management System")
root.geometry("900x650")
init_db()

# Login/Signup Frames
login_frame = tk.Frame(root)
signup_frame = tk.Frame(root)
main_frame = tk.Frame(root)

for frame in (login_frame, signup_frame, main_frame):
    frame.grid(row=0, column=0, sticky="nsew")

# --- Signup ---
tk.Label(signup_frame, text="Signup", font=("Arial", 20)).pack(pady=20)
su_username = tk.Entry(signup_frame, width=30)
su_password = tk.Entry(signup_frame, show="*", width=30)
tk.Label(signup_frame, text="Username").pack()
su_username.pack(pady=5)
tk.Label(signup_frame, text="Password").pack()
su_password.pack(pady=5)
tk.Button(signup_frame, text="Create Account", command=signup).pack(pady=10)
tk.Button(signup_frame, text="Back to Login", command=lambda: login_frame.tkraise()).pack()

# --- Login ---
tk.Label(login_frame, text="Login", font=("Arial", 20)).pack(pady=20)
username_entry = tk.Entry(login_frame, width=30)
password_entry = tk.Entry(login_frame, show="*", width=30)
tk.Label(login_frame, text="Username").pack()
username_entry.pack(pady=5)
tk.Label(login_frame, text="Password").pack()
password_entry.pack(pady=5)
tk.Button(login_frame, text="Login", command=login).pack(pady=10)
tk.Button(login_frame, text="Signup", command=lambda: signup_frame.tkraise()).pack()

# --- Main Tabs ---
notebook = ttk.Notebook(main_frame)
notebook.pack(pady=10, expand=True, fill="both")

# Add Student Tab
student_tab = tk.Frame(notebook)
notebook.add(student_tab, text="Manage Students")

tk.Label(student_tab, text="Roll No:").pack()
roll_no_entry = tk.Entry(student_tab)
roll_no_entry.pack(pady=2)
tk.Label(student_tab, text="Name:").pack()
name_entry = tk.Entry(student_tab)
name_entry.pack(pady=2)
tk.Button(student_tab, text="Add Student", command=add_student).pack(pady=5)
tk.Button(student_tab, text="Remove Selected", command=remove_student).pack(pady=5)

student_tree = ttk.Treeview(student_tab, columns=("ID", "Roll No", "Name"), show="headings")
student_tree.heading("ID", text="ID")
student_tree.heading("Roll No", text="Roll No")
student_tree.heading("Name", text="Name")
student_tree.pack(pady=10, fill="both", expand=True)

# Attendance Tab
attendance_tab = tk.Frame(notebook)
notebook.add(attendance_tab, text="Mark Attendance")

tk.Label(attendance_tab, text="Date:").pack()
date_entry = DateEntry(attendance_tab)
date_entry.pack(pady=5)

tk.Button(attendance_tab, text="Mark Attendance", command=mark_attendance).pack(pady=5)
tk.Button(attendance_tab, text="Export Attendance", command=export_attendance).pack(pady=5)
tk.Button(attendance_tab, text="Export to Excel", command=export_attendance_excel).pack(pady=5)

status_var = tk.StringVar()
tk.Label(attendance_tab, text="Set Status:").pack()
status_dropdown = ttk.Combobox(attendance_tab, textvariable=status_var, values=status_options, state="readonly")
status_dropdown.pack(pady=5)
status_dropdown.bind("<<ComboboxSelected>>", update_status)

attendance_tree = ttk.Treeview(attendance_tab, columns=("ID", "Roll No", "Name", "Status"), show="headings")
attendance_tree.heading("ID", text="ID")
attendance_tree.heading("Roll No", text="Roll No")
attendance_tree.heading("Name", text="Name")
attendance_tree.heading("Status", text="Status")
attendance_tree.pack(pady=10, fill="both", expand=True)

# Start at login
login_frame.tkraise()
fetch_students()
root.mainloop()
