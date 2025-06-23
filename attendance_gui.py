import tkinter as tk
from tkinter import messagebox
from openpyxl import Workbook

attendance = {}

def mark_attendance():
    name = entry_name.get().strip()
    if not name:
        messagebox.showwarning("Input Error", "Please enter a student name.")
        return
    if name in attendance:
        messagebox.showinfo("Info", f"{name} is already marked present.")
    else:
        attendance[name] = "Present"
        messagebox.showinfo("Success", f"Attendance marked for {name}.")
    entry_name.delete(0, tk.END)

def view_attendance():
    if not attendance:
        messagebox.showinfo("Attendance", "No records found.")
    else:
        records = "\n".join(f"{name}: {status}" for name, status in attendance.items())
        messagebox.showinfo("Attendance Records", records)

def save_and_exit():
    # Save to text file
    with open("attendance_gui.txt", "w") as file:
        for name, status in attendance.items():
            file.write(f"{name}: {status}\n")

    # Save to Excel file
    wb = Workbook()
    ws = wb.active
    ws.title = "Attendance"
    ws.append(["Name", "Status"])
    for name, status in attendance.items():
        ws.append([name, status])
    wb.save("attendance_gui.xlsx")

    root.destroy()

root = tk.Tk()
root.title("Student Attendance System")
root.geometry("400x300")
root.configure(bg="#f2f2f2")

tk.Label(root, text="Enter Student Name:", font=("Arial", 12), bg="#f2f2f2").pack(pady=10)
entry_name = tk.Entry(root, font=("Arial", 12), width=30)
entry_name.pack(pady=5)

tk.Button(root, text="Mark Attendance", font=("Arial", 11), command=mark_attendance, width=25, bg="#4CAF50", fg="white").pack(pady=10)
tk.Button(root, text="View Attendance", font=("Arial", 11), command=view_attendance, width=25, bg="#2196F3", fg="white").pack(pady=5)
tk.Button(root, text="Save & Exit", font=("Arial", 11), command=save_and_exit, width=25, bg="#f44336", fg="white").pack(pady=15)

root.mainloop()
