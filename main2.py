import tkinter as tk
from tkinter import messagebox
import openpyxl
from openpyxl import Workbook

# Function to save data to Excel
def save_to_excel():
    name = name_entry.get()
    age = age_entry.get()
    email = email_entry.get()

    if not name or not age or not email:
        messagebox.showerror("Input Error", "All fields are required!")
        return

    try:
        age = int(age)
    except ValueError:
        messagebox.showerror("Input Error", "Age must be an integer!")
        return

    # Create a new workbook and select the active worksheet
    wb = Workbook()
    ws = wb.active

    # Add headers
    ws.append(["Name", "Age", "Email"])

    # Append data
    ws.append([name, age, email])

    # Save the file
    filename = "user_data.xlsx"
    wb.save(filename)

    messagebox.showinfo("Success", f"Data saved to {filename}")

    # Clear the fields
    name_entry.delete(0, tk.END)
    age_entry.delete(0, tk.END)
    email_entry.delete(0, tk.END)

# Set up the main application window
root = tk.Tk()
root.title("Data Entry Form")

# Create and place labels and entry widgets for the form
tk.Label(root, text="Name").grid(row=0, column=0, padx=10, pady=10)
name_entry = tk.Entry(root)
name_entry.grid(row=0, column=1, padx=10, pady=10)

tk.Label(root, text="Age").grid(row=1, column=0, padx=10, pady=10)
age_entry = tk.Entry(root)
age_entry.grid(row=1, column=1, padx=10, pady=10)

tk.Label(root, text="Email").grid(row=2, column=0, padx=10, pady=10)
email_entry = tk.Entry(root)
email_entry.grid(row=2, column=1, padx=10, pady=10)

# Create and place the submit button
submit_button = tk.Button(root, text="Save to Excel", command=save_to_excel)
submit_button.grid(row=3, columnspan=2, pady=20)

# Run the application
root.mainloop()
