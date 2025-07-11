import tkinter as tk
from tkinter import messagebox, simpledialog
from PIL import Image, ImageTk
import matplotlib.pyplot as plt
from colorama import init, Fore
import csv
import os

from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage

# ====== Init Colorama ======
init(autoreset=True)

# ====== Globals ======
CSV_FILE = "transactions.csv"
EXCEL_FILE = "transactions.xlsx"
CHART_FILE = "chart.png"
transactions = []

# ====== Load CSV Transactions ======
def load_transactions():
    if not os.path.exists(CSV_FILE):
        return
    with open(CSV_FILE, mode='r', newline='', encoding='utf-8') as file:
        reader = csv.DictReader(file)
        for row in reader:
            try:
                row['amount'] = float(row['amount'])
                transactions.append(row)
            except:
                continue

# ====== Save CSV ======
def save_csv():
    with open(CSV_FILE, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.DictWriter(file, fieldnames=["type", "description", "amount", "date"])
        writer.writeheader()
        for t in transactions:
            writer.writerow(t)

# ====== Save Excel with Chart ======
def save_excel_with_chart():
    wb = Workbook()
    ws = wb.active
    ws.title = "Transactions"

    # Add headers
    ws.append(["Type", "Description", "Amount", "Date"])
    for t in transactions:
        ws.append([t["type"], t["description"], t["amount"], t["date"]])

    # Save chart image
    if os.path.exists(CHART_FILE):
        img = ExcelImage(CHART_FILE)
        img.anchor = "F2"
        ws.add_image(img)

    wb.save(EXCEL_FILE)
    print(Fore.GREEN + f"📊 Excel saved as {EXCEL_FILE} including chart.")

# ====== Add Transaction ======
def add_transaction(t_type):
    description = simpledialog.askstring("Input", f"Enter {t_type} description:")
    if not description:
        return
    try:
        amount = float(simpledialog.askstring("Input", f"Enter {t_type} amount:"))
    except:
        messagebox.showerror("Error", "Amount must be a number.")
        return
    date = simpledialog.askstring("Input", "Enter date (e.g. 2025-07-11):")
    if not date:
        return
    transactions.append({
        "type": t_type,
        "description": description,
        "amount": amount,
        "date": date
    })
    save_csv()
    messagebox.showinfo("Success", f"{t_type.capitalize()} added successfully!")

# ====== Totals ======
def calculate_totals():
    income = sum(t['amount'] for t in transactions if t['type'] == 'income')
    expense = sum(t['amount'] for t in transactions if t['type'] == 'expense')
    balance = income - expense
    count = len(transactions)
    return income, expense, balance, count

# ====== Max Expense ======
def show_max_expense():
    max_amount = 0
    max_expense = None
    for t in transactions:
        if t['type'] == 'expense' and t['amount'] > max_amount:
            max_amount = t['amount']
            max_expense = t
    if max_expense:
        msg = (
            f"Highest Expense:\n"
            f"Description: {max_expense['description']}\n"
            f"Amount: ${max_expense['amount']:.2f}\n"
            f"Date: {max_expense['date']}"
        )
        messagebox.showinfo("Max Expense", msg)
    else:
        messagebox.showinfo("Max Expense", "No expenses found.")

# ====== Show Report ======
def show_report():
    income, expense, balance, count = calculate_totals()
    msg = (
        f"Total Income: ${income:.2f}\n"
        f"Total Expenses: ${expense:.2f}\n"
        f"Balance: ${balance:.2f}\n"
        f"Total Transactions: {count}"
    )
    messagebox.showinfo("Financial Report", msg)

# ====== Charts ======
def show_charts():
    income = sum(t['amount'] for t in transactions if t['type'] == 'income')
    expense = sum(t['amount'] for t in transactions if t['type'] == 'expense')

    if income == 0 and expense == 0:
        messagebox.showinfo("No Data", "No income or expense to show.")
        return

    plt.figure(figsize=(10, 4))

    # Pie
    plt.subplot(1, 2, 1)
    plt.pie([income, expense], labels=["Income", "Expense"], colors=["#4CAF50", "#F44336"], autopct="%1.1f%%")
    plt.title("Income vs Expense")

    # Bar
    bar_data = {}
    for t in transactions:
        key = t["description"]
        bar_data[key] = bar_data.get(key, 0) + t["amount"]
    sorted_items = sorted(bar_data.items(), key=lambda x: x[1], reverse=True)[:5]
    keys = [item[0] for item in sorted_items]
    values = [item[1] for item in sorted_items]

    plt.subplot(1, 2, 2)
    plt.bar(keys, values, color="#2196F3")
    plt.xticks(rotation=45)
    plt.title("Top 5 Items")

    plt.tight_layout()
    plt.savefig(CHART_FILE)
    save_excel_with_chart()
    plt.show()

# ====== GUI Setup ======
def create_gui():
    root = tk.Tk()
    root.title("MoneyFlow Tracker")
    root.geometry("400x440")

    try:
        bg = Image.open("money.jpg")
        bg = bg.resize((400, 440), Image.Resampling.LANCZOS)
        bg_photo = ImageTk.PhotoImage(bg)
        bg_label = tk.Label(root, image=bg_photo)
        bg_label.place(x=0, y=0, relwidth=1, relheight=1)
    except:
        root.config(bg="#eeeeee")

    tk.Label(root, text="MoneyFlow Tracker", font=("Arial", 16, "bold"),
             bg="#333333", fg="white").pack(pady=15)

    tk.Button(root, text="Add Income", width=25, bg="#4CAF50", fg="white",
              command=lambda: add_transaction("income")).pack(pady=5)

    tk.Button(root, text="Add Expense", width=25, bg="#F44336", fg="white",
              command=lambda: add_transaction("expense")).pack(pady=5)

    tk.Button(root, text="Show Financial Report", width=25, bg="#FFC107", fg="black",
              command=show_report).pack(pady=5)

    tk.Button(root, text="Analyze with Chart", width=25, bg="#2196F3", fg="white",
              command=show_charts).pack(pady=5)

    tk.Button(root, text="Show Max Expense", width=25, bg="#FF5722", fg="white",
              command=show_max_expense).pack(pady=5)

    tk.Button(root, text="Exit", width=25, bg="#9C27B0", fg="white",
              command=root.quit).pack(pady=15)

    root.mainloop()

# ====== Start ======
if __name__ == "__main__":
    load_transactions()
    create_gui()
