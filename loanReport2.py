import tkinter as tk
from tkinter import filedialog
import pandas as pd
import numpy as np

try:
    root = tk.Tk()
    root.withdraw()

    file_path = filedialog.askopenfilename(
        parent=root,
        title="Select Excel File-Developed by Manish",
        filetypes=[("Excel files", "*.xls;*.xlsx")],
    )

    if file_path:
        df = pd.read_excel(file_path)
        print("Excel file loaded successfully.")
    else:
        print("No file selected.")

    df = df.iloc[1:, 2:]
    df.columns = df.iloc[0]
    df = df[1:]
    input("Press Enter to Continue: ")
    df.drop(columns=["Branch", "Purpose", "Disburse", "Overdue  Principal Amount", "Overdue  RI Amount", "Overdue  Other Int. Amount", "Address", "Obligor", "Collector"], inplace=True)
    df.rename(columns={
        "Disburse Date": "Disburse",
        "Int. Rate": "Int",
        "Mobile No.": "Contact",
        "Outstanding Principal Amount": "Principal",
        "Outstanding RI Amount": "RI",
        "Outstanding Other Int. Amount": "Fine",
        "Outstanding Total Amount": "Total",
        "Product Name": "Product"
    }, inplace=True)
    df = df.iloc[:-1]

    # Calculate 'Day' column and round to 2 decimal places
    day_column = df["RI"] / (df["Principal"] * df["Int"] / 36500)
    day_column = day_column.fillna(0)
    day_column.replace([np.inf, -np.inf], 0, inplace=True)
    df["Day"] = np.round(day_column, 2)    
    today = int(input("What Day Is Today? :"))
    last = int(input("Last Day Of Month? :"))
    df["Masanta"] = df["Day"] + last - today + 1
    df.sort_values(by="Masanta", ascending=False, inplace=True)
    df["Masanta RI"] = df["Masanta"] * (df["Principal"] * df["Int"] / 36500)
    df["Status"] = np.where(df["Masanta"] > 36, "Tageta", "")
    df["Product"] = df["Product"].str[-4:]

    # Calculations
    total_principal = df["Principal"].sum()
    total_principal_tageta = df[df["Status"] == "Tageta"]["Principal"].sum()
    total_customers = len(df)
    status_counts = df["Status"].value_counts().get("Tageta", 0)
    status_percentages = round((total_principal_tageta / total_principal) * 100, 2)
    
    # Create a summary DataFrame
    summary_df = pd.DataFrame({
        "Particulars": ["Total Principal Amount", "Total Due Amount", "Total Loanee", "Due Number", "Due Percentage"],
        "Data": [total_principal, total_principal_tageta, total_customers, status_counts, status_percentages]
    })

    save_path = filedialog.asksaveasfilename(
        parent=root,
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        title="Save Excel File-Developed By Manish"
    )

    # Check if a path was selected
    if save_path:
        # Save the DataFrames to the selected path with different sheet names
        with pd.ExcelWriter(save_path, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='loan_data', index=False)
            summary_df.to_excel(writer, sheet_name='dashboard', index=False)
            print(f"DataFrames saved successfully at: {save_path}")
    else:
        print("No file path selected.")

except Exception as e:
    print(f"An error occurred: {e}")

input("Task Completed Successfully")
