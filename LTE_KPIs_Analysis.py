import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import tkinter as tk
from tkinter import filedialog, messagebox, Listbox, Scrollbar
from tkinter import simpledialog
from tkinter import Label, Button, Frame
from PIL import Image, ImageTk

# Function to load and clean a single Excel file
def load_and_clean_file(file_path):
    data = pd.read_excel(file_path)
    
    # Check if data loaded successfully
    if data.empty:
        messagebox.showerror("Error", "The file is empty.")
        return None
    
    # Fill missing numeric data with column mean (only for numeric columns)
    numeric_columns = data.select_dtypes(include=[np.number])
    data[numeric_columns.columns] = numeric_columns.fillna(numeric_columns.mean())
    
    return data

# Function to plot time series for multiple KPIs with subplots
def plot_time_series_with_subplots(data1, data2, kpis):
    time1 = data1['Time']
    time2 = data2['Time']
    
    num_kpis = len(kpis)
    fig, axes = plt.subplots(nrows=num_kpis, ncols=1, figsize=(12, num_kpis * 4))
    
    # If there's only one KPI, axes won't be an array. Make it a list for consistency.
    if num_kpis == 1:
        axes = [axes]
    
    # Plot each KPI in its own subplot
    for i, kpi in enumerate(kpis):
        ax = axes[i]
        
        if kpi in data1.columns:
            ax.plot(time1, data1[kpi], label=f"{kpi} - First Log File", linestyle='-', marker='o')
        if kpi in data2.columns:
            ax.plot(time2, data2[kpi], label=f"{kpi} - Second Log File", linestyle='--', marker='x')
        
        ax.set_xlabel("Time")
        ax.set_ylabel(f"{kpi} Value")
        ax.set_title(f"Comparison of {kpi} Over Time")
        ax.grid(True)
        ax.legend(loc="best")
    
    plt.tight_layout()  # Adjust layout to avoid overlap
    plt.show()

# Function to handle file selection and KPI selection
def select_files_and_kpis(file_entry):
    # Ask the user to select the file
    file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel Files", "*.xlsx")])
    if file_path:
        file_entry.delete(0, tk.END)  # Clear the entry field
        file_entry.insert(0, file_path)  # Insert the file path into the entry

# Function to handle the file selection and KPI comparison
def load_and_compare_files(file1_entry, file2_entry):
    file1_path = file1_entry.get()
    file2_path = file2_entry.get()

    if not file1_path or not file2_path:
        messagebox.showerror("Error", "Please select both files.")
        return

    # Load and clean the data from both files
    data1 = load_and_clean_file(file1_path)
    data2 = load_and_clean_file(file2_path)

    if data1 is None or data2 is None:
        return
    
    # Automatically retrieve the common KPI columns from both files (excluding the 'Time' column)
    common_columns = list(set(data1.columns).intersection(set(data2.columns)))
    if 'Time' in common_columns:
        common_columns.remove('Time')  # Exclude 'Time' from the selectable KPIs

    # Ask the user how many KPIs they want to compare
    num_kpis = simpledialog.askinteger("Number of KPIs", "Enter the number of KPIs to compare (1 or more):")
    
    if num_kpis is None or num_kpis < 1:
        messagebox.showerror("Error", "Please enter a valid number of KPIs.")
        return
    
    selected_kpis = []
    
    # Create a listbox for the user to select KPIs
    kpi_window = tk.Toplevel()
    kpi_window.title("Select KPIs")
    kpi_window.geometry("300x400")
    
    listbox = Listbox(kpi_window, selectmode=tk.MULTIPLE)
    scrollbar = Scrollbar(kpi_window)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    listbox.config(yscrollcommand=scrollbar.set)
    scrollbar.config(command=listbox.yview)
    
    for kpi in common_columns:
        listbox.insert(tk.END, kpi)
    
    listbox.pack(expand=True, fill=tk.BOTH)
    
    def on_select_kpis():
        selected_indices = listbox.curselection()
        if len(selected_indices) != num_kpis:
            messagebox.showerror("Error", f"Please select exactly {num_kpis} KPIs.")
            return
        for idx in selected_indices:
            selected_kpis.append(common_columns[idx])
        kpi_window.destroy()
        # Plot the selected KPIs
        plot_time_series_with_subplots(data1, data2, selected_kpis)

    # Add a button to confirm the selection
    btn_confirm = tk.Button(kpi_window, text="Confirm Selection", command=on_select_kpis)
    btn_confirm.pack()

# Main GUI window
def main_gui():
    root = tk.Tk()
    root.title("LTE KPIs Analysis Tool")
    root.geometry("500x300")

    # Add the logo image
    logo_img = Image.open("vodafone_logo.png")  
    logo_img = logo_img.resize((50, 50), Image.Resampling.LANCZOS)
    logo_photo = ImageTk.PhotoImage(logo_img)
    logo_label = Label(root, image=logo_photo)
    logo_label.place(x=10)

    # Add the title and subtitle
    label_title = tk.Label(root, text="LTE KPIs Analysis Tool", font=("Arial", 16, "bold"))
    label_subtitle = tk.Label(root, text="Created by Salah Bendary", font=("Arial", 10))

    label_title.grid(row=0, column=0, columnspan=2, pady=10)
    label_subtitle.grid(row=1, column=0, columnspan=2, pady=5)

    # Labels and textboxes for file uploads
    label_file1 = tk.Label(root, text="Upload First Excel File:", font=("Arial", 10))
    label_file2 = tk.Label(root, text="Upload Second Excel File:", font=("Arial", 10))
    
    label_file1.grid(row=2, column=0, padx=10, pady=5, sticky="w")
    
    file1_entry = tk.Entry(root, width=30)
    file1_entry.grid(row=2, column=1, padx=10, pady=5)
    
    btn_browse1 = tk.Button(root, text="Browse...", command=lambda: select_files_and_kpis(file1_entry))
    btn_browse1.grid(row=2, column=2, padx=10, pady=5)

    label_file2.grid(row=3, column=0, padx=10, pady=5, sticky="w")

    file2_entry = tk.Entry(root, width=30)
    file2_entry.grid(row=3, column=1, padx=10, pady=5)
    
    btn_browse2 = tk.Button(root, text="Browse...", command=lambda: select_files_and_kpis(file2_entry))
    btn_browse2.grid(row=3, column=2, padx=10, pady=5)

    # Button to trigger the file selection and analysis
    btn_analyze = tk.Button(root, text="Select Files and KPIs", command=lambda: load_and_compare_files(file1_entry, file2_entry))
    btn_analyze.grid(row=4, column=0, columnspan=3, pady=20)

    root.mainloop()

if __name__ == "__main__":
    main_gui()
