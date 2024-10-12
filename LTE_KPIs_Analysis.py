import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import folium
from folium.plugins import HeatMap
from geopy.distance import geodesic
import webbrowser
import os
import numpy as np
import ttkbootstrap as ttk
from tkinter import filedialog, messagebox, Listbox, MULTIPLE, END
from ttkbootstrap.constants import *

# Function to load and clean a single Excel file
def load_and_clean_file(file_path):
    print(f"Loading file from: {file_path}")
    data = pd.read_excel(file_path)
    
    # Check if data loaded successfully
    if data.empty:
        print("The file is empty.")
        return None
    
    # Fill missing numeric data with column mean (only for numeric columns)
    numeric_columns = data.select_dtypes(include=[np.number])
    data[numeric_columns.columns] = numeric_columns.fillna(numeric_columns.mean())
    
    return data

# Function to plot time series for two files and selected KPIs
def plot_time_series(data1, data2, kpis):
    time1 = data1['Time']
    time2 = data2['Time']
    
    plt.figure(figsize=(12, 8))
    
    # Plot for First Log File
    for kpi in kpis:
        if kpi in data1.columns:
            plt.plot(time1, data1[kpi], label=f"{kpi} - First Log File", linestyle='-', marker='o')

    # Plot for Second Log File
    for kpi in kpis:
        if kpi in data2.columns:
            plt.plot(time2, data2[kpi], label=f"{kpi} - Second Log File", linestyle='--', marker='x')
    
    plt.xlabel("Time")
    plt.ylabel("KPI Value")
    plt.legend(loc="best")
    plt.title("KPI Trends Over Time for Two Files (Selected KPIs)")
    plt.grid(True)
    plt.tight_layout()  # Adjust layout to avoid clipping
    plt.show()

# Function to create heatmap of signal strength
def plot_heatmap(data):
    map_center = [data['Latitude'].mean(), data['Longitude'].mean()]
    m = folium.Map(location=map_center, zoom_start=13)
    
    heat_data = [[row['Latitude'], row['Longitude'], row['Serving Cell RSRP (dBm)']] for index, row in data.iterrows()]
    HeatMap(heat_data).add_to(m)
    
    # Save and open in browser
    m.save("heatmap.html")
    webbrowser.open("heatmap.html")

# GUI functions
def browse_file1():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        file1_var.set(file_path)
        update_generate_button_state()

def browse_file2():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        file2_var.set(file_path)
        update_generate_button_state()

def update_generate_button_state():
    if file1_var.get() and file2_var.get():
        generate_button.config(state=NORMAL)
    else:
        generate_button.config(state=DISABLED)

def generate_analysis():
    file1_path = file1_var.get()
    file2_path = file2_var.get()
    
    if not file1_path or not file2_path:
        messagebox.showerror("Error", "Please select both files.")
        return
    
    # Load and clean data from both files
    data1 = load_and_clean_file(file1_path)
    data2 = load_and_clean_file(file2_path)
    
    if data1 is None or data2 is None:
        messagebox.showerror("Error", "Error loading the files. Please check the file contents.")
        return
    
    # Get selected KPIs for comparison
    selected_kpis = []
    if kpi_var.get() == 'Single KPI':
        selected_kpis = [single_kpi_var.get()]
    else:  # Multiple KPIs
        selected_indices = kpi_listbox.curselection()
        selected_kpis = [kpi_listbox.get(i) for i in selected_indices]
    
    # Validate KPI selection
    if not selected_kpis:
        messagebox.showerror("Error", "Please select at least one KPI for visualization.")
        return
    
    # Plot Time Series for selected KPIs
    plot_time_series(data1, data2, selected_kpis)
    
    # Heatmap for file1
    plot_heatmap(data1)
    
    messagebox.showinfo("Success", "Analysis Complete!")

# Tkinter window setup
root = ttk.Window(themename="flatly")
root.title("LTE KPI Analysis Tool")
root.geometry("800x800")

# Create a frame to organize file selection components
file_frame = ttk.Frame(root)
file_frame.pack(pady=10, fill='x')

# Add the main label at the top of the window
creator_label = ttk.Label(file_frame, text="LTE KPI Analysis Tool", font=("Arial", 14, "bold"))
creator_label.pack(padx=5, pady=5)

# Add the sub-label below the main label
sub_creator_label = ttk.Label(file_frame, text="Created by Salah Bendary", font=("Arial", 10))
sub_creator_label.pack(padx=5, pady=0)  

file1_var = ttk.StringVar()
file2_var = ttk.StringVar()

# First Log File selection
ttk.Label(file_frame, text="Select First Log File (Excel Format)", font=("Arial", 12)).pack(pady=5)
ttk.Entry(file_frame, textvariable=file1_var, width=50).pack(pady=2)
ttk.Button(file_frame, text="Browse...", command=browse_file1, width=10).pack(pady=2)

# Second Log File selection
ttk.Label(file_frame, text="Select Second Log File (Excel Format)", font=("Arial", 12)).pack(pady=5)
ttk.Entry(file_frame, textvariable=file2_var, width=50).pack(pady=2)
ttk.Button(file_frame, text="Browse...", command=browse_file2, width=10).pack(pady=2)

# KPI categories and mappings
kpi_categories = {
    'Signal Quality': ['Serving Cell RSRP (dBm)', 'Serving Cell RSRQ (dB)', 'RS SINR Carrier 1 (dB)'],
    'Throughput': ['RLC Throughput DL (kbps)', 'PDSCH Phy Throughput (kbps)', 'PUSCH Phy Throughput (kbps)'],
    'Interference': ['Serving Cell Channel RSSI (dBm)', 'PDSCH BLER Carrier 1 (%)'],
    'Neighbor Cells': ['Neighbor Cell RSRP (dBm): N1', 'Neighbor Cell RSRQ (dB): N1'],
}

# KPI selection
kpi_frame = ttk.Frame(root)
kpi_frame.pack(pady=15, fill='x')

ttk.Label(kpi_frame, text="Select KPI Visualization", font=("Arial", 14)).pack(pady=5)

# Single or Multiple KPI Option
kpi_var = ttk.StringVar(value="Single KPI")
ttk.Radiobutton(kpi_frame, text="Single KPI", variable=kpi_var, value="Single KPI", bootstyle="info").pack()
ttk.Radiobutton(kpi_frame, text="Multiple KPIs", variable=kpi_var, value="Multiple KPIs", bootstyle="info").pack()

# Dropdown for Single KPI
single_kpi_var = ttk.StringVar()
single_kpi_dropdown = ttk.Combobox(kpi_frame, textvariable=single_kpi_var, values=[kpi for cat in kpi_categories.values() for kpi in cat], bootstyle="primary")
single_kpi_dropdown.pack(pady=5)

# Category selection dropdown
ttk.Label(kpi_frame, text="Select KPI Category", font=("Arial", 12)).pack(pady=5)
category_var = ttk.StringVar()
category_dropdown = ttk.Combobox(kpi_frame, textvariable=category_var, values=list(kpi_categories.keys()), state="readonly", bootstyle="primary")
category_dropdown.pack(pady=5)

# Listbox for multiple KPIs
kpi_listbox = Listbox(kpi_frame, selectmode=MULTIPLE, height=10, width=50)
kpi_listbox.pack(pady=5)

# Update KPI listbox based on category selection
def update_kpi_listbox(category):
    kpi_listbox.delete(0, END)  # Clear current KPI options
    for kpi in kpi_categories[category]:
        kpi_listbox.insert(END, kpi)  # Insert new KPIs based on category

category_dropdown.bind('<<ComboboxSelected>>', lambda event: update_kpi_listbox(category_var.get()))

# Generate analysis button in the right-center side
button_frame = ttk.Frame(root)
button_frame.pack(side="right", padx=1, pady=1)

generate_button = ttk.Button(button_frame, text="Generate Analysis", command=generate_analysis, state=DISABLED, width=30)
generate_button.pack(side="right", padx=10)

root.mainloop()
