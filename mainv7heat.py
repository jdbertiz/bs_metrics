import os
import csv
from openpyxl import load_workbook
from collections import defaultdict
import matplotlib.pyplot as plt
import pandas as pd
from matplotlib.backends.backend_pdf import PdfPages
from datetime import datetime
import seaborn as sns

# Path to the BSREPORTS folder
folder_path = 'BSREPORTS'

# Dictionary to store aggregated results
aggregated_data = defaultdict(lambda: [0, 0])  # key: (Content, Type), value: [Unique Viewers sum, Viewers sum]

# Loop through all files in the folder
for filename in os.listdir(folder_path):
    if filename.endswith('.xlsx'):
        file_path = os.path.join(folder_path, filename)
        print(f"\nProcessing file: {file_path}")
        
        workbook = load_workbook(filename=file_path)
        sheet_names = workbook.sheetnames
        print("Sheets in the workbook:", sheet_names)
        
        for sheet_name in sheet_names:
            if sheet_name.lower() == "popular content":
                sheet = workbook[sheet_name]
                print(f"\nSheet: {sheet_name}")
                
                # Start reading from row 7
                for row in sheet.iter_rows(min_row=7, values_only=True):
                    if all(cell is None for cell in row):
                        break
                    
                    cleaned_row = [cell for cell in row if cell is not None]
                    
                    if len(cleaned_row) >= 4:
                        content = cleaned_row[0]
                        page_type = cleaned_row[1]
                        try:
                            unique_viewers = int(cleaned_row[2])
                            viewers = int(cleaned_row[3])
                        except ValueError:
                            continue  # Skip if values are not numbers
                        
                        key = (content, page_type)
                        aggregated_data[key][0] += unique_viewers
                        aggregated_data[key][1] += viewers

# Prepare data for display and chart
data = []
type_totals = defaultdict(int)

for (content, page_type), (unique_viewers, viewers) in aggregated_data.items():
    data.append([content, page_type, unique_viewers, viewers])
    type_totals[page_type] += viewers

# Convert to DataFrame for clean table output
df = pd.DataFrame(data, columns=['Content', 'Type', 'Unique Viewers', 'Viewers'])

# Plot pie chart for Viewers by Type
labels = list(type_totals.keys())
sizes = list(type_totals.values())

plt.figure(figsize=(8, 6))
plt.pie(sizes, labels=labels, autopct='%1.1f%%', startangle=140)
plt.title('Viewers by Popular Type')
plt.axis('equal')  # Equal aspect ratio makes the pie chart circular
plt.tight_layout()

pdf_path = "report_viewers_summary.pdf"

# Function to parse date from filename
def extract_date_from_filename(filename):
    try:
        return datetime.strptime(filename.split('_')[-1].replace('.xlsx', ''), "%d-%b,%Y")
    except Exception:
        return None

# Identify the latest file by date in filename
latest_file = None
latest_date = None
for filename in os.listdir(folder_path):
    if filename.startswith('SiteAnalyticsData_') and filename.endswith('.xlsx'):
        file_date = extract_date_from_filename(filename)
        if file_date and (latest_date is None or file_date > latest_date):
            latest_date = file_date
            latest_file = filename

# Read data from the latest Usage by device sheet
device_usage_data = []

if latest_file:
    file_path = os.path.join(folder_path, latest_file)
    workbook = load_workbook(filename=file_path, data_only=True)
    print(f"\nProcessing 'Usage by device' from latest file: {latest_file}")

    if "Usage by device" in workbook.sheetnames:
        sheet = workbook["Usage by device"]

        # Collect the data and limit it to the last 30 entries
        rows = list(sheet.iter_rows(min_row=2, values_only=True))  # Assuming row 1 is headers
        last_30_rows = rows[-30:]  # Get the last 30 rows
        
        for row in last_30_rows:
            if row[0] is None:
                continue
            try:
                date = pd.to_datetime(row[0])
                visits = sum(cell if isinstance(cell, (int, float)) else 0 for cell in row[1:])  # sum visits from all devices
                device_usage_data.append([date, visits])
            except Exception:
                continue

# Create DataFrame and Heatmap if data exists
if device_usage_data:
    df_device = pd.DataFrame(device_usage_data, columns=["Date", "Total Visits"])
    df_device.sort_values("Date", inplace=True)
    df_device.set_index("Date", inplace=True)

    # Create pivot for heatmap: Assuming we want a week-wise heatmap
    df_device['Week'] = df_device.index.to_series().dt.to_period('W').astype(str)
    df_device['Day'] = df_device.index.to_series().dt.day_name()
    pivot = df_device.pivot_table(index='Day', columns='Week', values='Total Visits', aggfunc='sum')
    pivot = pivot.reindex(['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'])

    # Plot heatmap
    fig3, ax3 = plt.subplots(figsize=(12, 6))
    sns.heatmap(pivot, cmap="YlOrRd", linewidths=.5, annot=True, fmt=".0f", ax=ax3)
    ax3.set_title('Visits Heatmap by Weekday', fontsize=14)

# Now begin saving to PDF
with PdfPages(pdf_path) as pdf:
    # --- Page 1: Pie Chart ---
    fig1, ax1 = plt.subplots(figsize=(8, 6))
    ax1.pie(sizes, labels=labels, autopct='%1.1f%%', startangle=140)
    ax1.set_title('Viewers by Popular Type')
    ax1.axis('equal')
    pdf.savefig(fig1)
    plt.close(fig1)

    # --- Page 2: Table (plotted using matplotlib) ---
    fig2, ax2 = plt.subplots(figsize=(10, len(df) * 0.4 + 1))  # Adjust height to fit rows
    ax2.axis('off')
    ax2.set_title("Aggregated Content View Summary", fontsize=14, fontweight='bold')
    
    table = ax2.table(
        cellText=df.values,
        colLabels=df.columns,
        loc='center',
        cellLoc='left',
        colColours=['#CCCCCC'] * len(df.columns)
    )

    # Adjust font and row height
    table.auto_set_font_size(False)
    table.set_fontsize(9)
    table.scale(1.5, 1.5)  # Widen columns and increase row height
    table.auto_set_column_width(col=list(range(len(df.columns))))

    pdf.savefig(fig2)
    plt.close(fig2)

    # --- Page 3: Heatmap ---
    if device_usage_data:  # Only add the heatmap if we have data
        pdf.savefig(fig3)
        plt.close(fig3)

    # --- Page 4: Visits by Date Table ---
    if device_usage_data:
        fig4, ax4 = plt.subplots(figsize=(10, len(df_device) * 0.4 + 1))
        ax4.axis('off')
        ax4.set_title("Visits by Date", fontsize=14, fontweight='bold')

        df_table = df_device.reset_index()
        colLabels = df_table.columns

        table = ax4.table(
            cellText=df_table.values,
            colLabels=colLabels,
            loc='center',
            cellLoc='left',
            colColours=['#CCCCCC'] * len(colLabels)
        )

        # Adjust font size
        table.auto_set_font_size(False)
        table.set_fontsize(9)
        table.scale(1.5, 1.5)

    # Adjust column widths dynamically
    for i, col in enumerate(df_table.columns):
        max_length = max(df_table[col].apply(lambda x: len(str(x))))
        table.auto_set_column_width([i])  # Automatically adjust column width

    # Alternatively, manually set the width for the date column
    date_column_index = df_table.columns.get_loc('Date')  # Adjust this if needed
    table.auto_set_column_width([date_column_index])  # Automatically adjust date column width

    # Save to PDF
    pdf.savefig(fig4)
    plt.close(fig4)


print(f"\nPDF report saved as '{pdf_path}'")
