# Automating Brand Space Metrics

This Python script processes data from multiple Excel files (`.xlsx`) within a folder and generates a detailed PDF report. The report contains various visualizations and tables, including pie charts, heatmaps, and aggregated tables. It utilizes multiple libraries, including `openpyxl`, `matplotlib`, `pandas`, `seaborn`, and `PdfPages`.

## Prerequisites

Before running the script, ensure you have the following libraries installed:
pip install openpyxl matplotlib pandas seaborn

## Functionality Overview

### 1. **Data Aggregation**
   - The script reads all `.xlsx` files within the `BSREPORTS` folder.
   - It searches for a sheet named "Popular Content" in each workbook and aggregates data on content and page type, summing up unique viewers and total viewers.
   - Aggregated results are stored in a dictionary, which is later converted to a pandas DataFrame for clean output.

### 2. **Visualizations**
   The script generates the following charts and saves them into a PDF:

   - **Page 1**: A pie chart showing the distribution of viewers by type.
   - **Page 2**: A table summarizing the aggregated content view data.
   - **Page 3**: A heatmap of visits by date, generated from the "Usage by device" sheet of the latest file.
   - **Page 4**: A table summarizing visits by date, showing the total visits for each date.
   - **Page 5**: Heatmaps of visits by device (e.g., Desktop, Mobile, Tablet).
   - **Page 6**: A pie chart showing the distribution of views over different timeframes (from the "Usage by time" sheet).
   
### 3. **Processing the Latest Data**
   The script identifies the latest file (based on the date in the filename) and processes the "Usage by device" sheet, extracting data for generating heatmaps and total visits by date.

### 4. **PDF Report Generation**
   The script compiles all charts and tables into a PDF, which is saved with the name `report_viewers_summary.pdf`.

## Usage

To use this script:

1. Place your Excel files in a folder named `BSREPORTS`.
2. Run the script. The PDF report will be generated and saved in the current working directory.

`python your_script_name.py`

## Script Workflow

1. **File Processing**:
   - The script processes all files in the `BSREPORTS` folder.
   - For each `.xlsx` file, it looks for the sheet "Popular Content" and aggregates data on content and type.
   
2. **Data Aggregation**:
   - For each valid row in the "Popular Content" sheet, it extracts content, page type, unique viewers, and total viewers.
   - This data is aggregated and stored in a dictionary.
   
3. **Data Visualization**:
   - After the data is aggregated, the script creates:
     - A pie chart for viewer distribution by type.
     - A table summarizing the aggregated data.
     - Heatmaps for visits by date and device.
     - A pie chart for views by timeframe.
   
4. **Report Generation**:
   - The generated charts and tables are saved in a PDF document (`report_viewers_summary.pdf`).

## Example Output

- **Pie Chart** for "Viewers by Popular Type":
  
  ![Pie Chart Example](<img width="361" alt="image" src="https://github.com/user-attachments/assets/d1dcd5ed-196f-47a0-8190-2435e9029f54" />)

- **Aggregated Table** showing content view summary:

  | Content    | Type     | Unique Viewers | Viewers |

  |------------|----------|----------------|---------|

  | Content A  | Type 1   | 120            | 150     |

  | Content B  | Type 2   | 200            | 250     |
  
- **Heatmap** showing total visits by weekday:

  ![Heatmap Example](<img width="890" alt="image" src="https://github.com/user-attachments/assets/a0dfbba0-e727-4cf7-a340-fd96a54de620" />)
## Power Automate Workflow
1. For Custom Click Logs
   Create a custom list, that would be the storage path of the events received by the specific page
   <img width="959" alt="image" src="https://github.com/user-attachments/assets/df5cdbed-f3e4-48a8-9c5a-7165432c6054" />

3. 
## Notes

- The script assumes that the filenames in the `BSREPORTS` folder follow a specific format (`SiteAnalyticsData_<date>.xlsx`), where the date is used to identify the latest file.
- The script processes only `.xlsx` files and skips other file types.
- The final PDF report includes visualizations (charts) and data tables.
  

