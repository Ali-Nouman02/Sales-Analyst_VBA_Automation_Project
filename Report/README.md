# Report Module 

The Report Module processes and summarizes regional sales data based on user selection. It provides region-specific insights, generates tailored workbooks for regional managers and top management, and exports data as a CSV file for further use.
---

## What the Code does: 

### Regional Analysis
- The region is selected via a dropdown on the **Start Page**.
  ![create_report_dropdown](https://github.com/user-attachments/assets/33fe9a93-30d8-4d8f-bb8b-53bc2804606c)
- Extracts relevant sales data for the selected region (**Europe** or **America**) and populates it in the **Region tab**.
  ![region_tab](https://github.com/user-attachments/assets/25db37f4-7ccb-45e3-aefe-c0de62d13a70)
- Converts sales data from **Euro to Dollar** when applicable.  

### Explanation of the Pivot Tables
The **Region tab** contains three pivot tables providing detailed insights:  
1. **Sales by Customer**:  
   - Displays the total sales conducted by each customer in dollar terms.  

2. **Market Summary**:  
   - Shows the total value of sales and the count of rejected products in the selected market.  

3. **Complete Regional Summary**:  
   - Breaks down sales by customer, sales quantity, and dollar amount in a categorized format.

### Workbook Creation
1. **Regional Manager Workbook**:  
   - Transfers the first two pivot tables to a new workbook.  
   - File is named in the format: `RM_YEAR_MONTH_REGION` (e.g., `RM_201810America`).  
   - Saved in the current directory.  
[RM_201810America.xlsx](https://github.com/user-attachments/files/18136787/RM_201810America.xlsx)
2. **Top Management Workbook**:  
   - Transfers the third pivot table to a separate workbook.  
   - The workbook is **hardcoded** (contains no formulas).  
   - File is named in the format: `M_YEAR_MONTH_REGION` (e.g., `M_201810America`).  
   - Saved in the current directory.  
![topmanager_report](https://github.com/user-attachments/assets/f63cbba6-dca5-40c0-8b02-08e8f0c81ff1)

3. **Task Tracking**:  
   - After each workbook creation, a note is added under **Task Completed** on the **Start Page**.  
![task completion confirmation_2](https://github.com/user-attachments/assets/5647a6df-d514-4055-b1e3-7f266df7aa48)

### CSV File Creation
- A sub-procedure, `Create_CSV`, generates a CSV file based on the **Summary tab** created during the Import Module.  
- Uses `";"` as the delimiter.  
- File is saved in the current directory.  
![csv](https://github.com/user-attachments/assets/545d9ebf-7b77-4f9c-b383-ae79f967cf10)
![csv task completeion](https://github.com/user-attachments/assets/38506974-3fc2-4913-a90e-4444fae8fc7b)

---

## Code Workflow

1. **Region Selection**:  
   Select the desired region (Europe or America) on the **Start Page** dropdown.

2. **Data Processing**:  
   - Extract relevant data to the **Region tab**.  
   - Summarize data using three pivot tables.

3. **Workbook Creation**:  
   - Generate separate workbooks for Regional Managers and Top Management.

4. **CSV Generation**:  
   - Export summarized data to a semicolon-delimited CSV file.

5. **Task Tracking**:  
   - Record all completed tasks under **Task Completed**.

---

## File Naming Conventions
- **Regional Manager Workbook**: `RM_YEAR_MONTH_REGION`  
  Example: `RM_201810America`  
- **Top Management Workbook**: `M_YEAR_MONTH_REGION`  
  Example: `M_201810America`  
- **CSV File**: `summary.csv` (saved in the current directory).

---

## Attachments
- **Code**:  
  Include the VBA code used for the Report Module.

- **Pictures**:  
  Add screenshots of the generated pivot tables, workbooks, and CSV outputs.

---

## Benefits
- Streamlined region-specific reporting.  
- Tailored workbooks for different audiences (Regional Managers and Top Management).  
- Easy CSV export for database uploads.  
- Automatic task tracking for better transparency.


