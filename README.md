# Employee Attendance Tracker

This project is a Python-based script that reads an attendance sheet from an Excel file, processes the data to calculate the number of days each employee was present, and generates a new Excel report summarizing the attendance. It uses the `openpyxl` library for handling Excel files.

## Features
- Reads attendance data from an Excel file.
- Processes and counts the number of days each employee is marked as "Present."
- Exports the results into a new Excel file with a summary of employee attendance.

## Requirements
- Python 3.7 or higher
- `openpyxl` library

Install the required library with the following command:

```bash
pip install openpyxl
```

## File Structure
- **attendance_sheet.xlsx**: Input Excel file containing attendance data.
- **employee_attendance_report.xlsx**: Output Excel file summarizing employee attendance.

## Usage

### Input File Format
The input file (`attendance_sheet.xlsx`) should contain an `Attendance` sheet with the following columns:

| Column Name      | Description                           |
|------------------|---------------------------------------|
| Employee Name    | Name of the employee (Column C)       |
| Attendance Status| Status of attendance (e.g., Present) (Column D) |

### Script Execution
1. Place the input file `attendance_sheet.xlsx` in the same directory as the script.
2. Run the script using the following command:

```bash
python attendance_tracker.py
```

3. The output file, `employee_attendance_report.xlsx`, will be generated in the same directory.

### Output File Format
The output file contains two columns:

| Column Name      | Description                           |
|------------------|---------------------------------------|
| Employee Name    | Name of the employee                 |
| Number of Days Present | Total days the employee was marked "Present" |


## Customization
- Modify the input file name if the attendance file is named differently.
- Adjust column indexes in the script if the input file structure changes.


## Contributing
Feel free to fork the repository and submit pull requests with improvements or additional features.

---

### Author
Ahmed Fuseini
