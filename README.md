# Summer Camp Timetable Generator

This Python script, `generate_timetables.py`, is designed to automate the creation of personalized timetables for students attending a music summer camp. It processes a master schedule from an Excel file and generates individual Excel timetables for each student.

### **Core Functionality**

The script's primary goal is to transform a complex, multi-student master schedule into simple, easy-to-read timetables for each participant. It intelligently parses the master schedule, identifies which activities each student is assigned to, and compiles this information into a neatly formatted, individual Excel file.

### **Inputs**

The script relies on three main input files:

1.  **`flute-time-table.xlsx`**: This is the master Excel timetable.
    *   Each sheet within this workbook represents a single day of the camp.
    *   The first row of each sheet lists the names of the teachers or instructors.
    *   The first column lists the time slots for the day.
    *   The cells in the schedule contain the activities, which can be assigned to individual students (e.g., `F1`), groups (e.g., `Group 1`, `Group 2, 3 Acting Class`), or are common to all (e.g., `Lunch`, `Welcome Speech`).

2.  **`student_mapping.csv`**: A CSV file that maps unique student IDs (e.g., `F1`) to their full names (e.g., `John Doe`). This is used to create user-friendly filenames for the generated timetables.

3.  **`group_mapping.csv`**: A CSV file that defines the composition of various student groups. It maps group numbers to the student IDs that belong to them.

### **Process Flow**

The script executes the following steps:

1.  **Load Mappings**: It begins by reading the `student_mapping.csv` and `group_mapping.csv` files to understand student names and their group affiliations.
2.  **Read Master Timetable**: It opens the `flute-time-table.xlsx` workbook and processes each sheet (day) one by one. A key step here is to **unmerge** any merged cells in the Excel sheet to ensure that every cell has a value, which simplifies data extraction.
3.  **Identify All Students**: The script scans the entire schedule to compile a unique list of all student IDs mentioned.
4.  **Generate Individual Schedules**: For each unique student, the script iterates through every time slot of every day and determines their activity based on a clear hierarchy:
    *   **Direct Match**: It first looks for an activity cell that explicitly contains the student's ID.
    *   **Group Match**: If no direct match is found, it checks if the activity is assigned to a group the student belongs to.
    *   **Common Activity**: If neither of the above applies, it checks for camp-wide activities like "Lunch" or "Master Class".
    *   **Free Time**: If no specific activity is found for a time slot, it is marked as "Free Time".
5.  **Create Output Files**:
    *   For each student, a new Excel workbook is created.
    *   The student's personalized schedule is written to a single sheet titled "Full Timetable". The layout is organized with time slots in the first column and the days of the week spread across the subsequent columns.
    *   To improve readability, activities that span multiple time slots are merged into a single cell.
    *   The columns are auto-sized to fit the content.
    *   The resulting file is named using the student's full name (e.g., `John Doe_timetable.xlsx`) and saved.

### **Output**

The script's output is a collection of Excel files, one for each student, saved in a directory named `student_timetables`. Each file provides a clear and personalized schedule for the entire duration of the camp.

### **Dependencies**

The script requires the `openpyxl` library to handle Excel files. This dependency is listed in the `requirements.txt` file.