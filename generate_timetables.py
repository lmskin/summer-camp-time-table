import re
import os
import itertools
import datetime
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Alignment

def process_sheet(sheet):
    """
    Processes a single sheet to extract its data, handling merged cells by
    unmerging them and filling the values.
    """
    # Create a copy of the merged cell ranges to iterate over, as unmerging
    # will modify the sheet's merged_cells attribute.
    merged_ranges = list(sheet.merged_cells.ranges)
    for merged_cell_range in merged_ranges:
        min_col, min_row, max_col, max_row = merged_cell_range.bounds
        top_left_cell_value = sheet.cell(row=min_row, column=min_col).value
        sheet.unmerge_cells(str(merged_cell_range))
        for row in sheet.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
            for cell in row:
                cell.value = top_left_cell_value

    # Read the data from the now unmerged sheet into a 2D list
    data = []
    for row in sheet.iter_rows():
        data.append([cell.value if cell.value is not None else "" for cell in row])
    return data

def generate_timetables():
    """
    Reads an Excel file with multiple sheets (each representing a date) and
    generates an individual Excel timetable for each student. Each student's
    timetable will contain multiple sheets, corresponding to the dates in the
    input file.
    """
    input_filename = 'template-time-table.xlsx'
    try:
        # Load the workbook once to process all sheets.
        # data_only=True ensures we get cell values instead of formulas.
        workbook = load_workbook(input_filename, data_only=True)
    except FileNotFoundError:
        print(f"Error: {input_filename} not found. Make sure the file is in the same directory.")
        return

    # Process all sheets and store their data in a dictionary
    processed_sheets = {}
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        processed_sheets[sheet_name] = process_sheet(sheet)

    # Find all unique students across all processed sheets
    all_students = set()
    for sheet_name, sheet_data in processed_sheets.items():
        schedule_rows = sheet_data[2:]  # Schedule data starts from the third row
        for row in schedule_rows:
            for cell in row[1:]:
                if isinstance(cell, str) and re.match(r'F\d+', cell.strip()):
                    all_students.add(cell.strip())

    common_activities = [
        "Welcome Speech",
        "Lunch",
        "Break",
        "Ensemble Coaching",
        "Faculty Rehearsal",
        "Workshop"
    ]

    output_dir = "student_timetables"
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Generate one multi-sheet Excel file for each student
    for student in sorted(list(all_students)):
        student_wb = Workbook()  # Create a new workbook for the student
        student_wb.remove(student_wb.active)  # Remove the default sheet

        # Iterate through each day's schedule to create a sheet for it
        for sheet_name, sheet_data in processed_sheets.items():
            student_ws = student_wb.create_sheet(title=sheet_name)
            student_ws.append(['Time', 'Activity'])

            teachers = [str(name).strip() for name in sheet_data[0][1:]]
            schedule_rows = sheet_data[2:]

            student_schedule = []
            for row in schedule_rows:
                time_val = row[0]
                if isinstance(time_val, datetime.time):
                    time = time_val.strftime('%H:%M')
                else:
                    time = str(time_val).strip()

                activities = [str(act).strip() for act in row[1:]]
                activity_found_for_time = False

                if student in activities:
                    teacher_index = activities.index(student)
                    teacher = teachers[teacher_index]
                    student_schedule.append((time, f"Class with {teacher}"))
                    activity_found_for_time = True
                
                if not activity_found_for_time:
                    for common_activity in common_activities:
                        if common_activity in activities:
                            student_schedule.append((time, common_activity))
                            activity_found_for_time = True
                            break
                
                if not activity_found_for_time:
                    student_schedule.append((time, "Free Time"))

            # Write the generated schedule to the student's sheet for the day
            row_to_write = 2
            for activity, group in itertools.groupby(student_schedule, key=lambda x: x[1]):
                group_list = list(group)
                row_span = len(group_list)
                
                start_time = group_list[0][0]
                student_ws.cell(row=row_to_write, column=1, value=start_time)
                student_ws.cell(row=row_to_write, column=2, value=activity)
                
                for i in range(1, row_span):
                    student_ws.cell(row=row_to_write + i, column=1, value=group_list[i][0])

                if row_span > 1:
                    student_ws.merge_cells(start_row=row_to_write, start_column=2, end_row=row_to_write + row_span - 1, end_column=2)
                    merged_cell = student_ws.cell(row=row_to_write, column=2)
                    merged_cell.alignment = Alignment(vertical='center')
                
                row_to_write += row_span

            student_ws.column_dimensions['A'].width = 15
            student_ws.column_dimensions['B'].width = 30
        
        # Save the student's complete multi-sheet timetable
        file_path = os.path.join(output_dir, f'{student}_timetable.xlsx')
        student_wb.save(file_path)

    print(f"Successfully generated timetables for {len(all_students)} students in the '{output_dir}' directory.")


if __name__ == '__main__':
    generate_timetables() 