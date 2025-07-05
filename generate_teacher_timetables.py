import re
import os
import itertools
import datetime
import csv
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

def sanitize_filename(filename):
    """
    Removes characters from a string that are not allowed in file names.
    """
    return re.sub(r'[\\/*?:"<>|\n]', '', filename)

def process_sheet(sheet):
    """
    Processes a single sheet to extract its data, handling merged cells by
    unmerging them and filling the values.
    """
    merged_ranges = list(sheet.merged_cells.ranges)
    for merged_cell_range in merged_ranges:
        min_col, min_row, max_col, max_row = merged_cell_range.bounds
        top_left_cell_value = sheet.cell(row=min_row, column=min_col).value
        sheet.unmerge_cells(str(merged_cell_range))
        for row in sheet.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
            for cell in row:
                cell.value = top_left_cell_value

    data = []
    for row in sheet.iter_rows():
        data.append([cell.value if cell.value is not None else "" for cell in row])
    return data

def load_student_name_mapping(filename="student_mapping.csv"):
    """
    Loads student_no to student_name mappings from the specified CSV file.
    """
    name_map = {}
    try:
        with open(filename, mode='r', encoding='utf-8-sig') as infile:
            reader = csv.DictReader(infile)
            for row in reader:
                student_no = row.get('student_no')
                student_name = row.get('student_name')
                if student_no and student_name:
                    name_map[student_no.strip()] = student_name.strip()
    except FileNotFoundError:
        print(f"Warning: {filename} not found. Student numbers will be used in filenames.")
    except Exception as e:
        print(f"An error occurred while reading {filename}: {e}")
    return name_map

def generate_teacher_timetables():
    """
    Reads an Excel file with multiple sheets (each representing a date) and
    generates an individual Excel timetable for each teacher.
    """
    input_filename = "flute-campA-time-table.xlsx"
    music_instrument = input_filename.split('-')[0].capitalize()

    try:
        workbook = load_workbook(input_filename, data_only=True)
        print(f"Processing teacher timetables for {music_instrument}...")
    except FileNotFoundError:
        print(f"Error: {input_filename} not found. Make sure the file is in the same directory.")
        return

    student_name_map = load_student_name_mapping()

    processed_sheets = {}
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        processed_sheets[sheet_name] = process_sheet(sheet)

    all_teachers = set()
    for sheet_name, sheet_data in processed_sheets.items():
        if sheet_data:
            # Teacher names are in the first row, starting from the second column
            teacher_row = sheet_data[0]
            for teacher_name in teacher_row[1:]:
                if teacher_name and isinstance(teacher_name, str) and teacher_name.strip():
                    all_teachers.add(teacher_name.strip())

    output_dir = "teacher_timetables"
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Determine start date based on filename
    start_date = None
    if 'campa' in input_filename.lower():
        start_date = datetime.date(2025, 7, 14)
    elif 'campb' in input_filename.lower():
        start_date = datetime.date(2025, 7, 21)

    for teacher in sorted(list(all_teachers)):
        all_time_slots = set()
        daily_schedules = {}

        for sheet_name in workbook.sheetnames:
            if sheet_name not in processed_sheets:
                continue
            
            sheet_data = processed_sheets[sheet_name]
            header = [str(h).strip() for h in sheet_data[0]]
            
            try:
                # Find the column index for the current teacher
                teacher_col_index = header.index(teacher)
            except ValueError:
                # Teacher not present in this sheet
                continue

            schedule_rows = sheet_data[2:]
            teacher_schedule = []
            
            for row in schedule_rows:
                time_val = row[0]
                time = time_val.strftime('%H:%M') if isinstance(time_val, datetime.time) else str(time_val).strip()
                
                if not time:
                    continue

                # Get the activity from the teacher's column
                activity = str(row[teacher_col_index]).strip() if len(row) > teacher_col_index else ""

                if activity:
                    # Find all student IDs (e.g., F1) in the activity string
                    student_ids = re.findall(r'\bF\d+\b', activity)
                    for student_id in student_ids:
                        # Replace each student ID with their name, if available
                        student_name = student_name_map.get(student_id, student_id)
                        activity = activity.replace(student_id, student_name)
                    
                    teacher_schedule.append((time, activity))
                else:
                    teacher_schedule.append((time, "Free Time"))

            daily_schedules[sheet_name] = teacher_schedule
            for time, _ in teacher_schedule:
                all_time_slots.add(time)
        
        if not daily_schedules:
            continue

        sorted_times = sorted(list(all_time_slots))
        time_to_row = {time: i + 3 for i, time in enumerate(sorted_times)}

        teacher_wb = Workbook()
        teacher_ws = teacher_wb.active
        teacher_ws.title = "Full Timetable"

        teacher_ws.cell(row=2, column=1, value="Time").font = Font(bold=True)
        for i, time in enumerate(sorted_times):
            teacher_ws.cell(row=i + 3, column=1, value=time)

        current_col = 2
        for day_index, sheet_name in enumerate(workbook.sheetnames):
            if sheet_name not in daily_schedules:
                continue
            
            if start_date:
                current_date = start_date + datetime.timedelta(days=day_index)
                header_text = current_date.strftime('%d %b (%A)')
            else:
                header_text = sheet_name

            teacher_ws.cell(row=1, column=current_col, value=header_text).font = Font(bold=True)
            todays_schedule = daily_schedules[sheet_name]

            for activity, group in itertools.groupby(todays_schedule, key=lambda x: x[1]):
                group_list = list(group)
                row_span = len(group_list)
                start_time = group_list[0][0]
                
                if start_time not in time_to_row: continue

                start_row = time_to_row[start_time]
                cell = teacher_ws.cell(row=start_row, column=current_col, value=activity)

                if row_span > 1:
                    end_row = start_row + row_span - 1
                    teacher_ws.merge_cells(start_row=start_row, start_column=current_col, end_row=end_row, end_column=current_col)
                    cell.alignment = Alignment(vertical='center')

            current_col += 1

        # Set column widths
        for column in teacher_ws.columns:
            column_letter = get_column_letter(column[0].column)
            if column[0].column == 1:  # First column
                max_length = 0
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = (max_length + 2)
                teacher_ws.column_dimensions[column_letter].width = adjusted_width
            else:  # Other columns
                teacher_ws.column_dimensions[column_letter].width = 35

        sanitized_file_name = sanitize_filename(teacher)
        file_path = os.path.join(output_dir, f'{sanitized_file_name}_timetable.xlsx')
        teacher_wb.save(file_path)

    print(f"Successfully generated timetables for {len(all_teachers)} teachers in the '{output_dir}' directory.")

if __name__ == '__main__':
    generate_teacher_timetables()
