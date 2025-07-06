import re
import os
import itertools
import datetime
import csv
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Alignment, Font, Border, Side
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

def load_student_name_mapping(filename=os.path.join("input", "student_mapping.csv")):
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

def generate_teacher_timetables(input_filename):
    """
    Reads an Excel file with multiple sheets (each representing a date) and
    generates an individual Excel timetable for each teacher.
    """
    music_instrument = os.path.basename(input_filename).split('-')[0].capitalize()

    try:
        workbook = load_workbook(input_filename, data_only=True)
        print(f"\nProcessing teacher timetables for {os.path.basename(input_filename)}...")
    except FileNotFoundError:
        print(f"Error: {input_filename} not found.")
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

    # Define border style
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    # Determine start date and camp name based on filename
    start_date = None
    camp_name = ""
    if 'campa' in input_filename.lower():
        start_date = datetime.date(2025, 7, 14)
        camp_name = "CampA"
    elif 'campb' in input_filename.lower():
        start_date = datetime.date(2025, 7, 21)
        camp_name = "CampB"

    for teacher in sorted(list(all_teachers)):
        all_time_slots = set()
        daily_schedules = {}

        for day_index, sheet_name in enumerate(workbook.sheetnames):
            if sheet_name not in processed_sheets:
                continue
            
            is_day_6 = (day_index + 1) == 6
            sheet_data = processed_sheets[sheet_name]
            header = [str(h).strip() for h in sheet_data[0]]
            
            try:
                # Find the column index for the current teacher
                teacher_col_index = header.index(teacher)
            except ValueError:
                if is_day_6:
                    teacher_col_index = 1  # Assume the second column for Day 6 activities
                else:
                    # Teacher not present in this sheet
                    continue

            schedule_rows = sheet_data[2:]
            teacher_schedule = []
            
            for row in schedule_rows:
                time_val = row[0]
                time = time_val.strftime('%H:%M') if isinstance(time_val, datetime.time) else str(time_val).strip()
                
                if not time:
                    continue

                current_time_obj = None
                try:
                    current_time_obj = datetime.datetime.strptime(time, '%H:%M').time()
                    if current_time_obj >= datetime.time(22, 0):
                        continue

                    if is_day_6 and current_time_obj < datetime.time(11, 0):
                        continue
                except ValueError:
                    pass

                # Get the activity from the teacher's column
                activity = str(row[teacher_col_index]).strip() if len(row) > teacher_col_index else ""

                if is_day_6:
                    if "Lunch" in activity and "Dress Up, Warm Up" in activity:
                        activity = "Lunch"
                    elif "Concert call time" in activity:
                        activity = "Lunch"
                else: # Special processing only for days other than Day 6
                    if activity.lower().startswith("workshop") or activity.lower().startswith("briefing for saturday"):
                        activity = ""

                    if activity:
                        # Rename group activities like "Group 6" to "Group 6 Ensemble Coaching"
                        activity = re.sub(r'(Group\s+\d+)(?! Ensemble Coaching)', r'\1 Ensemble Coaching', activity)

                        # Find all student IDs (e.g., F1) in the activity string
                        student_ids = re.findall(r'\bF\d+\b', activity)
                        for student_id in student_ids:
                            # Replace each student ID with their name, if available
                            student_name = student_name_map.get(student_id, student_id)
                            activity = activity.replace(student_id, student_name)
                
                if activity:
                    teacher_schedule.append((time, activity))
                else:
                    teacher_schedule.append((time, ""))

            # For weekdays, manually add the evening merge block
            if not is_day_6:
                evening_times = [
                    "19:00", "19:15", "19:30", "19:45",
                    "20:00", "20:15", "20:30", "20:45",
                    "21:00", "21:15", "21:30", "21:45"
                ]
                for evening_time in evening_times:
                    teacher_schedule.append((evening_time, "EVENING_MERGE_BLOCK"))

            # For Saturday, manually add the morning merge block
            if is_day_6:
                morning_times = ["10:00", "10:15", "10:30", "10:45"]
                for morning_time in morning_times:
                    teacher_schedule.append((morning_time, "SATURDAY_MORNING_MERGE_BLOCK"))

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
                header_text = current_date.strftime('%d %B (%A)')
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
                
                cell_activity = activity
                if activity == "EVENING_MERGE_BLOCK" or activity == "SATURDAY_MORNING_MERGE_BLOCK":
                    cell_activity = ""

                # Ensure room information is on a new line, but only if it's not already.
                if '(' in cell_activity and ')' in cell_activity and '\n(' not in cell_activity:
                    cell_activity = cell_activity.replace('(', '\n(', 1)

                # Ensure "Ensemble Coaching" is on a new line
                if 'Ensemble Coaching' in cell_activity and '\nEnsemble Coaching' not in cell_activity:
                    cell_activity = cell_activity.replace(' Ensemble Coaching', '\nEnsemble Coaching', 1)

                cell = teacher_ws.cell(row=start_row, column=current_col, value=cell_activity)

                if row_span > 1:
                    end_row = start_row + row_span - 1
                    teacher_ws.merge_cells(start_row=start_row, start_column=current_col, end_row=end_row, end_column=current_col)

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

        # Apply borders and alignment to all cells
        for row in teacher_ws.iter_rows(min_row=1, max_row=teacher_ws.max_row, min_col=1, max_col=teacher_ws.max_column):
            for cell in row:
                cell.border = thin_border
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

        sanitized_file_name = sanitize_filename(teacher)
        camp_part = f"_{camp_name}" if camp_name else ""
        file_path = os.path.join(output_dir, f'{sanitized_file_name}{camp_part}_timetable.xlsx')
        teacher_wb.save(file_path)

    print(f"Successfully generated timetables for {len(all_teachers)} teachers for {os.path.basename(input_filename)} in the '{output_dir}' directory.")

if __name__ == '__main__':
    input_dir = "input"
    if not os.path.isdir(input_dir):
        print(f"Error: Input directory '{input_dir}' not found or is not a directory.")
    else:
        # Regex to match the expected filename format for specific instruments
        filename_pattern = re.compile(r"(cello|flute|harp)-(camp[ab])\-time-table\.xlsx", re.IGNORECASE)
        
        timetable_files = [f for f in os.listdir(input_dir) if filename_pattern.match(f)]

        if not timetable_files:
            print(f"No timetable files matching the pattern '{{music-instrument}}-{{campA or campB}}-time-table.xlsx' were found in '{input_dir}'.")
        else:
            print(f"Found {len(timetable_files)} timetable file(s) to process: {', '.join(sorted(timetable_files))}")
            for filename in sorted(timetable_files):
                full_path = os.path.join(input_dir, filename)
                generate_teacher_timetables(full_path)
