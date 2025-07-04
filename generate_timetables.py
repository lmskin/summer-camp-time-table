import re
import os
import itertools
import datetime
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

def sanitize_filename(filename):
    """
    Removes characters from a string that are not allowed in file names.
    This includes newline characters, which can cause issues with file paths.
    """
    return re.sub(r'[\\/*?:"<>|\\n]', '', filename)

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

def load_group_mappings(filename="mapping.xlsx"):
    """
    Loads group-to-student mappings from the 'group' sheet in mapping.xlsx.
    Returns a dictionary where keys are group names and values are lists of student IDs.
    """
    group_mappings = {}
    try:
        workbook = load_workbook(filename, data_only=True)
        if "group" not in workbook.sheetnames:
            print("Warning: 'group' sheet not found in mapping.xlsx. No group activities will be added.")
            return group_mappings
        
        sheet = workbook["group"]
        # Assumes header in the first row, data starts from the second
        header = [cell.value for cell in sheet[1]]
        try:
            group_col_idx = header.index("group_number") + 1
            student_col_idx = header.index("student_no") + 1
        except ValueError:
            print("Warning: 'group_number' or 'student_no' column not found in 'group' sheet of mapping.xlsx.")
            return group_mappings
            
        for row in sheet.iter_rows(min_row=2):
            group_number = row[group_col_idx - 1].value
            student_ids_str = row[student_col_idx - 1].value
            if group_number and student_ids_str:
                group_name = f"Group {group_number}"
                student_ids = [s.strip() for s in str(student_ids_str).split(',')]
                if group_name not in group_mappings:
                    group_mappings[group_name] = []
                group_mappings[group_name].extend(student_ids)
    except FileNotFoundError:
        print("Warning: mapping.xlsx not found. No group activities will be added.")
    
    return group_mappings

def generate_timetables():
    """
    Reads an Excel file with multiple sheets (each representing a date) and
    generates an individual Excel timetable for each student. Each student's
    timetable will contain multiple sheets, corresponding to the dates in the
    input file.
    """
    input_filename = "flute-time-table.xlsx"
    
    # Extract instrument name, e.g., 'cello' from 'cello-time-table.xlsx'
    music_instrument = input_filename.replace('-time-table.xlsx', '')
    
    try:
        # Load the workbook once to process all sheets.
        # data_only=True ensures we get cell values instead of formulas.
        workbook = load_workbook(input_filename, data_only=True)
        print(f"Processing timetable for {music_instrument.capitalize()}...")
    except FileNotFoundError:
        # This case is less likely now but good to keep as a safeguard
        print(f"Error: {input_filename} not found. Make sure the file is in the same directory.")
        return

    # Load group mappings
    group_mappings = load_group_mappings()
    
    # Create a reverse mapping from student to their groups for efficient lookup
    student_to_groups = {}
    for group, students in group_mappings.items():
        for s in students:
            if s not in student_to_groups:
                student_to_groups[s] = []
            student_to_groups[s].append(group)

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
                if isinstance(cell, str):
                    # Use regex to find all student IDs (e.g., F1) in a cell
                    found_students = re.findall(r'\bF\d+\b', cell)
                    for s in found_students:
                        all_students.add(s)

    common_activities = [
        "Welcome Speech",
        "Lunch",
        "Break",
        "Ensemble Coaching",
        "Workshop",
        "Toilet Break"
    ]

    output_dir = "student_timetables"
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Generate one single-sheet Excel file for each student
    for student in sorted(list(all_students)):
        # Pre-process to gather all time slots and daily schedules for the student
        all_time_slots = set()
        daily_schedules = {}

        # Using workbook.sheetnames to preserve the order of days
        for sheet_name in workbook.sheetnames:
            if sheet_name not in processed_sheets:
                continue
            
            sheet_data = processed_sheets[sheet_name]
            teachers = [str(name).strip() for name in sheet_data[0][1:]]
            schedule_rows = sheet_data[2:]
            
            student_schedule = []
            for row in schedule_rows:
                time_val = row[0]
                time = time_val.strftime('%H:%M') if isinstance(time_val, datetime.time) else str(time_val).strip()

                activities = [str(act).strip() for act in row[1:]]
                activity_found_for_timeslot = False

                for i, activity in enumerate(activities):
                    if not activity:
                        continue

                    # Priority 1: Direct student match
                    if not activity_found_for_timeslot and re.search(r'\b' + re.escape(student) + r'\b', activity):
                        if activity.strip() == student:
                            teacher = teachers[i]
                            desc = f"Class with {teacher}" if teacher else f"Practice ({music_instrument} practice room)"
                            student_schedule.append((time, desc))
                        else:
                            student_schedule.append((time, activity))
                        activity_found_for_timeslot = True

                    # Priority 2: Complex group match
                    if not activity_found_for_timeslot and activity.lower().startswith('group') and "," in activity:
                        activity_body = activity[len('Group'):].strip()
                        split_index = -1
                        for i_rev in range(len(activity_body) - 1, -1, -1):
                            if activity_body[i_rev].isdigit():
                                split_index = i_rev + 1
                                break
                        if split_index != -1:
                            group_part = activity_body[:split_index]
                            activity_name = activity_body[split_index:].strip()
                            group_numbers_str = ''.join(filter(lambda x: x.isdigit() or x == ',', group_part))
                            activity_groups = {f"Group {num.strip()}" for num in group_numbers_str.split(',') if num.strip()}
                            student_groups = set(student_to_groups.get(student, []))
                            if not student_groups.isdisjoint(activity_groups):
                                student_schedule.append((time, activity_name))
                                activity_found_for_timeslot = True

                    # Priority 3: Simple group match
                    if not activity_found_for_timeslot and student in student_to_groups:
                        student_groups = student_to_groups[student]
                        if activity in student_groups:
                            student_schedule.append((time, activity))
                            activity_found_for_timeslot = True
                
                # After checking all cells in the row for the student's activities
                if not activity_found_for_timeslot:
                    # Check for common activities only if no specific activity was found
                    common_activity_found = False
                    for common_activity in common_activities:
                        if common_activity in activities:
                            student_schedule.append((time, common_activity))
                            common_activity_found = True
                            break
                    if not common_activity_found:
                        student_schedule.append((time, "Free Time"))

            daily_schedules[sheet_name] = student_schedule
            for time, _ in student_schedule:
                all_time_slots.add(time)
        
        sorted_times = sorted(list(all_time_slots))
        time_to_row = {time: i + 3 for i, time in enumerate(sorted_times)}

        student_wb = Workbook()
        student_ws = student_wb.active
        student_ws.title = "Full Timetable"

        student_ws.cell(row=2, column=1, value="Time").font = Font(bold=True)
        for i, time in enumerate(sorted_times):
            student_ws.cell(row=i + 3, column=1, value=time)

        current_col = 2
        for sheet_name in workbook.sheetnames:
            if sheet_name not in daily_schedules:
                continue
            
            student_ws.cell(row=1, column=current_col, value=sheet_name).font = Font(bold=True)
            student_ws.cell(row=2, column=current_col, value="Activity").font = Font(bold=True)

            todays_schedule = daily_schedules[sheet_name]

            for activity, group in itertools.groupby(todays_schedule, key=lambda x: x[1]):
                group_list = list(group)
                row_span = len(group_list)
                start_time = group_list[0][0]
                
                if start_time not in time_to_row: continue

                start_row = time_to_row[start_time]
                cell = student_ws.cell(row=start_row, column=current_col, value=activity)

                if row_span > 1:
                    end_row = start_row + row_span - 1
                    student_ws.merge_cells(start_row=start_row, start_column=current_col, end_row=end_row, end_column=current_col)
                    cell.alignment = Alignment(vertical='center')

            current_col += 2

        # Autofit all columns
        for column in student_ws.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2)
            student_ws.column_dimensions[column_letter].width = adjusted_width

        sanitized_student_name = sanitize_filename(student)
        file_path = os.path.join(output_dir, f'{sanitized_student_name}_timetable.xlsx')
        student_wb.save(file_path)

    print(f"Successfully generated timetables for {len(all_students)} students in the '{output_dir}' directory.")


if __name__ == '__main__':
    generate_timetables() 