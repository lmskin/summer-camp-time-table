import os
import glob
import win32com.client as win32
import pythoncom

def convert_excel_to_pdf(xlsx_file_path, pdf_file_path):
    """
    Convert Excel file to PDF using Excel COM automation.
    This preserves all formatting, merged cells, and styling.
    """
    excel_app = None
    workbook = None
    
    try:
        # Initialize COM
        pythoncom.CoInitialize()
        
        # Convert paths to absolute paths
        abs_xlsx_path = os.path.abspath(xlsx_file_path)
        abs_pdf_path = os.path.abspath(pdf_file_path)
        
        # Create Excel application
        excel_app = win32.Dispatch("Excel.Application")
        excel_app.Visible = False
        excel_app.DisplayAlerts = False
        excel_app.ScreenUpdating = False
        
        # Open workbook
        workbook = excel_app.Workbooks.Open(abs_xlsx_path)
        
        # Get the active worksheet
        worksheet = workbook.ActiveSheet
        
        # Set page setup for landscape and fit to one page
        worksheet.PageSetup.Orientation = 2  # xlLandscape
        worksheet.PageSetup.Zoom = False
        worksheet.PageSetup.FitToPagesWide = 1
        worksheet.PageSetup.FitToPagesTall = 1
        worksheet.PageSetup.PrintArea = worksheet.UsedRange.Address
        
        # Set margins (in inches)
        worksheet.PageSetup.LeftMargin = excel_app.InchesToPoints(0.3)
        worksheet.PageSetup.RightMargin = excel_app.InchesToPoints(0.3)
        worksheet.PageSetup.TopMargin = excel_app.InchesToPoints(0.3)
        worksheet.PageSetup.BottomMargin = excel_app.InchesToPoints(0.3)
        worksheet.PageSetup.HeaderMargin = excel_app.InchesToPoints(0.1)
        worksheet.PageSetup.FooterMargin = excel_app.InchesToPoints(0.1)
        
        # Export to PDF with minimal parameters to avoid version compatibility issues
        workbook.ExportAsFixedFormat(0, abs_pdf_path)  # 0 = xlTypePDF
        
        # Close workbook and quit Excel
        workbook.Close(SaveChanges=False)
        excel_app.Quit()
        
        return True
        
    except Exception as e:
        print(f"Error converting {xlsx_file_path} to PDF: {e}")
        return False
        
    finally:
        # Clean up resources
        try:
            if workbook is not None:
                workbook.Close(SaveChanges=False)
        except:
            pass
        
        try:
            if excel_app is not None:
                excel_app.Quit()
        except:
            pass
        
        try:
            pythoncom.CoUninitialize()
        except:
            pass

def convert_teacher_timetables_to_pdf():
    """
    Converts all Excel files in the teacher_timetables folder to PDF format.
    """
    teacher_timetables_dir = "teacher_timetables"
    
    # Check if the teacher_timetables directory exists
    if not os.path.exists(teacher_timetables_dir):
        print(f"Error: Directory '{teacher_timetables_dir}' not found.")
        return
    
    # Find all Excel files in the teacher_timetables directory
    xlsx_pattern = os.path.join(teacher_timetables_dir, "*.xlsx")
    xlsx_files = glob.glob(xlsx_pattern)
    
    if not xlsx_files:
        print(f"No Excel files found in '{teacher_timetables_dir}' directory.")
        return
    
    print(f"Found {len(xlsx_files)} Excel files to convert to PDF:")
    for file in xlsx_files:
        print(f"  - {os.path.basename(file)}")
    
    converted_count = 0
    failed_count = 0
    
    for xlsx_file in xlsx_files:
        # Create PDF filename by replacing .xlsx extension with .pdf
        pdf_file = xlsx_file.replace('.xlsx', '.pdf')
        
        print(f"\nConverting: {os.path.basename(xlsx_file)} -> {os.path.basename(pdf_file)}")
        
        # Convert to PDF
        if convert_excel_to_pdf(xlsx_file, pdf_file):
            print(f"✅ Successfully converted: {os.path.basename(pdf_file)}")
            converted_count += 1
        else:
            print(f"❌ Failed to convert: {os.path.basename(xlsx_file)}")
            failed_count += 1
    
    print(f"\n{'='*50}")
    print(f"Conversion Summary:")
    print(f"✅ Successfully converted: {converted_count} files")
    print(f"❌ Failed conversions: {failed_count} files")
    print(f"📁 PDF files saved in: {teacher_timetables_dir}")

if __name__ == '__main__':
    convert_teacher_timetables_to_pdf() 