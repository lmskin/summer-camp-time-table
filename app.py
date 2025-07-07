import streamlit as st
import os
import re
import zipfile
import time
from generate_student_timetables import generate_timetables as generate_student_timetables
from generate_teacher_timetables import generate_teacher_timetables

def main():
    st.set_page_config(page_title="Summer Camp Timetable Generator", layout="wide")
    st.title("Summer Camp Timetable Generator")

    # Ensure necessary directories exist
    os.makedirs("input", exist_ok=True)
    os.makedirs("student_timetables", exist_ok=True)
    os.makedirs("teacher_timetables", exist_ok=True)

    st.sidebar.header("Instructions")
    st.sidebar.info(
        "1. **Upload mapping files** using the uploader below.\n\n"
        "2. **Upload your timetable file.** The filename must be in the format: "
        "`{music-instrument}-{campA or campB}-time-table.xlsx`.\n"
        "   For example: `flute-campA-time-table.xlsx`.\n\n"
        "3. **Select which timetables** you want to generate (Student, Teacher, or Both).\n\n"
        "4. **Click 'Generate Timetables'**.\n\n"
        "5. **Download the generated files** as a ZIP archive."
    )
    
    st.sidebar.markdown("---")
    st.sidebar.header("Upload Mapping Files")
    st.sidebar.info(
        "Upload all `student_mapping`, `group_mapping`, `room_mapping`, and `room_no_mapping` CSV files for your camps."
    )
    mapping_files = st.sidebar.file_uploader("Upload Mapping CSVs", type=["csv"], accept_multiple_files=True)


    uploaded_file = st.file_uploader("Upload Excel Timetable", type=["xlsx"])
    
    generation_option = st.radio(
        "Select which timetables to generate:",
        ('Student Timetables', 'Teacher Timetables', 'Both')
    )

    if st.button("Generate Timetables"):
        if uploaded_file is not None:
            # Save mapping files first, so they are available to the generator scripts.
            if mapping_files:
                st.write(f"DEBUG: Found {len(mapping_files)} mapping files to upload:")
                for mapping_file in mapping_files:
                    st.write(f"  - {mapping_file.name}")
                    # Save each mapping file to the 'input' directory.
                    mapping_filepath = os.path.join("input", mapping_file.name)
                    with open(mapping_filepath, "wb") as f:
                        f.write(mapping_file.getbuffer())
                    
                    # Verify the file was saved correctly
                    if os.path.exists(mapping_filepath):
                        file_size = os.path.getsize(mapping_filepath)
                        st.write(f"  ✓ Saved {mapping_file.name} ({file_size} bytes)")
                    else:
                        st.error(f"  ✗ Failed to save {mapping_file.name}")
                
                st.sidebar.success(f"{len(mapping_files)} mapping files uploaded successfully.")
                
                # List all files in input directory for verification
                st.write("DEBUG: Files in input directory after upload:")
                for file in os.listdir("input"):
                    st.write(f"  - {file}")
                    
            else:
                # Warn the user if no mapping files are provided, as it can affect output.
                st.warning("Warning: No mapping files uploaded. Name and room mappings may not work correctly.")

            filename = uploaded_file.name
            
            # Pattern to validate filename based on project conventions.
            # It allows for instrument names with spaces and optional versioning.
            filename_pattern = re.compile(r"[\w\s]+-(camp[ab])-time-table.*\.xlsx", re.IGNORECASE)
            
            if filename_pattern.match(filename):
                # Save the uploaded file to the 'input' directory so scripts can find it
                # and associated mapping files.
                input_filepath = os.path.join("input", filename)
                with open(input_filepath, "wb") as f:
                    f.write(uploaded_file.getbuffer())
                
                st.info(f"Processing '{filename}'...")
                
                # Clean output directories before generation to prevent mixing results
                clear_output_dirs()

                with st.spinner('Generating timetables... Please wait.'):
                    try:
                        if generation_option in ('Student Timetables', 'Both'):
                            st.write("Generating student timetables...")
                            generate_student_timetables(input_filepath)
                        
                        if generation_option in ('Teacher Timetables', 'Both'):
                            st.write("Generating teacher timetables...")
                            generate_teacher_timetables(input_filepath)
                        
                        st.success("Timetable generation complete!")
                        
                        # Create a zip file of the output for easy download
                        zip_path = create_zip_of_output(generation_option)
                        
                        if zip_path:
                            with open(zip_path, "rb") as f:
                                st.download_button(
                                    label="Download Timetables ZIP",
                                    data=f,
                                    file_name=os.path.basename(zip_path),
                                    mime="application/zip"
                                )
                            os.remove(zip_path) # Clean up the created zip file
                        else:
                            st.warning("No timetables were generated. Please check the input file and console logs for errors.")
                            
                    except Exception as e:
                        st.error(f"An error occurred during timetable generation: {e}")
                        st.exception(e) # Provides a full traceback in the UI for debugging

            else:
                st.error(f"Invalid filename format. Please use the format '{{music-instrument}}-{{campA or campB}}-time-table.xlsx'.")
        else:
            st.warning("Please upload a timetable file first.")

def clear_output_dirs():
    """Removes all files from the output directories to ensure a clean run."""
    for folder in ["student_timetables", "teacher_timetables"]:
        if os.path.exists(folder):
            for filename in os.listdir(folder):
                file_path = os.path.join(folder, filename)
                try:
                    if os.path.isfile(file_path) or os.path.islink(file_path):
                        os.unlink(file_path)
                except Exception as e:
                    st.error(f"Failed to delete {file_path}. Reason: {e}")

def create_zip_of_output(option):
    """Zips the contents of the relevant output directories."""
    timestamp = time.strftime("%Y%m%d-%H%M%S")
    zip_filename = f"timetables_{timestamp}.zip"
    
    dirs_to_zip = []
    if option in ('Student Timetables', 'Both'):
        dirs_to_zip.append("student_timetables")
    if option in ('Teacher Timetables', 'Both'):
        dirs_to_zip.append("teacher_timetables")
        
    files_were_added = False
    with zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for dir_to_zip in dirs_to_zip:
            if os.path.exists(dir_to_zip):
                for root, _, files in os.walk(dir_to_zip):
                    for file in files:
                        if file.endswith('.xlsx'):
                            file_path = os.path.join(root, file)
                            # Add file to zip, using a relative path inside the zip
                            zipf.write(file_path, os.path.relpath(file_path, start=os.curdir))
                            files_were_added = True
    
    if files_were_added:
        return zip_filename
    else:
        # If no files were generated, no zip is created
        return None

if __name__ == '__main__':
    main() 