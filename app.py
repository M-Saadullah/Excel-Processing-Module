import streamlit as st
import os
import shutil
import subprocess
import sys
from pathlib import Path
import time
import glob

# Configure the page
st.set_page_config(
    page_title="AI Document Processor",
    page_icon="🤖",
    layout="wide"
)

def setup_directories():
    """Create necessary directories if they don't exist."""
    base_dir = Path(__file__).parent
    input_folder = base_dir / "Input_Folder"
    updated_excel_workbooks = base_dir / "Updated_excel_workbooks"
    data_sources = base_dir / "data_sources"
    output_folder = base_dir / "output_folder"
    html_outputs = base_dir / "html_outputs"
    
    # Create directories
    input_folder.mkdir(exist_ok=True)
    updated_excel_workbooks.mkdir(exist_ok=True)
    data_sources.mkdir(exist_ok=True)
    output_folder.mkdir(exist_ok=True)
    html_outputs.mkdir(exist_ok=True)
    
    return input_folder, updated_excel_workbooks, data_sources, output_folder, html_outputs

def clear_directories(input_folder, updated_excel_workbooks, output_folder,html_outputs):
    """Clear previous files from input and output folders on Windows."""
    
    for folder in [input_folder, updated_excel_workbooks, output_folder, html_outputs]:
        folder_path = Path(folder)
        
        # Check if directory exists
        if not folder_path.exists():
            print(f"Directory does not exist: {folder}")
            continue
            
        if not folder_path.is_dir():
            print(f"Path is not a directory: {folder}")
            continue
            
        # Delete all files in the directory
        files_deleted = 0
        for file in folder_path.glob("*"):
            if file.is_file():
                try:
                    file.unlink()  # This works perfectly on Windows
                    files_deleted += 1
                except Exception as e:
                    print(f"Error deleting {file.name}: {e}")
        
        print(f"Cleared {files_deleted} files from: {folder}")


def run_main_processor():
    """Run the main processing function."""
    try:
        # Import and run your main function
        from excel_processor import main  # Replace 'your_main_script' with your actual script name
        main()
        return True
    except Exception as e:
        st.error(f"Error during processing: {str(e)}")
        return False

def get_output_files(updated_excel_workbooks):
    """Get all Excel files from the output folder."""
    excel_files = list(updated_excel_workbooks.glob("*.xlsx")) + list(updated_excel_workbooks.glob("*.xls"))
    return excel_files

def main():
    st.title("🤖 AI Document Processor")
    st.markdown("Upload your Excel workbook and text data source to process them with AI.")
    
    # Setup directories
    input_folder, updated_excel_workbooks, data_sources,output_folder,html_outputs = setup_directories()
    
    # Create two columns for file uploads
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📊 Excel Workbook")
        excel_file = st.file_uploader(
            "Choose an Excel file",
            type=['xlsx', 'xls'],
            help="Upload the Excel workbook you want to process"
        )
    
    with col2:
        st.subheader("📄 Text Data Sources")
        txt_files = st.file_uploader(
            "Choose text files (you can select multiple)",
            type=['txt'],
            accept_multiple_files=True,
            help="Upload one or more text files containing your data sources. Hold Ctrl/Cmd to select multiple files in the file browser."
        )
            
            # Show current selection count
        if txt_files:
            st.success(f"✅ {len(txt_files)} text file(s) selected")
    
    # Processing section
    st.markdown("---")
    
    if excel_file is not None and txt_files is not None:
        st.success("✅ Both files uploaded successfully!")
        
        # Display file information
        with st.expander("📋 File Information"):
            col1, col2 = st.columns(2)
            with col1:
                st.write(f"**Excel File:** {excel_file.name}")
                st.write(f"**Size:** {excel_file.size:,} bytes")
            with col2:
                st.write(f"**Text File:** {txt_files.count}")
        
        # Process button
        if st.button("🚀 Start Processing", type="primary", use_container_width=True):
            
            # Clear previous files
            clear_directories(input_folder, updated_excel_workbooks, output_folder,html_outputs)
            
            # Save uploaded files
# Save uploaded files
            try:
                # Save Excel file to Input_Folder
                excel_path = input_folder / excel_file.name
                with open(excel_path, "wb") as f:
                    f.write(excel_file.getbuffer())
                
                # Save all text files to data_sources folder
                saved_txt_files = []
                for txt_file in txt_files:
                    txt_path = data_sources / txt_file.name
                    with open(txt_path, "wb") as f:
                        f.write(txt_file.getbuffer())
                    saved_txt_files.append(txt_file.name)
                
                st.success(f"✅ Files saved successfully! ({len(saved_txt_files)} text files)")
                
                # Optional: Show which files were saved
                with st.expander("📁 Saved Files Details"):
                    st.write(f"**Excel file saved:** {excel_file.name}")
                    st.write(f"**Text files saved ({len(saved_txt_files)}):**")
                    for i, filename in enumerate(saved_txt_files, 1):
                        st.write(f"  {i}. {filename}")
                
            except Exception as e:
                st.error(f"❌ Error saving files: {str(e)}")
                return
            
            # Show processing status
            with st.status("🔄 Processing files...", expanded=True) as status:
                st.write("📁 Files uploaded and saved")
                st.write("🤖 Starting AI processing...")
                
                # Run the main processing function
                success = run_main_processor()
                
                if success:
                    st.write("✅ AI processing completed!")
                    st.write("📊 Excel updating completed!")
                    status.update(label="✅ Processing completed successfully!", state="complete")
                else:
                    status.update(label="⚠️ Processing completed with issues", state="error")
                    return
            
            # Check for output files
            output_files = get_output_files(updated_excel_workbooks)
            
            if output_files:
                st.success("🎉 Processing completed! Your files are ready for download.")
                
                # Download section
                st.markdown("---")
                st.subheader("📥 Download Processed Files")
                
                for output_file in output_files:
                    with open(output_file, "rb") as file:
                        st.download_button(
                            label=f"📊 Download {output_file.name}",
                            data=file.read(),
                            file_name=output_file.name,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
            else:
                st.warning("⚠️ No output files were generated. Please check the processing logs.")
    
    elif excel_file is not None or txt_files is not None:
        missing = []
        if excel_file is None:
            missing.append("Excel workbook")
        if txt_files is None:
            missing.append("Text data source")
        
        st.info(f"📋 Please upload the missing file(s): {', '.join(missing)}")
    
    else:
        st.info("📋 Please upload both an Excel workbook and a text data source file to begin processing.")
    
    # Instructions section
    with st.expander("ℹ️ How to use this application"):
        st.markdown("""
        ### Steps:
        1. **Upload Excel Workbook**: Select your Excel file (.xlsx or .xls) that needs processing
        2. **Upload Text Data Sources**: Select one or more text files (.txt) containing your data sources
        3. **Start Processing**: Click the "Start Processing" button to run the AI processing pipeline
        4. **Download Results**: Once processing is complete, download your processed Excel file(s)
        
        ### Requirements:
        - Excel files must be in .xlsx or .xls format
        - Text files must be in .txt format
        - You can upload multiple text files at once
        - Make sure your OpenAI API key is properly configured
        
        ### Processing Pipeline:
        The system will automatically:
        - Place your Excel file in the Input_Folder
        - Place all your text files in the data_sources folder
        - Run the complete AI processing workflow
        - Update the Excel file with processed data
        - Clean up input files after processing
        - Make the results available for download from the Output_folder
        - Remove downloaded files after successful download
        
        ### Automatic Cleanup:
        - After processing: All input Excel and text files are automatically removed
        - After download: Downloaded output files are automatically cleaned up
        - This keeps your directories clean and prevents file accumulation
        """)

if __name__ == "__main__":
    main()