import streamlit as st
import base64
import json
import requests
import re
import tempfile
import os
from io import BytesIO

# Page configuration
st.set_page_config(
    page_title="AI Excel Converter",
    page_icon="📊",
    layout="wide"
)

# Get MIME type based on file extension
def get_mime_type(file_name):
    ext = os.path.splitext(file_name)[1].lower()
    mime_types = {
        ".pdf": "application/pdf",
        ".jpg": "image/jpeg",
        ".jpeg": "image/jpeg",
        ".png": "image/png"
    }
    if ext not in mime_types:
        st.error(f"Unsupported file format: {ext}")
        return None
    return mime_types[ext]

# Build prompt for Gemini API
def build_prompt(user_prompt):
    return f"""
    I need Python code that extracts all text data from the attached file and creates an Excel file following this EXACT code structure:

    ```
    import openpyxl
    from openpyxl.styles import Font, Alignment, Border, Side
    import io

    def create_excel_report(buffer):
        try:
            # Create workbook
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Extracted Data"
            
            # Define styles
            bold_font = Font(bold=True)
            center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
            left_align = Alignment(horizontal='left', vertical='center', wrap_text=True)
            thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                                top=Side(style='thin'), bottom=Side(style='thin'))
            
            # YOUR CODE HERE: Extract and format all content from the source file
            
            # Save file to buffer
            wb.save(buffer)
            print("Successfully created Excel file in memory buffer")
            return True
        except Exception as e:
            print(f"Error: {{e}}")
            return False

    # Execution
    buffer = io.BytesIO()
    create_excel_report(buffer)
    buffer.seek(0) # Reset buffer position to beginning
    ```

    Requirements:
    1. Extract ALL text/tables from the file
    2. Format with proper headings, alignment, borders
    3. DO NOT change function structure
    4. MUST save to buffer with wb.save(buffer)
    5. Ensure buffer has data by verifying it after save
    6. Return buffer at the end of the function
    7. Don't explain anything, just give code

    User instructions: {user_prompt}
    """

# Call Gemini API
def call_gemini_api(api_key, prompt, file_data, mime_type):
    headers = {
        "Content-Type": "application/json"
    }
    
    payload = {
        "contents": [{
            "parts": [
                {"text": prompt},
                {
                    "inline_data": {
                        "mime_type": mime_type,
                        "data": file_data
                    }
                }
            ]
        }],
        "generationConfig": {
            "temperature": 0.3,
            "topP": 0.95,
            "maxOutputTokens": 8192
        }
    }
    
    response = requests.post(
        f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-pro-exp-03-25:generateContent?key={api_key}",
        headers=headers,
        json=payload
    )
    
    if response.status_code != 200:
        st.error(f"API Error: {response.text}")
        return None
    
    return response.json()

# Extract code from API response
def extract_code(response):
    try:
        content_parts = response["candidates"][0]["content"]["parts"]
        full_text = ""
        for part in content_parts:
            if "text" in part:
                full_text += part["text"]
        
        # Look for Python code blocks
        code_pattern = r'``````'
        matches = re.findall(code_pattern, full_text, re.DOTALL)
        if matches:
            return matches[0].strip()
        
        # If Python blocks not found, look for generic code blocks
        generic_pattern = r'``````'
        matches = re.findall(generic_pattern, full_text, re.DOTALL)
        if matches:
            return matches[0].strip()
        
        # Clean up any remaining markdown
        clean_text = full_text.strip()
        clean_text = re.sub(r'^```python', '', clean_text, flags=re.MULTILINE)
        clean_text = re.sub(r'^```', '', clean_text, flags=re.MULTILINE)
        clean_text = re.sub(r'```$', '', clean_text, flags=re.MULTILINE)
        
        return clean_text
    
    except (KeyError, IndexError) as e:
        st.error(f"Unable to extract code from response: {str(e)}")
        return None

# Execute generated code
def execute_code(code):
    try:
        # Use BytesIO for in-memory file
        buffer = BytesIO()
        
        # Prepare namespace
        namespace = {
            'os': os,
            'io': __import__('io'),
            'openpyxl': __import__('openpyxl'),
            'BytesIO': BytesIO,
            'buffer': buffer
        }
        
        # Ensure code saves to buffer
        if 'buffer = io.BytesIO()' not in code and 'buffer = BytesIO()' not in code:
            modified_code = code.replace('def create_excel_report(buffer):',
                                        'def create_excel_report(buffer=buffer):')
        else:
            modified_code = code
        
        # Add command to ensure buffer is returned
        if 'return buffer' not in modified_code:
            lines = modified_code.split('\n')
            for i in range(len(lines)-1, -1, -1):
                if 'buffer.seek(0)' in lines[i]:
                    lines.insert(i+1, '    return buffer')
                    break
            modified_code = '\n'.join(lines)
        
        # Capture output for debugging
        import sys
        from io import StringIO
        old_stdout = sys.stdout
        captured_output = StringIO()
        sys.stdout = captured_output
        
        try:
            # Execute and get return value if any
            exec(modified_code, namespace)
            
            # Check buffer from namespace
            if 'buffer' in namespace and isinstance(namespace['buffer'], BytesIO):
                buffer = namespace['buffer']
                # Ensure pointer is at start of buffer
                buffer.seek(0)
                
                # Check if buffer has data
                if buffer.getbuffer().nbytes == 0:
                    st.warning("Buffer empty after code execution")
                    # Try to find other variables in namespace that might contain Excel data
                    for var_name, var_value in namespace.items():
                        if isinstance(var_value, BytesIO) and var_value.getbuffer().nbytes > 0:
                            buffer = var_value
                            buffer.seek(0)
                            st.info(f"Found alternative buffer: {var_name}")
                            break
        finally:
            sys.stdout = old_stdout
            execution_log = captured_output.getvalue()
            st.code(execution_log, language="bash")  # Show log for debugging
        
        if buffer.getbuffer().nbytes > 0:
            return True, buffer, "Excel file generated successfully!"
        else:
            return False, None, "Empty buffer after code execution"
            
    except Exception as e:
        return False, None, f"Error executing code: {str(e)}"

# Initialize session state
if 'api_key' not in st.session_state:
    st.session_state.api_key = ""
if 'generated_code' not in st.session_state:
    st.session_state.generated_code = ""
if 'file_processed' not in st.session_state:
    st.session_state.file_processed = False

# Main UI
st.title("AI Excel Converter")

# API Key section
with st.expander("API Settings", expanded=True):
    api_key = st.text_input("Gemini API Key",
                           value=st.session_state.api_key,
                           type="password",
                           help="Enter your Gemini API key")
    
    # Save API key when changed
    if api_key != st.session_state.api_key:
        st.session_state.api_key = api_key

# File selection
st.subheader("Select Input File")
uploaded_file = st.file_uploader("Choose PDF/Image", type=["pdf", "png", "jpg", "jpeg"])

# Prompt
st.subheader("Processing Request")
prompt_text = st.text_area("",
                          value="Read file then create code to create Excel file with full data from image without editing or deleting anything, full text.",
                          height=100)

# Progress and status indicators
progress_placeholder = st.empty()
status_placeholder = st.empty()

# Code display area
st.subheader("Generated Code")
code_area = st.text_area("", value=st.session_state.generated_code, height=300)
if code_area != st.session_state.generated_code and code_area.strip() != "":
    st.session_state.generated_code = code_area

# Action buttons (simplified)
col1, col2, col3 = st.columns(3)
with col1:
    run_prompt_button = st.button("Generate Code", use_container_width=True)
with col2:
    run_code_button = st.button("Execute Code",
                               disabled=not st.session_state.generated_code,
                               use_container_width=True)
with col3:
    reset_button = st.button("Reset", use_container_width=True)

# Button handlers
if run_prompt_button:
    if not api_key:
        st.error("Please enter an API Key")
    elif not uploaded_file:
        st.error("Please select an input file")
    else:
        # Show progress bar
        progress_bar = progress_placeholder.progress(0)
        
        try:
            # Save temporary file
            with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(uploaded_file.name)[1]) as tmp:
                tmp.write(uploaded_file.getbuffer())
                temp_file_path = tmp.name
            
            status_placeholder.info("Processing file...")
            progress_bar.progress(20)
            
            # Read and encode file
            with open(temp_file_path, "rb") as f:
                file_data = base64.b64encode(f.read()).decode("utf-8")
            
            mime_type = get_mime_type(uploaded_file.name)
            if not mime_type:
                progress_placeholder.empty()
                status_placeholder.error("Unsupported file format")
                os.unlink(temp_file_path)
                st.stop()
            
            status_placeholder.info("Creating prompt...")
            progress_bar.progress(30)
            
            # Build prompt
            prompt = build_prompt(prompt_text)
            
            status_placeholder.info("Sending request to Gemini API...")
            progress_bar.progress(40)
            
            # Call API
            response = call_gemini_api(api_key, prompt, file_data, mime_type)
            if not response:
                progress_placeholder.empty()
                status_placeholder.error("Error calling Gemini API")
                os.unlink(temp_file_path)
                st.stop()
            
            status_placeholder.info("Extracting code...")
            progress_bar.progress(80)
            
            # Extract code
            generated_code = extract_code(response)
            if not generated_code:
                progress_placeholder.empty()
                status_placeholder.error("Unable to extract code from API response")
                os.unlink(temp_file_path)
                st.stop()
            
            # Save to session state
            st.session_state.generated_code = generated_code
            st.session_state.file_processed = True
            
            # Delete temp file
            os.unlink(temp_file_path)
            
            progress_bar.progress(100)
            status_placeholder.success("Code generated successfully")
            
            # Force UI update
            st.rerun()
            
        except Exception as e:
            progress_placeholder.empty()
            status_placeholder.error(f"Error: {str(e)}")

if run_code_button:
    if not st.session_state.generated_code:
        st.error("No code to execute")
    else:
        # Show progress bar
        progress_bar = progress_placeholder.progress(0)
        status_placeholder.info("Executing code...")
        
        # Update progress in steps
        for i in range(1, 5):
            progress_bar.progress(i * 20)
        
        # Run code and get buffer instead of saving file
        success, excel_buffer, message = execute_code(st.session_state.generated_code)
        
        progress_bar.progress(100)
        
        if success:
            # Save buffer in session state for later use
            st.session_state.excel_buffer = excel_buffer
            
            # Create filename based on input file (if any)
            if uploaded_file:
                base_name = os.path.splitext(uploaded_file.name)[0]  # Get just the filename
                excel_file_name = f"{base_name}.xlsx"
            else:
                excel_file_name = "converted_data.xlsx"
            
            buffer_size = excel_buffer.getbuffer().nbytes
            
            if buffer_size > 0:
                # Show download button
                st.download_button(
                    label="Download Excel File",
                    data=excel_buffer,
                    file_name=excel_file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                status_placeholder.success(f"Excel file created ({buffer_size} bytes). Click button to download.")
            else:
                status_placeholder.error("Excel file is empty (0 bytes). Error in file creation process.")
        else:
            status_placeholder.error(message)

if reset_button:
    # Reset session state
    st.session_state.generated_code = ""
    st.session_state.file_processed = False
    if 'excel_buffer' in st.session_state:
        del st.session_state.excel_buffer
    
    progress_placeholder.empty()
    status_placeholder.info("All fields reset")
    
    # Force UI update
    st.rerun()

# Footer
st.markdown("---")
st.caption("AI Excel Converter | Powered by Gemini API")
