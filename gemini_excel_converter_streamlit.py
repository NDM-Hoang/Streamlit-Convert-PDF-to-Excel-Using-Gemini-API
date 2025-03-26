import streamlit as st
import base64
import json
import os
import requests
import re
import tempfile
from pathlib import Path
from func_timeout import func_timeout, FunctionTimedOut

# Cài đặt trang
st.set_page_config(
    page_title="AI Excel Converter",
    page_icon="📊",
    layout="wide"
)

# Hàm lấy MIME type dựa trên phần mở rộng file
def get_mime_type(file_name):
    ext = os.path.splitext(file_name)[1].lower()
    mime_types = {
        ".pdf": "application/pdf",
        ".jpg": "image/jpeg",
        ".jpeg": "image/jpeg",
        ".png": "image/png"
    }
    
    if ext not in mime_types:
        st.error(f"Định dạng file không được hỗ trợ: {ext}")
        return None
    
    return mime_types[ext]

# Hàm xây dựng prompt cho Gemini API
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
            
            # Save file to buffer - DO NOT CHANGE THIS LINE
            wb.save(buffer)
            return buffer
        except Exception as e:
            print(f"Error: {{e}}")
            return None
    
    # This is how the function will be called - DO NOT CHANGE
    buffer = io.BytesIO()
    result = create_excel_report(buffer)
    buffer.seek(0)  # Reset buffer position
    ```
    
    Requirements:
    1. Extract ALL text/tables from the file
    2. Format with proper headings, alignment, borders
    3. DO NOT change function structure or parameters
    4. MUST save to buffer with wb.save(buffer)
    5. MUST return buffer at the end of the function
    6. Don't explain anything, just give code
    
    User instructions: {user_prompt}
    """

# Hàm gọi Gemini API
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

# Hàm trích xuất code từ phản hồi API
def extract_code(response):
    try:
        content_parts = response["candidates"][0]["content"]["parts"]
        full_text = ""
        for part in content_parts:
            if "text" in part:
                full_text += part["text"]
        
        # Tìm khối code Python
        code_pattern = r'``````'
        matches = re.findall(code_pattern, full_text, re.DOTALL)
        if matches:
            return matches[0].strip()
        
        # Thử tìm code trong khối ``````
        generic_pattern = r'``````'
        matches = re.findall(generic_pattern, full_text, re.DOTALL)
        if matches and len(matches) > 0:
            # Lấy khối code dài nhất (có thể là code chính)
            longest_match = max(matches, key=len)
            return longest_match.strip()
        
        # Làm sạch các dấu hiệu markdown còn sót
        clean_text = full_text.strip()
        clean_text = re.sub(r'^```python', '', clean_text, flags=re.MULTILINE)
        clean_text = re.sub(r'^```', '', clean_text, flags=re.MULTILINE)
        clean_text = re.sub(r'```$', '', clean_text, flags=re.MULTILINE)
        
        return clean_text
    
    except (KeyError, IndexError) as e:
        st.error(f"Không thể trích xuất code từ phản hồi: {str(e)}")
        return None


# Hàm thực thi code
def execute_code(code):
    try:
        # Sử dụng BytesIO cho file trong bộ nhớ
        from io import BytesIO
        buffer = BytesIO()
        
        # Chuẩn bị namespace 
        namespace = {
            'os': os,
            'io': __import__('io'),
            'openpyxl': __import__('openpyxl'),
            'BytesIO': BytesIO,
            'buffer': buffer
        }
        
        # Đảm bảo code có đoạn lưu vào buffer
        if 'buffer = io.BytesIO()' not in code and 'buffer = BytesIO()' not in code:
            modified_code = code.replace('def create_excel_report(buffer):', 
                                       'def create_excel_report(buffer=buffer):')
        else:
            modified_code = code
            
        # Thêm lệnh để đảm bảo buffer được trả về
        if 'return buffer' not in modified_code:
            lines = modified_code.split('\n')
            for i in range(len(lines)-1, -1, -1):
                if 'buffer.seek(0)' in lines[i]:
                    lines.insert(i+1, '    return buffer')
                    break
            modified_code = '\n'.join(lines)
            
        # Bắt output để debug
        import sys
        from io import StringIO
        old_stdout = sys.stdout
        captured_output = StringIO()
        sys.stdout = captured_output
        
        try:
            # Thực thi và lấy giá trị trả về nếu có
            result = exec(modified_code, namespace)
            
            # Kiểm tra buffer từ result hoặc namespace
            if 'buffer' in namespace and isinstance(namespace['buffer'], BytesIO):
                buffer = namespace['buffer']
                
            # Đảm bảo vị trí con trỏ ở đầu buffer
            buffer.seek(0)
            
            # Kiểm tra xem buffer có dữ liệu không
            if buffer.getbuffer().nbytes == 0:
                st.warning("Buffer trống sau khi thực thi code")
                # Thử tìm biến khác trong namespace có thể chứa dữ liệu Excel
                for var_name, var_value in namespace.items():
                    if isinstance(var_value, BytesIO) and var_value.getbuffer().nbytes > 0:
                        buffer = var_value
                        buffer.seek(0)
                        st.info(f"Đã tìm thấy buffer thay thế: {var_name}")
                        break
        finally:
            sys.stdout = old_stdout
        
        execution_log = captured_output.getvalue()
        st.code(execution_log, language="bash")  # Hiển thị log để debug
        
        if buffer.getbuffer().nbytes > 0:
            return True, buffer, "Excel file generated successfully!"
        else:
            return False, None, "Buffer trống sau khi thực thi code"
            
    except Exception as e:
        return False, None, f"Error executing code: {str(e)}"

def execute_code_with_timeout(code, timeout_seconds=30):
    try:
        # Sử dụng BytesIO cho file trong bộ nhớ
        from io import BytesIO
        buffer = BytesIO()
        
        # Chuẩn bị namespace
        namespace = {
            'os': os,
            'io': __import__('io'),
            'openpyxl': __import__('openpyxl'),
            'BytesIO': BytesIO,
            'buffer': buffer
        }
        
        # Định nghĩa hàm thực thi code
        def run_code():
            # Đảm bảo code sử dụng buffer
            modified_code = code
            if 'buffer = io.BytesIO()' not in modified_code and 'buffer = BytesIO()' not in modified_code:
                modified_code = modified_code.replace('def create_excel_report(buffer):', 
                                                    'def create_excel_report(buffer=buffer):')
            
            # Thêm lệnh return buffer nếu chưa có
            if 'return buffer' not in modified_code:
                lines = modified_code.split('\n')
                for i in range(len(lines)-1, -1, -1):
                    if 'buffer.seek(0)' in lines[i]:
                        lines.insert(i+1, '    return buffer')
                        break
                modified_code = '\n'.join(lines)
            
            # Thực thi code
            exec(modified_code, namespace)
            
            # Debug info
            st.write("Debug info:")
            st.json({
                "Platform": os.name,
                "Python version": sys.version,
                "Buffer size": buffer.getbuffer().nbytes if buffer else 0,
                "Namespace keys": list(namespace.keys())
            })
            
            # Lấy buffer từ namespace
            if 'buffer' in namespace and isinstance(namespace['buffer'], BytesIO):
                buffer = namespace['buffer']
                buffer.seek(0)
            
            return buffer
        
        # Chạy với timeout
        result_buffer = func_timeout(timeout_seconds, run_code)
        
        if result_buffer.getbuffer().nbytes > 0:
            return True, result_buffer, "Excel file generated successfully!"
        else:
            return False, None, "Buffer trống sau khi thực thi code"
            
    except FunctionTimedOut:
        return False, None, f"Thực thi code vượt quá thời gian giới hạn ({timeout_seconds} giây)"
    except Exception as e:
        return False, None, f"Error executing code: {str(e)}"

# Hàm lưu và tải API key
def save_api_key(api_key):
    try:
        config_dir = Path.home() / ".excel_converter"
        config_dir.mkdir(exist_ok=True)
        config_file = config_dir / "config.json"
        config = {"api_key": api_key}
        with open(config_file, "w") as f:
            json.dump(config, f)
    except Exception:
        pass

def load_api_key():
    try:
        config_file = Path.home() / ".excel_converter" / "config.json"
        if config_file.exists():
            with open(config_file, "r") as f:
                config = json.load(f)
                if "api_key" in config:
                    return config["api_key"]
    except Exception:
        pass
    return ""

# Khởi tạo session state
if 'api_key' not in st.session_state:
    st.session_state.api_key = load_api_key()
if 'generated_code' not in st.session_state:
    st.session_state.generated_code = ""
if 'excel_file_path' not in st.session_state:
    st.session_state.excel_file_path = ""
if 'execution_result' not in st.session_state:
    st.session_state.execution_result = ""
if 'file_processed' not in st.session_state:
    st.session_state.file_processed = False

# UI chính
st.title("AI Excel Converter")

# Phần API Key
with st.expander("Cài đặt API", expanded=True):
    api_key = st.text_input("Gemini API Key", 
                           value=st.session_state.api_key,
                           type="password",
                           help="Nhập Gemini API key của bạn")
    
    # Lưu API key khi thay đổi
    if api_key != st.session_state.api_key:
        st.session_state.api_key = api_key
        save_api_key(api_key)

# Phần chọn file
st.subheader("Chọn file đầu vào")
uploaded_file = st.file_uploader("Chọn PDF/Ảnh", type=["pdf", "png", "jpg", "jpeg"])

# Prompt
st.subheader("Yêu cầu xử lý")
prompt_text = st.text_area(
    "Prompt",  # Thêm label
    value="Read file then create code to create Excel file with full data from image without editing or deleting anything, full text.",
    height=100,
    label_visibility="collapsed"  # Ẩn label nhưng vẫn tuân thủ accessibility
)

# Thêm vào UI sau phần prompt
st.subheader("Cài đặt thực thi")
timeout_seconds = st.slider("Thời gian timeout (giây)", min_value=5, max_value=120, value=30, step=5)

# Thanh tiến trình và trạng thái
progress_placeholder = st.empty()
status_placeholder = st.empty()

# Khu vực hiển thị code
st.subheader("Code sinh ra")
code_area = st.text_area(
    "Generated Code",  # Thêm label
    value=st.session_state.generated_code, 
    height=300,
    label_visibility="collapsed"  # Ẩn label
)

if code_area != st.session_state.generated_code and code_area.strip() != "":
    st.session_state.generated_code = code_area

# Nút chức năng
col1, col2, col3, col4 = st.columns(4)

with col1:
    run_prompt_button = st.button("Chạy Prompt", use_container_width=True)

with col2:
    run_code_button = st.button("Chạy Code", 
                              disabled=not st.session_state.generated_code, 
                              use_container_width=True)

with col3:
    retry_prompt_button = st.button("Chạy lại Prompt", use_container_width=True)

with col4:
    reset_button = st.button("Reset", use_container_width=True)

# Xử lý khi nhấn nút
if run_prompt_button or retry_prompt_button:
    if not api_key:
        st.error("Vui lòng nhập API Key")
    elif not uploaded_file:
        st.error("Vui lòng chọn file đầu vào")
    else:
        # Hiển thị thanh tiến trình
        progress_bar = progress_placeholder.progress(0)
        
        try:
            # Lưu file tạm thời
            with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(uploaded_file.name)[1]) as tmp:
                tmp.write(uploaded_file.getbuffer())
                temp_file_path = tmp.name
            
            status_placeholder.info("Đang xử lý file...")
            progress_bar.progress(20)
            
            # Đọc và mã hóa file
            with open(temp_file_path, "rb") as f:
                file_data = base64.b64encode(f.read()).decode("utf-8")
            
            mime_type = get_mime_type(uploaded_file.name)
            if not mime_type:
                progress_placeholder.empty()
                status_placeholder.error("Định dạng file không được hỗ trợ")
                os.unlink(temp_file_path)
                st.stop()
            
            status_placeholder.info("Đang tạo prompt...")
            progress_bar.progress(30)
            
            # Xây dựng prompt mới không cần đường dẫn đầu ra
            prompt = build_prompt(prompt_text)
            
            status_placeholder.info("Đang gửi yêu cầu đến Gemini API...")
            progress_bar.progress(40)
            
            # Gọi API
            response = call_gemini_api(api_key, prompt, file_data, mime_type)
            if not response:
                progress_placeholder.empty()
                status_placeholder.error("Lỗi khi gọi Gemini API")
                os.unlink(temp_file_path)
                st.stop()
            
            status_placeholder.info("Đang trích xuất code...")
            progress_bar.progress(80)
            
            # Trích xuất code
            generated_code = extract_code(response)
            if not generated_code:
                progress_placeholder.empty()
                status_placeholder.error("Không thể trích xuất code từ phản hồi API")
                os.unlink(temp_file_path)
                st.stop()
            
            # Lưu vào session state
            st.session_state.generated_code = generated_code
            st.session_state.file_processed = True
            
            # Xóa file tạm
            os.unlink(temp_file_path)
            
            progress_bar.progress(100)
            status_placeholder.success("Đã tạo code thành công")
            
            # Buộc chạy lại để cập nhật giao diện
            st.rerun()
            
        except Exception as e:
            progress_placeholder.empty()
            status_placeholder.error(f"Lỗi: {str(e)}")

if run_code_button:
    if not st.session_state.generated_code:
        st.error("Không có code để thực thi")
    else:
        # Hiển thị thanh tiến trình
        progress_bar = progress_placeholder.progress(0)
        status_placeholder.info("Đang thực thi code...")
        
        # Cập nhật tiến trình theo từng bước
        for i in range(1, 5):
            progress_bar.progress(i * 20)
        
        try:
            # Chạy code với timeout
            success, excel_buffer, message = execute_code_with_timeout(
                st.session_state.generated_code, 
                timeout_seconds=timeout_seconds
            )
            
            progress_bar.progress(100)
            
            if success:
                # Tạo tên file theo tên file đầu vào (nếu có)
                if uploaded_file:
                    base_name = os.path.splitext(uploaded_file.name)[0]  # Lấy chỉ phần tên file
                    excel_file_name = f"{base_name}.xlsx"
                else:
                    excel_file_name = "converted_data.xlsx"
                
                buffer_size = excel_buffer.getbuffer().nbytes
                if buffer_size > 0:
                    # Hiển thị nút tải xuống
                    st.download_button(
                        label="Tải xuống file Excel",
                        data=excel_buffer,
                        file_name=excel_file_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                    status_placeholder.success(f"Excel file đã được tạo ({buffer_size} bytes). Nhấn nút để tải xuống.")
                else:
                    status_placeholder.error("File Excel trống (0 bytes). Có lỗi trong quá trình tạo file.")
            else:
                status_placeholder.error(f"Thực thi code thất bại: {message}")
                
        except Exception as e:
            progress_bar.progress(100)
            status_placeholder.error(f"Lỗi không mong đợi: {str(e)}")

if reset_button:
    # Reset session state
    st.session_state.generated_code = ""
    st.session_state.excel_file_path = ""
    st.session_state.execution_result = ""
    st.session_state.file_processed = False
    progress_placeholder.empty()
    status_placeholder.info("Đã reset các trường nhập liệu")
    
    # Buộc chạy lại để cập nhật giao diện
    st.rerun()

# Footer
st.markdown("---")
st.caption("AI Excel Converter | Powered by Gemini API")
