import streamlit as st
from datetime import datetime
import pytz
import unicodedata
import re
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
import os

# Google API configuration
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
CREDENTIALS = Credentials.from_service_account_info(st.secrets["google"], scopes=SCOPES)

# Google Drive & Sheet mapping
SHEET_MAPPING = {
    "Hồ sơ gia đình": ("TTHC_GIA_DINH", "1MD1jvHEXX1CbHN-ISydzTa6t69xRNfHY"),
    "Quyết định-Hợp đồng-Khác": ("QD_HD_KHAC", "1zPfNupTE8P5XnXqhRaoZ3k5lHwwC5WFa"),
    "Bằng cấp-Chứng chỉ": ("BANG_CAP", "1b0KkS894xUH_-y5UqScyvjiUkgcdOUgZ"),
    "CME": ("CME", "1z2J5raE8UprnDd5-s8Jcoy0ucIh2Wnp_"),
    "Khen Thưởng": ("KHEN_THUONG", "14s0HEXN10gLVXw7mT0-7fkeWciZDxGNY"),
}

# Dropdown options for each menu
CATEGORY_OPTIONS = {
    "Hồ sơ gia đình": ["Giấy tờ HC", "Học tập Gia Lộc", "Học tập Gia Phú"],
    "Quyết định-Hợp đồng-Khác": ["Quyết định", "Hợp đồng", "Chứng nhận", "Thông báo", "Cam kết", "Khác"],
    "Bằng cấp-Chứng chỉ": ["Bằng TN", "Chứng chỉ", "Chứng nhận", "Công nhận", "Bảng điểm", "Khác"],
    "CME": ["CME", "Chứng chỉ", "Chứng nhận", "Khác"],
    "Khen Thưởng": ["Giấy khen", "Bằng khen", "Danh hiệu", "Biểu dương", "Khác"],
}

SPREADSHEET_ID = "1H-7ycEtf8lFQqLCEbeLkRS61rBY3XZWtTkQuEV7GATY"

def normalize_text_to_title(text):
    """Convert text to lowercase, remove accents, and capitalize each word."""
    text = unicodedata.normalize('NFD', text)
    text = text.encode('ascii', 'ignore').decode("utf-8")
    text = re.sub(r'[^\w\s]', '', text)
    return text.title().replace(" ", "_")  # Capitalize each word and replace spaces with underscores

def get_next_stt(sheet_name):
    """Retrieve the next available STT (ID) from the Google Sheet."""
    client = gspread.authorize(CREDENTIALS)
    sheet = client.open_by_key(SPREADSHEET_ID).worksheet(sheet_name)
    data = sheet.get_all_values()  
    if len(data) > 2:  
        return len(data) - 2 + 1  
    return 1  

def save_to_google_sheet(sheet_name, stt, date, document_name, hyperlink, category, year, issuing_place):
    """Save data to the correct sheet in Google Sheets."""
    client = gspread.authorize(CREDENTIALS)
    sheet = client.open_by_key(SPREADSHEET_ID).worksheet(sheet_name)
    sheet.append_row(
        [stt, date, f'=HYPERLINK("{hyperlink}";"{stt}. {document_name}")', year, category, issuing_place],
        value_input_option="USER_ENTERED"
    )

def upload_to_google_drive(folder_id, file, file_name):
    """Upload file to the correct Google Drive folder and return file ID."""
    service = build("drive", "v3", credentials=CREDENTIALS)
    media = MediaFileUpload(file, resumable=True)
    file_metadata = {"name": file_name, "parents": [folder_id]}
    uploaded_file = service.files().create(body=file_metadata, media_body=media, fields="id").execute()
    return uploaded_file.get("id")

# Get current datetime in Vietnam timezone
def get_vietnam_time():
    vietnam_tz = pytz.timezone("Asia/Ho_Chi_Minh")
    return datetime.now(vietnam_tz).strftime("%d/%m/%Y %H:%M")

# Streamlit UI
st.title("Quản lý TL-HS gia đình")

# Menu chọn loại hồ sơ
category_type = st.sidebar.selectbox(
    "Chọn loại hồ sơ",
    list(SHEET_MAPPING.keys())
)

# Lấy sheet name và folder tương ứng
sheet_name, folder_id = SHEET_MAPPING[category_type]

# Get next STT
stt = get_next_stt(sheet_name)

# Input fields
date = st.text_input("Ngày", get_vietnam_time())  
document_name = st.text_input("Tên tài liệu/hồ sơ")

# Dropdown for "Loại"
category = st.selectbox("Loại", CATEGORY_OPTIONS[category_type])

# Year input for "Năm TL/HS" using selectbox
current_year = datetime.now().year
years = [str(y) for y in range(1900, current_year + 1)]
year = st.selectbox("Năm TL/HS", reversed(years))

# Text input for "Nơi ban hành"
issuing_place = st.text_input("Nơi ban hành")

# File upload
uploaded_file = st.file_uploader("Đính kèm tài liệu/hồ sơ", type=["pdf", "docx", "xlsx", "png", "jpg", "jpeg"])

if st.button("Lưu"):
    if not document_name or not uploaded_file or not category or not year or not issuing_place:
        st.error("Vui lòng nhập đầy đủ thông tin và tải lên file!")
    else:
        # Normalize and save the file with Title Case
        normalized_name = normalize_text_to_title(document_name) + os.path.splitext(uploaded_file.name)[1]
        file_with_stt = f"{stt}. {normalized_name}"

        with open(file_with_stt, "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        # Upload file to correct Google Drive folder
        file_id = upload_to_google_drive(folder_id, file_with_stt, file_with_stt)
        
        # Generate hyperlink
        file_link = f"https://drive.google.com/file/d/{file_id}/view"
        
        # Save data to the correct Google Sheet
        save_to_google_sheet(sheet_name, stt, date, document_name, file_link, category, year, issuing_place)
        
        # Cleanup
        os.remove(file_with_stt)
        
        st.success(f"Dữ liệu đã được lưu vào '{sheet_name}' thành công với STT: {stt}!")
        st.info(f"File đã được tải lên Google Drive với tên: {file_with_stt}")
