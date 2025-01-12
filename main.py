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

# Google Drive and Sheet setup
DRIVE_FOLDER_ID = "1MD1jvHEXX1CbHN-ISydzTa6t69xRNfHY"
SPREADSHEET_ID = "1H-7ycEtf8lFQqLCEbeLkRS61rBY3XZWtTkQuEV7GATY"

def normalize_text_to_title(text):
    """Convert text to lowercase, remove accents, and capitalize each word."""
    text = unicodedata.normalize('NFD', text)
    text = text.encode('ascii', 'ignore').decode("utf-8")
    text = re.sub(r'[^\w\s]', '', text)
    return text.title().replace(" ", "_")  # Capitalize each word and replace spaces with underscores

def save_to_google_sheet(date, document_name, hyperlink, category, year):
    """Save data to Google Sheet."""
    client = gspread.authorize(CREDENTIALS)
    sheet = client.open_by_key(SPREADSHEET_ID).sheet1
    # Append data to the sheet
    sheet.append_row(
        [date, f'=HYPERLINK("{hyperlink}";"{document_name}")', year, category],
        value_input_option="USER_ENTERED"
    )

def upload_to_google_drive(file, file_name):
    """Upload file to Google Drive and return file ID."""
    service = build("drive", "v3", credentials=CREDENTIALS)
    media = MediaFileUpload(file, resumable=True)
    file_metadata = {"name": file_name, "parents": [DRIVE_FOLDER_ID]}
    uploaded_file = service.files().create(body=file_metadata, media_body=media, fields="id").execute()
    return uploaded_file.get("id")

# Get current datetime in Vietnam timezone
def get_vietnam_time():
    vietnam_tz = pytz.timezone("Asia/Ho_Chi_Minh")
    return datetime.now(vietnam_tz).strftime("%d/%m/%Y %H:%M")

# Streamlit UI
st.title("Quản lý TL-HS gia đình")

# Input fields
date = st.text_input("Ngày", get_vietnam_time())  # Default to Vietnam time
document_name = st.text_input("Tên tài liệu/hồ sơ")

# Dropdown for "Loại"
category = st.selectbox(
    "Loại",
    ["Giấy tờ HC", "Học tập Gia Lộc", "Học tập Gia Phú"]
)

# Year input for "Năm TL/HS" using selectbox
current_year = datetime.now().year
years = [str(y) for y in range(1900, current_year + 1)]  # Generate a list of years from 1900 to current year
year = st.selectbox("Năm TL/HS", reversed(years))  # Show years in descending order

# File upload
uploaded_file = st.file_uploader("Đính kèm tài liệu/hồ sơ", type=["pdf", "docx", "xlsx", "png", "jpg", "jpeg"])

if st.button("Lưu"):
    if not document_name or not uploaded_file or not category or not year:
        st.error("Vui lòng nhập đầy đủ thông tin và tải lên file!")
    else:
        # Normalize and save the file with Title Case
        normalized_name = normalize_text_to_title(document_name) + os.path.splitext(uploaded_file.name)[1]
        with open(normalized_name, "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        # Upload file to Google Drive
        file_id = upload_to_google_drive(normalized_name, normalized_name)
        
        # Generate hyperlink
        file_link = f"https://drive.google.com/file/d/{file_id}/view"
        
        # Save data to Google Sheet
        save_to_google_sheet(date, document_name, file_link, category, year)
        
        # Cleanup
        os.remove(normalized_name)
        
        st.success("Dữ liệu đã được lưu thành công!")
        st.info(f"File đã được tải lên Google Drive với liên kết: {file_link}")
