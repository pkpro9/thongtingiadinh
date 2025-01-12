import streamlit as st
from datetime import datetime
import pytz  # For timezone handling
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

def normalize_text(text):
    """Convert text to lowercase and remove accents."""
    text = unicodedata.normalize('NFD', text)
    text = text.encode('ascii', 'ignore').decode("utf-8")
    text = re.sub(r'[^\w\s]', '', text)
    return text.lower().replace(" ", "_")

def extract_year(date_str):
    """Extract the year from a date string in dd/mm/yyyy HH:MM format."""
    try:
        return datetime.strptime(date_str, "%d/%m/%Y %H:%M").year
    except ValueError:
        return ""

def save_to_google_sheet(date, document_name, hyperlink, category, year):
    """Save data to Google Sheet with corrected hyperlink format."""
    client = gspread.authorize(CREDENTIALS)
    sheet = client.open_by_key(SPREADSHEET_ID).sheet1
    # Append data as a row to the sheet
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
st.title("Upload Hồ sơ và Lưu trữ Google Drive/Sheet")

# Input fields
date = st.text_input("Ngày", get_vietnam_time())  # Default to Vietnam time
document_name = st.text_input("Tên tài liệu/hồ sơ")

# Dropdown for "Loại"
category = st.selectbox(
    "Loại",
    ["Học tập Gia Lộc", "Học Tập Gia Phú", "Giấy tờ HC"]
)

# File upload
uploaded_file = st.file_uploader("Đính kèm tài liệu/hồ sơ", type=["pdf", "docx", "xlsx", "png", "jpg", "jpeg"])

if st.button("Lưu"):
    if not document_name or not uploaded_file or not category:
        st.error("Vui lòng nhập đầy đủ thông tin và tải lên file!")
    else:
        # Extract year from the date
        year = extract_year(date)
        
        # Normalize and save the file
        normalized_name = normalize_text(document_name) + os.path.splitext(uploaded_file.name)[1]
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
