import streamlit as st
import pandas as pd
import io
import re
import time
import os
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/presentations',
    'https://www.googleapis.com/auth/drive'
]

def extract_id(url, type='sheet'):
    if type == 'sheet':
        match = re.search(r'/spreadsheets/d/([a-zA-Z0-9-_]+)', url)
    elif type == 'slides':
        match = re.search(r'/presentation/d/([a-zA-Z0-9-_]+)', url)
    elif type == 'folder':
        match = re.search(r'/folders/([a-zA-Z0-9-_]+)', url)
    return match.group(1) if match else url.split('/')[-1].split('?')[0]

def col_to_letter(n):
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string

st.title("📄 Professional PDF Generator")
st.write("Now with Dynamic Filenames and True Auto-Resume.")

# --- UI SECTION 1: THE INPUTS ---
st.header("1. File Links")
sheet_url = st.text_input("Google Sheet URL")
slide_url = st.text_input("Google Slide Template URL")
folder_url = st.text_input("Google Drive Output Folder URL")

if st.button("Connect Google Account & Load Sheet"):
    if not os.path.exists("client_secret.json"):
        st.error("❌ client_secret.json not found!")
    elif sheet_url and slide_url and folder_url:
        with st.spinner("Connecting..."):
            try:
                flow = InstalledAppFlow.from_client_secrets_file('client_secret.json', SCOPES)
                creds = flow.run_local_server(port=8080, prompt='consent') 
                
                st.session_state['sheets_service'] = build('sheets', 'v4', credentials=creds)
                st.session_state['slides_service'] = build('slides', 'v1', credentials=creds)
                st.session_state['drive_service'] = build('drive', 'v3', credentials=creds)
                
                st.session_state['sheet_id'] = extract_id(sheet_url, 'sheet')
                st.session_state['slide_id'] = extract_id(slide_url, 'slides')
                st.session_state['folder_id'] = extract_id(folder_url, 'folder')
                
                st.success("✅ Connected successfully! Proceed to Step 2.")
            except Exception as e:
                st.error(f"Error: {e}")

# --- UI SECTION 2: THE CONFIGURATION ---
if 'sheet_id' in st.session_state:
    st.header("2. Smart Configuration")
    
    st.info("💡 **Naming Rule:** Type your desired PDF name below. Use `<<Column Name>>` to insert data from your sheet. You can combine multiple columns and manual text.")
    
    # NEW: Dynamic Flexible Filename Template
    filename_template = st.text_input(
        "PDF Filename Template:", 
        value="<<School ID>> - <<School Name>> - Partnership Letter"
    )

    st.warning("⚠️ **Important:** Make sure your Google Sheet has exactly two columns named **Status** and **PDF Link** to track progress.")

    if st.button("Start Generating / Resume Processing"):
        
        # 1. LIVE FETCH (Fixes the resume bug)
        status_text = st.empty()
        status_text.text("Fetching latest data from Google Sheets...")
        
        result = st.session_state['sheets_service'].spreadsheets().values().get(
            spreadsheetId=st.session_state['sheet_id'], range='Sheet1'
        ).execute()
        values = result.get('values', [])
        
        if not values:
            st.error("Your Google Sheet is empty!")
            st.stop()
            
        headers = values[0]
        
        if 'Status' not in headers or 'PDF Link' not in headers:
            st.error("❌ Could not find 'Status' or 'PDF Link' columns. Please add them to your sheet header.")
            st.stop()

        # Fixes the "Invisible Cell" bug by forcing all rows to be the same length
        max_cols = len(headers)
        padded_rows = []
        for r in values[1:]:
            padded_rows.append(r + [''] * (max_cols - len(r)))
            
        df = pd.DataFrame(padded_rows, columns=headers)

        status_idx = headers.index('Status') + 1
        link_idx = headers.index('PDF Link') + 1
        status_letter = col_to_letter(status_idx)
        link_letter = col_to_letter(link_idx)

        total_rows = len(df)
        progress_bar = st.progress(0)
        
        # 2. THE GENERATION LOOP
        for index, row in df.iterrows():
            
            # TRUE RESUME LOGIC
            current_status = str(row.get('Status', '')).strip().lower()
            if 'success' in current_status:
                progress_bar.progress((index + 1) / total_rows)
                continue

            try:
                # DYNAMIC FILENAME CREATION
                filename = filename_template
                for h in headers:
                    val = str(row.get(h, "")).strip()
                    # Replace the tag with actual data
                    filename = filename.replace(f"<<{h}>>", val)
                
                # Clean up the name and add .pdf
                filename = filename.strip()
                if not filename.lower().endswith(".pdf"):
                    filename += ".pdf"
                
                status_text.text(f"Processing ({index+1}/{total_rows}): {filename}")

                # Copy Template
                copy_body = {'name': filename.replace('.pdf', '')}
                copied_file = st.session_state['drive_service'].files().copy(
                    fileId=st.session_state['slide_id'], body=copy_body
                ).execute()
                temp_slide_id = copied_file['id']

                # Replace Text in Slide
                requests = []
                for h in headers:
                    val = str(row.get(h, "")).strip()
                    requests.append({
                        'replaceAllText': {
                            'containsText': {'text': f"<<{h}>>", 'matchCase': False},
                            'replaceText': val
                        }
                    })
                
                if requests:
                    st.session_state['slides_service'].presentations().batchUpdate(
                        presentationId=temp_slide_id, body={'requests': requests}
                    ).execute()

                # Export to PDF
                request = st.session_state['drive_service'].files().export_media(
                    fileId=temp_slide_id, mimeType='application/pdf'
                )
                fh = io.BytesIO()
                downloader = MediaIoBaseDownload(fh, request)
                done = False
                while not done:
                    _, done = downloader.next_chunk()

                # Upload PDF to Drive Folder
                fh.seek(0)
                file_metadata = {'name': filename, 'parents': [st.session_state['folder_id']]}
                media = MediaIoBaseUpload(fh, mimetype='application/pdf', resumable=True)
                uploaded_pdf = st.session_state['drive_service'].files().create(
                    body=file_metadata, media_body=media, fields='webViewLink'
                ).execute()
                pdf_link = uploaded_pdf.get('webViewLink')

                # Delete Temp Slide
                st.session_state['drive_service'].files().delete(fileId=temp_slide_id).execute()

                # UPDATE GOOGLE SHEET LIVE
                sheet_row = index + 2
                
                st.session_state['sheets_service'].spreadsheets().values().update(
                    spreadsheetId=st.session_state['sheet_id'],
                    range=f"Sheet1!{status_letter}{sheet_row}",
                    valueInputOption="USER_ENTERED",
                    body={"values": [["Success"]]}
                ).execute()

                st.session_state['sheets_service'].spreadsheets().values().update(
                    spreadsheetId=st.session_state['sheet_id'],
                    range=f"Sheet1!{link_letter}{sheet_row}",
                    valueInputOption="USER_ENTERED",
                    body={"values": [[pdf_link]]}
                ).execute()

                progress_bar.progress((index + 1) / total_rows)
                time.sleep(1) 

            except Exception as e:
                st.error(f"Error on row {index+1}: {e}")
                try:
                    st.session_state['sheets_service'].spreadsheets().values().update(
                        spreadsheetId=st.session_state['sheet_id'],
                        range=f"Sheet1!{status_letter}{index+2}",
                        valueInputOption="USER_ENTERED",
                        body={"values": [[f"Failed: {str(e)[:50]}"]]}
                    ).execute()
                except: pass

        status_text.text("Done!")
        st.success("🎉 All pending rows processed!")
        st.balloons()