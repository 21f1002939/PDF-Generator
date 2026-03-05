import streamlit as st
import pandas as pd
import io
import re
import time
import json
from google_auth_oauthlib.flow import Flow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

# --- 1. CONFIGURATION & SCOPES ---
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/presentations',
    'https://www.googleapis.com/auth/drive'
]

# Load secrets from Streamlit Vault
if "google_secret" in st.secrets:
    client_config = json.loads(st.secrets["google_secret"])
else:
    st.error("❌ Missing 'google_secret' in Streamlit Secrets! Please add it in Advanced Settings.")
    st.stop()

# --- 2. HELPER FUNCTIONS ---
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

# --- 3. CLOUD AUTHENTICATION ---
st.title("📄 Professional PDF Generator")

# CHANGE THIS to your actual Streamlit URL (e.g., https://ashish-pdf.streamlit.app)
# This MUST match what you entered in Google Cloud Console exactly.
redirect_uri = "https://auto-pdf-generator.streamlit.app/" 

flow = Flow.from_client_config(
    client_config,
    scopes=SCOPES,
    redirect_uri=redirect_uri
)

if 'creds' not in st.session_state:
    auth_url, _ = flow.authorization_url(prompt='consent', access_type='offline')
    st.info("Please connect your Google account to begin.")
    st.link_button("🔑 Connect Google Account", auth_url)

    # Check if we are coming back from Google with a code
    code = st.query_params.get("code")
    if code:
        try:
            flow.fetch_token(code=code)
            st.session_state['creds'] = flow.credentials
            st.success("✅ Authentication Successful!")
            st.rerun()
        except Exception as e:
            st.error(f"Authentication Error: {e}")
            st.stop()
else:
    # If already logged in, show the app
    creds = st.session_state['creds']
    sheets_service = build('sheets', 'v4', credentials=creds)
    slides_service = build('slides', 'v1', credentials=creds)
    drive_service = build('drive', 'v3', credentials=creds)

    st.sidebar.success("Account Connected")
    if st.sidebar.button("Logout / Reset"):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.rerun()

    # --- 4. THE MAIN INTERFACE ---
    st.header("1. File Setup")
    sheet_url = st.text_input("Google Sheet URL")
    slide_url = st.text_input("Google Slide Template URL")
    folder_url = st.text_input("Google Drive Output Folder URL")

    if sheet_url and slide_url and folder_url:
        sheet_id = extract_id(sheet_url, 'sheet')
        slide_id = extract_id(slide_url, 'slides')
        folder_id = extract_id(folder_url, 'folder')

        st.header("2. PDF Configuration")
        st.info("Use `<<Column Name>>` in the filename template.")
        filename_template = st.text_input("PDF Filename Template:", value="<<School ID>> - Document")

        if st.button("🚀 Start Generating / Resume"):
            status_display = st.empty()
            progress_bar = st.progress(0)
            
            # Fetch latest data
            result = sheets_service.spreadsheets().values().get(
                spreadsheetId=sheet_id, range='Sheet1'
            ).execute()
            values = result.get('values', [])
            
            if not values:
                st.error("The Google Sheet is empty!")
            else:
                headers = values[0]
                if 'Status' not in headers or 'PDF Link' not in headers:
                    st.error("Missing 'Status' or 'PDF Link' columns in Header!")
                    st.stop()

                # Padding rows for consistent lengths
                max_cols = len(headers)
                df_rows = [r + [''] * (max_cols - len(r)) for r in values[1:]]
                df = pd.DataFrame(df_rows, columns=headers)

                status_idx = headers.index('Status') + 1
                link_idx = headers.index('PDF Link') + 1
                status_col_letter = col_to_letter(status_idx)
                link_col_letter = col_to_letter(link_idx)

                for index, row in df.iterrows():
                    # Resume Logic
                    if str(row.get('Status', '')).strip().lower() == 'success':
                        progress_bar.progress((index + 1) / len(df))
                        continue

                    try:
                        # Dynamic Filename
                        fname = filename_template
                        for h in headers:
                            fname = fname.replace(f"<<{h}>>", str(row.get(h, "")))
                        fname = fname.strip() + ".pdf" if not fname.lower().endswith(".pdf") else fname

                        status_display.text(f"Processing: {fname}")

                        # 1. Copy Slide
                        copy_body = {'name': fname.replace('.pdf', '')}
                        temp_slide = drive_service.files().copy(fileId=slide_id, body=copy_body).execute()
                        temp_id = temp_slide['id']

                        # 2. Replace Text
                        reqs = []
                        for h in headers:
                            reqs.append({
                                'replaceAllText': {
                                    'containsText': {'text': f"<<{h}>>", 'matchCase': False},
                                    'replaceText': str(row.get(h, ""))
                                }
                            })
                        slides_service.presentations().batchUpdate(presentationId=temp_id, body={'requests': reqs}).execute()

                        # 3. Export PDF
                        export_req = drive_service.files().export_media(fileId=temp_id, mimeType='application/pdf')
                        pdf_content = io.BytesIO()
                        downloader = MediaIoBaseDownload(pdf_content, export_req)
                        done = False
                        while not done:
                            _, done = downloader.next_chunk()

                        # 4. Upload to Folder
                        pdf_content.seek(0)
                        meta = {'name': fname, 'parents': [folder_id]}
                        media = MediaIoBaseUpload(pdf_content, mimetype='application/pdf', resumable=True)
                        uploaded = drive_service.files().create(body=meta, media_body=media, fields='webViewLink').execute()
                        link = uploaded.get('webViewLink')

                        # 5. Cleanup & Update Sheet
                        drive_service.files().delete(fileId=temp_id).execute()
                        
                        row_num = index + 2
                        sheets_service.spreadsheets().values().update(
                            spreadsheetId=sheet_id, range=f"Sheet1!{status_col_letter}{row_num}",
                            valueInputOption="USER_ENTERED", body={"values": [["Success"]]}
                        ).execute()
                        sheets_service.spreadsheets().values().update(
                            spreadsheetId=sheet_id, range=f"Sheet1!{link_col_letter}{row_num}",
                            valueInputOption="USER_ENTERED", body={"values": [[link]]}
                        ).execute()

                    except Exception as e:
                        st.error(f"Row {index+1} Error: {e}")

                    progress_bar.progress((index + 1) / len(df))

                status_display.text("All Finished!")
                st.balloons()
