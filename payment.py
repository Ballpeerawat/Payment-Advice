import os
import base64
import datetime
import fitz
import threading
import customtkinter as ctk
from tkinter import messagebox
from tkcalendar import DateEntry
from email import message_from_bytes
# from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

from pathlib import Path

# ---------- CONFIG ----------
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']
DOWNLOAD_FOLDER = str(Path.home() / "Downloads" / "Download Payment Adv")
SUBJECT_KEYWORD = "[Group Account] SCB Business Anywhere: New Transfer Received (ได้รับเงินโอนเข้าบัญชี)"

# ---------- PDF NAME EXTRACT ----------
def extract_info_from_pdf(filepath):
    try:
        doc = fitz.open(filepath)
        text = "\n".join([page.get_text() for page in doc])
        doc.close()

        lines = [line.strip() for line in text.split('\n') if line.strip()]
        
        ref = None
        date_str = None
        name = None

        for i, line in enumerate(lines):
            if line == "Product Name" and i + 1 < len(lines):
                ref = lines[i + 1]
            if line == "Value Date" and i + 1 < len(lines):
                date_str = lines[i + 1]
            if line.startswith("เรียน "):
                name = line.replace("เรียน ", "").strip()

        if ref and date_str:
            date_obj = datetime.datetime.strptime(date_str, "%d/%m/%Y")
            date_formatted = date_obj.strftime("%Y%m%d")
            filename = f"{ref}_{name}_{date_formatted}.pdf"
            return filename
        else:
            return None

    except Exception as e:
        print(f"❌ PDF Error: {e}")
        return None


# ---------- GMAIL AUTH ----------
global_creds = None
def get_gmail_service():
    creden_dict = {"installed":{"client_id":"643248617111-oa9ulfu99mhugui3cea39gp94cklghh4.apps.googleusercontent.com",
                    "project_id":"haupcar-staging","auth_uri":"https://accounts.google.com/o/oauth2/auth",
                    "token_uri":"https://oauth2.googleapis.com/token","auth_provider_x509_cert_url":"https://www.googleapis.com/oauth2/v1/certs",
                    "client_secret":"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","redirect_uris":["http://localhost"]}}
    global global_creds
    if global_creds:
        return build('gmail', 'v1', credentials=global_creds)

    flow = InstalledAppFlow.from_client_config(creden_dict, SCOPES)
    global_creds = flow.run_local_server(port=0)
    return build('gmail', 'v1', credentials=global_creds)

# ---------- DOWNLOAD AND RENAME ----------
def download_pdfs(service, start_date, end_date, update_progress):
    query = f'subject:"{SUBJECT_KEYWORD}" after:{start_date} before:{end_date}'
    results = service.users().messages().list(userId='me', q=query).execute()
    messages = results.get('messages', [])

    if not os.path.exists(DOWNLOAD_FOLDER):
        os.makedirs(DOWNLOAD_FOLDER)

    total = len(messages)
    count = 0

    for msg in messages:
        msg_data = service.users().messages().get(userId='me', id=msg['id'], format='raw').execute()
        msg_bytes = base64.urlsafe_b64decode(msg_data['raw'].encode('ASCII'))
        email_message = message_from_bytes(msg_bytes)

        for part in email_message.walk():
            filename = part.get_filename()
            if filename and filename.lower().endswith('.pdf'):
                temp_path = os.path.join(DOWNLOAD_FOLDER, f"temp_{count}.pdf")
                with open(temp_path, 'wb') as f:
                    f.write(part.get_payload(decode=True))

                new_name = extract_info_from_pdf(temp_path)
                if new_name:
                    new_path = os.path.join(DOWNLOAD_FOLDER, new_name)
                    if os.path.exists(new_path):
                        os.remove(new_path)
                    os.rename(temp_path, new_path)
                else:
                    dest_path = os.path.join(DOWNLOAD_FOLDER, filename)
                    if os.path.exists(dest_path):
                        os.remove(dest_path)
                    os.rename(temp_path, dest_path)

                count += 1
                update_progress(count, total)

    return count, total

# ---------- CUSTOM POPUP ----------
def show_custom_popup(title, message, level="info", on_close=None):
    popup = ctk.CTkToplevel()
    popup.geometry("400x230")
    popup.title(title)
    popup.resizable(False, False)
    popup.grab_set()  # modal window

    colors = {
        "info": ("#004085", "#cce5ff"),
        "warning": ("#856404", "#fff3cd"),
        "error": ("#721c24", "#f8d7da"),
        "success": ("#155724", "#d4edda")
    }
    fg, bg = colors.get(level, ("#000000", "#ffffff"))

    popup.configure(fg_color=bg)

    label_title = ctk.CTkLabel(popup, text=title, font=("Helvetica", 18, "bold"), text_color=fg)
    label_title.pack(pady=(20, 10))

    label_msg = ctk.CTkLabel(popup, text=message, font=("Helvetica", 14), text_color=fg, wraplength=300)
    label_msg.pack(padx=20, pady=(0, 20))

    def on_ok():
        popup.destroy()
        if on_close:
            on_close()

    btn_ok = ctk.CTkButton(
        popup,
        text="OK",
        command=on_ok,
        fg_color=fg,
        hover_color=fg,
        text_color=bg,
        width=120,      # ⬅️ ปรับความกว้าง
        height=50,      # ⬅️ ปรับความสูง
        font=("Helvetica", 12, "bold")  # ⬅️ เพิ่มขนาดฟอนต์ (ถ้าต้องการ)
    )
    btn_ok.pack(pady=20)

# ---------- MAIN FUNCTION ----------
def run_download(start_str, end_str, update_progress, update_status, on_done):
    try:
        start = datetime.datetime.strptime(start_str, "%d/%m/%Y")
        end = datetime.datetime.strptime(end_str, "%d/%m/%Y") + datetime.timedelta(days=1)
        start_query = int(start.timestamp())
        end_query = int(end.timestamp())

        update_status("🔐 Connecting to Gmail...")
        service = get_gmail_service()
        update_status("📥 Starting PDF download...")

        count, total = download_pdfs(service, start_query, end_query, update_progress)

        if total == 0:
            on_done(0, total, "No emails matched the criteria for this date range.")
        elif count == 0:
            on_done(0, total, "No PDF files found in the downloaded emails.")
        else:
            on_done(count, total, "")

    except Exception as e:
        on_done(-1, -1, str(e))

# ---------- GUI ----------
def launch_gui():
    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")

    # Colors as requested
    BACKGROUND_COLOR = "#F5EFFF"
    BUTTON_COLOR = "#441752"
    PRIMARY_COLOR = "#441752"
    TEXT_COLOR_BUTTON = "#F5EFFF"

    root = ctk.CTk()
    root.title("📥 Gmail Payment Advice Downloader")
    root.geometry("600x350")
    root.configure(fg_color=BACKGROUND_COLOR)

    title = ctk.CTkLabel(
        root,
        text="📥 Download Payment Advice PDFs from Gmail",
        font=("Helvetica", 20, "bold"),
        text_color=PRIMARY_COLOR)
    title.pack(pady=(20, 10))

    frame_date = ctk.CTkFrame(root, fg_color=BACKGROUND_COLOR)
    frame_date.pack(pady=20)

    start_label = ctk.CTkLabel(frame_date, text="Start Date", font=("Helvetica", 14, "bold"), text_color=PRIMARY_COLOR)
    start_label.grid(row=0, column=0)
    end_label = ctk.CTkLabel(frame_date, text="End Date", font=("Helvetica", 14, "bold"), text_color=PRIMARY_COLOR)
    end_label.grid(row=0, column=1)

    date_start = DateEntry(frame_date, width=12, background='darkblue', foreground='white', borderwidth=2, 
                            date_pattern='dd/mm/yyyy', font=("Helvetica", 14))
    date_start.grid(row=1, column=0, padx=10, pady=5)
    date_end = DateEntry(frame_date, width=12, background='darkblue', foreground='white', borderwidth=2, 
                            date_pattern='dd/mm/yyyy', font=("Helvetica", 14))
    date_end.grid(row=1, column=1, padx=10, pady=5)

    label_range = ctk.CTkLabel(root, text="Please select start and end dates", font=("Helvetica", 16), text_color=PRIMARY_COLOR)
    label_range.pack(pady=(5, 0))

    progress_var = ctk.DoubleVar()
    progress_bar = ctk.CTkProgressBar(
        root,
        variable=progress_var,
        progress_color='#B33791',
        fg_color="#F4CCE9"
    )
    progress_bar.pack(fill='x', padx=20, pady=(15, 0))

    label_progress = ctk.CTkLabel(root, text="0 / 0", font=("Helvetica", 14), text_color=PRIMARY_COLOR)
    label_progress.pack(pady=(5, 15))

    def update_progress(current, total):
        if total > 0:
            progress_var.set(current / total)
            label_progress.configure(text=f"Downloading... {current} / {total}")
        else:
            progress_var.set(0)
            label_progress.configure(text="0 / 0")
        root.update()

    def update_status(msg):
        label_range.configure(text=msg)
        root.update()

    def download_done(count, total, err_msg):
        if count == -1:
            show_custom_popup("Error", f"An error occurred during download:\n{err_msg}", level="error")
            label_range.configure(text="Error occurred")
            progress_var.set(0)
            label_progress.configure(text="0 / 0")
        elif total == 0:
            show_custom_popup("Warning", "No emails matched the criteria for the selected date range.", level="warning")
            label_range.configure(text="No emails in selected date range")
            progress_var.set(0)
            label_progress.configure(text="0 / 0")
        elif count == 0:
            show_custom_popup("Warning", "No PDF files found in the downloaded emails.", level="warning")
            label_range.configure(text="No PDF files found in emails")
            progress_var.set(0)
            label_progress.configure(text="0 / 0")
        else:
            # Show success popup and exit program on OK
            show_custom_popup(
                "Success",
                f"Downloaded {count} files out of {total} emails\n\nSaved to:\n{DOWNLOAD_FOLDER}",
                level="success",
                on_close=root.destroy  # close app on OK
            )
            label_range.configure(text=f"Download completed {count} / {total}")
            progress_var.set(1)
            label_progress.configure(text=f"{count} / {total}")

    def on_click():
        progress_var.set(0)
        label_progress.configure(text="0 / 0")
        label_range.configure(text="Please select start and end dates")
        start = date_start.get_date().strftime("%d/%m/%Y")
        end = date_end.get_date().strftime("%d/%m/%Y")
        if start and end:
            label_range.configure(text=f"Date range: {start} to {end}")
            threading.Thread(target=run_download, args=(start, end, update_progress, update_status, download_done), daemon=True).start()
        else:
            messagebox.showwarning("Missing input", "Please select both start and end dates")

    download_btn = ctk.CTkButton(
        root,
        text="🚀 Download",
        command=on_click,
        corner_radius=10,
        fg_color=BUTTON_COLOR,
        hover_color=PRIMARY_COLOR,
        text_color=TEXT_COLOR_BUTTON,
        width=150,
        height=40,
        font=("Helvetica", 12, "bold")
    )
    download_btn.pack(pady=10)

    root.mainloop()

if __name__ == "__main__":
    launch_gui()
