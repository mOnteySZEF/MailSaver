import os
import io
import base64
import re
import threading
from datetime import datetime, timedelta
from tkinter import *
from tkinter import ttk, filedialog
from tkcalendar import DateEntry
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

SCOPES = [
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/drive.readonly'
]

SZKOLA_PREFIXES = ['1AP', '2AP', '2BP', '3AI', '3BI', '1AI', '2AI', '2BI']

class MailSaverApp:
    def __init__(self, root):
        self.root = root
        self.root.title("MailSaver - E-Mail Downloader by mOntey")
        self.root.geometry("760x580")
        self.root.resizable(False, False)
        self.root.iconbitmap("MailSaver.ico")
        self.root.configure(bg="#f0f4f7")

        self.search_type = StringVar(value="fraza")
        self.search_phrase = StringVar()
        self.sender_filter = StringVar()
        self.export_logs = BooleanVar()
        self.folder_path = StringVar(value=os.path.abspath("pobrane"))

        today = datetime.today()
        self.date_from = StringVar(value=today.strftime('%Y-%m-%d'))
        self.date_to = StringVar(value=(today + timedelta(days=1)).strftime('%Y-%m-%d'))

        style = ttk.Style(root)
        style.configure('TButton', font=('Segoe UI', 10), padding=5)
        style.configure('TLabel', background='#f0f4f7', font=('Segoe UI', 10))
        style.configure('TLabelframe', background='#f0f4f7', font=('Segoe UI', 11, 'bold'))
        style.configure('TLabelframe.Label', background='#f0f4f7', font=('Segoe UI', 11, 'bold'))

        search_frame = ttk.Labelframe(self.root, text="Ustawienia wyszukiwania", padding=10)
        search_frame.pack(fill=X, padx=10, pady=10)

        ttk.Radiobutton(search_frame, text="Fraza", variable=self.search_type, value="fraza",
                        command=self.toggle_phrase).grid(row=0, column=0, sticky=W, padx=5, pady=2)
        ttk.Radiobutton(search_frame, text="Szkola", variable=self.search_type, value="szkola",
                        command=self.toggle_phrase).grid(row=0, column=1, sticky=W, padx=5, pady=2)

        self.phrase_label = ttk.Label(search_frame, text="Fraza w tytule e-mailu:")
        self.phrase_label.grid(row=1, column=0, sticky=W, padx=5, pady=2)
        self.phrase_entry = ttk.Entry(search_frame, textvariable=self.search_phrase, width=40)
        self.phrase_entry.grid(row=1, column=1, columnspan=2, sticky=W, padx=5, pady=2)

        ttk.Label(search_frame, text="Filtr nadawcy (opcjonalne):").grid(row=2, column=0, sticky=W, padx=5, pady=2)
        ttk.Entry(search_frame, textvariable=self.sender_filter, width=40).grid(row=2, column=1,
                        columnspan=2, sticky=W, padx=5, pady=2)

        ttk.Label(search_frame, text="Data od:").grid(row=3, column=0, sticky=W, padx=5, pady=2)
        self.date_from_entry = DateEntry(search_frame, textvariable=self.date_from,
                                         date_pattern='yyyy-mm-dd', width=12)
        self.date_from_entry.grid(row=3, column=1, sticky=W, padx=5)

        ttk.Label(search_frame, text="Data do:").grid(row=3, column=2, sticky=W, padx=5, pady=2)
        self.date_to_entry = DateEntry(search_frame, textvariable=self.date_to,
                                       date_pattern='yyyy-mm-dd', width=12)
        self.date_to_entry.grid(row=3, column=3, sticky=W, padx=5)
        self.date_to_entry.set_date(today + timedelta(days=1))

        ttk.Checkbutton(search_frame, text="Eksportuj logi", variable=self.export_logs).grid(
                        row=4, column=1, sticky=W, padx=5, pady=5)

        ttk.Label(search_frame, text="Folder zapisu:").grid(row=5, column=0, sticky=W, padx=5, pady=2)
        ttk.Entry(search_frame, textvariable=self.folder_path, width=35).grid(row=5, column=1,
                        sticky=W, padx=5, pady=2)
        ttk.Button(search_frame, text="Wybierz...", command=self.choose_folder).grid(row=5,
                        column=2, sticky=W, padx=5)

        ttk.Button(search_frame, text="Start", command=self.start_thread).grid(row=6,
                        column=1, pady=10)

        log_frame = ttk.Labelframe(self.root, text="Logi", padding=10)
        log_frame.pack(fill=BOTH, expand=True, padx=10, pady=(0,10))
        self.log_area = Text(log_frame, height=15, width=90, state=DISABLED,
                             bg='#ffffff', fg='#000000')
        self.log_area.pack(fill=BOTH, expand=True)

        self.toggle_phrase()

    def toggle_phrase(self):
        if self.search_type.get() == 'szkola':
            self.phrase_label.grid_remove()
            self.phrase_entry.grid_remove()
        else:
            self.phrase_label.grid()
            self.phrase_entry.grid()

    def log(self, msg):
        self.log_area.config(state=NORMAL)
        self.log_area.insert(END, msg + "\n")
        self.log_area.see(END)
        self.log_area.config(state=DISABLED)

    def choose_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.folder_path.set(path)

    def start_thread(self):
        threading.Thread(target=self.run, daemon=True).start()

    def authenticate(self):
        creds = None
        if os.path.exists('token.json'):
            creds = Credentials.from_authorized_user_file('token.json', SCOPES)
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
                creds = flow.run_local_server(port=0)
            with open('token.json', 'w') as token:
                token.write(creds.to_json())
        return creds

    def run(self):
        stype = self.search_type.get()
        phrase = self.search_phrase.get().strip()
        sender = self.sender_filter.get().strip()
        base = self.folder_path.get()
        export = self.export_logs.get()
        try:
            df = datetime.strptime(self.date_from.get(), '%Y-%m-%d')
            dt = datetime.strptime(self.date_to.get(), '%Y-%m-%d') + timedelta(days=1)
        except:
            self.log('❗ Nieprawidłowe daty')
            return

        if stype == 'fraza' and not phrase:
            self.log('❗ Fraza wymagana')
            return

        creds = self.authenticate()
        gmail = build('gmail', 'v1', credentials=creds)
        drive = build('drive', 'v3', credentials=creds)

        self.log('🔍 Szukanie w Gmail...')
        query_parts = [
            f"after:{df.strftime('%Y/%m/%d')}",
            f"before:{dt.strftime('%Y/%m/%d')}"
        ]
        if stype == 'fraza':
            query_parts.append(f"subject:{phrase}")
        if sender:
            query_parts.append(f"from:{sender}")
        query = ' '.join(query_parts)
        self.log('Zapytanie Gmail: ' + query)
        msgs = gmail.users().messages().list(userId='me', q=query, maxResults=200).execute().get('messages', [])

        for m in msgs:
            detail = gmail.users().messages().get(userId='me', id=m['id'], format='full').execute()
            headers = detail['payload']['headers']
            subj = next((h['value'] for h in headers if h['name']=='Subject'), '')

            if stype == 'szkola' and not any(subj.startswith(p) for p in SZKOLA_PREFIXES):
                continue
            folder = next((p for p in SZKOLA_PREFIXES if subj.startswith(p)), subj[:3] or 'Nieznane')
            target = os.path.join(base, folder, subj)
            os.makedirs(target, exist_ok=True)

            snippet = detail.get('snippet', '')
            links = re.findall(r'(https?://\S+)', snippet)
            if links:
                with open(os.path.join(target, 'linki.txt'), 'w', encoding='utf-8') as f:
                    f.write("\n".join(links))

            parts = detail['payload'].get('parts', [])
            for part in parts:
                fid = part.get('body', {}).get('attachmentId')
                name = part.get('filename')
                if fid and name:
                    att = gmail.users().messages().attachments().get(
                        userId='me', messageId=m['id'], id=fid).execute()
                    data = att.get('data')
                    if data:
                        content = base64.urlsafe_b64decode(data)
                        with open(os.path.join(target, name), 'wb') as f:
                            f.write(content)
                        self.log('📁 Pobrano: ' + name)

        self.log('🔍 Szukanie w Drive (Udostępnione)...')
        dq = ["sharedWithMe"]
        if stype == 'fraza':
            dq.append(f"name contains '{phrase}'")
        if stype == 'szkola':
            cond = [f"name contains '{p}'" for p in SZKOLA_PREFIXES]
            dq.append('(' + ' or '.join(cond) + ')')
        dquery = ' and '.join(dq)
        self.log('Zapytanie Drive: ' + dquery)
        files = drive.files().list(q=dquery, fields='files(id,name)').execute().get('files', [])
        for f in files:
            fid, name = f['id'], f['name']
            folder = next((p for p in SZKOLA_PREFIXES if name.startswith(p)), name[:3] or 'Nieznane')
            target = os.path.join(base, folder, name)
            os.makedirs(target, exist_ok=True)
            fh = io.BytesIO()
            downloader = MediaIoBaseDownload(fh, drive.files().get_media(fileId=fid))
            done = False
            while not done:
                _, done = downloader.next_chunk()
            with open(os.path.join(target, name), 'wb') as out:
                out.write(fh.getvalue())
            self.log('📂 Drive: ' + name)

        self.log('✅ Gotowe')
        if export:
            with open(os.path.join(base, 'logi.txt'), 'w', encoding='utf-8') as logf:
                logf.write(self.log_area.get('1.0', END))

if __name__ == '__main__':
    root = Tk()
    MailSaverApp(root)
    root.mainloop()