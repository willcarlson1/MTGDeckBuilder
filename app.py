import os
from flask import Flask, render_template, request, redirect, url_for

try:
    import win32com.client
except ImportError:  # allows repository browsing on non-Windows systems
    win32com = None

app = Flask(__name__)

def move_messages_by_subject(subjects):
    if win32com is None:
        raise RuntimeError("win32com is unavailable. This app must run on Windows with Outlook installed.")

    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")

    inbox = namespace.GetDefaultFolder(6)  # olFolderInbox
    dest = namespace.GetDefaultFolder(3)   # olFolderDeletedItems

    messages = inbox.Items
    for mail in list(messages):
        try:
            subject_clean = str(mail.Subject).strip().lower()
        except Exception:
            continue
        if subject_clean in subjects:
            mail.Move(dest)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        text = request.form.get('subjects', '')
        subjects = [s.strip().lower() for s in text.splitlines() if s.strip()]
        if subjects:
            move_messages_by_subject(subjects)
        return redirect(url_for('done'))
    return render_template('index.html')

@app.route('/done')
def done():
    return 'Task complete – please manually review Deleted Items.'

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
