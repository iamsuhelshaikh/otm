# Onboarding Template Maker (OTM)

> Simple GUI tool to create onboarding email templates and manage a small CSV database.  
> Exports clipboard-friendly HTML fragments (headings preserved) and supports saving a persistent CSV database path.

<img width="1152" height="754" alt="image" src="https://github.com/user-attachments/assets/b0e6a7f1-b10b-4be2-8504-3e49ec7cc4cc" />

## Features
- Create/Edit onboarding records (Full name, Username, Password, PC Name, Ext, DDI, Email)
- Preview formatted email with bold/underline headings
- Copy preview to clipboard as HTML fragment (Windows/Outlook-friendly)
- Save / update / delete records in a CSV database
- Persist selected CSV path across app restarts
- Save newest records at top of the DB
- Build a single-file Windows EXE with PyInstaller
- Taskbar/Pin friendly (AppID + proper .ico handling)


## Requirements
- Python 3.10+ (3.11/3.12/3.13 tested)
- tkinter (standard with CPython)
- pyperclip (optional; used if available)
- PyInstaller (for building EXE)
