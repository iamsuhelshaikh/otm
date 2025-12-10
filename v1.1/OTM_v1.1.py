#!/usr/bin/env python3

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import csv
import pathlib
import sys
import os
import ctypes
import json

# Optional dependency for simple clipboard copy if installed
try:
    import pyperclip
    _HAVE_PYPERCLIP = True
except Exception:
    _HAVE_PYPERCLIP = False

# Config and DB paths
DEFAULT_DB = pathlib.Path("onboarding_db.csv")
CONFIG_FILE = pathlib.Path.home() / '.otm_config.json'
DB_FILE = DEFAULT_DB

FIELDS = ["full_name", "username", "password", "pc_name", "ext", "ddi", "email"]

# ----------------- Config helpers -----------------
def load_config():
    global DB_FILE
    try:
        if CONFIG_FILE.exists():
            with CONFIG_FILE.open('r', encoding='utf-8') as f:
                data = json.load(f)
            p = data.get('db_path')
            if p:
                p = pathlib.Path(p)
                # If configured path exists and appears to be a file, use it
                if p.exists():
                    DB_FILE = p
                else:
                    # If configured path doesn't exist, fall back to default but keep config
                    DB_FILE = pathlib.Path(p)
        else:
            DB_FILE = DEFAULT_DB
    except Exception:
        # On any config error, fall back to default DB
        DB_FILE = DEFAULT_DB

def save_config():
    try:
        with CONFIG_FILE.open('w', encoding='utf-8') as f:
            json.dump({'db_path': str(DB_FILE)}, f)
    except Exception:
        # silently ignore config save failures
        pass

# ----------------- Utilities -----------------

def ensure_db(path: pathlib.Path):
    if not path.exists():
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open("w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, FIELDS)
            writer.writeheader()

# load config early and ensure DB exists
load_config()
ensure_db(DB_FILE)

def make_username(full_name: str) -> str:
    """Return username with capitalized initials, joined by dots.
    e.g. 'suhel shaikh' -> 'Suhel.Shaikh'"""
    parts = [p for p in full_name.strip().split() if p]
    return ".".join([p.capitalize() for p in parts]) if parts else ""

def make_display_username(full_name: str) -> str:
    # keep this for compatibility (same as make_username now)
    return make_username(full_name)

def load_records():
    ensure_db(DB_FILE)
    with DB_FILE.open("r", newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))

def save_records(records):
    ensure_db(DB_FILE)
    # atomic write: write to temp file then replace
    tmp = DB_FILE.with_suffix('.tmp')
    with tmp.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, FIELDS)
        writer.writeheader()
        writer.writerows(records)
    try:
        os.replace(tmp, DB_FILE)
    except Exception:
        # fallback to non-atomic if replace fails
        with DB_FILE.open("w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, FIELDS)
            writer.writeheader()
            writer.writerows(records)

def open_csv_default():
    path = DB_FILE.resolve()
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)
        elif sys.platform == "darwin":
            # use subprocess to avoid shell-escaping issues
            import subprocess
            subprocess.run(["open", str(path)])
        else:
            import subprocess
            subprocess.run(["xdg-open", str(path)])
    except Exception as e:
        messagebox.showerror("Open CSV", f"Could not open CSV: {e}")

# ----------------- CF_HTML clipboard helpers (Windows) -----------------
def strip_tags(html: str) -> str:
    import re
    text = re.sub(r'<br\s*/?>', '\n', html, flags=re.I)
    text = re.sub(r'<[^>]+>', '', text)
    return text

def _make_cf_html_bytes(fragment_html: str) -> bytes:
    if isinstance(fragment_html, str):
        fragment_bytes = fragment_html.encode('utf-8')
    else:
        fragment_bytes = fragment_html
    start_fragment_marker = b"<!--StartFragment-->"
    end_fragment_marker = b"<!--EndFragment-->"
    html_body = b"<html><body>" + start_fragment_marker + fragment_bytes + end_fragment_marker + b"</body></html>"
    header_template = b"Version:0.9\r\nStartHTML:%08d\r\nEndHTML:%08d\r\nStartFragment:%08d\r\nEndFragment:%08d\r\n"
    header_len = len(header_template % (0,0,0,0))
    start_html = header_len
    start_fragment = start_html + html_body.find(start_fragment_marker) + len(start_fragment_marker)
    end_fragment = start_html + html_body.find(end_fragment_marker)
    end_html = start_html + len(html_body)
    full = header_template % (start_html, end_html, start_fragment, end_fragment) + html_body
    return full

def _set_clipboard_html_win32(html_fragment: str, plain_text: str = None):
    if plain_text is None:
        plain_text = strip_tags(html_fragment)
    data_cf_html = _make_cf_html_bytes(html_fragment)
    data_unicode = plain_text.encode('utf-16le')
    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32
    CF_UNICODETEXT = 13
    GMEM_MOVEABLE = 0x0002
    reg_html = user32.RegisterClipboardFormatW("HTML Format")
    if not user32.OpenClipboard(0):
        raise RuntimeError("OpenClipboard failed")
    try:
        if not user32.EmptyClipboard():
            raise RuntimeError("EmptyClipboard failed")
        h_global_html = kernel32.GlobalAlloc(GMEM_MOVEABLE, len(data_cf_html) + 1)
        if not h_global_html:
            raise RuntimeError("GlobalAlloc failed for HTML")
        p_html = kernel32.GlobalLock(h_global_html)
        if not p_html:
            kernel32.GlobalFree(h_global_html)
            raise RuntimeError("GlobalLock failed for HTML")
        ctypes.memmove(p_html, data_cf_html, len(data_cf_html))
        kernel32.GlobalUnlock(h_global_html)
        if not user32.SetClipboardData(reg_html, h_global_html):
            kernel32.GlobalFree(h_global_html)
            raise RuntimeError("SetClipboardData failed for HTML")
        h_global_txt = kernel32.GlobalAlloc(GMEM_MOVEABLE, len(data_unicode) + 2)
        if not h_global_txt:
            raise RuntimeError("GlobalAlloc failed for text")
        p_txt = kernel32.GlobalLock(h_global_txt)
        if not p_txt:
            kernel32.GlobalFree(h_global_txt)
            raise RuntimeError("GlobalLock failed for text")
        ctypes.memmove(p_txt, data_unicode, len(data_unicode))
        ctypes.memmove(p_txt + len(data_unicode), b'\x00\x00', 2)
        kernel32.GlobalUnlock(h_global_txt)
        if not user32.SetClipboardData(CF_UNICODETEXT, h_global_txt):
            kernel32.GlobalFree(h_global_txt)
            raise RuntimeError("SetClipboardData failed for text")
    finally:
        user32.CloseClipboard()

def set_html_clipboard_fragment(html_fragment: str, plain_text: str = None):
    if sys.platform.startswith("win"):
        try:
            # try pywin32 if present
            import win32clipboard, win32con  # type: ignore
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            cf_html = win32clipboard.RegisterClipboardFormat("HTML Format")
            data = _make_cf_html_bytes(html_fragment)
            win32clipboard.SetClipboardData(cf_html, data)
            plain = strip_tags(html_fragment) if plain_text is None else plain_text
            win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT, plain)
            win32clipboard.CloseClipboard()
            return
        except Exception:
            try:
                _set_clipboard_html_win32(html_fragment, plain_text)
                return
            except Exception:
                pass
    # fallback: regular plain text via Tk clipboard (non-windows or failure)
    try:
        root.clipboard_clear()
        root.clipboard_append(strip_tags(html_fragment) if plain_text is None else plain_text)
    except Exception:
        pass

# ----------------- GUI -----------------
root = tk.Tk()
root.title("Onboarding Template Maker V1.1")
# match screenshot proportions and keep 'held' as requested
#root.geometry("1100x680")
#root.minsize(900, 600)
#root.resizable(False, False)  # hold layout as in image

style = ttk.Style(root)
try:
    style.theme_use('clam')
except Exception:
    pass

try:
    root.option_add("*Font", ("Segoe UI", 10))
except Exception:
    try:
        root.option_add("*Font", ("Helvetica", 10))
    except Exception:
        pass

# Top-level paned (vertical) -> top (form+preview) and bottom (db)
paned = ttk.Panedwindow(root, orient='vertical')
paned.pack(fill='both', expand=True)

top_frame = ttk.Frame(paned, padding=10)
paned.add(top_frame, weight=7)
bottom_frame = ttk.Frame(paned, padding=10)
paned.add(bottom_frame, weight=3)

# horizontal paned for top area (form left / preview right)
top_paned = ttk.Panedwindow(top_frame, orient='horizontal')
top_paned.pack(fill='both', expand=True)
left_frame = ttk.Frame(top_paned)
right_frame = ttk.Frame(top_paned)
top_paned.add(left_frame, weight=3)
top_paned.add(right_frame, weight=5)

# Create/Edit frame (left)
create_frame = ttk.LabelFrame(left_frame, text='Create / Edit Record', padding=12)
create_frame.pack(fill='both', expand=True)
# grid columns: 0 labels, 1 entries, 2 copy-buttons
create_frame.columnconfigure(0, weight=3)
create_frame.columnconfigure(1, weight=8)
create_frame.columnconfigure(2, weight=0)

# Copy button style (subtle)
style.configure('Copy.TButton', relief='flat', padding=1)

# Bigger copy icon style
style.configure('CopyBig.TButton', font=("Segoe UI Emoji", 12), padding=2, relief='flat')

# Reduce Treeview font slightly to avoid horizontal scroll
style.configure('Treeview', font=(None, 9))

# Helper to build labeled entries with visible copy button using emoji
def lfield(parent, row, label_text, show=None):
    lbl = ttk.Label(parent, text=label_text)
    lbl.grid(row=row, column=0, sticky='w', pady=6, padx=(0,6))
    ent = ttk.Entry(parent, show=show)
    ent.grid(row=row, column=1, sticky='we', pady=6)
    copy_btn = ttk.Button(parent, text='📑', width=3, style='CopyBig.TButton',
                          command=lambda e=ent, b=None: copy_from_widget(e))
    copy_btn.grid(row=row, column=2, sticky='w', padx=(6,0))
    copy_btn.configure(takefocus=False)
    return ent

# Build fields with proper tab order
ent_full = lfield(create_frame, 1, 'Full Name')
ent_pc = lfield(create_frame, 2, 'PC Name')
ent_pass = lfield(create_frame, 3, 'Password')
ent_ext = lfield(create_frame, 4, 'Ext')
ent_ddi = lfield(create_frame, 5, 'DDI')

# Username & Email labels with bigger copy buttons
lbl_username = ttk.Label(create_frame, text='Username: ')
lbl_username.grid(row=6, column=0, columnspan=2, sticky='w', pady=(6,0))
btn_copy_username = ttk.Button(create_frame, text='📑', width=3, style='CopyBig.TButton',
                               command=lambda: copy_from_widget(lbl_username))
btn_copy_username.grid(row=6, column=2, sticky='w', pady=(6,0))
btn_copy_username.configure(takefocus=False)

lbl_email = ttk.Label(create_frame, text='Email: ')
lbl_email.grid(row=7, column=0, columnspan=2, sticky='w', pady=(0,8))
btn_copy_email = ttk.Button(create_frame, text='📑', width=3, style='CopyBig.TButton',
                            command=lambda: copy_from_widget(lbl_email))
btn_copy_email.grid(row=7, column=2, sticky='w', pady=(0,8))
btn_copy_email.configure(takefocus=False)

# Buttons row (now 4 buttons: Save | Update | Copy Preview | Clear)
btns_row = ttk.Frame(create_frame)
btns_row.grid(row=20, column=0, columnspan=3, sticky='we', pady=(6,0))
# make four equal columns for neat alignment
for i in range(4):
    btns_row.columnconfigure(i, weight=1)

# Copy helpers for entries/labels
def _copy_text_to_clipboard(text):
    try:
        if _HAVE_PYPERCLIP:
            pyperclip.copy(text)
        else:
            root.clipboard_clear()
            root.clipboard_append(text)
    except Exception:
        try:
            root.clipboard_clear()
            root.clipboard_append(text)
        except Exception:
            messagebox.showerror("Copy", "Could not copy to clipboard.")

def flash_button_minimal(btn):
    # very subtle visual feedback by briefly changing relief (best-effort)
    try:
        orig = btn.cget('style')
        # no heavy change; skip for minimal look
    except Exception:
        pass

def copy_from_widget(widget):
    try:
        if isinstance(widget, (ttk.Entry, tk.Entry)):
            val = widget.get()
        elif isinstance(widget, tk.Text):
            val = widget.get("1.0", tk.END).rstrip("\n")
        elif isinstance(widget, (ttk.Label, tk.Label)):
            val = widget.cget("text")
            # strip "Username: " or "Email: " prefix
            if ":" in val:
                val = val.split(":", 1)[1].strip()
        else:
            val = str(widget)
        _copy_text_to_clipboard(val)
    except Exception as e:
        messagebox.showerror("Copy", f"Copy failed: {e}")

# Derived fields behavior
def update_derived(event=None):
    full = ent_full.get().strip()
    if full:
        u = make_username(full)
        lbl_username.config(text=f'Username: {u}')
        lbl_email.config(text=f'Email: {u}@totalassist.co.uk')
    else:
        lbl_username.config(text='Username: ')
        lbl_email.config(text='Email: ')

# Preview and formatting
def format_email_text(rec: dict) -> str:
    # Use stored username if present, else compute from full_name
    username = rec.get('username') or make_username(rec.get('full_name',''))
    display_user = username
    text = f"""Hi,

Please see below for {rec.get('full_name','')}'s login details.

PC login:
PC Name: {rec.get('pc_name','')}
Username: {display_user}
Password: {rec.get('password','')}

Itris:
Username: {display_user}
Password: {rec.get('password','')}

Telephone and DDI:
Ext: {rec.get('ext','')}
DDI: {rec.get('ddi','')}

Thanks,
"""
    return text

def format_html_lines_for_fragment(lines):
    """Convert a list of plain-text lines to a simple HTML fragment preserving our headings
       (wrap heading lines with <b><u>...</u></b>). Return HTML fragment string."""
    out = []
    for raw in lines:
        s = raw.rstrip("\n")
        if not s:
            out.append("<br/>")
            continue
        # detect heading lines
        if s.startswith("PC login:") or s.startswith("Itris:") or s.startswith("Telephone and DDI:"):
            out.append(f"<b><u>{escape_html(s)}</u></b><br/>")
        else:
            out.append(f"{escape_html(s)}<br/>")
    return "".join(out)

def escape_html(s):
    import html
    return html.escape(s)

# Right: Preview frame
preview_labelframe = ttk.LabelFrame(right_frame, text='Preview', padding=8)
preview_labelframe.pack(fill='both', expand=True)

preview_labelframe.pack_propagate(False)
preview_labelframe.config(height=240)   # adjust number if needed (lower = less height)

preview_text = tk.Text(preview_labelframe, wrap='word', font=('Calibri', 11), undo=True)
preview_vsb = ttk.Scrollbar(preview_labelframe, orient='vertical', command=preview_text.yview)
preview_text.configure(yscrollcommand=preview_vsb.set)
preview_vsb.pack(side='right', fill='y')
preview_text.pack(side='left', fill='both', expand=True)
preview_text.config(state='disabled')

# Tag for headings (bold + underline in UI)
preview_text.tag_configure('heading', font=('Calibri', 11, 'bold', 'underline'))

def show_preview(sample_when_empty=True):
    full = ent_full.get().strip()
    if not full:
        if sample_when_empty:
            sample = {
                'full_name': "Name Surname",
                'username': 'Name.Surname',
                'password': 'Total1999!',
                'pc_name': 'Total-DXXX',
                'ext': '000',
                'ddi': '01234 56789',
                'email': 'Name.Surname@totalassist.co.uk'
            }
            txt = format_email_text(sample)
        else:
            txt = ''
    else:
        tmp = {
            'full_name': full,
            'username': make_username(full),
            'password': ent_pass.get().strip(),
            'pc_name': ent_pc.get().strip(),
            'ext': ent_ext.get().strip(),
            'ddi': ent_ddi.get().strip(),
            'email': f"{make_username(full)}@totalassist.co.uk",
        }
        txt = format_email_text(tmp)
    # Insert with tagging
    preview_text.config(state='normal')
    preview_text.delete('1.0', tk.END)
    for line in txt.splitlines(True):
        if line.strip().startswith("PC login:") or line.strip().startswith("Itris:") or line.strip().startswith("Telephone and DDI:"):
            start = preview_text.index(tk.END)
            preview_text.insert(tk.END, line)
            end = preview_text.index(tk.END)
            preview_text.tag_add('heading', f"{start} linestart", f"{end} lineend")
        else:
            preview_text.insert(tk.END, line)
    preview_text.config(state='disabled')

# Intercept Ctrl+C inside preview_text to put HTML fragment on clipboard (Windows)
def preview_copy_handler(event=None):
    try:
        try:
            sel = preview_text.get("sel.first", "sel.last")
        except tk.TclError:
            sel = preview_text.get("1.0", tk.END).rstrip("\n")

        lines = sel.splitlines(True)

        # Send ONLY fragment — not full HTML
        fragment = format_html_lines_for_fragment(lines)

        set_html_clipboard_fragment(fragment, plain_text=sel)

        return "break"
    except Exception:
        return

# Bind Ctrl+C (and Ctrl+Insert) for preview_text
preview_text.bind("<Control-c>", lambda e: preview_copy_handler(e))
preview_text.bind("<Control-C>", lambda e: preview_copy_handler(e))
preview_text.bind("<Control-Insert>", lambda e: preview_copy_handler(e))

# Fields change -> update derived and preview
def on_field_change(event=None):
    update_derived()
    show_preview()

for w in (ent_full, ent_pc, ent_pass, ent_ext, ent_ddi):
    w.bind('<KeyRelease>', on_field_change)

# Buttons for create/edit
def clear_form():
    for w in (ent_full, ent_pc, ent_pass, ent_ext, ent_ddi):
        w.delete(0, tk.END)
    update_derived()
    show_preview()

def save_new():
    full = ent_full.get().strip()
    pc = ent_pc.get().strip()
    pw = ent_pass.get().strip()
    ext = ent_ext.get().strip()
    ddi = ent_ddi.get().strip()
    if not (full and pc and pw):
        messagebox.showwarning('Missing fields', 'Please provide Full Name, PC Name and Password.')
        return
    username = make_username(full)
    rec = {
        'full_name': full,
        'username': username,
        'password': pw,
        'pc_name': pc,
        'ext': ext,
        'ddi': ddi,
        'email': f'{username}@totalassist.co.uk',
    }
    recs = load_records()
    # insert new record at top (index 0)
    recs.insert(0, rec)
    save_records(recs)
    refresh_data()
    clear_form()

current_index = None
def update_record():
    global current_index
    if current_index is None:
        messagebox.showwarning('No selection', 'Select a record from the database to update.')
        return
    full = ent_full.get().strip()
    pc = ent_pc.get().strip()
    pw = ent_pass.get().strip()
    ext = ent_ext.get().strip()
    ddi = ent_ddi.get().strip()
    if not (full and pc and pw):
        messagebox.showwarning('Missing fields', 'Please provide Full Name, PC Name and Password.')
        return
    username = make_username(full)
    rec = {
        'full_name': full,
        'username': username,
        'password': pw,
        'pc_name': pc,
        'ext': ext,
        'ddi': ddi,
        'email': f'{username}@totalassist.co.uk',
    }
    recs = load_records()
    if 0 <= current_index < len(recs):
        recs[current_index] = rec
        save_records(recs)
        refresh_data()
        clear_form()
    else:
        messagebox.showerror('Update', 'Selected record index is invalid.')

# Create the four aligned buttons
btn_save = ttk.Button(btns_row, text='Save', command=save_new)
btn_save.grid(row=0, column=0, sticky='we', padx=6)
btn_update = ttk.Button(btns_row, text='Update', command=update_record)
btn_update.grid(row=0, column=1, sticky='we', padx=6)

# Copy preview button: copies full preview (HTML if supported)
def copy_preview_action():
    try:
        txt = preview_text.get("1.0", tk.END).rstrip("\n")
        lines = txt.splitlines(True)
        html_fragment = "<div>" + format_html_lines_for_fragment(lines) + "</div>"
        set_html_clipboard_fragment(html_fragment, plain_text=txt)
        messagebox.showinfo("Copied", "Preview copied to clipboard (HTML if supported).")
    except Exception as e:
        messagebox.showerror("Copy", f"Could not copy: {e}")

btn_copy_preview = ttk.Button(btns_row, text='Copy', command=copy_preview_action)
btn_copy_preview.grid(row=0, column=2, sticky='we', padx=6)
btn_clear = ttk.Button(btns_row, text='Clear', command=clear_form)
btn_clear.grid(row=0, column=3, sticky='we', padx=6)

# ----------------- Bottom: Database -----------------
lbl_db_header = ttk.Label(bottom_frame, text='Search / Records', font=(None, 11, 'bold'))
lbl_db_header.pack(anchor='w')

search_bar = ttk.Frame(bottom_frame)
search_bar.pack(fill='x', pady=(6,8))

ent_search = ttk.Entry(search_bar)
ent_search.pack(side='left', fill='x', expand=True, padx=(0,6))

btn_find = ttk.Button(search_bar, text='Find', width=10)
btn_find.pack(side='left', padx=2)

btn_open = ttk.Button(search_bar, text='Open CSV', width=12, command=open_csv_default)
btn_open.pack(side='left', padx=2)

# Treeview area
tree_frame = ttk.Frame(bottom_frame)
tree_frame.pack(fill='both', expand=True)
cols = ["full_name", "username", "password", "pc_name", "email", "ext", "ddi"]
# show headings and no tree column
tree = ttk.Treeview(tree_frame, columns=cols, show='headings', selectmode='browse')
vsb = ttk.Scrollbar(tree_frame, orient='vertical', command=tree.yview)
tree.configure(yscrollcommand=vsb.set)
vsb.pack(side='right', fill='y')
tree.pack(side='left', fill='both', expand=True)

# Headings
tree.heading('full_name', text='Full Name')
tree.heading('username', text='Username')
tree.heading('password', text='Password')
tree.heading('pc_name', text='PC Name')
tree.heading('email', text='Email')
tree.heading('ext', text='Ext')
tree.heading('ddi', text='DDI')

# Column sizes chosen to avoid horizontal scrollbar in the fixed window
tree.column('full_name', width=260, anchor='w')
tree.column('username', width=160, anchor='w')
tree.column('password', width=120, anchor='w')
tree.column('pc_name', width=140, anchor='w')
tree.column('email', width=240, anchor='w')
tree.column('ext', width=70, anchor='w')
tree.column('ddi', width=120, anchor='w')

def populate_tree(records):
    tree.delete(*tree.get_children())
    for i, r in enumerate(records):
        tree.insert('', 'end', iid=str(i), values=(r.get('full_name',''), r.get('username',''),
                                                  r.get('password',''), r.get('pc_name',''), r.get('email',''),
                                                  r.get('ext',''), r.get('ddi','')))

def refresh_data():
    try:
        populate_tree(load_records())
    except Exception as e:
        messagebox.showerror("Load", f"Could not load data: {e}")

def on_tree_select(event=None):
    selected = tree.focus()   # get item ID
    if not selected:
        return

    values = tree.item(selected, "values")
    if not values or len(values) < 7:
        return

    # unpack in SAME order as Treeview columns
    full, user, pwd, pc, email, ext, ddi = values

    # populate fields
    ent_full.delete(0, tk.END); ent_full.insert(0, full)
    ent_pc.delete(0, tk.END); ent_pc.insert(0, pc)
    ent_pass.delete(0, tk.END); ent_pass.insert(0, pwd)
    ent_ext.delete(0, tk.END); ent_ext.insert(0, ext)
    ent_ddi.delete(0, tk.END); ent_ddi.insert(0, ddi)

    # username/email labels
    lbl_username.config(text=f"Username: {user}")
    lbl_email.config(text=f"Email: {email}")

    # update preview
    show_preview()


tree.bind('<<TreeviewSelect>>', on_tree_select)

# Bottom DB buttons row (Delete | Refresh | Locate CSV)
btns_db = ttk.Frame(bottom_frame)
btns_db.pack(fill='x', pady=(8,0))
btns_db.columnconfigure(0, weight=1)
btns_db.columnconfigure(1, weight=1)
btns_db.columnconfigure(2, weight=1)

def delete_record():
    sel = tree.selection()
    if not sel:
        messagebox.showinfo('No selection', 'Select a record to delete.')
        return
    if not messagebox.askyesno('Confirm delete', 'Delete selected record?'):
        return
    idx = int(sel[0])
    recs = load_records()
    if 0 <= idx < len(recs):
        del recs[idx]
        save_records(recs)
    refresh_data()
    clear_form()

def do_search(event=None):
    q = ent_search.get().strip().lower()
    recs = load_records()
    if q:
        filtered = []
        for r in recs:
            # guard against missing fields
            fn = (r.get('full_name') or '').lower()
            un = (r.get('username') or '').lower()
            pc = (r.get('pc_name') or '').lower()
            if q in fn or q in un or q in pc:
                filtered.append(r)
        recs = filtered
    populate_tree(recs)

btn_find.config(command=do_search)

# Live search: call do_search on typing and on Enter
ent_search.bind('<KeyRelease>', lambda e: do_search())
ent_search.bind('<Return>', lambda e: do_search())

btn_delete = ttk.Button(btns_db, text='Delete', command=delete_record)
btn_delete.grid(row=0, column=0, sticky='we', padx=6, pady=4)
btn_refresh = ttk.Button(btns_db, text='Refresh', command=refresh_data)
btn_refresh.grid(row=0, column=1, sticky='we', padx=6, pady=4)

def locate_csv():
    global DB_FILE
    p = filedialog.askopenfilename(title="Locate CSV database", filetypes=[("CSV files","*.csv"), ("All files","*.*")])
    if not p:
        return
    newpath = pathlib.Path(p)
    try:
        if not newpath.exists() or newpath.stat().st_size == 0:
            with newpath.open("w", newline="", encoding="utf-8") as f:
                writer = csv.DictWriter(f, FIELDS)
                writer.writeheader()
        DB_FILE = newpath
        ensure_db(DB_FILE)
        save_config()  # persist the user's choice
        refresh_data()
        messagebox.showinfo("Locate CSV", f"Using database: {DB_FILE}")
    except Exception as e:
        messagebox.showerror("Locate CSV", f"Could not use file: {e}")

btn_locate = ttk.Button(btns_db, text='Locate CSV', command=locate_csv)
btn_locate.grid(row=0, column=2, sticky='we', padx=6, pady=4)

# Double-click copies the whole selected record preview to clipboard (HTML if possible)
def on_tree_double(event):
    sel = tree.selection()
    if not sel:
        return
    idx = int(sel[0])
    recs = load_records()
    if 0 <= idx < len(recs):
        rec = recs[idx]
        txt = format_email_text(rec)
        # produce HTML fragment for the whole record, preserving headings
        lines = txt.splitlines(True)
        html_fragment = "<div>" + format_html_lines_for_fragment(lines) + "</div>"
        set_html_clipboard_fragment(html_fragment, plain_text=txt)
        messagebox.showinfo("Copied", "Selected record preview copied to clipboard (HTML if supported).")

tree.bind("<Double-1>", on_tree_double)

# Initial population and preview
refresh_data()
update_derived()
show_preview()

"""# Set initial sash position to visually match screenshot proportions
root.update_idletasks()
try:
    # place the top_paned sash so left size is roughly 36% (form) and right 64% (preview)
    total_width = top_paned.winfo_width()
    left_width = int(total_width * 0.37)
    top_paned.sash_place(0, left_width, 10)
    # vertical paned: give top area more room
    total_height = paned.winfo_height()
    top_height = int(total_height * 0.60)
    paned.sash_place(0, 10, top_height)
except Exception:
    pass
"""

root.mainloop()
