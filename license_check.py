# ============================================================
#  LICENSE CHECKER
#  Bundled with the tool. Safe to share — cannot generate keys.
#  © 2024 Suderson Jagadeeshan
# ============================================================

import hashlib
import datetime
import os
import tkinter as tk
from tkinter import messagebox, simpledialog

LICENSE_FILE = "license.key"
CONTACT      = "suderson123@gmail.com"

# ── Same secret as generator (must match exactly) ────────────
_SECRET = "SJ@IEC61850#CID$COMPARE!2024_SUDERSON"

def _make_hash(user_code, expiry_str):
    raw  = f"{user_code}|{expiry_str}|{_SECRET}"
    full = hashlib.sha256(raw.encode()).hexdigest().upper()
    return f"{full[:4]}-{full[4:8]}"

def validate_key(key):
    """
    Validate a license key.
    Returns (True, days_remaining, expiry_date) or (False, reason, None)
    """
    key = key.strip().upper()

    # ── Format check ──────────────────────────────────────────
    # Expected: SJ-USERCODE-XXXX-XXXX-EXPyyyymmdd
    # Example:  SJ-SUDERS-05B5-EB7F-EXP20260529
    parts = key.split('-')
    if len(parts) != 5:
        return False, "Invalid key format.", None
    if parts[0] != 'SJ':
        return False, "Invalid key format.", None
    if not parts[4].startswith('EXP') or len(parts[4]) != 11:
        return False, "Invalid key format.", None

    user_code  = parts[1]
    hash_part  = f"{parts[2]}-{parts[3]}"
    expiry_str = parts[4][3:]   # strip 'EXP'

    # ── Date check ────────────────────────────────────────────
    try:
        expiry = datetime.datetime.strptime(expiry_str, '%Y%m%d').date()
    except ValueError:
        return False, "Invalid expiry date in key.", None

    today = datetime.date.today()
    if today > expiry:
        days_over = (today - expiry).days
        return False, f"License expired {days_over} day(s) ago ({expiry.strftime('%d %b %Y')}).", expiry

    # ── Hash check (tamper detection) ─────────────────────────
    expected_hash = _make_hash(user_code, expiry_str)
    if hash_part != expected_hash:
        return False, "License key is invalid or has been tampered with.", None

    days_remaining = (expiry - today).days
    return True, days_remaining, expiry


def load_saved_key():
    """Load key from license.key file if it exists."""
    if os.path.exists(LICENSE_FILE):
        with open(LICENSE_FILE, 'r') as f:
            return f.read().strip()
    return None


def save_key(key):
    """Save key to license.key file."""
    with open(LICENSE_FILE, 'w') as f:
        f.write(key.strip())


def prompt_for_key(parent=None):
    """Show a dialog asking the user to enter their license key."""
    # Always create a fresh Tk root for the license dialog
    win = tk.Tk()
    win.title("License Activation")
    win.resizable(False, False)

    # Centre the window
    w, h = 480, 280
    win.update_idletasks()
    x = (win.winfo_screenwidth()  - w) // 2
    y = (win.winfo_screenheight() - h) // 2
    win.geometry(f"{w}x{h}+{x}+{y}")
    win.lift()
    win.focus_force()

    result = {'key': None}

    tk.Label(win, text="IEC 61850 CID / SCD Comparison Tool",
             font=("Arial", 11, "bold")).pack(pady=(18, 2))
    tk.Label(win, text="Please enter your license key to continue.",
             font=("Arial", 9)).pack()
    tk.Label(win, text=f"Contact {CONTACT} to obtain a key.",
             font=("Arial", 9, "italic"), fg="gray").pack(pady=(2, 12))

    entry_var = tk.StringVar()
    entry = tk.Entry(win, textvariable=entry_var, width=44,
                     font=("Courier", 10), justify='center')
    entry.pack(padx=20)
    entry.focus()

    msg_var = tk.StringVar()
    msg_label = tk.Label(win, textvariable=msg_var,
                         font=("Arial", 9), fg="red")
    msg_label.pack(pady=(6, 0))

    def on_activate():
        key = entry_var.get().strip()
        if not key:
            msg_var.set("Please enter a license key.")
            return
        valid, info, expiry = validate_key(key)
        if valid:
            save_key(key)
            result['key'] = key
            win.destroy()
        else:
            msg_var.set(f"❌  {info}")

    tk.Button(win, text="Activate", command=on_activate,
              bg="#4CAF50", fg="white", font=("Arial", 10, "bold"),
              width=14).pack(pady=(10, 4))

    tk.Label(win, text="© 2024 Suderson Jagadeeshan",
             font=("Arial", 8), fg="gray").pack()

    win.protocol("WM_DELETE_WINDOW", lambda: exit(0))
    win.wait_window()

    return result['key']


def show_expiry_warning(days_remaining, expiry):
    """Show a warning when license is close to expiry (≤ 3 days)."""
    if days_remaining <= 3:
        messagebox.showwarning(
            "License Expiring Soon",
            f"Your license expires in {days_remaining} day(s) "
            f"({expiry.strftime('%d %b %Y')}).\n\n"
            f"Please contact {CONTACT} to renew."
        )


def show_expired_dialog(reason):
    """Show expiry/invalid dialog and exit."""
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror(
        "License Error",
        f"{reason}\n\nPlease contact:\n{CONTACT}\nto obtain or renew your license."
    )
    root.destroy()
    exit(0)


def check_license():
    """
    Main entry point. Call this at the top of cid_gui.py before
    building the main window.

    Returns True if license is valid. Never returns False —
    exits the application instead.
    """
    key = load_saved_key()

    # No saved key → prompt user to enter one
    if not key:
        key = prompt_for_key()
        if not key:
            exit(0)

    # Validate
    valid, info, expiry = validate_key(key)

    if not valid:
        # If expired, delete saved key so they're prompted again after renewal
        if os.path.exists(LICENSE_FILE):
            os.remove(LICENSE_FILE)
        show_expired_dialog(info)

    # Warn if expiring soon
    show_expiry_warning(info, expiry)   # info = days_remaining when valid=True

    return True
