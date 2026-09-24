# ============================================================
#   PASSWORD MANAGER  (Python + Excel + Fernet)
#   Storage : passwords.xlsx  (Sheet: Passwords)
#   Key     : derived from a single Master Code (never stored)
# ============================================================

from cryptography.fernet import Fernet, InvalidToken
import base64
import hashlib
import getpass
from pathlib import Path
from openpyxl import Workbook, load_workbook
import subprocess
import os




# ---------- Constants ----------
EXCEL_FILE = "passwords.xlsx"
BACKUP_FILE = r"C:\Users\dilsh\iCloudDrive\Managing_Pass\passwords.xlsx"
SHEET_NAME = "Passwords"
HEADERS = ["Serial No", "Platform", "Username", "Hash Password"]


# ============================================================
#   KEY DERIVATION
# ============================================================
def derive_key(master_code: str) -> bytes:
    """Turn any Master Code text into a valid Fernet key (32 bytes, base64)."""
    raw = hashlib.sha256(master_code.encode()).digest()   # 32 raw bytes
    return base64.urlsafe_b64encode(raw)                  # base64 for Fernet


def encrypt_password(plain: str, master_code: str) -> str:
    """Encrypt a plain password -> string safe to store in Excel."""
    key = derive_key(master_code)
    f = Fernet(key)
    return f.encrypt(plain.encode()).decode()


def decrypt_password(token: str, master_code: str) -> str | None:
    """Decrypt a stored token. Returns None if Master Code is wrong."""
    key = derive_key(master_code)
    f = Fernet(key)
    try:
        return f.decrypt(token.encode()).decode()
    except InvalidToken:
        return None

def copy_to_clipboard(text: str) -> bool:
    """Copy given text to Windows clipboard using clip command."""
    try:
        p = subprocess.Popen("clip", stdin=subprocess.PIPE, shell=True)
        p.communicate(input=text.encode("utf-16le"))
        return p.returncode == 0
    except Exception:
        return False




# ============================================================
#   EXCEL HELPERS
# ============================================================
def save_workbook(wb) -> None:
    """Save to primary file AND iCloud backup simultaneously."""
    wb.save(EXCEL_FILE)
    try:
        backup_path = Path(BACKUP_FILE)
        if backup_path.parent and not backup_path.parent.exists():
            backup_path.parent.mkdir(parents=True, exist_ok=True)
        wb.save(BACKUP_FILE)
    except Exception as e:
        print(f"  ! Warning: primary saved, but backup save failed: {e}")


def sync_files_on_load() -> None:
    """Keep primary and backup in sync (newer wins, missing copied)."""
    primary = Path(EXCEL_FILE)
    backup = Path(BACKUP_FILE)
    try:
        if not primary.exists() and not backup.exists():
            return  # both missing -> will be created fresh
        if primary.exists() and not backup.exists():
            if backup.parent and not backup.parent.exists():
                backup.parent.mkdir(parents=True, exist_ok=True)
            wb_tmp = load_workbook(EXCEL_FILE)
            wb_tmp.save(BACKUP_FILE)
            print(f"[info] Backup created: {BACKUP_FILE}")
            return
        if backup.exists() and not primary.exists():
            wb_tmp = load_workbook(BACKUP_FILE)
            wb_tmp.save(EXCEL_FILE)
            print(f"[info] Restored primary from backup: {BACKUP_FILE}")
            return
        # both exist -> newer wins
        p_time = primary.stat().st_mtime
        b_time = backup.stat().st_mtime
        if abs(p_time - b_time) < 1:
            return  # same age, assume in sync
        if b_time > p_time:
            wb_tmp = load_workbook(BACKUP_FILE)
            wb_tmp.save(EXCEL_FILE)
            print("[info] Primary updated from newer backup.")
        else:
            wb_tmp = load_workbook(EXCEL_FILE)
            wb_tmp.save(BACKUP_FILE)
            print("[info] Backup updated from newer primary.")
    except Exception as e:
        print(f"  ! Warning: sync failed: {e}")


def load_or_create_workbook():
    """Open passwords.xlsx, creating it (with headers) if missing."""
    sync_files_on_load()
    if not Path(EXCEL_FILE).exists():
        wb = Workbook()
        ws = wb.active
        ws.title = SHEET_NAME
        ws.append(HEADERS)
        save_workbook(wb)
        print(f"[info] Created new file: {EXCEL_FILE}")
        return wb

    wb = load_workbook(EXCEL_FILE)
    if SHEET_NAME not in wb.sheetnames:
        ws = wb.create_sheet(SHEET_NAME)
        ws.append(HEADERS)
        save_workbook(wb)
    return wb


def find_row(ws, serial: int):
    """Return the row index (1-based) whose Serial No == serial, else None."""
    for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        try:
            if row[0] is not None and int(row[0]) == int(serial):
                return i
        except (ValueError, TypeError):
            continue
    return None


def get_next_serial(ws) -> int:
    """Auto-assign next Serial No as max(existing) + 1, or 1 if empty."""
    max_serial = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        try:
            if row[0] is not None and str(row[0]).strip() != "":
                val = int(row[0])
                if val > max_serial:
                    max_serial = val
        except (ValueError, TypeError):
            continue
    return max_serial + 1


def append_entry(ws, serial, platform, username, hash_password):
    ws.append([serial, platform, username, hash_password])


# ============================================================
#   INPUT HELPERS
# ============================================================
def ask_int(prompt: str) -> int | None:
    """Ask for an integer. Returns None if user types 'back' to cancel."""
    while True:
        val = input(prompt).strip()
        if val.lower() in ('back', 'cancel', 'exit'):
            return None
        if val.isdigit():
            return int(val)
        print("  ! Please enter a valid number or type 'back' to cancel.")


def ask_text(prompt: str, allow_empty: bool = False) -> str | None:
    """Ask for text. Returns None if user types 'back' to cancel."""
    while True:
        val = input(prompt).strip()
        if val.lower() in ('back', 'cancel', 'exit'):
            return None
        if val or allow_empty:
            return val
        print("  ! This field cannot be empty or type 'back' to cancel.")


# ============================================================
#   SESSION MASTER CODE (ask once, reuse, lockable)
# ============================================================
SESSION_MASTER = None


def session_login():
    """First-time login: set master (new file) or verify against data."""
    global SESSION_MASTER
    wb = load_or_create_workbook()
    ws = wb[SHEET_NAME]
    has_data = any(
        r[3] is not None and str(r[3]).strip() != ""
        for r in ws.iter_rows(min_row=2, values_only=True)
    )
    if not has_data:
        print("\n===== PASSWORD MANAGER =====")
        print("  No passwords saved yet. Set your Master Code.")
        while True:
            m1 = ask_secret("Set Master Code     : ")
            if m1 is None:
                print("  Master Code is required to continue.")
                continue
            m2 = ask_secret("Confirm Master Code : ")
            if m2 is None:
                print("  Master Code is required to continue.")
                continue
            if m1 != m2:
                print("  ! Codes do not match. Try again.")
                continue
            SESSION_MASTER = m1
            print("  ✔ Master Code set for this session.")
            return
    else:
        print("\n===== PASSWORD MANAGER =====")
        while True:
            m = ask_secret("Enter Master Code (session login): ")
            if m is None:
                print("  Master Code is required to continue.")
                continue
            # verify against first non-empty row
            ok = False
            for r in ws.iter_rows(min_row=2, values_only=True):
                if r[3] is not None and str(r[3]).strip() != "":
                    if decrypt_password(str(r[3]), m) is not None:
                        ok = True
                    break
            if ok:
                SESSION_MASTER = m
                print("  ✔ Unlocked for this session.")
                return
            print("  ! Invalid Master Code. Try again.")


def require_unlocked():
    """If locked, prompt for master until correct."""
    global SESSION_MASTER
    if SESSION_MASTER is not None:
        return
    print("\n  [Locked] Enter Master Code to unlock.")
    while True:
        m = ask_secret("Master Code : ")
        if m is None:
            print("  Still locked.")
            continue
        wb = load_or_create_workbook()
        ws = wb[SHEET_NAME]
        rows = [r for r in ws.iter_rows(min_row=2, values_only=True)
                if r[3] is not None and str(r[3]).strip() != ""]
        if not rows:
            # No entries yet: same as first-time setup, ask Set + Confirm
            m2 = ask_secret("Confirm Master Code : ")
            if m2 is None:
                print("  Still locked.")
                continue
            if m != m2:
                print("  ! Codes do not match. Try again.")
                continue
            SESSION_MASTER = m
            print("  ✔ Unlocked.")
            return
        if decrypt_password(str(rows[0][3]), m) is not None:
            SESSION_MASTER = m
            print("  ✔ Unlocked.")
            return
        print("  ! Invalid Master Code. Try again.")


def lock_session():
    """Lock the app (clear session master from memory)."""
    global SESSION_MASTER
    SESSION_MASTER = None
    print("  [Locked]. Master Code required to continue.")


def ask_secret(prompt: str) -> str | None:
    """Ask for a secret. Returns None if user types 'back' to cancel."""
    while True:
        val = getpass.getpass(prompt).strip()
        if val.lower() in ('back', 'cancel', 'exit'):
            return None
        if val:
            return val
        print("  ! This field cannot be empty or type 'back' to cancel.")


# ============================================================
#   FEATURE 1 : ADD PASSWORD  (Serial No auto-assigned)
# ============================================================
def add_password():
    require_unlocked()
    print("\n--- ADD PASSWORD ---")
    print("  (Type 'back' at any prompt to return to main menu)")

    wb = load_or_create_workbook()
    ws = wb[SHEET_NAME]

    serial = get_next_serial(ws)
    print(f"  Serial No (auto) : {serial}")

    platform = ask_text("Platform Name  : ")
    if platform is None:
        print("  Cancelled. Returning to main menu.")
        return

    username = ask_text("Username       : ")
    if username is None:
        print("  Cancelled. Returning to main menu.")
        return

    password = ask_secret("Password       : ")
    if password is None:
        print("  Cancelled. Returning to main menu.")
        return

    # Reuse session master; confirm only if user wants a different one
    print(f"  Using session Master Code. (type 'change' to use a different one, Enter to keep)")
    master = SESSION_MASTER
    choice = ask_text("Master Code [Enter=keep session] : ", allow_empty=True)
    if choice is None:
        print("  Cancelled. Returning to main menu.")
        return
    if choice != "" and choice.lower() != "change":
        master = choice
    elif choice.lower() == "change":
        new_m = ask_secret("New Master Code: ")
        if new_m is None:
            print("  Cancelled. Returning to main menu.")
            return
        master = new_m

    hash_password = encrypt_password(password, master)
    # verify round-trip before saving (catches any crypto error)
    if decrypt_password(hash_password, master) != password:
        print("  ! Encryption verify failed. Nothing saved.")
        return

    append_entry(ws, serial, platform, username, hash_password)

    save_workbook(wb)
    print(f"  ✔ Password saved with Serial No {serial}.")

    # Allow user to navigate back to main menu
    while True:
        back_choice = input("\n  Press Enter to return to main menu or type 'add' to add another password: ").strip().lower()
        if back_choice == "" or back_choice == "back":
            return
        elif back_choice == "add":
            # Allow adding another password without returning to main menu
            add_password()
            return
        else:
            print("  ! Invalid choice. Press Enter to go back.")


# ============================================================
#   FEATURE 2 : VIEW PASSWORD
# ============================================================
def view_password():
    require_unlocked()
    print("\n--- VIEW PASSWORD ---")
    print("  (Type 'back' at any prompt to return to main menu)")
    serial = ask_int("Serial No      : ")
    if serial is None:
        print("  Cancelled. Returning to main menu.")
        return

    master = SESSION_MASTER

    wb = load_or_create_workbook()
    ws = wb[SHEET_NAME]

    row_idx = find_row(ws, serial)
    if row_idx is None:
        print("  ! No entry found with that Serial No.")
        # Allow user to navigate back to main menu
        while True:
            back_choice = input("\n  Press Enter to return to main menu or type 'view' to view another password: ").strip().lower()
            if back_choice == "" or back_choice == "back":
                return
            elif back_choice == "view":
                view_password()
                return
            else:
                print("  ! Invalid choice. Press Enter to go back.")
        return

    row = ws[row_idx]
    platform      = row[1].value
    username      = row[2].value
    hash_password = row[3].value

    plain = decrypt_password(hash_password, master)
    if plain is None:
        print("  ! Invalid Master Code.")
        # Allow user to navigate back to main menu
        while True:
            back_choice = input("\n  Press Enter to return to main menu or type 'view' to try again: ").strip().lower()
            if back_choice == "" or back_choice == "back":
                return
            elif back_choice == "view":
                view_password()
                return
            else:
                print("  ! Invalid choice. Press Enter to go back.")
        return

    # Badge formatting styles: Pure White background with True pitch-black text
    # (Note: ANSI 'bold black' turns gray in Windows Terminal; true RGB 0;0;0 guarantees deep pure black)
    WHITE_BADGE = "\033[48;2;255;255;255;38;2;0;0;0m"
    RESET_STYLE = "\033[0m"

    raw_lines = [
        ("Serial No : ", str(serial), None),
        ("Platform  : ", str(platform), None),
        ("Username  : ", str(username), WHITE_BADGE),
        ("Password  : ", str(plain), WHITE_BADGE),
    ]

    card_items = []
    for label, val, style in raw_lines:
        if style:
            vis_text = f"{label}  {val}  "
            fmt_text = f"{label}{style}  {val}  {RESET_STYLE}"
        else:
            vis_text = f"{label}{val}"
            fmt_text = f"{label}{val}"
        card_items.append((vis_text, fmt_text))

    max_vis = max(max(len(v) for v, _ in card_items), 34)
    width = max_vis + 6

    header_text = " CREDENTIAL CARD "
    dash_total = width - len(header_text)
    left_dash = dash_total // 2
    right_dash = dash_total - left_dash

    print()
    print("  +" + "-" * width + "+")
    print("  |" + (" " * left_dash) + header_text + (" " * right_dash) + "|")
    print("  +" + "-" * width + "+")
    print("  |" + (" " * width) + "|")
    for i, (vis_text, fmt_text) in enumerate(card_items):
        padding = width - len(vis_text) - 4
        print(f"  |    {fmt_text}" + (" " * padding) + "|")
        if i in (1, 2):  # clean gap between platform & username, and between username & password
            print("  |" + (" " * width) + "|")
    print("  |" + (" " * width) + "|")
    print("  +" + "-" * width + "+")

    # --- View-screen options: Delete / Update / Copy ---
    while True:
        print()
        print("  +---------------------------------------------------------------------------+")
        print("  |  [1] Delete |  [2] Update |  [3] Copy  |  [view] Another |  [Enter] Back  |")
        print("  +---------------------------------------------------------------------------+")
        back_choice = input("  Enter your choice: ").strip().lower()
        if back_choice == "" or back_choice == "back":
            return
        elif back_choice == "view":
            view_password()
            return
        elif back_choice == "1":
            confirm = input(f"  ! Delete entry Serial No {serial}? (y/n): ").strip().lower()
            if confirm == "y":
                wb2 = load_or_create_workbook()
                ws2 = wb2[SHEET_NAME]
                row_to_del = find_row(ws2, serial)
                if row_to_del is None:
                    print("  ! Entry no longer exists.")
                else:
                    ws2.delete_rows(row_to_del, 1)
                    save_workbook(wb2)
                    print(f"  ✔ Entry Serial No {serial} deleted.")
                return
            else:
                print("  Cancelled. Entry not deleted.")
                continue
        elif back_choice == "2":
            update_password_entry(serial)
            return
        elif back_choice == "3":
            if copy_to_clipboard(plain):
                print("  ✔ Password copied to clipboard successfully!")
            else:
                print("  ! Could not copy to clipboard.")
            continue
        else:
            print("  ! Invalid choice.")


# ============================================================
#   FEATURE 2b : UPDATE ENTRY (called from View screen)
#   Updates Username and/or Password (press Enter to keep old)
# ============================================================
def update_password_entry(serial: int):
    require_unlocked()
    print(f"\n--- UPDATE ENTRY (Serial No {serial}) ---")
    print("  (Press Enter to keep current value, or type 'back' to cancel)")

    wb = load_or_create_workbook()
    ws = wb[SHEET_NAME]
    row_idx = find_row(ws, serial)
    if row_idx is None:
        print("  ! No entry found with that Serial No.")
        return

    current_platform = ws.cell(row=row_idx, column=2).value or ""
    current_username = ws.cell(row=row_idx, column=3).value or ""

    print(f"  Platform         : {current_platform}")
    print(f"  Current Username : {current_username}")

    # 1. Update Username
    while True:
        user_input = input("  New Username     : ").strip()
        if user_input.lower() in ('back', 'cancel', 'exit'):
            print("  Cancelled. Nothing changed.")
            return
        if user_input == "":
            new_username = current_username
            break
        else:
            new_username = user_input
            break

    # 2. Update Password
    while True:
        pass_input = getpass.getpass("  New Password     : ").strip()
        if pass_input.lower() in ('back', 'cancel', 'exit'):
            print("  Cancelled. Nothing changed.")
            return
        if pass_input == "":
            new_password = None  # keep current password
            break
        else:
            new_password = pass_input
            break

    # Check if anything changed
    changed = False

    if new_username != current_username:
        ws.cell(row=row_idx, column=3, value=new_username)
        changed = True

    if new_password is not None:
        new_master = SESSION_MASTER
        new_token = encrypt_password(new_password, new_master)
        if decrypt_password(new_token, new_master) != new_password:
            print("  ! Encryption verify failed. Nothing saved.")
            return
        ws.cell(row=row_idx, column=4, value=new_token)
        changed = True

    if changed:
        save_workbook(wb)
        print(f"  ✔ Entry Serial No {serial} updated successfully.")
    else:
        print("  No changes made.")


# ============================================================
#   FEATURE 3 : LIST ALL PASSWORDS
# ============================================================
def list_passwords():
    print("\n" + "=" * 70)
    print("  LIST ALL PASSWORDS")
    print("  (Shows Serial No, Platform, and Username only)")
    print("=" * 70)

    wb = load_or_create_workbook()
    ws = wb[SHEET_NAME]

    # Get all rows except header
    rows = list(ws.iter_rows(min_row=2, values_only=True))

    if not rows:
        print()
        print("  ! No passwords saved yet.")
        print()
    else:
        # Define column widths
        col1_width = 12  # Serial No
        col2_width = 25  # Platform
        col3_width = 25  # Username

        # Display header row
        print()
        header = f"  {'Serial No':<{col1_width}} | {'Platform':<{col2_width}} | {'Username':<{col3_width}}"
        print(header)
        print("  " + "-" * (col1_width + col2_width + col3_width + 4))  # +4 for " | " separators

        # Display each entry
        for row in rows:
            serial = row[0]
            platform = row[1] if row[1] else "N/A"
            username = row[2] if row[2] else "N/A"
            print(f"  {str(serial):<{col1_width}} | {str(platform):<{col2_width}} | {str(username):<{col3_width}}")

        print()
        print("  " + "-" * (col1_width + col2_width + col3_width + 4))
        print(f"  Total entries: {len(rows)}")
        print()

    print("=" * 70)

    # Allow user to navigate: 1=Main menu, 2=Refresh, 3=View by Serial No
    while True:
        print()
        print("  +---------------------------------------------------------------------------+")
        print("  |  [Enter/1] Main Menu      |  [2/list] Refresh List    |  [3] View by ID   |")
        print("  +---------------------------------------------------------------------------+")
        back_choice = input("  Enter your choice: ").strip().lower()
        if back_choice == "" or back_choice == "1" or back_choice == "back":
            return
        elif back_choice == "2" or back_choice == "list":
            list_passwords()
            return
        elif back_choice == "3":
            view_password()
            return
        else:
            print("  ! Invalid choice. Enter 1, 2, 3, Enter, or 'list'.")


# ============================================================
#   FEATURE 4 : UPDATE MASTER CODE
# ============================================================
def update_master_code_menu():
    while True:
        print()
        print("  +---------------------------------------------------------------------------+")
        print("  |                            UPDATE MASTER CODE                             |")
        print("  +---------------------------------------------------------------------------+")
        print("  |  [1] For All Passwords    |  [2] For Specific Serial  |  [3] Back         |")
        print("  +---------------------------------------------------------------------------+")
        choice = input("  Enter your choice: ").strip()

        if choice == "1":
            update_master_code_all()
        elif choice == "2":
            update_master_code_one()
        elif choice == "3" or choice.lower() == "back":
            return
        else:
            print("  ! Invalid choice.")


def update_master_code_all():
    global SESSION_MASTER
    require_unlocked()
    print("\n--- UPDATE MASTER CODE : FOR ALL ---")
    print("  (Type 'back' at any prompt to cancel)")
    old_master = SESSION_MASTER
    print("  Old Master Code: [using session login]")
    new_master = ask_secret("New Master Code : ")
    if new_master is None:
        print("  Cancelled.")
        return
    confirm = ask_secret("Confirm New     : ")
    if confirm is None:
        print("  Cancelled.")
        return
    if new_master != confirm:
        print("  ! New codes do not match. Aborted.")
        return

    wb = load_or_create_workbook()
    ws = wb[SHEET_NAME]

    rows = list(ws.iter_rows(min_row=2))
    if not rows:
        print("  ! No passwords saved yet.")
        return

    # First verify old master decrypts every row (abort if any fails)
    plains = []
    for r in rows:
        token = r[3].value
        if token is None or str(token).strip() == "":
            print(f"  ! Serial No {r[0].value} has empty password. Aborted.")
            return
        plain = decrypt_password(str(token), old_master)
        if plain is None:
            print(f"  ! Invalid Old Master Code (failed at Serial No {r[0].value}). Aborted, nothing changed.")
            return
        plains.append(plain)

    # Re-encrypt all with new master
    for r, plain in zip(rows, plains):
        r[3].value = encrypt_password(plain, new_master)

    save_workbook(wb)
    # new master becomes the session master
    SESSION_MASTER = new_master
    print(f"  ✔ Master Code updated for all {len(rows)} entries. Session updated.")


def update_master_code_one():
    require_unlocked()
    print("\n--- UPDATE MASTER CODE : FOR SPECIFIC SERIAL NO ---")
    print("  (Type 'back' at any prompt to cancel)")
    serial = ask_int("Serial No         : ")
    if serial is None:
        print("  Cancelled.")
        return
    old_master = SESSION_MASTER
    print("  Old Master Code: [using session login]")
    new_master = ask_secret("New Master Code   : ")
    if new_master is None:
        print("  Cancelled.")
        return
    confirm = ask_secret("Confirm New       : ")
    if confirm is None:
        print("  Cancelled.")
        return
    if new_master != confirm:
        print("  ! New codes do not match. Aborted.")
        return

    wb = load_or_create_workbook()
    ws = wb[SHEET_NAME]
    row_idx = find_row(ws, serial)
    if row_idx is None:
        print("  ! No entry found with that Serial No.")
        return

    token = ws.cell(row=row_idx, column=4).value
    plain = decrypt_password(str(token), old_master)
    if plain is None:
        print("  ! Invalid Old Master Code. Nothing changed.")
        return

    ws.cell(row=row_idx, column=4, value=encrypt_password(plain, new_master))
    save_workbook(wb)
    print(f"  ✔ Master Code updated for Serial No {serial}.")


# ============================================================
#   MAIN MENU
# ============================================================
def main():
    global SESSION_MASTER
    os.system("")  # Enable ANSI color escape sequences on Windows console
    session_login()
    while True:
        print()
        print("  +---------------------------------------------------------------------------+")
        print("  |                             PASSWORD MANAGER                              |")
        print("  +---------------------------------------------------------------------------+")
        print("  |  [1] Add Password       |  [2] View Password       |  [3] List Passwords  |")
        print("  |  [4] Master Code        |  [5] Lock Session        |  [6] Exit            |")
        print("  +---------------------------------------------------------------------------+")
        choice = input("  Enter your choice: ").strip()

        if choice == "1":
            add_password()
        elif choice == "2":
            view_password()
        elif choice == "3":
            require_unlocked()
            list_passwords()
        elif choice == "4":
            update_master_code_menu()
        elif choice == "5":
            lock_session()
            require_unlocked()
        elif choice == "6":
            SESSION_MASTER = None
            print("Bye. (session cleared)")
            break
        else:
            print("  ! Invalid choice.")


if __name__ == "__main__":
    main()