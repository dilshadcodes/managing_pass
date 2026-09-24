# 🔐 Password Manager (Python + Excel + Fernet)

A lightweight, secure, and offline command-line Password Manager built in Python. Passwords are encrypted using symmetric AES encryption via **Fernet** (`cryptography`) and securely stored in an Excel workbook (`passwords.xlsx`) with automated cloud backup sync support.

---

## ✨ Features

- **Robust Encryption**: Passwords are encrypted using **Fernet (AES-128-CBC with SHA-256 HMAC)**.
- **Zero-Storage Master Key**: Your Master Code is never written to disk or stored anywhere in plaintext or hashed form; encryption keys are derived on the fly in memory.
- **Excel Spreadsheet Storage**: Passwords and metadata are saved neatly to `passwords.xlsx` using `openpyxl`.
- **Automatic Cloud Backup Sync**: Automatically syncs with a configured cloud folder (e.g., iCloud Drive) on start and save, keeping the newest copy synchronized.
- **Credential Card Display**: View passwords in formatted high-contrast credential cards.
- **One-Click Clipboard Copy**: Instantly copy passwords to the Windows clipboard (`clip`).
- **Update & Delete**: Easily edit usernames, passwords, or remove entries directly from the view screen.
- **Session Lock & Unlock**: Lock the app session anytime without closing the terminal.
- **Master Key Rotation**: Re-encrypt a single entry or all saved entries under a new Master Code.

---

## 📋 Requirements

- **Python 3.10+** (recommended)
- Windows OS (for clipboard copy integration and ANSI support)
- Python packages:
  - `cryptography`
  - `openpyxl`

---

## 🚀 Getting Started

### 1. Clone the Repository
```bash
git clone https://github.com/dilshadcodes/managing_pass.git
cd managing_pass
```

### 2. Set Up a Virtual Environment
Create and activate a virtual environment:

- **Windows (PowerShell):**
  ```powershell
  python -m venv venv
  .\venv\Scripts\Activate.ps1
  ```

- **Windows (Command Prompt):**
  ```cmd
  python -m venv venv
  .\venv\Scripts\activate.bat
  ```

- **macOS / Linux:**
  ```bash
  python3 -m venv venv
  source venv/bin/activate
  ```

### 3. Install Dependencies
Install the required packages:
```bash
pip install cryptography openpyxl
```

### 4. Run the Application
```bash
python password_manager.py
```

---

## 📖 Usage Guide

### First-Time Launch
1. On your first run, if `passwords.xlsx` does not exist or is empty, the application will prompt you to **Set Master Code** and confirm it.
2. Remember this Master Code! Because it is never saved to disk, if you lose it, encrypted passwords cannot be recovered.

### Main Menu Options
When launched, you are greeted with the interactive menu:

| Option | Action | Description |
| :---: | :--- | :--- |
| **`[1]`** | **Add Password** | Auto-assigns an ID, prompts for platform, username, and password. |
| **`[2]`** | **View Password** | Look up by Serial No to view credentials, copy password to clipboard, edit, or delete. |
| **`[3]`** | **List Passwords** | Lists all saved platforms and usernames without revealing passwords. |
| **`[4]`** | **Master Code** | Rotate / update your master code for all entries or a single entry. |
| **`[5]`** | **Lock Session** | Locks the terminal session; requires re-entering the Master Code. |
| **`[6]`** | **Exit** | Clears session data from memory and exits safely. |

---

## 🔒 Security Architecture

1. **Key Derivation**:
   - Master code string $\to$ `SHA-256` digest (32 bytes) $\to$ URL-safe Base64 encoded Fernet key.
2. **Encrypted at Rest**:
   - Only the ciphertexts are stored inside the Excel file under the `Hash Password` column.
   - Master passwords or derivation salts are never written to disk.
3. **Round-Trip Verification**:
   - Every encryption operation runs an immediate test decryption verification before committing to file.

---

## 🛡️ .gitignore Configuration

Sensitive files like `passwords.xlsx`, backups, `.env` files, and virtual environments are explicitly excluded via `.gitignore` to prevent confidential data from leaking into public Git repositories.

---

## 📄 License

This project is open source and available under the [MIT License](LICENSE).
