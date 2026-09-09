# Folder Sync Comparison Tool

Compares two folder trees -- which may live on two different machines (local,
RDP, or an SFTP server), across Windows and Mac -- and produces an Excel
report showing what needs to be copied where to bring them back in sync.

This tool **only compares and reports** -- it never copies, deletes, or
modifies any files. You do the actual copying yourself, guided by the report.

## What it reports

1. Files/folders that exist in B but not A -> copy B -> A
2. Files/folders that exist in A but not B -> copy A -> B
3. Files that exist in both, where A's copy is newer -> copy A -> B
4. Files that exist in both, where B's copy is newer -> copy B -> A
5. Type conflicts (a file on one side, a folder of the same name on the other)
6. Name differences (same file matched across sides, but differing case or
   Unicode accent encoding)

All output goes into a single `folder_sync_report.xlsx` with a summary sheet
plus one sheet per category.

---

## Quick start: comparing an SFTP folder against a local folder

This is the setup where you're running an SFTP server on one machine (with a
root folder shared over SFTP) and running this script on a **different**
computer to compare a local folder against that SFTP root.

The client machine running this script can be **either Windows or Mac** --
pick the commands for your OS at each step below. (Windows commands use the
plain Command Prompt, `cmd.exe` -- not PowerShell.)

### 1. Install Python

You need Python 3.9 or newer, **on the computer where you'll run this
script** (the client machine, not the SFTP server).

**Windows:**
- Download the installer from [python.org/downloads](https://www.python.org/downloads/)
- During install, check the box **"Add python.exe to PATH"**
- Open Command Prompt (search "cmd" in the Start menu) and verify:
  ```bat
  python --version
  ```

**Mac:**
- Download the installer from [python.org/downloads](https://www.python.org/downloads/)
  (or, if you use [Homebrew](https://brew.sh/): `brew install python`)
- Open Terminal and verify:
  ```bash
  python3 --version
  ```

### 2. Get the code

**Windows (cmd):**
```bat
git clone https://github.com/aashil2495/folder_sync_tool.git
cd folder_sync_tool
```

**Mac (Terminal):**
```bash
git clone https://github.com/aashil2495/folder_sync_tool.git
cd folder_sync_tool
```

(No `git`? Instead click the green **Code** button on the GitHub page ->
**Download ZIP** -> extract it -> open a terminal in that folder.)

### 3. Create a virtual environment and install dependencies

A virtual environment keeps this project's Python packages separate from
everything else on your machine. Run these one at a time:

**Windows (cmd):**
```bat
python -m venv .venv
.venv\Scripts\activate.bat
pip install -r requirements.txt
```

**Mac (Terminal):**
```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

If it worked, your terminal prompt now starts with `(.venv)`.

### 4. Create your `.env` config file

**Windows (cmd):**
```bat
copy .env.example .env
```

**Mac (Terminal):**
```bash
cp .env.example .env
```

Open the new `.env` file in a text editor (Notepad/VS Code on Windows,
TextEdit/VS Code on Mac) and set these values:

```ini
# The SFTP server -- the machine that has the root folder you're comparing against
SFTP_HOST=192.168.1.50          # the SFTP server machine's IP address or hostname
SFTP_PORT=22
SFTP_USERNAME=your_sftp_username
SFTP_PASSWORD=your_sftp_password
SFTP_ROOT_FOLDER=/                # the folder on the SFTP server to compare -- "/" means its whole shared root

# The folder on THIS computer (the one running the script) to compare against it
# Windows example: LOCAL_ROOT_FOLDER=C:\Users\you\Documents\MyFolder
# Mac example:     LOCAL_ROOT_FOLDER=/Users/you/Documents/MyFolder
LOCAL_ROOT_FOLDER=C:\Users\you\Documents\MyFolder

# Optional: friendly names shown in the Excel report
LABEL_A=SFTP Server
LABEL_B=My Computer
```

Notes:
- `SFTP_HOST` is the IP address or hostname of the machine running the SFTP
  server -- not "localhost", unless the script happens to run on that same
  machine.
- `SFTP_ROOT_FOLDER` uses forward slashes (`/`) regardless of what OS the
  SFTP server itself runs on -- most SFTP servers present their shared root
  as `/`. Check your SFTP server's own documentation/settings if you're
  unsure what path to use here.
- `LOCAL_ROOT_FOLDER` uses whatever path style your client OS uses --
  backslashes and a drive letter on Windows (`C:\...`), forward slashes on
  Mac (`/Users/...`).
- If your SFTP server uses a private key instead of a password, leave
  `SFTP_PASSWORD` blank and set `SFTP_PRIVATE_KEY_PATH` to the key file path
  instead (e.g. `C:\Users\you\.ssh\id_rsa` on Windows, `/Users/you/.ssh/id_rsa`
  on Mac).
- **Never commit or share your `.env` file** -- it holds your real
  credentials. It's already excluded via `.gitignore`.

### 5. Run it

**Windows (cmd):**
```bat
python folder_sync_tool.py
```

**Mac (Terminal):**
```bash
python3 folder_sync_tool.py
```

You'll see progress printed to the terminal, then a line like:

```
Excel report written to /path/to/folder_sync_tool/folder_sync_report.xlsx
```

Open that file in Excel. Check the **Summary** sheet first for counts per
category, then the individual sheets for exactly which files to copy where.

### Troubleshooting

- **`ModuleNotFoundError`**: you forgot to activate the virtual environment
  or run `pip install -r requirements.txt`. Re-run step 3.
- **Authentication / connection errors**: double check `SFTP_HOST`,
  `SFTP_PORT`, `SFTP_USERNAME`, and `SFTP_PASSWORD`/`SFTP_PRIVATE_KEY_PATH` in
  `.env`. Confirm you can reach the SFTP server from this machine at all --
  e.g. test with an SFTP client like [WinSCP](https://winscp.net/) (Windows)
  or [Cyberduck](https://cyberduck.io/)/FileZilla (Mac) first, using the same
  host/port/credentials.
- **`NotADirectoryError`**: `LOCAL_ROOT_FOLDER` doesn't point at a real
  folder on this machine -- check the path, and make sure it uses the right
  slash direction for your OS.
- Nothing shows up as different but you expected changes: check
  `TIMESTAMP_TOLERANCE_SECONDS` in `.env` and make sure both machines' clocks
  are reasonably accurate.

---

## Cross-platform handling (automatic)

- **Case-insensitive matching**: `Report.pdf` and `report.pdf` are treated as
  the same file (both NTFS and APFS/HFS+ are case-insensitive by default).
  Exact-case differences are still called out on their own sheet.
- **Unicode normalization**: Mac has historically stored accented filenames
  (e.g. `café.txt`) in NFD (decomposed) form, Windows in NFC (composed).
  Paths are normalized before comparison so these aren't flagged as
  different files.
- **Junk file exclusion**: `.DS_Store`, `._*` AppleDouble sidecars,
  `Thumbs.db`, `desktop.ini`, `.git`, `__pycache__`, etc. are ignored
  automatically.
- **Timezone-safe timestamps**: comparisons use raw epoch mtimes, so it
  works even if the two machines are in different timezones.

## Configuration reference

All variables are read from `.env` (see `.env.example` for the full list
with defaults):

- `TIMESTAMP_TOLERANCE_SECONDS` -- files whose modified times differ by less
  than this are treated as identical (absorbs clock drift between machines).
  Default `2`.
- `CASE_INSENSITIVE_MATCH` -- match filenames regardless of case. Default
  `true`.
- `INCLUDE_IDENTICAL_SHEET` -- also list files identical on both sides as
  their own sheet. Default `false`.

## Requirements

- Python 3.9+
- `openpyxl` (Excel report generation)
- `paramiko` (SFTP support)
- `python-dotenv` (loads `.env`)

## Security note

Never commit `.env` or the generated `.xlsx` report if it contains sensitive
paths or data -- both are excluded by `.gitignore` by default.
