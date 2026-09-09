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
computer to compare a local folder against that SFTP root. This uses `MODE=sftp_compare`.

### 1. Install Python

You need Python 3.9 or newer, **on the computer where you'll run this
script** (the client machine, not the SFTP server).

- Download from [python.org/downloads](https://www.python.org/downloads/)
- During install on Windows, check the box **"Add python.exe to PATH"**
- Verify it worked by opening a terminal (PowerShell) and running:
  ```powershell
  python --version
  ```

### 2. Get the code

```powershell
git clone https://github.com/aashil2495/folder_sync_tool.git
cd folder_sync_tool
```

(No `git`? Instead click the green **Code** button on the GitHub page ->
**Download ZIP** -> extract it -> open a terminal in that folder.)

### 3. Create a virtual environment and install dependencies

A virtual environment keeps this project's Python packages separate from
everything else on your machine. Run these one at a time:

```powershell
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

If it worked, your terminal prompt now starts with `(.venv)`.

### 4. Create your `.env` config file

```powershell
copy .env.example .env
```

Open the new `.env` file in a text editor (Notepad, VS Code, whatever you
have) and set these values:

```ini
MODE=sftp_compare

# The SFTP server -- the machine that has the root folder you're comparing against
SFTP_HOST=192.168.1.50          # the SFTP server machine's IP address or hostname
SFTP_PORT=22
SFTP_USERNAME=your_sftp_username
SFTP_PASSWORD=your_sftp_password
SFTP_ROOT_FOLDER=/                # the folder on the SFTP server to compare -- "/" means its whole shared root

# The folder on THIS computer (the one running the script) to compare against it
LOCAL_ROOT_FOLDER=C:\Users\you\Documents\MyFolder

# Optional: friendly names shown in the Excel report
LABEL_A=SFTP Server
LABEL_B=My Computer
```

Notes:
- `SFTP_HOST` is the IP address or hostname of the machine running the SFTP
  server -- not "localhost", unless the script happens to run on that same
  machine.
- `SFTP_ROOT_FOLDER` uses forward slashes (`/`), even though this is a
  Windows path underneath -- most SFTP servers present their shared root as
  `/`. Check your SFTP server's own documentation/settings if you're unsure
  what path to use here.
- If your SFTP server uses a private key instead of a password, leave
  `SFTP_PASSWORD` blank and set `SFTP_PRIVATE_KEY_PATH` to the key file path
  instead (e.g. `C:\Users\you\.ssh\id_rsa`).
- **Never commit or share your `.env` file** -- it holds your real
  credentials. It's already excluded via `.gitignore`.

### 5. Run it

```powershell
python folder_sync_tool.py
```

You'll see progress printed to the terminal, then a line like:

```
Excel report written to C:\Users\you\folder_sync_tool\folder_sync_report.xlsx
```

Open that file in Excel. Check the **Summary** sheet first for counts per
category, then the individual sheets for exactly which files to copy where.

### Troubleshooting

- **`ModuleNotFoundError`**: you forgot to activate the virtual environment
  or run `pip install -r requirements.txt`. Re-run step 3.
- **Authentication / connection errors**: double check `SFTP_HOST`,
  `SFTP_PORT`, `SFTP_USERNAME`, and `SFTP_PASSWORD`/`SFTP_PRIVATE_KEY_PATH` in
  `.env`. Confirm you can reach the SFTP server from this machine at all --
  e.g. test with an SFTP client like [WinSCP](https://winscp.net/) or
  FileZilla first, using the same host/port/credentials.
- **`NotADirectoryError`**: `LOCAL_ROOT_FOLDER` doesn't point at a real
  folder on this machine -- check the path.
- Nothing shows up as different but you expected changes: check
  `TIMESTAMP_TOLERANCE_SECONDS` in `.env` and make sure both machines' clocks
  are reasonably accurate.

---

## The `scan` mode

Besides `sftp_compare`, the tool also has a standalone `scan` mode: it walks
one local folder and saves its contents (relative paths, sizes,
last-modified times) to a small JSON "manifest" file. This is useful on its
own as a point-in-time inventory/audit of a folder, independent of any
comparison.

```ini
MODE=scan
SCAN_SIDE_LABEL=A
SCAN_ROOT_FOLDER=C:\path\to\folder_to_scan
```

```bash
python folder_sync_tool.py
# -> produces manifest_A.json
```

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
- `paramiko` (SFTP support -- only exercised by `sftp_compare` mode)
- `python-dotenv` (loads `.env`)

## Security note

Never commit `.env`, manifest JSON files, or the generated `.xlsx` report if
they contain sensitive paths or data -- all three are excluded by
`.gitignore` by default.
