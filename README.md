# Folder Sync Comparison Tool

Compares two folder trees -- which may live on two different machines (local,
RDP, or an SFTP server), across Windows and Mac -- and produces an Excel
report showing what needs to be copied where to bring them back in sync.

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

## Modes

Set `MODE` in your `.env` file to one of:

| Mode             | Use when...                                                                 |
|------------------|------------------------------------------------------------------------------|
| `scan`           | You can only see one folder from this machine. Produces a small JSON manifest to transfer to the other machine. |
| `compare`        | You have both manifest JSON files (produced by `scan`) on one machine. Diffs them and produces the Excel report. |
| `direct_compare` | Both folders are reachable from one machine at once (e.g. a mapped network drive). No manifest step needed. |
| `sftp_compare`   | Folder A lives on an SFTP server, folder B is local to this machine. Connects over SFTP, walks both folders, and compares in one run. |

## Setup

```bash
git clone <this-repo-url>
cd folder-sync-tool
python -m venv .venv
.venv\Scripts\activate        # Windows
# source .venv/bin/activate   # macOS/Linux

pip install -r requirements.txt

copy .env.example .env        # Windows
# cp .env.example .env        # macOS/Linux
```

Then edit `.env` with your real folder paths and (if using `sftp_compare`)
your SFTP host/credentials.

**`.env` is gitignored and must never be committed** -- it's the only place
credentials should live. `.env.example` documents every variable with safe
placeholder values and is the file that gets committed.

## Usage

```bash
python folder_sync_tool.py
```

It reads `MODE` and all other settings from `.env`, runs the comparison, and
writes `folder_sync_report.xlsx` (path/name configurable via
`EXCEL_REPORT_PATH`) to the working directory.

### Typical two-machine workflow (no shared access, no SFTP)

```bash
# On machine A
MODE=scan
SCAN_SIDE_LABEL=A
SCAN_ROOT_FOLDER=<A's folder path>
python folder_sync_tool.py
# -> produces manifest_A.json

# On machine B
MODE=scan
SCAN_SIDE_LABEL=B
SCAN_ROOT_FOLDER=<B's folder path>
python folder_sync_tool.py
# -> produces manifest_B.json

# Copy manifest_A.json and manifest_B.json onto the same machine, then:
MODE=compare
MANIFEST_A_PATH=manifest_A.json
MANIFEST_B_PATH=manifest_B.json
python folder_sync_tool.py
# -> produces folder_sync_report.xlsx
```

### SFTP workflow

Folder A is on the SFTP server, folder B is local to the machine running the
script:

```bash
MODE=sftp_compare
SFTP_HOST=sftp.example.com
SFTP_PORT=22
SFTP_USERNAME=your_username
SFTP_PASSWORD=your_password        # or leave blank and set SFTP_PRIVATE_KEY_PATH
SFTP_ROOT_FOLDER=/remote/path/to/folder_a
LOCAL_ROOT_FOLDER=C:\path\to\folder_b
python folder_sync_tool.py
```

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
