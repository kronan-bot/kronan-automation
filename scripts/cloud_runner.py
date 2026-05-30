"""
cloud_runner.py — GitHub Actions version of kronan_runner.py

Checks Dropbox for new .xlsx sales reports, processes them,
deploys to Netlify, then moves the file to Processed/.

Required environment variables (GitHub Secrets):
  DROPBOX_REFRESH_TOKEN  — from one-time OAuth setup
  DROPBOX_APP_KEY        — your Dropbox app key
  DROPBOX_APP_SECRET     — your Dropbox app secret

Optional:
  NETLIFY_TOKEN          — Netlify personal access token
  NETLIFY_SITE_ID        — Netlify site ID (auto-discovered if omitted)
"""

import os, sys, subprocess, shutil, tempfile
from pathlib import Path
from datetime import datetime

import dropbox
from dropbox.files import WriteMode, FolderMetadata, FileMetadata
from dropbox.exceptions import ApiError

# ── Paths ──────────────────────────────────────────────────────────────────
REPO_ROOT   = Path(__file__).parent.parent.resolve()
SCRIPT_DIR  = REPO_ROOT / 'scripts'
DATA_DIR    = REPO_ROOT / 'data'
DATA_DIR.mkdir(exist_ok=True)

# Tell the existing scripts where to find data files
os.environ['KRONAN_BASE'] = str(DATA_DIR)

# ── Dropbox config ────────────────────────────────────────────────────────────
DROPBOX_REPORTS_FOLDER   = '/Krónan Reports'
DROPBOX_PROCESSED_FOLDER = '/Krónan Reports/Processed'

def get_dbx():
    return dropbox.Dropbox(
        oauth2_refresh_token=os.environ['DROPBOX_REFRESH_TOKEN'],
        app_key=os.environ['DROPBOX_APP_KEY'],
        app_secret=os.environ['DROPBOX_APP_SECRET'],
    )

# ── List .xlsx files in the reports folder ────────────────────────────────────
def list_new_reports(dbx):
    """Yields (dropbox_path, local_name, is_package) for each unprocessed report."""
    try:
        result = dbx.files_list_folder(DROPBOX_REPORTS_FOLDER)
    except ApiError as e:
        print(f'✗ Could not list Dropbox folder: {e}')
        sys.exit(1)

    for entry in result.entries:
        name = entry.name

        # Normal .xlsx file
        if isinstance(entry, FileMetadata) and name.lower().endswith('.xlsx'):
            yield entry.path_display, name, False

        # Power Automate "folder-as-file" package: a directory named *.xlsx
        # with the actual bytes inside a file called "undefined"
        elif isinstance(entry, FolderMetadata) and name.lower().endswith('.xlsx'):
            try:
                inner = dbx.files_list_folder(entry.path_display)
                undef = next(
                    (e for e in inner.entries
                     if isinstance(e, FileMetadata) and e.name == 'undefined'),
                    None
                )
                if undef:
                    yield undef.path_display, name, True   # path is to "undefined" file
            except ApiError:
                pass


# ── Run a script ──────────────────────────────────────────────────────────────
def run(script, *args):
    cmd = [sys.executable, str(SCRIPT_DIR / script)] + list(args)
    print(f'  $ {" ".join(cmd)}')
    r = subprocess.run(cmd)
    return r.returncode


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    dbx = get_dbx()

    reports = list(list_new_reports(dbx))

    if not reports:
        print('No new reports found in Dropbox. Nothing to do.')
        _set_output('processed', 'false')
        return

    processed_any = False

    for dropbox_path, name, is_package in reports:
        print(f'\n→ Found: {name}')

        # 1. Download to a temp file
        tmp = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
        tmp.close()
        try:
            dbx.files_download_to_file(tmp.name, dropbox_path)
            print(f'  Downloaded ({os.path.getsize(tmp.name):,} bytes)')
        except ApiError as e:
            print(f'  ✗ Download failed: {e}')
            os.unlink(tmp.name)
            continue

        # 2. Process with kronan_master.py
        rc = run('kronan_master.py', tmp.name)
        os.unlink(tmp.name)
        if rc != 0:
            print(f'  ✗ kronan_master.py failed (exit {rc})')
            continue

        # 3. Update dashboard HTML
        rc = run('update_dashboard_data.py')
        if rc != 0:
            print(f'  ✗ update_dashboard_data.py failed (exit {rc})')
            continue

        # 4. Deploy to Netlify
        rc = run('netlify_deploy.py')
        if rc != 0:
            print(f'  ⚠ netlify_deploy.py failed (exit {rc}) — data saved, deploy skipped')

        # 5. Move to Processed/ in Dropbox
        ts   = datetime.now().strftime('%Y%m%d_%H%M%S')
        dest = f'{DROPBOX_PROCESSED_FOLDER}/{ts}_{name}'
        try:
            # For packages, move the parent folder; for normal files, move the file
            if is_package:
                parent = str(Path(dropbox_path).parent)   # e.g. /Krónan Reports/foo.xlsx
                dbx.files_move_v2(parent, dest)
            else:
                dbx.files_move_v2(dropbox_path, dest)
            print(f'  ✓ Moved → Processed/{ts}_{name}')
        except ApiError as e:
            print(f'  ⚠ Could not move to Processed/: {e}')

        print(f'  ✅ Done: {name}')
        processed_any = True

    _set_output('processed', 'true' if processed_any else 'false')


def _set_output(key, value):
    """Write a GitHub Actions step output."""
    output_file = os.environ.get('GITHUB_OUTPUT')
    if output_file:
        with open(output_file, 'a') as f:
            f.write(f'{key}={value}\n')
    print(f'[output] {key}={value}')


if __name__ == '__main__':
    main()
