import os, sys, json
from pathlib import Path
import dropbox
from dropbox.exceptions import ApiError

DATA_DIR = Path(__file__).parent.parent / 'data'
DATA_DIR.mkdir(exist_ok=True)
OUT = DATA_DIR / 'dropbox_listing.txt'

def get_dbx():
    return dropbox.Dropbox(
        oauth2_refresh_token=os.environ['DROPBOX_REFRESH_TOKEN'],
        app_key=os.environ['DROPBOX_APP_KEY'],
        app_secret=os.environ['DROPBOX_APP_SECRET'],
    )

def main():
    lines = []
    try:
        dbx = get_dbx()
        lines.append('Connected to Dropbox OK')
        # Check account info
        try:
            acct = dbx.users_get_current_account()
            lines.append(f'Account: {acct.name.display_name} <{acct.email}>')
        except Exception as e:
            lines.append(f'Account info error: {e}')

        # List /Kronan Reports
        folder = '/Kronan Reports'
        for folder_path in ['/Krónan Reports', '/Kronan Reports', '/kronan reports']:
            try:
                result = dbx.files_list_folder(folder_path)
                lines.append(f'Folder {folder_path!r}: {len(result.entries)} entries, has_more={result.has_more}')
                for entry in result.entries:
                    lines.append(f'  [{type(entry).__name__}] {entry.path_display!r}')
            except ApiError as e:
                lines.append(f'Folder {folder_path!r}: ERROR {e}')

        # Also list root
        try:
            root = dbx.files_list_folder('')
            lines.append(f'Root listing: {len(root.entries)} entries')
            for entry in root.entries:
                lines.append(f'  [{type(entry).__name__}] {entry.path_display!r}')
        except ApiError as e:
            lines.append(f'Root listing error: {e}')

    except Exception as e:
        lines.append(f'FATAL: {e}')

    out_text = '\n'.join(lines)
    print(out_text)
    OUT.write_text(out_text)
    print(f'Written to {OUT}')

if __name__ == '__main__':
    main()
