import os, dropbox
from dropbox.exceptions import ApiError

dbx = dropbox.Dropbox(
    oauth2_refresh_token=os.environ['DROPBOX_REFRESH_TOKEN'],
    app_key=os.environ['DROPBOX_APP_KEY'],
    app_secret=os.environ['DROPBOX_APP_SECRET'],
)

# List Processed folder to find June 1 file
try:
    result = dbx.files_list_folder('/Kronan Reports/Processed')
    june1 = next((e.path_display for e in result.entries if '2026-06-01' in e.name), None)
except ApiError:
    try:
        result = dbx.files_list_folder('/Krónan Reports/Processed')
        june1 = next((e.path_display for e in result.entries if '2026-06-01' in e.name), None)
    except ApiError as e:
        print(f'Could not list Processed: {e}')
        june1 = None

if june1:
    dest = '/Krónan Reports/2026-06-01_Krónan söluskýrsla.xlsx'
    try:
        dbx.files_move_v2(june1, dest)
        print(f'Moved {june1} back to main folder')
    except ApiError as e:
        print(f'Move failed: {e}')
else:
    print('June 1 file not found in Processed')
