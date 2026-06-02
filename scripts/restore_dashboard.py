import os, dropbox

BASE = os.environ.get('KRONAN_BASE', 'data')
os.makedirs(BASE, exist_ok=True)
DEST = os.path.join(BASE, 'Krónan_Dashboard.html')

dbx = dropbox.Dropbox(
    oauth2_refresh_token=os.environ['DROPBOX_REFRESH_TOKEN'],
    app_key=os.environ['DROPBOX_APP_KEY'],
    app_secret=os.environ['DROPBOX_APP_SECRET'],
)
try:
    dbx.files_download_to_file(DEST, '/Krónan_Dashboard_clean.html')
    print(f'✅ Dashboard restored from Dropbox ({os.path.getsize(DEST):,} bytes)')
except Exception as e:
    print(f'⚠ Could not restore dashboard: {e}')
