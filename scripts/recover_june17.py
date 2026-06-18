import os, sys, dropbox
from dropbox.exceptions import ApiError

dbx = dropbox.Dropbox(
    oauth2_refresh_token=os.environ["DROPBOX_REFRESH_TOKEN"],
    app_key=os.environ["DROPBOX_APP_KEY"],
    app_secret=os.environ["DROPBOX_APP_SECRET"],
)

june17_src = None
for folder in ["/Kr\u00f3nan Reports/Processed", "/Kronan Reports/Processed"]:
    try:
        result = dbx.files_list_folder(folder)
        for e in result.entries:
            if "2026-06-17" in e.name:
                june17_src = e.path_display
                print("Found:", june17_src)
                break
        if june17_src:
            break
    except ApiError as ex:
        print("Cannot list", folder, str(ex))

if not june17_src:
    print("June 17 file not found in Processed/ - already moved or missing")
    sys.exit(0)

dest = "/Kr\u00f3nan Reports/2026-06-17_Kr\u00f3nan s\u00f6lusk\u00fdrsla.xlsx"
try:
    dbx.files_move_v2(june17_src, dest)
    print("Moved back to:", dest)
except ApiError as ex:
    print("Move failed:", str(ex))
    sys.exit(1)
