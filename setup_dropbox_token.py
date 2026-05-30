"""
Run this ONCE on your Mac to get a Dropbox refresh token.
The refresh token never expires — save it as a GitHub secret.

Usage:
  python setup_dropbox_token.py

You'll need:
  - DROPBOX_APP_KEY and DROPBOX_APP_SECRET from your Dropbox app
    (create one at https://www.dropbox.com/developers/apps)
"""

import os, sys

try:
    import dropbox
    from dropbox import DropboxOAuth2FlowNoRedirect
except ImportError:
    print("Installing dropbox SDK...")
    os.system(f"{sys.executable} -m pip install dropbox")
    import dropbox
    from dropbox import DropboxOAuth2FlowNoRedirect

APP_KEY    = input("Enter your Dropbox App Key: ").strip()
APP_SECRET = input("Enter your Dropbox App Secret: ").strip()

auth_flow = DropboxOAuth2FlowNoRedirect(
    APP_KEY,
    APP_SECRET,
    token_access_type='offline'   # gives a refresh token that never expires
)

authorize_url = auth_flow.start()
print(f"\n1. Open this URL in your browser:\n   {authorize_url}\n")
print("2. Click 'Allow' (you may need to log in)")
print("3. Copy the authorization code shown\n")

auth_code = input("Enter the authorization code: ").strip()

try:
    oauth_result = auth_flow.finish(auth_code)
except Exception as e:
    print(f"Error: {e}")
    sys.exit(1)

print("\n" + "="*60)
print("✅ SUCCESS — add these three values as GitHub Secrets:")
print("="*60)
print(f"\nDROPBOX_APP_KEY     = {APP_KEY}")
print(f"DROPBOX_APP_SECRET  = {APP_SECRET}")
print(f"DROPBOX_REFRESH_TOKEN = {oauth_result.refresh_token}")
print("\nThe refresh token never expires.")
print("="*60)
