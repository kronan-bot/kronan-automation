#!/usr/bin/env python3
"""Deploy Krónan Dashboard to Netlify (uses only stdlib — no pip needed)."""

import os
import sys
import json
import zipfile
import tempfile
import urllib.request
import urllib.error

# ── Config — reads from env vars (GitHub Actions) or falls back to local files ─
TOKEN   = os.environ.get('NETLIFY_TOKEN') or "nfp_5oshjYHvSV8iJ1AV625vTaTAUofby8hfa416"
SITE_ID = os.environ.get('NETLIFY_SITE_ID')  # set in GitHub secrets

_BASE = os.environ.get('KRONAN_BASE')
if _BASE:
    SITE_ID_FILE = os.path.join(_BASE, '.netlify_site_id')
    DASHBOARD    = os.path.join(_BASE, 'Krónan_Dashboard.html')
else:
    SITE_ID_FILE = os.path.expanduser("~/Documents/Krónan/scripts/.netlify_site_id")
    DASHBOARD    = os.path.expanduser("~/Documents/Krónan/Krónan_Dashboard.html")
# ───────────────────────────────────────────────────────────────────────────────

HEADERS_FILE = """\
/*
  Content-Type: text/html; charset=utf-8
  X-Content-Type-Options: nosniff
"""


def netlify_request(method, path, data=None, content_type="application/json"):
    url = f"https://api.netlify.com/api/v1{path}"
    headers = {
        "Authorization": f"Bearer {TOKEN}",
        "Content-Type": content_type,
    }
    if isinstance(data, dict):
        data = json.dumps(data).encode()
    req = urllib.request.Request(url, data=data, headers=headers, method=method)
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            return json.loads(resp.read())
    except urllib.error.HTTPError as e:
        print(f"[netlify] HTTP {e.code}: {e.read().decode()}")
        sys.exit(1)


def get_or_create_site():
    # 1. Env var (GitHub Actions)
    if SITE_ID:
        return SITE_ID

    # 2. Local file
    id_path = os.path.expanduser(SITE_ID_FILE)
    if os.path.exists(id_path):
        with open(id_path) as f:
            return f.read().strip()

    # 3. Look up by site name
    try:
        sites = netlify_request("GET", "/sites?filter=all")
        for s in sites:
            if s.get('name') == 'kronan-dashboard':
                site_id = s['id']
                os.makedirs(os.path.dirname(id_path), exist_ok=True)
                with open(id_path, 'w') as f:
                    f.write(site_id)
                print(f"[netlify] Found existing site: kronan-dashboard (id={site_id})")
                return site_id
    except Exception:
        pass

    # 4. Create new site
    data = netlify_request("POST", "/sites", {"name": "kronan-dashboard"})
    site_id = data["id"]
    url     = data.get("ssl_url") or data.get("url", "")
    os.makedirs(os.path.dirname(id_path), exist_ok=True)
    with open(id_path, "w") as f:
        f.write(site_id)
    print(f"[netlify] Created new site: {url}  (id={site_id})")
    return site_id


def deploy(site_id):
    dashboard_path = os.path.expanduser(DASHBOARD)
    if not os.path.exists(dashboard_path):
        print(f"[netlify] ERROR: Dashboard not found at {dashboard_path}")
        sys.exit(1)

    tmp = tempfile.NamedTemporaryFile(suffix=".zip", delete=False)
    tmp.close()
    with zipfile.ZipFile(tmp.name, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.write(dashboard_path, "index.html")
        zf.writestr("_headers", HEADERS_FILE)

    try:
        with open(tmp.name, "rb") as f:
            zip_data = f.read()
        data = netlify_request(
            "POST",
            f"/sites/{site_id}/deploys",
            data=zip_data,
            content_type="application/zip",
        )
    finally:
        os.unlink(tmp.name)

    url = data.get("ssl_url") or data.get("url") or f"https://{site_id}.netlify.app"
    print(f"[netlify] Deployed → {url}")
    return url


if __name__ == "__main__":
    site_id = get_or_create_site()
    deploy(site_id)
