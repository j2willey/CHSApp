"""
rollback_hosting.py
Lists recent Firebase Hosting releases for camphisierraapp and lets you
roll back to any previous version without needing index.html.

Usage:
    python rollback_hosting.py
"""

import json, sys, time, urllib.request, urllib.error

SERVICE_ACCOUNT = r'C:\Users\Michael\Documents\Claude\Projects\CHS app\serviceAccountKey.json'
SITE_ID = 'camphisierraapp'

# ── Auth (same pattern as deploy_index.py) ────────────────────────────────────
import google.auth
import google.auth.transport.requests
from google.oauth2 import service_account

SCOPES = ['https://www.googleapis.com/auth/firebase', 'https://www.googleapis.com/auth/cloud-platform']
creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT, scopes=SCOPES)
creds.refresh(google.auth.transport.requests.Request())
token = creds.token

BASE = 'https://firebasehosting.googleapis.com/v1beta1'

def api(method, url, body=None):
    data = json.dumps(body).encode() if body else None
    headers = {'Authorization': f'Bearer {token}', 'Content-Type': 'application/json'}
    req = urllib.request.Request(url, data=data, headers=headers, method=method)
    with urllib.request.urlopen(req) as r:
        return json.loads(r.read())

# ── List releases ─────────────────────────────────────────────────────────────
print("Fetching release history...\n")
resp = api('GET', f'{BASE}/sites/{SITE_ID}/releases?pageSize=20')
releases = resp.get('releases', [])

if not releases:
    print("No releases found.")
    sys.exit(1)

print(f"{'#':<4} {'Date':<25} {'Version':<30} {'Status'}")
print("-" * 80)
for i, rel in enumerate(releases):
    create_time = rel.get('releaseTime', '')[:19].replace('T', ' ')
    version_name = rel.get('version', {}).get('name', '')
    version_id = version_name.split('/')[-1] if version_name else 'N/A'
    status = rel.get('type', 'DEPLOY')
    print(f"{i:<4} {create_time:<25} {version_id:<30} {status}")

print()
choice = input("Enter # to roll back to that release (or q to quit): ").strip()
if choice.lower() == 'q':
    sys.exit(0)

try:
    idx = int(choice)
    target = releases[idx]
except (ValueError, IndexError):
    print("Invalid selection.")
    sys.exit(1)

version_name = target.get('version', {}).get('name', '')
if not version_name:
    print("Could not find version name for that release.")
    sys.exit(1)

version_id = version_name.split('/')[-1]
create_time = target.get('releaseTime', '')[:19].replace('T', ' ')
print(f"\nRolling back to: {create_time}  ({version_id})")
confirm = input("Confirm? (y/n): ").strip().lower()
if confirm != 'y':
    print("Aborted.")
    sys.exit(0)

# ── Create new release pointing to old version ────────────────────────────────
release = api('POST', f'{BASE}/sites/{SITE_ID}/releases?versionName={version_name}')
print(f"\n✓ Rolled back. Live at https://{SITE_ID}.web.app")
print(f"  Release: {release.get('name','')}")
