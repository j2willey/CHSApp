"""
find_real_app.py
Shows what files are in each Firebase Hosting release so we can find
the last version that had the real camp app, then rolls back to it.

Usage:
    python find_real_app.py
"""

import json, sys, urllib.request

SERVICE_ACCOUNT = r'C:\Users\Michael\Documents\Claude\Projects\CHS app\serviceAccountKey.json'
SITE_ID = 'camphisierraapp'

import google.auth, google.auth.transport.requests
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

# Fetch releases
print("Fetching releases...\n")
releases = []
page_token = None
while len(releases) < 100:
    url = f'{BASE}/sites/{SITE_ID}/releases?pageSize=25'
    if page_token:
        url += f'&pageToken={page_token}'
    resp = api('GET', url)
    releases.extend(resp.get('releases', []))
    page_token = resp.get('nextPageToken')
    if not page_token:
        break

# Show files in first 10 releases
print(f"Showing files in the 10 most recent releases:\n")
for i, rel in enumerate(releases[:10]):
    create_time = rel.get('releaseTime', '')[:16].replace('T', ' ')
    version_name = rel.get('version', {}).get('name', '')
    version_id = version_name.split('/')[-1] if version_name else 'N/A'
    rel_type = rel.get('type', 'DEPLOY')

    print(f"[{i}] {create_time} UTC  —  {version_id}  ({rel_type})")

    if rel_type != 'DEPLOY' or not version_name:
        print("     (no files — rollback/delete entry)")
        continue

    try:
        files_resp = api('GET', f'{BASE}/{version_name}/files?pageSize=100')
        files = files_resp.get('files', [])
        for f in files:
            path = f.get('path', '?')
            status = f.get('status', '')
            print(f"     {path}  [{status}]")
        if not files:
            print("     (no files listed)")
    except Exception as e:
        print(f"     (error: {e})")
    print()

print("\n" + "="*60)
choice = input("Enter # to roll back to that version (or q to quit): ").strip()
if choice.lower() == 'q':
    sys.exit(0)

try:
    idx = int(choice)
    target = releases[idx]
except (ValueError, IndexError):
    print("Invalid.")
    sys.exit(1)

version_name = target.get('version', {}).get('name', '')
version_id = version_name.split('/')[-1]
create_time = target.get('releaseTime', '')[:16].replace('T', ' ')
print(f"\nRolling back to: {create_time}  ({version_id})")
confirm = input("Confirm? (y/n): ").strip().lower()
if confirm != 'y':
    print("Aborted.")
    sys.exit(0)

release = api('POST', f'{BASE}/sites/{SITE_ID}/releases?versionName={version_name}')
print(f"\n✓ Done. Live at https://{SITE_ID}.web.app")
