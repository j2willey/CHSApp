"""
check_missing_headshots.py
Pulls the staff list from Firestore (Week_1 session) and compares
against the headshots_resized folder to find who's missing a photo.

Usage:
    python check_missing_headshots.py
"""

import os
import firebase_admin
from firebase_admin import credentials, firestore

SA_KEY       = r'C:\Users\Michael\Documents\Claude\Projects\CHS app\serviceAccountKey.json'
HEADSHOTS_DIR = r'C:\Users\Michael\Documents\Claude\Projects\CHS app\headshots_resized'
SESSION      = "Week_1"

cred = credentials.Certificate(SA_KEY)
firebase_admin.initialize_app(cred)
db = firestore.client()

# All staff in Firestore
print(f"Fetching staff from Firestore ({SESSION})...\n")
docs = db.collection('sessions').document(SESSION).collection('staff').stream()
staff = {doc.id: doc.to_dict() for doc in docs}

# All headshots on disk
headshots = {
    os.path.splitext(f)[0]
    for f in os.listdir(HEADSHOTS_DIR)
    if f.endswith('.jpg') and not f.startswith('_')
}

# Compare
missing = sorted(sid for sid in staff if sid not in headshots)
extra   = sorted(sid for sid in headshots if sid not in staff)

print(f"Total staff in Firestore: {len(staff)}")
print(f"Total headshots on disk:  {len(headshots)}")
print()

if missing:
    print(f"MISSING headshots ({len(missing)}) — in Firestore but no photo file:")
    for sid in missing:
        name = staff[sid].get('name', sid)
        has_photo = bool(staff[sid].get('photo'))
        print(f"  {sid}  (photo URL set: {has_photo})")
else:
    print("✅ No missing headshots — everyone has a file.")

print()
if extra:
    print(f"EXTRA files ({len(extra)}) — headshot on disk but no Firestore record:")
    for sid in extra:
        print(f"  {sid}")
