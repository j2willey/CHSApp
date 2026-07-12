# Fix photos stuck in Firebase Storage CDN cache by uploading to a fresh path.
# Uploads corrected images to staff/headshots/2026/<id>.jpg (new path = no cache)
# then updates Firestore to point to the new URLs.

import os, firebase_admin
from firebase_admin import credentials, firestore, storage

SA_KEY      = r'C:\Users\Michael\Documents\Claude\Projects\CHS app\serviceAccountKey.json'
BUCKET      = 'camphisierraapp.firebasestorage.app'
RESIZED_DIR = r'C:\Users\Michael\Documents\Claude\Projects\CHS app\headshots_resized'
NEW_PREFIX  = 'staff/headshots/2026'   # fresh path — CDN has never cached these URLs
BASE_URL    = f'https://storage.googleapis.com/{BUCKET}/{NEW_PREFIX}'

SESSIONS = ["Week_1","Eagle_Achievement_Camp","Week_3","Week_4","Week_5","Week_6"]

# All IDs with CDN-cached wrong images
FIX_IDS = [
    "daniel-scott-rifle",
    "danica-spears-scoutcraft",
    "matthew-lewis-archery",
    "stan-searing-climbing",
    "cody-bui-foxfire",
    "evan-godard-shotgun",
    "ian-horsely-kitchen",
]

cred = credentials.Certificate(SA_KEY)
firebase_admin.initialize_app(cred, {'storageBucket': BUCKET})
db     = firestore.client()
bucket = storage.bucket()

print(f"Uploading to {NEW_PREFIX}/...\n")
for staff_id in FIX_IDS:
    local = os.path.join(RESIZED_DIR, f'{staff_id}.jpg')
    if not os.path.exists(local):
        print(f'  SKIP (no local file): {staff_id}')
        continue

    # Upload to new path
    blob = bucket.blob(f'{NEW_PREFIX}/{staff_id}.jpg')
    blob.upload_from_filename(local, content_type='image/jpeg')
    blob.make_public()
    new_url = f'{BASE_URL}/{staff_id}.jpg'
    kb = os.path.getsize(local) // 1024
    print(f'  uploaded {staff_id} ({kb}KB)')

    # Patch Firestore with the new URL
    db.collection('staff').document(staff_id).set({'photo': new_url}, merge=True)
    for session in SESSIONS:
        ref = db.collection('sessions').document(session).collection('staff').document(staff_id)
        if ref.get().exists:
            ref.set({'photo': new_url}, merge=True)
    print(f'  ✓ Firestore patched → {new_url}')

print('\nDone. Hard-refresh the app (Ctrl+Shift+R).')
