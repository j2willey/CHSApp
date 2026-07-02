# ── FIX ALL PHOTO ISSUES ─────────────────────────────────────────────────────
# Handles all 9 corrections in one run:
#   1. Daniel S.     (Rifle)      - was showing Danica's photo
#   2. Matthew L.    (Archery)    - was showing Miguel's photo
#   3. Evan Godard   (Shotgun)    - had no photo
#   4. Stan Searing  (Climbing)   - was showing Saanvi Sharma's photo
#   5. Sam Jacobson  (Foxfire)    - was showing Katherine Rico → CLEAR photo
#   6. Cody B.       (Foxfire)    - was showing Chris Brodack's photo
#   7. Sean Miller   (Excursions) - was showing Sam McPheeters → CLEAR photo
#   8. Ian Horsely   (Kitchen)    - was showing wrong Ian
#   9. Lincoln H.    (Kitchen)    - display name was "John H." → "Lincoln H."
#
# Uses a fresh Storage path (staff/headshots/2026/) to bypass CDN cache.
# Run: python fix_all_photos.py

import os, firebase_admin
from firebase_admin import credentials, firestore, storage

SA_KEY      = r'C:\Users\Michael\Documents\Claude\Projects\CHS app\serviceAccountKey.json'
BUCKET      = 'camphisierraapp.firebasestorage.app'
RESIZED_DIR = r'C:\Users\Michael\Documents\Claude\Projects\CHS app\headshots_resized'
NEW_PREFIX  = 'staff/headshots/2026'
BASE_URL    = f'https://storage.googleapis.com/{BUCKET}/{NEW_PREFIX}'
SESSIONS    = ["Week_1","Eagle_Achievement_Camp","Week_3","Week_4","Week_5","Week_6"]

cred = credentials.Certificate(SA_KEY)
firebase_admin.initialize_app(cred, {'storageBucket': BUCKET})
db     = firestore.client()
bucket = storage.bucket()

# ── Helper ────────────────────────────────────────────────────────────────────
def patch_firestore(staff_id, updates):
    """Patch global staff doc + all session docs. Handles missing docs safely."""
    set_data    = {k: v for k, v in updates.items() if v is not None}
    delete_keys = [k for k, v in updates.items() if v is None]

    ref = db.collection('staff').document(staff_id)
    if set_data:
        ref.set(set_data, merge=True)
    if delete_keys and ref.get().exists:
        ref.update({k: firestore.DELETE_FIELD for k in delete_keys})

    for session in SESSIONS:
        ref = db.collection('sessions').document(session).collection('staff').document(staff_id)
        doc = ref.get()
        if not doc.exists:
            continue
        if set_data:
            ref.set(set_data, merge=True)
        if delete_keys:
            ref.update({k: firestore.DELETE_FIELD for k in delete_keys})

def upload_and_patch(staff_id):
    """Upload corrected image to fresh Storage path, patch Firestore URL."""
    local = os.path.join(RESIZED_DIR, f'{staff_id}.jpg')
    if not os.path.exists(local):
        print(f'  SKIP — no local file: {staff_id}')
        return
    blob = bucket.blob(f'{NEW_PREFIX}/{staff_id}.jpg')
    blob.upload_from_filename(local, content_type='image/jpeg')
    blob.make_public()
    url = f'{BASE_URL}/{staff_id}.jpg'
    patch_firestore(staff_id, {'photo': url})
    kb = os.path.getsize(local) // 1024
    print(f'  ✓ {staff_id} ({kb}KB) → {url}')

# ── 1–4, 6, 8: Upload corrected images + patch Firestore ─────────────────────
print('=== Uploading corrected headshots ===')
for staff_id in [
    'daniel-scott-rifle',       # 1. Daniel S.
    'matthew-lewis-archery',    # 2. Matthew L.
    'evan-godard-shotgun',      # 3. Evan Godard
    'stan-searing-climbing',    # 4. Stan Searing
    'cody-bui-foxfire',         # 6. Cody B.
    'ian-horsely-kitchen',      # 8. Ian Horsely
    'danica-spears-scoutcraft', # also fix Danica while we're at it
]:
    upload_and_patch(staff_id)

# ── 5: Sam Jacobson — clear photo (no real photo exists) ─────────────────────
print('\n=== Clearing photos (no real photo available) ===')
patch_firestore('sam-jacobson-foxfire', {'photo': None})
print('  ✓ sam-jacobson-foxfire — photo cleared')

# ── 7: Sean Miller — clear photo ─────────────────────────────────────────────
patch_firestore('sean-miller-excursions', {'photo': None})
print('  ✓ sean-miller-excursions — photo cleared')

# ── 9: Lincoln H. — display name fix ─────────────────────────────────────────
print('\n=== Name fix ===')
patch_firestore('john-h-kitchen', {'name': 'Lincoln H.'})
print('  ✓ john-h-kitchen → "Lincoln H."')

print('\n✅ All done. Hard-refresh the app (Ctrl+Shift+R).')
