# Fix incorrect photo URL assignments and display names in Firestore.
# Also re-uploads corrected image files to Firebase Storage.
# Run from your machine: python fix_photo_assignments.py
#
# What this fixes:
#   Rifle      - Daniel S.     was showing Danica's photo
#   Archery    - Matthew L.    was showing Miguel L.'s photo
#   Shotgun    - Evan Godard   had no photo despite one being uploaded
#   Climbing   - Stan S.       was showing Saanvi Sharma's photo
#   Foxfire    - Sam Jacobson  was showing Katherine Rico's photo → cleared (no real photo)
#   Foxfire    - Cody B.       was showing Chris Brodack's photo
#   Excursions - Sean Miller   was showing Sam McPheeters' photo
#   Kitchen    - Ian H.        was showing wrong Ian → corrected to Ian Horsely
#   Kitchen    - Lincoln H.    display name was "John H." → changed to "Lincoln H."

import os
import firebase_admin
from firebase_admin import credentials, firestore, storage

SA_KEY      = r'C:\Users\Michael\Documents\Claude\Projects\CHS app\serviceAccountKey.json'
BUCKET      = 'camphisierraapp.firebasestorage.app'
BASE_URL    = f'https://storage.googleapis.com/{BUCKET}/staff/headshots'
RESIZED_DIR = r'C:\Users\Michael\Documents\Claude\Projects\CHS app\headshots_resized'

SESSIONS = [
    "Week_1", "Eagle_Achievement_Camp",
    "Week_3", "Week_4", "Week_5", "Week_6",
]

# ── Files to re-upload to Firebase Storage ────────────────────────────────────
# These local files were corrected — overwrite the wrong images in Storage.
UPLOAD_IDS = [
    "daniel-scott-rifle",
    "danica-spears-scoutcraft",   # was showing Daniel's photo — fixed
    "matthew-lewis-archery",
    "stan-searing-climbing",
    "cody-bui-foxfire",
    "evan-godard-shotgun",
    "ian-horsely-kitchen",
]

# ── Firestore corrections ─────────────────────────────────────────────────────
# Each entry: (staff_doc_id, field_updates)
# Use photo=None to clear the photo field entirely.
CORRECTIONS = [
    ("daniel-scott-rifle",    {"photo": f"{BASE_URL}/daniel-scott-rifle.jpg"}),
    ("matthew-lewis-archery", {"photo": f"{BASE_URL}/matthew-lewis-archery.jpg"}),
    ("stan-searing-climbing", {"photo": f"{BASE_URL}/stan-searing-climbing.jpg"}),
    ("sam-jacobson-foxfire",  {"photo": None}),   # no real photo — clear it
    ("cody-bui-foxfire",      {"photo": f"{BASE_URL}/cody-bui-foxfire.jpg"}),
    ("ian-horsely-kitchen",   {"photo": f"{BASE_URL}/ian-horsely-kitchen.jpg"}),
    ("john-h-kitchen",        {"name": "Lincoln H."}),
    ("evan-godard-shotgun",   {"photo": f"{BASE_URL}/evan-godard-shotgun.jpg"}),
    ("sean-miller-excursions", {"photo": None}),   # no photo available — clear it
]

# ── Init Firebase ─────────────────────────────────────────────────────────────
cred = credentials.Certificate(SA_KEY)
firebase_admin.initialize_app(cred, {'storageBucket': BUCKET})
db     = firestore.client()
bucket = storage.bucket()

# ── Step 1: Upload corrected images to Firebase Storage ───────────────────────
print("=== Uploading corrected headshots to Firebase Storage ===\n")
for staff_id in UPLOAD_IDS:
    local_path = os.path.join(RESIZED_DIR, f"{staff_id}.jpg")
    if not os.path.exists(local_path):
        print(f"  SKIP (no local file): {staff_id}")
        continue
    blob = bucket.blob(f"staff/headshots/{staff_id}.jpg")
    blob.upload_from_filename(local_path, content_type="image/jpeg")
    blob.make_public()
    kb = os.path.getsize(local_path) // 1024
    print(f"  ✓ {staff_id} ({kb}KB)")

# ── Step 2: Patch Firestore photo URLs and names ──────────────────────────────
print("\n=== Patching Firestore documents ===")
for staff_id, updates in CORRECTIONS:
    print(f"\n{staff_id}")

    set_data    = {k: v for k, v in updates.items() if v is not None}
    delete_keys = [k for k, v in updates.items() if v is None]

    # Global staff/{id} doc
    ref = db.collection('staff').document(staff_id)
    if set_data:
        ref.set(set_data, merge=True)
    if delete_keys:
        if ref.get().exists:
            ref.update({key: firestore.DELETE_FIELD for key in delete_keys})
        else:
            print(f"  (no global doc — skipping field delete)")
    print(f"  ✓ staff/{staff_id}")

    # Per-session docs
    for session in SESSIONS:
        ref = db.collection('sessions').document(session) \
                .collection('staff').document(staff_id)
        doc = ref.get()
        if not doc.exists:
            continue
        if set_data:
            ref.set(set_data, merge=True)
        if delete_keys:
            ref.update({key: firestore.DELETE_FIELD for key in delete_keys})
        print(f"  ✓ {session}")

print("\nDone. Refresh the app to see all changes.")
