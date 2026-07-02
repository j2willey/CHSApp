"""
CHS Staff Headshots — Upload to Firebase Storage & Update Firestore
Run from your machine: python upload_headshots.py

Resized images (320x320, ~19KB each) are in the outputs/headshots_resized/ folder.
Adjust IMAGES_DIR below if needed.
"""

import os
import sys
import firebase_admin
from firebase_admin import credentials, storage, firestore

# ── Config ────────────────────────────────────────────────────────────────────
SERVICE_ACCOUNT = r'C:\Users\Michael\Documents\Claude\Projects\CHS app\serviceAccountKey.json'
STORAGE_BUCKET   = 'camphisierraapp.firebasestorage.app'

# Resized images — in the same folder as this script
IMAGES_DIR = r'C:\Users\Michael\Documents\Claude\Projects\CHS app\headshots_resized'

# Staff IDs to upload (filename matches <staff-id>.jpg)
STAFF_IDS = [
    "angie-sullivan-trail-to-eagle",
    "bruce-lee-management",
    "chris-murphy-nature",
    "dani-didoszak-management",
    "joe-beard-shotgun",
    "joey-newman-solo-campers",
    "john-velasquez-adult-education",
    "mj-winter-management",
    "peter-mcginnis-management",
    "rebecca-rosen-observatory",
    "lorien-bower-archery",
]

# ── Init Firebase ─────────────────────────────────────────────────────────────
cred = credentials.Certificate(SERVICE_ACCOUNT)
firebase_admin.initialize_app(cred, {'storageBucket': STORAGE_BUCKET})
bucket = storage.bucket()
db = firestore.client()

# ── Upload ────────────────────────────────────────────────────────────────────
print(f"Uploading {len(STAFF_IDS)} headshots...\n")
urls = {}

for staff_id in STAFF_IDS:
    img_path = os.path.join(IMAGES_DIR, f"{staff_id}.jpg")
    if not os.path.exists(img_path):
        print(f"  SKIP {staff_id} — file not found at {img_path}")
        continue
    try:
        blob = bucket.blob(f"staff/headshots/{staff_id}.jpg")
        blob.upload_from_filename(img_path, content_type='image/jpeg')
        blob.make_public()
        url = blob.public_url
        urls[staff_id] = url

        # Update Firestore staff document
        db.collection('staff').document(staff_id).set({'photo': url}, merge=True)
        kb = os.path.getsize(img_path) // 1024
        print(f"  ✓ {staff_id} ({kb}KB) → {url}")
    except Exception as e:
        print(f"  ✗ {staff_id}: {e}")

print(f"\nDone. {len(urls)}/{len(STAFF_IDS)} uploaded.")

# ── Unmatched headshots ───────────────────────────────────────────────────────
print("""
Note — 1 headshot still unmatched:
  • 2026 Emile.jpg — no 'Emile' found in STAFF array
""")
