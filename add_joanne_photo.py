# Resize + upload Joanne Stephenson's headshot and patch Firestore.
#
# BEFORE RUNNING:
#   Save the photo Michael pasted in chat to:
#     C:\Users\Michael\Documents\Claude\Projects\CHS app\joanne_source.jpg
#
# Then run locally (sandbox cannot reach Firebase/Firestore):
#   cd "C:\Users\Michael\Documents\Claude\Projects\CHS app"
#   python add_joanne_photo.py
#
# What it does:
#   1. Center-crops + resizes joanne_source.jpg to 320x320 JPEG
#      -> headshots_resized/joanne-stephenson-medic.jpg
#   2. Uploads it to Firebase Storage: staff/headshots/joanne-stephenson-medic.jpg
#   3. Patches the `photo` field on the global staff doc + Week_4 + Week_6

import os
from PIL import Image
import firebase_admin
from firebase_admin import credentials, firestore, storage

BASE        = r'C:\Users\Michael\Documents\Claude\Projects\CHS app'
SA_KEY      = os.path.join(BASE, 'serviceAccountKey.json')
BUCKET      = 'camphisierraapp.firebasestorage.app'
BASE_URL    = f'https://storage.googleapis.com/{BUCKET}/staff/headshots'
RESIZED_DIR = os.path.join(BASE, 'headshots_resized')

STAFF_ID    = 'joanne-stephenson-medic'
SOURCE      = os.path.join(BASE, 'joanne_source.jpg')
SESSIONS    = ["Week_4", "Week_6"]   # Joanne is only on these two weeks

# ── Step 1: center-crop to square + resize to 320x320 ─────────────────────────
if not os.path.exists(SOURCE):
    raise SystemExit(f"Source photo not found: {SOURCE}\n"
                     f"Save the pasted photo there first, then re-run.")

img = Image.open(SOURCE).convert('RGB')
w, h = img.size
side = min(w, h)
left = (w - side) // 2
top  = (h - side) // 2
img = img.crop((left, top, left + side, top + side)).resize((320, 320), Image.LANCZOS)

os.makedirs(RESIZED_DIR, exist_ok=True)
out_path = os.path.join(RESIZED_DIR, f'{STAFF_ID}.jpg')
img.save(out_path, 'JPEG', quality=85)
print(f"✓ Resized -> {out_path} ({os.path.getsize(out_path)//1024}KB)")

# ── Step 2 + 3: upload to Storage + patch Firestore ───────────────────────────
cred = credentials.Certificate(SA_KEY)
firebase_admin.initialize_app(cred, {'storageBucket': BUCKET})
db     = firestore.client()
bucket = storage.bucket()

blob = bucket.blob(f'staff/headshots/{STAFF_ID}.jpg')
blob.upload_from_filename(out_path, content_type='image/jpeg')
blob.make_public()
print(f"✓ Uploaded to Storage: staff/headshots/{STAFF_ID}.jpg")

photo_url = f'{BASE_URL}/{STAFF_ID}.jpg'

# Global staff/{id} doc
db.collection('staff').document(STAFF_ID).set({'photo': photo_url}, merge=True)
print(f"✓ staff/{STAFF_ID}")

# Per-session docs
for session in SESSIONS:
    db.collection('sessions').document(session) \
      .collection('staff').document(STAFF_ID).set({'photo': photo_url}, merge=True)
    print(f"✓ {session}")

print("\nDone. Refresh the app.")
