# Re-upload Stan Searing's headshot and force a cache-bust in Firestore.
import os, firebase_admin
from firebase_admin import credentials, firestore, storage

SA_KEY      = r'C:\Users\Michael\Documents\Claude\Projects\CHS app\serviceAccountKey.json'
BUCKET      = 'camphisierraapp.firebasestorage.app'
RESIZED_DIR = r'C:\Users\Michael\Documents\Claude\Projects\CHS app\headshots_resized'
STAFF_ID    = 'stan-searing-climbing'
SESSIONS    = ["Week_1","Eagle_Achievement_Camp","Week_3","Week_4","Week_5","Week_6"]

cred = credentials.Certificate(SA_KEY)
firebase_admin.initialize_app(cred, {'storageBucket': BUCKET})
db     = firestore.client()
bucket = storage.bucket()

# Upload correct photo to Firebase Storage
local = os.path.join(RESIZED_DIR, f'{STAFF_ID}.jpg')
blob  = bucket.blob(f'staff/headshots/{STAFF_ID}.jpg')
blob.upload_from_filename(local, content_type='image/jpeg')
blob.make_public()
print(f'Uploaded {STAFF_ID}.jpg')

# Patch Firestore with ?v=2 to bypass CDN cache
url = f'https://storage.googleapis.com/{BUCKET}/staff/headshots/{STAFF_ID}.jpg?v=2'
db.collection('staff').document(STAFF_ID).set({'photo': url}, merge=True)
print(f'  ✓ staff/{STAFF_ID}')
for session in SESSIONS:
    ref = db.collection('sessions').document(session).collection('staff').document(STAFF_ID)
    if ref.get().exists:
        ref.set({'photo': url}, merge=True)
        print(f'  ✓ {session}')

print('\nDone. Hard-refresh the app (Ctrl+Shift+R).')
