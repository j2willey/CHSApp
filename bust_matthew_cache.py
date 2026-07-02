# Force the app to load a fresh copy of Matthew Lewis's photo
# by appending ?v=2 to his photo URL in Firestore.
# Run: python bust_matthew_cache.py

import firebase_admin
from firebase_admin import credentials, firestore

SA_KEY   = r'C:\Users\Michael\Documents\Claude\Projects\CHS app\serviceAccountKey.json'
BUCKET   = 'camphisierraapp.firebasestorage.app'
BASE_URL = f'https://storage.googleapis.com/{BUCKET}/staff/headshots'

STAFF_ID = "matthew-lewis-archery"
NEW_URL  = f"{BASE_URL}/{STAFF_ID}.jpg?v=2"

SESSIONS = [
    "Week_1", "Eagle_Achievement_Camp",
    "Week_3", "Week_4", "Week_5", "Week_6",
]

cred = credentials.Certificate(SA_KEY)
firebase_admin.initialize_app(cred)
db = firestore.client()

print(f"Setting photo URL to: {NEW_URL}\n")

# Global doc
db.collection('staff').document(STAFF_ID).set({'photo': NEW_URL}, merge=True)
print(f"  ✓ staff/{STAFF_ID}")

# Per-session docs
for session in SESSIONS:
    ref = db.collection('sessions').document(session).collection('staff').document(STAFF_ID)
    if ref.get().exists:
        ref.set({'photo': NEW_URL}, merge=True)
        print(f"  ✓ {session}")

print("\nDone. Hard-refresh the app (Ctrl+Shift+R) to clear any local browser cache too.")
