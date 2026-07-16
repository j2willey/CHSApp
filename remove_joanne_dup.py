# Remove the duplicate Joanne doc (old ID) — keeps joanne-stephenson-medic.
# The first add script used "joanne-stephenson-nurse"; the ID was later changed
# to "joanne-stephenson-medic", leaving a stale duplicate. This deletes the old one.
# Deletes are no-ops if the doc doesn't exist, so this is safe to run.
#
# Run locally:
#   cd "C:\Users\Michael\Documents\Claude\Projects\CHS app"
#   python remove_joanne_dup.py
import firebase_admin
from firebase_admin import credentials, firestore

SA_KEY = r'C:\Users\Michael\Documents\Claude\Projects\CHS app\serviceAccountKey.json'

STALE_ID = "joanne-stephenson-nurse"   # the duplicate to remove
SESSIONS = [
    "Week_1", "Eagle_Achievement_Camp",
    "Week_3", "Week_4", "Week_5", "Week_6",
]

cred = credentials.Certificate(SA_KEY)
firebase_admin.initialize_app(cred)
db = firestore.client()

print(f"Deleting duplicate {STALE_ID} ...")
# Global doc
db.collection('staff').document(STALE_ID).delete()
print(f"  ✓ staff/{STALE_ID}")
# Per-session docs
for session in SESSIONS:
    db.collection('sessions').document(session) \
      .collection('staff').document(STALE_ID).delete()
    print(f"  ✓ {session}")

print("\nDone. Refresh the app. Joanne should now appear once (joanne-stephenson-medic).")
