# Remove Matthew Lewis (Archery) and Isabella DeMarco (Aquatics)
# from Week 4, Week 5, and Week 6 only. Leaves earlier weeks intact.
# Deletes are no-ops if the doc doesn't exist, so this is safe.
#
# Run locally:
#   cd "C:\Users\Michael\Documents\Claude\Projects\CHS app"
#   python remove_matthew_isabella.py
import firebase_admin
from firebase_admin import credentials, firestore

SA_KEY = r'C:\Users\Michael\Documents\Claude\Projects\CHS app\serviceAccountKey.json'

STAFF_IDS = ["matthew-lewis-archery", "isabella-demarco-aquatics"]
SESSIONS  = ["Week_4", "Week_5", "Week_6"]

cred = credentials.Certificate(SA_KEY)
firebase_admin.initialize_app(cred)
db = firestore.client()

for staff_id in STAFF_IDS:
    print(f"\n{staff_id}")
    for session in SESSIONS:
        db.collection('sessions').document(session) \
          .collection('staff').document(staff_id).delete()
        print(f"  ✓ removed from {session}")

print("\nDone. Refresh the app.")
