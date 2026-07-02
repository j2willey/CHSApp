# Remove a staff member from all sessions in Firestore.
import firebase_admin
from firebase_admin import credentials, firestore

STAFF_ID = "evan-f-kitchen"

SESSIONS = [
    "Week_1", "Eagle_Achievement_Camp",
    "Week_3", "Week_4", "Week_5", "Week_6",
]

SA_KEY = r'C:\Users\Michael\Documents\Claude\Projects\CHS app\serviceAccountKey.json'

cred = credentials.Certificate(SA_KEY)
firebase_admin.initialize_app(cred)
db = firestore.client()

print(f"Deleting {STAFF_ID} from all sessions...")
for session in SESSIONS:
    ref = db.collection('sessions').document(session).collection('staff').document(STAFF_ID)
    ref.delete()
    print(f"  deleted from {session}")

print("\nDone.")
