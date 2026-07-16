# Add Joanne Stephenson (Camp Nurse) to Week 4 & Week 6,
# and (re)add Brent Nicolai to Week 5.
# Staff docs feed BOTH the org chart and the appreciation picker automatically.
# Run locally (sandbox cannot reach Firestore):
#   cd "C:\Users\Michael\Documents\Claude\Projects\CHS app"
#   python add_nurse_and_brent.py
import firebase_admin
from firebase_admin import credentials, firestore

SA_KEY = r'C:\Users\Michael\Documents\Claude\Projects\CHS app\serviceAccountKey.json'

cred = credentials.Certificate(SA_KEY)
firebase_admin.initialize_app(cred)
db = firestore.client()

NEW_STAFF = [
    {
        "id": "joanne-stephenson-medic",
        "data": {"name": "Joanne Stephenson", "role": "Camp Nurse", "area": "Medic"},
        "sessions": ["Week_4", "Week_6"],
    },
    {
        "id": "brent-nicolai-medic",
        "data": {"name": "Brent Nicolai", "area": "Medic"},
        "sessions": ["Week_5"],
    },
]

for staff in NEW_STAFF:
    print(f"\n{staff['id']}  ({', '.join(staff['sessions'])})")
    for session in staff["sessions"]:
        db.collection('sessions').document(session).collection('staff') \
          .document(staff['id']).set(staff['data'], merge=True)
        print(f"  ✓ {session}")

print("\nDone. Refresh the app.")
