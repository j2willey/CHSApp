# Patch name/area fields for newly added staff who were created with only a photo URL.
import firebase_admin
from firebase_admin import credentials, firestore

SA_KEY = r'C:\Users\Michael\Documents\Claude\Projects\CHS app\serviceAccountKey.json'
SESSIONS_ALL = ["Week_1", "Eagle_Achievement_Camp", "Week_3", "Week_4", "Week_5", "Week_6"]
SESSIONS_BRENT = ["Week_3", "Week_5"]

cred = credentials.Certificate(SA_KEY)
firebase_admin.initialize_app(cred)
db = firestore.client()

NEW_STAFF = [
    {
        "id": "griffen-goza-kitchen",
        "data": {"name": "Griffen Goza", "area": "Kitchen"},
        "sessions": SESSIONS_ALL,
    },
    {
        "id": "brent-nicolai-medic",
        "data": {"name": "Brent Nicolai", "area": "Medic"},
        "sessions": SESSIONS_BRENT,
    },
    {
        "id": "lukas-winter-camp-wide",
        "data": {"name": "Lukas Winter", "area": "Camp-Wide"},
        "sessions": SESSIONS_ALL,
    },
    {
        "id": "rob-mautino-adult-education",
        "data": {"name": "Rob Mautino", "area": "Adult Education"},
        "sessions": SESSIONS_ALL,
    },
    {
        "id": "mike-murphy-van-driver",
        "data": {"name": "Mike Murphy", "area": "Van Driver"},
        "sessions": SESSIONS_ALL,
    },
]

for staff in NEW_STAFF:
    print(f"\n{staff['id']}  ({', '.join(staff['sessions'])})")
    for session in staff["sessions"]:
        db.collection('sessions').document(session).collection('staff') \
          .document(staff['id']).set(staff['data'], merge=True)
        print(f"  ✓ {session}")

print("\nDone. Refresh the app.")
