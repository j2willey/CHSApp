import firebase_admin
from firebase_admin import credentials, firestore

cred = credentials.Certificate(
    r'C:\Users\Michael\Documents\Claude\Projects\CHS app\serviceAccountKey.json'
)
firebase_admin.initialize_app(cred)
db = firestore.client()

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

SESSIONS = [
    "Week_1",
    "Eagle_Achievement_Camp",
    "Week_3",
    "Week_4",
    "Week_5",
    "Week_6",
]

BASE_URL = "https://storage.googleapis.com/camphisierraapp.firebasestorage.app/staff/headshots"

for session in SESSIONS:
    print(f"\nsessions/{session}/staff:")
    for staff_id in STAFF_IDS:
        photo_url = f"{BASE_URL}/{staff_id}.jpg"
        db.collection('sessions').document(session).collection('staff').document(staff_id).set(
            {'photo': photo_url}, merge=True
        )
        print(f"  ✓ {staff_id}")

print("\nDone. Refresh the app to see headshots.")
