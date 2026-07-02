# Upload/overwrite headshots to Firebase Storage and patch all 6 session Firestore docs.
import os, firebase_admin
from firebase_admin import credentials, firestore, storage

SESSIONS = [
    "Week_1", "Eagle_Achievement_Camp",
    "Week_3", "Week_4", "Week_5", "Week_6",
]

RESIZED_DIR = r'C:\Users\Michael\Documents\Claude\Projects\CHS app\headshots_resized'
BUCKET_NAME = 'camphisierraapp.firebasestorage.app'
BASE_URL = f'https://storage.googleapis.com/{BUCKET_NAME}/staff/headshots'
SA_KEY   = r'C:\Users\Michael\Documents\Claude\Projects\CHS app\serviceAccountKey.json'

# All IDs to upload (existing = overwrite, new = create)
UPLOAD_IDS = [
    # Original batch 1 (already done but included for safety)
    "adriana-chapa-handicraft",
    "allison-rahn-climbing",
    "danica-spears-scoutcraft",      # corrected photo
    "devin-bayless-climbing",        # corrected photo
    "isabella-demarco-aquatics",
    "sam-mcpheeters-trail-to-eagle",
    "sean-velasquez-welding",
    # Batch 2 - new staff
    "aaron-velasquez-nature",
    "adam-geilhuffe-excursions",
    "anthony-onodera-pinecrest",
    "austyn-pollard-kitchen",
    "ben-rotte-scoutcraft",
    "ben-summer-nature",
    "caelin-slade-aquatics",         # new
    "chris-brodack-kitchen",
    "chris-brodack-trading-post",
    "cody-bui-foxfire",              # corrected photo
    "daniel-scott-rifle",            # corrected photo (was Danica's file)
    "doug-wiley-maintenance",
    "eamonn-donnelly-kitchen",
    "eamonn-donnelly-trading-post",
    "ian-han-trail-to-eagle",        # swapped
    "ian-horsely-kitchen",           # swapped
    "jake-oxtal-foxfire",
    "john-stephenson-rifle",
    "justin-h-kitchen",
    "kai-grist-kitchen",
    "katherine-rico-welding",        # Sam Rico
    "matthew-lewis-archery",
    "nick-brandis-pinecrest",
    "olivia-englehorn-handicraft",
    "olivier-rollet-kitchen",
    "patrick-loweth-aquatics",
    "philp-notaro-excursions",
    "ryan-kenny-trail-to-eagle",
    "saanvi-sharma-nature",
    "sam-jacobson-foxfire",
    "shivani-agarwala-scoutcraft",
    "stan-searing-climbing",         # corrected photo
    "tobias-bone-aquatics",
    "tristan-sheppy-climbing",
    "victor-sanchez-kitchen",
    # New additions
    "ed-snyder-observatory",         # new staff member
    # Batch 3 - new headshots
    "ben-willey-kitchen",
    "eric-brewer-maintenance",
    "griffen-goza-kitchen",
    "john-h-kitchen",
    "kylie-morgan-handicraft",
    "miguel-l-kitchen",
    "morgan-de-la-torre-trail-to-eagle",
    "lukas-winter-camp-wide",
    "rob-mautino-adult-education",
    "mike-murphy-van-driver",
]

# Brent Nicolai (medic) — Week 3 and Week 5 only
BRENT_ID = "brent-nicolai-medic"
BRENT_SESSIONS = ["Week_3", "Week_5"]

cred = credentials.Certificate(SA_KEY)
firebase_admin.initialize_app(cred, {'storageBucket': BUCKET_NAME})
db = firestore.client()
bucket = storage.bucket()

print(f"Uploading {len(UPLOAD_IDS)} headshots to Firebase Storage...")
for staff_id in UPLOAD_IDS:
    local_path = os.path.join(RESIZED_DIR, f'{staff_id}.jpg')
    if not os.path.exists(local_path):
        print(f'  SKIP (no file): {staff_id}')
        continue
    blob = bucket.blob(f'staff/headshots/{staff_id}.jpg')
    blob.upload_from_filename(local_path, content_type='image/jpeg')
    blob.make_public()
    print(f'  uploaded {staff_id}')

print(f"\nPatching Firestore ({len(SESSIONS)} sessions)...")
for session in SESSIONS:
    print(f'  {session}')
    for staff_id in UPLOAD_IDS:
        photo_url = f'{BASE_URL}/{staff_id}.jpg'
        db.collection('sessions').document(session).collection('staff').document(staff_id).set(
            {'photo': photo_url}, merge=True
        )

# Rename Katherine Rico -> Sam Rico
print("\nRenaming Katherine Rico -> Sam Rico...")
for session in SESSIONS:
    db.collection('sessions').document(session).collection('staff').document('katherine-rico-welding').set(
        {'name': 'Sam Rico'}, merge=True
    )

# Upload and patch Brent (Week 3 & Week 5 only)
print(f"\nUploading {BRENT_ID} (medic, Week 3 & Week 5 only)...")
local_path = os.path.join(RESIZED_DIR, f'{BRENT_ID}.jpg')
if os.path.exists(local_path):
    blob = bucket.blob(f'staff/headshots/{BRENT_ID}.jpg')
    blob.upload_from_filename(local_path, content_type='image/jpeg')
    blob.make_public()
    print(f'  uploaded {BRENT_ID}')
    for session in BRENT_SESSIONS:
        photo_url = f'{BASE_URL}/{BRENT_ID}.jpg'
        db.collection('sessions').document(session).collection('staff').document(BRENT_ID).set(
            {'photo': photo_url}, merge=True
        )
        print(f'  patched {session}')
else:
    print(f'  SKIP (no file): {BRENT_ID}')

print(f"\nDone. Refresh the app to see updated headshots.")
