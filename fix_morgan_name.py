# Shorten Morgan de la Torre's display name to "Morgan T." in Firestore.
import firebase_admin
from firebase_admin import credentials, firestore

SA_KEY   = r'C:\Users\Michael\Documents\Claude\Projects\CHS app\serviceAccountKey.json'
STAFF_ID = 'morgan-de-la-torre-trail-to-eagle'
SESSIONS = ["Week_1","Eagle_Achievement_Camp","Week_3","Week_4","Week_5","Week_6"]

cred = credentials.Certificate(SA_KEY)
firebase_admin.initialize_app(cred)
db = firestore.client()

# "Morgan Torre" → shortName() → "Morgan T." in the app
NEW_NAME = "Morgan Torre"

db.collection('staff').document(STAFF_ID).set({'name': NEW_NAME}, merge=True)
print(f'  ✓ staff/{STAFF_ID}')

for session in SESSIONS:
    ref = db.collection('sessions').document(session).collection('staff').document(STAFF_ID)
    if ref.get().exists:
        ref.set({'name': NEW_NAME}, merge=True)
        print(f'  ✓ {session}')

print('\nDone.')
