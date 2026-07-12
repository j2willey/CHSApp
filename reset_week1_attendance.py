"""
reset_week1_attendance.py
Clears spotsTotal and spotsRemaining from every meritBadge doc in
sessions/Week_1/meritBadges so attendance counts start fresh.

Usage:
    python reset_week1_attendance.py <serviceAccountKey.json>
"""

import sys
import firebase_admin
from firebase_admin import credentials, firestore
from google.cloud.firestore_v1 import DELETE_FIELD

def main():
    if len(sys.argv) != 2:
        print("Usage: python reset_week1_attendance.py <serviceAccountKey.json>")
        sys.exit(1)

    sa_path = sys.argv[1]

    if not firebase_admin._apps:
        cred = credentials.Certificate(sa_path)
        firebase_admin.initialize_app(cred)
    db = firestore.client()

    mb_ref = db.collection("sessions").document("Week_1").collection("meritBadges")
    docs = list(mb_ref.stream())

    if not docs:
        print("⚠️  No meritBadge docs found under sessions/Week_1/meritBadges")
        sys.exit(1)

    print(f"Found {len(docs)} merit badge docs — clearing attendance counts...\n")

    cleared = 0
    for doc in docs:
        data = doc.to_dict()
        name = data.get("name", doc.id)
        had_fields = "spotsTotal" in data or "spotsRemaining" in data

        doc.reference.update({
            "spotsTotal":     DELETE_FIELD,
            "spotsRemaining": DELETE_FIELD,
        })

        status = "cleared" if had_fields else "no counts (skipped update)"
        print(f"  {'✅' if had_fields else '⬜'} {name}: {status}")
        if had_fields:
            cleared += 1

    print(f"\nDone. Cleared attendance counts from {cleared}/{len(docs)} docs.")

if __name__ == "__main__":
    main()
