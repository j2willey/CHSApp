"""
reset_all_attendance.py
Clears spotsTotal, spotsRemaining, and perBlock from every meritBadge doc
across all camp sessions, so attendance counts start fresh before a new push.

Usage:
    python reset_all_attendance.py <serviceAccountKey.json>
"""

import sys
import firebase_admin
from firebase_admin import credentials, firestore
from google.cloud.firestore_v1 import DELETE_FIELD

WEEKS = ["Week_1", "Week_3", "Week_4", "Week_5", "Week_6"]

def main():
    if len(sys.argv) != 2:
        print("Usage: python reset_all_attendance.py <serviceAccountKey.json>")
        sys.exit(1)

    sa_path = sys.argv[1]

    if not firebase_admin._apps:
        cred = credentials.Certificate(sa_path)
        firebase_admin.initialize_app(cred)
    db = firestore.client()

    total_cleared = 0

    for week in WEEKS:
        mb_ref = db.collection("sessions").document(week).collection("meritBadges")
        docs = list(mb_ref.stream())

        if not docs:
            print(f"\n{week}: ⚠️  No meritBadge docs found — skipping")
            continue

        print(f"\n{week} ({len(docs)} badges)")
        cleared = 0
        for doc in docs:
            data = doc.to_dict()
            name = data.get("name", doc.id)
            had_fields = any(f in data for f in ("spotsTotal", "spotsRemaining", "perBlock"))

            doc.reference.update({
                "spotsTotal":     DELETE_FIELD,
                "spotsRemaining": DELETE_FIELD,
                "perBlock":       DELETE_FIELD,
            })

            print(f"  {'✅' if had_fields else '⬜'} {name}")
            if had_fields:
                cleared += 1

        print(f"  → Cleared {cleared}/{len(docs)} docs")
        total_cleared += cleared

    print(f"\nDone. Cleared {total_cleared} badge docs total across {len(WEEKS)} weeks.")

if __name__ == "__main__":
    main()
