"""
push_skills_to_firestore.py
Reads CHS_Open_Program_Skills_of_the_Day.csv and pushes every activity row
to Firestore under sessions/{session}/skills.

Each row → one Firestore doc with fields:
  day, order, area, title, description, icon, location, steps (empty list)

Doc IDs are stable: e.g. Mon_01, Mon_02 … so re-runs are idempotent.

Usage:
    python push_skills_to_firestore.py <CSV_path> <serviceAccountKey.json>

Example:
    python push_skills_to_firestore.py "CHS_Open_Program_Skills_of_the_Day.csv" "%USERPROFILE%\\secrets\\camphisierraapp-deploy-key.json"
"""

import csv, sys, os, re
from collections import defaultdict

AREA_ICONS = {
    "aquatics":           "🏊",
    "nature":             "🌿",
    "foxfire":            "🔥",
    "handicraft":         "🎨",
    "rifle":              "🎯",
    "trail to eagle":     "⭐",
    "climbing":           "🧗",
    "archery":            "🏹",
    "scoutcraft":         "🪢",
    "blackfoot meadow":   "🏕️",
    "the kiosk":          "🏪",
    "excursions":         "🚵",
    "dining hall deck":   "🍽️",
    "program meadow":     "🌄",
    "downriver of nature":"🪓",
    "tbd":                "🗓️",
}

DAY_ORDER = ["Sun","Mon","Tue","Wed","Thu","Fri","Sat"]


def icon_for_area(area):
    if not area:
        return "📅"
    return AREA_ICONS.get(area.lower(), "🏕️")


def session_doc_id(session_raw):
    """'Week_1' stays 'Week_1'; spaces → underscores as a safety net."""
    return session_raw.strip().replace(" ", "_")


def parse_csv(csv_path):
    """Returns {session_doc_id: [row_dict, ...]} sorted by day then order."""
    sessions = defaultdict(list)
    with open(csv_path, newline='', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        for row in reader:
            session = session_doc_id(row.get("Session", "").strip())
            day     = row.get("Day", "").strip()
            order   = row.get("#", "").strip()
            area    = row.get("Area", "").strip()
            activity= row.get("Activity", "").strip()
            desc    = row.get("Description", "").strip()

            if not session or not day or not activity:
                continue

            try:
                order_int = int(order)
            except ValueError:
                order_int = 0

            sessions[session].append({
                "day":         day,
                "order":       order_int,
                "area":        area,
                "title":       activity,
                "description": desc,
                "icon":        icon_for_area(area),
                "location":    area if area else "Camp",
                "steps":       [],
            })

    # Sort each session's list by day then order
    for session in sessions:
        sessions[session].sort(
            key=lambda r: (DAY_ORDER.index(r["day"]) if r["day"] in DAY_ORDER else 99, r["order"])
        )
    return sessions


def push_to_firestore(sessions, service_account_path):
    import firebase_admin
    from firebase_admin import credentials, firestore as fs

    if not firebase_admin._apps:
        cred = credentials.Certificate(service_account_path)
        firebase_admin.initialize_app(cred)
    db = fs.client()

    for session_id, rows in sorted(sessions.items()):
        print(f"\n{session_id}  ({len(rows)} activities)")

        skills_ref = db.collection("sessions").document(session_id).collection("skills")

        # Delete all existing skill docs first so removed activities don't linger
        existing = list(skills_ref.stream())
        if existing:
            batch = db.batch()
            for doc in existing:
                batch.delete(doc.reference)
            batch.commit()
            print(f"  🗑️  Deleted {len(existing)} existing skill docs")

        # Write new docs in batches of 500
        batch = db.batch()
        count = 0
        for row in rows:
            doc_id = f"{row['day']}_{row['order']:02d}"
            batch.set(skills_ref.document(doc_id), row)
            count += 1
            if count % 500 == 0:
                batch.commit()
                batch = db.batch()

        if count % 500 != 0:
            batch.commit()

        print(f"  ✅ Wrote {count} skill docs")


def main():
    if len(sys.argv) != 3:
        print("Usage: python push_skills_to_firestore.py <CSV_path> <serviceAccountKey.json>")
        sys.exit(1)

    csv_path = os.path.expandvars(os.path.expanduser(sys.argv[1]))
    sa_path  = os.path.expandvars(os.path.expanduser(sys.argv[2]))

    if not os.path.exists(csv_path):
        print(f"❌  CSV not found: {csv_path}")
        sys.exit(1)
    if not os.path.exists(sa_path):
        print(f"❌  Service account key not found: {sa_path}")
        sys.exit(1)

    print("Parsing CSV…")
    sessions = parse_csv(csv_path)
    print(f"  Found {len(sessions)} sessions: {', '.join(sorted(sessions))}")

    print("\nPushing to Firestore…")
    push_to_firestore(sessions, sa_path)
    print("\nDone.")


if __name__ == "__main__":
    main()
