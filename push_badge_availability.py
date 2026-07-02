"""
push_badge_availability.py
Reads a single-week Class_Attendee_Counts CSV, auto-detects the week from the
sheet header, computes spotsTotal/spotsRemaining for each merit badge (column P
open seats floored at 0 — negatives treated as full), and patches
sessions/{week}/meritBadges in Firestore.

Usage:
    python push_badge_availability.py <Class_Attendee_Counts*.csv> <serviceAccountKey.json>

The CSV must contain exactly one "Week N" header row. That week name is used
as the Firestore session doc ID (e.g. "Week 1" → sessions/Week_1).
"""

import csv, sys, re, os
from collections import defaultdict

# ── Class # → badge display name (must match name field in Firestore meritBadges docs) ──
CLASS_TO_BADGE = {
    "CHSM95": "Swimming",
    "CHSM85": "Lifesaving",   "CHSM86": "Lifesaving",
    "CHSM84": "Kayaking",
    "CHSM76": "Canoeing",
    "CHSM89": "Rowing",
    "CHSM27": "First Aid",
    "CHSM92": "Small Boat Sailing",
    "CHSM1":  "Archery",      "CHSM8":  "Archery",      "CHSM9":  "Archery",
    "CHSM42": "Rifle Shooting","CHSM61": "Rifle Shooting",
    "CHSM62": "Rifle Shooting","CHSM63": "Rifle Shooting",
    "CHSM4":  "Shotgun Shooting","CHSM7": "Shotgun Shooting",
    "CHSM20": "Climbing",     "CHSM64": "Climbing",     "CHSM65": "Climbing",
    "CHSM31": "Geocaching",
    "CHSM38": "Orienteering",
    "CHSM80": "Exploration",
    "CHSM90": "Search and Rescue",
    "CHSM79": "Cycling + Multisport",
    "CHSM6":  "Hiking / Backpacking",
    "CHSM24": "Environmental Science",
    "CHSM87": "Mammal Study + Fish & Wildlife",
    "CHSM81": "Forestry",
    "CHSM93": "Soil & Water + Geology",
    "CHSM94": "Space Exploration",
    "CHSM68": "Sustainability",
    "CHSM14": "Astronomy",
    "CHSM74": "Advanced Astronomy",
    "CHSM41": "Pottery",
    "CHSM40": "Photography",
    "CHSM54": "Woodcarving",
    "CHSM36": "Metalwork / Forging",
    "CHSM11": "Advanced Metalworking",
    "CHSM2":  "Trail to Tenderfoot",
    "CHSM3":  "Trail to Second Class",
    "CHSM56": "Trail to First Class",
    "CHSM96": "Wilderness Survival",
    "CHSM15": "Camping",
    "CHSM78": "Citizenship in the World",  "CHSM99": "Citizenship in the World",
    "CHSM102":"Citizenship in the World",
    "CHSM77": "Citizenship in the Nation", "CHSM98": "Citizenship in the Nation",
    "CHSM21": "Communication",             "CHSM97": "Communication",
    "CHSM44": "Salesmanship / Entrepreneurship",
    "CHSM69": "SPL 101",
    "CHSM32": "American Indian Culture",
    "CHSM12": "Scouting / American Heritage",
    "CHSM75": "Archaeology",
    "CHSM72": "Fire Safety",
    "CHSM59": "Welding",
    "CHSM73": "Nuclear Science", "CHSM100": "Nuclear Science",
    "CHSM71": "Electronics",     "CHSM101": "Electronics",
    "CHSM91": "Signs, Signals, and Codes",
    "CHSM50": "Theater",
    "CHSM82": "Game Design + Chess", "CHSM83": "Game Design + Chess",
}


def parse_csv_by_week(csv_path):
    """Returns {week_name: {badge_name: {max, open, perBlock: {block_num: {cap, remaining}}}}}

    Per-block remaining is summed raw (across all sections in the same block) and
    floored at 0 after aggregation — so if Block 6 has two sections that together
    are overbooked, remaining shows 0, not a negative.
    """
    weeks = {}
    current_week = None

    with open(csv_path, newline='', encoding='utf-8-sig') as f:
        for line in f:
            line = line.rstrip('\n')
            # Detect week header line
            m = re.match(r'^(Week \d+)', line)
            if m:
                current_week = m.group(1)
                weeks[current_week] = defaultdict(lambda: {"max": 0, "open": 0, "perBlock": {}})
                continue

            if current_week is None:
                continue

            row = next(csv.reader([line]))
            if len(row) < 16:
                continue
            class_num = row[7].strip()
            if not class_num.startswith("CHS"):
                continue
            badge = CLASS_TO_BADGE.get(class_num)
            if badge is None:
                continue

            try:
                max_s  = int(row[13]) if row[13].strip() else 0
                open_s = int(row[15]) if row[15].strip() else 0
            except ValueError:
                continue

            # Extract block number from period column (e.g. "Block 3" → 3)
            period = row[9].strip() if len(row) > 9 else ''
            bm = re.match(r'Block\s+(\d+)', period)
            block_num = bm.group(1) if bm else None

            # Aggregate totals (open floored per-section)
            weeks[current_week][badge]["max"]  += max_s
            weeks[current_week][badge]["open"] += max(0, open_s)

            # Per-block totals — sum raw first, floor after all sections are in
            if block_num:
                pb = weeks[current_week][badge]["perBlock"]
                if block_num not in pb:
                    pb[block_num] = {"cap": 0, "remaining": 0}
                pb[block_num]["cap"]       += max_s
                pb[block_num]["remaining"] += max(0, open_s)  # floor per-section, same as aggregate

    return weeks


def session_doc_id(week_name):
    return week_name.replace(' ', '_')  # "Week 6" → "Week_6"


def push_to_firestore(weeks, service_account_path):
    import firebase_admin
    from firebase_admin import credentials, firestore as fs
    import time

    if not firebase_admin._apps:
        cred = credentials.Certificate(service_account_path)
        firebase_admin.initialize_app(cred)
    db = fs.client()

    for week_name, badges in sorted(weeks.items()):
        doc_id = session_doc_id(week_name)
        print(f"\n{week_name} ({doc_id})")

        # Read existing meritBadges docs to find IDs by name
        mb_ref = db.collection('sessions').document(doc_id).collection('meritBadges')
        existing = {doc.get('name'): doc.id for doc in mb_ref.stream() if doc.get('name')}

        if not existing:
            print(f"  ⚠️  No meritBadges docs found — skipping (run normal push first?)")
            continue

        updated = 0
        for badge_name, counts in badges.items():
            doc_id_badge = existing.get(badge_name)
            if not doc_id_badge:
                # Try case-insensitive match
                for stored_name, stored_id in existing.items():
                    if stored_name.lower() == badge_name.lower():
                        doc_id_badge = stored_id
                        break

            if not doc_id_badge:
                print(f"  ⚠️  No doc found for: {badge_name}")
                continue

            mb_ref.document(doc_id_badge).update({
                'spotsTotal':     counts['max'],
                'spotsRemaining': counts['open'],
                'perBlock':       counts['perBlock'],
            })
            updated += 1

        print(f"  ✅ Updated {updated} of {len(existing)} badge docs")

        # Stamp lastSync on the session doc so the app header shows the correct time
        now_ms = int(time.time() * 1000)
        db.collection('sessions').document(doc_id).update({
            'lastSync':       now_ms,
            'lastSyncSource': 'Class_Attendee_Counts',
        })
        print(f"  🕒 lastSync updated")


def main():
    if len(sys.argv) != 3:
        print("Usage: python push_badge_availability.py <Class_Attendee_Counts*.csv> <serviceAccountKey.json>")
        sys.exit(1)

    # Expand %USERPROFILE% / %VAR% (CMD) and ~ so paths work from any shell
    csv_path = os.path.expandvars(os.path.expanduser(sys.argv[1]))
    sa_path  = os.path.expandvars(os.path.expanduser(sys.argv[2]))

    print("Parsing CSV...")
    weeks = parse_csv_by_week(csv_path)

    if not weeks:
        print("❌  No week header found in CSV. Expected a line starting with 'Week N'.")
        sys.exit(1)

    if len(weeks) > 1:
        print(f"⚠️  Multiple weeks found: {', '.join(sorted(weeks))}. Expected a single-week CSV.")
        print("    Proceeding — will push all weeks found.")
    else:
        week_name = next(iter(weeks))
        print(f"  Detected: {week_name} ({len(weeks[week_name])} badges)")

    print("\nPushing to Firestore...")
    push_to_firestore(weeks, sa_path)
    print("\nDone.")


if __name__ == '__main__':
    main()
