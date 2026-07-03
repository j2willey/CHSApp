"""
push_skills_to_firestore.py
Reads CHS_Open_Program_Skills_of_the_Day.csv and pushes ONE doc per session/day
to Firestore under sessions/{session}/skills/{day}.

Each doc has fields:
  day, title, icon, location, description, steps (list of all activity strings)

Doc IDs are the day abbreviation ("Mon", "Tue", etc.) so re-runs are idempotent
and the app's SKILLS.find(s=>s.day===currentDay) lookup works correctly.

Usage:
    python push_skills_to_firestore.py <CSV_path> <serviceAccountKey.json>

Example:
    python push_skills_to_firestore.py "CHS_Open_Program_Skills_of_the_Day.csv" "%USERPROFILE%\\secrets\\camphisierraapp-deploy-key.json"
"""

import csv, sys, os
from collections import defaultdict, OrderedDict

DAY_ORDER = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]


def session_doc_id(session_raw):
    """'Week_1' stays 'Week_1'; spaces → underscores as a safety net."""
    return session_raw.strip().replace(" ", "_")


def fmt_step(area, activity, description):
    """Format a single activity into a step string."""
    if area:
        step = f"{area}: {activity}"
        if description:
            step += f" — {description}"  # em dash
    else:
        step = activity
        if description:
            step += f" — {description}"
    return step


def parse_csv(csv_path):
    """
    Returns {session_doc_id: OrderedDict{day: {"day":..., "steps":[...]}}}
    sorted by day then by row order (#).
    """
    # Two-level: session -> day -> list of (order, step_string)
    raw = defaultdict(lambda: defaultdict(list))

    with open(csv_path, newline='', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        for row in reader:
            session = session_doc_id(row.get("Session", "").strip())
            day     = row.get("Day", "").strip()
            order   = row.get("#", "").strip()
            area    = row.get("Area", "").strip()
            activity = row.get("Activity", "").strip()
            desc    = row.get("Description", "").strip()

            if not session or not day or not activity:
                continue

            try:
                order_int = int(order)
            except ValueError:
                order_int = 999

            step = fmt_step(area, activity, desc)
            raw[session][day].append((order_int, step))

    # Build final structure: one doc per session/day
    result = {}
    for session_id, days in raw.items():
        result[session_id] = {}
        for day, step_list in days.items():
            # Sort steps by their original order number
            step_list.sort(key=lambda t: t[0])
            steps = [s for _, s in step_list]
            result[session_id][day] = {
                "day":         day,
                "title":       "Open Program",
                "icon":        "\U0001f3d5️",  # 🏕️
                "location":    "All Program Areas",
                "description": "All areas open this afternoon — visit any activity, no sign-up needed!",
                "steps":       steps,
            }

    return result


def push_to_firestore(sessions, service_account_path):
    import firebase_admin
    from firebase_admin import credentials, firestore as fs

    if not firebase_admin._apps:
        cred = credentials.Certificate(service_account_path)
        firebase_admin.initialize_app(cred)
    db = fs.client()

    for session_id, days in sorted(sessions.items()):
        print(f"\n{session_id}  ({len(days)} days)")

        skills_ref = db.collection("sessions").document(session_id).collection("skills")

        # Delete ALL existing skill docs first (removes stale per-activity docs
        # like Mon_01, Mon_02 that were written by the old script format)
        existing = list(skills_ref.stream())
        if existing:
            batch = db.batch()
            for doc in existing:
                batch.delete(doc.reference)
            batch.commit()
            print(f"  \U0001f5d1️  Deleted {len(existing)} existing skill docs")

        # Write one doc per day, doc ID = the day abbreviation
        batch = db.batch()
        count = 0
        for day in sorted(days.keys(), key=lambda d: DAY_ORDER.index(d) if d in DAY_ORDER else 99):
            doc_data = days[day]
            batch.set(skills_ref.document(day), doc_data)
            count += 1
            step_count = len(doc_data["steps"])
            print(f"    {day}: {step_count} activities")

        batch.commit()
        print(f"  ✅ Wrote {count} day docs")


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
    total_days = sum(len(d) for d in sessions.values())
    print(f"  Found {len(sessions)} sessions, {total_days} session-day combos")
    for sid in sorted(sessions):
        days = sorted(sessions[sid].keys(), key=lambda d: DAY_ORDER.index(d) if d in DAY_ORDER else 99)
        print(f"    {sid}: {', '.join(days)}")

    print("\nPushing to Firestore…")
    push_to_firestore(sessions, sa_path)
    print("\nDone.")


if __name__ == "__main__":
    main()
