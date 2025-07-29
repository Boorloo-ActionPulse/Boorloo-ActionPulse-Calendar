import logging
from uuid import uuid4
from datetime import datetime, timezone
from google.oauth2.service_account import Credentials
import gspread
from ics import Calendar, Event
from ics.grammar.parse import ContentLine
import pytz
import os

# === CONFIGURATION ===
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SERVICE_ACCOUNT_FILE = os.path.join(SCRIPT_DIR, "boorloo-actionpulse-calendar-42fd39789b53.json")
SPREADSHEET_NAME = "BoorlooActionPulseCalendar"
TIMEZONE = "Australia/Perth"
REPO_ROOT = os.path.abspath(os.path.join(SCRIPT_DIR, os.pardir))
OUTPUT_DIR  = os.path.join(REPO_ROOT, "calendar")
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "calendar.ics")
ERROR_LOG_FILE = "generate_calendar_errors.log"

# === SETUP LOGGING ===
logging.basicConfig(filename=ERROR_LOG_FILE, level=logging.WARNING,
                    format="%(asctime)s %(levelname)s %(message)s")

# === SETUP AUTH ===
scopes = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly"
]
credentials = Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=scopes)
gc = gspread.authorize(credentials)

# === READ DATA ===
sheet = gc.open(SPREADSHEET_NAME).sheet1
rows = sheet.get_all_records()

# === GENERATE CALENDAR ===
calendar = Calendar()
calendar.extra.append(
    ContentLine(name="PRODID", value="-//Boorloo ActionPulse//Calendar Generator 1.0//EN")
)
tz = pytz.timezone(TIMEZONE)

for row in rows:
    try:
        # Skip if missing required fields
        if not row.get("Start Date") or not row.get("Start Time"):
            logging.warning(f"Skipping row with missing start: {row}")
            continue

        # Parse start/end with correct format
        start = tz.localize(datetime.strptime(
            f"{row['Start Date']} {row['Start Time']}", "%d/%m/%Y %I:%M %p"))
        end = tz.localize(datetime.strptime(
            f"{row['End Date']} {row['End Time']}", "%d/%m/%Y %I:%M %p"))

        # Skip invalid times
        if end <= start:
            logging.warning(f"Skipping row with end ≤ start: {row}")
            continue

        # Build event
        event = Event(
            name=str(row.get("Title", "Untitled Event")),
            begin=start,
            end=end,
            location=str(row.get("Location", "")),
            description=str(row.get("Description", "")),
            url=str(row.get("URL", ""))
        )

        # Stable UID and timestamps
        now_utc = datetime.now(timezone.utc)
        iso_now = now_utc.strftime("%Y%m%dT%H%M%SZ")
        event.uid = f"{uuid4()}@boorloo-actionpulse.org"
        event.extra.append(ContentLine(name="DTSTAMP", value=iso_now))
        event.extra.append(ContentLine(name="CREATED", value=iso_now))
        event.extra.append(ContentLine(name="LAST-MODIFIED", value=iso_now))

        # Optional recurrence
        if row.get("Recurrence Rule"):
            event.extra.append(
                ContentLine(name="RRULE", value=str(row["Recurrence Rule"]))
            )

        calendar.events.add(event)

    except Exception as e:
        logging.warning(f"Error processing row {row}: {e}")

# === WRITE FILE ===
raw = calendar.serialize()
raw = raw.replace('\r\n', '\n')
raw = raw.replace('\nBEGIN:VEVENT', '\n\nBEGIN:VEVENT')
raw = raw.replace('END:VEVENT\n', 'END:VEVENT\n\n')
import re
raw = re.sub(r'\n{3,}', r'\n\n', raw)
raw = raw.replace('\n', '\r\n')

# make sure dir exists
os.makedirs(OUTPUT_DIR, exist_ok=True)

with open(OUTPUT_FILE, "w", newline="") as f:
    f.write(raw)   # or calendar.serialize()

print(f"ICS file written to {OUTPUT_FILE}")
