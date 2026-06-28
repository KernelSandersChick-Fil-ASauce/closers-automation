"""
The Closers - Website Booking Checker
Reads businesses from businesses.csv, checks each site for booking/contact forms,
and writes results to results.xlsx.
"""

import csv
import time
import requests
import pandas as pd
from bs4 import BeautifulSoup

# ── Settings ──────────────────────────────────────────────────────────────────

INPUT_FILE  = "businesses.csv"
OUTPUT_FILE = "results.xlsx"

# Words that suggest a booking system or contact form exists
BOOKING_KEYWORDS = [
    "book", "booking", "schedule", "appointment", "reserve",
    "contact us", "contact form", "request a consultation",
    "book now", "book online", "schedule online", "calendly",
    "acuityscheduling", "mindbody", "vagaro", "square appointments",
    "zocdoc", "fresha", "get in touch", "send us a message",
]

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    )
}

# ── Core logic ────────────────────────────────────────────────────────────────

def check_website(url: str) -> dict:
    """
    Visit a URL and look for signs of a booking/contact system.
    Returns a dict with: has_booking (bool), notes (str).
    """
    if not url or not url.strip():
        return {"has_booking": None, "notes": "No website URL provided"}

    url = url.strip()
    if not url.startswith("http"):
        url = "https://" + url

    try:
        response = requests.get(url, headers=HEADERS, timeout=10)
        response.raise_for_status()
    except requests.exceptions.ConnectionError:
        return {"has_booking": None, "notes": "Could not connect to website"}
    except requests.exceptions.Timeout:
        return {"has_booking": None, "notes": "Website timed out"}
    except requests.exceptions.HTTPError as e:
        return {"has_booking": None, "notes": f"HTTP error: {e}"}
    except Exception as e:
        return {"has_booking": None, "notes": f"Error: {e}"}

    soup = BeautifulSoup(response.text, "html.parser")
    page_text = soup.get_text(separator=" ").lower()

    # Check for booking keywords in visible page text
    found_keywords = [kw for kw in BOOKING_KEYWORDS if kw in page_text]

    # Check for <form> tags (contact/booking forms)
    forms = soup.find_all("form")

    # Check for iframe embeds (Calendly, Acuity, etc.)
    iframes = soup.find_all("iframe")
    booking_iframes = [
        f.get("src", "") for f in iframes
        if any(s in (f.get("src", "") or "").lower()
               for s in ["calendly", "acuity", "mindbody", "vagaro", "square"])
    ]

    has_booking = bool(found_keywords or forms or booking_iframes)

    notes_parts = []
    if found_keywords:
        notes_parts.append(f"Keywords found: {', '.join(found_keywords[:5])}")
    if forms:
        notes_parts.append(f"{len(forms)} form(s) on page")
    if booking_iframes:
        notes_parts.append(f"Booking embed: {booking_iframes[0][:60]}")
    if not notes_parts:
        notes_parts.append("No booking signals detected")

    return {"has_booking": has_booking, "notes": "; ".join(notes_parts)}


def run():
    print(f"Reading businesses from '{INPUT_FILE}'...\n")

    rows = []
    with open(INPUT_FILE, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            rows.append(row)

    results = []
    for i, row in enumerate(rows, start=1):
        name = row.get("business_name", "").strip()
        url  = row.get("website_url", "").strip()

        print(f"[{i}/{len(rows)}] Checking: {name} ({url or 'no URL'}) ...", end=" ", flush=True)
        result = check_website(url)

        booking_label = (
            "Yes" if result["has_booking"] is True
            else "No" if result["has_booking"] is False
            else "Unknown"
        )

        print(booking_label)

        results.append({
            "Business Name":    name,
            "Website URL":      url,
            "Has Booking System": booking_label,
            "Notes":            result["notes"],
        })

        time.sleep(1)  # be polite — wait 1 second between requests

    df = pd.DataFrame(results)

    # Write once with auto-sized columns
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Results")
        ws = writer.sheets["Results"]
        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col) + 4
            ws.column_dimensions[col[0].column_letter].width = min(max_len, 60)

    print(f"\nDone! Results saved to '{OUTPUT_FILE}'.")
    no_booking = [r for r in results if r["Has Booking System"] in ("No", "Unknown")]
    print(f"Leads without booking systems: {len(no_booking)} / {len(results)}")


if __name__ == "__main__":
    run()
