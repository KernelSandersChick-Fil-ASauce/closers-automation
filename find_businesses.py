"""
The Closers - Phase 2: Automatic Business Finder
Searches Google Places for local businesses, checks each site for booking systems,
and writes everything to results.xlsx.

Usage:
    python3 find_businesses.py "chiropractors in Santa Monica CA" --max 20
    python3 find_businesses.py "med spas in Beverly Hills CA" --max 50
"""

import argparse
import re
import time
import requests
import pandas as pd
from urllib.parse import urljoin, urlparse

# ── YOUR API KEY ───────────────────────────────────────────────────────────────
# Paste your Google Places API key here (see README_API_SETUP.txt for instructions)
GOOGLE_API_KEY = "AIzaSyA-BrHr1cTOjB9-qyPNXb0CEw5BD7eG1g8"
# ──────────────────────────────────────────────────────────────────────────────

OUTPUT_FILE = "results.xlsx"

# Places API (New) endpoints
PLACES_TEXT_SEARCH_URL = "https://places.googleapis.com/v1/places:searchText"
PLACES_DETAILS_URL     = "https://places.googleapis.com/v1/places/{place_id}"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    )
}

EMAIL_RE = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}")

# Reuse booking-check logic from Phase 1
from check_websites import check_website


def find_email(website: str) -> str:
    """
    Try to find a publicly listed email on a business's website.
    Checks the homepage and /contact page. Returns the first email found, or "".
    """
    if not website or not website.strip():
        return ""

    url = website.strip()
    if not url.startswith("http"):
        url = "https://" + url

    # Pages to check in order
    base = f"{urlparse(url).scheme}://{urlparse(url).netloc}"
    pages_to_try = [url, urljoin(base, "/contact"), urljoin(base, "/contact-us"), urljoin(base, "/about")]

    seen = set()
    for page in pages_to_try:
        if page in seen:
            continue
        seen.add(page)
        try:
            r = requests.get(page, headers=HEADERS, timeout=8)
            if r.status_code != 200:
                continue

            # Look for mailto: links first (most reliable)
            from bs4 import BeautifulSoup
            soup = BeautifulSoup(r.text, "html.parser")
            for a in soup.find_all("a", href=True):
                href = a["href"]
                if href.startswith("mailto:"):
                    email = href.replace("mailto:", "").split("?")[0].strip()
                    if email and "@" in email:
                        return email

            # Fall back to regex scan of page text
            emails = EMAIL_RE.findall(r.text)
            for email in emails:
                # Filter out common false positives
                if any(skip in email.lower() for skip in [
                    "example.", "sentry.", "wix.", "squarespace.", "wordpress.",
                    ".png", ".jpg", ".gif", ".svg", ".webp", ".css", ".js",
                    "noreply", "no-reply", "donotreply", "support@wix",
                ]):
                    continue
                # Must have a real TLD (2-6 chars) and no file extensions
                domain_part = email.split("@")[-1]
                if "." not in domain_part:
                    continue
                return email

        except Exception:
            continue

    return ""


def search_places(query: str, max_results: int) -> list[dict]:
    """
    Search Google Places (New API) for businesses matching the query.
    Handles pagination to get up to max_results.
    """
    places = []
    next_page_token = None

    while len(places) < max_results:
        payload = {"textQuery": query, "pageSize": min(20, max_results - len(places))}
        if next_page_token:
            payload["pageToken"] = next_page_token

        headers = {
            "Content-Type": "application/json",
            "X-Goog-Api-Key": GOOGLE_API_KEY,
            "X-Goog-FieldMask": "places.id,places.displayName,places.formattedAddress,places.rating,nextPageToken",
        }

        response = requests.post(PLACES_TEXT_SEARCH_URL, json=payload, headers=headers, timeout=10)
        data = response.json()

        if "error" in data:
            print(f"\nERROR from Google Places: {data['error'].get('message', data['error'])}")
            return []

        places.extend(data.get("places", []))
        next_page_token = data.get("nextPageToken")

        if not next_page_token or len(places) >= max_results:
            break

        time.sleep(2)

    return places[:max_results]


def get_place_details(place_id: str) -> dict:
    """
    Fetch phone number and website for a single place using Place Details (New API).
    """
    url = PLACES_DETAILS_URL.format(place_id=place_id)
    headers = {
        "X-Goog-Api-Key": GOOGLE_API_KEY,
        "X-Goog-FieldMask": "nationalPhoneNumber,websiteUri",
    }
    response = requests.get(url, headers=headers, timeout=10)
    result = response.json()
    return {
        "phone":   result.get("nationalPhoneNumber", ""),
        "website": result.get("websiteUri", ""),
    }


def run(query: str, max_results: int):
    if GOOGLE_API_KEY == "YOUR_API_KEY_HERE":
        print("ERROR: You haven't added your Google API key yet.")
        print("Open find_businesses.py and paste your key where it says YOUR_API_KEY_HERE\n")
        print("See README_API_SETUP.txt for step-by-step instructions.")
        return

    print(f"\nSearching Google Places for: \"{query}\"")
    print(f"Max results: {max_results}\n")

    places = search_places(query, max_results)
    if not places:
        print("No results found.")
        return

    print(f"Found {len(places)} businesses. Fetching details and checking websites...\n")

    results = []
    for i, place in enumerate(places, start=1):
        name    = place.get("displayName", {}).get("text", "")
        address = place.get("formattedAddress", "")
        rating  = place.get("rating", "")

        print(f"[{i}/{len(places)}] {name}")

        # Get phone + website from Place Details API
        details = get_place_details(place["id"])
        phone   = details["phone"]
        website = details["website"]

        # Run booking check from Phase 1
        booking = check_website(website)
        booking_label = (
            "Yes"     if booking["has_booking"] is True
            else "No" if booking["has_booking"] is False
            else "Unknown"
        )

        # Try to find a contact email on their website
        email = find_email(website) if website else ""

        print(f"         Website: {website or '(none)'} | Booking: {booking_label} | Email: {email or '(not found)'}")

        # Only keep leads — businesses without a working booking/contact system
        if booking_label == "Yes":
            print(f"         ✓ Already has booking — skipping")
            continue

        results.append({
            "Business Name":      name,
            "Address":            address,
            "Phone":              phone,
            "Website":            website,
            "Owner Email":        email,
            "Rating":             rating,
            "Has Booking System": booking_label,
            "Notes":              booking["notes"],
        })

        time.sleep(0.5)  # avoid hitting rate limits

    # Append to existing spreadsheet or create new one
    new_df = pd.DataFrame(results)
    try:
        existing_df = pd.read_excel(OUTPUT_FILE)
        # Drop duplicates by business name + address
        combined = pd.concat([existing_df, new_df], ignore_index=True)
        combined = combined.drop_duplicates(subset=["Business Name", "Address"], keep="first")
    except FileNotFoundError:
        combined = new_df

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        combined.to_excel(writer, index=False, sheet_name="Results")
        ws = writer.sheets["Results"]
        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col) + 4
            ws.column_dimensions[col[0].column_letter].width = min(max_len, 60)

    print(f"\nDone! Results saved to '{OUTPUT_FILE}'.")
    no_booking = [r for r in results if r["Has Booking System"] in ("No", "Unknown")]
    print(f"Leads without booking systems: {len(no_booking)} / {len(results)}")
    print(f"Total businesses in spreadsheet: {len(combined)}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Find local businesses and check for booking systems.")
    parser.add_argument("query",       help='Search query, e.g. "chiropractors in Santa Monica CA"')
    parser.add_argument("--max",       type=int, default=20, help="Max number of businesses to check (default: 20)")
    args = parser.parse_args()

    run(args.query, args.max)
