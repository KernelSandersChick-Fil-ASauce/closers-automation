"""
The Closers - Phase 3: Personalized Cold Email Generator
Reads results.xlsx, scrapes real content from each lead's website,
and uses Claude AI to write a short personalized pitch email.

Usage:
    python3 generate_pitches.py

Requires:
    ANTHROPIC_API_KEY environment variable (get one at console.anthropic.com)
"""

import os
import time
import requests
import pandas as pd
from bs4 import BeautifulSoup
import anthropic

# ── Settings ──────────────────────────────────────────────────────────────────

INPUT_FILE  = "results.xlsx"
OUTPUT_FILE = "results.xlsx"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    )
}

# ── Website content scraper ───────────────────────────────────────────────────

def scrape_website_summary(url: str) -> dict:
    """
    Pull a small snapshot of real content from a business's website:
    page title, meta description, and first ~600 chars of visible text.
    Returns a dict with keys: title, description, body_text.
    """
    empty = {"title": "", "description": "", "body_text": ""}

    if not url or not url.strip():
        return empty

    url = url.strip()
    if not url.startswith("http"):
        url = "https://" + url

    try:
        response = requests.get(url, headers=HEADERS, timeout=10)
        response.raise_for_status()
    except Exception:
        return empty

    soup = BeautifulSoup(response.text, "html.parser")

    # Page title
    title = soup.title.string.strip() if soup.title and soup.title.string else ""

    # Meta description
    meta = soup.find("meta", attrs={"name": "description"})
    description = meta.get("content", "").strip() if meta else ""

    # Visible body text — strip scripts/styles/nav first
    for tag in soup(["script", "style", "nav", "footer", "header"]):
        tag.decompose()
    body_text = " ".join(soup.get_text(separator=" ").split())[:600]

    return {
        "title":       title[:120],
        "description": description[:200],
        "body_text":   body_text,
    }


# ── Email generator ───────────────────────────────────────────────────────────

def generate_email(client: anthropic.Anthropic, business: dict, site: dict) -> str:
    """
    Use Claude to write a short, personalized cold email for one lead.
    """
    name    = str(business.get("Business Name", "") or "")
    address = str(business.get("Address", "") or "")
    gap     = str(business.get("Has Booking System", "No") or "No")
    notes   = str(business.get("Notes", "") or "")
    website = str(business.get("Website", "") or "")

    has_website = bool(website.strip())

    # Build context string from scraped content
    site_context = ""
    if site.get("title"):
        site_context += f"Page title: {site['title']}\n"
    if site.get("description"):
        site_context += f"Meta description: {site['description']}\n"
    if site.get("body_text"):
        site_context += f"Website excerpt: {site['body_text'][:400]}\n"

    prompt = f"""You write cold outreach emails for a small business called "The Closers."
The Closers helps local businesses of ANY type get more customers. Two main services:
1. An AI receptionist (Bland.ai) — answers calls 24/7, books appointments or takes messages automatically, so the business never misses a lead
2. A simple professional landing page (Carrd) — for businesses with no web presence at all

The two solutions aren't one-size-fits-all. Think about what actually makes sense for this specific type of business:
- If they have NO website: pitch the landing page (Carrd)
- If they have a website but no way to contact/book them: pitch the AI receptionist OR a contact form — whichever fits their industry better. For example, a restaurant needs reservations more than "appointments." An auto shop needs people to call or book a drop-off. A law firm needs a contact form. Use your judgment.

Here is the lead:

Business name: {name}
Location: {address}
Has a website: {"Yes" if has_website else "No"}
Booking/contact gap found: {gap}
Checker notes: {notes}
{site_context}

Write a SHORT cold email (4-6 sentences max) that:
- First, silently figure out what type of business this is from the info above — do NOT state this in the email
- Opens with one sentence referencing something real and specific about their business (from the website content above). NOT generic filler like "I came across your business online." Something that shows you actually looked.
- Names the specific gap in one plain sentence (e.g. "I noticed there's no way to book online" or "couldn't find a way to reach you from the site")
- Briefly explains what The Closers can set up for them and why it matters for their specific type of business
- Ends with a simple, low-pressure CTA (e.g. "Worth a quick 10-min call?")
- Sounds like a real person — conversational, direct, no marketing buzzwords, no emojis
- Does NOT include a subject line, sign-off, or placeholder brackets like [Name]

Just write the email body. Nothing else."""

    response = client.messages.create(
        model="claude-opus-4-8",
        max_tokens=400,
        messages=[{"role": "user", "content": prompt}],
    )

    return response.content[0].text.strip()


# ── Main ──────────────────────────────────────────────────────────────────────

def run():
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        print("ERROR: ANTHROPIC_API_KEY environment variable not set.")
        print("Get a free API key at: https://console.anthropic.com")
        print("Then run:  export ANTHROPIC_API_KEY='your-key-here'")
        print("And run this script again.")
        return

    client = anthropic.Anthropic(api_key=api_key)

    print(f"Reading '{INPUT_FILE}'...")
    df = pd.read_excel(INPUT_FILE)

    if "Draft Pitch Email" not in df.columns:
        df["Draft Pitch Email"] = ""

    # Only process leads without a booking system
    leads = df[df["Has Booking System"].isin(["No", "Unknown"])].index.tolist()
    print(f"Found {len(leads)} leads to write pitches for.\n")

    if not leads:
        print("No leads to process. Run find_businesses.py first to find some.")
        return

    for i, idx in enumerate(leads, start=1):
        row = df.loc[idx]
        name    = str(row.get("Business Name", "") or "")
        website = str(row.get("Website", "") or "") if pd.notna(row.get("Website", "")) else ""

        # Skip rows that already have a valid pitch email
        raw = df.at[idx, "Draft Pitch Email"]
        existing = "" if pd.isna(raw) else str(raw).strip()
        if existing and not existing.startswith("[Error"):
            print(f"[{i}/{len(leads)}] {name} — already done, skipping")
            continue

        print(f"[{i}/{len(leads)}] {name}")

        # Scrape website for real content
        print(f"         Scraping website...", end=" ", flush=True)
        site = scrape_website_summary(website)
        print("done")

        # Generate email with Claude
        print(f"         Writing email...", end=" ", flush=True)
        try:
            email = generate_email(client, dict(row), site)
            df.at[idx, "Draft Pitch Email"] = email
            print("done")
        except Exception as e:
            df.at[idx, "Draft Pitch Email"] = f"[Error generating email: {e}]"
            print(f"error: {e}")

        time.sleep(0.5)  # be gentle on rate limits

    # Write updated Excel file
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Results")
        ws = writer.sheets["Results"]
        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col) + 4
            ws.column_dimensions[col[0].column_letter].width = min(max_len, 80)

    print(f"\nDone! Pitch emails saved to '{OUTPUT_FILE}'.")
    print("Review and edit each email before sending — they're drafts, not final copy.")


if __name__ == "__main__":
    run()
