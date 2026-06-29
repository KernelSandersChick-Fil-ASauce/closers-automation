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
    page title, meta description, and first ~400 words of visible text.
    Returns a dict with keys: title, description, body_text, services.
    """
    empty = {"title": "", "description": "", "body_text": "", "services": ""}

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

    # Visible body text — remove scripts/styles first
    for tag in soup(["script", "style", "nav", "footer", "header"]):
        tag.decompose()
    body_text = " ".join(soup.get_text(separator=" ").split())[:600]

    # Look for service-related words in the text
    service_keywords = [
        "chiropractic", "massage", "facial", "botox", "filler", "laser",
        "wellness", "acupuncture", "physical therapy", "cosmetic", "spa",
        "restaurant", "salon", "barbershop", "dental", "optometry",
        "auto repair", "gym", "pilates", "yoga", "med spa", "medspa",
        "dermatology", "aesthetic"
    ]
    found_services = [kw for kw in service_keywords if kw in body_text.lower()]
    services = ", ".join(found_services[:4]) if found_services else ""

    return {
        "title":       title[:120],
        "description": description[:200],
        "body_text":   body_text,
        "services":    services,
    }


# ── Email generator ───────────────────────────────────────────────────────────

def generate_email(client: anthropic.Anthropic, business: dict, site: dict) -> str:
    """
    Use Claude to write a short, personalized cold email for one lead.
    """
    name    = business.get("Business Name", "")
    address = business.get("Address", "")
    gap     = business.get("Has Booking System", "No")
    notes   = business.get("Notes", "")
    website = business.get("Website", "")

    # Determine which solution to pitch
    if not website or website.strip() == "":
        solution_focus = (
            "They don't have a website at all. Pitch a simple, professional "
            "landing page built with Carrd — looks great, takes 1-2 days, and "
            "costs almost nothing."
        )
    else:
        solution_focus = (
            "They have a website but no online booking or contact form. Pitch "
            "an AI receptionist using Bland.ai that answers calls 24/7 and "
            "books appointments automatically, so they never miss a new customer."
        )

    # Build context string from scraped content
    site_context = ""
    if site["title"]:
        site_context += f"Page title: {site['title']}\n"
    if site["description"]:
        site_context += f"Meta description: {site['description']}\n"
    if site["services"]:
        site_context += f"Services mentioned: {site['services']}\n"
    if site["body_text"]:
        site_context += f"Website excerpt: {site['body_text'][:400]}\n"

    prompt = f"""You write cold outreach emails for a small business called "The Closers."
The Closers helps local businesses get more customers by setting up either:
1. An AI receptionist (using Bland.ai) that answers calls and books appointments automatically
2. A simple, professional landing page (using Carrd) if they don't have a good website

Here is a lead you need to write an email for:

Business name: {name}
Location: {address}
Gap found: {gap} booking/contact system
Checker notes: {notes}
{site_context}

What to pitch for THIS lead:
{solution_focus}

Write a SHORT cold email (4-6 sentences max) that:
- Opens by referencing something real and specific about their business (use the website content above — NOT generic filler like "I came across your business")
- Names the specific gap you found in one sentence
- Briefly explains what The Closers does and why it matters for them specifically
- Ends with a simple, low-pressure call to action (e.g., "Worth a quick 10-min call?")
- Sounds like a real person, NOT a marketing email — conversational, no buzzwords, no emojis
- Does NOT include a subject line, signature, or placeholders like [Name]

Just write the email body text. Nothing else."""

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
        name    = row.get("Business Name", "")
        website = row.get("Website", "") if pd.notna(row.get("Website")) else ""

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
