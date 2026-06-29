"""
The Closers - Phase 4: Email Sender with Approval Flow
Reviews each draft pitch email one by one, lets you approve/skip/edit,
then sends approved emails via Gmail and tracks sent status in results.xlsx.

Setup (one time only):
    1. Go to myaccount.google.com/apppasswords
    2. Create an app password for "Mail"
    3. Run: export GMAIL_APP_PASSWORD='your-16-char-password'

Usage:
    python3 send_pitches.py
"""

import os
import smtplib
import pandas as pd
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime

# ── Settings ──────────────────────────────────────────────────────────────────

FROM_EMAIL  = "hqsummitos@gmail.com"
INPUT_FILE  = "results.xlsx"
OUTPUT_FILE = "results.xlsx"

# ── Email sender ──────────────────────────────────────────────────────────────

def send_email(app_password: str, to_email: str, subject: str, body: str) -> bool:
    """Send one email via Gmail SMTP. Returns True on success."""
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"]    = FROM_EMAIL
    msg["To"]      = to_email
    msg.attach(MIMEText(body, "plain"))

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(FROM_EMAIL, app_password)
            server.sendmail(FROM_EMAIL, to_email, msg.as_string())
        return True
    except Exception as e:
        print(f"         Send failed: {e}")
        return False


# ── Approval loop ─────────────────────────────────────────────────────────────

def approval_loop(df: pd.DataFrame, app_password: str) -> pd.DataFrame:
    """
    Walk through every unsent row that has a draft email.
    Show each one, wait for S/K/E input, then send or skip.
    """
    # Find rows with a draft email that haven't been sent yet
    has_draft = df["Draft Pitch Email"].apply(
        lambda x: bool(x) and not pd.isna(x) and not str(x).startswith("[Error")
    )
    not_sent = df["Email Status"].apply(
        lambda x: str(x).strip().lower() not in ("sent", "skipped")
        if not pd.isna(x) else True
    )
    pending = df[has_draft & not_sent].index.tolist()

    if not pending:
        print("No pending emails to review. Run generate_pitches.py first to create drafts.")
        return df

    print(f"Found {len(pending)} emails to review.\n")

    sent_count    = 0
    skipped_count = 0

    for i, idx in enumerate(pending, start=1):
        row          = df.loc[idx]
        name         = str(row.get("Business Name", "") or "Unknown Business")
        address      = str(row.get("Address", "") or "")
        website      = str(row.get("Website", "") or "")
        draft_email  = str(row["Draft Pitch Email"])

        print("━" * 60)
        print(f"[{i}/{len(pending)}] {name}")
        if address:
            print(f"Address : {address}")
        if website:
            print(f"Website : {website}")
        print("━" * 60)
        print(draft_email)
        print()

        # Pre-fill email from spreadsheet if we found one automatically
        found_email = str(row.get("Owner Email", "") or "").strip()
        if found_email and "@" in found_email:
            prompt_text = f"Send to [{found_email}] (Enter to confirm, or type a different address, or 'skip'): "
            answer = input(prompt_text).strip()
            if answer.lower() == "skip":
                to_email = ""
            elif answer == "":
                to_email = found_email
            else:
                to_email = answer
        else:
            to_email = input("Send to (email address, or press Enter to skip): ").strip()
        if not to_email:
            df.at[idx, "Email Status"] = "Skipped"
            df.at[idx, "Sent At"]      = ""
            print("Skipped.\n")
            skipped_count += 1
            continue

        # Auto-generate subject line from business name
        default_subject = f"Quick idea for {name}"
        subject_input   = input(f"Subject [{default_subject}]: ").strip()
        subject         = subject_input if subject_input else default_subject

        # Option to edit the body before sending
        print("\n[S]end as-is  [E]dit body  [K]skip")
        choice = input("> ").strip().lower()

        if choice == "k":
            df.at[idx, "Email Status"] = "Skipped"
            df.at[idx, "Sent At"]      = ""
            print("Skipped.\n")
            skipped_count += 1
            continue

        if choice == "e":
            print("\nPaste your edited email body below.")
            print("When done, type END on a new line and press Enter:\n")
            lines = []
            while True:
                line = input()
                if line.strip().upper() == "END":
                    break
                lines.append(line)
            draft_email = "\n".join(lines)
            print("\nUpdated email:")
            print("─" * 40)
            print(draft_email)
            print("─" * 40)
            confirm = input("\nSend this? [Y/N]: ").strip().lower()
            if confirm != "y":
                df.at[idx, "Email Status"] = "Skipped"
                df.at[idx, "Sent At"]      = ""
                print("Skipped.\n")
                skipped_count += 1
                continue
            # Save the edited version back to the spreadsheet
            df.at[idx, "Draft Pitch Email"] = draft_email

        # Send
        print(f"Sending to {to_email}...", end=" ", flush=True)
        success = send_email(app_password, to_email, subject, draft_email)

        if success:
            df.at[idx, "Email Status"] = "Sent"
            df.at[idx, "Sent At"]      = datetime.now().strftime("%Y-%m-%d %H:%M")
            print("Sent!\n")
            sent_count += 1
        else:
            df.at[idx, "Email Status"] = "Failed"
            df.at[idx, "Sent At"]      = ""
            print("Failed — check your app password and try again.\n")

    print("━" * 60)
    print(f"Session complete: {sent_count} sent, {skipped_count} skipped.")
    return df


# ── Main ──────────────────────────────────────────────────────────────────────

def run():
    app_password = os.environ.get("GMAIL_APP_PASSWORD")
    if not app_password:
        print("ERROR: GMAIL_APP_PASSWORD environment variable not set.")
        print("\nTo set it up:")
        print("  1. Go to myaccount.google.com/apppasswords")
        print("  2. Sign in, then create a password for 'Mail'")
        print("  3. Copy the 16-character password it gives you")
        print("  4. Run: export GMAIL_APP_PASSWORD='xxxx xxxx xxxx xxxx'")
        print("  5. Then run this script again.")
        return

    print(f"Reading '{INPUT_FILE}'...")
    df = pd.read_excel(INPUT_FILE)

    # Add tracking columns if they don't exist
    if "Email Status" not in df.columns:
        df["Email Status"] = ""
    if "Sent At" not in df.columns:
        df["Sent At"] = ""

    print(f"Sending from: {FROM_EMAIL}\n")

    df = approval_loop(df, app_password)

    # Save back to Excel
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Results")
        ws = writer.sheets["Results"]
        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col) + 4
            ws.column_dimensions[col[0].column_letter].width = min(max_len, 80)

    print(f"\nResults saved to '{OUTPUT_FILE}'.")


if __name__ == "__main__":
    run()
