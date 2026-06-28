HOW TO GET YOUR FREE GOOGLE PLACES API KEY
==========================================

Google gives you $200/month in free API credits — enough for thousands of searches.

STEP 1 — Create a Google Cloud project
---------------------------------------
1. Go to: https://console.cloud.google.com
2. Sign in with your Google account
3. Click "Select a project" at the top → "New Project"
4. Name it "The Closers" → click Create

STEP 2 — Enable the Places API
--------------------------------
1. In the left menu go to: APIs & Services → Library
2. Search for "Places API"
3. Click it → click "Enable"

STEP 3 — Create your API key
------------------------------
1. Go to: APIs & Services → Credentials
2. Click "+ Create Credentials" → "API Key"
3. Copy the key that appears (looks like: AIzaSy...)

STEP 4 — Restrict your key (recommended for security)
-------------------------------------------------------
1. Click "Edit API Key" after creating it
2. Under "API restrictions" → select "Restrict key"
3. Choose "Places API" from the dropdown
4. Click Save

STEP 5 — Add your key to the script
-------------------------------------
1. Open find_businesses.py
2. Find this line near the top:
       GOOGLE_API_KEY = "YOUR_API_KEY_HERE"
3. Replace YOUR_API_KEY_HERE with your actual key:
       GOOGLE_API_KEY = "AIzaSyXXXXXXXXXXXXXXXXX"
4. Save the file

STEP 6 — Run it!
-----------------
In Terminal:
    cd "/Users/natannegash/Closers Automation"
    python3 find_businesses.py "chiropractors in Santa Monica CA" --max 20

That's it! Results will appear in results.xlsx.


PRICING NOTE
------------
Each search costs about $0.017 (less than 2 cents).
The free $200/month credit = roughly 11,000 searches free.
You won't be charged unless you exceed that and have billing enabled.
