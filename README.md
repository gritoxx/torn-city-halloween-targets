🎃 **Halloween Target List — Google Sheets script** by Ray Welker · **v1.1**

Builds a sortable list of Halloween targets in your own Google Sheet: **name, user ID, profile link, FF score, estimated total stats, level, faction, status, and time since last action.**

**How it works**
• Pulls targets in your FF/level range from **FFScouter's API**
• Optionally looks up faction + status with the **Torn API** ('user' → 'profile')
• Set filters (min/max FF, level, inactive only, factionless) in the Settings tab
• A full list of 50 takes about a minute

**Setup**
1. New Google Sheet → Extensions → Apps Script → paste the script → Save
2. Reload the sheet → **Halloween Targets** menu → Set API key
3. Menu → Refresh list
Needs a **Limited** key registered on ffscouter.com (the one FF Scouter uses).

**Rules & API guidelines**
• API only — no scraping, never loads torn.com pages
• Read-only — you attack by hand from the profile link
• Stays under Torn's 100 requests/min, stops on any key error
• Requests tagged "HalloweenSheet" so you can see them in your key log
• Everyone runs **their own copy** — I never see your key or data

**Terms of Use**
• Data Storage: Persistent – until you delete it from your sheet
• Data Sharing: Nobody
• Purpose: Competitive advantage [personal Halloween target list]
• Key Storage & Sharing: Stored in your private Apps Script settings / Not shared — sent only to ffscouter.com and api.torn.com
• Key Access Level: Limited (FFScouter requirement); Torn calls use 'profile' only
• FFScouter Data Policy & Terms: <https://ffscouter.com>

**Changelog**
• **v1.1** — Profile links now work in non-English Google Sheets (no more HYPERLINK errors). Paste the new code over the old, then run Refresh list.
• **v1.0** — First release.

Full details in the **About** tab once installed. Share your sheet view-only — editors can change the script.
