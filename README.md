# HYROX 12-Week Training Logger

A mobile-first workout logger that uses Google Sheets as the backend. Open it on your phone, tap a session, fill in what you did, and see your progress over 12 weeks.

## Setup (10 minutes, one time)

### Step 1: Create a Google Sheet
1. Go to [sheets.google.com](https://sheets.google.com) and create a new blank spreadsheet
2. Name it something like "HYROX Training Log"

### Step 2: Open Apps Script
1. In the spreadsheet, go to **Extensions > Apps Script**
2. This opens the script editor in a new tab

### Step 3: Add the backend code
1. In the script editor, you'll see a file called `Code.gs` with an empty function
2. **Delete everything** in that file
3. Copy the entire contents of the `Code.gs` file from this folder and paste it in

### Step 4: Add the frontend
1. In the script editor, click the **+** next to "Files" on the left
2. Select **HTML**
3. Name it `Index` (not Index.html — Apps Script adds the .html automatically)
4. **Delete everything** in the new file
5. Copy the entire contents of `Index.html` from this folder and paste it in

### Step 5: Initialize the sheets
1. In the script editor, select `setupSheets` from the function dropdown (next to the Run button)
2. Click **Run**
3. It will ask you to authorize — click through the permissions (it needs access to the spreadsheet)
4. This creates all the sheets (Dashboard, RunLog, KBLog, StationLog, Benchmarks) pre-filled with your 12-week program

### Step 6: Deploy as a web app
1. Click **Deploy > New deployment**
2. Click the gear icon next to "Select type" and choose **Web app**
3. Set "Execute as" to **Me**
4. Set "Who has access" to **Only myself** (or "Anyone with the link" if you want to share)
5. Click **Deploy**
6. Copy the web app URL it gives you

### Step 7: Add to your phone
1. Open the URL on your phone's browser (Safari on iPhone, Chrome on Android)
2. **iPhone**: Tap the share button > "Add to Home Screen"
3. **Android**: Tap the three dots menu > "Add to Home screen"
4. Now it works like an app — tap the icon to open it

## How to use

- **Week tab**: See all 6 sessions for the current week. Tap any session to log it. Green check = logged.
- **Log session**: Fill in what you actually did. Pre-filled with planned workouts. Tap the emoji for how it felt. Hit Save.
- **Progress tab**: Charts showing your Tuesday pace, KB weights, session completion, long run distances, and TGU weights over 12 weeks.
- **Data tab**: "Initialize Sheets" button (first time only) and a link to view the raw Google Sheet.

## Updating the deployment

If I make changes to the code and you need to update:
1. Open Apps Script (Extensions > Apps Script from the Sheet)
2. Replace the contents of Code.gs and/or Index.html
3. Click **Deploy > Manage deployments**
4. Click the pencil icon on your deployment
5. Change "Version" to **New version**
6. Click **Deploy**

The URL stays the same — no need to re-bookmark.
