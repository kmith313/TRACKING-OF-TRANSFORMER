# Transformer Tracking + Complaint / PR Status (GitHub Pages)

The frontend runs on GitHub Pages; the backend is a Google Apps Script that you deploy once.
Each organisation gives its own Google Sheet link to the app, and the app builds everything else.

## Folder
- All the files next to `index.html` (`index.html`, `complain.js`, `config.js`, `sw.js`, `manifest.webmanifest`, `icon-192.png`, `icon-512.png`, `emblem.png`) : upload them to GitHub, no sub-folder needed
- `apps-script/Code.gs` : backend code to paste into Google Apps Script

## Setup (developer, one time)
**Step 1 - Backend**
1. Go to https://script.google.com and create a new project.
2. Paste the whole of `apps-script/Code.gs` (remove the default code) and Save.
3. Deploy -> New deployment -> type **Web app** -> Execute as: **Me**, Who has access: **Anyone** -> Deploy.
4. Allow the permissions (Google Sheets / Drive). Copy the resulting `.../exec` URL.

**Step 2 - Config**
Open `config.js` and replace `PASTE_YOUR_APPS_SCRIPT_EXEC_URL_HERE` with that URL.

**Step 3 - Google sign-in for automatic sharing (recommended)**
This is what lets the Admin do nothing except enter the link, workshops, ID and password.
1. Go to https://console.cloud.google.com, create a project, and enable the **Google Drive API**.
2. APIs & Services -> OAuth consent screen: User type **External**, fill in the app name and your email, add the scope `.../auth/drive`, and add your admins as **Test users** (or publish the app).
3. Credentials -> Create credentials -> **OAuth client ID** -> type **Web application** -> Authorized JavaScript origins: `https://<username>.github.io` -> Create.
4. Copy the Client ID into `config.js` (`TT_GOOGLE_CLIENT_ID`).
Note: until Google verifies the app, the sign-in screen shows "Google hasn't verified this app" - the Admin presses Advanced -> Continue. Unverified apps are limited to 100 test users.
If you skip this step, the app still works, but the Admin must share the sheet manually (Anyone with the link -> Editor).

**Step 4 - GitHub**
1. Create a new repository and upload the files of this folder.
2. Settings -> Pages -> Deploy from branch -> `main` / root.
3. After a minute the app opens at `https://<username>.github.io/<repo>/`. Share this link for installing (Add to Home screen).

After changing Code.gs: Deploy -> Manage deployments -> Edit -> New version.

## How the app works
**Home page (the only page everybody sees)**
- **Complaint status** (no login): the person selects their **DISCOM**, types a complaint / PR number (3+ characters) and sees the status.
  The search covers the data of **every Division Head (Admin) registered under that DISCOM**, whichever workshop pasted it.
- **Staff login**: Admin and users log in with just ID + password. Nothing else to select or paste.
- Small text at the bottom: **Division Head? Setup here**.

**Division Head (Admin) setup**
1. Press **Setup here**, then **Sign in with Google** (the Gmail that owns the Google Sheet) and allow access.
2. Paste the Google Sheet link, enter the **DISCOM**, the names of all workshops, and the **Profit Center**. The Profit Center is always the Admin ID. Choose a password and press **Create**.
3. The app shares the sheet, creates the tabs `DATA`, `Users`, `Drivers`, `Config`, `ARCHIVE` (data only up to **REPLACEMENT DATE + TIME**), registers the Admin, and logs in.
4. In the **Users** tab the Admin creates users (each ID must be unique, e.g. start with the Profit Center). Users log in directly from the home page.

**Pasting data**: columns in this order: Workshop, DIVISION, SUBSTATION, ... TIME. A leading **DISCOM** column (before Workshop) is accepted and ignored; the DISCOM is the one entered at setup.

## Who has installed it? (Registry)
Every setup adds a row to a Google Sheet in YOUR Drive named **TF-Tracker - Registry** (tab *Admins*): time, the Admin's verified Google email, Profit Center, DISCOM, workshops and Sheet link.
The tab *Users* maps every User ID to its sheet (this is how login works with just ID + password). The registry is created automatically on the first setup.
To find it: in Apps Script, select the function `registryUrl` and press Run; the link appears in the Execution log.
For anyone with a Gmail to be able to become an Admin, publish the OAuth app (Google Auth Platform -> Audience -> Publish app); in Testing mode only listed Test users can sign in.

## Forgot password
- **Admin:** home page -> *Staff login* -> **Admin forgot password?** -> enter the Profit Center, a new password, then sign in with the *same Google account used at setup*. Google confirms the e-mail (only e-mail is checked, no Drive access) and the new password is set.
  If the Admin was created before e-mail tracking existed, press *Setup here* once again with the same sheet link, Profit Center and password - this saves the e-mail.
- **User:** the Admin opens *Users* and resets the password; the user logs in with the temporary password and must choose a new one.

## Summary / speed
- Summary now shows Return (Today / Total) and a card layout per workshop; the old table is under *Show as table*.
- Data Sheet has a RETURN DATE column (19th column, added automatically to older sheets).
- Last views (summary, data sheet, driver list) are saved in the browser and shown instantly while fresh data loads.
- Driver and JE numbers have a **Call** button in pending lists and result lists.

## Complaint status: how it is so fast
- No special paste format is needed. The Admin keeps pasting into the same DATA columns (Workshop ... REPLACEMENT DATE, TIME, RETURN DATE).
- When a DISCOM is chosen, the browser downloads a small index of that DISCOM (complaint / PR no, place, substation, feeder, capacity, dates). It is saved on the phone (IndexedDB) and searched locally while typing, so results appear instantly.
- On the next visit only a tiny version check runs; the index is downloaded again only when data changed (or after 30 minutes).
- JE / driver names and mobile numbers are NOT in the index. They are fetched from the server only when a complaint is opened.
- If a DISCOM has more than 40,000 rows, or the index is not ready yet, the server searches instead (slower, same result).
- To turn the local index off, add `window.TT_LOCAL_SEARCH = false;` to config.js.
- Keep the DATA tab clean: no thousands of empty formatted rows, one row per PR. Rows older than 2 years are moved to ARCHIVE by the Admin tools, which keeps everything fast.

## Please note
- The Sheet is shared as "Anyone with the link", so anyone who has the link can open the Sheet directly (the Users tab also holds password hashes). Share the link only with your own people and use strong passwords.
- Free Google accounts have daily Apps Script limits; for many organisations a Workspace account is better.

## Install app (PWA)
Home page / header par **Install app** button aata hai (Chrome/Edge desktop + Android). iPhone Safari me Share -> Add to Home Screen ka help dikhta hai.

## Capacity & Pending for return
- CAPACITY cell me jo bhi text ho ("25 KVA 11/0.4 KV TRANSFORMER REPAIRED"), app KVA number (25) khud nikaal leta hai.
- Workshop poora naam (e.g. "EWC Azamgarh") paste karo to registered workshop naam me badal jata hai.
- Capacity-wise summary me **PENDING FOR RETURN** = replacement ho chuka lekin RETURN DATE abhi nahi.

## Comparison tab
From–To dates ke PR (PR DATE) ko pichhle saal ki **same dates** se compare karta hai. 4 type: Division wise, Substation wise (division naam left me), Day wise (har din ka this year vs last year), Workshop wise. Capacity 10/16/25/63, <100 KVA, 100..1000, ≥100 KVA, TOTAL aur Inc/Dec %. Kisi bhi number par click karke list khul jati hai. Division wise summary Summary me ab chhote icon ke peeche (click par khulti hai).

## DISCOM head, Creator, Division name, delete protection
- **Creator** = Google account jo Apps Script (Code.gs) ka owner hai. Home page par Creator panel me Google sign-in karke kisi bhi DISCOM head ko email se access de sakta hai. Ek **one-time password** registry sheet ("DiscomHeads" tab) me rehta hai aur head ko email bhi hota hai. Head pehli login par naya password set karta hai.
- **DISCOM head** (ID jaise `DH-...`) apne poore DISCOM ka data ek saath dekhta hai: Workshop division (EWD) -> workshops, Today PR/Issue, Till date PR/Issue/Pending issue, capacity wise (Pending issue, PR, Average) aur division x capacity. Return / pending replacement nahi dikhte. Data Sheet / Driver / Users tabs nahi hote.
- DISCOM head ka password bhoolne par Division Head ki tarah Google email verification se reset.
- **Division Head** ko pehli login par apne division ka naam (EWD se shuru) dena padta hai.
- Date paste: `10/18/2026` (ERP, MM/DD/YYYY) -> `18.10.2026`.
- Workshop naam me EWC/EWD alag ho to bhi wahi workshop maana jata hai. Galat workshop ki rows paste hone par warning aati hai aur wo save nahi hotin.
- Google Drive me sheet file ka delete roka nahi ja sakta; "Delete all data" action par email OTP + password lagta hai.

## Daily backup, restore, REMARK
- **Backup:** App Creator panel -> "Daily backup" -> *Turn on* (a nightly trigger). Har organisation ki **ek hi backup file** roz update hoti hai (DATA, ARCHIVE, Drivers, Users, Config). Agar sheet me achanak records bahut kam ho jayein (<70%), backup **overwrite nahi hota** ("HELD") - taaki galti se khali sheet purana backup na mita de. Backup files Creator ke Drive me bante hain (script Creator ke account se chalti hai).
- **Data delete ho gaya:** Admin -> Users (Admin) tab -> "Daily backup & restore" -> Gmail OTP + password se restore.
- **Poori Google Sheet delete ho gayi:** nayi khali sheet banao (Anyone with link - Editor) -> Home -> "Google Sheet deleted?" -> Admin ID, password, nayi sheet ka link, same Google account se verify.
- **REMARK:** Data Sheet ke aakhir me naya REMARK column - JE likh sakta hai ki transformer replace kyu nahi hua. Purani sheets me column apne aap add ho jata hai.
- **Speed:** summary server par cache hoti hai (data badalte hi refresh) aur login ke saath hi aa jati hai.

## Comparison - day wise
- **Unit head (Admin / workshop user):** har din ka damage (PR) - pichhla saal vs is saal, workshop wise, aakhir me TOTAL (inc/dec count + %).
- **DISCOM head:** har division ka pichhla saal | is saal | inc/dec (count) | %, din ke hisaab se, aur total.

## Kitty tab (daily) & day-wise layout
- **Kitty** tab (next to Comparison) is date-wise. A workshop enters that day's Issue, Repair, Issue to other workshop and Receive from other workshop (and the Previous year kitty). Balance kitty = Opening + Repair - Issue - Issue to other workshop + Receive from other workshop; the balance of a day becomes the opening kitty of the next day automatically. Kitty difference = Balance - Previous year kitty. Capacities: 10, 16, 25, 63, 100, 160, 250, 400, 630, 800, 1000.
- **Edit opening kitty** (workshop) -> **Send approval**. The Division Head (Admin) gets a small popup "<workshop> requested to change opening kitty" with the old and new values and taps Approve / Reject. Only after approval does the opening kitty change.
- The Division Head sees one table per item (workshops as columns + TOTAL) and a **Comparison** button next to Balance kitty (Balance | Last year | Difference) with Print and a close button.
- Sheet tabs created automatically: `KittyDay`, `KittyOpen`, `KittyReq` (all included in the daily backup).
- A JE/AE who is transferred keeps the same User ID: the new Division Head can add the same ID in his workshop (no "already exists" error for another organisation). At login the password decides which organisation opens.
- Day-wise comparison: left column is this year's date; each workshop (or division for the DISCOM head) shows Previous year | This year | Diff; Inc/Dec % appears only in the final total group. On phones the Division / Substation / Day wise buttons now fit on one row.

### Kitty update
- **Previous year kitty is automatic**: it is the balance kitty of the same date last year; if that date was not filled, the nearest earlier date (up to 5 days back) is used. While the app has no data of last year, the Division Head fills it with **Last year kitty** (date-wise, all workshops in one table, saved to the `KittyPrev` tab).
- The kitty tables save automatically (about one second after typing stops).
- The entry tables work like Excel: arrow keys and Enter move between cells, drag or Shift+click selects a block, Ctrl+C copies, Ctrl+V pastes a block from Excel (a single value fills a whole selection), Delete clears the selection, Ctrl+A selects all.
- Print layouts of Summary, Comparison and Kitty were made compact (landscape, tiles and tables sized to the page).

## Admin tab: users with several workshops, workshop edit/delete, storage and new sheet
- **Additional charge**: in Admin > Users every non-admin user has an *Additional charge* button (next to Reset password / Deactivate). Tick two or more workshops: the user sees the Summary, Comparison, Kitty and Data Sheet of all of them and can edit and paste data for each (rows of any other workshop are refused with a warning). The *Create a new user* form also takes several workshops.
- **Workshops** can be renamed (every record, driver, kitty entry and user is updated) or deleted (only when the workshop has no records).
- The *Danger zone - delete all data* box was removed. The Admin tab is split into Sheet & storage / Workshops / Users / Backup / Maintenance.
- **Storage**: when the Google Sheet is 90% full (the Google Sheets limit is 20 million cells per file) the Admin sees an English notification. Admin > Sheet & storage: paste the link of a new empty Google Sheet (Anyone with the link - Editor), tap Check, tick the permission box, enter the password and tap *Move data now*. The last 1 year of data (plus every pending record), users, drivers, kitty and settings are copied to the new sheet, which becomes the live one. The old sheet stays as it is; Summary, Comparison and the Data Sheet's *Old data* button read finished records older than the move date from it on demand. Everybody logs in again after the move.

## DISCOM head: tabs chosen by the app Creator

In the Creator panel every DISCOM head has a **Tabs allowed** row (Summary, Comparison, Kitty, Data Sheet, Driver Master). The **Admin** tab can never be given. Default for a new head: Summary + Comparison. Changes are enforced by the server at once.

- **Summary** - every division of the DISCOM one under the other (workshops inside each division): PR today, issued today, PR till date, issued till date, pending replacement, capacity wise tables. Return / pending-for-return are not shown.
- **Comparison** - day wise damage (PR) of every division, this year against the same dates last year.
- **Kitty, Data Sheet, Driver Master** - read only. A *Division* box on top lets the head pick one division at a time; everything is the same as the Division Head's own screen, but nothing can be edited.

Storage meter and the 90% alert use 20 million cells per Google Sheet (no record estimate is shown).

## Wrong / duplicate PR

Every Driver Master gets a default firm **DUPLICATE/WRONG PR** for each workshop (added automatically the first time the Driver Master / Data is opened; it cannot be edited or deleted). Choose it as *Issued to firm* for a wrong or duplicate PR:

- the Issue, Replacement, Time and Return dates are cleared and cannot be kept;
- the record is not counted in PR, issue, pending or comparison figures; it is counted only in its own column **Wrong / dup. PR** (workshop cards, tables and the top tiles - right after *Total return*; also in the DISCOM head's Summary);
- tap the number to open the list - the **Remark** of every record is visible to everyone who can open the list.

Kitty: the figures are typed in a flat grid like the Data Sheet; on phones the keyboard's Next / Enter key moves down the column.

## Home-page search scope

The fast browser index holds only the **last 31 days** of PRs plus **every PR that is not replaced yet** (wrong/duplicate PRs excluded). If nothing is found there the app automatically searches the older records of the DATA tab on the server (a moment slower, first 60 results). ARCHIVE and old Google Sheets are not searched.

## Pasting rows of another workshop

If a workshop user pastes rows that belong to another workshop of the same Admin, a **new** PR is saved under that workshop (it appears in that workshop's Data Sheet and Summary). If the PR already exists there (DATA or ARCHIVE) the row is **skipped** and never changed. Existing records of other workshops can still be edited only by their own workshop or the Admin.

## Share button (home screen)

The small **Share** button on the home screen opens three choices: Android user (shares the APK link), Windows / Mac user and iPhone user (share the website link). Upload your APK to the repository root, next to `index.html`, and write its exact file name in `config.js` (`TT_APK_URL`, default `TF-Tracker.apk`). On phones the normal share sheet opens; on a computer the link is copied.

## Pending for issue - date wise report (Summary tab)

Replaces the old workshop table. For every workshop: PR date (rows) x capacity 10 / 16 / 25 / 63 / 100 / 160 / 250 / 400 / 630 (columns), counting PRs whose transformer is **not issued yet** (till the "To" date; wrong / duplicate PRs are not counted). A workshop sees its own table; the Admin sees one table per workshop plus an ALL WORKSHOPS table. Tap a number to open that list and fill firm / driver / date.

## Wrong / duplicate PR - what changes in the screens

Choosing the firm **DUPLICATE/WRONG PR** switches the date fields off at once (pending cards, edit form and Data Sheet cells are locked and cleared) and the driver list is replaced by two boxes, **JE name (D)** and **JE mobile**, that the user types. The Driver Master default entry is Firm *DUPLICATE/WRONG PR*, driver *JE(D)*, no mobile.

## Big paste

A paste of thousands of rows is saved in parts of 1000 rows with a progress line ("Saving... 5000 of 28000 rows").

## Background saving

Switching tabs never asks for confirmation: pasted rows keep saving in the background. The Data Sheet / Driver Master tab button shows *(saving...)* or *(N pending)*. If the server is busy or the network drops, the save is retried automatically (2, 4, 8, 15 seconds); if it still fails the tab shows the error with a **Retry now** button and the app tries again after 1 minute. The browser warns only when the whole page is being closed while rows are unsaved.

## Financial year (April-March) archive + faster saving
- The **DATA** tab is meant to hold only the **current financial year** (1 April - 31 March) plus every record that is not finished yet. Finished records of earlier financial years (wrong PR, or RETURN DATE entered) go to **ARCHIVE** (Admin > Maintenance > *Archive*; run it once after 1 April each year, after taking a backup). Records that are replaced but never returned stay in DATA (they count in "Pending for return") and move only after 2 years.
- Summary / Excel download / Comparison read the ARCHIVE automatically when a date range goes before 1 April. DISCOM head Summary / Comparison do the same.
- Public complaint status search covers the DATA tab only: finished complaints of earlier financial years are no longer found after archiving.
- To go back to the old rule (2 years) set `var ARCH_FY = false;` in Code.gs.
- Saving a cell change in the Data Sheet (driver, firm, dates, remark ...) now patches the cached index instead of discarding it, so the next "Loading data..." no longer re-reads the whole sheet. Adding / deleting rows or changing workshop, division, substation, capacity or PR NO still rebuilds it once.
- Summary: "Division wise and capacity wise PR" was removed; "Pending for issue - date wise" is a tap-to-open block; workshop cards are shown one below the other. Comparison: "Substation wise" was removed.


## Monthly hot/cold data (Convert) + Kitty 1-day-1-cell
- Admin > Maintenance > **Convert to monthly mode** (once): DATA keeps only the current + previous month and all unfinished records;
  older finished records go to ARCHIVE (1 record = 1 cell); tab **STATS** keeps month-wise PR counts so last year's comparison needs no big read.
  A hidden backup tab `DATA_backup_YYYYMMDD` is created (delete it later). At the first login of every new month the previous month is moved automatically.
- Kitty: new tab **KittyDays** (1 day = 1 cell). The old `KittyDay` tab is copied once on first use and then left untouched.
- Limits: complaint search and the Data sheet show only the hot DATA; an archived PR returns to DATA if its row is pasted/saved again.


## Test: an Admin installs his own server copy (install-test.html)
- Creator (once): Google Cloud Console > project of the OAuth client in config.js > enable **Apps Script API**; OAuth consent screen > Data access > add the scopes
  `.../auth/script.projects` and `.../auth/script.deployments`; add the test Admin's Google e-mail as a *Test user*.
- Admin: turn ON "Google Apps Script API" at script.google.com/home/usersettings, open `install-test.html` on the website, sign in, then in the script editor run `authorizeAll` once.
- The page prints the Admin's own /exec link and an app link `index.html?api=<link>` (asks for a confirmation, nothing is saved).
- The script is read from `apps-script/Code.gs` of the website, so keep the newest Code.gs there.

## Admin's own server ("My server")
Admin -> Users/Admin screen -> **My server**: installs a copy of Code.gs into the Admin's own Google account (same Google Sheet, nothing is moved).
- Needs: `apps-script/Code.gs` (newest) on the website; Admin turns ON "Google Apps Script API" at script.google.com/home/usersettings; Creator enables the Apps Script API + the two scopes in the Cloud project.
- Central script keeps a `Servers` tab in the Registry sheet (Sheet ID | URL | Key | Status). While Active, login answers `{route:[...]}` and the app switches to the Admin's server; central refuses all other calls for that sheet.
- New User IDs created on an own server are sent to central (`regSync`, with the key); if central is unreachable they wait in a queue and are sent later.
- "Switch back to the shared server" is in the same screen.
- Limits: DISCOM / Creator views read the sheet directly (cache 2 min for such sheets); "move to new sheet" and "forgot password by e-mail" are not available on an own server yet.
