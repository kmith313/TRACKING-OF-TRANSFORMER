# Transformer Tracking + Complaint / PR Status (GitHub Pages)

The frontend runs on GitHub Pages; the backend is a Google Apps Script that you deploy once.
Each organisation gives its own Google Sheet link to the app, and the app builds everything else.

## Folder
- `index.html`, `complain.js`, `config.js`, `sw.js`, `manifest.webmanifest`, `icons/` : files to upload to GitHub
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

## Admin flow (when someone installs the app)
1. On first open, press **I am the Admin**.
2. **Sign in with Google** (the Gmail account that owns the Google Sheet) and allow access. The Google password is never seen by the app.
3. Paste the link of the Google Sheet, enter the names of all workshops, and choose the Admin ID + password. Press **Create**.
4. The app automatically shares the sheet (Anyone with the link -> Editor), creates the tabs `DATA`, `Users`, `Drivers` (Driver Master), `Config`, `ARCHIVE`
   (data columns go only up to **REPLACEMENT DATE + TIME**), logs the Admin in, and discards the Google access.
5. In the app's **Users** tab the Admin copies the "Link for users" and sends it to the users.

## Who has installed it? (Admin registry)
Every time someone sets up the app, a row is added to a Google Sheet in YOUR Drive named **Transformer Tracking - Admin Registry**:
time, the Admin's verified Google email, Admin ID, organisation, workshops and Sheet link. The sheet is created automatically on the first setup.
To find it: in Apps Script, select the function `registryUrl` and press Run; the link appears in the Execution log.
For anyone with a Gmail to be able to become an Admin, publish the OAuth app (Google Auth Platform -> Audience -> Publish app); in Testing mode only listed Test users can sign in.

## User flow
The user opens the link (or presses **I am a User** and pastes the Sheet link), lands directly on the **login page**, and logs in with ID + password.

## Complaint / PR status
It is on the login page itself and works **without login**. The data comes from the `DATA` tab of the Google Sheet the Admin connected
(fresh data on every search, auto-refresh every 5 minutes, offline cache).

## Please note
- The Sheet is shared as "Anyone with the link", so anyone who has the link can open the Sheet directly (the Users tab also holds password hashes). Share the link only with your own people and use strong passwords.
- Free Google accounts have daily Apps Script limits; for many organisations a Workspace account is better.
