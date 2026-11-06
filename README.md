# Valley View Sunday School: installable app

Two parts:
- `apps-script/Code.gs` runs inside your Google Sheet (the backend: schedule, calendar, emails, sign-in).
- `github-pages/` is the app teachers install on their phones (free hosting on GitHub Pages).

Your backend address is already filled in, in both places:
`https://script.google.com/macros/s/AKfycbzBaM05VE2DM7TF2mZVSMM29DfTGQJkWGK84z_baeGSUisNamgZP2E2SZGKVoo3t_XS0w/exec`

## Step 1: Update the Google Apps Script (keep the SAME web address)
1. Sheet ▸ **Extensions ▸ Apps Script**.
2. Replace everything in `Code.gs` with `apps-script/Code.gs`.
3. **Delete `Index.html` from the Apps Script project** (the app no longer lives there).
4. Click **Deploy ▸ Manage deployments ▸ pencil (edit) ▸ Version: New version ▸ Deploy**.
   Do not click "New deployment", or the address changes. Settings stay: Execute as **Me**, access **Anyone**
   (the sign-in code protects the data, not the link).
5. Approve any permission prompts the first time (mail, calendar, drive, external requests).
6. Back in the sheet, reload and run **Sunday School Tools ▸ Check Web App Setup** to confirm.

## Step 2: Put the app on GitHub Pages
1. Create a free GitHub account, then **New repository** (for example `valley-view-sunday-school`). Public is fine.
2. Upload every file inside `github-pages/` to the top level of the repo: `index.html`, `manifest.json`, `sw.js`,
   the five `.png` icons, and `.nojekyll`.
3. Repo **Settings ▸ Pages ▸ Build and deployment**: Source = *Deploy from a branch*, Branch = `main` / `(root)` ▸ Save.
4. After a minute your app is live at `https://YOUR-NAME.github.io/valley-view-sunday-school/`.
5. Copy that address into `CONFIG.GITHUB_PAGES_APP_URL` in `Code.gs`, save, and deploy a **New version** again.
   Email "Open App" buttons will then open the app.

## Step 3: Teachers install it
1. Send teachers the app address.
2. They open it, enter the email address that is on the teaching schedule, and type the 6-digit code that is emailed to them.
3. **iPhone:** Safari ▸ Share ▸ **Add to Home Screen**. **Android:** Chrome ▸ ⋮ ▸ **Install app** (or the Install button on the sign-in screen).
They stay signed in for 90 days.

## Who can get in
- Only people whose email is in column G of the **Schedule** sheet, or who hold a Sunday School president / counselor / secretary
  calling on the **Contact** sheet. Everyone else just sees "a code is on its way" and never gets one.
- Add a new teacher: put their email on the schedule. (Allow up to 2 minutes.) Remove one: delete the email and they are locked out within 2 minutes.
- Lost phone or want to reset everyone: **Sunday School Tools ▸ Sign Out All Devices**.
- Teachers can RSVP / request a sub only for their own assignments. The presidency can act on any assignment.

## Attendance
This app never reads or writes attendance data, and the **admin** attendance address (the one with `key=...`) is not in this project
at all. Signed-in people are given only the three public check-in links, and only these:
- `.../exec?class=Sunday%20School`
- `.../exec?class=Gospel%20Foundation`
- `.../exec?class=Temple%20Prep`

Safeguards: the server strips any `key` or other query text from the attendance address before sending it, the app refuses any class
name other than those three, and nothing is sent until someone has signed in. The public check-in address is not in the GitHub files.

## Notes
- Anyone who has a class check-in link or QR code can open that class's check-in page, as before. The admin page is never part of this project.
- Teaching aids: the sheet's **Teaching Aids** tab must have the "date / week / lesson" style header in cell A1 for new submissions
  to land in the right columns (same as the earlier package).
- Updating the app later: upload the new `index.html` to GitHub. The installed app picks it up the next time it is opened online.
