# Time Register (GitHub Pages + Google Sheets)

The page (`index.html`, `config.js`) is static and hosted on GitHub Pages. All data lives in a Google Sheet that only the Apps Script (`Code.gs`) can read. People sign in with a username and a password kept in the sheet.

```
Browser (GitHub Pages)  ->  Apps Script web app (API)  ->  Google Sheet
```

## 1. Back end (Google Sheet + Apps Script)

1. Create a new Google Sheet.
2. **Extensions > Apps Script**. Replace the contents of `Code.gs` with the provided `Code.gs`. There is **no** Index file in this version.
3. Select `setup` in the function dropdown and click **Run**. Approve the permissions. This creates the Attendance Logs, Employees and Audit Log sheets, plus a first administrator row (username `admin`).
4. Open the **Employees** sheet and type a password in column **F (Password)** of that row. You can also change the username in column C (use lowercase letters, no spaces).
5. **Deploy > New deployment > Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**
6. Click **Deploy** and copy the web app URL (it ends in `/exec`). Opening it in a browser should show `{"ok":true,...}`.

If you used the earlier version with hashed PINs: running `setup` again turns column F into a plain **Password** column. Type each person's password into column F again.

## 2. Front end (GitHub)

1. Open `config.js` and paste your web app URL between the quotes.
2. Create a new GitHub repository (Public or Private) and upload `index.html`, `config.js`, `Code.gs` and this README.
3. In the repository, go to **Settings > Pages**. Under **Build and deployment**, choose **Deploy from a branch**, branch `main`, folder `/ (root)`, and save.
4. After a minute your page is at `https://<your-username>.github.io/<repository-name>/`.

Private repositories can only use GitHub Pages on paid GitHub plans. A Public repository is fine, because it contains no data or secrets.

## 3. Add people

You can do it in two ways:

- **In the sheet:** add a row in **Employees** with ID, name, username (lowercase), role (`Employee` or `Administrator`), Active (`Yes`) and the password in column F. Changing a password later is just editing that cell.
- **In the app:** sign in as admin and use the **Employees** tab to add people, reset passwords, and deactivate accounts. People can change their own password with **Change password** at the top right.

## Updating later

When you change `Code.gs`, use **Deploy > Manage deployments > pencil icon > Version: New version > Deploy**. The URL stays the same.

## Security notes

- The web app URL is public, so anyone can try to sign in. Every action except login needs a signed session token that expires after 7 days. A deactivated user is locked out immediately.
- Passwords are stored as plain text in the Employees sheet, so keep that Google Sheet private (owner only). After 5 wrong attempts a username is locked for 15 minutes.
- This is sensible protection for a small team's timesheets. It is not designed for highly sensitive data. Use passwords of 6 or more characters.
- Employees can only see and edit their own entries. Administrators can see and edit everyone's. Every edit and delete is recorded in the Audit Log sheet.
- The Google Sheet is never exposed. Do not share it with employees.

## If sign-in says "Wrong username or password"

1. In **Employees**, the username in column C must match what you type (lowercase, no spaces), Active in column E must be `Yes`, and the password in column F must not be empty.
2. Check for a space before or after the password. The script trims spaces, but a visible typo is the most common cause.
3. After 5 wrong tries the username is locked for 15 minutes. Wait, then try again.
4. After editing `Code.gs`, redeploy with **Deploy > Manage deployments > pencil icon > Version: New version > Deploy**, or the live API still runs the old code.
