White Line Trucking Management Portal
Permanent Report Storage Version

WHAT THIS VERSION FIXES
- Reports stay visible after switching pages.
- Reports stay visible after closing and reopening the browser.
- Reports can stay visible after Streamlit Cloud goes to sleep/reboots, as long as GitHub storage is configured.

IMPORTANT
Streamlit Cloud can reset local files when the app sleeps, reboots, or redeploys. This version can save the latest uploaded reports back into your GitHub repository so they remain until a new report is uploaded.

REQUIRED FILES FOR GITHUB / STREAMLIT
Upload these files to the GitHub repository:
- app.py
- requirements.txt
- white_line_trucking_logo.png

STREAMLIT CLOUD SECRETS REQUIRED FOR PERMANENT STORAGE
In Streamlit Cloud, open the app settings and add these secrets:

GITHUB_TOKEN = "paste_your_github_token_here"
GITHUB_REPO = "your-github-username/white-line-trucking-management-portal"
GITHUB_BRANCH = "main"
GITHUB_STORAGE_PATH = "portal_storage"

HOW TO CREATE THE GITHUB TOKEN
1. Go to GitHub.
2. Click your profile picture.
3. Go to Settings.
4. Go to Developer settings.
5. Go to Personal access tokens.
6. Choose Fine-grained tokens.
7. Generate a new token.
8. Repository access: choose only the white-line-trucking-management-portal repository.
9. Repository permissions: set Contents to Read and write.
10. Generate the token and copy it.
11. Paste it into Streamlit Secrets as GITHUB_TOKEN.

WHAT GETS SAVED
The app saves the processed report data for:
- Cash Flow
- A/R weekly income manual entry
- A/P
- P&L CM
- P&L YTD

WHERE IT GETS SAVED
A folder will appear in GitHub called:
portal_storage

Inside that folder, the latest processed reports are saved as small .pkl files. Do not delete these files unless you want to clear the saved portal reports.

LOCAL TESTING
Double-click START_LOCAL.bat to run locally.
Local-only storage works on your computer, but true permanent Streamlit Cloud storage requires the GitHub secrets above.

NETWORK TESTING
Double-click START_LAN_SHARE.bat to allow computers on the same office network to access the portal.

CLOUD DEPLOYMENT
1. Upload app.py, requirements.txt, and white_line_trucking_logo.png to GitHub.
2. Configure the Streamlit secrets above.
3. Redeploy or reboot the Streamlit app.
4. Upload each report once.
5. Close/reopen the website to confirm the reports remain.


Theme: White and blue branding for White Line Trucking LLC.

A/R NOTE
The A/R page is manual entry only. Enter the weekly income amount and click Save weekly income. The saved number stays visible until it is updated or replaced.
