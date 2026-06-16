White Line Trucking Management Portal - Permanent Storage Version
================================================================

This version is set up to keep the latest uploaded reports active until a new file/report is uploaded.

Pages included:
1. Home
2. Cash Flow
3. A/R - manual weekly income entry
4. A/P
5. P&L CM
6. P&L YTD

WHAT STAYS SAVED
----------------
- Cash Flow uploaded workbook values
- A/R manual weekly income number
- A/P uploaded report output
- P&L CM uploaded report view
- P&L YTD uploaded report view

HOW IT WORKS
------------
On local computer testing, reports save in a local saved_reports folder.
On Streamlit Cloud, true permanent storage requires GitHub storage secrets.
Once the secrets are configured, the user only uploads the report in the portal.
The portal automatically saves the latest report to the GitHub repository in a portal_storage folder.
No one needs to manually commit reports after each upload.

REQUIRED FILES FOR GITHUB / STREAMLIT
-------------------------------------
Upload all files from this folder to your GitHub repository:
- app.py
- requirements.txt
- white_line_trucking_logo.png
- STREAMLIT_SECRETS_TEMPLATE.txt
- README_SETUP.txt

STREAMLIT SECRETS REQUIRED FOR PERMANENT STORAGE
------------------------------------------------
In Streamlit Cloud, open the app, go to Settings > Secrets, and add:

GITHUB_TOKEN = "paste_your_github_token_here"
GITHUB_REPO = "your-github-username/white-line-trucking-management-portal"
GITHUB_BRANCH = "main"
GITHUB_STORAGE_PATH = "portal_storage"

The GitHub token must have Contents: Read and Write permission for this specific repository.

TOKEN SETUP SUMMARY
-------------------
1. Go to GitHub.
2. Open Settings.
3. Go to Developer settings.
4. Open Personal access tokens.
5. Choose Fine-grained tokens.
6. Generate a new token.
7. Select only the White Line Trucking portal repository.
8. Give Repository permissions > Contents > Read and write.
9. Copy the token.
10. Paste it into Streamlit Secrets as GITHUB_TOKEN.

IMPORTANT
---------
If the Streamlit app resets after a day, that means Streamlit Cloud restarted the app container.
The GitHub storage secrets are what make the reports come back after that restart.

After secrets are set up, the normal user workflow is simple:
1. Open the portal.
2. Upload or update the report.
3. The portal saves it automatically.
4. It stays visible until a new report is uploaded.
