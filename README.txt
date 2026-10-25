AUDITFLOW - GITHUB PAGES (IN-BROWSER) VERSION
=============================================

What this is
------------
A static website. Every tool (Lead Schedule, Audit Sampling, Fixed Assets)
runs inside the visitor's browser using Python (Pyodide). The client's
Excel file is never uploaded anywhere. There is no server to run or pay for.

Put it online at your GitHub address (about 5 minutes)
------------------------------------------------------
1. Sign in at github.com as camayankchopra.
2. Click New repository. Name it:  auditflow   (Public). Create it.
3. Click "uploading an existing file".
4. Unzip auditflow-site.zip on your PC. Open the unzipped folder and drag
   EVERYTHING INSIDE it (index.html, assets, py, pyodide, wheels, templates,
   manifest.json, .nojekyll ...) into the GitHub upload box.
   - The folder structure must be kept. If the browser upload misses the
     sub-folders, use GitHub Desktop (drag the folder in, Commit, Push).
   - The largest file (pyodide.asm.wasm) is about 9 MB; GitHub's limit is 25 MB
     per file for browser upload, so this is fine.
5. Click Commit changes.
6. In the repository: Settings > Pages > Build and deployment >
   Source: "Deploy from a branch" > Branch: main, folder: / (root) > Save.
7. Wait 1-2 minutes. Your site is live at:

       https://camayankchopra.github.io/auditflow/

First visit downloads about 14 MB (the Python engine) and then the browser
caches it; later visits start in a few seconds.

Set your own name and contact e-mail
------------------------------------
Open assets/config.js in the site folder and edit:

    window.AF = { local:true, brand:"AuditFlow",
                  feedbackEmail:"you@yourfirm.com", feedbackUrl:"" };

- feedbackEmail : the "Request a demo" and Feedback buttons open an e-mail to this.
- feedbackUrl   : optional. Paste a Google Form link to collect feedback in a
                  form instead of e-mail.
Or set CONTACT_EMAIL / BRAND_NAME / FEEDBACK_URL before running
python build_static.py to regenerate the site.

Adding the next tool later
--------------------------
Add the tool in the main project (bridge.py + app.py + template), run
python build_static.py, and upload the new auditflow-site contents over the
old ones (GitHub keeps history, so you can roll back).

Good to know
------------
- No login or access-code gate: it is a public website. Anyone with the link
  can use the tools. That is fine for a trial; for paying clients you will
  eventually want accounts (needs a server).
- No feedback is stored on a server; feedback goes by e-mail / your form.
- Works in current Chrome, Edge, Firefox and Safari. Needs a desktop or laptop.
