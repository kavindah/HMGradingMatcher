GRADE VERIFICATION  -  free web app
====================================

Upload an APPROVED REFERENCE spec PDF and the SPEC to check (same grade-rule
category). Download a colour-coded Excel: green where the spec's size-to-size
grade increments match the reference, red where they differ, plus an
"Exceptions" sheet listing every mismatch.

FILES
  streamlit_app.py        the web page (two uploads -> verify -> download)
  grade_verify.py         the comparison engine
  requirements.txt        Python packages the host installs
  .streamlit/config.toml  hides the menu/fork, sets upload size (KEEP the dot-folder)

--------------------------------------------------------------------
TRY IT LOCALLY (optional)
--------------------------------------------------------------------
  pip install -r requirements.txt
  streamlit run streamlit_app.py

--------------------------------------------------------------------
PUT IT ONLINE FREE (Streamlit Community Cloud)
--------------------------------------------------------------------
1. Free GitHub account: https://github.com
2. New repository, e.g. "grade-verifier".
3. Upload these files INCLUDING the .streamlit folder and its config.toml:
       streamlit_app.py
       grade_verify.py
       requirements.txt
       .streamlit/config.toml
   (Dot-folders are easy to miss when drag-uploading - make sure it lands.)
4. https://share.streamlit.io -> sign in with GitHub -> Create app.
       Repository:      your-username/grade-verifier
       Branch:          main
       Main file path:  streamlit_app.py
5. Deploy. You get a public URL like https://grade-verifier.streamlit.app
6. Updating: edit files on GitHub; it redeploys. Config changes need a
   Reboot (app menu -> Reboot app).

Alternative free host: Hugging Face Spaces (choose the Streamlit SDK), upload
the same files.

--------------------------------------------------------------------
PRIVACY / CONFIDENTIALITY
--------------------------------------------------------------------
- Both PDFs are processed IN MEMORY and never written to the server's disk;
  the bytes are released after the Excel is built, and "Clear files" wipes them.
- A public app still means uploads travel to the host's servers. For
  confidential tech packs, keep the repo private with Streamlit's viewer
  allow-list (Settings -> Sharing), or run it on an internal company machine.
