# Doc Builder (Streamlit)

A small Streamlit app that lets a user upload images and text, then
generates a properly formatted `.docx` — images are auto-resized to fit
the page width while keeping their aspect ratio.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

Then open the local URL it prints (usually http://localhost:8501).

## Put it on GitHub

```bash
git init
git add .
git commit -m "Initial doc builder app"
git branch -M main
git remote add origin https://github.com/<your-username>/<repo-name>.git
git push -u origin main
```

## Deploy for free on Streamlit Community Cloud

1. Go to https://share.streamlit.io and sign in with GitHub.
2. Click **New app**, pick your repo/branch, and set the main file to `app.py`.
3. Click **Deploy**. It builds from `requirements.txt` automatically.
4. Every time you `git push` to `main`, the live app redeploys.

## How the image handling works

- `PIL.Image` opens each upload and reads its real pixel dimensions.
- If the image is larger than needed for print quality (150 DPI at the
  target width), it's downscaled — keeps the .docx file size small.
- Aspect ratio is preserved when computing the placed width/height, and
  height is capped so a very tall image can't run off the page — width
  is then rescaled to match.
- `python-docx`'s `add_picture(..., width=Inches(w))` places the image;
  height is derived automatically from the aspect ratio.

## Extending it

- Multiple images per row: build a table with `doc.add_table(rows=1, cols=2)`
  and insert one image per cell instead of stacking vertically.
- Custom templates: start from an existing branded `.docx` with
  `Document("template.docx")` instead of `Document()`, and append content.
- PDF export too: after saving the docx, convert with LibreOffice
  (`soffice --headless --convert-to pdf`) if you also want a PDF download.
