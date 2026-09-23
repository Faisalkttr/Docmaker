"""
Streamlit app: upload images + text, generate a polished .docx.

Run locally:   streamlit run app.py
Deploy:        push to GitHub -> connect repo at share.streamlit.io
"""

import io
from PIL import Image, ImageOps
import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn

st.set_page_config(page_title="Doc Builder", page_icon="📄")
st.title("📄 Document Builder")
st.caption("Fill in the details, upload images, download a Word doc.")

# ---------- Inputs ----------
title = st.text_input("Document title", "My Report")
intro = st.text_area("Intro / summary paragraph", "")

uploaded_files = st.file_uploader(
    "Upload images (in the order you want them to appear)",
    type=["png", "jpg", "jpeg"],
    accept_multiple_files=True,
)

captions = []
if uploaded_files:
    st.subheader("Captions (optional)")
    for f in uploaded_files:
        captions.append(st.text_input(f"Caption for {f.name}", key=f.name))

layout = st.radio(
    "Layout",
    ["4 photos per page (2\u00d72 grid)", "1 photo per page (full width)"],
    index=0,
)

# ---- Page / grid geometry (Letter, 1" margins -> 6.5in usable width) ----
PAGE_USABLE_WIDTH_IN = 6.5
CONTENT_HEIGHT_BUDGET_IN = 7.3  # both rows + captions must fit even on the
                                 # heading page; leaves headroom for title/intro
GRID_COLS = 2
GRID_ROWS = 2
CELL_PADDING_IN = 0.15  # room for cell margins so images don't touch borders
CAPTION_ALLOWANCE_IN = 0.3  # per row, for the caption line + spacing
GRID_CELL_MAX_W_IN = PAGE_USABLE_WIDTH_IN / GRID_COLS - CELL_PADDING_IN
GRID_CELL_MAX_H_IN = CONTENT_HEIGHT_BUDGET_IN / GRID_ROWS - CAPTION_ALLOWANCE_IN
SINGLE_MAX_W_IN = PAGE_USABLE_WIDTH_IN
SINGLE_MAX_H_IN = 8.0


def resized_image_stream(uploaded_file, max_width_in, max_height_in, dpi=150):
    """
    Open the uploaded image, auto-rotate per EXIF, downscale if needed,
    and return (BytesIO, width_in, height_in) that fits within
    max_width_in x max_height_in while preserving aspect ratio.
    """
    img = Image.open(uploaded_file)
    # Phone photos store rotation as an EXIF tag rather than rotating the
    # actual pixels. Bake that rotation into the pixels now, BEFORE we
    # read width/height or resize — otherwise portrait photos end up
    # embedded sideways.
    img = ImageOps.exif_transpose(img)
    img = img.convert("RGB")
    w_px, h_px = img.size
    aspect = h_px / w_px  # height / width

    # Fit within the box, preserving aspect ratio (contain, not crop).
    width_in = max_width_in
    height_in = width_in * aspect
    if height_in > max_height_in:
        height_in = max_height_in
        width_in = height_in / aspect

    # Downscale pixels to roughly match the placed size at `dpi`, so a
    # 6000px phone photo doesn't bloat the file for a 3-inch cell.
    max_px_width = max(int(width_in * dpi), 1)
    if w_px > max_px_width:
        new_w = max_px_width
        new_h = int(new_w * aspect)
        img = img.resize((new_w, new_h), Image.LANCZOS)

    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=88)
    buf.seek(0)
    return buf, width_in, height_in


def _prevent_row_split(row):
    """Force a table row to stay together on one page (no splitting a row's
    image from its caption across a page boundary)."""
    tr_pr = row._tr.get_or_add_trPr()
    cant_split = tr_pr.makeelement(qn("w:cantSplit"), {})
    tr_pr.append(cant_split)


def _strip_table_borders(table):
    """Remove visible grid lines so the table reads as pure layout, not a data table."""
    tbl_pr = table._tbl.tblPr
    borders = tbl_pr.makeelement(qn("w:tblBorders"), {})
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = borders.makeelement(qn(f"w:{edge}"), {qn("w:val"): "none"})
        borders.append(el)
    tbl_pr.append(borders)


def _add_image_to_cell(cell, img_stream, width_in, height_in, caption):
    paragraph = cell.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run()
    run.add_picture(img_stream, width=Inches(width_in), height=Inches(height_in))
    if caption:
        cap = cell.add_paragraph(caption)
        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cap.runs[0].italic = True
        cap.runs[0].font.size = Pt(9)


def build_docx(title, intro, files, captions, grid_layout=True):
    doc = Document()

    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)

    heading = doc.add_heading(title, level=1)
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER

    if intro.strip():
        doc.add_paragraph(intro.strip())

    items = list(zip(files, captions))

    if not grid_layout:
        # ---- One photo per page, full width ----
        for f, caption in items:
            f.seek(0)
            img_stream, w, h = resized_image_stream(f, SINGLE_MAX_W_IN, SINGLE_MAX_H_IN)
            doc.add_picture(img_stream, width=Inches(w), height=Inches(h))
            doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
            if caption:
                cap = doc.add_paragraph(caption)
                cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
                cap.runs[0].italic = True
                cap.runs[0].font.size = Pt(9)
            doc.add_paragraph()
        out = io.BytesIO()
        doc.save(out)
        out.seek(0)
        return out

    # ---- 4 photos per page, 2x2 grid ----
    per_page = GRID_COLS * GRID_ROWS
    groups = [items[i:i + per_page] for i in range(0, len(items), per_page)]

    for group_index, group in enumerate(groups):
        table = doc.add_table(rows=GRID_ROWS, cols=GRID_COLS)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = True
        _strip_table_borders(table)
        for row in table.rows:
            _prevent_row_split(row)

        for i, (f, caption) in enumerate(group):
            row, col = divmod(i, GRID_COLS)
            cell = table.cell(row, col)
            f.seek(0)
            img_stream, w, h = resized_image_stream(f, GRID_CELL_MAX_W_IN, GRID_CELL_MAX_H_IN)
            _add_image_to_cell(cell, img_stream, w, h, caption)

        # Page break between groups (not after the last one).
        if group_index < len(groups) - 1:
            doc.add_page_break()

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out


if st.button("Generate document", type="primary"):
    if not uploaded_files and not intro.strip():
        st.warning("Add some text or at least one image first.")
    else:
        grid_layout = layout.startswith("4 photos")
        docx_bytes = build_docx(title, intro, uploaded_files or [], captions, grid_layout=grid_layout)
        st.success("Document generated.")
        st.download_button(
            "⬇️ Download .docx",
            data=docx_bytes,
            file_name=f"{title or 'document'}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
