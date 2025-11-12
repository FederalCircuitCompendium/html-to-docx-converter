import streamlit as st
from io import BytesIO
from pathlib import Path
import html as py_html
import tempfile
from typing import Optional


from docx import Document

APP_DIR = Path(__file__).parent
REF_DOCX = APP_DIR / "assets" / "reference.docx"

st.set_page_config(page_title="HTML → DOCX", page_icon="📄", layout="wide")
st.title("HTML → DOCX (templated)")

# --- Inputs ---
title = st.text_input("Document title", value="", placeholder="Enter title…")
html_body = st.text_area("HTML", value="", height=420, placeholder="Paste raw HTML here…")
start_level = st.number_input("Demote body headings to start at level", min_value=1, max_value=9, value=1, step=1)
strong_emph = st.checkbox("Map bold/italic to Word Strong/Emphasis", value=True)

col1, col2 = st.columns([1,1])
with col1:
    st.caption("• Title is injected as `<h1>` → Word Heading 1\n"
               "• Metadata: Title + language = en-US\n"
               f"• Template: {'found ✅' if REF_DOCX.exists() else 'not found (ok)'}")

def apply_language_en_us(doc: Document):
    ns = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    def tag_runs(container):
        for p in container.paragraphs:
            for r in p.runs:
                if r._element.rPr is None:
                    r._element._add_rPr()
                r._element.rPr.set(ns + "lang", "en-US")
    tag_runs(doc)
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                tag_runs(cell)

def remap_headings_except_first(doc: Document, start_level: int):
    delta = max(0, start_level - 1)
    if delta == 0: return
    for i, p in enumerate(doc.paragraphs):
        if i == 0:  # keep the title line
            continue
        s = p.style
        if not s: continue
        nm = getattr(s, "name", str(s))
        if nm.startswith("Heading "):
            try:
                lvl = int(nm.split()[-1])
            except Exception:
                continue
            p.style = f"Heading {min(9, lvl + delta)}"

def bold_italic_to_character_styles(doc: Document, use=True):
    if not use: return
    for p in doc.paragraphs:
        for r in p.runs:
            if r.bold:
                try: r.style = "Strong"
                except: pass
            if r.italic:
                try: r.style = "Emphasis"
                except: pass

def try_pandoc_convert(html_str: str, title: str, out_path: Path, reference: Optional[Path]):
    try:
        import pypandoc
        try:
            pypandoc.get_pandoc_version()
        except OSError:
            pypandoc.download_pandoc()
        extra_args = ["--metadata=lang=en-US", f"--metadata=title={title}"]
        if reference and reference.exists():
            extra_args.append(f"--reference-doc={reference}")
        pypandoc.convert_text(html_str, to="docx", format="html",
                              outputfile=str(out_path), extra_args=extra_args)
        return True, ""
    except Exception as e:
        return False, str(e)

def fallback_htmldocx(html_body: str, title: str, out_path: Path):
    from htmldocx import HtmlToDocx
    doc = Document()
    p = doc.add_paragraph(title)
    try: p.style = "Heading 1"
    except: pass
    HtmlToDocx().add_html_to_document(html_body, doc)
    doc.core_properties.title = title
    doc.core_properties.language = "en-US"
    apply_language_en_us(doc)
    doc.save(out_path)

def build_docx(title: str, body_html: str, start_level: int, strong_emph: bool) -> bytes:
    if not title: title = "Converted Document"
    combined_html = f"<h1>{py_html.escape(title)}</h1>\n{body_html or '<p>(empty)</p>'}"
    with tempfile.TemporaryDirectory() as td:
        out_path = Path(td) / "out.docx"
        ok, err = try_pandoc_convert(combined_html, title, out_path, REF_DOCX if REF_DOCX.exists() else None)
        if not ok:
            # Pandoc failed; fall back
            fallback_htmldocx(body_html, title, out_path)
            doc = Document(str(out_path))
        else:
            doc = Document(str(out_path))

        if start_level > 1:
            remap_headings_except_first(doc, start_level)
        if strong_emph:
            bold_italic_to_character_styles(doc, True)
        doc.core_properties.language = "en-US"
        apply_language_en_us(doc)
        bio = BytesIO()
        doc.save(bio)
        return bio.getvalue()

if st.button("Convert"):
    if not title and not html_body:
        st.warning("Please provide a title or some HTML.")
    else:
        try:
            data = build_docx(title, html_body, int(start_level), bool(strong_emph))
            filename = (title or "Converted Document").strip().replace("/", "-") + ".docx"
            st.download_button("Download .docx", data=data, file_name=filename, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            st.success("Document ready.")
        except Exception as e:
            st.error(f"Conversion failed: {e}")
