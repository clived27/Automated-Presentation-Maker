"""
generate-ppt.py
Serverless function: accepts a POST request with JSON body,
downloads a master PPTX template, fills in placeholders, and
returns a downloadable .pptx file.

Expected JSON body:
{
    "template_url": "https://...",
    "date": "Sunday, 29 June 2026",
    "sections": [
        {
            "name": "Entrance",
            "song": {
                "title": "Song Title",
                "lyrics": [
                    { "label": "Verse 1", "text": "Line 1\nLine 2" },
                    { "label": "Chorus",  "text": "Chorus line 1\nChorus line 2" },
                    { "label": "Verse 2", "text": "Line 3\nLine 4" }
                ]
            }
        },
        ... (9 sections total, one per Mass part)
    ]
}

Slide mapping:
  Slide 0  → Cover Page  (replace {{DATE}})
  Slide 1  → Entrance
  Slide 2  → Lord Have Mercy
  Slide 3  → Gloria
  Slide 4  → Acclamation
  Slide 5  → Offertory
  Slide 6  → Holy Holy
  Slide 7  → Proclamation
  Slide 8  → Communion
  Slide 9  → Recessional
"""

import copy
import json
import io
import traceback
from http.server import BaseHTTPRequestHandler

import requests
from pptx import Presentation
from lxml import etree


# ---------------------------------------------------------------------------
# XML namespace constant
# ---------------------------------------------------------------------------
_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


# ---------------------------------------------------------------------------
# Helpers — text manipulation
# ---------------------------------------------------------------------------

def _para_full_text(para) -> str:
    """Concatenate all run texts in a paragraph (handles split runs)."""
    return "".join(run.text for run in para.runs)


def _frame_contains(text_frame, placeholder: str) -> bool:
    """
    Return True if *placeholder* appears anywhere in the text frame.
    Checks the concatenated paragraph text to handle placeholders that
    python-pptx has split across multiple runs.
    """
    for para in text_frame.paragraphs:
        if placeholder in _para_full_text(para):
            return True
    return False


def _replace_in_para(para, old: str, new: str) -> bool:
    """
    Replace *old* with *new* within a single paragraph.
    If the placeholder is split across runs, consolidates all run text
    into the first run (preserving its formatting) and clears the rest.
    Returns True if a replacement was made.
    """
    full = _para_full_text(para)
    if old not in full:
        return False

    replaced = full.replace(old, new)
    runs = para.runs
    if runs:
        runs[0].text = replaced
        for run in runs[1:]:
            run.text = ""
    return True


def _replace_text_in_frame(text_frame, old: str, new: str) -> bool:
    """Replace *old* with *new* across all paragraphs in a text frame."""
    replaced = False
    for para in text_frame.paragraphs:
        if _replace_in_para(para, old, new):
            replaced = True
    return replaced


def _set_frame_text(text_frame, new_text: str):
    """
    Replace the entire content of *text_frame* with *new_text*.
    Each \\n becomes a new paragraph, preserving the run properties
    (font size, bold, colour, etc.) of the first run found in the frame.
    """
    # Snapshot rPr (run properties) from the first non-empty run
    saved_rPr = None
    for para in text_frame.paragraphs:
        for run in para.runs:
            rPr_el = run._r.find(f"{{{_NS}}}rPr")
            if rPr_el is not None:
                saved_rPr = copy.deepcopy(rPr_el)
            break
        if saved_rPr is not None:
            break

    # Work directly on the txBody XML element
    txBody = text_frame._txBody

    # Remove every existing <a:p> element
    for p_el in txBody.findall(f"{{{_NS}}}p"):
        txBody.remove(p_el)

    # Re-build one <a:p> per line
    for line in new_text.split("\n"):
        p_el = etree.SubElement(txBody, f"{{{_NS}}}p")
        r_el = etree.SubElement(p_el, f"{{{_NS}}}r")

        if saved_rPr is not None:
            r_el.insert(0, copy.deepcopy(saved_rPr))

        t_el = etree.SubElement(r_el, f"{{{_NS}}}t")
        t_el.text = line


# ---------------------------------------------------------------------------
# Safe slide duplication
# ---------------------------------------------------------------------------

def _duplicate_slide(prs: Presentation, slide_index: int):
    """
    Safely duplicate the slide at *slide_index* and insert the copy
    immediately after it.

    Strategy:
      1. Add a fresh slide using the *same layout* as the template slide.
         python-pptx registers it properly (relationships, rIds, etc.).
      2. Copy the shapes from the template slide's spTree into the new slide's
         spTree using lxml deep-copy (copies XML nodes, not python-pptx
         wrapper objects — avoids broken internal references).
      3. Reposition the new sldId entry in the presentation's sldIdLst so the
         slide sits immediately after the original.

    Returns the new slide object (already at slide_index + 1).
    """
    template_slide = prs.slides[slide_index]
    layout = template_slide.slide_layout

    # Step 1: Add a properly initialised slide at the end of the deck
    new_slide = prs.slides.add_slide(layout)

    # Step 2: Replace the new slide's shape tree with a copy of the template's
    src_spTree = template_slide.shapes._spTree
    dst_spTree = new_slide.shapes._spTree

    # Remove everything currently in the destination spTree
    for child in list(dst_spTree):
        dst_spTree.remove(child)

    # Deep-copy each child XML node from source to destination.
    # We copy the lxml elements directly (not python-pptx objects), which is
    # safe because lxml elements carry no back-references to the presentation.
    for child in src_spTree:
        dst_spTree.append(copy.deepcopy(child))

    # Step 3: Move the new slide's sldId entry to slide_index + 1
    sldIdLst = prs.slides._sldIdLst
    new_sldId = sldIdLst[-1]          # add_slide always appends at the end
    sldIdLst.remove(new_sldId)
    sldIdLst.insert(slide_index + 1, new_sldId)

    return prs.slides[slide_index + 1]


# ---------------------------------------------------------------------------
# Section slide filling
# ---------------------------------------------------------------------------

def _fill_section_slide(prs: Presentation, slide_index: int, song: dict):
    """
    Fill a Mass-section slide and duplicate it for additional verses.

    Lyrics list format: [{ "label": "Verse 1", "text": "..." }, ...]
    Items whose label contains "chorus" are treated as the shared chorus
    and appended (with a blank line) after every verse on each slide.

    Verse 1  → fills the original template slide
    Verse 2+ → each gets a fresh duplicate of the template slide
    """
    lyrics = song.get("lyrics", [])

    # Clear placeholder and return early if no lyrics provided
    if not lyrics:
        slide = prs.slides[slide_index]
        for shape in slide.shapes:
            if shape.has_text_frame and _frame_contains(shape.text_frame, "{{LYRICS}}"):
                _replace_text_in_frame(shape.text_frame, "{{LYRICS}}", "")
        return

    # Split into verses and chorus
    chorus_text = ""
    verses = []
    for item in lyrics:
        label = item.get("label", "").lower()
        text  = item.get("text", "").strip()
        if "chorus" in label:
            chorus_text = text
        elif text:
            verses.append(text)

    # Edge case: only a chorus was provided
    if not verses and chorus_text:
        verses = [chorus_text]
        chorus_text = ""

    def _build_text(verse_text: str) -> str:
        return f"{verse_text}\n\n{chorus_text}" if chorus_text else verse_text

    # --- Verse 1 → original slide ---
    original_slide = prs.slides[slide_index]
    for shape in original_slide.shapes:
        if shape.has_text_frame and _frame_contains(shape.text_frame, "{{LYRICS}}"):
            _set_frame_text(shape.text_frame, _build_text(verses[0]))
            break

    # --- Verses 2+ → duplicated slides ---
    for extra_idx, verse_text in enumerate(verses[1:], start=1):
        # Duplicate the slide that is now at (slide_index + extra_idx - 1)
        # so the new copy lands at (slide_index + extra_idx)
        _duplicate_slide(prs, slide_index + extra_idx - 1)
        dup_slide = prs.slides[slide_index + extra_idx]

        for shape in dup_slide.shapes:
            if shape.has_text_frame and _frame_contains(shape.text_frame, "{{LYRICS}}"):
                _set_frame_text(shape.text_frame, _build_text(verse_text))
                break


# ---------------------------------------------------------------------------
# Core generation
# ---------------------------------------------------------------------------

SECTION_ORDER = [
    "Entrance",
    "Lord Have Mercy",
    "Gloria",
    "Acclamation",
    "Offertory",
    "Holy Holy",
    "Proclamation",
    "Communion",
    "Recessional",
]


def generate_presentation(template_url: str, date: str, sections: list) -> io.BytesIO:
    """
    Download the PPTX template from *template_url*, fill all placeholders,
    and return the finished presentation as an in-memory BytesIO buffer.
    """
    # 1. Download template entirely into memory — no temp files
    resp = requests.get(template_url, timeout=30)
    resp.raise_for_status()
    template_bytes = io.BytesIO(resp.content)

    # 2. Open presentation
    prs = Presentation(template_bytes)

    # 3. Cover page — replace {{DATE}} on slide 0
    cover = prs.slides[0]
    for shape in cover.shapes:
        if shape.has_text_frame:
            _replace_text_in_frame(shape.text_frame, "{{DATE}}", date)

    # 4. Build section name → song lookup (case-insensitive keys)
    section_map = {
        s.get("name", "").strip().lower(): s.get("song", {})
        for s in sections
    }

    # 5. Fill each Mass-section slide, tracking index offset from duplications
    slide_offset = 0
    for section_idx, section_name in enumerate(SECTION_ORDER):
        current_index = 1 + section_idx + slide_offset
        song = section_map.get(section_name.lower(), {})

        before = len(prs.slides)
        _fill_section_slide(prs, current_index, song)
        slide_offset += len(prs.slides) - before

    # 6. Save to in-memory buffer and return
    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return output


# ---------------------------------------------------------------------------
# HTTP handler — Vercel-compatible + local dev server
# ---------------------------------------------------------------------------

class handler(BaseHTTPRequestHandler):
    """
    Handles POST /generate-ppt (or any path) for both Vercel serverless
    deployment and local development (python generate-ppt.py).
    """

    # Suppress default request log lines — we print our own
    def log_message(self, fmt, *args):  # noqa: N802
        print(f"[{self.client_address[0]}] {fmt % args}", flush=True)

    def do_GET(self):  # noqa: N802
        self._send_text(200, "generate-ppt is running. POST JSON to this endpoint.")

    def do_POST(self):  # noqa: N802
        content_length = int(self.headers.get("Content-Length", 0))
        raw_body = self.rfile.read(content_length)

        # Parse JSON
        try:
            body = json.loads(raw_body)
        except json.JSONDecodeError as exc:
            self._send_json_error(400, f"Invalid JSON: {exc}")
            return

        template_url = body.get("template_url", "").strip()
        date         = body.get("date", "")
        sections     = body.get("sections", [])

        if not template_url:
            self._send_json_error(400, "Missing required field: template_url")
            return

        print(f"[generate-ppt] Generating for date='{date}', sections={len(sections)}", flush=True)

        try:
            pptx_buf = generate_presentation(template_url, date, sections)
        except requests.HTTPError as exc:
            msg = f"Failed to download template: {exc}"
            print(f"[generate-ppt] ERROR: {msg}", flush=True)
            self._send_json_error(502, msg)
            return
        except Exception:
            tb = traceback.format_exc()
            print(f"[generate-ppt] TRACEBACK:\n{tb}", flush=True)
            self._send_json_error(500, f"Presentation generation failed:\n{tb}")
            return

        pptx_bytes = pptx_buf.read()
        print(f"[generate-ppt] Success — sending {len(pptx_bytes):,} bytes", flush=True)

        self.send_response(200)
        self.send_header(
            "Content-Type",
            "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
        self.send_header("Content-Disposition", 'attachment; filename="Mass_Presentation.pptx"')
        self.send_header("Content-Length", str(len(pptx_bytes)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(pptx_bytes)

    def do_OPTIONS(self):  # noqa: N802
        """Handle CORS preflight requests."""
        self.send_response(204)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "POST, GET, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()

    def _send_json_error(self, code: int, message: str):
        body = json.dumps({"error": message}).encode()
        self.send_response(code)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(body)

    def _send_text(self, code: int, message: str):
        body = message.encode()
        self.send_response(code)
        self.send_header("Content-Type", "text/plain")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)


# ---------------------------------------------------------------------------
# Local dev entry-point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    from http.server import HTTPServer

    port = 8000
    print(f"[generate-ppt] Starting local dev server -> http://localhost:{port}", flush=True)
    server = HTTPServer(("", port), handler)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n[generate-ppt] Shutting down.", flush=True)
