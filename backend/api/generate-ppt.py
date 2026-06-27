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
from http.server import BaseHTTPRequestHandler

import requests
from pptx import Presentation
from pptx.util import Pt
from lxml import etree


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _replace_text_in_frame(text_frame, old: str, new: str) -> bool:
    """
    Replace *old* with *new* inside a text frame while preserving the
    formatting of the first run that contains the placeholder.
    Returns True if a replacement was made.
    """
    for para in text_frame.paragraphs:
        for run in para.runs:
            if old in run.text:
                run.text = run.text.replace(old, new)
                return True
    return False


def _frame_contains(text_frame, placeholder: str) -> bool:
    """Return True if *placeholder* appears anywhere in the text frame."""
    for para in text_frame.paragraphs:
        for run in para.runs:
            if placeholder in run.text:
                return True
    return False


def _set_frame_text(text_frame, new_text: str):
    """
    Replace the entire content of a text frame with *new_text*.
    Multi-line input (split on \\n) creates separate paragraph elements so
    that line breaks render correctly in PowerPoint.
    Formatting (font name, size, bold, colour) is copied from the first
    existing run.
    """
    # Snapshot formatting from the first run (if available)
    first_run = None
    for para in text_frame.paragraphs:
        for run in para.runs:
            first_run = run
            break
        if first_run:
            break

    # Clear existing paragraphs by removing all <a:p> children from the txBody
    txBody = text_frame._txBody
    for p_elem in txBody.findall(
        ".//{http://schemas.openxmlformats.org/drawingml/2006/main}p"
    ):
        txBody.remove(p_elem)

    lines = new_text.split("\n")
    NS = "http://schemas.openxmlformats.org/drawingml/2006/main"

    for line in lines:
        p_elem = etree.SubElement(txBody, f"{{{NS}}}p")
        r_elem = etree.SubElement(p_elem, f"{{{NS}}}r")

        # Copy run properties from first_run if we have them
        if first_run is not None and first_run._r.find(f"{{{NS}}}rPr") is not None:
            rPr = copy.deepcopy(first_run._r.find(f"{{{NS}}}rPr"))
            r_elem.insert(0, rPr)

        t_elem = etree.SubElement(r_elem, f"{{{NS}}}t")
        t_elem.text = line


def _duplicate_slide(prs: Presentation, slide_index: int):
    """
    Duplicate the slide at *slide_index* and insert it immediately after.
    Returns the new (duplicate) slide object.
    """
    template_slide = prs.slides[slide_index]
    slide_layout = template_slide.slide_layout

    # Add a blank new slide using the same layout
    new_slide = prs.slides.add_slide(slide_layout)

    # Deep-copy the XML tree from the template slide into the new slide
    new_slide._element.spTree.clear()
    for child in template_slide._element.spTree:
        new_slide._element.spTree.append(copy.deepcopy(child))

    # Move the new slide to immediately after the template slide in the XML
    xml_slides = prs.slides._sldIdLst
    # The new slide was appended at the end — move it to slide_index + 1
    new_slide_elem = xml_slides[-1]
    xml_slides.remove(new_slide_elem)
    xml_slides.insert(slide_index + 1, new_slide_elem)

    return prs.slides[slide_index + 1]


def _fill_section_slide(prs: Presentation, slide_index: int, song: dict):
    """
    Fill a single Mass-section slide (and duplicate it as needed).

    *song* is expected to have a "lyrics" key: a list of dicts, each with
    "label" and "text".  Labels containing "chorus" (case-insensitive) are
    treated as the chorus and appended after every verse.

    The first verse replaces {{LYRICS}} on the original slide.
    Each additional verse gets a duplicated slide.
    """
    lyrics = song.get("lyrics", [])
    if not lyrics:
        # Nothing to fill — just clear the placeholder
        slide = prs.slides[slide_index]
        for shape in slide.shapes:
            if shape.has_text_frame and _frame_contains(shape.text_frame, "{{LYRICS}}"):
                _replace_text_in_frame(shape.text_frame, "{{LYRICS}}", "")
        return

    # Separate verses and chorus
    chorus_text = ""
    verses = []
    for item in lyrics:
        label = item.get("label", "").lower()
        text  = item.get("text", "")
        if "chorus" in label:
            chorus_text = text
        else:
            verses.append(text)

    # If there are no explicit verses (only a chorus), treat the chorus as the
    # sole "verse" so we still fill the slide.
    if not verses:
        verses = [chorus_text]
        chorus_text = ""

    def _build_slide_text(verse_text: str) -> str:
        if chorus_text:
            return verse_text + "\n\n" + chorus_text
        return verse_text

    # --- Fill the original slide with Verse 1 ---
    original_slide = prs.slides[slide_index]
    for shape in original_slide.shapes:
        if shape.has_text_frame and _frame_contains(shape.text_frame, "{{LYRICS}}"):
            _set_frame_text(shape.text_frame, _build_slide_text(verses[0]))
            break

    # --- Duplicate slide for each additional verse ---
    for i, verse_text in enumerate(verses[1:], start=1):
        # After each duplication the new slide is at slide_index + i
        _duplicate_slide(prs, slide_index + i - 1)
        dup_slide = prs.slides[slide_index + i]

        for shape in dup_slide.shapes:
            if shape.has_text_frame and _frame_contains(shape.text_frame, "{{LYRICS}}"):
                _set_frame_text(shape.text_frame, _build_slide_text(verse_text))
                break


# ---------------------------------------------------------------------------
# Core generation function
# ---------------------------------------------------------------------------

# Ordered list of section names matching slides 1-9 (0-indexed: 1 to 9)
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
    Download *template_url*, fill placeholders, and return the finished
    presentation as a BytesIO buffer.
    """
    # 1. Download template into memory
    response = requests.get(template_url, timeout=30)
    response.raise_for_status()
    template_bytes = io.BytesIO(response.content)

    # 2. Open the presentation
    prs = Presentation(template_bytes)

    # 3. Fill Cover Page — Slide 0: replace {{DATE}}
    cover_slide = prs.slides[0]
    for shape in cover_slide.shapes:
        if shape.has_text_frame:
            _replace_text_in_frame(shape.text_frame, "{{DATE}}", date)

    # 4. Build a lookup: section name → song dict (case-insensitive)
    section_map = {}
    for section in sections:
        name = section.get("name", "").strip()
        section_map[name.lower()] = section.get("song", {})

    # 5. Fill each Mass-section slide.
    #    We must process slides in REVERSE order when duplicating so that
    #    earlier slide indices are not shifted before we process them.
    #    Instead, we process forward and track the current index offset
    #    caused by any inserted duplicate slides.
    slide_offset = 0  # grows whenever we add duplicate slides

    for section_idx, section_name in enumerate(SECTION_ORDER):
        # The base slide index in the ORIGINAL template (1-9) + any offset
        # introduced by previously duplicated slides.
        current_slide_index = 1 + section_idx + slide_offset

        song = section_map.get(section_name.lower(), {})

        # Count slides before filling so we can calculate offset increase
        slides_before = len(prs.slides)
        _fill_section_slide(prs, current_slide_index, song)
        slides_after = len(prs.slides)

        slide_offset += slides_after - slides_before

    # 6. Save to a BytesIO buffer and return
    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return output


# ---------------------------------------------------------------------------
# HTTP handler (works with Vercel / simple http.server)
# ---------------------------------------------------------------------------

class handler(BaseHTTPRequestHandler):
    """
    Minimal HTTP handler compatible with the Vercel Python runtime.
    For local testing you can also run this file directly.
    """

    def do_POST(self):  # noqa: N802
        content_length = int(self.headers.get("Content-Length", 0))
        raw_body = self.rfile.read(content_length)

        try:
            body = json.loads(raw_body)
        except json.JSONDecodeError as exc:
            self._send_error(400, f"Invalid JSON: {exc}")
            return

        template_url = body.get("template_url")
        date         = body.get("date", "")
        sections     = body.get("sections", [])

        if not template_url:
            self._send_error(400, "Missing required field: template_url")
            return

        try:
            pptx_buffer = generate_presentation(template_url, date, sections)
        except requests.HTTPError as exc:
            self._send_error(502, f"Failed to download template: {exc}")
            return
        except Exception as exc:  # noqa: BLE001
            self._send_error(500, f"Presentation generation failed: {exc}")
            return

        pptx_bytes = pptx_buffer.read()

        self.send_response(200)
        self.send_header(
            "Content-Type",
            "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
        self.send_header("Content-Disposition", 'attachment; filename="presentation.pptx"')
        self.send_header("Content-Length", str(len(pptx_bytes)))
        self.end_headers()
        self.wfile.write(pptx_bytes)

    def do_GET(self):  # noqa: N802
        self.send_response(200)
        self.send_header("Content-Type", "text/plain")
        self.end_headers()
        self.wfile.write(b"generate-ppt is running. Send a POST request with JSON body.")

    def _send_error(self, code: int, message: str):
        body = json.dumps({"error": message}).encode()
        self.send_response(code)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)


# ---------------------------------------------------------------------------
# Local dev server entry-point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    from http.server import HTTPServer

    port = 8000
    print(f"Starting local dev server on http://localhost:{port}")
    server = HTTPServer(("", port), handler)
    server.serve_forever()
