"""Final PDF: signed template pages 1-2 (coversheet, declaration), report body, template pages 4-5
(AI Evidence, AI Disclosure).

Usage: python stitch.py <template.pdf> <body.pdf> <out.pdf>

The template PDF is the ITSMA4-14 one, so the module code and lecturer are swapped on page 1
in the same font, size and colour.
"""
import sys

import pymupdf

ARIAL_BOLD = "/System/Library/Fonts/Supplemental/Arial Bold.ttf"
NAVY = (0, 0x20 / 255, 0x60 / 255)
SWAP = {"ITSMA4-14": "ITCDA4-34", "Mr Sydney Khumalo": "Mr Gilbert Lekula (Ngobeni)"}

tpl = pymupdf.open(sys.argv[1])
body = pymupdf.open(sys.argv[2])

cover = tpl[0]
spans = [s for b in cover.get_text("dict")["blocks"] for l in b.get("lines", []) for s in l["spans"]
         if s["text"].strip() in SWAP]
assert len(spans) == len(SWAP), spans
for s in spans:
    cover.add_redact_annot(pymupdf.Rect(s["bbox"]))
cover.apply_redactions(images=pymupdf.PDF_REDACT_IMAGE_NONE, graphics=pymupdf.PDF_REDACT_LINE_ART_NONE)
for s in spans:
    cover.insert_text(s["origin"], SWAP[s["text"].strip()], fontsize=s["size"],
                      fontname="ArialBold", fontfile=ARIAL_BOLD, color=NAVY)

out = pymupdf.open()
out.insert_pdf(tpl, from_page=0, to_page=1)
out.insert_pdf(body)

# number the AI pages the same way as the body pages
num = [s for b in body[1].get_text("dict")["blocks"] for l in b.get("lines", []) for s in l["spans"]
       if s["text"].strip() == "4"]
num = min(num, key=lambda s: s["origin"][1])
right, base, size = num["bbox"][2], num["origin"][1], num["size"]
first_ai = out.page_count
out.insert_pdf(tpl, from_page=3, to_page=4)
for i in range(first_ai, out.page_count):
    p, t = out[i], str(i + 1)
    w = pymupdf.get_text_length(t, fontname="helv", fontsize=size)
    p.insert_text((right - w, base), t, fontname="helv", fontsize=size)
out.save(sys.argv[3], garbage=3, deflate=True)
print("pages", out.page_count)
