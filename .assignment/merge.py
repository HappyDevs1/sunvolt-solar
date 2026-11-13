"""Place the report body (body.docx, from BODY_OUT=... build_report.py) into the signed Eduvos template.

Usage: python merge.py <body.docx> <out.docx> [--pdf-body]

--pdf-body writes only the Assessment page and the report body, numbered from page 3, for
stitch.py to place between the template's own PDF pages (LibreOffice does not lay out the
coversheet the same way Pages does, so the coversheet pages come from template/signed-fixed.pdf).

Result: coversheet and declaration (template pages 1-2), the Assessment page with the report body,
then the AI Evidence and AI Disclosure pages. The body keeps its own margins and its landscape
dashboard page through separate sections.
"""
import copy
import re
import sys
import zipfile
from pathlib import Path

from lxml import etree

here = Path(__file__).parent
tpl_dir = here / "template" / "x"
body = zipfile.ZipFile(sys.argv[1])
out = Path(sys.argv[2])
pdf_body = "--pdf-body" in sys.argv

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
q = lambda tag: "{%s}%s" % (W, tag)

# ---------------------------------------------------------------- body
broot = etree.fromstring(body.read("word/document.xml"))
bbody = broot.find(q("body"))

# every body paragraph gets an explicit style, and the body's Normal becomes SVNormal so it
# does not clash with the template's Normal (Times New Roman 12)
for p in bbody.iter(q("p")):
    ppr = p.find(q("pPr"))
    if ppr is None:
        ppr = etree.SubElement(p, q("pPr"))
        p.remove(ppr)
        p.insert(0, ppr)
    if ppr.find(q("pStyle")) is None:
        ps = etree.Element(q("pStyle"))
        ps.set(q("val"), "Normal")
        ppr.insert(0, ps)
for ps in bbody.iter(q("pStyle")):
    if ps.get(q("val")) == "Normal":
        ps.set(q("val"), "SVNormal")

# images get new relationship ids
brels = etree.fromstring(body.read("word/_rels/document.xml.rels"))
new_rels, media = [], {}
for i, rel in enumerate(r for r in brels if r.get("Type").endswith("/image")):
    old, new = rel.get("Id"), "rIdSV%d" % (i + 1)
    target = "media/sv_" + Path(rel.get("Target")).name
    media[target] = body.read("word/" + rel.get("Target"))
    new_rels.append('<Relationship Id="%s" Type="%s" Target="%s"/>' % (new, rel.get("Type"), target))
    for el in broot.iter():
        if el.get("{%s}embed" % "http://schemas.openxmlformats.org/officeDocument/2006/relationships") == old:
            el.set("{%s}embed" % R, new)

children = list(bbody)
body_sect = children[-1]
assert body_sect.tag == q("sectPr")
inner = "".join(etree.tostring(c, encoding="unicode") for c in children[:-1])
body_sect_xml = etree.tostring(body_sect, encoding="unicode")

# ---------------------------------------------------------------- template
doc = (tpl_dir / "word/document.xml").read_text()
doc = doc.replace("<w:t>ITSMA4-14</w:t>", "<w:t>ITCDA4-34</w:t>", 1)
doc = doc.replace("<w:t>Mr Sydney Khumalo</w:t>", "<w:t>Mr Gilbert Lekula (Ngobeni)</w:t>", 1)
assert "ITSMA" not in doc and "Khumalo" not in doc

final_sect = re.search(r"<w:sectPr>.*?</w:sectPr>", doc[doc.rindex("<w:sectPr>"):], re.S).group(0)


def para_with_sect(sect):
    return "<w:p><w:pPr>%s</w:pPr></w:p>" % sect


def page_break_para_before(text):
    """Start/end index of the page-break paragraph that comes just before `text`."""
    i = doc.index(text)
    br = doc.rindex('<w:br w:type="page"/>', 0, i)
    return doc.rindex("<w:p>", 0, br), doc.index("</w:p>", br) + len("</w:p>")

# 1) end the coversheet section where the template breaks to the Assessment page
s, e = page_break_para_before("<w:t>Assessment</w:t>")
doc = doc[:s] + para_with_sect(final_sect) + doc[e:]
body_start = s + len(para_with_sect(final_sect))

# 2) replace the placeholder (and the blank lines and page break after it) with the body,
#    then end the body's last section so the AI pages start on a new page
ph = doc.index("[This section is reserved for your assessment submission content]")
s = doc.rindex("<w:p>", 0, ph)
_, e = page_break_para_before("AI Evidence")
body_sect_xml = re.sub(r' xmlns:\w+="[^"]+"', "", body_sect_xml)
doc = doc[:s] + inner + para_with_sect(body_sect_xml) + doc[e:]
body_end = s + len(inner) + len(para_with_sect(body_sect_xml))

if pdf_body:
    content = doc[body_start:body_end - len(para_with_sect(body_sect_xml))]
    first = content.index("<w:sectPr")
    first = content.index(">", first) + 1
    content = (content[:first] + '<w:headerReference w:type="default" r:id="rId4"/>'
               + content[first:])
    content = re.sub(r"(<w:pgMar [^>]*/>)", r'\1<w:pgNumType w:start="3"/>', content, count=1)
    doc = doc[:doc.index("<w:body>") + 8] + content + body_sect_xml + "</w:body></w:document>"

# 3) the AI pages use the plain header, not the coversheet's first-page logo header
last = doc.rindex("<w:sectPr>")
doc = doc[:last] + doc[last:].replace('<w:titlePg w:val="1"/>', "", 1)

# namespaces the picture markup needs
root_open = re.search(r"<w:document[^>]*>", doc).group(0)
extra = ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'
doc = doc.replace(root_open, root_open[:-1] + extra + ">", 1)

# tidy two schema errors that come from Pages' Word export
doc = re.sub(r"<w:pgMar ((?:(?!w:gutter)[^>])*)/>", r'<w:pgMar \1 w:gutter="0"/>', doc)
settings = re.sub(r"<w:displayBackgroundShape[^>]*/>", "", (tpl_dir / "word/settings.xml").read_text())

# ---------------------------------------------------------------- styles
bs = body.read("word/styles.xml").decode()
styles = (tpl_dir / "word/styles.xml").read_text()
add = []
for sid in ["Normal", "Heading1", "Heading2", "Heading3", "ListBullet", "TableGrid"]:
    m = re.search(r'<w:style [^>]*w:styleId="%s".*?</w:style>' % sid, bs, re.S)
    assert m, sid
    st = m.group(0)
    if sid == "Normal":
        st = st.replace(' w:default="1"', "").replace('w:styleId="Normal"', 'w:styleId="SVNormal"')
        st = st.replace('<w:name w:val="Normal"/>', '<w:name w:val="SV Normal"/>')
    st = st.replace('<w:basedOn w:val="Normal"/>', '<w:basedOn w:val="SVNormal"/>')
    st = st.replace('<w:next w:val="Normal"/>', '<w:next w:val="SVNormal"/>')
    st = st.replace('<w:basedOn w:val="TableNormal"/>', '<w:basedOn w:val="Table Normal"/>')
    # theme font attributes would pick up the template's theme (Times), so use the named fonts only
    st = re.sub(r' w:(ascii|hAnsi|eastAsia|cs)Theme="[^"]*"', "", st)
    add.append(st)
styles = styles.replace("</w:styles>", "".join(add) + "</w:styles>")

rels = (tpl_dir / "word/_rels/document.xml.rels").read_text().replace(
    "</Relationships>",
    '<Relationship Id="rIdNum1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>'
    + "".join(new_rels) + "</Relationships>")
ct = (tpl_dir / "[Content_Types].xml").read_text().replace(
    "</Types>",
    '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>')

# page number, right-aligned, in the default header (first page keeps the Eduvos logo header)
rpr = '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="18"/></w:rPr>'
field = ('<w:p><w:pPr><w:jc w:val="right"/></w:pPr>'
         '<w:r>%s<w:fldChar w:fldCharType="begin"/></w:r>'
         '<w:r>%s<w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
         '<w:r>%s<w:fldChar w:fldCharType="separate"/></w:r>'
         '<w:r>%s<w:t>1</w:t></w:r>'
         '<w:r>%s<w:fldChar w:fldCharType="end"/></w:r></w:p>') % ((rpr,) * 5)
hdr = re.sub(r"<w:p>.*</w:p>", field, (tpl_dir / "word/header1.xml").read_text(), flags=re.S)

replaced = {
    "word/settings.xml": settings,
    "word/document.xml": doc,
    "word/styles.xml": styles,
    "word/_rels/document.xml.rels": rels,
    "[Content_Types].xml": ct,
    "word/header1.xml": hdr,
}
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
    for f in sorted(tpl_dir.rglob("*")):
        if f.is_dir():
            continue
        name = f.relative_to(tpl_dir).as_posix()
        z.writestr(name, replaced.get(name, f.read_bytes()))
    z.writestr("word/numbering.xml", body.read("word/numbering.xml"))
    for target, data in media.items():
        z.writestr("word/" + target, data)
print("written", out)
