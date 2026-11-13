# SunVolt Energy: Home Solar Backup Analysis

ITCDA4-34 project. Customer analysis, a CEO/investor dashboard, and a presentation story for SunVolt Energy.

## Folders

| Folder | Contents |
|---|---|
| `data/raw/` | Original dataset (unchanged) |
| `data/clean/` | Cleaned data (26 duplicate rows removed, 1,000 customers) |
| `analysis/` | Scripts that clean the data, build the charts and build the report |
| `charts/` | Question 1 charts |
| `dashboard/` | Question 2 dashboard (HTML and PNG) |
| `report/` | Word document (Questions 1 to 3) and PowerPoint deck with speaker notes (Question 4) |

## Rebuild

```bash
pip install -r requirements.txt
python analysis/sunvolt_analysis.py
python analysis/build_report.py
npm install pptxgenjs react react-dom react-icons sharp
node analysis/build_deck.js
python analysis/add_deck_charts.py
```

### Submission file (Eduvos template)

The report is placed into the signed Eduvos assessment template in `.assignment/template/` (coversheet, declaration, Assessment page, AI Evidence and AI Disclosure pages). Needs LibreOffice and PyMuPDF.

```bash
BODY_OUT=.assignment/body.docx python analysis/build_report.py
python .assignment/merge.py .assignment/body.docx "report/ITCDA4-34 - Project 2 - Online - EDUV13317839.docx"
python .assignment/merge.py .assignment/body.docx /tmp/bodypdf.docx --pdf-body
soffice --headless --convert-to pdf --outdir /tmp /tmp/bodypdf.docx
python .assignment/stitch.py .assignment/template/signed-fixed.pdf /tmp/bodypdf.pdf "report/ITCDA4-34 - Project 2 - Online - EDUV13317839.pdf"
```

The PDF takes the coversheet pages from the Pages-exported `signed-fixed.pdf` because LibreOffice pushes the coversheet's group table onto an extra page.

The deck's charts are added with python-pptx because charts written by pptxgenjs show up blank in Keynote.
