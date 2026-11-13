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

The deck's charts are added with python-pptx because charts written by pptxgenjs show up blank in Keynote.
