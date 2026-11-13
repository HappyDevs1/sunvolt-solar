"""Adds the native charts to report/SunVolt_Presentation.pptx (run after build_deck.js).

python-pptx charts are used because they display in both PowerPoint and Keynote.
"""
from pathlib import Path

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION, XL_MARKER_STYLE, XL_TICK_LABEL_POSITION, XL_TICK_MARK
from pptx.util import Inches, Pt

DECK = Path(__file__).resolve().parents[1] / "report" / "SunVolt_Presentation.pptx"

BLUE, ORANGE, GREY = RGBColor(0x2A, 0x78, 0xD6), RGBColor(0xEB, 0x68, 0x34), RGBColor(0xC9, 0xC8, 0xC2)
INK, INK2, LINE = RGBColor(0x0B, 0x0B, 0x0B), RGBColor(0x52, 0x51, 0x4E), RGBColor(0xE4, 0xE2, 0xDC)
X, Y, W, H = Inches(0.6), Inches(1.7), Inches(7.6), Inches(4.9)


def base(chart, title=None, legend=False, legend_pos=XL_LEGEND_POSITION.TOP):
    chart.font.name = "Arial"
    chart.font.size = Pt(14)
    chart.font.color.rgb = INK2
    chart.has_title = bool(title)
    if title:
        chart.chart_title.text_frame.text = title
        r = chart.chart_title.text_frame.paragraphs[0].runs[0]
        r.font.size, r.font.bold, r.font.color.rgb = Pt(14), False, INK2
    chart.has_legend = legend
    if legend:
        chart.legend.position = legend_pos
        chart.legend.include_in_layout = False
        chart.legend.font.size = Pt(14)
    cat = chart.category_axis
    cat.has_major_gridlines = False
    cat.format.line.color.rgb = GREY
    cat.tick_labels.font.size = Pt(14)


def hide_value_axis(chart, max_val):
    va = chart.value_axis
    va.has_major_gridlines = False
    va.tick_label_position = XL_TICK_LABEL_POSITION.NONE
    va.major_tick_mark = va.minor_tick_mark = XL_TICK_MARK.NONE
    va.format.line.fill.background()
    va.minimum_scale, va.maximum_scale = 0, max_val


def labels(plot, fmt):
    plot.has_data_labels = True
    dl = plot.data_labels
    dl.number_format, dl.number_format_is_linked = fmt, False
    dl.position = XL_LABEL_POSITION.OUTSIDE_END
    dl.font.size, dl.font.bold, dl.font.color.rgb = Pt(14), True, INK


def bar(slide, cats, vals, colours, fmt, max_val, title, horizontal=True):
    cd = CategoryChartData()
    cd.categories = cats
    cd.add_series("Value", vals)
    kind = XL_CHART_TYPE.BAR_CLUSTERED if horizontal else XL_CHART_TYPE.COLUMN_CLUSTERED
    gf = slide.shapes.add_chart(kind, X, Y, W, H, cd)
    ch = gf.chart
    base(ch, title)
    hide_value_axis(ch, max_val)
    plot = ch.plots[0]
    plot.gap_width = 45
    labels(plot, fmt)
    for i, c in enumerate(colours):
        pt = plot.series[0].points[i]
        pt.format.fill.solid()
        pt.format.fill.fore_color.rgb = c
    return gf


prs = Presentation(DECK)
slides = prs.slides

# Slide 3: purchase rate by region
g = bar(slides[2], ["Pacific", "Europe", "North America"], [58.9, 49.3, 43.3],
        [BLUE, GREY, ORANGE], '0"%"', 70, "Purchase rate by region", horizontal=False)
g.name = "Region purchase rate chart"

# Slide 4: within vs beyond 5 km by region
cd = CategoryChartData()
cd.categories = ["Pacific", "Europe", "North America"]
cd.add_series("Within 5 km", (68.09, 50.75, 51.34))
cd.add_series("Beyond 5 km", (50.0, 38.24, 27.49))
g = slides[3].shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, X, Y, W, H, cd)
g.name = "Distance by region chart"
ch = g.chart
base(ch, legend=True)
hide_value_axis(ch, 80)
plot = ch.plots[0]
plot.gap_width, plot.overlap = 60, -5
labels(plot, '0"%"')
for s, c in zip(plot.series, [BLUE, ORANGE]):
    s.format.fill.solid()
    s.format.fill.fore_color.rgb = c

# Slide 5: married vs single by age (line)
cd = CategoryChartData()
cd.categories = ["Under 30", "30-39", "40-49", "50-59", "60+"]
cd.add_series("Single", (39.6, 59.3, 54.4, 52.9, 51.3))
cd.add_series("Married", (46.7, 59.7, 40.3, 41.9, 22.4))
g = slides[4].shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS, X, Y, W, H, cd)
g.name = "Married vs single by age chart"
ch = g.chart
base(ch, "Purchase rate by age group", legend=True, legend_pos=XL_LEGEND_POSITION.BOTTOM)
va = ch.value_axis
va.minimum_scale, va.maximum_scale, va.major_unit = 0, 70, 10
va.tick_labels.number_format, va.tick_labels.number_format_is_linked = '0"%"', False
va.tick_labels.font.size = Pt(12)
va.has_major_gridlines = True
va.major_gridlines.format.line.color.rgb = LINE
va.format.line.fill.background()
for s, c in zip(ch.plots[0].series, [GREY, ORANGE]):
    s.smooth = False
    s.format.line.color.rgb = c
    s.format.line.width = Pt(3)
    s.marker.style = XL_MARKER_STYLE.CIRCLE
    s.marker.size = 10
    s.marker.format.fill.solid()
    s.marker.format.fill.fore_color.rgb = c
    s.marker.format.line.color.rgb = c

# Slide 6: untapped segments
g = bar(slides[5], ["Own 2+\nvehicles", "Married,\n2+ dependents", "Married,\naged 40+", "North America,\nbeyond 5 km"],
        [56, 43, 42, 35], [ORANGE] * 4, '"+"0', 70, "Extra sales if each group bought at our 48% average", horizontal=False)
g.name = "Untapped segments chart"

prs.save(DECK)
print("charts added to", DECK)
