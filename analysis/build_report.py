"""Builds report/SunVolt_Project_Report.docx from the charts and dashboard.
Run from the project root after sunvolt_analysis.py:  python analysis/build_report.py

With BODY_OUT=<path> set, the cover page is left out and the file is written to that path,
ready for .assignment/merge.py to place into the Eduvos submission template.
"""
import os
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

ROOT = Path(__file__).resolve().parents[1]
CH = ROOT / "charts"
BODY_OUT = os.environ.get("BODY_OUT")
OUT = Path(BODY_OUT) if BODY_OUT else ROOT / "report" / "SunVolt_Project_Report.docx"

NAVY = RGBColor(0x14, 0x30, 0x4F)
GREY = RGBColor(0x52, 0x51, 0x4E)

doc = Document()
st = doc.styles["Normal"]
st.font.name = "Arial"
st.font.size = Pt(11)
st.paragraph_format.space_after = Pt(6)
st.paragraph_format.line_spacing = 1.15
for name, size in [("Heading 1", 16), ("Heading 2", 13), ("Heading 3", 11.5)]:
    h = doc.styles[name]
    h.font.name = "Arial"
    h.font.size = Pt(size)
    h.font.color.rgb = NAVY
    h.font.bold = True
    rpr = h.element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rpr.append(rf)
    for a in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rf.set(qn(a), "Arial")

sec = doc.sections[0]
sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)  # A4
A4_W, A4_H = Cm(21.0), Cm(29.7)
for s in doc.sections:
    s.left_margin = s.right_margin = Cm(2.2)
    s.top_margin = s.bottom_margin = Cm(2)

fig_no = [0]


def p(text, bold_lead=None, style=None):
    para = doc.add_paragraph(style=style)
    if bold_lead:
        para.add_run(bold_lead).bold = True
    para.add_run(text)
    return para


def bullets(items):
    for it in items:
        if isinstance(it, tuple):
            p(it[1], it[0], "List Bullet")
        else:
            p(it, style="List Bullet")


def fig(name, caption, width=16.5):
    fig_no[0] += 1
    doc.add_picture(str(name), width=Cm(width))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    c = doc.add_paragraph()
    c.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = c.add_run(f"Figure {fig_no[0]}: {caption}")
    r.italic = True
    r.font.size = Pt(9)
    r.font.color.rgb = GREY


def table(rows, widths, header=True):
    t = doc.add_table(rows=len(rows), cols=len(rows[0]))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            cell = t.cell(i, j)
            cell.width = Cm(widths[j])
            cell.text = ""
            run = cell.paragraphs[0].add_run(str(val))
            run.font.size = Pt(9.5)
            if header and i == 0:
                run.bold = True
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                shd = OxmlElement("w:shd")
                shd.set(qn("w:val"), "clear")
                shd.set(qn("w:fill"), "14304F")
                cell._tc.get_or_add_tcPr().append(shd)
    doc.add_paragraph()
    return t


# ------------------------------------------------------------------ cover
if not BODY_OUT:
    t = doc.add_paragraph()
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    t.paragraph_format.space_before = Pt(120)
    r = t.add_run("SunVolt Energy\nInternational Home Solar Backup System Analysis")
    r.bold = True
    r.font.size = Pt(22)
    r.font.color.rgb = NAVY
    t = doc.add_paragraph()
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = t.add_run("ITCDA4-34 Project")
    r.font.size = Pt(14)
    t = doc.add_paragraph()
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    t.paragraph_format.space_before = Pt(40)
    for line in ["Student name: ____________________", "Student number: ____________________", "Date: ____________________"]:
        t.add_run(line + "\n").font.size = Pt(12)
    doc.add_page_break()
else:
    t = doc.add_paragraph()
    r = t.add_run("SunVolt Energy: International Home Solar Backup System Analysis")
    r.bold = True
    r.font.size = Pt(14)
    r.font.color.rgb = NAVY

# ------------------------------------------------------------------ data prep
doc.add_heading("Data Preparation and Approach", 1)
p("The dataset was loaded and checked in Python (pandas) before any analysis was done. The main steps were:")
bullets([
    "The file had 1,026 rows. 26 of these were exact duplicate customers, so they were removed, leaving 1,000 unique customers.",
    "There were no missing values. Age and income bands were checked against the raw Age and Annual Income columns and matched.",
    "The main measure used throughout is the purchase rate: the percentage of customers in a group who bought a solar backup system. Across all 1,000 customers this is 48.1% (481 buyers).",
])
p("I first explored every column against the purchase outcome (exploratory analysis) to find out where the real differences were. "
  "Only the charts that show a clear, useful pattern were kept and turned into explanatory charts. Each one has a title that states the finding, "
  "grey for context, blue for groups that buy more than average and orange for groups that buy less. The dashed line on the bar charts is the 48% company average, "
  "so the reader can see straight away who is above or below it.")

# ------------------------------------------------------------------ Q1
doc.add_heading("Question 1: Exploratory and Explanatory Data Analysis", 1)

doc.add_heading("I. Customer Demographics and Solar System Purchases", 2)
doc.add_heading("Which demographics are most associated with purchases?", 3)
fig(CH / "q1_01_demographics.png", "Purchase rate by age group, annual income and gender")
bullets([
    ("Age is the strongest demographic. ", "Customers aged 30-39 buy at 59%, well above every other age group. "
     "After 40 the rate falls to 46%, and customers aged 60+ buy at only 32%. Buyers are also younger on average (43 years vs 45 for non-buyers)."),
    ("Income works like a threshold, not a ladder. ", "Customers earning under 30k buy at 40%, but every band from 30k upwards sits at about 49-50%. "
     "Earning 100k+ does not make someone more likely to buy than earning 40k. Once a customer can afford the system, income stops mattering."),
    ("Gender makes almost no difference. ", "Women buy at 49% and men at 47%, so there is no reason to split marketing by gender."),
])
fig(CH / "q1_02_education_occupation.png", "Purchase rate by education level and occupation")
p("Education and occupation add to this picture. Customers with a Bachelors or Graduate degree buy at 54-55%, while those with only partial high school buy at 26%. "
  "Professionals are the best occupation group (54%) and, surprisingly, Management is the weakest (42%). "
  "So the typical buyer is a customer in their 30s, with a degree and a professional job.")

doc.add_heading("Patterns among customers who are not buying", 3)
fig(CH / "q1_03_buyer_gap.png", "Groups furthest above and below the company purchase rate", width=14.5)
fig(CH / "q1_04_married_by_age.png", "Purchase rate by age group and marital status", width=14)
p("Figure 3 brings every factor onto one scale. The groups that are not buying share a clear profile: older, married, bigger households with more cars. "
  "Figure 4 shows that the drop after 40 is really a married customer problem. Single customers keep buying at 51-54% at every age, "
  "but married customers fall to 40% at 40-49 and 22% at 60+. The 384 married customers aged 40+ make up 38% of all customers and buy at just 37%. "
  "Other low groups are households with 5 dependents (22%), customers with 2+ vehicles (37%), those earning under 30k (40%) and those with partial high school education (26%).")
p("How SunVolt can engage these groups more effectively:", bold_lead=None)
bullets([
    ("Married customers over 40: ", "use family-focused messaging about protecting the whole household during power cuts, and target them through channels they already use such as email, community events and home shows."),
    ("Customers aged 60+: ", "offer in-home demonstrations, simple maintenance plans and a longer warranty, since this group is likely to value reassurance and low effort."),
    ("Large and lower-income households: ", "introduce monthly payment plans or rent-to-own options so the upfront cost is not the barrier."),
    ("Keep the strongest buyers: ", "continue digital and referral campaigns aimed at customers in their 30s and professionals, as they convert easily and can refer others."),
])

doc.add_heading("II. Geographical Insights", 2)
doc.add_heading("How purchases vary across regions", 3)
fig(CH / "q1_05_region.png", "Customer base size and purchase rate by region")
p("Purchases vary a lot by region. The Pacific has the highest purchase rate at 59% (113 of 192 customers), Europe is close to the average at 49% (148 of 300), "
  "and North America has the lowest rate at 43% (220 of 508). This matters because North America is SunVolt's biggest market: it holds half of all customers and still produces the most buyers in total, "
  "but it converts the worst. If North America reached the Pacific's rate, SunVolt would gain about 79 extra sales from customers it already has.")
fig(CH / "q1_06_region_distance.png", "Purchase rate by region and distance to the nearest installer")
p("Figure 6 helps explain the gap. In North America, customers within 5 km of an installer buy at around 43-68%, but those 5-10 km away buy at 29% and those 10+ km away at 24%. "
  "About a third of North American customers (171) live more than 5 km from an installer, compared with only 34 in Europe. "
  "North America is also where married customers lag most (36% vs 54% for singles). The Pacific stays strong at every distance, with 73% of customers within 1 km buying. "
  "Europe is a lower-income market (average income of about 41k, compared with about 63k elsewhere), and its 104 customers under 30k buy at 38%.")

doc.add_heading("Highest and lowest regions, and strategies for underperformers", 3)
p("Highest: Pacific (59%). Lowest: North America (43%), followed by Europe (49%).")
bullets([
    ("North America, expand installer coverage: ", "add partner installers or mobile installation teams in areas more than 5 km from a branch, and run local campaigns in those suburbs once coverage is in place."),
    ("North America, target families: ", "build a family backup package with a payment plan and promote it to married households, the group furthest behind in this region."),
    ("Europe, make it affordable: ", "offer a smaller entry-level system and financing aimed at the large low-income group, rather than pushing premium packages."),
    ("Pacific, protect and learn: ", "use Pacific customers for referral rewards and case studies, and copy the Pacific's sales and installation approach in the other regions."),
])

doc.add_heading("III. Purchase Influences", 2)
fig(CH / "q1_07_household_influences.png", "Purchase rate by marital status, homeownership and number of dependents")
bullets([
    ("Marital status has a clear effect. ", "Single customers buy at 54%, married customers at 43%, an 11 percentage point gap across a large group (538 married customers)."),
    ("Dependents matter, mainly at the extremes. ", "Households with 1 dependent buy most (57%), while those with 5 dependents buy least (22%). "
     "Figure 8 shows the effect is driven by married customers: single customers with dependents keep buying at 57-62%, but married customers with 2 or more dependents drop to 34%."),
    ("Homeownership has almost no effect. ", "Homeowners buy at 48% and renters at 49%. This is surprising, because homeowners are the ones who can install a system permanently. "
     "It means SunVolt is not yet converting its most natural market, as 683 customers (68%) own their homes."),
])
fig(CH / "q1_08_dependents_by_marital.png", "Purchase rate by number of dependents and marital status", width=14)
fig(CH / "q1_09_vehicles.png", "Purchase rate by number of vehicles owned", width=12)
doc.add_heading("Untapped segments", 3)
p("The table below lists large segments that buy well below average. The last column shows how many extra sales SunVolt would gain if each segment simply reached the 48% company average. "
  "The segments overlap, so the numbers should not be added together.")
table([
    ["Segment", "Customers", "Purchase rate", "Extra sales at 48%"],
    ["Households with 2+ vehicles", "486", "37%", "+56"],
    ["Married with 2+ dependents", "314", "34%", "+43"],
    ["Married, aged 40+", "384", "37%", "+42"],
    ["Married homeowners", "441", "44%", "+20"],
], [6.5, 2.8, 3, 3.5])
p("These are households that clearly have the space and the need for backup power (several family members, several cars, their own home) but are not buying. "
  "A likely reason is the cost of a system big enough for a large household. SunVolt could target them with a larger family bundle on a payment plan, "
  "and with messages about keeping the whole household running and adding value to the home.")

doc.add_heading("IV. Distance from Nearest Installer and Purchasing Behaviour", 2)
fig(CH / "q1_10_distance.png", "Purchase rate by distance from the nearest installer", width=14)
p("Customers who live closer to an installer buy more. Within 5 km the purchase rate is 53%, but beyond 5 km it drops to 36%, and for customers more than 10 km away it is only 30%. "
  "Inside the 5 km zone the pattern is not perfectly smooth (1-2 km is lower than 2-5 km), so the clear break is at about 5 km rather than every extra kilometre. "
  "Buyers also live in closer distance bands on average than non-buyers.")
fig(CH / "q1_11_distance_by_region.png", "Purchase rate within vs beyond 5 km, by region", width=14)
p("The effect is strongest in North America, where buying falls from 51% to 27% beyond 5 km. In the Pacific it still falls, but only to 50%, which suggests the Pacific handles distant customers better.")
p("Recommendations by distance:")
bullets([
    ("Within 5 km: ", "run neighbourhood campaigns, offer quick installation dates and reward referrals, since these customers already convert well and are the cheapest to serve."),
    ("5 to 10 km: ", "remove the extra cost of distance with free or flat-fee installation and group installs in the same area on set days to keep costs down."),
    ("More than 10 km: ", "partner with local electricians, offer a remote (video) site assessment, and consider a new branch or mobile team where enough distant customers cluster, starting with North America."),
])
p("Note: these findings show which groups are linked with buying, not proof of what causes it. Some groups are small (for example Europe beyond 10 km has only 18 customers), "
  "so those results should be treated with care.", bold_lead=None).runs[0].italic = True

# ------------------------------------------------------------------ Q2 (landscape)
new = doc.add_section()
new.orientation = WD_ORIENT.LANDSCAPE
new.page_width, new.page_height = A4_H, A4_W
new.left_margin = new.right_margin = Cm(1.5)
new.top_margin = new.bottom_margin = Cm(1.5)
doc.add_heading("Question 2: Dashboard for the CEO and Investors", 1)
doc.add_picture(str(ROOT / "dashboard" / "sunvolt_dashboard.png"), height=Cm(14.8))
doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
fig_no[0] += 1
c = doc.add_paragraph()
c.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = c.add_run(f"Figure {fig_no[0]}: SunVolt Energy sales growth dashboard (interactive version: dashboard/sunvolt_dashboard.html)")
r.italic = True
r.font.size = Pt(9)
r.font.color.rgb = GREY

new = doc.add_section()
new.orientation = WD_ORIENT.PORTRAIT
new.page_width, new.page_height = A4_W, A4_H
new.left_margin = new.right_margin = Cm(2.2)
new.top_margin = new.bottom_margin = Cm(2)

doc.add_heading("Chart selection", 2)
p("The CEO and investors need to know where the growth is and what to do about it, not every detail from Question 1. "
  "I therefore kept only four charts that build one argument, and left out charts such as gender and education that do not lead to a decision:")
table([
    ["Dashboard chart", "From Q1", "Question it answers for the CEO"],
    ["1. Purchase rate by region", "Figure 5", "Where are we winning and losing?"],
    ["2. Within vs beyond 5 km by region", "Figure 11", "Why is North America behind?"],
    ["3. Gap from average", "Figure 3", "Who are our best and worst customers?"],
    ["4. Untapped segments", "Q1 III table", "How many sales are we missing, and where?"],
], [5.8, 2.6, 7.6])

doc.add_heading("Design choices", 2)
bullets([
    ("Layout and flow: ", "the dashboard reads top to bottom and left to right like a story: the headline message first, then five KPI cards, then Where, Why there, Who, Untapped, and finally What we recommend. "
     "Each card is numbered so the presenter can walk the audience through it in order."),
    ("Headline as the main message: ", "the banner states the recommendation in one sentence, so an executive who only reads the top still gets the point."),
    ("Chart types: ", "horizontal bars for regions and segments because the labels are long and easy to compare; grouped bars to compare near vs far within each region; "
     "a diverging bar chart for the gap from average so above and below average are seen at a glance. No pie charts, 3D or dual axes were used."),
    ("Colour: ", "only three colours with fixed meanings across every chart: blue for strong groups, orange for problem or opportunity groups and grey for context. "
     "Blue and orange are also safe for colour-blind readers. A dark navy banner gives the page a professional, investor-ready look."),
    ("Decluttering: ", "gridlines, axis lines and chart borders were removed, and values are labelled directly on the bars so no one has to read an axis. Titles state the finding instead of just naming the chart."),
    ("Interactivity: ", "in the HTML version, hovering over any bar shows the exact rate and number of customers, so investors can check the detail without cluttering the view."),
])

doc.add_heading("Recommended actions", 2)
bullets([
    ("Expand installer reach in North America: ", "add partner installers or mobile teams for the 171 customers beyond 5 km who buy at only 27%. Bringing them to the near rate of 51% would add about 41 sales."),
    ("Remove the cost of distance: ", "offer free or flat-fee installation and remote site assessments for customers more than 5 km away."),
    ("Launch a family and multi-car bundle: ", "a larger system on a payment plan, marketed to married households over 40 and homes with 2+ vehicles, which buy at 34-37%."),
    ("Protect the core buyers: ", "keep campaigns and referral rewards aimed at customers aged 30-39, professionals and Pacific customers, who buy at 54-59%."),
])

# ------------------------------------------------------------------ Q3
doc.add_page_break()
doc.add_heading("Question 3: Story and Storyboard", 1)
doc.add_heading("a. The Big Idea", 2)
bi = doc.add_paragraph()
bi.paragraph_format.left_indent = Cm(0.8)
bi.paragraph_format.right_indent = Cm(0.8)
r = bi.add_run("SunVolt should bring installers closer to North American customers and design a family offer for married households, "
               "because this is where we are losing the most sales; if we act, we can move our purchase rate from 48% towards the 59% the Pacific already achieves.")
r.bold = True
r.font.color.rgb = NAVY
p("This Big Idea follows the three parts taught in the module: it gives SunVolt's point of view (focus on distance and families in North America), "
  "it says what is at stake (sales we are already losing in our biggest market) and it is one complete sentence.")

doc.add_heading("a. The 3-minute story", 2)
story = [
    ("Setting the scene. ",
     "Good morning. Over the last year SunVolt has built a base of 1,000 customers across three regions. Electricity prices keep rising and families want power they can count on, "
     "so demand for solar backup is there. But right now, fewer than half of our customers, 48%, actually buy a system. "
     "That means more than half of the people who already know us are walking away."),
    ("Where the problem is. ",
     "When we split this by region, the picture changes. In the Pacific, 59% of customers buy. In Europe it is 49%. In North America it is only 43%. "
     "And North America is not a small market. It is half of our entire customer base. Our biggest region is our weakest one."),
    ("Why it is happening. ",
     "The data points to two reasons. The first is distance. Customers who live within 5 km of an installer buy at 53%. Beyond 5 km, that falls to 36%. "
     "In North America the drop is even sharper, from 51% down to 27%, and one in three North American customers lives beyond that 5 km line. "
     "The second reason is the type of household. Our best buyers are people in their 30s with smaller households. "
     "Married customers over 40, and married families with two or more children, buy at only 34 to 37%. "
     "These households have the most to lose in a power cut, yet they are the ones not buying, most likely because a system big enough for a family feels too expensive or too much effort."),
    ("What it is costing us. ",
     "If North America simply matched the Pacific, we would win about 79 more sales from customers we already have, without spending a cent on finding new ones. "
     "If our married families and multi-car households bought at our own average, that is roughly another 40 to 55 sales in each group."),
    ("What we should do. ",
     "So I am asking for three decisions today. First, approve partner installers or mobile install teams for North American areas more than 5 km from a branch. "
     "Second, remove the cost of distance by offering free or flat-fee installation for those customers. "
     "Third, launch a family backup bundle with a monthly payment plan, aimed at married households and homes with more than one car. "
     "At the same time we keep doing what already works with our customers in their 30s and in the Pacific."),
    ("Closing. ",
     "We do not have a demand problem. We have a reach problem and an offer problem, and both can be fixed. "
     "If we bring installers closer and build an offer for families, we can move from 48% towards the 59% the Pacific already shows is possible. "
     "I suggest we start with a three-month pilot in North America and report back on the results. Thank you."),
]
for lead, text in story:
    p(text, lead)
p("(Approximately 460 words, which is about three minutes when spoken at a normal presenting pace.)").runs[0].italic = True

doc.add_heading("b. Storyboard", 2)
p("The storyboard below plans each slide of the presentation before it is built. It follows the story arc above: the setting, rising tension, the climax, then the resolution and call to action. "
  "Each slide has one message, written as the slide title.")
table([
    ["#", "Story stage", "Slide title (main message)", "Visual", "Speaker focus"],
    ["1", "Setting", "Demand for solar backup is there, but fewer than half of our customers buy", "Three KPI tiles: 1,000 customers, 481 buyers, 48%", "Context and why it matters now"],
    ["2", "Rising action", "Our biggest market, North America, converts the worst", "Region bar chart, North America in orange", "Size vs performance of each region"],
    ["3", "Rising action", "Customers more than 5 km from an installer rarely buy", "Within vs beyond 5 km bars by region", "Distance as the first barrier"],
    ["4", "Rising action", "Married families over 40 are the largest group not buying", "Married vs single line chart", "Household type as the second barrier"],
    ["5", "Climax", "We are leaving around 79 sales on the table in North America alone", "Untapped segments bar chart with extra sales", "The cost of doing nothing"],
    ["6", "Resolution", "Three actions to close the gap", "Three icons: installers, free installation, family bundle", "What we will do and for whom"],
    ["7", "Call to action", "Approve a 3-month North America pilot", "Big Idea sentence with 48% to 59% target", "The decision we need today"],
], [0.8, 2.3, 5.2, 4.6, 3.6])

doc.save(OUT)
print("saved", OUT)
