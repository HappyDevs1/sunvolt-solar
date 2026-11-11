"""SunVolt Energy customer analysis (ITCDA4-34 project, Question 1).

Cleans the raw customer file and produces the charts used in the report.
Run from the project root:  python analysis/sunvolt_analysis.py
"""
from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
RAW = ROOT / "data" / "raw" / "SunVolt Solar Customer Data.xlsx"
CLEAN = ROOT / "data" / "clean"
CHARTS = ROOT / "charts"

# Colours: grey for context, blue for strong groups, orange for weak groups
GREY = "#c9c8c2"
BLUE = "#2a78d6"
ORANGE = "#eb6834"
INK = "#0b0b0b"
INK2 = "#52514e"

plt.rcParams.update({
    "font.family": "Arial",
    "font.size": 11,
    "axes.spines.top": False,
    "axes.spines.right": False,
    "axes.spines.left": False,
    "axes.edgecolor": "#8a8984",
    "axes.labelcolor": INK2,
    "xtick.color": INK2,
    "ytick.color": INK2,
    "figure.dpi": 150,
    "savefig.bbox": "tight",
    "savefig.facecolor": "white",
})

ORDERS = {
    "Age Group": ["Under 30", "30-39", "40-49", "50-59", "60+"],
    "Income Band": ["<30k", "30k-49k", "50k-69k", "70k-99k", "100k+"],
    "Distance from Nearest Installer": ["0-1 km", "1-2 km", "2-5 km", "5-10 km", "10+ km"],
    "Education Level": ["Partial High School", "High School", "Partial College", "Bachelors", "Graduate Degree"],
}


# ---------------------------------------------------------------- cleaning
def load_clean():
    df = pd.read_excel(RAW)
    before = len(df)
    df = df.drop_duplicates()
    print(f"Removed {before - len(df)} duplicate rows, {len(df)} customers left")
    assert df["Customer ID"].is_unique
    assert df.isna().sum().sum() == 0
    CLEAN.mkdir(parents=True, exist_ok=True)
    df.to_csv(CLEAN / "sunvolt_clean.csv", index=False)
    df.to_excel(CLEAN / "sunvolt_clean.xlsx", index=False)
    return df


def rate_table(df, col):
    t = df.groupby(col)["Purchase Flag"].agg(customers="count", buyers="sum")
    t["rate"] = t["buyers"] / t["customers"] * 100
    if col in ORDERS:
        t = t.reindex(ORDERS[col])
    return t


# ---------------------------------------------------------------- helpers
def bar_panel(ax, t, avg, colours, title=None, label_fmt="{:.0f}%"):
    labels = [str(i) for i in t.index]
    bars = ax.bar(labels, t["rate"], color=colours, width=0.65, edgecolor="white", linewidth=2)
    for b, v, n in zip(bars, t["rate"], t["customers"]):
        ax.text(b.get_x() + b.get_width() / 2, v + 1.2, label_fmt.format(v),
                ha="center", va="bottom", fontsize=10, color=INK, fontweight="bold",
                bbox=dict(facecolor="white", edgecolor="none", pad=0.5))
        ax.text(b.get_x() + b.get_width() / 2, 2, f"n={n}", ha="center", va="bottom",
                fontsize=8, color="white" if colours[list(bars).index(b)] != GREY else INK2)
    ax.axhline(avg, color=INK2, linestyle="--", linewidth=1, zorder=0)
    ax.set_ylim(0, 75)
    ax.set_yticks([])
    ax.tick_params(axis="x", length=0)
    if title:
        ax.set_title(title, loc="left", fontsize=11, color=INK2)


def pick(t, high=None, low=None):
    cols = []
    for i in t.index:
        if high is not None and i in high:
            cols.append(BLUE)
        elif low is not None and i in low:
            cols.append(ORANGE)
        else:
            cols.append(GREY)
    return cols


def headline(fig, title, subtitle):
    h = fig.get_figheight()
    fig.suptitle(title, x=0.01, y=1 + 0.6 / h, ha="left", fontsize=14, fontweight="bold", color=INK)
    fig.text(0.01, 1 + 0.25 / h, subtitle, ha="left", fontsize=10, color=INK2)


def save(fig, name):
    CHARTS.mkdir(exist_ok=True)
    fig.savefig(CHARTS / name)
    plt.close(fig)
    print("saved", name)


# ---------------------------------------------------------------- charts
def chart_demographics(df, avg):
    fig, axes = plt.subplots(1, 3, figsize=(13, 4), gridspec_kw={"width_ratios": [5, 5, 2]})
    t = rate_table(df, "Age Group")
    bar_panel(axes[0], t, avg, pick(t, ["30-39"], ["60+"]), "Age group")
    t = rate_table(df, "Income Band")
    bar_panel(axes[1], t, avg, pick(t, None, ["<30k"]), "Annual income")
    t = rate_table(df, "Gender")
    bar_panel(axes[2], t, avg, [GREY, GREY], "Gender")
    axes[0].text(4.45, avg + 1, f"avg {avg:.0f}%", fontsize=8, color=INK2, ha="right")
    headline(fig, "Customers aged 30-39 buy the most; age and low income matter, gender does not",
             "Purchase rate (% of customers who bought a solar backup system). Dashed line = company average.")
    save(fig, "q1_01_demographics.png")


def chart_education_occupation(df, avg):
    fig, axes = plt.subplots(1, 2, figsize=(13, 4))
    t = rate_table(df, "Education Level")
    t.index = ["Partial\nHigh School", "High\nSchool", "Partial\nCollege", "Bachelors", "Graduate\nDegree"]
    bar_panel(axes[0], t, avg, [ORANGE, GREY, GREY, BLUE, BLUE], "Education level")
    t = rate_table(df, "Occupation").sort_values("rate")
    t.index = [i.replace("Skilled Manual", "Skilled\nManual") for i in t.index]
    bar_panel(axes[1], t, avg, pick(t, ["Professional"], ["Management"]), "Occupation")
    headline(fig, "Degree holders and professionals are the strongest buyers",
             "Purchase rate by education and occupation. Dashed line = company average.")
    save(fig, "q1_02_education_occupation.png")


def chart_gap(df, avg):
    rows = []
    for col in ["Age Group", "Income Band", "Gender", "Marital Status", "Education Level", "Occupation",
                "Homeowner Status", "Household Dependents", "Vehicles Owned",
                "Distance from Nearest Installer", "Region"]:
        t = rate_table(df, col)
        for idx, r in t.iterrows():
            if r["customers"] >= 75:
                name = {"Homeowner Status": "Homeowner", "Household Dependents": "Dependents",
                        "Vehicles Owned": "Vehicles", "Distance from Nearest Installer": "Distance",
                        "Marital Status": "Marital", "Education Level": "Education",
                        "Income Band": "Income", "Age Group": "Age"}.get(col, col)
                rows.append((f"{name}: {idx}", r["rate"] - avg, int(r["customers"])))
    g = pd.DataFrame(rows, columns=["group", "gap", "n"]).sort_values("gap")
    g = pd.concat([g.head(9), g.tail(9)])
    fig, ax = plt.subplots(figsize=(10, 7))
    cols = [ORANGE if v < 0 else BLUE for v in g["gap"]]
    ax.barh(g["group"], g["gap"], color=cols, height=0.7, edgecolor="white", linewidth=2)
    for y, (v, n) in enumerate(zip(g["gap"], g["n"])):
        ax.text(v + (0.6 if v > 0 else -0.6), y, f"{v:+.0f} pts  (n={n})", va="center",
                ha="left" if v > 0 else "right", fontsize=9, color=INK)
    ax.axvline(0, color=INK2, linewidth=1)
    ax.set_xlim(-34, 22)
    ax.set_xticks([])
    ax.spines["bottom"].set_visible(False)
    ax.tick_params(axis="y", length=0)
    headline(fig, "Who buys and who does not: biggest gaps from the 48% average",
             "Percentage points above (blue) or below (orange) the company purchase rate. Groups with at least 75 customers.")
    save(fig, "q1_03_buyer_gap.png")


def chart_married_age(df):
    t = (pd.crosstab(df["Age Group"], df["Marital Status"], values=df["Purchase Flag"], aggfunc="mean") * 100)
    t = t.reindex(ORDERS["Age Group"])
    fig, ax = plt.subplots(figsize=(8, 4))
    ax.plot(t.index, t["Single"], color=GREY, linewidth=2.5, marker="o", markersize=8)
    ax.plot(t.index, t["Married"], color=ORANGE, linewidth=2.5, marker="o", markersize=8)
    for lab, col in [("Single", INK2), ("Married", ORANGE)]:
        for x, v in zip(t.index, t[lab]):
            other = t.loc[x, "Married" if lab == "Single" else "Single"]
            ax.text(x, v + (3 if v >= other else -6), f"{v:.0f}%", ha="center", fontsize=9, color=INK)
    ax.text(4.12, t["Single"].iloc[-1], "Single", va="center", color=INK2, fontweight="bold")
    ax.text(4.12, t["Married"].iloc[-1], "Married", va="center", color=ORANGE, fontweight="bold")
    ax.set_ylim(10, 75)
    ax.set_yticks([])
    ax.tick_params(axis="x", length=0)
    headline(fig, "Married customers over 40 are the largest group not buying",
             "Purchase rate by age group and marital status. 384 married customers aged 40+ buy at only 37%.")
    save(fig, "q1_04_married_by_age.png")


def chart_region(df, avg):
    t = rate_table(df, "Region").sort_values("rate")
    fig, axes = plt.subplots(1, 2, figsize=(12, 3.6), gridspec_kw={"wspace": 0.35})
    cols = pick(t, ["Pacific"], ["North America"])
    axes[0].barh(t.index, t["customers"], color=cols, height=0.6)
    for y, (n, b) in enumerate(zip(t["customers"], t["buyers"])):
        axes[0].text(n + 8, y, f"{n} customers\n{b} buyers", va="center", fontsize=10, color=INK)
    axes[0].set_xlim(0, 680)
    axes[0].set_title("Size of customer base", loc="left", fontsize=11, color=INK2)
    axes[1].barh(t.index, t["rate"], color=cols, height=0.6)
    for y, v in enumerate(t["rate"]):
        axes[1].text(v + 1, y, f"{v:.0f}%", va="center", fontsize=10, fontweight="bold", color=INK)
    axes[1].axvline(avg, color=INK2, linestyle="--", linewidth=1)
    axes[1].set_xlim(0, 75)
    axes[1].set_title("Purchase rate", loc="left", fontsize=11, color=INK2)
    for ax in axes:
        ax.set_xticks([])
        ax.spines["bottom"].set_visible(False)
        ax.tick_params(axis="y", length=0)
    headline(fig, "North America is our biggest market but converts the worst",
             "Half of all customers live in North America, yet only 43% buy, compared with 59% in the Pacific.")
    save(fig, "q1_05_region.png")


def chart_region_distance(df):
    t = (pd.crosstab(df["Region"], df["Distance from Nearest Installer"], values=df["Purchase Flag"],
                     aggfunc="mean") * 100)[ORDERS["Distance from Nearest Installer"]]
    n = pd.crosstab(df["Region"], df["Distance from Nearest Installer"])[ORDERS["Distance from Nearest Installer"]]
    t = t.reindex(["Pacific", "Europe", "North America"])
    n = n.reindex(t.index)
    fig, ax = plt.subplots(figsize=(10, 3.6))
    ax.imshow(t.values, cmap="Blues", vmin=0, vmax=90, aspect="auto")
    for i in range(t.shape[0]):
        for j in range(t.shape[1]):
            v = t.values[i, j]
            ax.text(j, i - 0.1, f"{v:.0f}%", ha="center", va="center", fontsize=11, fontweight="bold",
                    color="white" if v > 55 else INK)
            ax.text(j, i + 0.22, f"n={n.values[i, j]}", ha="center", va="center", fontsize=8,
                    color="white" if v > 55 else INK2)
    ax.set_xticks(range(t.shape[1]), t.columns)
    ax.set_yticks(range(t.shape[0]), t.index)
    ax.tick_params(length=0)
    for s in ax.spines.values():
        s.set_visible(False)
    headline(fig, "In North America, customers more than 5 km from an installer rarely buy",
             "Purchase rate by region and distance to the nearest installer (darker = higher). n = number of customers.")
    save(fig, "q1_06_region_distance.png")


def chart_influences(df, avg):
    fig, axes = plt.subplots(1, 3, figsize=(13, 4), gridspec_kw={"width_ratios": [2, 2, 6]})
    t = rate_table(df, "Marital Status")
    bar_panel(axes[0], t, avg, [ORANGE, BLUE], "Marital status")
    t = rate_table(df, "Homeowner Status")
    t.index = ["Renter", "Owner"]
    bar_panel(axes[1], t, avg, [GREY, GREY], "Homeowner")
    t = rate_table(df, "Household Dependents")
    bar_panel(axes[2], t, avg, [GREY, BLUE, GREY, GREY, GREY, ORANGE], "Number of household dependents")
    headline(fig, "Marital status and family size affect buying; owning a home does not",
             "Purchase rate by household factor. Dashed line = company average.")
    save(fig, "q1_07_household_influences.png")


def chart_dependents_marital(df):
    t = (pd.crosstab(df["Household Dependents"], df["Marital Status"], values=df["Purchase Flag"],
                     aggfunc="mean") * 100)
    fig, ax = plt.subplots(figsize=(8, 4))
    ax.plot(t.index, t["Single"], color=GREY, linewidth=2.5, marker="o", markersize=8)
    ax.plot(t.index, t["Married"], color=ORANGE, linewidth=2.5, marker="o", markersize=8)
    for lab in ["Single", "Married"]:
        for x, v in zip(t.index, t[lab]):
            other = t.loc[x, "Married" if lab == "Single" else "Single"]
            ax.text(x, v + (3 if v >= other else -6), f"{v:.0f}%", ha="center", fontsize=9, color=INK)
    ax.text(5.15, t["Single"].iloc[-1], "Single", va="center", color=INK2, fontweight="bold")
    ax.text(5.15, t["Married"].iloc[-1], "Married", va="center", color=ORANGE, fontweight="bold")
    ax.set_xticks(range(6))
    ax.set_xlabel("Number of household dependents")
    ax.set_ylim(5, 75)
    ax.set_yticks([])
    ax.tick_params(axis="x", length=0)
    headline(fig, "Married families with 2 or more children are an untapped segment",
             "Purchase rate by dependents and marital status. 314 married customers with 2+ dependents buy at 34%.")
    save(fig, "q1_08_dependents_by_marital.png")


def chart_vehicles(df, avg):
    t = rate_table(df, "Vehicles Owned")
    fig, ax = plt.subplots(figsize=(7, 3.8))
    bar_panel(ax, t, avg, [BLUE, BLUE, ORANGE, ORANGE, ORANGE])
    ax.set_xlabel("Vehicles owned")
    headline(fig, "Households with 2 or more vehicles buy far less often",
             "Purchase rate by vehicles owned. 486 customers own 2+ vehicles and only 37% of them bought.")
    save(fig, "q1_09_vehicles.png")


def chart_distance(df, avg):
    t = rate_table(df, "Distance from Nearest Installer")
    fig, ax = plt.subplots(figsize=(8, 4))
    bar_panel(ax, t, avg, [BLUE, BLUE, BLUE, ORANGE, ORANGE])
    near = df[df["Installer Distance Order"] <= 3]["Purchase Flag"].mean() * 100
    far = df[df["Installer Distance Order"] > 3]["Purchase Flag"].mean() * 100
    ax.annotate(f"Within 5 km: {near:.0f}%", xy=(1, 66), ha="center", fontsize=10, color=BLUE, fontweight="bold")
    ax.annotate(f"Beyond 5 km: {far:.0f}%", xy=(3.5, 66), ha="center", fontsize=10, color=ORANGE, fontweight="bold")
    ax.set_xlabel("Distance from nearest installer")
    headline(fig, "Purchases drop sharply once customers live more than 5 km from an installer",
             "Purchase rate by distance. Dashed line = company average.")
    save(fig, "q1_10_distance.png")


def chart_distance_region(df):
    df = df.assign(Band=df["Installer Distance Order"].le(3).map({True: "Within 5 km", False: "Beyond 5 km"}))
    t = (pd.crosstab(df["Region"], df["Band"], values=df["Purchase Flag"], aggfunc="mean") * 100)
    n = pd.crosstab(df["Region"], df["Band"])
    order = ["Pacific", "Europe", "North America"]
    t, n = t.reindex(order), n.reindex(order)
    fig, ax = plt.subplots(figsize=(9, 4))
    x = range(len(order))
    w = 0.36
    b1 = ax.bar([i - w / 2 for i in x], t["Within 5 km"], w, color=BLUE, edgecolor="white", linewidth=2, label="Within 5 km")
    b2 = ax.bar([i + w / 2 for i in x], t["Beyond 5 km"], w, color=ORANGE, edgecolor="white", linewidth=2, label="Beyond 5 km")
    for bars, col in [(b1, "Within 5 km"), (b2, "Beyond 5 km")]:
        for b, v, k in zip(bars, t[col], n[col]):
            ax.text(b.get_x() + b.get_width() / 2, v + 1.2, f"{v:.0f}%", ha="center", fontsize=10, fontweight="bold")
            ax.text(b.get_x() + b.get_width() / 2, 2, f"n={k}", ha="center", fontsize=8, color="white")
    ax.set_xticks(list(x), order)
    ax.set_ylim(0, 80)
    ax.set_yticks([])
    ax.tick_params(axis="x", length=0)
    ax.legend(frameon=False, loc="upper right", ncol=2)
    headline(fig, "Distance hurts most in North America, where 171 customers live beyond 5 km",
             "Purchase rate within vs beyond 5 km of an installer, by region.")
    save(fig, "q1_11_distance_by_region.png")


def summary_tables(df):
    out = ROOT / "analysis" / "summary_tables.xlsx"
    with pd.ExcelWriter(out) as xw:
        for col in ["Age Group", "Income Band", "Gender", "Education Level", "Occupation", "Marital Status",
                    "Homeowner Status", "Household Dependents", "Vehicles Owned",
                    "Distance from Nearest Installer", "Region"]:
            t = rate_table(df, col).round(1)
            t.to_excel(xw, sheet_name=col[:31])
    print("saved", out.name)


def main():
    df = load_clean()
    avg = df["Purchase Flag"].mean() * 100
    print(f"Overall purchase rate {avg:.1f}%")
    chart_demographics(df, avg)
    chart_education_occupation(df, avg)
    chart_gap(df, avg)
    chart_married_age(df)
    chart_region(df, avg)
    chart_region_distance(df)
    chart_influences(df, avg)
    chart_dependents_marital(df)
    chart_vehicles(df, avg)
    chart_distance(df, avg)
    chart_distance_region(df)
    summary_tables(df)


if __name__ == "__main__":
    main()
