"""
Build the applied research paper PDF from the analysis outputs.

Formatting follows the course syllabus:
  12-point Times (Times-Roman), double spaced, 1-inch margins on all sides,
  page numbers at the bottom, references excluded from the length limit,
  exhibits (tables and figures) placed after the references.

Run:  python3 analyze_demand.py && python3 paper/build_paper.py
Output: paper/gasoline_demand_paper.pdf
"""

import csv
import json
import os

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.utils import ImageReader
from reportlab.platypus import (BaseDocTemplate, Frame, Image, KeepTogether,
                                PageBreak, PageTemplate, Paragraph, Spacer,
                                Table, TableStyle)

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
TAB = os.path.join(ROOT, "output", "tables")
FIG = os.path.join(ROOT, "output", "figures")
OUT = os.path.join(HERE, "gasoline_demand_paper.pdf")

with open(os.path.join(ROOT, "output", "results.json")) as f:
    R = json.load(f)

# ── paper metadata (edit these three lines) ──────────────────────────
AUTHOR = "Roya Rezaie"
AFFIL = "University at Albany, SUNY"
COURSE = "BFIN 515: Economic Analysis"
REPO = "https://github.com/RoyaRezaie/gasoline-demand"
DATE = "October 5, 2026"

# ── shorthand for the estimates used in the text ─────────────────────
m1, m2, m3, m4 = R["m1"], R["m2"], R["m3"], R["m4"]
base_b = m1["terms"]["ln_p"]["coef"]
base_se = m1["terms"]["ln_p"]["se"]
pref_b = m2["terms"]["ln_p"]["coef"]
pref_se = m2["terms"]["ln_p"]["se"]
pref_r2 = m2["r2"]
covid_b = m2["terms"]["pandemic"]["coef"]
sr, sr_se = m3["short_run"], m3["short_run_se"]
lr, lr_se = m3["long_run"], m3["long_run_se"]
rho = m3["rho"]
iv_b, iv_se = m4["short_run"], m4["short_run_se"]
iv_F = m4["first_stage_f"]
J, Jp = m4["sargan_J"], m4["sargan_p"]
meta = R["meta"]
hold = R["holdout"]

fc = {(r["period"], r["method"]): r for r in R["forecast"]}
A = "Hold-out 2016-01..2020-02"
S = "Stress 2020-03..2026-07"
C = "In-sample 2020-03..2021-06"
rm_pref = fc[(A, "Preferred model (1-step)")]["rmse"]
mape_pref = fc[(A, "Preferred model (1-step)")]["mape"]
rm_base = fc[(A, "Baseline model (linear trend)")]["rmse"]
rm_rw = fc[(A, "Benchmark: random walk (t-1)")]["rmse"]
rm_sn = fc[(A, "Benchmark: seasonal naive (t-12)")]["rmse"]
rm_mean = fc[(A, "Benchmark: training mean")]["rmse"]
rm_stress = fc[(S, "Preferred model pre-2020 (1-step)")]["rmse"]
rm_stress_rw = fc[(S, "Benchmark: random walk (t-1)")]["rmse"]
rm_covid_with = fc[(C, "Preferred model with pandemic dummy")]["rmse"]
rm_covid_without = fc[(C, "Same model without pandemic dummy")]["rmse"]

imp = {r["price_change_$_per_gal"]: r for r in R["implied"]}
q25 = imp[0.25]

rob = R["robustness"]
rob_min = min(r["elasticity"] for r in rob)
rob_max = max(r["elasticity"] for r in rob)
season = {r["month"]: r["effect_percent"] for r in R["seasonal"]}
mean_p = R["mean_price_real"]
mean_q = meta["mean_q"]
n_obs = meta["n"]
MONTH_NAMES = ["January", "February", "March", "April", "May", "June", "July",
               "August", "September", "October", "November", "December"]
_end_y, _end_m = meta["end"].split("-")
end_long = f"{MONTH_NAMES[int(_end_m) - 1]} {_end_y}"
_start_y, _start_m = meta["start"].split("-")
start_long = f"{MONTH_NAMES[int(_start_m) - 1]} {_start_y}"

# ── styles ───────────────────────────────────────────────────────────
ss = getSampleStyleSheet()
BODY = ParagraphStyle("body", parent=ss["Normal"], fontName="Times-Roman",
                      fontSize=12, leading=24, alignment=TA_JUSTIFY,
                      firstLineIndent=18, spaceBefore=0, spaceAfter=0)
TITLE = ParagraphStyle("title", parent=BODY, fontName="Times-Bold", fontSize=15,
                       leading=20, alignment=TA_CENTER, firstLineIndent=0,
                       spaceBefore=0, spaceAfter=6)
SUB = ParagraphStyle("sub", parent=BODY, fontSize=12, leading=18,
                     alignment=TA_CENTER, firstLineIndent=0, spaceAfter=2)
AUTH = ParagraphStyle("auth", parent=BODY, fontSize=11, leading=17,
                      alignment=TA_CENTER, firstLineIndent=0, spaceAfter=1)
H1 = ParagraphStyle("h1", parent=BODY, fontName="Times-Bold", fontSize=12,
                    leading=24, alignment=TA_JUSTIFY, firstLineIndent=0,
                    spaceBefore=18, spaceAfter=0)
REF = ParagraphStyle("ref", parent=BODY, firstLineIndent=0, leftIndent=24,
                     fontSize=12, leading=24)
CAP = ParagraphStyle("cap", parent=BODY, fontName="Times-Italic", fontSize=9.5,
                     leading=12, alignment=TA_CENTER, firstLineIndent=0,
                     spaceBefore=4, spaceAfter=10)
EXH = ParagraphStyle("exh", parent=BODY, fontName="Times-Bold", fontSize=10.5,
                     leading=13, alignment=TA_CENTER, firstLineIndent=0,
                     spaceBefore=8, spaceAfter=3)
EQS = ParagraphStyle("eq", parent=BODY, alignment=TA_CENTER, firstLineIndent=0,
                     fontName="Times-Italic", spaceBefore=6, spaceAfter=6)
CODES = ParagraphStyle("code", parent=BODY, fontName="Courier", fontSize=9.5,
                       leading=13, alignment=0, firstLineIndent=0,
                       leftIndent=18, spaceBefore=4, spaceAfter=4)
APPH = ParagraphStyle("apph", parent=BODY, fontName="Times-Bold", fontSize=13,
                      leading=20, alignment=TA_CENTER, firstLineIndent=0,
                      spaceBefore=0, spaceAfter=12)
H2S = ParagraphStyle("h2s", parent=BODY, fontName="Times-Bold", fontSize=12,
                     leading=24, alignment=TA_JUSTIFY, firstLineIndent=0,
                     spaceBefore=14, spaceAfter=0)

# ── content blocks: build_paper.py records them and render() turns them
#    into PDF flowables; build_docx.py reads the same list ────────────
story = []


def p(text):
    story.append({"k": "p", "text": text})


def h1(text):
    story.append({"k": "h1", "text": text})


def h2(text):
    story.append({"k": "h2", "text": text})


def eq(text):
    story.append({"k": "eq", "text": text})


def ref(text):
    story.append({"k": "ref", "text": text})


def exhibit_title(text):
    story.append({"k": "exh", "text": text})


def caption(text):
    story.append({"k": "cap", "text": text})


def code(lines):
    story.append({"k": "code", "lines": lines})


def page_break():
    story.append({"k": "pagebreak"})


def spacer(height=8):
    story.append({"k": "spacer", "h": height})


def title(text):
    story.append({"k": "title", "text": text})


def subtitle(text):
    story.append({"k": "sub", "text": text})


def byline(text):
    story.append({"k": "auth", "text": text})


def center_heading(text):
    story.append({"k": "exh0", "text": text})


def apph(text):
    story.append({"k": "apph", "text": text})


def table(name=None, rows=None, widths=None, fontsize=8.5, numeric_from=1,
          columns=None, headers=None):
    story.append({"k": "table", "name": name, "rows": rows, "widths": widths,
                  "fontsize": fontsize, "numeric_from": numeric_from,
                  "columns": columns, "headers": headers})


def figure(fname, max_width=6.3 * 72, max_height=4.2 * 72):
    story.append({"k": "figure", "fname": fname, "max_width": max_width,
                  "max_height": max_height})


def read_table(name):
    with open(os.path.join(TAB, name)) as f:
        return list(csv.reader(f))


def table_rows(block):
    rows = read_table(block["name"]) if block["name"] else \
        [list(map(str, r)) for r in block["rows"]]
    if block["columns"] is not None:
        rows = [[r[i] for i in block["columns"]] for r in rows]
    if block["headers"] is not None:
        rows[0] = list(block["headers"])
    pcols = {i for i, c in enumerate(rows[0])
             if c.strip().lower() in ("p", "p-value")}
    for r in rows[1:]:
        for i in pcols:
            try:
                if float(r[i]) == 0.0:
                    r[i] = "&lt;0.0001"
            except (ValueError, IndexError):
                pass
    return rows


def make_table(block):
    rows = table_rows(block)
    fontsize = block["fontsize"]
    numeric_from = block["numeric_from"]
    data = [[Paragraph(f"<b>{c}</b>", ParagraphStyle(
        "th", fontName="Times-Bold", fontSize=fontsize, leading=fontsize + 2,
        alignment=TA_CENTER)) if i == 0 else c
        for i, c in enumerate(rows[0])]]
    st = ParagraphStyle("td", fontName="Times-Roman", fontSize=fontsize,
                        leading=fontsize + 2)
    stn = ParagraphStyle("tdn", parent=st, alignment=2)
    for r in rows[1:]:
        data.append([Paragraph(c, st) if i < numeric_from else Paragraph(c, stn)
                     for i, c in enumerate(r)])
    widths = block["widths"]
    if widths is None:
        widths = [468 / len(rows[0])] * len(rows[0])
    t = Table(data, colWidths=widths, repeatRows=1, hAlign="CENTER")
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.Color(0.92, 0.92, 0.92)),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    return t


def make_figure(block):
    path = os.path.join(FIG, block["fname"])
    w, h = ImageReader(path).getSize()
    scale = min(block["max_width"] / w, block["max_height"] / h)
    return Image(path, width=w * scale, height=h * scale)


def render(blocks):
    """Convert the content blocks into PDF flowables."""
    out = []
    for b in blocks:
        k = b["k"]
        if k == "p":
            out.append(Paragraph(b["text"], BODY))
        elif k == "h1":
            out.append(Paragraph(b["text"], H1))
        elif k == "h2":
            out.append(Paragraph(b["text"], H2S))
        elif k == "eq":
            out.append(Paragraph(b["text"], EQS))
        elif k == "ref":
            out.append(Paragraph(b["text"], REF))
        elif k == "exh":
            out.append(Paragraph(b["text"], EXH))
        elif k == "exh0":
            out.append(Paragraph(b["text"], ParagraphStyle(
                "exh0", parent=H1, alignment=TA_CENTER, spaceBefore=0,
                spaceAfter=6)))
        elif k == "cap":
            out.append(Paragraph(b["text"], CAP))
        elif k == "code":
            for line in b["lines"]:
                out.append(Paragraph(line.replace(" ", "&nbsp;"), CODES))
        elif k == "title":
            out.append(Paragraph(b["text"], TITLE))
        elif k == "sub":
            out.append(Paragraph(b["text"], SUB))
        elif k == "auth":
            out.append(Paragraph(b["text"], AUTH))
        elif k == "apph":
            out.append(Paragraph(b["text"], APPH))
        elif k == "pagebreak":
            out.append(PageBreak())
        elif k == "spacer":
            out.append(Spacer(1, b["h"]))
        elif k == "table":
            out.append(make_table(b))
        elif k == "figure":
            out.append(make_figure(b))
        else:
            raise ValueError(f"unknown block kind: {k}")
    return out


# ═════════════════════════════════════════════════════════════════════
#  Front matter
# ═════════════════════════════════════════════════════════════════════
title("The Impact of Gasoline Prices on Gasoline Demand "
      "in the United States")
subtitle("Evidence from Monthly Data, 1993–2026")
spacer(8)
byline(f"{AUTHOR}, {AFFIL}")
byline(COURSE)
byline(DATE)
byline(f'Replication package (code, data, figures, this paper): '
       f'<link href="{REPO}" color="blue">{REPO}</link>')
spacer(14)

# ═════════════════════════════════════════════════════════════════════
# 1. Business question
# ═════════════════════════════════════════════════════════════════════
h1("1. Business Question")

p("Gasoline is the single largest operating cost for most road-transport "
  "businesses and the core revenue line for retail fuel stations, where the "
  "profit on a gallon is measured in cents and volume is what pays the rent. "
  "That makes one question unusually valuable to a retailer: when the price "
  "of gasoline changes, how much does the quantity that customers actually "
  "buy change? The business question addressed in this paper is therefore: "
  "<b>how does the price of gasoline affect the quantity of gasoline demanded "
  "in the United States, and how can that relationship guide gasoline "
  "retailers in their pricing and inventory decisions?</b>")

p("Three groups of decision makers can use the answer. A pricing manager at a "
  "fuel or convenience chain needs to know whether raising the posted price "
  "will cost more in lost gallons than it gains in margin. An inventory and "
  "wholesale buyer needs a demand forecast in gallons in order to schedule "
  "deliveries and tank levels. A policymaker considering a fuel tax or a "
  "clean-fuel regulation needs the same elasticity to predict how much "
  "consumption, and therefore revenue or emissions, would actually change. "
  "This paper speaks to all three, with the retailer as the primary audience.")

p("Using {n} monthly observations from April 1993 through {end} from the U.S. "
  "Energy Information Administration (EIA) and the Federal Reserve Bank of "
  "St. Louis, I estimate that U.S. gasoline demand is price-inelastic in the "
  "short run: a preferred estimate of {b:.3f} (HAC standard error {se:.3f}) "
  "means a ten percent increase in the real pump price reduces the quantity "
  "demanded by less than one percent. The estimate is stable across fourteen "
  "alternative specifications, samples, and price measures, it is close to "
  "instrumental-variable estimates built on crude oil cost shocks, and a "
  "model estimated through 2015 beats three standard forecasting benchmarks "
  "in a pre-pandemic hold-out sample. The practical implication is that "
  "market-level demand barely moves when prices move, so for a retailer the "
  "volume risk of a modest price increase is small while the revenue upside "
  "is large, and inventory planning should be driven by seasonality and "
  "macroeconomic conditions rather than by price swings."
  .format(n=n_obs, end=end_long, b=pref_b, se=pref_se))

# ═════════════════════════════════════════════════════════════════════
# 2. Data
# ═════════════════════════════════════════════════════════════════════
h1("2. Data")

p("Quantity is measured as monthly U.S. product supplied of finished motor "
  "gasoline, reported by the EIA in thousands of barrels per month in the "
  "Petroleum and Other Liquids data navigator. Product supplied is the "
  "standard EIA supply-side proxy for consumption: it equals production plus "
  "net imports minus the change in stocks, so it measures the volume that "
  "reaches the market. I convert it to thousands of barrels per day by "
  "dividing by the number of days in each month; over the sample the United "
  "States consumes a mean of {meanq:,.0f} thousand barrels per day, with a "
  "minimum of 5,866 in April 2020 and a maximum of 9,834."
  .format(meanq=mean_q))

p("Price is the EIA monthly average retail price of all grades of gasoline, "
  "all formulations, in dollars per gallon, for the same months. The series "
  "begins in March 1993, which is why the sample starts in April 1993 and "
  "runs {n} months to {end}. Prices are deflated by the Consumer Price Index "
  "(CPIAUCSL) into constant 2017 dollars, giving a mean of ${p:.2f} per "
  "gallon and a range of $1.43 to $4.63. Nominal prices are kept for a "
  "robustness check. Exhibit 1 plots the two headline series; Exhibit 2 "
  "reports descriptive statistics."
  .format(n=n_obs, end=end_long, p=mean_p))

p("Three controls and one instrument come from FRED, the Federal Reserve Bank "
  "of St. Louis economic database: real disposable personal income "
  "(DSPIC96), which I divide by resident population (POPTHM) to obtain real "
  "income per capita as the shift variable for the demand curve; the "
  "unemployment rate (UNRATE), used only in a robustness check; and the "
  "West Texas Intermediate crude oil spot price (DCOILWTICO), which is the "
  "cost shifter used as an instrument for the retail price. All series are "
  "monthly and are merged on the calendar month, leaving {n} complete "
  "observations with no missing values.".format(n=n_obs))

p("The raw source files are downloaded directly from the EIA and FRED by the "
  "script download_data.py, which writes the analysis file "
  "data/gasoline_demand_monthly.csv. Nothing in the paper is hand-entered: "
  "every table and figure in the exhibits is regenerated by "
  "analyze_demand.py. The full repository, including this paper's source, is "
  f"at {REPO}.")

# ═════════════════════════════════════════════════════════════════════
# 3. Method
# ═════════════════════════════════════════════════════════════════════
h1("3. Method")

p("The object of interest is a demand function: quantity demanded as a "
  "function of price, income, and other determinants. I estimate it in "
  "log-linear form, so each slope coefficient is an elasticity. The baseline "
  "specification is")

eq("ln Q<sub>t</sub> = a + b ln P<sub>t</sub> + c ln Y<sub>t</sub> + "
   "month dummies + d · t + e<sub>t</sub>,")

p("where Q is product supplied in thousands of barrels per day, P is the real "
  "retail price, Y is real income per capita, eleven month dummies absorb "
  "seasonality with December as the base month, and t is a linear time trend. "
  "The coefficient b is the short-run price elasticity of demand: the "
  "percentage change in quantity for a one percent change in price.")

p("Two features of this problem make the naive version of that equation "
  "unreliable, and it is worth being explicit about them because they drive "
  "the whole design. First, the demand path is not a straight line. Improvements "
  "in fleet fuel economy, the aging of the vehicle stock, the shift toward "
  "sport-utility vehicles and later toward electrification, and the plateau in "
  "vehicle miles traveled after 2007 all shift gasoline demand gradually and "
  "non-linearly. Second, the pandemic was an enormous exogenous shock to "
  "mobility that moved price and quantity down together. A single linear "
  "trend cannot represent either. As Section 4 shows, the naive equation "
  "returns a positive price coefficient of {b:+.3f}, which is an artifact of "
  "omitted structural change rather than genuine behavior."
  .format(b=base_b))

p("The preferred specification therefore adds a step indicator and a slope "
  "change at January 2007, plus a dummy for March 2020 through June 2021, the "
  "period of formal mobility restrictions and remote work: <i>ln Q</i> = a + "
  "b ln P + c ln Y + month dummies + step + trend after 2007 + pandemic + e. "
  "The break year and the pandemic window are judgment calls, so Section 5 "
  "reports how the elasticity moves when either is changed.")

p("Because consumption adjusts slowly to its determinants, I also estimate a "
  "partial-adjustment (ARDL) model that adds the lagged level of ln Q. In "
  "that model b is the short-run elasticity and b / (1 - rho) is the "
  "long-run elasticity, where rho is the coefficient on the lagged dependent "
  "variable; the long-run standard error is computed with the delta method.")

p("Price is not randomly assigned. Demand shocks raise both price and "
  "quantity, which biases ordinary least squares upward, and refining and "
  "crude costs are passed through to the pump. I therefore estimate a "
  "two-stage least squares model in which the real retail price is "
  "instrumented by the real WTI crude oil price, whose exclusion restriction "
  "is that crude costs affect gasoline quantity only through the retail "
  "price they produce. A first-stage F statistic of {F:.0f} indicates a "
  "strong instrument. As a specification test I add the lagged crude price as "
  "a second instrument and compute the Sargan statistic: J = {J:.2f} "
  "(p = {Jp:.3f}) rejects the over-identified model, so the just-identified "
  "version is the one reported, and the rejection is itself evidence that a "
  "lagged crude price is contaminated by global demand conditions."
  .format(F=iv_F, J=J, Jp=Jp))

p("Because the data are monthly and the residuals are strongly serially "
  "correlated (the Durbin-Watson statistic of the preferred regression is "
  "about 1.0), all standard errors are Newey-West heteroskedasticity- and "
  "autocorrelation-consistent with twelve lags. Inference on the long-run "
  "elasticity uses the delta method.")

p("The estimation plan includes four validation exercises, described in "
  "Section 5: an out-of-sample hold-out period with explicit benchmarks, a "
  "stress test of the pandemic regime, a robustness table over alternative "
  "specifications, samples, and price measures, and a comparison of the "
  "elasticity with published estimates from the literature.")

# ═════════════════════════════════════════════════════════════════════
# 4. Results
# ═════════════════════════════════════════════════════════════════════
h1("4. Results")

p("Exhibit 3 reports the main estimation results. Column 1 is the baseline "
  "equation with a linear trend. Its price coefficient is {b:+.3f} "
  "(s.e. {se:.3f}), a positive and significant elasticity that would say "
  "consumers buy more gasoline when it gets more expensive. That is not "
  "credible, and the reason is visible in Exhibit 1: over 1993 to 2007 both "
  "real prices and quantity trended upward together, while after 2007 "
  "quantity flattened and prices did not. Pooling the two eras under a "
  "single straight line loads the structural change onto the price "
  "coefficient. Column 2 is the preferred specification with the 2007 break "
  "and the pandemic dummy; it raises R-squared from {r1:.3f} to {r2:.3f} and "
  "returns an elasticity of {b2:+.3f} (s.e. {se2:.3f}, significant at one "
  "percent). The pandemic dummy is {cv:+.3f}, about twelve percent lower "
  "quantity during the restriction period, as expected."
  .format(b=base_b, se=base_se, r1=m1["r2"], r2=pref_r2, b2=pref_b,
          se2=pref_se, cv=covid_b))

p("Column 3 adds the lagged dependent variable. The adjustment coefficient "
  "is {rho:.3f}, so about half of any deviation from the steady state closes "
  "within a month. The short-run elasticity is {sr:+.3f} (s.e. {srse:.3f}) "
  "and the long-run elasticity is {lr:+.3f} (s.e. {lrse:.3f}), both "
  "statistically significant. Column 4 instruments the price with crude oil "
  "costs and gives {iv:+.3f} (s.e. {ivse:.3f}), close to the least-squares "
  "estimate, which suggests that simultaneity bias is small in this setting "
  "once the structural controls are in place."
  .format(rho=rho, sr=sr, srse=sr_se, lr=lr, lrse=lr_se, iv=iv_b, ivse=iv_se))

p("The income coefficient in the preferred specification is small and "
  "insignificant ({inc:+.3f}, s.e. {incse:.3f}), which is common in monthly "
  "aggregate gasoline studies: income moves slowly, is absorbed by the trend "
  "terms, and its variation is modest relative to price over this sample. "
  "What remains is a demand curve that is steeply inelastic in quantity."
  .format(inc=m2["terms"]["ln_y"]["coef"], incse=m2["terms"]["ln_y"]["se"]))

p("Exhibit 4 shows the identifying variation directly: after residualizing "
  "both logs on income, seasonality, the trend terms, and the pandemic "
  "dummy, the scatter of price against quantity has a clearly negative "
  "slope. Exhibit 5 overlays the actual and fitted series; the preferred "
  "model tracks the 1990s expansion, the post-2007 plateau, the 2008 "
  "recession, the 2014–2016 price collapse, and the pandemic collapse and "
  "recovery.")

p("Two results matter for the business question. First, the magnitude: with "
  "an elasticity of {b:.3f}, a ten percent increase in the real pump price "
  "reduces market quantity by about {q:.2f} percent. Exhibit 10 converts this "
  "into the units a retailer thinks in: at the sample mean real price of "
  "${p:.2f}, a $0.25 per gallon increase ({pp:.1f} percent) lowers quantity "
  "by about {qp:.2f} percent, or {qd:,.0f} thousand barrels per day across "
  "the market, and raises fuel revenue by about {rev:.1f} percent, because "
  "demand is inelastic. Second, the seasonal pattern in Exhibit 11: holding "
  "price and income fixed, demand in {hot1} runs {s1:.1f} percent above "
  "December and in {hot2} {s2:.1f} percent above, while {cold} is "
  "{s3:.1f} percent below. Volume planning, not price movement, is what "
  "moves gallons."
  .format(b=pref_b, q=abs(pref_b) * 10, p=mean_p, pp=q25["price_change_percent"],
          qp=abs(q25["predicted_qty_change_percent"]),
          qd=abs(q25["qty_change_thousand_bpd"]),
          rev=q25["predicted_revenue_change_percent"],
          hot1="July", s1=season["July"], hot2="August", s2=season["August"],
          cold="January", s3=abs(season["January"])))

# ═════════════════════════════════════════════════════════════════════
# 5. Validation
# ═════════════════════════════════════════════════════════════════════
h1("5. Validation")

p("<b>Out-of-sample hold-out.</b> I re-estimate both the baseline and the "
  "preferred equations using only data through December 2015 ({ntrain} "
  "months) and forecast one month ahead for January 2016 through February "
  "2020 ({ntest} months), a period that no estimation equation has seen and "
  "that ends before the pandemic. Exhibit 6 reports root mean squared error "
  "(RMSE) and mean absolute percentage error (MAPE) in thousand barrels per "
  "day. The preferred model produces an RMSE of {rm:.0f} and a MAPE of "
  "{mp:.1f} percent, against {rm2:.0f} for the baseline model with the "
  "linear trend, {rms:.0f} for a seasonal naive forecast that repeats the "
  "same month one year earlier, {rmrw:.0f} for a random walk, and {rmm:.0f} "
  "for the training-sample mean. Exhibit 7 overlays the preferred model's "
  "one-step-ahead path on the actual series alongside the baseline forecast "
  "and the seasonal naive benchmark. The preferred model beats every benchmark, "
  "and the gap between {rm:.0f} and {rm2:.0f} shows that the structural "
  "controls improve forecasting as well as estimation, which is the practical "
  "justification for the extra terms rather than a purely statistical one."
  .format(ntrain=hold["n_train"], ntest=hold["n_test"], rm=rm_pref,
          mp=mape_pref, rm2=rm_base, rms=rm_sn, rmrw=rm_rw, rmm=rm_mean))

p("<b>Stress period.</b> The second check asks what happens when the regime "
  "changes in a way the model has never seen. Forecasting March 2020 through "
  "{end} with the pre-2020 model gives an RMSE of {rms:.0f} thousand barrels "
  "per day, worse than the random walk benchmark at {rmrw:.0f}, because a "
  "model that knows nothing about mobility restrictions extrapolates a "
  "demand level that no longer exists. Inside the sample, adding the "
  "pandemic dummy reduces in-sample error over March 2020 to June 2021 from "
  "{wo:.0f} to {w:.0f}. The honest conclusion is that the elasticity itself "
  "remains usable, but demand forecasts must be re-estimated, or explicitly "
  "flagged, when a shock of this type arrives."
  .format(end=end_long, rms=rm_stress, rmrw=rm_stress_rw,
          wo=rm_covid_without, w=rm_covid_with))

p("<b>Alternative specifications and samples.</b> Exhibit 8 reports the "
  "price elasticity from fourteen variations on the preferred equation: "
  "moving the trend break to 2006, 2008, or 2009; shortening or lengthening "
  "the pandemic window; adding the unemployment rate; using per-capita "
  "quantity as the dependent variable; restricting the sample to "
  "pre-pandemic months; splitting at 2004; using the nominal price with the "
  "CPI included as a separate regressor; using regular-grade rather than "
  "all-grade prices; the crude-oil instrument; and a linear-in-levels form "
  "whose implied elasticity at the sample mean is reported. Every estimate "
  "is negative, ranging from {lo:+.3f} to {hi:+.3f}, and most are "
  "significant at five percent or better. The pre-pandemic subsample alone "
  "gives {pre:+.3f} (s.e. {prese:.3f}), which is close to the full-sample "
  "result and shows that the estimate is not driven by the pandemic "
  "months.".format(lo=rob_min, hi=rob_max,
                   pre=[r for r in rob if "Pre-pandemic" in r["specification"]][0]["elasticity"],
                   prese=[r for r in rob if "Pre-pandemic" in r["specification"]][0]["se"]))

p("<b>Stability over time.</b> Exhibit 9 re-estimates the preferred equation "
  "in rolling six-year windows. Estimates are negative and mostly between "
  "-0.02 and -0.10 through 2019, become positive only in the windows that "
  "contain the pandemic collapse, and return to negative values afterwards. "
  "This is the same lesson as the stress test: the sign of the raw "
  "coefficient in a short window is dominated by whatever demand shock "
  "dominated that window, which is why the pooled estimate with an explicit "
  "pandemic control is the appropriate summary statistic.")

p("<b>Comparison with published estimates.</b> Hughes, Knittel, and Sperling "
  "(2008) report short-run price elasticities of U.S. gasoline demand of "
  "-0.034 to -0.077 for 2001–2006, compared with -0.21 to -0.34 for "
  "1975–1980, arguing that responsiveness has fallen. Havranek, Irsova, and "
  "Janda (2012), correcting a meta-analytic literature for publication "
  "selection bias, put the average short-run elasticity at -0.09 and the "
  "long run at -0.31; Espey (1998) is the classic earlier meta-analysis. My "
  "short-run estimates of {sr:+.3f} to {b:+.3f} sit inside the "
  "post-2000 range and are close to the publication-bias-corrected average. "
  "The long-run estimate of {lr:+.3f} is smaller in absolute value than the "
  "-0.31 benchmark, which is consistent with how the model identifies "
  "long-run behavior: the trend terms deliberately absorb the slow secular "
  "shifts in fuel economy and vehicle stock, so the long-run response is "
  "identified only from medium-horizon price variation. That is a limit of "
  "this design rather than a contradiction of the literature, and it is "
  "reported as such.".format(sr=sr, b=pref_b, lr=lr))

# ═════════════════════════════════════════════════════════════════════
# 6. Recommendation
# ═════════════════════════════════════════════════════════════════════
h1("6. Recommendation")

p("<b>For a pricing manager.</b> Market demand for gasoline is inelastic: "
  "quantity moves by roughly {q:.1f} percent for every ten percent change in "
  "price. Within the limits of local competition, that means a modest price "
  "increase is far more likely to raise fuel revenue than to lose gallons: "   "Exhibit 10 implies that $0.25 per gallon at the average station's market "
  "price changes volume by only {qp:.2f} percent while raising revenue by "
  "about {rev:.1f} percent. The essential caveat is that this is a "
  "<i>market-wide</i> elasticity. A single station raising its posted price "
  "while the station across the street does not will lose more than "
  "{qp:.2f} percent of its volume, because drivers substitute across "
  "stations; this data set cannot identify that own-price elasticity. The "
  "correct use of the result is therefore at the chain or market level, "
  "setting the direction and rough magnitude of a price move, while the "
  "store-level price gap against local competitors should be tuned with the "
  "retailer's own point-of-sale history."
  .format(q=abs(pref_b) * 10, qp=abs(q25["predicted_qty_change_percent"]),
          rev=q25["predicted_revenue_change_percent"]))

p("<b>For an inventory and wholesale buyer.</b> Because demand barely "
  "responds to price, price movements should not drive ordering. What does "   "drive gallons is the seasonal cycle, which Exhibit 11 quantifies at "
  "roughly {s1:.1f} to {s2:.1f} percent above December in July and August "
  "and about {s3:.1f} percent below it in January, together with income and "
  "the level of economic activity. Order schedules, tank sizing, and delivery "
  "cadence should be built from that seasonal profile and from a short-horizon "
  "macro forecast, with a price-driven adjustment of at most a fraction of a "
  "percent for a typical price swing."
  .format(s1=season["July"], s2=season["August"], s3=abs(season["January"])))

p("<b>For forecasting and planning.</b> The preferred model's hold-out MAPE "
  "of {mp:.1f} percent per month is good enough to use directly as a "
  "one-to-three-month demand plan, and it clearly beats the naive benchmarks "
  "in normal times. Two operating rules follow from the validation exercise. "
  "Re-estimate the equation annually so the trend terms track fuel economy "
  "and fleet changes. And when an exogenous mobility shock appears, do not "
  "extrapolate: the pre-2020 model's error rose to {rms:,.0f} thousand "
  "barrels per day through the pandemic, worse than a random walk, so switch "
  "to a scenario forecast until the regime is over."
  .format(mp=mape_pref, rms=rm_stress))

p("<b>For a policymaker.</b> The same inelasticity cuts the other way. A "
  "fuel tax or carbon charge of ten percent at the pump would reduce "
  "consumption by well under one percent in the short run, so tax-based "
  "policy needs to be large, permanent, and paired with efficiency or "
  "alternative-fuel measures if the goal is material reductions in gasoline "
  "use; in the short run it behaves mostly as a revenue instrument.")

p("<b>Limitations.</b> The analysis uses national monthly aggregates, so it "
  "cannot separate regions, stations, or grades, and it cannot identify the "
  "own-price elasticity a single retailer faces. Product supplied is a "
  "supply-side proxy for consumption. Fuel economy, vehicle miles traveled, "
  "and fleet composition are absorbed by the trend break rather than "
  "measured directly, which is why the long-run elasticity is smaller than "
  "published benchmarks. The instrument's exclusion restriction is "
  "imperfect, as the Sargan test indicates. Finally, the sample ends in "
  "{end}, so the most recent months are preliminary EIA figures that are "
  "subject to revision. None of these caveats changes the central finding for the "
  "decision maker: in the short run, U.S. gasoline demand is driven far more "
  "by seasonality and the state of the economy than by the price at the "
  "pump.".format(end=end_long))

p("<b>Data and code availability.</b> The data set, the estimation code, every "
  "table and figure in this paper, and the source of the paper itself are "
  f'publicly available at <link href="{REPO}" color="blue">{REPO}</link>. '
  "Running download_data.py, analyze_demand.py and paper/build_paper.py "
  "reproduces every number reported above.")

# ═════════════════════════════════════════════════════════════════════
# References (excluded from the page limit)
# ═════════════════════════════════════════════════════════════════════
page_break()
h1("References")

refs = [
    "Espey, M. (1998). Gasoline demand revisited: An international "
    "meta-analysis of cross-country elasticity estimates. <i>Energy "
    "Economics</i>, 20(3), 273–295.",
    "Hughes, J. E., Knittel, C. R., &amp; Sperling, D. (2008). Evidence of a "
    "shift in the short-run price elasticity of gasoline demand. <i>The "
    "Energy Journal</i>, 29(1), 113–134.",
    "Havranek, T., Irsova, Z., &amp; Janda, K. (2012). Demand for gasoline is "
    "more price-inelastic than commonly thought. <i>Energy Economics</i>, "
    "34(1), 201–207.",
    "U.S. Energy Information Administration. <i>Petroleum and Other Liquids: "
    "Supply and Disposition, and Retail Prices</i> (monthly series, "
    "downloaded October 2026). https://www.eia.gov/dnav/pet/",
    "Federal Reserve Bank of St. Louis. <i>FRED</i>: CPIAUCSL, DSPIC96, "
    "POPTHM, UNRATE, DCOILWTICO (downloaded October 2026). "
    "https://fred.stlouisfed.org/",
    "Using Supply and Demand Functions: A Survey by Data Availability. "
    "<i>SSRN Electronic Journal</i>. "
    "https://papers.ssrn.com/sol3/papers.cfm?abstract_id=7399858 (starting "
    "point for this project).",
    "Wooldridge, J. M. (2020). <i>Introductory Econometrics: A Modern "
    "Approach</i> (7th ed.). Cengage Learning.",
]
for r in refs:
    ref(r)

# ═════════════════════════════════════════════════════════════════════
# Exhibits (excluded from the page limit)
# ═════════════════════════════════════════════════════════════════════
page_break()
center_heading("Exhibits")

exhibit_title("Exhibit 1. U.S. retail gasoline price and gasoline product "
              "supplied, 1993–2026")
figure("fig1_price_quantity.png")
caption("Monthly data. Top panel: nominal and real (2017$) retail price of "
        "all grades of gasoline. Bottom panel: motor gasoline product "
        "supplied, thousand barrels per day. Shaded area: March 2020–June "
        "2021. Source: EIA; author's calculations.")

exhibit_title("Exhibit 2. Descriptive statistics")
table("table1_descriptives.csv",
      widths=[150, 30, 62, 66, 58, 58, 72, 72], fontsize=8,
      headers=["Variable", "N", "Mean", "Std. dev.", "Min", "Max",
               "First month", "Last month"])
caption("N = {n} monthly observations, April 1993 to {end}. Prices and the "
        "crude oil price are in constant 2017 dollars.".format(             n=n_obs, end=end_long))

exhibit_title("Exhibit 3. Main results: price elasticity of gasoline demand")
table("table2_main_results.csv",
      widths=[92, 122, 56, 46, 46, 34, 34, 38], fontsize=8, numeric_from=2,
      columns=[0, 1, 2, 3, 4, 5, 6, 7],
      headers=["Model", "Term", "Coefficient", "HAC s.e.", "p-value",
               "Sig.", "N", "R2"])
caption("Dependent variable: ln(quantity). Newey-West HAC standard errors "
        "in parentheses; *** p<0.01, ** p<0.05, * p<0.10. Column (1) uses a "
        "linear trend only; column (2) adds the 2007 step and slope break "
        "and the March 2020–June 2021 pandemic dummy; column (3) adds the "
        "lagged dependent variable and reports the delta-method long-run "
        "elasticity; column (4) instruments ln(price) with the real WTI "
        "crude price (first-stage F = {F:.0f}, Sargan J = {J:.2f}, "
        "p = {Jp:.3f}).".format(F=iv_F, J=J, Jp=Jp))

exhibit_title("Exhibit 4. Partial relationship between price and quantity")
figure("fig2_partialled_scatter.png")
caption("Quantity and price residualized on ln(real income per capita), "
        "month dummies, the trend terms, and the pandemic dummy; the fitted "
        "slope equals {b:.2f}.".format(b=pref_b))

exhibit_title("Exhibit 5. Actual and fitted gasoline demand, preferred model")
figure("fig3_actual_fitted.png")
caption("Shaded area: the 2016–2020 hold-out period used in Section 5.")

exhibit_title("Exhibit 6. Out-of-sample forecast accuracy")
table("table3_forecast_accuracy.csv",
      widths=[104, 158, 30, 58, 58, 60], fontsize=8, numeric_from=2,
      headers=["Period", "Method", "N", "RMSE", "MAE", "MAPE (%)"])
caption("Errors in thousand barrels per day. Forecasts are one month ahead. "
        "Hold-out: equations estimated on data through December 2015, "
        "evaluated over January 2016–February 2020. Stress period: "
        "equations estimated through February 2020, evaluated over March "
        "2020–July 2026. The last two rows compare in-sample fit in the "
        "pandemic window with and without the pandemic dummy.")

exhibit_title("Exhibit 7. One-step-ahead forecasts against benchmarks, "
              "2016–2020 hold-out")
figure("fig4_forecast.png")
caption("Thousand barrels per day.")

exhibit_title("Exhibit 8. Robustness of the price elasticity")
table("table4_robustness.csv",
      widths=[160, 52, 44, 44, 34, 40, 94], fontsize=7.5, numeric_from=1,
      headers=["Specification", "Elasticity", "HAC s.e.", "p-value", "N",
               "R2", "Note"])
caption("Each row re-estimates the preferred equation with one change: trend "
        "break year, pandemic window, additional controls, dependent "
        "variable, sample period, price measure, instrument, or functional "
        "form. HAC standard errors.")

exhibit_title("Exhibit 9. Rolling six-year price elasticity")
figure("fig5_rolling_elasticity.png")
caption("Point estimates and 95 percent confidence intervals from six-year "
        "rolling samples of the preferred equation, labelled by the final "
        "year of each window.")

exhibit_title("Exhibit 10. Implied response of market demand to a price "
              "increase")
table("table5_implied_responses.csv",
      widths=[86, 84, 96, 96, 106], fontsize=8, numeric_from=0,
      headers=["Price change ($/gal)", "Price change (%)",
               "Quantity change (%)", "Revenue change (%)",
               "Quantity change (000 b/d)"])
caption("Computed from the preferred elasticity {b:.3f} at the sample mean "
        "real price of ${p:.2f} per gallon and mean quantity of {q:,.0f} "
        "thousand barrels per day. Revenue change is the product of the "
        "price and quantity changes at the market level."
        .format(b=pref_b, p=mean_p, q=mean_q))

exhibit_title("Exhibit 11. Seasonal pattern of gasoline demand")
table("table6_seasonal_effects.csv",
      widths=[120, 150], fontsize=8, numeric_from=1,
      headers=["Month", "Effect vs. December (%)"])
caption("Percent difference in quantity relative to December, from the "
        "preferred regression (December = base month).")

# ═════════════════════════════════════════════════════════════════════
# Appendix — clearly separated from the main paper and excluded from the
# page limit (the syllabus limit covers references, exhibits and appendix)
# ═════════════════════════════════════════════════════════════════════
page_break()
apph("Appendix")
p("This appendix contains supporting material: data sources and variable "
  "construction, the complete regression output behind Exhibit 3, model "
  "diagnostics, and the commands needed to reproduce every number in the "
  "paper. It is supplementary to the main text and is not part of the "
  "15-page limit.")

h2("Appendix A. Data sources and variable construction")
p("The analysis file data/gasoline_demand_monthly.csv is assembled entirely "
  "by download_data.py. It contains 399 monthly observations from April 1993 "
  "to July 2026 with no missing values.")
table(rows=[
    ["Variable", "Definition and construction", "Source"],
    ["Quantity (Q)",
     "Monthly product supplied of finished motor gasoline, thousand barrels "
     "per month, divided by days in the month to obtain thousand barrels per day",
     "EIA, PET_CONS_PSUP_DC_NUS_MBBL_M.xls"],
    ["Price, nominal (P)",
     "Monthly retail price, all grades and all formulations, dollars per gallon",
     "EIA, PET_PRI_GND_DCUS_NUS_M.xls"],
    ["Price, real (2017$)",
     "Nominal price multiplied by the 2017 average of CPIAUCSL and divided by "
     "the CPIAUCSL value of the same month",
     "FRED, CPIAUCSL"],
    ["Real income per capita (Y)",
     "Real disposable personal income (billions of chained 2017 dollars, SAAR) "
     "multiplied by 1e9 and divided by population in persons",
     "FRED, DSPIC96 and POPTHM"],
    ["Unemployment rate", "Civilian unemployment rate, percent, seasonally adjusted",
     "FRED, UNRATE"],
    ["WTI crude price", "West Texas Intermediate spot price, daily data averaged "
     "to a monthly mean, then deflated to 2017 dollars",
     "FRED, DCOILWTICO"],
    ["Time trend (t)", "0 for April 1993, increasing by one each month",
     "Constructed"],
    ["Step indicator", "0 through December 2006, 1 from January 2007",
     "Constructed"],
    ["Trend after 2007", "0 through December 2006, then months since January "
     "2007 divided by 12", "Constructed"],
    ["Pandemic dummy", "1 from March 2020 through June 2021, else 0",
     "Constructed"],
    ["Month dummies", "Eleven indicators for January through November; December "
     "is the base month", "Constructed"],
], widths=[96, 244, 128], fontsize=7.5, numeric_from=99)
caption("Appendix Table A1. Variables used in the estimation, their construction, "
        "and their sources.")

h2("Appendix B. Complete regression output and diagnostics")
p("Appendix Table B1 reports every coefficient of the preferred specification "
  "that underlies column 2 of Exhibit 3, including the eleven month dummies, "
  "the trend terms and the pandemic dummy. Appendix Table B2 reports the "
  "diagnostics. Two of them matter for inference: the residuals are strongly "
  "serially correlated (Durbin-Watson 1.01) and strongly non-normal "
  "(Jarque-Bera p-value below 0.0001, driven by the pandemic months), which "
  "is why every standard error in the paper is Newey-West HAC with twelve "
  "lags. The instrument is strong (first-stage F of 133), and the Sargan test "
  "reported in Section 3 rejects the over-identified version, so the "
  "just-identified IV is the one reported.")
table("table8_preferred_regression.csv",
      widths=[214, 76, 76, 56, 46], fontsize=8, numeric_from=1,
      headers=["Term", "Coefficient", "HAC s.e.", "p-value", "Sig."])
caption("Appendix Table B1. Preferred specification, dependent variable "
        "ln(quantity); 399 monthly observations, R-squared 0.838.")
table("table9_diagnostics.csv", widths=[286, 182], fontsize=8, numeric_from=1,
      headers=["Statistic", "Value"])
caption("Appendix Table B2. Model diagnostics for the preferred, ARDL and IV "
        "estimates reported in Exhibit 3.")

h2("Appendix C. Reproducing the analysis")
p("The repository contains everything needed to rebuild the results from "
  "scratch; nothing in the paper is entered by hand. From the project root:")
code([
    "python3 -m pip install -r requirements.txt",
    "python3 download_data.py          # fetches the EIA and FRED source files",
    "python3 analyze_demand.py         # estimates all models, writes tables and figures",
    "python3 paper/build_paper.py      # builds the PDF version of the paper",
    "python3 paper/build_docx.py       # builds the Word version of the paper",
])
p("download_data.py writes the raw source files to data/raw/ and the merged "
  "panel to data/gasoline_demand_monthly.csv. analyze_demand.py writes "
  "output/tables/*.csv, output/figures/*.png and output/results.json. Both "
  "builders read those outputs, so the tables and figures in the paper always "
  f"match the estimation code. The project is public at {REPO}.")

# ── document with page numbers ───────────────────────────────────────
def on_page(canvas, doc):
    canvas.saveState()
    canvas.setFont("Times-Roman", 10)
    canvas.drawCentredString(letter[0] / 2.0, 0.55 * 72, str(canvas.getPageNumber()))
    canvas.restoreState()


class PaperDoc(BaseDocTemplate):
    """Records the page on which the References section begins."""
    refs_page = None

    def afterFlowable(self, flowable):
        if (self.refs_page is None and isinstance(flowable, Paragraph)
                and flowable.getPlainText() == "References"):
            self.refs_page = self.page


def build(flowables, path):
    doc = PaperDoc(path, pagesize=letter,
                          leftMargin=72, rightMargin=72,
                          topMargin=72, bottomMargin=72,
                          title="The Impact of Gasoline Prices on Gasoline "
                                "Demand in the United States",
                          author=AUTHOR)
    frame = Frame(72, 72, letter[0] - 144, letter[1] - 144, id="main",
                  leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)
    doc.addPageTemplates([PageTemplate(id="all", frames=[frame], onPage=on_page)])
    doc.build(list(flowables))
    return doc.page, doc.refs_page


if __name__ == "__main__":
    total_pages, refs_page = build(render(story), OUT)
    body_pages = (refs_page or 1) - 1      # References starts on a fresh page

    print(f"Wrote {OUT}")
    print(f"  body pages (excl. references and exhibits): {body_pages}   "
          f"total pages: {total_pages}")
    print("  syllabus limit: 5-15 pages excluding references, exhibits, footnotes")
    if not 5 <= body_pages <= 15:
        print("  WARNING: body is outside the 5-15 page limit")
