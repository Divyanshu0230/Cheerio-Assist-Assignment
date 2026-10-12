"""Build the Cheerio Assist cost workbook.

Blue cells are inputs. Every other number is a formula so a reviewer
can change an assumption and watch the margin move.

Upload the xlsx to Google Sheets with File > Import > Upload.
"""

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.fill import PatternFillProperties, ColorChoice
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.chart.legend import Legend
from openpyxl.utils import get_column_letter
from openpyxl.chart.series import SeriesLabel
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.chart.marker import DataPoint as DP
from copy import copy

wb = Workbook()

# --- styles ---
ink = "1A1814"
green = "0E3B2E"
green_mid = "0E6B4F"
paper = "F7F4EC"
input_fill = PatternFill("solid", fgColor="D9E8F6")
input_font = Font(name="Calibri", size=11, color="1F4E79", bold=True)
formula_fill = PatternFill("solid", fgColor="FFFcf7")
formula_font = Font(name="Calibri", size=11, color=ink)
label_font = Font(name="Calibri", size=11, color=ink)
header_fill = PatternFill("solid", fgColor=green)
header_font = Font(name="Calibri", size=11, color="FFFFFF", bold=True)
section_font = Font(name="Calibri", size=14, color=green, bold=True)
title_font = Font(name="Calibri", size=18, color=green, bold=True)
note_font = Font(name="Calibri", size=10, color="5C564C", italic=True)
link_font = Font(name="Calibri", size=10, color="1F4E79", italic=True)
big_font = Font(name="Calibri", size=20, color=green, bold=True)
thin = Border(
    left=Side(style="thin", color="E4DDD0"),
    right=Side(style="thin", color="E4DDD0"),
    top=Side(style="thin", color="E4DDD0"),
    bottom=Side(style="thin", color="E4DDD0"),
)
wrap = Alignment(wrap_text=True, vertical="center")
left = Alignment(vertical="center", wrap_text=True)
center = Alignment(vertical="center", horizontal="center", wrap_text=True)

INR = '₹#,##0'
INR2 = '₹#,##0.00'
PCT = "0.0%"
NUM = "#,##0.0"
NUM0 = "#,##0"
USD = '"$"#,##0.00'
USDsmall = '"$"#,##0.000000'
FX = "0.00"


def apply_input(cell, fmt=None):
    cell.font = input_font
    cell.fill = input_fill
    cell.border = thin
    cell.alignment = center
    if fmt:
        cell.number_format = fmt


def apply_formula(cell, fmt=None):
    cell.font = formula_font
    cell.fill = formula_fill
    cell.border = thin
    cell.alignment = center
    if fmt:
        cell.number_format = fmt


def apply_label(cell):
    cell.font = label_font
    cell.alignment = left


def header_row(ws, row, cols, fill=header_fill):
    for col in range(1, cols + 1):
        cell = ws.cell(row, col)
        cell.font = header_font
        cell.fill = fill
        cell.alignment = center
        cell.border = thin


def set_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def note(ws, cell, text):
    ws[cell] = text
    ws[cell].font = note_font
    ws[cell].alignment = wrap


# =============================================================================
# Assumptions
# =============================================================================
assumptions = wb.active
assumptions.title = "Assumptions"

set_widths(assumptions, [38, 28, 16, 18, 18, 18, 62])
assumptions.row_dimensions[1].height = 24
assumptions["A1"] = "Cheerio Assist — inputs"
assumptions["A1"].font = title_font
note(
    assumptions,
    "A2",
    "Blue cells are the only inputs. Change them and Summary, Volume, Cost, Billing, Roadmap, and Audit all move. "
    "Black cells on the other sheets are formulas. Do not paste values over them. "
    "These volumes are a guesstimate. Cheerio did not publish ticket counts. If you have the export, replace the blue cells.",
)
assumptions.merge_cells("A2:G2")
assumptions.row_dimensions[2].height = 36

assumptions["A4"] = "Company"
assumptions["A4"].font = section_font

assumptions["A5"] = "Active businesses"
assumptions["C5"] = 5000
apply_input(assumptions["C5"], NUM0)
assumptions["D5"] = "businesses"
assumptions["G5"] = "Given in the brief. Treat as the businesses that log in during a month, not signups all-time."
apply_label(assumptions["A5"])
assumptions["G5"].font = note_font
assumptions["G5"].alignment = wrap

assumptions["A6"] = "Already solved with no human, before Assist"
assumptions["C6"] = 0.08
apply_input(assumptions["C6"], PCT)
assumptions["D6"] = "share of queries"
assumptions["G6"] = "Help docs already exist, so not every query becomes a ticket today. 8% is a judgment, not a measurement."
apply_label(assumptions["A6"])
assumptions["G6"].font = note_font
assumptions["G6"].alignment = wrap

assumptions["A7"] = "Month-3 share of queries that open Assist"
assumptions["C7"] = 0.70
apply_input(assumptions["C7"], PCT)
assumptions["D7"] = "share of queries"
assumptions["G7"] = "The other 30% still email or WhatsApp the support line and never see the panel."
apply_label(assumptions["A7"])
assumptions["G7"].font = note_font
assumptions["G7"].alignment = wrap

assumptions["A9"] = "Who raises the queries"
assumptions["A9"].font = section_font
for col, text in enumerate(["Segment", "Share of businesses", "Queries per business / month", "", "", "", "Why this query rate"], 1):
    assumptions.cell(10, col, text)
header_row(assumptions, 10, 7)

segments = [
    (11, "Onboarding or a new channel", 0.20, 2.5, "First 45 days, or they added a channel this month. Setup, first template, first workflow."),
    (12, "Steady weekly use", 0.65, 1.0, "A campaign, a template, or a silent bot in a typical month. Not every week."),
    (13, "Light or seasonal", 0.15, 0.4, "They pay, they rarely log in. A bill or a login problem a couple of times a year."),
]
for row, name, share, qpb, why in segments:
    assumptions.cell(row, 1, name).font = label_font
    assumptions.cell(row, 2, share)
    apply_input(assumptions.cell(row, 2), PCT)
    assumptions.cell(row, 3, qpb)
    apply_input(assumptions.cell(row, 3), "0.0")
    assumptions.cell(row, 7, why).font = note_font
    assumptions.cell(row, 7).alignment = wrap
    assumptions.row_dimensions[row].height = 28

assumptions["A14"] = "Shares must equal 100%"
assumptions["B14"] = "=SUM(B11:B13)"
apply_formula(assumptions["B14"], PCT)
apply_label(assumptions["A14"])
assumptions["G14"] = '=IF(ABS(B14-1)<0.0001,"OK","Fix the shares. They have to sum to 100%.")'
assumptions["G14"].font = formula_font
assumptions["G14"].alignment = left

assumptions["A16"] = "What the queries are, and how many Assist can finish"
assumptions["A16"].font = section_font
headers = [
    "Issue",
    "Share of queries",
    "Playbook can finish",
    "Live in month 1",
    "Live in month 2",
    "Live in month 3",
    "What the rate means",
]
for col, text in enumerate(headers, 1):
    assumptions.cell(17, col, text)
header_row(assumptions, 17, 7)

issues = [
    (18, "Template rejected", 0.26, 0.72, 1, 1, 1,
     "Name format, Utility vs Marketing, missing variable example. A second rejection goes to a person. I would not let Assist keep editing."),
    (19, "WhatsApp setup stuck", 0.20, 0.48, 0, 1, 1,
     "Assist can show the incomplete step. It cannot make someone a Facebook admin, or free a number that is still on the WhatsApp app."),
    (20, "Workflow did not reply", 0.16, 0.62, 0, 0, 1,
     "Toggle off, keyword does not match, or the template is not approved. Redesigning the flow stays with a person."),
    (21, "Campaign did not send", 0.12, 0.75, 1, 1, 1,
     "Usually the template. Status is a fact Assist can show. This is separate from a money question."),
    (22, "Why was I charged?", 0.06, 0.10, 1, 1, 1,
     "Assist explains the debit and stops. Almost everyone with a money question still wants a person. 10% accept the explanation."),
    (23, "Integration not syncing", 0.10, 0.35, 0, 0, 1,
     "Empty Shopify connection or blank webhook is checkable. The other company's error is not."),
    (24, "Other", 0.10, 0.10, 0, 0, 1,
     "Bugs, access, and one-off asks. A playbook should barely touch these."),
]
for row, name, share, rate, m1, m2, m3, why in issues:
    assumptions.cell(row, 1, name).font = label_font
    assumptions.cell(row, 2, share)
    apply_input(assumptions.cell(row, 2), PCT)
    assumptions.cell(row, 3, rate)
    apply_input(assumptions.cell(row, 3), PCT)
    for col, flag in ((4, m1), (5, m2), (6, m3)):
        assumptions.cell(row, col, flag)
        apply_input(assumptions.cell(row, col), "0")
    assumptions.cell(row, 7, why).font = note_font
    assumptions.cell(row, 7).alignment = wrap
    assumptions.row_dimensions[row].height = 36

assumptions["A25"] = "Shares must equal 100%"
assumptions["B25"] = "=SUM(B18:B24)"
apply_formula(assumptions["B25"], PCT)
apply_label(assumptions["A25"])
assumptions["G25"] = '=IF(ABS(B25-1)<0.0001,"OK","Fix the shares. They have to sum to 100%.")'
assumptions["G25"].font = formula_font

assumptions["A27"] = "Human support"
assumptions["A27"].font = section_font
people = [
    (28, "Agent monthly CTC", 40000, INR, "INR / month", "Assumption for an L1 agent. Not a salary survey. Edit it."),
    (29, "Load for benefits, tools, and a lead's time", 0.20, PCT, "share of CTC", "Loaded cost = CTC × (1 + this)."),
    (30, "Tickets one agent finishes in a day", 10, "0", "tickets", "How-to tickets. Deep bugs are slower. If you believe 8, type 8."),
    (31, "Working days in a month", 22, "0", "days", ""),
]
for row, label, value, fmt, unit, why in people:
    assumptions.cell(row, 1, label).font = label_font
    assumptions.cell(row, 3, value)
    apply_input(assumptions.cell(row, 3), fmt)
    assumptions.cell(row, 4, unit).font = note_font
    assumptions.cell(row, 7, why).font = note_font

assumptions["A33"] = "Model — gpt-4o-mini"
assumptions["A33"].font = section_font
note(
    assumptions,
    "A34",
    "Prices from OpenAI's API pricing page, checked 12 Oct 2026: gpt-4o-mini input $0.15, cached input $0.075, output $0.60 per 1M tokens. "
    "https://developers.openai.com/api/docs/pricing  —  USD/INR is the 12 Oct 2026 market, about ₹95.75 to ₹95.83 per dollar. Both are inputs.",
)
assumptions.merge_cells("A34:G34")
assumptions.row_dimensions[34].height = 32

model_rows = [
    (35, "Input price", 0.15, USD, "USD per 1M tokens"),
    (36, "Cached input price", 0.075, USD, "USD per 1M tokens"),
    (37, "Output price", 0.60, USD, "USD per 1M tokens"),
    (38, "Rupees per US dollar", 95.80, FX, "INR"),
    (39, "Cached tokens per session", 2500, NUM0, "tokens"),
    (40, "Fresh input tokens per session", 900, NUM0, "tokens"),
    (41, "Output tokens per session", 450, NUM0, "tokens"),
    (42, "Share of sessions with a follow-up turn", 0.20, PCT, "share"),
    (43, "Follow-up fresh input tokens", 600, NUM0, "tokens"),
    (44, "Follow-up output tokens", 350, NUM0, "tokens"),
]
for row, label, value, fmt, unit in model_rows:
    assumptions.cell(row, 1, label).font = label_font
    assumptions.cell(row, 3, value)
    apply_input(assumptions.cell(row, 3), fmt)
    assumptions.cell(row, 4, unit).font = note_font

assumptions["G39"] = "Playbook text plus a short account snapshot. Cached because the playbook repeats across sessions."
assumptions["G39"].font = note_font
assumptions["G39"].alignment = wrap
assumptions["G40"] = "The user's words plus the live status that changed this session."
assumptions["G40"].font = note_font

assumptions["A46"] = "Other monthly cost"
assumptions["A46"].font = section_font
other = [
    (47, "Hosting, logs, and search index", 15000, INR, "A small flat cost. Tokens are not the bill that matters."),
    (48, "Playbook owner", 55000, INR, "One person who edits the five playbooks and reads the weekly sample. Without them the answers rot."),
    (49, "Owner hours in a month", 176, "0", "22 × 8. Used only if someone later prices a human hour."),
]
for row, label, value, fmt, why in other:
    assumptions.cell(row, 1, label).font = label_font
    assumptions.cell(row, 3, value)
    apply_input(assumptions.cell(row, 3), fmt)
    assumptions.cell(row, 7, why).font = note_font
    assumptions.cell(row, 7).alignment = wrap

assumptions["A51"] = "The price I would not charge"
assumptions["A51"].font = section_font
note(
    assumptions,
    "A52",
    "Kept here so the Billing sheet can show the counterfactual. The recommendation is to include Assist free. A fee is an input so you can see the tickets come back.",
)
assumptions.merge_cells("A52:G52")

assumptions["A53"] = "Fee per Assist resolution"
assumptions["C53"] = 49
apply_input(assumptions["C53"], INR)
assumptions["D53"] = "INR"
apply_label(assumptions["A53"])
assumptions["G53"] = "Charged only when Assist closes it. Sessions that still need a person are free in this counterfactual."
assumptions["G53"].font = note_font

assumptions["A54"] = "Adoption if that fee exists"
assumptions["C54"] = 0.35
apply_input(assumptions["C54"], PCT)
apply_label(assumptions["A54"])
assumptions["G54"] = "Half of the free-case adoption. People who will not pay walk back to the human inbox, which is still free."
assumptions["G54"].font = note_font
assumptions["G54"].alignment = wrap

assumptions["A56"] = "Roadmap adoption, on the issues marked live that month"
assumptions["A56"].font = section_font
assumptions["A57"] = "Month 1 adoption"
assumptions["C57"] = 0.45
apply_input(assumptions["C57"], PCT)
apply_label(assumptions["A57"])
assumptions["G57"] = "Template, campaign-did-not-send, and the charge question. The panel is new."
assumptions["G57"].font = note_font

assumptions["A58"] = "Month 2 adoption"
assumptions["C58"] = 0.60
apply_input(assumptions["C58"], PCT)
apply_label(assumptions["A58"])
assumptions["G58"] = "WhatsApp setup joins. Adoption rises because the panel is already on the failed campaign."
assumptions["G58"].font = note_font

assumptions["A59"] = "Month 3 adoption"
assumptions["C59"] = "=C7"
apply_formula(assumptions["C59"], PCT)
apply_label(assumptions["A59"])
assumptions["G59"] = "Tied to the month-3 input above so the roadmap cannot drift from the main case."
assumptions["G59"].font = note_font

assumptions["A61"] = "Audit thresholds and a worked example"
assumptions["A61"].font = section_font
note(
    assumptions,
    "A62",
    "The worked example is sample close data so the Audit sheet's PASS/ALERT logic is visible. Replace it with the real invoice and the real sample when the month closes.",
)
assumptions.merge_cells("A62:G62")

audit_inputs = [
    (63, "Invoice may exceed logged token cost by", 0.15, PCT, "Alert if the OpenAI invoice is higher than logs × (1 + this)."),
    (64, "Repeat contact within 7 days — alert above", 0.12, PCT, "Same issue, another contact. Deflection that bounces back."),
    (65, "Worked example — repeat rate observed", 0.07, PCT, "Sample. Type the measured rate at month close."),
    (66, "Worked example — wrong one-click fixes in 30", 0, "0", "Sample. One wrong edit is enough to pull that playbook's button."),
    (67, "Worked example — wallet or send actions by Assist", 0, "0", "Must stay zero. Assist does not move money or message guests."),
    (68, "Worked example — OpenAI invoice", 280, INR, "Sample invoice. Expected log cost is about ₹265, so 280 should PASS."),
]
for row, label, value, fmt, why in audit_inputs:
    assumptions.cell(row, 1, label).font = label_font
    assumptions.cell(row, 3, value)
    apply_input(assumptions.cell(row, 3), fmt)
    assumptions.cell(row, 7, why).font = note_font
    assumptions.cell(row, 7).alignment = wrap

assumptions.row_dimensions[16].height = 20
assumptions.auto_filter.ref = "A17:G24"
assumptions.freeze_panes = "A4"
assumptions.page_setup.orientation = "landscape"
assumptions.page_setup.fitToPage = True
assumptions.page_setup.fitToWidth = 1
assumptions.page_setup.fitToHeight = 3
assumptions.page_setup.paperSize = assumptions.PAPERSIZE_A4
assumptions.sheet_properties.tabColor = "1F4E79"
assumptions.oddHeader.left.text = "Cheerio Assist  ·  Assumptions  ·  blue = input"
assumptions.print_title_rows = "1:2"
assumptions.page_setup.horizontalCentered = True
assumptions.sheet_view.showGridLines = False
assumptions.sheet_view.zoomScale = 110

# data validation: live flags 0 or 1
dv = DataValidation(type="list", formula1='"0,1"', allow_blank=False)
dv.error = "Use 0 or 1"
dv.errorTitle = "Live flag"
dv.add("D18:F24")
assumptions.add_data_validation(dv)

# =============================================================================
# Volume
# =============================================================================
volume = wb.create_sheet("Volume")
set_widths(volume, [34, 22, 24, 28, 26, 28, 24, 26])
volume["A1"] = "Query volume"
volume["A1"].font = title_font
note(
    volume,
    "A2",
    "Every number on this sheet is a formula pointed at Assumptions. "
    "A query becomes a human ticket unless someone finishes it with no ticket. "
    "Today that is only the unaided share. In month 3 it is the unaided share for people who never open Assist, plus the playbook for people who do.",
)
volume.merge_cells("A2:H2")
volume.row_dimensions[2].height = 32

volume["A4"] = "Queries this month"
volume["A4"].font = section_font
for col, text in enumerate(["Segment", "Share of businesses", "Businesses", "Queries per business", "Queries"], 1):
    volume.cell(5, col, text)
header_row(volume, 5, 5)

for i, src in enumerate((11, 12, 13)):
    row = 6 + i
    volume.cell(row, 1, f"=Assumptions!A{src}")
    volume.cell(row, 2, f"=Assumptions!B{src}")
    volume.cell(row, 3, f"=Assumptions!$C$5*B{row}")
    volume.cell(row, 4, f"=Assumptions!C{src}")
    volume.cell(row, 5, f"=C{row}*D{row}")
    for col in range(1, 6):
        apply_formula(volume.cell(row, col))
    volume.cell(row, 2).number_format = PCT
    volume.cell(row, 3).number_format = NUM0
    volume.cell(row, 4).number_format = "0.0"
    volume.cell(row, 5).number_format = NUM

volume["A9"] = "Total"
volume["B9"] = "=SUM(B6:B8)"
volume["C9"] = "=SUM(C6:C8)"
volume["E9"] = "=SUM(E6:E8)"
for col in (1, 2, 3, 5):
    apply_formula(volume.cell(9, col))
volume["B9"].number_format = PCT
volume["C9"].number_format = NUM0
volume["E9"].number_format = NUM
volume["A9"].font = Font(name="Calibri", size=11, color=ink, bold=True)

volume["A11"] = "Today, before Assist"
volume["A11"].font = section_font
today_rows = [
    (12, "Queries", "=E9", NUM, "From the segments above."),
    (13, "Already closed with no human", "=B12*Assumptions!C6", NUM, "The 8% (or whatever you typed)."),
    (14, "Human tickets", "=B12-B13", NUM, "This is the queue."),
    (15, "Human tickets per 100 active businesses", "=B14/Assumptions!C5*100", "0.0", "North star, today."),
]
for row, label, formula, fmt, why in today_rows:
    volume.cell(row, 1, label).font = label_font
    volume.cell(row, 2, formula)
    apply_formula(volume.cell(row, 2), fmt)
    volume.cell(row, 5, why).font = note_font

volume["A17"] = "Month 3, with Assist"
volume["A17"].font = section_font
m3_headers = [
    "Issue",
    "Queries",
    "Playbook finish rate",
    "Adoption on this issue",
    "Share closed with no human",
    "Closed with no human",
    "Human tickets",
    "Closed by Assist itself",
]
for col, text in enumerate(m3_headers, 1):
    volume.cell(18, col, text)
header_row(volume, 18, 8)
volume.row_dimensions[18].height = 30

for i in range(7):
    row = 19 + i
    src = 18 + i
    volume.cell(row, 1, f"=Assumptions!A{src}")
    volume.cell(row, 2, f"=$E$9*Assumptions!B{src}")
    volume.cell(row, 3, f"=Assumptions!C{src}")
    volume.cell(row, 4, f"=Assumptions!$C$7*Assumptions!F{src}")
    volume.cell(row, 5, f"=(1-D{row})*Assumptions!$C$6+D{row}*C{row}")
    volume.cell(row, 6, f"=B{row}*E{row}")
    volume.cell(row, 7, f"=B{row}-F{row}")
    volume.cell(row, 8, f"=B{row}*D{row}*C{row}")
    for col in range(1, 9):
        apply_formula(volume.cell(row, col))
    volume.cell(row, 2).number_format = NUM
    volume.cell(row, 3).number_format = PCT
    volume.cell(row, 4).number_format = PCT
    volume.cell(row, 5).number_format = PCT
    volume.cell(row, 6).number_format = NUM
    volume.cell(row, 7).number_format = NUM
    volume.cell(row, 8).number_format = NUM

volume["A26"] = "Month 3 total"
volume["B26"] = "=SUM(B19:B25)"
volume["F26"] = "=SUM(F19:F25)"
volume["G26"] = "=SUM(G19:G25)"
volume["H26"] = "=SUM(H19:H25)"
for col in (1, 2, 6, 7, 8):
    apply_formula(volume.cell(26, col))
    volume.cell(26, col).font = Font(name="Calibri", size=11, color=ink, bold=True)
volume["B26"].number_format = NUM
volume["F26"].number_format = NUM
volume["G26"].number_format = NUM
volume["H26"].number_format = NUM

tail = [
    (28, "Human tickets per 100 businesses, month 3", "=G26/Assumptions!C5*100", "0.0"),
    (29, "Drop in human tickets vs today", "=(B14-G26)/B14", PCT),
    (30, "Queries closed by Assist itself", "=H26", NUM),
    (31, "Sessions that open Assist (we pay tokens on these)", "=E9*Assumptions!C7", NUM),
]
for row, label, formula, fmt in tail:
    volume.cell(row, 1, label).font = label_font
    volume.cell(row, 2, formula)
    apply_formula(volume.cell(row, 2), fmt)

note(
    volume,
    "A33",
    "Share closed with no human = (1 − adoption) × unaided rate + adoption × playbook rate. "
    "People who never open Assist keep today's unaided rate. People who open it get the playbook rate instead, not on top. "
    "Closed by Assist itself = queries × adoption × playbook rate. That is the slice we compare with a human agent.",
)
volume.merge_cells("A33:H33")
volume.row_dimensions[33].height = 36
volume.freeze_panes = "A4"
volume.sheet_properties.tabColor = "0E6B4F"
volume.sheet_view.showGridLines = False
volume.page_setup.orientation = "landscape"
volume.page_setup.fitToPage = True
volume.page_setup.fitToWidth = 1
volume.page_setup.fitToHeight = 1
volume.page_setup.paperSize = volume.PAPERSIZE_A4
volume.oddHeader.left.text = "Cheerio Assist  ·  Volume"

# =============================================================================
# Cost
# =============================================================================
cost = wb.create_sheet("Cost")
set_widths(cost, [62, 22, 78])
cost["A1"] = "What it costs to run, versus agents"
cost["A1"].font = title_font
note(
    cost,
    "A2",
    "Token math is here because the brief asks for it. It is not the decision. The playbook owner and the agents we still need are the decision. "
    "Fractional FTE is the fair comparison. Whole people are who you would actually roster. Both are shown.",
)
cost.merge_cells("A2:C2")
cost.row_dimensions[2].height = 32

cost["A4"] = "One human ticket"
cost["A4"].font = section_font
cost_people = [
    (5, "Loaded cost of one agent per month", "=Assumptions!C28*(1+Assumptions!C29)", INR, "CTC × (1 + load)."),
    (6, "Tickets one agent finishes in a month", "=Assumptions!C30*Assumptions!C31", NUM0, "Days × tickets per day."),
    (7, "Cost of one human ticket", "=B5/B6", INR2, "This is the number Assist has to beat."),
]
for row, label, formula, fmt, why in cost_people:
    cost.cell(row, 1, label).font = label_font
    cost.cell(row, 2, formula)
    apply_formula(cost.cell(row, 2), fmt)
    cost.cell(row, 3, why).font = note_font

cost["A9"] = "Tokens"
cost["A9"].font = section_font
cost["A10"] = "USD per Assist session"
cost["B10"] = (
    "=Assumptions!C39/1000000*Assumptions!C36"
    "+Assumptions!C40/1000000*Assumptions!C35"
    "+Assumptions!C41/1000000*Assumptions!C37"
    "+Assumptions!C42*(Assumptions!C43/1000000*Assumptions!C35+Assumptions!C44/1000000*Assumptions!C37)"
)
apply_formula(cost["B10"], USDsmall)
apply_label(cost["A10"])
cost["C10"] = "Cached tokens at the cached price, fresh tokens at the input price, output at the output price, plus a follow-up on some sessions."
cost["C10"].font = note_font
cost["C10"].alignment = wrap
cost.row_dimensions[10].height = 32

cost["A11"] = "INR per Assist session"
cost["B11"] = "=B10*Assumptions!C38"
apply_formula(cost["B11"], INR2)
apply_label(cost["A11"])

cost["A12"] = "Sessions in month 3"
cost["B12"] = "=Volume!B31"
apply_formula(cost["B12"], NUM)
apply_label(cost["A12"])
cost["C12"] = "One session per query that opens Assist, including the ones a person still has to finish."
cost["C12"].font = note_font

cost["A13"] = "Token bill in month 3"
cost["B13"] = "=B12*B11"
apply_formula(cost["B13"], INR2)
apply_label(cost["A13"])

cost["A15"] = "Assist, fully loaded, month 3"
cost["A15"].font = section_font
for row, label, formula, fmt, why in [
    (16, "Hosting, logs, and search", "=Assumptions!C47", INR, ""),
    (17, "Playbook owner", "=Assumptions!C48", INR, ""),
    (18, "Assist operating cost", "=B13+B16+B17", INR, "Tokens + hosting + the owner."),
]:
    cost.cell(row, 1, label).font = label_font
    cost.cell(row, 2, formula)
    apply_formula(cost.cell(row, 2), fmt)
    cost.cell(row, 3, why).font = note_font

cost["A20"] = "Company cost — fractional load"
cost["A20"].font = section_font
company = [
    (21, "Agent load today, FTE", "=Volume!B14/B6", "0.00", "Human tickets today ÷ monthly capacity."),
    (22, "Agent load in month 3, FTE", "=Volume!G26/B6", "0.00", "Tickets Assist did not close."),
    (23, "Human payroll today", "=B21*B5", INR, ""),
    (24, "Human payroll in month 3", "=B22*B5", INR, ""),
    (25, "Total support cost in month 3", "=B24+B18", INR, "Remaining agents + Assist operating cost."),
    (26, "Monthly saving vs today", "=B23-B25", INR, "The number I would manage the quarter against."),
]
for row, label, formula, fmt, why in company:
    cost.cell(row, 1, label).font = label_font
    cost.cell(row, 2, formula)
    apply_formula(cost.cell(row, 2), fmt)
    cost.cell(row, 3, why).font = note_font

cost["A28"] = "Same volume, two ways of finishing it"
cost["A28"].font = section_font
note(
    cost,
    "A29",
    "Same volume means the queries Assist itself closes. Human cost uses the per-ticket rate. Assist cost loads the entire operating bill onto that slice, tokens included. "
    "A portion of those users might have coped with the docs anyway. The company saving in row 26 already nets that out. This block is the unit comparison the brief asks for.",
)
cost.merge_cells("A29:C29")
cost.row_dimensions[29].height = 32

same = [
    (30, "Queries Assist closes", "=Volume!H26", NUM, ""),
    (31, "Assist cost per one of those", "=B18/B30", INR2, "Full operating cost ÷ Assist closures."),
    (32, "Human cost for the same queries", "=B30*B7", INR, ""),
    (33, "Difference on that slice", "=B32-B18", INR, "Human cost minus the whole Assist bill."),
]
for row, label, formula, fmt, why in same:
    cost.cell(row, 1, label).font = label_font
    cost.cell(row, 2, formula)
    apply_formula(cost.cell(row, 2), fmt)
    cost.cell(row, 3, why).font = note_font

cost["A35"] = "If you can only hire whole people"
cost["A35"].font = section_font
whole = [
    (36, "Agents to roster today", "=CEILING(B21,1)", "0", "Rounded up."),
    (37, "Agents to roster in month 3", "=CEILING(B22,1)", "0", "Rounded up. Do not drop headcount until the ticket rate actually moves."),
    (38, "Payroll today, whole people", "=B36*B5", INR, ""),
    (39, "Month-3 cost, whole people + Assist", "=B37*B5+B18", INR, ""),
    (40, "Saving, whole people", "=B38-B39", INR, ""),
]
for row, label, formula, fmt, why in whole:
    cost.cell(row, 1, label).font = label_font
    cost.cell(row, 2, formula)
    apply_formula(cost.cell(row, 2), fmt)
    cost.cell(row, 3, why).font = note_font

note(
    cost,
    "A42",
    "Meta's WhatsApp message rates are not in this sheet. Assist does not send those messages. "
    "A search summary called a dollar rate card a rupee rate. I left it out rather than build a margin on a mislabeled number.",
)
cost.merge_cells("A42:C42")
cost.row_dimensions[42].height = 28
cost.freeze_panes = "A4"
cost.sheet_properties.tabColor = "C2410C"
cost.sheet_view.showGridLines = False
cost.page_setup.orientation = "landscape"
cost.page_setup.fitToPage = True
cost.page_setup.fitToWidth = 1
cost.page_setup.fitToHeight = 1
cost.page_setup.paperSize = cost.PAPERSIZE_A4
cost.oddHeader.left.text = "Cheerio Assist  ·  Cost"

# =============================================================================
# Roadmap
# =============================================================================
roadmap = wb.create_sheet("Roadmap")
set_widths(roadmap, [32, 16, 16, 16, 18, 18, 18, 36])
roadmap["A1"] = "How the 3-month move actually happens"
roadmap["A1"].font = title_font
note(
    roadmap,
    "A2",
    "Month 1 is only the failed campaign and the template behind it, plus the charge question on that same screen. "
    "Month 2 adds WhatsApp setup. Month 3 adds workflows and a thin pass at integrations and everything else. "
    "Adoption is an input on Assumptions. Issues that are not live that month stay at today's unaided rate.",
)
roadmap.merge_cells("A2:H2")
roadmap.row_dimensions[2].height = 32

roadmap["A4"] = "Per issue"
roadmap["A4"].font = section_font
for col, text in enumerate(
    ["Issue", "Queries", "Month 1 human tickets", "Month 2 human tickets", "Month 3 human tickets",
     "Month 1 adoption used", "Month 2 adoption used", "Month 3 adoption used"],
    1,
):
    roadmap.cell(5, col, text)
header_row(roadmap, 5, 8)
roadmap.row_dimensions[5].height = 30

# Human tickets for an issue at a given adoption:
# queries * (1 - ((1-adoption)*unaided + adoption*rate))
# adoption used = month_adoption * live_flag
for i in range(7):
    row = 6 + i
    src = 18 + i
    vrow = 19 + i
    roadmap.cell(row, 1, f"=Assumptions!A{src}")
    roadmap.cell(row, 2, f"=Volume!B{vrow}")
    roadmap.cell(row, 6, f"=Assumptions!$C$57*Assumptions!D{src}")
    roadmap.cell(row, 7, f"=Assumptions!$C$58*Assumptions!E{src}")
    roadmap.cell(row, 8, f"=Assumptions!$C$59*Assumptions!F{src}")
    roadmap.cell(row, 3, f"=B{row}*(1-((1-F{row})*Assumptions!$C$6+F{row}*Assumptions!C{src}))")
    roadmap.cell(row, 4, f"=B{row}*(1-((1-G{row})*Assumptions!$C$6+G{row}*Assumptions!C{src}))")
    roadmap.cell(row, 5, f"=B{row}*(1-((1-H{row})*Assumptions!$C$6+H{row}*Assumptions!C{src}))")
    for col in range(1, 9):
        apply_formula(roadmap.cell(row, col))
    roadmap.cell(row, 2).number_format = NUM
    for col in (3, 4, 5):
        roadmap.cell(row, col).number_format = NUM
    for col in (6, 7, 8):
        roadmap.cell(row, col).number_format = PCT

roadmap["A13"] = "Total human tickets"
roadmap["B13"] = "=SUM(B6:B12)"
roadmap["C13"] = "=SUM(C6:C12)"
roadmap["D13"] = "=SUM(D6:D12)"
roadmap["E13"] = "=SUM(E6:E12)"
for col in range(1, 6):
    apply_formula(roadmap.cell(13, col))
    roadmap.cell(13, col).font = Font(name="Calibri", size=11, color=ink, bold=True)
for col in range(2, 6):
    roadmap.cell(13, col).number_format = NUM

roadmap["A15"] = "Against today"
roadmap["A15"].font = section_font
for col, text in enumerate(["", "Today", "Month 1", "Month 2", "Month 3"], 1):
    roadmap.cell(16, col, text)
header_row(roadmap, 16, 5)

roadmap["A17"] = "Human tickets"
roadmap["B17"] = "=Volume!B14"
roadmap["C17"] = "=C13"
roadmap["D17"] = "=D13"
roadmap["E17"] = "=E13"
roadmap["A18"] = "Per 100 active businesses"
roadmap["B18"] = "=B17/Assumptions!$C$5*100"
roadmap["C18"] = "=C17/Assumptions!$C$5*100"
roadmap["D18"] = "=D17/Assumptions!$C$5*100"
roadmap["E18"] = "=E17/Assumptions!$C$5*100"
roadmap["A19"] = "Drop vs today"
roadmap["B19"] = 0
roadmap["C19"] = "=(B17-C17)/B17"
roadmap["D19"] = "=(B17-D17)/B17"
roadmap["E19"] = "=(B17-E17)/B17"
for row in (17, 18, 19):
    apply_label(roadmap.cell(row, 1))
    for col in range(2, 6):
        apply_formula(roadmap.cell(row, col))
for col in range(2, 6):
    roadmap.cell(17, col).number_format = NUM
    roadmap.cell(18, col).number_format = "0.0"
    roadmap.cell(19, col).number_format = PCT
# B19 is a typed 0 but should look like a formula result. It's an input-like constant.
# Make it a formula so the blue-input rule stays clean.
roadmap["B19"] = "=0"
apply_formula(roadmap["B19"], PCT)

roadmap["A21"] = "Month 3 on this sheet matches Volume"
roadmap["B21"] = '=IF(ABS(E13-Volume!G26)<0.05,"PASS","ALERT")'
apply_formula(roadmap["B21"])
apply_label(roadmap["A21"])

note(
    roadmap,
    "A23",
    "Gate for the quarter: if month 1 does not cut human tickets by about 10%, do not start the setup playbook. Fix the campaign screen first. "
    "The 10% is a management rule, not a formula. Month 1 in this model is a bit above that, which is why month 2 is allowed.",
)
roadmap.merge_cells("A23:H23")
roadmap.row_dimensions[23].height = 32
roadmap.freeze_panes = "A4"
roadmap.sheet_properties.tabColor = "9A6B2F"
roadmap.sheet_view.showGridLines = False
roadmap.page_setup.orientation = "landscape"
roadmap.page_setup.fitToPage = True
roadmap.page_setup.fitToWidth = 1
roadmap.page_setup.fitToHeight = 1
roadmap.page_setup.paperSize = roadmap.PAPERSIZE_A4
roadmap.oddHeader.left.text = "Cheerio Assist  ·  Roadmap"

# =============================================================================
# Billing
# =============================================================================
billing = wb.create_sheet("Billing")
set_widths(billing, [56, 24, 70])
billing["A1"] = "How to charge"
billing["A1"].font = title_font
note(
    billing,
    "A2",
    "Include Assist on every paid plan. Do not sell it per conversation, per credit, or as a bot add-on. "
    "Growth already includes 4 support calls a month and Pro includes 10 (cheerio.ai/pricing, checked 12 Oct 2026). Escalation uses that entitlement. "
    "A new fee on the fix trains the owner to skip it and message a person, which is the path that costs about ₹218.",
)
billing.merge_cells("A2:C2")
billing.row_dimensions[2].height = 48

billing["A4"] = "Free Assist — company outcome vs doing nothing"
billing["A4"].font = section_font
free_rows = [
    (5, "Human payroll today", "=Cost!B23", INR),
    (6, "Month-3 total (agents still needed + Assist)", "=Cost!B25", INR),
    (7, "Net saving per month", "=B5-B6", INR),
    (8, "Assist operating cost", "=Cost!B18", INR),
    (9, "Saving per rupee of Assist cost", "=IF(B8=0,\"n/a\",B7/B8)", "0.00"),
]
for row, label, formula, fmt in free_rows:
    billing.cell(row, 1, label).font = label_font
    billing.cell(row, 2, formula)
    apply_formula(billing.cell(row, 2), fmt)

billing["C7"] = "This is the margin. It is payroll you do not add, not a new line of revenue."
billing["C7"].font = note_font
billing["C7"].alignment = wrap
billing["C9"] = "Above 1 means the operating cost comes back as avoided payroll."
billing["C9"].font = note_font

billing["A11"] = "If we metered it anyway"
billing["A11"].font = section_font
note(
    billing,
    "A12",
    "Fee and the lower adoption are inputs on Assumptions (the blue cells under “The price I would not charge”). "
    "Every issue is live, same as month 3. Only adoption changes. Hosting and the playbook owner stay, because you still run the system.",
)
billing.merge_cells("A12:C12")
billing.row_dimensions[12].height = 32

for col, text in enumerate(["Issue", "Human tickets if metered", "Assist closures if metered"], 1):
    billing.cell(14, col, text)
header_row(billing, 14, 3)

for i in range(7):
    row = 15 + i
    src = 18 + i
    vrow = 19 + i
    billing.cell(row, 1, f"=Assumptions!A{src}")
    # adoption is C54 for every issue (product is fully live)
    billing.cell(
        row,
        2,
        f"=Volume!B{vrow}*(1-((1-Assumptions!$C$54)*Assumptions!$C$6+Assumptions!$C$54*Assumptions!C{src}))",
    )
    billing.cell(row, 3, f"=Volume!B{vrow}*Assumptions!$C$54*Assumptions!C{src}")
    for col in range(1, 4):
        apply_formula(billing.cell(row, col))
    billing.cell(row, 2).number_format = NUM
    billing.cell(row, 3).number_format = NUM

billing["A22"] = "Metered total"
billing["B22"] = "=SUM(B15:B21)"
billing["C22"] = "=SUM(C15:C21)"
for col in range(1, 4):
    apply_formula(billing.cell(22, col))
    billing.cell(22, col).font = Font(name="Calibri", size=11, color=ink, bold=True)
billing["B22"].number_format = NUM
billing["C22"].number_format = NUM

metered = [
    (24, "Human payroll if metered", "=B22/Cost!B6*Cost!B5", INR, "Tickets ÷ capacity × loaded agent cost."),
    (25, "Token bill if metered", "=Volume!E9*Assumptions!C54*Cost!B11", INR, "Fewer sessions, same price per session."),
    (26, "Assist operating cost if metered", "=B25+Assumptions!C47+Assumptions!C48", INR, "Tokens fall. The owner does not."),
    (27, "Total cost if metered", "=B24+B26", INR, ""),
    (28, "Fee revenue", "=C22*Assumptions!C53", INR, "Closures × the fee. Failed sessions are not billed."),
    (29, "Fee revenue minus Assist cost", "=B28-B26", INR, "Gross margin on the fee, before the tickets that come back."),
    (30, "Margin on the fee", "=IF(B28=0,\"n/a\",(B28-B26)/B28)", PCT, "At ₹49 this is negative. The owner costs more than the fee collects."),
    (31, "Net benefit vs doing nothing", "=(Cost!B23-B27)+B28", INR, "Payroll avoided, minus nothing, plus the fee."),
    (32, "Free Assist net benefit, for comparison", "=B7", INR, "Same definition, no fee, higher adoption."),
    (33, "Free beats the fee by", "=B32-B31", INR, "Positive means I would not turn the meter on."),
]
for row, label, formula, fmt, why in metered:
    billing.cell(row, 1, label).font = label_font
    billing.cell(row, 2, formula)
    apply_formula(billing.cell(row, 2), fmt)
    billing.cell(row, 3, why).font = note_font
    billing.cell(row, 3).alignment = wrap

note(
    billing,
    "A35",
    "At the ₹49 input, fee revenue does not cover the playbook owner, so the margin on the fee is negative. "
    "Raise that blue cell until the margin turns positive and row 33 still loses to giving Assist away, because human tickets come back. Row 33 is the comparison.",
)
billing.merge_cells("A35:C35")
billing.row_dimensions[35].height = 32
billing.freeze_panes = "A4"
billing.sheet_properties.tabColor = "1F4E79"
billing.sheet_view.showGridLines = False
billing.page_setup.orientation = "landscape"
billing.page_setup.fitToPage = True
billing.page_setup.fitToWidth = 1
billing.page_setup.fitToHeight = 1
billing.page_setup.paperSize = billing.PAPERSIZE_A4
billing.oddHeader.left.text = "Cheerio Assist  ·  Billing"

# =============================================================================
# Audit
# =============================================================================
audit = wb.create_sheet("Audit")
set_widths(audit, [62, 16, 78])
audit["A1"] = "Monthly close"
audit["A1"].font = title_font
note(
    audit,
    "A2",
    "Run the structural checks every time the sheet changes. Run the worked-example checks on the first Monday of the month with real numbers typed into the blue cells on Assumptions. "
    "A FAIL on wallet or send actions stops the rollout the same day. A FAIL on a playbook turns off that playbook's button and leaves the explanation in place.",
)
audit.merge_cells("A2:C2")
audit.row_dimensions[2].height = 36

audit["A4"] = "Structural — these should pass before anyone trusts a margin"
audit["A4"].font = section_font
for col, text in enumerate(["Check", "Result", "What it is looking at"], 1):
    audit.cell(5, col, text)
header_row(audit, 5, 3)

structural = [
    (6, "Business segments sum to 100%", '=IF(ABS(Assumptions!B14-1)<0.0001,"PASS","ALERT")', "Assumptions, who raises the queries."),
    (7, "Issue shares sum to 100%", '=IF(ABS(Assumptions!B25-1)<0.0001,"PASS","ALERT")', "Assumptions, what the queries are."),
    (8, "Month-3 roadmap matches the volume sheet", '=IF(Roadmap!B21="PASS","PASS","ALERT")', "Two formulas, one queue. They have to agree."),
    (9, "Month-3 human tickets are below today", '=IF(Volume!G26<Volume!B14,"PASS","ALERT")', "If this fails, Assist is adding work."),
    (10, "Monthly saving is positive", '=IF(Cost!B26>0,"PASS","ALERT")', "Fractional load. The number on Cost row 26."),
    (11, "Free Assist beats the metered fee", '=IF(Billing!B33>0,"PASS","ALERT")', "If this fails, look again at the fee. The recommendation might be wrong."),
    (12, "Token price inputs are present", '=IF(AND(Assumptions!C35>0,Assumptions!C37>0,Assumptions!C38>0),"PASS","ALERT")', "A zero price would fake the margin."),
]
for row, label, formula, why in structural:
    audit.cell(row, 1, label).font = label_font
    audit.cell(row, 2, formula)
    apply_formula(audit.cell(row, 2))
    audit.cell(row, 3, why).font = note_font

audit["A14"] = "This month's close — sample numbers until the first real month"
audit["A14"].font = section_font
for col, text in enumerate(["Check", "Result", "Rule"], 1):
    audit.cell(15, col, text)
header_row(audit, 15, 3)

monthly = [
    (16, "OpenAI invoice vs logged token cost",
     '=IF(Assumptions!C68<=Cost!B13*(1+Assumptions!C63),"PASS","ALERT")',
     "Invoice ≤ logged cost × (1 + tolerance). Logs are Cost!B13. Invoice and tolerance are blue inputs."),
    (17, "Repeat contact within 7 days",
     '=IF(Assumptions!C65<=Assumptions!C64,"PASS","ALERT")',
     "Observed rate ≤ the alert line. Over the line, that playbook loses its one-click button."),
    (18, "Wrong one-click fixes in the sample of 30",
     '=IF(Assumptions!C66=0,"PASS","ALERT")',
     "Any number above zero: turn the button off for that playbook the same day. Advice can stay."),
    (19, "Wallet, plan, or send actions taken by Assist",
     '=IF(Assumptions!C67=0,"PASS","ALERT")',
     "Must be zero. This is a test in the product, not a sample. Non-zero stops rollout."),
]
for row, label, formula, why in monthly:
    audit.cell(row, 1, label).font = label_font
    audit.cell(row, 2, formula)
    apply_formula(audit.cell(row, 2))
    audit.cell(row, 3, why).font = note_font
    audit.cell(row, 3).alignment = wrap
    audit.row_dimensions[row].height = 28

audit["A21"] = "Logged token cost the invoice is checked against"
audit["B21"] = "=Cost!B13"
apply_formula(audit["B21"], INR2)
apply_label(audit["A21"])
audit["C21"] = "Formula. Do not type over it. Type the invoice on Assumptions."
audit["C21"].font = note_font

audit["A23"] = "What a person does on the first Monday"
audit["A23"].font = section_font
steps = [
    (24, "1", "Export Assist sessions: issue, tokens, resolved or handed off, and whether a ticket opened on the same issue within 7 days."),
    (25, "2", "Sum the tokens and compare with the OpenAI invoice. Type the invoice into the blue cell. The check above moves."),
    (26, "3", "Read 30 sessions marked resolved, stratified across the live playbooks. Count wrong one-click fixes, not clumsy wording."),
    (27, "4", "Confirm the automated test: Assist called no wallet, plan, or send endpoint. Type the count."),
    (28, "5", "Look at human tickets per 100 active businesses. If month 2 is not below month 1, do not add a playbook. Repair the one that is live."),
    (29, "6", "Do not bank the headcount saving until that rate has moved. The roster line on Cost is a ceiling for a reason."),
]
for row, n, text in steps:
    audit.cell(row, 1, text).font = label_font
    audit.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
    audit.row_dimensions[row].height = 20

note(
    audit,
    "A31",
    "Alert destination: one channel the support lead already reads. A dashboard nobody opens is not an audit. "
    "The playbook owner is the person named on Assumptions. If that seat is empty, Assist should explain and should not edit.",
)
audit.merge_cells("A31:C31")
audit.row_dimensions[31].height = 32

# conditional formatting for PASS / ALERT
pass_fill = PatternFill("solid", fgColor="D9EFE4")
pass_font = Font(name="Calibri", size=11, color="0E6B4F", bold=True)
alert_fill = PatternFill("solid", fgColor="FDE8DF")
alert_font = Font(name="Calibri", size=11, color="9F2D12", bold=True)
for rng in ("B6:B12", "B16:B19", "Roadmap!B21"):
    # Roadmap range can't be applied from this sheet in one rule easily; handle audit ranges here
    pass

for rng in ("B6:B12", "B16:B19"):
    audit.conditional_formatting.add(
        rng,
        CellIsRule(operator="equal", formula=['"PASS"'], fill=pass_fill, font=pass_font),
    )
    audit.conditional_formatting.add(
        rng,
        CellIsRule(operator="equal", formula=['"ALERT"'], fill=alert_fill, font=alert_font),
    )

roadmap.conditional_formatting.add(
    "B21",
    CellIsRule(operator="equal", formula=['"PASS"'], fill=pass_fill, font=pass_font),
)
roadmap.conditional_formatting.add(
    "B21",
    CellIsRule(operator="equal", formula=['"ALERT"'], fill=alert_fill, font=alert_font),
)

audit.freeze_panes = "A4"
audit.sheet_properties.tabColor = "9F2D12"
audit.sheet_view.showGridLines = False
audit.page_setup.orientation = "landscape"
audit.page_setup.fitToPage = True
audit.page_setup.fitToWidth = 1
audit.page_setup.fitToHeight = 1
audit.page_setup.paperSize = audit.PAPERSIZE_A4
audit.oddHeader.left.text = "Cheerio Assist  ·  Audit"

# =============================================================================
# Summary
# =============================================================================
summary = wb.create_sheet("Summary", 0)
set_widths(summary, [46, 22, 46, 22, 28])
summary.row_dimensions[1].height = 26
summary["A1"] = "Cheerio Assist — the numbers"
summary["A1"].font = title_font
note(
    summary,
    "A2",
    "Start here. Every figure is a formula. Blue inputs live on Assumptions. "
    "North star: human tickets per 100 active businesses. Target for month 3 is the drop in row 8, not a round number chosen first.",
)
summary.merge_cells("A2:E2")
summary.row_dimensions[2].height = 28

summary["A4"] = "The queue"
summary["A4"].font = section_font
summary["C4"] = "The money"
summary["C4"].font = section_font

queue = [
    (5, "Active businesses", "=Assumptions!C5", NUM0),
    (6, "Queries this month", "=Volume!E9", NUM0),
    (7, "Human tickets per 100 businesses, today", "=Volume!B15", "0.0"),
    (8, "Human tickets per 100 businesses, month 3", "=Volume!B28", "0.0"),
    (9, "Drop in human tickets", "=Volume!B29", PCT),
]
for row, label, formula, fmt in queue:
    summary.cell(row, 1, label).font = label_font
    summary.cell(row, 2, formula)
    apply_formula(summary.cell(row, 2), fmt)

money = [
    (5, "Cost of one human ticket", "=Cost!B7", INR2),
    (6, "Token cost of one Assist session", "=Cost!B11", INR2),
    (7, "Full Assist cost per query it closes", "=Cost!B31", INR2),
    (8, "Support cost today", "=Cost!B23", INR),
    (9, "Support cost in month 3", "=Cost!B25", INR),
    (10, "Saving per month", "=Cost!B26", INR),
]
for row, label, formula, fmt in money:
    summary.cell(row, 3, label).font = label_font
    summary.cell(row, 4, formula)
    apply_formula(summary.cell(row, 4), fmt)

summary["A11"] = "Month 1 drop"
summary["B11"] = "=Roadmap!C19"
apply_formula(summary["B11"], PCT)
apply_label(summary["A11"])
summary["A12"] = "Month 2 drop"
summary["B12"] = "=Roadmap!D19"
apply_formula(summary["B12"], PCT)
apply_label(summary["A12"])
summary["C11"] = "Saving per rupee spent on Assist"
summary["D11"] = "=Billing!B9"
apply_formula(summary["D11"], "0.00")
apply_label(summary["C11"])
summary["C12"] = "Free beats a ₹49 fee by"
summary["D12"] = "=Billing!B33"
apply_formula(summary["D12"], INR)
apply_label(summary["C12"])

summary["A14"] = "All structural checks"
summary["B14"] = '=IF(COUNTIF(Audit!B6:B12,"ALERT")+COUNTIF(Audit!B16:B19,"ALERT")=0,"PASS","ALERT")'
apply_formula(summary["B14"])
apply_label(summary["A14"])
summary.conditional_formatting.add(
    "B14",
    CellIsRule(operator="equal", formula=['"PASS"'], fill=pass_fill, font=pass_font),
)
summary.conditional_formatting.add(
    "B14",
    CellIsRule(operator="equal", formula=['"ALERT"'], fill=alert_fill, font=alert_font),
)

# Chart data
summary["A16"] = "Cost comparison for the chart"
summary["A16"].font = section_font
for col, text in enumerate(["", "Human payroll", "Assist operating cost"], 1):
    summary.cell(17, col, text)
header_row(summary, 17, 3)
summary["A18"] = "Today"
summary["B18"] = "=Cost!B23"
summary["C18"] = 0
summary["A19"] = "Month 3, Assist included free"
summary["B19"] = "=Cost!B24"
summary["C19"] = "=Cost!B18"
summary["A20"] = "Month 3, if we charge ₹49"
summary["B20"] = "=Billing!B24"
summary["C20"] = "=Billing!B26"
for row in (18, 19, 20):
    apply_label(summary.cell(row, 1))
    for col in (2, 3):
        apply_formula(summary.cell(row, col), INR)
# C18 is a real zero via formula
summary["C18"] = "=0"
apply_formula(summary["C18"], INR)

chart = BarChart()
chart.type = "col"
chart.grouping = "stacked"
chart.title = "Monthly support cost"
chart.y_axis.title = "INR"
chart.x_axis.title = None
data = Reference(summary, min_col=2, max_col=3, min_row=17, max_row=20)
cats = Reference(summary, min_col=1, min_row=18, max_row=20)
chart.add_data(data, from_rows=False, titles_from_data=True)
chart.set_categories(cats)
chart.shape = 4
chart.legend.position = "b"
chart.y_axis.numFmt = "₹#,##0"
chart.style = 10
chart.overlap = 100
chart.width = 18
chart.height = 8
summary.add_chart(chart, "A22")

note(
    summary,
    "A38",
    "Read the third bar as a warning, not a plan. The fee brings in revenue on the Billing sheet and still costs more overall, because human tickets come back. "
    "Sources worth knowing: OpenAI gpt-4o-mini prices (developers.openai.com/api/docs/pricing, 12 Oct 2026); "
    "USD/INR about 95.8 on 12 Oct 2026; Cheerio list prices and included calls on cheerio.ai/pricing the same day. "
    "Agent salary, query mix, resolution rates, and adoption are assumptions.",
)
summary.merge_cells("A38:E38")
summary.row_dimensions[38].height = 48

summary.freeze_panes = "A4"
summary.sheet_properties.tabColor = "0E3B2E"
summary.sheet_view.showGridLines = False
summary.page_setup.orientation = "landscape"
summary.page_setup.fitToPage = True
summary.page_setup.fitToWidth = 1
summary.page_setup.fitToHeight = 1
summary.page_setup.paperSize = summary.PAPERSIZE_A4
summary.oddHeader.left.text = "Cheerio Assist  ·  Summary"
summary.page_setup.horizontalCentered = True
summary.sheet_view.zoomScale = 120
summary.oddFooter.center.text = "Blue cells are on the Assumptions tab. Everything here is a formula."

# Print order is sheet order. Summary first. Good.
# Give Assumptions a print area that includes notes.
assumptions.print_area = "A1:G68"
summary.print_area = "A1:E38"
volume.print_area = "A1:H33"
cost.print_area = "A1:C42"
billing.print_area = "A1:C35"
roadmap.print_area = "A1:H23"
audit.print_area = "A1:C31"

# Freeze header hint on summary row heights
for row in range(5, 13):
    summary.row_dimensions[row].height = 20

out = "/Users/divyanshu/Desktop/Cheerio Assist-AI-Powered Self-Service Support/Cheerio-Assist-Cost-Model.xlsx"
wb.save(out)
print(out)
