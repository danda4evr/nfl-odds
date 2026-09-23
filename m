import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

FONT = "Arial"
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "NFL Week 4 Leans"

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(name=FONT, bold=True, color="FFFFFF", size=11)
TITLE_FONT = Font(name=FONT, bold=True, size=16, color="1F4E78")
SUB_FONT = Font(name=FONT, italic=True, size=10, color="595959")
INPUT_FONT = Font(name=FONT, color="0000FF")
FORMULA_FONT = Font(name=FONT, color="000000")
LABEL_FONT = Font(name=FONT, bold=True)
GREEN_FILL = PatternFill("solid", fgColor="C6EFCE")
YELLOW_FILL = PatternFill("solid", fgColor="FFF2CC")
GREY_FILL = PatternFill("solid", fgColor="F2F2F2")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

widths = [10, 22, 10, 11, 11, 11, 13, 8, 7, 16, 11, 46]
for i, w in enumerate(widths, start=1):
    ws.column_dimensions[get_column_letter(i)].width = w

ws.merge_cells("A1:L1")
ws["A1"] = "NFL Week 4 \u2014 Odds & Leans"
ws["A1"].font = TITLE_FONT
ws.merge_cells("A2:L2")
ws["A2"] = "Games Thu 9/24 \u2013 Mon 9/28, 2026. Market lines checked 9/23/26 \u2014 lines move, always confirm before betting."
ws["A2"].font = SUB_FONT
ws.row_dimensions[1].height = 24
ws.row_dimensions[2].height = 16

headers = ["Date", "Matchup\n(Away @ Home)", "Kickoff\n(ET)", "Bet365\nML", "FanDuel\nML",
           "Best Price\nYou Found", "Market Ref.\nML (Away/Home)", "Spread", "Total",
           "My Lean", "Confidence", "Why"]
hr = 4
for i, h in enumerate(headers, start=1):
    ws.cell(row=hr, column=i, value=h)
    c = ws.cell(row=hr, column=i)
    c.fill = HEADER_FILL
    c.font = HEADER_FONT
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = BORDER
ws.row_dimensions[hr].height = 32

# date, matchup, kickoff, market ML away/home, spread, total, lean, confidence, why
games = [
    ("Thu 9/24", "Atlanta Falcons @ Green Bay Packers", "8:15 PM",
     "+220 / -250", "GB -5 (-110)", "43", "Green Bay", "Medium-High",
     "Model gives GB ~71% to win. Falcons' offense has been inconsistent; Packers strong at home."),
    ("Sun 9/27", "Carolina Panthers @ Cleveland Browns", "1:00 PM",
     "-148 / +130", "CAR -2.5 (-118)", "42.5", "Carolina", "Medium",
     "Bryce Young/Panthers offense trending up; Browns are one of the league's weaker teams this year."),
    ("Sun 9/27", "Cincinnati Bengals @ Pittsburgh Steelers", "1:00 PM",
     "-184 / +160", "CIN -3.5 (-108)", "42.5", "Cincinnati", "Medium-High",
     "Model (~63%) and market both favor CIN by a similar margin \u2014 no disagreement, just a solid road favorite."),
    ("Sun 9/27", "Houston Texans @ Indianapolis Colts", "1:00 PM",
     "-148 / +130", "HOU -2.5 (-115)", "42.5", "Houston", "Medium-High",
     "Model (~57%) agrees with the market line favoring Houston on the road."),
    ("Sun 9/27", "LA Chargers @ Buffalo Bills", "1:00 PM",
     "+300 / -335", "BUF -7 (-110)", "50", "Buffalo", "High",
     "Model (~75%) and market both strongly favor Buffalo at home \u2014 the chalk play of the week."),
    ("Sun 9/27", "New England Patriots @ Jacksonville Jaguars", "1:00 PM",
     "+135 / -155", "JAC -3 (-105)", "45.5", "Jacksonville", "Medium",
     "Model (~58%) and market agree JAC is the play at home; not a huge edge either way."),
    ("Sun 9/27", "New York Jets @ Detroit Lions", "1:00 PM",
     "+260 / -300", "DET -6.5 (-110)", "47.5", "Detroit", "Medium-High",
     "Model (~73%) backs Detroit comfortably; Jets have been shaky on offense."),
    ("Sun 9/27", "Seattle Seahawks @ Washington Commanders", "1:00 PM",
     "-298 / +270", "SEA -7 (-105)", "40", "Seattle", "High",
     "Seattle is the ROAD favorite here by a wide margin \u2014 model (~76%) and market fully agree."),
    ("Sun 9/27", "Tennessee Titans @ New York Giants", "1:00 PM",
     "+128 / -135", "NYG -3 (+100)", "39.5", "NY Giants", "Medium",
     "Model (~57%) lines up almost exactly with the market price \u2014 fair value, modest lean."),
    ("Sun 9/27", "Arizona Cardinals @ San Francisco 49ers", "4:05 PM",
     "+370 / -400", "SF -8.5 (-105)", "48", "San Francisco", "Medium-High",
     "Large home favorite; no model number available for this one, but the market is lopsided for a reason."),
    ("Sun 9/27", "Minnesota Vikings @ Tampa Bay Buccaneers", "4:05 PM",
     "-120 / +106", "MIN -1.5 (-108)", "42.5", "Toss-up \u2014 lean TB (+106)", "Low",
     "Essentially a coin flip. If betting at all, the home dog at plus money is the value side; smallest stake."),
    ("Sun 9/27", "Baltimore Ravens @ Dallas Cowboys", "4:25 PM",
     "-176 / +155", "BAL -3.5 (-102)", "52.5", "Baltimore", "Medium",
     "Ravens favored on the road; CHECK Lamar Jackson's injury/practice status before kickoff \u2014 it swings this line."),
    ("Sun 9/27", "Las Vegas Raiders @ New Orleans Saints", "4:25 PM",
     "+154 / -168", "NO -3 (-112)", "44", "New Orleans", "Medium",
     "Saints have looked more organized on both sides of the ball recently; modest home favorite."),
    ("Sun 9/27", "LA Rams @ Denver Broncos", "8:20 PM",
     "-137 / +120", "LA -2.5 (-112)", "45.5", "LA Rams", "Medium",
     "Rams favored on the road in a fairly close, no-strong-signal matchup."),
    ("Mon 9/28", "Philadelphia Eagles @ Chicago Bears", "8:15 PM",
     "-225 / +194", "PHI -4.5 (-107)", "41", "Philadelphia", "Medium-High",
     "Eagles have been the more complete team this season and are sizable road favorites here."),
]

first = hr + 1
conf_fill = {"High": GREEN_FILL, "Medium-High": GREEN_FILL, "Medium": YELLOW_FILL, "Low": GREY_FILL}

for i, row in enumerate(games):
    r = first + i
    date, matchup, kickoff, mkt_ml, spread, total, lean, conf, why = row
    ws.cell(row=r, column=1, value=date).font = FORMULA_FONT
    ws.cell(row=r, column=2, value=matchup).font = FORMULA_FONT
    ws.cell(row=r, column=3, value=kickoff).font = FORMULA_FONT
    # Bet365 / FanDuel: blank blue input cells
    ws.cell(row=r, column=4).font = INPUT_FONT
    ws.cell(row=r, column=5).font = INPUT_FONT
    d_ref, e_ref = f"D{r}", f"E{r}"
    ws.cell(row=r, column=6, value=(
        f'=IF(AND({d_ref}="",{e_ref}=""),"",IF({d_ref}="",{e_ref},IF({e_ref}="",{d_ref},MAX({d_ref},{e_ref}))))'))
    ws.cell(row=r, column=6).font = FORMULA_FONT
    ws.cell(row=r, column=7, value=mkt_ml).font = FORMULA_FONT
    ws.cell(row=r, column=8, value=spread).font = FORMULA_FONT
    ws.cell(row=r, column=9, value=total).font = FORMULA_FONT
    ws.cell(row=r, column=10, value=lean).font = Font(name=FONT, bold=True)
    conf_cell = ws.cell(row=r, column=11, value=conf)
    conf_cell.font = FORMULA_FONT
    conf_cell.fill = conf_fill.get(conf, GREY_FILL)
    conf_cell.alignment = Alignment(horizontal="center")
    why_cell = ws.cell(row=r, column=12, value=why)
    why_cell.font = Font(name=FONT, size=9)
    why_cell.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[r].height = 44
    for c in range(1, 13):
        ws.cell(row=r, column=c).border = BORDER

last = first + len(games) - 1

# highlight best price cell green when D or E matches it (simple line-shopping visual)
for r in range(first, last + 1):
    for col_letter in ("D", "E"):
        ref = f"{col_letter}{r}"
        ws.conditional_formatting.add(
            ref, FormulaRule(formula=[f'AND({ref}<>"",{ref}=F{r})'], fill=GREEN_FILL))

dv = DataValidation(type="list", formula1='"High,Medium-High,Medium,Low"', allow_blank=True)
ws.add_data_validation(dv)
dv.add(f"K{first}:K{last}")

ws.freeze_panes = f"A{first}"

# Legend
lr = last + 3
ws.cell(row=lr, column=1, value="Legend").font = Font(name=FONT, bold=True, size=12)
legend = [
    ("Blue cells (Bet365/FanDuel)", "Paste the exact current line from each site \u2014 these are the only two things you type in."),
    ("Best Price You Found", "Auto-highlights whichever of your two pasted numbers is better once you fill both in."),
    ("Market Ref. ML", "The moneyline I found via research as of 9/23/26, for context \u2014 not guaranteed to match either book exactly."),
    ("Confidence", "How much the model/market/context line up behind the lean, not a guarantee \u2014 sports are inherently uncertain."),
]
for i, (a, b) in enumerate(legend):
    ws.cell(row=lr+1+i, column=1, value=a).font = LABEL_FONT
    c2 = ws.cell(row=lr+1+i, column=2, value=b)
    c2.alignment = Alignment(wrap_text=True)
    ws.merge_cells(start_row=lr+1+i, start_column=2, end_row=lr+1+i, end_column=8)

wb.save("/home/claude/nfl_week4_leans.xlsx")
print("saved")
