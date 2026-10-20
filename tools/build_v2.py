#!/usr/bin/env python3
"""build_v2.py - builds deck/CFO_CoS_deck_v2.pptx from deck/CFO_CoS_deck_repaired.pptx.

Edits in priority order (core arguments first, then structure, then polish). Reproducible: run it again to rebuild.
usage: python3 -I tools/build_v2.py [src] [dst]
"""
import copy
import sys

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

SRC = sys.argv[1] if len(sys.argv) > 1 else "deck/CFO_CoS_deck_repaired.pptx"
DST = sys.argv[2] if len(sys.argv) > 2 else "deck/CFO_CoS_deck_v2.pptx"

PURPLE = RGBColor(0x82, 0x0A, 0xD1)
INK = RGBColor(0x11, 0x05, 0x2C)
BODY = RGBColor(0x3B, 0x2A, 0x54)
MUTED = RGBColor(0x6B, 0x5B, 0x82)
LILAC = RGBColor(0xE2, 0xD9, 0xF3)
ORANGE = RGBColor(0xD9, 0x77, 0x06)
GREY = RGBColor(0x6B, 0x5B, 0x82)

prs = Presentation(SRC)
S = list(prs.slides)  # original order, 1-based helpers below


def sh(slide, name):
    for s in slide.shapes:
        if s.name == name:
            return s
    raise KeyError(name)


def _first_run(tf):
    for p in tf.paragraphs:
        for r in p.runs:
            return p, r
    p = tf.paragraphs[0]
    return p, p.add_run()


def set_rich(tf, paragraphs, size=None):
    """paragraphs: list of paragraphs; each is a str or a list of (text, bold) segments. Keeps the first run's formatting."""
    p0, r0 = _first_run(tf)
    rpr0 = copy.deepcopy(r0._r.get_or_add_rPr())
    ppr0 = copy.deepcopy(p0._p.get_or_add_pPr())
    # wipe
    for p in list(tf.paragraphs):
        tf._txBody.remove(p._p)
    for para in paragraphs:
        p_el = copy.deepcopy(p0._p)
        for child in list(p_el):
            if child.tag != qn("a:pPr"):
                p_el.remove(child)
        tf._txBody.append(p_el)
        segs = [(para, None)] if isinstance(para, str) else para
        for text, bold in segs:
            r_el = copy.deepcopy(r0._r)
            for child in list(r_el):
                r_el.remove(child)
            rpr = copy.deepcopy(rpr0)
            if bold is not None:
                rpr.set("b", "1" if bold else "0")
            if size is not None:
                rpr.set("sz", str(int(size * 100)))
            r_el.append(rpr)
            t = r_el.makeelement(qn("a:t"), {})
            t.text = text
            r_el.append(t)
            p_el.append(r_el)


def set_text(shape_or_cell, text, size=None):
    tf = shape_or_cell.text_frame
    set_rich(tf, text if isinstance(text, list) else [text], size=size)


def cell(table, r, c):
    return table.cell(r, c)


def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


def clone_slide(src):
    new = prs.slides.add_slide(src.slide_layout)
    for s in list(new.shapes):
        s._element.getparent().remove(s._element)
    for s in src.shapes:
        if s.shape_type == 13:
            continue  # pictures need rels; none cloned
        new.shapes._spTree.append(copy.deepcopy(s._element))
    bg = src._element.cSld.find(qn("p:bg"))
    if bg is not None:
        new._element.cSld.insert(0, copy.deepcopy(bg))
    return new


def add_table_row(table, copy_from=-1):
    tbl = table._tbl
    tr = copy.deepcopy(tbl.tr_lst[copy_from])
    tbl.append(tr)
    return len(tbl.tr_lst) - 1


def delete_slide(slide):
    sldIdLst = prs.slides._sldIdLst
    for sldId in list(sldIdLst):
        if prs.slides.part.related_part(sldId.rId) is slide.part if hasattr(prs.slides.part, "related_part") else False:
            pass
    # python-pptx: find by rId lookup
    for sldId in list(sldIdLst):
        if prs.part.related_part(sldId.rId) == slide.part:
            prs.part.drop_rel(sldId.rId)
            sldIdLst.remove(sldId)
            return


def move_slide(slide, new_index):
    sldIdLst = prs.slides._sldIdLst
    for sldId in list(sldIdLst):
        if prs.part.related_part(sldId.rId) == slide.part:
            sldIdLst.remove(sldId)
            sldIdLst.insert(new_index, sldId)
            return


def add_textbox(slide, left, top, width, height, text, size=13.5, bold=False, color=BODY, align=PP_ALIGN.LEFT, font="DM Sans"):
    tb = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = tb.text_frame
    tf.word_wrap = True
    paras = text if isinstance(text, list) else [text]
    first = True
    for para in paras:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        segs = [(para, bold)] if isinstance(para, str) else para
        for t, b in segs:
            r = p.add_run()
            r.text = t
            r.font.size = Pt(size)
            r.font.bold = bool(b)
            r.font.name = font
            r.font.color.rgb = color
    return tb


# ============================================================ 1. Title
s = S[0]
set_text(sh(s, "TextBox 6"), "Chief of Staff to the CFO: a planning engine, a Board rhythm and an HQ charter")
set_text(sh(s, "TextBox 7"), "Case resolution, 60 days into the role. Prepared by Cassio Modelli for the Nu panel, October 2026. "
         "Nu figures are from public Q2'26 filings; my assumptions are listed on slide 3.")
set_text(sh(s, "TextBox 11"), "Bring planning in-house and answer $60M of growth asks inside a $25M cost cap.")
set_text(sh(s, "TextBox 14"), "Settle two opposing capital views and get the Board book out on time.")
set_text(sh(s, "TextBox 17"), "Lead legal, tax and campus terms across jurisdictions without line authority.")
set_text(sh(s, "TextBox 8"), "Four deliverables: (1) Planning and prioritization engine  (2) C-suite alignment and Board rhythm  "
         "(3) HQ execution roadmap  (4) How I lead")
notes(s, "Takeaway: three live problems, four deliverables, one ask at the end. "
         "Hardest question: 'Why should we believe a 60-day-old view?' Answer: because every number here is either a public Nu figure "
         "with its source or an assumption I list on slide 3 and would verify in week one.")

# ============================================================ 2. Executive summary and ask
s = S[1]
set_text(sh(s, "TextBox 7"), "Fund FY27's $25M through two business-case gates, lock the Board plan 16 days out, and charter the HQ project this week")
set_text(sh(s, "TextBox 9"), "$25M in two gates")
set_text(sh(s, "TextBox 10"), "Every ask gets a business case; $15M at FY27 start, $10M at Q1 close on indicators. No flat haircuts.")
set_text(sh(s, "TextBox 12"), "T-16 lock, T-7 book")
set_text(sh(s, "TextBox 13"), "One options memo settles the capital dispute at T-16; the Board book ships at T-7 with written Q&A.")
set_text(sh(s, "TextBox 15"), "4 gates, 120 days")
set_text(sh(s, "TextBox 16"), "Tax substance, grant terms, officer liability and lease, each with an owner and an exit test.")
set_text(sh(s, "TextBox 17"), "DELIVERABLE 4 // LEADERSHIP")
set_text(sh(s, "TextBox 18"), "Rigor in the squads")
set_text(sh(s, "TextBox 19"), "I work inside the business units with the numbers, not above them with a veto.")
set_text(sh(s, "TextBox 20"), "THE ASK")
set_text(sh(s, "TextBox 21"), "Approve the $25M FY27 cap in two gates, with a CFO- and IR-gated valve of up to $15M that is not in plan; "
         "the T-16 / T-7 Board calendar; and the HQ project charter. I report on all three at Day 90.", size=15)
notes(s, "Takeaway: three decisions, each with a number. "
         "Hardest question: 'Why $25M and not a reallocation inside the existing base?' Answer: the cap is the CFO's constraint from the "
         "FY26 efficiency-ratio guide of about 20%; the model reallocates inside it every quarter, which is the point.")

# ============================================================ 3. (new) Assumptions slide, cloned from the scenario table slide (old 14)
assump = clone_slide(S[13])
set_text(sh(assump, "TextBox 26"), "02 // WHAT I AM ASSUMING, AND WHAT I WOULD VERIFY IN WEEK ONE")
set_text(sh(assump, "TextBox 27"), "Key assumptions, separated from the facts they sit on")
t = sh(assump, "Table 25").table if any(x.name == "Table 25" for x in assump.shapes) else [x for x in assump.shapes if x.has_table][0].table
hdr = ["#", "Assumption or fact", "Class", "Basis"]
rows = [
    ["A1", "Nu Q2'26: 138.9M customers, monthly ARPAC $17.1, cost to serve $1.0, efficiency ratio 19.5%, ROE 33%, net income $1.06B.", "FACT", "Q2'26 earnings release and 6-K"],
    ["A2", "Management expects the FY26 efficiency ratio near 20% and guided US spend at no more than about 100 bps of it.", "FACT (wording to re-check)", "Q2'26 earnings call"],
    ["A3", "The $60M asks and the $25M cap are the case's figures. At Nu's scale $35M is about 15 bps of annualized revenue; the stakes are the operating model and the precedent, not the amount.", "ASSUMPTION", "Case brief; my calculation on Q2'26 revenue"],
    ["A4", "The two MTeam members disagree about the growth-opex envelope (core markets vs new bets), not about regulatory capital by entity.", "ASSUMPTION", "Case brief is silent; verify in week one"],
    ["A5", "$8M of the asks can be met on the shared global platform at near-zero marginal cost; $12M lack a business case today and can be prepared in 60 days.", "ASSUMPTION", "My estimate; verify with Engineering and the GMs"],
    ["A6", "Any spend above $25M is released only from H1 profit above plan, sized so the efficiency ratio stays inside the FY26 guide, and only with CFO and IR agreement.", "RULE (proposed)", "To agree with the CFO"],
    ["A7", "Board meeting in 4 weeks; Investor Day 8 Dec 2026. The Board approves the international capital plan before the Investor Day narrative is locked.", "FACT (dates) + ASSUMPTION (sequence)", "Case brief; Nu IR"],
    ["A8", "The HQ jurisdiction is open. Nu's revenue is above the EUR 750M Pillar Two threshold; the 2026 transitional substance carve-out is 9.4% of payroll and 7.4% of tangible assets.", "FACT (rule) + ASSUMPTION (site)", "OECD GloBE rules; Q2'26 revenue"],
    ["A9", "Option scores and weights on slide 6 are my judgment; third-party benchmarks are directional, never load-bearing.", "ASSUMPTION", "Slide 6; ledger"],
    ["A10", "The cap is the FY27 growth-opex envelope: gate 1 opens at FY27 start (1 Jan 2027), gate 2 at Q1 close (31 Mar) on Q1 indicators, the A6 valve at H1 close. The hub needs 6-8 people.", "ASSUMPTION (fiscal calendar, sizing)", "Nu reports on a calendar fiscal year; sizing is mine"],
]
# table currently has 1 header + 5 body rows (6); add rows to reach 1 + 9
while len(t.rows) < 1 + len(rows):
    add_table_row(t)
for c, h in enumerate(hdr):
    set_text(cell(t, 0, c), h)
for r, row in enumerate(rows, start=1):
    for c, val in enumerate(row):
        set_text(cell(t, r, c), [[(val, c in (0, 2))]], size=10)
# column widths: #, statement, class, basis
widths = [0.5, 7.5, 1.7, 2.3]
for c, w in enumerate(widths):
    t.columns[c].width = Inches(w)
for r in range(len(t.rows)):
    t.rows[r].height = Inches(0.44 if r else 0.36)
tbl_shape = [x for x in assump.shapes if x.has_table][0]
tbl_shape.top = Inches(1.3)
tbl_shape.height = Inches(0.36 + 0.44 * len(rows))
# remove the decorative separator rectangles cloned from the scenario slide (they no longer line up)
for x in list(assump.shapes):
    if x.shape_type == 1 and x.name.startswith("Rectangle"):
        x._element.getparent().remove(x._element)
set_text(sh(assump, "TextBox 28"), [[("Week one: ", True), ("the real MTeam positions; the Board and committee calendar; who owns the planning model today; local-CFO decision rights in Mexico and Colombia; the HQ project's current state (sites, counterparties, anything signed); the planning cycle time and forecast accuracy today.", False)]])
sh(assump, "TextBox 28").top = Inches(6.62)
set_text(sh(assump, "TextBox 28"), [[("Week one: ", True), ("the real MTeam positions; the Board and committee calendar; who owns the planning model today; local-CFO decision rights in Mexico and Colombia; the HQ project's current state; today's planning cycle time and forecast accuracy.", False)]], size=11)
notes(assump, "Takeaway: facts and assumptions are kept apart on purpose; the brief asked for it and the CFO will test it. "
      "Hardest question: 'Which assumption, if wrong, breaks the plan?' Answer: A4. If the dispute is about regulatory capital by entity, "
      "slide 12 becomes the main event and the opex split is secondary.")

# ============================================================ 4. Scale challenge (old 3)
s = S[2]
set_text(sh(s, "TextBox 5"), "03 // WHY THE CAP MATTERS AT NU'S SCALE")
set_text(sh(s, "TextBox 6"), "The $35M gap is small against a ~$23B revenue base; the precedent and the FY26 efficiency guide are what matter")
set_text(sh(s, "TextBox 8"), "THE GAP")
set_text(sh(s, "TextBox 10"), "$60M of asks vs a $25M cost cap: about 15 bps of annualized revenue, under 1% of ~$4.6B opex")
set_text(sh(s, "TextBox 11"), "1. Keep the efficiency ratio near 20% [1]")
set_text(sh(s, "TextBox 12"), "Q2'26 came in at 19.5% (17.6% in Q1), with international expansion named as one driver. Management expects about 20% for FY26. Unvetted overhead is what moves that number the wrong way.")
set_text(sh(s, "TextBox 13"), "2. Protect the cost-to-serve edge [1]")
set_text(sh(s, "TextBox 14"), "Monthly cost to serve per active customer is $1.0 (Q2'26, up from $0.8 a year earlier) against $17.1 of monthly ARPAC. The margin is wide; it stays wide only if central costs scale slower than customers.")
set_text(sh(s, "TextBox 15"), "3. Set the precedent for the first in-house cycle")
set_text(sh(s, "TextBox 16"), "This is the first budget cycle planned centrally. How the $35M is closed decides whether GMs bring real cases next year or pad their asks. The process is the product.")
set_text(sh(s, "TextBox 7"), "Sources: [1] Nu Holdings Q2'26 earnings release and 6-K (138.9M customers; cost to serve $1.0; ARPAC $17.1; efficiency ratio 19.5%; FY26 ~20% per the Q2'26 call). Materiality: my calculation on annualized Q2'26 managerial revenue of ~$23.5B.")
notes(s, "Takeaway: do not oversell $35M to a CFO who reports $5.9B a quarter; sell the precedent. "
      "Hardest question: 'If it is immaterial, why not just approve the $60M?' Answer: because the efficiency ratio is guided at ~20% and "
      "every unvetted central dollar is a dollar that is not in a customer-facing cohort.")

# ============================================================ 5. Test, learn, scale (old 4)
s = S[3]
set_text(sh(s, "TextBox 5" if any(x.name == "TextBox 5" for x in s.shapes) else "TextBox 4"), "04 // DELIVERABLE 1: THE PLANNING PRINCIPLE")
title = [x for x in s.shapes if x.has_text_frame and x.text_frame.text.startswith("Use Nu")][0]
set_text(title, "Budgets should be underwritten like credit lines: start small, gate on data, raise fast when the cohort proves out")
box = [x for x in s.shapes if x.has_text_frame and x.text_frame.text.startswith("Why This Fits")][0]
set_text(box, [[("Why this fits Nu: ", True), ("in credit, Nu starts customers on small limits and raises them on repayment behavior. Applying the same logic to budgets replaces one annual negotiation with quarterly evidence. It is Nu's own test-and-learn discipline, applied to Nu's own money.", False)]])
for x in s.shapes:
    if not x.has_text_frame:
        continue
    tx = x.text_frame.text
    if tx.startswith("2. DYNAMIC"):
        set_text(x, "2. WHAT CHANGES FOR THE CFO")
    elif tx.startswith("+30% Higher"):
        set_text(x, "Evidence replaces negotiation")
    elif tx.startswith("McKinsey research"):
        set_text(x, "Each release decision reads the same four indicators for every unit: activation, 30-day retention, CAC vs plan, unit cost vs plan. Reallocation across units happens every quarter, not once a year.")
    elif tx.startswith("Nu's 20-F filings"):
        set_text(x, "Nu's 20-F describes low-and-grow underwriting: small initial limits, raised as repayment data comes in. Losses stay controlled because exposure grows only behind evidence.")
    elif tx.startswith("Instead of fighting"):
        set_text(x, "GMs stop fighting for one annual cheque. They get a funded Q1 start and a published rule for unlocking the next stage, so the conversation moves from positions to indicators.")
    elif tx.startswith("Data & Implications"):
        set_text(x, "Sources: [1] Nu Holdings Form 20-F (credit underwriting; 'low-and-grow'). [2] Directional only, not load-bearing: multi-year studies of active capital reallocation (e.g. McKinsey, 1,600 US companies) associate it with higher shareholder returns.")
notes(s, "Takeaway: this is the one idea the rest of the deck hangs on. "
      "Hardest question: 'Credit lines have repayment data in 30 days; a new vertical does not. What do you gate on?' Answer: leading "
      "indicators (activation, retention, CAC, unit cost) in the first quarter; payback is the lagging confirmation at the second gate.")

# ============================================================ 6. Options (old 5)
s = S[4]
set_text(sh(s, "TextBox 31"), "05 // DELIVERABLE 1: THREE WAYS TO CLOSE $35M")
set_text(sh(s, "TextBox 32"), "Three ways to close $35M: staged readiness wins on return and ownership; flat cuts win only on speed")
tbl = [x for x in s.shapes if x.has_table][0].table
set_text(cell(tbl, 1, 0), [[("1. Equal percentage cut", True), (" (every unit gets 42% of its ask)", False)]])
set_text(cell(tbl, 2, 0), [[("2. Core markets first", True), (" (fund Brazil, Mexico, Colombia; pause new bets)", False)]])
set_text(cell(tbl, 3, 0), [[("3. Business-case readiness + staged release", True), (" (recommended)", False)]])
set_text(sh(s, "TextBox 34"), [[("How to read it: ", True), ("scores and weights are my judgment (assumption A9). The ranking does not change if any weight moves by 10 points. Option 1 loses on return and ownership: across-the-board cuts teach GMs to pad next year's ask and starve launches below viable scale. Option 2 is honest but gives the new bets nothing to earn.", False)]])
set_text(sh(s, "TextBox 33"), "Scores 0-100 per criterion; weighted total = sum of score x weight. Re-run with each weight +/-10 points: option 3 stays first in every case.")
notes(s, "Takeaway: show the options you rejected and why. "
      "Hardest question: 'Your scores are made up.' Answer: yes, they are judgment, labeled as such; the test that matters is that "
      "the winner survives every +/-10 point weight change, and it does.")

# ============================================================ 7. Bridge (old 6): native chart replaces the picture
s = S[5]
for x in list(s.shapes):
    if x.shape_type == 13:
        x._element.getparent().remove(x._element)
set_text(sh(s, "TextBox 2") if any(x.name == "TextBox 2" for x in s.shapes) else [x for x in s.shapes if x.has_text_frame][0],
         "06 // DELIVERABLE 1: THE $60M BRIDGE")
title = [x for x in s.shapes if x.has_text_frame and x.text_frame.text.startswith("Sequence")][0]
set_text(title, "$60M of asks become $25M of funded spend: $8M delivered on the platform, $12M sent back for a case, $15M conditional")
cap = [x for x in s.shapes if x.has_text_frame and x.text_frame.text.startswith("Zero Arbitrary")][0]
set_text(cap, [[("No arbitrary cuts: ", True), ("$8M of country asks are met by reusing the shared global code base at near-zero marginal cost (A5); $12M of early ideas enter a 60-day business-case track with my team; $15M is conditional on H1 profit above plan under rule A6. $25M is funded now in two gates.", False)]])
cd = CategoryChartData()
cd.categories = ["GM asks", "Delivered on the shared platform", "Sent back for a 60-day case", "Conditional on H1 (rule A6)", "Funded now (2 gates)"]
cd.add_series("base", (0, 52, 40, 25, 0))
cd.add_series("value", (60, 8, 12, 15, 25))
gf = s.shapes.add_chart(XL_CHART_TYPE.BAR_STACKED, Inches(0.9), Inches(1.75), Inches(11.5), Inches(4.4), cd)
ch = gf.chart
ch.has_legend = False
ch.has_title = False
ca = ch.category_axis
ca.reverse_order = True
ca.tick_labels.font.size = Pt(12)
ca.tick_labels.font.name = "DM Sans"
ca.tick_labels.font.color.rgb = INK
ca.format.line.fill.background()
va = ch.value_axis
va.visible = False
va.has_major_gridlines = False
va.maximum_scale = 62
plot = ch.plots[0]
plot.gap_width = 60
plot.overlap = 100
base, val = plot.series
base.format.fill.background()
base.format.line.fill.background()
colors = [RGBColor(0x5B, 0x4B, 0x7A), RGBColor(0x8B, 0x5C, 0xF6), RGBColor(0x8B, 0x5C, 0xF6), ORANGE, PURPLE]
labels = ["$60M", "-$8M", "-$12M", "-$15M", "$25M"]
for i, pt in enumerate(val.points):
    pt.format.fill.solid()
    pt.format.fill.fore_color.rgb = colors[i]
    dl = pt.data_label
    dl.has_text_frame = True
    dl.text_frame.text = labels[i]
    dl.position = XL_LABEL_POSITION.INSIDE_END  # stacked bars allow ctr, inEnd, inBase only
    r = dl.text_frame.paragraphs[0].runs[0]
    r.font.size = Pt(13)
    r.font.bold = True
    r.font.name = "DM Sans"
    r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
notes(s, "Takeaway: the whole $60M is answered; nothing is 'cut', each piece has a path. "
      "Hardest question: 'The $8M platform saving is convenient. Where does it come from?' Answer: it is my estimate (A5) from the asks "
      "that duplicate features the global code base already ships; Engineering confirms it in week one or the number moves to the case track.")

# ============================================================ 8. Hub-and-spoke and gates (old 7)
s = S[6]
set_text(sh(s, "TextBox 3"), "07 // DELIVERABLE 1: THE OPERATING MODEL AND THE RELEASE RULE")
set_text(sh(s, "TextBox 4"), "Hub-and-spoke: one model and one calendar centrally, finance partners inside each market and vertical")
set_text(sh(s, "TextBox 7"), "One source of truth, local speed")
set_text(sh(s, "TextBox 11"), [[("Central planning hub (6-8 people, A10): ", True), ("takes the annual cycle and the single model in-house from the business units; reads Nu's data platform, not spreadsheets. KPIs: cycle time under 40 days, forecast accuracy, 100% of spend with an approved case.", False)]], size=12.5)
set_text(sh(s, "TextBox 13"), [[("Embedded finance partners: ", True), ("solid line to the hub, dotted line to the GM; in place by Day 95. They build the cases with the GMs and own the indicator read-outs at each gate.", False)]], size=12.5)
set_text(sh(s, "TextBox 9"), "Two gates inside the cap, one valve above it")
set_text(sh(s, "TextBox 15"), [[("Gate 1, $15M at FY27 start (1 Jan): ", True), ("released to the vetted cases ($10.5M core markets, $4.5M new bets).", False)]], size=12.5)
set_text(sh(s, "TextBox 17"), [[("Gate 2, $10M at Q1 close (31 Mar): ", True), ("completes the $25M cap; released to units whose Q1 indicators (activation, 30-day retention, CAC and unit cost vs plan) hold.", False)]], size=12.5)
set_text(sh(s, "TextBox 19"), [[("The A6 valve: ", True), ("up to $15M in H2, only from H1 profit above plan, sized to stay inside the efficiency guide, with CFO and IR sign-off. Not in the base plan; disclosed as upside spend, not a change to guidance.", False)]], size=12.5)
set_text(sh(s, "TextBox 5"), "Sources: Nu Q2'26 release (efficiency ratio 19.5%; FY26 guide ~20% per the call). Cycle-time target is mine; the directional benchmark for central planning hubs (Hackett Group) is in the ledger and not load-bearing.")
notes(s, "Takeaway: the cap is hard; anything above it is a rule with a formula, not a hope. "
      "Hardest question: 'Rule A6 is a loophole GMs will game.' Answer: it is funded only from profit that is already in the bank, sized "
      "by formula, and the CFO and IR sign it; a GM cannot trigger it.")

# ============================================================ 9. Cadence table (old 12) edited, will be moved after slide 8
s = S[11]
set_text([x for x in s.shapes if x.has_text_frame and x.text_frame.text.startswith("11 //")][0], "08 // DELIVERABLE 1: THE CADENCE")
title = [x for x in s.shapes if x.has_text_frame and x.text_frame.text.startswith("How Our Team")][0]
set_text(title, "Four meetings, each with an owner and an output; nothing else is added")
tbl = [x for x in s.shapes if x.has_table][0].table
tshape = [x for x in s.shapes if x.has_table][0]
tshape.top = Inches(1.6)
for x in list(s.shapes):
    if x.shape_type == 1 and x.name.startswith("Rectangle"):
        x._element.getparent().remove(x._element)
set_text(cell(tbl, 0, 2), "What happens")
set_text(cell(tbl, 0, 3), "Owner and output")
set_text(cell(tbl, 1, 2), "Finance partners join the country and vertical stand-ups and model unit economics with the GMs before anything reaches formal review.")
set_text(cell(tbl, 1, 3), "Finance partner. Output: a business case in the standard template, co-signed by the GM.")
set_text(cell(tbl, 2, 2), "I meet each GM privately to stress-test assumptions and design the staged option before it is seen by the CFO.")
set_text(cell(tbl, 2, 3), "Me. Output: no surprises at the T-16 lock; every ask has a conditional-yes path.")
set_text(cell(tbl, 3, 0), [[("3. HQ pod", True), (" (Deliverable 3)", False)]])
set_text(cell(tbl, 3, 2), "Legal, Tax, Workplace and People leads around one decision log with 48-hour escalation to the CFO.")
set_text(cell(tbl, 3, 3), "Me (chair). Output: the gate pack for the next HQ gate, with owner and exit test.")
set_text(cell(tbl, 4, 2), "Read the four indicators per unit against plan and release or hold the next stage.")
set_text(cell(tbl, 4, 3), "CFO decides on my recommendation. Output: release memo, published to the GMs the same day.")
foot = [x for x in s.shapes if x.has_text_frame and x.text_frame.text.startswith("Data & Implications")][0]
set_text(foot, "Decision roles follow a RAPID-style split: finance partners recommend, GMs provide input, I assemble the pack, the CFO decides, the hub performs.")
foot.top = Inches(6.95)
notes(s, "Takeaway: four meetings, each with an owner and an output; nothing else is added to the calendar. "
      "Hardest question: 'How many people is this?' Answer: the hub is small (model, template, pack); the partners are the finance people "
      "already attached to markets, re-pointed at cases and gates.")

# ============================================================ 10. Core vs new bets (old 8)
s = S[7]
set_text([x for x in s.shapes if x.has_text_frame and x.text_frame.text.startswith("07 //")][0], "09 // DELIVERABLE 2: SETTLING THE TWO VIEWS")
title = [x for x in s.shapes if x.has_text_frame and x.text_frame.text.startswith("Settle")][0]
set_text(title, "Core markets and new bets are judged on different indicators, so both executives can sign the FY27 opex lock")
for x in s.shapes:
    if not x.has_text_frame:
        continue
    tx = x.text_frame.text
    if tx.startswith("HORIZON 2"):
        set_text(x, "HORIZON 2 & 3: NEW BETS (NEW VERTICALS AND MARKETS)")
    elif tx.startswith("70% of Q1"):
        set_text(x, "70% of the gate-1 release ($10.5M)")
    elif tx.startswith("30% of Q1"):
        set_text(x, "30% of the gate-1 release ($4.5M)")
    elif tx.startswith("Nu filings show mature"):
        set_text(x, "Mature cohorts reach ~$31 of monthly ARPAC at eight years vs $17.1 for the group, at 33% ROE (Q2'26). Proven markets get the first, largest release [1].")
    elif tx.startswith("HBR's Three-Horizons"):
        set_text(x, "A new bet judged on mature-market profit in its first quarter always loses. New bets are graded on activation, retention and unit cost until a cohort can show payback.")
    elif tx.startswith("Full-Year Path: $16.5M"):
        set_text(x, "Full year: $16.5M in the cap + up to $5.0M (rule A6)", size=13)
        x.text_frame.word_wrap = True
    elif tx.startswith("Full-Year Path: $8.5M"):
        set_text(x, "Full year: $8.5M in the cap + up to $10.0M (rule A6)", size=13)
        x.text_frame.word_wrap = True
    elif tx.startswith("Why Both Executives Agree"):
        set_text(x, [[("Why both can sign: ", True), ("core markets get 70% of the first release to protect the year; new bets get a protected $4.5M start plus first claim on up to $14M across gate 2 and the A6 valve, if it opens. I write the one-page options memo; the CFO recommends; the CEO decides; the Board chair gets the pre-read. This is opex (A4); equity for Mexico and the US charter is the Board's own decision, next slide but one.", False)]], size=12.5)
    elif tx.startswith("Data & Implications"):
        set_text(x, "Sources: [1] Nu Q2'26 release (ARPAC $17.1; ROE 33%); Q1'26 earnings presentation cohort slide ($30.8 monthly ARPAC at 96 months since origination vs $15.9 total). The horizon split follows the three-horizons framing (Baghai, Coley and White), used here as a frame, not as evidence.")
notes(s, "Takeaway: the fix for two opposing executives is a shared yardstick, not a compromise number. "
      "Hardest question: 'What if the dispute is really about capital for Mexico and the US?' Answer: then slide 12 is the decision and "
      "this slide only sets the opex envelope; I check which it is in week one (A4).")

# ============================================================ 11. (new) Capital lens, cloned from the 3-card slide (old 11)
cap_s = clone_slide(S[10])
set_text(sh(cap_s, "TextBox 4"), "11 // DELIVERABLE 2: THE CAPITAL QUESTION THE BOARD WILL ACTUALLY ASK")
set_text(sh(cap_s, "TextBox 5"), "Equity by entity is the Board's larger decision than the opex cap; it gets its own ranking, yardstick and owner")
set_text(sh(cap_s, "TextBox 7"), "1. WHAT IS ON THE TABLE [1]")
set_text(sh(cap_s, "TextBox 8"), "Four uses of equity")
set_text(sh(cap_s, "TextBox 9"), [[("Brazil deepening; Nu Mexico, authorized as a bank in July 2026 and operating as one since August ($4.2B planned to 2030); the US de novo charter, to be capitalized within 12 months of the 29 January 2026 OCC approval; and the buyback of up to $1.0B announced in June 2026.", False)]])
set_text(sh(cap_s, "TextBox 10"), "2. THE YARDSTICK")
set_text(sh(cap_s, "TextBox 11"), "Return on allocated capital")
set_text(sh(cap_s, "TextBox 12"), [[("Rank each use by after-tax return on incremental allocated equity against that market's cost of equity, with capital headroom by entity under a downside case. A use that beats its cost of equity creates value even if it dilutes the 33% group ROE; the buyback is the bar every growth use must clear.", False)]])
set_text(sh(cap_s, "TextBox 13"), "3. WHO DECIDES, AND WHEN")
set_text(sh(cap_s, "TextBox 14"), "Before Investor Day")
set_text(sh(cap_s, "TextBox 15"), [[("The CFO recommends the ranking, the CEO decides, the Board approves at this meeting, so the Investor Day narrative (8 Dec 2026) rests on an approved plan. My job: one page per use with the same four lines (return, capital, headroom, risk), and the chair's pre-read at T-9.", False)]])
set_text(sh(cap_s, "TextBox 6"), "Sources: [1] Nu Q2'26 release (ROE 33%); Nu announcements: Mexico banking authorization (10 Jul 2026, operations as a bank from 6 Aug) and $4.2B plan to 2030; OCC conditional approval of Nubank N.A. (29 Jan 2026); buyback of up to $1.0B (4 Jun 2026); Investor Day 8 Dec 2026. Cost-of-equity figures are to be set with Treasury (ASSUMPTION).")
notes(cap_s, "Takeaway: show the panel you know the Board topic is capital, not a $25M opex line. "
      "Hardest question: 'Where should the next dollar of equity go?' Answer: wherever return on allocated capital beats the local cost of "
      "equity by the widest margin after the downside; today that case is strongest in Brazil deepening and Mexico, with the US sized to its "
      "100 bps envelope until its credit models mature. "
      "Recent context to have ready: on 30 Sep 2026 Nu said it is not pursuing a transaction with Monzo and that its capital allocation "
      "framework is unchanged; the buyback runs 4 Jun 2026 to 3 Jun 2027; Q3'26 results are due 12 Nov 2026, before Investor Day. "
      "Board question: 'You want the ranking approved but show no return or cost-of-equity figures.' Answer: the Board approves the "
      "yardstick and the Mexico and US capitalization sequence at this meeting; Treasury's numbers are in the pack by T-16 and the ranking "
      "page is the one I present.")

# ============================================================ 12. Board rhythm (old 9)
s = S[8]
set_text([x for x in s.shapes if x.has_text_frame and x.text_frame.text.startswith("08 //")][0], "12 // DELIVERABLE 2: THE 28-DAY BOARD RHYTHM")
title = [x for x in s.shapes if x.has_text_frame and x.text_frame.text.startswith("Our Governance Team")][0]
set_text(title, "The Board book ships at T-7 with written Q&A, after a T-16 executive lock and a T-10 input freeze")
for x in s.shapes:
    if not x.has_text_frame:
        continue
    tx = x.text_frame.text
    if tx.startswith("Issue a locked slide outline"):
        set_text(x, "I issue the locked outline and standard data tables today, so units stop building rival decks. Owner: me, with the Corporate Secretary.")
    elif tx.startswith("Run 1-on-1 pre-alignment"):
        set_text(x, "One-to-one pre-alignment at T-21; the capital options memo settles the 70/30 lock with CEO and CFO at T-16.")
    elif tx.startswith("Freeze inputs at Day 10"):
        set_text(x, "Inputs freeze at T-10 for single-pen editing; committee chairs pre-briefed at T-9; book on the portal at T-7.")
    elif tx.startswith("Answer routine data questions"):
        set_text(x, "Directors' data questions answered in writing by T-3, so the live session is spent on the capital plan, not read-outs.")
    elif tx.startswith("80% Strategy"):
        set_text(x, "Strategy, not read-outs")
    elif tx.startswith("Statutory Entity Rule"):
        set_text(x, [[("Why this meeting matters more than most: ", True), ("it is the last Board before Investor Day (8 Dec 2026). The capital plan approved here is the one the market hears. Statutory governance for the expanding entity set is on slide 14.", False)]])
    elif tx.startswith("Data & Implications"):
        set_text(x, "Sources: Nu IR calendar (Investor Day 8 Dec 2026). The T-7 book and written-Q&A practice follows common director guidance (e.g. NACD); the time-split figures often quoted for it are directional and not used here.")
notes(s, "Takeaway: a Board rhythm is a calendar with owners, not a sprint metaphor. "
      "Dates around it: Q3'26 results 12 Nov 2026; Investor Day 8 Dec 2026 (announced 2 Sep, first ever, New York). "
      "Hardest question: 'Materials are already late. What do you cut?' Answer: the outline is locked today; anything not in it goes to "
      "the appendix or the next meeting; the capital plan and statutory items are the two things that cannot slip.")

# ============================================================ 13. Statutory governance (old 17, Appendix A) promoted and corrected
s = S[16]
set_text([x for x in s.shapes if x.has_text_frame and x.text_frame.text.startswith("APPENDIX A")][0], "13 // DELIVERABLE 2: STATUTORY GOVERNANCE ACROSS NEW ENTITIES")
title = [x for x in s.shapes if x.has_text_frame and x.text_frame.text.startswith("Clear Rules")][0]
set_text(title, "The model: open an entity only when a license or tax substance requires it; officers and signing follow the stage")
tbl = [x for x in s.shapes if x.has_table][0].table
set_text(cell(tbl, 0, 0), "Stage")
set_text(cell(tbl, 0, 1), "Trigger for a local entity")
set_text(cell(tbl, 0, 2), "Operating setup")
set_text(cell(tbl, 0, 3), "Officers, signing and insurance")
set_text(cell(tbl, 1, 0), [[("1. Market test", True), (" (new countries)", False)]])
set_text(cell(tbl, 1, 1), "No local license needed, no taxable presence created; served from an existing licensed entity or a partner bank.", size=11)
set_text(cell(tbl, 1, 2), "No local entity. Cross-border or partner model.", size=11)
set_text(cell(tbl, 1, 3), [[("Central signing only. No local director.", True), (" Prevents an accidental permanent establishment.", False)]], size=11)
set_text(cell(tbl, 2, 0), [[("2. Licensed scale-up", True), (" (US charter, new verticals)", False)]], size=11)
set_text(cell(tbl, 2, 1), "A regulator requires a local licensed entity (deposits, lending, custody), or substance rules require local payroll and assets.", size=11)
set_text(cell(tbl, 2, 2), "Dedicated licensed entity with its own capital plan and headroom.", size=11)
set_text(cell(tbl, 2, 3), [[("Fit-and-proper directors;", True), (" two central officers plus one local leader; dual signature above a set limit; D&O bound before activation.", False)]], size=11)
set_text(cell(tbl, 3, 0), [[("3. Core hub", True), (" (Mexico as a bank since Aug 2026; Colombia; the HQ)", False)]], size=11)
set_text(cell(tbl, 3, 1), "Full banking or fintech license in force; local CFO in place (already so in Mexico and Colombia).", size=11)
set_text(cell(tbl, 3, 2), "Full licensed entity; regional campus where the HQ decision lands.", size=11)
set_text(cell(tbl, 3, 3), [[("Audit and Risk Committee oversight;", True), (" dual signature above limits set by category; director liability insurance; annual review of powers of attorney.", False)]], size=11)
set_text(cell(tbl, 1, 0), [[("1. Market test", True), (" (new countries)", False)]], size=11)
for c, w in enumerate([1.9, 3.7, 2.5, 3.9]):
    tbl.columns[c].width = Inches(w)
for r in range(len(tbl.rows)):
    tbl.rows[r].height = Inches(0.55 if r == 0 else 1.25)
tshape = [x for x in s.shapes if x.has_table][0]
tshape.top = Inches(1.75)
tshape.height = Inches(0.55 + 1.25 * 3)
for x in list(s.shapes):
    if x.shape_type == 1 and x.name.startswith("Rectangle"):
        x._element.getparent().remove(x._element)
foot = [x for x in s.shapes if x.has_text_frame and x.text_frame.text.startswith("Data & Implications")][0]
set_text(foot, "Sources: Nu announcements (US launch via Lead Bank, 10 Sep 2026; OCC conditional approval, 29 Jan 2026; Mexico bank authorization, 10 Jul 2026; local finance leadership in Mexico and Colombia per the CFO appointment release). Signing limits are proposals (ASSUMPTION).")
foot.top = Inches(6.75)
notes(s, "Takeaway: for a bank the entity question is a licensing question first. "
      "Hardest question: 'Would you be a statutory director of the HQ entity?' Answer: no. Statutory officers must pass fit-and-proper "
      "tests and carry personal liability; the Chief of Staff runs the process, the licensed executives hold the seats.")

# ============================================================ 14. GM pushback playbook (old 13) moved to D2
s = S[12]
set_text(sh(s, "TextBox 6"), "10 // DELIVERABLE 2: WHEN A SENIOR GM PUSHES BACK")
set_text(sh(s, "TextBox 7"), "When a senior GM pushes back: validate the goal privately, move to criteria, offer a conditional yes, co-present the result")
set_text(sh(s, "TextBox 10"), "STEP 1 // DE-ESCALATE")
set_text(sh(s, "TextBox 11"), "Separate the person from the numbers")
set_text(sh(s, "TextBox 12"), "I meet the GM privately first and confirm the strategic goal before any budget line is discussed. Nobody is corrected in front of the CFO.")
set_text(sh(s, "TextBox 13"), "STEP 2 // REFRAME")
set_text(sh(s, "TextBox 14"), "From positions to criteria")
set_text(sh(s, "TextBox 15"), "Move from 'I need $25M' to 'which indicators would prove this works?' The published rule does the arguing, not me.")
set_text(sh(s, "TextBox 16"), "STEP 3 // CO-DESIGN")
set_text(sh(s, "TextBox 17"), "Offer a conditional yes")
set_text(sh(s, "TextBox 18"), "Replace a flat no with a staged yes: the $4.5M Q1 start today and a reserved claim on gate 2 once the indicators hold.")
set_text(sh(s, "TextBox 19"), "STEP 4 // INSTITUTIONALIZE")
set_text(sh(s, "TextBox 20"), "Co-present at the lock")
set_text(sh(s, "TextBox 21"), "The GM and I present the staged plan together at the T-16 lock. There is no back channel to the CEO because there is nothing left to appeal.")
set_text(sh(s, "TextBox 9"), [[("Why this works without line authority: ", True), ("fair process (early engagement, a clear explanation, published expectations) is what makes senior people accept an outcome they argued against. The rule is public on day one; the conversation is private; the sign-off is joint.", False)]])
set_text(sh(s, "TextBox 8"), "Source: Kim and Mauborgne, 'Fair Process: Managing in the Knowledge Economy', Harvard Business Review. Everything else on this slide is practice, not research.")
notes(s, "Takeaway: four moves, in order, every time. "
      "Hardest question: 'Give me a real example.' Answer: [Pit wall: insert one case from your career here: the leader, what they asked for, "
      "what you offered instead, and the result.]")

# ============================================================ 15. HQ gates (old 10)
s = S[9]
set_text([x for x in s.shapes if x.has_text_frame and x.text_frame.text.startswith("09 //")][0], "14 // DELIVERABLE 3: THE HQ CHARTER")
title = [x for x in s.shapes if x.has_text_frame and x.text_frame.text.startswith("Govern the New")][0]
set_text(title, "HQ charter: four gates in 120 days from signature, each with an owner, an exit test and a stop condition")
for x in s.shapes:
    if not x.has_text_frame:
        continue
    tx = x.text_frame.text
    if tx.startswith("1. PROTECT"):
        set_text(x, "1. TAX SUBSTANCE (GATE 1, D+30) [1]")
    elif tx.startswith("Real Payroll"):
        set_text(x, "Incentives tied to real payroll", size=14)
    elif tx.startswith("The Data: Under 2026 OECD"):
        set_text(x, [[("The rule: ", True), ("Nu is above the EUR 750M Pillar Two threshold. In 2026 only 9.4% of payroll and 7.4% of tangible assets are carved out of the 15% top-up base, so incentives without people and assets on the ground are worth less than they look.", False)], [("My action: ", True), ("campus build-out grants and payroll credits. Tax owns the gate; exit test: modeled effective rate holds under Pillar Two.", False)]], size=12)
    elif tx.startswith("2. PREVENT"):
        set_text(x, "2. GRANT TERMS (GATE 2, D+60) [2]")
    elif tx.startswith("Proportional Hiring"):
        set_text(x, "Proportional, clawback-proof", size=14)
    elif tx.startswith("The Data: Site Selection"):
        set_text(x, [[("The risk: ", True), ("municipal grants carry headcount clawbacks that bite when hiring pace or hybrid work changes.", False)], [("My action: ", True), ("proportional scaling (85% of headcount keeps 85% of the grant) and a 24-month grace period. Legal owns the gate; exit test: no cliff-edge clawback in the signed term sheet.", False)]], size=12)
    elif tx.startswith("3. SHIELD"):
        set_text(x, "3. OFFICERS (GATE 3, D+90) [3]")
    elif tx.startswith("Dual Control"):
        set_text(x, "Fit-and-proper officers, D&O", size=14)
    elif tx.startswith("The Data: Civil-"):
        set_text(x, [[("The risk: ", True), ("statutory directors carry personal liability for tax and labor filings; early commercial signing can create a taxable presence before the entity is ready.", False)], [("My action: ", True), ("officers who pass the regulator's test, dual signature above category limits, D&O bound before activation. Legal and People own the gate. Gate 4 (D+120): the lease.", False)]], size=12)
    elif tx.startswith("4-Stage Execution Gate"):
        set_text(x, [[("Leading it as an interim across non-finance teams: ", True), ("one weekly pod (Legal, Tax, People, Workplace), one decision log, 48-hour escalation to the CFO, and a five-item risk register reviewed at every gate: tax residence, clawback, officer liability, site politics, key-person dependence on me. Jurisdiction is open (A8).", False)]], size=12.5)
    elif tx.startswith("Data & Implications"):
        set_text(x, "Sources: [1] OECD GloBE Model Rules, Art. 5.3 and the Art. 9.2 transitional rates for 2026. [2] Location-incentive practice (e.g. Deloitte, Site Selection) is directional. [3] Cross-border director-liability guidance (e.g. NACD, Baker McKenzie) is directional. No savings figure is claimed until a site is chosen.")
notes(s, "Takeaway: the HQ is a governance project before it is a real-estate project. "
      "Hardest question: 'What is the dollar value of the incentives?' Answer: none is claimed yet; it depends on the site, and any figure "
      "before gate 1 would be a guess. The exit test at gate 1 produces the number. "
      "Tax question: 'Nu Holdings is Cayman-incorporated; does a corporate HQ move the place of effective management, and which "
      "jurisdiction collects the Pillar Two top-up?' Answer: gate 1's exit test includes a tax-residence opinion; the carve-out only bites "
      "where a qualified domestic minimum top-up tax applies (Brazil has one since 2025), so it is a site-selection input, not a saving.")

# ============================================================ 16. D4 one slide (old 11), first person
s = S[10]
set_text(sh(s, "TextBox 4"), "15 // DELIVERABLE 4: HOW I LEAD")
set_text(sh(s, "TextBox 5"), "How I lead without line authority: three rules I hold myself to, and what I would not touch in the first 90 days")
set_text(sh(s, "TextBox 7"), "1. FAIR PROCESS FIRST")
set_text(sh(s, "TextBox 8"), "Publish the rule, then negotiate")
set_text(sh(s, "TextBox 9"), [[("The rule: ", True), ("people accept outcomes they argued against when the criteria were public before the decision. I publish the scoring rubric on day one and never change it mid-cycle.", False)], [("In practice: ", True), ("every GM sees the same four indicators and the same release rule.", False)]])
set_text(sh(s, "TextBox 10"), "2. NO PUBLIC AMBUSHES")
set_text(sh(s, "TextBox 11"), "Fix the numbers privately")
set_text(sh(s, "TextBox 12"), [[("The rule: ", True), ("model gaps are closed one-to-one before anything reaches the CFO or the Board. Credibility with senior people comes from never surprising them in a room.", False)], [("In practice: ", True), ("the bi-weekly GM one-to-ones and the T-16 lock exist for this.", False)]])
set_text(sh(s, "TextBox 13"), "3. HANDS-ON RIGOR")
set_text(sh(s, "TextBox 14"), "Own the model and the clauses")
set_text(sh(s, "TextBox 15"), [[("The rule: ", True), ("strategic influence is only trusted when the hands-on FP&A and governance rigor is mine. I build and check the model and read the clauses myself; I do not delegate what can embarrass the CFO.", False)], [("What I would not touch: ", True), ("local-CFO decision rights in Mexico and Colombia, credit-model governance, the IR calendar.", False)]])
set_text(sh(s, "TextBox 6"), "Sources: the fair-process principle follows Kim and Mauborgne (HBR). The rest is how I work; examples from my own record are in the speaker notes and are the first thing I would share in Q&A.")
notes(s, "Takeaway: three rules, each tied to a mechanism on an earlier slide. "
      "Examples to give in Q&A [Pit wall: insert three from your career, one per rule: the situation, your role, the decision, the result].")

# ============================================================ 17. Scenarios and reversal triggers (old 14)
s = S[13]
set_text(sh(s, "TextBox 26"), "16 // RISKS, SCENARIOS AND REVERSAL TRIGGERS")
set_text(sh(s, "TextBox 27"), "If Q1 indicators miss, gate 2 is held; the A6 valve opens only on an H1 profit beat, inside the efficiency guide")
tbl = [x for x in s.shapes if x.has_table][0].table
set_text(cell(tbl, 0, 0), "Dimension")
set_text(cell(tbl, 0, 1), "Downside (market slows)")
set_text(cell(tbl, 0, 2), "Base plan")
set_text(cell(tbl, 0, 3), "Upside (H1 beats plan)")
SZ = 11
set_text(cell(tbl, 1, 0), [[("1. Funded spend vs the $25M cap", True)]], size=SZ)
set_text(cell(tbl, 1, 1), [[("$15.0M (gate 2 held at Q1 close)", False)]], size=SZ)
set_text(cell(tbl, 1, 2), [[("$25.0M ($15M gate 1 + $10M gate 2)", False)]], size=SZ)
set_text(cell(tbl, 1, 3), [[("$25.0M", False)]], size=SZ)
set_text(cell(tbl, 2, 0), [[("2. Release above the cap (rule A6)", True)]], size=SZ)
set_text(cell(tbl, 2, 1), [[("$0", False)]], size=SZ)
set_text(cell(tbl, 2, 2), [[("$0 (not assumed)", False)]], size=SZ)
set_text(cell(tbl, 2, 3), [[("Up to $15.0M from H1 profit above plan, inside the ~20% FY26 guide", False)]], size=SZ)
set_text(cell(tbl, 3, 0), [[("3. Reversal triggers (stop or hold)", True)]], size=SZ)
set_text(cell(tbl, 3, 1), [[("A unit below plan on two of four indicators: its gate-2 share is held", False)]], size=SZ)
set_text(cell(tbl, 3, 2), [[("Efficiency ratio above the FY26 guide at Q1 close: gate 2 held for all units", False)]], size=SZ)
set_text(cell(tbl, 3, 3), [[("Credit cost or NPL 90+ above plan: rule A6 is not used", False)]], size=SZ)
set_text(cell(tbl, 4, 0), [[("4. Top risks and owners", True)]], size=SZ)
set_text(cell(tbl, 4, 1), [[("GMs game the indicators: finance partners own the read-out", False)]], size=SZ)
set_text(cell(tbl, 4, 2), [[("Board book slips: outline locked today; anything outside it goes to the appendix", False)]], size=SZ)
set_text(cell(tbl, 4, 3), [[("HQ: tax residence, clawback, officer liability, site politics, key-person risk", False)]], size=SZ)
set_text(cell(tbl, 5, 0), [[("5. What I report at Day 90", True)]], size=SZ)
set_text(cell(tbl, 5, 1), [[("First monthly indicator read-outs per unit (gate 2 decided at Q1 close)", False)]], size=SZ)
set_text(cell(tbl, 5, 2), [[("Board delivered at T-0; capital plan approved; cycle-time and forecast baselines", False)]], size=SZ)
set_text(cell(tbl, 5, 3), [[("HQ gates 1 and 2 cleared; gate 3 pack ready", False)]], size=SZ)
for r in range(len(tbl.rows)):
    tbl.rows[r].height = Inches(0.5 if r == 0 else 0.78)
tshape = [x for x in s.shapes if x.has_table][0]
tshape.top = Inches(1.88)
tshape.height = Inches(0.5 + 0.78 * 5)
for x in list(s.shapes):
    if x.shape_type == 1 and x.name.startswith("Rectangle"):
        x._element.getparent().remove(x._element)
set_text(sh(s, "TextBox 28"), [[("For the Board: ", True), ("funding above $15M is released only on evidence, so margins are protected automatically if conditions tighten, with no mid-year freeze. No value figures are claimed beyond the spend lines; they would be guesses before HQ gate 1.", False)]], size=11.5)
sh(s, "TextBox 28").top = Inches(6.55)
notes(s, "Takeaway: the downside is handled by the rule, not by a meeting. "
      "Hardest question: 'What would make you stop?' Answer: two of four indicators below plan for a unit, the efficiency ratio above the "
      "guide for the company, or credit cost above plan for rule A6. All three are observable at quarter close. "
      "Context a short-seller would raise: NPL 90+ rose 35 bps to 6.9% in Q2'26 (management called it seasonal); the Brazil election "
      "runoff is 25 Oct 2026; sell-side is split (JPMorgan raised its target in October, Itau BBA downgraded on consumer-credit risk). "
      "Short-seller question: 'The A6 valve spends H1 beats instead of letting them reach net income; is that not the padding you claim to "
      "remove?' Answer: it is capped, formula-sized inside the ~20% guide, CFO- and IR-signed, and disclosed as upside spend, not a change "
      "to guidance; and it only exists in the year the beat exists.")

# ============================================================ 18. 90-day plan (old 15)
s = S[14]
set_text(sh(s, "TextBox 4"), "17 // THE NEXT 90 DAYS")
set_text(sh(s, "TextBox 5"), "Days 61-150: Board locked by Day 88, hub in place by Day 95, HQ gate 2 cleared by Day 120")
set_text(sh(s, "TextBox 13"), [[("Today (T-28): ", True), ("Board outline issued; HQ charter signed; gate 1 launched.", False)]])
set_text(sh(s, "TextBox 15"), [[("Day 72 (T-16): ", True), ("options memo settles the 70/30 lock with CEO, CFO, MTeam.", False)]])
set_text(sh(s, "TextBox 17"), [[("Day 88 (T-0): ", True), ("Board delivered; $15M gate-1 release memo published.", False)]])
set_text(sh(s, "TextBox 19"), [[("Day 95: ", True), ("hub stood up; finance partners embedded in each market and vertical.", False)]])
set_text(sh(s, "TextBox 21"), [[("Day 105: ", True), ("squad clinics running; $12M case track under way.", False)]])
set_text(sh(s, "TextBox 23"), [[("Day 120 (HQ D+60): ", True), ("gate 2 cleared: proportional grant terms, 24-month grace.", False)]])
set_text(sh(s, "TextBox 25"), [[("Day 133 (1 Jan): ", True), ("FY27 starts; $15M gate 1 opens to the vetted cases.", False)]])
set_text(sh(s, "TextBox 27"), [[("Day 150 (HQ D+90): ", True), ("gate 3 cleared; Day-90 report to the CFO.", False)]])
set_text(sh(s, "TextBox 29"), [[("Beyond: ", True), ("gate 2 at Q1 close (31 Mar); HQ gate 4 at D+120.", False)]])
set_text(sh(s, "TextBox 17"), [[("Day 88 (T-0): ", True), ("Board approves the FY27 plan; gate-1 allocation memo published.", False)]])
set_text(sh(s, "TextBox 7"), "Lock the Board and gate 1")
set_text(sh(s, "TextBox 9"), "Stand up the hub; HQ gate 2")
set_text(sh(s, "TextBox 11"), "FY27 opens; HQ gate 3")
set_text(sh(s, "TextBox 5"), "Days 61-150 (today to mid-January): Board approved by Day 88, hub in place by Day 95, FY27 gate 1 open on 1 January")
notes(s, "Takeaway: one clock. Tenure days on this slide (today is Day 60, 20 Oct 2026); HQ gates are D+n from the charter signed today; "
      "funding gates follow the FY27 fiscal calendar (A10). "
      "Hardest question: 'Why does the first release wait until January?' Answer: because the cap is FY27's envelope; the Board approves the "
      "plan on Day 88, the cases are vetted in December, and nobody is asked to prove a Q1 indicator in 37 days.")

# ============================================================ 19. Closing = the ask (old 20)
s = S[19]
for x in s.shapes:
    if not x.has_text_frame:
        continue
    tx = x.text_frame.text
    if tx.startswith("Ready to Launch"):
        set_text(x, "The ask: approve the $25M FY27 cap in two gates with the A6 valve stated, the T-16 / T-7 Board calendar, and the HQ charter. I report back at Day 90.", size=34)
    elif tx.startswith("Thank you"):
        set_text(x, "KPIs at Day 90: 100% of gate-1 spend with an approved case; Board book at T-7; HQ gates 1 and 2 cleared. Stop rules: two of four indicators below plan, or the efficiency ratio above the FY26 guide.")
        x.text_frame.word_wrap = True
        x.left = Inches(1.5)
        x.width = Inches(10.33)
        x.height = Inches(0.9)
notes(s, "Close on the decision, not on thanks. Then open Q&A with the assumptions slide on screen.")

# ============================================================ Appendix: divider (old 16), ledger (old 18); AI slide (old 19) removed
s = S[15]
for x in s.shapes:
    if x.has_text_frame and x.text_frame.text.startswith("Appendix A"):
        set_text(x, ["Appendix A: Sources and assumptions ledger   |   Appendix B: Option scores and weight sensitivity"])
notes(s, "Divider. Not presented; the two appendix pages are backup for Q&A.")
s = S[17]
set_text([x for x in s.shapes if x.has_text_frame and x.text_frame.text.startswith("APPENDIX B")][0], "APPENDIX A // SOURCES AND ASSUMPTIONS LEDGER")
title = [x for x in s.shapes if x.has_text_frame and x.text_frame.text.startswith("Complete Ledger")][0]
set_text(title, "Every number in the deck is a sourced fact, a labeled assumption, or a directional benchmark that carries no weight")
tbl = [x for x in s.shapes if x.has_table][0].table
set_text(cell(tbl, 0, 0), "Topic")
set_text(cell(tbl, 0, 1), "Source and class")
set_text(cell(tbl, 0, 2), "Data point")
set_text(cell(tbl, 0, 3), "Where it is used")
set_text(cell(tbl, 1, 0), [[("1. Nu scale and efficiency", True)]])
set_text(cell(tbl, 1, 1), [[("Nu Q2'26 release and 6-K; 20-F", True), (" | FACT", False)]])
LZ = 10.5
set_text(cell(tbl, 1, 0), [[("1. Nu scale and efficiency", True)]], size=LZ)
set_text(cell(tbl, 1, 1), [[("Nu Q2'26 release and 6-K; 20-F", True), (" | FACT", False)]], size=LZ)
set_text(cell(tbl, 1, 2), [[("138.9M customers; ARPAC $17.1; cost to serve $1.0; efficiency ratio 19.5% (FY26 ~20% per the call); ROE 33%; net income $1.06B; low-and-grow underwriting.", False)]], size=LZ)
set_text(cell(tbl, 1, 3), [[("Slides 4, 5, 8, 10, 17: the cap, the planning principle, the 70/30 release.", False)]], size=LZ)
set_text(cell(tbl, 2, 0), [[("2. Corporate events", True)]], size=LZ)
set_text(cell(tbl, 2, 1), [[("Nu announcements; OCC; Nu IR", True), (" | FACT", False)]], size=LZ)
set_text(cell(tbl, 2, 2), [[("Mexico bank authorization 10 Jul 2026 (bank operations from 6 Aug), $4.2B plan to 2030; OCC conditional approval 29 Jan 2026 (capitalize within 12 months); US launch via Lead Bank 10 Sep 2026; buyback up to $1.0B (4 Jun 2026); CFO since 13 Jul 2026; Q3'26 results 12 Nov 2026; Investor Day 8 Dec 2026.", False)]], size=LZ)
set_text(cell(tbl, 2, 3), [[("Slides 12, 13, 14: the capital question and the Board calendar.", False)]], size=LZ)
set_text(cell(tbl, 3, 0), [[("3. HQ tax rules", True)]], size=LZ)
set_text(cell(tbl, 3, 1), [[("OECD GloBE Model Rules", True), (" | FACT (rule); ASSUMPTION (site)", False)]], size=LZ)
set_text(cell(tbl, 3, 2), [[("EUR 750M threshold; 15% minimum; 2026 transitional carve-out 9.4% of payroll and 7.4% of tangible assets (Art. 9.2).", False)]], size=LZ)
set_text(cell(tbl, 3, 3), [[("Slide 15, gate 1. No savings figure is claimed before a site is chosen.", False)]], size=LZ)
set_text(cell(tbl, 4, 0), [[("4. Directional benchmarks", True)]], size=LZ)
set_text(cell(tbl, 4, 1), [[("Hackett, McKinsey, Gartner, HBR, NACD, Deloitte", True), (" | CONTEXT, not load-bearing", False)]], size=LZ)
set_text(cell(tbl, 4, 2), [[("Central hubs shorten planning cycles; active reallocators earn higher returns; fair process wins acceptance; T-7 Board books shift meeting time to strategy.", False)]], size=LZ)
set_text(cell(tbl, 4, 3), [[("Named in footers only. No headline rests on them.", False)]], size=LZ)
for c, w in enumerate([1.8, 2.4, 4.6, 3.2]):
    tbl.columns[c].width = Inches(w)
for r in range(len(tbl.rows)):
    tbl.rows[r].height = Inches(0.5 if r == 0 else 1.05)
tshape = [x for x in s.shapes if x.has_table][0]
tshape.top = Inches(1.65)
tshape.height = Inches(0.5 + 1.05 * 4)
for x in list(s.shapes):
    if x.shape_type == 1 and x.name.startswith("Rectangle"):
        x._element.getparent().remove(x._element)
foot = [x for x in s.shapes if x.has_text_frame and x.text_frame.text.startswith("Independent Audit")][0]
set_text(foot, "My assumptions (A3-A6, A9, A10) are on slide 3 with their basis. Research assisted by AI tools; every figure, calculation and recommendation was checked by me.")
foot.top = Inches(6.65)
notes(s, "Backup. If asked about any number, this is the page.")

# Appendix B: option scores (clone of the options slide) kept as backup with the sensitivity table note
appb = clone_slide(S[4])
set_text(sh(appb, "TextBox 31"), "APPENDIX B // OPTION SCORES AND WEIGHT SENSITIVITY")
set_text(sh(appb, "TextBox 32"), "Option 3 stays first under every +/-10 point change to any single weight")
set_text(sh(appb, "TextBox 34"), [[("Sensitivity (my calculation): ", True), ("moving any one weight by +/-10 points and rebalancing the others pro rata leaves the order 3 > 2 > 1 unchanged; the smallest gap between option 3 and option 2 across all cases is above 20 points. Scores are judgment (A9).", False)]])
set_text(sh(appb, "TextBox 33"), "Weighted total = sum of score x weight. Base weights: cap 30%, growth return 25%, GM ownership 20%, agility 15%, speed 10%.")
notes(appb, "Backup for the options slide.")

# remove the AI-disclosure slide (old 19)
delete_slide(S[18])

# ============================================================ Reorder
order = [S[0], S[1], assump, S[2], S[3], S[4], S[5], S[6], S[11], S[7], S[12], cap_s, S[8], S[16], S[9], S[10], S[13], S[14], S[19], S[15], S[17], appb]
for idx, sl in enumerate(order):
    move_slide(sl, idx)

# ============================================================ Titles and headings always wrap (the source boxes had wrap off);
# slide titles stay inside the 12.0in content grid instead of running to the slide edge
for sl in prs.slides:
    for x in sl.shapes:
        if not x.has_text_frame:
            continue
        for p in x.text_frame.paragraphs:
            for r in p.runs:
                if r.font.size is not None and r.font.size.pt >= 14:
                    x.text_frame.word_wrap = True
                if r.font.size is not None and r.font.size.pt >= 25 and x.left < Inches(1.0) and x.width > Inches(12.3):
                    x.width = Inches(12.0)
                break
            break

# options slide and its appendix clone: drop the stray separator rectangles under the last row
for sl in (S[4], appb):
    for x in list(sl.shapes):
        if x.shape_type == 1 and x.name.startswith("Rectangle"):
            x._element.getparent().remove(x._element)

# planning-principle slide (old 4): align card 1 with cards 2 and 3
sl = S[3]
cards = sorted([x for x in sl.shapes if x.shape_type == 1 and x.name.startswith("Rounded") and x.width < Inches(5)], key=lambda c: c.left)
if len(cards) == 3:
    ref = cards[1]
    c1 = cards[0]
    dy = ref.top - c1.top
    c1.top, c1.height = ref.top, ref.height
    for x in sl.shapes:
        if x.has_text_frame and x.left < c1.left + c1.width and x.left >= c1.left - Inches(0.1) and c1.top - Inches(0.5) < x.top + Inches(0.5) < c1.top + c1.height and x.top > Inches(2.0):
            x.top = x.top + dy

# ============================================================ Slide numbers (bottom right) on every slide
for i, sl in enumerate(prs.slides, 1):
    dark = i in (1, 19)
    add_textbox(sl, 12.4, 7.05, 0.8, 0.3, str(i), size=9, color=(RGBColor(0xC0, 0x84, 0xFC) if dark else MUTED), align=PP_ALIGN.RIGHT)

prs.save(DST)
print("saved", DST, "slides:", len(prs.slides))
