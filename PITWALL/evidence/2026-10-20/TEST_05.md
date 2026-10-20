# TEST_05: deck v5 (three-agent review applied) and the v5 talk track and Q&A (2026-10-20)

Input: the assessed action plan in `review/2026-10-20_three-agent-review.md`, items 0-11b. The pit wall's instruction:
"incorporate the proposed changes. ensure no key data point, statement or assumption reference gets lost."

Files under test:
- `deck/CFO_CoS_deck_v5.pptx`, built by `tools/build_v5.py` from the v4 template, with packaging by `tools/pptx_safe.py`.
  sha256 9afe3bafea576edc38cfd4f4339d80c8aa3082e8d1cfe85d2000c6b5775b4e81.
- `review/2026-10-20_v5_talk-track-and-QA.md`.

## Action plan items: applied or not

| # | Item | Status in v5 |
|---|---|---|
| 0 | Opex ~$3.2B, not ~$4.6B; $35M about 1.1% of opex, about 0.2 points of efficiency ratio | Applied: slide 4, A3, ledger, notes, numbers sheet. Q1 also has $25M at about 0.15 points |
| 1 | 14 presented slides; six slides moved to backup; duplicate options table cut | Applied: slides 1-14, divider, backups A-G (G is new) |
| 2 | D2 anchored on capital; 70/30 as the second half | Applied: slide 8 |
| 3 | "H2 reserve", "first/second release", spoken "days out"; A-numbers only as references | Applied |
| 4 | Reserve, release amounts and fair process each defined once | Applied: reserve on slide 6, releases on slide 6, fair process on slide 9 (slide 12 cites it) |
| 5 | The CFO decides, Risk can block, IR sets disclosure; A2 labeled FACT (management guidance); R3 checkpoints and 15-minute cut | Applied |
| 6 | D4 retitled to the brief; derailer, hiring and triage lines | Applied: slide 12 title and label use the brief's wording; hiring line on slide 7; triage on slide 1; derailer is a slot only the pit wall can fill |
| 7 | Short assertion titles, sentence case, no card scaffolding labels, "check" not "cheque" | Applied |
| 8 | 0.15 points; Q2 real-estate timing on the HQ slide; "smart efficiency" and "owners" | Applied: A3, slide 11 gate 4, slide 4 |
| 9 | "Confirming now" for "week one"; "nothing left to appeal" replaced | Applied: slides 3 and 9 |
| 10 | Q&A: Brazil CFO, pushback on the CFO, stakeholder map | Applied: Q17, Q18, Q20, backup G |
| 11 | Speaker notes: one plain sentence per slide, no coaching, no placeholders, no "Pit wall" | Applied. A first pass still had coaching ("say so", "Do not oversell"); it was found and rewritten in this run |
| 11b | Spoken sentences of 30 words or fewer; Q&A answers under 60 seconds | Applied: no spoken sentence over 30 words (scripted check). Q&A Say lines run 46-112 words, about 21-52 seconds |
| 12 | Personal examples and a timed dry run | NOT DONE: needs the pit wall (B16, B8) |

## Gates on v5

- **OOXML validation:** "All validations PASSED!"
- **Strict package check:** no structural problems. It covers content types, relationship targets, r:id resolution, unique
  shape ids, no control characters and no embedded fonts. sldIds 256-277 are in presentation order.
- **Render:** LibreOffice converted all 22 slides, and every slide was inspected.
  - Defect found: slide 21's two-line title overlapped its table.
  - Fixed: shorter title.
  - Guard added: `build_v5.py` now warns when a title's predicted last line reaches the next shape. Shown working by
    rebuilding with the old title, which gave "slide 21: shape 'Table 3' at 1.50in collides with title ending 1.53in".
  - Text fit check: no overflow predicted.
- **Capability scan:** 44 pass, 1 fail (the personal example, B5/B16), the same as v4. See `scan_v5.txt`. The first v5
  scan showed 37 pass, 8 fail. The 7 failures that were not the personal example were handled as follows:
  - **Fixed in the deck (real losses):**
    - The D4 label now uses the brief's "Leadership style and operating philosophy".
    - "Test-and-learn" is back on slide 4.
    - The hub's KPIs are named as KPIs on slide 7.
  - **Widened in REGRESSION.md (equivalent wording; the change log is in the file):**
    - "equal cut" (option A).
    - "staged release" / "two releases" (the allocation rule).
    - "stop rules" (R2's replacement for "reversal triggers").
  - **Narrowed to its purpose:** the stale cost-to-serve guard no longer fires on "$0.8 a year earlier". Proven by test
    strings: "cost to serve is $0.80 per month", "$0.8 monthly cost to serve" and "cost to serve $0.8." are still caught;
    "cost to serve $1.0 ($0.8 a year earlier)" is not.
- **Dates:** every tenure day was recomputed from Day 60 = Tue 20 Oct 2026.
  - Weekday checks: 17 Nov (Tue), 19 Nov (Thu), 1 Jan 2027 (Fri), 31 Mar 2027 (Wed), 12 Nov (Thu), 8 Dec (Tue).
  - Dates that fall on a weekend (Day 72 Sun, Day 78 Sat, Day 120 Sat) are shown as "Early Nov" or "Mid-Dec", not as
    exact dates.
  - Backup F gained the missing Day 90 row (Thu 19 Nov, HQ D+30, gate 1).
  - "Day 90 report" was renamed "90-day report" or "in 90 days". Tenure Day 90 is 19 Nov, but the report is due 90 days
    from today (mid-January).

## Preservation check (v4 to v5)

Script: `tools/preservation_check.py` (usage: `python3 -I tools/preservation_check.py OLD.pptx NEW.pptx [extra text files...]`). It extracts tokens from all v4 slide text, tables, chart data and
speaker notes:
- dollar amounts, percentages and bps;
- T-n, D+n and Day n;
- A1-A10;
- dates, quarters, and figures with one decimal place.

Each token is looked up in v5's slides and notes, plus the v5 talk track and Q&A, ignoring spaces.

- **Result: 142 v4 tokens; 1 not found verbatim, "29 January 2026".** v5 writes it "29 Jan 2026", so it is a format
  difference only.
- **Intentional changes** (listed in section 6 of the talk-track file):
  - $4.6B became $3.2B: correction.
  - "Day-90" became "90-day": disambiguation.
  - "20 points" became 24.2: computed exactly.
- **v4 notes content not on any v5 slide** is preserved in the Q&A:
  - Monzo (30 Sep 2026); NPL 90+ 6.9%, up 35 bps; election runoff 25 Oct 2026; JPMorgan and Itaú BBA.
  - Brazil QDMTT from 2025; Cayman / place of effective management.
  - Special meeting and written consent; FDIC and Fed pending.
  - Investor Day announced 2 Sep; buyback to 3 Jun 2027; the 30-day repayment question; the 37 days; $5.9B a quarter.
- **The v4 talk track's 133 tokens were checked against v5 plus the new file.** The only items not found verbatim are
  the old timing table's cumulative minutes (replaced) and date spellings ("10 September 2026", "31 March", "8 December",
  "29 January 2026"). All of those appear in v5 in short form.

## Talk track

- **Length** (`tools/talk_timing.py`): 3,134 spoken words, about 24.1 minutes at 130 wpm. Story slots are counted at their stated length.
- **Clock checkpoints:** slide 8 starts at 10.9 minutes (target 13), slide 11 at 16.8 (target 19.5) and slide 13 at 21.6
  (target 23).
- **Q&A:** 36 questions, each with why they ask, position, spoken answer, tone and the slide to show.
- **Pronouns:** none gendered.
- **Unverified items are flagged in the text:**
  - The call-summary figures in Q35 (12.4% risk-adjusted NIM, "same discipline", the Desenrola split) are held as "verify
    before quoting" and kept out of the spoken answer.
  - Campus program building details are marked as context only.

## Still NOT proven

- That PowerPoint opens v5. There is no PowerPoint here; v5 uses the same packaging as the v3 fix, and you have not yet
  confirmed that one either.
- The personal examples and derailer (B16). The timed dry run (B8). Q3'26 figures (due 12 Nov 2026).
