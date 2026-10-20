# TEST_02 - gate run on the CFO Chief of Staff deck v2 (2026-10-20, built 10:05-11:05 EDT)

File under test: `deck/CFO_CoS_deck_v2.pptx` sha256 9e79926f593bacb700d47d62bc211213a7851e3009efd38637b7086c84b64461
Built by `tools/build_v2.py` from the repaired v1 copy (0012dc44...). Reproducible: run the script again to rebuild.
Criteria: `PITWALL/briefs/deck-v1-gates.md`, unchanged from TEST_01. Coordinator: the reviewing session. Tester's step (dry run): NOT RUN.

## What changed between v1 and v2 (priority order, as built)

1. Facts and cap logic (core argument): cost to serve $1.0 and 138.9M customers (Q2'26); cohort ROE replaced by group ROE 33%; the FY26
   ~20% efficiency guide frames the cap; Stage 3 became rule A6 (above the cap only from H1 profit above plan, sized to stay inside the
   guide, CFO and IR sign-off, not in the base plan); "margin neutral" removed; all unsupported value figures ($15.6M, $36.4M, $61.2M,
   $6.2M, $21.4M) removed; the OECD carve-out described correctly; one clock (tenure days, HQ gates as D+n from the charter).
2. Structure: new slide 3 (assumptions A1-A9 and the week-one list); new slide 11 (capital allocation across entities: what is on the
   table, the yardstick, who decides before Investor Day); statutory governance promoted from the appendix to slide 13 with license-first
   logic; Deliverable 4 cut from three slides to one, first person; the cadence table moved to D1 and the pushback playbook to D2;
   closing slide is the ask with KPIs and stop rules; AI-disclosure slide and "Independent Audit Verdict" removed, one honest line kept.
3. Evidence hygiene: third-party statistics removed from every headline and card; four sources kept (Nu filings, OECD GloBE rules, Kim and
   Mauborgne fair process, directional planning benchmarks in the ledger only); every "proves" removed; "seamless", "connective tissue",
   "catalyst", "bulletproof" removed.
4. Polish: native bridge chart replaces the raster image; speaker notes on all 22 slides (takeaway plus hardest question and answer);
   slide numbers; name, date and fact-base line on the title; options slide states the weight sensitivity.

## G1 Scope (capability scan)

`python3 -I tools/deck_capability_check.py deck/CFO_CoS_deck_v2.pptx` -> `scan_v2.txt` (this folder). RESULT 44 pass, 1 fail (exit 1).

| Line group | v1 repaired | v2 |
|---|---|---|
| Section A (brief, 21 lines) | 20 PASS, 1 FAIL | 21 PASS |
| Section B (reviewer, 16 lines) | 4 PASS, 12 FAIL | 15 PASS, 1 FAIL |
| Section C (guard, 8 lines) | 0 PASS, 8 FAIL | 8 PASS |

Remaining FAIL: "a personal example backs the leadership claims". The examples are the candidate's own and cannot be written by the
reviewer; slots are marked in the speaker notes of slides 14 and 16 (board decision B5).

Regex corrections recorded (the gate brief requires it): three lines were spelled for v1's wording and were widened to accept
equivalent forms ("staged release", "pushes back", "planning hub"); none was loosened past the capability it names.

**G1 outcome: PASS** (all section-A lines pass; section B's one FAIL is a board item, not a scope gap).

## G2 Facts and arithmetic

| Slide | v2 value | Latest public figure | Status |
|---|---|---|---|
| 3, 4, 21 | cost to serve $1.0; 138.9M customers; ARPAC $17.1; efficiency ratio 19.5%, FY26 ~20%; ROE 33%; net income $1.06B | Q2'26 release and call | PASS |
| 10 | mature cohorts "well above the $17.1 group ARPAC" (no figure) | Nu cohort disclosures; exact figure not quoted | PASS (direction only; cite the page if a number is added) |
| 11, 13, 21 | Mexico bank authorization 10 Jul 2026, $4.2B to 2030; OCC conditional approval 27 Jan 2026, capitalize within 12 months; US launch via Lead Bank 10 Sep 2026; buyback up to $1.0B (Jun 2026); CFO 13 Jul 2026; Investor Day 8 Dec 2026 | Nu announcements, OCC, Nu IR (fact base 20 Oct 2026) | PASS (re-verify the day before the panel) |
| 15 | EUR 750M threshold; 9.4% payroll / 7.4% tangible assets (2026) as an exclusion from the top-up base | OECD GloBE Art. 9.2 | PASS |
| 10, 15, 17 | no value figures claimed | - | PASS |

Arithmetic re-run in code: option scores 57.75 / 63.75 / 91.75 (slide shows 57.8 / 63.8 / 91.8) TIES; bridge 60 - 8 - 12 - 15 = 25 TIES;
gate math 10.5 + 4.5 = 15, 16.5 + 8.5 = 25, 5 + 10 = 15 TIES; weight sensitivity (each weight +/-10 points, others rebalanced pro rata):
order 3 > 2 > 1 in all ten cases, smallest gap between options 3 and 2 = 24.2 points, which supports the "above 20 points" statement on
appendix B. Materiality: $35M = 14.9 bps of annualized Q2'26 revenue, 0.76% of implied opex (slide 4 says "about 15 bps, under 1%").
Guard: 8 of 8 forbid lines pass.

**G2 outcome: PASS.**

## G3 Decision logic

| Criterion | v2 | Status |
|---|---|---|
| (a) one envelope | $25M hard cap in two gates everywhere; above-cap release only under rule A6, labeled "not in the base plan" on slides 3, 8, 10, 17 | PASS |
| (b) one clock | tenure days on slide 18; HQ gates D+30/60/90/120 from the charter signed on Day 60; Day 120 = D+60, Day 150 = D+90, gate 4 at Day 180 stated | PASS |
| (c) observable gate criteria | gate 2 reads activation, 30-day retention, CAC vs plan, unit cost vs plan at quarter close; payback is the lagging confirmation | PASS |
| (d) Board topic | slide 11 answers capital allocation across entities; assumption A4 states the reading of the MTeam dispute | PASS |
| (e) options, weights, sensitivity | three options, weights disclosed, scores labeled judgment (A9), sensitivity stated and verified | PASS |

**G3 outcome: PASS.**

## G4 Panel readiness

| Criterion | v2 | Evidence |
|---|---|---|
| (a) validates and opens | PASS: "All validations PASSED!"; LibreOffice converts | validator; render PDF |
| (b) renders clean | PASS after three passes: all 22 slides inspected; last pass fixed the cadence title, the 90-day card 3 bullets and the rule A6 bullet | render inspection |
| (c) notes on every core slide | PASS: 22 of 22 | `check: notes` |
| (d) dry run 25-30 min | NOT RUN | INCONCLUSIVE = FAIL (board B8) |
| (e) red-team answers on slides | 9 of 10 answerable from slides or notes; the "give me a real example" question needs the candidate's own case (B5) | notes |
| raster chart | PASS: native bar chart on slide 7 | `check: no_raster_charts` |

**G4 outcome: FAIL** on (d) only. The file is otherwise ready for the dry run.

## Branches taken on the board without a recorded pit-wall call (to confirm or reverse)

B1 rule A6 (formula-sized reinvestment inside the guide) rather than deletion; B2 capital slide added; B3 AI slide cut, one line kept;
B4 jurisdiction left open and said so; B5 one first-person slide with example slots in the notes; B6 "I" voice; B7 benchmarks culled to
four. Each is one edit in `tools/build_v2.py` to reverse.

## Still NOT proven after this run

- That the deck fits 30 minutes (no dry run; 19 presented slides at ~1.5 min each is the estimate, not a measurement).
- The candidate's three personal examples (slides 14 and 16 notes).
- Corporate-event facts should be re-pulled the day before the panel (Q3'26 results and any new announcement would supersede them).
