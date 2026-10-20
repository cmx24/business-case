# Gate brief: CFO Chief of Staff deck, version 1 (the file received 2026-10-20)

Four gates, in order. Each gate has a pre-registered pass criterion (written before the run, not edited after), one check per
pass/fail point, the evidence it produces, and a branch table with the action per outcome. INCONCLUSIVE counts as FAIL.
Results are recorded in `PITWALL/evidence/2026-10-20/TEST_01.md`; the status column here is copied from there after the run.

Ship rule: the deck goes to Renata only when G1-G4 are all PASS on the same file (same sha256). Anything not proven at G4 is
either removed from the deck or labeled on the slide as an assumption.

## G1 Scope: does the deck answer the brief?

- **Criterion (pre-registered):** every `capability:` line in section A of `PITWALL/REGRESSION.md` passes; `max_slides` passes.
- **Check:** `python3 -I tools/deck_capability_check.py deck/<file>.pptx` (symbol scan; under 10 s; no judgment of quality).
- **Evidence:** the scan's PASS/FAIL lines and RESULT line, pasted into TEST_<nn>.md.

| Outcome | Branch |
|---|---|
| All section-A lines PASS | G1 PASS; go to G2. |
| One or more section-A lines FAIL | Add the missing content (the line names it); re-run; a line that keeps failing after two builds goes to the board. |
| A line fails because the regex is wrong, not the deck | Fix the line in REGRESSION.md, note it in the evidence file, re-run. Never loosen a line to make a real gap pass. |

## G2 Facts and arithmetic: is every number current and reproducible?

- **Criterion:** (a) every Nu figure equals the latest published quarter (Q2'26 on the review date; Q3'26 once released);
  (b) every calculation on a slide is re-run in code and ties to the slide within rounding; (c) every figure that is not a FACT with a
  source is labeled ASSUMPTION or ESTIMATE on the slide or the assumptions slide; (d) all `forbid:` guard lines in section C pass.
- **Check:** fact table (deck value vs. release value, with the source); the arithmetic script in the evidence file; the scan's section C.
- **Evidence:** the fact table and the script output in TEST_<nn>.md.

| Outcome | Branch |
|---|---|
| (a)-(d) all hold | G2 PASS. |
| A Nu figure is stale | Replace with the latest figure and its source; re-run. No deck goes out quoting a superseded quarter. |
| A calculation does not tie | Fix the slide or the model; if the number has no model behind it (e.g. "$15.6M saving"), delete it or move it to the assumptions slide as an ESTIMATE with its basis. |
| A third-party statistic cannot be sourced to a document you have read | Remove it from headlines; keep at most four sourced benchmarks (board decision B7). |

## G3 Decision logic: is the recommendation consistent and does it answer the decision the panel will judge?

- **Criterion:** (a) the funding envelope is one number everywhere (no path exceeds the cap without a stated, agreed rule);
  (b) one clock for every timeline (gate days and tenure days reconcile); (c) the gate criteria are observable within the gate's own
  window; (d) the Board's actual topic (capital allocation across entities) is either answered or explicitly scoped out on the
  assumptions slide; (e) at least three options compared, one recommended, with weights disclosed and a weight sensitivity stated.
- **Check:** reviewer's consistency pass on slides 2, 6, 7, 8, 10, 14, 15 plus the capital-vs-opex read of slide 8.
- **Evidence:** the consistency table in TEST_<nn>.md.

| Outcome | Branch |
|---|---|
| (a)-(e) all hold | G3 PASS. |
| Cap breached in any scenario | Board decision B1: formula-sized reinvestment rule, or delete Stage 3. |
| Two clocks | Relabel all gate days to one clock (recommended: "D+n from charter") and reconcile slide 15. |
| Gate criterion not observable in its window | Replace with leading indicators (activation, 30-day retention, CAC vs plan, unit cost vs plan); payback becomes a lagging confirmation. |
| Capital topic unanswered | Board decision B2: add the capital slide or scope it out explicitly. |

## G4 Panel readiness: does it open, render, fit 30 minutes and survive Q&A?

- **Criterion:** (a) the .pptx validates against the OOXML schema and opens without repair in PowerPoint and LibreOffice;
  (b) every slide renders with no overflow, overlap or clipping; (c) every core slide has speaker notes with the hardest question and its
  answer; (d) a timed dry run in front of a stand-in panel lands between 25 and 30 minutes (the Tester's step; one run, time-boxed, with
  the stopwatch time as the evidence); (e) each of the ten red-team questions in the review has an answer that is already on a slide.
- **Check:** validator; full render and inspection; `check: notes`; the dry run; the Q&A walk-through.
- **Evidence:** validator line, render notes per slide, dry-run time, Q&A table in TEST_<nn>.md.

| Outcome | Branch |
|---|---|
| (a)-(e) all hold | G4 PASS; the file (by sha256) is the one sent to Renata 24 h before the panel. |
| Schema or render defect | Fix mechanically without touching wording (prove the text set is unchanged), re-validate, re-render the changed slides only. |
| No notes | Write them; the note for each slide is the oral takeaway plus the hardest question and its answer. |
| Dry run not done or outside 25-30 min | INCONCLUSIVE = FAIL. Cut or merge slides until it fits (the review's 14-slide storyline is the fallback), re-run. |
| A red-team question has no answer on a slide | Add the answer to the slide or the notes; an answer that only exists in your head is not proven. |

## Status after the 2026-10-20 run (copied from TEST_01.md)

| Gate | Status on v1 (original file) | Status on v1 repaired copy | Main blocker |
|---|---|---|---|
| G1 Scope | FAIL | FAIL | No assumptions slide (an explicit instruction of the brief). Everything else in section A passes. |
| G2 Facts and arithmetic | FAIL | FAIL | Cost to serve $0.80 (Q2'25) instead of $1.0 (Q2'26); 140M+; cohort ROE; three value figures with no model; eight guard lines hit. Arithmetic that exists ties. |
| G3 Decision logic | FAIL | FAIL | $40M upside spend vs $25M cap labeled "margin neutral"; two clocks on slides 10 and 15; Q2 gate not observable in 37 days; capital topic unanswered. Option set and weights pass. |
| G4 Panel readiness | FAIL | FAIL | Original: schema invalid, four render defects. Repaired: schema valid, renders clean, but no notes on any slide, raster chart on slide 6, dry run NOT RUN (INCONCLUSIVE = FAIL). |

Still NOT proven after this run, said plainly: that the deck fits 30 minutes; that the presenter can answer the ten red-team questions
from the slides; that any of the eight board decisions has been taken.

## Status after the v2 build (copied from TEST_02.md, 2026-10-20 11:05 EDT)

| Gate | Status on v2 (`deck/CFO_CoS_deck_v2.pptx`, sha256 9e79926f...) | Remaining blocker |
|---|---|---|
| G1 Scope | PASS | None in section A. Section B: the personal example is the candidate's to add (B5). |
| G2 Facts and arithmetic | PASS | Re-pull corporate-event facts the day before the panel. |
| G3 Decision logic | PASS | None. |
| G4 Panel readiness | FAIL | Dry run not run (B8). Everything else passes. |

Ship rule unchanged: the file goes to Renata only when G4 passes on the same sha256.

## Status after the v3 build (copied from TEST_03.md, 2026-10-20 ~12:00 EDT)

v3 adds four parallel checks (web fact-freshness, unanchored red-team review, fresh-eyes visual QA, talk track with timing) and their fixes.

| Gate | Status on v3 (`deck/CFO_CoS_deck_v3.pptx`) | Remaining blocker |
|---|---|---|
| G1 Scope | PASS | Section B: the personal example remains the candidate's (B5). |
| G2 Facts and arithmetic | PASS | Every Nu fact web-verified; the OCC date corrected to 29 Jan 2026; Q3'26 results due 12 Nov 2026 (re-pull then). |
| G3 Decision logic | PASS | Red-team blockers closed: one clock (FY27 fiscal calendar, A10), the A6 valve stated in the ask, in-housing specifics on slide 8. |
| G4 Panel readiness | FAIL | Dry run still not run (B8). Talk track measures 26-27 minutes at 130 wpm, which is an estimate, not a measurement. |
