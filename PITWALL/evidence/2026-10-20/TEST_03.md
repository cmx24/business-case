# TEST_03 - gate run on the CFO Chief of Staff deck v3 (2026-10-20, built 11:10-12:15 EDT)

File under test: `deck/CFO_CoS_deck_v3.pptx` sha256 cf3ee6d687eee7866a58c0ad43ad5b6138277f75e5d1d8ff875110c3dbaf251e
Built by `tools/build_v2.py` from the repaired v1 copy. Criteria: `PITWALL/briefs/deck-v1-gates.md`, unchanged.
Coordinator: the reviewing session, with four parallel agents (fact freshness on the web; unanchored red-team review; fresh-eyes
visual QA of the 22 renders; talk track and timing). Tester's step (dry run): NOT RUN.

## What changed between v2 and v3

From the red-team review (three blockers, six majors, nine minors; all blockers and majors closed except the ledger format):
- One clock: the funding gates follow the FY27 fiscal calendar (gate 1 at FY27 start on 1 Jan 2027, gate 2 at Q1 close on 31 Mar on Q1
  indicators, the A6 valve at H1 close); the 90-day plan runs in tenure days from today (Day 60) and the HQ gates are D+n from the charter
  (new assumption A10). The "37 days to prove Q1 payback" problem no longer exists.
- The A6 valve is stated in the ask on slides 2 and 19 ("$25M FY27 cap in two gates, with a CFO- and IR-gated valve of up to $15M that
  is not in plan"), so the panel is not left to discover a $40M envelope on slide 8.
- In-housing specifics on slide 8: what moves in-house (the annual cycle and the single model, from the business units), hub of 6-8
  people (assumption), partners solid line to the hub and dotted to the GM, in place by Day 95; the model reads the data platform.
- Opex and capital kept apart: slide 10 labels are "of the gate-1 release", the new-bets card no longer lists the US, and the US charter
  capital sits only on slide 12. Slide 10's title says "FY27 opex lock"; slide 12's title no longer contradicts it.
- Slide 17 title separates gate 2 (Q1 indicators) from the A6 valve (H1 beat).
- Slide 14 is labeled "the model" and no longer uses the US or Brazil as examples it cannot verify; signing limits are "by category".
- Pushback slide moved to slide 11, after "settling the two views".
- Minors: "answer $60M of asks" on the title; "I assemble the pack" (no veto); "Strategy, not read-outs"; "DELIVERABLE 4 // LEADERSHIP";
  Capital One reference removed; opex base shown ($4.6B); AI line only once (slide 21).
- Not done (accepted): ledger rows per claim with confidence and sensitivity (M6); the Nu Lens scoring table (M5) beyond the one line on
  data on slide 8. Both are backup-page work; the storyline does not depend on them.

From the fact check (12 facts, all web-verified against Nu's 6-Ks, releases and the OCC letter): one correction, the OCC conditional
approval is dated 29 Jan 2026 (deck said 27 Jan); two refinements, Mexico "authorized July 2026, operating as a bank from 6 Aug" and the
mature-cohort ARPAC now quoted as ~$31 at 96 months vs $15.9 total from the Q1'26 cohort slide. Q3'26 results are not out (due 12 Nov
2026). October news (Monzo not pursued, election runoff 25 Oct, split sell-side, NPL 90+ at 6.9%) is in the notes of slides 12 and 17.

From the visual QA (10 fixes, 5 nits): titles kept inside the 12.0in content grid; two overflowing "Full year" lines; two long labels;
the planning-principle card alignment; stray separators under two tables; the dangling divider on slide 20.

## G1 Scope

`scan_v3.txt`: RESULT 44 pass, 1 fail. Sections A and C pass in full; section B fails only on the candidate's personal example (B5).
Regex note: the "planning KPIs" line was widened from a 80- to a 200-character window after the hub sentence grew; the content is on
slide 8. **PASS.**

## G2 Facts and arithmetic

All 12 Nu and OECD facts CONFIRMED by the fact-check agent against primary sources (6-K filings of 13 Aug, 1 Jun, 4 Jun and 20 Jul
2026; nu.com newsroom; OCC letter ca1362; Investor Day release of 2 Sep 2026), with the OCC date corrected. Arithmetic unchanged from
TEST_02 and still ties. Guard lines 8 of 8 pass. **PASS.**

## G3 Decision logic

(a) one envelope: $25M in two gates, A6 valve named in the ask and marked "not in the base plan" on slides 3, 8, 10, 17, 19. PASS.
(b) one clock: fiscal dates for funding gates, tenure days for the plan, D+n for HQ gates; reconciled on slide 18. PASS.
(c) observable criteria: Q1 indicators judged at Q1 close. PASS. (d) Board topic: slide 12, with A4. PASS. (e) options and
sensitivity: unchanged, verified. PASS. **PASS.**

## G4 Panel readiness

(a) validates: "All validations PASSED!". (b) renders clean: all 22 slides inspected after the v3 build, the two slides changed in
the last pass (14, 15) re-rendered and clean. (c) notes: 22 of 22. (d) dry run: NOT RUN; the talk track measures 3,400-3,500 spoken
words, 26-27 minutes at 130 wpm, which is an estimate. (e) red-team questions: 13 of 15 answerable from slides or notes; two need the
candidate's own examples. **FAIL on (d) only.**

## Still NOT proven

- The timed dry run (B8). - The candidate's three examples (B5). - Q3'26 figures (not published until 12 Nov 2026).
