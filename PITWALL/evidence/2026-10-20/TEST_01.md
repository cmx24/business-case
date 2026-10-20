# TEST_01 - gate run on the CFO Chief of Staff deck v1 (2026-10-20)

Files under test (sha256):
- original, as received: `deck/CFO_CoS_deck_original.pptx` c053b402c09253585c88f63be7b92cd1bdce792b9fe1a67e0527f5e0dc26b265
- repaired copy (mechanical fixes only, wording unchanged): `deck/CFO_CoS_deck_repaired.pptx` 0012dc44ca71f9798e9fa3820b09e77fc3664d264186cd9dec73a6c692e0140e

Criteria: `PITWALL/briefs/deck-v1-gates.md` (written before the checks below were run against the repaired copy; the original was
scanned once before the regex fixes noted under G1, and that first run is kept in the git history of this folder).
Coordinator: the reviewing session. Tester's step (dry run): NOT RUN, nobody was available to time it.

## G1 Scope (capability scan)

Command: `python3 -I tools/deck_capability_check.py deck/<file>.pptx` (validator path via `PPTX_VALIDATOR`). Full outputs:
`scan_original.txt`, `scan_repaired.txt` (this folder). Exit code 1 on both (fail lines present).

| Line group | Original | Repaired |
|---|---|---|
| Section A (from the brief, 21 lines) | 20 PASS, 1 FAIL | 20 PASS, 1 FAIL |
| Section B (reviewer's list, 16 lines) | 3 PASS, 13 FAIL | 4 PASS, 12 FAIL |
| Section C (guard, 8 forbid lines) | 0 PASS, 8 FAIL | 0 PASS, 8 FAIL |
| RESULT | 23 pass, 22 fail | 24 pass, 21 fail |

Section A failure: "assumptions are listed and separated from facts" (no slide matches). This is an explicit instruction of the brief.

Regex corrections made between the first and second run, recorded as the gate brief requires: the US-launch line had a comma inside its
regex (split by the grammar) and was rewritten; "personal example" matched the bare word "at"; "capital lens" matched "roac" inside
"approach"; "reversal triggers" matched "triggers corporate tax"; "planning KPIs" matched Hackett prose. Each was tightened, which turned
four false PASSes into FAILs. No line was loosened.

**G1 outcome: FAIL** (one section-A line). Branch: add the assumptions slide (draft in the review, section 6); re-run.

## G2 Facts and arithmetic

Fact table (Nu figures on the slides vs. the Q2'26 release; Q3'26 not yet published on the review date):

| Slide | Deck | Latest public figure | Status |
|---|---|---|---|
| 3, 18 | ~$0.80 monthly cost to serve | $1.0 (Q2'26); $0.8 was Q2'25 | FAIL (stale) |
| 3, 18 | 140M+ users | 138.9M customers (Q2'26) | FAIL (rounded above the reported figure) |
| 3, 18 | 19.5% efficiency ratio | 19.5% (Q2'26); ~20% expected FY26 | PASS |
| 8 | "mature cohorts ... 32%+ return on equity" | Group ROE 33% (Q2'26); ROE is not disclosed per cohort | FAIL (wrong attribution) |
| 8, 18 | $27+ mature-cohort ARPAC | Disclosed in earnings materials; exact figure and quarter not verified here | UNVERIFIED (cite the page) |
| 8, 18 | $17 company ARPAC | $17.1 (Q2'26) | PASS |
| 10 | OECD 2026 transitional SBIE 9.4% payroll / 7.4% tangible assets | Matches the Art. 9.2 schedule | PASS (rate); FAIL (description: it is an exclusion from the top-up base, not a "backing" requirement) |
| 10, 14 | "$15.6M saving" / "$15.6M HQ 5-year net value" | No model or source anywhere in the deck | FAIL (unsupported) |
| 14 | 3-year value $16.8M / $36.4M / $61.2M; HQ value $6.2M / $21.4M | No model or source | FAIL (unsupported) |

Arithmetic re-run in code (python, this session):

```
weighted scores: 57.75 / 63.75 / 91.75  (slide 5 shows 57.8 / 63.8 / 91.8)        TIES
bridge: 60 - 8 - 12 - 15 = 25                                                     TIES
stage 1: 10.5 + 4.5 = 15.0; base year 16.5 + 8.5 = 25.0; H2 5 + 10 = 15          TIES
Q2 split: core 6.0, new bets 4.0; new-bets Q2+H2 claim 14.0                        TIES
upside total spend: 25 + 15 = 40 vs cap 25                                          BREACH (see G3)
materiality vs Q2'26 annualized revenue ~$23.5B, implied opex ~$4.6B:
  $25M = 10.6 bps of revenue, 0.55% of opex; $35M = 14.9 bps, 0.76%; $60M = 25.5 bps, 1.31%
gate clocks: project day 30/60/90/120 -> tenure day 90/120/150/180 (slide 15 says 120/135/150 for gates 2-4)   MISMATCH
```

Guard (section C): 8 of 8 forbid lines hit on both files (seamless x2, "connective tissue", "data proves" x2, $0.80 cost to serve x2,
140M+ x2, cohort ROE, "Independent Audit Verdict", "Margin Neutral").

**G2 outcome: FAIL.** Branches: replace stale figures with Q2'26 values and sources; delete or model the five unsupported value
figures; fix the OECD description; clear the guard lines (board decision B7 for the benchmark cull).

## G3 Decision logic

| Criterion | Finding | Status |
|---|---|---|
| (a) one envelope everywhere | Slides 7 and 14: $15M + $10M = $25M, then "up to $15M in H2" = $40M; slide 14 calls it "100% margin neutral" with no rule | FAIL (board decision B1) |
| (b) one clock | Slide 10 uses project days; slide 15 uses tenure days; gates 3 and 4 land 15 and 30 days early on slide 15 | FAIL |
| (c) observable gate criteria | "90% of Q1 customer payback" judged on Day 125, 37 days after the Q1 release on Day 88 | FAIL |
| (d) Board topic answered or scoped out | Board topic is international capital allocation across entities; deck answers with a 70/30 split of a $15M opex tranche; no assumption stated | FAIL (board decision B2) |
| (e) three options, weights, sensitivity | Three options with weights on slide 5; no weight sensitivity stated; scores are unlabeled judgment | PARTIAL (add the ±10 pp sensitivity line and the ASSUMPTION label) |

**G3 outcome: FAIL.**

## G4 Panel readiness

| Criterion | Original | Repaired | Evidence |
|---|---|---|---|
| (a) validates and opens | FAIL: `p:regular` after `p:bold` in `embeddedFontLst`; LibreOffice: "source file could not be loaded" | PASS: validator "All validations PASSED!"; LibreOffice converts | validator output in scan files; render PDF produced |
| (b) renders clean | FAIL: slide 3 overflow, slide 7 overlap, slide 19 clipped, cards misaligned on 2, 11, 13, 15 | PASS: all 20 slides re-rendered and inspected after the fixes | render inspection, this session |
| (c) notes on every core slide | FAIL: 0 of 20 | FAIL: 0 of 20 | `check: notes` |
| (d) dry run 25-30 min | NOT RUN | NOT RUN | INCONCLUSIVE = FAIL (board decision B8) |
| (e) red-team answers on slides | 3 of 10 answerable from the slides (ask; pre-Board routine; GM pushback steps) | same | review section 7 |
| raster chart | FAIL: slide 6 is a picture | FAIL | `check: no_raster_charts` |

**G4 outcome: FAIL** on both files. The repaired copy clears (a) and (b) only.

Text-set check for the repair (the "never change wording" rule): markitdown text of original and repaired, sorted, identical; the
unsorted diff shows reading-order changes only where boxes were moved. Verified this session.

## Still NOT proven after this run (said plainly)

- That the deck fits 30 minutes (no dry run).
- That the presenter can answer the red-team questions from the slides (7 of 10 have no slide).
- The exact mature-cohort ARPAC figure and its quarter (not verified here).
- Any of the eight board decisions (none recorded).

## Next run

TEST_02 on v2: same four gates, same commands; add the dry-run time and the sha256 of the file that goes to Renata.
