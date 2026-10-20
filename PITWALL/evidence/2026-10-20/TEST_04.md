# TEST_04 - PitWall review of the external feedback on deck v3 and the talk track (2026-10-20)

Input: an external LLM review titled "Master LLM input payload" (assessment, red-team matrix, slide-by-slide tracked edits, a
python-pptx reorder script, and a fully rewritten talk track with personal examples filled in). PitWall rules applied: every claim is
checked against the files before it is accepted ("the evidence wins"); an item whose evidence fails is rejected even if it sounds right;
anything only the pit wall can know goes to the board.

Files under test: `deck/CFO_CoS_deck_v3.pptx` (input) -> `deck/CFO_CoS_deck_v4.pptx` (output, built by `tools/build_v2.py`);
`review/2026-10-20_talk-track-and-QA.md` (revised in place).

## Verdict per item

| # | Feedback item | Evidence check | Verdict | What changed |
|---|---|---|---|---|
| A | "Physical slide order is wrong" (7 slides) plus a reorder script | Read the deck's presentation order (`sldIdLst`, what PowerPoint shows): slides 1-22 run 01 // to 17 // then the close and the appendix, in order. The reviewer's "physical" positions match the order of the internal relationship IDs, which reflect build order for the three cloned slides. Ran the reviewer's script on a copy of v3: it moved the appendix divider to slide 3, the capital slide to 9, the ledger to 12, the close to 19 and "How I lead" to 20. | **REJECT. Do not run the script.** | Hardening only: v4 renumbers the relationship IDs so they ascend in presentation order (rId2-rId23), so no tool can misread the order again. Display order unchanged. |
| B | Hub 6-8 -> 4-6 people, reframed as unblocking squads; cycle time 40 -> 30 days | Hub size is an assumption either way (A10, "sizing is mine"); the smaller cell fits Nu's "smart efficiency" value and the squad framing. The 30-day target has no basis anywhere in the deck or the sources; the only benchmark in the ledger (directional) supports 40, and this is the first in-house cycle. | **ACCEPT (size, framing) / REJECT (30 days)** | A10, slide 8, talk slide 8, Q15, numbers sheet: lean 4-6 person hub that unblocks squads; cycle time stays under 40 days, baseline measured in week one. |
| C | Fallback for the $8M platform estimate | Valid; the rule was already in the talk track and notes but not on the slide. The reviewer's added sentence duplicated the existing one. | **ACCEPT (merged, not duplicated)** | A5 and slide 7 caption: scope Engineering cannot confirm moves to the case track and the cap does not grow. Talk slide 7: one clause with the $5M/$3M example. |
| D | Add Risk to the A6 valve sign-off; NPL 90+ condition | Valid governance gap: the valve is conditioned on credit metrics, so Risk should hold a veto. The reviewer's slide 17 wording ("locked, requires joint sign-off") would imply the valve can open with NPL above plan; rejected that wording. The "<=20%" hard threshold was kept as "inside the ~20% guide", matching the source. | **ACCEPT (with wording fix)** | Slide 2 ask, A6, slide 8, slide 17 row 3 ("A6 stays shut; Risk holds a veto"), speaker notes of slides 8 and 17, talk slides 2, 8, 17, Q2, numbers sheet. |
| E | OCC capitalization deadline; "this Board is the final statutory window" | The 12-month condition from the 29 Jan 2026 approval is verified (OCC letter ca1362, fact check). "Final statutory window" is not supported: no rule makes this meeting the last chance (special meeting or written consent remain), and FDIC and Fed approvals were still pending at the last check. | **ACCEPT (date) / REJECT ("final statutory window")** | Slide 12 card 3: "the US charter must be capitalized by about 29 Jan 2027 (OCC condition)"; speaker notes carry the ~10-week arithmetic and the caution; talk slide 12 and numbers sheet. |
| F | Relabel A2 as management guidance | Verified: the fact check found the guidance in the Q2'26 call transcript (13 Aug 2026). "US organic run-rate spend" is not the source's wording; the call says the US entry stays within about 100 bps of efficiency-ratio impact. | **ACCEPT (source wording)** | A2: "MANAGEMENT GUIDANCE, Q2'26 earnings call (13 Aug 2026)"; Q6 updated. |
| G | Rewrite of the slide 2 cap sentence ("our twenty percent guidance ... reallocates dynamically") | Drops the case brief's anchor (the CFO's cap for public-market expectations and margin discipline), turns "about 20%" into a hard number, and adds a filler word. | **REJECT** | Kept the current sentence; only Risk added to the ask. |
| H | Personal examples written into slides 11 and 16 and Q7 | These are specific career claims (a $10M VP ask, interchange-rebate double count, an uncapped indemnity) written by a third party. Nothing in the files shows they happened; presenting invented history at an executive interview is the highest-severity risk in the package. | **REJECT (held for the pit wall)** | Placeholders kept. The three candidate stories are copied to the board (B16) for you to confirm, edit to the true version, or replace. |
| I | Red-team matrix (three challenges) | Challenges 1 and 3 are already answered (Q2; slide 12 and Q3). Challenge 2, "you are cutting asks by 58%", had no prepared answer; 58% = $35M / $60M checks. | **ACCEPT (challenge 2)** | New Q16. |
| J | Interruption protocol | Useful and consistent with the timing table. | **ACCEPT (simplified)** | New section 2b: jump rule from slides 4-6, drop the examples on slide 16 if behind, hard stop at minute 28 goes to the ask. |
| K | Q10 addition ("operational force multiplier") | Point is fair; the phrase is jargon. | **ACCEPT (plain wording)** | One sentence: AI made research and scenario testing faster; judgment and accountability stay with me. |
| L | Timing "28.9 minutes, 3,756 words" | Their count includes the invented examples and the rejected rewrites. | **REPLACED** | Recounted from the text: 3,631 words, 27.9 minutes at 130 wpm, after trimming three filler phrases. |

## Gate run on v4 (sha256 b0547526386d64a1032d0aa91aded343cc956024535df580fed1ec4f202960c3)

- OOXML validation: "All validations PASSED!"
- Strict package check (content types, every relationship target and r:id resolves, unique shape ids, no control characters, no
  embedded fonts): clean.
- Presentation order: 22 slides in narrative order; slide relationship IDs rId2-rId23 ascend in that order.
- Render: all 22 slides converted; the six changed slides (2, 3, 7, 8, 12, 17) inspected, no overflow or overlap.
- Capability scan: 44 pass, 1 fail (unchanged: the personal example, B5/B16).
- Talk track: no gendered pronouns, no banned words, 7 example placeholders, 27.9 minutes.

## Still NOT proven

- That PowerPoint opens v4 (no PowerPoint here; same package build as the v3 fix, which you have not yet confirmed).
- The three personal examples (B16). - The dry run (B8). - Q3'26 figures (12 Nov 2026).
