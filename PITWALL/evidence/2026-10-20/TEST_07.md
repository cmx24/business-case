# TEST_07: deck v7 (2-3 slides per deliverable) and the v7 talk track (2026-10-20)

**Request (pit wall):** "the deck seems very light on deliverable 3. I need all deliverables to have at least 2-3 slides of
focus on the response, with a good assessment, recommendation, rationale/foundation based on data, and raising possible
risks/issues. have Fable assess the scope, the challenges for each pillar and the deliverables in that context to reframe
the deck." Correction during the run: "D3 doesn't have to get the deepest treatment - I just want all of them to have the
appropriate depth."

**Files under test:**
- `review/2026-10-20_fable-reframe.md`: the Fable assessment and specification.
- `deck/CFO_CoS_deck_v7.pptx`: built by `tools/build_v7.py`; sha256 46104000b7e77a6ead34527632c6522518eaf47c4af28a1538e45802377050b1.
- `review/2026-10-20_v7_talk-track-and-QA.md`: renamed from v6 and rewritten.

## Fable's proposal, assessed

**Adopted:**
- The 15-slide storyline: D1 4-6, D2 7-9, D3 10-12, D4 13-14, close 15.
- The per-pillar diagnosis lines: "what is broken today" on slide 4, "why the book is late" on slide 9, and the D3
  assessment on slide 10.
- Risk strips on slides 6, 9 and 12.
- Capacity allocation on slide 6.
- The non-convergence rule on slide 9.
- The site method on slide 10: hard tests, then criteria.
- Gates with "what we negotiate" and walk-aways on slide 11.
- The officer model and the disagreement rule on slide 12.
- The per-problem balance table on slide 14.
- Backups I and J.
- Q42-Q55.

**Adapted or rejected:**

| Fable proposal | Call | Reason |
|---|---|---|
| Hub and HQ run-cost "inside existing central budgets, not the $25M" | REJECTED; replaced by A12 | The brief caps cost growth "across all central and regional operations". A12 says any net new cost is funded inside the $25M before the GM releases; if material, the second release shrinks |
| Slide 8 title "senior executive" | ADAPTED | The title keeps the brief's "senior GM"; the MTeam line goes in the strip |
| Derailer placeholder printed on slide 14 | REJECTED | No placeholder may appear in a file sent to Renata; the slot stays in the talk track |
| Agreement checklist in backup J | ADAPTED | Slide 10 already names the agreement components and slide 11 the walk-aways; backup J holds the scorecard and the risk register |
| Talk-time check in REGRESSION.md | NOT ADDED | The checker reads decks, not scripts; timing is enforced by `tools/talk_timing.py` and recorded here |
| "Treaty hub" as a site option | DROPPED | Vague and unsourced; slide 10 says where Nu already has licensed substance or is building |

## Gates on v7

- **OOXML validation:** PASSED.
- **Strict package check:** no structural problems.
- **Fit check:** no overflow predicted (after four card fixes). The title-collision guard raised no warnings.
- **Render:** LibreOffice converted all 26 slides, and every slide was inspected.
  - Fixed after the first render: oversized dark bands on slides 11 and 14.
  - The preview PDF matches the committed deck byte for byte.
- **Capability scan:** 52 pass, 1 fail (the personal example, B16). See `scan_v7.txt`.
  - Presented-slide limit 14 -> 15 and total 23 -> 26, on the pit wall's instruction.
  - Five new lines: site method, walk-aways, disagreement rule, capacity, per-deliverable risk strips. All pass on v7
    and the last four fail on v6, so they measure the new depth.
  - The risk-strip line was tightened after it first matched "Risk can block" on slide 2.
  - Checker self-test: PASS.
- **Placeholders in the deck file:** 0, counting slide text and notes ("your example", "your watch-out", "Pit wall").
- **Preservation:**
  - v6 to v7 (slides and notes against v7 plus the talk track): 192 tokens, 0 missing.
  - v4 to v7: 1 missing, a format-only date ("29 January 2026" is written "29 Jan 2026").
  - Old v6 talk track to v7: 184 tokens; only the old timing table's minutes are missing.
- **Talk track:** 3,129 words, about 24.1 minutes. No sentence is over 30 words.
  - Opening 3.6 min; D1 5.5; D2 5.5; D3 4.9; D4 3.4; close 1.2.
  - Checkpoints: slide 7 at 9.1 (target 10), slide 10 at 14.6 (16), slide 13 at 19.6 (21), slide 15 at 22.9 (24).
- **Q&A:** 55 questions, each with why they ask, position, spoken answer, tone and the slide to show. Every slide
  reference was renumbered to v7 and checked.

## Still NOT proven

- That PowerPoint opens v7 (same packaging as v3-v6).
- The personal stories, the derailer, the Visa specifics, the Q55 example (B16, B22).
- The dry run (B8).
- A12 and A13 (to confirm with the CFO).
- Jurisdiction-specific officer liability (Legal owns it per site; the slide says so).
