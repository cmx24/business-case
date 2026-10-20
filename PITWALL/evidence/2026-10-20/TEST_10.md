# TEST_10: final pass on the candidate's own restructure (deck v10) (2026-10-20)

**Input:** the candidate's uploaded file (21 slides: 14 presented, divider, backups A, B, D, F, H, J), kept as
`deck/upload_final/uploaded.pptx`. It is the candidate's rewrite of v9, not v9 itself. Instruction: "take a final pass.
make minor tweaks to ensure cohesion and depth. ensure numbers and data are included where applicable."

**Output:** `deck/CFO_CoS_deck_v10.pptx`, produced in place by `tools/final_pass_upload.py` (the candidate's layout is kept;
text is replaced at run level). (final: reserve $10M; no speaker notes anywhere, on the pit wall's instruction).

## What changed (all minor, in the candidate's voice)

- **Cohesion of the funding story (slide 6 and backup A).** The chart's reserve bar was drawn at $10M while labeled $5M; the pit wall set the reserve to $10M, so bar, label, card, footer and backup A now agree. The 2H checkpoint was "decided on Q4 data" for a 3Q-4Q release, which cannot happen; it now reads
  "decided at the H1 close on H1 data", with the reserve conditions the assumptions table already carries (profit above plan,
  the efficiency-ratio guide in force, NPL 90+ on plan; the CFO decides, Risk can block). Backup A's envelope row now matches
  (second release on Q1 data; checkpoint up to $25M and reserve $5M at H1 close). A footer states the arithmetic: the two
  releases are the cap; the checkpoint and the reserve sit outside the plan; every dollar of the $60M has a path ($15M + $10M + $25M + $10M); and the cohort data behind 70/30.
- **Numbers added where the slide had none.** Slide 2's success measures now carry targets (100% of discretionary spend with
  a case, cycle time under 40 days, book at T-7 every quarter, HQ gates 1-2 by mid-December, GM rating twice a year) and a
  one-line statement of the decisions requested, which the restructure had lost with the old closing slide. Slide 10 got
  back the Pillar Two footer (Cayman incorporation, EUR 750M, 15%, 9.4% / 7.4%, Brazil top-up from 2025). Slide 3's KPI is
  now 100% of discretionary spend, not "% of relevant spend".
- **Depth restored on D3.** Slide 11 now states what happens when functions disagree (both views in the log, to the CFO
  within 48 hours with a recommendation; legality vetoes stay with Legal and Compliance) and that gate decisions are taken
  at the monthly SteerCo chaired by the CFO.
- **Consistency.** "CoS & GC" / "GC Team" to "CoS & CG" throughout; "Mteam" to "MTeam"; "120d algo approach" to "120-day
  gated approach"; "go/non-go" to "go/no-go"; grant grace 12-24 months everywhere (slide 12, backups F and J); "I" to "we"
  on slide 8; "me" to "the CoS" in backup J; typos fixed (ships / abd / "better even ."); page numbers now run 1-21 (they
  ran 1, 2, 6, 4, 5, 5, 7, 9, 8, ...); the divider lists only the six backups that exist; backup D's "where used" column
  points at the new slide numbers; backup F no longer cites the removed $12M case track or squad clinics.
- **Speaker notes:** removed from every slide (slides only, per the pit wall). Slide 13's footer no longer promises them.

## Not changed, on purpose

- The candidate's new funding structure ($15M + $10M inside the cap; up to $25M at a 2H checkpoint; $5M reserve) and the
  removal of the capital-ranking slide, the options scoring columns, the closing slide and backups C, E, G, I. Those are the
  candidate's calls. Two consequences are reported, not reversed: the capital lens (return on allocated capital against cost
  of equity) is no longer on any slide, and "pushback" as a word is gone (slide 9 covers the content).
- The McKinsey / Deloitte / DirectorPrep references the candidate added on slides 5, 7 and backup D were not verified here.

## Gates

- OOXML validation: PASSED. Strict package check: no structural problems.
- Render: all 21 slides rendered and inspected; four fit defects introduced by longer text were fixed (slides 2, 3, 6, 11).
- Capability scan: 49 pass, 4 fail. The four: GM pushback wording; the capital lens; first-person D4 voice (the deck now
  speaks as "the team" without "we work / we bring"); the personal example. The first two follow from the candidate's
  restructure; the last is still the candidate's to write.
- Preservation (upload to v10): 118 tokens; 3 differ, all intentional ($12M case track removed from backup F; "A1-A13"
  label no longer used because the table has no A-numbers).
- Placeholders in the file: none.

## Still NOT proven

- That PowerPoint opens v10 (the upload came from PowerPoint, so the packaging risk is lower than for the generated decks).
- The personal stories and the dry run (B16, B8).
