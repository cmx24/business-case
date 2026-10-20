# Board: pit wall decisions for the CFO Chief of Staff deck

Only the pit wall (Cassio) can take these. Each row names the call, the reviewer's recommendation, and the pre-decided branch per answer,
so the next build never stops to ask. Record the answer as "Pit wall HH:MM: <answer>" next to the row, the way Turbo's code cites it.

| # | Decision | Recommendation | If YES | If NO | Blocks |
|---|---|---|---|---|---|
| B1 | Keep the Stage 3 "up to $15M in H2" pool? | Keep it only as an efficiency-ratio-neutral reinvestment rule agreed with the CFO and IR (sized by formula, e.g. 50% of H1 profit above plan, cap $15M). | Rewrite slides 2, 6, 7, 8, 14 with the formula; the cap line reads "$25M hard cap + conditional reinvestment inside guidance". | Delete Stage 3 from slides 6, 7, 8, 14; the bridge ends at $25M with $15M "not funded this year". | G3 |
| B2 | Add a capital-allocation lens (equity by entity, cost of equity, US capitalization, buyback) to answer the Board's actual topic? | Yes, one slide, labeled ASSUMPTION where the case brief is silent. | New slide 9 per the storyline in the review (section 5). | State on the assumptions slide: "I read the MTeam dispute as the growth-opex envelope; capital allocation by entity is out of scope here." | G3 |
| B3 | Keep the AI usage disclosure (Appendix C) and the "Independent Audit Verdict" line? | Cut both; keep one honest line in the sources footer. | Keep the table but remove model names and the "Human Executive Review" column; delete the audit verdict. | Delete slide 19 and the verdict line on slide 18; add the footer line. | G4 |
| B4 | Which jurisdiction(s) does the HQ slide name? | Name the options you would actually consider and mark the pick as ASSUMPTION; the brief says "regional hub/campus" and "multiple jurisdictions". | Slide 10 gets an options row with constraints per jurisdiction. | Slide 10 stays jurisdiction-neutral and says so in one line. | G3 |
| B5 | Deliverable 4 in first person with your own examples (three)? | Yes. Pick three: one on getting a senior leader to accept less money, one on leading without authority, one on hands-on FP&A or governance rigor. | Replace slides 11-13 with one slide; move the cadence table to D1 and the pushback playbook to D2. | Keep three slides but add one personal example per slide and switch to "I". | G4 |
| B6 | Voice of the whole deck: "I" or "our office"? | "I" for recommendations, the ask and D4; "the planning function" or "the office" for the operating model. | Edit titles and the ask accordingly. | Keep "we", but D4 must still be "I". | G4 |
| B7 | Replace every third-party benchmark except four you can discuss? | Yes. Keep: Nu filings; OECD Pillar Two transitional rates; the fair-process paper (Kim and Mauborgne); one planning benchmark you have read in full. | Strip the rest from headlines and label the remainder "directional" in the ledger. | Keep them, but move every statistic off headlines into the ledger with a "not load-bearing" tag. | G2 |
| B8 | Who does the 30-minute dry run (the Tester's step of G4) and when? | A stand-in panel, timed, at least 24 hours before the file goes to Renata. | Record it in `PITWALL/evidence/<date>/TEST_02.md`. | G4 stays INCONCLUSIVE, which counts as FAIL; the deck is not "proven" for the panel. | G4 |

Open as of 2026-10-20: all eight. No answer has been recorded yet.

**Provisional branches taken in the v2 build (2026-10-20 10:05-11:05 EDT), because the pit wall asked for as much as possible before
13:00 and these were the recommended branches.** Each is a single edit in `tools/build_v2.py` to reverse; record "Pit wall HH:MM" next
to the row to confirm or reverse.

| # | Branch taken in v2 | Where |
|---|---|---|
| B1 | YES: rule A6 (formula-sized, inside the FY26 guide, CFO and IR sign-off, not in the base plan) | slides 3, 8, 10, 17 |
| B2 | YES: capital-allocation slide added, A4 states the reading of the dispute | slides 3, 11 |
| B3 | NO (cut): AI slide and audit verdict removed; one line kept on slides 20 and 21 | slides 20, 21 |
| B4 | NO: jurisdiction left open and said so (A8) | slides 3, 15 |
| B5 | YES: one first-person slide; the three examples are still yours to add, slots in the notes of slides 14 and 16 | slide 16 |
| B6 | YES: "I" for recommendations, the ask and D4 | throughout |
| B7 | YES: four sources kept; the rest named in footers only | slides 5, 12, 21 |
| B8 | NOT DONE: the dry run needs you and a stopwatch | TEST_02 G4 |

**Added in the v3 build (red-team and fact-check fixes, recommended branches taken):**

| # | Decision | Branch taken in v3 | Where |
|---|---|---|---|
| B9 | Which clock do the funding gates follow? | FY27 fiscal calendar: gate 1 at FY27 start (1 Jan 2027), gate 2 at Q1 close (31 Mar), A6 valve at H1 close; tenure days for the 90-day plan; HQ gates D+n from the charter (A10) | slides 2, 3, 8, 17, 18, 19 |
| B10 | Is the A6 valve stated in the ask itself? | YES: "$25M FY27 cap in two gates, with a CFO- and IR-gated valve of up to $15M that is not in plan" | slides 2, 19 |
| B11 | Hub size and reporting lines? | Hub of 6-8 people (A10, assumption); partners solid line to the hub, dotted to the GM; in place by Day 95 | slides 3, 8 |
| B12 | Where does the GM-pushback slide sit? | After "settling the two views" (slide 11), before the capital question | slide 11 |
| B13 | Keep the Capital One reference on the planning-principle slide? | NO (read as flattery); replaced with "Nu's own test-and-learn discipline" | slide 5 |
| B14 | Mature-cohort ARPAC figure? | YES, now that it is sourced: ~$31 at 96 months vs $15.9 total, Q1'26 earnings presentation cohort slide | slide 10 |

**Added after the PitWall review of external feedback (TEST_04, v4 build):**

| # | Decision | Branch taken in v4 | Where |
|---|---|---|---|
| B15 | Hub size and cycle-time target | 4-6 people (was 6-8); cycle time stays under 40 days, because 30 has no basis | A10, slide 8, Q15 |
| B16 | The three personal examples the external review wrote for you | NOT inserted. Use them only if they are true; otherwise replace them with your own, at about the same length (35-45 words each). The candidates, verbatim from the review: (1) slide 11 / Q7: "In commercial finance, when an enterprise VP demanded ten million dollars upfront for an unproven channel, I locked two million for initial activation and gated the remaining eight million on cohort CAC beating our target; they delivered the proof and unlocked the balance cleanly." (2) slide 16, fair process: "At an enterprise financial unit, publishing a transparent project scoring matrix upfront resulted in losing divisions fully supporting the final prioritization without post-decision appeals." (3) slide 16, no ambushes: "When a division lead's revenue forecast double-counted interchange rebates, I reviewed it with them bilaterally the evening prior, correcting the baseline quietly before the C-suite review." (4) slide 16, rigor: "In cross-border partner negotiations, I caught an uncapped liability indemnity in an affiliate schedule, redrafting the carve-out before signature and insulating the balance sheet." | talk slides 11, 16; Q7 |
| B17 | Risk in the A6 valve sign-off | YES: joint CFO, IR and Risk sign-off; NPL 90+ above plan keeps the valve shut | slides 2, 3, 8, 17 |
| B18 | Call the November Board "the final statutory window" for US capitalization? | NO: state the date (capital in by about 29 Jan 2027) and the ~10 weeks; the stronger claim is not supported | slide 12 and notes |

**Added in the v5 build (three-agent review applied, TEST_05):**

| # | Decision | Branch taken in v5 | Where |
|---|---|---|---|
| B19 | Presented length | 14 slides plus backups A-G; talk track 3,250 words (about 25.0 minutes) | deck v5, talk track section 2 |
| B20 | Decision rights for the H2 reserve | The CFO decides; Risk can block; IR sets disclosure (replaces "joint CFO, IR and Risk sign-off", B17) | slides 2, 6, 13; A6 |
| B21 | Deliverable 2 order | Capital ranking first; the 70/30 opex split second, on the same slide | slide 8 |
| B22 | Derailer | NOT inserted: a slot on slide 12 and in Q19, with guidance on choosing a true one. Yours to write, like B16 | talk slide 12, Q19 |

**Pit wall input, 2026-10-20 (after v5):** the Deliverable 4 anecdote. You have seen big companies at local, regional and
global level: rigid governance and bureaucracy at Citi; Itaú BBA's agile, client-centric DNA slowed by group politics and
too many cooks; and Visa's tech-company essence now coexisting with the governance a global-scale business needs. It is
used as the opening of talk slide 12 and in the new Q37 and Q38. "Politics" is left out of the spoken wording; the script
says "too many cooks in the kitchen" instead (reason in the slide 12 delivery note). This does not close B16: the three
rule examples, the derailer (B22) and your Visa specifics in Q38 are still yours to write.

**Added in the v6 build (TEST_06):**

| # | Decision | Branch taken in v6 | Where |
|---|---|---|---|
| B23 | AI in the operating model | Two agents: a PMO agent (your example: status without templates) and a review agent (case checks, Board-book and directors' drafts). Every figure is signed by a named person; Nu-approved tools only (A11) | slides 7, 10, 11, 12, 14; backup H; Q39-Q41 |
| B24 | Where the AI lines sit in the talk | Spoken on slides 7, 10, 11, 12 and 14. To make room, five lines moved to Q&A or to the slides | talk track; timing 25.4 min |

**Added in the v7 build (TEST_07):**

| # | Decision | Branch taken in v7 | Where |
|---|---|---|---|
| B25 | Slides per deliverable | 2-3 each, on the pit wall's instruction: D1 3, D2 3, D3 3, D4 2; 15 presented (supersedes B19's 14) | deck v7 |
| B26 | Hub and HQ net new cost against the cap | Inside the $25M (A12), funded before the GM releases; if material, the second release shrinks. Fable's "outside the cap" proposal rejected | slide 3, slide 6, backup A, Q49 |
| B27 | HQ site choice | Four hard tests, then five weighted criteria (weights are proposals, A13); two sites to term sheet; the CEO and CFO pick; no site named | slide 10, backup J, Q42 |
| B28 | Who signs for the HQ | Licensed executives after gate 3; the Chief of Staff holds no power of attorney; Legal and Compliance veto on legality only | slide 12, Q45, Q46 |
