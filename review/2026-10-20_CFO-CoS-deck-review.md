# Review: "CFO Chief of Staff – Corporate Governance Strategy" deck

**Reviewed against:** Business Case – Chief of Staff to the CFO & Global Corporate Governance (Nu scope PDF, reporting to Rob Livingston, 30 min presentation + 30 min Q&A)
**Deck reviewed:** `CFO_Chief_of_Staff__Corporate_Governance_Strategy.pptx` (20 slides: 15 core, 1 divider, 3 appendix, 1 close)
**Method:** Nubank Business Case Architect skill (quality gates, red team, Nu fact base as of 20 Oct 2026). Every slide rendered and inspected; all arithmetic recomputed in code; file validated against the OOXML schema.
**Fact base:** Nu Q2'26 (quarter ended 30 Jun 2026). Q3'26 results were not yet published on the review date.

---

## 1. Verdict

**Evidence Sufficient with Qualifications. Not yet ready for a CFO panel. Fixable in one working day.**

The skeleton is right: the deck answers all four deliverables in order, leads with a recommendation, uses a staged-funding logic that genuinely fits Nu's test-and-learn DNA, and ends with an explicit ask. That is better than most candidate decks.

Three things would get it marked down by this audience:

1. **The load-bearing numbers are unsourced, stale, or self-contradicting.** Cost-to-serve is quoted at ~$0.80 (Q2'26 reported $1.0). "$15.6M saving", "$36.4M net value", "$61.2M net value" appear with no derivation anywhere. The $25M cap is breached in the upside case ($40M) while the slide says "100% margin neutral". The brief explicitly asks for assumptions to be separated from facts and listed; the deck has no assumptions list.
2. **The Board topic is "international capital allocation"; the deck answers with an opex split.** A 70/30 split of a $15M quarterly opex tranche is not capital allocation at a bank with $39B of credit, a new Mexican bank, a US de novo charter to capitalize by ~Jan 2027, and a $1B buyback. The CFO will notice within 30 seconds. Either state the assumption openly or add a capital lens.
3. **Deliverable 4 ("briefly share your leadership style") gets 3 of 15 core slides and contains zero personal evidence.** It is a literature review (Edmondson, Kim & Mauborgne, Ariely, Kotter, Bain). The brief says "beyond your professional experience, we encourage you to use public data". The panel wants to hear what you have actually done.

Plus one mechanical blocker: **the .pptx file is malformed** (invalid element order in the embedded-font list). LibreOffice refuses to open it; PowerPoint will likely show a "repair" prompt. A repaired copy with the layout defects also fixed is in `deck/`.

---

## 2. Scorecard against the four deliverables

| Deliverable (from the scope) | Coverage | Grade | Main gap |
|---|---|---|---|
| 1. Corporate planning framework, capacity allocation, the $35M gap | Slides 4–7 | B | Good logic (business-case readiness + staged release). Cap breached in upside; "quarterly" gates do not fit the day-count; no KPIs for the planning function itself; no RACI/decision rights. |
| 2. Stakeholder alignment, GM pushback, pre-Board routine, statutory governance | Slides 8–9, 13, App. A | B− | Pre-Board routine is concrete and good. MTeam dispute is assumed to be "core vs new bets" without saying so. Capital allocation conflated with opex. Statutory governance (explicitly in Pillar 2) is buried in Appendix A and its threshold logic is wrong for a regulated bank. |
| 3. HQ/campus project: gates, statutory officers, cross-functional leadership | Slide 10 (+ row on 12) | C+ | Four gates are sensible. No jurisdiction options, no tax-residence/substance/PE risk statement, no regulator fit-and-proper for officers, no risk register, no "how I lead Legal/Tax/People as an interim without authority". OECD rule is mis-described; $15.6M is unsupported. |
| 4. Leadership style and operating philosophy ("briefly") | Slides 11–13 | C | Three slides, all third-party citations, "we/our team" voice, no career examples, no "what I would not touch". |
| Assumptions clearly separated from facts and listed (general instruction) | — | F | Not present. Appendix B is a benchmark ledger, not an assumptions list. |
| Scenario/risk view | Slide 14 | C | Three scenarios but no risk register, no reversal triggers, no kill criteria, and the value numbers have no model behind them. |

---

## 3. Ranked list of issues (fix in this order)

### Must fix before sending (24h before the panel)

**1. File integrity.** `ppt/presentation.xml` lists the DM Sans embedded font as `<bold>` before `<regular>`; the schema requires regular → bold → italic → boldItalic. The package also uses non-standard namespace prefixes in `[Content_Types].xml`. Result: LibreOffice cannot load it ("source file could not be loaded"); PowerPoint will probably repair silently or prompt. Fixed in `deck/CFO_CoS_deck_repaired.pptx` (validates clean). If you re-export from your HTML/CSS tool, re-validate.

**2. Add the assumptions list the brief asks for.** One slide in the core deck (after the executive summary) or a full page in the appendix with a FACT / ASSUMPTION label on every number that drives a decision. Draft in Section 6 below.

**3. Refresh and correct Nu facts.**

| Deck says | Latest public figure (Q2'26 release unless noted) | Action |
|---|---|---|
| "~$0.80 monthly cost-to-serve" (slides 3, 18) | **$1.0** in Q2'26 (Q2'25 was $0.8). | Update. Quoting the year-old figure to the CFO who just reported the new one is a credibility killer. |
| "140M+ users" (slides 3, 18) | **138.9M** customers at Q2'26; Q3'26 not yet reported. | Say "~139M (Q2'26)". Do not round above the reported number. |
| "19.5% efficiency ratio" | 19.5% Q2'26 (17.6% Q1'26; ~20% expected for FY26 per the Q2 call) | Correct, but note management guided ~20% FY26 and attributed the Q2 rise partly to international expansion. That is your real cap argument. |
| "mature LatAm cohorts … 32%+ return on equity" (slide 8) | Group ROE 33% Q2'26. ROE is not disclosed by cohort. | A cohort has ARPAC, not ROE. Say "group ROE 33%" and "mature-cohort ARPAC" separately, with the source slide of the earnings deck. |
| "$27+ mature monthly revenue per active customer" | Nu has disclosed mature-cohort ARPAC in earnings materials; verify the exact figure and quarter before quoting. | Cite the deck page. Group ARPAC $17.1 (Q2'26) is the comparator. |
| "Core Markets (Brazil, Mexico, Colombia) vs New Bets (US, Global & Verticals)" | US launched 10 Sep 2026 via partner bank (Lead Bank) with account, card and Nu Global live; OCC conditional approval for Nubank N.A. (27 Jan 2026), capitalization deadline ~Jan 2027; management guided US spend at ≤ ~100 bps of efficiency ratio. Mexico received bank authorization 10 Jul 2026. | Acknowledge these. A "$4.5M Q1 pilot runway for New Bets incl. US" is not credible next to a ~100 bps envelope. Label the case numbers as hypothetical and tie your cap logic to the efficiency-ratio guide instead. |
| CFO not named | Rob Livingston, CFO since 13 Jul 2026 (ex-Visa, 18 yrs Capital One). | Address the CFO's lens explicitly: test-and-learn, cohort discipline, investor framing. The staged-funding idea maps directly onto Capital One's playbook. Say so once. |
| Board calendar | Investor Day scheduled 8 Dec 2026 (New York). | The "4-week Board" sits right before Investor Day. Your Board rhythm slide should say what the Board must approve before the Investor Day narrative is locked. |

**4. Fix the cap breach.** Slide 7 and 14: $15M (Q1) + $10M (Q2) = $25M, then "up to $15M in H2" = $40M total. Slide 14 calls the upside "100% margin neutral" with no basis. Options: (a) define Stage 3 as efficiency-ratio-neutral reinvestment, sized by formula (e.g., 50% of H1 pre-tax profit above plan, capped at $15M) and pre-agreed with the CFO and IR so it is inside guidance; or (b) drop Stage 3 and keep $25M as the hard cap with reallocation inside it. Pick one and make slides 2, 6, 7, 8, 14 consistent.

**5. Reframe materiality on slide 3.** Against Q2'26 annualized managerial revenue (~$23.5B) and the 19.5% efficiency ratio (implied opex ~$4.6B/yr):

| Amount | bps of annualized revenue | % of implied annual opex |
|---|---|---|
| $25M | ~11 bps | ~0.5% |
| $35M | ~15 bps | ~0.8% |
| $60M | ~26 bps | ~1.3% |

"$35M in unvetted overhead directly erodes our margin edge" will be challenged. The honest and stronger framing: the stakes are the operating model and the precedent (first centralized cycle, Board and IR credibility, keeping FY26 at ~20%), not the $35M itself.

**6. Separate capital allocation from the opex cap.** The MTeam dispute is "capital prioritization"; the Board topic is "international capital allocation across expanding legal entities". Add one slide (or a clearly labeled assumption): what capital decisions are actually on the Board agenda (equity for Nu Mexico as a bank, US charter capitalization, Colombia, buyback pace), the criterion to rank them (return on allocated capital vs. cost of equity by market; capital headroom by entity under the downside), and who decides what (CFO recommends, CEO decides, Board approves). Then say: "the $25M growth-opex envelope is the operating expression of that priority order".

**7. Cut the unverifiable precision.** The deck leans on 12+ third-party statistics with suspicious exactness: Hackett "35% / 29% / 56→40 days", Gartner "34%", Gartner "20–30% padding / 40% survival", McKinsey "1,600 companies / +30% TSR", Edmondson "−47%", HBR conflict "−50%", Bain RAPID "−40%", NACD "70%→80%", Site Selection "30% clawbacks", PwC "$150–250k". A panel with ex-consultants will probe one of them, and if it does not hold, the whole evidence base is discounted. (The McKinsey finding, for example, is about multi-year reallocation across business units, 1990–2005, not "quarterly gates".) Keep at most four sources you can defend verbatim, label the rest "directional", and never put a third-party statistic in a headline. Lead with Nu's own numbers and your own experience.

**8. Rebuild Deliverable 4 as one slide, first person.** Three principles, each with a real example (situation, your role, decision, result, what it taught you about leading without authority), plus "what I would deliberately not touch in the first 90 days". Move the operating-rhythm table (slide 12) into Deliverable 1 where it belongs; fold the GM-pushback playbook (slide 13) into Deliverable 2 as the scope asks.

**9. Fix the timeline arithmetic.**
- Slide 10 uses project days (Gate 1 Day 30 … Gate 4 Day 120); slide 15 uses tenure days. If the project launches on tenure Day 60, Gate 3 is Day 150 and Gate 4 is Day 180, not 135/150. Pick one clock and label it ("D+30 from charter").
- Slide 15 releases the "Q2" tranche on Day 125, 37 days after the "Q1" release on Day 88. Nobody can measure "90% of Q1 customer payback" in 37 days. Either use calendar quarters or rename the gate criteria to leading indicators you can observe in 6–8 weeks (activation rate, 30-day retention, CAC vs plan, unit cost vs plan) and say payback is a lagging confirmation.

**10. Add the three slides every C-level reader looks for and cannot find:**
- **The ask** with amount, scope, timing, KPIs and kill criteria (the one-liner on slide 2 is good; expand it into the closing slide instead of "Thank you").
- **Risks and reversal triggers** (GM gaming of gates, sandbagged cases, Board pre-read slips, HQ political/tax risk, key-person dependency) with the observable metric that would make you stop or change course.
- **Decision rights (RACI)** for the budget cycle and for the HQ project: group CFO, local CFOs (Mexico and Colombia already have local finance leadership), GMs, CEO, Board committees.

### Should fix

**11. Promote statutory governance out of Appendix A and fix its logic.** Pillar 2 explicitly names "statutory governance across expanding legal entities". The current rule (no local entity under $5M revenue or 15 staff) is wrong for a regulated financial business: whether Nu needs a local entity is driven first by licensing and regulatory requirements (you cannot take deposits or lend cross-border in most markets), then by tax substance, and only then by scale. Restate as: licensing need → tax/substance → scale thresholds → officer and signing controls. Keep the dual-signature and D&O points; add regulator fit-and-proper approval for statutory officers and who serves as director (and whether the CoS should ever be one).

**12. Make the HQ slide a real project page.** Name the decision options (which jurisdiction(s) are in scope and why), the hard constraints (regulatory permissions, tax residence of a Cayman holding company with Brazilian operations, Brazil's domestic minimum top-up tax in force since 2025, permanent-establishment risk from early signing), the gates with owners and exit criteria, the risk register (top 5), and your interim-leadership model for Legal, Tax, People and Workplace (weekly pod, decision log, 48h escalation is a good start; add who the escalation goes to). Correct the OECD description: the substance-based income exclusion removes 9.4% of eligible payroll and 7.4% of tangible-asset carrying value (2026 transitional rates) from the income subject to top-up tax; it is not a "backing" requirement. Either derive the "$15.6M" or delete it.

**13. Replace the scenario table (slide 14) with something a CFO can check.** Rows 4 and 5 ("HQ 5-year net value $6.2M / $15.6M / $21.4M", "3-year value $16.8M / $36.4M / $61.2M") have no model. Either put a one-page model in the appendix (inputs, formula, outputs) or remove the value rows and keep the spend rows, which are arithmetically sound.

**14. Language.** Remove "seamless" (slides 9 title, 15, 20), "connective tissue" and "strategic engine" (quoted from the brief; sound borrowed), "catalyst", "bulletproof", and every "proves". Switch Deliverable 4 and the ask to the first person. Make every title an assertion with a number where possible (proposed titles in Section 4).

**15. Appendix C (AI disclosure) and the "Independent Audit Verdict".** Decide this deliberately. As drafted, the table reads as "Gemini did the research, the model, the operating design and the writing; a human polished", and slide 18 ends with an AI-generated "Evidence Sufficient" self-grade. Recommended: one honest line in the sources footer ("Research assisted by AI tools; all analysis, numbers and recommendations are mine and were checked by me") and delete the four-row table and the self-audit line. If you keep the full disclosure, drop the model names and the "Human Executive Review" column.

### Polish

**16. Visual defects found in the render** (fixed in the repaired copy unless marked):
- Slide 3: "$60M Business Unit Requests vs. $25M Cost Target" overflows the dark stat card on both sides.
- Slide 7: Stage 2 bullet runs into Stage 3 ("targets." overlaps the next heading).
- Slide 19: the governance note is cut off at the bottom of the slide.
- Slides 2, 11, 13, 15: one card in each row sits lower/higher and is a different height from its siblings.
- Slides 4, 7, 8, 9, 10, 11, 15: ~1 inch of dead space between title and content while the footer is cramped (not fixed; design choice).
- Slide 6: the "waterfall" is a raster image, not a native chart, and is not a waterfall (each reduction is drawn against a full-width track, so −$8M looks like 13% of something). Rebuild as a native bridge chart: $60M → −$8M → −$12M → −$15M → $25M, with what happens to each piece labeled (not fixed).
- No slide numbers (the "01 //" labels are offset by one from the real slide number). Add page numbers for Q&A navigation.
- No speaker notes. Add, per slide, the oral takeaway and the hardest question it invites.
- Title slide has no candidate name, date, or "prepared for" line.
- Fonts: DM Sans and Space Grotesk are embedded (with the broken entry). If you cannot guarantee the embed, switch body text to Arial or Calibri so the deck renders identically on the panel's machines.

---

## 4. Slide-by-slide edits

Format: current → proposed. "Cut" means remove; "Keep" means fine as is.

| # | Current title | Proposed title (assertion) | Edits |
|---|---|---|---|
| 1 | Connecting Financial Discipline, Board Rhythm, and Ways of Working as Nu Scales | Chief of Staff to the CFO: a planning engine, a Board rhythm, and an HQ charter for the next 90 days | Add name, date, "Prepared for the Nu panel; all Nu figures from Q2'26 public filings; assumptions listed on slide 3". Keep the three-pillar strip. |
| 2 | Our Office Acts as the Connective Tissue Between Ambition and Discipline | Fund $25M now through staged business-case gates, lock the Board plan 16 days out, and charter the HQ project this week | Good ask line; keep it. Replace the metaphor title. Add the one number per deliverable (e.g., "$25M across 2 gates", "T-16 executive lock, T-7 book", "4 gates, 120 days", "3 rules, 3 examples"). |
| NEW | — | What I am assuming, and what I would verify in week one | The brief demands it. Use the list in Section 6. |
| 3 | Scaling Past 140M Customers Requires Shared Rules Across All Regions | The $35M gap is small against a ~$23B revenue base; the precedent and the FY26 ~20% efficiency guide are what matter | Replace ~$0.80 with $1.0 (Q2'26); 140M+ with ~139M; drop the PwC and Hackett cards (or label directional). Add the materiality table. |
| 4 | Use Nu's "Test, Learn, and Scale" Credit Method to Allocate Budgets | Budgets should be underwritten like credit lines: small, data-gated, and raised fast when the cohort proves out | Keep the analogy, it is the best idea in the deck. Cut the "+30% TSR" McKinsey card or move to footnote as directional. Add one line tying it to the CFO's Capital One/Visa background (one sentence, no flattery). |
| 5 | Rigorous Business Case Readiness and Staged Funding Beat Flat Haircuts | Three ways to close $35M: staged readiness wins on return and GM ownership, flat cuts win only on speed | Arithmetic checks (57.8 / 63.8 / 91.8). Add a weight-sensitivity line: "winner unchanged if any weight moves ±10 pp" (verify in code). State that the scores are your judgment (ASSUMPTION). Rename "Core Markets Only" to "Core markets first" since the deck later funds new bets anyway. |
| 6 | Sequence the $60M Pipeline by Business Case Readiness and Shared Tech | $60M of asks become $25M of funded spend: $8M delivered on the shared platform, $12M sent back for a 60-day case, $15M conditional on H1 | Rebuild as a native bridge chart. Label the status of each piece (delivered at ~$0 marginal cost / deferred / conditional). Add "ASSUMPTION: the $8M platform-reuse estimate is mine". |
| 7 | Combine a Central Modeling Hub With Embedded Business Unit Partners | Hub-and-spoke: one model and one calendar centrally, finance partners inside each market and vertical | Fix the Stage 3 cap breach (Section 3, item 4). Replace "Q1/Q2" with gate dates on a calendar. Replace "90% of Q1 customer payback" with observable leading indicators. Add the planning function's own KPIs: cycle time, forecast accuracy, % of spend with an approved case, variance to plan. Add a mini-RACI. |
| 8 | Settle the Core vs. New Vertical Standoff Using Horizon-Specific Metrics | Core markets and new bets are judged on different metrics, so both executives can win the Q1 lock | State the assumption that the dispute is about the growth-opex envelope. Fix "32%+ ROE" (group ROE 33%). Add the capital lens or point to the new capital slide. Say who breaks a tie (CFO recommends, CEO decides) and that the options memo goes to both executives before the Day-16 lock. |
| 9 | Our Governance Team Drives a 28-Day Sprint to Deliver One Unified Story | The Board book ships T-7 with written Q&A, after a T-16 executive lock and a T-10 input freeze | Remove "seamless". Add owners (Corporate Secretary / General Counsel, committee chairs), the Investor Day (8 Dec 2026) dependency, and what the Board must decide vs. note. Move the "statutory entity rule" to the promoted governance slide. |
| 10 | Govern the New Headquarters Through Four Risk-Clearing Decision Gates | HQ charter: four gates in 120 days, each with an owner, an exit test, and a kill condition | Add jurisdiction options and constraints, risk register, interim-leadership model. Fix the OECD description. Derive or delete $15.6M. Make the gate days consistent with slide 15. |
| NEW | — | Statutory governance for new entities: license need first, substance second, scale third | Promote from Appendix A with corrected logic; add officer fit-and-proper, D&O, dual signature, and who may be a director. |
| 11–13 | Three slides of behavioral science | How I lead without line authority: three rules, three examples | One slide, first person, with career examples and "what I would not touch". Cut the citations to a footnote. |
| 12 (table) | How Our Team Integrates Into Daily Business Rhythms | The operating cadence: weekly squad clinics, bi-weekly GM 1:1s, weekly HQ pod, quarterly gate review | Keep the table; move it into Deliverable 1; delete the "Measurable Impact (Backed by Research)" column and replace with "Owner / Output". |
| 13 (playbook) | Four-Step Playbook for Turning Senior GM Pushback Into Shared Alignment | When a senior GM pushes back: validate the goal privately, move to criteria, offer a conditional yes, co-present the result | Keep the four steps (they are good). Move into Deliverable 2. Replace citations with one real example of you doing it. |
| 14 | How Our Staged Governance Model Performs Under Every Market Scenario | If H1 misses, Q2 funding holds at $15M automatically; if it beats, up to $15M more is released inside the efficiency guide | Keep rows 1–3 (sound). Delete rows 4–5 or back them with an appendix model. Add reversal triggers per scenario. |
| NEW | — | Risks and what would make me change course | Top 5 risks, owner, early-warning metric, threshold, response. |
| 15 | Our 90-Day Roadmap Across Planning, Board Rhythm, and the HQ Project | Days 61–150: Board locked by Day 88, hub in place by Day 95, HQ Gate 2 cleared by Day 120 | Fix gate-day consistency and the Q2 tranche timing. Add a column "what I would not touch yet". |
| 16 | Appendices divider | Keep | — |
| 17 | Appendix A | Keep as backup once promoted | Correct the threshold logic. |
| 18 | Appendix B ledger | Sources and assumptions ledger | Add columns: classification (FACT/ASSUMPTION/ESTIMATE), period, confidence. Delete the "Independent Audit Verdict" line. |
| 19 | Appendix C AI disclosure | Cut, or reduce to one line in the footer | See item 15. |
| 20 | Ready to Launch … Today | The ask: approve $25M in two gates, the T-16/T-7 Board calendar, and the HQ charter; I report back at Day 90 | Replace "Thank you" with the decision requested, KPIs and kill criteria. |

---

## 5. Proposed storyline for 30 minutes (14 core slides, ~2 min each, plus backup)

1. Title (name, date, fact base)
2. The answer and the ask (three decisions, one number each)
3. Assumptions and what I would verify in week one
4. Where Nu stands and why the cap matters (Q2'26 scorecard, materiality, FY26 ~20% guide)
5. Deliverable 1: underwrite budgets like credit lines (the idea)
6. Deliverable 1: three options compared, with weight sensitivity
7. Deliverable 1: the $60M bridge to $25M and the two gates (native chart)
8. Deliverable 1: hub-and-spoke operating model, RACI, cadence, the function's KPIs
9. Deliverable 2: capital allocation vs. opex envelope, and how the MTeam dispute gets decided
10. Deliverable 2: the 28-day Board rhythm with owners and the Investor Day dependency
11. Deliverable 2: statutory governance rules for new entities and officers
12. Deliverable 3: HQ charter, four gates, risk register, interim-leadership model
13. Deliverable 4: how I lead without line authority (three rules, three examples)
14. Risks, reversal triggers, and the 90-day plan on one page

Backup (not presented): sources and assumptions ledger; scenario and value model; subsidiary matrix; the GM-pushback playbook; the operating-rhythm table; the behavioral-science references.

---

## 6. Draft assumptions list (for the new slide the brief requires)

Label every item on the slide. "FACT" items carry a source; "ASSUMPTION" items are yours.

| # | Statement | Class | Source / basis |
|---|---|---|---|
| A1 | Nu Q2'26: ~139M customers, ARPAC $17.1, cost to serve $1.0, efficiency ratio 19.5%, ROE 33%, net income $1.06B | FACT | Q2'26 earnings release / 6-K |
| A2 | Management expects FY26 efficiency ratio ~20% and guided US spend at ≤ ~100 bps of the efficiency ratio | FACT (verify wording) | Q2'26 earnings call |
| A3 | The $60M / $25M figures in the case are hypothetical and far below Nu's actual international investment (e.g., Mexico $4.2B through 2030) | ASSUMPTION (case framing) | Case brief; Nu Feb 2026 announcement |
| A4 | The two MTeam members' dispute is about the growth-opex envelope split between core markets and new bets, not about regulatory capital | ASSUMPTION | Case brief (not specified) |
| A5 | $8M of the $60M asks can be delivered on the shared global platform at ~zero marginal cost | ASSUMPTION | Candidate estimate; to verify with Engineering in week one |
| A6 | $12M of asks lack a business case today and can be prepared within 60 days | ASSUMPTION | Candidate estimate |
| A7 | Stage 3 (up to $15M) is funded only from H1 pre-tax profit above plan and is efficiency-ratio neutral | ASSUMPTION / RULE | To be agreed with CFO and IR |
| A8 | The quarterly Board meeting is in 4 weeks and Investor Day is 8 Dec 2026; the Board must approve the international capital plan before the Investor Day narrative is locked | FACT (dates) + ASSUMPTION (sequence) | Case brief; Nu IR |
| A9 | The HQ jurisdiction is [X]; Nu Holdings' group revenue is above the €750M Pillar Two threshold | ASSUMPTION / FACT | Case brief; Q2'26 revenue |
| A10 | The 2026 transitional substance-based income exclusion is 9.4% of eligible payroll and 7.4% of tangible-asset carrying value | FACT | OECD GloBE Model Rules Art. 9.2 |
| A11 | Option scores on the prioritization matrix are the candidate's judgment | ASSUMPTION | Slide 6 |
| A12 | All third-party benchmarks (Hackett, Gartner, McKinsey, HBR, Bain, NACD, PwC, Deloitte) are directional and not load-bearing | CONTEXT | Appendix ledger |

Week-one verification list: the actual MTeam positions; the Board agenda and committee calendar; the current planning calendar and who owns the model; local-CFO decision rights in Mexico and Colombia; the HQ project's current state (sites, counterparties, any signed LOIs); the planning function's current cycle time and forecast accuracy.

---

## 7. Red team: questions this panel will ask, and the answer the deck should already contain

1. **"Cost to serve is $1.0, not $0.80. Did you read the Q2 release?"** Fix the number; say why it rose (real estate, marketing timing, international expansion) and why that supports a cap on unvetted overhead.
2. **"Your upside case spends $40M against a $25M cap. How is that margin neutral?"** Have the Stage 3 formula and the efficiency-ratio math on the slide.
3. **"The Board topic is capital allocation across entities. Where is the capital?"** Have the capital-lens slide: return on allocated capital vs. cost of equity by market, capital headroom by entity, the US capitalization deadline, buyback vs growth.
4. **"Why should a GM trust a gate they cannot pass in 37 days?"** Switch to leading indicators and calendar quarters.
5. **"Where did $15.6M come from?"** Derive it or delete it.
6. **"You cite twelve consultancies. Which of those studies have you read?"** Keep four you can discuss.
7. **"Tell me about a time you got a senior leader to accept less money."** This is the Deliverable 4 example the deck currently lacks.
8. **"What would you not touch in your first 90 days?"** Add the line (e.g., local-CFO decision rights in Mexico and Colombia; the credit model governance; the IR calendar).
9. **"Who signs for the HQ entity, and can you be a statutory director?"** Have the officer and fit-and-proper answer ready; the honest answer is usually "not me".
10. **"Did AI write this?"** One honest sentence, then back to the content. The current three-model table invites a longer conversation than you want.

---

## 8. What was fixed in the repaired copy (`deck/CFO_CoS_deck_repaired.pptx`)

Mechanical only; no wording changed, text content verified identical to the original.

- Embedded-font list element order corrected (file validates against the OOXML schema; LibreOffice now opens it).
- Slide 3: stat-card text boxes narrowed to sit inside the card.
- Slide 7: right-card headings aligned with the left card; Stage 2 and Stage 3 bullets given room; both cards extended so nothing overlaps.
- Slides 2, 11, 13, 15: cards and their text boxes aligned to a common top and height.
- Slide 19: table rows tightened and the governance note moved back inside the slide.

Not fixed (content decisions for the author): dead space under titles, the raster waterfall, missing slide numbers and speaker notes, fonts.

---

## 9. What I checked and how

- Rendered all 20 slides (LibreOffice, after repair) and inspected each.
- Recomputed: weighted scores (57.75 / 63.75 / 91.75 → shown as 57.8 / 63.8 / 91.8, correct); bridge $60 − 8 − 12 − 15 = $25M (correct); stage math $10.5 + $4.5 = $15M, $16.5 + $8.5 = $25M, H2 $5 + $10 = $15M (internally consistent); total upside spend $40M vs $25M cap (breach); gate-day mapping (project Day 90/120 ≠ tenure Day 135/150).
- Checked the OECD 2026 transitional substance carve-out rates (9.4% payroll, 7.4% tangible assets: correct) and the description of how they apply (incorrect in the deck).
- Checked Nu figures against the Q2'26 fact base in the business-case skill (snapshot 20 Oct 2026). Items I could not verify from that base are marked "verify" above rather than asserted.
- Searched for a "PitWall" skill as requested: none exists in this workspace or the plugin catalog (the only match is a Formula 1 telemetry plugin), so the review uses the Nubank Business Case Architect skill only.
