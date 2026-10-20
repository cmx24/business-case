# REGRESSION.md - capabilities the CFO Chief of Staff deck must keep (checked by tools/deck_capability_check.py)

Grammar (same as Turbo's PITWALL): one capability per line, `capability: <words> => <pattern>[, <pattern>]`.
A capability PASSES when at least one pattern matches the deck's text (regex, case-insensitive, matched per slide; tables included).
Two extra line kinds for a document: `forbid: <words> => <pattern>` passes when NO slide matches (the package --guard analogue),
and `check: <words> => <key>[=<value>]` runs a structural check (`max_slides=N`, `notes`, `validates`, `no_raster_charts`).

Sources of the lines: section A is derived from the case brief (the "old registry"); section B is the reviewer's list
(the "builders' list"); section C is the guard. Edit the words freely; keep the grammar.

## A. From the case brief (deck/Nu_case_scope.pdf)

- capability: D1 names the budget gap => \$35M
- capability: D1 states the $60M asks against the $25M cap => \$60M.{0,80}\$25M, \$25M.{0,80}\$60M
- capability: D1 describes in-housing a central planning function => in-hous, central (modeling|planning) hub, hub-and-spoke
- capability: D1 compares option A (flat haircut) => equal percentage, pro-rata, flat haircut
- capability: D1 compares option B (core only) => core markets only, core markets first
- capability: D1 compares option C (staged) => staged (funding|release), business[- ]case readiness
- capability: D1 states a capacity or capital allocation rule (tranches or stages) => capacity allocation, allocat.{0,40}(rule|criteri|framework), tranche, stage [1-3] \(
- capability: D2 handles pushback from a senior GM => pushback, push(es|ed)? back
- capability: D2 gains buy-in without line authority => without (direct )?line authority, without authority, line authority
- capability: D2 structures a pre-Board alignment routine => pre-alignment, pre-brief, board (book|portal|calendar|rhythm|sprint)
- capability: D2 covers statutory governance across legal entities => statutory, legal entit
- capability: D2 addresses the two opposing MTeam views => mteam, opposing views, standoff, two executives, both executives
- capability: D3 defines milestone gates for the HQ project => gate [1-4], decision gates, milestone gate
- capability: D3 states statutory officer implications => statutory (officer|director), director (liability|insurance), d&o
- capability: D3 shows cross-functional leadership => cross-functional, cross-vertical, legal, tax
- capability: D3 covers regulatory and campus infrastructure terms => campus, lease, grant, incentive
- capability: D4 shares leadership style and operating philosophy => leadership style, operating philosophy, how I lead
- capability: D4 balances strategic influence with hands-on FP&A or governance rigor => technical rigor, fp&a, hands-on
- capability: assumptions are listed and separated from facts (general instruction) => key assumptions, assumptions? (list|ledger|slide), what (i|we) (am|are) assuming, assumption:
- capability: public data is cited with a source (general instruction) => 20-f, 6-k, earnings release, filings, investor relations
- capability: the deck fits the 30-minute format (presentation, not memo) => check: max_slides=22

## B. Reviewer's list (what a CFO panel looks for)

- capability: an explicit ask with a decision requested => authorize, approve, the ask, decision requested
- capability: the one number that matters ($25M) appears in the deck => \$25M
- capability: risks and reversal triggers are stated => reversal trigger, kill criteri, what would make (me|us) (stop|change), risk register, top (three|3|five|5) risks
- capability: decision rights are mapped (RACI, RAPID or equivalent) => \braci\b, \brapid\b, decision rights, who decides
- capability: the planning function has its own KPIs => (kpi|metric)s?.{0,200}(planning|finance) (function|hub), (planning|finance) (function|hub).{0,200}(kpi|metric), forecast accuracy.{0,60}(target|kpi|owner)
- capability: a capital lens exists next to the opex cap => return on allocated capital, cost of equity, capital headroom, \broac\b
- capability: Nu's latest cost to serve is used (Q2'26: $1.0) => \$1\.0\b.{0,60}cost.to.serve, cost.to.serve.{0,60}\$1\.0\b
- capability: the efficiency-ratio guide (~20% FY26) frames the cap => ~?20%.{0,40}efficiency, efficiency.{0,60}20%
- capability: the CFO's lens is addressed (Livingston, test-and-learn) => livingston, capital one, test-and-learn
- capability: the US launch and its spend envelope are acknowledged => lead bank, 100 bps, nubank\s*n\.a\.
- capability: Investor Day sits in the Board calendar => investor day
- capability: first-person voice in Deliverable 4 => \bI (lead|would|will|believe|balance|pair)\b, my (leadership|style|practice)
- capability: a personal example backs the leadership claims => when I was at, in my (previous|last|prior) role, \bI (led|built|negotiated|ran|closed|cut|took over|turned)\b
- capability: every core slide has speaker notes => check: notes
- capability: the file validates against the OOXML schema => check: validates
- capability: charts are native, not raster images => check: no_raster_charts

## C. Guard (forbidden content)

- forbid: banned word "seamless" => seamless
- forbid: borrowed brief phrases ("connective tissue", "strategic engine") => connective tissue, strategic engine
- forbid: "proves" as a verb for a third-party statistic => \b(data|research|study|studies) proves?\b
- forbid: stale cost to serve ($0.80 is Q2'25) => \$0\.80?\b.{0,40}cost|cost.to.serve.{0,40}\$0\.80?\b
- forbid: customer count rounded above the reported figure => 140M\+
- forbid: ROE attributed to a cohort => cohorts?.{0,80}return on equity, return on equity.{0,80}cohort
- forbid: an unverifiable AI self-audit verdict => independent audit verdict
- forbid: cap breach stated as margin neutral => margin neutral
