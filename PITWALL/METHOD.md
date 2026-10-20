# PitWall method, as applied to the CFO Chief of Staff case deck

The PitWall folder of the FC27-Editor-Turbo project lives outside git (`turbo_dev\PITWALL\` on Cassio's PC), so this note
reconstructs the method from what the repository itself says about it. Every rule below is cited to a file on branch
`origin/2.21/play` of `cmx24/FC27-Editor-Turbo` (commit 2456cd0 lineage, fetched 2026-10-20). Where the deck review
needs a rule the repo does not state, it is marked ADAPTED.

## 1. The folder and its artefacts (what PITWALL contains)

| Artefact | What it is in Turbo | Source | Deck equivalent in this repo |
|---|---|---|---|
| `PITWALL\REGRESSION.md` | The list of capabilities that must keep working, as lines `capability: <words> => <function>[, <function>]`; a scan checks each one mechanically and prints PASS/FAIL and one RESULT line, exit 1 on any fail. | `tools/nav_capability_check.py` docstring lines 5-14 and 265-291 | `PITWALL/REGRESSION.md`, checked by `tools/deck_capability_check.py` |
| `PITWALL\capabilities\<track>.md` | The builders' own capability lists per work track, same line grammar. | `tools/nav_capability_check.py` lines 12-13, 358 | `PITWALL/capabilities/` (one file per deliverable, if the deck is split across authors) |
| `PITWALL\board\ia-*.md` | The board: one row per work track (the `ia-*` tracks are the GUI information-architecture pages), each with its blockers; the scan reads its capability lines alongside the builders' lists. | `tools/nav_capability_check.py` lines 13, 274; `tools/club_fields_livecheck.py` line 26 ("blocker 1 of the board row"); `tools/profile_livecheck.py` line 2 ("brief 2.20-IA-pages, track ia-profile") | `PITWALL/board/pitwall-decisions.md` (decisions only the pit wall can take, with blockers per gate) |
| `PITWALL\briefs\<version>-G4.md` | The gate brief of a release; "in-game status is filled in at G4". | `CHANGELOG.md` line 18 (2.21.0) | `PITWALL/briefs/deck-v1-gates.md` |
| `evidence\<date>\TEST_<nn>.md` | Evidence of one launch: criterion written before, results after, never edited after the launch. | `tools/session_template.md` sections 0 and 4; `docs/HANDOVER.md` 00f | `PITWALL/evidence/2026-10-20/TEST_01.md` |

## 2. Roles

| Role | Rule | Source |
|---|---|---|
| **Pit wall** (Cassio) | Takes the timestamped calls that change scope or risk; code cites them as "Pit wall HH:MM". No launch without his chat "go". | `turbogui/src/core/callnames.h` line 148; `tools/turbo_launch.ps1` line 104; `tools/crest2d_livecheck.py` line 12; `tools/session_template.md` binding rules |
| **Coordinator** | The only one who drives the live system and writes state; agents never do. | `tools/GAME_STATE.md`; `docs/TACTICS_OVERHAUL_DECISION.md` section 10 |
| **Builders** | Write the code and publish their capability lines; hand the coordinator a scripted check of at most 2 minutes. | `tools/nav_capability_check.py` line 12; DECISION section 10 |
| **Tester** | Runs the live steps exactly as scripted, files the screenshot; his own careers are never touched. | `tools/crest2d_livecheck.py` lines 10-19; `CHANGELOG.md` line 32 |

ADAPTED for the deck: pit wall = Cassio (the candidate and owner of the deck); coordinator = the reviewer running the gates;
builders = whoever drafts slides; tester = the person who does the timed dry run in front of a stand-in panel.

## 3. Gates

1. **Pre-registered criterion.** The gate and its pass criterion are written in the evidence file before the run and never edited
   after it. No "BASELINE NOT FOUND" left in a criterion. (`tools/session_template.md` section 0)
2. **One pass/fail point per check, time-boxed, with a screenshot or an exact log line as the evidence.** (`session_template.md` section 2;
   brief Part C.1 rule 4)
3. **Branch table.** Every outcome of a gate has a pre-decided branch, so the implementer never stops to ask.
   (`docs/TACTICS_OVERHAUL_DECISION.md` intro and section 4)
4. **INCONCLUSIVE counts as FAIL** for the route it was meant to prove. (`CHANGELOG.md` line 10, gate G-T1)
5. **The one real risk is named, detected by the first gate, and escalated honestly.** (DECISION intro, R-G1-1)
6. **Ship only when working / ship only what is proven.** Failing items are fixed or removed, never shipped with a caveat.
   (brief Part C.1 rule 5; DECISION section 10; `session_template.md` binding rules)
7. **Get in, prove the point, get out.** Short runs; never keep the system running for probing. (`docs/HANDOVER.md` 00f, Cassio 20:38;
   Part C.1 rule 4)

## 4. Evidence and wording

1. Every claim carries a status: **Proven in game** or **Not proven** in each CHANGELOG release section; **UNPROVEN** items are listed
   first under Unreleased with the gate that failed. "Nothing here is shipped or claimed as a feature." (`CHANGELOG.md` lines 3-14, 34-45)
2. Knob states are **PROVEN / LOCATED / UNLOCATED**; the UI shows honest badges and a yellow "no proven in-game effect" line for
   anything not proven. (`docs/HANDOVER.md` 00f; `docs/GUI_AUDIT.md` line 36; `CHANGELOG.md` 2.1.0 "honest badges")
3. **The evidence wins** over any note that cites it; **where an audit downgrades a claim, the audit wins.**
   (`docs/TACTICS_OVERHAUL_BRIEF.md` line 19; DECISION section 9 on `T<n>_verify.md`)
4. The report says plainly what is **still NOT proven after this run**. (`session_template.md` section 4)
5. Report only what was actually run, versus what is offline-only or untested. (`docs/RESUME-PROMPT.md` STEP 1)
6. A decision document picks one option, records why the others were rejected, and lists residual risks each with a mitigation.
   (DECISION sections 2.1-2.2)

## 5. Review

1. Every track gets an **adversarial review plus a fix pass**; findings are classed **blocker / major** (and lesser), each closed or
   proven correct in the live system. (`docs/HANDOVER.md` 00 "Reviews"; 00d "review workflow: 3 lenses + skeptics")
2. A regression run of a candidate can find the defect that the next version fixes; the candidate is then never published.
   (`CHANGELOG.md` 2.17.1)
3. A handover is updated after every session and ends with a prompt for the next operator. (`docs/HANDOVER.md` header)

## 6. Safety rules carried over (ADAPTED)

Turbo's write-safety rules (opt-in, kill switch, read-back, undo, backup + md5 before and after) have one deck analogue:
**never edit the author's wording during a mechanical repair, and prove it.** The repaired deck in `deck/` was checked line-for-line
against the original (identical text set; only reading order changed where boxes were moved).

## 7. How the gates map onto this deck

| Gate | Question | Check | Evidence |
|---|---|---|---|
| G1 Scope | Does the deck cover every deliverable and instruction in the case brief? | `tools/deck_capability_check.py` against `PITWALL/REGRESSION.md` (symbol scan, no judgment) | `PITWALL/evidence/2026-10-20/TEST_01.md` |
| G2 Facts and arithmetic | Is every Nu figure current and every calculation reproducible? | Fact table against the Q2'26 release; arithmetic re-run in code | same file |
| G3 Decision logic | Is the recommendation internally consistent and does it answer the decision the panel will actually judge? | Cap consistency, capital vs opex, timeline clocks, option set | same file |
| G4 Panel readiness | Does the file open, render clean, fit 30 minutes, and survive the red-team questions? | OOXML validation, full render, dry run, Q&A bank | same file; the dry run is the Tester's step |

The branch tables for each gate are in `PITWALL/briefs/deck-v1-gates.md`.
