# TEST_06: deck v6 (two AI agents in the operating model) and the v6 talk track (2026-10-20)

Request (pit wall): "find ways to insert AI/LLM tooling in how myself and my team will look to implement and run governance
... think of 2 real life day to day examples ... focus on use cases that have some data or reference behind why they are
best practices."

Files under test:
- `deck/CFO_CoS_deck_v6.pptx`, built by `tools/build_v6.py`. sha256 d6dd67539be6478e50b0aa94d3dade15bb7e22f3e20dc01d1922831f0adf2c5d.
- `review/2026-10-20_v6_talk-track-and-QA.md`, renamed from the v5 file and edited.

## The two use cases and their evidence

| Use case | Day to day | Evidence (how checked) |
|---|---|---|
| PMO agent (the pit wall's example) | Reads status where the work happens for the HQ gates, the releases and the Board countdown; updates the gate tracker and decision log; flags dates at risk; drafts the pod note. Replaces status templates | Gartner press release, 20 Mar 2019: 80% of today's project-management tasks eliminated by 2030 as AI takes over data collection, tracking and reporting. This is a forecast, labeled as one. Found via search; the gartner.com page could not be fetched from here (proxy 403). PMI 2023 Annual Global Survey: reporting is project managers' top AI use (34%). This is secondary: a press report of the survey, flagged as such in the numbers sheet |
| Review agent | Checks each GM case against the template and the four indicators and ties figures to the data platform before the clinic; drafts the Board book and directors' answers, with every figure linked to its source | Noy and Zhang, *Science*, Jul 2023 (453 professionals; writing about 40% faster, quality up about 18%). Dell'Acqua et al., HBS/BCG working paper, 2023 (758 consultants; inside the frontier 12.2% more tasks, about 25% faster, more than 40% higher quality; 19 percentage points less accurate outside it). Gartner press release, 11 Sep 2024 (58% of finance functions use AI; anomaly and error detection 39%). All found via search results quoting the primary pages; the PDFs could not be fetched from here |
| Nu context | | Q2'26 release: AI agents handle more than 60% of support conversations in Brazil (search result quoting the release). Nu newsroom, 26 Jun 2024: Hyperplane acquisition "to accelerate AI-first strategy" |

Guardrails on the slides:
- Nu-approved tools inside Nu's environment (new assumption A11: Security and Risk approve access first). The Board book
  holds non-public results.
- A named person signs every figure and answer.
- The PMO agent only reads: it never changes a date or a decision.
- A wrong number that reaches the CFO sends that step back to manual.

## Where it appears

**Slides:**
- Slide 7: a hub card bullet.
- Slide 10: a new band on the Board-book drafts and directors' answers. The cards were shortened to make room.
- Slide 11: the pod's decision log is kept by the PMO agent.
- Slide 12: rule 3, "AI does the first pass; I still do the last".
- Slide 14: "PMO agent live, status templates retired" by late November.
- Backup H: new slide (what each agent does, what it replaces, evidence, guardrails, measures).
- Backup A: A11.
- Backup D: a new ledger row.
- Divider: lists H.

**Talk track and Q&A:**
- Talk track: slides 7, 10, 11, 12 and 14.
- Q&A: Q10, Q30 and Q37 updated; Q39, Q40 and Q41 added (41 questions).

## Gates

- **OOXML validation:** PASSED.
- **Strict package check:** no structural problems.
- **Render:** all 23 slides converted. Slides 3, 7, 10, 11, 12, 14, 15, 16, 19 and 23 were inspected, with no overflow or
  overlap. Backup H's font was enlarged from 10 to 11pt after the first render showed spare room. The fit check predicted
  no overflow, and the title-collision guard raised no warnings.
- **Capability scan:** 47 pass, 1 fail (the personal example, B16). See `scan_v6.txt`.
  - The 30-minute line now counts presented slides before the backup divider, limit 14 (`max_presented`, new in
    `deck_capability_check.py`). That is stricter than the old total of 22: v4 fails it.
  - Caveat: v4 counts as 22 presented slides, not its actual 19, because v4's divider did not start with "Backup". It
    fails either way.
  - The total-slide memo guard was raised from 22 to 23 for backup H.
  - The checker self-test, extended with both outcomes of the new check, passes.
  - Two AI capability lines were added, and both pass.
- **Preservation:**
  - v5 to v6 (slides, notes and the talk track): 182 tokens, 0 missing.
  - v4 to v6: 1 missing, "29 January 2026", which is written "29 Jan 2026" (format only).
- **Talk track** (`tools/talk_timing.py`):
  - 3,300 words, about 25.4 minutes. To make room, five lines moved to Q&A answers (Q4, Q29, Q33) or are already on the
    slides.
  - Checkpoints: slide 8 at 10.9 (target 13), slide 11 at 17.1 (19.5), slide 13 at 22.8 (23).
  - No sentence over 30 words. New Q&A answers run 27-51 seconds.

## Still NOT proven

- That PowerPoint opens v6 (same packaging as v3-v5).
- The personal examples, the derailer and the Visa specifics (B16, B22).
- The dry run (B8).
- The PMI 34% figure in the original report.
- A11 (Nu's actual approved-tool policy).
