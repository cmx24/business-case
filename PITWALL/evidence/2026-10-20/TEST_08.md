# TEST_08: deck v8 (one way of deciding, used three times) and the v8 talk track (2026-10-20)

Request (pit wall): "incorporate and take a final pass at making the deck the best it can be. have Fable lead this and Opus executes."

**Fable led:**
- `review/2026-10-20_fable-reassessment.md`: verdicts; revisions 1-7 plus the optional slide 10 line; the talk-track changes.
- `review/2026-10-20_fable-final-pass.md`: ten punch-list items.

**Opus executed:**
- `tools/build_v8.py` produced `deck/CFO_CoS_deck_v8.pptx` (sha256 3bae6ed8ed42f9d408af7a7fb116f0a1be315f6a41feb6734d0dd3e3fe07152c).
- `review/2026-10-20_v8_talk-track-and-QA.md` (renamed from v7 and edited).

## Applied

**Reassessment:**
- Slide 2: thesis and through-line strip.
- Slide 5: five answers named on the bridge, with an honest "stop".
- Slide 7: options for the CEO, with the 70/30 as the recommendation.
- Slide 10: the shortlist line.
- Slide 11: decision gates (can we / what may we offer / who may bind Nu / do we commit), with gate 0 = the charter.
- Slide 12: the posture title ("experts own the judgment, I own the process, the CFO owns residual risk").
- Slide 14: the matrix split.
- Slide 15: the closing echo.

**Final pass, all ten items:**
- Slide 9: becomes the standing quarterly routine.
- Slide 2: tiles consistent with slides 7, 11, 12 and 14.
- Slides 15 and 25: "guide in force" instead of the FY26 guide at an FY27 date.
- Slide 4: the footer labels which base each calculation uses.
- Slide 15: "the Board on 17 Nov".
- Speaker notes: synced on six slides.
- Slide 12: risk card "decide on".
- Slide 6: decision split.
- Slide 13: footer drops the talk-only promise.

**Adapted:**
- Slide 13 title: Fable's "A few clear rules, published early, and nothing more: three I hold myself to" lost the brief's Deliverable 4 words. Used "Influence with hands-on rigor: a few clear rules, published early" instead.
- Slide 2 D2 tile: "at T-16" became "at the lock" (T-numbers stay on the Board slide).
- Orphan words on three titles: shortened to one line (slides 12, 13, 15).
- Closing: kept the A4 question instead of "Thank you", per R3.

## Gates

- OOXML validation: PASSED.
- Strict package check: no structural problems.
- Fit check: no overflow predicted (one slide 9 card shortened). No title-collision warnings.
- Render: all 26 slides converted; presented slides 1-15 and backup I inspected. The preview PDF matches the committed deck.
- Capability scan: 52 pass, 1 fail (the personal example, B16). See `scan_v8.txt`.
- Placeholders in the deck (slide text and notes): 0.
- Preservation v7 to v8 (deck plus talk track): 193 tokens, 0 missing.
- Talk track: 3,300 words, about 25.4 minutes; no sentence over 30 words.
  - Checkpoints: slide 7 at 9.3 (target 10), slide 10 at 15.2 (16), slide 13 at 20.5 (21), slide 15 at 24.1 (24.5).

## Still NOT proven

- That PowerPoint opens v8.
- The story slots, the derailer and the Visa specifics (B16, B22).
- The dry run (B8).
- A12 and A13.
