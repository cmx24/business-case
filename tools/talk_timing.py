import re, sys
p = sys.argv[1]; t = open(p).read()
sec = t.split("## 1. Talk track")[1].split("## 2. Timing")[0]
parts = re.split(r"^### (Slide \d+: .+)$", sec, flags=re.M)
rows, total, cum, long_s = [], 0, 0.0, []
for title, body in zip(parts[1::2], parts[2::2]):
    spoken = " ".join(l[1:].strip() for l in body.splitlines() if l.startswith(">"))
    spoken = re.sub(r"\[(Pause\.|Switch to slide \d+\.)\]", " ", spoken)
    ph = [int(n) for n in re.findall(r"\[your [^\]]*?about (\d+) words[^\]]*\]", spoken)]
    plain = re.sub(r"\[your [^\]]*\]", " ", spoken)
    n = len(plain.split()) + sum(ph)
    for s in re.split(r"(?<=[.?!])\s+", plain):
        if len(s.split()) > 30: long_s.append((title.split(":")[0], len(s.split()), s[:80]))
    m = n / 130; cum += m; total += n
    rows.append(f"| {title.split(': ',1)[0].replace('Slide ','')}. {title.split(': ',1)[1]} | {n}{' (incl. '+str(sum(ph))+' story)' if ph else ''} | {m:.1f} | {cum:.1f} |")
tbl = "| Slide | Words | Minutes | Cumulative |\n|---|---|---|---|\n" + "\n".join(rows) + f"\n| **Total** | **{total}** | **{total/130:.1f}** | |"
print(tbl); print("long sentences:", long_s)
if len(sys.argv) > 2:
    note = (f"\n\n**Total: {total} spoken words, {total/130:.1f} minutes at 130 words per minute.** The story slots are counted at their stated "
            "length; a longer story comes out of the buffer. That leaves about five minutes of the 30 for the walk-in, two question "
            "pauses and one or two interruptions. A comfortable pace in an interview is often slower than 130, so time yourself once "
            "(board B8) and cut from the 'If you are behind' list rather than speaking faster.")
    open(p, "w").write(t.replace("TIMING_TABLE", tbl + note))
