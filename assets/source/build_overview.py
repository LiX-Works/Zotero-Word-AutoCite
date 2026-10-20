"""Zotero Word AutoCite overview, before-and-after design.

Usage: py build_overview.py [docked|beforeafter] -> writes overview.svg/.html next to this file.
docked:      a dark AutoCite bar with the three inputs docked on top and a zoomed Word page under it.
beforeafter: the plan as Word comments (left) becomes live Zotero fields (right) through a dark band.
The default variant is beforeafter.
Body text is deliberately blurred; citations use bracketed numbers.
"""
import random
import sys
from pathlib import Path

VAR = sys.argv[1] if len(sys.argv) > 1 else "beforeafter"
if VAR not in {"docked", "beforeafter"}:
    raise SystemExit("variant must be 'docked' or 'beforeafter'")
SANS = "Arial,Helvetica,sans-serif"
SERIF = "'Times New Roman',Times,serif"
INK, SUB = "#1F2933", "#475467"
NAVY0, NAVY1 = "#1E355F", "#2F5B96"
FIELD, FIELD_BG = "#B4232F", "#FCE4E6"
AMBER, AMBER_T, AMBER_BG = "#C07A2C", "#8A5317", "#FDF3E7"
WORDBLUE = "#2B579A"
BAR = "#C3CAD6"

out = []
add = out.append


def text(x, y, s, size=22, weight=400, fill=INK, anchor="start", family=SANS):
    s = s.replace("&", "&amp;")
    add(f'<text x="{x}" y="{y}" font-family="{family}" font-size="{size}" font-weight="{weight}" '
        f'fill="{fill}" text-anchor="{anchor}">{s}</text>')


def words(x0, x1, y, seed, h=11):
    """A blurred line of text made of word-like segments between x0 and x1."""
    r = random.Random(seed)
    x = x0
    segs = []
    while x < x1 - 20:
        w = r.choice((34, 42, 50, 58, 66, 74, 88))
        w = min(w, x1 - x)
        segs.append(f'<rect x="{x}" y="{y - h / 2}" width="{w}" height="{h}" rx="3"/>')
        x += w + 12
    add(f'<g fill="{BAR}" filter="url(#blur)">' + "".join(segs) + "</g>")


def field(x, y, label, size=25):
    w = len(label) * size * 0.5 + 22
    add(f'<rect x="{x}" y="{y - 20}" width="{w:.0f}" height="34" rx="4" fill="{FIELD_BG}" stroke="{FIELD}" stroke-width="1.8"/>')
    text(x + w / 2, y + 6, label, size, 700, FIELD, "middle", SERIF)
    return w


def bibliography(x, y, w, n=5):
    text(x, y, "References", 26, 700, INK, "start", SERIF)
    add(f'<rect x="{x - 12}" y="{y + 16}" width="{w + 24}" height="{n * 34 + 16}" rx="6" fill="{FIELD_BG}" opacity=".7"/>')
    for k in range(n):
        yy = y + 44 + k * 34
        text(x, yy + 7, f"{k + 1}.", 22, 700, FIELD, "start", SERIF)
        words(x + 34, x + w - 10 - (k % 3) * 40, yy, 100 + k, 10)


def person(cx, cy, s=1.0, color=AMBER):
    add(f'<circle cx="{cx}" cy="{cy - 9 * s}" r="{8.5 * s}" fill="{color}"/>')
    add(f'<rect x="{cx - 14 * s}" y="{cy + 2 * s}" width="{28 * s}" height="{17 * s}" rx="{8 * s}" fill="{color}"/>')


def book(cx, cy, color="#FFFFFF"):
    add(f'<path d="M{cx - 20} {cy - 10}q10 -6 19 0v24q-9 -5 -19 0z" fill="{color}"/>')
    add(f'<path d="M{cx + 1} {cy - 10}q9 -6 19 0v24q-10 -5 -19 0z" fill="{color}" opacity=".75"/>')


def doc(cx, cy, color="#FFFFFF"):
    add(f'<path d="M{cx - 14} {cy - 18}h20l10 10v26h-30z" fill="{color}"/>')
    text(cx + 1, cy + 9, "W", 15, 700, WORDBLUE, "middle")


def tab(x, y, w, label, kind):
    fill, stroke, col = ((AMBER_BG, AMBER, AMBER_T) if kind == "plan" else ("#FFFFFF", "#AFC1DA", INK))
    add(f'<path d="M{x} {y + 64}V{y + 14}a14 14 0 0 1 14 -14H{x + w - 14}a14 14 0 0 1 14 14V{y + 64}Z" '
        f'fill="{fill}" stroke="{stroke}" stroke-width="2"/>')
    if kind == "plan":
        person(x + 34, y + 32)
    elif kind == "lib":
        book(x + 34, y + 32, "#2E6B5B")
    else:
        doc(x + 34, y + 32, WORDBLUE)
    text(x + 62, y + 40, label, 24, 700, col)


def steps(x, y, dx=0, dy=0, color="#FFFFFF", num="#1E355F"):
    for k, s in enumerate(("Match", "Insert", "Refresh", "Verify")):
        cx, cy = x + k * dx, y + k * dy
        add(f'<circle cx="{cx}" cy="{cy}" r="15" fill="{color}"/>')
        text(cx, cy + 7, str(k + 1), 19, 700, num, "middle")
        text(cx + 26, cy + 8, s, 24, 600, color)


if VAR == "docked":
    W, H = 1600, 650
    # inputs docked on the bar's top edge
    tab(150, 28, 330, "Your citation plan", "plan")
    tab(500, 28, 290, "Zotero library", "lib")
    tab(810, 28, 270, "Manuscript", "doc")
    # the dark focal bar
    add(f'<rect x="40" y="90" width="1520" height="110" rx="18" fill="url(#navy)"/>')
    add(f'<rect x="40" y="90" width="1520" height="110" rx="18" fill="url(#sheen)"/>')
    text(76, 158, "Zotero Word AutoCite", 32, 700, "#FFFFFF")
    steps(780, 145, dx=190)
    # the Word page coming out of the bar
    px, py, pw, ph = 110, 200, 1380, 420
    add(f'<rect x="{px}" y="{py - 10}" width="{pw}" height="{ph + 10}" fill="#FFFFFF" stroke="#D0D5DD" stroke-width="1.6" filter="url(#shadow)"/>')
    add(f'<rect x="38" y="186" width="1524" height="16" fill="#FFFFFF"/>')  # tuck the page under the bar
    add(f'<path d="M40 192H1560" stroke="#D0D5DD" stroke-width="0"/>')
    bx0, bx1 = px + 60, px + 860
    rows = [(258, None), (310, ("[1]", 0.42)), (362, None), (414, ("[2–4]", 0.55)), (466, None),
            (518, ("[5]", 0.30)), (570, None)]
    for i, (y, fld) in enumerate(rows):
        end = bx1 - (0 if i % 3 else 120)
        if fld:
            label, at = fld
            fx = bx0 + at * (bx1 - bx0)
            words(bx0, fx - 6, y, i)
            w = field(fx, y, label)
            if label == "[2–4]":
                # the sentence ends at this field; a short label sits beside it
                add(f'<rect x="{fx + w + 16:.0f}" y="{y - 16}" width="196" height="32" rx="16" fill="{FIELD}"/>')
                text(fx + w + 114, y + 7, "live Zotero field", 21, 700, "#FFFFFF", "middle")
            else:
                words(fx + w + 8, end, y, 50 + i)
        else:
            words(bx0, end, y, i)
    bibliography(px + 930, 280, 380)
else:
    W, H = 1380, 700
    lx, ly, lw, lh = 40, 86, 500, 530
    bx, bw = 590, 240
    rx, rw = 880, 460
    # headings: the human places the citations first, then AutoCite inserts them
    person(58, 44, 1.05)
    text(84, 54, "Human decides which paper goes where", 25, 700, AMBER_T)
    text(bx + bw / 2, 54, "AutoCite inserts", 25, 700, NAVY0, "middle")
    text(rx + rw, 54, "Live Zotero fields", 25, 700, INK, "end")
    # left page: the plan written as Word comments, framed in the human colour
    add(f'<rect x="{lx}" y="{ly}" width="{lw}" height="{lh}" rx="6" fill="#FFFFFF" stroke="{AMBER}" stroke-width="2.6" filter="url(#shadow)"/>')
    add(f'<rect x="{lx + 318}" y="{ly + 2}" width="{lw - 320}" height="{lh - 4}" fill="#F7F3EC"/>')
    rows = [150, 202, 254, 306, 358, 410, 462, 514, 566]
    marks = {1: "Paper 1", 3: "Papers 2–4", 5: "Paper 5"}
    for i, y in enumerate(rows):
        end = lx + 296 - (i % 3) * 26
        words(lx + 26, end, y, i)
        if i in marks:
            add(f'<rect x="{end - 50}" y="{y - 13}" width="54" height="26" fill="#FFE58F" opacity=".8"/>')
            add(f'<path d="M{end + 4} {y}H{lx + 330}" stroke="{AMBER}" stroke-width="1.8" stroke-dasharray="5 4"/>')
            add(f'<rect x="{lx + 330}" y="{y - 22}" width="156" height="44" rx="8" fill="#FFFFFF" stroke="{AMBER}" stroke-width="2"/>')
            text(lx + 408, y + 8, marks[i], 22, 700, AMBER_T, "middle")
    # the dark band
    add(f'<rect x="{bx}" y="{ly}" width="{bw}" height="{lh}" rx="18" fill="url(#navy)"/>')
    add(f'<rect x="{bx}" y="{ly}" width="{bw}" height="{lh}" rx="18" fill="url(#sheen)"/>')
    text(bx + bw / 2, ly + 56, "Zotero Word", 28, 700, "#FFFFFF", "middle")
    text(bx + bw / 2, ly + 90, "AutoCite", 28, 700, "#FFFFFF", "middle")
    steps(bx + 46, ly + 172, dy=68)
    # Zotero library: a full card docked under the band
    cy0 = ly + lh - 30
    add(f'<rect x="{bx - 14}" y="{cy0}" width="{bw + 28}" height="88" rx="16" fill="#FFFFFF" stroke="{FIELD}" stroke-width="2.6" filter="url(#shadow)"/>')
    # official Zotero app icon from zotero.org/support/brand, embedded unmodified
    import base64
    icon = base64.b64encode((Path(__file__).resolve().parent / "zotero-brand" / "zotero_256x256x32.png").read_bytes()).decode()
    add(f'<image href="data:image/png;base64,{icon}" x="{bx}" y="{cy0 + 12}" width="64" height="64"/>')
    text(bx + 72, cy0 + 40, "Zotero library", 24, 700, FIELD)
    text(bx + 72, cy0 + 68, "references, PDFs", 20, 400, SUB)
    # right page with live fields and bibliography
    add(f'<rect x="{rx}" y="{ly}" width="{rw}" height="{lh}" rx="6" fill="#FFFFFF" stroke="#D0D5DD" stroke-width="1.6" filter="url(#shadow)"/>')
    fields = {1: "[1]", 3: "[2–4]", 5: "[5]"}
    for i, y in enumerate(rows[:6]):
        end = rx + rw - 34 - (i % 3) * 36
        if i in fields:
            # each field opens its line, so the connector meets it at the page edge
            fx = rx + 28
            w = field(fx, y, fields[i])
            words(fx + w + 10, end, y, 60 + i)
            # straight connector from the decision, through the band's ports, to the field
            add(f'<path d="M{lx + 486} {y}H{bx}" stroke="{FIELD}" stroke-width="2.4"/>')
            add(f'<path d="M{bx + bw} {y}H{fx - 4}" stroke="{FIELD}" stroke-width="2.4"/>')
            for px_ in (bx, bx + bw):
                add(f'<circle cx="{px_}" cy="{y}" r="7" fill="#FFFFFF" stroke="{FIELD}" stroke-width="3"/>')
        else:
            words(rx + 28, end, y, i)
    bibliography(rx + 36, 470, rw - 72, 3)

defs = f"""<defs>
  <linearGradient id="navy" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="{NAVY0}"/><stop offset="1" stop-color="{NAVY1}"/></linearGradient>
  <linearGradient id="sheen" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#fff" stop-opacity=".14"/><stop offset=".5" stop-color="#fff" stop-opacity="0"/></linearGradient>
  <filter id="shadow" x="-10%" y="-10%" width="120%" height="130%"><feDropShadow dx="0" dy="6" stdDeviation="10" flood-color="#1F2933" flood-opacity=".14"/></filter>
  <filter id="blur"><feGaussianBlur stdDeviation="1.6"/></filter>
</defs>"""
svg = (f'<svg xmlns="http://www.w3.org/2000/svg" width="{W}" height="{H}" viewBox="0 0 {W} {H}" role="img" aria-labelledby="t d">\n'
       '<title id="t">Zotero Word AutoCite overview</title>\n'
       '<desc id="d">Your citation plan, your Zotero library, and your manuscript go into Zotero Word AutoCite, which matches, '
       'inserts, refreshes, and verifies. The result is a Word document whose citations are live Zotero fields, with a '
       'refreshable bibliography. Body text is blurred; the page is an illustration.</desc>\n'
       f'{defs}<rect width="{W}" height="{H}" fill="#FFFFFF"/>\n' + "\n".join(out) + "\n</svg>\n")
here = Path(__file__).resolve().parent
(here / "overview.svg").write_text(svg, encoding="utf-8")
(here / "overview.html").write_text(
    '<!doctype html><meta charset="utf-8"><style>html,body{margin:0;background:#fff}svg{display:block}</style>' + svg,
    encoding="utf-8")
print(VAR, W, H)
