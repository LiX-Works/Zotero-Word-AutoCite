"""Build the README overview figure (assets/overview.png) with the Python standard library.

Writes overview.svg and an overview.html wrapper next to the --output path; render the
HTML to PNG at 2x (3200 x 1400), e.g. with headless Chrome or Edge. The Word page is a
stylized illustration.
"""
from pathlib import Path
import argparse
import math
from xml.sax.saxutils import escape

cli = argparse.ArgumentParser(description=__doc__)
cli.add_argument("--output", type=Path, default=Path(__file__).resolve().parent / "overview.svg")
args = cli.parse_args()

W, H = 1600, 700
FONT = "Arial,Helvetica,sans-serif"
C = dict(
    ink="#16203A", sub="#2C3850", line="#2E3A55", card="#FFFFFF", cardStroke="#C3CDDD",
    amber="#E8960F", amberText="#A9650A", lead0="#22386A", lead1="#3866A6",
    plan="#5A3CC2", planTint="#F1ECFD", refs="#2F7D6B", doc="#2B579A",
    field="#B4232F", fieldTint="#FBE3E4", bar="#D5DBE5",
)

out = []
add = out.append


def f(v):
    return f"{v:.1f}".rstrip("0").rstrip(".")


def text(x, y, s, size=26, weight=700, fill=C["ink"], anchor="start", extra=""):
    add(f'<text x="{f(x)}" y="{f(y)}" font-size="{size}" font-weight="{weight}" fill="{fill}" '
        f'text-anchor="{anchor}" {extra}>{escape(s)}</text>')


def line(d, stroke, width, extra=""):
    add(f'<path d="{d}" fill="none" stroke="{stroke}" stroke-width="{width}" '
        f'stroke-linecap="round" stroke-linejoin="round" {extra}/>')


def head(x2, y2, ux, uy, color, size=22, half=11):
    bx, by = x2 - size * ux, y2 - size * uy
    add(f'<path d="M{f(x2)} {f(y2)} L{f(bx - half * uy)} {f(by + half * ux)} '
        f'L{f(bx + half * uy)} {f(by - half * ux)}Z" fill="{color}"/>')


def arrow(points, color=C["line"], width=5.5):
    # The shaft stops inside the head, so the tip stays a sharp point.
    (x1, y1), (x2, y2) = points[-2], points[-1]
    n = math.hypot(x2 - x1, y2 - y1)
    ux, uy = (x2 - x1) / n, (y2 - y1) / n
    end = (x2 - 13 * ux, y2 - 13 * uy)
    d = "M" + " L".join(f"{f(a)} {f(b)}" for a, b in points[:-1] + [end])
    add(f'<path d="{d}" fill="none" stroke="{color}" stroke-width="{width}" stroke-linejoin="round"/>')
    head(x2, y2, ux, uy, color)


def card(x, y, w, h, rx=22):
    add(f'<rect x="{f(x)}" y="{f(y)}" width="{f(w)}" height="{f(h)}" rx="{rx}" fill="{C["card"]}" '
        f'stroke="{C["cardStroke"]}" stroke-width="1.8" filter="url(#shadow)"/>')


# ---------- icons (solid) ----------
def icon_plan(cx, cy, col):
    for k in range(3):
        y = cy - 16 + k * 16
        add(f'<circle cx="{cx - 14}" cy="{y}" r="4.5" fill="{col}"/>')
        line(f"M{cx - 4} {y}h22", col, 5)


def icon_refs(cx, cy, col):
    for k, (dx, h) in enumerate(((-16, 40), (-5, 44), (6, 38))):
        add(f'<rect x="{cx + dx}" y="{cy + 22 - h}" width="10" height="{h}" rx="2" fill="{col}" opacity="{1 - k * .2}"/>')
    add(f'<rect x="{cx + 17}" y="{cy - 14}" width="10" height="36" rx="2" fill="{col}" transform="rotate(14 {cx + 22} {cy + 22})"/>')


def icon_doc(cx, cy, col):
    add(f'<path d="M{cx - 17} {cy - 22} h24 l10 10 v34 h-34z" fill="{col}"/>')
    text(cx - 1, cy + 9, "W", 18, 700, "#FFFFFF", "middle")


# ---------- layout ----------
IN_X, IN_W, IN_H = 24, 390, 130
IN_Y = [62, 222, 382]
AG = dict(x=510, y=140, w=316, h=354)
PG = dict(x=884, y=36, w=400, h=556)
MIDY = AG["y"] + AG["h"] / 2

inputs = [
    (icon_plan, C["plan"], "Citation plan", "Notes, tables, Word comments"),
    (icon_refs, C["refs"], "References", "Zotero items or PDFs"),
    (icon_doc, C["doc"], "Manuscript", "Your Word document"),
]
BUS = 462
for y, (icon, col, title, sub) in zip(IN_Y, inputs):
    card(IN_X, y, IN_W, IN_H)
    icon(IN_X + 50, y + IN_H / 2, col)
    text(IN_X + 92, y + 56, title, 28)
    text(IN_X + 92, y + 92, sub, 21, 400, C["sub"])
    line(f"M{IN_X + IN_W + 7} {y + IN_H / 2}H{BUS}", C["line"], 5.5)
# the citation plan carries the user's decisions
tx, ty = IN_X + IN_W - 150, IN_Y[0] - 16
add(f'<rect x="{tx}" y="{ty}" width="168" height="34" rx="17" fill="{C["plan"]}"/>')
text(tx + 84, ty + 23, "Your decisions", 18, 700, "#FFFFFF", "middle")
line(f"M{BUS} {IN_Y[0] + IN_H / 2}V{IN_Y[2] + IN_H / 2}", C["line"], 5.5)
arrow([(BUS, MIDY), (AG["x"] - 7, MIDY)])

# ---------- agent ----------
add(f'<rect x="{AG["x"]}" y="{AG["y"]}" width="{AG["w"]}" height="{AG["h"]}" rx="28" fill="url(#leadGrad)" filter="url(#leadShadow)"/>')
add(f'<rect x="{AG["x"]}" y="{AG["y"]}" width="{AG["w"]}" height="{AG["h"]}" rx="28" fill="url(#sheen)"/>')
text(AG["x"] + AG["w"] / 2, AG["y"] + 62, "AI Agent", 38, 700, "#FFFFFF", "middle")
line(f"M{AG['x'] + 36} {AG['y'] + 92}H{AG['x'] + AG['w'] - 36}", "#FFFFFF", 1.5, 'opacity=".35"')
steps = ["Match references", "Insert Zotero fields", "Refresh in Word", "Verify bibliography"]
for k, s in enumerate(steps):
    y = AG["y"] + 140 + k * 56
    add(f'<circle cx="{AG["x"] + 50}" cy="{y - 8}" r="17" fill="#FFFFFF" opacity=".95"/>')
    text(AG["x"] + 50, y, str(k + 1), 21, 700, C["lead0"], "middle")
    text(AG["x"] + 80, y, s, 23, 700, "#FFFFFF")

arrow([(AG["x"] + AG["w"] + 7, MIDY), (PG["x"] - 7, MIDY)])

# ---------- Word page ----------
px, py, pw, ph = PG["x"], PG["y"], PG["w"], PG["h"]
add(f'<rect x="{px}" y="{py}" width="{pw}" height="{ph}" rx="14" fill="#FFFFFF" stroke="{C["cardStroke"]}" stroke-width="1.8" filter="url(#pageShadow)"/>')
add(f'<path d="M{px} {py + 46} V{py + 14} a14 14 0 0 1 14 -14 H{px + pw - 14} a14 14 0 0 1 14 14 V{py + 46}Z" fill="{C["doc"]}"/>')
text(px + 24, py + 31, "manuscript.docx", 20, 700, "#FFFFFF")
fields = {}


def body_row(y, parts):
    x = px + 34
    for kind, val in parts:
        if kind == "bar":
            add(f'<rect x="{f(x)}" y="{y - 6}" width="{val}" height="12" rx="6" fill="{C["bar"]}"/>')
            x += val + 10
        else:
            w = len(val) * 11 + 18
            add(f'<rect x="{f(x)}" y="{y - 15}" width="{f(w)}" height="30" rx="6" fill="{C["fieldTint"]}" stroke="{C["field"]}" stroke-width="1.6"/>')
            text(x + w / 2, y + 7, val, 19, 700, C["field"], "middle")
            fields[val] = (x, y, w)
            x += w + 10


rows = [
    [("bar", 320)],
    [("bar", 130), ("field", "[1]"), ("bar", 120)],
    [("bar", 330)],
    [("bar", 80), ("field", "[2–4]"), ("bar", 150)],
    [("bar", 280)],
    [("bar", 190), ("field", "[5]"), ("bar", 70)],
    [("bar", 220)],
]
for k, parts in enumerate(rows):
    body_row(py + 92 + k * 36, parts)
# bibliography
by = py + 92 + len(rows) * 36 + 16
text(px + 34, by, "References", 22, 700, C["ink"])
add(f'<rect x="{px + 24}" y="{by + 16}" width="{pw - 48}" height="{5 * 30 + 14}" rx="8" fill="{C["fieldTint"]}" opacity=".55"/>')
for k, w in enumerate((270, 230, 280, 210, 250)):
    y = by + 46 + k * 30
    text(px + 42, y + 6, f"{k + 1}.", 17, 700, C["field"])
    add(f'<rect x="{px + 70}" y="{y - 5}" width="{w}" height="11" rx="5.5" fill="{C["bar"]}"/>')

# ---------- callouts ----------
def callout(y, title, sub, target):
    cx0 = px + pw + 40
    add(f'<rect x="{cx0}" y="{y - 44}" width="240" height="88" rx="18" fill="#FFFFFF" stroke="{C["field"]}" stroke-width="2.2" filter="url(#shadow)"/>')
    # refresh glyph
    gx, gy = cx0 + 36, y
    add(f'<path d="M{gx + 13} {gy - 8} A15 15 0 1 0 {gx + 15} {gy + 4}" fill="none" stroke="{C["field"]}" stroke-width="4" stroke-linecap="round"/>')
    add(f'<path d="M{gx + 17} {gy - 18} L{gx + 16} {gy - 4} L{gx + 3} {gy - 9}Z" fill="{C["field"]}"/>')
    text(cx0 + 62, y - 6, title, 20, 700, C["ink"])
    text(cx0 + 62, y + 22, sub, 18, 400, C["sub"])
    tx, ty = target
    line(f"M{cx0 - 4} {y}L{f(tx)} {f(ty)}", C["field"], 2.2, 'stroke-dasharray="6 5"')
    add(f'<circle cx="{f(tx)}" cy="{f(ty)}" r="4.5" fill="{C["field"]}"/>')


fx, fy, fw = fields["[2–4]"]
callout(py + 150, "Live Zotero fields", "not typed numbers", (px + pw - 30, fy))
line(f"M{f(fx + fw + 4)} {fy}H{px + pw - 30}", C["field"], 2.2, 'stroke-dasharray="6 5"')
callout(by + 90, "Refreshable", "bibliography", (px + pw - 24, by + 90))

# ---------- loop ----------
ly = H - 46
arrow([(px + pw / 2, py + ph + 7), (px + pw / 2, ly), (AG["x"] + AG["w"] / 2, ly), (AG["x"] + AG["w"] / 2, AG["y"] + AG["h"] + 7)], C["amber"])
lx = (px + pw / 2 + AG["x"] + AG["w"] / 2) / 2
add(f'<rect x="{f(lx - 112)}" y="{ly - 22}" width="224" height="44" rx="22" fill="#FFFFFF"/>')
text(lx, ly + 9, "adjust & refresh", 25, 700, C["amberText"], "middle")

defs = f"""<defs>
  <linearGradient id="leadGrad" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="{C['lead0']}"/><stop offset="1" stop-color="{C['lead1']}"/></linearGradient>
  <linearGradient id="sheen" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#fff" stop-opacity=".16"/><stop offset=".45" stop-color="#fff" stop-opacity="0"/></linearGradient>
  <filter id="shadow" x="-20%" y="-30%" width="140%" height="170%"><feDropShadow dx="0" dy="5" stdDeviation="8" flood-color="#16203A" flood-opacity=".10"/></filter>
  <filter id="pageShadow" x="-20%" y="-20%" width="140%" height="140%"><feDropShadow dx="0" dy="10" stdDeviation="14" flood-color="#16203A" flood-opacity=".16"/></filter>
  <filter id="leadShadow" x="-30%" y="-30%" width="160%" height="170%"><feDropShadow dx="0" dy="10" stdDeviation="14" flood-color="#22386A" flood-opacity=".32"/></filter>
</defs>"""
svg = (f'<svg xmlns="http://www.w3.org/2000/svg" width="{W}" height="{H}" viewBox="0 0 {W} {H}" '
       f'font-family="{FONT}" role="img" aria-labelledby="title desc">\n'
       '<title id="title">Zotero Word AutoCite overview</title>\n'
       '<desc id="desc">Your citation plan, your references (Zotero items or PDFs), and your Word manuscript go to an AI agent, '
       'which matches references, inserts Zotero fields, refreshes in Word, and verifies the bibliography. The result is a '
       'Word document with live, refreshable Zotero citation fields and bibliography; existing citations can be adjusted and '
       'refreshed again. The page is a stylized illustration.</desc>\n'
       f'<rect width="{W}" height="{H}" fill="#FFFFFF"/>\n{defs}\n' + "\n".join(out) + "\n</svg>\n")
args.output.write_text(svg, encoding="utf-8")
args.output.with_suffix(".html").write_text(
    '<!doctype html><meta charset="utf-8"><style>html,body{margin:0;background:#fff}svg{display:block}</style>' + svg,
    encoding="utf-8")
print(args.output)
