import re, difflib
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

SRC = "/root/.claude/uploads/a7ca4361-b1fd-50b0-bc6b-9819d20c31ff/86ac40bd-Rubriques_Hierarchie_De_Gestion_1_1.md"
OUT = "/home/user/pigment/Mapping_Rubriques_Hierarchie_Gestion.xlsx"

# ---------- 1. Parse ----------
items = []
for l in open(SRC, encoding="utf-8"):
    l = l.strip()
    if not l or re.match(r"^lundi \d+ mai 2026", l) or re.match(r"^\d+/33$", l):
        continue
    m = re.match(r"^([A-Z0-9][A-Za-z0-9_.\-]*) - (.*)$", l)
    if m:
        items.append([m.group(1), m.group(2)])
    else:  # label wrapped on the next line
        items[-1][1] += " " + l

# ---------- 2. Hierarchy rules ----------
# NODES "code:depth"  -> node (has children) at this depth
# LEAVES "code=depth" -> leaf forced at this depth
# anything else       -> leaf, child of the last open node
SPEC = """
B1:0 A10R:1 A10:2 A100:3 A1008:3 A150:1 A15R:1 A20R:1 A20:2 A2R0:3 A200:4 A210:4 A220:4 A2R08:3 P5520=2
A25R:1 A2R5:2 A250:3 A3010=3 A2R58:2 A3008:3 A35:1 A350:2 A3509:2 A37:1 A40R:1 A400:2 A4009:2 A405:2 A4059:2
B2:0 B20:1 A6R1:2 A610:3 A620:3 A640:3 A6R19:2 A6109:3 A6209:3 A6409:3 P6R2:2 P62:3 P6R4:3
B30:1 A63:2 A630:3 A6309:3 A6R5:2 A650:3 A6509:3 P65:3 P63:2 P6R3:2 P6R5:2 P60:2 P6R0:3
B3:0 A1509:1 A5R5:1 A55:2 A70:2 P5510=2 P7010=2 A71:1 A755:1 A7559:2 B25:1 P6R6:2 P50:1 P52:1
B4:0 P20:1 P25:1 P2R0:2 P2R5:2
B5:0 B50:1 A4R5:2 A45:3 A80:3 P45:3 P8010=3 P3R0:2 P35:3 P3R0L=3 P4R4:2 P36:3 P4R4L=3
B51:1 P7R0:2 B52:1 A4R3:2 A4R4:2 B53:1 A8R5:2 A85:3 A90:3 P7540=3 P90:3
P1:0 P10:1 P100:2 P101:2 P104:2 P105:2 P1R0:2 P102:3 P103:3 P106:3 P15:1 P1540=2 P1R5:2 P1R9:1 A95:2 P95:2
R:0 R1:1 R00:2 R01:3 R10:4 R10B=5 R11:4 R2R0:4 R2D15R:5 R2R1:5 R2R2:5 R2R3:5 R275:6 R2D50=6
R03:3 R30:4 R32:4 R34:4 R342:5 R3446=5 R3R4:5 R04:2 R40:3 R42:3 R420:4 R424:4 R4R2:4 R4999=4 R44:3
R05:2 R5525=3 R5R1:3 R5R10:4 R5R15:4 R5R18:4 R5999=4 R5R19:4 R5R2:3
R06:2 R601=3 R610:3 R6335=3 R650:3 R660:3 R670:3 R6R1:3
R07:2 R720:3 R730:3 R740:3 R7R1:3 R7R2:3 R2:1 R3:1
XSN:0
XTF:0 XTFA:1 XTFA900:2 XTFA905:3 XTFA907=3 XTFA910:2 XTFA915:2 XTFA920=2 XTFCAF:2 XTFA000=3
XTFA1:3 XTFA105:4 XTFA110:4 XTFA195=4 XTFA2:3 XTFA11:4 XTFA120:5 XTFA125=5 XTFA130:5 XTFA135:5 XTFA136=5
XTFA140:4 XTFA142=4 XTFA145:4 XTFA200:3 XTFA3:3 XTFA30:4 XTFA310:5 XTFA315=4 XTFA350:3 XTFA353=3
XTFA50:3 XTFA600=3 XTFD100=1
XTFF:1 XTFF1:2 XTFF10:3 XTFF11:3 XTFF100=2 XTFF3=2 XTFF4:2 XTFF5=2 XTFF6:2 XTFF7:2 XTFF8:2 XTFF730:3 XTFF805=3 XTFF9=2
XTFI:1 XTFI1:2 XTFI10:3 XTFI2:2 XTFI3:2 XTFI20:3 XTFI015:4 XTFI025=4 XTFI4:2 XTFI150:3 XTFI155=3 XTFI420=2
XTFI5:2 XTFI500:3 XTFI6:2 XTFI600=2 XTFI7:2 XTFI8:2 XTFI810:3
XTS:0 XTZ:0 XTZ00:1 XTFD200TE:2 XTZ10:1 XTFD30EF:2 XTFD315EF=2
XZ55:0 ZDAEFN:0 ZIMPOT:0 ZKPI:0 ZKPI00:1 ZKPI10:1 ZKPI20:1 ZKPI202:2 ZKRN:0 ZPLAN:0 ZTM:0
"""
nodes, forced = {}, {}
for tok in SPEC.split():
    if ":" in tok:
        c, d = tok.split(":"); nodes[c] = int(d)
    else:
        c, d = tok.split("="); forced[c] = int(d)

# Items whose placement is a deduction the source .md cannot confirm (indentation lost)
UNCERTAIN_RANGES = [("P5520",), ("A3010",), ("A3008", "A30109"), ("P5510", "P7010"), ("B25",), ("P52",),
    ("P8010", "P8030L"), ("P7540", "P7620"), ("P1R9",), ("R10B",), ("R275", "R2D95"), ("R3446",), ("R5525",),
    ("R5999",), ("R601", "R6120"), ("R6335",), ("XTFA901",), ("XTFA905", "XTFA907"), ("XTFA920",),
    ("XTFA11", "XTFA136"), ("XTFA195",), ("XTFA142",), ("XTFA315", "XTFA999"), ("XTFD100",), ("XTFF100", "XTFF205"),
    ("XTFF3",), ("XTFF5",), ("XTFF805",), ("XTFF9",), ("XTFI010",), ("XTFI030",), ("XTFI110",), ("XTFI020",),
    ("XTFI155", "XTFI420"), ("XTFI205", "XTFI615"), ("XTFD30EF", "XTFI140EF"), ("ZIMPOT", "XITPT80"),
    ("ZKPI202", "FTMC"), ("ZTM", "TM4400")]
codes = [c for c, _ in items]
idx = {c: i for i, c in enumerate(codes)}
uncertain = set()
for r in UNCERTAIN_RANGES:
    a, b = idx[r[0]], idx[r[-1]]
    uncertain.update(codes[a:b + 1])
CONFIRMED = set(codes[:idx["A20"] + 1])  # visible on the screenshot (page 1)

stack, rows = [], []
for code, label in items:
    is_node = code in nodes
    d = nodes.get(code, forced.get(code))
    if d is None:
        d = stack[-1][0] + 1
    while stack and stack[-1][0] >= d:
        stack.pop()
    assert (stack[-1][0] if stack else -1) == d - 1, (code, d, stack[-1:])
    parent = stack[-1][1] if stack else ""
    path = [c for _, c in stack] + [code]
    rows.append(dict(code=code, label=label, depth=d, parent=parent, node=is_node, path=path))
    if is_node:
        stack.append((d, code))
for c in nodes:
    assert c in idx, c
for r in rows:  # a declared node without children would be a spec error
    if r["node"]:
        assert any(x["parent"] == r["code"] for x in rows), r["code"]

def status(c):
    if c in CONFIRMED: return "Confirmé (capture)"
    if c in uncertain: return "À vérifier"
    return "Déduit"

# ---------- 3. Labels ----------
SPLIT_FIX = {  # manual short-label cut where the automatic split is wrong
    "A35109": "dep GW stés MEE", "A6436": "Dériv act.JV - BFR(chge) - CFH", "A6437": "Dériv act.JV - BFR(chge) - FVH",
    "A6438": "Dériv act. JV - BFR(chge) - NH", "P6436": "Dériv pas.JV - BFR(chge) - CFH", "P6437": "Dériv pas.JV- BFR(chge) - FVH",
    "P6438": "Dériv pas. JV - BFR(chge) - NH", "A70109": "Écrêtement - Net IDA - C", "P10": "Capitaux propres - Gpe",
    "R7220": "ID régul N-1", "A15109": "dep GW", "XIA315BN": "Diff. temporelles \"régul\"",
}
def split_label(t, code=None):
    if code in SPLIT_FIX and t.startswith(SPLIT_FIX[code] + " - "):
        return SPLIT_FIX[code], t[len(SPLIT_FIX[code]) + 3:]
    parts = t.split(" - ")
    if len(parts) == 1:
        return t, t
    best = None
    for i in range(1, len(parts)):
        s, l = " - ".join(parts[:i]), " - ".join(parts[i:])
        score = difflib.SequenceMatcher(None, s.lower(), l.lower()[:max(len(s), 1) + 5]).ratio()
        if len(s) > 30: score -= 1  # short labels are max 30 chars in the source
        if s.count("(") != s.count(")") or l.count("(") != l.count(")"): score -= 1
        if best is None or score > best[0]:
            best = (score, s, l)
    return best[1], best[2]

# ---------- 4. Excel ----------
HDR = Font(bold=True, color="FFFFFF"); FILL = PatternFill("solid", fgColor="1F4E78")
WARN = PatternFill("solid", fgColor="FFF2CC")
def header(ws, cols):
    ws.append(cols)
    for c in ws[1]:
        c.font, c.fill, c.alignment = HDR, FILL, Alignment(vertical="center")
    ws.freeze_panes = "A2"; ws.auto_filter.ref = ws.dimensions
def widths(ws, w):
    for i, x in enumerate(w, 1):
        ws.column_dimensions[get_column_letter(i)].width = x

maxd = max(r["depth"] for r in rows)
wb = Workbook()

# Sheet 1: flattened mapping, one row per leaf
ws = wb.active; ws.title = "Mapping_Niveaux"
cols = [f"Niveau {i}" for i in range(maxd + 1)] + ["Code feuille", "Profondeur", "Statut"]
header(ws, cols)
for r in rows:
    if r["node"]: continue
    p = r["path"] + [""] * (maxd + 1 - len(r["path"]))
    ws.append(p + [r["code"], r["depth"], status(r["code"])])
    if status(r["code"]) == "À vérifier":
        for c in ws[ws.max_row]: c.fill = WARN
header_done = ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{ws.max_row}"
widths(ws, [12] * (maxd + 1) + [14, 11, 18])

# Sheet 2: parent / child
ws = wb.create_sheet("Parent_Enfant")
header(ws, ["Ordre", "Code", "Code parent", "Niveau", "Type", "Statut"])
for i, r in enumerate(rows, 1):
    ws.append([i, r["code"], r["parent"], r["depth"], "Noeud" if r["node"] else "Feuille", status(r["code"])])
    if status(r["code"]) == "À vérifier":
        for c in ws[ws.max_row]: c.fill = WARN
ws.auto_filter.ref = f"A1:F{ws.max_row}"
widths(ws, [8, 14, 14, 8, 10, 18])

# Sheet 3: labels (for Pigment upload)
ws = wb.create_sheet("Libelles")
header(ws, ["Code", "Libellé court", "Libellé long", "Code - Libellé", "Libellé source complet"])
for r in rows:
    s, l = split_label(r["label"], r["code"])
    ws.append([r["code"], s, l, f'{r["code"]} - {l}', r["label"]])
ws.auto_filter.ref = f"A1:E{ws.max_row}"
widths(ws, [14, 32, 70, 80, 90])

# Sheet 4: one list per level (Pigment dimension lists)
ws = wb.create_sheet("Listes_par_niveau")
header(ws, ["Niveau", "Code", "Code parent", "Libellé long"])
for d in range(maxd + 1):
    for r in rows:
        if r["depth"] == d:
            ws.append([d, r["code"], r["parent"], split_label(r["label"], r["code"])[1]])
ws.auto_filter.ref = f"A1:D{ws.max_row}"
widths(ws, [8, 14, 14, 80])

# Sheet 5: readme
ws = wb.create_sheet("Lisez-moi")
notes = [
    "Source : Rubriques_Hierarchie_De_Gestion (export PDF converti en .md, 33 pages, 1 772 rubriques).",
    "L'indentation du PDF a été perdue dans le .md : la hiérarchie est reconstruite à partir de l'ordre des rubriques et de la logique des codes.",
    "Logique : une rubrique 'nœud' (ex. A1008) regroupe les rubriques qui la suivent (A10108, A10109, ...), elle-même rattachée à un nœud de niveau supérieur (A10 > A10R > B1).",
    "Statut 'Confirmé (capture)' : visible sur la capture de la page 1. 'Déduit' : logique de code claire. 'À vérifier' (surligné jaune) : rattachement ambigu sans l'indentation.",
    "Sections techniques (ZIMPOT, ZTM, ZKRN, ZPLAN, XZ55, ZDAEFN, XSN, XTS) : rubriques mises à plat sous leur rubrique racine.",
    "Mapping_Niveaux : 1 ligne par rubrique feuille, avec toute la chaîne de ses parents (Niveau 0 = racine). Hiérarchie irrégulière : colonnes vides au-delà de la profondeur de la feuille.",
    "Parent_Enfant : format parent/enfant (recommandé pour Pigment). Libelles : code + libellés pour l'import. Listes_par_niveau : une liste par niveau avec son parent.",
    f"Profondeur max : {maxd}. Nœuds : {sum(r['node'] for r in rows)}. Feuilles : {sum(not r['node'] for r in rows)}. À vérifier : {sum(status(r['code'])=='À vérifier' for r in rows)}.",
]
for n in notes: ws.append([n])
ws.column_dimensions["A"].width = 160
wb.move_sheet("Lisez-moi", offset=-4)
wb.save(OUT)
print("OK", OUT, len(rows), "maxdepth", maxd, notes[-1])
