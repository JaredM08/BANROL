#!/usr/bin/env python3
"""
MOTOR — WEEKLY BAN-ROL REPORT (QST)
Recibe los 2 PDF (inv705 + inv315, en cualquier orden) y produce los 2 Excel
actualizados respetando fórmulas, formato y celdas fijas.

Uso CLI (para probar local):
    python3 banrol_motor.py inv705.pdf inv315.pdf --out ./salida

Uso microservicio (para n8n):
    python3 banrol_motor.py --serve         # levanta en :8000
    POST /generate  (multipart: 2 archivos PDF)  ->  devuelve un ZIP con los 2 .xlsx

Reglas de negocio (confirmadas):
  WEEKS OF COVERAGE, por tela:
    B MIN, C MAX, F ON ORDER, E STOCK(=ON HAND)  <- inv705  (valores)
    H WEEKLY AVERAGE = monthly avg / 4           <- inv315  (fórmula =avg/4)
    D TARGET, G TOTAL, I WEEKS OF COVERAGE        = fórmulas (no se tocan)
    J LEAD TIME, K MINIMUN BUY, L NOTES           = fijos    (no se tocan)
    A1 = fecha del reporte
  MIN-MAX: mismas columnas; se reemplaza la fecha + ON HAND en la última columna.
"""
import os, re, io, sys, zipfile, argparse, datetime as dt
import pdfplumber, openpyxl
from openpyxl.utils import get_column_letter

# Al empaquetar con PyInstaller, los archivos van a una carpeta temporal (_MEIPASS)
BASE = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
HERE = BASE
TPL_WEEKS  = os.path.join(HERE, "templates", "weeks_template.xlsx")
TPL_MINMAX = os.path.join(HERE, "templates", "minmax_template.xlsx")
ALT = {"SDBR-300-BK (SDBR-72-BK)": "SDBR-72-BK"}   # nombres alternos en MIN-MAX

# ---------------- lectura de PDFs ----------------
def _text(src):
    with pdfplumber.open(src) as pdf:
        return "\n".join((p.extract_text() or "") for p in pdf.pages)

def classify(t):
    if "LIST OF INVENTORY BY WAREHOUSE" in t or "inv705" in t: return "inv705"
    if "Min/Max Proj" in t or "Monthly" in t or "inv315" in t: return "inv315"
    return None

def parse_705(t):
    """{PART: {onhand, onorder, min, max}}"""
    out, rx = {}, re.compile(r'^([A-Z0-9][A-Z0-9\-/]+)\s+.*?((?:-?\d[\d,]*\s+){7}-?\d[\d,]*)\s*$')
    for ln in t.splitlines():
        m = rx.match(ln.strip())
        if not m: continue
        n = [int(x.replace(",", "")) for x in m.group(2).split()]
        if len(n) == 8:
            out[m.group(1)] = dict(onhand=n[0], onorder=n[4], min=n[6], max=n[7])
    return out

def parse_315(t):
    """{PART: monthly_average}"""
    out, rx = {}, re.compile(r'^([A-Z][A-Z0-9\-/]+)\s+(\d+)\s+([\d,]+)\s')
    for ln in t.splitlines():
        m = rx.match(ln.strip())
        if m: out[m.group(1)] = int(m.group(3).replace(",", ""))
    return out

def date_from_705(t):
    m = re.search(r'RUN DATE:\s*(\d{2})/(\d{2})/(\d{2})', t)
    if m:
        mm, dd, yy = m.groups()
        return dt.date(2000 + int(yy), int(mm), int(dd))
    return dt.date.today()

# ---------------- construcción de Excel ----------------
def build_weeks(d705, avg, fecha):
    wb = openpyxl.load_workbook(TPL_WEEKS); ws = wb["Sheet1"]
    ws["A1"] = dt.datetime.combine(fecha, dt.time())
    for r in range(6, ws.max_row + 1):
        c = ws.cell(r, 1).value
        if not isinstance(c, str) or c.startswith("("): continue
        c = c.strip(); d = d705.get(c)
        if d:
            ws.cell(r, 2).value = d["min"]      # B
            ws.cell(r, 3).value = d["max"]      # C
            ws.cell(r, 5).value = d["onhand"]   # E (valor)
            ws.cell(r, 6).value = d["onorder"]  # F
        if c in avg:
            ws.cell(r, 8).value = f"={avg[c]}/4" # H (fórmula)
    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()

def build_minmax(d705, fecha):
    wb = openpyxl.load_workbook(TPL_MINMAX); ws = wb["Sheet1"]
    col = ws.max_column
    ws.cell(3, col).value = dt.datetime.combine(fecha, dt.time())
    ws.cell(3, col).number_format = "m/d/yyyy"
    for r in range(4, ws.max_row + 1):
        c = ws.cell(r, 1).value
        if not isinstance(c, str) or c.startswith("("): continue
        d = d705.get(ALT.get(c.strip(), c.strip()))
        ws.cell(r, col).value = d["onhand"] if d else 0
    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()

def run(pdf_sources):
    """pdf_sources: lista de rutas o file-like. Devuelve (fecha, weeks_bytes, minmax_bytes)."""
    t705 = t315 = None
    for s in pdf_sources:
        t = _text(s); k = classify(t)
        if k == "inv705": t705 = t
        elif k == "inv315": t315 = t
    if t705 is None: raise ValueError("No se encontró el inv705 entre los PDF.")
    d705 = parse_705(t705)
    avg  = parse_315(t315) if t315 else {}
    fecha = date_from_705(t705)
    return fecha, build_weeks(d705, avg, fecha), build_minmax(d705, fecha)

# ---------------- CLI ----------------
def main_cli(a):
    fecha, wks, mm = run(a.pdfs)
    os.makedirs(a.out, exist_ok=True)
    p1 = os.path.join(a.out, f"WEEKS_OF_COVERAGE_{fecha:%m-%d-%y}.xlsx")
    p2 = os.path.join(a.out, f"MIN-MAX_{fecha:%m-%d-%y}.xlsx")
    open(p1, "wb").write(wks); open(p2, "wb").write(mm)
    print("OK", fecha, "->", p1, "|", p2)

# ---------------- microservicio ----------------
PAGE = """<!doctype html><html lang=es><head><meta charset=utf-8>
<meta name=viewport content="width=device-width,initial-scale=1">
<title>BANROL — generar reporte</title><style>
*{box-sizing:border-box;font-family:system-ui,Segoe UI,Roboto,sans-serif}
body{margin:0;background:#0f172a;color:#e2e8f0;display:flex;min-height:100vh;
align-items:center;justify-content:center;padding:24px}
.card{background:#1e293b;border:1px solid #334155;border-radius:16px;padding:32px;
max-width:440px;width:100%;box-shadow:0 10px 40px rgba(0,0,0,.4)}
h1{margin:0 0 4px;font-size:22px}p.sub{margin:0 0 24px;color:#94a3b8;font-size:14px}
label{display:block;border:2px dashed #475569;border-radius:12px;padding:28px;
text-align:center;cursor:pointer;transition:.15s;color:#94a3b8}
label:hover{border-color:#38bdf8;color:#e2e8f0}
input[type=file]{display:none}
#names{font-size:13px;color:#38bdf8;margin-top:12px;min-height:18px;text-align:center}
button{width:100%;margin-top:20px;background:#38bdf8;color:#0f172a;border:0;
border-radius:10px;padding:14px;font-size:16px;font-weight:600;cursor:pointer}
button:disabled{background:#334155;color:#64748b;cursor:not-allowed}
a.dl{display:block;background:#22c55e;color:#052e16;text-decoration:none;
border-radius:10px;padding:14px;font-weight:600;text-align:center;margin-top:12px}
.err{color:#f87171;margin-top:16px;font-size:14px}
</style></head><body><div class=card>
<h1>Reporte BANROL</h1><p class=sub>Sube los 2 PDF (inv705 + inv315). El orden no importa.</p>
%%BODY%%
</div></body></html>"""

FORM = """<form method=post action=/ui enctype=multipart/form-data>
<label for=f>📄 Da clic para elegir los 2 PDF<div id=names></div></label>
<input id=f name=pdfs type=file accept=.pdf multiple onchange="
document.getElementById('names').textContent=[...this.files].map(x=>x.name).join(' · ');
document.getElementById('go').disabled=this.files.length<2;">
<button id=go type=submit disabled>Generar Excel</button></form>
%%MSG%%"""

def create_app():
    import base64
    from flask import Flask, request, jsonify
    app = Flask(__name__)

    @app.get("/")
    def home():
        return PAGE.replace("%%BODY%%", FORM.replace("%%MSG%%", ""))

    @app.post("/ui")
    def ui():
        files = request.files.getlist("pdfs")
        try:
            if len(files) < 2:
                raise ValueError("Sube los 2 PDF.")
            srcs = [io.BytesIO(f.read()) for f in files]
            fecha, wks, mm = run(srcs)
        except Exception as e:
            msg = f'<div class=err>⚠️ {e}</div>'
            return PAGE.replace("%%BODY%%", FORM.replace("%%MSG%%", msg)), 400
        mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        def link(name, data):
            b = base64.b64encode(data).decode()
            return f'<a class=dl download="{name}" href="data:{mime};base64,{b}">⬇ {name}</a>'
        body = (f'<p class=sub>Listo — reporte del {fecha:%d/%m/%Y}. Descarga ambos:</p>'
                + link(f"WEEKS_OF_COVERAGE_{fecha:%m-%d-%y}.xlsx", wks)
                + link(f"MIN-MAX_{fecha:%m-%d-%y}.xlsx", mm)
                + '<a class=dl style="background:#334155;color:#e2e8f0;margin-top:20px" href="/">↩ Generar otro</a>')
        return PAGE.replace("%%BODY%%", body)

    @app.post("/generate")
    def generate():
        files = list(request.files.values())
        if len(files) < 2:
            return {"error": "Envía los 2 PDF (inv705 + inv315)."}, 400
        srcs = [io.BytesIO(f.read()) for f in files]
        fecha, wks, mm = run(srcs)
        b64 = lambda b: base64.b64encode(b).decode()
        # Devuelve los 2 archivos SUELTOS (base64), sin zip
        return jsonify({
            "date": f"{fecha:%m-%d-%y}",
            "files": [
                {"filename": f"WEEKS_OF_COVERAGE_{fecha:%m-%d-%y}.xlsx", "b64": b64(wks)},
                {"filename": f"MIN-MAX_{fecha:%m-%d-%y}.xlsx",           "b64": b64(mm)},
            ],
        })
    @app.get("/health")
    def health(): return {"ok": True}
    return app

def serve(port, open_browser=False):
    if open_browser:
        import threading, webbrowser
        threading.Timer(1.5, lambda: webbrowser.open(f"http://localhost:{port}")).start()
    create_app().run(host="0.0.0.0", port=port)

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("pdfs", nargs="*", help="rutas de los 2 PDF")
    ap.add_argument("--out", default="./salida")
    ap.add_argument("--serve", action="store_true")
    ap.add_argument("--port", type=int, default=8000)
    a = ap.parse_args()
    frozen = getattr(sys, "frozen", False)   # True cuando corre como .exe
    if a.pdfs:
        main_cli(a)
    elif a.serve or frozen:
        # doble clic al .exe -> levanta el programa y abre el navegador
        serve(a.port, open_browser=True)
    else:
        ap.print_help()
