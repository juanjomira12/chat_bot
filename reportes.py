"""Generación de reportes en Excel y Word con gráficas."""

from collections import defaultdict
from datetime import datetime
from io import BytesIO
from pathlib import Path

import matplotlib

matplotlib.use("Agg")  # sin ventanas: solo genera imágenes
import matplotlib.pyplot as plt
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor
from openpyxl import Workbook
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Alignment, Font, PatternFill

LOGO = Path(__file__).resolve().parent / "logo.png"
AZUL, DORADO, VERDE, ROJO = "1F4E9A", "C9A24A", "2E8B57", "C0392B"
PALETA = ["#1F4E9A", "#C9A24A", "#4A90D9", "#E3C77A", "#2E8B57", "#C0392B",
          "#7F8C8D", "#8E44AD", "#16A085", "#D35400"]
MESES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]
MONEDA = '"$"#,##0'


def _pesos(v: float) -> str:
    return "$" + f"{v:,.0f}".replace(",", ".")


def _mes(clave: str) -> str:  # "2026-11" -> "Nov 2026"
    return f"{MESES[int(clave[5:7]) - 1]} {clave[:4]}"


def analizar(filas: list[tuple]) -> dict:
    """filas: (fecha, tipo, categoria, descripcion, monto)."""
    ingresos = sum(f[4] for f in filas if f[1] == "ingreso")
    gastos = sum(f[4] for f in filas if f[1] == "gasto")
    por_cat: dict[str, float] = defaultdict(float)
    por_mes: dict[str, list[float]] = defaultdict(lambda: [0.0, 0.0])
    for fecha, tipo, cat, _, monto in filas:
        if tipo == "gasto":
            por_cat[cat or "Otros"] += monto
            por_mes[fecha[:7]][1] += monto
        else:
            por_mes[fecha[:7]][0] += monto
    return {
        "ingresos": ingresos,
        "gastos": gastos,
        "balance": ingresos - gastos,
        "categorias": sorted(por_cat.items(), key=lambda x: x[1], reverse=True),
        "meses": sorted(por_mes.items()),
    }


# ------------------------------------------------------------------ gráficas


def _grafica_torta(categorias) -> BytesIO:
    fig, ax = plt.subplots(figsize=(7, 4.2))
    valores = [v for _, v in categorias]
    ax.pie(valores, colors=PALETA[: len(valores)], autopct=lambda p: f"{p:.0f}%" if p >= 4 else "",
           startangle=90, pctdistance=0.78, wedgeprops={"width": 0.45, "edgecolor": "white"},
           textprops={"color": "white", "fontsize": 9, "weight": "bold"})
    ax.legend([f"{c} ({_pesos(v)})" for c, v in categorias], loc="center left",
              bbox_to_anchor=(1, 0.5), frameon=False, fontsize=9)
    ax.set_title("Gastos por categoría", fontsize=13, weight="bold", color="#1F4E9A")
    ax.axis("equal")
    return _a_png(fig)


def _grafica_meses(meses) -> BytesIO:
    fig, ax = plt.subplots(figsize=(7, 3.8))
    etiquetas = [_mes(m) for m, _ in meses]
    x = range(len(meses))
    ax.bar([i - 0.2 for i in x], [v[0] for _, v in meses], 0.4, label="Ingresos", color="#2E8B57")
    ax.bar([i + 0.2 for i in x], [v[1] for _, v in meses], 0.4, label="Gastos", color="#C0392B")
    ax.set_xticks(list(x), etiquetas)
    ax.yaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(lambda v, _: _pesos(v)))
    ax.spines[["top", "right"]].set_visible(False)
    ax.grid(axis="y", alpha=0.3)
    ax.legend(frameon=False)
    ax.set_title("Ingresos vs gastos por mes", fontsize=13, weight="bold", color="#1F4E9A")
    return _a_png(fig)


def _a_png(fig) -> BytesIO:
    buf = BytesIO()
    fig.tight_layout()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf


# --------------------------------------------------------------------- Excel


def _encabezado(ws, fila: int, titulos: list[str]) -> None:
    for col, t in enumerate(titulos, start=1):
        c = ws.cell(row=fila, column=col, value=t)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=AZUL)
        c.alignment = Alignment(horizontal="center")


def generar_excel(filas: list[tuple], periodo: str) -> BytesIO:
    d = analizar(filas)
    wb = Workbook()
    ws = wb.active
    ws.title = "Resumen"
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 16

    ws["A1"] = "MIRA – Finanzas Personales"
    ws["A1"].font = Font(bold=True, size=16, color=AZUL)
    ws["A2"] = f"Periodo: {periodo}"
    ws["A3"] = f"Generado: {datetime.now():%Y-%m-%d %H:%M}"
    ws["A2"].font = ws["A3"].font = Font(italic=True, color="666666")

    _encabezado(ws, 5, ["Concepto", "Valor"])
    ws["A6"], ws["B6"] = "Ingresos", d["ingresos"]
    ws["A7"], ws["B7"] = "Gastos", d["gastos"]
    ws["A8"], ws["B8"] = "Balance", "=B6-B7"
    ws["A9"], ws["B9"] = "% gastado", '=IF(B6>0,B7/B6,"")'
    for r in (6, 7, 8):
        ws[f"B{r}"].number_format = MONEDA
    ws["B9"].number_format = "0%"
    ws["A8"].font = ws["B8"].font = Font(bold=True)

    # Gastos por categoría + torta
    fila = 11
    _encabezado(ws, fila, ["Categoría", "Total", "%"])
    ini = fila + 1
    for i, (cat, total) in enumerate(d["categorias"]):
        r = ini + i
        ws.cell(row=r, column=1, value=cat)
        ws.cell(row=r, column=2, value=total).number_format = MONEDA
        ws.cell(row=r, column=3, value=f"=B{r}/SUM(B${ini}:B${ini + len(d['categorias']) - 1})").number_format = "0%"
    fin = ini + max(len(d["categorias"]), 1) - 1
    if d["categorias"]:
        torta = PieChart()
        torta.title = "Gastos por categoría"
        torta.add_data(Reference(ws, min_col=2, min_row=fila, max_row=fin), titles_from_data=True)
        torta.set_categories(Reference(ws, min_col=1, min_row=ini, max_row=fin))
        torta.dataLabels = DataLabelList()
        torta.dataLabels.showPercent = True
        torta.height, torta.width = 8, 13
        ws.add_chart(torta, "F2")
    else:
        ws.cell(row=ini, column=1, value="Sin gastos en el periodo")

    # Resumen mensual + barras
    fila_m = fin + 3
    _encabezado(ws, fila_m, ["Mes", "Ingresos", "Gastos", "Balance"])
    for i, (mes, (ing, gas)) in enumerate(d["meses"]):
        r = fila_m + 1 + i
        ws.cell(row=r, column=1, value=_mes(mes))
        for col, v in ((2, ing), (3, gas), (4, f"=B{r}-C{r}")):
            ws.cell(row=r, column=col, value=v).number_format = MONEDA
    if d["meses"]:
        barras = BarChart()
        barras.type = "col"
        barras.title = "Ingresos vs gastos por mes"
        fin_m = fila_m + len(d["meses"])
        barras.add_data(Reference(ws, min_col=2, max_col=3, min_row=fila_m, max_row=fin_m), titles_from_data=True)
        barras.set_categories(Reference(ws, min_col=1, min_row=fila_m + 1, max_row=fin_m))
        barras.series[0].graphicalProperties.solidFill = VERDE
        barras.series[1].graphicalProperties.solidFill = ROJO
        barras.y_axis.numFmt = MONEDA
        barras.height, barras.width = 8, 13
        ws.add_chart(barras, "F19")

    # Hoja de movimientos
    wm = wb.create_sheet("Movimientos")
    _encabezado(wm, 1, ["Fecha", "Tipo", "Categoría", "Descripción", "Monto"])
    for i, (fecha, tipo, cat, desc, monto) in enumerate(filas, start=2):
        wm.cell(row=i, column=1, value=datetime.fromisoformat(fecha)).number_format = "yyyy-mm-dd hh:mm"
        wm.cell(row=i, column=2, value=tipo.capitalize())
        wm.cell(row=i, column=3, value=cat or "Otros")
        wm.cell(row=i, column=4, value=desc)
        c = wm.cell(row=i, column=5, value=monto)
        c.number_format = MONEDA
        c.font = Font(color=ROJO if tipo == "gasto" else VERDE)
    for col, w in zip("ABCDE", (18, 10, 18, 32, 14)):
        wm.column_dimensions[col].width = w
    wm.freeze_panes = "A2"
    wm.auto_filter.ref = f"A1:E{max(len(filas) + 1, 2)}"

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ---------------------------------------------------------------------- Word


def _tabla(doc, encabezados: list[str], filas: list[list[str]]) -> None:
    t = doc.add_table(rows=1, cols=len(encabezados))
    t.style = "Light Grid Accent 1"
    for c, txt in zip(t.rows[0].cells, encabezados):
        c.text = txt
    for fila in filas:
        celdas = t.add_row().cells
        for c, txt in zip(celdas, fila):
            c.text = txt
    for row in t.rows:
        for c in row.cells:
            for p in c.paragraphs:
                for r in p.runs:
                    r.font.size = Pt(10)


def generar_word(filas: list[tuple], periodo: str) -> BytesIO:
    d = analizar(filas)
    doc = Document()
    if LOGO.exists():
        doc.add_picture(str(LOGO), width=Inches(1.4))
        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    titulo = doc.add_heading("Reporte de Finanzas Personales", 0)
    titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    info = doc.add_paragraph(f"Periodo: {periodo}  ·  Generado: {datetime.now():%Y-%m-%d %H:%M}")
    info.alignment = WD_ALIGN_PARAGRAPH.CENTER
    info.runs[0].font.color.rgb = RGBColor(0x66, 0x66, 0x66)

    doc.add_heading("Resumen", 1)
    pct = f"{d['gastos'] / d['ingresos']:.0%}" if d["ingresos"] else "—"
    _tabla(doc, ["Concepto", "Valor"], [
        ["Ingresos", _pesos(d["ingresos"])],
        ["Gastos", _pesos(d["gastos"])],
        ["Balance", ("-" if d["balance"] < 0 else "") + _pesos(abs(d["balance"]))],
        ["% de ingresos gastado", pct],
    ])
    consejo = doc.add_paragraph()
    if d["balance"] < 0:
        consejo.add_run("⚠ Gastaste más de lo que ingresaste. Revisa las categorías con mayor gasto.").bold = True
    elif d["ingresos"] and d["balance"] / d["ingresos"] >= 0.2:
        consejo.add_run("✔ ¡Muy bien! Ahorraste al menos el 20% de tus ingresos.").bold = True
    elif d["ingresos"]:
        consejo.add_run("Intenta ahorrar al menos el 20% de tus ingresos (regla 50/30/20).").bold = True

    doc.add_heading("Gastos por categoría", 1)
    if d["categorias"]:
        doc.add_picture(_grafica_torta(d["categorias"]), width=Inches(6))
        _tabla(doc, ["Categoría", "Total", "%"],
               [[c, _pesos(v), f"{v / d['gastos']:.0%}"] for c, v in d["categorias"]])
    else:
        doc.add_paragraph("Sin gastos en el periodo.")

    if d["meses"]:
        doc.add_heading("Evolución mensual", 1)
        doc.add_picture(_grafica_meses(d["meses"]), width=Inches(6))
        _tabla(doc, ["Mes", "Ingresos", "Gastos", "Balance"],
               [[_mes(m), _pesos(i), _pesos(g), ("-" if i - g < 0 else "") + _pesos(abs(i - g))]
                for m, (i, g) in d["meses"]])

    doc.add_heading("Movimientos", 1)
    ultimos = filas[-100:]
    if len(filas) > 100:
        doc.add_paragraph(f"Se muestran los últimos 100 de {len(filas)} movimientos.")
    _tabla(doc, ["Fecha", "Tipo", "Categoría", "Descripción", "Monto"],
           [[f[:16].replace("T", " "), t.capitalize(), c or "Otros", desc or "", _pesos(m)]
            for f, t, c, desc, m in ultimos])

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf
