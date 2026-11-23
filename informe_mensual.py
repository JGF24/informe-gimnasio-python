"""
INFORME MENSUAL AUTOMÁTICO DE UN GIMNASIO
=========================================
Lee el Excel de socios y en segundos genera:
  1. Un informe en Excel con resumen, gráficos, morosos, socios en riesgo de baja
     y movimientos del mes (altas y bajas).
  2. Un email de recordatorio de pago, ya redactado, para cada socio moroso.

Uso:
  python informe_mensual.py                      -> usa datos/socios.xlsx y el mes actual
  python informe_mensual.py --mes 2026-10        -> otro mes
  python informe_mensual.py --entrada otro.xlsx  -> otro archivo de socios
"""
import argparse
import calendar
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

# ---------------------------------------------------------------
# CONFIGURACIÓN (lo que un cliente cambiaría para su negocio)
# ---------------------------------------------------------------
PRECIOS = {"Básico": 29, "Premium": 45, "Élite": 79}
NOMBRE_NEGOCIO = "Forja Gym"
VISITAS_MINIMAS = 2  # Con menos visitas al mes, el socio está "en riesgo de baja"

COLOR_PRINCIPAL = "FF4D00"
COLOR_SUAVE = "FFE3D6"
BASE = Path(__file__).parent


# ---------------------------------------------------------------
# 1. LEER LOS DATOS
# ---------------------------------------------------------------
@dataclass
class Socio:
    id: int
    nombre: str
    apellidos: str
    email: str
    telefono: str
    plan: str
    alta: date
    baja: date | None
    ultimo_pago: date
    visitas: int

    @property
    def nombre_completo(self) -> str:
        return f"{self.nombre} {self.apellidos}"

    @property
    def cuota(self) -> int:
        return PRECIOS.get(self.plan, 0)


def a_fecha(valor) -> date | None:
    """Excel puede devolver datetime, date o texto: lo convertimos todo a date."""
    if valor is None or valor == "":
        return None
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, date):
        return valor
    return datetime.strptime(str(valor).strip(), "%d/%m/%Y").date()


def leer_socios(ruta: Path) -> list[Socio]:
    hoja = load_workbook(ruta, data_only=True).active
    socios = []
    for fila in hoja.iter_rows(min_row=2, values_only=True):
        if fila[0] is None:  # Fila vacía
            continue
        socios.append(Socio(
            id=int(fila[0]), nombre=fila[1], apellidos=fila[2], email=fila[3],
            telefono=str(fila[4]), plan=fila[5], alta=a_fecha(fila[6]),
            baja=a_fecha(fila[7]), ultimo_pago=a_fecha(fila[8]), visitas=int(fila[9] or 0),
        ))
    return socios


# ---------------------------------------------------------------
# 2. CALCULAR LOS INDICADORES DEL MES
# ---------------------------------------------------------------
def analizar(socios: list[Socio], anio: int, mes: int) -> dict:
    inicio = date(anio, mes, 1)
    fin = date(anio, mes, calendar.monthrange(anio, mes)[1])

    activos = [s for s in socios if s.alta <= fin and (s.baja is None or s.baja > fin)]
    altas = [s for s in socios if inicio <= s.alta <= fin]
    bajas = [s for s in socios if s.baja and inicio <= s.baja <= fin]
    morosos = [s for s in activos if s.ultimo_pago < inicio]
    en_riesgo = [s for s in activos if s.visitas < VISITAS_MINIMAS and s not in morosos]

    por_plan = {plan: sum(1 for s in activos if s.plan == plan) for plan in PRECIOS}
    ingresos_plan = {plan: n * PRECIOS[plan] for plan, n in por_plan.items()}

    return {
        "inicio": inicio, "fin": fin,
        "activos": activos, "altas": altas, "bajas": bajas,
        "morosos": sorted(morosos, key=lambda s: s.ultimo_pago),
        "en_riesgo": sorted(en_riesgo, key=lambda s: s.visitas),
        "por_plan": por_plan,
        "ingresos_plan": ingresos_plan,
        "ingresos_previstos": sum(ingresos_plan.values()),
        "deuda_total": sum(s.cuota for s in morosos),
        "visitas_media": round(sum(s.visitas for s in activos) / len(activos), 1) if activos else 0,
    }


# ---------------------------------------------------------------
# 3. CREAR EL EXCEL DEL INFORME
# ---------------------------------------------------------------
TITULO = Font(size=18, bold=True, color=COLOR_PRINCIPAL)
NEGRITA = Font(bold=True)
CABECERA_FONT = Font(bold=True, color="FFFFFF")
CABECERA_FILL = PatternFill("solid", fgColor=COLOR_PRINCIPAL)
FILA_ALTERNA = PatternFill("solid", fgColor="F6F6F6")
KPI_FILL = PatternFill("solid", fgColor=COLOR_SUAVE)
BORDE = Border(bottom=Side(style="thin", color="DDDDDD"))
EUROS = '#,##0 "€"'


def escribir_tabla(ws, fila_inicio: int, cabecera: list, filas: list, anchos: list, formatos: dict | None = None):
    """Escribe una tabla con cabecera de color, filas alternas y anchos de columna."""
    formatos = formatos or {}
    for col, texto in enumerate(cabecera, start=1):
        c = ws.cell(row=fila_inicio, column=col, value=texto)
        c.font, c.fill = CABECERA_FONT, CABECERA_FILL
        c.alignment = Alignment(horizontal="center")
    for i, datos in enumerate(filas, start=1):
        for col, valor in enumerate(datos, start=1):
            c = ws.cell(row=fila_inicio + i, column=col, value=valor)
            c.border = BORDE
            if i % 2 == 0:
                c.fill = FILA_ALTERNA
            if col in formatos:
                c.number_format = formatos[col]
    for col, ancho in enumerate(anchos, start=1):
        ws.column_dimensions[ws.cell(row=1, column=col).column_letter].width = ancho
    ws.freeze_panes = ws.cell(row=fila_inicio + 1, column=1)
    if filas:
        ultima = ws.cell(row=fila_inicio + len(filas), column=len(cabecera)).coordinate
        ws.auto_filter.ref = f"A{fila_inicio}:{ultima}"


def hoja_resumen(wb: Workbook, r: dict, nombre_mes: str):
    ws = wb.active
    ws.title = "Resumen"
    ws.sheet_view.showGridLines = False
    ws["A1"] = f"{NOMBRE_NEGOCIO} · Informe de {nombre_mes}"
    ws["A1"].font = TITULO
    ws["A2"] = f"Generado automáticamente el {date.today():%d/%m/%Y}"
    ws["A2"].font = Font(italic=True, color="888888")

    kpis = [
        ("Socios activos", len(r["activos"]), None),
        ("Altas del mes", len(r["altas"]), None),
        ("Bajas del mes", len(r["bajas"]), None),
        ("Ingresos previstos", r["ingresos_previstos"], EUROS),
        ("Socios morosos", len(r["morosos"]), None),
        ("Deuda pendiente", r["deuda_total"], EUROS),
        ("En riesgo de baja", len(r["en_riesgo"]), None),
        ("Visitas medias por socio", r["visitas_media"], None),
    ]
    for i, (nombre, valor, formato) in enumerate(kpis, start=4):
        ws.cell(row=i, column=1, value=nombre).font = NEGRITA
        c = ws.cell(row=i, column=2, value=valor)
        c.font = Font(bold=True, size=13)
        c.fill = KPI_FILL
        c.alignment = Alignment(horizontal="center")
        if formato:
            c.number_format = formato
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 16

    # Tabla de datos por plan (alimenta los gráficos)
    fila = 14
    for col, texto in enumerate(["Plan", "Socios", "Ingresos"], start=1):
        c = ws.cell(row=fila, column=col, value=texto)
        c.font, c.fill = CABECERA_FONT, CABECERA_FILL
    for i, plan in enumerate(PRECIOS, start=1):
        ws.cell(row=fila + i, column=1, value=plan)
        ws.cell(row=fila + i, column=2, value=r["por_plan"][plan])
        ws.cell(row=fila + i, column=3, value=r["ingresos_plan"][plan]).number_format = EUROS
    ws.column_dimensions["C"].width = 14

    categorias = Reference(ws, min_col=1, min_row=fila + 1, max_row=fila + len(PRECIOS))

    barras = BarChart()
    barras.title = "Socios activos por plan"
    barras.add_data(Reference(ws, min_col=2, min_row=fila, max_row=fila + len(PRECIOS)), titles_from_data=True)
    barras.set_categories(categorias)
    barras.legend = None
    barras.height, barras.width = 7.5, 13
    barras.series[0].graphicalProperties.solidFill = COLOR_PRINCIPAL
    barras.dataLabels = DataLabelList()
    barras.dataLabels.showVal = True
    ws.add_chart(barras, "E4")

    tarta = PieChart()
    tarta.title = "Reparto de ingresos"
    tarta.add_data(Reference(ws, min_col=3, min_row=fila, max_row=fila + len(PRECIOS)), titles_from_data=True)
    tarta.set_categories(categorias)
    tarta.height, tarta.width = 7.5, 13
    tarta.dataLabels = DataLabelList()
    tarta.dataLabels.showPercent = True
    ws.add_chart(tarta, "E20")


def hoja_morosos(wb: Workbook, r: dict):
    ws = wb.create_sheet("Morosos")
    filas = [[s.id, s.nombre_completo, s.email, s.telefono, s.plan, s.ultimo_pago,
              (r["inicio"] - s.ultimo_pago).days, s.cuota] for s in r["morosos"]]
    escribir_tabla(ws, 1,
                   ["ID", "Socio", "Email", "Teléfono", "Plan", "Último pago", "Días sin pagar", "Importe"],
                   filas, [6, 28, 36, 12, 10, 13, 15, 11], {6: "DD/MM/YYYY", 8: EUROS})
    total = len(filas) + 2
    ws.cell(row=total, column=7, value="TOTAL").font = NEGRITA
    c = ws.cell(row=total, column=8, value=f"=SUM(H2:H{total - 1})")
    c.font, c.number_format = NEGRITA, EUROS


def hoja_riesgo(wb: Workbook, r: dict):
    ws = wb.create_sheet("Riesgo de baja")
    filas = [[s.id, s.nombre_completo, s.email, s.telefono, s.plan, s.visitas, s.alta] for s in r["en_riesgo"]]
    escribir_tabla(ws, 1, ["ID", "Socio", "Email", "Teléfono", "Plan", "Visitas este mes", "Socio desde"],
                   filas, [6, 28, 36, 12, 10, 16, 13], {7: "DD/MM/YYYY"})


def hoja_movimientos(wb: Workbook, r: dict):
    ws = wb.create_sheet("Altas y bajas")
    filas = ([["Alta", s.alta, s.nombre_completo, s.plan] for s in r["altas"]] +
             [["Baja", s.baja, s.nombre_completo, s.plan] for s in r["bajas"]])
    filas.sort(key=lambda f: f[1])
    escribir_tabla(ws, 1, ["Tipo", "Fecha", "Socio", "Plan"], filas, [8, 13, 28, 10], {2: "DD/MM/YYYY"})
    for fila in ws.iter_rows(min_row=2, max_col=1):
        fila[0].font = Font(bold=True, color="1E9E5A" if fila[0].value == "Alta" else "D0312D")


# ---------------------------------------------------------------
# 4. REDACTAR LOS EMAILS DE RECORDATORIO
# ---------------------------------------------------------------
PLANTILLA_EMAIL = """Para: {email}
Asunto: Recordatorio de pago de tu cuota de {mes} – {negocio}

Hola, {nombre}:

Te escribimos porque todavía no hemos recibido el pago de tu cuota del plan {plan} ({importe} €) correspondiente a {mes}.

Puedes pagarla en recepción o desde la app de socio. Si ya lo has hecho, ignora este mensaje.

¡Te esperamos en el gimnasio!

El equipo de {negocio}
"""


def generar_emails(r: dict, carpeta: Path, nombre_mes: str) -> int:
    carpeta.mkdir(parents=True, exist_ok=True)
    for s in r["morosos"]:
        texto = PLANTILLA_EMAIL.format(email=s.email, nombre=s.nombre, plan=s.plan, importe=s.cuota,
                                       mes=nombre_mes, negocio=NOMBRE_NEGOCIO)
        (carpeta / f"{s.id:04d}_{s.nombre}_{s.apellidos.split()[0]}.txt").write_text(texto, encoding="utf-8")
    return len(r["morosos"])


# ---------------------------------------------------------------
# PROGRAMA PRINCIPAL
# ---------------------------------------------------------------
MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
         "agosto", "septiembre", "octubre", "noviembre", "diciembre"]


def main():
    parser = argparse.ArgumentParser(description="Genera el informe mensual del gimnasio.")
    parser.add_argument("--entrada", default=BASE / "datos" / "socios.xlsx", type=Path)
    parser.add_argument("--mes", default=date.today().strftime("%Y-%m"), help="Formato AAAA-MM")
    args = parser.parse_args()

    anio, mes = map(int, args.mes.split("-"))
    nombre_mes = f"{MESES[mes - 1]} de {anio}"

    print(f"📂 Leyendo {args.entrada.name}...")
    socios = leer_socios(args.entrada)
    r = analizar(socios, anio, mes)

    salida = BASE / "salida"
    salida.mkdir(exist_ok=True)
    wb = Workbook()
    hoja_resumen(wb, r, nombre_mes)
    hoja_morosos(wb, r)
    hoja_riesgo(wb, r)
    hoja_movimientos(wb, r)
    ruta_informe = salida / f"informe_{anio}-{mes:02d}.xlsx"
    wb.save(ruta_informe)

    n_emails = generar_emails(r, salida / f"emails_{anio}-{mes:02d}", nombre_mes)

    print(f"""
✅ Informe de {nombre_mes} generado
   Socios activos ....... {len(r['activos'])}
   Altas / bajas ........ {len(r['altas'])} / {len(r['bajas'])}
   Ingresos previstos ... {r['ingresos_previstos']:,} €
   Morosos .............. {len(r['morosos'])} (deuda: {r['deuda_total']:,} €)
   En riesgo de baja .... {len(r['en_riesgo'])}

📊 Excel:  {ruta_informe}
✉️  {n_emails} emails de recordatorio en: {salida / f'emails_{anio}-{mes:02d}'}""".replace(",", "."))


if __name__ == "__main__":
    main()
