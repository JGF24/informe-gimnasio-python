"""
Genera un Excel de ejemplo con socios ficticios de un gimnasio.

Sirve para probar la automatización sin datos reales de clientes.
Uso:  python generar_datos.py
"""
import random
from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

random.seed(42)  # Misma semilla = mismos datos cada vez

NOMBRES = ["Lucía", "Hugo", "Martina", "Mateo", "Sofía", "Leo", "Julia", "Daniel", "Paula", "Pablo",
           "Valeria", "Álvaro", "Carla", "Adrián", "Sara", "Manuel", "Elena", "Diego", "Noa", "Javier",
           "Irene", "Marcos", "Claudia", "Sergio", "Alba", "Iván", "Marta", "Rubén", "Laura", "Carlos"]
APELLIDOS = ["García", "Rodríguez", "González", "Fernández", "López", "Martínez", "Sánchez", "Pérez",
             "Gómez", "Martín", "Jiménez", "Ruiz", "Hernández", "Díaz", "Moreno", "Muñoz", "Álvarez",
             "Romero", "Alonso", "Navarro"]
PLANES = ["Básico", "Premium", "Élite"]
PESOS_PLANES = [45, 40, 15]  # % aproximado de socios en cada plan


def fecha_aleatoria(desde: date, hasta: date) -> date:
    return desde + timedelta(days=random.randint(0, (hasta - desde).days))


def generar(num_socios: int = 180) -> Path:
    hoy = date.today()
    inicio_mes = hoy.replace(day=1)

    wb = Workbook()
    ws = wb.active
    ws.title = "Socios"
    cabecera = ["ID", "Nombre", "Apellidos", "Email", "Teléfono", "Plan",
                "Fecha alta", "Fecha baja", "Último pago", "Visitas este mes"]
    ws.append(cabecera)

    for i in range(1, num_socios + 1):
        nombre = random.choice(NOMBRES)
        apellidos = f"{random.choice(APELLIDOS)} {random.choice(APELLIDOS)}"
        email = f"{nombre.lower()}.{apellidos.split()[0].lower()}{i}@example.com"
        email = (email.replace("á", "a").replace("é", "e").replace("í", "i")
                      .replace("ó", "o").replace("ú", "u"))
        telefono = f"6{random.randint(10_000_000, 99_999_999)}"
        plan = random.choices(PLANES, weights=PESOS_PLANES)[0]

        # 15 % de altas nuevas este mes, el resto en los últimos 2 años
        if random.random() < 0.15:
            alta = fecha_aleatoria(inicio_mes, hoy)
        else:
            alta = fecha_aleatoria(hoy - timedelta(days=730), inicio_mes - timedelta(days=1))

        # 10 % se ha dado de baja (algunas este mes)
        baja = None
        if random.random() < 0.10:
            baja = fecha_aleatoria(max(alta, hoy - timedelta(days=90)), hoy)

        # 12 % no ha pagado este mes (morosos)
        if random.random() < 0.12:
            ultimo_pago = fecha_aleatoria(inicio_mes - timedelta(days=60), inicio_mes - timedelta(days=1))
        else:
            desde = max(alta, inicio_mes)
            ultimo_pago = fecha_aleatoria(desde, min(desde + timedelta(days=5), hoy))

        # Visitas: algunos socios casi no vienen (riesgo de baja)
        visitas = 0 if baja else random.choices([0, 1, 2, random.randint(3, 22)], weights=[8, 6, 6, 80])[0]

        ws.append([i, nombre, apellidos, email, telefono, plan, alta, baja, ultimo_pago, visitas])

    # Un poco de formato para que se vea bien al abrirlo
    for celda in ws[1]:
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = PatternFill("solid", fgColor="FF4D00")
    for col, ancho in zip("ABCDEFGHIJ", [6, 12, 24, 38, 12, 10, 12, 12, 12, 16]):
        ws.column_dimensions[col].width = ancho
    for fila in ws.iter_rows(min_row=2, min_col=7, max_col=9):
        for celda in fila:
            celda.number_format = "DD/MM/YYYY"

    salida = Path(__file__).parent / "datos" / "socios.xlsx"
    salida.parent.mkdir(exist_ok=True)
    wb.save(salida)
    return salida


if __name__ == "__main__":
    ruta = generar()
    print(f"✔ Datos de ejemplo creados en: {ruta}")
