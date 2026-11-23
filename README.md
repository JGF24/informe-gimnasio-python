# 📊 Informe mensual automático para gimnasios

Script en **Python** que convierte el Excel de socios de un gimnasio en un **informe mensual completo en segundos**, con trabajo que normalmente lleva horas hacer a mano.

## ¿Qué hace?
A partir de un único Excel con los socios, genera automáticamente:

| Resultado | Descripción |
|---|---|
| **Hoja "Resumen"** | Socios activos, altas, bajas, ingresos previstos, deuda pendiente… con **gráficos** de socios por plan y reparto de ingresos |
| **Hoja "Morosos"** | Socios que no han pagado este mes, días de retraso e importe total adeudado |
| **Hoja "Riesgo de baja"** | Socios que apenas vienen al gimnasio (para contactarlos antes de que se den de baja) |
| **Hoja "Altas y bajas"** | Movimientos del mes ordenados por fecha |
| **Emails de recordatorio** | Un email personalizado y ya redactado para cada socio moroso |

## Cómo usarlo
```bash
pip install -r requirements.txt
python generar_datos.py          # crea datos de ejemplo (opcional)
python informe_mensual.py        # genera el informe del mes actual
python informe_mensual.py --mes 2026-10   # informe de otro mes
```
Los resultados aparecen en la carpeta `salida/`.

## Adaptable a cualquier negocio
Precios, nombre del negocio, umbral de "riesgo de baja" y textos de los emails se cambian en la sección **CONFIGURACIÓN** al principio de `informe_mensual.py`. La misma idea sirve para academias, clínicas, coworkings o cualquier negocio con clientes que pagan una cuota.

**Ampliaciones posibles:** envío real de los emails (SMTP/Gmail), lectura desde Google Sheets, ejecución automática cada día 1 de mes, o exportar el informe a PDF.

## Tecnologías
Python 3 · openpyxl (Excel y gráficos) · argparse

> Los datos de ejemplo son ficticios y se generan con `generar_datos.py`.
