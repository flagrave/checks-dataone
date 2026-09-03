"""
Pone =SUM(Jan..Dec) en la columna `Full Year` de la hoja PreFC Y1.

Usa Excel via COM (no openpyxl) porque el libro tiene un pivot table, la
tabla `TablaSBP` y las partes del add-in de SAP Analysis: openpyxl las
tiraria al reescribir el xlsx.

Se saltean las filas que tienen Full Year cargado y los doce meses en cero
(las de Gross Profit sin fraseo mensual): ahi el =SUM las dejaria en cero y
se perderia el dato.
"""

import os
import sys

import win32com.client as win32

BOOK = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Checksibf.xlsx")
SHEET = "PreFC Y1"
FIRST_ROW, LAST_ROW = 3, 534
COL_JAN, COL_DEC, COL_FY = 14, 25, 26  # N, Y, Z
FORMULA = "=SUM(TablaSBP[@[Jan]:[Dec]])"

xl = win32.DispatchEx("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
autofill_prev = xl.AutoCorrect.AutoFillFormulasInLists
# Sin esto, la tabla propaga la formula a TODA la columna y pisa las filas
# que queremos dejar como valor.
xl.AutoCorrect.AutoFillFormulasInLists = False

wb = None
try:
    wb = xl.Workbooks.Open(BOOK, UpdateLinks=0)
    ws = wb.Worksheets(SHEET)

    # Snapshot: valor previo de Full Year y suma de meses, por fila.
    before = {}
    for r in range(FIRST_ROW, LAST_ROW + 1):
        months = [ws.Cells(r, c).Value for c in range(COL_JAN, COL_DEC + 1)]
        months = [m if isinstance(m, (int, float)) else 0 for m in months]
        fy = ws.Cells(r, COL_FY).Value
        before[r] = (fy if isinstance(fy, (int, float)) else 0, sum(months), any(months))

    escritas, salteadas = 0, []
    for r in range(FIRST_ROW, LAST_ROW + 1):
        fy_prev, suma, tiene_meses = before[r]
        if not tiene_meses and fy_prev != 0:
            salteadas.append((r, fy_prev))
            continue
        ws.Cells(r, COL_FY).Formula = FORMULA
        escritas += 1

    xl.CalculateFullRebuild()

    # Verificacion: la formula tiene que dar la suma de meses.
    malas = []
    cambiadas = []
    for r in range(FIRST_ROW, LAST_ROW + 1):
        if any(r == s[0] for s in salteadas):
            continue
        fy_prev, suma, _ = before[r]
        ahora = ws.Cells(r, COL_FY).Value
        ahora = ahora if isinstance(ahora, (int, float)) else 0
        if abs(ahora - suma) > 0.005:
            malas.append((r, suma, ahora))
        if abs(ahora - fy_prev) > 0.005:
            cambiadas.append((r, fy_prev, ahora))

    if malas:
        print("ABORTADO: la formula no da la suma de meses en estas filas:")
        for m in malas[:20]:
            print("   fila %d  esperado %.2f  obtenido %.2f" % m)
        wb.Close(SaveChanges=False)
        wb = None
        sys.exit(1)

    wb.Save()
    wb.Close(SaveChanges=False)
    wb = None

    print("Formula puesta: %s" % FORMULA)
    print("Filas con =SUM : %d" % escritas)
    print("Filas salteadas: %d  (Full Year cargado, meses en cero)" % len(salteadas))
    for r, v in salteadas:
        print("   fila %-4d Full Year %.2f  -> queda como valor" % (r, v))
    print()
    print("Filas donde el total CAMBIO: %d" % len(cambiadas))
    for r, antes, ahora in cambiadas:
        print("   fila %-4d %14.2f -> %14.2f   (delta %+.2f)" % (r, antes, ahora, ahora - antes))

finally:
    if wb is not None:
        wb.Close(SaveChanges=False)
    xl.AutoCorrect.AutoFillFormulasInLists = autofill_prev
    xl.Quit()
