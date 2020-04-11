import os
from openpyxl import load_workbook

def crearcarpetas(wb):
  grupos = wb.sheetnames
  for hoja in grupos:
    grupo = wb[hoja]
    carpeta = grupo.title

    try:
      os.stat(carpeta)
    except:
      os.mkdir(carpeta)

    for fila in grupo.rows:
      for columna in fila:
        nombre = columna.value
        nombre = nombre.strip()
        subcarpeta = carpeta + os.sep + nombre

        try:
          os.stat(subcarpeta)
        except:
          os.mkdir(subcarpeta)

if __name__ == "__main__":
  wb = load_workbook(filename='nombres.xlsx', read_only=True)
  crearcarpetas(wb)