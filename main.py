# Importamos una libreria para tratar archivos .xlsx y c

from openpyxl import load_workbook
import xlsxwriter

# Creamos unas variables que contienen el libro y la hoja que deseamos leer

file_path='Libro1.xlsx'
SHEET='Hoja1'

# Con el load_workbook abrimos el file_path solo como lectura

workbook=load_workbook(file_path,read_only=True)

# Luego abrimos la hoja que vamos a leer

sheet=workbook[SHEET]

# Quiero ver el contenido de la hoja para saber que valores estaré usando 

for row in sheet.iter_rows():
  print(row[0].value)
  print(row[1].value)
  print(row[2].value)

# Le asigno a cada celda una variable para poder usar su valor luego

a1=sheet['A1'].value
a2=sheet['A2'].value
a3=sheet['A3'].value

# Con estas variables vamos a crear una matriz

#m1=np.array([a1,a2,a3])

"""
# Creamos un libro en blanco

copia1=openpyxl.Workbook()

# Accedemos a la hoja en blanco que este archivo recien creado tiene

hoja1=copia1.active

# Le cambiamos el nombre a la hoja

hoja1.title='hoja1'

# Escribiremos en las celdas de esta ultima hoja creada
"""

workbook = xlsxwriter.Workbook("hoja1.xlsx")
worksheet = workbook.add_worksheet()

m1 = [[a1,a2,a3]]

row = 0

for col, data in enumerate(m1):
  worksheet.write_column(row, col, data)

workbook.close()


