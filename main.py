from openpyxl import Workbook
from openpyxl.styles import *
from openpyxl.drawing.image import Image
import datetime

class Presupuesto():

    libro = Workbook()
    hoja1 = libro.active

    LOGO_LONIDA = Image("images\logo.png")
    DIRECCION_NEGOCIO = "ALBERT EINSTEIN 502-SALTA"
    CELULAR = "CEL 3875381011"
    BORDE_TABLA = Border(left=Side(border_style='thin', color='000000'), right=Side(border_style='thin', color='000000'), top=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'))
    
    
    hoja1.row_dimensions[2].height = 95
    hoja1.row_dimensions[6].height = 20
    hoja1.column_dimensions["A"].width = 15
    hoja1.column_dimensions["B"].width = 55
    hoja1.column_dimensions["C"].width = 13
    hoja1.column_dimensions["D"].width = 20
    hoja1["A6"].font = Font(size=20)
    hoja1["A8"].font = Font(size=14)
    hoja1["A9"].font = Font(size=9, bold=True, italic=True)
    hoja1["B9"].font = Font(size=9, bold=True, italic=True)
    hoja1["C9"].font = Font(size=9, bold=True, italic=True)
    hoja1["D9"].font = Font(size=9, bold=True, italic=True)
    hoja1["A9"].border = BORDE_TABLA
    hoja1["B9"].border = BORDE_TABLA
    hoja1["C9"].border = BORDE_TABLA
    hoja1["D9"].border = BORDE_TABLA
    for i in range(2,9):
        hoja1.merge_cells(f"A{i}:B{i}")
    fecha = datetime.datetime.now()
    fecha_actual = fecha.strftime("%d/%m/%Y")

    hoja1.add_image(LOGO_LONIDA, "A2")
    hoja1["D5"] = "Fecha: " + fecha_actual
    hoja1["A6"] = "P R E S U P U E S T O"
    hoja1["A4"] = DIRECCION_NEGOCIO
    hoja1["A5"] = CELULAR
    hoja1["A7"] = "SRES."
    hoja1["A9"] = "CANTIDAD"
    hoja1["B9"] = "ACCESORIOS"
    hoja1["C9"] = "PRECIO"
    hoja1["D9"] = "TOTAL"
    
    def __init__(self, vendedor, cliente):
        self.hoja1["A3"] = "DE " + vendedor
        self.hoja1["A8"] = cliente
        self.libro.save("A.xlsx")

    def formato_de_tabla_accesorios(self, fila=10):
        self.hoja1[f"A{fila}"].alignment = Alignment(horizontal='center', vertical='center')
        self.hoja1[f"C{fila}"].alignment = Alignment(horizontal='right', vertical='center')
        self.hoja1[f"A{fila}"].border = self.BORDE_TABLA
        self.hoja1[f"B{fila}"].border = self.BORDE_TABLA
        self.hoja1[f"C{fila}"].border = self.BORDE_TABLA
        self.hoja1[f"D{fila}"].border = self.BORDE_TABLA
        self.hoja1[f"A{fila}"].font = Font(size=9, italic=True)
        self.hoja1[f"B{fila}"].font = Font(size=9, italic=True)
        self.hoja1[f"C{fila}"].font = Font(size=9, italic=True)
        self.hoja1[f"D{fila}"].font = Font(size=9, italic=True)
        self.libro.save("A.xlsx")

    def colocar_accesorio(self, fila, cantidad, accesorio, precio):
        self.hoja1[f"A{fila}"] = cantidad
        self.hoja1[f"B{fila}"] = accesorio
        self.hoja1[f"C{fila}"] = precio
        self.hoja1[f"D{fila}"] = f"=A{fila}*C{fila}"
        self.libro.save("A.xlsx")
    
    def formato_total_pago(self, fila, contado):
        self.hoja1[f"C{fila}"].border = self.BORDE_TABLA
        self.hoja1[f"C{fila}"].font = Font(size=14, bold=True)
        self.hoja1[f"C{fila}"] = "TOTAL"
        self.hoja1[f"D{fila}"].border = self.BORDE_TABLA
        self.hoja1[f"D{fila}"].font = Font(size=14, bold=True)
        self.hoja1[f"D{fila}"] = f"=SUM(D10:D{fila-1})"
        self.hoja1[f"B{fila+2}"].font = Font(size=14, bold=True)
        self.hoja1[f"B{fila+2}"]= contado
        self.libro.save("A.xlsx")
    
    libro.save("A.xlsx")


print("BINEVENIDO A LONXCEL, una aplicación para hacer presupuestos LONIDA")
print()
nombre_vendedor = input("Ingrese el nombre del vendedor que realiza el presupuesto: ").upper()
nombre_cliente = input("¿Para quién desea realizar el presupuesto? ").upper()
apunte = Presupuesto(nombre_vendedor, nombre_cliente)
numero_de_accesorio = 0
fila = 9
total = 0
while True:
    fila = fila + 1 
    numero_de_accesorio = numero_de_accesorio + 1
    print(f"Accesorio número {numero_de_accesorio}")
    accesorio = input(f"Ingrese el accesorio (para dejar de ingresar coloque 'NINGUNO'): ").upper()
    if accesorio == "NINGUNO":
        break
    cantidad_accesorio = input("Ingrese la cantidad del accesorio: ")
    precio = input("Ingrese el precio del producto: ")
    precio_pesos =  "$ " + precio
    
    cantidad_accesorio = int(cantidad_accesorio)
    precio = int(precio)
    total = total + (cantidad_accesorio * precio)

    apunte.formato_de_tabla_accesorios(fila)
    apunte.colocar_accesorio(fila, cantidad_accesorio, accesorio, precio)
    

print(f"El total es {total}")
total_contado = input("Ingrese cuánto sería el pago total de contado: ")
pago_contado = f"PAGO CONTADO $ {total_contado}"
apunte.formato_total_pago(fila, pago_contado)












