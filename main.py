from openpyxl import Workbook
from openpyxl.styles import *
import openpyxl.drawing.image as imagen
from tkinter import *
from tkinter import filedialog
import subprocess as sb
from tkinter import messagebox
import datetime


class Presupuesto():

    libro = Workbook()
    hoja1 = libro.active

    LOGO_LONIDA = imagen.Image("images\logo.png")
    DIRECCION_NEGOCIO = "ALBERT EINSTEIN 502-SALTA"
    CELULAR = "CEL 3875381011"
    BORDE_TABLA = Border(left=Side(border_style='thin', color='000000'),
                         right=Side(border_style='thin', color='000000'),
                         top=Side(border_style='thin', color='000000'),
                         bottom=Side(border_style='thin', color='000000')
    )
    
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
    
    def __init__(self, vendedor="DANIEL ANTONIO MEJIA", cliente="JOAQUIN SAJAMA"):
        self.hoja1["A3"] = "DE " + vendedor
        self.hoja1["A8"] = cliente
        self.nombre_presupuesto = "PRESUPUESTO " + cliente
        self.libro.save(f"{self.nombre_presupuesto}.xlsx")

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
        self.libro.save(f"{self.nombre_presupuesto}.xlsx")

    def colocar_accesorio(self, fila, cantidad, accesorio, precio):
        self.hoja1[f"A{fila}"] = cantidad
        self.hoja1[f"B{fila}"] = accesorio
        self.hoja1[f"C{fila}"] = precio
        self.hoja1[f"D{fila}"] = f"=A{fila}*C{fila}"
        self.libro.save(f"{self.nombre_presupuesto}.xlsx")
    
    def formato_total_pago(self, fila):
        self.hoja1[f"C{fila}"].border = self.BORDE_TABLA
        self.hoja1[f"D{fila}"].border = self.BORDE_TABLA
        self.hoja1[f"C{fila}"].font = Font(size=14, bold=True)
        self.hoja1[f"D{fila}"].font = Font(size=14, bold=True)
        self.hoja1[f"C{fila}"] = "TOTAL"
        self.hoja1[f"D{fila}"] = f"=SUM(D10:D{fila-1})"
        self.libro.save(f"{self.nombre_presupuesto}.xlsx")


def mostrar_entry():

    def mostrar_entradas(event):

        def abrir_archivo():
            archivo = filedialog.askopenfilename(
                title="Seleccionar archivo Excel",
                filetypes=[("Archivos Excel", "*.xlsx")]
            )
            if archivo:
                try:
                    sb.run(["start", archivo], shell=True)  # Abrir con app predeterminada
                    messagebox.showinfo("Éxito", "Archivo abierto correctamente")
                except Exception as e:
                    messagebox.showerror("Error", f"No se pudo abrir el archivo:\n{str(e)}")
                    
        global fila
        global aux
        global aux1
        global lista_aux
        global lista_aux1
        global fila
        aux = aux + 1
    
        texto = entrada.get()
        texto = texto.upper()
    
        if aux == 1:
            lista_aux.append(texto)
            etiqueta_indicaciones.config(text="Ingrese el nombre del cliente")
        elif aux == 2:
            lista_aux.append(texto)
            etiqueta_indicaciones.config(text="Ingrese el accesorio (para dejar de ingresar coloque 'Ninguno')")
        elif aux >= 3:
            aux1 = aux1 + 1
            prespuesto = Presupuesto(lista_aux[0],lista_aux[1])
            
            if aux1 == 1 and texto != "NINGUNO":
                lista_aux1.append(texto)
                etiqueta_indicaciones.config(text="Ingrese la cantidad del accesorio")
            elif aux1 == 2:
                lista_aux1.append(texto)
                etiqueta_indicaciones.config(text="Ingrese el precio del producto")
            elif aux1 == 3:
                lista_aux1.append(texto)
                fila = fila + 1 
                prespuesto.formato_de_tabla_accesorios(fila)
                prespuesto.colocar_accesorio(fila, lista_aux1[1], lista_aux1[0], lista_aux1[2])
                aux1 = 0
                lista_aux1 = []
                etiqueta_indicaciones.config(text="Ingrese el accesorio (para dejar de ingresar coloque 'Ninguno')")
            else:
                prespuesto.formato_total_pago(fila+1)
                etiqueta_indicaciones.config(text="¡Tu presupuesto ya está listo, hecha un vistazo!")
                boton_abrir = Button(ventana, text="Abrir Archivo",font="Georgia 12",bg="#2794C2",command=abrir_archivo)
                boton_abrir.pack(pady=30)
              
        etiqueta_mostrar_entrada.config(text=f"Ingresaste: {texto}")
        entrada.delete(0,END)
    
    etiqueta_indicaciones = Label(ventana)
    etiqueta_indicaciones = Label(ventana,text="Ingrese el nombre del vendedor, que realizará el presupuesto",font="Georgia 14",bg="#92D8F5")
    etiqueta_indicaciones.pack(pady=15)

    entrada = Entry(ventana,font= "Georgia 14",bg="#AFE1F6")
    entrada.config(width=40)
    entrada.pack(pady=20)

    etiqueta_mostrar_entrada = Label(ventana,text="",font="Georgia 14",bg="#92D8F5")
    etiqueta_mostrar_entrada.pack(pady=25)
    
    entrada.bind("<Return>", lambda event: mostrar_entradas(event))
    entrada.focus_set()
    

fila = 9
aux = 0
aux1 = 0
lista_aux =[]
lista_aux1 = []

ventana = Tk()

ancho_pantalla = ventana.winfo_screenwidth()
alto_pantalla = ventana.winfo_screenheight()
x = (ancho_pantalla - 900) // 2
y = (alto_pantalla - 750) // 2
ventana.geometry(f"{900}x{750}+{x}+{y}")

ventana.title("LONXCEL")
ventana.iconbitmap("images\logo.ico")
ventana.geometry("900x700")
ventana.config(bg="#92D8F5")

etiqueta_bienvenida = Label(ventana,text="¡Bienvenido a Lonxcel!",font="Georgia 18",bg="#4ABAE7")
etiqueta_bienvenida.config(height=2,width=20)
etiqueta_bienvenida.pack(pady=10)

etiqueta_aviso = Label(ventana, text="Para realizar un presupuesto, pulse 'Comenzar'",font="Georgia 14",bg="#92D8F5" )
etiqueta_aviso.pack(pady=12)

boton_comenzar = Button(ventana,text="Comenzar",font="Georgia 12",bg="#2794C2",command=mostrar_entry)
boton_comenzar.config(height=1,width=10)
boton_comenzar.pack(pady=15)

ventana.mainloop()








