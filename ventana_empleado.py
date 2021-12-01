from tkinter import *
from tkinter.ttk import *

from numpy import integer
from ventana_empleado import *
from ver_errores_persona import *
import tkinter as tk
from functools import *
from analisis import *


def call_result(caja_de_texto,_nombre_persona, mañana, tarde, noche):  
    num1 = int(mañana)
    num2 = int(tarde)
    num3 = int(noche)

    horasyminutos = calcular_horas_trabajadas(_nombre_persona)
    sueldo = calcular_sueldo(horasyminutos,_nombre_persona,num1,num2,num3)
    
    caja_de_texto.delete(0,tk.END)
    caja_de_texto.insert(tk.END, f"Sueldo calculado: {sueldo} ARS") 

def crear_ventana_empleado(_nombre_persona):
    if((_nombre_persona != '') or (_nombre_persona != 'Seleccione a un empleado')):

        #ventana 
        ventana= Tk()
        ventana.title(f"Empleado {_nombre_persona}")
        ventana.geometry("1235x500")

        #Frames
        frame1=Frame(ventana)
        frame1.grid(row=1)

        frame2=Frame(ventana)
        frame2.grid(row=2)

        frame3=Frame(ventana)
        frame2.grid(row=3)

        #Elementos
        titulo = Label(frame1, text=f"Calculo de sueldo de {_nombre_persona}", font=("Arial",19),) #Titulo esta contenido en el frame principal
        titulo.grid(column=0,row=0,padx=10,pady=10,)

        #Tabla con los horarios que no son erroneos
        tabla = ttk.Treeview(frame1 , height=10)
        tabla.grid(column=0, row=1, sticky='nsew')

        ladox = Scrollbar(frame1, orient = HORIZONTAL, command= tabla.xview)
        ladox.grid(column=0, row = 2, sticky='ew') 

        ladoy = Scrollbar(frame1, orient =VERTICAL, command = tabla.yview)
        ladoy.grid(column = 1, row = 1, sticky='ns')

        tabla.configure(xscrollcommand = ladox.set, yscrollcommand = ladoy.set)

        # Busca el path de la carpeta raiz
        path_raiz = os.getcwd()
        ruta= f'{path_raiz}\\ExcelsPersonas\\{_nombre_persona}\\{_nombre_persona}_sin_errores.xlsx'


        #Carga el contenido del excel en la tabla
        df=pd.read_excel(ruta)

        tabla['column'] = list(df.columns)
        tabla['show'] = "headings"  #encabezado
        
        for columna in tabla['column']:
            tabla.heading(columna, text= columna)
        
        df_fila = df.to_numpy().tolist()
        for fila in df_fila:
            tabla.insert('', 'end', values =fila)

        #Labels y campos de texto para llenar con los precios de las horas
        Label(frame2, text="Precio p/hora mañana:").grid(row=0, column=0)
        entrada_mañana = tk.Entry(frame2, width=40)
        entrada_mañana.grid(padx=3, row=0, column=1)

        Label(frame2, text="Precio p/hora tarde:").grid(row=1, column=0)
        entrada_tarde = tk.Entry(frame2, width=40)
        entrada_tarde.grid(padx=3, row=1, column=1)

        Label(frame2, text="Precio p/hora noche:").grid(row=2, column=0)
        entrada_noche = tk.Entry(frame2, width=40)
        entrada_noche.grid(padx=3, row=2, column=1)

        indica = Label(frame3, fg= 'white', bg='gray26', text=ruta , font= ('Arial',10,'bold') )
        indica.grid(column=0, row = 5)

        caja_de_texto = Listbox(frame2,width=50)
        caja_de_texto.grid(padx=20, column=2, row=0, rowspan=3, columnspan=4)

        #call_result = partial(caja_de_texto, precio_m,precio_t,precio_n)

        boton_calcular = Button(frame2, text="Calcular", command=lambda:call_result(caja_de_texto,_nombre_persona,entrada_mañana.get(),entrada_tarde.get(),entrada_noche.get()))
        boton_calcular.grid(padx=5,pady=3, row=3, column=1)

        ventana=mainloop()
    else:
        print("No se imprimio la ventana porque esta vacio el dato2")


if __name__=="__main__":
    crear_ventana_empleado('ALMARAZ ANGELICA')
    
