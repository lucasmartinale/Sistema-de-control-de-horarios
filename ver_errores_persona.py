from tkinter import Tk, Label, Button, Frame,  messagebox, filedialog, ttk, Scrollbar, VERTICAL, HORIZONTAL
import pandas as pd
import os

def ventana_errores(_nombre_persona):
    if((_nombre_persona != '') or (_nombre_persona != 'Seleccione a un empleado')):

        #Se crea la ventana
        ventana = Tk()
        ventana.config(bg='black')
        ventana.geometry('600x400')
        ventana.minsize(width=600, height=400)
        ventana.title(f'Errores de horario de {_nombre_persona}')


        ventana.columnconfigure(0, weight = 25)
        ventana.rowconfigure(0, weight= 25)
        ventana.columnconfigure(0, weight = 1)
        ventana.rowconfigure(1, weight= 1)

        #Crea los frames
        frame1 = Frame(ventana)
        frame1.grid(column=0,row=0,sticky='nsew')
        frame2 = Frame(ventana)
        frame2.grid(column=0,row=1,sticky='nsew')

        frame1.columnconfigure(0, weight = 1)
        frame1.rowconfigure(0, weight= 1)

        frame2.columnconfigure(0, weight = 1)
        frame2.rowconfigure(0, weight= 1)
        frame2.columnconfigure(1, weight = 1)
        frame2.rowconfigure(0, weight= 1)

        frame2.columnconfigure(2, weight = 1)
        frame2.rowconfigure(0, weight= 1)

        frame2.columnconfigure(3, weight = 2)
        frame2.rowconfigure(0, weight= 1)


        tabla = ttk.Treeview(frame1 , height=10)
        tabla.grid(column=0, row=0, sticky='nsew')

        ladox = Scrollbar(frame1, orient = HORIZONTAL, command= tabla.xview)
        ladox.grid(column=0, row = 1, sticky='ew') 

        ladoy = Scrollbar(frame1, orient =VERTICAL, command = tabla.yview)
        ladoy.grid(column = 1, row = 0, sticky='ns')

        tabla.configure(xscrollcommand = ladox.set, yscrollcommand = ladoy.set)

        # Busca el path de la carpeta raiz
        path_raiz = os.getcwd()
        ruta= f'{path_raiz}\\ExcelsPersonas\\{_nombre_persona}\\{_nombre_persona}_horarios.xlsx'

        indica = Label(frame2, fg= 'white', bg='gray26', text=ruta , font= ('Arial',10,'bold') )
        indica.grid(column=3, row = 0)
        
        df=pd.read_excel(ruta)

        tabla['column'] = list(df.columns)
        tabla['show'] = "headings"  #encabezado
        
        for columna in tabla['column']:
            tabla.heading(columna, text= columna)
        
        df_fila = df.to_numpy().tolist()
        for fila in df_fila:
            tabla.insert('', 'end', values =fila)

        ventana.mainloop()
    else:
        print("No se imprimio la ventana porque esta vacio el dato")

if __name__=="__main__":
    ventana_errores("ALMARAZ ANGELICA")
    