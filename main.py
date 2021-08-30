from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd

import tkinter as tk
from tkinter import ttk

def limpiezaExcel(df):
    df=df.drop(columns=['Num', 'Department','Verifycode','Device ID','Device Name','UserExtFmt'])
    
    #Extraer fecha y hora
    df['Fecha']=df['Date/Time'].str.extract('(....-..-..)',expand=True)
    df['Hora']=df['Date/Time'].str.extract('(..:..:..)',expand=True)

    caja_de_texto.insert(0,"Se ha realizado limpieza de datos")
    df.to_excel('limpio.xlsx',index=False)
    return df
    

def extraerListaNombres(df):
    #borrar todas las columnas que no sean nombres
    df=df.drop(columns=['ID', 'Date/Time','Clock-in/out','Fecha','Hora'])

    #quitar duplicados
    df = df.drop_duplicates()

    #exportar lista de nombres en un excel
    df.to_excel('ListaNombres.xlsx',index=False)
    caja_de_texto.insert(0,"Se ha extraido la lista de personas")


def abrirArchivo():
     path_archivo = filedialog.askopenfilename(title="Abrir", initialdir=r"C:\Users\Lucas\Desktop\Practicas comunitarias Don Bosco")
     print(path_archivo) #en archivo se guarda el path del archivo que quiero abrir
     df = pd.read_excel(path_archivo)    
     messagebox.showinfo(message="El archivo ha sido importado", title="Horarios cargados")

     #Acciones que suceden despues de abrir el libro
     df= limpiezaExcel(df)
     extraerListaNombres(df)

def cargarListaPersonas():
    excelNombres = pd.read_excel('ListaNombres.xlsx')  
    excelNombres.reset_index(drop=True, inplace=True)
    nombres=excelNombres["Name"]
    vlist=[]

    for i in nombres:
        vlist.append(i)
    
    return vlist




if __name__ == "__main__":
    raiz=Tk()
    raiz.geometry('410x350')
    raiz.title("Sistema de control de horarios - Hogar Don Bosco")

    #Frames
    frame = Frame(raiz, width=500,height=100)
    frame.config(bg="deep sky blue")   
    frame.pack()

    frame2 = Frame(raiz,width=500,height=100)  
    frame2.config(bg="gold")   
    frame2.pack()

    frame3 = Frame(raiz,width=500,height=100)  
    frame3.config(bg="SlateBlue2")   
    frame3.pack()

    frame4 = Frame(raiz,width=500,height=450)  
    frame4.config(bg="SpringGreen2")   
    frame4.pack()
    
    #Elementos de la ventana
    label = Label(frame,text="Hogar Don Bosco")
    label.config(font=("Arial",24)) 
    label.pack()
    
    # ...................ComboBox....................

    vlist=cargarListaPersonas()
    
    Combo = ttk.Combobox(frame3, values = vlist)
    Combo.config(width=200)
    Combo.set("Seleccione a un empleado")
    Combo.pack(padx = 5, pady = 5)
    #...........................................

    Button(frame2, text="Abrir archivo", command=abrirArchivo).pack()
    
    caja_de_texto = tk.Listbox(frame4)
    caja_de_texto.place(x=5, y=25, width=400, height=200)   
    
    raiz.mainloop()




