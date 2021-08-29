from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd

import tkinter as tk
from tkinter import ttk

def limpiezaExcel(df):
    df=df.drop(columns=['Num', 'Department','Verifycode','Device ID','Device Name','UserExtFmt'])
    print(df)
    
    #Extraer fecha y hora
    df['Fecha']=df['Date/Time'].str.extract('(....-..-..)',expand=True)
    df['Hora']=df['Date/Time'].str.extract('(..:..:..)',expand=True)

    print(df)
    caja_de_texto.insert(0,"Se ha realizado limpieza de datos")
    return df
    #df.to_excel('limpio.xlsx',index=False)

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



if __name__ == "__main__":
    raiz=Tk()
    raiz.geometry('410x310')
    raiz.title("Sistema de control de horarios - Hogar Don Bosco")

    #Frames
    frame = Frame(raiz, width=500,height=100)
    frame.config(bg="lightblue")   
    frame.pack()

    frame2 = Frame(raiz,width=500,height=100)  
    frame.config(bg="lightgreen")   
    frame2.pack()

    frame3 = Frame(raiz,width=500,height=450)  
    frame.config(bg="lightblue")   
    frame3.pack()
    
    #Elementos de la ventana
    label = Label(frame,text="Hogar Don Bosco")
    label.config(font=("Arial",24)) 
    label.pack()
    
    
    Button(frame2, text="Abrir archivo", command=abrirArchivo).pack()
    
    caja_de_texto = tk.Listbox(frame3)
    caja_de_texto.place(x=5, y=25, width=400, height=200)   
    
    raiz.mainloop()




