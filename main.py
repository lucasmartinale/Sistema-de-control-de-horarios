from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd

import tkinter as tk
from tkinter import ttk


#Crea un excel para cada empleado con sus respectivas entradas y salidas
def dividirEmpleadosEnExcels():
    excelLimpio = pd.read_excel('limpio.xlsx')  
    listaNombres = pd.read_excel('ListaNombres.xlsx')  

    #guardo los nombres en un array
    vector_nombres = list(listaNombres['Name'])

    #recorrer listaNombres
    for i in vector_nombres :

        filas_persona = excelLimpio['Name'] == i
        filas_que_cumplen = excelLimpio[filas_persona]
        filas_que_cumplen.to_excel(i+'_horarios.xlsx', index=False)



def definirTurno():
    excelLimpio = pd.read_excel('limpio.xlsx')  
    #Se filtran las filas que son C/in (osea entradas)
    entradas_cin = excelLimpio['Clock-in/out'] == "C/In"
    entradas_limpias=excelLimpio[entradas_cin]
    entradas_limpias.to_excel('soloEntradas.xlsx',index=False)

    
    excelLimpio = pd.read_excel('soloEntradas.xlsx') 
    excelLimpio['Turno']="Nada"

    for i, fila in excelLimpio.iterrows():
        #con ese numero determinar el turno
        #extraemos los dos primeros caracteres
        subcadena = fila['Hora'][0:2]
        #lo convertimos en numero
        numero_hora=int(subcadena)
        #en base al numero determinamos el turno
        
        if 20<=numero_hora<=23:
            excelLimpio.at[i,'Turno']='Nocturno'
        elif 0<=numero_hora<=3:
            excelLimpio.at[i,'Turno']='Nocturno'
        elif 4<=numero_hora<=12:
            excelLimpio.at[i,'Turno']='Mañana'
        elif 13<=numero_hora<=19:
            excelLimpio.at[i,'Turno']='Tarde'
        else:
            print("Error")
    excelLimpio.to_excel('soloEntradas.xlsx',index=False)

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
     definirTurno()

def cargarListaPersonas():
    excelNombres = pd.read_excel('ListaNombres.xlsx')  
    excelNombres.reset_index(drop=True, inplace=True)
    nombres=excelNombres["Name"]
    vlist=[]

    for i in nombres:
        vlist.append(i)
    
    return vlist

if __name__ == "__main__":
    dividirEmpleadosEnExcels()  #TODO: CAMBIARLE EL LUGAR EN DONDE SE IMPLEMENTA, ESTA ACA SOLO PARA TESTEAR
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




