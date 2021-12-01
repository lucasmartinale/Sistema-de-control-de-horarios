from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd
import tkinter as tk
from tkinter import ttk
import os
from shutil import rmtree
from os import remove
from ventana_empleados import *


def procesarHorarioPersona(_nombre_persona):
    print("-------->procesarHorarioPersona")
    
    # Busca el path de la carpeta raiz
    path_raiz = os.getcwd()

    # Abro el archivo de la persona
    df = pd.read_excel(path_raiz +
                       '\\ExcelsPersonas\\'+_nombre_persona+'\\'+_nombre_persona+'_horarios.xlsx')
    print("Abrio el archivo de excel de "+_nombre_persona)

    # guardo los nombres en un array
    lista_cin_cout = list(df['Clock-in/out'])

    #Se crea dataframe para los horarios que no tienen errores
    df_sin_errores = pd.DataFrame(columns = ['Name', 'ID', 'Date/Time', 'Clock-in/out', 'Fecha','Hora'])
    #Se crea un dataframe para los horarios con errores
    df_errores = pd.DataFrame(columns = ['Name', 'ID', 'Date/Time', 'Clock-in/out', 'Fecha','Hora'])
    
    cant_filas=range(df.shape[0])  #es un rango
    print("Cantidad de filas:", cant_filas)
    for i in cant_filas:

        if(i != df.shape[0]-1):   #Si no es la ultima fila

            #¿es un c/in?
            if(df.loc[i]["Clock-in/out"] == "C/In"):

                # ¿El siguiente es un c/out?
                #Si
                if(df.loc[i+1]["Clock-in/out"] == "C/Out"):
                    #Se guardan las 2 filas en excel sin_errores
                    #Guarda primera fila
                    fila_a_insertar = {'Name': df.loc[i]['Name'],
                                        'ID': df.loc[i]['ID'],
                                        'Date/Time': df.loc[i]['Date/Time'],
                                        'Clock-in/out': df.loc[i]['Clock-in/out'],
                                        'Fecha': df.loc[i]['Fecha'],
                                        'Hora': df.loc[i]['Hora'],
                                        }

                    df_sin_errores=df_sin_errores.append(fila_a_insertar, ignore_index = True)
                    #Guarda segunda fila
                    fila_a_insertar = {'Name': df.loc[i+1]['Name'],
                                        'ID': df.loc[i+1]['ID'],
                                        'Date/Time': df.loc[i+1]['Date/Time'],
                                        'Clock-in/out': df.loc[i+1]['Clock-in/out'],
                                        'Fecha': df.loc[i+1]['Fecha'],
                                        'Hora': df.loc[i+1]['Hora'],
                                        }
                    
                    df_sin_errores=df_sin_errores.append(fila_a_insertar, ignore_index = True)
                    

                #No
                else:
                    #Guarda primera fila en archivo de errores
                    fila_a_insertar = {'Name': df.loc[i]['Name'],
                                        'ID': df.loc[i]['ID'],
                                        'Date/Time': df.loc[i]['Date/Time'],
                                        'Clock-in/out': df.loc[i]['Clock-in/out'],
                                        'Fecha': df.loc[i]['Fecha'],
                                        'Hora': df.loc[i]['Hora'],
                                        }

                    df_errores=df_errores.append(fila_a_insertar, ignore_index = True)
                    #Guarda segunda fila en archivo de errores
                    fila_a_insertar = {'Name': df.loc[i+1]['Name'],
                                        'ID': df.loc[i+1]['ID'],
                                        'Date/Time': df.loc[i+1]['Date/Time'],
                                        'Clock-in/out': df.loc[i+1]['Clock-in/out'],
                                        'Fecha': df.loc[i+1]['Fecha'],
                                        'Hora': df.loc[i+1]['Hora'],
                                        }
                    
                    df_errores=df_errores.append(fila_a_insertar, ignore_index = True)
            #Es un c/out
            else:
                #Va a el excel de errores
                fila_a_insertar = {'Name': df.loc[i]['Name'],
                                        'ID': df.loc[i]['ID'],
                                        'Date/Time': df.loc[i]['Date/Time'],
                                        'Clock-in/out': df.loc[i]['Clock-in/out'],
                                        'Fecha': df.loc[i]['Fecha'],
                                        'Hora': df.loc[i]['Hora'],
                                        }
                    
                df_errores=df_errores.append(fila_a_insertar, ignore_index = True)

    
    df_sin_errores.to_excel(path_raiz +'\\ExcelsPersonas\\'+_nombre_persona+'\\'+_nombre_persona+'_sin_errores.xlsx', index=False)
    df_errores.to_excel(path_raiz +'\\ExcelsPersonas\\'+_nombre_persona+'\\'+_nombre_persona+'_errores.xlsx', index=False)


def procesarTodosLosHorarios():
    print("-------->ProcesarTodosLosHorarios")
        
    listaNombres = pd.read_excel('ListaNombres.xlsx')
    vector_nombres = list(listaNombres['Name'])

    # Dentro del while busca el nombre que coincide con el elemento de la lista de personas
    for i in vector_nombres:
        procesarHorarioPersona(i)
    
    #Se avisa por medio de la caja de texto que se ha terminado de procesar
    caja_de_texto.insert(tk.END, "Se han procesado los horarios de todos los empleados")
    caja_de_texto.itemconfigure(tk.END, bg="#00aa00", fg="#fff")
    
    #Se habilita el boton para ver los empleados (porque ya estan listos los datos)
    boton_empleados.config(state="normal")
    print("Se ha procesado toda la informacion")


# Crea un excel para cada empleado con sus respectivas entradas y salidas
def dividirEmpleadosEnExcels():
    print("-------->dividirEmpleadosEnExcels")
    
    excelLimpio = pd.read_excel('limpio.xlsx')
    listaNombres = pd.read_excel('ListaNombres.xlsx')

    # guardo los nombres en un array
    vector_nombres = list(listaNombres['Name'])

    # Crea la carpeta ExcelsPersonas para meter los excels de manera organizada
    os.makedirs('ExcelsPersonas', exist_ok=True)

    # recorrer listaNombres
    for i in vector_nombres:
        filas_persona = excelLimpio['Name'] == i
        filas_que_cumplen = excelLimpio[filas_persona]

        #Crea una carpeta por persona
        os.makedirs('ExcelsPersonas/'+i)


        # Busca el path de la carpeta raiz
        path_excelsPersonas = os.getcwd()
        # Crea el excel en la carpeta ExcelPersonas que esta dentro de la carpeta raiz
        filas_que_cumplen.to_excel(
            path_excelsPersonas+'\\ExcelsPersonas\\'+i+'\\'+i+'_horarios.xlsx', index=False)



def limpiezaExcel(df):
    print("-------->limpiezaExcel")
    
    df = df.drop(columns=['Num', 'Department', 'Verifycode',
                 'Device ID', 'Device Name', 'UserExtFmt'])

    # Extraer fecha y hora
    df['Fecha'] = df['Date/Time'].str.extract('(....-..-..)', expand=True)
    df['Hora'] = df['Date/Time'].str.extract('(..:..:..)', expand=True)

    
    df.to_excel('limpio.xlsx', index=False)
    return df


def extraerListaNombres(df):
    print("-------->extraerListaNombres")
    # borrar todas las columnas que no sean nombres
    df = df.drop(columns=['ID', 'Date/Time', 'Clock-in/out', 'Fecha', 'Hora'])

    # quitar duplicados
    df = df.drop_duplicates()

    # exportar lista de nombres en un excel
    df.to_excel('ListaNombres.xlsx', index=False)
    

def abrirArchivo():
    print("-------->abrirArchivo")
    
    path_archivo = filedialog.askopenfilename(
        title="Abrir", initialdir=r"C:\Users\Lucas\Desktop\Todo COMUNITARIAS\Practicas comunitarias Don Bosco")
    # en archivo se guarda el path del archivo que quiero abrir
    print(path_archivo)
    df = pd.read_excel(path_archivo)
    messagebox.showinfo(message="El archivo ha sido importado",
                        title="Horarios cargados")

    # Acciones que suceden despues de abrir el libro
    df = limpiezaExcel(df)  # guarda el limpio en df
    extraerListaNombres(df)
    # definirTurno()
    dividirEmpleadosEnExcels()
    procesarTodosLosHorarios()


def limpiarCarpeta():
    try:
        remove("limpio.xlsx")
    except:
        print("No se pudo eliminar limpio.xlsx")

    try:        
        remove("ListaNombres.xlsx")
    except:
        print("No se pudo borrar ListaNombres.xlsx")

    try:
        remove("soloEntradas.xlsx")
    except:
        print("No se pudo borrar soloEntradas.xlsx")

    try:
        path_raiz = os.getcwd()
        rmtree(path_raiz+'\\ExcelsPersonas')
        print('Se borro ExcelsPersonas')
    except:
        print('No existe ExcelsPersonas asi que no se borra')


if __name__ == "__main__":
    #Antes que nada borrar todas las carpetas y archivos de alguna ejecucion pasada
    limpiarCarpeta()

    #Interfaz
    raiz = Tk()
    raiz.geometry('410x310')
    raiz.title("Sistema de control de horarios - Hogar Don Bosco")
    raiz.resizable(0, 0)

    # Frames
    frame1=Frame(raiz)
    frame1.grid(row=1)


    frame2=Frame(raiz)
    frame2.grid(row=2)


    frame3=Frame(raiz)
    frame3.grid(row=3)

    frame4=Frame(raiz)
    frame4.grid(row=4)

    # Elementos de la ventana
    label = Label(frame1, text="Hogar Don Bosco")
    label.config(font=("Arial", 24))
    label.pack()

    # ...........................................

    Button(frame2, text="Abrir archivo", command=abrirArchivo).grid(padx=5,pady=5, row=0, column=1)

    caja_de_texto = tk.Listbox(frame3, width=66)
    caja_de_texto.grid(padx=5,pady=5, row=0, column=1)

    boton_empleados = Button(frame4, text="Ver Empleados",state="disabled", command=ventana_seleccion_empleados)
    boton_empleados.grid(padx=5,pady=5, row=0, column=1)

    raiz.mainloop()
