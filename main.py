from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd
import tkinter as tk
from tkinter import ttk
import os
from shutil import rmtree
from os import remove


def procesarHorarioPersona(_nombre_persona):
    print("-------->procesarHorarioPersona")
    caja_de_texto.insert(0, "Procesar horario de persona")

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


        else:  #Si es la ultima fila
            print(".:Ultima Fila:.")
            if(df.loc[i]["Clock-in/out"] == "C/In"):
                print("ES C/IN")
                print(df.loc[i]["Date/Time"])
            else:
                print("ES C/OUT")

    df_sin_errores.to_excel(path_raiz +'\\ExcelsPersonas\\'+_nombre_persona+'\\'+_nombre_persona+'_sin_errores.xlsx', index=False)
    df_errores.to_excel(path_raiz +'\\ExcelsPersonas\\'+_nombre_persona+'\\'+_nombre_persona+'_errores.xlsx', index=False)
   
    


def procesarTodosLosHorarios():
    print("-------->ProcesarTodosLosHorarios")
    caja_de_texto.insert(0, "Procesar todos los horarios")
    # TODO: Tomar la lista de personas y ponerla en un vector
    listaNombres = pd.read_excel('ListaNombres.xlsx')
    vector_nombres = list(listaNombres['Name'])

    # TODO: Hacer un while que tome los excels de las personas
    # Dentro del while busca el nombre que coincide con el elemento de la lista de personas
    for i in vector_nombres:
        procesarHorarioPersona(i)


# Crea un excel para cada empleado con sus respectivas entradas y salidas
def dividirEmpleadosEnExcels():
    print("-------->dividirEmpleadosEnExcels")
    caja_de_texto.insert(0, "Dividir los empleados en excels")

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


def definirTurno():
    print("-------->definirTurno")
    excelLimpio = pd.read_excel('limpio.xlsx')
    # Se filtran las filas que son C/in (osea entradas)
    entradas_cin = excelLimpio['Clock-in/out'] == "C/In"
    entradas_limpias = excelLimpio[entradas_cin]
    entradas_limpias.to_excel('soloEntradas.xlsx', index=False)

    excelLimpio = pd.read_excel('soloEntradas.xlsx')
    excelLimpio['Turno'] = "Nada"

    for i, fila in excelLimpio.iterrows():
        # con ese numero determinar el turno
        # extraemos los dos primeros caracteres
        subcadena = fila['Hora'][0:2]
        # lo convertimos en numero
        numero_hora = int(subcadena)
        # en base al numero determinamos el turno

        if 20 <= numero_hora <= 23:
            excelLimpio.at[i, 'Turno'] = 'Nocturno'
        elif 0 <= numero_hora <= 3:
            excelLimpio.at[i, 'Turno'] = 'Nocturno'
        elif 4 <= numero_hora <= 12:
            excelLimpio.at[i, 'Turno'] = 'Mañana'
        elif 13 <= numero_hora <= 19:
            excelLimpio.at[i, 'Turno'] = 'Tarde'
        else:
            print("Error")
    excelLimpio.to_excel('soloEntradas.xlsx', index=False)


def limpiezaExcel(df):
    print("-------->limpiezaExcel")
    caja_de_texto.insert(0, "Limpieza de Excel")
    df = df.drop(columns=['Num', 'Department', 'Verifycode',
                 'Device ID', 'Device Name', 'UserExtFmt'])

    # Extraer fecha y hora
    df['Fecha'] = df['Date/Time'].str.extract('(....-..-..)', expand=True)
    df['Hora'] = df['Date/Time'].str.extract('(..:..:..)', expand=True)

    caja_de_texto.insert(0, "Se ha realizado limpieza de datos")
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
    caja_de_texto.insert(0, "Se ha extraido la lista de personas")


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


def cargarListaPersonas():
    excelNombres = pd.read_excel('ListaNombres.xlsx')
    excelNombres.reset_index(drop=True, inplace=True)
    nombres = excelNombres["Name"]
    vlist = []

    for i in nombres:
        vlist.append(i)

    return vlist


if __name__ == "__main__":
    #Antes que nada borrar todas las carpetas y archivos de alguna ejecucion pasada
    try:
        remove("limpio.xlsx")
    except:
        pass

    try:        
        remove("ListaNombres.xlsx")
    except:
        pass

    try:
        remove("soloEntradas.xlsx")
    except:
        pass

    try:
        path_raiz = os.getcwd()
        rmtree(path_raiz+'\\ExcelsPersonas')
        print('Se borro ExcelsPersonas')
    except:
        print('No existe ExcelsPersonas asi que no se borra')

    #Interfaz
    raiz = Tk()
    raiz.geometry('410x350')
    raiz.title("Sistema de control de horarios - Hogar Don Bosco")

    # Frames
    frame = Frame(raiz, width=500, height=100)
    frame.config(bg="deep sky blue")
    frame.pack()

    frame2 = Frame(raiz, width=500, height=100)
    frame2.config(bg="gold")
    frame2.pack()

    frame3 = Frame(raiz, width=500, height=100)
    frame3.config(bg="SlateBlue2")
    frame3.pack()

    frame4 = Frame(raiz, width=500, height=450)
    frame4.config(bg="SpringGreen2")
    frame4.pack()

    # Elementos de la ventana
    label = Label(frame, text="Hogar Don Bosco")
    label.config(font=("Arial", 24))
    label.pack()

    # ...........................................

    Button(frame2, text="Abrir archivo", command=abrirArchivo).pack()

    caja_de_texto = tk.Listbox(frame4)
    caja_de_texto.place(x=5, y=25, width=400, height=200)

    raiz.mainloop()
