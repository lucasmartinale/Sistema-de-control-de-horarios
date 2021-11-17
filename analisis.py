import pandas as pd
import os
from datetime import timedelta
from datetime import datetime

def AnalisisHoras(_nombre_persona):

    # Busca el path de la carpeta raiz
    path_raiz = os.getcwd()

    # Abro el archivo sin errores de la persona
    df = pd.read_excel(path_raiz +
                       '\\ExcelsPersonas\\'+_nombre_persona+'\\'+_nombre_persona+'_sin_errores.xlsx')
    print("Abrio el archivo de excel de "+_nombre_persona)

    #cant_filas=range(df.shape[0])

    #Se crea el dataframe para el analisis
    df_analisis = pd.DataFrame(columns = ['Inicio', 'Fin', 'Total'])
    
    horas_trabajadas_totales=0

    #Busco la cantidad de horas que hay entre cada par
    for i in range(0,df.shape[0],2):
        
        print("--------------")
        #print(i)
        entrada=df.loc[i]["Date/Time"]
        #print(entrada)
        t1=pd.to_datetime(entrada)
        

        salida=df.loc[i+1]["Date/Time"]
        #print(salida)
        t2=pd.to_datetime(salida)

        horas_trabajadas_par= t2-t1
        print("Par: {horas}".format(horas=horas_trabajadas_par))

        #Va a el excel de errores
        fila_a_insertar = {'Name': df.loc[i]['Name'], 
                            'Inicio': t1,
                            'Fin': t2, 
                            'Total': horas_trabajadas_par
                            }

                
        #Añade la columna al dataframe de analisis
        df_analisis=df_analisis.append(fila_a_insertar, ignore_index = True)        

        #Summa las horas para obtener las totales
        horas_trabajadas_totales+= int(horas_trabajadas_par.seconds//3600)
        

    print("Horas totales trabajadas: {horas}".format(horas=horas_trabajadas_totales))

    #Exporta excel a la carpeta del empleado
    print("Exportando excel")
    df_analisis.to_excel(path_raiz +'\\ExcelsPersonas\\'+_nombre_persona+'\\'+_nombre_persona+'_analisis.xlsx', index=False)


AnalisisHoras("ALMARAZ ANGELICA")