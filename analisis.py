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


    #Se crea el dataframe para el analisis
    df_analisis = pd.DataFrame(columns = ['Inicio', 'Fin', 'Total','Minutos restantes','Minutos'])
    
    horas_trabajadas_totales=0

    #Busco la cantidad de horas que hay entre cada par
    for i in range(0,df.shape[0],2):
        
        #Horario de entrada
        entrada=df.loc[i]["Date/Time"]
        entrada=pd.to_datetime(entrada)
        
        #Horario de salida
        salida=df.loc[i+1]["Date/Time"]
        salida=pd.to_datetime(salida)

        #Diferencia entre la hora de entrada y la de salida
        horas_trabajadas_par= salida-entrada

        #Se calculan las horas de esa diferencia
        duration_in_s = horas_trabajadas_par.total_seconds() 
        hours = divmod(duration_in_s, 3600)[0]

        #Se Calculan los minutos de esa diferencia que no cuentan en las horas calculadas en el anterior
        minutes = divmod(duration_in_s, 60)[0]
        minutos_restantes=minutes - (hours*60)
       
        #Va a el excel de errores
        fila_a_insertar = { 'Inicio': entrada,
                            'Fin': salida, 
                            'Total': hours,
                            'Minutos restantes':minutos_restantes
                            }

        #Añade la columna al dataframe de analisis
        df_analisis=df_analisis.append(fila_a_insertar, ignore_index = True)        

        #Suma las horas para obtener las totales
        horas_trabajadas_totales+= int(horas_trabajadas_par.seconds//3600)
        
    #Horas totales
    horas_totales = int(df_analisis['Total'].sum())

    #Separa los minutos que valen la pena contar de los que no
    df_analisis.loc[(df_analisis['Minutos restantes'] < 20) , 'Minutos'] = 0  
    df_analisis.loc[(df_analisis['Minutos restantes'] >= 20) , 'Minutos'] =  df_analisis['Minutos restantes']

    #Ahora si, la suma definitiva de horas y minutos
    total_minutos_validos=df_analisis['Minutos'].sum()
    nuevas_horas = int(total_minutos_validos//60)
    nuevos_minutos= int(total_minutos_validos % 60)

    #A las horas totales se le suman las horas formadas por los minutos que vale la pena contar
    nuevas_horas=nuevas_horas+horas_totales
    print("{horas} horas y {minutos} minutos".format(horas=nuevas_horas,minutos=nuevos_minutos))

    #Exportar excel
    df_analisis.to_excel(path_raiz +'\\ExcelsPersonas\\'+_nombre_persona+'\\'+_nombre_persona+'_analisis.xlsx', index=False)

AnalisisHoras("ALMARAZ ANGELICA")