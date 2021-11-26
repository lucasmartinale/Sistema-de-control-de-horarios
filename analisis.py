from typing import Annotated
import pandas as pd
import os
from datetime import timedelta
from datetime import datetime

PRECIO_HORA_MAÑANA=100
PRECIO_HORA_TARDE=101
PRECIO_HORA_NOCHE=150

def calcular_horas_trabajadas(_nombre_persona):

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
    #print("{horas} horas y {minutos} minutos".format(horas=nuevas_horas,minutos=nuevos_minutos))

    #Exportar excel
    df_analisis.to_excel(path_raiz +'\\ExcelsPersonas\\'+_nombre_persona+'\\'+_nombre_persona+'_analisis.xlsx', index=False)
    
    
    return {'horas':nuevas_horas  ,'minutos':nuevos_minutos}

#Le digo una hora, me dice el turno
def calcular_turno(hora):
    if 22 <= hora <= 23:
        turno='noche'
    elif 0 <= hora <= 5:
        turno='noche'
    elif 6 <= hora <= 12:
        turno='mañana'
    elif 13 <= hora <= 21:
        turno = 'tarde'
    else:
        print("Error")

    return turno

#Le paso la lista de horas me devuelve la lista de turnos
def calcular_total_a_pagar(lista_turnos_horas):
    total_precio_mañana =lista_turnos_horas.count("mañana") * PRECIO_HORA_MAÑANA
    total_precio_tarde =lista_turnos_horas.count("tarde") * PRECIO_HORA_TARDE
    total_precio_noche =lista_turnos_horas.count("noche") * PRECIO_HORA_NOCHE
    return total_precio_mañana + total_precio_tarde + total_precio_noche    

def calcular_sueldo(horasyminutos, _nombre_persona):
    # Busca el path de la carpeta raiz
    path_raiz = os.getcwd()

    # Abro el archivo sin errores de la persona
    df = pd.read_excel(path_raiz +
                       '\\ExcelsPersonas\\'+_nombre_persona+'\\'+_nombre_persona+'_analisis.xlsx')
    print("Abrio el archivo de excel de "+_nombre_persona)

    
    df['Total a pagar'] = 0
    df['Cantidad horas noche'] = 0
    df['Cantidad horas tarde'] = 0
    df['Cantidad horas mañana'] = 0

    
    #Recorre todas las filas
    for i in range(0,df.shape[0]):
        #Traemos el campo 'Inicio' de la fila i
        entrada=df.loc[i]["Inicio"]
        entrada=pd.to_datetime(entrada)

        lista_turnos_horas=[]
        #Agarro la primer hora y calculo el precio de esa hora
        lista_turnos_horas.append(calcular_turno(entrada.hour) )
        
        hora_siguiente=entrada
        #Recorre las horas trabajadas agregando a la lista la palabra tarde noche o mañana
        for j in range(int(df.loc[i]['Total'])-1):
            
                      
            #Suma una hora a la 
            hora_siguiente= hora_siguiente + timedelta(hours=1)
            
            print(f"hora siguiente {hora_siguiente}")
            
            #Calcula turno de esa hora siguiente y la añade a la lista
            lista_turnos_horas.append(calcular_turno(hora_siguiente.hour) )

        
        df.at[i, 'Cantidad horas noche'] = lista_turnos_horas.count('noche')
        df.at[i, 'Cantidad horas tarde'] = lista_turnos_horas.count('tarde')
        df.at[i, 'Cantidad horas mañana'] = lista_turnos_horas.count('mañana')

        print(lista_turnos_horas)
        total_a_pagar = calcular_total_a_pagar(lista_turnos_horas)
        print(f"cantidad a pagar: {total_a_pagar}")
        
        df.at[i, 'Total a pagar'] = total_a_pagar


    #Exportar excel
    df.to_excel(path_raiz +'\\ExcelsPersonas\\'+_nombre_persona+'\\'+_nombre_persona+'_analisis.xlsx', index=False)


if __name__=="__main__":
    horasyminutos = calcular_horas_trabajadas("ALMARAZ ANGELICA")
    calcular_sueldo(horasyminutos,"ALMARAZ ANGELICA")