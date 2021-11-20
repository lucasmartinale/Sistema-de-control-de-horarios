from tkinter import *
from tkinter.ttk import *
from ventana_empleado import *
from ver_errores_persona import *
import tkinter as tk


#Devuelve la lista de personas (Se usa para la GUI)
def cargarListaPersonas():
    excelNombres = pd.read_excel('ListaNombres.xlsx')
    excelNombres.reset_index(drop=True, inplace=True)
    nombres = excelNombres["Name"]
    vlist = []

    for i in nombres:
        vlist.append(i)

    return vlist
 

def selection_changed(combo,calcular_sueldo_boton,ver_errores_boton):
    opcion = combo.get()
    ver_errores_boton.configure(state=tk.NORMAL, command=ventana_errores(opcion))
    print("Se abrio la ventana de sueldo")
    #calcular_sueldo_boton.configure(state=tk.NORMAL, command=)   


def ventana_seleccion_empleados():
        #ventana 
    root= Tk()
    root.title("Empleados")

    #Frame que va a contener a todos
    frame_principal= Frame(root) #el frame va a estar en el root
    frame_principal.grid()
    frame_principal.config(width=480, height=320)

    #TODO: centrar el texto en el grid
    titulo = Label(frame_principal, text="Empleados", font=("Arial",24),) #Titulo esta contenido en el frame principal
    titulo.grid(column=0,row=0,padx=10,pady=10,)

    #Botones
    calcular_sueldo_boton= Button(frame_principal,text="Calcular Sueldo", state=tk.DISABLED)
    calcular_sueldo_boton.grid(column=0,row=2,padx=10, pady=10)
    

    ver_errores_boton= Button(frame_principal,text="Ver errores")
    ver_errores_boton.grid(column=1,row=2,padx=10, pady=10)


    #Combobox
    vlist=cargarListaPersonas()

    combo = Combobox(frame_principal, values = vlist, state="readonly")
    combo.config(width=50)
    #combo.set("Seleccione a un empleado")
    combo.grid(column=0,row=1, columnspan=2) #columspan es para que ocupe mas de una grilla
    combo.current(0)
    combo.bind("<<ComboboxSelected>>", lambda e:selection_changed(combo,calcular_sueldo_boton,ver_errores_boton))

    root=mainloop()

if __name__=="__main__":
    ventana_seleccion_empleados()
