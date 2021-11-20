from tkinter import *
from tkinter.ttk import *

def crear_ventana_empleado():
    #ventana 
    root= Tk()
    root.title("Empleados")

"""     #Frame que va a contener a todos
    frame_principal= Frame(root) #el frame va a estar en el root
    frame_principal.grid()
    frame_principal.config(width=480, height=320)

    #TODO: centrar el texto en el grid
    titulo = Label(frame_principal, text="Empleados", font=("Arial",24),) #Titulo esta contenido en el frame principal
    titulo.grid(column=0,row=0,padx=10,pady=10,)

    #Combobox
    vlist=["Lucas","Martin"]

    combo = Combobox(frame_principal, values = vlist)
    combo.config(width=50)
    combo.set("Seleccione a un empleado")
    combo.grid(column=0,row=1, columnspan=2) #columspan es para que ocupe mas de una grilla

    #Botones
    calcular_sueldo_boton= Button(frame_principal,text="Calcular Sueldo")
    calcular_sueldo_boton.grid(column=0,row=2,padx=10, pady=10)

    ver_errores_boton= Button(frame_principal,text="Ver errores")
    ver_errores_boton.grid(column=1,row=2,padx=10, pady=10) """

if __name__=="__main__":
    crear_ventana_empleado()
    root=mainloop()
