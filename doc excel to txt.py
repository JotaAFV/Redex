from tkinter import *
from tkinter import ttk, messagebox, simpledialog, N, S, W, E
import pandas as pd
import tkinter as tk
from tkinter.filedialog import askopenfilename, askdirectory
import numpy as np
import re
import pickle




class App(tk.Frame):
    def __init__(self, master):
        super().__init__(master)
        self.grid(padx=25,pady=20)

        self.widgets()
        self.list_xlsx = []
        self.fold_list = []

    def data_for_ejec(self):
        dir_fold = self.fold_list[-1]
        with open("files/directory.pickle", "wb") as file:
            pickle.dump(dir_fold, file)
        file.close()

    def widgets(self):
        self.confi = tk.Button(self)
        self.confi["text"] = "Configuracion"
        self.confi["command"] = self.confis
        self.confi.grid(column=0, row=0, padx=10, pady=10, sticky="W")

        self.search_f1 = tk.Button(self)
        self.search_f1["text"] = "Seleccionar doc excel"
        self.search_f1["command"] = self.search
        self.search_f1.grid(column=0, row=1, padx=10, pady=10, sticky="NSW")

        self.conv = tk.Button(self)
        self.conv["text"] = "Convertir"
        self.conv["command"] = self.convert
        self.conv.grid(column=2, row=1, padx=10, pady=10)
        

        self.quit = tk.Button(self, text="QUIT", fg="red",command=root.destroy)
        self.quit.grid(column=1, row=3, sticky="S")

    def confis(self):
        ventana_secundaria = tk.Toplevel()
        ventana_secundaria.title("Ventana secundaria")
        ventana_secundaria.config(width=300, height=200)
        # Crear un botón dentro de la ventana secundaria
        boton_busca = ttk.Button(
            ventana_secundaria,
            text="Elegir carpeta Final", 
            command=self.search_folder
        )
        boton_busca.place(x=75, y=75)         
        # para cerrar la misma.
        boton_cerrar = ttk.Button(
            ventana_secundaria,
            text="Cerrar ventana", 
            command=ventana_secundaria.destroy
        )
        boton_cerrar.place(x=150, y=195)

    def search(self):
        self.archivo = askopenfilename(initialdir="/", title="Seleccionar archivo")
        if self.archivo:
            nombre_archivo = self.archivo
        else:
            nombre_archivo = ""
        
        self.list_xlsx.append(nombre_archivo)
    
    def search_folder(self):
        self.carpeta = askdirectory(initialdir="C:/Users/USUARIO/Documents/VS", title="Seleccionar capeta")
        if self.carpeta:
            nombre_carpeta = self.carpeta
        else:
            nombre_carpeta = ""
        
        self.fold_list.append(nombre_carpeta)
        self.data_for_ejec()

    def show_message_box(self):
        self.box = messagebox.showinfo("El archivo fue creado", f"IM{self.year}{self.month}.txt fue creado")
    def show_input_box(self):
        self.month = simpledialog.askstring("Mes", "Mes (en dos cifras):")
        self.year = simpledialog.askstring("Año", "Año (Ultimas dos cifras):")
     
    def convert(self):
        self.show_input_box()
        dir_exl = self.list_xlsx[-1]
        archivo_excel = pd.read_excel(f"{dir_exl}")

        dia = archivo_excel["Dia"].values 
        debe = archivo_excel["Debe"].values
        haber = archivo_excel["Haber"].values
        concepto = archivo_excel["Concepto"].values
        rut = archivo_excel["RUT"].values
        mon = archivo_excel["Moneda"].values
        total = archivo_excel["Total"].values
        codiva = archivo_excel["CodigoIVA"].values
        iva = archivo_excel["IVA"].values
        cotiz = archivo_excel["Cotizacion"].values
        libro = archivo_excel["Libro"].values

        f_list= ['Dia, Debe, Haber, Concepto, RUT, Moneda,  Total, CodigoIVA, IVA, Cotizacion, Libro']

        for x in range(len(total)):
            #dia = re.search("",fecha[x])
            
            if np.isnan(rut[x]) and np.isnan(cotiz[x]):
                f_list.append(f'\n{dia[x]},{debe[x]},{haber[x]},"{concepto[x]}",,{mon[x]},{total[x]},{codiva[x]},{iva[x]},,{libro[x]}')    
            elif np.isnan(rut[x]):
                f_list.append(f'\n{dia[x]},{debe[x]},{haber[x]},"{concepto[x]}",,{mon[x]},{total[x]},{codiva[x]},{iva[x]},{cotiz[x]},{libro[x]}')
            elif np.isnan(cotiz[x]):
                f_list.append(f'\n{dia[x]},{debe[x]},{haber[x]},"{concepto[x]}",,{mon[x]},{total[x]},{codiva[x]},{iva[x]},,{libro[x]}')

        dir_fin = r"C:\Users\USUARIO\Documents\GitHub\Tkinter"
        
        with open("files/directory.pickle", "rb") as file:
            dir_fin = pickle.load(file)
        
        f = open (f"{dir_fin}/IM{self.year}{self.month}.txt", "w")
        f.writelines(f_list)
        self.show_message_box()
        f.close()

root = tk.Tk()
myapp = App(root)
root.grid_columnconfigure(0, weight=1)
root.grid_columnconfigure(1, weight=1)
root.grid_rowconfigure(0, weight=1)
root.grid_rowconfigure(1, weight=1)
myapp.mainloop()
