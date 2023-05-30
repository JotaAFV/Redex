from tkinter import *
from tkinter import ttk, messagebox, simpledialog, Label
from PIL import ImageTk, Image
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

        self.list_xlsx = []
        self.fold_list = []
        with open("C:/Users/USUARIO/Documents/VS/1.0/files/directory.pickle", "rb") as file:
            dir_fin = pickle.load(file)
            self.fold_list.append(dir_fin)
        self.widgets()

    def data_for_ejec(self):
        dir_fold = self.fold_list[-1]
        with open("C:/Users/USUARIO/Documents/VS/1.0/files/directory.pickle", "wb") as file:
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
        #self.dir1 = tk.Label(self, text=str(self.list_xlsx[-1]))


        self.conv = tk.Button(self)
        self.conv["text"] = "Convertir"
        self.conv["command"] = self.convert
        self.conv.grid(column=2, row=1, padx=10, pady=10)
        

        self.quit = tk.Button(self, text="QUIT", fg="red",command=root.destroy)
        self.quit.grid(column=1, row=3, sticky="S")

    def confis(self):
        ventana_secundaria = tk.Toplevel(root)
        ventana_secundaria.title("Configuracion")
        ventana_secundaria.config(width=300, height=200)
        # Crear un botón dentro de la ventana secundaria
        boton_busca = ttk.Button(ventana_secundaria, text="Elegir carpeta Final", command=self.search_folder)
        boton_busca.place(x=20, y=40)
        text1 = ttk.Label(ventana_secundaria, text="Carpeta Actual:", font="Red").place(x=20, y=70)
        try:
            dir1 = ttk.Label(ventana_secundaria, text=str(self.fold_list[-1])).place(x=20, y=90)     
        except IndexError:
            dir1 = ttk.Label(ventana_secundaria, text="No hay carpeta aun").place(x=20, y=90)

        # Boton crear libro
        boton_crea = ttk.Button(ventana_secundaria, text="Crear Libro base", command=self.create_book)
        boton_crea.place(x=20, y=150)
        # para cerrar la misma.
        boton_cerrar = ttk.Button(ventana_secundaria, text="Cerrar ventana", command=ventana_secundaria.destroy)
        boton_cerrar.place(x=110, y=175)

    def search(self):
        try:
            last_search = self.list_xlsx[-1]
        except IndexError:
            last_search = "/"
        self.archivo = askopenfilename(initialdir=f"{last_search}", title="Seleccionar archivo")
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

    def create_book(self):
        data = {
        "Dia": [],
        "Debe": [],
        "Haber": [],
        "Concepto": [],
        "RUT": [],
        "Moneda": [],
        "Total": [],
        "CodigoIVA": [],
        "IVA": [],
        "Cotizacion": [],
        "Libro": []
        }
        df = pd.DataFrame(data)

        writer = pd.ExcelWriter("C:/Users/USUARIO/Documents/VS/1.0/pre_import.xlsx", engine='xlsxwriter')
        df.to_excel(writer, sheet_name='Hoja1', index=False)
        writer.close()

    def convert(self):
        self.show_input_box()
        dir_exl = self.list_xlsx[-1]
        archivo_excel = pd.read_excel(f"{dir_exl}")

        fecha = archivo_excel["Dia"].values 
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
            fech = re.search(r"^(\d+)/?(\d+)/?(\d+)", str(fecha[x]))
            
            if not fech:
                dia = f"{1:02}"
            else:
                dia = fech.group(1)

            if np.isnan(rut[x]) and np.isnan(cotiz[x]):
                f_list.append(f'\n{dia},{debe[x]},{haber[x]},"{concepto[x]}",,{mon[x]},{total[x]},{codiva[x]},{iva[x]},,{libro[x]}')    
            elif np.isnan(rut[x]):
                f_list.append(f'\n{dia},{debe[x]},{haber[x]},"{concepto[x]}",,{mon[x]},{total[x]},{codiva[x]},{iva[x]},{cotiz[x]},{libro[x]}')
            elif np.isnan(cotiz[x]):
                f_list.append(f'\n{dia},{debe[x]},{haber[x]},"{concepto[x]}",,{mon[x]},{total[x]},{codiva[x]},{iva[x]},,{libro[x]}')

        dir_fin = r"C:\Users\USUARIO\Documents\GitHub\Tkinter"
        
        with open("C:/Users/USUARIO/Documents/VS/1.0/files/directory.pickle", "rb") as file:
            dir_fin = pickle.load(file)
        
        try:
            f = open (f"{dir_fin}/IM{self.year}{self.month}.txt", "w")
        except FileNotFoundError:
            self.warn = messagebox.showinfo("Erroro", f"No se pudo encontrar la dreccion final para el archivo IM{self.year}{self.month}.txt")


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
