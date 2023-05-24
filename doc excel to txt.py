from tkinter import *
from tkinter import ttk, messagebox, simpledialog
import pandas as pd
import tkinter as tk
from tkinter.filedialog import askopenfilename
import numpy as np




class App(tk.Frame):
    def __init__(self, master):
        super().__init__(master)
        self.pack()
        self.widgets()
        self.list_xlsx = []

    def widgets(self):
        self.search_f1 = tk.Button(self)
        self.search_f1["text"] = "Seleccionar doc excel"
        self.search_f1["command"] = self.search
        self.search_f1.pack(side="left")

        self.conv = tk.Button(self)
        self.conv["text"] = "Convertir"
        self.conv["command"] = self.convert
        self.conv.pack(side="right")
        

        self.quit = tk.Button(self, text="QUIT", fg="red",command=root.destroy)
        self.quit.pack(side="bottom")


    def search(self):
        self.archivo = askopenfilename(initialdir="/", title="Seleccionar archivo")
        if self.archivo:
            nombre_archivo = self.archivo
        else:
            nombre_archivo = ""
        
        self.list_xlsx.append(nombre_archivo)

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
            
            if np.isnan(rut[x]) and np.isnan(cotiz[x]):
                f_list.append(f'\n{dia[x]},{debe[x]},{haber[x]},"{concepto[x]}",,{mon[x]},{total[x]},{codiva[x]},{iva[x]},,{libro[x]}')    
            elif np.isnan(rut[x]):
                f_list.append(f'\n{dia[x]},{debe[x]},{haber[x]},"{concepto[x]}",,{mon[x]},{total[x]},{codiva[x]},{iva[x]},{cotiz[x]},{libro[x]}')
            elif np.isnan(cotiz[x]):
                f_list.append(f'\n{dia[x]},{debe[x]},{haber[x]},"{concepto[x]}",,{mon[x]},{total[x]},{codiva[x]},{iva[x]},,{libro[x]}')

        dir_fin = r"C:\Users\USUARIO\Documents\VS\Tkinter"
        f = open (f"{dir_fin}/IM{self.year}{self.month}.txt", "w")
        f.writelines(f_list)
        self.show_message_box()
        f.close()

root = tk.Tk()
myapp = App(root)
myapp.mainloop()
