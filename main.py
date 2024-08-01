import tkinter as tk
import tkinter.ttk as ttk
import tkinter.filedialog as filedialog
import os
import pandas as pd
import openpyxl

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Hello, Tkinter!")
        self.geometry("300x100")
        self.configure(bg="#000")
        
        self.label = tk.Label(self, text="Conciliaciones bancarias", font=("Helvetica", 30), bg="#fff", fg="#000")
        self.label.pack(pady=0, fill=tk.X)
        
        self.frame_izquierdo = tk.Frame(self, bg="#fff")
        self.frame_izquierdo.pack(side=tk.LEFT, fill=tk.BOTH)
        
        self.lbl_izq = tk.Label(self.frame_izquierdo, text="", bg="#fff")
        self.lbl_izq.pack(pady=10)
        
        self.button = tk.Button(self.frame_izquierdo, text="Seleccionar archivo", command=self.elegir_archivo)
        self.button.pack(pady=10, fill=tk.X)
        
        self.lista_frame = tk.Frame(self.frame_izquierdo)
        self.lista_frame.pack(fill=tk.BOTH, expand=True)
        
        self.lista = ttk.Treeview(self.lista_frame)
        self.lista.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.scrollbar_x = ttk.Scrollbar(self.lista_frame, orient="horizontal", command=self.lista.xview)
        self.scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.lista.configure(xscrollcommand=self.scrollbar_x.set)
        
        self.scrollbar_y = ttk.Scrollbar(self.lista_frame, orient="vertical", command=self.lista.yview)
        self.scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.lista.configure(yscrollcommand=self.scrollbar_y.set)
        
        self.frame_izquierdo.pack_propagate(False)
        self.frame_izquierdo.config(width=self.winfo_screenwidth()//2)
        
        self.bind("<Left>", self.scroll_izq)
        self.bind("<Right>", self.scroll_der)
        self.bind("<Up>", self.scroll_arriba)
        self.bind("<Down>", self.scroll_abajo)
    
    def scroll_izq(self, evento):
        self.lista.xview_scroll(-15, "units")
    
    def scroll_der(self, evento):
        self.lista.xview_scroll(15, "units")
    
    def scroll_arriba(self, evento):
        self.lista.yview_scroll(-15, "units")
    
    def scroll_abajo(self, evento):
        self.lista.yview_scroll(15, "units")
        
    def elegir_archivo(self):
        try:
            ruta_archivo = filedialog.askopenfilename(
                title="Elegir archivo",
                filetypes=[("Archivos de Excel", "*.xlsx")]
            )
            
            if ruta_archivo:
                print("Ruta del archivo: "+ruta_archivo)
                nombre_archivo = os.path.basename(ruta_archivo)
                self.lbl_izq.config(text="Archivo seleccionado: "+nombre_archivo)
                
                # Se lee el archivo excel y se guarda en un DataFrame
                archivo = pd.read_excel(ruta_archivo)
                
                # Se eliminan filas y columnas con valores nulos
                archivo = archivo.dropna(how="all")
                archivo = archivo.dropna(axis=1, how="all")
                
                # Se muestra el contenido del archivo en la lista
                self.mostrar_datos(archivo)
            else:
                print("No se ha seleccionado ningún archivo")
                self.lbl_izq.config(text="No se ha seleccionado ningún archivo")
        except Exception as e:
            print("EXCEPCIÓN: "+ str(e))
            self.lbl_izq.config(text="Error al leer el archivo")
    
    def mostrar_datos(self, archivo):
        # Limpiar la lista
        self.lista.delete(*self.lista.get_children())
        
        # Configurar columnas de la lista
        self.lista["column"] = list(archivo.columns)
        self.lista["show"] = "headings"
                
        # Configurar encabezados de las columnas
        for col in archivo.columns:
            self.lista.heading(col, text=col)
        
        # Insertar datos en la lista
        for _, row in archivo.iterrows():
            self.lista.insert("", "end", values=list(row))

if __name__ == "__main__":
    app = App()
    app.mainloop()