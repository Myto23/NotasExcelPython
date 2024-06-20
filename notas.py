import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

class NotaIngresoApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Ingreso de Notas")
        
        # Variables
        self.nombres = []
        self.notas = pd.DataFrame()
        self.curso = ""
        self.filename = ""
        self.ponderaciones = {}
        
        # Interface Setup
        self.setup_interface()
        
    def setup_interface(self):
        # Botones y entradas para cargar nombres y curso
        ttk.Button(self.root, text="Cargar Nombres y Notas", command=self.cargar_nombres_y_notas).grid(row=0, column=0, pady=10, padx=10, sticky="ew")
        ttk.Label(self.root, text="Nombre del Curso:").grid(row=1, column=0, padx=10, sticky="w")
        self.curso_entry = ttk.Entry(self.root)
        self.curso_entry.grid(row=1, column=0, padx=10, sticky="ew")
        ttk.Button(self.root, text="Guardar Curso", command=self.guardar_curso).grid(row=1, column=1, padx=10, sticky="ew")
        
        # Área para mostrar notas de estudiantes
        self.notas_text = tk.Text(self.root, height=20, width=50)
        self.notas_text.grid(row=2, column=0, columnspan=5, padx=10, pady=10, sticky="ew")
        
        # Botones para agregar, editar y eliminar notas
        ttk.Button(self.root, text="Agregar Estudiante", command=self.agregar_estudiante).grid(row=3, column=0, padx=10, pady=10, sticky="ew")
        ttk.Button(self.root, text="Agregar Nota", command=self.agregar_nota).grid(row=3, column=1, padx=10, pady=10, sticky="ew")
        ttk.Button(self.root, text="Editar Nota", command=self.editar_nota).grid(row=3, column=2, padx=10, pady=10, sticky="ew")
        ttk.Button(self.root, text="Eliminar Estudiante", command=self.eliminar_estudiante).grid(row=3, column=3, padx=10, pady=10, sticky="ew")
        ttk.Button(self.root, text="Eliminar Nota", command=self.eliminar_nota).grid(row=3, column=4, padx=10, pady=10, sticky="ew")  # Botón para eliminar notas
        
        # Botón para promediar las notas
        ttk.Button(self.root, text="Promediar Notas", command=self.mostrar_promedio_notas).grid(row=4, column=0, columnspan=5, padx=10, pady=10, sticky="ew")
        
    def cargar_nombres_y_notas(self):
        self.filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if self.filename.endswith('.xlsx'):
            df = pd.read_excel(self.filename)
            self.nombres = df.iloc[:, 0].tolist()
            self.notas = df.set_index(df.columns[0])
            self.actualizar_interfaz_notas()
            messagebox.showinfo("Info", "Nombres y notas cargados correctamente.")
        else:
            messagebox.showwarning("Warning", "Seleccione un archivo Excel válido.")
        
    def guardar_curso(self):
        self.curso = self.curso_entry.get()
        if self.curso:
            messagebox.showinfo("Info", f"Curso '{self.curso}' guardado correctamente.")
        else:
            messagebox.showwarning("Warning", "Ingrese un nombre de curso válido.")
            
    def actualizar_interfaz_notas(self):
        self.notas_text.delete(1.0, tk.END)
        for nombre, fila in self.notas.iterrows():
            notas_text_format = f"{nombre}: "
            for nota in fila:
                color = "green" if round(nota, 2) >= 4.0 else "red"
                notas_text_format += f"{round(nota, 2)} " if nota else "- "
                self.notas_text.insert(tk.END, notas_text_format, color)
                notas_text_format = ""
            self.notas_text.insert(tk.END, "\n")
            
    def agregar_estudiante(self):
        nombre_estudiante = simpledialog.askstring("Agregar Estudiante", "Ingrese el nombre del estudiante:")
        if nombre_estudiante:
            if nombre_estudiante in self.nombres:
                messagebox.showwarning("Warning", "El estudiante ya está en la lista.")
            else:
                self.nombres.append(nombre_estudiante)
                num_notas = simpledialog.askinteger("Agregar Estudiante", f"Cuantas notas desea agregar para '{nombre_estudiante}'?")
                if num_notas:
                    notas_estudiante = {}
                    for i in range(num_notas):
                        nueva_nota = self.ask_for_valid_float(f"Ingrese la nota {i + 1} para '{nombre_estudiante}':", "Nota inválida. Debe estar entre 10 y 70.", 10, 70)
                        if nueva_nota is not None:
                            notas_estudiante[f"Nota {i + 1}"] = nueva_nota
                    self.notas = pd.concat([self.notas, pd.DataFrame(notas_estudiante, index=[nombre_estudiante])])
                    self.guardar_notas_en_excel()
                    self.actualizar_interfaz_notas()
                    messagebox.showinfo("Info", f"Estudiante '{nombre_estudiante}' agregado correctamente.")
            
    def ask_for_valid_float(self, message, error_message, min_value, max_value):
        while True:
            nota = simpledialog.askfloat("Agregar Nota", message)
            if nota is None:
                return None
            elif min_value <= nota <= max_value:
                return nota
            else:
                messagebox.showerror("Error", error_message)
                
    def agregar_nota(self):
        nombre_estudiante = simpledialog.askstring("Agregar Nota", "Ingrese el nombre del estudiante:")
        if nombre_estudiante:
            if nombre_estudiante not in self.nombres:
                messagebox.showwarning("Warning", "El estudiante no está en la lista.")
            else:
                if len(self.notas.columns) == 0:
                    nueva_nota = self.ask_for_valid_float("Ingrese la nueva nota:", "Nota inválida. Debe estar entre 10 y 70.", 10, 70)
                    if nueva_nota is not None:
                        self.notas[nueva_nota] = 0
                else:
                    nueva_nota = self.ask_for_valid_float("Ingrese la nueva nota:", "Nota inválida. Debe estar entre 10 y 70.", 10, 70)
                    if nueva_nota is not None:
                        nueva_columna = max(self.notas.columns) + 1
                        self.notas[nueva_columna] = 0
                self.guardar_notas_en_excel()
                self.actualizar_interfaz_notas()
                messagebox.showinfo("Info", f"Nota para '{nombre_estudiante}' agregada correctamente.")
            
    def editar_nota(self):
        nombre_estudiante = simpledialog.askstring("Editar Nota", "Ingrese el nombre del estudiante:")
        if nombre_estudiante and nombre_estudiante in self.nombres:
            columna = simpledialog.askinteger("Editar Nota", "Ingrese el número de la nota a editar:") - 1
            if columna is not None and 0 <= columna < len(self.notas.columns):
                nueva_nota = self.ask_for_valid_float("Ingrese la nueva nota:", "Nota inválida. Debe estar entre 10 y 70.", 10, 70)
                if nueva_nota is not None:
                    self.notas.loc[nombre_estudiante, self.notas.columns[columna]] = nueva_nota
                    self.guardar_notas_en_excel()
                    self.actualizar_interfaz_notas()
                    messagebox.showinfo("Info", "Nota editada correctamente.")
            else:
                messagebox.showwarning("Warning", "Número de nota inválido.")
        elif nombre_estudiante:
            messagebox.showwarning("Warning", "El estudiante no existe.")
            
    def eliminar_estudiante(self):
        nombre_estudiante = simpledialog.askstring("Eliminar Estudiante", "Ingrese el nombre del estudiante a eliminar:")
        if nombre_estudiante and nombre_estudiante in self.nombres:
            confirmacion = messagebox.askyesno("Eliminar Estudiante", f"¿Está seguro de eliminar a '{nombre_estudiante}'?")
            if confirmacion:
                self.nombres.remove(nombre_estudiante)
                self.notas = self.notas.drop(nombre_estudiante)
                self.guardar_notas_en_excel()
                self.actualizar_interfaz_notas()
                messagebox.showinfo("Info", f"Estudiante '{nombre_estudiante}' eliminado correctamente.")
        elif nombre_estudiante:
            messagebox.showwarning("Warning", "El estudiante no existe.")
            
    def eliminar_nota(self):
        nombre_estudiante = simpledialog.askstring("Eliminar Nota", "Ingrese el nombre del estudiante:")
        if nombre_estudiante and nombre_estudiante in self.nombres:
            columna = simpledialog.askinteger("Eliminar Nota", "Ingrese el número de la nota a eliminar:") - 1
            if columna is not None and 0 <= columna < len(self.notas.columns):
                confirmacion = messagebox.askyesno("Eliminar Nota", f"¿Está seguro de eliminar la nota {columna + 1} de '{nombre_estudiante}'?")
                if confirmacion:
                    self.notas.loc[nombre_estudiante, self.notas.columns[columna]] = None
                    self.guardar_notas_en_excel()
                    self.actualizar_interfaz_notas()
                    messagebox.showinfo("Info", "Nota eliminada correctamente.")
            else:
                messagebox.showwarning("Warning", "Número de nota inválido.")
        elif nombre_estudiante:
            messagebox.showwarning("Warning", "El estudiante no existe.")
            
    def promediar_notas(self):
        if not self.notas.empty:
            promedios = self.notas.mean(axis=1)
            resultados = []
            for nombre, promedio in promedios.items():
                estado = "Aprobado" if promedio >= 4.0 else "Reprobado"
                resultados.append({"Nombre": nombre, "Estado": estado, "Promedio": promedio})
            return resultados
        else:
            messagebox.showwarning("Warning", "No hay notas para promediar.")
            return []

    def mostrar_promedio_notas(self):
        resultados_promedio = self.promediar_notas()
        if resultados_promedio:
            mensaje = "Promedios de Notas:\n"
            for resultado in resultados_promedio:
                promedio_redondeado = round(resultado['Promedio'])
                estado = "Aprobado" if promedio_redondeado >= 40 else "Reprobado"
                mensaje += f"Nombre: {resultado['Nombre']}, Promedio: {promedio_redondeado}, Estado: {estado}\n"
            messagebox.showinfo("Info", mensaje)
        else:
            messagebox.showwarning("Warning", "No hay notas para mostrar promedio.")

    def guardar_notas_en_excel(self):
        if self.filename.endswith('.xlsx'):
            self.notas.to_excel(self.filename, index=True)
            messagebox.showinfo("Info", "Notas guardadas en el archivo Excel.")
        else:
            messagebox.showwarning("Warning", "Seleccione un archivo Excel válido para guardar las notas.")

root = tk.Tk()
app = NotaIngresoApp(root)
root.mainloop()
