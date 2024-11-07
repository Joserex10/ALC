import tkinter as tk
from tkinter import filedialog, messagebox
import numpy as np
import pandas as pd
from sklearn.decomposition import PCA
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from mpl_toolkits.mplot3d import Axes3D
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.chart import BarChart, Reference

# Función para mostrar los autovalores y autovectores de manera organizada
def format_values(values, is_vector=False):
    if is_vector:
        formatted = '\n'.join(['  '.join([f'{num:.4f}' for num in row]) for row in values])
    else:
        formatted = ', '.join([f'{num:.4f}' for num in values])
    return formatted

# Función para aplicar PCA a un archivo CSV
def apply_pca():
    file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
    if not file_path:
        return
    
    try:
        # Cargar los datos
        data = pd.read_csv(file_path, header=None)

        # Verificar que los datos sean numéricos
        if not np.all(np.isreal(data.values)):
            messagebox.showerror("Error", "El archivo debe contener únicamente datos numéricos para aplicar PCA.")
            return

        # Mostrar los datos originales en un gráfico 3D
        show_original_data_3d(data)
        
        # Aplicar PCA
        pca = PCA(n_components=2)  # Puedes ajustar el número de componentes
        principal_components = pca.fit_transform(data)
        
        # Obtener autovalores y autovectores
        autovalores = pca.explained_variance_
        autovectores = pca.components_
        
        # Mostrar el gráfico de varianza explicada
        plot_explained_variance(pca)
        
        # Crear una nueva ventana para mostrar los resultados después del PCA
        result_window = tk.Toplevel(root)
        result_window.title("Resultados de PCA")
        result_window.geometry("800x700")  # Ajustar el tamaño de la ventana
        result_window.configure(bg=bg_color)
        
        # Crear un cuadro para los autovalores y autovectores
        results_frame = tk.Frame(result_window, bg=bg_color)
        results_frame.pack(pady=20)

        # Formatear los autovalores y autovectores para que se vean mejor
        autovalores_str = "Autovalores:\n" + format_values(autovalores)
        autovectores_str = "Autovectores:\n" + format_values(autovectores, is_vector=True)
        
        # Mostrar los autovalores y autovectores con un fondo oscuro y fuente clara
        text_box = tk.Text(results_frame, wrap=tk.WORD, width=90, height=10, font=("Courier New", 12), bg=bg_color, fg=text_color, bd=0, highlightthickness=0)
        text_box.insert(tk.END, autovalores_str + "\n\n" + autovectores_str)
        text_box.config(state=tk.DISABLED)  # Hacer que el cuadro de texto sea de solo lectura
        text_box.pack(pady=10)
        
        # Gráfica de los componentes principales (datos transformados por PCA)
        fig, ax = plt.subplots(figsize=(8, 6))
        scatter = ax.scatter(principal_components[:, 0], principal_components[:, 1], 
                            c=['red' if x > 0 else 'blue' for x in principal_components[:, 0]], edgecolor='k', s=50)
        ax.set_title("Datos proyectados en los dos primeros componentes principales", fontsize=14)
        ax.set_xlabel("Componente Principal 1 (PC1)", fontsize=12)
        ax.set_ylabel("Componente Principal 2 (PC2)", fontsize=12)
        ax.grid(True)

        # Integrar la gráfica en Tkinter
        canvas = FigureCanvasTkAgg(fig, master=result_window)
        canvas.draw()
        canvas.get_tk_widget().pack(pady=10, fill=tk.BOTH, expand=True)
        
        # Botón para guardar los datos transformados
        save_btn = tk.Button(result_window, text="Guardar Datos Transformados", font=("Arial", 12), bg=btn_color, fg=text_color, command=lambda: save_transformed_data(principal_components, autovalores, autovectores, pca))
        save_btn.pack(pady=20)
        
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error al aplicar PCA:\n{e}")

# Función para mostrar los datos originales en 3D
def show_original_data_3d(data):
    original_window = tk.Toplevel(root)
    original_window.title("Datos Originales - Gráfico 3D")
    original_window.geometry("800x600")
    
    fig = plt.figure(figsize=(8, 6))
    ax = fig.add_subplot(111, projection='3d')
    
    # Asumimos que los datos tienen al menos 3 columnas para el gráfico 3D
    ax.scatter(data.iloc[:, 0], data.iloc[:, 1], data.iloc[:, 2], c='blue', marker='o', s=50)
    ax.set_title("Datos Originales en 3D")
    ax.set_xlabel("Variable 1")
    ax.set_ylabel("Variable 2")
    ax.set_zlabel("Variable 3")

    # Integrar la gráfica 3D en Tkinter
    canvas = FigureCanvasTkAgg(fig, master=original_window)
    canvas.draw()
    canvas.get_tk_widget().pack(pady=10, fill=tk.BOTH, expand=True)

# Función para mostrar el gráfico de varianza explicada con porcentajes en las barras
def plot_explained_variance(pca):
    explained_variance = pca.explained_variance_ratio_ * 100  # Convertir a porcentaje
    components = np.arange(1, len(explained_variance) + 1)
    
    fig, ax = plt.subplots(figsize=(8, 6))
    bars = ax.bar(components, explained_variance, color='cyan')
    ax.set_title('Varianza Explicada por Cada Componente')
    ax.set_xlabel('Componente Principal')
    ax.set_ylabel('Varianza Explicada (%)')
    ax.set_xticks(components)
    
    # Añadir los valores de porcentaje sobre las barras
    for bar in bars:
        height = bar.get_height()
        ax.annotate(f'{height:.2f}%',  # Mostrar el valor formateado como porcentaje
                    xy=(bar.get_x() + bar.get_width() / 2, height),
                    xytext=(0, 3),  # Desplazar el texto hacia arriba
                    textcoords="offset points",
                    ha='center', va='bottom')

    # Mostrar gráfico en ventana separada
    var_window = tk.Toplevel(root)
    var_window.title("Varianza Explicada")
    canvas = FigureCanvasTkAgg(fig, master=var_window)
    canvas.draw()
    canvas.get_tk_widget().pack(pady=10, fill=tk.BOTH, expand=True)


# Función para guardar los datos transformados y mejorar el archivo Excel con gráficos
def save_transformed_data(principal_components, autovalores, autovectores, pca):
    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
    if save_path:
        # Crear DataFrame con PC1, PC2 y los autovalores/autovectores
        transformed_df = pd.DataFrame(np.round(principal_components, 3), columns=['PC1', 'PC2'])
        transformed_df['Autovalores'] = np.nan
        transformed_df.loc[0, 'Autovalores'] = autovalores[0]
        transformed_df.loc[1, 'Autovalores'] = autovalores[1]
        
        # Crear archivo Excel con gráficos
        with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
            transformed_df.to_excel(writer, index=False, sheet_name='Datos Transformados')
            
            # Hoja para los autovectores
            autovectores_df = pd.DataFrame(autovectores, columns=[f"Componente {i+1}" for i in range(autovectores.shape[1])])
            autovectores_df.to_excel(writer, index=False, sheet_name='Autovectores')
            
            # Crear una hoja para la varianza explicada
            explained_variance_df = pd.DataFrame(pca.explained_variance_ratio_, columns=["Varianza Explicada"])
            explained_variance_df.to_excel(writer, sheet_name='Varianza Explicada', index_label="Componente")
        
        # Añadir gráfico de barras a la hoja de varianza explicada en Excel
        wb = openpyxl.load_workbook(save_path)
        sheet = wb["Varianza Explicada"]
        chart = BarChart()
        data = Reference(sheet, min_col=2, min_row=1, max_row=sheet.max_row)
        chart.add_data(data, titles_from_data=True)
        chart.title = "Varianza Explicada por Componente"
        chart.x_axis.title = "Componente Principal"
        chart.y_axis.title = "Varianza Explicada (%)"
        sheet.add_chart(chart, "E5")
        wb.save(save_path)
        
        messagebox.showinfo("Guardado", f"Datos transformados y gráficos guardados en: {save_path}")

# Crear la ventana principal
root = tk.Tk()
root.title("Analizador de Archivos - PCA")
root.geometry("500x500")

# Colores vibrantes
bg_color = "#1e1e1e"  # Fondo más oscuro
btn_color = "#5bc0de"  # Azul vibrante
text_color = "#f8f9fa"  # Blanco claro

# Configurar el fondo
root.configure(bg=bg_color)

# Título
title_label = tk.Label(root, text="Analizador de Archivos - PCA", font=("Helvetica", 18, "bold"), bg=bg_color, fg=text_color)
title_label.pack(pady=50)

# Botón para aplicar PCA
pca_btn = tk.Button(root, text="Aplicar PCA a Archivo CSV", font=("Arial", 14), bg=btn_color, fg=text_color, padx=10, pady=5)
pca_btn.config(bd=0, relief=tk.FLAT)
pca_btn.config(command=apply_pca)
pca_btn.pack(padx=10, pady=50)

root.mainloop()
