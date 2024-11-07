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

        if data.shape[1] < 3:
            messagebox.showerror("Error", "El archivo debe tener al menos 3 columnas para graficar en 3D.")
            return
        
        # Mostrar los datos originales en un gráfico 3D
        show_original_data_3d(data)

        # Selección de número de componentes por el usuario
        n_components = select_components(data.shape[1])
        if n_components is None:
            return
        
        # Aplicar PCA
        pca = PCA(n_components=n_components)
        principal_components = pca.fit_transform(data)
        
        # Obtener autovalores y autovectores
        autovalores = pca.explained_variance_
        autovectores = pca.components_
        
        # Mostrar el gráfico de varianza explicada
        plot_explained_variance(pca)
        
        # Crear una nueva ventana para mostrar los resultados después del PCA
        result_window = tk.Toplevel(root)
        result_window.title("Resultados de PCA")
        result_window.geometry("800x700")  
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
        text_box.config(state=tk.DISABLED)  
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