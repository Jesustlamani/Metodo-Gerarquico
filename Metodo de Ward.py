import pandas as pd
from scipy.cluster.hierarchy import linkage, dendrogram, fcluster
import matplotlib.pyplot as plt
import openpyxl
import tkinter as tk
from tkinter import filedialog

def cargar_csv():
    # Abrir cuadro de diálogo para seleccionar archivo CSV
    archivo_csv = filedialog.askopenfilename(
        title="Seleccionar archivo CSV", filetypes=[("Archivos CSV", "*.csv")]
    )
    if archivo_csv:
        cargar_datos(archivo_csv)
    else:
        mensaje.set("No se seleccionó ningún archivo CSV.")

def cargar_datos(archivo_csv):
    # Cargar los datos del CSV
    datos = pd.read_csv(archivo_csv)

    # Seleccionar las columnas numéricas relevantes
    columnas_numericas = [
        "Student Age",
        "Sex",
        "Graduated high-school type",
        "scholarship type",
        "additional work",
        "regular artistic or sport activity",
        "do you have a partner",
        "total salary if available",
        "transportation to the university",
        "accomodation type in cyprus",
        "mothersae education",
        "fathersae education",
        "number of sister/brother",
        "parental status",
        "mothersae ocupation",
        "farsae ocupation",
        "weekly study hours",
        "reading frecuency (non-scientify book/journals)",
        "reading frequency (scientific books/journals)",
        "Attendance to the seminars/conferences related to the department",
        "Impact of your projects/activities on your success",
        "Attendance to classes",
        "Preparation to midterm exams 1",
        "Preparation to midterm exams 2",
        "Taking notes in classes",
        "Listening in classes",
        "Discussion improves my interest and success in the course",
        "Flip-classroom",
        "Cumulative grade point average in the last semester",
        "Expected Cumulative grade point average in the graduation",
    ]

    columnas_existentes = datos.columns.tolist()
    columnas_numericas = [
        col for col in columnas_numericas if col in columnas_existentes
    ]
    datos_numericos = datos[columnas_numericas]

    # Calcular la matriz de distancias
    matriz_distancias = datos_numericos.values

    # Aplicar el método de Ward
    Z = linkage(matriz_distancias, method="ward")

    # Graficar el dendrograma
    plt.figure(figsize=(25, 10))
    plt.title("Dendrograma")
    plt.xlabel("Objetos")
    plt.ylabel("Distancia")
    dendrogram(Z, leaf_rotation=90, leaf_font_size=8)

    # Obtener los límites del eje y para determinar la altura máxima
    ylim = plt.ylim()
    altura_maxima = ylim[1]

    # Dibujar las líneas de corte en el dendrograma
    lineas_corte_altura = [20, 15, 10]
    colores = ["r", "g", "b"]
    for i, altura in enumerate(lineas_corte_altura):
        plt.axhline(
            y=altura,
            c=colores[i],
            linestyle="--",
            label=f"Línea de corte {i+1} (t={altura})",
        )

    # Ajustar los límites del eje y para incluir las líneas de corte
    plt.ylim(0, altura_maxima)

    # Agregar una leyenda
    plt.legend()

    # Generar las líneas de corte y los archivos Excel
    for altura in [20, 15, 10]:
        grupo = fcluster(Z, t=altura, criterion="distance")
        datos[f"Grupo_{altura}"] = grupo
        grupos = datos.groupby(f"Grupo_{altura}")

        # Crear un libro de Excel
        escritor = pd.ExcelWriter(f"grupos_corte_{altura}.xlsx", engine="openpyxl")

        # Agregar una hoja para cada grupo
        for nombre_grupo, grupo_datos in grupos:
            grupo_datos.to_excel(
                escritor, sheet_name=f"Grupo_{nombre_grupo}", index=False
            )

        # Guardar el libro de Excel
        escritor._save()
        escritor.close()

        mensaje.set(f"Libros de Excel creados correctamente :D.")

    plt.show()

def salir():
    ventana.destroy()

# Crear ventana principal
ventana = tk.Tk()
ventana.title("AGRUPAMIENTO JERARQUICO")
ventana.geometry("800x600")  # Tamaño de la ventana

# Colores para la GUI
color_fondo = "#FFD700"  # Amarillo
color_botones = "#FF6347"  # Rojo

# Configurar color de fondo
ventana.configure(bg=color_fondo)

# Crear título
etiqueta_titulo = tk.Label(ventana, text="AGRUPAMIENTO JERARQUICO", font=("Arial", 20, "bold"), bg=color_fondo)
etiqueta_titulo.pack(pady=20)

# Crear botón para cargar CSV
boton_cargar_csv = tk.Button(ventana, text="Cargar CSV", command=cargar_csv, bg=color_botones, fg="white", font=("Arial", 14))
boton_cargar_csv.pack(pady=20)

# Crear botón para salir del programa
boton_salir = tk.Button(ventana, text="SALIR", command=salir, bg=color_botones, fg="white", font=("Arial", 14))
boton_salir.pack(pady=10)

# Mensaje para mostrar el resultado
mensaje = tk.StringVar()
etiqueta_mensaje = tk.Label(ventana, textvariable=mensaje, font=("Arial", 12), bg=color_fondo)
etiqueta_mensaje.pack(pady=10)

# Ejecutar el bucle principal de la ventana
ventana.mainloop()