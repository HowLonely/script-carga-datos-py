import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from tkcalendar import DateEntry
import pandas as pd
from datetime import datetime
from tkinter import ttk

import xlwings as xw

# Variables globales para las rutas de los archivos
path_archivo_forms = ""
path_archivo_final = ""
hoja_final = "Listado SKU (2808)"

# Variables globales para columnas necesarias
columnas_necesarias_form = ["Hora de inicio", "¿Cual es tu nombre?", "¿Cual es el código de stock?", "¿Material es encontrado?", "¿Caja Cerrada?",
                              "¿Descripción Extendida del material?", "¿Número de Parte?, Escríbalo tal como se desarrolla en el componente físico.",
                              "¿Fabricante?", "¿Modelo? Detalle modelo de componente.", "¿Cantidad Contabilizada? Información relacionada a inventario realizado.",
                              "Comentario Adicional"]

columnas_necesarias_final = []
df_final = None
df_columns_final = []
relaciones = {}

diccionario_mapeo = {}

# Función para seleccionar el archivo de formularios
def seleccionar_archivo_forms():
    global path_archivo_forms
    archivo = filedialog.askopenfilename(
        title="Seleccionar archivo de formularios",
        filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")]
    )
    if archivo:
        path_archivo_forms = archivo
        archivo_forms_label.config(text=f"Archivo de formularios: {archivo}")
    else:
        messagebox.showwarning("Selección de archivo", "No se ha seleccionado ningún archivo.")

# Función para seleccionar el archivo final
def seleccionar_archivo_final():
    global path_archivo_final
    archivo = filedialog.askopenfilename(
        title="Seleccionar archivo final",
        filetypes=[("Excel Files", "*.xlsm"), ("All Files", "*.*")]
    )
    if archivo:
        path_archivo_final = archivo
        archivo_final_label.config(text=f"Archivo final: {archivo}")

        # Cargar el archivo final después de seleccionar
        global df_final, df_columns_final
        df_final = pd.read_excel(path_archivo_final, sheet_name=hoja_final, header=1)  # header = 1 ya que aquí están los nombres de columnas
        df_columns_final = df_final.columns.to_list()


        global columnas_necesarias_final

        columnas_necesarias_final = ["Status Planificacion (KDM)", "Turno (KDM)", "Semana (KDM)", "Responsable (KDM)", "Fecha (KDM)", "Meses"
                                        "SKU", "Ubicacion", "Número de Parte", "Fabricante (s)", "Descripción Extendida", "Modelo", 
                                        "Cantidad Contabilizada", "Comentarios adicionales ", "CAJA CERRADA?", "Status KDM"]

        # Relaciones
        global relaciones
        relaciones = {
            "caja_cerrada": ["¿Caja Cerrada?", "CAJA CERRADA?"],
            "desc_ext": ["¿Descripción Extendida del material?", "¿Número de Parte?, Escríbalo tal como se desarrolla en el componente físico.", 
                         "¿Fabricante?", "Descripción Extendida"],
            "n_parte": ["¿Número de Parte?, Escríbalo tal como se desarrolla en el componente físico.", "Número de Parte"],
            "fabricante": ["¿Fabricante?", "Fabricante (s)"],
            "modelo": ["¿Modelo? Detalle modelo de componente.", "Modelo"],
            "cant_contab": ["¿Cantidad Contabilizada? Información relacionada a inventario realizado.", "Cantidad Contabilizada"],
            "coment_adicional": ["Comentario Adicional", "Comentarios adicionales "],
            "contado": ["¿Material es encontrado?", "Status KDM"],
            "responsable": ["¿Cual es tu nombre?", "Responsable (KDM)"],
            "fecha": ["Hora de inicio", "Fecha (KDM)"],
            #FALTA MESES IMPORTANTE
        }
    else:
        messagebox.showwarning("Selección de archivo", "No se ha seleccionado ningún archivo.")

# Función para procesar el archivo de formularios
def proceso_form(fecha_obj):
    global columnas_necesarias_form

    # Cargar el archivo de formularios
    df_form = pd.read_excel(path_archivo_forms)
    df_columns_form = df_form.columns.tolist()
    
    # Convertir la columna 'Hora de inicio' a datetime para comparar solo la fecha
    df_form['Hora de inicio'] = pd.to_datetime(df_form[df_columns_form[1]], errors='coerce')

    # Filtrar solo las filas que coinciden con la fecha ingresada
    df_filtrado = df_form[df_form['Hora de inicio'].dt.date == fecha_obj.date()]

    # Verificar si se encontraron resultados
    if not df_filtrado.empty:
        # Ordenar por 'Hora de inicio' de antiguo a más reciente
        df_filtrado = df_filtrado.sort_values(by='Hora de inicio', ascending=True)

        # Seleccionar solo las columnas necesarias
        registros_necesarios = df_filtrado[df_filtrado.columns[df_filtrado.columns.isin(columnas_necesarias_form)]]

        #print(f"Registros encontrados para la fecha {fecha_obj.strftime('%d-%m-%Y')} ordenados de antiguo a reciente:")

        # Convertir a lista de diccionarios para mayor claridad en el resultado
        registros_list_form = registros_necesarios.to_dict('records')


        for registro in registros_list_form:
            print("[DEBUG FORMS]", registro)
            print("")

        return registros_list_form
    else:
        print(f"[ERROR] No se encontraron registros para la fecha {fecha_obj.strftime('%d-%m-%Y')}.")
        return []

# Función para crear el diccionario de mapeo
def crear_diccionario_mapeo(registros):
    mapeo = {}

    for registro in registros:
        sku = registro.get("¿Cual es el código de stock?")
        if sku is None:
            continue

        print
        
        fila_final = df_final[df_final["SKU"] == sku]
        if fila_final.empty:
            continue
        
        fila_index = fila_final.index[0]

        mapeo[sku] = {}

        
        for clave, valores in relaciones.items():
            columnas_origen = valores[:-1]
            columna_destino = valores[-1]

            datos_origen = [registro.get(col) for col in columnas_origen if col in registro]

            if clave == "desc_ext":
                mapeo[sku][df_columns_final[df_columns_final.index(columna_destino)]] = ' '.join(str(dato) for dato in datos_origen if pd.notna(dato))
            elif clave == "fecha":
                fecha_valor = registro.get("Hora de inicio")
                
                if pd.notna(fecha_valor):
                    # Convertir la fecha al formato DD/MM/YYYY
                    fecha_formateada = fecha_valor.strftime('%d/%m/%Y')
                    mapeo[sku][df_columns_final[df_columns_final.index(columna_destino)]] = fecha_formateada
                else:
                    mapeo[sku][df_columns_final[df_columns_final.index(columna_destino)]] = None
            elif clave == "contado":
                
                # Crear un diccionario de mapeo para facilitar la conversión
                mapeo_contado = {
                    "SI": "Contado",
                    "NO": "No Contado",
                    "Contado": "Contado",
                    "Buscado/No encontrado": "Buscado,No Encontrado"
                }

                contado_form_valor = registro.get("¿Material es encontrado?")

                if contado_form_valor in mapeo_contado:
                    mapeo[sku][df_columns_final[df_columns_final.index(columna_destino)]] = mapeo_contado[contado_form_valor]
                else:
                    mapeo[sku][df_columns_final[df_columns_final.index(columna_destino)]] = "No Contado"  # Valor por defecto si no se encuentra en el mapeo

            else:
                mapeo[sku][df_columns_final[df_columns_final.index(columna_destino)]] = datos_origen[0] if datos_origen else None

        # Copiar el valor de la columna "Stor. Bin" a la columna "Ubicación"
        stor_bin_val = df_final.at[fila_index, "Stor. Bin"]
        # Agregar el valor de la columna "Ubicación" al mapeo
        mapeo[sku][df_columns_final[df_columns_final.index("Ubicacion")]] = stor_bin_val

    
    #print(f"mapeo: {mapeo}" )
    return mapeo

# Función para construir y mostrar la tabla en la interfaz gráfica
def construir_tabla(diccionario_mapeo):
    global tree, columnas

    # Eliminar cualquier fila existente en el treeview
    for item in tree.get_children():
        tree.delete(item)

    if diccionario_mapeo:
        # Agregar la columna SKU
        columnas = ["SKU"] + sorted(set(columna for valor in diccionario_mapeo.values() for columna in valor.keys()))
        tree["columns"] = columnas

        # Crear encabezados de columnas
        for col in columnas:
            tree.heading(col, text=col)
            tree.column(col, width=100, anchor="w")

        # Insertar datos en la tabla
        for sku, datos in diccionario_mapeo.items():
            valores = [sku] + [datos.get(col, "") for col in columnas[1:]]
            tree.insert("", "end", values=valores)
        
        tree.bind("<Double-1>", on_double_click)


def on_double_click(event):
    item = tree.selection()
    if not item:
        return
    item = item[0]
    
    column_id = tree.identify_column(event.x)
    col_index = int(column_id.split('#')[-1]) - 1  # Convertir a índice de columna (0 basado)
    
    if col_index == 0:  # Si es la columna SKU (no editable)
        return

    # Obtener el valor actual de la celda
    viejo_valor = tree.item(item, "values")[col_index]
    
    # Mostrar un cuadro de diálogo para ingresar el nuevo valor
    nuevo_valor = simpledialog.askstring("Editar valor", "Nuevo valor:", initialvalue=viejo_valor)
    
    if nuevo_valor is not None:
        # Actualizar el valor en el Treeview
        valores = list(tree.item(item, "values"))
        valores[col_index] = nuevo_valor
        tree.item(item, values=valores)
        
        # Actualizar el diccionario de mapeo
        sku = float(valores[0])
        columna = columnas[col_index]
        
        if sku in diccionario_mapeo:
            #print("editado sku: ", sku)
            diccionario_mapeo[sku][columna] = nuevo_valor
        else:
            # Crear un nuevo registro si el SKU no existe, lo cual debería ser raro
            diccionario_mapeo[sku] = {columna: nuevo_valor}

# Función para manejar la eliminación de un registro
def eliminar_registro():
    item = tree.selection()
    if not item:
        messagebox.showwarning("Eliminar registro", "No se ha seleccionado ningún registro para eliminar.")
        return
    
    item = item[0]
    
    # Obtener el SKU del registro seleccionado
    sku = float(tree.item(item, "values")[0])
    
    # Eliminar el registro del diccionario mapeo
    if sku in diccionario_mapeo:
        del diccionario_mapeo[sku]
    
    # Eliminar el registro del Treeview
    tree.delete(item)
    
    messagebox.showinfo("Eliminar registro", "El registro ha sido eliminado exitosamente.")


def ejecutar_mapeo_consola():
    # Imprimir el mapeo
    for sku, valores in diccionario_mapeo.items():
        print(f"[DEBUG] SKU: {sku}")
        for col_final, valor in valores.items():
            print(f"  {col_final}: {valor}")

def ejecutar_proceso():
    global path_archivo_forms, path_archivo_final, diccionario_mapeo

    if not path_archivo_forms or not path_archivo_final:
        messagebox.showwarning("Error", "Debe seleccionar ambos archivos.")
        return

    fecha_str = fecha_entry.get()
    try:
        fecha_obj = datetime.strptime(fecha_str, '%d/%m/%Y')
    except ValueError:
        messagebox.showerror("Error", "Formato de fecha inválido. Use DD/MM/YYYY.")
        return

    registros = proceso_form(fecha_obj)
    diccionario_mapeo = crear_diccionario_mapeo(registros)
    construir_tabla(diccionario_mapeo)

def actualizar_registros():
    global diccionario_mapeo, df_final, df_columns_final, path_archivo_final, hoja_final
    
    # Verificar que haya datos para actualizar
    if not diccionario_mapeo or df_final is None:
        messagebox.showwarning("Actualizar registros", "No hay datos para actualizar.")
        return

    # Abrir el archivo Excel con xlwings
    try:
        wb = xw.Book(path_archivo_final)
        sheet = wb.sheets[hoja_final]
    except Exception as e:
        messagebox.showerror("Error", f"Error al abrir el archivo: {e}")
        return

    # Iterar sobre el diccionario de mapeo para realizar actualizaciones
    for sku, datos_mapeo in diccionario_mapeo.items():
        # Verificar si el SKU existe en el DataFrame
        idx = df_final[df_final[df_columns_final[30]] == sku].index

        print(f"[ACTUALIZACION SKU: {sku}]")

        if not idx.empty:
            print("")
            fila_index = idx[0]
            fila_excel = fila_index + 2  # Ajuste según encabezado
            
            # Iterar sobre las columnas a actualizar para este SKU
            for columna, valor in datos_mapeo.items():
                if columna in df_columns_final and pd.notna(valor) and valor != '-' and valor != '':
                    col_index = df_columns_final.index(columna)

                    # Actualizar la celda en Excel
                    print(f"'{sku}', Columna '{columna} Columna idx {col_index + 1}' en la fila {fila_excel + 1}")
                    print(f"[Antiguo]  con valor: {sheet.cells[int(fila_excel), int(col_index)].value}")
                    sheet.cells[int(fila_excel), int(col_index)].value = valor
                    print(f"[Actualizado] '{sku}', Columna '{columna} Columna idx {col_index}' en la fila {fila_excel + 1} con valor: {valor}")

    # Guardar el libro de trabajo actualizado
    try:
        wb.save()
        wb.close()
        messagebox.showinfo("Actualizar registros", "Los registros se han actualizado exitosamente.")
    except Exception as e:
        messagebox.showerror("Error", f"Error al guardar el archivo: {e}")



# Configuración de la ventana principal
root = tk.Tk()
root.title("Interfaz de Carga de Datos")

# Etiquetas y botones para seleccionar archivos
archivo_forms_label = tk.Label(root, text="Archivo de formularios: No seleccionado")
archivo_forms_label.pack()

btn_seleccionar_forms = tk.Button(root, text="Seleccionar archivo de formularios", command=seleccionar_archivo_forms)
btn_seleccionar_forms.pack()

archivo_final_label = tk.Label(root, text="Archivo final: No seleccionado")
archivo_final_label.pack()

btn_seleccionar_final = tk.Button(root, text="Seleccionar archivo final", command=seleccionar_archivo_final)
btn_seleccionar_final.pack()

# Campo de entrada para la fecha
tk.Label(root, text="Ingrese la fecha (DD/MM/YYYY):").pack()
fecha_entry = tk.Entry(root)
fecha_entry.pack()

# Botón para ejecutar el proceso
btn_ejecutar = tk.Button(root, text="Ejecutar Proceso", command=ejecutar_proceso)
btn_ejecutar.pack()

# Tabla para mostrar los resultados
tree = ttk.Treeview(root, show="headings")
tree.pack(fill="both", expand=True)

btn_mapeo_consola = tk.Button(root, text="Mapeo Consola", command=ejecutar_mapeo_consola)
btn_mapeo_consola.pack()

btn_eliminar = tk.Button(root, text="Eliminar registro", command=eliminar_registro)
btn_eliminar.pack()

# Añadir un botón para actualizar los registros
btn_actualizar = tk.Button(root, text="Actualizar registros", command=actualizar_registros)
btn_actualizar.pack()

# Ejecutar la interfaz gráfica
root.mainloop()


# RECOMENDACIONES FINALES
# UTILIZAR COMO ARCHIVO MAESTRO, UNA COPIA DEL ARCHIVO MAESTRO ORIGINAL EN CASO DE ERRORES
# CUANDO SE AGREGEN COLUMNAS O MODIFIQUEN TENER MAS RELEVANCIA A LO ANTERIOR
# EN CASO DE CAMBIO DE NOMBRE DE UNA COLUMNA AVISARME