import pandas as pd
from datetime import timedelta
import re
import warnings
import unicodedata
import functools
from itertools import groupby
from operator import itemgetter
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import logging
import os

# Configurar logging
logging.basicConfig(level=logging.INFO)

# Ignorar advertencias de futuras versiones
warnings.simplefilter(action='ignore', category=FutureWarning)

# 1. Función para convertir tiempo a timedelta
@functools.lru_cache(maxsize=None)
def time_to_timedelta(time_str):
    try:
        parts = time_str.split(':')
        if len(parts) == 4:
            hours, minutes, seconds, frames = map(int, parts)
        elif len(parts) == 3:
            hours, minutes, seconds = map(int, parts)
            frames = 0
        else:
            raise ValueError("Formato de tiempo incorrecto")
        total_seconds = hours * 3600 + minutes * 60 + seconds + frames / 24
        return timedelta(seconds=total_seconds)
    except Exception as e:
        logging.error(f"Error al convertir tiempo: {time_str} - {e}")
        return timedelta(0)

# 2. Dividir diálogos que excedan los 60 caracteres (excluyendo contenido entre paréntesis)
def dividir_dialogo(dialogo, max_caracteres=60):
    dialogo_sin_parentesis = re.sub(r'\(.*?\)', '', dialogo)
    if len(dialogo_sin_parentesis) <= max_caracteres:
        return [dialogo]

    palabras = dialogo.split()
    lineas = []
    linea_actual = ''
    for palabra in palabras:
        temp = f"{linea_actual} {palabra}".strip()
        temp_sin_parentesis = re.sub(r'\(.*?\)', '', temp)
        if len(temp_sin_parentesis) > max_caracteres:
            if linea_actual:
                lineas.append(linea_actual)
            linea_actual = palabra
        else:
            linea_actual = temp
    if linea_actual:
        lineas.append(linea_actual)
    return lineas

# 3. Expandir diálogos
def expandir_dialogos(df_original):
    nuevas_filas = []
    for idx, row in df_original.iterrows():
        lineas_divididas = dividir_dialogo(row['DIÁLOGO'])
        for linea in lineas_divididas:
            new_row = row.copy()
            new_row['DIÁLOGO'] = linea
            nuevas_filas.append(new_row)
    return pd.DataFrame(nuevas_filas)

# 4. Limpiar columnas de texto
def clean_text(text):
    if isinstance(text, str):
        text = ''.join(ch for ch in text if unicodedata.category(ch)[0] != 'C')
        return text
    else:
        return text

# 5. Optimizar la división de *takes* en una escena
def optimizar_takes_escena(intervenciones_escena, max_duracion_take=30, max_lineas_take=10, max_lineas_consecutivas=5, max_lineas_por_personaje=5):
    intervenciones = []
    for idx, row in intervenciones_escena.iterrows():
        intervenciones.append({
            'idx': idx,
            'in_td': row['in_td'],
            'out_td': row['out_td'],
            'duracion': row['duracion'],
            'personaje': row['PERSONAJE'],
            'dialogo': row['DIÁLOGO'],
            'IN': row['IN'],
            'OUT': row['OUT'],
            'SCENE': row['SCENE']
        })

    if not intervenciones:
        return []

    intervenciones_sorted = sorted(intervenciones, key=lambda x: (x['in_td'], x['out_td']))
    bloques = [list(group) for key, group in groupby(intervenciones_sorted, key=lambda x: (x['in_td'], x['out_td']))]
    n = len(bloques)

    @functools.lru_cache(maxsize=None)
    def dp(pos):
        if pos >= n:
            return [], (0, 0)

        best_takes = None
        best_cost = None

        for end in range(pos + 1, n + 1):
            take_bloques = bloques[pos:end]
            take_intervenciones = [intervencion for bloque in take_bloques for intervencion in bloque]

            duracion_take = take_intervenciones[-1]['out_td'] - take_intervenciones[0]['in_td']
            if duracion_take.total_seconds() > max_duracion_take:
                break

            if len(take_intervenciones) > max_lineas_take:
                break

            valid = True
            personaje_lineas_consecutivas = {}
            personaje_lineas_totales = {}
            for i, intervencion in enumerate(take_intervenciones):
                personaje = intervencion['personaje']
                # Contar líneas totales por personaje
                personaje_lineas_totales[personaje] = personaje_lineas_totales.get(personaje, 0) + 1
                if personaje_lineas_totales[personaje] > max_lineas_por_personaje:
                    valid = False
                    break
                # Contar líneas consecutivas por personaje
                if i == 0 or personaje != take_intervenciones[i - 1]['personaje']:
                    personaje_lineas_consecutivas[personaje] = 1
                else:
                    personaje_lineas_consecutivas[personaje] += 1
                if personaje_lineas_consecutivas[personaje] > max_lineas_consecutivas:
                    valid = False
                    break
            if not valid:
                continue

            next_takes, next_cost = dp(end)

            personajes_en_take = set([intervencion['personaje'] for intervencion in take_intervenciones])
            total_takes_por_personaje = next_cost[0] + len(personajes_en_take)
            total_takes = next_cost[1] + 1

            current_cost = (total_takes_por_personaje, total_takes)

            if best_cost is None or current_cost < best_cost:
                best_takes = [tuple(take_intervenciones)] + list(next_takes)
                best_cost = current_cost

        if best_takes is None:
            return [], (float('inf'), float('inf'))

        return best_takes, best_cost

    best_takes, _ = dp(0)

    takes = []
    take_id = 1
    for take_intervenciones in best_takes:
        take_intervenciones = list(take_intervenciones)
        takes.append({
            'take': take_id,
            'in': take_intervenciones[0]['in_td'],
            'out': take_intervenciones[-1]['out_td'],
            'scene': take_intervenciones[0]['SCENE'],
            'lineas': take_intervenciones
        })
        take_id += 1

    return takes

# 6. Asignar *takes* optimizados
def asignar_takes_optimizado(df):
    df = df.reset_index(drop=True)

    escenas = df['SCENE'].unique()
    all_takes = []
    take_global_id = 1

    for escena in escenas:
        intervenciones_escena = df[df['SCENE'] == escena]
        takes_escena = optimizar_takes_escena(intervenciones_escena)
        for take in takes_escena:
            take['take'] = take_global_id
            take_global_id += 1
            all_takes.append(take)

    take_list = []
    for take in all_takes:
        for linea in take['lineas']:
            take_list.append({
                'TAKE': take['take'],
                'IN': linea['IN'],
                'OUT': linea['OUT'],
                'PERSONAJE': linea['personaje'],
                'DIÁLOGO': linea['dialogo'],
                'DURACIÓN': linea['duracion'],
                'SCENE': take['scene']
            })

    return pd.DataFrame(take_list)

# 7. Calcular el total de *takes* por personaje
def calcular_total_takes_por_personaje(df_takes):
    takes_por_personaje = df_takes.groupby('PERSONAJE')['TAKE'].nunique().reset_index()
    takes_por_personaje.rename(columns={'TAKE': 'TOTAL_TAKES'}, inplace=True)
    suma_total_takes = takes_por_personaje['TOTAL_TAKES'].sum()
    return takes_por_personaje, suma_total_takes

# 8. Leer archivo
def leer_archivo(file_path):
    try:
        return pd.read_excel(file_path)
    except FileNotFoundError:
        messagebox.showerror("Error", f"El archivo '{file_path}' no se encontró.")
    except Exception as e:
        logging.error(f"Error al leer el archivo Excel: {e}")
    return None

# 9. Función para transformar Excel a TXT
def transformar_excel_a_txt(ruta_excel, ruta_salida_txt):
    # Obtener el nombre del archivo sin la extensión y en mayúsculas
    nombre_archivo = os.path.splitext(os.path.basename(ruta_excel))[0].upper()

    # Leer datos del Excel
    df = pd.read_excel(ruta_excel)

    # Asegurarnos de que las columnas necesarias existen
    columnas_necesarias = ["TAKE", "IN", "OUT", "PERSONAJE", "DIÁLOGO", "DURACIÓN", "SCENE"]
    for col in columnas_necesarias:
        if col not in df.columns:
            raise ValueError(f"Falta la columna requerida: {col}")

    # Convertir todo en texto para prevenir problemas con valores no string
    df["DIÁLOGO"] = df["DIÁLOGO"].astype(str)

    # Reemplazar ":" por " " en IN y OUT
    df["IN"] = df["IN"].str.replace(":", " ", regex=False)
    df["OUT"] = df["OUT"].str.replace(":", " ", regex=False)

    # Agrupar por TAKE
    agrupado_por_take = df.groupby("TAKE")

    with open(ruta_salida_txt, "w", encoding="utf-8") as archivo_salida:
        # Escribir el nombre del archivo al principio
        archivo_salida.write(f"{nombre_archivo}\n\n")

        for take, grupo in agrupado_por_take:
            archivo_salida.write(f"TAKE {take}\n")
            archivo_salida.write(f"{grupo.iloc[0]['IN']}\n")  # TC de IN

            dialogo_actual = ""  # Para acumular el diálogo combinado
            personaje_actual = None

            for idx, fila in grupo.iterrows():
                personaje = fila["PERSONAJE"]
                dialogo = formatear_dialogo(fila["DIÁLOGO"], acumulado=True)

                # Si el personaje actual es el mismo, acumular diálogo con espacio
                if personaje == personaje_actual:
                    dialogo_actual += f" {dialogo}"  # Concatenar con un espacio
                else:
                    # Si hay un personaje previo, escribir su diálogo acumulado
                    if personaje_actual:
                        archivo_salida.write(f"{personaje_actual}:\t{dialogo_actual}\n")

                    # Cambiar al nuevo personaje y reiniciar diálogo acumulado
                    personaje_actual = personaje
                    dialogo_actual = dialogo

            # Escribir el último diálogo acumulado del TAKE
            if personaje_actual:
                archivo_salida.write(f"{personaje_actual}:\t{dialogo_actual}\n")

            # Escribir el OUT del último diálogo del TAKE
            archivo_salida.write(f"{grupo.iloc[-1]['OUT']}\n\n")

    print(f"Archivo de texto generado en: {ruta_salida_txt}")

# Función auxiliar para formatear diálogos
def formatear_dialogo(dialogo, tab_size=4, acumulado=False):
    """
    Formatea el diálogo. Si es acumulado, no se añaden '<<', solo se concatenan las líneas con un espacio.
    """
    dialogo = str(dialogo).replace("“", '"').replace("”", '"')  # Reemplaza comillas tipográficas por comillas rectas
    if acumulado:
        return dialogo.replace("\n", " ")  # Reemplaza saltos de línea por espacios
    else:
        lineas = dialogo.split("\n")
        tab = " " * tab_size
        if len(lineas) == 1:
            return lineas[0]
        else:
            return lineas[0] + "\n" + "\n".join([f"{tab}<< {linea}" for linea in lineas[1:]])

# 10. Procesar archivo
def procesar_archivo(file_path, selected_personajes, status_label, window, process_button):
    def update_status(text):
        window.after(0, lambda: status_label.config(text=text))

    def show_info(title, message):
        window.after(0, lambda: messagebox.showinfo(title, message))

    def show_error(title, message):
        window.after(0, lambda: messagebox.showerror(title, message))

    def enable_process_button():
        window.after(0, lambda: process_button.config(state=tk.NORMAL))

    df = leer_archivo(file_path)
    if df is None:
        enable_process_button()
        return

    columnas_necesarias = {'IN', 'OUT', 'PERSONAJE', 'DIÁLOGO', 'SCENE'}
    columnas_faltantes = columnas_necesarias - set(df.columns)
    if columnas_faltantes:
        show_error("Error", f"Faltan las siguientes columnas: {', '.join(columnas_faltantes)}")
        enable_process_button()
        return

    # Filtrar los personajes seleccionados
    df = df[df['PERSONAJE'].isin(selected_personajes)].reset_index(drop=True)

    update_status("Convirtiendo tiempos...")
    df['in_td'] = df['IN'].apply(time_to_timedelta)
    df['out_td'] = df['OUT'].apply(time_to_timedelta)
    df['duracion'] = (df['out_td'] - df['in_td']).dt.total_seconds()
    df = df.sort_values(by=['in_td', 'out_td']).reset_index(drop=True)

    update_status("Dividiendo diálogos largos...")
    df = expandir_dialogos(df)

    update_status("Limpiando texto...")
    df['DIÁLOGO'] = df['DIÁLOGO'].apply(clean_text)
    df['PERSONAJE'] = df['PERSONAJE'].apply(clean_text)

    update_status("Asignando *takes* optimizados...")
    df_prop_optimizada = asignar_takes_optimizado(df)

    df_prop_optimizada['DURACIÓN'] = df_prop_optimizada['DURACIÓN'].astype(float)

    update_status("Calculando resumen de *takes* por personaje...")
    takes_por_personaje_optimizada, suma_total_takes_optimizada = calcular_total_takes_por_personaje(df_prop_optimizada)

    # Obtener el nombre base del archivo de entrada
    base_name = os.path.splitext(os.path.basename(file_path))[0]

    # Crear nombres de archivos de salida
    output_excel = f"{base_name}_TAKEO.xlsx"
    output_txt = f"{base_name}_DIALOG.txt"

    update_status(f"Exportando a Excel '{output_excel}'...")
    try:
        with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
            df_prop_optimizada.to_excel(writer, sheet_name='Optimizada_Takes', index=False)
            takes_por_personaje_optimizada.to_excel(writer, sheet_name='Resumen', index=False)
            workbook  = writer.book
            worksheet = writer.sheets['Resumen']
            last_row = len(takes_por_personaje_optimizada) + 1
            worksheet.write(f'A{last_row + 1}', 'Suma total de Takes:')
            worksheet.write(f'B{last_row + 1}', suma_total_takes_optimizada)
        update_status(f"Exportación a Excel completada: '{output_excel}'")
    except Exception as e:
        logging.error(f"Error al exportar a Excel: {e}")
        show_error("Error", f"Error al exportar a Excel: {e}")
        enable_process_button()
        return

    # Transformar Excel a TXT
    update_status(f"Transformando Excel a TXT '{output_txt}'...")
    try:
        transformar_excel_a_txt(output_excel, output_txt)
        update_status(f"Exportación a TXT completada: '{output_txt}'")
        show_info("Éxito", f"La propuesta optimizada y su resumen han sido exportados a '{output_excel}'\nEl archivo de diálogo ha sido generado: '{output_txt}'")
    except Exception as e:
        logging.error(f"Error al transformar Excel a TXT: {e}")
        show_error("Error", f"Error al transformar Excel a TXT: {e}")
        enable_process_button()
        return

    enable_process_button()

# 11. Seleccionar archivo y mostrar ventana de personajes
def seleccionar_archivo(entry_label):
    file_path = filedialog.askopenfilename(
        title="Seleccionar archivo Excel",
        filetypes=(("Archivos Excel", "*.xlsx *.xls"), ("Todos los archivos", "*.*"))
    )
    if file_path:
        entry_label.config(text=file_path)
        crear_ventana_personajes(file_path)

# 12. Crear ventana para seleccionar personajes
def crear_ventana_personajes(file_path):
    df = leer_archivo(file_path)
    if df is None:
        return

    if 'PERSONAJE' not in df.columns:
        messagebox.showerror("Error", "El archivo no contiene la columna 'PERSONAJE'")
        return

    personajes = sorted(df['PERSONAJE'].dropna().unique())

    # Crear ventana nueva
    window = tk.Toplevel()
    window.title("Seleccionar Personajes")

    # Variable para controlar si el procesamiento está en curso
    processing = [False]  # Usamos una lista para que sea mutable

    # Función para manejar el cierre de la ventana
    def on_closing():
        if processing[0]:
            messagebox.showwarning("Advertencia", "El procesamiento está en curso, por favor espera a que termine.")
        else:
            window.destroy()

    window.protocol("WM_DELETE_WINDOW", on_closing)

    # Marco para búsqueda
    search_frame = tk.Frame(window)
    search_frame.pack(pady=5)

    search_label = tk.Label(search_frame, text="Buscar:")
    search_label.pack(side=tk.LEFT)

    search_var = tk.StringVar()
    search_entry = tk.Entry(search_frame, textvariable=search_var)
    search_entry.pack(side=tk.LEFT)

    # Marco para botones de seleccionar/deseleccionar todos
    button_frame = tk.Frame(window)
    button_frame.pack(pady=5)

    select_all_button = ttk.Button(button_frame, text="Seleccionar Todos")
    select_all_button.pack(side=tk.LEFT, padx=5)

    deselect_all_button = ttk.Button(button_frame, text="Deseleccionar Todos")
    deselect_all_button.pack(side=tk.LEFT, padx=5)

    # Marco para checkboxes
    checkbox_frame = tk.Frame(window)
    checkbox_frame.pack()

    # Añadir un canvas y un scrollbar
    canvas = tk.Canvas(checkbox_frame)
    scrollbar = ttk.Scrollbar(checkbox_frame, orient="vertical", command=canvas.yview)
    scrollable_frame = tk.Frame(canvas)

    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(
            scrollregion=canvas.bbox("all")
        )
    )

    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    # Variables y widgets de checkboxes
    checkbox_vars = {}
    checkbox_widgets = {}

    # Función para actualizar checkboxes según búsqueda
    def update_checkboxes(*args):
        search_text = search_var.get().lower()
        for personaje, var in checkbox_vars.items():
            if search_text in personaje.lower():
                checkbox_widgets[personaje].pack(anchor='w')
            else:
                checkbox_widgets[personaje].pack_forget()

    search_var.trace('w', update_checkboxes)

    # Crear checkboxes para cada personaje
    for personaje in personajes:
        var = tk.BooleanVar(value=True)
        checkbox = tk.Checkbutton(scrollable_frame, text=personaje, variable=var)
        checkbox.pack(anchor='w')
        checkbox_vars[personaje] = var
        checkbox_widgets[personaje] = checkbox

    # Función para seleccionar todos
    def select_all():
        for var in checkbox_vars.values():
            var.set(True)

    # Función para deseleccionar todos
    def deselect_all():
        for var in checkbox_vars.values():
            var.set(False)

    select_all_button.config(command=select_all)
    deselect_all_button.config(command=deselect_all)

    # Botón para iniciar procesamiento
    process_button = ttk.Button(window, text="Iniciar Procesamiento")
    process_button.pack(pady=10)

    # Etiqueta de estado
    status_label = tk.Label(window, text="Esperando...", fg="green")
    status_label.pack(pady=5)

    # Función para iniciar procesamiento
    def iniciar():
        selected_personajes = [p for p, var in checkbox_vars.items() if var.get()]
        if not selected_personajes:
            messagebox.showwarning("Advertencia", "Debe seleccionar al menos un personaje para procesar.")
            return

        # Deshabilitar el botón y marcar que el procesamiento está en curso
        process_button.config(state=tk.DISABLED)
        processing[0] = True

        # Iniciar procesamiento
        threading.Thread(target=procesar_archivo, args=(file_path, selected_personajes, status_label, window, process_button), daemon=True).start()

    process_button.config(command=iniciar)

# 13. Crear interfaz gráfica principal
def crear_interfaz():
    root = tk.Tk()
    root.title("Optimización de Takes")

    root.geometry("500x200")
    root.resizable(False, False)

    frame = tk.Frame(root, padx=20, pady=20)
    frame.pack(fill=tk.BOTH, expand=True)

    file_label = tk.Label(frame, text="No se ha seleccionado ningún archivo", wraplength=400, justify="left")
    file_label.pack(pady=(0, 10))

    select_button = ttk.Button(frame, text="Seleccionar Archivo Excel", command=lambda: seleccionar_archivo(file_label), width=25)
    select_button.pack()

    root.mainloop()

# Ejecutar la interfaz gráfica
if __name__ == "__main__":
    crear_interfaz()
