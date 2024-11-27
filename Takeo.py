import pandas as pd
from datetime import timedelta
import re
import warnings
import unicodedata
import sys
import functools
from itertools import groupby
from operator import itemgetter
import tkinter as tk
from tkinter import filedialog, messagebox
import threading

# Aumentar el límite de recursión si es necesario
sys.setrecursionlimit(100000)

# Ignorar advertencias de futuras versiones
warnings.simplefilter(action='ignore', category=FutureWarning)

# 1. Función para convertir tiempo a timedelta
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
        print(f"Error al convertir tiempo: {time_str} - {e}")
        return timedelta(0)

# 5. Dividir diálogos que excedan los 60 caracteres (excluyendo contenido entre paréntesis)
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

def expandir_dialogos(df_original):
    # Lista para almacenar todas las nuevas filas
    nuevas_filas = []

    for idx, row in df_original.iterrows():
        # Dividir el diálogo si excede el límite de caracteres
        lineas_divididas = dividir_dialogo(row['DIÁLOGO'])
        for linea in lineas_divididas:
            new_row = row.copy()
            new_row['DIÁLOGO'] = linea
            nuevas_filas.append(new_row)

    # Crear un nuevo DataFrame con las filas expandidas
    df_expanded = pd.DataFrame(nuevas_filas)
    return df_expanded

# Limpiar columnas de texto
def clean_text(text):
    if isinstance(text, str):
        # Remover caracteres de control
        text = ''.join(ch for ch in text if unicodedata.category(ch)[0] != 'C')
        return text
    else:
        return text

# 3. Función para optimizar la división de *takes* en una escena
def optimizar_takes_escena(intervenciones_escena, max_duracion_take=30, max_lineas_take=10, max_lineas_consecutivas=5):
    # Crear una lista de intervenciones con la información necesaria
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

    # Agrupar intervenciones por tiempo de IN y OUT
    intervenciones_sorted = sorted(intervenciones, key=lambda x: (x['in_td'], x['out_td']))
    bloques = [list(group) for key, group in groupby(intervenciones_sorted, key=lambda x: (x['in_td'], x['out_td']))]

    n = len(bloques)

    # Función recursiva con memoización
    @functools.lru_cache(maxsize=None)
    def dp(pos):
        if pos >= n:
            return [], (0, 0)  # No hay más bloques

        best_takes = None
        best_cost = None

        for end in range(pos + 1, n + 1):
            # Construir un posible *take* desde 'pos' hasta 'end'
            take_bloques = bloques[pos:end]
            take_intervenciones = [intervencion for bloque in take_bloques for intervencion in bloque]

            # Verificar restricciones
            duracion_take = take_intervenciones[-1]['out_td'] - take_intervenciones[0]['in_td']
            if duracion_take.total_seconds() > max_duracion_take:
                break  # Excede la duración máxima

            if len(take_intervenciones) > max_lineas_take:
                break  # Excede el número máximo de líneas

            # Verificar líneas consecutivas por personaje
            valid = True
            personaje_lineas_consecutivas = {}
            for i, intervencion in enumerate(take_intervenciones):
                personaje = intervencion['personaje']
                if i == 0 or personaje != take_intervenciones[i - 1]['personaje']:
                    personaje_lineas_consecutivas[personaje] = 1
                else:
                    personaje_lineas_consecutivas[personaje] += 1
                if personaje_lineas_consecutivas[personaje] > max_lineas_consecutivas:
                    valid = False
                    break  # Excede el máximo de líneas consecutivas por personaje
            if not valid:
                continue

            # Obtener el resultado de los bloques restantes
            next_takes, next_cost = dp(end)

            # Calcular el costo actual
            personajes_en_take = set([intervencion['personaje'] for intervencion in take_intervenciones])
            total_takes_por_personaje = next_cost[0] + len(personajes_en_take)
            total_takes = next_cost[1] + 1  # Se añade un nuevo take

            current_cost = (total_takes_por_personaje, total_takes)

            if best_cost is None or current_cost < best_cost:
                best_takes = [tuple(take_intervenciones)] + list(next_takes)
                best_cost = current_cost

        if best_takes is None:
            # No es posible dividir desde esta posición
            return [], (float('inf'), float('inf'))

        return best_takes, best_cost

    # Obtener la mejor división de *takes* y el costo mínimo
    best_takes, _ = dp(0)

    # Construir la lista de *takes* con sus intervenciones
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

# 4. Función para asignar *takes* optimizados
def asignar_takes_optimizado(df, omit_rotulo=False):
    df = df.reset_index(drop=True)
    if omit_rotulo:
        df = df[df['PERSONAJE'].str.upper() != 'ROTULO'].reset_index(drop=True)
    
    escenas = df['SCENE'].unique()
    all_takes = []
    take_global_id = 1

    for escena in escenas:
        intervenciones_escena = df[df['SCENE'] == escena]
        takes_escena = optimizar_takes_escena(intervenciones_escena)
        # Actualizar el ID global de *takes*
        for take in takes_escena:
            take['take'] = take_global_id
            take_global_id += 1
            all_takes.append(take)

    # Construir el DataFrame de resultados
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

    df_takes = pd.DataFrame(take_list)
    return df_takes

# 7. Calcular el total de *takes* por personaje para la propuesta optimizada
def calcular_total_takes_por_personaje(df_takes):
    takes_por_personaje = df_takes.groupby('PERSONAJE')['TAKE'].nunique().reset_index()
    takes_por_personaje.rename(columns={'TAKE': 'TOTAL_TAKES'}, inplace=True)
    suma_total_takes = takes_por_personaje['TOTAL_TAKES'].sum()
    return takes_por_personaje, suma_total_takes

# Función principal para procesar el archivo
def procesar_archivo(file_path, omit_rotulo, status_label):
    try:
        status_label.config(text="Leyendo el archivo Excel...")
        df = pd.read_excel(file_path)
    except FileNotFoundError:
        messagebox.showerror("Error", f"El archivo '{file_path}' no se encontró.")
        return
    except Exception as e:
        messagebox.showerror("Error", f"Error al leer el archivo Excel: {e}")
        return

    # Verificar que todas las columnas necesarias existan
    columnas_necesarias = ['IN', 'OUT', 'PERSONAJE', 'DIÁLOGO', 'SCENE']
    for columna in columnas_necesarias:
        if columna not in df.columns:
            messagebox.showerror("Error", f"La columna '{columna}' no se encuentra en el archivo Excel.")
            return

    # Convertir tiempos a timedelta
    status_label.config(text="Convirtiendo tiempos...")
    df['in_td'] = df['IN'].apply(time_to_timedelta)
    df['out_td'] = df['OUT'].apply(time_to_timedelta)
    df['duracion'] = (df['out_td'] - df['in_td']).dt.total_seconds()
    df = df.sort_values(by=['in_td', 'out_td']).reset_index(drop=True)

    # Identificar Cambios de Escena basados en la columna 'SCENE'
    df['cambio_escena'] = df['SCENE'] != df['SCENE'].shift(1)
    df['cambio_escena'].fillna(False, inplace=True)  # La primera fila no es un cambio de escena

    # Dividir diálogos que excedan los 60 caracteres
    status_label.config(text="Dividiendo diálogos largos...")
    df = expandir_dialogos(df)

    # Limpiar las columnas de texto después de expandir diálogos
    status_label.config(text="Limpiando texto...")
    df['DIÁLOGO'] = df['DIÁLOGO'].apply(clean_text)
    df['PERSONAJE'] = df['PERSONAJE'].apply(clean_text)

    # Asignar los *takes* optimizados
    status_label.config(text="Asignando *takes* optimizados...")
    df_prop_optimizada = asignar_takes_optimizado(df, omit_rotulo=omit_rotulo)

    # Calcular y limpiar después de asignar *takes*
    df_prop_optimizada['DURACIÓN'] = df_prop_optimizada['DURACIÓN'].astype(float)

    # Calcular el total de *takes* por personaje
    status_label.config(text="Calculando resumen de *takes* por personaje...")
    takes_por_personaje_optimizada, suma_total_takes_optimizada = calcular_total_takes_por_personaje(df_prop_optimizada)

    # Exportar la propuesta optimizada y el resumen a Excel
    status_label.config(text="Exportando a Excel...")
    try:
        with pd.ExcelWriter('prop_optimizada.xlsx', engine='xlsxwriter') as writer:
            # Exportar los *takes* de la propuesta optimizada
            df_prop_optimizada.to_excel(writer, sheet_name='Optimizada_Takes', index=False)

            # Exportar el resumen de *takes* por personaje
            takes_por_personaje_optimizada.to_excel(writer, sheet_name='Resumen', index=False)

            # Añadir la suma total de takes
            workbook  = writer.book
            worksheet = writer.sheets['Resumen']
            # Escribir la suma total de takes debajo de la tabla
            last_row = len(takes_por_personaje_optimizada) + 1
            worksheet.write(f'A{last_row + 1}', 'Suma total de Takes:')
            worksheet.write(f'B{last_row + 1}', suma_total_takes_optimizada)
        status_label.config(text="Exportación completada: 'prop_optimizada.xlsx'")
        messagebox.showinfo("Éxito", "La propuesta optimizada y su resumen han sido exportados a 'prop_optimizada.xlsx'")
    except Exception as e:
        messagebox.showerror("Error", f"Error al exportar a Excel: {e}")
        return

    # Mostrar el resumen de *takes* por personaje en la consola
    print("Propuesta Optimizada:")
    print(takes_por_personaje_optimizada)
    print(f"Suma total de *Takes*: {suma_total_takes_optimizada}\n")

# Función para manejar el botón de selección de archivo
def seleccionar_archivo(entry_label):
    file_path = filedialog.askopenfilename(
        title="Seleccionar archivo Excel",
        filetypes=(("Archivos Excel", "*.xlsx *.xls"), ("Todos los archivos", "*.*"))
    )
    if file_path:
        entry_label.config(text=file_path)

# Función para manejar el botón de procesamiento
def iniciar_procesamiento(file_label, omit_rotulo_var, status_label):
    file_path = file_label.cget("text")
    omit_rotulo = omit_rotulo_var.get()

    if not file_path or file_path == "No se ha seleccionado ningún archivo":
        messagebox.showwarning("Advertencia", "Por favor, selecciona un archivo Excel antes de procesar.")
        return

    # Ejecutar el procesamiento en un hilo separado para mantener la GUI responsive
    threading.Thread(target=procesar_archivo, args=(file_path, omit_rotulo, status_label), daemon=True).start()

# Configuración de la interfaz gráfica con Tkinter
def crear_interfaz():
    root = tk.Tk()
    root.title("Optimización de *Takes*")

    # Tamaño de la ventana
    root.geometry("500x250")
    root.resizable(False, False)

    # Estilo de la fuente
    fuente = ("Helvetica", 12)

    # Frame principal
    frame = tk.Frame(root, padx=20, pady=20)
    frame.pack(fill=tk.BOTH, expand=True)

    # Selección de archivo
    file_label = tk.Label(frame, text="No se ha seleccionado ningún archivo", wraplength=400, justify="left")
    file_label.pack(pady=(0, 10))

    select_button = tk.Button(frame, text="Seleccionar Archivo Excel", command=lambda: seleccionar_archivo(file_label), width=25, font=fuente)
    select_button.pack()

    # Checkbox para omitir ROTULO
    omit_rotulo_var = tk.BooleanVar()
    omit_rotulo_checkbox = tk.Checkbutton(
        frame,
        text="Omitir ROTULO",
        variable=omit_rotulo_var,
        font=fuente
    )
    omit_rotulo_checkbox.pack(pady=(10, 10))

    # Botón para iniciar procesamiento
    process_button = tk.Button(
        frame,
        text="Iniciar Procesamiento",
        command=lambda: iniciar_procesamiento(file_label, omit_rotulo_var, status_label),
        width=20,
        font=fuente,
        bg="blue",
        fg="white"
    )
    process_button.pack(pady=(0, 10))

    # Label de estado
    status_label = tk.Label(frame, text="Esperando...", font=fuente, fg="green", wraplength=400, justify="left")
    status_label.pack(pady=(10, 0))

    root.mainloop()

# Ejecutar la interfaz gráfica
if __name__ == "__main__":
    crear_interfaz()
