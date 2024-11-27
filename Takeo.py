import pandas as pd
from datetime import timedelta
import re
import warnings
import unicodedata
import sys
import functools
from itertools import groupby
from operator import itemgetter

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

# 2. Leer y procesar los datos
try:
    df = pd.read_excel('ANNA IS MISSING-02.xlsx')
except FileNotFoundError:
    print("Error: El archivo 'ANNA IS MISSING-02.xlsx' no se encontró.")
    exit(1)
except Exception as e:
    print(f"Error al leer el archivo Excel: {e}")
    exit(1)

# Verificar que todas las columnas necesarias existan
columnas_necesarias = ['IN', 'OUT', 'PERSONAJE', 'DIÁLOGO', 'SCENE']
for columna in columnas_necesarias:
    if columna not in df.columns:
        print(f"Error: La columna '{columna}' no se encuentra en el archivo Excel.")
        exit(1)

df['in_td'] = df['IN'].apply(time_to_timedelta)
df['out_td'] = df['OUT'].apply(time_to_timedelta)
df['duracion'] = (df['out_td'] - df['in_td']).dt.total_seconds()
df = df.sort_values(by=['in_td', 'out_td']).reset_index(drop=True)

# 2.1. Identificar Cambios de Escena basados en la columna 'SCENE'
df['cambio_escena'] = df['SCENE'] != df['SCENE'].shift(1)
df['cambio_escena'].fillna(False, inplace=True)  # La primera fila no es un cambio de escena

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

# 6. Aplicar la expansión de diálogos antes de asignar takes
df = expandir_dialogos(df)

# Limpiar las columnas de texto después de expandir diálogos
df['DIÁLOGO'] = df['DIÁLOGO'].apply(clean_text)
df['PERSONAJE'] = df['PERSONAJE'].apply(clean_text)

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
                best_takes = [take_intervenciones] + next_takes
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
def asignar_takes_optimizado(df):
    df = df.reset_index(drop=True)
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

# 6. Asignar los *takes* optimizados después de expandir diálogos
df_prop_optimizada = asignar_takes_optimizado(df)

# 8. Calcular y limpiar después de asignar *takes*
df_prop_optimizada['DURACIÓN'] = df_prop_optimizada['DURACIÓN'].astype(float)

# 9. Calcular el total de *takes* por personaje
takes_por_personaje_optimizada, suma_total_takes_optimizada = calcular_total_takes_por_personaje(df_prop_optimizada)

# 10. Exportar la propuesta optimizada y el resumen a Excel
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
    print("La propuesta optimizada y su resumen han sido exportados a 'prop_optimizada.xlsx'")
except Exception as e:
    print(f"Error al exportar a Excel: {e}")
    exit(1)

# 11. Mostrar el resumen de *takes* por personaje en la consola
print("Propuesta Optimizada:")
print(takes_por_personaje_optimizada)
print(f"Suma total de *Takes*: {suma_total_takes_optimizada}\n")
