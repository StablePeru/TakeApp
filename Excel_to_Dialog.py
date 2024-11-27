import pandas as pd
import os

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

# Ruta al archivo Excel y al archivo de salida
ruta_excel = "prop_optimizada.xlsx"
ruta_salida_txt = "salida_formateada.txt"

# Llamar a la función
transformar_excel_a_txt(ruta_excel, ruta_salida_txt)
