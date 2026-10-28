# Este archivo lee la información del Excel de configuración.

import os
from openpyxl import load_workbook


def cargar_parametros():

    # Buscamos el archivo Excel dentro de la carpeta config
    ruta = os.path.join("config", "Parametros.xlsx")

    # Abrimos el archivo para leer sus datos
    excel = load_workbook(ruta, data_only=True)

    # Leemos los parámetros generales del programa
    hoja = excel["Parametros"]
    parametros = {}

    for fila in hoja.iter_rows(min_row=2, values_only=True):
        if fila[0] is not None:
            clave = str(fila[0]).strip()
            parametros[clave] = fila[1]

    # Leemos los cargos y guardamos sus datos por código
    hoja = excel["Cargos"]
    cargos = {}

    for fila in hoja.iter_rows(min_row=2, values_only=True):
        if fila[0] is not None:
            codigo = str(fila[0]).strip().upper()
            cargos[codigo] = fila[1:]

    # Leemos los centros de costo y sus datos
    hoja = excel["Centros_Costo"]
    centros = {}

    for fila in hoja.iter_rows(min_row=2, values_only=True):
        if fila[0] is not None:
            codigo = str(fila[0]).strip().upper()
            centros[codigo] = fila[1:]

    # Leemos los permisos y perfiles de cada cargo
    hoja = excel["Matriz_Roles"]
    matriz = []

    for fila in hoja.iter_rows(min_row=2, values_only=True):
        if fila[0] is not None:
            # Quitamos los espacios sobrantes de los textos
            fila_limpia = tuple(
                valor.strip() if isinstance(valor, str) else valor
                for valor in fila
            )
            matriz.append(fila_limpia)

    # Leemos qué campos son obligatorios para cada novedad
    hoja = excel["Campos_Obligatorios"]
    campos_obligatorios = {}

    for fila in hoja.iter_rows(min_row=2, values_only=True):
        if fila[0] is not None and fila[1] is not None:
            tipo = str(fila[0]).strip().upper()
            campos = str(fila[1]).split(";")

            # Guardamos los campos sin espacios sobrantes
            campos_obligatorios[tipo] = [
                campo.strip() for campo in campos
            ]

    # Cerramos el Excel después de leer toda la información
    excel.close()

    # Entregamos todos los datos al programa principal
    return (
        parametros,
        cargos,
        centros,
        matriz,
        campos_obligatorios
    )