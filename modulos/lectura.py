
# Este archivo busca y lee las novedades del archivo JSON.

import json
import re
from pathlib import Path
from modulos.excepciones import ErrorArchivoEntrada


# Buscamos la carpeta principal del proyecto.
CARPETA_PROYECTO = Path(__file__).resolve().parent.parent


def leer_novedades():
    # Buscamos la carpeta donde están los archivos JSON.
    carpeta = CARPETA_PROYECTO / "entrada"

    if not carpeta.is_dir():
        raise ErrorArchivoEntrada(
            "No existe la carpeta de entrada."
        )

    # Buscamos los archivos que terminan en .json.
    archivos = [
        archivo
        for archivo in carpeta.iterdir()
        if archivo.is_file() and archivo.suffix.lower() == ".json"
    ]

    if not archivos:
        raise ErrorArchivoEntrada(
            "No se encontró ningún archivo JSON en la carpeta entrada."
        )

    # Buscamos la fecha que aparece en el nombre del archivo.
    def obtener_fecha(archivo):
        coincidencia = re.search(
            r"(\d{4})-(\d{2})-(\d{2})",
            archivo.name
        )

        if coincidencia:
            return tuple(map(int, coincidencia.groups()))

        # Si no tiene fecha, queda por debajo de los archivos con fecha.
        return (0, 0, 0)

    archivos_con_fecha = [
        archivo for archivo in archivos
        if obtener_fecha(archivo) != (0, 0, 0)
    ]

    if not archivos_con_fecha:
        raise ErrorArchivoEntrada(
            "Los archivos JSON deben tener una fecha en el nombre, "
            "por ejemplo: novedades_2026-10-23.json."
        )

    # Elegimos el archivo que tiene la fecha más reciente en el nombre.
    nombre_reciente = max(
        archivos_con_fecha,
        key=lambda archivo: obtener_fecha(archivo)
    )

    try:
        # Abrimos el archivo y leemos su contenido.
        with nombre_reciente.open("r", encoding="utf-8") as archivo:
            datos = json.load(archivo)

    except (json.JSONDecodeError, OSError) as error:
        raise ErrorArchivoEntrada(
            f"No se pudo leer el archivo {nombre_reciente.name}: {error}"
        ) from error

    # Comprobamos que el contenido sea un objeto.
    if not isinstance(datos, dict):
        raise ErrorArchivoEntrada(
            "El contenido del JSON debe ser un objeto."
        )

    # Obtenemos la lista de novedades.
    novedades = datos.get("novedades")

    if not isinstance(novedades, list):
        raise ErrorArchivoEntrada(
            "El archivo JSON debe contener una lista llamada novedades."
        )

    # Guardamos el nombre del archivo que se leyó.
    print("Archivo procesado:", nombre_reciente.name)
    leer_novedades.archivo_procesado = nombre_reciente.name

    return novedades
