# Importamos las herramientas para leer archivos JSON
import json
import os

# Importamos los errores propios del proyecto
from modulos.excepciones import ErrorArchivoEntrada


def leer_novedades():
    # Indicamos la carpeta donde están los archivos JSON
    carpeta = "entrada"

    # Buscamos los archivos que terminan en .json
    if not os.path.isdir(carpeta):
        raise ErrorArchivoEntrada(
            "No existe la carpeta de entrada."
        )

    archivos = [
        nombre
        for nombre in os.listdir(carpeta)
        if nombre.lower().endswith(".json")
        and os.path.isfile(os.path.join(carpeta, nombre))
    ]

    # Verificamos que exista al menos un archivo JSON
    if not archivos:
        raise ErrorArchivoEntrada(
            "No se encontró ningún archivo JSON en la carpeta entrada."
        )

    # Elegimos el archivo JSON modificado más recientemente
    nombre_reciente = max(
        archivos,
        key=lambda nombre: os.path.getmtime(
            os.path.join(carpeta, nombre)
        )
    )

    # Guardamos la ruta del archivo seleccionado
    ruta_archivo = os.path.join(carpeta, nombre_reciente)

    try:
        # Abrimos el archivo y leemos su contenido
        with open(ruta_archivo, "r", encoding="utf-8") as archivo:
            datos = json.load(archivo)

    except (json.JSONDecodeError, OSError) as error:
        # Avisamos si el archivo está dañado o no se puede leer
        raise ErrorArchivoEntrada(
            f"No se pudo leer el archivo {nombre_reciente}: {error}"
        ) from error

    # Comprobamos que el contenido sea un objeto
    if not isinstance(datos, dict):
        raise ErrorArchivoEntrada(
            "El contenido del JSON debe ser un objeto."
        )

    # Obtenemos la lista de novedades
    novedades = datos.get("novedades")

    # Verificamos que novedades sea una lista
    if not isinstance(novedades, list):
        raise ErrorArchivoEntrada(
            "El archivo JSON debe contener una lista llamada novedades."
        )

    # Mostramos el nombre del archivo procesado
    print("Archivo procesado:", nombre_reciente)
    leer_novedades.archivo_procesado = nombre_reciente

    # Devolvemos la lista de novedades
    return novedades