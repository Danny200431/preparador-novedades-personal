# Importamos las herramientas del proyecto
from datetime import datetime
from modulos.lectura import leer_novedades
from modulos.configuracion import cargar_parametros
from modulos.procesamiento import procesar_novedad, extraer_fecha_vinculacion
from modulos.reporte import crear_reporte
from modulos.registro import crear_registro, guardar_mensaje
from modulos.excepciones import ErrorProyecto

# Preparamos los datos de la ejecución
resultados = []
parametros = {}
inicio = datetime.now()
archivo_entrada = ""

print("PREPARADOR DE NOVEDADES DE PERSONAL")
print("----------------------------------")

try:
    # Leemos las novedades del archivo JSON
    novedades = leer_novedades()
    archivo_entrada = getattr(
        leer_novedades, "archivo_procesado", ""
    )

    guardar_mensaje(
        "Archivo de entrada leído correctamente",
        "INFO",
        "lectura",
        "leer_novedades",
        archivo_entrada
    )

    print("Novedades encontradas:", len(novedades))

    # Cargamos la configuración del Excel
    parametros, cargos, centros, matriz, campos_obligatorios = cargar_parametros()

    guardar_mensaje(
        "Configuración cargada correctamente",
        "INFO",
        "configuracion",
        "cargar_parametros"
    )

    # Revisamos cada novedad
    for posicion, novedad in enumerate(novedades, start=1):
        correcta, mensaje, datos = procesar_novedad(
            novedad,
            parametros,
            cargos,
            centros,
            matriz,
            campos_obligatorios
        )

        # Conservamos el consecutivo original si existe
        consecutivo = novedad.get("consecutivo", posicion)

        # Preparamos los datos para el reporte
        resultado = novedad.copy()
        resultado.update(datos)

        resultado["consecutivo"] = consecutivo
        resultado["datos"] = datos
        resultado["tipo"] = novedad.get("tipo", "Sin tipo")
        resultado["numero_documento"] = novedad.get(
            "numero_documento", "Sin documento"
        )
        resultado["fecha_vinculacion_previa"] = extraer_fecha_vinculacion(
            novedad.get("observaciones", "")
        )
        resultado["resultado"] = "ACEPTADA" if correcta else "RECHAZADA"
        resultado["mensaje"] = mensaje

        # Guardamos el motivo y la clasificación del rechazo
        if not correcta:
            resultado["motivo"] = mensaje
            resultado["clasificacion"] = datos.get(
                "clasificacion", "VALIDACION"
            )
            resultado["campo"] = datos.get("campo", "")

        resultados.append(resultado)

        # Guardamos el resultado de cada novedad en el registro
        guardar_mensaje(
            mensaje,
            "INFO" if correcta else "WARNING",
            "procesamiento",
            "procesar_novedad",
            novedad.get("numero_documento", consecutivo)
        )

        print(
            resultado["tipo"], "-",
            resultado["numero_documento"], "-",
            resultado["resultado"], "-", mensaje
        )

    # Creamos el reporte final
    ruta_reporte = crear_reporte(
        resultados,
        parametros,
        archivo_entrada
    )
    print("Reporte guardado en:", ruta_reporte)

except ErrorProyecto as error:
    guardar_mensaje(
        str(error),
        "ERROR",
        "main",
        "ejecucion",
        detalle_tecnico=type(error).__name__
    )
    print("Error del proyecto:", error)

except Exception as error:
    guardar_mensaje(
        str(error),
        "CRITICAL",
        "main",
        "ejecucion",
        detalle_tecnico=type(error).__name__
    )
    print("Ocurrió un error inesperado:", error)

finally:
    # Creamos el registro incluso si ocurre un error
    try:
        ruta_registro = crear_registro(
            resultados,
            parametros,
            inicio,
            archivo_entrada
        )
        print("Registro guardado en:", ruta_registro)
    except Exception as error:
        print("No se pudo guardar el registro:", error)

    # Mostramos el tiempo de ejecución
    fin = datetime.now()
    print("Inicio:", inicio.strftime("%Y-%m-%d %H:%M:%S"))
    print("Fin:", fin.strftime("%Y-%m-%d %H:%M:%S"))
    print(
        "Duración en segundos:",
        round((fin - inicio).total_seconds(), 2)
    )
    print("Total revisado:", len(resultados))