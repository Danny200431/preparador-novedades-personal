# Este archivo crea el reporte de Excel con los resultados
# de las novedades de personal.

import os
from datetime import datetime
import pandas as pd


def clasificar_rechazo(motivo):
    """Clasifica el rechazo como validación o regla de negocio."""

    texto = str(motivo or "").strip().lower()

    motivos_regla_negocio = (
        "el área del cargo no coincide con el centro de costo",
        "el area del cargo no coincide con el centro de costo",
        "el cargo no existe en el catálogo",
        "el cargo no existe en el catalogo",
        "el centro de costo no existe en el catálogo",
        "el centro de costo no existe en el catalogo",
        "el cargo y el área no están permitidos",
        "el cargo y el area no estan permitidos",
    )

    if texto in motivos_regla_negocio:
        return "REGLA_NEGOCIO"

    return "VALIDACION"


def obtener_valor_recibido(resultado, datos, motivo, campo):
    """Busca el valor original relacionado con el rechazo."""

    novedad_original = resultado.get("novedad_original", {})

    if not isinstance(novedad_original, dict):
        novedad_original = {}

    fuentes = [novedad_original, resultado, datos]

    # Si ya existe un valor recibido, lo utilizamos.
    for fuente in fuentes:
        valor = fuente.get("valor_recibido")

        if valor is not None and str(valor).strip() != "":
            return valor

    texto = str(motivo or "").strip().lower()
    campo = str(campo or "").strip().lower()

    # Identificamos el campo que produjo el rechazo.
    if "falta el campo obligatorio:" in texto:
        nombre_campo = texto.split(
            "falta el campo obligatorio:", 1
        )[1].strip()

        for fuente in fuentes:
            valor = fuente.get(nombre_campo)

            if valor is not None and str(valor).strip() != "":
                return valor

        return "Campo vacío"

    if "la cc debe tener" in texto:
        campos = ["numero_documento"]
    elif (
        "fecha efectiva ya pasó" in texto
        or "fecha efectiva ya paso" in texto
        or "fecha supera el límite permitido" in texto
        or "fecha supera el limite permitido" in texto
    ):
        campos = ["fecha_efectiva"]
    elif (
        "el área del cargo no coincide" in texto
        or "el area del cargo no coincide" in texto
    ):
        campos = ["centro_costo", "cargo"]
    elif "el centro de costo no existe" in texto:
        campos = ["centro_costo"]
    elif "el cargo no existe" in texto:
        campos = ["cargo", "codigo_cargo"]
    elif "teléfono" in texto or "telefono" in texto:
        campos = ["telefono"]
    else:
        equivalencias = {
            "tipo": ["tipo", "tipo_novedad"],
            "tipo_novedad": ["tipo_novedad", "tipo"],
            "numero_documento": ["numero_documento"],
            "documento": ["numero_documento"],
            "fecha_efectiva": ["fecha_efectiva"],
            "centro_costo": ["centro_costo"],
            "cargo": ["cargo", "codigo_cargo"],
            "cargo/centro_costo": [
                "cargo", "codigo_cargo", "centro_costo"
            ],
            "telefono": ["telefono"],
            "correo_corporativo": ["correo_corporativo"],
            "nombre": [
                "nombre_completo", "nombre", "nombres"
            ],
        }

        campos = equivalencias.get(campo, [])

    # Buscamos el dato en las fuentes disponibles.
    for nombre_campo in campos:
        for fuente in fuentes:
            valor = fuente.get(nombre_campo)

            if valor is not None and str(valor).strip() != "":
                return valor

    return ""


def crear_reporte(resultados, parametros, archivo_entrada=""):
    """Crea el archivo Excel con las cuatro hojas solicitadas."""

    os.makedirs("salida", exist_ok=True)

    aceptadas = []
    rechazadas = []
    correos = {}

    # Separamos las novedades aceptadas y rechazadas.
    for resultado in resultados:
        datos = resultado.get("datos", {})

        if not isinstance(datos, dict):
            datos = {}

        if resultado.get("resultado") == "ACEPTADA":
            aceptadas.append(resultado)
        else:
            motivo = resultado.get(
                "motivo",
                resultado.get("mensaje", "")
            )

            clasificacion = str(
                resultado.get(
                    "clasificacion",
                    datos.get("clasificacion", "")
                )
            ).strip().upper()

            if clasificacion not in (
                "VALIDACION",
                "REGLA_NEGOCIO"
            ):
                clasificacion = clasificar_rechazo(motivo)

            campo = resultado.get(
                "campo",
                datos.get("campo", "")
            )

            valor_recibido = obtener_valor_recibido(
                resultado,
                datos,
                motivo,
                campo
            )

            rechazadas.append({
                "consecutivo": resultado.get("consecutivo", ""),
                "tipo": resultado.get(
                    "tipo",
                    resultado.get("tipo_novedad", "")
                ),
                "numero_documento": resultado.get(
                    "numero_documento", ""
                ),
                "nombre": resultado.get(
                    "nombre_completo",
                    resultado.get("nombre", "")
                ),
                "clasificacion": clasificacion,
                "campo": campo,
                "motivo": motivo,
                "valor_recibido": valor_recibido,
            })

    # Columnas de la hoja de novedades aceptadas.
    columnas_carga = [
        "consecutivo",
        "tipo_novedad",
        "tipo_documento",
        "numero_documento",
        "nombre_completo",
        "correo_corporativo",
        "codigo_cargo",
        "nombre_cargo",
        "nivel",
        "centro_costo",
        "area",
        "perfil_sistema",
        "requiere_aprobacion",
        "fecha_efectiva",
        "telefono",
        "fecha_vinculacion_previa",
    ]

    filas_carga = []

    for resultado in aceptadas:
        datos = resultado.get("datos", {})

        if not isinstance(datos, dict):
            datos = {}

        fila = {}

        # Tomamos los valores preparados para el sistema.
        # El consecutivo se conserva desde el resultado original.
        for columna in columnas_carga:
            if columna == "consecutivo":
                fila[columna] = resultado.get("consecutivo", "")
            else:
                fila[columna] = datos.get(
                    columna,
                    resultado.get(columna, "")
                )

        filas_carga.append(fila)

    df_carga = pd.DataFrame(
        filas_carga,
        columns=columnas_carga
    )

    # Columnas de la hoja de rechazados.
    columnas_rechazados = [
        "consecutivo",
        "tipo",
        "numero_documento",
        "nombre",
        "clasificacion",
        "campo",
        "motivo",
        "valor_recibido",
    ]

    df_rechazados = pd.DataFrame(
        rechazadas,
        columns=columnas_rechazados
    )

    # Buscamos correos repetidos entre las novedades aceptadas.
    for resultado in aceptadas:
        datos = resultado.get("datos", {})

        if not isinstance(datos, dict):
            datos = {}

        correo = str(
            datos.get(
                "correo_corporativo",
                resultado.get("correo_corporativo", "")
            )
        ).strip().lower()

        if correo and "@" in correo:
            correo_base = correo.split("@", 1)[0]

            correos.setdefault(correo_base, []).append({
                "consecutivo": resultado.get("consecutivo", ""),
                "correo": correo,
            })

    filas_colisiones = []

    for correo_base, registros in correos.items():
        if len(registros) > 1:
            filas_colisiones.append({
                "correo_base": correo_base,
                "consecutivos": ", ".join(
                    str(registro["consecutivo"])
                    for registro in registros
                ),
                "correos_finales": ", ".join(
                    registro["correo"]
                    for registro in registros
                ),
            })

    columnas_colisiones = [
        "correo_base",
        "consecutivos",
        "correos_finales",
    ]

    df_colisiones = pd.DataFrame(
        filas_colisiones,
        columns=columnas_colisiones
    )

    # Contamos las novedades por tipo.
    tipos = {}

    for resultado in resultados:
        tipo = str(
            resultado.get("tipo", "Sin tipo")
        ).strip().upper()

        if tipo not in tipos:
            tipos[tipo] = {
                "tipo_novedad": tipo,
                "recibidas": 0,
                "validas": 0,
                "rechazadas": 0,
            }

        tipos[tipo]["recibidas"] += 1

        if resultado.get("resultado") == "ACEPTADA":
            tipos[tipo]["validas"] += 1
        else:
            tipos[tipo]["rechazadas"] += 1

    filas_resumen = []

    for tipo, datos_tipo in tipos.items():
        recibidas = datos_tipo["recibidas"]
        validas = datos_tipo["validas"]

        porcentaje = (
            round(validas / recibidas * 100, 2)
            if recibidas else 0
        )

        filas_resumen.append({
            "tipo_novedad": tipo,
            "recibidas": recibidas,
            "validas": validas,
            "rechazadas": datos_tipo["rechazadas"],
            "porcentaje_exito": porcentaje,
        })

    # Añadimos la fila total.
    filas_resumen.append({
        "tipo_novedad": "TOTAL",
        "recibidas": len(resultados),
        "validas": len(aceptadas),
        "rechazadas": len(rechazadas),
        "porcentaje_exito": (
            round(len(aceptadas) / len(resultados) * 100, 2)
            if resultados else 0
        ),
    })

    # Buscamos el motivo de rechazo más frecuente.
    motivos = [
        registro["motivo"]
        for registro in rechazadas
        if registro["motivo"]
    ]

    motivo_frecuente = (
        pd.Series(motivos).value_counts().idxmax()
        if motivos else ""
    )

    filas_resumen.append({
        "tipo_novedad": "Motivo de rechazo más frecuente",
        "recibidas": "",
        "validas": "",
        "rechazadas": "",
        "porcentaje_exito": motivo_frecuente,
    })

    filas_resumen.append({
        "tipo_novedad": "Archivo de entrada",
        "recibidas": "",
        "validas": "",
        "rechazadas": "",
        "porcentaje_exito": (
            os.path.basename(str(archivo_entrada))
            if archivo_entrada else ""
        ),
    })

    df_resumen = pd.DataFrame(filas_resumen)

    # Creamos el nombre del archivo Excel.
    fecha_actual = datetime.now().strftime("%Y%m%d")

    ruta_salida = os.path.join(
        "salida",
        f"carga_novedades_{fecha_actual}.xlsx"
    )

    # Guardamos las cuatro hojas en el mismo archivo.
    with pd.ExcelWriter(
        ruta_salida,
        engine="openpyxl"
    ) as escritor:
        df_carga.to_excel(
            escritor,
            sheet_name="Carga_Sistema",
            index=False
        )

        df_rechazados.to_excel(
            escritor,
            sheet_name="Rechazados",
            index=False
        )

        df_colisiones.to_excel(
            escritor,
            sheet_name="Colisiones_Correo",
            index=False
        )

        df_resumen.to_excel(
            escritor,
            sheet_name="Resumen",
            index=False
        )

        # Damos formato básico a las hojas del reporte.
        for hoja in escritor.book.worksheets:
            hoja.freeze_panes = "A2"

            for celda in hoja[1]:
                celda.font = celda.font.copy(
                    bold=True,
                    color="FFFFFF"
                )
                celda.fill = __import__(
                    "openpyxl"
                ).styles.PatternFill(
                    fill_type="solid",
                    fgColor="1F4E78"
                )

            # Ajustamos el ancho de las columnas al contenido.
            for columna in hoja.columns:
                letra = columna[0].column_letter
                ancho = max(
                    len(str(celda.value or ""))
                    for celda in columna
                )
                hoja.column_dimensions[letra].width = min(
                    ancho + 2, 45
                )

    return ruta_salida