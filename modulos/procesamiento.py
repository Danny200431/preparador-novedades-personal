# Este archivo revisa, valida y prepara las novedades de personal.

import re
from datetime import datetime

from modulos.validacion import (
    PATRONES,
    validar_tipo,
    validar_campos,
    validar_documento,
    validar_fecha,
    validar_centro_costo,
    validar_telefono,
    validar_regla_negocio,
)


# Buscamos la fecha más antigua en las observaciones.
def extraer_fecha_vinculacion(observaciones):
    texto = str(observaciones or "")

    # Buscamos las fechas usando el patrón compartido.
    fechas = PATRONES["fecha_observaciones"].findall(texto)
    fechas_validas = []

    for fecha_texto in fechas:
        fecha_convertida = None

        for formato in ("%d/%m/%Y", "%Y-%m-%d"):
            try:
                fecha_convertida = datetime.strptime(
                    fecha_texto, formato
                ).date()
                break
            except ValueError:
                continue

        if fecha_convertida is not None:
            fechas_validas.append((fecha_convertida, fecha_texto))

    if not fechas_validas:
        return ""

    fecha_mas_antigua = min(
        fechas_validas,
        key=lambda elemento: elemento[0]
    )

    return fecha_mas_antigua[1]


# Guardamos el motivo, la clasificación y el campo del rechazo.
def rechazar_novedad(mensaje, clasificacion="VALIDACION", campo=""):
    return False, mensaje, {
        "clasificacion": clasificacion,
        "campo": campo
    }


# Revisamos y preparamos una novedad.
def procesar_novedad(
    novedad,
    parametros,
    cargos,
    centros,
    matriz,
    campos_obligatorios
):
    # Validamos el tipo de novedad.
    correcta, mensaje = validar_tipo(
        novedad, campos_obligatorios
    )

    if not correcta:
        return rechazar_novedad(
            mensaje, "VALIDACION", "tipo"
        )

    # Revisamos los campos obligatorios del Excel.
    correcta, mensaje = validar_campos(
        novedad, campos_obligatorios
    )

    if not correcta:
        campo = mensaje.split(": ", 1)[1] if ": " in mensaje else ""
        return rechazar_novedad(
            mensaje, "VALIDACION", campo
        )

    # Validamos el documento.
    correcta, mensaje = validar_documento(novedad)

    if not correcta:
        return rechazar_novedad(
            mensaje, "VALIDACION", "numero_documento"
        )

    # Validamos la fecha efectiva.
    correcta, mensaje = validar_fecha(
        novedad,
        parametros.get("dias_maximo_futuro", 90)
    )

    if not correcta:
        return rechazar_novedad(
            mensaje, "VALIDACION", "fecha_efectiva"
        )

    # Validamos el centro de costo cuando fue informado.
    if str(novedad.get("centro_costo", "")).strip():
        correcta, mensaje = validar_centro_costo(
            novedad, centros
        )

        if not correcta:
            if (
                "no existe en el catálogo" in mensaje.lower()
                or "no existe en el catalogo" in mensaje.lower()
            ):
                clasificacion = "REGLA_NEGOCIO"
            else:
                clasificacion = "VALIDACION"

            return rechazar_novedad(
                mensaje, clasificacion, "centro_costo"
            )

    # Validamos el teléfono.
    correcta, mensaje = validar_telefono(
        novedad, parametros
    )

    if not correcta:
        return rechazar_novedad(
            mensaje, "VALIDACION", "telefono"
        )

    # Revisamos las reglas del cargo y el centro de costo.
    correcta, mensaje = validar_regla_negocio(
        novedad, cargos, centros, matriz
    )

    if not correcta:
        return rechazar_novedad(
            mensaje, "REGLA_NEGOCIO", "cargo/centro_costo"
        )

    # Copiamos los datos para preparar la novedad aceptada.
    datos = novedad.copy()

    # Organizamos el tipo de novedad.
    tipo = str(novedad.get("tipo", "")).strip().upper()
    datos["tipo"] = tipo
    datos["tipo_novedad"] = tipo

    # Limpiamos el documento.
    documento = str(
        novedad.get("numero_documento", "")
    ).strip().upper()

    datos["numero_documento"] = re.sub(
        r"[.\-\s]", "", documento
    )

    # Organizamos el nombre completo.
    nombres = " ".join(
        str(novedad.get("nombres", "")).split()
    )
    apellidos = " ".join(
        str(novedad.get("apellidos", "")).split()
    )

    nombre_completo = " ".join(
        (nombres + " " + apellidos).split()
    ).title()

    if novedad.get("nombre_completo"):
        nombre_completo = " ".join(
            str(novedad["nombre_completo"]).split()
        ).title()

    datos["nombre_completo"] = nombre_completo

    # Buscamos los datos del cargo.
    cargo = str(
        novedad.get("cargo", "")
    ).strip().upper()

    datos["codigo_cargo"] = cargo

    if cargo in cargos:
        informacion_cargo = cargos[cargo]
        datos["nombre_cargo"] = informacion_cargo[0]
        datos["nivel"] = informacion_cargo[1]
        datos["area"] = informacion_cargo[2]

    # Buscamos el perfil y la aprobación en la matriz de roles.
    for fila in matriz:
        codigo = str(fila[0]).strip().upper()
        area = str(fila[1]).strip().upper()

        if (
            codigo == cargo
            and area == str(
                datos.get("area", "")
            ).strip().upper()
        ):
            datos["perfil_sistema"] = fila[2]
            datos["requiere_aprobacion"] = fila[3]
            break

    # Organizamos el centro de costo.
    datos["centro_costo"] = str(
        novedad.get("centro_costo", "")
    ).strip().upper()

    # Dejamos solamente los números del teléfono.
    telefono = str(
        novedad.get("telefono", "")
    ).strip()

    datos["telefono"] = PATRONES[
        "telefono_no_numerico"
    ].sub("", telefono)

    # Creamos el correo para las altas cuando no fue informado.
    correo = str(
        novedad.get("correo_corporativo", "")
    ).strip().lower()

    if tipo == "ALTA" and not correo and nombres and apellidos:
        nombre_correo = ".".join(nombres.lower().split())
        apellido_correo = ".".join(apellidos.lower().split())

        dominio = str(
            parametros.get("dominio_correo", "empresa.com")
        ).strip().lstrip("@")

        correo = (
            f"{nombre_correo}.{apellido_correo}@{dominio}"
        )

    # Conservamos el correo recibido o generado.
    # No rechazamos la novedad por el patrón de correo añadido.
    datos["correo_corporativo"] = correo

    # Convertimos la fecha efectiva al formato AAAA-MM-DD.
    fecha_texto = str(
        novedad.get("fecha_efectiva", "")
    ).strip()

    for formato in ("%d/%m/%Y", "%Y-%m-%d"):
        try:
            fecha = datetime.strptime(
                fecha_texto, formato
            )
            datos["fecha_efectiva"] = fecha.strftime(
                "%Y-%m-%d"
            )
            break
        except ValueError:
            continue

    # Extraemos la fecha de vinculación previa.
    datos["fecha_vinculacion_previa"] = (
        extraer_fecha_vinculacion(
            novedad.get("observaciones", "")
        )
    )

    # Indicamos que la novedad fue aceptada.
    datos["clasificacion"] = ""
    datos["campo"] = ""

    return True, "La novedad pasó las validaciones", datos