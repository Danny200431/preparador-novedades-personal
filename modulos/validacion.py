# Este archivo contiene las funciones que revisan los datos
# de cada novedad de personal.

import re
from datetime import datetime, date, timedelta


# Preparamos las expresiones regulares una sola vez.
# Después las reutilizamos en las diferentes validaciones.

PATRONES = {
    # Documento de identidad: entre 6 y 10 números.
    "documento_cc": re.compile(r"^[0-9]{6,10}$"),

    # Documento de extranjería: números o una letra seguida de números.
    "documento_ce": re.compile(r"^[0-9]{6,10}$|^[A-Z][0-9]{6,9}$"),

    # Pasaporte: entre 6 y 12 caracteres, con al menos una letra.
    "documento_pa": re.compile(r"^(?=.*[A-Z])[A-Z0-9]{6,12}$"),

    # Centro de costo: CC- seguido de cuatro números.
    "centro_costo": re.compile(r"^CC-[0-9]{4}$"),

    # Permite encontrar fechas en los dos formatos admitidos.
    "fecha_observaciones": re.compile(
        r"\b\d{2}/\d{2}/\d{4}\b|\b\d{4}-\d{2}-\d{2}\b"
    ),

    # Permite separar los números de un teléfono.
    "telefono_no_numerico": re.compile(r"\D"),

    # Formato general de un correo electrónico.
    "correo": re.compile(
        r"^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$"
    ),
}


def validar_tipo(novedad, campos_obligatorios):
    """Comprueba que el tipo de novedad esté configurado."""

    tipo = str(novedad.get("tipo", "")).strip().upper()

    if tipo not in campos_obligatorios:
        return False, "El tipo de novedad no es válido"

    return True, ""


def validar_campos(novedad, campos_obligatorios):
    """Comprueba los campos obligatorios definidos en el Excel."""

    tipo = str(novedad.get("tipo", "")).strip().upper()
    campos = campos_obligatorios.get(tipo, [])

    for campo in campos:
        valor = novedad.get(campo, "")

        if valor is None or not str(valor).strip():
            return False, f"Falta el campo obligatorio: {campo}"

    return True, ""


def validar_documento(novedad):
    """Comprueba el tipo y el formato del documento."""

    tipo_documento = str(
        novedad.get("tipo_documento", "")
    ).strip().upper()

    documento = str(
        novedad.get("numero_documento", "")
    ).strip().upper()

    # Quitamos puntos, guiones y espacios antes de validar.
    documento = re.sub(r"[.\-\s]", "", documento)

    if not documento:
        return False, "Falta el número de documento"

    if tipo_documento == "CC":
        if not PATRONES["documento_cc"].fullmatch(documento):
            return False, (
                "La CC debe tener entre 6 y 10 dígitos numéricos"
            )

    elif tipo_documento == "CE":
        if not PATRONES["documento_ce"].fullmatch(documento):
            return False, "La CE no cumple el formato permitido"

    elif tipo_documento == "PA":
        if not PATRONES["documento_pa"].fullmatch(documento):
            return False, (
                "El PA debe tener entre 6 y 12 caracteres "
                "y al menos una letra"
            )

    else:
        return False, "El tipo de documento no es válido"

    return True, ""


def validar_fecha(novedad, dias_maximo_futuro):
    """Comprueba el formato y el límite de la fecha efectiva."""

    fecha_texto = str(
        novedad.get("fecha_efectiva", "")
    ).strip()

    fecha = None

    # Aceptamos los dos formatos de fecha establecidos.
    for formato in ("%d/%m/%Y", "%Y-%m-%d"):
        try:
            fecha = datetime.strptime(
                fecha_texto, formato
            ).date()
            break
        except ValueError:
            continue

    if fecha is None:
        return False, "La fecha no tiene un formato válido"

    hoy = date.today()

    if fecha < hoy:
        return False, "La fecha efectiva ya pasó"

    try:
        dias_maximos = int(dias_maximo_futuro)
    except (ValueError, TypeError):
        return False, "El límite de días futuros no es válido"

    if dias_maximos < 0:
        return False, "El límite de días futuros no es válido"

    fecha_maxima = hoy + timedelta(days=dias_maximos)

    if fecha > fecha_maxima:
        return False, "La fecha supera el límite permitido"

    return True, ""


def validar_centro_costo(novedad, centros):
    """Comprueba el formato y la existencia del centro de costo."""

    centro = str(
        novedad.get("centro_costo", "")
    ).strip().upper()

    if not PATRONES["centro_costo"].fullmatch(centro):
        return False, (
            "El centro de costo no tiene el formato CC-0000"
        )

    if centro not in centros:
        return False, "El centro de costo no existe en el catálogo"

    return True, ""


def validar_telefono(novedad, parametros):
    """Comprueba la longitud del teléfono cuando está informado."""

    telefono = str(novedad.get("telefono", "")).strip()

    # El teléfono puede estar vacío si no es obligatorio.
    if not telefono:
        return True, ""

    digitos = PATRONES["telefono_no_numerico"].sub("", telefono)

    try:
        longitud = int(parametros.get("longitud_telefono", 10))
    except (ValueError, TypeError):
        return False, "La longitud configurada del teléfono no es válida"

    if len(digitos) != longitud:
        return False, f"El teléfono debe tener {longitud} dígitos"

    return True, ""


def validar_correo(correo):
    """Comprueba que el correo tenga un formato válido."""

    correo = str(correo or "").strip()

    if not correo:
        return False, "El correo está vacío"

    if not PATRONES["correo"].fullmatch(correo):
        return False, "El correo electrónico no tiene un formato válido"

    return True, ""


def validar_regla_negocio(novedad, cargos, centros, matriz):
    """Comprueba que el cargo y el centro de costo sean compatibles."""

    cargo = str(novedad.get("cargo", "")).strip().upper()
    centro = str(novedad.get("centro_costo", "")).strip().upper()
    tipo = str(novedad.get("tipo", "")).strip().upper()

    # Para los retiros no se aplica esta comprobación.
    if tipo == "RETIRO":
        return True, ""

    # Si falta uno de los datos, los campos obligatorios
    # se encargan de informar el problema cuando corresponda.
    if not cargo or not centro:
        return True, ""

    if cargo not in cargos:
        return False, "El cargo no existe en el catálogo"

    if centro not in centros:
        return False, "El centro de costo no existe en el catálogo"

    area_cargo = str(cargos[cargo][2]).strip().upper()
    area_centro = str(centros[centro][1]).strip().upper()

    if area_cargo != area_centro:
        return False, (
            "El área del cargo no coincide con el centro de costo"
        )

    # Revisamos que el cargo y el área estén permitidos
    # en la matriz de roles del Excel.
    for fila in matriz:
        codigo_cargo = str(fila[0]).strip().upper()
        area = str(fila[1]).strip().upper()

        if codigo_cargo == cargo and area == area_centro:
            return True, ""

    return False, "El cargo y el área no están permitidos"