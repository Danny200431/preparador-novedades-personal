# Creamos un error general para nuestro proyecto
class ErrorProyecto(Exception):
    pass


# Error cuando la configuración no existe o está incorrecta
class ErrorConfiguracion(ErrorProyecto):
    pass


# Error cuando el archivo de entrada no existe o no se puede leer
class ErrorArchivoEntrada(ErrorProyecto):
    pass


# Error cuando una novedad tiene datos incorrectos
class RegistroInvalido(ErrorProyecto):
    pass


# Error cuando una novedad incumple una regla del negocio
class ErrorReglaNegocio(ErrorProyecto):
    pass


# Error cuando ocurre un problema al escribir un archivo
class ErrorEscritura(ErrorProyecto):
    pass


# Conservamos esta excepción por compatibilidad con el código existente
class ErrorArchivoSalida(ErrorEscritura):
    pass