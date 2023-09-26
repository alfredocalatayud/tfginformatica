import pandas as pd
import requests
from progress.bar import Bar
import openpyxl as op

# Diccionario de descripciones de estados HTTP
descripciones_http = {
    100: "Continuar",
    101: "Cambiando protocolos",
    102: "Procesando",
    200: "OK",
    201: "Creado",
    202: "Aceptado",
    203: "Información no autoritativa",
    204: "Sin contenido",
    205: "Restablecer contenido",
    206: "Contenido parcial",
    207: "Estado multiestado",
    208: "Estado ya informado",
    226: "IM utilizado",
    300: "Múltiples opciones",
    301: "Movido permanentemente",
    302: "Encontrado",
    303: "Vea otros",
    304: "No modificado",
    305: "Usar proxy",
    307: "Redirección temporal",
    308: "Redirección permanente",
    400: "Solicitud incorrecta",
    401: "No autorizado",
    402: "Pago requerido",
    403: "Prohibido",
    404: "No encontrado",
    405: "Método no permitido",
    406: "No aceptable",
    407: "Autenticación de proxy requerida",
    408: "Solicitud de tiempo de espera",
    409: "Conflicto",
    410: "Ya no disponible",
    411: "Longitud requerida",
    412: "Falló la precondición",
    413: "Solicitud de entidad demasiado grande",
    414: "URI demasiado larga",
    415: "Tipo de medio no soportado",
    416: "Rango solicitado no satisfactorio",
    417: "Falló la expectativa",
    418: "Soy una tetera",
    421: "Solicitud mal dirigida",
    422: "Entidad no procesable",
    423: "Bloqueado",
    424: "Fallo dependiente",
    426: "Actualización requerida",
    428: "Precondición obligatoria",
    429: "Demasiadas solicitudes",
    431: "Campos de encabezado de solicitud demasiado grandes",
    451: "No disponible por razones legales",
    500: "Error interno del servidor",
    501: "No implementado",
    502: "Puerta de enlace incorrecta",
    503: "Servicio no disponible",
    504: "Tiempo de espera de la puerta de enlace",
    505: "Versión HTTP no compatible",
    506: "Variante también negocia",
    507: "Espacio insuficiente de almacenamiento",
    508: "Ciclo detectado",
    510: "No extendido",
    511: "Autenticación de red requerida",
}

db_xlsx = "ayuntamientos_web_v2.xlsx"
# workbook = op.load_workbook(db_xlsx)
#
# sheet = workbook.active
#
# column_name = "WebAyuntamiento"
#
# webs = []
#
# for row in sheet.iter_rows(values_only=True):
#     row_list = list(row)  # Convierte la tupla en una lista
#     if column_name in sheet[1]:  # Verifica si el nombre de la columna está en la primera fila
#         web = row_list[sheet[1].index(column_name)]  # Utiliza index en la lista
#         webs.append(web)
#
# for web in webs:
#     print(web)
#
# workbook.close()

try:
    df = pd.read_excel(db_xlsx)
except FileNotFoundError:
    print(f"El archivo {db_xlsx} no se encontró.")
    exit()
except Exception as e:
    print(f"Ocurrió un error al leer el archivo: {str(e)}")
    exit()

l_http_status = []

urls = df["WebAyuntamiento"]

n_urls = len(urls)
bar = Bar('Probando URLs:', max=n_urls)

n_urls = 0
n_urls_off = 0
urls_off = []

for url in urls:
    bar.next()
    try:
        response = requests.get(url)
        http_status = response.status_code
        l_http_status.append(http_status)

        if http_status != 200:
            n_urls_off += 1
            description = descripciones_http.get(http_status, "Estado desconocido")
            urls_off.append((url, http_status, description))
    except Exception as e:
        http_status = None
        l_http_status.append(http_status)

    # n_urls += 1
    #
    # if n_urls >= 100:
    #     break

bar.finish()

with open('informe.txt', 'w') as txtfile:
    txtfile.write("Informe de Estados HTTP:\n\n")
    txtfile.write(f"Número de URLs revisadas: {n_urls}\n")
    txtfile.write(f"Número de URLs con estado distinto de 200: {n_urls_off}\n\n")

    if n_urls_off > 0:
        txtfile.write("URLs con estado distinto de 200:\n")
        for url, status, description in urls_off:
            txtfile.write(f"URL: {url}\nEstado HTTP: {status} ({description})\n\n")

print("Informe generado exitosamente en 'informe.txt'")