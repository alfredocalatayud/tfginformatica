import pandas as pd
import requests
from progress.bar import Bar
import openpyxl as op

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
bar = Bar('Probando URLs:', max=40)

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
            urls_off.append((url, http_status))
    except Exception as e:
        http_status = None
        l_http_status.append(http_status)

    n_urls += 1

    if n_urls >= 40:
        break

bar.finish()

with open('informe.txt', 'w') as txtfile:
    txtfile.write("Informe de Estados HTTP:\n\n")
    txtfile.write(f"Número de URLs revisadas: {n_urls}\n")
    txtfile.write(f"Número de URLs con estado distinto de 200: {n_urls_off}\n\n")

    if n_urls_off > 0:
        txtfile.write("URLs con estado distinto de 200:\n")
        for url, status in urls_off:
            txtfile.write(f"URL: {url}\nEstado HTTP: {status}\n\n")

print("Informe generado exitosamente en 'informe.txt'")