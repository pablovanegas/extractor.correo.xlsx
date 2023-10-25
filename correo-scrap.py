import pandas as pd
import re

# Ruta del archivo Excel
xlsx_path = 'C:/Users/juanp/Downloads/base.xlsx'

# Inicializamos una lista para almacenar todos los correos encontrados
all_correos = []

# Cargamos el archivo Excel
xlsx = pd.ExcelFile(xlsx_path)

# Obtenemos los nombres de todas las hojas
sheet_names = xlsx.sheet_names

# Patrón para buscar direcciones de correo electrónico
email_pattern = r'\S+@\S+'

# Iteramos sobre cada hoja del Excel
for sheet in sheet_names:
    # Cargamos los datos de la hoja actual
    data_sheet = pd.read_excel(xlsx, sheet_name=sheet)

    # Iteramos sobre todas las columnas de la hoja
    for col_name in data_sheet.columns:
        # Verificamos si el nombre de la columna contiene "correo" o "correo electrónico"
        if 'correo' in col_name.lower() or 'correo electrónico' in col_name.lower():
            # Extraemos los correos de la columna encontrada
            correos_raw = data_sheet[col_name].dropna().astype(str)
            
            # Buscamos los correos en cada entrada y los añadimos a la lista de correos
            for entry in correos_raw:
                correos = re.findall(email_pattern, entry)
                all_correos.extend(correos)

# Eliminamos posibles duplicados y ordenamos los correos
all_correos_unique = sorted(set(all_correos))

# Guardamos los correos en un archivo de texto
output_file_path = 'C:/Users/juanp/Downloads/correos_todas_hojas.txt'
with open(output_file_path, 'w', encoding='utf-8') as file:
    file.write("\n".join(all_correos_unique))

# Devolvemos la ruta del archivo creado y el número de correos únicos extraídos
(output_file_path, len(all_correos_unique))
