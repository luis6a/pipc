import requests
import os
from docxtpl import DocxTemplate, InlineImage
from docx import Document
from docx.shared import Inches

#################### CONFIGURACION DE USUARIO ####################

access_token = 'patwubbJ7vFDQCigL.8ac92361a7f8f099407c42a654ac0e166c794c07850f89723e862492a54c408b'
base_id = 'appy6PazgVEt6DWzU'
table_name = 'tblGLidPHPZP7M7ds'

# Valor específico para buscar
target_value = 'BBVA'  # Valor específico que deseas buscar en la columna "Categoría"
target_column = 'Categoría'  # Cambia 'Categoría' al nombre de la columna que contiene el valor

# URL para acceder a la tabla en Airtable
url = f'https://api.airtable.com/v0/{base_id}/{table_name}'

# Cabeceras para la solicitud
headers = {
    'Authorization': f'Bearer {access_token}'
}

# Parámetros para filtrar registros
params = {
    'filterByFormula': f"{{{target_column}}}='{target_value}'"
}

# Realiza la solicitud GET a la API de Airtable
response = requests.get(url, headers=headers, params=params)

# Verifica si la solicitud fue exitosa
if response.status_code == 200:
    print("Acceso a la base de datos exitoso.")
    data = response.json()
    records = data['records']
    print(f"Total de registros encontrados: {len(records)}")
else:
    print(f"Error: {response.status_code}")
    print("Revisa el token de acceso, el ID de la base y el nombre de la tabla.")
    records = []

# Filtrar el registro específico
target_record = None
if records:
    target_record = records[0]  # Obtiene el primer registro que coincide con el filtro
    print(f"Registro encontrado: {target_record}")
else:
    print(f"No se encontró ningún registro con {target_column} = {target_value}")

if not target_record:
    print(f'No se encontró ningún registro con {target_column} = {target_value}')
    exit()

# Crear una carpeta para almacenar las imágenes descargadas
os.makedirs('images', exist_ok=True)

def download_image(url, filename):
    response = requests.get(url)
    if response.status_code == 200:
        with open(filename, 'wb') as f:
            f.write(response.content)
        print(f"Imagen descargada: {filename}")
    else:
        print(f'Error al descargar la imagen: {response.status_code}')

def replace_placeholder(doc, placeholder, text):
    for paragraph in doc.paragraphs:
        if placeholder in paragraph.text:
            paragraph.text = paragraph.text.replace(placeholder, text)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if placeholder in cell.text:
                    cell.text = cell.text.replace(placeholder, text)

def add_image_at_placeholder(doc, placeholder, image_path, width=None):
    for paragraph in doc.paragraphs:
        if placeholder in paragraph.text:
            run = paragraph.clear().add_run()
            run.add_picture(image_path, width=width)
            break

# Descargar imágenes para el registro específico
fields = target_record['fields']
for key, value in fields.items():
    if isinstance(value, list) and all(isinstance(item, dict) for item in value):
        for item in value:
            if 'url' in item:
                url = item['url']
                filename = os.path.join('images', os.path.basename(url))
                download_image(url, filename)

# Crear una carpeta para los documentos de salida
os.makedirs('output', exist_ok=True)

# Seleccionar la plantilla correcta basada en el valor de la columna "Categoría"
category = fields.get('Categoría')  # Cambia 'Categoría' al nombre de tu columna en Airtable
print(f"Categoría del registro: {category}")

# Mapeo de categorías a plantillas
template_map = {
    'DHL': 'DHL_template.docx',
    'BBVA': 'BBVA_template.docx',
    'GDL': 'GDL_template.docx'
}

# Seleccionar la plantilla correspondiente
template_name = template_map.get(category, 'default_template.docx')  # Usa 'default_template.docx' como fallback
template_path = f'templates/{template_name}'  # Asegúrate de que las plantillas estén en esta carpeta
print(f"Plantilla seleccionada: {template_path}")

# Cargar la plantilla de Word
doc = Document(template_path)

# Reemplazar los placeholders con los valores de Airtable
for key, value in fields.items():
    if isinstance(value, list) and all(isinstance(item, dict) for item in value):
        for item in value:
            if 'url' in item:
                url = item['url']
                filename = os.path.join('images', os.path.basename(url))
                add_image_at_placeholder(doc, f'{{{{{key}}}}}', filename, width=Inches(2.0))
    else:
        replace_placeholder(doc, f'{{{{{key}}}}}', str(value))

# Guardar el documento
output_path = f'output/{fields.get("Name", "document")}_{category}.docx'  # Cambia 'Name' a una columna que tenga un identificador único o un nombre
doc.save(output_path)
print(f"Documento guardado en: {output_path}")