import os
import shutil
import pandas as pd
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
from pyairtable import Api

#################### CONFIGURACION DE USUARIO ####################

# Configuración de Airtable
AIRTABLE_API_KEY = 'patwubbJ7vFDQCigL.8ac92361a7f8f099407c42a654ac0e166c794c07850f89723e862492a54c408b'
BASE_ID = 'appy6PazgVEt6DWzU'
TABLE_NAME = 'tblGLidPHPZP7M7ds'

# Tablas relacionadas en Airtable
INMUEBLE_TABLE = 'Inmueble'
POBLACION_TABLE = 'Población'
BRIGADAS_TABLE = 'Brigadas'
INVENTARIO_TABLE = 'Inventario'
EQUIPO_TABLE = 'Equipo'
OTRO_EQUIPO_TABLE = 'Otro Equipo'
RIESGOS_TABLE = 'Riesgos'
GRI_TABLE = 'GRI'
GASOLINERA_TABLE = 'Gasolinera'
PERITO_TABLE = 'Perito'

# Valores específicos para buscar en Airtable
CATEGORIA_BUSCAR = 'GASOLINERA'
NOMBRE_COMERCIAL_BUSCAR = 'SAN SEBASTIAN E.S. 7335 VALERO'

# Ruta de salida
OUTPUT_PATH = '.\Outputs'

# Ruta fichero Excel (para datos adicionales si es necesario)
EXCEL_PATH = '.\Inputs\BD.xlsx'

# Rutas de plantillas Word (mantén las que necesites)
GASOLINERA_WORD_PTLL_PATH = '.\Inputs\Templates\Gasolinera.docx'
GAS_MF_WORD_PTLL_PATH = '.\Inputs\Templates\MF Gasolinera.docx'
GRI_GAS_WORD_PTLL_PATH = '.\Inputs\Templates\GRI Gasolinera.docx'
CIDUR_PTLL_PATH = '.\Inputs\Templates\Cidur.docx'
CIDUR_MF_PTLL_PATH = '.\Inputs\Templates\MF Cidur.docx'
CIDUR_GRI_PTLL_PATH = '.\Inputs\Templates\GRI Cidur.docx'
GDL_PTLL_PATH = '.\Inputs\Templates\Farmcia_GDL.docx'
GDL_MF_PTLL_PATH = '.\Inputs\Templates\MF Farmcia_GDL.docx'
GDL_GRI_PTLL_PATH = '.\Inputs\Templates\GRI Farmcia_GDL.docx'
GENERAL_PTLL_PATH = '.\Inputs\Templates\General.docx'
GENERAL_MF_PTLL_PATH = '.\Inputs\Templates\MF General.docx'
GENERAL_GRI_PTLL_PATH = '.\Inputs\Templates\GRI General.docx'
BBVA_PTLL_PATH = '.\Inputs\Templates\BBVA.docx'
BBVA_MF_PTLL_PATH = '.\Inputs\Templates\MF BBVA.docx'
BBVA_GRI_PTLL_PATH = '.\Inputs\Templates\GRI BBVA.docx'
COMPARTAMOS_PTLL_PATH = '.\Inputs\Templates\Compartamos.docx'
COMPARTAMOS_MF_PTLL_PATH = '.\Inputs\Templates\MF Compartamos.docx'
COMPARTAMOS_GRI_PTLL_PATH = '.\Inputs\Templates\GRI Compartamos.docx'
ALL_GOWER_PTLL_PATH = '.\Inputs\Templates\Cartas Gower.docx'
ALL_NOE_PTLL_PATH = '.\Inputs\Templates\Cartas Noe.docx'
UVP_PTLL_PATH = '.\Inputs\Templates\PIPC UVP.docx'
DHL_PTLL_PATH = '.\Inputs\Templates\DHL.docx'

# Ruta imágenes
IMAGES_PATH = '.\Inputs\Images'

#################### CONFIGURACION DE USUARIO ####################

# Eliminar y crear carpetas

def eliminar_crear_carpetas(path):
    if os.path.exists(path):
        shutil.rmtree(path)
    os.mkdir(path)

# Función para obtener datos de tablas relacionadas
def obtener_datos_relacionados(api, nombre_comercial, tabla_nombre):
    tabla = api.table(BASE_ID, tabla_nombre)
    
    # Usar diferente columna de búsqueda según la tabla
    if tabla_nombre == INMUEBLE_TABLE:
        # Para la tabla Inmueble, buscamos por Nombre Comercial
        formula = f"{{Nombre Comercial}}='{nombre_comercial}'"
    else:
        # Para las otras tablas, buscamos por el campo Inmueble
        formula = f"{{Inmueble}}='{nombre_comercial}'"
    
    records = tabla.all(formula=formula)
    
    if not records:
        print(f"No se encontraron registros en la tabla {tabla_nombre} para {nombre_comercial}.")
        return {}
    
    # Si hay múltiples registros, los combinamos en un solo diccionario
    if len(records) > 1:
        print(f"Se encontraron {len(records)} registros en la tabla {tabla_nombre} para {nombre_comercial}.")
        # Para listas o campos que pueden tener múltiples valores, podemos combinarlos
        combined_fields = {}
        for idx, record in enumerate(records, 1):
            for key, value in record['fields'].items():
                # Si es una lista, extendemos
                if isinstance(value, list) and key in combined_fields and isinstance(combined_fields[key], list):
                    combined_fields[key].extend(value)
                # Si no es una lista o el campo no existe, lo asignamos directamente
                # O añadimos un sufijo numérico para diferenciar
                elif key in combined_fields:
                    combined_fields[f"{key}_{idx}"] = value
                else:
                    combined_fields[key] = value
        return combined_fields
    
    # Si solo hay un registro, lo devolvemos directamente
    return records[0]['fields']

# Función para obtener datos de Airtable y tablas relacionadas
def obtener_datos_airtable():
    api = Api(AIRTABLE_API_KEY)
    table = api.table(BASE_ID, TABLE_NAME)

    # Filtrar por categoría y nombre comercial
    formula = f"AND({{Categoría}}='{CATEGORIA_BUSCAR}', {{Nombre Comercial}}='{NOMBRE_COMERCIAL_BUSCAR}')"
    records = table.all(formula=formula)

    if not records:
        print("No se encontraron registros con los criterios especificados.")
        return None

    # Tomamos el primer registro que coincida
    datos_principales = records[0]['fields']
    nombre_comercial = datos_principales.get('Nombre Comercial', '')

 # Obtenemos datos de tablas relacionadas
    tablas_relacionadas = {
        'inmueble': INMUEBLE_TABLE,
        'poblacion': POBLACION_TABLE,
        'brigadas': BRIGADAS_TABLE,
        'inventario': INVENTARIO_TABLE,
        'equipo': EQUIPO_TABLE,
        'otro_equipo': OTRO_EQUIPO_TABLE,
        'riesgos': RIESGOS_TABLE,
        'gri': GRI_TABLE,
        'gasolinera': GASOLINERA_TABLE,
        'perito': PERITO_TABLE
    }
    
    datos_completos = {}  # Creamos un diccionario vacío para los datos completos
    
    # Añadir datos de tablas relacionadas con prefijos para evitar colisiones
    for prefijo, tabla in tablas_relacionadas.items():
        datos_tabla = obtener_datos_relacionados(api, nombre_comercial, tabla)
        # Añadimos los datos con prefijo para evitar colisiones
        for k, v in datos_tabla.items():
            datos_completos[f"{prefijo}_{k}"] = v
            # También añadimos los campos sin prefijo para facilitar el acceso en la plantilla
            # Esto podría sobreescribir datos si hay nombres de campo repetidos entre tablas
            if k not in datos_completos:  # Solo añadimos si no existe ya para evitar sobrescribir
                datos_completos[k] = v
    
    return datos_completos

# Función para cargar imágenes
def cargar_imagen(docx_tpl, imagen_nombre_campo, imagen_default, medida_mm, tipo_medida='height', datos_airtable=None):
    try:
        img_nombre = datos_airtable.get(f'{imagen_nombre_campo}_nombre', imagen_default)
        img_path = os.path.join(IMAGES_PATH, img_nombre)
        
        if os.path.exists(img_path):
            # Aplicar altura o ancho según corresponda
            if tipo_medida.lower() == 'width':
                return InlineImage(docx_tpl, img_path, width=Mm(medida_mm))
            else:  # Por defecto usa height
                return InlineImage(docx_tpl, img_path, height=Mm(medida_mm))
        else:
            print(f'Advertencia: No se encontró la imagen {img_nombre}')
            return ''
    except Exception as e:
        print(f'Advertencia: No se pudo cargar la imagen {img_nombre}: {e}')
        return ''

# Función para crear ficheros Word
def crear_word(datos_airtable):
    if not datos_airtable:
        print("No hay datos para procesar.")
        return

    # Determinar qué plantillas usar basado en la categoría
    categoria = datos_airtable.get('Categoría', '')

    if categoria == 'GASOLINERA':
        plantillas = [GASOLINERA_WORD_PTLL_PATH,
                      GAS_MF_WORD_PTLL_PATH, GRI_GAS_WORD_PTLL_PATH]
    elif categoria == 'BANCO':
        plantillas = [CIDUR_PTLL_PATH, CIDUR_MF_PTLL_PATH,
                      CIDUR_GRI_PTLL_PATH, ALL_GOWER_PTLL_PATH, ALL_NOE_PTLL_PATH]
    elif categoria == 'GDL':
        plantillas = [GDL_PTLL_PATH, GDL_MF_PTLL_PATH,
                      GDL_GRI_PTLL_PATH, ALL_GOWER_PTLL_PATH, ALL_NOE_PTLL_PATH]
    elif categoria == 'GENERAL':
        plantillas = [GENERAL_PTLL_PATH, GENERAL_MF_PTLL_PATH,
                      GENERAL_GRI_PTLL_PATH, ALL_GOWER_PTLL_PATH, ALL_NOE_PTLL_PATH]
    elif categoria == 'BBVA':
        plantillas = [BBVA_PTLL_PATH, BBVA_MF_PTLL_PATH,
                      BBVA_GRI_PTLL_PATH, ALL_GOWER_PTLL_PATH, ALL_NOE_PTLL_PATH]
    elif categoria == 'COMPARTAMOS':
        plantillas = [COMPARTAMOS_PTLL_PATH, COMPARTAMOS_MF_PTLL_PATH,
                      COMPARTAMOS_GRI_PTLL_PATH, ALL_GOWER_PTLL_PATH, ALL_NOE_PTLL_PATH]
    elif categoria == 'UVP':
        plantillas = [UVP_PTLL_PATH, GENERAL_MF_PTLL_PATH,
                      GENERAL_GRI_PTLL_PATH, ALL_GOWER_PTLL_PATH, ALL_NOE_PTLL_PATH]
    elif categoria == 'DHL':
        plantillas = [DHL_PTLL_PATH, GENERAL_MF_PTLL_PATH,
                      GENERAL_GRI_PTLL_PATH, ALL_GOWER_PTLL_PATH, ALL_NOE_PTLL_PATH]
    else:
        print(f"No se encontraron plantillas para la categoría: {categoria}")
        return

    for idx, plantilla_path in enumerate(plantillas, start=1):
        # Cargar plantilla
        docx_tpl = DocxTemplate(plantilla_path)

        # Cargar las imágenes
        logo1 = cargar_imagen(docx_tpl, 'logo1', 'logo1.jpg', 145, 'height',datos_airtable)
        logo2 = cargar_imagen(docx_tpl, 'logo2', 'logo2.jpg', 15, 'width',datos_airtable)

        # Mapeo de nombres de campos y sus equivalentes con valores predeterminados
        campo_mapping = {
            # Campos de la tabla Inmueble (ejemplo)
            'nombre_comercial': ('Nombre Comercial', ''),
            'razon_social': ('Razón Social', ''),
            'RFC': ('RFC', ''),
            # Campos de la tabla Población (ejemplo)
            'representante_legal': ('Representante Legal', ''),
            'hombres': ('No Hombres', ''),
            # Campos de la tabla Brigadas (ejemplo)
            'responsable_pipc': ('Coordinador', ''),
            'coord_puesto': ('Puesto Coordinador', ''),
            # Campos de la tabla Inventario (ejemplo)
            'ruta_evac': ('Ruta evacuación', ''),
            'salida_emerg': ('Salida Emergencia', '')
        }

        # Crear contexto y aplicar mapeo en una sola operación
        context = {
            **datos_airtable,
            **{key_context: datos_airtable.get(key_airtable, default_value) 
               for key_context, (key_airtable, default_value) in campo_mapping.items()},
            'logo1': logo1,
            'logo2': logo2
        }

        try:
            # Renderizar documento
            docx_tpl.render(context)

            # Determinar nombre del archivo de salida
            if idx == 1:
                nombre_pipc = f'1. PIPC {datos_airtable["Nombre Comercial"]}.docx'
            elif idx == 2:
                nombre_pipc = f'2. MEMORIA FOTOGRAFICA {datos_airtable["Nombre Comercial"]}.docx'
            elif idx == 3:
                nombre_pipc = f'3. RIESGO DE INCENDIO {datos_airtable["Nombre Comercial"]}.docx'

            # Guardar el documento
            docx_tpl.save(os.path.join(OUTPUT_PATH, nombre_pipc))
            print(f"Documento creado: {nombre_pipc}")

        except Exception as e:
            print(f'Error al guardar el documento: {str(e)}')

# Función principal
def main():
    # Eliminar y volver a crear carpeta 'Outputs'
    eliminar_crear_carpetas(OUTPUT_PATH)

    # Obtener datos de Airtable
    datos_airtable = obtener_datos_airtable()

    # Crear ficheros Word
    crear_word(datos_airtable)


if __name__ == '__main__':
    main()
