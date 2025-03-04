import os
import shutil
import pandas as pd
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
from pyairtable import Table

#################### CONFIGURACION DE USUARIO ####################

# Configuración de Airtable
AIRTABLE_API_KEY = 'patwubbJ7vFDQCigL.8ac92361a7f8f099407c42a654ac0e166c794c07850f89723e862492a54c408b'
BASE_ID = 'appy6PazgVEt6DWzU'
TABLE_NAME = 'tblGLidPHPZP7M7ds'

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

# Función para obtener datos de Airtable

def obtener_datos_airtable():
    table = Table(AIRTABLE_API_KEY, BASE_ID, TABLE_NAME)

    # Filtrar por categoría y nombre comercial
    formula = f"AND({{Categoría}}='{CATEGORIA_BUSCAR}', {{Nombre Comercial}}='{NOMBRE_COMERCIAL_BUSCAR}')"
    records = table.all(formula=formula)

    if not records:
        print("No se encontraron registros con los criterios especificados.")
        return None

    # Tomamos el primer registro que coincida
    return records[0]['fields']

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

        # Preparar imágenes (si es necesario)
        # Ejemplo:
        try:
            img_logo1 = datos_airtable.get('logo1_nombre', 'logo1.jpg')
            img_path_logo1 = os.path.join(
                IMAGES_PATH, img_logo1)
            if os.path.exists(img_path_logo1):
                logo1 = InlineImage(docx_tpl, img_path_logo1, height=Mm(145))
            else:
                print(
                    f'Advertencia: No se encontró la imagen {datos_airtable.get("logo1", "")}')
                logo1 = ''
        except Exception as e:
            print(
                f'Advertencia: No se pudo cargar la imagen {datos_airtable.get("logo1", "")}: {e}')
            logo1 = ''

        try:
            img_logo2 = datos_airtable.get('logo2_nombre', 'logo2.jpg')
            img_path_logo2 = os.path.join(
                IMAGES_PATH, img_logo2)
            if os.path.exists(img_path_logo2):
                logo2 = InlineImage(docx_tpl, img_path_logo2, height=Mm(15))
            else:
                print(
                    f'Advertencia: No se encontró la imagen {datos_airtable.get("logo2", "")}')
                logo2 = ''
        except Exception as e:
            print(
                f'Advertencia: No se pudo cargar la imagen {datos_airtable.get("logo2", "")}: {e}')
            logo2 = ''

        # Crear contexto
        context = datos_airtable.copy()
        # Ya tienes todos los valores en context gracias al .copy(), pero si necesitas 
        # asegurarte de que existan con valores predeterminados, usa:
        context['nombre_comercial'] = datos_airtable.get('Nombre Comercial', '')
        context['razon_social'] = datos_airtable.get('Razón Social', '')
        context['RFC'] = datos_airtable.get('RFC', '')
        context['logo1'] = logo1
        context['logo2'] = logo2
        # Agrega aquí más procesamiento de imágenes si es necesario

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
