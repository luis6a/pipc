import os
import shutil
import stat
import time
import pandas as pd
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
from pyairtable import Api
from dotenv import load_dotenv

#################### CONFIGURACION DE USUARIO ####################

# Cargar variables de entorno desde .env
load_dotenv()

# Configuración de Airtable
AIRTABLE_API_KEY = os.getenv("AIRTABLE_TOKEN")
BASE_ID = 'appy6PazgVEt6DWzU'
TABLE_NAME = 'tblGLidPHPZP7M7ds'

# Tablas relacionadas en Airtablevenv\Scripts\activate
INMUEBLE_TABLE = 'Inmueble'
POBLACION_TABLE = 'Poblacion'
OTRO_EQUIPO_TABLE = 'Otro Equipo'
GASOLINERA_TABLE = 'Gasolinera'
PERITO_TABLE = 'Perito'

# Valores específicos para buscar en Airtable
CATEGORIA_BUSCAR = 'GENERAL'
NOMBRE_COMERCIAL_BUSCAR = 'COMEX TEPETITLA'

# Ruta base absoluta
BASE_DIR = "C:/Users/luis6/OneDrive/Documentos/GitHub/PIPC"

# Ruta de salida
OUTPUT_PATH = f"{BASE_DIR}/Outputs"

# Ruta fichero Excel
EXCEL_PATH = f"{BASE_DIR}/Inputs/BD.xlsx"

# Ruta plantillas ficheros Word
GASOLINERA_WORD_PTLL_PATH = f"{BASE_DIR}/Inputs/Templates/Gasolinera.docx"
GASOLINERA_MF_WORD_PTLL_PATH = f"{BASE_DIR}/Inputs/Templates/MF Gasolinera.docx"
GASOLINERA_GRI_WORD_PTLL_PATH = f"{BASE_DIR}/Inputs/Templates/GRI Gasolinera.docx"
CIDUR_PTLL_PATH = f"{BASE_DIR}/Inputs/Templates/Banorte.docx"
CIDUR_MF_PTLL_PATH = f"{BASE_DIR}/Inputs/Templates/MF Banorte.docx"
CIDUR_GRI_PTLL_PATH = f"{BASE_DIR}/Inputs/Templates/GRI Banorte.docx"
GDL_PTLL_PATH = f"{BASE_DIR}/Inputs/Templates/Farmcia_GDL.docx"
GDL_MF_PTLL_PATH = f"{BASE_DIR}/Inputs/Templates/MF Farmcia_GDL.docx"
GDL_GRI_PTLL_PATH = f"{BASE_DIR}/Inputs/Templates/GRI Farmcia_GDL.docx"
GENERAL_PTLL_PATH = f"{BASE_DIR}/Inputs/Templates/General.docx"
GENERAL_MF_PTLL_PATH = f"{BASE_DIR}/Inputs/Templates/MF General.docx"
GENERAL_GRI_PTLL_PATH = f"{BASE_DIR}/Inputs/Templates/GRI General.docx"
BBVA_PTLL_PATH = f"{BASE_DIR}/Inputs/Templates/BBVA.docx"
BBVA_MF_PTLL_PATH = f"{BASE_DIR}/Inputs/Templates/MF BBVA.docx"
BBVA_GRI_PTLL_PATH = f"{BASE_DIR}/Inputs/Templates/GRI BBVA.docx"
ALL_NOE_PTLL_PATH = f"{BASE_DIR}/Inputs/Templates/Cartas Noe.docx"
UVP_PTLL_PATH = f"{BASE_DIR}/Inputs/Templates/PIPC UVP.docx"
UVP_MF_PTLL_PATH = f"{BASE_DIR}/Inputs/Templates/MF UVP.docx"
UVP_GRI_PTLL_PATH = f"{BASE_DIR}/Inputs/Templates/GRI UVP.docx"
DHL_PTLL_PATH = f"{BASE_DIR}/Inputs/Templates/DHL.docx"
DHL_MF_PTLL_PATH = f"{BASE_DIR}/Inputs/Templates/MF DHL.docx"
DHL_GRI_PTLL_PATH = f"{BASE_DIR}/Inputs/Templates/GRI DHL.docx"
GASERA_WORD_PTLL_PATH = f"{BASE_DIR}/Inputs/Templates/Gasera.docx"
GASERA_MF_WORD_PTLL_PATH = f"{BASE_DIR}/Inputs/Templates/MF Gasera.docx"
ESTAFETA_PTLL_PATH = f"{BASE_DIR}/Inputs/Templates/Estafeta.docx"
ESTAFETA_MF_PTLL_PATH = f"{BASE_DIR}/Inputs/Templates/MF Estafeta.docx"
ESTAFETA_GRI_PTLL_PATH = f"{BASE_DIR}/Inputs/Templates/GRI Estafeta.docx"
FEDEX_PTLL_PATH = f"{BASE_DIR}/Inputs/Templates/Fedex.docx"
FEDEX_MF_PTLL_PATH = f"{BASE_DIR}/Inputs/Templates/MF Fedex.docx"
FEDEX_GRI_PTLL_PATH = f"{BASE_DIR}/Inputs/Templates/GRI Fedex.docx"
SMARTFIT_PTLL_PATH = f"{BASE_DIR}/Inputs/Templates/SmartFit.docx"
SMARTFIT_MF_PTLL_PATH = f"{BASE_DIR}/Inputs/Templates/MF SmartFit.docx"
SMARTFIT_GRI_PTLL_PATH = f"{BASE_DIR}/Inputs/Templates/GRI SmartFit.docx"
AXA_PTLL_PATH = f"{BASE_DIR}/Inputs/Templates/Axa.docx"
AXA_MF_PTLL_PATH = f"{BASE_DIR}/Inputs/Templates/MF Axa.docx"
AXA_GRI_PTLL_PATH = f"{BASE_DIR}/Inputs/Templates/GRI Axa.docx"
CONTINGENCIA_PTLL_PATH = f"{BASE_DIR}/Inputs/Templates/Contingencia.docx"
ALL_LEVANTAMIENTO_PTLL_PATH = f"{BASE_DIR}/Inputs/Templates/Levantamiento.docx"

# Ruta imágenes
IMAGES_PATH = f"{BASE_DIR}/Inputs/Images"

#################### CONFIGURACION DE USUARIO ####################

# Eliminar y crear carpetas

def eliminar_crear_carpetas(path):
    if os.path.exists(path):

        def on_rm_exc(func, path, exc_info):
            os.chmod(path, stat.S_IWRITE)
            func(path)

        try:
            shutil.rmtree(path, onexc=on_rm_exc)
        except PermissionError:
            time.sleep(1)
            shutil.rmtree(path, onexc=on_rm_exc)

    os.makedirs(path, exist_ok=True)

# Función para obtener datos de tablas relacionadas

def obtener_datos_relacionados(api, nombre_comercial, tabla_nombre):
    tabla = api.table(BASE_ID, tabla_nombre)

    # Usar diferente columna de búsqueda según la tabla
    if tabla_nombre == INMUEBLE_TABLE:
        # Para la tabla Inmueble, buscamos por Nombre Comercial
        formula = f"{{nombre_comercial}}='{nombre_comercial}'"
    else:
        # Para las otras tablas, buscamos por el campo Inmueble
        formula = f"{{Inmueble}}='{nombre_comercial}'"

    records = tabla.all(formula=formula)

    if not records:
        print(
            f"No se encontraron registros en la tabla {tabla_nombre} para {nombre_comercial}")
        return {}

    # Si hay múltiples registros, los combinamos en un solo diccionario
    if len(records) > 1:
        print(
            f"Se encontraron {len(records)} registros en la tabla {tabla_nombre} para {nombre_comercial}")
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
    formula = f"AND({{Categoria}}='{CATEGORIA_BUSCAR}', {{nombre_comercial}}='{NOMBRE_COMERCIAL_BUSCAR}')"
    records = table.all(formula=formula)

    if not records:
        print("No se encontraron registros con los criterios especificados")
        return None

    # Tomamos el primer registro que coincida
    datos_principales = records[0]['fields']
    nombre_comercial = datos_principales.get('nombre_comercial', '')

 # Obtenemos datos de tablas relacionadas
    tablas_relacionadas = {
        'inmueble': INMUEBLE_TABLE,
        'poblacion': POBLACION_TABLE,
        'otro_equipo': OTRO_EQUIPO_TABLE,
        'gasolinera': GASOLINERA_TABLE,
        'perito': PERITO_TABLE
    }

    datos_completos = {}  # Creamos un diccionario vacío para los datos completos

    # Añadir datos principales primero (sin prefijo)
    datos_completos.update(datos_principales)

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
        img_nombre = datos_airtable.get(
            f'{imagen_nombre_campo}_nombre', imagen_default)
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
        print("No hay datos para procesar")
        return

    # Determinar qué plantillas usar basado en la categoría
    categoria = datos_airtable.get('Categoria', '')

    if categoria == 'GASOLINERA':
        plantillas = [GASOLINERA_WORD_PTLL_PATH,
                      GASOLINERA_MF_WORD_PTLL_PATH, GASOLINERA_GRI_WORD_PTLL_PATH, ALL_NOE_PTLL_PATH, ALL_LEVANTAMIENTO_PTLL_PATH]
    elif categoria == 'BANORTE':
        plantillas = [CIDUR_PTLL_PATH, CIDUR_MF_PTLL_PATH,
                      CIDUR_GRI_PTLL_PATH, ALL_NOE_PTLL_PATH, ALL_LEVANTAMIENTO_PTLL_PATH]
    elif categoria == 'GDL':
        plantillas = [GDL_PTLL_PATH, GDL_MF_PTLL_PATH,
                      GDL_GRI_PTLL_PATH, ALL_NOE_PTLL_PATH, ALL_LEVANTAMIENTO_PTLL_PATH]
    elif categoria == 'GENERAL':
        plantillas = [GENERAL_PTLL_PATH, GENERAL_MF_PTLL_PATH,
                      GENERAL_GRI_PTLL_PATH, ALL_NOE_PTLL_PATH, ALL_LEVANTAMIENTO_PTLL_PATH]
    elif categoria == 'BBVA':
        plantillas = [BBVA_PTLL_PATH, BBVA_MF_PTLL_PATH,
                      BBVA_GRI_PTLL_PATH, ALL_NOE_PTLL_PATH, ALL_LEVANTAMIENTO_PTLL_PATH]
    elif categoria == 'UVP':
        plantillas = [UVP_PTLL_PATH, UVP_MF_PTLL_PATH,
                      UVP_GRI_PTLL_PATH, ALL_NOE_PTLL_PATH, ALL_LEVANTAMIENTO_PTLL_PATH]
    elif categoria == 'DHL':
        plantillas = [DHL_PTLL_PATH, DHL_MF_PTLL_PATH, DHL_GRI_PTLL_PATH,
                      ALL_NOE_PTLL_PATH, ALL_LEVANTAMIENTO_PTLL_PATH]
    elif categoria == 'GASERA':
        plantillas = [GASERA_WORD_PTLL_PATH, GASERA_MF_WORD_PTLL_PATH,
                      GENERAL_GRI_PTLL_PATH, ALL_NOE_PTLL_PATH, ALL_LEVANTAMIENTO_PTLL_PATH]
    elif categoria == 'ESTAFETA':
        plantillas = [ESTAFETA_PTLL_PATH, ESTAFETA_MF_PTLL_PATH,
                      ESTAFETA_GRI_PTLL_PATH, ALL_NOE_PTLL_PATH, ALL_LEVANTAMIENTO_PTLL_PATH]
    elif categoria == 'FEDEX':
        plantillas = [FEDEX_PTLL_PATH, FEDEX_MF_PTLL_PATH,
                      FEDEX_GRI_PTLL_PATH, ALL_NOE_PTLL_PATH, ALL_LEVANTAMIENTO_PTLL_PATH]
    elif categoria == 'SMARTFIT':
        plantillas = [SMARTFIT_PTLL_PATH, SMARTFIT_MF_PTLL_PATH,
                      SMARTFIT_GRI_PTLL_PATH, ALL_NOE_PTLL_PATH, ALL_LEVANTAMIENTO_PTLL_PATH]
    elif categoria == 'AXA':
        plantillas = [AXA_PTLL_PATH, AXA_MF_PTLL_PATH,
                      AXA_GRI_PTLL_PATH, ALL_NOE_PTLL_PATH, ALL_LEVANTAMIENTO_PTLL_PATH]
    elif categoria == 'CONTINGENCIA':
        plantillas = [CONTINGENCIA_PTLL_PATH, GENERAL_MF_PTLL_PATH,
                      GENERAL_GRI_PTLL_PATH, ALL_NOE_PTLL_PATH, ALL_LEVANTAMIENTO_PTLL_PATH]
    else:
        print(f"No se encontraron plantillas para la categoría: {categoria}")
        return

    for idx, plantilla_path in enumerate(plantillas, start=1):
        # Cargar plantilla
        docx_tpl = DocxTemplate(plantilla_path)

        # Cargar las imágenes
        logo1 = cargar_imagen(docx_tpl, 'logo1', 'logo1.jpg',
                              150, 'width', datos_airtable)
        logo2 = cargar_imagen(docx_tpl, 'logo2', 'logo2.jpg',
                              15, 'height', datos_airtable)
        fachada = cargar_imagen(docx_tpl, 'fachada',
                                'fachada.jpg', 150, 'width', datos_airtable)
        mapa = cargar_imagen(docx_tpl, 'mapa', 'mapa.png',
                             144, 'width', datos_airtable)
        esc_emer = cargar_imagen(docx_tpl, 'esc_emer',
                                 'esc_emer.jpg', 50, 'height', datos_airtable)
        mueble1 = cargar_imagen(docx_tpl, 'mueble1',
                                'mueble (1).jpg', 50, 'height', datos_airtable)
        mueble2 = cargar_imagen(docx_tpl, 'mueble2',
                                'mueble (2).jpg', 50, 'height', datos_airtable)
        mueble3 = cargar_imagen(docx_tpl, 'mueble3',
                                'mueble (3).jpg', 50, 'height', datos_airtable)
        mueble4 = cargar_imagen(docx_tpl, 'mueble4',
                                'mueble (4).jpg', 50, 'height', datos_airtable)
        venteo = cargar_imagen(
            docx_tpl, 'venteo', 'venteo.jpg', 50, 'height', datos_airtable)
        manguera = cargar_imagen(docx_tpl, 'manguera',
                                 'manguera.jpg', 50, 'height', datos_airtable)
        electrico = cargar_imagen(
            docx_tpl, 'electrico', 'electrico.jpg', 50, 'height', datos_airtable)
        banio = cargar_imagen(docx_tpl, 'banio', 'banio.jpg',
                              50, 'height', datos_airtable)
        cisterna = cargar_imagen(docx_tpl, 'cisterna',
                                 'cisterna.jpg', 50, 'height', datos_airtable)
        sismo = cargar_imagen(docx_tpl, 'sismo', 'sismo.png',
                              155, 'width', datos_airtable)
        inundacion = cargar_imagen(
            docx_tpl, 'inundacion', 'inundacion.png', 155, 'width', datos_airtable)
        torm_elect = cargar_imagen(
            docx_tpl, 'torm_elect', 'torm_elect.png', 155, 'width', datos_airtable)
        incendio = cargar_imagen(docx_tpl, 'incendio',
                                 'incendio.png', 155, 'width', datos_airtable)
        influenza = cargar_imagen(
            docx_tpl, 'influenza', 'influenza.png', 155, 'width', datos_airtable)
        radiacion = cargar_imagen(
            docx_tpl, 'radiacion', 'radiacion.png', 155, 'width', datos_airtable)
        ext1 = cargar_imagen(docx_tpl, 'ext1', 'ext (1).jpg',
                             50, 'height', datos_airtable)
        ext2 = cargar_imagen(docx_tpl, 'ext2', 'ext (2).jpg',
                             50, 'height', datos_airtable)
        ext3 = cargar_imagen(docx_tpl, 'ext3', 'ext (3).jpg',
                             50, 'height', datos_airtable)
        ext4 = cargar_imagen(docx_tpl, 'ext4', 'ext (4).jpg',
                             50, 'height', datos_airtable)
        ext5 = cargar_imagen(docx_tpl, 'ext5', 'ext (5).jpg',
                             50, 'height', datos_airtable)
        ext6 = cargar_imagen(docx_tpl, 'ext6', 'ext (6).jpg',
                             50, 'height', datos_airtable)
        botiquin = cargar_imagen(docx_tpl, 'botiquin',
                                 'botiquin.jpg', 50, 'height', datos_airtable)
        botiquin1 = cargar_imagen(
            docx_tpl, 'botiquin1', 'botiquin1.jpg', 50, 'height', datos_airtable)
        ruta1 = cargar_imagen(
            docx_tpl, 'ruta1', 'ruta (1).jpg', 50, 'height', datos_airtable)
        ruta2 = cargar_imagen(
            docx_tpl, 'ruta2', 'ruta (2).jpg', 50, 'height', datos_airtable)
        ruta3 = cargar_imagen(
            docx_tpl, 'ruta3', 'ruta (3).jpg', 50, 'height', datos_airtable)
        ruta4 = cargar_imagen(
            docx_tpl, 'ruta4', 'ruta (4).jpg', 50, 'height', datos_airtable)
        salida = cargar_imagen(
            docx_tpl, 'salida', 'salida.jpg', 50, 'height', datos_airtable)
        alarma = cargar_imagen(
            docx_tpl, 'alarma', 'alarma.jpg', 50, 'height', datos_airtable)
        alarma1 = cargar_imagen(docx_tpl, 'alarma1',
                                'alarma1.jpg', 50, 'height', datos_airtable)
        prohib1 = cargar_imagen(docx_tpl, 'prohib1',
                                'prohib (1).jpg', 50, 'height', datos_airtable)
        prohib2 = cargar_imagen(docx_tpl, 'prohib2',
                                'prohib (2).jpg', 50, 'height', datos_airtable)
        prohib3 = cargar_imagen(docx_tpl, 'prohib3',
                                'prohib (3).jpg', 50, 'height', datos_airtable)
        prohib4 = cargar_imagen(docx_tpl, 'prohib4',
                                'prohib (4).jpg', 50, 'height', datos_airtable)
        layout = cargar_imagen(
            docx_tpl, 'layout', 'layout.png', 155, 'width', datos_airtable)
        cap1 = cargar_imagen(docx_tpl, 'cap1', 'cap (1).jpg',
                             60, 'height', datos_airtable)
        cap2 = cargar_imagen(docx_tpl, 'cap2', 'cap (2).jpg',
                             60, 'height', datos_airtable)
        cap3 = cargar_imagen(docx_tpl, 'cap3', 'cap (3).jpg',
                             60, 'height', datos_airtable)
        cap4 = cargar_imagen(docx_tpl, 'cap4', 'cap (4).jpg',
                             60, 'height', datos_airtable)
        cap5 = cargar_imagen(docx_tpl, 'cap5', 'cap (5).jpg',
                             60, 'height', datos_airtable)
        cap6 = cargar_imagen(docx_tpl, 'cap6', 'cap (6).jpg',
                             60, 'height', datos_airtable)
        cap7 = cargar_imagen(docx_tpl, 'cap7', 'cap (7).jpg',
                             60, 'height', datos_airtable)
        cap8 = cargar_imagen(docx_tpl, 'cap8', 'cap (8).jpg',
                             60, 'height', datos_airtable)
        cap9 = cargar_imagen(docx_tpl, 'cap9', 'cap (9).jpg',
                             60, 'height', datos_airtable)
        cap10 = cargar_imagen(
            docx_tpl, 'cap10', 'cap (10).jpg', 60, 'height', datos_airtable)
        cap11 = cargar_imagen(
            docx_tpl, 'cap11', 'cap (11).jpg', 60, 'height', datos_airtable)
        cap12 = cargar_imagen(
            docx_tpl, 'cap12', 'cap (12).jpg', 60, 'height', datos_airtable)
        cap13 = cargar_imagen(
            docx_tpl, 'cap13', 'cap (13).jpg', 60, 'height', datos_airtable)
        cap14 = cargar_imagen(
            docx_tpl, 'cap14', 'cap (14).jpg', 60, 'height', datos_airtable)
        cap15 = cargar_imagen(
            docx_tpl, 'cap15', 'cap (15).jpg', 60, 'height', datos_airtable)
        cap16 = cargar_imagen(
            docx_tpl, 'cap16', 'cap (16).jpg', 60, 'height', datos_airtable)
        cap17 = cargar_imagen(
            docx_tpl, 'cap17', 'cap (17).jpg', 60, 'height', datos_airtable)
        cap18 = cargar_imagen(
            docx_tpl, 'cap18', 'cap (18).jpg', 60, 'height', datos_airtable)
        cap19 = cargar_imagen(
            docx_tpl, 'cap19', 'cap (19).jpg', 60, 'height', datos_airtable)
        cap20 = cargar_imagen(
            docx_tpl, 'cap20', 'cap (20).jpg', 60, 'height', datos_airtable)
        cap21 = cargar_imagen(
            docx_tpl, 'cap21', 'cap (21).jpg', 60, 'height', datos_airtable)
        cap22 = cargar_imagen(
            docx_tpl, 'cap22', 'cap (22).jpg', 60, 'height', datos_airtable)
        cap23 = cargar_imagen(
            docx_tpl, 'cap23', 'cap (23).jpg', 60, 'height', datos_airtable)
        cap24 = cargar_imagen(
            docx_tpl, 'cap24', 'cap (24).jpg', 60, 'height', datos_airtable)
        sim1 = cargar_imagen(docx_tpl, 'sim1', 'sim (1).jpg',
                             60, 'height', datos_airtable)
        sim2 = cargar_imagen(docx_tpl, 'sim2', 'sim (2).jpg',
                             60, 'height', datos_airtable)
        sim3 = cargar_imagen(docx_tpl, 'sim3', 'sim (3).jpg',
                             60, 'height', datos_airtable)
        sim4 = cargar_imagen(docx_tpl, 'sim4', 'sim (4).jpg',
                             60, 'height', datos_airtable)
        sim5 = cargar_imagen(docx_tpl, 'sim5', 'sim (5).jpg',
                             60, 'height', datos_airtable)
        sim6 = cargar_imagen(docx_tpl, 'sim6', 'sim (6).jpg',
                             60, 'height', datos_airtable)
        sim7 = cargar_imagen(docx_tpl, 'sim7', 'sim (7).jpg',
                             60, 'height', datos_airtable)
        sim8 = cargar_imagen(docx_tpl, 'sim8', 'sim (8).jpg',
                             60, 'height', datos_airtable)
        sim9 = cargar_imagen(docx_tpl, 'sim9', 'sim (9).jpg',
                             60, 'height', datos_airtable)
        sim10 = cargar_imagen(
            docx_tpl, 'sim10', 'sim (10).jpg', 60, 'height', datos_airtable)
        sim11 = cargar_imagen(
            docx_tpl, 'sim11', 'sim (11).jpg', 60, 'height', datos_airtable)
        sim12 = cargar_imagen(
            docx_tpl, 'sim12', 'sim (12).jpg', 60, 'height', datos_airtable)
        techo = cargar_imagen(docx_tpl, 'techo', 'techo.jpg',
                              50, 'height', datos_airtable)
        techo1 = cargar_imagen(
            docx_tpl, 'techo1', 'techo1.jpg', 50, 'height', datos_airtable)
        pisos = cargar_imagen(docx_tpl, 'pisos', 'pisos.jpg',
                              50, 'height', datos_airtable)
        pisos1 = cargar_imagen(
            docx_tpl, 'pisos1', 'pisos1.jpg', 50, 'height', datos_airtable)
        puerta = cargar_imagen(
            docx_tpl, 'puerta', 'puerta.jpg', 50, 'height', datos_airtable)
        estantes = cargar_imagen(docx_tpl, 'estantes',
                                 'estantes.jpg', 50, 'height', datos_airtable)
        site = cargar_imagen(docx_tpl, 'site', 'site.jpg',
                             50, 'height', datos_airtable)
        dh = cargar_imagen(docx_tpl, 'dh', 'dh.jpg',
                           50, 'height', datos_airtable)
        dh1 = cargar_imagen(docx_tpl, 'dh1', 'dh1.jpg',
                            50, 'height', datos_airtable)
        ventanas = cargar_imagen(docx_tpl, 'ventanas',
                                 'ventanas.jpg', 50, 'height', datos_airtable)
        compresor = cargar_imagen(
            docx_tpl, 'compresor', 'compresor.jpg', 50, 'height', datos_airtable)
        quimicos = cargar_imagen(docx_tpl, 'quimicos',
                                 'quimicos.jpg', 50, 'height', datos_airtable)
        tanques_gaso = cargar_imagen(
            docx_tpl, 'tanques_gaso', 'tanques_gaso.jpg', 50, 'height', datos_airtable)
        tanques_gaso1 = cargar_imagen(
            docx_tpl, 'tanques_gaso1', 'tanques_gaso1.jpg', 50, 'height', datos_airtable)
        paro = cargar_imagen(docx_tpl, 'paro', 'paro.jpg',
                             50, 'height', datos_airtable)
        trampa_grasa = cargar_imagen(
            docx_tpl, 'trampa_grasa', 'trampa_grasa.jpg', 50, 'height', datos_airtable)
        planta = cargar_imagen(
            docx_tpl, 'planta', 'planta.jpg', 50, 'height', datos_airtable)
        deposito = cargar_imagen(docx_tpl, 'deposito',
                                 'deposito.jpg', 50, 'height', datos_airtable)
        mapa_satel = cargar_imagen(
            docx_tpl, 'mapa_satel', 'mapa_satel.png', 144, 'width', datos_airtable)
        plano = cargar_imagen(docx_tpl, 'plano', 'plano.jpg',
                              160, 'width', datos_airtable)
        inmueble1 = cargar_imagen(
            docx_tpl, 'inmueble1', 'inmueble (1).jpg', 50, 'height', datos_airtable)
        inmueble2 = cargar_imagen(
            docx_tpl, 'inmueble2', 'inmueble (2).jpg', 50, 'height', datos_airtable)
        inmueble3 = cargar_imagen(
            docx_tpl, 'inmueble3', 'inmueble (3).jpg', 50, 'height', datos_airtable)
        inmueble4 = cargar_imagen(
            docx_tpl, 'inmueble4', 'inmueble (4).jpg', 50, 'height', datos_airtable)
        banio1 = cargar_imagen(
            docx_tpl, 'banio1', 'banio1.jpg', 50, 'height', datos_airtable)
        electrico1 = cargar_imagen(
            docx_tpl, 'electrico1', 'electrico1.jpg', 50, 'height', datos_airtable)
        fachada1 = cargar_imagen(docx_tpl, 'fachada1',
                                 'fachada1.jpg', 50, 'height', datos_airtable)
        bateria = cargar_imagen(docx_tpl, 'bateria',
                                'bateria.jpg', 50, 'height', datos_airtable)
        acta1 = cargar_imagen(docx_tpl, 'acta1', 'acta1.png',
                              155, 'width', datos_airtable)
        acta2 = cargar_imagen(docx_tpl, 'acta2', 'acta2.png',
                              155, 'width', datos_airtable)
        acta = cargar_imagen(docx_tpl, 'acta', 'acta.png',
                             200, 'height', datos_airtable)
        crono_anual = cargar_imagen(
            docx_tpl, 'crono_anual', 'crono_anual.png', 155, 'width', datos_airtable)
        mantto1 = cargar_imagen(docx_tpl, 'mantto1',
                                'mantto1.png', 155, 'width', datos_airtable)
        mantto2 = cargar_imagen(docx_tpl, 'mantto2',
                                'mantto2.png', 155, 'width', datos_airtable)
        simulacro = cargar_imagen(
            docx_tpl, 'simulacro', 'simulacro.png', 155, 'width', datos_airtable)
        capacitacion = cargar_imagen(
            docx_tpl, 'capacitacion', 'capacitacion.png', 155, 'width', datos_airtable)
        inv_quim = cargar_imagen(docx_tpl, 'inv_quim',
                                 'inv_quim.png', 155, 'width', datos_airtable)
        inv_emer = cargar_imagen(docx_tpl, 'inv_emer',
                                 'inv_emer.png', 155, 'width', datos_airtable)
        bit_emer = cargar_imagen(docx_tpl, 'bit_emer',
                                 'bit_emer.png', 155, 'width', datos_airtable)
        insp_bot = cargar_imagen(docx_tpl, 'insp_bot',
                                 'insp_bot.png', 155, 'width', datos_airtable)
        insp_ext = cargar_imagen(docx_tpl, 'insp_ext',
                                 'insp_ext.png', 149, 'width', datos_airtable)
        insp_ext1 = cargar_imagen(
            docx_tpl, 'insp_ext1', 'insp_ext1.png', 149, 'width', datos_airtable)
        insp_dh = cargar_imagen(docx_tpl, 'insp_dh',
                                'insp_dh.png', 155, 'width', datos_airtable)
        insp_lamp = cargar_imagen(
            docx_tpl, 'insp_lamp', 'insp_lamp.png', 155, 'width', datos_airtable)
        insp_alarm = cargar_imagen(
            docx_tpl, 'insp_alarm', 'insp_alarm.png', 155, 'width', datos_airtable)
        ev_sim1 = cargar_imagen(docx_tpl, 'ev_sim1',
                                'ev_sim1.png', 155, 'width', datos_airtable)
        ev_sim2 = cargar_imagen(docx_tpl, 'ev_sim2',
                                'ev_sim2.png', 155, 'width', datos_airtable)
        visitas = cargar_imagen(docx_tpl, 'visitas',
                                'visitas.png', 155, 'width', datos_airtable)
        dir_emer = cargar_imagen(docx_tpl, 'dir_emer',
                                 'dir_emer.png', 155, 'width', datos_airtable)
        corresp1 = cargar_imagen(docx_tpl, 'corresp1',
                                 'corresp1.png', 155, 'width', datos_airtable)
        corresp2 = cargar_imagen(docx_tpl, 'corresp2',
                                 'corresp2.png', 155, 'width', datos_airtable)
        corresp3 = cargar_imagen(docx_tpl, 'corresp3',
                                 'corresp3.png', 155, 'width', datos_airtable)
        carta_respon = cargar_imagen(
            docx_tpl, 'carta_respon', 'carta_respon.png', 155, 'width', datos_airtable)
        registro1 = cargar_imagen(
            docx_tpl, 'registro1', 'registro1.png', 155, 'width', datos_airtable)
        registro2 = cargar_imagen(
            docx_tpl, 'registro2', 'registro2.png', 155, 'width', datos_airtable)
        ries_circ = cargar_imagen(
            docx_tpl, 'ries_circ', 'ries_circ.png', 155, 'width', datos_airtable)
        mapa_ext = cargar_imagen(docx_tpl, 'mapa_ext',
                                 'mapa_ext.png', 155, 'width', datos_airtable)
        rec_ext = cargar_imagen(docx_tpl, 'rec_ext',
                                'rec_ext.png', 155, 'width', datos_airtable)
        mayor_ries = cargar_imagen(
            docx_tpl, 'mayor_ries', 'mayor_ries.png', 155, 'width', datos_airtable)
        menor_ries = cargar_imagen(
            docx_tpl, 'menor_ries', 'menor_ries.png', 155, 'width', datos_airtable)
        zona_evac = cargar_imagen(
            docx_tpl, 'zona_evac', 'zona_evac.png', 155, 'width', datos_airtable)
        firma = cargar_imagen(docx_tpl, 'firma', 'firma.png',
                              14, 'height', datos_airtable)
        layout1 = cargar_imagen(docx_tpl, 'layout1',
                                'layout (1).png', 160, 'width', datos_airtable)
        layout2 = cargar_imagen(docx_tpl, 'layout2',
                                'layout (2).png', 160, 'width', datos_airtable)
        layout3 = cargar_imagen(docx_tpl, 'layout3',
                                'layout (3).png', 160, 'width', datos_airtable)
        layout4 = cargar_imagen(docx_tpl, 'layout4',
                                'layout (4).png', 160, 'width', datos_airtable)
        layout5 = cargar_imagen(docx_tpl, 'layout5',
                                'layout (5).png', 160, 'width', datos_airtable)
        layout6 = cargar_imagen(docx_tpl, 'layout6',
                                'layout (6).png', 160, 'width', datos_airtable)
        layout7 = cargar_imagen(docx_tpl, 'layout7',
                                'layout (7).png', 160, 'width', datos_airtable)
        layout8 = cargar_imagen(docx_tpl, 'layout8',
                                'layout (8).png', 160, 'width', datos_airtable)
        layout9 = cargar_imagen(docx_tpl, 'layout9',
                                'layout (9).png', 160, 'width', datos_airtable)
        layout10 = cargar_imagen(
            docx_tpl, 'layout10', 'layout (10).png', 160, 'width', datos_airtable)
        layout11 = cargar_imagen(
            docx_tpl, 'layout11', 'layout (11).png', 160, 'width', datos_airtable)
        layout12 = cargar_imagen(
            docx_tpl, 'layout12', 'layout (12).png', 160, 'width', datos_airtable)
        ev_sim3 = cargar_imagen(docx_tpl, 'ev_sim3',
                                'ev_sim (3).png', 155, 'width', datos_airtable)
        ev_sim4 = cargar_imagen(docx_tpl, 'ev_sim4',
                                'ev_sim (4).png', 155, 'width', datos_airtable)
        ev_sim5 = cargar_imagen(docx_tpl, 'ev_sim5',
                                'ev_sim (5).png', 155, 'width', datos_airtable)
        ev_sim6 = cargar_imagen(docx_tpl, 'ev_sim6',
                                'ev_sim (6).png', 155, 'width', datos_airtable)
        ev_sim7 = cargar_imagen(docx_tpl, 'ev_sim7',
                                'ev_sim (7).png', 155, 'width', datos_airtable)
        ev_sim8 = cargar_imagen(docx_tpl, 'ev_sim8',
                                'ev_sim (8).png', 155, 'width', datos_airtable)
        acta3 = cargar_imagen(docx_tpl, 'acta3', 'acta3.png',
                              155, 'width', datos_airtable)
        acta4 = cargar_imagen(docx_tpl, 'acta4', 'acta4.png',
                              155, 'width', datos_airtable)
        acta5 = cargar_imagen(docx_tpl, 'acta5', 'acta5.png',
                              155, 'width', datos_airtable)
        plan1 = cargar_imagen(
            docx_tpl, 'plan1', 'plan (1).jpg', 50, 'width', datos_airtable)
        plan2 = cargar_imagen(
            docx_tpl, 'plan2', 'plan (2).jpg', 50, 'width', datos_airtable)
        plan3 = cargar_imagen(
            docx_tpl, 'plan3', 'plan (3).jpg', 50, 'width', datos_airtable)
        lampara = cargar_imagen(docx_tpl, 'lampara',
                                'lampara.jpg', 50, 'heigth', datos_airtable)
        lampara1 = cargar_imagen(docx_tpl, 'lampara1',
                                 'lampara1.jpg', 50, 'heigth', datos_airtable)
        bombas = cargar_imagen(
            docx_tpl, 'bombas', 'bombas.jpg', 50, 'heigth', datos_airtable)
        atencion_clientes = cargar_imagen(
            docx_tpl, 'atencion_clientes', 'atencion_clientes.jpg', 50, 'heigth', datos_airtable)
        bardas = cargar_imagen(
            docx_tpl, 'bardas', 'bardas.jpg', 50, 'heigth', datos_airtable)
        sis_inc = cargar_imagen(docx_tpl, 'sis_inc',
                                'sis_inc.jpg', 50, 'heigth', datos_airtable)
        sis_inc1 = cargar_imagen(docx_tpl, 'sis_inc1',
                                 'sis_inc1.jpg', 50, 'heigth', datos_airtable)
        zona_sec = cargar_imagen(docx_tpl, 'zona_sec',
                                 'zona_sec.jpg', 50, 'heigth', datos_airtable)
        zona_sec1 = cargar_imagen(
            docx_tpl, 'zona_sec1', 'zona_sec1.jpg', 50, 'heigth', datos_airtable)
        punto_reun = cargar_imagen(
            docx_tpl, 'punto_reun', 'punto_reun.jpg', 50, 'heigth', datos_airtable)
        valvulas = cargar_imagen(docx_tpl, 'valvulas',
                                 'valvulas.jpg', 50, 'heigth', datos_airtable)
        bomberos = cargar_imagen(docx_tpl, 'bomberos',
                                 'bomberos.jpg', 50, 'heigth', datos_airtable)
        caldera = cargar_imagen(docx_tpl, 'caldera',
                                'caldera.jpg', 50, 'heigth', datos_airtable)
        brigadas = cargar_imagen(docx_tpl, 'brigadas',
                                 'brigadas.jpg', 50, 'heigth', datos_airtable)
        hidrante = cargar_imagen(docx_tpl, 'hidrante',
                                 'hidrante.jpg', 50, 'heigth', datos_airtable)

        try:
            # Renderizar documento - pasamos todos los datos directamente
            docx_tpl.render({**datos_airtable,
                            'logo1': logo1,
                             'logo2': logo2,
                             'fachada': fachada,
                             'mapa': mapa,
                             'esc_emer': esc_emer,
                             'mueble1': mueble1,
                             'mueble2': mueble2,
                             'venteo': venteo,
                             'manguera': manguera,
                             'electrico': electrico,
                             'banio': banio,
                             'cisterna': cisterna,
                             'sismo': sismo,
                             'inundacion': inundacion,
                             'torm_elect': torm_elect,
                             'incendio': incendio,
                             'influenza': influenza,
                             'radiacion': radiacion,
                             'ext1': ext1,
                             'ext2': ext2,
                             'ext3': ext3,
                             'ext4': ext4,
                             'ext5': ext5,
                             'ext6': ext6,
                             'botiquin': botiquin,
                             'ruta1': ruta1,
                             'ruta2': ruta2,
                             'ruta3': ruta3,
                             'salida': salida,
                             'alarma': alarma,
                             'alarma1': alarma1,
                             'prohib1': prohib1,
                             'prohib2': prohib2,
                             'prohib3': prohib3,
                             'prohib4': prohib4,
                             'layout': layout,
                             'cap1': cap1,
                             'cap2': cap2,
                             'cap3': cap3,
                             'cap4': cap4,
                             'cap5': cap5,
                             'cap6': cap6,
                             'cap7': cap7,
                             'cap8': cap8,
                             'cap9': cap9,
                             'cap10': cap10,
                             'cap11': cap11,
                             'cap12': cap12,
                             'sim1': sim1,
                             'sim2': sim2,
                             'sim3': sim3,
                             'sim4': sim4,
                             'sim5': sim5,
                             'sim6': sim6,
                             'techo': techo,
                             'pisos': pisos,
                             'puerta': puerta,
                             'estantes': estantes,
                             'site': site,
                             'dh': dh,
                             'ventanas': ventanas,
                             'compresor': compresor,
                             'quimicos': quimicos,
                             'tanques_gaso': tanques_gaso,
                             'tanques_gaso1': tanques_gaso1,
                             'paro': paro,
                             'trampa_grasa': trampa_grasa,
                             'planta': planta,
                             'deposito': deposito,
                             'mapa_satel': mapa_satel,
                             'plano': plano,
                             'inmueble1': inmueble1,
                             'inmueble2': inmueble2,
                             'banio1': banio1,
                             'electrico1': electrico1,
                             'fachada1': fachada1,
                             'bateria': bateria,
                             'acta1': acta1,
                             'acta2': acta2,
                             'acta': acta,
                             'crono_anual': crono_anual,
                             'mantto1': mantto1,
                             'mantto2': mantto2,
                             'simulacro': simulacro,
                             'capacitacion': capacitacion,
                             'inv_quim': inv_quim,
                             'inv_emer': inv_emer,
                             'bit_emer': bit_emer,
                             'insp_bot': insp_bot,
                             'insp_ext': insp_ext,
                             'insp_ext1': insp_ext1,
                             'insp_dh': insp_dh,
                             'insp_lamp': insp_lamp,
                             'insp_alarm': insp_alarm,
                             'ev_sim1': ev_sim1,
                             'ev_sim2': ev_sim2,
                             'visitas': visitas,
                             'dir_emer': dir_emer,
                             'corresp1': corresp1,
                             'corresp2': corresp2,
                             'corresp3': corresp3,
                             'carta_respon': carta_respon,
                             'registro1': registro1,
                             'registro2': registro2,
                             'ries_circ': ries_circ,
                             'mapa_ext': mapa_ext,
                             'rec_ext': rec_ext,
                             'mayor_ries': mayor_ries,
                             'menor_ries': menor_ries,
                             'zona_evac': zona_evac,
                             'firma': firma,
                             'layout1': layout1,
                             'layout2': layout2,
                             'layout3': layout3,
                             'layout4': layout4,
                             'layout5': layout5,
                             'layout6': layout6,
                             'layout7': layout7,
                             'layout8': layout8,
                             'layout9': layout9,
                             'layout10': layout10,
                             'layout11': layout11,
                             'layout12': layout12,
                             'ev_sim3': ev_sim3,
                             'ev_sim4': ev_sim4,
                             'ev_sim5': ev_sim5,
                             'ev_sim6': ev_sim6,
                             'ev_sim7': ev_sim7,
                             'ev_sim8': ev_sim8,
                             'acta3': acta3,
                             'acta4': acta4,
                             'acta5': acta5,
                             'plan1': plan1,
                             'plan2': plan2,
                             'plan3': plan3,
                             'lampara': lampara,
                             'bombas': bombas,
                             'atencion_clientes': atencion_clientes,
                             'bardas': bardas,
                             'sis_inc': sis_inc,
                             'zona_sec': zona_sec,
                             'punto_reun': punto_reun,
                             'valvulas': valvulas,
                             'lampara1': lampara1,
                             'dh1': dh1,
                             'sis_inc1': sis_inc1,
                             'zona_sec1': zona_sec1,
                             'mueble3': mueble3,
                             'mueble4': mueble4,
                             'inmueble3': inmueble3,
                             'inmueble4': inmueble4,
                             'botiquin1': botiquin1,
                             'ruta4': ruta4,
                             'cap13': cap13,
                             'cap14': cap14,
                             'cap15': cap15,
                             'cap16': cap16,
                             'cap17': cap17,
                             'cap18': cap18,
                             'cap19': cap19,
                             'cap20': cap20,
                             'cap21': cap21,
                             'cap22': cap22,
                             'cap23': cap23,
                             'cap24': cap24,
                             'sim7': sim7,
                             'sim8': sim8,
                             'sim9': sim9,
                             'sim10': sim10,
                             'sim11': sim11,
                             'sim12': sim12,
                             'techo1': techo1,
                             'pisos1': pisos1,
                             'bomberos': bomberos,
                             'caldera': caldera,
                             'brigadas': brigadas,
                             'hidrante': hidrante
                             })

            # Determinar nombre del archivo de salida
            nombre_comercial = datos_airtable.get(
                "nombre_comercial", "Documento")
            if idx == 1:
                nombre_pipc = f'1. PIPC {nombre_comercial}.docx'
            elif idx == 2:
                nombre_pipc = f'2. MEMORIA FOTOGRAFICA {nombre_comercial}.docx'
            elif idx == 3:
                nombre_pipc = f'3. RIESGO DE INCENDIO {nombre_comercial}.docx'
            elif idx == 4:
                nombre_pipc = f'4. CARTA NOE {nombre_comercial}.docx'
            elif idx == 5:
                nombre_pipc = f'5. LEVANTAMIENTO {nombre_comercial}.docx'
            else:
                nombre_pipc = f'{idx}. DOCUMENTO {nombre_comercial}.docx'

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
