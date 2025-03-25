import os
import shutil
import stat
import pandas as pd
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm

#################### CONFIGURACION DE USUARIO ####################

# Ruta base absoluta
BASE_DIR = "C:/Users/gutie/OneDrive/Documentos/GitHub/Proyecto_PIPC"

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
DHL_PTLL_PATH = f"{BASE_DIR}/Inputs/Templates/DHL.docx"
GASERA_WORD_PTLL_PATH = f"{BASE_DIR}/Inputs/Templates/Gasera.docx"
GASERA_MF_WORD_PTLL_PATH = f"{BASE_DIR}/Inputs/Templates/MF Gasera.docx"
ESTAFETA_PTLL_PATH = f"{BASE_DIR}/Inputs/Templates/Estafeta.docx"
ESTAFETA_MF_PTLL_PATH = f"{BASE_DIR}/Inputs/Templates/MF Estafeta.docx"
ESTAFETA_GRI_PTLL_PATH = f"{BASE_DIR}/Inputs/Templates/GRI Estafeta.docx"
ALL_LEVANTAMIENTO_PTLL_PATH = f"{BASE_DIR}/Inputs/Templates/Levantamiento.docx"

# Ruta imágenes
IMAGES_PATH = f"{BASE_DIR}/Inputs/Images"


#################### CONFIGURACION DE USUARIO ####################

# Eliminar y crear carpetas


def eliminar_crear_carpetas(path):
    # Función para manejar errores de permisos
    def on_rm_error(func, path, exc_info):
        import stat
        # Cambiar los permisos de la carpeta a escritura
        os.chmod(path, stat.S_IWRITE)
        func(path)  # Intentar eliminar la carpeta nuevamente

    # Verificar si la carpeta existe y eliminarla
    if os.path.exists(path):
        shutil.rmtree(path, onerror=on_rm_error)  # Llamar a rmtree con manejo de errores

    # Crear carpeta de salida
    os.mkdir(path)  # Crear la carpeta nueva

# Leer datos de Excel y pasarlo a formato dataframe 'df'


def leer_bd(path, worksheet):
    # Convertir Excel a dataframe
    excel_df = pd.read_excel(path, worksheet)

    return excel_df

# Rutina para crear ficheros Word para cada PIPC


def crear_word(df_pipc):
    # Iteramos sobre cada pipc
    for idx, r_val in df_pipc.iterrows():
        # Cargar plantilla
        if r_val['pipc'] == 'GASOLINERA':
            plantillas = [GASOLINERA_WORD_PTLL_PATH, GASOLINERA_MF_WORD_PTLL_PATH, GASOLINERA_GRI_WORD_PTLL_PATH, ALL_NOE_PTLL_PATH, ALL_LEVANTAMIENTO_PTLL_PATH]
        
        elif r_val['pipc'] == 'BANORTE':
            plantillas = [CIDUR_PTLL_PATH,
                          CIDUR_MF_PTLL_PATH, CIDUR_GRI_PTLL_PATH, ALL_NOE_PTLL_PATH, ALL_LEVANTAMIENTO_PTLL_PATH]
        
        elif r_val['pipc'] == 'GDL':
            plantillas = [GDL_PTLL_PATH, GDL_MF_PTLL_PATH, GDL_GRI_PTLL_PATH, ALL_NOE_PTLL_PATH, ALL_LEVANTAMIENTO_PTLL_PATH]
        
        elif r_val['pipc'] == 'GENERAL':
            plantillas = [GENERAL_PTLL_PATH,
                          GENERAL_MF_PTLL_PATH, GENERAL_GRI_PTLL_PATH, ALL_NOE_PTLL_PATH, ALL_LEVANTAMIENTO_PTLL_PATH]
        
        elif r_val['pipc'] == 'BBVA':
            plantillas = [BBVA_PTLL_PATH,
                          BBVA_MF_PTLL_PATH, BBVA_GRI_PTLL_PATH, ALL_NOE_PTLL_PATH, ALL_LEVANTAMIENTO_PTLL_PATH]
        
        elif r_val['pipc'] == 'UVP':
            plantillas = [UVP_PTLL_PATH,
                          GENERAL_MF_PTLL_PATH, GENERAL_GRI_PTLL_PATH, ALL_NOE_PTLL_PATH, ALL_LEVANTAMIENTO_PTLL_PATH]
            
        elif r_val['pipc'] == 'DHL':
            plantillas = [DHL_PTLL_PATH,
                          GENERAL_MF_PTLL_PATH, GENERAL_GRI_PTLL_PATH, ALL_NOE_PTLL_PATH, ALL_LEVANTAMIENTO_PTLL_PATH]
            
        elif r_val['pipc'] == 'GASERA':
            plantillas = [GASERA_WORD_PTLL_PATH,
                          GASERA_MF_WORD_PTLL_PATH, GENERAL_GRI_PTLL_PATH, ALL_NOE_PTLL_PATH, ALL_LEVANTAMIENTO_PTLL_PATH]
            
        elif r_val['pipc'] == 'ESTAFETA':
            plantillas = [ESTAFETA_PTLL_PATH,
                          ESTAFETA_MF_PTLL_PATH, ESTAFETA_GRI_PTLL_PATH, ALL_NOE_PTLL_PATH, ALL_LEVANTAMIENTO_PTLL_PATH]

        for idx, l_tpl in enumerate(plantillas, start=1):
            # Cargar plantilla
            docx_tpl = DocxTemplate(l_tpl)

            # Añadir imagen
            try:
                img_path_logo1 = os.path.join(IMAGES_PATH, r_val['logo1'])
                if os.path.exists(img_path_logo1):
                    logo1 = InlineImage(
                        docx_tpl, img_path_logo1, width=Mm(155))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["logo1"]}')
                    logo1 = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["logo1"]}: {e}')
                logo1 = ''

            try:
                img_path_logo2 = os.path.join(IMAGES_PATH, r_val["logo2"])
                if os.path.exists(img_path_logo2):
                    logo2 = InlineImage(
                        docx_tpl, img_path_logo2, height=Mm(15))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["logo2"]}')
                    logo2 = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["logo2"]}: {e}')
                logo2 = ''

            try:
                img_path_fachada = os.path.join(IMAGES_PATH, r_val["fachada"])
                if os.path.exists(img_path_fachada):
                    fachada = InlineImage(
                        docx_tpl, img_path_fachada, height=Mm(90))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["fachada"]}')
                    fachada = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["fachada"]}: {e}')
                fachada = ''

            try:
                img_path_mapa = os.path.join(IMAGES_PATH, r_val["mapa"])
                if os.path.exists(img_path_mapa):
                    mapa = InlineImage(docx_tpl, img_path_mapa, width=Mm(155))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["mapa"]}')
                    mapa = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["mapa"]}: {e}')
                mapa = ''

            try:
                img_path_esc_emer = os.path.join(
                    IMAGES_PATH, r_val["esc_emer"])
                if os.path.exists(img_path_esc_emer):
                    esc_emer = InlineImage(
                        docx_tpl, img_path_esc_emer, height=Mm(50))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["esc_emer"]}')
                    esc_emer = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["esc_emer"]}: {e}')
                esc_emer = ''

            try:
                img_path_mueble1 = os.path.join(IMAGES_PATH, r_val["mueble1"])
                if os.path.exists(img_path_mueble1):
                    mueble1 = InlineImage(
                        docx_tpl, img_path_mueble1, height=Mm(50))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["mueble1"]}')
                    mueble1 = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["mueble1"]}: {e}')
                mueble1 = ''

            try:
                img_path_mueble2 = os.path.join(IMAGES_PATH, r_val["mueble2"])
                if os.path.exists(img_path_mueble2):
                    mueble2 = InlineImage(
                        docx_tpl, img_path_mueble2, height=Mm(50))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["mueble2"]}')
                    mueble2 = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["mueble2"]}: {e}')
                mueble2 = ''

            try:
                img_path_venteo = os.path.join(IMAGES_PATH, r_val["venteo"])
                if os.path.exists(img_path_venteo):
                    venteo = InlineImage(
                        docx_tpl, img_path_venteo, height=Mm(50))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["venteo"]}')
                    venteo = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["venteo"]}: {e}')
                venteo = ''

            try:
                img_path_manguera = os.path.join(
                    IMAGES_PATH, r_val["manguera"])
                if os.path.exists(img_path_manguera):
                    manguera = InlineImage(
                        docx_tpl, img_path_manguera, height=Mm(50))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["manguera"]}')
                    manguera = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["manguera"]}: {e}')
                manguera = ''

            try:
                img_path_electrico = os.path.join(
                    IMAGES_PATH, r_val["electrico"])
                if os.path.exists(img_path_electrico):
                    electrico = InlineImage(
                        docx_tpl, img_path_electrico, height=Mm(50))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["electrico"]}')
                    electrico = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["electrico"]}: {e}')
                electrico = ''

            try:
                img_path_banio = os.path.join(IMAGES_PATH, r_val["banio"])
                if os.path.exists(img_path_banio):
                    banio = InlineImage(
                        docx_tpl, img_path_banio, height=Mm(50))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["banio"]}')
                    banio = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["banio"]}: {e}')
                banio = ''

            try:
                img_path_cisterna = os.path.join(
                    IMAGES_PATH, r_val["cisterna"])
                if os.path.exists(img_path_cisterna):
                    cisterna = InlineImage(
                        docx_tpl, img_path_cisterna, height=Mm(50))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["cisterna"]}')
                    cisterna = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["cisterna"]}: {e}')
                cisterna = ''

            try:
                img_path_sismo = os.path.join(IMAGES_PATH, r_val["sismo"])
                if os.path.exists(img_path_sismo):
                    sismo = InlineImage(
                        docx_tpl, img_path_sismo, width=Mm(155))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["sismo"]}')
                    sismo = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["sismo"]}: {e}')
                sismo = ''

            try:
                img_path_inundacion = os.path.join(
                    IMAGES_PATH, r_val["inundacion"])
                if os.path.exists(img_path_inundacion):
                    inundacion = InlineImage(
                        docx_tpl, img_path_inundacion, width=Mm(155))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["inundacion"]}')
                    inundacion = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["inundacion"]}: {e}')
                inundacion = ''

            try:
                img_path_torm_elect = os.path.join(
                    IMAGES_PATH, r_val["torm_elect"])
                if os.path.exists(img_path_torm_elect):
                    torm_elect = InlineImage(
                        docx_tpl, img_path_torm_elect, width=Mm(155))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["torm_elect"]}')
                    torm_elect = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["torm_elect"]}: {e}')
                torm_elect = ''

            try:
                img_path_incendio = os.path.join(
                    IMAGES_PATH, r_val["incendio"])
                if os.path.exists(img_path_incendio):
                    incendio = InlineImage(
                        docx_tpl, img_path_incendio, width=Mm(155))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["incendio"]}')
                    incendio = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["incendio"]}: {e}')
                incendio = ''

            try:
                img_path_influenza = os.path.join(
                    IMAGES_PATH, r_val["influenza"])
                if os.path.exists(img_path_influenza):
                    influenza = InlineImage(
                        docx_tpl, img_path_influenza, width=Mm(155))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["influenza"]}')
                    influenza = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["influenza"]}: {e}')
                influenza = ''

            try:
                img_path_radiacion = os.path.join(
                    IMAGES_PATH, r_val["radiacion"])
                if os.path.exists(img_path_radiacion):
                    radiacion = InlineImage(
                        docx_tpl, img_path_radiacion, width=Mm(155))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["radiacion"]}')
                    radiacion = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["radiacion"]}: {e}')
                radiacion = ''

            try:
                img_path_ext1 = os.path.join(IMAGES_PATH, r_val["ext1"])
                if os.path.exists(img_path_ext1):
                    ext1 = InlineImage(docx_tpl, img_path_ext1, height=Mm(50))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["ext1"]}')
                    ext1 = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["ext1"]}: {e}')
                ext1 = ''

            try:
                img_path_ext2 = os.path.join(IMAGES_PATH, r_val["ext2"])
                if os.path.exists(img_path_ext2):
                    ext2 = InlineImage(docx_tpl, img_path_ext2, height=Mm(50))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["ext2"]}')
                    ext2 = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["ext2"]}: {e}')
                ext2 = ''

            try:
                img_path_ext3 = os.path.join(IMAGES_PATH, r_val["ext3"])
                if os.path.exists(img_path_ext3):
                    ext3 = InlineImage(docx_tpl, img_path_ext3, height=Mm(50))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["ext3"]}')
                    ext3 = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["ext3"]}: {e}')
                ext3 = ''

            try:
                img_path_ext4 = os.path.join(IMAGES_PATH, r_val["ext4"])
                if os.path.exists(img_path_ext4):
                    ext4 = InlineImage(docx_tpl, img_path_ext4, height=Mm(50))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["ext4"]}')
                    ext4 = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["ext4"]}: {e}')
                ext4 = ''

            try:
                img_path_botiquin = os.path.join(
                    IMAGES_PATH, r_val["botiquin"])
                if os.path.exists(img_path_botiquin):
                    botiquin = InlineImage(
                        docx_tpl, img_path_botiquin, height=Mm(50))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["botiquin"]}')
                    botiquin = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["botiquin"]}: {e}')
                botiquin = ''

            try:
                img_path_ruta1 = os.path.join(IMAGES_PATH, r_val["ruta1"])
                if os.path.exists(img_path_ruta1):
                    ruta1 = InlineImage(
                        docx_tpl, img_path_ruta1, height=Mm(50))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["ruta1"]}')
                    ruta1 = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["ruta1"]}: {e}')
                ruta1 = ''

            try:
                img_path_ruta2 = os.path.join(IMAGES_PATH, r_val["ruta2"])
                if os.path.exists(img_path_ruta2):
                    ruta2 = InlineImage(
                        docx_tpl, img_path_ruta2, height=Mm(50))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["ruta2"]}')
                    ruta2 = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["ruta2"]}: {e}')
                ruta2 = ''

            try:
                img_path_ruta3 = os.path.join(IMAGES_PATH, r_val["ruta3"])
                if os.path.exists(img_path_ruta3):
                    ruta3 = InlineImage(
                        docx_tpl, img_path_ruta3, height=Mm(50))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["ruta3"]}')
                    ruta3 = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["ruta3"]}: {e}')
                ruta3 = ''

            try:
                img_path_salida = os.path.join(IMAGES_PATH, r_val["salida"])
                if os.path.exists(img_path_salida):
                    salida = InlineImage(
                        docx_tpl, img_path_salida, height=Mm(50))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["salida"]}')
                    salida = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["salida"]}: {e}')
                salida = ''

            try:
                img_path_alarma = os.path.join(IMAGES_PATH, r_val["alarma"])
                if os.path.exists(img_path_alarma):
                    alarma = InlineImage(
                        docx_tpl, img_path_alarma, height=Mm(50))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["alarma"]}')
                    alarma = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["alarma"]}: {e}')
                alarma = ''

            try:
                img_path_prohib1 = os.path.join(IMAGES_PATH, r_val["prohib1"])
                if os.path.exists(img_path_prohib1):
                    prohib1 = InlineImage(
                        docx_tpl, img_path_prohib1, height=Mm(50))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["prohib1"]}')
                    prohib1 = ''
            except Exception as e:
                print(
                    f'A  dvertencia: No se pudo cargar la imagen {r_val["prohib1"]}: {e}')
                prohib1 = ''

            try:
                img_path_prohib2 = os.path.join(IMAGES_PATH, r_val["prohib2"])
                if os.path.exists(img_path_prohib2):
                    prohib2 = InlineImage(
                        docx_tpl, img_path_prohib2, height=Mm(50))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["prohib2"]}')
                    prohib2 = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["prohib2"]}: {e}')
                prohib2 = ''

            try:
                img_path_prohib3 = os.path.join(IMAGES_PATH, r_val["prohib3"])
                if os.path.exists(img_path_prohib3):
                    prohib3 = InlineImage(
                        docx_tpl, img_path_prohib3, height=Mm(50))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["prohib3"]}')
                    prohib3 = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["prohib3"]}: {e}')
                prohib3 = ''

            try:
                img_path_prohib4 = os.path.join(IMAGES_PATH, r_val["prohib4"])
                if os.path.exists(img_path_prohib4):
                    prohib4 = InlineImage(
                        docx_tpl, img_path_prohib4, height=Mm(50))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["prohib4"]}')
                    prohib4 = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["prohib4"]}: {e}')
                prohib4 = ''

            try:
                img_path_layout = os.path.join(IMAGES_PATH, r_val["layout"])
                if os.path.exists(img_path_layout):
                    layout = InlineImage(
                        docx_tpl, img_path_layout, width=Mm(155))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["layout"]}')
                    layout = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["layout"]}: {e}')
                layout = ''

            try:
                img_path_cap1 = os.path.join(IMAGES_PATH, r_val["cap1"])
                if os.path.exists(img_path_cap1):
                    cap1 = InlineImage(docx_tpl, img_path_cap1, height=Mm(60))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["cap1"]}')
                    cap1 = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["cap1"]}: {e}')
                cap1 = ''

            try:
                img_path_cap2 = os.path.join(IMAGES_PATH, r_val["cap2"])
                if os.path.exists(img_path_cap2):
                    cap2 = InlineImage(docx_tpl, img_path_cap2, height=Mm(60))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["cap2"]}')
                    cap2 = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["cap2"]}: {e}')
                cap2 = ''

            try:
                img_path_cap3 = os.path.join(IMAGES_PATH, r_val["cap3"])
                if os.path.exists(img_path_cap3):
                    cap3 = InlineImage(docx_tpl, img_path_cap3, height=Mm(60))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["cap3"]}')
                    cap3 = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["cap3"]}: {e}')
                cap3 = ''

            try:
                img_path_cap4 = os.path.join(IMAGES_PATH, r_val["cap4"])
                if os.path.exists(img_path_cap4):
                    cap4 = InlineImage(docx_tpl, img_path_cap4, height=Mm(60))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["cap4"]}')
                    cap4 = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["cap4"]}: {e}')
                cap4 = ''

            try:
                img_path_cap5 = os.path.join(IMAGES_PATH, r_val["cap5"])
                if os.path.exists(img_path_cap5):
                    cap5 = InlineImage(docx_tpl, img_path_cap5, height=Mm(60))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["cap5"]}')
                    cap5 = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["cap5"]}: {e}')
                cap5 = ''

            try:
                img_path_cap6 = os.path.join(IMAGES_PATH, r_val["cap6"])
                if os.path.exists(img_path_cap6):
                    cap6 = InlineImage(docx_tpl, img_path_cap6, height=Mm(60))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["cap6"]}')
                    cap6 = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["cap6"]}: {e}')
                cap6 = ''

            try:
                img_path_cap7 = os.path.join(IMAGES_PATH, r_val["cap7"])
                if os.path.exists(img_path_cap7):
                    cap7 = InlineImage(docx_tpl, img_path_cap7, height=Mm(60))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["cap7"]}')
                    cap7 = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["cap7"]}: {e}')
                cap7 = ''

            try:
                img_path_cap8 = os.path.join(IMAGES_PATH, r_val["cap8"])
                if os.path.exists(img_path_cap8):
                    cap8 = InlineImage(docx_tpl, img_path_cap8, height=Mm(60))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["cap8"]}')
                    cap8 = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["cap8"]}: {e}')
                cap8 = ''

            try:
                img_path_cap9 = os.path.join(IMAGES_PATH, r_val["cap9"])
                if os.path.exists(img_path_cap9):
                    cap9 = InlineImage(docx_tpl, img_path_cap9, height=Mm(60))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["cap9"]}')
                    cap9 = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["cap9"]}: {e}')
                cap9 = ''

            try:
                img_path_cap10 = os.path.join(IMAGES_PATH, r_val["cap10"])
                if os.path.exists(img_path_cap10):
                    cap10 = InlineImage(
                        docx_tpl, img_path_cap10, height=Mm(60))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["cap10"]}')
                    cap10 = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["cap10"]}: {e}')
                cap10 = ''

            try:
                img_path_cap11 = os.path.join(IMAGES_PATH, r_val["cap11"])
                if os.path.exists(img_path_cap11):
                    cap11 = InlineImage(
                        docx_tpl, img_path_cap11, height=Mm(60))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["cap11"]}')
                    cap11 = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["cap11"]}: {e}')
                cap11 = ''

            try:
                img_path_cap12 = os.path.join(IMAGES_PATH, r_val["cap12"])
                if os.path.exists(img_path_cap12):
                    cap12 = InlineImage(
                        docx_tpl, img_path_cap12, height=Mm(60))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["cap12"]}')
                    cap12 = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["cap12"]}: {e}')
                cap12 = ''

            try:
                img_path_sim1 = os.path.join(IMAGES_PATH, r_val["sim1"])
                if os.path.exists(img_path_sim1):
                    sim1 = InlineImage(docx_tpl, img_path_sim1, height=Mm(60))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["sim1"]}')
                    sim1 = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["sim1"]}: {e}')
                sim1 = ''

            try:
                img_path_sim2 = os.path.join(IMAGES_PATH, r_val["sim2"])
                if os.path.exists(img_path_sim2):
                    sim2 = InlineImage(docx_tpl, img_path_sim2, height=Mm(60))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["sim2"]}')
                    sim2 = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["sim2"]}: {e}')
                sim2 = ''

            try:
                img_path_sim3 = os.path.join(IMAGES_PATH, r_val["sim3"])
                if os.path.exists(img_path_sim3):
                    sim3 = InlineImage(docx_tpl, img_path_sim3, height=Mm(60))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["sim3"]}')
                    sim3 = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["sim3"]}: {e}')
                sim3 = ''

            try:
                img_path_sim4 = os.path.join(IMAGES_PATH, r_val["sim4"])
                if os.path.exists(img_path_sim4):
                    sim4 = InlineImage(docx_tpl, img_path_sim4, height=Mm(60))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["sim4"]}')
                    sim4 = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["sim4"]}: {e}')
                sim4 = ''

            try:
                img_path_sim5 = os.path.join(IMAGES_PATH, r_val["sim5"])
                if os.path.exists(img_path_sim5):
                    sim5 = InlineImage(docx_tpl, img_path_sim5, height=Mm(60))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["sim5"]}')
                    sim5 = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["sim5"]}: {e}')
                sim5 = ''

            try:
                img_path_sim6 = os.path.join(IMAGES_PATH, r_val["sim6"])
                if os.path.exists(img_path_sim6):
                    sim6 = InlineImage(docx_tpl, img_path_sim6, height=Mm(60))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["sim6"]}')
                    sim6 = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["sim6"]}: {e}')
                sim6 = ''

            try:
                img_path_techo = os.path.join(IMAGES_PATH, r_val["techo"])
                if os.path.exists(img_path_techo):
                    techo = InlineImage(
                        docx_tpl, img_path_techo, height=Mm(50))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["techo"]}')
                    techo = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["techo"]}: {e}')
                techo = ''

            try:
                img_path_pisos = os.path.join(IMAGES_PATH, r_val["pisos"])
                if os.path.exists(img_path_pisos):
                    pisos = InlineImage(
                        docx_tpl, img_path_pisos, height=Mm(50))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["pisos"]}')
                    pisos = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["pisos"]}: {e}')
                pisos = ''

            try:
                img_path_puerta = os.path.join(IMAGES_PATH, r_val["puerta"])
                if os.path.exists(img_path_puerta):
                    puerta = InlineImage(
                        docx_tpl, img_path_puerta, height=Mm(50))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["puerta"]}')
                    puerta = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["puerta"]}: {e}')
                puerta = ''

            try:
                img_path_estantes = os.path.join(
                    IMAGES_PATH, r_val["estantes"])
                if os.path.exists(img_path_estantes):
                    estantes = InlineImage(
                        docx_tpl, img_path_estantes, height=Mm(50))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["estantes"]}')
                    estantes = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["estantes"]}: {e}')
                estantes = ''

            try:
                img_path_site = os.path.join(IMAGES_PATH, r_val["site"])
                if os.path.exists(img_path_site):
                    site = InlineImage(docx_tpl, img_path_site, height=Mm(50))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["site"]}')
                    site = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["site"]}: {e}')
                site = ''

            try:
                img_path_dh = os.path.join(IMAGES_PATH, r_val["dh"])
                if os.path.exists(img_path_dh):
                    dh = InlineImage(docx_tpl, img_path_dh, height=Mm(50))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["dh"]}')
                    dh = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["dh"]}: {e}')
                dh = ''

            try:
                img_path_ventanas = os.path.join(
                    IMAGES_PATH, r_val["ventanas"])
                if os.path.exists(img_path_ventanas):
                    ventanas = InlineImage(
                        docx_tpl, img_path_ventanas, height=Mm(50))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["ventanas"]}')
                    ventanas = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["ventanas"]}: {e}')
                ventanas = ''

            try:
                img_path_compresor = os.path.join(
                    IMAGES_PATH, r_val["compresor"])
                if os.path.exists(img_path_compresor):
                    compresor = InlineImage(
                        docx_tpl, img_path_compresor, height=Mm(50))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["compresor"]}')
                    compresor = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["compresor"]}: {e}')
                compresor = ''

            try:
                img_path_quimicos = os.path.join(
                    IMAGES_PATH, r_val["quimicos"])
                if os.path.exists(img_path_quimicos):
                    quimicos = InlineImage(
                        docx_tpl, img_path_quimicos, height=Mm(50))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["quimicos"]}')
                    quimicos = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["quimicos"]}: {e}')
                quimicos = ''

            try:
                img_path_tanques_gaso = os.path.join(
                    IMAGES_PATH, r_val["tanques_gaso"])
                if os.path.exists(img_path_tanques_gaso):
                    tanques_gaso = InlineImage(
                        docx_tpl, img_path_tanques_gaso, height=Mm(50))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["tanques_gaso"]}')
                    tanques_gaso = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["tanques_gaso"]}: {e}')
                tanques_gaso = ''

            try:
                img_path_paro = os.path.join(IMAGES_PATH, r_val["paro"])
                if os.path.exists(img_path_paro):
                    paro = InlineImage(docx_tpl, img_path_paro, height=Mm(50))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["paro"]}')
                    paro = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["paro"]}: {e}')
                paro = ''

            try:
                img_path_trampa_grasa = os.path.join(
                    IMAGES_PATH, r_val["trampa_grasa"])
                if os.path.exists(img_path_trampa_grasa):
                    trampa_grasa = InlineImage(
                        docx_tpl, img_path_trampa_grasa, height=Mm(50))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["trampa_grasa"]}')
                    trampa_grasa = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["trampa_grasa"]}: {e}')
                trampa_grasa = ''

            try:
                img_path_planta = os.path.join(IMAGES_PATH, r_val["planta"])
                if os.path.exists(img_path_planta):
                    planta = InlineImage(
                        docx_tpl, img_path_planta, height=Mm(50))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["planta"]}')
                    planta = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["planta"]}: {e}')
                planta = ''

            try:
                img_path_deposito = os.path.join(
                    IMAGES_PATH, r_val["deposito"])
                if os.path.exists(img_path_deposito):
                    deposito = InlineImage(
                        docx_tpl, img_path_deposito, height=Mm(50))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["deposito"]}')
                    deposito = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["deposito"]}: {e}')
                deposito = ''

            try:
                img_path_mapa_satel = os.path.join(
                    IMAGES_PATH, r_val["mapa_satel"])
                if os.path.exists(img_path_mapa_satel):
                    mapa_satel = InlineImage(
                        docx_tpl, img_path_mapa_satel, width=Mm(155))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["mapa_satel"]}')
                    mapa_satel = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["mapa_satel"]}: {e}')
                mapa_satel = ''

            try:
                img_path_plano = os.path.join(IMAGES_PATH, r_val["plano"])
                if os.path.exists(img_path_plano):
                    plano = InlineImage(
                        docx_tpl, img_path_plano, width=Mm(155))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["plano"]}')
                    plano = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["plano"]}: {e}')
                plano = ''

            try:
                img_path_inmueble1 = os.path.join(
                    IMAGES_PATH, r_val["inmueble1"])
                if os.path.exists(img_path_inmueble1):
                    inmueble1 = InlineImage(
                        docx_tpl, img_path_inmueble1, height=Mm(50))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["inmueble1"]}')
                    inmueble1 = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["inmueble1"]}: {e}')
                inmueble1 = ''

            try:
                img_path_inmueble2 = os.path.join(
                    IMAGES_PATH, r_val["inmueble2"])
                if os.path.exists(img_path_inmueble2):
                    inmueble2 = InlineImage(
                        docx_tpl, img_path_inmueble2, height=Mm(50))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["inmueble2"]}')
                    inmueble2 = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["inmueble2"]}: {e}')
                inmueble2 = ''

            try:
                img_path_banio1 = os.path.join(IMAGES_PATH, r_val["banio1"])
                if os.path.exists(img_path_banio1):
                    banio1 = InlineImage(
                        docx_tpl, img_path_banio1, height=Mm(50))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["banio1"]}')
                    banio1 = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["banio1"]}: {e}')
                banio1 = ''

            try:
                img_path_electrico1 = os.path.join(
                    IMAGES_PATH, r_val["electrico1"])
                if os.path.exists(img_path_electrico1):
                    electrico1 = InlineImage(
                        docx_tpl, img_path_electrico1, height=Mm(50))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["electrico1"]}')
                    electrico1 = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["electrico1"]}: {e}')
                electrico1 = ''

            try:
                img_path_fachada1 = os.path.join(
                    IMAGES_PATH, r_val["fachada1"])
                if os.path.exists(img_path_fachada1):
                    fachada1 = InlineImage(
                        docx_tpl, img_path_fachada1, height=Mm(50))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["fachada1"]}')
                    fachada1 = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["fachada1"]}: {e}')
                fachada1 = ''

            try:
                img_path_bateria = os.path.join(IMAGES_PATH, r_val["bateria"])
                if os.path.exists(img_path_bateria):
                    bateria = InlineImage(
                        docx_tpl, img_path_bateria, height=Mm(50))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["bateria"]}')
                    bateria = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["bateria"]}: {e}')
                bateria = ''

            try:
                img_path_acta1 = os.path.join(IMAGES_PATH, r_val["acta1"])
                if os.path.exists(img_path_acta1):
                    acta1 = InlineImage(
                        docx_tpl, img_path_acta1, width=Mm(155))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["acta1"]}')
                    acta1 = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["acta1"]}: {e}')
                acta1 = ''

            try:
                img_path_acta2 = os.path.join(IMAGES_PATH, r_val["acta2"])
                if os.path.exists(img_path_acta2):
                    acta2 = InlineImage(
                        docx_tpl, img_path_acta2, width=Mm(155))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["acta2"]}')
                    acta2 = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["acta2"]}: {e}')
                acta2 = ''

            try:
                img_path_crono_anual = os.path.join(
                    IMAGES_PATH, r_val["crono_anual"])
                if os.path.exists(img_path_crono_anual):
                    crono_anual = InlineImage(
                        docx_tpl, img_path_crono_anual, width=Mm(155))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["crono_anual"]}')
                    crono_anual = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["crono_anual"]}: {e}')
                crono_anual = ''

            try:
                img_path_mantto1 = os.path.join(IMAGES_PATH, r_val["mantto1"])
                if os.path.exists(img_path_mantto1):
                    mantto1 = InlineImage(
                        docx_tpl, img_path_mantto1, width=Mm(155))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["mantto1"]}')
                    mantto1 = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["mantto1"]}: {e}')
                mantto1 = ''

            try:
                img_path_mantto2 = os.path.join(IMAGES_PATH, r_val["mantto2"])
                if os.path.exists(img_path_mantto2):
                    mantto2 = InlineImage(
                        docx_tpl, img_path_mantto2, width=Mm(155))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["mantto2"]}')
                    mantto2 = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["mantto2"]}: {e}')
                mantto2 = ''

            try:
                img_path_simulacro = os.path.join(
                    IMAGES_PATH, r_val["simulacro"])
                if os.path.exists(img_path_simulacro):
                    simulacro = InlineImage(
                        docx_tpl, img_path_simulacro, width=Mm(155))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["simulacro"]}')
                    simulacro = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["simulacro"]}: {e}')
                simulacro = ''

            try:
                img_path_capacitacion = os.path.join(
                    IMAGES_PATH, r_val["capacitacion"])
                if os.path.exists(img_path_capacitacion):
                    capacitacion = InlineImage(
                        docx_tpl, img_path_capacitacion, width=Mm(155))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["capacitacion"]}')
                    capacitacion = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["capacitacion"]}: {e}')
                capacitacion = ''

            try:
                img_path_inv_quim = os.path.join(
                    IMAGES_PATH, r_val["inv_quim"])
                if os.path.exists(img_path_inv_quim):
                    inv_quim = InlineImage(
                        docx_tpl, img_path_inv_quim, width=Mm(155))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["inv_quim"]}')
                    inv_quim = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["inv_quim"]}: {e}')
                inv_quim = ''

            try:
                img_path_inv_emer = os.path.join(
                    IMAGES_PATH, r_val["inv_emer"])
                if os.path.exists(img_path_inv_emer):
                    inv_emer = InlineImage(
                        docx_tpl, img_path_inv_emer, width=Mm(155))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["inv_emer"]}')
                    inv_emer = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["inv_emer"]}: {e}')
                inv_emer = ''

            try:
                img_path_bit_emer = os.path.join(
                    IMAGES_PATH, r_val["bit_emer"])
                if os.path.exists(img_path_bit_emer):
                    bit_emer = InlineImage(
                        docx_tpl, img_path_bit_emer, width=Mm(155))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["bit_emer"]}')
                    bit_emer = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["bit_emer"]}: {e}')
                bit_emer = ''

            try:
                img_path_insp_bot = os.path.join(
                    IMAGES_PATH, r_val["insp_bot"])
                if os.path.exists(img_path_insp_bot):
                    insp_bot = InlineImage(
                        docx_tpl, img_path_insp_bot, width=Mm(155))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["insp_bot"]}')
                    insp_bot = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["insp_bot"]}: {e}')
                insp_bot = ''

            try:
                img_path_insp_ext = os.path.join(
                    IMAGES_PATH, r_val["insp_ext"])
                if os.path.exists(img_path_insp_ext):
                    insp_ext = InlineImage(
                        docx_tpl, img_path_insp_ext, width=Mm(149))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["insp_ext"]}')
                    insp_ext = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["insp_ext"]}: {e}')
                insp_ext = ''

            try:
                img_path_insp_dh = os.path.join(IMAGES_PATH, r_val["insp_dh"])
                if os.path.exists(img_path_insp_dh):
                    insp_dh = InlineImage(
                        docx_tpl, img_path_insp_dh, width=Mm(155))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["insp_dh"]}')
                    insp_dh = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["insp_dh"]}: {e}')
                insp_dh = ''

            try:
                img_path_insp_lamp = os.path.join(
                    IMAGES_PATH, r_val["insp_lamp"])
                if os.path.exists(img_path_insp_lamp):
                    insp_lamp = InlineImage(
                        docx_tpl, img_path_insp_lamp, width=Mm(155))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["insp_lamp"]}')
                    insp_lamp = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["insp_lamp"]}: {e}')
                insp_lamp = ''

            try:
                img_path_insp_alarm = os.path.join(
                    IMAGES_PATH, r_val["insp_alarm"])
                if os.path.exists(img_path_insp_alarm):
                    insp_alarm = InlineImage(
                        docx_tpl, img_path_insp_alarm, width=Mm(155))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["insp_alarm"]}')
                    insp_alarm = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["insp_alarm"]}: {e}')
                insp_alarm = ''

            try:
                img_path_ev_sim1 = os.path.join(IMAGES_PATH, r_val["ev_sim1"])
                if os.path.exists(img_path_ev_sim1):
                    ev_sim1 = InlineImage(
                        docx_tpl, img_path_ev_sim1, width=Mm(155))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["ev_sim1"]}')
                    ev_sim1 = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["ev_sim1"]}: {e}')
                ev_sim1 = ''

            try:
                img_path_ev_sim2 = os.path.join(IMAGES_PATH, r_val["ev_sim2"])
                if os.path.exists(img_path_ev_sim2):
                    ev_sim2 = InlineImage(
                        docx_tpl, img_path_ev_sim2, width=Mm(155))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["ev_sim2"]}')
                    ev_sim2 = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["ev_sim2"]}: {e}')
                ev_sim2 = ''

            try:
                img_path_visitas = os.path.join(IMAGES_PATH, r_val["visitas"])
                if os.path.exists(img_path_visitas):
                    visitas = InlineImage(
                        docx_tpl, img_path_visitas, width=Mm(155))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["visitas"]}')
                    visitas = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["visitas"]}: {e}')
                visitas = ''

            try:
                img_path_dir_emer = os.path.join(
                    IMAGES_PATH, r_val["dir_emer"])
                if os.path.exists(img_path_dir_emer):
                    dir_emer = InlineImage(
                        docx_tpl, img_path_dir_emer, width=Mm(155))
                else:
                    print(
                        f'Advertencia: No se encontró la imagen {r_val["dir_emer"]}')
                    dir_emer = ''
            except Exception as e:
                print(
                    f'Advertencia: No se pudo cargar la imagen {r_val["dir_emer"]}: {e}')
                dir_emer = ''

            try:	
                img_path_corresp1 = os.path.join(IMAGES_PATH, r_val["corresp1"])	
                if os.path.exists(img_path_corresp1):	
                    corresp1 = InlineImage(docx_tpl, img_path_corresp1, width=Mm(155))	
                else:	
                    print(f'Advertencia: No se encontró la imagen {r_val["corresp1"]}')	
                    corresp1 = ''	
            except Exception as e:	
                print(f'Advertencia: No se pudo cargar la imagen {r_val["corresp1"]}: {e}')	
                corresp1 = ''

            try:	
                img_path_corresp2 = os.path.join(IMAGES_PATH, r_val["corresp2"])	
                if os.path.exists(img_path_corresp2):	
                    corresp2 = InlineImage(docx_tpl, img_path_corresp2, width=Mm(155))	
                else:	
                    print(f'Advertencia: No se encontró la imagen {r_val["corresp2"]}')	
                    corresp2 = ''	
            except Exception as e:	
                print(f'Advertencia: No se pudo cargar la imagen {r_val["corresp2"]}: {e}')	
                corresp2 = ''
            
            try:	
                img_path_corresp3 = os.path.join(IMAGES_PATH, r_val["corresp3"])	
                if os.path.exists(img_path_corresp3):	
                    corresp3 = InlineImage(docx_tpl, img_path_corresp3, width=Mm(155))	
                else:	
                    print(f'Advertencia: No se encontró la imagen {r_val["corresp3"]}')	
                    corresp3 = ''	
            except Exception as e:	
                print(f'Advertencia: No se pudo cargar la imagen {r_val["corresp3"]}: {e}')	
                corresp3 = ''
            
            try:	
                img_path_carta_respon = os.path.join(IMAGES_PATH, r_val["carta_respon"])	
                if os.path.exists(img_path_carta_respon):	
                    carta_respon = InlineImage(docx_tpl, img_path_carta_respon, width=Mm(155))	
                else:	
                    print(f'Advertencia: No se encontró la imagen {r_val["carta_respon"]}')	
                    carta_respon = ''	
            except Exception as e:	
                print(f'Advertencia: No se pudo cargar la imagen {r_val["carta_respon"]}: {e}')	
                carta_respon = ''
            
            try:	
                img_path_registro1 = os.path.join(IMAGES_PATH, r_val["registro1"])	
                if os.path.exists(img_path_registro1):	
                    registro1 = InlineImage(docx_tpl, img_path_registro1, width=Mm(155))	
                else:	
                    print(f'Advertencia: No se encontró la imagen {r_val["registro1"]}')	
                    registro1 = ''	
            except Exception as e:	
                print(f'Advertencia: No se pudo cargar la imagen {r_val["registro1"]}: {e}')	
                registro1 = ''
            
            try:	
                img_path_registro2 = os.path.join(IMAGES_PATH, r_val["registro2"])	
                if os.path.exists(img_path_registro2):	
                    registro2 = InlineImage(docx_tpl, img_path_registro2, width=Mm(155))	
                else:	
                    print(f'Advertencia: No se encontró la imagen {r_val["registro2"]}')	
                    registro2 = ''	
            except Exception as e:	
                print(f'Advertencia: No se pudo cargar la imagen {r_val["registro2"]}: {e}')	
                registro2 = ''
            
            try:	
                img_path_ries_circ = os.path.join(IMAGES_PATH, r_val["ries_circ"])	
                if os.path.exists(img_path_ries_circ):	
                    ries_circ = InlineImage(docx_tpl, img_path_ries_circ, width=Mm(155))	
                else:	
                    print(f'Advertencia: No se encontró la imagen {r_val["ries_circ"]}')	
                    ries_circ = ''	
            except Exception as e:	
                print(f'Advertencia: No se pudo cargar la imagen {r_val["ries_circ"]}: {e}')	
                ries_circ = ''
            
            try:	
                img_path_mapa_ext = os.path.join(IMAGES_PATH, r_val["mapa_ext"])
                if os.path.exists(img_path_mapa_ext):
                    mapa_ext = InlineImage(docx_tpl, img_path_mapa_ext, width=Mm(155))	
                else:	
                    print(f'Advertencia: No se encontró la imagen {r_val["mapa_ext"]}')	
                    mapa_ext = ''	
            except Exception as e:	
                print(f'Advertencia: No se pudo cargar la imagen {r_val["mapa_ext"]}: {e}')	
                mapa_ext = ''
            
            try:	
                img_path_rec_ext = os.path.join(IMAGES_PATH, r_val["rec_ext"])	
                if os.path.exists(img_path_rec_ext):	
                    rec_ext = InlineImage(docx_tpl, img_path_rec_ext, width=Mm(155))	
                else:	
                    print(f'Advertencia: No se encontró la imagen {r_val["rec_ext"]}')	
                    rec_ext = ''	
            except Exception as e:	
                print(f'Advertencia: No se pudo cargar la imagen {r_val["rec_ext"]}: {e}')	
                rec_ext = ''
            
            try:	
                img_path_mayor_ries = os.path.join(IMAGES_PATH, r_val["mayor_ries"])	
                if os.path.exists(img_path_mayor_ries):	
                    mayor_ries = InlineImage(docx_tpl, img_path_mayor_ries, width=Mm(155))	
                else:	
                    print(f'Advertencia: No se encontró la imagen {r_val["mayor_ries"]}')	
                    mayor_ries = ''	
            except Exception as e:	
                print(f'Advertencia: No se pudo cargar la imagen {r_val["mayor_ries"]}: {e}')	
                mayor_ries = ''
            
            try:	
                img_path_menor_ries = os.path.join(IMAGES_PATH, r_val["menor_ries"])	
                if os.path.exists(img_path_menor_ries):	
                    menor_ries = InlineImage(docx_tpl, img_path_menor_ries, width=Mm(155))	
                else:	
                    print(f'Advertencia: No se encontró la imagen {r_val["menor_ries"]}')	
                    menor_ries = ''	
            except Exception as e:	
                print(f'Advertencia: No se pudo cargar la imagen {r_val["menor_ries"]}: {e}')	
                menor_ries = ''
            
            try:	
                img_path_zona_evac = os.path.join(IMAGES_PATH, r_val["zona_evac"])	
                if os.path.exists(img_path_zona_evac):	
                    zona_evac = InlineImage(docx_tpl, img_path_zona_evac, width=Mm(155))	
                else:	
                    print(f'Advertencia: No se encontró la imagen {r_val["zona_evac"]}')	
                    zona_evac = ''	
            except Exception as e:	
                print(f'Advertencia: No se pudo cargar la imagen {r_val["zona_evac"]}: {e}')	
                zona_evac = ''

            try:	
                img_path_firma = os.path.join(IMAGES_PATH, r_val["firma"])	
                if os.path.exists(img_path_firma):	
                    firma = InlineImage(docx_tpl, img_path_firma, width=Mm(70))	
                else:	
                    print(f'Advertencia: No se encontró la imagen {r_val["firma"]}')	
                    firma = ''	
            except Exception as e:	
                print(f'Advertencia: No se pudo cargar la imagen {r_val["firma"]}: {e}')	
                firma = ''

            try:	
                img_path_layout1 = os.path.join(IMAGES_PATH, r_val["layout1"])	
                if os.path.exists(img_path_layout1):	
                    layout1 = InlineImage(docx_tpl, img_path_layout1, height=Mm(100))	
                else:	
                    print(f'Advertencia: No se encontró la imagen {r_val["layout1"]}')	
                    layout1 = ''	
            except Exception as e:	
                print(f'Advertencia: No se pudo cargar la imagen {r_val["layout1"]}: {e}')	
                layout1 = ''
            
            try:	
                img_path_layout2 = os.path.join(IMAGES_PATH, r_val["layout2"])	
                if os.path.exists(img_path_layout2):	
                    layout2 = InlineImage(docx_tpl, img_path_layout2, height=Mm(100))	
                else:	
                    print(f'Advertencia: No se encontró la imagen {r_val["layout2"]}')	
                    layout2 = ''	
            except Exception as e:	
                print(f'Advertencia: No se pudo cargar la imagen {r_val["layout2"]}: {e}')	
                layout2 = ''
            
            try:	
                img_path_layout3 = os.path.join(IMAGES_PATH, r_val["layout3"])
                if os.path.exists(img_path_layout3):	
                    layout3 = InlineImage(docx_tpl, img_path_layout3, height=Mm(100))	
                else:	
                    print(f'Advertencia: No se encontró la imagen {r_val["layout3"]}')	
                    layout3 = ''	
            except Exception as e:	
                print(f'Advertencia: No se pudo cargar la imagen {r_val["layout3"]}: {e}')	
                layout3 = ''
            
            try:	
                img_path_layout4 = os.path.join(IMAGES_PATH, r_val["layout4"])	
                if os.path.exists(img_path_layout4):
                    layout4 = InlineImage(docx_tpl, img_path_layout4, height=Mm(100))	
                else:	
                    print(f'Advertencia: No se encontró la imagen {r_val["layout4"]}')	
                    layout4 = ''	
            except Exception as e:	
                print(f'Advertencia: No se pudo cargar la imagen {r_val["layout4"]}: {e}')	
                layout4 = ''
            
            try:	
                img_path_layout5 = os.path.join(IMAGES_PATH, r_val["layout5"])	
                if os.path.exists(img_path_layout5):	
                    layout5 = InlineImage(docx_tpl, img_path_layout5, height=Mm(100))	
                else:	
                    print(f'Advertencia: No se encontró la imagen {r_val["layout5"]}')	
                    layout5 = ''	
            except Exception as e:	
                print(f'Advertencia: No se pudo cargar la imagen {r_val["layout5"]}: {e}')
                layout5 = ''
            
            try:	
                img_path_layout6 = os.path.join(IMAGES_PATH, r_val["layout6"])	
                if os.path.exists(img_path_layout6):	
                    layout6 = InlineImage(docx_tpl, img_path_layout6, height=Mm(100))	
                else:	
                    print(f'Advertencia: No se encontró la imagen {r_val["layout6"]}')	
                    layout6 = ''	
            except Exception as e:	
                print(f'Advertencia: No se pudo cargar la imagen {r_val["layout6"]}: {e}')
                layout6 = ''
            
            try:	
                img_path_layout7 = os.path.join(IMAGES_PATH, r_val["layout7"])	
                if os.path.exists(img_path_layout7):	
                    layout7 = InlineImage(docx_tpl, img_path_layout7, height=Mm(100))
                else:	
                    print(f'Advertencia: No se encontró la imagen {r_val["layout7"]}')
                    layout7 = ''	
            except Exception as e:
                    print(f'Advertencia: No se pudo cargar la imagen {r_val["layout7"]}: {e}')	
                    layout7 = ''
            
            try:	
                img_path_layout8 = os.path.join(IMAGES_PATH, r_val["layout8"])
                if os.path.exists(img_path_layout8):	
                    layout8 = InlineImage(docx_tpl, img_path_layout8, height=Mm(100))	
                else:	
                    print(f'Advertencia: No se encontró la imagen {r_val["layout8"]}')	
                    layout8 = ''	
            except Exception as e:	
                print(f'Advertencia: No se pudo cargar la imagen {r_val["layout8"]}: {e}')
                layout8 = ''
            
            try:	
                img_path_layout9 = os.path.join(IMAGES_PATH, r_val["layout9"])	
                if os.path.exists(img_path_layout9):	
                    layout9 = InlineImage(docx_tpl, img_path_layout9, height=Mm(100))	
                else:	
                    print(f'Advertencia: No se encontró la imagen {r_val["layout9"]}')
                    layout9 = ''	
            except Exception as e:	
                print(f'Advertencia: No se pudo cargar la imagen {r_val["layout9"]}: {e}')	
                layout9 = ''
            
            try:	
                img_path_layout10 = os.path.join(IMAGES_PATH, r_val["layout10"])
                if os.path.exists(img_path_layout10):	
                    layout10 = InlineImage(docx_tpl, img_path_layout10, height=Mm(100))
                else:	
                    print(f'Advertencia: No se encontró la imagen {r_val["layout10"]}')	
                    layout10 = ''	
            except Exception as e:	
                print(f'Advertencia: No se pudo cargar la imagen {r_val["layout10"]}: {e}')	
                layout10 = ''
            
            try:	
                img_path_layout11 = os.path.join(IMAGES_PATH, r_val["layout11"])	
                if os.path.exists(img_path_layout11):	
                    layout11 = InlineImage(docx_tpl, img_path_layout11, height=Mm(100))	
                else:	
                    print(f'Advertencia: No se encontró la imagen {r_val["layout11"]}')	
                    layout11 = ''	
            except Exception as e:	
                print(f'Advertencia: No se pudo cargar la imagen {r_val["layout11"]}: {e}')	
                layout11 = ''
            
            try:	
                img_path_layout12 = os.path.join(IMAGES_PATH, r_val["layout12"])	
                if os.path.exists(img_path_layout12):
                        layout12 = InlineImage(docx_tpl, img_path_layout12, height=Mm(100))
                else:	
                    print(f'Advertencia: No se encontró la imagen {r_val["layout12"]}')	
                    layout12 = ''	
            except Exception as e:	
                print(f'Advertencia: No se pudo cargar la imagen {r_val["layout12"]}: {e}')	
                layout12 = ''

            try:	
                img_path_ev_sim3 = os.path.join(IMAGES_PATH, r_val["ev_sim3"])	
                if os.path.exists(img_path_ev_sim3):	
                    ev_sim3 = InlineImage(docx_tpl, img_path_ev_sim3, width=Mm(155))	
                else:	
                    print(f'Advertencia: No se encontró la imagen {r_val["ev_sim3"]}')	
                    ev_sim3 = ''	
            except Exception as e:	
                print(f'Advertencia: No se pudo cargar la imagen {r_val["ev_sim3"]}: {e}')	
                ev_sim3 = ''
            
            try:	
                img_path_ev_sim4 = os.path.join(IMAGES_PATH, r_val["ev_sim4"])	
                if os.path.exists(img_path_ev_sim4):	
                    ev_sim4 = InlineImage(docx_tpl, img_path_ev_sim4, width=Mm(155))	
                else:	
                    print(f'Advertencia: No se encontró la imagen {r_val["ev_sim4"]}')	
                    ev_sim4 = ''	
            except Exception as e:	
                print(f'Advertencia: No se pudo cargar la imagen {r_val["ev_sim4"]}: {e}')	
                ev_sim4 = ''
            
            try:	
                img_path_ev_sim5 = os.path.join(IMAGES_PATH, r_val["ev_sim5"])	
                if os.path.exists(img_path_ev_sim5):	
                    ev_sim5 = InlineImage(docx_tpl, img_path_ev_sim5, width=Mm(155))	
                else:	
                    print(f'Advertencia: No se encontró la imagen {r_val["ev_sim5"]}')	
                    ev_sim5 = ''	
            except Exception as e:
                print(f'Advertencia: No se pudo cargar la imagen {r_val["ev_sim5"]}: {e}')	
                ev_sim5 = ''
            
            try:	
                img_path_ev_sim6 = os.path.join(IMAGES_PATH, r_val["ev_sim6"])	
                if os.path.exists(img_path_ev_sim6):	
                    ev_sim6 = InlineImage(docx_tpl, img_path_ev_sim6, width=Mm(155))
                else:	
                    print(f'Advertencia: No se encontró la imagen {r_val["ev_sim6"]}')	
                    ev_sim6 = ''	
            except Exception as e:	
                print(f'Advertencia: No se pudo cargar la imagen {r_val["ev_sim6"]}: {e}')	
                ev_sim6 = ''
            
            try:	
                img_path_ev_sim7 = os.path.join(IMAGES_PATH, r_val["ev_sim7"])	
                if os.path.exists(img_path_ev_sim7):	
                    ev_sim7 = InlineImage(docx_tpl, img_path_ev_sim7, width=Mm(155))	
                else:	
                    print(f'Advertencia: No se encontró la imagen {r_val["ev_sim7"]}')	
                    ev_sim7 = ''	
            except Exception as e:	
                print(f'Advertencia: No se pudo cargar la imagen {r_val["ev_sim7"]}: {e}')	
                ev_sim7 = ''
            
            try:	
                img_path_ev_sim8 = os.path.join(IMAGES_PATH, r_val["ev_sim8"])	
                if os.path.exists(img_path_ev_sim8):	
                    ev_sim8 = InlineImage(docx_tpl, img_path_ev_sim8, width=Mm(155))	
                else:	
                    print(f'Advertencia: No se encontró la imagen {r_val["ev_sim8"]}')
                    ev_sim8 = ''	
            except Exception as e:
                print(f'Advertencia: No se pudo cargar la imagen {r_val["ev_sim8"]}: {e}')	
                ev_sim8 = ''

            try:	
                img_path_acta3 = os.path.join(IMAGES_PATH, r_val["acta3"])	
                if os.path.exists(img_path_acta3):	
                    acta3 = InlineImage(docx_tpl, img_path_acta3, width=Mm(155))	
                else:	
                    print(f'Advertencia: No se encontró la imagen {r_val["acta3"]}')	
                    acta3 = ''	
            except Exception as e:	
                print(f'Advertencia: No se pudo cargar la imagen {r_val["acta3"]}: {e}')	
                acta3 = ''
            
            try:	
                img_path_acta4 = os.path.join(IMAGES_PATH, r_val["acta4"])	
                if os.path.exists(img_path_acta4):	
                    acta4 = InlineImage(docx_tpl, img_path_acta4, width=Mm(155))	
                else:	
                    print(f'Advertencia: No se encontró la imagen {r_val["acta4"]}')	
                    acta4 = ''	
            except Exception as e:	
                print(f'Advertencia: No se pudo cargar la imagen {r_val["acta4"]}: {e}')	
                acta4 = ''


            # Crear contexto
            context = {
                # 'id': r_val['id'],
                'pipc': r_val['pipc'],
                'razon_social': r_val['razon_social'],
                'nombre_comercial': r_val['nombre_comercial'],
                'rfc': r_val['rfc'],
                'codigo_gasolinera': r_val['codigo_gasolinera'],
                'giro_comercial': r_val['giro_comercial'],
                'descripcion_actividades': r_val['descripcion_actividades'],
                'calle': r_val['calle'],
                'no_exterior': r_val['no_exterior'],
                'no_interior': r_val['no_interior'],
                'colonia_barrio': r_val['colonia_barrio'],
                'municipio': r_val['municipio'],
                'estado': r_val['estado'],
                'codigo_postal': r_val['codigo_postal'],
                'telefono': r_val['telefono'],
                'email': r_val['email'],
                'antiguedad_inmueble': r_val['antiguedad_inmueble'],
                'inicio_operaciones': r_val['inicio_operaciones'],
                'registro_perito' : r_val['registro_perito'],
                'no_registro' : r_val['no_registro'],
                'terreno_m2': r_val['terreno_m2'],
                'construccion_m2': r_val['construccion_m2'],
                'edificios': r_val['edificios'],
                'niveles': r_val['niveles'],
                'accesos': r_val['accesos'],
                'salidas_emergencia': r_val['salidas_emergencia'],
                'escaleras': r_val['escaleras'],
                'escaleras_emergencia': r_val['escaleras_emergencia'],
                'estacionamiento': r_val['estacionamiento'],
                'representante_legal': r_val['representante_legal'],
                'responsable_pipc': r_val['responsable_pipc'],
                'trabajadores': r_val['trabajadores'],
                'poblacion_discapacidad': r_val['poblacion_discapacidad'],
                'hombres': r_val['hombres'],
                'mujeres': r_val['mujeres'],
                'hombre_dicapacidad': r_val['hombre_dicapacidad'],
                'mujeres_discapacidad': r_val['mujeres_discapacidad'],
                'turnos': r_val['turnos'],
                'visitantes': r_val['visitantes'],
                'proveedores': r_val['proveedores'],
                'dias': r_val['dias'],
                'horario': r_val['horario'],
                'senial_informacion': r_val['senial_informacion'],
                'botiquin_emer': r_val['botiquin_emer'],
                'ubicación_bot': r_val['ubicación_bot'],
                'extintor': r_val['extintor'],
                'ubicación_ext': r_val['ubicación_ext'],
                'ext_pqs': r_val['ext_pqs'],
                'ext_co2': r_val['ext_co2'],
                'paros_emergencia': r_val['paros_emergencia'],
                'ubicación_pe': r_val['ubicación_pe'],
                'venteo_sist': r_val['venteo_sist'],
                'ubicación_vent': r_val['ubicación_vent'],
                'planta_emer': r_val['planta_emer'],
                'ubicación_plant': r_val['ubicación_plant'],
                'met_alarma': r_val['met_alarma'],
                'tipo_alarma': r_val['tipo_alarma'],
                'silbato': r_val['silbato'],
                'estrobo': r_val['estrobo'],
                'ubicacion_alarma': r_val['ubicacion_alarma'],
                'dh_int' : r_val['dh_int'],
                'dh_ext' : r_val['dh_ext'],
                'detectores_humo': r_val['detectores_humo'],
                'ubicación_dh': r_val['ubicación_dh'],
                'ruta_evac': r_val['ruta_evac'],
                'escaleras': r_val['escaleras'],
                'salida_emerg': r_val['salida_emerg'],
                'zona_menor_rg': r_val['zona_menor_rg'],
                'punto_reunion': r_val['punto_reunion'],
                'sismo_incendio': r_val['sismo_incendio'],
                'riesgo_electrico': r_val['riesgo_electrico'],
                'senial_prohibicion': r_val['senial_prohibicion'],
                'no_fumar': r_val['no_fumar'],
                'area_restrig': r_val['area_restrig'],
                'apague_motor': r_val['apague_motor'],
                'no_celular': r_val['no_celular'],
                'no_gorra_lentes': r_val['no_gorra_lentes'],
                'uso_epp': r_val['uso_epp'],
                'detectores_ext' : r_val['detectores_ext'],
                'detectores_mov': r_val['detectores_mov'],
                'ubicación_mov': r_val['ubicación_mov'],
                'site_emer': r_val['site_emer'],
                'ubicacion_site': r_val['ubicacion_site'],
                'hidrantes': r_val['hidrantes'],
                'aspersores': r_val['aspersores'],
                'bomberos': r_val['bomberos'],
                'ubicacion_bombero': r_val['ubicacion_bombero'],
                'detector_gas': r_val['detector_gas'],
                'ubicación_gas': r_val['ubicación_gas'],
                'equipo_brigada': r_val['equipo_brigada'],
                'ubicacion_brigada': r_val['ubicacion_brigada'],
                'lampara': r_val['lampara'],
                'ubicacion_lampara': r_val['ubicacion_lampara'],
                'baterias': r_val['baterias'],
                'ubiacion_baterias': r_val['ubiacion_baterias'],
                'tambo_arena': r_val['tambo_arena'],
                'ubicacion_tambo': r_val['ubicacion_tambo'],
                'tanques': r_val['tanques'],
                'tanque_1': r_val['tanque_1'],
                'tanque_2': r_val['tanque_2'],
                'tanque_3': r_val['tanque_3'],
                'riesgo1' : r_val['riesgo1'],
                'riesgo2' : r_val['riesgo2'],
                'riesgo3' : r_val['riesgo3'],
                'riesgo4' : r_val['riesgo4'],
                'medidas_ries1' : r_val['medidas_ries1'],
                'medidas_ries2' : r_val['medidas_ries2'],
                'medidas_ries3' : r_val['medidas_ries3'],
                'medidas_ries4' : r_val['medidas_ries4'],
                'dia': r_val['dia'],
                'mes': r_val['mes'],
                'anio': r_val['anio'],
                'coordenadas': r_val['coordenadas'],
                'norte': r_val['norte'],
                'sur': r_val['sur'],
                'este': r_val['este'],
                'oeste': r_val['oeste'],
                'ref_llegar': r_val['ref_llegar'],
                'ley': r_val['ley'],
                'reglamento': r_val['reglamento'],
                'm_dir': r_val['m_dir'],
                'coord_suplente': r_val['coord_suplente'],
                'coord_sup_puesto' : r_val['coord_sup_puesto'],
                'm_jef_caj': r_val['m_jef_caj'],
                'evacuacion': r_val['evacuacion'],
                'evac_puesto': r_val['evac_puesto'],
                'm_evac': r_val['m_evac'],
                'evac_suplente': r_val['evac_suplente'],
                'supl_evac_pue': r_val['supl_evac_pue'],
                'm_supl_evac': r_val['m_supl_evac'],
                'incendios': r_val['incendios'],
                'incen_puesto': r_val['incen_puesto'],
                'm_inc': r_val['m_inc'],
                'inc_suplente': r_val['inc_suplente'],
                'supl_inc_puesto': r_val['supl_inc_puesto'],
                'm_supl_inc': r_val['m_supl_inc'],
                'primeros_auxilios': r_val['primeros_auxilios'],
                'prim_aux_puesto': r_val['prim_aux_puesto'],
                'm_prim_aux': r_val['m_prim_aux'],
                'aux_suplente': r_val['aux_suplente'],
                'supl_prim_aux_puesto': r_val['supl_prim_aux_puesto'],
                'm_supl_paux': r_val['m_supl_paux'],
                'busqueda': r_val['busqueda'],
                'busq_puesto': r_val['busq_puesto'],
                'm_busq': r_val['m_busq'],
                'busq_suplente': r_val['busq_suplente'],
                'supl_busq_puesto': r_val['supl_busq_puesto'],
                'm_supl_busq': r_val['m_supl_busq'],
                'descrip_gi': r_val['descrip_gi'],
                'gas_inflamable': r_val['gas_inflamable'],
                'valor_gi': r_val['valor_gi'],
                'descrp_li': r_val['descrp_li'],
                'liquido_inflamable': r_val['liquido_inflamable'],
                'valor_li': r_val['valor_li'],
                'descrip_lc': r_val['descrip_lc'],
                'liquido_combustible': r_val['liquido_combustible'],
                'valor_lc': r_val['valor_lc'],
                'descrip_sc': r_val['descrip_sc'],
                'solido_combusible': r_val['solido_combusible'],
                'valor_sc': r_val['valor_sc'],
                'valor_gri' : r_val['valor_gri'],
                'tipo_riesgo': r_val['tipo_riesgo'],
                'nombre1': r_val['nombre1'],
                'puesto1': r_val['puesto1'],
                'm1': r_val['m1'],
                'nombre2': r_val['nombre2'],
                'puesto2': r_val['puesto2'],
                'm2': r_val['m2'],
                'barretas' : r_val['barretas'],
                'barretas_ubicacion' : r_val['barretas_ubicacion'],
                'banderines' : r_val['banderines'],
                'banderines_ubicacion' : r_val['banderines_ubicacion'],
                'casco' : r_val['casco'],
                'casco_ubicacion' : r_val['casco_ubicacion'],
                'guantes' : r_val['guantes'],
                'guantes_ubicacion' : r_val['guantes_ubicacion'],
                'linterna' : r_val['linterna'],
                'linterna_ubicacion' : r_val['linterna_ubicacion'],
                'pala' : r_val['pala'],
                'pala_ubicacion' : r_val['pala_ubicacion'],
                'pico' : r_val['pico'],
                'pico_ubicacion' : r_val['pico_ubicacion'],
                'camilla' : r_val['camilla'],
                'camilla_ubicacion' : r_val['camilla_ubicacion'],
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
                'botiquin': botiquin,
                'ruta1': ruta1,
                'ruta2': ruta2,
                'ruta3': ruta3,
                'salida': salida,
                'alarma': alarma,
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
                'acta4': acta4
            }

            try:

                # Renderizamos usando el contexto creado
                docx_tpl.render(context)

                # Guardar documento
                if idx == 1:
                    nombre_pipc = '1. PIPC ' + \
                        r_val['nombre_comercial'] + '.docx'
                elif idx == 2:
                    nombre_pipc = '2. MEMORIA FOTOGRAFICA ' + \
                        r_val['nombre_comercial'] + '.docx'
                elif idx == 3:
                    nombre_pipc = '3. RIESGO DE INCENDIO ' + \
                        r_val['nombre_comercial'] + '.docx'
                    
                elif idx == 4:
                    nombre_pipc = '4. CARTAS GOWER ' + \
                        r_val['nombre_comercial'] + '.docx'
                    
                elif idx == 5:
                    nombre_pipc = '5. CARTAS NOE ' + \
                        r_val['nombre_comercial'] + '.docx'
                    
                elif idx == 6:
                    nombre_pipc = '6. LEVANTAMIENTO ' + \
                        r_val['nombre_comercial'] + '.docx'

                # Guardar el documento con un nombre único
                docx_tpl.save(OUTPUT_PATH + '\\' + nombre_pipc)

            except Exception as e:
                print(f'Error al guardar el documento: {str(e)}')

# Rutina principal


def main():
    # Eliminar y volver a crear carpeta 'Outputs'
    eliminar_crear_carpetas(OUTPUT_PATH)

    # Leer datos de DB Excel
    df_bd = leer_bd(EXCEL_PATH, 'DATOS')

    # Crear ficheros Word
    crear_word(df_bd)


if __name__ == '__main__':
    main()
