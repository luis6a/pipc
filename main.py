import os
import shutil
import pandas as pd
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm

#################### CONFIGURACION DE USUARIO ####################

# Ruta de salida
OUTPUT_PATH = '.\Outputs'

# Ruta fichero Excel
EXCEL_PATH = '.\Inputs\BD.xlsx'

# Ruta plantillas ficheros Word
GASOLINERA_WORD_PTLL_PATH = '.\Inputs\Templates\Gasolinera.docx'
GAS_MF_WORD_PTLL_PATH = '.\Inputs\Templates\MF Gasolinera.docx'
GRI_GAS_WORD_PTLL_PATH = '.\Inputs\Templates\GRI Gasolinera.docx'
CIDUR_PTLL_PATH = '.\Inputs\Templates\Cidur.docx'
CIDUR_MF_PTLL_PATH = '.\Inputs\Templates\MF Cidur.docx'
CIDUR_GRI_PTLL_PATH = '.\Inputs\Templates\GRI Cidur.docx'

# Ruta imágenes
IMAGES_PATH = '.\Inputs\Images'

#################### CONFIGURACION DE USUARIO ####################

# Eliminar y crear carpetas


def eliminar_crear_carpetas(path):
    # Verficiar si la carpeta existe y eliminarla
    if (os.path.exists(path)):
        shutil.rmtree(path)

    # Crear carpeta de salida
    os.mkdir(OUTPUT_PATH)

# Leer datos de Excel y pasarlo a formato dataframe 'df'


def leer_bd(path, worksheet):
    # Convertir Excel a dataframe
    excel_df = pd.read_excel(path, worksheet)

    return excel_df

# Rutina para crear ficheros Word para cada PIPC


def crear_word(df_pipc):
    # Interamos sobre cada pipc
    for r_idx, r_val in df_pipc.iterrows():
        # Cargar plantilla
        if (r_val['pipc'] == 'GASOLINERA'):
            l_tpl = GASOLINERA_WORD_PTLL_PATH  # and GAS_MF_WORD_PTLL_PATH
        elif (r_val['pipc'] == 'BANCO'):
            l_tpl = CIDUR_PTLL_PATH  # and CIDUR_MF_PTLL_PATH

        # Procesamos plantilla
        docx_tpl = DocxTemplate(l_tpl)

        # Añadir imagen
        img_path_logo1 = os.path.join(IMAGES_PATH + '\\' + r_val['logo1'])
        logo1 = InlineImage(docx_tpl, img_path_logo1, width=Mm(145))

        img_path_logo2 = os.path.join(IMAGES_PATH + '\\' + r_val['logo2'])
        logo2 = InlineImage(docx_tpl, img_path_logo2, height=Mm(25))

        img_path_fachada = os.path.join(IMAGES_PATH + '\\' + r_val['fachada'])
        fachada = InlineImage(docx_tpl, img_path_fachada, height=Mm(105))

        img_path_mapa = os.path.join(IMAGES_PATH + '\\' + r_val['mapa'])
        mapa = InlineImage(docx_tpl, img_path_mapa, height=Mm(95))

        img_path_esc_emer = os.path.join(
            IMAGES_PATH + '\\' + r_val['esc_emer'])
        esc_emer = InlineImage(docx_tpl, img_path_esc_emer, height=Mm(60))

        img_path_mueble1 = os.path.join(IMAGES_PATH + '\\' + r_val['mueble1'])
        mueble1 = InlineImage(docx_tpl, img_path_mueble1, height=Mm(50))

        img_path_mueble2 = os.path.join(IMAGES_PATH + '\\' + r_val['mueble2'])
        mueble2 = InlineImage(docx_tpl, img_path_mueble2, height=Mm(50))

        img_path_venteo = os.path.join(IMAGES_PATH + '\\' + r_val['venteo'])
        venteo = InlineImage(docx_tpl, img_path_venteo, height=Mm(60))

        img_path_manguera = os.path.join(
            IMAGES_PATH + '\\' + r_val['manguera'])
        manguera = InlineImage(docx_tpl, img_path_manguera, height=Mm(50))

        img_path_electrico = os.path.join(
            IMAGES_PATH + '\\' + r_val['electrico'])
        electrico = InlineImage(docx_tpl, img_path_electrico, height=Mm(60))

        img_path_banio = os.path.join(IMAGES_PATH + '\\' + r_val['banio'])
        banio = InlineImage(docx_tpl, img_path_banio, height=Mm(50))

        img_path_cisterna = os.path.join(
            IMAGES_PATH + '\\' + r_val['cisterna'])
        cisterna = InlineImage(docx_tpl, img_path_cisterna, height=Mm(60))

        img_path_sismo = os.path.join(IMAGES_PATH + '\\' + r_val['sismo'])
        sismo = InlineImage(docx_tpl, img_path_sismo, height=Mm(70))

        img_path_inundacion = os.path.join(
            IMAGES_PATH + '\\' + r_val['inundacion'])
        inundacion = InlineImage(docx_tpl, img_path_inundacion, height=Mm(70))

        img_path_torm_elect = os.path.join(
            IMAGES_PATH + '\\' + r_val['torm_elect'])
        torm_elect = InlineImage(docx_tpl, img_path_torm_elect, height=Mm(70))

        img_path_incendio = os.path.join(
            IMAGES_PATH + '\\' + r_val['incendio'])
        incendio = InlineImage(docx_tpl, img_path_incendio, height=Mm(70))

        img_path_influenza = os.path.join(
            IMAGES_PATH + '\\' + r_val['influenza'])
        influenza = InlineImage(docx_tpl, img_path_influenza, height=Mm(70))

        img_path_radiacion = os.path.join(
            IMAGES_PATH + '\\' + r_val['radiacion'])
        radiacion = InlineImage(docx_tpl, img_path_radiacion, height=Mm(70))

        img_path_ext1 = os.path.join(IMAGES_PATH + '\\' + r_val['ext1'])
        ext1 = InlineImage(docx_tpl, img_path_ext1, height=Mm(68))

        img_path_ext2 = os.path.join(IMAGES_PATH + '\\' + r_val['ext2'])
        ext2 = InlineImage(docx_tpl, img_path_ext2, height=Mm(68))

        img_path_ext3 = os.path.join(IMAGES_PATH + '\\' + r_val['ext3'])
        ext3 = InlineImage(docx_tpl, img_path_ext3, height=Mm(68))

        img_path_ext4 = os.path.join(IMAGES_PATH + '\\' + r_val['ext4'])
        ext4 = InlineImage(docx_tpl, img_path_ext4, height=Mm(62))

        img_path_botiquin = os.path.join(
            IMAGES_PATH + '\\' + r_val['botiquin'])
        botiquin = InlineImage(docx_tpl, img_path_botiquin, height=Mm(60))

        img_path_ruta1 = os.path.join(IMAGES_PATH + '\\' + r_val['ruta1'])
        ruta1 = InlineImage(docx_tpl, img_path_ruta1, height=Mm(53))

        img_path_ruta2 = os.path.join(IMAGES_PATH + '\\' + r_val['ruta2'])
        ruta2 = InlineImage(docx_tpl, img_path_ruta2, height=Mm(53))

        img_path_ruta3 = os.path.join(IMAGES_PATH + '\\' + r_val['ruta3'])
        ruta3 = InlineImage(docx_tpl, img_path_ruta3, height=Mm(53))

        img_path_salida = os.path.join(IMAGES_PATH + '\\' + r_val['salida'])
        salida = InlineImage(docx_tpl, img_path_salida, height=Mm(53))

        img_path_alarma = os.path.join(IMAGES_PATH + '\\' + r_val['alarma'])
        alarma = InlineImage(docx_tpl, img_path_alarma, height=Mm(60))

        img_path_prohib1 = os.path.join(IMAGES_PATH + '\\' + r_val['prohib1'])
        prohib1 = InlineImage(docx_tpl, img_path_prohib1, height=Mm(53))

        img_path_prohib2 = os.path.join(IMAGES_PATH + '\\' + r_val['prohib2'])
        prohib2 = InlineImage(docx_tpl, img_path_prohib2, height=Mm(53))

        img_path_prohib3 = os.path.join(IMAGES_PATH + '\\' + r_val['prohib3'])
        prohib3 = InlineImage(docx_tpl, img_path_prohib3, height=Mm(53))

        img_path_prohib4 = os.path.join(IMAGES_PATH + '\\' + r_val['prohib4'])
        prohib4 = InlineImage(docx_tpl, img_path_prohib4, height=Mm(53))

        img_path_layout = os.path.join(IMAGES_PATH + '\\' + r_val['layout'])
        layout = InlineImage(docx_tpl, img_path_layout, height=Mm(158))

        img_path_cap1 = os.path.join(IMAGES_PATH + '\\' + r_val['cap1'])
        cap1 = InlineImage(docx_tpl, img_path_cap1, height=Mm(60))

        img_path_cap2 = os.path.join(IMAGES_PATH + '\\' + r_val['cap2'])
        cap2 = InlineImage(docx_tpl, img_path_cap2, height=Mm(60))

        img_path_cap3 = os.path.join(IMAGES_PATH + '\\' + r_val['cap3'])
        cap3 = InlineImage(docx_tpl, img_path_cap3, height=Mm(60))

        img_path_cap4 = os.path.join(IMAGES_PATH + '\\' + r_val['cap4'])
        cap4 = InlineImage(docx_tpl, img_path_cap4, height=Mm(60))

        img_path_cap5 = os.path.join(IMAGES_PATH + '\\' + r_val['cap5'])
        cap5 = InlineImage(docx_tpl, img_path_cap5, height=Mm(60))

        img_path_cap6 = os.path.join(IMAGES_PATH + '\\' + r_val['cap6'])
        cap6 = InlineImage(docx_tpl, img_path_cap6, height=Mm(60))

        img_path_cap7 = os.path.join(IMAGES_PATH + '\\' + r_val['cap7'])
        cap7 = InlineImage(docx_tpl, img_path_cap7, height=Mm(60))

        img_path_cap8 = os.path.join(IMAGES_PATH + '\\' + r_val['cap8'])
        cap8 = InlineImage(docx_tpl, img_path_cap8, height=Mm(60))

        img_path_cap9 = os.path.join(IMAGES_PATH + '\\' + r_val['cap9'])
        cap9 = InlineImage(docx_tpl, img_path_cap9, height=Mm(60))

        img_path_cap10 = os.path.join(IMAGES_PATH + '\\' + r_val['cap10'])
        cap10 = InlineImage(docx_tpl, img_path_cap10, height=Mm(60))

        img_path_cap11 = os.path.join(IMAGES_PATH + '\\' + r_val['cap11'])
        cap11 = InlineImage(docx_tpl, img_path_cap11, height=Mm(60))

        img_path_cap12 = os.path.join(IMAGES_PATH + '\\' + r_val['cap12'])
        cap12 = InlineImage(docx_tpl, img_path_cap12, height=Mm(60))

        img_path_sim1 = os.path.join(IMAGES_PATH + '\\' + r_val['sim1'])
        sim1 = InlineImage(docx_tpl, img_path_sim1, height=Mm(60))

        img_path_sim2 = os.path.join(IMAGES_PATH + '\\' + r_val['sim2'])
        sim2 = InlineImage(docx_tpl, img_path_sim2, height=Mm(60))

        img_path_sim3 = os.path.join(IMAGES_PATH + '\\' + r_val['sim3'])
        sim3 = InlineImage(docx_tpl, img_path_sim3, height=Mm(60))

        img_path_sim4 = os.path.join(IMAGES_PATH + '\\' + r_val['sim4'])
        sim4 = InlineImage(docx_tpl, img_path_sim4, height=Mm(60))

        img_path_sim5 = os.path.join(IMAGES_PATH + '\\' + r_val['sim5'])
        sim5 = InlineImage(docx_tpl, img_path_sim5, height=Mm(60))

        img_path_sim6 = os.path.join(IMAGES_PATH + '\\' + r_val['sim6'])
        sim6 = InlineImage(docx_tpl, img_path_sim6, height=Mm(60))

        img_path_techo = os.path.join(IMAGES_PATH + '\\' + r_val['techo'])
        techo = InlineImage(docx_tpl, img_path_techo, height=Mm(70))

        img_path_pisos = os.path.join(IMAGES_PATH + '\\' + r_val['pisos'])
        pisos = InlineImage(docx_tpl, img_path_pisos, height=Mm(50))

        img_path_puerta = os.path.join(IMAGES_PATH + '\\' + r_val['puerta'])
        puerta = InlineImage(docx_tpl, img_path_puerta, height=Mm(60))

        img_path_estantes = os.path.join(
            IMAGES_PATH + '\\' + r_val['estantes'])
        estantes = InlineImage(docx_tpl, img_path_estantes, height=Mm(60))

        img_path_site = os.path.join(IMAGES_PATH + '\\' + r_val['site'])
        site = InlineImage(docx_tpl, img_path_site, height=Mm(70))

        img_path_dh = os.path.join(IMAGES_PATH + '\\' + r_val['dh'])
        dh = InlineImage(docx_tpl, img_path_dh, height=Mm(70))

        img_path_ventanas = os.path.join(
            IMAGES_PATH + '\\' + r_val['ventanas'])
        ventanas = InlineImage(docx_tpl, img_path_ventanas, height=Mm(60))

        img_path_compresor = os.path.join(
            IMAGES_PATH + '\\' + r_val['compresor'])
        compresor = InlineImage(docx_tpl, img_path_compresor, height=Mm(50))

        img_path_quimicos = os.path.join(
            IMAGES_PATH + '\\' + r_val['quimicos'])
        quimicos = InlineImage(docx_tpl, img_path_quimicos, height=Mm(60))

        img_path_tanques_gaso = os.path.join(
            IMAGES_PATH + '\\' + r_val['tanques_gaso'])
        tanques_gaso = InlineImage(
            docx_tpl, img_path_tanques_gaso, height=Mm(60))

        img_path_paro = os.path.join(IMAGES_PATH + '\\' + r_val['paro'])
        paro = InlineImage(docx_tpl, img_path_paro, height=Mm(60))

        img_path_trampa_grasa = os.path.join(
            IMAGES_PATH + '\\' + r_val['trampa_grasa'])
        trampa_grasa = InlineImage(
            docx_tpl, img_path_trampa_grasa, height=Mm(60))

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
            'mod_estructurales': r_val['mod_estructurales'],
            'mod_arquitect': r_val['mod_arquitect'],
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
            'paros_emergencia': r_val['paros_emergencia'],
            'ubicación_pe': r_val['ubicación_pe'],
            'venteo_sist': r_val['venteo_sist'],
            'ubicación_vent': r_val['ubicación_vent'],
            'planta': r_val['planta'],
            'ubicación_plant': r_val['ubicación_plant'],
            'met_alarma': r_val['met_alarma'],
            'tipo_alarma': r_val['tipo_alarma'],
            'ubicacion_alarma': r_val['ubicacion_alarma'],
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
            'apague_motor': r_val['apague_motor'],
            'no_celular': r_val['no_celular'],
            'no_gorra_lentes': r_val['no_gorra_lentes'],
            'uso_epp': r_val['uso_epp'],
            'detectores_mov': r_val['detectores_mov'],
            'site_emer': r_val['site_emer'],
            'hidrantes': r_val['hidrantes'],
            'aspersores': r_val['aspersores'],
            'bomberos': r_val['bomberos'],
            'detector_gas': r_val['detector_gas'],
            'equipo_brigada': r_val['equipo_brigada'],
            'lampara': r_val['lampara'],
            'baterias': r_val['baterias'],
            'tambo_arena': r_val['tambo_arena'],
            'tanques': r_val['tanques'],
            'tanque_1': r_val['tanque_1'],
            'tanque_2': r_val['tanque_2'],
            'tanque_3': r_val['tanque_3'],
            'dia': r_val['dia'],
            'mes': r_val['mes'],
            'anio': r_val['anio'],
            'coordenadas': r_val['coordenadas'],
            'norte': r_val['norte'],
            'sur': r_val['sur'],
            'este': r_val['este'],
            'oeste': r_val['oeste'],
            'ley': r_val['ley'],
            'reglamento': r_val['reglamento'],
            'coord_suplente': r_val['coord_suplente'],
            'evacuacion': r_val['evacuacion'],
            'evac_suplente': r_val['evac_suplente'],
            'incendios': r_val['incendios'],
            'inc_suplente': r_val['inc_suplente'],
            'primeros_auxilios': r_val['primeros_auxilios'],
            'aux_suplente': r_val['aux_suplente'],
            'busqueda': r_val['busqueda'],
            'busq_suplente': r_val['busq_suplente'],
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
            'tipo_riesgo': r_val['tipo_riesgo'],
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
        }

        # Renderizamos usando el contexto creado
        docx_tpl.render(context)

        # Guardar documento
        if (pd.notna(r_val['pipc'])):
            nombre_pipc = '1. PIPC ' + \
                r_val['nombre_comercial'] + '.docx'
        else:
            nombre_pipc = '1. PIPC ' + \
                r_val['nombre_comercial'] + '.docx'
        docx_tpl.save(OUTPUT_PATH + '\\' + nombre_pipc)

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
