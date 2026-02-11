import streamlit as st
import pdfplumber
import re
import os
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
from io import BytesIO
from datetime import date

# --- CONFIGURACIÓN ---
st.set_page_config(layout="wide", page_title="Sistema de Inspección GAF", page_icon="🏢")

# --- LISTA DE PAÍSES (ORDENADA ALFABÉTICAMENTE) ---
PAISES_ORIGEN = [
    "Elige un país...", "Andorra", "Angola", "Anguila", "Antártida", "Antigua y Barbuda", "Antillas Neerlandesas", "Arabia Saudita", 
    "Argelia", "Argentina", "Armenia", "Aruba", "Australia", "Austria", "Azerbaiyán", "Bahamas", "Bahréin", 
    "Bangladesh", "Barbados", "Bélgica", "Belice", "Benín", "Bermudas", "Bielorrusia", "Bolivia", "Bonaire", 
    "Bosnia y Herzegovina", "Botswana", "Brasil", "Brunéi", "Bulgaria", "Burkina Faso", "Burundi", "Bután", 
    "Cabo Verde", "Caimán", "Camboya", "Camerún", "Canadá", "Canal", "Chad", "Chile", "China", "Chipre", 
    "Ciudad del Vaticano", "Cocos", "Colombia", "Comoras", "Comunidad Europea", "Congo", "Cook", "Corea del Norte", 
    "Corea del Sur", "Costa Rica", "Costa de Marfil", "Croacia", "Cuba", "Curazao", "Dinamarca", "Djibouti", 
    "Dominica", "Ecuador", "Egipto", "El Salvador", "Emiratos Árabes Unidos", "Eritrea", "Eslovenia", "España", 
    "Estado Federado de Micronesia", "Estados Unidos de América", "Estonia", "Etiopía", "Fidji", "Filipinas", 
    "Finlandia", "Francia", "Franja de Gaza", "Gabonesa", "Gambia", "Georgia", "Georgia del Sur e Islas Sandwich del Sur", 
    "Ghana", "Gibraltar", "Granada", "Grecia", "Groenlandia", "Guadalupe", "Guam", "Guatemala", "Guernsey", 
    "Guinea", "Guinea Ecuatorial", "Guinea-Bissau", "Guyana", "Guyana Francesa", "Haití", "Honduras", "Hong Kong", 
    "Hungría", "India", "Indonesia", "Irak", "Irán", "Irlanda", "Islandia", "Isla Bouvet", "Isla Feroe", "Isla de Man", 
    "Islas Aland", "Islas Australianas Cocos", "Islas Caimán", "Islas Comoras", "Islas Cook", "Islas Heard y Mcdonald", 
    "Islas Malvinas", "Islas Marianas Septentrionales", "Islas Marshall", "Islas Normandas Canal", "Islas Salomón", 
    "Islas Svalbard y Jan Mayen", "Islas Tokelau", "Islas Wallis y Futuna", "Islas del Canal", "Israel", "Italia", 
    "Jamaica", "Japón", "Jersey", "Jordania", "Kazakhstan", "Keeling Cocos", "Kenya", "Kiribati", "Kuwait", 
    "Kyrgyzstan", "Lesotho", "Letonia", "Liberia", "Libia", "Liechtenstein", "Lituania", "Luxemburgo", "Líbano", 
    "MLI Malí", "Macao", "Macedonia", "Madagascar", "Malasia", "Malawi", "Maldivas", "Malta", "Marruecos", 
    "Martinica", "Mauritania", "Mauricio", "Mayotte", "Moldavia", "Mongolia", "Monserrat", "Montenegro", 
    "Mozambique", "Myanmar", "México", "Mónaco", "Namibia", "Nauru", "Navidad", "Nepal", "Nicaragua", "Niger", 
    "Nigeria", "Nive", "Norfolk", "Noruega", "Nueva Caledonia", "Nueva Zelanda", "Omán", "Pacífico", "Países Bajos", 
    "Países no declarados", "Pakistán", "Palau", "Palestina", "Panamá", "Papúa Nueva Guinea", "Paraguay", "Perú", 
    "Pitcairns", "Polinesia Francesa", "Polonia", "Portugal", "Puerto Rico", "Qatar", "Reino Unido de la Gran Bretaña e Irlanda del Norte", 
    "Reino de Tonga", "República Árabe Saharavi Democrática", "República Centroafricana", "República Checa", 
    "República Democrática Popular Laos", "República Dominicana", "República Eslovaca", "República Popular del Congo", 
    "República Ruandesa", "República Togolesa", "República de Djibouti", "República de Serbia", "República del Congo", 
    "Reunión (Departamento de la) (Francia)", "Rumania", "Rusia", "Sahara Occidental", "Samoa", "Samoa Americana", 
    "San Bartolomé", "San Cristóbal y Nieves", "San Eustaquio y Saba", "San Marino", "San Martín", "San Pedro y Miquelón", 
    "San Vicente y las Granadinas", "Santa Elena", "Santa Lucía", "Santo Tomé y Príncipe", "Senegal", "Seychelles", 
    "Sierra Leona", "Singapur", "Sint Maarten", "Siria", "Somalia", "Sri Lanka", "Sudáfrica", "Sudán", "Sudán del Sur", 
    "Suecia", "Suiza", "Surinam", "Swazilandia", "Tadjikistan", "Tailandia", "Taiwán", "Tanzania", 
    "Territorios Británicos del Océano Índico", "Territorios Franceses, Australes y Antárticos", "Timor Oriental", 
    "Togo", "Tonga", "Trinidad y Tobago", "Turcas y Caicos", "Turkmenistan", "Turquía", "Tuvalu", "Túnez", "Ucrania", 
    "Uganda", "Uruguay", "Uzbejistan", "Vanuatu", "Venezuela", "Vietnam", "Vírgenes. Islas", "Yemen", "Zambia", 
    "Zimbabwe", "Zona Neutral Iraq-Arabia Saudita", "Zona del Canal de Panamá"
]

# --- DISEÑO UI / UX ---
st.markdown("""
    <style>
    .stApp, .block-container { background-color: #f0f4f8 !important; }
    .stApp p, .stApp span, .stApp label, div[data-testid="stMarkdownContainer"] { color: #212529 !important; }
    .stApp h1, .stApp h2, .stApp h3 { color: #0d2b49 !important; font-weight: 800 !important; }
    .block-container { background-color: #ffffff !important; padding: 3rem !important; border-radius: 12px !important; box-shadow: 0px 8px 24px rgba(0,0,0,0.08) !important; max-width: 1100px !important; margin-top: 2rem !important; border-top: 8px solid #0d2b49 !important; }
    .seccion-header { background-color: #e9ecef !important; color: #0d2b49 !important; padding: 12px 18px !important; border-left: 6px solid #0d2b49 !important; font-weight: 700 !important; text-transform: uppercase !important; font-size: 15px !important; margin-top: 35px !important; margin-bottom: 20px !important; border-radius: 4px !important; }
    .stTextInput input, .stTextArea textarea, .stDateInput input { background-color: #ffffff !important; color: #000000 !important; -webkit-text-fill-color: #000000 !important; border: 1px solid #ced4da !important; border-radius: 6px !important; padding: 10px !important; }
    .stTextInput input:disabled, .stTextArea textarea:disabled { background-color: #eef2f5 !important; color: #000000 !important; -webkit-text-fill-color: #000000 !important; opacity: 1 !important; cursor: not-allowed; font-weight: 500 !important; }
    .stTextInput input:focus { border-color: #1a73e8 !important; box-shadow: 0 0 0 1px #1a73e8 !important; }
    div[data-baseweb="select"] > div { background-color: #ffffff !important; border: 1px solid #ced4da !important; border-radius: 6px !important; }
    div[data-baseweb="select"] span, div[data-baseweb="select"] div { color: #000000 !important; -webkit-text-fill-color: #000000 !important; }
    div[data-baseweb="popover"], div[data-baseweb="popover"] ul, div[data-baseweb="popover"] li { background-color: #ffffff !important; color: #000000 !important; }
    div[data-baseweb="popover"] li:hover { background-color: #e9ecef !important; color: #0d2b49 !important; }
    [data-testid="stFileUploader"] section { background-color: #ffffff !important; border: 2px dashed #1a73e8 !important; border-radius: 8px !important; }
    [data-testid="stFileUploader"] section * { color: #0d2b49 !important; }
    [data-testid="stFileUploader"] button { background-color: #e9ecef !important; color: #000000 !important; border: 1px solid #ced4da !important; }
    label p { font-weight: 600 !important; color: #495057 !important; font-size: 13px !important; margin-bottom: 5px !important; }
    .stCheckbox p { font-weight: bold !important; color: #0d2b49 !important; }
    div[data-testid="stExpander"] details summary p { color: #0d2b49 !important; font-weight: 800 !important; font-size: 16px !important; }
    div[data-testid="stExpander"] details summary svg { fill: #0d2b49 !important; color: #0d2b49 !important; }
    div[data-testid="stExpander"] { border: 2px solid #1a73e8 !important; border-radius: 8px !important; background-color: #f8fbff !important; margin-bottom: 15px !important; }
    [data-testid="stExpanderDetails"] { border-top: 1px solid #dee2e6 !important; background-color: #ffffff !important; padding: 20px !important; border-radius: 0 0 8px 8px !important; }
    .stButton button { background-color: #1a73e8 !important; border: none !important; border-radius: 8px !important; padding: 15px !important; transition: all 0.3s ease; }
    .stButton button p { color: #ffffff !important; font-weight: bold !important; font-size: 14px !important; }
    .stButton button:hover { background-color: #1557b0 !important; transform: translateY(-2px); box-shadow: 0 4px 10px rgba(26,115,232,0.4); }
    .stDownloadButton button { background-color: #28a745 !important; border: none !important; border-radius: 8px !important; padding: 15px !important; width: 100% !important; }
    .stDownloadButton button p { color: #ffffff !important; font-weight: bold !important; font-size: 15px !important; }
    .stDownloadButton button:hover { background-color: #218838 !important; box-shadow: 0 4px 10px rgba(40,167,69,0.4) !important; }
    #MainMenu {visibility: hidden;} footer {visibility: hidden;} header {visibility: hidden;}
    </style>
""", unsafe_allow_html=True)

TEMPLATE_PATH = "assets/plantilla.docx"
FIRMA_PATH = "assets/firma.png"

# --- LÓGICA DE EXTRACCIÓN ---
def limpiar_direccion_avanzada(texto_domicilio):
    datos = {"calle": "", "num_ext": "", "num_int": "", "colonia": "", "cp": "", "municipio": "", "estado": "CIUDAD DE MEXICO"}
    if not texto_domicilio: return datos
    match_cp = re.search(r'C\.?P\.?\s*(\d{5})', texto_domicilio)
    if match_cp: datos['cp'] = match_cp.group(1)
    match_ext = re.search(r'No\.?\s*Ext\.?[\s\.]*([^\n]+?)(?=\s+No\.|\s+COLONIA|\s+COL)', texto_domicilio, re.IGNORECASE)
    if match_ext: datos['num_ext'] = match_ext.group(1).strip()
    match_int = re.search(r'No\.?\s*Int\.?[\s\.]*([^\n]+?)(?=\s+COLONIA|\s+COL)', texto_domicilio, re.IGNORECASE)
    if match_int: datos['num_int'] = match_int.group(1).strip()
    match_bloque = re.search(r'COLONIA\s+(.*?)\s+C\.?P\.?', texto_domicilio, re.IGNORECASE)
    if match_bloque:
        bloque = match_bloque.group(1).strip()
        partes = bloque.split(',')
        if len(partes) >= 1: datos['colonia'] = partes[0].strip()
        if len(partes) >= 2: datos['municipio'] = partes[1].strip()
        else:
            if "AZCAPOTZALCO" in bloque.upper(): datos['municipio'] = "AZCAPOTZALCO"
    match_calle = re.split(r'No\.?\s*Ext', texto_domicilio, flags=re.IGNORECASE)
    if match_calle: datos['calle'] = match_calle[0].replace("DOMICILIO:", "").strip()
    return datos

def extraer_info_pedimento(pdf_file):
    datos_grales = {"pedimento": "", "rfc": "", "nombre_social": "", "calle": "", "num_ext": "", "num_int": "", "cp": "", "colonia": "", "municipio": "", "estado": "CIUDAD DE MEXICO", "acuse_valor": ""}
    partidas = []
    texto_completo = ""
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            texto_completo += page.extract_text() + "\n"
    
    # EXTRACCIÓN DE NOMBRE CORREGIDA
    try:
        parte_1 = texto_completo.split("NOMBRE, DENOMINACION O RAZON SOCIAL:")[1]
        nombre_sucio = parte_1.split("DOMICILIO:")[0]
        # Limpiamos la palabra "CURP:" y saltos de línea
        nombre_limpio = nombre_sucio.replace("CURP:", "").replace("CURP", "").replace("\n", " ")
        # Quitamos espacios múltiples
        datos_grales['nombre_social'] = re.sub(r'\s+', ' ', nombre_limpio).strip()
    except:
        datos_grales['nombre_social'] = ""
        
    try:
        bloque_domicilio = texto_completo.split("DOMICILIO:")[1][:300] 
        bloque_domicilio_lineal = bloque_domicilio.replace("\n", " ")
        datos_direccion = limpiar_direccion_avanzada(bloque_domicilio_lineal)
        datos_grales.update(datos_direccion)
    except:
        pass
    match_ped = re.search(r'NUM\.?\s*PEDIMENTO:?\s*([\d\s]{15,})', texto_completo)
    if match_ped: datos_grales['pedimento'] = match_ped.group(1).replace(" ", "")
    match_rfc = re.search(r'RFC:?\s*([A-Z0-9]{12,13})', texto_completo)
    if match_rfc: datos_grales['rfc'] = match_rfc.group(1)
    try:
        if "ACUSE DE VALOR" in texto_completo:
            bloque_acuse = texto_completo.split("ACUSE DE VALOR")[1][:200]
            match_cove = re.search(r'(COVE[A-Z0-9]+)', bloque_acuse)
            if match_cove:
                datos_grales['acuse_valor'] = match_cove.group(1).strip()
            else:
                palabras_ignoradas = ["VINCULACION", "INCOTERM", "TRANSPORTE", "IDENTIFICACION"]
                cadenas = re.findall(r'[A-Z0-9]{10,}', bloque_acuse)
                for cadena in cadenas:
                    if cadena not in palabras_ignoradas:
                        datos_grales['acuse_valor'] = cadena.strip()
                        break
    except Exception as e:
        datos_grales['acuse_valor'] = ""
    match_total = re.search(r'NUM\. TOTAL DE PARTIDAS:\s*(\d+)', texto_completo)
    total_esperado = int(match_total.group(1)) if match_total else 999
    patron_partida = re.compile(r'\b(\d{3})\s+(\d{8})\b')
    matches = patron_partida.finditer(texto_completo)
    secuencias_vistas = set()
    for match in matches:
        sec = match.group(1)
        fraccion = match.group(2)
        if sec in secuencias_vistas: continue
        if len(partidas) >= total_esperado: break
        secuencias_vistas.add(sec)
        inicio_desc = match.end()
        chunk = texto_completo[inicio_desc:inicio_desc+400]
        lineas = chunk.split('\n')
        desc_lineas = []
        for l in lineas[:5]:
            if len(l.strip()) > 3 and not re.match(r'^[\d\.\s]+$', l):
                if "CHN" not in l and "IGI" not in l and "IVA" not in l:
                    desc_lineas.append(l.strip())
        desc_limpia = " ".join(desc_lineas).replace(fraccion, "").strip()
        partidas.append({"secuencia": sec, "fraccion": fraccion, "producto": desc_limpia})
    return datos_grales, partidas

# --- INTERFAZ GRÁFICA ---
st.markdown("<h2>📄 Generador de Solicitudes GAF</h2>", unsafe_allow_html=True)
col_file, col_info = st.columns([1, 2])
with col_file: archivo_pdf = st.file_uploader("📂 Arrastra tu Pedimento (PDF) aquí", type="pdf")

if archivo_pdf:
    if 'datos_base' not in st.session_state or st.session_state.get('pdf_name') != archivo_pdf.name:
        with st.spinner("Analizando y estructurando datos..."):
            st.session_state['datos_base'], st.session_state['partidas'] = extraer_info_pedimento(archivo_pdf)
            st.session_state['pdf_name'] = archivo_pdf.name

    datos = st.session_state['datos_base']
    partidas = st.session_state['partidas']

    # 1. ENCABEZADO
    st.markdown('<div class="seccion-header">1. ENCABEZADO DE SOLICITUD</div>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1: v_tipo_servicio = st.text_input("Solicitud (Servicio):", value="")
    with c2: v_giro = st.text_input("Giro de la empresa:", value="LOGISTICA Y DISTRIBUCION")
    c3, c4 = st.columns(2)
    with c3: v_fecha_sol = st.date_input("Fecha de Solicitud:", value=date.today())
    with c4: v_inspector = st.selectbox("Inspector de Ingreso:", ["Elige un elemento", "Jose de Jesus Leon", "Nicole Danae Jimenez", "Aranza Jimenez"])
    c5, c6 = st.columns(2)
    with c5: v_folio_contrato = st.text_input("Folio de Contrato de Prestación de Servicios:")
    with c6: v_fecha_contrato = st.date_input("Fecha De Contrato:", value=None)

    # 2. SOLICITANTE
    st.markdown('<div class="seccion-header">2. DATOS DEL SOLICITANTE (FISCAL)</div>', unsafe_allow_html=True)
    v_razon = st.text_input("Nombre y/o razón social del domicilio fiscal:", value=datos.get('nombre_social'))
    col_d1, col_d2, col_d3 = st.columns([3, 1, 1])
    v_calle = col_d1.text_input("Calle:", value=datos.get('calle'))
    v_ext = col_d2.text_input("No. Ext.:", value=datos.get('num_ext'))
    v_int = col_d3.text_input("No. Int.:", value=datos.get('num_int'))
    col_d4, col_d5 = st.columns([1, 2])
    v_cp = col_d4.text_input("C. P.:", value=datos.get('cp'))
    v_colonia = col_d5.text_input("Colonia:", value=datos.get('colonia'))
    col_d6, col_d7 = st.columns([1, 1])
    v_estado = col_d6.text_input("Estado:", value=datos.get('estado'))
    v_municipio = col_d7.text_input("Alcaldía o municipio:", value=datos.get('municipio'))
    col_c1, col_c2 = st.columns([1, 1])
    v_telefono = col_c1.text_input("Teléfono:")
    v_correo = col_c2.text_input("Correo electrónico:")
    v_rfc = st.text_input("R. F. C.:", value=datos.get('rfc'))
    v_rep_legal = st.text_input("Nombre del Representante Legal:")
    v_gestor = st.selectbox("Nombre del gestor o apoderado que realiza el trámite (Opcional):", ["Elige un elemento", "Jose Leon", "Danna Jimenez", "Aranza Jimenez"])
    v_correo_gestor = st.selectbox("Correo electrónico (Gestor):", ["Elige un elemento", "jose@gaf.com", "danna@gaf.com", "aranza@gaf.com"])

    # 3. PRODUCTOS
    st.markdown('<div class="seccion-header">3. DATOS DEL PRODUCTO (PARTIDAS DETECTADAS)</div>', unsafe_allow_html=True)
    st.info(f"✅ Se detectaron exitosamente **{len(partidas)} partidas**.")
    partidas_editadas = []
    if partidas:
        for index, p in enumerate(partidas):
            with st.expander(f"📦 Editar Partida {p['secuencia']} - Fracción: {p['fraccion']}", expanded=(index==0)):
                p_prod = st.text_area("Producto:", value=p['producto'], height=60, key=f"prod_{index}")
                p_marca = st.text_input("Marca:", key=f"marca_{index}")
                p_modelo = st.text_input("Modelo(s):", key=f"mod_{index}")
                p_pais = st.selectbox("País de origen:", PAISES_ORIGEN, key=f"pais_{index}")
                p_pedimento = st.text_input("*Pedimento:", value=datos.get('pedimento'), key=f"ped_{index}")
                p_factura = st.text_input("*Factura y/o lista de empaque:", value=datos.get('acuse_valor', ''), key=f"fac_{index}")
                p_lote = st.text_input("*Tamaño del lote:", key=f"lote_{index}")
                p_umc = st.text_input("*UMC: (Unidad de Medida Comercial)", value="PIEZA", key=f"umc_{index}")
                p_fraccion = st.text_input("*Fracción arancelaria:", value=p['fraccion'], key=f"frac_{index}")
                p_folios = st.text_input("Folio(s):", key=f"fol_{index}")
                partidas_editadas.append({
                    'secuencia': p['secuencia'], 'producto': p_prod, 'marca': p_marca,
                    'modelo': p_modelo, 'pais': p_pais, 'pedimento': p_pedimento,
                    'factura': p_factura, 'lote': p_lote, 'umc': p_umc, 
                    'fraccion': p_fraccion, 'folios': p_folios
                })

    # 4. INSPECCIÓN
    st.markdown('<div class="seccion-header">4. DATOS DEL DOMICILIO DE INSPECCIÓN</div>', unsafe_allow_html=True)
    usar_fiscal = st.checkbox("📍 Usar el mismo domicilio fiscal capturado arriba", value=True)
    i_calle = st.text_input("Calle:", value=v_calle, disabled=usar_fiscal, key="i_calle")
    ci1, ci2, ci3 = st.columns([1, 1, 2])
    i_ext = ci1.text_input("No. Ext.:", value=v_ext, disabled=usar_fiscal, key="i_ext")
    i_int = ci2.text_input("No. Int.:", value=v_int, disabled=usar_fiscal, key="i_int")
    i_col = ci3.text_input("Colonia:", value=v_colonia, disabled=usar_fiscal, key="i_col")
    ci4, ci5 = st.columns(2)
    i_cp = ci4.text_input("C.P.:", value=v_cp, disabled=usar_fiscal, key="i_cp")
    i_mun = ci5.text_input("Alcaldía o Municipio:", value=v_municipio, disabled=usar_fiscal, key="i_mun")
    ci6, ci7 = st.columns(2)
    i_tel = ci6.text_input("Teléfono (Inspección):", key="i_tel")
    i_edo = ci7.text_input("Estado:", value=v_estado if usar_fiscal else "", disabled=usar_fiscal, key="i_edo")
    i_rep_insp = st.selectbox("Nombre del representante autorizado para recibir la visita de inspección:", ["Elige un elemento", "Jose ", "Danna ", "Aranza"])
    i_correo_insp = st.selectbox("Correo electrónico (Inspección):", ["Elige un elemento", "anna@gaf.com", "danna@gaf.com", "aranza@gaf.com"])
    v_fecha_prog = st.date_input("Fecha programada para la inspección:")
    v_observaciones = st.text_area("Observaciones:", height=60)
    st.markdown("---")
    v_firmante = st.selectbox("RESPONSABLE DEL ORGANISMO DE INSPECCIÓN:", ["Jose de Jesus Leon", "Nicole Danae Jimenez", "Aranza Jimenez"])

    # 5. GENERACIÓN
    st.markdown('<div class="seccion-header">5. DESCARGAR DOCUMENTOS WORD</div>', unsafe_allow_html=True)
    cols_botones = st.columns(len(partidas_editadas) if len(partidas_editadas) > 0 else 1)
    
    for idx, p_data in enumerate(partidas_editadas):
        with cols_botones[idx % len(cols_botones)]:
            if st.button(f"⚙️ GENERAR PARTIDA {p_data['secuencia']}", key=f"btn_gen_{idx}", use_container_width=True):
                
                marca_dictamen = "X" if "DICTAMEN" in v_tipo_servicio.upper() else ""
                marca_constancia = "X" if "CONSTANCIA" in v_tipo_servicio.upper() else ""

                gestor_final = "" if v_gestor == "Elige un elemento" else v_gestor
                correo_gestor_final = "" if v_correo_gestor == "Elige un elemento" else v_correo_gestor
                rep_ins_final = "" if i_rep_insp == "Elige un elemento" else i_rep_insp
                correo_ins_final = "" if i_correo_insp == "Elige un elemento" else i_correo_insp
                inspector_final = "" if v_inspector == "Elige un elemento" else v_inspector
                pais_final = "" if p_data['pais'] == "Elige un país..." else p_data['pais']
                
                contexto = {
                    'giro': v_giro, 'solicitud': v_tipo_servicio, 'fecha_solicitud': str(v_fecha_sol),
                    'inspector': inspector_final, 'folio_contrato': v_folio_contrato, 
                    'fecha_contrato': str(v_fecha_contrato) if v_fecha_contrato else "",
                    'nombre_social': v_razon, 'calle': v_calle, 'num_ext': v_ext, 'num_int': v_int,
                    'cp': v_cp, 'colonia': v_colonia, 'estado': v_estado, 'municipio': v_municipio,
                    'telefono': v_telefono, 'correo': v_correo, 'rfc': v_rfc,
                    'representante': v_rep_legal, 'gestor': gestor_final, 'correo_gestor': correo_gestor_final,
                    
                    'producto': p_data['producto'], 'marca': p_data['marca'], 'modelo': p_data['modelo'],
                    'pais': pais_final, 'pedimento': p_data['pedimento'], 'factura': p_data['factura'],
                    'lote': p_data['lote'], 'umc': p_data['umc'], 'fraccion': p_data['fraccion'], 'folios': p_data['folios'],
                    
                    'calle_ins': i_calle, 'ext_ins': i_ext, 'int_ins': i_int, 'col_ins': i_col,
                    'cp_ins': i_cp, 'mun_ins': i_mun, 'tel_ins': i_tel, 'edo_ins': i_edo,
                    'rep_ins': rep_ins_final, 'correo_ins': correo_ins_final,
                    'fecha_programada': str(v_fecha_prog), 'observaciones': v_observaciones,
                    'nombre_firmante': v_firmante,
                    'dictamen': marca_dictamen,
                    'constancia': marca_constancia
                }
                
                doc = DocxTemplate(TEMPLATE_PATH)
                if os.path.exists(FIRMA_PATH):
                    contexto['imagen_firma'] = InlineImage(doc, FIRMA_PATH, width=Mm(35))
                else:
                    contexto['imagen_firma'] = ""
                doc.render(contexto)
                buffer = BytesIO()
                doc.save(buffer)
                buffer.seek(0)
                
                st.download_button(
                    label=f"⬇️ GUARDAR PARTIDA {p_data['secuencia']}",
                    data=buffer,
                    file_name=f"Solicitud_{v_rfc}_Partida_{p_data['secuencia']}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key=f"dl_real_{idx}",
                    use_container_width=True
                )

else:
    st.markdown("""
    <div style="text-align: center; padding: 50px; background-color: #ffffff; border-radius: 10px; border: 2px dashed #1a73e8;">
        <h2 style="color: #0d2b49;">👋 Bienvenido al Gestor de Solicitudes</h2>
        <p style="font-size: 16px; color: #495057;">Por favor, arrastra un archivo <b>PDF de Pedimento</b> en la barra superior para comenzar a extraer los datos.</p>
    </div>
    """, unsafe_allow_html=True)