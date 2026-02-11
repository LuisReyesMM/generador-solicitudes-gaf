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

# --- LISTA DE PAÍSES ---
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
    div[data-baseweb="select"] > div { background-color: #ffffff !important; border: 1px solid #ced4da !important; border-radius: 6px !important; }
    div[data-baseweb="select"] span, div[data-baseweb="select"] div { color: #000000 !important; -webkit-text-fill-color: #000000 !important; }
    
    /* TRADUCCIÓN CAJA DE CARGA */
    [data-testid="stFileUploader"] section { background-color: #ffffff !important; border: 2px dashed #1a73e8 !important; border-radius: 8px !important; }
    [data-testid="stFileUploader"] small, [data-testid="stFileUploader"] span { display: none !important; }
    [data-testid="stFileUploader"] section::before { content: "Arrastra y suelta tu archivo PDF aquí"; display: block; font-size: 16px; font-weight: 600; margin-bottom: 5px; color: #0d2b49 !important; }
    [data-testid="stFileUploader"] section::after { content: "Límite 200MB por archivo · Solo PDF"; display: block; font-size: 13px; color: #495057 !important; margin-top: 5px; }
    [data-testid="stFileUploader"] button { background-color: #e9ecef !important; color: transparent !important; border: 1px solid #ced4da !important; position: relative; }
    [data-testid="stFileUploader"] button::after { content: "Buscar archivo"; color: #000000 !important; position: absolute; left: 50%; top: 50%; transform: translate(-50%, -50%); font-weight: 600; font-size: 14px; display: block;}
    
    .stButton button { background-color: #1a73e8 !important; border: none !important; border-radius: 8px !important; padding: 15px !important; width: 100%; }
    .stButton button p { color: #ffffff !important; font-weight: bold !important; font-size: 14px !important; display: block !important;}
    .stDownloadButton button { background-color: #28a745 !important; border: none !important; border-radius: 8px !important; padding: 20px !important; width: 100% !important; font-size: 18px !important; }
    .stDownloadButton button p { color: #ffffff !important; font-weight: bold !important; font-size: 18px !important; display: block !important;}
    .stDownloadButton button:hover { background-color: #218838 !important; transform: translateY(-2px); box-shadow: 0 4px 12px rgba(33,136,56,0.4); }
    
    #MainMenu {visibility: hidden;} footer {visibility: hidden;} header {visibility: hidden;}
    </style>
""", unsafe_allow_html=True)

TEMPLATE_PATH = "assets/plantilla.docx"
FIRMA_PATH = "assets/firma.png"

# --- LÓGICA DE EXTRACCIÓN AVANZADA ---
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
    match_calle = re.split(r'No\.?\s*Ext', texto_domicilio, flags=re.IGNORECASE)
    if match_calle: datos['calle'] = match_calle[0].replace("DOMICILIO:", "").strip()
    return datos

def extraer_info_pedimento(pdf_file):
    datos_grales = {"pedimento": "", "rfc": "", "nombre_social": "", "calle": "", "num_ext": "", "num_int": "", "cp": "", "colonia": "", "municipio": "", "estado": "CIUDAD DE MEXICO", "factura_auto": ""}
    partidas_detectadas = []
    texto_completo = ""
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            texto_completo += page.extract_text() + "\n"
    
    # 1. NOMBRE Y RFC
    try:
        parte_nombre = texto_completo.split("NOMBRE, DENOMINACION O RAZON SOCIAL:")[1]
        nombre_final = parte_nombre.split("DOMICILIO:")[0].replace("CURP:", "").replace("\n", " ")
        datos_grales['nombre_social'] = re.sub(r'\s+', ' ', nombre_final).strip()
    except: pass

    try:
        bloque_domicilio = texto_completo.split("DOMICILIO:")[1][:300] 
        bloque_domicilio_lineal = bloque_domicilio.replace("\n", " ")
        datos_direccion = limpiar_direccion_avanzada(bloque_domicilio_lineal)
        datos_grales.update(datos_direccion)
    except: pass
    
    match_rfc = re.search(r'RFC:?\s*([A-Z0-9]{12,13})', texto_completo)
    if match_rfc: datos_grales['rfc'] = match_rfc.group(1)
    
    # 2. PEDIMENTO (CORREGIDO: MANTIENE ESPACIOS)
    match_ped = re.search(r'NUM\.?\s*PEDIMENTO:?\s*([\d\s]{15,})', texto_completo)
    if match_ped: 
        datos_grales['pedimento'] = match_ped.group(1).strip()

    # 3. PRIORIDAD FACTURA (GODF sobre COVE)
    try:
        bloque = texto_completo.split("NUM. CFDI O DOCUMENTO EQUIVALENTE")[1].split("TRANSPORTE")[0]
        lineas = [l.strip() for l in bloque.split('\n') if len(l.strip()) > 5]
        cove_val, otro_val = None, None
        for l in lineas:
            if "COVE" in l: cove_val = l
            elif any(char.isdigit() for char in l): otro_val = l
        datos_grales['factura_auto'] = otro_val if otro_val else (cove_val if cove_val else "")
    except: pass

    # 4. PARTIDAS Y DESCRIPCIÓN ULTRA-PRECISA
    match_total = re.search(r'NUM\. TOTAL DE PARTIDAS:\s*(\d+)', texto_completo)
    total_ped = int(match_total.group(1)) if match_total else 999
    
    patron = re.compile(r'\b(\d{3})\s+(\d{8})\b')
    vistos = set()
    
    # PALABRAS DE FRENO (STOP WORDS)
    # Si encontramos esto, dejamos de leer la descripción inmediatamente
    palabras_freno = [
        "CLAVE", "NUM. PERMISO", "FIRMA DESCARGO", "VAL. COM.", "CANTIDAD UMT", 
        "IDENTIF", "COMPLEMENTO", "OBSERVACIONES", "VIN", "MARCA", "MODELO",
        "VALOR", "IMPORTE", "PRECIO", "NOM-", "TASA"
    ]
    
    # PALABRAS BASURA DENTRO DE LA LÍNEA
    palabras_basura = ["CHN", "IGI", "IVA", "CON."]

    for m in patron.finditer(texto_completo):
        if m.group(1) not in vistos and len(partidas_detectadas) < total_ped:
            vistos.add(m.group(1))
            
            inicio_desc = m.end()
            chunk = texto_completo[inicio_desc:inicio_desc+1500] 
            lineas = chunk.split('\n')
            desc_lineas = []
            
            for l in lineas[:30]: # Revisamos hasta 30 líneas
                l_limpia = l.strip()
                
                # 1. FRENO DE MANO: Si encontramos una palabra clave de otra sección, PARAMOS.
                if any(kw in l_limpia for kw in palabras_freno):
                    break
                
                # 2. FRENO DE NÚMEROS: Si la línea parece ser solo valores numéricos (ej: 76530 76530), PARAMOS.
                # Detecta líneas que son mayormente dígitos y puntos
                if re.match(r'^[\d\.\s,]+$', l_limpia) and len(l_limpia) > 5:
                    break

                # 3. LIMPIEZA DE INICIO: Quitamos números de columna colados al principio (ej: "0 1 1 FORMAS...")
                l_sin_numeros = re.sub(r'^[\d\s]+', '', l_limpia)
                
                # 4. FILTRADO DE CONTENIDO
                if len(l_sin_numeros) > 2:
                    if not any(b in l_sin_numeros for b in palabras_basura):
                        desc_lineas.append(l_sin_numeros)
            
            desc_final = " ".join(desc_lineas).replace(m.group(2), "").strip()
            
            partidas_detectadas.append({"secuencia": m.group(1), "fraccion": m.group(2), "producto": desc_final})
            
    return datos_grales, partidas_detectadas

# --- INTERFAZ ---
st.markdown("<h2>📄 Generador de Solicitudes GAF</h2>", unsafe_allow_html=True)
archivo_pdf = st.file_uploader("Subir Pedimento", type="pdf")

if archivo_pdf:
    if 'cache_data' not in st.session_state or st.session_state.get('pdf_name') != archivo_pdf.name:
        with st.spinner("Analizando PDF..."):
            st.session_state['cache_data'], st.session_state['cache_partidas'] = extraer_info_pedimento(archivo_pdf)
            st.session_state['pdf_name'] = archivo_pdf.name

    d = st.session_state['cache_data']
    all_partidas = st.session_state['cache_partidas']
    
    # --- SECCIÓN 1 Y 2 ---
    st.markdown('<div class="seccion-header">1. ENCABEZADO Y DATOS DEL SOLICITANTE</div>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    v_tipo_servicio = c1.text_input("Solicitud (Servicio):", placeholder="Ej. Dictamen, Constancia...")
    v_giro = c2.text_input("Giro de la empresa:", value="LOGISTICA Y DISTRIBUCION")
    c3, c4 = st.columns(2)
    v_fecha_sol = c3.date_input("Fecha de Solicitud:", value=date.today())
    v_inspector = c4.selectbox("Inspector de Ingreso:", ["Elige un elemento", "Jose de Jesus Leon", "Nicole Danae Jimenez", "Aranza Jimenez"])
    c5, c6 = st.columns(2)
    v_folio_contrato = c5.text_input("Folio de Contrato de Prestación de Servicios:")
    v_fecha_contrato = c6.date_input("Fecha De Contrato:", value=None)

    v_razon = st.text_input("Nombre y/o razón social del domicilio fiscal:", value=d['nombre_social'])
    col_d1, col_d2, col_d3 = st.columns([3, 1, 1])
    v_calle = col_d1.text_input("Calle:", value=d.get('calle', ''))
    v_ext = col_d2.text_input("No. Ext.:", value=d.get('num_ext', ''))
    v_int = col_d3.text_input("No. Int.:", value=d.get('num_int', ''))
    col_d4, col_d5 = st.columns([1, 2])
    v_cp = col_d4.text_input("C. P.:", value=d.get('cp', ''))
    v_colonia = col_d5.text_input("Colonia:", value=d.get('colonia', ''))
    col_d6, col_d7 = st.columns([1, 1])
    v_estado = col_d6.text_input("Estado:", value=d.get('estado', ''))
    v_municipio = col_d7.text_input("Alcaldía o municipio:", value=d.get('municipio', ''))
    
    v_rfc = st.text_input("R. F. C.:", value=d['rfc'])
    v_rep_legal = st.text_input("Nombre del Representante Legal:")
    c_gestor1, c_gestor2 = st.columns(2)
    v_gestor = c_gestor1.selectbox("Nombre del gestor o apoderado (Opcional):", ["Elige un elemento", "Gestor A", "Gestor B"])
    v_correo_gestor = c_gestor2.selectbox("Correo electrónico (Gestor):", ["Elige un elemento", "gestor@empresa.com", "contacto@empresa.com"])

    # --- SECCIÓN: SELECCIÓN ---
    st.markdown('<div class="seccion-header">2. SELECCIONAR PARTIDAS A PROCESAR</div>', unsafe_allow_html=True)
    st.info("💡 **Marca las casillas** de las partidas que deseas incluir en el documento único.")
    
    partidas_seleccionadas = []
    cols_check = st.columns(min(len(all_partidas), 4) if len(all_partidas) > 0 else 1)
    
    for idx, p in enumerate(all_partidas):
        with cols_check[idx % len(cols_check)]:
            if st.checkbox(f"Partida {p['secuencia']}", value=True, key=f"chk_{p['secuencia']}"):
                partidas_seleccionadas.append(p['secuencia'])

    if not partidas_seleccionadas:
        st.warning("⚠️ Debes seleccionar al menos una partida para generar el documento.")
    else:
        # --- SECCIÓN 3: EDICIÓN ---
        st.markdown('<div class="seccion-header">3. DATOS DE LOS PRODUCTOS SELECCIONADOS</div>', unsafe_allow_html=True)
        
        datos_editados = {}

        for p in all_partidas:
            if p['secuencia'] in partidas_seleccionadas:
                with st.expander(f"📦 Editar Partida {p['secuencia']} - Fracción: {p['fraccion']}"):
                    p_prod = st.text_area("Producto:", value=p.get('producto', ''), height=60, key=f"prod_{p['secuencia']}")
                    p_marca = st.text_input("Marca:", key=f"marca_{p['secuencia']}")
                    p_modelo = st.text_input("Modelo(s):", key=f"mod_{p['secuencia']}")
                    p_pais = st.selectbox("País de origen:", PAISES_ORIGEN, key=f"pais_{p['secuencia']}")
                    p_pedimento = st.text_input("*Pedimento:", value=d['pedimento'], key=f"ped_{p['secuencia']}")
                    p_factura = st.text_input("*Factura y/o lista de empaque:", value=d['factura_auto'], key=f"f_{p['secuencia']}")
                    p_lote = st.text_input("*Tamaño del lote:", key=f"lote_{p['secuencia']}")
                    p_umc = st.text_input("*UMC: (Unidad de Medida Comercial)", value="PIEZA", key=f"umc_{p['secuencia']}")
                    p_folios = st.text_input("Folio(s):", key=f"fol_{p['secuencia']}")
                    
                    datos_editados[p['secuencia']] = {
                        'secuencia': p['secuencia'], 'producto': p_prod, 'marca': p_marca,
                        'modelo': p_modelo, 'pais': p_pais if p_pais != "Elige un país..." else "", 
                        'pedimento': p_pedimento, 'factura': p_factura, 'lote': p_lote, 'umc': p_umc, 
                        'fraccion': p['fraccion'], 'folios': p_folios
                    }

        # --- SECCIÓN 4: INSPECCIÓN ---
        st.markdown('<div class="seccion-header">4. DATOS DE INSPECCIÓN</div>', unsafe_allow_html=True)
        usar_fiscal = st.checkbox("📍 Usar el mismo domicilio fiscal", value=True)
        i_calle = st.text_input("Calle (Insp):", value=v_calle, disabled=usar_fiscal, key="i_calle")
        ci1, ci2, ci3 = st.columns([1, 1, 2])
        i_ext = ci1.text_input("No. Ext.:", value=v_ext, disabled=usar_fiscal, key="i_ext")
        i_int = ci2.text_input("No. Int.:", value=v_int, disabled=usar_fiscal, key="i_int")
        i_col = ci3.text_input("Colonia:", value=v_colonia, disabled=usar_fiscal, key="i_col")
        ci4, ci5 = st.columns(2)
        i_cp = ci4.text_input("C.P.:", value=v_cp, disabled=usar_fiscal, key="i_cp")
        i_mun = ci5.text_input("Alcaldía/Municipio:", value=v_municipio, disabled=usar_fiscal, key="i_mun")
        ci6, ci7 = st.columns(2)
        i_tel = ci6.text_input("Teléfono:", key="i_tel")
        i_edo = ci7.text_input("Estado:", value=v_estado if usar_fiscal else "", disabled=usar_fiscal, key="i_edo")
        i_rep_insp = st.selectbox("Representante inspección:", ["Elige un elemento", "Representante 1", "Representante 2"])
        i_correo_insp = st.selectbox("Correo inspección:", ["Elige un elemento", "correo1@empresa.com", "correo2@empresa.com"])
        v_fecha_prog = st.date_input("Fecha programada:")
        v_observaciones = st.text_area("Observaciones:", height=60)
        v_firmante = st.selectbox("RESPONSABLE FIRMANTE:", ["Jose de Jesus Leon", "Nicole Danae Jimenez", "Aranza Jimenez"])

        # --- SECCIÓN 5: DESCARGA FINAL ---
        st.markdown('<div class="seccion-header">5. DESCARGAR DOCUMENTO FINAL</div>', unsafe_allow_html=True)
        
        lista_final_para_word = []
        for p in all_partidas:
            if p['secuencia'] in partidas_seleccionadas:
                lista_final_para_word.append(datos_editados[p['secuencia']])

        if st.button("📄 GENERAR DOCUMENTO UNIFICADO (WORD)", type="primary", use_container_width=True):
            
            marca_dictamen = "X" if "DICTAMEN" in v_tipo_servicio.upper() else ""
            marca_constancia = "X" if "CONSTANCIA" in v_tipo_servicio.upper() else ""
            gestor_final = "" if v_gestor == "Elige un elemento" else v_gestor
            correo_gestor_final = "" if v_correo_gestor == "Elige un elemento" else v_correo_gestor
            rep_ins_final = "" if i_rep_insp == "Elige un elemento" else i_rep_insp
            correo_ins_final = "" if i_correo_insp == "Elige un elemento" else i_correo_insp
            inspector_final = "" if v_inspector == "Elige un elemento" else v_inspector
            
            contexto = {
                'solicitud': v_tipo_servicio, 'giro': v_giro, 'fecha_solicitud': str(v_fecha_sol),
                'inspector': inspector_final, 'folio_contrato': v_folio_contrato, 
                'fecha_contrato': str(v_fecha_contrato) if v_fecha_contrato else "",
                'nombre_social': v_razon, 'calle': v_calle, 'num_ext': v_ext, 'num_int': v_int,
                'cp': v_cp, 'colonia': v_colonia, 'estado': v_estado, 'municipio': v_municipio,
                'telefono': "", 'correo': "", 'rfc': v_rfc,
                'representante': v_rep_legal, 'gestor': gestor_final, 'correo_gestor': correo_gestor_final,
                'calle_ins': i_calle, 'ext_ins': i_ext, 'int_ins': i_int, 'col_ins': i_col,
                'cp_ins': i_cp, 'mun_ins': i_mun, 'tel_ins': i_tel, 'edo_ins': i_edo,
                'rep_ins': rep_ins_final, 'correo_ins': correo_ins_final,
                'fecha_programada': str(v_fecha_prog), 'observaciones': v_observaciones,
                'nombre_firmante': v_firmante, 'dictamen': marca_dictamen, 'constancia': marca_constancia,
                'lista_productos': lista_final_para_word
            }
            
            doc = DocxTemplate(TEMPLATE_PATH)
            if os.path.exists(FIRMA_PATH):
                contexto['imagen_firma'] = InlineImage(doc, FIRMA_PATH, width=Mm(35))
            else:
                contexto['imagen_firma'] = ""
            
            doc.render(contexto)
            buf = BytesIO()
            doc.save(buf)
            buf.seek(0)
            
            st.download_button(
                label="⬇️ DESCARGAR WORD CON TODAS LAS PARTIDAS",
                data=buf,
                file_name=f"Solicitud_Unificada_{v_rfc}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )

else:
    st.markdown("""
    <div style="text-align: center; padding: 50px; background-color: #ffffff; border-radius: 10px; border: 2px dashed #1a73e8;">
        <h2 style="color: #0d2b49;">👋 Bienvenido al Gestor de Solicitudes</h2>
        <p style="font-size: 16px; color: #495057;">Por favor, arrastra un archivo <b>PDF de Pedimento</b> en la barra superior para comenzar a extraer los datos.</p>
    </div>
    """, unsafe_allow_html=True)
