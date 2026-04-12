import streamlit as st 
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from io import BytesIO
import os
from datetime import datetime
import sib_api_v3_sdk
from sib_api_v3_sdk.rest import ApiException
import base64

# =========================
# 🔧 CONFIG GLOBAL EXCEL
# =========================
RUTA_EXCEL = r"C:\Sistema de Diplomas EAC\alumnos.xlsx"

import ssl
import certifi
os.environ['SSL_CERT_FILE'] = certifi.where()
ssl._create_default_https_context = ssl.create_default_context

# =========================
# 🔍 VERIFICACIÓN POR QR (AQUÍ VA)
# =========================
params = st.query_params

if "verificar" in params:
    codigo = params["verificar"]

    try:
        if os.path.exists(RUTA_EXCEL):
            df = pd.read_excel(RUTA_EXCEL)
        else:
            st.error("❌ Base de datos no encontrada")
            st.stop()

        if "Codigo_Verificacion" in df.columns:

            resultado = df[df['Codigo_Verificacion'] == codigo]

            if not resultado.empty:
                alumno = resultado.iloc[0]

                st.set_page_config(layout="centered")

                st.markdown("## ✅ DIPLOMA VERIFICADO")
                st.markdown("---")

                st.write(f"👤 **Nombre:** {alumno['Nombre_Completo']}")
                st.write(f"🎓 **Nivel:** {alumno['Nivel']}")
                st.write(f"🏅 **Tipo:** {alumno['Tipo']}")
                st.write(f"📅 **Fecha:** {alumno['Fecha_Curso']}")

            else:
                st.error("❌ Código no válido")

        else:
            st.error("❌ Base de datos inválida")

    except Exception as e:
        st.error(f"❌ Error: {str(e)}")

    st.stop()

# =========================
# 🔥 FONDO DIFUMINADO
# =========================
def set_bg():
    try:
        with open("assets/fondo.jpg", "rb") as f:
            data = base64.b64encode(f.read()).decode()

        st.markdown(f"""
        <style>
        .stApp {{
            background: linear-gradient(rgba(255,255,255,0.00), rgba(255,255,255,0.00)),
            url("data:image/jpg;base64,{data}");
            background-size: cover;
            background-position: center;
        }}
        </style>
        """, unsafe_allow_html=True)
    except:
        pass

set_bg()

# =========================
# CONFIG
# =========================
st.set_page_config(
    page_title="Sistema de Diplomas EAC",
    page_icon="⛪",
    layout="centered"
)

# =========================
# SIDEBAR
# =========================
try:
    st.sidebar.image("assets/logo.png", use_container_width=True)
except:
    st.sidebar.warning("⚠️ Logo no encontrado")

st.sidebar.markdown("## 🎓 Panel Administrativo")
st.sidebar.markdown("Sistema de Diplomas EAC")

# =========================
# 🎛️ MENÚ PROFESIONAL
# =========================

if "menu" not in st.session_state:
    st.session_state["menu"] = "Inicio"

st.sidebar.markdown("## 📂 Navegación")

def menu_button(label, key, icon):
    if st.sidebar.button(f"{icon}  {label}", key=key):
        st.session_state["menu"] = label

menu_button("Inicio", "btn_inicio", "🏠")
menu_button("Alumnos", "btn_alumnos", "📊")
menu_button("Diplomas", "btn_diplomas", "📄")
menu_button("Envíos", "btn_envios", "📨")

menu = st.session_state["menu"]



# =========================
# 🎨 ESTILO PREMIUM CATÓLICO
# =========================
st.markdown("""
<style>

/* Fondo general */
.stApp {
    background-color: #f8f5ef;
    color: #2c2c2c;
}

/* Header */
.main-header {
    font-size: 2.8rem;
    font-weight: 700;
    text-align: center;
    color: #1e3a5f;
}

/* Subheader */
.sub-header {
    font-size: 1.2rem;
    text-align: center;
    color: #c9a227;
    font-style: italic;
}

/* Tarjeta */
.block-container {
    background-color: rgba(255,255,255,0.96);
    padding: 2rem;
    border-radius: 14px;
    box-shadow: 0px 6px 25px rgba(0,0,0,0.08);
}

/* Botones generales */
.stButton>button {
    background: linear-gradient(135deg, #d4af37, #b8962e);
    color: white;
    border-radius: 10px;
    font-weight: bold;
}

/* Sidebar fondo */
section[data-testid="stSidebar"] {
    background: linear-gradient(180deg,#1e3a5f,#14263f);
    color: white;
}

/* Botones Sidebar PRO */
section[data-testid="stSidebar"] .stButton>button {
    width: 100%;
    margin-bottom: 10px;
    padding: 12px;
    border-radius: 10px;
    font-weight: 600;
    background: linear-gradient(135deg, #d4af37, #b8962e);
    color: white;
    border: none;
    transition: all 0.2s ease-in-out;
}

section[data-testid="stSidebar"] .stButton>button:hover {
    transform: scale(1.03);
    background: linear-gradient(135deg, #e6c24c, #c9a227);
}

/* Caja info */
.info-box {
    background: linear-gradient(135deg,#fdfaf4,#fff);
    padding: 18px;
    border-radius: 10px;
    border-left: 5px solid #d4af37;
}

/* Radios */
div[role="radiogroup"] {
    background-color: #f1ede5;
    padding: 10px;
    border-radius: 8px;
}

</style>
""", unsafe_allow_html=True)
# =========================
# RESTO DE TU CÓDIGO
# (NO LO TOQUÉ)
# =========================

# Diccionario de frases
FRASES = {
    "Participacion": "Por haber Participado en el Curso Virtual de Apologética Católica Método Padre Luis Toro",
    "Reconocimiento": "Por haber aprobado satisfactoriamente el Curso Virtual de Apologética Católica, Método Padre Luis Toro"
}

# =========================
# 📊 LEER EXCEL LOCAL (BASE PRINCIPAL)
# =========================
def leer_excel_local():
    try:
        if not os.path.exists(RUTA_EXCEL):
            st.warning("⚠️ Archivo Excel no existe, creando base inicial...")

            df_base = pd.DataFrame(columns=[
                "Nombre_Completo",
                "Email",
                "Fecha_Curso",
                "PDF_Enviado",
                "Nivel",
                "Tipo",
                "Codigo_Verificacion"
            ])

            df_base.to_excel(RUTA_EXCEL, index=False)
            return df_base

        df = pd.read_excel(RUTA_EXCEL)

        return df

    except Exception as e:
        st.error(f"❌ Error leyendo Excel: {str(e)}")
        return None


def generar_pdf(nombre, frase, nivel, tipo, es_apologista=False):
    """Genera el PDF personalizado sobre la plantilla"""
    
    template_path = "plantillas/diploma_base.pdf"
    
    if not os.path.exists(template_path):
        st.error(f"❌ No se encuentra la plantilla: {template_path}")
        return None
    
    output_filename = f"generados/Diploma_{nombre.replace(' ', '_')}_N{nivel}.pdf"
    
    try:
        reader = PdfReader(template_path)
        writer = PdfWriter()
        
        page = reader.pages[0]
        width = float(page.mediabox.width)
        height = float(page.mediabox.height)
        
        packet = BytesIO()
        c = canvas.Canvas(packet, pagesize=(width, height))

        # =========================
        # 🔐 QR + CÓDIGO
        # =========================
        import uuid
        import qrcode

        codigo = str(uuid.uuid4())[:8].upper()
        url = f"https://TU-APP.streamlit.app/?verificar={codigo}"

        qr = qrcode.make(url)
        qr_path = f"generados/qr_{codigo}.png"
        qr.save(qr_path)

        # =========================
        # TEXTO DIPLOMA
        # =========================

        # =========================
        # ✨ TEXTO SUPERIOR DINÁMICO
        # =========================

        if tipo == "Participacion":
            texto_superior = "OTORGA EL PRESENTE CERTIFICADO DE PARTICIPACIÓN"

        elif tipo == "Reconocimiento":
            texto_superior = "OTORGA EL PRESENTE CERTIFICADO DE RECONOCIMIENTO"

        elif tipo == "Apologista":
            texto_superior = "OTORGA EL PRESENTE CERTIFICADO DE APOLOGISTA"

        else:
            texto_superior = "OTORGA EL PRESENTE CERTIFICADO"

        c.setFillColorRGB(0.83, 0.69, 0.22)  # dorado elegante
        c.setFont("Times-Bold", 16)

        text_width_top = c.stringWidth(texto_superior, "Times-Bold", 16)
        x_top = (width - text_width_top) / 2

        c.drawString(x_top, 355, texto_superior)

        c.setFillColorRGB(0, 0, 0)  # volver a negro

        # Nombre
        c.setFont("Times-BoldItalic", 20)
        nombre_texto = f"A: {nombre.upper()}"
        text_width = c.stringWidth(nombre_texto, "Times-BoldItalic", 19)
        x_nombre = (width - text_width) / 2
        c.drawString(x_nombre, 330, nombre_texto)

        # =========================
        # 🔧 FRASE DINÁMICA SEGÚN TIPO
        # =========================
        c.setFont("Times-Roman", 16)

        if tipo == "Apologista":
            frase_apologista = "POR HABER APROBADO SATISFACTORIAMENTE LOS 3 NIVELES DEL CURSO VIRTUAL DE APOLOGÉTICA CATÓLICA, MÉTODO PADRE LUIS TORO"
    
            palabras = frase_apologista.split()
        else:
            palabras = frase.split()

        mitad = len(palabras) // 2
        linea1 = " ".join(palabras[:mitad])
        linea2 = " ".join(palabras[mitad:])

        x1 = (width - c.stringWidth(linea1, "Times-Roman", 14)) / 2
        c.drawString(x1, 310, linea1)

        x2 = (width - c.stringWidth(linea2, "Times-Roman", 14)) / 2
        c.drawString(x2, 290, linea2)

        # =========================
        # 🔧 NIVEL (OCULTO PARA APOSTOLISTA)
        # =========================
        if tipo != "Apologista":
            c.setFont("Times-Bold", 16)
            nivel_texto = f"NIVEL {nivel}"
            x_nivel = (width - c.stringWidth(nivel_texto, "Times-Bold", 16)) / 2
            c.drawString(x_nivel, 270, nivel_texto)

        # =========================
        # 🏅 SELLO POR NIVEL
        # =========================

        if not es_apologista:
            if nivel == "1":
                sello_path = "assets/sello_nivel1.png"
            elif nivel == "2":
                sello_path = "assets/sello_nivel2.png"
            elif nivel == "3":
                sello_path = "assets/sello_nivel3.png"
            else:
                sello_path = None

            if sello_path and os.path.exists(sello_path):
                c.drawImage(sello_path, 60, 180, 90, 90, mask='auto')  # ← ajustamos luego si quieres

        # =========================
        # 📅 FECHA (TU VERSIÓN PERFECTA)
        # =========================
        from datetime import datetime

        meses = [
            "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
            "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
        ]

        hoy = datetime.now()
        fecha_texto = f"{hoy.day} de {meses[hoy.month - 1]} de {hoy.year}"

        c.setFont("Times-Bold", 14)

        text_width_fecha = c.stringWidth(fecha_texto, "Times-Bold", 14)
        margen_derecho = 80

        x_fecha = width - text_width_fecha - margen_derecho
        y_fecha = 400  # ← misma posición que ya te gustaba

        c.drawString(x_fecha, y_fecha, fecha_texto)

        # =========================
        # 📱 QR EN EL DIPLOMA
        # =========================
        # 🔥 NUEVA POSICIÓN PRO
        x_qr = width - 130   # ← mueve a la izquierda
        y_qr = 200           # ← sube el QR

        c.drawImage(qr_path, x_qr, y_qr, 80, 80)

        c.save()
        packet.seek(0)
        
        new_pdf = PdfReader(packet)
        page.merge_page(new_pdf.pages[0])
        writer.add_page(page)
        
        with open(output_filename, "wb") as output_file:
            writer.write(output_file)
        
        return output_filename, codigo
        
    except Exception as e:
        st.error(f"❌ Error generando PDF: {str(e)}")
        return None

# =========================
# 📧 ENVÍO CON OUTLOOK (HOTMAIL)
# =========================
def enviar_email_api(destinatario, archivo_pdf, nombre):

    import sib_api_v3_sdk
    from sib_api_v3_sdk.rest import ApiException
    from sib_api_v3_sdk import Configuration, ApiClient
    from sib_api_v3_sdk.api import transactional_emails_api
    from sib_api_v3_sdk.models import SendSmtpEmail, SendSmtpEmailAttachment
    import base64
    import os

    configuration = Configuration()
    configuration.api_key['api-key'] = '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'

    api_instance = transactional_emails_api.TransactionalEmailsApi(ApiClient(configuration))

    try:
        with open(archivo_pdf, "rb") as f:
            contenido = base64.b64encode(f.read()).decode()

        attachment = SendSmtpEmailAttachment(
            content=contenido,
            name=os.path.basename(archivo_pdf)
        )

        email = SendSmtpEmail(
            to=[{"email": destinatario, "name": nombre}],
            sender={"email": "erickmazierlg@gmail.com", "name": "Escuela Apologética"},
            subject="🎓 Tu Diploma Oficial",
            html_content=f"""
            <p>Hola {nombre},</p>
            <p>Adjunto encontrarás tu diploma.</p>
            <p>Dios te bendiga 🙏</p>
            """,
            attachment=[attachment]
        )

        api_instance.send_transac_email(email)

        st.success(f"📧 Enviado a {destinatario}")
        return True

    except ApiException as e:
        st.error(f"❌ Error API: {str(e)}")
        return False

# ============ INTERFAZ PRINCIPAL ============

st.markdown('<div class="main-header">⛪ Escuela de Apologética Católica</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">"Método Padre Luis Toro" - Sistema de Certificados</div>', unsafe_allow_html=True)

st.markdown("---")

# Selección de Nivel y Tipo
col1, col2 = st.columns(2)

with col1:
    st.subheader("📚 Seleccionar Nivel")
    nivel = st.radio(
        "Nivel del curso:",
        ["1", "2", "3"],
        horizontal=True,
        help="Selecciona el nivel del diplomado"
    )

with col2:
    st.subheader("🎯 Tipo de Certificado")
    if nivel == "3":
        tipo = st.radio(
            "Tipo:",
            ["Participacion", "Reconocimiento", "Apologista"],
            help="Nivel 3 tiene opción especial de Apologista"
        )
    else:
        tipo = st.radio(
            "Tipo:",
            ["Participacion", "Reconocimiento"],
            help="Selecciona si participó o aprobó"
        )

# Mostrar frase resultante
frase_base = FRASES["Participacion"] if tipo == "Participacion" else FRASES["Reconocimiento"]
frase_completa = f"{frase_base}"

st.markdown('<div class="info-box">', unsafe_allow_html=True)
st.markdown(f"**📝 Texto que aparecerá en el diploma:**")
st.markdown(f"*A: [NOMBRE DEL ALUMNO]*")
st.markdown(f"*{frase_completa}*")
st.markdown(f"***NIVEL {nivel}***")
if tipo == "Apologista":
    st.markdown(f"*(Certificado especial de Apologista)*")
st.markdown('</div>', unsafe_allow_html=True)

st.markdown("---")

# Botón para cargar datos
if st.button("📊 CARGAR ALUMNOS", type="secondary"):
    with st.spinner("Cargando base de datos..."):
        df = leer_excel_local()
        if df is not None:
            st.session_state['df'] = df
            st.session_state['nivel'] = nivel
            st.session_state['tipo'] = tipo
            st.success(f"✅ {len(df)} alumnos cargados correctamente")

# Mostrar tabla y procesar
if 'df' in st.session_state:
    df = st.session_state['df']
    
    st.subheader("📋 Lista de Alumnos")
    
    # 🔧 PROTECCIÓN SI EXCEL ESTÁ VACÍO
    if df is not None and not df.empty:

        # Mostrar tabla editable
        edited_df = st.data_editor(
            df[['Nombre_Completo', 'Email', 'Fecha_Curso', 'PDF_Enviado']],
            disabled=['PDF_Enviado'],
            hide_index=True,
            use_container_width=True
        )
        
        # Filtrar solo los no enviados
        pendientes = df[
            (df['PDF_Enviado'] != 'SÍ') &
            (df['Nivel'].astype(str) == str(nivel)) &
            (df['Tipo'] == tipo)
        ]
        
        st.markdown(f"**🎯 Alumnos pendientes: {len(pendientes)}**")
        
        col_gen, col_test = st.columns(2)
        
        with col_gen:
            if st.button("🚀 GENERAR TODOS LOS DIPLOMAS", type="primary"):
                if len(pendientes) == 0:
                    st.warning("No hay alumnos pendientes")
                else:
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    archivos_generados = []
                    
                    for i, (idx, row) in enumerate(pendientes.iterrows()):
                        # Actualizar progreso
                        progress = (i + 1) / len(pendientes)
                        progress_bar.progress(progress)
                        status_text.text(f"Generando: {row['Nombre_Completo']}...")
                        
                        # Generar PDF
                        es_apologista = (row['Tipo'] == "Apologista")
                        pdf_path, codigo = generar_pdf(
                            row['Nombre_Completo'],
                            frase_completa,
                            str(row['Nivel']),
                            row['Tipo'],
                            es_apologista
                        )
                        
                        if pdf_path:
                            df.loc[idx, 'Codigo_Verificacion'] = codigo

                            archivos_generados.append({
                                'nombre': row['Nombre_Completo'],
                                'email': row['Email'],
                                'archivo': pdf_path
                            })
                    
                    # Guardar en session para descarga
                    st.session_state['archivos_generados'] = archivos_generados
                    
                    st.success(f"🎉 ¡{len(archivos_generados)} diplomas generados!")
                    st.balloons()

                    df.to_excel(RUTA_EXCEL, index=False)
        
        with col_test:
            if st.button("🧪 GENERAR PRUEBA (1 solo)", type="secondary"):
                if len(pendientes) > 0:
                    primer_alumno = pendientes.iloc[0]
                    es_apologista = (st.session_state['tipo'] == "Apologista")
                    
                    pdf_path = generar_pdf(
                        primer_alumno['Nombre_Completo'],
                        frase_completa,
                        st.session_state['nivel'],
                        st.session_state['tipo'],
                        es_apologista
                    )
                    
                    if pdf_path:
                        st.success(f"✅ Prueba generada: {pdf_path}")
                        # Ofrecer descarga
                        with open(pdf_path, "rb") as f:
                            st.download_button(
                                "⬇️ Descargar prueba",
                                f,
                                file_name=f"Prueba_Diploma.pdf",
                                mime="application/pdf"
                            )

    else:
        st.warning("⚠️ El archivo Excel está vacío. Agrega alumnos para continuar.")

# =========================
# 📊 PANEL DE CONTROL
# =========================

df_control = leer_excel_local()

if df_control is not None and not df_control.empty:

    if 'PDF_Enviado' not in df_control.columns:
        df_control['PDF_Enviado'] = 'No'

    df_control['PDF_Enviado'] = df_control['PDF_Enviado'].fillna('No')
    df_control['PDF_Enviado'] = df_control['PDF_Enviado'].astype(str).str.strip()

    filtro_total = df_control[
        (df_control['Nivel'].astype(str) == str(nivel)) &
        (df_control['Tipo'] == tipo)
    ]

    filtro_enviados = filtro_total[
        filtro_total['PDF_Enviado'] == 'Sí'
    ]

    filtro_pendientes = filtro_total[
        filtro_total['PDF_Enviado'] != 'Sí'
    ]

    total = len(filtro_total)
    enviados = len(filtro_enviados)
    pendientes = len(filtro_pendientes)

    hoy = min(300, pendientes)

    st.markdown("### 📊 Estado del Lote")

    col1, col2, col3, col4 = st.columns(4)

    col1.metric("Total", total)
    col2.metric("Enviados", enviados)
    col3.metric("Pendientes", pendientes)
    col4.metric("Se envían hoy", hoy)

# =========================
# 📧 ENVÍO PROFESIONAL (MAX 300)
# =========================

st.markdown("---")
st.subheader("📧 Envío de Diplomas")

if st.button("📨 ENVIAR HASTA 300 PENDIENTES", type="primary"):

    df = leer_excel_local()

    # =========================
    # 🔧 NORMALIZAR EXCEL (PASO 5)
    # =========================
    if 'PDF_Enviado' not in df.columns:
        df['PDF_Enviado'] = 'No'

    df['PDF_Enviado'] = df['PDF_Enviado'].fillna('No')
    df['PDF_Enviado'] = df['PDF_Enviado'].astype(str).str.strip().str.upper()
    df['PDF_Enviado'] = df['PDF_Enviado'].replace({
        '': 'No',
        'nan': 'No',
        'None': 'No'
    })

    # =========================
    # 🎯 FILTRO POR NIVEL + TIPO
    # =========================
    pendientes = df[
        (df['PDF_Enviado'] != 'SÍ') &
        (df['Nivel'].astype(str) == str(nivel)) &
        (df['Tipo'] == tipo)
    ]

    total_pendientes = len(pendientes)
    limite = 300

    lote_envio = pendientes.head(limite)

    st.info(f"📦 Pendientes totales: {total_pendientes}")
    st.info(f"🚀 Se enviarán ahora: {len(lote_envio)}")

    enviados = 0
    errores = 0

    # =========================
    # 🚀 ENVÍO
    # =========================
    for idx, row in lote_envio.iterrows():

        nombre_archivo = row['Nombre_Completo'].strip().replace(' ', '_')   
        archivo_pdf = f"generados/Diploma_{nombre_archivo}_N{row['Nivel']}.pdf"

        if os.path.exists(archivo_pdf):

            ok = enviar_email_api(
                row['Email'],
                archivo_pdf,
                row['Nombre_Completo']
            )

            if ok:
                df.loc[idx, 'PDF_Enviado'] = 'Sí'
                enviados += 1
            else:
                errores += 1

        else:
            st.warning(f"⚠️ PDF no encontrado: {row['Nombre_Completo']}")
            errores += 1

    # =========================
    # 💾 GUARDAR EXCEL
    # =========================
    df.to_excel(RUTA_EXCEL, index=False)

    st.success(f"✅ Enviados correctamente: {enviados}")
    st.info("📌 Puedes volver mañana para continuar con el siguiente lote")

    if errores > 0:
        st.warning(f"⚠️ Errores: {errores}")

st.markdown("---")
st.caption("🏛️ Parroquia Santísima Trinidad - Diócesis de Acarigua-Araure - Venezuela")
st.caption("Desarrollado Por Erick Mazier con ❤️ para la Escuela de Apologética Católica Metodo Padre Luis Toro")