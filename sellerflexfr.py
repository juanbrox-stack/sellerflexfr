import streamlit as st
import pandas as pd
import re
import io
import os
import requests
from datetime import datetime

# Librerías para envío de email en segundo plano con adjuntos
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

st.set_page_config(page_title="Seller Flex FR Automator", layout="wide", page_icon="📦")

st.title("📦 Automatización Seller Flex FR - Cecopartners")
st.markdown("""
Esta aplicación descarga automáticamente el **Stock de Francia en tiempo real desde Cecotec Cloud**, comprueba la disponibilidad (filtrando referencias con stock ≤ 2) y segmenta la información para Cecopartners, cancelaciones por email directo y almacén.
""")

st.sidebar.header("Carga de Ficheros Principales")

# 1. Ya no se pide el archivo de Stock en la barra lateral. Se descarga automáticamente por URL.
st.sidebar.info("🌐 Stock de Francia: Se descarga automáticamente desde Cecotec Cloud en tiempo real.")

# 2. Segundo cargador (Pedidos de la lista de recogida) con su captura de pantalla correspondiente
st.sidebar.markdown("---")
pedidos_recoger_file = st.sidebar.file_uploader("1. Pedidos de la lista de recogida (Excel/CSV)", type=["csv", "xlsx"])
if os.path.exists("pedidoslistarecogida.png"):
    st.sidebar.image("pedidoslistarecogida.png", caption="Ayuda: Archivo de pedidos de la lista de recogida", use_container_width=True)

# 3. Tercer cargador (Fichero con ID pedidos) con su captura de pantalla correspondiente
st.sidebar.markdown("---")
listar_recogida_file = st.sidebar.file_uploader("2. Fichero con ID pedidos (Excel/CSV)", type=["csv", "xlsx"])
if os.path.exists("idpedidos.png"):
    st.sidebar.image("idpedidos.png", caption="Ayuda: Archivo con IDs de pedido y envío", use_container_width=True)


def download_stock_from_url():
    """Descarga el stock en tiempo real desde la URL de Cecotec Cloud manejando separadores y encodings."""
    url = "https://cecobi.cecotec.cloud/ws/getstocksabanaFranciaX.php"
    try:
        response = requests.get(url, timeout=30)
        response.raise_for_status()  # Lanza un error si la descarga falla (HTTP 4xx o 5xx)
        
        # El archivo devuelto es un CSV. Probamos codificaciones típicas europeas
        csv_data = response.content
        for enc in ['utf-8', 'latin-1', 'iso-8859-1']:
            try:
                # Intentamos leer primero con punto y coma (común en Cecotec)
                df = pd.read_csv(io.BytesIO(csv_data), sep=';', encoding=enc, on_bad_lines='skip')
                if len(df.columns) > 1:
                    return df
            except Exception:
                try:
                    # Fallback a coma tradicional si falla
                    df = pd.read_csv(io.BytesIO(csv_data), sep=',', encoding=enc, on_bad_lines='skip')
                    if len(df.columns) > 1:
                        return df
                except Exception:
                    continue
        return None
    except Exception as e:
        st.error(f"❌ Error crítico al conectar o descargar el Stock desde Cecotec Cloud: {e}")
        return None

def load_data(file):
    """Carga los ficheros subidos por el usuario (Excel o CSV)."""
    if file is not None:
        if file.name.endswith('.csv') or file.name.endswith('.txt'):
            try:
                return pd.read_csv(file, sep=None, engine='python')
            except Exception:
                file.seek(0)
                return pd.read_csv(file, encoding='latin-1', sep=None, engine='python')
        else:
            return pd.read_excel(file)
    return None

def clean_sku(sku):
    """Limpia el SKU quitando el prefijo 'FR' y los ceros sobrantes a la izquierda."""
    if pd.isna(sku):
        return ""
    sku_str = str(sku).strip()
    if sku_str.upper().startswith("FR"):
        sku_str = sku_str[2:]
    sku_str = sku_str.lstrip('0')
    return sku_str

def to_excel(df):
    """Convierte un DataFrame en un archivo de Excel descargable en memoria."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Datos')
    return output.getvalue()


def enviar_correo_background(excel_data, filename):
    """Envía el correo directamente desde el servidor SMTP usando los Secrets de Streamlit."""
    try:
        if "email" not in st.secrets:
            st.error("❌ Error: No se ha encontrado la sección [email] en los Secrets de Streamlit.")
            return False
            
        smtp_server = st.secrets["email"]["smtp_server"]
        smtp_port = st.secrets["email"]["smtp_port"]
        sender_email = st.secrets["email"]["sender_email"]
        sender_password = st.secrets["email"]["sender_password"]
        
        destinatario = "juanbrox@cecotec.es"
        fecha_hoy = datetime.now().strftime("%d/%m/%Y")
        asunto = f"Cancel orders {fecha_hoy}"
        cuerpo_mensaje = "Good morning,\n\nI attach one order to cancel.\n\nBest regards."
        
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = destinatario
        msg['Subject'] = asunto
        msg.attach(MIMEText(cuerpo_mensaje, 'plain'))
        
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(excel_data)
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename="{filename}"')
        msg.attach(part)
        
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, destinatario, msg.as_string())
        server.quit()
        return True
    except Exception as e:
        st.error(f"❌ Error al enviar el correo a través del servidor: {e}")
        return False


# Verificar que los 2 ficheros manuales requeridos han sido subidos
if pedidos_recoger_file and listar_recogida_file:
    
    st.info("📥 Conectando con Cecotec Cloud y procesando ficheros... Por favor, espera.")
    
    # 1. Descarga automática del Stock desde la URL
    df_stock = download_stock_from_url()
    
    # 2. Carga de los DataFrames manuales
    df_pedidos_recoger = load_data(pedidos_recoger_file)
    df_listar_recogida = load_data(listar_recogida_file)
    
    # Estructura fija en memoria de la plantilla de Cecopartners
    columnas_plantilla = [
        'article', 'quantity', 'customer_name', 'nif', 'attention_of_customer', 
        'address', 'postal_code', 'phone', 'city', 'country_code', 
        'customer_mail', 'comment', 'addressee_order_number'
    ]
    
    if df_stock is None or df_stock.empty:
        st.error("No se pudo obtener o procesar el archivo de Stock automático desde la URL. Por favor, revisa la conexión o el servidor Cecobi.")
    else:
        try:
            # ---- PROCESAMIENTO 1: Parsear el Fichero de la Opción 2 (ID Pedidos) ----
            col_listar_a = df_listar_recogida.columns[0] 
            
            def parse_listar_recogida(text):
                text = str(text).strip()
                n_pedido = text.split()[0] if len(text.split()) > 0 else ""
                match = re.search(r'(?:ID de envío:|ID de envio:)\s*([A-Za-z0-9]+)', text)
                id_envio = match.group(1) if match else ""
                return pd.Series([n_pedido, id_envio])

            df_listar_recogida[['Num_Pedido_LR', 'Id_Envio_LR']] = df_listar_recogida[col_listar_a].apply(parse_listar_recogida)
            mapa_envio_pedido = dict(zip(df_listar_recogida['Id_Envio_LR'].str.strip(), df_listar_recogida['Num_Pedido_LR'].str.strip()))

            # ---- PROCESAMIENTO 2: Limpieza y Mapeo en la Opción 1 (Pedidos de Lista de Recogida) ----
            df_pedidos_recoger['SKU_Limpio'] = df_pedidos_recoger['SKU'].apply(clean_sku)
            df_pedidos_recoger['Identificador_Clean'] = df_pedidos_recoger['Identificador de pedido'].astype(str).str.strip()
            df_pedidos_recoger['Número_Pedido_Final'] = df_pedidos_recoger['Identificador_Clean'].map(mapa_envio_pedido).fillna("")
            
            # Mapeos auxiliares de Zona e Identificador de Envío vinculados al Número de Pedido de Amazon
            mapa_pedido_a_zona = dict(zip(df_pedidos_recoger['Número_Pedido_Final'].str.strip(), df_pedidos_recoger['Zona'].astype(str).str.strip()))
            mapa_pedido_a_envio = dict(zip(df_pedidos_recoger['Número_Pedido_Final'].str.strip(), df_pedidos_recoger['Identificador de pedido'].astype(str).str.strip()))

            # ---- PROCESAMIENTO 3: Mapeo del Stock Descargado de la URL ----
            col_stock_ref = df_stock.columns[0]
            
            col_stock_cant = None
            for col in df_stock.columns:
                if any(x in col.lower() for x in ['disponible', 'physique', 'stock', 'cantidad', 'unidades']):
                    col_stock_cant = col
                    break
            if not col_stock_cant:
                col_stock_cant = df_stock.columns[1]
            
            df_stock[col_stock_ref] = df_stock[col_stock_ref].astype(str).str.strip().apply(lambda x: x.lstrip('0'))
            df_stock[col_stock_cant] = pd.to_numeric(df_stock[col_stock_cant], errors='coerce').fillna(0)
            
            mapa_referencias_stock = dict(zip(df_stock[col_stock_ref], df_stock[col_stock_cant]))
            df_pedidos_recoger['Stock_Actual'] = df_pedidos_recoger['SKU_Limpio'].map(mapa_referencias_stock).fillna(0)
            
            # Regla de corte automatizada: stock > 2 disponible. Si stock <= 2, se va a cancelaciones.
            df_pedidos_recoger['Disponible'] = df_pedidos_recoger['Stock_Actual'] > 2
            
            # ---- FILTRADO Y DIVISIÓN ----
            df_ok = df_pedidos_recoger[df_pedidos_recoger['Disponible'] == True].copy()
            df_cancel = df_pedidos_recoger[df_pedidos_recoger['Disponible'] == False].copy()
            
            # ---- CONSTRUCCIÓN DE LOS FICHEROS DE SALIDA ----
            
            # 1. FICHERO FINAL CECOPARTNERS
            df_subida_plantilla = pd.DataFrame(columns=columnas_plantilla)
            if not df_ok.empty:
                df_subida_plantilla['article'] = df_ok['SKU_Limpio']
                df_subida_plantilla['quantity'] = df_ok['Unidades'].fillna(1).astype(int)
                df_subida_plantilla['customer_name'] = 'AMAZON FLEX'
                df_subida_plantilla['nif'] = ''
                df_subida_plantilla['attention_of_customer'] = 'AMAZON FLEX'
                df_subida_plantilla['address'] = 'Cam.Real de Madrid 117'
                df_subida_plantilla['postal_code'] = 46292
                df_subida_plantilla['phone'] = 0
                df_subida_plantilla['city'] = 'MASSALAVÉS'
                df_subida_plantilla['country_code'] = 'ES'
                df_subida_plantilla['comment'] = 0
                df_subida_plantilla['addressee_order_number'] = df_ok['Número_Pedido_Final']
                df_subida_plantilla['customer_mail'] = df_subida_plantilla['addressee_order_number'].astype(str) + '@sellerflexfr.com'
            
            # 2. FICHERO DE CANCELACIONES
            if not df_cancel.empty:
                df_cancelaciones = pd.DataFrame({
                    'Node ID': 'SRAN',
                    'Order number': df_cancel['Número_Pedido_Final'],
                    'Shipment ID': df_cancel['Identificador de pedido'],
                    'ASIN': df_cancel['FNSKU'],
                    'Reason': 'OOO'
                })
            else:
                df_cancelaciones = pd.DataFrame(columns=['Node ID', 'Order number', 'Shipment ID', 'ASIN', 'Reason'])
                
            # ---- RENDERIZADO EN STREAMLIT ----
            st.success("✨ ¡Ficheros base calculados con éxito! Stock descargado automáticamente desde la nube de Cecotec.")
            
            pestana1, pestana2, pestana3 = st.tabs(["📤 Fichero Subida (Plantilla Cecopartners)", "❌ Cancelaciones (OOO)", "🇫🇷 D-PEDIDOS Francia"])
            
            with pestana1:
                st.subheader("Fichero Resultante Cecopartners")
                st.dataframe(df_subida_plantilla)
                st.download_button(
                    label="📥 Descargar Fichero Subida Cecopartners (Excel)",
                    data=to_excel(df_subida_plantilla),
                    file_name="SELLER_FLEX_FR_PROCESADO.xlsx",
                    mime="application/vnd.ms-excel"
                )
                
            with pestana2:
                st.subheader("Fichero de Cancelaciones")
                st.dataframe(df_cancelaciones)
                
                if not df_cancelaciones.empty:
                    excel_cancelados = to_excel(df_cancelaciones)
                    nombre_archivo_cancelados = "CancelOrders_SellerFlexFR.xlsx"
                    
                    st.download_button(
                        label="📥 Descargar Fichero Cancelaciones (Excel)",
                        data=excel_cancelados,
                        file_name=nombre_archivo_cancelados,
                        mime="application/vnd.ms-excel"
                    )
                    
                    st.markdown("---")
                    st.subheader("📧 Servidor Automatizado de Correo")
                    st.write("Presiona el siguiente botón para enviar directamente el fichero de cancelaciones a **fr-sellerflex-support@amazon.com** utilizando los recursos del servidor:")
                    
                    if st.button("🚀 Enviar Fichero de Cancelación Directamente", type="primary"):
                        with st.spinner("Conectando con el servidor SMTP y enviando correo..."):
                            exito = enviar_correo_background(excel_cancelados, nombre_archivo_cancelados)
                            if exito:
                                st.success("📬 ¡Correo enviado con éxito! El archivo Excel de cancelaciones ha sido adjuntado de manera automática.")
                else:
                    st.info("No se han detectado pedidos para cancelar.")
                
            with pestana3:
                st.subheader("Generación de Fichero Definitivo de Almacén")
                st.markdown("""
                **Paso intermedio:** Descarga primero el archivo de Cecopartners de la pestaña 1, súbelo a su plataforma, y cuando te devuelvan el **fichero con las referencias D**, cárgalo aquí abajo para estructurar el definitivo de almacén:
                """)
                
                cecopartners_downloaded_file = st.file_uploader("Subir Fichero Descargado de Cecopartners (Excel/CSV) para cruzar las D", type=["csv", "xlsx"], key="ceco_almacen")
                
                if cecopartners_downloaded_file is not None:
                    df_ceco_in = load_data(cecopartners_downloaded_file)
                    
                    if df_ceco_in is not None and not df_ceco_in.empty:
                        try:
                            dict_cols_ceco = {col.lower(): col for col in df_ceco_in.columns}
                            col_ceco_ref_d = dict_cols_ceco.get('referencia', df_ceco_in.columns[3] if len(df_ceco_in.columns) > 3 else df_ceco_in.columns[0])
                            col_ceco_pedido = dict_cols_ceco.get('número de pedido de cliente', dict_cols_ceco.get('addressee_order_number', df_ceco_in.columns[-1]))
                            
                            df_almacen_fr = pd.DataFrame()
                            pedidos_ceco_limpios = df_ceco_in[col_ceco_pedido].astype(str).str.strip()
                            
                            df_almacen_fr['P'] = pedidos_ceco_limpios.map(mapa_pedido_a_zona).fillna("")
                            df_almacen_fr['U'] = pedidos_ceco_limpios.map(mapa_pedido_a_envio).fillna("")
                            df_almacen_fr['Agencia'] = 'AMZN_FR_SH_SD'
                            df_almacen_fr['REFERENCIA'] = df_ceco_in[col_ceco_ref_d].astype(str).str.strip()
                            df_almacen_fr['NOMBRE DEL CLIENTE'] = 'AMAZON FLEX'
                            df_almacen_fr['ESTADO'] = 'ESPERANDO ETIQUETA'
                            df_almacen_fr['NÚMERO DE PEDIDO DE CLIENTE'] = pedidos_ceco_limpios
                            
                            st.success("🎉 ¡Fichero definitivo para Almacén Francia generado con las 'D' cruzadas exitosamente!")
                            st.dataframe(df_almacen_fr)
                            
                            st.download_button(
                                label="📥 Descargar D-PEDIDOS Almacén Definitivo (Excel)",
                                data=to_excel(df_almacen_fr),
                                file_name="D-PEDIDOS_FLEX_FR_DEFINITIVO.xlsx",
                                mime="application/vnd.ms-excel"
                            )
                        except Exception as ex_cruce:
                            st.error(f"Error procesando el fichero devuelto de Cecopartners: {ex_cruce}")
                    else:
                        st.warning("El archivo de Cecopartners subido está vacío o no es válido.")
                else:
                    st.info("⏳ Esperando que subas el archivo descargado de Cecopartners con las referencias D para generar el listado definitivo de almacén.")

        except Exception as e:
            st.error(f"Error estructural en las columnas de los ficheros: {e}")
            st.warning("Asegúrate de que las columnas coincidan con las estructuras estándar.")
else:
    st.info("👋 Por favor, carga los 2 archivos manuales requeridos en la barra lateral para empezar a operar.")
