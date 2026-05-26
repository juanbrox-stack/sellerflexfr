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
Esta aplicación procesa los ficheros de Seller Flex, comprueba la disponibilidad en el Stock de Francia (filtrando referencias con stock ≤ 2) y segmenta la información para Cecopartners, cancelaciones por email directo y almacén.
""")

st.sidebar.header("Carga de Ficheros Principales")

# Opción híbrida de Stock (Automatizado por URL o Manual por archivo)
st.sidebar.subheader("⚙️ Configuración de Stock FR")
subida_manual_stock = st.sidebar.checkbox("Subir fichero de Stock FR manualmente", value=False)

stock_file = None
if subida_manual_stock:
    stock_file = st.sidebar.file_uploader("1. Selecciona el Fichero de Stock FR (CSV)", type=["csv", "txt"])
else:
    st.sidebar.info("🌐 Stock de Francia: Configurado en automático desde Cecotec Cloud en tiempo real.")

# 2. Segundo cargador (Pedidos de la lista de recogida)
st.sidebar.markdown("---")
pedidos_recoger_file = st.sidebar.file_uploader("2. Pedidos de la lista de recogida (Excel/CSV)", type=["csv", "xlsx"])
if os.path.exists("pedidoslistarecogida.png"):
    st.sidebar.image("pedidoslistarecogida.png", caption="Ayuda: Archivo de pedidos de la lista de recogida", use_container_width=True)

# 3. Tercer cargador (Fichero con ID pedidos)
st.sidebar.markdown("---")
listar_recogida_file = st.sidebar.file_uploader("3. Fichero con ID pedidos (Excel/CSV)", type=["csv", "xlsx"])
if os.path.exists("idpedidos.png"):
    st.sidebar.image("idpedidos.png", caption="Ayuda: Archivo con IDs de pedido y envío", use_container_width=True)


def download_stock_from_url():
    """Descarga el stock en tiempo real desde la URL de Cecotec Cloud."""
    url = "https://cecobi.cecotec.cloud/ws/getstocksabanaFranciaX.php"
    try:
        response = requests.get(url, timeout=30)
        response.raise_for_status()
        return process_csv_bytes(response.content)
    except Exception as e:
        st.error(f"❌ Error crítico al conectar o descargar el Stock automático desde Cecotec Cloud: {e}")
        return None

def process_csv_bytes(csv_bytes):
    """Procesa un bloque de bytes de CSV manejando encodings y delimitadores europeos."""
    for enc in ['utf-8', 'latin-1', 'iso-8859-1']:
        try:
            df = pd.read_csv(io.BytesIO(csv_bytes), sep=';', encoding=enc, on_bad_lines='skip')
            if len(df.columns) > 1:
                return df
        except Exception:
            try:
                df = pd.read_csv(io.BytesIO(csv_bytes), sep=',', encoding=enc, on_bad_lines='skip')
                if len(df.columns) > 1:
                    return df
            except Exception:
                continue
    return None

def load_data(file):
    """Carga los ficheros subidos de forma manual por el usuario (Excel o CSV)."""
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


def enviar_correo_background(excel_data, filename, extra_file=None):
    """Envía el correo desde el servidor SMTP incluyendo destinatarios en CC y el segundo archivo adjunto diario."""
    try:
        if "email" not in st.secrets:
            st.error("❌ Error: No se ha encontrado la sección [email] en los Secrets de Streamlit.")
            return False
            
        smtp_server = st.secrets["email"]["smtp_server"]
        smtp_port = st.secrets["email"]["smtp_port"]
        sender_email = st.secrets["email"]["sender_email"]
        sender_password = st.secrets["email"]["sender_password"]
        
        destinatario = "fr-sellerflex-support@amazon.com"
        cc_emails = ["juanbrox@cecotec.es", "antoniodiaz@cecotec.es"]
        
        fecha_hoy = datetime.now().strftime("%d/%m/%Y")
        asunto = f"Cancel orders {fecha_hoy}"
        cuerpo_mensaje = "Good morning,\n\nI attach one order to cancel.\n\nBest regards."
        
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = destinatario
        msg['Cc'] = ", ".join(cc_emails)
        msg['Subject'] = asunto
        msg.attach(MIMEText(cuerpo_mensaje, 'plain'))
        
        # Adjunto 1: Excel de cancelaciones calculadas
        part1 = MIMEBase('application', 'octet-stream')
        part1.set_payload(excel_data)
        encoders.encode_base64(part1)
        part1.add_header('Content-Disposition', f'attachment; filename="{filename}"')
        msg.attach(part1)
        
        # Adjunto 2: El pantallazo/fichero diario variable de transportistas que subes en la pestaña 3
        if extra_file is not None:
            part2 = MIMEBase('application', 'octet-stream')
            part2.set_payload(extra_file.read())
            encoders.encode_base64(part2)
            part2.add_header('Content-Disposition', f'attachment; filename="{extra_file.name}"')
            msg.attach(part2)
        
        todos_los_destinatarios = [destinatario] + cc_emails
        
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, todos_los_destinatarios, msg.as_string())
        server.quit()
        return True
    except Exception as e:
        st.error(f"❌ Error al enviar el correo a través del servidor: {e}")
        return False


# Control condicional de la carga de ficheros obligatorios
ficheros_listos = False
if subida_manual_stock:
    if stock_file and pedidos_recoger_file and listar_recogida_file:
        ficheros_listos = True
else:
    if pedidos_recoger_file and listar_recogida_file:
        ficheros_listos = True

if ficheros_listos:
    st.info("Procesando datos... Por favor, espera.")
    
    # 1. Obtención del DataFrame de Stock
    if subida_manual_stock:
        df_stock = process_csv_bytes(stock_file.read())
    else:
        df_stock = download_stock_from_url()
        
    # 2. Carga del resto de ficheros manuales
    df_pedidos_recoger = load_data(pedidos_recoger_file)
    df_listar_recogida = load_data(listar_recogida_file)
    
    # Estructura fija de Cecopartners
    columnas_plantilla = [
        'article', 'quantity', 'customer_name', 'nif', 'attention_of_customer', 
        'address', 'postal_code', 'phone', 'city', 'country_code', 
        'customer_mail', 'comment', 'addressee_order_number'
    ]
    
    if df_stock is None or df_stock.empty:
        st.error("No se pudo obtener o procesar el Stock. Si estás usando la opción manual, verifica que el CSV sea válido.")
    else:
        try:
            # ---- PROCESAMIENTO 1: Separar la columna A de ListarRecogida usando espacios ----
            col_listar_a = df_listar_recogida.columns[0] 
            
            def parse_listar_recogida(text):
                text = str(text).strip()
                n_pedido = text.split()[0] if len(text.split()) > 0 else ""
                match = re.search(r'(?:ID de envío:|ID de envio:)\s*([A-Za-z0-9]+)', text)
                id_envio = match.group(1) if match else ""
                return pd.Series([n_pedido, id_envio])

            df_listar_recogida[['Num_Pedido_LR', 'Id_Envio_LR']] = df_listar_recogida[col_listar_a].apply(parse_listar_recogida)
            mapa_envio_pedido = dict(zip(df_listar_recogida['Id_Envio_LR'].str.strip(), df_listar_recogida['Num_Pedido_LR'].str.strip()))

            # ---- PROCESAMIENTO 2: Limpieza y Mapeo en Pedidos de Lista de Recogida ----
            df_pedidos_recoger['SKU_Limpio'] = df_pedidos_recoger['SKU'].apply(clean_sku)
            df_pedidos_recoger['Identificador_Clean'] = df_pedidos_recoger['Identificador de pedido'].astype(str).str.strip()
            df_pedidos_recoger['Número_Pedido_Final'] = df_pedidos_recoger['Identificador_Clean'].map(mapa_envio_pedido).fillna("")
            
            # MAPAS DE MEMORIA GLOBALES INDEXADOS POR EL NÚMERO DE PEDIDO LARGO (Clave común con Cecopartners)
            mapa_pedido_largo_a_zona = dict(zip(df_pedidos_recoger['Número_Pedido_Final'].astype(str).str.strip(), df_pedidos_recoger['Zona'].astype(str).str.strip()))
            mapa_pedido_largo_a_envio = dict(zip(df_pedidos_recoger['Número_Pedido_Final'].astype(str).str.strip(), df_pedidos_recoger['Identificador de pedido'].astype(str).str.strip()))

            # ---- PROCESAMIENTO 3: Mapeo de Unidades de Stock ----
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
            
            # Regla de corte: stock > 2 disponible. Si es <= 2, se va a cancelaciones.
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
            st.success("✨ ¡Cálculos base realizados correctamente!")
            
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
                    st.write("Presiona el siguiente botón para enviar directamente el fichero de cancelaciones a **fr-sellerflex-support@amazon.com** con copia a Juan y Antonio:")
                    
                    if st.button("🚀 Enviar Fichero de Cancelación Directamente", type="primary"):
                        with st.spinner("Conectando con el servidor SMTP y enviando correo..."):
                            exito = enviar_correo_background(excel_cancelados, nombre_archivo_cancelados, None)
                            if exito:
                                st.success("📬 ¡Correo enviado con éxito! El soporte de Amazon ha recibido el fichero.")
                else:
                    st.info("No se han detectado pedidos para cancelar.")
                
            with pestana3:
                st.subheader("Generación de Fichero Definitivo de Almacén")
                st.markdown("""
                **Paso intermedio:** Descarga primero el archivo de Cecopartners de la pestaña 1, súbelo a su plataforma, y cuando te devuelvan el **fichero con las referencias D**, cárgalo aquí abajo para estructurar el definitivo de almacén:
                """)
                
                # Cargador del fichero devuelto de Cecopartners con las D
                cecopartners_downloaded_file = st.file_uploader("Subir Fichero Descargado de Cecopartners (Detalle de Transportistas) para añadir las D", type=["csv", "xlsx"], key="ceco_almacen")
                
                # Cargador del pantallazo/fichero diario variable solicitado por el usuario
                st.markdown("---")
                fichero_nuevo_diario = st.file_uploader("Adjuntar el nuevo fichero diario o pantallazo de transportistas", type=["png", "jpg", "jpeg", "pdf", "xlsx", "csv"], key="fichero_diario_transp")
                
                # Botón integrado para enviar el correo incluyendo el pantallazo diario si el usuario lo necesita
                if fichero_nuevo_diario is not None:
                    st.info(f"📂 Archivo diario cargado listo para ser enviado: **{fichero_nuevo_diario.name}**")
                    if st.button("📧 Enviar Correo con Pantallazo Diario adjunto", key="send_extra_mail"):
                        with st.spinner("Enviando correo con el archivo diario adjunto..."):
                            # Si hay cancelaciones acumuladas las mandamos, si no mandamos un buffer vacío junto con el pantallazo
                            excel_to_send = to_excel(df_cancelaciones) if not df_cancelaciones.empty else to_excel(pd.DataFrame([{"Info": "No cancel orders today"}]))
                            exito_mail = enviar_correo_background(excel_to_send, "CancelOrders.xlsx", fichero_nuevo_diario)
                            if exito_mail:
                                st.success("📬 ¡Correo con el Pantallazo Diario enviado con éxito a soporte!")

                if cecopartners_downloaded_file is not None:
                    df_ceco_in = load_data(cecopartners_downloaded_file)
                    
                    if df_ceco_in is not None and not df_ceco_in.empty:
                        try:
                            # Sincronización robusta de columnas sin importar mayúsculas
                            dict_cols_ceco = {col.lower().strip(): col for col in df_ceco_in.columns}
                            
                            col_ceco_ref_d = dict_cols_ceco.get('referencia', df_ceco_in.columns[3] if len(df_ceco_in.columns) > 3 else df_ceco_in.columns[0])
                            
                            # Identificamos la columna del número de pedido (ej: addressee_order_number o número de pedido de cliente)
                            col_ceco_pedido = dict_cols_ceco.get('addressee_order_number', dict_cols_ceco.get('número de pedido de cliente', None))
                            if not col_ceco_pedido:
                                # Fallback dinámico si no encuentra los nombres estándar
                                for c in df_ceco_in.columns:
                                    if any(x in c.lower() for x in ['pedido', 'number', 'addressee']):
                                        col_ceco_pedido = c
                                        break
                            if not col_ceco_pedido:
                                col_ceco_pedido = df_ceco_in.columns[-1]
                            
                            # Construir el fichero definitivo alineando con la estructura del almacén
                            df_almacen_fr = pd.DataFrame()
                            pedidos_ceco_limpios = df_ceco_in[col_ceco_pedido].astype(str).str.strip()
                            
                            # ¡Mapeo Corregido e Infalible! Cruzamos por el número de pedido largo de Amazon
                            df_almacen_fr['P'] = pedidos_ceco_limpios.map(mapa_pedido_largo_a_zona).fillna("")
                            df_almacen_fr['U'] = pedidos_ceco_limpios.map(mapa_pedido_largo_a_envio).fillna("")
                            df_almacen_fr['Agencia'] = 'AMZN_FR_SH_SD'
                            df_almacen_fr['REFERENCIA'] = df_ceco_in[col_ceco_ref_d].astype(str).str.strip()
                            df_almacen_fr['NOMBRE DEL CLIENTE'] = 'AMAZON FLEX'
                            df_almacen_fr['ESTADO'] = 'ESPERANDO ETIQUETA'
                            df_almacen_fr['NÚMERO DE PEDIDO DE CLIENTE'] = pedidos_ceco_limpios
                            
                            st.success("🎉 ¡Fichero definitivo para Almacén Francia generado con las 'D', 'P' y 'U' cruzadas exitosamente!")
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
                    st.info("⏳ Esperando que subas el archivo con las referencias D de Cecopartners para mapear el listado definitivo de almacén.")

        except Exception as e:
            st.error(f"Error estructural en las columnas de los ficheros: {e}")
            st.warning("Asegúrate de que las columnas coincidan con las estructuras estándar.")
else:
    st.sidebar.warning("Falta subir archivos obligatorios.")
    st.info("👋 Por favor, carga los archivos requeridos en la barra lateral para empezar a operar.")
