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
Esta aplicación procesa los ficheros de Seller Flex, comprueba la disponibilidad en el Stock de Francia (filtrando referencias con stock disponible ≤ 2) y segmenta la información para Cecopartners, cancelaciones por email directo y almacén.
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


def enviar_correo_background(excel_data, filename, destinatario, extra_file=None):
    """Envía el correo desde el servidor SMTP incluyendo destinatarios dinámicos y CC asignados."""
    try:
        if "email" not in st.secrets:
            st.error("❌ Error: No se ha encontrado la sección [email] en los Secrets de Streamlit.")
            return False
            
        smtp_server = st.secrets["email"]["smtp_server"]
        smtp_port = st.secrets["email"]["smtp_port"]
        sender_email = st.secrets["email"]["sender_email"]
        sender_password = st.secrets["email"]["sender_password"]
        
        cc_emails = ["juanbrox@cecotec.es", "antoniodiaz@cecotec.es"]
        
        fecha_hoy = datetime.now().strftime("%d/%m/%Y")
        
        if "support" in destinatario:
            asunto = f"Cancel orders {fecha_hoy}"
            cuerpo_mensaje = "Good morning,\n\nI attach one order to cancel.\n\nBest regards."
        else:
            asunto = f"SELLER FLEX FR - {fecha_hoy}"
            cuerpo_mensaje = f"Buenos días,\n\nAdjunto el archivo definitivo de pedidos procesado para el almacén junto al pantallazo diario de transportistas.\n\nSaludos cordiales."
        
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = destinatario
        msg['Cc'] = ", ".join(cc_emails)
        msg['Subject'] = asunto
        msg.attach(MIMEText(cuerpo_mensaje, 'plain'))
        
        # Adjunto 1: Fichero Excel Calculado
        part1 = MIMEBase('application', 'octet-stream')
        part1.set_payload(excel_data)
        encoders.encode_base64(part1)
        part1.add_header('Content-Disposition', f'attachment; filename="{filename}"')
        msg.attach(part1)
        
        # Adjunto 2: Pantallazo o archivo diario adicional si el usuario lo sube
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


# Generación de variables de fecha dinámicas
fecha_aaaammdd = datetime.now().strftime("%Y%m%d")

# Control condicional de la carga de ficheros obligatorios en barra lateral
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
    
    # Estructura fija de Cecopartners para la pestaña 1
    columnas_plantilla = [
        'article', 'quantity', 'customer_name', 'nif', 'attention_of_customer', 
        'address', 'postal_code', 'phone', 'city', 'country_code', 
        'customer_mail', 'comment', 'addressee_order_number'
    ]
    
    if df_stock is None or df_stock.empty:
        st.error("No se pudo obtener o procesar el Stock. Si estás usando la opción manual, verifica que el CSV sea válido.")
    else:
        try:
            # ---- PROCESAMIENTO 1: Extracción limpia por espacios de la columna A de ListarRecogida ----
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
            
            # PROTECCIÓN DE REPETIDOS (Fichero Recogida): Agrupamos en listas por pedido para consumirlos uno a uno de forma secuencial
            mapa_pedido_largo_a_zonas_list = df_pedidos_recoger.groupby('Número_Pedido_Final')['Zona'].apply(lambda x: x.astype(str).str.strip().tolist()).to_dict()
            
            dict_cols_pr = {col.lower().strip(): col for col in df_pedidos_recoger.columns}
            col_agencia_origen = dict_cols_pr.get('agencia', None)
            if col_agencia_origen:
                mapa_pedido_largo_a_agencias_list = df_pedidos_recoger.groupby('Número_Pedido_Final')[col_agencia_origen].apply(lambda x: x.astype(str).str.strip().tolist()).to_dict()
            else:
                mapa_pedido_largo_a_agencias_list = {}

            # ---- PROCESAMIENTO 3: Mapeo de Unidades de Stock CORREGIDO ----
            col_stock_ref = df_stock.columns[0]
            
            # AJUSTE FIJO: Lectura estricta sobre la columna F 'StockDisponible'
            col_stock_cant = None
            for col in df_stock.columns:
                if 'disponible' in col.lower() or 'stockdisponible' in col.lower():
                    col_stock_cant = col
                    break
            
            if not col_stock_cant:
                col_stock_cant = df_stock.columns[5] if len(df_stock.columns) > 5 else df_stock.columns[1]
            
            df_stock[col_stock_ref] = df_stock[col_stock_ref].astype(str).str.strip().apply(lambda x: x.lstrip('0'))
            df_stock[col_stock_cant] = pd.to_numeric(df_stock[col_stock_cant], errors='coerce').fillna(0)
            
            mapa_referencias_stock = dict(zip(df_stock[col_stock_ref], df_stock[col_stock_cant]))
            df_pedidos_recoger['Stock_Actual'] = df_pedidos_recoger['SKU_Limpio'].map(mapa_referencias_stock).fillna(0)
            
            # Regla de corte: stock disponible > 2 disponible. Si es <= 2, va a cancelaciones.
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
                df_subida_plantilla['addressee_order_number'] = df_ok['Número_Pedido_Final'].astype(str).str.strip()
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
            st.success("✨ ¡Cálculos base realizados correctamente bajo la columna StockDisponible!")
            
            pestana1, pestana2, pestana3 = st.tabs(["📤 Fichero Subida (Plantilla Cecopartners)", "❌ Cancelaciones (OOO)", "🇫🇷 D-PEDIDOS Francia"])
            
            with pestana1:
                st.subheader("Fichero Resultante Cecopartners")
                st.dataframe(df_subida_plantilla)
                st.download_button(
                    label="📥 Descargar Fichero Subida Cecopartners (Excel)",
                    data=to_excel(df_subida_plantilla),
                    file_name=f"{fecha_aaaammdd}_SELLER_FLEX_FR_PROCESADO.xlsx",
                    mime="application/vnd.ms-excel"
                )
                
            with pestana2:
                st.subheader("Fichero de Cancelaciones")
                st.dataframe(df_cancelaciones)
                
                if not df_cancelaciones.empty:
                    excel_cancelados = to_excel(df_cancelaciones)
                    nombre_archivo_cancelados = f"{fecha_aaaammdd}_CancelOrders_SellerFlexFR.xlsx"
                    st.download_button(
                        label="📥 Descargar Fichero Cancelaciones (Excel)",
                        data=excel_cancelados,
                        file_name=nombre_archivo_cancelados,
                        mime="application/vnd.ms-excel"
                    )
                    
                    st.markdown("---")
                    st.subheader("📧 Correo de Cancelación para Amazon")
                    st.write("Manda el listado de cancelaciones por falta de stock directamente a **fr-sellerflex-support@amazon.com**:")
                    if st.button("🚀 Enviar Cancelaciones a Amazon Support", type="primary", key="send_cancel_btn"):
                        with st.spinner("Enviando correo..."):
                            exito_amazon = enviar_correo_background(excel_cancelados, nombre_archivo_cancelados, "fr-sellerflex-support@amazon.com")
                            if exito_amazon:
                                st.success("📬 ¡Correo enviado con éxito a soporte de Amazon!")
                else:
                    st.info("No se han detectado pedidos para cancelar hoy.")
                
            with pestana3:
                st.subheader("Generación de Fichero Definitivo de Almacén")
                
                st.subheader("📧 Comunicación Automatizada con Almacén Francia")
                st.write("Adjunta tu pantallazo diario de transportistas y envíalo directamente a **almacenfrancia@cecotec.es**:")
                fichero_nuevo_diario = st.file_uploader("Adjuntar el nuevo fichero diario o pantallazo de transportistas", type=["png", "jpg", "jpeg", "pdf", "xlsx", "csv"], key="fichero_diario_transp")
                
                if st.button("🚀 Enviar Fichero y Pantallazo Diario a Almacén Francia", key="send_global_mail_btn", type="primary"):
                    with st.spinner("Conectando con el servidor SMTP y enviando correo al Almacén..."):
                        if 'df_almacen_final_mem' in st.session_state:
                            excel_to_send = st.session_state['df_almacen_final_mem']
                            filename_to_send = f"{fecha_aaaammdd}_D-PEDIDOS_FLEX_FR_DEFINITIVO.xlsx"
                        else:
                            excel_to_send = to_excel(pd.DataFrame([{"Aviso": "Verificar listado definitivo adjunto"}]))
                            filename_to_send = f"{fecha_aaaammdd}_Pre-Alerta_Transporte.xlsx"
                            
                        exito_mail = enviar_correo_background(excel_to_send, filename_to_send, "almacenfrancia@cecotec.es", fichero_nuevo_diario)
                        if exito_mail:
                            st.success("📬 ¡Correo enviado con éxito a almacenfrancia@cecotec.es con copia a Juan y Antonio!")

                st.markdown("---")
                st.markdown("""
                Descarga primero el archivo de Cecopartners de la pestaña 1, súbelo a su plataforma, y cuando te devuelvan el **fichero con las referencias D**, cárgalo aquí abajo junto al listado de envíos globales de Amazon:
                """)
                
                col_c1, col_c2 = st.columns(2)
                with col_c1:
                    cecopartners_downloaded_file = st.file_uploader("1. Subir Fichero Descargado de Cecopartners (Con las D)", type=["csv", "xlsx"], key="ceco_almacen")
                with col_c2:
                    sran_shipments_file = st.file_uploader("2. Subir Nuevo Informe Global de Envíos (CSV de Amazon SRAN)", type=["csv", "txt"], key="sran_shipments")

                if cecopartners_downloaded_file is not None and sran_shipments_file is not None:
                    df_ceco_in = load_data(cecopartners_downloaded_file)
                    
                    try:
                        sran_bytes = sran_shipments_file.read()
                        df_sran = pd.read_csv(io.BytesIO(sran_bytes), sep=';', encoding='utf-8')
                        if len(df_sran.columns) <= 1:
                            df_sran = pd.read_csv(io.BytesIO(sran_bytes), sep=',', encoding='utf-8')
                    except Exception:
                        sran_shipments_file.seek(0)
                        sran_bytes = sran_shipments_file.read()
                        df_sran = pd.read_csv(io.BytesIO(sran_bytes), sep=';', encoding='latin-1', on_bad_lines='skip')
                    
                    if df_ceco_in is not None and not df_ceco_in.empty and df_sran is not None and not df_sran.empty:
                        try:
                            df_almacen_fr = df_ceco_in.copy()
                            df_sran.columns = [str(c).replace('env√≠o', 'envío').strip() for c in df_sran.columns]
                            
                            col_sran_pedido = 'Identificador de pedido de cliente'
                            col_sran_envio = 'Identificador de envío'
                            
                            # ---- BLINDAJE EXTRA DE REPETIDOS (Amazon SRAN): Agrupamos en listas ordenadas ----
                            mapa_sran_u_list = df_sran.groupby(col_sran_pedido)[col_sran_envio].apply(lambda x: x.astype(str).str.strip().tolist()).to_dict()
                            
                            # Buscemos la columna de cruce clave de Cecopartners ('NÚMERO DE LÍNEA DE PEDIDO DE CLIENTE')
                            dict_cols_ceco = {col.lower().strip(): col for col in df_almacen_fr.columns}
                            clave_ceco = dict_cols_ceco.get('número de línea de pedido de cliente', df_almacen_fr.columns[-9] if len(df_almacen_fr.columns) > 9 else df_almacen_fr.columns[-1])
                            
                            lista_final_u = []
                            lista_final_p = []
                            lista_final_agencia = []
                            
                            # Duplicamos listas locales consumibles (`.pop(0)`) para asegurar identificadores únicos en filas iguales
                            mapas_u_consumibles = {k: list(v) for k, v in mapa_sran_u_list.items()}
                            mapas_p_consumibles = {k: list(v) for k, v in mapa_pedido_largo_a_zonas_list.items()}
                            mapas_agencia_consumibles = {k: list(v) for k, v in mapa_pedido_largo_a_agencias_list.items()}
                            
                            for idx, row in df_almacen_fr.iterrows():
                                codigo_pedido = str(row[clave_ceco]).strip()
                                
                                # Asignación secuencial e independiente de U (Identificador Envío)
                                if codigo_pedido in mapas_u_consumibles and len(mapas_u_consumibles[codigo_pedido]) > 0:
                                    envio_asignado = mapas_u_consumibles[codigo_pedido].pop(0)
                                else:
                                    envio_asignado = ""
                                lista_final_u.append(envio_asignado)
                                
                                # Asignación secuencial e independiente de P (Zona)
                                if codigo_pedido in mapas_p_consumibles and len(mapas_p_consumibles[codigo_pedido]) > 0:
                                    zona_asignada = mapas_p_consumibles[codigo_pedido].pop(0)
                                else:
                                    zona_asignada = ""
                                lista_final_p.append(zona_asignada)
                                
                                # Asignación secuencial e independiente de Agencia
                                if codigo_pedido in mapas_agencia_consumibles and len(mapas_agencia_consumibles[codigo_pedido]) > 0:
                                    agencia_asignada = mapas_agencia_consumibles[codigo_pedido].pop(0)
                                else:
                                    agencia_asignada = 'AMZN_FR_SH_SD'
                                lista_final_agencia.append(agencia_asignada)
                            
                            # Inyectamos de forma limpia los resultados
                            df_almacen_fr['U'] = lista_final_u
                            df_almacen_fr['P'] = lista_final_p
                            df_almacen_fr['Agencia'] = lista_final_agencia
                            
                            # Borrado automático de columnas basura
                            columnas_a_borrar = [c for c in df_almacen_fr.columns if str(c).upper() in ['FALSE', 'DISPONIBLE']]
                            if columnas_a_borrar:
                                df_almacen_fr.drop(columns=columnas_a_borrar, inplace=True)
                            
                            # Reordenación para fijar P, U y Agencia al principio del Excel resultante
                            cols_ordenadas = ['P', 'U', 'Agencia'] + [c for c in df_almacen_fr.columns if c not in ['P', 'U', 'Agencia']]
                            df_almacen_fr = df_almacen_fr[cols_ordenadas]
                            
                            st.subheader("📝 Tabla Editable de Almacén Francia")
                            st.markdown("Cada número de pedido repetido ha recibido su identificador de envío `U` correspondiente de forma independiente.")
                            
                            # Editor interactivo en pantalla
                            df_editable_final = st.data_editor(df_almacen_fr, use_container_width=True, num_rows="dynamic")
                            
                            st.session_state['df_almacen_final_mem'] = to_excel(df_editable_final)
                            
                            st.download_button(
                                label="📥 Descargar D-PEDIDOS Almacén Definitivo (Excel)",
                                data=st.session_state['df_almacen_final_mem'],
                                file_name=f"{fecha_aaaammdd}_D-PEDIDOS_FLEX_FR_DEFINITIVO.xlsx",
                                mime="application/vnd.ms-excel"
                            )
                        except Exception as ex_cruce:
                            st.error(f"Error procesando el cruce secuencial de Almacén: {ex_cruce}")
                    else:
                        st.warning("Verifica que los archivos subidos contengan registros válidos.")
                else:
                    st.info("⏳ Por favor, sube el archivo de Cecopartners (1) y el Nuevo CSV de Envíos de Amazon (2) en esta pestaña.")

        except Exception as e:
            st.error(f"Error estructural en el procesamiento: {e}")
else:
    st.sidebar.warning("Falta subir archivos obligatorios.")
    st.info("👋 Por favor, carga los archivos requeridos en la barra lateral para empezar a operar.")
