import streamlit as st
import pandas as pd
import re
import io
import os
import requests
from datetime import datetime, date

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

def enviar_correo_almacen_francia(excel_data, filename, archivo_imagen=None):
    """Envía el fichero final D-PEDIDOS a Almacén Francia con el asunto actualizado que incluye la fecha actual de envío."""
    try:
        if "email" not in st.secrets:
            st.error("❌ Error: No se ha encontrado la sección [email] en los Secrets de Streamlit.")
            return False
            
        smtp_server = st.secrets["email"]["smtp_server"]
        smtp_port = st.secrets["email"]["smtp_port"]
        sender_email = st.secrets["email"]["sender_email"]
        sender_password = st.secrets["email"]["sender_password"]
        
        destinatario = "almacenfrancia@cecotec.es"
        cc_emails = ["juanbrox@cecotec.es", "antoniodiaz@cecotec.es"]
        
        # Fecha de envío actual formateada
        fecha_actual = datetime.now().strftime("%d/%m/%Y")
        fecha_hoy = fecha_actual
        asunto = f"Fichero D-PEDIDOS FLEX FR - Envío: {fecha_actual}"
        
        cuerpo_mensaje = f"Buenos días,\n\nAdjunto el archivo definitivo D-PEDIDOS de Seller Flex Francia con las referencias D conciliadas para su preparación correspondiente a la fecha de envío {fecha_hoy}.\n"
        if archivo_imagen is not None:
            cuerpo_mensaje += "\nSe adjunta también la captura con los detalles de transporte y recogida solicitados.\n"
        cuerpo_mensaje += "\nUn saludo."
        
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = destinatario
        msg['Cc'] = ", ".join(cc_emails)
        msg['Subject'] = asunto
        msg.attach(MIMEText(cuerpo_mensaje, 'plain'))
        
        # Adjuntar Excel
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(excel_data)
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename="{filename}"')
        msg.attach(part)
        
        # Adjuntar captura de imagen si existe
        if archivo_imagen is not None:
            img_part = MIMEBase('application', 'octet-stream')
            img_part.set_payload(archivo_imagen.getvalue())
            encoders.encode_base64(img_part)
            img_part.add_header('Content-Disposition', f'attachment; filename="{archivo_imagen.name}"')
            msg.attach(img_part)
            
        todos_los_destinatarios = [destinatario] + cc_emails
        
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, todos_los_destinatarios, msg.as_string())
        server.quit()
        return True
    except Exception as e:
        st.error(f"❌ Error al enviar el correo al almacén: {e}")
        return False

st.sidebar.header("Carga de Ficheros Principales")
st.sidebar.info("🌐 Stock de Francia: Se descarga automáticamente desde Cecotec Cloud en tiempo real.")

st.sidebar.markdown("---")
pedidos_recoger_file = st.sidebar.file_uploader("1. Pedidos de la lista de recogida (Excel/CSV)", type=["csv", "xlsx"])
if os.path.exists("pedidoslistarecogida.png"):
    st.sidebar.image("pedidoslistarecogida.png", caption="Ayuda: Archivo de pedidos de la lista de recogida", use_container_width=True)

st.sidebar.markdown("---")
listar_recogida_file = st.sidebar.file_uploader("2. Fichero con ID pedidos (Excel/CSV)", type=["csv", "xlsx"])
if os.path.exists("idpedidos.png"):
    st.sidebar.image("idpedidos.png", caption="Ayuda: Archivo con IDs de pedido y envío", use_container_width=True)

def load_excel_or_csv(file):
    if file is None:
        return None
    name = file.name.lower()
    if name.endswith('.xlsx') or name.endswith('.xls'):
        return pd.read_excel(file)
    for enc in ['utf-8', 'latin-1', 'iso-8859-1']:
        try:
            file.seek(0)
            return pd.read_csv(file, sep=None, engine='python', encoding=enc, on_bad_lines='skip')
        except Exception:
            continue
    return None

def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Datos')
    return output.getvalue()

def clean_sku(sku_val):
    if pd.isna(sku_val):
        return ""
    s = str(sku_val).strip().upper()
    s = re.sub(r'^(FR|IT|DE|ES|S)[\s\-_]*', '', s)
    if '.' in s:
        s = s.split('.')[0]
    return s.zfill(5)

@st.cache_data(ttl=600)
def fetch_stock_from_cloud():
    try:
        url_stock = "https://cloud.cecotec.tech/s/FranciaStockRealTime/download"
        res = requests.get(url_stock, timeout=15)
        if res.status_code == 200:
            df = pd.read_csv(io.StringIO(res.text), sep=';', encoding='latin-1')
            return df
    except Exception:
        pass
    return None

df_stock = fetch_stock_from_cloud()

if pedidos_recoger_file and listar_recogida_file:
    df_pedidos_recoger = load_excel_or_csv(pedidos_recoger_file)
    df_listar_recogida = load_excel_or_csv(listar_recogida_file)
    
    if df_stock is None:
        st.error("❌ No se ha podido descargar el stock en tiempo real desde Cecotec Cloud.")
    elif df_pedidos_recoger is not None and df_listar_recogida is not None:
        try:
            df_pedidos_recoger.columns = [c.strip() for c in df_pedidos_recoger.columns]
            df_listar_recogida.columns = [c.strip() for c in df_listar_recogida.columns]
            df_stock.columns = [c.strip() for c in df_stock.columns]
            
            col_listar_a = df_listar_recogida.columns[0]
            
            def parse_listar_recogida(text):
                text = str(text).strip()
                partes = text.split()
                n_pedido = partes[0] if len(partes) > 0 else ""
                id_envio = text[-9:].strip() if len(text) >= 9 else ""
                return pd.Series([n_pedido, id_envio])

            df_listar_recogida[['Num_Pedido_LR', 'Id_Envio_LR']] = df_listar_recogida[col_listar_a].apply(parse_listar_recogida)
            mapa_envio_pedido = dict(zip(df_listar_recogida['Id_Envio_LR'].str.strip(), df_listar_recogida['Num_Pedido_LR'].str.strip()))

            df_pedidos_recoger['SKU_Limpio'] = df_pedidos_recoger['SKU'].apply(clean_sku)
            df_pedidos_recoger['Identificador_Clean'] = df_pedidos_recoger['Identificador de pedido'].astype(str).str.strip()
            df_pedidos_recoger['Número_Pedido_Final'] = df_pedidos_recoger['Identificador_Clean'].map(mapa_envio_pedido).fillna("")
            
            # Persistencia en st.session_state para evitar pérdidas en los refrescos de la pestaña 3
            mapa_p = {}
            mapa_u = {}
            for _, r in df_pedidos_recoger.iterrows():
                ped_f = str(r.get('Número_Pedido_Final', '')).strip()
                ident = str(r.get('Identificador de pedido', '')).strip()
                zona_val = str(r.get('Zona', '')).strip()
                
                if ped_f and ped_f != "nan" and ped_f != "":
                    mapa_p[ped_f] = zona_val
                    mapa_u[ped_f] = ident
                if ident and ident != "nan" and ident != "":
                    mapa_p[ident] = zona_val
                    mapa_u[ident] = ident
                    
            st.session_state['mapa_pedido_a_zona'] = mapa_p
            st.session_state['mapa_pedido_a_envio'] = mapa_u

            col_stock_ref = df_stock.columns[0]
            col_stock_cant = df_stock.columns[2] if len(df_stock.columns) > 2 else df_stock.columns[-1]
            dict_stock = dict(zip(df_stock[col_stock_ref].apply(clean_sku), pd.to_numeric(df_stock[col_stock_cant], errors='coerce').fillna(0)))

            lineas_ok = []
            lineas_ooo = []
            
            for _, row_p in df_pedidos_recoger.iterrows():
                sku_l = row_p['SKU_Limpio']
                unidades_pedidas = pd.to_numeric(row_p['Unidades'], errors='coerce')
                if pd.isna(unidades_pedidas): unidades_pedidas = 1
                stock_dispo = dict_stock.get(sku_l, 0)
                
                if stock_dispo > 2 and stock_dispo >= unidades_pedidas:
                    lineas_ok.append(row_p)
                    dict_stock[sku_l] = stock_dispo - unidades_pedidas
                else:
                    lineas_ooo.append(row_p)

            t1, t2, t3 = st.tabs(["📋 Subida Cecopartners", "❌ Cancelaciones OOO", "🚚 Fichero Almacén (D-PEDIDOS)"])
            fecha_hoy_str = datetime.now().strftime("%Y%m%d")

            with t1:
                if len(lineas_ok) > 0:
                    df_ok = pd.DataFrame(lineas_ok)
                    df_subida = pd.DataFrame()
                    df_subida['article'] = df_ok['SKU_Limpio']
                    df_subida['quantity'] = df_ok['Unidades']
                    df_subida['customer_name'] = 'AMAZON FLEX'
                    df_subida['attention_of_customer'] = 'AMAZON FLEX'
                    df_subida['address'] = 'Cam.Real de Madrid 117'
                    df_subida['postal_code'] = '46292'
                    df_subida['city'] = 'Massalavés'
                    df_subida['country_code'] = 'ES'
                    df_subida['addressee_order_number'] = df_ok['Número_Pedido_Final']
                    
                    st.success(f"🎉 ¡Líneas con stock correcto validadas: {len(df_subida)} hileras encontradas!")
                    st.dataframe(df_subida, use_container_width=True)
                    
                    nombre_archivo_subida = f"SELLER_FLEX_FR_{fecha_hoy_str}.xlsx"
                    st.download_button(
                        label="📥 Descargar Subida Cecopartners (Excel)",
                        data=to_excel(df_subida),
                        file_name=nombre_archivo_subida,
                        mime="application/vnd.ms-excel",
                        use_container_width=True
                    )
                else:
                    st.warning("No se encontraron líneas con stock suficiente disponible en Francia (Stock > 2).")

            with t2:
                if len(lineas_ooo) > 0:
                    df_ooo = pd.DataFrame(lineas_ooo)
                    df_cancelaciones = pd.DataFrame()
                    df_cancelaciones['Node ID'] = 'SRAN'
                    df_cancelaciones['Order number'] = df_ooo['Número_Pedido_Final']
                    df_cancelaciones['Shipment ID'] = df_ooo['Identificador de pedido']
                    df_cancelaciones['ASIN'] = df_ooo['FNSKU']
                    df_cancelaciones['Reason'] = 'OOO'
                    
                    st.error(f"⚠️ Se han detectado {len(df_cancelaciones)} líneas en rotura o stock crítico (Stock ≤ 2).")
                    st.dataframe(df_cancelaciones, use_container_width=True)
                    
                    nombre_archivo_ooo = f"CANCELACIONES_FLEX_FR_{fecha_hoy_str}.xlsx"
                    st.download_button(
                        label="📥 Descargar Archivo de Cancelaciones (Excel)",
                        data=to_excel(df_cancelaciones),
                        file_name=nombre_archivo_ooo,
                        mime="application/vnd.ms-excel",
                        use_container_width=True
                    )
                else:
                    st.success("✅ ¡Excelente! Cero hileras en rotura de stock para el día de hoy.")

            with t3:
                st.markdown("""
                **Paso intermedio:** Descarga primero el archivo de Cecopartners de la pestaña 1, súbelo a su plataforma, y cuando te devuelvan el **fichero con las referencias D**, cárgalo aquí abajo para estructurar el definitivo de almacén:
                """)
                
                cecopartners_downloaded_file = st.file_uploader("Subir Fichero Descargado de Cecopartners (Excel/CSV) para cruzar las D", type=["csv", "xlsx"], key="ceco_almacen")
                
                if cecopartners_downloaded_file is not None:
                    df_ceco_in = load_excel_or_csv(cecopartners_downloaded_file)
                    
                    if df_ceco_in is not None and not df_ceco_in.empty:
                        try:
                            dict_cols_ceco = {col.lower(): col for col in df_ceco_in.columns}
                            col_ceco_pedido = dict_cols_ceco.get('número de línea de pedido de cliente', dict_cols_ceco.get('número de pedido de cliente', dict_cols_ceco.get('addressee_order_number', df_ceco_in.columns[-1])))
                            
                            pedidos_ceco_limpios = df_ceco_in[col_ceco_pedido].astype(str).str.strip()
                            
                            mapa_zona_session = st.session_state.get('mapa_pedido_a_zona', {})
                            mapa_envio_session = st.session_state.get('mapa_pedido_a_envio', {})
                            
                            # Generar hileras de cruce calculadas
                            col_P = pedidos_ceco_limpios.map(mapa_zona_session).fillna("")
                            col_U = pedidos_ceco_limpios.map(mapa_envio_session).fillna("")
                            col_Agencia = ['AMZN_FR_SH_SD'] * len(df_ceco_in)
                            
                            # Borrar la columna A (primera columna del archivo)
                            df_ceco_restante = df_ceco_in.iloc[:, 1:].copy()
                            
                            # Reensamblar con P, U, Agencia al inicio y el resto de Cecopartners intacto hacia la derecha
                            df_almacen_fr = pd.DataFrame()
                            df_almacen_fr['P'] = col_P
                            df_almacen_fr['U'] = col_U
                            df_almacen_fr['Agencia'] = col_Agencia
                            df_almacen_fr = pd.concat([df_almacen_fr, df_ceco_restante], axis=1)
                            
                            nombre_archivo_dpedidos = f"D-PEDIDOS_FLEX_FR_{fecha_hoy_str}.xlsx"
                            
                            st.success("🎉 ¡Fichero definitivo para Almacén Francia generado con las 'D' cruzadas exitosamente!")
                            st.dataframe(df_almacen_fr, use_container_width=True)
                            
                            st.download_button(
                                label="📥 Descargar D-PEDIDOS Almacén (Excel)",
                                data=to_excel(df_almacen_fr),
                                file_name=nombre_archivo_dpedidos,
                                mime="application/vnd.ms-excel",
                                use_container_width=True
                            )
                            
                            st.markdown("---")
                            st.subheader("📧 Envío Automatizado a Almacén Francia (Con Detalles de Transporte)")
                            st.write("Puedes adjuntar de manera opcional una captura de pantalla o imagen con los detalles del transporte/recogida para el almacén:")
                            
                            archivo_captura_transporte = st.file_uploader(
                                "📸 Adjuntar captura de detalles de transporte (Opcional)", 
                                type=["png", "jpg", "jpeg"],
                                key="captura_transporte_almacen"
                            )
                            
                            if archivo_captura_transporte is not None:
                                st.image(archivo_captura_transporte, caption="Vista previa de los detalles de transporte adjuntos", width=400)
                            
                            if st.button("🚀 Enviar D-PEDIDOS + Transporte a Almacén Francia", type="primary", use_container_width=True):
                                with st.spinner("Enviando correo con los adjuntos cargados..."):
                                    excel_bytes = to_excel(df_almacen_fr)
                                    exito = enviar_correo_almacen_francia(excel_bytes, nombre_archivo_dpedidos, archivo_captura_transporte)
                                    if exito:
                                        st.success("📬 ¡Correo enviado con éxito al Almacén de Francia! Se ha incluido el fichero definitivo con el asunto fechado y la captura de transporte adjunta.")
                                        
                        except Exception as ex_cruce:
                            st.error(f"Error procesando el fichero devuelto de Cecopartners: {ex_cruce}")
                    else:
                        st.warning("El archivo de Cecopartners subido está vacío o no es válido.")
                else:
                    st.info("⏳ Esperando que subas el archivo descargado de Cecopartners con las referencias D para generar el listado definitivo de almacén.")
        except Exception as e:
            st.error(f"Error estructural en las columnas de los ficheros: {e}")
else:
    st.info("👋 Por favor, carga los 2 archivos manuales requeridos en la barra lateral para empezar a operar.")
