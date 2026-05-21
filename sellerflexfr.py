import streamlit as st
import pandas as pd
import re
import io
import os

st.set_page_config(page_title="Seller Flex FR Automator", layout="wide", page_icon="📦")

st.title("📦 Automatización Seller Flex FR - Cecopartners")
st.markdown("""
Esta aplicación procesa los ficheros de Seller Flex, comprueba la disponibilidad en el Stock de Francia (filtrando referencias con stock ≤ 2) y genera la documentación para Cecopartners, cancelaciones y almacén.
""")

st.sidebar.header("Carga de Ficheros")

# 1. Primer cargador
stock_file = st.sidebar.file_uploader("1. Fichero de Stock FR (CSV)", type=["csv", "txt"])

# 2. Segundo cargador (Fase 2) con su captura de pantalla correspondiente
st.sidebar.markdown("---")
pedidos_recoger_file = st.sidebar.file_uploader("2. Pedidos de la lista de recogida (Excel/CSV)", type=["csv", "xlsx"])
if os.path.exists("pedidoslistarecogida.png"):
    st.sidebar.image("pedidoslistarecogida.png", caption="Ayuda: Archivo de pedidos de la lista de recogida", use_container_width=True)

# 3. Tercer cargador con su captura de pantalla correspondiente
st.sidebar.markdown("---")
listar_recogida_file = st.sidebar.file_uploader("3. Fichero con ID pedidos (Excel/CSV)", type=["csv", "xlsx"])
if os.path.exists("idpedidos.png"):
    st.sidebar.image("idpedidos.png", caption="Ayuda: Archivo con IDs de pedido y envío", use_container_width=True)


def load_data(file, is_stock=False):
    """Carga los ficheros manejando errores de codificación, separadores y filas malformadas."""
    if file is not None:
        if file.name.endswith('.csv') or file.name.endswith('.txt'):
            if is_stock:
                for enc in ['utf-8', 'latin-1', 'iso-8859-1']:
                    try:
                        file.seek(0)
                        return pd.read_csv(file, sep=';', encoding=enc, on_bad_lines='skip')
                    except Exception:
                        try:
                            file.seek(0)
                            return pd.read_csv(file, sep=',', encoding=enc, on_bad_lines='skip')
                        except Exception:
                            continue
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

# Verificar que los 3 ficheros requeridos han sido subidos
if stock_file and pedidos_recoger_file and listar_recogida_file:
    
    st.info("Procesando ficheros... Por favor, espera.")
    
    # Carga de los DataFrames
    df_stock = load_data(stock_file, is_stock=True)
    df_pedidos_recoger = load_data(pedidos_recoger_file)
    df_listar_recogida = load_data(listar_recogida_file)
    
    # Estructura fija en memoria de la plantilla de Cecopartners
    columnas_plantilla = [
        'article', 'quantity', 'customer_name', 'nif', 'attention_of_customer', 
        'address', 'postal_code', 'phone', 'city', 'country_code', 
        'customer_mail', 'comment', 'addressee_order_number'
    ]
    
    if df_stock is None or df_stock.empty:
        st.error("No se pudo procesar correctamente el archivo de Stock. Verifica que sea un CSV válido.")
    else:
        try:
            # ---- PROCESAMIENTO 1: Parsear el Fichero Opcion 3 (ID Pedidos) ----
            col_listar_a = df_listar_recogida.columns[0] 
            
            def parse_listar_recogida(text):
                text = str(text).strip()
                n_pedido = text.split()[0] if len(text.split()) > 0 else ""
                match = re.search(r'(?:ID de envío:|ID de envio:)\s*([A-Za-z0-9]+)', text)
                id_envio = match.group(1) if match else ""
                return pd.Series([n_pedido, id_envio])

            df_listar_recogida[['Num_Pedido_LR', 'Id_Envio_LR']] = df_listar_recogida[col_listar_a].apply(parse_listar_recogida)
            mapa_envio_pedido = dict(zip(df_listar_recogida['Id_Envio_LR'].str.strip(), df_listar_recogida['Num_Pedido_LR'].str.strip()))

            # ---- PROCESAMIENTO 2: Limpieza y Mapeo en Opcion 2 (Pedidos de Lista de Recogida) ----
            df_pedidos_recoger['SKU_Limpio'] = df_pedidos_recoger['SKU'].apply(clean_sku)
            df_pedidos_recoger['Identificador_Clean'] = df_pedidos_recoger['Identificador de pedido'].astype(str).str.strip()
            df_pedidos_recoger['Número_Pedido_Final'] = df_pedidos_recoger['Identificador_Clean'].map(mapa_envio_pedido).fillna("")
            
            # ---- PROCESAMIENTO 3: Mapeo de Unidades de Stock Disponibles ----
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
            
            # NUEVA REGLA DE CORTE: Si el stock actual es estrictamente superior a 2, está disponible. 
            # Si es igual o inferior a 2, pasa a cancelados.
            df_pedidos_recoger['Disponible'] = df_pedidos_recoger['Stock_Actual'] > 2
            
            # ---- FILTRADO Y DIVISIÓN ----
            df_ok = df_pedidos_recoger[df_pedidos_recoger['Disponible'] == True].copy()
            df_cancel = df_pedidos_recoger[df_pedidos_recoger['Disponible'] == False].copy()
            
            # ---- CONSTRUCCIÓN DE LOS FICHEROS DE SALIDA ----
            
            # 1. FICHERO FINAL CECOPARTNERS (Generado a partir de los datos aptos de la Opción 2)
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
            
            # 3. FICHERO D-PEDIDOS Almacén Francia
            df_almacen_fr = pd.DataFrame()
            if not df_ok.empty:
                df_almacen_fr['P'] = df_ok['Zona']
                df_almacen_fr['U'] = df_ok['Identificador de pedido']
                df_almacen_fr['Agencia'] = 'AMZN_FR_SH'
                df_almacen_fr['REFERENCIA'] = [f"D26600{114+i}-1 SGA" for i in range(len(df_ok))]
                df_almacen_fr['NOMBRE DEL CLIENTE'] = 'AMAZON FLEX'
                df_almacen_fr['ESTADO'] = 'ESPERANDO ETIQUETA'
                df_almacen_fr['NÚMERO DE PEDIDO DE CLIENTE'] = df_ok['Número_Pedido_Final']
                
            # ---- RENDERIZADO EN STREAMLIT ----
            st.success("✨ ¡Ficheros generados con éxito! Comprobación de existencias completada (Mínimo > 2 unidades).")
            
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
                    st.download_button(
                        label="📥 Descargar Fichero Cancelaciones (Excel)",
                        data=to_excel(df_cancelaciones),
                        file_name="CancelOrders_SellerFlexFR.xlsx",
                        mime="application/vnd.ms-excel"
                    )
                else:
                    st.info("No se han detectado pedidos para cancelar (Todos tienen stock suficiente).")
                
            with pestana3:
                st.subheader("D-PEDIDOS Almacén Francia")
                st.dataframe(df_almacen_fr)
                if not df_almacen_fr.empty:
                    st.download_button(
                        label="📥 Descargar D-PEDIDOS Almacén (Excel)",
                        data=to_excel(df_almacen_fr),
                        file_name="D-PEDIDOS_FLEX_FR.xlsx",
                        mime="application/vnd.ms-excel"
                    )

        except Exception as e:
            st.error(f"Error estructural en las columnas de los ficheros: {e}")
            st.warning("Asegúrate de que las columnas coincidan exactamente con la estructura de las capturas de pantalla guía.")
else:
    st.info("👋 Por favor, carga los 3 archivos requeridos en la barra lateral para procesar las fases.")
