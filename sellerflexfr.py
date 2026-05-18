import streamlit as st
import pandas as pd
import re
import io

st.set_page_config(page_title="Seller Flex FR Automator", layout="wide", page_icon="📦")

st.title("📦 Automatización Seller Flex FR - Cecopartners")
st.markdown("""
Sube los ficheros necesarios para procesar los pedidos, verificar stock y generar los documentos de subida, cancelaciones y almacén de Francia.
""")

st.sidebar.header("Carga de Ficheros")

# 1. Cargar Fichero de Stock FR
stock_file = st.sidebar.file_uploader("1. Fichero de Stock FR (CSV o Excel)", type=["csv", "xlsx"])

# 2. Cargar Fichero PedidosRecoger
pedidos_recoger_file = st.sidebar.file_uploader("2. Fichero PedidosRecoger (Excel/CSV)", type=["csv", "xlsx"])

# 3. Cargar Fichero ListarRecogida
listar_recogida_file = st.sidebar.file_uploader("3. Fichero ListarRecogida (Excel/CSV)", type=["csv", "xlsx"])

# 4. Cargar Fichero Plantilla/Base (SELLER_FLEX_FR)
plantilla_file = st.sidebar.file_uploader("4. Fichero Plantilla SELLER_FLEX_FR (Excel/CSV)", type=["csv", "xlsx"])

def load_data(file):
    if file is not None:
        if file.name.endswith('.csv'):
            return pd.read_csv(file)
        else:
            return pd.read_excel(file)
    return None

def clean_sku(sku):
    """Limpia el SKU quitando el prefijo FR y los ceros a la izquierda"""
    if pd.isna(sku):
        return ""
    sku_str = str(sku).strip()
    if sku_str.upper().startswith("FR"):
        sku_str = sku_str[2:]  # Quita 'FR'
    sku_str = sku_str.lstrip('0')  # Quita ceros a la izquierda
    return sku_str

# Función para convertir dataframe a Excel descargable
def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Datos')
    processed_data = output.getvalue()
    return processed_data

if stock_file and pedidos_recoger_file and listar_recogida_file and plantilla_file:
    
    st.info("Procesando datos... Por favor, espera.")
    
    # Cargar DataFrames
    df_stock = load_data(stock_file)
    df_pedidos_recoger = load_data(pedidos_recoger_file)
    df_listar_recogida = load_data(listar_recogida_file)
    df_plantilla = load_data(plantilla_file)
    
    try:
        # ---- PROCESAMIENTO 1: Limpieza de SKU en PedidosRecoger ----
        # Identificar columnas (ajustar nombres según cabeceras reales)
        col_sku = [c for c in df_pedidos_recoger.columns if 'SKU' in c.upper()][0]
        col_id_envio = [c for c in df_pedidos_recoger.columns if 'IDENTIFICADOR' in c.upper() or 'ENVÍO' in c.upper() or 'ID' in c.upper()][4] # O la que corresponda
        # Forzar por el nombre exacto de tus muestras:
        if 'SKU' in df_pedidos_recoger.columns: col_sku = 'SKU'
        if 'Identificador de pedido' in df_pedidos_recoger.columns: col_id_envio = 'Identificador de pedido'
        if 'Zona' in df_pedidos_recoger.columns: col_zona = 'Zona' # ID de lista de recogida (P....)
        
        df_pedidos_recoger['SKU_Limpio'] = df_pedidos_recoger[col_sku].apply(clean_sku)
        
        # ---- PROCESAMIENTO 2: Cruzar con Stock ----
        # Asumiendo que en df_stock la columna A es la primera columna
        col_stock_ref = df_stock.columns[0]
        # Asegurar stock como string limpio
        df_stock[col_stock_ref] = df_stock[col_stock_ref].astype(str).str.strip().apply(lambda x: x.lstrip('0'))
        
        # Lista de referencias con stock disponible
        referencias_disponibles = set(df_stock[col_stock_ref].unique())
        
        # Clasificar pedidos de PedidosRecoger
        df_pedidos_recoger['Disponible'] = df_pedidos_recoger['SKU_Limpio'].isin(referencias_disponibles)
        
        # ---- PROCESAMIENTO 3: Separar columna A de ListarRecogida ----
        col_listar_a = df_listar_recogida.columns[0] # "ID del pedido" según tu muestra
        
        # Extraer Número de pedido e Identificador de envío usando Regex desde la columna combinada
        # Ejemplo: "406-4244722-3978723 ID de envío: TJhpw0q4D"
        def parse_listar_recogida(text):
            text = str(text)
            n_pedido = text.split()[0] if len(text.split()) > 0 else ""
            id_envio = ""
            match = re.search(r'(?:ID de envío:|ID de envio:)\s*([A-Za-z0-9]+)', text)
            if match:
                id_envio = match.group(1)
            else:
                # Intento secundario si viene al final
                parts = text.split()
                if len(parts) >= 4:
                    id_envio = parts[-1]
            return pd.Series([n_pedido, id_envio])

        df_listar_recogida[['Número de Pedido', 'Identificador de Envío']] = df_listar_recogida[col_listar_a].apply(parse_listar_recogida)
        
        # ---- PROCESAMIENTO 4: Mapear Disponibilidad a ListarRecogida ----
        # Cruzamos por el Identificador de Envío / Pedido
        envios_disponibles = df_pedidos_recoger[df_pedidos_recoger['Disponible'] == True][col_id_envio].unique()
        envios_cancelar = df_pedidos_recoger[df_pedidos_recoger['Disponible'] == False][col_id_envio].unique()
        
        df_listar_recogida['Estado_Stock'] = df_listar_recogida['Identificador de Envío'].apply(
            lambda x: 'OK' if x in envios_disponibles else ('CANCELAR' if x in envios_cancelar else 'OK') # Por defecto OK si no se cruza, o ajustar criterio
        )
        
        # ---- GENERACIÓN DE FICHEROS FINALES ----
        
        # 1. Fichero Cecopartners (Subida)
        # Información: ID de lista de recogida (P....), Identificador de pedido, y agencia (AMZN_FR_SH_SD)
        # Cruzamos ListarRecogida con PedidosRecoger para sacar la Zona (P...)
        df_mismatch_zona = df_pedidos_recoger[[col_id_envio, col_zona]].drop_duplicates()
        df_cecopartners = df_listar_recogida[df_listar_recogida['Estado_Stock'] == 'OK'].merge(
            df_mismatch_zona, left_on='Identificador de Envío', right_on=col_id_envio, how='left'
        )
        
        df_subida_ceco = pd.DataFrame({
            'ID Lista Recogida': df_cecopartners[col_zona],
            'Identificador de Pedido': df_cecopartners['Identificador de Envío'],
            'Agencia': 'AMZN_FR_SH_SD'
        }).dropna(subset=['Identificador de Pedido'])
        
        # 2. Fichero de Cancelaciones (Igual al modelo 20260515_CancelOrders)
        df_cancelados_raw = df_listar_recogida[df_listar_recogida['Estado_Stock'] == 'CANCELAR']
        df_cancel_merge = df_cancelados_raw.merge(df_pedidos_recoger, left_on='Identificador de Envío', right_on=col_id_envio, how='inner')
        
        df_cancelaciones = pd.DataFrame({
            'Node ID': 'SRAN',
            'Order number': df_cancel_merge['Número de Pedido'],
            'Shipment ID': df_cancel_merge['Identificador de Envío'],
            'ASIN': df_cancel_merge['FNSKU'] if 'FNSKU' in df_cancel_merge.columns else '',
            'Reason': 'OOO'
        }).drop_duplicates()
        
        # 3. Fichero D-PEDIDOS para Almacén Francia
        # Asignar un número incremental ficticio o basado en el ERP para el código D-
        df_d_pedidos = df_listar_recogida[df_listar_recogida['Estado_Stock'] == 'OK'].copy()
        
        # Recreamos la estructura simulada del D-PEDIDOS basándonos en tu plantilla
        df_almacen_fr = pd.DataFrame()
        if not df_d_pedidos.empty:
            df_almacen_fr['P'] = df_cecopartners[col_zona]
            df_almacen_fr['U'] = df_cecopartners['Identificador de Envío']
            df_almacen_fr['Agencia'] = 'AMZN_FR_SH'
            # Generar el código D- incremental (ej: D26600114-1, etc.)
            df_almacen_fr['REFERENCIA'] = [f"D26600{114+i}-1 SGA" for i in range(len(df_almacen_fr))]
            df_almacen_fr['NOMBRE DEL CLIENTE'] = 'AMAZON FLEX'
            df_almacen_fr['ESTADO'] = 'ESPERANDO ETIQUETA'
            df_almacen_fr['NÚMERO DE PEDIDO DE CLIENTE'] = df_cecopartners['Número de Pedido']
        
        # ---- INTERFAZ DE RESULTADOS ----
        st.success("¡Ficheros procesados correctamente!")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.subheader("1. Subida Cecopartners")
            st.dataframe(df_subida_ceco.head())
            st.download_button(
                label="📥 Descargar Subida Cecopartners",
                data=to_excel(df_subida_ceco),
                file_name="Subida_Cecopartners_FLEX.xlsx",
                mime="application/vnd.ms-excel"
            )
            
        with col2:
            st.subheader("2. Cancelaciones")
            st.dataframe(df_cancelaciones.head())
            st.download_button(
                label="📥 Descargar Cancelaciones (OOO)",
                data=to_excel(df_cancelaciones),
                file_name="Cancel Orders_SellerFlexFR.xlsx",
                mime="application/vnd.ms-excel"
            )
            
        with col3:
            st.subheader("3. D-PEDIDOS Almacén FR")
            st.dataframe(df_almacen_fr.head())
            st.download_button(
                label="📥 Descargar D-PEDIDOS Almacén",
                data=to_excel(df_almacen_fr),
                file_name="D-PEDIDOS_FLEX_FR.xlsx",
                mime="application/vnd.ms-excel"
            )

    except Exception as e:
        st.error(f"Error al procesar los archivos: {e}")
        st.warning("Asegúrate de que las cabeceras y formatos de los archivos subidos coinciden con las estructuras esperadas.")
else:
    st.info("Por favor, sube los 4 archivos requeridos en la barra lateral para empezar.")