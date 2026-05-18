import streamlit as st
import pandas as pd
import re
import io

st.set_page_config(page_title="Seller Flex FR Automator", layout="wide", page_icon="📦")

st.title("📦 Automatización Seller Flex FR - Cecopartners")
st.markdown("""
Esta aplicación procesa los ficheros de Seller Flex, comprueba la disponibilidad en el Stock de Francia y genera por separado los ficheros de subida, almacén y cancelaciones de forma estricta.
""")

st.sidebar.header("Carga de Ficheros")

# Componentes de subida en la barra lateral
stock_file = st.sidebar.file_uploader("1. Fichero de Stock FR (CSV)", type=["csv", "txt"])
pedidos_recoger_file = st.sidebar.file_uploader("2. Fichero PedidosRecoger (Excel/CSV)", type=["csv", "xlsx"])
listar_recogida_file = st.sidebar.file_uploader("3. Fichero ListarRecogida (Excel/CSV)", type=["csv", "xlsx"])
plantilla_file = st.sidebar.file_uploader("4. Fichero Plantilla SELLER_FLEX_FR (Excel/CSV)", type=["csv", "xlsx"])

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

# Verificar que todos los ficheros requeridos han sido subidos
if stock_file and pedidos_recoger_file and listar_recogida_file and plantilla_file:
    
    st.info("Procesando ficheros... Por favor, espera.")
    
    # 1. Carga de los DataFrames
    df_stock = load_data(stock_file, is_stock=True)
    df_pedidos_recoger = load_data(pedidos_recoger_file)
    df_listar_recogida = load_data(listar_recogida_file)
    df_plantilla = load_data(plantilla_file)
    
    if df_stock is None or df_stock.empty:
        st.error("No se pudo procesar correctamente el archivo de Stock. Verifica que sea un CSV válido.")
    else:
        try:
            # ---- PROCESAMIENTO 1: Parsear el Fichero ListarRecogida Primero ----
            col_listar_a = df_listar_recogida.columns[0] 
            
            def parse_listar_recogida(text):
                text = str(text).strip()
                n_pedido = text.split()[0] if len(text.split()) > 0 else ""
                match = re.search(r'(?:ID de envío:|ID de envio:)\s*([A-Za-z0-9]+)', text)
                id_envio = match.group(1) if match else ""
                return pd.Series([n_pedido, id_envio])

            df_listar_recogida[['Num_Pedido_LR', 'Id_Envio_LR']] = df_listar_recogida[col_listar_a].apply(parse_listar_recogida)
            mapa_envio_pedido = dict(zip(df_listar_recogida['Id_Envio_LR'].str.strip(), df_listar_recogida['Num_Pedido_LR'].str.strip()))

            # ---- PROCESAMIENTO 2: Limpieza y Mapeo en PedidosRecoger ----
            df_pedidos_recoger['SKU_Limpio'] = df_pedidos_recoger['SKU'].apply(clean_sku)
            df_pedidos_recoger['Identificador_Clean'] = df_pedidos_recoger['Identificador de pedido'].astype(str).str.strip()
            df_pedidos_recoger['Número_Pedido_Final'] = df_pedidos_recoger['Identificador_Clean'].map(mapa_envio_pedido).fillna("")
            
            # ---- PROCESAMIENTO 3: Comprobación de Stock FR ----
            col_stock_ref = df_stock.columns[0]
            df_stock[col_stock_ref] = df_stock[col_stock_ref].astype(str).str.strip().apply(lambda x: x.lstrip('0'))
            referencias_disponibles = set(df_stock[col_stock_ref].unique())
            
            # Clasificamos disponibilidad individual
            df_pedidos_recoger['Disponible'] = df_pedidos_recoger['SKU_Limpio'].isin(referencias_disponibles)
            
            # ---- FILTRADO ABSOLUTO (Garantiza separación total de registros) ----
            df_ok = df_pedidos_recoger[df_pedidos_recoger['Disponible'] == True].copy().reset_index(drop=True)
            df_cancel = df_pedidos_recoger[df_pedidos_recoger['Disponible'] == False].copy().reset_index(drop=True)
            
            # ---- CONSTRUCCIÓN DE LOS 3 FICHEROS DE SALIDA ----
            
            # 1. FICHERO FINAL CECOPARTNERS (Sólo los 5 aprobados con stock)
            df_subida_plantilla = pd.DataFrame(columns=df_plantilla.columns)
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
            
            # 2. FICHERO DE CANCELACIONES (Estructura Solicitada Exacta - Recoge al SKU 2748)
            df_cancelaciones = pd.DataFrame(columns=['Node ID', 'Order number', 'Shipment ID', 'ASIN', 'Reason'])
            if not df_cancel.empty:
                df_cancelaciones['Node ID'] = 'SRAN'
                df_cancelaciones['Order number'] = df_cancel['Número_Pedido_Final']
                df_cancelaciones['Shipment ID'] = df_cancel['Identificador de pedido']
                df_cancelaciones['ASIN'] = df_cancel['FNSKU']
                df_cancelaciones['Reason'] = 'OOO'
            
            # 3. FICHERO D-PEDIDOS Almacén Francia (Sólo los 5 aprobados con stock)
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
            st.success("✨ ¡Ficheros perfectamente divididos y filtrados por stock!")
            
            pestana1, pestana2, pestana3 = st.tabs(["📤 Fichero Subida (Plantilla Cecopartners)", "❌ Cancelaciones (OOO)", "🇫🇷 D-PEDIDOS Francia"])
            
            with pestana1:
                st.subheader(f"Fichero Resultante Cecopartners ({len(df_subida_plantilla)} Pedidos Aptos)")
                st.dataframe(df_subida_plantilla)
                st.download_button(
                    label="📥 Descargar Fichero Subida Cecopartners (Excel)",
                    data=to_excel(df_subida_plantilla),
                    file_name="SELLER_FLEX_FR_PROCESADO.xlsx",
                    mime="application/vnd.ms-excel"
                )
                
            with pestana2:
                st.subheader(f"Fichero de Cancelaciones ({len(df_cancelaciones)} Pedido Sin Stock)")
                st.dataframe(df_cancelaciones)
                if not df_cancelaciones.empty:
                    st.download_button(
                        label="📥 Descargar Fichero Cancelaciones (Excel)",
                        data=to_excel(df_cancelaciones),
                        file_name="20260518_CancelOrders_SellerFlexFR.xlsx",
                        mime="application/vnd.ms-excel"
                    )
                else:
                    st.info("No se han detectado pedidos sin stock.")
                
            with pestana3:
                st.subheader(f"D-PEDIDOS Almacén Francia ({len(df_almacen_fr)} Pedidos)")
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
            st.warning("Verifica que los ficheros mantengan los nombres de columnas estándar.")
else:
    st.info("👋 Por favor, sube los 4 archivos solicitados en la barra lateral para generar la documentación.")