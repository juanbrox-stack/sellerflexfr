import streamlit as st
import pandas as pd
import re
import io

st.set_page_config(page_title="Seller Flex FR Automator", layout="wide", page_icon="📦")

st.title("📦 Automatización Seller Flex FR - Cecopartners")
st.markdown("""
Esta aplicación procesa los ficheros de Seller Flex, comprueba la disponibilidad en el Stock de Francia y genera automáticamente los tres ficheros resultantes.
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
            # Si es el stock, aplicamos un control estricto de errores de parsing
            if is_stock:
                for enc in ['utf-8', 'latin-1', 'iso-8859-1']:
                    try:
                        file.seek(0)
                        # Intentamos primero con punto y coma (común en Europa) omitiendo líneas rotas
                        return pd.read_csv(file, sep=';', encoding=enc, on_bad_lines='skip')
                    except Exception:
                        try:
                            file.seek(0)
                            # Si falla, intentamos con coma tradicional omitiendo líneas rotas
                            return pd.read_csv(file, sep=',', encoding=enc, on_bad_lines='skip')
                        except Exception:
                            continue
            
            # Carga estándar para el resto de archivos CSV
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
    
    # 1. Carga de los DataFrames utilizando el control de errores para Stock
    df_stock = load_data(stock_file, is_stock=True)
    df_pedidos_recoger = load_data(pedidos_recoger_file)
    df_listar_recogida = load_data(listar_recogida_file)
    df_plantilla = load_data(plantilla_file)
    
    if df_stock is None or df_stock.empty:
        st.error("No se pudo procesar correctamente el archivo de Stock. Verifica que sea un CSV válido.")
    else:
        try:
            # ---- PROCESAMIENTO 1: Limpieza y Normalización de PedidosRecoger ----
            df_pedidos_recoger['SKU_Limpio'] = df_pedidos_recoger['SKU'].apply(clean_sku)
            
            # ---- PROCESAMIENTO 2: Extraer Referencias del Stock FR ----
            col_stock_ref = df_stock.columns[0]
            df_stock[col_stock_ref] = df_stock[col_stock_ref].astype(str).str.strip().apply(lambda x: x.lstrip('0'))
            referencias_disponibles = set(df_stock[col_stock_ref].unique())
            
            # Clasificar la disponibilidad de cada artículo en PedidosRecoger
            df_pedidos_recoger['Disponible'] = df_pedidos_recoger['SKU_Limpio'].isin(referencias_disponibles)
            
            # ---- PROCESAMIENTO 3: Segmentación de la columna A en ListarRecogida ----
            col_listar_a = df_listar_recogida.columns[0] # Normalizado como 'ID del pedido'
            
            def parse_listar_recogida(text):
                text = str(text).strip()
                # Extrae el número de pedido largo que va antes del primer espacio
                n_pedido = text.split()[0] if len(text.split()) > 0 else ""
                # Extrae mediante Expresión Regular el ID de envío (ej: TJhpw0q4D)
                match = re.search(r'(?:ID de envío:|ID de envio:)\s*([A-Za-z0-9]+)', text)
                id_envio = match.group(1) if match else ""
                return pd.Series([n_pedido, id_envio])

            df_listar_recogida[['Número de Pedido', 'Identificador de Envío']] = df_listar_recogida[col_listar_a].apply(parse_listar_recogida)
            
            # ---- PROCESAMIENTO 4: Cruzar Disponibilidad General ----
            envios_disponibles = set(df_pedidos_recoger[df_pedidos_recoger['Disponible'] == True]['Identificador de pedido'].unique())
            envios_cancelar = set(df_pedidos_recoger[df_pedidos_recoger['Disponible'] == False]['Identificador de pedido'].unique())
            
            # Asignar estado basándonos en el ID de envío extraído
            def mapear_estado(id_envio):
                if id_envio in envios_cancelar:
                    return 'CANCELAR'
                return 'OK'

            df_listar_recogida['Estado_Stock'] = df_listar_recogida['Identificador de Envío'].apply(mapear_estado)
            
            # ---- CONSTRUCCIÓN DE LOS 3 FICHEROS DE SALIDA ----
            
            # Mapeo de la Zona (P...) uniendo por Identificador de Pedido
            df_zona_mapping = df_pedidos_recoger[['Identificador de pedido', 'Zona', 'FNSKU', 'SKU_Limpio']].drop_duplicates()
            df_master = df_listar_recogida.merge(df_zona_mapping, left_on='Identificador de Envío', right_on='Identificador de pedido', how='left')
            
            # 1. Fichero Resultante Cecopartners (Subida)
            df_ok = df_master[df_master['Estado_Stock'] == 'OK']
            df_subida_ceco = pd.DataFrame({
                'ID de lista de recogida (P….)': df_ok['Zona'],
                'Identificador de pedido': df_ok['Identificador de Envío'],
                'Agencia': 'AMZN_FR_SH_SD'
            }).dropna(subset=['Identificador de pedido']).drop_duplicates()
            
            # 2. Fichero de Cancelaciones (Estructura Modelo OOO)
            df_cancel = df_master[df_master['Estado_Stock'] == 'CANCELAR']
            df_cancelaciones = pd.DataFrame({
                'Node ID': 'SRAN',
                'Order number': df_cancel['Número de Pedido'],
                'Shipment ID': df_cancel['Identificador de Envío'],
                'ASIN': df_cancel['FNSKU'],
                'Reason': 'OOO'
            }).dropna(subset=['Shipment ID']).drop_duplicates()
            
            # 3. Fichero D-PEDIDOS Almacén Francia (Utilizando el formato de la plantilla cargada)
            df_almacen_fr = pd.DataFrame(columns=df_plantilla.columns)
            if not df_ok.empty:
                df_almacen_fr['P'] = df_ok['Zona']
                df_almacen_fr['U'] = df_ok['Identificador de Envío']
                df_almacen_fr['Agencia'] = 'AMZN_FR_SH'
                # Replicación del consecutivo de referencia del ERP Francia D26600...
                df_almacen_fr['REFERENCIA'] = [f"D26600{114+i}-1 SGA" for i in range(len(df_ok))]
                df_almacen_fr['NOMBRE DEL CLIENTE'] = 'AMAZON FLEX'
                df_almacen_fr['ESTADO'] = 'ESPERANDO ETIQUETA'
                df_almacen_fr['NÚMERO DE PEDIDO DE CLIENTE'] = df_ok['Número de Pedido']
                df_almacen_fr['NÚMERO DE PEDIDO'] = [f"D26600{114+i}" for i in range(len(df_ok))]
                df_almacen_fr['ARTÍCULOS'] = df_ok['SKU_Limpio']
                
            # ---- RENDERIZADO DE LA INTERFAZ DE USUARIO EN STREAMLIT ----
            st.success("✨ ¡Ficheros cruzados y procesados correctamente sin errores de formato!")
            
            pestana1, pestana2, pestana3 = st.tabs(["📤 Subida Cecopartners", "❌ Cancelaciones", "🇫🇷 D-PEDIDOS Francia"])
            
            with pestana1:
                st.subheader("Datos de Subida a Cecopartners")
                st.dataframe(df_subida_ceco)
                st.download_button(
                    label="📥 Descargar Subida Cecopartners (Excel)",
                    data=to_excel(df_subida_ceco),
                    file_name="Subida_Cecopartners_FLEX.xlsx",
                    mime="application/vnd.ms-excel"
                )
                
            with pestana2:
                st.subheader("Pedidos sin Stock (Para Cancelar)")
                st.dataframe(df_cancelaciones)
                if not df_cancelaciones.empty:
                    st.download_button(
                        label="📥 Descargar Fichero Cancelaciones (Excel)",
                        data=to_excel(df_cancelaciones),
                        file_name="CancelOrders_SellerFlexFR.xlsx",
                        mime="application/vnd.ms-excel"
                    )
                else:
                    st.info("No se han detectado pedidos sin stock. ¡Fichero de cancelaciones vacío!")
                
            with pestana3:
                st.subheader("D-PEDIDOS Almacén Francia")
                st.dataframe(df_almacen_fr)
                if not df_almacen_fr.empty:
                    st.download_button(
                        label="📥 Descargar D-PEDIDOS Almacén FR (Excel)",
                        data=to_excel(df_almacen_fr),
                        file_name="D-PEDIDOS_FLEX_FR.xlsx",
                        mime="application/vnd.ms-excel"
                    )

        except Exception as e:
            st.error(f"Error estructural en las columnas de los ficheros: {e}")
            st.warning("Verifica que los ficheros de Recogida y Pedidos mantengan los nombres de columnas estándar ('SKU', 'Zona', 'Identificador de pedido').")
else:
    st.info("👋 Por favor, sube los 4 archivos solicitados en la barra lateral para generar la documentación de Cecopartners.")