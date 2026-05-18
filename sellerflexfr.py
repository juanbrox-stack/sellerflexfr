import streamlit as st
import pandas as pd
import re
import io
from datetime import date

st.set_page_config(page_title="Seller Flex FR Automator", layout="wide", page_icon="📦")

st.title("📦 Automatización Seller Flex FR — Cecopartners")
st.markdown("Procesa los ficheros de Seller Flex, comprueba stock en Francia y genera los documentos de subida y almacén.")

# ──────────────────────────────────────────────
# HELPERS
# ──────────────────────────────────────────────

def load_excel_or_csv(file):
    if file is None:
        return None
    name = file.name.lower()
    if name.endswith('.xlsx') or name.endswith('.xls'):
        return pd.read_excel(file)
    for enc in ['utf-8', 'latin-1']:
        try:
            file.seek(0)
            return pd.read_csv(file, sep=None, engine='python', encoding=enc, on_bad_lines='skip')
        except Exception:
            continue
    return None

def load_stock(file):
    """Carga el CSV de stock con separador ; y encoding latin-1."""
    if file is None:
        return None
    for enc in ['utf-8', 'latin-1', 'iso-8859-1']:
        try:
            file.seek(0)
            df = pd.read_csv(file, sep=';', encoding=enc, on_bad_lines='skip')
            if len(df.columns) > 1:
                return df
        except Exception:
            pass
        try:
            file.seek(0)
            df = pd.read_csv(file, sep=',', encoding=enc, on_bad_lines='skip')
            if len(df.columns) > 1:
                return df
        except Exception:
            pass
    return None

def clean_sku(sku):
    """
    Elimina el prefijo FR del SKU conservando los ceros siguientes,
    para que el formato coincida con Referencia del stock (FR02748 -> 02748).
    """
    if pd.isna(sku):
        return ""
    s = str(sku).strip()
    if s.upper().startswith("FR"):
        s = s[2:]
    return s  # NO lstrip('0'): el stock usa 02748, no 2748

def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Datos')
    return output.getvalue()

def today_str():
    return date.today().strftime("%Y%m%d")

def parse_listar(text):
    """Extrae Num_Pedido e Id_Envio de una fila de ListarRecogida."""
    text = str(text).strip()
    m_pedido = re.match(r'(\d{3}-\d{7}-\d{7})', text)
    n_pedido = m_pedido.group(1) if m_pedido else text.split()[0]
    m_envio  = re.search(r'ID de env[íi]o:\s*([A-Za-z0-9]+)', text)
    id_envio = m_envio.group(1) if m_envio else ""
    return pd.Series([n_pedido, id_envio])

# ──────────────────────────────────────────────
# TABS DE FASES
# ──────────────────────────────────────────────

tab_fase1, tab_fase2 = st.tabs(["🚀 Fase 1 — Subida Cecopartners", "🏭 Fase 2 — D-PEDIDOS Almacén Francia"])

# ══════════════════════════════════════════════
# FASE 1
# ══════════════════════════════════════════════
with tab_fase1:
    st.header("Fase 1: Generar fichero de subida y cancelaciones")

    with st.sidebar:
        st.header("📂 Fase 1 — Ficheros")
        stock_file     = st.file_uploader("1. Stock FR (CSV)",                   type=["csv", "txt"],  key="stock")
        pedidos_file   = st.file_uploader("2. PedidosRecoger (Excel/CSV)",       type=["csv", "xlsx"], key="pedidos")
        listar_file    = st.file_uploader("3. ListarRecogida (Excel/CSV)",       type=["csv", "xlsx"], key="listar")
        plantilla_file = st.file_uploader("4. Plantilla SELLER_FLEX_FR (Excel)", type=["csv", "xlsx"], key="plantilla")

    if st.button("▶️ Procesar Fase 1", type="primary", key="btn_fase1"):
        if not all([stock_file, pedidos_file, listar_file, plantilla_file]):
            st.error("Sube los 4 ficheros de la barra lateral para continuar.")
        else:
            with st.spinner("Procesando..."):
                try:
                    df_stock     = load_stock(stock_file)
                    df_pedidos   = load_excel_or_csv(pedidos_file)
                    df_listar    = load_excel_or_csv(listar_file)
                    df_plantilla = load_excel_or_csv(plantilla_file)

                    if df_stock is None or df_stock.empty:
                        st.error("No se pudo leer el fichero de stock.")
                        st.stop()

                    # ── 1. PARSEAR ListarRecogida ──
                    col_lr = df_listar.columns[0]
                    df_listar[['Num_Pedido', 'Id_Envio']] = df_listar[col_lr].apply(parse_listar)
                    mapa_envio_pedido = dict(zip(
                        df_listar['Id_Envio'].str.strip(),
                        df_listar['Num_Pedido'].str.strip()
                    ))

                    # ── 2. LIMPIAR PedidosRecoger ──
                    df_pedidos['SKU_Clean']     = df_pedidos['SKU'].apply(clean_sku)
                    df_pedidos['IdPedido_Clean']= df_pedidos['Identificador de pedido'].astype(str).str.strip()
                    df_pedidos['Num_Pedido']    = df_pedidos['IdPedido_Clean'].map(mapa_envio_pedido).fillna("")

                    # ── 3. CRUZAR CON STOCK ──
                    col_ref = df_stock.columns[0]
                    df_stock[col_ref] = df_stock[col_ref].astype(str).str.strip()

                    # Usar solo refs con stock disponible > 0 si existe esa columna
                    if 'StockDisponible' in df_stock.columns:
                        refs_con_stock = set(
                            df_stock[df_stock['StockDisponible'].fillna(0) > 0][col_ref].unique()
                        )
                    else:
                        refs_con_stock = set(df_stock[col_ref].unique())

                    df_pedidos['Disponible'] = df_pedidos['SKU_Clean'].isin(refs_con_stock)

                    df_ok     = df_pedidos[df_pedidos['Disponible']].copy()
                    df_cancel = df_pedidos[~df_pedidos['Disponible']].copy()

                    # ── 4. FICHERO SUBIDA CECOPARTNERS ──
                    if not df_ok.empty:
                        subida_data = {
                            'article':               df_ok['SKU_Clean'].tolist(),
                            'quantity':              df_ok['Unidades'].fillna(1).astype(int).tolist(),
                            'customer_name':         ['AMAZON FLEX'] * len(df_ok),
                            'nif':                   [''] * len(df_ok),
                            'attention_of_customer': ['AMAZON FLEX'] * len(df_ok),
                            'address':               ['Cam.Real de Madrid 117'] * len(df_ok),
                            'postal_code':           [46292] * len(df_ok),
                            'phone':                 [0] * len(df_ok),
                            'city':                  ['MASSALAVÉS'] * len(df_ok),
                            'country_code':          ['ES'] * len(df_ok),
                            'comment':               [0] * len(df_ok),
                            'addressee_order_number': df_ok['Num_Pedido'].tolist(),
                        }
                        df_subida = pd.DataFrame(subida_data)
                        df_subida['customer_mail'] = (
                            df_subida['addressee_order_number'].astype(str) + '@sellerflexfr.com'
                        )
                        for col in df_plantilla.columns:
                            if col not in df_subida.columns:
                                df_subida[col] = ''
                        df_subida = df_subida[df_plantilla.columns]
                    else:
                        df_subida = pd.DataFrame(columns=df_plantilla.columns)

                    # ── 5. FICHERO CANCELACIONES ──
                    asin_col  = next(
                        (c for c in ['ASIN', 'FNSKU', 'asin', 'fnsku'] if c in df_cancel.columns),
                        None
                    )
                    asin_vals = df_cancel[asin_col].tolist() if asin_col else [''] * len(df_cancel)

                    if not df_cancel.empty:
                        df_cancelaciones = pd.DataFrame({
                            'Node ID':      ['SRAN'] * len(df_cancel),
                            'Order number': df_cancel['Num_Pedido'].tolist(),
                            'Shipment ID':  df_cancel['IdPedido_Clean'].tolist(),
                            'ASIN':         asin_vals,
                            'Reason':       ['OOO'] * len(df_cancel),
                        })
                    else:
                        df_cancelaciones = pd.DataFrame(
                            columns=['Node ID', 'Order number', 'Shipment ID', 'ASIN', 'Reason']
                        )

                    # Guardar también el mapa para la Fase 2
                    st.session_state['mapa_envio_pedido'] = mapa_envio_pedido
                    st.session_state['df_pedidos_ok']     = df_ok
                    st.session_state['df_subida']         = df_subida
                    st.session_state['df_cancelaciones']  = df_cancelaciones
                    st.session_state['procesado_f1']      = True

                except Exception as e:
                    st.error(f"Error procesando los ficheros: {e}")
                    import traceback; st.code(traceback.format_exc())

    # ── MOSTRAR RESULTADOS FASE 1 ──
    if st.session_state.get('procesado_f1'):
        df_subida        = st.session_state['df_subida']
        df_cancelaciones = st.session_state['df_cancelaciones']

        st.success(
            f"✅ Procesado: **{len(df_subida)} pedidos aptos** · **{len(df_cancelaciones)} cancelaciones**"
        )

        res1, res2 = st.tabs(["📤 Subida Cecopartners", "❌ Cancelaciones (OOO)"])

        with res1:
            st.subheader(f"Fichero Subida Cecopartners ({len(df_subida)} líneas)")
            st.dataframe(df_subida, use_container_width=True)
            st.download_button(
                "📥 Descargar Subida Cecopartners (Excel)",
                data=to_excel(df_subida),
                file_name=f"SELLER_FLEX_FR_{today_str()}.xlsx",
                mime="application/vnd.ms-excel",
                key="dl_subida"
            )

        with res2:
            st.subheader(f"Cancelaciones OOO ({len(df_cancelaciones)} líneas)")
            if df_cancelaciones.empty:
                st.info("🎉 No hay pedidos sin stock.")
            else:
                st.dataframe(df_cancelaciones, use_container_width=True)
                st.download_button(
                    "📥 Descargar Cancelaciones (Excel)",
                    data=to_excel(df_cancelaciones),
                    file_name=f"{today_str()}_CancelOrders_SellerFlexFR.xlsx",
                    mime="application/vnd.ms-excel",
                    key="dl_cancel"
                )
    else:
        st.info("👆 Sube los 4 ficheros en la barra lateral y pulsa **Procesar Fase 1**.")


# ══════════════════════════════════════════════
# FASE 2
# ══════════════════════════════════════════════
with tab_fase2:
    st.header("Fase 2: Generar D-PEDIDOS para Almacén Francia")
    st.markdown("""
    Sube el fichero descargado de Cecopartners (con la REFERENCIA `D-XXXXXX-N SGA` ya asignada)
    junto con ListarRecogida y PedidosRecoger. Se generará el **D-PEDIDOS** con esta estructura:

    | Col | Cabecera | Fuente |
    |---|---|---|
    | A | **P** | Zona (P…) de PedidosRecoger, cruzada por Shipment ID |
    | B | **U** | Shipment ID — cruzado via ListarRecogida por nº pedido |
    | C | **Agencia** | `AMZN_FR_SH_SD` (siempre fijo) |
    | D… | *(resto del fichero Cecopartners)* | Todas las columnas tal cual |

    El cruce: `NÚMERO DE LÍNEA DE PEDIDO DE CLIENTE` (Cecopartners) → ListarRecogida → Shipment ID + Zona P
    """)

    col1, col2 = st.columns(2)
    with col1:
        ceco_result_file = st.file_uploader(
            "📄 Fichero descargado de Cecopartners (con REFERENCIA D-…)",
            type=["csv", "xlsx"], key="ceco_result"
        )
    with col2:
        listar_f2_file = st.file_uploader(
            "📄 ListarRecogida (col A: nº pedido + ID de envío)",
            type=["csv", "xlsx"], key="listar_f2"
        )
        pedidos_f2_file = st.file_uploader(
            "📄 PedidosRecoger (para obtener la zona P…)",
            type=["csv", "xlsx"], key="pedidos_f2"
        )

    if st.button("▶️ Generar D-PEDIDOS", type="primary", key="btn_fase2"):
        if not all([ceco_result_file, listar_f2_file, pedidos_f2_file]):
            st.error("Sube los 3 ficheros para continuar.")
        else:
            with st.spinner("Generando D-PEDIDOS..."):
                try:
                    df_ceco     = load_excel_or_csv(ceco_result_file)
                    df_listar2  = load_excel_or_csv(listar_f2_file)
                    df_pedidos2 = load_excel_or_csv(pedidos_f2_file)

                    # ── 1. Parsear ListarRecogida: col A → Num_Pedido + Id_Envio ──
                    # Formato: "406-8213474-0209101 ID de envío: Tcg94cqKD   Listo..."
                    col_lr2 = df_listar2.columns[0]
                    df_listar2[['Num_Pedido', 'Id_Envio']] = df_listar2[col_lr2].apply(parse_listar)

                    # Mapa num_pedido → Shipment ID  (para rellenar columna U)
                    mapa_pedido_shipment = dict(zip(
                        df_listar2['Num_Pedido'].str.strip(),
                        df_listar2['Id_Envio'].str.strip()
                    ))

                    # ── 2. Obtener zona P de PedidosRecoger ──
                    # Columna Zona (P…) cruzada por Identificador de pedido (Shipment ID)
                    zona_col = next(
                        (c for c in df_pedidos2.columns if 'zona' in c.lower()),
                        df_pedidos2.columns[0]
                    )
                    id_col2 = next(
                        (c for c in df_pedidos2.columns if 'identificador de pedido' in c.lower()),
                        None
                    )
                    if id_col2 is None:
                        # Fallback: columna con valores que empiezan por T (Shipment IDs)
                        for c in df_pedidos2.columns:
                            sample = df_pedidos2[c].dropna().astype(str).head(5).tolist()
                            if any(v.startswith('T') and len(v) > 5 for v in sample):
                                id_col2 = c
                                break

                    # Mapa Shipment ID → Zona P
                    mapa_shipment_zona = dict(zip(
                        df_pedidos2[id_col2].astype(str).str.strip(),
                        df_pedidos2[zona_col].astype(str).str.strip()
                    ))

                    # ── 3. Localizar columna clave en Cecopartners ──
                    # "NÚMERO DE LÍNEA DE PEDIDO DE CLIENTE" contiene el nº pedido (406-XXXXX)
                    linea_col = next(
                        (c for c in df_ceco.columns
                         if 'línea de pedido de cliente' in c.lower()
                         or 'linea de pedido de cliente' in c.lower()),
                        None
                    )
                    if linea_col is None:
                        st.error(
                            f"No se encontró 'NÚMERO DE LÍNEA DE PEDIDO DE CLIENTE' en Cecopartners.\n"
                            f"Columnas: {df_ceco.columns.tolist()}"
                        )
                        st.stop()

                    # ── 4. Calcular U y P para cada fila de Cecopartners ──
                    num_pedidos = df_ceco[linea_col].astype(str).str.strip()
                    shipment_ids = num_pedidos.map(mapa_pedido_shipment).fillna("")
                    zonas        = shipment_ids.map(mapa_shipment_zona).fillna("")

                    # ── 5. Montar D-PEDIDOS: P | U | Agencia | (cols Cecopartners sin la primera FALSE) ──
                    cols_ceco = list(df_ceco.columns[1:])  # saltar col A (FALSE/checkbox)
                    df_dpedidos = pd.DataFrame({'P': zonas, 'U': shipment_ids, 'Agencia': 'AMZN_FR_SH_SD'})
                    for c in cols_ceco:
                        df_dpedidos[c] = df_ceco[c].values

                    # ── 6. Diagnóstico de cruces no encontrados ──
                    sin_cruce = df_dpedidos[df_dpedidos['U'] == ""]
                    if not sin_cruce.empty:
                        st.warning(
                            f"⚠️ {len(sin_cruce)} fila(s) sin Shipment ID en ListarRecogida "
                            f"(nº pedido no encontrado). Revisa que los ficheros sean del mismo día."
                        )

                    st.session_state['df_dpedidos']  = df_dpedidos
                    st.session_state['procesado_f2'] = True

                except Exception as e:
                    st.error(f"Error generando D-PEDIDOS: {e}")
                    import traceback; st.code(traceback.format_exc())

    # ── MOSTRAR RESULTADOS FASE 2 ──
    if st.session_state.get('procesado_f2'):
        df_dp = st.session_state['df_dpedidos']
        st.success(f"✅ D-PEDIDOS generado: **{len(df_dp)} líneas**")
        st.dataframe(df_dp, use_container_width=True)
        st.download_button(
            "📥 Descargar D-PEDIDOS Almacén (Excel)",
            data=to_excel(df_dp),
            file_name=f"D-PEDIDOS_FLEX_FR_{today_str()}.xlsx",
            mime="application/vnd.ms-excel",
            key="dl_dpedidos"
        )
    else:
        st.info("👆 Sube los 3 ficheros y pulsa **Generar D-PEDIDOS**.")