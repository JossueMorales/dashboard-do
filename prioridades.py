import streamlit as st
import pandas as pd
import datetime
import urllib.request
import urllib.parse
import openpyxl

# ¡URL TATUADA EN EL CÓDIGO! Ya no lee archivos viejos.
URL_OFICIAL_APPS_SCRIPT = "https://script.google.com/macros/s/AKfycbx0byaIIxNbFxyHuh28moJcagPbtavg27XHRNQ-qAFVBrDbNR0qFK7qL2YYfEih6gZb/exec"

def enviar_actualizacion(url, row, col, value):
    import requests
    payload = {"row": row, "column": col, "value": value, "sheetName": "Junta de Seguimiento"}
    try:
        res = requests.post(url, json=payload)
        if res.status_code == 200:
            result = res.json()
            if result.get("status") == "success":
                return True
    except:
        pass
    return False

def limpiar_celda(val):
    if pd.isna(val) or val is None:
        return ""
    if isinstance(val, (datetime.datetime, pd.Timestamp)):
        return val.strftime('%Y-%m-%d')
    if isinstance(val, (int, float)):
        if val == int(val): return str(int(val))
        return str(val)
    val_str = str(val).strip()
    if val_str.lower() in ["nan", "none"]: return ""
    if val_str.endswith(" 00:00:00"): return val_str.replace(" 00:00:00", "")
    return val_str

st.set_page_config(page_title="Prioridades DO 2026", layout="wide", page_icon="🎯")
st.title("🎯 Centro de Control: Acuerdos y Prioridades DO 2026")
st.markdown("---")

@st.cache_data(ttl=5)
def cargar_datos_nuevos():
    url_xlsx = "https://docs.google.com/spreadsheets/d/16vDONi15PWK-TTMxDpCCHoV-JFsH1z29PWrxYNhxedU/export?format=xlsx"
    xlsx_path = "sheet_cache.xlsx"
    urllib.request.urlretrieve(url_xlsx, xlsx_path)
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    sheet = wb["Junta de Seguimiento"]
    data = []
    hyperlinks = {}
    for r_idx, row in enumerate(sheet.iter_rows(values_only=False), start=1):
        row_data = []
        for c_idx, cell in enumerate(row, start=1):
            val = cell.value
            if isinstance(val, (int, float)) and cell.number_format and '%' in cell.number_format:
                val = val * 100
            row_data.append(val)
            if cell.hyperlink and cell.hyperlink.target:
                hyperlinks[(r_idx - 1, c_idx - 1)] = cell.hyperlink.target
        data.append(row_data)
    return pd.DataFrame(data), hyperlinks

try:
    df_raw, hyperlinks = cargar_datos_nuevos()
    
    # ACUERDOS
    df_acuerdos = df_raw.iloc[4:12, 2:5].copy() 
    df_acuerdos.columns = ["Acuerdos", "Actividad", "Fecha Actividad"]
    df_acuerdos = df_acuerdos.dropna(how="all")
    df_acuerdos = df_acuerdos[df_acuerdos["Actividad"].astype(str).str.strip().str.lower() != "none"]
    df_acuerdos = df_acuerdos[df_acuerdos["Actividad"].astype(str).str.strip() != "nan"]
    df_acuerdos = df_acuerdos[df_acuerdos["Actividad"].astype(str).str.strip() != ""]

    # PRIORIDADES
    fila_encabezado_p = 14
    for idx, row in df_raw.iterrows():
        if any('minuta' in str(x).lower() or 'prioridad' in str(x).lower() for x in row.values) and idx > 10:
            fila_encabezado_p = idx
            break
            
    df_prioridades = df_raw.iloc[fila_encabezado_p+1:].copy()
    df_prioridades.columns = df_raw.iloc[fila_encabezado_p].astype(str).str.strip()
    df_prioridades = df_prioridades.loc[:, df_prioridades.columns.notna() & (df_prioridades.columns != "nan") & (df_prioridades.columns != "none") & (df_prioridades.columns != "")]
    df_prioridades = df_prioridades.loc[:, ~df_prioridades.columns.astype(str).str.contains('^Unnamed|^None', case=False)]
    
    col_id = df_prioridades.columns[0]
    df_prioridades = df_prioridades[df_prioridades[col_id].astype(str).str.strip().str.lower() != "none"]
    df_prioridades = df_prioridades[df_prioridades[col_id].astype(str).str.strip() != "nan"]
    df_prioridades = df_prioridades[df_prioridades[col_id].astype(str).str.strip() != ""]

    col_link_idx = None
    for c_idx, col_name in enumerate(df_raw.iloc[fila_encabezado_p]):
        if str(col_name).strip().lower() == 'link':
            col_link_idx = c_idx
            break

    urls_kanban_raw = {}
    if col_link_idx is not None:
        col_link_name = df_raw.iloc[fila_encabezado_p][col_link_idx]
        for idx in df_prioridades.index:
            val_link = df_prioridades.loc[idx, col_link_name]
            if (idx, col_link_idx) in hyperlinks and hyperlinks[(idx, col_link_idx)]:
                urls_kanban_raw[idx] = hyperlinks[(idx, col_link_idx)]
            elif val_link and "http" in str(val_link):
                urls_kanban_raw[idx] = str(val_link).strip()
            else:
                urls_kanban_raw[idx] = "https://docs.google.com/spreadsheets/d/16vDONi15PWK-TTMxDpCCHoV-JFsH1z29PWrxYNhxedU/edit"

        sheet_edit_url = "https://docs.google.com/spreadsheets/d/16vDONi15PWK-TTMxDpCCHoV-JFsH1z29PWrxYNhxedU/edit"
        df_prioridades[col_link_name] = [
            f"{hyperlinks[(idx, col_link_idx)]}#🔗 {val}" if (idx, col_link_idx) in hyperlinks and hyperlinks[(idx, col_link_idx)]
            else f"{sheet_edit_url}#🔗 {val}" if val and str(val).strip().lower() != "none" and str(val).strip() != ""
            else ""
            for idx, val in zip(df_prioridades.index, df_prioridades[col_link_name])
        ]

    st.sidebar.header("⚙️ Filtros de Búsqueda")
    
    col_tipo = [c for c in df_prioridades.columns if 'tipo' in c.lower()]
    if col_tipo:
        nombre_tipo = col_tipo[0]
        tipos = [x for x in df_prioridades[nombre_tipo].dropna().unique().tolist() if str(x).strip() != "" and str(x).lower() != "nan" and str(x).lower() != "none"]
        if tipos:
            tipo_sel = st.sidebar.multiselect(f"Filtrar por {nombre_tipo}:", tipos, default=tipos)
            df_prioridades_filtrado = df_prioridades[df_prioridades[nombre_tipo].isin(tipo_sel)]
        else:
            df_prioridades_filtrado = df_prioridades.copy()
    else:
        df_prioridades_filtrado = df_prioridades.copy()

    col_estatus = [c for c in df_prioridades.columns if 'estatus' in c.lower() or 'estado' in c.lower()]
    estatus_sel = []
    if col_estatus:
        nombre_estatus = col_estatus[0]
        estatus_list = [x for x in df_prioridades[nombre_estatus].dropna().unique().tolist() if str(x).strip() != "" and str(x).lower() != "nan" and str(x).lower() != "none"]
        if estatus_list:
            estatus_sel = st.sidebar.multiselect(f"Filtrar por {nombre_estatus}:", estatus_list, default=estatus_list)
            df_prioridades_filtrado = df_prioridades_filtrado[df_prioridades_filtrado[nombre_estatus].isin(estatus_sel)]
            
    st.sidebar.markdown("---")
    st.sidebar.subheader("🔌 Conexión Activa")
    st.sidebar.success("Conectado a Google Workspace (Seguro)")
    st.sidebar.link_button("✏️ Editar Archivo Original", "https://docs.google.com/spreadsheets/d/16vDONi15PWK-TTMxDpCCHoV-JFsH1z29PWrxYNhxedU/edit")

    st.header("📝 1. Acuerdos de la Última Junta")
    config_columnas_a = {}
    for col in df_acuerdos.columns:
        df_acuerdos[col] = df_acuerdos[col].apply(limpiar_celda)
        if df_acuerdos[col].str.contains("http://|https://", case=False).any():
            config_columnas_a[col] = st.column_config.LinkColumn(col, display_text="🔗 Abrir Archivo")
            
    st.data_editor(df_acuerdos, use_container_width=True, hide_index=True, column_config=config_columnas_a, key="editor_acuerdos")
    st.markdown("<br><hr><br>", unsafe_allow_html=True)
    
    st.header("📊 2. Prioridades / Relevantes Operativos")
    col1, col2, col3 = st.columns(3)
    with col1: st.metric(label="Total de Prioridades", value=len(df_prioridades_filtrado))
    with col2:
        col_pct = [c for c in df_prioridades.columns if '%' in c or 'avance' in c.lower()]
        if col_pct:
            nombre_pct = col_pct[0]
            df_prioridades_filtrado[nombre_pct] = pd.to_numeric(df_prioridades_filtrado[nombre_pct].astype(str).str.replace('[\%,]', '', regex=True), errors='coerce').fillna(0)
            promedio_avance = int(df_prioridades_filtrado[nombre_pct].mean()) if not df_prioridades_filtrado.empty else 0
            st.metric(label="Avance Promedio Global", value=f"{promedio_avance}%")
        else:
            st.metric(label="Avance Promedio Global", value="N/A")
    with col3:
        if col_estatus:
            nombre_estatus = col_estatus[0]
            en_proceso = len(df_prioridades_filtrado[df_prioridades_filtrado[nombre_estatus].astype(str).str.contains("Proceso|Activo|Iniciado|Aprobación", case=False, na=False)])
            st.metric(label="Tareas Activas", value=en_proceso)
            
    st.markdown("<br>", unsafe_allow_html=True)
    tab_tabla, tab_kanban = st.tabs(["🗂️ Vista de Tabla (Editable)", "📋 Tablero Kanban"])

    with tab_tabla:
        df_tabla_display = df_prioridades_filtrado.copy()
        for col in df_tabla_display.columns:
            if col_pct and col == col_pct[0]: continue
            df_tabla_display[col] = df_tabla_display[col].apply(limpiar_celda)
            
        config_columnas_p = {}
        if col_pct:
            config_columnas_p[col_pct[0]] = st.column_config.ProgressColumn(col_pct[0], format="%d%%", min_value=0, max_value=100)
            
        for col in df_tabla_display.columns:
            if df_tabla_display[col].astype(str).str.contains("http://|https://", case=False).any():
                if df_tabla_display[col].astype(str).str.contains("#", case=False).any():
                    config_columnas_p[col] = st.column_config.LinkColumn(col, display_text=r"^.*#(.*)$")
                else:
                    config_columnas_p[col] = st.column_config.LinkColumn(col, display_text="🔗 Abrir Archivo")
            
        st.data_editor(df_tabla_display, use_container_width=True, hide_index=True, column_config=config_columnas_p, key="editor_prioridades")

    with tab_kanban:
        nom_estatus = [c for c in df_prioridades.columns if 'estatus' in c.lower()][0]
        nom_entregable = [c for c in df_prioridades.columns if 'entregable' in c.lower() or 'proyecto' in c.lower()][0]
        nom_minuta = [c for c in df_prioridades.columns if 'minuta' in c.lower() or 'actividad' in c.lower()][0]
        nom_comentarios = [c for c in df_prioridades.columns if 'comentario' in c.lower() or 'compromiso' in c.lower()][0]
        
        excel_col_comentarios = None
        for c_idx, name in enumerate(df_raw.iloc[fila_encabezado_p]):
            if str(name).strip() == nom_comentarios:
                excel_col_comentarios = c_idx + 1
                break
        
        columnas_kanban = estatus_sel if estatus_sel else [e for e in df_prioridades[nom_estatus].unique() if str(e).strip() != "" and str(e).lower() != "nan"]
        
        if columnas_kanban:
            cols = st.columns(len(columnas_kanban))
            for i, estado in enumerate(columnas_kanban):
                with cols[i]:
                    st.markdown(f"<h3 style='text-align: center; color: #1E90FF;'>📂 {estado}</h3>", unsafe_allow_html=True)
                    st.markdown("---")
                    
                    df_estado = df_prioridades_filtrado[df_prioridades_filtrado[nom_estatus] == estado]
                    for idx, row in df_estado.iterrows():
                        pct = row.get(nombre_pct, 0) if col_pct else 0
                        tipo = row.get(nombre_tipo, "") if col_tipo else ""
                        comentarios_txt = limpiar_celda(row.get(nom_comentarios, "Sin comentarios."))
                        
                        with st.expander(f"📌 **{row[nom_minuta]}**\n\n*{row[nom_entregable]}* 📊 {int(pct)}%"):
                            st.markdown(f"**🔹 Tipo:** {tipo}")
                            
                            nuevo_comentario = st.text_area(
                                "📝 Detalles (Modifica Excel):", value=comentarios_txt, key=f"txt_{idx}", height=80
                            )
                            
                            if nuevo_comentario != comentarios_txt:
                                if st.button("💾 Guardar en Excel", key=f"btn_save_{idx}", use_container_width=True):
                                    if excel_col_comentarios:
                                        with st.spinner("Guardando..."):
                                            if enviar_actualizacion(URL_OFICIAL_APPS_SCRIPT, idx + 1, excel_col_comentarios, nuevo_comentario):
                                                st.success("✅ ¡Excel actualizado!")
                                                st.cache_data.clear()
                                                st.rerun()
                            
                            notas_personales = st.text_area(
                                "🔒 Mis Notas para la Tarea:", value="", placeholder="Notas adicionales solo para ti...", key=f"notas_local_{idx}", height=80
                            )
                            
                            st.markdown("---")
                            target_url = urls_kanban_raw.get(idx, "https://docs.google.com/spreadsheets/d/16vDONi15PWK-TTMxDpCCHoV-JFsH1z29PWrxYNhxedU/edit")
                            
                            titulo_tarea = f"📌 {row[nom_minuta]} ({row[nom_entregable]})"
                            detalles_tarea = f"Detalles: {nuevo_comentario}\n\nNotas: {notas_personales}\n\nLink: {target_url}"
                            
                            # AQUI GENERAMOS EL LINK AL APPS SCRIPT CORRECTO SIEMPRE
                            params = urllib.parse.urlencode({"action": "createTask", "title": titulo_tarea, "notes": detalles_tarea})
                            tasks_url = f"{URL_OFICIAL_APPS_SCRIPT}?{params}"
                            
                            b_col1, b_col2 = st.columns(2)
                            with b_col1: st.link_button("🔗 Ver Archivo", target_url, use_container_width=True)
                            with b_col2: st.link_button("📥 Crear Tarea (Real)", tasks_url, use_container_width=True)
        else:
            st.warning("Selecciona al menos un estatus en los filtros laterales para visualizar el tablero Kanban.")
    
    actualizado = False
    if "editor_acuerdos" in st.session_state:
        edits_a = st.session_state.editor_acuerdos.get("edited_rows", {})
        if edits_a:
            col_mapping_acuerdos = {"Acuerdos": 3, "Actividad": 4, "Fecha Actividad": 5}
            for row_offset, col_changes in edits_a.items():
                original_row_idx = df_acuerdos.index[row_offset]
                excel_row = original_row_idx + 1
                for col_name, new_val in col_changes.items():
                    excel_col = col_mapping_acuerdos.get(col_name)
                    if excel_col:
                        with st.spinner("Sincronizando..."):
                            if enviar_actualizacion(URL_OFICIAL_APPS_SCRIPT, excel_row, excel_col, new_val): actualizado = True
                                
    if "editor_prioridades" in st.session_state:
        edits_p = st.session_state.editor_prioridades.get("edited_rows", {})
        if edits_p:
            for row_offset, col_changes in edits_p.items():
                original_row_idx = df_prioridades_filtrado.index[row_offset]
                excel_row = original_row_idx + 1
                for col_name, new_val in col_changes.items():
                    excel_col = None
                    for c_idx, name in enumerate(df_raw.iloc[fila_encabezado_p]):
                        if str(name).strip() == col_name:
                            excel_col = c_idx + 1
                            break
                    if excel_col:
                        with st.spinner("Sincronizando..."):
                            if enviar_actualizacion(URL_OFICIAL_APPS_SCRIPT, excel_row, excel_col, new_val): actualizado = True
                                
    if actualizado:
        st.success("✅ ¡Google Sheets actualizado con éxito!")
        st.cache_data.clear()
        st.rerun()

except Exception as e:
    st.error("Ocurrió un detalle al intentar procesar las celdas de tu nuevo Google Sheet.")
    st.warning(f"Error técnico: {e}")