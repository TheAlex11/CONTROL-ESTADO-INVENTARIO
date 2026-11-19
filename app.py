import streamlit as st
import pandas as pd
import openpyxl
from datetime import datetime
import os
import sqlite3

# ==============================================================================
# CONFIGURACIÓN GENERAL Y SEGURIDAD
# ==============================================================================
st.set_page_config(
    page_title="Control de Estado de Activos | Tecnología",
    page_icon="🖥️",
    layout="wide",
    initial_sidebar_state="collapsed"
)

EXCEL_PATH = "ESTADO_TIENDAS 2026.xlsx"
SHEET_NAME = "Control de Equipos"
DB_PATH = "historial_reportes.db"
ADMIN_PASSWORD = "Rivier@25"  # <-- Cambia esta clave por la contraseña que desees

ESTADOS_GESTION_TI = [
    "🔴 Pendiente / Abierto",
    "🟡 En Revisión / Diagnóstico",
    "🟠 Pendiente de Repuesto / Compra",
    "🔵 Pendiente de Cambio / Envío de Equipo",
    "🟢 Solucionado / Operativo"
]

# Periféricos que NO manejan placa de Activo Fijo
PERIFERICOS_SIN_ACTIVO = {
    "Teclado",
    "Mouse",
    "Pantalla",
    "Cables (Poder/Vid/Red)",
    "PadMouse"
}

# ==============================================================================
# MAPEO EXACTO DE COLUMNAS DEL EXCEL (ESTADO_TIENDAS 2026.xlsx)
# ==============================================================================
MAPEO_EQUIPOS = {
    "POS I (Caja 1)": {
        "Computador": {"activo_col": "D", "estado_col": "E"},
        "Teléfono": {"activo_col": "F", "estado_col": "G"},
        "Biométrico": {"activo_col": "H", "estado_col": "I"},
        "Lector de Código": {"activo_col": "J", "estado_col": "K"},
        "Cajón Monedero": {"activo_col": "L", "estado_col": "M"},
        "Impresora POS": {"activo_col": "N", "estado_col": "O"},
        "Teclado": {"activo_col": None, "estado_col": "X"},
        "Mouse": {"activo_col": None, "estado_col": "Y"},
        "Pantalla": {"activo_col": None, "estado_col": "Z"},
    },
    "POS II (Caja 2)": {
        "Computador": {"activo_col": "P", "estado_col": "Q"},
        "Lector de Código": {"activo_col": "R", "estado_col": "S"},
        "Cajón Monedero": {"activo_col": "T", "estado_col": "U"},
        "Impresora POS": {"activo_col": "V", "estado_col": "W"},
        "Teclado": {"activo_col": None, "estado_col": "AA"},
        "Mouse": {"activo_col": None, "estado_col": "AB"},
        "Pantalla": {"activo_col": None, "estado_col": "AC"},
    },
    "Equipos Generales e Infraestructura": {
        "Impresora Láser": {"activo_col": "AD", "estado_col": "AE"},
        "Pantalla Publicitaria I": {"activo_col": "AF", "estado_col": "AG"},
        "Pantalla Publicitaria II": {"activo_col": "AH", "estado_col": "AI"},
        "Equipo de Sonido": {"activo_col": "AJ", "estado_col": "AK"},
        "Cables (Poder/Vid/Red)": {"activo_col": None, "estado_col": "AL"},
        "PadMouse": {"activo_col": None, "estado_col": "AM"},
        "AP (Access Point)": {"activo_col": "AN", "estado_col": "AO"},
        "SWITCH": {"activo_col": "AP", "estado_col": "AQ"},
        "Monitor de Cámaras": {"activo_col": "AR", "estado_col": "AS"},
    },
    "Caja Adicional (Caja 3 / 4 / 5)": {
        "Computador": {"activo_col": None, "estado_col": None},
        "Pantalla": {"activo_col": None, "estado_col": None},
        "Teclado": {"activo_col": None, "estado_col": None},
        "Mouse": {"activo_col": None, "estado_col": None},
        "Lector de Código": {"activo_col": None, "estado_col": None},
        "Cajón Monedero": {"activo_col": None, "estado_col": None},
        "Impresora POS": {"activo_col": None, "estado_col": None},
    }
}

# ==============================================================================
# BASE DE DATOS DE AUDITORÍA Y GESTIÓN TI (SQLITE)
# ==============================================================================
def init_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS reportes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha_hora TEXT,
            fila_excel INTEGER,
            c_costos TEXT,
            tienda TEXT,
            ciudad TEXT,
            ubicacion TEXT,
            caja_extra TEXT,
            equipo TEXT,
            estado TEXT,
            activo_fijo TEXT,
            observacion TEXT,
            reportado_por TEXT,
            estado_ti TEXT DEFAULT '🔴 Pendiente / Abierto',
            solucion_ti TEXT DEFAULT '',
            fecha_cierre_ti TEXT DEFAULT ''
        )
    """)
    columnas_existentes = [col[1] for col in c.execute("PRAGMA table_info(reportes)").fetchall()]
    nuevas_cols = {
        "fila_excel": "INTEGER",
        "estado_ti": "TEXT DEFAULT '🔴 Pendiente / Abierto'",
        "solucion_ti": "TEXT DEFAULT ''",
        "fecha_cierre_ti": "TEXT DEFAULT ''"
    }
    for col_name, col_def in nuevas_cols.items():
        if col_name not in columnas_existentes:
            c.execute(f"ALTER TABLE reportes ADD COLUMN {col_name} {col_def}")
    conn.commit()
    conn.close()

def guardar_log_db(data):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    estado_inicial_ti = (
        "🟢 Solucionado / Operativo" if data["estado"] == "BUENO" else "🔴 Pendiente / Abierto"
    )
    c.execute("""
        INSERT INTO reportes (
            fecha_hora, fila_excel, c_costos, tienda, ciudad, ubicacion, caja_extra,
            equipo, estado, activo_fijo, observacion, reportado_por, estado_ti
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        data["fecha_hora"], data["fila_excel"], data["c_costos"], data["tienda"],
        data["ciudad"], data["ubicacion"], data["caja_extra"], data["equipo"],
        data["estado"], data["activo_fijo"], data["observacion"],
        data["reportado_por"], estado_inicial_ti
    ))
    conn.commit()
    conn.close()

# ==============================================================================
# FUNCIONES DE LECTURA Y ESCRITURA SOBRE EXCEL MAESTRO
# ==============================================================================
def obtener_tiendas_desde_excel():
    """Lee dinámicamente las tiendas (filas 8 a 42) del formato oficial."""
    if not os.path.exists(EXCEL_PATH):
        return []
    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
    ws = wb[SHEET_NAME]
    tiendas = []
    for r in range(8, 43):
        cc = ws.cell(r, 1).value
        nombre = ws.cell(r, 2).value
        ciudad = ws.cell(r, 3).value
        if nombre:
            tiendas.append({
                "fila": r,
                "c_costos": str(cc) if cc is not None else "",
                "tienda": str(nombre).strip(),
                "ciudad": str(ciudad).strip() if ciudad else ""
            })
    wb.close()
    return tiendas

def actualizar_excel_maestro(fila_tienda, ubicacion, caja_extra, equipo, estado, activo_fijo, observacion, reportado_por, es_cierre_ti=False):
    """Actualiza la celda exacta de Estado, Activo Fijo y Observaciones en el .xlsx"""
    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb[SHEET_NAME]

    config_col = MAPEO_EQUIPOS[ubicacion][equipo]
    estado_col = config_col["estado_col"]
    activo_col = config_col["activo_col"]

    # 1. Actualizar columna de Estado si existe en la matriz
    if estado_col:
        ws[f"{estado_col}{fila_tienda}"] = estado.upper()

    # 2. Actualizar columna de Activo Fijo / Placa solo si aplica y se ingresó
    if activo_col and activo_fijo and activo_fijo not in ["N/A", "SIN ACTIVO"]:
        ws[f"{activo_col}{fila_tienda}"] = activo_fijo.strip().upper()

    # 3. Registrar novedad o solución TI en OBSERVACIONES (Columna AT)
    fecha_corta = datetime.now().strftime("%d/%m/%Y")
    etiqueta_ubic = caja_extra if caja_extra else ubicacion
    detalle_placa = f" [Placa: {activo_fijo.upper()}]" if (activo_fijo and activo_fijo not in ["N/A", "SIN ACTIVO"]) else ""
    detalle_obs = f" - {observacion.strip()}" if observacion.strip() else ""

    if es_cierre_ti:
        nueva_nota = f"[{fecha_corta} TI] {etiqueta_ubic} - {equipo} -> {estado.upper()}{detalle_placa}{detalle_obs}"
    elif estado.upper() in ["REGULAR", "MALO"] or observacion.strip():
        nueva_nota = f"[{fecha_corta}] {etiqueta_ubic} - {equipo} ({estado.upper()}){detalle_placa}{detalle_obs} (Rep: {reportado_por})"
    else:
        nueva_nota = ""

    if nueva_nota:
        obs_actual = ws[f"AT{fila_tienda}"].value
        if obs_actual and str(obs_actual).strip():
            ws[f"AT{fila_tienda}"] = f"{str(obs_actual).strip()} | {nueva_nota}"
        else:
            ws[f"AT{fila_tienda}"] = nueva_nota

    # 4. Actualizar fecha de emisión del formato (Celda G1)
    ws["G1"] = datetime.now()

    wb.save(EXCEL_PATH)
    wb.close()

# ==============================================================================
# INICIALIZACIÓN
# ==============================================================================
init_db()

if "admin_autenticado" not in st.session_state:
    st.session_state.admin_autenticado = False

# ==============================================================================
# BARRA LATERAL: ACCESO EXCLUSIVO TECNOLOGÍA / ADMIN
# ==============================================================================
with st.sidebar:
    st.markdown("### 🔐 Acceso Tecnología")
    st.caption("Área restringida para auditoría, gestión de casos y descarga de Excel.")
    if not st.session_state.admin_autenticado:
        pwd = st.text_input("Contraseña de Administrador", type="password", key="pwd_input")
        if st.button("Ingresar a Panel TI", use_container_width=True):
            if pwd == ADMIN_PASSWORD:
                st.session_state.admin_autenticado = True
                st.rerun()
            else:
                st.error("Contraseña incorrecta.")
    else:
        st.success("Sesión activa: Tecnología")
        vista_seleccionada = st.radio(
            "Navegación",
            ["📋 Formulario Tiendas", "📊 Panel Control Tecnología (Excel)"],
            index=1
        )
        if st.button("Cerrar Sesión TI", use_container_width=True):
            st.session_state.admin_autenticado = False
            st.rerun()

if not st.session_state.admin_autenticado:
    vista_actual = "📋 Formulario Tiendas"
else:
    vista_actual = vista_seleccionada

# ==============================================================================
# VISTA 1: PORTAL DE REPORTE PARA TIENDAS (SIN VISIBILIDAD DE INVENTARIO)
# ==============================================================================
if vista_actual == "📋 Formulario Tiendas":
    col_izq, col_centro, col_der = st.columns([1, 2.2, 1])

    with col_centro:
        st.markdown("## 🖥️ Control de Estado de Activos")
        st.markdown(
            "Reporte novedades o cambios en el estado operativo de los equipos tecnológicos de su tienda. "
            "La información se envía directamente al área de **Tecnología**."
        )

        tiendas = obtener_tiendas_desde_excel()
        if not tiendas:
            st.error(f"No se encontró el archivo maestro `{EXCEL_PATH}` en el servidor.")
            st.stop()

        opciones_tiendas = {
            f"{t['tienda']} ({t['ciudad']}) — CC: {t['c_costos']}": t for t in tiendas
        }

        with st.container(border=True):
            st.subheader("1. Tienda Con Novedad")
            tienda_label = st.selectbox(
                "Seleccione la Tienda / Oficina *",
                options=["-- Seleccione su tienda --"] + list(opciones_tiendas.keys())
            )

            st.subheader("2. Ubicación del equipo")
            ubicacion = st.selectbox(
                "¿Dónde se encuentra el equipo? *",
                options=list(MAPEO_EQUIPOS.keys())
            )

            caja_extra = ""
            if ubicacion == "Caja Adicional (Caja 3 / 4 / 5)":
                caja_extra = st.selectbox(
                    "Especifique el número de caja *",
                    options=["Caja 3", "Caja 4", "Caja 5"]
                )

            equipos_disponibles = list(MAPEO_EQUIPOS[ubicacion].keys())
            equipo = st.selectbox(
                "Seleccione el equipo o periférico afectado *",
                options=equipos_disponibles
            )

            st.subheader("3. Nuevo Estado Operativo")
            estado = st.radio(
                "Seleccione el estado actual del equipo *",
                options=["Bueno", "Regular", "Malo"],
                horizontal=True
            )

            # VALIDACIÓN CONDICIONAL:
            # Solo pide Activo Fijo / Placa si es Regular o Malo Y el equipo maneja activo fijo
            requiere_activo = (estado in ["Regular", "Malo"]) and (equipo not in PERIFERICOS_SIN_ACTIVO)
            activo_fijo = ""

            if requiere_activo:
                st.warning("⚠️ Al reportar este equipo en estado **Regular** o **Malo**, es obligatorio ingresar el número de **Activo Fijo / Placa**.")
                activo_fijo = st.text_input(
                    "Número de Activo Fijo / Placa *",
                    placeholder="Ej: R05177, A-1702, 7730..."
                )
            elif estado in ["Regular", "Malo"] and equipo in PERIFERICOS_SIN_ACTIVO:
                st.info(f"ℹ️ El periférico **{equipo}** no requiere número de Activo Fijo / Placa. Solo describa la novedad.")

            observacion = st.text_area(
                "Detalle de la novedad o falla" + (" *" if estado in ["Regular", "Malo"] else " (Opcional)"),
                placeholder="Ej: El clic izquierdo del mouse dejó de funcionar / El teclado tiene teclas pegadas...",
                height=90
            )

            reportado_por = st.text_input(
                "Nombre de quien reporta *",
                placeholder="Nombre y apellido del colaborador"
            )

            st.markdown("---")
            enviar = st.button("📤 Enviar Actualización a Tecnología", type="primary", use_container_width=True)

            if enviar:
                errores = []
                if tienda_label == "-- Seleccione su tienda --":
                    errores.append("Debe seleccionar una Tienda / Oficina.")
                if requiere_activo and not activo_fijo.strip():
                    errores.append("Debe ingresar el número de **Activo Fijo / Placa** del equipo.")
                if estado in ["Regular", "Malo"] and not observacion.strip():
                    errores.append("Debe describir brevemente la falla presentada.")
                if not reportado_por.strip():
                    errores.append("Debe indicar el nombre de quien realiza el reporte.")

                if errores:
                    for err in errores:
                        st.error(err)
                else:
                    tienda_info = opciones_tiendas[tienda_label]
                    fecha_now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    placa_guardada = (
                        activo_fijo.strip().upper()
                        if requiere_activo
                        else ("SIN ACTIVO" if equipo in PERIFERICOS_SIN_ACTIVO else "N/A")
                    )

                    try:
                        actualizar_excel_maestro(
                            fila_tienda=tienda_info["fila"],
                            ubicacion=ubicacion,
                            caja_extra=caja_extra,
                            equipo=equipo,
                            estado=estado,
                            activo_fijo=placa_guardada,
                            observacion=observacion,
                            reportado_por=reportado_por,
                            es_cierre_ti=False
                        )

                        guardar_log_db({
                            "fecha_hora": fecha_now,
                            "fila_excel": tienda_info["fila"],
                            "c_costos": tienda_info["c_costos"],
                            "tienda": tienda_info["tienda"],
                            "ciudad": tienda_info["ciudad"],
                            "ubicacion": ubicacion,
                            "caja_extra": caja_extra,
                            "equipo": equipo,
                            "estado": estado.upper(),
                            "activo_fijo": placa_guardada,
                            "observacion": observacion.strip(),
                            "reportado_por": reportado_por.strip()
                        })

                        st.success(
                            f"✅ **Reporte registrado correctamente.**\n\n"
                            f"**Tienda:** {tienda_info['tienda']} | **Equipo:** {equipo} | **Nuevo Estado:** {estado.upper()}\n\n"
                            f"El área de Tecnología ha recibido la actualización."
                        )
                    except PermissionError:
                        st.error("❌ El archivo Excel maestro está abierto en este momento por otro usuario. Ciérrelo e intente nuevamente.")
                    except Exception as e:
                        st.error(f"❌ Ocurrió un error al guardar el reporte: {e}")

# ==============================================================================
# VISTA 2: PANEL EXCLUSIVO TECNOLOGÍA / ADMINISTRADOR
# ==============================================================================
elif vista_actual == "📊 Panel Control Tecnología (Excel)":
    st.title("📊 Panel de Control y Atención TI — Tecnología")
    st.caption("Gestión de requerimientos reportados por tiendas y descarga directa del formato maestro actualizado.")

    conn = sqlite3.connect(DB_PATH)
    df_logs = pd.read_sql_query("SELECT * FROM reportes ORDER BY id DESC", conn)
    conn.close()

    m1, m2, m3, m4 = st.columns(4)
    total_rep = len(df_logs)
    pendientes = len(df_logs[df_logs["estado_ti"] != "🟢 Solucionado / Operativo"]) if not df_logs.empty else 0
    solucionados = len(df_logs[df_logs["estado_ti"] == "🟢 Solucionado / Operativo"]) if not df_logs.empty else 0
    malos_activos = len(df_logs[(df_logs["estado"] == "MALO") & (df_logs["estado_ti"] != "🟢 Solucionado / Operativo")]) if not df_logs.empty else 0

    m1.metric("Total Reportes", total_rep)
    m2.metric("⏳ Casos Abiertos / En Gestión", pendientes)
    m3.metric("🔴 Equipos MALOS sin resolver", malos_activos)
    m4.metric("✅ Casos Solucionados", solucionados)

    st.markdown("---")

    # ==========================================================================
    # MÓDULO DE ATENCIÓN Y SOLUCIÓN DE REQUERIMIENTOS (ADMIN TI)
    # ==========================================================================
    if not df_logs.empty:
        with st.expander("🛠️ Atender / Actualizar Estado de un Requerimiento", expanded=True):
            opciones_casos = {
                f"ID #{row['id']} | {row['tienda']} — {row['equipo']} ({row['ubicacion']}) | Estado actual: {row['estado_ti']}": row
                for _, row in df_logs.iterrows()
            }

            caso_sel_label = st.selectbox(
                "Seleccione el requerimiento a gestionar:",
                options=list(opciones_casos.keys())
            )
            caso = opciones_casos[caso_sel_label]

            st.info(
                f"**Detalle del reporte:** Tienda **{caso['tienda']}** | Equipo: **{caso['equipo']}** | "
                f"Placa reportada: **{caso['activo_fijo']}** | Reportado por: **{caso['reportado_por']}**\n\n"
                f"**Falla reportada:** {caso['observacion'] if caso['observacion'] else 'Sin observación'}"
            )

            col_g1, col_g2 = st.columns(2)
            with col_g1:
                idx_actual = ESTADOS_GESTION_TI.index(caso["estado_ti"]) if caso["estado_ti"] in ESTADOS_GESTION_TI else 0
                nuevo_estado_ti = st.selectbox(
                    "Nuevo Estado de Gestión TI *",
                    options=ESTADOS_GESTION_TI,
                    index=idx_actual
                )

            nueva_placa = caso["activo_fijo"]
            actualizar_excel_a_bueno = False

            with col_g2:
                if nuevo_estado_ti == "🟢 Solucionado / Operativo":
                    actualizar_excel_a_bueno = st.checkbox(
                        "Actualizar estado del equipo en el Excel a BUENO",
                        value=True
                    )
                    if caso["equipo"] not in PERIFERICOS_SIN_ACTIVO:
                        nueva_placa = st.text_input(
                            "Activo Fijo / Placa final (si se cambió el equipo, ingrese la nueva placa):",
                            value=caso["activo_fijo"] if caso["activo_fijo"] not in ["N/A", "SIN ACTIVO"] else ""
                        )

            solucion_ti = st.text_input(
                "Diagnóstico o Solución aplicada por Tecnología *",
                value=caso["solucion_ti"] if caso["solucion_ti"] else "",
                placeholder="Ej: Se envió mouse nuevo / Se reinstaló controlador POS..."
            )

            if st.button("💾 Guardar Gestión TI y Sincronizar Excel", type="primary"):
                if not solucion_ti.strip():
                    st.error("Por favor ingrese el diagnóstico o acción realizada por Tecnología.")
                else:
                    try:
                        fecha_cierre = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        nuevo_estado_equipo = "BUENO" if (nuevo_estado_ti == "🟢 Solucionado / Operativo" and actualizar_excel_a_bueno) else caso["estado"]
                        placa_final = nueva_placa.strip().upper() if (nueva_placa and nueva_placa.strip()) else caso["activo_fijo"]

                        fila_excel = caso["fila_excel"]
                        if pd.isna(fila_excel) or not fila_excel:
                            tiendas_map = {t["tienda"]: t["fila"] for t in obtener_tiendas_desde_excel()}
                            fila_excel = tiendas_map.get(caso["tienda"])

                        if fila_excel:
                            nota_excel = f"{nuevo_estado_ti.split(' ', 1)[1]}: {solucion_ti.strip()}"
                            actualizar_excel_maestro(
                                fila_tienda=int(fila_excel),
                                ubicacion=caso["ubicacion"],
                                caja_extra=caso["caja_extra"] if caso["caja_extra"] else "",
                                equipo=caso["equipo"],
                                estado=nuevo_estado_equipo,
                                activo_fijo=placa_final,
                                observacion=nota_excel,
                                reportado_por="Tecnología",
                                es_cierre_ti=True
                            )

                        conn = sqlite3.connect(DB_PATH)
                        c = conn.cursor()
                        c.execute("""
                            UPDATE reportes
                            SET estado_ti = ?, solucion_ti = ?, fecha_cierre_ti = ?,
                                estado = ?, activo_fijo = ?
                            WHERE id = ?
                        """, (
                            nuevo_estado_ti, solucion_ti.strip(), fecha_cierre,
                            nuevo_estado_equipo, placa_final, int(caso["id"])
                        ))
                        conn.commit()
                        conn.close()

                        st.success("✅ Gestión guardada y sincronizada con el formato Excel.")
                        st.rerun()
                    except PermissionError:
                        st.error("❌ Cierre el archivo Excel maestro antes de guardar los cambios.")
                    except Exception as e:
                        st.error(f"❌ Error al actualizar el caso: {e}")

    st.markdown("---")

    # Botón de descarga directa del Excel oficial actualizado
    col_dl1, col_dl2 = st.columns([2, 1])
    with col_dl1:
        st.subheader("Historial de Requerimientos y Seguimiento TI")
    with col_dl2:
        if os.path.exists(EXCEL_PATH):
            with open(EXCEL_PATH, "rb") as f:
                excel_bytes = f.read()
            st.download_button(
                label="📥 Descargar Excel Oficial Actualizado (.xlsx)",
                data=excel_bytes,
                file_name=f"ESTADO_TIENDAS_2026_ACTUALIZADO_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )

    if df_logs.empty:
        st.info("Aún no se han recibido reportes desde el portal de tiendas.")
    else:
        f1, f2, f3 = st.columns(3)
        with f1:
            filtro_gestion = st.multiselect("Filtrar por Gestión TI", ESTADOS_GESTION_TI, default=ESTADOS_GESTION_TI)
        with f2:
            filtro_tienda = st.multiselect("Filtrar por Tienda", sorted(df_logs["tienda"].unique()))
        with f3:
            buscar_placa = st.text_input("Buscar por Activo Fijo / Placa", placeholder="Ej: R05177")

        df_filtrado = df_logs[df_logs["estado_ti"].isin(filtro_gestion)]
        if filtro_tienda:
            df_filtrado = df_filtrado[df_filtrado["tienda"].isin(filtro_tienda)]
        if buscar_placa.strip():
            df_filtrado = df_filtrado[df_filtrado["activo_fijo"].str.contains(buscar_placa.strip(), case=False, na=False)]

        st.dataframe(
            df_filtrado[[
                "id", "fecha_hora", "tienda", "ubicacion", "caja_extra",
                "equipo", "estado", "activo_fijo", "observacion",
                "estado_ti", "solucion_ti", "fecha_cierre_ti", "reportado_por"
            ]],
            use_container_width=True,
            hide_index=True
        )
