import streamlit as st
import pandas as pd
import openpyxl
from datetime import datetime
import os
import sqlite3
from io import BytesIO

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
ADMIN_PASSWORD = os.getenv("TI_ADMIN_PASS", "Tecnologia2026*")  # Cambia esta clave a tu gusto

# ==============================================================================
# MAPEO EXACTO DE COLUMNAS DEL EXCEL (ESTADO_TIENDAS 2026.xlsx)
# ==============================================================================
# Estructura:
# "Ubicación": {
#     "Nombre Equipo": {"activo_col": "Letra o None", "estado_col": "Letra o None"}
# }
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
# BASE DE DATOS DE AUDITORÍA (SQLITE)
# ==============================================================================
def init_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS reportes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha_hora TEXT,
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
            revisado_ti INTEGER DEFAULT 0
        )
    """)
    conn.commit()
    conn.close()

def guardar_log_db(data):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("""
        INSERT INTO reportes (
            fecha_hora, c_costos, tienda, ciudad, ubicacion, caja_extra,
            equipo, estado, activo_fijo, observacion, reportado_por
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        data["fecha_hora"], data["c_costos"], data["tienda"], data["ciudad"],
        data["ubicacion"], data["caja_extra"], data["equipo"], data["estado"],
        data["activo_fijo"], data["observacion"], data["reportado_por"]
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

def actualizar_excel_maestro(fila_tienda, ubicacion, caja_extra, equipo, estado, activo_fijo, observacion, reportado_por):
    """Actualiza la celda exacta de Estado, Activo Fijo y Observaciones en el .xlsx"""
    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb[SHEET_NAME]

    config_col = MAPEO_EQUIPOS[ubicacion][equipo]
    estado_col = config_col["estado_col"]
    activo_col = config_col["activo_col"]

    # 1. Actualizar columna de Estado si existe en la matriz
    if estado_col:
        ws[f"{estado_col}{fila_tienda}"] = estado.upper()

    # 2. Actualizar columna de Activo Fijo / Placa si se ingresó y la columna existe
    if activo_col and activo_fijo:
        ws[f"{activo_col}{fila_tienda}"] = activo_fijo.strip().upper()

    # 3. Registrar novedad en OBSERVACIONES (Columna AT) cuando es REGULAR o MALO, o si hay nota
    fecha_corta = datetime.now().strftime("%d/%m/%Y")
    etiqueta_ubic = caja_extra if caja_extra else ubicacion
    detalle_placa = f" [Placa: {activo_fijo.upper()}]" if activo_fijo else ""
    detalle_obs = f" - {observacion.strip()}" if observacion.strip() else ""

    if estado.upper() in ["REGULAR", "MALO"] or observacion.strip():
        nueva_nota = f"[{fecha_corta}] {etiqueta_ubic} - {equipo} ({estado.upper()}){detalle_placa}{detalle_obs} (Rep: {reportado_por})"
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
    st.caption("Área restringida para auditoría de base de datos y descarga de Excel.")
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

# Si no es admin, siempre ve únicamente el formulario de reporte
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
            st.subheader("2. Equipo Afectado")
                ubicacion = st.selectbox(
                "¿Dónde se encuentra el equipo? *",
                options=list(MAPEO_EQUIPOS.keys())
            )
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

            # VALIDACIÓN CONDICIONAL: Solo pide Activo Fijo / Placa si es Regular o Malo
            activo_fijo = ""
            if estado in ["Regular", "Malo"]:
                st.warning("⚠️ Al reportar un equipo en estado **Regular** o **Malo**, es obligatorio ingresar el número de **Activo Fijo / Placa**.")
                activo_fijo = st.text_input(
                    "Número de Activo Fijo / Placa *",
                    placeholder="Ej: R05177, A-1702, 7730..."
                )

            observacion = st.text_area(
                "Detalle de la novedad o falla" + (" *" if estado in ["Regular", "Malo"] else " (Opcional)"),
                placeholder="Ej: El clic izquierdo del mouse dejó de funcionar / El equipo de sonido presenta interferencia...",
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
                if estado in ["Regular", "Malo"] and not activo_fijo.strip():
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

                    try:
                        # 1. Actualizar Excel Oficial
                        actualizar_excel_maestro(
                            fila_tienda=tienda_info["fila"],
                            ubicacion=ubicacion,
                            caja_extra=caja_extra,
                            equipo=equipo,
                            estado=estado,
                            activo_fijo=activo_fijo,
                            observacion=observacion,
                            reportado_por=reportado_por
                        )

                        # 2. Guardar en historial de auditoría
                        guardar_log_db({
                            "fecha_hora": fecha_now,
                            "c_costos": tienda_info["c_costos"],
                            "tienda": tienda_info["tienda"],
                            "ciudad": tienda_info["ciudad"],
                            "ubicacion": ubicacion,
                            "caja_extra": caja_extra,
                            "equipo": equipo,
                            "estado": estado.upper(),
                            "activo_fijo": activo_fijo.strip().upper() if activo_fijo else "N/A",
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
    st.title("📊 Panel de Control y Auditoría — Tecnología")
    st.caption("Visualización restringida de novedades reportadas y descarga directa del formato maestro actualizado.")

    conn = sqlite3.connect(DB_PATH)
    df_logs = pd.read_sql_query("SELECT * FROM reportes ORDER BY id DESC", conn)
    conn.close()

    # Métricas superiores
    m1, m2, m3, m4 = st.columns(4)
    total_rep = len(df_logs)
    malos = len(df_logs[df_logs["estado"] == "MALO"]) if not df_logs.empty else 0
    regulares = len(df_logs[df_logs["estado"] == "REGULAR"]) if not df_logs.empty else 0
    buenos = len(df_logs[df_logs["estado"] == "BUENO"]) if not df_logs.empty else 0

    m1.metric("Total Reportes Recibidos", total_rep)
    m2.metric("🔴 Equipos MALOS", malos)
    m3.metric("🟡 Equipos REGULARES", regulares)
    m4.metric("🟢 Reportes BUENOS", buenos)

    st.markdown("---")

    # Botón de descarga directa del Excel oficial actualizado
    col_dl1, col_dl2 = st.columns([2, 1])
    with col_dl1:
        st.subheader("Historial de Actualizaciones Recibidas")
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
        # Filtros para Tecnología
        f1, f2, f3 = st.columns(3)
        with f1:
            filtro_estado = st.multiselect("Filtrar por Estado", ["MALO", "REGULAR", "BUENO"], default=["MALO", "REGULAR", "BUENO"])
        with f2:
            filtro_tienda = st.multiselect("Filtrar por Tienda", sorted(df_logs["tienda"].unique()))
        with f3:
            buscar_placa = st.text_input("Buscar por Activo Fijo / Placa", placeholder="Ej: R05177")

        df_filtrado = df_logs[df_logs["estado"].isin(filtro_estado)]
        if filtro_tienda:
            df_filtrado = df_filtrado[df_filtrado["tienda"].isin(filtro_tienda)]
        if buscar_placa.strip():
            df_filtrado = df_filtrado[df_filtrado["activo_fijo"].str.contains(buscar_placa.strip(), case=False, na=False)]

        st.dataframe(
            df_filtrado[[
                "id", "fecha_hora", "c_costos", "tienda", "ciudad",
                "ubicacion", "caja_extra", "equipo", "estado", "activo_fijo",
                "observacion", "reportado_por"
            ]],
            use_container_width=True,
            hide_index=True
        )
