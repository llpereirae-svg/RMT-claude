"""Procesador de Resumen Mensual de Transacciones — UI corporativa local."""
from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from time import perf_counter

import pandas as pd
import streamlit as st

from src import db, auth
from src.ats_models import ATSProcesado
from src.ats_parser import parse_ats
from src.consolidator import consolidar, exportar_excel
from src.forms import calcular_formularios, exportar_formulario
from src.parser import parse

ROOT = Path(__file__).resolve().parent
DATA_IN = ROOT / "data" / "rmt_in"
DATA_OUT = ROOT / "data" / "rmt_out"
DEFAULT_CLIENT_ROOT = ROOT / "data" / "clientes"
# Template incluido en el repo. En local puedes apuntar a otra ruta desde el sidebar.
REPO_TEMPLATE = ROOT / "templates" / "FORMULARIO_IVA.xlsx"
LOCAL_TEMPLATE = Path(r"C:\Users\llper\OneDrive\Memin\Descargas\FORMULARIO IVA.xlsx")
DEFAULT_TEMPLATE = LOCAL_TEMPLATE if LOCAL_TEMPLATE.exists() else REPO_TEMPLATE
MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


# ─── Hidratar usuarios desde st.secrets (Streamlit Cloud) ────────────────────
def _hydrate_users_from_secrets() -> dict:
    """En Streamlit Cloud, carga st.secrets['users'] al store en memoria de auth.
    Devuelve un diagnostico {found, source, count, error}.
    En local con archivo .streamlit/users.json existente, no hace nada (file wins)."""
    diag = {"found": False, "source": None, "count": 0, "error": None}

    # Si ya hay archivo local con usuarios, ese gana (modo dev).
    if auth.USERS_FILE.exists():
        try:
            raw = json.loads(auth.USERS_FILE.read_text(encoding="utf-8"))
            diag.update(found=True, source="file", count=len(raw))
        except (json.JSONDecodeError, OSError) as exc:
            diag.update(error=f"file_read_error: {exc}")
        return diag

    # Caso cloud: leer de Secrets y cargar al store en memoria.
    try:
        secret_users = st.secrets.get("users")
    except Exception as exc:
        diag.update(error=f"secrets_access_error: {type(exc).__name__}: {exc}")
        return diag

    if not secret_users:
        diag.update(error="no 'users' key in st.secrets")
        return diag

    try:
        content = {}
        for username, info in dict(secret_users).items():
            content[str(username)] = dict(info)
        auth.use_memory_store(content)
        diag.update(found=True, source="secrets", count=len(content))
    except Exception as exc:
        diag.update(error=f"hydration_error: {type(exc).__name__}: {exc}")
    return diag


_users_diag = _hydrate_users_from_secrets()


# ─── Panel de diagnostico (URL ?debug=secrets) ──────────────────────────────
if st.query_params.get("debug") == "secrets":
    st.title("Diagnóstico de autenticación")
    st.json(_users_diag)
    st.write("Usuarios cargados (solo nombres, sin hashes):")
    try:
        st.write(list(auth._load_users_raw().keys()))
    except Exception as exc:
        st.error(f"No se pueden listar usuarios: {exc}")
    st.write("USERS_FILE existe:", auth.USERS_FILE.exists())
    st.write("USERS_FILE path:", str(auth.USERS_FILE))
    st.caption("Quita el `?debug=secrets` del URL para volver al login normal.")
    st.stop()

st.set_page_config(
    page_title="RMT Suite",
    layout="wide",
    initial_sidebar_state="collapsed",  # sidebar como íconos por defecto
)

# ─── Estilos mínimos (sin envoltorios HTML) ──────────────────────────────────
st.markdown(
    """
    <style>
    .block-container { padding-top: 0.6rem; padding-bottom: 1.5rem; max-width: 100%; }

    /* Header corporativo SAP S/4HANA */
    .corp-header {
        background: linear-gradient(180deg, #354a5f 0%, #2c3e50 100%);
        color: #ffffff;
        padding: 0.7rem 1.1rem;
        border-radius: 4px;
        margin-bottom: 0.7rem;
        display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 1rem;
    }
    .corp-header .title { color: #fff; font-size: 1rem; font-weight: 600; letter-spacing: 0.3px; margin: 0; }
    .corp-header .sub { color: #c5cdd6; font-size: 0.74rem; margin-top: 2px; }
    .corp-kpis { display: flex; gap: 1.5rem; }
    .corp-kpi { text-align: right; }
    .corp-kpi .v { font-size: 1.1rem; font-weight: 600; line-height: 1.1; }
    .corp-kpi .l { font-size: 0.66rem; color: #c5cdd6; text-transform: uppercase; letter-spacing: 0.5px; }

    /* Cintilla de declaración */
    .ribbon {
        background: #ffffff; border: 1px solid #d9d9d9; border-radius: 3px;
        padding: 0.6rem 0.9rem; margin-bottom: 0.7rem;
    }

    /* Botón primario azul SAP */
    .stButton > button[kind="primary"] {
        background: #0a6ed1; border-color: #0a6ed1; color: #fff; font-weight: 600;
    }
    .stButton > button[kind="primary"]:hover { background: #085caf; border-color: #085caf; }
    .stButton > button { border-radius: 3px; }

    /* Tabla y métricas */
    div[data-testid="stDataFrame"] { border: 1px solid #d9d9d9; border-radius: 3px; }
    [data-testid="stMetric"] {
        background: #fafafa; border: 1px solid #ebebeb; border-radius: 3px; padding: 0.45rem 0.7rem;
    }
    [data-testid="stMetricLabel"] {
        font-size: 0.7rem; color: #6a6d70; text-transform: uppercase; letter-spacing: 0.4px;
    }
    [data-testid="stMetricValue"] { font-size: 1.05rem; font-weight: 600; color: #32363a; }

    /* Sección título dentro de un container */
    .panel-title {
        font-size: 0.74rem; font-weight: 600; text-transform: uppercase; letter-spacing: 0.7px;
        color: #6a6d70; margin: 0 0 0.5rem 0;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ─── Estado de sesión ────────────────────────────────────────────────────────
if "selected_archivo" not in st.session_state:
    st.session_state.selected_archivo = None
if "user" not in st.session_state:
    st.session_state.user = None  # auth.Usuario | None


def _flash(level: str, msg: str) -> None:
    {"success": st.success, "warning": st.warning, "error": st.error, "info": st.info}[level](msg)


# ─── Login y cambio forzado de contraseña ───────────────────────────────────
def _render_login() -> None:
    """Pantalla de inicio de sesión. Si autentica, escribe en session_state.user."""
    _c1, c2, _c3 = st.columns([1, 1.2, 1])
    with c2:
        st.markdown("""
        <div style="margin-top:3rem; text-align:center;">
            <div style="font-size:1.4rem; font-weight:600; color:#32363a;">RMT Suite</div>
            <div style="font-size:0.85rem; color:#6a6d70; margin-top:0.2rem;">
                Iniciar sesión para acceder al procesador
            </div>
        </div>
        """, unsafe_allow_html=True)
        with st.container(border=True):
            with st.form("login_form", clear_on_submit=False):
                username = st.text_input("Usuario (RUC o cédula)", autocomplete="username")
                password = st.text_input("Contraseña", type="password", autocomplete="current-password")
                submitted = st.form_submit_button("Ingresar", type="primary", use_container_width=True)
            if submitted:
                user = auth.authenticate(username, password)
                if user is None:
                    st.error("Usuario o contraseña incorrectos.")
                else:
                    st.session_state.user = user
                    st.session_state.selected_archivo = None
                    st.rerun()
        st.caption("¿No tienes cuenta? Solicita una a tu administrador.")


def _render_password_change(user: "auth.Usuario") -> None:
    """Pantalla que fuerza el cambio de contraseña en el primer login."""
    _c1, c2, _c3 = st.columns([1, 1.2, 1])
    with c2:
        st.markdown(f"""
        <div style="margin-top:2rem; text-align:center;">
            <div style="font-size:1.2rem; font-weight:600; color:#32363a;">
                Cambia tu contraseña provisional
            </div>
            <div style="font-size:0.85rem; color:#6a6d70; margin-top:0.3rem;">
                Hola <b>{user.display_name}</b>. Antes de continuar define una contraseña propia.
            </div>
        </div>
        """, unsafe_allow_html=True)
        with st.container(border=True):
            with st.form("pwd_change_form", clear_on_submit=False):
                new1 = st.text_input("Contraseña nueva", type="password",
                                     autocomplete="new-password",
                                     help="Mínimo 8 caracteres. Evita contraseñas predecibles.")
                new2 = st.text_input("Repetir contraseña nueva", type="password",
                                     autocomplete="new-password")
                submitted = st.form_submit_button("Guardar contraseña", type="primary",
                                                  use_container_width=True)
            if submitted:
                if len(new1) < 8:
                    st.error("La contraseña debe tener al menos 8 caracteres.")
                elif new1 != new2:
                    st.error("Las dos contraseñas no coinciden.")
                else:
                    auth.change_password(user.username, new1)
                    st.session_state.user = auth.get_user(user.username)
                    st.success("Contraseña actualizada. Redirigiendo…")
                    st.rerun()


# Gate de autenticación: si no hay sesión, mostrar login y detener
if st.session_state.user is None:
    _render_login()
    st.stop()

# Si el usuario tiene contraseña provisional, forzar cambio
if st.session_state.user.must_change_password:
    _render_password_change(st.session_state.user)
    st.stop()


# ─── Sidebar: configuración técnica + usuario actual ─────────────────────────
USER = st.session_state.user
USER_DATA_ROOT = auth.user_data_root(USER.username, DEFAULT_CLIENT_ROOT)

with st.sidebar:
    st.markdown(f"**{USER.display_name}**")
    st.caption(f"Usuario `{USER.username}`")
    if st.button("Cerrar sesión", use_container_width=True):
        st.session_state.user = None
        st.session_state.selected_archivo = None
        st.rerun()
    st.markdown("---")
    st.markdown("### Configuración")
    data_root = Path(
        st.text_input("Carpeta de bases por cliente", value=str(USER_DATA_ROOT),
                      help="Cada usuario tiene su propia carpeta. No mezcles aquí datos de otros usuarios.")
    )
    template_path = Path(
        st.text_input("Plantilla 104/103", value=str(DEFAULT_TEMPLATE))
    )
    crear_carpetas = st.checkbox("Autorizar creación de carpetas locales", value=True)


# ─── Datos del feed ─────────────────────────────────────────────────────────
feed_rows: list[dict] = db.latest_files(data_root) if data_root.exists() else []
feed_df = pd.DataFrame(feed_rows) if feed_rows else pd.DataFrame(
    columns=["id", "ruc", "cliente_nombre", "periodo_inicio", "periodo_fin",
             "nombre", "estado", "creado_en", "db_path"]
)

# KPIs globales
n_clientes = feed_df["ruc"].nunique() if not feed_df.empty else 0
n_archivos = len(feed_df)
n_periodos = feed_df["periodo_inicio"].nunique() if not feed_df.empty else 0
ultimo = str(feed_df["creado_en"].max())[:16] if not feed_df.empty else "—"

# ─── Header ──────────────────────────────────────────────────────────────────
st.markdown(
    f"""
    <div class="corp-header">
        <div>
            <div class="title">RMT Suite</div>
            <div class="sub">Procesador local de Resumen Mensual de Transacciones · SRI Ecuador</div>
        </div>
        <div class="corp-kpis">
            <div class="corp-kpi"><div class="v">{n_clientes}</div><div class="l">Clientes</div></div>
            <div class="corp-kpi"><div class="v">{n_archivos}</div><div class="l">Archivos</div></div>
            <div class="corp-kpi"><div class="v">{n_periodos}</div><div class="l">Períodos</div></div>
            <div class="corp-kpi"><div class="v">{ultimo}</div><div class="l">Último proc.</div></div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)


# ─── Cintilla de declaración + carga ─────────────────────────────────────────
with st.container():
    rib_a, rib_b, rib_c, rib_d, rib_e = st.columns([1.3, 1.3, 1.2, 1, 1])
    with rib_a:
        tipo_ruc = st.selectbox(
            "Tipo de contribuyente",
            ["Persona jurídica", "Persona natural"],
            label_visibility="visible",
        )
    with rib_b:
        contri_especial = st.checkbox(
            "Contribuyente / actividad especial",
            help="Si está marcado, las ventas 0% van al casillero 405/415 (con derecho a CT). "
                 "Si no, al 403/413.",
        )
    with rib_c:
        obligado_103 = st.checkbox("Obligado Formulario 103")
    with rib_d:
        cargar_clicked = st.button("Cargar RMT", type="primary", use_container_width=True)
    with rib_e:
        consolidar_clicked = st.button("Consolidar", use_container_width=True)

    # Segunda fila: empleados (solo si obligado 103)
    tiene_empleados = False
    base_iess = 0.0
    retencion_empleados = 0.0
    if obligado_103:
        emp_a, emp_b, emp_c = st.columns([1, 1, 1])
        with emp_a:
            tiene_empleados = st.checkbox("Empleados en relación de dependencia")
        if tiene_empleados:
            with emp_b:
                base_iess = st.number_input("Base aportable IESS", min_value=0.0, step=10.0, format="%.2f")
                st.caption(f"Aporte personal 9.45 %: {base_iess * 0.0945:,.2f}")
            with emp_c:
                retencion_empleados = st.number_input("Retención a empleados", min_value=0.0, step=1.0, format="%.2f")


# ─── Diálogo modal: cargar archivos ──────────────────────────────────────────
@st.dialog("Cargar archivos RMT", width="large")
def _dialog_cargar():
    uploaded = st.file_uploader(
        "Archivos RMT del SRI (.xlsx / .xlsm)",
        type=["xlsx", "xlsm"],
        accept_multiple_files=True,
    )
    ats_files = st.file_uploader(
        "ATS XML opcional (mismo período del RMT) — refina 332, 531 y 407/408",
        type=["xml"],
        accept_multiple_files=True,
    )
    st.caption("Los archivos se guardan en `data/rmt_in/` y se procesan al instante. El ATS se cruza por RUC.")
    procesar = st.button("Procesar", type="primary", disabled=not uploaded, use_container_width=True)

    if procesar and uploaded:
        if not crear_carpetas:
            st.error("Active la autorización para crear carpetas locales (sidebar).")
            return
        if not template_path.exists():
            st.error(f"Plantilla no encontrada: {template_path}")
            return

        DATA_IN.mkdir(parents=True, exist_ok=True)
        DATA_OUT.mkdir(parents=True, exist_ok=True)
        data_root.mkdir(parents=True, exist_ok=True)
        progreso = st.empty()
        resultados: list[tuple[str, str]] = []

        # Indexar los ATS subidos por RUC para cruzar con cada RMT
        ats_by_ruc: dict[str, ATSProcesado] = {}
        for af in ats_files or []:
            ats_path = DATA_IN / af.name
            ats_path.write_bytes(af.getbuffer())
            try:
                ats = parse_ats(ats_path)
                if ats.ruc_informante:
                    ats_by_ruc[ats.ruc_informante] = ats
            except Exception as exc:
                resultados.append(("warning", f"ATS {af.name}: no se pudo parsear ({exc}). Se procesará sin ATS."))

        for i, file in enumerate(uploaded, 1):
            progreso.info(f"Procesando {i}/{len(uploaded)}: {file.name}")
            t0 = perf_counter()
            local_path = DATA_IN / file.name
            local_path.write_bytes(file.getbuffer())
            try:
                rmt = parse(local_path)
                ats = ats_by_ruc.get(rmt.ruc)
                calculado = calcular_formularios(
                    rmt,
                    ats=ats,
                    obligado_103=obligado_103,
                    base_iess_empleados=base_iess if tiene_empleados else 0.0,
                    retencion_empleados=retencion_empleados if tiene_empleados else 0.0,
                    ventas_0_con_credito=contri_especial,
                )
                db_path = db.client_db_path(rmt.ruc, rmt.cliente_nombre, data_root)
                archivo_id = db.save_result(db_path, rmt, calculado)
                output = DATA_OUT / f"formularios_{rmt.ruc}_{rmt.periodo_inicio}_{archivo_id}.xlsx"
                exportar_formulario(template_path, output, calculado)
                dt = perf_counter() - t0
                tag = " · con ATS" if ats else ""
                resultados.append(("success" if dt < 10 else "warning",
                                   f"{rmt.cliente_nombre} · {rmt.periodo_label}{tag} · {dt:.2f} s"))
            except Exception as exc:
                resultados.append(("error", f"{file.name}: {exc}"))

        progreso.empty()
        for lvl, msg in resultados:
            _flash(lvl, msg)
        if resultados and all(lvl != "error" for lvl, _ in resultados):
            if st.button("Cerrar y ver feed", type="primary", use_container_width=True):
                st.rerun()


if cargar_clicked:
    _dialog_cargar()


# ─── Layout principal: 2 columnas (feed + detalle) ───────────────────────────
col_feed, col_detalle = st.columns([3, 2], gap="medium")


# ─── Feed ────────────────────────────────────────────────────────────────────
with col_feed:
    with st.container(border=True):
        st.markdown('<div class="panel-title">Feed de procesamientos</div>', unsafe_allow_html=True)

        # Filtros compactos al tope del feed
        if not feed_df.empty:
            f1, f2, f3 = st.columns([1.2, 1.2, 2])
            with f1:
                rucs = ["Todos"] + sorted(feed_df["ruc"].dropna().unique().tolist())
                sel_ruc = st.selectbox("RUC", rucs, key="f_ruc", label_visibility="collapsed")
            with f2:
                periodos = ["Todos"] + sorted(feed_df["periodo_inicio"].dropna().unique().tolist(), reverse=True)
                sel_periodo = st.selectbox("Período", periodos, key="f_periodo", label_visibility="collapsed")
            with f3:
                sel_cliente = st.text_input("Cliente", "", placeholder="Buscar por razón social…",
                                            key="f_cliente", label_visibility="collapsed")
        else:
            sel_ruc, sel_periodo, sel_cliente = "Todos", "Todos", ""

        # Aplicar filtros
        view = feed_df.copy()
        if not view.empty:
            if sel_ruc != "Todos":
                view = view[view["ruc"] == sel_ruc]
            if sel_periodo != "Todos":
                view = view[view["periodo_inicio"] == sel_periodo]
            if sel_cliente.strip():
                view = view[view["cliente_nombre"].str.contains(sel_cliente.strip(), case=False, na=False)]

        st.caption(f"{len(view)} archivo(s)")

        if view.empty:
            st.info("Cargue un RMT con el botón **Cargar RMT** arriba.")
        else:
            display = view[["cliente_nombre", "ruc", "periodo_inicio", "periodo_fin",
                            "nombre", "estado", "creado_en"]].copy()
            display.columns = ["Cliente", "RUC", "Per. inicio", "Per. fin", "Archivo", "Estado", "Procesado"]
            display["Procesado"] = display["Procesado"].astype(str).str[:16]

            event = st.dataframe(
                display,
                use_container_width=True,
                hide_index=True,
                on_select="rerun",
                selection_mode="single-row",
                height=min(500, 80 + 35 * len(display)),
            )

            if event and event.selection and event.selection["rows"]:
                row_idx = event.selection["rows"][0]
                row = view.iloc[row_idx]
                st.session_state.selected_archivo = (row["db_path"], int(row["id"]))


# ─── Detalle ─────────────────────────────────────────────────────────────────
with col_detalle:
    with st.container(border=True):
        st.markdown('<div class="panel-title">Detalle del archivo</div>', unsafe_allow_html=True)

        if not st.session_state.selected_archivo:
            st.caption("Seleccione una fila del feed para ver casilleros, documentos y descarga.")
        else:
            db_path_str, archivo_id = st.session_state.selected_archivo
            db_path = Path(db_path_str)

            sel_row = view[view["id"] == archivo_id] if not view.empty else feed_df[feed_df["id"] == archivo_id]
            if not sel_row.empty:
                r = sel_row.iloc[0]
                st.markdown(f"**{r['cliente_nombre']}** · RUC {r['ruc']}")
                rango = r['periodo_inicio'] + (f" a {r['periodo_fin']}" if r['periodo_fin'] != r['periodo_inicio'] else "")
                st.caption(f"Período {rango}")

            tab_104, tab_103, tab_docs = st.tabs(["Form 104", "Form 103", "Documentos"])

            with tab_104:
                cas_104 = db.casilleros(db_path, archivo_id, "104")
                if cas_104:
                    df104 = pd.DataFrame(cas_104)
                    df104.columns = ["Casillero", "Valor"]
                    df104["Valor"] = df104["Valor"].map(lambda x: f"{x:,.2f}")
                    st.dataframe(df104, use_container_width=True, hide_index=True, height=340)
                else:
                    st.caption("Sin casilleros calculados.")

            with tab_103:
                cas_103 = db.casilleros(db_path, archivo_id, "103")
                if cas_103:
                    df103 = pd.DataFrame(cas_103)
                    df103.columns = ["Casillero", "Valor"]
                    df103["Valor"] = df103["Valor"].map(lambda x: f"{x:,.2f}")
                    st.dataframe(df103, use_container_width=True, hide_index=True, height=340)
                else:
                    st.caption("No declara 103 o sin retenciones.")

            with tab_docs:
                resumen = db.documentos_resumen(db_path, archivo_id)
                if resumen:
                    st.dataframe(pd.DataFrame(resumen), use_container_width=True, hide_index=True, height=340)
                else:
                    st.caption("Sin documentos.")

            # Reprocesar y Descarga
            if not sel_row.empty:
                r = sel_row.iloc[0]
                col_reproc, col_dl = st.columns(2)

                with col_reproc:
                    if st.button("Reprocesar con config actual", use_container_width=True,
                                 help="Re-aplica los toggles de la cintilla superior (tipo, 103, contri especial) al RMT original."):
                        local_rmt = DATA_IN / r["nombre"]
                        if not local_rmt.exists():
                            st.error(f"No se encontró el archivo original `{r['nombre']}` en `{DATA_IN}`. "
                                     "Vuelve a cargarlo desde Cargar RMT.")
                        else:
                            try:
                                rmt2 = parse(local_rmt)
                                # Buscar ATS del mismo RUC en data/rmt_in (best-effort)
                                ats2 = None
                                for xml in DATA_IN.glob("*.xml"):
                                    try:
                                        cand = parse_ats(xml)
                                        if cand.ruc_informante == rmt2.ruc:
                                            ats2 = cand
                                            break
                                    except Exception:
                                        continue
                                calc2 = calcular_formularios(
                                    rmt2,
                                    ats=ats2,
                                    obligado_103=obligado_103,
                                    base_iess_empleados=base_iess if tiene_empleados else 0.0,
                                    retencion_empleados=retencion_empleados if tiene_empleados else 0.0,
                                    ventas_0_con_credito=contri_especial,
                                )
                                new_id = db.save_result(db_path, rmt2, calc2)
                                output2 = DATA_OUT / f"formularios_{rmt2.ruc}_{rmt2.periodo_inicio}_{new_id}.xlsx"
                                if template_path.exists():
                                    exportar_formulario(template_path, output2, calc2)
                                st.session_state.selected_archivo = (str(db_path), new_id)
                                tag = " · con ATS" if ats2 else ""
                                st.success(f"Reprocesado · 103={'sí' if obligado_103 else 'no'} · "
                                           f"contri especial={'sí' if contri_especial else 'no'}{tag}")
                                st.rerun()
                            except Exception as exc:
                                st.error(f"Error al reprocesar: {exc}")

                with col_dl:
                    output_files = list(DATA_OUT.glob(f"formularios_{r['ruc']}_{r['periodo_inicio']}_{archivo_id}.xlsx"))
                    if output_files:
                        with open(output_files[0], "rb") as fh:
                            st.download_button(
                                "Descargar Excel",
                                fh.read(),
                                file_name=output_files[0].name,
                                mime=MIME_XLSX,
                                use_container_width=True,
                            )


# ─── Consolidado ─────────────────────────────────────────────────────────────
if consolidar_clicked:
    with st.spinner("Generando consolidado…"):
        t0 = perf_counter()
        df = consolidar(
            data_root,
            ruc=sel_ruc if 'sel_ruc' in dir() and sel_ruc != "Todos" else None,
            periodo=sel_periodo if 'sel_periodo' in dir() and sel_periodo != "Todos" else None,
        )
        dt = perf_counter() - t0
    if df.empty:
        st.warning(f"Consolidado vacío con los filtros actuales ({dt:.2f} s).")
    else:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        destino = DATA_OUT / f"consolidado_{ts}.xlsx"
        exportar_excel(df, destino)
        msg = f"Consolidado generado · {len(df):,} fila(s) en {dt:.2f} s"
        (st.warning if dt > 10 else st.success)(msg)
        with open(destino, "rb") as fh:
            st.download_button("Descargar consolidado", fh.read(),
                               file_name=destino.name, mime=MIME_XLSX)
