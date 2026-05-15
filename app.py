"""Interfaz Streamlit local para procesar RMT y generar formularios 104/103."""
from __future__ import annotations

from datetime import datetime
from pathlib import Path
from time import perf_counter

import pandas as pd
import streamlit as st

from src import db
from src.consolidator import consolidar, exportar_excel
from src.forms import calcular_formularios, exportar_formulario
from src.parser import parse

ROOT = Path(__file__).resolve().parent
DATA_IN = ROOT / "data" / "rmt_in"
DATA_OUT = ROOT / "data" / "rmt_out"
DEFAULT_CLIENT_ROOT = ROOT / "data" / "clientes"
DEFAULT_TEMPLATE = Path(r"C:\Users\llper\OneDrive\Memin\Descargas\FORMULARIO IVA.xlsx")

st.set_page_config(page_title="Resumen Mensual Optimizado", layout="wide")

st.title("Resumen Mensual Optimizado")
st.caption("Procesamiento local de RMT, consolidación por cliente y generación de Formularios 104 / 103.")

# ─── Sidebar: configuración ──────────────────────────────────────────────────
with st.sidebar:
    st.header("Configuración")

    data_root = Path(
        st.text_input(
            "Carpeta local para bases por cliente",
            value=str(DEFAULT_CLIENT_ROOT),
            help="Cada RUC tendrá su propia subcarpeta con un SQLite.",
        )
    )
    crear_carpetas = st.checkbox("Autorizar creación de carpetas locales", value=True)

    template_path = Path(
        st.text_input(
            "Plantilla Excel (Formularios 104 / 103)",
            value=str(DEFAULT_TEMPLATE),
        )
    )

    st.divider()
    st.subheader("Declaración")

    tipo_ruc = st.radio(
        "Tipo de contribuyente",
        ["Persona jurídica", "Persona natural"],
    )

    obligado_103 = st.checkbox("Obligado a declarar Formulario 103", value=False)

    tiene_empleados = False
    base_iess = 0.0
    retencion_empleados = 0.0

    if obligado_103:
        tiene_empleados = st.checkbox("Tiene empleados en relación de dependencia")
        if tiene_empleados:
            base_iess = st.number_input(
                "Base aportable IESS",
                min_value=0.0,
                step=10.0,
                format="%.2f",
            )
            st.caption(f"Aporte personal 9.45 %: {base_iess * 0.0945:,.2f}")
            retencion_empleados = st.number_input(
                "Total retenido a empleados",
                min_value=0.0,
                step=1.0,
                format="%.2f",
            )

# ─── Tabs principales ────────────────────────────────────────────────────────
tab_procesar, tab_feed, tab_consolidado = st.tabs(["Procesar", "Feed", "Consolidado"])

# ─── Tab 1: Procesar ─────────────────────────────────────────────────────────
with tab_procesar:
    st.subheader("Cargar y procesar RMT")

    uploaded = st.file_uploader(
        "Seleccione uno o varios archivos RMT exportados del portal SRI",
        type=["xlsx", "xlsm"],
        accept_multiple_files=True,
    )

    col_proc, col_info = st.columns([1, 3])
    with col_proc:
        procesar = st.button("Procesar", type="primary", disabled=not uploaded)
    with col_info:
        st.caption(f"Entrada temporal: `{DATA_IN}`")

    if procesar and uploaded:
        if not crear_carpetas:
            st.error("Active la autorización para crear carpetas locales antes de procesar.")
        elif not template_path.exists():
            st.error(f"No se encontró la plantilla: {template_path}")
        else:
            DATA_IN.mkdir(parents=True, exist_ok=True)
            DATA_OUT.mkdir(parents=True, exist_ok=True)
            data_root.mkdir(parents=True, exist_ok=True)

            for file in uploaded:
                with st.container(border=True):
                    start = perf_counter()
                    local_path = DATA_IN / file.name
                    local_path.write_bytes(file.getbuffer())

                    try:
                        rmt = parse(local_path)
                        calculado = calcular_formularios(
                            rmt,
                            obligado_103=obligado_103,
                            base_iess_empleados=base_iess if tiene_empleados else 0.0,
                            retencion_empleados=retencion_empleados if tiene_empleados else 0.0,
                        )
                        db_path = db.client_db_path(rmt.ruc, rmt.cliente_nombre, data_root)
                        archivo_id = db.save_result(db_path, rmt, calculado)

                        output = (
                            DATA_OUT
                            / f"formularios_{rmt.ruc}_{rmt.periodo_inicio}_{archivo_id}.xlsx"
                        )
                        exportar_formulario(template_path, output, calculado)

                        elapsed = perf_counter() - start

                        st.markdown(f"**{file.name}**")
                        c1, c2, c3, c4 = st.columns(4)
                        c1.metric("RUC", rmt.ruc)
                        c2.metric("Período", rmt.periodo_label)
                        c3.metric("Documentos", len(rmt.documentos))
                        c4.metric("Anulados", len(rmt.anulados))

                        if elapsed > 10:
                            st.warning(f"Procesado en {elapsed:.1f} s (superó el umbral de 10 s).")
                        else:
                            st.success(f"Procesado en {elapsed:.2f} s")

                        col_a, col_b = st.columns(2)
                        col_a.caption(f"Base de datos del cliente: `{db_path}`")
                        col_b.caption(f"Formulario exportado: `{output}`")

                        with open(output, "rb") as fh:
                            st.download_button(
                                label="Descargar Excel del formulario",
                                data=fh.read(),
                                file_name=output.name,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key=f"dl_{archivo_id}",
                            )

                        if calculado.advertencias:
                            with st.expander(f"Advertencias ({len(calculado.advertencias)})"):
                                for adv in calculado.advertencias:
                                    st.write(f"- {adv}")
                    except Exception as exc:
                        st.error(f"Error procesando {file.name}: {exc}")
                        st.exception(exc)

# ─── Tab 2: Feed ─────────────────────────────────────────────────────────────
with tab_feed:
    st.subheader("Archivos procesados")

    start = perf_counter()
    feed = db.latest_files(data_root)
    elapsed = perf_counter() - start
    st.caption(f"Consulta del feed: {elapsed:.3f} s")

    if not feed:
        st.info("Aún no hay archivos procesados.")
    else:
        feed_df = pd.DataFrame(feed)

        col_a, col_b = st.columns(2)
        with col_a:
            ruc_options = ["Todos"] + sorted(feed_df["ruc"].dropna().unique().tolist())
            sel_ruc = st.selectbox("Filtrar por RUC", ruc_options)
        with col_b:
            periodos = sorted(feed_df["periodo_inicio"].dropna().unique().tolist(), reverse=True)
            sel_periodo = st.selectbox("Filtrar por período", ["Todos"] + periodos)

        view = feed_df.copy()
        if sel_ruc != "Todos":
            view = view[view["ruc"] == sel_ruc]
        if sel_periodo != "Todos":
            view = view[view["periodo_inicio"] == sel_periodo]

        st.caption(f"{len(view)} archivo(s)")

        for _, row in view.iterrows():
            with st.container(border=True):
                c1, c2, c3, c4 = st.columns([3, 1, 1, 1])
                c1.markdown(f"**{row['cliente_nombre']}**")
                c1.caption(row["nombre"])
                c2.metric("RUC", row["ruc"])
                c3.metric("Período", row["periodo_inicio"])
                c4.metric("Estado", row["estado"])

                with st.expander("Ver casilleros calculados y documentos"):
                    db_path = Path(row["db_path"])
                    archivo_id = int(row["id"])

                    st.markdown("**Formulario 104**")
                    cas_104 = db.casilleros(db_path, archivo_id, "104")
                    if cas_104:
                        st.dataframe(
                            pd.DataFrame(cas_104),
                            use_container_width=True,
                            hide_index=True,
                        )
                    else:
                        st.caption("Sin casilleros calculados.")

                    if obligado_103:
                        st.markdown("**Formulario 103**")
                        cas_103 = db.casilleros(db_path, archivo_id, "103")
                        if cas_103:
                            st.dataframe(
                                pd.DataFrame(cas_103),
                                use_container_width=True,
                                hide_index=True,
                            )
                        else:
                            st.caption("Sin casilleros calculados.")

                    st.markdown("**Resumen de documentos**")
                    resumen = db.documentos_resumen(db_path, archivo_id)
                    if resumen:
                        st.dataframe(
                            pd.DataFrame(resumen),
                            use_container_width=True,
                            hide_index=True,
                        )

# ─── Tab 3: Consolidado ──────────────────────────────────────────────────────
with tab_consolidado:
    st.subheader("Consolidado de documentos")

    col_a, col_b, col_c = st.columns([1, 1, 1])
    with col_a:
        ruc_filtro = st.text_input("RUC (opcional)", "")
    with col_b:
        periodo_filtro = st.text_input("Período YYYY-MM (opcional)", "")
    with col_c:
        st.write("")
        st.write("")
        consultar = st.button("Consultar")

    if consultar:
        start = perf_counter()
        df = consolidar(data_root, ruc=ruc_filtro or None, periodo=periodo_filtro or None)
        elapsed = perf_counter() - start
        st.caption(f"Consulta: {elapsed:.3f} s")

        if elapsed > 10:
            st.warning("La consulta superó 10 segundos.")

        if df.empty:
            st.info("Sin resultados con esos filtros.")
        else:
            st.dataframe(df, use_container_width=True, hide_index=True)
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            destino = DATA_OUT / f"consolidado_{ts}.xlsx"
            exportar_excel(df, destino)
            st.success(f"Exportado: {destino}")
            with open(destino, "rb") as fh:
                st.download_button(
                    "Descargar Excel consolidado",
                    fh.read(),
                    file_name=destino.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
