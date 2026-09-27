# -*- coding: utf-8 -*-
"""
🏗️ Pipeline de Análisis de Ventas — App Streamlit
==================================================
Ejecuta el pipeline completo del notebook PROanalisispower por etapas,
manteniendo la lógica original INTACTA (incluido el guardado/relectura
de Excel entre pasos, del que depende el comportamiento validado).

Etapas:
  1-4  → Carga, limpieza, fechas, mapeo EMP        → resultado_pasos_1_a_4.xlsx
  5    → Unificación NOMCLIE (CLIENTES_UNICOS)  🔍 → resultado_paso_5.xlsx
  6    → Construcción de CU (clave única)          → resultado_paso_6.xlsx
  7    → Unificación DESCRIP/FAM (DESCRIP_UNICAS)🔍→ resultado_paso_7.xlsx
  8-9  → Tasas PAR y BCV (TASAS.xlsx)              → resultado_pasos_8_y_9.xlsx
  7.1  → Estandarización IGTF (2024/2025 → 0)     → (dentro de resultado_paso_7.xlsx)
  10-13→ Consolidación final (TOTAL FACT + MONEDA) → Analisis_Consolidado_Final.xlsx
  CASHEA → Tipo de venta SOCIOS/NETAS/CASHEA       → ..._con_tipo_venta.xlsx
  DEF  → COSTO X CANT, negativos DEV, reorden      → Analisis_Final_DEF.xlsx

🔍 = etapas con panel de auditoría dedicado.
"""
import io
import os
import contextlib
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

import etapas

# ─────────────────────────── Configuración ───────────────────────────
st.set_page_config(
    page_title="Pipeline Análisis de Ventas",
    page_icon="🏗️",
    layout="wide",
)

# Definición del pipeline: (id, etiqueta, función, archivos_requeridos, archivo_salida)
PIPELINE = [
    ("p1_4",   "Pasos 1 a 4 — Carga y limpieza",          "etapa_1_a_4",
     ["__PRINCIPAL__"],                                   "resultado_pasos_1_a_4.xlsx"),
    ("p5",     "Paso 5 — Unificación de clientes 🔍",      "etapa_5",
     ["resultado_pasos_1_a_4.xlsx", "CLIENTES_UNICOS.xlsx"], "resultado_paso_5.xlsx"),
    ("p6",     "Paso 6 — Clave única (CU)",                "etapa_6",
     ["resultado_paso_5.xlsx"],                           "resultado_paso_6.xlsx"),
    ("p7",     "Paso 7 — Unificación DESCRIP/FAM + IGTF 🔍",      "etapa_7",
     ["resultado_paso_6.xlsx", "DESCRIP_UNICAS.xlsx"],    "resultado_paso_7.xlsx"),
    ("p8_9",   "Pasos 8 y 9 — Tasas PAR / BCV",            "etapa_8_y_9",
     ["resultado_paso_7.xlsx", "TASAS.xlsx"],             "resultado_pasos_8_y_9.xlsx"),
    ("p10_13", "Pasos 10 a 13 — Consolidación + MONEDA",      "etapa_10_a_13",
     ["resultado_pasos_8_y_9.xlsx"],                      "Analisis_Consolidado_Final.xlsx"),
    ("cashea", "Tipo de venta (SOCIOS/NETAS/CASHEA)",      "etapa_tipo_venta_cashea",
     ["Analisis_Consolidado_Final.xlsx"],                 "Analisis_Consolidado_Final_con_tipo_venta.xlsx"),
    ("def",    "Transformación final (DEF)",               "etapa_transformacion_final",
     ["Analisis_Consolidado_Final_con_tipo_venta.xlsx"],  "Analisis_Final_DEF.xlsx"),
]

NOMBRE_PRINCIPAL = "archivo_principal.xlsx"  # nombre canónico interno del listado

# ─────────────────────────── Estado de sesión ───────────────────────────
if "workdir" not in st.session_state:
    st.session_state.workdir = tempfile.mkdtemp(prefix="pipeline_ventas_")
if "logs" not in st.session_state:
    st.session_state.logs = {}      # id_etapa -> texto del log
if "estado" not in st.session_state:
    st.session_state.estado = {}    # id_etapa -> "ok" | "error"

WORKDIR = Path(st.session_state.workdir)


# ─────────────────────────── Utilidades ───────────────────────────
def existe(nombre: str) -> bool:
    if nombre == "__PRINCIPAL__":
        return (WORKDIR / NOMBRE_PRINCIPAL).exists()
    return (WORKDIR / nombre).exists()


def guardar_subida(uploaded, nombre_destino: str):
    """Guarda un archivo subido en el directorio de trabajo con nombre canónico."""
    destino = WORKDIR / nombre_destino
    destino.write_bytes(uploaded.getbuffer())
    return destino


@contextlib.contextmanager
def en_workdir():
    """Ejecuta código con el workdir como directorio actual (los scripts usan rutas relativas)."""
    previo = os.getcwd()
    os.chdir(WORKDIR)
    try:
        yield
    finally:
        os.chdir(previo)


def ejecutar_etapa(etapa_id: str, nombre_funcion: str) -> bool:
    """Ejecuta una etapa capturando su salida (los print originales) en un log."""
    buffer = io.StringIO()
    ok = True
    try:
        with en_workdir(), contextlib.redirect_stdout(buffer):
            funcion = getattr(etapas, nombre_funcion)
            if nombre_funcion == "etapa_1_a_4":
                funcion(NOMBRE_PRINCIPAL)
            else:
                funcion()
    except SystemExit:
        ok = False
    except Exception as e:  # noqa: BLE001
        buffer.write(f"\n❌ Excepción no controlada: {e}")
        ok = False
    st.session_state.logs[etapa_id] = buffer.getvalue()
    st.session_state.estado[etapa_id] = "ok" if ok else "error"
    return ok


@st.cache_data(show_spinner=False)
def leer_excel_cacheado(ruta: str, mtime: float) -> pd.DataFrame:
    """Lee un Excel del workdir. mtime invalida el caché cuando el archivo cambia."""
    return pd.read_excel(ruta, dtype={
        "DOCUMENTO": str, "CODPROD": str, "CODVEND": str,
        "RIF": str, "FACTAFECT": str, "CU": str,
    })


def df_de(nombre_archivo: str) -> pd.DataFrame:
    ruta = WORKDIR / nombre_archivo
    return leer_excel_cacheado(str(ruta), ruta.stat().st_mtime)


def boton_descarga(nombre_archivo: str, etiqueta: str | None = None, key: str | None = None):
    ruta = WORKDIR / nombre_archivo
    if ruta.exists():
        st.download_button(
            label=etiqueta or f"⬇️ Descargar {nombre_archivo}",
            data=ruta.read_bytes(),
            file_name=nombre_archivo,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=key or f"dl_{nombre_archivo}",
        )


def a_excel_bytes(df: pd.DataFrame, hoja: str = "Auditoria") -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=hoja)
    return buf.getvalue()


# ─────────────────────────── Sidebar: archivos de entrada ───────────────────────────
with st.sidebar:
    st.header("📂 Archivos de entrada")
    st.caption("Subí los 4 archivos. Se guardan solo durante la sesión.")

    up_principal = st.file_uploader(
        "1️⃣ Listado principal (hoja 'Empresas')", type=["xlsx"], key="up_main")
    if up_principal:
        guardar_subida(up_principal, NOMBRE_PRINCIPAL)
        st.success(f"✅ {up_principal.name}")

    up_clientes = st.file_uploader(
        "2️⃣ CLIENTES_UNICOS.xlsx", type=["xlsx"], key="up_cli")
    if up_clientes:
        guardar_subida(up_clientes, "CLIENTES_UNICOS.xlsx")
        st.success(f"✅ {up_clientes.name}")

    up_descrip = st.file_uploader(
        "3️⃣ DESCRIP_UNICAS.xlsx", type=["xlsx"], key="up_desc")
    if up_descrip:
        guardar_subida(up_descrip, "DESCRIP_UNICAS.xlsx")
        st.success(f"✅ {up_descrip.name}")

    up_tasas = st.file_uploader(
        "4️⃣ TASAS.xlsx (hojas PAR y BCV)", type=["xlsx"], key="up_tasas")
    if up_tasas:
        guardar_subida(up_tasas, "TASAS.xlsx")
        st.success(f"✅ {up_tasas.name}")

    st.divider()
    if st.button("🗑️ Reiniciar sesión (borra todo)", use_container_width=True):
        for f in WORKDIR.glob("*"):
            f.unlink(missing_ok=True)
        st.session_state.logs = {}
        st.session_state.estado = {}
        st.cache_data.clear()
        st.rerun()

# ─────────────────────────── Cabecera ───────────────────────────
st.title("🏗️ Pipeline de Análisis de Ventas")
st.caption(
    "Lógica original del notebook, sin modificaciones. "
    "Cada etapa genera su Excel intermedio, descargable para auditoría."
)

# Barra de progreso del pipeline
completadas = sum(1 for *_, salida in PIPELINE if existe(salida))
st.progress(completadas / len(PIPELINE),
            text=f"Progreso: {completadas} de {len(PIPELINE)} etapas completadas")

# ─────────────────────────── Pestañas principales ───────────────────────────
tab_pipeline, tab_audit5, tab_audit7, tab_resultados = st.tabs([
    "▶️ Pipeline", "🔍 Auditoría Paso 5 (Clientes)",
    "🔍 Auditoría Paso 7 (Descripciones)", "📦 Resultados y descargas",
])

# ═══════════════════════════ TAB: PIPELINE ═══════════════════════════
with tab_pipeline:
    col_run_all, _ = st.columns([1, 2])
    with col_run_all:
        entradas_listas = all(
            existe(a) for a in ["__PRINCIPAL__", "CLIENTES_UNICOS.xlsx",
                                "DESCRIP_UNICAS.xlsx", "TASAS.xlsx"]
        )
        if st.button("🚀 Ejecutar pipeline completo", type="primary",
                     disabled=not entradas_listas, use_container_width=True):
            barra = st.progress(0.0, text="Iniciando…")
            for i, (eid, etiqueta, fn, _reqs, _salida) in enumerate(PIPELINE):
                barra.progress(i / len(PIPELINE), text=f"Ejecutando: {etiqueta}")
                if not ejecutar_etapa(eid, fn):
                    st.error(f"❌ El pipeline se detuvo en: {etiqueta}. Revisá el log abajo.")
                    break
            else:
                barra.progress(1.0, text="✅ Pipeline completado")
                st.balloons()
            st.cache_data.clear()
            st.rerun()
        if not entradas_listas:
            st.info("⬅️ Subí los 4 archivos de entrada en la barra lateral para habilitar la ejecución.")

    st.divider()

    # Etapas individuales
    for eid, etiqueta, fn, reqs, salida in PIPELINE:
        requisitos_ok = all(existe(r) for r in reqs)
        estado = st.session_state.estado.get(eid)
        hecho = existe(salida)

        icono = "✅" if hecho else ("❌" if estado == "error" else "⬜")
        with st.expander(f"{icono} {etiqueta}", expanded=(estado == "error")):
            c1, c2, c3 = st.columns([1, 1, 2])
            with c1:
                if st.button("▶️ Ejecutar", key=f"run_{eid}", disabled=not requisitos_ok):
                    with st.spinner(f"Ejecutando {etiqueta}…"):
                        ejecutar_etapa(eid, fn)
                    st.cache_data.clear()
                    st.rerun()
            with c2:
                boton_descarga(salida, "⬇️ Excel de esta etapa", key=f"dl_{eid}")
            with c3:
                if not requisitos_ok:
                    faltan = [("Listado principal" if r == "__PRINCIPAL__" else r)
                              for r in reqs if not existe(r)]
                    st.warning(f"Falta: {', '.join(faltan)}")

            log = st.session_state.logs.get(eid)
            if log:
                st.code(log, language=None)

# ═══════════════════════════ TAB: AUDITORÍA PASO 5 ═══════════════════════════
with tab_audit5:
    st.subheader("🔍 Auditoría de la estandarización de clientes (Paso 5)")
    if not existe("resultado_paso_5.xlsx"):
        st.info("Ejecutá el Paso 5 para habilitar esta auditoría.")
    else:
        df5 = df_de("resultado_paso_5.xlsx")
        cols_aud = [c for c in ["EMP", "RIF", "NOMCLIE_ORIGINAL", "NOMCLIE",
                                "UNIFICADO", "NOMVEND"] if c in df5.columns]

        total = len(df5)
        unif_si = int((df5.get("UNIFICADO") == "SI").sum())
        unif_no = total - unif_si
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Filas totales", f"{total:,}")
        m2.metric("Unificadas (SI)", f"{unif_si:,}")
        m3.metric("Sin unificar (NO)", f"{unif_no:,}")
        m4.metric("% cobertura", f"{(unif_si / total * 100):.1f}%" if total else "—")

        vista = st.radio(
            "Vista", ["❗ Solo NO unificados (a revisar)", "✏️ Nombre cambió",
                      "Todos"], horizontal=True, key="vista5")

        base = df5[cols_aud].copy()
        if vista.startswith("❗"):
            filtrado = base[base["UNIFICADO"] == "NO"]
        elif vista.startswith("✏️"):
            filtrado = base[
                (base["UNIFICADO"] == "SI")
                & (base["NOMCLIE_ORIGINAL"].astype(str) != base["NOMCLIE"].astype(str))
            ]
        else:
            filtrado = base

        # Vista única por cliente para auditar más rápido
        agrupar = st.toggle("Agrupar por RIF (una fila por cliente)", value=True, key="grp5")
        if agrupar and "RIF" in filtrado.columns:
            filtrado = (filtrado
                        .drop_duplicates(subset=["RIF", "NOMCLIE_ORIGINAL", "NOMCLIE"])
                        .sort_values(["UNIFICADO", "NOMCLIE"]))

        busqueda = st.text_input("Buscar cliente o RIF", key="q5")
        if busqueda:
            q = busqueda.strip().upper()
            filtrado = filtrado[
                filtrado["NOMCLIE"].astype(str).str.upper().str.contains(q, na=False)
                | filtrado["NOMCLIE_ORIGINAL"].astype(str).str.upper().str.contains(q, na=False)
                | filtrado["RIF"].astype(str).str.upper().str.contains(q, na=False)
            ]

        st.caption(f"{len(filtrado):,} filas en la vista actual")
        st.dataframe(filtrado, use_container_width=True, height=420)
        st.download_button(
            "⬇️ Descargar esta vista de auditoría (Excel)",
            data=a_excel_bytes(filtrado, "Auditoria_Paso5"),
            file_name="auditoria_paso_5.xlsx",
            key="dl_aud5",
        )

# ═══════════════════════════ TAB: AUDITORÍA PASO 7 ═══════════════════════════
with tab_audit7:
    st.subheader("🔍 Auditoría de la estandarización de descripciones (Paso 7)")
    if not existe("resultado_paso_7.xlsx"):
        st.info("Ejecutá el Paso 7 para habilitar esta auditoría.")
    else:
        df7 = df_de("resultado_paso_7.xlsx")
        cols_aud7 = [c for c in ["EMP", "TIPO", "CODPROD", "DESCRIP", "FAM",
                                 "CAMBIO_EJECUTADO"] if c in df7.columns]

        total7 = len(df7)
        cambio_si = int((df7.get("CAMBIO_EJECUTADO") == "SI").sum())
        cambio_no = total7 - cambio_si
        n1, n2, n3, n4 = st.columns(4)
        n1.metric("Filas totales", f"{total7:,}")
        n2.metric("Actualizadas (SI)", f"{cambio_si:,}")
        n3.metric("Sin match (NO)", f"{cambio_no:,}")
        n4.metric("% cobertura", f"{(cambio_si / total7 * 100):.1f}%" if total7 else "—")

        vista7 = st.radio(
            "Vista", ["❗ Solo sin match (a incorporar al maestro)", "Todos"],
            horizontal=True, key="vista7")

        base7 = df7[cols_aud7].copy()
        if vista7.startswith("❗"):
            filtrado7 = base7[base7["CAMBIO_EJECUTADO"] == "NO"]
        else:
            filtrado7 = base7

        agrupar7 = st.toggle("Agrupar por EMP+TIPO+CODPROD (una fila por producto)",
                             value=True, key="grp7")
        if agrupar7:
            claves = [c for c in ["EMP", "TIPO", "CODPROD"] if c in filtrado7.columns]
            if claves:
                filtrado7 = (filtrado7
                             .drop_duplicates(subset=claves)
                             .sort_values(claves))

        if "EMP" in filtrado7.columns:
            emps = sorted(filtrado7["EMP"].dropna().astype(str).unique().tolist())
            sel_emp = st.multiselect("Filtrar por empresa", emps, key="emp7")
            if sel_emp:
                filtrado7 = filtrado7[filtrado7["EMP"].astype(str).isin(sel_emp)]

        busqueda7 = st.text_input("Buscar por CODPROD o descripción", key="q7")
        if busqueda7:
            q7 = busqueda7.strip().upper()
            filtrado7 = filtrado7[
                filtrado7["CODPROD"].astype(str).str.upper().str.contains(q7, na=False)
                | filtrado7["DESCRIP"].astype(str).str.upper().str.contains(q7, na=False)
            ]

        st.caption(f"{len(filtrado7):,} filas en la vista actual")
        st.dataframe(filtrado7, use_container_width=True, height=420)
        st.download_button(
            "⬇️ Descargar esta vista de auditoría (Excel)",
            data=a_excel_bytes(filtrado7, "Auditoria_Paso7"),
            file_name="auditoria_paso_7.xlsx",
            key="dl_aud7",
        )

# ═══════════════════════════ TAB: RESULTADOS ═══════════════════════════
with tab_resultados:
    st.subheader("📦 Archivos generados en esta sesión")
    generados = [(etq, sal) for _, etq, _, _, sal in PIPELINE if existe(sal)]
    if not generados:
        st.info("Todavía no hay archivos generados. Ejecutá el pipeline.")
    else:
        for etq, sal in generados:
            c1, c2 = st.columns([3, 1])
            with c1:
                tam = (WORKDIR / sal).stat().st_size / 1024
                st.write(f"**{sal}** — {etq} · {tam:,.0f} KB")
            with c2:
                boton_descarga(sal, "⬇️ Descargar", key=f"res_{sal}")

        if existe("Analisis_Final_DEF.xlsx"):
            st.divider()
            st.subheader("👀 Vista previa del archivo final (Analisis_Final_DEF.xlsx)")
            df_final = df_de("Analisis_Final_DEF.xlsx")
            st.caption(f"{len(df_final):,} filas × {len(df_final.columns)} columnas — mostrando primeras 200")
            st.dataframe(df_final.head(200), use_container_width=True, height=420)
            g1, g2 = st.columns(2)
            with g1:
                if "TIPO DE VENTA" in df_final.columns:
                    st.markdown("**Líneas por TIPO DE VENTA**")
                    st.bar_chart(df_final["TIPO DE VENTA"].value_counts())
            with g2:
                if "MONEDA" in df_final.columns and "CU" in df_final.columns:
                    # MONEDA se propaga a todas las líneas: contar por factura (filas -A)
                    facturas = df_final[df_final["CU"].astype(str).str.endswith("-A")]
                    st.markdown("**Facturas por MONEDA**")
                    st.bar_chart(facturas["MONEDA"].value_counts())
