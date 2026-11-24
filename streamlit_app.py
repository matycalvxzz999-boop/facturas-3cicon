import io
import os
import zipfile
from datetime import date, datetime
from decimal import Decimal
from typing import Any, Optional
from zoneinfo import ZoneInfo

import altair as alt
import pandas as pd
import streamlit as st
from dotenv import load_dotenv

from extractor_facturas import (
    FacturaExtraida,
    extraer_factura,
    generar_nombre_archivo,
    solo_digitos,
    to_dec,
)
from planilla_writer import ErrorPlanilla, PlanillaMemoria

load_dotenv()

ZONA_ARG = ZoneInfo("America/Argentina/Buenos_Aires")

st.set_page_config(
    page_title="3CICON — Carga de Facturas IVA Compras",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    .main { background-color: #F8F9FA; }
    .stButton>button { border-radius: 6px; font-weight: 500; }
    .stDownloadButton>button {
        border-radius: 6px;
        font-weight: 600;
        background-color: #1F4E78;
        color: white;
    }
    div[data-testid="stMetricValue"] { font-size: 1.45rem; color: #1F4E78; }
    </style>
    """,
    unsafe_allow_html=True,
)


def _get_secret_or_env(key: str) -> Optional[str]:
    try:
        val = st.secrets.get(key)
        if val is not None and str(val).strip():
            return str(val).strip()
    except Exception:
        pass
    env_val = os.getenv(key)
    return env_val.strip() if env_val and env_val.strip() else None


# --- AUTENTICACIÓN ---
if "autenticado" not in st.session_state:
    st.session_state.autenticado = False

app_password = _get_secret_or_env("APP_PASSWORD")
if app_password and not st.session_state.autenticado:
    c1, c2, c3 = st.columns([1, 2, 1])
    with c2:
        st.markdown("<br><br>", unsafe_allow_html=True)
        st.markdown("## Acceso — 3CICON IVA Compras")
        st.write("Ingresá la contraseña para continuar.")
        pwd = st.text_input("Contraseña:", type="password")
        if st.button("Ingresar", type="primary", use_container_width=True):
            if pwd == app_password:
                st.session_state.autenticado = True
                st.rerun()
            else:
                st.error("Contraseña incorrecta.")
    st.stop()


# --- UTILIDADES ---
def _parse_fecha_celda(val: Any) -> Optional[date]:
    if isinstance(val, date) and not isinstance(val, datetime):
        return val
    if isinstance(val, datetime):
        return val.date()
    txt = str(val or "").strip()
    if not txt:
        return None
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(txt[:10], fmt).date()
        except ValueError:
            pass
    return None


def _fila_a_factura(row: pd.Series) -> FacturaExtraida:
    alic_raw = row.get("alicuota_iva")
    alic_val = None
    if pd.notna(alic_raw) and str(alic_raw).strip() != "":
        try:
            d_al = Decimal(str(alic_raw).replace("%", "").strip())
            alic_val = d_al if d_al > 0 else None
        except Exception:
            alic_val = None

    cod_raw = row.get("codigo_tipo")
    cod_val = int(cod_raw) if pd.notna(cod_raw) and str(cod_raw).strip() != "" else 1

    pv_raw = row.get("punto_venta")
    pv_val = int(pv_raw) if pd.notna(pv_raw) and str(pv_raw).strip() != "" else None

    nro_raw = row.get("numero")
    nro_val = int(nro_raw) if pd.notna(nro_raw) and str(nro_raw).strip() != "" else None

    return FacturaExtraida(
        fecha_emision=_parse_fecha_celda(row.get("fecha_emision")),
        codigo_tipo=cod_val,
        tipo_letra=str(row.get("tipo_letra") or "A"),
        punto_venta=pv_val,
        numero=nro_val,
        cuit_emisor=solo_digitos(str(row.get("cuit_emisor") or "")),
        razon_social_emisor=str(row.get("razon_social_emisor") or "").strip().upper(),
        neto_gravado=to_dec(row.get("neto_gravado")),
        alicuota_iva=alic_val,
        importe_iva=to_dec(row.get("importe_iva")),
        imp_internos_no_gravado=to_dec(row.get("imp_internos_no_gravado")),
        perc_iibb_caba=to_dec(row.get("perc_iibb_caba")),
        perc_iibb_pba=to_dec(row.get("perc_iibb_pba")),
        perc_iva=to_dec(row.get("perc_iva")),
        total=to_dec(row.get("total")),
        metodo=str(row.get("metodo") or "MANUAL"),
        observaciones=str(row.get("observaciones") or ""),
    )


def _crear_zip_renombrado(items_zip: list[tuple[str, bytes]]) -> bytes:
    buf = io.BytesIO()
    usados: dict[str, int] = {}
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for nombre, contenido in items_zip:
            base, ext = os.path.splitext(nombre)
            if nombre in usados:
                usados[nombre] += 1
                nombre_final = f"{base}_{usados[nombre]}{ext}"
            else:
                usados[nombre] = 0
                nombre_final = nombre
            z.writestr(nombre_final, contenido)
    return buf.getvalue()


# --- BARRA LATERAL ---
with st.sidebar:
    st.markdown("### 3CICON — IVA Compras")
    st.caption("Carga automática de facturas a Excel + Renombrado")
    st.divider()
    vista = st.radio(
        "Navegación",
        ["Cargar y Procesar Facturas", "Tablero de Control"],
        label_visibility="collapsed",
    )
    st.divider()
    if st.button("Limpiar lote actual", use_container_width=True):
        for k in ("lote_df", "archivos_raw", "excel_salida", "zip_salida", "reporte_escritura"):
            st.session_state.pop(k, None)
        st.rerun()
    if app_password and st.button("Cerrar Sesión", use_container_width=True):
        st.session_state.autenticado = False
        st.rerun()


# --- VISTA 1: CARGAR Y PROCESAR ---
if vista == "Cargar y Procesar Facturas":
    st.title("Carga de Facturas → Planilla IVA Compras")
    st.write("Subí la planilla oficial del mes y las facturas (PDF o fotos). Revisá la tabla y descargá el Excel actualizado.")

    col_u1, col_u2 = st.columns(2, gap="large")
    with col_u1:
        st.subheader("1. Planilla oficial (.xlsx)")
        excel_up = st.file_uploader(
            "Subir planilla 3CICON IVA COMPRAS",
            type=["xlsx"],
            key="planilla_up",
        )
        planilla_estado = None
        if excel_up is not None:
            try:
                pm_check = PlanillaMemoria(excel_up.getvalue())
                planilla_estado = pm_check.resumen_estado()
                st.success(
                    f"Planilla válida — Cargadas: {planilla_estado['cantidad_cargadas']} | "
                    f"Próxima fila libre: {planilla_estado['proxima_fila']} | "
                    f"Filas disponibles: {planilla_estado['filas_libres']}"
                )
            except ErrorPlanilla as err_p:
                st.error(str(err_p))

    with col_u2:
        st.subheader("2. Facturas nuevas (PDF / JPG / PNG)")
        facturas_up = st.file_uploader(
            "Podés arrastrar varias facturas juntas",
            type=["pdf", "jpg", "jpeg", "png"],
            accept_multiple_files=True,
            key="facturas_up",
        )

    if facturas_up:
        if st.button(f"Procesar {len(facturas_up)} factura(s) con IA + QR", type="primary", use_container_width=True):
            filas_lote = []
            archivos_raw: dict[str, bytes] = {}
            progreso = st.progress(0)
            estado_txt = st.empty()

            for idx, archivo in enumerate(facturas_up, start=1):
                estado_txt.text(f"Analizando [{idx}/{len(facturas_up)}]: {archivo.name} ...")
                contenido = archivo.getvalue()
                archivos_raw[archivo.name] = contenido
                ext = os.path.splitext(archivo.name)[1].lower()
                try:
                    fac = extraer_factura(contenido, archivo.name)
                    filas_lote.append(
                        {
                            "incluir": True,
                            "archivo_original": archivo.name,
                            "fecha_emision": fac.fecha_emision.strftime("%Y-%m-%d") if fac.fecha_emision else "",
                            "razon_social_emisor": fac.razon_social_emisor,
                            "cuit_emisor": fac.cuit_emisor,
                            "tipo_letra": fac.tipo_letra or "A",
                            "codigo_tipo": fac.codigo_tipo or 1,
                            "punto_venta": fac.punto_venta or 0,
                            "numero": fac.numero or 0,
                            "total": float(fac.total),
                            "alicuota_iva": float(fac.alicuota_iva) if fac.alicuota_iva is not None else 0.0,
                            "importe_iva": float(fac.importe_iva),
                            "imp_internos_no_gravado": float(fac.imp_internos_no_gravado),
                            "perc_iibb_caba": float(fac.perc_iibb_caba),
                            "perc_iibb_pba": float(fac.perc_iibb_pba),
                            "perc_iva": float(fac.perc_iva),
                            "nombre_renombrado": generar_nombre_archivo(fac, ext),
                            "metodo": fac.metodo,
                            "observaciones": fac.observaciones,
                        }
                    )
                except Exception as e_fac:
                    filas_lote.append(
                        {
                            "incluir": False,
                            "archivo_original": archivo.name,
                            "fecha_emision": "",
                            "razon_social_emisor": "",
                            "cuit_emisor": "",
                            "tipo_letra": "A",
                            "codigo_tipo": 1,
                            "punto_venta": 0,
                            "numero": 0,
                            "total": 0.0,
                            "alicuota_iva": 21.0,
                            "importe_iva": 0.0,
                            "imp_internos_no_gravado": 0.0,
                            "perc_iibb_caba": 0.0,
                            "perc_iibb_pba": 0.0,
                            "perc_iva": 0.0,
                            "nombre_renombrado": archivo.name,
                            "metodo": "ERROR",
                            "observaciones": str(e_fac),
                        }
                    )
                progreso.progress(idx / len(facturas_up))

            estado_txt.empty()
            st.session_state.lote_df = pd.DataFrame(filas_lote)
            st.session_state.archivos_raw = archivos_raw
            st.session_state.pop("excel_salida", None)
            st.session_state.pop("zip_salida", None)
            st.session_state.pop("reporte_escritura", None)
            st.success("Extracción terminada. Revisá la tabla de abajo antes de generar el Excel.")

    # --- REVISIÓN EDITABLE ---
    if "lote_df" in st.session_state and not st.session_state.lote_df.empty:
        st.divider()
        st.subheader("3. Revisar y Editar antes de escribir en la Planilla")
        st.caption("Podés corregir cualquier monto, fecha, CUIT o alícuota. Destildá 'Incluir' si querés omitir una fila.")

        df_edit = st.data_editor(
            st.session_state.lote_df,
            use_container_width=True,
            num_rows="dynamic",
            key="editor_facturas",
            column_config={
                "incluir": st.column_config.CheckboxColumn("Incluir", default=True),
                "archivo_original": st.column_config.TextColumn("Archivo original", disabled=True),
                "fecha_emision": st.column_config.TextColumn("Fecha (AAAA-MM-DD)"),
                "razon_social_emisor": st.column_config.TextColumn("Proveedor"),
                "cuit_emisor": st.column_config.TextColumn("CUIT (11 díg.)"),
                "tipo_letra": st.column_config.SelectboxColumn("Letra", options=["A", "B", "C", "M", "OTRO"]),
                "codigo_tipo": st.column_config.NumberColumn("Cód. AFIP", step=1),
                "punto_venta": st.column_config.NumberColumn("Pto. Vta", step=1),
                "numero": st.column_config.NumberColumn("Nº Comp.", step=1),
                "total": st.column_config.NumberColumn("Total ($)", format="$ %.2f"),
                "alicuota_iva": st.column_config.NumberColumn("IVA %", format="%.2f"),
                "importe_iva": st.column_config.NumberColumn("IVA ($)", format="$ %.2f"),
                "imp_internos_no_gravado": st.column_config.NumberColumn("Col T (Imp/NG)", format="$ %.2f"),
                "perc_iibb_caba": st.column_config.NumberColumn("Col U (IIBB CABA)", format="$ %.2f"),
                "perc_iibb_pba": st.column_config.NumberColumn("Col V (IIBB PBA)", format="$ %.2f"),
                "perc_iva": st.column_config.NumberColumn("Col W (Perc. IVA)", format="$ %.2f"),
                "nombre_renombrado": st.column_config.TextColumn("Nombre renombrado"),
                "metodo": st.column_config.TextColumn("Método", disabled=True),
                "observaciones": st.column_config.TextColumn("Observaciones"),
            },
        )
        st.session_state.lote_df = df_edit

        st.divider()
        st.subheader("4. Generar Excel Actualizado y Archivos Renombrados")

        nc_neg = st.checkbox("Cargar Notas de Crédito en negativo", value=True)

        if st.button("Escribir en Planilla y Preparar Descargas", type="primary", use_container_width=True):
            if excel_up is None:
                st.error("Primero subí la planilla oficial (.xlsx) en el paso 1.")
            else:
                try:
                    planilla = PlanillaMemoria(excel_up.getvalue())
                    filas_incluidas = df_edit[df_edit["incluir"] == True].copy()

                    # Ordenar cronológicamente antes de escribir
                    pares = []
                    for _, row in filas_incluidas.iterrows():
                        fac = _fila_a_factura(row)
                        pares.append((row, fac))
                    pares.sort(
                        key=lambda x: (
                            x[1].fecha_emision or date(2099, 12, 31),
                            x[1].punto_venta or 0,
                            x[1].numero or 0,
                        )
                    )

                    reporte = []
                    items_zip = []
                    archivos_raw = st.session_state.get("archivos_raw", {})

                    for row, fac in pares:
                        orig = str(row.get("archivo_original") or "")
                        ext = os.path.splitext(orig)[1].lower() or ".pdf"
                        nombre_ren = generar_nombre_archivo(fac, ext)
                        estado, fila_nro, msg = planilla.agregar_factura(fac, nc_en_negativo=nc_neg)
                        reporte.append(
                            {
                                "Archivo": orig,
                                "Estado": estado,
                                "Fila Excel": fila_nro or "",
                                "Proveedor": fac.razon_social_emisor,
                                "Comprobante": f"{fac.punto_venta or 0:04d}-{fac.numero or 0:08d}",
                                "Total": float(fac.total),
                                "Nombre Final": nombre_ren,
                                "Detalle": msg,
                            }
                        )
                        if estado == "OK" and orig in archivos_raw:
                            items_zip.append((nombre_ren, archivos_raw[orig]))

                    st.session_state.excel_salida = planilla.exportar_bytes()
                    st.session_state.nombre_excel_salida = excel_up.name
                    st.session_state.zip_salida = _crear_zip_renombrado(items_zip) if items_zip else None
                    st.session_state.reporte_escritura = pd.DataFrame(reporte)
                    st.success("¡Listo! Ya podés descargar la planilla actualizada y el ZIP con las facturas renombradas.")
                except ErrorPlanilla as err_e:
                    st.error(str(err_e))

        if "reporte_escritura" in st.session_state:
            st.dataframe(st.session_state.reporte_escritura, use_container_width=True)

            col_d1, col_d2 = st.columns(2)
            with col_d1:
                st.download_button(
                    label="Descargar Planilla Excel Actualizada (.xlsx)",
                    data=st.session_state.excel_salida,
                    file_name=st.session_state.get("nombre_excel_salida", "3CICON_IVA_COMPRAS_ACTUALIZADO.xlsx"),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True,
                )
            with col_d2:
                if st.session_state.get("zip_salida"):
                    fecha_hoy = datetime.now(ZONA_ARG).strftime("%Y%m%d_%H%M")
                    st.download_button(
                        label="Descargar Facturas Renombradas (.zip)",
                        data=st.session_state.zip_salida,
                        file_name=f"Facturas_Renombradas_{fecha_hoy}.zip",
                        mime="application/zip",
                        use_container_width=True,
                    )

# --- VISTA 2: TABLERO DE CONTROL ---
elif vista == "Tablero de Control":
    st.title("Tablero de Control — Lote de Compras")
    if "lote_df" not in st.session_state or st.session_state.lote_df.empty:
        st.info("Procesá al menos una factura en 'Cargar y Procesar Facturas' para ver el tablero de control.")
    else:
        df_t = st.session_state.lote_df[st.session_state.lote_df["incluir"] == True].copy()
        tot_general = df_t["total"].sum()
        tot_iva = df_t["importe_iva"].sum()
        tot_imp = (
            df_t["imp_internos_no_gravado"].sum()
            + df_t["perc_iibb_caba"].sum()
            + df_t["perc_iibb_pba"].sum()
            + df_t["perc_iva"].sum()
        )
        tot_neto = max(0.0, tot_general - tot_iva - tot_imp)

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total Facturado", f"$ {tot_general:,.2f}")
        m2.metric("IVA Crédito Fiscal", f"$ {tot_iva:,.2f}")
        m3.metric("Impuestos / Percepciones", f"$ {tot_imp:,.2f}")
        m4.metric("Comprobantes", f"{len(df_t)}")

        st.divider()
        c_g1, c_g2 = st.columns(2)

        with c_g1:
            st.subheader("Composición del Gasto (%)")
            df_comp = pd.DataFrame(
                {
                    "Concepto": ["Neto Gravado", "IVA", "Impuestos / Percepciones"],
                    "Monto": [tot_neto, tot_iva, tot_imp],
                }
            )
            df_comp = df_comp[df_comp["Monto"] > 0]
            if not df_comp.empty:
                df_comp["Porcentaje"] = (df_comp["Monto"] / df_comp["Monto"].sum()) * 100
                df_comp["Etiqueta"] = df_comp["Porcentaje"].map(lambda p: f"{p:.1f}%")
                base = alt.Chart(df_comp).encode(
                    theta=alt.Theta("Monto:Q", stack=True),
                    color=alt.Color("Concepto:N", scale=alt.Scale(range=["#1F4E78", "#2ECC71", "#E67E22"])),
                    tooltip=[
                        "Concepto",
                        alt.Tooltip("Monto:Q", format="$,.2f"),
                        alt.Tooltip("Porcentaje:Q", format=".1f", title="%"),
                    ],
                )
                pie = base.mark_arc(innerRadius=50, outerRadius=115)
                txt = base.mark_text(radius=82, size=12, fontWeight="bold", color="#FFFFFF").encode(text="Etiqueta:N")
                st.altair_chart((pie + txt).properties(height=340), use_container_width=True)

        with c_g2:
            st.subheader("Top Proveedores por Monto")
            df_prov = (
                df_t.groupby("razon_social_emisor", as_index=False)["total"]
                .sum()
                .sort_values("total", ascending=False)
                .head(10)
            )
            if not df_prov.empty:
                bar = (
                    alt.Chart(df_prov)
                    .mark_bar(color="#1F4E78")
                    .encode(
                        x=alt.X("total:Q", title="Total ($)"),
                        y=alt.Y("razon_social_emisor:N", sort="-x", title="Proveedor"),
                        tooltip=["razon_social_emisor", alt.Tooltip("total:Q", format="$,.2f")],
                    )
                    .properties(height=340)
                )
                st.altair_chart(bar, use_container_width=True)
