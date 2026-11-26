import io
import json
import os
import zipfile
from datetime import date, datetime
from decimal import Decimal
from typing import Any, Optional

import altair as alt
import pandas as pd
import streamlit as st
from dotenv import load_dotenv
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from supabase import Client, create_client

from calculos_fiscales import (
    ComprobanteFiscal,
    LineaAlicuotaIVA,
    a_decimal,
    solo_digitos,
    validar_y_calcular_comprobante,
)
from duplicados import (
    ConfigDuplicados,
    RegistroFacturaExistente,
    calcular_dhash_imagen,
    calcular_sha256,
    construir_identidad_fiscal,
    detectar_duplicado,
)
from extractor_facturas import (
    FacturaExtraida,
    extraer_factura,
    generar_nombre_archivo,
    to_dec,
)
from fiscal_config import (
    ParametrosFiscales,
    TIPOS_COMPROBANTE_ARCA,
    ZONA_ARG,
    formatear_fecha_ar,
    formatear_moneda_ar,
    periodo_fiscal_mensual,
)
from planilla_writer import ErrorPlanilla, PlanillaMemoria

load_dotenv()

MAX_UPLOAD_MB = 15

st.set_page_config(
    page_title="3CICON — Portal de Facturas",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# --- ESTILOS DARK MODE EJECUTIVO ---
st.markdown(
    """
    <style>
    .stApp {
        background: radial-gradient(circle at top right, #161E31 0%, #0B0F19 55%, #070A12 100%);
        color: #F8FAFC;
    }
    [data-testid="stSidebar"], [data-testid="collapsedControl"] { display: none; }
    header[data-testid="stHeader"] { background: transparent; }
    h1, h2, h3, h4, p, span, label, .stMarkdown { color: #F1F5F9; }
    .stCaption, [data-testid="stCaptionContainer"] { color: #94A3B8 !important; }

    [data-testid="stVerticalBlockBorderWrapper"] {
        background-color: rgba(17, 24, 39, 0.85);
        border: 1px solid #1F2937 !important;
        border-radius: 16px !important;
        padding: 1.35rem 1.45rem !important;
        box-shadow: 0 10px 30px rgba(0, 0, 0, 0.35);
    }
    .stTextInput > div > div > input,
    .stNumberInput > div > div > input,
    .stDateInput > div > div > input {
        background-color: #0F172A !important;
        color: #F8FAFC !important;
        border: 1px solid #334155 !important;
        border-radius: 10px !important;
    }
    [data-testid="stFileUploaderDropzone"] {
        background-color: #0F172A !important;
        border: 1.5px dashed #334155 !important;
        border-radius: 12px !important;
    }
    [data-testid="stFileUploaderDropzone"]:hover { border-color: #6366F1 !important; }

    .stButton > button, .stFormSubmitButton > button {
        border-radius: 10px !important;
        font-weight: 600 !important;
        padding: 0.62rem 1.15rem !important;
        background-color: #1E293B !important;
        color: #F8FAFC !important;
        border: 1px solid #334155 !important;
    }
    .stButton > button:hover, .stFormSubmitButton > button:hover {
        border-color: #6366F1 !important;
        background-color: #273549 !important;
    }
    .stButton > button[kind="primary"], .stFormSubmitButton > button[kind="primary"] {
        background: linear-gradient(135deg, #6366F1 0%, #4F46E5 100%) !important;
        border: none !important;
        color: #FFFFFF !important;
        box-shadow: 0 4px 14px rgba(99, 102, 241, 0.35);
    }
    .stDownloadButton > button {
        border-radius: 10px !important;
        font-weight: 600 !important;
        padding: 0.7rem 1.2rem !important;
        background: linear-gradient(135deg, #10B981 0%, #059669 100%) !important;
        color: #FFFFFF !important;
        border: none !important;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background-color: #111827;
        padding: 6px;
        border-radius: 12px;
        border: 1px solid #1F2937;
        width: fit-content;
    }
    .stTabs [data-baseweb="tab"] {
        background-color: transparent;
        border-radius: 8px;
        padding: 0.55rem 1.25rem;
        font-weight: 600;
        color: #94A3B8;
        border: none;
    }
    .stTabs [aria-selected="true"] {
        background-color: #6366F1 !important;
        color: #FFFFFF !important;
    }
    div[data-testid="stMetric"] {
        background-color: #111827;
        border: 1px solid #1F2937;
        border-radius: 14px;
        padding: 1rem 1.2rem;
    }
    div[data-testid="stMetricValue"] { color: #FFFFFF !important; font-weight: 700 !important; }
    div[data-testid="stMetricLabel"] { color: #94A3B8 !important; }

    .login-brand { text-align: center; margin-bottom: 1.2rem; }
    .login-pill {
        display: inline-block;
        background: rgba(99, 102, 241, 0.15);
        color: #A5B4FC;
        border: 1px solid rgba(99, 102, 241, 0.35);
        padding: 0.25rem 0.8rem;
        border-radius: 999px;
        font-size: 0.78rem;
        font-weight: 700;
        letter-spacing: 1px;
        text-transform: uppercase;
        margin-bottom: 0.75rem;
    }
    .status-ok {
        display: inline-block;
        background: rgba(16, 185, 129, 0.15);
        color: #34D399;
        border: 1px solid rgba(16, 185, 129, 0.35);
        padding: 0.18rem 0.65rem;
        border-radius: 999px;
        font-size: 0.78rem;
        font-weight: 600;
    }
    .app-title { font-size: 1.6rem; font-weight: 700; color: #FFFFFF; margin: 0; }
    .app-sub { font-size: 0.92rem; color: #94A3B8; margin: 0.15rem 0 0 0; }
    .step-tag {
        color: #818CF8;
        font-size: 0.78rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.8px;
        margin-bottom: 0.25rem;
    }
    .step-title { color: #F8FAFC; font-size: 1.12rem; font-weight: 600; margin: 0 0 0.75rem 0; }
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


@st.cache_resource(show_spinner=False)
def _crear_cliente_supabase(url: str, key: str) -> Client:
    return create_client(url, key)


def obtener_supabase() -> Optional[Client]:
    url = _get_secret_or_env("SUPABASE_URL")
    key = _get_secret_or_env("SUPABASE_KEY")
    if not url or not key:
        return None
    return _crear_cliente_supabase(url, key)


def _sanitizar_celda_excel(valor: Any) -> str:
    txt = "" if valor is None else str(valor)
    return ("'" + txt) if txt.startswith(("=", "+", "-", "@", "\t", "\r")) else txt


def _validar_archivo_subido(file_bytes: bytes, filename: str) -> Optional[str]:
    if len(file_bytes) > MAX_UPLOAD_MB * 1024 * 1024:
        return f"El archivo {filename} supera el límite de {MAX_UPLOAD_MB} MB."
    ext = os.path.splitext(filename)[1].lower()
    if ext == ".pdf" and not file_bytes.startswith(b"%PDF"):
        return f"El archivo {filename} no tiene firma válida de PDF."
    if ext in (".jpg", ".jpeg") and not file_bytes.startswith(b"\xff\xd8\xff"):
        return f"El archivo {filename} no tiene firma válida de JPG."
    if ext == ".png" and not file_bytes.startswith(b"\x89PNG"):
        return f"El archivo {filename} no tiene firma válida de PNG."
    return None


# --- AUTENTICACIÓN Y ROLES ---
def _autenticar_usuario(usuario: str, password: str) -> Optional[tuple[str, str]]:
    """Soporta APP_USERS (JSON: {'usuario': {'password': '...', 'rol': 'admin'}}) o APP_PASSWORD simple."""
    users_json = _get_secret_or_env("APP_USERS")
    if users_json:
        try:
            users_map = json.loads(users_json)
            u_norm = usuario.strip().lower()
            if u_norm in users_map and str(users_map[u_norm].get("password", "")) == password:
                return u_norm, str(users_map[u_norm].get("rol", "operador"))
            return None
        except Exception:
            pass

    app_password = _get_secret_or_env("APP_PASSWORD")
    if app_password and password == app_password:
        return (usuario.strip() or "operador"), "admin"
    return None


if "autenticado" not in st.session_state:
    st.session_state.autenticado = False
    st.session_state.usuario_actual = ""
    st.session_state.rol_actual = "operador"

if (_get_secret_or_env("APP_PASSWORD") or _get_secret_or_env("APP_USERS")) and not st.session_state.autenticado:
    st.markdown("<br><br><br>", unsafe_allow_html=True)
    _, col_center, _ = st.columns([1.3, 1.4, 1.3])
    with col_center:
        with st.container(border=True):
            st.markdown(
                """
                <div class="login-brand">
                    <div class="login-pill">3CICON SAS</div>
                    <div style="font-size:1.7rem;font-weight:700;color:#FFF;">Portal de Facturas</div>
                    <div style="font-size:0.92rem;color:#94A3B8;margin-top:0.35rem;">Ingresá tu usuario y clave</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            with st.form("login_form", border=False):
                usr = st.text_input("Usuario", placeholder="Tu nombre o usuario")
                pwd = st.text_input("Contraseña", type="password", placeholder="••••••••")
                if st.form_submit_button("Ingresar", type="primary", use_container_width=True):
                    auth = _autenticar_usuario(usr, pwd)
                    if auth:
                        st.session_state.autenticado = True
                        st.session_state.usuario_actual = auth[0]
                        st.session_state.rol_actual = auth[1]
                        st.rerun()
                    else:
                        st.error("Usuario o contraseña incorrectos.")
    st.stop()


# --- CAPA DE PERSISTENCIA SUPABASE ---
def cargar_facturas_db() -> list[dict[str, Any]]:
    sb = obtener_supabase()
    if sb is None:
        return []
    resp = sb.table("facturas_compras").select("*").order("fecha_emision", desc=True).limit(1000).execute()
    return resp.data or []


def cargar_auditoria_db() -> list[dict[str, Any]]:
    sb = obtener_supabase()
    if sb is None:
        return []
    resp = sb.table("auditoria_facturas").select("*").order("creado_en", desc=True).limit(300).execute()
    return resp.data or []


def registrar_auditoria(
    accion: str,
    identidad_canonica: Optional[str],
    motivo: str = "",
    detalle: Optional[dict[str, Any]] = None,
) -> None:
    sb = obtener_supabase()
    if sb is None:
        return
    sb.table("auditoria_facturas").insert(
        {
            "identidad_canonica": identidad_canonica,
            "accion": accion,
            "usuario": st.session_state.get("usuario_actual") or "operador",
            "rol": st.session_state.get("rol_actual") or "operador",
            "motivo": motivo,
            "detalle_json": detalle or {},
        }
    ).execute()


def db_a_registros_existentes(filas_db: list[dict[str, Any]]) -> list[RegistroFacturaExistente]:
    salida: list[RegistroFacturaExistente] = []
    for r in filas_db:
        try:
            f_em = date.fromisoformat(str(r["fecha_emision"])[:10])
            salida.append(
                RegistroFacturaExistente(
                    id_registro=str(r.get("id", "")),
                    cuit_emisor=str(r.get("cuit_emisor", "")),
                    razon_social_emisor=str(r.get("razon_social_emisor", "")),
                    codigo_tipo=int(r.get("codigo_tipo", 1)),
                    punto_venta=int(r.get("punto_venta", 0)),
                    numero=int(r.get("numero", 0)),
                    fecha_emision=f_em,
                    total=a_decimal(str(r.get("total", "0.00"))),
                    cae=str(r.get("cae") or ""),
                    sha256=str(r.get("sha256") or ""),
                    dhash=r.get("dhash"),
                    descartada=bool(r.get("descartada", False)),
                )
            )
        except Exception:
            continue
    return salida


# --- CONVERSIÓN Y VALIDACIÓN DE FILAS ---
def _parse_fecha(val: Any) -> Optional[date]:
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


def _fila_a_comprobante_fiscal(row: pd.Series) -> ComprobanteFiscal:
    alic_dec = a_decimal(str(row.get("alicuota_iva") or "0"))
    neto_dec = a_decimal(str(row.get("neto_gravado") or "0"))
    iva_dec = a_decimal(str(row.get("importe_iva") or "0"))
    cod_tipo = int(row.get("codigo_tipo") or 1)
    tipo_info = TIPOS_COMPROBANTE_ARCA.get(cod_tipo)

    lineas_iva: list[LineaAlicuotaIVA] = []
    no_grav = a_decimal(str(row.get("imp_internos_no_gravado") or "0"))

    if tipo_info and tipo_info.discrimina_iva:
        if neto_dec > Decimal("0.00") or iva_dec > Decimal("0.00"):
            lineas_iva.append(LineaAlicuotaIVA(alicuota=alic_dec, neto_gravado=neto_dec, importe_iva=iva_dec))
    else:
        # En Factura B/C el total sin impuestos va como no_gravado si no hay líneas IVA
        if neto_dec > Decimal("0.00"):
            no_grav = no_grav + neto_dec

    percepciones_iibb: dict[str, Decimal] = {}
    p_caba = a_decimal(str(row.get("perc_iibb_caba") or "0"))
    p_pba = a_decimal(str(row.get("perc_iibb_pba") or "0"))
    if p_caba > Decimal("0.00"):
        percepciones_iibb["CABA"] = p_caba
    if p_pba > Decimal("0.00"):
        percepciones_iibb["PBA"] = p_pba

    return ComprobanteFiscal(
        fecha_emision=_parse_fecha(row.get("fecha_emision")),
        codigo_tipo=cod_tipo,
        punto_venta=int(row.get("punto_venta") or 0) or None,
        numero=int(row.get("numero") or 0) or None,
        cuit_emisor=solo_digitos(str(row.get("cuit_emisor") or "")),
        razon_social_emisor=str(row.get("razon_social_emisor") or "").strip().upper(),
        cuit_cliente="30716137488",
        lineas_iva=lineas_iva,
        no_gravado=no_grav,
        percepcion_iva=a_decimal(str(row.get("perc_iva") or "0")),
        percepciones_iibb=percepciones_iibb,
        total_declarado=a_decimal(str(row.get("total") or "0")),
        cae=solo_digitos(str(row.get("cae") or "")),
        moneda=str(row.get("moneda") or "PES").upper().strip(),
    )


def _fila_a_factura_legacy(row: pd.Series) -> FacturaExtraida:
    alic_raw = a_decimal(str(row.get("alicuota_iva") or "0"))
    cod_tipo = int(row.get("codigo_tipo") or 1)
    tipo_info = TIPOS_COMPROBANTE_ARCA.get(cod_tipo)
    alic_val = alic_raw if (tipo_info and tipo_info.discrimina_iva and alic_raw > Decimal("0.00")) else None
    return FacturaExtraida(
        fecha_emision=_parse_fecha(row.get("fecha_emision")),
        codigo_tipo=cod_tipo,
        tipo_letra=tipo_info.letra if tipo_info else str(row.get("tipo_letra") or "A"),
        punto_venta=int(row.get("punto_venta") or 0) or None,
        numero=int(row.get("numero") or 0) or None,
        cuit_emisor=solo_digitos(str(row.get("cuit_emisor") or "")),
        razon_social_emisor=_sanitizar_celda_excel(str(row.get("razon_social_emisor") or "").strip().upper()),
        neto_gravado=to_dec(row.get("neto_gravado")),
        alicuota_iva=alic_val,
        importe_iva=to_dec(row.get("importe_iva")),
        imp_internos_no_gravado=to_dec(row.get("imp_internos_no_gravado")),
        perc_iibb_caba=to_dec(row.get("perc_iibb_caba")),
        perc_iibb_pba=to_dec(row.get("perc_iibb_pba")),
        perc_iva=to_dec(row.get("perc_iva")),
        total=to_dec(row.get("total")),
        cae=solo_digitos(str(row.get("cae") or "")),
        moneda=str(row.get("moneda") or "PES"),
        metodo=str(row.get("metodo") or "MANUAL"),
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


def _generar_libro_iva_compras_xlsx(filas_db: list[dict[str, Any]], titulo_periodo: str) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Libro_IVA_Compras"

    borde = Border(
        left=Side(style="thin", color="D0D7DE"),
        right=Side(style="thin", color="D0D7DE"),
        top=Side(style="thin", color="D0D7DE"),
        bottom=Side(style="thin", color="D0D7DE"),
    )
    fmt_ars = '$ #,##0.00;-$ #,##0.00;"$ 0.00"'

    ws.merge_cells("A1:N1")
    ws["A1"] = f"3CICON SAS — LIBRO IVA COMPRAS ({titulo_periodo})"
    ws["A1"].font = Font(bold=True, size=14, color="1F4E78")

    headers = [
        "Fecha", "Período", "Tipo", "Pto. Vta", "Número", "CUIT", "Proveedor",
        "Neto Gravado", "No Grav. / Otros", "IVA Total", "Perc. IVA",
        "Perc. IIBB CABA", "Perc. IIBB PBA", "Total",
    ]
    ws.row_dimensions[3].height = 24
    for col_i, h in enumerate(headers, start=1):
        c = ws.cell(row=3, column=col_i, value=h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = borde

    fila = 4
    for r in filas_db:
        if r.get("descartada"):
            continue
        tipo_obj = TIPOS_COMPROBANTE_ARCA.get(int(r.get("codigo_tipo", 1)))
        signo = tipo_obj.signo if tipo_obj else 1
        vals = [
            str(r.get("fecha_emision", ""))[:10],
            str(r.get("periodo_fiscal", "")),
            tipo_obj.descripcion if tipo_obj else str(r.get("codigo_tipo", "")),
            int(r.get("punto_venta", 0)),
            int(r.get("numero", 0)),
            str(r.get("cuit_emisor", "")),
            _sanitizar_celda_excel(r.get("razon_social_emisor", "")),
            float(a_decimal(str(r.get("neto_gravado", "0"))) * signo),
            float((a_decimal(str(r.get("no_gravado", "0"))) + a_decimal(str(r.get("imp_internos_otros", "0")))) * signo),
            float(a_decimal(str(r.get("iva_total", "0"))) * signo),
            float(a_decimal(str(r.get("percepcion_iva", "0"))) * signo),
            float(a_decimal(str(r.get("percepcion_iibb_caba", "0"))) * signo),
            float(a_decimal(str(r.get("percepcion_iibb_pba", "0"))) * signo),
            float(a_decimal(str(r.get("total", "0"))) * signo),
        ]
        for col_i, v in enumerate(vals, start=1):
            c = ws.cell(row=fila, column=col_i, value=v)
            c.border = borde
            if col_i >= 8:
                c.number_format = fmt_ars
                c.alignment = Alignment(horizontal="right")
        fila += 1

    # Fila de totales con fórmulas reales
    ws.cell(row=fila, column=1, value="TOTAL").font = Font(bold=True, color="1F4E78")
    for col_i in range(8, 15):
        col_l = get_column_letter(col_i)
        c = ws.cell(row=fila, column=col_i, value=f"=SUM({col_l}4:{col_l}{max(4, fila - 1)})")
        c.font = Font(bold=True, color="1F4E78")
        c.number_format = fmt_ars
        c.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        c.border = borde

    for col_i in range(1, 15):
        ws.column_dimensions[get_column_letter(col_i)].width = 16 if col_i != 7 else 34

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# --- BARRA SUPERIOR ---
sb_ok = obtener_supabase() is not None
badge_db = "Supabase Conectada" if sb_ok else "Sin Supabase (solo memoria)"

c_top1, c_top2, c_top3 = st.columns([4.4, 1.1, 1], vertical_alignment="center")
with c_top1:
    st.markdown(
        f"""
        <div>
            <div class="app-title">3CICON — Gestión de Facturas</div>
            <div class="app-sub">Validación fiscal ARCA, detección de duplicados y exportación automática &nbsp; <span class="status-ok">● {badge_db}</span></div>
        </div>
        """,
        unsafe_allow_html=True,
    )
with c_top2:
    if st.button("Limpiar lote", use_container_width=True):
        for k in ("lote_df", "archivos_raw", "excel_salida", "zip_salida", "reporte_escritura", "decisiones_dup"):
            st.session_state.pop(k, None)
        st.rerun()
with c_top3:
    if st.button("Salir", use_container_width=True):
        st.session_state.autenticado = False
        st.rerun()

st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)

tab_carga, tab_tablero, tab_auditoria = st.tabs(["Cargar Facturas", "Tablero e Histórico", "Auditoría"])


# ==============================================================================
# PESTAÑA 1: CARGAR, VALIDAR Y EXPORTAR
# ==============================================================================
with tab_carga:
    st.markdown("<div style='height: 6px;'></div>", unsafe_allow_html=True)
    col_p1, col_p2 = st.columns(2, gap="medium")

    with col_p1:
        with st.container(border=True):
            st.markdown('<div class="step-tag">Paso 1</div>', unsafe_allow_html=True)
            st.markdown('<div class="step-title">Planilla oficial 3CICON (.xlsx)</div>', unsafe_allow_html=True)
            excel_up = st.file_uploader("Planilla Excel", type=["xlsx"], key="planilla_up", label_visibility="collapsed")
            if excel_up is not None:
                try:
                    est = PlanillaMemoria(excel_up.getvalue()).resumen_estado()
                    st.success(
                        f"Cargadas: {est['cantidad_cargadas']}  •  Próxima fila: {est['proxima_fila']}  •  Libres: {est['filas_libres']}"
                    )
                except ErrorPlanilla as err_p:
                    st.error(str(err_p))

    with col_p2:
        with st.container(border=True):
            st.markdown('<div class="step-tag">Paso 2</div>', unsafe_allow_html=True)
            st.markdown('<div class="step-title">Comprobantes (PDF / Fotos) o Carga Manual</div>', unsafe_allow_html=True)
            facturas_up = st.file_uploader(
                "Facturas",
                type=["pdf", "jpg", "jpeg", "png"],
                accept_multiple_files=True,
                key="facturas_up",
                label_visibility="collapsed",
            )
            if st.button("Agregar factura manualmente (sin archivo)", use_container_width=True):
                nueva_fila = {
                    "incluir": True,
                    "archivo_original": "CARGA_MANUAL",
                    "fecha_emision": datetime.now(ZONA_ARG).strftime("%Y-%m-%d"),
                    "razon_social_emisor": "",
                    "cuit_emisor": "",
                    "tipo_letra": "A",
                    "codigo_tipo": 1,
                    "punto_venta": 1,
                    "numero": 1,
                    "neto_gravado": "0.00",
                    "alicuota_iva": "21.00",
                    "importe_iva": "0.00",
                    "imp_internos_no_gravado": "0.00",
                    "perc_iibb_caba": "0.00",
                    "perc_iibb_pba": "0.00",
                    "perc_iva": "0.00",
                    "total": "0.00",
                    "cae": "",
                    "moneda": "PES",
                    "sha256": "",
                    "dhash": "",
                    "metodo": "MANUAL",
                }
                if "lote_df" not in st.session_state or st.session_state.lote_df.empty:
                    st.session_state.lote_df = pd.DataFrame([nueva_fila])
                else:
                    st.session_state.lote_df = pd.concat(
                        [st.session_state.lote_df, pd.DataFrame([nueva_fila])], ignore_index=True
                    )
                st.rerun()

    if facturas_up:
        st.markdown("<div style='height: 8px;'></div>", unsafe_allow_html=True)
        if st.button(f"Procesar {len(facturas_up)} factura(s) con QR + IA", type="primary", use_container_width=True):
            filas_lote = []
            archivos_raw: dict[str, bytes] = st.session_state.get("archivos_raw", {})
            progreso = st.progress(0)
            estado_txt = st.empty()

            for idx, archivo in enumerate(facturas_up, start=1):
                estado_txt.info(f"Analizando {idx}/{len(facturas_up)}: {archivo.name}")
                contenido = archivo.getvalue()
                err_file = _validar_archivo_subido(contenido, archivo.name)
                if err_file:
                    st.error(err_file)
                    continue

                archivos_raw[archivo.name] = contenido
                sha = calcular_sha256(contenido)
                ext = os.path.splitext(archivo.name)[1].lower()
                dhash = calcular_dhash_imagen(contenido) if ext in (".jpg", ".jpeg", ".png") else ""

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
                            "neto_gravado": str(fac.neto_gravado),
                            "alicuota_iva": str(fac.alicuota_iva if fac.alicuota_iva is not None else Decimal("0.00")),
                            "importe_iva": str(fac.importe_iva),
                            "imp_internos_no_gravado": str(fac.imp_internos_no_gravado),
                            "perc_iibb_caba": str(fac.perc_iibb_caba),
                            "perc_iibb_pba": str(fac.perc_iibb_pba),
                            "perc_iva": str(fac.perc_iva),
                            "total": str(fac.total),
                            "cae": fac.cae,
                            "moneda": fac.moneda or "PES",
                            "sha256": sha,
                            "dhash": dhash or "",
                            "metodo": fac.metodo,
                        }
                    )
                except Exception as e_fac:
                    st.error(f"No se pudo leer {archivo.name}: {e_fac}")
                progreso.progress(idx / len(facturas_up))

            estado_txt.empty()
            progreso.empty()
            if filas_lote:
                st.session_state.lote_df = pd.DataFrame(filas_lote)
                st.session_state.archivos_raw = archivos_raw
                st.session_state.decisiones_dup = {}
                st.success("Lectura completada. Revisá validaciones y duplicados abajo antes de confirmar.")

    # --- PASO 3: REVISIÓN, VALIDACIÓN FISCAL Y DUPLICADOS ---
    if "lote_df" in st.session_state and not st.session_state.lote_df.empty:
        st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
        with st.container(border=True):
            st.markdown('<div class="step-tag">Paso 3</div>', unsafe_allow_html=True)
            st.markdown('<div class="step-title">Revisión editable (montos en texto exacto Decimal)</div>', unsafe_allow_html=True)

            df_edit = st.data_editor(
                st.session_state.lote_df,
                use_container_width=True,
                num_rows="dynamic",
                key="editor_facturas_v3",
                column_config={
                    "incluir": st.column_config.CheckboxColumn("Incluir", default=True),
                    "archivo_original": st.column_config.TextColumn("Archivo", disabled=True),
                    "fecha_emision": st.column_config.TextColumn("Fecha (AAAA-MM-DD)"),
                    "razon_social_emisor": st.column_config.TextColumn("Proveedor"),
                    "cuit_emisor": st.column_config.TextColumn("CUIT"),
                    "tipo_letra": st.column_config.SelectboxColumn("Letra", options=["A", "B", "C", "M"]),
                    "codigo_tipo": st.column_config.NumberColumn("Cód. ARCA", step=1),
                    "punto_venta": st.column_config.NumberColumn("Pto. Vta", step=1),
                    "numero": st.column_config.NumberColumn("Nº Comp.", step=1),
                    "neto_gravado": st.column_config.TextColumn("Neto ($)"),
                    "alicuota_iva": st.column_config.TextColumn("IVA %"),
                    "importe_iva": st.column_config.TextColumn("IVA ($)"),
                    "imp_internos_no_gravado": st.column_config.TextColumn("Col T (NG/Otros)"),
                    "perc_iibb_caba": st.column_config.TextColumn("Col U (IIBB CABA)"),
                    "perc_iibb_pba": st.column_config.TextColumn("Col V (IIBB PBA)"),
                    "perc_iva": st.column_config.TextColumn("Col W (Perc. IVA)"),
                    "total": st.column_config.TextColumn("Total ($)"),
                    "cae": st.column_config.TextColumn("CAE (14 díg.)"),
                    "moneda": st.column_config.TextColumn("Moneda"),
                    "sha256": None,
                    "dhash": None,
                    "metodo": st.column_config.TextColumn("Origen", disabled=True),
                },
            )
            st.session_state.lote_df = df_edit

            # Ejecutar motor fiscal + motor de duplicados sobre cada fila incluida
            existentes_db = db_a_registros_existentes(cargar_facturas_db())
            evaluaciones = []
            if "decisiones_dup" not in st.session_state:
                st.session_state.decisiones_dup = {}

            for idx, row in df_edit.iterrows():
                if not row.get("incluir", True):
                    continue
                comp = _fila_a_comprobante_fiscal(row)
                val_f = validar_y_calcular_comprobante(comp, ParametrosFiscales())

                cand = RegistroFacturaExistente(
                    id_registro=f"LOTE-{idx}",
                    cuit_emisor=comp.cuit_emisor,
                    razon_social_emisor=comp.razon_social_emisor,
                    codigo_tipo=comp.codigo_tipo or 1,
                    punto_venta=comp.punto_venta or 0,
                    numero=comp.numero or 0,
                    fecha_emision=comp.fecha_emision or date.today(),
                    total=comp.total_declarado,
                    cae=comp.cae,
                    sha256=str(row.get("sha256") or ""),
                    dhash=str(row.get("dhash") or "") or None,
                )
                dup_res = detectar_duplicado(cand, existentes_db, ConfigDuplicados())
                evaluaciones.append((idx, row, comp, val_f, dup_res))

            st.markdown("#### Resultado de Validación Fiscal y Duplicados")
            for idx, row, comp, val_f, dup_res in evaluaciones:
                etiqueta = f"Fila {idx + 1} — {comp.razon_social_emisor or 'SIN PROVEEDOR'} ({comp.punto_venta or 0:04d}-{comp.numero or 0:08d}) | Total $ {formatear_moneda_ar(comp.total_declarado)}"
                if dup_res.estado == "DUPLICADA":
                    st.error(f"🚫 **DUPLICADA EXACTA (Bloqueada):** {etiqueta}\n\n{dup_res.senales[0].explicacion}")
                elif dup_res.estado == "POSIBLE_DUPLICADO":
                    motivos_txt = "\n".join(f"- {s.explicacion}" for s in dup_res.senales)
                    st.warning(f"⚠️ **POSIBLE DUPLICADO:** {etiqueta}\n\n{motivos_txt}")
                    c_d1, c_d2 = st.columns([1.2, 2])
                    dec = c_d1.selectbox(
                        "Decisión humana",
                        ["Pendiente de revisión", "No es duplicado (permitir)", "Es duplicado (descartar)"],
                        key=f"dec_dup_{idx}",
                    )
                    mot = c_d2.text_input("Motivo obligatorio (queda en auditoría)", key=f"mot_dup_{idx}")
                    st.session_state.decisiones_dup[idx] = (dec, mot)
                elif val_f.estado in ("ERROR", "INCONSISTENCIA"):
                    msgs = " | ".join(val_f.errores + val_f.inconsistencias)
                    st.error(f"❌ **{val_f.estado} FISCAL:** {etiqueta} — {msgs}")
                elif val_f.estado == "ADVERTENCIA":
                    st.warning(f"⚠️ **ADVERTENCIA:** {etiqueta} — {' | '.join(val_f.advertencias)}")
                else:
                    st.success(f"✅ **OK y ÚNICA:** {etiqueta} (Período {val_f.resumen.periodo_fiscal})")

        # --- PASO 4: CONFIRMAR, GUARDAR EN SUPABASE Y EXPORTAR ---
        st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
        with st.container(border=True):
            st.markdown('<div class="step-tag">Paso 4</div>', unsafe_allow_html=True)
            st.markdown('<div class="step-title">Confirmar en Base de Datos y Exportar Excel + ZIP</div>', unsafe_allow_html=True)

            nc_neg = st.checkbox("Cargar Notas de Crédito en negativo en la planilla 3CICON", value=True)
            permitir_inconsistentes = st.checkbox("Permitir guardar comprobantes con advertencia/inconsistencia revisada", value=False)

            if st.button("Confirmar, Guardar en Supabase y Generar Descargas", type="primary", use_container_width=True):
                planilla = PlanillaMemoria(excel_up.getvalue()) if excel_up is not None else None
                sb = obtener_supabase()
                reporte = []
                items_zip = []
                archivos_raw = st.session_state.get("archivos_raw", {})

                for idx, row, comp, val_f, dup_res in evaluaciones:
                    orig = str(row.get("archivo_original") or "MANUAL")
                    id_f = construir_identidad_fiscal(comp.cuit_emisor, comp.codigo_tipo, comp.punto_venta, comp.numero)
                    clave_can = id_f.clave_canonica() if id_f else None

                    if dup_res.bloquea_guardado:
                        registrar_auditoria("BLOQUEO_DUPLICADO_EXACTO", clave_can, dup_res.senales[0].explicacion)
                        reporte.append({"Archivo": orig, "Estado": "BLOQUEADA (DUPLICADA)", "Detalle": dup_res.senales[0].explicacion})
                        continue

                    dec_txt, mot_txt = st.session_state.decisiones_dup.get(idx, ("", ""))
                    if dup_res.estado == "POSIBLE_DUPLICADO":
                        if dec_txt == "Pendiente de revisión" or not mot_txt.strip():
                            reporte.append({"Archivo": orig, "Estado": "OMITIDA", "Detalle": "Falta decisión o motivo para posible duplicado."})
                            continue
                        if dec_txt == "Es duplicado (descartar)":
                            registrar_auditoria("DESCARTADA_DUPLICADO", clave_can, mot_txt.strip())
                            reporte.append({"Archivo": orig, "Estado": "DESCARTADA", "Detalle": mot_txt.strip()})
                            continue
                        registrar_auditoria("APROBADA_NO_DUPLICADO", clave_can, mot_txt.strip())

                    if val_f.estado == "ERROR" or (val_f.estado == "INCONSISTENCIA" and not permitir_inconsistentes):
                        reporte.append({"Archivo": orig, "Estado": f"BLOQUEADA ({val_f.estado})", "Detalle": " | ".join(val_f.errores + val_f.inconsistencias)})
                        continue

                    # 1. Guardar en Supabase (con UNIQUE sobre identidad_canonica)
                    if sb is not None and clave_can:
                        try:
                            sb.table("facturas_compras").insert(
                                {
                                    "identidad_canonica": clave_can,
                                    "cuit_emisor": comp.cuit_emisor,
                                    "razon_social_emisor": comp.razon_social_emisor,
                                    "codigo_tipo": comp.codigo_tipo,
                                    "punto_venta": comp.punto_venta,
                                    "numero": comp.numero,
                                    "fecha_emision": comp.fecha_emision.isoformat(),
                                    "periodo_fiscal": val_f.resumen.periodo_fiscal,
                                    "moneda": comp.moneda,
                                    "neto_gravado": float(val_f.resumen.total_neto_gravado),
                                    "no_gravado": float(val_f.resumen.total_no_gravado),
                                    "exento": float(val_f.resumen.total_exento),
                                    "iva_total": float(val_f.resumen.total_iva_discriminado),
                                    "percepcion_iva": float(comp.percepcion_iva),
                                    "percepcion_iibb_caba": float(comp.percepciones_iibb.get("CABA", Decimal("0.00"))),
                                    "percepcion_iibb_pba": float(comp.percepciones_iibb.get("PBA", Decimal("0.00"))),
                                    "imp_internos_otros": float(val_f.resumen.total_otros_tributos),
                                    "total": float(comp.total_declarado),
                                    "cae": comp.cae,
                                    "sha256": str(row.get("sha256") or ""),
                                    "dhash": str(row.get("dhash") or "") or None,
                                    "estado_fiscal": val_f.estado,
                                    "estado_duplicado": dup_res.estado,
                                    "motivo_decision": mot_txt.strip(),
                                    "creado_por": st.session_state.get("usuario_actual") or "operador",
                                }
                            ).execute()
                            registrar_auditoria("ALTA_FACTURA", clave_can, mot_txt.strip() or "Carga confirmada")
                        except Exception as e_db:
                            reporte.append({"Archivo": orig, "Estado": "ERROR DB (UNIQUE)", "Detalle": str(e_db)})
                            continue

                    # 2. Escribir en planilla oficial 3CICON si fue subida
                    fac_leg = _fila_a_factura_legacy(row)
                    ext = os.path.splitext(orig)[1].lower() or ".pdf"
                    nombre_ren = generar_nombre_archivo(fac_leg, ext)
                    detalle_excel = "Guardada en DB."
                    if planilla is not None:
                        est_xl, fila_xl, msg_xl = planilla.agregar_factura(fac_leg, nc_en_negativo=nc_neg)
                        detalle_excel = f"DB OK | Excel: {est_xl} ({msg_xl})"
                    if orig in archivos_raw:
                        items_zip.append((nombre_ren, archivos_raw[orig]))

                    reporte.append({"Archivo": orig, "Estado": "CONFIRMADA", "Detalle": detalle_excel, "Nombre Final": nombre_ren})

                if planilla is not None:
                    st.session_state.excel_salida = planilla.exportar_bytes()
                    st.session_state.nombre_excel_salida = excel_up.name
                st.session_state.zip_salida = _crear_zip_renombrado(items_zip) if items_zip else None
                st.session_state.reporte_escritura = pd.DataFrame(reporte)
                st.success("Proceso finalizado. Revisá el reporte y descargá tus archivos.")

            if "reporte_escritura" in st.session_state:
                st.dataframe(st.session_state.reporte_escritura, use_container_width=True, hide_index=True)
                col_d1, col_d2 = st.columns(2, gap="medium")
                with col_d1:
                    if st.session_state.get("excel_salida"):
                        st.download_button(
                            label="Descargar Planilla 3CICON Actualizada (.xlsx)",
                            data=st.session_state.excel_salida,
                            file_name=st.session_state.get("nombre_excel_salida", "3CICON_IVA_COMPRAS.xlsx"),
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
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


# ==============================================================================
# PESTAÑA 2: TABLERO HISTÓRICO Y LIBRO IVA COMPRAS
# ==============================================================================
with tab_tablero:
    st.markdown("<div style='height: 6px;'></div>", unsafe_allow_html=True)
    filas_hist = cargar_facturas_db()
    if not filas_hist:
        with st.container(border=True):
            st.info("Todavía no hay facturas guardadas en Supabase.")
    else:
        periodos = sorted({str(r.get("periodo_fiscal", "")) for r in filas_hist if r.get("periodo_fiscal")}, reverse=True)
        c_f1, c_f2 = st.columns(2)
        per_sel = c_f1.selectbox("Filtrar por período fiscal", ["Todos"] + periodos)
        filas_filt = filas_hist if per_sel == "Todos" else [r for r in filas_hist if r.get("periodo_fiscal") == per_sel]

        proveedores = sorted({str(r.get("razon_social_emisor", "")) for r in filas_filt})
        prov_sel = c_f2.selectbox("Filtrar por proveedor", ["Todos"] + proveedores)
        if prov_sel != "Todos":
            filas_filt = [r for r in filas_filt if r.get("razon_social_emisor") == prov_sel]

        tot_g = sum((a_decimal(str(r.get("total", "0"))) for r in filas_filt), Decimal("0.00"))
        tot_iva = sum((a_decimal(str(r.get("iva_total", "0"))) for r in filas_filt), Decimal("0.00"))
        tot_perc = sum(
            (
                a_decimal(str(r.get("percepcion_iva", "0")))
                + a_decimal(str(r.get("percepcion_iibb_caba", "0")))
                + a_decimal(str(r.get("percepcion_iibb_pba", "0")))
                for r in filas_filt
            ),
            Decimal("0.00"),
        )

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total Comprobantes", f"$ {formatear_moneda_ar(tot_g)}")
        m2.metric("IVA Total", f"$ {formatear_moneda_ar(tot_iva)}")
        m3.metric("Percepciones", f"$ {formatear_moneda_ar(tot_perc)}")
        m4.metric("Cantidad", f"{len(filas_filt)}")

        st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
        with st.container(border=True):
            st.markdown('<div class="step-title">Exportar Libro IVA Compras y Backup</div>', unsafe_allow_html=True)
            libro_bytes = _generar_libro_iva_compras_xlsx(filas_filt, per_sel)
            c_b1, c_b2 = st.columns(2)
            c_b1.download_button(
                "Descargar Libro IVA Compras (.xlsx con fórmulas)",
                data=libro_bytes,
                file_name=f"Libro_IVA_Compras_3CICON_{per_sel}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
            c_b2.download_button(
                "Descargar Backup Completo (.json)",
                data=json.dumps(filas_hist, ensure_ascii=False, indent=2).encode("utf-8"),
                file_name=f"Backup_3CICON_{datetime.now(ZONA_ARG).strftime('%Y%m%d_%H%M')}.json",
                mime="application/json",
                use_container_width=True,
            )

        df_hist_view = pd.DataFrame(
            [
                {
                    "Fecha": str(r.get("fecha_emision", ""))[:10],
                    "Período": r.get("periodo_fiscal"),
                    "Proveedor": r.get("razon_social_emisor"),
                    "CUIT": r.get("cuit_emisor"),
                    "Comprobante": f"{int(r.get('punto_venta', 0)):04d}-{int(r.get('numero', 0)):08d}",
                    "Total": float(r.get("total", 0)),
                    "IVA": float(r.get("iva_total", 0)),
                    "Estado": r.get("estado_fiscal"),
                }
                for r in filas_filt
            ]
        )
        st.dataframe(df_hist_view, use_container_width=True, hide_index=True)


# ==============================================================================
# PESTAÑA 3: AUDITORÍA
# ==============================================================================
with tab_auditoria:
    st.markdown("<div style='height: 6px;'></div>", unsafe_allow_html=True)
    eventos = cargar_auditoria_db()
    if not eventos:
        with st.container(border=True):
            st.info("Todavía no hay eventos registrados en la auditoría.")
    else:
        df_aud = pd.DataFrame(
            [
                {
                    "Fecha/Hora": str(e.get("creado_en", ""))[:19].replace("T", " "),
                    "Usuario": e.get("usuario"),
                    "Rol": e.get("rol"),
                    "Acción": e.get("accion"),
                    "Identidad Fiscal": e.get("identidad_canonica"),
                    "Motivo": e.get("motivo"),
                }
                for e in eventos
            ]
        )
        st.dataframe(df_aud, use_container_width=True, hide_index=True)
