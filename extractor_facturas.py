import base64
import io
import json
import logging
import os
import random
import re
import time
import unicodedata
from datetime import date
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from typing import Literal, Optional
from urllib.parse import unquote

import cv2
import numpy as np
from google import genai
from google.genai import types
from PIL import Image, ImageOps
from pydantic import BaseModel, Field, field_validator
from pypdf import PdfReader

logger = logging.getLogger("facturas_3cicon.extractor")

CUIT_CLIENTE_DEFAULT = "30716137488"  # 3CICON SAS

TIPOS_AFIP = {
    1: ("FACTURA", "A"),
    2: ("NOTA DE DEBITO", "A"),
    3: ("NOTA DE CREDITO", "A"),
    4: ("RECIBO", "A"),
    6: ("FACTURA", "B"),
    7: ("NOTA DE DEBITO", "B"),
    8: ("NOTA DE CREDITO", "B"),
    9: ("RECIBO", "B"),
    11: ("FACTURA", "C"),
    12: ("NOTA DE DEBITO", "C"),
    13: ("NOTA DE CREDITO", "C"),
    15: ("RECIBO", "C"),
    51: ("FACTURA", "M"),
    52: ("NOTA DE DEBITO", "M"),
    53: ("NOTA DE CREDITO", "M"),
    201: ("FCE", "A"),
    202: ("NDE", "A"),
    203: ("NCE", "A"),
    206: ("FCE", "B"),
    207: ("NDE", "B"),
    208: ("NCE", "B"),
    211: ("FCE", "C"),
    212: ("NDE", "C"),
    213: ("NCE", "C"),
}


def solo_digitos(txt: str) -> str:
    return re.sub(r"\D", "", txt or "")


def cuit_valido(cuit: str) -> bool:
    c = solo_digitos(cuit)
    if len(c) != 11:
        return False
    coef = (5, 4, 3, 2, 7, 6, 5, 4, 3, 2)
    resto = 11 - (sum(int(a) * b for a, b in zip(c[:10], coef)) % 11)
    dv = 0 if resto == 11 else 9 if resto == 10 else resto
    return dv == int(c[10])


def formatear_cuit(cuit: str) -> str:
    c = solo_digitos(cuit)
    return f"{c[:2]}-{c[2:10]}-{c[10:]}" if len(c) == 11 else c


def quitar_acentos(texto: str) -> str:
    salida = []
    for ch in unicodedata.normalize("NFC", texto or ""):
        if ch in "ñÑ":
            salida.append(ch)
        else:
            d = unicodedata.normalize("NFD", ch)
            salida.append("".join(c for c in d if unicodedata.category(c) != "Mn"))
    return "".join(salida)


def to_dec(val) -> Decimal:
    if val is None or val == "":
        return Decimal("0.00")
    if isinstance(val, Decimal):
        return val.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    try:
        limpio = str(val).replace("$", "").strip()
        return Decimal(limpio).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    except (InvalidOperation, ValueError):
        return Decimal("0.00")


def es_nota_credito(codigo_tipo: Optional[int]) -> bool:
    return codigo_tipo in TIPOS_AFIP and TIPOS_AFIP[codigo_tipo][0].startswith(("NOTA DE CREDITO", "NCE"))


class FacturaExtraida(BaseModel):
    fecha_emision: Optional[date] = Field(default=None, description="Fecha de emisión en formato AAAA-MM-DD")
    codigo_tipo: Optional[int] = Field(default=None, description="Código numérico AFIP (1=FA A, 3=NC A, 6=FA B, 11=FA C, etc.)")
    tipo_letra: Optional[Literal["A", "B", "C", "M", "OTRO"]] = Field(default="A")
    punto_venta: Optional[int] = Field(default=None, description="Punto de venta entero sin ceros a la izquierda")
    numero: Optional[int] = Field(default=None, description="Número de comprobante entero sin ceros a la izquierda")
    cuit_emisor: str = Field(default="", description="CUIT del proveedor/emisor (11 dígitos sin guiones)")
    razon_social_emisor: str = Field(default="", description="Razón social del proveedor/emisor en MAYÚSCULAS")
    cuit_cliente: str = Field(default="", description="CUIT del cliente/receptor (11 dígitos sin guiones)")
    neto_gravado: Decimal = Field(default=Decimal("0.00"), description="Importe neto gravado")
    alicuota_iva: Optional[Decimal] = Field(default=None, description="Alícuota de IVA (21, 10.5, 27, 0 o null si es B/C/Monotributo)")
    importe_iva: Decimal = Field(default=Decimal("0.00"), description="Importe total de IVA discriminado")
    imp_internos_no_gravado: Decimal = Field(default=Decimal("0.00"), description="Columna T: impuestos internos, no gravado, exento u otros tributos")
    perc_iibb_caba: Decimal = Field(default=Decimal("0.00"), description="Columna U: Percepción IIBB CABA / AGIP")
    perc_iibb_pba: Decimal = Field(default=Decimal("0.00"), description="Columna V: Percepción IIBB Provincia de Buenos Aires / ARBA / BSAS")
    perc_iva: Decimal = Field(default=Decimal("0.00"), description="Columna W: Percepción de IVA")
    total: Decimal = Field(default=Decimal("0.00"), description="Importe total del comprobante")
    cae: str = Field(default="", description="Número de CAE de 14 dígitos si figura")
    moneda: str = Field(default="PES", description="Moneda: PES, DOL, etc.")
    metodo: str = Field(default="GEMINI")
    observaciones: str = Field(default="")

    @field_validator(
        "neto_gravado",
        "importe_iva",
        "imp_internos_no_gravado",
        "perc_iibb_caba",
        "perc_iibb_pba",
        "perc_iva",
        "total",
        mode="before",
    )
    @classmethod
    def _val_dec(cls, v):
        return to_dec(v)

    @field_validator("alicuota_iva", mode="before")
    @classmethod
    def _val_alic(cls, v):
        if v is None or v == "":
            return None
        try:
            d = Decimal(str(v).replace("%", "").strip())
            return d
        except (InvalidOperation, ValueError):
            return None

    @field_validator("cuit_emisor", "cuit_cliente", "cae", mode="before")
    @classmethod
    def _val_digitos(cls, v):
        return solo_digitos(str(v or ""))


SYSTEM_PROMPT = f"""
Sos un auditor contable argentino experto en comprobantes de AFIP/ARCA.
Extraé con máxima precisión los datos de esta factura/comprobante de compra para la empresa cliente 3CICON SAS (CUIT {CUIT_CLIENTE_DEFAULT}).

REGLAS ESTRICTAS:
1. Emisor vs Cliente:
   - El CUIT {CUIT_CLIENTE_DEFAULT} (3CICON SAS) es el CLIENTE (receptor), NUNCA el emisor.
   - El emisor es el PROVEEDOR que emite la factura.
2. Números:
   - Devolvé decimales puros con punto decimal (ej: 154200.47). Sin signo $ ni separadores de miles.
   - Punto de venta y número de comprobante como enteros (ej: 00004-00084169 -> punto_venta=4, numero=84169).
3. Fecha de emisión: formato ISO AAAA-MM-DD. Si dice 07/11/26 es 2026-11-07.
4. IVA y Alícuota:
   - Facturas A o M: indicá alicuota_iva (21, 10.5, 27, etc.) e importe_iva.
   - Facturas B, C o Monotributo: alicuota_iva = null, importe_iva = 0.
5. Clasificación de impuestos y percepciones:
   - perc_iibb_caba: Percepción IIBB CABA / Ciudad Autónoma / AGIP.
   - perc_iibb_pba: Percepción IIBB Provincia de Buenos Aires / BSAS / PBA / ARBA (ej: "ISIB BSAS G.20 2.5%").
   - perc_iva: Percepción de IVA.
   - imp_internos_no_gravado: Impuestos internos, conceptos no gravados, exentos, tasas municipales u otros tributos.
6. Devolvé ÚNICAMENTE JSON válido bajo el esquema FacturaExtraida.
"""


def _obtener_api_key() -> Optional[str]:
    api_key = os.getenv("GEMINI_API_KEY")
    if api_key and api_key.strip():
        return api_key.strip()
    try:
        import streamlit as st

        secret_key = st.secrets.get("GEMINI_API_KEY")
        if secret_key and str(secret_key).strip():
            return str(secret_key).strip()
    except Exception:
        pass
    return None


def _parsear_qr_afip(dato: str) -> Optional[dict]:
    m = re.search(r"[?&]p=([A-Za-z0-9_\-+/=%]+)", dato or "")
    if not m:
        return None
    b64 = unquote(m.group(1))
    b64 += "=" * (-len(b64) % 4)
    try:
        j = json.loads(base64.urlsafe_b64decode(b64.replace("+", "-").replace("/", "_")))
        return {
            "fecha": date.fromisoformat(str(j["fecha"])[:10]),
            "cuit": f'{int(j["cuit"]):011d}',
            "pto_vta": int(j["ptoVta"]),
            "tipo": int(j["tipoCmp"]),
            "nro": int(j["nroCmp"]),
            "importe": to_dec(j["importe"]),
            "moneda": str(j.get("moneda", "PES")),
            "cae": str(int(j["codAut"])) if j.get("codAut") else "",
        }
    except Exception:
        return None


def _decodificar_qr_en_imagen(img_bgr: np.ndarray) -> Optional[dict]:
    det = cv2.QRCodeDetector()
    variantes = [
        img_bgr,
        cv2.cvtColor(img_bgr, cv2.COLOR_BGR2GRAY),
        cv2.resize(img_bgr, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC),
    ]
    for var in variantes:
        try:
            dato, _, _ = det.detectAndDecode(var)
            if dato:
                qr = _parsear_qr_afip(dato)
                if qr:
                    return qr
        except cv2.error:
            pass
    return None


def _buscar_qr_en_bytes(file_bytes: bytes, es_pdf: bool) -> Optional[dict]:
    try:
        if es_pdf:
            reader = PdfReader(io.BytesIO(file_bytes))
            for page in reader.pages[:2]:
                for img_file in page.images:
                    arr = np.frombuffer(img_file.data, np.uint8)
                    img_bgr = cv2.imdecode(arr, cv2.IMREAD_COLOR)
                    if img_bgr is not None:
                        qr = _decodificar_qr_en_imagen(img_bgr)
                        if qr:
                            return qr
        else:
            pil = ImageOps.exif_transpose(Image.open(io.BytesIO(file_bytes))).convert("RGB")
            img_bgr = cv2.cvtColor(np.array(pil), cv2.COLOR_RGB2BGR)
            return _decodificar_qr_en_imagen(img_bgr)
    except Exception:
        return None
    return None


def generar_nombre_archivo(f: FacturaExtraida, extension: str) -> str:
    """Formato pedido por Martín: AAAAMMDD RAZON SOCIAL PTOVTA NRO.ext"""
    fecha_str = f.fecha_emision.strftime("%Y%m%d") if f.fecha_emision else "00000000"
    rs = quitar_acentos(f.razon_social_emisor or "SIN_PROVEEDOR").upper()
    rs = re.sub(r"[^A-Z0-9Ñ ]", "", rs)
    rs = re.sub(r"\s+", " ", rs).strip()[:40] or "SIN PROVEEDOR"
    pv = f"{f.punto_venta:04d}" if f.punto_venta is not None else "0000"
    nro = f"{f.numero:08d}" if f.numero is not None else "00000000"
    ext = extension if extension.startswith(".") else f".{extension}"
    return f"{fecha_str} {rs} {pv} {nro}{ext.lower()}"


def extraer_factura(file_bytes: bytes, filename: str) -> FacturaExtraida:
    api_key = _obtener_api_key()
    if not api_key:
        raise ValueError("Falta configurar GEMINI_API_KEY en Streamlit Secrets o .env.")

    ext = os.path.splitext(filename)[1].lower()
    es_pdf = ext == ".pdf"

    qr = _buscar_qr_en_bytes(file_bytes, es_pdf=es_pdf)

    if es_pdf:
        contenido_doc = types.Part.from_bytes(data=file_bytes, mime_type="application/pdf")
    else:
        pil = ImageOps.exif_transpose(Image.open(io.BytesIO(file_bytes))).convert("RGB")
        pil.thumbnail((2000, 2000), Image.Resampling.LANCZOS)
        contenido_doc = pil

    client = genai.Client(
        api_key=api_key,
        http_options=types.HttpOptions(timeout=45000),
    )

    modelos_candidatos = ["gemini-3.8-flash", "gemini-3.6-flash"]
    ultimo_error: Optional[Exception] = None
    factura: Optional[FacturaExtraida] = None

    for modelo in modelos_candidatos:
        for intento in range(2):
            try:
                response = client.models.generate_content(
                    model=modelo,
                    contents=[
                        contenido_doc,
                        SYSTEM_PROMPT + "\nDevolvé el resultado en JSON bajo el esquema FacturaExtraida.",
                    ],
                    config=types.GenerateContentConfig(
                        response_mime_type="application/json",
                        response_schema=FacturaExtraida,
                        temperature=0.0,
                        thinking_config=types.ThinkingConfig(thinking_level="low"),
                    ),
                )
                if not response.text:
                    raise ValueError(f"Respuesta vacía de {modelo}.")
                factura = FacturaExtraida.model_validate_json(response.text)
                break
            except Exception as e:
                ultimo_error = e
                err_str = str(e)
                if any(c in err_str for c in ("429", "RESOURCE_EXHAUSTED", "404", "NOT_FOUND")):
                    break
                if any(c in err_str for c in ("503", "UNAVAILABLE", "504", "DEADLINE")):
                    time.sleep(2.0 * (intento + 1) + random.uniform(0.1, 0.4))
                    continue
                raise RuntimeError(f"Error al analizar {filename}: {err_str}") from e
        if factura is not None:
            break

    if factura is None:
        raise RuntimeError(f"No se pudo procesar {filename} con IA. Detalle: {ultimo_error}")

    # Si hubo QR de AFIP, sus datos mandan por ser 100% exactos
    if qr:
        factura.fecha_emision = qr["fecha"]
        factura.cuit_emisor = qr["cuit"]
        factura.punto_venta = qr["pto_vta"]
        factura.numero = qr["nro"]
        factura.codigo_tipo = qr["tipo"]
        factura.total = qr["importe"]
        factura.moneda = qr["moneda"]
        if qr["cae"]:
            factura.cae = qr["cae"]
        if qr["tipo"] in TIPOS_AFIP:
            factura.tipo_letra = TIPOS_AFIP[qr["tipo"]][1]
        factura.metodo = "QR+GEMINI"
    else:
        factura.metodo = "GEMINI"

    # Normalizar razón social
    factura.razon_social_emisor = re.sub(r"\s+", " ", (factura.razon_social_emisor or "").strip()).upper()

    # Validaciones de aviso
    avisos = []
    if not cuit_valido(factura.cuit_emisor):
        avisos.append(f"Revisar CUIT emisor ({factura.cuit_emisor}).")
    if factura.moneda not in ("PES", "", None):
        avisos.append(f"Moneda {factura.moneda}: verificar conversión a ARS.")
    if factura.observaciones:
        avisos.append(factura.observaciones)
    factura.observaciones = " | ".join(avisos)

    return factura
