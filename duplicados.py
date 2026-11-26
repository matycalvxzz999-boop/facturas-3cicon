from __future__ import annotations

import hashlib
import io
from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal
from typing import Literal, Optional

from PIL import Image

from calculos_fiscales import a_decimal, cuit_valido, solo_digitos
from fiscal_config import TIPOS_COMPROBANTE_ARCA, formatear_fecha_ar, formatear_moneda_ar

CERO = Decimal("0.00")

# Confusiones típicas de OCR (0/O, 1/I/L, 5/S, 8/B, 2/Z, 6/G)
_MAPA_OCR = str.maketrans({
    "O": "0", "Q": "0", "D": "0",
    "I": "1", "L": "1", "|": "1",
    "Z": "2",
    "S": "5",
    "G": "6",
    "B": "8",
})


@dataclass(frozen=True)
class IdentidadFiscal:
    cuit_emisor: str
    codigo_tipo: int
    punto_venta: int
    numero: int

    def clave_canonica(self) -> str:
        return f"{self.cuit_emisor}|{self.codigo_tipo}|{self.punto_venta}|{self.numero}"


def construir_identidad_fiscal(
    cuit_emisor: str,
    codigo_tipo: Optional[int],
    punto_venta: Optional[int | str],
    numero: Optional[int | str],
) -> Optional[IdentidadFiscal]:
    """Normaliza la identidad fiscal canónica: CUIT 11 dígitos válido + tipo ARCA + pto vta + nro sin ceros."""
    cuit = solo_digitos(cuit_emisor)
    if not cuit_valido(cuit):
        return None
    if codigo_tipo is None or int(codigo_tipo) not in TIPOS_COMPROBANTE_ARCA:
        return None
    pv_str = solo_digitos(str(punto_venta or ""))
    nro_str = solo_digitos(str(numero or ""))
    if not pv_str or not nro_str:
        return None
    pv_int, nro_int = int(pv_str), int(nro_str)
    if pv_int <= 0 or nro_int <= 0:
        return None
    return IdentidadFiscal(cuit_emisor=cuit, codigo_tipo=int(codigo_tipo), punto_venta=pv_int, numero=nro_int)


def calcular_sha256(contenido: bytes) -> str:
    return hashlib.sha256(contenido).hexdigest()


def calcular_dhash_imagen(contenido: bytes, hash_size: int = 8) -> Optional[str]:
    """Hash perceptual (dHash 64-bit) para detectar fotos/escaneos repetidos aunque cambie el nombre o compresión."""
    try:
        img = Image.open(io.BytesIO(contenido)).convert("L").resize((hash_size + 1, hash_size), Image.Resampling.LANCZOS)
        pixels = list(img.getdata())
        bits = []
        for row in range(hash_size):
            offset = row * (hash_size + 1)
            for col in range(hash_size):
                bits.append("1" if pixels[offset + col] > pixels[offset + col + 1] else "0")
        return f"{int(''.join(bits), 2):016x}"
    except Exception:
        return None


def distancia_hamming_hex(h1: Optional[str], h2: Optional[str]) -> Optional[int]:
    if not h1 or not h2 or len(h1) != len(h2):
        return None
    try:
        return bin(int(h1, 16) ^ int(h2, 16)).count("1")
    except ValueError:
        return None


def distancia_levenshtein(a: str, b: str) -> int:
    if a == b:
        return 0
    if not a:
        return len(b)
    if not b:
        return len(a)
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, start=1):
        curr = [i]
        for j, cb in enumerate(b, start=1):
            costo = 0 if ca == cb else 1
            curr.append(min(prev[j] + 1, curr[j - 1] + 1, prev[j - 1] + costo))
        prev = curr
    return prev[-1]


def normalizar_numero_ocr(valor: Optional[int | str]) -> str:
    txt = str(valor or "").upper().strip().translate(_MAPA_OCR)
    dig = solo_digitos(txt)
    return str(int(dig)) if dig else ""


@dataclass
class RegistroFacturaExistente:
    id_registro: str
    cuit_emisor: str
    razon_social_emisor: str
    codigo_tipo: int
    punto_venta: int
    numero: int
    fecha_emision: date
    total: Decimal
    cae: str = ""
    sha256: str = ""
    dhash: Optional[str] = None
    es_pago_parcial: bool = False
    descartada: bool = False

    def __post_init__(self):
        self.cuit_emisor = solo_digitos(self.cuit_emisor)
        self.total = a_decimal(self.total)
        self.cae = solo_digitos(self.cae)

    @property
    def identidad(self) -> Optional[IdentidadFiscal]:
        return construir_identidad_fiscal(self.cuit_emisor, self.codigo_tipo, self.punto_venta, self.numero)


@dataclass
class ConfigDuplicados:
    ventana_dias_fecha_cercana: int = 7
    tolerancia_importe: Decimal = Decimal("0.06")
    max_distancia_numero: int = 2
    max_distancia_dhash: int = 6


@dataclass
class SenalDuplicado:
    regla: str
    severidad: Literal["EXACTA", "BLANDA"]
    id_factura_original: str
    explicacion: str


@dataclass
class ResultadoDeteccionDuplicados:
    estado: Literal["DUPLICADA", "POSIBLE_DUPLICADO", "UNICA"]
    identidad_canonica: Optional[str]
    bloquea_guardado: bool
    senales: list[SenalDuplicado] = field(default_factory=list)
    factura_original: Optional[RegistroFacturaExistente] = None


def _es_relacion_nc_nd(cod_a: int, cod_b: int) -> bool:
    """True si uno es Factura y el otro es Nota de Crédito o Débito (no son duplicados entre sí)."""
    ta = TIPOS_COMPROBANTE_ARCA.get(cod_a)
    tb = TIPOS_COMPROBANTE_ARCA.get(cod_b)
    if not ta or not tb:
        return False
    return (ta.signo != tb.signo) or (cod_a != cod_b and ("NOTA DE DEBITO" in ta.descripcion or "NOTA DE DEBITO" in tb.descripcion))


def detectar_duplicado(
    candidata: RegistroFacturaExistente,
    existentes: list[RegistroFacturaExistente],
    cfg: Optional[ConfigDuplicados] = None,
) -> ResultadoDeteccionDuplicados:
    """
    Evalúa una factura candidata contra el histórico existente.
    - DUPLICADA (exacta): misma identidad fiscal canónica (CUIT + tipo + pto_vta + numero). Bloquea guardado.
    - POSIBLE_DUPLICADO: señales blandas para revisión humana lado a lado.
    - UNICA: sin señales.
    """
    cfg = cfg or ConfigDuplicados()

    # Casos legítimos que nunca deben marcarse como duplicado de factura:
    if candidata.es_pago_parcial:
        id_cand = candidata.identidad
        return ResultadoDeteccionDuplicados(
            estado="UNICA",
            identidad_canonica=id_cand.clave_canonica() if id_cand else None,
            bloquea_guardado=False,
            senales=[],
        )

    tipo_cand = TIPOS_COMPROBANTE_ARCA.get(candidata.codigo_tipo)
    if tipo_cand is not None and not tipo_cand.es_comprobante_fiscal:
        return ResultadoDeteccionDuplicados(
            estado="UNICA",
            identidad_canonica=None,
            bloquea_guardado=False,
            senales=[],
        )

    id_cand = candidata.identidad
    clave_cand = id_cand.clave_canonica() if id_cand else None
    senales_blandas: list[SenalDuplicado] = []
    primera_orig_blanda: Optional[RegistroFacturaExistente] = None

    for ex in existentes:
        if ex.descartada or ex.es_pago_parcial:
            continue
        tipo_ex = TIPOS_COMPROBANTE_ARCA.get(ex.codigo_tipo)
        if tipo_ex is not None and not tipo_ex.es_comprobante_fiscal:
            continue

        id_ex = ex.identidad
        clave_ex = id_ex.clave_canonica() if id_ex else None

        # --- 1. DUPLICADO EXACTO (identidad fiscal canónica) ---
        if clave_cand and clave_ex and clave_cand == clave_ex:
            msg = (
                f"Duplicado exacto: ya existe {tipo_ex.descripcion if tipo_ex else ex.codigo_tipo} "
                f"{ex.punto_venta:04d}-{ex.numero:08d} del proveedor {ex.razon_social_emisor} "
                f"(CUIT {ex.cuit_emisor}, ID {ex.id_registro}, fecha {formatear_fecha_ar(ex.fecha_emision)}, "
                f"total $ {formatear_moneda_ar(ex.total)})."
            )
            return ResultadoDeteccionDuplicados(
                estado="DUPLICADA",
                identidad_canonica=clave_cand,
                bloquea_guardado=True,
                senales=[SenalDuplicado("IDENTIDAD_FISCAL_EXACTA", "EXACTA", ex.id_registro, msg)],
                factura_original=ex,
            )

        es_mismo_archivo = bool(candidata.sha256 and ex.sha256 and candidata.sha256 == ex.sha256)
        dist_img = distancia_hamming_hex(candidata.dhash, ex.dhash)
        es_imagen_similar = dist_img is not None and dist_img <= cfg.max_distancia_dhash

        if es_mismo_archivo:
            senales_blandas.append(
                SenalDuplicado(
                    "ARCHIVO_SHA256_REPETIDO",
                    "BLANDA",
                    ex.id_registro,
                    f"El archivo subido es idéntico (SHA-256) al comprobante ya cargado ID {ex.id_registro} ({ex.razon_social_emisor}).",
                )
            )
            primera_orig_blanda = primera_orig_blanda or ex
            continue

        if es_imagen_similar:
            senales_blandas.append(
                SenalDuplicado(
                    "IMAGEN_PERCEPTUAL_SIMILAR",
                    "BLANDA",
                    ex.id_registro,
                    f"La imagen es visualmente casi idéntica (distancia perceptual {dist_img}) al comprobante ID {ex.id_registro}.",
                )
            )
            primera_orig_blanda = primera_orig_blanda or ex
            continue

        # --- 2. Mismo CAE con distinta identidad ---
        if candidata.cae and ex.cae and len(candidata.cae) == 14 and candidata.cae == ex.cae:
            senales_blandas.append(
                SenalDuplicado(
                    "MISMO_CAE_DISTINTA_IDENTIDAD",
                    "BLANDA",
                    ex.id_registro,
                    f"Mismo CAE ({candidata.cae}) que el comprobante ID {ex.id_registro} pero con distinta identidad fiscal.",
                )
            )
            primera_orig_blanda = primera_orig_blanda or ex
            continue

        # A partir de acá solo comparamos si es el mismo proveedor (mismo CUIT)
        if not candidata.cuit_emisor or candidata.cuit_emisor != ex.cuit_emisor:
            continue

        # Relación legítima Factura <-> NC/ND o distinto tipo (ej. MiPyME vs Factura A)
        if _es_relacion_nc_nd(candidata.codigo_tipo, ex.codigo_tipo) or candidata.codigo_tipo != ex.codigo_tipo:
            continue

        mismo_importe = abs(candidata.total - ex.total) <= cfg.tolerancia_importe
        dif_dias = abs((candidata.fecha_emision - ex.fecha_emision).days)
        fecha_cercana = dif_dias <= cfg.ventana_dias_fecha_cercana

        # --- 3. Mismo CUIT y mismo número pero distinto importe ---
        if candidata.numero == ex.numero and not mismo_importe:
            senales_blandas.append(
                SenalDuplicado(
                    "MISMO_CUIT_Y_NUMERO_DISTINTO_IMPORTE",
                    "BLANDA",
                    ex.id_registro,
                    f"Mismo proveedor y Nº {candidata.numero} que ID {ex.id_registro}, pero con importe distinto "
                    f"($ {formatear_moneda_ar(candidata.total)} vs $ {formatear_moneda_ar(ex.total)}).",
                )
            )
            primera_orig_blanda = primera_orig_blanda or ex
            continue

        # --- 4. Número casi igual (1 o 2 caracteres de diferencia, típico error OCR) en fecha cercana ---
        nro_a = normalizar_numero_ocr(candidata.numero)
        nro_b = normalizar_numero_ocr(ex.numero)
        dist_nro = distancia_levenshtein(nro_a, nro_b)
        if 1 <= dist_nro <= cfg.max_distancia_numero and len(nro_a) >= 4 and len(nro_b) >= 4 and fecha_cercana:
            senales_blandas.append(
                SenalDuplicado(
                    "NUMERO_CASI_IGUAL_OCR",
                    "BLANDA",
                    ex.id_registro,
                    f"Número muy similar ({candidata.numero} vs {ex.numero}, distancia {dist_nro}) al comprobante ID {ex.id_registro} "
                    f"del mismo proveedor (posible error de lectura OCR).",
                )
            )
            primera_orig_blanda = primera_orig_blanda or ex
            continue

        # --- 5. Mismo CUIT, mismo importe y fecha cercana (±N días) con número distinto ---
        if mismo_importe and fecha_cercana and candidata.numero != ex.numero:
            senales_blandas.append(
                SenalDuplicado(
                    "MISMO_CUIT_IMPORTE_Y_FECHA_CERCANA",
                    "BLANDA",
                    ex.id_registro,
                    f"Mismo proveedor e importe ($ {formatear_moneda_ar(candidata.total)}) con fecha cercana "
                    f"({dif_dias} día(s) de diferencia respecto a ID {ex.id_registro}, Nº {ex.numero}).",
                )
            )
            primera_orig_blanda = primera_orig_blanda or ex

    if senales_blandas:
        return ResultadoDeteccionDuplicados(
            estado="POSIBLE_DUPLICADO",
            identidad_canonica=clave_cand,
            bloquea_guardado=False,
            senales=senales_blandas,
            factura_original=primera_orig_blanda,
        )

    return ResultadoDeteccionDuplicados(
        estado="UNICA",
        identidad_canonica=clave_cand,
        bloquea_guardado=False,
        senales=[],
    )


def comparar_qr_con_ia(datos_qr: dict, datos_ia: dict, tolerancia: Decimal = Decimal("0.06")) -> list[str]:
    """
    Compara los datos del QR oficial ARCA (fuente principal) contra los extraídos por Gemini/OCR.
    Devuelve lista de discrepancias para revisión humana (no oculta diferencias).
    """
    discrepancias: list[str] = []
    if not datos_qr:
        return discrepancias

    if datos_ia.get("cuit_emisor") and solo_digitos(datos_ia["cuit_emisor"]) != solo_digitos(str(datos_qr.get("cuit", ""))):
        discrepancias.append(f"CUIT difiere: QR={datos_qr.get('cuit')} vs IA={datos_ia.get('cuit_emisor')}.")
    if datos_ia.get("codigo_tipo") is not None and int(datos_ia["codigo_tipo"]) != int(datos_qr.get("tipo", -1)):
        discrepancias.append(f"Tipo comprobante difiere: QR={datos_qr.get('tipo')} vs IA={datos_ia.get('codigo_tipo')}.")
    if datos_ia.get("punto_venta") is not None and int(datos_ia["punto_venta"]) != int(datos_qr.get("pto_vta", -1)):
        discrepancias.append(f"Punto de venta difiere: QR={datos_qr.get('pto_vta')} vs IA={datos_ia.get('punto_venta')}.")
    if datos_ia.get("numero") is not None and int(datos_ia["numero"]) != int(datos_qr.get("nro", -1)):
        discrepancias.append(f"Número difiere: QR={datos_qr.get('nro')} vs IA={datos_ia.get('numero')}.")
    if datos_ia.get("fecha_emision") and str(datos_ia["fecha_emision"])[:10] != str(datos_qr.get("fecha", ""))[:10]:
        discrepancias.append(f"Fecha difiere: QR={datos_qr.get('fecha')} vs IA={datos_ia.get('fecha_emision')}.")
    if datos_ia.get("total") is not None and datos_qr.get("importe") is not None:
        dif = abs(a_decimal(datos_ia["total"]) - a_decimal(datos_qr["importe"]))
        if dif > tolerancia:
            discrepancias.append(f"Total difiere: QR={datos_qr.get('importe')} vs IA={datos_ia.get('total')} (dif {dif}).")
    if datos_ia.get("cae") and datos_qr.get("cae") and solo_digitos(datos_ia["cae"]) != solo_digitos(str(datos_qr.get("cae"))):
        discrepancias.append(f"CAE difiere: QR={datos_qr.get('cae')} vs IA={datos_ia.get('cae')}.")

    return discrepancias
