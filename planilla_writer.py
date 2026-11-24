from __future__ import annotations

import re
import zipfile
from datetime import date
from decimal import Decimal
from io import BytesIO
from typing import Any, Optional

from lxml import etree

from extractor_facturas import FacturaExtraida, es_nota_credito, solo_digitos, to_dec

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
N = {"m": NS}

COLUMNAS_ENTRADA = {"E", "F", "I", "J", "K", "L", "M", "N", "P", "T", "U", "V", "W"}
COLUMNAS_VACIAS_REQUERIDAS = {"E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "T", "U", "V", "W"}


class ErrorPlanilla(Exception):
    pass


def _col(ref: str) -> str:
    m = re.match(r"[A-Z]+", ref or "")
    return m.group(0) if m else ""


def _col_idx(col: str) -> int:
    n = 0
    for ch in col:
        n = n * 26 + ord(ch) - 64
    return n


class PlanillaMemoria:
    """Edita el XML interno de la planilla de 3CICON en memoria sin tocar fórmulas, estilos ni logo."""

    def __init__(self, excel_bytes: bytes, hoja: Optional[str] = None, texto_fila_total: str = "TOTAL"):
        self.hoja_objetivo = hoja
        self.texto_fila_total = texto_fila_total.upper()

        try:
            with zipfile.ZipFile(BytesIO(excel_bytes), "r") as z:
                self._infos = z.infolist()
                self._partes = {i.filename: z.read(i.filename) for i in self._infos}
        except zipfile.BadZipFile as e:
            raise ErrorPlanilla("El archivo subido no es un Excel .xlsx válido.") from e

        self._originales = dict(self._partes)
        self._ubicar_hoja()
        self._indexar()
        self._validar_firma()
        self.filas_escritas: list[int] = []

    def _ubicar_hoja(self) -> None:
        wb = etree.fromstring(self._partes["xl/workbook.xml"])
        rels = etree.fromstring(self._partes["xl/_rels/workbook.xml.rels"])
        hojas = wb.find("m:sheets", N)
        hoja = next((h for h in hojas if self.hoja_objetivo in (None, h.get("name"))), None)
        if hoja is None:
            raise ErrorPlanilla(f"No existe la hoja '{self.hoja_objetivo}' en el Excel.")

        rid = hoja.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
        destino = next(r.get("Target") for r in rels if r.get("Id") == rid)
        self.parte_hoja = destino.lstrip("/") if destino.startswith("/") else "xl/" + destino
        self.xml = etree.fromstring(self._partes[self.parte_hoja])
        self.sheetdata = self.xml.find("m:sheetData", N)
        self.filas = {int(r.get("r")): r for r in self.sheetdata.findall("m:row", N)}

        self.compartidas: list[str] = []
        if "xl/sharedStrings.xml" in self._partes:
            ss = etree.fromstring(self._partes["xl/sharedStrings.xml"])
            self.compartidas = ["".join(si.itertext()) for si in ss.findall("m:si", N)]

    def _celda(self, fila: int, col: str, crear: bool = False):
        r = self.filas.get(fila)
        if r is None:
            return None
        ref = f"{col}{fila}"
        for c in r.findall("m:c", N):
            if c.get("r") == ref:
                return c
        if not crear:
            return None

        nueva = etree.Element(f"{{{NS}}}c", r=ref)
        pos = 0
        for i, c in enumerate(r.findall("m:c", N)):
            if _col_idx(_col(c.get("r"))) < _col_idx(col):
                pos = i + 1
        r.insert(pos, nueva)
        return nueva

    def _valor(self, fila: int, col: str):
        c = self._celda(fila, col)
        if c is None:
            return None
        t, v = c.get("t"), c.find("m:v", N)
        if t == "inlineStr":
            is_el = c.find("m:is", N)
            return "".join(is_el.itertext()) if is_el is not None else ""
        if v is None or v.text is None:
            return None
        if t == "s":
            return self.compartidas[int(v.text)]
        if t in ("str", "e", "b"):
            return v.text
        try:
            return float(v.text)
        except ValueError:
            return v.text

    def _validar_firma(self) -> None:
        esperado = {
            ("E", "CONTROL DE FACTURAS Y GASTOS"),
            ("F", "PROVEEDOR"),
            ("N", "IMPORTE TOTAL"),
            ("F", "FACTURA COMPRAS"),
        }
        presentes = {
            (c, str(self._valor(r, c)).strip().upper())
            for r in range(1, 12)
            for c in "EFN"
            if isinstance(self._valor(r, c), str)
        }
        faltan = sorted(t for _, t in esperado - presentes)
        if faltan:
            raise ErrorPlanilla(
                "Este Excel no parece ser la planilla oficial de IVA Compras (faltan los rótulos: "
                + ", ".join(faltan)
                + ")."
            )

    def _indexar(self) -> None:
        self.fila_total = next(
            (
                r
                for r in sorted(self.filas)
                if str(self._valor(r, "M") or "").strip().upper().startswith(self.texto_fila_total)
            ),
            None,
        )
        if self.fila_total is None:
            raise ErrorPlanilla("No se encontró la fila 'TOTAL COMPRAS' en la columna M.")

        rotulo = next(
            (
                r
                for r in sorted(self.filas)
                if str(self._valor(r, "F") or "").strip().upper() == "FACTURA COMPRAS"
            ),
            None,
        )
        self.fila_inicio = (rotulo + 1) if rotulo else 9

        def es_dato(r: int) -> bool:
            return any(isinstance(self._valor(r, c), float) for c in "EJMN") or bool(self._valor(r, "F"))

        datos = [r for r in range(self.fila_inicio, self.fila_total) if r in self.filas and es_dato(r)]
        self.ultima_fila_datos = max(datos) if datos else self.fila_inicio - 1
        self.cantidad_cargadas = len(datos)

        self.claves: dict[tuple[str, int, int], int] = {}
        self.nombres: dict[str, str] = {}
        for r in datos:
            try:
                i, j, k = (int(self._valor(r, c)) for c in "IJK")
                pv, nro = int(self._valor(r, "L")), int(self._valor(r, "M"))
            except (TypeError, ValueError):
                continue
            cuit = f"{i:02d}{j:08d}{k}"
            self.claves[(cuit, pv, nro)] = r
            if isinstance(self._valor(r, "F"), str):
                self.nombres.setdefault(cuit, self._valor(r, "F"))

        def vacia(r: int) -> bool:
            return r in self.filas and all(self._valor(r, c) in (None, "") for c in COLUMNAS_VACIAS_REQUERIDAS)

        self.filas_libres = [r for r in range(self.ultima_fila_datos + 1, self.fila_total) if vacia(r)]

    def resumen_estado(self) -> dict[str, Any]:
        return {
            "cantidad_cargadas": self.cantidad_cargadas,
            "ultima_fila_datos": self.ultima_fila_datos,
            "proxima_fila": self.filas_libres[0] if self.filas_libres else None,
            "filas_libres": len(self.filas_libres),
            "fila_total": self.fila_total,
        }

    @staticmethod
    def _poner(c, valor) -> None:
        if c.find("m:f", N) is not None:
            raise ErrorPlanilla(f"La celda {c.get('r')} contiene una fórmula; no se modifica.")
        for hijo in list(c):
            c.remove(hijo)
        if "t" in c.attrib:
            del c.attrib["t"]
        if valor is None:
            return
        if isinstance(valor, str):
            c.set("t", "inlineStr")
            is_ = etree.SubElement(c, f"{{{NS}}}is")
            t = etree.SubElement(is_, f"{{{NS}}}t")
            t.text = valor
            if valor != valor.strip():
                t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        else:
            val_str = format(valor.normalize(), "f") if isinstance(valor, Decimal) else str(valor)
            etree.SubElement(c, f"{{{NS}}}v").text = val_str

    def agregar_factura(self, f: FacturaExtraida, nc_en_negativo: bool = True) -> tuple[str, Optional[int], str]:
        cuit = solo_digitos(f.cuit_emisor)
        if len(cuit) != 11 or f.fecha_emision is None or f.punto_venta is None or f.numero is None:
            return "ERROR", None, "Faltan datos obligatorios (fecha, CUIT de 11 dígitos, punto de venta o número)."

        clave = (cuit, int(f.punto_venta), int(f.numero))
        if clave in self.claves:
            return "DUPLICADA", self.claves[clave], f"Ya existe en la fila {self.claves[clave]}."

        if not self.filas_libres:
            return "SIN_ESPACIO", None, "No quedan filas libres antes de TOTAL COMPRAS."

        fila = self.filas_libres.pop(0)
        signo = Decimal(-1) if (nc_en_negativo and es_nota_credito(f.codigo_tipo)) else Decimal(1)

        def q(x: Decimal) -> Decimal:
            return (to_dec(x) * signo).quantize(Decimal("0.01"))

        nombre = self.nombres.get(cuit) or (f.razon_social_emisor or "SIN PROVEEDOR").strip().upper()
        pv = {
            "E": (f.fecha_emision - date(1899, 12, 30)).days,
            "F": nombre,
            "I": int(cuit[:2]),
            "J": int(cuit[2:10]),
            "K": int(cuit[10]),
            "L": int(f.punto_venta),
            "M": int(f.numero),
            "N": q(f.total),
            "P": f.alicuota_iva,  # None limpia el 21% por defecto de la plantilla (para Factura B/C/Monotributo)
            "T": q(f.imp_internos_no_gravado) or None,
            "U": q(f.perc_iibb_caba) or None,
            "V": q(f.perc_iibb_pba) or None,
            "W": q(f.perc_iva) or None,
        }

        for col, valor in pv.items():
            if col not in COLUMNAS_ENTRADA:
                continue
            if valor is None and col != "P":
                continue
            self._poner(self._celda(fila, col, crear=True), valor)

        self.claves[clave] = fila
        self.nombres.setdefault(cuit, nombre)
        self.filas_escritas.append(fila)
        return "OK", fila, f"Escrita en fila {fila}."

    def exportar_bytes(self) -> bytes:
        nuevas = dict(self._originales)
        nuevas[self.parte_hoja] = etree.tostring(self.xml, xml_declaration=True, encoding="UTF-8", standalone=True)

        wb = self._originales["xl/workbook.xml"].decode("utf-8")
        if "fullCalcOnLoad" not in wb:
            wb = re.sub(r"<calcPr([^>]*?)/>", r'<calcPr\1 fullCalcOnLoad="1"/>', wb, count=1)
        nuevas["xl/workbook.xml"] = wb.encode("utf-8")

        buf = BytesIO()
        with zipfile.ZipFile(buf, "w") as z:
            for info in self._infos:
                z.writestr(info, nuevas[info.filename], compress_type=info.compress_type)
        return buf.getvalue()
