from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from typing import Literal, Optional

from fiscal_config import (
    ParametrosFiscales,
    RegimenRetencion,
    TIPOS_COMPROBANTE_ARCA,
    ZONA_ARG,
    periodo_fiscal_mensual,
    redondear_moneda,
)

CERO = Decimal("0.00")


def solo_digitos(txt: str) -> str:
    return re.sub(r"\D", "", txt or "")


def a_decimal(valor: Decimal | int | str | None) -> Decimal:
    if isinstance(valor, float):
        raise TypeError("No se permite float para valores monetarios; usá Decimal o str.")
    if valor is None or valor == "":
        return CERO
    if isinstance(valor, Decimal):
        if valor.is_nan() or valor.is_infinite():
            raise ValueError("Valor Decimal inválido (NaN/Infinity).")
        return redondear_moneda(valor)
    if isinstance(valor, int):
        return redondear_moneda(Decimal(valor))
    limpio = str(valor).replace("$", "").strip()
    try:
        return redondear_moneda(Decimal(limpio))
    except InvalidOperation as e:
        raise ValueError(f"Monto inválido: {valor!r}") from e


def cuit_valido(cuit: str) -> bool:
    c = solo_digitos(cuit)
    if len(c) != 11:
        return False
    coef = (5, 4, 3, 2, 7, 6, 5, 4, 3, 2)
    resto = 11 - (sum(int(a) * b for a, b in zip(c[:10], coef)) % 11)
    dv = 0 if resto == 11 else 9 if resto == 10 else resto
    return dv == int(c[10])


def cae_valido(cae: str) -> bool:
    c = solo_digitos(cae)
    return len(c) == 14


@dataclass(frozen=True)
class LineaAlicuotaIVA:
    alicuota: Decimal
    neto_gravado: Decimal
    importe_iva: Decimal

    def __post_init__(self):
        object.__setattr__(self, "alicuota", a_decimal(self.alicuota))
        object.__setattr__(self, "neto_gravado", a_decimal(self.neto_gravado))
        object.__setattr__(self, "importe_iva", a_decimal(self.importe_iva))


@dataclass(frozen=True)
class ComprobanteAsociado:
    codigo_tipo: int
    punto_venta: int
    numero: int


@dataclass
class ComprobanteFiscal:
    fecha_emision: Optional[date]
    codigo_tipo: Optional[int]
    punto_venta: Optional[int]
    numero: Optional[int]
    cuit_emisor: str
    razon_social_emisor: str
    cuit_cliente: str = ""
    lineas_iva: list[LineaAlicuotaIVA] = field(default_factory=list)
    no_gravado: Decimal = CERO
    exento: Decimal = CERO
    percepcion_iva: Decimal = CERO
    percepciones_iibb: dict[str, Decimal] = field(default_factory=dict)
    otras_percepciones: Decimal = CERO
    impuestos_internos: Decimal = CERO
    otros_tributos: Decimal = CERO
    descuentos: Decimal = CERO
    total_declarado: Decimal = CERO
    cae: str = ""
    vto_cae: Optional[date] = None
    moneda: str = "PES"
    comprobante_asociado: Optional[ComprobanteAsociado] = None
    alicuotas_padron_proveedor: dict[str, Decimal] = field(default_factory=dict)

    def __post_init__(self):
        self.cuit_emisor = solo_digitos(self.cuit_emisor)
        self.cuit_cliente = solo_digitos(self.cuit_cliente)
        self.cae = solo_digitos(self.cae)
        self.no_gravado = a_decimal(self.no_gravado)
        self.exento = a_decimal(self.exento)
        self.percepcion_iva = a_decimal(self.percepcion_iva)
        self.percepciones_iibb = {k.upper().strip(): a_decimal(v) for k, v in (self.percepciones_iibb or {}).items()}
        self.otras_percepciones = a_decimal(self.otras_percepciones)
        self.impuestos_internos = a_decimal(self.impuestos_internos)
        self.otros_tributos = a_decimal(self.otros_tributos)
        self.descuentos = a_decimal(self.descuentos)
        self.total_declarado = a_decimal(self.total_declarado)
        self.alicuotas_padron_proveedor = {
            k.upper().strip(): a_decimal(v) for k, v in (self.alicuotas_padron_proveedor or {}).items()
        }


@dataclass(frozen=True)
class RetencionCalculada:
    regimen: str
    impuesto: str
    jurisdiccion: str
    base_imponible: Decimal
    alicuota_aplicada: Decimal
    importe: Decimal


@dataclass
class ResumenCalculoFiscal:
    periodo_fiscal: str
    signo_contable: int
    total_neto_gravado: Decimal
    total_no_gravado: Decimal
    total_exento: Decimal
    total_iva_discriminado: Decimal
    iva_por_alicuota: dict[str, Decimal]
    neto_por_alicuota: dict[str, Decimal]
    credito_fiscal_iva: Decimal
    total_percepciones_iibb: Decimal
    total_percepciones: Decimal
    total_otros_tributos: Decimal
    total_descuentos: Decimal
    total_calculado: Decimal
    total_declarado: Decimal
    diferencia_reconciliacion: Decimal
    retenciones: list[RetencionCalculada]
    total_retenciones: Decimal
    importe_a_pagar: Decimal
    impacto_iva_compras: Decimal


@dataclass
class ResultadoValidacionFiscal:
    estado: Literal["OK", "ADVERTENCIA", "INCONSISTENCIA", "ERROR"]
    errores: list[str]
    advertencias: list[str]
    inconsistencias: list[str]
    resumen: ResumenCalculoFiscal


def _alicuota_vigente(tasa: Decimal, fecha: date, params: ParametrosFiscales) -> bool:
    for cfg in params.alicuotas_iva:
        if cfg.tasa == tasa and cfg.esta_vigente(fecha):
            return True
    return False


def _calcular_retenciones(
    comp: ComprobanteFiscal,
    total_neto_gravado: Decimal,
    total_iva: Decimal,
    total_comprobante: Decimal,
    params: ParametrosFiscales,
) -> list[RetencionCalculada]:
    if not params.empresa_es_agente_retencion or not comp.fecha_emision:
        return []
    tipo = TIPOS_COMPROBANTE_ARCA.get(comp.codigo_tipo or -1)
    if tipo is None or tipo.signo <= 0:
        return []

    resultados: list[RetencionCalculada] = []
    for reg in params.regimenes_retencion:
        if not reg.esta_vigente(comp.fecha_emision):
            continue
        if reg.base_calculo == "NETO_GRAVADO":
            base = total_neto_gravado
        elif reg.base_calculo == "IVA":
            base = total_iva
        else:
            base = total_comprobante

        base_sujeta = base - reg.minimo_no_sujeto
        if base_sujeta <= CERO:
            continue

        clave_padron = f"{reg.impuesto}_{reg.jurisdiccion}".upper()
        alicuota = (
            comp.alicuotas_padron_proveedor.get(clave_padron, reg.alicuota)
            if reg.usa_padron
            else reg.alicuota
        )
        importe = redondear_moneda(base_sujeta * alicuota / Decimal("100"))
        if importe > CERO:
            resultados.append(
                RetencionCalculada(
                    regimen=reg.nombre,
                    impuesto=reg.impuesto,
                    jurisdiccion=reg.jurisdiccion,
                    base_imponible=redondear_moneda(base_sujeta),
                    alicuota_aplicada=alicuota,
                    importe=importe,
                )
            )
    return resultados


def validar_y_calcular_comprobante(
    comp: ComprobanteFiscal,
    params: Optional[ParametrosFiscales] = None,
    hoy: Optional[date] = None,
) -> ResultadoValidacionFiscal:
    params = params or ParametrosFiscales()
    hoy_ref = hoy or datetime.now(ZONA_ARG).date()

    errores: list[str] = []
    advertencias: list[str] = []
    inconsistencias: list[str] = []

    if comp.fecha_emision is None:
        errores.append("Falta la fecha de emisión.")
        periodo = ""
    else:
        periodo = periodo_fiscal_mensual(comp.fecha_emision)
        if comp.fecha_emision < date(2015, 1, 1) or comp.fecha_emision > date(hoy_ref.year + 1, 12, 31):
            errores.append(f"Fecha de emisión fuera de rango válido: {comp.fecha_emision.isoformat()}.")

    tipo = TIPOS_COMPROBANTE_ARCA.get(comp.codigo_tipo) if comp.codigo_tipo is not None else None
    if tipo is None:
        errores.append(f"Código de tipo de comprobante ARCA inválido o faltante: {comp.codigo_tipo}.")
        signo = 1
    else:
        signo = tipo.signo
        if not tipo.es_comprobante_fiscal:
            errores.append(f"El comprobante '{tipo.descripcion}' (código {tipo.codigo}) no es un comprobante fiscal de IVA Compras.")

    if comp.punto_venta is None or not (1 <= comp.punto_venta <= 99999):
        errores.append(f"Punto de venta inválido (debe estar entre 1 y 99999): {comp.punto_venta}.")
    if comp.numero is None or not (1 <= comp.numero <= 99999999):
        errores.append(f"Número de comprobante inválido (debe estar entre 1 y 99999999): {comp.numero}.")

    if not comp.cuit_emisor:
        errores.append("Falta el CUIT del emisor.")
    elif not cuit_valido(comp.cuit_emisor):
        errores.append(f"CUIT del emisor inválido (dígito verificador incorrecto): {comp.cuit_emisor}.")

    if not comp.razon_social_emisor.strip():
        errores.append("Falta la razón social del emisor.")

    cuit_emp = solo_digitos(params.cuit_empresa)
    if comp.cuit_cliente:
        if not cuit_valido(comp.cuit_cliente):
            errores.append(f"CUIT del cliente inválido: {comp.cuit_cliente}.")
        elif cuit_emp and comp.cuit_cliente != cuit_emp:
            errores.append(
                f"El comprobante está emitido a otro cliente (CUIT {comp.cuit_cliente}) y no a {params.razon_social_empresa} ({cuit_emp})."
            )

    if comp.cae:
        if not cae_valido(comp.cae):
            errores.append(f"CAE inválido (debe tener exactamente 14 dígitos): {comp.cae}.")
    else:
        advertencias.append("No se informó CAE.")

    if comp.vto_cae is not None and comp.fecha_emision is not None:
        if comp.vto_cae < comp.fecha_emision:
            errores.append(
                f"Fecha de vencimiento del CAE ({comp.vto_cae.isoformat()}) anterior a la fecha de emisión ({comp.fecha_emision.isoformat()})."
            )

    moneda_norm = (comp.moneda or "PES").upper().strip()
    if moneda_norm != params.moneda_operativa:
        inconsistencias.append(
            f"Comprobante en moneda '{moneda_norm}' (la empresa opera solo en {params.moneda_operativa}): requiere revisión manual."
        )

    campos_no_negativos = [
        ("no_gravado", comp.no_gravado),
        ("exento", comp.exento),
        ("percepcion_iva", comp.percepcion_iva),
        ("otras_percepciones", comp.otras_percepciones),
        ("impuestos_internos", comp.impuestos_internos),
        ("otros_tributos", comp.otros_tributos),
        ("descuentos", comp.descuentos),
        ("total_declarado", comp.total_declarado),
    ]
    for nombre_campo, val in campos_no_negativos:
        if val < CERO:
            errores.append(f"El campo '{nombre_campo}' no puede ser negativo ({val}); el signo lo determina el tipo de comprobante.")

    for jur, val in comp.percepciones_iibb.items():
        if val < CERO:
            errores.append(f"La percepción IIBB ({jur}) no puede ser negativa ({val}).")

    neto_por_alicuota: dict[str, Decimal] = {}
    iva_por_alicuota: dict[str, Decimal] = {}
    total_neto_gravado = CERO
    total_iva_discriminado = CERO

    for linea in comp.lineas_iva:
        if linea.neto_gravado < CERO or linea.importe_iva < CERO:
            errores.append(f"Línea de IVA {linea.alicuota}% con importe negativo (neto={linea.neto_gravado}, iva={linea.importe_iva}).")
            continue

        clave_al = f"{linea.alicuota:.2f}"
        neto_por_alicuota[clave_al] = redondear_moneda(neto_por_alicuota.get(clave_al, CERO) + linea.neto_gravado)
        iva_por_alicuota[clave_al] = redondear_moneda(iva_por_alicuota.get(clave_al, CERO) + linea.importe_iva)
        total_neto_gravado = redondear_moneda(total_neto_gravado + linea.neto_gravado)
        total_iva_discriminado = redondear_moneda(total_iva_discriminado + linea.importe_iva)

        if comp.fecha_emision and not _alicuota_vigente(linea.alicuota, comp.fecha_emision, params):
            inconsistencias.append(
                f"Alícuota de IVA {linea.alicuota}% no figura como vigente para la fecha {comp.fecha_emision.isoformat()}."
            )

        if tipo is None or tipo.discrimina_iva:
            iva_esperado = redondear_moneda(linea.neto_gravado * linea.alicuota / Decimal("100"))
            dif_iva = redondear_moneda(linea.importe_iva - iva_esperado)
            if abs(dif_iva) > params.tolerancia_iva_alicuota:
                inconsistencias.append(
                    f"IVA {linea.alicuota}% inconsistente: declarado {linea.importe_iva}, esperado {iva_esperado} (diferencia {dif_iva})."
                )

    credito_fiscal_iva = CERO
    if tipo is not None:
        if not tipo.discrimina_iva:
            if total_iva_discriminado > CERO:
                inconsistencias.append(
                    f"{tipo.descripcion} no discrimina IVA, pero tiene IVA declarado por {total_iva_discriminado}."
                )
            credito_fiscal_iva = CERO
        elif tipo.es_factura_m:
            if params.politica_factura_m == "COMPUTAR_CREDITO":
                credito_fiscal_iva = total_iva_discriminado
                advertencias.append("Factura M: crédito fiscal computado según parámetro configurado (validar retenciones con contador).")
            elif params.politica_factura_m == "NO_COMPUTAR":
                credito_fiscal_iva = CERO
                advertencias.append("Factura M: crédito fiscal no computado según parámetro configurado.")
            else:
                credito_fiscal_iva = CERO
                advertencias.append("Factura M detectada: requiere validación del contador (cómputo de IVA y retención RG 1575).")
        else:
            credito_fiscal_iva = total_iva_discriminado

        if tipo.signo == -1 and comp.comprobante_asociado is None:
            advertencias.append("Nota de Crédito sin comprobante asociado informado.")

    total_percepciones_iibb = redondear_moneda(sum(comp.percepciones_iibb.values(), CERO))
    total_percepciones = redondear_moneda(comp.percepcion_iva + total_percepciones_iibb + comp.otras_percepciones)
    total_otros_tributos = redondear_moneda(comp.impuestos_internos + comp.otros_tributos)

    total_calculado = redondear_moneda(
        total_neto_gravado
        + comp.no_gravado
        + comp.exento
        + total_iva_discriminado
        + total_percepciones
        + total_otros_tributos
        - comp.descuentos
    )
    diferencia = redondear_moneda(comp.total_declarado - total_calculado)

    if abs(diferencia) > params.tolerancia_reconciliacion:
        inconsistencias.append(
            f"El comprobante no cierra: total declarado {comp.total_declarado} vs calculado {total_calculado} (diferencia {diferencia})."
        )

    retenciones = _calcular_retenciones(comp, total_neto_gravado, total_iva_discriminado, comp.total_declarado, params)
    total_retenciones = redondear_moneda(sum((r.importe for r in retenciones), CERO))
    importe_a_pagar = redondear_moneda(comp.total_declarado - total_retenciones) if signo > 0 else CERO
    impacto_iva_compras = redondear_moneda(comp.total_declarado * Decimal(signo))

    if errores:
        estado: Literal["OK", "ADVERTENCIA", "INCONSISTENCIA", "ERROR"] = "ERROR"
    elif inconsistencias:
        estado = "INCONSISTENCIA"
    elif advertencias:
        estado = "ADVERTENCIA"
    else:
        estado = "OK"

    resumen = ResumenCalculoFiscal(
        periodo_fiscal=periodo,
        signo_contable=signo,
        total_neto_gravado=total_neto_gravado,
        total_no_gravado=comp.no_gravado,
        total_exento=comp.exento,
        total_iva_discriminado=total_iva_discriminado,
        iva_por_alicuota=iva_por_alicuota,
        neto_por_alicuota=neto_por_alicuota,
        credito_fiscal_iva=credito_fiscal_iva,
        total_percepciones_iibb=total_percepciones_iibb,
        total_percepciones=total_percepciones,
        total_otros_tributos=total_otros_tributos,
        total_descuentos=comp.descuentos,
        total_calculado=total_calculado,
        total_declarado=comp.total_declarado,
        diferencia_reconciliacion=diferencia,
        retenciones=retenciones,
        total_retenciones=total_retenciones,
        importe_a_pagar=importe_a_pagar,
        impacto_iva_compras=impacto_iva_compras,
    )

    return ResultadoValidacionFiscal(
        estado=estado,
        errores=errores,
        advertencias=advertencias,
        inconsistencias=inconsistencias,
        resumen=resumen,
    )
