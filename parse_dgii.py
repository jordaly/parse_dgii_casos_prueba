import csv
import argparse
from pathlib import Path
from datetime import datetime
from typing import Any, TypedDict, get_type_hints


CODIGO_EMPRESA = 1


TYPE_MAPPINGS = {
    "int": int,
    "bigint": int,
    "numeric": float,
    "varchar": str,
    "date": str,
    "datetime": str,
    "bit": bool,
}


class Line(TypedDict):
    name: str
    value: Any | None
    level: int
    origin_col: int


class Encabezado(TypedDict):
    ActividadEconomica: str
    BancoPago: str
    CantidadBulto: float
    CantidadIntentosEnviosDgii: int
    CantidadIntentosEnviosReceptor: int
    CodEmpresa: int
    CodigoEstadoValidacionDgii: int
    CodigoInternoComprador: str
    CodigoModificacion: int
    CodigoSeguridadeCF: str
    CodigoVendedor: str
    CodUsuarioCreador: str
    CodUsuarioEnvioDgii: str
    CodUsuarioEnvioReceptor: str
    CondicionesEntrega: str
    Conductor: str
    ContactoComprador: str
    ContactoEntrega: str
    CorreoComprador: str
    CorreoEmisor: str
    DireccionComprador: str
    DireccionDestino: str
    DireccionEmisor: str
    DireccionEntrega: str
    DocumentoTransporte: int  # BIGINT
    eNCF: str
    EnviaraDgiiPorResumen: bool
    EnviarAReceptor: bool
    ErrorEnvioDgii: str
    EstadoAprobacionComercial: int
    Estatus: bool
    EstatusEnvioDgii: str
    EstatusEnvioReceptor: str
    FechaActualizacionEstadoValidacionDgii: str  # DATETIME
    FechaCreacion: str  # DATETIME
    FechaDesde: str  # DATE
    FechaEmbarque: str  # DATE
    FechaEmision: str  # DATE
    FechaEntrega: str  # DATE
    FechaHasta: str  # DATE
    FechaHoraFirma: str  # DATETIME
    FechaLimitePago: str  # DATE
    FechaNCFModificado: str  # DATE
    FechaOrdenCompra: str  # DATE
    FechaVencimientoSecuencia: str  # DATE
    Ficha: str
    Flete: float
    IdentificadorExtranjero: str
    IndicadorEnvioDiferido: int
    IndicadorMontoGravado: int
    IndicadorNotaCredito: int
    IndicadorServicioTodoIncluido: int
    Informacionadicionalcomprador: str
    InformacionAdicionalEmisor: str
    ITBIS1: int
    ITBIS2: int
    ITBIS3: int
    MensajeErrorEnvioDgii: str
    MontoAvancePago: float
    MontoExento: float
    MontoExentoOtraMoneda: float
    MontoGravado1OtraMoneda: float
    MontoGravado2OtraMoneda: float
    MontoGravado3OtraMoneda: float
    MontoGravadoI1: float
    MontoGravadoI2: float
    MontoGravadoI3: float
    MontoGravadoTotal: float
    MontoGravadoTotalOtraMoneda: float
    MontoImpuestoAdicional: float
    MontoImpuestoAdicionalOtraMoneda: float
    MontoNoFacturable: float
    MontoPeriodo: float
    MontoTotal: float
    MontoTotalOtraMoneda: float
    Municipio: str
    MunicipioComprador: str
    NCFModificado: str
    NombreArchivo: str
    NombreComercial: str
    NombreCompaniaTransportista: str
    NombreDispositivoCreador: str
    NombreDispositivoEnvioDgii: str
    NombreDispositivoEnvioReceptor: str
    NombrePuertoDesembarque: str
    NombrePuertoEmbarque: str
    NombrePuertoSalida: str
    Nota_Interna: str
    NumeroAlbaran: str
    NumeroContenedor: str
    NumeroCuentaPago: str
    NumeroEmbarque: str
    NumeroFacturaInterna: str
    NumeroOrdenCompra: str
    NumeroPedidoInterno: int  # BIGINT
    NumeroReferencia: int  # BIGINT
    NumeroViaje: str
    OtrosGastos: float
    PaisComprador: str
    PaisDestino: str
    PaisOrigen: str
    PesoBruto: float
    PesoNeto: float
    Placa: str
    Provincia: str
    ProvinciaComprador: str
    RazonModificacion: str
    RazonSocialComprador: str
    RazonSocialEmisor: str
    RegimenAduanero: str
    ResponsablePago: str
    RNCComprador: str
    RNCEmisor: str
    RNCIdentificacionCompaniaTransportista: str
    RNCOtroContribuyente: str
    RutaTransporte: str
    RutaVenta: str
    SaldoAnterior: float
    SecuenciaUtilizada: bool
    Seguro: float
    Sucursal: str
    TelefonoAdicional: str
    TerminoPago: str
    TipoCambio: float
    TipoCuentaPago: str
    TipoeCF: str
    TipoIngresos: str
    TipoMoneda: str
    TipoPago: int
    TotalCif: float
    TotalFob: float
    TotalISRPercepcion: float
    TotalISRRetencion: float
    TotalITBIS: float
    TotalITBIS1: float
    TotalITBIS1OtraMoneda: float
    TotalITBIS2: float
    TotalITBIS2OtraMoneda: float
    TotalITBIS3: float
    TotalITBIS3OtraMoneda: float
    TotalITBISOtraMoneda: float
    TotalITBISPercepcion: float
    TotalITBISRetenido: float
    TotalPaginas: int
    TrackId: str
    UltimaFechaEnvioDgii: str  # DATETIME
    UltimaFechaEnvioReceptor: str  # DATETIME
    UnidadBulto: int
    UnidadPesoBruto: int
    UnidadPesoNeto: int
    UnidadVolumen: int
    ValorPagar: float
    Version: float
    ViaTransporte: str
    VolumenBulto: float
    WebSite: str
    ZonaTransporte: str
    ZonaVenta: str


class FormaPago(TypedDict):
    eNCF: str
    CodEmpresa: int
    FormaPago: int
    MontoPago: float


class TelefonoEmisor(TypedDict):
    eNCF: str
    CodEmpresa: int
    TelefonoEmisor: str


class InpuestosAdicionales(TypedDict):
    eNCF: str
    CodEmpresa: int
    TipoImpuesto: str
    TasaImpuestoAdicional: float
    MontoImpuestoSelectivoConsumoEspecifico: float
    MontoImpuestoSelectivoConsumoAdvalorem: float
    OtrosImpuestosAdicionales: float


class InpuestosAdicionalesOtraMoneda(TypedDict):
    eNCF: str
    CodEmpresa: int
    TipoImpuestoOtraMoneda: str
    TasaImpuestoAdicionalOtraMoneda: float
    MontoImpuestoSelectivoConsumoEspecificoOtraMoneda: float
    MontoImpuestoSelectivoConsumoAdvaloremOtraMoneda: float
    OtrosImpuestosAdicionalesOtraMoneda: float


class DetalleEncabezado(TypedDict):
    CantidadItem: float
    CodEmpresa: int
    eNCF: str
    IndicadorBienoServicio: int
    IndicadorFacturacion: int
    MontoItem: float
    NombreItem: str
    NumeroLinea: int
    PrecioUnitarioItem: float
    CantidadReferencia: float
    DescripcionItem: str
    DescuentoMonto: float
    DescuentoOtraMoneda: float
    FechaElaboracion: str  # DATE
    FechaVencimientoItem: str  # DATE
    GradosAlcohol: float
    IndicadorAgenteRetencionoPercepcion: int
    Liquidacion: int
    MontoISRRetenido: float
    MontoITBISRetenido: float
    MontoItemOtraMoneda: float
    PesoNetoKilogramo: float
    PesoNetoMineria: float
    PrecioOtraMoneda: float
    PrecioUnitarioReferencia: float
    RecargoMonto: float
    RecargoOtraMoneda: float
    TipoAfiliacion: int
    UnidadMedida: int
    UnidadReferencia: int


class DetalleItem(TypedDict):
    eNCF: str
    CodEmpresa: int
    NumeroLinea: int
    TipoCodigo: str
    CodigoItem: str


class DetalleSubcantidad(TypedDict):
    eNCF: str
    CodEmpresa: int
    NumeroLinea: int
    Subcantidad: float
    CodigoSubcantidad: int


class DetalleSubdescuento(TypedDict):
    eNCF: str
    CodEmpresa: int
    NumeroLinea: int

    TipoSubDescuento: str
    SubDescuentoPorcentaje: float
    MontoSubDescuento: float


class DetalleSubrecargo(TypedDict):
    eNCF: str
    CodEmpresa: int
    NumeroLinea: int

    TipoSubRecargo: str
    SubRecargoPorcentaje: float
    MontosubRecargo: float


class DetalleInpuestosAdicionales(TypedDict):
    eNCF: str
    CodEmpresa: int
    NumeroLinea: int
    TipoImpuesto: str


class DescuentoRecargo(TypedDict):
    eNCF: str
    CodEmpresa: int
    NumeroLineaDoR: str  # NumeroLineaDoR
    TipoAjuste: str
    IndicadorNorma1007: int
    DescripcionDescuentooRecargo: str
    TipoValor: str
    ValorDescuentooRecargo: float
    MontoDescuentooRecargo: float
    MontoDescuentooRecargoOtraMoneda: float
    IndicadorFacturacionDescuentooRecargo: int


class Paginacion(TypedDict):
    eNCF: str
    CodEmpresa: int
    PaginaNo: int
    NoLineaDesde: str
    NoLineaHasta: str
    SubtotalMontoGravadoPagina: float
    SubtotalMontoGravado1Pagina: float
    SubtotalMontoGravado2Pagina: float
    SubtotalMontoGravado3Pagina: float
    SubtotalExentoPagina: float
    SubtotalItbisPagina: float
    SubtotalItbis1Pagina: float
    SubtotalItbis2Pagina: float
    SubtotalItbis3Pagina: float
    SubtotalImpuestoAdicionalPagina: float
    SubtotalImpuestoAdicionalPaginaTabla: float
    SubtotalImpuestoSelectivoConsumoEspecificoPagina: float
    SubtotalOtrosImpuesto: float
    MontoSubtotalPagina: float
    SubtotalMontoNoFacturablePagina: float


def load_exel(file_path: str):
    import openpyxl
    # Load the workbook and select the active worksheet
    workbook = openpyxl.load_workbook(file_path, data_only=True, read_only=True)
    sheet = workbook.active

    # Read data from the worksheet
    data = []

    keys = []

    # Get the header row
    for cell in sheet[1]:
        keys.append(cell.value)

    for row in sheet.iter_rows(min_row=2, values_only=True):
        row_data = {}
        for i, cell in enumerate(row):
            cell_content: str | int | float | bool | None = None

            if isinstance(cell, bool):
                cell_content = bool(cell)

            if isinstance(cell, str):
                cell_content = cell.strip()

            if isinstance(cell, datetime):
                cell_content = cell.strftime("%Y-%m-%d")

            # if cell_content can be converted to a number, do it
            if isinstance(cell_content, str) and cell_content.isdigit():
                cell_content = float(cell_content)

                if cell_content.is_integer():
                    cell_content = int(cell_content)

            if cell_content is None:
                cell_content = "#e"

            row_data[keys[i]] = cell_content

        data.append(row_data)

    return data


def read_csv_as_dicts(
    file: str, delimiter: str = ",", encoding: str = "latin-1"
) -> list[dict[str, str]]:
    with open(file, newline="", encoding=encoding) as f:
        reader = csv.DictReader(f, delimiter=delimiter)
        return [dict(row) for row in reader]


def save_csv(data: list[dict]) -> str:
    keys = data[0].keys()

    fname = "res.csv"

    with open(fname, "w") as f:
        f.write("|".join(keys) + "\n")

        for row in data:
            f.write("|".join([str(row[key]) for key in keys]) + "\n")

    return fname


def get_encf(lines: list[Line]):
    return str(lines[3]["value"])


def get_numero_linea_det(lines: list[Line]):
    try:
        return int(lines[0]["value"])
    except Exception:
        return None


def group_section(
    lines: list[Line],
    start: int | None = None,
    end: int | None = None,
    ignore_none: bool = True,
    level: int = 1,
) -> dict[int, list[Line]]:

    groups = {}

    if start is not None and end is not None:
        _slice = lines[start:end]
    elif end is None and start is not None:
        _slice = lines[start:]
    elif start is None and end is not None:
        _slice = lines[:end]
    else:
        _slice = lines

    for line in _slice:
        if line["value"] is None and ignore_none:
            continue

        group = int(line["name"].split("[")[level].split("]")[0])

        if group in groups:
            groups[group].append(line)
        else:
            groups[group] = [line]

    return groups


def generar_insert(data: dict[str, Any], table_name: str) -> str:
    columns_present = sorted(
        set([key if key != "NumeroLineaDoR" else "NumeroLinea" for key in data.keys()])
    )

    if "NumeroLineaDoR" in data:
        data["NumeroLinea"] = data["NumeroLineaDoR"]
        del data["NumeroLineaDoR"]

    for key in data.keys():
        if "Fecha" in key:
            try:
                day, month, year = data[key].split("-")
                data[key] = f"{year}-{month}-{day}"
            except Exception:
                pass

    values = [
        f"'{data[name]}'" if isinstance(data[name], str) else str(data[name])
        for name in columns_present
    ]

    query = f"""\

INSERT INTO {table_name}(
    {"\n    ,".join(columns_present)}
)
VALUES (
    {"\n    ,".join(values)}
)

"""
    query = query.replace("'GETDATE()'", "GETDATE()")

    with open("res.sql", "a") as f:
        f.write(query)

    return query


def procesar_encabezado(lines: list[Line]):
    level_0 = []

    for line in lines:
        if line["level"] == 0:
            level_0.append(line)

    # Si tipo ecf = 32 y montototal < 250000 true else false
    enviar_a_dgii_por_resumen = str(
        int(lines[2]["value"]) == 32 and float(lines[139]["value"]) < 250_000
    )

    encabezado: Encabezado = {
        "CodEmpresa": CODIGO_EMPRESA,
        "FechaHoraFirma": "GETDATE()",
        "FechaCreacion": "GETDATE()",
        "Estatus": "True",
        "EnviaraDgiiPorResumen": enviar_a_dgii_por_resumen,
        "EnviarAReceptor": "False",
        "EstatusEnvioDgii": "Pendiente",
        "EstatusEnvioReceptor": "Pendiente",
        "NombreDispositivoCreador": "Interno",
        "CantidadIntentosEnviosDgii": 0,
        "CantidadIntentosEnviosReceptor": 0,
        "NombreArchivo": str(level_0[0]["value"]) + ".xml",
        "CodigoSeguridadeCF": "",
        "CodigoEstadoValidacionDgii": 0,
        "CodUsuarioCreador": "Sist",
    }
    hints = get_type_hints(Encabezado)

    for line in level_0:
        if line["value"] is None:
            continue

        name = line["name"].split("[")[0].strip()

        if name in hints:
            encabezado[name] = hints[name](line["value"])

    generar_insert(encabezado, "Comprobantes_Emitidos")


def procesar_forma_pago(lines: list[Line]):

    encf = get_encf(lines=lines)

    parsed_input: list[FormaPago] = []
    groups = group_section(lines=lines, start=12, end=26)
    hints = get_type_hints(FormaPago)

    for group in groups.keys():
        lines_group = groups[group]
        record = {}

        for line in lines_group:
            name = line["name"].split("[")[0].strip()

            if name in hints:
                record[name] = hints[name](line["value"])

        if record:
            record["eNCF"] = encf
            record["CodEmpresa"] = CODIGO_EMPRESA
            parsed_input.append(record)

    for forma_pago in parsed_input:
        generar_insert(forma_pago, "Comprobantes_Emitidos_Formas_Pago")


def procesar_telefonos_emisor(lines: list[Line]):

    encf = get_encf(lines=lines)
    groups = group_section(lines=lines, start=39, end=42)
    hints = get_type_hints(TelefonoEmisor)
    parsed_input: list[TelefonoEmisor] = []

    for group in groups.keys():
        lines_group = groups[group]
        record = {}
        for line in lines_group:
            name = line["name"].split("[")[0].strip()

            if name in hints:
                record[name] = hints[name](line["value"])

        if record:
            record["eNCF"] = encf
            record["CodEmpresa"] = CODIGO_EMPRESA
            parsed_input.append(record)

    for telefono in parsed_input:
        generar_insert(telefono, "Comprobantes_Emitidos_Telefonos_Emisor")


def procesar_inpuestos_adicionales(lines: list[Line]):

    encf = get_encf(lines=lines)
    groups = group_section(lines=lines, start=119, end=139)

    parsed_input: list[InpuestosAdicionales] = []
    hints = get_type_hints(InpuestosAdicionales)

    for group in groups.keys():
        lines_group = groups[group]
        record = {}

        for line in lines_group:
            name = line["name"].split("[")[0].strip()

            if name in hints:
                record[name] = hints[name](line["value"])

        if record:
            record["eNCF"] = encf
            record["CodEmpresa"] = CODIGO_EMPRESA
            parsed_input.append(record)

    for inpuesto in parsed_input:
        generar_insert(inpuesto, "Comprobantes_Emitidos_Impuestos_Adicionales")


def procesar_inpuestos_adicionales_otra_moneda(lines: list[Line]):
    encf = get_encf(lines=lines)

    groups = group_section(lines=lines, start=161, end=181)

    parsed_input: list[InpuestosAdicionales] = []
    hints = get_type_hints(InpuestosAdicionales)

    for group in groups.keys():
        lines_group = groups[group]

        record = {}

        for line in lines_group:
            name = line["name"].split("[")[0].strip()

            if name in hints:
                record[name] = hints[name](lines["value"])

        if record:
            record["eNCF"] = encf
            record["CodEmpresa"] = CODIGO_EMPRESA
            parsed_input.append(record)

    for inpuesto_otra_moneda in parsed_input:
        generar_insert(
            inpuesto_otra_moneda,
            "Comprobantes_Emitidos_Impuestos_Adicionales_Otra_Moneda",
        )


def procesar_detalle_encabezado(sublines: list[Line], encf: str):
    groups = group_section(lines=sublines)

    if not groups:
        return

    hints = get_type_hints(DetalleEncabezado)
    detalle = {}
    for line in groups[list(groups.keys())[0]]:

        name = line["name"].split("[")[0].strip()

        if name in hints:
            detalle[name] = hints[name](line["value"])

    if detalle:
        detalle["eNCF"] = encf
        detalle["CodEmpresa"] = CODIGO_EMPRESA

        generar_insert(detalle, "Comprobantes_Emitidos_Detalle")


def procesar_detalle_item(sublines: list[Line], encf: str, numero_linea: int):

    groups = group_section(lines=sublines, start=1, end=11, level=2)

    parsed_input: list[DetalleItem] = []
    hints = get_type_hints(DetalleItem)

    for group in groups.keys():
        lines_group = groups[group]
        record = {}

        for line in lines_group:

            name = line["name"].split("[")[0].strip()

            if name in hints:
                record[name] = hints[name](line["value"])

        if record:
            record["eNCF"] = encf
            record["CodEmpresa"] = CODIGO_EMPRESA
            record["NumeroLinea_Detalle"] = numero_linea
            parsed_input.append(record)

    for item in parsed_input:
        generar_insert(item, "Comprobantes_Emitidos_Detalle_Item")


def procesar_detalle_subcantidad(sublines: list[Line], encf: str, numero_linea: int):

    groups = group_section(lines=sublines, start=22, end=32, level=2)

    parsed_input: list[DetalleSubcantidad] = []
    hints = get_type_hints(DetalleSubcantidad)

    for group in groups.keys():
        lines_group = groups[group]

        record = {}

        for line in lines_group:

            name = line["name"].split("[")[0].strip()

            if name in hints:
                record[name] = hints[name](line["value"])

        if record:
            record["eNCF"] = encf
            record["CodEmpresa"] = CODIGO_EMPRESA
            record["NumeroLinea_Detalle"] = numero_linea
            parsed_input.append(record)

    for subcantidad in parsed_input:
        generar_insert(subcantidad, "Comprobantes_Emitidos_Detalle_Subcantidad")


def procesar_detalle_subdescuento(sublines: list[Line], encf: str, numero_linea: int):

    groups = group_section(lines=sublines, start=42, end=57, level=2)

    parsed_input: list[DetalleSubdescuento] = []
    hints = get_type_hints(DetalleSubdescuento)
    for group in groups.keys():
        lines_group = groups[group]
        record = {}

        for line in lines_group:

            name = line["name"].split("[")[0].strip()

            if name in hints:
                record[name] = hints[name](line["value"])

        if record:
            record["eNCF"] = encf
            record["CodEmpresa"] = CODIGO_EMPRESA
            record["NumeroLinea_Detalle"] = numero_linea
            parsed_input.append(record)

    for subdescuento in parsed_input:
        generar_insert(subdescuento, "Comprobantes_Emitidos_Detalle_SubDescuento")


def procesar_detalle_subrecargo(sublines: list[Line], encf: str, numero_linea: int):

    groups = group_section(lines=sublines, start=58, end=73, level=2)

    parsed_input: list[DetalleSubrecargo] = []
    hints = get_type_hints(DetalleSubrecargo)
    for group in groups.keys():
        lines_group = groups[group]

        record = {}

        for line in lines_group:

            name = line["name"].split("[")[0].strip()

            if name in hints:
                record[name] = hints[name](line["value"])

        if record:
            record["eNCF"] = encf
            record["CodEmpresa"] = CODIGO_EMPRESA
            record["NumeroLinea_Detalle"] = numero_linea
            parsed_input.append(record)

    for subrecargo in parsed_input:
        generar_insert(subrecargo, "Comprobantes_Emitidos_Detalle_SubRecargo")


def procesar_detalle_impuestos_adicionales(
    sublines: list[Line], encf: str, numero_linea: int
):

    groups = group_section(lines=sublines, start=73, end=75, level=2)

    parsed_input: list[DetalleInpuestosAdicionales] = []
    hints = get_type_hints(DetalleInpuestosAdicionales)
    for group in groups.keys():
        lines_group = groups[group]
        record = {}

        for line in lines_group:

            name = line["name"].split("[")[0].strip()

            if name in hints:
                record[name] = hints[name](line["value"])

        if record:
            record["eNCF"] = encf
            record["CodEmpresa"] = CODIGO_EMPRESA
            record["NumeroLinea_Detalle"] = numero_linea
            parsed_input.append(record)

    for impuesto in parsed_input:
        generar_insert(impuesto, "Comprobantes_Emitidos_Detalle_Impuestos_Adicionales")


def procesar_detalle(lines: list[Line]):
    encf = get_encf(lines=lines)

    groups = group_section(lines=lines, start=182, end=5142, ignore_none=False)

    for group in groups.keys():
        lines_group = groups[group]

        numero_linea = get_numero_linea_det(lines=lines_group)

        if numero_linea is not None:
            procesar_detalle_encabezado(sublines=lines_group, encf=encf)
            procesar_detalle_item(
                sublines=lines_group, encf=encf, numero_linea=numero_linea
            )
            procesar_detalle_subcantidad(
                sublines=lines_group, encf=encf, numero_linea=numero_linea
            )
            procesar_detalle_subdescuento(
                sublines=lines_group, encf=encf, numero_linea=numero_linea
            )
            procesar_detalle_subrecargo(
                sublines=lines_group, encf=encf, numero_linea=numero_linea
            )
            procesar_detalle_impuestos_adicionales(
                sublines=lines_group, encf=encf, numero_linea=numero_linea
            )


def procesar_descuento_recargo(lines: list[Line]):
    encf = get_encf(lines=lines)

    groups = group_section(lines=lines, start=5157, end=5175)

    parsed_input: list[DescuentoRecargo] = []
    hints = get_type_hints(DescuentoRecargo)
    for group in groups.keys():
        lines_group = groups[group]

        record = {}

        for line in lines_group:

            name = line["name"].split("[")[0].strip()
            if name in hints:
                record[name] = hints[name](line["value"])

        if record:
            record["eNCF"] = encf
            record["CodEmpresa"] = CODIGO_EMPRESA

            parsed_input.append(record)

    for descuento_o_recargo in parsed_input:
        generar_insert(
            descuento_o_recargo,
            "Comprobantes_Emitidos_Descuento_Recargo",
        )


def procesar_paginacion(lines: list[Line]):
    encf = get_encf(lines=lines)

    groups = group_section(lines=lines, start=5175, end=5210)

    parsed_input: list[Paginacion] = []

    hints = get_type_hints(Paginacion)

    for group in groups.keys():
        lines_group = groups[group]

        record = {}
        for line in lines_group:

            name = line["name"].split("[")[0].strip()

            if name in hints:
                record[name] = hints(name)(line["value"])

        if record:
            record["eNCF"] = encf
            record["CodEmpresa"] = CODIGO_EMPRESA

            parsed_input.append(record)

    for pagina in parsed_input:
        generar_insert(pagina, "Comprobantes_Emitidos_Paginacion")


def main():
    global CODIGO_EMPRESA

    parser = argparse.ArgumentParser(
        prog="Herramienta de transformacion DGII",
        description="Transforma archivos csv en sentensias insert en sql de acuerdo a la documentacion de la DGII.",
    )

    parser.add_argument("filepath", type=str)

    parser.add_argument("-e", "--empresa", type=int, default=1)

    args = parser.parse_args()

    path_input = args.filepath

    CODIGO_EMPRESA = args.empresa

    if path_input.endswith(".xlsx"):
        print("Loading data from excel file...")
        data = load_exel(path_input)

        print("Saving parsed excel file into a csv file\n")
        path_input = save_csv(data)

    filepath = Path(path_input)

    if not filepath.exists():
        print("ERROR: El archivo no existe.")
        return

    if not path_input.endswith(".csv"):
        print("ERROR: El archivo no es un archivo .csv")
        return

    data = read_csv_as_dicts(file=filepath, delimiter="|")

    lines = []

    keys = [key for key in data[0].keys()]

    with open("res.sql", "w") as f:
        f.write(
            """\
DECLARE @errormensage varchar(max)

BEGIN TRY
BEGIN TRANSACTION

"""
        )

    for index, row in enumerate(data):
        lines.clear()

        for i, key in enumerate(keys):
            number_of_levels = len(key.split("[")) - 1

            value = row[key]

            if value.strip() == "#e":
                value = None

            lines.append(
                Line(name=key, value=value, level=number_of_levels, origin_col=i)
            )

        eNCF = get_encf(lines)

        print(f"Generating ({index + 1}) {eNCF=}")

        with open("res.sql", "a") as f:
            f.write("-- " + "=" * 80)
            f.write(f" -- Transaccon {index + 1}, {eNCF=} --\n")

        procesar_encabezado(lines)
        procesar_forma_pago(lines)
        procesar_telefonos_emisor(lines)
        procesar_inpuestos_adicionales(lines)
        procesar_inpuestos_adicionales_otra_moneda(lines)
        procesar_detalle(lines)
        procesar_descuento_recargo(lines)
        procesar_paginacion(lines)


    with open("res.sql", "a") as f:
        f.write(
            """\


COMMIT TRANSACTION
END TRY
BEGIN CATCH
    ROLLBACK TRANSACTION;
    set @errormensage = ERROR_MESSAGE()
    RAISERROR(@errormensage,15,217)
END CATCH"""
        )

    print("Done!\n")

    print(f"Result saved at: {Path("res.sql").resolve()}")


if __name__ == "__main__":
    main()

