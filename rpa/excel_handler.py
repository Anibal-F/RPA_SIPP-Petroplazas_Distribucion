import openpyxl
from pathlib import Path
from typing import List, Tuple

FOLIO_COL        = 4    # Column D  — folio de factura
CC_COL           = 32   # Column AF — CC OC
OBS_COL          = 33   # Column AG — Observaciones OC
SUBTOTAL_COL     = 34   # Column AH — Subtotal OC
DESCUENTO_COL    = 35   # Column AI — Descuento OC
IVA_COL          = 36   # Column AJ — IVA (16%) OC
GASTOS_ENVIO_COL = 37   # Column AK — Gastos de Envío OC
TOTAL_OC_COL     = 38   # Column AL — Total OC
HEADER_ROW     = 8
DATA_START_ROW = 9


class ExcelHandler:
    def __init__(self, filepath: str):
        self.filepath = filepath
        self.wb = openpyxl.load_workbook(filepath)
        self.ws = self.wb.active

    def get_folios(self) -> List[Tuple[int, str]]:
        """Return list of (row_number, folio) from column D starting at row 9."""
        result = []
        for row in range(DATA_START_ROW, self.ws.max_row + 1):
            val = self.ws.cell(row=row, column=FOLIO_COL).value
            if val is not None and str(val).strip():
                result.append((row, str(val).strip()))
        return result

    def ensure_headers(self):
        """Write column headers to row 8 if not already present."""
        headers = {
            CC_COL:           "CC OC",
            OBS_COL:          "Observaciones OC",
            SUBTOTAL_COL:     "Subtotal OC",
            DESCUENTO_COL:    "Descuento OC",
            IVA_COL:          "IVA OC",
            GASTOS_ENVIO_COL: "Gastos Envío OC",
            TOTAL_OC_COL:     "Total OC",
        }
        for col, title in headers.items():
            if not self.ws.cell(row=HEADER_ROW, column=col).value:
                self.ws.cell(row=HEADER_ROW, column=col, value=title)

    def write_result(
        self,
        row: int,
        cc: str,
        observaciones: str,
        subtotal: str = "",
        descuento: str = "",
        iva: str = "",
        gastos_envio: str = "",
        total_oc: str = "",
    ):
        self.ws.cell(row=row, column=CC_COL,           value=cc)
        self.ws.cell(row=row, column=OBS_COL,          value=observaciones)
        self.ws.cell(row=row, column=SUBTOTAL_COL,     value=subtotal)
        self.ws.cell(row=row, column=DESCUENTO_COL,    value=descuento)
        self.ws.cell(row=row, column=IVA_COL,          value=iva)
        self.ws.cell(row=row, column=GASTOS_ENVIO_COL, value=gastos_envio)
        self.ws.cell(row=row, column=TOTAL_OC_COL,     value=total_oc)

    def save(self):
        self.wb.save(self.filepath)
        self.wb.close()
