"""Lectura/escritura genérica de catálogos CSV (Distribucion/ y CuentasContables/)."""
import csv
from pathlib import Path


def read_catalog(path: Path, columns: list[str]) -> list[dict]:
    if not path.exists():
        return []
    try:
        with open(path, encoding="utf-8-sig", newline="") as f:
            return [{c: (row.get(c) or "") for c in columns} for row in csv.DictReader(f)]
    except UnicodeDecodeError:
        # Algunos catálogos (ej. Cuentas_GastoEstaciones.csv) vienen en cp1252.
        with open(path, encoding="cp1252", newline="") as f:
            return [{c: (row.get(c) or "") for c in columns} for row in csv.DictReader(f)]


def write_catalog(path: Path, columns: list[str], rows: list[dict]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=columns)
        writer.writeheader()
        writer.writerows(rows)
