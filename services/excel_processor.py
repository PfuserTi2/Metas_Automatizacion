from io import BytesIO
from pathlib import Path
import re
import pandas as pd


def infer_period(filename):
    stem = Path(filename).stem
    m = re.search(r"(?<!\d)(\d{2})(\d{2})(\d{4})(?!\d)", stem)
    if m:
        day, month, year = map(int, m.groups())
        if 1 <= month <= 12:
            return year, month
    m = re.search(r"(?<!\d)(\d{4})[-_](\d{1,2})(?!\d)", stem)
    if m:
        return int(m.group(1)), int(m.group(2))
    return None, None


def _find_header_row(raw, required_column):
    target = required_column.strip().lower()
    for i in range(len(raw)):
        values = {str(v).strip().lower() for v in raw.iloc[i].tolist() if pd.notna(v)}
        if target in values:
            return i
    raise ValueError(f'No se encontró una fila de encabezados con "{required_column}".')


def _unique_columns(columns):
    seen, result = {}, []
    for value in columns:
        base = str(value).strip() if pd.notna(value) else ""
        base = base or "SIN_NOMBRE"
        n = seen.get(base, 0)
        seen[base] = n + 1
        result.append(base if n == 0 else f"{base}_{n}")
    return result


def _office_id(value):
    if pd.isna(value):
        return None
    n = pd.to_numeric(value, errors="coerce")
    return None if pd.isna(n) else int(n)


def _meta(value):
    # 0 es una meta válida. Se trunca, no se redondea.
    if pd.isna(value) or str(value).strip() == "":
        return None
    if isinstance(value, (int, float)):
        return int(float(value))
    s = str(value).strip().replace(" ", "")
    if "." in s and "," in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    return int(float(s))


def process_excel(file_bytes, filename, year, month, config):
    source = config["source"]
    sheet = source["preferred_sheet"]
    office_col = source["office_id_column"]
    meta_col = source["meta_column"]

    companies = [c for c in config["companies"] if c.get("active", True)]
    if not companies:
        raise ValueError("No hay empresas activas.")

    raw = pd.read_excel(
        BytesIO(file_bytes),
        sheet_name=sheet,
        header=None,
        engine="openpyxl"
    )

    header = _find_header_row(raw, office_col)
    df = raw.iloc[header + 1:].copy()
    df.columns = _unique_columns(raw.iloc[header].tolist())

    if office_col not in df.columns:
        raise ValueError(f'No se encontró "{office_col}".')
    if meta_col not in df.columns:
        raise ValueError(f'No se encontró "{meta_col}" en "{sheet}".')

    office_meta = {}
    ignored_non_numeric = []
    ignored_without_meta = []

    for _, row in df.iterrows():
        raw_code = row[office_col]

        # El segundo encabezado CODIGO termina el bloque real de oficinas.
        if str(raw_code).strip().upper() == office_col.upper():
            break

        office = _office_id(raw_code)

        # Solo códigos numéricos son oficinas.
        if office is None:
            if pd.notna(raw_code) and str(raw_code).strip():
                ignored_non_numeric.append(str(raw_code).strip())
            continue

        raw_meta = row[meta_col]

        # Meta vacía = fila auxiliar. Meta 0 se conserva.
        if pd.isna(raw_meta) or str(raw_meta).strip() == "":
            ignored_without_meta.append(office)
            continue

        meta = _meta(raw_meta)
        if meta is None:
            raise ValueError(f"La oficina {office} tiene una meta inválida: {raw_meta!r}.")

        if office in office_meta and office_meta[office] != meta:
            raise ValueError(
                f"La oficina {office} aparece con dos metas diferentes "
                f"({office_meta[office]} y {meta})."
            )

        office_meta[office] = meta

    if not office_meta:
        raise ValueError("No se encontraron oficinas válidas.")

    records = []
    for office, meta in office_meta.items():
        for company in companies:
            records.append({
                "ID Oficina": office,
                "Año": int(year),
                "Mes": int(month),
                "ID Empresa": int(company["id"]),
                "Meta": int(meta),
            })

    result = pd.DataFrame(
        records,
        columns=["ID Oficina", "Año", "Mes", "ID Empresa", "Meta"]
    )

    audit = {
        "offices": len(office_meta),
        "companies": len(companies),
        "records": len(result),
        "zero_meta_offices": sum(1 for x in office_meta.values() if x == 0),
        "low_meta_offices": sorted(
            [(office, meta) for office, meta in office_meta.items() if 0 <= meta < 100],
            key=lambda x: (x[1], x[0])
        ),
        "ignored_non_numeric_codes": ignored_non_numeric,
        "ignored_rows_without_meta": ignored_without_meta,
        "office_meta": sorted(office_meta.items(), key=lambda x: x[0]),
    }

    return result, audit, header + 1
