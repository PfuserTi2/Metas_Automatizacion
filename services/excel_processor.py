
import re
import unicodedata
import pandas as pd

def _norm(v):
    if pd.isna(v): return ""
    s = re.sub(r"\s+", " ", str(v).strip()).lower()
    return "".join(ch for ch in unicodedata.normalize("NFD", s) if unicodedata.category(ch) != "Mn")

def _company_id_from_text(text):
    s=_norm(text)
    # Match explicit IDs or common company names
    m=re.search(r"\b(?:empresa|id empresa|codigo empresa|c[oó]digo empresa)\s*[:#-]?\s*(13|2|252|31)\b", s)
    if m: return int(m.group(1))
    names={13:["bancolombia"],2:["western union","western"],252:["proteccion virtual"],31:["sura"]}
    for cid, vals in names.items():
        if any(x in s for x in vals): return cid
    return None

def _find_header_row(raw, required=("codigo",)):
    for i in range(min(len(raw),100)):
        vals=[_norm(x) for x in raw.iloc[i].tolist()]
        if all(any(req==v or req in v for v in vals) for req in required):
            return i
    return None

def _find_col(columns, patterns):
    for c in columns:
        s=_norm(c)
        if all(p in s for p in patterns):
            return c
    return None

def _extract_office_block(df):
    # Find the first header containing CODIGO and a Presupuesto TRX header.
    for i in range(min(len(df),100)):
        vals=[_norm(x) for x in df.iloc[i].tolist()]
        if any(v=="codigo" or v.startswith("codigo ") for v in vals) and any("presupuesto" in v for v in vals):
            header=i
            out=df.iloc[header+1:].copy()
            out.columns=[str(x).strip() for x in df.iloc[header].tolist()]
            # Stop at next row containing a repeated CODIGO header
            stop=None
            for j in range(len(out)):
                vals2=[_norm(x) for x in out.iloc[j].tolist()]
                if any(v=="codigo" for v in vals2):
                    stop=j; break
            if stop is not None: out=out.iloc[:stop]
            return out
    raise ValueError("No se encontró el bloque de oficinas con CODIGO y columnas de presupuesto.")

def _presupuesto_columns(columns, companies):
    """
    Relaciona cada ID de empresa con su columna de presupuesto.
    Acepta encabezados como:
      - Presupuesto TRX Bancolombia
      - Presupuesto Bancolombia
      - Presupuesto Sura
    No exige que todas las empresas tengan literalmente la palabra TRX.
    """
    mapping={}
    names={
        13:["bancolombia"],
        2:["western union","western"],
        252:["proteccion virtual"],
        31:["sura"]
    }
    for c in columns:
        s=_norm(c)
        if "presupuesto" not in s:
            continue
        for cid in companies:
            if cid not in names:
                continue
            if any(name in s for name in names[cid]):
                mapping[cid]=c
                break
    return mapping

def process_excel(file, year, month, companies):
    xls=pd.ExcelFile(file)
    sheet="Informe de Gestión" if "Informe de Gestión" in xls.sheet_names else xls.sheet_names[0]
    raw=pd.read_excel(file,sheet_name=sheet,header=None)
    df=_extract_office_block(raw)
    office_col=next((c for c in df.columns if _norm(c)=="codigo"), None)
    if office_col is None: raise ValueError("No se encontró la columna CODIGO.")
    active=[c for c in companies if c.get("active",True)]
    ids={int(c["id"]) for c in active}
    meta_cols=_presupuesto_columns(df.columns, ids)
    missing=ids-set(meta_cols)
    if missing:
        raise ValueError("No se encontraron columnas de presupuesto para las empresas: "+", ".join(map(str,sorted(missing))))
    records=[]
    offices=[]
    for _,row in df.iterrows():
        if pd.isna(row[office_col]): continue
        try:
            office=int(float(str(row[office_col]).replace(",","").strip()))
        except: continue
        offices.append(office)
        for c in active:
            val=row[meta_cols[int(c["id"])]]
            if pd.isna(val) or str(val).strip()=="":
                # Missing is distinct from zero; preserve valid zero.
                raise ValueError(f"La oficina {office} no tiene valor en {meta_cols[int(c['id'])]}.")
            try:
                meta=int(float(str(val).replace(",","").strip()))
            except:
                raise ValueError(f"Valor no numérico para oficina {office}, empresa {c['id']}: {val}")
            records.append({"ID Oficina":office,"Año":int(year),"Mes":int(month),"ID Empresa":int(c["id"]),"Meta":meta})
    result=pd.DataFrame(records)
    # Orden: empresa completa y, dentro de ella, todas sus oficinas.
    company_order={13:0, 2:1, 252:2, 31:3}
    result["_orden_empresa"]=result["ID Empresa"].map(company_order).fillna(999)
    result=result.sort_values(["_orden_empresa","ID Oficina"], kind="stable").drop(columns=["_orden_empresa"]).reset_index(drop=True)
    audit=result.groupby("ID Oficina")["Meta"].agg(["min","max"]).reset_index()
    return result, audit, {"sheet":sheet,"offices":len(set(offices)),"records":len(result),"meta_columns":meta_cols}

