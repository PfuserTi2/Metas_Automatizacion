from pathlib import Path
import io
import json
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font
from services.excel_processor import process_excel, infer_period

BASE = Path(__file__).resolve().parent
CONFIG = json.loads((BASE / "config" / "config.json").read_text(encoding="utf-8"))

st.set_page_config(page_title="Automatización de Metas", page_icon="📊", layout="wide")
st.title("📊 Automatización de Metas — V4")

st.info(
    "Regla: la Meta es Equivalentes Esperadas de cada oficina. "
    "Se trunca la parte decimal, sin redondear. "
    "0, 3, 6 y cualquier otro valor son metas válidas; no existe un mínimo."
)

file = st.file_uploader("Carga el Excel mensual", type=["xlsx", "xlsm"])

if file:
    detected_year, detected_month = infer_period(file.name)
    c1, c2 = st.columns(2)
    with c1:
        year = st.number_input("Año", 2020, 2100, detected_year or 2026)
    with c2:
        month = st.number_input("Mes", 1, 12, detected_month or 1)

    if st.button("🚀 Procesar", type="primary"):
        try:
            df, audit, header = process_excel(
                file.getvalue(), file.name, int(year), int(month), CONFIG
            )

            st.success("Reporte procesado correctamente.")

            a, b, c, d = st.columns(4)
            a.metric("Oficinas", audit["offices"])
            b.metric("Metas en 0", audit["zero_meta_offices"])
            c.metric("Empresas", audit["companies"])
            d.metric("Registros", f'{audit["records"]:,}')

            st.subheader("🔎 Auditoría de oficinas")
            audit_df = pd.DataFrame(audit["office_meta"], columns=["ID Oficina", "Meta"])
            st.dataframe(audit_df, use_container_width=True, hide_index=True, height=350)

            if audit["low_meta_offices"]:
                st.warning(
                    "Se encontraron metas pequeñas. Todas se conservaron; "
                    "no se aplica ningún filtro por valor."
                )
                st.dataframe(
                    pd.DataFrame(audit["low_meta_offices"], columns=["ID Oficina", "Meta"]),
                    use_container_width=True, hide_index=True
                )

            if audit["ignored_non_numeric_codes"]:
                st.caption(
                    "Códigos no numéricos ignorados: "
                    + ", ".join(audit["ignored_non_numeric_codes"])
                )

            if audit["ignored_rows_without_meta"]:
                st.warning(
                    "Filas con CODIGO numérico pero sin Equivalentes Esperadas: "
                    + ", ".join(map(str, audit["ignored_rows_without_meta"]))
                )

            st.subheader("📄 Resultado final")
            st.dataframe(df, use_container_width=True, hide_index=True, height=500)

            csv = df.to_csv(index=False, sep=";", encoding="utf-8-sig").encode("utf-8-sig")

            bio = io.BytesIO()
            with pd.ExcelWriter(bio, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name="RESULTADO_METAS", index=False)
                audit_df.to_excel(writer, sheet_name="AUDITORIA_OFICINAS", index=False)

            bio.seek(0)
            wb = load_workbook(bio)
            for ws in [wb["RESULTADO_METAS"], wb["AUDITORIA_OFICINAS"]]:
                for cell in ws[1]:
                    cell.font = Font(bold=True)
                ws.freeze_panes = "A2"
                ws.auto_filter.ref = ws.dimensions

            out = io.BytesIO()
            wb.save(out)

            x, y = st.columns(2)
            with x:
                st.download_button(
                    "⬇️ Descargar Excel", out.getvalue(),
                    f"METAS_{int(year)}_{int(month):02d}.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            with y:
                st.download_button(
                    "⬇️ Descargar CSV", csv,
                    f"METAS_{int(year)}_{int(month):02d}.csv",
                    "text/csv", use_container_width=True
                )
        except Exception as e:
            st.error("No fue posible procesar el archivo.")
            st.exception(e)
