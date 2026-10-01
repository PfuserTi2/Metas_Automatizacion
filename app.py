
import streamlit as st
from services.excel_processor import process_excel
import pandas as pd
import io

st.set_page_config(page_title="Automatización de Metas", layout="wide")
st.title("📊 Automatización de Metas")
st.caption("Fuente de meta: presupuesto de cada empresa por oficina. El resultado se ordena por empresa y luego por oficina.")

uploaded=st.file_uploader("Carga el Excel de cumplimiento", type=["xlsx","xls"])
c1,c2=st.columns(2)
year=c1.number_input("Año", min_value=2000, max_value=2100, value=2026)
month=c2.number_input("Mes", min_value=1, max_value=12, value=9)

companies=[
 {"id":13,"name":"Bancolombia","active":True},
 {"id":2,"name":"Western Union","active":True},
 {"id":252,"name":"Proteccion","active":True},
 {"id":31,"name":"Sura","active":True},
]
if uploaded and st.button("🚀 Procesar"):
    try:
        result,audit,info=process_excel(uploaded,year,month,companies)
        st.success("Procesamiento completado.")
        a,b,c=st.columns(3)
        a.metric("Oficinas", info["offices"])
        b.metric("Empresas", len([x for x in companies if x["active"]]))
        c.metric("Registros", info["records"])
        st.info("Las metas se toman exclusivamente de la columna de presupuesto correspondiente a cada empresa. Equivalentes Esperadas no se utiliza.")
        st.write("### Columnas de presupuesto detectadas")
        st.json({str(k):str(v) for k,v in info["meta_columns"].items()})
        st.write("### Auditoría por oficina")
        audit.columns=["ID Oficina","Meta mínima","Meta máxima"]
        st.dataframe(audit,use_container_width=True)
        st.write("### Resultado")
        st.dataframe(result,use_container_width=True)
        # Exportación real a Excel: cada campo queda en su propia columna.
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            result.to_excel(writer, index=False, sheet_name="Metas")
            audit.to_excel(writer, index=False, sheet_name="Auditoria")
        output.seek(0)
        st.download_button(
            "⬇️ Descargar Excel",
            output.getvalue(),
            "metas_resultado.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(str(e))
