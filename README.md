# Automatización de Metas V5

La meta ahora se obtiene exclusivamente de la columna **Presupuesto TRX** correspondiente a cada empresa y oficina.

- No usa Equivalentes Esperadas.
- La meta es por oficina + empresa.
- Cero y metas pequeñas son válidas.
- No hay filtro por valor mínimo.
- Conserva la parte entera del presupuesto sin redondear.
- Detecta el bloque correcto de oficinas y se detiene ante el siguiente encabezado CODIGO.

## Ejecución
```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python -m streamlit run app.py
```


## Nota V6
La detección de columnas de presupuesto es flexible: no exige que todas tengan la palabra `TRX`. Por ejemplo, `Presupuesto Sura` se asigna automáticamente al ID Empresa 31.


## V7
La descarga ahora es un archivo Excel `.xlsx` real, con las columnas separadas: ID Oficina, Año, Mes, ID Empresa y Meta. También incluye una hoja Auditoria.


## V8
El resultado se ordena por empresa en este orden: 13, 2, 252 y 31. Dentro de cada empresa aparecen todas sus oficinas ordenadas por ID de oficina.


## V9
La empresa 252 corresponde a **Protección Virtual** y utiliza exclusivamente la columna **Presupuesto TRX Protección Virtual**.
