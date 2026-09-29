# AutomatizacionMetas_V4

## Estructura
- `app.py`: interfaz.
- `requirements.txt`: librerías.
- `config/config.json`: empresas y nombres de columnas.
- `services/excel_processor.py`: lógica de lectura y cálculo.
- `data/entrada`: archivos originales (opcional).
- `data/procesados`: archivos procesados (opcional).
- `data/resultados`: resultados (opcional).
- `logs`: registros futuros.
- `tests`: pruebas futuras.

## Regla de negocio
- Oficina = `CODIGO` numérico del primer bloque de `Informe de Gestión`.
- Meta = `Equivalentes Esperadas`.
- Se trunca la parte decimal, no se redondea.
- `0`, `3`, `6` y cualquier otro valor son metas válidas.
- No existe mínimo de meta.
- La meta pertenece a la oficina y se replica para cada empresa activa.
- El segundo encabezado `CODIGO` termina el bloque de oficinas.

## Ejecutar
### Opción automática en Windows
Doble clic en `iniciar_windows.bat`.

### Opción manual
```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python -m streamlit run app.py
```
