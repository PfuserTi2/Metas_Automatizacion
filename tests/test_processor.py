
# Prueba mínima de importación
from services.excel_processor import _company_id_from_text
assert _company_id_from_text("Presupuesto TRX Bancolombia") == 13
assert _company_id_from_text("Presupuesto TRX Western Union") == 2

assert _company_id_from_text("Presupuesto Sura") == 31

assert _company_id_from_text("Presupuesto TRX Protección Virtual") == 252
