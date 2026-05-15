# Resumen Mensual Optimizado

Aplicacion local para procesar el Resumen Mensual de Transacciones (RMT) del SRI Ecuador, consolidar bases por cliente/RUC y generar una salida Excel para formularios 104 y 103 usando la plantilla oficial de trabajo.

## Stack

- Python 3.11+
- Streamlit para interfaz web local
- SQLite por cliente/RUC
- openpyxl y pandas para lectura, calculo y exportacion Excel

## Flujo principal

1. El usuario carga uno o varios RMT en Excel (`.xlsx` o `.xlsm`).
2. La app detecta encabezado, periodo, RUC, cliente y bloques del reporte.
3. Se procesan ventas, compras, notas de credito/debito, retenciones, anulados, gastos de viaje y liquidacion aduanera.
4. Se calcula una primera salida de casilleros para formulario 104 y, si aplica, 103.
5. Cada cliente queda guardado en su propia base SQLite dentro de `data/clientes/<RUC>_<cliente>/rmt.db`.
6. El feed permite revisar procesamientos, casilleros y resumen de documentos.

## Ejecutar localmente

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
streamlit run app.py
```

Abrir: http://localhost:8501

## Archivos locales sensibles

No se suben al repositorio:

- `data/rmt_in/*`
- `data/rmt_out/*`
- `*.db`, `*.sqlite`, `*.sqlite3`

## Estructura

```text
app.py                  Interfaz Streamlit
src/parser.py           Parser del RMT Excel por bloques
src/forms.py            Calculo y exportacion 104/103
src/db.py               SQLite por cliente/RUC
src/consolidator.py     Consultas y exportacion consolidada
src/models.py           Modelos internos
tests/                  Pruebas de parser y calculo base
```

## Estado actual

Version funcional inicial con RMT Excel real de prueba. Puntos que requieren validacion tributaria fina:

- Clasificacion de exportaciones entre bienes y servicios usando ATS.
- Mapeo definitivo de todos los codigos de retencion IR al formulario 103.
- Reglas especificas de credito tributario cuando el sustento no sea `01` o `02`.
