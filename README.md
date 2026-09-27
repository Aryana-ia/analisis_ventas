# 🏗️ Pipeline de Análisis de Ventas — Streamlit

App que ejecuta el pipeline del notebook `PROanalisispower` por etapas,
con la lógica original **sin modificar** (incluido el guardado/relectura
de Excel entre pasos, del que depende el comportamiento validado).

## Estructura

| Archivo | Qué es |
|---|---|
| `app.py` | Interfaz Streamlit: uploads, ejecución por etapa, auditorías, descargas |
| `etapas.py` | Las 8 celdas del notebook, verbatim, cada una envuelta en una función |
| `requirements.txt` | Dependencias |
| `runtime.txt` | Fija Python 3.12 en Streamlit Community Cloud |

## Cambios de plomería (NO de lógica) respecto al notebook

1. Cada celda es ahora una función en `etapas.py` (cuerpo idéntico).
2. El nombre `'Listado Julio 2026 (12).xlsx'` pasó a ser parámetro
   (`etapa_1_a_4(archivo_principal)`) para poder subir cualquier listado mensual.
3. `exit()` → `raise SystemExit(1)` (mismo efecto; la app lo captura y muestra el log).
4. Los `print()` originales se capturan y se muestran como log de cada etapa.
5. El paso **TIPO DE VENTA** simple quedó fuera; solo se ejecuta
   **TIPO DE VENTA CON CASHEA**, según lo indicado.

## Uso local

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy en Streamlit Community Cloud

1. Subir esta carpeta a un repo de GitHub.
2. New app → seleccionar repo → `app.py`.
3. `runtime.txt` ya fija Python 3.12 (evita el problema de incompatibilidad
   con 3.14 que apareció en el deploy anterior).

## Flujo de trabajo

1. Subir los 4 archivos en la barra lateral: listado principal,
   `CLIENTES_UNICOS.xlsx`, `DESCRIP_UNICAS.xlsx`, `TASAS.xlsx`.
2. **🚀 Ejecutar pipeline completo**, o etapa por etapa.
3. Auditar en las pestañas 🔍 (Paso 5: clientes / Paso 7: descripciones):
   métricas de cobertura, filtro de no-unificados, búsqueda, y descarga
   de la vista filtrada como Excel para incorporar al maestro.
4. Descargar cualquier intermedio o el `Analisis_Final_DEF.xlsx` en 📦 Resultados.

## Historial de versiones

### v3 — IGTF y MONEDA
- **Paso 7.1 (IGTF):** los listados 2024/2025 no traen `IGTF`. Si falta, se crea
  con valor `0` inmediatamente después de `RETENCION`, igual que en el formato 2026.
  Si ya existe, no se toca. Así `resultado_paso_7.xlsx` en adelante tiene siempre
  la misma estructura para Power BI.
- **Pasos 10-13 (MONEDA):** nueva columna por factura, calculada sobre los pagos
  consolidados (los mismos valores MAX que recibe la fila `-A`):

  | Grupo | Columnas |
  |---|---|
  | BS  | `BSEFECTCONT`, `OTROSBSCONT`, `BSEFECTCRED`, `OTROSBSCRED` |
  | USD | `DOLAREFECTCONT`, `DOLAEXTCONT`, `DOLAREFECTCRED`, `DOLAREXTCRED`, `ANTICIPOCON`, `ANTICIPOCXC` |

  Solo BS > 0 → `BS` · solo USD > 0 → `USD` · ambos → `MIXTO` · ninguno → `SIN PAGO`.
  La etiqueta se propaga a **todas las líneas** de la factura (es un atributo, no un
  monto: no duplica importes y permite filtrar por moneda a nivel producto).
  Para contar facturas por moneda en Power BI, filtrar filas cuya `CU` termina en `-A`.
- **DEF:** `MONEDA` agregada a `ORDEN_FINAL` (después de `ANTICIPOCXC`); sin esto
  la transformación final la descartaba.
- **App:** gráfico de facturas por MONEDA en la pestaña Resultados.

### v2 — CONSTRUYE ALIANZA
- Nueva empresa `CONSTALIANZA` en mapa de empresas, Paso 5, Paso 6 y exclusiones de NETAS.
- Fix de comas faltantes en `clientes_a_prefijar` (Paso 5).
