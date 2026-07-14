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

## Nota importante (bug latente en el notebook, NO corregido aquí)

En el Paso 5, la lista `clientes_a_prefijar` tiene una **coma faltante** entre
`'JUAN ALBERTO REINA OLAYA'` y `'CONSTRUCIONES BARINAS'`: Python concatena ambos
strings y ninguno de los dos recibe el prefijo `.` por esta vía. Se mantuvo tal
cual porque la consigna fue no modificar la lógica. Corregirlo = agregar una coma
en `etapas.py`, función `etapa_5`.
