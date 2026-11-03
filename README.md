# Proyecto 5 · Distribución y monitoreo de pozos de riego agrícola

Aplicación web (Streamlit) que carga 15,000 títulos de concesión del REPDA (CONAGUA) en memoria RAM, los ordena
con un método elegido al azar según el campo clave sorteado y ejecuta una búsqueda lineal o binaria, midiendo
tiempos en nanosegundos, comparaciones y movimientos.

## Archivos

| Archivo | Para qué sirve |
|---|---|
| `app.py` | Interfaz web (Streamlit) |
| `proyecto_pozos.py` | Algoritmos de ordenamiento y búsqueda, carga de datos y versión de consola |
| `Pozos_Riego_15000.xlsx` | Datos del proyecto (hoja `Datos`), formato de registro y mapeo de campos |
| `preparar_datos.py` | Genera el Excel a partir de `ReporteRepda.csv` |
| `requirements.txt` | Dependencias |

## Datos

- **Fuente:** reporte público del Registro Público de Derechos de Agua (REPDA – CONAGUA), `ReporteRepda.csv`.
- **Selección:** títulos de uso agrícola con extracción de aguas subterráneas y folio único; muestra aleatoria
  reproducible de 15,000 (semilla 2026).
- **Campos reales:** `id_titulo_concesion`, `volumen_extraido_m3`, `fecha_concesion` (fecha de registro),
  `nombre_titular`, `estado`, `uso_principal`.
- **Campos simulados** (el REPDA no los publica): `municipio`, `profundidad_pozo_m`, `estatus_pozo`,
  `caudal_litros_segundo`. La hoja `Mapeo_Campos` del Excel documenta cada transformación.

## Algoritmos

| Campo clave | Métodos del sorteo |
|---|---|
| `id_titulo_concesion` | Radix Sort, Counting Sort |
| `volumen_extraido_m3` | Bucket Sort |
| `fecha_concesion` | Quick Sort, Merge Sort, Heap Sort |

Búsqueda lineal sobre cualquier campo (en texto ignora mayúsculas y acentos y acepta coincidencias parciales) y
búsqueda binaria sobre el campo clave ordenado en la fase anterior.

## Ejecutar

```
py -m pip install -r requirements.txt
py -m streamlit run app.py
```

Versión de consola: `py proyecto_pozos.py`

## Uso de Claude

Proyecto desarrollado con Claude (Anthropic) como asistente de programación: análisis del documento, mapeo de
campos, scripts en Python, interfaz web y despliegue.
