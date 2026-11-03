"""
Proyecto 5 - Interfaz web (Streamlit). Reutiliza las funciones de proyecto_pozos.py.
Ejecutar local:  streamlit run app.py
"""
import json
import os
import random
import time
from collections import Counter

import pandas as pd
import streamlit as st

import proyecto_pozos as pp  # mismas funciones: cargar, algoritmos, búsquedas, constantes

ARCHIVO = "Pozos_Riego_15000.xlsx"
CONTADOR = "contador_web.json"  # archivo propio de la web (ignora el contador_metodos.json de las pruebas)

st.set_page_config(page_title="Pozos de Riego - REPDA", page_icon="💧", layout="wide")

# ───────────── Fichas teóricas de cada método (para interpretar los resultados) ─────────────
INFO = {
    "Radix Sort": dict(
        tipo="Distribución (no compara elementos)", mejor="O(d·(n+b))", prom="O(d·(n+b))", peor="O(d·(n+b))",
        espacio="O(n+b)", estable="Sí",
        idea="Ordena dígito por dígito, del menos al más significativo, con un conteo estable en cada posición.",
        nota="No compara elementos entre sí, por eso las comparaciones son 0. Los movimientos equivalen a "
             "n × (número de dígitos del mayor valor)."),
    "Counting Sort": dict(
        tipo="Distribución (no compara elementos)", mejor="O(n+k)", prom="O(n+k)", peor="O(n+k)",
        espacio="O(n+k)", estable="Sí",
        idea="Cuenta cuántas veces aparece cada valor y usa esos conteos para colocar cada registro en su posición final.",
        nota="No compara elementos entre sí, por eso las comparaciones son 0. Los movimientos suman los "
             "incrementos del arreglo de conteo y las ubicaciones finales (2 × n). Requiere memoria proporcional "
             "al valor máximo (k)."),
    "Bucket Sort": dict(
        tipo="Distribución por cubetas", mejor="O(n+k)", prom="O(n+k)", peor="O(n²)",
        espacio="O(n+k)", estable="Sí (con inserción)",
        idea="Reparte los registros en cubetas por rango de valor (escala logarítmica, porque el volumen está muy "
             "sesgado) y ordena cada cubeta por inserción.",
        nota="Las comparaciones ocurren solo dentro de cada cubeta. Si muchos registros caen en la misma cubeta, "
             "se acerca al peor caso O(n²)."),
    "Quick Sort": dict(
        tipo="Comparación (divide y vencerás)", mejor="O(n log n)", prom="O(n log n)", peor="O(n²)",
        espacio="O(log n)", estable="No",
        idea="Elige un pivote (mediana de tres), reparte los elementos a ambos lados con partición de Hoare y "
             "repite en cada lado.",
        nota="Suele ser el más rápido entre los de comparación, pero un pivote desafortunado lo lleva a O(n²)."),
    "Merge Sort": dict(
        tipo="Comparación (divide y vencerás)", mejor="O(n log n)", prom="O(n log n)", peor="O(n log n)",
        espacio="O(n)", estable="Sí",
        idea="Divide la lista a la mitad, ordena cada mitad y las mezcla en orden.",
        nota="Tiempo predecible en todos los casos, a cambio de memoria adicional para las mezclas."),
    "Heap Sort": dict(
        tipo="Comparación (montículo)", mejor="O(n log n)", prom="O(n log n)", peor="O(n log n)",
        espacio="O(1)", estable="No",
        idea="Construye un montículo máximo y extrae repetidamente la raíz para colocarla al final.",
        nota="No usa memoria adicional, pero hace más comparaciones y accesos dispersos que Quick o Merge."),
}

# Métodos que se pueden comparar en el análisis extra, por campo (se excluyen combinaciones inviables en memoria)
APLICABLES = {
    "id_titulo_concesion": ["Radix Sort", "Counting Sort", "Quick Sort", "Merge Sort", "Heap Sort"],
    "volumen_extraido_m3": ["Radix Sort", "Bucket Sort", "Quick Sort", "Merge Sort", "Heap Sort"],
    "fecha_concesion": ["Radix Sort", "Quick Sort", "Merge Sort", "Heap Sort"],
}

AYUDA = {
    "id_titulo_concesion": "Número entero de 6 dígitos",
    "volumen_extraido_m3": "Número entero en m³ (sin comas)",
    "fecha_concesion": "Fecha como número AAAAMMDD",
    "profundidad_pozo_m": "Número decimal exacto, con punto (ej. 120.5)",
    "caudal_litros_segundo": "Número decimal exacto, con punto (ej. 35.8)",
}
AYUDA_TEXTO = "Texto: no distingue mayúsculas ni acentos y acepta coincidencias parciales"


# ───────────── Datos y contadores ─────────────
@st.cache_resource
def datos():
    t0 = time.perf_counter_ns()
    filas = pp.cargar(ARCHIVO)
    return filas, time.perf_counter_ns() - t0


@st.cache_resource
def contador_global():
    """Acumulado de TODOS los visitantes mientras el servidor siga encendido."""
    if os.path.exists(CONTADOR):
        try:
            return Counter(json.load(open(CONTADOR)))
        except Exception:
            pass
    return Counter()


def guardar_contador(c):
    try:
        json.dump(dict(c), open(CONTADOR, "w"), indent=2)
    except OSError:
        pass  # algunos hospedajes son de solo lectura; el contador sigue en memoria


def a_df(filas):
    return pd.DataFrame(filas, columns=pp.CAMPOS)


def ejecutar_sorteo(filas):
    """Disparador aleatorio de ordenación + disparador aleatorio de búsqueda."""
    k_sel = random.randint(1, 3)
    campo = pp.CLAVES[k_sel]
    metodo = random.choice(pp.METODOS[k_sel])
    idx = pp.CAMPOS.index(campo)
    ct = pp.Cont()
    t0 = time.perf_counter_ns()
    ordenado = pp.ALGORITMOS[metodo](filas, idx, ct)
    t = time.perf_counter_ns() - t0

    ss = st.session_state
    ss.n_ejec = ss.get("n_ejec", 0) + 1
    ss.orden = dict(k=k_sel, campo=campo, metodo=metodo, idx=idx, t=t, c=ct.c, m=ct.m, datos=ordenado)
    ss.tipo_busqueda = random.choice(["Lineal", "Binaria"])
    ss.resultado = None
    ss.setdefault("registro", []).append(
        {"Ejecución #": ss.n_ejec, "Campo Clave Seleccionado (1-3)": k_sel, "Nombre del Campo": campo,
         "Método de Ordenamiento Asignado": metodo, "Tiempo de Ejecución (ns)": t})
    ss.setdefault("stats", []).append({"Método": metodo, "t": t, "c": ct.c, "m": ct.m})
    ss.setdefault("cont_sesion", Counter())[metodo] += 1
    cg = contador_global()
    cg[metodo] += 1
    guardar_contador(cg)


def mostrar_alfa(titulo, cont):
    if not cont:
        st.write("Aún no hay selecciones.")
        return
    top = max(cont.values())
    alfas = [m for m, n in cont.items() if n == top]
    st.success(f"{titulo}: **{', '.join(alfas)}** ({top} selección/es" + (", empate)" if len(alfas) > 1 else ")"))
    df = pd.DataFrame({"Método": list(cont), "Selecciones": list(cont.values())}).sort_values(
        "Selecciones", ascending=False)
    st.dataframe(df, hide_index=True, use_container_width=True)
    st.bar_chart(df.set_index("Método"))


# ───────────────────────────────── Interfaz ─────────────────────────────────
st.title("💧 Distribución y monitoreo de pozos de riego agrícola")
st.caption("Proyecto 5 · Registro Público de Derechos de Agua (REPDA – CONAGUA) · Sector primario: gestión del agua")

with st.expander("Introducción del proyecto", expanded=True):
    st.markdown(
        "La CONAGUA, a través del **REPDA**, administra los títulos de concesión y los volúmenes de extracción "
        "de los pozos destinados al riego agrícola. Esta aplicación carga **15,000 títulos** en memoria RAM, "
        "los **ordena** con un método elegido al azar según el campo clave sorteado y luego realiza una "
        "**búsqueda** (lineal o binaria), midiendo todo en nanosegundos. El sorteo se ejecuta automáticamente "
        "al abrir la página y puede repetirse."
    )

with st.expander("Origen de los datos: qué campos son reales y cuáles simulados"):
    st.markdown(
        "Los registros provienen del reporte público del REPDA (aprovechamientos de uso **agrícola** con "
        "extracción de **aguas subterráneas**). Ese reporte **no publica** municipio, profundidad, estatus ni "
        "caudal del pozo, por lo que esos cuatro campos se **simularon** para completar la estructura de 10 "
        "campos pedida en el proyecto."
    )
    st.dataframe(pd.DataFrame([
        ("id_titulo_concesion", "Real", "Folio numérico del número de título"),
        ("volumen_extraido_m3", "Real", "Volumen de extracción anual, redondeado"),
        ("fecha_concesion", "Real", "Fecha de registro del título (AAAAMMDD)"),
        ("nombre_titular", "Real", "Titular del título"),
        ("estado", "Real", "Entidad federativa tomada de la clave del título"),
        ("uso_principal", "Real", "Uso agrícola"),
        ("municipio", "Simulado", "Municipio agrícola del estado, al azar"),
        ("profundidad_pozo_m", "Simulado", "Valor aleatorio entre 15 y 450 m"),
        ("estatus_pozo", "Simulado", "Activo, Mantenimiento, Suspendido o Clausurado"),
        ("caudal_litros_segundo", "Simulado", "Estimado a partir del volumen anual (12 h/día de bombeo)"),
    ], columns=["Campo", "Origen", "Detalle"]), hide_index=True, use_container_width=True)

filas, t_carga0 = datos()
t_carga = st.session_state.get("t_carga", t_carga0)
st.success(f"Se cargaron **{len(filas):,}** registros en memoria RAM en **{t_carga:,} nanosegundos**.")
cc1, cc2 = st.columns([3, 1])
cc1.caption("El tiempo mostrado inicialmente es el de la primera carga del servidor; después los datos permanecen "
            "en RAM. Puedes medir una carga nueva con el botón.")
if cc2.button("⏱ Medir la carga de nuevo"):
    t0 = time.perf_counter_ns()
    pp.cargar(ARCHIVO)
    st.session_state.t_carga = time.perf_counter_ns() - t0
    st.rerun()

c1, c2 = st.columns(2)
with c1:
    st.subheader("Primeros 30 registros")
    st.dataframe(a_df(filas[:30]), use_container_width=True, hide_index=True)
with c2:
    st.subheader("Últimos 30 registros")
    st.dataframe(a_df(filas[-30:]), use_container_width=True, hide_index=True)

# ── 1. Ordenación (se dispara sola al abrir la página; el botón solo la repite)
st.divider()
st.header("1 · Disparador aleatorio de ordenación")
if "orden" not in st.session_state:
    with st.spinner("Sorteando campo y método, y ordenando 15,000 registros…"):
        ejecutar_sorteo(filas)
elif st.button("🔁 Repetir sorteo de ordenación", type="primary"):
    with st.spinner("Ordenando 15,000 registros…"):
        ejecutar_sorteo(filas)

o = st.session_state.orden
n_ej = st.session_state.n_ejec
st.info(f"Campo clave sorteado: **{o['k']}** ({o['campo']}) · Método sorteado: **{o['metodo']}**")
st.success(
    f"Se ordenaron {len(o['datos']):,} registros por el campo **{o['campo']}** usando el método "
    f"**{o['metodo']}** en **{o['t']:,}** nanosegundos. Comparaciones: **{o['c']:,}**, "
    f"Intercambios/Movimientos: **{o['m']:,}**"
)
m1, m2, m3 = st.columns(3)
m1.metric("Tiempo (ns)", f"{o['t']:,}")
m2.metric("Comparaciones", f"{o['c']:,}")
m3.metric("Intercambios / Movimientos", f"{o['m']:,}")

inf = INFO[o["metodo"]]
with st.container(border=True):
    st.markdown(f"**¿Cómo interpretar este resultado? · {o['metodo']}** — {inf['tipo']}")
    st.markdown(f"- **Idea:** {inf['idea']}")
    st.markdown(f"- **Complejidad:** mejor {inf['mejor']} · promedio {inf['prom']} · peor {inf['peor']} · "
                f"espacio {inf['espacio']} · estable: {inf['estable']}")
    st.markdown(f"- **En esta ejecución:** {o['t'] / len(o['datos']):,.0f} ns por registro; "
                f"{o['c'] / len(o['datos']):,.1f} comparaciones y {o['m'] / len(o['datos']):,.1f} movimientos "
                f"por registro.")
    st.markdown(f"- **Nota:** {inf['nota']}")

f1, f2 = st.columns(2)
f1.markdown("**Primeros 10 ordenados**")
f1.dataframe(a_df(o["datos"][:10]), hide_index=True, use_container_width=True)
f2.markdown("**Últimos 10 ordenados**")
f2.dataframe(a_df(o["datos"][-10:]), hide_index=True, use_container_width=True)

# ── 2. Búsqueda
st.divider()
st.header("2 · Disparador aleatorio de búsqueda")
tipo = st.session_state.tipo_busqueda
st.info(f"Tipo de búsqueda sorteada: **{tipo}**")

if tipo == "Binaria":
    campo_b, base, idx_b = o["campo"], o["datos"], o["idx"]
    st.write(f"Se usará el campo clave previamente ordenado: **{campo_b}**")
else:
    campo_b = st.selectbox("Campo donde buscar (clave o informativo)", pp.CAMPOS, key=f"campo_busq_{n_ej}")
    base, idx_b = filas, pp.CAMPOS.index(campo_b)

es_texto = campo_b in pp.CAMPOS_TEXTO
ejemplo = str(random.Random(n_ej).choice(base)[idx_b])
st.caption("Formato esperado: " + (AYUDA_TEXTO if es_texto else AYUDA[campo_b]))
with st.form(f"form_busqueda_{n_ej}"):
    valor_txt = st.text_input(f"Dato a buscar en {campo_b}", placeholder=f"ej. {ejemplo}",
                              key=f"valor_{n_ej}_{campo_b}")
    enviar = st.form_submit_button("🔎 Buscar")

if enviar and valor_txt.strip():
    limpio = valor_txt.strip()
    if not es_texto:
        limpio = limpio.replace(" ", "")
        limpio = limpio.replace(",", ".") if campo_b in ("profundidad_pozo_m", "caudal_litros_segundo") \
            else limpio.replace(",", "")
    try:
        valor = pp.convertir(limpio, campo_b)
    except ValueError:
        st.error("Valor no válido para ese campo. Revisa el formato esperado.")
    else:
        if tipo == "Binaria":
            fn = pp.busqueda_binaria
        elif es_texto:
            fn = pp.busqueda_lineal_texto
        else:
            fn = pp.busqueda_lineal
        t0 = time.perf_counter_ns()
        hallados, cts = fn(base, idx_b, valor)
        t_b = time.perf_counter_ns() - t0
        st.session_state.resultado = dict(tipo=tipo, campo=campo_b, valor=valor, hallados=hallados,
                                          c=cts.c, t=t_b, texto=es_texto and tipo == "Lineal")

r = st.session_state.get("resultado")
if r:
    msg = (f"Búsqueda {r['tipo']} en {r['campo']} = {r['valor']}: **{len(r['hallados']):,}** coincidencia(s), "
           f"**{r['c']:,}** comparaciones, **{r['t']:,}** nanosegundos.")
    if r["texto"]:
        msg += " (Sin distinguir mayúsculas/acentos; coincidencia parcial.)"
    (st.success if r["hallados"] else st.warning)(msg)
    if r["hallados"]:
        st.dataframe(a_df(r["hallados"][:50]), hide_index=True, use_container_width=True)
        if len(r["hallados"]) > 50:
            st.caption(f"Se muestran 50 de {len(r['hallados']):,} coincidencias.")
    else:
        st.caption("Sin resultados: revisa el formato o prueba con un dato que exista (el ejemplo gris del campo existe).")

# ── 3. Registro y resumen por método
st.divider()
st.header("3 · Registro de ejecuciones")
reg = pd.DataFrame(st.session_state.registro)
st.dataframe(reg, hide_index=True, use_container_width=True)
st.download_button("⬇ Descargar registro (CSV)", reg.to_csv(index=False).encode("utf-8-sig"),
                   "registro_ejecuciones.csv", "text/csv")

st.subheader("Resumen por método (esta sesión)")
est = pd.DataFrame(st.session_state.stats)
resumen = est.groupby("Método").agg(
    Veces=("t", "size"), **{"Tiempo promedio (ns)": ("t", "mean"), "Tiempo mínimo (ns)": ("t", "min"),
                            "Tiempo máximo (ns)": ("t", "max"), "Comparaciones (prom.)": ("c", "mean"),
                            "Movimientos (prom.)": ("m", "mean")}).round(0).astype(int).reset_index()
st.dataframe(resumen, hide_index=True, use_container_width=True)
st.caption("Cada método se ejecuta sobre el campo que le tocó en el sorteo, así que estos tiempos no son "
           "directamente comparables entre métodos. Para una comparación justa usa el análisis siguiente.")

# ── 4. Análisis comparativo extra
st.divider()
st.header("4 · Análisis comparativo (extra)")
st.caption("Ejecuta varios métodos sobre el mismo campo para compararlos en igualdad de condiciones. "
           "No modifica el sorteo, el registro ni el conteo del método Alfa.")
campo_cmp = st.selectbox("Campo a comparar", list(APLICABLES), index=list(APLICABLES).index(o["campo"]))
if st.button("▶ Comparar métodos"):
    idx_c = pp.CAMPOS.index(campo_cmp)
    filas_cmp = []
    with st.spinner("Ejecutando métodos…"):
        for met in APLICABLES[campo_cmp]:
            ctc = pp.Cont()
            t0 = time.perf_counter_ns()
            res = pp.ALGORITMOS[met](filas, idx_c, ctc)
            tc = time.perf_counter_ns() - t0
            ok = all(res[i][idx_c] <= res[i + 1][idx_c] for i in range(len(res) - 1))
            filas_cmp.append({"Método": met, "Tiempo (ns)": tc, "Comparaciones": ctc.c,
                              "Movimientos": ctc.m, "Orden correcto": "Sí" if ok else "No"})
    d = pd.DataFrame(filas_cmp).sort_values("Tiempo (ns)").reset_index(drop=True)
    d["Veces más lento que el mejor"] = (d["Tiempo (ns)"] / d["Tiempo (ns)"].min()).round(1)
    st.session_state.comparacion = (campo_cmp, d)

if st.session_state.get("comparacion"):
    cc, dcmp = st.session_state.comparacion
    st.markdown(f"**Resultado sobre `{cc}`** — el más rápido fue **{dcmp.iloc[0]['Método']}**.")
    st.dataframe(dcmp, hide_index=True, use_container_width=True)
    st.caption("Los tiempos varían un poco entre corridas por la carga del servidor; repite para confirmar tendencias.")

# ── 5. Método Alfa
st.divider()
st.header("★ Método Alfa")
a1, a2 = st.columns(2)
with a1:
    st.subheader("Esta sesión")
    mostrar_alfa("Método Alfa (tu sesión)", st.session_state.cont_sesion)
with a2:
    st.subheader("Todos los visitantes")
    mostrar_alfa("Método Alfa (acumulado)", contador_global())
    st.caption("El acumulado se reinicia si el servidor se reinicia.")
