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
CONTADOR = "contador_metodos.json"

st.set_page_config(page_title="Pozos de Riego - REPDA", page_icon="💧", layout="wide")


# ───────────── Datos y contador global (compartido por todos los visitantes) ─────────────
@st.cache_resource
def datos():
    t0 = time.perf_counter_ns()
    filas = pp.cargar(ARCHIVO)
    return filas, time.perf_counter_ns() - t0


@st.cache_resource
def contador_global():
    if os.path.exists(CONTADOR):
        try:
            return Counter(json.load(open(CONTADOR)))
        except Exception:
            pass
    return Counter()


def guardar_contador(c):
    try:
        json.dump(c, open(CONTADOR, "w"), indent=2)
    except OSError:
        pass  # algunos hospedajes son de solo lectura; el contador sigue en memoria


def a_df(filas):
    return pd.DataFrame(filas, columns=pp.CAMPOS)


def ejecutar_ordenacion(filas):
    """Disparador aleatorio de ordenación + disparador aleatorio de búsqueda."""
    k_sel = random.randint(1, 3)
    campo = pp.CLAVES[k_sel]
    metodo = random.choice(pp.METODOS[k_sel])
    idx = pp.CAMPOS.index(campo)
    ct = pp.Cont()
    t0 = time.perf_counter_ns()
    ordenado = pp.ALGORITMOS[metodo](filas, idx, ct)
    t = time.perf_counter_ns() - t0
    st.session_state.orden = dict(k=k_sel, campo=campo, metodo=metodo, idx=idx, t=t, c=ct.c, m=ct.m, datos=ordenado)
    st.session_state.tipo_busqueda = random.choice(["Lineal", "Binaria"])
    st.session_state.resultado = None
    st.session_state.setdefault("registro", []).append(
        {"Ejecución #": len(st.session_state.get("registro", [])) + 1, "Campo Clave Seleccionado (1-3)": k_sel,
         "Nombre del Campo": campo, "Método de Ordenamiento Asignado": metodo, "Tiempo de Ejecución (ns)": t})
    cont = contador_global()
    cont[metodo] += 1
    guardar_contador(dict(cont))


# ───────────────────────────────── Interfaz ─────────────────────────────────
st.title("💧 Distribución y monitoreo de pozos de riego agrícola")
st.caption("Proyecto 5 · Registro Público de Derechos de Agua (REPDA – CONAGUA) · Sector primario: gestión del agua")

with st.expander("Introducción del proyecto", expanded=True):
    st.markdown(
        "La CONAGUA, a través del **REPDA**, administra los títulos de concesión y los volúmenes de extracción "
        "de los pozos destinados al riego agrícola. Esta aplicación carga **15,000 títulos** en memoria RAM, "
        "los **ordena** con un método elegido al azar según el campo clave sorteado y luego realiza una "
        "**búsqueda** (lineal o binaria), midiendo todo en nanosegundos."
    )

filas, t_carga = datos()
st.success(f"Se cargaron **{len(filas):,}** registros en memoria RAM en **{t_carga:,} nanosegundos**.")

c1, c2 = st.columns(2)
with c1:
    st.subheader("Primeros 30 registros")
    st.dataframe(a_df(filas[:30]), use_container_width=True, hide_index=True)
with c2:
    st.subheader("Últimos 30 registros")
    st.dataframe(a_df(filas[-30:]), use_container_width=True, hide_index=True)

st.divider()
st.header("1 · Disparador aleatorio de ordenación")
if st.button("🎲 Ejecutar ordenación aleatoria", type="primary"):
    with st.spinner("Ordenando 15,000 registros…"):
        ejecutar_ordenacion(filas)

o = st.session_state.get("orden")
if o:
    st.info(f"Campo clave sorteado: **{o['k']}** ({o['campo']}) · Método sorteado: **{o['metodo']}**")
    st.success(
        f"Se ordenaron {len(o['datos']):,} registros por el campo **{o['campo']}** usando el método "
        f"**{o['metodo']}** en **{o['t']:,}** nanosegundos. Comparaciones: **{o['c']:,}**, "
        f"Intercambios/Movimientos: **{o['m']:,}**"
    )
    f1, f2 = st.columns(2)
    f1.markdown("**Primeros 10 ordenados**"); f1.dataframe(a_df(o["datos"][:10]), hide_index=True, use_container_width=True)
    f2.markdown("**Últimos 10 ordenados**"); f2.dataframe(a_df(o["datos"][-10:]), hide_index=True, use_container_width=True)

    st.divider()
    st.header("2 · Disparador aleatorio de búsqueda")
    tipo = st.session_state.tipo_busqueda
    st.info(f"Tipo de búsqueda sorteada: **{tipo}**")

    if tipo == "Binaria":
        campo_b, base, idx_b = o["campo"], o["datos"], o["idx"]
        st.write(f"Se usará el campo clave previamente ordenado: **{campo_b}**")
    else:
        campo_b = st.selectbox("Campo donde buscar (clave o informativo)", pp.CAMPOS)
        base, idx_b = filas, pp.CAMPOS.index(campo_b)

    ejemplo = str(random.Random(len(st.session_state.get("registro", []))).choice(base)[idx_b])
    valor_txt = st.text_input(f"Dato a buscar en {campo_b}", placeholder=f"ej. {ejemplo}")
    if st.button("🔎 Buscar") and valor_txt.strip():
        try:
            valor = pp.convertir(valor_txt, campo_b)
        except ValueError:
            st.error("Valor no válido para ese campo.")
        else:
            fn = pp.busqueda_binaria if tipo == "Binaria" else pp.busqueda_lineal
            t0 = time.perf_counter_ns()
            hallados, cts = fn(base, idx_b, valor)
            t_b = time.perf_counter_ns() - t0
            st.session_state.resultado = dict(tipo=tipo, campo=campo_b, valor=valor, hallados=hallados, c=cts.c, t=t_b)

    r = st.session_state.get("resultado")
    if r:
        msg = (f"Búsqueda {r['tipo']} en {r['campo']} = {r['valor']}: **{len(r['hallados'])}** coincidencia(s), "
               f"**{r['c']:,}** comparaciones, **{r['t']:,}** nanosegundos.")
        (st.success if r["hallados"] else st.warning)(msg)
        if r["hallados"]:
            st.dataframe(a_df(r["hallados"][:50]), hide_index=True, use_container_width=True)
            if len(r["hallados"]) > 50:
                st.caption(f"Se muestran 50 de {len(r['hallados']):,} coincidencias.")

    st.divider()
    st.header("3 · Registro de ejecuciones")
    reg = pd.DataFrame(st.session_state.registro)
    st.dataframe(reg, hide_index=True, use_container_width=True)
    st.download_button("⬇ Descargar registro (CSV)", reg.to_csv(index=False).encode("utf-8-sig"),
                       "registro_ejecuciones.csv", "text/csv")

st.divider()
st.header("★ Método Alfa")
cont = contador_global()
if cont:
    alfa, n = cont.most_common(1)[0]
    st.success(f"Método Alfa: **{alfa}** ({n} selección/es en total)")
    st.bar_chart(pd.Series(dict(cont), name="Selecciones"))
else:
    st.write("Aún no hay ejecuciones. Pulsa el botón de ordenación aleatoria.")
