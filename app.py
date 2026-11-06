"""
Proyecto 5 - Interfaz web (Streamlit). Reutiliza las funciones de proyecto_pozos.py.
Ejecutar local:  streamlit run app.py
"""
import json
import os
import random
import re
import time
from collections import Counter

import pandas as pd
import streamlit as st

import proyecto_pozos as pp  # mismas funciones: cargar, algoritmos, búsquedas, constantes

ARCHIVO = "Pozos_Riego_15000.xlsx"
CONTADOR = "contador_web.json"   # Método Alfa general
PERFILES = "perfiles_web.json"   # perfiles de usuario y su Método Alfa individual
PATRON_ALIAS = re.compile(r"[\w .\-]{2,30}")

st.set_page_config(page_title="Pozos de Riego · REPDA", page_icon="💧", layout="wide")

# ───────────── Estilo ─────────────
st.markdown("""
<style>
footer {visibility: hidden;}
.block-container {padding-top: 2.2rem; max-width: 1280px;}
.banner {background: linear-gradient(90deg, #14304D 0%, #1F4E79 100%); padding: 1.5rem 1.8rem;
         border-radius: 6px; margin-bottom: 1.2rem;}
.banner h1 {color: #FFFFFF !important; font-size: 1.85rem !important; margin: 0 !important;
            padding: 0 !important; letter-spacing: .2px;}
.banner p {color: #D6E2EF !important; margin: .4rem 0 0 0 !important; font-size: .95rem;}
.seccion {border-left: 4px solid #1F4E79; padding-left: .75rem; margin: .3rem 0 .9rem 0;
          font-size: 1.35rem; font-weight: 600; letter-spacing: .2px;}
div[data-testid="stMetric"] {border: 1px solid rgba(128,128,128,.35); border-radius: 6px;
                             padding: .6rem .9rem;}
div.stButton > button, div[data-testid="stFormSubmitButton"] > button {border-radius: 4px; font-weight: 600;}
button[data-baseweb="tab"] {font-weight: 600;}
</style>
""", unsafe_allow_html=True)


def seccion(texto):
    st.markdown(f'<div class="seccion">{texto}</div>', unsafe_allow_html=True)


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


# ───────────── Datos, contador general y perfiles ─────────────
def cargar_json(ruta, defecto):
    if os.path.exists(ruta):
        try:
            with open(ruta, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return defecto


def guardar_json(ruta, obj):
    try:
        with open(ruta, "w", encoding="utf-8") as f:
            json.dump(obj, f, ensure_ascii=False, indent=2)
    except OSError:
        pass  # algunos hospedajes son de solo lectura; los datos siguen en memoria


@st.cache_resource
def datos():
    t0 = time.perf_counter_ns()
    filas = pp.cargar(ARCHIVO)
    return filas, time.perf_counter_ns() - t0


@st.cache_resource
def contador_global():
    """Método Alfa general: acumulado de TODAS las ordenaciones mientras el servidor siga encendido."""
    crudo = cargar_json(CONTADOR, {})
    return Counter({m: n for m, n in crudo.items() if m in pp.ALGORITMOS and isinstance(n, int)}) \
        if isinstance(crudo, dict) else Counter()


@st.cache_resource
def perfiles_global():
    """Perfiles: {clave: {"nombre": alias, "conteo": {método: veces}}}."""
    crudo, ok = cargar_json(PERFILES, {}), {}
    if isinstance(crudo, dict):
        for k, v in crudo.items():
            if isinstance(v, dict) and isinstance(v.get("nombre"), str) and isinstance(v.get("conteo"), dict):
                ok[k] = {"nombre": v["nombre"][:30],
                         "conteo": {m: n for m, n in v["conteo"].items()
                                    if m in pp.ALGORITMOS and isinstance(n, int)}}
    return ok


def entrar_perfil(alias):
    """Crea el perfil si no existe, o inicia sesión si ya existe. Devuelve (ok, mensaje)."""
    alias = " ".join(alias.split())
    if not PATRON_ALIAS.fullmatch(alias):
        return False, "El alias debe tener de 2 a 30 caracteres: letras, números, espacios, punto, guion o guion bajo."
    clave = alias.casefold()
    perfiles = perfiles_global()
    if clave in perfiles:
        st.session_state.perfil = clave
        return True, f"Sesión iniciada en el perfil `{perfiles[clave]['nombre']}`."
    perfiles[clave] = {"nombre": alias, "conteo": {}}
    guardar_json(PERFILES, perfiles)
    st.session_state.perfil = clave
    return True, f"Perfil `{alias}` creado."


def a_df(filas):
    return pd.DataFrame(filas, columns=pp.CAMPOS)


def ejecutar_sorteo(filas):
    """Disparador aleatorio de ORDENACIÓN (campo clave + método). La búsqueda se sortea aparte."""
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
    ss.resultado = None  # el resultado anterior pudo hacerse sobre otra ordenación
    ss.setdefault("registro", []).append(
        {"Ejecución #": ss.n_ejec, "Campo Clave Seleccionado (1-3)": k_sel, "Nombre del Campo": campo,
         "Método de Ordenamiento Asignado": metodo, "Tiempo de Ejecución (ns)": t})
    ss.setdefault("stats", []).append({"Método": metodo, "t": t, "c": ct.c, "m": ct.m})

    cg = contador_global()          # Método Alfa general
    cg[metodo] += 1
    guardar_json(CONTADOR, dict(cg))
    perfiles = perfiles_global()    # Método Alfa del perfil activo (si hay uno)
    p = perfiles.get(ss.get("perfil"))
    if p is not None:
        p["conteo"][metodo] = p["conteo"].get(metodo, 0) + 1
        guardar_json(PERFILES, perfiles)


def sortear_busqueda():
    """Disparador aleatorio de BÚSQUEDA, independiente del de ordenación."""
    ss = st.session_state
    ss.n_busq = ss.get("n_busq", 0) + 1
    ss.tipo_busqueda = random.choice(["Lineal", "Binaria"])
    ss.resultado = None


def alfa_de(conteo):
    top = max(conteo.values())
    return [m for m, n in conteo.items() if n == top], top


def mostrar_alfa(titulo, conteo):
    conteo = {m: n for m, n in conteo.items() if n > 0}
    if not conteo:
        st.write("Aún no hay ordenaciones registradas.")
        return
    alfas, top = alfa_de(conteo)
    st.success(f"{titulo}: **{', '.join(alfas)}** ({top} selección/es" + (", empate)" if len(alfas) > 1 else ")"))
    df = pd.DataFrame({"Método": list(conteo), "Selecciones": list(conteo.values())}).sort_values(
        "Selecciones", ascending=False)
    st.dataframe(df, hide_index=True, use_container_width=True)
    st.bar_chart(df.set_index("Método"))


ss = st.session_state

# ───────────── Barra lateral: perfil de usuario ─────────────
with st.sidebar:
    st.markdown("### Perfil de usuario")
    st.caption("Crea un perfil para llevar tu propio Método Alfa, separado del general. Los perfiles no tienen "
               "contraseña y sus alias son visibles para los demás visitantes: usa un alias, no datos personales.")
    with st.form("form_perfil"):
        alias_in = st.text_input("Alias del perfil", max_chars=30, placeholder="ej. Equipo 3")
        entrar = st.form_submit_button("Entrar o crear perfil")
    if entrar:
        ok, texto = entrar_perfil(alias_in)
        (st.success if ok else st.error)(texto)
    activo = perfiles_global().get(ss.get("perfil"))
    if activo:
        st.markdown(f"Perfil activo: **`{activo['nombre']}`**")
        st.caption("Tu Método Alfa individual está en la pestaña «5 · Método Alfa».")
        if st.button("Cerrar perfil", key="btn_cerrar_perfil"):
            ss.perfil = None
            st.rerun()
    else:
        ss.perfil = None
        st.info("Sin perfil activo. Las ordenaciones solo se cuentan en el Método Alfa general.")

# ───────────────────────────────── Encabezado ─────────────────────────────────
st.markdown(
    '<div class="banner"><h1>Distribución y monitoreo de pozos de riego agrícola</h1>'
    '<p>Proyecto 5 · Registro Público de Derechos de Agua (REPDA – CONAGUA) · Sector primario: gestión del agua</p>'
    '</div>', unsafe_allow_html=True)

filas, t_carga0 = datos()
t_datos, t_ord, t_bus, t_cmp, t_alfa, t_bib = st.tabs(
    ["1 · Datos", "2 · Ordenación", "3 · Búsqueda", "4 · Análisis comparativo", "5 · Método Alfa", "6 · Bibliografía"])

# ───────────────────────────── 1 · Datos ─────────────────────────────
with t_datos:
    seccion("Presentación del proyecto y carga de datos")
    with st.expander("Introducción del proyecto", expanded=True):
        st.markdown(
            "La CONAGUA, a través del **REPDA**, administra los títulos de concesión y los volúmenes de extracción "
            "de los pozos destinados al riego agrícola. Esta aplicación carga **15,000 títulos** en memoria RAM. "
            "Permite **ordenarlos** con un método elegido al azar según el campo clave sorteado y realizar una "
            "**búsqueda** (lineal o binaria) también sorteada, midiendo todo en nanosegundos. La ordenación y la "
            "búsqueda se sortean por separado, y solo se ejecutan cuando el usuario lo solicita."
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

    t_carga = ss.get("t_carga", t_carga0)
    st.success(f"Se cargaron **{len(filas):,}** registros en memoria RAM en **{t_carga:,} nanosegundos**.")
    cc1, cc2 = st.columns([3, 1])
    cc1.caption("El tiempo mostrado inicialmente es el de la primera carga del servidor; después los datos "
                "permanecen en RAM. Puedes medir una carga nueva con el botón.")
    if cc2.button("Medir la carga de nuevo", key="btn_carga"):
        t0 = time.perf_counter_ns()
        pp.cargar(ARCHIVO)
        ss.t_carga = time.perf_counter_ns() - t0
        st.rerun()

    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Primeros 30 registros")
        st.dataframe(a_df(filas[:30]), use_container_width=True, hide_index=True)
    with c2:
        st.subheader("Últimos 30 registros")
        st.dataframe(a_df(filas[-30:]), use_container_width=True, hide_index=True)

# ───────────────────────────── 2 · Ordenación ─────────────────────────────
with t_ord:
    seccion("Disparador aleatorio de ordenación")
    st.write("Al pulsar el botón se sortea el campo clave (1 a 3) y el método de ordenamiento que le corresponde. "
             "La ordenación solo se ejecuta cuando la solicitas.")
    etiqueta = "Repetir ordenación aleatoria" if "orden" in ss else "Ejecutar ordenación aleatoria"
    if st.button(etiqueta, type="primary", key="btn_orden"):
        with st.spinner("Sorteando campo y método, y ordenando 15,000 registros…"):
            ejecutar_sorteo(filas)

    if "orden" not in ss:
        st.info("Aún no se ha ejecutado ninguna ordenación.")
    else:
        o = ss.orden
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
            st.markdown(f"**Interpretación del resultado · {o['metodo']}** — {inf['tipo']}")
            st.markdown(f"- **Idea:** {inf['idea']}")
            st.markdown(f"- **Complejidad:** mejor {inf['mejor']} · promedio {inf['prom']} · peor {inf['peor']} · "
                        f"espacio {inf['espacio']} · estable: {inf['estable']}")
            st.markdown(f"- **En esta ejecución:** {o['t'] / len(o['datos']):,.0f} ns por registro; "
                        f"{o['c'] / len(o['datos']):,.1f} comparaciones y {o['m'] / len(o['datos']):,.1f} "
                        f"movimientos por registro.")
            st.markdown(f"- **Nota:** {inf['nota']}")

        f1, f2 = st.columns(2)
        f1.markdown("**Primeros 10 ordenados**")
        f1.dataframe(a_df(o["datos"][:10]), hide_index=True, use_container_width=True)
        f2.markdown("**Últimos 10 ordenados**")
        f2.dataframe(a_df(o["datos"][-10:]), hide_index=True, use_container_width=True)

        st.divider()
        seccion("Registro de ejecuciones")
        reg = pd.DataFrame(ss.registro)
        st.dataframe(reg, hide_index=True, use_container_width=True)
        st.download_button("Descargar registro (CSV)", reg.to_csv(index=False).encode("utf-8-sig"),
                           "registro_ejecuciones.csv", "text/csv")

        st.subheader("Resumen por método (esta sesión)")
        est = pd.DataFrame(ss.stats)
        resumen = est.groupby("Método").agg(
            Veces=("t", "size"), **{"Tiempo promedio (ns)": ("t", "mean"), "Tiempo mínimo (ns)": ("t", "min"),
                                    "Tiempo máximo (ns)": ("t", "max"), "Comparaciones (prom.)": ("c", "mean"),
                                    "Movimientos (prom.)": ("m", "mean")}).round(0).astype(int).reset_index()
        st.dataframe(resumen, hide_index=True, use_container_width=True)
        st.caption("Cada método se ejecuta sobre el campo que le tocó en el sorteo, así que estos tiempos no son "
                   "directamente comparables entre métodos. Para una comparación justa usa la pestaña 4.")

# ───────────────────────────── 3 · Búsqueda ─────────────────────────────
with t_bus:
    seccion("Disparador aleatorio de búsqueda")
    st.write("El tipo de búsqueda (lineal o binaria) se sortea de forma independiente a la ordenación.")
    if st.button("Sortear tipo de búsqueda", type="primary", key="btn_busq"):
        sortear_busqueda()

    tipo = ss.get("tipo_busqueda")
    if tipo is None:
        st.info("Aún no se ha sorteado el tipo de búsqueda.")
    else:
        st.info(f"Tipo de búsqueda sorteada: **{tipo}**")
        if tipo == "Binaria" and "orden" not in ss:
            st.warning("La búsqueda binaria requiere un campo clave ordenado. Ejecuta primero la ordenación en la "
                       "pestaña «2 · Ordenación» y regresa aquí.")
        else:
            n_b, n_e = ss.get("n_busq", 0), ss.get("n_ejec", 0)
            if tipo == "Binaria":
                o = ss.orden
                campo_b, base, idx_b = o["campo"], o["datos"], o["idx"]
                st.write(f"Se usará el campo clave de la última ordenación: **{campo_b}** (método {o['metodo']}).")
            else:
                campo_b = st.selectbox("Campo donde buscar (clave o informativo)", pp.CAMPOS,
                                       key=f"campo_busq_{n_b}")
                base, idx_b = filas, pp.CAMPOS.index(campo_b)

            es_texto = campo_b in pp.CAMPOS_TEXTO
            ejemplo = str(random.Random(n_b * 1000 + n_e).choice(base)[idx_b])
            st.caption("Formato esperado: " + (AYUDA_TEXTO if es_texto else AYUDA[campo_b]))
            with st.form(f"form_busqueda_{n_b}_{n_e}"):
                valor_txt = st.text_input(f"Dato a buscar en {campo_b}", placeholder=f"ej. {ejemplo}",
                                          key=f"valor_{n_b}_{n_e}_{campo_b}")
                enviar = st.form_submit_button("Buscar")

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
                    ss.resultado = dict(tipo=tipo, campo=campo_b, valor=valor, hallados=hallados,
                                        c=cts.c, t=t_b, texto=es_texto and tipo == "Lineal")

            r = ss.get("resultado")
            if r:
                msg = (f"Búsqueda {r['tipo']} en {r['campo']} = {r['valor']}: **{len(r['hallados']):,}** "
                       f"coincidencia(s), **{r['c']:,}** comparaciones, **{r['t']:,}** nanosegundos.")
                if r["texto"]:
                    msg += " (Sin distinguir mayúsculas/acentos; coincidencia parcial.)"
                (st.success if r["hallados"] else st.warning)(msg)
                if r["hallados"]:
                    st.dataframe(a_df(r["hallados"][:50]), hide_index=True, use_container_width=True)
                    if len(r["hallados"]) > 50:
                        st.caption(f"Se muestran 50 de {len(r['hallados']):,} coincidencias.")
                else:
                    st.caption("Sin resultados: revisa el formato o prueba con un dato que exista "
                               "(el ejemplo gris del campo existe).")

# ───────────────────────────── 4 · Análisis comparativo ─────────────────────────────
with t_cmp:
    seccion("Análisis comparativo (complementario)")
    st.caption("Ejecuta varios métodos sobre el mismo campo para compararlos en igualdad de condiciones. "
               "No modifica el registro ni el Método Alfa.")
    campo_cmp = st.selectbox("Campo a comparar", list(APLICABLES), key="campo_cmp")
    if st.button("Comparar métodos", type="primary", key="btn_cmp"):
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
        ss.comparacion = (campo_cmp, d)

    if ss.get("comparacion"):
        cc, dcmp = ss.comparacion
        st.markdown(f"**Resultado sobre `{cc}`** — el más rápido fue **{dcmp.iloc[0]['Método']}**.")
        st.dataframe(dcmp, hide_index=True, use_container_width=True)
        st.caption("Los tiempos varían un poco entre corridas por la carga del servidor; repite para confirmar "
                   "tendencias.")
    else:
        st.info("Aún no se ha ejecutado ninguna comparación.")

# ───────────────────────────── 5 · Método Alfa ─────────────────────────────
with t_alfa:
    seccion("Método Alfa")
    st.write("El Método Alfa es el método de ordenamiento más veces seleccionado por el sorteo. Se lleva por "
             "separado para tu perfil y para el conjunto de todas las ordenaciones.")
    a1, a2 = st.columns(2)
    with a1:
        st.subheader("Perfil activo")
        activo = perfiles_global().get(ss.get("perfil"))
        if activo:
            mostrar_alfa(f"Método Alfa de `{activo['nombre']}`", activo["conteo"])
        else:
            st.info("Crea o inicia un perfil en la barra lateral para ver tu Método Alfa individual. Solo se "
                    "cuentan las ordenaciones realizadas con el perfil activo.")
    with a2:
        st.subheader("General")
        mostrar_alfa("Método Alfa general", contador_global())
        st.caption("Incluye todas las ordenaciones de todos los visitantes. Se reinicia si el servidor se reinicia.")

    perfiles = perfiles_global()
    if perfiles:
        st.divider()
        st.subheader("Perfiles registrados")
        filas_p = []
        for p in perfiles.values():
            cont = {m: n for m, n in p["conteo"].items() if n > 0}
            filas_p.append({"Perfil": p["nombre"], "Ordenaciones": sum(cont.values()),
                            "Método Alfa": ", ".join(alfa_de(cont)[0]) if cont else "—"})
        st.dataframe(pd.DataFrame(filas_p).sort_values("Ordenaciones", ascending=False),
                     hide_index=True, use_container_width=True)

# ───────────────────────────── 6 · Bibliografía ─────────────────────────────
with t_bib:
    seccion("Bibliografía y recursos utilizados")

    st.subheader("Fuentes de datos")
    st.markdown(
        "- Comisión Nacional del Agua [CONAGUA]. (s. f.). *Registro Público de Derechos de Agua (REPDA)*. "
        "Gobierno de México. Reporte de títulos de concesión (`ReporteRepda.csv`). "
        "[https://www.gob.mx/conagua](https://www.gob.mx/conagua)\n"
        "- Gobierno de México. (s. f.). *Catálogo de Datos Abiertos del Gobierno Federal*. "
        "[https://www.datos.gob.mx/](https://www.datos.gob.mx/)"
    )
    st.caption("Los campos municipio, profundidad, estatus y caudal del pozo no provienen de estas fuentes: "
               "fueron simulados (ver «Origen de los datos» en la pestaña 1).")

    st.subheader("Herramientas de software")
    st.markdown(
        "- Python Software Foundation. *Python*. [https://www.python.org](https://www.python.org)\n"
        "- Streamlit Inc. *Streamlit* (interfaz web y Streamlit Community Cloud para el despliegue público). "
        "[https://docs.streamlit.io](https://docs.streamlit.io)\n"
        "- The pandas development team. *pandas* (manejo de tablas de datos). "
        "[https://pandas.pydata.org](https://pandas.pydata.org)\n"
        "- Gazoni, E. y Clark, C. *openpyxl* (lectura y escritura de archivos Excel). "
        "[https://openpyxl.readthedocs.io](https://openpyxl.readthedocs.io)\n"
        "- GitHub, Inc. *GitHub* (repositorio del código). [https://github.com](https://github.com)"
    )

    st.subheader("Uso de inteligencia artificial")
    st.markdown(
        "- Anthropic. (2026). *Claude* (modelo Claude Sonnet 5.5) [Asistente de inteligencia artificial]. "
        "[https://claude.ai](https://claude.ai)"
    )
    st.markdown(
        "Claude se utilizó como asistente de programación durante todo el proyecto para: analizar el documento de "
        "requerimientos; identificar los campos del Excel y mapearlos a las columnas del REPDA; escribir los "
        "scripts en Python que preparan los 15,000 registros; implementar los algoritmos de ordenamiento (Radix, "
        "Counting, Bucket, Quick, Merge y Heap Sort) y de búsqueda (lineal y binaria) con su medición en "
        "nanosegundos; construir la interfaz web con Streamlit, incluidos los perfiles de usuario y el rediseño "
        "visual; depurar errores; preparar el despliegue público; y proponer mejoras a partir de una "
        "retroalimentación integral del proyecto."
    )
