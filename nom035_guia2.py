"""
NOM-035-STPS-2018 | RFRANYUTTI, CONCIENCIA VERDE Y LABORAL S.C.
================================================================
MÓDULO 0 + 2  —  Panel Operativo + Cuestionario Virtual
GUÍA II       —  Identificación y Análisis de los Factores de
                 Riesgo Psicosocial y Evaluación del Entorno
                 Organizacional
(Para centros de trabajo con 16 a 50 trabajadores)
"""

import streamlit as st
import pandas as pd
import os, re, time, sys, io, textwrap, sqlite3, threading
from datetime import datetime

# ── Modo empleado: detectar ?cliente=XXXX en la URL ───────────────────────────
st.set_page_config(page_title="NOM-035 · Guía II", page_icon="🏭",
                   layout="centered", initial_sidebar_state="collapsed")

_params        = st.query_params
_MODO_EMPLEADO = "cliente" in _params
_CLIENTE_URL   = str(_params.get("cliente", "")).upper() if _MODO_EMPLEADO else ""

# ── Pantalla de carga inicial ────────────────────────────────────────────────
if "app_loaded_g2" not in st.session_state:
    st.session_state.app_loaded_g2 = False

if not st.session_state.app_loaded_g2:
    _load_ph = st.empty()
    _load_ph.markdown("""
    <div class="loading-overlay">
        <div class="loading-logo" style="font-family:Montserrat,sans-serif;font-size:1.1rem;
             font-weight:700;color:#4b694e;letter-spacing:.08em;">RFRANYUTTI</div>
        <div style="font-size:.72rem;color:#888;font-family:Montserrat,sans-serif;
             font-weight:500;margin-top:.3rem;letter-spacing:.1em;">
            CONCIENCIA VERDE Y LABORAL S.C.</div>
        <div class="loading-ring"></div>
        <div class="loading-txt">NOM-035-STPS-2018 · Guía II · Cargando...</div>
    </div>
    """, unsafe_allow_html=True)
    import time as _t; _t.sleep(1.2)
    _load_ph.empty()
    st.session_state.app_loaded_g2 = True

# ── Clientes ──────────────────────────────────────────────────────────────────
CLIENTES = {
    "FRUCO": {
        "razon": "FRUTAS CONCENTRADAS, S.A.P.I. DE C.V.",
        "logo":  "assets/logos/fruco.png",
        "opciones": ["FRUTAS CONCENTRADAS, S.A.P.I. DE C.V."],
    },
    "QUALTIA": {
        "razon": "QUALTIA ALIMENTOS Y OPERACIONES, S. DE R.L. DE C.V.",
        "logo":  "assets/logos/qualtia.png",
        "opciones": [
            "QUALTIA ALIMENTOS Y OPERACIONES, S. DE R.L. DE C.V.",
            "QUALTIA ALIMENTOS OPERACIONES, S. DE R.L. DE C.V. (CEDIS Y SERVICIOS AUXILIARES)",
        ],
    },
    "DIABLOS": {
        "razon": "CENTRO DEPORTIVO ALFREDO HARP HELÚ, S.A. DE C.V.",
        "logo":  "assets/logos/Diablos.png",
        "opciones": ["CENTRO DEPORTIVO ALFREDO HARP HELÚ, S.A. DE C.V."],
    },
}
LOGO_RF = "assets/logos/rfranyutti.gif"

def excel_path(cliente_key: str, razon_social: str = "") -> str:
    if cliente_key == "QUALTIA" and "CEDIS" in razon_social.upper():
        return "data/g2_resultados_QUALTIA_CEDIS.xlsx"
    return f"data/g2_resultados_{cliente_key.upper()}.xlsx"

# ── Catálogos ──────────────────────────────────────────────────────────────────
SEL = "— Selecciona —"
OPC_RESP    = ["Siempre", "Casi siempre", "Algunas veces", "Casi nunca", "Nunca"]
OPC_SEXO    = [SEL, "Femenino", "Masculino"]
OPC_EDAD    = [SEL,"15 - 19","20 - 24","25 - 29","30 - 34","35 - 39","40 - 44",
               "45 - 49","50 - 54","55 - 59","60 - 64","65 - 69","70 o más"]
OPC_ECIVIL  = [SEL,"Soltero","Casado","Unión libre","Divorciado","Viudo"]
OPC_ESTUD   = [SEL,"Sin formación","Primaria","Secundaria","Preparatoria o Bachillerato",
               "Técnico Superior","Licenciatura","Maestría","Doctorado"]
OPC_PUESTO  = [SEL,"Operativo","Supervisor","Profesional o técnico","Gerente"]
OPC_CONTRAT = [SEL,"Por obra o proyecto","Tiempo indeterminado",
               "Por tiempo determinado (temporal)","Honorarios"]
OPC_PERSONAL= [SEL,"Sindicalizado","Confianza","Ninguno"]
OPC_JORNADA = [SEL,"Fijo diurno (entre las 6:00 y 20:00 hrs)",
               "Fijo nocturno (entre las 20:00 y 6:00 hrs)",
               "Fijo mixto (combinación de nocturno y diurno)"]
OPC_TPUESTO = [SEL,"Menos de 6 meses","Entre 6 meses y 1 año","Entre 1 a 4 años",
               "Entre 5 a 9 años","Entre 10 a 14 años","Entre 15 a 19 años",
               "Entre 20 a 24 años","25 años o más"]
OPC_EXP     = [SEL,"Menos de 6 meses","Entre 6 meses y 1 año","Entre 1 a 4 años",
               "Entre 5 a 9 años","Entre 10 a 14 años","Entre 15 a 19 años",
               "Entre 20 a 24 años","25 años o más"]

# Escala de respuestas → puntaje
ESCALA = {"Siempre": 4, "Casi siempre": 3, "Algunas veces": 2,
          "Casi nunca": 1, "Nunca": 0}

# ── PREGUNTAS OFICIALES GUÍA II NOM-035 (46 ítems) ────────────────────────────
# Estructura: {categoria, dominio, pregunta}
# Categorías: 1=Ambiente, 2=Actividad, 3=Tiempo, 4=Liderazgo
# Dominios: 1=Carga, 2=Control, 3=Jornada, 4=Interf, 5=Liderazgo, 6=Relaciones, 7=Violencia

PREGUNTAS_G2 = [
    # CATEGORÍA 1 — Condiciones en el ambiente de trabajo
    {"id":1,  "cat":1, "dom":0, "cat_nombre":"Condiciones en el ambiente de trabajo",
     "dom_nombre":"Condiciones del ambiente",
     "texto":"Mi trabajo me exige hacer mucho esfuerzo físico."},
    {"id":2,  "cat":1, "dom":0, "cat_nombre":"Condiciones en el ambiente de trabajo",
     "dom_nombre":"Condiciones del ambiente",
     "texto":"Me preocupa sufrir un accidente en mi trabajo."},
    {"id":3,  "cat":1, "dom":0, "cat_nombre":"Condiciones en el ambiente de trabajo",
     "dom_nombre":"Condiciones del ambiente",
     "texto":"Considero que en mi trabajo se presentan riesgos que pueden afectar mi salud."},
    {"id":4,  "cat":1, "dom":0, "cat_nombre":"Condiciones en el ambiente de trabajo",
     "dom_nombre":"Condiciones del ambiente",
     "texto":"En mi trabajo puedo tener poco o nada de movimiento."},

    # CATEGORÍA 2 — Factores propios de la actividad / DOMINIO 1: Carga de trabajo
    {"id":5,  "cat":2, "dom":1, "cat_nombre":"Factores propios de la actividad",
     "dom_nombre":"Carga de trabajo",
     "texto":"Por la cantidad de trabajo que tengo debo quedarme tiempo adicional a mi turno."},
    {"id":6,  "cat":2, "dom":1, "cat_nombre":"Factores propios de la actividad",
     "dom_nombre":"Carga de trabajo",
     "texto":"Por la cantidad de trabajo que tengo debo llevarme trabajo a casa."},
    {"id":7,  "cat":2, "dom":1, "cat_nombre":"Factores propios de la actividad",
     "dom_nombre":"Carga de trabajo",
     "texto":"Debo atender varias actividades al mismo tiempo."},
    {"id":8,  "cat":2, "dom":1, "cat_nombre":"Factores propios de la actividad",
     "dom_nombre":"Carga de trabajo",
     "texto":"En mi trabajo debo hacer esfuerzo mental importante para recordar muchas cosas."},
    {"id":9,  "cat":2, "dom":1, "cat_nombre":"Factores propios de la actividad",
     "dom_nombre":"Carga de trabajo",
     "texto":"En mi trabajo me exigen concentrarme demasiado."},
    {"id":10, "cat":2, "dom":1, "cat_nombre":"Factores propios de la actividad",
     "dom_nombre":"Carga de trabajo",
     "texto":"Mi trabajo exige un nivel de atención muy elevado."},
    {"id":11, "cat":2, "dom":1, "cat_nombre":"Factores propios de la actividad",
     "dom_nombre":"Carga de trabajo",
     "texto":"En mi trabajo soy responsable de cosas de mucho valor."},
    {"id":12, "cat":2, "dom":1, "cat_nombre":"Factores propios de la actividad",
     "dom_nombre":"Carga de trabajo",
     "texto":"En mi trabajo soy responsable de la seguridad de otros."},
    {"id":13, "cat":2, "dom":1, "cat_nombre":"Factores propios de la actividad",
     "dom_nombre":"Carga de trabajo",
     "texto":"En mi trabajo me dan órdenes contradictorias."},
    {"id":14, "cat":2, "dom":1, "cat_nombre":"Factores propios de la actividad",
     "dom_nombre":"Carga de trabajo",
     "texto":"En mi trabajo hay situaciones que me hacen tener reacciones emocionales que afectan mi desempeño (enfado, llanto, etc.)."},
    {"id":15, "cat":2, "dom":1, "cat_nombre":"Factores propios de la actividad",
     "dom_nombre":"Carga de trabajo",
     "texto":"Mi trabajo permite que desarrolle nuevas habilidades."},
    {"id":16, "cat":2, "dom":1, "cat_nombre":"Factores propios de la actividad",
     "dom_nombre":"Carga de trabajo",
     "texto":"En mi trabajo puedo aplicar mis habilidades y mis conocimientos."},

    # DOMINIO 2: Falta de control sobre el trabajo
    {"id":17, "cat":2, "dom":2, "cat_nombre":"Factores propios de la actividad",
     "dom_nombre":"Falta de control sobre el trabajo",
     "texto":"Me informan con claridad cuáles son mis funciones."},
    {"id":18, "cat":2, "dom":2, "cat_nombre":"Factores propios de la actividad",
     "dom_nombre":"Falta de control sobre el trabajo",
     "texto":"Me explican claramente los resultados que debo obtener en mi trabajo."},
    {"id":19, "cat":2, "dom":2, "cat_nombre":"Factores propios de la actividad",
     "dom_nombre":"Falta de control sobre el trabajo",
     "texto":"Me informan con quién puedo resolver los problemas o asuntos de trabajo."},
    {"id":20, "cat":2, "dom":2, "cat_nombre":"Factores propios de la actividad",
     "dom_nombre":"Falta de control sobre el trabajo",
     "texto":"Me permiten organizar mi trabajo como mejor lo considero."},
    {"id":21, "cat":2, "dom":2, "cat_nombre":"Factores propios de la actividad",
     "dom_nombre":"Falta de control sobre el trabajo",
     "texto":"Me permiten tomar decisiones en mi trabajo."},
    {"id":22, "cat":2, "dom":2, "cat_nombre":"Factores propios de la actividad",
     "dom_nombre":"Falta de control sobre el trabajo",
     "texto":"Me proporcionan capacitación necesaria para hacer mi trabajo."},

    # CATEGORÍA 3 — Organización del tiempo de trabajo / DOMINIO 3: Jornada
    {"id":23, "cat":3, "dom":3, "cat_nombre":"Organización del tiempo de trabajo",
     "dom_nombre":"Jornada de trabajo",
     "texto":"Mi jornada de trabajo me permite atender mis necesidades personales."},
    {"id":24, "cat":3, "dom":3, "cat_nombre":"Organización del tiempo de trabajo",
     "dom_nombre":"Jornada de trabajo",
     "texto":"Puedo tomar pausas cuando las necesito en mi trabajo."},
    {"id":25, "cat":3, "dom":3, "cat_nombre":"Organización del tiempo de trabajo",
     "dom_nombre":"Jornada de trabajo",
     "texto":"Mi trabajo me permite tener vacaciones."},

    # DOMINIO 4: Interferencia en la relación trabajo-familia
    {"id":26, "cat":3, "dom":4, "cat_nombre":"Organización del tiempo de trabajo",
     "dom_nombre":"Interferencia en la relación trabajo-familia",
     "texto":"Puedo atender las necesidades de mis familiares o personales cuando lo necesito."},
    {"id":27, "cat":3, "dom":4, "cat_nombre":"Organización del tiempo de trabajo",
     "dom_nombre":"Interferencia en la relación trabajo-familia",
     "texto":"Puedo separar fácilmente los problemas familiares de los laborales."},
    {"id":28, "cat":3, "dom":4, "cat_nombre":"Organización del tiempo de trabajo",
     "dom_nombre":"Interferencia en la relación trabajo-familia",
     "texto":"Cuando estoy en casa sigo pensando en el trabajo."},
    {"id":29, "cat":3, "dom":4, "cat_nombre":"Organización del tiempo de trabajo",
     "dom_nombre":"Interferencia en la relación trabajo-familia",
     "texto":"Hay momentos en que necesitaría estar en el trabajo y en casa al mismo tiempo."},
    {"id":30, "cat":3, "dom":4, "cat_nombre":"Organización del tiempo de trabajo",
     "dom_nombre":"Interferencia en la relación trabajo-familia",
     "texto":"Mi trabajo me quita tiempo que quisiera dedicar a mi familia o a mis actividades personales."},

    # CATEGORÍA 4 — Liderazgo y relaciones / DOMINIO 5: Liderazgo
    {"id":31, "cat":4, "dom":5, "cat_nombre":"Liderazgo y relaciones en el trabajo",
     "dom_nombre":"Liderazgo",
     "texto":"Mi jefe tiene en cuenta mis puntos de vista y opiniones."},
    {"id":32, "cat":4, "dom":5, "cat_nombre":"Liderazgo y relaciones en el trabajo",
     "dom_nombre":"Liderazgo",
     "texto":"Mi jefe me comunica a tiempo la información que necesito para hacer mi trabajo."},
    {"id":33, "cat":4, "dom":5, "cat_nombre":"Liderazgo y relaciones en el trabajo",
     "dom_nombre":"Liderazgo",
     "texto":"Mi jefe me permite participar en la toma de decisiones del trabajo."},
    {"id":34, "cat":4, "dom":5, "cat_nombre":"Liderazgo y relaciones en el trabajo",
     "dom_nombre":"Liderazgo",
     "texto":"Mi jefe distribuye el trabajo de forma equitativa entre sus subordinados."},
    {"id":35, "cat":4, "dom":5, "cat_nombre":"Liderazgo y relaciones en el trabajo",
     "dom_nombre":"Liderazgo",
     "texto":"Mi jefe me orienta y da retroalimentación sobre mi trabajo."},
    {"id":36, "cat":4, "dom":5, "cat_nombre":"Liderazgo y relaciones en el trabajo",
     "dom_nombre":"Liderazgo",
     "texto":"Mi jefe me ayuda a solucionar los problemas que se presentan en el trabajo."},
    {"id":37, "cat":4, "dom":5, "cat_nombre":"Liderazgo y relaciones en el trabajo",
     "dom_nombre":"Liderazgo",
     "texto":"En mi empresa, el jefe nos da el reconocimiento que merecemos."},
    {"id":38, "cat":4, "dom":5, "cat_nombre":"Liderazgo y relaciones en el trabajo",
     "dom_nombre":"Liderazgo",
     "texto":"En la empresa donde trabajo existe justicia en la forma de resolver los conflictos."},
    {"id":39, "cat":4, "dom":5, "cat_nombre":"Liderazgo y relaciones en el trabajo",
     "dom_nombre":"Liderazgo",
     "texto":"En mi trabajo existe comunicación eficiente entre compañeros."},

    # DOMINIO 6: Relaciones en el trabajo
    {"id":40, "cat":4, "dom":6, "cat_nombre":"Liderazgo y relaciones en el trabajo",
     "dom_nombre":"Relaciones en el trabajo",
     "texto":"En mi trabajo hay apoyo entre compañeros cuando alguien está en problemas."},
    {"id":41, "cat":4, "dom":6, "cat_nombre":"Liderazgo y relaciones en el trabajo",
     "dom_nombre":"Relaciones en el trabajo",
     "texto":"En mi trabajo los compañeros tienen en cuenta mis puntos de vista y opiniones."},
    {"id":42, "cat":4, "dom":6, "cat_nombre":"Liderazgo y relaciones en el trabajo",
     "dom_nombre":"Relaciones en el trabajo",
     "texto":"En mi trabajo puedo confiar en mis compañeros."},
    {"id":43, "cat":4, "dom":6, "cat_nombre":"Liderazgo y relaciones en el trabajo",
     "dom_nombre":"Relaciones en el trabajo",
     "texto":"Cuando tengo que hacer tareas difíciles mis compañeros me apoyan."},

    # DOMINIO 7: Violencia laboral (ítems 44-46) — MÁS DELICADOS
    {"id":44, "cat":4, "dom":7, "cat_nombre":"Liderazgo y relaciones en el trabajo",
     "dom_nombre":"Violencia laboral",
     "texto":"En mi trabajo me ignoran o me excluyen."},
    {"id":45, "cat":4, "dom":7, "cat_nombre":"Liderazgo y relaciones en el trabajo",
     "dom_nombre":"Violencia laboral",
     "texto":"En mi trabajo soy objeto de burlas, calumnias, falsedades, humillaciones o ridiculizaciones."},
    {"id":46, "cat":4, "dom":7, "cat_nombre":"Liderazgo y relaciones en el trabajo",
     "dom_nombre":"Violencia laboral",
     "texto":"En mi trabajo me exigen responsabilidades que no corresponden a mis funciones o se me asignan tareas sin los recursos necesarios."},
]

# ── Ítems de calificación INVERTIDA (respuesta favorable = Siempre → menor riesgo)
# Estos ítems se puntúan al revés: Siempre=0, Casi siempre=1... Nunca=4
ITEMS_DIRECTOS = {1,2,3,4,5,6,7,8,9,10,11,12,13,14,28,29,30,44,45,46}
ITEMS_INVERSOS = {15,16,17,18,19,20,21,22,23,24,25,26,27,31,32,33,34,35,36,37,38,39,40,41,42,43}

ESCALA_DIRECTA  = {"Siempre":4,"Casi siempre":3,"Algunas veces":2,"Casi nunca":1,"Nunca":0}
ESCALA_INVERSA  = {"Siempre":0,"Casi siempre":1,"Algunas veces":2,"Casi nunca":3,"Nunca":4}

# Tabla 5 NOM-035 — Puntos de corte Guía II (46 ítems, puntaje máx = 184)
TABLA5 = [
    (0,  20,  "NULO",     "#4B694E", "#D6E4D8"),
    (21, 45,  "BAJO",     "#69A2D8", "#EBF3FB"),
    (46, 70,  "MEDIO",    "#C8A600", "#FFF8DC"),
    (71, 90,  "ALTO",     "#E07820", "#FDEBD0"),
    (91, 999, "MUY ALTO", "#A20000", "#FDDEDE"),
]

# Tabla de puntos de corte por DOMINIO
TABLA_DOM = {
    # dom_id: (nulo_max, bajo_max, medio_max, alto_max)
    0: (4,  8,  12, 16),    # Condiciones ambiente (4 ítems, max 16)
    1: (8,  16, 24, 32),    # Carga de trabajo (12 ítems puntaje max 48)
    2: (2,  4,  8,  10),    # Control (6 ítems max 24)
    3: (2,  4,  6,  8),     # Jornada (3 ítems max 12)
    4: (3,  6,  9,  12),    # Interferencia (5 ítems max 20)
    5: (4,  8,  14, 22),    # Liderazgo (9 ítems max 36)
    6: (2,  4,  6,  8),     # Relaciones (4 ítems max 16)
    7: (0,  2,  4,  6),     # Violencia (3 ítems max 12) — umbral bajo
}

def nivel_riesgo(puntaje: int) -> dict:
    """Clasifica el nivel de riesgo global según Tabla 5 NOM-035."""
    for pmin, pmax, nivel, color, bg in TABLA5:
        if pmin <= puntaje <= pmax:
            return {"nivel": nivel, "color": color, "bg": bg,
                    "pmin": pmin, "pmax": pmax}
    return {"nivel": "MUY ALTO", "color": "#A20000", "bg": "#FDDEDE",
            "pmin": 91, "pmax": 999}

def nivel_dominio(dom_id: int, puntaje: int) -> str:
    cortes = TABLA_DOM.get(dom_id, (4, 8, 12, 16))
    if puntaje <= cortes[0]:   return "NULO"
    if puntaje <= cortes[1]:   return "BAJO"
    if puntaje <= cortes[2]:   return "MEDIO"
    if puntaje <= cortes[3]:   return "ALTO"
    return "MUY ALTO"

def calcular_puntaje(respuestas: list) -> dict:
    """
    Calcula puntaje total, por dominio y por categoría.
    respuestas: lista de 46 respuestas en orden (índice 0 = ítem 1)
    """
    if len(respuestas) < 46:
        return {}

    puntaje_total = 0
    por_dominio   = {0:0, 1:0, 2:0, 3:0, 4:0, 5:0, 6:0, 7:0}
    por_categoria = {1:0, 2:0, 3:0, 4:0}
    items_dom     = {0:[], 1:[], 2:[], 3:[], 4:[], 5:[], 6:[], 7:[]}
    alerta_violencia = False

    for preg in PREGUNTAS_G2:
        idx   = preg["id"] - 1
        resp  = respuestas[idx] if idx < len(respuestas) else "Nunca"
        id_p  = preg["id"]
        escala = ESCALA_DIRECTA if id_p in ITEMS_DIRECTOS else ESCALA_INVERSA
        pts   = escala.get(resp, 0)
        puntaje_total        += pts
        por_dominio[preg["dom"]] += pts
        por_categoria[preg["cat"]] += pts
        items_dom[preg["dom"]].append({"id": id_p, "resp": resp, "pts": pts})
        # Alerta violencia: ítem 44-46 con puntaje > 0 (es decir no "Nunca")
        if preg["dom"] == 7 and pts > 0:
            alerta_violencia = True

    nivel = nivel_riesgo(puntaje_total)
    niveles_dom = {dom: nivel_dominio(dom, pts)
                   for dom, pts in por_dominio.items()}

    return {
        "puntaje_total":   puntaje_total,
        "nivel":           nivel["nivel"],
        "color":           nivel["color"],
        "bg":              nivel["bg"],
        "por_dominio":     por_dominio,
        "niveles_dom":     niveles_dom,
        "por_categoria":   por_categoria,
        "items_dom":       items_dom,
        "alerta_violencia": alerta_violencia,
    }

# ── Excel ──────────────────────────────────────────────────────────────────────
def init_excel(path: str):
    os.makedirs("data", exist_ok=True)
    if not os.path.exists(path):
        cols = ["Folio","Fecha","Cliente","Razón Social","Nombre","Sexo","Edad",
                "Estado Civil","Nivel Estudios","Estud. Status","Puesto","Área",
                "Contratación","Tipo Personal","Jornada","Rotación Turnos",
                "Tiempo Puesto","Experiencia",
                "Puntaje Total","Nivel Riesgo",
                "Cat1 Ambiente","Cat2 Actividad","Cat3 Tiempo","Cat4 Liderazgo",
                "Dom0 Ambiente","Dom1 Carga","Dom2 Control","Dom3 Jornada",
                "Dom4 Interferencia","Dom5 Liderazgo","Dom6 Relaciones","Dom7 Violencia",
                "Nivel Dom0","Nivel Dom1","Nivel Dom2","Nivel Dom3",
                "Nivel Dom4","Nivel Dom5","Nivel Dom6","Nivel Dom7",
                "Alerta Violencia"]
        for i in range(1, 47):
            cols.append(f"P{i:02d}")
        pd.DataFrame(columns=cols).to_excel(path, index=False)

def guardar(data: dict) -> bool:
    path = excel_path(data["cliente"], data.get("razon", ""))
    init_excel(path)
    res  = data["resultado"]
    for intento in range(5):
        try:
            df   = pd.read_excel(path)
            fila = {
                "Folio":         data["folio"],
                "Fecha":         datetime.now().strftime("%Y-%m-%d %H:%M"),
                "Cliente":       data["cliente"],
                "Razón Social":  data["razon"],
                "Nombre":        data["nombre"],
                "Sexo":          data["sexo"],
                "Edad":          data["edad"],
                "Estado Civil":  data["ecivil"],
                "Nivel Estudios":data["estudios"],
                "Estud. Status": data["estatus"],
                "Puesto":        data["puesto"],
                "Área":          data["area"],
                "Contratación":  data["contrat"],
                "Tipo Personal": data["personal"],
                "Jornada":       data["jornada"],
                "Rotación Turnos": data["rotacion"],
                "Tiempo Puesto": data["tpuesto"],
                "Experiencia":   data["exp"],
                "Puntaje Total": res["puntaje_total"],
                "Nivel Riesgo":  res["nivel"],
                "Cat1 Ambiente": res["por_categoria"].get(1, 0),
                "Cat2 Actividad":res["por_categoria"].get(2, 0),
                "Cat3 Tiempo":   res["por_categoria"].get(3, 0),
                "Cat4 Liderazgo":res["por_categoria"].get(4, 0),
                "Dom0 Ambiente": res["por_dominio"].get(0, 0),
                "Dom1 Carga":    res["por_dominio"].get(1, 0),
                "Dom2 Control":  res["por_dominio"].get(2, 0),
                "Dom3 Jornada":  res["por_dominio"].get(3, 0),
                "Dom4 Interferencia": res["por_dominio"].get(4, 0),
                "Dom5 Liderazgo":res["por_dominio"].get(5, 0),
                "Dom6 Relaciones":res["por_dominio"].get(6, 0),
                "Dom7 Violencia":res["por_dominio"].get(7, 0),
                "Nivel Dom0":    res["niveles_dom"].get(0, ""),
                "Nivel Dom1":    res["niveles_dom"].get(1, ""),
                "Nivel Dom2":    res["niveles_dom"].get(2, ""),
                "Nivel Dom3":    res["niveles_dom"].get(3, ""),
                "Nivel Dom4":    res["niveles_dom"].get(4, ""),
                "Nivel Dom5":    res["niveles_dom"].get(5, ""),
                "Nivel Dom6":    res["niveles_dom"].get(6, ""),
                "Nivel Dom7":    res["niveles_dom"].get(7, ""),
                "Alerta Violencia": "SÍ — URGENTE" if res["alerta_violencia"] else "No",
            }
            for i, resp in enumerate(data.get("respuestas", []), 1):
                fila[f"P{i:02d}"] = resp
            df = pd.concat([df, pd.DataFrame([fila])], ignore_index=True)
            df.to_excel(path, index=False)
            return True
        except PermissionError:
            if intento < 4:
                time.sleep(1)
            else:
                st.error("⚠ El archivo Excel está abierto. Ciérralo e intenta de nuevo.")
                return False
    return False

# ── Folio con SQLite AUTOINCREMENT ────────────────────────────────────────────
_folio_lock = threading.Lock()

def _db_path(cliente_key: str) -> str:
    os.makedirs("data", exist_ok=True)
    return f"data/g2_folios_{cliente_key.upper()}.db"

def _init_db(db_path: str):
    with sqlite3.connect(db_path, timeout=30) as con:
        con.execute("""CREATE TABLE IF NOT EXISTS folios (
            id    INTEGER PRIMARY KEY AUTOINCREMENT,
            razon TEXT NOT NULL,
            ts    TEXT NOT NULL
        )""")
        con.commit()

def folio_nuevo(cliente_key: str, razon_social: str) -> str:
    db        = _db_path(cliente_key)
    _init_db(db)
    razon_key = razon_social.strip().upper()
    with _folio_lock:
        with sqlite3.connect(db, timeout=30, check_same_thread=False) as con:
            cur = con.execute(
                "INSERT INTO folios (razon, ts) VALUES (?, ?)",
                (razon_key, datetime.now().isoformat())
            )
            con.commit()
            n = con.execute(
                "SELECT COUNT(*) FROM folios WHERE razon=? AND id<=?",
                (razon_key, cur.lastrowid)
            ).fetchone()[0]
    return str(n).zfill(3)

def idx_de(lst, val):
    return lst.index(val) if val in lst else 0

def solo_letras(t):
    return re.sub(r"[^A-Za-záéíóúÁÉÍÓÚüÜñÑ\s;]", "", t).upper().strip()

def trabajador_ya_registrado(ap1, ap2, nom, razon, cliente_key) -> bool:
    path = excel_path(cliente_key, razon)
    if not os.path.exists(path):
        return False
    try:
        df = pd.read_excel(path)
        if df.empty or "Nombre" not in df.columns:
            return False
        nombre_nuevo = f"{ap1}; {ap2}; {nom}".strip().upper()
        mask = (
            (df["Nombre"].str.strip().str.upper() == nombre_nuevo) &
            (df["Razón Social"].str.strip().str.upper() == razon.strip().upper())
        )
        return mask.any()
    except Exception:
        return False

def _img_b64(path: str) -> str:
    import base64, mimetypes
    try:
        mime = mimetypes.guess_type(path)[0] or "image/png"
        with open(path, "rb") as f:
            return f"data:{mime};base64,{base64.b64encode(f.read()).decode()}"
    except Exception:
        return ""

# ── CSS ────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@300;400;500;600;700&display=swap');
:root{--v:#4b694e;--v2:#6a9370;--rj:#a20000;--fo:#f1f1f1;--bl:#ffffff;
      --tx:#1a1a1a;--br:#d0d0d0;--am:#c8a600;}
html,body,.stApp{background:var(--fo)!important;font-family:'Montserrat',sans-serif!important;}
#MainMenu,footer,header{visibility:hidden;}
.block-container{padding:1.6rem 2rem!important;max-width:860px!important;margin:auto;}
.folio-box{font-size:.78rem;font-weight:600;color:#888;text-align:right;
           letter-spacing:.06em;text-transform:uppercase;line-height:1.4;}
.folio-num{font-size:1.25rem;font-weight:700;color:var(--tx);}
.ptitulo{font-size:1.35rem;font-weight:700;color:var(--v);border-bottom:2px solid var(--v);
         padding-bottom:.4rem;margin-bottom:1.2rem;}
.pcard{background:var(--bl);border-radius:13px;padding:1.4rem 1.8rem;
       margin-bottom:1rem;box-shadow:0 2px 8px rgba(0,0,0,.07);}
.bv-box{background:var(--bl);border-radius:15px;padding:2rem 2.5rem;
        text-align:center;box-shadow:0 4px 14px rgba(0,0,0,.09);margin-bottom:1.4rem;}
.bv-tit{font-size:1.55rem;font-weight:700;color:var(--v);text-transform:uppercase;
        letter-spacing:.05em;margin-bottom:.3rem;}
.bv-sub{font-size:.88rem;font-weight:500;color:#666;text-transform:uppercase;
        letter-spacing:.1em;margin-bottom:1.2rem;}
.priv{background:#f8f8f8;border:1px solid var(--br);border-radius:9px;
      padding:1rem 1.4rem;text-align:left;font-size:.92rem;color:#555;
      margin-bottom:1.2rem;line-height:1.8;}
.slabel{font-size:.78rem;font-weight:700;color:var(--v);text-transform:uppercase;
        letter-spacing:.1em;margin:.9rem 0 .25rem 0;}
.cfm-box{background:var(--bl);border:2px solid var(--v);border-radius:13px;
         padding:2rem 2.4rem;text-align:center;font-size:1rem;line-height:1.9;
         color:var(--tx);margin-bottom:1.4rem;box-shadow:0 4px 14px rgba(75,105,78,.14);}
.av-box{background:var(--bl);border:2px solid var(--am);border-radius:13px;
        padding:1.7rem 2.2rem;text-align:center;font-size:1rem;color:var(--tx);
        line-height:1.9;margin-bottom:1.4rem;}
.av-tit{font-size:1.05rem;font-weight:700;color:#7a5900;margin-bottom:.5rem;}
.pq-card{background:var(--bl);border-radius:13px;padding:1.7rem 2.1rem;
         margin-bottom:.85rem;box-shadow:0 3px 11px rgba(0,0,0,.08);}
.pq-sec{font-size:.75rem;font-weight:700;color:var(--v);text-transform:uppercase;
        letter-spacing:.1em;margin-bottom:.2rem;}
.pq-dom{font-size:.72rem;color:#888;margin-bottom:.2rem;}
.pq-txt{font-size:1.05rem;font-weight:600;color:var(--tx);line-height:1.6;}
.pq-num{font-size:.7rem;color:#aaa;margin-bottom:.3rem;}
.err-r{color:var(--rj);font-size:.88rem;font-weight:700;margin-top:.4rem;}
.prog-txt{font-size:.82rem;color:#888;text-align:right;margin-bottom:.5rem;font-weight:500;}
.fin-box{background:var(--v);color:#fff;border-radius:17px;padding:3rem 2rem;
         text-align:center;box-shadow:0 6px 22px rgba(75,105,78,.24);}
.fin-tit{font-size:1.35rem;font-weight:700;letter-spacing:.05em;margin-bottom:.6rem;}
.fin-sub{font-size:.96rem;opacity:.84;}
.alerta-viol{background:#fff0f0;border:2px solid #a20000;border-radius:11px;
             padding:1rem 1.4rem;margin:.8rem 0;font-size:.9rem;color:#a20000;
             font-weight:700;line-height:1.7;}
.dup-alert{background:#fff0f0;border:2px solid #a20000;border-radius:11px;
           padding:1rem 1.4rem;margin:.8rem 0;font-size:.88rem;color:#a20000;
           font-weight:600;line-height:1.7;}
.stButton>button{font-family:'Montserrat',sans-serif!important;font-weight:700!important;
  font-size:.93rem!important;letter-spacing:.05em!important;border-radius:9px!important;
  border:none!important;padding:.65rem 1.8rem!important;cursor:pointer!important;
  background:var(--v)!important;color:#fff!important;transition:opacity .2s!important;}
.stButton>button:hover{opacity:.84!important;}
div[role="radiogroup"]>label{background:#fafafa;border:1.5px solid var(--br);
  border-radius:7px;padding:.5rem 1rem!important;margin-bottom:.35rem!important;
  font-size:.96rem!important;font-weight:500;cursor:pointer;
  transition:border-color .14s,background .14s;}
div[role="radiogroup"]>label:hover{border-color:var(--v2);background:#f0f6f0;}
hr.div{border:none;border-top:1.5px solid var(--br);margin:.8rem 0;}
.stTextInput input{text-transform:uppercase!important;font-family:Montserrat,sans-serif!important;
  font-weight:600!important;letter-spacing:.04em!important;}
@keyframes rf-pulse{0%,100%{opacity:1;transform:scale(1);}50%{opacity:.55;transform:scale(.96);}}
@keyframes rf-spin{0%{transform:rotate(0deg);}100%{transform:rotate(360deg);}}
.loading-overlay{position:fixed;inset:0;background:rgba(241,241,241,.92);
  display:flex;flex-direction:column;align-items:center;justify-content:center;z-index:9999;}
.loading-logo{animation:rf-pulse 1.4s ease-in-out infinite;}
.loading-ring{width:48px;height:48px;border:4px solid #d0d0d0;border-top-color:var(--v);
  border-radius:50%;animation:rf-spin .9s linear infinite;margin-top:1.2rem;}
.loading-txt{font-size:.82rem;font-weight:600;color:#4b694e;margin-top:.7rem;letter-spacing:.06em;}
</style>
""", unsafe_allow_html=True)

# ── Estado ─────────────────────────────────────────────────────────────────────
_cliente_def  = _CLIENTE_URL if (_MODO_EMPLEADO and _CLIENTE_URL in CLIENTES) else "FRUCO"
_pantalla_def = "bienvenida" if _MODO_EMPLEADO else "panel"

DEF = dict(
    pantalla=_pantalla_def, cliente_key=_cliente_def,
    razon=CLIENTES[_cliente_def]["opciones"][0],
    areas=["Producción","Administración","Recursos Humanos","Ventas","Operaciones"],
    folio="001",
    ap1="", ap2="", nom="",
    sexo=SEL, edad=SEL, ecivil=SEL,
    estudios=SEL, estatus="Terminada",
    puesto=SEL, area=SEL,
    contrat=SEL, personal=SEL,
    jornada=SEL, rotacion="Sí",
    tpuesto=SEL, exp=SEL,
    preg_idx=0,
    respuestas=[],
    err=False, res=None, modal=None,
    form_v=0,
)
for k, v in DEF.items():
    if k not in st.session_state:
        st.session_state[k] = v
S = st.session_state

if _MODO_EMPLEADO and S.get("pantalla") == "panel":
    S["pantalla"]    = "bienvenida"
    S["cliente_key"] = _cliente_def
    S["razon"]       = CLIENTES[_cliente_def]["opciones"][0]

TODAS_KEYS_FORM = [
    "d_ap1","d_ap2","d_nom",
    "w_sexo","w_edad","w_ecivil","w_estudios","w_estatus",
    "w_puesto","w_area","w_contrat","w_personal",
    "w_jornada","w_rotacion","w_tpuesto","w_exp",
]

def limpiar():
    for k in ["ap1","ap2","nom"]: S[k] = ""
    for k in ["sexo","edad","ecivil","estudios","puesto","area",
              "contrat","personal","jornada","tpuesto","exp"]: S[k] = SEL
    S.estatus = "Terminada"; S.rotacion = "Sí"
    for wk in TODAS_KEYS_FORM:
        if wk in st.session_state: del st.session_state[wk]

def borrar_formulario():
    limpiar()
    S.form_v = S.get("form_v", 0) + 1

def reset_cuestionario():
    S.preg_idx   = 0
    S.respuestas = []
    S.err        = False
    S.res        = None

def header(mostrar_folio=False):
    rf_src  = _img_b64(LOGO_RF)
    cli_src = _img_b64(CLIENTES[S.cliente_key]["logo"])
    folio_html = (
        f'<div class="folio-box" style="margin-left:auto;text-align:right;">'
        f'NO. DE CUESTIONARIO<br>'
        f'<span class="folio-num">{S.folio}</span></div>'
    ) if (mostrar_folio and not _MODO_EMPLEADO) else ""
    rf_tag  = (f'<img src="{rf_src}" alt="RFRANYUTTI" '
               f'style="height:52px;max-width:120px;width:auto;object-fit:contain;">'
               ) if rf_src else "<b>RFRANYUTTI</b>"
    cli_tag = (f'<img src="{cli_src}" alt="{S.cliente_key}" '
               f'style="height:52px;max-width:120px;width:auto;object-fit:contain;">'
               ) if cli_src else f"<b>{S.cliente_key}</b>"
    st.markdown(f"""
    <div style="display:flex;flex-wrap:wrap;align-items:center;justify-content:space-between;
                gap:.5rem;padding:.5rem 0 .6rem 0;margin-bottom:.2rem;">
        <div style="display:flex;align-items:center;gap:.8rem;flex-wrap:wrap;">
            {rf_tag}
            <div style="width:1px;height:44px;background:#ccc;flex-shrink:0;"></div>
            {cli_tag}
        </div>
        <div style="flex-shrink:0;">{folio_html}</div>
    </div>
    <hr class="div">
    """, unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# PANEL DEL OPERATIVO
# ══════════════════════════════════════════════════════════════════════════════
if S.pantalla == "panel":
    st.markdown('<div class="ptitulo">⚙ PANEL DEL OPERATIVO · NOM-035-STPS-2018 · GUÍA II</div>',
                unsafe_allow_html=True)

    if S.modal == "borrar":
        st.warning("⚠ **LOS REGISTROS SE ELIMINARÁN PERMANENTEMENTE.**\n\n¿Desea continuar?")
        c1, c2 = st.columns(2)
        with c1:
            if st.button("CONTINUAR — BORRAR TODO"):
                p = excel_path(S.cliente_key, S.razon)
                if os.path.exists(p): os.remove(p)
                S.modal = None; st.success("✓ Registros eliminados."); st.rerun()
        with c2:
            if st.button("CANCELAR"): S.modal = None; st.rerun()
        st.stop()

    if S.modal == "terminar":
        st.info("ℹ **¿Desea cerrar la sesión de registros?**")
        c1, c2 = st.columns(2)
        with c1:
            if st.button("CONTINUAR"):
                S.modal = None; st.success("✓ Sesión cerrada."); st.rerun()
        with c2:
            if st.button("CANCELAR"): S.modal = None; st.rerun()
        st.stop()

    st.markdown('<div class="pcard">', unsafe_allow_html=True)
    st.markdown("**1. Selección de cliente**")
    ck = st.selectbox("Cliente", list(CLIENTES.keys()),
                      index=list(CLIENTES.keys()).index(S.cliente_key))
    S.cliente_key = ck
    inf = CLIENTES[ck]
    S.razon = st.selectbox("Razón social", inf["opciones"])
    try: st.image(inf["logo"], width=120)
    except: st.caption(f"Logo no encontrado: {inf['logo']}")
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="pcard">', unsafe_allow_html=True)
    st.markdown("**2. Departamentos / Áreas**")
    st.caption("Una área por línea.")
    at = st.text_area("", "\n".join(S.areas), height=120, label_visibility="collapsed")
    S.areas = [a.strip() for a in at.split("\n") if a.strip()]
    st.caption(f"{len(S.areas)} área(s) configuradas.")
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="pcard">', unsafe_allow_html=True)
    st.markdown("**3. Guía a aplicar**")
    st.info("**GUÍA II** — Identificación y Análisis de los Factores de Riesgo Psicosocial "
            "y Evaluación del Entorno Organizacional\n\n"
            "Para centros de trabajo con **16 a 50 trabajadores** · **46 preguntas**")
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("**4. Acciones**")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        _p_raw = excel_path(S.cliente_key, S.razon)
        if os.path.exists(_p_raw):
            with open(_p_raw, "rb") as _f:
                st.download_button(
                    label="⬇️ DATOS (Excel)",
                    data=_f.read(),
                    file_name=os.path.basename(_p_raw),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
        else:
            st.button("⬇️ DATOS (Excel)", disabled=True, use_container_width=True,
                      help=f"Sin datos aún para {S.cliente_key}.")
    with c2:
        if st.button("ABRIR WORD", use_container_width=True):
            st.info("Usa el botón INFORME WORD COMPLETO en la sección 5.")
    with c3:
        if st.button("BORRAR REGISTROS", use_container_width=True):
            S.modal = "borrar"; st.rerun()
    with c4:
        if st.button("TERMINAR REGISTROS", use_container_width=True):
            S.modal = "terminar"; st.rerun()

    # ── Generar Reportes ──────────────────────────────────────────────────────
    st.markdown('<hr class="div">', unsafe_allow_html=True)
    st.markdown("**5. Generar Reportes**")
    _p_rep   = excel_path(S.cliente_key, S.razon)
    _hay_dat = os.path.exists(_p_rep)
    cr1, cr2 = st.columns(2)
    with cr1:
        if st.button("📊  EXCEL MEJORADO + GRÁFICAS", use_container_width=True,
                     disabled=not _hay_dat):
            with st.spinner("Generando Excel con análisis completo..."):
                try:
                    from generar_reporte_g2 import generar_excel_g2
                    out = generar_excel_g2(_p_rep, S.cliente_key, S.razon)
                    if out:
                        with open(out, "rb") as f:
                            st.download_button("⬇️  DESCARGAR EXCEL", data=f.read(),
                                file_name=os.path.basename(out),
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True)
                    else:
                        st.warning("Sin datos suficientes.")
                except Exception as e:
                    st.error(f"Error: {e}")
    with cr2:
        if st.button("📄  INFORME WORD COMPLETO", use_container_width=True,
                     disabled=not _hay_dat):
            with st.spinner("Generando informe Word..."):
                try:
                    from generar_reporte_g2 import generar_word_g2
                    out = generar_word_g2(_p_rep, S.cliente_key, S.razon,
                                          logo_rf=LOGO_RF,
                                          logo_cliente=CLIENTES[S.cliente_key]["logo"])
                    if out:
                        with open(out, "rb") as f:
                            st.download_button("⬇️  DESCARGAR WORD", data=f.read(),
                                file_name=os.path.basename(out),
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                use_container_width=True)
                    else:
                        st.warning("Sin datos suficientes.")
                except Exception as e:
                    st.error(f"Error: {e}")
    if not _hay_dat:
        st.caption("⚠ Registra al menos un cuestionario antes de generar reportes.")

    # ── Link para empleados ───────────────────────────────────────────────────
    st.markdown('<hr class="div">', unsafe_allow_html=True)
    st.markdown("**6. Link para empleados**")
    _base = st.text_input("URL base de tu app", value="http://localhost:8501",
                          help="Pega tu URL de Streamlit Cloud aquí.")
    _link = f"{_base.rstrip('/')}/?cliente={S.cliente_key}"
    st.markdown(f"""
    <div style="background:#f0f6f0;border:2px solid #4b694e;border-radius:10px;
                padding:1rem 1.4rem;margin:.4rem 0 .8rem 0;font-family:Montserrat,sans-serif;">
        <div style="font-size:.72rem;font-weight:700;color:#4b694e;text-transform:uppercase;
                    letter-spacing:.08em;margin-bottom:.35rem;">
            🔗 Link empleados — {S.cliente_key}
        </div>
        <div style="font-size:.9rem;font-weight:600;word-break:break-all;background:#fff;
                    border-radius:6px;padding:.45rem .8rem;border:1px solid #d0d0d0;">
            {_link}
        </div>
        <div style="font-size:.75rem;color:#666;margin-top:.4rem;">
            Comparte por WhatsApp · Sesiones independientes · Folio automático
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.code(_link, language=None)

    st.markdown('<hr class="div">', unsafe_allow_html=True)
    st.markdown("**7. Cuestionario modo local**")
    if st.button("🟢  INICIAR CUESTIONARIO PARA EMPLEADO", use_container_width=True):
        S.folio = folio_nuevo(S.cliente_key, S.razon)
        limpiar(); reset_cuestionario()
        S.pantalla = "bienvenida"; st.rerun()

# ══════════════════════════════════════════════════════════════════════════════
# BIENVENIDA
# ══════════════════════════════════════════════════════════════════════════════
elif S.pantalla == "bienvenida":
    header(False)
    st.markdown("""
    <div class="bv-box">
      <div class="bv-tit">NOM-035-STPS-2018</div>
      <div class="bv-sub">Cuestionario para identificar y analizar los factores<br>
        de riesgo psicosocial y evaluar el entorno organizacional</div>
      <p style="font-size:.83rem;color:#444;line-height:1.8;text-align:left;">
        Este cuestionario tiene como objetivo identificar y analizar los posibles factores
        de riesgo psicosocial en su centro de trabajo, así como evaluar el entorno
        organizacional. Sus respuestas son confidenciales y se utilizarán exclusivamente
        para mejorar las condiciones de trabajo. No existen respuestas correctas o incorrectas.
        Por favor responda con sinceridad basándose en su experiencia personal.
      </p>
      <div class="bv-sub" style="margin-top:.9rem;">Cláusula de Privacidad</div>
      <div class="priv">
        La información proporcionada será tratada con estricta confidencialidad. Los datos y
        resultados se utilizarán exclusivamente para fines internos de diagnóstico y mejora del
        ambiente laboral, conforme a la normativa vigente en materia de protección de datos
        personales.
      </div>
    </div>
    """, unsafe_allow_html=True)
    acepta = st.checkbox("He leído y acepto la cláusula de privacidad *")
    if acepta:
        if st.button("CONTINUAR →", use_container_width=True):
            if not _MODO_EMPLEADO:
                S.pantalla = "datos"
            else:
                S.folio    = folio_nuevo(S.cliente_key, S.razon)
                limpiar(); reset_cuestionario()
                S.pantalla = "datos"
            st.rerun()

# ══════════════════════════════════════════════════════════════════════════════
# DATOS GENERALES
# ══════════════════════════════════════════════════════════════════════════════
elif S.pantalla == "datos":
    header(True)
    st.markdown('<div class="slabel">Información General del Trabajador</div>',
                unsafe_allow_html=True)
    err = []

    st.markdown('<div class="slabel">Nombre</div>', unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)
    with c1: ap1 = st.text_input("PRIMER APELLIDO", value=S.ap1, key=f"d_ap1_{S.form_v}",
                                  help="Solo letras.")
    with c2: ap2 = st.text_input("SEGUNDO APELLIDO", value=S.ap2, key=f"d_ap2_{S.form_v}",
                                  help="Solo letras.")
    with c3: nom = st.text_input("NOMBRE(S)", value=S.nom, key=f"d_nom_{S.form_v}",
                                  help="Solo letras.")

    _hay_err = False
    for _val, _lbl in [(ap1,"Primer Apellido"),(ap2,"Segundo Apellido"),(nom,"Nombre(s)")]:
        if _val and re.search(r"[^A-Za-záéíóúÁÉÍÓÚüÜñÑ\s]", _val):
            st.markdown(f'<p style="color:#a20000;font-size:.78rem;font-weight:700;margin:0;">'
                        f'⚠ {_lbl}: solo letras.</p>', unsafe_allow_html=True)
            _hay_err = True

    c1, c2, c3 = st.columns(3)
    with c1: sexo   = st.selectbox("SEXO", OPC_SEXO, key=f"w_sexo_{S.form_v}",
                                    index=idx_de(OPC_SEXO, S.sexo))
    with c2: edad   = st.selectbox("EDAD (años)", OPC_EDAD, key=f"w_edad_{S.form_v}",
                                    index=idx_de(OPC_EDAD, S.edad))
    with c3: ecivil = st.selectbox("ESTADO CIVIL", OPC_ECIVIL, key=f"w_ecivil_{S.form_v}",
                                    index=idx_de(OPC_ECIVIL, S.ecivil))

    st.markdown('<div class="slabel">Nivel de Estudios</div>', unsafe_allow_html=True)
    c1, c2 = st.columns([2, 1])
    with c1: estudios = st.selectbox("NIVEL", OPC_ESTUD, key=f"w_estudios_{S.form_v}",
                                      index=idx_de(OPC_ESTUD, S.estudios))
    with c2:
        if estudios not in [SEL, "Sin formación"]:
            estatus = st.radio("", ["Terminada","Incompleta"], horizontal=True,
                               key=f"w_estatus_{S.form_v}",
                               index=0 if S.estatus == "Terminada" else 1)
        else:
            estatus = "N/A"
            k_est = f"w_estatus_{S.form_v}"
            if k_est in st.session_state: del st.session_state[k_est]

    c1, c2 = st.columns(2)
    with c1: puesto = st.selectbox("PUESTO", OPC_PUESTO, key=f"w_puesto_{S.form_v}",
                                    index=idx_de(OPC_PUESTO, S.puesto))
    with c2:
        aopts = [SEL] + S.areas
        area  = st.selectbox("DEPARTAMENTO / ÁREA", aopts, key=f"w_area_{S.form_v}",
                              index=idx_de(aopts, S.area))

    c1, c2 = st.columns(2)
    with c1: contrat  = st.selectbox("TIPO DE CONTRATACIÓN", OPC_CONTRAT,
                                      key=f"w_contrat_{S.form_v}",
                                      index=idx_de(OPC_CONTRAT, S.contrat))
    with c2: personal = st.selectbox("TIPO DE PERSONAL", OPC_PERSONAL,
                                      key=f"w_personal_{S.form_v}",
                                      index=idx_de(OPC_PERSONAL, S.personal))

    c1, c2 = st.columns([3, 1])
    with c1: jornada = st.selectbox("TIPO DE JORNADA DE TRABAJO", OPC_JORNADA,
                                     key=f"w_jornada_{S.form_v}",
                                     index=idx_de(OPC_JORNADA, S.jornada))
    with c2:
        st.markdown('<div class="slabel">Rotación de Turnos</div>', unsafe_allow_html=True)
        rotacion = st.radio("", ["Sí","No"], horizontal=True, key=f"w_rotacion_{S.form_v}",
                            index=0 if S.rotacion == "Sí" else 1)

    c1, c2 = st.columns(2)
    with c1: tpuesto = st.selectbox("TIEMPO EN EL PUESTO ACTUAL", OPC_TPUESTO,
                                     key=f"w_tpuesto_{S.form_v}",
                                     index=idx_de(OPC_TPUESTO, S.tpuesto))
    with c2: exp     = st.selectbox("TIEMPO DE EXPERIENCIA LABORAL", OPC_EXP,
                                     key=f"w_exp_{S.form_v}",
                                     index=idx_de(OPC_EXP, S.exp))

    st.markdown("<br>", unsafe_allow_html=True)
    c1, c2 = st.columns([3, 1])
    with c1: iniciar = st.button("INICIAR CUESTIONARIO", use_container_width=True)
    with c2:
        if st.button("BORRAR DATOS", use_container_width=True):
            borrar_formulario(); st.rerun()

    if iniciar:
        a1 = solo_letras(ap1.strip())
        a2 = solo_letras(ap2.strip())
        nn = solo_letras(nom.strip())
        if _hay_err:
            err.append("Corrige los campos de nombre: solo letras.")
        else:
            if not a1: err.append("El Primer Apellido es obligatorio.")
            if not a2: err.append("El Segundo Apellido es obligatorio.")
            if not nn: err.append("El campo Nombre(s) es obligatorio.")
        if sexo     == SEL: err.append("Selecciona el Sexo.")
        if edad     == SEL: err.append("Selecciona la Edad.")
        if ecivil   == SEL: err.append("Selecciona el Estado Civil.")
        if estudios == SEL: err.append("Selecciona el Nivel de Estudios.")
        if puesto   == SEL: err.append("Selecciona el Puesto.")
        if area     == SEL: err.append("Selecciona el Área.")
        if contrat  == SEL: err.append("Selecciona el Tipo de Contratación.")
        if personal == SEL: err.append("Selecciona el Tipo de Personal.")
        if jornada  == SEL: err.append("Selecciona el Tipo de Jornada.")
        if tpuesto  == SEL: err.append("Selecciona el Tiempo en el Puesto Actual.")
        if exp      == SEL: err.append("Selecciona el Tiempo de Experiencia Laboral.")
        if err:
            for e in err:
                st.markdown(f'<p style="color:#a20000;font-size:.82rem;font-weight:600;">'
                            f'⚠ {e}</p>', unsafe_allow_html=True)
        else:
            if trabajador_ya_registrado(a1, a2, nn, S.razon, S.cliente_key):
                st.markdown(f"""
                <div class="dup-alert">
                    ⛔ Registro duplicado — {a1} {a2}, {nn} ya tiene cuestionario
                    registrado en {S.razon}.
                </div>""", unsafe_allow_html=True)
            else:
                S.ap1=a1; S.ap2=a2; S.nom=nn
                S.sexo=sexo; S.edad=edad; S.ecivil=ecivil
                S.estudios=estudios; S.estatus=estatus
                S.puesto=puesto; S.area=area
                S.contrat=contrat; S.personal=personal
                S.jornada=jornada; S.rotacion=rotacion
                S.tpuesto=tpuesto; S.exp=exp
                S.pantalla = "confirmar"; st.rerun()

# ══════════════════════════════════════════════════════════════════════════════
# CONFIRMACIÓN
# ══════════════════════════════════════════════════════════════════════════════
elif S.pantalla == "confirmar":
    header(True)
    st.markdown("""
    <div class="cfm-box">
      <p style="font-size:.97rem;font-weight:700;text-transform:uppercase;color:#4b694e;">
        Confirmación de Datos</p>
      <p>Al pulsar <strong>"ACEPTAR"</strong> usted declara que los datos ingresados
      en el apartado de información general son correctos y reflejan su situación actual
      en la empresa.</p>
    </div>
    """, unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        if st.button("ACEPTAR", use_container_width=True):
            S.pantalla = "aviso"; st.rerun()
    with c2:
        if st.button("← REGRESAR Y EDITAR", use_container_width=True):
            S.pantalla = "datos"; st.rerun()

# ══════════════════════════════════════════════════════════════════════════════
# AVISO DE NO RETROCESO
# ══════════════════════════════════════════════════════════════════════════════
elif S.pantalla == "aviso":
    header(True)
    st.markdown("""
    <div class="av-box">
      <div class="av-tit">⚠ ANTES DE COMENZAR — LEA CON ATENCIÓN</div>
      <p>Este cuestionario consta de <strong>46 preguntas</strong>. Una vez que avance
      a la siguiente pregunta, <strong>no podrá regresar a la anterior.</strong><br>
      Por favor responda con sinceridad basándose en su experiencia real en el trabajo.</p>
    </div>
    """, unsafe_allow_html=True)
    if st.button("COMENZAR CUESTIONARIO →", use_container_width=True):
        reset_cuestionario()
        S.pantalla = "preguntas"; st.rerun()

# ══════════════════════════════════════════════════════════════════════════════
# PREGUNTAS — 46 ítems, una a la vez, sin retroceso
# ══════════════════════════════════════════════════════════════════════════════
elif S.pantalla == "preguntas":
    header(True)

    total_pregs = len(PREGUNTAS_G2)
    idx         = S.preg_idx

    if idx >= total_pregs:
        S.pantalla = "fin"; st.rerun()

    preg = PREGUNTAS_G2[idx]

    st.markdown(
        f'<div class="prog-txt">{preg["cat_nombre"]} · '
        f'Pregunta {idx + 1} de {total_pregs}</div>',
        unsafe_allow_html=True)
    st.progress(idx / total_pregs)

    st.markdown(f"""
    <div class="pq-card">
      <div class="pq-num">Pregunta {idx + 1} / {total_pregs}</div>
      <div class="pq-sec">{preg["cat_nombre"]}</div>
      <div class="pq-dom">Dominio: {preg["dom_nombre"]}</div>
      <div class="pq-txt">{preg["texto"]}</div>
    </div>
    """, unsafe_allow_html=True)

    resp = st.radio(
        "",
        OPC_RESP,
        index=None,
        horizontal=False,
        key=f"q_{idx}",
        label_visibility="collapsed",
    )

    if S.err:
        st.markdown('<p class="err-r">⚠ SELECCIONA UNA RESPUESTA ANTES DE CONTINUAR</p>',
                    unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    if st.button("SIGUIENTE →", use_container_width=True, key=f"sig_{idx}"):
        if resp is None:
            S.err = True; st.rerun()
        else:
            S.err        = False
            S.respuestas = S.respuestas + [resp]
            S.preg_idx   = idx + 1

            if S.preg_idx >= total_pregs:
                # Calcular resultado
                resultado = calcular_puntaje(S.respuestas)
                S.res     = resultado
                guardar(dict(
                    folio=S.folio, cliente=S.cliente_key, razon=S.razon,
                    nombre=f"{S.ap1}; {S.ap2}; {S.nom}",
                    sexo=S.sexo, edad=S.edad, ecivil=S.ecivil,
                    estudios=S.estudios, estatus=S.estatus,
                    puesto=S.puesto, area=S.area,
                    contrat=S.contrat, personal=S.personal,
                    jornada=S.jornada, rotacion=S.rotacion,
                    tpuesto=S.tpuesto, exp=S.exp,
                    respuestas=S.respuestas,
                    resultado=resultado,
                ))
                S.pantalla = "fin"
            st.rerun()

# ══════════════════════════════════════════════════════════════════════════════
# FIN
# ══════════════════════════════════════════════════════════════════════════════
elif S.pantalla == "fin":
    header(False)
    st.markdown("""
    <div class="fin-box">
      <div class="fin-tit">✓ CUESTIONARIO FINALIZADO CORRECTAMENTE</div>
      <div class="fin-sub">AGRADECEMOS TU PARTICIPACIÓN</div>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)

    if _MODO_EMPLEADO:
        st.markdown("""
        <div style="text-align:center;font-size:.92rem;color:#666;
                    font-family:Montserrat,sans-serif;padding:1rem 0;">
            Puedes cerrar esta ventana.
        </div>
        """, unsafe_allow_html=True)
    else:
        # Solo el operativo ve el resultado
        if S.res:
            r   = S.res
            col = r.get("color", "#4b694e")
            bg  = r.get("bg",    "#D6E4D8")
            with st.expander("Ver resultado — uso interno operativo"):
                st.markdown(f"""
                <div style="background:{col};color:#fff;border-radius:13px;
                            padding:1.5rem 2rem;text-align:center;margin-bottom:1rem;">
                    <div style="font-size:.75rem;opacity:.8;text-transform:uppercase;
                                letter-spacing:.1em;">Resultado · Guía II · NOM-035</div>
                    <div style="font-size:1.4rem;font-weight:700;margin-top:.4rem;">
                        {r['nivel']}</div>
                    <div style="font-size:.9rem;opacity:.85;margin-top:.3rem;">
                        Puntaje total: {r['puntaje_total']} puntos</div>
                </div>
                """, unsafe_allow_html=True)
                if r.get("alerta_violencia"):
                    st.markdown("""
                    <div class="alerta-viol">
                        🚨 ALERTA — VIOLENCIA LABORAL DETECTADA<br>
                        Este trabajador reportó conductas de violencia laboral (ítems 44-46).
                        Se requiere intervención inmediata conforme al numeral 8.4.c
                        de la NOM-035-STPS-2018.
                    </div>
                    """, unsafe_allow_html=True)

        c1, c2 = st.columns(2)
        with c1:
            if st.button("REGISTRAR OTRO EMPLEADO", use_container_width=True):
                S.pantalla = "panel"; st.rerun()
        with c2:
            if st.button("VOLVER AL PANEL", use_container_width=True):
                S.pantalla = "panel"; st.rerun()
