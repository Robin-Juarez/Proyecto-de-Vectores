#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PROYECTO VECTORES - Álgebra Lineal  (100% Python)
=================================================

QUÉ HACE ESTE PROGRAMA
----------------------
Abre una ventana de escritorio donde se escriben de 2 a 5 vectores (cada uno de 2 a 5
componentes), se elige una operación y el programa muestra:
  * el resultado de la operación,
  * SIEMPRE una gráfica:
        - 2 componentes  -> gráfica 2D con flechas,
        - 3 componentes  -> gráfica 3D con flechas (se gira con el mouse),
        - 4 o 5 componentes (o componentes distintos entre vectores) -> gráfica de barras
          con el valor de cada componente (no se pueden dibujar flechas en 4D/5D),
  * el resultado de TODAS las operaciones,
  * el procedimiento completo, paso a paso, en lenguaje sencillo.

OPERACIONES: suma, resta, multiplicación por escalar, magnitud y vector unitario.

REGLAS (se mantienen siempre)
-----------------------------
  - Solo NÚMEROS ENTEROS: sin letras, sin signo negativo y sin punto decimal.
    (El mínimo permitido está en la constante VALOR_MINIMO.)
  - Para SUMAR o RESTAR todos los vectores deben tener el mismo número de componentes.

BONO: CARGAR VECTORES DESDE UN ARCHIVO
--------------------------------------
Botón "Cargar vectores desde archivo". Formatos: .txt .csv .pdf .docx .xlsx .xlsm
Una línea (o fila) = un vector. Los números van separados por coma, punto y coma o espacio.
Se acepta una etiqueta opcional al inicio. Ejemplo de archivo de texto:
        A = (3, 2, 1)
        B: 1 4 2
        4, 5, 6
Las líneas sin ningún número (títulos) se ignoran. Si una línea tiene letras, negativos
o decimales, el programa avisa en qué línea está el error y NO carga nada.

LIBRERÍAS
---------
  Obligatorias:  numpy, matplotlib            (tkinter viene con Python)
  Opcionales (solo para leer ese tipo de archivo):
        pypdf (PDF)   python-docx (Word)   openpyxl (Excel)

INSTALACIÓN Y USO
-----------------
    pip install numpy matplotlib pypdf python-docx openpyxl
    python Taller_de_Vectores_RJ_python.py

COLORES: A verde, B rojo, C ámbar, D morado, E turquesa, resultado azul.

POSIBLES CAMBIOS (sin alterar el funcionamiento del programa)
-------------------------------------------------------------
  1. Cálculos: numpy -> listas de Python + math.sqrt. Basta reescribir sumar(), restar(),
     por_escalar(), magnitud() y unitario() manteniendo lo que reciben y devuelven.
  2. Gráficas: matplotlib -> plotly o pyqtgraph. Solo cambian dibujar_2d(), dibujar_3d()
     y dibujar_barras(); los datos (construir_segmentos) quedan iguales.
  3. Barras (4D/5D) -> gráfica de "coordenadas paralelas" (una línea por vector):
     solo se reemplaza dibujar_barras().
  4. Ventana: tkinter -> PyQt/customtkinter: solo se cambia la clase App.
  5. Lectura de archivos: pypdf -> pdfplumber; python-docx -> docx2txt; openpyxl ->
     pandas.read_excel. Solo cambia extraer_lineas(); parsear_vectores() sigue igual.
  6. Permitir el cero: poner VALOR_MINIMO = 0 (el vector unitario ya evita dividir entre 0).
  7. Más vectores/componentes: cambiar MAX_VECTORES / MAX_COMP y agregar nombres, ejes y colores.
  8. parsear_vectores(): las expresiones regulares (re) se pueden cambiar por split() e int().
"""

# ---------------------------------------------------------------------------
# IMPORTACIÓN DE LIBRERÍAS
# ---------------------------------------------------------------------------
import os                          # nombres y extensiones de archivos
import random                      # vectores al azar
import re                          # expresiones regulares para leer los archivos
import sys                         # terminar el programa si falta una librería obligatoria
import tkinter as tk               # ventana y widgets básicos
from tkinter import ttk, filedialog, messagebox   # widgets modernos, selector de archivos, ventanas de aviso

try:
    import numpy as np                                     # cálculos con vectores
    import matplotlib
    matplotlib.use("TkAgg")                                # dibujar las gráficas dentro de tkinter
    from matplotlib.figure import Figure                   # lienzo de la gráfica
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
    from mpl_toolkits.mplot3d import Axes3D                # noqa: F401  (activa los ejes 3D)
except ImportError:
    print("Faltan librerías. Instálalas con:  pip install numpy matplotlib")
    sys.exit(1)

# Librerías OPCIONALES: si no están, el programa funciona igual y solo avisa
# cuando intentes abrir ese tipo de archivo.
try:
    from pypdf import PdfReader                            # leer PDF
except ImportError:
    PdfReader = None
try:
    import docx                                            # leer Word (.docx)
except ImportError:
    docx = None
try:
    import openpyxl                                        # leer Excel (.xlsx)
except ImportError:
    openpyxl = None

# ---------------------------------------------------------------------------
# CONSTANTES
# ---------------------------------------------------------------------------
NOMBRES = ["A", "B", "C", "D", "E"]            # nombres de los vectores
EJES = ["x", "y", "z", "w", "u"]               # nombres de las componentes
SUB = "₁₂₃₄₅"                                  # subíndices: a₁, a₂, ...
COLORES = {"A": "#0e9f6e", "B": "#e11d48", "C": "#d97706", "D": "#7c3aed", "E": "#0891b2", "R": "#1d6fe0"}  # colores de los vectores y del resultado
ROJO_AVISO = "#c0263b"                         # color de los mensajes de error
MIN_VECTORES, MAX_VECTORES = 2, 5              # cantidad de vectores permitida
MIN_COMP, MAX_COMP = 2, 5                      # componentes por vector permitidas
VALOR_MINIMO = 0                               # valor mínimo de una componente (1 = naturales; 0 = permite el cero)
MAX_DIGITOS = 5                                # máximo de dígitos por número
MSG_ENTERO = (f"Solo se permiten números enteros mayores o iguales a {VALOR_MINIMO} "
              "(sin letras, signos negativos ni punto decimal).")

# Menú de operaciones: (identificador, texto del botón)
OPERACIONES = [("suma", "Suma  A + B + …"), ("resta", "Resta  A − B − …"), ("escalar", "Escalar  k · A"),
               ("magnitud", "Magnitud  |A|"), ("unitario", "Vector unitario  Â")]


# ===========================================================================
# 1) VALIDACIÓN Y FORMATO
# ===========================================================================
def validar_tecla(texto_propuesto):
    """
    Se ejecuta cada vez que el usuario escribe o pega algo en una caja numérica.
    Recibe el texto que QUEDARÍA en la caja y solo lo acepta si está vacío o son únicamente
    dígitos 0-9 (máximo MAX_DIGITOS). Así se bloquean letras, negativos, espacios y el punto decimal.
    """
    return texto_propuesto == "" or (texto_propuesto.isascii() and texto_propuesto.isdigit()
                                     and len(texto_propuesto) <= MAX_DIGITOS)


def es_natural(texto):
    """Dice si el texto es un entero válido (solo dígitos y >= VALOR_MINIMO). Vacío NO es válido."""
    return texto.isascii() and texto.isdigit() and int(texto) >= VALOR_MINIMO



def fmt(x):
    """Formato de número: máximo 4 decimales y sin ceros sobrantes (3.6056, no 3.605551)."""
    r = round(float(x), 4)
    if r == 0:
        r = 0.0                                # evita mostrar "-0"
    return str(int(r)) if r == int(r) else str(r)


def fmt_vec(v):
    """Formato de vector: [3, 2] -> "(3, 2)"."""
    return "(" + ", ".join(fmt(x) for x in v) + ")"


# ===========================================================================
# 2) OPERACIONES MATEMÁTICAS  (cada vector es un arreglo de numpy)
# ===========================================================================
def sumar(vectores):
    """Suma de vectores: (a1+b1, a2+b2, ...). np.sum con axis=0 suma posiciones de igual índice."""
    return np.sum(vectores, axis=0)


def restar(vectores):
    """Resta de vectores: A − B − C ... Parte del primero y le resta los demás posición por posición."""
    resultado = vectores[0].copy()
    for v in vectores[1:]:
        resultado = resultado - v
    return resultado


def por_escalar(k, v):
    """Multiplicación por escalar: (k·a1, k·a2, ...). Estira el vector k veces."""
    return k * v


def magnitud(v):
    """Magnitud (longitud): raíz de (a1² + a2² + ...). Es el teorema de Pitágoras generalizado."""
    return float(np.sqrt(np.sum(v ** 2)))


def unitario(v):
    """Vector unitario: el vector dividido entre su magnitud (mide 1). Si la magnitud es 0 devuelve ceros."""
    m = magnitud(v)
    return v / m if m > 0 else np.zeros_like(v)


def calcular(datos):
    """
    Calcula TODAS las operaciones y devuelve un diccionario.
    Si los vectores no tienen el mismo número de componentes, suma y resta quedan en None.
    Escalar, magnitud y unitario se aplican al vector objetivo (A, B, C, D o E).
    """
    vs = datos["vs"]
    objetivo = vs[datos["tgt"]]
    iguales = len(set(datos["dims"])) == 1
    return {
        "suma": sumar(vs) if iguales else None,
        "resta": restar(vs) if iguales else None,
        "escalar": por_escalar(datos["k"], objetivo),
        "magnitud": magnitud(objetivo),
        "unitario": unitario(objetivo),
    }


def dim_grafica(datos, op):
    """
    Número de componentes con el que se opera/dibuja (0 si no se puede operar).
    Suma/resta: el número común de componentes (0 si son distintos).
    Escalar/magnitud/unitario: los componentes del vector objetivo.
    """
    if op in ("suma", "resta"):
        return datos["dims"][0] if len(set(datos["dims"])) == 1 else 0
    return datos["dims"][datos["tgt"]]


def mensaje_dimensiones(datos, op):
    """Explica por qué no se puede sumar o restar (distinto número de componentes)."""
    detalle = ", ".join(f"{NOMBRES[i]} tiene {datos['dims'][i]}" for i in range(datos["n"]))
    verbo = "sumar" if op == "suma" else "restar"
    return f"Para {verbo}, todos los vectores deben tener el mismo número de componentes ({detalle})."


# ===========================================================================
# 3) TEXTOS: RESULTADO, CUADRITOS Y PROCEDIMIENTO PASO A PASO
# ===========================================================================
def texto_resultado(datos, res, op):
    """Texto del resultado principal, por ejemplo "A + B = (4, 6)"."""
    if dim_grafica(datos, op) == 0 and op in ("suma", "resta"):
        return "No se puede " + ("sumar" if op == "suma" else "restar") + ": los vectores tienen distinto número de componentes."
    t = NOMBRES[datos["tgt"]]
    todos = [NOMBRES[i] for i in range(datos["n"])]
    if op == "suma":
        return f"{' + '.join(todos)} = {fmt_vec(res['suma'])}"
    if op == "resta":
        return f"{' − '.join(todos)} = {fmt_vec(res['resta'])}"
    if op == "escalar":
        return f"{datos['k']}·{t} = {fmt_vec(res['escalar'])}"
    if op == "magnitud":
        return f"|{t}| = {fmt(res['magnitud'])}"
    return f"{t}\u0302 = {fmt_vec(res['unitario'])}"


def texto_cuadritos(datos, res):
    """Resultado de TODAS las operaciones (pestaña "Todas las operaciones")."""
    n, vs, k = datos["n"], datos["vs"], datos["k"]
    todos = [NOMBRES[i] for i in range(n)]
    sin_def = "No definida: componentes distintos"
    lineas = [
        "SUMA\n    " + " + ".join(todos) + " = " + (fmt_vec(res["suma"]) if res["suma"] is not None else sin_def),
        "RESTA\n    " + " − ".join(todos) + " = " + (fmt_vec(res["resta"]) if res["resta"] is not None else sin_def),
        f"ESCALAR (k = {k})\n" + "\n".join(f"    k·{NOMBRES[i]} = {fmt_vec(k * vs[i])}" for i in range(n)),
        "MAGNITUD\n" + "\n".join(f"    |{NOMBRES[i]}| = {fmt(magnitud(vs[i]))}" for i in range(n)),
        "VECTOR UNITARIO\n" + "\n".join(f"    {NOMBRES[i]}\u0302 = {fmt_vec(unitario(vs[i]))}" for i in range(n)),
    ]
    return "\n\n".join(lineas)


def construir_pasos(datos, res, op):
    """
    Construye el PROCEDIMIENTO paso a paso. Devuelve una lista de pasos; cada paso es
    (título, [líneas de cálculo], explicación o None). Muestra fórmula, sustitución con
    los números reales, cálculo posición por posición y una explicación sencilla.
    """
    n, vs, t = datos["n"], datos["vs"], datos["tgt"]
    I = list(range(n))
    pd = dim_grafica(datos, op)
    J = list(range(pd))
    T = vs[t]
    low = lambda i: NOMBRES[i].lower()
    nom_t = NOMBRES[t]

    direcciones = f" ({', '.join(EJES[:pd])})" if pd else ""
    paso1 = ("Escribe tus vectores", [f"{NOMBRES[i]} = {fmt_vec(vs[i])}" for i in I],
             "Un vector es una flecha. Sus números (componentes) dicen cuánto avanza en cada "
             f"dirección{direcciones} desde el origen.")

    if pd == 0 and op in ("suma", "resta"):
        verbo = "sumar" if op == "suma" else "restar"
        return [paso1, (f"No se puede {verbo}", [mensaje_dimensiones(datos, op)],
                        "Sumar o restar se hace posición por posición. Si un vector tiene más posiciones "
                        "que otro, a algunas no les toca pareja. Cambia el número de componentes para que sean iguales.")]

    def grafica(texto_flechas):
        """Último paso: explica la gráfica según el tipo (2D, 3D o barras)."""
        if pd == 2:
            return ("Míralo en la gráfica", [], texto_flechas)
        if pd == 3:
            return ("Míralo en la gráfica 3D", [], texto_flechas + " Puedes girar la gráfica arrastrando con el mouse.")
        return ("Míralo en la gráfica de barras", [],
                "Con 4 o 5 componentes no se pueden dibujar flechas. La gráfica muestra, con barras, el valor "
                "de cada componente de cada vector (y del resultado en azul).")

    # ---- SUMA / RESTA ----
    if op in ("suma", "resta"):
        o = " + " if op == "suma" else " − "
        r = res[op]
        letras = o.join(NOMBRES[i] for i in I)
        general = "(" + ", ".join(o.join(low(i) + SUB[j] for i in I) for j in J) + ")"
        sustit = "(" + ", ".join(o.join(fmt(vs[i][j]) for i in I) for j in J) + ")"
        por_pos = [f"{EJES[j]}: " + o.join(fmt(vs[i][j]) for i in I) + f" = {fmt(r[j])}" for j in J]
        if op == "suma":
            return [paso1,
                    ("Elige la regla", [f"{letras} = {general}"],
                     "Se suma cada posición con su pareja: " + ", ".join(f"{EJES[j]} con {EJES[j]}" for j in J) + "."),
                    ("Cambia las letras por tus números", [sustit], None),
                    ("Calcula posición por posición", por_pos, None),
                    ("Resultado", [f"{letras} = {fmt_vec(r)}"], None),
                    grafica("Cada flecha punteada es el mismo vector colocado donde termina el anterior. "
                            "La flecha azul es el atajo directo desde el origen hasta el final del recorrido.")]
        return [paso1,
                ("Elige la regla", [f"{letras} = {general}"],
                 "Restar es sumar el vector contrario: el mismo recorrido, pero en sentido opuesto."),
                ("Cambia las letras por tus números", [sustit], None),
                ("Calcula posición por posición", por_pos, None),
                ("Resultado", [f"{letras} = {fmt_vec(r)}"],
                 "Los números pueden salir negativos: la flecha apunta hacia atrás en esa dirección. Es normal."),
                grafica("La flecha punteada es el opuesto (al revés) colocado en la punta de A. "
                        "El resultado azul va del origen hasta su final.")]

    # ---- ESCALAR ----
    k = datos["k"]
    if op == "escalar":
        r = res["escalar"]
        return [paso1,
                ("Elige la regla", [f"k · {nom_t} = (" + ", ".join(f"k·{low(t)}{SUB[j]}" for j in J) + ")"],
                 "Un escalar es un número normal. Multiplicarlo por un vector lo estira, sin cambiar hacia dónde apunta."),
                (f"Usa k = {k}", [f"{k} · {fmt_vec(T)}"], None),
                ("Multiplica cada posición", [f"{EJES[j]}: {k} × {fmt(T[j])} = {fmt(r[j])}" for j in J], None),
                ("Resultado", [f"{k}·{nom_t} = {fmt_vec(r)}"], None),
                grafica(f"Como k = {k} es mayor que 1, la flecha azul es más larga que {nom_t}." if k > 1
                        else "Con k = 1 el vector queda igual.")]

    suma_cuadrados = float(np.sum(T ** 2))

    # ---- MAGNITUD ----
    if op == "magnitud":
        return [paso1,
                ("Elige la regla (Pitágoras)", [f"|{nom_t}| = √(" + " + ".join(f"{low(t)}{SUB[j]}²" for j in J) + ")"],
                 "La magnitud es la longitud de la flecha, como medirla con una regla. Sus componentes son los "
                 "lados de un triángulo (o caja) y la flecha es la diagonal."),
                ("Eleva cada posición al cuadrado", [f"{EJES[j]}²: {fmt(T[j])}² = {fmt(T[j] ** 2)}" for j in J], None),
                ("Suma los cuadrados", [" + ".join(fmt(x * x) for x in T) + f" = {fmt(suma_cuadrados)}"], None),
                ("Saca la raíz cuadrada", [f"|{nom_t}| = √{fmt(suma_cuadrados)} = {fmt(res['magnitud'])}"],
                 "Es siempre mayor o igual que 0: una longitud no puede ser negativa."),
                grafica("La franja azul sobre la flecha es la longitud medida; las líneas punteadas "
                        "muestran los puntos que están a esa misma distancia del origen.")]

    # ---- VECTOR UNITARIO ----
    u = res["unitario"]
    return [paso1,
            ("Elige la regla", [f"{nom_t}\u0302 = {nom_t} / |{nom_t}|"],
             "El vector unitario apunta igual que el original pero mide exactamente 1. Sirve para indicar solo la dirección."),
            ("Calcula primero la magnitud", [f"|{nom_t}| = √{fmt(suma_cuadrados)} = {fmt(res['magnitud'])}"], None),
            ("Divide cada posición entre la magnitud", [f"{EJES[j]}: {fmt(T[j])} ÷ {fmt(res['magnitud'])} = {fmt(u[j])}" for j in J], None),
            ("Resultado", [f"{nom_t}\u0302 = {fmt_vec(u)}"], None),
            ("Comprueba que mide 1", ["√(" + " + ".join(fmt(x) + "²" for x in u) + f") = {fmt(magnitud(u))}"],
             "Si da 1, está bien hecho. Pequeñas diferencias vienen del redondeo."),
            grafica("La flecha azul llega justo a la circunferencia (o esfera) de radio 1.")]


# ===========================================================================
# 4) GRÁFICAS (matplotlib): 2D, 3D y barras
# ===========================================================================
def puntos_componentes(v):
    """
    Puntos del "camino por los ejes" hasta la punta del vector: origen -> eje x -> ... -> punta.
    Sirve para dibujar las líneas punteadas de componentes en 2D y 3D.
    """
    pts = [np.zeros(len(v))]
    for j in range(len(v)):
        p = np.zeros(len(v))
        p[:j + 1] = v[:j + 1]
        pts.append(p)
    return pts


def anillos(agregar, r, color, pd=2):
    """
    Dibuja la circunferencia de radio r (2D) o tres círculos máximos que forman una esfera (3D).
    Calcula 73 puntos con seno y coseno para cerrar cada círculo.
    """
    a = np.linspace(0, 2 * np.pi, 73)
    c, s, z = r * np.cos(a), r * np.sin(a), np.zeros_like(a)
    planos = [(c, s)] if pd == 2 else [(c, s, z), (c, z, s), (z, c, s)]
    for p in planos:
        agregar(list(zip(*p)), color, ":", w=1.3, alpha=0.7)


def construir_segmentos(datos, res, op, pd, mostrar_comp):
    """
    Arma la lista de "segmentos" que se dibujan (sirve para 2D y 3D porque trabaja con pd
    coordenadas). Cada segmento es un diccionario: pts, color, estilo, flecha, grosor, alpha, etq.
    Primero van los vectores originales desde el origen, luego las líneas de construcción
    (método punta-cola, círculos...) y al final el resultado en azul.
    """
    seg = []
    t = datos["tgt"]
    I = [i for i in range(datos["n"]) if datos["dims"][i] == pd]      # vectores que sí se dibujan
    vs = [datos["vs"][i] for i in I]
    O = np.zeros(pd)
    AZUL = COLORES["R"]

    def agregar(pts, color, estilo="-", flecha=False, w=1.5, alpha=1.0, etq=""):
        seg.append(dict(pts=[np.array(p, dtype=float) for p in pts], color=color, estilo=estilo,
                        flecha=flecha, w=w, alpha=alpha, etq=etq))

    for k, v in enumerate(vs):
        agregar([O, v], COLORES[NOMBRES[I[k]]], flecha=True, w=2.5, etq=NOMBRES[I[k]])

    if op == "suma":
        p = vs[0]
        for k in range(1, len(vs)):                                   # método punta-cola
            q = p + vs[k]
            agregar([p, q], COLORES[NOMBRES[I[k]]], "--", w=1.5, alpha=0.8)
            p = q
        agregar([O, res["suma"]], AZUL, flecha=True, w=3.2, etq="Σ")
    elif op == "resta":
        p = vs[0]
        for k in range(1, len(vs)):
            q = p - vs[k]
            agregar([p, q], COLORES[NOMBRES[I[k]]], "--", w=1.5, alpha=0.8)
            p = q
        agregar([O, res["resta"]], AZUL, flecha=True, w=3.2, etq="Δ")
    elif op == "escalar":
        agregar([O, res["escalar"]], AZUL, flecha=True, w=3.2, etq=f"k·{NOMBRES[t]}")
    elif op == "magnitud":
        agregar([O, datos["vs"][t]], AZUL, w=7, alpha=0.4, etq=f"|{NOMBRES[t]}| = {fmt(res['magnitud'])}")
        anillos(agregar, res["magnitud"], AZUL, pd)
    else:
        agregar([O, res["unitario"]], AZUL, flecha=True, w=3.2, etq=f"{NOMBRES[t]}\u0302")
        anillos(agregar, 1.0, AZUL, pd)

    if mostrar_comp or op == "magnitud":                              # líneas de componentes
        for i in ([t] if op == "magnitud" else I):
            agregar(puntos_componentes(datos["vs"][i]), COLORES[NOMBRES[i]], ":", w=1.3, alpha=0.8)
    return seg


def limites(seg):
    """Rango común de los ejes para que todo quepa, con margen y dejando el 0 visible."""
    todos = np.concatenate([np.array(s["pts"]).ravel() for s in seg])
    lo, hi = min(0.0, float(todos.min())), max(1.0, float(todos.max()))
    pad = (hi - lo) * 0.12 + 0.6
    return (lo - pad if lo < 0 else -pad), hi + pad


def dibujar_2d(fig, seg):
    """Gráfica 2D: flechas con ax.annotate y líneas con ax.plot, con la misma escala en x e y."""
    ax = fig.add_subplot(111)
    lim = limites(seg)
    for s in seg:
        P = np.array(s["pts"])
        if s["flecha"]:
            ax.annotate("", xy=P[1], xytext=P[0],
                        arrowprops=dict(arrowstyle="-|>", color=s["color"], lw=s["w"], shrinkA=0, shrinkB=0, alpha=s["alpha"]))
        else:
            ax.plot(P[:, 0], P[:, 1], linestyle=s["estilo"], color=s["color"], lw=s["w"], alpha=s["alpha"])
        if s["etq"]:
            ax.text(P[-1, 0], P[-1, 1], "  " + s["etq"], color=s["color"], fontsize=12, fontweight="bold")
    ax.set_xlim(*lim)
    ax.set_ylim(*lim)
    ax.set_aspect("equal", adjustable="box")
    ax.axhline(0, color="black", lw=1.1)
    ax.axvline(0, color="black", lw=1.1)
    ax.grid(True, alpha=0.3)
    ax.set_xlabel("x")
    ax.set_ylabel("y")


def dibujar_3d(fig, seg):
    """
    Gráfica 3D: flechas con ax.quiver y líneas con ax.plot. Con el mouse se puede girar,
    acercar y mover la gráfica (la barra de herramientas de matplotlib también ayuda).
    """
    ax = fig.add_subplot(111, projection="3d")
    lo, hi = limites(seg)
    for s in seg:
        P = np.array(s["pts"])
        if s["flecha"]:
            d = P[1] - P[0]
            ax.quiver(P[0][0], P[0][1], P[0][2], d[0], d[1], d[2], color=s["color"],
                      linewidth=s["w"], alpha=s["alpha"], arrow_length_ratio=0.12)
        else:
            ax.plot(P[:, 0], P[:, 1], P[:, 2], linestyle=s["estilo"], color=s["color"], lw=s["w"], alpha=s["alpha"])
        if s["etq"]:
            ax.text(P[-1][0], P[-1][1], P[-1][2], "  " + s["etq"], color=s["color"], fontsize=11, fontweight="bold")
    ax.set_xlim(lo, hi)
    ax.set_ylim(lo, hi)
    ax.set_zlim(lo, hi)
    ax.set_box_aspect((1, 1, 1))                                       # mismos tamaños en x, y, z
    ax.set_xlabel("x")
    ax.set_ylabel("y")
    ax.set_zlabel("z")


def dibujar_barras(fig, datos, res, op, pd):
    """
    Gráfica de barras para cuando NO se pueden dibujar flechas (4 o 5 componentes, o vectores
    con distinto número de componentes). Cada grupo de barras es una componente (x, y, z, w, u);
    cada color es un vector; el resultado (si es un vector) va en azul.
    Posible cambio: reemplazar por coordenadas paralelas (una línea por vector).
    """
    ax = fig.add_subplot(111)
    series = [(NOMBRES[i], datos["vs"][i], COLORES[NOMBRES[i]]) for i in range(datos["n"])]
    t = NOMBRES[datos["tgt"]]
    etiquetas = {"suma": "Σ (suma)", "resta": "Δ (resta)", "escalar": f"k·{t}", "unitario": f"{t}\u0302 (unitario)"}
    if op in etiquetas and isinstance(res[op], np.ndarray):
        series.append((etiquetas[op], res[op], COLORES["R"]))
    m = max(len(v) for _, v, _ in series)
    ancho = 0.8 / len(series)
    for s_i, (nombre, v, color) in enumerate(series):
        x = np.arange(len(v)) + (s_i - (len(series) - 1) / 2) * ancho
        barras = ax.bar(x, v, ancho * 0.95, color=color, label=nombre)
        try:
            ax.bar_label(barras, labels=[fmt(a) for a in v], fontsize=7, padding=1)   # valor sobre cada barra
        except AttributeError:                                                        # matplotlib antiguo: sin etiquetas
            pass
    ax.set_xticks(range(m))
    ax.set_xticklabels(EJES[:m])
    ax.set_ylabel("valor de la componente")
    ax.grid(True, axis="y", alpha=0.3)
    ax.legend(fontsize=8)
    if pd == 0 and op in ("suma", "resta"):
        ax.set_title(mensaje_dimensiones(datos, op), color=ROJO_AVISO, fontsize=9, wrap=True)
    elif op == "magnitud":
        ax.set_title(f"Componentes de los vectores   |{t}| = {fmt(res['magnitud'])}", fontsize=10)
    else:
        ax.set_title("Gráfica de componentes (4 o 5 componentes no se dibujan como flechas)", fontsize=10)


def dibujar(fig, datos, res, op, mostrar_comp):
    """
    Dibuja SIEMPRE una gráfica: 2D con 2 componentes, 3D con 3 componentes y barras en
    cualquier otro caso (4 o 5 componentes, o suma/resta con componentes distintos).
    """
    fig.clear()
    pd = dim_grafica(datos, op)
    if pd in (2, 3):
        seg = construir_segmentos(datos, res, op, pd, mostrar_comp)
        (dibujar_2d if pd == 2 else dibujar_3d)(fig, seg)
    else:
        dibujar_barras(fig, datos, res, op, pd)


# ===========================================================================
# 5) LEER VECTORES DESDE DOCUMENTOS (txt, csv, pdf, docx, xlsx)
# ===========================================================================
def celda_a_texto(c):
    """Convierte una celda de Excel a texto; un 3.0 se escribe "3" (los decimales reales se rechazarán después)."""
    if isinstance(c, float) and c.is_integer():
        return str(int(c))
    return str(c)


def extraer_lineas(ruta):
    """
    Lee un archivo y devuelve su contenido como lista de líneas de texto.
    Según la extensión usa: open (txt/csv), pypdf (pdf), python-docx (docx) u openpyxl (xlsx).
    Posible cambio: usar otra librería por formato; solo hay que seguir devolviendo una lista de líneas.
    """
    ext = os.path.splitext(ruta)[1].lower()
    if ext in (".txt", ".csv", ".tsv", ".dat", ".md", ""):
        with open(ruta, "rb") as f:
            crudo = f.read()
        if crudo.startswith((b"\xff\xfe", b"\xfe\xff")):                    # archivo UTF-16 (Bloc de notas)
            return crudo.decode("utf-16").splitlines()
        try:
            return crudo.decode("utf-8-sig").splitlines()
        except UnicodeDecodeError:
            return crudo.decode("cp1252", errors="replace").splitlines()   # ANSI de Windows
    if ext == ".pdf":
        if PdfReader is None:
            raise ImportError("Para leer PDF instala pypdf:  pip install pypdf")
        lector = PdfReader(ruta)
        return "\n".join((p.extract_text() or "") for p in lector.pages).splitlines()
    if ext == ".docx":
        if docx is None:
            raise ImportError("Para leer Word instala python-docx:  pip install python-docx")
        d = docx.Document(ruta)
        lineas = [p.text for p in d.paragraphs]
        for tabla in d.tables:                                              # cada fila de tabla = una línea
            for fila in tabla.rows:
                lineas.append(", ".join(c.text for c in fila.cells))
        return lineas
    if ext in (".xlsx", ".xlsm"):
        if openpyxl is None:
            raise ImportError("Para leer Excel instala openpyxl:  pip install openpyxl")
        libro = openpyxl.load_workbook(ruta, data_only=True, read_only=True)
        lineas = []
        for fila in libro.active.iter_rows(values_only=True):               # cada fila = un vector
            celdas = [celda_a_texto(c) for c in fila if c is not None and str(c).strip() != ""]
            if celdas:
                lineas.append(", ".join(celdas))
        libro.close()
        return lineas
    if ext in (".doc", ".xls"):
        raise ValueError("Los formatos antiguos .doc y .xls no se pueden leer. Guárdalos como .docx o .xlsx.")
    raise ValueError(f"Formato no compatible ({ext}). Usa .txt, .csv, .pdf, .docx o .xlsx.")


def parsear_vectores(lineas):
    """
    Convierte las líneas de texto en una lista de vectores (listas de enteros) y VALIDA todo:
      - solo dígitos (sin letras, sin signo negativo, sin punto decimal),
      - cada número entre VALOR_MINIMO y MAX_DIGITOS dígitos,
      - de MIN_COMP a MAX_COMP componentes por vector y de MIN_VECTORES a MAX_VECTORES vectores.
    Se acepta una etiqueta opcional ("A = (3, 2)", "B: 1 4"). Las líneas sin números se ignoran.
    Si algo está mal lanza ValueError indicando la línea; así no se carga nada a medias.
    """
    vectores = []
    for num, linea in enumerate(lineas, 1):
        linea = linea.strip()
        if not re.search(r"\d", linea):                                     # títulos o líneas vacías
            continue
        linea = re.sub(r"^\s*(?:vector\s+)?[A-Za-z]\s*[=:]\s*", "", linea, flags=re.I)   # quita "A =" / "B:"
        limpio = re.sub(r"[()\[\]{}]", " ", linea)                           # quita paréntesis y corchetes
        v = []
        for tok in (t for t in re.split(r"[,;\s|]+", limpio) if t):
            if not re.fullmatch(r"[0-9]+", tok):                            # solo dígitos ASCII
                motivo = ("es negativo" if tok[0] in "-−–" else
                          "tiene punto decimal" if "." in tok else "tiene letras o símbolos")
                raise ValueError(f"Línea {num}: «{tok}» {motivo}. Solo se permiten enteros sin signo.")
            if len(tok) > MAX_DIGITOS or int(tok) < VALOR_MINIMO:
                raise ValueError(f"Línea {num}: «{tok}» debe estar entre {VALOR_MINIMO} y {'9' * MAX_DIGITOS}.")
            v.append(int(tok))
        if not MIN_COMP <= len(v) <= MAX_COMP:
            raise ValueError(f"Línea {num}: el vector tiene {len(v)} componentes; deben ser de {MIN_COMP} a {MAX_COMP}.")
        vectores.append(v)
    if not MIN_VECTORES <= len(vectores) <= MAX_VECTORES:
        raise ValueError(f"El archivo tiene {len(vectores)} vectores; deben ser de {MIN_VECTORES} a {MAX_VECTORES}.")
    return vectores


# ===========================================================================
# 6) VENTANA PRINCIPAL (tkinter)
# ===========================================================================
class App(tk.Tk):
    """
    Ventana principal. Crea los controles; cada vez que el usuario cambia algo se llama a
    actualizar(), que lee y valida los datos, calcula y vuelve a mostrar resultado, gráfica y
    procedimiento.
    """

    def __init__(self):
        """Crea la ventana, las variables que guardan lo que escribe el usuario y los controles."""
        super().__init__()
        self.title("Proyecto Vectores - Álgebra Lineal")
        self.geometry("1400x860")
        self.silencio = False                                     # True = no recalcular (al cambiar varias cajas a la vez)

        self.n_vec = tk.IntVar(value=2)                           # cantidad de vectores (2 a 5)
        self.op = tk.StringVar(value="suma")                      # operación elegida
        self.k_var = tk.StringVar(value="2")                      # escalar k
        self.tgt_var = tk.StringVar(value="A")                    # vector objetivo
        self.mostrar_comp = tk.BooleanVar(value=False)            # mostrar líneas de componentes
        self.explicar = tk.BooleanVar(value=True)                 # explicar como a un principiante
        self.dims = [tk.StringVar(value="2") for _ in range(MAX_VECTORES)]     # componentes de cada vector
        iniciales = [[3, 2, 1, 4, 2], [1, 4, 2, 3, 5], [2, 1, 3, 5, 4], [4, 3, 5, 1, 2], [5, 2, 4, 3, 1]]
        self.vals = [[tk.StringVar(value=str(x)) for x in fila] for fila in iniciales]   # 5 vectores x 5 cajas
        self.entradas = [[None] * MAX_COMP for _ in range(MAX_VECTORES)]
        self.marcos_vec = []
        self.cierres = []

        self.crear_widgets()
        for fila in self.vals:                                    # al escribir en cualquier caja se recalcula
            for var in fila:
                var.trace_add("write", lambda *a: self.actualizar())
        self.k_var.trace_add("write", lambda *a: self.actualizar())
        self.actualizar_visibilidad()
        self.actualizar()

    # ------------------------------------------------------------------ construcción de la interfaz
    def crear_widgets(self):
        """Construye la interfaz: a la izquierda gráfica y pestañas de texto; a la derecha menú y cajas."""
        izq = ttk.Frame(self)
        izq.pack(side="left", fill="both", expand=True, padx=8, pady=8)
        der = ttk.Frame(self)
        der.pack(side="right", fill="y", padx=8, pady=8)

        nb = ttk.Notebook(izq, height=240)                        # pestañas de texto
        nb.pack(side="bottom", fill="x")
        self.txt_ops = self.crear_texto(nb, "Todas las operaciones")
        self.txt_pasos = self.crear_texto(nb, "Procedimiento paso a paso")

        self.fig = Figure(figsize=(6.5, 4.8), dpi=100, facecolor="white")      # gráfica de matplotlib en tkinter
        self.canvas = FigureCanvasTkAgg(self.fig, master=izq)
        NavigationToolbar2Tk(self.canvas, izq).update()           # zoom, mover, guardar imagen
        self.canvas.get_tk_widget().pack(side="top", fill="both", expand=True)

        vcmd = (self.register(validar_tecla), "%P")               # validación: solo dígitos
        icmd = (self.register(self.aviso_tecla),)                 # qué hacer si se rechaza una tecla

        m1 = ttk.LabelFrame(der, text="Cantidad de vectores")
        m1.pack(fill="x", pady=3)
        for v in range(MIN_VECTORES, MAX_VECTORES + 1):
            ttk.Radiobutton(m1, text=str(v), value=v, variable=self.n_vec,
                            command=self.cambio_cantidad).pack(side="left", padx=10, pady=4)

        m2 = ttk.LabelFrame(der, text="Operación")
        m2.pack(fill="x", pady=3)
        for clave, texto in OPERACIONES:
            ttk.Radiobutton(m2, text=texto, value=clave, variable=self.op,
                            command=self.cambio_operacion).pack(anchor="w", padx=8)

        self.marco_k = ttk.Frame(der)                             # k y vector objetivo (solo cuando se necesitan)
        self.marco_k.pack(fill="x", pady=3)
        self.fila_k = ttk.Frame(self.marco_k)
        ttk.Label(self.fila_k, text="k =").pack(side="left")
        self.entrada_k = tk.Entry(self.fila_k, textvariable=self.k_var, width=6, justify="center", validate="key",
                                  validatecommand=vcmd, invalidcommand=icmd)
        self.entrada_k.pack(side="left", padx=4)
        self.fila_t = ttk.Frame(self.marco_k)
        ttk.Label(self.fila_t, text="Vector:").pack(side="left")
        self.combo_t = ttk.Combobox(self.fila_t, textvariable=self.tgt_var, values=NOMBRES[:2], width=4, state="readonly")
        self.combo_t.pack(side="left", padx=4)
        self.combo_t.bind("<<ComboboxSelected>>", lambda e: self.actualizar())

        self.zona = ttk.LabelFrame(der, text="Vectores (componentes y valores)")
        self.zona.pack(fill="x", pady=3)
        for i in range(MAX_VECTORES):                             # una fila por vector: componentes + cajas
            fila = ttk.Frame(self.zona)
            self.marcos_vec.append(fila)
            ttk.Label(fila, text=NOMBRES[i], foreground=COLORES[NOMBRES[i]], font=("Consolas", 12, "bold"), width=2).grid(row=0, column=0)
            combo = ttk.Combobox(fila, textvariable=self.dims[i], values=[str(d) for d in range(MIN_COMP, MAX_COMP + 1)],
                                 width=2, state="readonly")
            combo.grid(row=0, column=1, padx=3)
            combo.bind("<<ComboboxSelected>>", lambda e: self.cambio_componentes())
            ttk.Label(fila, text="= (", font=("Consolas", 12, "bold")).grid(row=0, column=2)
            for j in range(MAX_COMP):                             # 5 cajas; se muestran solo las necesarias
                self.entradas[i][j] = tk.Entry(fila, textvariable=self.vals[i][j], width=5, justify="center",
                                               validate="key", validatecommand=vcmd, invalidcommand=icmd,
                                               font=("Consolas", 12))
            self.cierres.append(ttk.Label(fila, text=")", font=("Consolas", 12, "bold")))

        ttk.Checkbutton(der, text="Mostrar componentes (líneas punteadas)",
                        variable=self.mostrar_comp, command=self.actualizar).pack(anchor="w", pady=2)
        ttk.Checkbutton(der, text="Explicar como a un principiante", variable=self.explicar,
                        command=self.actualizar).pack(anchor="w", pady=2)
        ttk.Button(der, text="Vectores al azar", command=self.vectores_al_azar).pack(fill="x", pady=2)
        ttk.Button(der, text="📂 Cargar vectores desde archivo (txt, pdf, docx, xlsx)",
                   command=self.cargar_archivo).pack(fill="x", pady=2)
        ttk.Button(der, text="¿Cómo debe ser el archivo?", command=self.ayuda_archivo).pack(fill="x", pady=2)

        self.lbl_aviso = tk.Label(der, text="", fg=ROJO_AVISO, wraplength=400, justify="left")
        self.lbl_aviso.pack(fill="x")
        self.lbl_res = tk.Label(der, text="", fg=COLORES["R"], font=("Consolas", 13, "bold"),
                                wraplength=400, justify="left", relief="groove", padx=8, pady=8)
        self.lbl_res.pack(fill="x", pady=6)

    def crear_texto(self, notebook, titulo):
        """Crea una pestaña con un cuadro de texto de solo lectura, barra de desplazamiento y estilos."""
        marco = ttk.Frame(notebook)
        notebook.add(marco, text=titulo)
        texto = tk.Text(marco, wrap="word", font=("Segoe UI", 10), state="disabled", padx=8, pady=6)
        barra = ttk.Scrollbar(marco, command=texto.yview)
        texto.configure(yscrollcommand=barra.set)
        barra.pack(side="right", fill="y")
        texto.pack(side="left", fill="both", expand=True)
        texto.tag_configure("titulo", font=("Segoe UI", 11, "bold"))
        texto.tag_configure("mono", font=("Consolas", 11))
        texto.tag_configure("easy", foreground="#5b3fc4", background="#efe9ff", lmargin1=24, lmargin2=24)
        return texto

    # ------------------------------------------------------------------ eventos del menú
    def aviso_tecla(self):
        """Se ejecuta cuando validar_tecla rechaza una tecla: muestra el aviso y hace sonar la campana."""
        self.lbl_aviso.config(text=MSG_ENTERO)
        self.bell()

    def cambio_cantidad(self):
        """Cambió la cantidad de vectores: muestra u oculta filas y recalcula."""
        self.actualizar_visibilidad()
        self.actualizar()

    def cambio_componentes(self):
        """Cambió el número de componentes de algún vector: muestra solo las cajas necesarias."""
        self.actualizar_visibilidad()
        self.actualizar()

    def cambio_operacion(self):
        """Cambió la operación: muestra u oculta k y el selector de vector según haga falta."""
        self.actualizar_visibilidad()
        self.actualizar()

    def actualizar_visibilidad(self):
        """
        Muestra u oculta controles: k solo en escalar; selector de vector en escalar, magnitud
        y unitario; tantas filas de vectores como la cantidad elegida; tantas cajas como componentes.
        """
        op, n = self.op.get(), self.n_vec.get()
        for fila in (self.fila_k, self.fila_t):
            fila.pack_forget()
        if op == "escalar":
            self.fila_k.pack(side="left", padx=4)
        if op in ("escalar", "magnitud", "unitario"):
            self.fila_t.pack(side="left", padx=4)
        self.combo_t.config(values=NOMBRES[:n])
        if self.tgt_var.get() not in NOMBRES[:n]:
            self.tgt_var.set("A")
        for fila in self.marcos_vec:
            fila.pack_forget()
        for i in range(n):
            self.marcos_vec[i].pack(fill="x", pady=3, padx=4)
        for i in range(MAX_VECTORES):
            d = int(self.dims[i].get())
            for j in range(MAX_COMP):
                caja = self.entradas[i][j]
                caja.grid_forget()
                if j < d:
                    caja.grid(row=0, column=j + 3, padx=2)
            self.cierres[i].grid(row=0, column=d + 3)

    def vectores_al_azar(self):
        """Llena todas las cajas con enteros del 1 al 9. Desactiva el recálculo y recalcula una sola vez."""
        self.silencio = True
        for fila in self.vals:
            for var in fila:
                var.set(str(random.randint(1, 9)))
        self.silencio = False
        self.actualizar()

    # ------------------------------------------------------------------ bono: archivos
    def ayuda_archivo(self):
        """Ventana con el formato que debe tener el archivo."""
        messagebox.showinfo("Formato del archivo",
                            "Una línea (o fila) = un vector. Números separados por coma, punto y coma o espacio.\n"
                            f"De {MIN_VECTORES} a {MAX_VECTORES} vectores, de {MIN_COMP} a {MAX_COMP} componentes cada uno.\n\n"
                            "Ejemplo:\n    A = (3, 2, 1)\n    B: 1 4 2\n    4, 5, 6\n\n"
                            "Formatos: .txt .csv .pdf .docx .xlsx\n"
                            "No se aceptan letras, negativos ni decimales (las líneas sin números, como títulos, se ignoran).")

    def cargar_archivo(self):
        """
        Pide un archivo, lo lee con extraer_lineas(), lo valida con parsear_vectores() y, si todo
        está bien, llena la pantalla con esos vectores. Si algo falla muestra el motivo y no cambia nada.
        """
        ruta = filedialog.askopenfilename(
            title="Elige el archivo con los vectores",
            filetypes=[("Archivos compatibles", "*.txt *.csv *.pdf *.docx *.xlsx *.xlsm"), ("Todos los archivos", "*.*")])
        if not ruta:                                              # el usuario canceló
            return
        try:
            vectores = parsear_vectores(extraer_lineas(ruta))
        except ImportError as e:
            messagebox.showerror("Falta una librería", str(e))
            return
        except (ValueError, OSError) as e:
            messagebox.showerror("No se pudo cargar el archivo", str(e))
            return
        except Exception as e:                                    # archivo dañado u otro error inesperado
            messagebox.showerror("No se pudo leer el archivo", f"{type(e).__name__}: {e}")
            return
        self.aplicar_vectores(vectores)
        messagebox.showinfo("Archivo cargado", f"Se cargaron {len(vectores)} vectores desde:\n{os.path.basename(ruta)}")

    def aplicar_vectores(self, vectores):
        """Escribe en la pantalla la cantidad de vectores, sus componentes y sus valores; luego recalcula."""
        self.silencio = True
        self.n_vec.set(len(vectores))
        for i, v in enumerate(vectores):
            self.dims[i].set(str(len(v)))
            for j, x in enumerate(v):
                self.vals[i][j].set(str(x))
        self.silencio = False
        self.actualizar_visibilidad()
        self.actualizar()

    # ------------------------------------------------------------------ lectura y validación de datos
    def leer_datos(self):
        """
        Lee las cajas visibles y las valida (las inválidas se pintan de rojo). Si todo está bien devuelve
        {n, dims, vs (arreglos de numpy), k, tgt}; si hay algún error devuelve None.
        """
        n = self.n_vec.get()
        ok = True
        dims, vs = [], []
        for i in range(MAX_VECTORES):
            d = int(self.dims[i].get())
            for j in range(MAX_COMP):
                visible = i < n and j < d
                valido = es_natural(self.vals[i][j].get())
                self.entradas[i][j].config(bg="white" if (valido or not visible) else "#ffd6da")
                if visible and not valido:
                    ok = False
            if i < n and ok:
                dims.append(d)
                vs.append(np.array([int(self.vals[i][j].get()) for j in range(d)], dtype=float))
        k_texto = self.k_var.get()
        k_valido = es_natural(k_texto)
        self.entrada_k.config(bg="white" if k_valido else "#ffd6da")
        if self.op.get() == "escalar" and not k_valido:
            ok = False
        if not ok:
            return None
        return {"n": n, "dims": dims, "vs": vs, "k": int(k_texto) if k_valido else 1,
                "tgt": NOMBRES.index(self.tgt_var.get())}

    # ------------------------------------------------------------------ actualizar todo
    def actualizar(self):
        """
        Recalcula y muestra TODO: resultado, gráfica, todas las operaciones y procedimiento.
        Si hay un campo inválido muestra el aviso y conserva lo último calculado.
        """
        if self.silencio:
            return
        datos = self.leer_datos()
        if datos is None:
            self.lbl_aviso.config(text=MSG_ENTERO + " Corrige el campo marcado en rojo.")
            return
        self.lbl_aviso.config(text="")
        op = self.op.get()
        res = calcular(datos)
        self.lbl_res.config(text=texto_resultado(datos, res, op))
        dibujar(self.fig, datos, res, op, self.mostrar_comp.get())
        self.canvas.draw_idle()
        self.escribir(self.txt_ops, texto_cuadritos(datos, res))
        self.escribir_pasos(construir_pasos(datos, res, op))

    def escribir(self, widget, texto):
        """Reemplaza el contenido de un cuadro de texto de solo lectura."""
        widget.config(state="normal")
        widget.delete("1.0", "end")
        widget.insert("end", texto, "mono")
        widget.config(state="disabled")

    def escribir_pasos(self, pasos):
        """Muestra el procedimiento: título en negrita, cálculos en monoespaciado y explicación opcional."""
        t = self.txt_pasos
        t.config(state="normal")
        t.delete("1.0", "end")
        for numero, (titulo, lineas, explicacion) in enumerate(pasos, 1):
            t.insert("end", f"Paso {numero}. {titulo}\n", "titulo")
            for linea in lineas:
                t.insert("end", "    " + linea + "\n", "mono")
            if explicacion and self.explicar.get():
                t.insert("end", explicacion + "\n", "easy")
            t.insert("end", "\n")
        t.config(state="disabled")


# ===========================================================================
# 7) PROGRAMA PRINCIPAL
# ===========================================================================
def main():
    """Crea la ventana y la deja funcionando hasta que el usuario la cierre."""
    App().mainloop()


# Este bloque se ejecuta solo cuando se corre el archivo directamente
if __name__ == "__main__":
    main()