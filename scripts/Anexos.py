"""
Genera los tres anexos del estudio de conexion, en LaTeX y con el mismo estilo
que el informe principal.

    python3 scripts/Anexos.py

NO toca PowerFactory. Lee lo que ya esta generado:
  - resultados/Resultados_Flujo_Carga.xlsx
  - resultados/Resultados_Cortocircuito.xlsx
  - resultados/Coordinacion_de_Protecciones/Resultados_Coordinacion_Protecciones.xlsx
  - los diagramas unifilares SVG que exportan Flujo_carga.py y Cortocircuito.py
  - las figuras de informe/ que genera Graficas_Informe.py

Salidas (en informe/anexos/):
  Anexo_A_Flujo_Carga.tex
  Anexo_B_Cortocircuito.tex
  Anexo_C_Coordinacion_Protecciones.tex
  figuras/*.pdf   diagramas unifilares convertidos de SVG a PDF

SOBRE LA COMPILACION A PDF
Los .tex quedan listos para compilar. Si hay un motor de LaTeX instalado
(pdflatex), el script los compila y deja los PDF al lado; si no lo hay --- que es
el caso habitual en una maquina sin distribucion de TeX --- los deja listos para
subir a Overleaf junto con la carpeta figuras/, y lo avisa por consola. No se
genera el PDF por otra via para no producir un documento con un aspecto distinto
al del informe principal, que es justamente lo que se quiere evitar.

SOBRE LOS DIAGRAMAS UNIFILARES
Los scripts de simulacion exportan un SVG por cada combinacion simulada: son 72
en flujo de carga y 93 en cortocircuito. Incluirlos todos haria un anexo
inmanejable, asi que se selecciona un subconjunto representativo --- los
escenarios del numeral 8.2 en flujo de carga, y las barras de mayor nivel de
falla en cortocircuito --- y se deja constancia en el propio anexo de cuantos
diagramas existen en total y donde estan los demas. La seleccion se controla con
las constantes ESCENARIOS_FLUJO y BARRAS_CORTO de este archivo.
"""

import math
import re
import shutil
import subprocess
from collections import defaultdict
from pathlib import Path

import openpyxl

from parametros import P

BASE_DIR = Path(__file__).resolve().parent.parent
RES_FLUJO = BASE_DIR / "resultados" / "Resultados_Flujo_Carga.xlsx"
RES_CORTO = BASE_DIR / "resultados" / "Resultados_Cortocircuito.xlsx"
RES_COORD = BASE_DIR / "resultados" / "Coordinacion_de_Protecciones" / "Resultados_Coordinacion_Protecciones.xlsx"
SVG_FLUJO = BASE_DIR / "resultados" / "graficos_red"
SVG_CORTO = BASE_DIR / "resultados" / "graficos_red_corto"
INFORME_DIR = BASE_DIR / "informe"
SALIDA_DIR = INFORME_DIR / "anexos"
FIGURAS_DIR = SALIDA_DIR / "figuras"

CASO_BASE = "CasoBase"
CASO_CARGA_PURA = "CargaPura"

# Diagramas unifilares que se incluyen en el Anexo A: los que corresponden a los
# tres escenarios del numeral 8.2, en el ano base.
ESCENARIOS_FLUJO = [
    ("Anio2026_Hora15_CasoBase", "Escenario 8.2.1 --- carga pura, hora de m\\'axima demanda"),
    ("Anio2026_Hora12_NetworkVariation", "Escenario 8.2.2 --- m\\'axima generaci\\'on (hora 12)"),
    ("Anio2026_Hora15_NetworkVariation", "Escenario 8.2.3 --- m\\'axima demanda con generaci\\'on"),
]

# Barras cuyo diagrama de falla se incluye en el Anexo B: la de mayor nivel de
# falla de la red, la del punto de conexion y la de baja tension del proyecto.
BARRAS_CORTO = [
    ("PTO BOYACA 13.2kV", "Trifasico", "Barra de mayor nivel de falla de la subestaci\\'on"),
    ("P1 15344 13.2kV", "Trifasico", "Punto de conexi\\'on del proyecto"),
    ("Puerto Boyaca 800V", "Trifasico", "Barra de baja tensi\\'on del centro de transformaci\\'on"),
]


# ---------------------------------------------------------------------------
# Utilidades
# ---------------------------------------------------------------------------
def escapar(texto):
    """Escapa los caracteres que LaTeX interpreta como ordenes."""
    if texto is None:
        return ""
    s = str(texto)
    for a, b in (("\\", r"\textbackslash{}"), ("&", r"\&"), ("%", r"\%"), ("$", r"\$"),
                 ("#", r"\#"), ("_", r"\_"), ("{", r"\{"), ("}", r"\}"), ("~", r"\textasciitilde{}"),
                 ("^", r"\textasciicircum{}")):
        s = s.replace(a, b)
    return s


def num(valor, decimales=4):
    if valor is None:
        return "---"
    if isinstance(valor, bool):
        return "S\\'i" if valor else "No"
    if isinstance(valor, (int, float)):
        return f"{valor:.{decimales}f}"
    return escapar(valor)


def leer_hoja(ruta, hoja):
    wb = openpyxl.load_workbook(ruta, data_only=True)
    if hoja not in wb.sheetnames:
        return [], []
    filas = list(wb[hoja].iter_rows(values_only=True))
    if not filas:
        return [], []
    return list(filas[0]), filas[1:]


def convertir_svg(origen, destino):
    """SVG -> PDF, para poder incluirlo con pdfLaTeX (que no lee SVG)."""
    try:
        from svglib.svglib import svg2rlg
        from reportlab.graphics import renderPDF
    except ImportError:
        return None, ("Falta la libreria svglib/reportlab para convertir los diagramas "
                      "unifilares de SVG a PDF. Instalala con: pip install svglib reportlab")
    try:
        dibujo = svg2rlg(str(origen))
        if dibujo is None:
            return None, f"No se pudo leer el SVG '{origen.name}'"
        renderPDF.drawToFile(dibujo, str(destino))
        return destino, None
    except Exception as exc:
        return None, f"Error convirtiendo '{origen.name}': {exc}"


def buscar_svg(carpeta, patrones):
    """Primer SVG de la carpeta cuyo nombre contenga TODOS los patrones."""
    for archivo in sorted(carpeta.glob("*.svg")):
        if all(p.lower() in archivo.name.lower() for p in patrones):
            return archivo
    return None


def tabla_latex(encabezados, filas, alineacion=None, nota=None, ancho_small=True):
    """Tabla longtable: se parte sola entre paginas, que es lo que hace falta en
    un anexo donde algunas tablas tienen decenas de filas."""
    align = alineacion or ("l" + "r" * (len(encabezados) - 1))
    out = []
    if ancho_small:
        out.append(r"\small")
    out.append(r"\begin{longtable}{" + align + "}")
    out.append(r"\toprule")
    out.append(" & ".join(r"\textbf{" + escapar(h) + "}" for h in encabezados) + r" \\")
    out.append(r"\midrule")
    out.append(r"\endfirsthead")
    out.append(r"\toprule")
    out.append(" & ".join(r"\textbf{" + escapar(h) + "}" for h in encabezados) + r" \\")
    out.append(r"\midrule")
    out.append(r"\endhead")
    out.append(r"\bottomrule")
    out.append(r"\endfoot")
    for fila in filas:
        out.append(" & ".join(fila) + r" \\")
    out.append(r"\end{longtable}")
    if nota:
        out.append(r"\vspace{-0.5em}")
        out.append(r"{\small\itshape " + nota + r"}")
        out.append("")
    return "\n".join(out)


def preambulo(titulo, subtitulo):
    """Mismo preambulo que el informe principal, para que los anexos salgan con
    identica tipografia, margenes y encabezado."""
    return r"""% ============================================================================
% """ + titulo + r"""
% Proyecto: Generacion Distribuida 990 kW - Puerto Boyaca (EBSA), circuito 15344
% Generado automaticamente por scripts/Anexos.py - no editar a mano.
% Compilador recomendado: pdfLaTeX.
% ============================================================================
\documentclass[11pt,a4paper]{article}

\usepackage[spanish,es-tabla]{babel}
\usepackage[utf8]{inputenc}
\usepackage[T1]{fontenc}
\usepackage{textcomp}

\usepackage[a4paper,margin=2.5cm]{geometry}
\usepackage{fancyhdr}
\usepackage{titlesec}
\usepackage{setspace}
\onehalfspacing

\usepackage{booktabs}
\usepackage{longtable}
\usepackage{array}
\usepackage{tabularx}
\usepackage{graphicx}
\usepackage{caption}
\usepackage{float}
\usepackage{siunitx}
\sisetup{output-decimal-marker={,}, group-separator={.}}

\usepackage{enumitem}
\usepackage[dvipsnames]{xcolor}
\usepackage[colorlinks=true,linkcolor=NavyBlue,citecolor=NavyBlue,urlcolor=NavyBlue]{hyperref}

\graphicspath{{figuras/}{../}}

\pagestyle{fancy}
\fancyhf{}
\fancyhead[L]{""" + subtitulo + r"""}
\fancyhead[R]{\thepage}
\fancyfoot[C]{\small Confidencial -- Uso exclusivo tr\'amite de conexi\'on ante EBSA}
\renewcommand{\headrulewidth}{0.4pt}

\titleformat{\section}{\normalfont\Large\bfseries}{\thesection.}{0.6em}{}
\titleformat{\subsection}{\normalfont\large\bfseries}{\thesubsection}{0.6em}{}

\begin{document}

\begin{titlepage}
\centering
\vspace*{3cm}
{\Large\bfseries Estudio de Conexi\'on Simplificado\par}
\vspace{0.4cm}
{\large Generaci\'on Distribuida 990\,kW --- Puerto Boyac\'a\par}
\vspace{0.2cm}
{\large Circuito 15344 --- EBSA --- Solicitud N.\textdegree{} 16869\par}
\vspace{2.5cm}
{\Huge\bfseries """ + titulo + r"""\par}
\vspace{2cm}
{\large Resoluci\'on CREG 174 de 2021\par}
\vfill
{\small Documento generado autom\'aticamente a partir de los resultados de simulaci\'on.\par}
\end{titlepage}

\tableofcontents
\newpage
"""


# ---------------------------------------------------------------------------
# Anexo A - Flujo de carga
# ---------------------------------------------------------------------------
def anexo_flujo_carga():
    avisos = []
    h, resumen = leer_hoja(RES_FLUJO, "Resumen")
    iA, iC, iH = h.index("Anio"), h.index("Caso"), h.index("Hora")
    iConv = h.index("Convergio")
    iFD = h.index("Factor_Demanda")
    iFG = h.index("Factor_Generacion") if "Factor_Generacion" in h else None

    anios = sorted({r[iA] for r in resumen})
    casos = sorted({r[iC] for r in resumen})
    horas = sorted({r[iH] for r in resumen})
    no_convergieron = [r for r in resumen if not r[iConv]]

    cuerpo = [preambulo("Anexo A --- Simulaci\\'on de Flujo de Carga",
                        "Anexo A --- Flujo de Carga")]

    cuerpo.append(r"""\section{Alcance y metodolog\'ia}

Este anexo recoge el detalle completo de las simulaciones de flujo de carga que
sustentan las Secciones 8.1, 8.3 y 8.6 del informe principal. El cuerpo del
informe presenta las conclusiones; aqu\'i est\'an los resultados que las
respaldan, escenario por escenario y elemento por elemento.

\subsection{Herramienta y procedimiento}

Los c\'alculos se ejecutaron en \textbf{DIgSILENT PowerFactory 2024}, mediante
una rutina propia en Python que recorre autom\'aticamente todas las
combinaciones de a\~no, caso de red y hora. Por cada combinaci\'on la rutina:

\begin{enumerate}
\item activa el caso de red correspondiente y \emph{a continuaci\'on} relee los
elementos calc-relevantes --- el orden importa, porque una variaci\'on de red
puede agregar o quitar equipos;
\item fija la demanda de cada carga a su valor horario multiplicado por el
factor de crecimiento del a\~no simulado;
\item despacha cada generador a su potencia nominal multiplicada por el factor
solar de esa hora;
\item ejecuta el flujo de carga y exporta el diagrama unifilar resultante;
\item registra tensi\'on, \'angulo, cargabilidad, corrientes y p\'erdidas de
cada elemento.
\end{enumerate}

Al terminar cada caso, la demanda y el despacho vuelven a su valor original, de
modo que el modelo no queda alterado por la corrida.""")

    cuerpo.append(r"\subsection{Escenarios simulados}")
    cuerpo.append(
        f"Se simularon \\textbf{{{len(resumen)} escenarios}}: "
        f"{len(anios)} a\\~nos $\\times$ {len(casos)} casos de red $\\times$ {len(horas)} horas. "
        + (r"\textbf{Todos convergieron.}" if not no_convergieron
           else f"\\textbf{{{len(no_convergieron)} no convergieron}} (ver tabla).")
    )
    cuerpo.append("")

    filas = []
    for anio in anios:
        for caso in casos:
            sub = [r for r in resumen if r[iA] == anio and r[iC] == caso]
            if not sub:
                continue
            fg = f"{min(r[iFG] for r in sub):.3f}--{max(r[iFG] for r in sub):.3f}" if iFG else "---"
            filas.append([str(anio), escapar(caso), str(len(sub)),
                          f"{sub[0][iFD]:.4f}", fg,
                          "todos" if all(r[iConv] for r in sub) else "parcial"])
    cuerpo.append(tabla_latex(
        ["Año", "Caso de red", "Horas", "Factor demanda", "Factor generación", "Convergencia"],
        filas, alineacion="llrrrl",
        nota="El factor de generaci\\'on es el del perfil solar gaussiano: 0 en el "
             "escenario de carga pura y entre 0.056 y 1.000 en los dem\\'as."))

    # --- figuras ---
    cuerpo.append(r"\newpage\section{Curvas de demanda y generaci\'on}")
    for archivo, pie, etiqueta in (
        ("fig_demanda_cargas.png",
         "Curvas de demanda de cada una de las 8 cargas del \\'area de influencia, agrupadas por circuito.",
         "fig:anexoA-demanda-cargas"),
        ("fig_demanda_generacion.png",
         "Demanda agregada del \\'area frente a la generaci\\'on, con los escenarios del numeral 8.2 marcados.",
         "fig:anexoA-demanda-generacion"),
    ):
        if (INFORME_DIR / archivo).exists():
            cuerpo.append(figura(archivo, pie, etiqueta))
        else:
            avisos.append(f"Falta la figura {archivo}: corre antes scripts/Graficas_Informe.py")

    cuerpo.append(r"\newpage\section{Resultados gr\'aficos}")
    for archivo, pie, etiqueta in (
        ("fig_tension.png",
         "Perfiles de tensi\\'on en el punto de conexi\\'on y en el extremo del circuito.",
         "fig:anexoA-tension"),
        ("fig_cargabilidad.png",
         "Nivel de carga del tramo de evacuaci\\'on y del transformador del proyecto.",
         "fig:anexoA-cargabilidad"),
        ("fig_perdidas.png",
         "P\\'erdidas activas totales y balance del proyecto frente al escenario de carga pura.",
         "fig:anexoA-perdidas"),
    ):
        if (INFORME_DIR / archivo).exists():
            cuerpo.append(figura(archivo, pie, etiqueta))
        else:
            avisos.append(f"Falta la figura {archivo}: corre antes scripts/Graficas_Informe.py")

    # --- diagramas unifilares ---
    cuerpo.append(r"\newpage\section{Diagramas unifilares por escenario}")
    total_svg = len(list(SVG_FLUJO.glob("*.svg")))
    cuerpo.append(
        f"La rutina de simulaci\\'on export\\'o \\textbf{{{total_svg} diagramas unifilares}}, uno por "
        f"cada combinaci\\'on de a\\~no, caso y hora, disponibles en "
        f"\\texttt{{resultados/graficos\\_red/}}. A continuaci\\'on se reproducen los "
        f"correspondientes a los tres escenarios del numeral 8.2 en el a\\~no base."
    )
    cuerpo.append("")
    for patron, descripcion in ESCENARIOS_FLUJO:
        origen = buscar_svg(SVG_FLUJO, patron.split("_"))
        if origen is None:
            avisos.append(f"No se encontro el diagrama de flujo para '{patron}'")
            continue
        destino = FIGURAS_DIR / f"unifilar_{patron}.pdf"
        _, error = convertir_svg(origen, destino)
        if error:
            avisos.append(error)
            continue
        cuerpo.append(figura(destino.name, descripcion, f"fig:uni-{patron}", ancho=0.92))

    # --- tablas de resultados ---
    cuerpo.append(r"\newpage\section{Resultados num\'ericos}")
    cuerpo.append(seccion_tabla_elementos(RES_FLUJO, "Nodos",
        "Tensiones por barra", ["U_pu", "U_kV", "Angulo_deg"],
        "Tensi\\'on en cada barra del sistema, para todas las combinaciones simuladas."))
    cuerpo.append(seccion_tabla_elementos(RES_FLUJO, "Lineas",
        "Cargabilidad y p\\'erdidas por l\\'inea",
        ["Cargabilidad_%", "P_perdidas_MW", "Q_perdidas_Mvar"],
        "Nivel de carga y p\\'erdidas de cada l\\'inea."))
    cuerpo.append(seccion_tabla_elementos(RES_FLUJO, "Transformadores_2Devanados",
        "Transformador del proyecto", ["Cargabilidad_%", "I_HV_kA", "I_LV_kA"],
        "Resultados del transformador de elevaci\\'on del proyecto."))
    cuerpo.append(seccion_tabla_elementos(RES_FLUJO, "Generadores",
        "Despacho de los generadores", ["P_MW", "Q_Mvar", "I_kA"],
        "Potencia efectivamente despachada por cada generador en cada hora."))

    cuerpo.append(r"\end{document}")
    return "\n\n".join(cuerpo), avisos


def figura(archivo, pie, etiqueta, ancho=1.0):
    return (r"\begin{figure}[H]" "\n" r"\centering" "\n"
            rf"\includegraphics[width={ancho}\textwidth]{{{archivo}}}" "\n"
            rf"\caption{{{pie}}}" "\n" rf"\label{{{etiqueta}}}" "\n" r"\end{figure}")


def seccion_tabla_elementos(ruta, hoja, titulo, columnas, descripcion):
    """Tabla larga de una hoja de resultados, filtrada a las columnas utiles."""
    h, filas = leer_hoja(ruta, hoja)
    if not filas:
        return rf"\subsection{{{titulo}}}" "\n\n" + "Sin datos en la hoja " + escapar(hoja) + "."
    claves = [c for c in ("Nombre", "Anio", "Caso", "Hora", "Barra_Falla", "Tipo_Falla") if c in h]
    cols = claves + [c for c in columnas if c in h]
    idx = [h.index(c) for c in cols]
    datos = [[num(f[i], 4) if isinstance(f[i], float) else escapar(f[i]) for i in idx] for f in filas]
    cab = [c.replace("_", " ") for c in cols]
    return (rf"\subsection{{{titulo}}}" "\n\n" + descripcion + "\n\n"
            + tabla_latex(cab, datos, alineacion="l" * len(claves) + "r" * (len(cols) - len(claves))))


# ---------------------------------------------------------------------------
# Anexo B - Cortocircuito
# ---------------------------------------------------------------------------
def anexo_cortocircuito():
    avisos = []
    capacidad = P.red_or("CAPACIDAD_CORTE_KA")
    h, nodos = leer_hoja(RES_CORTO, "Nodos")
    iC, iB, iT, iI = h.index("Caso"), h.index("Barra_Falla"), h.index("Tipo_Falla"), h.index("Ikss_kA")
    d = {(r[iC], r[iB], r[iT]): r[iI] for r in nodos}
    casos = sorted({k[0] for k in d})
    caso_gd = next((c for c in casos if c != CASO_BASE), None)
    barras = sorted({k[1] for k in d})

    cuerpo = [preambulo("Anexo B --- An\\'alisis de Cortocircuito",
                        "Anexo B --- Cortocircuito")]

    cuerpo.append(r"""\section{Alcance y metodolog\'ia}

Este anexo recoge el detalle de los c\'alculos de cortocircuito que sustentan la
Secci\'on 8.4 del informe principal.

\subsection{M\'etodo de c\'alculo}

Se aplica el m\'etodo de la fuente de tensi\'on equivalente de la norma
\textbf{IEC 60909 / VDE 0102 Parte 0}, con factor de tensi\'on $c = 1.1$ para el
c\'alculo de corrientes m\'aximas:
\[
I''_{k} = \frac{c \cdot U_n}{\sqrt{3}\, Z_k}
\]

El barrido recorre \textbf{todas las barras del sistema} y, en cada una,
\textbf{los tres tipos de falla} (trif\'asica, bif\'asica y monof\'asica a
tierra), para el caso base y para el caso con el proyecto conectado. El
resultado no depende del nivel de demanda, por lo que no se repite por a\~no.

\subsection{Parametrizaci\'on de secuencia cero}

Las impedancias de secuencia cero del equivalente de red se ajustaron para
reproducir el dato oficial del Operador de Red: una corriente de falla
monof\'asica de \textbf{2.622\,kA} en el barraje de 115\,kV. Esta verificaci\'on
es lo que hace utilizables las corrientes de falla a tierra de este anexo para
el dise\~no de la malla de puesta a tierra.""")

    if (INFORME_DIR / "fig_cortocircuito.png").exists():
        cuerpo.append(r"\newpage\section{Resultados gr\'aficos}")
        cuerpo.append(figura("fig_cortocircuito.png",
                             "Nivel de falla por barra frente a la capacidad de corte, e incremento aportado por el proyecto.",
                             "fig:anexoB-corto"))

    # resumen comparativo
    cuerpo.append(r"\newpage\section{Niveles de falla por barra}")
    filas = []
    for b in barras:
        fila = [escapar(b)]
        for tf in ("Trifasico", "Bifasico", "Monofasico"):
            fila.append(num(d.get((CASO_BASE, b, tf)), 4))
            fila.append(num(d.get((caso_gd, b, tf)), 4))
        vals = [v for v in (d.get((caso_gd, b, tf)) for tf in ("Trifasico", "Bifasico", "Monofasico")) if v]
        fila.append(num(capacidad - max(vals), 2) if vals else "---")
        filas.append(fila)
    cuerpo.append(tabla_latex(
        ["Barra", "3F sin", "3F con", "2F sin", "2F con", "1F sin", "1F con", "Margen"],
        filas, alineacion="l" + "r" * 7,
        nota=f"Corrientes en kA. El margen se calcula contra la capacidad de corte de "
             f"{capacidad:.0f}\\,kA declarada por el OR. La barra de baja tensi\\'on del "
             f"proyecto no se rige por ese criterio: ver Secci\\'on 8.4 del informe."))

    # diagramas
    cuerpo.append(r"\newpage\section{Diagramas de falla}")
    total_svg = len(list(SVG_CORTO.glob("*.svg")))
    cuerpo.append(
        f"La rutina export\\'o \\textbf{{{total_svg} diagramas}}, uno por cada combinaci\\'on de "
        f"caso, barra fallada y tipo de falla, disponibles en "
        f"\\texttt{{resultados/graficos\\_red\\_corto/}}. Se reproducen los de las barras "
        f"m\\'as representativas."
    )
    cuerpo.append("")
    for barra, tipo, descripcion in BARRAS_CORTO:
        origen = buscar_svg(SVG_CORTO, [barra, tipo])
        if origen is None:
            avisos.append(f"No se encontro el diagrama de falla para '{barra}' / {tipo}")
            continue
        seguro = re.sub(r"[^A-Za-z0-9]+", "_", f"{barra}_{tipo}")
        destino = FIGURAS_DIR / f"falla_{seguro}.pdf"
        _, error = convertir_svg(origen, destino)
        if error:
            avisos.append(error)
            continue
        cuerpo.append(figura(destino.name,
                             f"{descripcion} --- falla {tipo.lower()} en {escapar(barra)}.",
                             f"fig:falla-{seguro}", ancho=0.92))

    cuerpo.append(r"\newpage\section{Resultados num\'ericos}")
    cuerpo.append(seccion_tabla_elementos(RES_CORTO, "Nodos",
        "Corriente de falla por barra", ["Ikss_kA", "Skss_MVA"],
        "Corriente de cortocircuito inicial sim\\'etrica y potencia de cortocircuito."))
    cuerpo.append(seccion_tabla_elementos(RES_CORTO, "Lineas",
        "Aporte de las l\\'ineas", ["Ikss_bus1_kA", "Ikss_bus2_kA"],
        "Contribuci\\'on de cada l\\'inea a la corriente de falla."))
    cuerpo.append(seccion_tabla_elementos(RES_CORTO, "Generadores",
        "Aporte de los generadores", ["Ikss_kA"],
        "Contribuci\\'on de cada generador. Los generadores existentes GD1 y GD2 "
        "est\\'an configurados sin aporte a cortocircuito."))

    cuerpo.append(r"\end{document}")
    return "\n\n".join(cuerpo), avisos


# ---------------------------------------------------------------------------
# Anexo C - Coordinacion de protecciones
# ---------------------------------------------------------------------------
def anexo_coordinacion():
    avisos = []
    cuerpo = [preambulo("Anexo C --- Coordinaci\\'on de Protecciones",
                        "Anexo C --- Coordinaci\\'on de Protecciones")]

    cuerpo.append(r"""\section{Alcance y metodolog\'ia}

Este anexo recoge el detalle de la comprobaci\'on de coordinaci\'on que sustenta
la Secci\'on 9.2 y el anexo de protecciones del informe principal.

\subsection{Qu\'e se comprueba}

No se trata de diseñar de nuevo la protecci\'on del circuito, sino de verificar
si el \textbf{ajuste ya aprobado y en operaci\'on} del reconectador de cabecera
del circuito 15344 sigue coordinando correctamente una vez se agrega el aporte
de falla del proyecto.

\subsection{F\'ormula aplicada}

Curva IEC Normal Inverse, seg\'un IEC 60255-151 --- la misma familia con la que
el Operador de Red ajust\'o la protecci\'on de cabecera:
\[
t(I) = \mathrm{Dial} \cdot \frac{k}{\left(I/I_{arranque}\right)^{\alpha} - 1},
\qquad k = 0.14,\quad \alpha = 0.02
\]

Si la corriente supera el umbral instant\'aneo, el rel\'e no espera la curva
temporizada y opera en tiempo definido corto.""")

    # ajustes desde el Excel de parametros
    cuerpo.append(r"\subsection{Ajustes evaluados}")
    filas_cab = [
        [escapar("Curva"), escapar(P.cabecera("CURVA")), "---"],
        ["Arranque 51", num(P.cabecera("PICKUP_51_A"), 1), "A"],
        ["Dial 51", num(P.cabecera("DIAL_51"), 2), "---"],
        ["Instantáneo 50", num(P.cabecera("INST_50_A"), 1), "A"],
        ["Arranque 51N", num(P.cabecera("PICKUP_51N_A"), 1), "A"],
        ["Dial 51N", num(P.cabecera("DIAL_51N"), 2), "---"],
        ["Instantáneo 50N", num(P.cabecera("INST_50N_A"), 1), "A"],
        ["Tiempo de recierre rápido", num(P.cabecera("TIEMPO_RECIERRE_S"), 1), "s"],
    ]
    cuerpo.append(tabla_latex(["Parámetro", "Valor", "Unidad"],
                              [[escapar(a), b, escapar(c)] for a, b, c in filas_cab],
                              alineacion="lrl",
                              nota="Ajustes reales de la protecci\\'on de cabecera, seg\\'un el documento "
                                   "de insumos del Operador de Red."))

    if (INFORME_DIR / "Curva_TCC_Cabecera_15344.png").exists():
        cuerpo.append(r"\newpage\section{Curva tiempo-corriente}")
        cuerpo.append(figura("Curva_TCC_Cabecera_15344.png",
                             "Curva TCC de la protecci\\'on de cabecera, con los puntos de falla evaluados.",
                             "fig:anexoC-tcc"))

    # tablas del Excel de coordinacion
    if RES_COORD.exists():
        cuerpo.append(r"\newpage\section{Resultados}")
        wb = openpyxl.load_workbook(RES_COORD, data_only=True)
        for hoja in wb.sheetnames:
            ws = wb[hoja]
            filas = [f for f in ws.iter_rows(values_only=True) if any(v is not None for v in f)]
            if not filas:
                continue
            cuerpo.append(rf"\subsection{{{escapar(hoja.replace('_', ' '))}}}")
            ancho = max(len(f) for f in filas)
            datos = [[num(v, 3) if isinstance(v, float) else escapar(v)
                      for v in list(f) + [None] * (ancho - len(f))] for f in filas]
            cuerpo.append(tabla_latex([f"C{i+1}" for i in range(ancho)], datos[1:],
                                      alineacion="l" * ancho) if ancho > 4
                          else tabla_latex([escapar(v) for v in datos[0]], datos[1:],
                                           alineacion="l" * ancho))
    else:
        avisos.append(f"No existe {RES_COORD.name}: corre antes scripts/Coordinacion_Protecciones.py")

    cuerpo.append(r"\end{document}")
    return "\n\n".join(cuerpo), avisos


# ---------------------------------------------------------------------------
def compilar(ruta_tex):
    """Compila a PDF si hay un motor de LaTeX. Devuelve (ruta_pdf | None, aviso)."""
    if shutil.which("pdflatex") is None:
        return None, None
    for _ in range(2):  # dos pasadas: la segunda resuelve el indice
        subprocess.run(["pdflatex", "-interaction=nonstopmode", ruta_tex.name],
                       cwd=ruta_tex.parent, capture_output=True)
    pdf = ruta_tex.with_suffix(".pdf")
    return (pdf, None) if pdf.exists() else (None, f"pdflatex no produjo {pdf.name}")


def main():
    SALIDA_DIR.mkdir(parents=True, exist_ok=True)
    FIGURAS_DIR.mkdir(parents=True, exist_ok=True)

    # las figuras del informe se copian a figuras/ para que los anexos sean
    # autocontenidos y se puedan subir a Overleaf como carpeta independiente
    for png in INFORME_DIR.glob("*.png"):
        shutil.copy2(png, FIGURAS_DIR / png.name)

    todos_avisos = []
    for nombre, generador in (
        ("Anexo_A_Flujo_Carga", anexo_flujo_carga),
        ("Anexo_B_Cortocircuito", anexo_cortocircuito),
        ("Anexo_C_Coordinacion_Protecciones", anexo_coordinacion),
    ):
        texto, avisos = generador()
        ruta = SALIDA_DIR / f"{nombre}.tex"
        ruta.write_text(texto, encoding="utf-8")
        lineas = len(texto.splitlines())
        pdf, aviso_pdf = compilar(ruta)
        if aviso_pdf:
            avisos.append(aviso_pdf)
        estado = f"-> {pdf.name}" if pdf else "(sin compilar: no hay pdflatex en esta maquina)"
        print(f"  {ruta.relative_to(BASE_DIR)}  [{lineas} lineas] {estado}")
        todos_avisos.extend(avisos)

    if todos_avisos:
        print("\nAvisos:")
        for a in dict.fromkeys(todos_avisos):
            print(f"  - {a}")

    if shutil.which("pdflatex") is None:
        print("\nNo hay LaTeX instalado en esta maquina, asi que los anexos quedan en .tex.")
        print("Para obtener los PDF: sube a Overleaf la carpeta informe/anexos/ completa")
        print("(los tres .tex y la subcarpeta figuras/) y compila cada uno con pdfLaTeX.")


if __name__ == "__main__":
    main()
