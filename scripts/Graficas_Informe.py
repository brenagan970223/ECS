"""
Genera las figuras del informe de estudio de conexion a partir de los
resultados ya exportados por Flujo_carga.py y Cortocircuito.py.

NO toca PowerFactory: se ejecuta como Python normal, igual que
Coordinacion_Protecciones.py, y solo lee los Excel de resultados.

    python3 scripts/Graficas_Informe.py

Figuras generadas (en informe/, para que pdfLaTeX las encuentre junto al .tex):

  fig_tension.png       Tension minima y maxima del sistema, hora a hora, por
                        caso de red y por ano. Con las bandas de limite
                        regulatorio (0.90-1.10 p.u.) marcadas.
  fig_cargabilidad.png  Cargabilidad maxima de lineas y del transformador del
                        proyecto, hora a hora, por caso y por ano.
  fig_perdidas.png      Perdidas activas totales del sistema, hora a hora, por
                        caso y por ano, mas el balance del proyecto frente al
                        escenario de carga pura.
  fig_cortocircuito.png Corriente de cortocircuito por barra, sin y con
                        proyecto, para los tres tipos de falla, contra la
                        capacidad de corte declarada por el OR.

CRITERIOS DE PRESENTACION
Las figuras van a un informe impreso/PDF, en blanco, asi que se usa una paleta
categorica validada para fondo claro y verificada para daltonismo (deuteranopia
y tritanopia): azul #2a78d6, naranja #eb6834, aguamarina #1baf7a. Los tres casos
de red conservan SIEMPRE el mismo color en todas las figuras (el color sigue a
la entidad, no a su posicion), para que el lector no tenga que releer la leyenda
en cada grafica. Los limites regulatorios van en gris discontinuo, nunca en un
color de serie, para que no compitan con los datos.
"""

import math
from collections import defaultdict
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import openpyxl

BASE_DIR = Path(__file__).resolve().parent.parent
RESULTADOS_FLUJO = BASE_DIR / "resultados" / "Resultados_Flujo_Carga.xlsx"
RESULTADOS_CORTO = BASE_DIR / "resultados" / "Resultados_Cortocircuito.xlsx"
SALIDA_DIR = BASE_DIR / "informe"

CASO_CARGA_PURA = "CargaPura"
CASO_BASE = "CasoBase"

# Paleta categorica validada para fondo claro (ver docstring). El orden es fijo.
COLOR = {
    "CargaPura": "#2a78d6",   # azul
    "CasoBase": "#eb6834",    # naranja
    "Proyecto": "#1baf7a",    # aguamarina
}
ETIQUETA = {
    "CargaPura": "Carga pura (sin generación)",
    "CasoBase": "Sin proyecto (solo GD1/GD2)",
    "Proyecto": "Con proyecto",
}
GRIS_LIMITE = "#8a8a85"
# El panel de balance mide EL EFECTO DEL PROYECTO, asi que sus barras llevan el
# color del proyecto: el color sigue a la entidad. El signo lo da la posicion
# respecto al cero, que es inequivoca y no necesita una segunda codificacion por
# color. Se evita asi que un mismo tono signifique "caso de red" en un panel y
# "signo" en otro, que es justo lo que confundiria al leer la figura completa.
GRIS_EJE = "#d5d5d0"
TINTA = "#0b0b0b"
TINTA_SEC = "#52514e"

CAPACIDAD_CORTE_KA = 10.0

# Barras y elementos que se grafican por nombre (ver cargar_flujo / figura_*)
BARRA_PC = "P1 15344 13.2kV"        # punto de conexion del proyecto (cabecera)
BARRA_EXTREMO = "P6 15344 13.2kV"   # extremo del circuito, maxima elevacion
LINEA_EVACUACION = "AL 1.32km"      # tramo que evacua la generacion del proyecto


def estilo_ejes(ax):
    """Ejes recesivos: el dato manda, la rejilla acompana."""
    ax.grid(True, axis="y", color=GRIS_EJE, linewidth=0.6, alpha=0.8)
    ax.set_axisbelow(True)
    for lado in ("top", "right"):
        ax.spines[lado].set_visible(False)
    for lado in ("left", "bottom"):
        ax.spines[lado].set_color(GRIS_EJE)
    ax.tick_params(colors=TINTA_SEC, labelsize=8, length=3, color=GRIS_EJE)
    ax.set_facecolor("white")


def rotulo_fila(fig, ax, texto):
    """Rotulo de fila al margen izquierdo de la figura, fuera del area de datos.

    Se usa en vez de un texto dentro del panel porque las lineas de limite
    regulatorio ocupan justamente la franja superior e inferior, y cualquier
    texto colocado ahi dentro acaba solapandose con ellas."""
    caja = ax.get_position()
    fig.text(0.008, caja.y0 + caja.height / 2, texto, rotation=90,
             va="center", ha="left", fontsize=8.5, color=TINTA_SEC)


def leer_hoja(wb, nombre):
    filas = list(wb[nombre].iter_rows(values_only=True))
    return list(filas[0]), filas[1:]


def clave_caso(nombre_caso):
    """Nombre de caso de PowerFactory -> clave de color/etiqueta."""
    if nombre_caso == CASO_CARGA_PURA:
        return "CargaPura"
    if nombre_caso == CASO_BASE:
        return "CasoBase"
    return "Proyecto"


def cargar_flujo():
    wb = openpyxl.load_workbook(RESULTADOS_FLUJO, data_only=True)

    h, nodos = leer_hoja(wb, "Nodos")
    iA, iC, iH, iU = h.index("Anio"), h.index("Caso"), h.index("Hora"), h.index("U_pu")
    # Se grafican DOS BARRAS CONCRETAS en vez de la envolvente min/max de todo el
    # sistema: la envolvente cambia de barra de una hora a otra (la barra que
    # marca el minimo no es siempre la misma), lo que produce una curva quebrada
    # que no corresponde al comportamiento de ningun elemento real. Las dos
    # barras elegidas son las que importan para este estudio: el punto de
    # conexion y el extremo del circuito, que es donde la inyeccion produce la
    # mayor elevacion de tension.
    u_pc, u_ext = {}, {}
    for r in nodos:
        if r[iU] is None:
            continue
        k = (r[iA], clave_caso(r[iC]), r[iH])
        if r[0] == BARRA_PC:
            u_pc[k] = r[iU]
        elif r[0] == BARRA_EXTREMO:
            u_ext[k] = r[iU]
    u_min, u_max = u_ext, u_pc

    h, lineas = leer_hoja(wb, "Lineas")
    iA, iC, iH, iL = h.index("Anio"), h.index("Caso"), h.index("Hora"), h.index("Cargabilidad_%")
    # Se grafica la LINEA DE EVACUACION del proyecto y no la mas cargada del
    # sistema: la mas cargada (Al 0.82km) pertenece a otro circuito del area y da
    # practicamente el mismo valor en los tres casos, de modo que las tres series
    # quedarian superpuestas y la grafica no mostraria nada. El tramo de
    # evacuacion es donde el efecto del proyecto es visible.
    carg_lin = {}
    carg_max_sistema = defaultdict(lambda: -math.inf)
    for r in lineas:
        if r[iL] is None:
            continue
        k = (r[iA], clave_caso(r[iC]), r[iH])
        if r[0] == LINEA_EVACUACION:
            carg_lin[k] = r[iL]
        carg_max_sistema[k] = max(carg_max_sistema[k], r[iL])

    h, tr2 = leer_hoja(wb, "Transformadores_2Devanados")
    iA, iC, iH, iL = h.index("Anio"), h.index("Caso"), h.index("Hora"), h.index("Cargabilidad_%")
    carg_tr = {}
    for r in tr2:
        if r[iL] is not None:
            carg_tr[(r[iA], clave_caso(r[iC]), r[iH])] = r[iL]

    perdidas = defaultdict(float)
    for hoja in ("Lineas", "Transformadores_2Devanados", "Transformadores_3Devanados"):
        h, filas = leer_hoja(wb, hoja)
        if not filas:
            continue
        iA, iC, iH = h.index("Anio"), h.index("Caso"), h.index("Hora")
        iP = h.index("P_perdidas_MW")
        for r in filas:
            if r[iP] is not None:
                perdidas[(r[iA], clave_caso(r[iC]), r[iH])] += r[iP] * 1000.0  # kW

    anios = sorted({k[0] for k in u_min})
    horas = sorted({k[2] for k in u_min})
    return anios, horas, u_min, u_max, carg_lin, carg_tr, perdidas, carg_max_sistema


def figura_tension(anios, horas, u_min, u_max, ruta):
    """Envolvente de tension del sistema: la minima y la maxima de todas las
    barras, hora a hora. Es la vista que permite verificar cumplimiento de un
    vistazo, porque si la envolvente no sale de la banda, ninguna barra lo hace."""
    fig, axes = plt.subplots(2, len(anios), figsize=(11, 6.4), sharex=True, sharey="row")
    fig.patch.set_facecolor("white")

    for col, anio in enumerate(anios):
        for fila, (datos, titulo) in enumerate((
            (u_max, f"Punto de conexión ({BARRA_PC})"),
            (u_min, f"Extremo del circuito ({BARRA_EXTREMO})"),
        )):
            ax = axes[fila][col]
            estilo_ejes(ax)
            for caso in ("CargaPura", "CasoBase", "Proyecto"):
                y = [datos.get((anio, caso, hh)) for hh in horas]
                if all(v is None or v == math.inf or v == -math.inf for v in y):
                    continue
                ax.plot(horas, y, color=COLOR[caso], linewidth=2.0,
                        marker="o", markersize=4, markeredgecolor="white",
                        markeredgewidth=0.8, label=ETIQUETA[caso], zorder=3)
            for limite in (0.90, 1.10):
                ax.axhline(limite, color=GRIS_LIMITE, linewidth=1.2, linestyle="--", zorder=2)
                ax.annotate(f"Límite {limite:.2f} p.u.", xy=(horas[-1], limite),
                            xytext=(-4, 4), textcoords="offset points", ha="right",
                            fontsize=7.5, color=TINTA_SEC)
            if col == 0:
                ax.set_ylabel("Tensión (p.u.)", fontsize=9, color=TINTA)
            if fila == 0:
                ax.set_title(f"Año {anio}", fontsize=11, color=TINTA, fontweight="bold", pad=10)
            if fila == 1:
                ax.set_xlabel("Hora del día", fontsize=9, color=TINTA)

    # el eje abarca toda la banda regulatoria, para que la holgura real del
    # proyecto frente a los limites se lea de un vistazo y no quede exagerada
    # por un zoom sobre los datos
    for fila in (0, 1):
        axes[fila][0].set_ylim(0.88, 1.12)

    manejadores, etiquetas = axes[0][0].get_legend_handles_labels()
    fig.legend(manejadores, etiquetas, loc="lower center", ncol=3, frameon=False,
               fontsize=9, bbox_to_anchor=(0.5, -0.01))
    fig.suptitle("Perfiles de tensión en las barras críticas, horas 6 a 17",
                 fontsize=12.5, color=TINTA, fontweight="bold", y=0.98)
    fig.tight_layout(rect=[0.035, 0.05, 1, 0.95])
    rotulo_fila(fig, axes[0][0], f"Punto de conexión ({BARRA_PC})")
    rotulo_fila(fig, axes[1][0], f"Extremo del circuito ({BARRA_EXTREMO})")
    fig.savefig(ruta, dpi=200, facecolor="white")
    plt.close(fig)
    return ruta


def figura_cargabilidad(anios, horas, carg_lin, carg_tr, carg_max_sistema, ruta):
    fig, axes = plt.subplots(2, len(anios), figsize=(11, 6.4), sharex=True, sharey="row")
    fig.patch.set_facecolor("white")

    for col, anio in enumerate(anios):
        # fila 0: lineas
        ax = axes[0][col]
        estilo_ejes(ax)
        for caso in ("CargaPura", "CasoBase", "Proyecto"):
            y = [carg_lin.get((anio, caso, hh)) for hh in horas]
            if all(v is None for v in y):
                continue
            ax.plot(horas, y, color=COLOR[caso], linewidth=2.0, marker="o",
                    markersize=4, markeredgecolor="white", markeredgewidth=0.8,
                    label=ETIQUETA[caso], zorder=3)
        ax.set_title(f"Año {anio}", fontsize=11, color=TINTA, fontweight="bold", pad=10)
        pico = max(carg_max_sistema.get((anio, c, hh), 0)
                   for c in ("CargaPura", "CasoBase", "Proyecto") for hh in horas)
        ax.annotate(f"Elemento más cargado de todo el sistema: {pico:.1f} %",
                    xy=(0.015, 0.80), xycoords="axes fraction",
                    fontsize=7.5, color=TINTA_SEC)
        if col == 0:
            ax.set_ylabel("Cargabilidad (%)", fontsize=9, color=TINTA)

        # fila 1: transformador del proyecto (solo existe en el caso con proyecto)
        ax = axes[1][col]
        estilo_ejes(ax)
        y = [carg_tr.get((anio, "Proyecto", hh)) for hh in horas]
        ax.plot(horas, y, color=COLOR["Proyecto"], linewidth=2.0, marker="o",
                markersize=4, markeredgecolor="white", markeredgewidth=0.8, zorder=3)
        ax.set_xlabel("Hora del día", fontsize=9, color=TINTA)
        if col == 0:
            ax.set_ylabel("Cargabilidad (%)", fontsize=9, color=TINTA)

    for fila in (0, 1):
        for col in range(len(anios)):
            axes[fila][col].axhline(100, color=GRIS_LIMITE, linewidth=1.2, linestyle="--", zorder=2)
        axes[fila][0].annotate("Límite 100 %", xy=(horas[0], 100),
                               xytext=(0, 4), textcoords="offset points",
                               fontsize=7.5, color=TINTA_SEC)
    axes[0][0].set_ylim(0, 108)
    axes[1][0].set_ylim(0, 108)

    manejadores, etiquetas = axes[0][0].get_legend_handles_labels()
    fig.legend(manejadores, etiquetas, loc="lower center", ncol=3, frameon=False,
               fontsize=9, bbox_to_anchor=(0.5, -0.01))
    fig.suptitle("Nivel de carga de los elementos — condición normal (N)",
                 fontsize=12.5, color=TINTA, fontweight="bold", y=0.98)
    fig.tight_layout(rect=[0.035, 0.05, 1, 0.95])
    rotulo_fila(fig, axes[0][0], f"Línea de evacuación ({LINEA_EVACUACION})")
    rotulo_fila(fig, axes[1][0], "Transformador del proyecto (1250 kVA)")
    fig.savefig(ruta, dpi=200, facecolor="white")
    plt.close(fig)
    return ruta


def figura_perdidas(anios, horas, perdidas, ruta):
    """Dos vistas: el nivel de perdidas hora a hora, y el balance del proyecto
    frente al escenario de carga pura (que es la comparacion que responde a la
    pregunta del numeral 8.6: que le aporta la generacion distribuida a la red)."""
    fig, axes = plt.subplots(2, len(anios), figsize=(11, 6.4), sharex=True, sharey="row")
    fig.patch.set_facecolor("white")

    for col, anio in enumerate(anios):
        ax = axes[0][col]
        estilo_ejes(ax)
        for caso in ("CargaPura", "CasoBase", "Proyecto"):
            y = [perdidas.get((anio, caso, hh)) for hh in horas]
            if all(v is None for v in y):
                continue
            ax.plot(horas, y, color=COLOR[caso], linewidth=2.0, marker="o",
                    markersize=4, markeredgecolor="white", markeredgewidth=0.8,
                    label=ETIQUETA[caso], zorder=3)
        ax.set_title(f"Año {anio}", fontsize=11, color=TINTA, fontweight="bold", pad=10)
        if col == 0:
            ax.set_ylabel("Pérdidas (kW)", fontsize=9, color=TINTA)

        # balance frente a carga pura
        ax = axes[1][col]
        estilo_ejes(ax)
        delta = [perdidas.get((anio, "Proyecto", hh), 0) - perdidas.get((anio, "CargaPura", hh), 0)
                 for hh in horas]
        ax.bar(horas, delta, color=COLOR["Proyecto"], width=0.62, zorder=3)
        ax.axhline(0, color=GRIS_LIMITE, linewidth=1.0, zorder=2)
        ax.set_xlabel("Hora del día", fontsize=9, color=TINTA)
        if col == 0:
            ax.set_ylabel("Δ pérdidas (kW)", fontsize=9, color=TINTA)
            ax.annotate("por debajo de cero = el proyecto reduce las pérdidas\ndel sistema respecto al escenario sin generación",
                        xy=(0.015, 0.06), xycoords="axes fraction",
                        fontsize=7.5, color=TINTA_SEC)

    manejadores, etiquetas = axes[0][0].get_legend_handles_labels()
    fig.legend(manejadores, etiquetas, loc="lower center", ncol=3, frameon=False,
               fontsize=9, bbox_to_anchor=(0.5, -0.01))
    fig.suptitle("Análisis de pérdidas del sistema",
                 fontsize=12.5, color=TINTA, fontweight="bold", y=0.98)
    fig.tight_layout(rect=[0.035, 0.05, 1, 0.95])
    rotulo_fila(fig, axes[0][0], "Pérdidas activas totales")
    rotulo_fila(fig, axes[1][0], "Balance frente a carga pura")
    fig.savefig(ruta, dpi=200, facecolor="white")
    plt.close(fig)
    return ruta


def figura_cortocircuito(ruta):
    wb = openpyxl.load_workbook(RESULTADOS_CORTO, data_only=True)
    h, filas = leer_hoja(wb, "Nodos")
    iC, iB, iT, iI = h.index("Caso"), h.index("Barra_Falla"), h.index("Tipo_Falla"), h.index("Ikss_kA")
    d = {(r[iC], r[iB], r[iT]): r[iI] for r in filas}
    caso_gd = next(c for c in {k[0] for k in d} if c != CASO_BASE)

    # solo barras de la red del OR (la de 800 V es de la instalacion del proyecto
    # y se rige por otro criterio, por eso se excluye de la comparacion vs 10 kA)
    barras = sorted(
        {k[1] for k in d if d.get((CASO_BASE, k[1], "Trifasico")) is not None},
        key=lambda b: d.get((CASO_BASE, b, "Trifasico"), 0),
    )

    fig, axes = plt.subplots(1, 2, figsize=(11, 6.2),
                            gridspec_kw={"width_ratios": [1.35, 1]})
    fig.patch.set_facecolor("white")

    # --- panel izquierdo: nivel de falla por barra, sin vs con proyecto ---
    ax = axes[0]
    estilo_ejes(ax)
    ax.grid(True, axis="x", color=GRIS_EJE, linewidth=0.6, alpha=0.8)
    ax.grid(False, axis="y")
    y = range(len(barras))
    sin = [d.get((CASO_BASE, b, "Trifasico")) for b in barras]
    con = [d.get((caso_gd, b, "Trifasico")) for b in barras]
    ax.barh([v + 0.19 for v in y], sin, height=0.34, color=COLOR["CasoBase"],
            label="Sin proyecto", zorder=3)
    ax.barh([v - 0.19 for v in y], con, height=0.34, color=COLOR["Proyecto"],
            label="Con proyecto", zorder=3)
    ax.axvline(CAPACIDAD_CORTE_KA, color=GRIS_LIMITE, linewidth=1.4, linestyle="--", zorder=4)
    ax.annotate(f"Capacidad de corte {CAPACIDAD_CORTE_KA:.0f} kA",
                xy=(CAPACIDAD_CORTE_KA, len(barras) - 0.6), xytext=(-6, 0),
                textcoords="offset points", fontsize=8, color=TINTA_SEC,
                ha="right", rotation=90, va="top")
    ax.set_yticks(list(y))
    ax.set_yticklabels(barras, fontsize=8)
    ax.set_xlabel("Corriente de cortocircuito trifásica $I''_{k}$ (kA)",
                  fontsize=9, color=TINTA)
    ax.set_xlim(0, 11.2)
    ax.legend(frameon=False, fontsize=8.5, loc="lower right")
    ax.set_title("Nivel de falla por barra", fontsize=11, color=TINTA,
                 fontweight="bold", pad=10)

    # --- panel derecho: incremento porcentual que aporta el proyecto ---
    ax = axes[1]
    estilo_ejes(ax)
    ax.grid(True, axis="x", color=GRIS_EJE, linewidth=0.6, alpha=0.8)
    ax.grid(False, axis="y")
    delta = [100.0 * (c - s) / s if s else 0 for s, c in zip(sin, con)]
    ax.barh(list(y), delta, height=0.55, color=COLOR["Proyecto"], zorder=3)
    ax.set_yticks(list(y))
    ax.set_yticklabels([])
    ax.set_xlabel("Incremento aportado por el proyecto (%)", fontsize=9, color=TINTA)
    ax.set_title("Aporte del proyecto", fontsize=11, color=TINTA, fontweight="bold", pad=10)
    for i, v in zip(y, delta):
        ax.annotate(f"{v:+.2f} %", xy=(v, i), xytext=(4, 0), textcoords="offset points",
                    fontsize=7.5, color=TINTA_SEC, va="center")
    ax.set_xlim(0, max(delta) * 1.45 if delta else 1)

    fig.suptitle("Contribución a la corriente de cortocircuito — falla trifásica",
                 fontsize=12.5, color=TINTA, fontweight="bold", y=0.975)
    fig.tight_layout(rect=[0, 0, 1, 0.94])
    fig.savefig(ruta, dpi=200, facecolor="white")
    plt.close(fig)
    return ruta


def main():
    SALIDA_DIR.mkdir(parents=True, exist_ok=True)
    anios, horas, u_min, u_max, carg_lin, carg_tr, perdidas, carg_max_sistema = cargar_flujo()
    print(f"Anos: {anios} | horas: {horas[0]}-{horas[-1]}")

    rutas = [
        figura_tension(anios, horas, u_min, u_max, SALIDA_DIR / "fig_tension.png"),
        figura_cargabilidad(anios, horas, carg_lin, carg_tr, carg_max_sistema,
                            SALIDA_DIR / "fig_cargabilidad.png"),
        figura_perdidas(anios, horas, perdidas, SALIDA_DIR / "fig_perdidas.png"),
        figura_cortocircuito(SALIDA_DIR / "fig_cortocircuito.png"),
    ]
    for r in rutas:
        print(f"  generada: {r.relative_to(BASE_DIR)}")


if __name__ == "__main__":
    main()
