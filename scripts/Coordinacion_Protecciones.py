"""
Comprobacion de coordinacion de proteccion de cabecera del circuito 15344
(EACP simplificado) - CREG 174 de 2021 numeral 9.2 / Acuerdo CNO 2121 de 2026.

MODULO DE POST-PROCESO PURO: no se conecta a PowerFactory, no lee ni escribe
nada del modelo de red. Se corre como script de Python normal (no dentro de
PowerFactory), igual que el script que genera curva_demanda_generacion.png.
Solo lee:
  - resultados/Resultados_Cortocircuito.xlsx (ya generado por Cortocircuito.py
    dentro de PowerFactory) - hoja "Nodos", para la Ikss de la barra de
    cabecera del circuito 15344, en CasoBase (sin proyecto) y con la Network
    Variation del proyecto activa (con proyecto).
  - Los ajustes REALES de la proteccion de cabecera del circuito 15344 (curva
    IEC Normal Inverse ya aprobada y en operacion por EBSA), tomados de
    DOC_REF_EST_1787069892338.pdf (pag. 4, "CURVA AJUSTADA EN LA PROTECCION
    DEL CIRCUITO Y RECONECTADOR") y ya registrados en
    informe/Parametros_Proyecto.xlsx (fila Condiciones_operativas). Estos son
    datos de placa fijos del reconectador, no resultados de simulacion, por
    eso van hardcodeados abajo (con su fuente documentada).

Alcance (leer antes de usar): esto es una COMPROBACION puntual de si el
ajuste ya existente en cabecera del circuito 15344 sigue coordinando bien al
agregar el aporte de falla del proyecto - NO es un EACP completo (que
cubriria ademas las funciones de proteccion internas de la planta en todos
los niveles de tension). Ver README.md, seccion "Comprobacion de
Coordinacion de Protecciones (EACP simplificado)", para la justificacion
completa de por que este alcance es adecuado para un proyecto de este
tamano.

Metodologia (formulas):
  1. Curva IEC Normal Inverse (IEC 60255-151), la misma familia de curva con
     la que EBSA ya ajusto la proteccion de cabecera ("Curva IEC Norm
     Inver" en la ficha de insumo):

         t(I) = Dial * k / [ (I / Ipickup)^alpha - 1 ]      (I > Ipickup)

     con k = 0.14 y alpha = 0.02 (constantes de la norma para esta familia
     de curva, no calibrables). Si la corriente supera el umbral
     instantaneo (50 de fase / 50N de neutro), el rele no espera la curva
     temporizada: dispara en un tiempo definido muy corto (aqui se usa
     TIEMPO_MIN_INSTANTANEO_S como aproximacion, tipico de un reconectador).
  2. Para cada caso (CasoBase / con proyecto) y tipo de falla (trifasico ->
     funcion 51/50 de fase; monofasico -> funcion 51N/50N de neutro), se
     evalua la Ikss de PowerFactory en la formula de arriba y se compara el
     tiempo de disparo resultante y la zona de operacion (temporizada vs.
     instantanea).
  3. Si ambos casos (sin y con proyecto) caen en la MISMA zona de operacion
     con el MISMO tiempo de disparo (dentro de un margen despreciable), la
     conclusion es que el aporte de falla del proyecto NO altera la
     coordinacion ya aprobada por EBSA en cabecera.
  4. Aparte, se documenta el aporte de falla propio del sistema de
     generacion (el configurado en PowerFactory y el resultado de
     simulacion ya validado) como insumo para el diseno de la malla de
     puesta a tierra (el Anexo 1 del Acuerdo 2121 liga explicitamente el
     diseno de puesta a tierra con el sistema de protecciones propuesto).

Salidas (todo dentro de resultados/Coordinacion_de_Protecciones/):
  - Curva_TCC_Cabecera_15344.png: curva tiempo-corriente (log-log) de fase y
    neutro de la proteccion de cabecera, con los 4 puntos evaluados marcados.
  - Resultados_Coordinacion_Protecciones.xlsx: tabla comparativa sin/con
    proyecto, el aporte de falla del sistema de generacion, y la conclusion.
"""

from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import openpyxl
from openpyxl.styles import Font

BASE_DIR = Path(__file__).resolve().parent.parent
RESULTADOS_CORTO_PATH = BASE_DIR / "resultados" / "Resultados_Cortocircuito.xlsx"
CARPETA_SALIDA = BASE_DIR / "resultados" / "Coordinacion_de_Protecciones"
GRAFICOS_DIR = CARPETA_SALIDA
RESULTADOS_PATH = CARPETA_SALIDA / "Resultados_Coordinacion_Protecciones.xlsx"

CASO_BASE = "CasoBase"

# --- Ajustes REALES de la proteccion de cabecera del circuito 15344 -------
# Fuente: DOC_REF_EST_1787069892338.pdf (pag. 4) / Parametros_Proyecto.xlsx,
# fila "Condiciones_operativas". Curva IEC Normal Inverse (IEC 60255-151).
PICKUP_51_A = 240.0     # sobrecorriente de fase, temporizada (51)
DIAL_51 = 0.1
INST_50_A = 1500.0      # instantaneo de fase (50)
PICKUP_51N_A = 60.0     # sobrecorriente de neutro/tierra, temporizada (51N)
DIAL_51N = 0.1
INST_50N_A = 300.0      # instantaneo de neutro/tierra (50N)
TIEMPO_RECIERRE_RAPIDO_S = 2.0

K_IEC_NI = 0.14
ALPHA_IEC_NI = 0.02
TIEMPO_MIN_INSTANTANEO_S = 0.03  # tiempo de operacion tipico de un instantaneo/reconectador

# Barra que representa la cabecera del circuito 15344 (primer punto del
# feeder, justo aguas abajo del reconectador "Cab 15344"). AJUSTAR aqui si
# el nombre exacto de la barra en el modelo de PowerFactory cambia.
BARRA_CABECERA = "P1 15344 13.2kV"

# --- Aporte de falla propio del sistema de generacion, ya validado --------
# Fuente: Parametros_Proyecto.xlsx, filas "Metodo_calculo_corto_detallado" /
# "Confirmacion_fix_corto_generador".
IK3PF_CONFIGURADO_A_POR_INVERSOR = 238.2   # Ik"3PF configurado en cada ElmPvsys
N_INVERSORES = 3
POTENCIA_INVERSOR_KW = 330.0
TENSION_AC_INVERSOR_V = 800.0
IKSS_GENERADOR_VALIDADO_KA = 0.7146        # Trifasico, barra propia del generador (Cortocircuito.py)
LIMITE_REGULATORIO_PU = 1.1                # Anexo 1, Acuerdo CNO 2121 de 2026


def tiempo_disparo_ni(i_a, pickup_a, dial, inst_a):
    """Tiempo de disparo (s) de una funcion de sobrecorriente IEC Normal
    Inverse con instantaneo asociado. None si no dispara (I <= pickup)."""
    if i_a >= inst_a:
        return TIEMPO_MIN_INSTANTANEO_S, "Instantaneo"
    if i_a <= pickup_a:
        return None, "No dispara"
    t = dial * K_IEC_NI / ((i_a / pickup_a) ** ALPHA_IEC_NI - 1)
    return t, "Temporizado"


def curva_tcc(pickup_a, dial, inst_a, i_max_a):
    """Puntos (I, t) de la curva temporizada, desde justo encima del pickup
    hasta el instantaneo, para graficar la rama IEC Normal Inverse."""
    i_vals = np.logspace(np.log10(pickup_a * 1.01), np.log10(inst_a), 200)
    t_vals = [dial * K_IEC_NI / ((i / pickup_a) ** ALPHA_IEC_NI - 1) for i in i_vals]
    return i_vals, t_vals


def leer_ikss_cabecera():
    """Lee Resultados_Cortocircuito.xlsx y devuelve
    {(caso_normalizado, tipo_falla): Ikss_kA} para BARRA_CABECERA.
    caso_normalizado es 'CasoBase' o 'Con_Proyecto' (cualquier caso distinto
    de CasoBase se agrupa como 'Con_Proyecto', igual que en las hojas
    Generadores_Con_Proyecto de Flujo_carga.py / Cortocircuito.py)."""
    if not RESULTADOS_CORTO_PATH.exists():
        raise FileNotFoundError(
            f"No se encontro {RESULTADOS_CORTO_PATH}. Corre primero Cortocircuito.py dentro de PowerFactory."
        )
    wb = openpyxl.load_workbook(RESULTADOS_CORTO_PATH, data_only=True)
    ws = wb["Nodos"]
    headers = [c.value for c in ws[1]]
    idx = {h: i for i, h in enumerate(headers)}

    datos = {}
    variaciones_encontradas = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[idx["Nombre"]] != BARRA_CABECERA:
            continue
        caso = row[idx["Caso"]]
        caso_norm = CASO_BASE if caso == CASO_BASE else "Con_Proyecto"
        if caso_norm == "Con_Proyecto":
            variaciones_encontradas.add(caso)
        tipo = row[idx["Tipo_Falla"]]
        ikss = row[idx["Ikss_kA"]]
        # si hay mas de una Network Variation "con proyecto", se deja el
        # primer valor encontrado y se avisa - este modulo asume una sola
        # variation representando "con proyecto conectado"
        datos.setdefault((caso_norm, tipo), ikss)

    if len(variaciones_encontradas) > 1:
        print(
            f"AVISO: se encontro mas de una Network Variation con datos en {BARRA_CABECERA} "
            f"({variaciones_encontradas}) - se uso la primera encontrada para 'Con_Proyecto'."
        )
    if not datos:
        raise ValueError(
            f"No hay filas para la barra '{BARRA_CABECERA}' en la hoja Nodos de {RESULTADOS_CORTO_PATH}. "
            "Revisa BARRA_CABECERA (nombre exacto de la barra en el modelo)."
        )
    return datos


def graficar_tcc(comparativa):
    GRAFICOS_DIR.mkdir(parents=True, exist_ok=True)

    fig, ax = plt.subplots(figsize=(9, 6.5))

    # curva de fase (51/50)
    i51, t51 = curva_tcc(PICKUP_51_A, DIAL_51, INST_50_A, INST_50_A)
    ax.plot(i51, t51, color="tab:blue", label="Fase 51 (temporizada)")
    ax.axvline(INST_50_A, color="tab:blue", linestyle="--", linewidth=1, label="Fase 50 (instantaneo, 1500A)")

    # curva de neutro (51N/50N)
    i51n, t51n = curva_tcc(PICKUP_51N_A, DIAL_51N, INST_50N_A, INST_50N_A)
    ax.plot(i51n, t51n, color="tab:orange", label="Neutro 51N (temporizada)")
    ax.axvline(INST_50N_A, color="tab:orange", linestyle="--", linewidth=1, label="Neutro 50N (instantaneo, 300A)")

    marcadores = {
        ("CasoBase", "Trifasico"): ("o", "tab:blue", "Trifasico sin proyecto"),
        ("Con_Proyecto", "Trifasico"): ("s", "tab:blue", "Trifasico con proyecto"),
        ("CasoBase", "Monofasico"): ("o", "tab:orange", "Monofasico sin proyecto"),
        ("Con_Proyecto", "Monofasico"): ("s", "tab:orange", "Monofasico con proyecto"),
    }
    for fila in comparativa:
        clave = (fila["caso"], fila["tipo_falla"])
        if clave not in marcadores:
            continue
        marker, color, etiqueta = marcadores[clave]
        ax.scatter(
            [fila["ikss_a"]], [fila["tiempo_s"]], marker=marker, color=color,
            edgecolor="black", zorder=5, s=90, label=etiqueta,
        )

    ax.set_xscale("log")
    ax.set_yscale("log")
    ax.set_xlabel("Corriente (A, primario)")
    ax.set_ylabel("Tiempo de disparo (s)")
    ax.set_title("Curva TCC - Proteccion de cabecera circuito 15344 (IEC Normal Inverse)")
    ax.grid(True, which="both", linestyle=":", linewidth=0.5)

    handles, labels = ax.get_legend_handles_labels()
    vistos = dict(zip(labels, handles))
    ax.legend(vistos.values(), vistos.keys(), fontsize=8, loc="upper right")

    # los 4 puntos evaluados caen muy cerca entre si en la escala log del
    # grafico (el aporte del proyecto es marginal frente al nivel de falla
    # de la red) - se agrega una tabla de texto con los valores exactos para
    # que la comparacion sea legible aunque los marcadores se superpongan.
    texto = "Valores evaluados (A primario):\n" + "\n".join(
        f"  {f['tipo_falla']} {f['caso'].replace('_', ' ')}: {f['ikss_a']:.0f} A"
        for f in comparativa
    )
    ax.text(
        0.02, 0.03, texto, transform=ax.transAxes, fontsize=8, va="bottom", ha="left",
        bbox=dict(boxstyle="round", facecolor="white", alpha=0.85, edgecolor="gray"),
    )

    ruta = GRAFICOS_DIR / "Curva_TCC_Cabecera_15344.png"
    fig.tight_layout()
    fig.savefig(ruta, dpi=150)
    plt.close(fig)
    return ruta


def construir_comparativa(ikss_por_caso):
    ajustes = {
        "Trifasico": (PICKUP_51_A, DIAL_51, INST_50_A, "51/50 (fase)"),
        "Monofasico": (PICKUP_51N_A, DIAL_51N, INST_50N_A, "51N/50N (neutro)"),
    }
    filas = []
    base_por_tipo = {}
    for tipo, (pickup, dial, inst, funcion) in ajustes.items():
        for caso in (CASO_BASE, "Con_Proyecto"):
            ikss_kA = ikss_por_caso.get((caso, tipo))
            if ikss_kA is None:
                continue
            ikss_a = ikss_kA * 1000.0
            tiempo_s, zona = tiempo_disparo_ni(ikss_a, pickup, dial, inst)
            if caso == CASO_BASE:
                base_por_tipo[tipo] = ikss_a
            filas.append({
                "tipo_falla": tipo,
                "funcion": funcion,
                "caso": caso,
                "ikss_kA": ikss_kA,
                "ikss_a": ikss_a,
                "zona": zona,
                "tiempo_s": tiempo_s,
            })
    for fila in filas:
        base_a = base_por_tipo.get(fila["tipo_falla"])
        if base_a and fila["caso"] != CASO_BASE:
            fila["delta_a"] = fila["ikss_a"] - base_a
            fila["delta_pct"] = 100.0 * fila["delta_a"] / base_a
        else:
            fila["delta_a"] = None
            fila["delta_pct"] = None
    return filas


def evaluar_conclusion(comparativa):
    """Compara zona/tiempo de disparo sin vs. con proyecto, por tipo de
    falla, y arma la conclusion en texto."""
    por_tipo = {}
    for fila in comparativa:
        por_tipo.setdefault(fila["tipo_falla"], {})[fila["caso"]] = fila

    lineas = []
    requiere_ajuste = False
    for tipo, casos in por_tipo.items():
        base = casos.get(CASO_BASE)
        con_proy = casos.get("Con_Proyecto")
        if not base or not con_proy:
            continue
        mismo_zona = base["zona"] == con_proy["zona"]
        delta_pct = con_proy["delta_pct"] or 0.0
        if mismo_zona and abs(delta_pct) < 5.0:
            lineas.append(
                f"- Falla {tipo.lower()} en cabecera 15344: {base['ikss_a']:.0f}A -> {con_proy['ikss_a']:.0f}A "
                f"({delta_pct:+.2f}%), ambos en zona '{base['zona']}' con el mismo tiempo de disparo "
                f"({con_proy['tiempo_s']:.3f}s) -> NO requiere reajuste."
            )
        else:
            requiere_ajuste = True
            lineas.append(
                f"- Falla {tipo.lower()} en cabecera 15344: {base['ikss_a']:.0f}A -> {con_proy['ikss_a']:.0f}A "
                f"({delta_pct:+.2f}%), cambia de zona '{base['zona']}' a '{con_proy['zona']}' o el tiempo "
                f"de disparo varia de forma significativa -> REVISAR pickup/dial de esta funcion."
            )
    conclusion_general = (
        "No se requiere reajuste de la proteccion de cabecera del circuito 15344: el aporte de falla "
        "del proyecto es marginal frente al nivel de cortocircuito de la red de EBSA y ambos escenarios "
        "(sin y con proyecto) siguen operando en la misma zona de la curva ya aprobada."
        if not requiere_ajuste else
        "Se recomienda revisar el ajuste de cabecera del circuito 15344 con el OR antes de energizar el proyecto."
    )
    return "\n".join(lineas), conclusion_general, requiere_ajuste


def aporte_falla_generacion():
    corriente_nominal_inversor_a = (POTENCIA_INVERSOR_KW * 1000.0) / (np.sqrt(3) * TENSION_AC_INVERSOR_V)
    pu_configurado = IK3PF_CONFIGURADO_A_POR_INVERSOR / corriente_nominal_inversor_a
    return {
        "ik3pf_configurado_a_por_inversor": IK3PF_CONFIGURADO_A_POR_INVERSOR,
        "n_inversores": N_INVERSORES,
        "corriente_nominal_inversor_a": corriente_nominal_inversor_a,
        "pu_configurado": pu_configurado,
        "limite_regulatorio_pu": LIMITE_REGULATORIO_PU,
        "ikss_validado_kA": IKSS_GENERADOR_VALIDADO_KA,
    }


def exportar_resultados(comparativa, texto_conclusion, conclusion_general, aporte):
    wb = openpyxl.Workbook()

    ws1 = wb.active
    ws1.title = "Comparativa_Cabecera_15344"
    headers = ["Tipo_Falla", "Funcion", "Caso", "Ikss_kA", "Ikss_A", "Zona_Operacion", "Tiempo_disparo_s", "Delta_vs_SinProyecto_A", "Delta_vs_SinProyecto_%"]
    ws1.append(headers)
    for cell in ws1[1]:
        cell.font = Font(bold=True)
    for fila in comparativa:
        ws1.append([
            fila["tipo_falla"], fila["funcion"], fila["caso"], fila["ikss_kA"], fila["ikss_a"],
            fila["zona"], fila["tiempo_s"], fila["delta_a"], fila["delta_pct"],
        ])
    ws1.auto_filter.ref = ws1.dimensions
    for col in ws1.columns:
        max_len = max(len(str(c.value)) if c.value is not None else 0 for c in col)
        ws1.column_dimensions[col[0].column_letter].width = max_len + 2

    ws2 = wb.create_sheet("Aporte_Falla_Generacion")
    ws2.append(["Parametro", "Valor", "Unidad", "Fuente"])
    for cell in ws2[1]:
        cell.font = Font(bold=True)
    ws2.append(["Ik\"3PF configurado por inversor", aporte["ik3pf_configurado_a_por_inversor"], "A", "PowerFactory - ElmPvsys, pestana Short-Circuit VDE/IEC"])
    ws2.append(["Cantidad de inversores", aporte["n_inversores"], "-", "Ficha tecnica Huawei SUN2000-330KTL-H1"])
    ws2.append(["Corriente nominal del inversor (calculada)", aporte["corriente_nominal_inversor_a"], "A", "330kW / (raiz(3) x 800V), cos(phi)=1"])
    ws2.append(["Aporte configurado en p.u. de la corriente nominal", aporte["pu_configurado"], "p.u.", "Calculado"])
    ws2.append(["Limite regulatorio (Anexo 1, Acuerdo CNO 2121/2026)", aporte["limite_regulatorio_pu"], "p.u.", "Anexo_1_Acuerdo_2121.pdf"])
    ws2.append(["Ikss trifasico validado en barra propia del generador", aporte["ikss_validado_kA"], "kA", "Resultados_Cortocircuito.xlsx (corrida real)"])
    ws2.append([
        "Nota", "Este valor (corriente de falla trifasica real del sistema de generacion, segun la "
        "simulacion) es el insumo a usar para el diseno de la malla de puesta a tierra del proyecto.", "-", "-",
    ])
    for col in ws2.columns:
        max_len = max(len(str(c.value)) if c.value is not None else 0 for c in col)
        ws2.column_dimensions[col[0].column_letter].width = min(max_len + 2, 100)

    ws3 = wb.create_sheet("Conclusion")
    ws3.append(["Detalle por tipo de falla"])
    ws3["A1"].font = Font(bold=True)
    for linea in texto_conclusion.split("\n"):
        ws3.append([linea])
    ws3.append([])
    ws3.append(["Conclusion general"])
    ws3[f"A{ws3.max_row}"].font = Font(bold=True)
    ws3.append([conclusion_general])
    for col in ws3.columns:
        ws3.column_dimensions[col[0].column_letter].width = 120

    try:
        wb.save(RESULTADOS_PATH)
    except PermissionError:
        alterna = RESULTADOS_PATH.with_name(f"{RESULTADOS_PATH.stem}_bloqueado{RESULTADOS_PATH.suffix}")
        wb.save(alterna)
        print(f"'{RESULTADOS_PATH.name}' esta abierto en Excel - se guardo como '{alterna.name}'. Cierra el original y corre de nuevo.")
        return alterna
    return RESULTADOS_PATH


def main():
    CARPETA_SALIDA.mkdir(parents=True, exist_ok=True)
    ikss_por_caso = leer_ikss_cabecera()
    comparativa = construir_comparativa(ikss_por_caso)
    texto_conclusion, conclusion_general, _ = evaluar_conclusion(comparativa)
    aporte = aporte_falla_generacion()

    ruta_grafico = graficar_tcc(comparativa)
    ruta_excel = exportar_resultados(comparativa, texto_conclusion, conclusion_general, aporte)

    print("=== Comprobacion de coordinacion de protecciones - cabecera circuito 15344 ===")
    for fila in comparativa:
        print(
            f"{fila['tipo_falla']:>10} | {fila['caso']:>13} | {fila['ikss_a']:>8.1f} A | "
            f"{fila['zona']:>12} | t={fila['tiempo_s']:.3f}s" if fila["tiempo_s"] is not None
            else f"{fila['tipo_falla']:>10} | {fila['caso']:>13} | {fila['ikss_a']:>8.1f} A | {fila['zona']:>12}"
        )
    print()
    print(texto_conclusion)
    print()
    print("CONCLUSION:", conclusion_general)
    print()
    print(f"Grafico: {ruta_grafico}")
    print(f"Excel:   {ruta_excel}")


if __name__ == "__main__":
    main()
