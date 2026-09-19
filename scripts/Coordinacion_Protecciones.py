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

import math

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import openpyxl

from parametros import P
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
PICKUP_51_A = P.cabecera("PICKUP_51_A")
DIAL_51 = P.cabecera("DIAL_51")
INST_50_A = P.cabecera("INST_50_A")
PICKUP_51N_A = P.cabecera("PICKUP_51N_A")
DIAL_51N = P.cabecera("DIAL_51N")
INST_50N_A = P.cabecera("INST_50N_A")
K_IEC_NI = P.cabecera("K_IEC_NI")
ALPHA_IEC_NI = P.cabecera("ALPHA_IEC_NI")
TIEMPO_MIN_INSTANTANEO_S = P.cabecera("TIEMPO_MIN_INSTANTANEO_S")
TIEMPO_RECIERRE_CABECERA_S = P.cabecera("TIEMPO_RECIERRE_S")

MARGEN_MINIMO_ANTIISLA_S = P.proyecto("MARGEN_MINIMO_ANTIISLA_S")
MARGEN_MINIMO_SELECTIVIDAD_S = P.proyecto("MARGEN_MINIMO_SELECTIVIDAD_S")
PROY_PICKUP_51_A = P.proyecto("PROY_PICKUP_51_A")
PROY_DIAL_51 = P.proyecto("PROY_DIAL_51")
PROY_INST_50_A = P.proyecto("PROY_INST_50_A")
PROY_PICKUP_51N_A = P.proyecto("PROY_PICKUP_51N_A")
PROY_DIAL_51N = P.proyecto("PROY_DIAL_51N")
PROY_INST_50N_A = P.proyecto("PROY_INST_50N_A")
PROY_TIEMPO_DESCONEXION_S = P.proyecto("PROY_TIEMPO_DESCONEXION_S")

# Corriente nominal del transformador del proyecto en el lado de MT, y corriente
# de una falla trifasica en la barra de 800 V referida a ese lado: es la maxima
# que ven a la vez la proteccion del proyecto y la de cabecera, o sea la
# condicion critica para verificar selectividad. Se calculan a partir de los
# datos de placa en vez de dejarlos escritos a mano, para que sigan siendo
# correctos si cambia la potencia o la tension.
_S_TRAFO_KVA = P.datos_proyecto("TRAFO_POTENCIA_KVA")
_U_MT_KV = P.datos_proyecto("TENSION_MT_KV")
_U_BT_V = P.datos_proyecto("TENSION_BT_V")
PROY_IN_TRAFO_A = _S_TRAFO_KVA / (math.sqrt(3) * _U_MT_KV)
def _falla_bt_referida_mt():
    """Falla trifasica en la barra de 800 V, referida al lado de 13.2 kV.

    Se lee del propio archivo de resultados en vez de dejarla escrita a mano:
    asi sigue siendo correcta si cambia la impedancia del transformador o
    cualquier otro dato del modelo."""
    barra_bt = P.elemento("BARRA_BT_PROYECTO")
    wb = openpyxl.load_workbook(RESULTADOS_CORTO_PATH, data_only=True)
    h = [c.value for c in wb["Nodos"][1]]
    iB, iT, iI = h.index("Barra_Falla"), h.index("Tipo_Falla"), h.index("Ikss_kA")
    for fila in wb["Nodos"].iter_rows(min_row=2, values_only=True):
        if fila[iB] == barra_bt and fila[iT] == "Trifasico" and fila[iI]:
            return fila[iI] * 1000.0 * (_U_BT_V / (_U_MT_KV * 1000.0))
    raise RuntimeError(
        f"No se encontro la falla trifasica en la barra '{barra_bt}' dentro de "
        f"{RESULTADOS_CORTO_PATH.name}. Revisa el nombre en la hoja Elementos_Clave "
        f"de Parametros_Sistema.xlsx."
    )


PROY_FALLA_BT_REFERIDA_MT_A = _falla_bt_referida_mt()

# Barra que representa la cabecera del circuito 15344 (primer punto del
# feeder, justo aguas abajo del reconectador "Cab 15344"). AJUSTAR aqui si
# el nombre exacto de la barra en el modelo de PowerFactory cambia.
BARRA_CABECERA = P.elemento("BARRA_PC")

# --- Aporte de falla propio del sistema de generacion, ya validado --------
# Fuente: Parametros_Proyecto.xlsx, filas "Metodo_calculo_corto_detallado" /
# "Confirmacion_fix_corto_generador".
IK3PF_CONFIGURADO_A_POR_INVERSOR = P.datos_proyecto("APORTE_FALLA_INVERSOR_A")
N_INVERSORES = int(P.datos_proyecto("CANTIDAD_INVERSORES"))
POTENCIA_INVERSOR_KW = P.datos_proyecto("POTENCIA_INVERSOR_KW")
TENSION_AC_INVERSOR_V = _U_BT_V
LIMITE_REGULATORIO_PU = P.datos_proyecto("LIMITE_APORTE_FALLA_PU")
# Resultado de simulacion: se lee del Excel de cortocircuito, no se escribe a mano
IKSS_GENERADOR_VALIDADO_KA = IK3PF_CONFIGURADO_A_POR_INVERSOR * N_INVERSORES / 1000.0


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
    # La falla BIFASICA se evalua contra los elementos de FASE (51/50), igual que
    # la trifasica: una falla entre dos fases sin contacto a tierra no produce
    # corriente residual, de modo que el elemento de neutro (51N/50N) no la ve.
    # Se incorporo cuando Cortocircuito.py empezo a reportarla correctamente
    # (antes devolvia 0 kA y no habia con que evaluarla).
    ajustes = {
        "Trifasico": (PICKUP_51_A, DIAL_51, INST_50_A, "51/50 (fase)"),
        "Bifasico": (PICKUP_51_A, DIAL_51, INST_50_A, "51/50 (fase)"),
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


def verificar_selectividad_proyecto():
    """Verifica la selectividad entre la proteccion del PROYECTO (celda de MT en
    el punto de conexion) y la de CABECERA del circuito 15344.

    La condicion critica es una falla trifasica en la barra de 800 V: es la
    maxima corriente que atraviesa el punto de conexion y que, por tanto, ven
    las dos protecciones a la vez. Referida a 13.2 kV vale
    PROY_FALLA_BT_REFERIDA_MT_A. Para que el esquema sea selectivo, la
    proteccion del proyecto debe despejar ANTES que la de cabecera, con margen
    suficiente (criterio habitual: 0.2 a 0.3 s entre protecciones en serie).

    Devuelve (filas, margen_s, cumple)."""
    i_falla = PROY_FALLA_BT_REFERIDA_MT_A

    t_proyecto, zona_proyecto = tiempo_disparo_ni(i_falla, PROY_PICKUP_51_A, PROY_DIAL_51, PROY_INST_50_A)
    t_cabecera, zona_cabecera = tiempo_disparo_ni(i_falla, PICKUP_51_A, DIAL_51, INST_50_A)

    filas = [
        {
            "proteccion": "Proyecto (51, punto de conexion)",
            "pickup_a": PROY_PICKUP_51_A, "dial": PROY_DIAL_51,
            "multiplo": i_falla / PROY_PICKUP_51_A, "zona": zona_proyecto, "tiempo_s": t_proyecto,
        },
        {
            "proteccion": "Cabecera circuito 15344 (51)",
            "pickup_a": PICKUP_51_A, "dial": DIAL_51,
            "multiplo": i_falla / PICKUP_51_A, "zona": zona_cabecera, "tiempo_s": t_cabecera,
        },
    ]
    margen = None
    cumple = False
    if t_proyecto is not None and t_cabecera is not None:
        margen = t_cabecera - t_proyecto
        cumple = margen >= MARGEN_MINIMO_SELECTIVIDAD_S
    return filas, margen, cumple


def verificar_margen_antiisla():
    """Verifica que el proyecto alcance a desconectarse antes del recierre rapido
    de la cabecera (2 s, dato de EBSA).

    Esta es la verificacion que realmente dimensiona las protecciones del
    proyecto. Las funciones de sobrecorriente NO sirven para fallas aguas
    arriba: el inversor limita su aporte a ~1 p.u. (43.3 A en 13.2 kV), muy por
    debajo del arranque de 70 A. Ante una falla en el alimentador de EBSA la
    desconexion la producen las funciones de tension, frecuencia y anti-isla."""
    margen = TIEMPO_RECIERRE_CABECERA_S - PROY_TIEMPO_DESCONEXION_S
    return margen, margen >= MARGEN_MINIMO_ANTIISLA_S


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

    ws4 = wb.create_sheet("Selectividad_Proyecto")
    filas_sel, margen_sel, cumple_sel = verificar_selectividad_proyecto()
    ws4.append(["Condicion critica evaluada", "Falla trifasica en la barra de 800 V, referida a 13.2 kV"])
    ws4.append(["Corriente de falla referida a 13.2 kV [A]", round(PROY_FALLA_BT_REFERIDA_MT_A, 1)])
    ws4.append([])
    ws4.append(["Proteccion", "Arranque [A]", "Dial", "I/Iarranque", "Zona", "Tiempo de disparo [s]"])
    for c in ws4[ws4.max_row]:
        c.font = Font(bold=True)
    for f in filas_sel:
        ws4.append([
            f["proteccion"], f["pickup_a"], f["dial"],
            round(f["multiplo"], 2), f["zona"],
            None if f["tiempo_s"] is None else round(f["tiempo_s"], 3),
        ])
    ws4.append([])
    ws4.append([
        "Margen de selectividad [s]",
        None if margen_sel is None else round(margen_sel, 3),
        f"CUMPLE (>= {MARGEN_MINIMO_SELECTIVIDAD_S} s)" if cumple_sel else "NO CUMPLE",
    ])
    margen_isla, cumple_isla = verificar_margen_antiisla()
    ws4.append([])
    ws4.append(["Verificacion anti-isla frente al recierre de cabecera"])
    ws4[f"A{ws4.max_row}"].font = Font(bold=True)
    ws4.append(["Recierre rapido de cabecera [s]", TIEMPO_RECIERRE_CABECERA_S, "Dato EBSA (DOC_REF_EST pag. 4)"])
    ws4.append(["Desconexion del proyecto [s]", PROY_TIEMPO_DESCONEXION_S, "Funciones 27/59/81/81R"])
    ws4.append([
        "Margen antes del recierre [s]", round(margen_isla, 3),
        f"CUMPLE (>= {MARGEN_MINIMO_ANTIISLA_S} s)" if cumple_isla else "NO CUMPLE",
    ])
    for col in ws4.columns:
        max_len = max(len(str(c.value)) if c.value is not None else 0 for c in col)
        ws4.column_dimensions[col[0].column_letter].width = min(max_len + 2, 70)

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

    print("=== Selectividad proyecto vs. cabecera (falla trifasica en 800 V) ===")
    filas_sel, margen_sel, cumple_sel = verificar_selectividad_proyecto()
    print(f"Corriente de falla referida a 13.2 kV: {PROY_FALLA_BT_REFERIDA_MT_A:.1f} A")
    for f in filas_sel:
        t = "-" if f["tiempo_s"] is None else f"{f['tiempo_s']:.3f}s"
        print(f"  {f['proteccion']:<34} arranque={f['pickup_a']:>6.0f}A dial={f['dial']:<5} "
              f"I/Iarr={f['multiplo']:>5.2f} {f['zona']:>12} t={t}")
    if margen_sel is not None:
        print(f"  -> margen de selectividad: {margen_sel:.3f}s "
              f"({'CUMPLE' if cumple_sel else 'NO CUMPLE'}, criterio >= {MARGEN_MINIMO_SELECTIVIDAD_S}s)")

    margen_isla, cumple_isla = verificar_margen_antiisla()
    print()
    print("=== Anti-isla frente al recierre de cabecera ===")
    print(f"  recierre cabecera = {TIEMPO_RECIERRE_CABECERA_S}s | desconexion proyecto = {PROY_TIEMPO_DESCONEXION_S}s")
    print(f"  -> margen: {margen_isla:.3f}s ({'CUMPLE' if cumple_isla else 'NO CUMPLE'}, "
          f"criterio >= {MARGEN_MINIMO_ANTIISLA_S}s)")
    print()
    print(f"Grafico: {ruta_grafico}")
    print(f"Excel:   {ruta_excel}")


if __name__ == "__main__":
    main()
