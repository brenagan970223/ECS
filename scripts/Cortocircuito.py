"""
Analisis de cortocircuito en TODAS las barras del sistema, para cada CASO
DE RED (CasoBase + Network Variations listadas en Network_Variations.xlsx)
y para 3 TIPOS DE FALLA: trifasico, bifasico y monofasico (a tierra).

Recicla la misma logica y estructura de Flujo_carga.py: mismo manejo de
Network Variations (activar/desactivar, tolerante a las que no existan),
misma relectura de elementos DESPUES de activar cada caso (una Network
Variation puede agregar equipos nuevos), y resultados organizados por tipo
de elemento en formato largo Nombre | Caso | Barra_Falla | Tipo_Falla |
<variables>.

A proposito NO reutiliza Parametros_Demanda.xlsx ni el barrido por hora:
con el metodo de calculo usado aqui (IEC 60909 / VDE 0102 Part 0,
ildfinit=False, ver configurar_shc), la corriente de cortocircuito se
calcula con una fuente de tension equivalente y NO depende de plini/qlini
de las cargas. Agregar la dimension Hora repetiria el mismo Ikss 12 veces
por barra/caso/tipo de falla, sin aportar informacion nueva - decision
confirmada con el usuario 2026-09-13.

Motivacion (CREG 174 de 2021, numeral 8.4 "Calculo contribucion a la
corriente de cortocircuito"): verificar que la conexion del generador no
incremente la corriente de cortocircuito (Icc) en las subestaciones de la
zona de influencia por encima de la capacidad de corte de los
interruptores existentes, calculando la intensidad de fase maxima ante
falla polifasica (trifasica/bifasica) o falla a tierra (monofasica) en
cada barra.

Modo asistido (igual que Flujo_carga.py): se corre desde dentro de
PowerFactory. Se ejecuta manualmente UNA vez; a partir de ahi el barrido
sobre todas las barras, tipos de falla y casos de red corre de manera
desatendida (sin intervencion manual por combinacion).

Pasos:
  1. Conexion a PowerFactory (proyecto y caso de estudio ya activos). Se
     borran los SVG de resultados/graficos_red_corto/ de corridas
     anteriores.
  2. Lectura y emparejamiento de Network Variations (caso base siempre se
     corre; una Network Variation del Excel que no exista en PowerFactory
     se excluye, avisa, y no detiene el barrido).
  3. Por cada CASO DE RED: se activa el caso, y RECIEN DESPUES se listan
     los elementos calc-relevantes (nodos, lineas, transformadores 2 y 3
     devanados, generadores) - igual que en Flujo_carga.py, porque esa
     lista cambia segun el caso activo.
  4. Por cada BARRA del caso y por cada TIPO DE FALLA (3psc/2psc/spgf):
     ejecuta ComShc con esa barra como punto de falla, exporta el
     diagrama a SVG, y si el calculo es valido registra por cada elemento
     sus variables tecnicas.
  5. Al terminar, desactiva todas las Network Variations (vuelve al caso
     base) y exporta a Resultados_Cortocircuito.xlsx.

IMPORTANTE - transparencia sobre que esta verificado y que no:
  - CONFIRMADO (documentacion/ejemplos reales de scripting DIgSILENT):
    comando 'ComShc'; atributos shcobj (barra de falla), iopt_allbus=0
    (falla en una barra especifica), iopt_shc con codigos '3psc'
    (trifasico), '2psc' (bifasico), 'spgf' (monofasico a tierra);
    resultado m:Ikss (corriente de cortocircuito simetrica inicial, kA)
    y m:Skss (potencia de cortocircuito, MVA) leidos del ElmTerm fallado.
  - METODO DE CALCULO: se fuerza iopt_mde = 0, que corresponde a
    "VDE 0102 Part 0 / DIN EN 60909-0" (la adopcion europea/aleman de
    IEC 60909 - en la practica el mismo metodo de calculo). Es el metodo
    por defecto de esta instalacion de PowerFactory, confirmado con el
    usuario. El script imprime el valor de iopt_mde al arrancar para que
    quede visible en la Output Window en cada corrida.
  - MEJOR ESFUERZO, NO VERIFICADO todavia en un proyecto real: los
    nombres de atributo de corriente de aporte en lineas/transformadores/
    generadores (m:Ikss:bus1, m:Ikss:bushv, etc.) se extrapolaron del
    patron ':busX' que SI esta validado para flujo de carga en este mismo
    proyecto (m:P:bus1, m:I:bus1). Si salen vacios o con nombre
    incorrecto, avisa para ajustarlos - no detienen el barrido
    (usan safe_get, igual que el resto del proyecto).
  - Se exporta un SVG por CADA combinacion caso x barra fallada x tipo de
    falla (uno por cada ComShc.Execute(), converja o no) - confirmado con
    el usuario, a pesar de que son muchos archivos para un sistema con
    varias barras. Se exporta DESPUES de calcular esa falla puntual (no
    antes), para que el diagrama refleje el resultado de ESE calculo y
    no un estado en blanco o de un calculo anterior.

Variables registradas por tipo de elemento:
  Nodos (ElmTerm):        Ikss_kA (m:Ikss), Skss_MVA (m:Skss) - SOLO de la
                          barra realmente fallada en esa combinacion (una
                          fila por Caso x Barra_Falla x Tipo_Falla, igual
                          cantidad de filas que la hoja Resumen). No se
                          registran las demas barras: el Ikss/Skss de una
                          barra que NO es la fallada, durante el calculo
                          de otra barra, no es su propio nivel de
                          cortocircuito (para eso hay que fallarla a
                          ella, y eso ya se hace en su propia iteracion) -
                          incluirlas solo inflaba la hoja sin aportar
                          informacion.
  Lineas (ElmLne):        Nodo_I, Nodo_J, Ikss_bus1_kA, Ikss_bus2_kA
  Transformadores 2 devanados (ElmTr2): Ikss_HV_kA, Ikss_LV_kA
  Transformadores 3 devanados (ElmTr3): Ikss_HV_kA, Ikss_MV_kA, Ikss_LV_kA
  Generadores (ElmGenstat / ElmSym / ElmPvsys): Barra (bus1, nombre de
                          la barra a la que esta conectado), Ikss_kA
"""

import datetime as dt
import re
import sys
from pathlib import Path

try:
    import powerfactory as pf
except ImportError:
    sys.path.append(r"C:\Program Files\DIgSILENT\PowerFactory 2024\Python\3.12")
    import powerfactory as pf

import openpyxl
from openpyxl.styles import Font

BASE_DIR = Path(r"C:\Users\USUARIO\Desktop\Puerto Boyaca\ECS")
VARIATIONS_PATH = BASE_DIR / "inputs" / "Network_Variations.xlsx"
VARIATIONS_SHEET = "Network_Variations"
RESULTADOS_PATH = BASE_DIR / "resultados" / "Resultados_Cortocircuito.xlsx"
GRAFICOS_DIR = BASE_DIR / "resultados" / "graficos_red_corto"
CASO_BASE = "CasoBase"

# tipo mostrado en tablas/nombres de archivo -> codigo iopt_shc de PowerFactory
TIPOS_FALLA = [
    ("Trifasico", "3psc"),
    ("Bifasico", "2psc"),
    ("Monofasico", "spgf"),
]


def sanitizar_nombre_archivo(texto):
    return re.sub(r'[\\/:*?"<>|]', "_", str(texto))


def exportar_diagrama_red(app, nombre_proyecto, nombre_caso, barra_falla, tipo_falla):
    """Exporta el/los diagramas unifilares (SetDeskpage) activos a SVG, una
    vez por cada combinacion caso x barra fallada x tipo de falla (una
    imagen por cada calculo de ComShc). Ver docstring del modulo."""
    try:
        desktop = app.GetGraphicsBoard()
        paginas = list(desktop.GetContents("*.SetDeskpage", 1)) if desktop is not None else []
        if not paginas:
            app.PrintError("No se encontro ningun diagrama de red (SetDeskpage) abierto para exportar a SVG.")
            return

        marca_tiempo = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        base = sanitizar_nombre_archivo(f"{nombre_proyecto}_{marca_tiempo}_{nombre_caso}_{barra_falla}_{tipo_falla}")

        for pagina in paginas:
            desktop.Show(pagina)
            sufijo = f"_{sanitizar_nombre_archivo(pagina.loc_name)}" if len(paginas) > 1 else ""
            ruta = GRAFICOS_DIR / f"{base}{sufijo}.svg"

            com_wr = app.GetFromStudyCase("ComWr")
            com_wr.SetAttribute("iopt_rd", "svg")
            com_wr.SetAttribute("iopt_savas", 0)
            com_wr.SetAttribute("f", str(ruta))
            com_wr.Execute()
    except Exception as exc:
        app.PrintError(
            f"No se pudo exportar el diagrama de red a SVG (caso {nombre_caso}, barra {barra_falla}, "
            f"falla {tipo_falla}): {exc}"
        )


def limpiar_carpeta_graficos(app):
    if not GRAFICOS_DIR.exists():
        return
    borrados = 0
    for archivo in GRAFICOS_DIR.glob("*.svg"):
        try:
            archivo.unlink()
            borrados += 1
        except OSError as exc:
            app.PrintError(f"No se pudo borrar el SVG anterior '{archivo.name}': {exc}")
    app.PrintInfo(f"Limpieza de graficos_red_corto: {borrados} SVG de corridas anteriores borrados.")


def emparejar_por_nombre(objetos, nombres, descripcion):
    """Empareja nombres del Excel con objetos de PowerFactory por coincidencia
    de substring en loc_name. Devuelve (mapeo, faltantes); un nombre ambiguo
    (coincide con mas de un objeto) si detiene la ejecucion."""
    mapeo = {}
    faltantes = []
    for nombre in nombres:
        candidatos = [obj for obj in objetos if nombre in obj.loc_name]
        if len(candidatos) == 1:
            mapeo[nombre] = candidatos[0]
        elif len(candidatos) == 0:
            faltantes.append(nombre)
        else:
            raise RuntimeError(
                f"Nombre de {descripcion} ambiguo '{nombre}': coincide con {[c.loc_name for c in candidatos]}"
            )
    return mapeo, faltantes


def leer_nombres_columna(path, sheet_name):
    if not path.exists():
        return []
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[sheet_name]
    nombres = []
    for (valor,) in ws.iter_rows(min_row=2, max_col=1, values_only=True):
        if valor is not None and str(valor).strip():
            nombres.append(str(valor).strip())
    return nombres


def listar_network_variations(app):
    carpeta = app.GetProjectFolder("scheme")
    if carpeta is None:
        return []
    return list(carpeta.GetContents("*.IntScheme", 1))


def desactivar_todas_variations(variations_pf, activas):
    for variacion in variations_pf:
        if variacion in activas:
            variacion.Deactivate()


def activar_caso_red(app, nombre_caso, variacion_objetivo, variations_pf):
    activas = list(app.GetActiveNetworkVariations())
    desactivar_todas_variations(variations_pf, activas)
    if variacion_objetivo is not None:
        if variacion_objetivo.Activate():
            raise RuntimeError(f"No se pudo activar la Network Variation '{nombre_caso}'.")


def safe_get(obj, attr):
    try:
        return obj.GetAttribute(attr)
    except Exception:
        return None


# Atributos candidatos para leer la corriente de cortocircuito, en orden de
# preferencia. Se prueban uno por uno y se usa el primero que devuelva un valor
# distinto de None y de cero.
#
# POR QUE UNA LISTA Y NO UN SOLO ATRIBUTO: 'm:Ikss' devuelve el valor correcto
# para falla trifasica y monofasica, pero en la falla BIFASICA devolvia 0.0000
# en las 31 combinaciones barra x caso de la corrida del 2026-09-19. Una
# corriente de falla bifasica de cero es fisicamente imposible en una red
# energizada (no depende de la secuencia cero: Ik2 = raiz(3)*c*Un/|Z1+Z2|, que
# para Z2 ~ Z1 da Ik2 ~ 0.866*Ik3), asi que no es un problema del modelo sino
# de que en una falla desbalanceada PowerFactory reporta el resultado en las
# variables POR FASE y deja 'm:Ikss' sin poblar. En una falla bifasica las
# fases falladas son dos (por defecto L2-L3), por eso se toma el MAXIMO de las
# corrientes de fase: sirve igual para trifasica (las tres iguales), monofasica
# (solo la fallada) y bifasica (las dos falladas).
ATRIBUTOS_IKSS_BARRA = ["m:Ikss", "m:Ikss:A", "m:Ikss:B", "m:Ikss:C"]
ATRIBUTOS_SKSS_BARRA = ["m:Skss", "m:Skss:A", "m:Skss:B", "m:Skss:C"]


def leer_maximo_por_fase(obj, atributos):
    """Devuelve (valor, atributo_usado) tomando el mayor valor no nulo entre los
    atributos candidatos. Si el primero (el agregado, ej. 'm:Ikss') ya trae un
    valor util, se usa ese; si viene nulo o en cero, se cae a las variables por
    fase. Devuelve (None, None) si ninguno da un valor."""
    agregado = safe_get(obj, atributos[0])
    if agregado not in (None, 0):
        return agregado, atributos[0]

    mejor, mejor_attr = None, None
    for attr in atributos[1:]:
        valor = safe_get(obj, attr)
        if valor not in (None, 0) and (mejor is None or valor > mejor):
            mejor, mejor_attr = valor, attr
    if mejor is not None:
        return mejor, mejor_attr
    # ningun candidato dio valor util: se devuelve el agregado tal cual vino
    # (puede ser 0 o None) para que quede registrado y se marque como invalido
    return agregado, atributos[0]


def diagnosticar_atributos(app, obj, atributos, etiqueta):
    """Imprime en la Output Window que atributos candidatos traen valor y cuales
    no, para una barra concreta. Se llama una sola vez por tipo de falla: con
    eso queda documentado en el log de la corrida cual es la variable que
    realmente usa esta version de PowerFactory para cada tipo de falla, sin
    tener que adivinarlo."""
    partes = []
    for attr in atributos:
        valor = safe_get(obj, attr)
        partes.append(f"{attr}={'None' if valor is None else f'{valor:.6f}'}")
    app.PrintInfo(f"  [diagnostico {etiqueta}] {obj.loc_name}: " + " | ".join(partes))


# ---------------------------------------------------------------------------
# Diagnostico de los datos de secuencia que realmente tiene el modelo
# ---------------------------------------------------------------------------
# Datos de referencia del OR para el equivalente de red de 115 kV
# (DOC_REF_EST_1787069892338.pdf pag. 4, "CORRIENTE DE CORTOCIRCUITO EN
# BARRAJE 115kV DE LA SE PUERTO BOYACA").
IK3_REF_115KV_KA = 6.002
IK1_REF_115KV_KA = 2.622
XR_REF_115KV = 10.0

# Atributos que se inspeccionan por tipo de objeto. Se prueban todos y se
# reporta cuales existen en esta version de PowerFactory y cuales no: asi el
# diagnostico sirve aunque algun nombre de atributo cambie entre versiones.
ATRIBUTOS_XNET = [
    ("ikss", "Ikss max [kA] - cortocircuito TRIFASICO", "directa"),
    ("snss", "Skss max [MVA] - potencia de cortocircuito", "directa"),
    ("rntxn", "R/X max - relacion resistencia/reactancia", "directa"),
    ("z2tz1", "Z2/Z1 max - secuencia inversa", "inversa"),
    ("x0tx1", "X0/X1 max", "CERO"),
    ("r0tx0", "R0/X0 max", "CERO"),
]
ATRIBUTOS_TYPLNE = [
    ("rline", "R' [Ohm/km]", "directa"),
    ("xline", "X' [Ohm/km]", "directa"),
    ("rline0", "R0' [Ohm/km]", "CERO"),
    ("xline0", "X0' [Ohm/km]", "CERO"),
]
ATRIBUTOS_TYPTR2 = [
    ("strn", "Potencia nominal [MVA]", "datos"),
    ("tr2cn_h", "Grupo de conexion lado HV", "datos"),
    ("tr2cn_l", "Grupo de conexion lado LV", "datos"),
    ("uktr", "uk [%] - tension de cortocircuito", "directa"),
    ("uk0tr", "uk0 [%] - tension de cortocircuito", "CERO"),
    ("ur0tr", "uR0 [%] - componente resistiva", "CERO"),
]
ATRIBUTOS_TYPTR3 = [
    ("strn3_h", "Potencia nominal HV [MVA]", "datos"),
    ("tr3cn_h", "Grupo de conexion HV", "datos"),
    ("tr3cn_m", "Grupo de conexion MV", "datos"),
    ("tr3cn_l", "Grupo de conexion LV", "datos"),
    ("uktr3_h", "uk HV-MV [%]", "directa"),
    ("uktr3_m", "uk MV-LV [%]", "directa"),
    ("uktr3_l", "uk LV-HV [%]", "directa"),
    ("uk0tr3_h", "uk0 HV-MV [%]", "CERO"),
    ("uk0tr3_m", "uk0 MV-LV [%]", "CERO"),
    ("uk0tr3_l", "uk0 LV-HV [%]", "CERO"),
]


def relacion_z0z1_requerida(ik3_ka, ik1_ka):
    """Z0/Z1 que hace falta para reproducir una falla monofasica dada.

    De las formulas IEC 60909 para falla trifasica y monofasica a tierra,
    asumiendo Z2 = Z1 (valido en redes de distribucion):

        Ik3 = c*Un / (raiz(3)*Z1)
        Ik1 = raiz(3)*c*Un / |2*Z1 + Z0|

    dividiendo una entre otra y despejando:

        Z0/Z1 = 3/(Ik1/Ik3) - 2
    """
    if not ik3_ka or not ik1_ka:
        return None
    return 3.0 / (ik1_ka / ik3_ka) - 2.0


def _volcar_atributos(app, obj, atributos, sangria="      "):
    """Imprime los atributos que existen, agrupados por secuencia, y devuelve
    un dict {nombre: valor} con los que si trajeron valor."""
    encontrados = {}
    for attr, descripcion, grupo in atributos:
        valor = safe_get(obj, attr)
        if valor is None:
            app.PrintInfo(f"{sangria}[{grupo:>7}] {attr:<10} = (no existe en esta version)   {descripcion}")
        else:
            encontrados[attr] = valor
            texto = f"{valor:.6g}" if isinstance(valor, (int, float)) else str(valor)
            app.PrintInfo(f"{sangria}[{grupo:>7}] {attr:<10} = {texto:<22} {descripcion}")
    return encontrados


def diagnosticar_datos_secuencia(app):
    """Recorre el modelo ANTES de calcular y reporta que datos de secuencia
    tiene puesto cada elemento, en que objeto estan y que hay que cambiar.

    Existe porque la corrida anterior mostro Ik1 exactamente igual a Ik3 en
    las 15 barras de media tension, senal de que la secuencia cero no esta
    parametrizada; este diagnostico dice DONDE esta el dato que falta, con el
    nombre y la ruta exacta del objeto, para no tener que buscarlo a mano."""
    app.PrintInfo("")
    app.PrintInfo("=" * 78)
    app.PrintInfo("DIAGNOSTICO: datos de secuencia que esta usando el modelo")
    app.PrintInfo("=" * 78)

    acciones = []

    # ---- 1. Equivalente de red (ElmXnet) ----
    app.PrintInfo("")
    app.PrintInfo("[1] EQUIVALENTE DE RED (clase ElmXnet)")
    app.PrintInfo("-" * 78)
    xnets = list(app.GetCalcRelevantObjects("*.ElmXnet", 1))
    if not xnets:
        app.PrintError("    No se encontro ningun ElmXnet calc-relevante.")
    for xnet in xnets:
        barra = nombre_nodo(safe_get(xnet, "bus1")) or "?"
        app.PrintInfo(f"    Objeto : {xnet.loc_name}")
        app.PrintInfo(f"    Ruta   : {xnet.GetFullName()}")
        app.PrintInfo(f"    Barra  : {barra}")
        valores = _volcar_atributos(app, xnet, ATRIBUTOS_XNET)

        z0z1_req = relacion_z0z1_requerida(IK3_REF_115KV_KA, IK1_REF_115KV_KA)
        x0tx1_actual = valores.get("x0tx1")
        app.PrintInfo("")
        app.PrintInfo(f"      Referencia del OR: Ik3 = {IK3_REF_115KV_KA} kA, Ik1 = {IK1_REF_115KV_KA} kA, X/R = {XR_REF_115KV}")
        app.PrintInfo(f"      -> X0/X1 necesario para reproducir esa Ik1: {z0z1_req:.3f}")
        if x0tx1_actual is None:
            app.PrintError("      -> No se pudo leer 'x0tx1'. Buscar en la pestana de cortocircuito "
                           "el campo X0/X1 (o el de corriente monofasica) y ajustarlo a mano.")
        elif abs(x0tx1_actual - 1.0) < 1e-6:
            acciones.append(
                f"ElmXnet '{xnet.loc_name}': poner x0tx1 = {z0z1_req:.3f} (hoy vale {x0tx1_actual:.3f}, "
                f"el valor por defecto, que hace Ik1 = Ik3) y r0tx0 = {1.0 / XR_REF_115KV:.3f}"
            )
            app.PrintError(f"      -> ACCION: x0tx1 vale {x0tx1_actual:.3f} (por defecto). "
                           f"Cambiarlo a {z0z1_req:.3f}. Con 1.0, la falla monofasica sale igual a la trifasica.")
        elif abs(x0tx1_actual - z0z1_req) > 0.05:
            acciones.append(
                f"ElmXnet '{xnet.loc_name}': x0tx1 vale {x0tx1_actual:.3f} pero la referencia del OR pide {z0z1_req:.3f}"
            )
            app.PrintError(f"      -> ACCION: x0tx1 vale {x0tx1_actual:.3f}, se esperaba {z0z1_req:.3f}.")
        else:
            app.PrintInfo(f"      -> OK: x0tx1 = {x0tx1_actual:.3f} coincide con la referencia del OR.")

    # ---- 2. Transformadores ----
    for clase, atributos, etiqueta in (
        ("*.ElmTr2", ATRIBUTOS_TYPTR2, "TRANSFORMADORES DE 2 DEVANADOS"),
        ("*.ElmTr3", ATRIBUTOS_TYPTR3, "TRANSFORMADORES DE 3 DEVANADOS"),
    ):
        app.PrintInfo("")
        app.PrintInfo(f"[2] {etiqueta} (los datos viven en el TIPO, no en el elemento)")
        app.PrintInfo("-" * 78)
        for tr in app.GetCalcRelevantObjects(clase, 1):
            tipo = safe_get(tr, "typ_id")
            app.PrintInfo(f"    Elemento: {tr.loc_name}")
            if tipo is None:
                app.PrintError("      Sin tipo asignado (typ_id vacio): no se pueden leer sus impedancias.")
                continue
            app.PrintInfo(f"    Tipo    : {tipo.loc_name}")
            app.PrintInfo(f"    Ruta    : {tipo.GetFullName()}")
            valores = _volcar_atributos(app, tipo, atributos)
            # aviso si la secuencia cero esta igual a la directa o vacia
            for attr_dir, attr_cero in (("uktr", "uk0tr"), ("uktr3_h", "uk0tr3_h")):
                if attr_dir in valores and attr_cero in valores:
                    if not valores[attr_cero]:
                        acciones.append(f"Tipo '{tipo.loc_name}': {attr_cero} esta en cero o vacio")
                        app.PrintError(f"      -> ACCION: {attr_cero} = 0. Revisar la secuencia cero de este tipo.")
                    elif abs(valores[attr_cero] - valores[attr_dir]) < 1e-9:
                        app.PrintInfo(f"      -> Nota: {attr_cero} = {attr_dir}. Correcto SOLO si el OR lo declara asi "
                                      f"(para T9, EBSA reporta secuencia 0 igual a la directa).")

    # ---- 3. Lineas ----
    app.PrintInfo("")
    app.PrintInfo("[3] LINEAS (los datos viven en el TIPO de conductor)")
    app.PrintInfo("-" * 78)
    tipos_vistos = {}
    for ln in app.GetCalcRelevantObjects("*.ElmLne", 1):
        tipo = safe_get(ln, "typ_id")
        if tipo is None:
            app.PrintError(f"    Linea '{ln.loc_name}': sin tipo asignado.")
            continue
        tipos_vistos.setdefault(tipo.GetFullName(), (tipo, []))[1].append(ln.loc_name)
    for _, (tipo, lineas) in tipos_vistos.items():
        app.PrintInfo(f"    Tipo  : {tipo.loc_name}   (usado por {len(lineas)} lineas: {', '.join(lineas[:4])}{'...' if len(lineas) > 4 else ''})")
        app.PrintInfo(f"    Ruta  : {tipo.GetFullName()}")
        valores = _volcar_atributos(app, tipo, ATRIBUTOS_TYPLNE)
        r1, x1 = valores.get("rline"), valores.get("xline")
        r0, x0 = valores.get("rline0"), valores.get("xline0")
        if r1 and x1:
            if not r0 or not x0:
                acciones.append(f"Tipo de linea '{tipo.loc_name}': secuencia cero vacia o en cero (rline0/xline0)")
                app.PrintError("      -> ACCION: la secuencia cero esta vacia. Valores tipicos de linea aerea de "
                               "distribucion: R0' = 2 a 3 x R', X0' = 3 x X'. EBSA no los entrega, asi que "
                               "habria que asumirlos y declararlo en el informe.")
            else:
                app.PrintInfo(f"      -> Relaciones actuales: R0'/R' = {r0 / r1:.2f}, X0'/X' = {x0 / x1:.2f} "
                              f"(tipico en linea aerea: 2-3 y ~3)")
                if abs(r0 / r1 - 1.0) < 1e-6 and abs(x0 / x1 - 1.0) < 1e-6:
                    acciones.append(f"Tipo de linea '{tipo.loc_name}': secuencia cero igual a la directa (valores por defecto)")
                    app.PrintError("      -> ACCION: secuencia cero identica a la directa, son los valores por "
                                   "defecto. En una linea aerea real X0' es del orden de 3 veces X'.")

    # ---- resumen ----
    app.PrintInfo("")
    app.PrintInfo("=" * 78)
    if acciones:
        app.PrintInfo(f"RESUMEN: {len(acciones)} punto(s) a corregir antes de confiar en las fallas monofasicas")
        for i, accion in enumerate(acciones, 1):
            app.PrintError(f"  {i}. {accion}")
    else:
        app.PrintInfo("RESUMEN: no se detectaron datos de secuencia cero sin parametrizar.")
    app.PrintInfo("=" * 78)
    app.PrintInfo("")


def nombre_nodo(cubicle):
    try:
        return cubicle.cterm.loc_name
    except Exception:
        return None


def configurar_shc(shc, barra):
    """Configuracion base de ComShc: falla en una barra especifica (no todas
    las barras a la vez), impedancia de falla nula (falla franca, para
    obtener la Icc maxima que pide CREG 174), asunciones de calculo "todo
    incluido" (cargas, capacitancia de lineas, magnetizacion de
    transformadores, shunts), y metodo de calculo IEC 60909 / VDE 0102
    Part 0 (iopt_mde = 0) - ver nota en el docstring del modulo."""
    shc.iopt_allbus = 0
    shc.shcobj = barra
    shc.iopt_mde = 0  # IEC 60909 / VDE 0102 Part 0 (DIN EN 60909-0)
    shc.Rf = 0
    shc.Xf = 0
    shc.ildfinit = False
    shc.cfac_full = 1
    shc.ilngLoad = True
    shc.ilngLneCap = True
    shc.ilngTrfMag = True
    shc.ilngShnt = True


def main():
    app = pf.GetApplication()
    if app is None:
        raise RuntimeError(
            "CONEXION FALLIDA: GetApplication() devolvio None. "
            "Si esto corre fuera de PowerFactory, cierra cualquier instancia grafica abierta."
        )

    project = app.GetActiveProject()
    if project is None:
        raise RuntimeError("No hay ningun proyecto activo. Abre el proyecto en PowerFactory antes de correr el script.")

    study_case = app.GetActiveStudyCase()
    if study_case is None:
        raise RuntimeError("No hay ningun caso de estudio activo. Activa un Study Case antes de correr el script.")

    nombre_proyecto = project.loc_name
    GRAFICOS_DIR.mkdir(parents=True, exist_ok=True)
    limpiar_carpeta_graficos(app)

    app.PrintInfo("===================================")
    app.PrintInfo(f"Proyecto activo: {project}")
    app.PrintInfo(f"Caso de estudio activo: {study_case}")
    app.PrintInfo(
        "Metodo de calculo de cortocircuito: iopt_mde = 0 (IEC 60909 / VDE 0102 Part 0), "
        "forzado por el script en cada ejecucion de ComShc."
    )

    # --- lectura y emparejamiento de Network Variations (caso base siempre se corre) ---
    variations_pf = listar_network_variations(app)
    nombres_variations = leer_nombres_columna(VARIATIONS_PATH, VARIATIONS_SHEET)
    mapeo_variations, variations_faltantes = emparejar_por_nombre(variations_pf, nombres_variations, "Network Variation")
    if variations_faltantes:
        app.PrintError(
            f"Network Variations del Excel que NO existen en PowerFactory (se excluyen, "
            f"no detienen la ejecucion): {variations_faltantes}"
        )

    casos = [(CASO_BASE, None)] + [(obj.loc_name, obj) for obj in mapeo_variations.values()]
    app.PrintInfo(f"Casos de red a simular: {[nombre for nombre, _ in casos]}")

    app.PrintInfo("Asegurando estado inicial: desactivando Network Variations antes de iniciar el barrido...")
    activar_caso_red(app, CASO_BASE, None, variations_pf)

    # Se diagnostica con el CasoBase ya activo (recien desactivadas las Network
    # Variations), que es el estado en que la red base esta completa.
    diagnosticar_datos_secuencia(app)

    shc = app.GetFromStudyCase("ComShc")

    filas_resumen = []
    filas_nodos = []
    filas_lineas = []
    filas_trafos_2w = []
    filas_trafos_3w = []
    filas_generadores = []
    tipos_diagnosticados = set()  # para volcar el diagnostico de atributos una sola vez por tipo de falla

    for nombre_caso, variacion_obj in casos:
        activar_caso_red(app, nombre_caso, variacion_obj, variations_pf)
        app.PrintInfo(f"--- Caso de red: {nombre_caso} ---")

        # se relee la red DESPUES de activar el caso (ver Flujo_carga.py: una
        # Network Variation puede agregar/quitar equipos)
        terminals = list(app.GetCalcRelevantObjects("*.ElmTerm", 1))
        lines = list(app.GetCalcRelevantObjects("*.ElmLne", 1))
        trafos_2w = list(app.GetCalcRelevantObjects("*.ElmTr2", 1))
        trafos_3w = list(app.GetCalcRelevantObjects("*.ElmTr3", 1))
        generadores = (
            list(app.GetCalcRelevantObjects("*.ElmGenstat", 1))
            + list(app.GetCalcRelevantObjects("*.ElmSym", 1))
            + list(app.GetCalcRelevantObjects("*.ElmPvsys", 1))
        )
        app.PrintInfo(
            f"[{nombre_caso}] elementos: {len(terminals)} nodos, {len(lines)} lineas, "
            f"{len(trafos_2w)} transformadores 2 dev., {len(trafos_3w)} transformadores 3 dev., "
            f"{len(generadores)} generadores."
        )

        for barra in terminals:
            for nombre_tipo, codigo_tipo in TIPOS_FALLA:
                configurar_shc(shc, barra)
                shc.iopt_shc = codigo_tipo

                err = shc.Execute()
                # ComShc con IEC 60909/VDE 0102 es un calculo directo (no iterativo),
                # a diferencia de ComLdf: no "converge" o no, simplemente calcula bien
                # o da error (configuracion invalida, dato faltante, etc.)
                ejecuto_ok = err == 0

                # se exporta una imagen por CADA calculo (valido o no), como en
                # Flujo_carga.py, para poder revisar cada falla individualmente
                exportar_diagrama_red(app, nombre_proyecto, nombre_caso, barra.loc_name, nombre_tipo)

                if not ejecuto_ok:
                    filas_resumen.append([nombre_caso, barra.loc_name, nombre_tipo, False, "ComShc devolvio error"])
                    app.PrintError(
                        f"Caso {nombre_caso}, barra {barra.loc_name}, falla {nombre_tipo}: "
                        f"CALCULO NO VALIDO (codigo {err})."
                    )
                    continue

                # la primera vez que se calcula cada tipo de falla se vuelca al log
                # que atributos traen valor, para dejar documentado cual usa esta
                # version de PowerFactory (ver ATRIBUTOS_IKSS_BARRA)
                if nombre_tipo not in tipos_diagnosticados:
                    tipos_diagnosticados.add(nombre_tipo)
                    app.PrintInfo(f"  Atributos disponibles para falla {nombre_tipo}:")
                    diagnosticar_atributos(app, barra, ATRIBUTOS_IKSS_BARRA, nombre_tipo)

                ikss, attr_ikss = leer_maximo_por_fase(barra, ATRIBUTOS_IKSS_BARRA)
                skss, _ = leer_maximo_por_fase(barra, ATRIBUTOS_SKSS_BARRA)

                # VALIDACION FISICA: una corriente de falla nula en una barra
                # energizada no existe. Antes el script marcaba Valido=True con
                # solo mirar el codigo de retorno de ComShc, y por eso las 93
                # combinaciones de la corrida anterior salieron "validas" pese a
                # que las 31 bifasicas venian en cero.
                if ikss in (None, 0):
                    filas_resumen.append([
                        nombre_caso, barra.loc_name, nombre_tipo, False,
                        f"Ikss nula o vacia: ningun atributo de {ATRIBUTOS_IKSS_BARRA} trajo valor",
                    ])
                    app.PrintError(
                        f"Caso {nombre_caso}, barra {barra.loc_name}, falla {nombre_tipo}: "
                        f"ComShc ejecuto sin error pero Ikss vino nula/cero - RESULTADO NO UTILIZABLE. "
                        f"Revisar que variable reporta la corriente para este tipo de falla."
                    )
                else:
                    filas_resumen.append([nombre_caso, barra.loc_name, nombre_tipo, True, attr_ikss])

                # solo se registra la barra realmente fallada (Nombre == Barra_Falla):
                # m:Ikss/m:Skss en las demas barras durante ESTE calculo puntual no es
                # el nivel de cortocircuito propio de esas otras barras (para eso hay
                # que fallarlas a ellas, y eso ya se hace en su propia iteracion) -
                # registrarlas aqui solo inflaba la hoja sin aportar informacion.
                filas_nodos.append([
                    barra.loc_name, nombre_caso, barra.loc_name, nombre_tipo,
                    ikss, skss, attr_ikss,
                ])

                for ln in lines:
                    filas_lineas.append([
                        ln.loc_name, nombre_nodo(ln.bus1), nombre_nodo(ln.bus2), nombre_caso, barra.loc_name, nombre_tipo,
                        safe_get(ln, "m:Ikss:bus1"), safe_get(ln, "m:Ikss:bus2"),
                    ])

                for tr in trafos_2w:
                    filas_trafos_2w.append([
                        tr.loc_name, nombre_caso, barra.loc_name, nombre_tipo,
                        safe_get(tr, "m:Ikss:bushv"), safe_get(tr, "m:Ikss:buslv"),
                    ])

                for tr in trafos_3w:
                    filas_trafos_3w.append([
                        tr.loc_name, nombre_caso, barra.loc_name, nombre_tipo,
                        safe_get(tr, "m:Ikss:bushv"), safe_get(tr, "m:Ikss:busmv"), safe_get(tr, "m:Ikss:buslv"),
                    ])

                for gen in generadores:
                    filas_generadores.append([
                        gen.loc_name, nombre_nodo(gen.bus1), nombre_caso, barra.loc_name, nombre_tipo,
                        safe_get(gen, "m:Ikss:bus1"),
                    ])

    # deja PowerFactory en el caso base al terminar
    activar_caso_red(app, CASO_BASE, None, variations_pf)

    verificar_coherencia_fisica(app, filas_nodos)

    exportar_resultados(
        app, filas_resumen, filas_nodos, filas_lineas, filas_trafos_2w, filas_trafos_3w, filas_generadores,
        variations_faltantes,
    )
    app.PrintInfo("===================================")


def verificar_coherencia_fisica(app, filas_nodos):
    """Chequeos de coherencia sobre los resultados, ANTES de exportarlos.

    Son dos relaciones que la fisica impone y que sirven para detectar de una
    vez si algun tipo de falla quedo mal leido (que fue justo lo que paso con
    la bifasica en la corrida del 2026-09-19):

      1. Ik2 / Ik3 ~ 0.866  (raiz(3)/2), porque la falla bifasica solo depende
         de las secuencias positiva y negativa: Ik2 = raiz(3)*c*Un/|Z1+Z2| y
         Z2 ~ Z1 en una red de distribucion.
      2. Ik1 distinto de Ik3. Si dan EXACTAMENTE iguales en todas las barras,
         es senal de que las impedancias de secuencia cero no estan
         parametrizadas y PowerFactory esta asumiendo Z0 = Z1 por defecto; en
         ese caso las corrientes de falla monofasica NO sirven para disenar la
         malla de puesta a tierra, que depende del camino de secuencia cero.
    """
    # filas_nodos: [Nombre, Caso, Barra_Falla, Tipo_Falla, Ikss, Skss, Atributo]
    valores = {}
    for fila in filas_nodos:
        valores[(fila[1], fila[2], fila[3])] = fila[4]

    app.PrintInfo("--- Verificacion de coherencia fisica de los resultados ---")

    ratios_2 = []
    iguales_1_3 = 0
    comparadas = 0
    for (caso, barra, tipo), valor in valores.items():
        if tipo != "Trifasico" or not valor:
            continue
        ik3 = valor
        ik2 = valores.get((caso, barra, "Bifasico"))
        ik1 = valores.get((caso, barra, "Monofasico"))
        if ik2:
            ratios_2.append(ik2 / ik3)
        if ik1:
            comparadas += 1
            if abs(ik1 / ik3 - 1) < 1e-6:
                iguales_1_3 += 1

    if not ratios_2:
        app.PrintError(
            "  [1] No hay ninguna corriente de falla BIFASICA utilizable: todas vinieron nulas o en cero. "
            "Revisar el diagnostico de atributos impreso mas arriba."
        )
    else:
        promedio = sum(ratios_2) / len(ratios_2)
        fuera = [r for r in ratios_2 if not (0.80 <= r <= 0.95)]
        mensaje = f"  [1] Ik2/Ik3: promedio {promedio:.4f} sobre {len(ratios_2)} barras (esperado ~0.866)."
        if fuera:
            app.PrintError(mensaje + f" {len(fuera)} barras fuera del rango 0.80-0.95 - REVISAR.")
        else:
            app.PrintInfo(mensaje + " OK.")

    if comparadas and iguales_1_3 == comparadas:
        app.PrintError(
            f"  [2] Ik1 es EXACTAMENTE igual a Ik3 en las {comparadas} barras comparadas. Eso indica que las "
            f"impedancias de secuencia cero NO estan parametrizadas (PowerFactory esta usando Z0 = Z1 por "
            f"defecto). Las corrientes de falla monofasica de esta corrida NO son utilizables para el diseno "
            f"de la malla de puesta a tierra."
        )
    elif comparadas:
        app.PrintInfo(
            f"  [2] Ik1 difiere de Ik3 en {comparadas - iguales_1_3} de {comparadas} barras: la secuencia cero "
            f"si esta siendo considerada. OK."
        )


def guardar_workbook_seguro(app, wb, ruta):
    try:
        wb.save(ruta)
        return ruta
    except PermissionError:
        alterna = ruta.with_name(f"{ruta.stem}_{dt.datetime.now().strftime('%Y%m%d_%H%M%S')}{ruta.suffix}")
        wb.save(alterna)
        app.PrintError(
            f"No se pudo guardar en '{ruta.name}' (permiso denegado - seguramente esta abierto en Excel). "
            f"CIERRA ese archivo antes de la proxima corrida. Los resultados de esta corrida se guardaron en: {alterna}"
        )
        return alterna


def exportar_resultados(
    app, filas_resumen, filas_nodos, filas_lineas, filas_trafos_2w, filas_trafos_3w, filas_generadores,
    variations_faltantes,
):
    wb = openpyxl.Workbook()

    def hoja(nombre, headers, filas):
        ws = wb.create_sheet(nombre) if nombre != "Resumen" else wb.active
        ws.title = nombre
        ws.append(headers)
        for cell in ws[1]:
            cell.font = Font(bold=True)
        for fila in filas:
            ws.append(fila)
        ws.auto_filter.ref = ws.dimensions
        for col in ws.columns:
            max_len = max(len(str(c.value)) if c.value is not None else 0 for c in col)
            ws.column_dimensions[col[0].column_letter].width = max_len + 2

    # "Valido" ahora refleja tambien la validacion fisica del resultado (Ikss no
    # nula), no solo el codigo de retorno de ComShc; "Detalle" dice por que fallo
    # o de que atributo se leyo la corriente.
    hoja("Resumen", ["Caso", "Barra_Falla", "Tipo_Falla", "Valido", "Detalle"], filas_resumen)
    hoja(
        "Variations_Faltantes",
        ["Network_Variation_no_encontrada_en_PowerFactory"],
        [[nombre] for nombre in variations_faltantes],
    )
    hoja(
        "Nodos",
        ["Nombre", "Caso", "Barra_Falla", "Tipo_Falla", "Ikss_kA", "Skss_MVA", "Atributo_leido"],
        filas_nodos,
    )
    hoja(
        "Lineas",
        ["Nombre", "Nodo_I", "Nodo_J", "Caso", "Barra_Falla", "Tipo_Falla", "Ikss_bus1_kA", "Ikss_bus2_kA"],
        filas_lineas,
    )
    hoja(
        "Transformadores_2Devanados",
        ["Nombre", "Caso", "Barra_Falla", "Tipo_Falla", "Ikss_HV_kA", "Ikss_LV_kA"],
        filas_trafos_2w,
    )
    hoja(
        "Transformadores_3Devanados",
        ["Nombre", "Caso", "Barra_Falla", "Tipo_Falla", "Ikss_HV_kA", "Ikss_MV_kA", "Ikss_LV_kA"],
        filas_trafos_3w,
    )
    hoja(
        "Generadores",
        ["Nombre", "Barra", "Caso", "Barra_Falla", "Tipo_Falla", "Ikss_kA"],
        filas_generadores,
    )
    # ver nota en Flujo_carga.py: mismo desglose, para separar a simple vista
    # las plantas ya existentes (CasoBase) del escenario con el generador del
    # proyecto activo (Network Variation).
    hoja(
        "Generadores_Existentes",
        ["Nombre", "Barra", "Caso", "Barra_Falla", "Tipo_Falla", "Ikss_kA"],
        [fila for fila in filas_generadores if fila[2] == CASO_BASE],
    )
    hoja(
        "Generadores_Con_Proyecto",
        ["Nombre", "Barra", "Caso", "Barra_Falla", "Tipo_Falla", "Ikss_kA"],
        [fila for fila in filas_generadores if fila[2] != CASO_BASE],
    )

    ruta_final = guardar_workbook_seguro(app, wb, RESULTADOS_PATH)
    app.PrintInfo(f"Resultados exportados a: {ruta_final}")


main()
