"""
Barrido horario de flujo de carga (horas segun Parametros_Demanda.xlsx),
repetido por cada CASO DE RED (caso base + Network Variations listadas en
Network_Variations.xlsx), con resultados organizados por TIPO DE ELEMENTO:
una tabla para nodos, una para lineas, una para transformadores y una para
generadores.

Cada tabla tiene el formato: Nombre | Caso | Hora | <variables tecnicas
del tipo> (formato largo/tidy, facil de filtrar o de llevar a tabla
dinamica).

Modo asistido (igual que test_conexion.py): se corre desde dentro de
PowerFactory, reutilizando el proyecto y caso de estudio ya activos. No
crea objetos de red, solo actualiza plini/qlini de las cargas existentes
y activa/desactiva Network Variations ya creadas en PowerFactory.

Pasos:
  1. Conexion a PowerFactory (proyecto y caso de estudio ya activos). Se
     borran los SVG de resultados/graficos_red/ de corridas anteriores
     (esa carpeta no acumula imagenes viejas entre corridas).
  2. Lectura de la demanda por hora y por carga desde Parametros_Demanda.xlsx,
     y de la lista de Network Variations (IntScheme) existentes en el
     proyecto vs. las que se quieren simular (Network_Variations.xlsx).
  3. Por cada CASO DE RED (CasoBase + cada Network Variation que haga match
     entre el Excel y PowerFactory): se ACTIVA el caso primero, y RECIEN
     DESPUES se listan los elementos calc-relevantes de la red (nodos,
     lineas, transformadores 2 y 3 devanados, generadores, cargas). Esto
     es clave: una Network Variation puede agregar o quitar equipos, asi
     que esa lista cambia segun el caso activo y no se puede calcular una
     sola vez al principio (si se hiciera antes de activar, los equipos
     que solo existen dentro de una Network Variation nunca apareceran
     en los resultados de ese caso).
  4. Por cada HORA dentro de ese caso: actualiza plini/qlini de las
     cargas, ejecuta ComLdf.
  5. Cada vez que se ejecuta ComLdf (converja o no), exporta el diagrama
     unifilar activo a SVG en resultados/graficos_red/, con nombre
     <Proyecto>_<timestamp>_Hora<N>_<Caso>.svg.
  6. Si converge, registra por cada elemento sus variables tecnicas
     (segun el tipo); si no converge, esa combinacion caso/hora queda
     marcada y se continua con la siguiente.
  7. Al terminar, desactiva todas las Network Variations (vuelve al caso
     base) para dejar PowerFactory en un estado limpio.
  8. Exporta a Resultados_Flujo_Carga.xlsx: hojas Resumen, Cargas_Faltantes,
     Variations_Faltantes, Nodos, Lineas, Transformadores_2Devanados,
     Transformadores_3Devanados, Generadores.

Una Network Variation del Excel que no exista en PowerFactory NO detiene
el barrido: se avisa y se excluye (igual que con las cargas faltantes).
Si Network_Variations.xlsx no existe todavia, se corre solo el CasoBase.

La exportacion a SVG usa el comando ComWr de PowerFactory (metodo
documentado por DIgSILENT para exportar graficos via Python), combinado
con SetDesktop.Show() para seleccionar el diagrama de red. Esta
combinacion especifica no se ha podido probar en un proyecto real todavia:
si falla, avisa por PrintError pero NO detiene el resto del barrido.

Variables registradas por tipo de elemento:
  Nodos (ElmTerm):        U_pu (m:u), U_kV (calculado = m:u x uknom, ver
                          nota de unidades abajo), Angulo_deg (m:phiu)
  Lineas (ElmLne):        Nodo_I, Nodo_J (nombres de las barras en cada extremo),
                          Cargabilidad_% (m:loading), I_bus1_kA, I_bus2_kA,
                          P_perdidas_MW, Q_perdidas_Mvar
  Transformadores 2 devanados (ElmTr2): Cargabilidad_% (m:loading), I_HV_kA, I_LV_kA,
                          P_perdidas_MW, Q_perdidas_Mvar
  Transformadores 3 devanados (ElmTr3): Cargabilidad_% (m:loading),
                          I_HV_kA, I_MV_kA, I_LV_kA, P_perdidas_MW, Q_perdidas_Mvar
  Generadores (ElmGenstat / ElmSym / ElmPvsys): Barra (bus1, nombre de
                          la barra a la que esta conectado), P_MW, Q_MW,
                          Cargabilidad_% (m:loading), I_kA

NOTA DE UNIDADES (importante, corregido 2026-09-18): en este proyecto de
PowerFactory, los resultados de FLUJO DE CARGA (m:P, m:Q, m:I de ComLdf)
vienen en kW/kvar/A, NO en MW/Mvar/kA como se asumio originalmente -
verificado cruzando generadores, lineas y el transformador propio contra
calculos fisicos esperados (ver Informe_Estudio_Conexion.tex). Por eso
este script divide esos valores entre 1000 antes de guardarlos, para que
las hojas de resultados queden en MW/Mvar/kA de verdad (consistente con
Cortocircuito.py, cuyas variables de cortocircuito - m:Ikss, m:Ikss:busX -
SI vienen nativamente en kA sin necesidad de ajuste, porque el modulo de
cortocircuito usa una convencion de unidades fija independiente de esta
configuracion del proyecto). Tambien se corrigio U_kV: el atributo m:U
resulto ser tension fase-neutro (fase-fase / raiz(3)), no fase-fase; se
reemplazo por m:u (pu, verificado correcto) multiplicado por la tension
nominal de la barra (uknom).
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
DEMANDA_PATH = BASE_DIR / "inputs" / "Parametros_Demanda.xlsx"
DEMANDA_SHEET = "Demanda"
VARIATIONS_PATH = BASE_DIR / "inputs" / "Network_Variations.xlsx"
VARIATIONS_SHEET = "Network_Variations"
GRAFICOS_DIR = BASE_DIR / "resultados" / "graficos_red"
CASO_BASE = "CasoBase"

# Factor de crecimiento de demanda a aplicar sobre Parametros_Demanda.xlsx (que
# representa el ano t=2026). Cambiar a 1.0104 para correr el escenario del ano
# t+x=2028 (incremento del 1.04% de demanda solicitado por EBSA para esta
# solicitud, ver DOC_REF_EST_1787069892338.pdf pag. 4). Dejar en 1.0 para el
# ano t. El nombre del archivo de resultados incluye el factor usado, para no
# confundir corridas de distintos anos.
ESCALA_DEMANDA = 1.0
ANIO_ETIQUETA = "2026" if ESCALA_DEMANDA == 1.0 else "2028"
RESULTADOS_PATH = BASE_DIR / "resultados" / f"Resultados_Flujo_Carga_{ANIO_ETIQUETA}.xlsx"


def sanitizar_nombre_archivo(texto):
    return re.sub(r'[\\/:*?"<>|]', "_", str(texto))


def limpiar_carpeta_graficos(app):
    """Borra los SVG de corridas anteriores en GRAFICOS_DIR antes de generar
    los de esta corrida, para que la carpeta no acumule imagenes viejas."""
    if not GRAFICOS_DIR.exists():
        return
    borrados = 0
    for archivo in GRAFICOS_DIR.glob("*.svg"):
        try:
            archivo.unlink()
            borrados += 1
        except OSError as exc:
            app.PrintError(f"No se pudo borrar el SVG anterior '{archivo.name}': {exc}")
    app.PrintInfo(f"Limpieza de graficos_red: {borrados} SVG de corridas anteriores borrados.")


def exportar_diagrama_red(app, nombre_proyecto, nombre_caso, hora):
    """Exporta el/los diagramas unifilares (SetDeskpage) actualmente en el
    Graphics Board a SVG, uno por hora/caso simulado.

    Basado en el metodo documentado por DIgSILENT para exportar graficos via
    Python (ComWr con iopt_rd='svg'), combinado con SetDesktop.Show() para
    seleccionar el diagrama de red en vez de un plot (se sustituye 'GrpPage'
    por 'SetDeskpage', segun esa misma documentacion). No verificado en un
    proyecto real todavia: si el nombre de clase de la pagina no coincide en
    esta version de PowerFactory, avisa por PrintError y no detiene el barrido.
    """
    try:
        desktop = app.GetGraphicsBoard()
        paginas = list(desktop.GetContents("*.SetDeskpage", 1)) if desktop is not None else []
        if not paginas:
            app.PrintError("No se encontro ningun diagrama de red (SetDeskpage) abierto para exportar a SVG.")
            return

        etiqueta_caso = nombre_caso if nombre_caso == CASO_BASE else f"NetworkVariation_{nombre_caso}"
        marca_tiempo = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        base = sanitizar_nombre_archivo(f"{nombre_proyecto}_{marca_tiempo}_Hora{hora}_{etiqueta_caso}")

        for pagina in paginas:
            desktop.Show(pagina)
            sufijo = f"_{sanitizar_nombre_archivo(pagina.loc_name)}" if len(paginas) > 1 else ""
            ruta = GRAFICOS_DIR / f"{base}{sufijo}.svg"

            com_wr = app.GetFromStudyCase("ComWr")
            com_wr.SetAttribute("iopt_rd", "svg")
            com_wr.SetAttribute("iopt_savas", 0)
            com_wr.SetAttribute("f", str(ruta))
            com_wr.Execute()
    except Exception as exc:  # no debe detener el barrido por un fallo de exportacion grafica
        app.PrintError(f"No se pudo exportar el diagrama de red a SVG (caso {nombre_caso}, hora {hora}): {exc}")


def leer_demanda(path, sheet_name, escala=1.0):
    """Lee Carga | Hora | P_MW | Q_MW -> {hora: {carga: (P_MW, Q_MW)}}.

    escala multiplica P y Q de todas las cargas (ver ESCALA_DEMANDA) para
    proyectar la demanda del ano t=2026 a otro ano (ej. 1.0104 para t+x=2028)
    sin necesidad de un segundo archivo de insumos."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[sheet_name]
    demanda = {}
    for carga, hora, p_mw, q_mw in ws.iter_rows(min_row=2, values_only=True):
        if hora is None:
            continue
        demanda.setdefault(int(hora), {})[str(carga)] = (float(p_mw) * escala, float(q_mw) * escala)
    return demanda


def emparejar_por_nombre(objetos, nombres, descripcion):
    """Empareja nombres del Excel con objetos de PowerFactory por coincidencia
    de substring en loc_name (ej. '15344-1' con 'Carga_15344-1', o
    'GenSolar1' con una Network Variation llamada igual).

    Devuelve (mapeo, faltantes). Los nombres que todavia no existen en
    PowerFactory (modelo/proyecto en construccion) no detienen el barrido:
    se reportan como faltantes y se excluyen del calculo. Un nombre ambiguo
    (coincide con mas de un objeto) si detiene la ejecucion, porque ahi no
    hay forma segura de saber a cual objeto se refiere."""
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


def leer_nombres_columna(path, sheet_name, columna_header="Nombre"):
    """Lee una columna de nombres (una hoja simple de una sola columna) desde
    un Excel opcional. Si el archivo no existe todavia, devuelve lista vacia
    en vez de fallar (permite correr solo con el caso base)."""
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
    """Deja activa unicamente la Network Variation objetivo (o ninguna, para
    el caso base). Se desactivan primero todas las demas para no mezclar
    variations entre corridas."""
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


def safe_get_mw(obj, attr):
    """Como safe_get, pero convierte kW/kvar/A a MW/Mvar/kA (ver nota de
    unidades en el docstring del modulo): los resultados de flujo de carga
    (ComLdf) en este proyecto vienen en kW/A, no en MW/kA."""
    valor = safe_get(obj, attr)
    return None if valor is None else valor / 1000.0


def nombre_nodo(cubicle):
    """Nombre de la barra (ElmTerm) conectada a una cubicula (bus1/bus2/bushv/...)."""
    try:
        return cubicle.cterm.loc_name
    except Exception:
        return None


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

    # --- lectura de la demanda por hora (no depende de PowerFactory) ---
    demanda = leer_demanda(DEMANDA_PATH, DEMANDA_SHEET, ESCALA_DEMANDA)
    app.PrintInfo(f"Escala de demanda aplicada: {ESCALA_DEMANDA} (ano {ANIO_ETIQUETA})")
    horas = sorted(demanda.keys())
    nombres_carga = sorted({carga for valores in demanda.values() for carga in valores})

    # --- lectura y emparejamiento de Network Variations (caso base siempre se corre) ---
    variations_pf = listar_network_variations(app)
    nombres_variations = leer_nombres_columna(VARIATIONS_PATH, VARIATIONS_SHEET)
    mapeo_variations, variations_faltantes = emparejar_por_nombre(variations_pf, nombres_variations, "Network Variation")
    if variations_faltantes:
        app.PrintError(
            f"Network Variations del Excel que NO existen en PowerFactory (se excluyen, "
            f"no detienen la ejecucion): {variations_faltantes}"
        )

    # se usa el loc_name exacto de PowerFactory (no el texto de busqueda del Excel)
    # para que el nombre del caso en tablas y en el SVG sea el nombre real de la Network Variation
    casos = [(CASO_BASE, None)] + [(obj.loc_name, obj) for obj in mapeo_variations.values()]
    app.PrintInfo(f"Casos de red a simular: {[nombre for nombre, _ in casos]}")

    # medida de seguridad: se desactivan todas las Network Variations que hagan
    # match ANTES de arrancar el barrido, sin importar el estado en que haya
    # quedado PowerFactory (de una corrida anterior, de trabajo manual, etc.)
    app.PrintInfo("Asegurando estado inicial: desactivando Network Variations antes de iniciar el barrido...")
    activar_caso_red(app, CASO_BASE, None, variations_pf)

    ldf = app.GetFromStudyCase("ComLdf")

    filas_resumen = []
    filas_nodos = []
    filas_lineas = []
    filas_trafos_2w = []
    filas_trafos_3w = []
    filas_generadores = []
    cargas_faltantes_todas = []  # [Caso, Carga] - una Network Variation puede agregar cargas que en CasoBase no existen

    for nombre_caso, variacion_obj in casos:
        activar_caso_red(app, nombre_caso, variacion_obj, variations_pf)
        app.PrintInfo(f"--- Caso de red: {nombre_caso} ---")

        # se relee la red DESPUES de activar el caso: una Network Variation puede
        # agregar/quitar equipos (transformadores, cargas, generadores, etc.), asi
        # que la lista de "calc relevant objects" cambia segun el caso activo y
        # NO se puede calcular una sola vez antes del bucle.
        loads = list(app.GetCalcRelevantObjects("*.ElmLod", 1))
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
            f"{len(generadores)} generadores, {len(loads)} cargas."
        )
        # Diagnostico: lista cada generador detectado con su clase, si esta
        # fuera de servicio, y en que barra esta conectado - util para
        # confirmar si un generador que deberia aparecer (ej. una planta ya
        # existente en el circuito) esta siendo excluido por estar en otro
        # caso de red, fuera de servicio, o en un grid no calc-relevante.
        for gen in generadores:
            try:
                bus = gen.GetCubicle(0).cterm.loc_name
            except Exception:
                bus = "?"
            fuera_servicio = bool(safe_get(gen, "outserv"))
            app.PrintInfo(
                f"  [{nombre_caso}] generador detectado: '{gen.loc_name}' "
                f"(clase {gen.GetClassName()}, barra '{bus}', "
                f"fuera de servicio: {fuera_servicio}, ruta: {gen.GetFullName()})"
            )

        mapeo_cargas, cargas_faltantes_caso = emparejar_por_nombre(loads, nombres_carga, "carga")
        app.PrintInfo(f"[{nombre_caso}] cargas mapeadas: { {k: v.loc_name for k, v in mapeo_cargas.items()} }")
        if cargas_faltantes_caso:
            app.PrintError(
                f"[{nombre_caso}] cargas del Excel que NO existen en este caso (se excluyen, "
                f"no detienen la ejecucion): {cargas_faltantes_caso}"
            )
            cargas_faltantes_todas.extend([nombre_caso, nombre] for nombre in cargas_faltantes_caso)

        for hora in horas:
            for nombre, (p_mw, q_mw) in demanda[hora].items():
                carga_obj = mapeo_cargas.get(nombre)
                if carga_obj is None:
                    continue  # carga aun no montada en PowerFactory, ya reportada arriba
                carga_obj.SetAttribute("plini", p_mw)
                carga_obj.SetAttribute("qlini", q_mw)

            err = ldf.Execute()
            convergio = err == 0
            filas_resumen.append([nombre_caso, hora, convergio])

            exportar_diagrama_red(app, nombre_proyecto, nombre_caso, hora)

            if not convergio:
                app.PrintError(f"Caso {nombre_caso}, hora {hora}: FLUJO DE CARGA NO CONVERGIO (codigo {err}).")
                continue

            app.PrintInfo(f"Caso {nombre_caso}, hora {hora}: flujo de carga convergio.")

            for term in terminals:
                u_pu = safe_get(term, "m:u")
                uknom = safe_get(term, "uknom")
                u_kv = None if (u_pu is None or uknom is None) else u_pu * uknom
                filas_nodos.append([
                    term.loc_name, nombre_caso, hora,
                    u_pu, u_kv, safe_get(term, "m:phiu"),
                ])

            for ln in lines:
                p_perdidas = (safe_get_mw(ln, "m:P:bus1") or 0) + (safe_get_mw(ln, "m:P:bus2") or 0)
                q_perdidas = (safe_get_mw(ln, "m:Q:bus1") or 0) + (safe_get_mw(ln, "m:Q:bus2") or 0)
                filas_lineas.append([
                    ln.loc_name, nombre_nodo(ln.bus1), nombre_nodo(ln.bus2), nombre_caso, hora,
                    safe_get(ln, "m:loading"),
                    safe_get_mw(ln, "m:I:bus1"), safe_get_mw(ln, "m:I:bus2"),
                    p_perdidas, q_perdidas,
                ])

            for tr in trafos_2w:
                p_perdidas = (safe_get_mw(tr, "m:P:bushv") or 0) + (safe_get_mw(tr, "m:P:buslv") or 0)
                q_perdidas = (safe_get_mw(tr, "m:Q:bushv") or 0) + (safe_get_mw(tr, "m:Q:buslv") or 0)
                filas_trafos_2w.append([
                    tr.loc_name, nombre_caso, hora,
                    safe_get(tr, "m:loading"),
                    safe_get_mw(tr, "m:I:bushv"), safe_get_mw(tr, "m:I:buslv"),
                    p_perdidas, q_perdidas,
                ])

            for tr in trafos_3w:
                p_perdidas = (
                    (safe_get_mw(tr, "m:P:bushv") or 0) + (safe_get_mw(tr, "m:P:busmv") or 0) + (safe_get_mw(tr, "m:P:buslv") or 0)
                )
                q_perdidas = (
                    (safe_get_mw(tr, "m:Q:bushv") or 0) + (safe_get_mw(tr, "m:Q:busmv") or 0) + (safe_get_mw(tr, "m:Q:buslv") or 0)
                )
                filas_trafos_3w.append([
                    tr.loc_name, nombre_caso, hora,
                    safe_get(tr, "m:loading"),
                    safe_get_mw(tr, "m:I:bushv"), safe_get_mw(tr, "m:I:busmv"), safe_get_mw(tr, "m:I:buslv"),
                    p_perdidas, q_perdidas,
                ])

            for gen in generadores:
                filas_generadores.append([
                    gen.loc_name, nombre_nodo(gen.bus1), nombre_caso, hora,
                    safe_get_mw(gen, "m:P:bus1"), safe_get_mw(gen, "m:Q:bus1"),
                    safe_get(gen, "m:loading"), safe_get_mw(gen, "m:I:bus1"),
                ])

    # deja PowerFactory en el caso base al terminar
    activar_caso_red(app, CASO_BASE, None, variations_pf)

    exportar_resultados(
        app, filas_resumen, filas_nodos, filas_lineas, filas_trafos_2w, filas_trafos_3w, filas_generadores,
        cargas_faltantes_todas, variations_faltantes,
    )
    app.PrintInfo("===================================")


def guardar_workbook_seguro(app, wb, ruta):
    """Guarda el workbook; si el archivo esta bloqueado (ej. abierto en Excel),
    no se pierden los resultados calculados: se guarda una copia con timestamp
    junto al archivo original y se avisa claramente por que paso."""
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
    cargas_faltantes, variations_faltantes,
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

    hoja("Resumen", ["Caso", "Hora", "Convergio"], filas_resumen)
    hoja(
        "Cargas_Faltantes",
        ["Caso", "Carga_no_encontrada_en_PowerFactory"],
        cargas_faltantes,
    )
    hoja(
        "Variations_Faltantes",
        ["Network_Variation_no_encontrada_en_PowerFactory"],
        [[nombre] for nombre in variations_faltantes],
    )
    hoja("Nodos", ["Nombre", "Caso", "Hora", "U_pu", "U_kV", "Angulo_deg"], filas_nodos)
    hoja(
        "Lineas",
        ["Nombre", "Nodo_I", "Nodo_J", "Caso", "Hora", "Cargabilidad_%", "I_bus1_kA", "I_bus2_kA", "P_perdidas_MW", "Q_perdidas_Mvar"],
        filas_lineas,
    )
    hoja(
        "Transformadores_2Devanados",
        ["Nombre", "Caso", "Hora", "Cargabilidad_%", "I_HV_kA", "I_LV_kA", "P_perdidas_MW", "Q_perdidas_Mvar"],
        filas_trafos_2w,
    )
    hoja(
        "Transformadores_3Devanados",
        ["Nombre", "Caso", "Hora", "Cargabilidad_%", "I_HV_kA", "I_MV_kA", "I_LV_kA", "P_perdidas_MW", "Q_perdidas_Mvar"],
        filas_trafos_3w,
    )
    hoja(
        "Generadores",
        ["Nombre", "Barra", "Caso", "Hora", "P_MW", "Q_Mvar", "Cargabilidad_%", "I_kA"],
        filas_generadores,
    )
    # Mismos datos que "Generadores", mas faciles de leer separados: en
    # "Generadores_Existentes" (Caso == CasoBase, sin la Network Variation del
    # proyecto activa) solo deberian aparecer las plantas que YA estan en el
    # circuito (ej. GD1/GD2); en "Generadores_Con_Proyecto" (Network Variation
    # activa) aparecen esas mismas plantas MAS el generador del proyecto.
    hoja(
        "Generadores_Existentes",
        ["Nombre", "Barra", "Caso", "Hora", "P_MW", "Q_Mvar", "Cargabilidad_%", "I_kA"],
        [fila for fila in filas_generadores if fila[2] == CASO_BASE],
    )
    hoja(
        "Generadores_Con_Proyecto",
        ["Nombre", "Barra", "Caso", "Hora", "P_MW", "Q_Mvar", "Cargabilidad_%", "I_kA"],
        [fila for fila in filas_generadores if fila[2] != CASO_BASE],
    )

    ruta_final = guardar_workbook_seguro(app, wb, RESULTADOS_PATH)
    app.PrintInfo(f"Resultados exportados a: {ruta_final}")


main()
