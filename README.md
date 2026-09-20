# Estudio de Conexión GD 990kW — EBSA (Puerto Boyacá)

Automatización en Python del estudio de conexión simplificado (circular CREG 021 de 2022 / CREG 174 de 2021) para la solicitud de generación distribuida (990 kW, Puerto Boyacá, circuito 15344), sobre PowerFactory 2024.

Todo el cálculo de ingeniería (flujo de carga, tensiones, corrientes, cargabilidad, pérdidas) lo hace **PowerFactory**. Python solo orquesta: lee insumos de Excel, mueve parámetros hacia PowerFactory, dispara el flujo de carga, y exporta lo que PowerFactory calculó.

---

## 📊 Estado del proyecto

*Última actualización: 2026-09-19*

**El estudio técnico está cerrado.** Las simulaciones están corridas y validadas, el informe redactado y los anexos armados. Lo que falta es material de terceros y la compilación final.

| Análisis | Resultado | Estado |
|---|---|---|
| 8.1 Validación de modelación | Desviación +0.129 % en tensión (límite 10 %) | ✅ |
| 8.3 Perfiles de tensión | 0.9363–1.0115 p.u. (límite 0.90–1.10) | ✅ |
| 8.3 Nivel de carga | Máximo 80.3 % (límite 100 %) | ✅ |
| 8.4 Cortocircuito | Máximo 8.164 kA (capacidad 10 kA) | ✅ |
| 8.5 Funcionamiento en isla | Déficit de 3.975 MW: isla inviable | ✅ |
| 8.6 Pérdidas | Reduce 2.28 kW frente a carga pura | ✅ |
| 9.2 Coordinación | No requiere reajuste de cabecera | ✅ |

**Simulaciones:** 72 escenarios de flujo de carga (2 años × 3 casos × 12 horas) y 93 de cortocircuito. Todos convergieron.

**Checklist del informe:** 19 secciones Completo · 1 Parcial · 3 No aplica
**Causales de rechazo:** 6 Cumple · 1 Pendiente · 2 No aplica

---

## 🚧 Lo que falta

### 🔴 Bloqueante para radicar

**1. Compilar los anexos a PDF.**
Los tres anexos están escritos en LaTeX y con todas sus figuras convertidas, pero **no hay PDF** porque esta máquina no tiene motor de LaTeX instalado. Dos caminos:

- **Overleaf:** sube la carpeta `informe/anexos/` completa (los tres `.tex` + la subcarpeta `figuras/`) y compila cada uno con pdfLaTeX.
- **Local:** instala un motor de LaTeX (`brew install basictex`) y vuelve a correr `python3 scripts/Anexos.py` — el script detecta `pdflatex` y compila solo.

**2. Verificar el plazo de radicación.**
EBSA emitió el documento de insumos el **11/08/2026** y otorga **5 meses**, es decir hasta **~11/01/2027**. Nadie ha confirmado esta fecha contra el Anexo 5 de CREG 174. Es causal de rechazo automática, independiente de toda la ingeniería.

### 🟠 Datos de terceros pendientes

| # | Qué falta | A quién | Impacto |
|---|---|---|---|
| 1 | **Protocolo de pruebas del transformador** (impedancia real) | Ryctel / Tesla | Hoy va con `u_k = 6 %` **asumido**. Al recibirlo hay que sustituirlo y re-verificar 8.4 |
| 2 | **Curva real del reconectador** de cabecera | EBSA | Hoy se usan los ajustes de la ficha (pág. 4). Sirve para la coordinación definitiva |
| 3 | **¿El recierre tiene supervisión de tensión?** | EBSA | Se asumió que **no** (hipótesis conservadora). Confirmarlo solo puede mejorar el margen |
| 4 | **Marca y modelo del relé** + su parametrización | Tú / proveedor | Es el único `\pendiente{}` que queda en el informe |

### 🟡 Revisión manual en Overleaf

El informe compila sin errores (entornos balanceados, sin referencias rotas, todas las imágenes presentes), pero conviene revisar a mano:

- Saltos de página y ubicación de figuras flotantes
- Que las tablas largas no se partan en mal sitio
- La numeración de secciones tras los cambios recientes

---

## 📄 Páginas y anexos: qué existe y qué no

### Informe principal — `informe/Informe_Estudio_Conexion.tex`

**~1290 líneas. Completo salvo un marcador.**

| Sección | Estado |
|---|---|
| 1. Resumen ejecutivo | ✅ Redactado |
| 2. Descripción y ubicación | ✅ Con mapa satelital extraído de la memoria |
| 3. Parámetros eléctricos (Tablas 1–5) | ✅ |
| 4. Información de entrada y supuestos | ✅ Incluye modelación del despacho solar |
| 5.1–5.6 Análisis (8.1 a 8.6) | ✅ Con 6 figuras y su análisis |
| 5.7–5.8 Estabilidad y evaluación económica | ✅ No aplican, justificado |
| 6. Conclusiones | ✅ 6 conclusiones + 6 recomendaciones |
| Anexo de protecciones | ⚠️ Queda 1 `\pendiente{}`: hoja de ajustes del relé |
| Checklist de causales de rechazo | ✅ |

**Imágenes que necesita** (todas en `informe/`, ninguna falta):
`ubicacion_proyecto.jpg` · `diagrama_unifilar.pdf` · `fig_demanda_cargas.png` · `fig_demanda_generacion.png` · `fig_tension.png` · `fig_cargabilidad.png` · `fig_perdidas.png` · `fig_cortocircuito.png` · `Curva_TCC_Cabecera_15344.png`

### Anexos — `informe/anexos/`

**Los tres existen en `.tex` con todas sus figuras. Falta compilarlos.**

| Anexo | Líneas | Contenido | PDF |
|---|---|---|---|
| **A — Flujo de carga** | 2426 | Metodología, los 72 escenarios, curvas de demanda, 3 figuras de resultados, **3 diagramas unifilares** y tablas completas de nodos, líneas, transformadores y generadores | ❌ |
| **B — Cortocircuito** | 1665 | Método IEC 60909, parametrización de secuencia cero, niveles de falla de las 16 barras × 3 tipos, **3 diagramas de falla** y tablas de aportes | ❌ |
| **C — Coordinación** | 221 | Ajustes evaluados, curva TCC, verificación de selectividad y margen anti-isla | ❌ |

`informe/anexos/figuras/` contiene **15 archivos**: los 6 diagramas de PowerFactory ya convertidos de SVG a PDF y las 9 gráficas.

> **Por qué los diagramas están en PDF y no en SVG:** pdfLaTeX no lee SVG. `scripts/Anexos.py` los convierte con `svglib`.

> **Por qué no están los 165 diagramas:** la simulación exporta uno por escenario (72 de flujo + 93 de cortocircuito). Incluirlos todos haría el anexo inmanejable, así que se selecciona un subconjunto representativo — configurable en las constantes `ESCENARIOS_FLUJO` y `BARRAS_CORTO` de `Anexos.py` — y el propio anexo deja constancia de cuántos hay en total y dónde están los demás.

---

## ✅ Correcciones ya aplicadas (no rehacer)

Problemas detectados y resueltos, para que no se vuelvan a levantar:

- **Falla bifásica en 0 kA** — `Cortocircuito.py` leía `m:Ikss`, que PowerFactory no puebla en fallas desbalanceadas. Ahora lee las variables por fase. Ratio Ik2/Ik3 = 0.8660 exacto (√3/2).
- **Secuencia cero sin parametrizar** — la falla monofásica salía idéntica a la trifásica. Corregido con `Z0/Z1 = 4.867` en el equivalente de 115 kV, validado contra el dato del OR (2.622 kA).
- **Impedancia del transformador** — el modelo tenía ~3.9 %, irreal para 1250 kVA. Ajustada a 6 % típico.
- **Generación plana** — el despacho no seguía ningún perfil, lo que impedía construir los escenarios 8.2. Ahora sigue una campana de Gauss y existe el caso de carga pura.
- **Conclusión de cortocircuito invertida** — el informe decía que la memoria *sobreestima* 1.5×; en realidad la **subestima un 26 %** (11.2 vs 14.164 kA).
- **Parámetros hardcodeados** — trasladados a `inputs/Parametros_Sistema.xlsx`.

---

## 📋 Correcciones pendientes en la memoria de cálculo

La memoria se elaboró antes del estudio y con supuestos simplificados. **Cuando ambos documentos discrepen, manda el estudio de conexión.**

El listado completo de las 10 correcciones a trasladar está en **`informe/Correcciones_Memoria_Calculo.md`**. La más relevante: la memoria **subestima** el cortocircuito en 800 V (11.2 kA frente a los 14.164 kA reales).

## Modo de ejecución: asistido

Todos los scripts que tocan PowerFactory se ejecutan **manualmente, dentro de PowerFactory** (Data Manager → Python Script → ejecutar), reutilizando el proyecto y el caso de estudio que ya tengas abiertos en la GUI. Ninguno se lanza en modo "engine" (headless) desde una terminal externa — eso requeriría cerrar cualquier instancia gráfica abierta y no es el flujo de trabajo que usamos aquí.

## Estructura de carpetas

```
ECS/
├── scripts/       Código Python (ver abajo)
├── inputs/        Insumos en Excel (los editas tú a mano)
├── resultados/    Todo lo que generan los scripts (se sobrescribe en cada corrida)
├── informe/       El informe LaTeX, sus figuras y los anexos
│   └── anexos/    Los tres anexos en LaTeX + figuras/ con los diagramas convertidos a PDF
├── env_pf/        Entorno virtual Python 3.12 que usa PowerFactory (no versionado)
└── powerfactory-tools/   Clon de referencia (github.com/ieeh-tu-dresden/powerfactory-tools),
                          consultado para ver convenciones de la API de PowerFactory
                          (nombres de atributos, clases, comandos). No se usa como dependencia:
                          esa librería asume modo "engine" (GetApplicationExt), incompatible
                          con nuestro modo asistido (GetApplication).
```

## Scripts (`scripts/`)

### Nota: GD1 y GD2 (generadores ya existentes en el circuito 15344)
EBSA reporta (`DOC_REF_EST_1787069892338.pdf`, "Generadores Distribuidos Conectados: 2") dos generadores ya existentes en el circuito, de 330 kVA cada uno. **Se montaron a mano directamente en PowerFactory** (no vía script — decisión explícita), en la red base `Red_Puerto_Boyaca_15344`, porque existen independientemente de si el proyecto nuevo se conecta.

**Ya están montados y todos los resultados actuales los incluyen.** Según el modelo, `GD P4 15344` está conectado en la barra **P5 15344 13.2kV** y `GD P5 15344` en **P6 15344 13.2kV** (confirmado contra el unifilar: los nombres no coinciden con las barras, pero la conexión es la correcta).

Características acordadas: cosφ=1, despacho según el mismo perfil solar del proyecto escalado a su potencia, y aporte a cortocircuito = "No Short-Circuit Contribution" — por eso aparecen con `Ikss = 0` en todos los resultados de falla, lo cual es correcto y no un error.

### `test_conexion.py`
Prueba mínima: conecta a PowerFactory (`GetApplication()`), confirma que hay un proyecto activo. Úsalo cuando algo no conecta, para descartar problemas de PowerFactory antes de correr el script grande.

### `Flujo_carga.py`
El script principal. Por cada **año del horizonte** (año t y año t+x, ver abajo), por cada **caso de red** (`CasoBase` + cada Network Variation que definas en `inputs/Network_Variations.xlsx`) y por cada **hora** (las que traiga `inputs/Parametros_Demanda.xlsx`):

1. Activa ese caso de red (activa la Network Variation correspondiente, o ninguna para `CasoBase`).
2. **Recién después de activar el caso**, relee los elementos de la red (nodos, líneas, transformadores, generadores, cargas) — no antes. Esto es a propósito: una Network Variation puede agregar o quitar equipos, así que la lista de equipos "calc-relevantes" depende de qué caso esté activo en ese momento.
3. Actualiza `plini`/`qlini` (P/Q) de cada carga según la demanda de esa hora, **multiplicada por el factor de crecimiento del año que se esté simulando**.
4. Ejecuta el flujo de carga (`ComLdf`).
5. Exporta el diagrama unifilar activo a SVG (`resultados/graficos_red/`), converja o no.
6. Si converge, registra tensiones, corrientes, cargabilidad y pérdidas de cada elemento.

Al terminar cada caso devuelve `plini`/`qlini` a su valor original (así el modelo no queda con la demanda proyectada del último año como si fuera su estado normal), al final desactiva todas las Network Variations (vuelve a `CasoBase`), y guarda todo en `resultados/Resultados_Flujo_Carga.xlsx`.

**Escenario año t vs. año t+x — automático, en una sola corrida.** No hay que editar constantes ni relanzar el script: la lista `ESCENARIOS_ANIO` al inicio del archivo define los años y su factor de crecimiento de demanda, y el barrido los recorre todos seguidos, escalando por sí solo la demanda de cada carga del sistema.

```python
ESCENARIOS_ANIO = [
    (2026, 1.0),      # año t   - demanda tal cual en Parametros_Demanda.xlsx
    (2028, 1.0104),   # año t+x - demanda del año t incrementada 1.04%
]
```

El factor es **acumulado respecto al año base**, no anual compuesto: `1.0104` es el incremento total del 1.04% que pidió EBSA para esta solicitud entre 2026 y 2028. Si el OR cambia el horizonte o la tasa, se edita solo esa lista — agregar o quitar filas `(año, factor)` basta para que el barrido corra esos años. El Excel de insumos nunca se toca: siempre representa el año base.

Los resultados de todos los años quedan en **un solo archivo**, `resultados/Resultados_Flujo_Carga.xlsx`, diferenciados por la columna **`Anio`** que lleva cada hoja (con autofiltro activado). La hoja `Escenarios_Anio` deja registrado qué factor se aplicó a cada año, para que la trazabilidad del escalado quede en el mismo archivo de resultados.

El cortocircuito (`Cortocircuito.py`) no necesita nada de esto — con el método IEC 60909/VDE 0102 usado, el resultado no depende del nivel de demanda, así que no se repite por año.

**Corrección de unidades (2026-09-18)**: en este proyecto, las variables de resultado de `ComLdf` (`m:P:busX`, `m:Q:busX`, `m:I:busX`) vienen en **kW/kvar/A**, no en MW/Mvar/kA como es lo habitual en PowerFactory — el script ahora las convierte dividiendo entre 1000 (función `safe_get_mw`). Aparte, `m:U` en `ElmTerm` resultó ser tensión fase-neutro, no fase-fase; `U_kV` ahora se calcula como `m:u` (p.u., ya venía correcto) × `uknom` (tensión nominal). Antes de este fix, los resultados de `Flujo_carga.py` estaban mal por un factor de 1000 (P/Q/I) y de √3 (U_kV) — **no era un error del modelo en PowerFactory, era un error de interpretación de unidades en el script**. `Cortocircuito.py` nunca tuvo este problema: sus variables (`m:Ikss`, `m:Ikss:busX`) siempre vinieron en kA nativamente. Si corriste `Flujo_carga.py` antes de esta fecha, vuelve a correrlo — los resultados viejos en `resultados/Resultados_Flujo_Carga_*.xlsx` no son confiables.

**Tolerante a un modelo incompleto**: si una carga o una Network Variation del Excel todavía no existe en PowerFactory (en ese caso de red particular), no detiene el barrido — la excluye, avisa por consola, y la deja registrada en las hojas `Cargas_Faltantes` (con el `Caso` en que faltó) / `Variations_Faltantes` del resultado.

**Emparejamiento por nombre**: las cargas y Network Variations se buscan en PowerFactory por coincidencia de nombre (substring), no por una lista fija en el código. Si cambias la topología o agregas cargas nuevas, solo tienes que reflejarlo en los Excel de `inputs/` — no hay que tocar el script.

### `Cortocircuito.py`
Análisis de cortocircuito (CREG 174 de 2021, numeral 8.4 — contribución a la corriente de cortocircuito) en **todas las barras del sistema**, para 3 tipos de falla (**trifásico, bifásico, monofásico a tierra**) y por cada **caso de red** (mismo esquema `CasoBase` + Network Variations que `Flujo_carga.py`, incluyendo la relectura de elementos después de activar cada caso).

Por cada caso: por cada barra, por cada tipo de falla, ejecuta `ComShc` con esa barra como punto de falla y registra la contribución de corriente de cada elemento. Se ejecuta una sola vez manualmente y de ahí corre desatendido sobre todas las combinaciones. Guarda en `resultados/Resultados_Cortocircuito.xlsx`.

**Advertencia de transparencia** (léela antes de confiar en los resultados): el comando `ComShc`, los tipos de falla (`3psc`/`2psc`/`spgf`) y `m:Ikss`/`m:Skss` en barras están confirmados contra documentación y ejemplos reales de scripting de PowerFactory. Pero:
- **Método de cálculo**: se fuerza `iopt_mde = 0` (IEC 60909 / VDE 0102 Part 0, DIN EN 60909-0) en cada ejecución de `ComShc` — es el método por defecto de esta instalación, confirmado con el usuario. El script imprime esto en la Output Window de PowerFactory al arrancar.
- Los nombres de atributo de corriente de aporte en líneas/transformadores/generadores (`m:Ikss:bus1`, `m:Ikss:bushv`, etc.) **ya están confirmados** en corridas reales: las hojas `Lineas` y `Generadores` traen valores.
- **Falla bifásica:** `m:Ikss` no se puebla en fallas desbalanceadas, así que el script prueba una lista de atributos candidatos (`ATRIBUTOS_IKSS_BARRA`) y cae a las variables por fase. La columna `Atributo_leido` de la hoja `Nodos` registra de dónde salió cada valor.
- **Diagnóstico de secuencia:** al arrancar, el script recorre el modelo y reporta qué datos de secuencia tiene cada elemento, en qué objeto viven y cuál hay que cambiar, con ruta completa. Al terminar verifica dos relaciones físicas: `Ik2/Ik3 ≈ 0.866` y que `Ik1 ≠ Ik3`.
- Exporta un SVG por cada combinación caso × barra fallada × tipo de falla (uno por cada cálculo de `ComShc`), en `resultados/graficos_red_corto/` — son muchos archivos, confirmado así con el usuario.

### `parametros.py` — lectura centralizada de parámetros

No se ejecuta solo: lo importan los demás scripts. Lee `inputs/Parametros_Sistema.xlsx` y expone los valores por su nombre:

```python
from parametros import P
P.cabecera("PICKUP_51_A")      # 240.0
P.red_or("CAPACIDAD_CORTE_KA") # 10.0
P.elemento("BARRA_PC")         # "P1 15344 13.2kV"
```

**Por qué existe:** antes, valores como los ajustes de protección (240 A, dial 0.1, 1500 A), la capacidad de corte de la subestación o los nombres de las barras estaban escritos dentro de cada script, a veces repetidos en varios. Para cambiar un ajuste había que buscarlo en el código, y para saber de dónde salía un número no había forma. Ahora se edita el Excel. Si falta una clave, el error dice exactamente qué hoja y qué fila revisar.

### `Graficas_Informe.py` — figuras del informe

Se corre como Python normal. Lee los Excel de resultados y genera seis figuras en `informe/`:

| Figura | Qué muestra |
|---|---|
| `fig_demanda_cargas.png` | Las 8 cargas del área, agrupadas por circuito |
| `fig_demanda_generacion.png` | Demanda agregada vs. generación, con los escenarios 8.2 marcados |
| `fig_tension.png` | Perfiles en el punto de conexión y en el extremo del circuito |
| `fig_cargabilidad.png` | Tramo de evacuación y transformador del proyecto |
| `fig_perdidas.png` | Pérdidas totales y balance frente a carga pura |
| `fig_cortocircuito.png` | Nivel de falla por barra vs. capacidad de corte |

Usa una paleta categórica validada para fondo claro y verificada contra deuteranopia y tritanopia. Los tres casos de red conservan el mismo color en todas las figuras, para no tener que releer la leyenda en cada una.

**Dos decisiones de presentación que conviene conocer**, porque cambian lo que muestra cada figura:
- **Tensión:** se grafican dos barras concretas (P1 y P6) en vez de la envolvente mín/máx del sistema. La envolvente cambia de barra de una hora a otra, así que produce una curva quebrada que no describe a ningún elemento real.
- **Cargabilidad:** se grafica la línea de evacuación y no la más cargada del sistema. La más cargada (`Al 0.82km`) pertenece a otro circuito y da el mismo valor en los tres casos — las tres series quedaban superpuestas y la figura no mostraba nada. El valor del elemento más cargado va anotado aparte.

### `Anexos.py` — los tres anexos en LaTeX

Se corre como Python normal, **después** de `Graficas_Informe.py` y `Coordinacion_Protecciones.py`. Genera en `informe/anexos/`:

- `Anexo_A_Flujo_Carga.tex`, `Anexo_B_Cortocircuito.tex`, `Anexo_C_Coordinacion_Protecciones.tex`
- `figuras/` con los diagramas de PowerFactory convertidos de SVG a PDF (pdfLaTeX no lee SVG) y las gráficas copiadas

Usa el mismo preámbulo que el informe principal, para que salgan con idéntica tipografía y márgenes. Si hay `pdflatex` en la máquina los compila; si no, deja los `.tex` listos para Overleaf y lo avisa.

## Inputs (`inputs/`)

### `Parametros_Sistema.xlsx` — parámetros que antes estaban en el código

Siete hojas, cada fila con **valor, unidad, descripción y fuente documental**:

| Hoja | Qué contiene |
|---|---|
| `Proteccion_Cabecera` | Ajustes reales del reconectador de EBSA (curva, pickups, diales, instantáneos, recierre) |
| `Proteccion_Proyecto` | Ajustes propuestos para el punto de conexión y criterios de margen |
| `Datos_Proyecto` | Potencia, inversores, tensiones, datos del transformador, aporte de falla |
| `Red_OR` | Capacidad de corte, cortocircuito del equivalente de 115 kV, límites regulatorios |
| `Elementos_Clave` | Nombre exacto de cada barra y elemento en el modelo de PowerFactory |
| `Despacho_Solar` | Hora pico, sigma de la campana, si se simula el caso de carga pura |
| `Anios_Analisis` | Los años del horizonte y su factor de demanda |

**Para cambiar un ajuste se edita aquí, no en el código.**

### `Parametros_Demanda.xlsx` — hoja `Demanda`
Demanda horaria por carga. Columnas: `Carga | Hora | P_MW | Q_MW`. Formato largo (una fila por carga+hora), pensado para filtrar/dinamizar fácil en Excel.
Fuente: tabla "CARGA PTO BOYACA 34,5kV" de `DOC_REF_EST_1787069892338.pdf` (pág. 4), transcrita y verificada valor por valor. Cubre las 8 cargas del sistema (15302-1/2, 15303-1/2, 15340-1/2, 15344-1/2) para horas 6 a 17 — no solo las que ya estén montadas en PowerFactory.

### `Network_Variations.xlsx` — hoja `Network_Variations`
Columna: `Nombre`. Lista de las Network Variations (`IntScheme`) que quieres simular además del caso base. Si el archivo no existe o está vacío, el barrido corre solo `CasoBase`.

### `Insumos_Red_15344.xlsx`
Insumos de referencia del montaje original del circuito 15344 (barras, líneas, transformador, cargas hora 6, equivalente de red). Histórico — ya no lo consume ningún script, queda como documentación de los datos de partida.

## Outputs (`resultados/`)

### `Resultados_Flujo_Carga.xlsx`
**Un solo archivo con todos los años del horizonte.** Una hoja por tipo de elemento, formato largo `Nombre | Anio | Caso | Hora | <variables>` — para separar años se filtra por la columna `Anio` (2026 / 2028), no por archivo:

| Hoja | Variables |
|---|---|
| `Resumen` | `Anio, Caso, Hora, Convergio, Factor_Demanda` |
| `Escenarios_Anio` | `Anio, Factor_Demanda_aplicado, Variacion_%_vs_ano_base, Descripcion` — deja registrado en el propio archivo de resultados con qué factor se escaló cada año |
| `Nodos` | `U_pu, U_kV, Angulo_deg` |
| `Lineas` | `Nodo_I, Nodo_J, Cargabilidad_%, I_bus1_kA, I_bus2_kA, P_perdidas_MW, Q_perdidas_Mvar` |
| `Transformadores_2Devanados` | `Cargabilidad_%, I_HV_kA, I_LV_kA, P_perdidas_MW, Q_perdidas_Mvar` |
| `Transformadores_3Devanados` | igual, con lado MV adicional |
| `Generadores` | `Barra` (nodo al que está conectado), `P_MW, Q_Mvar, Cargabilidad_%, I_kA` — todos los generadores (`ElmGenstat`/`ElmSym`/`ElmPvsys`) de cada caso, en una sola tabla larga |
| `Generadores_Existentes` | mismos datos, filtrados solo a `Caso == CasoBase` (Network Variation del proyecto desactivada) — deberían aparecer únicamente las plantas que ya existen en el circuito (ej. GD1/GD2), no la del proyecto |
| `Generadores_Con_Proyecto` | mismos datos, filtrados a `Caso != CasoBase` — las plantas existentes más el generador del proyecto |
| `Cargas_Faltantes` / `Variations_Faltantes` | qué del Excel no se encontró en PowerFactory (las cargas faltantes se reportan por año y caso) |

Si el archivo está abierto en Excel al momento de guardar, el script no pierde los resultados: guarda una copia con timestamp y avisa que cierres el original.

> Los archivos `Resultados_Flujo_Carga_2026.xlsx` y `Resultados_Flujo_Carga.xlsx` que existían antes son de corridas viejas (uno por año / anterior al fix de unidades). Quedan reemplazados por este archivo único en la primera corrida del script actualizado.

### `graficos_red/`
Un SVG por cada combinación año × caso × hora simulada, nombre `<Proyecto>_<timestamp>_Anio<AAAA>_Hora<N>_<Caso>.svg` (o `..._NetworkVariation_<Nombre>.svg` cuando hay una variation activa). El año va en el nombre porque la carpeta se limpia una sola vez al arrancar: sin él, los diagramas del segundo año se confundirían con los del primero.

### `Resultados_Cortocircuito.xlsx`
Formato largo `Nombre | Caso | Barra_Falla | Tipo_Falla | <variables>`:

| Hoja | Variables |
|---|---|
| `Resumen` | `Caso, Barra_Falla, Tipo_Falla, Valido` |
| `Nodos` | `Ikss_kA, Skss_MVA` (solo de la barra realmente fallada — mismo número de filas que `Resumen`, no una fila por cada barra del sistema) |
| `Lineas` | `Nodo_I, Nodo_J, Ikss_bus1_kA, Ikss_bus2_kA` |
| `Transformadores_2Devanados` | `Ikss_HV_kA, Ikss_LV_kA` |
| `Transformadores_3Devanados` | `Ikss_HV_kA, Ikss_MV_kA, Ikss_LV_kA` |
| `Generadores` | `Barra` (nodo al que está conectado), `Ikss_kA` |
| `Generadores_Existentes` / `Generadores_Con_Proyecto` | mismo desglose que en `Resultados_Flujo_Carga.xlsx`, filtrado por `Caso == CasoBase` / `Caso != CasoBase` |
| `Variations_Faltantes` | qué Network Variation del Excel no se encontró en PowerFactory |

### `graficos_red_corto/`
Un SVG por cada combinación caso × barra fallada × tipo de falla (uno por cada cálculo de cortocircuito).

## Informe de conexión (`informe/`)

### `Estructura_Informe_Conexion.xlsx`
Checklist de todas las secciones que debe tener el informe final, según los "Lineamientos y contenido estudio de conexión simplificado" de CREG 174 de 2021 (numerales 6 a 10 y Anexo I). El proyecto cae en la ruta **"Solicitudes N2"** (GD < 1MW, conexión en 13.2kV) — no aplican estabilidad (8.7) ni despachos de generación (7.4, solo N3/N4).

- **`Estructura_Informe`**: cada sección del informe con su numeral CREG 174, qué contenido pide, de dónde sale el dato (o qué script), y su `Estado` (Completo / Parcial / Falta / No aplica — con color).
- **`Escenarios_8.2`**: los 3 escenarios que exige CREG 174 (carga pura, máxima diferencia, máxima demanda+generación) y cómo identificarlos dentro de `Parametros_Demanda.xlsx` — **no son simplemente correr las 12 horas**, hay que mapear cada uno a una hora/condición específica.
- **`Causales_Rechazo`**: la Tabla 6 de CREG 174 (motivos de rechazo del estudio) como checklist de autoverificación antes de radicar.

Se actualiza a mano a medida que se completa cada parte del análisis.

### `Parametros_Proyecto.xlsx`
Cuestionario reutilizable con los datos del proyecto recopilados en conversación (año t/t+x, coordenadas, vida útil, datos de placa, certificaciones, hallazgos técnicos, etc.), hoja `Cuestionario_Proyecto` con columnas `Parametro | Pregunta | Valor | Unidad | Estado (Respondida/Pendiente) | Fuente | Notas`. Sirve como bitácora central: cualquier dato que se resuelva en conversación se guarda aquí antes de reflejarse en el `.tex`, así no se repite la pregunta ni se pierde el porqué de cada decisión.

### `Informe_Estudio_Conexion.tex`
Informe completo en LaTeX (listo para pegar en Overleaf, compilador pdfLaTeX), con la misma estructura que `Estructura_Informe_Conexion.xlsx`. Incluye portada, tabla de control de versiones, datos de placa completos (panel/inversor/transformador), curva de generación mensual y anual con degradación del panel, gráfica de demanda vs. generación con los escenarios 8.2 marcados (`curva_demanda_generacion.png` — **debe subirse junto con el .tex a Overleaf**), justificación técnica de por qué no se hace análisis N-1 (topología radial), metodología detallada de cálculo de cortocircuito por cadena de impedancias (IEC 60909, verificada contra PowerFactory), y la corrección exacta a aplicar en PowerFactory para el aporte a cortocircuito del generador. Las secciones que aún faltan están marcadas con `\pendiente{...}` / `\parcial{...}` — resaltadas en el texto y listadas al inicio del documento (`\listoftodos`). Incluye un anexo para el estudio de coordinación de protecciones (siguiente paso del proyecto).

## Comprobación de Coordinación de Protecciones (EACP simplificado) — `scripts/Coordinacion_Protecciones.py`

**Estado: implementado y corrido (2026-09-18).** Es una **metodología programada**: no hay ajuste manual ni cálculo subjetivo — el script aplica la fórmula normalizada IEC 60255-151 automáticamente sobre los datos ya calculados y validados, así que es reproducible cada vez que cambien los resultados de `Cortocircuito.py`.

**A diferencia de `Flujo_carga.py`/`Cortocircuito.py`, este script se corre como Python normal, NO dentro de PowerFactory** (no toca el modelo, solo lee el Excel de resultados ya exportado):
```
./env_pf/Scripts/python.exe scripts/Coordinacion_Protecciones.py
```

**Qué es y qué NO es**: el Acuerdo CNO 2121 de 2026 (que sustituye a los Acuerdos 2087/1862/1749/...) exige EACP a partir de 0.1 MW, así que un proyecto de 990 kW sí aplica. Pero un EACP completo cubre todas las funciones de protección de la planta (27/59/50/51/50N/51N/59N/ROCOF/anti-isla, etc.) en varios niveles de tensión. Lo que se pide aquí es más acotado: una **comprobación puntual** de si la protección de cabecera del circuito 15344 (el reconectador de la S/E Puerto Boyacá, ya ajustado y en operación) sigue coordinando bien una vez se agrega el aporte de falla del proyecto — no un rediseño de esa protección ni de las protecciones internas de la planta (esas ya están definidas por el fabricante del inversor/tablero, ver certificación anti-isla en el informe, sección 8.5).

**Por qué es viable hacerlo como comprobación y no como estudio completo**: el Anexo 1 del Acuerdo 2121 confirma que los sistemas de generación basados en inversores **no superan 1.1 p.u. de su corriente nominal como aporte a la falla** (pág. de "Sistema de interrupción" del anexo) — es una limitación de la electrónica de potencia, no del ajuste. Como ya validamos con `Cortocircuito.py` que el inversor está configurado así (`Ik"3PF=238.2A` por inversor ≈ 1.08 p.u. de su corriente nominal), el aporte de falla del proyecto es pequeño y acotado por diseño frente a los cientos/miles de amperios de la red de EBSA — es exactamente el caso donde una comprobación dirigida (¿el aporte nuevo saca la curva de cabecera de su zona de operación esperada?) tiene sentido, en vez de rehacer la coordinación completa del circuito.

**Datos de entrada y por qué esos y no otros**:
1. **Corriente de cortocircuito en la barra de cabecera del circuito 15344** (o la barra calc-relevante más cercana a esa protección en el modelo), trifásica y monofásica, en los casos `CasoBase` (sin proyecto) y `Con_Proyecto` (Network Variation activa) — sale de `resultados/Resultados_Cortocircuito.xlsx`, hoja `Nodos`. Es el dato correcto porque es exactamente lo que "ve" el reconectador ante una falla: si el nivel de falla con proyecto se mantiene dentro del mismo tramo de la curva (mismo tiempo de disparo, sin cruzar el instantáneo 50/50N de forma distinta a como ya opera hoy), la coordinación no se altera. No se recalcula a mano porque PowerFactory ya lo tiene con el modelo completo de la red (más preciso que una cadena de impedancias simplificada).
2. **Curva ajustada de la protección de cabecera 15344** (ya en `Parametros_Proyecto.xlsx`, fila `Condiciones_operativas`, tomada de `DOC_REF_EST_1787069892338.pdf` pág. 4): curva IEC Normal Inverse, Pickup51=240A, Pickup51N=60A, Dial51=0.1, Dial51N=0.1, instantáneo 50=1500A, 50N=300A, recierre rápido=2s. Es el dato correcto porque es el ajuste **real, ya aprobado y operando** de EBSA — no un ajuste propuesto por nosotros; la comprobación es "¿este ajuste ya existente sigue sirviendo con el proyecto conectado?", no "¿qué ajuste habría que poner?".
3. **Aporte de falla configurado en el inversor** (`Ik"3PF=238.2A`, `Ik"2PF=Ik"1PF≈0` por inversor, ×3 inversores) y su resultado de simulación validado (`Ikss=0.7146kA` trifásico en la barra propia del generador) — de `Parametros_Proyecto.xlsx` y `Resultados_Cortocircuito.xlsx`. Es el dato pedido explícitamente para la sección de "aporte de falla del sistema de generación" y quedará documentado como el valor a usar después en el diseño de la malla de puesta a tierra (el Anexo 1 del Acuerdo 2121 liga explícitamente el diseño de puesta a tierra con los sistemas de protección propuestos).
4. **Requisito de equipo de corte según potencia** (Anexo 1, Acuerdo 2121): para 0.25 MW < P < 1 MW se exige "Interruptor con unidades de disparo, Reconectador o Interruptor de potencia" en el PC, con capacidad de interrupción acorde al nivel de cortocircuito del estudio de conexión — esto conecta con el hallazgo pendiente de 9.2 (no hay documentado un equipo de corte en MV/13.2kV del transformador propio) y se retoma aquí como una de las conclusiones/recomendaciones.

**Salida esperada del módulo**:
- Una gráfica TCC (tiempo vs. corriente, log-log) con la curva ajustada de cabecera 15344 (fase 51/50 y neutro 51N/50N, fórmula IEC 60255 Normal Inverse) y dos líneas verticales marcando la Ikss trifásica en cabecera sin proyecto vs. con proyecto.
- Una tabla comparativa (kA y tiempo de disparo correspondiente) sin proyecto vs. con proyecto.
- La sección de aporte de falla (punto 3 arriba) con el valor a trasladar a puesta a tierra.
- Conclusión corta: si el desplazamiento de corriente cae dentro de la misma banda de tiempo de disparo → no se requiere reajuste de cabecera; si no → qué parámetro (pickup/dial) requeriría revisión, y recomendaciones (ej. equipo de corte en el PC del punto 4).
- Todo esto se integra en el Anexo del `.tex` (`Informe_Estudio_Conexion.tex`, sección "Comprobación de coordinación de protecciones en cabecera del circuito 15344").

**Fórmula usada (IEC 60255-151, "IEC Normal Inverse" — la misma familia con la que EBSA ya ajustó la protección de cabecera)**:

```
t(I) = Dial * k / ( (I / Ipickup)^alpha - 1 ),   I > Ipickup      k = 0.14, alpha = 0.02
```

Si `I` supera el umbral instantáneo (50/50N), el relé no espera la curva temporizada: opera en `~0.03s` (tiempo definido corto, típico de un reconectador). Es la misma fórmula que el reconectador de cabecera ya tiene programada — no se propone ni calibra nada nuevo, solo se evalúa con las corrientes de falla simuladas.

**Resultado de la corrida actual** (con los datos ya disponibles en `resultados/Resultados_Cortocircuito.xlsx`, barra `P1 15344 13.2kV`):

| Falla | Sin proyecto | Con proyecto | Δ | Zona (ambos) | Conclusión |
|---|---|---|---|---|---|
| Trifásica | 4112 A | 4155 A | +1.05% | Instantáneo (50) | Mismo tiempo de disparo (0.03s) → no requiere reajuste |
| Monofásica | 4112 A | 4118 A | +0.13% | Instantáneo (50N) | Mismo tiempo de disparo (0.03s) → no requiere reajuste |

El aporte de falla propio del sistema de generación (`Ik"3PF=238.2A` por inversor = exactamente 1.00 p.u. de su corriente nominal, `330kW/(√3×800V)`) queda muy por debajo del límite regulatorio de 1.1 p.u. (Anexo 1, Acuerdo CNO 2121 de 2026), y el `Ikss=0.7146kA` trifásico validado en la barra propia del generador queda documentado como insumo para el diseño de la malla de puesta a tierra.

**Salidas** (todo en `resultados/Coordinacion_de_Protecciones/`): `Curva_TCC_Cabecera_15344.png` (curva TCC con los 4 puntos evaluados) y `Resultados_Coordinacion_Protecciones.xlsx` (hojas `Comparativa_Cabecera_15344`, `Aporte_Falla_Generacion`, `Conclusion`).

**Restricción de siempre**: este módulo solo lee resultados ya exportados (`Resultados_Cortocircuito.xlsx`) y datos de la ficha de EBSA — no toca el modelo de PowerFactory ni monta elementos.

## Flujo de trabajo típico

**Orden completo, de principio a fin:**

```
# 1. Dentro de PowerFactory (modo asistido)
Flujo_carga.py          # 72 escenarios: 2 años × 3 casos × 12 horas
Cortocircuito.py        # arranca con el diagnóstico de secuencia cero

# 2. Fuera de PowerFactory (Python normal)
python3 scripts/Coordinacion_Protecciones.py   # lee el Excel de cortocircuito
python3 scripts/Graficas_Informe.py            # genera las 6 figuras del informe
python3 scripts/Anexos.py                      # genera los 3 anexos en LaTeX
```

Los tres últimos **no tocan PowerFactory**: solo leen los Excel de resultados. Si cambia algo del modelo, hay que repetir la secuencia completa en ese orden, porque cada paso consume lo que produjo el anterior.

**Flujo de carga en detalle:**
1. Editar `inputs/Parametros_Demanda.xlsx` (año base) y/o `inputs/Network_Variations.xlsx` si cambia algo. El horizonte, la tasa de crecimiento y el perfil de despacho se editan en `inputs/Parametros_Sistema.xlsx`.
2. Correr `Flujo_carga.py` dentro de PowerFactory — **una sola vez**: corre los dos años y los tres casos de red seguidos.
3. Revisar `resultados/Resultados_Flujo_Carga.xlsx` (filtrar por `Anio` y `Caso`) y `resultados/graficos_red/`.

**Cortocircuito:**
1. Editar `inputs/Network_Variations.xlsx` si aplica (el método de cálculo IEC 60909/VDE 0102 ya lo fuerza el script, no requiere configuración previa).
2. Correr `Cortocircuito.py` dentro de PowerFactory.
3. Revisar `resultados/Resultados_Cortocircuito.xlsx` y `resultados/graficos_red_corto/`.
