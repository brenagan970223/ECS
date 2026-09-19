# Estudio de Conexión GD 990kW — EBSA (Puerto Boyacá)

Automatización en Python del estudio de conexión simplificado (circular CREG 021 de 2022 / CREG 174 de 2021) para la solicitud de generación distribuida (990 kW, Puerto Boyacá, circuito 15344), sobre PowerFactory 2024.

Todo el cálculo de ingeniería (flujo de carga, tensiones, corrientes, cargabilidad, pérdidas) lo hace **PowerFactory**. Python solo orquesta: lee insumos de Excel, mueve parámetros hacia PowerFactory, dispara el flujo de carga, y exporta lo que PowerFactory calculó.

## Modo de ejecución: asistido

Todos los scripts que tocan PowerFactory se ejecutan **manualmente, dentro de PowerFactory** (Data Manager → Python Script → ejecutar), reutilizando el proyecto y el caso de estudio que ya tengas abiertos en la GUI. Ninguno se lanza en modo "engine" (headless) desde una terminal externa — eso requeriría cerrar cualquier instancia gráfica abierta y no es el flujo de trabajo que usamos aquí.

## Estructura de carpetas

```
ECS/
├── scripts/       Código Python (ver abajo)
├── inputs/        Insumos en Excel (los editas tú a mano)
├── resultados/    Todo lo que generan los scripts (se sobrescribe en cada corrida)
├── informe/       Estructura y seguimiento del informe de estudio de conexión (CREG 174)
├── env_pf/        Entorno virtual Python 3.12 que usa PowerFactory para correr estos scripts
└── powerfactory-tools/   Clon de referencia (github.com/ieeh-tu-dresden/powerfactory-tools),
                          consultado para ver convenciones de la API de PowerFactory
                          (nombres de atributos, clases, comandos). No se usa como dependencia:
                          esa librería asume modo "engine" (GetApplicationExt), incompatible
                          con nuestro modo asistido (GetApplication).
```

## Scripts (`scripts/`)

### Nota: GD1 y GD2 (generadores ya existentes en el circuito 15344)
EBSA reporta (`DOC_REF_EST_1787069892338.pdf`, "Generadores Distribuidos Conectados: 2") dos generadores ya existentes en el circuito — GD1 (330kVA, barra P4_15344) y GD2 (330kVA, barra P5_15344) — que no estaban modelados. **A diferencia de los demás elementos de este proyecto, estos se montan a mano directamente en PowerFactory** (no vía script — decisión explícita del usuario), en la red base `Red_Puerto_Boyaca_15344` (no en `Network_Variations.xlsx`, porque ya existen independientemente de si el proyecto nuevo se conecta). Ver `informe/Parametros_Proyecto.xlsx` (fila `GD_existentes_circuito`) para la ficha de características genéricas acordadas (cosφ=1, despacho igual al perfil solar propio escalado a 330kW, aporte a cortocircuito = "No Short-Circuit Contribution"). **Una vez montados manualmente**, hay que repetir `Flujo_carga.py`/`Cortocircuito.py` — todos los resultados previos, calculados sin GD1/GD2, quedan desactualizados.

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
- Los nombres de atributo de corriente de aporte en líneas/transformadores/generadores (`m:Ikss:bus1`, `m:Ikss:bushv`, etc.) son una extrapolación del patrón `:busX` que sí está validado para flujo de carga — no confirmados todavía en un proyecto real. Si salen vacíos, avisa para ajustarlos.
- Exporta un SVG por cada combinación caso × barra fallada × tipo de falla (uno por cada cálculo de `ComShc`), en `resultados/graficos_red_corto/` — son muchos archivos, confirmado así con el usuario.

## Inputs (`inputs/`)

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

**Flujo de carga:**
1. Editar `inputs/Parametros_Demanda.xlsx` (año base) y/o `inputs/Network_Variations.xlsx` si cambia algo. Si cambia el horizonte o la tasa de crecimiento, editar `ESCENARIOS_ANIO` en `Flujo_carga.py`.
2. Correr `Flujo_carga.py` dentro de PowerFactory — **una sola vez**: corre los dos años seguidos, escalando la demanda por sí solo.
3. Revisar `resultados/Resultados_Flujo_Carga.xlsx` (filtrar por la columna `Anio`) y `resultados/graficos_red/`.

**Cortocircuito:**
1. Editar `inputs/Network_Variations.xlsx` si aplica (el método de cálculo IEC 60909/VDE 0102 ya lo fuerza el script, no requiere configuración previa).
2. Correr `Cortocircuito.py` dentro de PowerFactory.
3. Revisar `resultados/Resultados_Cortocircuito.xlsx` y `resultados/graficos_red_corto/`.
