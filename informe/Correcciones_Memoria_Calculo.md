# Correcciones a aplicar a la Memoria de Cálculo

**Proyecto:** Sistema Solar Fotovoltaico On-Grid 990 kW — Puerto Boyacá, circuito 15344 (EBSA)
**Solicitud:** 16869
**Documento de origen:** `informe/Memoria de Calculo/Memoria de calculo.pdf`
**Última actualización de este listado:** 2026-09-19

---

## Para qué sirve este documento

La memoria de cálculo se elaboró **antes** del estudio de conexión y con supuestos
simplificados. El estudio de conexión, en cambio, se hizo con el modelo completo
de la red de EBSA en PowerFactory, con los insumos oficiales del OR.

**Cuando los dos documentos discrepen, manda el estudio de conexión.** Este archivo
recoge, una por una, las correcciones que hay que trasladar a la memoria para que
ambos documentos queden coherentes antes de radicar.

Cada ítem indica: qué dice la memoria, qué dice el estudio, y por qué el estudio
tiene la razón.

---

## 1. Corriente de cortocircuito en la barra de 800 V — **la memoria la SUBESTIMA**

| | Valor | Método |
|---|---|---|
| **Memoria** | 11.2 kA | Solo el transformador propio (1250 kVA, Z=8 %), asumiendo bus infinito del lado de 13.2 kV |
| **Estudio de conexión** | **14.164 kA** | Red completa modelada en PowerFactory (equivalente 115 kV + T9 + líneas del 15344 + transformador propio con Z=6 %), método IEC 60909 |

**Diferencia: el valor real es un 26 % MÁS ALTO que el de la memoria.**

Esto es importante y va en contra de la intuición habitual. Se suele asumir que
calcular solo con el transformador es "conservador" porque ignora la impedancia de
la red aguas arriba. Pero aquí el efecto dominante es otro: **la memoria usó Z=8 %
y la impedancia real adoptada es 6 %**, y una impedancia menor produce una corriente
de falla mayor. Ese efecto pesa más que la impedancia de la red.

> ⚠️ **Verificar el dimensionamiento de todo el equipo de baja tensión** contra
> 14.164 kA y no contra 11.2 kA. Según la cotización Ryctel C26-688, el barraje del
> tablero de BT tiene aisladores certificados a 25 kA y el interruptor ACB tiene
> Icu = 50 kA a 1000 V: **ambos siguen cumpliendo**, pero el margen es menor del
> que suponía la memoria.

---

## 2. Impedancia del transformador de elevación

| | Valor |
|---|---|
| **Memoria** | Z = 8 % |
| **Estudio de conexión** | **Z = 6 %** (valor típico normalizado adoptado) |

**Estado: supuesto, no dato medido.** A la fecha **no se dispone del protocolo de
pruebas de laboratorio del transformador**. La cotización del proveedor (Ryctel
C26-688) es un documento comercial y no incluye la impedancia.

Se adoptó 6 % por ser el valor típico de mercado para un transformador de
distribución trifásico de 1250 kVA en aceite a 13.2 kV. El modelo de PowerFactory
tenía originalmente una impedancia implícita de ~3.9 %, que no es realista.

> 🔴 **ACCIÓN PENDIENTE (ambos documentos):** cuando el fabricante entregue el
> protocolo de pruebas, sustituir el 6 % por el valor medido en los dos documentos
> y volver a verificar las conclusiones de cortocircuito.

---

## 3. Datos de placa del transformador — completar

La memoria no documenta el grupo de conexión. Según la cotización Ryctel C26-688,
ítem 3 (referencia `TRF-1250-13200-800-A-K1`, marca Tesla):

- Potencia: **1250 kVA**
- Relación: **13200 / 800-462 V**
- Grupo de conexión: **Dyn5** ← dato que la memoria no tenía
- Tipo: aceite convencional, devanado en aluminio, pantalla electrostática, factor K1
- Incluye DPS y termómetro de dos contactos

El grupo **Dyn5** no es un detalle menor: el devanado en delta del lado de 13.2 kV
**bloquea la secuencia cero**, lo que significa que el proyecto no aporta corriente
a las fallas a tierra del alimentador de EBSA. Eso condiciona todo el esquema de
protección de tierra (ver ítem 7).

---

## 4. Punto de conexión — inconsistencia a resolver

La memoria y algunos cálculos manuales sitúan el punto de conexión en la barra
**P3**. El modelo de PowerFactory lo tiene en **P1 15344 13.2 kV**, que es la barra
de cabecera del circuito.

La diferencia no es cosmética: cambia la impedancia vista desde el proyecto y, por
tanto, el nivel de cortocircuito disponible.

- Icc trifásica en **P1** = 4.112 kA (sin proyecto)
- Icc trifásica en **P3** = 1.236 kA

> 🔴 **ACCIÓN PENDIENTE:** verificar en terreno / con EBSA cuál es el punto de
> conexión real y unificar el dato en los dos documentos. Los resultados de este
> estudio corresponden a **P1**.

---

## 5. Generadores distribuidos ya existentes en el circuito

EBSA reporta en su documento de insumos (`DOC_REF_EST_1787069892338.pdf`, pág. 4)
que el circuito 15344 ya tiene **2 generadores distribuidos conectados**, de
330 kVA cada uno.

La memoria de cálculo **no los considera**. El estudio de conexión sí los modela,
y su presencia afecta:

- la capacidad disponible en el punto de conexión,
- el perfil de tensión del tramo,
- el balance de potencia para el análisis anti-isla.

> ⚠️ Revisar si algún cálculo de la memoria (especialmente de regulación de tensión)
> debe rehacerse teniéndolos en cuenta.

---

## 6. Aporte del proyecto a la corriente de falla

Valor validado en el estudio de conexión, a trasladar a la memoria:

| Concepto | Valor |
|---|---|
| Aporte por inversor (Ik″3PF) | 238.2 A |
| Corriente nominal del inversor | 238.2 A → **1.00 p.u.** |
| Cantidad de inversores | 3 |
| **Aporte total en 800 V** | **714.6 A (0.7146 kA)** |
| Aporte referido a 13.2 kV | 43.3 A |

Cumple con holgura el límite regulatorio de **1.1 p.u.** que fija el Anexo 1 del
Acuerdo CNO 2121 de 2026 para generación basada en inversores.

**Dato clave para el diseño de protecciones:** en media tensión el proyecto solo
aporta 43.3 A a una falla. Eso significa que sus protecciones de sobrecorriente
**no pueden detectar fallas en la red de EBSA** — la desconexión ante fallas
externas depende íntegramente de las funciones de tensión, frecuencia y anti-isla.

---

## 7. Corriente de falla a tierra — **pendiente de recalcular**

| | Valor |
|---|---|
| **Memoria** (`If_13.2kV_referencia_memoria`) | 2.799 kA |
| **Estudio de conexión** | ⏳ **Pendiente** |

**Por qué está pendiente:** el modelo de PowerFactory tiene las impedancias de
secuencia cero sin parametrizar, de modo que la corriente de falla monofásica sale
exactamente igual a la trifásica en las 15 barras de media tensión. Eso no es un
resultado físico válido.

Se está corrigiendo con el dato que el propio EBSA entrega (Ik1 = 2.622 kA en el
barraje de 115 kV), fijando Z0/Z1 = 4.867 en el equivalente de red.

> 🔴 **ACCIÓN PENDIENTE:** cuando se complete esa corrección, trasladar aquí el
> valor definitivo. **Es el insumo para el diseño de la malla de puesta a tierra**,
> así que el diseño del SPT no debe cerrarse hasta tener este número.

---

## 8. Celda de media tensión — cambio de equipo

La cotización Ryctel C26-688 ofrecía originalmente una **celda de seccionador
fusible** de 17 kV / 16 kA con fusibles tipo HH (ítem 2, referencia `CF-17KV-SIE`).

**Esa configuración no cumple** el Anexo 1 del Acuerdo CNO 2121 de 2026, que para
proyectos entre 0.25 MW y 1 MW exige en el punto de conexión un *interruptor con
unidades de disparo, reconectador o interruptor de potencia*.

> ✅ **DECISIÓN TOMADA:** se adopta la **celda con reconectador e interruptor con
> unidades de disparo**, opción que el mismo proveedor ofrece.
>
> ⚠️ La cotización advierte que **no incluye la parametrización del relé**. Esa
> actividad debe contratarse aparte, y los ajustes deben tomarse de la sección
> "Sistema de protecciones del proyecto en el punto de conexión" del informe de
> estudio de conexión.

---

## 9. Perfil de generación para análisis de despacho

La memoria calcula energía anual a partir de horas pico solar equivalentes, lo cual
es correcto para estimación energética. Pero el estudio de conexión necesita el
**perfil horario de potencia**, que es distinto.

Se adoptó una **campana de Gauss centrada en la hora 12**, con σ = 2.5, aplicada de
forma proporcional a la potencia nominal de cada generador. Los factores resultantes
son:

| Hora | 6 | 7 | 8 | 9 | 10 | 11 | 12 | 13 | 14 | 15 | 16 | 17 |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| Factor | 0.056 | 0.135 | 0.278 | 0.487 | 0.726 | 0.923 | **1.000** | 0.923 | 0.726 | 0.487 | 0.278 | 0.135 |

> ⚠️ Si la memoria contiene un perfil horario distinto, unificar ambos o justificar
> la diferencia.

---

## 10. Degradación del panel

**Criterio adoptado en el estudio:** la garantía de rendimiento del fabricante a
**30 años** (87.4 % de la potencia inicial), en lugar de un modelo de degradación
lineal por tramos.

Es el criterio más defendible ante el OR porque se apoya en un compromiso
contractual verificable y coincide con la vida útil declarada del proyecto.

> ⚠️ Verificar que la memoria use el mismo criterio en su cálculo de energía.

---

## Resumen de acciones

| # | Acción | Bloqueante | Depende de |
|---|---|---|---|
| 1 | Actualizar Icc en 800 V: 11.2 → **14.164 kA** | Sí | — |
| 2 | Sustituir Z=8 % por el valor del protocolo de pruebas | Sí | Fabricante |
| 3 | Añadir grupo de conexión **Dyn5** y datos de placa | No | — |
| 4 | Unificar punto de conexión (**P1** vs P3) | Sí | Verificación en terreno |
| 5 | Incorporar GD1/GD2 a los cálculos que aplique | No | — |
| 6 | Documentar aporte de falla: **714.6 A / 1.0 p.u.** | No | — |
| 7 | Recalcular corriente de falla a tierra | **Sí** | Corrección de secuencia cero |
| 8 | Cambiar celda de MT a reconectador | Sí | Proveedor |
| 9 | Unificar perfil horario de generación | No | — |
| 10 | Verificar criterio de degradación a 30 años | No | — |
