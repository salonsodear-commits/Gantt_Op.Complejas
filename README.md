# Diagrama de Gantt — Proyecto Datos Op. Complejas (versión mejorada)

Herramienta de seguimiento de proyecto convertida en un tablero **dinámico, profesional, escalable y fácil de mantener**, respetando al 100 % la planificación original.

> **Archivo entregable:** `Diagrama_de_Gantt___Proyecto_Datos_Op._Complejas_080626.xlsx`
> **Documentación (Excel aparte):** `Guia_Diagrama_de_Gantt.xlsx`
> **Macro opcional (write-back + log):** `macro_sprints.bas`
> **Respaldo original intacto:** `original_backup.xlsx`
> **Generador reproducible:** `build_xlsx.py`

Hojas del libro: `Proyecto Datos (OKR´s)` (Gantt original), `Tablero`,
`Entregas Próximas`, `Sprints`, y `_Datos` (motor, oculta) — más las
3 hojas ocultas originales que no se tocaron.

---

## 1. Principio rector: cero cambios destructivos

Todas las mejoras son **aditivas**. No se modificó ningún dato, fecha, nombre de tarea,
responsable, estado, entregable, sprint ni fórmula existente.

La edición se hizo a nivel XML del `.xlsx` (no se reabrió con librerías que reescriben el
archivo) para **preservar byte a byte** el contenido original: imagen incrustada,
comentarios, `customXml`, complementos web (`webextensions`), configuración de impresión,
formato condicional original del Gantt y los **nombres definidos relativos**
(`task_start`, `task_end`, `task_progress`) que dan vida a las barras del cronograma.

Sobre la hoja Gantt original (`Proyecto Datos (OKR´s)`) el único cambio es la **adición**
de reglas de formato condicional. Se verificó por diff que no se quitó ni alteró nada más.

---

## 2. Qué se agregó

### A) Formato condicional sobre el Gantt (Parte 1)
- **Semáforo en la columna F (Fecha de entrega)** para filas de tarea:
  - Vencido · Próximo a vencer (≤ 7 días) · En plazo · Completado.
- **Escala de color por Sprint** (columna E): diferenciación visual automática que se
  oscurece a medida que crece el número de sprint.

### B) Hoja **Tablero** (resumen ejecutivo — Parte 4)
- KPIs: Avance global, Tareas, Completadas, En curso, Próximas (≤7 d) y Vencidas.
- **Avance por Módulo (OKR)** con barra de datos. *El % se recalcula de forma
  independiente y correcta* (no usa las celdas de promedio de las filas de fase).
- **Resumen por Sprint** (Parte 2) con barra de datos: Sprint · Tareas · Pendientes ·
  Vencidas · Próx. a vencer · Próxima entrega · % Avance. La lista es **dinámica**: solo
  aparecen los sprints que existen y se agregan solos al usarse, conservando el histórico.

### C) Hoja **Entregas Próximas** (Parte 2)
- Lista de **detalle de entregas** ordenada por Sprint y luego por Fecha de entrega, con
  semáforo de estado. Columnas: Sprint · Entregable (Módulo) · Tarea · Responsable · Fecha
  de entrega · Estado · Días restantes · Progreso.
- La pantalla queda **inmovilizada solo en el encabezado** (6 filas) para ver ~25
  entregas a la vez; el resto se desplaza.
- Franja superior con totales (entregas, vencidas, próximas, avance, próxima entrega).
  El **resumen por sprint** está en la hoja *Tablero*.
- **Umbral “Próximo a vencer”** editable (celda `I3`, por defecto 7 días).

### D) Documentación **Guía** (archivo Excel aparte)
`Guia_Diagrama_de_Gantt.xlsx`: criterios, fórmulas, leyenda de colores, escalabilidad y
observaciones detectadas. Se entrega como libro independiente.

### E) Hoja **_Datos** (oculta)
Motor de cálculo que lee el Gantt y prepara los datos (incluye el listado dinámico de
sprints existentes). **No editar.**

---

## 3. Criterios y fórmulas clave

- **Fecha de entrega** = columna `F` (Fecha) si existe; si está vacía se usa `L` (FIN)
  como respaldo.
- **Estado** (semáforo):
  - `Completado`: Progreso ≥ 100 %.
  - `Vencido`: fecha de entrega anterior a HOY y Progreso < 100 %.
  - `Próximo a vencer`: faltan entre 0 y el umbral de días.
  - `En curso`: dentro de plazo.
- **Tarea vs. fase**: las filas de fase tienen *texto* en la columna INICIO (K) y las
  tareas tienen una *fecha*. `EsTarea = Y(hay tarea; no es fase; hay fecha)`.
- **Módulo de cada tarea**: `LOOKUP(2;1/(...))` localiza el último título de fase por
  encima de la fila.
- **Ordenamiento sin macros**: `SMALL + MATCH + INDEX` sobre una *clave de orden* =
  `Sprint × 1.000.000 + Fecha + Fila/100.000`.
- **Métricas**: `COUNTIF/COUNTIFS`, `MINIFS` y `AVERAGEIF`.
- El libro fuerza recálculo al abrir (`fullCalcOnLoad`), por lo que todos los valores se
  completan automáticamente en Excel.

---

## 4. Escalabilidad (Parte 3)

Todas las fórmulas abarcan hasta la **fila 200** de la hoja Gantt. Se pueden agregar
sprints, tareas, entregables, responsables y fechas nuevos: el **Tablero** y
**Entregas Próximas** se actualizan solos, sin tocar fórmulas. Para superar las 200 filas,
basta con ampliar el rango en la hoja `_Datos`.

**Sprints dinámicos:** la lista de sprints del *Tablero* muestra únicamente los que
existen en los datos. Cuando una tarea pase a usar el Sprint 2, este aparecerá
automáticamente junto al Sprint 1 (el histórico no se pierde). Las tareas sin número de
sprint se agrupan en “(Sin sprint)”.

---

## 5. Observaciones detectadas (informadas, **no modificadas**)

- **Promedios de fase**: las celdas `J27` y `J36` (fases 4 y 5) promedian rangos que no
  corresponden a sus tareas. El Tablero ya recalcula el avance por módulo de forma
  correcta. Si se desea corregir el origen: `J27 → =PROMEDIO(J28:J35)` y
  `J36 → =PROMEDIO(J37:J47)`.
- **Fechas a revisar** (posibles errores de carga): `K13` (19/11/2026, posterior a su FIN),
  `K40` (01/12/2026) y filas con INICIO mayor que FIN (41, 42, 46, 47), que producen DÍAS
  negativos en la columna N.
- **Filas 79–80**: textos sueltos sin fecha; quedan excluidos automáticamente de los cálculos.

---

## 6. Reproducir

```bash
pip install openpyxl
python3 build_xlsx.py   # lee original_backup.xlsx y genera el archivo mejorado + la Guía
```

---

# Requerimientos adicionales (v3)

Todo lo siguiente es **aditivo** y se calcula a partir de los campos ya existentes
(`INICIO` K, `FIN` L, `Fecha` F, `Progreso` J, `Asignado` I) + `HOY()` + un calendario de
sprints editable. No se agregó ninguna columna a la hoja Gantt (no se desplaza la línea de
tiempo): toda la lógica vive en la hoja oculta `_Datos`.

## 1) Gestión dinámica de tareas
- **Cómo agregar una tarea:** escribir en una fila nueva del Gantt — `Tarea` (G), `Inicio`
  (K) y `Fecha` (F) y/o `Fin` (L). *Recomendado:* insertar la fila **dentro de un módulo**
  para que Excel copie las fórmulas de esa fila.
- **Se refleja solo en:** las barras del Gantt (se **amplió el formato condicional de
  barras a las filas 56–200**, reutilizando los estilos de barra originales `dxfId 6/7/8`),
  el `Tablero`, `Entregas Próximas` y el `Historial`. Todos los rangos llegan a la fila
  200, así que **no hay que ajustar fórmulas**.
- **Estructura de datos:** sin cambios en el Gantt. En `_Datos` se detecta cada tarea con
  `EsTarea = Y(hay Tarea; no es fila de fase; hay fecha)` y se le asigna un **ID** estable
  `="T"&TEXTO(fila)`.

## 2) Priorización avanzada en "Entregas Próximas"
Columnas nuevas (todas calculadas): `Dur.est = FinPlan−Inicio+1`, `Días real = HOY−Inicio`,
`Avance esperado = (HOY−Inicio)/(FinPlan−Inicio)`, `Variación = Progreso−Esperado`,
`Atraso(d) = HOY−FechaObjetivo`, `Desvío plan = FinPlan−FechaObjetivo`.

**Score de prioridad (0–100)** y clasificación automática:
```
Score = 35%·Urgencia + 30%·Atraso + 20%·Riesgo + 15%·Carga
≥70 Crítica · ≥45 Alta · ≥25 Media · resto Baja
```
- *Urgencia*: por días restantes (vencida=1; ≤3d=0,9; ≤7=0,7; ≤14=0,45; resto 0,2).
- *Atraso*: `MEDIANA(0; MAX(−Variación; Atraso/30); 1)`.
- *Riesgo*: bloqueo=1; (avance<50% y ≤7d)=0,7; resto 0,3.
- *Carga*: `MEDIANA(0; (PendientesDelResponsable−4)/8; 1)`.

La lista del detalle se **ordena por prioridad** (`LARGE` + `MATCH` + `INDEX`).

**Cuellos de botella** (columna *Alerta*): `ATRASO`, `BLOQUEO` (debía iniciar y sigue en
0 %), `SOBRECARGA` (responsable con > 8 pendientes), `DESVÍO-PLAN`. *El modelo no tiene
columna de dependencias; el bloqueo se infiere del inicio vencido sin avance.* El `Tablero`
suma estas alertas en el panel **Alertas / Cuellos de botella**.

## 3) Sprints por reunión (1 sprint = 1 reunión de cierre)
Siguiendo el concepto del adjunto: **un sprint es un período de tiempo que cierra en una
reunión**, no una tarea. En la hoja **Sprints** hay un **calendario editable** (Sprint 1, 2,
3… con su *fecha de cierre* = la reunión):
```
Sprint actual  = primer sprint cuya reunión aún no pasó
               = MÍN(nº de sprints; CONTAR.SI(cierres; "<"&HOY) + 1)
Sprint vigente = SI(Completada; planificado; MÁX(planificado; Sprint actual))
```
**Regla (igual que Scrum):** al cerrar la reunión, las tareas *Completadas* quedan
registradas en su sprint; las **no completadas pasan automáticamente al sprint siguiente**
(*Sprint vigente*), mostrando el **desvío en días**. El **Sprint planificado (col E del
Gantt) no se modifica** → es el histórico. `Tablero` y `Entregas` agrupan por *Sprint
vigente*.

El calendario viene pre-cargado **semanal** (cierres los martes: 16/06, 23/06, 30/06…),
adaptado a las reuniones que figuran en el Gantt; **editá esas fechas** a tu cadencia real.

**Uso de la hoja `Sprints` (3 pasos, columnas a la derecha del detalle):**
1. **Pasar al siguiente** → lista `Sí` (1 clic) para empujar la tarea al sprint siguiente,
   además del paso automático al cerrar la reunión.
2. **Motivo del desvío** → lista desplegable **editable** (validación de datos): *Dependencia
   bloqueada · Sobrecarga · Falta de información · Reestimación · Recurso no disponible ·
   Prioridad reasignada · Bloqueo técnico · Otro*.
3. **Comentario (libre)** → **texto totalmente editable y personalizable** (sin validación),
   para notas que no estén en la lista.

Las tareas se **agrupan por módulo** con un encabezado, y las movidas se resaltan en ámbar.
**Botones opcionales (.xlsm):** `PasarAlSiguiente` e `IndicarMotivo` (`macro_sprints.bas`).

## 4) Trazabilidad — hoja "Sprints"
La tabla *Seguimiento de tareas por sprint* muestra, por tarea: **ID, Tarea, Sprint
planificado → Sprint vigente, Estado, Desvío (días), Pasar al siguiente, Motivo del desvío
y Comentario**. Las tareas que se movieron se resaltan. Es a la vez el panel de acción
(pasar al siguiente, elegir motivo, comentar) y el registro del desvío.

**Bitácora persistente + usuario (opcional):** `macro_sprints.bas` (`ReasignarSprintsFisico`)
reasigna físicamente el nº de sprint en el Gantt y agrega una línea con `Now()` y
`Application.UserName` en una hoja *Log Movimientos*. Requiere `.xlsm`. *No se incrustó como
`.xlsm` para no romper la compatibilidad del `.xlsx` actual.*

## Cómo validar cada función
1. **Tarea dinámica:** insertar una fila dentro de un módulo, completar G/K/F → aparece la
   barra en el Gantt, sube el contador de *Tareas* del Tablero y aparece en Entregas.
2. **Priorización:** cambiar un `Progreso` o una `Fecha` → cambian Score, Prioridad,
   Variación y el orden de la lista. Poner avance 0 con inicio pasado → marca `BLOQUEO`.
3. **Sprints por reunión:** en la hoja *Sprints*, editar la *fecha de cierre* de una reunión
   para que ya haya pasado → sube el *Sprint actual* y las tareas no completadas de ese
   sprint muestran *Sprint vigente* = siguiente, con su *Desvío (días)*.
4. **Motivo del desvío:** elegir un motivo en la columna *Motivo del desvío* (lista
   desplegable) de la fila de la tarea movida.

---

# Backlog + Sprints por frecuencia (v4)

Rediseño según el print solicitado (backlog de OKRs + sprints por frecuencia).

- **Hoja `Backlog General`**: mis OKRs **por proyecto** con sus **subtareas**. Por subtarea:
  `Estado` (Completo / No se realizó), y si no se realizó la `Acción`
  (**Pasa al siguiente** / **Baja prioridad** / **Cancelado**) + `Comentario` (texto libre).
  Muestra Frecuencia, Sprint planificado/vigente, **Mora** y **Nuevo**.
  - *Baja prioridad* → sigue **Pendiente** pero **deja de contar días de mora**.
  - *Pasa al siguiente* → aparece en el sprint siguiente con la misma lógica.
- **Hoja `Sprints`**: tres tableros **Semanal / Quincenal / Mensual** (apertura por sprint).
  La **frecuencia es por Proyecto** (tabla editable) y cada frecuencia tiene su calendario
  editable (inicio + duración 7/14/30). El *Sprint actual* se calcula solo.
- **Proyecto nuevo `Tablero Forecast`**: agregado como **solicitud nueva de la semana**
  (marcado en Comentario + "Nuevo"), con la **misma clasificación que `Tablero BI OPEX`**
  (Semanal) y 3 subtareas: *armado de procedimiento*; *slicer funcional con % de incremento
  sobre venta según factores*; *factores de clientes con % de incremento sobre venta*.
- **Tablero**: "Avance por Sprint" pasó a **"Avance por Frecuencia"** (Semanal/Quincenal/Mensual).

Decisiones (según *mejor práctica*, confirmadas): una sola hoja de Sprints con 3 secciones;
frecuencia por Proyecto; se mantuvieron y adaptaron Tablero y Entregas Próximas.

---

# Ajustes (v5)

- **Tablero Forecast dentro de los Tableros (25%)**: el proyecto nuevo se reubicó en el
  bloque de Tableros (filas 48–51 del Gantt, dentro del merge `H22:H51` del 25%). **No
  figura aparte** y el **total general sigue siendo 100%** (no suma peso propio).
- **Backlog trae los comentarios del Gantt** (1ª hoja, columna *Comentarios*) en
  *Comentario (Gantt)*, y agrega una *Justificación / Comentario* libre. Así queda el
  **historial claro** de qué no se hizo y por qué (aunque esté justificado).
- **Frecuencias editables por validación de datos** (lista Semanal/Quincenal/Mensual) en la
  tabla *Frecuencia por Proyecto* de la hoja Sprints.
- **"Nuevo" editable y por sprint**: la columna *Sol. en Sprint* indica el sprint en que se
  solicitó; *Nuevo* muestra "Nuevo S#" solo mientras el sprint actual de su frecuencia
  coincide (ej.: nuevo en S1 deja de serlo en S2).
