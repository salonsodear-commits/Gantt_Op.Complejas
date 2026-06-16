# Diagrama de Gantt — Proyecto Datos Op. Complejas (versión mejorada)

Herramienta de seguimiento de proyecto convertida en un tablero **dinámico, profesional, escalable y fácil de mantener**, respetando al 100 % la planificación original.

> **Archivo entregable:** `Diagrama_de_Gantt___Proyecto_Datos_Op._Complejas_080626.xlsx`
> **Documentación (Excel aparte):** `Guia_Diagrama_de_Gantt.xlsx`
> **Macro opcional (write-back + log):** `macro_sprints.bas`
> **Respaldo original intacto:** `original_backup.xlsx`
> **Generador reproducible:** `build_xlsx.py`

Hojas del libro: `Proyecto Datos (OKR´s)` (Gantt original), `Tablero`,
`Entregas Próximas`, `Historial de Sprints`, y `_Datos` (motor, oculta) — más las
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

## 3) Automatización de sprint (sin intervención manual)
En **Historial de Sprints** se configuran `Inicio Sprint 1` y `Duración` (editables).
```
Sprint actual = MÁX(1; ENTERO((HOY − InicioSprint1)/Duración) + 1)
Sprint vigente = SI(Completada; planificado; MÁX(planificado; Sprint actual))
```
Una tarea no completada cuyo sprint ya cerró pasa **automáticamente** al sprint actual (por
fórmula, se recalcula sola). El **Sprint planificado original (col E del Gantt) no se
modifica** → queda como histórico. El `Tablero` agrupa por *Sprint vigente*.

**Postergar manualmente por cuello de botella — hoja `Postergar`:** en la fila de cada
tarea hay dos listas desplegables:
- **Postergar (nº sprints)** → `0,1,2,3` (1 = próximo sprint).
- **Motivo del cuello de botella** → lista editable: *Dependencia bloqueada · Sobrecarga del
  responsable · Falta de información/insumos · Reestimación/mayor alcance · Recurso no
  disponible · Prioridad reasignada · Bloqueo técnico · Otro*.

`Sprint vigente = MÁX(planificado; Sprint actual) + Postergar`. Al elegir las opciones, el
sprint vigente se recalcula solo y el movimiento se registra en `Historial de Sprints` con
ese motivo. **Botón de 1 clic (opcional):** la macro `PostergarTareaActual` de
`macro_sprints.bas` rellena esas celdas automáticamente — basta insertar un *Botón de
formulario* y asignarle la macro (requiere `.xlsm`).

## 4) Trazabilidad — hoja "Historial de Sprints"
Registra cada reasignación con: **ID, Tarea, Responsable, Sprint origen, Sprint destino,
Fecha del movimiento** (cierre del sprint origen) **y Motivo**. Se completa solo con las
tareas cuyo *vigente ≠ planificado*.

**Registro persistente + usuario (opcional):** `macro_sprints.bas` reasigna físicamente el
sprint en el Gantt y agrega una línea con `Now()` y `Application.UserName` en una hoja
*Log Movimientos*. Requiere guardar como `.xlsm` (Alt+F11 → Insertar módulo → pegar →
ejecutar `ReasignarSprints`; opcional `Workbook_Open`). *No se incrustó como `.xlsm`
para no romper la compatibilidad del `.xlsx` actual.*

## Cómo validar cada función
1. **Tarea dinámica:** insertar una fila dentro de un módulo, completar G/K/F → aparece la
   barra en el Gantt, sube el contador de *Tareas* del Tablero y aparece en Entregas.
2. **Priorización:** cambiar un `Progreso` o una `Fecha` → cambian Score, Prioridad,
   Variación y el orden de la lista. Poner avance 0 con inicio pasado → marca `BLOQUEO`.
3. **Automatización de sprint:** en *Historial de Sprints*, fijar `Inicio Sprint 1` a una
   fecha tal que el Sprint 1 ya haya cerrado → las tareas no completadas pasan a *Sprint
   vigente* = actual y aparecen listadas en *Movimientos entre sprints*.
4. **Trazabilidad:** revisar la tabla de *Movimientos*; cada fila muestra origen→destino,
   fecha y motivo.
