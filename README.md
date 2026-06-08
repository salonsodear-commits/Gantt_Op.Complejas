# Diagrama de Gantt — Proyecto Datos Op. Complejas (versión mejorada)

Herramienta de seguimiento de proyecto convertida en un tablero **dinámico, profesional, escalable y fácil de mantener**, respetando al 100 % la planificación original.

> **Archivo entregable:** `Diagrama_de_Gantt___Proyecto_Datos_Op._Complejas_080626.xlsx`
> **Respaldo original intacto:** `original_backup.xlsx`
> **Generador reproducible:** `build_xlsx.py`

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
- **Avance por Sprint** con barra de datos.

### C) Hoja **Entregas Próximas** (Parte 2)
- **Resumen por Sprint**: tareas, pendientes, vencidas, próximas a vencer, próxima
  entrega y % de avance, más fila TOTAL.
- **Detalle de entregas** ordenado por Sprint y luego por Fecha de entrega, con semáforo
  de estado. Columnas: Sprint · Entregable (Módulo) · Tarea · Responsable · Fecha de
  entrega · Estado · Días restantes · Progreso.
- **Umbral “Próximo a vencer”** editable (celda `I4`, por defecto 7 días).

### D) Hoja **Guía**
Documentación dentro del propio libro: criterios, fórmulas, leyenda de colores,
escalabilidad y observaciones detectadas.

### E) Hoja **_Datos** (oculta)
Motor de cálculo que lee el Gantt y prepara los datos. **No editar.**

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
python3 build_xlsx.py   # lee original_backup.xlsx y genera el archivo mejorado
```
