# PDPPlanner — tablero Kanban inicial

Este archivo es una vista de seguimiento versionada del proyecto. Las incidencias de GitHub son la fuente de detalle; actualizar este tablero cuando cambie el estado de una tarea.

## Backlog

- [ ] [#2 — Añadir seguimiento de pedidos SAP al Gantt](https://github.com/0Neto0/PDPPlanner/issues/2) — funcionalidad de pedidos, dependiente de los lectores y de las pruebas de integridad del Excel.
- [ ] [#5 — Modelo de dominio y motor de capacidad/planificación](https://github.com/0Neto0/PDPPlanner/issues/5) — alcance global de PDPPlanner.

## Análisis

- [ ] [#3 — Auditoría de fuentes y contrato de datos](https://github.com/0Neto0/PDPPlanner/issues/3) — auditoría inicial documentada; quedan por acordar reglas de datos incompletos y planificación.

## En desarrollo

- [ ] [#4 — Lectores de Excel y validación de datos](https://github.com/0Neto0/PDPPlanner/issues/4) — lector SAP de solo lectura y primer lector de operaciones de `PdP PRODUCCIÓN` implementados. El lector de producción no fija bloques de 8 filas: reconoce operaciones por la columna E y conserva el contexto de OT cuando la celda está combinada.

## En pruebas

- [ ] Lector SAP: pruebas de cabeceras, filas incompletas, estados desconocidos y pedido asociado a varias OT.
- [ ] Lector de producción: pruebas sintéticas de bloques de OT con longitudes distintas, celdas OT vacías por combinación y operaciones sin OT.
- [ ] Confirmar CI de la última revisión antes de mover lectores a Validación.

## Validación

- [ ] Integridad de los libros Excel tras cualquier generación de salida.
- [ ] Validación de resultados del motor de capacidad y planificación contra casos esperados.
- [ ] Comparar lectores contra los dos libros fuente reales y documentar discrepancias sin modificar los originales.

## Finalizado

- [x] Base inicial del paquete Python, configuración de Pytest/Ruff y GitHub Actions.
- [x] Primera implementación de normalización y resumen de estados SAP con pruebas; CI inicial confirmado en verde.
- [x] Hoja de ruta global y auditoría inicial de los libros fuente documentadas.

## Regla de trabajo

No mover una tarea a **Finalizado** solo porque el código esté escrito. Requiere pruebas ejecutadas y criterios de aceptación cumplidos. Los ficheros fuente nunca se sobrescriben.
