# PDPPlanner — tablero Kanban inicial

Este archivo es una vista de seguimiento versionada del proyecto. Las incidencias de GitHub son la fuente de detalle; actualizar este tablero cuando cambie el estado de una tarea.

## Backlog

- [ ] [#2 — Añadir seguimiento de pedidos SAP al Gantt](https://github.com/0Neto0/PDPPlanner/issues/2) — funcionalidad de pedidos, dependiente de los lectores y de las pruebas de integridad del Excel.
- [ ] [#5 — Modelo de dominio y motor de capacidad/planificación](https://github.com/0Neto0/PDPPlanner/issues/5) — alcance global de PDPPlanner.

## Análisis

- [ ] [#3 — Auditoría de fuentes y contrato de datos](https://github.com/0Neto0/PDPPlanner/issues/3) — auditoría inicial documentada; quedan por acordar las reglas de bloques de OT, datos incompletos y planificación.

## En desarrollo

- [ ] [#4 — Lectores de Excel y validación de datos](https://github.com/0Neto0/PDPPlanner/issues/4) — iniciado el lector de solo lectura para SAP 2024, con incidencias explícitas de calidad de datos.

## En pruebas

- [ ] Lector SAP: pruebas de cabeceras, filas incompletas, estados desconocidos y pedido asociado a varias OT. Pendiente comprobar el resultado de CI para estos nuevos commits.

## Validación

- [ ] Integridad de los libros Excel tras cualquier generación de salida.
- [ ] Validación de resultados del motor de capacidad y planificación contra casos esperados.

## Finalizado

- [x] Base inicial del paquete Python, configuración de Pytest/Ruff y GitHub Actions.
- [x] Primera implementación de normalización y resumen de estados SAP con pruebas; CI inicial confirmado en verde.
- [x] Hoja de ruta global y auditoría inicial de los libros fuente documentadas.

## Regla de trabajo

No mover una tarea a **Finalizado** solo porque el código esté escrito. Requiere pruebas ejecutadas y criterios de aceptación cumplidos. Los ficheros fuente nunca se sobrescriben.
