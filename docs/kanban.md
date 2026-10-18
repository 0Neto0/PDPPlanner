# PDPPlanner — tablero Kanban inicial

Este archivo es una vista de seguimiento versionada del proyecto. Las incidencias de GitHub son la fuente de detalle; actualizar este tablero cuando cambie el estado de una tarea.

## Backlog

- [ ] [#2 — Añadir seguimiento de pedidos SAP al Gantt](https://github.com/0Neto0/PDPPlanner/issues/2) — funcionalidad de pedidos, dependiente de la validación de lectores y de las pruebas de integridad del Excel.
- [ ] [#5 — Modelo de dominio y motor de capacidad/planificación](https://github.com/0Neto0/PDPPlanner/issues/5) — alcance global de PDPPlanner.

## Análisis

- [ ] [#3 — Auditoría de fuentes y contrato de datos](https://github.com/0Neto0/PDPPlanner/issues/3) — auditoría inicial documentada; quedan por cerrar reglas de datos incompletos y planificación.

## En desarrollo

- [ ] [#4 — Lectores de Excel y validación de datos](https://github.com/0Neto0/PDPPlanner/issues/4) — lectores iniciales para SAP `2024`, operaciones de `PdP PRODUCCIÓN`, `Listado Pedidos` y calendario de `Vacaciones-Bajas`. Todos trabajan en solo lectura y devuelven incidencias de validación.

## En pruebas

- [ ] SAP: cabeceras, filas incompletas, estados desconocidos y pedido asociado a varias OT.
- [ ] Producción: longitudes variables de bloques de OT, celdas OT vacías por combinación y operaciones sin OT.
- [ ] Listado de pedidos: correspondencia de columnas OT/producto/cliente/estado y progreso.
- [ ] Calendario de personal: lectura de fechas y marcas de indisponibilidad.
- [ ] Confirmar CI de la última revisión antes de mover lectores a Validación.

## Validación

- [ ] Ejecutar lectores contra los dos libros fuente reales y revisar el resumen de incidencias.
- [ ] Integridad de los libros Excel tras cualquier generación de salida.
- [ ] Validación de resultados del motor de capacidad y planificación contra casos esperados.

## Finalizado

- [x] Base inicial del paquete Python, configuración de Pytest/Ruff y GitHub Actions.
- [x] Primera implementación de normalización y resumen de estados SAP con pruebas; CI inicial confirmado en verde.
- [x] Hoja de ruta global y auditoría inicial de los libros fuente documentadas.

## Regla de trabajo

No mover una tarea a **Finalizado** solo porque el código esté escrito. Requiere pruebas ejecutadas y criterios de aceptación cumplidos. Los ficheros fuente nunca se sobrescriben.
