# Seguimiento de pedidos SAP en el Gantt

## Objetivo

Añadir una línea informativa `PEDIDOS` dentro de cada bloque de OT de la hoja `PdP PRODUCCIÓN`. La línea muestra el porcentaje de pedidos pendientes y fechas previstas/reales de recepción en el Gantt.

## Archivos y campos

### Libro de Gestión de proyectos

- Hoja de salida: `PdP PRODUCCIÓN`.
- OT: columna B.
- Línea nueva: columna E = `PEDIDOS`; columna F = porcentaje pendiente.
- La salida debe ser una copia nueva; nunca sobrescribir el original.

### Libro Seguimiento de pedidos SAP

- Hoja: `2024`.
- OT: columna H.
- Identificador de pedido SAP: columna E.
- Estado: columna J.
- Fecha prevista: columna B.
- Fecha real: columna D.
- Columna C: ignorada intencionadamente.

## Cálculo

- Contar pedidos distintos asociados a cada OT, no unidades.
- Excluir los pedidos anulados del denominador.
- Pendientes: estados normalizados `EN TRÁMITE` y `PARCIALMENTE RECIBIDO`.
- Recibidos: estado normalizado `RECIBIDO`.
- Normalizar espacios, mayúsculas y tildes/variantes de escritura.
- Los estados desconocidos deben informarse para revisión; no clasificarlos silenciosamente.

\[
\text{porcentaje pendiente} = \frac{\text{pedidos pendientes}}{\text{pedidos no anulados}} \times 100
\]

Si una OT no tiene pedidos válidos, el porcentaje debe mostrarse como `N/D` (no como 0 %), para diferenciar ausencia de datos de una OT completamente recibida.

## Marcas del Gantt

Para cada OT, representar las fechas previstas (columna B de SAP) y reales (columna D) de sus pedidos. Como una OT puede tener varios pedidos, puede tener varias marcas de cada tipo. La fecha real vacía no debe generar marca real.

## Seguridad de los datos del Excel

La hoja `PdP PRODUCCIÓN` tiene muchas columnas de fechas, celdas combinadas y fórmulas. Insertar filas con una librería Excel puede dejar referencias, combinaciones o formato condicional sin actualizar. Antes de generar el archivo final, se debe probar el método de inserción y verificar fórmulas, estilos, celdas combinadas y formato del Gantt. Los libros de origen se mantienen intactos.

## Validación

- OT con varios pedidos en estados distintos.
- OT sin pedidos.
- Pedidos anulados.
- Estados con espacios o diferencias de mayúsculas.
- Fechas previstas y reales vacías.
- Pedidos con fechas distintas para una misma OT.
- Revisión de fórmulas, estilos, combinaciones y formato condicional tras generar la copia.
