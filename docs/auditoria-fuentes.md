# Auditoría inicial de fuentes de datos

Fecha de inspección: 2026-10-18. Revisión de solo lectura de los dos libros aportados. No se han guardado cambios en los ficheros fuente.

## 1. Libro de Gestión de proyectos

Fichero inspeccionado: `000-Gestión de proyectos -2026 -260702(2).xlsx`

| Hoja | Dimensión detectada | Observaciones |
|---|---:|---|
| `OTS-MES` | 9 × 14 | Sin celdas combinadas detectadas |
| `OTS EN CURSO` | 78 × 28 | 8 reglas de formato condicional |
| `Listado Pedidos` | 237 × 44 | 2 rangos combinados; fórmulas `#REF!` preexistentes |
| `PdP PRODUCCIÓN` | 3297 × 478 (A:RJ) | 100 rangos combinados, 4263 reglas de formato condicional, 12.722 fórmulas en el libro completo |
| `Vacaciones-Bajas` | 290 × 534 | 146 rangos combinados; calendario/capacidad de personal |

El libro contiene 95 nombres definidos y 2 vínculos externos. El código que procese o guarde una copia debe tratar estos elementos con cuidado y no debe eliminar vínculos/nombres de forma inadvertida.

### Hoja `PdP PRODUCCIÓN`

- El encabezado de las operaciones está en la fila 19; la OT está en la columna B.
- La fila 20 es una fila oculta usada por fórmulas de la plantilla.
- La escala temporal comienza en la zona de columnas J en adelante; el área de fechas se extiende hasta la columna RJ.
- Se observan bloques de trabajo de longitudes variables: aunque muchos bloques completos ocupan ocho filas, también existen OT con 1, 3, 7 y 16 filas. El lector no debe imponer una cantidad fija de filas.
- El lector inicial `read_production_operations` lee las operaciones a partir de la fila 29, identifica la máquina/recurso en E y las horas presupuestadas en F, y devuelve incidencias cuando una operación no puede asociarse a una OT. Las fórmulas y celdas de la escala temporal no se interpretan ni modifican.
- Hay columnas/filas ocultas y fórmulas que dependen de nombres definidos. La inserción de filas sin estrategia y pruebas específicas podría dañar referencias, fórmulas o formato.

### Hoja `Listado Pedidos`

El lector inicial `read_production_order_list` lee OT (B), pedido de compras (F), producto (G), cliente (H), progreso por fases (I:L) y estado (N). Se abre con `data_only=True`: no recalcula fórmulas. Los valores calculados que falten en el fichero no se deben inventar ni sustituir por resultados estimados.

La hoja contiene 1.011 celdas cuyo valor es `#REF!` en la fuente inspeccionada. Estos errores son preexistentes y deben reportarse como incidencia de origen, no atribuirse a PDPPlanner ni repararse automáticamente.

### Hoja `Vacaciones-Bajas`

El lector inicial `read_staff_calendar` lee la identidad del personal desde B:D y las fechas visibles de la fila 12 desde J en adelante. Las celdas no vacías del calendario se devuelven como marcas de indisponibilidad, sin inferir el motivo. La semántica de cada tipo de marca y su impacto exacto en capacidad debe confirmarse antes de alimentar el motor de planificación.

## 2. Libro de seguimiento SAP

Fichero inspeccionado: `Seguimiento_pedidos SAP(1).xlsx`

Hoja principal: `2024`, dimensión detectada 5.502 × 22. La fila 5 contiene los encabezados y los registros empiezan en la fila 6.

| Campo | Columna | Uso acordado |
|---|---|---|
| Plazo de entrega / fecha prevista | B | Fecha prevista de recepción |
| Fecha de entrega | D | Fecha real de recepción |
| Nº pedido SAP | E | Identificador del pedido |
| O.T. | H | Relación con la OT |
| Estado | J | Clasificación del pedido |
| Fecha respuesta proveedor | C | Ignorar expresamente |

En las filas 6–5502 hay 5.497 filas examinadas. En el estado actual de la fuente, 358 filas no tienen OT y 367 no tienen identificador de pedido SAP; deben revisarse/filtrarse según reglas explícitas, no asociarse a una OT por inferencia. Se detectó un número de pedido SAP repetido en dos OT diferentes (`4500518152`, OT `2300056` y `2300069`, ambas con estado `RECIBIDO`). El procesamiento debe señalar este caso como posible anomalía, no fusionar automáticamente pedidos entre OT distintas sin una regla acordada.

Estados observados con más frecuencia incluyen `RECIBIDO`, `EN TRAMITE`, `PARCIALMENTE RECIBIDO`, `ANULADO`, `RECIBIDO ERRONEO`, `LISTO PARA RECOGER` y `pendiente comprobar calidad`. Debe normalizarse texto y mantener visibles los estados no reconocidos.

## 3. Estado de implementación y protección de datos

- Se han añadido lectores iniciales de solo lectura para las cuatro fuentes: SAP `2024`, operaciones de `PdP PRODUCCIÓN`, `Listado Pedidos` y calendario `Vacaciones-Bajas`.
- Hay pruebas sintéticas para los lectores; todavía falta ejecutar una revisión de resultados sobre los libros fuente completos y clasificar las incidencias reales.
- Los lectores devuelven registros tipados y objetos `ValidationIssue`; no guardan los libros ni insertan filas.
- Los errores `#REF!` de origen no deben confundirse con errores creados por el código.
- La salida modificada del PdP se desarrollará en un hito posterior, mediante una copia nueva y pruebas de regresión.
- Antes de cualquier escritura, comparar fórmulas, nombres definidos, vínculos externos, rangos combinados, dimensiones y reglas de formato condicional.

## 4. Decisiones pendientes de confirmar antes de automatizar

- Regla exacta para identificar inicio y fin de un bloque de OT cuando la OT está repetida en varias filas o hay filas incompletas.
- Qué hacer con registros SAP sin OT o sin número de pedido.
- Qué regla aplicar si el mismo número de pedido SAP aparece asociado a varias OT.
- Semántica de las marcas de calendario de personal y su efecto en la capacidad disponible.
- Reglas completas del motor de capacidad y planificación (prioridades, precedencias, disponibilidad por recurso, duraciones y fechas bloqueadas).
- Política para fórmulas `#REF!` ya presentes: reportarlas sin intentar repararlas automáticamente.

## 5. Siguiente entrega técnica

Completar un informe consolidado de validación y ejecutar los lectores contra las fuentes reales. Después, diseñar el modelo de capacidad y el motor de planificación a partir de datos comprobados. La escritura del Gantt y el motor de planificación no deben acoplarse directamente a openpyxl.
