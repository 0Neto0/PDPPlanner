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
- Hay 100 celdas combinadas y 4263 reglas de formato condicional en la hoja.
- Se observan bloques de trabajo con una fila de cabecera/operación y siete filas de secciones productivas (ocho filas en total para muchos bloques completos). Sin embargo, la fuente también contiene OT con 1, 3, 7 y 16 filas. Por tanto, **no se debe asumir que cada OT ocupa exactamente ocho filas**: el lector debe identificar el bloque real y reportar excepciones.
- Hay columnas/filas ocultas y fórmulas que dependen de nombres definidos. La inserción de filas sin estrategia y pruebas específicas podría dañar referencias, fórmulas o formato.

### Errores preexistentes detectados

La hoja `Listado Pedidos` contiene 1.011 celdas cuyo valor es `#REF!` (incluye un bloque de columnas AD:AJ y más referencias en filas posteriores). Estos errores ya están presentes en el libro fuente inspeccionado. Deben registrarse como incidencia de origen y no atribuirse a PDPPlanner.

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

## 3. Reglas de protección de los datos

1. Abrir los libros fuente en modo de lectura lógica y no guardarlos encima de los originales.
2. La primera versión de los lectores debe producir objetos de dominio e informe de validación; no debe editar celdas.
3. Separar los errores preexistentes del Excel de los errores detectados por el lector.
4. La salida modificada del PdP se desarrollará en un hito posterior, mediante una copia nueva y pruebas de regresión.
5. Antes de cualquier escritura, comparar fórmulas, nombres definidos, vínculos externos, rangos combinados, dimensiones y reglas de formato condicional.

## 4. Decisiones pendientes de confirmar antes de automatizar

- Regla exacta para identificar inicio y fin de un bloque de OT cuando la OT está repetida en varias filas o hay filas incompletas.
- Qué hacer con registros SAP sin OT o sin número de pedido.
- Qué regla aplicar si el mismo número de pedido SAP aparece asociado a varias OT.
- Reglas completas del motor de capacidad y planificación (prioridades, precedencias, disponibilidad por recurso, duraciones y fechas bloqueadas).
- Política para fórmulas `#REF!` ya presentes: reportarlas sin intentar repararlas automáticamente.

## 5. Siguiente entrega técnica

Implementar lectores independientes y de solo lectura para:
- `PdP PRODUCCIÓN` y su estructura de OT;
- `Listado Pedidos`;
- `Vacaciones-Bajas` y disponibilidad por recurso/fecha;
- seguimiento SAP `2024`.

Los lectores devolverán datos tipados y advertencias de validación. La escritura del Gantt y el motor de planificación no deben acoplarse a openpyxl directamente.
