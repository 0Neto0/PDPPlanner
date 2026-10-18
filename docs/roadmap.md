# PDPPlanner — hoja de ruta global

## Objetivo del proyecto

Construir PDPPlanner como herramienta de apoyo a la planificación de producción, no como una utilidad limitada a añadir marcas de pedidos en el Excel. El procesamiento del PdP y las fechas de pedidos es un componente de la solución completa.

## Hitos

### Hito 0 — Bootstrap y calidad técnica
- Estructura del paquete Python y configuración del proyecto.
- Configuración de Pytest, Ruff y, cuando proceda, MyPy.
- Modelo inicial de dominio.
- GitHub Actions para ejecutar comprobaciones automáticas.
- Flujo de ramas: `main`, `develop` y ramas `feature/*`.

### Hito 1 — Requisitos y arquitectura
- Documentar las reglas de planificación y las fuentes de datos.
- Separar lectura de Excel, modelo de dominio, reglas de planificación y generación de resultados.
- Mantener trazabilidad mediante issues, commits, pruebas y documentación.
- Confirmar y documentar la regla de organización de las OT en grupos de 8 líneas, sin codificar supuestos no validados.

### Hito 2 — Lectura y validación de datos
- Leer la hoja `PdP PRODUCCIÓN`.
- Leer `Listado Pedidos` y `Vacaciones-Bajas`.
- Leer el seguimiento SAP, hoja `2024`.
- Validar encabezados, OT, fechas, ausencias, duplicados y errores de datos.
- No modificar nunca los libros fuente.

### Hito 3 — Modelo de dominio y reglas de negocio
- Representar OT, pedidos, calendarios, disponibilidad/capacidad y elementos de planificación.
- Centralizar normalización y validación de datos.
- Informar de anomalías y estados desconocidos en lugar de clasificarlos silenciosamente.

### Hito 4 — Motor de capacidad y planificación
- Implementar y probar las reglas de capacidad y planificación de producción.
- Considerar las restricciones y calendarios que se hayan confirmado en los requisitos.
- Validar resultados con casos representativos y resultados esperados acordados.

### Hito 5 — Integración de pedidos en el PdP
- Añadir una línea informativa `PEDIDOS` por OT en `PdP PRODUCCIÓN`.
- Columna E = `PEDIDOS`; columna F = porcentaje de pedidos pendientes.
- Contar pedidos distintos, excluir anulados y tratar `RECIBIDO ERRÓNEO` como pendiente.
- Representar fechas previstas y reales de recepción en el Gantt.
- Conservar fórmulas, estilos, celdas combinadas y formato condicional; generar una copia de salida.

### Hito 6 — Validación integral
- Pruebas unitarias, de integración y de regresión.
- Verificar cálculos de pedidos y reglas del motor de planificación.
- Comparar el libro generado con el original: fórmulas, combinaciones, estilos, formatos y datos.
- Revisar manualmente los estados desconocidos y las anomalías reportadas.

### Hito 7 — Entrega y uso
- Documentar instalación, configuración, ejecución y resolución de errores.
- Proporcionar ejemplos reproducibles y explicar limitaciones conocidas.
- Preparar la integración mediante Pull Request; no integrar en `main` sin validación.

## Tablero Kanban

Usar estas columnas para las tareas e hitos:
1. Backlog
2. Análisis
3. En desarrollo
4. En pruebas
5. Validación
6. Finalizado

Cada cambio debe tener issue o tarea trazable, pruebas pertinentes y documentación actualizada.

## Estado inicial conocido

- Existe una primera implementación de normalización y resumen de estados de pedidos con pruebas.
- La integración de los pedidos en el Excel sigue pendiente de implementación y validación.
- La lectura completa de las fuentes, el motor de capacidad/planificación y la validación integral siguen formando parte del alcance global.
- No considerar terminado un hito hasta comprobar su resultado real.
