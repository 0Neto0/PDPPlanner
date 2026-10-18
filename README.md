# PDPPlanner

PDPPlanner es un proyecto para desarrollar una herramienta de apoyo a la planificación global de producción, incluyendo lectura y validación de los libros de gestión, modelado de OT y operaciones, capacidad por recurso, reglas de planificación y seguimiento de pedidos SAP.

La funcionalidad de pedidos en el Gantt es un componente, no el alcance completo del proyecto.

## Estado

El repositorio contiene una primera base para normalizar estados SAP y calcular resúmenes por OT, además de pruebas y GitHub Actions. Los lectores completos, el modelo de dominio, el motor de capacidad/planificación y la generación validada de libros siguen pendientes.

Consulta:
- [Hoja de ruta](docs/roadmap.md)
- [Auditoría inicial de fuentes](docs/auditoria-fuentes.md)
- [Reglas de pedidos SAP en el Gantt](docs/seguimiento-pedidos-gantt.md)

## Entorno de desarrollo

Requiere Python 3.11 o posterior.

```bash
python -m venv .venv
# Windows:
.venv\\Scripts\\activate
# Linux/macOS:
source .venv/bin/activate
python -m pip install -e ".[dev]"
pytest
ruff check .
```

## Principios del proyecto

- No sobrescribir los Excel fuente.
- Separar lectura, dominio, planificación y generación de salida.
- Validar datos incompletos y anomalías explícitamente.
- Mantener pruebas, documentación y trazabilidad en cada cambio.
- Trabajar mediante `feature/*` → `develop` → `main`; no integrar en `main` sin validación.
