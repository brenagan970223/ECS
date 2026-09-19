import sys

try:
    import powerfactory as pf
except ImportError:
    sys.path.append(r"C:\Program Files\DIgSILENT\PowerFactory 2024\Python\3.12")
    import powerfactory as pf

app = pf.GetApplication()

if app is None:
    raise RuntimeError(
        "CONEXION FALLIDA: GetApplication() devolvio None. "
        "Si esto corre fuera de PowerFactory, cierra cualquier instancia grafica abierta."
    )

project = app.GetActiveProject()

app.PrintInfo("===================================")
app.PrintInfo("CONEXION EXITOSA A POWERFACTORY")
app.PrintInfo(f"Proyecto activo: {project}")
app.PrintInfo("===================================")
