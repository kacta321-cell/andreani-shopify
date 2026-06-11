"""
Shopify CSV -> Andreani Excel (CLI).
Uso: python convertir_andreani.py [ruta_csv] [carpeta_salida]
  ruta_csv      : CSV exportado de Shopify (default: orders_export*.csv más reciente en esta carpeta o en Downloads)
  carpeta_salida: dónde guardar el Excel (default: misma carpeta del CSV)

La lógica de conversión vive en convertir_core.py (compartida con la web app).
"""
import sys, os, urllib.request
from pathlib import Path

SCRIPT_DIR    = Path(__file__).parent
GITHUB_RAW    = "https://raw.githubusercontent.com/kacta321-cell/andreani-shopify/main"
JSON_PATH     = SCRIPT_DIR / "andreani_localidades.json"
TEMPLATE_PATH = SCRIPT_DIR / "EnvioMasivoExcelPaquetes.xlsx"


def descargar_si_falta(path, url):
    if not path.exists():
        print(f"Descargando {path.name}...")
        urllib.request.urlretrieve(url, path)


# Auto-descarga de dependencias de datos
descargar_si_falta(JSON_PATH,     f"{GITHUB_RAW}/andreani_localidades.json")
descargar_si_falta(TEMPLATE_PATH, f"{GITHUB_RAW}/EnvioMasivoExcelPaquetes.xlsx")

# Asegurar openpyxl
try:
    import openpyxl  # noqa: F401
except ImportError:
    print("Instalando openpyxl...")
    os.system(f'"{sys.executable}" -m pip install openpyxl -q')

from convertir_core import convertir

# Detectar CSV
if len(sys.argv) > 1:
    csv_path = Path(sys.argv[1])
else:
    candidatos = (
        list(SCRIPT_DIR.glob("orders_export*.csv")) +
        list(Path.home().glob("Downloads/orders_export*.csv"))
    )
    if not candidatos:
        print("ERROR: No se encontró orders_export*.csv.")
        print("Uso: python convertir_andreani.py ruta/al/archivo.csv")
        sys.exit(1)
    csv_path = max(candidatos, key=lambda p: p.stat().st_mtime)
    print(f"Usando CSV: {csv_path}")

output_dir  = Path(sys.argv[2]) if len(sys.argv) > 2 else csv_path.parent
output_path = output_dir / "Andreani_envio.xlsx"

n = convertir(str(csv_path), str(TEMPLATE_PATH), str(JSON_PATH), str(output_path))
print(f"Pedidos procesados: {n}")
print(f"Listo: {output_path}")
