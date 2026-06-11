"""
Shopify CSV -> Andreani Excel
Uso: python convertir_andreani.py [ruta_csv] [carpeta_salida]
  ruta_csv      : archivo CSV exportado desde Shopify (default: busca orders_export*.csv en carpeta actual)
  carpeta_salida: donde guardar el Excel (default: misma carpeta que el CSV)
"""
import csv, shutil, json, unicodedata, sys, os, glob, urllib.request
from pathlib import Path

# ── CONFIG (modificar si cambian las medidas del paquete) ─────────────────────
PESO          = 500
ALTO          = 15
ANCHO         = 20
PROF          = 10
TEL_COD       = '11'
TEL_NUM       = '00000000'
EMAIL_DEFAULT = 'sinmail@sinmail.com'

# ── RUTAS ────────────────────────────────────────────────────────────────────
SCRIPT_DIR   = Path(__file__).parent
GITHUB_RAW   = "https://raw.githubusercontent.com/kacta321-cell/andreani-shopify/main"
JSON_PATH    = SCRIPT_DIR / "andreani_localidades.json"
TEMPLATE_PATH = SCRIPT_DIR / "EnvioMasivoExcelPaquetes.xlsx"

def descargar_si_falta(path, url):
    if not path.exists():
        print(f"Descargando {path.name}...")
        urllib.request.urlretrieve(url, path)
        print(f"  OK: {path.name}")

# ── AUTO-DESCARGA DE DEPENDENCIAS ─────────────────────────────────────────────
descargar_si_falta(JSON_PATH,     f"{GITHUB_RAW}/andreani_localidades.json")
descargar_si_falta(TEMPLATE_PATH, f"{GITHUB_RAW}/EnvioMasivoExcelPaquetes.xlsx")

# Verificar openpyxl
try:
    import openpyxl
    from openpyxl.styles import Alignment
except ImportError:
    print("Instalando openpyxl...")
    os.system(f"{sys.executable} -m pip install openpyxl -q")
    import openpyxl
    from openpyxl.styles import Alignment

# ── DETECTAR CSV ──────────────────────────────────────────────────────────────
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

# ── LOOKUP ANDREANI ───────────────────────────────────────────────────────────
with open(JSON_PATH, encoding='utf-8') as f:
    lookup_data = json.load(f)
cp_lookup = lookup_data['por_cp']
todas     = lookup_data['lista']

def norm(s):
    s = s.upper().strip()
    s = unicodedata.normalize('NFD', s)
    return ''.join(c for c in s if unicodedata.category(c) != 'Mn')

ciudad_lookup = {}
for loc in todas:
    parts = loc.split('/')
    if len(parts) == 3:
        k = norm(parts[1].strip())
        if k not in ciudad_lookup:
            ciudad_lookup[k] = loc

def clean_cp(z):
    z = str(z).replace("'", "").strip()
    return z, ''.join(filter(str.isdigit, z))[:4]

def buscar_localidad(province, city, zipcode):
    cp_raw, cp4 = clean_cp(zipcode)
    if cp_raw in cp_lookup:  return cp_lookup[cp_raw]
    if cp4   in cp_lookup:  return cp_lookup[cp4]
    city_clean = norm(''.join(c for c in city if not c.isdigit()).strip())
    if city_clean in ciudad_lookup: return ciudad_lookup[city_clean]
    return f"{norm(province)} / {norm(city)} / {cp_raw}"

# ── HELPERS ───────────────────────────────────────────────────────────────────
def split_name(full):
    parts = full.strip().split(' ', 1)
    return (parts[0], parts[1]) if len(parts) > 1 else (parts[0], '')

def split_phone(phone):
    phone = phone.strip()
    if not phone: return TEL_COD, TEL_NUM
    for prefix in ['+54', '54']:
        if phone.startswith(prefix):
            phone = phone[len(prefix):]
            break
    if phone.startswith('9'): phone = phone[1:]
    if phone.startswith('11') and len(phone) >= 10: return '011', phone[2:]
    elif len(phone) >= 10: return phone[:3], phone[3:]
    return TEL_COD, phone or TEL_NUM

def clean_total(val):
    s = str(val).replace("'", "").replace('$', '').replace(' ', '').strip()
    if ',' in s and '.' in s:
        s = s.replace('.', '').replace(',', '.')
    elif ',' in s:
        s = s.replace(',', '.')
    try:    return round(float(s), 2)
    except: return 0.0

def clean_numero(addr1, addr2):
    if addr2:
        digits = ''.join(filter(str.isdigit, str(addr2).strip()))
        if digits: return digits[:20]
    if ',' in str(addr1):
        after  = addr1.split(',', 1)[1].strip()
        digits = ''.join(filter(str.isdigit, after))
        if digits: return digits[:20]
    return '0'

def clean_email(email):
    email = email.strip()
    return email if email and '@' in email else EMAIL_DEFAULT

# ── LEER CSV ──────────────────────────────────────────────────────────────────
with open(csv_path, encoding='utf-8-sig') as f:
    rows = list(csv.DictReader(f))

orders = {}
for r in rows:
    if r['Name'] not in orders:
        orders[r['Name']] = r
orders_list = list(orders.values())
print(f"Pedidos encontrados: {len(orders_list)}")

# ── GENERAR XLSX ──────────────────────────────────────────────────────────────
shutil.copy(TEMPLATE_PATH, output_path)
wb = openpyxl.load_workbook(output_path)
ws = wb['A domicilio']

for i, order in enumerate(orders_list):
    row = 3 + i
    nombre, apellido = split_name(order.get('Shipping Name', ''))
    cod_area, numero = split_phone(order.get('Phone', ''))
    localidad  = buscar_localidad(
        order.get('Shipping Province Name', ''),
        order.get('Shipping City', ''),
        order.get('Shipping Zip', '')
    )
    total     = int(clean_total(order.get('Total', 0)))
    nro_calle = clean_numero(order.get('Shipping Address1', ''), order.get('Shipping Address2', ''))
    email     = clean_email(order.get('Email', ''))

    data = [
        '',           # A — Paquete Guardado
        PESO,         # B — Peso
        ALTO,         # C — Alto
        ANCHO,        # D — Ancho
        PROF,         # E — Profundidad
        total,        # F — Valor declarado
        order.get('Name', ''),              # G — Numero Interno
        nombre,                             # H — Nombre
        apellido,                           # I — Apellido
        order.get('Billing Company', ''),   # J — DNI (Shopify lo guarda en Billing Company)
        email,                              # K — Email
        cod_area,                           # L — Celular código
        numero,                             # M — Celular número
        order.get('Shipping Address1', ''), # N — Calle
        nro_calle,                          # O — Número (solo dígitos)
        '',                                 # P — Piso
        '',                                 # Q — Departamento
        localidad,                          # R — Provincia / Localidad / CP
        order.get('Notes', ''),             # S — Observaciones
    ]

    for col, val in enumerate(data, 1):
        cell = ws.cell(row=row, column=col, value=val)
        cell.alignment = Alignment(vertical='center')
        if col == 6:
            cell.number_format = '0'

wb.save(output_path)
print(f"Listo: {output_path}")
