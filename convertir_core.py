"""
Lógica de conversión Shopify -> Andreani.
Fuente única usada tanto por el CLI (convertir_andreani.py) como por la web app (index.html via Pyodide).
No tiene side-effects al importar: solo define funciones.
"""
import csv, shutil, json, unicodedata

# ── CONFIG (medidas del paquete) ──────────────────────────────────────────────
PESO          = 500
ALTO          = 15
ANCHO         = 20
PROF          = 10
TEL_COD       = '11'
TEL_NUM       = '00000000'
EMAIL_DEFAULT = 'sinmail@sinmail.com'


def norm(s):
    s = (s or '').upper().strip()
    s = unicodedata.normalize('NFD', s)
    return ''.join(c for c in s if unicodedata.category(c) != 'Mn')


def clean_cp(z):
    z = str(z).replace("'", "").strip()
    return z, ''.join(filter(str.isdigit, z))[:4]


def split_name(full):
    parts = (full or '').strip().split(' ', 1)
    return (parts[0], parts[1]) if len(parts) > 1 else (parts[0], '')


def split_phone(phone):
    phone = (phone or '').strip()
    if not phone:
        return TEL_COD, TEL_NUM
    for prefix in ['+54', '54']:
        if phone.startswith(prefix):
            phone = phone[len(prefix):]
            break
    if phone.startswith('9'):
        phone = phone[1:]
    phone = ''.join(filter(str.isdigit, phone))
    if phone.startswith('11') and len(phone) >= 10:
        return '011', phone[2:]
    elif len(phone) >= 10:
        return phone[:3], phone[3:]
    return TEL_COD, phone or TEL_NUM


def pick_phone(order):
    """Toma el primer teléfono no vacío: Phone -> Shipping Phone -> Billing Phone."""
    for key in ('Phone', 'Shipping Phone', 'Billing Phone'):
        val = (order.get(key, '') or '').strip()
        if val:
            return val
    return ''


def clean_total(val):
    s = str(val).replace("'", "").replace('$', '').replace(' ', '').strip()
    if ',' in s and '.' in s:
        s = s.replace('.', '').replace(',', '.')
    elif ',' in s:
        s = s.replace(',', '.')
    try:
        return round(float(s), 2)
    except Exception:
        return 0.0


def clean_numero(addr1, addr2):
    if addr2:
        digits = ''.join(filter(str.isdigit, str(addr2).strip()))
        if digits:
            return digits[:20]
    if ',' in str(addr1):
        after  = addr1.split(',', 1)[1].strip()
        digits = ''.join(filter(str.isdigit, after))
        if digits:
            return digits[:20]
    return '0'


def clean_email(email):
    email = (email or '').strip()
    return email if email and '@' in email else EMAIL_DEFAULT


def _build_lookups(lookup_data):
    cp_lookup = lookup_data['por_cp']
    todas     = lookup_data['lista']
    ciudad_lookup = {}
    for loc in todas:
        parts = loc.split('/')
        if len(parts) == 3:
            k = norm(parts[1].strip())
            if k not in ciudad_lookup:
                ciudad_lookup[k] = loc
    return cp_lookup, ciudad_lookup


def _buscar_localidad(cp_lookup, ciudad_lookup, province, city, zipcode):
    cp_raw, cp4 = clean_cp(zipcode)
    if cp_raw in cp_lookup:
        return cp_lookup[cp_raw]
    if cp4 in cp_lookup:
        return cp_lookup[cp4]
    city_clean = norm(''.join(c for c in city if not c.isdigit()).strip())
    if city_clean in ciudad_lookup:
        return ciudad_lookup[city_clean]
    return f"{norm(province)} / {norm(city)} / {cp_raw}"


def convertir(csv_path, template_path, json_path, output_path):
    """Convierte el CSV de Shopify al Excel de Andreani. Devuelve la cantidad de pedidos."""
    import openpyxl
    from openpyxl.styles import Alignment

    with open(json_path, encoding='utf-8') as f:
        lookup_data = json.load(f)
    cp_lookup, ciudad_lookup = _build_lookups(lookup_data)

    with open(csv_path, encoding='utf-8-sig') as f:
        rows = list(csv.DictReader(f))

    orders = {}
    for r in rows:
        if r['Name'] not in orders:
            orders[r['Name']] = r
    orders_list = list(orders.values())

    shutil.copy(template_path, output_path)
    wb = openpyxl.load_workbook(output_path)
    ws = wb['A domicilio']

    for i, order in enumerate(orders_list):
        row = 3 + i
        nombre, apellido = split_name(order.get('Shipping Name', ''))
        cod_area, numero = split_phone(pick_phone(order))
        localidad  = _buscar_localidad(
            cp_lookup, ciudad_lookup,
            order.get('Shipping Province Name', '') or order.get('Shipping Province', ''),
            order.get('Shipping City', ''),
            order.get('Shipping Zip', '')
        )
        total     = int(clean_total(order.get('Total', 0)))
        nro_calle = clean_numero(order.get('Shipping Address1', ''), order.get('Shipping Address2', ''))
        email     = clean_email(order.get('Email', ''))

        data = [
            '',                                 # A — Paquete Guardado
            PESO, ALTO, ANCHO, PROF,            # B-E — medidas
            total,                              # F — Valor declarado
            order.get('Name', ''),              # G — Numero Interno
            nombre,                             # H — Nombre
            apellido,                           # I — Apellido
            order.get('Billing Company', ''),   # J — DNI
            email,                              # K — Email
            cod_area,                           # L — Celular código
            numero,                             # M — Celular número
            order.get('Shipping Address1', ''), # N — Calle
            nro_calle,                          # O — Número
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
    return len(orders_list)
