"""
LA POSTA — Importar productos desde lista de precios Excel
Ejecutar: python importar_productos.py
"""

import openpyxl
import json
import urllib.request
import urllib.parse

API_URL = 'https://script.google.com/macros/s/AKfycbzYU4oUPCb7jNORognxXS-gVDVk8SBqU6H7Kp24eWxMtZD9pZ5wjMGpXK35aOTjLHkb0Q/exec'
EXCEL_PATH = r'C:\Users\jgriffo\Downloads\2026-06 Lista de precios.xlsx'

def p(val):
    """Convierte a float, devuelve 0 si no es número."""
    try:
        v = float(val)
        return v if v > 0 else 0
    except (TypeError, ValueError):
        return 0

# ── Parsear Excel ──────────────────────────────────────────────────────────
wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
ws = wb.active

productos = []

def add(nombre, categoria, unidad, precio, costo=0, precio_oferta='', cant_min=''):
    if not nombre or precio == 0:
        return
    productos.append({
        'nombre': nombre,
        'categoria': categoria,
        'unidad': unidad,
        'precio': precio,
        'costo': costo,
        'activo': True,
        'precioOferta': precio_oferta,
        'cantMinOferta': cant_min,
        'fechaFinOferta': ''
    })

rows = list(ws.iter_rows(values_only=True))

# ── FIAMBRE (unidad=kg, precio=por kg extrapolado desde 100gr) ────────────
fiambres = [
    ('JAMON COCIDO',     'Fiambre'),
    ('MORTADELA',        'Fiambre'),
    ('PALETA',           'Fiambre'),
    ('QUESO DE MAQUINA', 'Fiambre'),
    ('SALAME',           'Fiambre'),
    ('CHEDDAR',          'Fiambre'),
]
for row in rows:
    nombre = str(row[0]).strip().upper() if row[0] else ''
    for prod_nombre, cat in fiambres:
        if nombre == prod_nombre:
            costo   = p(row[2])
            p100gr  = p(row[4])   # precio x 100gr
            p250gr  = p(row[5])   # precio x 250gr
            precio_kg = p100gr * 10
            oferta_kg = round(p250gr * 4) if p250gr > 0 else ''
            cant_min  = 0.250 if p250gr > 0 else ''
            add(prod_nombre.title(), cat, 'kg', precio_kg, costo * 10, oferta_kg, cant_min)

# ── PICADITA (unidad=100gr, precio directo) ───────────────────────────────
picaditas = [
    'ACEITUNAS VERDES', 'CHIZITOS', 'MANI CERVECERO', 'MANI COMUN',
    'PALITOS SALADOS', 'PAPAS FRITAS', 'LONGANIZA', 'SALAMIN'
]
for row in rows:
    nombre = str(row[0]).strip().upper() if row[0] else ''
    if nombre in picaditas:
        costo  = p(row[2])
        precio = p(row[4])
        add(nombre.title(), 'Picadita', '100gr', precio, costo / 10)

# ── QUESOS (unidad=kg) ────────────────────────────────────────────────────
quesos_map = {
    'CREMOSO PTA DE AGUA': ('Quesos', 4),
    'CREMOSO LECTAR':      ('Quesos', 4),
    'REGGIANITO':          ('Quesos', 4),
}
for row in rows:
    nombre = str(row[0]).strip().upper() if row[0] else ''
    if nombre in quesos_map:
        cat, col = quesos_map[nombre]
        costo  = p(row[2])
        precio = p(row[col])
        add(nombre.title(), cat, 'kg', precio, costo)

# ── HUEVOS ────────────────────────────────────────────────────────────────
for row in rows:
    nombre = str(row[0]).strip().upper() if row[0] else ''
    if nombre == 'HUEVOS':
        costo   = p(row[2])
        p_maple = p(row[4])
        p_2mp   = p(row[5])
        if p_maple > 0:
            oferta = round(p_2mp / 2) if p_2mp > 0 else ''
            cant   = 2 if p_2mp > 0 else ''
            add('Huevos (maple)', 'Huevos', 'un', p_maple, costo, oferta, cant)
    if nombre == '1/2 DOCENA':
        add('Huevos (1/2 docena)', 'Huevos', 'un', p(row[4]), 0)
    if nombre == 'X DOCENA':
        add('Huevos (docena)', 'Huevos', 'un', p(row[4]), 0)

# ── PASTAS ────────────────────────────────────────────────────────────────
pastas_map = {
    'RAVIOLES':    ('Pastas', 5, 3400),
    'RAVIOLONES':  ('Pastas', 5, 4100),
    'SORRENTINOS': ('Pastas', 5, 4100),
    'FIDEOS':      ('Pastas', 6, 2600),
}
for row in rows:
    nombre = str(row[0]).strip().upper() if row[0] else ''
    if nombre in pastas_map:
        cat, col, costo = pastas_map[nombre]
        precio = p(row[col])
        add(nombre.title(), cat, 'porción', precio, costo)
    if nombre == 'NOQUIS 1/2KG':
        add('Ñoquis 1/2 kg', 'Pastas', 'un', p(row[5]), 2200)

# ── SALSAS / OTROS ────────────────────────────────────────────────────────
salsas_map = {
    'SALSA LISTA':   ('Almacén', 4, 'un'),
    'PURE DE TOMATE': ('Almacén', 4, 'un'),
    'ACEITE':        ('Almacén', 4, 'un'),
    'ARROZ 1/2 KG':  ('Almacén', 4, 'un'),
    'CALDO MAGUI':   ('Almacén', 4, 'un'),
    'CARBON':        ('Almacén', 4, 'un'),
}
for row in rows:
    nombre = str(row[0]).strip().upper() if row[0] else ''
    if nombre in salsas_map:
        cat, col, unidad = salsas_map[nombre]
        costo = p(row[2])
        precio = p(row[col])
        add(nombre.title(), cat, unidad, precio, costo)

# ── TAPAS ─────────────────────────────────────────────────────────────────
tapas_map = {
    'PASCUALINA HOJALDRE': 2400,
    'PASCUALINA CRIOLLA':  0,
    'TORTA FRITA':         2050,
    'EMPANADA HOJALDRE':   2050,
    'EMPANADA CRIOLLA':    1500,
}
for row in rows:
    nombre = str(row[0]).strip().upper() if row[0] else ''
    if nombre in tapas_map:
        precio = p(row[4])
        add(nombre.title(), 'Tapas', 'un', precio, tapas_map[nombre])

# ── CONGELADOS (unidad=kg) ────────────────────────────────────────────────
congelados = [
    'MUZARELITAS', 'NUGGETS', 'MEDALLONES DE POLLO', 'PATITAS',
    'FORMITAS', 'CARITAS', 'PAPAS BASTON', 'PAPAS NOISETTE'
]
for row in rows:
    nombre = str(row[0]).strip().upper() if row[0] else ''
    if nombre in congelados:
        costo  = p(row[2])
        precio = p(row[4])
        add(nombre.title(), 'Congelados', 'kg', precio, costo)

# ── POLLO (unidad=kg, con oferta x2kg) ───────────────────────────────────
pollo_list = [
    'ALITAS', 'CARCAZA', 'MENUDOS', 'MILANESAS DE POLLO',
    'MILANESAS DE CERDO', 'PATA MUSLO', 'POLLO ENTERO',
    'POLLO TROZADO', 'PECHUGA', 'SUPREMA'
]
for row in rows:
    nombre = str(row[0]).strip().upper() if row[0] else ''
    if nombre in pollo_list:
        p1kg = p(row[4])
        p2kg = p(row[5])
        if p1kg > 0:
            oferta = round(p2kg / 2) if p2kg > 0 and str(row[5]) != '-' else ''
            cant   = 2 if oferta else ''
            add(nombre.title(), 'Pollo', 'kg', p1kg, 0, oferta, cant)

# ── CERDO / OTROS ─────────────────────────────────────────────────────────
cerdo_list = [
    'BONDIOLA', 'CARRE', 'CHORIZO COMUN', 'CHORIZO COLORADO',
    'CHORI RELLENO', 'CUERITO', 'SALCHICHA', 'HUESO',
    'HAMBURGUESA DE POLLO', 'MATAMBRE', 'MATAMBRE / MANTA',
    'PECHITO DE CERDO', 'PANCETA AHUMADA', 'PANCETA SALADA',
    'PATA DE CHANCHO', 'ROSCA'
]
for row in rows:
    nombre = str(row[0]).strip().upper() if row[0] else ''
    if nombre in cerdo_list:
        costo  = p(row[2])
        precio = p(row[4])
        add(nombre.title(), 'Cerdo', 'kg', precio, costo)

# ── Mostrar resumen ────────────────────────────────────────────────────────
print(f"\nProductos a importar: {len(productos)}")
for prod in productos:
    oferta_info = f" | Oferta: {prod['precioOferta']}/kg x{prod['cantMinOferta']}" if prod['precioOferta'] else ''
    print(f"  [{prod['categoria']}] {prod['nombre']} — {prod['unidad']} — ${prod['precio']}{oferta_info}")

# ── Enviar al backend ─────────────────────────────────────────────────────
print(f"\nEnviando {len(productos)} productos al backend...")

payload = json.dumps({
    'action': 'bulkImportProductos',
    'data': {
        'productos': productos,
        'limpiarAntes': True   # borra los productos de prueba antes de importar
    }
})

data = urllib.parse.urlencode({'payload': payload}).encode('utf-8')
req = urllib.request.Request(API_URL, data=data, method='POST')
req.add_header('Content-Type', 'application/x-www-form-urlencoded')

try:
    with urllib.request.urlopen(req, timeout=60) as response:
        result = json.loads(response.read().decode('utf-8'))
        if result.get('ok'):
            print(f"\nOK: {result['data']['importados']} productos cargados.")
        else:
            print(f"\nERROR backend: {result.get('error')}")
except Exception as e:
    print(f"\nERROR conexion: {e}")
