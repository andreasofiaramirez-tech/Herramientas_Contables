# mappings.py
# ==============================================================================
# ARCHIVO DE CONFIGURACIÓN Y DATOS MAESTROS
# ==============================================================================
import re

def normalize_account(acc):
    """Normaliza cuentas eliminando caracteres no numéricos."""
    return re.sub(r'\D', '', str(acc))

# --- CONSTANTES DE TOLERANCIA ---
TOLERANCIA_MAX_BS = 0.02
TOLERANCIA_MAX_USD = 0.50

# --- LISTAS DE SELECCIÓN PARA APP.PY ---
# Unificamos Sillaca dentro de Febeca Quincalla
LISTA_EMPRESAS = [
    "MAYOR BEVAL, C.A", 
    "FEBECA, C.A", 
    "FEBECA, C.A (QUINCALLA)", 
    "PRISMA, C.A"
]

CODIGOS_EMPRESAS = {
    "FEBECA, C.A": "004",
    "MAYOR BEVAL, C.A": "207",
    "PRISMA, C.A": "298",
    "FEBECA, C.A (QUINCALLA)": "071"
}

# --- DIRECTORIOS DE CUENTAS (PAQUETE CC) ---
CUENTAS_CONOCIDAS = {normalize_account(acc) for acc in [
    '1.1.3.01.1.001', '1.1.3.01.1.901', '7.1.3.45.1.997', '6.1.1.12.1.001',
    '4.1.1.22.4.001', '2.1.3.04.1.001', '7.1.3.19.1.012', '2.1.2.05.1.108',
    '6.1.1.19.1.001', '4.1.1.21.4.001', '2.1.3.04.1.006', '2.1.3.01.1.012',
    '7.1.3.04.1.004', '7.1.3.06.1.998', '1.1.1.04.6.003', '1.1.4.01.7.020',
    '1.1.4.01.7.021', '1.1.1.02.1.004', '1.1.1.02.1.007', '1.1.1.02.1.009',
    '1.1.1.02.1.016', '1.1.1.02.1.112', '1.1.1.02.1.124', '1.1.1.02.1.132',
    '1.1.1.02.6.002', '1.1.1.02.6.003', '1.1.1.02.6.005', '1.1.1.02.6.010',
    '1.1.1.03.6.012', '1.1.1.03.6.024', '1.1.1.03.6.026', '1.1.1.03.6.031',
    # Bancos adicionales
    '1.1.1.02.1.002', '1.1.1.02.1.005', '1.1.1.02.6.001', '1.1.1.02.1.003',
    '4.1.1.21.4.001', '2.1.3.04.1.001', '4.1.1.22.4.001',
    # Cuentas Grupos Nuevos
    '1.9.1.01.3.008', '1.9.1.01.3.009', '7.1.3.01.1.001',
    '1.1.4.01.7.044', '2.1.2.05.1.005'
]}

CUENTAS_BANCO = {normalize_account(acc) for acc in [
    '1.1.4.01.7.020', '1.1.4.01.7.021', '1.1.1.02.1.004', '1.1.1.02.1.007',
    '1.1.1.02.1.009', '1.1.1.02.1.016', '1.1.1.02.1.112', '1.1.1.02.1.124',
    '1.1.1.02.1.132', '1.1.1.02.6.002', '1.1.1.02.6.003', '1.1.1.02.6.005',
    '1.1.1.02.6.010', '1.1.1.03.6.012', '1.1.1.03.6.024', '1.1.1.03.6.026',
    '1.1.1.03.6.031',
    '1.1.1.02.1.002', '1.1.1.02.1.005', '1.1.1.02.6.001', '1.1.1.02.1.003'
]}

# --- NOMBRES OFICIALES DE CUENTAS (CUADRE CB-CG) ---
NOMBRES_CUENTAS_OFICIALES = {
    '1.1.1.02.1.000': 'Bancos del País',
    '1.1.1.02.1.002': 'Banco Venezolano de Credito, S.A.',
    '1.1.1.02.1.003': 'Banco de Venezuela, S.A. Banco Universal',
    '1.1.1.02.1.004': 'Banco Provincial, S.A. Banco Universal',
    '1.1.1.02.1.005': 'Banesco, C.A. Banco Universal',
    '1.1.1.02.1.006': 'Banco Provincial S.A. Banco Universal',
    '1.1.1.02.1.008': 'Banco Bicentenario, Banco Universal',
    '1.1.1.02.1.009': 'Banco Mercantil C.A. Banco Universal',
    '1.1.1.02.1.010': 'Banco del Caribe C.A. Banco Universal',
    '1.1.1.02.1.015': 'Banco Exterior S.A. Banco Universal',
    '1.1.1.02.1.018': 'Banco Nacional de Cdto.C.A. Bco.Univer.',
    '1.1.1.02.1.019': 'Banco Caroní, C.A. Banco Universal',
    '1.1.1.02.1.021': 'Bancamiga Banco Universal, C.A.',
    '1.1.1.02.1.022': 'Banco Sofitasa Banco Universal, C.A.',
    '1.1.1.02.1.111': 'Banco Exterior C.A. Banco Universal',
    '1.1.1.02.1.112': 'Venezolano de Credito S.A. Bco.Universal',
    '1.1.1.02.1.115': 'Bicentenario Banco Universal, C.A.',
    '1.1.1.02.1.116': 'Venezolano de Credito S.A. Bco.Universal',
    '1.1.1.02.1.124': 'Banco Nacional de Cdto.C.A. Bco.Univer.',
    '1.1.1.02.1.126': 'Del Sur Banco Universal, C.A.',
    '1.1.1.02.1.132': 'Banplus, C.A Banco Universal',
    '1.1.1.02.6.000': 'Bancos del Pais en Monedas Extranjeras',
    '1.1.1.02.6.001': 'Banco Mercantil Banco Universal',
    '1.1.1.02.6.002': 'Banco Nacional de Crédito C.A.',
    '1.1.1.02.6.003': 'Banco de Venezuela S.A. Banco Universal',
    '1.1.1.02.6.005': 'Banesco, C.A. Banco Universal',
    '1.1.1.02.6.006': 'Bancaribe',
    '1.1.1.02.6.010': 'Banplus (US$)',
    '1.1.1.02.6.011': 'Bancamiga Banco Universal',
    '1.1.1.02.6.013': 'Banco Provincial, s.a.Banco Universal',
    '1.1.1.02.6.015': 'Banco Sofitasa, Banco Universal, C.A.',
    '1.1.1.02.6.017': 'Banco Banesco (US$) - Cta. Custodia',
    '1.1.1.02.6.210': 'Banplus (EUR)',
    '1.1.1.02.6.213': 'Banco de Venezuela (EUR)',
    '1.1.1.02.6.214': 'Banco Sofitasa, Banco Universal C.A(COP)',
    '1.1.1.02.6.995': 'Bancos del País - Dif.en Cambio No Reali',
    '1.1.1.03.0.000': 'Bancos del Exterior',
    '1.1.1.03.6.000': 'Bancos del Exterior',
    '1.1.1.03.6.002': 'Amerant Bank, N.A.',
    '1.1.1.03.6.012': 'Banesco, S.A. Panamá',
    '1.1.1.03.6.015': 'Santander Private Banking',
    '1.1.1.03.6.024': 'Venezolano de Crédito Cayman Branch',
    '1.1.1.03.6.026': 'Banco Mercantil Panamá',
    '1.1.1.03.6.028': 'FACEBANK International',
    '1.1.1.06.6.000': 'Monedero Electrónico - Moneda Extranjera',
    '1.1.1.06.6.001': 'PayPal',
    '1.1.1.06.6.002': 'Creska',
    '1.1.1.06.6.003': 'Zinli',
    '1.1.4.01.7.020': 'Servicios de Administración de Fondos -Z',
    '1.1.4.01.7.021': 'Servicios de Administración de Fondos - USDT',
    '1.1.1.01.6.001': 'Cuenta Dolares',
    '1.1.1.01.6.002': 'Cuenta Euros'
}

# --- MAPEOS DE CUADRE (CB vs CG) ---
MAPEO_CB_CG_BEVAL = {
    "0102E":  {"cta": "1.1.1.02.6.003", "moneda": "USD"},
    "0102EU": {"cta": "1.1.1.02.6.213", "moneda": "EUR"},
    "0102L":  {"cta": "1.1.1.02.1.003", "moneda": "VES"},
    "0104L":  {"cta": "1.1.1.02.1.002", "moneda": "VES"},
    "0105E":  {"cta": "1.1.1.02.6.001", "moneda": "USD"},
    "0105L":  {"cta": "1.1.1.02.1.009", "moneda": "VES"},
    "0108E":  {"cta": "1.1.1.02.6.013", "moneda": "USD"},
    "0108L":  {"cta": "1.1.1.02.1.004", "moneda": "VES"},
    "0114E":  {"cta": "1.1.1.02.6.006", "moneda": "USD"},
    "0114L":  {"cta": "1.1.1.02.1.010", "moneda": "VES"},
    "0115L":  {"cta": "1.1.1.02.1.015", "moneda": "VES"},
    "0134E":  {"cta": "1.1.1.02.6.017", "moneda": "USD"},
    "0134EC": {"cta": "1.1.1.02.6.005", "moneda": "USD"},
    "0134L":  {"cta": "1.1.1.02.1.005", "moneda": "VES"},
    "0137CP": {"cta": "1.1.1.02.6.214", "moneda": "COP"},
    "0137E":  {"cta": "1.1.1.02.6.015", "moneda": "USD"},
    "0137L":  {"cta": "1.1.1.02.1.022", "moneda": "VES"},
    "0172E":  {"cta": "1.1.1.02.6.011", "moneda": "USD"},
    "0172L":  {"cta": "1.1.1.02.1.021", "moneda": "VES"},
    "0174E":  {"cta": "1.1.1.02.6.010", "moneda": "USD"},
    "0174EU": {"cta": "1.1.1.02.6.210", "moneda": "EUR"},
    "0174L":  {"cta": "1.1.1.02.1.132", "moneda": "VES"},
    "0175L":  {"cta": "1.1.1.02.1.008", "moneda": "VES"},
    "0191E":  {"cta": "1.1.1.02.6.002", "moneda": "USD"},
    "0191L":  {"cta": "1.1.1.02.1.018", "moneda": "VES"},
    "0201E":  {"cta": "1.1.1.03.6.012", "moneda": "USD"},
    "0202E":  {"cta": "1.1.1.03.6.002", "moneda": "USD"},
    "0203E":  {"cta": "1.1.4.01.7.020", "moneda": "USD"},
    "0204E":  {"cta": "1.1.1.03.6.028", "moneda": "USD"},
    "0205E":  {"cta": "1.1.1.03.6.026", "moneda": "USD"},
    "0206E":  {"cta": "1.1.1.06.6.001", "moneda": "USD"},
    "0207E":  {"cta": "1.1.4.01.7.021", "moneda": "USD"},
    "0209E":  {"cta": "1.1.1.01.6.001", "moneda": "USD"},
    "0210EU": {"cta": "1.1.1.01.6.002", "moneda": "EUR"},
    "0211E":  {"cta": "1.1.1.03.6.015", "moneda": "USD"},
    "0501E":  {"cta": "1.1.1.03.6.024", "moneda": "USD"},
    "2407E":  {"cta": "1.1.1.06.6.003", "moneda": "USD"},
}

MAPEO_CB_CG_FEBECA = {
    "0102EU": {"cta": "1.1.1.02.6.213", "moneda": "EUR"},
    "0102L":  {"cta": "1.1.1.02.1.016", "moneda": "VES"},
    "0104L":  {"cta": "1.1.1.02.1.112", "moneda": "VES"},
    "0104LP": {"cta": "1.1.1.02.1.116", "moneda": "VES"}, 
    "0105E":  {"cta": "1.1.1.02.6.001", "moneda": "USD"},
    "0105L":  {"cta": "1.1.1.02.1.009", "moneda": "VES"},
    "0108E":  {"cta": "1.1.1.02.6.013", "moneda": "USD"},
    "0108L":  {"cta": "1.1.1.02.1.004", "moneda": "VES"},
    "0114E":  {"cta": "1.1.1.02.6.006", "moneda": "USD"},
    "0114L":  {"cta": "1.1.1.02.1.011", "moneda": "VES"},
    "0115L":  {"cta": "1.1.1.02.1.111", "moneda": "VES"},
    "0134E2": {"cta": "1.1.1.02.6.005", "moneda": "USD"},
    "0134EC": {"cta": "1.1.1.02.6.005", "moneda": "USD"},
    "0134L":  {"cta": "1.1.1.02.1.007", "moneda": "VES"},
    "0134L2": {"cta": "1.1.1.02.1.007", "moneda": "VES"}, 
    "0137CP": {"cta": "1.1.1.02.6.214", "moneda": "COP"},
    "0137E":  {"cta": "1.1.1.02.6.015", "moneda": "USD"},
    "0137L":  {"cta": "1.1.1.02.1.022", "moneda": "VES"},
    "0172E":  {"cta": "1.1.1.02.6.011", "moneda": "USD"},
    "0172L":  {"cta": "1.1.1.02.1.021", "moneda": "VES"},
    "0174E":  {"cta": "1.1.1.02.6.010", "moneda": "USD"},
    "0174EU": {"cta": "1.1.1.02.6.210", "moneda": "EUR"},
    "0174L":  {"cta": "1.1.1.02.1.132", "moneda": "VES"},
    "0175L":  {"cta": "1.1.1.02.1.115", "moneda": "VES"},
    "0191E":  {"cta": "1.1.1.02.6.002", "moneda": "USD"},
    "0191L":  {"cta": "1.1.1.02.1.124", "moneda": "VES"},
    "0201E":  {"cta": "1.1.1.03.6.012", "moneda": "USD"},
    "0202E":  {"cta": "1.1.1.03.6.002", "moneda": "USD"},
    "0202E2": {"cta": "1.1.1.03.6.002", "moneda": "USD"},
    "0202E3": {"cta": "1.1.1.03.6.002", "moneda": "USD"},
    "0203E":  {"cta": "1.1.4.01.7.020", "moneda": "USD"},
    "0204E":  {"cta": "1.1.1.03.6.028", "moneda": "USD"},
    "0205E":  {"cta": "1.1.1.03.6.026", "moneda": "USD"},
    "0206E":  {"cta": "1.1.1.06.6.001", "moneda": "USD"},
    "0207E":  {"cta": "1.1.4.01.7.021", "moneda": "USD"},
    "0211E":  {"cta": "1.1.1.03.6.015", "moneda": "USD"},
    "0212E":  {"cta": "1.1.1.03.6.031", "moneda": "USD"},
    "0501E":  {"cta": "1.1.1.03.6.024", "moneda": "USD"},
    "2407E":  {"cta": "1.1.1.06.6.003", "moneda": "USD"},
}

MAPEO_CB_CG_PRISMA = {
    "0104L":  {"cta": "1.1.1.02.1.112", "moneda": "VES"},
    "0105E":  {"cta": "1.1.1.02.6.001", "moneda": "USD"},
    "0105L":  {"cta": "1.1.1.02.1.003", "moneda": "VES"},
    "0114L":  {"cta": "1.1.1.02.1.005", "moneda": "VES"},
    "0174E":  {"cta": "1.1.1.02.6.010", "moneda": "USD"},
    "0174L":  {"cta": "1.1.1.02.1.132", "moneda": "VES"},
    "0175L":  {"cta": "1.1.1.02.1.115", "moneda": "VES"},
    "0205E":  {"cta": "1.1.1.03.6.026", "moneda": "USD"},
    "0209E":  {"cta": "1.1.1.01.6.001", "moneda": "USD"},
    "0501E":  {"cta": "1.1.1.03.6.024", "moneda": "USD"},
}

# QUINCALLA USA EL MISMO DICCIONARIO QUE FEBECA + SILLACA (UNIFICADO)
MAPEO_CB_CG_QUINCALLA = MAPEO_CB_CG_FEBECA

