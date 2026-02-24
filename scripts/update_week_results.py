
import sys
import os

# Añadimos la ruta de los scripts para poder importar el módulo principal
sys.path.append(r'c:\Users\Equipo1\Documents\Antigravity\Skills\scripts')
from actualizar_pick_and_roll import update_pick_and_roll

# Ruta del Excel en red
PATH_EXCEL = r'\\172.16.80.129\g\EXCEL PROGRAMAS\PICK AND ROLL\DATOS PICK AND ROLL.xlsx'

# --- DATOS EXTRAÍDOS (Jornada 19 - 22/02/2026) ---

# TERCERA FEB - GRUPO D-B
tercera_resultados = [
    ("CBA SPAIN - D.R. PEÑARROYA", "54 - 87", "21/02/2026"),
    ("HACHE P. MORALEJA - BAUBLOCK GYMNÁSTICA", "58 - 81", "21/02/2026"),
    ("I. CB ALCALÁ CAJA87 - BOSCO MÉRIDA", "50 - 83", "21/02/2026"),
    ("SAN ANTONIO CÁCERES - CB DOS HERMANAS", "62 - 73", "22/02/2026"),
    ("HUELVA COMERCIO VIRIDIS - BC BADAJOZ", "52 - 73", "22/02/2026"),
    ("CB SAN FERNANDO - SAGRADO CORAZÓN", "85 - 79", "22/02/2026"),
    ("CB CORIA - S.D. ALJARAQUE", "APLAZADO", "26/03/2026")
]

tercera_clasificacion = [
    ("1", "VÍTALY LA MAR BC BADAJOZ", "18", "16", "2", "34"),
    ("2", "BOSCO MÉRIDA PATRIMONIO", "18", "14", "4", "32"),
    ("3", "BAUBLOCK GYMNÁSTICA", "19", "13", "6", "32"),
    ("4", "CB SAN FERNANDO", "19", "11", "8", "30"),
    ("5", "SAN ANTONIO CÁCERES", "18", "11", "7", "29"),
    ("6", "LITHIUM IBERIA SAGRADO", "18", "11", "7", "29"),
    ("7", "HUELVA COMERCIO VIRIDIS", "18", "10", "8", "28"),
    ("8", "ATICA SEVILLA CB CORIA", "17", "10", "7", "27"),
    ("9", "INSOLAC CB ALCALÁ CAJA87", "18", "8", "10", "26"),
    ("10", "D.R. PEÑARROYA", "18", "7", "11", "25"),
    ("11", "EIFFAGE CB DOS HERMANAS", "17", "6", "11", "23"),
    ("12", "CBA SPAIN", "18", "3", "15", "21"),
    ("13", "HACHE P. MORALEJA", "18", "2", "16", "20"),
    ("14", "S.D. ALJARAQUE", "16", "3", "13", "19")
]

# N1 ANDALUZA - GRUPO A
n1_resultados = [
    ("CB GIBRALEÓN - CB CIUDAD DE PALOS", "80 - 60", "21/02/2026"),
    ("CB CORIA - CB LEPE ALIUS", "89 - 60", "21/02/2026"),
    ("LA PALMA 95 - NAUTICO SEVILLA", "74 - 56", "21/02/2026"),
    ("HUELVA LA LUZ - ÉCIJA BASKET", "67 - 100", "22/02/2026"),
    ("CB FRESAS - XEREZ CD", "59 - 64", "22/02/2026"),
    ("CDBC PILAS - CIRCULO MERCANTIL", "57 - 86", "22/02/2026"),
    ("RC LABRADORES - AD REM ONUBA", "APLAZADO", "10/03/2026")
]

n1_clasificacion = [
    ("1", "AD REM ONUBA", "17", "16", "1", "33"),
    ("2", "CIRCULO MERCANTIL", "18", "15", "3", "33"),
    ("3", "ATICA SEVILLA CB CORIA", "18", "12", "6", "30"),
    ("4", "CB FRESAS - SAFA REYES", "19", "10", "9", "29"),
    ("5", "BARNETO MODAS LA PALMA 95", "17", "12", "5", "29"),
    ("6", "CLUB NAUTICO SEVILLA", "18", "10", "8", "28"),
    ("7", "CB LEPE ALIUS EL JAMÓN", "17", "10", "7", "27"),
    ("8", "XEREZ CLUB DEPORTIVO", "19", "8", "11", "27"),
    ("9", "CDBC PILAS", "17", "9", "8", "26"),
    ("10", "CB GIBRALEÓN", "18", "8", "10", "26"),
    ("11", "RC LABRADORES", "16", "8", "8", "24"),
    ("12", "ÉCIJA BASKET", "18", "5", "13", "23"),
    ("13", "CB CIUDAD DE PALOS", "20", "2", "18", "22"),
    ("14", "HUELVA LA LUZ", "18", "0", "18", "18")
]

# --- EJECUCIÓN ---

data_to_update = {
    "TERCERA FEB": tercera_resultados,
    "CLASIFICACION TERCERA FEB": tercera_clasificacion,
    "N1 ANDALUZA": n1_resultados,
    "CLASIFICACION N1 ANDALUZA": n1_clasificacion
}

print("Iniciando actualización de datos del fin de semana...")
update_pick_and_roll(PATH_EXCEL, data_to_update)
