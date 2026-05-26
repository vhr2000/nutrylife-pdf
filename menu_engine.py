"""
Nutrylife — menu_engine.py v2.7 - DIVERSIFICACIÓN MEJORADA
"""

import os
import pandas as pd
from itertools import product
from typing import Dict, List, Tuple, Any, Set

# ============================================================================
# RESTRICCIONES
# ============================================================================

RESTRICTION_ALIASES = {
    "huevo": {"huevo", "huevos"},
    "lacteo": {"lacteo", "lácteo", "lacteos", "lácteos", "leche", "queso", "yogurt", "yogur", "mantequilla", "crema", "quesillo", "ricotta", "mozzarella"},
    "gluten": {"gluten", "pan", "pasta", "harina", "trigo", "tortilla"},
    "pescado": {"pescado", "salmón", "salmon", "atún", "atun", "reineta", "merluza", "jurel"},
    "mariscos": {"marisco", "mariscos", "camarón", "camaron", "camarones", "ostión", "ostion", "choritos", "machas"},
    "vacuno": {"vacuno", "carne", "posta", "asiento", "lomo", "filete", "plateada"},
    "cerdo": {"cerdo", "jamón", "jamon", "jamón serrano", "tocino", "panceta", "longaniza"},
    "frutos_secos": {"nuez", "nueces", "almendra", "almendras", "maní", "mani", "avellana", "pistacho", "castaña", "frutos secos"},
    "soya": {"soya", "soja", "tofu", "salsa de soya", "salsa de soja"},
    "picante": {"ají", "aji", "merkén", "merquen", "chile", "picante"},
}

SPECIAL_RESTRICTIONS = {
    "vegetariano": {"pollo", "pavo", "vacuno", "carne", "pescado", "salmon", "atun", "mariscos", "camarones", "cerdo", "jamon"},
    "vegano": {"pollo", "pavo", "vacuno", "carne", "pescado", "salmon", "atun", "mariscos", "camarones", "cerdo", "jamon", "huevo", "lacteo", "leche", "queso", "yogurt", "mantequilla", "crema"},
}

# ============================================================================
# UTILIDADES
# ============================================================================

def normalizar_texto(texto: str) -> str:
    if not isinstance(texto, str):
        return ""
    texto = texto.lower().strip()
    return texto.replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u")

def obtener_numero(valor: Any, default: float = 0.0) -> float:
    if valor is None or pd.isna(valor):
        return default
    try:
        return float(valor)
    except (ValueError, TypeError):
        return default

def obtener_id_plato_base(id_opcion: str) -> str:
    if not isinstance(id_opcion, str):
        return str(id_opcion)
    return id_opcion.split('-')[0] if '-' in id_opcion else id_opcion

def obtener_tipo_desde_id(id_opcion: str) -> str:
    if not isinstance(id_opcion, str) or len(id_opcion) == 0:
        return None
    primer_char = id_opcion[0].upper()
    if primer_char == 'D':
        return "Desayuno"
    elif primer_char == 'A':
        return "Almuerzo"
    elif primer_char == 'C':
        return "Cena"
    return None

# ============================================================================
# CARGA DE BASE DE DATOS
# ============================================================================

def cargar_base_platos(excel_path: str = None) -> Tuple[List[Dict[str, Any]], List[str]]:
    if excel_path is None:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        excel_path = os.path.join(script_dir, "data", "20260525_Base_Platos_v1.xlsx")
    
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Base no encontrada: {excel_path}")
    
    df = pd.read_excel(excel_path, sheet_name="BD Completa", skiprows=3)
    
    column_mapping = {
        "ID opción": "id", "ID plato": "id_plato", "Tipo comida": "tipo_comida",
        "Plato": "nombre", "Versión": "version", "Ingredientes": "ingredientes",
        "Proteínas (g)": "proteinas", "Carbohidratos (g)": "carbohidratos",
        "Grasas (g)": "grasas", "Kcal calculada": "kcal", "Restricciones": "restricciones",
        "Imagen": "imagen",
    }
    
    for col in column_mapping.keys():
        if col not in df.columns:
            raise ValueError(f"Columna no encontrada: {col}")
    
    df = df.rename(columns=column_mapping)
    
    platos = []
    advertencias = []
    
    for idx, row in df.iterrows():
        tipo_comida = normalizar_texto(str(row["tipo_comida"]))
        
        if "desayuno" in tipo_comida:
            tipo_comida_std = "Desayuno"
        elif "almuerzo" in tipo_comida:
            tipo_comida_std = "Almuerzo"
        elif "cena" in tipo_comida:
            tipo_comida_std = "Cena"
        else:
            continue
        
        version_raw = str(row["version"]).strip() if pd.notna(row["version"]) else "Media"
        restricciones_raw = str(row["restricciones"]) if pd.notna(row["restricciones"]) else ""
        restricciones = [r.strip() for r in restricciones_raw.split(",") if r.strip()]
        
        id_opcion = str(row["id"]).strip()
        
        tipo_esperado = obtener_tipo_desde_id(id_opcion)
        if tipo_esperado != tipo_comida_std:
            advertencias.append(f"Fila {idx+1}: ID {id_opcion} no coincide. Excluida.")
            continue
        
        platos.append({
            "id": id_opcion,
            "id_plato": str(row["id_plato"]).strip(),
            "id_base": obtener_id_plato_base(id_opcion),
            "tipo_comida": tipo_comida_std,
            "nombre": str(row["nombre"]).strip() if pd.notna(row["nombre"]) else "Sin nombre",
            "version": version_raw.capitalize(),
            "imagen": str(row["imagen"]).strip() if pd.notna(row["imagen"]) else None,
            "ingredientes": str(row["ingredientes"]).strip() if pd.notna(row["ingredientes"]) else "",
            "proteinas": obtener_numero(row["proteinas"]),
            "carbohidratos": obtener_numero(row["carbohidratos"]),
            "grasas": obtener_numero(row["grasas"]),
            "kcal": obtener_numero(row["kcal"]),
            "restricciones": restricciones,
        })
    
    if not platos:
        raise ValueError("No hay platos válidos")
    
    return platos, advertencias

# ============================================================================
# VALIDACIONES
# ============================================================================

def validar_distribuciones(datos: Dict[str, Any]) -> None:
    for nombre, dist in [
        ("prot_distribucion", datos.get("prot_distribucion", {})),
        ("carb_distribucion", datos.get("carb_distribucion", {})),
        ("gras_distribucion", datos.get("gras_distribucion", {})),
    ]:
        if not dist or sum(dist.values()) != 100:
            raise ValueError(f"{nombre} es inválida")

def validar_objetivos_diarios(datos: Dict[str, Any]) -> Tuple[float, float, float, float]:
    kcal = obtener_numero(datos.get("kcal"))
    proteinas = obtener_numero(datos.get("proteinas"))
    carbos = obtener_numero(datos.get("carbos"))
    grasas = obtener_numero(datos.get("grasas"))
    if kcal <= 0 or proteinas <= 0 or carbos <= 0 or grasas <= 0:
        raise ValueError("Objetivos inválidos")
    return kcal, proteinas, carbos, grasas

def calcular_metas_por_comida(datos: Dict[str, Any]) -> Dict[str, Dict[str, float]]:
    kcal, proteinas, carbos, grasas = validar_objetivos_diarios(datos)
    validar_distribuciones(datos)
    metas = {}
    for comida in ["desayuno", "almuerzo", "cena"]:
        metas[comida] = {
            "proteinas": (proteinas * datos["prot_distribucion"][comida]) / 100,
            "carbohidratos": (carbos * datos["carb_distribucion"][comida]) / 100,
            "grasas": (grasas * datos["gras_distribucion"][comida]) / 100,
        }
    return metas

# ============================================================================
# RESTRICCIONES
# ============================================================================

def mapear_restriccion_a_canonica(restriccion_usuario: str) -> Set[str]:
    rest_norm = normalizar_texto(restriccion_usuario)
    for canonica, aliases in RESTRICTION_ALIASES.items():
        if rest_norm in {normalizar_texto(a) for a in aliases}:
            return {canonica}
    return {rest_norm}

def construir_restricciones_efectivas(restricciones_usuario: List[str]) -> Set[str]:
    restricciones_efectivas = set()
    for rest in restricciones_usuario:
        if rest and isinstance(rest, str):
            restricciones_efectivas.update(mapear_restriccion_a_canonica(rest))
    return restricciones_efectivas

def plato_viola_restriccion(plato: Dict[str, Any], restricciones_efectivas: Set[str]) -> bool:
    if not restricciones_efectivas:
        return False
    
    ingredientes_texto = normalizar_texto(plato.get("ingredientes", ""))
    
    for rest_plato_norm in [normalizar_texto(r) for r in plato.get("restricciones", [])]:
        if mapear_restriccion_a_canonica(rest_plato_norm) & restricciones_efectivas:
            return True
    
    for restriccion in restricciones_efectivas:
        keywords = SPECIAL_RESTRICTIONS.get(restriccion, RESTRICTION_ALIASES.get(restriccion, set()))
        if any(normalizar_texto(k) in ingredientes_texto for k in keywords):
            return True
    
    return False

def filtrar_por_restricciones(platos: List[Dict[str, Any]], restricciones_usuario: List[str]) -> List[Dict[str, Any]]:
    restricciones_efectivas = construir_restricciones_efectivas(restricciones_usuario)
    return platos if not restricciones_efectivas else [p for p in platos if not plato_viola_restriccion(p, restricciones_efectivas)]

def separar_por_tipo_comida(platos: List[Dict[str, Any]]) -> Dict[str, List[Dict]]:
    por_tipo = {"Desayuno": [], "Almuerzo": [], "Cena": []}
    for plato in platos:
        if plato.get("tipo_comida") in por_tipo:
            por_tipo[plato["tipo_comida"]].append(plato)
    for tipo, lista in por_tipo.items():
        if not lista:
            raise ValueError(f"No hay platos válidos para {tipo}")
    return por_tipo

# ============================================================================
# VALIDACIONES DE CALIDAD DE PAUTA
# ============================================================================

def almuerzo_cena_distintos(almuerzo: Dict, cena: Dict) -> bool:
    """
    Verifica que almuerzo y cena sean platos distintos.
    Compara por nombre normalizado.
    """
    nombre_a = normalizar_texto(almuerzo.get("nombre", ""))
    nombre_c = normalizar_texto(cena.get("nombre", ""))
    return nombre_a != nombre_c

def desayuno_distinto_dia_anterior(desayuno_actual: Dict, desayuno_anterior: Dict) -> bool:
    """
    Verifica que desayuno no se repita en días consecutivos.
    Compara por nombre normalizado.
    """
    nombre_actual = normalizar_texto(desayuno_actual.get("nombre", ""))
    nombre_anterior = normalizar_texto(desayuno_anterior.get("nombre", ""))
    return nombre_actual != nombre_anterior

# ============================================================================
# SCORING
# ============================================================================

def calcular_score(combo: Tuple[Dict, Dict, Dict], objetivos: Dict[str, float], metas_por_comida: Dict[str, Dict[str, float]]) -> float:
    desayuno, almuerzo, cena = combo
    
    total_prot = desayuno["proteinas"] + almuerzo["proteinas"] + cena["proteinas"]
    total_carb = desayuno["carbohidratos"] + almuerzo["carbohidratos"] + cena["carbohidratos"]
    total_gras = desayuno["grasas"] + almuerzo["grasas"] + cena["grasas"]
    total_kcal = desayuno["kcal"] + almuerzo["kcal"] + cena["kcal"]
    
    error_prot_diario = abs(total_prot - objetivos["proteinas"])
    error_carb_diario = abs(total_carb - objetivos["carbos"])
    error_gras_diario = abs(total_gras - objetivos["grasas"])
    error_kcal_pct = abs(total_kcal - objetivos["kcal"]) / objetivos["kcal"] if objetivos["kcal"] > 0 else 0
    
    error_prot_comida = sum([abs(desayuno["proteinas"] - metas_por_comida["desayuno"]["proteinas"]),
                             abs(almuerzo["proteinas"] - metas_por_comida["almuerzo"]["proteinas"]),
                             abs(cena["proteinas"] - metas_por_comida["cena"]["proteinas"])])
    
    error_carb_comida = sum([abs(desayuno["carbohidratos"] - metas_por_comida["desayuno"]["carbohidratos"]),
                             abs(almuerzo["carbohidratos"] - metas_por_comida["almuerzo"]["carbohidratos"]),
                             abs(cena["carbohidratos"] - metas_por_comida["cena"]["carbohidratos"])])
    
    error_gras_comida = sum([abs(desayuno["grasas"] - metas_por_comida["desayuno"]["grasas"]),
                             abs(almuerzo["grasas"] - metas_por_comida["almuerzo"]["grasas"]),
                             abs(cena["grasas"] - metas_por_comida["cena"]["grasas"])])
    
    return (error_prot_diario * 100 + error_prot_comida * 60 +
            error_carb_diario * 40 + error_carb_comida * 25 +
            error_gras_diario * 30 + error_gras_comida * 15 +
            error_kcal_pct * 100)

def clasificar_estado(combo: Tuple[Dict, Dict, Dict], objetivos: Dict[str, float]) -> str:
    desayuno, almuerzo, cena = combo
    total_prot = desayuno["proteinas"] + almuerzo["proteinas"] + cena["proteinas"]
    total_carb = desayuno["carbohidratos"] + almuerzo["carbohidratos"] + cena["carbohidratos"]
    total_gras = desayuno["grasas"] + almuerzo["grasas"] + cena["grasas"]
    total_kcal = desayuno["kcal"] + almuerzo["kcal"] + cena["kcal"]
    
    tolerancias = {"prot_g": 5, "carb_g": 10, "gras_g": 8, "kcal_pct": 0.05}
    error_prot = abs(total_prot - objetivos["proteinas"])
    error_carb = abs(total_carb - objetivos["carbos"])
    error_gras = abs(total_gras - objetivos["grasas"])
    error_kcal_pct = abs(total_kcal - objetivos["kcal"]) / objetivos["kcal"]
    
    if (error_prot <= tolerancias["prot_g"] and error_carb <= tolerancias["carb_g"] and 
        error_gras <= tolerancias["gras_g"] and error_kcal_pct <= tolerancias["kcal_pct"]):
        return "aprobable"
    if (error_prot > tolerancias["prot_g"] * 2 or error_carb > tolerancias["carb_g"] * 2 or 
        error_gras > tolerancias["gras_g"] * 2 or error_kcal_pct > tolerancias["kcal_pct"] * 2):
        return "fuera_de_rango"
    return "requiere_revision"

# ============================================================================
# GENERACIÓN DIVERSIFICADA
# ============================================================================

def generar_combinaciones_rankeadas(desayunos: List[Dict], almuerzos: List[Dict], cenas: List[Dict],
                                     objetivos: Dict[str, float], metas_por_comida: Dict[str, Dict[str, float]],
                                     pool_size: int = 150) -> List[Tuple[Tuple[Dict, Dict, Dict], float, str]]:
    """
    Genera combinaciones D+A+C con DIVERSIFICACIÓN (Opción A).
    
    v2.7: Pool diversificado especialmente en desayunos.
    """
    # PASO 1: Generar TODAS las combos y rankearlas
    all_combos_scored = []
    for idx, combo in enumerate(product(desayunos, almuerzos, cenas)):
        score = calcular_score(combo, objetivos, metas_por_comida)
        estado = clasificar_estado(combo, objetivos)
        all_combos_scored.append((combo, score, estado, idx))
    
    all_combos_scored.sort(key=lambda x: (x[1], x[3]))
    
    # PASO 2: Diversificación GARANTIZADA de d_bases con ORDEN INTERCALADO
    # Garantizar representación de TODOS los d_bases intercalando en el pool
    
    d_bases_included = set()
    selected_by_d_base = {}  # d_base -> lista de posiciones
    
    # Fase 1: Agrupar combos por d_base
    for pos, (combo, score, estado, _) in enumerate(all_combos_scored):
        d_base = combo[0]["id_base"]
        if d_base not in selected_by_d_base:
            selected_by_d_base[d_base] = []
        selected_by_d_base[d_base].append(pos)
    
    # Fase 2: Intercalar d_bases para diversidad
    pool_final_idx = []
    max_per_d_base = max(5, pool_size // len(selected_by_d_base))  # Al menos 5 combos por d_base
    
    for d_base in sorted(selected_by_d_base.keys()):
        for i in range(min(max_per_d_base, len(selected_by_d_base[d_base]))):
            pool_final_idx.append(selected_by_d_base[d_base][i])
            if len(pool_final_idx) >= pool_size:
                break
        if len(pool_final_idx) >= pool_size:
            break
    
    # PASO 7: Retornar pool en orden de índices (mantiene diversidad intercalada)
    pool_final = [all_combos_scored[idx] for idx in pool_final_idx]
    
    return [(c, s, e) for c, s, e, _ in pool_final[:pool_size]]

# ============================================================================
# SALIDA Y MENÚ
# ============================================================================

def _formato_plato_salida(plato: Dict[str, Any]) -> Dict[str, Any]:
    return {
        "id": plato["id"],
        "nombre": plato["nombre"],
        "desc": plato.get("ingredientes", ""),
        "prot": plato["proteinas"],
        "kcal": plato["kcal"],
        "imagen": plato["imagen"],
        "tipo_comida": plato["tipo_comida"],
        "version": plato["version"],
        "ingredientes": plato.get("ingredientes", ""),
        "proteinas": plato["proteinas"],
        "carbohidratos": plato["carbohidratos"],
        "grasas": plato["grasas"],
        "restricciones": ", ".join(plato.get("restricciones", [])),
    }

def calcular_penalizacion_semanal(combo: Tuple[Dict, Dict, Dict],
                                 combos_usados: List[Tuple[Dict, Dict, Dict]]) -> float:
    """
    Penalizaciones para maximizar variedad y calidad de pauta.
    
    Reglas en orden de importancia:
    1. CRÍTICA: Almuerzo != Cena en mismo día
    2. CRÍTICA: Desayuno no en días consecutivos
    3. FUERTE: No repetir id_base
    """
    d_actual, a_actual, c_actual = combo
    d_base_actual = d_actual["id_base"]
    a_base_actual = a_actual["id_base"]
    c_base_actual = c_actual["id_base"]
    
    penalizacion = 0.0
    
    # ===== REGLA 3 (CRÍTICA): Almuerzo != Cena en mismo día =====
    if not almuerzo_cena_distintos(a_actual, c_actual):
        penalizacion += 500000  # PENALIZACIÓN EXTREMA - casi prohibición
    
    # ===== REGLA 2 (CRÍTICA): Desayuno no en días consecutivos =====
    if combos_usados:
        d_ayer, a_ayer, c_ayer = combos_usados[-1]
        if not desayuno_distinto_dia_anterior(d_actual, d_ayer):
            penalizacion += 400000  # PENALIZACIÓN EXTREMA - casi prohibición
    
    # ===== REGLA 1 (FUERTE): No repetir bases =====
    for d, a, c in combos_usados:
        if d["id_base"] == d_base_actual:
            penalizacion += 5000
        if a["id_base"] == a_base_actual:
            penalizacion += 4000
        if c["id_base"] == c_base_actual:
            penalizacion += 4000
        
        if (d["id"] == d_actual["id"] and
            a["id"] == a_actual["id"] and
            c["id"] == c_actual["id"]):
            penalizacion += 200000
    
    # Extra penalty para no repetir día anterior (bases)
    if combos_usados:
        d_ayer, a_ayer, c_ayer = combos_usados[-1]
        if a_ayer["id_base"] == a_base_actual:
            penalizacion += 3000
        if c_ayer["id_base"] == c_base_actual:
            penalizacion += 3000
    
    return penalizacion

def generar_menu_semanal(datos: Dict[str, Any], excel_path: str = None, top_n: int = 10, pool_size: int = 150) -> Dict[str, Any]:
    """
    Genera menú semanal con pool DIVERSIFICADO y penalizaciones agresivas.
    """
    platos_base, advertencias_carga = cargar_base_platos(excel_path)
    
    validar_distribuciones(datos)
    kcal, proteinas, carbos, grasas = validar_objetivos_diarios(datos)
    
    objetivos = {"kcal": kcal, "proteinas": proteinas, "carbos": carbos, "grasas": grasas}
    restricciones_usuario = datos.get("restricciones", [])
    platos_validos = filtrar_por_restricciones(platos_base, restricciones_usuario)
    
    por_tipo = separar_por_tipo_comida(platos_validos)
    desayunos = por_tipo["Desayuno"]
    almuerzos = por_tipo["Almuerzo"]
    cenas = por_tipo["Cena"]
    
    metas_por_comida = calcular_metas_por_comida(datos)
    
    # Pool diversificado
    combinaciones_pool = generar_combinaciones_rankeadas(
        desayunos, almuerzos, cenas, objetivos, metas_por_comida, pool_size=pool_size
    )
    
    if not combinaciones_pool:
        raise ValueError("No se generaron combinaciones válidas")
    
    # Selección semanal
    dias_semana = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
    menu_dias = []
    combos_usados = []
    
    for nombre_dia in dias_semana:
        combos_evaluadas = []
        
        # Evaluar TODAS las combos con penalizaciones
        for combo, score_nutri, estado in combinaciones_pool:
            penalizacion = calcular_penalizacion_semanal(combo, combos_usados)
            score_ajustado = score_nutri + penalizacion
            combos_evaluadas.append((combo, score_nutri, estado, score_ajustado))
        
        # Ordenar por score ajustado
        combos_evaluadas.sort(key=lambda x: (x[3], x[1]))
        
        # Seleccionar: la mejor respetando reglas
        combo, score_nutri, estado, score_ajustado = combos_evaluadas[0]
        
        desayuno, almuerzo, cena = combo
        combos_usados.append(combo)
        
        total_prot = desayuno["proteinas"] + almuerzo["proteinas"] + cena["proteinas"]
        total_carb = desayuno["carbohidratos"] + almuerzo["carbohidratos"] + cena["carbohidratos"]
        total_gras = desayuno["grasas"] + almuerzo["grasas"] + cena["grasas"]
        total_kcal = desayuno["kcal"] + almuerzo["kcal"] + cena["kcal"]
        
        menu_dias.append({
            "dia": nombre_dia,
            "desayuno": _formato_plato_salida(desayuno),
            "almuerzo": _formato_plato_salida(almuerzo),
            "cena": _formato_plato_salida(cena),
            "totales": {"proteinas": round(total_prot, 1), "carbohidratos": round(total_carb, 1),
                       "grasas": round(total_gras, 1), "kcal": round(total_kcal, 1)},
            "objetivos": objetivos,
            "metas_por_comida": metas_por_comida,
            "score": round(score_nutri, 2),
            "score_nutricional": round(score_nutri, 2),
            "estado": estado,
        })
    
    resumen = {
        "objetivo_diario": objetivos,
        "distribucion": {
            "proteinas": datos.get("prot_distribucion"),
            "carbohidratos": datos.get("carb_distribucion"),
            "grasas": datos.get("gras_distribucion"),
        },
        "restricciones": restricciones_usuario,
        "advertencias_carga": advertencias_carga,
        "advertencias": [],
    }
    
    return {"dias": menu_dias, "resumen": resumen}

# ============================================================================
# DEMO
# ============================================================================

if __name__ == "__main__":
    datos_demo = {
        "kcal": 2500,
        "proteinas": 160,
        "carbos": 250,
        "grasas": 85,
        "prot_distribucion": {"desayuno": 20, "almuerzo": 40, "cena": 40},
        "carb_distribucion": {"desayuno": 25, "almuerzo": 45, "cena": 30},
        "gras_distribucion": {"desayuno": 20, "almuerzo": 35, "cena": 45},
        "restricciones": [],
    }
    
    print("=" * 80)
    print("MENU ENGINE v2.7 - COMBINACIONES DIVERSIFICADAS")
    print("=" * 80)
    
    try:
        resultado = generar_menu_semanal(datos_demo, pool_size=200)
        
        print("\n[1] Menú semanal:")
        for dia in resultado["dias"]:
            d_id = dia["desayuno"]["id"]
            a_id = dia["almuerzo"]["id"]
            c_id = dia["cena"]["id"]
            print(f"  {dia['dia']}: {d_id:6s} + {a_id:6s} + {c_id:6s} (score={dia['score']:8.1f})")
        
        print("\n[2] Variedad:")
        d_bases = set(d["desayuno"]["id"].split('-')[0] for d in resultado["dias"])
        a_bases = set(d["almuerzo"]["id"].split('-')[0] for d in resultado["dias"])
        c_bases = set(d["cena"]["id"].split('-')[0] for d in resultado["dias"])
        print(f"  Desayunos: {len(d_bases)}/7 únicos")
        print(f"  Almuerzos: {len(a_bases)}/7 únicos")
        print(f"  Cenas: {len(c_bases)}/7 únicos")
        
        print("\n✅ Menú generado")
        
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()
