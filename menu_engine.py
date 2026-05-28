"""
Nutrylife — menu_engine.py v3.0 - OPTIMIZADO PARA RENDER

Motor de generación de menús semanales - VERSIÓN OPTIMIZADA:
- Preselecciona candidatos por comida (no genera todas las combos)
- Reduce combinaciones de 598,500 a ~3,888 (99.4% menos)
- Mantiene todas las reglas funcionales intactas
- Compatible con Render gratuito
- Determinístico, sin random
- Salida compatible con PPTX

CAMBIOS PRINCIPALES:
1. Función seleccionar_candidatos_por_comida() nueva
2. Preselecciona ~12 desayunos, ~18 almuerzos, ~18 cenas
3. Genera solo 3.888 combos máximo (vs 598.500)
4. Mantiene diversidad por nombre normalizado
5. Parámetros configurables con defaults
"""

import os
import heapq
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
# VALIDACIONES DE CALIDAD
# ============================================================================

def almuerzo_cena_distintos(almuerzo: Dict, cena: Dict) -> bool:
    nombre_a = normalizar_texto(almuerzo.get("nombre", ""))
    nombre_c = normalizar_texto(cena.get("nombre", ""))
    return nombre_a != nombre_c

def desayuno_distinto_dia_anterior(desayuno_actual: Dict, desayuno_anterior: Dict) -> bool:
    nombre_actual = normalizar_texto(desayuno_actual.get("nombre", ""))
    nombre_anterior = normalizar_texto(desayuno_anterior.get("nombre", ""))
    return nombre_actual != nombre_anterior

# ============================================================================
# PRESELECCIÓN DE CANDIDATOS (NUEVA v3.0)
# ============================================================================

def calcular_score_individual(plato: Dict, meta_comida: Dict) -> float:
    """Score de cercanía a meta de UNA comida."""
    error_prot = abs(plato["proteinas"] - meta_comida["proteinas"])
    error_carb = abs(plato["carbohidratos"] - meta_comida["carbohidratos"])
    error_gras = abs(plato["grasas"] - meta_comida["grasas"])
    return error_prot * 100 + error_carb * 40 + error_gras * 30

def seleccionar_candidatos_por_comida(
    platos: List[Dict],
    tipo_comida: str,
    meta_comida: Dict[str, float],
    max_candidatos: int = 22
) -> List[Dict]:
    """
    Preselecciona candidatos maximizando diversidad de NOMBRES, no solo score.
    
    ESTRATEGIA v3.0:
    1. Agrupar por nombre normalizado
    2. Tomar MEJOR de cada nombre (garantiza diversidad de nombres)
    3. Si hay menos nombres que max_candidatos, agregar más versiones
    4. Rankear por score para mantener determinismo
    
    IMPORTANTE: Garantiza 1+ plato de cada nombre único, luego llena con alternativas.
    """
    if len(platos) <= max_candidatos:
        return sorted(platos, key=lambda p: calcular_score_individual(p, meta_comida))
    
    # Agrupar por nombre normalizado
    nombres_normalizados = {}
    for plato in platos:
        nombre_norm = normalizar_texto(plato["nombre"])
        if nombre_norm not in nombres_normalizados:
            nombres_normalizados[nombre_norm] = []
        nombres_normalizados[nombre_norm].append(plato)
    
    # Paso 1: Garantizar 1 mejor versión de CADA nombre
    resultado = []
    usados = set()
    for nombre_norm, versiones in sorted(nombres_normalizados.items()):
        mejor = min(versiones, key=lambda p: calcular_score_individual(p, meta_comida))
        resultado.append(mejor)
        usados.add(mejor["id"])
    
    # Paso 2: Llenar cupo restante con más versiones
    if len(resultado) < max_candidatos:
        for plato in sorted(platos, key=lambda p: calcular_score_individual(p, meta_comida)):
            if plato["id"] not in usados:
                resultado.append(plato)
                usados.add(plato["id"])
                if len(resultado) >= max_candidatos:
                    break
    
    # Retornar ordenado por score
    return sorted(resultado[:max_candidatos], key=lambda p: calcular_score_individual(p, meta_comida))

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
# GENERACIÓN DE COMBINACIONES (OPTIMIZADO v3.0)
# ============================================================================

def generar_combinaciones_rankeadas(
    desayunos: List[Dict],
    almuerzos: List[Dict],
    cenas: List[Dict],
    objetivos: Dict[str, float],
    metas_por_comida: Dict[str, Dict[str, float]],
    pool_size: int = 100,
    max_desayunos: int = 22,
    max_almuerzos: int = 24,
    max_cenas: int = 24
) -> List[Tuple[Tuple[Dict, Dict, Dict], float, str]]:
    """
    v3.3: Pool con diversidad garantizada por desayuno, almuerzo y cena.

    El pool se construye en dos partes:
    A) Mitad inferior: mejores combos globales (score nutricional).
    B) Mitad superior: diversidad por desayuno — para cada desayuno candidato,
       se incluyen sus mejores combos, garantizando que la selección semanal
       siempre tenga combos válidas aunque el día anterior usara D001.
    """
    # PASO 1: Preseleccionar candidatos
    desayunos_cand = seleccionar_candidatos_por_comida(
        desayunos, "Desayuno", metas_por_comida["desayuno"], max_desayunos
    )
    almuerzos_cand = seleccionar_candidatos_por_comida(
        almuerzos, "Almuerzo", metas_por_comida["almuerzo"], max_almuerzos
    )
    cenas_cand = seleccionar_candidatos_por_comida(
        cenas, "Cena", metas_por_comida["cena"], max_cenas
    )

    # PASO 2: Generar todas las combos válidas (almuerzo != cena por nombre)
    combos_válidas = []
    for combo in product(desayunos_cand, almuerzos_cand, cenas_cand):
        if almuerzo_cena_distintos(combo[1], combo[2]):
            score = calcular_score(combo, objetivos, metas_por_comida)
            estado = clasificar_estado(combo, objetivos)
            combos_válidas.append((combo, score, estado))

    combos_válidas.sort(key=lambda x: x[1])

    # PASO 3: Construir pool con tres capas de diversidad
    #
    # Capa 1: Para cada DESAYUNO candidato → sus mejores 7 combos
    #   Garantiza que para cualquier día anterior, el pool tiene combos
    #   con desayuno distinto.
    #
    # Capa 2: Para cada ALMUERZO candidato → su mejor combo (con cualquier desayuno)
    #   Garantiza que todos los almuerzos están representados.
    #
    # Capa 3: Relleno con mejores globales no incluidas aún.

    # Agrupar combos por desayuno y por almuerzo
    combos_por_desayuno: Dict[str, List] = {}
    combos_por_almuerzo: Dict[str, List] = {}
    for item in combos_válidas:
        d_n = normalizar_texto(item[0][0]["nombre"])
        a_n = normalizar_texto(item[0][1]["nombre"])
        combos_por_desayuno.setdefault(d_n, []).append(item)
        combos_por_almuerzo.setdefault(a_n, []).append(item)

    pool_ids: set = set()
    pool: List = []

    def _add(item):
        obj_id = id(item[0])
        if obj_id not in pool_ids:
            pool.append(item)
            pool_ids.add(obj_id)

    # Capa 1: mejores 7 combos por desayuno
    cuota_d = 7
    for d_nombre in sorted(combos_por_desayuno.keys()):
        for item in combos_por_desayuno[d_nombre][:cuota_d]:
            _add(item)

    # Capa 2: mejor combo por almuerzo (para asegurar que todos los almuerzos entran)
    for a_nombre in sorted(combos_por_almuerzo.keys()):
        _add(combos_por_almuerzo[a_nombre][0])   # ya ordenado por score

    # Capa 3: rellenar con mejores globales hasta tope generoso
    # El pool necesita ser amplio para que combos_busqueda tenga cobertura real
    # de todos los pares A/C, especialmente con bases donde pocos platos dominan.
    tope = max(pool_size * 15, 1500)
    for item in combos_válidas:
        if len(pool) >= tope:
            break
        _add(item)

    # Log de diversidad
    n_d = len(set(normalizar_texto(c[0]["nombre"]) for c, _, _ in pool))
    n_a = len(set(normalizar_texto(c[1]["nombre"]) for c, _, _ in pool))
    n_c = len(set(normalizar_texto(c[2]["nombre"]) for c, _, _ in pool))
    print(f"[menu_engine v3.7] Candidatos D/A/C = {len(desayunos_cand)}/{len(almuerzos_cand)}/{len(cenas_cand)}")
    print(f"[menu_engine v3.7] Combinaciones evaluadas = {len(combos_válidas)}")
    print(f"[menu_engine v3.7] Nombres únicos en pool D/A/C = {n_d}/{n_a}/{n_c}")

    return pool

# ============================================================================
# PENALIZACIONES
# ============================================================================

# ============================================================================
# REGLAS SEMANALES DE CALIDAD - v3.0
# ============================================================================

def obtener_nombre_normalizado_plato(plato: Dict) -> str:
    """Retorna nombre normalizado para agrupar platos visualmente idénticos."""
    return normalizar_texto(plato.get("nombre", ""))

def plato_excede_maximo_semanal(plato_nombre_norm: str, conteo_semana: Dict[str, int], max_apariciones: int = 3) -> bool:
    """Verifica si un plato ya apareció max_apariciones veces en la semana."""
    return conteo_semana.get(plato_nombre_norm, 0) >= max_apariciones

def plato_en_dia_anterior(plato_nombre_norm: str, combos_usados: List[Tuple[Dict, Dict, Dict]]) -> bool:
    """Verifica si el plato apareció en almuerzo o cena del día anterior."""
    if not combos_usados:
        return False
    
    d_ayer, a_ayer, c_ayer = combos_usados[-1]
    a_ayer_norm = obtener_nombre_normalizado_plato(a_ayer)
    c_ayer_norm = obtener_nombre_normalizado_plato(c_ayer)
    
    return plato_nombre_norm == a_ayer_norm or plato_nombre_norm == c_ayer_norm

# ============================================================================
# VALIDACIÓN DE REGLAS DURAS - v3.0
# ============================================================================

def cumple_reglas_duras_semana(
    combo: Tuple[Dict, Dict, Dict],
    combos_usados: List[Tuple[Dict, Dict, Dict]],
    conteo_platos_semana: Dict[str, int],
    max_apariciones: int = 3
) -> bool:
    """
    v3.0: Valida reglas DURAS (no penalizaciones).
    
    Reglas que DEBEN cumplirse (retorna False si alguna viola):
    1. Almuerzo != Cena en mismo día (por nombre normalizado)
    2. Desayuno != desayuno del día anterior
    3. Almuerzo/Cena no aparecen en el día anterior
    4. Conteo semanal de almuerzo < max_apariciones
    5. Conteo semanal de cena < max_apariciones
    """
    d_actual, a_actual, c_actual = combo
    d_norm = obtener_nombre_normalizado_plato(d_actual)
    a_norm = obtener_nombre_normalizado_plato(a_actual)
    c_norm = obtener_nombre_normalizado_plato(c_actual)
    
    # REGLA 1: Almuerzo != Cena en mismo día
    if a_norm == c_norm:
        return False
    
    # Reglas que dependen del día anterior
    if combos_usados:
        d_ayer, a_ayer, c_ayer = combos_usados[-1]
        d_ayer_norm = obtener_nombre_normalizado_plato(d_ayer)
        a_ayer_norm = obtener_nombre_normalizado_plato(a_ayer)
        c_ayer_norm = obtener_nombre_normalizado_plato(c_ayer)
        
        # REGLA 2: Desayuno != desayuno anterior
        if d_norm == d_ayer_norm:
            return False
        
        # REGLA 3: Almuerzo/Cena no aparecen en día anterior
        if a_norm == a_ayer_norm or a_norm == c_ayer_norm:
            return False
        if c_norm == a_ayer_norm or c_norm == c_ayer_norm:
            return False
    
    # REGLA 4 y 5: Máximo 3 apariciones semanales
    if conteo_platos_semana.get(a_norm, 0) >= max_apariciones:
        return False
    if conteo_platos_semana.get(c_norm, 0) >= max_apariciones:
        return False
    
    return True

def calcular_penalizacion_semanal_v2(
    combo: Tuple[Dict, Dict, Dict],
    combos_usados: List[Tuple[Dict, Dict, Dict]],
    conteo_semana: Dict[str, int],
    max_apariciones_plato: int = 3
) -> float:
    """
    v3.0: Penalizaciones con reglas semanales de calidad por nombre normalizado.
    
    Reglas:
    1. Almuerzo != Cena en mismo día (por nombre normalizado)
    2. Desayuno no en días consecutivos
    3. No repetir bases
    4. Máximo 3 apariciones semanales por plato (almuerzo+cena)
    5. No plato almuerzo/cena en días consecutivos
    """
    d_actual, a_actual, c_actual = combo
    d_base_actual = d_actual["id_base"]
    a_base_actual = a_actual["id_base"]
    c_base_actual = c_actual["id_base"]
    a_norm = obtener_nombre_normalizado_plato(a_actual)
    c_norm = obtener_nombre_normalizado_plato(c_actual)
    
    penalizacion = 0.0
    
    # REGLA 1: Almuerzo != Cena en mismo día (por nombre normalizado)
    if a_norm == c_norm:
        penalizacion += 500000
    
    # REGLA 2: Desayuno no en días consecutivos
    if combos_usados:
        d_ayer, a_ayer, c_ayer = combos_usados[-1]
        if not desayuno_distinto_dia_anterior(d_actual, d_ayer):
            penalizacion += 1000000  # CRÍTICA: Evitar desayuno repetido
    
    # REGLA 3: Máximo 3 apariciones semanales por plato (almuerzo+cena)
    if plato_excede_maximo_semanal(a_norm, conteo_semana, max_apariciones_plato):
        penalizacion += 350000  # Penalización muy alta
    if plato_excede_maximo_semanal(c_norm, conteo_semana, max_apariciones_plato):
        penalizacion += 350000  # Penalización muy alta
    
    # REGLA 4: No plato almuerzo/cena en días consecutivos
    if plato_en_dia_anterior(a_norm, combos_usados):
        penalizacion += 300000  # Penalización alta
    if plato_en_dia_anterior(c_norm, combos_usados):
        penalizacion += 300000  # Penalización alta
    
    # REGLA 5: No repetir bases (antigua)
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
    
    # Extra penalty para no repetir día anterior
    if combos_usados:
        d_ayer, a_ayer, c_ayer = combos_usados[-1]
        if a_ayer["id_base"] == a_base_actual:
            penalizacion += 3000
        if c_ayer["id_base"] == c_base_actual:
            penalizacion += 3000
    
    return penalizacion

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

def generar_menu_semanal(
    datos: Dict[str, Any],
    excel_path: str = None,
    top_n: int = 10,
    pool_size: int = 100,
    max_desayunos: int = 22,
    max_almuerzos: int = 24,
    max_cenas: int = 24
) -> Dict[str, Any]:
    """
    v3.0: Genera menú semanal optimizado para Render.
    
    Parámetros nuevos (con defaults):
    - pool_size: tamaño del pool final (default 100)
    - max_desayunos: candidatos por preselección (default 12)
    - max_almuerzos: candidatos por preselección (default 18)
    - max_cenas: candidatos por preselección (default 18)
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
    
    # Pool optimizado
    combinaciones_pool = generar_combinaciones_rankeadas(
        desayunos, almuerzos, cenas, objetivos, metas_por_comida,
        pool_size=pool_size,
        max_desayunos=max_desayunos,
        max_almuerzos=max_almuerzos,
        max_cenas=max_cenas
    )
    
    if not combinaciones_pool:
        raise ValueError("No se generaron combinaciones válidas")
    
    # ── Selección semanal v3.7: BEAM SEARCH CON REINTENTOS ESCALONADOS ──────
    #
    # Cambios respecto a v3.6:
    # 1. Construcción de combos_busqueda parametrizable (tamaño objetivo).
    # 2. Beam search encapsulado en función _correr_beam(combos, width).
    # 3. Poda diversificada: mitad por score, mitad por diversidad estructural
    #    (estados que han usado más nombres únicos de A/C). Esto evita que el
    #    beam quede dominado por estados con mismos pocos platos buenos.
    # 4. Reintentos escalonados antes de activar fallback:
    #       intento 1: combos=800,  width=120
    #       intento 2: combos=1500, width=200
    #       intento 3: combos=2500, width=300
    # 5. Fallback solo si los 3 intentos fallan.

    dias_semana = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
    advertencias_reglas = []

    pool_sorted = sorted(combinaciones_pool, key=lambda x: x[1])

    # Agrupar pool por almuerzo/cena/par para reusar en cada construcción
    por_a: Dict[str, list] = {}
    por_c: Dict[str, list] = {}
    por_ac: Dict[tuple, tuple] = {}
    for item in pool_sorted:
        a_n = normalizar_texto(item[0][1]["nombre"])
        c_n = normalizar_texto(item[0][2]["nombre"])
        por_a.setdefault(a_n, []).append(item)
        por_c.setdefault(c_n, []).append(item)
        if (a_n, c_n) not in por_ac:
            por_ac[(a_n, c_n)] = item

    def _construir_combos_busqueda(tam_objetivo: int) -> list:
        """Construye combos_busqueda con cobertura diversa hasta tam_objetivo."""
        vistas: set = set()
        cb: list = []

        def _add(item):
            key = (item[0][0]["id"], item[0][1]["id"], item[0][2]["id"])
            if key not in vistas:
                vistas.add(key)
                cb.append(item)

        # Capa 1: mejores globales (~20% del tamaño objetivo)
        for item in pool_sorted[:max(200, tam_objetivo // 5)]:
            _add(item)

        # Capa 2: top-N combos por almuerzo_norm — cobertura cruzada
        top_por_a = max(5, tam_objetivo // 100)
        for lst in por_a.values():
            for item in lst[:top_por_a]:
                _add(item)

        # Capa 3: top-N combos por cena_norm
        top_por_c = max(5, tam_objetivo // 100)
        for lst in por_c.values():
            for item in lst[:top_por_c]:
                _add(item)

        # Capa 4: mejor combo por par (almuerzo_norm, cena_norm)
        for item in por_ac.values():
            _add(item)

        # Capa 5: relleno con mejores globales no incluidas
        for item in pool_sorted:
            if len(cb) >= tam_objetivo:
                break
            _add(item)

        cb.sort(key=lambda x: x[1])

        # Precomputar nombres normalizados una sola vez
        precomputado = []
        for combo, score_nutri, estado_nutri in cb:
            d, a, c = combo
            precomputado.append((
                d, a, c, score_nutri, estado_nutri,
                normalizar_texto(d["nombre"]),
                normalizar_texto(a["nombre"]),
                normalizar_texto(c["nombre"]),
            ))
        return precomputado

    EstadoInicial = {
        "dias": [],
        "conteo_ac": {},
        "conteo_d": {},          # contador de desayunos por nombre normalizado
        "ac_ayer": set(),
        "d_ayer": "",
        "score_total": 0.0,
        "ids_key": (),
    }

    # Penalizaciones suaves por repetición (se suman al score nutricional
    # para crear un "score ajustado" de estado — no son reglas duras).
    # Esto inclina el beam hacia estados más variados sin bloquear el algoritmo.
    PEN_D_2 = 2000    # segunda aparición del mismo desayuno
    PEN_D_3 = 8000    # tercera aparición del mismo desayuno
    PEN_AC_2 = 1000   # segunda aparición del mismo plato A/C
    PEN_AC_3 = 4000   # tercera aparición del mismo plato A/C

    def _score_ajustado(estado):
        """Score nutricional + penalizaciones suaves por repetición."""
        pen = 0
        for d_n, cnt in estado["conteo_d"].items():
            if cnt >= 3: pen += PEN_D_3
            elif cnt == 2: pen += PEN_D_2
        for ac_n, cnt in estado["conteo_ac"].items():
            if cnt >= 3: pen += PEN_AC_3
            elif cnt == 2: pen += PEN_AC_2
        return estado["score_total"] + pen

    def _key_score(estado):
        # Score ajustado (incluye penalizaciones suaves de diversidad)
        # luego desempate por más desayunos únicos, más A/C únicos, estabilidad
        return (
            _score_ajustado(estado),
            -len(estado["conteo_d"]),    # más desayunos únicos = mejor
            -len(estado["conteo_ac"]),   # más A/C únicos = mejor
            estado["ids_key"],
        )

    def _key_diversidad(estado):
        # Diversidad estructural: más nombres únicos primero,
        # luego menor concentración, luego score nutricional puro
        max_conc_ac = max(estado["conteo_ac"].values()) if estado["conteo_ac"] else 0
        max_conc_d  = max(estado["conteo_d"].values())  if estado["conteo_d"]  else 0
        return (
            -len(estado["conteo_d"]),    # más desayunos únicos = mejor
            -len(estado["conteo_ac"]),   # más A/C únicos = mejor
            max_conc_d,                  # menor concentración de desayuno
            max_conc_ac,                 # menor concentración A/C
            estado["score_total"],
            estado["ids_key"],
        )

    def _correr_beam(combos_pre: list, beam_width: int):
        """
        Ejecuta beam search completo. Retorna (estados_completos, ultimo_beam).
        Si estados_completos es no vacío, hay solución; si no, ultimo_beam tiene
        los mejores estados parciales para el fallback.
        """
        beam = [EstadoInicial]

        for dia_nombre in dias_semana:
            nuevos_estados = []

            for estado in beam:
                ec = estado["conteo_ac"]
                eac = estado["ac_ayer"]
                eda = estado["d_ayer"]
                ed = estado["conteo_d"]     # contador de desayunos
                es = estado["score_total"]
                edias = estado["dias"]
                eids = estado["ids_key"]

                for d, a, c, score_nutri, estado_nutri, d_n, a_n, c_n in combos_pre:
                    # Reglas duras (filtro estricto)
                    if a_n == c_n:                                   continue  # R1
                    if eda and d_n == eda:                           continue  # R2
                    if a_n in eac or c_n in eac:                     continue  # R3
                    if ec.get(a_n, 0) >= 3:                          continue  # R4
                    if ec.get(c_n, 0) >= 3:                          continue  # R5
                    if ed.get(d_n, 0) >= 2:                          continue  # R6: max 2 apariciones por desayuno

                    nuevo_conteo = dict(ec)
                    nuevo_conteo[a_n] = nuevo_conteo.get(a_n, 0) + 1
                    nuevo_conteo[c_n] = nuevo_conteo.get(c_n, 0) + 1

                    nuevo_conteo_d = dict(ed)
                    nuevo_conteo_d[d_n] = nuevo_conteo_d.get(d_n, 0) + 1

                    nuevos_estados.append({
                        "dias": edias + [(dia_nombre, d, a, c, score_nutri, estado_nutri)],
                        "conteo_ac": nuevo_conteo,
                        "conteo_d": nuevo_conteo_d,
                        "ac_ayer": {a_n, c_n},
                        "d_ayer": d_n,
                        "score_total": es + score_nutri,
                        "ids_key": eids + ((d["id"], a["id"], c["id"]),),
                    })

            if not nuevos_estados:
                # No hay caminos en este día — retornar último beam para fallback
                return [], beam

            # Poda diversificada: mitad por score, mitad por diversidad estructural.
            # Esto evita que el beam quede dominado por estados que comparten los
            # mismos pocos platos buenos y se quedan sin opciones más adelante.
            mitad = beam_width // 2
            mejores_score = heapq.nsmallest(mitad, nuevos_estados, key=_key_score)
            ids_score = {id(e) for e in mejores_score}
            restantes = [e for e in nuevos_estados if id(e) not in ids_score]
            mejores_div = heapq.nsmallest(beam_width - mitad, restantes, key=_key_diversidad)
            beam = mejores_score + mejores_div

        estados_completos = [e for e in beam if len(e["dias"]) == 7]
        return estados_completos, beam

    # Reintentos escalonados
    intentos = [
        (800,  120),
        (1500, 200),
        (2500, 300),
    ]

    print(f"[menu_engine v3.7] Candidatos A/C únicos en pool = {len(por_a)}/{len(por_c)}")
    estados_completos = []

    for n_intento, (n_combos, bw) in enumerate(intentos, start=1):
        combos_busqueda_pre = _construir_combos_busqueda(n_combos)
        n_combos_real = len(combos_busqueda_pre)
        BEAM_WIDTH_USADO = bw
        estados_completos, ultimo_beam = _correr_beam(combos_busqueda_pre, bw)
        exito = len(estados_completos) > 0
        print(f"[menu_engine v3.7] Intento beam {n_intento}: combos={n_combos_real}, width={bw}, éxito={exito}")
        if exito:
            break

    print(f"[menu_engine v3.7] Combos búsqueda semanal = {len(combos_busqueda_pre)}")
    print(f"[menu_engine v3.7] Beam width = {BEAM_WIDTH_USADO}")

    # Helper: materializa una tupla ligera de día en el dict pesado de salida
    def _materializar_dia(dia_tuple):
        nombre_dia, d, a, c, score_nutri, estado_nutri = dia_tuple
        total_d = d["proteinas"] + a["proteinas"] + c["proteinas"]
        total_carb = d["carbohidratos"] + a["carbohidratos"] + c["carbohidratos"]
        total_gras = d["grasas"] + a["grasas"] + c["grasas"]
        total_kcal = d["kcal"] + a["kcal"] + c["kcal"]
        return {
            "dia": nombre_dia,
            "desayuno": _formato_plato_salida(d),
            "almuerzo": _formato_plato_salida(a),
            "cena": _formato_plato_salida(c),
            "totales": {
                "proteinas": round(total_d, 1),
                "carbohidratos": round(total_carb, 1),
                "grasas": round(total_gras, 1),
                "kcal": round(total_kcal, 1),
            },
            "objetivos": objetivos,
            "metas_por_comida": metas_por_comida,
            "score": round(score_nutri, 2),
            "score_nutricional": round(score_nutri, 2),
            "estado": estado_nutri,
        }

    if estados_completos:
        # Hay solución bajo reglas duras — elegir la mejor por score
        estados_completos.sort(key=_key_score)
        mejor = estados_completos[0]
        menu_dias = [_materializar_dia(dt) for dt in mejor["dias"]]
        conteo_platos_semana = mejor["conteo_ac"]
        conteo_desayunos_semana = mejor["conteo_d"]
        print(f"[menu_engine v3.7] Beam completó semana con reglas duras = Sí")
    else:
        # Los 3 intentos fallaron — fallback relajado
        advertencias_reglas.append(
            "Semana incompleta con reglas duras tras 3 intentos; completada con fallback relajado"
        )
        print(f"[menu_engine v3.7] Beam completó semana con reglas duras = No (fallback activado)")

        mejor_parcial = max(ultimo_beam, key=lambda e: len(e["dias"])) if ultimo_beam else EstadoInicial
        menu_dias = [_materializar_dia(dt) for dt in mejor_parcial["dias"]]
        conteo_platos_semana = dict(mejor_parcial.get("conteo_ac", {}))
        conteo_desayunos_semana = dict(mejor_parcial.get("conteo_d", {}))

        # Reconstruir contexto del último día del estado parcial
        combos_fb_usados_raw = []
        for dia_tuple in mejor_parcial["dias"]:
            _, d, a, c, _, _ = dia_tuple
            combos_fb_usados_raw.append((
                normalizar_texto(d["nombre"]),
                {normalizar_texto(a["nombre"]), normalizar_texto(c["nombre"])},
            ))

        # Completar días faltantes con greedy relajado (penalizaciones, no reglas duras)
        for dia_nombre in dias_semana[len(menu_dias):]:
            d_ant = combos_fb_usados_raw[-1][0] if combos_fb_usados_raw else ""
            ac_ant = combos_fb_usados_raw[-1][1] if combos_fb_usados_raw else set()

            candidatas = []
            for d, a, c, score_nutri, estado_nutri, d_n, a_n, c_n in combos_busqueda_pre:
                if a_n == c_n:
                    continue
                pen = 0
                if d_n == d_ant: pen += 100000
                if a_n in ac_ant or c_n in ac_ant: pen += 50000
                if conteo_platos_semana.get(a_n, 0) >= 3: pen += 80000
                if conteo_platos_semana.get(c_n, 0) >= 3: pen += 80000
                if conteo_desayunos_semana.get(d_n, 0) >= 2: pen += 100000
                candidatas.append((d, a, c, score_nutri, estado_nutri, d_n, a_n, c_n, score_nutri + pen))

            candidatas.sort(key=lambda x: x[8])
            d, a, c, score_nutri, estado_nutri, d_n, a_n, c_n, _ = candidatas[0]

            menu_dias.append(_materializar_dia((dia_nombre, d, a, c, score_nutri, estado_nutri)))
            conteo_platos_semana[a_n] = conteo_platos_semana.get(a_n, 0) + 1
            conteo_platos_semana[c_n] = conteo_platos_semana.get(c_n, 0) + 1
            conteo_desayunos_semana[d_n] = conteo_desayunos_semana.get(d_n, 0) + 1
            combos_fb_usados_raw.append((d_n, {a_n, c_n}))

    resumen = {
        "objetivo_diario": objetivos,
        "distribucion": {
            "proteinas": datos.get("prot_distribucion"),
            "carbohidratos": datos.get("carb_distribucion"),
            "grasas": datos.get("gras_distribucion"),
        },
        "restricciones": restricciones_usuario,
        "advertencias_carga": advertencias_carga,
        "advertencias": advertencias_reglas,
    }

    # Logs finales — variedad semanal
    d_unicos = len(conteo_desayunos_semana)
    ac_unicos = len(conteo_platos_semana)
    print(f"[menu_engine v3.7] Desayunos únicos en semana = {d_unicos}/7")
    print(f"[menu_engine v3.7] Platos A/C únicos en semana = {ac_unicos}/14")
    print(f"[menu_engine v3.7] Conteo desayunos = {dict(sorted(conteo_desayunos_semana.items(), key=lambda x: -x[1]))}")
    print(f"[menu_engine v3.7] Conteo platos A/C = {dict(sorted(conteo_platos_semana.items(), key=lambda x: -x[1]))}")

    # ── Generar opciones por categoría ───────────────────────────────────────
    # Catálogo de hasta 7 platos únicos por nombre normalizado para cada
    # tipo de comida. Respetan restricciones del paciente y están rankeados
    # por cercanía a la meta nutricional individual. No aplican reglas semanales.
    def _generar_opciones_categoria(platos_validos, meta_comida, max_opciones=7):
        """Devuelve hasta max_opciones platos, 1 por nombre normalizado, por score."""
        # Agrupar por nombre normalizado — quedarnos con la mejor versión de cada nombre
        mejores_por_nombre: Dict[str, Dict] = {}
        for plato in platos_validos:
            nombre_norm = normalizar_texto(plato["nombre"])
            if nombre_norm not in mejores_por_nombre:
                mejores_por_nombre[nombre_norm] = plato
            else:
                # Mantener la versión con mejor score individual
                if calcular_score_individual(plato, meta_comida) < calcular_score_individual(mejores_por_nombre[nombre_norm], meta_comida):
                    mejores_por_nombre[nombre_norm] = plato

        # Rankear por score y devolver los mejores max_opciones
        ordenados = sorted(mejores_por_nombre.values(),
                           key=lambda p: calcular_score_individual(p, meta_comida))
        return [_formato_plato_salida(p) for p in ordenados[:max_opciones]]

    opciones_desayunos = _generar_opciones_categoria(desayunos, metas_por_comida["desayuno"])
    opciones_almuerzos = _generar_opciones_categoria(almuerzos, metas_por_comida["almuerzo"])
    opciones_cenas     = _generar_opciones_categoria(cenas,     metas_por_comida["cena"])

    # Advertir si alguna categoría quedó con menos de 7 por restricciones
    for tipo, lista in [("desayunos", opciones_desayunos),
                        ("almuerzos", opciones_almuerzos),
                        ("cenas",     opciones_cenas)]:
        if len(lista) < 7:
            resumen["advertencias"].append(
                f"opciones.{tipo}: solo {len(lista)} platos únicos disponibles (restricciones aplicadas)"
            )

    print(f"[menu_engine v3.7] Opciones D/A/C = {len(opciones_desayunos)}/{len(opciones_almuerzos)}/{len(opciones_cenas)}")

    return {
        "dias": menu_dias,
        "resumen": resumen,
        "opciones": {
            "desayunos": opciones_desayunos,
            "almuerzos": opciones_almuerzos,
            "cenas":     opciones_cenas,
        },
    }

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
    print("MENU ENGINE v3.7 - OPTIMIZADO PARA RENDER + REGLAS SEMANALES")
    print("=" * 80)
    
    try:
        resultado = generar_menu_semanal(datos_demo, pool_size=100)
        
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
        
        print("\n✅ Menú generado exitosamente")
        
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()
