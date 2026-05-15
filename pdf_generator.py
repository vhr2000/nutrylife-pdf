"""
Nutrylife — menu_engine.py
Motor de menú: calcula gramajes exactos según macros del plan
y arma el menú de 7 días con restricciones aplicadas.
"""

# ============================================================
# BASE DE DATOS NUTRICIONAL (valores estándar por 100g o unidad)
# Fuente: USDA / Minsal Chile
# Formato: { nombre: { kcal, prot, grasa, carbo, unidad, equiv } }
# unidad: "100g" o "unidad"
# equiv: gramos por unidad (si aplica)
# ============================================================

ALIMENTOS = {
    # ---- PROTEÍNAS ANIMALES ----
    "pollo_pechuga": {
        "nombre": "Pollo (pechuga)",
        "kcal": 165, "prot": 31.0, "grasa": 3.6, "carbo": 0.0,
        "unidad": "100g", "categoria": "Proteínas de tierra"
    },
    "pollo_muslo": {
        "nombre": "Pollo (muslo sin piel)",
        "kcal": 177, "prot": 24.0, "grasa": 8.0, "carbo": 0.0,
        "unidad": "100g", "categoria": "Proteínas de tierra"
    },
    "vacuno_magro": {
        "nombre": "Vacuno (corte magro)",
        "kcal": 215, "prot": 26.0, "grasa": 12.0, "carbo": 0.0,
        "unidad": "100g", "categoria": "Proteínas de tierra"
    },
    "cerdo_lomo": {
        "nombre": "Cerdo (lomo)",
        "kcal": 242, "prot": 27.0, "grasa": 14.0, "carbo": 0.0,
        "unidad": "100g", "categoria": "Proteínas de tierra"
    },
    "pavo_pechuga": {
        "nombre": "Pavo (pechuga)",
        "kcal": 135, "prot": 29.0, "grasa": 1.0, "carbo": 0.0,
        "unidad": "100g", "categoria": "Proteínas de tierra"
    },
    "huevo": {
        "nombre": "Huevo entero",
        "kcal": 72, "prot": 6.0, "grasa": 5.0, "carbo": 0.4,
        "unidad": "unidad", "equiv": 50, "categoria": "Proteínas de tierra"
    },
    "jamon_artesanal": {
        "nombre": "Jamón artesanal",
        "kcal": 145, "prot": 18.0, "grasa": 7.0, "carbo": 1.5,
        "unidad": "100g", "categoria": "Proteínas de tierra"
    },
    # ---- PROTEÍNAS DEL MAR ----
    "salmon": {
        "nombre": "Salmón",
        "kcal": 208, "prot": 20.0, "grasa": 13.0, "carbo": 0.0,
        "unidad": "100g", "categoria": "Proteínas del mar"
    },
    "merluza": {
        "nombre": "Merluza",
        "kcal": 85, "prot": 18.0, "grasa": 1.2, "carbo": 0.0,
        "unidad": "100g", "categoria": "Proteínas del mar"
    },
    "atun_agua": {
        "nombre": "Atún en agua",
        "kcal": 116, "prot": 25.5, "grasa": 1.0, "carbo": 0.0,
        "unidad": "100g", "categoria": "Proteínas del mar"
    },
    "reineta": {
        "nombre": "Reineta",
        "kcal": 95, "prot": 19.0, "grasa": 2.0, "carbo": 0.0,
        "unidad": "100g", "categoria": "Proteínas del mar"
    },
    "congrio": {
        "nombre": "Congrio",
        "kcal": 112, "prot": 19.0, "grasa": 3.5, "carbo": 0.0,
        "unidad": "100g", "categoria": "Proteínas del mar"
    },
    "camarones": {
        "nombre": "Camarones",
        "kcal": 99, "prot": 20.0, "grasa": 1.7, "carbo": 0.9,
        "unidad": "100g", "categoria": "Proteínas del mar"
    },
    "sardinas": {
        "nombre": "Sardinas",
        "kcal": 208, "prot": 24.6, "grasa": 11.5, "carbo": 0.0,
        "unidad": "100g", "categoria": "Proteínas del mar"
    },
    # ---- LÁCTEOS Y QUESOS ----
    "queso_philadelphia": {
        "nombre": "Queso Philadelphia",
        "kcal": 342, "prot": 6.0, "grasa": 33.0, "carbo": 4.0,
        "unidad": "100g", "categoria": "Lácteos y quesos"
    },
    "queso_chacra": {
        "nombre": "Queso chacra",
        "kcal": 350, "prot": 22.0, "grasa": 28.0, "carbo": 2.0,
        "unidad": "100g", "categoria": "Lácteos y quesos"
    },
    "queso_mantecoso": {
        "nombre": "Queso mantecoso",
        "kcal": 380, "prot": 20.0, "grasa": 32.0, "carbo": 1.5,
        "unidad": "100g", "categoria": "Lácteos y quesos"
    },
    "quesillo": {
        "nombre": "Quesillo",
        "kcal": 98, "prot": 11.0, "grasa": 5.0, "carbo": 2.0,
        "unidad": "100g", "categoria": "Lácteos y quesos"
    },
    "yogurt_natural": {
        "nombre": "Yogurt natural entero",
        "kcal": 61, "prot": 3.5, "grasa": 3.3, "carbo": 4.7,
        "unidad": "100g", "categoria": "Lácteos y quesos"
    },
    # ---- GRASAS ----
    "palta": {
        "nombre": "Palta",
        "kcal": 160, "prot": 2.0, "grasa": 15.0, "carbo": 9.0,
        "unidad": "100g", "categoria": "Grasas naturales"
    },
    "aceite_oliva": {
        "nombre": "Aceite de oliva",
        "kcal": 884, "prot": 0.0, "grasa": 100.0, "carbo": 0.0,
        "unidad": "100g", "categoria": "Grasas naturales"
    },
    "mantequilla": {
        "nombre": "Mantequilla",
        "kcal": 717, "prot": 0.9, "grasa": 81.0, "carbo": 0.1,
        "unidad": "100g", "categoria": "Grasas naturales"
    },
    "almendras": {
        "nombre": "Almendras",
        "kcal": 579, "prot": 21.0, "grasa": 50.0, "carbo": 22.0,
        "unidad": "100g", "categoria": "Frutos secos"
    },
    "nueces": {
        "nombre": "Nueces",
        "kcal": 654, "prot": 15.0, "grasa": 65.0, "carbo": 14.0,
        "unidad": "100g", "categoria": "Frutos secos"
    },
    # ---- VERDURAS ----
    "espinaca": {
        "nombre": "Espinaca",
        "kcal": 23, "prot": 2.9, "grasa": 0.4, "carbo": 3.6,
        "unidad": "100g", "categoria": "Verduras"
    },
    "lechuga": {
        "nombre": "Lechuga",
        "kcal": 15, "prot": 1.4, "grasa": 0.2, "carbo": 2.9,
        "unidad": "100g", "categoria": "Verduras"
    },
    "tomate": {
        "nombre": "Tomate",
        "kcal": 18, "prot": 0.9, "grasa": 0.2, "carbo": 3.9,
        "unidad": "100g", "categoria": "Verduras"
    },
    "brocoli": {
        "nombre": "Brócoli",
        "kcal": 34, "prot": 2.8, "grasa": 0.4, "carbo": 6.6,
        "unidad": "100g", "categoria": "Verduras"
    },
    "coliflor": {
        "nombre": "Coliflor",
        "kcal": 25, "prot": 1.9, "grasa": 0.3, "carbo": 5.0,
        "unidad": "100g", "categoria": "Verduras"
    },
    "zapallito": {
        "nombre": "Zapallito italiano",
        "kcal": 17, "prot": 1.2, "grasa": 0.3, "carbo": 3.1,
        "unidad": "100g", "categoria": "Verduras"
    },
    "pimiento": {
        "nombre": "Pimiento rojo",
        "kcal": 31, "prot": 1.0, "grasa": 0.3, "carbo": 6.0,
        "unidad": "100g", "categoria": "Verduras"
    },
    "champinones": {
        "nombre": "Champiñones",
        "kcal": 22, "prot": 3.1, "grasa": 0.3, "carbo": 3.3,
        "unidad": "100g", "categoria": "Verduras"
    },
    "esparragos": {
        "nombre": "Espárragos",
        "kcal": 20, "prot": 2.2, "grasa": 0.1, "carbo": 3.9,
        "unidad": "100g", "categoria": "Verduras"
    },
    "pepino": {
        "nombre": "Pepino",
        "kcal": 16, "prot": 0.7, "grasa": 0.1, "carbo": 3.6,
        "unidad": "100g", "categoria": "Verduras"
    },
    "cebolla": {
        "nombre": "Cebolla",
        "kcal": 40, "prot": 1.1, "grasa": 0.1, "carbo": 9.3,
        "unidad": "100g", "categoria": "Verduras"
    },
    "apio": {
        "nombre": "Apio",
        "kcal": 16, "prot": 0.7, "grasa": 0.2, "carbo": 3.0,
        "unidad": "100g", "categoria": "Verduras"
    },
    "rucula": {
        "nombre": "Rúcula",
        "kcal": 25, "prot": 2.6, "grasa": 0.7, "carbo": 3.7,
        "unidad": "100g", "categoria": "Verduras"
    },
    "berenjena": {
        "nombre": "Berenjena",
        "kcal": 25, "prot": 1.0, "grasa": 0.2, "carbo": 5.9,
        "unidad": "100g", "categoria": "Verduras"
    },
    # ---- FRUTAS ----
    "limon": {
        "nombre": "Limón",
        "kcal": 29, "prot": 1.1, "grasa": 0.3, "carbo": 9.3,
        "unidad": "100g", "categoria": "Frutas"
    },
    "frutilla": {
        "nombre": "Frutilla",
        "kcal": 32, "prot": 0.7, "grasa": 0.3, "carbo": 7.7,
        "unidad": "100g", "categoria": "Frutas"
    },
    "arandano": {
        "nombre": "Arándano",
        "kcal": 57, "prot": 0.7, "grasa": 0.3, "carbo": 14.5,
        "unidad": "100g", "categoria": "Frutas"
    },
    "frambuesa": {
        "nombre": "Frambuesa",
        "kcal": 52, "prot": 1.2, "grasa": 0.7, "carbo": 11.9,
        "unidad": "100g", "categoria": "Frutas"
    },
}


# ============================================================
# MENÚ BASE DE 7 DÍAS
# Cada comida tiene: nombre, ingredientes principales con key y gramos base
# Los gramos base son para 100g de proteína/día — se escalan según el plan
# ============================================================

MENU_BASE = [
    {
        "dia": "Lunes",
        "desayuno": {
            "nombre": "Omelette mediterráneo",
            "proteina_key": "huevo",
            "proteina_unidades": 3,  # unidades, no gramos
            "acompanamiento": "champiñones salteados, queso philadelphia",
            "verduras": ["champinones"],
            "grasa_extra": ["aceite_oliva"],
        },
        "almuerzo": {
            "nombre": "Salmón al horno con espárragos",
            "proteina_key": "salmon",
            "verduras": ["esparragos", "tomate"],
            "grasa_extra": ["aceite_oliva"],
        },
        "cena": {
            "nombre": "Pollo salteado con pimientos",
            "proteina_key": "pollo_pechuga",
            "verduras": ["pimiento", "zapallito", "cebolla"],
            "grasa_extra": ["aceite_oliva"],
        },
    },
    {
        "dia": "Martes",
        "desayuno": {
            "nombre": "Huevos revueltos con jamón y queso",
            "proteina_key": "huevo",
            "proteina_unidades": 3,
            "acompanamiento": "jamón artesanal, queso mantecoso",
            "verduras": [],
            "grasa_extra": ["mantequilla"],
        },
        "almuerzo": {
            "nombre": "Pollo al ajillo con coliflor",
            "proteina_key": "pollo_pechuga",
            "verduras": ["coliflor", "tomate"],
            "grasa_extra": ["aceite_oliva"],
        },
        "cena": {
            "nombre": "Pavo plancha con ensalada",
            "proteina_key": "pavo_pechuga",
            "verduras": ["rucula", "tomate", "cebolla"],
            "grasa_extra": ["aceite_oliva"],
        },
    },
    {
        "dia": "Miércoles",
        "desayuno": {
            "nombre": "Bowl de palta y huevo pochado",
            "proteina_key": "huevo",
            "proteina_unidades": 3,
            "acompanamiento": "palta, tomate cherry",
            "verduras": ["tomate"],
            "grasa_extra": ["palta"],
        },
        "almuerzo": {
            "nombre": "Carne mechada con ensalada",
            "proteina_key": "vacuno_magro",
            "verduras": ["tomate", "cebolla", "pimiento"],
            "grasa_extra": ["aceite_oliva", "palta"],
        },
        "cena": {
            "nombre": "Tímbal de salmón con palta",
            "proteina_key": "salmon",
            "verduras": ["pimiento", "cebolla"],
            "grasa_extra": ["palta", "aceite_oliva"],
        },
    },
    {
        "dia": "Jueves",
        "desayuno": {
            "nombre": "Tortilla de champiñones y queso",
            "proteina_key": "huevo",
            "proteina_unidades": 3,
            "acompanamiento": "champiñones, queso philadelphia",
            "verduras": ["champinones"],
            "grasa_extra": ["aceite_oliva"],
        },
        "almuerzo": {
            "nombre": "Merluza gratinada con espárragos",
            "proteina_key": "merluza",
            "verduras": ["esparragos", "tomate"],
            "grasa_extra": ["mantequilla", "aceite_oliva"],
        },
        "cena": {
            "nombre": "Pollo al horno con pimientos",
            "proteina_key": "pollo_pechuga",
            "verduras": ["pimiento", "zapallito", "tomate"],
            "grasa_extra": ["aceite_oliva"],
        },
    },
    {
        "dia": "Viernes",
        "desayuno": {
            "nombre": "Queso chacra con palta y huevo duro",
            "proteina_key": "huevo",
            "proteina_unidades": 2,
            "acompanamiento": "queso chacra, palta",
            "verduras": ["tomate"],
            "grasa_extra": ["palta"],
        },
        "almuerzo": {
            "nombre": "Vacuno al horno con zapallito gratinado",
            "proteina_key": "vacuno_magro",
            "verduras": ["zapallito", "cebolla"],
            "grasa_extra": ["mantequilla", "aceite_oliva"],
        },
        "cena": {
            "nombre": "Atún con ensalada de tomate y palta",
            "proteina_key": "atun_agua",
            "verduras": ["tomate", "cebolla"],
            "grasa_extra": ["palta", "aceite_oliva"],
        },
    },
    {
        "dia": "Sábado",
        "desayuno": {
            "nombre": "Omelette con pollo desmenuzado",
            "proteina_key": "huevo",
            "proteina_unidades": 2,
            "acompanamiento": "pollo desmenuzado, palta, mayonesa",
            "verduras": [],
            "grasa_extra": ["palta"],
        },
        "almuerzo": {
            "nombre": "Reineta al sartén con vegetales",
            "proteina_key": "reineta",
            "verduras": ["zapallito", "pimiento", "cebolla"],
            "grasa_extra": ["aceite_oliva", "mantequilla"],
        },
        "cena": {
            "nombre": "Tomates rellenos con quesillo y atún",
            "proteina_key": "atun_agua",
            "verduras": ["tomate"],
            "grasa_extra": ["aceite_oliva", "quesillo"],
        },
    },
    {
        "dia": "Domingo",
        "desayuno": {
            "nombre": "Yogurt natural con berries y almendras",
            "proteina_key": "yogurt_natural",
            "acompanamiento": "berries, almendras, chía",
            "verduras": [],
            "grasa_extra": ["almendras"],
        },
        "almuerzo": {
            "nombre": "Pollo al curry con zapallito",
            "proteina_key": "pollo_pechuga",
            "verduras": ["zapallito", "cebolla"],
            "grasa_extra": ["aceite_oliva"],
        },
        "cena": {
            "nombre": "Rolls de salmón en lámina de palta",
            "proteina_key": "salmon",
            "verduras": ["pepino"],
            "grasa_extra": ["palta"],
        },
    },
]


# ============================================================
# MOTOR DE CÁLCULO
# ============================================================

def calcular_gramos_proteina(prot_key: str, prot_objetivo_g: float) -> dict:
    """
    Dado un alimento y el objetivo de proteína en gramos,
    calcula cuántos gramos (o unidades) se necesitan y los macros resultantes.
    """
    alimento = ALIMENTOS[prot_key]
    prot_por_100g = alimento["prot"]

    if alimento["unidad"] == "unidad":
        prot_por_unidad = prot_por_100g * alimento["equiv"] / 100
        unidades = round(prot_objetivo_g / prot_por_unidad)
        # Tope máximo razonable: 4 huevos en desayuno, 6 en otras comidas
        unidades = max(1, min(unidades, 4))
        gramos_total = unidades * alimento["equiv"]
        factor = gramos_total / 100
        return {
            "cantidad": unidades,
            "unidad": "unidades",
            "gramos": gramos_total,
            "kcal": round(alimento["kcal"] * factor),
            "prot": round(alimento["prot"] * factor, 1),
            "grasa": round(alimento["grasa"] * factor, 1),
            "carbo": round(alimento["carbo"] * factor, 1),
        }
    else:
        gramos = round(prot_objetivo_g / prot_por_100g * 100 / 25) * 25
        gramos = max(50, gramos)
        factor = gramos / 100
        return {
            "cantidad": gramos,
            "unidad": "g",
            "gramos": gramos,
            "kcal": round(alimento["kcal"] * factor),
            "prot": round(alimento["prot"] * factor, 1),
            "grasa": round(alimento["grasa"] * factor, 1),
            "carbo": round(alimento["carbo"] * factor, 1),
        }


def distribuir_proteina_dia(prot_total_g: float) -> tuple:
    """
    Distribuye la proteína del día en 3 comidas.
    Desayuno: 25%, Almuerzo: 40%, Cena: 35%
    """
    desayuno = round(prot_total_g * 0.25)
    almuerzo = round(prot_total_g * 0.40)
    cena = prot_total_g - desayuno - almuerzo
    return desayuno, almuerzo, cena


def aplicar_restricciones(menu: list, restricciones: list) -> list:
    """
    Revisa el menú y reemplaza verduras restringidas por alternativas.
    """
    if not restricciones:
        return menu

    rest_lower = [r.lower() for r in restricciones]
    VERDURAS_ALTERNATIVAS = ["tomate", "pimiento", "zapallito", "champinones", "cebolla"]

    menu_limpio = []
    for dia in menu:
        dia_limpio = dict(dia)
        for comida in ["desayuno", "almuerzo", "cena"]:
            c = dict(dia[comida])
            verduras_ok = []
            for v in c.get("verduras", []):
                nombre_v = ALIMENTOS.get(v, {}).get("nombre", v).lower()
                restringida = any(r in nombre_v or nombre_v in r for r in rest_lower)
                if not restringida and v not in verduras_ok:
                    verduras_ok.append(v)
                else:
                    # Buscar alternativa que no esté ya en la lista
                    for alt in VERDURAS_ALTERNATIVAS:
                        alt_nombre = ALIMENTOS.get(alt, {}).get("nombre", alt).lower()
                        no_restringida = not any(r in alt_nombre for r in rest_lower)
                        if alt not in verduras_ok and no_restringida:
                            verduras_ok.append(alt)
                            break
            c["verduras"] = verduras_ok
            dia_limpio[comida] = c
        menu_limpio.append(dia_limpio)
    return menu_limpio


def calcular_menu_completo(datos: dict) -> list:
    """
    Genera el menú de 7 días con macros calculados para cada comida.
    """
    prot_total = float(datos.get("proteinas", 120))
    carbo_total = float(datos.get("carbos", 100))
    grasa_total = float(datos.get("grasas", 75))
    kcal_total = float(datos.get("kcal", 1580))
    restricciones = datos.get("restricciones", [])

    # Distribuir proteína del día
    prot_desayuno, prot_almuerzo, prot_cena = distribuir_proteina_dia(prot_total)

    # Aplicar restricciones al menú base
    menu = aplicar_restricciones(MENU_BASE, restricciones)

    resultado = []
    for dia in menu:
        dia_calculado = {"dia": dia["dia"], "comidas": []}

        total_dia = {"kcal": 0, "prot": 0.0, "grasa": 0.0, "carbo": 0.0}

        for comida_tipo, prot_obj, label in [
            ("desayuno", prot_desayuno, "Desayuno"),
            ("almuerzo", prot_almuerzo, "Almuerzo"),
            ("cena", prot_cena, "Cena"),
        ]:
            c = dia[comida_tipo]
            prot_key = c["proteina_key"]

            # Calcular proteína
            macros = calcular_gramos_proteina(prot_key, prot_obj)

            # Construir descripción del plato
            alimento = ALIMENTOS[prot_key]
            if alimento["unidad"] == "unidad":
                cantidad_str = f"{macros['cantidad']} {macros['unidad']}"
            else:
                cantidad_str = f"{macros['cantidad']}g"

            # Verduras del plato
            verduras_nombres = [
                ALIMENTOS[v]["nombre"] for v in c.get("verduras", []) if v in ALIMENTOS
            ]
            verduras_str = ", ".join(verduras_nombres) if verduras_nombres else ""

            # Acompañamiento extra
            acomp = c.get("acompanamiento", "")

            # Descripción completa
            desc_parts = [f"{cantidad_str} de {alimento['nombre'].lower()}"]
            if acomp:
                desc_parts.append(acomp)
            if verduras_str:
                desc_parts.append(verduras_str)
            descripcion = " · ".join(desc_parts)

            # Totales acumulados del día
            total_dia["kcal"] += macros["kcal"]
            total_dia["prot"] += macros["prot"]
            total_dia["grasa"] += macros["grasa"]
            total_dia["carbo"] += macros["carbo"]

            dia_calculado["comidas"].append({
                "tipo": label,
                "nombre": c["nombre"],
                "descripcion": descripcion,
                "macros": macros,
            })

        dia_calculado["total"] = {
            "kcal": round(kcal_total),
            "prot": round(prot_total, 1),
            "grasa": round(grasa_total, 1),
            "carbo": round(carbo_total, 1),
        }

        resultado.append(dia_calculado)

    return resultado
