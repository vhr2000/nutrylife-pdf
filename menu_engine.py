"""
Nutrylife — menu_engine.py
Base de datos real de platos de Mimi + motor de cálculo de gramajes.
Fuente: 20250515_Platos_Mimi_v5.xlsx
"""

DESAYUNOS = {
    "D001": {
        "nombre": "Huevos revueltos con jamón y queso",
        "restricciones": ["huevo", "lacteo"],
        "imagen": "D001.jpg",
        "liviana":   "2 huevos, 1 lámina de jamón, queso mantecoso, ciboulette",
        "estandar":  "3 huevos, 2 láminas de jamón, 1 lámina de queso mantecoso, ciboulette",
        "completa":  "4 huevos, 2 láminas de jamón, 1 lámina de queso mantecoso, ciboulette",
    },
    "D002": {
        "nombre": "Queso chacra con palta y huevo duro",
        "restricciones": ["lacteo"],
        "imagen": "D002.jpg",
        "liviana":   "50g queso chacra, 1 huevo duro, 1/2 palta, 1 rebanada pan low carb",
        "estandar":  "100g queso chacra, 2 huevos duros, 1/2 palta, 2 rebanadas pan low carb",
        "completa":  "150g queso chacra, 3 huevos duros, 1/2 palta, 2 rebanadas pan low carb",
    },
    "D003": {
        "nombre": "Jamón artesanal mix",
        "restricciones": ["huevo"],
        "imagen": "D003.jpg",
        "liviana":   "2 láminas jamón, 1 huevo revuelto, mayonesa, 1/2 palta, espinaca baby",
        "estandar":  "4 láminas jamón, 2 huevos revueltos, 2 cucharadas mayonesa, 1 palta, espinaca baby",
        "completa":  "5 láminas jamón, 3 huevos revueltos, 2 cucharadas mayonesa, 1 palta, espinaca baby",
    },
    "D004": {
        "nombre": "Tortilla de espinaca rellena de jamón y queso philadelphia",
        "restricciones": ["huevo"],
        "imagen": "D004.jpg",
        "liviana":   "2 huevos, champiñones salteados, jamón, queso philadelphia",
        "estandar":  "3 huevos, champiñones salteados, jamón, queso philadelphia, 1 rebanada pan low carb",
        "completa":  "4 huevos, champiñones salteados, jamón, queso philadelphia, 2 rebanadas pan low carb",
    },
    "D005": {
        "nombre": "Huevos a la copa",
        "restricciones": ["huevo"],
        "imagen": "D005.jpg",
        "liviana":   "2 huevos, 3 cucharadas aceite oliva, sal y pimienta a gusto",
        "estandar":  "3 huevos, 3 cucharadas aceite oliva, sal y pimienta a gusto",
        "completa":  "4 huevos, 3 cucharadas aceite oliva, sal y pimienta a gusto",
    },
    "D006": {
        "nombre": "Omelette con pollo desmenuzado y palta",
        "restricciones": ["huevo"],
        "imagen": "D006.jpg",
        "liviana":   "2 huevos, 50g pollo desmenuzado, 1/2 palta, 2 cucharadas mayonesa, sal y pimienta",
        "estandar":  "2 huevos, 100g pollo desmenuzado, 1 palta, 3 cucharadas mayonesa, sal y pimienta",
        "completa":  "3 huevos, 150g pollo desmenuzado, 1 palta, 3 cucharadas mayonesa, sal y pimienta",
    },
    "D007": {
        "nombre": "Quesillo con tomate y jamón",
        "restricciones": ["lacteo"],
        "imagen": "D007.jpg",
        "liviana":   "100g queso chacra, 2 láminas de jamón, 5 tomates cherry, 20cc aceite oliva, orégano",
        "estandar":  "200g queso chacra, 3 láminas de jamón, 10 tomates cherry, 30cc aceite oliva, orégano, albahaca",
        "completa":  "250g queso chacra, 4 láminas de jamón, 10 tomates cherry, 30cc aceite oliva, orégano, albahaca",
    },
    "D008": {
        "nombre": "Salmón ahumado con palta y aceitunas",
        "restricciones": ["pescado"],
        "imagen": "D008.jpg",
        "liviana":   "100g salmón ahumado, 30g palta, 4 aceitunas",
        "estandar":  "150g salmón ahumado, 1/2 palta, 6 aceitunas",
        "completa":  "200g salmón ahumado, 1 palta, 8 aceitunas",
    },
    "D009": {
        "nombre": "Huevos revueltos con tocino y palta",
        "restricciones": ["huevo"],
        "imagen": "D009.jpg",
        "liviana":   "2 huevos, 2 láminas de tocino, 1/3 palta",
        "estandar":  "3 huevos, 3 láminas de tocino, 1/2 palta",
        "completa":  "4 huevos, 4 láminas de jamón, 1 palta",
    },
    "D010": {
        "nombre": "Salmón y huevo pochado",
        "restricciones": ["pescado"],
        "imagen": "D010.jpg",
        "liviana":   "50g salmón ahumado, 1 huevo pochado",
        "estandar":  "50g salmón ahumado, 2 huevos pochados",
        "completa":  "100g salmón ahumado, 2 huevos pochados",
    },
    "D011": {
        "nombre": "Champiñones rellenos con pollo y queso",
        "restricciones": ["lacteo"],
        "imagen": "D011.jpg",
        "liviana":   "4 champiñones, 50g pollo desmenuzado, 50g queso parmesano",
        "estandar":  "6 champiñones, 100g pollo desmenuzado, 50g queso parmesano",
        "completa":  "8 champiñones, 150g pollo desmenuzado, 80g queso parmesano",
    },
    "D012": {
        "nombre": "Taco de lechuga relleno con carne, queso y palta",
        "restricciones": ["lacteo"],
        "imagen": "D012.jpg",
        "liviana":   "1 hoja de lechuga, 80g carne, 1 lámina de queso mantecoso, 1 cucharada de palta",
        "estandar":  "1 hoja de lechuga, 100g carne, 2 láminas de queso mantecoso, 2 cucharadas de palta",
        "completa":  "1 hoja de lechuga, 150g carne, 2 láminas de queso mantecoso, 3 cucharadas de palta",
    },
    "D013": {
        "nombre": "Hamburguesa casera con tomate y palta",
        "restricciones": ["vacuno"],
        "imagen": "D013.jpg",
        "liviana":   "100g hamburguesa vacuno, 2 rebanadas tomate, 1/3 palta",
        "estandar":  "150g hamburguesa vacuno, 3 rebanadas tomate, 1/2 palta",
        "completa":  "200g hamburguesa vacuno, 4 rebanadas tomate, 1 palta",
    },
    "D014": {
        "nombre": "Rollitos de jamón y queso",
        "restricciones": ["lacteo"],
        "imagen": "D014.jpg",
        "liviana":   "3 láminas de jamón pavo cocido, 2 láminas queso mantecoso",
        "estandar":  "4 láminas de jamón pavo cocido, 3 láminas queso mantecoso",
        "completa":  "5 láminas de jamón pavo cocido, 4 láminas queso mantecoso",
    },
    "D015": {
        "nombre": "Tomate, queso mozzarella y pollo",
        "restricciones": ["lacteo"],
        "imagen": "D015.jpg",
        "liviana":   "1 tomate cortado, 50g pollo desmenuzado, 30g queso mozzarella",
        "estandar":  "1 tomate cortado, 100g pollo desmenuzado, 50g queso mozzarella",
        "completa":  "1 tomate cortado, 150g pollo desmenuzado, 70g queso mozzarella",
    },
    "D016": {
        "nombre": "Muffins de champiñones y jamón",
        "restricciones": ["huevo"],
        "imagen": "D016.jpg",
        "liviana":   "2 huevos, 1 lámina de jamón, 2 champiñones",
        "estandar":  "3 huevos, 2 láminas de jamón, 3 champiñones",
        "completa":  "4 huevos, 2 láminas de jamón, 4 champiñones",
    },
    "D017": {
        "nombre": "Pudin de chia cremoso con whey y frutos rojos",
        "restricciones": ["frutos secos"],
        "imagen": "D017.jpg",
        "liviana":   "Leche de almendras 240cc, 1/2 scoop whey, 1 cucharada chia",
        "estandar":  "Leche de almendras 240cc, 1 scoop whey, 1 cucharada chia, 50g frambuesas",
        "completa":  "Leche de almendras 240cc, 1 1/2 scoop whey, 1 cucharada chia, 50g frambuesas",
    },
    "D018": {
        "nombre": "Yogurt griego con whey protein",
        "restricciones": ["lacteo"],
        "imagen": "D018.jpg",
        "liviana":   "200cc yogurt griego sin azucar, 1/2 scoop proteinas, 50g frambuesas",
        "estandar":  "200cc yogurt griego sin azucar, 1 scoop proteinas",
        "completa":  "200cc yogurt griego sin azucar, 1 1/2 scoop proteinas",
    },
    "D019": {
        "nombre": "Atun con palta",
        "restricciones": ["pescado"],
        "imagen": None,
        "liviana":   "1 lata de atun al agua, 1/3 de palta",
        "estandar":  "1 1/2 latas de atun al agua, 1/2 palta",
        "completa":  "2 latas de atun al agua, 1 palta",
    },
    "D020": {
        "nombre": "Queso cabra con tomate y albahaca",
        "restricciones": ["lacteo"],
        "imagen": "D020.jpg",
        "liviana":   "100g queso cabra, 5 tomates cherry, 5 hojas albahaca, 20cc aceite oliva",
        "estandar":  "150g queso cabra, 8 tomates cherry, 5 hojas albahaca, 25cc aceite oliva",
        "completa":  "200g queso cabra, 10 tomates cherry, 10 hojas albahaca, 30cc aceite oliva",
    },
    "D021": {
        "nombre": "Salmon ahumado con palta, huevo, aceitunas, tomate y pepino",
        "restricciones": ["pescado", "huevo"],
        "imagen": "D021.jpg",
        "liviana":   "50g salmon ahumado, 1 huevo duro, 1/3 palta, 1/3 tomate, 1/3 pepino, 3 aceitunas",
        "estandar":  "80g salmon ahumado, 2 huevos duros, 1/3 palta, 1/3 tomate, 1/3 pepino, 4 aceitunas",
        "completa":  "100g salmon ahumado, 3 huevos duros, 1/3 palta, 1/3 tomate, 1/3 pepino, 5 aceitunas",
    },
}

ALMUERZOS_CENAS = {
    "AC001": {
        "nombre": "Carne mechada con ensalada de apio y palta",
        "tipo": "almuerzo",
        "proteina_nombre": "Vacuno",
        "proteina_g_100g": 26,
        "ingredientes_fijos": "Apio picado, 1/2 palta, aceite oliva, limon, sal y pimienta a gusto",
        "restricciones": [],
        "imagen": "AC001.jpg",
    },
    "AC002": {
        "nombre": "Pollo con arroz y ensalada de tomate y espinaca",
        "tipo": "almuerzo",
        "proteina_nombre": "Pollo",
        "proteina_g_100g": 23,
        "ingredientes_fijos": "3/4 taza arroz, 1 taza lechuga, 1/2 tomate, 1/2 espinaca, 2 cucharadas aceite oliva, vinagreta",
        "restricciones": [],
        "imagen": "AC002.jpg",
    },
    "AC003": {
        "nombre": "Salmon a la mantequilla con papas salteadas",
        "tipo": "almuerzo",
        "proteina_nombre": "Salmon",
        "proteina_g_100g": 17,
        "ingredientes_fijos": "1 papa salteada en mantequilla, 5 tomates cherry, 5 hojas albahaca, 30cc aceite oliva, sal y pimienta",
        "restricciones": [],
        "imagen": "AC003.jpg",
    },
    "AC004": {
        "nombre": "Malaya con queso y ensalada de brocoli",
        "tipo": "almuerzo",
        "proteina_nombre": "Vacuno",
        "proteina_g_100g": 26,
        "ingredientes_fijos": "100g queso mozzarella, 1/2 taza brocoli, limon, aceite oliva, sal y pimienta",
        "restricciones": ["lacteo"],
        "imagen": "AC004.jpg",
    },
    "AC005": {
        "nombre": "Merluza a la mantequilla con papa cocida",
        "tipo": "almuerzo",
        "proteina_nombre": "Merluza",
        "proteina_g_100g": 14,
        "ingredientes_fijos": "100cc crema de leche, 100g queso mozzarella, 1 papa cocida, ensalada mix, sal y pimienta",
        "restricciones": ["lacteo"],
        "imagen": "AC005.jpg",
    },
    "AC006": {
        "nombre": "Pollo al curry",
        "tipo": "almuerzo",
        "proteina_nombre": "Pollo",
        "proteina_g_100g": 23,
        "ingredientes_fijos": "100g crema de leche o coco, 20g curry en pasta, 20cc aceite oliva, mix de hojas verdes, sal, limon",
        "restricciones": [],
        "imagen": "AC006.jpg",
    },
    "AC007": {
        "nombre": "Pollo salteado con verduras",
        "tipo": "cena",
        "proteina_nombre": "Pollo",
        "proteina_g_100g": 23,
        "ingredientes_fijos": "Lechuga, 1/2 tomate, cebolla morada, 30cc aceite oliva, sal y pimienta a gusto",
        "restricciones": [],
        "imagen": "AC007.jpg",
    },
    "AC008": {
        "nombre": "Timbal de salmon con palta",
        "tipo": "cena",
        "proteina_nombre": "Salmon",
        "proteina_g_100g": 17,
        "ingredientes_fijos": "1 pimenton, 1/2 cebolla, 20cc aceite oliva, perejil, sal y pimienta a gusto",
        "restricciones": [],
        "imagen": "AC008.jpg",
    },
    "AC009": {
        "nombre": "Pavo plancha con ensalada",
        "tipo": "cena",
        "proteina_nombre": "Pavo",
        "proteina_g_100g": 24,
        "ingredientes_fijos": "1/4 cebolla picada, rucula, berros, tomate cherry, 20cc aceite oliva, 4 aceitunas, sal y pimienta",
        "restricciones": [],
        "imagen": "AC009.jpg",
    },
    "AC010": {
        "nombre": "Pollo al horno con ensalada",
        "tipo": "cena",
        "proteina_nombre": "Pollo",
        "proteina_g_100g": 23,
        "ingredientes_fijos": "100g pepino, tomate cherry, lechuga, 20cc aceite oliva, cilantro, sal y pimienta a gusto",
        "restricciones": [],
        "imagen": "AC010.jpg",
    },
    "AC011": {
        "nombre": "Tomate con quesillo, atun y aceitunas",
        "tipo": "cena",
        "proteina_nombre": "Atun en agua",
        "proteina_g_100g": 25,
        "ingredientes_fijos": "Tomate cortado en 3, 1 pote de quesillo molido, 20cc aceite oliva, 4 aceitunas, sal y pimienta",
        "restricciones": ["lacteo"],
        "imagen": "AC011.jpg",
    },
    "AC012": {
        "nombre": "Sushi con arroz",
        "tipo": "cena",
        "proteina_nombre": "Salmon",
        "proteina_g_100g": 17,
        "ingredientes_fijos": "8 piezas, arroz para sushi, salsa de soja",
        "restricciones": ["soja", "gluten"],
        "imagen": "AC012.jpg",
    },
    "AC013": {
        "nombre": "Caldo de hueso",
        "tipo": "cena",
        "proteina_nombre": "Osobuco",
        "proteina_g_100g": 18,
        "ingredientes_fijos": "1 litro agua, 2 horas coccion en olla a presion a fuego bajo, sal y pimienta a gusto",
        "restricciones": [],
        "imagen": "AC013.jpg",
    },
    "AC014": {
        "nombre": "Salmon con esparragos gratinados",
        "tipo": "cena",
        "proteina_nombre": "Salmon",
        "proteina_g_100g": 17,
        "ingredientes_fijos": "6 esparragos cocidos salteados, sal y pimienta a gusto",
        "restricciones": [],
        "imagen": "AC014.jpg",
    },
    "AC015": {
        "nombre": "Merluza con pure de coliflor",
        "tipo": "almuerzo",
        "proteina_nombre": "Merluza",
        "proteina_g_100g": 14,
        "ingredientes_fijos": "100g coliflor, 20cc crema de leche, 30g queso parmesano",
        "restricciones": ["lacteo"],
        "imagen": "AC015.jpg",
    },
    "AC016": {
        "nombre": "Carne asada con tortilla de porotos verdes",
        "tipo": "almuerzo",
        "proteina_nombre": "Vacuno",
        "proteina_g_100g": 26,
        "ingredientes_fijos": "2 huevos, 1 taza porotos verdes, sal y pimienta a gusto",
        "restricciones": [],
        "imagen": "AC016.jpg",
    },
    "AC017": {
        "nombre": "Cerdo asado con ensalada chilena",
        "tipo": "almuerzo",
        "proteina_nombre": "Cerdo",
        "proteina_g_100g": 21,
        "ingredientes_fijos": "1 tomate, 1/3 cebolla, sal, 20cc aceite oliva",
        "restricciones": [],
        "imagen": "AC017.jpg",
    },
    "AC018": {
        "nombre": "Ceviche de reineta",
        "tipo": "cena",
        "proteina_nombre": "Reineta",
        "proteina_g_100g": 17,
        "ingredientes_fijos": "30g cebolla morada, 20g pimenton rojo, 10g pimenton verde, 20g cilantro, jugo de 2 limones, sal y pimienta a gusto",
        "restricciones": [],
        "imagen": "AC018.jpg",
    },
    "AC019": {
        "nombre": "Hamburguesas de vacuno con palta, tomate y mayonesa",
        "tipo": "almuerzo",
        "proteina_nombre": "Vacuno",
        "proteina_g_100g": 26,
        "ingredientes_fijos": "1/2 tomate, 1/2 palta, 20cc mayonesa (opcional)",
        "restricciones": [],
        "imagen": "AC019.jpg",
    },
    "AC020": {
        "nombre": "Curry de pollo con verduras salteadas",
        "tipo": "almuerzo",
        "proteina_nombre": "Pollo",
        "proteina_g_100g": 23,
        "ingredientes_fijos": "20g curry en polvo, 20cc aceite de coco, 1 diente ajo, 1 taza brocoli cortado, 30g cebollin, 100cc leche de coco, cilantro, sal y pimienta a gusto",
        "restricciones": [],
        "imagen": None,
    },
    "AC021": {
        "nombre": "Pechuga de pollo rellena de queso y jamon",
        "tipo": "almuerzo",
        "proteina_nombre": "Pollo",
        "proteina_g_100g": 23,
        "ingredientes_fijos": "2 laminas de queso mantecoso, 2 laminas de jamon pavo, 10cc aceite oliva, sal y pimienta a gusto",
        "restricciones": [],
        "imagen": "AC021.jpg",
    },
}


def elegir_version_desayuno(plato, prot_objetivo_g):
    if prot_objetivo_g < 25:
        return plato.get("liviana") or plato["estandar"]
    elif prot_objetivo_g > 35:
        return plato.get("completa") or plato["estandar"]
    else:
        return plato["estandar"]


def calcular_gramaje(proteina_g_100g, prot_objetivo_g):
    gramos = (prot_objetivo_g / proteina_g_100g) * 100
    gramos_redondeado = round(gramos / 25) * 25
    return max(100, min(400, gramos_redondeado))


def distribuir_proteina(prot_total_g):
    prot_desayuno = round(prot_total_g * 0.25)
    prot_almuerzo = round(prot_total_g * 0.40)
    prot_cena = prot_total_g - prot_desayuno - prot_almuerzo
    return prot_desayuno, prot_almuerzo, prot_cena


def filtrar_platos(platos, restricciones, tipo=None):
    rest_lower = [r.lower().strip() for r in restricciones]
    resultado = []
    for pid, plato in platos.items():
        if tipo and plato.get("tipo") != tipo:
            continue
        restricciones_plato = [r.lower() for r in plato.get("restricciones", [])]
        if not any(r in restricciones_plato for r in rest_lower):
            resultado.append(pid)
    return resultado


def seleccionar_menu_7_dias(datos):
    restricciones = datos.get("restricciones", [])
    prot_total = float(datos.get("proteinas", 120))
    prot_desayuno, prot_almuerzo, prot_cena = distribuir_proteina(prot_total)

    desayunos_ok = filtrar_platos(DESAYUNOS, restricciones)
    almuerzos_ok = filtrar_platos(ALMUERZOS_CENAS, restricciones, tipo="almuerzo")
    cenas_ok = filtrar_platos(ALMUERZOS_CENAS, restricciones, tipo="cena")

    dias_semana = ["Lunes", "Martes", "Miercoles", "Jueves", "Viernes", "Sabado", "Domingo"]
    menu = []

    for i, dia in enumerate(dias_semana):
        d_id = desayunos_ok[i % len(desayunos_ok)]
        a_id = almuerzos_ok[i % len(almuerzos_ok)]
        c_id = cenas_ok[i % len(cenas_ok)]

        desayuno = DESAYUNOS[d_id]
        almuerzo = ALMUERZOS_CENAS[a_id]
        cena = ALMUERZOS_CENAS[c_id]

        gramaje_almuerzo = calcular_gramaje(almuerzo["proteina_g_100g"], prot_almuerzo)
        gramaje_cena = calcular_gramaje(cena["proteina_g_100g"], prot_cena)
        ingredientes_desayuno = elegir_version_desayuno(desayuno, prot_desayuno)

        menu.append({
            "dia": dia,
            "desayuno": {
                "id": d_id,
                "nombre": desayuno["nombre"],
                "ingredientes": ingredientes_desayuno,
                "imagen": desayuno.get("imagen"),
            },
            "almuerzo": {
                "id": a_id,
                "nombre": almuerzo["nombre"],
                "proteina_nombre": almuerzo["proteina_nombre"],
                "gramaje": gramaje_almuerzo,
                "ingredientes_fijos": almuerzo["ingredientes_fijos"],
                "imagen": almuerzo.get("imagen"),
            },
            "cena": {
                "id": c_id,
                "nombre": cena["nombre"],
                "proteina_nombre": cena["proteina_nombre"],
                "gramaje": gramaje_cena,
                "ingredientes_fijos": cena["ingredientes_fijos"],
                "imagen": cena.get("imagen"),
            },
        })

    return menu
