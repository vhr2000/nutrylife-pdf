from menu_engine import generar_menu_semanal


datos = {
    "nombre": "Paciente Demo Sin Restricciones",
    "edad": "42",
    "sexo": "femenino",
    "fecha": "2026-05-25",
    "peso": "78",
    "talla": "172",
    "grasa": "30",
    "foco": "Recomposición corporal",
    "estilo_vida": "Moderadamente activo",
    "patologias": "—",

    "kcal": "1800",
    "proteinas": "140",
    "carbos": "130",
    "grasas": "65",

    "version_plan": "con_carbos",
    "suplementos": ["magnesio", "omega3"],
    "restricciones": [],
    "nota_personal": "",

    "prot_distribucion": {
        "desayuno": 30,
        "almuerzo": 45,
        "cena": 25
    },
    "carb_distribucion": {
        "desayuno": 25,
        "almuerzo": 40,
        "cena": 35
    },
    "gras_distribucion": {
        "desayuno": 40,
        "almuerzo": 35,
        "cena": 25
    }
}


resultado = generar_menu_semanal(datos)

print("\n=== MENÚ GENERADO: SIN RESTRICCIONES ===\n")

for dia in resultado["dias"]:
    desayuno = dia["desayuno"]
    almuerzo = dia["almuerzo"]
    cena = dia["cena"]
    totales = dia["totales"]

    print(f'{dia["dia"]}: {desayuno["id"]} + {almuerzo["id"]} + {cena["id"]}')
    print(f'  Desayuno: {desayuno["nombre"]} ({desayuno["version"]})')
    print(f'  Almuerzo: {almuerzo["nombre"]} ({almuerzo["version"]})')
    print(f'  Cena:     {cena["nombre"]} ({cena["version"]})')
    print(
        f'  Totales: P {totales["proteinas"]}g · '
        f'C {totales["carbohidratos"]}g · '
        f'G {totales["grasas"]}g · '
        f'{totales["kcal"]} kcal'
    )
    print(f'  Estado: {dia["estado"]} · Score: {dia["score"]}')
    print()

print("=== RESUMEN ===")
print(resultado["resumen"])
