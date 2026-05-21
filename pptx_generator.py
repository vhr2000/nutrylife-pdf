"""
pptx_generator.py
Genera el plan nutricional en formato PPTX portrait (A4 vertical)
a partir de los datos del formulario web de Myriam Márquez Nutrición.

Llamado desde app.py igual que lo hacía pdf_generator.py.
Requiere: pip install python-pptx requests
"""

import os
import random
import subprocess
import tempfile
import json
from datetime import datetime

# ── python-pptx ──────────────────────────────────────────────
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
from pptx.oxml.ns import qn
from pptx.oxml import parse_xml
from lxml import etree
import copy

# ── Dimensiones A4 portrait ──────────────────────────────────
W = Inches(8.27)   # 210 mm
H = Inches(11.69)  # 297 mm

# ── Paleta ───────────────────────────────────────────────────
VANILLA    = RGBColor(0xFA, 0xF6, 0xEC)
PISTACHO   = RGBColor(0x6B, 0x8E, 0x5A)
PIST_SOFT  = RGBColor(0xEA, 0xF0, 0xE5)
LILA       = RGBColor(0x7A, 0x5F, 0xA0)
LILA_SOFT  = RGBColor(0xED, 0xE8, 0xF5)
OSCURO     = RGBColor(0x3A, 0x3A, 0x3A)
GRIS       = RGBColor(0x8A, 0x8A, 0x8A)
GRIS_SOFT  = RGBColor(0xE8, 0xE4, 0xDA)
BLANCO     = RGBColor(0xFF, 0xFF, 0xFF)
OSCURO_BG  = RGBColor(0x3A, 0x3A, 0x3A)

# ── Fuentes ──────────────────────────────────────────────────
SERIF = "Georgia"
SANS  = "Calibri"

# ── Imágenes de platos (rutas relativas al servidor Render) ──
IMG_BASE = os.path.join(os.path.dirname(__file__), "imagenes")

# ── Base de datos de platos ──────────────────────────────────
DESAYUNOS = [
    {"id": "D001", "nombre": "Omelette mediterráneo",
     "desc": "3 huevos con espinaca, champiñones y queso de cabra. Aceite de oliva, cúrcuma.",
     "prot": 22, "kcal": 320},
    {"id": "D002", "nombre": "Bowl de palta y huevo",
     "desc": "½ palta machacada sobre pan low carb. 2 huevos pochados y semillas de sésamo.",
     "prot": 20, "kcal": 340},
    {"id": "D003", "nombre": "Yogurt con berries",
     "desc": "Yogurt entero sin azúcar, puñado de berries, almendras tostadas y chía.",
     "prot": 15, "kcal": 280},
    {"id": "D004", "nombre": "Huevos revueltos con salmón",
     "desc": "3 huevos revueltos en mantequilla con salmón ahumado, cebollín y tomate cherry.",
     "prot": 28, "kcal": 360, "fijo": True},
    {"id": "D005", "nombre": "Batido proteico",
     "desc": "Proteína en polvo + leche de almendras + frambuesas + mantequilla de almendras.",
     "prot": 25, "kcal": 300},
    {"id": "D006", "nombre": "Plato frío proteico",
     "desc": "150g pollo desmenuzado o salmón ahumado, palta, tomate cherry, aceitunas.",
     "prot": 30, "kcal": 350},
    {"id": "D007", "nombre": "Tostadas de almendra",
     "desc": "Pan de almendra con queso de cabra, palta y 2 huevos duros. Aceite de oliva.",
     "prot": 24, "kcal": 380},
    {"id": "D008", "nombre": "Bowl de semillas",
     "desc": "Yogurt griego con chía, linaza, coco rallado, nueces y berries. Canela.",
     "prot": 18, "kcal": 290},
    {"id": "D009", "nombre": "Huevos con champiñones",
     "desc": "3 huevos al plato con champiñones salteados, espinaca y cúrcuma.",
     "prot": 22, "kcal": 310},
    {"id": "D010", "nombre": "Wrap de lechuga con atún",
     "desc": "Hojas de lechuga rellenas con atún, palta, pepino y salsa de yogurt.",
     "prot": 26, "kcal": 290},
]

ALMUERZOS = [
    {"id": "AC001", "nombre": "Salmón al horno con espárragos",
     "desc": "150g salmón al horno con limón y eneldo. Espárragos asados con almendras.",
     "prot": 35, "kcal": 420},
    {"id": "AC002", "nombre": "Pollo al ajillo con coliflor",
     "desc": "150g pechuga al ajillo. Coliflor asada con cúrcuma. Ensalada de tomate y palta.",
     "prot": 38, "kcal": 400},
    {"id": "AC003", "nombre": "Bowl de huevo y vegetales",
     "desc": "3 huevos cocidos. Mix de hojas verdes, brócoli, champiñones y sésamo.",
     "prot": 20, "kcal": 350},
    {"id": "AC004", "nombre": "Hamburguesas caseras",
     "desc": "Carne molida con cilantro y zapallo italiano. Ensalada de pepino y palta.",
     "prot": 40, "kcal": 480},
    {"id": "AC005", "nombre": "Pescado blanco al wok",
     "desc": "150g reineta al sartén. Mix de pimentón, zapallo y brócoli con jengibre.",
     "prot": 32, "kcal": 380},
    {"id": "AC006", "nombre": "Vacuno con vegetales asados",
     "desc": "150g lomo de vacuno. Vegetales asados: pimentón, zapallo, cebolla. Rúcula.",
     "prot": 38, "kcal": 450},
    {"id": "AC007", "nombre": "Pollo al curry con coliflor",
     "desc": "150g pollo en salsa de curry y leche de coco. Arroz de coliflor y pepino.",
     "prot": 36, "kcal": 420},
    {"id": "AC008", "nombre": "Atún con ensalada fresca",
     "desc": "1 lata atún en aceite de oliva. Mix de hojas, pepino, tomate, aceitunas.",
     "prot": 30, "kcal": 320},
    {"id": "AC009", "nombre": "Pechuga rellena con espinaca",
     "desc": "Pechuga rellena con espinaca y queso de cabra al horno. Vegetales asados.",
     "prot": 40, "kcal": 430},
    {"id": "AC010", "nombre": "Cerdo al limón con brócoli",
     "desc": "150g filete de cerdo al limón y hierbas. Brócoli al vapor con ajo y oliva.",
     "prot": 35, "kcal": 410},
]

CENAS = [
    {"id": "AC011", "nombre": "Salmón al sartén con espinaca",
     "desc": "130g salmón. Espinaca salteada con ajo. Tomatitos cherry. Limón y oliva.",
     "prot": 30, "kcal": 360},
    {"id": "AC012", "nombre": "Crema de zapallo",
     "desc": "Crema de zapallo con caldo de huesos, jengibre y cúrcuma. 2 huevos duros.",
     "prot": 18, "kcal": 290},
    {"id": "AC013", "nombre": "Tortilla de verduras",
     "desc": "3 huevos con pimentón, cebolla, espinaca y queso de cabra. Ensalada verde.",
     "prot": 22, "kcal": 310},
    {"id": "AC014", "nombre": "Pollo desmenuzado con palta",
     "desc": "120g pollo desmenuzado, palta, limón y cilantro sobre hojas verdes.",
     "prot": 32, "kcal": 340},
    {"id": "AC015", "nombre": "Merluza al vapor",
     "desc": "130g merluza al vapor con limón y hierbas. Brócoli y coliflor con ajo.",
     "prot": 28, "kcal": 300},
    {"id": "AC016", "nombre": "Bowl nocturno liviano",
     "desc": "Yogurt griego con pepino, menta y oliva. 2 huevos duros y aceitunas.",
     "prot": 20, "kcal": 260},
    {"id": "AC017", "nombre": "Vacuno liviano con ensalada",
     "desc": "120g filete magro de vacuno. Ensalada de hojas verdes, tomate y palta.",
     "prot": 30, "kcal": 350},
    {"id": "AC018", "nombre": "Caldo de huesos con vegetales",
     "desc": "Caldo casero con verduras de temporada y jengibre. 2 huevos duros.",
     "prot": 16, "kcal": 220},
    {"id": "AC019", "nombre": "Revuelto de claras con vegetales",
     "desc": "5 claras revueltas con espinaca, tomate y queso cottage. Tostada low carb.",
     "prot": 25, "kcal": 270},
    {"id": "AC020", "nombre": "Reineta con ensalada mediterránea",
     "desc": "130g reineta al horno. Ensalada de pepino, aceitunas, tomate y feta.",
     "prot": 28, "kcal": 310},
]

SUPLEMENTOS_INFO = {
    "magnesio":   {"nombre": "Magnesio Citrato",      "dosis": "400mg",    "momento": "Noche, antes de dormir",         "objetivo": "Calidad del sueño y recuperación muscular"},
    "omega3":     {"nombre": "Omega 3",               "dosis": "700mg",    "momento": "Con el almuerzo",                 "objetivo": "Antiinflamatorio y salud cardiovascular"},
    "vitamina_d": {"nombre": "Vitamina D",             "dosis": "5.000 UI", "momento": "Con el desayuno, 2-3 meses",     "objetivo": "Inmunidad, función hormonal y absorción de calcio"},
    "vitamina_c": {"nombre": "Vitamina C Liposomal",   "dosis": "1 cáps.",  "momento": "Con el desayuno",                 "objetivo": "Antioxidante, inmunidad y síntesis de colágeno"},
    "electrolitos":{"nombre": "Electrolitos",          "dosis": "Diario",   "momento": "Agua con limón y sal de mar",     "objetivo": "Hidratación celular y equilibrio de minerales"},
}

DIAS = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]


# ═══════════════════════════════════════════════════════════════
# HELPERS
# ═══════════════════════════════════════════════════════════════

def rgb(r, g, b):
    return RGBColor(r, g, b)

def add_textbox(slide, text, left, top, width, height,
                font_name=SANS, font_size=11, bold=False, italic=False,
                color=None, align=PP_ALIGN.LEFT, word_wrap=True,
                line_spacing=None):
    """Agrega un textbox al slide. Retorna el shape."""
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = word_wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.italic = italic
    if color:
        run.font.color.rgb = color
    if line_spacing:
        from pptx.util import Pt as PPt
        p.line_spacing = PPt(line_spacing)
    return txBox

def add_rect(slide, left, top, width, height, fill_color=None, line_color=None, line_width=Pt(0)):
    """Agrega un rectángulo."""
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        left, top, width, height
    )
    if fill_color:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_color
    else:
        shape.fill.background()
    if line_color:
        shape.line.color.rgb = line_color
        shape.line.width = line_width
    else:
        shape.line.fill.background()
    return shape

def add_line(slide, left, top, width):
    """Línea horizontal."""
    line = slide.shapes.add_shape(1, left, top, width, Pt(0.5))
    line.fill.solid()
    line.fill.fore_color.rgb = GRIS_SOFT
    line.line.fill.background()
    return line

def inches(val):
    return Inches(val)

def pt(val):
    return Pt(val)

def add_image_safe(slide, img_id, left, top, width, height):
    """Agrega imagen si existe, sino pone placeholder de color."""
    path = os.path.join(IMG_BASE, f"{img_id}.jpg")
    if not os.path.exists(path):
        path = os.path.join(IMG_BASE, f"{img_id}.JPG")
    if os.path.exists(path):
        try:
            slide.shapes.add_picture(path, left, top, width, height)
            return True
        except Exception:
            pass
    # Placeholder
    rect = add_rect(slide, left, top, width, height, fill_color=PIST_SOFT)
    add_textbox(slide, "📷", left, top + height//2 - Pt(12),
                width, Pt(24), font_size=16, align=PP_ALIGN.CENTER)
    return False

def footer(slide, seccion, num):
    """Footer estándar."""
    y_footer = H - inches(0.35)
    add_line(slide, inches(0.5), y_footer - Pt(2), W - inches(1.0))
    add_textbox(slide, "MYRIAM NUTRICIÓN", inches(0.5), y_footer,
                inches(3), Pt(14), font_size=7, color=GRIS)
    add_textbox(slide, seccion.upper(), inches(3), y_footer,
                inches(2.5), Pt(14), font_size=7, color=GRIS, align=PP_ALIGN.CENTER)
    add_textbox(slide, str(num), W - inches(1.0), y_footer,
                inches(0.5), Pt(14), font_size=7, color=GRIS, align=PP_ALIGN.RIGHT)

def section_header(slide, numero, titulo, subtitulo=None):
    """Header de sección estilo editorial."""
    add_textbox(slide, str(numero).zfill(2), inches(0.5), inches(0.35),
                inches(1), Pt(18), font_size=11, color=PISTACHO, bold=True)
    add_textbox(slide, titulo, inches(0.5), inches(0.6),
                W - inches(1.0), Pt(55), font_size=36, italic=True,
                font_name=SERIF, color=OSCURO)
    if subtitulo:
        add_textbox(slide, subtitulo, inches(0.5), inches(1.55),
                    W - inches(1.0), Pt(18), font_size=12, italic=True,
                    color=GRIS)
    add_line(slide, inches(0.5), inches(1.9), W - inches(1.0))

def new_slide(prs):
    """Slide en blanco con fondo vainilla."""
    layout = prs.slide_layouts[6]  # blank
    slide = prs.slides.add_slide(layout)
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = VANILLA
    return slide

def seleccionar_platos(lista, cantidad, fijo_key="fijo"):
    resultado = []
    disponibles = list(lista)
    fijo = next((p for p in disponibles if p.get(fijo_key)), None)
    if fijo:
        resultado.append(fijo)
        disponibles = [p for p in disponibles if not p.get(fijo_key)]
    random.shuffle(disponibles)
    while len(resultado) < cantidad and disponibles:
        resultado.append(disponibles.pop(0))
    return resultado[:cantidad]

def fecha_legible(fecha_str):
    try:
        dt = datetime.strptime(fecha_str, "%Y-%m-%d")
        meses = ["enero","febrero","marzo","abril","mayo","junio",
                 "julio","agosto","septiembre","octubre","noviembre","diciembre"]
        return f"{dt.day} de {meses[dt.month-1]} {dt.year}"
    except:
        return fecha_str or ""


# ═══════════════════════════════════════════════════════════════
# FUNCIÓN PRINCIPAL
# ═══════════════════════════════════════════════════════════════

def generar_pptx_completo(datos, output_path):
    """
    Genera el PPTX completo del plan nutricional.
    datos: dict con campos del formulario
    output_path: ruta donde guardar el .pptx
    """

    # ── Parsear datos ────────────────────────────────────────
    nombre     = datos.get("nombre", "Paciente")
    nombre1    = nombre.split()[0]
    edad       = datos.get("edad", "")
    peso       = datos.get("peso", "")
    talla_cm   = float(datos.get("talla", 0) or 0)
    talla      = f"{talla_cm/100:.2f}" if talla_cm else ""
    grasa      = datos.get("grasa", "")
    foco       = datos.get("foco", "")
    duracion   = datos.get("duracion", "14 días")
    fecha      = fecha_legible(datos.get("fecha", ""))
    kcal       = datos.get("kcal", "")
    proteinas  = datos.get("proteinas", "")
    carbos     = datos.get("carbos", "0")
    grasas_g   = datos.get("grasas", "")
    version    = datos.get("version_plan", "con_carbos")
    supls      = datos.get("suplementos", list(SUPLEMENTOS_INFO.keys()))
    nota       = datos.get("nota_personal", "")
    restricciones = datos.get("restricciones", [])

    # IMC
    try:
        imc = round(float(peso) / ((float(talla_cm)/100)**2), 1)
    except:
        imc = ""

    # ── Seleccionar platos ───────────────────────────────────
    desayunos_sel = seleccionar_platos(DESAYUNOS, 7)
    almuerzos_sel = seleccionar_platos(ALMUERZOS, 7)
    cenas_sel     = seleccionar_platos(CENAS, 7)

    # ── Crear presentación ───────────────────────────────────
    prs = Presentation()
    prs.slide_width  = W
    prs.slide_height = H

    pnum = 1  # contador de página

    # ════════════════════════════════════════════════════════
    # SLIDE 1 — PORTADA
    # ════════════════════════════════════════════════════════
    s = new_slide(prs)

    # Bloque verde cubre mitad superior
    COVER_H = inches(7.0)
    add_rect(s, 0, 0, W, COVER_H, fill_color=RGBColor(0x4A, 0x6B, 0x3A))

    # Imagen de portada si existe
    img_portada = os.path.join(IMG_BASE, "portada.jpg")
    if os.path.exists(img_portada):
        try:
            s.shapes.add_picture(img_portada, 0, 0, W, COVER_H)
        except:
            pass

    # Textos header sobre fondo verde
    add_textbox(s, "NUTRICION AVANZADA", inches(0.5), inches(0.38),
                W - inches(1.0), Pt(16), font_size=8.5,
                color=RGBColor(0xD4, 0xE0, 0xCB))
    add_textbox(s, "PLAN PERSONALIZADO", inches(0.5), inches(0.38),
                W - inches(1.0), Pt(16), font_size=8.5,
                color=RGBColor(0xD4, 0xE0, 0xCB), align=PP_ALIGN.RIGHT)

    # Titulo principal
    txb = s.shapes.add_textbox(inches(0.5), inches(2.5), W - inches(1.0), inches(2.5))
    tf = txb.text_frame
    tf.word_wrap = True
    p_title = tf.paragraphs[0]
    p_title.alignment = PP_ALIGN.CENTER
    run = p_title.add_run()
    run.text = "Plan Nutricional"
    run.font.name = SERIF
    run.font.size = Pt(60)
    run.font.italic = True
    run.font.color.rgb = BLANCO

    add_textbox(s, "MYRIAM MARQUEZ GARCIA  -  NUTRICIONISTA",
                inches(0.5), inches(6.2), W - inches(1.0), Pt(18),
                font_size=9, color=RGBColor(0xD4, 0xE0, 0xCB),
                align=PP_ALIGN.CENTER)

    # Bloque inferior datos
    add_line(s, inches(0.5), inches(7.35), W - inches(1.0))

    col_w = (W - inches(1.0)) / 3

    # Col 1: Paciente
    add_textbox(s, "PREPARADO PARA", inches(0.5), inches(7.45),
                col_w, Pt(14), font_size=7, color=GRIS)
    add_textbox(s, nombre, inches(0.5), inches(7.65),
                col_w, Pt(22), font_size=16, font_name=SERIF, color=OSCURO)

    # Col 2: Duracion
    add_textbox(s, "DURACION", inches(0.5) + col_w, inches(7.45),
                col_w, Pt(14), font_size=7, color=GRIS, align=PP_ALIGN.CENTER)
    add_textbox(s, duracion, inches(0.5) + col_w, inches(7.65),
                col_w, Pt(22), font_size=16, font_name=SERIF, color=OSCURO, align=PP_ALIGN.CENTER)

    # Col 3: Fecha
    add_textbox(s, "FECHA DE INICIO", inches(0.5) + col_w * 2, inches(7.45),
                col_w, Pt(14), font_size=7, color=GRIS, align=PP_ALIGN.RIGHT)
    add_textbox(s, fecha, inches(0.5) + col_w * 2, inches(7.65),
                col_w, Pt(22), font_size=16, font_name=SERIF, color=OSCURO, align=PP_ALIGN.RIGHT)

    add_line(s, inches(0.5), inches(8.2), W - inches(1.0))

    # Col 3: Fecha
    add_textbox(s, "FECHA DE INICIO", inches(0.5) + col_w*2, inches(4.18),
                col_w, Pt(14), font_size=7, color=GRIS, align=PP_ALIGN.RIGHT)
    add_textbox(s, fecha, inches(0.5) + col_w*2, inches(4.38),
                col_w, Pt(22), font_size=16, font_name=SERIF, color=OSCURO, align=PP_ALIGN.RIGHT)

    add_line(s, inches(0.5), inches(5.0), W - inches(1.0))

    # ════════════════════════════════════════════════════════
    # SLIDE 2 — ÍNDICE
    # ════════════════════════════════════════════════════════
    s = new_slide(prs)

    add_textbox(s, "01", inches(0.5), inches(0.35), inches(1), Pt(18),
                font_size=11, color=PISTACHO, bold=True)
    add_textbox(s, "Índice", inches(0.5), inches(0.6), W - inches(1.0), Pt(52),
                font_size=38, italic=True, font_name=SERIF, color=OSCURO)
    add_line(s, inches(0.5), inches(1.9), W - inches(1.0))

    secciones = [
        ("02", "Sobre ti"), ("03", "Filosofía del plan"),
        ("04", "Tu transición"), ("05", "Requerimientos y ritmo del día"),
        ("06", "Construye tus comidas"), ("07", f"Desayunos para {nombre1}"),
        ("08", f"Almuerzos para {nombre1}"), ("09", f"Cenas para {nombre1}"),
        ("10", "Menú semanal — Lunes a Domingo"), ("11", "Tablas de raciones"),
        ("12", "Vegetales y frutas"), ("13", "Suplementación"),
        ("14", "Tabla nutricional de proteínas"), ("15", "Indicaciones generales"),
        ("16", "Lista de compras"), ("17", "Entrenamiento"),
        ("18", "Recetas — Anexo"),
    ]

    for i, (num, titulo) in enumerate(secciones):
        y = inches(2.05) + i * inches(0.52)
        add_textbox(s, num, inches(0.5), y, inches(0.5), Pt(16),
                    font_size=9, color=PISTACHO)
        add_textbox(s, titulo, inches(1.1), y, W - inches(2.0), Pt(16),
                    font_size=12, color=OSCURO)
        add_line(s, inches(1.1), y + Pt(16), W - inches(2.0))

    add_textbox(s, "«Comer sin sentir hambre, es el primer paso para tu cambio de vida.»",
                inches(0.5), H - inches(0.8), W - inches(1.0), Pt(18),
                font_size=9, italic=True, color=GRIS, align=PP_ALIGN.CENTER)

    footer(s, "Índice", pnum); pnum += 1

    # ════════════════════════════════════════════════════════
    # SLIDE 3 — BIENVENIDA
    # ════════════════════════════════════════════════════════
    s = new_slide(prs)

    add_textbox(s, "—", inches(0.5), inches(0.35), inches(1), Pt(18),
                font_size=14, color=PISTACHO)
    add_textbox(s, f"Querida {nombre1},", inches(0.5), inches(0.6),
                W - inches(1.0), Pt(52), font_size=34, italic=True,
                font_name=SERIF, color=OSCURO)

    # Foto de Mimi (si existe)
    foto_mimi = os.path.join(IMG_BASE, "mimi.jpg")
    if os.path.exists(foto_mimi):
        try:
            s.shapes.add_picture(foto_mimi, W - inches(2.3), inches(0.5),
                                 inches(1.9), inches(1.9))
        except:
            add_rect(s, W - inches(2.3), inches(0.5), inches(1.9), inches(1.9),
                     fill_color=PIST_SOFT)
    else:
        add_rect(s, W - inches(2.3), inches(0.5), inches(1.9), inches(1.9),
                 fill_color=PIST_SOFT)

    add_textbox(s, "TU NUTRICIONISTA", W - inches(2.3), inches(2.45),
                inches(1.9), Pt(14), font_size=7, color=GRIS, align=PP_ALIGN.CENTER)
    add_textbox(s, "Myriam Márquez García", W - inches(2.4), inches(2.62),
                inches(2.1), Pt(16), font_size=10, color=OSCURO, align=PP_ALIGN.CENTER)

    parrafos = [
        "Este plan fue diseñado pensando en ti: en tu cuerpo, tus tiempos, tus gustos y, sobre todo, en tu objetivo. No es una dieta. Es una propuesta de alimentación pensada para acompañarte y ayudarte a construir hábitos sostenibles, con flexibilidad y sin culpa.",
        "Vas a encontrar acá tres cosas: un marco claro de qué comer y cuánto, un sistema de opciones e intercambios para que armes tus comidas con libertad, y recetas que pueden inspirarte cuando no sepas qué preparar.",
        "Lo más importante: este plan se ajusta a ti, no al revés. Si algo no calza con tu rutina, hablémoslo. Estoy aquí para acompañarte.",
    ]
    for i, p in enumerate(parrafos):
        add_textbox(s, p, inches(0.5), inches(1.8) + i * inches(1.1),
                    W - inches(3.0), Pt(36),
                    font_size=11.5, color=OSCURO, word_wrap=True)

    add_textbox(s, "Con cariño,", inches(0.5), inches(5.1), inches(3), Pt(20),
                font_size=13, italic=True, font_name=SERIF, color=OSCURO)
    add_textbox(s, "Myriam", inches(0.5), inches(5.35), inches(3), Pt(38),
                font_size=30, italic=True, font_name=SERIF, color=PISTACHO)

    footer(s, "Bienvenida", pnum); pnum += 1

    # ════════════════════════════════════════════════════════
    # SLIDE 4 — SOBRE TI
    # ════════════════════════════════════════════════════════
    s = new_slide(prs)
    section_header(s, 2, "Sobre ti", "tu punto de partida en este viaje")

    # Columna izquierda
    add_textbox(s, "DATOS PERSONALES", inches(0.5), inches(2.1),
                inches(3.5), Pt(14), font_size=8, color=GRIS)
    add_line(s, inches(0.5), inches(2.32), inches(3.5))

    datos_personales = [
        ("NOMBRE", nombre),
        ("EDAD", f"{edad} años" if edad else "—"),
        ("FOCO CLÍNICO", foco or "—"),
        ("ESTILO DE VIDA", datos.get("estilo_vida", "—")),
    ]
    for i, (lbl, val) in enumerate(datos_personales):
        y = inches(2.42) + i * inches(0.75)
        add_textbox(s, lbl, inches(0.5), y, inches(3.5), Pt(13),
                    font_size=7.5, color=GRIS)
        add_textbox(s, val, inches(0.5), y + Pt(13), inches(3.5), Pt(22),
                    font_size=13, color=OSCURO, bold=True)

    # Columna derecha — métricas
    add_textbox(s, "EVALUACIÓN INICIAL", inches(4.6), inches(2.1),
                inches(3.5), Pt(14), font_size=8, color=GRIS)
    add_line(s, inches(4.6), inches(2.32), inches(3.5))

    metricas = [
        (str(peso) + " kg", "PESO"),
        (str(talla) + " m", "TALLA"),
        (str(imc), "IMC"),
        (f"{grasa} %", "GRASA CORPORAL"),
    ]
    for i, (val, lbl) in enumerate(metricas):
        col = i % 2
        row = i // 2
        x = inches(4.6) + col * inches(1.9)
        y = inches(2.42) + row * inches(1.1)
        add_textbox(s, val, x, y, inches(1.8), Pt(42),
                    font_size=32, color=OSCURO, font_name=SANS)
        add_textbox(s, lbl, x, y, inches(1.8), Pt(14),
                    font_size=7, color=GRIS)

    # Nota clínica
    add_line(s, inches(0.5), inches(6.2), W - inches(1.0))
    add_textbox(s, "NOTA CLÍNICA", inches(0.5), inches(6.28),
                inches(3), Pt(14), font_size=7.5, color=GRIS)
    nota_clinica = datos.get("nota_personal", "") or f"{nombre} trabaja hacia: {foco}."
    add_textbox(s, nota_clinica, inches(0.5), inches(6.48),
                W - inches(1.0), Pt(50),
                font_size=11, italic=True, color=OSCURO, word_wrap=True)

    footer(s, "Sobre ti", pnum); pnum += 1

    # ════════════════════════════════════════════════════════
    # SLIDE 5 — FILOSOFÍA
    # ════════════════════════════════════════════════════════
    s = new_slide(prs)
    section_header(s, 3, "Filosofía del plan", "los principios que guían tu alimentación")

    filosofia = [
        ("01", "Proteína primero",
         "Eje central de cada comida. Fuente de saciedad, masa magra y reparación. Priorizamos cortes magros, huevos, pescados de calidad."),
        ("02", "Carbohidratos estratégicos",
         "Bajos durante el día, presentes alrededor del entrenamiento. Cambia la fisiología del cuerpo sin renunciar al placer ni a la energía."),
        ("03", "Grasas conscientes",
         "Aceite de oliva, palta, frutos secos. Saciedad real, sabor profundo y soporte hormonal. Sin miedo, con criterio."),
    ]

    for i, (num, titulo, texto) in enumerate(filosofia):
        y = inches(2.1) + i * inches(1.5)
        add_textbox(s, num, inches(0.5), y, inches(0.7), Pt(42),
                    font_size=30, italic=True, font_name=SERIF, color=PISTACHO)
        add_textbox(s, titulo, inches(1.3), y + Pt(6), W - inches(1.8), Pt(22),
                    font_size=15, bold=True, color=OSCURO)
        add_textbox(s, texto, inches(1.3), y + Pt(30), W - inches(1.8), Pt(44),
                    font_size=11.5, color=OSCURO, word_wrap=True)
        if i < 2:
            add_line(s, inches(0.5), y + inches(1.35), W - inches(1.0))

    add_textbox(s, "Este no es un plan rígido. Es un mapa flexible que se ajusta a tu vida.",
                inches(0.5), H - inches(0.85), W - inches(1.0), Pt(24),
                font_size=12, italic=True, color=PISTACHO, font_name=SERIF, align=PP_ALIGN.CENTER)

    footer(s, "Filosofía del plan", pnum); pnum += 1

    # ════════════════════════════════════════════════════════
    # SLIDE 6 — TRANSICIÓN
    # ════════════════════════════════════════════════════════
    s = new_slide(prs)
    section_header(s, 4, "Tu transición", "qué soltar y qué incorporar, sin dramas")

    suelta = [
        ("Azúcares refinados", "y endulzantes artificiales"),
        ("Harinas blancas", "pan, pasta, masas industriales"),
        ("Aceites vegetales", "canola, soya, maíz, girasol"),
        ("Lácteos procesados", "endulzados, light, saborizados"),
        ("Ultraprocesados", "snacks, embutidos industriales"),
        ("Bebidas azucaradas", "gaseosas, jugos, energéticas"),
    ]
    incorpora = [
        ("Proteínas de calidad", "carnes magras, pescados, huevos"),
        ("Vegetales sin almidón", "hojas verdes, crucíferas, raíces"),
        ("Grasas naturales", "oliva, palta, frutos secos, ghee"),
        ("Lácteos enteros", "yogurt natural, quesos artesanales"),
        ("Frutas de estación", "berries y cítricos preferentemente"),
        ("Hidratación", "agua, infusiones, café sin azúcar"),
    ]

    add_textbox(s, "SUELTA", inches(0.5), inches(2.05), inches(3.5), Pt(14),
                font_size=8, color=GRIS)
    add_line(s, inches(0.5), inches(2.27), inches(3.5))
    add_textbox(s, "INCORPORA", inches(4.3), inches(2.05), inches(3.5), Pt(14),
                font_size=8, color=PISTACHO)
    add_line(s, inches(4.3), inches(2.27), inches(3.5))

    for i, ((t1, s1), (t2, s2)) in enumerate(zip(suelta, incorpora)):
        y = inches(2.38) + i * inches(0.83)
        add_textbox(s, t1, inches(0.5), y, inches(3.5), Pt(20),
                    font_size=13, bold=True, color=OSCURO)
        add_textbox(s, s1, inches(0.5), y + Pt(20), inches(3.5), Pt(16),
                    font_size=10, italic=True, color=GRIS)
        add_textbox(s, t2, inches(4.3), y, inches(3.5), Pt(20),
                    font_size=13, bold=True, color=OSCURO)
        add_textbox(s, s2, inches(4.3), y + Pt(20), inches(3.5), Pt(16),
                    font_size=10, italic=True, color=GRIS)

    add_textbox(s, "No se trata de eliminar para siempre, sino de elegir mejor la mayor parte del tiempo.",
                inches(0.5), H - inches(0.85), W - inches(1.0), Pt(20),
                font_size=11, italic=True, color=PISTACHO, font_name=SERIF, align=PP_ALIGN.CENTER)

    footer(s, "Transición alimentaria", pnum); pnum += 1

    # ════════════════════════════════════════════════════════
    # SLIDE 7 — REQUERIMIENTOS + RITMO DEL DÍA
    # ════════════════════════════════════════════════════════
    s = new_slide(prs)
    section_header(s, 5, "Tus requerimientos", "los números y el ritmo que orientan tu día")

    # Macros grandes
    macros_data = [
        (f"{proteinas}g", "PROTEÍNAS", "20% de tu energía"),
        (f"{carbos}g", "CARBOHIDRATOS", "estratégicos, post-entreno"),
        (f"{grasas_g}g", "GRASAS", "fuentes naturales"),
    ]
    col_w3 = (W - inches(1.0)) / 3
    for i, (val, lbl, sub) in enumerate(macros_data):
        x = inches(0.5) + i * col_w3
        add_textbox(s, val, x, inches(2.05), col_w3, Pt(44),
                    font_size=36, color=OSCURO, font_name=SANS)
        add_textbox(s, lbl, x, inches(2.05), col_w3, Pt(14),
                    font_size=7, color=PISTACHO)
        add_textbox(s, sub, x, inches(2.85), col_w3, Pt(16),
                    font_size=10.5, italic=True, color=GRIS)
        if i < 2:
            add_line(s, x + col_w3 - Pt(2), inches(2.05), Pt(0.5))

    # Extras
    add_line(s, inches(0.5), inches(3.4), W - inches(1.0))
    extras = [
        (str(kcal), "CALORÍAS", "kcal/día"),
        ("2.5", "AGUA", "L / día"),
        ("≥ 30", "FIBRA", "g / día"),
    ]
    for i, (val, lbl, sub) in enumerate(extras):
        x = inches(0.5) + i * col_w3
        add_textbox(s, lbl, x, inches(3.5), col_w3, Pt(14),
                    font_size=7.5, color=GRIS, align=PP_ALIGN.CENTER)
        add_textbox(s, val, x, inches(3.68), col_w3, Pt(42),
                    font_size=34, color=OSCURO, align=PP_ALIGN.CENTER)
        add_textbox(s, sub, x, inches(4.45), col_w3, Pt(16),
                    font_size=11, italic=True, color=GRIS, align=PP_ALIGN.CENTER)

    # Ritmo del día
    add_line(s, inches(0.5), inches(5.2), W - inches(1.0))
    add_textbox(s, "TU RITMO DEL DÍA", inches(0.5), inches(5.28),
                inches(3), Pt(14), font_size=8, color=GRIS)

    horarios = [
        ("08:00", "Desayuno", "Proteína sólida, grasas buenas. Dentro de la primera hora después de despertar."),
        ("13:30", "Almuerzo", "El plato más completo del día. Aquí van los carbos si entrenaste en la mañana."),
        ("20:00", "Cena", "Versión más liviana del almuerzo. Proteína, vegetales, grasas. Cerrar cocina 3h antes."),
    ]
    for i, (hora, comida, desc) in enumerate(horarios):
        x = inches(0.5) + i * col_w3
        add_textbox(s, hora, x, inches(5.5), col_w3, Pt(28),
                    font_size=22, italic=True, font_name=SERIF, color=PISTACHO)
        add_textbox(s, comida, x, inches(5.95), col_w3, Pt(20),
                    font_size=14, bold=True, color=OSCURO)
        add_textbox(s, desc, x, inches(6.2), col_w3, Pt(36),
                    font_size=9.5, color=GRIS, word_wrap=True)

    footer(s, "Requerimientos", pnum); pnum += 1

    # ════════════════════════════════════════════════════════
    # SLIDE 8 — CONSTRUYE TUS COMIDAS
    # ════════════════════════════════════════════════════════
    s = new_slide(prs)
    section_header(s, 6, "Construye tus comidas", "el sistema simple para armar tu plato")

    add_textbox(s, "Cada comida sigue una fórmula simple: proteína + vegetales + grasa. Las cantidades varían según el momento del día.",
                inches(0.5), inches(2.05), W - inches(1.0), Pt(32),
                font_size=12, color=OSCURO, word_wrap=True)

    # Diagrama de torta simplificado con rectángulos de colores
    es_carbos = (version == "con_carbos")

    segmentos = [
        (PISTACHO, "50%", "Proteína",
         "Pollo, pescado, huevos, carne magra. La base de saciedad y masa muscular."),
        (LILA, "25%",
         "Carbohidratos" if es_carbos else "Vegetales",
         "Estratégicos post-entreno: arroz integral, camote, legumbres." if es_carbos else "Hojas verdes, crucíferas, zapallo italiano. Sin límite."),
        (RGBColor(0xC9, 0xB6, 0x9E), "25%", "Grasas",
         "Palta, aceite de oliva, frutos secos. Soporte hormonal y saciedad."),
    ]

    # Barras de color como "torta" visual
    bar_total_h = inches(3.5)
    bar_w = inches(1.0)
    bar_x = inches(1.5)
    bar_y = inches(2.8)

    proporciones = [0.5, 0.25, 0.25]
    acumulado = 0
    for i, ((color, pct, label, desc), prop) in enumerate(zip(segmentos, proporciones)):
        seg_h = int(bar_total_h * prop)
        add_rect(s, bar_x, bar_y + int(acumulado), bar_w, seg_h, fill_color=color)
        acumulado += seg_h

    # Leyenda derecha
    for i, (color, pct, label, desc) in enumerate(segmentos):
        y = inches(2.95) + i * inches(1.2)
        add_rect(s, inches(3.0), y, Pt(6), inches(0.7), fill_color=color)
        add_textbox(s, f"{pct} {label}", inches(3.15), y, W - inches(3.6), Pt(20),
                    font_size=15, bold=True, color=OSCURO)
        add_textbox(s, desc, inches(3.15), y + Pt(22), W - inches(3.6), Pt(36),
                    font_size=11, color=GRIS, word_wrap=True)

    footer(s, "Construye tus comidas", pnum); pnum += 1

    # ════════════════════════════════════════════════════════
    # SLIDES 9, 10, 11 — DESAYUNOS / ALMUERZOS / CENAS PARA TI
    # ════════════════════════════════════════════════════════
    for seccion_idx, (titulo, platos, color_acento, sec_num) in enumerate([
        (f"Desayunos para {nombre1}", desayunos_sel, PISTACHO, 7),
        (f"Almuerzos para {nombre1}", almuerzos_sel, PISTACHO, 8),
        (f"Cenas para {nombre1}", cenas_sel, LILA, 9),
    ]):
        s = new_slide(prs)
        section_header(s, sec_num, titulo,
                       "siete opciones seleccionadas para ti" if seccion_idx == 0
                       else "siete platos completos para el mediodía" if seccion_idx == 1
                       else "siete opciones livianas para cerrar el día")

        # Imagen rectangular redondeada (card) al costado
        IMG_W = inches(1.8)
        IMG_H = inches(1.1)
        TEXT_X = inches(0.5)
        TEXT_W = W - inches(1.0) - IMG_W - inches(0.2)

        for i, plato in enumerate(platos):
            y = inches(2.05) + i * inches(1.32)

            # Número
            add_textbox(s, f"0{i+1}", TEXT_X, y, inches(0.4), Pt(26),
                        font_size=22, italic=True, font_name=SERIF, color=color_acento)

            # Imagen rectangular
            img_x = W - IMG_W - inches(0.4)
            add_image_safe(s, plato["id"], img_x, y, IMG_W, IMG_H)

            # Nombre
            add_textbox(s, plato["nombre"], TEXT_X + inches(0.45), y + Pt(4),
                        TEXT_W - inches(0.45), Pt(20),
                        font_size=13, bold=True, color=OSCURO)

            # Macros
            macro_txt = f"{plato['prot']}g prot · {plato['kcal']} kcal"
            add_textbox(s, macro_txt, TEXT_X + inches(0.45), y + Pt(24),
                        TEXT_W - inches(0.45), Pt(14),
                        font_size=9, color=color_acento, italic=True)

            # Descripción
            add_textbox(s, plato["desc"], TEXT_X + inches(0.45), y + Pt(38),
                        TEXT_W - inches(0.45), Pt(28),
                        font_size=10, color=GRIS, word_wrap=True)

            if i < len(platos) - 1:
                add_line(s, TEXT_X, y + Pt(85), W - inches(1.0))

        footer(s, titulo, pnum); pnum += 1

    # ════════════════════════════════════════════════════════
    # SLIDES 12-18 — MENÚ SEMANAL (1 slide por día)
    # ════════════════════════════════════════════════════════
    for idx, dia in enumerate(DIAS):
        s = new_slide(prs)

        # Header con día destacado
        add_rect(s, 0, 0, W, inches(1.5), fill_color=PISTACHO)
        add_textbox(s, "10 · MENÚ SEMANAL", inches(0.5), inches(0.18),
                    W - inches(1.0), Pt(16), font_size=8,
                    color=RGBColor(0xD4,0xE0,0xCB))
        add_textbox(s, dia, inches(0.5), inches(0.38),
                    W - inches(1.0), Pt(62), font_size=48, italic=True,
                    font_name=SERIF, color=BLANCO)

        # Tres comidas
        comidas = [
            ("*", "DESAYUNO  -  08:00", desayunos_sel[idx], PIST_SOFT, PISTACHO),
            ("*", "ALMUERZO  -  13:30", almuerzos_sel[idx], PIST_SOFT, PISTACHO),
            ("*", "CENA  -  20:00", cenas_sel[idx], LILA_SOFT, LILA),
        ]

        IMG_W_M = inches(1.7)
        IMG_H_M = inches(1.05)

        for ci, (emoji, label, plato, bg, acento) in enumerate(comidas):
            y_base = inches(1.65) + ci * inches(3.28)

            # Badge de etiqueta
            add_rect(s, inches(0.5), y_base, inches(2.2), Pt(22), fill_color=bg)
            add_textbox(s, f"{emoji} {label}", inches(0.6), y_base + Pt(3),
                        inches(2.1), Pt(18), font_size=9, color=acento, bold=True)

            # Imagen a la derecha
            img_x = W - IMG_W_M - inches(0.4)
            add_image_safe(s, plato["id"], img_x, y_base + Pt(28), IMG_W_M, IMG_H_M)

            # Nombre plato
            add_textbox(s, plato["nombre"],
                        inches(0.5), y_base + Pt(30),
                        W - inches(1.0) - IMG_W_M - inches(0.3), Pt(26),
                        font_size=17, bold=True, color=OSCURO)

            # Descripción
            add_textbox(s, plato["desc"],
                        inches(0.5), y_base + Pt(58),
                        W - inches(1.0) - IMG_W_M - inches(0.3), Pt(32),
                        font_size=10.5, color=GRIS, word_wrap=True)

            # Macros
            add_textbox(s, f"{plato['prot']}g proteína · {plato['kcal']} kcal",
                        inches(0.5), y_base + Pt(92),
                        inches(3.5), Pt(16),
                        font_size=9.5, color=acento, italic=True)

            if ci < 2:
                add_line(s, inches(0.5), y_base + inches(1.32), W - inches(1.0))

        footer(s, f"Menú — {dia}", pnum); pnum += 1

    # ════════════════════════════════════════════════════════
    # SLIDE — TABLAS DE RACIONES
    # ════════════════════════════════════════════════════════
    s = new_slide(prs)
    section_header(s, 11, "Tablas de raciones", "intercambia con criterio, mantén el equilibrio")

    proteinas_tab = [
        ("Carnes magras (pollo, vacuno, pavo)", "100 g", "1 trozo medio"),
        ("Pescado blanco o salmón", "120 g", "1 filete chico"),
        ("Huevos enteros", "3 unidades", "tamaño grande"),
        ("Queso de cabra u oveja", "60 g", "1 trozo"),
        ("Jamón serrano artesanal", "60 g", "4 lonjas finas"),
        ("Proteína en polvo", "1 medida", "30 g aprox."),
    ]
    grasas_tab = [
        ("Palta", "½ unidad", "30 g aprox."),
        ("Aceite de oliva extra virgen", "1 cda", "5 g"),
        ("Frutos secos (almendras, nueces)", "5-8 unid.", "10 g"),
        ("Aceitunas verdes o negras", "5 grandes", "30 g"),
        ("Semillas (chía, linaza, sésamo)", "1 cda", "10 g"),
    ]

    def tabla_racion(slide, titulo, subtitulo, filas, y_start):
        add_textbox(slide, titulo, inches(0.5), y_start,
                    inches(4), Pt(14), font_size=8, color=PISTACHO)
        add_textbox(slide, subtitulo, W - inches(4.0), y_start,
                    inches(3.5), Pt(14), font_size=9, italic=True,
                    color=GRIS, align=PP_ALIGN.RIGHT)
        add_line(slide, inches(0.5), y_start + Pt(14), W - inches(1.0))

        cols = [inches(3.8), inches(1.8), inches(2.2)]
        headers = ["ALIMENTO", "CANTIDAD", "MEDIDA PRÁCTICA"]

        y = y_start + Pt(18)
        x_pos = [inches(0.5), inches(4.3), inches(6.1)]

        # Header
        for j, (hdr, xp) in enumerate(zip(headers, x_pos)):
            add_textbox(slide, hdr, xp, y, cols[j], Pt(13),
                        font_size=7.5, color=GRIS, bold=True)
        y += Pt(16)
        add_line(slide, inches(0.5), y, W - inches(1.0))
        y += Pt(4)

        for alimento, cantidad, medida in filas:
            add_textbox(slide, alimento, x_pos[0], y, cols[0], Pt(18),
                        font_size=11, color=OSCURO)
            add_textbox(slide, cantidad, x_pos[1], y, cols[1], Pt(18),
                        font_size=11, italic=True, color=GRIS)
            add_textbox(slide, medida, x_pos[2], y, cols[2], Pt(18),
                        font_size=11, italic=True, color=GRIS)
            y += Pt(22)
            add_line(slide, inches(0.5), y, W - inches(1.0))
            y += Pt(4)
        return y

    y_end = tabla_racion(s, "PROTEÍNAS", "una ración aporta ~20g de proteína",
                         proteinas_tab, inches(2.05))
    tabla_racion(s, "GRASAS", "una ración aporta ~10g de grasa",
                 grasas_tab, y_end + inches(0.25))

    footer(s, "Tablas de raciones", pnum); pnum += 1

    # ════════════════════════════════════════════════════════
    # SLIDE — VEGETALES Y FRUTAS
    # ════════════════════════════════════════════════════════
    s = new_slide(prs)
    section_header(s, 12, "Vegetales & frutas", "el color y la fibra de tu plato")

    vegetales = [
        ("Hojas verdes", "espinaca, lechuga, kale, rúcula, acelga"),
        ("Crucíferas", "brócoli, coliflor, repollo, coles de bruselas"),
        ("Zapallo italiano", "berenjenas, calabacín, pimentón"),
        ("Aromáticos", "ajo, cebolla, ciboulette, perejil"),
        ("Otros", "champiñones, espárragos, palmitos"),
    ]
    frutas = [
        ("Frutos rojos", "frambuesa, frutilla, mora, arándano"),
        ("Cítricos", "limón, mandarina, naranja en porción"),
        ("Manzana verde", "1 unidad pequeña"),
        ("Kiwi", "2 unidades pequeñas"),
        ("Ciruela", "1 unidad mediana"),
    ]

    # Dos columnas
    for col_idx, (datos_col, titulo_col, sub_col) in enumerate([
        (vegetales, "VEGETALES", "una ración: 1 taza crudo o ½ cocido"),
        (frutas, "FRUTAS", "máx. 1 ración al día, preferir berries"),
    ]):
        x = inches(0.5) + col_idx * inches(3.9)
        add_textbox(s, titulo_col, x, inches(2.05), inches(3.5), Pt(14),
                    font_size=8, color=GRIS)
        add_textbox(s, sub_col, x, inches(2.22), inches(3.5), Pt(14),
                    font_size=9.5, italic=True, color=GRIS)
        add_line(s, x, inches(2.42), inches(3.5))

        for i, (nombre_v, detalle) in enumerate(datos_col):
            y = inches(2.52) + i * inches(0.82)
            add_textbox(s, nombre_v, x, y, inches(3.5), Pt(20),
                        font_size=13, bold=True, color=OSCURO)
            add_textbox(s, detalle, x, y + Pt(20), inches(3.5), Pt(16),
                        font_size=10, italic=True, color=GRIS)

    # Superalimentos
    add_line(s, inches(0.5), inches(6.8), W - inches(1.0))
    add_textbox(s, "SÚPER ALIMENTOS — INTEGRA SEGÚN GUSTO",
                inches(0.5), inches(6.88), W - inches(1.0), Pt(14),
                font_size=7.5, color=GRIS)
    add_textbox(s, "· Cúrcuma  · Matcha  · Maqui  · Caldo de huesos  · Chucrut  · Jengibre  · Cacao puro",
                inches(0.5), inches(7.05), W - inches(1.0), Pt(18),
                font_size=11, italic=True, color=PISTACHO)

    footer(s, "Vegetales y frutas", pnum); pnum += 1

    # ════════════════════════════════════════════════════════
    # SLIDE — SUPLEMENTACIÓN
    # ════════════════════════════════════════════════════════
    s = new_slide(prs)
    section_header(s, 13, "Suplementación", "apoyo específico para tus objetivos")

    supls_activos = [k for k in supls if k in SUPLEMENTOS_INFO]

    for i, key in enumerate(supls_activos):
        sup = SUPLEMENTOS_INFO[key]
        y = inches(2.05) + i * inches(1.05)

        # Barra lila izquierda
        add_rect(s, inches(0.5), y, Pt(5), inches(0.75), fill_color=LILA)

        add_textbox(s, sup["nombre"], inches(0.75), y + Pt(4),
                    W - inches(1.25), Pt(22), font_size=15, bold=True, color=OSCURO)
        add_textbox(s,
                    f"Dosis: {sup['dosis']}   ·   Momento: {sup['momento']}",
                    inches(0.75), y + Pt(26), W - inches(1.25), Pt(16),
                    font_size=10.5, color=GRIS)
        add_textbox(s, f"→ {sup['objetivo']}", inches(0.75), y + Pt(42),
                    W - inches(1.25), Pt(16), font_size=10, italic=True, color=PISTACHO)
        if i < len(supls_activos) - 1:
            add_line(s, inches(0.75), y + Pt(58), W - inches(1.25))

    footer(s, "Suplementación", pnum); pnum += 1

    # ════════════════════════════════════════════════════════
    # SLIDE — TABLA NUTRICIONAL DE PROTEÍNAS
    # ════════════════════════════════════════════════════════
    s = new_slide(prs)
    section_header(s, 14, "Tabla nutricional", "valores de proteínas en crudo por 100g")

    tabla_prot = [
        ("Pechuga de pollo", "31 g", "3.6 g", "165 kcal"),
        ("Salmón fresco", "25 g", "13 g", "208 kcal"),
        ("Huevo entero", "13 g", "11 g", "155 kcal"),
        ("Vacuno magro (lomo)", "26 g", "7 g", "170 kcal"),
        ("Reineta / Merluza", "20 g", "2 g", "100 kcal"),
        ("Atún en agua", "29 g", "1 g", "128 kcal"),
        ("Pavo (pechuga)", "29 g", "1 g", "135 kcal"),
        ("Queso de cabra", "21 g", "20 g", "268 kcal"),
        ("Yogurt griego natural", "10 g", "0.7 g", "59 kcal"),
        ("Proteína en polvo (whey)", "75-80 g", "4-6 g", "370 kcal"),
        ("Claras de huevo", "11 g", "0.2 g", "52 kcal"),
        ("Carne molida magra", "20 g", "10 g", "170 kcal"),
    ]
    headers_tab = ["ALIMENTO", "PROTEÍNA", "GRASA", "CALORÍAS"]
    col_widths = [inches(3.2), inches(1.4), inches(1.4), inches(1.6)]
    x_cols = [inches(0.5), inches(3.7), inches(5.1), inches(6.5)]

    y = inches(2.05)
    # Header
    add_rect(s, inches(0.5), y, W - inches(1.0), Pt(20), fill_color=PISTACHO)
    for j, (hdr, xp, cw) in enumerate(zip(headers_tab, x_cols, col_widths)):
        add_textbox(s, hdr, xp, y + Pt(4), cw, Pt(14),
                    font_size=8, color=BLANCO, bold=True)
    y += Pt(22)

    for row in tabla_prot:
        bg = PIST_SOFT if tabla_prot.index(row) % 2 == 0 else VANILLA
        add_rect(s, inches(0.5), y, W - inches(1.0), Pt(20), fill_color=bg)
        for j, (cell, xp, cw) in enumerate(zip(row, x_cols, col_widths)):
            clr = PISTACHO if j > 0 else OSCURO
            add_textbox(s, cell, xp, y + Pt(4), cw, Pt(14),
                        font_size=10.5, color=clr)
        y += Pt(21)

    add_textbox(s, "* Valores aproximados en crudo. Pueden variar según corte, preparación y marca.",
                inches(0.5), y + Pt(8), W - inches(1.0), Pt(16),
                font_size=9, italic=True, color=GRIS)

    footer(s, "Tabla nutricional", pnum); pnum += 1

    # ════════════════════════════════════════════════════════
    # SLIDE — INDICACIONES
    # ════════════════════════════════════════════════════════
    s = new_slide(prs)
    section_header(s, 15, "Indicaciones", "lo que hace la diferencia, día a día")

    indicaciones = [
        ("01", "Calidad sobre cantidad", "Prefiere alimentos enteros, mínimamente procesados. La calidad de los ingredientes pesa más que el conteo exacto de gramos."),
        ("02", "Come sin pantallas", "Cuando sea posible, come lejos del computador y del teléfono. La saciedad llega antes cuando estás presente."),
        ("03", "Hidrátate antes de tener sed", "Empieza el día con un vaso de agua. Si te cuesta llegar a 2,5 L, suma infusiones, agua con limón o caldos."),
        ("04", "Mueve el cuerpo", "Tres sesiones de fuerza por semana son suficientes para empezar. Camina entre 7.000 y 10.000 pasos al día."),
        ("05", "Duerme como una prioridad", "Apunta a 7-8 horas. El cortisol mal regulado por mal sueño sabotea cualquier plan de alimentación."),
        ("06", "Sé flexible los fines de semana", "Una comida libre por semana es parte del plan, no una traición. Vuelve a la siguiente comida sin culpa."),
    ]

    col_w2 = (W - inches(1.0)) / 2
    for i, (num, titulo, texto) in enumerate(indicaciones):
        col = i % 2
        row = i // 2
        x = inches(0.5) + col * col_w2
        y = inches(2.1) + row * inches(1.65)

        add_textbox(s, num, x, y, inches(0.5), Pt(30),
                    font_size=24, italic=True, font_name=SERIF, color=PISTACHO)
        add_textbox(s, titulo, x + inches(0.55), y + Pt(5), col_w2 - inches(0.6), Pt(20),
                    font_size=13, bold=True, color=OSCURO)
        add_textbox(s, texto, x + inches(0.55), y + Pt(26), col_w2 - inches(0.6), Pt(42),
                    font_size=10.5, color=GRIS, word_wrap=True)

        if row < 2 and col == 1:
            add_line(s, inches(0.5), y + Pt(70), W - inches(1.0))

    footer(s, "Indicaciones", pnum); pnum += 1

    # ════════════════════════════════════════════════════════
    # SLIDE — LISTA DE COMPRAS
    # ════════════════════════════════════════════════════════
    s = new_slide(prs)
    section_header(s, 16, "Lista de compras", "una semana, organizada en cinco bloques")

    compras = {
        "PROTEÍNAS": ["Pechuga de pollo · 1 kg", "Salmón fresco · 600 g",
                      "Reineta o merluza · 600 g", "Huevos de campo · 2 docenas",
                      "Queso de cabra · 200 g", "Jamón serrano · 100 g"],
        "VEGETALES": ["Espinaca · 1 atado", "Brócoli · 1 unidad grande",
                      "Coliflor · 1 unidad", "Zapallo italiano · 3 unid.",
                      "Pimentón rojo · 2 unid.", "Mix de hojas verdes · 200 g"],
        "FRUTAS": ["Frambuesas · 1 caja", "Arándanos · 1 caja",
                   "Limón · 6 unid.", "Palta · 4 unid.", "Manzana verde · 3 unid."],
        "GRASAS Y SEMILLAS": ["Aceite de oliva AOVE · 1 botella",
                               "Almendras · 250 g", "Nueces · 200 g",
                               "Semillas de chía · 100 g", "Aceitunas · 1 frasco"],
        "DESPENSA": ["Harina de almendras · 250 g", "Cacao puro · 200 g",
                     "Vinagre de manzana · 1 botella", "Sal de mar · 1 frasco",
                     "Café orgánico · 250 g", "Stevia natural · 1 frasco"],
    }

    y_cursor = inches(2.08)
    for cat, items in compras.items():
        add_textbox(s, cat, inches(0.5), y_cursor, W - inches(1.0), Pt(14),
                    font_size=7.5, color=PISTACHO, bold=True)
        y_cursor += Pt(14)
        add_line(s, inches(0.5), y_cursor, W - inches(1.0))
        y_cursor += Pt(5)

        # Items en 2 columnas
        mid = (len(items) + 1) // 2
        col1 = items[:mid]
        col2 = items[mid:]
        col_w_c = (W - inches(1.0)) / 2

        max_rows = max(len(col1), len(col2))
        for r in range(max_rows):
            if r < len(col1):
                add_textbox(s, f"○  {col1[r]}", inches(0.5), y_cursor,
                            col_w_c, Pt(16), font_size=11, color=OSCURO)
            if r < len(col2):
                add_textbox(s, f"○  {col2[r]}", inches(0.5) + col_w_c, y_cursor,
                            col_w_c, Pt(16), font_size=11, color=OSCURO)
            y_cursor += Pt(18)
        y_cursor += Pt(8)

    footer(s, "Lista de compras", pnum); pnum += 1

    # ════════════════════════════════════════════════════════
    # SLIDE — ENTRENAMIENTO
    # ════════════════════════════════════════════════════════
    s = new_slide(prs)

    # Header pistacho
    add_rect(s, 0, 0, W, inches(1.5), fill_color=PISTACHO)
    add_textbox(s, "17 · ENTRENAMIENTO", inches(0.5), inches(0.2),
                W - inches(1.0), Pt(16), font_size=8.5, color=RGBColor(0xD4,0xE0,0xCB))
    add_textbox(s, "Lo que no debe faltar", inches(0.5), inches(0.42),
                W - inches(1.0), Pt(62), font_size=40, italic=True,
                font_name=SERIF, color=BLANCO)

    ejercicios = [
        ("01", "Sentadillas", "3-4 series x 12-15 rep", "Piernas - Gluteos"),
        ("02", "Puente de gluteos", "3-4 series x 15-20 rep", "Gluteos - Femorales"),
        ("03", "Flexiones (rodillas)", "3 series x 8-12 rep", "Pecho - Triceps"),
        ("04", "Remo con botellas", "3 series x 12-15 rep", "Espalda - Biceps"),
        ("05", "Zancadas alternas", "3 series x 10 rep/lado", "Piernas - Equilibrio"),
        ("06", "Plancha abdominal", "3 series x 30-45 seg", "Core - Estabilidad"),
    ]

    CARD_W = (W - inches(1.2)) / 3
    CARD_H = inches(3.5)

    for i, (emoji, nombre_e, series, foco_e) in enumerate(ejercicios):
        col = i % 3
        row = i // 3
        x = inches(0.5) + col * (CARD_W + inches(0.1))
        y = inches(1.65) + row * (CARD_H + inches(0.1))

        # Card blanca
        add_rect(s, x, y, CARD_W, CARD_H,
                 fill_color=BLANCO, line_color=GRIS_SOFT, line_width=Pt(0.5))

        # Numero grande (ancho suficiente para no partir el texto)
        add_textbox(s, emoji, x + Pt(8), y + Pt(8), inches(0.6), inches(0.45), font_size=26,
                    bold=True, color=PISTACHO, font_name=SERIF)

        # Nombre ejercicio
        add_textbox(s, nombre_e, x + Pt(8), y + Pt(48), CARD_W - Pt(16), Pt(24),
                    font_size=12, bold=True, color=OSCURO, word_wrap=True)

        # Badge series
        add_rect(s, x + Pt(8), y + Pt(76), CARD_W - Pt(16), Pt(20),
                 fill_color=PIST_SOFT)
        add_textbox(s, series, x + Pt(10), y + Pt(79), CARD_W - Pt(20), Pt(14),
                    font_size=9.5, color=PISTACHO, bold=True)

        # Foco muscular
        add_textbox(s, foco_e, x + Pt(8), y + CARD_H - Pt(22),
                    CARD_W - Pt(16), Pt(16),
                    font_size=9, italic=True, color=GRIS)

    # Banner tips inferior
    add_rect(s, 0, H - inches(0.75), W, inches(0.75), fill_color=LILA_SOFT)
    tips = "Calienta 5-7 min antes  |  Hidratate durante  |  Descansa 60-90 seg entre series  |  3 veces por semana es suficiente"
    add_textbox(s, tips, inches(0.3), H - inches(0.62),
                W - inches(0.6), Pt(22),
                font_size=9.5, color=LILA, bold=True, align=PP_ALIGN.CENTER)

    pnum += 1

    # ════════════════════════════════════════════════════════
    # SLIDE — RECETAS (ANEXO)
    # ════════════════════════════════════════════════════════
    s = new_slide(prs)
    section_header(s, 18, "Recetas", "preparaciones simples para tu día a día")
    add_textbox(s, "A N E X O", inches(0.5), inches(0.38),
                inches(2), Pt(14), font_size=9, color=LILA)

    recetas = [
        ("RECETA 01", "Hamburguesas caseras",
         "rinde 4 unidades · 25 minutos · ideal para almuerzo o cena",
         ["500 g carne molida magra", "1 zapallo italiano rallado",
          "1 diente de ajo picado", "Cilantro y perejil al gusto",
          "Sal de mar, pimienta", "1 cda aceite de oliva"],
         "Mezclar la carne con el zapallo, ajo y hierbas. Formar 4 hamburguesas y cocinar en sartén con aceite de oliva, 4 minutos por lado. Reposar 2 minutos antes de servir."),
        ("RECETA 02", "Pan de almendra y linaza",
         "rinde 1 porción · 5 minutos · al microondas",
         ["1 huevo de campo", "1 cda harina de almendra",
          "1 cda harina de linaza", "1 cda aceite de oliva", "Sal de mar a gusto"],
         "En un pote apto para microondas, batir el huevo con el aceite y la sal. Agregar las harinas y mezclar. Llevar al microondas 2 minutos a potencia alta. Desmoldar y cortar."),
        ("RECETA 03", "Brownie low carb",
         "rinde 6 porciones · 30 minutos · para tu antojo dulce sin culpa",
         ["1 taza harina de almendras", "½ taza cacao puro orgánico",
          "2 huevos de campo", "2 cdas aceite de coco",
          "Stevia o eritritol al gusto"],
         "Mezclar la harina de almendras y el cacao. Batir los huevos con el aceite y el endulzante. Unir todo, verter en molde y hornear a 180°C por 15-20 minutos."),
    ]

    for i, (rid, nombre_r, sub, ings, prep) in enumerate(recetas):
        y = inches(2.08) + i * inches(2.35)
        add_textbox(s, rid, inches(0.5), y, inches(2), Pt(14),
                    font_size=8, color=PISTACHO)
        add_textbox(s, nombre_r, inches(0.5), y + Pt(14), W - inches(1.0), Pt(30),
                    font_size=20, italic=True, font_name=SERIF, color=OSCURO)
        add_textbox(s, sub, inches(0.5), y + Pt(44), W - inches(1.0), Pt(16),
                    font_size=9.5, italic=True, color=GRIS)
        add_line(s, inches(0.5), y + Pt(62), W - inches(1.0))

        # Columna ingredientes
        add_textbox(s, "INGREDIENTES", inches(0.5), y + Pt(68),
                    inches(3.2), Pt(14), font_size=7.5, color=GRIS)
        for j, ing in enumerate(ings):
            add_textbox(s, f"— {ing}", inches(0.5), y + Pt(82) + j * Pt(18),
                        inches(3.2), Pt(17), font_size=10.5, color=OSCURO)

        # Columna preparación
        add_textbox(s, "PREPARACIÓN", inches(4.0), y + Pt(68),
                    W - inches(4.5), Pt(14), font_size=7.5, color=GRIS)
        add_textbox(s, prep, inches(4.0), y + Pt(82),
                    W - inches(4.5), Pt(70),
                    font_size=10.5, color=OSCURO, word_wrap=True)

        if i < len(recetas) - 1:
            add_line(s, inches(0.5), y + inches(2.25), W - inches(1.0))

    add_textbox(s, "📸  Más recetas en Instagram: @nutrylife_2.0",
                inches(0.5), H - inches(0.7), W - inches(1.0), Pt(20),
                font_size=11.5, italic=True, bold=True, color=LILA, align=PP_ALIGN.CENTER)

    footer(s, "Recetas — Anexo", pnum); pnum += 1

    # ════════════════════════════════════════════════════════
    # SLIDE FINAL — CIERRE
    # ════════════════════════════════════════════════════════
    s = new_slide(prs)
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = OSCURO_BG

    add_textbox(s, "UN ÚLTIMO RECORDATORIO", inches(0.5), inches(0.5),
                W - inches(1.0), Pt(18), font_size=8, color=GRIS, align=PP_ALIGN.CENTER)
    add_line(s, inches(0.5), inches(0.8), W - inches(1.0))

    # Texto grande
    txb = s.shapes.add_textbox(inches(0.5), inches(2.0), W - inches(1.0), inches(4.5))
    tf = txb.text_frame
    tf.word_wrap = True
    p_cierre = tf.paragraphs[0]
    p_cierre.alignment = PP_ALIGN.CENTER
    run = p_cierre.add_run()
    run.text = "Lograremos\ntu mejor\nversión."
    run.font.name = SERIF
    run.font.size = Pt(56)
    run.font.italic = True
    run.font.color.rgb = BLANCO

    add_line(s, inches(0.5), inches(7.2), W - inches(1.0))
    add_textbox(s, "SEGUIMOS EN CONTACTO", inches(0.5), inches(7.3),
                W - inches(1.0), Pt(16), font_size=7.5, color=GRIS, align=PP_ALIGN.CENTER)
    add_textbox(s, "Myriam Márquez García", inches(0.5), inches(7.55),
                W - inches(1.0), Pt(28), font_size=20, italic=True,
                font_name=SERIF, color=BLANCO, align=PP_ALIGN.CENTER)
    add_textbox(s, "nutricion.metodo@gmail.com   ·   @nutrylife_2.0",
                inches(0.5), inches(8.0), W - inches(1.0), Pt(18),
                font_size=10, italic=True, color=GRIS, align=PP_ALIGN.CENTER)

    # ── Guardar ─────────────────────────────────────────────
    prs.save(output_path)
    return output_path
