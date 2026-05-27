"""
pptx_generator.py
Genera el plan nutricional en formato PPTX portrait (Carta vertical)
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
from pathlib import Path

# ── Integración menu_engine (opcional, con fallback) ─────────
try:
    from menu_engine import generar_menu_semanal as _generar_menu
    _MENU_ENGINE_DISPONIBLE = True
except Exception as _menu_engine_import_exc:
    _generar_menu = None
    _MENU_ENGINE_DISPONIBLE = False
    _MENU_ENGINE_IMPORT_ERROR = _menu_engine_import_exc

# Fuente única del documento: el PDF de referencia usa una familia editorial tipo Garamond.
# En PPTX mantenemos EB Garamond en todo el texto; solo cambiamos tamaño, peso, cursiva y espaciado.

# ── Dimensiones Carta / Letter portrait ──────────────────────
# Tamaño carta vertical: 8.5 x 11 pulgadas.
W = Inches(8.5)
H = Inches(11.0)

# ── Paleta ───────────────────────────────────────────────────
VANILLA    = RGBColor(0xF3, 0xF0, 0xEA)  # fondo editorial cálido, similar al PDF de referencia
PISTACHO   = RGBColor(0x6F, 0x7F, 0x62)  # verde salvia sobrio
PIST_SOFT  = RGBColor(0xE7, 0xEA, 0xDF)
LILA       = RGBColor(0x7A, 0x5F, 0xA0)
LILA_SOFT  = RGBColor(0xED, 0xE8, 0xF5)
OSCURO     = RGBColor(0x34, 0x34, 0x34)
GRIS       = RGBColor(0x72, 0x72, 0x72)
GRIS_SOFT  = RGBColor(0xD8, 0xD2, 0xC8)
BLANCO     = RGBColor(0xFF, 0xFF, 0xFF)
OSCURO_BG  = RGBColor(0x34, 0x34, 0x34)

# ── Fuentes ──────────────────────────────────────────────────
SERIF = "EB Garamond"  # fuente editorial única para todo el documento
SANS  = "EB Garamond"  # misma familia: varía tamaño, peso y cursiva

# ── Imágenes de platos (rutas relativas al servidor Render) ──
IMG_BASE = os.path.join(os.path.dirname(__file__), "imagenes")
_IMAGE_INDEX_CACHE   = None   # índice de archivos del repo (construido una vez)
_IMAGE_RESOLVE_CACHE = {}     # img_ref -> ruta resuelta (evita re-buscar el mismo plato)
_IMAGE_CONVERT_CACHE = {}     # ruta_original -> ruta_convertida (evita re-convertir WEBP)

# Activar logs detallados de imágenes con: IMAGE_DEBUG=1 en variables de entorno de Render.
# En producción queda silencioso.
DEBUG_IMAGES = os.environ.get("IMAGE_DEBUG", "0") == "1"

# ── Base de datos de platos ──────────────────────────────────
DESAYUNOS = [
    {"id": "D001", "nombre": "Huevos revueltos con jamón y queso", "desc": "3 huevos, 2 láminas de jamón, 1 lámina de queso mantecoso, ciboulette.", "prot": 30, "kcal": 380, "contiene": ["huevo", "lácteos"]},
    {"id": "D002", "nombre": "Queso chacra con palta y huevo duro", "desc": "100g queso chacra, 2 huevos duros, 1/2 palta, 2 rebanadas pan low carb.", "prot": 30, "kcal": 400, "contiene": ["lácteos", "huevo", "palta"]},
    {"id": "D003", "nombre": "Jamón artesanal mix", "desc": "4 láminas jamón, 2 huevos revueltos, 2 cucharadas mayonesa, 1 palta, espinaca baby.", "prot": 28, "kcal": 420, "contiene": ["huevo", "palta"]},
    {"id": "D004", "nombre": "Tortilla de espinaca con jamón y queso", "desc": "3 huevos, champiñones salteados, jamón, queso philadelphia, 1 rebanada pan low carb.", "prot": 28, "kcal": 360, "contiene": ["huevo", "lácteos"]},
    {"id": "D005", "nombre": "Huevos a la copa", "desc": "3 huevos, 3 cucharadas aceite oliva, sal y pimienta a gusto.", "prot": 20, "kcal": 280, "contiene": ["huevo"]},
    {"id": "D006", "nombre": "Omelette con pollo y palta", "desc": "2 huevos, 100g pollo desmenuzado, 1 palta, 3 cucharadas mayonesa, sal y pimienta.", "prot": 35, "kcal": 450, "contiene": ["huevo", "palta"]},
    {"id": "D007", "nombre": "Quesillo con tomate y jamón", "desc": "200g queso chacra, 3 láminas de jamón, 10 tomates cherry, 30cc aceite oliva, orégano, albahaca.", "prot": 28, "kcal": 380, "contiene": ["lácteos"]},
    {"id": "D008", "nombre": "Salmón ahumado con palta y aceitunas", "desc": "150g salmón ahumado, 1/2 palta, 6 aceitunas.", "prot": 30, "kcal": 350, "contiene": ["pescado", "palta"]},
    {"id": "D009", "nombre": "Huevos revueltos con tocino y palta", "desc": "3 huevos, 3 láminas de tocino, 1/2 palta.", "prot": 28, "kcal": 400, "contiene": ["huevo", "palta", "cerdo"]},
    {"id": "D010", "nombre": "Salmón y huevo pochado", "desc": "50g salmón ahumado, 2 huevos pochados.", "prot": 25, "kcal": 300, "contiene": ["pescado", "huevo"]},
    {"id": "D011", "nombre": "Champiñones rellenos con pollo y queso", "desc": "6 champiñones, 100g pollo desmenuzado, 50g queso parmesano.", "prot": 32, "kcal": 350, "contiene": ["lácteos"]},
    {"id": "D012", "nombre": "Taco de lechuga con carne y queso", "desc": "1 hoja de lechuga, 100g carne, 2 láminas queso mantecoso, 2 cucharadas palta.", "prot": 30, "kcal": 380, "contiene": ["lácteos", "vacuno", "palta"]},
    {"id": "D013", "nombre": "Hamburguesa casera con tomate y palta", "desc": "1 1/2 hamburguesa vacuno, 3 rebanadas tomate, 1/2 palta rebanada.", "prot": 35, "kcal": 420, "contiene": ["vacuno", "palta"]},
    {"id": "D014", "nombre": "Rollitos de jamón y queso", "desc": "4 láminas de jamón pavo cocido, 3 láminas queso mantecoso.", "prot": 22, "kcal": 280, "contiene": ["lácteos"]},
    {"id": "D015", "nombre": "Tomate, queso mozzarella y pollo", "desc": "1 tomate cortado, 100g pollo desmenuzado, queso mozzarella 50g.", "prot": 30, "kcal": 320, "contiene": ["lácteos"]},
    {"id": "D016", "nombre": "Muffins de champiñones y jamón", "desc": "3 huevos, 2 láminas de jamón, 3 champiñones.", "prot": 24, "kcal": 300, "contiene": ["huevo"]},
    {"id": "D017", "nombre": "Pudín de chía con whey y frutos rojos", "desc": "Leche de almendras 240cc, 1 scoop whey, 1 cucharada chía, 50g frambuesas.", "prot": 28, "kcal": 280, "contiene": ["frutos secos"]},
    {"id": "D018", "nombre": "Yogurt griego con whey protein", "desc": "200cc yogurt griego sin azúcar, 1 scoop proteínas.", "prot": 30, "kcal": 250, "contiene": ["lácteos"]},
    {"id": "D019", "nombre": "Atún con palta", "desc": "1 1/2 latas de atún al agua, 1/2 palta.", "prot": 32, "kcal": 300, "contiene": ["pescado", "palta"]},
    {"id": "D020", "nombre": "Queso cabra, tomate y albahaca", "desc": "150g queso cabra, 8 tomates cherry, 5 hojas albahaca, 25cc aceite oliva.", "prot": 28, "kcal": 350, "contiene": ["lácteos"]},
    {"id": "D021", "nombre": "Salmón con palta, huevo y aceitunas", "desc": "Salmón 80g, 2 huevos duros, 1/3 palta, tomate, pepino, 4 aceitunas.", "prot": 30, "kcal": 380, "contiene": ["pescado", "huevo", "palta"]},
]

PLATOS_PRINCIPALES = [
    {"id": "AC001", "nombre": "Carne mechada con ensalada de apio y palta", "desc": "Vacuno mechado, apio picado, 1/2 palta, aceite oliva, limón, sal y pimienta.", "prot": 35, "kcal": 420, "contiene": ["vacuno", "palta"]},
    {"id": "AC002", "nombre": "Pollo con arroz y ensalada", "desc": "Pollo, 3/4 taza arroz, lechuga, tomate, espinaca, aceite oliva, vinagreta.", "prot": 32, "kcal": 400, "contiene": []},
    {"id": "AC003", "nombre": "Salmón a la mantequilla con papas", "desc": "Salmón a la mantequilla, 1 papa salteada, tomates cherry, albahaca, aceite oliva.", "prot": 30, "kcal": 450, "contiene": ["pescado", "lácteos"]},
    {"id": "AC004", "nombre": "Malaya con queso y ensalada de brócoli", "desc": "Vacuno, 100g queso mozzarella, 1/2 taza brócoli, limón, aceite oliva.", "prot": 38, "kcal": 460, "contiene": ["vacuno", "lácteos"]},
    {"id": "AC005", "nombre": "Merluza a la mantequilla con papa cocida", "desc": "Merluza, crema de leche, queso mozzarella, 1 papa cocida, ensalada mix.", "prot": 28, "kcal": 400, "contiene": ["pescado", "lácteos"]},
    {"id": "AC006", "nombre": "Pollo al curry", "desc": "Pollo, crema de leche o coco, curry en pasta, aceite oliva, mix de hojas verdes.", "prot": 34, "kcal": 420, "contiene": []},
    {"id": "AC007", "nombre": "Pollo salteado con verduras", "desc": "Pollo salteado, lechuga, tomate, cebolla morada, aceite oliva, sal y pimienta.", "prot": 32, "kcal": 350, "contiene": []},
    {"id": "AC008", "nombre": "Tímbal de salmón con palta", "desc": "Salmón en tímbal, pimentón, cebolla, aceite oliva, perejil, sal y pimienta.", "prot": 30, "kcal": 380, "contiene": ["pescado"]},
    {"id": "AC009", "nombre": "Pavo plancha con ensalada", "desc": "Pavo a la plancha, cebolla picada, rúcula, berros, tomate cherry, aceitunas.", "prot": 30, "kcal": 320, "contiene": []},
    {"id": "AC010", "nombre": "Pollo al horno con ensalada", "desc": "Pollo al horno, pepino, tomate cherry, lechuga, aceite oliva, cilantro.", "prot": 34, "kcal": 360, "contiene": []},
    {"id": "AC011", "nombre": "Tomate con quesillo, atún y aceitunas", "desc": "Tomate cortado, quesillo molido, atún, aceite oliva, aceitunas, sal y pimienta.", "prot": 28, "kcal": 300, "contiene": ["lácteos", "pescado"]},
    {"id": "AC012", "nombre": "Sushi con arroz", "desc": "8 piezas de sushi, arroz para sushi, salsa de soja.", "prot": 25, "kcal": 380, "contiene": ["pescado", "soya", "gluten"]},
    {"id": "AC013", "nombre": "Caldo de hueso", "desc": "Osobuco en 1 litro agua, 2 horas cocción en olla a presión, sal y pimienta.", "prot": 22, "kcal": 250, "contiene": ["vacuno"]},
    {"id": "AC014", "nombre": "Salmón con espárragos gratinados", "desc": "Salmón, 6 espárragos cocidos y salteados, sal y pimienta.", "prot": 30, "kcal": 350, "contiene": ["pescado"]},
    {"id": "AC015", "nombre": "Merluza con puré de coliflor", "desc": "Merluza, 100g coliflor, crema de leche, queso parmesano.", "prot": 28, "kcal": 320, "contiene": ["pescado", "lácteos"]},
    {"id": "AC016", "nombre": "Carne asada con tortilla de porotos verdes", "desc": "Vacuno asado, 2 huevos, 1 taza porotos verdes, sal y pimienta.", "prot": 38, "kcal": 450, "contiene": ["vacuno", "huevo"]},
    {"id": "AC017", "nombre": "Cerdo asado con ensalada chilena", "desc": "Cerdo asado, 1 tomate, 1/3 cebolla, sal, aceite de oliva.", "prot": 35, "kcal": 400, "contiene": ["cerdo"]},
    {"id": "AC018", "nombre": "Ceviche de reineta", "desc": "Reineta, cebolla morada, pimentón rojo y verde, cilantro, jugo de 2 limones.", "prot": 25, "kcal": 250, "contiene": ["pescado"]},
    {"id": "AC019", "nombre": "Hamburguesas de vacuno con palta y tomate", "desc": "Hamburguesas vacuno, 1/2 tomate, 1/2 palta, mayonesa opcional.", "prot": 35, "kcal": 440, "contiene": ["vacuno", "palta"]},
    {"id": "AC020", "nombre": "Curry de pollo con verduras salteadas", "desc": "Pollo al curry, aceite de coco, ajo, brócoli, cebollín, leche de coco, cilantro.", "prot": 34, "kcal": 420, "contiene": []},
    {"id": "AC021", "nombre": "Pechuga rellena de queso y jamón", "desc": "Pechuga de pollo, 2 láminas queso mantecoso, 2 láminas jamón pavo, aceite oliva.", "prot": 38, "kcal": 400, "contiene": ["lácteos"]},
]

# CENAS eliminado — se usa PLATOS_PRINCIPALES para almuerzos y cenas

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

def spaced(text):
    """Simula letter-spacing editorial para títulos pequeños."""
    return " ".join(str(text).upper())

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
    """Línea horizontal editorial: muy delgada, sólida y oscura, como el PDF de referencia."""
    line = slide.shapes.add_shape(1, left, top, width, Pt(0.35))
    line.fill.solid()
    line.fill.fore_color.rgb = OSCURO
    line.line.fill.background()
    return line

def inches(val):
    return Inches(val)

def pt(val):
    return Pt(val)

def _unique_paths(paths):
    """Devuelve rutas únicas existentes, preservando orden."""
    out = []
    seen = set()
    for x in paths:
        if not x:
            continue
        try:
            px = os.path.abspath(os.path.expanduser(str(x)))
        except Exception:
            continue
        if px not in seen:
            seen.add(px)
            out.append(px)
    return out


def image_search_bases():
    """Rutas base donde buscar imágenes — versión reducida para Render.

    Solo busca en las 6 ubicaciones realmente relevantes, en orden de prioridad.
    Evita iterar padres y decenas de subcarpetas que generaban logs masivos
    y eran inútiles en producción.
    """
    script_dir = os.path.dirname(os.path.abspath(__file__))
    cwd = os.getcwd()
    raw = [
        # 0) Override explícito por variable de entorno (Render / Flask / Streamlit).
        os.environ.get("IMAGE_DIR"),
        os.environ.get("IMAGES_DIR"),
        # 1) Canónica del proyecto: imagenes/ junto al script.
        IMG_BASE,
        os.path.join(script_dir, "imagenes"),
        # 2) Raíz del script (para imágenes en la raíz del repo).
        script_dir,
        # 3) Render: ruta fija del proyecto.
        "/opt/render/project/src/imagenes",
        "/opt/render/project/src",
        # 4) cwd e imagenes/ dentro de cwd (útil en desarrollo local).
        os.path.join(cwd, "imagenes"),
        cwd,
    ]
    return _unique_paths(raw)


def _norm_image_name(name):
    """Normaliza nombres para comparar: sin extensión, espacios, guiones ni underscores."""
    import re
    stem = os.path.splitext(os.path.basename(str(name)))[0]
    return re.sub(r"[^a-z0-9]", "", stem.lower())


def _image_refs_from_obj(img_ref):
    """Extrae posibles referencias de imagen desde un id string o desde el dict del plato."""
    refs = []
    if isinstance(img_ref, dict):
        for key in ["imagen", "image", "foto", "photo", "filename", "file", "img", "id"]:
            val = img_ref.get(key)
            if val:
                refs.append(str(val))
        # NO agregar nombre del plato — solo usar el ID para evitar matcheos incorrectos
    elif img_ref:
        refs.append(str(img_ref))
    # únicos preservando orden
    out=[]
    for r in refs:
        r = r.strip()
        if r and r not in out:
            out.append(r)
    return out


def image_id_candidates(img_ref):
    """Candidatos de nombre para una imagen de plato.

    Soporta:
    - ids exactos: D001, AC001, AC011
    - ids con sufijo de versión del engine: D004-F, A012-X, C002-E → D004, A012, C002
    - cenas antiguas: AC011 -> C001, Cena001, etc.
    - almuerzos alternativos: AC001 -> A001, Almuerzo001
    - variantes con/sin ceros: D001 -> D1, D01
    - si el plato trae campo imagen/foto/filename, también lo usa.
    """
    import re as _re_img
    refs = _image_refs_from_obj(img_ref)
    out = []

    def add(x):
        if x and str(x).strip() and str(x).strip() not in out:
            out.append(str(x).strip())

    # Normalizar refs: si el ID tiene sufijo de versión (ej: D004-F, A012-X, C002-E)
    # agregar también la versión sin sufijo como candidato prioritario.
    refs_expandidos = list(refs)
    for raw in refs:
        stem = os.path.splitext(os.path.basename(raw))[0]
        # Detectar patrón BASE-SUFIJO donde SUFIJO es 1 letra mayúscula (F, M, L, E, X, S, etc.)
        m = _re_img.match(r'^([A-Za-z0-9]+)-([A-Za-z])$', stem)
        if m:
            base_sin_sufijo = m.group(1)
            if base_sin_sufijo not in refs_expandidos:
                refs_expandidos.append(base_sin_sufijo)

    for raw in refs_expandidos:
        add(raw)
        stem = os.path.splitext(os.path.basename(raw))[0]
        add(stem)
        up = stem.upper()

        # Variantes D001 / D01 / D1 / Desayuno001
        if up.startswith("D") and up[1:].isdigit():
            n = int(up[1:])
            add(f"D{n:03d}"); add(f"D{n:02d}"); add(f"D{n}")
            add(f"Desayuno{n:03d}"); add(f"Desayuno{n}")

        # Variantes AC directas
        if up.startswith("AC") and up[2:].isdigit():
            n = int(up[2:])
            add(f"AC{n:03d}"); add(f"AC{n:02d}"); add(f"AC{n}")

        # Variantes A001 directas
        if up.startswith("A") and up[1:].isdigit():
            n = int(up[1:])
            add(f"A{n:03d}"); add(f"Almuerzo{n:03d}")

        # Variantes C001 directas
        # En el repo de Mimi las cenas se llaman AC011..AC020 (continuación de almuerzos)
        if up.startswith("C") and up[1:].isdigit():
            n = int(up[1:])
            add(f"C{n:03d}"); add(f"C{n:02d}"); add(f"C{n}")
            add(f"Cena{n:03d}"); add(f"Cena{n}")
            # Mapeo al sistema de Mimi: C001->AC011, C002->AC012...
            ac_n = n + 10
            add(f"AC{ac_n:03d}"); add(f"AC{ac_n:02d}"); add(f"AC{ac_n}")

    return out


def image_index():
    """Índice cacheado de imágenes del repositorio para no hacer os.walk por cada plato."""
    global _IMAGE_INDEX_CACHE
    if _IMAGE_INDEX_CACHE is not None:
        return _IMAGE_INDEX_CACHE

    exts = {".jpg", ".jpeg", ".png", ".webp"}
    index = {}
    scanned_roots = []

    for root_dir in image_search_bases():
        if not os.path.isdir(root_dir):
            continue
        root_dir = os.path.abspath(root_dir)
        if root_dir in scanned_roots:
            continue
        scanned_roots.append(root_dir)

        # Buscamos más profundo porque muchas apps guardan assets en static/uploads/... etc.
        max_depth = 8
        for cur, dirs, files in os.walk(root_dir):
            rel = os.path.relpath(cur, root_dir)
            depth = 0 if rel == "." else rel.count(os.sep) + 1
            if depth > max_depth:
                dirs[:] = []
                continue
            dirs[:] = [d for d in dirs if d not in {
                ".git", "__pycache__", "node_modules", "venv", ".venv", "site-packages",
                ".cache", ".next", "dist", "build"
            }]
            for fname in files:
                ext = os.path.splitext(fname)[1].lower()
                if ext in exts:
                    full = os.path.join(cur, fname)
                    # key por stem normalizado
                    stem_key = _norm_image_name(fname)
                    index.setdefault(stem_key, full)
                    # key por path relativo normalizado, por si viene 'imagenes/D001.jpg'
                    rel_key = _norm_image_name(os.path.relpath(full, root_dir))
                    index.setdefault(rel_key, full)

    _IMAGE_INDEX_CACHE = index
    print(f"📷 Índice de imágenes construido: {len(index)} claves / {len(set(index.values()))} archivos encontrados.")
    if index:
        muestra = list(dict.fromkeys(index.values()))[:25]
        print("📷 Muestra de imágenes detectadas:")
        for m in muestra:
            print(f"   - {m}")
    else:
        print("⚠️ No se detectó ninguna imagen. Rutas revisadas:")
        for b in image_search_bases():
            print(f"   - {b} {'✅' if os.path.isdir(b) else '❌'}")
    return index


def _maybe_convert_image_for_pptx(path):
    """Convierte a JPEG si python-pptx no soporta el formato real del archivo.

    Detecta el formato REAL del contenido (no solo la extensión) porque
    algunos archivos .jpg son en realidad WEBP u otros formatos.
    Cachea el resultado para no reconvertir el mismo archivo en cada slide.
    """
    global _IMAGE_CONVERT_CACHE
    if path in _IMAGE_CONVERT_CACHE:
        return _IMAGE_CONVERT_CACHE[path]

    result = path  # por defecto, devolver la ruta original
    try:
        from PIL import Image
        with Image.open(path) as im:
            fmt = im.format  # formato real: JPEG, PNG, WEBP, etc.
            if fmt in ("JPEG", "PNG", "BMP", "GIF", "TIFF"):
                result = path  # python-pptx lo soporta directamente
            else:
                # Formato no soportado (WEBP, AVIF, etc.) — convertir a JPEG
                tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".jpg")
                tmp.close()
                im.convert("RGB").save(tmp.name, "JPEG", quality=92)
                if DEBUG_IMAGES:
                    print(f"🔁 {fmt} convertido a JPEG: {path} -> {tmp.name}")
                result = tmp.name
    except Exception as e:
        if DEBUG_IMAGES:
            print(f"⚠️ No se pudo verificar/convertir {path}: {e}")
        result = path

    _IMAGE_CONVERT_CACHE[path] = result
    return result


def resolve_image_path(img_ref):
    """Busca una imagen por id/nombre.
    Prioridad: busca el ID exacto del plato primero (ej: AC005.jpg),
    luego intenta variantes alternativas.
    Cachea el resultado para no repetir búsquedas por el mismo ref.
    """
    global _IMAGE_RESOLVE_CACHE
    if not img_ref:
        return None

    # Clave de cache: dict → frozenset de items, string → string
    try:
        cache_key = frozenset(img_ref.items()) if isinstance(img_ref, dict) else str(img_ref)
    except Exception:
        cache_key = str(img_ref)

    if cache_key in _IMAGE_RESOLVE_CACHE:
        return _IMAGE_RESOLVE_CACHE[cache_key]

    exts = [".jpg", ".jpeg", ".png", ".webp", ".JPG", ".JPEG", ".PNG", ".WEBP"]

    # Extraer el ID exacto del plato
    if isinstance(img_ref, dict):
        exact_id = img_ref.get("id", "")
    else:
        exact_id = str(img_ref)

    result = None

    # PASO 0: buscar el ID exacto en TODAS las bases
    if exact_id:
        if DEBUG_IMAGES:
            print(f"🔍 Buscando {exact_id} en {len(list(image_search_bases()))} bases...")
        for base in image_search_bases():
            if not os.path.isdir(base):
                continue
            try:
                archivos = os.listdir(base)
                exact_lower = exact_id.lower()
                for fname in archivos:
                    name_no_ext, fext = os.path.splitext(fname)
                    if name_no_ext.lower() == exact_lower and fext.lower() in ['.jpg', '.jpeg', '.png', '.webp']:
                        found = os.path.join(base, fname)
                        if DEBUG_IMAGES:
                            print(f"✅ ENCONTRADO: {found}")
                        result = _maybe_convert_image_for_pptx(found)
                        break
                if result:
                    break
                # También intento directo por si acaso
                for ext in exts:
                    path = os.path.join(base, exact_id + ext)
                    if os.path.exists(path):
                        if DEBUG_IMAGES:
                            print(f"✅ ENCONTRADO directo: {path}")
                        result = _maybe_convert_image_for_pptx(path)
                        break
                if result:
                    break
            except Exception as e:
                if DEBUG_IMAGES:
                    print(f"   Error listando {base}: {e}")
                continue
        if not result and DEBUG_IMAGES:
            print(f"❌ {exact_id} no encontrado en ninguna base")

    if not result:
        # PASO 1: candidatos alternativos (fallback)
        names = image_id_candidates(img_ref)
        for base in image_search_bases():
            for name in names:
                root, ext = os.path.splitext(str(name))
                candidates = [str(name)] if ext else [root + e for e in exts]
                for c in candidates:
                    path = os.path.join(base, c)
                    if os.path.exists(path):
                        result = _maybe_convert_image_for_pptx(path)
                        break
                if result:
                    break
            if result:
                break

    if not result:
        wanted_norm = {_norm_image_name(n) for n in (image_id_candidates(img_ref) if not exact_id else [exact_id])}

        # PASO 2: coincidencia normalizada dentro de bases principales.
        for base in image_search_bases():
            if not os.path.isdir(base):
                continue
            try:
                for fname in os.listdir(base):
                    if os.path.splitext(fname)[1].lower() in [e.lower() for e in exts]:
                        fn = _norm_image_name(fname)
                        if fn in wanted_norm:
                            result = _maybe_convert_image_for_pptx(os.path.join(base, fname))
                            break
            except Exception:
                continue
            if result:
                break

    if not result:
        # PASO 3: índice cacheado con verificación estricta.
        names_all = image_id_candidates(img_ref)
        wanted_norm_all = {_norm_image_name(n) for n in names_all}
        index = image_index()
        for wn in wanted_norm_all:
            if wn in index:
                found_path = index[wn]
                found_norm = _norm_image_name(os.path.basename(found_path))
                if found_norm == wn:
                    result = _maybe_convert_image_for_pptx(found_path)
                    break

    _IMAGE_RESOLVE_CACHE[cache_key] = result
    return result

def add_image_safe(slide, img_ref, left, top, width, height):
    """Agrega imagen si existe; si no, pone placeholder. Logs solo con DEBUG_IMAGES=1."""
    path = resolve_image_path(img_ref)
    label = img_ref.get("id") if isinstance(img_ref, dict) else img_ref
    if path:
        try:
            add_picture_contain_rounded(slide, path, left, top, width, height)
            if DEBUG_IMAGES:
                print(f"✅ Imagen insertada: {label} -> {path}")
            return True
        except Exception as e:
            if DEBUG_IMAGES:
                print(f"⚠️ Se encontró pero NO se pudo insertar imagen {path}: {e}")

    if DEBUG_IMAGES:
        cand = ", ".join(image_id_candidates(img_ref))
        print(f"⚠️ Imagen no encontrada para {label}. Candidatos: {cand}")
        print("   Rutas base revisadas:")
        for b in image_search_bases():
            print(f"   - {b} {'✅' if os.path.isdir(b) else '❌'}")

    # Placeholder limpio
    add_rect(slide, left, top, width, height, fill_color=PIST_SOFT, line_color=GRIS_SOFT, line_width=Pt(0.5))
    add_textbox(slide, "imagen", left, top + height//2 - Pt(7),
                width, Pt(14), font_size=7, color=GRIS, align=PP_ALIGN.CENTER)
    return False

def footer(slide, seccion, num):
    """Footer estándar estilo editorial del PDF de referencia."""
    y_footer = H - inches(0.42)
    add_line(slide, inches(0.65), y_footer - Pt(3), W - inches(1.30))
    add_textbox(slide, "MYRIAM NUTRICIÓN", inches(0.65), y_footer,
                inches(2.2), Pt(12), font_size=6.5, color=GRIS)
    add_textbox(slide, seccion.upper(), inches(3.0), y_footer,
                inches(2.4), Pt(12), font_size=6.5, color=GRIS, align=PP_ALIGN.CENTER)
    # No agregamos número abajo a la derecha: basta con el número superior del header.


def section_header(slide, numero, titulo, subtitulo=None, subtitle_size=9.5):
    """Header editorial limpio. El contenido de cada slide debe empezar en inches(2.55) o más."""
    top_label = f"{str(numero).zfill(2)} · {titulo.upper()}"
    add_textbox(slide, spaced(top_label), inches(0.65), inches(0.38),
                inches(4.8), Pt(12), font_size=6.2, color=GRIS)
    add_textbox(slide, f"— {str(numero).zfill(2)} —", W - inches(1.25), inches(0.38),
                inches(0.6), Pt(12), font_size=5.5, color=GRIS, align=PP_ALIGN.RIGHT)
    add_textbox(slide, str(numero).zfill(2), inches(0.65), inches(0.95),
                inches(1.0), Pt(14), font_size=8.0, color=PISTACHO)
    add_textbox(slide, titulo, inches(0.65), inches(1.15),
                W - inches(1.30), Pt(44), font_size=30, italic=True,
                font_name=SERIF, color=OSCURO)
    add_line(slide, inches(0.65), inches(1.85), inches(2.45))
    if subtitulo:
        add_textbox(slide, subtitulo, inches(0.65), inches(2.00),
                    W - inches(1.30), Pt(22), font_size=subtitle_size, italic=True,
                    font_name=SERIF, color=GRIS)
    # La constante CONTENT_Y = inches(2.45) es el mínimo para el contenido de cada slide

def new_slide(prs):
    """Slide en blanco con fondo vainilla y bordes decorativos pistacho/lila."""
    layout = prs.slide_layouts[6]  # blank
    slide = prs.slides.add_slide(layout)
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = VANILLA
    # Bordes decorativos: pistacho arriba-izquierda, lila abajo-derecha
    add_rect(slide, 0, 0, Pt(4), inches(2.0), fill_color=PISTACHO)
    add_rect(slide, 0, 0, inches(2.0), Pt(4), fill_color=PISTACHO)
    add_rect(slide, W - Pt(4), H - inches(2.0), Pt(4), inches(2.0), fill_color=LILA)
    add_rect(slide, W - inches(2.0), H - Pt(4), inches(2.0), Pt(4), fill_color=LILA)
    return slide

def seleccionar_platos(lista, cantidad, fijo_key="fijo"):
    """Selecciona platos preservando el orden del catálogo.

    Importante para las fotos: cada texto debe corresponder al archivo de imagen
    con el mismo ID (D001, AC001, C001, etc.). Antes se usaba random.shuffle(),
    lo que podía desordenar la experiencia visual y dificultar depurar imágenes.
    """
    return list(lista)[:cantidad]

def fecha_legible(fecha_str):
    try:
        dt = datetime.strptime(fecha_str, "%Y-%m-%d")
        meses = ["enero","febrero","marzo","abril","mayo","junio",
                 "julio","agosto","septiembre","octubre","noviembre","diciembre"]
        return f"{dt.day} de {meses[dt.month-1]} {dt.year}"
    except:
        return fecha_str or ""


def clean_number_unit(value, unit):
    """Devuelve valor + unidad evitando duplicar kg/m/% si el formulario ya lo trae."""
    txt = str(value or "").strip()
    if not txt:
        return "—"
    low = txt.lower().replace(" ", "")
    if unit == "%":
        return txt if "%" in txt else f"{txt}%"
    if unit == "kg":
        return txt if "kg" in low else f"{txt} kg"
    if unit == "m":
        return txt if low.endswith("m") else f"{txt} m"
    return txt


def saludo_por_genero(datos, nombre1):
    """Devuelve Querido/Querida según campo sexo/genero/género si viene en datos.
    Si no viene, usa una heurística suave por terminación del primer nombre.
    """
    raw = str(datos.get("sexo") or datos.get("genero") or datos.get("género") or datos.get("gender") or "").strip().lower()
    if raw.startswith(("m", "h", "var", "hom", "masc")):
        return "Querido"
    if raw.startswith(("f", "muj", "fem")):
        return "Querida"
    # fallback imperfecto, solo si no hay dato explícito
    return "Querida" if str(nombre1).strip().lower().endswith("a") else "Querido"

def deduplicar_platos_por_nombre(platos):
    """Elimina duplicados visuales en los slides de opciones/catálogo.

    Criterio: nombre normalizado (sin tildes, minúsculas, sin signos).
    Conserva la primera aparición; si dos platos tienen el mismo nombre normalizado,
    prefiere el de mayor proteína.
    No afecta el menú semanal día a día — solo se usa en los slides de catálogo.
    """
    import unicodedata, re as _re_dedup
    def _norm(s):
        n = unicodedata.normalize("NFD", str(s).lower()).encode("ascii", "ignore").decode()
        return _re_dedup.sub(r"[^a-z0-9]", "", n)

    seen = {}   # nombre_norm -> índice en resultado
    result = []
    for plato in platos:
        key = _norm(plato.get("nombre", ""))
        if key not in seen:
            seen[key] = len(result)
            result.append(plato)
        else:
            # Si el duplicado tiene más proteína, reemplazar
            existing = result[seen[key]]
            try:
                if float(plato.get("prot", 0)) > float(existing.get("prot", 0)):
                    result[seen[key]] = plato
            except (TypeError, ValueError):
                pass
    return result


def find_image(*names):
    """Busca imágenes decorativas (portada, myriam, etc.) priorizando imagenes/.

    Estrategia:
    1. Prueba imagenes/<name>.<ext> directamente (más rápido y preciso).
    2. Cae a resolve_image_path para búsqueda amplia.
    Si no encuentra ninguna, retorna None silenciosamente (no rompe la generación).
    """
    _EXTS = [".jpg", ".jpeg", ".png", ".webp", ".JPG", ".JPEG", ".PNG", ".WEBP"]

    # Prioridad 1: imagenes/ junto al script (ruta canónica del proyecto)
    script_dir = os.path.dirname(os.path.abspath(__file__))
    imagenes_dir = os.path.join(script_dir, "imagenes")

    for name in names:
        stem = os.path.splitext(str(name))[0]  # acepta "portada.jpg" o "portada"
        # Buscar en imagenes/ primero
        for base in [imagenes_dir, IMG_BASE, os.path.join(os.getcwd(), "imagenes"),
                     "/opt/render/project/src/imagenes"]:
            for ext in _EXTS:
                candidate = os.path.join(base, stem + ext)
                if os.path.exists(candidate):
                    return _maybe_convert_image_for_pptx(candidate)
        # Fallback: resolve_image_path (búsqueda amplia con índice)
        found = resolve_image_path(name)
        if found:
            return found

    # No encontrada — silencio, el llamador decide si pone placeholder
    return None


def add_picture_contain(slide, path, left, top, width, height):
    """Inserta una imagen sin deformarla, contenida dentro del rectángulo indicado.

    python-pptx deforma la imagen si se pasan width y height simultáneamente.
    Esta función calcula el tamaño proporcional y centra la imagen dentro del marco.
    """
    try:
        from PIL import Image
        with Image.open(path) as im:
            img_w, img_h = im.size
        if not img_w or not img_h:
            slide.shapes.add_picture(path, left, top, width=width)
            return
        box_ratio = width / height
        img_ratio = img_w / img_h
        if img_ratio >= box_ratio:
            new_w = width
            new_h = int(width / img_ratio)
        else:
            new_h = height
            new_w = int(height * img_ratio)
        new_left = int(left + (width - new_w) / 2)
        new_top = int(top + (height - new_h) / 2)
        slide.shapes.add_picture(path, new_left, new_top, width=new_w, height=new_h)
    except Exception as e:
        print(f"⚠️ No se pudo preservar proporción de imagen {path}: {e}")
        slide.shapes.add_picture(path, left, top, width=width)


def add_picture_contain_rounded(slide, path, left, top, width, height, radius=0):
    """Inserta imagen centrada sin deformar aspect ratio."""
    try:
        from PIL import Image
        with Image.open(path) as im:
            img_w, img_h = im.size
        
        # Calcular tamaño que cabe sin deformar
        img_ratio = img_w / img_h
        box_ratio = int(width) / int(height)
        
        if img_ratio > box_ratio:
            # Imagen más ancha: ajustar por ancho, centrar vertical
            new_w = width
            new_h = int(width / img_ratio)
            new_left = left
            new_top = top + (height - new_h) // 2
        else:
            # Imagen más alta: ajustar por alto, centrar horizontal
            new_h = height
            new_w = int(height * img_ratio)
            new_left = left + (width - new_w) // 2
            new_top = top
        
        slide.shapes.add_picture(path, new_left, new_top, width=new_w, height=new_h)
    except Exception as e:
        print(f"⚠️ Error insertando {path}: {e}")
        try:
            slide.shapes.add_picture(path, left, top, width=width)
        except:
            pass


def add_picture_cover(slide, path, left, top, width, height):
    """Inserta imagen cubriendo el rectángulo, recortando al centro sin deformar.

    Útil para portada: queda llena como foto horizontal editorial, no estirada.
    """
    try:
        from PIL import Image
        with Image.open(path) as im:
            im = im.convert("RGB")
            img_w, img_h = im.size
            box_ratio = width / height
            img_ratio = img_w / img_h
            if img_ratio > box_ratio:
                # imagen demasiado ancha: recortar lados
                new_w = int(img_h * box_ratio)
                x0 = max(0, int((img_w - new_w) / 2))
                crop = im.crop((x0, 0, x0 + new_w, img_h))
            else:
                # imagen demasiado alta: recortar arriba/abajo
                new_h = int(img_w / box_ratio)
                y0 = max(0, int((img_h - new_h) / 2))
                crop = im.crop((0, y0, img_w, y0 + new_h))
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".jpg")
            tmp.close()
            crop.save(tmp.name, "JPEG", quality=92)
            slide.shapes.add_picture(tmp.name, left, top, width=width, height=height)
    except Exception as e:
        print(f"⚠️ No se pudo hacer cover de imagen {path}: {e}")
        slide.shapes.add_picture(path, left, top, width=width)


def add_pie_plate(slide, left, top, size, version="con_carbos"):
    """Dibuja un plato visto desde arriba, dividido en sectores con Pillow.

    Diseño elegante: plato con borde, sombra suave y etiquetas limpias.
    """
    import tempfile, math
    from PIL import Image, ImageDraw, ImageFont

    px = 1400
    img = Image.new("RGBA", (px, px), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    cx = cy = px // 2
    R_outer = int(px * 0.46)   # radio del plato
    R_inner = int(px * 0.10)   # círculo central (vacío decorativo)

    # Sombra del plato (círculo desplazado, gris muy suave)
    shadow_offset = int(px * 0.018)
    for sr in range(R_outer + shadow_offset, R_outer - 1, -1):
        alpha = max(0, 60 - (sr - R_outer) * 3)
        draw.ellipse([
            cx - sr + shadow_offset, cy - sr + shadow_offset,
            cx + sr + shadow_offset, cy + sr + shadow_offset
        ], fill=(150, 140, 130, alpha))

    # Colores de los segmentos — paleta del documento
    if version == "con_carbos":
        seg_data = [
            (180,  0, (111, 127, 98),  "50%",  "Proteína"),       # pistacho
            (270, 90, (122, 95, 160),  "25%",  "Carbohidratos"),   # lila
            (  0, 90, (201, 182, 158), "25%",  "Grasas"),          # beige/arena
        ]
    else:
        seg_data = [
            (180,  0, (111, 127, 98),  "50%",  "Proteína"),
            (270, 90, (138, 180, 122), "25%",  "Vegetales"),
            (  0, 90, (201, 182, 158), "25%",  "Grasas"),
        ]
    # (start_angle_deg, extent, color_rgb, pct, label)
    # Angles: Pillow 0° = right, CCW. Convertimos: start_from_top horario.
    # Proteína = 50% = 180° desde arriba (top=270 en Pillow)
    # Carbos/Veg = 25% = 90° siguientes
    # Grasas = 25° = 90° siguientes

    bbox = [cx - R_outer, cy - R_outer, cx + R_outer, cy + R_outer]

    pillow_angles = [
        (270, 270 + 180),   # Proteína 50%  (arriba hacia abajo, lado derecho)
        (270 + 180, 270 + 270), # Carbos 25%
        (270 + 270, 270 + 360), # Grasas 25%
    ]

    # Dibujar segmentos con borde blanco entre ellos
    for (start, end), (_, _, color, _, _) in zip(pillow_angles, seg_data):
        draw.pieslice(bbox, start=start, end=end,
                      fill=color + (255,), outline=(255,255,255,255), width=6)

    # Borde exterior del plato
    draw.ellipse(bbox, outline=(200, 195, 185, 255), width=12)

    # Círculo central blanco (centro del plato)
    inner_bbox = [cx - R_inner, cy - R_inner, cx + R_inner, cy + R_inner]
    draw.ellipse(inner_bbox, fill=(243, 240, 234, 255), outline=(220,215,205,255), width=4)

    # Fuente
    try:
        font_big   = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSerif-Bold.ttf", 80)
        font_small = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSerif.ttf", 58)
    except Exception:
        try:
            font_big   = ImageFont.truetype("DejaVuSerif-Bold.ttf", 80)
            font_small = ImageFont.truetype("DejaVuSerif.ttf", 58)
        except Exception:
            font_big = font_small = ImageFont.load_default()

    # Posiciones de etiquetas en cada segmento (x, y relativo al centro)
    label_positions = [
        (int(cx + R_outer * 0.55), int(cy)),               # Proteína (der, más afuera)
        (int(cx - R_outer * 0.42), int(cy + R_outer * 0.55)),  # Carbos (abajo izq)
        (int(cx - R_outer * 0.42), int(cy - R_outer * 0.55)),  # Grasas (arriba izq)
    ]

    def draw_centered_text(draw, text, cx, cy, font, color):
        bbox = draw.textbbox((0, 0), text, font=font)
        w, h = bbox[2] - bbox[0], bbox[3] - bbox[1]
        draw.text((cx - w // 2, cy - h // 2), text, font=font, fill=color)

    text_color = (255, 255, 255, 255)
    for (_, _, color, pct, label), (lx, ly) in zip(seg_data, label_positions):
        # Porcentaje grande
        draw_centered_text(draw, pct, lx, ly - 45, font_big, text_color)
        # Label pequeño
        draw_centered_text(draw, label, lx, ly + 50, font_small, text_color)

    # Guardar y retornar
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
    img.save(tmp.name, "PNG")
    slide.shapes.add_picture(tmp.name, left, top, width=size, height=size)
    return tmp.name


def diagnosticar_imagenes(platos):
    """Imprime un resumen de qué imágenes encuentra y cuáles faltan.
    Solo se ejecuta si IMAGE_DEBUG=1 está definido como variable de entorno.
    """
    if not DEBUG_IMAGES:
        return
    print("\n──────── Diagnóstico de imágenes ────────")
    print(f"__file__: {__file__}")
    print(f"cwd: {os.getcwd()}")
    print("Carpetas de búsqueda:")
    for b in image_search_bases():
        exists = "OK" if os.path.isdir(b) else "NO EXISTE"
        print(f"  - {b} [{exists}]")
    for p in platos:
        pid = p.get("id", "")
        path = resolve_image_path(pid)
        cand = ", ".join(image_id_candidates(pid))
        if path:
            print(f"✅ {pid} ({cand}) -> {path}")
        else:
            print(f"⚠️ {pid} ({cand}) -> no encontrada")
    print("────────────────────────────────────────\n")


# ═══════════════════════════════════════════════════════════════
# FUNCIÓN PRINCIPAL
# ═══════════════════════════════════════════════════════════════

def asegurar_distribuciones_menu_engine(datos):
    """Devuelve una copia de datos con distribuciones nutricionales garantizadas.

    menu_engine.py requiere prot_distribucion / carb_distribucion / gras_distribucion.
    Si el formulario no los envía, inyecta defaults conservadores.
    No muta el dict original.
    """
    d = dict(datos)
    if not d.get("prot_distribucion"):
        d["prot_distribucion"] = {"desayuno": 30, "almuerzo": 45, "cena": 25}
    if not d.get("carb_distribucion"):
        d["carb_distribucion"] = {"desayuno": 25, "almuerzo": 40, "cena": 35}
    if not d.get("gras_distribucion"):
        d["gras_distribucion"] = {"desayuno": 40, "almuerzo": 35, "cena": 25}
    # menu_engine requiere carbos > 0. Si el plan es sin_carbos usamos un mínimo simbólico.
    try:
        if float(d.get("carbos") or 0) <= 0:
            d["carbos"] = 1
    except (TypeError, ValueError):
        d["carbos"] = 1
    return d


def generar_pptx_completo(datos, output_path, excel_path=None):
    """
    Genera el PPTX completo del plan nutricional.
    datos: dict con campos del formulario
    output_path: ruta donde guardar el .pptx
    excel_path: ruta opcional al Excel de platos para menu_engine (None = auto)
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
    # Intenta usar menu_engine.py (nuevo motor con Excel).
    # Si falla por cualquier motivo, cae al sistema hardcodeado anterior.

    try:
        if not _MENU_ENGINE_DISPONIBLE:
            raise RuntimeError(
                f"menu_engine no disponible: {getattr(_MENU_ENGINE_IMPORT_ERROR, 'args', ('?',))[0]}"
            )
        datos_engine = asegurar_distribuciones_menu_engine(datos)
        menu_resultado = _generar_menu(datos_engine, excel_path=excel_path)
        dias_menu = menu_resultado.get("dias", [])
        if len(dias_menu) < 7:
            raise ValueError(f"menu_engine devolvió {len(dias_menu)} días (se esperaban 7)")
        desayunos_sel = [dia["desayuno"] for dia in dias_menu[:7]]
        almuerzos_sel = [dia["almuerzo"] for dia in dias_menu[:7]]
        cenas_sel     = [dia["cena"]     for dia in dias_menu[:7]]
        print("✅ Menú generado con menu_engine.py")

    except Exception as _engine_exc:
        print(f"⚠️ No se pudo usar menu_engine.py. Usando fallback antiguo: {_engine_exc}")

        # ── FALLBACK: lógica original con DESAYUNOS / PLATOS_PRINCIPALES ──
        def filtrar_por_restricciones(lista, restricciones_raw):
            """Excluye platos usando el campo 'contiene' explícito.
            Nunca muestra un plato con ingrediente restringido, aunque queden pocos."""
            if not restricciones_raw:
                return lista
            import unicodedata
            def norm(s):
                return unicodedata.normalize("NFD", s.lower()).encode("ascii", "ignore").decode()
            rest_norm = [norm(r) for r in restricciones_raw if r]
            aptos = []
            for plato in lista:
                contiene = [norm(c) for c in plato.get("contiene", [])]
                if not contiene:
                    texto = norm(plato.get("nombre", "") + " " + plato.get("desc", ""))
                    excluir = any(r in texto for r in rest_norm)
                else:
                    excluir = any(r in contiene for r in rest_norm)
                if not excluir:
                    aptos.append(plato)
            return aptos

        rest_lista = restricciones if isinstance(restricciones, list) else []

        def asegurar_7(aptos, originales):
            """Garantiza siempre 7 platos rotando si hay menos."""
            if len(aptos) >= 7:
                return aptos[:7]
            resultado = list(aptos)
            i = 0
            while len(resultado) < 7:
                resultado.append(aptos[i % len(aptos)] if aptos else originales[i % len(originales)])
                i += 1
            return resultado

        import random as _random_fb

        desayunos_filtrados  = filtrar_por_restricciones(DESAYUNOS, rest_lista)
        principales_filtrados = filtrar_por_restricciones(PLATOS_PRINCIPALES, rest_lista)

        pool_des = list(desayunos_filtrados)
        _random_fb.shuffle(pool_des)
        desayunos_sel = asegurar_7(pool_des, DESAYUNOS)

        pool = list(principales_filtrados)
        _random_fb.shuffle(pool)
        if len(pool) >= 14:
            almuerzos_sel = pool[:7]
            cenas_sel     = pool[7:14]
        else:
            almuerzos_sel = asegurar_7(pool, PLATOS_PRINCIPALES)
            restantes = [p for p in pool if p not in almuerzos_sel]
            if not restantes:
                restantes = pool
            cenas_sel = asegurar_7(restantes, PLATOS_PRINCIPALES)

        for nombre_cat, filtrados in [
            ("Desayunos", desayunos_filtrados),
            ("Almuerzos/Cenas", principales_filtrados),
        ]:
            if len(filtrados) < 7:
                print(f"⚠️ {nombre_cat}: solo {len(filtrados)} platos aptos, se completó con rotación.")

    # Diagnóstico visible en consola para depurar rutas/nombres de fotos.
    diagnosticar_imagenes(desayunos_sel + almuerzos_sel + cenas_sel)

    # ── Crear presentación ───────────────────────────────────
    prs = Presentation()
    prs.slide_width  = W
    prs.slide_height = H

    pnum = 1  # contador de página

    # ════════════════════════════════════════════════════════
    # SLIDE 1 — PORTADA
    # ════════════════════════════════════════════════════════
    s = new_slide(prs)

    # Microcabecera y línea superior, copiando proporciones del PDF
    add_line(s, inches(0.65), inches(0.80), W - inches(1.30))
    add_textbox(s, spaced("NUTRICIÓN AVANZADA"), inches(0.72), inches(1.06),
                inches(3.3), Pt(22), font_name=SANS, font_size=10.9, color=GRIS)
    add_textbox(s, spaced("PLAN PERSONALIZADO"), W - inches(3.35), inches(1.06),
                inches(2.7), Pt(22), font_name=SANS, font_size=10.9, color=GRIS, align=PP_ALIGN.RIGHT)

    # Imagen principal portada
    # Buscar portada directamente en las bases de imágenes
    img_portada = None
    _portada_names = ["portada.jpg", "portada.jpeg", "portada.png", "portada.webp",
                      "Portada.jpg", "PORTADA.jpg", "cover.jpg", "cover.jpeg", "cover.png"]
    _portada_bases = [
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "imagenes"),
        IMG_BASE,
        os.path.join(os.getcwd(), "imagenes"),
        "/opt/render/project/src/imagenes",
        "/opt/render/project/src",
    ]
    for base in _portada_bases:
        if not os.path.isdir(base):
            continue
        for nombre_p in _portada_names:
            p_candidate = os.path.join(base, nombre_p)
            if os.path.exists(p_candidate):
                img_portada = p_candidate
                break
        if img_portada:
            break
    if not img_portada:
        img_portada = find_image("portada", "foto_portada", "cover")
    x_img, y_img, w_img, h_img = inches(0.65), inches(1.76), W - inches(1.30), inches(3.75)
    if img_portada:
        try:
            add_picture_cover(s, img_portada, x_img, y_img, w_img, h_img)
        except Exception:
            add_rect(s, x_img, y_img, w_img, h_img, fill_color=PIST_SOFT, line_color=GRIS_SOFT, line_width=Pt(0.5))
    else:
        add_rect(s, x_img, y_img, w_img, h_img, fill_color=PIST_SOFT, line_color=GRIS_SOFT, line_width=Pt(0.5))
        add_textbox(s, "imagen de portada", x_img, y_img + h_img/2 - Pt(8), w_img, Pt(16),
                    font_name=SANS, font_size=8, color=GRIS, align=PP_ALIGN.CENTER)

    add_textbox(s, "Plan Nutricional", inches(0.65), inches(5.90),
                W - inches(1.30), Pt(64), font_size=36, italic=True,
                font_name=SERIF, color=OSCURO, align=PP_ALIGN.CENTER)

    # Bloque inferior de datos: letras del paciente al doble aprox.
    add_line(s, inches(0.65), inches(8.14), W - inches(1.30))
    col_w = (W - inches(1.30)) / 3
    labels = [("PREPARADO PARA", nombre, PP_ALIGN.LEFT),
              ("DURACIÓN", duracion, PP_ALIGN.CENTER),
              ("FECHA DE INICIO", fecha, PP_ALIGN.RIGHT)]
    for i, (lbl, val, al) in enumerate(labels):
        x = inches(0.65) + col_w * i
        add_textbox(s, spaced(lbl), x, inches(8.26), col_w, Pt(20),
                    font_name=SANS, font_size=9.9, color=GRIS, align=al)
        add_textbox(s, val, x, inches(8.54), col_w, Pt(28),
                    font_name=SANS, font_size=18, color=OSCURO, align=al)
    add_line(s, inches(0.65), inches(8.92), W - inches(1.30))

    add_textbox(s, "MYRIAM MÁRQUEZ GARCÍA — N U T R I C I O N I S T A",
                inches(0.65), inches(10.14), W - inches(1.30), Pt(24),
                font_name=SANS, font_size=9.9, color=GRIS, align=PP_ALIGN.CENTER)

    # ════════════════════════════════════════════════════════
    # SLIDE 2 — ÍNDICE
    # ════════════════════════════════════════════════════════
    s = new_slide(prs)

    # Header del índice, replicando la lógica del PDF: microtítulo arriba y número a la derecha.
    add_textbox(s, spaced("ÍNDICE"), inches(0.65), inches(0.40), inches(2.0), Pt(12),
                font_name=SANS, font_size=6.0, color=GRIS)
    add_textbox(s, "— 02 —", W - inches(1.25), inches(0.40), inches(0.6), Pt(12),
                font_name=SANS, font_size=6.0, color=GRIS, align=PP_ALIGN.RIGHT)

    add_textbox(s, "01", inches(0.82), inches(1.18), inches(0.65), Pt(22),
                font_name=SERIF, font_size=17, color=PISTACHO, word_wrap=False)
    add_textbox(s, "Índice", inches(1.52), inches(1.18), inches(3.5), Pt(30),
                font_name=SERIF, font_size=23, italic=True, color=OSCURO)

    secciones = [
        ("02", "Sobre ti"), ("03", "Filosofía del plan"), ("04", "Tu transición alimentaria"),
        ("05", "Tus requerimientos diarios"), ("06", "Tu ritmo del día"), ("07", "Construye tus comidas"),
        ("08", f"Desayunos para {nombre1}"), ("09", f"Almuerzos / Cenas para {nombre1}"),
        ("10", "Menú semanal — Lunes a Domingo"),
        ("12", "Tablas de raciones"), ("13", "Vegetales y frutas"),
        ("14", "Suplementación"), ("15", "Indicaciones generales"),
        ("16", "Lista de compras"), ("17", "Entrenamiento"), ("18", "Recetas — Anexo"),
    ]

    x_num = inches(0.82)
    x_txt = inches(1.60)
    y0 = inches(1.80)
    step = inches(0.415)  # ajustado para que quepan 17 ítems + frase abajo
    for i, (num, titulo) in enumerate(secciones):
        y = y0 + i * step
        add_textbox(s, num, x_num, y, inches(0.70), Pt(18),
                    font_name=SERIF, font_size=16.5, color=PISTACHO,
                    word_wrap=False)
        add_textbox(s, titulo, x_txt, y, W - inches(2.45), Pt(18),
                    font_name=SANS, font_size=15.5, color=OSCURO,
                    align=PP_ALIGN.LEFT, word_wrap=False)

    add_textbox(s, "«Comer sin sentir hambre, es el primer paso para tu cambio de vida.»",
                inches(0.78), inches(9.05), W - inches(1.56), Pt(32),
                font_name=SERIF, font_size=14.5, italic=True, color=OSCURO, align=PP_ALIGN.CENTER)
    add_textbox(s, "— MYRIAM MÁRQUEZ GARCÍA —",
                inches(0.78), inches(9.48), W - inches(1.56), Pt(14),
                font_name=SANS, font_size=7.8, color=GRIS, align=PP_ALIGN.CENTER)

    footer(s, "Índice", pnum); pnum += 1

    # ════════════════════════════════════════════════════════
    # SLIDE 3 — BIENVENIDA
    # ════════════════════════════════════════════════════════
    s = new_slide(prs)

    saludo = saludo_por_genero(datos, nombre1)
    add_textbox(s, spaced("BIENVENIDA"), inches(0.65), inches(0.40), inches(2.0), Pt(12),
                font_name=SANS, font_size=6.0, color=GRIS)
    add_textbox(s, "— 03 —", W - inches(1.25), inches(0.40), inches(0.6), Pt(12),
                font_name=SANS, font_size=6.0, color=GRIS, align=PP_ALIGN.RIGHT)
    add_line(s, inches(0.65), inches(1.55), inches(0.15))
    add_textbox(s, f"{saludo} {nombre1},", inches(0.65), inches(2.05),
                W - inches(1.30), Pt(54), font_size=30, italic=True,
                font_name=SERIF, color=OSCURO)

    # Foto de Myriam: la dejamos arriba-derecha, como firma visual profesional.
    # Abajo quedaría compitiendo con la firma; arriba acompaña la bienvenida sin romper el cierre.
    foto_m = find_image("myriam", "mimi", "foto_myriam", "nutricionista")
    fx, fy, fw, fh = W - inches(2.65), inches(2.45), inches(1.75), inches(1.35)
    if foto_m:
        try:
            add_picture_contain(s, foto_m, fx, fy, fw, fh)
        except Exception:
            add_rect(s, fx, fy, fw, fh, fill_color=PIST_SOFT, line_color=GRIS_SOFT, line_width=Pt(0.5))
    else:
        add_rect(s, fx, fy, fw, fh, fill_color=PIST_SOFT, line_color=GRIS_SOFT, line_width=Pt(0.5))
    add_textbox(s, spaced("TU NUTRICIONISTA"), fx, fy + fh + Pt(5), fw, Pt(12),
                font_name=SANS, font_size=5.8, color=GRIS, align=PP_ALIGN.CENTER)
    add_textbox(s, "Myriam Márquez García", fx - Pt(8), fy + fh + Pt(18), fw + Pt(16), Pt(18),
                font_name=SERIF, font_size=16.5, color=OSCURO, align=PP_ALIGN.CENTER)

    parrafos = [
        "Este plan fue diseñado pensando en ti: en tu cuerpo, tus tiempos, tus gustos y, sobre todo, en tu objetivo. No es una dieta. Es una propuesta de alimentación pensada para acompañarte durante las próximas semanas y ayudarte a construir hábitos sostenibles, con flexibilidad y sin culpa.",
        "Vas a encontrar acá tres cosas: un marco claro de qué comer y cuánto, un sistema de opciones e intercambios para que armes tus comidas con libertad, y una serie de recetas que pueden inspirarte cuando no sepas qué preparar.",
        "Lo más importante: este plan se ajusta a ti, no al revés. Si algo no calza con tu rutina, hablémoslo. Estoy aquí para acompañarte.",
    ]
    # Un solo textbox por párrafo, pero usando ancho reducido para respetar la foto
    texto_w_foto = fx - inches(0.65) - inches(0.15)
    textos_y = [inches(3.28), inches(4.62), inches(5.96)]
    for i, (par, ty) in enumerate(zip(parrafos, textos_y)):
        tw = texto_w_foto if i < 2 else W - inches(1.30)
        add_textbox(s, par, inches(0.65), ty, tw, inches(1.25),
                    font_name=SANS, font_size=12.8, color=OSCURO, word_wrap=True)

    add_textbox(s, "Con cariño,", inches(0.65), inches(8.90), inches(3), Pt(20),
                font_size=19.5, italic=True, font_name=SERIF, color=OSCURO)
    add_textbox(s, "Myriam", inches(0.65), inches(9.18), inches(3), Pt(38),
                font_size=42, italic=True, font_name=SERIF, color=PISTACHO)

    footer(s, "Bienvenida", pnum); pnum += 1

    # ════════════════════════════════════════════════════════
    # SLIDE 4 — SOBRE TI
    # ════════════════════════════════════════════════════════
    s = new_slide(prs)
    section_header(s, 2, "Sobre ti", "tu punto de partida en este viaje")

    # ── Columna izquierda: datos personales + estilo + patologías + foco ──
    y_head = inches(2.55)
    add_textbox(s, spaced("DATOS PERSONALES"), inches(0.65), y_head,
                inches(3.35), Pt(12), font_size=6.5, color=GRIS)
    add_line(s, inches(0.65), y_head + Pt(15), inches(3.00))

    # 1. Datos personales
    datos_izq = [
        ("NOMBRE", nombre),
        ("EDAD", f"{edad} años" if edad and "año" not in str(edad).lower() else (edad or "—")),
        ("ESTILO DE VIDA", datos.get("estilo_vida", "") or "—"),
        ("PATOLOGÍAS", datos.get("patologias", "") or "—"),
        ("FOCO CLÍNICO", foco or "—"),
    ]
    for i, (lbl, val) in enumerate(datos_izq):
        y = inches(2.90) + i * inches(0.65)
        add_textbox(s, spaced(lbl), inches(0.65), y, inches(3.25), Pt(11),
                    font_size=5.8, color=GRIS)
        add_textbox(s, str(val), inches(0.65), y + Pt(12), inches(3.35), Pt(22),
                    font_size=11.0, color=OSCURO, word_wrap=True)

    # ── Columna derecha: métricas + duración + próxima cita ──
    add_textbox(s, spaced("EVALUACIÓN INICIAL"), inches(4.55), y_head,
                inches(3.10), Pt(12), font_size=6.5, color=GRIS)
    add_line(s, inches(4.55), y_head + Pt(15), inches(3.05))

    metricas = [
        (clean_number_unit(peso, "kg"), "PESO"),
        (clean_number_unit(talla, "m"), "TALLA"),
        (str(imc) if imc != "" else "—", "IMC"),
        (clean_number_unit(grasa, "%"), "GRASA CORPORAL"),
    ]
    for i, (val, lbl) in enumerate(metricas):
        col = i % 2
        row = i // 2
        x = inches(4.55) + col * inches(1.60)
        y = inches(2.92) + row * inches(1.15)
        add_textbox(s, str(val), x, y, inches(1.52), Pt(34),
                    font_size=25, italic=True, color=OSCURO, font_name=SERIF, align=PP_ALIGN.LEFT)
        add_textbox(s, spaced(lbl), x, y + Pt(35), inches(1.52), Pt(12),
                    font_size=5.8, color=GRIS)

    # Duración y próxima cita en la columna derecha, debajo de métricas
    dur_y = inches(5.30)
    add_textbox(s, spaced("DURACIÓN DEL PLAN"), inches(4.55), dur_y,
                inches(3.0), Pt(12), font_size=5.8, color=GRIS)
    add_textbox(s, str(duracion), inches(4.55), dur_y + Pt(12),
                inches(3.0), Pt(20), font_size=12, color=OSCURO)

    # Próxima cita: fecha inicio + días de duración
    try:
        from datetime import timedelta
        dias_num = int(''.join(c for c in str(duracion) if c.isdigit()) or '30')
        fecha_str = str(fecha).strip()
        # Intentar varios formatos
        for fmt in ["%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"]:
            try:
                fecha_inicio_dt = datetime.strptime(fecha_str, fmt)
                break
            except ValueError:
                continue
        else:
            fecha_inicio_dt = datetime.now()
        proxima_cita = fecha_inicio_dt + timedelta(days=dias_num)
        proxima_cita_str = proxima_cita.strftime("%d de %B de %Y")
        # Traducir meses al español
        meses_es = {"January":"enero","February":"febrero","March":"marzo","April":"abril",
                    "May":"mayo","June":"junio","July":"julio","August":"agosto",
                    "September":"septiembre","October":"octubre","November":"noviembre","December":"diciembre"}
        for eng, esp in meses_es.items():
            proxima_cita_str = proxima_cita_str.replace(eng, esp)
    except Exception:
        proxima_cita_str = "—"

    cita_y = dur_y + inches(0.60)
    add_textbox(s, spaced("PRÓXIMA CITA"), inches(4.55), cita_y,
                inches(3.0), Pt(12), font_size=5.8, color=GRIS)
    add_textbox(s, proxima_cita_str, inches(4.55), cita_y + Pt(12),
                inches(3.0), Pt(20), font_size=12, italic=True, color=PISTACHO,
                font_name=SERIF)

    # ── Nota clínica al fondo ──
    add_line(s, inches(0.65), inches(6.55), W - inches(1.30))
    add_textbox(s, spaced("NOTA CLÍNICA"), inches(0.65), inches(6.68),
                inches(3), Pt(12), font_size=6.3, color=GRIS)
    nota_clinica = datos.get("nota_personal", "") or f"{nombre} trabaja hacia: {foco}."
    add_textbox(s, nota_clinica, inches(0.65), inches(6.92),
                W - inches(1.30), Pt(70),
                font_size=11.5, color=OSCURO, word_wrap=True)

    footer(s, "Sobre ti", pnum); pnum += 1

    # ════════════════════════════════════════════════════════
    # SLIDE 5 — FILOSOFÍA
    # ════════════════════════════════════════════════════════
    s = new_slide(prs)
    section_header(s, 3, "Filosofía del plan", "los principios que guían tu alimentación", subtitle_size=16.2)

    filosofia = [
        ("01", "Proteína primero",
         "Eje central de cada comida. Fuente de saciedad, masa magra y reparación. Priorizamos cortes magros, huevos, pescados de calidad."),
        ("02", "Carbohidratos estratégicos",
         "Bajos durante el día, presentes alrededor del entrenamiento. Cambia la fisiología del cuerpo sin renunciar al placer ni a la energía."),
        ("03", "Grasas conscientes",
         "Aceite de oliva, palta, frutos secos. Saciedad real, sabor profundo y soporte hormonal. Sin miedo, con criterio."),
    ]

    # Bajamos los bloques para que no choquen con el subtítulo y damos más aire vertical,
    # imitando el PDF: número a la izquierda, texto a la derecha y líneas finas separadoras.
    for i, (num, titulo, texto) in enumerate(filosofia):
        y = inches(3.10) + i * inches(1.82)
        add_textbox(s, num, inches(0.65), y, inches(0.70), Pt(44),
                    font_size=29, italic=True, font_name=SERIF, color=PISTACHO)
        add_textbox(s, titulo, inches(1.42), y + Pt(5), W - inches(2.05), Pt(24),
                    font_size=15.5, bold=True, color=OSCURO)
        add_textbox(s, texto, inches(1.42), y + Pt(36), W - inches(2.05), Pt(92),
                    font_size=16.5, color=OSCURO, word_wrap=True)
        if i < 2:
            add_line(s, inches(0.65), y + inches(1.58), W - inches(1.30))

    add_textbox(s, "Este no es un plan rígido.\nEs un mapa flexible que se ajusta a tu vida.",
                inches(0.65), H - inches(1.25), W - inches(1.30), Pt(42),
                font_size=13, italic=True, color=PISTACHO, font_name=SERIF, align=PP_ALIGN.CENTER)

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

    # Bajamos los encabezados de columna para liberar completamente el subtítulo.
    # En la versión anterior "SUELTA" / "INCORPORA" quedaban demasiado arriba
    # y se montaban sobre "qué soltar y qué incorporar, sin dramas".
    col_header_y = inches(2.30)
    col_line_y = inches(2.50)
    items_start_y = inches(2.90)
    items_gap_y = inches(0.80)

    add_textbox(s, "SUELTA", inches(0.5), col_header_y, inches(3.5), Pt(14),
                font_size=8, color=GRIS)
    add_line(s, inches(0.5), col_line_y, inches(3.5))
    add_textbox(s, "INCORPORA", inches(4.3), col_header_y, inches(3.5), Pt(14),
                font_size=8, color=PISTACHO)
    add_line(s, inches(4.3), col_line_y, inches(3.5))

    for i, ((t1, s1), (t2, s2)) in enumerate(zip(suelta, incorpora)):
        y = items_start_y + i * items_gap_y
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
    # SLIDE 7 — REQUERIMIENTOS
    # ════════════════════════════════════════════════════════
    s = new_slide(prs)
    section_header(s, 5, "Tus requerimientos", "los números que orientan tu día")

    macros_data = [
        (f"{proteinas}g", "PROTEÍNAS", "20% de tu energía"),
        (f"{carbos}g", "CARBOHIDRATOS", "estratégicos, post-entreno"),
        (f"{grasas_g}g", "GRASAS", "fuentes naturales"),
    ]
    col_w3 = (W - inches(1.30)) / 3
    for i, (val, lbl, sub) in enumerate(macros_data):
        x = inches(0.65) + i * col_w3
        # Etiqueta ARRIBA del número para no pisar
        add_textbox(s, spaced(lbl), x, inches(2.65), col_w3, Pt(14),
                    font_name=SANS, font_size=7.2, color=PISTACHO, align=PP_ALIGN.CENTER)
        # Número grande
        add_textbox(s, val, x, inches(2.85), col_w3, Pt(52),
                    font_name=SERIF, font_size=38, color=OSCURO, align=PP_ALIGN.CENTER)
        # Subtítulo debajo
        add_textbox(s, sub, x, inches(3.65), col_w3, Pt(22),
                    font_name=SERIF, font_size=10.5, italic=True, color=GRIS, align=PP_ALIGN.CENTER)

    add_line(s, inches(0.65), inches(5.05), W - inches(1.30))
    extras = [(str(kcal), "CALORÍAS", "kcal/día"), ("2.5", "AGUA", "L / día"), ("≥ 30", "FIBRA", "g / día")]
    for i, (val, lbl, sub) in enumerate(extras):
        x = inches(0.65) + i * col_w3
        add_textbox(s, spaced(lbl), x, inches(5.38), col_w3, Pt(14),
                    font_name=SANS, font_size=7.2, color=GRIS, align=PP_ALIGN.CENTER)
        add_textbox(s, val, x, inches(5.78), col_w3, Pt(48),
                    font_name=SERIF, font_size=36, color=OSCURO, align=PP_ALIGN.CENTER)
        add_textbox(s, sub, x, inches(6.52), col_w3, Pt(18),
                    font_name=SERIF, font_size=11, italic=True, color=GRIS, align=PP_ALIGN.CENTER)

    footer(s, "Requerimientos", pnum); pnum += 1

    # ════════════════════════════════════════════════════════
    # SLIDE 8 — TU RITMO DEL DÍA
    # ════════════════════════════════════════════════════════
    s = new_slide(prs)
    section_header(s, 6, "Tu ritmo del día", "tres tiempos, marcados por la luz y el cuerpo")


    ritmo_imgs = ["D001", "AC002", "AC005"]
    horarios = [
        ("08:00", "Desayuno", "Proteina solida, grasas buenas, opcional vegetales. Ojalá dentro de la primera hora después de despertar."),
        ("13:30", "Almuerzo", "El plato más completo del día: proteina, vegetales abundantes, grasas. Aquí pueden ir los carbohidratos si entrenaste en la mañana."),
        ("20:00", "Cena", "Versión más liviana del almuerzo. Proteina, vegetales, grasas. Idealmente cerrar la cocina 3 horas antes de dormir."),
    ]
    RITMO_IMG_W = inches(1.30)
    RITMO_IMG_H = inches(0.90)
    for i, (hora, comida, desc) in enumerate(horarios):
        y = inches(2.60) + i * inches(1.68)
        # Hora
        add_textbox(s, hora, inches(0.65), y, inches(1.10), Pt(34),
                    font_name=SERIF, font_size=22, italic=True, color=PISTACHO)
        # Nombre comida
        add_textbox(s, comida, inches(1.85), y + Pt(2), inches(1.5), Pt(22),
                    font_name=SANS, font_size=14, bold=True, color=OSCURO)
        # Descripción
        add_textbox(s, desc, inches(1.85), y + Pt(24), W - inches(2.50) - RITMO_IMG_W - inches(0.20), Pt(48),
                    font_name=SANS, font_size=10.5, color=GRIS, word_wrap=True)
        # Foto a la derecha
        img_x = W - RITMO_IMG_W - inches(0.65)
        add_image_safe(s, {"id": ritmo_imgs[i]}, img_x, y, RITMO_IMG_W, RITMO_IMG_H)
        if i < 2:
            add_line(s, inches(0.65), y + inches(1.15), W - inches(1.30))

    # Mensaje de flexibilidad abajo, separado de los horarios para evitar traslapes.
    add_textbox(s, "Si hay días que solo puedes hacer dos comidas,\nestá bien — escucha tu hambre y respétala.",
                inches(0.95), inches(8.05), W - inches(1.90), Pt(60),
                font_name=SERIF, font_size=17, italic=True, color=OSCURO, align=PP_ALIGN.CENTER)

    footer(s, "Ritmo del día", pnum); pnum += 1

    # ════════════════════════════════════════════════════════
    # SLIDE 9 — CONSTRUYE TUS COMIDAS
    # ════════════════════════════════════════════════════════
    s = new_slide(prs)
    section_header(s, 7, "Construye tus comidas", "el sistema simple para armar tu plato")

    # Texto introductorio — empieza en 2.55 para no pisar el subtítulo
    add_textbox(s, "Cada comida sigue una fórmula simple: proteína + vegetales + grasa. Las cantidades varían según el momento del día.",
                inches(0.65), inches(2.55), W - inches(1.30), Pt(36),
                font_name=SANS, font_size=11.5, color=OSCURO, word_wrap=True)

    # TORTA circular real — vista de plato desde arriba
    torta_size = inches(5.0)
    torta_x = inches(0.55)
    torta_y = inches(3.20)
    add_pie_plate(s, torta_x, torta_y, torta_size, version)

    # Leyenda a la derecha de la torta
    es_carbos = (version == "con_carbos")
    leyenda = [
        (PISTACHO,                          "50%", "Proteína",
         "Pollo, pescado, carne magra"),
        (LILA if es_carbos else RGBColor(0x8E,0xB4,0x7A),
                                            "25%",
         "Carbohidratos" if es_carbos else "Vegetales",
         "Camote, arroz integral" if es_carbos else "Hojas verdes, crucíferas"),
        (RGBColor(0xC9,0xB6,0x9E),          "25%", "Grasas",
         "Palta, oliva, frutos secos"),
    ]
    ley_x = torta_x + torta_size + inches(0.35)
    ley_w = W - ley_x - inches(0.45)
    for i, (color, pct, label, desc) in enumerate(leyenda):
        ly = torta_y + inches(0.45) + i * inches(1.55)
        add_rect(s, ley_x, ly, Pt(6), inches(1.1), fill_color=color)
        add_textbox(s, f"{pct}  {label}",
                    ley_x + Pt(14), ly,
                    ley_w, Pt(22),
                    font_name=SERIF, font_size=14, bold=True, color=OSCURO)
        add_textbox(s, desc,
                    ley_x + Pt(14), ly + Pt(26),
                    ley_w, Pt(34),
                    font_name=SANS, font_size=10.5, color=GRIS, word_wrap=True)

    footer(s, "Construye tus comidas", pnum); pnum += 1

    # ════════════════════════════════════════════════════════
    # PORTADA — TU SEMANA SUGERIDA
    # ════════════════════════════════════════════════════════
    s = new_slide(prs)
    add_textbox(s, "Tu semana", inches(0.65), inches(3.0),
                W - inches(1.30), Pt(65), font_size=48, italic=True,
                font_name=SERIF, color=OSCURO)
    add_textbox(s, "sugerida", inches(0.65), inches(4.0),
                W - inches(1.30), Pt(55), font_size=42, italic=True,
                font_name=SERIF, color=PISTACHO)
    add_rect(s, inches(0.65), inches(5.0), inches(2.5), Pt(3), fill_color=PISTACHO)
    add_textbox(s, "Un ejemplo de cómo armar tus días con las opciones de este plan.",
                inches(0.65), inches(5.30), W - inches(1.30), Pt(24),
                font_size=12, italic=True, font_name=SANS, color=GRIS)
    footer(s, "Tu semana sugerida", pnum); pnum += 1

    # ════════════════════════════════════════════════════════
    # SLIDES 12-18 — MENÚ SEMANAL (1 slide por día)
    # ════════════════════════════════════════════════════════
    for idx, dia in enumerate(DIAS):
        s = new_slide(prs)

        # ── Header limpio sin fondo de color ────────────────
        add_textbox(s, "10  ·  MENÚ SEMANAL", inches(0.55), inches(0.32),
                    inches(4), Pt(12), font_size=6.5, color=GRIS, font_name=SANS)
        add_textbox(s, dia, inches(0.55), inches(0.50),
                    W - inches(1.0), Pt(52), font_size=38, italic=True,
                    font_name=SERIF, color=OSCURO)
        add_line(s, inches(0.55), inches(1.42), W - inches(1.1))

        # ── Tres comidas — layout 2 columnas: texto izq, foto der ──
        comidas = [
            ("DESAYUNO  ·  08:00", desayunos_sel[idx], PISTACHO),
            ("ALMUERZO  ·  13:30", almuerzos_sel[idx], PISTACHO),
            ("CENA  ·  20:00",     cenas_sel[idx],     LILA),
        ]

        MENU_STEP = inches(3.12)
        IMG_W_M   = inches(2.10)   # foto más grande
        IMG_H_M   = inches(1.65)   # foto más alta

        for ci, (label, plato, acento) in enumerate(comidas):
            y_base = inches(1.55) + ci * MENU_STEP

            # Foto centrada verticalmente dentro del espacio de la comida
            img_x = W - IMG_W_M - inches(0.45)
            img_y_center = y_base + (MENU_STEP - IMG_H_M) // 2
            add_image_safe(s, plato, img_x, img_y_center, IMG_W_M, IMG_H_M)

            # Calcular centro vertical del bloque de texto para alinearlo con la foto
            text_w = W - inches(1.20) - IMG_W_M - inches(0.15)
            desc = plato.get("desc", "")
            import re as _re
            items_desc = [x.strip() for x in _re.split(r"[.,]", desc) if x.strip()][:5]
            n_items = len(items_desc)
            # Altura total del bloque: etiqueta + nombre + items + macros
            block_h = Pt(14) + Pt(26) + n_items * Pt(16) + Pt(8) + Pt(14)
            text_y_start = y_base + (MENU_STEP - block_h) // 2  # centrado vertical

            # Etiqueta
            add_textbox(s, label, inches(0.55), text_y_start,
                        inches(3.5), Pt(14), font_size=8, color=acento,
                        bold=True, font_name=SANS)

            # Nombre del plato
            add_textbox(s, plato["nombre"],
                        inches(0.55), text_y_start + Pt(18),
                        text_w, Pt(26),
                        font_name=SERIF, font_size=15, bold=True, color=OSCURO)

            # Ingredientes en lista vertical
            for ii, item in enumerate(items_desc):
                add_textbox(s, f"— {item}",
                            inches(0.55), text_y_start + Pt(46) + ii * Pt(16),
                            text_w, Pt(15),
                            font_name=SANS, font_size=9.5, color=GRIS)

            # Macros
            macro_y = text_y_start + Pt(46) + n_items * Pt(16) + Pt(8)
            add_textbox(s, f"{plato['prot']}g proteína  ·  {plato['kcal']} kcal",
                        inches(0.55), macro_y,
                        inches(3.5), Pt(14),
                        font_name=SERIF, font_size=8.5, color=acento, italic=True)

            if ci < 2:
                add_line(s, inches(0.55), y_base + MENU_STEP - Pt(4), W - inches(1.1))

        footer(s, f"Menú — {dia}", pnum); pnum += 1


    # ════════════════════════════════════════════════════════
    # PORTADA — TUS OPCIONES DE COMIDAS
    # ════════════════════════════════════════════════════════
    s = new_slide(prs)
    add_textbox(s, "Tus opciones", inches(0.65), inches(3.0),
                W - inches(1.30), Pt(65), font_size=48, italic=True,
                font_name=SERIF, color=OSCURO)
    add_textbox(s, "de comidas", inches(0.65), inches(4.0),
                W - inches(1.30), Pt(55), font_size=42, italic=True,
                font_name=SERIF, color=PISTACHO)
    add_rect(s, inches(0.65), inches(5.0), inches(2.5), Pt(3), fill_color=PISTACHO)
    add_textbox(s, "Todas las alternativas disponibles para armar tu semana.",
                inches(0.65), inches(5.30), W - inches(1.30), Pt(24),
                font_size=12, italic=True, font_name=SANS, color=GRIS)
    footer(s, "Tus opciones de comidas", pnum); pnum += 1

    # ════════════════════════════════════════════════════════
    # SLIDES 9, 10, 11 — DESAYUNOS / ALMUERZOS / CENAS PARA TI
    # ════════════════════════════════════════════════════════

    # Deduplicar por nombre normalizado solo para el catálogo visual.
    # El menú semanal día por día NO se toca — sigue usando desayunos_sel[idx], etc.
    _desayunos_cat   = deduplicar_platos_por_nombre(desayunos_sel)
    _almcenas_cat    = deduplicar_platos_por_nombre(almuerzos_sel + cenas_sel)

    for seccion_idx, (titulo, platos, color_acento, sec_num) in enumerate([
        (f"Desayunos para {nombre1}", _desayunos_cat, PISTACHO, 8),
        (f"Almuerzos / Cenas para {nombre1}", _almcenas_cat, PISTACHO, 9),
    ]):
        s = new_slide(prs)
        section_header(s, sec_num, titulo,
                       "opciones seleccionadas para ti" if seccion_idx == 0
                       else "opciones seleccionadas para tu semana")

        # Lista compacta estilo PDF: número lateral, texto al centro, foto pequeña a la derecha.
        IMG_W = inches(0.92)
        IMG_H = inches(0.58)
        TEXT_X = inches(0.65)
        NUM_W = inches(0.55)
        TEXT_W = W - inches(1.30) - IMG_W - NUM_W - inches(0.25)
        ROW_H = inches(0.88)

        # Paginar: 7 platos por slide
        PLATOS_POR_SLIDE = 7
        for i, plato in enumerate(platos):
            # Si es el plato 8 (o múltiplo de 7), crear nueva slide
            if i > 0 and i % PLATOS_POR_SLIDE == 0:
                footer(s, titulo, pnum); pnum += 1
                s = new_slide(prs)
                section_header(s, sec_num, titulo, "(continuación)")
            y = inches(2.75) + (i % PLATOS_POR_SLIDE) * ROW_H

            add_textbox(s, f"{i + 1:02d}", TEXT_X, y + Pt(8), NUM_W, Pt(24),
                        font_name=SERIF, font_size=20, italic=True, color=color_acento)

            img_x = W - IMG_W - inches(0.65)
            add_image_safe(s, plato, img_x, y + Pt(5), IMG_W, IMG_H)

            add_textbox(s, plato["nombre"], TEXT_X + NUM_W, y + Pt(2),
                        TEXT_W, Pt(18), font_name=SANS, font_size=11.5, bold=True, color=OSCURO)

            macro_txt = f"{plato['prot']}g prot · {plato['kcal']} kcal"
            add_textbox(s, macro_txt, TEXT_X + NUM_W, y + Pt(20),
                        TEXT_W, Pt(12), font_name=SERIF, font_size=8.3, color=color_acento, italic=True)

            add_textbox(s, plato["desc"], TEXT_X + NUM_W, y + Pt(34),
                        TEXT_W, Pt(23), font_name=SANS, font_size=8.4, color=GRIS, word_wrap=True)

            if i < len(platos) - 1:
                add_line(s, TEXT_X, y + ROW_H - Pt(4), W - inches(1.30))

        footer(s, titulo, pnum); pnum += 1

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
                         proteinas_tab, inches(2.55))
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
        add_textbox(s, titulo_col, x, inches(2.55), inches(3.5), Pt(14),
                    font_size=8, color=GRIS)
        add_textbox(s, sub_col, x, inches(2.72), inches(3.5), Pt(14),
                    font_size=9.5, italic=True, color=GRIS)
        add_line(s, x, inches(2.95), inches(3.5))

        for i, (nombre_v, detalle) in enumerate(datos_col):
            y = inches(3.08) + i * inches(0.82)
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
        y = inches(2.55) + i * inches(1.05)

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

    y = inches(2.55)
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
            add_textbox(s, cell, xp, y + Pt(3), cw, Pt(16),
                        font_size=10.5, color=clr)
        y += Pt(22)

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
        y = inches(2.55) + row * inches(1.55)

        add_textbox(s, num, x, y, inches(0.55), Pt(30),
                    font_size=22, italic=True, font_name=SERIF, color=PISTACHO,
                    word_wrap=False)
        add_textbox(s, titulo, x + inches(0.55), y + Pt(5), col_w2 - inches(0.6), Pt(20),
                    font_size=13, bold=True, color=OSCURO)
        add_textbox(s, texto, x + inches(0.55), y + Pt(26), col_w2 - inches(0.6), Pt(42),
                    font_size=10.5, color=GRIS, word_wrap=True)

        if row < 2 and col == 1:
            add_line(s, inches(0.5), y + Pt(80), W - inches(1.0))

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

    y_cursor = inches(2.55)
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
    # SLIDE — ENTRENAMIENTO (fondo claro, cards elegantes)
    # ════════════════════════════════════════════════════════
    s = new_slide(prs)
    section_header(s, 17, "Tu rutina de movimiento", "3 veces por semana · sin gimnasio · solo tú y tu cuerpo")

    ejercicios = [
        ("01", "Sentadillas",        "3 series", "12-15 rep",   "Piernas · Glúteos",
         "Pies al ancho de hombros. Baja hasta muslos paralelos al piso. Espalda recta."),
        ("02", "Puente de glúteos",  "3 series", "15-20 rep",   "Glúteos · Femorales",
         "Pies apoyados, rodillas dobladas. Aprieta glúteos al subir la cadera."),
        ("03", "Flexiones",          "3 series", "10-12 rep",   "Pecho · Tríceps",
         "Manos bajo hombros, cuerpo en tabla. Baja controlando hasta rozar el piso."),
        ("04", "Remo con botellas",  "3 series", "12-15 rep",   "Espalda · Bíceps",
         "Inclínate 45°, espalda recta. Jala los codos hacia atrás con control."),
        ("05", "Zancadas alternas",  "3 series", "10 rep/lado", "Piernas · Equilibrio",
         "Rodilla trasera cerca del piso. Pecho erguido en todo momento."),
        ("06", "Plancha abdominal",  "3 series", "30-45 seg",   "Core · Estabilidad",
         "Cuerpo en línea recta. Aprieta abdomen, respira con ritmo constante."),
    ]

    # Layout: 2 columnas × 3 filas, fondo claro, cards con borde suave
    CARD_W = (W - inches(1.20)) / 2
    CARD_H = inches(2.35)
    GAP_X  = inches(0.15)
    GAP_Y  = inches(0.12)

    for i, (num, nombre_e, series, reps, musculo, tip) in enumerate(ejercicios):
        col = i % 2
        row = i // 2
        x = inches(0.50) + col * (CARD_W + GAP_X)
        y = inches(2.55) + row * (CARD_H + GAP_Y)

        # Card blanca con borde suave
        add_rect(s, x, y, CARD_W, CARD_H,
                 fill_color=BLANCO, line_color=GRIS_SOFT, line_width=Pt(0.8))

        # Barra de acento izquierda
        acento = PISTACHO if i % 2 == 0 else LILA
        add_rect(s, x, y, Pt(5), CARD_H, fill_color=acento)

        # Layout: nombre a la izquierda, foto a la derecha
        EJ_IMG_W = Pt(70)
        EJ_IMG_H = Pt(70)

        # Número y nombre
        add_textbox(s, num, x + Pt(14), y + Pt(8), Pt(30), Pt(20),
                    font_name=SERIF, font_size=16, italic=True, bold=True,
                    color=acento, word_wrap=False)

        add_textbox(s, nombre_e, x + Pt(14), y + Pt(30),
                    CARD_W - EJ_IMG_W - Pt(30), Pt(20),
                    font_name=SERIF, font_size=13, bold=True, color=OSCURO,
                    word_wrap=False)

        # Músculo target
        add_textbox(s, musculo, x + Pt(14), y + Pt(50),
                    CARD_W - EJ_IMG_W - Pt(30), Pt(14),
                    font_name=SERIF, font_size=8.5, italic=True, color=GRIS)

        # Foto del ejercicio a la derecha (mantiene aspect ratio)
        ej_names = [f"ejercicio_{i+1:02d}", f"ej{i+1:02d}", f"ej_{i+1}",
                    nombre_e.lower().replace(" ","_").replace("(","").replace(")","")]
        ej_path = None
        for en in ej_names:
            ej_path = resolve_image_path(en)
            if ej_path:
                break
        img_x_ej = x + CARD_W - EJ_IMG_W - Pt(8)
        if ej_path:
            try:
                add_picture_contain_rounded(s, ej_path, img_x_ej, y + Pt(8), EJ_IMG_W, EJ_IMG_H)
            except:
                ej_path = None
        if not ej_path:
            colores_ej = [PISTACHO, LILA, RGBColor(0x8E,0xA8,0x7A),
                          RGBColor(0x5E,0x78,0x8A), RGBColor(0x9E,0x7A,0x5E), RGBColor(0x6B,0x7E,0x5A)]
            add_rect(s, img_x_ej, y + Pt(8), EJ_IMG_W, EJ_IMG_H, fill_color=colores_ej[i % 6])

        # Badges: series y reps
        badge_y = y + Pt(72)
        badge_w = (CARD_W - Pt(28)) / 2
        add_rect(s, x + Pt(14), badge_y, badge_w - Pt(4), Pt(22), fill_color=PIST_SOFT)
        add_textbox(s, series, x + Pt(16), badge_y + Pt(4),
                    badge_w - Pt(8), Pt(16),
                    font_name=SANS, font_size=10, bold=True, color=PISTACHO,
                    align=PP_ALIGN.CENTER)

        add_rect(s, x + Pt(14) + badge_w, badge_y, badge_w - Pt(4), Pt(22), fill_color=LILA_SOFT)
        add_textbox(s, reps, x + Pt(16) + badge_w, badge_y + Pt(4),
                    badge_w - Pt(8), Pt(16),
                    font_name=SANS, font_size=10, bold=True, color=LILA,
                    align=PP_ALIGN.CENTER)

        # Tip de técnica (sin línea para ahorrar espacio)
        add_textbox(s, tip, x + Pt(14), badge_y + Pt(28),
                    CARD_W - Pt(28), Pt(44),
                    font_name=SANS, font_size=8.5, color=GRIS, word_wrap=True)

    # Banner inferior con tips
    banner_y = H - inches(0.72)
    add_rect(s, 0, banner_y, W, inches(0.72), fill_color=PIST_SOFT)
    tips = "Calienta 5-7 min antes  |  Hidrátate durante  |  Descansa 60-90 seg entre series"
    add_textbox(s, tips, inches(0.3), banner_y + Pt(12),
                W - inches(0.6), Pt(20),
                font_size=9.5, color=PISTACHO, bold=True, font_name=SANS,
                align=PP_ALIGN.CENTER)

    pnum += 1

    # ════════════════════════════════════════════════════════
    # SLIDE — RECETAS (ANEXO)
    # ════════════════════════════════════════════════════════
    s = new_slide(prs)
    section_header(s, 18, "Recetas", "preparaciones simples para tu día a día")
    add_textbox(s, "A N E X O", W - inches(1.5), inches(0.38),
                inches(1.2), Pt(14), font_size=7, color=LILA,
                align=PP_ALIGN.RIGHT)

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

    # IDs de imagen para cada receta (usar platos del catálogo que coincidan visualmente)
    receta_img_ids = ["receta01", "receta02", "receta03"]  # hamburguesas, pan, brownie aprox

    for i, (rid, nombre_r, sub, ings, prep) in enumerate(recetas):
        # Cada receta en su propia slide
        if i > 0:
            s = new_slide(prs)
            add_textbox(s, "A N E X O", W - inches(1.5), inches(0.38),
                        inches(1.2), Pt(14), font_size=7, color=LILA,
                        align=PP_ALIGN.RIGHT)
            section_header(s, 18, "Recetas", "preparaciones simples para tu día a día")

        # Layout: foto ocupa la mitad derecha completa (sin solaparse con nada)
        # Texto a la izquierda en columna independiente
        IMG_R_W = inches(3.8)
        IMG_R_H = inches(4.20)
        img_rx  = W - IMG_R_W - inches(0.40)
        img_ry  = inches(2.55)

        # Ancho disponible para texto (columna izquierda)
        txt_w = W - IMG_R_W - inches(1.30)  # margen izq + espacio hasta la foto

        # Etiqueta de receta
        add_textbox(s, rid, inches(0.55), inches(2.55),
                    txt_w, Pt(14), font_size=8, color=PISTACHO, font_name=SANS)

        # Nombre grande
        add_textbox(s, nombre_r, inches(0.55), inches(2.75),
                    txt_w, Pt(26), font_size=18, italic=True,
                    font_name=SERIF, color=OSCURO, word_wrap=False)

        # Subtítulo
        add_textbox(s, sub, inches(0.55), inches(3.10),
                    txt_w, Pt(18), font_size=9, italic=True,
                    color=GRIS, font_name=SANS)

        add_line(s, inches(0.55), inches(3.58), txt_w)

        # Ingredientes debajo del subtítulo
        add_textbox(s, "INGREDIENTES", inches(0.55), inches(3.68),
                    txt_w, Pt(14), font_size=7.5, color=GRIS,
                    font_name=SANS, bold=True)
        for j, ing in enumerate(ings):
            add_textbox(s, f"— {ing}",
                        inches(0.55), inches(3.86) + j * inches(0.30),
                        txt_w, Pt(18),
                        font_size=11, color=OSCURO, font_name=SANS)

        # Preparación debajo de ingredientes
        y_prep = inches(3.86) + len(ings) * inches(0.30) + inches(0.15)
        add_line(s, inches(0.55), y_prep, txt_w)
        add_textbox(s, "PREPARACIÓN", inches(0.55), y_prep + Pt(6),
                    txt_w, Pt(14), font_size=7.5, color=GRIS,
                    font_name=SANS, bold=True)
        add_textbox(s, prep, inches(0.55), y_prep + Pt(22),
                    txt_w, H - y_prep - inches(0.9),
                    font_size=11, color=OSCURO, font_name=SANS, word_wrap=True)

        # Foto — se dibuja al final para quedar encima si hubiera solapamiento
        if i < len(receta_img_ids):
            add_image_safe(s, {"id": receta_img_ids[i]}, img_rx, img_ry, IMG_R_W, IMG_R_H)

        footer(s, f"Recetas — {rid}", pnum); pnum += 1

    # ════════════════════════════════════════════════════════
    # SLIDE FINAL — DESPEDIDA (todo en una sola slide)
    # ════════════════════════════════════════════════════════
    s = new_slide(prs)

    # Frase principal centrada
    txb = s.shapes.add_textbox(inches(0.5), inches(2.0), W - inches(1.0), inches(3.5))
    tf = txb.text_frame
    tf.word_wrap = True
    p_cierre = tf.paragraphs[0]
    p_cierre.alignment = PP_ALIGN.CENTER
    run = p_cierre.add_run()
    run.text = "Lograremos\ntu mejor\nversión."
    run.font.name = SERIF
    run.font.size = Pt(52)
    run.font.italic = True
    run.font.color.rgb = OSCURO

    # Línea pistacho decorativa
    add_rect(s, inches(2.5), inches(5.8), inches(3.5), Pt(2.5), fill_color=PISTACHO)

    # Instagram
    add_textbox(s, "Sígueme en Instagram", inches(0.5), inches(6.20),
                W - inches(1.0), Pt(24), font_size=14, italic=True,
                font_name=SERIF, color=GRIS, align=PP_ALIGN.CENTER)
    add_textbox(s, "@nutrylife_2.0", inches(0.5), inches(6.60),
                W - inches(1.0), Pt(36), font_size=26, bold=True,
                font_name=SERIF, color=PISTACHO, align=PP_ALIGN.CENTER)
    add_textbox(s, "Más recetas, tips y motivación para tu estilo de vida saludable",
                inches(0.5), inches(7.15), W - inches(1.0), Pt(18),
                font_size=10, italic=True, font_name=SANS, color=GRIS, align=PP_ALIGN.CENTER)

    # Línea + contacto
    add_line(s, inches(0.65), inches(7.65), W - inches(1.30))
    add_textbox(s, "SEGUIMOS EN CONTACTO", inches(0.5), inches(7.80),
                W - inches(1.0), Pt(14), font_size=7, color=GRIS,
                font_name=SANS, align=PP_ALIGN.CENTER)
    add_textbox(s, "Myriam Márquez García", inches(0.5), inches(8.00),
                W - inches(1.0), Pt(28), font_size=20, italic=True,
                font_name=SERIF, color=OSCURO, align=PP_ALIGN.CENTER)
    add_textbox(s, "nutricion.metodo@gmail.com   ·   @nutrylife_2.0",
                inches(0.5), inches(8.45), W - inches(1.0), Pt(18),
                font_size=10, italic=True, color=GRIS,
                font_name=SANS, align=PP_ALIGN.CENTER)

    # ── Guardar ─────────────────────────────────────────────
    prs.save(output_path)
    return output_path


def round_corners_image(image_path, radius=30):
    """Aplica esquinas redondeadas a una imagen con Pillow."""
    from PIL import Image, ImageDraw
    import tempfile
    img = Image.open(image_path).convert("RGBA")
    w, h = img.size
    # Crear máscara con esquinas redondeadas
    mask = Image.new("L", (w, h), 0)
    draw = ImageDraw.Draw(mask)
    r = min(radius, w // 4, h // 4)
    draw.rounded_rectangle([0, 0, w, h], radius=r, fill=255)
    # Aplicar máscara
