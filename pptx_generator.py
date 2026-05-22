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
_IMAGE_INDEX_CACHE = None

# ── Base de datos de platos ──────────────────────────────────
DESAYUNOS = [
    {"id": "D001", "nombre": "Huevos revueltos con salmón",
     "desc": "3 huevos revueltos en mantequilla con salmón ahumado, cebollín y tomate cherry.",
     "prot": 28, "kcal": 360,
     "contiene": ["huevo", "salmón", "lácteos"]},
    {"id": "D002", "nombre": "Huevos con champiñones",
     "desc": "3 huevos al plato con champiñones salteados, espinaca y cúrcuma.",
     "prot": 22, "kcal": 310,
     "contiene": ["huevo"]},
    {"id": "D003", "nombre": "Yogurt con berries",
     "desc": "Yogurt entero sin azúcar, puñado de berries, almendras tostadas y chía.",
     "prot": 15, "kcal": 280,
     "contiene": ["lácteos", "frutos secos"]},
    {"id": "D004", "nombre": "Wrap de lechuga con atún",
     "desc": "Hojas de lechuga rellenas con atún, palta, pepino y salsa de yogurt.",
     "prot": 26, "kcal": 290,
     "contiene": ["palta", "lácteos"]},
    {"id": "D005", "nombre": "Tostadas de almendra",
     "desc": "Pan de almendra con queso de cabra, palta y 2 huevos duros. Aceite de oliva.",
     "prot": 24, "kcal": 380,
     "contiene": ["huevo", "palta", "lácteos", "frutos secos"]},
    {"id": "D006", "nombre": "Bowl de semillas",
     "desc": "Yogurt griego con chía, linaza, coco rallado, nueces y berries. Canela.",
     "prot": 18, "kcal": 290,
     "contiene": ["lácteos", "frutos secos"]},
    {"id": "D007", "nombre": "Plato frío proteico",
     "desc": "150g pollo desmenuzado o salmón ahumado, palta, tomate cherry, aceitunas.",
     "prot": 30, "kcal": 350,
     "contiene": ["palta", "salmón"]},
    {"id": "D008", "nombre": "Omelette mediterráneo",
     "desc": "3 huevos con espinaca, champiñones y queso de cabra. Aceite de oliva, cúrcuma.",
     "prot": 22, "kcal": 320,
     "contiene": ["huevo", "lácteos"]},
    {"id": "D009", "nombre": "Bowl de palta y huevo",
     "desc": "½ palta machacada sobre pan low carb. 2 huevos pochados y semillas de sésamo.",
     "prot": 20, "kcal": 340,
     "contiene": ["huevo", "palta"]},
    {"id": "D010", "nombre": "Batido proteico",
     "desc": "Proteína en polvo + leche de almendras + frambuesas + mantequilla de almendras.",
     "prot": 25, "kcal": 300,
     "contiene": ["frutos secos"]},
]

ALMUERZOS = [
    {"id": "AC001", "nombre": "Vacuno con vegetales asados",
     "desc": "150g lomo de vacuno. Vegetales asados: pimentón, zapallo, cebolla. Rúcula.",
     "prot": 38, "kcal": 450,
     "contiene": ["vacuno"]},
    {"id": "AC002", "nombre": "Pollo al ajillo con coliflor",
     "desc": "150g pechuga al ajillo. Coliflor asada con cúrcuma. Ensalada de tomate y palta.",
     "prot": 38, "kcal": 400,
     "contiene": ["palta"]},
    {"id": "AC003", "nombre": "Cerdo al limón con brócoli",
     "desc": "150g filete de cerdo al limón y hierbas. Brócoli al vapor con ajo y oliva.",
     "prot": 35, "kcal": 410,
     "contiene": ["cerdo"]},
    {"id": "AC004", "nombre": "Pechuga rellena con espinaca",
     "desc": "Pechuga rellena con espinaca y queso de cabra al horno. Vegetales asados.",
     "prot": 40, "kcal": 430,
     "contiene": ["lácteos"]},
    {"id": "AC005", "nombre": "Hamburguesas caseras",
     "desc": "Carne molida con cilantro y zapallo italiano. Ensalada de pepino y palta.",
     "prot": 40, "kcal": 480,
     "contiene": ["vacuno", "palta"]},
    {"id": "AC006", "nombre": "Bowl de huevo y vegetales",
     "desc": "3 huevos cocidos. Mix de hojas verdes, brócoli, champiñones y sésamo.",
     "prot": 20, "kcal": 350,
     "contiene": ["huevo"]},
    {"id": "AC007", "nombre": "Salmón al horno con espárragos",
     "desc": "150g salmón al horno con limón y eneldo. Espárragos asados con almendras.",
     "prot": 35, "kcal": 420,
     "contiene": ["salmón", "frutos secos"]},
    {"id": "AC008", "nombre": "Atún con ensalada fresca",
     "desc": "1 lata atún en aceite de oliva. Mix de hojas, pepino, tomate, aceitunas.",
     "prot": 30, "kcal": 320,
     "contiene": ["pescado"]},
    {"id": "AC009", "nombre": "Pollo al curry con coliflor",
     "desc": "150g pollo en salsa de curry y leche de coco. Arroz de coliflor y pepino.",
     "prot": 36, "kcal": 420,
     "contiene": []},
    {"id": "AC010", "nombre": "Pescado blanco al wok",
     "desc": "150g reineta al sartén. Mix de pimentón, zapallo y brócoli con jengibre.",
     "prot": 32, "kcal": 380,
     "contiene": ["pescado"]},
]

CENAS = [
    {"id": "C001", "nombre": "Bowl nocturno liviano",
     "desc": "Yogurt griego con pepino, menta y oliva. 2 huevos duros y aceitunas.",
     "prot": 20, "kcal": 260,
     "contiene": ["huevo", "lácteos"]},
    {"id": "C002", "nombre": "Salmón al sartén con espinaca",
     "desc": "130g salmón. Espinaca salteada con ajo. Tomatitos cherry. Limón y oliva.",
     "prot": 30, "kcal": 360,
     "contiene": ["salmón"]},
    {"id": "C003", "nombre": "Pollo desmenuzado con palta",
     "desc": "120g pollo desmenuzado, palta, limón y cilantro sobre hojas verdes.",
     "prot": 32, "kcal": 340,
     "contiene": ["palta"]},
    {"id": "C004", "nombre": "Caldo de huesos con vegetales",
     "desc": "Caldo casero con verduras de temporada y jengibre. 2 huevos duros.",
     "prot": 16, "kcal": 220,
     "contiene": ["huevo", "vacuno"]},
    {"id": "C005", "nombre": "Reineta con ensalada mediterránea",
     "desc": "130g reineta al horno. Ensalada de pepino, aceitunas, tomate y feta.",
     "prot": 28, "kcal": 310,
     "contiene": ["pescado", "lácteos"]},
    {"id": "C006", "nombre": "Revuelto de claras con vegetales",
     "desc": "5 claras revueltas con espinaca, tomate y queso cottage. Tostada low carb.",
     "prot": 25, "kcal": 270,
     "contiene": ["huevo", "lácteos"]},
    {"id": "C007", "nombre": "Crema de zapallo",
     "desc": "Crema de zapallo con caldo de huesos, jengibre y cúrcuma. 2 huevos duros.",
     "prot": 18, "kcal": 290,
     "contiene": ["huevo"]},
    {"id": "C008", "nombre": "Tortilla de verduras",
     "desc": "3 huevos con pimentón, cebolla, espinaca y queso de cabra. Ensalada verde.",
     "prot": 22, "kcal": 310,
     "contiene": ["huevo", "lácteos"]},
    {"id": "C009", "nombre": "Merluza al vapor",
     "desc": "130g merluza al vapor con limón y hierbas. Brócoli y coliflor con ajo.",
     "prot": 28, "kcal": 300,
     "contiene": ["pescado"]},
    {"id": "C010", "nombre": "Vacuno liviano con ensalada",
     "desc": "120g filete magro de vacuno. Ensalada de hojas verdes, tomate y palta.",
     "prot": 30, "kcal": 350,
     "contiene": ["vacuno", "palta"]},
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
    """Rutas base donde buscar imágenes.

    Esta versión es más amplia porque en producción el script puede ejecutarse
    desde una carpeta distinta a la del repositorio. Busca en:
    - variable de entorno IMAGE_DIR / IMAGES_DIR, si existe
    - carpeta /imagenes junto al script
    - carpeta del script
    - carpeta actual de ejecución
    - padres cercanos del script y del cwd
    - subcarpetas típicas: imagenes, images, img, static/images, assets/images, etc.
    """
    script_dir = os.path.dirname(os.path.abspath(__file__))
    cwd = os.getcwd()
    raw = []

    # 0) Override explícito por variable de entorno, muy útil en Render/Flask/Streamlit.
    raw += [os.environ.get("IMAGE_DIR"), os.environ.get("IMAGES_DIR")]

    # 1) Carpetas directas más probables.
    raw += [IMG_BASE, script_dir, cwd, os.path.join(cwd, "imagenes")]

    # 2) Padres cercanos del script y del cwd.
    for base in [script_dir, cwd]:
        cur = Path(base)
        for _ in range(5):
            raw.append(str(cur))
            for sub in [
                "imagenes", "imagenes/comidas", "imagenes/platos",
                "images", "images/food", "img", "fotos",
                "static", "static/images", "static/imagenes", "static/img",
                "assets", "assets/images", "assets/imagenes", "public", "public/images",
            ]:
                raw.append(str(cur / sub))
            if cur.parent == cur:
                break
            cur = cur.parent

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
        # Agregar nombre del plato como fallback muy flexible.
        if img_ref.get("nombre"):
            refs.append(str(img_ref.get("nombre")))
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
    - cenas antiguas: AC011 -> C001, Cena001, etc.
    - almuerzos alternativos: AC001 -> A001, Almuerzo001
    - variantes con/sin ceros: D001 -> D1, D01
    - si el plato trae campo imagen/foto/filename, también lo usa.
    """
    refs = _image_refs_from_obj(img_ref)
    out = []

    def add(x):
        if x and str(x).strip() and str(x).strip() not in out:
            out.append(str(x).strip())

    for raw in refs:
        add(raw)
        stem = os.path.splitext(os.path.basename(raw))[0]
        add(stem)
        up = stem.upper()

        # Variantes D001 / D01 / D1 / Desayuno001
        if up.startswith("D") and up[1:].isdigit():
            n = int(up[1:])
            add(f"D{n:03d}"); add(f"D{n:02d}"); add(f"D{n}")
            add(f"Desayuno{n:03d}"); add(f"Desayuno{n}")

        # Variantes AC001 / A001 / Almuerzo001 y AC011 -> C001 / Cena001
        if up.startswith("AC") and up[2:].isdigit():
            n = int(up[2:])
            add(f"AC{n:03d}"); add(f"AC{n:02d}"); add(f"AC{n}")
            if 1 <= n <= 10:
                add(f"A{n:03d}"); add(f"A{n:02d}"); add(f"A{n}")
                add(f"Almuerzo{n:03d}"); add(f"Almuerzo{n}")
            if 11 <= n <= 20:
                c = n - 10
                add(f"C{c:03d}"); add(f"C{c:02d}"); add(f"C{c}")
                add(f"Cena{c:03d}"); add(f"Cena{c}")

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
    """Convierte WEBP a PNG temporal si python-pptx no puede insertarla."""
    ext = os.path.splitext(path)[1].lower()
    if ext != ".webp":
        return path
    try:
        from PIL import Image
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
        tmp.close()
        with Image.open(path) as im:
            im.convert("RGB").save(tmp.name, "PNG")
        print(f"🔁 WEBP convertido temporalmente a PNG: {path} -> {tmp.name}")
        return tmp.name
    except Exception as e:
        print(f"⚠️ No se pudo convertir WEBP {path}: {e}")
        return path


def resolve_image_path(img_ref):
    """Busca una imagen por id/nombre de forma robusta."""
    if not img_ref:
        return None

    exts = [".jpg", ".jpeg", ".png", ".webp", ".JPG", ".JPEG", ".PNG", ".WEBP"]
    names = image_id_candidates(img_ref)

    # 1) Coincidencia exacta en bases principales.
    for base in image_search_bases():
        for name in names:
            root, ext = os.path.splitext(str(name))
            candidates = [str(name)] if ext else [root + e for e in exts]
            for c in candidates:
                path = os.path.join(base, c)
                if os.path.exists(path):
                    return _maybe_convert_image_for_pptx(path)

    wanted_norm = {_norm_image_name(n) for n in names}

    # 2) Coincidencia normalizada dentro de bases principales.
    # Solo acepta si la norma del archivo coincide EXACTAMENTE con uno de los candidatos.
    for base in image_search_bases():
        if not os.path.isdir(base):
            continue
        try:
            for fname in os.listdir(base):
                if os.path.splitext(fname)[1].lower() in [e.lower() for e in exts]:
                    fn = _norm_image_name(fname)
                    if fn in wanted_norm:
                        return _maybe_convert_image_for_pptx(os.path.join(base, fname))
        except Exception:
            continue

    # 3) Búsqueda en índice cacheado con verificación estricta:
    # el nombre normalizado del archivo debe coincidir EXACTAMENTE,
    # no solo estar contenido. Esto evita que "c001" matchee "ac001".
    index = image_index()
    for wn in wanted_norm:
        if wn in index:
            # Verificar que el nombre del archivo normalizado sea igual al candidato
            found_path = index[wn]
            found_norm = _norm_image_name(os.path.basename(found_path))
            if found_norm == wn:
                return _maybe_convert_image_for_pptx(found_path)

    return None

def add_image_safe(slide, img_ref, left, top, width, height):
    """Agrega imagen si existe; si no, pone placeholder y deja diagnóstico en consola."""
    path = resolve_image_path(img_ref)
    label = img_ref.get("id") if isinstance(img_ref, dict) else img_ref
    if path:
        try:
            add_picture_cover(slide, path, left, top, width, height)
            print(f"✅ Imagen insertada: {label} -> {path}")
            return True
        except Exception as e:
            print(f"⚠️ Se encontró pero NO se pudo insertar imagen {path}: {e}")

    cand = ", ".join(image_id_candidates(img_ref))
    print(f"⚠️ Imagen no encontrada para {label}. Candidatos: {cand}")
    print("   Rutas base revisadas:")
    for b in image_search_bases():
        print(f"   - {b} {'✅' if os.path.isdir(b) else '❌'}")

    # Placeholder limpio: evita emojis/glyphs rotos en PowerPoint/LibreOffice.
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
    """Slide en blanco con fondo vainilla."""
    layout = prs.slide_layouts[6]  # blank
    slide = prs.slides.add_slide(layout)
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = VANILLA
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

def find_image(*names):
    """Busca imágenes habituales solo por coincidencia exacta/normalizada.

    Evita coincidencias parciales demasiado amplias, por ejemplo que `cover` tome
    una imagen interna del sistema o de una librería en vez de la portada real.
    Para portada usa `portada.jpg` o `portada.jpeg`; para Myriam usa `myriam.jpg`
    o `myriam.jpeg`.
    """
    for name in names:
        path = resolve_image_path(name)
        if path:
            print(f"✅ Imagen encontrada: {name} -> {path}")
            return path

    print(f"⚠️ Imagen no encontrada para alternativas {names}. Buscada en: {image_search_bases()}")
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
    """Imprime un resumen claro de qué imágenes encuentra y cuáles faltan."""
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
            print(f"✅ {p.get("id", p)} ({cand}) -> {path}")
        else:
            print(f"⚠️ {p.get("id", p)} ({cand}) -> no encontrada")
    print("────────────────────────────────────────\n")


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

    # ── Seleccionar platos respetando restricciones ─────────
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
            # Usar campo "contiene" si existe (explícito y confiable)
            contiene = [norm(c) for c in plato.get("contiene", [])]
            # Fallback: parsear nombre + desc si no hay campo contiene
            if not contiene:
                texto = norm(plato.get("nombre","") + " " + plato.get("desc",""))
                excluir = any(r in texto for r in rest_norm)
            else:
                excluir = any(r in contiene for r in rest_norm)
            if not excluir:
                aptos.append(plato)
        return aptos

    rest_lista = restricciones if isinstance(restricciones, list) else []
    desayunos_sel = filtrar_por_restricciones(DESAYUNOS, rest_lista)[:7]
    almuerzos_sel = filtrar_por_restricciones(ALMUERZOS, rest_lista)[:7]
    cenas_sel     = filtrar_por_restricciones(CENAS, rest_lista)[:7]

    # Advertencia en consola si quedan pocos platos
    for nombre_cat, sel in [("Desayunos", desayunos_sel),
                             ("Almuerzos", almuerzos_sel),
                             ("Cenas", cenas_sel)]:
        if len(sel) < 7:
            print(f"⚠️ {nombre_cat}: solo {len(sel)} platos aptos con las restricciones indicadas.")

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
    img_portada = find_image("portada", "foto_portada", "cover", "imagen_portada", "portada_plan", "cover_image")
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
        ("08", f"Desayunos para {nombre1}"), ("09", f"Almuerzos para {nombre1}"),
        ("10", f"Cenas para {nombre1}"), ("11", "Menú semanal — Lunes a Domingo"),
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

    # Encabezados más bajos para evitar choque con subtítulo. Líneas finas como el PDF.
    y_head = inches(2.62)
    add_textbox(s, spaced("DATOS PERSONALES"), inches(0.65), y_head,
                inches(3.35), Pt(12), font_size=6.5, color=GRIS)
    add_line(s, inches(0.65), y_head + Pt(15), inches(3.00))

    datos_personales = [
        ("NOMBRE", nombre),
        ("EDAD", f"{edad} años" if edad and "año" not in str(edad).lower() else (edad or "—")),
        ("FOCO CLÍNICO", foco or "—"),
        ("ESTILO DE VIDA", datos.get("estilo_vida", "—")),
    ]
    for i, (lbl, val) in enumerate(datos_personales):
        y = inches(2.98) + i * inches(0.62)
        add_textbox(s, spaced(lbl), inches(0.65), y, inches(3.25), Pt(11),
                    font_size=5.8, color=GRIS)
        add_textbox(s, str(val), inches(0.65), y + Pt(12), inches(3.35), Pt(20),
                    font_size=12.0, color=OSCURO)

    add_textbox(s, spaced("EVALUACIÓN INICIAL"), inches(4.55), y_head,
                inches(3.10), Pt(12), font_size=6.5, color=GRIS)
    add_line(s, inches(4.55), y_head + Pt(15), inches(3.05))

    metricas = [
        (clean_number_unit(peso, "kg"), "PESO"),
        (clean_number_unit(talla, "m"), "TALLA"),
        (str(imc) if imc != "" else "—", "IMC"),
        (clean_number_unit(grasa, "%"), "GRASA CORPORAL"),
    ]
    # Números en disposición 2x2, sin duplicar unidades y sin traslapar etiquetas.
    for i, (val, lbl) in enumerate(metricas):
        col = i % 2
        row = i // 2
        x = inches(4.55) + col * inches(1.60)
        y = inches(3.00) + row * inches(1.22)
        add_textbox(s, str(val), x, y, inches(1.52), Pt(34),
                    font_size=25, italic=True, color=OSCURO, font_name=SERIF, align=PP_ALIGN.LEFT)
        add_textbox(s, spaced(lbl), x, y + Pt(35), inches(1.52), Pt(12),
                    font_size=5.8, color=GRIS)

    # Nota clínica
    add_line(s, inches(0.65), inches(6.15), W - inches(1.30))
    add_textbox(s, spaced("NOTA CLÍNICA"), inches(0.65), inches(6.30),
                inches(3), Pt(12), font_size=6.3, color=GRIS)
    nota_clinica = datos.get("nota_personal", "") or f"{nombre} trabaja hacia: {foco}."
    add_textbox(s, nota_clinica, inches(0.65), inches(6.58),
                W - inches(1.30), Pt(80),
                font_size=11.8, color=OSCURO, word_wrap=True)

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
        add_textbox(s, val, x, inches(2.85), col_w3, Pt(52),
                    font_name=SERIF, font_size=38, color=OSCURO, align=PP_ALIGN.CENTER)
        add_textbox(s, spaced(lbl), x, inches(3.45), col_w3, Pt(14),
                    font_name=SANS, font_size=7.2, color=PISTACHO, align=PP_ALIGN.CENTER)
        add_textbox(s, sub, x, inches(3.68), col_w3, Pt(22),
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


    horarios = [
        ("08:00", "Desayuno", "Proteína sólida, grasas buenas, opcional vegetales. Ojalá dentro de la primera hora después de despertar."),
        ("13:30", "Almuerzo", "El plato más completo del día: proteína, vegetales abundantes, grasas. Aquí pueden ir los carbohidratos si entrenaste en la mañana."),
        ("20:00", "Cena", "Versión más liviana del almuerzo. Proteína, vegetales, grasas. Idealmente cerrar la cocina 3 horas antes de dormir."),
    ]
    for i, (hora, comida, desc) in enumerate(horarios):
        y = inches(2.60) + i * inches(1.58)
        add_textbox(s, hora, inches(0.80), y, inches(1.25), Pt(34),
                    font_name=SERIF, font_size=24, italic=True, color=PISTACHO)
        add_textbox(s, comida, inches(2.20), y + Pt(2), inches(1.6), Pt(22),
                    font_name=SANS, font_size=14, bold=True, color=OSCURO)
        add_textbox(s, desc, inches(3.80), y + Pt(2), W - inches(4.45), Pt(42),
                    font_name=SANS, font_size=10.8, color=GRIS, word_wrap=True)
        if i < 2:
            add_line(s, inches(0.80), y + inches(0.88), W - inches(1.60))

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
    # SLIDES 9, 10, 11 — DESAYUNOS / ALMUERZOS / CENAS PARA TI
    # ════════════════════════════════════════════════════════
    for seccion_idx, (titulo, platos, color_acento, sec_num) in enumerate([
        (f"Desayunos para {nombre1}", desayunos_sel, PISTACHO, 8),
        (f"Almuerzos para {nombre1}", almuerzos_sel, PISTACHO, 9),
        (f"Cenas para {nombre1}", cenas_sel, LILA, 10),
    ]):
        s = new_slide(prs)
        section_header(s, sec_num, titulo,
                       "siete opciones seleccionadas para ti" if seccion_idx == 0
                       else "siete platos completos para el mediodía" if seccion_idx == 1
                       else "siete opciones livianas para cerrar el día")

        # Lista compacta estilo PDF: número lateral, texto al centro, foto pequeña a la derecha.
        IMG_W = inches(0.92)
        IMG_H = inches(0.58)
        TEXT_X = inches(0.65)
        NUM_W = inches(0.55)
        TEXT_W = W - inches(1.30) - IMG_W - NUM_W - inches(0.25)
        ROW_H = inches(0.88)

        for i, plato in enumerate(platos):
            y = inches(2.60) + i * ROW_H

            add_textbox(s, f"0{i+1}", TEXT_X, y + Pt(8), NUM_W, Pt(24),
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

            # Etiqueta sin fondo — solo texto con acento de color
            add_textbox(s, label, inches(0.55), y_base,
                        inches(3.5), Pt(14), font_size=8, color=acento,
                        bold=True, font_name=SANS)

            # Foto grande a la derecha, alineada con el nombre del plato
            img_x = W - IMG_W_M - inches(0.45)
            add_image_safe(s, plato, img_x, y_base + Pt(2), IMG_W_M, IMG_H_M)

            # Nombre del plato
            add_textbox(s, plato["nombre"],
                        inches(0.55), y_base + Pt(18),
                        W - inches(1.20) - IMG_W_M - inches(0.15), Pt(26),
                        font_name=SERIF, font_size=15, bold=True, color=OSCURO)

            # Ingredientes en lista vertical (parsear la desc por comas/puntos)
            desc = plato.get("desc", "")
            # Dividir por punto o coma para generar lista
            import re as _re
            items_desc = [x.strip() for x in _re.split(r"[.,]", desc) if x.strip()]
            text_w = W - inches(1.20) - IMG_W_M - inches(0.15)
            for ii, item in enumerate(items_desc[:5]):
                add_textbox(s, f"— {item}",
                            inches(0.55), y_base + Pt(46) + ii * Pt(16),
                            text_w, Pt(15),
                            font_name=SANS, font_size=9.5, color=GRIS)

            # Macros
            add_textbox(s, f"{plato['prot']}g proteína  ·  {plato['kcal']} kcal",
                        inches(0.55), y_base + Pt(130),
                        inches(3.5), Pt(14),
                        font_name=SERIF, font_size=8.5, color=acento, italic=True)

            if ci < 2:
                add_line(s, inches(0.55), y_base + MENU_STEP - Pt(4), W - inches(1.1))

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
        y = inches(2.55) + row * inches(1.55)

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
    # SLIDE — ENTRENAMIENTO (diseño premium, gym de alta gama)
    # ════════════════════════════════════════════════════════
    s = new_slide(prs)

    # Fondo dividido: mitad oscura arriba (premium/gym), mitad vainilla abajo
    add_rect(s, 0, 0, W, inches(4.78), fill_color=OSCURO)

    # Línea accent pistacho a la izquierda del header
    add_rect(s, 0, 0, Pt(6), inches(4.78), fill_color=PISTACHO)

    # Número de sección y etiqueta
    add_textbox(s, "17", inches(0.25), inches(0.28), inches(1), Pt(16),
                font_size=8, color=PISTACHO, font_name=SANS)
    add_textbox(s, "ENTRENAMIENTO", inches(0.25), inches(0.42), inches(3), Pt(14),
                font_size=7, color=RGBColor(0xA0,0xA8,0x98), font_name=SANS)

    # Título principal grande en blanco
    add_textbox(s, "Tu rutina", inches(0.25), inches(0.75),
                W - inches(0.5), Pt(55), font_size=40, italic=True,
                font_name=SERIF, color=BLANCO)
    add_textbox(s, "de movimiento", inches(0.25), inches(1.55),
                W - inches(0.5), Pt(38), font_size=26, italic=True,
                font_name=SERIF, color=PISTACHO)

    # Subtítulo motivacional
    add_textbox(s, "3 veces por semana · sin gimnasio · solo tú y tu cuerpo",
                inches(0.25), inches(2.35), W - inches(0.5), Pt(22),
                font_size=11, color=RGBColor(0xC0,0xC8,0xB8), font_name=SANS, italic=True)

    # Franja divisoria
    add_line(s, inches(0.25), inches(2.80), W - inches(0.5))

    # ── 6 ejercicios en 2 filas × 3 columnas ────────────────
    # Diseño: fila superior sobre fondo oscuro (con acento verde),
    #         fila inferior sobre fondo vainilla (con acento lila)
    ejercicios = [
        ("01", "Sentadillas",          "3 × 12-15 rep", "Piernas · Glúteos",   PISTACHO),
        ("02", "Puente de glúteos",    "3 × 15-20 rep", "Glúteos · Femorales", PISTACHO),
        ("03", "Flexiones",            "3 × 10-12 rep", "Pecho · Tríceps",     PISTACHO),
        ("04", "Remo con botellas",    "3 × 12-15 rep", "Espalda · Bíceps",    LILA),
        ("05", "Zancadas alternas",    "3 × 10 rep/lado","Piernas · Equilibrio",LILA),
        ("06", "Plancha abdominal",    "3 × 30-45 seg",  "Core · Estabilidad",  LILA),
    ]

    CARD_W = (W - inches(1.0)) / 3
    GAP    = inches(0.12)

    for i, (num, nombre_e, series, musculo, acento) in enumerate(ejercicios):
        col = i % 3
        row = i // 3
        x = inches(0.35) + col * (CARD_W + GAP)

        if row == 0:
            # Fila superior: sobre fondo oscuro
            y = inches(2.90)
            CARD_H = inches(1.90)
            card_fill  = RGBColor(0x48, 0x52, 0x46)
            card_line  = PISTACHO
            txt_color  = BLANCO
            sub_color  = RGBColor(0xA8, 0xC0, 0xA0)
        else:
            # Fila inferior: sobre fondo vainilla — altura balanceada
            y = inches(4.85)
            CARD_H = inches(5.50)
            card_fill  = BLANCO
            card_line  = GRIS_SOFT
            txt_color  = OSCURO
            sub_color  = GRIS

        add_rect(s, x, y, CARD_W, CARD_H,
                 fill_color=card_fill, line_color=card_line, line_width=Pt(0.8))

        # Número con accent color
        add_textbox(s, num, x + Pt(8), y + Pt(6), Pt(30), Pt(22),
                    font_name=SERIF, font_size=18, italic=True, bold=True,
                    color=acento, word_wrap=False)

        # Nombre del ejercicio
        add_textbox(s, nombre_e, x + Pt(8), y + Pt(30),
                    CARD_W - Pt(16), Pt(28),
                    font_name=SERIF, font_size=13.5, bold=True,
                    color=txt_color, word_wrap=True)

        # Foco muscular en fila superior también
        if row == 0:
            add_textbox(s, musculo, x + Pt(8), y + Pt(60),
                        CARD_W - Pt(16), Pt(14),
                        font_name=SERIF, font_size=8.5, italic=True,
                        color=sub_color, word_wrap=False)

        # Series — pill badge
        badge_y = y + (Pt(80) if row == 0 else Pt(62))
        add_rect(s, x + Pt(8), badge_y, CARD_W - Pt(16), Pt(20),
                 fill_color=acento)
        add_textbox(s, series, x + Pt(10), badge_y + Pt(4),
                    CARD_W - Pt(20), Pt(14),
                    font_name=SANS, font_size=9.5, bold=True,
                    color=BLANCO, word_wrap=False)

        # Tip para ambas filas
        tips_ej = {
            "Sentadillas":         "Pies al ancho de hombros. Baja hasta muslos paralelos al piso.",
            "Puente de glúteos":   "Pies apoyados, rodillas dobladas. Aprieta glúteos al subir.",
            "Flexiones":           "Manos bajo hombros, cuerpo en tabla. Baja hasta rozar el piso.",
            "Remo con botellas":   "Inclínate 45°, espalda recta. Jala los codos hacia atrás.",
            "Zancadas alternas":   "Rodilla trasera cerca del piso. Pecho erguido siempre.",
            "Plancha abdominal":   "Cuerpo en línea recta. Aprieta abdomen y respira.",
        }
        tip = tips_ej.get(nombre_e, "")

        if row == 0 and tip:
            # Tip después del badge (que está en y+Pt(80)) — va en la parte baja de la card
            add_line(s, x + Pt(8), y + Pt(108), CARD_W - Pt(16))
            add_textbox(s, tip, x + Pt(8), y + Pt(116),
                        CARD_W - Pt(16), Pt(50),
                        font_name=SANS, font_size=8.5,
                        color=RGBColor(0xB8, 0xC8, 0xB0), word_wrap=True)

        if row == 1:
            # Fila inferior: músculo + tip
            add_textbox(s, musculo, x + Pt(8), y + Pt(112),
                        CARD_W - Pt(16), Pt(16),
                        font_name=SERIF, font_size=9.5, italic=True,
                        color=sub_color, word_wrap=False)
            if tip:
                add_line(s, x + Pt(8), y + Pt(134), CARD_W - Pt(16))
                add_textbox(s, tip, x + Pt(8), y + Pt(142),
                            CARD_W - Pt(16), Pt(42),
                            font_name=SANS, font_size=9, color=GRIS, word_wrap=True)

    # ── Banner inferior con los 3 pilares ───────────────────
    banner_y = H - inches(0.85)
    add_rect(s, 0, banner_y, W, inches(0.85), fill_color=PISTACHO)

    pilares = [
        ("ANTES", "Calienta 5-7 min"),
        ("DURANTE", "Hidratate bien"),
        ("DESPUÉS", "Descansa 60-90 seg"),
    ]
    pw = W / 3
    for i, (titulo_p, desc_p) in enumerate(pilares):
        px = i * pw
        if i > 0:
            add_rect(s, px, banner_y + Pt(10), Pt(0.5), Pt(30),
                     fill_color=RGBColor(0x8E,0xA8,0x7A))
        add_textbox(s, titulo_p, px, banner_y + Pt(8), pw, Pt(14),
                    font_size=7, color=RGBColor(0xD4,0xE8,0xD0), font_name=SANS,
                    bold=True, align=PP_ALIGN.CENTER)
        add_textbox(s, desc_p, px, banner_y + Pt(22), pw, Pt(18),
                    font_size=10, color=BLANCO, font_name=SERIF,
                    italic=True, align=PP_ALIGN.CENTER)

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

    # IDs de imagen para cada receta (usar platos del catálogo que coincidan visualmente)
    receta_img_ids = ["AC004", "D002", "D006"]  # hamburguesas, pan, brownie aprox

    for i, (rid, nombre_r, sub, ings, prep) in enumerate(recetas):
        # Cada receta en su propia slide
        if i > 0:
            s = new_slide(prs)
            add_textbox(s, "A N E X O", inches(0.5), inches(0.38),
                        inches(2), Pt(14), font_size=9, color=LILA)
            section_header(s, 18, "Recetas", "preparaciones simples para tu día a día")

        # Foto de la receta — grande, arriba a la derecha
        IMG_R_W = inches(3.0)
        IMG_R_H = inches(2.20)
        img_rx = W - IMG_R_W - inches(0.50)
        img_ry = inches(2.55)
        if i < len(receta_img_ids):
            add_image_safe(s, {"id": receta_img_ids[i]}, img_rx, img_ry, IMG_R_W, IMG_R_H)

        # Etiqueta
        add_textbox(s, rid, inches(0.55), inches(2.55), inches(2), Pt(14),
                    font_size=8, color=PISTACHO, font_name=SANS)

        # Nombre en grande — word_wrap=False para no partirse
        add_textbox(s, nombre_r, inches(0.55), inches(2.72),
                    W - IMG_R_W - inches(1.10), Pt(34),
                    font_size=22, italic=True, font_name=SERIF,
                    color=OSCURO, word_wrap=False)

        # Subtítulo
        add_textbox(s, sub, inches(0.55), inches(3.30),
                    W - IMG_R_W - inches(1.10), Pt(16),
                    font_size=9, italic=True, color=GRIS, font_name=SANS)

        add_line(s, inches(0.55), inches(3.58), W - inches(1.10))

        # Ingredientes (col izq)
        add_textbox(s, "INGREDIENTES", inches(0.55), inches(3.68),
                    inches(3.5), Pt(14), font_size=7.5, color=GRIS,
                    font_name=SANS, bold=True)
        for j, ing in enumerate(ings):
            add_textbox(s, f"— {ing}",
                        inches(0.55), inches(3.86) + j * inches(0.32),
                        inches(3.5), Pt(20),
                        font_size=11, color=OSCURO, font_name=SANS)

        # Preparación (col der)
        add_textbox(s, "PREPARACIÓN", inches(4.35), inches(3.68),
                    W - inches(4.85), Pt(14), font_size=7.5, color=GRIS,
                    font_name=SANS, bold=True)
        add_textbox(s, prep, inches(4.35), inches(3.86),
                    W - inches(4.85), H - inches(4.80),
                    font_size=11, color=OSCURO, font_name=SANS, word_wrap=True)

        footer(s, f"Recetas — {rid}", pnum); pnum += 1

    # Slide de cierre de recetas con Instagram
    s = new_slide(prs)
    add_line(s, inches(0.55), H / 2 - inches(0.5), W - inches(1.1))
    add_textbox(s, "Más recetas en Instagram", inches(0.5), H / 2 - inches(0.3),
                W - inches(1.0), Pt(30), font_size=22, italic=True,
                font_name=SERIF, color=OSCURO, align=PP_ALIGN.CENTER)
    add_textbox(s, "@nutrylife_2.0", inches(0.5), H / 2 + inches(0.35),
                W - inches(1.0), Pt(42), font_size=32, bold=True,
                font_name=SERIF, color=LILA, align=PP_ALIGN.CENTER)
    add_line(s, inches(0.55), H / 2 + inches(0.9), W - inches(1.1))
    footer(s, "Recetas — Anexo", pnum); pnum += 1

    # ════════════════════════════════════════════════════════
    # SLIDE FINAL — CIERRE
    # ════════════════════════════════════════════════════════
    s = new_slide(prs)   # fondo vainilla claro como el resto

    # Acento pistacho vertical izquierdo
    add_rect(s, 0, 0, Pt(5), H, fill_color=PISTACHO)

    add_textbox(s, "UN ÚLTIMO RECORDATORIO", inches(0.25), inches(0.55),
                W - inches(0.5), Pt(16), font_size=7.5, color=GRIS,
                font_name=SANS, align=PP_ALIGN.CENTER)
    add_line(s, inches(0.65), inches(0.85), W - inches(1.30))

    # Frase principal en oscuro (no blanco sobre fondo claro)
    txb = s.shapes.add_textbox(inches(0.5), inches(1.8), W - inches(1.0), inches(4.5))
    tf = txb.text_frame
    tf.word_wrap = True
    p_cierre = tf.paragraphs[0]
    p_cierre.alignment = PP_ALIGN.CENTER
    run = p_cierre.add_run()
    run.text = "Lograremos\ntu mejor\nversión."
    run.font.name = SERIF
    run.font.size = Pt(60)
    run.font.italic = True
    run.font.color.rgb = OSCURO

    # Línea pistacho decorativa bajo el texto
    add_rect(s, inches(2.5), inches(6.55), inches(3.5), Pt(2.5), fill_color=PISTACHO)

    add_line(s, inches(0.65), inches(7.2), W - inches(1.30))
    add_textbox(s, "SEGUIMOS EN CONTACTO", inches(0.5), inches(7.30),
                W - inches(1.0), Pt(16), font_size=7, color=GRIS,
                font_name=SANS, align=PP_ALIGN.CENTER)
    add_textbox(s, "Myriam Márquez García", inches(0.5), inches(7.55),
                W - inches(1.0), Pt(30), font_size=22, italic=True,
                font_name=SERIF, color=OSCURO, align=PP_ALIGN.CENTER)
    add_textbox(s, "nutricion.metodo@gmail.com   ·   @nutrylife_2.0",
                inches(0.5), inches(8.10), W - inches(1.0), Pt(18),
                font_size=10, italic=True, color=GRIS,
                font_name=SANS, align=PP_ALIGN.CENTER)

    # ── Guardar ─────────────────────────────────────────────
    prs.save(output_path)
    return output_path
