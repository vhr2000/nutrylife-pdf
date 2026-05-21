"""
Nutrylife — pdf_generator.py
Genera el PDF personalizado a partir de los datos del formulario de Mimi.
Versión adaptada para correr en Render.com (sin fuentes locales).
"""

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib.colors import HexColor
from reportlab.pdfgen import canvas
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, PageBreak,
    Frame, PageTemplate, Table, TableStyle,
    NextPageTemplate, HRFlowable
)
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY
import os

# ============================================================
# PALETA
# ============================================================
VAINILLA       = HexColor("#F5EFE0")
VAINILLA_SOFT  = HexColor("#FAF6EC")
VAINILLA_DARK  = HexColor("#EBE5D5")
PISTACHO       = HexColor("#A8C49A")
PISTACHO_DARK  = HexColor("#6B8E5A")
PISTACHO_SOFT  = HexColor("#D4E0CB")
LILA           = HexColor("#B8A5D1")
LILA_DARK      = HexColor("#7A5FA0")
TEXTO          = HexColor("#3A3A3A")
TEXTO_SOFT     = HexColor("#6B6B6B")
TEXTO_MUY_SOFT = HexColor("#9A9A9A")
LINEA          = HexColor("#D9D2C0")
LINEA_SOFT     = HexColor("#E8E2D2")

# ============================================================
# FUENTES — fallback a Helvetica/Times (siempre disponibles)
# Render no tiene Lora/Poppins instaladas, usamos las built-in
# ============================================================
FONT_SERIF      = "Times-Roman"
FONT_SERIF_ITAL = "Times-Italic"
FONT_SERIF_BOLD = "Times-Bold"
FONT_SANS       = "Helvetica"
FONT_SANS_BOLD  = "Helvetica-Bold"
FONT_SANS_ITAL  = "Helvetica-Oblique"


# ============================================================
# ESTILOS
# ============================================================
def make_styles():
    s = {}
    s["body"] = ParagraphStyle(
        "body", fontName=FONT_SANS, fontSize=10, leading=15,
        textColor=TEXTO, alignment=TA_LEFT, spaceAfter=6,
    )
    s["body_just"] = ParagraphStyle(
        "body_just", parent=s["body"], alignment=TA_JUSTIFY,
    )
    s["body_center"] = ParagraphStyle(
        "body_center", parent=s["body"], alignment=TA_CENTER,
    )
    s["h1"] = ParagraphStyle(
        "h1", fontName=FONT_SERIF_ITAL, fontSize=30, leading=34,
        textColor=TEXTO, alignment=TA_LEFT, spaceAfter=8,
    )
    s["h2"] = ParagraphStyle(
        "h2", fontName=FONT_SERIF, fontSize=16, leading=20,
        textColor=TEXTO, alignment=TA_LEFT, spaceAfter=6,
    )
    s["h3"] = ParagraphStyle(
        "h3", fontName=FONT_SANS_BOLD, fontSize=11, leading=14,
        textColor=TEXTO, alignment=TA_LEFT, spaceAfter=4,
    )
    s["eyebrow"] = ParagraphStyle(
        "eyebrow", fontName=FONT_SANS, fontSize=8, leading=11,
        textColor=LILA_DARK, alignment=TA_LEFT, spaceAfter=4,
    )
    s["subtitle"] = ParagraphStyle(
        "subtitle", fontName=FONT_SERIF_ITAL, fontSize=10, leading=14,
        textColor=TEXTO_SOFT, alignment=TA_LEFT, spaceAfter=8,
    )
    s["label"] = ParagraphStyle(
        "label", fontName=FONT_SANS, fontSize=8, leading=11,
        textColor=TEXTO_SOFT, alignment=TA_LEFT, spaceAfter=2,
    )
    s["metric"] = ParagraphStyle(
        "metric", fontName=FONT_SERIF_ITAL, fontSize=28, leading=32,
        textColor=TEXTO, alignment=TA_LEFT,
    )
    s["small"] = ParagraphStyle(
        "small", fontName=FONT_SANS, fontSize=8, leading=11,
        textColor=TEXTO_MUY_SOFT, alignment=TA_LEFT,
    )
    s["italic"] = ParagraphStyle(
        "italic", fontName=FONT_SERIF_ITAL, fontSize=10, leading=15,
        textColor=TEXTO, alignment=TA_LEFT, spaceAfter=6,
    )
    s["italic_center"] = ParagraphStyle(
        "italic_center", parent=s["italic"], alignment=TA_CENTER,
    )
    s["quote"] = ParagraphStyle(
        "quote", fontName=FONT_SERIF_ITAL, fontSize=12, leading=18,
        textColor=TEXTO, alignment=TA_CENTER, spaceBefore=8, spaceAfter=8,
    )
    return s


# ============================================================
# DECORACIÓN DE PÁGINAS
# ============================================================
def draw_cover_bg(c, doc):
    w, h = A4
    c.setFillColor(VAINILLA)
    c.rect(0, 0, w, h, fill=1, stroke=0)
    c.setFillColor(PISTACHO)
    c.setFillAlpha(0.28)
    c.rect(w * 0.5, h * 0.55, w * 0.5, h * 0.45, fill=1, stroke=0)
    c.setFillColor(LILA)
    c.setFillAlpha(0.22)
    c.circle(w * 0.75, h * 0.78, 4.2 * cm, fill=1, stroke=0)
    c.setFillAlpha(1)


def draw_page_bg(c, doc):
    w, h = A4
    c.setFillColor(VAINILLA_SOFT)
    c.rect(0, 0, w, h, fill=1, stroke=0)
    # Decoración sutil
    c.setFillColor(PISTACHO)
    c.setFillAlpha(0.08)
    c.circle(w + 0.5 * cm, h + 0.5 * cm, 4 * cm, fill=1, stroke=0)
    c.setFillAlpha(1)
    # Header
    c.setStrokeColor(LINEA)
    c.setLineWidth(0.5)
    c.line(2 * cm, h - 1.5 * cm, w - 2 * cm, h - 1.5 * cm)
    c.setFont(FONT_SANS, 7.5)
    c.setFillColor(TEXTO_SOFT)
    c.drawString(2 * cm, h - 1.1 * cm, "MYRIAM NUTRICIÓN")
    c.drawRightString(w - 2 * cm, h - 1.1 * cm, f"— {doc.page:02d} —")
    # Footer
    c.line(2 * cm, 1.6 * cm, w - 2 * cm, 1.6 * cm)
    c.setFont(FONT_SERIF_ITAL, 8)
    c.setFillColor(LILA_DARK)
    c.drawCentredString(w / 2, 1.1 * cm, "Logremos tu mejor versión")


def draw_cierre_bg(c, doc):
    w, h = A4
    c.setFillColor(VAINILLA)
    c.rect(0, 0, w, h, fill=1, stroke=0)
    c.setStrokeColor(LILA)
    c.setLineWidth(0.5)
    c.circle(w / 2, h * 0.35, 1.2 * cm, stroke=1, fill=0)
    c.setFillColor(PISTACHO_DARK)
    c.circle(w / 2, h * 0.35, 0.08 * cm, fill=1, stroke=0)


# ============================================================
# HELPERS
# ============================================================
def chapter_header(numero, titulo, subtitulo=None):
    elements = []
    elements.append(Paragraph(
        f"{numero:02d}",
        ParagraphStyle("ch_num", fontName=FONT_SANS, fontSize=9,
                       textColor=TEXTO_SOFT, alignment=TA_LEFT, spaceAfter=2)
    ))
    elements.append(Paragraph(
        titulo,
        ParagraphStyle("ch_h", fontName=FONT_SERIF_ITAL, fontSize=28,
                       textColor=TEXTO, alignment=TA_LEFT, leading=32, spaceAfter=6)
    ))
    elements.append(HRFlowable(width="20%", thickness=0.5, color=LINEA,
                               spaceBefore=2, spaceAfter=10, hAlign="LEFT"))
    if subtitulo:
        elements.append(Paragraph(
            subtitulo,
            ParagraphStyle("ch_sub", fontName=FONT_SERIF_ITAL, fontSize=9.5,
                           textColor=TEXTO_SOFT, alignment=TA_LEFT, spaceAfter=16)
        ))
    return elements


def caja(contenido, color_borde=PISTACHO, color_fondo=VAINILLA, ancho=17 * cm):
    """Caja con borde izquierdo de color."""
    data = [[Paragraph(contenido,
                       ParagraphStyle("caja_p", fontName=FONT_SANS, fontSize=10,
                                      textColor=TEXTO, leading=15))]]
    t = Table(data, colWidths=[ancho])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), color_fondo),
        ("LINEBEFORE", (0, 0), (0, 0), 3, color_borde),
        ("LEFTPADDING", (0, 0), (-1, -1), 14),
        ("RIGHTPADDING", (0, 0), (-1, -1), 14),
        ("TOPPADDING", (0, 0), (-1, -1), 12),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 12),
    ]))
    return t


# ============================================================
# SECCIONES FIJAS (siempre iguales, no dependen del paciente)
# ============================================================

CLAVES = [
    "Come cuando tengas hambre real.",
    "Come hasta sentirte saciado, ni más ni menos.",
    "No te peses todos los días. Tu progreso no cabe en una balanza.",
    "Evita las colaciones. Si tu desayuno y almuerzo son completos, no las vas a necesitar.",
    "No cuentes calorías. La calidad de lo que comes importa más que la cantidad.",
    "No te compares con nadie. Tu proceso es tuyo y lo vas a lograr.",
    "No le tengas miedo a las grasas ni a las proteínas. Témele al exceso de carbohidratos y a los alimentos procesados.",
    "Desayuna lo suficiente para no necesitar nada a media mañana.",
    "Almuerza lo suficiente para no necesitar nada a media tarde.",
    "Hidrátate bien: agua con o sin gas, infusiones, té, café, mate.",
    "Toma tu vitamina C con el desayuno, todos los días.",
    "Toma omega 3 todos los días.",
    "Toma magnesio con la cena o antes de dormir.",
    "Duerme al menos 8 horas. Apaga las pantallas una hora antes de acostarte.",
    "Si cenas, hazlo al menos dos horas antes de dormir.",
    "Si vas a hacer ejercicio, prioriza pesas, máquinas o el peso de tu propio cuerpo.",
    "Toma sol 10 minutos al día. Camina sobre el pasto cuando puedas.",
    "Confía en el proceso. Los cambios reales toman tiempo, pero llegan.",
]

SUPLEMENTOS_INFO = {
    "magnesio": {
        "nombre": "Magnesio Citrato",
        "tagline": "el regulador del cuerpo",
        "texto": (
            "El magnesio interviene en más de 300 procesos del cuerpo. Es clave en la producción "
            "de hormonas, la regulación del sistema inmune, el metabolismo, la pérdida de grasa "
            "corporal, el crecimiento muscular, el sueño y el estado de ánimo."
        ),
        "sintomas": "Insomnio, ansiedad, migrañas, arritmias, presión alta, fatiga crónica, calambres, hipotiroidismo.",
        "dosis": "400mg, 2 comprimidos al día.",
        "horario": "1 en la mañana, 1 antes de dormir.",
        "marca": "Pure Science Magnesio Citrato Quelado",
    },
    "omega3": {
        "nombre": "Omega 3",
        "tagline": "antiinflamatorio celular",
        "texto": (
            "Los ácidos grasos esenciales son antiinflamatorios potentes. Mejoran la salud cerebral, "
            "la memoria, el estado de ánimo, la calidad de la piel y el cabello. Apoyan el desarrollo "
            "y mantención de la masa muscular."
        ),
        "sintomas": "Especialmente importante en ansiedad, depresión, déficit atencional y deportistas.",
        "dosis": "700mg, 1 cápsula al día.",
        "horario": "Con el desayuno o almuerzo.",
        "marca": "My OmegaRed (incluye astaxantina + colina)",
    },
    "vitamina_d": {
        "nombre": "Vitamina D",
        "tagline": "el sol que nos falta",
        "texto": (
            "En Chile, por nuestra latitud, prácticamente toda la población tiene déficit. "
            "La vitamina D regula el sistema inmune, la salud ósea, el ánimo y la energía. "
            "Es uno de los suplementos más importantes que puedes tomar."
        ),
        "sintomas": "Fatiga persistente, baja inmunidad, dolor en huesos, debilidad muscular, depresión leve.",
        "dosis": "5.000 UI diarias durante 2 a 3 meses.",
        "horario": "Con el desayuno (mejor absorción con grasa).",
        "marca": "NOW Vitamin D-3 5000 IU",
    },
    "vitamina_c": {
        "nombre": "Vitamina C Liposomal",
        "tagline": "antioxidante y reparador",
        "texto": (
            "Antioxidante potente, refuerza el sistema inmune, mejora la cicatrización y retrasa "
            "el envejecimiento de la piel. La versión liposomal se absorbe mucho mejor que la "
            "vitamina C convencional."
        ),
        "sintomas": "Anemia, hematomas frecuentes, mala cicatrización, envejecimiento prematuro de la piel.",
        "dosis": "1 cápsula al día.",
        "horario": "Con el desayuno.",
        "marca": "Wellplus Vitamina C Plus Liposomal",
    },
    "electrolitos": {
        "nombre": "Electrolitos",
        "tagline": "lo que el agua sola no repone",
        "texto": (
            "En una alimentación baja en carbohidratos, el cuerpo elimina más sodio, potasio "
            "y magnesio. Reponer electrolitos previene los síntomas de adaptación. "
            "Si seguimos un plan sin carbohidratos, este suplemento no es opcional."
        ),
        "sintomas": "Dolor de cabeza, fatiga, calambres, estreñimiento, neblina mental, ansiedad inexplicable.",
        "dosis": "1 cucharadita de sal de mar en 500cc de agua con jugo de medio limón.",
        "horario": "Durante el día, especialmente en la mañana.",
        "marca": "ELEC-ViTaL o receta casera con sal de mar + limón",
    },
}

DESAYUNOS = [
    ("Omelette mediterráneo",
     "3 huevos batidos con espinaca, champiñones y queso de cabra. Aceite de oliva, sal de mar, una pizca de cúrcuma."),
    ("Bowl de palta y huevo",
     "Media palta machacada con limón y sal sobre 2 rebanadas de pan low carb. 2 huevos pochados."),
    ("Yogurt natural con berries",
     "Yogurt entero sin azúcar con berries, almendras tostadas y chía. Café o infusión."),
    ("Plato simple",
     "150g de pollo desmenuzado o salmón ahumado, palta, tomate cherry y aceite de oliva."),
    ("Batido proteico",
     "1 medida proteína en polvo + leche de almendras + frambuesas + mantequilla de almendras. Máx. 2 veces/semana."),
    ("Huevos revueltos con jamón",
     "3 huevos revueltos con 2 láminas de jamón artesanal y queso mantecoso. Sin pan."),
    ("Queso chacra con palta",
     "100g de queso chacra, 2 huevos duros, ½ palta, tomate en rodajas y aceite de oliva."),
    ("Tortilla de champiñones",
     "Tortilla de 3 huevos con champiñones, jamón y queso philadelphia. Café sin azúcar."),
]

ALMUERZOS = [
    ("Salmón al horno con espárragos",
     "150g de salmón al horno con limón. Espárragos asados. Hojas verdes con aceite de oliva."),
    ("Pollo al ajillo con coliflor",
     "150g de pechuga al ajillo. Coliflor asada con cúrcuma. Ensalada de tomate y palta."),
    ("Bowl de huevo y vegetales",
     "3 huevos cocidos. Mix de hojas verdes, brócoli salteado, champiñones. Aceite de oliva y limón."),
    ("Hamburguesas caseras",
     "Hamburguesas de carne molida con zapallo italiano. Ensalada de pepino y palta."),
    ("Pescado blanco al wok",
     "150g de reineta o merluza al sartén. Mix de pimentón, zapallo italiano y brócoli salteados."),
    ("Carne mechada con palta",
     "150g de carne mechada. Palta, tomate, aceite de oliva y limón. Solo para almuerzo."),
    ("Pavo plancha con ensalada",
     "150g de pavo a la plancha. Rúcula, berros, tomate cherry, cebolla morada y aceite de oliva."),
    ("Tímbal de salmón con palta",
     "150g de salmón. Pimentón, cebolla, perejil. Servir sobre palta machacada con limón."),
]


# ============================================================
# FUNCIÓN PRINCIPAL
# ============================================================

def generar_pdf_completo(datos: dict, output_path: str):
    """
    Genera el PDF personalizado con los datos del formulario de Mimi.

    datos = {
        'nombre': str,
        'edad': str,
        'fecha': str,
        'peso': str,
        'talla': str,
        'grasa': str (opcional),
        'foco': str,
        'duracion': str,
        'kcal': str,
        'proteinas': str,
        'carbos': str,
        'grasas': str,
        'version_plan': 'con_carbos' | 'sin_carbos',
        'suplementos': list[str],
        'restricciones': list[str],
        'nota_personal': str,
    }
    """

    styles = make_styles()

    doc = SimpleDocTemplate(
        output_path,
        pagesize=A4,
        leftMargin=2 * cm,
        rightMargin=2 * cm,
        topMargin=2.5 * cm,
        bottomMargin=2.2 * cm,
        title=f"Plan Nutricional — {datos.get('nombre', 'Paciente')}",
        author="Myriam Márquez García",
    )

    # Frames
    cover_frame = Frame(2*cm, 2*cm, A4[0]-4*cm, A4[1]-4*cm,
                        leftPadding=0, rightPadding=0,
                        topPadding=0, bottomPadding=0, id="cover")
    normal_frame = Frame(2*cm, 2.2*cm, A4[0]-4*cm, A4[1]-4.7*cm,
                         leftPadding=0, rightPadding=0,
                         topPadding=0, bottomPadding=0, id="normal")
    cierre_frame = Frame(2*cm, 2*cm, A4[0]-4*cm, A4[1]-4*cm,
                         leftPadding=0, rightPadding=0,
                         topPadding=0, bottomPadding=0, id="cierre")

    doc.addPageTemplates([
        PageTemplate(id="cover",  frames=[cover_frame],  onPage=draw_cover_bg),
        PageTemplate(id="normal", frames=[normal_frame], onPage=draw_page_bg),
        PageTemplate(id="cierre", frames=[cierre_frame], onPage=draw_cierre_bg),
    ])

    story = []

    # ----------------------------------------------------------
    # PORTADA
    # ----------------------------------------------------------
    story.append(Spacer(1, 1.2*cm))

    # Marca Myriam — logo tipográfico elegante
    marca_data = [[
        [
            Paragraph("Myriam Márquez",
                      ParagraphStyle("marca1", fontName=FONT_SERIF_ITAL, fontSize=18,
                                     textColor=TEXTO, leading=20, spaceAfter=2)),
            Paragraph("NUTRICIONISTA · LOW CARB",
                      ParagraphStyle("marca2", fontName=FONT_SANS, fontSize=8,
                                     textColor=PISTACHO_DARK, letterSpacing=2.5)),
        ],
        Paragraph("",
                  ParagraphStyle("marca3", fontName=FONT_SANS, fontSize=9))
    ]]
    mt_marca = Table(marca_data, colWidths=[10*cm, 7*cm])
    mt_marca.setStyle(TableStyle([
        ("LEFTPADDING", (0,0), (-1,-1), 0),
        ("RIGHTPADDING", (0,0), (-1,-1), 0),
        ("VALIGN", (0,0), (-1,-1), "BOTTOM"),
    ]))
    story.append(mt_marca)
    story.append(Spacer(1, 9*cm))

    story.append(Paragraph("Plan Nutricional",
        ParagraphStyle("cov_h", fontName=FONT_SERIF_ITAL, fontSize=48,
                       textColor=TEXTO, leading=52, alignment=TA_LEFT, spaceAfter=14)))

    story.append(Spacer(1, 1.5*cm))
    story.append(HRFlowable(width="100%", thickness=0.5, color=LINEA,
                             spaceBefore=0, spaceAfter=12))

    # Datos del paciente en 3 columnas
    nombre = datos.get("nombre", "—")
    duracion = datos.get("duracion", "14 días")
    fecha = datos.get("fecha", "")

    cover_data = [[
        [
            Paragraph("PREPARADO PARA",
                      ParagraphStyle("cl1", fontName=FONT_SANS, fontSize=7,
                                     textColor=TEXTO_SOFT, spaceAfter=3, letterSpacing=1)),
            Paragraph(nombre,
                      ParagraphStyle("cv1", fontName=FONT_SANS_BOLD, fontSize=12, textColor=TEXTO))
        ],
        [
            Paragraph("DURACIÓN",
                      ParagraphStyle("cl2", fontName=FONT_SANS, fontSize=7,
                                     textColor=TEXTO_SOFT, spaceAfter=3, letterSpacing=1)),
            Paragraph(duracion,
                      ParagraphStyle("cv2", fontName=FONT_SANS_BOLD, fontSize=12, textColor=TEXTO))
        ],
        [
            Paragraph("FECHA DE INICIO",
                      ParagraphStyle("cl3", fontName=FONT_SANS, fontSize=7,
                                     textColor=TEXTO_SOFT, spaceAfter=3, letterSpacing=1)),
            Paragraph(fecha if fecha else "—",
                      ParagraphStyle("cv3", fontName=FONT_SANS_BOLD, fontSize=12, textColor=TEXTO))
        ],
    ]]
    ct = Table(cover_data, colWidths=[5.5*cm, 5.5*cm, 6*cm])
    ct.setStyle(TableStyle([
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("LEFTPADDING", (0,0), (-1,-1), 0),
        ("TOPPADDING", (0,0), (-1,-1), 12),
        ("BOTTOMPADDING", (0,0), (-1,-1), 12),
        ("LINEBELOW", (0,0), (-1,-1), 0.5, LINEA),
    ]))
    story.append(ct)
    story.append(Spacer(1, 0.6*cm))
    story.append(Paragraph("MYRIAM MÁRQUEZ GARCÍA — NUTRICIONISTA",
        ParagraphStyle("dfin", fontName=FONT_SANS, fontSize=8,
                       textColor=TEXTO_SOFT, letterSpacing=1.5)))

    story.append(PageBreak())
    story.append(NextPageTemplate("normal"))

    # ----------------------------------------------------------
    # SOBRE TI
    # ----------------------------------------------------------
    story.extend(chapter_header(1, "Sobre ti", "tu punto de partida"))

    edad  = datos.get("edad", "—")
    peso  = datos.get("peso", "—")
    talla = datos.get("talla", "—")
    grasa = datos.get("grasa", "")
    foco  = datos.get("foco", "—")

    datos_data = [[
        [
            Paragraph("DATOS", ParagraphStyle("dl", fontName=FONT_SANS, fontSize=8,
                      textColor=TEXTO_SOFT, spaceAfter=8, letterSpacing=1.2)),
            Paragraph(f"<b>{nombre}</b>",
                      ParagraphStyle("dn", fontName=FONT_SANS, fontSize=10, textColor=TEXTO, spaceAfter=4)),
            Paragraph(f"Edad · {edad} años",
                      ParagraphStyle("dm1", fontName=FONT_SANS, fontSize=9, textColor=TEXTO_SOFT, spaceAfter=3)),
            Paragraph(f"Foco · {foco}",
                      ParagraphStyle("dm2", fontName=FONT_SANS, fontSize=9, textColor=TEXTO_SOFT)),
        ],
        [
            Paragraph("EVALUACIÓN INICIAL",
                      ParagraphStyle("el", fontName=FONT_SANS, fontSize=8,
                                     textColor=TEXTO_SOFT, spaceAfter=10, letterSpacing=1.2)),
            Table([[
                [
                    Paragraph(f"<font size='20'>{peso}</font>",
                              ParagraphStyle("pv", fontName=FONT_SERIF_ITAL, fontSize=20,
                                             textColor=TEXTO, leading=22)),
                    Paragraph("kg / PESO",
                              ParagraphStyle("pl", fontName=FONT_SANS, fontSize=7,
                                             textColor=TEXTO_SOFT, letterSpacing=1))
                ],
                [
                    Paragraph(f"<font size='20'>{talla}</font>",
                              ParagraphStyle("tv", fontName=FONT_SERIF_ITAL, fontSize=20,
                                             textColor=TEXTO, leading=22)),
                    Paragraph("cm / TALLA",
                              ParagraphStyle("tl", fontName=FONT_SANS, fontSize=7,
                                             textColor=TEXTO_SOFT, letterSpacing=1))
                ],
            ]], colWidths=[3.5*cm, 4*cm],
               style=TableStyle([
                   ("VALIGN", (0,0), (-1,-1), "TOP"),
                   ("LEFTPADDING", (0,0), (-1,-1), 0),
                   ("BOTTOMPADDING", (0,0), (-1,-1), 8),
               ])),
        ]
    ]]
    dt = Table(datos_data, colWidths=[8*cm, 9*cm])
    dt.setStyle(TableStyle([
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("LEFTPADDING", (0,0), (-1,-1), 0),
    ]))
    story.append(dt)

    story.append(Spacer(1, 0.6*cm))
    story.append(HRFlowable(width="100%", thickness=0.4, color=LINEA,
                            spaceBefore=0, spaceAfter=16))

    # ----------------------------------------------------------
    # REQUERIMIENTOS NUTRICIONALES — misma página que Sobre ti
    # ----------------------------------------------------------
    story.extend(chapter_header(2, "Tus requerimientos", "los números que orientan tu día"))

    kcal     = datos.get("kcal", "—")
    proteinas = datos.get("proteinas", "—")
    carbos   = datos.get("carbos", "—")
    grasas   = datos.get("grasas", "—")
    version  = datos.get("version_plan", "con_carbos")
    version_txt = "Low Carb Flexible" if version == "con_carbos" else "Keto / Sin carbohidratos"

    macros_data = [[
        [
            Paragraph("PROTEÍNAS",
                      ParagraphStyle("ml1", fontName=FONT_SANS_BOLD, fontSize=8,
                                     textColor=LILA_DARK, spaceAfter=4, letterSpacing=1.2)),
            Paragraph(f'<font size="36">{proteinas}</font><font size="14">g</font>',
                      ParagraphStyle("mv1", fontName=FONT_SERIF_ITAL, fontSize=36,
                                     textColor=TEXTO, leading=40, spaceAfter=4)),
            Paragraph("base de cada comida",
                      ParagraphStyle("md1", fontName=FONT_SERIF_ITAL, fontSize=8,
                                     textColor=TEXTO_SOFT)),
        ],
        [
            Paragraph("CARBOHIDRATOS",
                      ParagraphStyle("ml2", fontName=FONT_SANS_BOLD, fontSize=8,
                                     textColor=LILA_DARK, spaceAfter=4, letterSpacing=1.2)),
            Paragraph(f'<font size="36">{carbos}</font><font size="14">g</font>',
                      ParagraphStyle("mv2", fontName=FONT_SERIF_ITAL, fontSize=36,
                                     textColor=TEXTO, leading=40, spaceAfter=4)),
            Paragraph("estratégicos, post-entreno",
                      ParagraphStyle("md2", fontName=FONT_SERIF_ITAL, fontSize=8,
                                     textColor=TEXTO_SOFT)),
        ],
        [
            Paragraph("GRASAS",
                      ParagraphStyle("ml3", fontName=FONT_SANS_BOLD, fontSize=8,
                                     textColor=LILA_DARK, spaceAfter=4, letterSpacing=1.2)),
            Paragraph(f'<font size="36">{grasas}</font><font size="14">g</font>',
                      ParagraphStyle("mv3", fontName=FONT_SERIF_ITAL, fontSize=36,
                                     textColor=TEXTO, leading=40, spaceAfter=4)),
            Paragraph("fuentes naturales",
                      ParagraphStyle("md3", fontName=FONT_SERIF_ITAL, fontSize=8,
                                     textColor=TEXTO_SOFT)),
        ],
    ]]
    mt = Table(macros_data, colWidths=[5.6*cm, 5.6*cm, 5.6*cm])
    mt.setStyle(TableStyle([
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("LEFTPADDING", (0,0), (-1,-1), 0),
        ("LINEAFTER", (0,0), (-2,-1), 0.5, LINEA),
    ]))
    story.append(mt)
    story.append(Spacer(1, 0.5*cm))

    extras_data = [[
        [
            Paragraph("CALORÍAS", ParagraphStyle("e1l", fontName=FONT_SANS, fontSize=8,
                      textColor=TEXTO_SOFT, alignment=TA_CENTER, letterSpacing=1.5, spaceAfter=6)),
            Paragraph(f"<font size='18'>{kcal}</font>",
                      ParagraphStyle("e1v", fontName=FONT_SERIF_ITAL, fontSize=18,
                                     textColor=TEXTO, alignment=TA_CENTER, spaceAfter=2)),
            Paragraph("kcal / día",
                      ParagraphStyle("e1u", fontName=FONT_SERIF_ITAL, fontSize=8,
                                     textColor=TEXTO_SOFT, alignment=TA_CENTER)),
        ],
        [
            Paragraph("VERSIÓN",
                      ParagraphStyle("e2l", fontName=FONT_SANS, fontSize=8,
                                     textColor=TEXTO_SOFT, alignment=TA_CENTER,
                                     letterSpacing=1.5, spaceAfter=6)),
            Paragraph(version_txt,
                      ParagraphStyle("e2v", fontName=FONT_SERIF_ITAL, fontSize=11,
                                     textColor=PISTACHO_DARK, alignment=TA_CENTER,
                                     leading=14, spaceAfter=2)),
        ],
    ]]
    et = Table(extras_data, colWidths=[8.5*cm, 8.5*cm])
    et.setStyle(TableStyle([
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("TOPPADDING", (0,0), (-1,-1), 12),
        ("BOTTOMPADDING", (0,0), (-1,-1), 12),
        ("LINEABOVE", (0,0), (-1,0), 0.5, LINEA),
        ("LINEBELOW", (0,0), (-1,-1), 0.5, LINEA),
    ]))
    story.append(et)

    story.append(PageBreak())

    # ----------------------------------------------------------
    # LAS CLAVES
    # ----------------------------------------------------------
    story.extend(chapter_header(3, "Las claves", "principios de 15 años de experiencia clínica"))
    story.append(Paragraph(
        "Estas no son reglas rígidas. Son la base sobre la que se construye todo lo demás.",
        ParagraphStyle("cli", fontName=FONT_SANS, fontSize=9.5, textColor=TEXTO_SOFT,
                       alignment=TA_LEFT, spaceAfter=14, leading=14)
    ))

    rows = []
    for i in range(0, 18, 2):
        def cell(num, text):
            return [
                Paragraph(f"{num:02d}",
                          ParagraphStyle("cn", fontName=FONT_SERIF_ITAL, fontSize=16,
                                         textColor=PISTACHO_DARK, leading=18)),
                Paragraph(text,
                          ParagraphStyle("ct", fontName=FONT_SANS, fontSize=9,
                                         textColor=TEXTO, leading=13, spaceAfter=6))
            ]
        left_text  = CLAVES[i] if i < len(CLAVES) else ""
        right_text = CLAVES[i+1] if (i+1) < len(CLAVES) else ""
        rows.append([cell(i+1, left_text), cell(i+2, right_text) if right_text else [Paragraph("", styles["body"])]])

    t = Table(rows, colWidths=[8.5*cm, 8.5*cm])
    t.setStyle(TableStyle([
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("LEFTPADDING", (0,0), (-1,-1), 4),
        ("RIGHTPADDING", (0,0), (-1,-1), 10),
        ("TOPPADDING", (0,0), (-1,-1), 3),
        ("BOTTOMPADDING", (0,0), (-1,-1), 3),
    ]))
    story.append(t)
    story.append(PageBreak())

    # ----------------------------------------------------------
    # MENÚ 7 DÍAS CON MACROS CALCULADOS
    # ----------------------------------------------------------
    from menu_engine import seleccionar_menu_7_dias

    menu_calculado = seleccionar_menu_7_dias(datos)

    story.extend(chapter_header(4, "Tu menú semanal", "7 días con gramajes calculados para tu plan"))

    # Nota clave
    story.append(caja(
        '<b>Nota clínica:</b> Los gramajes de proteína están calculados exactamente para tu plan. '
        'Las verduras son libres — agrégalas según tu apetito. '
        '<b>Nunca incorporamos carne roja en la cena</b>: su digestión lenta genera vigilia.',
        color_borde=LILA, color_fondo=HexColor("#EFE6F5")
    ))
    story.append(Spacer(1, 0.4*cm))

    def render_comida(label, comida_data, es_almuerzo_cena=False):
        """Renderiza una comida del menú con foto si está disponible."""
        nombre = comida_data["nombre"]
        if es_almuerzo_cena:
            gramaje = comida_data["gramaje"]
            proteina = comida_data["proteina_nombre"]
            ingredientes_fijos = comida_data["ingredientes_fijos"]
            descripcion = f"{gramaje}g de {proteina.lower()}, {ingredientes_fijos}"
        else:
            descripcion = comida_data["ingredientes"]

        imagen_id = comida_data.get("imagen")
        imagen_path = None
        if imagen_id:
            # Buscar la imagen en la carpeta local (cuando corre en Render)
            import os
            posibles = [
                f"/opt/render/project/src/imagenes/{imagen_id}",
                f"/opt/render/project/src/{imagen_id}",
                f"/home/claude/imagenes/{imagen_id}",
            ]
            for p in posibles:
                if os.path.exists(p):
                    imagen_path = p
                    break

        if imagen_path:
            from reportlab.platypus import Image as RLImage
            try:
                img = RLImage(imagen_path, width=3.5*cm, height=3.5*cm)
                img.hAlign = 'RIGHT'
                fila = [[
                    [
                        Paragraph(label,
                                  ParagraphStyle("ct_label2", fontName=FONT_SANS_BOLD, fontSize=8,
                                                 textColor=PISTACHO_DARK, letterSpacing=1.2, spaceAfter=2)),
                        Paragraph(nombre,
                                  ParagraphStyle("ct_nombre2", fontName=FONT_SERIF_ITAL, fontSize=13,
                                                 textColor=TEXTO, leading=16, spaceAfter=3)),
                        Paragraph(descripcion,
                                  ParagraphStyle("ct_desc2", fontName=FONT_SANS, fontSize=9,
                                                 textColor=TEXTO_SOFT, leading=12)),
                    ],
                    img
                ]]
                ct = Table(fila, colWidths=[13*cm, 4*cm])
                ct.setStyle(TableStyle([
                    ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
                    ("LEFTPADDING", (0,0), (0,0), 0),
                    ("RIGHTPADDING", (0,0), (0,0), 0),
                    ("TOPPADDING", (0,0), (-1,-1), 8),
                    ("BOTTOMPADDING", (0,0), (-1,-1), 8),
                    ("LINEBELOW", (0,0), (-1,-1), 0.3, LINEA_SOFT),
                ]))
                story.append(ct)
                return
            except Exception:
                pass

        # Sin foto — solo texto
        story.append(Paragraph(label,
            ParagraphStyle("ct_label", fontName=FONT_SANS_BOLD, fontSize=8,
                           textColor=PISTACHO_DARK, letterSpacing=1.2, spaceAfter=2)))
        story.append(Paragraph(nombre,
            ParagraphStyle("ct_nombre", fontName=FONT_SERIF_ITAL, fontSize=13,
                           textColor=TEXTO, leading=16, spaceAfter=3)))
        story.append(Paragraph(descripcion,
            ParagraphStyle("ct_desc", fontName=FONT_SANS, fontSize=9,
                           textColor=TEXTO_SOFT, leading=12, spaceAfter=0)))
        story.append(HRFlowable(width="100%", thickness=0.3, color=LINEA_SOFT,
                               spaceBefore=8, spaceAfter=8))

    for i, dia_data in enumerate(menu_calculado):
        # Header del día — 50% más grande
        story.append(Paragraph(
            dia_data["dia"].upper(),
            ParagraphStyle("dia_eyebrow", fontName=FONT_SANS_BOLD, fontSize=14,
                          textColor=LILA_DARK, spaceAfter=4, letterSpacing=2)
        ))
        story.append(HRFlowable(width="100%", thickness=0.8, color=LILA,
                               spaceBefore=0, spaceAfter=10))

        render_comida("DESAYUNO", dia_data["desayuno"], es_almuerzo_cena=False)
        render_comida("ALMUERZO", dia_data["almuerzo"], es_almuerzo_cena=True)
        render_comida("CENA", dia_data["cena"], es_almuerzo_cena=True)

        # Total del día con nombre del día
        kcal = float(datos.get("kcal", 0))
        prot = float(datos.get("proteinas", 0))
        carbo = float(datos.get("carbos", 0))
        grasa = float(datos.get("grasas", 0))

        total_data = [[
            Paragraph(f"TOTAL DEL {dia_data['dia'].upper()}",
                      ParagraphStyle("total_lbl", fontName=FONT_SANS_BOLD, fontSize=8,
                                     textColor=TEXTO_SOFT, letterSpacing=1.5)),
            Paragraph(
                f'P {round(prot)}g  ·  C {round(carbo)}g  ·  G {round(grasa)}g  ·  '
                f'<b>{round(kcal)} kcal</b>',
                ParagraphStyle("total_val", fontName=FONT_SANS, fontSize=9.5,
                              textColor=TEXTO, alignment=TA_CENTER)
            )
        ]]
        tt = Table(total_data, colWidths=[4*cm, 13*cm])
        tt.setStyle(TableStyle([
            ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
            ("BACKGROUND", (0,0), (-1,-1), VAINILLA_DARK),
            ("LEFTPADDING", (0,0), (-1,-1), 10),
            ("RIGHTPADDING", (0,0), (-1,-1), 10),
            ("TOPPADDING", (0,0), (-1,-1), 8),
            ("BOTTOMPADDING", (0,0), (-1,-1), 8),
        ]))
        story.append(tt)
        story.append(Spacer(1, 0.5*cm))

        # Salto de página cada 2 días
        if i in [1, 3, 5]:
            story.append(PageBreak())
            if i != 5:
                story.extend(chapter_header(4, "Tu menú semanal", "(continuación)"))

    story.append(PageBreak())

    # ----------------------------------------------------------
    # NOTA PERSONAL (si la hay)
    # ----------------------------------------------------------
    nota_personal = datos.get("nota_personal", "").strip()
    restricciones = datos.get("restricciones", [])
    if nota_personal:
        story.extend(chapter_header(6, "Para ti", "nota personal de Myriam"))
        story.append(caja(
            f'<i>"{nota_personal}"</i>',
            color_borde=LILA, color_fondo=HexColor("#EFE6F5")
        ))
        story.append(PageBreak())
        cap_supl = 7
    else:
        cap_supl = 6

    # ----------------------------------------------------------
    # RESTRICCIONES (si las hay — alerta visible en el PDF)
    # ----------------------------------------------------------
    if restricciones:
        story.extend(chapter_header(cap_supl, "Tu menú personalizado",
                                    "ajustado a tus preferencias alimentarias"))
        rest_str = ", ".join([r.capitalize() for r in restricciones])
        story.append(caja(
            f'<b>Alimentos excluidos de tu plan:</b> {rest_str}.<br/><br/>'
            'Las opciones de desayuno, almuerzo y cena han sido adaptadas para '
            'no incluir estos ingredientes.',
            color_borde=PISTACHO, color_fondo=VAINILLA
        ))
        story.append(PageBreak())
        cap_supl += 1

    # ----------------------------------------------------------
    # SUPLEMENTACIÓN
    # ----------------------------------------------------------
    suplementos_activos = datos.get("suplementos", list(SUPLEMENTOS_INFO.keys()))
    story.extend(chapter_header(cap_supl, "Suplementación clínica",
                                "lo que tu plato, por sí solo, no alcanza a aportar"))
    story.append(Paragraph(
        "Estos suplementos son parte del tratamiento. Las dosis se ajustan en cada control.",
        ParagraphStyle("si", fontName=FONT_SANS, fontSize=9.5, textColor=TEXTO_SOFT,
                       alignment=TA_LEFT, spaceAfter=16, leading=14)
    ))

    for key in suplementos_activos:
        if key not in SUPLEMENTOS_INFO:
            continue
        s = SUPLEMENTOS_INFO[key]

        # Header
        h_data = [[
            Paragraph(f"{list(suplementos_activos).index(key)+1:02d}",
                      ParagraphStyle("sn", fontName=FONT_SERIF_ITAL, fontSize=24,
                                     textColor=PISTACHO_DARK, leading=26)),
            [
                Paragraph(s["nombre"],
                          ParagraphStyle("snom", fontName=FONT_SERIF_ITAL, fontSize=16,
                                         textColor=TEXTO, leading=20, spaceAfter=2)),
                Paragraph(f"<i>{s['tagline']}</i>",
                          ParagraphStyle("stag", fontName=FONT_SERIF_ITAL, fontSize=9,
                                         textColor=TEXTO_SOFT))
            ]
        ]]
        ht = Table(h_data, colWidths=[1.8*cm, 15.2*cm])
        ht.setStyle(TableStyle([
            ("VALIGN", (0,0), (-1,-1), "TOP"),
            ("LEFTPADDING", (0,0), (-1,-1), 0),
            ("BOTTOMPADDING", (0,0), (-1,-1), 6),
        ]))
        story.append(ht)

        # Cuerpo
        b_data = [[
            [
                Paragraph(s["texto"],
                          ParagraphStyle("sb", fontName=FONT_SANS, fontSize=9,
                                         textColor=TEXTO, leading=13,
                                         alignment=TA_JUSTIFY, spaceAfter=6)),
                Paragraph(f'<b>Si te falta:</b> <i>{s["sintomas"]}</i>',
                          ParagraphStyle("ss", fontName=FONT_SANS, fontSize=8.5,
                                         textColor=TEXTO_SOFT, leading=12))
            ],
            [
                Paragraph("DOSIS", ParagraphStyle("dl", fontName=FONT_SANS, fontSize=7,
                          textColor=PISTACHO_DARK, letterSpacing=1.2, spaceAfter=2)),
                Paragraph(s["dosis"],
                          ParagraphStyle("dv", fontName=FONT_SANS_BOLD, fontSize=9.5,
                                         textColor=TEXTO, leading=13, spaceAfter=6)),
                Paragraph("HORARIO", ParagraphStyle("hl", fontName=FONT_SANS, fontSize=7,
                          textColor=PISTACHO_DARK, letterSpacing=1.2, spaceAfter=2)),
                Paragraph(s["horario"],
                          ParagraphStyle("hv", fontName=FONT_SANS, fontSize=9.5,
                                         textColor=TEXTO, leading=13, spaceAfter=6)),
                Paragraph("MARCA", ParagraphStyle("ml", fontName=FONT_SANS, fontSize=7,
                          textColor=PISTACHO_DARK, letterSpacing=1.2, spaceAfter=2)),
                Paragraph(s["marca"],
                          ParagraphStyle("mv", fontName=FONT_SERIF_ITAL, fontSize=9.5,
                                         textColor=LILA_DARK, leading=13))
            ]
        ]]
        bt = Table(b_data, colWidths=[10.5*cm, 6.5*cm])
        bt.setStyle(TableStyle([
            ("VALIGN", (0,0), (-1,-1), "TOP"),
            ("LEFTPADDING", (0,0), (0,0), 0),
            ("RIGHTPADDING", (0,0), (0,0), 14),
            ("LEFTPADDING", (1,0), (1,0), 14),
            ("LINEBEFORE", (1,0), (1,0), 0.3, LINEA),
        ]))
        story.append(bt)
        story.append(Spacer(1, 0.3*cm))
        story.append(HRFlowable(width="100%", thickness=0.3, color=LINEA_SOFT,
                                spaceAfter=0.3*cm))

    story.append(PageBreak())

    # ----------------------------------------------------------
    # ANEXO: TABLA NUTRICIONAL
    # ----------------------------------------------------------
    story.append(NextPageTemplate("normal"))
    story.extend(chapter_header(99, "Tabla nutricional",
                                "valores de referencia por cada 100g o unidad"))

    story.append(Paragraph(
        "Usa esta tabla para entender qué contiene cada alimento de tu plan. "
        "Con práctica, podrás hacer tus propios intercambios sin perder el equilibrio.",
        ParagraphStyle("tn_intro", fontName=FONT_SANS, fontSize=9.5,
                      textColor=TEXTO_SOFT, alignment=TA_LEFT,
                      spaceAfter=16, leading=14)
    ))

    from menu_engine import ALMUERZOS_CENAS as ALIMENTOS_DB

    proteinas_vistas = {}
    for key, plato in ALIMENTOS_DB.items():
        nombre = plato["proteina_nombre"]
        g100 = plato["proteina_g_100g"]
        if nombre not in proteinas_vistas:
            proteinas_vistas[nombre] = g100

    porciones_ref = {
        "Pollo": "150-250g", "Salmon": "150-250g", "Vacuno": "150-250g",
        "Merluza": "150-300g", "Pavo": "150-250g", "Atun en agua": "1-2 latas (85g c/u)",
        "Osobuco": "250-350g", "Reineta": "150-300g", "Cerdo": "150-250g",
    }

    header_row = [
        Paragraph("PROTEINA PRINCIPAL", ParagraphStyle("th", fontName=FONT_SANS_BOLD,
                  fontSize=8, textColor=TEXTO_SOFT, letterSpacing=1)),
        Paragraph("g prot / 100g crudo", ParagraphStyle("th2", fontName=FONT_SANS_BOLD,
                  fontSize=8, textColor=PISTACHO_DARK, letterSpacing=1, alignment=TA_CENTER)),
        Paragraph("Porcion aprox.", ParagraphStyle("th3", fontName=FONT_SANS_BOLD,
                  fontSize=8, textColor=TEXTO_SOFT, letterSpacing=1, alignment=TA_CENTER)),
    ]
    filas = [header_row]

    for nombre, g100 in sorted(proteinas_vistas.items()):
        porcion = porciones_ref.get(nombre, "150-250g")
        fila = [
            Paragraph(nombre, ParagraphStyle("td1", fontName=FONT_SANS, fontSize=9, textColor=TEXTO)),
            Paragraph(f"{g100}g", ParagraphStyle("td2", fontName=FONT_SANS_BOLD, fontSize=9,
                      textColor=PISTACHO_DARK, alignment=TA_CENTER)),
            Paragraph(porcion, ParagraphStyle("td3", fontName=FONT_SANS, fontSize=9,
                      textColor=TEXTO_SOFT, alignment=TA_CENTER)),
        ]
        filas.append(fila)

    t = Table(filas, colWidths=[8*cm, 5*cm, 5*cm])
    estilo = [
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("LEFTPADDING", (0,0), (-1,-1), 8),
        ("RIGHTPADDING", (0,0), (-1,-1), 8),
        ("TOPPADDING", (0,0), (-1,-1), 6),
        ("BOTTOMPADDING", (0,0), (-1,-1), 6),
        ("LINEBELOW", (0,0), (-1,-1), 0.3, LINEA_SOFT),
        ("BACKGROUND", (0,0), (-1,0), VAINILLA_DARK),
        ("LINEBELOW", (0,0), (-1,0), 0.5, LINEA),
    ]
    t.setStyle(TableStyle(estilo))
    story.append(t)
    story.append(Spacer(1, 0.5*cm))

    story.append(Paragraph(
        "Fuente: USDA Food Database / Minsal Chile. Valores aproximados.",
        ParagraphStyle("tn_footer", fontName=FONT_SANS, fontSize=8,
                      textColor=TEXTO_MUY_SOFT, alignment=TA_CENTER,
                      fontStyle="italic")
    ))

    story.append(PageBreak())

    # ----------------------------------------------------------
    # CIERRE
    # ----------------------------------------------------------
    story.append(NextPageTemplate("cierre"))
    story.append(Spacer(1, 1.5*cm))
    story.append(Paragraph("UN ÚLTIMO RECORDATORIO",
        ParagraphStyle("clr", fontName=FONT_SANS, fontSize=7,
                       textColor=TEXTO_SOFT, alignment=TA_LEFT,
                       spaceAfter=10, letterSpacing=1.5)))
    story.append(Paragraph(
        "Lograremos<br/>tu mejor<br/>versión.",
        ParagraphStyle("clh", fontName=FONT_SERIF_ITAL, fontSize=52,
                       textColor=TEXTO, alignment=TA_LEFT, leading=58)
    ))
    story.append(Spacer(1, 3.5*cm))
    story.append(HRFlowable(width="100%", thickness=0.3, color=LINEA,
                             spaceBefore=10, spaceAfter=14))
    story.append(Paragraph("SEGUIMOS EN CONTACTO",
        ParagraphStyle("clc", fontName=FONT_SANS, fontSize=7,
                       textColor=TEXTO_SOFT, alignment=TA_CENTER,
                       spaceAfter=8, letterSpacing=1.5)))
    story.append(Paragraph("Myriam Márquez García",
        ParagraphStyle("cln", fontName=FONT_SERIF_ITAL, fontSize=16,
                       textColor=TEXTO, alignment=TA_CENTER, spaceAfter=6)))
    story.append(Paragraph(
        "<i>nutricion.metodo@gmail.com<br/>WhatsApp: +56 9 9733 2001<br/>@nutrylife.cl</i>",
        ParagraphStyle("clct", fontName=FONT_SERIF_ITAL, fontSize=10,
                       textColor=TEXTO_SOFT, alignment=TA_CENTER, leading=15)
    ))

    doc.build(story)
