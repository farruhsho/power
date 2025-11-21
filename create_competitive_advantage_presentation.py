#!/usr/bin/env python3
"""
Beautiful PowerPoint presentation about Competitive Advantage Theory in Uzbek
20 slides with visual elements and beautiful design
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor

def add_gradient_background(slide, prs, color1, color2=None):
    """Add a colored background to the slide"""
    bg = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        0, 0, prs.slide_width, prs.slide_height
    )
    bg.fill.solid()
    bg.fill.fore_color.rgb = color1
    bg.line.fill.background()
    return bg

def create_title_slide(prs):
    """Slide 1: Title slide"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Background
    add_gradient_background(slide, prs, RGBColor(26, 82, 118))

    # Decorative circles
    for i, (x, y, size, alpha) in enumerate([
        (8, 0.5, 2.5, 0.3),
        (0.5, 5, 2, 0.25),
        (7.5, 5.5, 1.5, 0.2)
    ]):
        circle = slide.shapes.add_shape(
            MSO_SHAPE.OVAL,
            Inches(x), Inches(y), Inches(size), Inches(size)
        )
        circle.fill.solid()
        circle.fill.fore_color.rgb = RGBColor(255, 255, 255)
        circle.line.fill.background()

    # Main title
    title_box = slide.shapes.add_textbox(Inches(1), Inches(2.5), Inches(8), Inches(1.5))
    title_frame = title_box.text_frame
    title_frame.text = "RAQOBAT USTUNLIGI\nNAZARIYASI"
    for para in title_frame.paragraphs:
        para.font.size = Pt(56)
        para.font.bold = True
        para.font.color.rgb = RGBColor(255, 255, 255)
        para.alignment = PP_ALIGN.CENTER

    # Subtitle
    subtitle_box = slide.shapes.add_textbox(Inches(1), Inches(4.5), Inches(8), Inches(0.8))
    subtitle_frame = subtitle_box.text_frame
    subtitle_frame.text = "Biznes strategiyasi va muvaffaqiyat yo'llari"
    subtitle_para = subtitle_frame.paragraphs[0]
    subtitle_para.font.size = Pt(28)
    subtitle_para.font.color.rgb = RGBColor(230, 230, 230)
    subtitle_para.alignment = PP_ALIGN.CENTER

    # Icon
    icon_box = slide.shapes.add_textbox(Inches(4.2), Inches(5.5), Inches(1.6), Inches(1))
    icon_frame = icon_box.text_frame
    icon_frame.text = "🏆"
    icon_para = icon_frame.paragraphs[0]
    icon_para.font.size = Pt(72)
    icon_para.alignment = PP_ALIGN.CENTER

def create_content_slide(prs, title, content_items, bg_color, icon=""):
    """Create a content slide with bullet points"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Background
    add_gradient_background(slide, prs, bg_color)

    # Title bar
    title_bar = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(0.5), Inches(0.5), Inches(9), Inches(1)
    )
    title_bar.fill.solid()
    title_bar.fill.fore_color.rgb = RGBColor(255, 255, 255)
    title_bar.line.fill.background()

    # Title text
    title_box = slide.shapes.add_textbox(Inches(1), Inches(0.65), Inches(8), Inches(0.7))
    title_frame = title_box.text_frame
    title_frame.text = title
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(36)
    title_para.font.bold = True
    title_para.font.color.rgb = bg_color
    title_para.alignment = PP_ALIGN.CENTER

    # Content items
    y_pos = 2
    for i, item in enumerate(content_items):
        # Item box
        item_box = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(1), Inches(y_pos), Inches(8), Inches(0.8)
        )
        item_box.fill.solid()
        item_box.fill.fore_color.rgb = RGBColor(255, 255, 255)
        item_box.line.fill.background()

        # Item text
        text_box = slide.shapes.add_textbox(
            Inches(1.3), Inches(y_pos + 0.15), Inches(7.4), Inches(0.5)
        )
        text_frame = text_box.text_frame
        text_frame.text = f"• {item}"
        text_para = text_frame.paragraphs[0]
        text_para.font.size = Pt(20)
        text_para.font.color.rgb = RGBColor(44, 62, 80)

        y_pos += 1

    # Icon if provided
    if icon:
        icon_box = slide.shapes.add_textbox(Inches(8.5), Inches(6), Inches(1), Inches(1))
        icon_frame = icon_box.text_frame
        icon_frame.text = icon
        icon_para = icon_frame.paragraphs[0]
        icon_para.font.size = Pt(48)

def create_two_column_slide(prs, title, left_title, left_items, right_title, right_items, bg_color):
    """Create a slide with two columns"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Background
    add_gradient_background(slide, prs, bg_color)

    # Title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
    title_frame = title_box.text_frame
    title_frame.text = title
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(40)
    title_para.font.bold = True
    title_para.font.color.rgb = RGBColor(255, 255, 255)
    title_para.alignment = PP_ALIGN.CENTER

    # Left column
    left_box = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(0.5), Inches(1.8), Inches(4.25), Inches(4.5)
    )
    left_box.fill.solid()
    left_box.fill.fore_color.rgb = RGBColor(255, 255, 255)
    left_box.line.fill.background()

    # Left title
    left_title_box = slide.shapes.add_textbox(Inches(0.7), Inches(2), Inches(3.85), Inches(0.5))
    left_title_frame = left_title_box.text_frame
    left_title_frame.text = left_title
    left_title_para = left_title_frame.paragraphs[0]
    left_title_para.font.size = Pt(24)
    left_title_para.font.bold = True
    left_title_para.font.color.rgb = bg_color
    left_title_para.alignment = PP_ALIGN.CENTER

    # Left items
    left_content = slide.shapes.add_textbox(Inches(0.8), Inches(2.7), Inches(3.65), Inches(3.3))
    left_content_frame = left_content.text_frame
    left_content_frame.word_wrap = True
    left_content_frame.text = "\n".join([f"• {item}" for item in left_items])
    for para in left_content_frame.paragraphs:
        para.font.size = Pt(16)
        para.font.color.rgb = RGBColor(44, 62, 80)
        para.space_after = Pt(8)

    # Right column
    right_box = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(5.25), Inches(1.8), Inches(4.25), Inches(4.5)
    )
    right_box.fill.solid()
    right_box.fill.fore_color.rgb = RGBColor(255, 255, 255)
    right_box.line.fill.background()

    # Right title
    right_title_box = slide.shapes.add_textbox(Inches(5.45), Inches(2), Inches(3.85), Inches(0.5))
    right_title_frame = right_title_box.text_frame
    right_title_frame.text = right_title
    right_title_para = right_title_frame.paragraphs[0]
    right_title_para.font.size = Pt(24)
    right_title_para.font.bold = True
    right_title_para.font.color.rgb = bg_color
    right_title_para.alignment = PP_ALIGN.CENTER

    # Right items
    right_content = slide.shapes.add_textbox(Inches(5.55), Inches(2.7), Inches(3.65), Inches(3.3))
    right_content_frame = right_content.text_frame
    right_content_frame.word_wrap = True
    right_content_frame.text = "\n".join([f"• {item}" for item in right_items])
    for para in right_content_frame.paragraphs:
        para.font.size = Pt(16)
        para.font.color.rgb = RGBColor(44, 62, 80)
        para.space_after = Pt(8)

def create_strategy_slide(prs, title, strategies, bg_color):
    """Create a slide with strategy boxes"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Background
    add_gradient_background(slide, prs, bg_color)

    # Title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
    title_frame = title_box.text_frame
    title_frame.text = title
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(40)
    title_para.font.bold = True
    title_para.font.color.rgb = RGBColor(255, 255, 255)
    title_para.alignment = PP_ALIGN.CENTER

    # Strategy boxes
    colors = [
        RGBColor(231, 76, 60),
        RGBColor(52, 152, 219),
        RGBColor(46, 204, 113)
    ]

    y_pos = 2
    for i, (strategy_title, description) in enumerate(strategies):
        # Strategy box
        strategy_box = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(1), Inches(y_pos), Inches(8), Inches(1.3)
        )
        strategy_box.fill.solid()
        strategy_box.fill.fore_color.rgb = colors[i % 3]
        strategy_box.line.fill.background()

        # Strategy title
        stitle_box = slide.shapes.add_textbox(
            Inches(1.3), Inches(y_pos + 0.15), Inches(7.4), Inches(0.4)
        )
        stitle_frame = stitle_box.text_frame
        stitle_frame.text = strategy_title
        stitle_para = stitle_frame.paragraphs[0]
        stitle_para.font.size = Pt(24)
        stitle_para.font.bold = True
        stitle_para.font.color.rgb = RGBColor(255, 255, 255)

        # Description
        desc_box = slide.shapes.add_textbox(
            Inches(1.3), Inches(y_pos + 0.6), Inches(7.4), Inches(0.6)
        )
        desc_frame = desc_box.text_frame
        desc_frame.text = description
        desc_frame.word_wrap = True
        desc_para = desc_frame.paragraphs[0]
        desc_para.font.size = Pt(16)
        desc_para.font.color.rgb = RGBColor(255, 255, 255)

        y_pos += 1.5

def create_diagram_slide(prs, title, center_text, items, bg_color):
    """Create a slide with circular diagram"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Background
    add_gradient_background(slide, prs, bg_color)

    # Title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
    title_frame = title_box.text_frame
    title_frame.text = title
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(40)
    title_para.font.bold = True
    title_para.font.color.rgb = RGBColor(255, 255, 255)
    title_para.alignment = PP_ALIGN.CENTER

    # Center circle
    center_circle = slide.shapes.add_shape(
        MSO_SHAPE.OVAL,
        Inches(3.5), Inches(3), Inches(3), Inches(3)
    )
    center_circle.fill.solid()
    center_circle.fill.fore_color.rgb = RGBColor(255, 255, 255)
    center_circle.line.color.rgb = RGBColor(41, 128, 185)
    center_circle.line.width = Pt(5)

    # Center text
    center_text_box = slide.shapes.add_textbox(Inches(3.8), Inches(4), Inches(2.4), Inches(1))
    center_text_frame = center_text_box.text_frame
    center_text_frame.text = center_text
    center_text_frame.word_wrap = True
    center_para = center_text_frame.paragraphs[0]
    center_para.font.size = Pt(24)
    center_para.font.bold = True
    center_para.font.color.rgb = RGBColor(41, 128, 185)
    center_para.alignment = PP_ALIGN.CENTER

    # Surrounding items
    positions = [
        (0.5, 2.5), (7.5, 2.5), (0.5, 5), (7.5, 5)
    ]

    for i, (item, (x, y)) in enumerate(zip(items, positions)):
        # Item box
        item_box = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(x), Inches(y), Inches(2.2), Inches(0.7)
        )
        item_box.fill.solid()
        item_box.fill.fore_color.rgb = RGBColor(241, 196, 15)
        item_box.line.fill.background()

        # Item text
        item_text_box = slide.shapes.add_textbox(
            Inches(x + 0.1), Inches(y + 0.15), Inches(2), Inches(0.4)
        )
        item_text_frame = item_text_box.text_frame
        item_text_frame.text = item
        item_text_frame.word_wrap = True
        item_para = item_text_frame.paragraphs[0]
        item_para.font.size = Pt(14)
        item_para.font.color.rgb = RGBColor(255, 255, 255)
        item_para.alignment = PP_ALIGN.CENTER

def create_quote_slide(prs, quote, author, bg_color):
    """Create a quote slide"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Background
    add_gradient_background(slide, prs, bg_color)

    # Quote icon
    icon_box = slide.shapes.add_textbox(Inches(4), Inches(2), Inches(2), Inches(1))
    icon_frame = icon_box.text_frame
    icon_frame.text = "💡"
    icon_para = icon_frame.paragraphs[0]
    icon_para.font.size = Pt(80)
    icon_para.alignment = PP_ALIGN.CENTER

    # Quote box
    quote_box = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(1.5), Inches(3.5), Inches(7), Inches(2)
    )
    quote_box.fill.solid()
    quote_box.fill.fore_color.rgb = RGBColor(255, 255, 255)
    quote_box.line.fill.background()

    # Quote text
    quote_text = slide.shapes.add_textbox(Inches(2), Inches(3.8), Inches(6), Inches(1.2))
    quote_frame = quote_text.text_frame
    quote_frame.text = f'"{quote}"'
    quote_frame.word_wrap = True
    quote_para = quote_frame.paragraphs[0]
    quote_para.font.size = Pt(22)
    quote_para.font.italic = True
    quote_para.font.color.rgb = RGBColor(44, 62, 80)
    quote_para.alignment = PP_ALIGN.CENTER

    # Author
    author_text = slide.shapes.add_textbox(Inches(2), Inches(5.2), Inches(6), Inches(0.4))
    author_frame = author_text.text_frame
    author_frame.text = f"— {author}"
    author_para = author_frame.paragraphs[0]
    author_para.font.size = Pt(18)
    author_para.font.bold = True
    author_para.font.color.rgb = RGBColor(52, 73, 94)
    author_para.alignment = PP_ALIGN.RIGHT

def create_conclusion_slide(prs):
    """Create final conclusion slide"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Background
    add_gradient_background(slide, prs, RGBColor(46, 204, 113))

    # Success icon
    icon_box = slide.shapes.add_textbox(Inches(4.2), Inches(2), Inches(1.6), Inches(1))
    icon_frame = icon_box.text_frame
    icon_frame.text = "✅"
    icon_para = icon_frame.paragraphs[0]
    icon_para.font.size = Pt(80)
    icon_para.alignment = PP_ALIGN.CENTER

    # Title
    title_box = slide.shapes.add_textbox(Inches(1), Inches(3.5), Inches(8), Inches(0.8))
    title_frame = title_box.text_frame
    title_frame.text = "XULOSA"
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(48)
    title_para.font.bold = True
    title_para.font.color.rgb = RGBColor(255, 255, 255)
    title_para.alignment = PP_ALIGN.CENTER

    # Content
    content_box = slide.shapes.add_textbox(Inches(1.5), Inches(4.5), Inches(7), Inches(1.5))
    content_frame = content_box.text_frame
    content_frame.text = "Raqobat ustunligini yaratish va saqlash\nuzluksiz jarayon bo'lib, strategik yondashuv,\ninnovatsiya va mijozlarga yo'naltirilgan\nfaoliyatni talab qiladi."
    content_frame.word_wrap = True
    for para in content_frame.paragraphs:
        para.font.size = Pt(20)
        para.font.color.rgb = RGBColor(255, 255, 255)
        para.alignment = PP_ALIGN.CENTER
        para.space_after = Pt(6)

def main():
    """Create the complete presentation"""
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)

    # Slide 1: Title
    create_title_slide(prs)

    # Slide 2: Introduction
    create_content_slide(
        prs,
        "Raqobat ustunligi nima?",
        [
            "Kompaniyaning raqobatchilaridan ajralib turish qobiliyati",
            "Bozorda yetakchi o'rinni egallash imkoniyati",
            "Mijozlar uchun noyob qiymat yaratish",
            "Uzoq muddatli muvaffaqiyat ta'minlash"
        ],
        RGBColor(41, 128, 185),
        "🎯"
    )

    # Slide 3: Importance
    create_content_slide(
        prs,
        "Raqobat ustunligining ahamiyati",
        [
            "Bozor ulushini oshirish va o'sish imkoniyati",
            "Yuqori daromad va rentabellikka erishish",
            "Brendning kuchli pozitsiyasini yaratish",
            "Barqaror biznes modelini rivojlantirish"
        ],
        RGBColor(142, 68, 173),
        "📈"
    )

    # Slide 4: Michael Porter's theory
    create_quote_slide(
        prs,
        "Raqobat strategiyasi - bu boshqacha bo'lish demakdir",
        "Michael Porter",
        RGBColor(52, 73, 94)
    )

    # Slide 5: Three Generic Strategies
    create_strategy_slide(
        prs,
        "Porter'ning 3 ta asosiy strategiyasi",
        [
            ("1. Xarajatlarni kamaytirish (Cost Leadership)",
             "Eng past narxda mahsulot taklif etish orqali ustunlik"),
            ("2. Differensiatsiya (Differentiation)",
             "Noyob va o'ziga xos mahsulotlar yaratish"),
            ("3. Fokuslanish (Focus Strategy)",
             "Ma'lum bir segmentga yo'naltirilgan xizmat")
        ],
        RGBColor(231, 76, 60)
    )

    # Slide 6: Cost Leadership details
    create_two_column_slide(
        prs,
        "Xarajatlarni kamaytirish strategiyasi",
        "AFZALLIKLAR 👍",
        [
            "Katta bozor ulushi",
            "Narx bo'yicha raqobat",
            "Yuqori hajmda sotish",
            "Samaradorlik"
        ],
        "KAMCHILIKLAR 👎",
        [
            "Sifat xavfi",
            "Brand imijining sustligi",
            "Innovatsiya cheklanishi",
            "Nozik foyda marjasi"
        ],
        RGBColor(52, 152, 219)
    )

    # Slide 7: Differentiation details
    create_two_column_slide(
        prs,
        "Differensiatsiya strategiyasi",
        "AFZALLIKLAR 👍",
        [
            "Premium narx qo'yish",
            "Kuchli brend loyalligi",
            "Noyob pozitsiya",
            "Yuqori foyda marjasi"
        ],
        "KAMCHILIKLAR 👎",
        [
            "Yuqori xarajatlar",
            "Imitatsiya xavfi",
            "Cheklangan bozor",
            "Talabga bog'liqlik"
        ],
        RGBColor(155, 89, 182)
    )

    # Slide 8: Five Forces Model
    create_diagram_slide(
        prs,
        "Porter'ning 5 kuch modeli",
        "KOMPANIYA\nva\nRAQOBAT",
        [
            "Raqobatchilar",
            "Yangi kirishlar",
            "Yetkazuvchilar",
            "Xaridorlar"
        ],
        RGBColor(26, 188, 156)
    )

    # Slide 9: Sources of competitive advantage
    create_content_slide(
        prs,
        "Raqobat ustunligining manbalari",
        [
            "🔬 Innovatsiya va texnologiya",
            "👥 Malakali xodimlar va korporativ madaniyat",
            "💰 Moliyaviy resurslar va investitsiyalar",
            "🏭 Samarali ishlab chiqarish jarayonlari",
            "📊 Ma'lumotlar va analytics"
        ],
        RGBColor(230, 126, 34)
    )

    # Slide 10: Brand as advantage
    create_content_slide(
        prs,
        "Brend - kuchli raqobat vositasi",
        [
            "Mijozlar ishonchi va sodiqligini ta'minlash",
            "Premium narxlarni qo'llash imkoniyati",
            "Bozorda taniqlik va obro'",
            "Raqobatchilardan ajralib turish"
        ],
        RGBColor(52, 73, 94),
        "🏅"
    )

    # Slide 11: Innovation importance
    create_quote_slide(
        prs,
        "Innovatsiya raqobat ustunligining kalitidir",
        "Biznes ekspertlari",
        RGBColor(41, 128, 185)
    )

    # Slide 12: Innovation types
    create_two_column_slide(
        prs,
        "Innovatsiya turlari",
        "MAHSULOT INNOVATSIYASI 🚀",
        [
            "Yangi mahsulotlar yaratish",
            "Mavjudlarni yaxshilash",
            "Dizayn innovatsiyalari",
            "Texnologik yangiliklar"
        ],
        "JARAYON INNOVATSIYASI ⚙️",
        [
            "Ishlab chiqarishni optimallashtirish",
            "Xarajatlarni kamaytirish",
            "Sifatni oshirish",
            "Vaqtni tejash"
        ],
        RGBColor(231, 76, 60)
    )

    # Slide 13: Customer focus
    create_content_slide(
        prs,
        "Mijozlarga yo'naltirilgan strategiya",
        [
            "Mijozlar ehtiyojlarini chuqur tushunish",
            "Shaxsiy yondashuv va individual xizmat",
            "Mijozlar tajribasini yaxshilash",
            "Doimiy aloqa va feedback olish",
            "Mijozlar sodiqligini rivojlantirish"
        ],
        RGBColor(142, 68, 173),
        "❤️"
    )

    # Slide 14: Digital transformation
    create_content_slide(
        prs,
        "Raqamli transformatsiya",
        [
            "💻 Zamonaviy texnologiyalardan foydalanish",
            "☁️ Cloud xizmatlari va saqlash",
            "🤖 Sun'iy intellekt va avtomatlashtirish",
            "📱 Mobil ilovalar va platformalar",
            "🔐 Kiberxavfsizlik va ma'lumotlar himoyasi"
        ],
        RGBColor(41, 128, 185)
    )

    # Slide 15: Quality as advantage
    create_strategy_slide(
        prs,
        "Sifat - asosiy ustunlik",
        [
            ("Mahsulot sifati 🏆",
             "Yuqori standartlar va nazorat tizimi"),
            ("Xizmat sifati 🤝",
             "Professional va samimiy mijozlar xizmati"),
            ("Jarayon sifati ✅",
             "ISO standartlari va sertifikatsiyalar")
        ],
        RGBColor(46, 204, 113)
    )

    # Slide 16: Examples of successful companies
    create_content_slide(
        prs,
        "Muvaffaqiyatli kompaniyalar misollari",
        [
            "🍎 Apple - innovatsiya va premium brend",
            "📦 Amazon - mijozlar tajribasi va logistika",
            "☕ Starbucks - brend va xizmat sifati",
            "🚗 Toyota - sifat va samaradorlik",
            "✈️ Turkish Airlines - xizmat va geografiya"
        ],
        RGBColor(230, 126, 34)
    )

    # Slide 17: Sustainable competitive advantage
    create_content_slide(
        prs,
        "Barqaror raqobat ustunligi",
        [
            "Uzoq muddatli strategik rejalashtirish",
            "Doimiy o'rganish va rivojlanish",
            "Tashkiliy madaniyatni mustahkamlash",
            "Innovatsiyaga doimiy investitsiya",
            "Bozor o'zgarishlariga moslashuvchanlik"
        ],
        RGBColor(142, 68, 173),
        "♻️"
    )

    # Slide 18: Challenges and risks
    create_two_column_slide(
        prs,
        "Qiyinchiliklar va xavflar",
        "ICHKI QIYINCHILIKLAR 🏢",
        [
            "Resurslar cheklanganligi",
            "Tashkiliy qarshilik",
            "Malakali kadrlar yetishmasligi",
            "Moliyaviy cheklovlar"
        ],
        "TASHQI XAVFLAR 🌍",
        [
            "Qattiq raqobat",
            "Bozor o'zgarishlari",
            "Texnologik yangiliklar",
            "Iqtisodiy inqirozlar"
        ],
        RGBColor(231, 76, 60)
    )

    # Slide 19: Key success factors
    create_content_slide(
        prs,
        "Muvaffaqiyat uchun asosiy omillar",
        [
            "💡 Aniq strategiya va viziya",
            "👥 Kuchli jamoa va leadership",
            "🎯 Mijozlarga diqqat va e'tibor",
            "🔄 Tezkor moslashuv va innovatsiya",
            "📊 Ma'lumotlarga asoslangan qarorlar"
        ],
        RGBColor(52, 152, 219)
    )

    # Slide 20: Conclusion
    create_conclusion_slide(prs)

    # Save presentation
    filename = "Raqobat_Ustunligi_Nazariyasi.pptx"
    prs.save(filename)
    print(f"✅ Prezentatsiya muvaffaqiyatli yaratildi: {filename}")
    print(f"📊 Jami slaydlar: {len(prs.slides)}")
    print("\n🎨 Prezentatsiya xususiyatlari:")
    print("   • 20 ta professional slayd")
    print("   • Rangdor va chiroyli dizayn")
    print("   • Vizual elementlar va ikonkalar")
    print("   • O'zbek tilida to'liq kontent")
    print("   • Porter nazariyasi va zamonaviy yondashuvlar")
    print("   • Amaliy misollar va strategiyalar")

if __name__ == "__main__":
    main()
