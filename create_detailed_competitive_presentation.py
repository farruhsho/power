#!/usr/bin/env python3
"""
Detailed PowerPoint presentation about Competitive Advantage Theory in Uzbek
20 slides with detailed explanations for each section
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor

def add_gradient_background(slide, prs, color1):
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
    for i, (x, y, size) in enumerate([
        (8, 0.5, 2.5),
        (0.5, 5, 2),
        (7.5, 5.5, 1.5)
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

def create_explanation_slide(prs, title, main_text, explanation, bg_color, icon=""):
    """Create a slide with detailed explanation"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Background
    add_gradient_background(slide, prs, bg_color)

    # Title bar
    title_bar = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(0.5), Inches(0.5), Inches(9), Inches(0.9)
    )
    title_bar.fill.solid()
    title_bar.fill.fore_color.rgb = RGBColor(255, 255, 255)
    title_bar.line.fill.background()

    # Title text
    title_box = slide.shapes.add_textbox(Inches(1), Inches(0.65), Inches(8), Inches(0.6))
    title_frame = title_box.text_frame
    title_frame.text = title
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(32)
    title_para.font.bold = True
    title_para.font.color.rgb = bg_color
    title_para.alignment = PP_ALIGN.CENTER

    # Main content box
    content_box = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(0.8), Inches(1.8), Inches(8.4), Inches(2)
    )
    content_box.fill.solid()
    content_box.fill.fore_color.rgb = RGBColor(255, 255, 255)
    content_box.line.fill.background()

    # Main text
    main_text_box = slide.shapes.add_textbox(Inches(1.2), Inches(2), Inches(7.6), Inches(1.6))
    main_text_frame = main_text_box.text_frame
    main_text_frame.text = main_text
    main_text_frame.word_wrap = True
    for para in main_text_frame.paragraphs:
        para.font.size = Pt(18)
        para.font.color.rgb = RGBColor(44, 62, 80)
        para.line_spacing = 1.3

    # Explanation section header
    exp_header = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(0.8), Inches(4.1), Inches(8.4), Inches(0.5)
    )
    exp_header.fill.solid()
    exp_header.fill.fore_color.rgb = RGBColor(241, 196, 15)
    exp_header.line.fill.background()

    exp_header_text = slide.shapes.add_textbox(Inches(1), Inches(4.2), Inches(8), Inches(0.3))
    exp_header_frame = exp_header_text.text_frame
    exp_header_frame.text = "📚 TUSHUNTIRISH:"
    exp_header_para = exp_header_frame.paragraphs[0]
    exp_header_para.font.size = Pt(20)
    exp_header_para.font.bold = True
    exp_header_para.font.color.rgb = RGBColor(255, 255, 255)

    # Explanation box
    exp_box = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(0.8), Inches(4.7), Inches(8.4), Inches(1.8)
    )
    exp_box.fill.solid()
    exp_box.fill.fore_color.rgb = RGBColor(255, 255, 255)
    exp_box.line.color.rgb = RGBColor(241, 196, 15)
    exp_box.line.width = Pt(3)

    # Explanation text
    exp_text_box = slide.shapes.add_textbox(Inches(1.2), Inches(4.9), Inches(7.6), Inches(1.4))
    exp_text_frame = exp_text_box.text_frame
    exp_text_frame.text = explanation
    exp_text_frame.word_wrap = True
    for para in exp_text_frame.paragraphs:
        para.font.size = Pt(16)
        para.font.color.rgb = RGBColor(52, 73, 94)
        para.line_spacing = 1.2

    # Icon if provided
    if icon:
        icon_box = slide.shapes.add_textbox(Inches(8.5), Inches(1.2), Inches(1), Inches(0.8))
        icon_frame = icon_box.text_frame
        icon_frame.text = icon
        icon_para = icon_frame.paragraphs[0]
        icon_para.font.size = Pt(48)

def create_detailed_list_slide(prs, title, items_with_explanations, bg_color, icon=""):
    """Create a slide with list items and their explanations"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Background
    add_gradient_background(slide, prs, bg_color)

    # Title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.4), Inches(9), Inches(0.7))
    title_frame = title_box.text_frame
    title_frame.text = title
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(32)
    title_para.font.bold = True
    title_para.font.color.rgb = RGBColor(255, 255, 255)
    title_para.alignment = PP_ALIGN.CENTER

    # Items
    y_pos = 1.3
    item_height = 1.15
    colors = [
        RGBColor(52, 152, 219),
        RGBColor(46, 204, 113),
        RGBColor(155, 89, 182),
        RGBColor(231, 76, 60),
        RGBColor(241, 196, 15)
    ]

    for i, (item_title, item_explanation) in enumerate(items_with_explanations[:5]):
        # Item box
        item_box = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(0.7), Inches(y_pos), Inches(8.6), Inches(item_height)
        )
        item_box.fill.solid()
        item_box.fill.fore_color.rgb = RGBColor(255, 255, 255)
        item_box.line.color.rgb = colors[i]
        item_box.line.width = Pt(4)

        # Item title
        title_text_box = slide.shapes.add_textbox(
            Inches(1), Inches(y_pos + 0.1), Inches(8), Inches(0.35)
        )
        title_text_frame = title_text_box.text_frame
        title_text_frame.text = item_title
        title_para = title_text_frame.paragraphs[0]
        title_para.font.size = Pt(18)
        title_para.font.bold = True
        title_para.font.color.rgb = colors[i]

        # Item explanation
        exp_text_box = slide.shapes.add_textbox(
            Inches(1), Inches(y_pos + 0.5), Inches(7.8), Inches(0.6)
        )
        exp_text_frame = exp_text_box.text_frame
        exp_text_frame.text = item_explanation
        exp_text_frame.word_wrap = True
        for para in exp_text_frame.paragraphs:
            para.font.size = Pt(14)
            para.font.color.rgb = RGBColor(52, 73, 94)

        y_pos += item_height + 0.1

    # Icon
    if icon:
        icon_box = slide.shapes.add_textbox(Inches(8.8), Inches(0.5), Inches(1), Inches(0.6))
        icon_frame = icon_box.text_frame
        icon_frame.text = icon
        icon_para = icon_frame.paragraphs[0]
        icon_para.font.size = Pt(36)

def create_strategy_detail_slide(prs, title, strategy_name, description, advantages, disadvantages, example, bg_color):
    """Create a detailed strategy explanation slide"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Background
    add_gradient_background(slide, prs, bg_color)

    # Title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.4), Inches(9), Inches(0.6))
    title_frame = title_box.text_frame
    title_frame.text = title
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(32)
    title_para.font.bold = True
    title_para.font.color.rgb = RGBColor(255, 255, 255)
    title_para.alignment = PP_ALIGN.CENTER

    # Strategy name box
    name_box = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(1), Inches(1.2), Inches(8), Inches(0.6)
    )
    name_box.fill.solid()
    name_box.fill.fore_color.rgb = RGBColor(255, 255, 255)
    name_box.line.fill.background()

    name_text = slide.shapes.add_textbox(Inches(1.2), Inches(1.35), Inches(7.6), Inches(0.35))
    name_frame = name_text.text_frame
    name_frame.text = strategy_name
    name_para = name_frame.paragraphs[0]
    name_para.font.size = Pt(22)
    name_para.font.bold = True
    name_para.font.color.rgb = bg_color
    name_para.alignment = PP_ALIGN.CENTER

    # Description
    desc_box = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(1), Inches(2), Inches(8), Inches(0.9)
    )
    desc_box.fill.solid()
    desc_box.fill.fore_color.rgb = RGBColor(255, 255, 255)
    desc_box.line.color.rgb = RGBColor(241, 196, 15)
    desc_box.line.width = Pt(3)

    desc_text = slide.shapes.add_textbox(Inches(1.3), Inches(2.15), Inches(7.4), Inches(0.6))
    desc_frame = desc_text.text_frame
    desc_frame.text = description
    desc_frame.word_wrap = True
    for para in desc_frame.paragraphs:
        para.font.size = Pt(15)
        para.font.color.rgb = RGBColor(44, 62, 80)

    # Advantages column
    adv_box = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(0.7), Inches(3.2), Inches(4.3), Inches(2.5)
    )
    adv_box.fill.solid()
    adv_box.fill.fore_color.rgb = RGBColor(46, 204, 113)
    adv_box.line.fill.background()

    adv_title = slide.shapes.add_textbox(Inches(1), Inches(3.4), Inches(3.7), Inches(0.3))
    adv_title_frame = adv_title.text_frame
    adv_title_frame.text = "✅ AFZALLIKLAR:"
    adv_title_para = adv_title_frame.paragraphs[0]
    adv_title_para.font.size = Pt(18)
    adv_title_para.font.bold = True
    adv_title_para.font.color.rgb = RGBColor(255, 255, 255)

    adv_text = slide.shapes.add_textbox(Inches(1), Inches(3.8), Inches(3.7), Inches(1.7))
    adv_text_frame = adv_text.text_frame
    adv_text_frame.text = "\n".join([f"• {adv}" for adv in advantages])
    adv_text_frame.word_wrap = True
    for para in adv_text_frame.paragraphs:
        para.font.size = Pt(14)
        para.font.color.rgb = RGBColor(255, 255, 255)
        para.space_after = Pt(4)

    # Disadvantages column
    dis_box = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(5.3), Inches(3.2), Inches(4.3), Inches(2.5)
    )
    dis_box.fill.solid()
    dis_box.fill.fore_color.rgb = RGBColor(231, 76, 60)
    dis_box.line.fill.background()

    dis_title = slide.shapes.add_textbox(Inches(5.6), Inches(3.4), Inches(3.7), Inches(0.3))
    dis_title_frame = dis_title.text_frame
    dis_title_frame.text = "⚠️ KAMCHILIKLAR:"
    dis_title_para = dis_title_frame.paragraphs[0]
    dis_title_para.font.size = Pt(18)
    dis_title_para.font.bold = True
    dis_title_para.font.color.rgb = RGBColor(255, 255, 255)

    dis_text = slide.shapes.add_textbox(Inches(5.6), Inches(3.8), Inches(3.7), Inches(1.7))
    dis_text_frame = dis_text.text_frame
    dis_text_frame.text = "\n".join([f"• {dis}" for dis in disadvantages])
    dis_text_frame.word_wrap = True
    for para in dis_text_frame.paragraphs:
        para.font.size = Pt(14)
        para.font.color.rgb = RGBColor(255, 255, 255)
        para.space_after = Pt(4)

    # Example box
    ex_box = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(0.7), Inches(6), Inches(8.6), Inches(0.7)
    )
    ex_box.fill.solid()
    ex_box.fill.fore_color.rgb = RGBColor(255, 255, 255)
    ex_box.line.color.rgb = RGBColor(52, 152, 219)
    ex_box.line.width = Pt(3)

    ex_text = slide.shapes.add_textbox(Inches(1), Inches(6.15), Inches(8.2), Inches(0.4))
    ex_text_frame = ex_text.text_frame
    ex_text_frame.text = f"💼 MISOL: {example}"
    ex_text_frame.word_wrap = True
    for para in ex_text_frame.paragraphs:
        para.font.size = Pt(15)
        para.font.color.rgb = RGBColor(52, 73, 94)

def create_porter_forces_slide(prs):
    """Create Porter's 5 Forces detailed slide"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Background
    add_gradient_background(slide, prs, RGBColor(26, 188, 156))

    # Title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.4), Inches(9), Inches(0.6))
    title_frame = title_box.text_frame
    title_frame.text = "PORTER'NING 5 KUCH MODELI"
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(36)
    title_para.font.bold = True
    title_para.font.color.rgb = RGBColor(255, 255, 255)
    title_para.alignment = PP_ALIGN.CENTER

    # Center - Company
    center = slide.shapes.add_shape(
        MSO_SHAPE.OVAL,
        Inches(3.5), Inches(3), Inches(3), Inches(2)
    )
    center.fill.solid()
    center.fill.fore_color.rgb = RGBColor(241, 196, 15)
    center.line.color.rgb = RGBColor(243, 156, 18)
    center.line.width = Pt(4)

    center_text = slide.shapes.add_textbox(Inches(3.8), Inches(3.7), Inches(2.4), Inches(0.6))
    center_frame = center_text.text_frame
    center_frame.text = "KOMPANIYA\nva\nRAQOBAT"
    for para in center_frame.paragraphs:
        para.font.size = Pt(18)
        para.font.bold = True
        para.font.color.rgb = RGBColor(255, 255, 255)
        para.alignment = PP_ALIGN.CENTER

    # Forces with detailed explanations
    forces = [
        ("Mavjud raqobatchilar",
         "Sohada faoliyat\nyuritayotgan kompaniyalar",
         Inches(3.5), Inches(1.2), RGBColor(231, 76, 60)),
        ("Yangi kirishlar xavfi",
         "Yangi kompaniyalarning\nbozorga kirish ehtimoli",
         Inches(0.5), Inches(3), RGBColor(52, 152, 219)),
        ("Yetkazuvchilar kuchi",
         "Ta'minotchilarning\nnarx belgilash qobiliyati",
         Inches(3.5), Inches(5.5), RGBColor(155, 89, 182)),
        ("Xaridorlar kuchi",
         "Mijozlarning\nsavdolashish qobiliyati",
         Inches(6.8), Inches(3), RGBColor(46, 204, 113)),
    ]

    for force_name, force_desc, x, y, color in forces:
        # Force box
        force_box = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            x, y, Inches(2.7), Inches(1.2)
        )
        force_box.fill.solid()
        force_box.fill.fore_color.rgb = color
        force_box.line.fill.background()

        # Force name
        name_box = slide.shapes.add_textbox(x + Inches(0.1), y + Inches(0.1), Inches(2.5), Inches(0.4))
        name_frame = name_box.text_frame
        name_frame.text = force_name
        name_frame.word_wrap = True
        for para in name_frame.paragraphs:
            para.font.size = Pt(16)
            para.font.bold = True
            para.font.color.rgb = RGBColor(255, 255, 255)
            para.alignment = PP_ALIGN.CENTER

        # Force description
        desc_box = slide.shapes.add_textbox(x + Inches(0.1), y + Inches(0.55), Inches(2.5), Inches(0.6))
        desc_frame = desc_box.text_frame
        desc_frame.text = force_desc
        desc_frame.word_wrap = True
        for para in desc_frame.paragraphs:
            para.font.size = Pt(12)
            para.font.color.rgb = RGBColor(255, 255, 255)
            para.alignment = PP_ALIGN.CENTER

def create_conclusion_slide(prs):
    """Create final conclusion slide"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Background
    add_gradient_background(slide, prs, RGBColor(46, 204, 113))

    # Icon
    icon_box = slide.shapes.add_textbox(Inches(4.2), Inches(1.5), Inches(1.6), Inches(1))
    icon_frame = icon_box.text_frame
    icon_frame.text = "🎓"
    icon_para = icon_frame.paragraphs[0]
    icon_para.font.size = Pt(80)
    icon_para.alignment = PP_ALIGN.CENTER

    # Title
    title_box = slide.shapes.add_textbox(Inches(1), Inches(2.8), Inches(8), Inches(0.6))
    title_frame = title_box.text_frame
    title_frame.text = "XULOSA"
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(48)
    title_para.font.bold = True
    title_para.font.color.rgb = RGBColor(255, 255, 255)
    title_para.alignment = PP_ALIGN.CENTER

    # Content box
    content_box = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(1), Inches(3.7), Inches(8), Inches(2.5)
    )
    content_box.fill.solid()
    content_box.fill.fore_color.rgb = RGBColor(255, 255, 255)
    content_box.line.color.rgb = RGBColor(39, 174, 96)
    content_box.line.width = Pt(5)

    # Key points
    key_points = """✅ Raqobat ustunligi - biznesning asosi

✅ Strategik yondashuv zarur

✅ Innovatsiya va moslashuvchanlik muhim

✅ Mijozlarga yo'naltirilganlik - muvaffaqiyat kaliti

✅ Doimiy takomillashtirish va rivojlanish"""

    content_text = slide.shapes.add_textbox(Inches(1.5), Inches(4), Inches(7), Inches(2))
    content_frame = content_text.text_frame
    content_frame.text = key_points
    content_frame.word_wrap = True
    for para in content_frame.paragraphs:
        para.font.size = Pt(18)
        para.font.color.rgb = RGBColor(44, 62, 80)
        para.space_after = Pt(8)

def main():
    """Create the complete detailed presentation"""
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)

    # Slide 1: Title
    create_title_slide(prs)

    # Slide 2: Introduction with explanation
    create_explanation_slide(
        prs,
        "Raqobat ustunligi nima?",
        "Raqobat ustunligi - bu kompaniyaning o'z raqobatchilaridan ajralib turadigan va unga bozorda yetakchi o'rinni egallash imkonini beradigan xususiyatlar, strategiyalar va resurslar majmuidir.",
        "Bu nafaqat mahsulot yoki xizmatning sifati, balki kompaniyaning umumiy strategiyasi, brendi, innovatsiyalari, xodimlarining malakasi va boshqa ko'plab omillarni o'z ichiga oladi. Raqobat ustunligi uzoq muddatli muvaffaqiyat va barqaror o'sishning kalitidir.",
        RGBColor(41, 128, 185),
        "🎯"
    )

    # Slide 3: Why it's important
    create_explanation_slide(
        prs,
        "Nima uchun bu muhim?",
        "Zamonaviy bozorda raqobat tobora kuchayib bormoqda. Har bir sohada yuzlab kompaniyalar bir xil mijozlar uchun kurashmoqda. Raqobat ustunligi - bu kompaniyaning omon qolishi va rivojlanishi uchun zarur shart.",
        "Raqobat ustunligiga ega kompaniyalar yuqori foyda oladi, bozor ulushini oshiradi, kuchli brend yaratadi va barqaror biznes modelini quradi. Bu nafaqat bugungi kun, balki kelajak uchun ham investitsiya hisoblanadi.",
        RGBColor(142, 68, 173),
        "📈"
    )

    # Slide 4: Michael Porter introduction
    create_explanation_slide(
        prs,
        "Michael Porter - raqobat strategiyasi ustasi",
        "Michael Porter - Harvard biznes maktabining professori, raqobat strategiyasi sohasining asoschisi. U 1980-yilda o'zining mashhur nazariyalarini ishlab chiqdi va biznes dunyosini butunlay o'zgartirdi.",
        "Porter'ning asosiy g'oyasi - raqobat strategiyasi bu boshqacha bo'lish demakdir. Kompaniya hammaga o'xshab ketmasligi, o'zining noyob yo'lini topishi va qiymat yaratishda alohida yondashuvga ega bo'lishi kerak.",
        RGBColor(52, 73, 94),
        "👨‍🏫"
    )

    # Slide 5: Three Generic Strategies overview
    create_explanation_slide(
        prs,
        "Porter'ning 3 asosiy strategiyasi",
        "Michael Porter uchta asosiy raqobat strategiyasini aniqlagan: 1) Xarajatlarni kamaytirish (Cost Leadership), 2) Differensiatsiya (Differentiation), 3) Fokuslanish (Focus). Har bir strategiya o'ziga xos yondashuvni talab qiladi.",
        "Bu strategiyalar har qanday biznes uchun asos bo'lib xizmat qiladi. Kompaniya ulardan birini tanlashi va barcha resurslarini shu yo'nalishda to'plashi kerak. Bir vaqtning o'zida bir nechta strategiyani amalga oshirishga urinish muvaffaqiyatsizlikka olib keladi.",
        RGBColor(231, 76, 60),
        "🎯"
    )

    # Slide 6: Cost Leadership Strategy
    create_strategy_detail_slide(
        prs,
        "1-STRATEGIYA: XARAJATLARNI KAMAYTIRISH",
        "💰 COST LEADERSHIP (Xarajat liderlik)",
        "Bu strategiya kompaniyaga bozorda eng past narxlarni taklif qilish orqali ustunlik qozonishga qaratilgan. Maqsad - ishlab chiqarish va operatsion xarajatlarni maksimal darajada kamaytirish.",
        [
            "Katta bozor ulushini egallash",
            "Ommaviy mijozlarga xizmat",
            "Narx bo'yicha kuchli raqobat",
            "Operatsion samaradorlik"
        ],
        [
            "Sifat pasayishi xavfi",
            "Kichik foyda marjasi",
            "Brand imijining zaif bo'lishi",
            "Innovatsiyalarga kam investitsiya"
        ],
        "Walmart, Costco - eng past narxlarda mahsulotlar taklif qiluvchi kompaniyalar",
        RGBColor(52, 152, 219)
    )

    # Slide 7: Differentiation Strategy
    create_strategy_detail_slide(
        prs,
        "2-STRATEGIYA: DIFFERENSIATSIYA",
        "⭐ DIFFERENTIATION (Farqlash)",
        "Differensiatsiya strategiyasi - bu kompaniyaning bozorda noyob va o'ziga xos mahsulot yoki xizmat yaratishi. Maqsad - raqobatchilardan ajralib turish va premium narx qo'yish imkoniyatini olish.",
        [
            "Premium narxlarni qo'llash",
            "Kuchli brand loyalligi",
            "Noyob bozor pozitsiyasi",
            "Yuqori foyda marjasi"
        ],
        [
            "Yuqori ishlab chiqarish xarajatlari",
            "Imitatsiya xavfi",
            "Cheklangan maqsadli auditoriya",
            "Bozor talabiga bog'liqlik"
        ],
        "Apple, Tesla, Mercedes-Benz - noyob mahsulotlar yaratuvchi kompaniyalar",
        RGBColor(155, 89, 182)
    )

    # Slide 8: Focus Strategy
    create_strategy_detail_slide(
        prs,
        "3-STRATEGIYA: FOKUSLANISH",
        "🎯 FOCUS STRATEGY (Fokus strategiya)",
        "Fokus strategiyasi - bu kompaniyaning barcha resurslarini ma'lum bir segmentga, nichga yoki geografik hududga yo'naltirishi. Bu strategiya kichik, lekin samarali biznes yuritishga imkon beradi.",
        [
            "Chuqur bozor bilimi",
            "Mijozlarning aniq ehtiyojlarini qondirib berish",
            "Yuqori mijoz sodiqligis",
            "Niche bozorda monopoliya"
        ],
        [
            "Cheklangan o'sish imkoniyatlari",
            "Bozor o'zgarishlariga sezgirlik",
            "Kichik miqyos",
            "Diversifikatsiya qiyinligi"
        ],
        "Rolls-Royce, Rolex - premium segmentga yo'naltirilgan kompaniyalar",
        RGBColor(230, 126, 34)
    )

    # Slide 9: Porter's 5 Forces
    create_porter_forces_slide(prs)

    # Slide 10: Five Forces explanation
    create_explanation_slide(
        prs,
        "5 kuch modeli tushuntirishi",
        "Porter'ning 5 kuch modeli kompaniyaga o'z sohalarida raqobat muhitini to'liq tahlil qilish imkonini beradi. Bu tahlil strategik qarorlar qabul qilishda muhim ahamiyatga ega.",
        "Har bir kuch kompaniyaning rentabelligiga ta'sir ko'rsatadi. Agar barcha 5 kuch kuchli bo'lsa, soha juda raqobatbardosh va foyda olish qiyin. Agar kuchlar zaif bo'lsa, kompaniya yuqori foyda olish imkoniyatiga ega. Maqsad - bu kuchlarni tushunish va ularni o'z foydasiga ishlatish.",
        RGBColor(26, 188, 156),
        "🔍"
    )

    # Slide 11: Sources of competitive advantage
    create_detailed_list_slide(
        prs,
        "Raqobat ustunligining manbalari",
        [
            ("🔬 Innovatsiya va texnologiya",
             "Yangi mahsulotlar, xizmatlar va jarayonlarni yaratish orqali raqobatchilardan oldinda borish"),
            ("👥 Inson resurslari",
             "Malakali, motivatsiyalangan va professional xodimlar - kompaniyaning asosiy boyligi"),
            ("💰 Moliyaviy quvvat",
             "Kuchli moliyaviy asosga ega bo'lish innovatsiyalarga investitsiya qilish imkonini beradi"),
            ("🏭 Operatsion samaradorlik",
             "Samarali jarayonlar, past xarajatlar va yuqori sifat kombinatsiyasi"),
            ("📊 Ma'lumotlar va analytics",
             "Big data va analytics yordamida to'g'ri qarorlar qabul qilish qobiliyati")
        ],
        RGBColor(230, 126, 34),
        "💎"
    )

    # Slide 12: Brand power
    create_explanation_slide(
        prs,
        "Brend - kuchli raqobat vositasi",
        "Brend - bu kompaniya haqidagi umumiy taassurot, mijozlar bilan bog'lanish va ishonch darajasi. Kuchli brend kompaniyaga ko'plab imkoniyatlar beradi va uzoq muddatli ustunlik yaratadi.",
        "Brendni qurish - bu uzoq jarayon. U sizdan izchillik, sifat, mijozlar bilan samimiy munosabat va o'z va'dalaringizni bajarishni talab qiladi. Kuchli brend kompaniyaga premium narxlar qo'yish, yangi bozorlarga oson kirish va raqobat sharoitida barqarorlikni ta'minlaydi.",
        RGBColor(52, 73, 94),
        "🏅"
    )

    # Slide 13: Innovation importance
    create_explanation_slide(
        prs,
        "Innovatsiya - kelajak kaliti",
        "Zamonaviy dunyoda innovatsiyasiz kompaniya o'sishi va hatto omon qolishi mumkin emas. Bozor tez o'zgarmoqda, texnologiyalar rivojlanmoqda va mijozlar kutishlari ortmoqda.",
        "Innovatsiya nafaqat yangi mahsulotlar yaratish, balki mavjud jarayonlarni yaxshilash, yangi biznes modellarga o'tish va mijozlar tajribasini o'zgartirish ham hisoblanadi. Innovator kompaniyalar doimo raqobatchilardan bir qadam oldinda boradilar.",
        RGBColor(41, 128, 185),
        "💡"
    )

    # Slide 14: Types of innovation
    create_detailed_list_slide(
        prs,
        "Innovatsiya turlari va ularning ahamiyati",
        [
            ("🚀 Mahsulot innovatsiyasi",
             "Yangi va yaxshilangan mahsulotlar yaratish, dizayn va funksionallikni yangilash"),
            ("⚙️ Jarayon innovatsiyasi",
             "Ishlab chiqarish va operatsion jarayonlarni optimallashtirish, avtomatlashtirish"),
            ("📱 Texnologik innovatsiya",
             "Yangi texnologiyalarni qo'llash va raqamli transformatsiya"),
            ("🎨 Marketing innovatsiyasi",
             "Yangi marketing strategiyalari, kanallar va mijozlar bilan muloqot usullari"),
            ("🔄 Biznes model innovatsiyasi",
             "Butunlay yangi biznes modellarini yaratish va joriy etish")
        ],
        RGBColor(142, 68, 173),
        "🎯"
    )

    # Slide 15: Customer-centric approach
    create_explanation_slide(
        prs,
        "Mijozlarga yo'naltirilgan strategiya",
        "Mijozlar - har qanday biznesning markazi. Mijozlarga yo'naltirilgan kompaniyalar o'z mijozlarining ehtiyojlari, istakları va muammolarini chuqur tushunadilar va ular uchun maksimal qiymat yaratishga harakat qiladilar.",
        "Bu yondashuv shunchaki yaxshi xizmat ko'rsatish emas. Bu butun kompaniya madaniyati, jarayonlar va qarorlarni mijozlar nuqtai nazaridan ko'rish demakdir. Mijozlarga yo'naltirilgan kompaniyalar yuqori sodiqlik, takroriy sotuvlar va ijobiy tavsiyalar oladi.",
        RGBColor(231, 76, 60),
        "❤️"
    )

    # Slide 16: Digital transformation
    create_detailed_list_slide(
        prs,
        "Raqamli transformatsiya - zaruryat",
        [
            ("💻 Cloud texnologiyalar",
             "Moslashuvchanlik, samaradorlik va xarajatlarni kamaytirish imkoniyati"),
            ("🤖 Sun'iy intellekt va ML",
             "Avtomatlashtirish, bashorat qilish va qarorlarni optimallashtirish"),
            ("📱 Mobil-first yondashuv",
             "Mijozlar uchun qulay mobil tajriba yaratish"),
            ("🔐 Kiberxavfsizlik",
             "Ma'lumotlar himoyasi va mijozlar ishonchini ta'minlash"),
            ("📊 Data Analytics",
             "Ma'lumotlarga asoslangan qarorlar qabul qilish va biznesni optimallashtirish")
        ],
        RGBColor(41, 128, 185),
        "🌐"
    )

    # Slide 17: Quality as foundation
    create_explanation_slide(
        prs,
        "Sifat - barcha narsaning asosi",
        "Sifat - bu kompaniyaning reputatsiyasi, brendi va uzoq muddatli muvaffaqiyatining asosi. Yuqori sifat mijozlar sodiqligini ta'minlaydi, qayta sotuvlarni oshiradi va ijobiy tavsiyalar yaratadi.",
        "Sifatni ta'minlash - bu tizimli yondashuv. Bu standartlar, jarayonlar, nazorat tizimlari va doimiy yaxshilanishni o'z ichiga oladi. ISO sertifikatlari, sifat menejment tizimi va xodimlarni o'qitish - sifatli mahsulot va xizmatlarning asosiy komponentlari.",
        RGBColor(46, 204, 113),
        "🏆"
    )

    # Slide 18: Success examples
    create_detailed_list_slide(
        prs,
        "Muvaffaqiyatli kompaniyalar va ularning sirları",
        [
            ("🍎 Apple Inc.",
             "Innovatsiya, dizayn va premium brend orqali raqobat ustunligi"),
            ("📦 Amazon",
             "Mijozlar tajribasi, logistika va texnologiya orqali bozor lideri"),
            ("☕ Starbucks",
             "Noyob tajriba yaratish va kuchli brend orqali premium pozitsiya"),
            ("🚗 Toyota",
             "Sifat, ishonchlilik va doimiy takomillashtirish falsafasi"),
            ("✈️ Turkish Airlines",
             "Xizmat sifati, geografik qamrov va milliy brand orqali o'sish")
        ],
        RGBColor(230, 126, 34),
        "⭐"
    )

    # Slide 19: Sustainable advantage
    create_explanation_slide(
        prs,
        "Barqaror raqobat ustunligi",
        "Raqobat ustunligini yaratish - bu faqat birinchi qadam. Muhimi - uni uzoq muddat saqlab qolish. Barqaror ustunlik uchun kompaniya doimiy ravishda o'zgarib, innovatsiyalar kiritib va bozor o'zgarishlariga tez moslashishi kerak.",
        "Barqaror ustunlik uchun: 1) Uzoq muddatli strategik reja, 2) Doimiy innovatsiya, 3) Kuchli tashkiliy madaniyat, 4) Moslashuvchanlik va chaqqonlik, 5) Investitsiya va resurslar boshqaruvi muhim ahamiyatga ega.",
        RGBColor(142, 68, 173),
        "♻️"
    )

    # Slide 20: Conclusion
    create_conclusion_slide(prs)

    # Save presentation
    filename = "Raqobat_Ustunligi_Batafsil.pptx"
    prs.save(filename)
    print(f"✅ Prezentatsiya muvaffaqiyatli yaratildi: {filename}")
    print(f"📊 Jami slaydlar: {len(prs.slides)}")
    print("\n🎨 Prezentatsiya xususiyatlari:")
    print("   • 20 ta professional slayd")
    print("   • Har bir mavzu batafsil tushuntirilgan")
    print("   • Tushuntirish bo'limlari bilan")
    print("   • Amaliy misollar va strategiyalar")
    print("   • Vizual elementlar va ranglar")
    print("   • O'zbek tilida to'liq kontent")
    print("   • Porter nazariyasi va zamonaviy yondashuvlar")

if __name__ == "__main__":
    main()
