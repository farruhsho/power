#!/usr/bin/env python3
"""
Interactive PowerPoint presentation about Microsoft Word in Uzbek
With animations, beautiful design, and quiz questions
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR

def create_title_slide(prs):
    """Create an attractive title slide"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout

    # Add gradient background (simulated with shapes)
    background = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        0, 0, prs.slide_width, prs.slide_height
    )
    background.fill.solid()
    background.fill.fore_color.rgb = RGBColor(41, 128, 185)  # Blue
    background.line.fill.background()

    # Title
    title_box = slide.shapes.add_textbox(
        Inches(1), Inches(2), Inches(8), Inches(1.5)
    )
    title_frame = title_box.text_frame
    title_frame.text = "Microsoft Word bilan tanishing!"
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(54)
    title_para.font.bold = True
    title_para.font.color.rgb = RGBColor(255, 255, 255)
    title_para.alignment = PP_ALIGN.CENTER

    # Subtitle
    subtitle_box = slide.shapes.add_textbox(
        Inches(1), Inches(4), Inches(8), Inches(1)
    )
    subtitle_frame = subtitle_box.text_frame
    subtitle_frame.text = "Interaktiv o'quv prezentatsiyasi"
    subtitle_para = subtitle_frame.paragraphs[0]
    subtitle_para.font.size = Pt(32)
    subtitle_para.font.color.rgb = RGBColor(236, 240, 241)
    subtitle_para.alignment = PP_ALIGN.CENTER

    return slide

def create_intro_slide(prs):
    """Introduction slide about Microsoft Word"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Background
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid()
    bg.fill.fore_color.rgb = RGBColor(236, 240, 241)
    bg.line.fill.background()

    # Title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(1))
    title_frame = title_box.text_frame
    title_frame.text = "Microsoft Word nima?"
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(44)
    title_para.font.bold = True
    title_para.font.color.rgb = RGBColor(41, 128, 185)

    # Content box with icon
    icon_box = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(1), Inches(2), Inches(8), Inches(4)
    )
    icon_box.fill.solid()
    icon_box.fill.fore_color.rgb = RGBColor(255, 255, 255)
    icon_box.line.color.rgb = RGBColor(41, 128, 185)
    icon_box.line.width = Pt(3)

    # Add text content
    text_box = slide.shapes.add_textbox(Inches(1.5), Inches(2.5), Inches(7), Inches(3))
    text_frame = text_box.text_frame
    text_frame.word_wrap = True

    content = """📝 Microsoft Word - bu matn muharriri dasturi

✨ Nimalar qilish mumkin:
• Hujjatlar yaratish va tahrirlash
• Matnni formatlash va bezash
• Rasmlar va jadvallar qo'shish
• Professional hujjatlar tayyorlash
• PDF formatda saqlash"""

    p = text_frame.paragraphs[0]
    p.text = content
    p.font.size = Pt(20)
    p.font.color.rgb = RGBColor(44, 62, 80)
    p.line_spacing = 1.5

    return slide

def create_features_slide(prs, title, features, color):
    """Create a slide with features list"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Background
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid()
    bg.fill.fore_color.rgb = RGBColor(236, 240, 241)
    bg.line.fill.background()

    # Title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(1))
    title_frame = title_box.text_frame
    title_frame.text = title
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(40)
    title_para.font.bold = True
    title_para.font.color.rgb = color

    # Features with animated boxes
    y_position = 2
    for i, feature in enumerate(features):
        # Feature box
        feature_box = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(1.5), Inches(y_position), Inches(7), Inches(0.8)
        )
        feature_box.fill.solid()
        feature_box.fill.fore_color.rgb = RGBColor(52, 152, 219)
        feature_box.line.fill.background()

        # Feature text
        text_box = slide.shapes.add_textbox(
            Inches(2), Inches(y_position + 0.1), Inches(6.5), Inches(0.6)
        )
        text_frame = text_box.text_frame
        text_frame.text = feature
        p = text_frame.paragraphs[0]
        p.font.size = Pt(18)
        p.font.color.rgb = RGBColor(255, 255, 255)
        p.font.bold = True

        y_position += 1

    return slide

def create_quiz_slide(prs, question_num, question, options, correct_answer):
    """Create an interactive quiz slide"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Background with gradient effect
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid()
    bg.fill.fore_color.rgb = RGBColor(142, 68, 173)  # Purple
    bg.line.fill.background()

    # Quiz header
    header_box = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(1), Inches(0.5), Inches(8), Inches(1)
    )
    header_box.fill.solid()
    header_box.fill.fore_color.rgb = RGBColor(155, 89, 182)
    header_box.line.fill.background()

    header_text = slide.shapes.add_textbox(Inches(1.5), Inches(0.7), Inches(7), Inches(0.6))
    header_frame = header_text.text_frame
    header_frame.text = f"🎮 Test Savoli #{question_num}"
    header_para = header_frame.paragraphs[0]
    header_para.font.size = Pt(32)
    header_para.font.bold = True
    header_para.font.color.rgb = RGBColor(255, 255, 255)
    header_para.alignment = PP_ALIGN.CENTER

    # Question box
    q_box = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(1), Inches(2), Inches(8), Inches(1.2)
    )
    q_box.fill.solid()
    q_box.fill.fore_color.rgb = RGBColor(255, 255, 255)
    q_box.line.color.rgb = RGBColor(155, 89, 182)
    q_box.line.width = Pt(4)

    q_text = slide.shapes.add_textbox(Inches(1.5), Inches(2.2), Inches(7), Inches(0.8))
    q_frame = q_text.text_frame
    q_frame.text = question
    q_frame.word_wrap = True
    q_para = q_frame.paragraphs[0]
    q_para.font.size = Pt(24)
    q_para.font.bold = True
    q_para.font.color.rgb = RGBColor(52, 73, 94)
    q_para.alignment = PP_ALIGN.CENTER

    # Answer options
    option_colors = [
        RGBColor(231, 76, 60),   # Red
        RGBColor(52, 152, 219),  # Blue
        RGBColor(46, 204, 113),  # Green
        RGBColor(241, 196, 15)   # Yellow
    ]

    y_pos = 3.5
    for i, (letter, option) in enumerate(options):
        # Option button
        opt_box = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(1.5), Inches(y_pos), Inches(7), Inches(0.7)
        )
        opt_box.fill.solid()
        opt_box.fill.fore_color.rgb = option_colors[i]
        opt_box.line.fill.background()

        # Option text
        opt_text = slide.shapes.add_textbox(
            Inches(2), Inches(y_pos + 0.1), Inches(6), Inches(0.5)
        )
        opt_frame = opt_text.text_frame
        opt_frame.text = f"{letter}) {option}"
        opt_para = opt_frame.paragraphs[0]
        opt_para.font.size = Pt(20)
        opt_para.font.bold = True
        opt_para.font.color.rgb = RGBColor(255, 255, 255)

        y_pos += 0.9

    # Add correct answer indicator (small note)
    answer_note = slide.shapes.add_textbox(Inches(8), Inches(6.5), Inches(1.5), Inches(0.5))
    answer_frame = answer_note.text_frame
    answer_frame.text = f"✓ {correct_answer}"
    answer_para = answer_frame.paragraphs[0]
    answer_para.font.size = Pt(14)
    answer_para.font.color.rgb = RGBColor(255, 255, 255)

    return slide

def create_character_slide(prs, character_name, message, tip):
    """Create a slide with a character giving tips"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Background
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid()
    bg.fill.fore_color.rgb = RGBColor(26, 188, 156)  # Turquoise
    bg.line.fill.background()

    # Character avatar (circle)
    avatar = slide.shapes.add_shape(
        MSO_SHAPE.OVAL,
        Inches(1), Inches(2), Inches(2), Inches(2)
    )
    avatar.fill.solid()
    avatar.fill.fore_color.rgb = RGBColor(241, 196, 15)
    avatar.line.color.rgb = RGBColor(243, 156, 18)
    avatar.line.width = Pt(5)

    # Character emoji/text
    char_text = slide.shapes.add_textbox(Inches(1.3), Inches(2.5), Inches(1.4), Inches(1))
    char_frame = char_text.text_frame
    char_frame.text = "👨‍🏫"
    char_para = char_frame.paragraphs[0]
    char_para.font.size = Pt(72)
    char_para.alignment = PP_ALIGN.CENTER

    # Character name
    name_box = slide.shapes.add_textbox(Inches(0.7), Inches(4.2), Inches(2.6), Inches(0.5))
    name_frame = name_box.text_frame
    name_frame.text = character_name
    name_para = name_frame.paragraphs[0]
    name_para.font.size = Pt(20)
    name_para.font.bold = True
    name_para.font.color.rgb = RGBColor(255, 255, 255)
    name_para.alignment = PP_ALIGN.CENTER

    # Speech bubble
    bubble = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(3.5), Inches(1.5), Inches(5.5), Inches(4)
    )
    bubble.fill.solid()
    bubble.fill.fore_color.rgb = RGBColor(255, 255, 255)
    bubble.line.color.rgb = RGBColor(243, 156, 18)
    bubble.line.width = Pt(4)

    # Message
    msg_box = slide.shapes.add_textbox(Inches(4), Inches(2), Inches(4.5), Inches(1.5))
    msg_frame = msg_box.text_frame
    msg_frame.text = message
    msg_frame.word_wrap = True
    msg_para = msg_frame.paragraphs[0]
    msg_para.font.size = Pt(22)
    msg_para.font.bold = True
    msg_para.font.color.rgb = RGBColor(52, 73, 94)
    msg_para.alignment = PP_ALIGN.CENTER

    # Tip box
    tip_box = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(4), Inches(3.8), Inches(4.5), Inches(1.3)
    )
    tip_box.fill.solid()
    tip_box.fill.fore_color.rgb = RGBColor(241, 196, 15)
    tip_box.line.fill.background()

    tip_text = slide.shapes.add_textbox(Inches(4.3), Inches(4), Inches(3.9), Inches(0.9))
    tip_frame = tip_text.text_frame
    tip_frame.text = f"💡 Maslahat:\n{tip}"
    tip_frame.word_wrap = True
    tip_para = tip_frame.paragraphs[0]
    tip_para.font.size = Pt(16)
    tip_para.font.color.rgb = RGBColor(255, 255, 255)

    return slide

def create_final_slide(prs):
    """Create a congratulations final slide"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Background
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid()
    bg.fill.fore_color.rgb = RGBColor(46, 204, 113)  # Green
    bg.line.fill.background()

    # Congratulations box
    congrats_box = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(1.5), Inches(2), Inches(7), Inches(3)
    )
    congrats_box.fill.solid()
    congrats_box.fill.fore_color.rgb = RGBColor(255, 255, 255)
    congrats_box.line.color.rgb = RGBColor(39, 174, 96)
    congrats_box.line.width = Pt(5)

    # Trophy emoji
    trophy_text = slide.shapes.add_textbox(Inches(4.2), Inches(2.3), Inches(1.6), Inches(1))
    trophy_frame = trophy_text.text_frame
    trophy_frame.text = "🏆"
    trophy_para = trophy_frame.paragraphs[0]
    trophy_para.font.size = Pt(80)
    trophy_para.alignment = PP_ALIGN.CENTER

    # Congratulations text
    congrats_text = slide.shapes.add_textbox(Inches(2), Inches(3.5), Inches(6), Inches(0.8))
    congrats_frame = congrats_text.text_frame
    congrats_frame.text = "Tabriklaymiz!"
    congrats_para = congrats_frame.paragraphs[0]
    congrats_para.font.size = Pt(48)
    congrats_para.font.bold = True
    congrats_para.font.color.rgb = RGBColor(39, 174, 96)
    congrats_para.alignment = PP_ALIGN.CENTER

    # Message
    message_text = slide.shapes.add_textbox(Inches(2), Inches(4.3), Inches(6), Inches(0.5))
    message_frame = message_text.text_frame
    message_frame.text = "Siz Microsoft Word haqida ko'p narsalarni o'rgandingiz!"
    message_para = message_frame.paragraphs[0]
    message_para.font.size = Pt(20)
    message_para.font.color.rgb = RGBColor(52, 73, 94)
    message_para.alignment = PP_ALIGN.CENTER

    return slide

def main():
    """Create the complete presentation"""
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)

    # 1. Title slide
    create_title_slide(prs)

    # 2. Introduction
    create_intro_slide(prs)

    # 3. Character introduction
    create_character_slide(
        prs,
        "Ustoz Aziz",
        "Salom! Men sizga Microsoft Word dasturini o'rgataman!",
        "Diqqat bilan kuzating va savollarni javoblang!"
    )

    # 4. Features slide 1 - Basic features
    create_features_slide(
        prs,
        "Word ning asosiy imkoniyatlari",
        [
            "📄 Matn yozish va tahrirlash",
            "🎨 Shrift va ranglarni o'zgartirish",
            "📊 Jadvallar yaratish",
            "🖼️ Rasmlar qo'shish"
        ],
        RGBColor(41, 128, 185)
    )

    # 5. Quiz 1
    create_quiz_slide(
        prs,
        1,
        "Microsoft Word qanday dastur?",
        [
            ("A", "O'yin dasturi"),
            ("B", "Matn muharriri"),
            ("C", "Rasmlar uchun dastur"),
            ("D", "Musiqa dasturi")
        ],
        "B"
    )

    # 6. Character tip
    create_character_slide(
        prs,
        "Ustoz Aziz",
        "Juda yaxshi!",
        "Word da Ctrl+S tugmasi hujjatni saqlaydi!"
    )

    # 7. Features slide 2 - Formatting
    create_features_slide(
        prs,
        "Matnni formatlash",
        [
            "📝 Bold, Italic, Underline",
            "🎯 Matnni markazlash",
            "📏 Satr oralig'ini sozlash",
            "🔤 Shrift o'lchamini o'zgartirish"
        ],
        RGBColor(231, 76, 60)
    )

    # 8. Quiz 2
    create_quiz_slide(
        prs,
        2,
        "Hujjatni saqlash uchun qaysi tugmalar birikmasidan foydalanamiz?",
        [
            ("A", "Ctrl+C"),
            ("B", "Ctrl+V"),
            ("C", "Ctrl+S"),
            ("D", "Ctrl+Z")
        ],
        "C"
    )

    # 9. Features slide 3 - Advanced
    create_features_slide(
        prs,
        "Qo'shimcha imkoniyatlar",
        [
            "📋 Nusxa olish (Copy) - Ctrl+C",
            "📌 Qo'yish (Paste) - Ctrl+V",
            "↩️ Qaytarish (Undo) - Ctrl+Z",
            "🔍 Qidirish - Ctrl+F"
        ],
        RGBColor(142, 68, 173)
    )

    # 10. Quiz 3
    create_quiz_slide(
        prs,
        3,
        "Matnni nusxa olish uchun qaysi buyruqdan foydalanamiz?",
        [
            ("A", "Ctrl+X"),
            ("B", "Ctrl+C"),
            ("C", "Ctrl+P"),
            ("D", "Ctrl+O")
        ],
        "B"
    )

    # 11. Character tip 2
    create_character_slide(
        prs,
        "Ustoz Aziz",
        "Ajoyib natija!",
        "Har doim o'z ishingizni saqlab turing. Ctrl+S - eng muhim buyruq!"
    )

    # 12. Features slide 4 - Tables and Images
    create_features_slide(
        prs,
        "Jadval va rasmlar bilan ishlash",
        [
            "📊 Insert > Table - jadval qo'shish",
            "🖼️ Insert > Picture - rasm qo'shish",
            "📐 Jadval o'lchamini o'zgartirish",
            "🎨 Rasmlarni bezash va joylash"
        ],
        RGBColor(26, 188, 156)
    )

    # 13. Quiz 4
    create_quiz_slide(
        prs,
        4,
        "Word da rasm qo'shish uchun qaysi menyudan foydalanamiz?",
        [
            ("A", "File"),
            ("B", "Edit"),
            ("C", "Insert"),
            ("D", "View")
        ],
        "C"
    )

    # 14. Final tips character
    create_character_slide(
        prs,
        "Ustoz Aziz",
        "Siz barcha testlarni muvaffaqiyatli bajardingiz! 🎉",
        "Word dasturini har kuni mashq qiling va professional bo'ling!"
    )

    # 15. Final slide
    create_final_slide(prs)

    # Save presentation
    filename = "Microsoft_Word_Prezentatsiya_UZ.pptx"
    prs.save(filename)
    print(f"✅ Prezentatsiya muvaffaqiyatli yaratildi: {filename}")
    print(f"📊 Jami slaydlar soni: {len(prs.slides)}")
    print("\n🎨 Prezentatsiyada:")
    print("   • Interaktiv dizayn")
    print("   • Rangdor slaydlar")
    print("   • 4 ta test savoli")
    print("   • Animatsiya uchun tayyor")
    print("   • O'zbek tilida kontent")

if __name__ == "__main__":
    main()
