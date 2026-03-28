"""アダムは騒音である - シュールなプレゼンテーション"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

# --- Color palette ---
BLACK = RGBColor(0, 0, 0)
WHITE = RGBColor(255, 255, 255)
RED = RGBColor(200, 30, 30)
GOLD = RGBColor(255, 215, 0)
DARK_RED = RGBColor(80, 10, 10)
GRAY = RGBColor(60, 60, 60)


def add_bg(slide, color):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_textbox(slide, left, top, width, height, text, font_size=32,
                color=WHITE, bold=False, alignment=PP_ALIGN.CENTER, font_name="Arial"):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(font_size)
    p.font.color.rgb = color
    p.font.bold = bold
    p.font.name = font_name
    p.alignment = alignment
    return txBox


def add_shape_with_text(slide, shape_type, left, top, width, height, text,
                        fill_color=None, font_size=24, font_color=WHITE, bold=False):
    shape = slide.shapes.add_shape(shape_type, left, top, width, height)
    if fill_color:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_color
    else:
        shape.fill.background()
    shape.line.fill.background()
    tf = shape.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(font_size)
    p.font.color.rgb = font_color
    p.font.bold = bold
    p.alignment = PP_ALIGN.CENTER
    tf.paragraphs[0].space_before = Pt(0)
    tf.paragraphs[0].space_after = Pt(0)
    return shape


# ============================================================
# Slide 1: タイトル - 「この男、騒音である。」
# ============================================================
slide1 = prs.slides.add_slide(prs.slide_layouts[6])  # blank
add_bg(slide1, BLACK)

add_textbox(slide1, Inches(1), Inches(0.5), Inches(11.333), Inches(1.5),
            "緊 急 報 告",
            font_size=28, color=RED, bold=True)

add_textbox(slide1, Inches(1), Inches(1.5), Inches(11.333), Inches(2.5),
            "この男、騒音である。",
            font_size=72, color=WHITE, bold=True)

# アダムの名前
add_textbox(slide1, Inches(2), Inches(4.2), Inches(9.333), Inches(1),
            "― 天界初の人間 アダム ―",
            font_size=36, color=GOLD, bold=True)

add_textbox(slide1, Inches(2), Inches(5.5), Inches(9.333), Inches(1),
            "※ 本プレゼンは天界騒音対策委員会の調査に基づきます",
            font_size=16, color=RGBColor(150, 150, 150))

# ============================================================
# Slide 2: 「騒音レベルの比較」
# ============================================================
slide2 = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide2, RGBColor(15, 15, 25))

add_textbox(slide2, Inches(0.5), Inches(0.3), Inches(12.333), Inches(1),
            "騒音レベル比較表",
            font_size=48, color=WHITE, bold=True)

add_textbox(slide2, Inches(0.5), Inches(1.2), Inches(12.333), Inches(0.6),
            "～ アダムはどれくらいうるさいのか ～",
            font_size=22, color=RGBColor(180, 180, 180))

# Bar chart style comparison
items = [
    ("ジェット機の離陸", "130 dB", 5.0, RGBColor(100, 100, 180)),
    ("ロックコンサート", "120 dB", 4.6, RGBColor(120, 100, 180)),
    ("アダムの通常会話", "347 dB", 11.0, RGBColor(255, 50, 50)),
    ("アダムの歌", "982 dB", 11.0, RGBColor(255, 0, 0)),
    ("図書館", "30 dB", 1.2, RGBColor(80, 180, 80)),
]

y_start = 2.0
for i, (label, db, bar_width, color) in enumerate(items):
    y = y_start + i * 1.05
    add_textbox(slide2, Inches(0.3), Inches(y), Inches(3.2), Inches(0.7),
                label, font_size=20, color=WHITE, alignment=PP_ALIGN.RIGHT)
    # bar
    bar = slide2.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(3.7), Inches(y + 0.1), Inches(bar_width), Inches(0.5))
    bar.fill.solid()
    bar.fill.fore_color.rgb = color
    bar.line.fill.background()
    # db label
    add_textbox(slide2, Inches(3.7 + bar_width + 0.1), Inches(y), Inches(2), Inches(0.7),
                db, font_size=18, color=color, bold=True, alignment=PP_ALIGN.LEFT)

add_textbox(slide2, Inches(1), Inches(6.8), Inches(11.333), Inches(0.5),
            "※ アダムの数値は推定値です。測定器が3台壊れました。",
            font_size=14, color=RGBColor(255, 100, 100))

# ============================================================
# Slide 3: 「被害報告」
# ============================================================
slide3 = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide3, RGBColor(20, 10, 10))

add_textbox(slide3, Inches(0.5), Inches(0.3), Inches(12.333), Inches(1),
            "被 害 報 告",
            font_size=52, color=RED, bold=True)

reports = [
    "✦ 天使の翼が音圧で2枚もげた",
    "✦ ルシファーが「あいつだけはマジで無理」と発言",
    "✦ 天界のステンドグラスが全壊（3回目）",
    "✦ アダムの歌を聴いた天使の87%が\n　 「地獄の方がまだ静か」と回答",
    "✦ チャーリーが耳栓を地獄の新産業にすることを検討中",
]

y = 1.5
for report in reports:
    add_textbox(slide3, Inches(1), Inches(y), Inches(11.333), Inches(1),
                report, font_size=28, color=WHITE, alignment=PP_ALIGN.LEFT)
    y += 1.05

add_textbox(slide3, Inches(1), Inches(6.7), Inches(11.333), Inches(0.5),
            "出典：天界安全衛生委員会 年次報告書（非公開）",
            font_size=14, color=RGBColor(150, 150, 150))

# ============================================================
# Slide 4: 「本人の反応」
# ============================================================
slide4 = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide4, RGBColor(10, 10, 20))

add_textbox(slide4, Inches(0.5), Inches(0.3), Inches(12.333), Inches(1),
            "本人の反応",
            font_size=48, color=GOLD, bold=True)

# Big quote
quote_shape = slide4.shapes.add_shape(
    MSO_SHAPE.ROUNDED_RECTANGLE,
    Inches(1.5), Inches(1.8), Inches(10.333), Inches(3))
quote_shape.fill.solid()
quote_shape.fill.fore_color.rgb = RGBColor(30, 30, 50)
quote_shape.line.color.rgb = GOLD

tf = quote_shape.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "「俺の歌が騒音？\n　ハッ、お前らの耳がついてこれねぇだけだろ！\n　俺は天界一のロックスターだぞ！！」"
p.font.size = Pt(32)
p.font.color.rgb = WHITE
p.font.bold = True
p.alignment = PP_ALIGN.CENTER

add_textbox(slide4, Inches(6), Inches(5.0), Inches(6), Inches(0.6),
            "― アダム（反省の色なし）",
            font_size=22, color=RGBColor(180, 180, 180), alignment=PP_ALIGN.RIGHT)

# Additional commentary
add_textbox(slide4, Inches(1.5), Inches(5.8), Inches(10.333), Inches(1.2),
            "委員会コメント：「予想通りの反応でした。\n"
            "　なお、この発言時の音量も147dBを記録しています。」",
            font_size=20, color=RGBColor(200, 200, 200))

# ============================================================
# Slide 5: 「結論」
# ============================================================
slide5 = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide5, BLACK)

add_textbox(slide5, Inches(0.5), Inches(0.5), Inches(12.333), Inches(1.5),
            "結　論",
            font_size=52, color=WHITE, bold=True)

# Big conclusion box
conclusion = slide5.shapes.add_shape(
    MSO_SHAPE.RECTANGLE,
    Inches(1.5), Inches(2.2), Inches(10.333), Inches(2.5))
conclusion.fill.solid()
conclusion.fill.fore_color.rgb = DARK_RED
conclusion.line.color.rgb = RED

tf = conclusion.text_frame
tf.word_wrap = True
tf.paragraphs[0].alignment = PP_ALIGN.CENTER
p = tf.paragraphs[0]
p.text = "アダムは騒音である。"
p.font.size = Pt(64)
p.font.color.rgb = WHITE
p.font.bold = True

add_textbox(slide5, Inches(1), Inches(5.0), Inches(11.333), Inches(0.8),
            "対策案：耳栓の全天使への配布（予算：天界GDP の 12%）",
            font_size=24, color=GOLD)

add_textbox(slide5, Inches(1), Inches(5.8), Inches(11.333), Inches(0.8),
            "しかし本人が耳栓を拒否しているため、根本的解決は不可能。",
            font_size=22, color=RGBColor(200, 200, 200))

add_textbox(slide5, Inches(1), Inches(6.6), Inches(11.333), Inches(0.5),
            "ご清聴ありがとうございました（耳、大丈夫ですか？）",
            font_size=18, color=RGBColor(120, 120, 120))

# Save
output_path = "/home/user/POWERPOINT/adam_souon.pptx"
prs.save(output_path)
print(f"Saved: {output_path}")
