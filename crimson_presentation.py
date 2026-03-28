"""クリムゾン（ヘルヴァボス）プレゼンテーション - 11枚構成"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
import os

IMG = "/home/user/POWERPOINT/画像"

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

# --- Color palette (mafia aesthetic) ---
BLACK = RGBColor(0, 0, 0)
NEAR_BLACK = RGBColor(12, 8, 8)
DARK_RED = RGBColor(100, 15, 15)
CRIMSON = RGBColor(165, 25, 25)
DEEP_CRIMSON = RGBColor(130, 10, 10)
GOLD = RGBColor(212, 175, 55)
DIM_GOLD = RGBColor(160, 130, 40)
WHITE = RGBColor(245, 240, 235)
LIGHT_GRAY = RGBColor(180, 175, 170)
DARK_GRAY = RGBColor(45, 40, 40)
NAVY = RGBColor(20, 22, 40)


def set_bg(slide, r, g, b):
    bg = slide.background.fill
    bg.solid()
    bg.fore_color.rgb = RGBColor(r, g, b)


def add_text(slide, left, top, width, height, text, size=32,
             color=WHITE, bold=False, align=PP_ALIGN.CENTER, name="Arial"):
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(size)
    p.font.color.rgb = color
    p.font.bold = bold
    p.font.name = name
    p.alignment = align
    return box


def add_multiline(slide, left, top, width, height, lines, size=20,
                  color=WHITE, bold=False, align=PP_ALIGN.LEFT, spacing=1.2, name="Arial"):
    """Add text box with multiple lines as separate paragraphs."""
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = box.text_frame
    tf.word_wrap = True
    for i, line in enumerate(lines):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = line
        p.font.size = Pt(size)
        p.font.color.rgb = color
        p.font.bold = bold
        p.font.name = name
        p.alignment = align
        p.space_after = Pt(size * (spacing - 1))
    return box


def add_rect(slide, left, top, width, height, fill_color, alpha=None):
    """Add a colored rectangle (decorative bar, overlay, etc.)."""
    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    shape.line.fill.background()
    if alpha is not None:
        # Set transparency via XML
        spPr = shape._element.spPr
        sf = spPr.find(qn('a:solidFill'))
        if sf is None:
            # solidFill is nested inside the shape fill
            ln_or_fill = spPr.findall('.//')
            for el in spPr.iter():
                if el.tag.endswith('}solidFill'):
                    sf = el
                    break
        if sf is not None:
            color_elem = sf.find(qn('a:srgbClr'))
            if color_elem is not None:
                a_elem = color_elem.makeelement(qn('a:alpha'), {})
                a_elem.set('val', str(int(alpha * 1000)))
                color_elem.append(a_elem)
    return shape


def add_img(slide, path, left, top, width=None, height=None):
    kwargs = {}
    if width: kwargs['width'] = Inches(width)
    if height: kwargs['height'] = Inches(height)
    return slide.shapes.add_picture(path, Inches(left), Inches(top), **kwargs)


def add_line_accent(slide, left, top, width, color=GOLD):
    """Thin gold accent line."""
    return add_rect(slide, left, top, width, 0.04, color)


# ============================================================
# Slide 1: TITLE
# ============================================================
s = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(s, 12, 8, 8)

# Full bleed image on right
add_img(s, f"{IMG}/crimson-1.JPG", 7.5, 0, height=7.5)
# Dark gradient overlay left
add_rect(s, 0, 0, 8.5, 7.5, BLACK, alpha=60)

# Accent lines
add_line_accent(s, 1.0, 1.8, 4.5)
add_line_accent(s, 1.0, 5.2, 4.5)

# Title text
add_text(s, 1.0, 2.1, 5.5, 1.5, "CRIMSON", size=72, color=CRIMSON, bold=True, name="Arial")
add_text(s, 1.0, 3.4, 5.5, 0.8, "Knolastname", size=36, color=GOLD, bold=False)
add_text(s, 1.0, 4.2, 5.5, 0.8, "Greed Ring's Most Feared Mafia Boss", size=20, color=LIGHT_GRAY)

add_text(s, 1.0, 6.0, 5.5, 0.5, "HELLUVA BOSS  |  Character Presentation", size=14, color=DIM_GOLD)

# ============================================================
# Slide 2: PROFILE
# ============================================================
s = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(s, 15, 10, 10)

add_rect(s, 0, 0, 13.333, 1.2, DARK_RED)
add_text(s, 0.5, 0.15, 12, 0.9, "PROFILE", size=42, color=WHITE, bold=True)
add_line_accent(s, 0.5, 1.2, 12.333)

# Image
add_img(s, f"{IMG}/crimson-4.PNG", 0.8, 1.8, height=5.0)

# Profile info right side
info_lines = [
    "Name :  Crimson (Knolastname)",
    "Species :  Imp Demon",
    "Ring :  Greed",
    "Role :  Mafia Boss / Crime Lord",
    "Family :  Moxxie (Son)",
    "VA :  Richard Steven Horvitz",
    "Debut :  \"Exes and Oohs\" (S2E6)",
]
add_multiline(s, 5.5, 2.0, 7.0, 4.5, info_lines, size=22, color=WHITE, spacing=1.8)

add_text(s, 5.5, 6.2, 7.0, 0.6,
         "Greed Ring of Hellを支配するマフィアの頂点に立つインプ。",
         size=16, color=LIGHT_GRAY, align=PP_ALIGN.LEFT)

# ============================================================
# Slide 3: APPEARANCE
# ============================================================
s = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(s, 10, 10, 18)

add_rect(s, 0, 0, 13.333, 1.2, NAVY)
add_text(s, 0.5, 0.15, 12, 0.9, "APPEARANCE", size=42, color=WHITE, bold=True)
add_line_accent(s, 0.5, 1.2, 12.333)

# Two images
add_img(s, f"{IMG}/crimson-3.PNG", 0.5, 1.5, height=5.5)
add_img(s, f"{IMG}/crimson-2.PNG", 4.5, 2.0, height=4.5)

# Description
desc_lines = [
    "Blitzoに近い長身のインプ",
    "白髪 / 白黒の縞模様の角",
    "黄色い強膜、金の牙（右）",
    "ダークレッドの肌",
    "黒いフェドーラ帽（赤白バンド）",
    "ネイビーブルーのコート＋赤いシャツ",
    "趾行性ではなく蹠行性の脚",
]
add_multiline(s, 8.0, 1.8, 4.8, 5.0, desc_lines, size=20, color=WHITE, spacing=1.6)

# ============================================================
# Slide 4: PERSONALITY - INTIMIDATION
# ============================================================
s = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(s, 5, 3, 3)

# Large dark image background
add_img(s, f"{IMG}/crimson-7.JPG", 0, 0, width=13.333)
# Overlay
add_rect(s, 0, 0, 13.333, 7.5, BLACK, alpha=55)

add_text(s, 0.5, 0.4, 12, 1.0, "PERSONALITY", size=48, color=CRIMSON, bold=True)
add_line_accent(s, 0.5, 1.3, 5.0)

traits = [
    "残忍かつ冷酷なマフィアのボス",
    "絶対的な恐怖で組織を支配",
    "しかし、必要な時には魅力的で温厚な顔も見せる",
    "サメ悪魔のギャング軍団を統率",
    "脅迫・暴力・心理操作を駆使",
    "金の牙のような凶悪なギャップに満ちている",
]
add_multiline(s, 0.8, 1.8, 6.5, 5.0, traits, size=22, color=WHITE, spacing=1.5)

# ============================================================
# Slide 5: MAFIA BOSS - POWER
# ============================================================
s = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(s, 8, 12, 8)

# Chair scene image
add_img(s, f"{IMG}/crimson-6.JPG", 6.5, 0.5, height=6.5)
add_rect(s, 6.0, 0, 0.06, 7.5, GOLD)

add_rect(s, 0, 0, 6.2, 7.5, NEAR_BLACK, alpha=85)

add_text(s, 0.5, 0.4, 5.5, 1.0, "THE DON", size=52, color=GOLD, bold=True)
add_line_accent(s, 0.5, 1.3, 4.5, CRIMSON)

power_lines = [
    "Greed Ringを牛耳る犯罪帝国の長",
    "",
    "サメ悪魔の大軍を従え、",
    "恐喝と暴力で勢力を拡大。",
    "",
    "財政危機に際してはチャズとの",
    "政略結婚すら画策する野心家。",
    "",
    "Strikerとその一味を雇い、",
    "Fizzarolliを人質にAsmodeus",
    "すら脅迫する大胆さ。",
]
add_multiline(s, 0.8, 1.8, 5.0, 5.5, power_lines, size=19, color=WHITE, spacing=1.2)

# ============================================================
# Slide 6: ABILITIES
# ============================================================
s = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(s, 15, 10, 10)

add_rect(s, 0, 0, 13.333, 1.2, DARK_RED)
add_text(s, 0.5, 0.15, 12, 0.9, "ABILITIES", size=42, color=WHITE, bold=True)
add_line_accent(s, 0.5, 1.2, 12.333)

# Image
add_img(s, f"{IMG}/crimson-8.JPG", 7.5, 1.5, width=5.3)

# Abilities list
abilities = [
    ["Leadership", "絶対的なカリスマで大組織を統率"],
    ["Manipulation", "長年にわたりモクシーの精神を支配\n脅迫と心理戦で周囲を操る"],
    ["Weapon Proficiency", "ナイフの投擲に秀でた腕前"],
    ["Acting", "温厚な紳士から冷酷な暴君まで\n自在に顔を使い分ける"],
]

y = 1.6
for title, desc in abilities:
    add_text(s, 0.8, y, 6.0, 0.5, title, size=24, color=GOLD, bold=True, align=PP_ALIGN.LEFT)
    add_text(s, 0.8, y + 0.45, 6.0, 0.8, desc, size=17, color=LIGHT_GRAY, align=PP_ALIGN.LEFT)
    y += 1.35

# ============================================================
# Slide 7: CRIMSON & MOXXIE (use crimson-love.jpg)
# ============================================================
s = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(s, 10, 15, 10)

# Full bg image
add_img(s, f"{IMG}/crimson-love.jpg", 0, 0, width=13.333)
add_rect(s, 0, 0, 13.333, 7.5, BLACK, alpha=50)

add_text(s, 0.5, 0.3, 12, 1.0, "CRIMSON  &  MOXXIE", size=48, color=GOLD, bold=True)
add_line_accent(s, 0.5, 1.2, 6.0)

# Two column feel
moxxie_lines = [
    "マフィアのボスの息子として、",
    "茨の道を歩ませなければならなかった。",
    "",
    "「殺らなきゃ殺られる」世界で",
    "息子が生き残れるよう厳しく育てた。",
    "",
    "家を出て結婚した息子を、",
    "すぐに追わず静かに見守っていた。",
    "",
    "不器用すぎる、歪んだ父の愛。",
]
add_multiline(s, 0.8, 1.6, 6.0, 5.5, moxxie_lines, size=21, color=WHITE, spacing=1.15)

# ============================================================
# Slide 8: DARK SIDE - VIOLENCE
# ============================================================
s = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(s, 5, 3, 3)

add_img(s, f"{IMG}/crimson-9.JPG", 0, 0, width=13.333)
add_rect(s, 0, 0, 13.333, 7.5, BLACK, alpha=45)

add_text(s, 0.5, 0.3, 12, 1.0, "THE DARK SIDE", size=48, color=CRIMSON, bold=True)
add_line_accent(s, 0.5, 1.2, 5.0)

dark_lines = [
    "妻を「息子の成長を阻害する存在」として排除",
    "",
    "モクシーへの身体的・精神的虐待",
    "",
    "組織のためなら息子すら駒として使う冷徹さ",
    "",
    "Fizzarolliを人質にし、七つの大罪の一柱",
    "Asmoデウスすら脅迫する豪胆さ",
]
add_multiline(s, 0.8, 1.6, 7.0, 5.5, dark_lines, size=21, color=WHITE, spacing=1.2)

# ============================================================
# Slide 9: KEY EPISODES
# ============================================================
s = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(s, 15, 10, 10)

add_rect(s, 0, 0, 13.333, 1.2, DARK_RED)
add_text(s, 0.5, 0.15, 12, 0.9, "KEY EPISODES", size=42, color=WHITE, bold=True)
add_line_accent(s, 0.5, 1.2, 12.333)

add_img(s, f"{IMG}/crimson-5.JPG", 8.5, 1.8, height=4.5)

episodes = [
    ["Exes and Oohs  (S2E6)", "初登場。モクシーをチャズと\n政略結婚させようと画策。"],
    ["Oops  (S2E7)", "Strikerを雇い、Fizzarolliを\n人質にしてAsmodeusを脅迫。"],
    ["Mammon's Magnificent Musical\nMid-Season Special", "音楽スペシャルでの登場。"],
]

y = 1.6
for title, desc in episodes:
    add_text(s, 0.8, y, 7.0, 0.6, title, size=22, color=GOLD, bold=True, align=PP_ALIGN.LEFT)
    add_text(s, 0.8, y + 0.55, 7.0, 0.9, desc, size=17, color=LIGHT_GRAY, align=PP_ALIGN.LEFT)
    y += 1.6

# ============================================================
# Slide 10: CHARACTER APPEAL
# ============================================================
s = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(s, 8, 5, 12)

add_img(s, f"{IMG}/crimson-8.JPG", 0, 0, width=7.0)
add_rect(s, 0, 0, 7.0, 7.5, BLACK, alpha=30)
add_rect(s, 6.8, 0, 0.06, 7.5, GOLD)

add_rect(s, 6.86, 0, 6.5, 7.5, NEAR_BLACK, alpha=85)

add_text(s, 7.2, 0.4, 5.5, 1.0, "WHY WE", size=28, color=LIGHT_GRAY, bold=False)
add_text(s, 7.2, 0.9, 5.5, 1.0, "LOVE HIM", size=48, color=CRIMSON, bold=True)
add_line_accent(s, 7.2, 1.8, 4.5)

appeal_lines = [
    "壊れきった魂の持ち主が持つ渋さと苦さ",
    "",
    "インプでありながら威厳を纏う貫禄",
    "",
    "歪んだ形で愛を表現する不器用さ",
    "",
    "PTSDを抱えながらも足掻き続ける強さ",
    "",
    "演技と話術で成り上がった努力の人",
    "",
    "「死ぬ技術」ではなく",
    "「生き延びる技術」を磨いた男",
]
add_multiline(s, 7.4, 2.2, 5.3, 5.0, appeal_lines, size=19, color=WHITE, spacing=1.1)

# ============================================================
# Slide 11: LAST PAGE - FANART (crimson-fanart.PNG)
# ============================================================
s = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(s, 12, 8, 8)

# Center the fanart
add_img(s, f"{IMG}/crimson-fanart.PNG", 3.5, 0.3, height=5.5)

# Decorative bars
add_rect(s, 0, 6.2, 13.333, 1.3, NEAR_BLACK)
add_line_accent(s, 0, 6.2, 13.333)

add_text(s, 0.5, 6.35, 8.0, 0.5,
         "Fanart by the presentation creator",
         size=18, color=GOLD, align=PP_ALIGN.LEFT)

add_text(s, 0.5, 6.75, 8.0, 0.5,
         "Thank you for watching.",
         size=16, color=LIGHT_GRAY, align=PP_ALIGN.LEFT)

add_text(s, 7.0, 6.35, 5.8, 0.5,
         "CRIMSON  -  Character Presentation",
         size=14, color=DIM_GOLD, align=PP_ALIGN.RIGHT)

add_text(s, 7.0, 6.75, 5.8, 0.5,
         "Helluva Boss  |  Vivziepop",
         size=12, color=LIGHT_GRAY, align=PP_ALIGN.RIGHT)


# === Save ===
out = "/home/user/POWERPOINT/crimson_presentation.pptx"
prs.save(out)
print(f"Saved: {out}")
