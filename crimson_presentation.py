"""クリムゾン（ヘルヴァボス）プレゼンテーション - 11枚構成"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

IMG = "/home/user/POWERPOINT/画像"

prs = Presentation()
W = 13.333
H = 7.5
prs.slide_width = Inches(W)
prs.slide_height = Inches(H)

# --- Colors ---
BLACK = RGBColor(0, 0, 0)
NEAR_BLACK = RGBColor(18, 12, 12)
DARK_RED = RGBColor(100, 15, 15)
CRIMSON = RGBColor(180, 30, 30)
GOLD = RGBColor(212, 175, 55)
DIM_GOLD = RGBColor(160, 130, 40)
WHITE = RGBColor(245, 240, 235)
LIGHT_GRAY = RGBColor(190, 185, 180)

BG = (18, 12, 12)  # Standard dark background


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
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = box.text_frame
    tf.word_wrap = True
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = line
        p.font.size = Pt(size)
        p.font.color.rgb = color
        p.font.bold = bold
        p.font.name = name
        p.alignment = align
        p.space_after = Pt(size * (spacing - 1))
    return box


def add_rect(slide, left, top, width, height, fill_color):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    shape.line.fill.background()
    return shape


def add_img(slide, path, left, top, width=None, height=None):
    kwargs = {}
    if width: kwargs['width'] = Inches(width)
    if height: kwargs['height'] = Inches(height)
    return slide.shapes.add_picture(path, Inches(left), Inches(top), **kwargs)


def gold_line(slide, left, top, width):
    return add_rect(slide, left, top, width, 0.04, GOLD)


def header_bar(slide, title):
    add_rect(slide, 0, 0, W, 1.2, DARK_RED)
    add_text(slide, 0.5, 0.15, 12, 0.9, title, size=42, color=WHITE, bold=True)
    gold_line(slide, 0.5, 1.2, W - 1.0)


# === Layout constants ===
# Left panel: image area 0 ~ 5.5 | gap | text area 6.0 ~ 12.8
# Right panel: text area 0.5 ~ 7.0 | gap | image area 7.5 ~ 12.8
IMG_LEFT_X = 0.8       # image on left side
TEXT_RIGHT_X = 6.0      # text when image is on left
TEXT_RIGHT_W = 6.8      # text width when image is on left
IMG_RIGHT_X = 8.0       # image on right side
TEXT_LEFT_X = 0.8       # text when image is on right
TEXT_LEFT_W = 6.8       # text width when image is on right
CONTENT_Y = 1.6         # y start below header


# ============================================================
# Slide 1: TITLE (テキスト左 + 画像右中央)
# ============================================================
s = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(s, *BG)

# Text on left
gold_line(s, 1.0, 2.0, 5.0)
add_text(s, 1.0, 2.3, 5.5, 1.5, "CRIMSON", size=76, color=CRIMSON, bold=True)
add_text(s, 1.0, 3.7, 5.5, 0.7, "Knolastname", size=36, color=GOLD)
add_text(s, 1.0, 4.5, 5.5, 0.7, "Greed Ring's Most Feared Mafia Boss",
         size=20, color=LIGHT_GRAY)
gold_line(s, 1.0, 5.4, 5.0)
add_text(s, 1.0, 6.2, 5.5, 0.5,
         "HELLUVA BOSS  |  Character Presentation", size=14, color=DIM_GOLD)

# Vertical gold accent
add_rect(s, 6.8, 0.5, 0.05, 6.5, GOLD)

# Image on right, vertically centered (924x904, ratio ~1.0, height=5.0 → width≈5.1)
add_img(s, f"{IMG}/crimson-1.JPG", 7.5, 1.25, height=5.0)

# ============================================================
# Slide 2: PROFILE (画像左 + テキスト右)
# ============================================================
s = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(s, *BG)

header_bar(s, "PROFILE")
# crimson-4.PNG: 697x916 (ratio 0.76), height=4.5 → width≈3.4
add_img(s, f"{IMG}/crimson-4.PNG", 1.0, 1.8, height=4.5)

info_lines = [
    "名前 :  Crimson（クリムゾン）",
    "種族 :  インプ",
    "リング :  強欲の環（Greed Ring）",
    "役割 :  マフィアボス / 犯罪王",
    "家族 :  モクシー（息子）",
    "声優 :  Richard Steven Horvitz",
    "初登場 :  \"Exes and Oohs\"（S2E6）",
]
add_multiline(s, 5.5, 1.8, 7.3, 4.0, info_lines, size=22, color=WHITE, spacing=1.5)

add_text(s, 5.5, 6.5, 7.3, 0.5,
         "強欲の環を支配するマフィアの頂点に立つインプ。",
         size=15, color=LIGHT_GRAY, align=PP_ALIGN.LEFT)

# ============================================================
# Slide 3: APPEARANCE (画像左 + テキスト右)
# ============================================================
s = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(s, *BG)

header_bar(s, "外見")
# crimson-3.PNG: transparent bg → dark panel behind
add_rect(s, 0.5, 1.5, 5.0, 5.7, NEAR_BLACK)
add_img(s, f"{IMG}/crimson-3.PNG", 0.8, 1.5, height=5.5)

desc_lines = [
    "Blitzoに近い長身のインプ",
    "白髪 / 白黒の縞模様の角",
    "黄色い強膜に金の牙（右側）",
    "ダークレッドの肌",
    "黒いフェドーラ帽（赤白バンド付き）",
    "ネイビーブルーのコートに赤いシャツ",
    "蹠行性（かかとを地面につける）の脚",
]
add_multiline(s, 7.0, 1.8, 5.8, 5.0, desc_lines, size=21, color=WHITE, spacing=1.6)

# ============================================================
# Slide 4: PERSONALITY (画像右 + テキスト左)
# ============================================================
s = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(s, *BG)

header_bar(s, "性格")
# crimson-7.JPG: square, height=4.5 → width≈4.5
add_img(s, f"{IMG}/crimson-7.JPG", 8.0, 1.8, height=4.8)

traits = [
    "◆ 残忍かつ冷酷なマフィアのボス",
    "◆ 絶対的な恐怖で組織を支配する",
    "◆ 必要な時には魅力的で温厚な顔も見せる",
    "◆ サメ悪魔のギャング軍団を統率",
    "◆ 脅迫・暴力・心理操作を駆使する策略家",
    "◆ 冷酷と魅力のギャップに満ちた男",
]
add_multiline(s, 0.8, 2.0, 6.8, 5.0, traits, size=22, color=WHITE, bold=True, spacing=1.5)

# ============================================================
# Slide 5: THE DON (画像右 + テキスト左)
# ============================================================
s = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(s, *BG)

add_rect(s, 0, 0, 6.5, H, NEAR_BLACK)
add_rect(s, 6.4, 0, 0.06, H, GOLD)

# crimson-6.JPG: landscape, height=6.5
add_img(s, f"{IMG}/crimson-6.JPG", 6.8, 0.5, height=6.5)

add_text(s, 0.5, 0.4, 5.5, 1.0, "THE DON", size=52, color=GOLD, bold=True)
gold_line(s, 0.5, 1.3, 5.0)

power_lines = [
    "強欲の環を牛耳る犯罪帝国の長。",
    "",
    "サメ悪魔の大軍を従え、",
    "恐喝と暴力で勢力を拡大する。",
    "",
    "財政危機に際しては",
    "チャズとの政略結婚すら画策する野心家。",
    "",
    "ストライカーを雇い、",
    "フィズロリを人質にアスモデウス",
    "すら脅迫する大胆さ。",
]
add_multiline(s, 0.8, 1.8, 5.2, 5.5, power_lines, size=20, color=WHITE, spacing=1.15)

# ============================================================
# Slide 6: ABILITIES (画像右 + テキスト左)
# ============================================================
s = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(s, *BG)

header_bar(s, "能力")
# crimson-8.JPG: 1334x750 (landscape), width=5.0 → height≈2.8
add_img(s, f"{IMG}/crimson-8.JPG", 7.8, 2.5, width=5.0)

abilities = [
    ["統率力", "絶対的なカリスマで巨大組織を率いる"],
    ["心理操作", "長年にわたりモクシーの精神を支配し、\n脅迫と心理戦で周囲を操る"],
    ["武器の扱い", "ナイフの投擲に秀でた腕前を持つ"],
    ["演技力", "温厚な紳士から冷酷な暴君まで、\n自在に顔を使い分ける"],
]

y = 1.7
for title, desc in abilities:
    add_text(s, 0.8, y, 6.5, 0.5, title, size=24, color=GOLD, bold=True, align=PP_ALIGN.LEFT)
    add_text(s, 0.8, y + 0.5, 6.5, 0.8, desc, size=17, color=LIGHT_GRAY, align=PP_ALIGN.LEFT)
    y += 1.4

# ============================================================
# Slide 7: CRIMSON & MOXXIE (画像左 + テキスト右) crimson-love.jpg
# ============================================================
s = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(s, *BG)

header_bar(s, "クリムゾンとモクシー")
# crimson-love.jpg: square-ish, height=4.8
add_img(s, f"{IMG}/crimson-love.jpg", 0.8, 1.8, height=4.8)

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
    "―― 不器用すぎる、歪んだ父の愛。",
]
add_multiline(s, 6.5, 1.8, 6.3, 5.5, moxxie_lines, size=22, color=WHITE, spacing=1.2)

# ============================================================
# Slide 8: DARK SIDE (画像左 + テキスト右)
# ============================================================
s = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(s, *BG)

header_bar(s, "THE DARK SIDE")
# crimson-9.JPG: landscape, height=4.0
add_img(s, f"{IMG}/crimson-9.JPG", 0.5, 2.0, height=4.0)

dark_lines = [
    "◆ 妻を「息子の成長を阻害する存在」",
    "　 として排除した",
    "",
    "◆ モクシーへの身体的・精神的虐待",
    "",
    "◆ 組織のためなら息子すら",
    "　 駒として使う冷徹さ",
    "",
    "◆ フィズロリを人質にし、",
    "　 七つの大罪の一柱アスモデウス",
    "　 すら脅迫する豪胆さ",
]
add_multiline(s, 7.0, 1.8, 5.8, 5.5, dark_lines, size=21, color=WHITE, bold=True, spacing=1.1)

# ============================================================
# Slide 9: KEY EPISODES (画像右 + テキスト左)
# ============================================================
s = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(s, *BG)

header_bar(s, "登場エピソード")
# crimson-5.JPG: square, height=4.5
add_img(s, f"{IMG}/crimson-5.JPG", 8.3, 2.0, height=4.5)

episodes = [
    ["Exes and Oohs（S2E6）",
     "初登場。モクシーをチャズと\n政略結婚させようと画策する。"],
    ["Oops（S2E7）",
     "ストライカーを雇い、フィズロリを\n人質にしてアスモデウスを脅迫する。"],
    ["Mammon's Magnificent Musical",
     "音楽スペシャルでの登場。"],
]

y = 1.8
for title, desc in episodes:
    add_text(s, 0.8, y, 7.0, 0.7, title, size=22, color=GOLD, bold=True, align=PP_ALIGN.LEFT)
    add_text(s, 0.8, y + 0.55, 7.0, 0.9, desc, size=17, color=LIGHT_GRAY, align=PP_ALIGN.LEFT)
    y += 1.7

# ============================================================
# Slide 10: WHY WE LOVE HIM (画像左 + テキスト右 / 参照レイアウト寄せ)
# ============================================================
s = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(s, *BG)

# crimson-2.PNG: 400x372 (ratio 1.08), width=5.0 → height≈4.65, vertically centered
add_img(s, f"{IMG}/crimson-2.PNG", 0.1, 1.4, width=5.0)

# 仕切り線を中央寄りに配置
add_rect(s, 5.45, 0, 0.04, H, GOLD)

add_text(s, 6.0, 0.5, 6.6, 0.6, "WHY WE", size=28, color=LIGHT_GRAY)
add_text(s, 6.0, 1.0, 6.6, 0.9, "LOVE HIM", size=52, color=CRIMSON, bold=True)
gold_line(s, 6.0, 1.9, 4.85)

appeal_lines = [
    "壊れきった魂の持ち主が持つ渋さと苦さ",
    "",
    "インプでありながら威厳を纏う貫禄",
    "",
    "歪んだ形で愛を表現する不器用さ",
    "",
    "傷を抱えながらも足掻き続ける強さ",
    "",
    "演技と話術で成り上がった努力の人",
    "",
    "「死ぬ技術」ではなく",
    "「生き延びる技術」を磨いた男",
]
add_multiline(s, 6.15, 2.35, 6.45, 4.95, appeal_lines, size=20, color=WHITE, spacing=1.1)

# ============================================================
# Slide 11: FANART (画像中央 + 下にクレジット)
# ============================================================
s = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(s, *BG)

add_img(s, f"{IMG}/crimson-fanart.PNG", 3.5, 0.3, height=5.5)

add_rect(s, 0, 6.1, W, 1.4, NEAR_BLACK)
gold_line(s, 0, 6.1, W)

add_text(s, 0.5, 6.25, 8.0, 0.5,
         "このパワーポイントの製作者が描いたファンアートです",
         size=18, color=GOLD, align=PP_ALIGN.LEFT)

add_text(s, 0.5, 6.75, 8.0, 0.5,
         "ご覧頂きありがとうございました。",
         size=16, color=LIGHT_GRAY, align=PP_ALIGN.LEFT)

add_text(s, 7.0, 6.75, 5.8, 0.5,
         "CRIMSON  —  Helluva Boss  |  Vivziepop",
         size=13, color=DIM_GOLD, align=PP_ALIGN.RIGHT)


# === Save ===
out = "/home/user/POWERPOINT/crimson_presentation.pptx"
prs.save(out)
print(f"Saved: {out}")
