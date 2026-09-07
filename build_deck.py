# -*- coding: utf-8 -*-
"""《中国艺术设计史》——「形式追随功能」主线课件生成脚本
纹样配图用 PIL 程序化绘制（抽象几何母题，非文物实拍），
课件用 python-pptx 生成。
"""
import os, math, random
from PIL import Image, ImageDraw

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

BASE = os.path.dirname(os.path.abspath(__file__))
ASSET = os.path.join(BASE, "assets")
os.makedirs(ASSET, exist_ok=True)

# ---------------- 配色（朱砂·黛青·赭金·宣纸） ----------------
PAPER   = (247, 242, 233)   # 宣纸米白
CARD    = (255, 253, 248)   # 卡片暖白
PRIMARY = (154, 59, 46)     # 朱砂红
PRIMARY_D = (122, 43, 33)
SECOND  = (39, 71, 83)      # 黛青
SECOND_D= (27, 50, 59)
ACCENT  = (194, 149, 63)    # 赭金
INK     = (42, 38, 32)      # 墨色
MUTED   = (134, 126, 114)   # 灰褐
LINEC   = (225, 216, 198)   # 卡边线

FONT_SERIF = "宋体"
FONT_BODY  = "微软雅黑"

SW, SH = 13.333, 7.5
MARGIN = 0.5

# ================= 纹样母题绘制 =================
def _c(size, bg=PAPER):
    img = Image.new("RGB", (size, size), bg)
    return img, ImageDraw.Draw(img)

def _frame(d, s, m=34, color=INK, w=5):
    d.rectangle([m, m, s-m, s-m], outline=color, width=w)

def motif_swirl(s=800):      # 旋涡纹（原始彩陶）
    img, d = _c(s)
    cx = cy = s//2
    r = int(s*0.42); step = r//6
    cols = [PRIMARY, SECOND, ACCENT, PRIMARY, SECOND, ACCENT]
    for i in range(6):
        rr = r - i*step
        d.ellipse([cx-rr, cy-rr, cx+rr, cy+rr], outline=cols[i], width=11)
    for k in range(12):
        a = k*math.pi/6
        d.line([cx+int(math.cos(a)*r*0.52), cy+int(math.sin(a)*r*0.52),
                cx+int(math.cos(a)*r), cy+int(math.sin(a)*r)],
               fill=ACCENT, width=6)
    d.ellipse([cx-16, cy-16, cx+16, cy+16], fill=PRIMARY)
    _frame(d, s)
    return img

def motif_wave(s=800):       # 水波纹（陶器）
    img, d = _c(s)
    for row in range(8):
        y0 = 60 + row*92
        pts = []
        for x in range(s+20):
            pts.append((x, y0 + int((18 if row % 2 else 14)*math.sin((x/ (42+row*3)) + row))))
        c = [SECOND, PRIMARY, ACCENT, SECOND, PRIMARY, ACCENT, SECOND, PRIMARY][row]
        d.line(pts, fill=c, width=10)
    _frame(d, s)
    return img

def motif_squareback(s=800):  # 回字纹·方胜
    img, d = _c(s)
    cx = cy = s//2
    cols = [PRIMARY, SECOND, ACCENT]
    for i in range(7):
        r = int(s*0.44) - i*int(s*0.055)
        d.rectangle([cx-r, cy-r, cx+r, cy+r], outline=cols[i%3], width=10)
    for dx, dy in [(-1,-1),(1,-1),(-1,1),(1,1)]:
        xx, yy = cx+dx*int(s*0.30), cy+dy*int(s*0.30)
        d.rectangle([xx-34, yy-34, xx+34, yy+34], outline=PRIMARY, width=8)
    _frame(d, s)
    return img

def motif_thunder(s=800):    # 雷纹·乳钉
    img, d = _c(s)
    n = 6; g = s//n
    for i in range(n):
        for j in range(n):
            x, y = i*g + 30, j*g + 30
            u = g - 60
            d.rectangle([x, y, x+u, y+u], outline=INK, width=3)
            d.ellipse([x+u//2-9, y+u//2-9, x+u//2+9, y+u//2+9], fill=PRIMARY)
    _frame(d, s)
    return img

def motif_diamond(s=800):    # 菱格纹（织锦）
    img, d = _c(s)
    step = 80
    for row in range(-1, 11):
        for col in range(-1, 11):
            x = col*step + (40 if row % 2 else 0)
            y = 60 + row*step
            p = [(x, y), (x+step//2, y-step//2), (x+step, y), (x+step//2, y+step//2)]
            d.polygon(p, outline=[SECOND, PRIMARY, ACCENT][(row+col) % 3], width=4)
            d.point((x+step//2, y), fill=ACCENT)
    _frame(d, s)
    return img

def motif_cloud(s=800):      # 云纹（瓦当）
    img, d = _c(s)
    cx = cy = s//2
    R = int(s*0.40)
    d.ellipse([cx-R, cy-R, cx+R, cy+R], outline=PRIMARY, width=10)
    for k in range(4):
        a = k*math.pi/2
        bx = cx + int(math.cos(a)*R*0.72)
        by = cy + int(math.sin(a)*R*0.72)
        for j in range(3):
            rr = 34 - j*8
            d.arc([bx-rr, by-rr, bx+rr, by+rr], 20, 320, fill=ACCENT if j else PRIMARY, width=7)
    d.ellipse([cx-18, cy-18, cx+18, cy+18], fill=SECOND)
    _frame(d, s)
    return img

def motif_lotus(s=800):      # 莲瓣纹
    img, d = _c(s)
    cx = cy = s//2
    R = int(s*0.20)
    d.ellipse([cx-R, cy-R, cx+R, cy+R], fill=ACCENT)
    for ring, (r, num, w, c) in enumerate([(0.28, 8, 66, PRIMARY), (0.40, 12, 58, SECOND)]):
        rr = int(s*r)
        for k in range(num):
            a = k*2*math.pi/num
            px, py = cx+int(math.cos(a)*rr), cy+int(math.sin(a)*rr)
            d.ellipse([px-w//2, py-w//2, px+w//2, py+w//2], outline=c, width=7)
    _frame(d, s)
    return img

def motif_roundel(s=800):    # 宝相花
    img, d = _c(s)
    cx = cy = s//2
    for ring, (r, num, w, c) in enumerate([(0.20, 6, 60, ACCENT), (0.34, 10, 64, PRIMARY), (0.44, 14, 46, SECOND)]):
        rr = int(s*r)
        for k in range(num):
            a = (k+ (ring*0.19))*2*math.pi/num
            px, py = cx+int(math.cos(a)*rr), cy+int(math.sin(a)*rr)
            d.ellipse([px-w//2, py-w//2, px+w//2, py+w//2], fill=c if ring < 2 else None, outline=c, width=5)
    d.ellipse([cx-16, cy-16, cx+16, cy+16], fill=INK)
    _frame(d, s)
    return img

def motif_crackle(s=800):    # 冰裂纹（宋瓷）
    img, d = _c(s)
    random.seed(7)
    pts = [(random.randint(60, s-60), random.randint(60, s-60)) for _ in range(26)]
    for i, p in enumerate(pts):
        x, y = p
        for _ in range(3):
            nx = x + random.randint(-70, 70); ny = y + random.randint(-70, 70)
            nx = max(60, min(s-60, nx)); ny = max(60, min(s-60, ny))
            d.line([x, y, nx, ny], fill=SECOND, width=2)
            x, y = nx, ny
    for x, y in pts:
        d.ellipse([x-4, y-4, x+4, y+4], fill=PRIMARY)
    _frame(d, s)
    return img

def motif_searock(s=800):    # 海水江崖
    img, d = _c(s)
    for row in range(7):
        y0 = s - 80 - row*70
        pts = [(x, y0 - int((26 if row % 2 else 20)*math.sin((x/58)+row*1.3))) for x in range(s+20)]
        pts.append((s+20, s)); pts.append((0, s))
        d.polygon(pts, outline=SECOND, width=4)
    rock = [(s//2, 120), (s//2-90, 320), (s//2-40, 320), (s//2-20, 420),
            (s//2+20, 420), (s//2+40, 320), (s//2+90, 320)]
    d.polygon(rock, fill=PRIMARY)
    d.ellipse([s//2-30, 150, s//2+30, 210], outline=ACCENT, width=5)
    _frame(d, s)
    return img

def motif_geo(s=800):        # 几何构成（近现代）
    img, d = _c(s)
    n = 4; g = s//n
    d.rectangle([g, g, 3*g, 3*g], fill=SECOND)
    for i in range(0, s, g):
        d.line([i, 0, i, s], fill=INK, width=4)
        d.line([0, i, s, i], fill=INK, width=4)
    d.ellipse([int(s*0.55), int(s*0.10), int(s*0.88), int(s*0.43)], fill=PRIMARY)
    d.polygon([(int(s*0.12), int(s*0.72)), (int(s*0.30), int(s*0.72)),
               (int(s*0.21), int(s*0.88))], fill=ACCENT)
    d.rectangle([0, 0, s, s], outline=INK, width=6)
    return img

MOTIFS = {
    "swirl":  motif_swirl, "wave": motif_wave, "squareback": motif_squareback,
    "thunder": motif_thunder, "diamond": motif_diamond, "cloud": motif_cloud,
    "lotus": motif_lotus, "roundel": motif_roundel, "crackle": motif_crackle,
    "searock": motif_searock, "geo": motif_geo,
}

def make_assets(size=800):
    for name, fn in MOTIFS.items():
        p = os.path.join(ASSET, f"{name}_{size}x{size}.png")
        fn(size).save(p)
    # 封面徽记（复用宝相花，更大）
    p = os.path.join(ASSET, f"emblem_{size}x{size}.png")
    motif_roundel(size).save(p)

# ---------------- PPTX 工具 ----------------
def rgb(hx):
    return RGBColor.from_string(hx)

def set_font(run, name, size, bold=False, color=None, italic=False):
    f = run.font
    f.name = name; f.size = Pt(size); f.bold = bold; f.italic = italic
    if color is not None:
        f.color.rgb = rgb(color)
    rPr = run._r.get_or_add_rPr()
    ea = rPr.find(qn('a:ea'))
    if ea is None:
        ea = rPr.makeelement(qn('a:ea'), {}); rPr.append(ea)
    ea.set('typeface', name)

def add_box(slide, x, y, w, h, anchor=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True; tf.vertical_anchor = anchor
    tf.margin_left = 0; tf.margin_right = 0; tf.margin_top = 0; tf.margin_bottom = 0
    return tf

def para(tf, text, *, size=14, bold=False, color="2A2620", font=FONT_BODY,
         align=PP_ALIGN.LEFT, first=False, after=6, leading=None, italic=False):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.alignment = align
    if after is not None: p.space_after = Pt(after)
    if leading is not None: p.line_spacing = leading
    r = p.add_run(); r.text = text
    set_font(r, font, size, bold, color, italic)
    return p

def bg(slide, hx):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = rgb(hx)

def shape(slide, kind, x, y, w, h, fill=None, line=None, line_w=1.0, radius=None):
    sp = slide.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid(); sp.fill.fore_color.rgb = rgb(fill)
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = rgb(line); sp.line.width = Pt(line_w)
    sp.shadow.inherit = False
    return sp

def footer(slide, n):
    add_box(slide, MARGIN, 7.02, 6.0, 0.3)
    tf = add_box(slide, MARGIN, 7.02, 6.0, 0.3)
    para(tf, "《中国艺术设计史》· 依「形式追随功能」重读", size=9, color="867E72", first=True, after=0)
    tf2 = add_box(slide, SW-MARGIN-1.4, 7.02, 1.4, 0.3)
    para(tf2, f"{n:02d} / 19", size=9, color="867E72", first=True, after=0, align=PP_ALIGN.RIGHT)

# ================= 生成配图资源 =================
make_assets()

# ================= 幻灯片构建 =================
prs = Presentation()
prs.slide_width = Inches(SW)
prs.slide_height = Inches(SH)
BLANK = prs.slide_layouts[6]

def new_slide():
    return prs.slides.add_slide(BLANK)

# ---- S1 封面 ----
s = new_slide(); bg(s, "F7F2E9")
# 顶部/底部朱砂色带
shape(s, MSO_SHAPE.RECTANGLE, 0, 0, SW, 0.18, fill="9A3B2E")
shape(s, MSO_SHAPE.RECTANGLE, 0, 0.18, SW, 0.06, fill="C2953F")
shape(s, MSO_SHAPE.RECTANGLE, 0, SH-0.18, SW, 0.18, fill="274753")
shape(s, MSO_SHAPE.RECTANGLE, 0, SH-0.24, SW, 0.06, fill="C2953F")
em = os.path.join(ASSET, "emblem_800x800.png")
s.shapes.add_picture(em, Inches(8.9), Inches(1.35), height=Inches(4.4))
tf = add_box(s, 0.9, 1.7, 7.2, 3.6)
para(tf, "中国艺术设计史", size=54, bold=True, color="2A2620", font=FONT_SERIF, first=True, after=10)
para(tf, "依「形式追随功能」重读十章", size=26, bold=True, color="9A3B2E", font=FONT_SERIF, after=14)
para(tf, "夏燕靖　著 · 第 3 版", size=15, color="867E72", after=4)
para(tf, "设计与技术 · 社会生活 · 审美观念 的关系主线", size=13, color="867E72", after=0)
footer(s, 1)

# ---- S2 目录 ----
s = new_slide(); bg(s, "F7F2E9")
shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, 0.6, 0.55, 0.09, 1.0, fill="9A3B2E")
tf = add_box(s, 0.9, 0.6, 9.0, 1.0)
para(tf, "目录", size=32, bold=True, color="2A2620", font=FONT_SERIF, first=True, after=0)
items = [
    ("01", "核心框架", "形式追随功能的三重内涵 —— 实用 · 礼制象征 · 审美生活"),
    ("02", "演进主线", "全书功能重心览 —— 生存巫术 → 工业现代（六阶段）"),
    ("03", "十章精读", "从萌芽到近现代，逐章看「功能如何决定形式」"),
    ("04", "知识拓展·出处", "每章一个知识点，附维基百科来源与参考文献"),
    ("05", "贯穿性结论", "三条主线与对现代设计的启示"),
]
y = 1.70
for num, t, d in items:
    shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, 0.9, y, 11.5, 0.94, fill="FFFDF8", line="E1D8C6", line_w=1.0)
    tf = add_box(s, 1.3, y+0.1, 0.9, 0.74)
    para(tf, num, size=22, bold=True, color="C2953F", font=FONT_SERIF, first=True, after=0)
    tf = add_box(s, 2.4, y+0.08, 9.6, 0.8)
    para(tf, t, size=17, bold=True, color="274753", font=FONT_SERIF, first=True, after=2)
    para(tf, d, size=11, color="867E72", after=0)
    y += 1.06
footer(s, 2)

# ---- S3 核心框架 ----
s = new_slide(); bg(s, "F7F2E9")
shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, 0.6, 0.55, 0.09, 1.0, fill="9A3B2E")
tf = add_box(s, 0.9, 0.6, 9.0, 1.0)
para(tf, "核心框架 · 形式追随功能", size=32, bold=True, color="2A2620", font=FONT_SERIF, first=True, after=0)
cards = [
    ("实用功能", "汲水 · 炊煮 · 坐卧 · 农战 · 生产", "器形直接由用途塑造，如尖底瓶为汲水而生、铁器随农战普及。"),
    ("礼制象征功能", "等级 · 通神 · 礼制 · 天人感应", "鼎尊簋、饕餮纹、服饰章纹皆由「礼」决定，是中国设计根本特色。"),
    ("审美·生活功能", "尚意 · 文人趣味 · 商品流通", "宋瓷之简约、明式家具之合度、版画之传布，形式追随审美与市场。"),
]
x = 0.7; w = 3.9
for i, (t, sub, body) in enumerate(cards):
    cx = x + i*(w+0.22)
    shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, cx, 1.8, w, 3.3, fill="FFFDF8", line="E1D8C6", line_w=1.0)
    top = ["9A3B2E", "274753", "C2953F"][i]
    shape(s, MSO_SHAPE.RECTANGLE, cx+0.25, 1.8, 0.5, 0.07, fill=top)
    tf = add_box(s, cx+0.28, 2.1, w-0.56, 2.8)
    para(tf, t, size=19, bold=True, color="2A2620", font=FONT_SERIF, first=True, after=4)
    para(tf, sub, size=11, color=top, bold=True, after=8)
    para(tf, body, size=12.5, color="2A2620", after=0, leading=1.25)
b = shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, 0.7, 5.45, SW-1.4, 1.1, fill="274753")
tf = add_box(s, 1.05, 5.6, SW-2.1, 0.85)
para(tf, "技术 · 材料是中介变量", size=14, bold=True, color="C2953F", first=True, after=2)
para(tf, "《考工记》「天有时 · 地有气 · 材有美 · 工有巧」——功能提出要求，材料与工艺决定形式的上限。",
     size=12, color="F7F2E9", after=0)
footer(s, 3)

# ---- S4 演进主线 ----
s = new_slide(); bg(s, "F7F2E9")
shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, 0.6, 0.55, 0.09, 1.0, fill="9A3B2E")
tf = add_box(s, 0.9, 0.6, 10.0, 1.0)
para(tf, "演进主线 · 全书功能重心一览", size=32, bold=True, color="2A2620", font=FONT_SERIF, first=True, after=0)
stages = [
    ("萌芽期", "生存·巫术", "工具·纹饰", "9A3B2E"),
    ("青铜期", "礼制·象征", "礼器", "274753"),
    ("铁器期", "实用·技术", "实用器·考工", "C2953F"),
    ("大一统", "大一统·宗教", "陵墓·石窟", "9A3B2E"),
    ("宋元明清", "审美·商品", "文人器物", "274753"),
    ("近现代", "工业·现代", "现代设计", "C2953F"),
]
colw = 1.9; gap = 0.17; x0 = 0.66
# 轴线
shape(s, MSO_SHAPE.RECTANGLE, x0+0.1, 2.05, len(stages)*colw+(len(stages)-1)*gap-0.2, 0.03, fill="E1D8C6")
for i, (era, fn, form, c) in enumerate(stages):
    cx = x0 + i*(colw+gap)
    dot = shape(s, MSO_SHAPE.OVAL, cx+colw/2-0.17, 1.9, 0.34, 0.34, fill=c)
    tfn = add_box(s, cx+colw/2-0.17, 1.94, 0.34, 0.26)
    para(tfn, str(i+1), size=11, bold=True, color="FFFFFF", first=True, after=0, align=PP_ALIGN.CENTER)
    tf = add_box(s, cx, 2.5, colw, 4.1)
    para(tf, era, size=15, bold=True, color="2A2620", font=FONT_SERIF, first=True, after=3, align=PP_ALIGN.CENTER)
    para(tf, "功能", size=9, color="867E72", after=1, align=PP_ALIGN.CENTER)
    para(tf, fn, size=12, bold=True, color=c, after=7, align=PP_ALIGN.CENTER)
    para(tf, "形式", size=9, color="867E72", after=1, align=PP_ALIGN.CENTER)
    para(tf, form, size=12, color="2A2620", after=0, align=PP_ALIGN.CENTER)
tf = add_box(s, 0.7, 6.35, SW-1.4, 0.6)
para(tf, "主线：功能重心沿「生存巫术 → 礼制象征 → 实用理性 → 大一统宗教 → 审美商品 → 工业现代」迁移，形式随之流转。",
     size=13, bold=True, color="9A3B2E", first=True, after=0)
footer(s, 4)

# ---- 章节幻灯片（S5-S14）----
CH = [
    ("第一章 · 艺术设计的萌芽", "原始社会", "生存 + 巫术",
     "打制石器随「砍砸·切割」成形；彩陶纹饰为「祭祀通神」而绘。",
     "实用与象征两条功能线，在起源处即并行萌发。", "swirl", "9A3B2E"),
    ("第二章 · 中华文明形成期", "新石器时代", "定居生活 + 礼制初现",
     "陶器器形随汲水·炊煮·储藏定型；玉器由工具实用转向礼制。",
     "玉器是史上首次「形式脱离实用、单向追随象征」。", "wave", "274753"),
    ("第三章 · 青铜时代", "夏·商·西周", "礼制象征主导",
     "鼎、尊、簋的器形与饕餮纹，由「列鼎等级制度」决定，而非实用。",
     "功能不单决定形式，还决定技术投向——青铜被礼制「征用」。", "thunder", "9A3B2E"),
    ("第四章 · 早期铁器时代", "春秋·战国", "实用 + 技术理性",
     "铁器随农战实用需求迅速普及；《考工记》立「天时·地气·材美·工巧」。",
     "《考工记》是「形式追随功能」的第一次理论化。", "diamond", "C2953F"),
    ("第五章 · 秦汉时期", "秦·汉", "大一统 + 天人感应",
     "陵墓空间、兵马俑、服饰章纹，随「事死如事生」与等级秩序而设。",
     "形式转为象征符码，统一而成熟。", "cloud", "274753"),
    ("第六章 · 魏晋南北朝", "六朝", "精神解脱 + 宗教",
     "玄学趋自然主义审美；佛教兴石窟造像；高坐家具随「垂足而坐」兴起。",
     "生活方式改变，形式立即追随。", "lotus", "9A3B2E"),
    ("第七章 · 隋唐时期", "隋·唐", "开放融合 + 帝国气象",
     "金银器、都城、服饰随中外交融与秩序需求，走向雄浑开放。",
     "形式承载「交流与秩序」的社会功能。", "roundel", "C2953F"),
    ("第八章 · 五代两宋辽金西夏", "五代·两宋·辽金西夏", "文人审美 + 市民生活",
     "宋瓷随「尚意·内敛」审美臻于简约纯净；雕版印刷随文化传播而行。",
     "形式追随审美功能，宋瓷达极简高峰。", "crackle", "274753"),
    ("第九章 · 元明清时期", "元·明·清", "商品化 + 宫廷",
     "明式家具结构即功能、比例合人体；青花瓷与版画随市场流通量产。",
     "明式家具是全书最纯粹的「形式追随功能」范本。", "searock", "9A3B2E"),
    ("第十章 · 近现代设计", "近现代", "工业化 + 现代生活",
     "形式随机器生产与大众需求而变，西方功能主义正式传入。",
     "完成「传统造物 → 现代设计」的转型。", "geo", "C2953F"),
]
for idx, (title, era, fn, form, judge, motif, accent) in enumerate(CH):
    s = new_slide(); bg(s, "F7F2E9")
    shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, 0.6, 0.55, 0.09, 1.0, fill=accent)
    tf = add_box(s, 0.9, 0.52, 10.4, 1.15)
    para(tf, title, size=24, bold=True, color="2A2620", font=FONT_SERIF, first=True, after=2)
    para(tf, era, size=12, color="867E72", after=0)
    # 左列：功能 / 形式 / 判断
    tf = add_box(s, 0.85, 1.9, 7.3, 4.9)
    para(tf, "功能重心", size=11, bold=True, color=accent, first=True, after=2)
    para(tf, fn, size=20, bold=True, color="2A2620", font=FONT_SERIF, after=10)
    para(tf, "形式如何追随", size=11, bold=True, color=accent, after=2)
    para(tf, form, size=14.5, color="2A2620", after=10, leading=1.25)
    para(tf, "关键判断", size=11, bold=True, color=accent, after=2)
    para(tf, judge, size=13, color="274753", after=0, leading=1.25, italic=True)
    # 右侧纹样图
    ip = os.path.join(ASSET, f"{motif}_800x800.png")
    shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, 8.55, 1.85, 4.15, 4.15, fill="FFFDF8", line="E1D8C6", line_w=1.0)
    s.shapes.add_picture(ip, Inches(8.72), Inches(2.02), height=Inches(3.81))
    footer(s, 5 + idx)

# ---- 知识拓展（据维基百科）----
KN = [
    ("01", "人面鱼纹彩陶盆", "仰韶文化彩陶代表作（约6000年前），1955年陕西西安半坡遗址出土，高16.5厘米、口径39.8厘米，现藏中国国家博物馆。", "仰韶文化", "9A3B2E"),
    ("02", "良渚玉琮", "良渚文化（约5300—4000年前，长江下游）玉琮“内圆外方”，体现“天圆地方”宇宙观，是通天地、祭神礼地的礼器。", "良渚文化", "274753"),
    ("03", "后母戊鼎", "商代晚期青铜方鼎，1939年河南安阳出土，通高133厘米、重832.84千克，中国最重商代青铜礼器，饰兽面纹与夔龙纹。", "后母戊鼎", "9A3B2E"),
    ("04", "考工记", "先秦工艺专著，提出“天有时、地有气、材有美、工有巧”，是“形式追随功能”的第一次系统理论化。", "考工记", "C2953F"),
    ("05", "秦始皇兵马俑", "1974年陕西临潼发现，属秦始皇陵陪葬坑，以千件等身陶俑再现统一帝国军阵与“事死如事生”观念。", "秦始皇兵马俑", "274753"),
    ("06", "云冈石窟", "位于今山西大同，始凿于北魏文成帝和平年间（460年起），历时60余年，现存主要洞窟45个，为北魏皇家石窟。", "云冈石窟", "9A3B2E"),
    ("07", "唐三彩", "唐代低温铅釉陶器，釉彩以黄、绿、白（褐）为主，多用于随葬明器，是盛唐气象与中外交流的物化。", "唐三彩", "C2953F"),
    ("08", "宋代开片瓷（哥窑·汝窑）", "宋代以自然“开片”为美：哥窑“金丝铁线、紫口铁足”，汝窑冰裂纹开片——尚意审美催生的釉色极简。", "哥窑", "274753"),
    ("09", "明式圈椅", "圈椅为明式家具代表，造型简练、以线为主，腿足“外圆内方”，通体榫卯不施钉，尺度合人体曲线。", "明式家具", "9A3B2E"),
    ("10", "形式追随功能", "美国建筑师路易斯·沙利文（1856—1924）1896年于《高层办公建筑的艺术考虑》提出，成现代功能主义口号。", "路易斯·沙利文", "C2953F"),
]

def knowledge_slide(n, title, sub, entries):
    s = new_slide(); bg(s, "F7F2E9")
    shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, 0.6, 0.55, 0.09, 1.0, fill="9A3B2E")
    tf = add_box(s, 0.9, 0.52, 11.6, 1.15)
    para(tf, title, size=26, bold=True, color="2A2620", font=FONT_SERIF, first=True, after=2)
    para(tf, sub, size=12, color="867E72", after=0)
    y = 1.70
    for num, name, fact, wiki, accent in entries:
        shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, 0.7, y, 11.93, 0.95, fill="FFFDF8", line="E1D8C6", line_w=1.0)
        shape(s, MSO_SHAPE.RECTANGLE, 0.7, y, 0.08, 0.95, fill=accent)
        tfn = add_box(s, 0.95, y+0.12, 0.6, 0.72)
        para(tfn, num, size=20, bold=True, color=accent, font=FONT_SERIF, first=True, after=0)
        tft = add_box(s, 1.7, y+0.07, 3.3, 0.82)
        para(tft, name, size=13, bold=True, color="2A2620", font=FONT_SERIF, first=True, after=3)
        para(tft, "出处：维基百科「" + wiki + "」", size=8.5, color=accent, after=0)
        tfb = add_box(s, 5.2, y+0.1, 7.15, 0.76)
        para(tfb, fact, size=10.5, color="2A2620", after=0, leading=1.15)
        y += 1.02
    footer(s, n)

knowledge_slide(15, "知识拓展 · 第一～五章", "每章一个关键知识点，附维基百科出处", KN[:5])
knowledge_slide(16, "知识拓展 · 第六～十章", "每章一个关键知识点，附维基百科出处", KN[5:])

# ---- S17 结论 ----
s = new_slide(); bg(s, "F7F2E9")
shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, 0.6, 0.55, 0.09, 1.0, fill="9A3B2E")
tf = add_box(s, 0.9, 0.6, 9.0, 1.0)
para(tf, "贯穿性结论", size=32, bold=True, color="2A2620", font=FONT_SERIF, first=True, after=0)
concl = [
    ("三重驱动", "中国的「形式追随功能」由实用、礼制象征、审美生活三重功能并行驱动。",
     "其中「礼制象征」是中国区别于西方功能主义的根本特色——形式会为「等级与礼制」服务，而非纯粹实用。"),
    ("技术中介", "技术是功能与形式之间的中介变量。",
     "功能提出要求，材料与工艺决定形式的上限，二者共同锁定形式的实现方式。"),
    ("演进主线", "功能重心的迁移，驱动形式一路流转。",
     "生存巫术 → 礼制象征 → 实用理性 → 大一统宗教 → 审美商品 → 工业现代；形式随之从工具纹饰走向现代设计。"),
]
y = 1.75
for i, (t, lead, body) in enumerate(concl):
    numk = ["9A3B2E", "C2953F", "274753"][i]
    shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, 0.7, y, 12.0, 1.55, fill="FFFDF8", line="E1D8C6", line_w=1.0)
    c = shape(s, MSO_SHAPE.OVAL, 0.95, y+0.27, 0.55, 0.55, fill=numk)
    tfn = add_box(s, 0.95, y+0.34, 0.55, 0.42)
    para(tfn, str(i+1), size=16, bold=True, color="FFFFFF", first=True, after=0, align=PP_ALIGN.CENTER)
    tf = add_box(s, 1.75, y+0.14, 10.7, 1.35)
    para(tf, t + "  ·  " + lead, size=15.5, bold=True, color="2A2620", font=FONT_SERIF, first=True, after=3)
    para(tf, body, size=12.5, color="2A2620", after=0, leading=1.2)
    y += 1.68
footer(s, 17)

# ---- S18 参考文献与出处 ----
s = new_slide(); bg(s, "F7F2E9")
shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, 0.6, 0.55, 0.09, 1.0, fill="C2953F")
tf = add_box(s, 0.9, 0.52, 11.6, 1.15)
para(tf, "参考文献与出处", size=32, bold=True, color="2A2620", font=FONT_SERIF, first=True, after=0)
para(tf, "知识点据维基百科中文条目整理，供课堂延伸查阅（点击链接可跳转）", size=12, color="867E72", after=0)
REFS = [
    ("仰韶文化", "https://zh.wikipedia.org/wiki/仰韶文化"),
    ("良渚文化", "https://zh.wikipedia.org/wiki/良渚文化"),
    ("后母戊鼎", "https://zh.wikipedia.org/wiki/后母戊鼎"),
    ("考工记", "https://zh.wikipedia.org/wiki/考工记"),
    ("秦始皇兵马俑", "https://zh.wikipedia.org/wiki/秦始皇兵马俑"),
    ("云冈石窟", "https://zh.wikipedia.org/wiki/云冈石窟"),
    ("唐三彩", "https://zh.wikipedia.org/wiki/唐三彩"),
    ("哥窑", "https://zh.wikipedia.org/wiki/哥窑"),
    ("汝窑", "https://zh.wikipedia.org/wiki/汝窑"),
    ("明式家具", "https://zh.wikipedia.org/wiki/明式家具"),
    ("路易斯·沙利文", "https://zh.wikipedia.org/wiki/路易斯·沙利文"),
]

def add_ref(tf, title, url, first):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.space_after = Pt(2)
    r = p.add_run(); r.text = title
    set_font(r, FONT_SERIF, 10.5, True, "274753")
    p2 = tf.add_paragraph(); p2.space_after = Pt(9)
    r2 = p2.add_run(); r2.text = url
    set_font(r2, FONT_BODY, 8, False, "9A3B2E")
    if url.startswith("http"):
        r2.font.underline = True
        r2.hyperlink.address = url

left = REFS[:6]
right = REFS[6:] + [("夏燕靖《中国艺术设计史》（第3版）", "全书内容主线与章节结构来源")]
for colx, w, entries in [(0.85, 5.7, left), (6.8, 5.85, right)]:
    tf = add_box(s, colx, 1.75, w, 5.0)
    first = True
    for title, u in entries:
        add_ref(tf, title, u, first)
        first = False
footer(s, 18)

# ---- S19 结束页 ----
s = new_slide(); bg(s, "274753")
shape(s, MSO_SHAPE.RECTANGLE, 0, 0, SW, 0.16, fill="C2953F")
shape(s, MSO_SHAPE.RECTANGLE, 0, SH-0.16, SW, 0.16, fill="9A3B2E")
em = os.path.join(ASSET, "emblem_800x800.png")
s.shapes.add_picture(em, Inches(9.4), Inches(3.6), height=Inches(2.2))
tf = add_box(s, 0.9, 2.3, 8.0, 2.6)
para(tf, "谢谢观看", size=46, bold=True, color="F7F2E9", font=FONT_SERIF, first=True, after=10)
para(tf, "形式追随功能，功能因时而变。", size=18, color="C2953F", after=8)
para(tf, "读懂一部造物史，即读懂一部「需求」如何塑造「形态」的历史。", size=14, color="F7F2E9", after=0)

out = os.path.join(BASE, "中国艺术设计史_形式追随功能_课件.pptx")
prs.save(out)
print("SAVED:", out)
print("SLIDES:", len(prs.slides._sldIdLst))