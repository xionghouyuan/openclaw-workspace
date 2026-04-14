#!/usr/bin/env python3
"""Bain风格幻灯片 PNG 预览 — 纯 PIL"""
from PIL import Image, ImageDraw, ImageFont

BAIN_RED   = (0xCC, 0x22, 0x29)
BLACK      = (0x1A, 0x1A, 0x1A)
DARK_GRAY  = (0x4A, 0x4A, 0x4A)
MED_GRAY   = (0x7F, 0x7F, 0x7F)
WHITE      = (0xFF, 0xFF, 0xFF)
LIGHT_GRAY = (0xF2, 0xF2, 0xF2)
BORDER_G   = (0xD9, 0xD9, 0xD9)
CHARCOAL   = (0x3D, 0x3D, 0x3D)

FCJK  = '/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc'
FCJKB = '/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc'
FDEV  = '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'
FDEVB = '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'
FGEORGIA = '/usr/share/fonts/truetype/dejavu/DejaVuSerif.ttf'

def fnt(path, size):
    try: return ImageFont.truetype(path, size)
    except: return ImageFont.load_default()

W, H = 1000, 750

def new_slide():
    img = Image.new('RGB', (W, H), WHITE)
    return img, ImageDraw.Draw(img)

def rect(d, x, y, w, h, fill=None, outline=None, width=1):
    if fill:    d.rectangle([x, y, x+w, y+h], fill=fill)
    if outline: d.rectangle([x, y, x+w, y+h], outline=outline, width=width)

def txt(d, text, x, y, font_path=FCJK, size=13, color=BLACK, bold=False):
    fp = FCJKB if bold else font_path
    d.text((x, y), text, font=fnt(fp, size), fill=color)

# ─── SLIDE 1 — COVER ───────────────────────────────────
def slide1():
    img, d = new_slide()
    rect(d, 0, 0, W, 5, fill=BAIN_RED)

    txt(d, '银浆印刷加热雨刮', 50, 138, size=46, bold=True)
    txt(d, '卧倒区域方案', 50, 205, size=46, bold=True)
    rect(d, 50, 318, 220, 8, fill=BAIN_RED)
    txt(d, '开发费用周期及计划  /  Development Cost Cycle & Plan', 50, 338, size=17, color=DARK_GRAY)

    stats = [('¥5~10万','总开发费用'),('-45天','进度偏差'),('11/15','OTS里程碑'),('3项','待决策')]
    x = 50
    for val, lbl in stats:
        rect(d, x, 428, 210, 130, fill=LIGHT_GRAY)
        d.text((x+12, 444), val, font=fnt(FCJKB, 36), fill=BAIN_RED)
        txt(d, lbl, x+12, 514, size=14, color=DARK_GRAY)
        x += 230

    d.text((50, 648), '2026.04.13', font=fnt(FDEV, 13), fill=MED_GRAY)
    d.text((40,  H-28), '银浆印刷加热雨刮卧倒区域方案', font=fnt(FDEV, 11), fill=MED_GRAY)
    d.text((W-60, H-28), '1 / 5', font=fnt(FDEV, 12), fill=MED_GRAY)
    img.save('/tmp/bain-slide1.png')
    print('slide1 done')

# ─── SLIDE 2 — OVERVIEW ────────────────────────────────
def slide2():
    img, d = new_slide()
    rect(d, 0, 0, W, 5, fill=BAIN_RED)
    d.text((40, 18), '01 / OVERVIEW', font=fnt(FDEV, 11), fill=MED_GRAY)
    txt(d, 'OTS里程碑 11/15，当前进度滞后 45 天，三项决策待确认', 40, 42, size=19, bold=True)
    rect(d, 40, 96, W-80, 2, fill=BORDER_G)

    # Cost table header
    rect(d, 40, 110, 440, 36, fill=BAIN_RED)
    d.text((55, 113), '开发费用 Development Cost', font=fnt(FCJK, 14), fill=WHITE)

    rows = [
        ('模具单侧 Left or Right', '¥20,000', LIGHT_GRAY),
        ('模具双侧 Left + Right',  '¥40,000', WHITE),
        ('试验单侧 Cabin',         '¥30,000', LIGHT_GRAY),
        ('试验双侧 Cabin',         '¥60,000', WHITE),
    ]
    y = 148
    for lbl, val, bg in rows:
        rect(d, 40, y, 440, 44, fill=bg, outline=BORDER_G, width=1)
        txt(d, lbl, 55, y+6, size=13, color=DARK_GRAY)
        clr = BAIN_RED if val in ('¥40,000','¥60,000') else BLACK
        d.text((390, y+4), val, font=fnt(FCJKB, 14), fill=clr)
        y += 46

    d.text((40, 337), '含模具 + 试验费用，详见技术规格', font=fnt(FDEV, 10), fill=MED_GRAY)

    # OTS card
    rect(d, 510, 110, 450, 155, fill=LIGHT_GRAY)
    rect(d, 510, 110, 5, 155, fill=BAIN_RED)
    txt(d, 'OTS 里程碑  /  OTS Milestone', 530, 115, size=13, color=BAIN_RED, bold=True)
    d.text((530, 148), '11/15', font=fnt(FCJKB, 52), fill=BAIN_RED)
    txt(d, '当前滞后 45 天，需紧急压缩周期', 530, 225, size=13, color=DARK_GRAY)

    # Decisions card
    rect(d, 510, 280, 450, 195, fill=WHITE, outline=BORDER_G, width=1)
    rect(d, 510, 280, 5, 195, fill=BAIN_RED)
    txt(d, '待决策事项  /  Decision Items', 530, 285, size=13, color=BAIN_RED, bold=True)
    decisions = [
        '① 尽快正式提供雨刮区域加热配置的技术要求',
        '② 确认该配置对应的具体销售区域和车型',
        '③ 确认是否需要开发左/右舵 Cabin 双侧配置',
    ]
    dy = 322
    for dec in decisions:
        txt(d, dec, 530, dy, size=12, color=DARK_GRAY); dy += 42
    d.text((530, 456), '需尽快确认，避免影响后续开发节点', font=fnt(FDEV, 10), fill=MED_GRAY)

    d.text((40, H-28), '银浆印刷加热雨刮卧倒区域方案 — 开发费用及里程碑', font=fnt(FDEV, 11), fill=MED_GRAY)
    d.text((W-60, H-28), '2 / 5', font=fnt(FDEV, 12), fill=MED_GRAY)
    img.save('/tmp/bain-slide2.png')
    print('slide2 done')

# ─── SLIDE 3 — GANTT ───────────────────────────────────
def slide3():
    img, d = new_slide()
    rect(d, 0, 0, W, 5, fill=BAIN_RED)
    d.text((40, 18), '02 / TIMELINE', font=fnt(FDEV, 11), fill=MED_GRAY)
    txt(d, '开发周期 M10 ~ M10 SOP，当前进度滞后 45 天风险突出', 40, 42, size=19, bold=True)
    rect(d, 40, 96, W-80, 2, fill=BORDER_G)

    phases = [
        ('M10 KO',        BAIN_RED),
        ('M11 B',         (0x7B, 0x3A, 0x9B)),
        ('M1-3 过程开发', (0x8B, 0x5E, 0x3C)),
        ('M4-6 S1',       (0xC0, 0x5C, 0x9C)),
        ('M7-8 S2',       (0xCC, 0x44, 0x4A)),
        ('M9 P1',         (0x1A, 0x7A, 0x3A)),
        ('M10 SOP',       (0x0A, 0x50, 0x30)),
    ]
    px = 40; pw = 134
    for lbl, col in phases:
        rect(d, px, 108, pw-6, 64, fill=col)
        parts = lbl.split(' ')
        d.text((px+4, 112), parts[0], font=fnt(FCJKB, 12), fill=WHITE)
        if len(parts) > 1:
            d.text((px+4, 140), parts[1], font=fnt(FCJK, 11), fill=WHITE)
        px += pw

    months = ['M10','M11','M1','M2','M3','M4','M5','M6','M7','M8','M9','M10']
    mx = 40; mw = 75
    for m in months:
        d.text((mx+18, 177), m, font=fnt(FDEV, 10), fill=MED_GRAY)
        rect(d, mx+mw, 177, 1, 4, fill=BORDER_G)
        mx += mw

    tasks = [
        ('数据设计', 0.20, BAIN_RED,  'M4~M5'),
        ('设变流程', 0.25, BAIN_RED,  'M5~M6'),
        ('商务议价', 0.30, BAIN_RED,  'M6~M7'),
        ('零件开发', 0.45, CHARCOAL, 'M7~M9'),
        ('型式试验', 0.55, CHARCOAL, 'M9~M11'),
        ('OTS认可',  0.72, BAIN_RED,  'M11~P+'),
    ]
    ty = 205
    for name, spct, col, dates in tasks:
        rect(d, 40, ty, 120, 52, fill=LIGHT_GRAY, outline=BORDER_G, width=1)
        txt(d, name, 45, ty+8, size=14)
        bar_l = int(165 + spct * (W - 165 - 40))
        rect(d, bar_l, ty+6, W-40-bar_l, 40, fill=col)
        d.text((bar_l+8, ty+10), dates, font=fnt(FCJK, 11), fill=WHITE)
        ty += 62

    rect(d, 40, H-50, W-80, 2, fill=BORDER_G)
    d.text((40, H-44), '当前滞后 45 天，关键路径为型式试验（50天）+ OTS认可（30天）',
           font=fnt(FCJK, 12), fill=DARK_GRAY)
    d.text((40, H-28), '银浆印刷加热雨刮卧倒区域方案 — 开发周期', font=fnt(FDEV, 11), fill=MED_GRAY)
    d.text((W-60, H-28), '3 / 5', font=fnt(FDEV, 12), fill=MED_GRAY)
    img.save('/tmp/bain-slide3.png')
    print('slide3 done')

# ─── SLIDE 4 — RISK ────────────────────────────────────
def slide4():
    img, d = new_slide()
    rect(d, 0, 0, W, 5, fill=BAIN_RED)
    d.text((40, 18), '03 / RISK & MITIGATION', font=fnt(FDEV, 11), fill=MED_GRAY)
    txt(d, '当前开发周期无法满足 SOP，OTS 里程碑 11/15 滞后 45 天', 40, 42, size=19, bold=True)
    rect(d, 40, 96, W-80, 2, fill=BORDER_G)

    # Risk card
    rect(d, 40, 110, 440, 245, fill=LIGHT_GRAY)
    rect(d, 40, 110, 5, 245, fill=BAIN_RED)
    txt(d, '关键风险  Key Risk', 60, 115, size=13, color=BAIN_RED, bold=True)
    d.text((60, 152), '-45天', font=fnt(FCJKB, 58), fill=BAIN_RED)
    txt(d, '当前进度滞后', 60, 235, size=16, color=DARK_GRAY)
    risks = [
        '• 整体周期无法满足 SOP 要求',
        '• 型式试验（50天）和 OTS 认可（30天）',
        '  为关键路径，无法压缩',
        '• 当前已滞后 45 天，需紧急干预',
    ]
    ry = 268
    for r in risks:
        txt(d, r, 60, ry, size=12, color=DARK_GRAY); ry += 30

    # Measures card
    rect(d, 510, 110, 450, 440, fill=WHITE, outline=BORDER_G, width=1)
    txt(d, '应对措施  Mitigation Measures', 530, 115, size=13, bold=True)
    measures = [
        ('措施 1', '目标压缩 30~45 天',
         '提前启动设计工作，压缩设计周期；提前发起紧急设变流程；推动商务议价；提前排期试验。'),
        ('措施 2', '提前 15 天达到量产',
         '关键型式试验完成后，采用过渡 OTS 认可流程，可提前 15 天达到量产条件。'),
    ]
    my = 155
    for num, title, desc in measures:
        rect(d, 510, my, 5, 75, fill=BAIN_RED)
        txt(d, num, 528, my, size=12, color=BAIN_RED, bold=True)
        txt(d, title, 528, my+28, size=15, bold=True)
        txt(d, desc, 528, my+65, size=12, color=DARK_GRAY)
        my += 145

    rect(d, 40, 558, W-80, 78, fill=LIGHT_GRAY)
    txt(d, '执行要点：提前启动设计、紧急设变、提前排期试验是压缩周期的关键', 55, 570, size=14, color=DARK_GRAY)
    d.text((40, H-28), '银浆印刷加热雨刮卧倒区域方案 — 风险与应对', font=fnt(FDEV, 11), fill=MED_GRAY)
    d.text((W-60, H-28), '4 / 5', font=fnt(FDEV, 12), fill=MED_GRAY)
    img.save('/tmp/bain-slide4.png')
    print('slide4 done')

# ─── SLIDE 5 — DECISIONS ──────────────────────────────
def slide5():
    img, d = new_slide()
    rect(d, 0, 0, W, 5, fill=BAIN_RED)
    d.text((40, 18), '04 / DECISION REQUIRED', font=fnt(FDEV, 11), fill=MED_GRAY)
    txt(d, '三项决策需尽快确认，影响后续开发节点和商务议价', 40, 42, size=19, bold=True)
    rect(d, 40, 96, W-80, 2, fill=BORDER_G)

    items = [
        ('01', '技术要求确认',
         '尽快正式提供雨刮区域加热配置的技术要求，包括功率、温度、位置等具体参数，为设计开发提供依据。',
         BAIN_RED),
        ('02', '销售区域与车型确认',
         '确认该配置对应的具体销售区域和车型，明确目标市场，以便进行针对性的开发和认证。',
         CHARCOAL),
        ('03', '左/右舵 Cabin 双侧配置',
         '确认是否需要开发左舵（Left）和右舵（Right）双侧配置，还是仅需单侧 Cabin 配置。',
         (0x8C, 0x8C, 0x8C)),
    ]
    y = 112
    for num, title, desc, col in items:
        rect(d, 40, y, 65, 155, fill=col)
        d.text((40+8, y+50), num, font=fnt(FCJKB, 26), fill=WHITE)
        rect(d, 110, y, 850, 155, fill=LIGHT_GRAY, outline=BORDER_G, width=1)
        txt(d, title, 125, y+10, size=17, bold=True)
        txt(d, desc, 125, y+55, size=13, color=DARK_GRAY)
        y += 168

    rect(d, 40, 593, W-80, 62, fill=WHITE, outline=BAIN_RED, width=2)
    rect(d, 40, 593, 5, 62, fill=BAIN_RED)
    txt(d, '请确认以上三项决策，以便尽快推进开发工作，避免影响 SOP 节点',
        55, 603, size=14, color=BAIN_RED, bold=True)
    d.text((40, H-28), '银浆印刷加热雨刮卧倒区域方案 — 待决策事项', font=fnt(FDEV, 11), fill=MED_GRAY)
    d.text((W-60, H-28), '5 / 5', font=fnt(FDEV, 12), fill=MED_GRAY)
    img.save('/tmp/bain-slide5.png')
    print('slide5 done')

slide1()
slide2()
slide3()
slide4()
slide5()
print('All done')
