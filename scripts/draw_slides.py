#!/usr/bin/env python3
"""Linear风格工作计划 — PIL直接绘制"""
from PIL import Image, ImageDraw, ImageFont
import os

# ── Linear Colors ──
BG       = (247, 248, 248)
SURFACE  = (255, 255, 255)
PANEL    = (243, 244, 245)
BORDER   = (230, 230, 230)
TEXT1    = (26, 26, 30)
TEXT2    = (60, 60, 67)
TEXT3    = (98, 102, 109)
TEXT4    = (138, 143, 152)
BRAND    = (94, 106, 210)
ACCENT   = (113, 112, 255)
GREEN    = (39, 166, 68)
EMERALD  = (16, 185, 129)
AMBER    = (245, 166, 35)
RED      = (229, 72, 77)
BROWN    = (139, 94, 60)
PINK     = (212, 96, 154)
PURPLE   = (124, 58, 237)
SOPG     = (26, 122, 58)
LB       = (227, 242, 253)
LG       = (232, 245, 233)
LR       = (255, 235, 238)
LA       = (255, 243, 224)

W, H = 1600, 900

FONT_CJK_REGULAR = '/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc'
FONT_CJK_BOLD = '/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc'
FONT_SANS = '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'
FONT_SANS_BOLD = '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'
FONT_MONO = '/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf'

def font_cjk(size, bold=False):
    return ImageFont.truetype(FONT_CJK_BOLD if bold else FONT_CJK_REGULAR, size)

def font_sans(size, bold=False):
    return ImageFont.truetype(FONT_SANS_BOLD if bold else FONT_SANS, size)

def font_mono(size):
    return ImageFont.truetype(FONT_MONO, size)

def draw_round_rect(draw, xy, r, fill, outline=None, width=1):
    x1,y1,x2,y2 = xy
    draw.rounded_rectangle(xy, r, fill=fill, outline=outline, width=width)

def draw_slide1():
    img = Image.new('RGB', (W, H), BG)
    d = ImageDraw.Draw(img)

    # Header
    draw_round_rect(d, (0, 0, W, 64), 0, fill=SURFACE)
    d.rectangle([(0, 63), (W, 64)], fill=BORDER)
    d.text((36, 22), '银浆印刷加热雨刮卧倒区域方案', font=font_cjk(15, True), fill=TEXT1)
    tag = '📋 开发费用周期及计划'
    tag_w = 150
    draw_round_rect(d, (36+200+10, 20, 36+200+10+tag_w, 44), 20, fill=(227, 242, 253))
    d.text((36+200+20, 22), tag, font=font_cjk(11, False), fill=BRAND)
    d.text((W-120, 24), '2026.04.13', font=font_cjk(12, False), fill=TEXT4)

    # Footer
    draw_round_rect(d, (0, H-52, W, H), 0, fill=SURFACE)
    d.rectangle([(0, H-53), (W, H-52)], fill=BORDER)
    d.text((36, H-32), 'xionghouyuan2 · Work Plan', font=font_cjk(11, False), fill=TEXT4)
    d.text((W-100, H-32), '2026.04.13', font=font_cjk(11, False), fill=TEXT4)

    # Main content centered
    cx, cy = W//2, H//2 - 20

    # Title
    title = '银浆印刷加热雨刮\n卧倒区域方案'
    for i, line in enumerate(title.split('\n')):
        tw = d.textlength(line, font=font_cjk(44, True)) if hasattr(d, 'textlength') else len(line)*44*0.6
        x = cx - tw//2
        d.text((x, cy-130+i*56), line, font=font_cjk(44, True), fill=TEXT1)

    subtitle = '开发费用周期及计划  /  Development Cost Cycle & Plan'
    sw = d.textlength(subtitle, font=font_cjk(18, False)) if hasattr(d, 'textlength') else len(subtitle)*9
    d.text((cx-sw//2, cy+20), subtitle, font=font_cjk(18, False), fill=TEXT3)

    # Divider line
    div_y = cy + 65
    d.rectangle([(cx-80, div_y), (cx+80, div_y+3)], fill=BRAND)

    # Stats row
    stats = [
        ('¥5~10万', '总开发费用', BRAND),
        ('6', '项目任务数', GREEN),
        ('-45天', '进度偏差', RED),
        ('3项', '待决策事项', AMBER),
    ]
    total_sw = 4*200 + 3*16
    start_x = cx - total_sw//2
    for i, (num, lbl, col) in enumerate(stats):
        x = start_x + i*216
        y = cy + 100
        draw_round_rect(d, (x, y, x+200, y+110), 12, fill=SURFACE, outline=BORDER, width=1)
        d.text((x+15, y+15), num, font=font_cjk(32, True), fill=col)
        d.text((x+15, y+60), lbl, font=font_cjk(12, False), fill=TEXT3)
        d.text((x+15, y+80), '含模具+试验' if i==0 else ('本周计划' if i==1 else ('当前滞后' if i==2 else '需尽快确认')), font=font_cjk(10, False), fill=TEXT4)

    return img

def draw_slide2():
    img = Image.new('RGB', (W, H), BG)
    d = ImageDraw.Draw(img)

    # Header
    draw_round_rect(d, (0, 0, W, 64), 0, fill=SURFACE)
    d.rectangle([(0, 63), (W, 64)], fill=BORDER)
    d.text((36, 22), '银浆印刷加热雨刮卧倒区域方案', font=font_cjk(15, True), fill=TEXT1)
    d.text((W-120, 24), '2026.04.13', font=font_cjk(12, False), fill=TEXT4)

    # Footer
    draw_round_rect(d, (0, H-52, W, H), 0, fill=SURFACE)
    d.rectangle([(0, H-53), (W, H-52)], fill=BORDER)
    d.text((36, H-32), 'xionghouyuan2 · Work Plan', font=font_cjk(11, False), fill=TEXT4)
    d.text((W-100, H-32), '2026.04.13', font=font_cjk(11, False), fill=TEXT4)

    # ── Gantt area (left 2/3) ──
    gx, gy = 28, 80
    gw = W//2 - 56

    d.text((gx, gy), '02 / Timeline', font=font_mono(10), fill=TEXT4)
    d.text((gx, gy+20), '开发周期甘特图', font=font_cjk(20, True), fill=TEXT1)

    # Phase legend
    phases = [
        ('M10 KO', AMBER), ('M11 B', PURPLE), ('M1-3 过程', BROWN),
        ('M4-6 S1', PINK), ('M7-8 S2', RED), ('M9 P1', GREEN), ('M10 SOP', SOPG),
    ]
    pw = gw / len(phases)
    py = gy + 58
    ph = 32
    for i, (lbl, col) in enumerate(phases):
        px = gx + pw*i
        draw_round_rect(d, (px, py, px+pw-2, py+ph), 4, fill=col)
        lw = d.textlength(lbl, font=font_cjk(9, True)) if hasattr(d,'textlength') else len(lbl)*5
        d.text((px+(pw-2)/2 - lw/2, py+8), lbl, font=font_cjk(9, True), fill=(255,255,255))

    # Month labels
    my = py + ph + 6
    months = ['M10','M11','M1','M4','M5','M6','M7','M8','M9','M10']
    mw = gw / 10
    for i, m in enumerate(months):
        mx = gx + mw*i
        col = BRAND if m == 'M4' else TEXT4
        bold = m == 'M4'
        d.text((mx, my), m, font=font_cjk(9, bold), fill=col)

    # Grid lines
    for i in range(11):
        vx = gx + mw*i
        d.rectangle([(vx, my+4), (vx+0.5, my+4+300)], fill=BORDER)

    # Task bars
    tasks = [
        ('数据设计',   RED,   0.30, 0.16, '20天  4/01~5/01'),
        ('设变流程',   RED,   0.38, 0.16, '20天  5/01~6/01'),
        ('商务议价',   RED,   0.46, 0.13, '25天  6/01~7/01'),
        ('零件开发',   RED,   0.54, 0.26, '50天  7/01~9/01'),
        ('型式试验',   RED,   0.70, 0.26, '50天  9/01~11/01'),
        ('OTS认可',    BRAND, 0.85, 0.13, '30天  11/01~11/30'),
    ]
    ty = my + 36
    for i, (name, col, left, width, lbl) in enumerate(tasks):
        # Task name
        d.text((8, ty+i*46), name, font=font_cjk(11, False), fill=TEXT2)
        # Bar
        bx = gx + gw*left
        bw = gw*width
        draw_round_rect(d, (bx, ty+i*46+3, bx+bw, ty+i*46+23), 4, fill=col)
        d.text((bx+6, ty+i*46+5), lbl, font=font_cjk(9, True), fill=(255,255,255))
        # Separator
        d.rectangle([(8, ty+i*46+32), (gx+gw, ty+i*46+33)], fill=BORDER)

    # Current line
    cx = gx + mw*3.4
    d.rectangle([(cx, my+4), (cx+1.5, my+4+280)], fill=RED)
    d.text((cx-10, my+290), '当前', font=font_cjk(9, True), fill=RED)

    # ── Right panel ──
    rx = W//2 + 20
    rw = W - rx - 28

    # Risk
    ry = gy
    d.text((rx, ry), '03 / Risk', font=font_mono(10), fill=TEXT4)
    d.text((rx, ry+20), '风险与决策', font=font_cjk(20, True), fill=TEXT1)

    risk_y = ry + 58
    draw_round_rect(d, (rx, risk_y, rx+rw, risk_y+120), 8, fill=LR, outline=BORDER)
    d.rectangle([(rx, risk_y), (rx+5, risk_y+120)], fill=RED)
    d.text((rx+14, risk_y+10), '⚠  进度风险', font=font_cjk(11, True), fill=RED)
    d.text((rx+14, risk_y+38), '-45天', font=font_cjk(30, True), fill=RED)
    d.text((rx+14, risk_y+80), '当前开发周期无法满足 SOP 要求，OTS 里程碑 11/15 滞后 45 天', font=font_cjk(11, False), fill=TEXT3)

    # Decision
    dec_y = risk_y + 135
    draw_round_rect(d, (rx, dec_y, rx+rw, dec_y+130), 8, fill=LB, outline=BORDER)
    d.rectangle([(rx, dec_y), (rx+5, dec_y+130)], fill=BRAND)
    d.text((rx+14, dec_y+10), '决策项 Decision', font=font_cjk(11, True), fill=BRAND)
    decisions = [
        '① 尽快正式提供雨刮区域加热配置的技术要求',
        '② 确认该配置对应的具体销售区域和车型',
        '③ 确认是否需要开发左/右舵 Cabin 双侧配置',
    ]
    for i, dec in enumerate(decisions):
        dy = dec_y + 40 + i*28
        d.text((rx+14, dy), dec, font=font_cjk(11, False), fill=TEXT2)
        if i < len(decisions)-1:
            d.rectangle([(rx+10, dy+22), (rx+rw-10, dy+23)], fill=BORDER)

    # Mindset
    mind_y = dec_y + 145
    draw_round_rect(d, (rx, mind_y, rx+rw, mind_y+70), 8, fill=TEXT1)
    d.text((rx+14, mind_y+18), '🎯', font=font_cjk(22, False))
    d.text((rx+50, mind_y+12), '执行要点', font=font_cjk(9, True), fill=ACCENT)
    d.text((rx+50, mind_y+30), '提前启动设计、紧急设变、提前排期试验是关键', font=font_cjk(12, False), fill=(220,220,235))

    return img

if __name__ == '__main__':
    img1 = draw_slide1()
    img1.save('/home/xionghouyuan2/slide1_direct.png')
    print(f'Slide1: {img1.size}')

    img2 = draw_slide2()
    img2.save('/home/xionghouyuan2/slide2_direct.png')
    print(f'Slide2: {img2.size}')
    print('Done')
