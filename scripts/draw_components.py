#!/usr/bin/env python3
"""Linear风格组件 — 单独渲染 PNG，透明背景"""
from PIL import Image, ImageDraw, ImageFont
import os

# ── Palette ──
SURFACE = (255,255,255); PANEL = (243,244,245); BORDER = (230,230,230)
TEXT1   = (26,26,30);      TEXT2  = (60,60,67);  TEXT3  = (98,102,109); TEXT4 = (138,143,152)
BRAND   = (94,106,210);    ACCENT = (113,112,255); GREEN  = (39,166,68)
EMERALD = (16,185,129);   AMBER  = (245,166,35);   RED   = (229,72,77)
BROWN   = (139,94,60);     PINK   = (212,96,154);   PURPLE = (124,58,237)
SOPG    = (26,122,58)
LB = (227,242,253); LG = (232,245,233); LR = (255,235,238); LA = (255,243,224)

# ── Fonts ──
FCJK = '/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc'
FCJKB= '/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc'
FS   = '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'
FSB  = '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'
FM   = '/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf'

def cjk(sz, bold=False):
    return ImageFont.truetype(FCJKB if bold else FCJK, sz)
def sn(sz, bold=False):
    return ImageFont.truetype(FSB if bold else FS, sz)
def mn(sz):
    return ImageFont.truetype(FM, sz)

def rr(draw, x1,y1,x2,y2, r, fc, oc=None):
    draw.rounded_rectangle((x1,y1,x2,y2), r, fill=fc, outline=oc)

def rect(draw, x1,y1,x2,y2, fc):
    draw.rectangle((x1,y1,x2,y2), fill=fc)

OUT = '/home/xionghouyuan2/components'
os.makedirs(OUT, exist_ok=True)

def save(img, name):
    p = f'{OUT}/{name}.png'
    img.save(p, 'PNG')
    sz = img.size
    print(f'Saved {name} ({sz[0]}x{sz[1]})')

# ════════════════════════════════════════
# STAT CARD
# ════════════════════════════════════════
def stat_card(num, label, sub, col):
    W,H = 200,120; img = Image.new('RGBA',(W,H),(0,0,0,0)); d = ImageDraw.Draw(img)
    rr(d, 0,0,W-1,H-1, 12, SURFACE, BORDER)
    rr(d, 0,0,W,6, 0, col)           # top accent
    d.text((16,18), num, font=cjk(28,True), fill=col)
    d.text((16,58), label, font=cjk(11), fill=TEXT3)
    d.text((16,78), sub, font=cjk(10), fill=TEXT4)
    save(img, f'stat-{num}')

stat_card('¥5~10万','总开发费用','含模具+试验',BRAND)
stat_card('6','项目任务数','本周计划',GREEN)
stat_card('-45天','进度偏差','当前滞后',RED)
stat_card('3项','待决策事项','需尽快确认',AMBER)

# ════════════════════════════════════════
# RISK CARD
# ════════════════════════════════════════
def risk_card():
    W,H = 480,140; img = Image.new('RGBA',(W,H),(0,0,0,0)); d = ImageDraw.Draw(img)
    rr(d, 0,0,W-1,H-1, 8, LR, BORDER)
    rect(d, 0,0,5,H, RED)
    d.text((16,12), '⚠  关键风险 Key Risk', font=cjk(12,True), fill=RED)
    d.text((16,46), '-45天', font=cjk(30,True), fill=RED)
    d.text((16,94), '当前开发周期无法满足 SOP 要求，OTS 里程碑 11/15 滞后 45 天', font=cjk(11), fill=TEXT3)
    save(img, 'risk-card')

risk_card()

# ════════════════════════════════════════
# DECISION CARD
# ════════════════════════════════════════
def decision_card():
    W,H = 480,140; img = Image.new('RGBA',(W,H),(0,0,0,0)); d = ImageDraw.Draw(img)
    rr(d, 0,0,W-1,H-1, 8, LB, BORDER)
    rect(d, 0,0,5,H, BRAND)
    d.text((16,12), '决策项 Decision Items', font=cjk(12,True), fill=BRAND)
    items = ['① 尽快正式提供雨刮区域加热配置的技术要求','② 确认该配置对应的具体销售区域和车型','③ 确认是否需要开发左/右舵 Cabin 双侧配置']
    for i,item in enumerate(items):
        d.text((16,42+i*28), item, font=cjk(11), fill=TEXT2)
    save(img, 'decision-card')

decision_card()

# ════════════════════════════════════════
# SECTION TITLE BAR
# ════════════════════════════════════════
def sec_title(lbl, title):
    W,H = 560,78; img = Image.new('RGBA',(W,H),(0,0,0,0)); d = ImageDraw.Draw(img)
    d.text((0,4), lbl, font=mn(10), fill=TEXT4)
    d.text((0,24), title, font=cjk(20,True), fill=TEXT1)
    rect(d, 0,72,W,73, BORDER)
    save(img, f'sec-{lbl[:2].lower()}')

sec_title('01 / OVERVIEW','项目概览 Project Overview')
sec_title('02 / COST','开发费用 Development Cost')
sec_title('03 / TIMELINE','开发周期甘特图 Development Gantt')
sec_title('04 / MEASURES','应对措施 Mitigation Measures')

# ════════════════════════════════════════
# COST TABLE
# ════════════════════════════════════════
def cost_table():
    W,H = 480,220; img = Image.new('RGBA',(W,H),(0,0,0,0)); d = ImageDraw.Draw(img)
    rr(d, 0,0,W-1,H-1, 8, SURFACE, BORDER)
    # Mold header
    rect(d, 0,0,W,40, PANEL)
    d.text((16,10), '🛠  模具/夹具费用', font=cjk(13,True), fill=TEXT1)
    rows = [('单侧 Cabin（Left or Right）','¥20,000',TEXT1),('双侧 Cabin（Left+Right）','¥40,000',BRAND)]
    for i,(lbl,val,col) in enumerate(rows):
        y = 50+i*50
        rr(d, 8,y,W-8,y+44, 6, PANEL, BORDER)
        d.text((20,y+8), lbl, font=cjk(12), fill=TEXT3)
        d.text((340,y+8), val, font=sn(16,True), fill=col)
        d.text((20,y+28), 'Left or Right' if i==0 else 'Left + Right', font=cjk(9), fill=TEXT4)
    # Test header
    rect(d, 0,105,W,145, PANEL)
    d.text((16,115), '🔬  试验费用', font=cjk(13,True), fill=TEXT1)
    rows2 = [('单侧 Cabin','¥30,000',TEXT1),('双侧 Cabin','¥60,000',BRAND)]
    for i,(lbl,val,col) in enumerate(rows2):
        y = 155+i*28
        rr(d, 8,y,W-8,y+24, 4, PANEL, BORDER)
        d.text((20,y+4), lbl, font=cjk(11), fill=TEXT3)
        d.text((380,y+4), val, font=sn(13,True), fill=col)
    save(img, 'cost-table')

cost_table()

# ════════════════════════════════════════
# OTS CARD
# ════════════════════════════════════════
def ots_card():
    W,H = 240,220; img = Image.new('RGBA',(W,H),(0,0,0,0)); d = ImageDraw.Draw(img)
    rr(d, 0,0,W-1,H-1, 8, LR, BORDER)
    rect(d, 0,0,5,H, RED)
    d.text((16,14), '⚠  OTS 里程碑', font=cjk(11,True), fill=RED)
    d.text((16,50), '11/15', font=cjk(44,True), fill=RED)
    d.text((16,120), '当前滞后 45 天', font=cjk(12), fill=TEXT3)
    d.text((16,155), '急需压缩周期', font=cjk(12), fill=TEXT3)
    save(img, 'ots-card')

ots_card()

# ════════════════════════════════════════
# PHASE LEGEND BAR
# ════════════════════════════════════════
def phase_bar():
    W,H = 900,62
    phases=[('M10 KO',AMBER),('M11 B',PURPLE),('M1-3 过程开发',BROWN),
            ('M4-6 S1',PINK),('M7-8 S2',RED),('M9 P1',GREEN),('M10 SOP',SOPG)]
    img = Image.new('RGBA',(W,H),(0,0,0,0)); d = ImageDraw.Draw(img)
    pw = W // len(phases)
    for i,(lbl,col) in enumerate(phases):
        x = i*pw
        rr(d, x+1,0,x+pw-2,36, 4, col)
        d.text((x+pw//2-22,10), lbl, font=cjk(9,True), fill=(255,255,255))
    months=['M10','M11','M1','M4','M5','M6','M7','M8','M9','M10']
    mw = W/10
    for i,m in enumerate(months):
        col = BRAND if m=='M4' else TEXT4
        b = m=='M4'
        d.text((int(i*mw+3),40), m, font=cjk(9,b), fill=col)
    save(img, 'phase-bar')

phase_bar()

# ════════════════════════════════════════
# TASK ROW
# ════════════════════════════════════════
def task_row(name, pct, col, dates, lp, wp, comp=False):
    W,H = 900,50; img = Image.new('RGBA',(W,H),(0,0,0,0)); d = ImageDraw.Draw(img)
    d.text((8,12), name, font=cjk(12), fill=TEXT2)           # name
    rr(d, 130,10,W-8,36, 4, PANEL, BORDER)                # track bg
    tw = W-138
    fw = int(tw*wp)
    if fw > 8:
        rr(d, 130,10,130+fw,36, 4, col)
        d.text((136,14), f'{pct}%  {dates}', font=cjk(9,True), fill=(255,255,255))
    save(img, f'task-{name[:2]}')

task_row('数据设计',20,RED,'4/01~5/01',0.30,0.16,True)
task_row('设变流程',20,RED,'5/01~6/01',0.38,0.16,True)
task_row('商务议价',25,RED,'6/01~7/01',0.46,0.13,True)
task_row('零件开发',50,RED,'7/01~9/01',0.54,0.26,True)
task_row('型式试验',50,RED,'9/01~11/01',0.70,0.26,True)
task_row('OTS认可',30,BRAND,'11/01~11/30',0.85,0.13,False)

# ════════════════════════════════════════
# MEASURE CARD
# ════════════════════════════════════════
def measure_card(n, title, desc, col, light):
    W,H = 420,180; img = Image.new('RGBA',(W,H),(0,0,0,0)); d = ImageDraw.Draw(img)
    rr(d, 0,0,W-1,H-1, 8, light, BORDER)
    rr(d, 0,0,W,6, 0, col)               # top accent
    cx,cy,cr = 24,32,17
    d.ellipse((cx-cr,cy-cr,cx+cr,cy+cr), fill=col)
    d.text((cx-5,cy-9), str(n), font=cjk(13,True), fill=(255,255,255))
    d.text((56,24), title, font=cjk(14,True), fill=col)
    words = desc.split('，')
    y = 62
    for w in words:
        d.text((16,y), w+'，', font=cjk(12), fill=TEXT2); y+=22
    save(img, f'measure-{n}')

measure_card(1,'目标压缩 30~45 天',
    '提前启动设计工作，压缩设计周期；提前发起紧急设变流程；推动商务议价；提前排期试验。',AMBER,LA)
measure_card(2,'提前 15 天达到量产',
    '关键型式试验完成后，采用过渡 OTS 认可流程，可提前 15 天达到量产条件。',EMERALD,LG)

# ════════════════════════════════════════
# MINDSET CARD
# ════════════════════════════════════════
def mindset_card():
    W,H = 900,72; img = Image.new('RGBA',(W,H),(0,0,0,0)); d = ImageDraw.Draw(img)
    rr(d, 0,0,W-1,H-1, 8, TEXT1)
    d.text((20,18), '🎯', font=cjk(26), fill=(255,255,255))
    d.text((66,12), '执行要点', font=cjk(9,True), fill=ACCENT)
    d.text((66,30), '提前启动设计、紧急设变、提前排期试验是关键', font=cjk(14), fill=(220,220,235))
    save(img, 'mindset-card')

mindset_card()

# ════════════════════════════════════════
# BOTTOM BAR
# ════════════════════════════════════════
def bottom_bar():
    W,H = 900,62; img = Image.new('RGBA',(W,H),(0,0,0,0)); d = ImageDraw.Draw(img)
    stats=[('¥5~10万','总费用',BRAND),('-45天','进度偏差',RED),('30~45天','目标压缩',EMERALD),('3项','待决策',AMBER)]
    for i,(num,lbl,col) in enumerate(stats):
        x = i*220
        rr(d, x,0,x+200,56, 8, PANEL, BORDER)
        d.text((x+16,6), num, font=cjk(22,True), fill=col)
        d.text((x+16,38), lbl, font=cjk(11), fill=TEXT3)
    save(img, 'bottom-bar')

bottom_bar()

# ════════════════════════════════════════
# COVER HERO
# ════════════════════════════════════════
def cover_hero():
    W,H = 800,300; img = Image.new('RGBA',(W,H),(0,0,0,0)); d = ImageDraw.Draw(img)
    d.text((0,0), '银浆印刷加热雨刮', font=cjk(44,True), fill=TEXT1)
    d.text((0,60), '卧倒区域方案', font=cjk(44,True), fill=TEXT1)
    d.text((0,140), '开发费用周期及计划  /  Development Cost Cycle & Plan', font=cjk(18), fill=TEXT3)
    rect(d, 0,196,200,199, BRAND)
    save(img, 'cover-hero')

cover_hero()

print('\n✅ All components saved to', OUT)
