# -*- coding: utf-8 -*-
"""TalkTally MVP v1.0 产品手册海报 · 四联竖版生成脚本"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

# ---------- 设计系统 ----------
BG      = "0C2333"   # 深海海军蓝（主背景，占比 ~70%）
CARD    = "16334C"   # 卡片 tint
CARD2   = "1E4465"   # 气泡/亮一档
LINE    = "2E5372"   # 细线
TEXT    = "FFFFFF"
MUTED   = "93A9BC"
ACCENT  = "F26B21"   # 唯一强调色（~5%）
DARKTXT = "0C2333"   # 白卡上的深色字
MUTED2  = "5A6B7A"   # 白卡上的次级字

EN      = "Arial"
EN_BLK  = "Arial Black"
CN      = "微软雅黑"

W, H, M = 9.0, 12.0, 0.5

prs = Presentation()
prs.slide_width  = Inches(W)
prs.slide_height = Inches(H)
BLANK = prs.slide_layouts[6]


def slide():
    s = prs.slides.add_slide(BLANK)
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = RGBColor.from_string(BG)
    return s


def _style_run(r, t, size, bold=False, color=TEXT, font=EN, ea=CN, spc=None, italic=False):
    r.text = t
    f = r.font
    f.size = Pt(size); f.bold = bold; f.italic = italic
    f.name = font
    f.color.rgb = RGBColor.from_string(color)
    rPr = r._r.get_or_add_rPr()
    if spc:
        rPr.set('spc', str(int(spc * 100)))
    e = rPr.find(qn('a:ea'))
    if e is None:
        e = rPr.makeelement(qn('a:ea'), {}); rPr.append(e)
    e.set('typeface', ea)


def txt(s, x, y, w, h, paras, anchor=MSO_ANCHOR.TOP, wrap=True):
    """paras: [ [run_dict, ...], ... ] 每个内层列表是一段"""
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    for m_ in ('margin_left', 'margin_right', 'margin_top', 'margin_bottom'):
        setattr(tf, m_, 0)
    for i, runs in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        runs = [dict(r) for r in runs]
        opts = {k: runs[0].pop(k) for k in ('align', 'line', 'after', 'before') if k in runs[0]}
        for rd in runs:
            for k in ('align', 'line', 'after', 'before'):
                rd.pop(k, None)
        for rd in runs:
            _style_run(p.add_run(), **rd)
        if 'align'  in opts: p.alignment    = opts['align']
        if 'line'   in opts: p.line_spacing = opts['line']
        if 'after'  in opts: p.space_after  = Pt(opts['after'])
        if 'before' in opts: p.space_before = Pt(opts['before'])
    return tb


def rect(s, x, y, w, h, fill=None, line=None, lw=0.75, shape=MSO_SHAPE.RECTANGLE, radius=None):
    sp = s.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    sp.shadow.inherit = False
    if radius is not None:
        try: sp.adjustments[0] = radius
        except Exception: pass
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid(); sp.fill.fore_color.rgb = RGBColor.from_string(fill)
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = RGBColor.from_string(line); sp.line.width = Pt(lw)
    return sp


def hair(s, x, y, w, color=LINE):
    return rect(s, x, y, w, 9525 / 914400, fill=color)  # 0.75pt


def R(t, size, bold=False, color=TEXT, font=EN, ea=CN, spc=None, **kw):
    d = dict(t=t, size=size, bold=bold, color=color, font=font, ea=ea, spc=spc)
    d.update(kw); return d


# ============================================================
# 第 1 页 · 封面
# ============================================================
s = slide()
txt(s, M, 0.5, 5.2, 0.3, [[R("TALKTALLY · VOICE-FIRST LEDGER", 9, color=MUTED, spc=2.2)]])
txt(s, W - M - 4.0, 0.5, 4.0, 0.3, [[R("MVP V1.0 · 2026", 9, color=MUTED, spc=2.2, align=PP_ALIGN.RIGHT)]])
hair(s, M, 0.95, W - 2 * M)

txt(s, M, 1.2, W - 2 * M, 1.75, [[
    R("LOGGED", 94, bold=True, font=EN_BLK),
    R(".", 94, bold=True, color=ACCENT, font=EN_BLK),
]])
txt(s, M, 2.8, 8.0, 0.55, [[R("语音记账产品手册", 27, bold=True)]])
txt(s, M, 3.33, 8.0, 0.3, [[R("A VOICE-FIRST EXPENSE LEDGER", 10, color=MUTED, spc=2.4)]])
txt(s, M, 3.76, 8.0, 0.4, [[R("从说出口，到心里有数", 15)]])
txt(s, M, 4.14, 8.0, 0.3, [[R("From spoken words to a clear ledger", 10, color=MUTED)]])

# ---- 手机端示意 ----
px, py, pw, ph = 2.85, 4.55, 3.3, 4.55
rect(s, px, py, pw, ph, fill=CARD, line=LINE, lw=1.0, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.10)
# 说话气泡
rect(s, px + 0.2, py + 0.3, pw - 0.4, 0.62, fill=CARD2, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5)
rect(s, px + 0.33, py + 0.52, 0.1, 0.1, fill=ACCENT, shape=MSO_SHAPE.OVAL)
txt(s, px + 0.55, py + 0.42, pw - 0.95, 0.4, [[R("买水 5 块，支付宝", 11.5, bold=True)]])
# 声波
base = py + 1.62
for i, hh in enumerate([0.14, 0.26, 0.38, 0.2, 0.46, 0.3, 0.5, 0.22, 0.36, 0.16]):
    rect(s, px + 0.62 + i * 0.21, base - hh, 0.045, hh,
         fill=TEXT if i != 4 else ACCENT)
# 按住说话
rect(s, px + 0.55, py + 1.85, pw - 1.1, 0.52, fill=ACCENT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5)
txt(s, px + 0.55, py + 1.97, pw - 1.1, 0.3, [[R("按住说话", 12, bold=True, align=PP_ALIGN.CENTER)]])
# 结果卡片
cx, cy, cw, ch = px + 0.23, py + 2.55, pw - 0.46, 1.72
rect(s, cx, cy, cw, ch, fill=TEXT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.09)
txt(s, cx + 0.2, cy + 0.16, cw - 0.4, 0.35, [[
    R("便利店", 14, bold=True, color=DARKTXT, ea=CN),
    R("   ¥5.00", 14, bold=True, color=ACCENT),
]])
txt(s, cx + 0.2, cy + 0.56, cw - 0.4, 0.3, [[R("餐饮 · 支付宝 · 今天", 9.5, color=MUTED2)]])
rect(s, cx + 0.2, cy + 0.92, 0.86, 0.28, fill=ACCENT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5)
txt(s, cx + 0.2, cy + 0.965, 0.86, 0.22, [[R("云端解析", 8.5, bold=True, align=PP_ALIGN.CENTER)]])
txt(s, cx + 0.2, cy + 1.28, cw - 0.4, 0.3, [[R("仅发送了这句话，音频留在本机", 8.5, color=MUTED2)]])

# ---- 标语行 ----
txt(s, M, 9.35, W - 2 * M, 0.35, [[
    R("SAY IT", 12, bold=True, spc=2),
    R("  →  ", 12, color=ACCENT),
    R("LOGGED", 12, bold=True, spc=2),
]])
txt(s, M, 9.82, W - 2 * M, 0.35, [[R("一句话入账 × 账本在本地 × 每周一句建议", 12, bold=True)]])
txt(s, M, 10.18, W - 2 * M, 0.3,  [[R("One-sentence entry × On-device ledger × Weekly nudge", 9.5, color=MUTED, spc=0.6)]])
hair(s, M, 10.9, W - 2 * M)
txt(s, M, 11.08, 5.0, 0.3, [[
    R("TalkTally", 12, bold=True, font=EN_BLK),
    R(".", 12, bold=True, color=ACCENT, font=EN_BLK),
]])
txt(s, W - M - 4.0, 11.12, 4.0, 0.3, [[R("PRODUCT MANUAL · V1.0", 9, color=MUTED, spc=2, align=PP_ALIGN.RIGHT)]])

# ============================================================
# 第 2 页 · 内页一：工作原理
# ============================================================
s = slide()
txt(s, M, 0.5, 5.6, 0.4, [[R("从说出口到心里有数", 16, bold=True)]])
txt(s, M, 0.94, 5.6, 0.3, [[R("FROM SPOKEN TO LOGGED", 9, color=MUTED, spc=2.2)]])
txt(s, W - M - 4.0, 0.52, 4.0, 0.4, [[
    R("TALKTALLY", 14, bold=True, font=EN_BLK),
    R(".", 14, bold=True, color=ACCENT, font=EN_BLK),
    R("   MVP", 9, color=MUTED, spc=1.5),
]])
hair(s, M, 1.38, W - 2 * M)
txt(s, M, 1.52, W - 2 * M, 0.3, [[
    R("语音引擎 × 本地账本 × 数据最小化", 9.5, bold=True, spc=0.8),
    R("   VOICE ENGINE · ON-DEVICE VAULT · MINIMAL DATA", 9.5, color=MUTED, spc=0.8),
]])

# 左栏
txt(s, M, 1.98, 3.9, 0.5, [[R("01", 26, bold=True, color=ACCENT, font=EN_BLK)]])
txt(s, M, 2.52, 3.9, 0.5, [[R("两层解析，一句话入账", 17, bold=True)]])
txt(s, M, 3.12, 3.9, 3.0, [[
    R("本地规则先判：覆盖「星巴克 38」「打车 25」这类高频句式，零延迟、可复现、离线可用。规则拿不准的长尾句才联网解析，且只发送当前这一句——识别在本机完成，音频从不出设备。", 10.5, color=MUTED, line=1.5, after=8),
    R("账本只写入本机：不建账号、不上传，随时可完整导出。", 10.5, color=MUTED, line=1.5),
]])

# 右栏两张卡
for i, (num, cn_t, en_t, body) in enumerate([
    ("01", "说一句", "SPEAK", "按住说话，端侧识别成文本；三秒得到一张结构化卡片：金额、分类、账户、时间。"),
    ("02", "存本地", "KEEP",  "账本写入本机数据库，可选本机加密；一键导出 CSV / Excel / JSON，永久免费。"),
]):
    cy0 = 1.98 + i * 2.12
    rect(s, 4.7, cy0, 3.8, 1.92, fill=CARD, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.07)
    txt(s, 4.95, cy0 + 0.2, 3.3, 0.35, [[
        R(num + "  ", 15, bold=True, color=ACCENT, font=EN_BLK),
        R(cn_t + "  ", 13, bold=True),
        R(en_t, 9, color=MUTED, spc=1.5),
    ]])
    txt(s, 4.95, cy0 + 0.66, 3.3, 1.1, [[R(body, 10, color=MUTED, line=1.45)]])

hair(s, M, 6.95, W - 2 * M)
txt(s, M, 7.12, 6.5, 0.4, [[R("从开口到复盘，五步走完", 15, bold=True)]])
txt(s, M, 7.54, 6.5, 0.3, [[R("FIVE STEPS, UTTERANCE TO REVIEW", 9, color=MUTED, spc=2)]])

steps = [
    ("01", "唤醒", "WAKE",    "按住小组件直接说"),
    ("02", "说出", "SPEAK",   "音频从不出设备"),
    ("03", "解析", "PARSE",   "规则先判，长尾联网"),
    ("04", "确认", "CONFIRM", "卡片可点改，可语音改"),
    ("05", "复盘", "REVIEW",  "每周一句周报建议"),
]
for i, (num, cn_t, en_t, d) in enumerate(steps):
    x = M + i * 1.6
    if i:
        rect(s, x - 0.06, 8.0, 9525 / 914400, 1.75, fill=LINE)
    txt(s, x, 8.0, 1.45, 0.4, [[R(num, 20, bold=True, color=ACCENT, font=EN_BLK)]])
    txt(s, x, 8.48, 1.45, 0.35, [[
        R(cn_t + "  ", 12, bold=True), R(en_t, 8, color=MUTED, spc=0.8)]])
    txt(s, x, 8.88, 1.42, 0.85, [[R(d, 9, color=MUTED, line=1.4)]])

hair(s, M, 10.05, W - 2 * M)
txt(s, M, 10.22, W - 2 * M, 0.3, [[
    R("每笔入账都标注解析路径", 9.5, bold=True),
    R("——本地规则或云端兜底，流水明细里随时可查。", 9.5, color=MUTED),
]])

hair(s, M, 11.0, W - 2 * M)
txt(s, M, 11.14, 5.0, 0.3, [[R("INSIDE · HOW IT WORKS", 8.5, color=MUTED, spc=2)]])
txt(s, W - M - 4.0, 11.14, 4.0, 0.3, [[R("01 / 04", 8.5, color=MUTED, spc=2, align=PP_ALIGN.RIGHT)]])

# ============================================================
# 第 3 页 · 内页二：场景与隐私
# ============================================================
s = slide()
txt(s, M, 0.5, 5.6, 0.4, [[R("语音记账，真实场景", 16, bold=True)]])
txt(s, M, 0.94, 5.6, 0.3, [[R("REAL-WORLD USE CASES", 9, color=MUTED, spc=2.2)]])
txt(s, W - M - 4.0, 0.52, 4.0, 0.4, [[
    R("TALKTALLY", 14, bold=True, font=EN_BLK),
    R(".", 14, bold=True, color=ACCENT, font=EN_BLK),
    R("   MVP", 9, color=MUTED, spc=1.5),
]])
hair(s, M, 1.38, W - 2 * M)

txt(s, M, 1.6, 8.0, 0.45, [[R("记账的敌人不是懒，是每一步都要点五下。", 15, bold=True)]])
txt(s, M, 2.04, 8.0, 0.3, [[R("The enemy of bookkeeping isn't laziness — it's five taps per transaction.", 9.5, color=MUTED)]])

cards = [
    ("01", "通勤白领", "DAILY COMMUTE", "", "外卖、咖啡、便利店，一天十几笔；按住说一句，金额、分类、账户都已归好。"),
    ("02", "精算学生", "TIGHT BUDGET",  "", "首月预算按周排好；每周一句周报，说清超支多少、花在了哪。"),
    ("03", "情侣合账", "SHARED COSTS", "V0.2", "各说各的，月底合并对账；两笔疑似同一笔，自动提示确认。"),
    ("04", "复杂交易", "COMPLEX ENTRIES", "V0.2", "代付、AA、退款，一句话说清两笔；分类与金额分别落账。"),
]
for i, (num, cn_t, en_t, tag, body) in enumerate(cards):
    x = M + (i % 2) * 4.1
    y = 2.45 + (i // 2) * 1.85
    rect(s, x, y, 3.9, 1.65, fill=CARD, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.07)
    title_runs = [
        R(num + "  ", 14, bold=True, color=ACCENT, font=EN_BLK),
        R(cn_t + "  ", 12.5, bold=True),
        R(en_t, 8, color=MUTED, spc=0.6),
    ]
    if tag:
        title_runs.append(R("  · " + tag, 8, bold=True, color=ACCENT))
    txt(s, x + 0.24, y + 0.14, 3.5, 0.5, [title_runs])
    txt(s, x + 0.24, y + 0.64, 3.45, 0.95, [[R(body, 9.5, color=MUTED, line=1.45)]])

hair(s, M, 6.3, W - 2 * M)
txt(s, M, 6.46, 6.5, 0.4, [[R("数据最小化三原则", 15, bold=True)]])
txt(s, M, 6.88, 6.5, 0.3, [[R("MINIMAL-DATA RULES", 9, color=MUTED, spc=2)]])

rules = [
    ("只发当句", "联网时仅发送当前这句话，不带账号、位置与历史流水"),
    ("零留存",   "服务端不留存、不训练；原始音频从不出设备"),
    ("可关闭",   "一键关闭云端解析，本地规则照样记完每一笔"),
]
for i, (lead, body) in enumerate(rules):
    y = 7.3 + i * 0.62
    rect(s, M, y + 0.09, 0.11, 0.11, fill=ACCENT)
    txt(s, M + 0.28, y, 7.7, 0.5, [[
        R(lead + "　", 11.5, bold=True),
        R("—  " + body, 10, color=MUTED),
    ]])

hair(s, M, 9.4, W - 2 * M)
rect(s, M, 9.66, 0.55, 0.07, fill=ACCENT)
txt(s, M, 9.84, 8.0, 0.5, [[R("说出去的话，收进本地的账。", 18, bold=True)]])
txt(s, M, 10.32, 8.0, 0.3, [[R("Spoken once, kept on device.", 10, color=MUTED)]])

hair(s, M, 11.0, W - 2 * M)
txt(s, M, 11.14, 5.0, 0.3, [[R("INSIDE · USE CASES", 8.5, color=MUTED, spc=2)]])
txt(s, W - M - 4.0, 11.14, 4.0, 0.3, [[R("02 / 04", 8.5, color=MUTED, spc=2, align=PP_ALIGN.RIGHT)]])

# ============================================================
# 第 4 页 · 封底
# ============================================================
s = slide()
txt(s, M, 0.5, 5.0, 0.4, [[
    R("TALKTALLY", 15, bold=True, font=EN_BLK),
    R(".", 15, bold=True, color=ACCENT, font=EN_BLK),
]])
txt(s, W - M - 4.0, 0.56, 4.0, 0.3, [[R("MVP V1.0 · 2026", 9.5, color=MUTED, spc=2.2, align=PP_ALIGN.RIGHT)]])
hair(s, M, 1.05, W - 2 * M)

txt(s, M, 2.3, 8.0, 1.6, [[
    R("账可以慢慢理，", 33, bold=True, line=1.28),
    R("话要当时就说。", 33, bold=True, line=1.28),
]])
txt(s, M, 4.05, 8.0, 0.35, [[R("Log it the moment you say it.", 12, color=MUTED, spc=0.5)]])

# 二维码占位
rect(s, M, 5.0, 0.55, 0.07, fill=ACCENT)
qx, qy, qs = M, 5.3, 1.75
rect(s, qx, qy, qs, qs, fill=TEXT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06)
mod = qs / 21.0
for r in range(21):
    for c in range(21):
        if (r < 7 and c < 7) or (r < 7 and c > 13) or (r > 13 and c < 7):
            continue
        if (r * 7 + c * 13 + (r * c) % 5) % 3 == 0:
            rect(s, qx + c * mod, qy + r * mod, mod * 0.92, mod * 0.92, fill=DARKTXT)
for (fr, fc) in [(0, 0), (0, 14), (14, 0)]:
    rect(s, qx + fc * mod + 0.5 * mod, qy + fr * mod + 0.5 * mod, 6 * mod, 6 * mod, fill=DARKTXT)
    rect(s, qx + fc * mod + 1.5 * mod, qy + fr * mod + 1.5 * mod, 4 * mod, 4 * mod, fill=TEXT)
    rect(s, qx + fc * mod + 2.5 * mod, qy + fr * mod + 2.5 * mod, 2 * mod, 2 * mod, fill=DARKTXT)
txt(s, qx, qy + qs + 0.12, qs, 0.25, [[R("YOUR QR", 8.5, color=MUTED, spc=2, align=PP_ALIGN.CENTER)]])

txt(s, 2.75, 5.6, 5.75, 0.45, [[R("扫码生成你的第一条语音记账", 16, bold=True)]])
txt(s, 2.75, 6.07, 5.75, 0.3, [[R("SCAN TO LOG YOUR FIRST LINE", 9.5, color=MUTED, spc=1.6)]])
txt(s, 2.75, 6.43, 5.75, 0.3, [[R("v1.0 已包含：语音入账 · 本地账本 · 周报建议", 10, color=MUTED)]])

hair(s, M, 8.1, W - 2 * M)
txt(s, M, 8.28, 8.0, 0.4, [[R("一句话入账 × 账本在本地 × 每周一句建议", 13, bold=True)]])
txt(s, M, 8.7, 8.0, 0.3, [[R("One-sentence entry × On-device ledger × Weekly nudge", 9.5, color=MUTED, spc=0.6)]])

txt(s, M, 9.4, W - 2 * M, 0.35, [[R("SAY IT ONCE · LOGGED FOR GOOD", 10, color=MUTED, spc=3, align=PP_ALIGN.CENTER)]])
txt(s, M, 9.92, W - 2 * M, 0.3, [[R("内容对应 TalkTally v1.0 产品文档 · docs/PRD.md", 9, color=MUTED, align=PP_ALIGN.CENTER)]])

hair(s, M, 11.0, W - 2 * M)
txt(s, M, 11.14, 6.0, 0.3, [[R("[YOUR WEBSITE]　　[YOUR EMAIL]", 9.5, color=MUTED, spc=1)]])
txt(s, W - M - 4.0, 11.14, 4.0, 0.3, [[R("BACK · 04 / 04", 8.5, color=MUTED, spc=2, align=PP_ALIGN.RIGHT)]])

prs.core_properties.title = "TalkTally MVP v1.0 产品手册海报"
prs.core_properties.author = "TalkTally"
OUT = r"D:\zcodeproject\TalkTally\TalkTally-MVP-产品手册海报.pptx"
prs.save(OUT)
print("saved:", OUT)
