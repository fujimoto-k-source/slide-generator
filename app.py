"""
Slide Generator — 世界最高峰の広告代理店クオリティ
DTP原則を徹底実装: トラッキング・ベースライングリッド・タイプスケール・行間
"""

import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from lxml import etree
import io, re
from datetime import datetime

# ─────────────────────────────────────────
st.set_page_config(page_title="Slide Generator", layout="centered",
                   initial_sidebar_state="collapsed")

st.markdown("""
<style>
#MainMenu, footer, header { visibility: hidden; }
.stApp { background:#FFF; font-family:'Helvetica Neue',Arial,sans-serif; }
.main .block-container { max-width:740px; padding-top:56px; padding-bottom:88px; }

.pg-eyebrow { font-size:10px; font-weight:700; letter-spacing:.2em;
              text-transform:uppercase; color:#CCC; margin-bottom:8px; }
.pg-title   { font-size:26px; font-weight:700; letter-spacing:-.3px;
              color:#111; margin-bottom:6px; line-height:1.2; }
.pg-sub     { font-size:13px; color:#999; margin-bottom:52px; line-height:1.7; }
.step-label { font-size:10px; font-weight:700; letter-spacing:.2em;
              text-transform:uppercase; color:#CCC; margin:40px 0 16px; display:block; }

/* Radio cards */
div[data-testid="stRadio"] label { display:none!important; }
div[data-testid="stRadio"] > div { display:flex; flex-direction:column; gap:8px; }
div[data-testid="stRadio"] > div > label {
    border:1px solid #EAEAEA!important; border-radius:2px!important;
    padding:18px 22px!important; background:#FAFAFA!important;
    cursor:pointer!important; transition:all .15s!important;
}
div[data-testid="stRadio"] > div > label:hover {
    border-color:#111!important; background:#FFF!important; }
div[data-testid="stRadio"] > div > label[data-checked="true"] {
    border-color:#111!important; background:#FFF!important;
    box-shadow:inset 3px 0 0 #111!important; }

/* Inputs */
label, .stTextArea label, .stTextInput label {
    font-size:10px!important; font-weight:700!important;
    letter-spacing:.15em!important; text-transform:uppercase!important; color:#BBBBBB!important; }
.stTextArea textarea {
    background:#F9F9F9!important; border:1px solid #E8E8E8!important;
    border-radius:2px!important; font-size:13.5px!important; color:#111!important;
    font-family:'Helvetica Neue',Arial,sans-serif!important;
    line-height:1.7!important; padding:18px!important; resize:vertical!important; }
.stTextArea textarea:focus { border-color:#111!important; box-shadow:none!important; }
.stTextInput input {
    background:#F9F9F9!important; border:1px solid #E8E8E8!important;
    border-radius:2px!important; font-size:13.5px!important; color:#111!important;
    padding:11px 15px!important; }
.stTextInput input:focus { border-color:#111!important; box-shadow:none!important; }

/* Buttons */
.stButton>button {
    background:#111!important; color:#FFF!important; border:none!important;
    border-radius:2px!important; padding:15px 32px!important;
    font-size:11px!important; font-weight:700!important; letter-spacing:.14em!important;
    text-transform:uppercase!important; width:100%!important; transition:opacity .15s!important; }
.stButton>button:hover { opacity:.65!important; }
.stDownloadButton>button {
    background:#FFF!important; color:#111!important;
    border:1.5px solid #111!important; border-radius:2px!important;
    padding:15px 32px!important; font-size:11px!important; font-weight:700!important;
    letter-spacing:.14em!important; text-transform:uppercase!important;
    width:100%!important; margin-top:10px!important; transition:background .15s!important; }
.stDownloadButton>button:hover { background:#F5F5F5!important; }

/* Swatch */
.swatch { display:inline-block; width:13px; height:13px; border-radius:1px;
          margin-right:8px; vertical-align:middle; border:1px solid rgba(0,0,0,.08); }
.divider { border:none; border-top:1px solid #F0F0F0; margin:44px 0; }
[data-testid="column"] { padding:0 6px!important; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────
#  ① DTPコアユーティリティ
# ─────────────────────────────────────────
W, H = 13.33, 7.5   # slide size in inches

def rgb(h: str) -> RGBColor:
    h = h.lstrip('#')
    return RGBColor(int(h[0:2],16), int(h[2:4],16), int(h[4:6],16))

def run_style(run, font='Calibri', size=None, bold=False, italic=False,
              color: RGBColor=None, tracking: int=0, ea='Meiryo UI'):
    """
    tracking: hundredths-of-a-point (OOXML `spc`)
      +100 = +1pt (wide, for caps labels)
      -80  = -0.8pt (tight, for large display type)
    """
    run.font.name = font
    if size is not None: run.font.size = Pt(size)
    run.font.bold  = bold
    run.font.italic = italic
    if color: run.font.color.rgb = color
    rPr = run._r.get_or_add_rPr()
    # Tracking
    if tracking != 0:
        rPr.set('spc', str(int(tracking)))
    # East-Asian font (Japanese support)
    for ch in list(rPr):
        if ch.tag == qn('a:ea'): rPr.remove(ch)
    ea_el = etree.SubElement(rPr, qn('a:ea'))
    ea_el.set('typeface', ea)

def para_style(p, align=PP_ALIGN.LEFT,
               sp_before: float=0, sp_after: float=0,
               line_pct: float=1.0):
    """
    line_pct: 1.0=single, 1.4=140%...
    """
    p.alignment  = align
    p.space_before = Pt(sp_before)
    p.space_after  = Pt(sp_after)
    if line_pct != 1.0:
        pPr = p._p.get_or_add_pPr()
        for ch in list(pPr):
            if ch.tag == qn('a:lnSpc'): pPr.remove(ch)
        lnSpc  = etree.SubElement(pPr, qn('a:lnSpc'))
        spcPct = etree.SubElement(lnSpc, qn('a:spcPct'))
        spcPct.set('val', str(int(line_pct * 100000)))

def shape_rect(slide, l, t, w, h, fill=None):
    s = slide.shapes.add_shape(1, Inches(l), Inches(t), Inches(w), Inches(h))
    if fill: s.fill.solid(); s.fill.fore_color.rgb = fill
    else:    s.fill.background()
    s.line.fill.background()
    return s

def textbox(slide, text, l, t, w, h,
            font='Calibri', size=16, bold=False, italic=False,
            color: RGBColor=None, align=PP_ALIGN.LEFT,
            tracking: int=0, line_pct: float=1.0,
            sp_before: float=0, wrap=True, ea='Meiryo UI'):
    tb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = wrap
    p  = tf.paragraphs[0]
    para_style(p, align=align, sp_before=sp_before, line_pct=line_pct)
    r = p.add_run(); r.text = text
    run_style(r, font=font, size=size, bold=bold, italic=italic,
              color=color, tracking=tracking, ea=ea)
    return tb

# ─────────────────────────────────────────
#  ② タイプスケール (黄金比ベース)
# ─────────────────────────────────────────
# Display  → 60pt  tracking -100
# H1       → 56pt  tracking -80
# H2       → 32pt  tracking -40
# Body     → 16pt  tracking   0   line-height 1.45
# Sub      → 13pt  tracking   0   line-height 1.38
# Caption  → 10pt  tracking +100
TS = dict(
    display  = dict(size=60, bold=True,  tracking=-100, line_pct=1.05),
    h1       = dict(size=56, bold=True,  tracking=-80,  line_pct=1.05),
    h2       = dict(size=32, bold=True,  tracking=-40,  line_pct=1.1),
    h2_light = dict(size=30, bold=False, tracking=-20,  line_pct=1.1),
    body     = dict(size=16, bold=False, tracking=0,    line_pct=1.45),
    sub      = dict(size=13, bold=False, tracking=0,    line_pct=1.38),
    caption  = dict(size=10, bold=False, tracking=100,  line_pct=1.0),
    label    = dict(size=10, bold=True,  tracking=120,  line_pct=1.0),
    num      = dict(size=10, bold=True,  tracking=80,   line_pct=1.0),
)

# Margin standard
LM = 0.75   # left margin
RM = 0.75   # right margin
CW = W - LM - RM   # content width = 11.83

# ─────────────────────────────────────────
#  ③ コンテンツブロック（全テーマ共通）
# ─────────────────────────────────────────
def content_block(slide, items, l, t, w, h, th):
    tb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = True

    for i, item in enumerate(items):
        p  = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        lv = item.get('level', 0)
        ib = item.get('bold', False)
        txt = item['text']

        if lv == 0 and ib:
            # ── アクセントヘッダー (### 見出し)
            para_style(p, sp_before=18 if i > 0 else 4,
                       sp_after=2, line_pct=1.25)
            r = p.add_run(); r.text = txt
            run_style(r, font='Calibri', size=17, bold=True,
                      color=th['accent'], tracking=-20, ea='Meiryo UI')

        elif lv == 0:
            # ── 主要箇条書き
            para_style(p, sp_before=9 if i > 0 else 0,
                       sp_after=0, line_pct=1.45)
            r = p.add_run()
            r.text = '\u2014\u2002' + txt   # —  + thin space
            run_style(r, font='Calibri', size=16, bold=False,
                      color=th['body'], tracking=0, ea='Meiryo UI')

        else:
            # ── サブ箇条書き (インデント)
            para_style(p, sp_before=4, sp_after=0, line_pct=1.38)
            r = p.add_run()
            r.text = '\u2003\u00B7\u2002' + txt  # em-space + middle-dot + thin-space
            run_style(r, font='Calibri', size=13, bold=False,
                      color=th['sub'], tracking=0, ea='Meiryo UI')

# ─────────────────────────────────────────
#  ④ テーマ設定
# ─────────────────────────────────────────
ACCENT_PRESETS = {
    'ネイビー':   '#0A2463',
    'ゴールド':   '#A67C3A',
    'クリムゾン': '#8C1C2E',
    'エメラルド': '#0B5E4A',
    'スレート':   '#2E3F4F',
}

def get_theme(key: str, accent_hex: str) -> dict:
    a = rgb(accent_hex)
    base = dict(
        accent = a,
        white  = rgb('FFFFFF'),
        ink    = rgb('0D0D0D'),    # near-black (softer than pure black)
        body   = rgb('2C2C2C'),    # body text
        sub    = rgb('717171'),    # sub-bullets / captions
        rules  = rgb('EBEBEB'),    # hairlines
        num_c  = rgb('CCCCCC'),    # slide numbers
        muted  = rgb('9A9A9A'),    # subtitles / dates
    )
    if key == 'MINIMAL':
        base.update(title_bg=rgb('FFFFFF'), hdr_bg=rgb('FFFFFF'), hdr_txt=rgb('0D0D0D'))
    elif key == 'NOIR':
        base.update(title_bg=rgb('0C0C0C'), hdr_bg=rgb('FFFFFF'), hdr_txt=rgb('0D0D0D'))
    elif key == 'BOLD':
        base.update(title_bg=rgb('FFFFFF'), hdr_bg=a,             hdr_txt=rgb('FFFFFF'))
    return base

# ─────────────────────────────────────────
#  ⑤ MINIMAL — 知性のグリッドシステム
#    大型ナンバリング × 水平リズム × 広大な余白
# ─────────────────────────────────────────
def minimal_title(sl, sd, th, author, today):
    # ─ 最上部ルール (アクセント)
    shape_rect(sl, 0, 0, W, 0.045, fill=th['accent'])

    # ─ 著者・日付 (キャプション, wide tracking)
    if author:
        textbox(sl, author.upper(), LM, 0.30, 5.0, 0.35,
                size=10, bold=True, color=th['muted'], tracking=120)
    textbox(sl, today, W-RM-3.0, 0.30, 3.0, 0.35,
            size=10, color=th['muted'], align=PP_ALIGN.RIGHT, tracking=80)

    # ─ タイトル (display, tight tracking)
    textbox(sl, sd['title'], LM, 2.05, CW, 1.75,
            size=60, bold=True, color=th['ink'], tracking=-100, line_pct=1.05)

    # ─ アクセントライン (タイトル直後)
    shape_rect(sl, LM, 3.95, 2.4, 0.022, fill=th['accent'])

    # ─ サブタイトル
    if sd.get('subtitle'):
        textbox(sl, sd['subtitle'], LM, 4.18, CW*0.75, 0.65,
                size=17, color=th['muted'], tracking=0, line_pct=1.5)

    # ─ 最下部ルール
    shape_rect(sl, 0, H-0.045, W, 0.045, fill=th['accent'])


def minimal_content(sl, sd, th, num):
    # ─ スライド番号 (小さく、右上、refined)
    textbox(sl, f'{num:02d}', W-RM-0.8, 0.22, 0.8, 0.38,
            size=10, bold=True, color=th['num_c'],
            align=PP_ALIGN.RIGHT, tracking=80)

    # ─ タイトル (H2 bold)
    textbox(sl, sd.get('title',''), LM, 0.20, CW-1.0, 0.82,
            size=32, bold=True, color=th['ink'], tracking=-40, line_pct=1.1)

    # ─ ヘアラインルール (全幅 / アクセント部分 2.2")
    shape_rect(sl, LM, 1.15, CW, 0.012, fill=th['rules'])
    shape_rect(sl, LM, 1.15, 2.2, 0.012, fill=th['accent'])

    # ─ コンテンツ
    items = sd.get('content', [])
    if items:
        content_block(sl, items, LM, 1.35, CW, H-1.35-0.45, th)

    # ─ 最下部ルール
    shape_rect(sl, 0, H-0.045, W, 0.045, fill=th['accent'])


# ─────────────────────────────────────────
#  ⑥ NOIR — シネマティック・ラグジュアリー
#    漆黒タイトル × 白コンテンツ × 鯯利なアクセント
# ─────────────────────────────────────────
def noir_title(sl, sd, th, author, today):
    # ─ 全面黒背景
    shape_rect(sl, 0, 0, W, H, fill=th['title_bg'])

    # ─ トップ: 極細アクセントライン
    shape_rect(sl, 0, 0, W, 0.045, fill=th['accent'])

    # ─ 右側: 装飾的ヴァーティカルライン (非常に細い、ほぼ見えない程度)
    shape_rect(sl, 12.15, 0.4, 0.018, H-0.8, fill=rgb('1C1C1C'))

    # ─ タイトル (display, 白, tight tracking)
    textbox(sl, sd['title'], LM, 1.70, 10.5, 2.0,
            size=60, bold=True, color=th['white'],
            tracking=-100, line_pct=1.05)

    # ─ アクセントライン (タイトル直後)
    shape_rect(sl, LM, 3.82, 2.0, 0.022, fill=th['accent'])

    # ─ サブタイトル
    if sd.get('subtitle'):
        textbox(sl, sd['subtitle'], LM, 4.06, 9.5, 0.65,
                size=17, color=th['muted'], tracking=0, line_pct=1.5)

    # ─ フッター
    parts = [p for p in [author, today] if p]
    if parts:
        textbox(sl, '  /  '.join(parts), LM, 7.02, CW, 0.28,
                size=10, color=rgb('3A3A3A'),
                align=PP_ALIGN.RIGHT, tracking=80)

    # ─ ボトムライン
    shape_rect(sl, 0, H-0.045, W, 0.045, fill=th['accent'])


def noir_content(sl, sd, th, num):
    # ─ 左: 極細アクセントストリップ
    shape_rect(sl, 0, 0, 0.048, H, fill=th['accent'])

    # ─ スライド番号 (右上, muted)
    textbox(sl, f'{num:02d}', W-RM-0.85, 0.20, 0.85, 0.40,
            size=10, bold=True, color=th['num_c'],
            align=PP_ALIGN.RIGHT, tracking=80)

    # ─ タイトル (H2 bold)
    textbox(sl, sd.get('title',''), LM, 0.20, CW-1.0, 0.82,
            size=32, bold=True, color=th['ink'], tracking=-40, line_pct=1.1)

    # ─ ダブルヘアライン (全幅グレー + 短めアクセント)
    shape_rect(sl, LM, 1.15, CW, 0.012, fill=th['rules'])
    shape_rect(sl, LM, 1.15, 1.8, 0.012, fill=th['accent'])

    # ─ コンテンツ
    items = sd.get('content', [])
    if items:
        content_block(sl, items, LM, 1.35, CW, H-1.35-0.45, th)

    # ─ ボトムライン
    shape_rect(sl, 0, H-0.045, W, 0.045, fill=th['accent'])


# ─────────────────────────────────────────
#  ⑦ BOLD — パネル分割 × ブランド刈印
#    クリエイティブエージェンシーの大胆な構成力
# ─────────────────────────────────────────
PANEL_W = 5.1   # left panel width

def bold_title(sl, sd, th, author, today):
    # ─ 左パネル (アクセントカラー)
    shape_rect(sl, 0, 0, PANEL_W, H, fill=th['accent'])

    # ─ パネル内: 著者 (最上部, 白, wide tracking)
    if author:
        textbox(sl, author.upper(), 0.55, 0.30, PANEL_W-0.65, 0.38,
                size=9, bold=True, color=rgb('FFFFFF'), tracking=120)
    else:
        textbox(sl, today, 0.55, 0.30, PANEL_W-0.65, 0.38,
                size=9, color=rgb('FFFFFF'), tracking=80)

    # ─ パネル内: アクセントライン (水平)
    shape_rect(sl, 0.55, 1.88, 1.8, 0.022, fill=rgb('FFFFFF'))

    # ─ パネル内: タイトル (白, large)
    textbox(sl, sd['title'], 0.55, 2.05, PANEL_W-0.7, 3.5,
            size=44, bold=True, color=th['white'],
            tracking=-80, line_pct=1.08)

    # ─ 右パネル: サブタイトル (中央寄り縦位置)
    if sd.get('subtitle'):
        textbox(sl, sd['subtitle'], PANEL_W+0.65, 2.2, W-PANEL_W-1.2, 2.0,
                size=20, color=th['ink'], tracking=-10, line_pct=1.5)

    # ─ 右パネル: 日付
    textbox(sl, today if author else '', W-RM-2.5, 7.0, 2.5, 0.30,
            size=10, color=th['muted'], align=PP_ALIGN.RIGHT, tracking=80)


def bold_content(sl, sd, th, num):
    # ─ ヘッダーバー (アクセントカラー)
    shape_rect(sl, 0, 0, W, 1.05, fill=th['hdr_bg'])

    # ─ タイトル (ヘッダー内, 白)
    textbox(sl, sd.get('title',''), LM, 0.18, CW-1.2, 0.72,
            size=26, bold=True, color=th['hdr_txt'],
            tracking=-30, line_pct=1.1)

    # ─ スライド番号 (ヘッダー右端, 白, 大きめ)
    textbox(sl, f'{num:02d}', W-RM-0.9, 0.24, 0.9, 0.5,
            size=18, bold=True,
            color=rgb('FFFFFF') if th['hdr_txt'] == th['white'] else th['num_c'],
            align=PP_ALIGN.RIGHT, tracking=60)

    # ─ コンテンツ
    items = sd.get('content', [])
    if items:
        content_block(sl, items, LM, 1.22, CW, H-1.22-0.45, th)

    # ─ ボトムライン
    shape_rect(sl, 0, H-0.045, W, 0.045, fill=th['accent'])


# ─────────────────────────────────────────
#  ⑧ マスタービルダー
# ─────────────────────────────────────────
TITLE_FN   = {'MINIMAL': minimal_title, 'NOIR': noir_title, 'BOLD': bold_title}
CONTENT_FN = {'MINIMAL': minimal_content,'NOIR': noir_content,'BOLD': bold_content}

def build_pptx(slides_data, theme_key, accent_hex, author='') -> io.BytesIO:
    prs = Presentation()
    prs.slide_width  = Inches(W)
    prs.slide_height = Inches(H)
    blank = prs.slide_layouts[6]
    th    = get_theme(theme_key, accent_hex)
    today = datetime.now().strftime('%Y.%m.%d')
    cnum  = 0

    for sd in slides_data:
        sl = prs.slides.add_slide(blank)
        if sd['type'] == 'title':
            TITLE_FN[theme_key](sl, sd, th, author, today)
        else:
            cnum += 1
            CONTENT_FN[theme_key](sl, sd, th, cnum)

    buf = io.BytesIO()
    prs.save(buf); buf.seek(0)
    return buf

# ─────────────────────────────────────────
#  ⑨ テキストパーサー
# ─────────────────────────────────────────
def parse_slides(text: str) -> list[dict]:
    lines = text.strip().split('\n')
    slides, cur = [], [None]

    def push():
        c = cur[0]
        if c and (c.get('title') or c.get('content')): slides.append(c)
        cur[0] = None

    def cs(title=''):
        return {'type':'content','title':title,'content':[]}

    for raw in lines:
        s = raw.strip()
        if re.match(r'^(-{3,}|={3,})$', s):            push(); continue
        if not s:                                        continue
        if s.startswith('# '):
            push(); t = s[2:].strip()
            cur[0] = ({'type':'title','title':t,'subtitle':'','content':[]}
                      if not slides else cs(t)); continue
        if s.startswith('## '):
            push(); cur[0] = cs(s[3:].strip()); continue
        if s.startswith('### '):
            if not cur[0]: cur[0] = cs(s[4:].strip())
            else: cur[0]['content'].append({'level':0,'text':s[4:].strip(),'bold':True})
            continue
        m = re.match(r'^[-*・•]\s+(.+)', s)
        if m:
            if not cur[0]: cur[0] = cs()
            cur[0]['content'].append({'level':0,'text':m.group(1).strip()}); continue
        m2 = re.match(r'^\s{2,}[-*・•]\s+(.+)', raw)
        if m2:
            if not cur[0]: cur[0] = cs()
            cur[0]['content'].append({'level':1,'text':m2.group(1).strip()}); continue
        m3 = re.match(r'^\d+[.。)]\s+(.+)', s)
        if m3:
            if not cur[0]: cur[0] = cs()
            cur[0]['content'].append({'level':0,'text':m3.group(1).strip()}); continue
        c = cur[0]
        if not c:   cur[0] = {'type':'title','title':s,'subtitle':'','content':[]}
        elif c['type'] == 'title' and not c.get('subtitle'): c['subtitle'] = s
        else: c['content'].append({'level':0,'text':s})

    push()
    return slides or [{'type':'title','title':'Presentation','subtitle':'','content':[]}]

# ─────────────────────────────────────────
#  ⑩ UI
# ─────────────────────────────────────────
st.markdown('<div class="pg-eyebrow">Slide Generator</div>', unsafe_allow_html=True)
st.markdown('<div class="pg-title">テキストを貿るだけで<br>コンペを制すスライドが完成する。</div>',
            unsafe_allow_html=True)
st.markdown('<div class="pg-sub">DTP原則を徹底実装。トラッキング・ベースライングリッド・タイプスケール・行間比率を自動制御。</div>',
            unsafe_allow_html=True)

# ── Step 01: テーマ ───────────────────────
st.markdown('<span class="step-label">Step 01 — Design Language</span>', unsafe_allow_html=True)

theme_raw = st.radio('テーマ', [
    'MINIMAL  ——  グリッドとナンバリング。知性を纚ったコーポレートデザイン。',
    'NOIR  ——  漆黒のタイトルスライド。圧倒的な存在感でコンペを制す。',
    'BOLD  ——  パネル分割と大型タイポグラフィー。ブランドを刈む。',
], label_visibility='collapsed')
theme_key = theme_raw.split()[0]

# ── Step 02: アクセントカラー ─────────────────
st.markdown('<span class="step-label">Step 02 — Accent Color</span>', unsafe_allow_html=True)

accent_label = st.radio('アクセントカラー',
    list(ACCENT_PRESETS.keys()),
    label_visibility='collapsed')
accent_hex = ACCENT_PRESETS[accent_label]

col1, col2 = st.columns([1, 2])
with col1:
    use_custom = st.checkbox('カスタムカラー')
with col2:
    if use_custom:
        cx = st.text_input('HEX', value='#C8A96E', label_visibility='collapsed')
        try: rgb(cx); accent_hex = cx
        except: st.warning('HEX形式エラー（例: #C8A96E）')

st.markdown(
    f'<div style="margin-top:4px;">'
    f'<span class="swatch" style="background:{accent_hex};"></span>'
    f'<span style="font-size:12px;color:#999;letter-spacing:.04em;">{accent_hex}</span>'
    f'</div>', unsafe_allow_html=True)

# ── Step 03: 作成者 ───────────────────────
st.markdown('<span class="step-label">Step 03 — Author / Company</span>', unsafe_allow_html=True)
c1, _ = st.columns([1, 1])
with c1:
    author = st.text_input('組織名・作成者',
                           placeholder='例：Creative Division',
                           label_visibility='visible')

# ── Step 04: テキスト入力 ────────────────
st.markdown('<hr class="divider">', unsafe_allow_html=True)
st.markdown('<span class="step-label">Step 04 — Content</span>', unsafe_allow_html=True)

text_input = st.text_area(
    'テキスト入力', height=360,
    placeholder=(
        '# プレゼンテーションタイトル\n'
        'サブタイトル・キャッチコピー\n\n'
        '## スライドの見出し\n'
        '- 伝えたいメッセージ\n'
        '- 重要なポイント\n'
        '  - 詳細・補足\n\n'
        '## 次のスライド\n'
        '### 強調したい小見出し\n'
        '- 項目A\n'
        '- 項目B\n\n'
        '---  （スライドを手動で区切る）'
    ), label_visibility='collapsed')

with st.expander('テキストの書き方ガイド'):
    st.markdown("""
| 書き方 | 効果 |
|--------|------|
| `# タイトル` | タイトルスライド |
| `## 見出し` | コンテンツスライド見出し |
| `### 小見出し` | アクセントカラーで強調 |
| `- 項目` | 箇条書き (—  マーカー) |
| `  - サブ` | インデント箇条書き (·  マーカー) |
| `---` | スライド強制分割 |

GeminiやNotebookLMの出力をそのまま貼り付けてOKです。
""")

st.markdown('<div style="height:6px;"></div>', unsafe_allow_html=True)

if st.button('スライドを生成する'):
    if not text_input.strip():
        st.warning('テキストを入力してください。')
    else:
        with st.spinner('生成中...'):
            slides   = parse_slides(text_input)
            pptx_buf = build_pptx(slides, theme_key, accent_hex, author)

        nt = len(slides)
        nc = sum(1 for s in slides if s['type'] != 'title')
        fname = f'slides_{datetime.now().strftime("%Y%m%d_%H%M")}.pptx'

        st.success(
            f'生成完了 — {nt} 枚（コンテンツ {nc} 枚）  '
            f'[ {theme_key}  ×  {accent_hex} ]'
        )
        st.download_button(
            'PowerPoint をダウンロード', data=pptx_buf, file_name=fname,
            mime='application/vnd.openxmlformats-officedocument.presentationml.presentation')

st.markdown(
    '<div style="text-align:center;font-size:10px;color:#E0E0E0;'
    'letter-spacing:.18em;text-transform:uppercase;margin-top:64px;">'
    'Powered by Claude</div>', unsafe_allow_html=True)
