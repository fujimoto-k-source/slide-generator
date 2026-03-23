"""
Slide Generator v2.1 — AI-Powered Presentation Engine
Gemini Content Analysis + Auto Charts + Smart Layout + AI Imagery
"""

import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from lxml import etree
import io, re, json, math
from datetime import datetime

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import numpy as np
from PIL import Image, ImageDraw, ImageFilter

try:
    from google import genai
    from google.genai import types as genai_types
    GEMINI_AVAILABLE = True
except ImportError:
    GEMINI_AVAILABLE = False

# =============================================
#  Page Config & CSS
# =============================================
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

div[data-testid="stRadio"] label { display:none!important; }
div[data-testid="stRadio"] > div { display:flex; flex-direction:column; gap:8px; }
div[data-testid="stRadio"] > div > label {
    border:1px solid #EAEAEA!important; border-radius:2px!important;
    padding:18px 22px!important; background:#FAFAFA!important;
    cursor:pointer!important; transition:all .15s!important; }
div[data-testid="stRadio"] > div > label:hover {
    border-color:#111!important; background:#FFF!important; }
div[data-testid="stRadio"] > div > label[data-checked="true"] {
    border-color:#111!important; background:#FFF!important;
    box-shadow:inset 3px 0 0 #111!important; }

label, .stTextArea label, .stTextInput label {
    font-size:10px!important; font-weight:700!important;
    letter-spacing:.15em!important; text-transform:uppercase!important; color:#BBB!important; }
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

.swatch { display:inline-block; width:13px; height:13px; border-radius:1px;
          margin-right:8px; vertical-align:middle; border:1px solid rgba(0,0,0,.08); }
.divider { border:none; border-top:1px solid #F0F0F0; margin:44px 0; }
[data-testid="column"] { padding:0 6px!important; }

.preview-card { background:#FAFAFA; border:1px solid #EAEAEA; border-radius:2px;
                padding:12px 16px; margin:6px 0; }
.preview-tag  { font-size:9px; font-weight:700; letter-spacing:.15em;
                text-transform:uppercase; color:#BBB; }
.preview-name { font-size:13px; font-weight:600; color:#222; margin:3px 0 0; }
</style>
""", unsafe_allow_html=True)

# =============================================
#  DTP Core Utilities
# =============================================
W, H = 13.33, 7.5

def _rgb(h):
    h = h.lstrip('#')
    return RGBColor(int(h[0:2],16), int(h[2:4],16), int(h[4:6],16))

def _hext(h):
    h = h.lstrip('#')
    return (int(h[0:2],16), int(h[2:4],16), int(h[4:6],16))

def _hexm(h):
    return h if h.startswith('#') else f'#{h}'

def _run(run, font='Calibri', size=None, bold=False, italic=False,
         color=None, tracking=0, ea='Meiryo UI'):
    run.font.name = font
    if size: run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    if color: run.font.color.rgb = color
    rPr = run._r.get_or_add_rPr()
    if tracking:
        rPr.set('spc', str(int(tracking)))
    for ch in list(rPr):
        if ch.tag == qn('a:ea'): rPr.remove(ch)
    etree.SubElement(rPr, qn('a:ea')).set('typeface', ea)

def _para(p, align=PP_ALIGN.LEFT, sb=0, sa=0, lh=1.0):
    p.alignment = align
    p.space_before = Pt(sb)
    p.space_after = Pt(sa)
    if lh != 1.0:
        pPr = p._p.get_or_add_pPr()
        for ch in list(pPr):
            if ch.tag == qn('a:lnSpc'): pPr.remove(ch)
        spc = etree.SubElement(etree.SubElement(pPr, qn('a:lnSpc')), qn('a:spcPct'))
        spc.set('val', str(int(lh * 100000)))

def _rect(sl, l, t, w, h, fill=None):
    s = sl.shapes.add_shape(1, Inches(l), Inches(t), Inches(w), Inches(h))
    if fill:
        s.fill.solid(); s.fill.fore_color.rgb = fill
    else:
        s.fill.background()
    s.line.fill.background()
    return s

def _tb(sl, txt, l, t, w, h, size=16, bold=False, italic=False,
        color=None, align=PP_ALIGN.LEFT, tracking=0, lh=1.0,
        font='Calibri', ea='Meiryo UI', sb=0):
    tb = sl.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]
    _para(p, align=align, sb=sb, lh=lh)
    r = p.add_run(); r.text = str(txt)
    _run(r, font=font, size=size, bold=bold, italic=italic,
         color=color, tracking=tracking, ea=ea)
    return tb

def _img(sl, data, l, t, w, h):
    if data is None: return
    if isinstance(data, Image.Image):
        buf = io.BytesIO()
        data.convert('RGB').save(buf, format='PNG')
        buf.seek(0)
        sl.shapes.add_picture(buf, Inches(l), Inches(t), Inches(w), Inches(h))
    elif isinstance(data, (bytes, bytearray)):
        sl.shapes.add_picture(io.BytesIO(data), Inches(l), Inches(t), Inches(w), Inches(h))
    elif isinstance(data, io.BytesIO):
        data.seek(0)
        sl.shapes.add_picture(data, Inches(l), Inches(t), Inches(w), Inches(h))

# =============================================
#  Theme
# =============================================
LM = 0.75
RM = 0.75
CW = W - LM - RM

ACCENT_PRESETS = {
    'Navy':     '#0A2463',
    'Gold':     '#A67C3A',
    'Crimson':  '#8C1C2E',
    'Emerald':  '#0B5E4A',
    'Slate':    '#2E3F4F',
}

def _theme(ahex):
    a = _rgb(ahex)
    return dict(
        accent=a, ahex=ahex,
        white=_rgb('FFFFFF'), ink=_rgb('0D0D0D'),
        body=_rgb('2C2C2C'), sub=_rgb('717171'),
        rules=_rgb('EBEBEB'), numc=_rgb('CCCCCC'),
        muted=_rgb('9A9A9A'), bg=_rgb('F5F5F5'),
    )

# =============================================
#  Gemini API
# =============================================
ANALYSIS_PROMPT = """You are a presentation design expert.
Analyze the following text and output a professional slide deck structure in JSON.

RULES:
1. Assign the best layout type to each slide
2. Extract chart_data when numerical data exists
3. Add image_prompt (in English, minimal corporate style) for visual slides
4. First slide MUST be "title_hero"
5. Last slide MUST be "closing"
6. NEVER omit any information from the text. Reflect ALL content.
7. All image_prompt must include: "minimal geometric corporate illustration, clean white background, subtle shadows, no text, no labels"
8. Output slide titles, subtitles, bullets, stat_label, step titles/descriptions, card titles/descriptions, and ALL text content in the SAME LANGUAGE as the input text. If the input is Japanese, all output text must be Japanese.

LAYOUT TYPES:
- title_hero: Title slide (cover)
- stat_callout: Big KPI/number emphasis (stat_value + stat_label required)
- split_image: Text left + image right (image_prompt required)
- process_flow: Steps/flow (steps array required, each with label/title/desc)
- comparison: Compare/contrast (left/right required, each with title + items array)
- chart_full: Chart-dominant (chart_data required)
- grid_cards: Multiple item cards (cards array required, each with title + desc)
- closing: Ending slide

chart_data structure:
{"type": "bar"|"line"|"pie"|"horizontal_bar", "labels": [...], "values": [...], "unit": "..."}
Multi-series: {"type":"line","labels":[...],"datasets":[{"name":"...","values":[...]}],"unit":"..."}

Output: {"slides": [...]}

TEXT:
"""

def _init_gemini(key):
    if not GEMINI_AVAILABLE:
        return None
    try:
        return genai.Client(api_key=key)
    except Exception as e:
        st.error(f'Gemini 初期化エラー: {e}')
        return None

def _analyze(client, text):
    try:
        r = client.models.generate_content(
            model='gemini-2.0-flash',
            contents=ANALYSIS_PROMPT + text,
            config=genai_types.GenerateContentConfig(
                response_mime_type='application/json',
            ),
        )
        data = json.loads(r.text)
        return data.get('slides', data if isinstance(data, list) else [])
    except Exception as e:
        st.error(f'分析エラー: {e}')
        return None

def _gen_image(client, prompt):
    if not client or not prompt:
        return None
    full = (
        f"Create a minimal, professional illustration for a presentation slide. "
        f"Style: clean geometric corporate art, white background, subtle gradients, "
        f"elegant shadows, absolutely no text or labels in the image. "
        f"Subject: {prompt}"
    )
    # Try Gemini image generation
    try:
        r = client.models.generate_content(
            model='gemini-2.0-flash-exp',
            contents=full,
            config=genai_types.GenerateContentConfig(
                response_modalities=['IMAGE', 'TEXT'],
            ),
        )
        for part in r.candidates[0].content.parts:
            if hasattr(part, 'inline_data') and part.inline_data and part.inline_data.data:
                return Image.open(io.BytesIO(part.inline_data.data))
    except Exception:
        pass
    # Fallback: Imagen
    try:
        r = client.models.generate_images(
            model='imagen-3.0-generate-002',
            prompt=full,
            config=genai_types.GenerateImagesConfig(number_of_images=1),
        )
        if r.generated_images:
            return r.generated_images[0].image
    except Exception:
        pass
    return None

# =============================================
#  Abstract Visual Fallback (Pillow)
# =============================================
def _abstract(ahex, w=1200, h=675):
    img = Image.new('RGBA', (w, h), (255,255,255,255))
    base = _hext(ahex)

    ov = Image.new('RGBA', (w,h), (0,0,0,0))
    od = ImageDraw.Draw(ov)

    cx, cy, r = int(w*0.72), int(h*0.62), int(w*0.38)
    for i in range(r, 0, -3):
        a = int(22 * (i/r))
        od.ellipse([cx-i, cy-i, cx+i, cy+i], fill=(*base, a))

    cx2, cy2, r2 = int(w*0.18), int(h*0.22), int(w*0.18)
    for i in range(r2, 0, -3):
        a = int(14 * (i/r2))
        od.ellipse([cx2-i, cy2-i, cx2+i, cy2+i], fill=(*base, a))

    img = Image.alpha_composite(img, ov).convert('RGB')
    img = img.filter(ImageFilter.GaussianBlur(radius=35))

    buf = io.BytesIO()
    img.save(buf, format='PNG', quality=95)
    buf.seek(0)
    return buf.getvalue()

# =============================================
#  Chart Engine (matplotlib)
# =============================================
def _sfloats(lst):
    out = []
    for v in (lst or []):
        try: out.append(float(v))
        except: out.append(0)
    return out

def _cbase(ahex, w=7, h=4):
    fig, ax = plt.subplots(figsize=(w, h), dpi=200)
    fig.patch.set_facecolor('white')
    fig.patch.set_alpha(0)
    ax.set_facecolor('white')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_color('#E8E8E8')
    ax.spines['bottom'].set_color('#E8E8E8')
    ax.tick_params(colors='#888888', labelsize=9)
    return fig, ax

def _csave(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format='png', bbox_inches='tight', transparent=True, dpi=200)
    plt.close(fig)
    buf.seek(0)
    return buf.getvalue()

def _cbar(data, ahex):
    labels = data.get('labels', [])
    values = _sfloats(data.get('values', []))
    if not labels or not values: return None
    fig, ax = _cbase(ahex)
    accent = _hexm(ahex)
    n = len(values)
    colors = ['#DCDCDC'] * n
    if values:
        colors[values.index(max(values))] = accent
    ax.bar(labels, values, color=colors, width=0.55, edgecolor='none')
    mx = max(values) if values else 1
    for i, (l, v) in enumerate(zip(labels, values)):
        ax.text(i, v + mx*0.02, f'{v:,.0f}', ha='center', va='bottom',
                fontsize=9, fontweight='bold', color='#333')
    ax.set_ylim(0, mx*1.18)
    u = data.get('unit', '')
    if u: ax.set_ylabel(u, fontsize=9, color='#AAA')
    fig.tight_layout(pad=0.8)
    return _csave(fig)

def _cline(data, ahex):
    labels = data.get('labels', [])
    datasets = data.get('datasets', [])
    if not datasets:
        v = _sfloats(data.get('values', []))
        if v: datasets = [{'name': '', 'values': v}]
    if not labels or not datasets: return None
    fig, ax = _cbase(ahex)
    accent = _hexm(ahex)
    palette = [accent, '#AAAAAA', '#CCCCCC', '#666666']
    for i, ds in enumerate(datasets):
        vals = _sfloats(ds.get('values', []))
        c = palette[i % len(palette)]
        ax.plot(labels[:len(vals)], vals, color=c, linewidth=2.5,
                marker='o', markersize=5, markerfacecolor='white',
                markeredgecolor=c, markeredgewidth=2)
        if vals:
            ax.text(len(vals)-1, vals[-1], f'  {vals[-1]:,.0f}',
                    va='center', fontsize=8, fontweight='bold', color=c)
    u = data.get('unit', '')
    if u: ax.set_ylabel(u, fontsize=9, color='#AAA')
    fig.tight_layout(pad=0.8)
    return _csave(fig)

def _cpie(data, ahex):
    labels = data.get('labels', [])
    values = _sfloats(data.get('values', []))
    if not labels or not values: return None
    fig, ax = _cbase(ahex, w=5, h=4)
    base = _hext(ahex)
    colors = []
    for i in range(len(values)):
        f = 1.0 - i*0.14
        c = tuple(min(255, int(v*f + 255*(1-f)*0.3)) for v in base)
        colors.append(f'#{c[0]:02x}{c[1]:02x}{c[2]:02x}')
    wedges, _, autotexts = ax.pie(
        values, labels=None, autopct='%1.0f%%', colors=colors, startangle=90,
        wedgeprops=dict(width=0.42, edgecolor='white', linewidth=2.5), pctdistance=0.78)
    for t in autotexts:
        t.set_fontsize(9); t.set_fontweight('bold'); t.set_color('#333')
    ax.legend(labels, loc='center', fontsize=8, frameon=False)
    fig.tight_layout(pad=0.5)
    return _csave(fig)

def _chbar(data, ahex):
    labels = data.get('labels', [])
    values = _sfloats(data.get('values', []))
    if not labels or not values: return None
    fig, ax = _cbase(ahex)
    accent = _hexm(ahex)
    y = range(len(labels))
    ax.barh(list(y), values, color=accent, height=0.5, edgecolor='none', alpha=0.85)
    ax.set_yticks(list(y)); ax.set_yticklabels(labels, fontsize=9)
    ax.invert_yaxis()
    mx = max(values) if values else 1
    for bar, val in zip(ax.patches, values):
        ax.text(bar.get_width() + mx*0.02, bar.get_y()+bar.get_height()/2,
                f'{val:,.0f}', va='center', fontsize=9, fontweight='bold', color='#333')
    ax.spines['left'].set_visible(False)
    fig.tight_layout(pad=0.8)
    return _csave(fig)

def _make_chart(cd, ahex):
    if not cd: return None
    t = cd.get('type', 'bar')
    fn = {'bar': _cbar, 'line': _cline, 'pie': _cpie,
          'donut': _cpie, 'horizontal_bar': _chbar}.get(t, _cbar)
    try:
        return fn(cd, ahex)
    except Exception:
        return None

# =============================================
#  Layout: Title Hero
# =============================================
def _lay_title(sl, sd, th, ahex, img=None):
    _rect(sl, 0, 0, W, 0.04, fill=th['accent'])

    if img:
        _img(sl, img, W*0.48, 0.04, W*0.52, H-0.08)
        # Fade overlay (white gradient)
        ov = Image.new('RGBA', (400, 675), (255,255,255,0))
        d = ImageDraw.Draw(ov)
        for x in range(400):
            a = int(255 * (1 - x/400))
            d.line([(x,0),(x,675)], fill=(255,255,255,a))
        buf = io.BytesIO(); ov.save(buf, format='PNG'); buf.seek(0)
        _img(sl, buf.getvalue(), W*0.42, 0.04, W*0.16, H-0.08)

    _tb(sl, sd.get('title',''), LM, 1.6, 6.8, 2.4,
        size=52, bold=True, color=th['ink'], tracking=-100, lh=1.08)
    _rect(sl, LM, 4.15, 2.0, 0.025, fill=th['accent'])
    sub = sd.get('subtitle','')
    if sub:
        _tb(sl, sub, LM, 4.40, 6.2, 0.9,
            size=16, color=th['muted'], lh=1.55)
    _rect(sl, 0, H-0.04, W, 0.04, fill=th['accent'])

# =============================================
#  Layout: Stat Callout
# =============================================
def _lay_stat(sl, sd, th, ahex, chart=None):
    _tb(sl, sd.get('title',''), LM, 0.35, CW, 0.5,
        size=11, bold=True, color=th['muted'], tracking=120)
    _rect(sl, LM, 0.92, CW, 0.008, fill=th['rules'])

    _tb(sl, sd.get('stat_value',''), LM, 1.15, 6.0, 2.6,
        size=96, bold=True, color=th['accent'], tracking=-80, lh=1.0)
    _tb(sl, sd.get('stat_label',''), LM, 3.55, 5.5, 0.5,
        size=16, color=th['sub'])

    bullets = sd.get('bullets', [])
    y = 4.25
    for b in bullets[:3]:
        _tb(sl, f'\u2014  {b}', LM, y, 5.5, 0.4,
            size=13, color=th['body'], lh=1.4)
        y += 0.48

    if chart:
        _img(sl, chart, 7.0, 1.2, 5.6, 4.6)
    _rect(sl, 0, H-0.04, W, 0.04, fill=th['accent'])

# =============================================
#  Layout: Split Image
# =============================================
def _lay_split(sl, sd, th, ahex, img=None):
    _tb(sl, sd.get('title',''), LM, 0.35, 5.8, 0.7,
        size=28, bold=True, color=th['ink'], tracking=-40, lh=1.12)
    _rect(sl, LM, 1.15, 1.5, 0.02, fill=th['accent'])

    bullets = sd.get('bullets', [])
    y = 1.50
    for b in bullets:
        _tb(sl, f'\u2014  {b}', LM, y, 5.6, 0.52,
            size=15, color=th['body'], lh=1.45)
        y += 0.62

    if img:
        _img(sl, img, 7.2, 0.45, 5.4, 6.6)
    else:
        _rect(sl, 7.2, 0.45, 5.4, 6.6, fill=th['bg'])
    _rect(sl, 0, H-0.04, W, 0.04, fill=th['accent'])

# =============================================
#  Layout: Process Flow
# =============================================
def _lay_flow(sl, sd, th, ahex):
    _tb(sl, sd.get('title',''), LM, 0.35, CW, 0.7,
        size=28, bold=True, color=th['ink'], tracking=-40, lh=1.12)
    _rect(sl, LM, 1.15, CW, 0.008, fill=th['rules'])
    _rect(sl, LM, 1.15, 1.5, 0.008, fill=th['accent'])

    steps = sd.get('steps', [])
    n = len(steps)
    if n == 0: return
    _rect(sl, 0, H-0.04, W, 0.04, fill=th['accent'])

    gap = 0.35
    sw = min(2.8, (CW - gap*(n-1)) / n)
    total = n*sw + (n-1)*gap
    sx = LM + (CW - total) / 2
    circ = 0.65
    cy_top = 2.0

    for i, step in enumerate(steps):
        x = sx + i*(sw + gap)
        cx = x + sw/2

        oval = sl.shapes.add_shape(
            MSO_SHAPE.OVAL,
            Inches(cx - circ/2), Inches(cy_top),
            Inches(circ), Inches(circ))
        oval.fill.solid(); oval.fill.fore_color.rgb = th['accent']
        oval.line.fill.background()

        lbl = step.get('label', f'{i+1:02d}')
        _tb(sl, lbl, cx-circ/2, cy_top+0.12, circ, circ-0.12,
            size=18, bold=True, color=th['white'], align=PP_ALIGN.CENTER)

        if i < n-1:
            ax1 = cx + circ/2 + 0.06
            ncx = sx + (i+1)*(sw+gap) + sw/2
            ax2 = ncx - circ/2 - 0.06
            _rect(sl, ax1, cy_top + circ/2 - 0.006, ax2-ax1, 0.012, fill=th['rules'])

        _tb(sl, step.get('title',''), x, cy_top + circ + 0.20, sw, 0.5,
            size=14, bold=True, color=th['ink'], align=PP_ALIGN.CENTER, tracking=-10, lh=1.12)

        desc = step.get('desc','')
        if desc:
            _tb(sl, desc, x, cy_top + circ + 0.72, sw, 1.2,
                size=11, color=th['sub'], align=PP_ALIGN.CENTER, lh=1.42)

# =============================================
#  Layout: Comparison
# =============================================
def _lay_compare(sl, sd, th, ahex):
    _tb(sl, sd.get('title',''), LM, 0.35, CW, 0.7,
        size=28, bold=True, color=th['ink'], tracking=-40, lh=1.12)
    _rect(sl, LM, 1.15, CW, 0.008, fill=th['rules'])

    hw = (CW - 0.5) / 2
    left = sd.get('left', {})
    right = sd.get('right', {})

    _rect(sl, LM, 1.45, hw, 5.3, fill=th['bg'])
    _tb(sl, left.get('title','Before'), LM+0.3, 1.72, hw-0.6, 0.5,
        size=18, bold=True, color=th['accent'], tracking=-20)
    y = 2.35
    for item in left.get('items', []):
        _tb(sl, f'\u2014  {item}', LM+0.3, y, hw-0.6, 0.4,
            size=13, color=th['body'], lh=1.4)
        y += 0.50

    mid = LM + hw + 0.18
    _rect(sl, mid, 1.65, 0.012, 4.8, fill=th['accent'])

    rx = LM + hw + 0.5
    _rect(sl, rx, 1.45, hw, 5.3, fill=th['bg'])
    _tb(sl, right.get('title','After'), rx+0.3, 1.72, hw-0.6, 0.5,
        size=18, bold=True, color=th['accent'], tracking=-20)
    y = 2.35
    for item in right.get('items', []):
        _tb(sl, f'\u2014  {item}', rx+0.3, y, hw-0.6, 0.4,
            size=13, color=th['body'], lh=1.4)
        y += 0.50

    _rect(sl, 0, H-0.04, W, 0.04, fill=th['accent'])

# =============================================
#  Layout: Chart Full
# =============================================
def _lay_chart(sl, sd, th, ahex, chart=None):
    _tb(sl, sd.get('title',''), LM, 0.35, CW-1, 0.7,
        size=28, bold=True, color=th['ink'], tracking=-40, lh=1.12)
    _rect(sl, LM, 1.15, CW, 0.008, fill=th['rules'])
    _rect(sl, LM, 1.15, 1.5, 0.008, fill=th['accent'])

    if chart:
        _img(sl, chart, LM+0.2, 1.50, CW-0.4, 5.0)

    bullets = sd.get('bullets', [])
    if bullets and not chart:
        y = 1.60
        for b in bullets:
            _tb(sl, f'\u2014  {b}', LM, y, CW, 0.4,
                size=14, color=th['body'], lh=1.45)
            y += 0.55
    _rect(sl, 0, H-0.04, W, 0.04, fill=th['accent'])

# =============================================
#  Layout: Grid Cards
# =============================================
def _lay_grid(sl, sd, th, ahex):
    _tb(sl, sd.get('title',''), LM, 0.35, CW, 0.7,
        size=28, bold=True, color=th['ink'], tracking=-40, lh=1.12)
    _rect(sl, LM, 1.15, CW, 0.008, fill=th['rules'])
    _rect(sl, LM, 1.15, 1.5, 0.008, fill=th['accent'])

    cards = sd.get('cards', [])[:6]
    n = len(cards)
    if n == 0: return
    _rect(sl, 0, H-0.04, W, 0.04, fill=th['accent'])

    cols = min(n, 3)
    rows = math.ceil(n / cols)
    gap = 0.22
    cw = (CW - gap*(cols-1)) / cols
    ch = 2.15 if rows == 1 else 1.85
    sy = 1.50

    for i, card in enumerate(cards):
        col = i % cols
        row = i // cols
        x = LM + col*(cw + gap)
        y = sy + row*(ch + gap)

        _rect(sl, x, y, cw, ch, fill=th['bg'])
        _rect(sl, x, y, cw, 0.032, fill=th['accent'])

        _tb(sl, f'{i+1:02d}', x+0.2, y+0.22, 0.5, 0.3,
            size=9, bold=True, color=th['numc'], tracking=80)
        _tb(sl, card.get('title',''), x+0.2, y+0.52, cw-0.4, 0.4,
            size=14, bold=True, color=th['ink'], tracking=-10, lh=1.15)
        desc = card.get('desc','')
        if desc:
            _tb(sl, desc, x+0.2, y+0.95, cw-0.4, ch-1.15,
                size=11, color=th['sub'], lh=1.4)

# =============================================
#  Layout: Closing
# =============================================
def _lay_close(sl, sd, th, ahex, author=''):
    _rect(sl, 0, 0, W, H, fill=th['accent'])
    _tb(sl, sd.get('title','Thank You'), LM, 2.0, CW, 2.2,
        size=52, bold=True, color=th['white'],
        align=PP_ALIGN.CENTER, tracking=-80, lh=1.05)
    _rect(sl, W/2-1.0, 4.35, 2.0, 0.018, fill=th['white'])
    sub = sd.get('subtitle','')
    if sub:
        _tb(sl, sub, LM, 4.65, CW, 0.6,
            size=14, color=th['white'], align=PP_ALIGN.CENTER, lh=1.5)
    if author:
        _tb(sl, author.upper(), LM, 6.5, CW, 0.4,
            size=10, bold=True, color=th['white'],
            align=PP_ALIGN.CENTER, tracking=120)

# =============================================
#  Master Builder
# =============================================
def _build(slides_data, ahex, author, client=None, gen_img=True, prog=None):
    prs = Presentation()
    prs.slide_width = Inches(W)
    prs.slide_height = Inches(H)
    blank = prs.slide_layouts[6]
    th = _theme(ahex)
    total = len(slides_data)

    for i, sd in enumerate(slides_data):
        sl = prs.slides.add_slide(blank)
        lay = sd.get('layout', 'split_image')

        if prog:
            prog(i, total, f'スライド {i+1}/{total} を生成中: {lay}')

        chart = _make_chart(sd.get('chart_data'), ahex)

        img = None
        ip = sd.get('image_prompt')
        if ip and gen_img and client:
            img = _gen_image(client, ip)
        if img is None and lay in ('title_hero', 'split_image'):
            img = _abstract(ahex)

        if lay == 'title_hero':
            _lay_title(sl, sd, th, ahex, img)
        elif lay == 'stat_callout':
            _lay_stat(sl, sd, th, ahex, chart)
        elif lay == 'split_image':
            _lay_split(sl, sd, th, ahex, img)
        elif lay == 'process_flow':
            _lay_flow(sl, sd, th, ahex)
        elif lay == 'comparison':
            _lay_compare(sl, sd, th, ahex)
        elif lay == 'chart_full':
            _lay_chart(sl, sd, th, ahex, chart)
        elif lay == 'grid_cards':
            _lay_grid(sl, sd, th, ahex)
        elif lay == 'closing':
            _lay_close(sl, sd, th, ahex, author)
        else:
            _lay_split(sl, sd, th, ahex, img)

    buf = io.BytesIO()
    prs.save(buf); buf.seek(0)
    return buf

# =============================================
#  UI
# =============================================
st.markdown('<div class="pg-eyebrow">Slide Generator</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="pg-title">テキストを貼るだけで<br>'
    'コンペを制すスライドが完成する。</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="pg-sub">'
    'AI構成分析 / 自動チャート / スマートレイアウト / AI画像生成</div>',
    unsafe_allow_html=True)

# Step 01
st.markdown('<span class="step-label">Step 01 ─ Gemini API キー</span>',
            unsafe_allow_html=True)
api_key = st.text_input(
    'API Key', type='password',
    placeholder='Gemini API キー（Google AI Studio で取得）',
    label_visibility='collapsed',
    help='https://aistudio.google.com/apikey から無料で取得できます')

# Step 02
st.markdown('<span class="step-label">Step 02 ─ アクセントカラー</span>',
            unsafe_allow_html=True)
accent_label = st.radio('Accent', list(ACCENT_PRESETS.keys()),
                         label_visibility='collapsed')
accent_hex = ACCENT_PRESETS[accent_label]

col1, col2 = st.columns([1, 2])
with col1:
    use_custom = st.checkbox('カスタムカラー')
with col2:
    if use_custom:
        cx = st.text_input('HEX', value='#C8A96E', label_visibility='collapsed')
        try:
            _rgb(cx); accent_hex = cx
        except:
            st.warning('HEX形式が正しくありません')

st.markdown(
    f'<div style="margin-top:4px;">
    f'<span class="swatch" style="background:{accent_hex};"></span>
    f'<span style="font-size:12px;color:#999;letter-spacing:.04em;">{accent_hex}</span>
    f'</div>', unsafe_allow_html=True)

# Step 03
st.markdown('<span class="step-label">Step 03 ─ 作成者 / 会社名</span>',
            unsafe_allow_html=True)
c1, _ = st.columns([1, 1])
with c1:
    author = st.text_input('作成者', placeholder='例: クリエイティブ事業部',
                            label_visibility='visible')

# Step 04
st.markdown('<hr class="divider">', unsafe_allow_html=True)
st.markdown('<span class="step-label">Step 04 ─ コンテンツ入力</span>',
            unsafe_allow_html=True)

text_input = st.text_area(
    'Content', height=360,
    placeholder=(
        '# プレゼンタイトル\n'
        'サブタイトル\n\n'
        '## 市場概況\n'
        '- 市場規模は2024年時点で3.2兆円\n'
        '- 前年比130%の成長率\n\n'
        '## 実現ステップ\n'
        '1. 企画\n2. 設計\n3. 実装\n\n'
        'Gemini / NotebookLM の出力をそのまま貼り付けOK'),
    label_visibility='collapsed')

# Options
st.markdown('<span class="step-label">オプション</span>', unsafe_allow_html=True)
gen_images = st.checkbox('AI画像を生成する', value=True,
    help='Geminiでスライドごとにコンセプト画像を自動生成します')

st.markdown('<div style="height:6px;"></div>', unsafe_allow_html=True)

# Generate
if st.button('スライドを生成'):
    if not text_input.strip():
        st.warning('テキストを入力してください。')
    elif not api_key.strip():
        st.warning('Gemini APIキーを入力してください。')
    else:
        client = _init_gemini(api_key.strip())
        if not client:
            st.error('Geminiの初期化に失敗しました。APIキーを確認してください。')
        else:
            bar = st.progress(0, text='コンテンツを分析中...')
            bar.progress(10, text='AIがコンテンツの構造を分析しています...')
            slides = _analyze(client, text_input)

            if not slides:
                st.error('コンテンツの分析に失敗しました。')
            else:
                bar.progress(25, text=f'{len(slides)}枚のスライドを構成しました')

                with st.expander(f'スライド構成（{len(slides)}枚）',
                                 expanded=False):
                    for i, sd in enumerate(slides):
                        lay = sd.get('layout', '')
                        ttl = sd.get('title', '')
                        st.markdown(
                            f'<div class="preview-card">
                            f'<div class="preview-tag">{i+1:02d} / {lay}</div>
                            f'<div class="preview-name">{ttl}</div>
                            f'</div>', unsafe_allow_html=True)

                def update(i, total, msg):
                    pct = 25 + int(65 * (i+1) / total)
                    bar.progress(pct, text=msg)

                pptx_buf = _build(
                    slides, accent_hex, author,
                    client=client, gen_img=gen_images,
                    prog=update)

                bar.progress(100, text='生成完了')

                nt = len(slides)
                layouts = len(set(sd.get('layout','') for sd in slides))
                fname = f'slides_{datetime.now().strftime("%Y%m%d_%H%M")}.pptx'

                st.success(f'{nt}枚のスライド / {layouts}種類のレイアウト / {accent_hex}')
                st.download_button(
                    'PowerPoint をダウンロード',
                    data=pptx_buf, file_name=fname,
                    mime='application/vnd.openxmlformats-officedocument.presentationml.presentation')

st.markdown(
    '<div style="text-align:center;font-size:10px;color:#E0E0E0;
    'letter-spacing:.18em;text-transform:uppercase;margin-top:64px;">
    'Powered by Gemini + Claude</div>', unsafe_allow_html=True)

