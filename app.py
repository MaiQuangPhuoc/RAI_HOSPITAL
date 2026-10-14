"""PDF -> KDパンチ入力シート (demo).   Chạy:  streamlit run app.py

Cấu trúc thư mục (đường dẫn tương đối so với app.py):
    input/
        app.py
        data/
            dic_output.txt
            sheet/
                template.xlsx      (file xlsx chỉ có tới dòng 3)
(nếu không có thư mục 'input' thì dùng  data/ cạnh app.py)
"""
import ast
import base64
import colorsys
import io
import re
import time
from concurrent.futures import ThreadPoolExecutor
from functools import lru_cache
from pathlib import Path

import streamlit as st
import streamlit.components.v1 as components
from datetime import datetime
from html import escape
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# ======================= CẤU HÌNH =======================
from pathlib import Path

BASE = Path(__file__).resolve().parent  # D:\rikai\RAI-LLM\code

# File từ điển
DIC_PATH = BASE / 'dic_output.txt'

# File Excel
DATA_DIR = BASE.parent / 'data'         # D:\rikai\RAI-LLM\data
SHEET_DIR = DATA_DIR
TEMPLATE_NAME = 'sheet.xlsx'
TEMPLATE_PATH = SHEET_DIR / TEMPLATE_NAME

DIC_NAME = 'dic_output.txt'
SHEET_NAME = '入力シート'

SECONDS_PER_FILE = 3
FIRST_DATA_ROW = 4
FREEZE_ROWS = 3
FREEZE_COLS = 6
LEFT_PERCENT = 30

RESULT_NAME = 'KDパンチ入力シート_ketqua.xlsx'

# Số trong tên file -> tên dict
# 1 -> dic_4, 2 -> dic_6, 3 -> dic_8, 4 -> dic_10, 5 -> dic_12
DIC_BY_NUMBER = {n: f'dic_{2 + 2 * n}' for n in range(1, 6)}

# Dòng lẻ: chỉ điền các cột này
ODD_DEFAULTS = {
    '回目': 2,
    '入力日': '4/9',
    '入力者': 'iida',
    'Verify': 0
}

GREEN, GREEN_DARK = '#217346', '#185c37'
# ========================================================


def find_template():
    """Tìm file Excel mẫu."""
    if TEMPLATE_PATH.exists():
        return TEMPLATE_PATH

    raise FileNotFoundError(
        f'テンプレートExcelファイルが見つかりません: {TEMPLATE_PATH}'
    )


def find_dic():
    """Tìm file từ điển."""
    if DIC_PATH.exists():
        return DIC_PATH

    raise FileNotFoundError(
        f'辞書ファイルが見つかりません: {DIC_PATH}'
    )


def load_dicts(path):
    """Đọc file dic_output.txt (dạng dic_4 = {...}) -> {'dic_4': {...}, ...}"""
    tree = ast.parse(Path(path).read_text(encoding='utf-8'))
    return {n.targets[0].id: ast.literal_eval(n.value)
            for n in tree.body if isinstance(n, ast.Assign)}


def file_number(filename):
    """abx_1.pdf -> 1 ; data-1.pdf -> 1 ; không có số -> None"""
    stem = Path(filename).stem
    found = re.findall(r'[_\-\s](\d+)', stem)
    if found:
        return int(found[-1])
    found = re.findall(r'\d+', stem)
    return int(found[0]) if found else None


def header_keys(ws):
    """Tên cột từ dòng 2/3 (cùng logic bước trước), tên trùng thì thêm _2, _3."""
    row2 = [c.value for c in ws[2]]
    row3 = [c.value for c in ws[3]]
    zip_idx = next(i for i, v in enumerate(row2) if v == '郵便番号')
    names = [(row2[i] if row2[i] is not None else row3[i]) if i <= zip_idx else row2[i]
             for i in range(len(row2))]
    count, keys = {}, []
    for n in names:
        count[n] = count.get(n, 0) + 1
        keys.append(n if count[n] == 1 else f'{n}_{count[n]}')
    return keys


# ---------------------------- Màu / viền từ file xlsx ----------------------------
def _lum(hexrgb):
    r, g, b = (int(hexrgb[i:i + 2], 16) for i in (0, 2, 4))
    return (0.299 * r + 0.587 * g + 0.114 * b) / 255


@lru_cache
def _tint(hexrgb, tint):
    r, g, b = (int(hexrgb[i:i + 2], 16) / 255 for i in (0, 2, 4))
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    l = l * (1 + tint) if tint < 0 else l * (1 - tint) + tint
    r, g, b = colorsys.hls_to_rgb(h, min(max(l, 0), 1), s)
    return '%02x%02x%02x' % (round(r * 255), round(g * 255), round(b * 255))


def theme_palette(wb):
    """Bảng màu theme của file xlsx, theo thứ tự chỉ số theme của Excel."""
    xml = getattr(wb, 'loaded_theme', None)
    if not xml:
        return []
    s = xml.decode('utf-8', 'ignore') if isinstance(xml, bytes) else xml
    m = re.search(r'<a:clrScheme.*?</a:clrScheme>', s, re.S)
    if not m:
        return []
    d = dict(re.findall(r'<a:(dk1|lt1|dk2|lt2|accent[1-6])>\s*<a:(?:srgbClr val="|sysClr val="\w+" lastClr=")'
                        r'([0-9A-Fa-f]{6})', m.group(0)))
    order = ['lt1', 'dk1', 'lt2', 'dk2'] + [f'accent{i}' for i in range(1, 7)]
    return [d.get(k) for k in order]


def color_hex(color, pal):
    """Màu openpyxl (rgb/theme+tint) -> 'rrggbb' hoặc None."""
    if color is None:
        return None
    if color.type == 'rgb' and isinstance(color.rgb, str):
        return None if color.rgb == '00000000' else color.rgb[-6:].lower()
    if color.type == 'theme' and pal and color.theme is not None and color.theme < len(pal) and pal[color.theme]:
        base = pal[color.theme].lower()
        return _tint(base, color.tint) if color.tint else base
    return None


BORDER_W = {'thin': '1px solid', 'hair': '1px solid', 'dotted': '1px dotted', 'dashed': '1px dashed',
            'dashDot': '1px dashed', 'dashDotDot': '1px dashed', 'medium': '2px solid',
            'mediumDashed': '2px dashed', 'mediumDashDot': '2px dashed', 'thick': '3px solid',
            'double': '3px double', 'slantDashDot': '2px dashed', 'mediumDashDotDot': '2px dashed'}


def border_css(side, b, pal):
    if b is None or not b.style:
        return None
    col = color_hex(b.color, pal)
    col = f'#{col}' if col and _lum(col) > 0.15 else 'var(--bord)'   # viền đen -> đổi màu theo sáng/tối
    return f'border-{side}:{BORDER_W.get(b.style, "1px solid")} {col}'


# ---------------------------- Giao diện sheet kiểu Excel (2 khung) ----------------------------
THEMES = {
    'light': dict(bg='#ffffff', fg='#0a0a0a', grid='#c3c9cf', hdr='#eef1f3', hdr_fg='#3c4043',
                  band='#00ccff', border='#b8bec4', alt='#ffeeb8', bord='#000000', scroll='#a9b0b6'),
    'dark':  dict(bg='#1b1d1c', fg='#ffffff', grid='#4a4f4c', hdr='#2a2d2c', hdr_fg='#d0d4d2',
                  band='#0b7a99', border='#4a4f4c', alt='#54491f', bord='#d8d8d8', scroll='#6a706c'),
}
HDR_H = 32       # chiều cao hàng chữ cái cột (A, B, C...)
RN_W = 54        # độ rộng cột số dòng
MIN_ROW = 36     # chiều cao tối thiểu mỗi dòng


def excel_css(dark):
    t = THEMES['dark' if dark else 'light']
    return f"""<style>
:root{{--bg:{t['bg']};--fg:{t['fg']};--grid:{t['grid']};--hdr:{t['hdr']};--hdr-fg:{t['hdr_fg']};
      --band:{t['band']};--border:{t['border']};--alt:{t['alt']};--bord:{t['bord']};
      --green:{GREEN};--green-d:{GREEN_DARK};--frz:#8a9096}}
*{{box-sizing:border-box}}
html,body{{height:100%}}
body{{margin:0;background:var(--bg);color:var(--fg);display:flex;flex-direction:column;
     font:600 15px 'Segoe UI','Yu Gothic UI','Meiryo',Calibri,Arial,sans-serif}}
.fxbar{{height:46px;flex:none;display:flex;align-items:center;gap:12px;padding:0 14px;background:var(--hdr);
       border:1px solid var(--border);border-bottom:none;border-radius:12px 12px 0 0;font-size:16px;font-weight:700}}
.nb{{min-width:130px;height:32px;display:flex;align-items:center;justify-content:space-between;gap:8px;
    padding:0 12px;background:var(--bg);border:1px solid var(--border);border-radius:8px}}
.nb i{{font-style:normal;color:var(--hdr-fg);font-size:10px}}
.fxi{{color:var(--green);font:italic 700 18px serif}}
#fx{{flex:1;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}}

.panes{{flex:1;min-height:0;display:flex;border:1px solid var(--border);background:var(--bg)}}
.pane{{overflow-x:scroll;position:relative;background:var(--bg)}}
#pl{{width:{LEFT_PERCENT}%;flex:none;overflow-y:hidden}}
#pr{{flex:1;min-width:0;overflow-y:scroll}}
.split{{flex:none;width:12px;cursor:col-resize;background:var(--hdr);border-left:1px solid var(--border);
       border-right:1px solid var(--border);display:flex;align-items:center;justify-content:center;
       user-select:none;transition:background .15s}}
.split::after{{content:'';width:4px;height:56px;border-radius:2px;background:var(--hdr-fg);opacity:.55}}
.split:hover,.split.on{{background:var(--green)}}
.split:hover::after,.split.on::after{{background:#fff;opacity:1}}

table{{border-collapse:separate;border-spacing:0;table-layout:fixed}}
td,th{{border-right:1px solid var(--grid);border-bottom:1px solid var(--grid);padding:2px 8px;overflow:hidden;
      background:var(--bg);color:var(--fg);vertical-align:bottom;font-weight:600}}
td.hd{{background:var(--band)}}
td.alt{{background:var(--alt);border-bottom:2px solid var(--bord)}}
td.sel{{outline:3px solid var(--green);outline-offset:-3px}}
td.fr{{position:sticky;z-index:3;font-weight:700}}
td.fb{{box-shadow:0 3px 0 var(--frz)}}
th{{background:var(--hdr);color:var(--hdr-fg);font-size:14px;font-weight:700;text-align:center}}
th.fr{{position:sticky;z-index:5}}
th.rn{{position:sticky;left:0;z-index:4;width:{RN_W}px}}
th.rn.fr{{z-index:6}}
th.fb{{box-shadow:0 3px 0 var(--frz)}}

.tab{{height:46px;flex:none;background:var(--hdr);border:1px solid var(--border);border-top:none;
     display:flex;align-items:center;gap:8px;padding-left:14px;border-radius:0 0 12px 12px}}
.tab .ic{{width:30px;text-align:center;color:var(--hdr-fg);font-size:18px}}
.tab .t{{background:var(--bg);color:var(--green);font-weight:700;font-size:15px;padding:11px 22px;
        border-bottom:4px solid var(--green);border-radius:8px 8px 0 0}}
::-webkit-scrollbar{{width:14px;height:14px}}
::-webkit-scrollbar-thumb{{background:{t['scroll']};border-radius:8px;border:3px solid var(--hdr)}}
::-webkit-scrollbar-thumb:hover{{background:var(--green)}}
::-webkit-scrollbar-track{{background:var(--hdr)}}
</style>"""


SCRIPT = """<script>
const FR=__FR__;
const pl=document.getElementById('pl'),pr=document.getElementById('pr'),sp=document.getElementById('sp'),
      panes=document.getElementById('panes'),nb=document.getElementById('nb'),fx=document.getElementById('fx');

function layout(){
  const L=[...document.querySelectorAll('#tl tr')],R=[...document.querySelectorAll('#tr tr')];
  L.concat(R).forEach(t=>t.style.height=t.dataset.h+'px');
  const hs=L.map((t,i)=>Math.max(t.offsetHeight,R[i]?R[i].offsetHeight:0));
  L.forEach((t,i)=>{t.style.height=hs[i]+'px';if(R[i])R[i].style.height=hs[i]+'px';});
  let top=0;
  for(let i=0;i<=FR;i++){
    [L[i],R[i]].forEach(tr=>{if(tr)tr.querySelectorAll('.fr').forEach(x=>x.style.top=top+'px');});
    top+=hs[i];
  }
}
let raf=0;function relayout(){cancelAnimationFrame(raf);raf=requestAnimationFrame(layout);}
function fit(){
  try{const f=window.frameElement;if(f){const h=Math.max(560,window.parent.innerHeight-110);
    if(Math.abs(f.offsetHeight-h)>2){f.style.height=h+'px';f.height=h;}}}catch(e){}
}
pr.addEventListener('scroll',()=>{pl.scrollTop=pr.scrollTop;});
pl.addEventListener('wheel',e=>{                       // lăn chuột ở khung trái -> cuộn dọc cả 2 khung
  if(e.shiftKey){pl.scrollLeft+=e.deltaY;}else{pr.scrollTop+=e.deltaY;}
  e.preventDefault();},{passive:false});

let drag=false;                                         // kéo thanh giữa để đổi tỉ lệ trái/phải
sp.addEventListener('mousedown',e=>{drag=true;sp.classList.add('on');document.body.style.userSelect='none';e.preventDefault();});
document.addEventListener('mousemove',e=>{if(!drag)return;const r=panes.getBoundingClientRect();
  pl.style.width=Math.min(Math.max(e.clientX-r.left,180),r.width-260)+'px';});
document.addEventListener('mouseup',()=>{if(!drag)return;drag=false;sp.classList.remove('on');
  document.body.style.userSelect='';layout();});

function pick(td){document.querySelectorAll('td.sel').forEach(x=>x.classList.remove('sel'));
  td.classList.add('sel');nb.textContent=td.dataset.ref;fx.textContent=td.innerText;}
panes.addEventListener('click',e=>{const td=e.target.closest('td');if(td)pick(td);});

window.addEventListener('message',e=>{                  // nút "lên đầu trang" ở góc phải dưới
  if(e.data==='kd-scroll-top'){pr.scrollTo({top:0,behavior:'smooth'});pl.scrollTo({top:0,behavior:'smooth'});}});
window.addEventListener('resize',()=>{fit();relayout();});
if(document.fonts)document.fonts.ready.then(relayout);
fit();layout();
const first=document.querySelector('td[data-ref]');if(first)pick(first);
</script>"""


def sheet_to_html(ws, last_row, min_rows=20):
    """Vẽ sheet thành 2 khung (trái = cột A-F, phải = G...), ô gộp, viền, màu, dòng lẻ tô màu."""
    pal = theme_palette(ws.parent)
    n_rows, n_cols = max(last_row + 5, min_rows), ws.max_column
    last_val = last_row

    widths = {}
    for cd in ws.column_dimensions.values():
        for c in range(cd.min or 1, min(cd.max or cd.min or 1, n_cols) + 1):
            widths[c] = cd.width
    px = {c: int((widths.get(c) or 8.43) * 9 + 8) for c in range(1, n_cols + 1)}
    row_h = {}
    for r in range(1, n_rows + 1):
        h = ws.row_dimensions[r].height
        row_h[r] = max(MIN_ROW, round(h * 1.333)) if h else MIN_ROW

    def render(c_from, c_to, with_rn, tid):
        cols_range = range(c_from, c_to + 1)
        # ô gộp, cắt theo khung (ô gộp vắt qua ranh giới F|G được tách làm 2)
        span, skip, virt = {}, set(), {}
        for m in ws.merged_cells.ranges:
            if m.max_col < c_from or m.min_col > c_to:
                continue
            a, e = max(m.min_col, c_from), min(m.max_col, c_to)
            span[(m.min_row, a)] = (m.max_row - m.min_row + 1, e - a + 1)
            if m.min_col < c_from:
                virt[(m.min_row, a)] = (m.min_row, m.min_col)
            skip |= {(rr, cc) for rr in range(m.min_row, m.max_row + 1) for cc in range(a, e + 1)} - {(m.min_row, a)}

        total = (RN_W if with_rn else 0) + sum(px[c] for c in cols_range)
        cols = ('<col style="width:%dpx">' % RN_W if with_rn else '') + ''.join(
            f'<col style="width:{px[c]}px">' for c in cols_range)
        rows = []
        # hàng chữ cái cột
        ths = ('<th class="rn fr corner"></th>' if with_rn else '') + ''.join(
            f'<th class="fr">{get_column_letter(c)}</th>' for c in cols_range)
        rows.append(f'<tr data-h="{HDR_H}" style="height:{HDR_H}px">{ths}</tr>')

        for r in range(1, n_rows + 1):
            fbot = ' fb' if r == FREEZE_ROWS else ''
            frz = ' fr' if r <= FREEZE_ROWS else ''
            tds = [f'<th class="rn{frz}{fbot}">{r}</th>'] if with_rn else []
            is_alt = r >= FIRST_DATA_ROW and (r - FIRST_DATA_ROW) % 2 == 1 and r <= last_val
            for c in cols_range:
                if (r, c) in skip:
                    continue
                rs, cs = span.get((r, c), (1, 1))
                r2, c2 = r + rs - 1, c + cs - 1
                is_virtual = (r, c) in virt
                cell = ws.cell(*virt.get((r, c), (r, c)))
                al, ft, fl = cell.alignment, cell.font, cell.fill
                v = None if is_virtual else cell.value
                css, classes = [], []

                fill = None
                if fl and fl.fill_type == 'solid':
                    fill = color_hex(fl.fgColor, pal)
                    if fill is None and fl.fgColor.type == 'theme' and r <= 3:
                        classes.append('hd')
                if fill:
                    css.append(f'background:#{fill}')
                    css.append('color:' + ('#fff' if _lum(fill) < 0.45 else '#000'))
                fc = color_hex(ft.color, pal) if ft and ft.color is not None else None
                if fc and (fill or _lum(fc) > 0.3):
                    css.append(f'color:#{fc}')
                if ft and ft.bold:
                    css.append('font-weight:800')
                if is_alt:
                    classes.append('alt')
                if r2 <= FREEZE_ROWS:
                    classes.append('fr')
                    if r2 == FREEZE_ROWS:
                        classes.append('fb')

                # viền từ file xlsx (vẽ 1 lần cho mỗi đường biên)
                pos = ws.cell(row=r, column=c)
                right = ws.cell(row=r, column=c2).border.right
                bottom = ws.cell(row=r2, column=c).border.bottom
                left = pos.border.left if (c == c_from or not ws.cell(row=r, column=c - 1).border.right.style) else None
                top = pos.border.top if (r == 1 or not ws.cell(row=r - 1, column=c).border.bottom.style) else None
                for side, bd in (('right', right), ('bottom', bottom), ('left', left), ('top', top)):
                    s = border_css(side, bd, pal)
                    if s:
                        css.append(s)

                ha = al.horizontal if al.horizontal not in (None, 'general') else \
                    ('right' if isinstance(v, (int, float, datetime)) else 'left')
                css.append(f'text-align:{ha}')
                css.append('vertical-align:' + {'center': 'middle', 'top': 'top'}.get(al.vertical, 'bottom'))
                css.append('white-space:pre-wrap' if al.wrap_text else 'white-space:nowrap')
                if isinstance(v, datetime):
                    v = f'{v.month}/{v.day}'
                text = '' if v is None else escape(str(v)).replace('\n', '<br>')
                src_r, src_c = virt.get((r, c), (r, c))
                ref = f'{get_column_letter(src_c)}{src_r}'
                if is_virtual:
                    for m in ws.merged_cells.ranges:
                        if (m.min_row, m.min_col) == (src_r, src_c):
                            ref = m.coord
                elif rs > 1 or cs > 1:
                    ref = f'{get_column_letter(c)}{r}:{get_column_letter(c2)}{r2}'
                attr = (f' rowspan="{rs}"' if rs > 1 else '') + (f' colspan="{cs}"' if cs > 1 else '')
                cls = f' class="{" ".join(classes)}"' if classes else ''
                tds.append(f'<td{cls}{attr} data-ref="{ref}" style="{";".join(css)}">{text}</td>')
            rows.append(f'<tr data-h="{row_h[r]}" style="height:{row_h[r]}px">{"".join(tds)}</tr>')
        return (f'<table id="{tid}" style="width:{total}px"><colgroup>{cols}</colgroup>'
                f'<tbody>{"".join(rows)}</tbody></table>')

    left_to = min(FREEZE_COLS, n_cols)
    left_html = render(1, left_to, True, 'tl')
    right_html = render(left_to + 1, n_cols, False, 'tr') if n_cols > left_to else '<table id="tr"></table>'

    return ('<div class="fxbar"><div class="nb"><span id="nb">A1</span><i>▼</i></div>'
              '<div class="fxi">fx</div><div id="fx"></div></div>'
            + f'<div class="panes" id="panes"><div class="pane" id="pl">{left_html}</div>'
              f'<div class="split" id="sp" title="左右の比率をドラッグして調整"></div>'
              f'<div class="pane" id="pr">{right_html}</div></div>'
            + f'<div class="tab"><span class="ic">＋</span><span class="ic">≡</span>'
              f'<div class="t">{escape(ws.title)}</div></div>'
            + SCRIPT.replace('__FR__', str(FREEZE_ROWS)))


# ---------------------------- Xem nội dung PDF + nút lên đầu trang ----------------------------
def pdf_viewer_html(data):
    b64 = base64.b64encode(data).decode()
    return ('<body style="margin:0"><iframe id="f" style="width:100%;height:100vh;border:0"></iframe><script>'
            'const b=atob("__B64__"),a=new Uint8Array(b.length);'
            'for(let i=0;i<b.length;i++)a[i]=b.charCodeAt(i);'
            'document.getElementById("f").src=URL.createObjectURL(new Blob([a],{type:"application/pdf"}));'
            '</script></body>').replace('__B64__', b64)


def open_pdf(f):
    """Click tên file -> mở cửa sổ hiển thị nội dung PDF."""
    dialog = getattr(st, 'dialog', None) or getattr(st, 'experimental_dialog', None)
    data = f.getvalue()
    if dialog is None:                                    # Streamlit cũ: hiện trong expander
        with st.expander(f'📄 {f.name}', expanded=True):
            components.html(pdf_viewer_html(data), height=650)
        return
    try:
        deco = dialog(f.name, width='large')
    except TypeError:
        deco = dialog(f.name)

    @deco
    def _show():
        components.html(pdf_viewer_html(data), height=650)
    _show()


TOP_BTN = """<script>
(function(){
  const d=window.parent.document;
  if(d.getElementById('kd-top-btn'))return;
  const b=d.createElement('button');
  b.id='kd-top-btn';b.title='先頭へ';b.textContent='↑';
  b.style.cssText='position:fixed;right:28px;bottom:28px;width:54px;height:54px;border-radius:50%;border:none;'
    +'background:#217346;color:#fff;font-size:26px;font-weight:700;cursor:pointer;z-index:999999;'
    +'box-shadow:0 6px 16px rgba(0,0,0,.35);transition:transform .15s,background .15s';
  b.onmouseenter=()=>{b.style.background='#185c37';b.style.transform='translateY(-3px)';};
  b.onmouseleave=()=>{b.style.background='#217346';b.style.transform='none';};
  b.onclick=()=>{
    const m=d.querySelector('[data-testid="stMain"]')||d.querySelector('section.main');
    if(m)m.scrollTo({top:0,behavior:'smooth'});
    d.querySelectorAll('iframe').forEach(f=>{try{f.contentWindow.postMessage('kd-scroll-top','*');}catch(e){}});
  };
  d.body.appendChild(b);
})();
</script>"""


# ---------------------------- Giao diện Streamlit ----------------------------
def app_css(dark):
    if dark:
        bg, fg, card, border, muted = '#121715', '#f2f6f3', '#1c2420', '#2f3b34', '#a9bbb0'
    else:
        bg, fg, card, border, muted = '#f4f7f5', '#13211a', '#ffffff', '#d5e0d9', '#5d6f64'
    return f"""<style>
.stApp{{background:{bg}}}
.stApp, .stApp p, .stApp label, .stApp h1, .stApp h2, .stApp h3, .stApp span, .stApp small{{color:{fg}}}
header[data-testid="stHeader"]{{background:transparent}}
.block-container{{padding:1.5rem 2rem 3rem;max-width:100%}}
h1{{font-weight:700;letter-spacing:-0.5px}}

/* khung kéo thả file */
[data-testid="stFileUploaderDropzone"]{{background:{card};border:2px dashed {GREEN};border-radius:14px;padding:1.2rem}}
[data-testid="stFileUploaderDropzone"] small{{color:{muted}}}
[data-testid="stFileUploaderFile"]{{background:{card};border:1px solid {border};border-radius:10px;padding:.2rem .6rem}}

/* nút xanh */
.stButton > button, .stDownloadButton > button,
[data-testid="stFileUploaderDropzone"] button{{
  background:{GREEN};color:#fff !important;border:none;border-radius:10px;
  font-weight:600;padding:.55rem 1.3rem;box-shadow:0 2px 6px rgba(33,115,70,.25);transition:all .15s}}
.stButton > button *, .stDownloadButton > button *,
[data-testid="stFileUploaderDropzone"] button *{{color:#fff !important}}
.stButton > button:hover:not(:disabled), .stDownloadButton > button:hover:not(:disabled),
[data-testid="stFileUploaderDropzone"] button:hover{{background:{GREEN_DARK};transform:translateY(-1px);
  box-shadow:0 4px 10px rgba(33,115,70,.35)}}
.stButton > button:disabled{{background:{border};color:{muted} !important;box-shadow:none;cursor:not-allowed}}
.stButton > button:disabled *{{color:{muted} !important}}

/* nút xem file: viền xanh, nền trắng */
[class*="st-key-view_"] button{{background:{card};border:2px solid {GREEN};box-shadow:none}}
[class*="st-key-view_"] button, [class*="st-key-view_"] button *{{color:{GREEN} !important}}
[class*="st-key-view_"] button:hover:not(:disabled){{background:{GREEN};box-shadow:none}}
[class*="st-key-view_"] button:hover:not(:disabled) *{{color:#fff !important}}

/* progress + toggle */
.stProgress > div > div > div > div{{background:{GREEN}}}
[data-baseweb="checkbox"] > div:first-child{{background-color:{GREEN} !important}}
</style>"""


def build(items, dicts):
    """items: [(số, tên file)]. Trả về bytes xlsx và HTML của sheet."""
    wb = load_workbook(find_template())
    ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb.active
    keys = header_keys(ws)

    r = FIRST_DATA_ROW
    for n, _ in sorted(items):
        dic = dicts[DIC_BY_NUMBER[n]]
        even = [dic.get(k) for k in keys]                 # dòng chẵn: giá trị từ dic
        odd = [ODD_DEFAULTS.get(k) for k in keys]         # dòng lẻ: mặc định
        for vals in (even, odd):
            for col, v in enumerate(vals, start=1):
                if v is not None:
                    ws.cell(row=r, column=col, value=v)
            r += 1

    last_row = r - 1
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue(), sheet_to_html(ws, last_row)


def make(items):
    """Chạy nền: tạo xlsx + HTML. Không gọi st.* trong hàm này."""
    return build(items, load_dicts(find_dic())) if items else (None, None)


def main():
    st.set_page_config(page_title='PDF → KDパンチ入力シート', page_icon='📗', layout='wide')

    c1, c2 = st.columns([6, 1])
    dark = c2.toggle('🌙 ダークモード', key='dark')
    st.markdown(app_css(dark), unsafe_allow_html=True)    # CSS theo chế độ sáng/tối
    c1.title('📗 PDF → KDパンチ入力シート')

    files = st.file_uploader('PDFファイルを1件または複数選択', type='pdf',
                             accept_multiple_files=True)

    if files:
        with st.expander('👁 ファイル名をクリックして内容を確認', expanded=False):
            cols = st.columns(min(len(files), 4))
            for i, f in enumerate(files):
                if cols[i % len(cols)].button(f'📄 {f.name}', key=f'view_{i}', use_container_width=True):
                    open_pdf(f)

    if st.button('処理を実行', type='primary', disabled=not files):
        items, errors = [], []
        for f in files:
            n = file_number(f.name)
            if n is None:
                errors.append(f'{f.name}: ファイル名に番号が見つかりません')
            elif n not in DIC_BY_NUMBER:
                errors.append(f'{f.name}: 番号 {n} に対応する辞書がありません')
            else:
                items.append((n, f.name))

        total = SECONDS_PER_FILE * len(files)
        bar = st.progress(0.0, text='処理中...')
        try:
            with ThreadPoolExecutor(max_workers=1) as ex, st.spinner('データを読み込んでいます...'):
                fut = ex.submit(make, items)
                for s in range(1, total + 1):
                    time.sleep(1)
                    bar.progress(s / total, text=f'処理中... {s}/{total}s')
                data, sheet = fut.result()
            st.session_state['result'] = (data, sheet, errors)
        except Exception as e:                            # thiếu file, sai path, sai sheet...
            st.session_state['result'] = None
            st.error(f'エラー: {e!r}')
        bar.empty()

    res = st.session_state.get('result')
    if res:
        data, sheet, errors = res
        for e in errors:
            st.warning(e)
        if data is not None:
            st.subheader('結果シート')
            components.html(excel_css(dark) + sheet, height=760)
            st.download_button('⬇ XLSXをダウンロード', data, file_name=RESULT_NAME,
                               mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    components.html(TOP_BTN, height=0)                    # nút ↑ góc phải dưới


if __name__ == '__main__':
    main()


# python -m streamlit run app.py