#!/usr/bin/env python3
"""PyQt6 시트 계산기 (Excel 함수 지원) - pip install PyQt6"""
import sys, re, math, json, csv, statistics, datetime, random, time
from pathlib import Path
from PyQt6.QtCore import Qt, QAbstractTableModel, QModelIndex, QSettings, QEvent, QTimer, QPoint
from PyQt6.QtGui import QColor, QAction, QKeySequence, QFont, QBrush, QCursor, QPixmap, QPainter, QIcon
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QTableView,
    QLineEdit, QLabel, QComboBox, QPushButton, QDialog, QFormLayout, QFontComboBox, QSpinBox,
    QColorDialog, QScrollArea, QFileDialog, QAbstractItemView, QListWidget, QInputDialog, QStyledItemDelegate, QMessageBox, QMenu, QSplitter, QToolButton)

# ───────────── 다국어 (ko, ja, en) ─────────────
LANGS = [("한국어", "ko"), ("日本語", "ja"), ("English", "en")]
LANG = "ko"
S = {
 "file": ("파일", "ファイル", "File"), "new": ("새로 만들기", "新規", "New"),
 "open": ("CSV 열기", "CSVを開く", "Open CSV"), "save": ("CSV 저장", "CSVを保存", "Save CSV"),
 "quit": ("종료", "終了", "Quit"), "theme": ("테마", "テーマ", "Theme"), "lang": ("언어", "言語", "Language"),
 "custom": ("사용자 지정...", "カスタマイズ...", "Customize..."),
 "default": ("기본 테마", "デフォルトテーマ", "Default"), "count": ("개수", "個数", "Count"),
 "sum": ("합계", "合計", "Sum"), "avg": ("평균", "平均", "Average"),
 "title": ("XcelBox", "XcelBox", "XcelBox"),
 "reset": ("기본값으로 초기화", "デフォルトにリセット", "Reset to default"),
 "apply": ("적용", "適用", "Apply"), "close": ("닫기", "閉じる", "Close"),
 "colors": ("색상", "色", "Colors"), "fonts": ("글꼴 / 크기", "フォント / サイズ", "Fonts / Size"),
 "fill": ("선택한 모든 셀에 입력", "選択セルすべてに入力", "Fill all selected cells"),
 "ph": ("값 또는 SUM(A1:A3) 처럼 입력", "値または SUM(A1:A3) のように入力", "Enter a value or e.g. SUM(A1:A3)"),
 "view": ("보기", "表示", "View"), "memo": ("메모리", "メモリ", "Memory"),
 "m_save": ("현재 시트 저장", "現在のシートを保存", "Save current sheet"), "m_load": ("불러오기", "読み込み", "Load"),
 "m_over": ("덮어쓰기", "上書き", "Overwrite"), "m_ren": ("제목 변경", "タイトル変更", "Rename"),
 "m_del": ("삭제", "削除", "Delete"), "m_ask": ("제목을 입력하세요", "タイトルを入力してください", "Enter a title"),
 "c_window_bg": ("창 배경", "ウィンドウ背景", "Window background"), "c_window_text": ("창 글자", "ウィンドウ文字", "Window text"),
 "c_cell_bg": ("셀 배경", "セル背景", "Cell background"), "c_cell_text": ("셀 글자", "セル文字", "Cell text"),
 "c_grid": ("격자선", "グリッド線", "Grid lines"), "c_header_bg": ("헤더 배경", "ヘッダー背景", "Header background"),
 "c_header_text": ("헤더 글자", "ヘッダー文字", "Header text"), "c_select_bg": ("선택 배경", "選択背景", "Selection background"),
 "c_select_text": ("선택 글자", "選択文字", "Selection text"), "c_input_bg": ("입력창 배경", "入力欄背景", "Input background"),
 "c_input_text": ("입력창 글자", "入力欄文字", "Input text"), "c_button_bg": ("버튼 배경", "ボタン背景", "Button background"),
 "c_button_text": ("버튼 글자", "ボタン文字", "Button text"), "c_status_bg": ("상태바 배경", "ステータスバー背景", "Status bar background"),
 "c_status_text": ("상태바 글자", "ステータスバー文字", "Status bar text"),
 "f_ui": ("전체 UI", "全体UI", "General UI"), "f_cell": ("셀", "セル", "Cells"), "f_header": ("헤더", "ヘッダー", "Headers"),
 "f_formula": ("수식 입력줄", "数式バー", "Formula bar"), "f_status": ("상태바", "ステータスバー", "Status bar"),
 "settings": ("설정", "設定", "Settings"), "cancel": ("취소", "キャンセル", "Cancel"),
 "m_clear": ("전체 삭제", "すべて削除", "Delete all"),
 "m_clear_ask": ("저장된 메모리를 모두 삭제할까요?\n이 작업은 되돌릴 수 없습니다.",
                 "保存されたメモリをすべて削除しますか?\nこの操作は元に戻せません。",
                 "Delete all saved memories?\nThis cannot be undone."),
 "current_work": ("현재 작업 (Current Work)", "現在の作業 (Current Work)", "Current Work"),
 "cell_bg_btn": ("셀 배경색", "セル背景色", "Cell BG Color"),
 "cell_fg_btn": ("셀 글자색", "セル文字色", "Cell Text Color"),
 "cell_color_reset": ("셀 색상 초기화", "セル色リセット", "Reset Cell Colors"),
 "export_cfg": ("설정 내보내기", "設定エクスポート", "Export Config"),
 "import_cfg": ("설정 불러오기", "設定インポート", "Import Config"),
}
def tr(k): return S[k][[c for _, c in LANGS].index(LANG)]

# ───────────── 테마 ─────────────
THEMES = {
 "default": ("#f0f0f0", "#000000", "#ffffff", "#e1e1e1", "#3399ff", "#d0d0d0"),
 "Light": ("#ffffff", "#202124", "#ffffff", "#f1f3f4", "#1a73e8", "#dadce0"),
 "Dark Gray": ("#2b2b2b", "#e0e0e0", "#333333", "#3c3c3c", "#5a7ab5", "#4a4a4a"),
 "Latte": ("#eff1f5", "#4c4f69", "#e6e9ef", "#ccd0da", "#1e66f5", "#bcc0cc"),
 "Tokyo Night": ("#1a1b26", "#c0caf5", "#16161e", "#24283b", "#7aa2f7", "#292e42"),
 "Frappé": ("#303446", "#c6d0f5", "#292c3c", "#414559", "#8caaee", "#51576d"),
 "Macchiato": ("#24273a", "#cad3f5", "#1e2030", "#363a4f", "#8aadf4", "#494d64"),
 "Pastel Mint": ("#e8f8f1", "#2f4f43", "#f6fffb", "#c8efdf", "#7fd6b4", "#b7e4d2"),
 "Pastel Pink": ("#fff0f5", "#5c3a47", "#fffafc", "#ffd9e6", "#f7a8c4", "#f3c6d6"),
 "Retro": ("#2d1b00", "#ffb000", "#1a1000", "#3d2800", "#ff8c00", "#5a3c00"),
 "Cyberpunk": ("#0d0221", "#00f0ff", "#12032e", "#241047", "#ff00a0", "#3a1a6e"),
 "Classic Windows": ("#c0c0c0", "#000000", "#ffffff", "#c0c0c0", "#000080", "#808080"),
 "Sage Green": ("#e9efe6", "#34423a", "#f4f8f1", "#cfdccb", "#87a96b", "#b8c8b2"),
 "Dark Academia": ("#2a211b", "#d9c8a9", "#241c17", "#3a2d24", "#8b5a2b", "#4d3c30"),
 "Japanese Cherry Blossom": ("#fff5f7", "#5a3a44", "#ffffff", "#fde0e8", "#f4a1b8", "#f2cdd6"),
 "Nordic Forest": ("#1f2d27", "#d8e2dc", "#18241f", "#2c4037", "#4f8a6b", "#3b5448"),
 "Mountain Lake": ("#e6f0f3", "#243b4a", "#f5fafb", "#c5dde4", "#3b8ea5", "#b0ccd5"),
 "Ocean": ("#0a2540", "#d6ecff", "#07203a", "#12395f", "#1ca3ec", "#1d4a75"),
 "Blue Sky": ("#e3f2fd", "#0d3c61", "#ffffff", "#bbdefb", "#42a5f5", "#b3d4ef"),
 "Deep Blue": ("#0b1230", "#dbe4ff", "#08102a", "#16224f", "#3d5afe", "#26346e"),
 "Discord": ("#313338", "#dbdee1", "#2b2d31", "#1e1f22", "#5865f2", "#3f4147"),
 "Pokémon": ("#fff8dc", "#222222", "#ffffff", "#ffcb05", "#3b4cca", "#e6c200"),
 "Ark": ("#10161a", "#c7e6f0", "#0b1013", "#1b2830", "#e07b00", "#2a3b45"),
 "Atom One Dark": ("#282c34", "#abb2bf", "#21252b", "#2c313a", "#61afef", "#3e4451"),
 "Black Mint": ("#0a0a0a", "#c9f5e1", "#000000", "#141414", "#3eb489", "#1f2f28"),
}
COLOR_KEYS = ["window_bg", "window_text", "cell_bg", "cell_text", "grid", "header_bg", "header_text",
              "select_bg", "select_text", "input_bg", "input_text", "button_bg", "button_text", "status_bg", "status_text"]
FONT_KEYS = ["ui", "cell", "header", "formula", "status"]
DEF_FONTS = {"ui": ["맑은 고딕", 10], "cell": ["맑은 고딕", 11], "header": ["맑은 고딕", 10], "formula": ["맑은 고딕", 11], "status": ["맑은 고딕", 10]}

def contrast(c):
    q = QColor(c)
    return "#000000" if 0.299 * q.red() + 0.587 * q.green() + 0.114 * q.blue() > 150 else "#ffffff"

def make_colors(t):
    bg, tx, cell, hd, ac, gr = THEMES[t]
    return dict(window_bg=bg, window_text=tx, cell_bg=cell, cell_text=tx, grid=gr, header_bg=hd, header_text=tx,
                select_bg=ac, select_text=contrast(ac), input_bg=cell, input_text=tx, button_bg=hd,
                button_text=tx, status_bg=hd, status_text=tx)

MEMO_FILE = Path.home() / ".pysheet_calc_memos.json"

QSS = """
QWidget{background:%(window_bg)s;color:%(window_text)s;%(f_ui)s}
QMenuBar::item:selected,QMenu::item:selected{background:%(select_bg)s;color:%(select_text)s}
QMenu{border:1px solid %(grid)s}
QTableView{background:%(cell_bg)s;color:%(cell_text)s;gridline-color:%(grid)s;selection-background-color:%(select_bg)s;selection-color:%(select_text)s;border:1px solid %(grid)s;%(f_cell)s}
QHeaderView{%(f_header)s}
QHeaderView::section{background:%(header_bg)s;color:%(header_text)s;border:1px solid %(grid)s;padding:2px 6px;%(f_header)s}
QTableCornerButton::section{background:%(header_bg)s;border:1px solid %(grid)s}
QLineEdit,QSpinBox,QComboBox,QFontComboBox{background:%(input_bg)s;color:%(input_text)s;border:1px solid %(grid)s;padding:2px 4px}
QLineEdit#formula{%(f_formula)s}
QComboBox QAbstractItemView{background:%(input_bg)s;color:%(input_text)s;selection-background-color:%(select_bg)s;selection-color:%(select_text)s}
QPushButton,QToolButton{background:%(button_bg)s;color:%(button_text)s;border:1px solid %(grid)s;padding:4px 12px}
QPushButton:hover,QToolButton:hover{border-color:%(select_bg)s}
QPushButton::menu-button,QToolButton::menu-button{border:none;width:16px}
QLabel#status{background:%(status_bg)s;color:%(status_text)s;padding:4px 8px;%(f_status)s}
QScrollBar:vertical{background:%(window_bg)s;width:12px}
QScrollBar:horizontal{background:%(window_bg)s;height:12px}
QScrollBar::handle{background:%(grid)s;border-radius:5px;min-height:24px;min-width:24px}
QScrollBar::add-line,QScrollBar::sub-line{width:0;height:0}
QScrollArea{border:none}
QListWidget{background:%(cell_bg)s;color:%(cell_text)s;border:1px solid %(grid)s;%(f_cell)s}
QListWidget::item{padding:4px}
QListWidget::item:selected{background:%(select_bg)s;color:%(select_text)s}
"""

def make_color_icon(color_hex):
    pixmap = QPixmap(14, 14)
    pixmap.fill(Qt.GlobalColor.transparent)
    painter = QPainter(pixmap)
    painter.setBrush(QBrush(QColor(color_hex if color_hex else "#ffffff")))
    painter.setPen(QColor("#666666"))
    painter.drawRect(1, 1, 11, 11)
    painter.end()
    return QIcon(pixmap)

def toggle_f4_ref(text, pos):
    m = re.finditer(r"(\$?[A-Z]{1,3})(\$?\d+)", text)
    for match in m:
        start, end = match.span()
        if start <= pos <= end:
            col_part, row_part = match.group(1), match.group(2)
            has_c_abs = col_part.startswith("$")
            has_r_abs = row_part.startswith("$")
            c_name = col_part.lstrip("$")
            r_name = row_part.lstrip("$")
            
            if not has_c_abs and not has_r_abs:
                new_ref = f"${c_name}${r_name}"
            elif has_c_abs and has_r_abs:
                new_ref = f"{c_name}${r_name}"
            elif not has_c_abs and has_r_abs:
                new_ref = f"${c_name}{r_name}"
            else:
                new_ref = f"{c_name}{r_name}"
            
            new_text = text[:start] + new_ref + text[end:]
            return new_text, start + len(new_ref)
    return text, pos

def col_name(n):
    s = ""; n += 1
    while n: n, r = divmod(n - 1, 26); s = chr(65 + r) + s
    return s

def col_idx(s):
    c = 0
    for ch in s: c = c * 26 + ord(ch) - 64
    return c - 1

AUTO = re.compile(r"\s*(?:(?:\$?[A-Za-z]{1,3}\$?\d+|\d+\.?\d*|\.\d+|[-+*/^()])\s*)+")

def is_formula(t):
    if not t: return False
    if t.startswith("="): return True
    stripped = t.strip().upper()
    if any(stripped.startswith(k + "(") for k in F.keys()): return True
    if AUTO.fullmatch(t) and re.search(r"[-+*/^A-Za-z]", t): return True
    return False

def shift_formula(formula, dr, dc):
    if not is_formula(formula):
        return formula

    has_eq = formula.startswith("=")
    f_body = formula[1:] if has_eq else formula

    def replace_ref(m):
        c_abs, c_str, r_abs, r_str = m.group(1), m.group(2), m.group(3), m.group(4)
        c_val = col_idx(c_str)
        r_val = int(r_str) - 1

        if not c_abs:
            c_val = max(0, c_val + dc)
        if not r_abs:
            r_val = max(0, r_val + dr)

        new_c = f"${c_str}" if c_abs else col_name(c_val)
        new_r = f"${r_str}" if r_abs else str(r_val + 1)
        return new_c + new_r

    pattern = r"(\$?)([A-Z]{1,3})(\$?)([1-9]\d*)"
    new_body = re.sub(pattern, replace_ref, f_body)
    return "=" + new_body if has_eq else new_body

class Err(Exception):
    def __init__(s, code): super().__init__(code); s.code = code

def flat(a):
    for x in a:
        if isinstance(x, list): yield from x
        else: yield x
def isnum(x): return isinstance(x, (int, float)) and not isinstance(x, bool)
def nums(a): return [x for x in flat(a) if isnum(x)]
def num(x):
    if x in (None, ""): return 0
    try: return float(x)
    except (TypeError, ValueError): raise Err("#VALUE!")

def fmt(v):
    if v is None: return ""
    if isinstance(v, bool): return "TRUE" if v else "FALSE"
    if isinstance(v, Dt): return v.show()
    if isinstance(v, float): return str(int(v)) if v == int(v) and abs(v) < 1e15 else f"{v:.10g}"
    return str(v)

def crit(c):
    op, v = "=", c
    if isinstance(c, str):
        m = re.match(r"(<=|>=|<>|<|>|=)?(.*)", c); op, v = m.group(1) or "=", m.group(2)
        try: v = float(v)
        except ValueError: pass
    f = {"=": lambda a, b: a == b, "<>": lambda a, b: a != b, "<": lambda a, b: a < b,
         ">": lambda a, b: a > b, "<=": lambda a, b: a <= b, ">=": lambda a, b: a >= b}[op]
    def t(x):
        try: return x is not None and f(x, v)
        except TypeError: return False
    return t

def ROUND(x, n=0):
    m = 10 ** int(n); x = num(x)
    return math.copysign(math.floor(abs(x) * m + 0.5) / m, x)

F = dict(
    SUM=lambda *a: sum(nums(a)), AVERAGE=lambda *a: sum(nums(a)) / len(nums(a)),
    COUNT=lambda *a: len(nums(a)), COUNTA=lambda *a: sum(1 for x in flat(a) if x not in (None, "")),
    MIN=lambda *a: min(nums(a) or [0]), MAX=lambda *a: max(nums(a) or [0]),
    PRODUCT=lambda *a: math.prod(nums(a)), MEDIAN=lambda *a: statistics.median(nums(a)),
    IF=lambda c, a=True, b=False: a if c else b, ROUND=ROUND,
    ABS=lambda x: abs(num(x)), SQRT=lambda x: math.sqrt(num(x)), POWER=lambda x, y: num(x) ** num(y),
    MOD=lambda x, y: num(x) % num(y), INT=lambda x: math.floor(num(x)),
    LOG=lambda x, b=10: math.log(num(x), num(b)), LN=lambda x: math.log(num(x)), EXP=lambda x: math.exp(num(x)),
    PI=lambda: math.pi, SIN=lambda x: math.sin(num(x)), COS=lambda x: math.cos(num(x)), TAN=lambda x: math.tan(num(x)),
    AND=lambda *a: all(flat(a)), OR=lambda *a: any(flat(a)), NOT=lambda x: not x,
    CONCAT=lambda *a: "".join(fmt(x) for x in flat(a)), CONCATENATE=lambda *a: "".join(fmt(x) for x in flat(a)),
    LEN=lambda s: len(fmt(s)), UPPER=lambda s: fmt(s).upper(), LOWER=lambda s: fmt(s).lower(),
    LEFT=lambda s, n=1: fmt(s)[:int(n)], RIGHT=lambda s, n=1: fmt(s)[-int(n):] if int(n) else "",
    SUMIF=lambda r, c, s=None: sum(y for x, y in zip(r, s if s is not None else r) if crit(c)(x) and isnum(y)),
    COUNTIF=lambda r, c: sum(1 for x in r if crit(c)(x)),
    TRUE=True, FALSE=False,
)

EPOCH = datetime.datetime(1899, 12, 30)
class Dt(float):
    def show(s):
        d = EPOCH + datetime.timedelta(days=float(s))
        return d.strftime("%Y-%m-%d" if float(s) == int(s) else "%Y-%m-%d %H:%M")
class Rng(list): nr = nc = 1

def todate(x): return EPOCH + datetime.timedelta(days=num(x))
def serial(d): return (d - EPOCH).total_seconds() / 86400
def DATE(y, m, d):
    y, m = int(y) + (int(m) - 1) // 12, (int(m) - 1) % 12 + 1
    return Dt(serial(datetime.datetime(y, m, 1)) + int(d) - 1)
def _rd(x, n, f): m = 10 ** int(n); x = num(x); return math.copysign(f(abs(x) * m) / m, x)
def _var(a, p):
    x = nums(a); m = sum(x) / len(x); return sum((i - m) ** 2 for i in x) / (len(x) if p else len(x) - 1)
def _idx(a):
    ts = [(r, crit(c)) for r, c in zip(a[::2], a[1::2])]
    return [i for i in range(len(a[0])) if all(t(r[i]) for r, t in ts)]
def _find(v, keys, approx):
    if not approx:
        for i, k in enumerate(keys):
            if k == v or (isinstance(k, str) and isinstance(v, str) and k.lower() == v.lower()): return i
        raise Err("#N/A")
    idx = None
    for i, k in enumerate(keys):
        try:
            if k is not None and k <= v: idx = i
        except TypeError: pass
    if idx is None: raise Err("#N/A")
    return idx
def VLOOKUP(v, rng, col, approx=True):
    if not 1 <= int(col) <= rng.nc: raise Err("#REF!")
    i = _find(v, [rng[r * rng.nc] for r in range(rng.nr)], bool(approx)); return rng[i * rng.nc + int(col) - 1]
def HLOOKUP(v, rng, row, approx=True):
    if not 1 <= int(row) <= rng.nr: raise Err("#REF!")
    i = _find(v, rng[:rng.nc], bool(approx)); return rng[(int(row) - 1) * rng.nc + i]
def INDEX(rng, r, c=None):
    r = int(r)
    if c is None: r, c = (1, r) if rng.nr == 1 else (r, 1)
    if not (1 <= r <= rng.nr and 1 <= int(c) <= rng.nc): raise Err("#REF!")
    return rng[(r - 1) * rng.nc + int(c) - 1]
def IFS(*a):
    for c, v in zip(a[::2], a[1::2]):
        if c: return v
    raise Err("#N/A")
def IFERROR(a, b):
    try: return a()
    except Exception: return b()
def ISERROR(a):
    try: a(); return False
    except Exception: return True
def TEXT(v, f):
    v = num(v); pct = "%" in f; v *= 100 if pct else 1
    dec = len(f.split(".")[1].rstrip("%")) if "." in f else 0
    return (f"{v:,.{dec}f}" if "," in f else f"{v:.{dec}f}") + ("%" if pct else "")

F.update(
    IF=lambda c, a=lambda: True, b=lambda: False: (a if c else b)(), IFERROR=IFERROR, ISERROR=ISERROR, IFS=IFS,
    XOR=lambda *a: sum(bool(x) for x in flat(a)) % 2 == 1, CHOOSE=lambda i, *a: a[int(i) - 1],
    ISNUMBER=lambda x: isnum(x), ISTEXT=lambda x: isinstance(x, str),
    ISEVEN=lambda x: int(num(x)) % 2 == 0, ISODD=lambda x: int(num(x)) % 2 == 1,
    SUMSQ=lambda *a: sum(x * x for x in nums(a)),
    SUMPRODUCT=lambda *a: sum(math.prod(num(x) for x in t) for t in zip(*a)),
    ROUNDUP=lambda x, n=0: _rd(x, n, math.ceil), ROUNDDOWN=lambda x, n=0: _rd(x, n, math.floor),
    TRUNC=lambda x, n=0: _rd(x, n, math.floor),
    CEILING=lambda x, s=1: math.ceil(num(x) / s) * s, FLOOR=lambda x, s=1: math.floor(num(x) / s) * s,
    SIGN=lambda x: (num(x) > 0) - (num(x) < 0), FACT=lambda x: math.factorial(int(num(x))),
    GCD=lambda *a: math.gcd(*[int(x) for x in nums(a)]), LCM=lambda *a: math.lcm(*[int(x) for x in nums(a)]),
    QUOTIENT=lambda x, y: math.trunc(num(x) / num(y)), LOG10=lambda x: math.log10(num(x)),
    ASIN=lambda x: math.asin(num(x)), ACOS=lambda x: math.acos(num(x)), ATAN=lambda x: math.atan(num(x)),
    ATAN2=lambda y, x: math.atan2(num(y), num(x)), DEGREES=lambda x: math.degrees(num(x)),
    RADIANS=lambda x: math.radians(num(x)), RAND=random.random, RANDBETWEEN=lambda a, b: random.randint(int(a), int(b)),
    COUNTBLANK=lambda r: sum(1 for x in flat([r]) if x in (None, "")),
    MODE=lambda *a: statistics.mode(nums(a)),
    LARGE=lambda r, k: sorted(nums([r]), reverse=True)[int(k) - 1], SMALL=lambda r, k: sorted(nums([r]))[int(k) - 1],
    RANK=lambda x, r, o=0: 1 + sum(1 for y in nums([r]) if (y < x if o else y > x)),
    STDEV=lambda *a: math.sqrt(_var(a, 0)), STDEVP=lambda *a: math.sqrt(_var(a, 1)),
    VAR=lambda *a: _var(a, 0), VARP=lambda *a: _var(a, 1),
    AVERAGEIF=lambda r, c, a=None: (lambda v: sum(v) / len(v))(
        [y for x, y in zip(r, r if a is None else a) if crit(c)(x) and isnum(y)]),
    SUMIFS=lambda s, *a: sum(s[i] for i in _idx(a) if isnum(s[i])), COUNTIFS=lambda *a: len(_idx(a)),
    AVERAGEIFS=lambda s, *a: (lambda v: sum(v) / len(v))([s[i] for i in _idx(a) if isnum(s[i])]),
    PROPER=lambda s: fmt(s).title(), TRIM=lambda s: " ".join(fmt(s).split()),
    MID=lambda s, i, n: fmt(s)[int(i) - 1:int(i) - 1 + int(n)],
    FIND=lambda t, s, st=1: fmt(s).index(fmt(t), int(st) - 1) + 1,
    SEARCH=lambda t, s, st=1: fmt(s).lower().index(fmt(t).lower(), int(st) - 1) + 1,
    SUBSTITUTE=lambda s, o, n, k=None: fmt(s).replace(fmt(o), fmt(n)), REPT=lambda s, n: fmt(s) * int(n),
    EXACT=lambda a, b: fmt(a) == fmt(b), VALUE=lambda s: num(s), TEXT=TEXT,
    TEXTJOIN=lambda d, ig, *a: fmt(d).join(fmt(x) for x in flat(a) if not (ig and x in (None, ""))),
    VLOOKUP=VLOOKUP, HLOOKUP=HLOOKUP, INDEX=INDEX, MATCH=lambda v, r, t=1: _find(v, list(r), bool(t)) + 1,
    ROWS=lambda r: getattr(r, "nr", 1), COLUMNS=lambda r: getattr(r, "nc", 1),
    TODAY=lambda: Dt(int(serial(datetime.datetime.now()))), NOW=lambda: Dt(serial(datetime.datetime.now())),
    DATE=DATE, YEAR=lambda x: todate(x).year, MONTH=lambda x: todate(x).month, DAY=lambda x: todate(x).day,
    WEEKDAY=lambda x: (todate(x).weekday() + 1) % 7 + 1, DAYS=lambda e, s: num(e) - num(s),
)

LZ = re.compile(r"\b(IFERROR|ISERROR|IF)\(")
def lazy(e):
    out, i = "", 0
    while True:
        m = LZ.search(e, i)
        if not m: return out + e[i:]
        j, d, q, st, parts = m.end(), 1, None, m.end(), []
        while j < len(e):
            ch = e[j]
            if q:
                if ch == q: q = None
            elif ch in "\"'": q = ch
            elif ch == "(": d += 1
            elif ch == ")":
                d -= 1
                if d == 0: break
            elif ch == "," and d == 1: parts.append(e[st:j]); st = j + 1
            j += 1
        parts.append(e[st:j])
        a = [lazy(p) for p in parts]
        a = [a[0]] + [f"lambda:({x})" for x in a[1:]] if m.group(1) == "IF" else [f"lambda:({x})" for x in a]
        out += e[i:m.start()] + m.group(1) + "(" + ",".join(a) + ")"
        i = j + 1

def ref(a):
    m = re.fullmatch(r"([A-Z]+)(\d+)", a); c = 0
    for ch in m.group(1): c = c * 26 + ord(ch) - 64
    return int(m.group(2)) - 1, c - 1

class Sheet:
    def __init__(s):
        s.raw = {}
        s.style = {}
        s.cache = {}
        s.stack = set()

    def value(s, k):
        if k in s.cache: return s.cache[k]
        if k in s.stack: raise Err("#REF!")
        s.stack.add(k)
        try: v = s.compute(s.raw.get(k, ""))
        except Err as e: v = e
        except ZeroDivisionError: v = Err("#DIV/0!")
        except NameError: v = Err("#NAME?")
        except SyntaxError: v = Err("#SYNTAX?")
        except (ValueError, OverflowError): v = Err("#NUM!")
        except Exception: v = Err("#VALUE!")
        finally: s.stack.discard(k)
        s.cache[k] = v
        return v

    def compute(s, t):
        if t == "": return None
        if t.startswith("'"): return t[1:]
        
        f_str = t[1:] if t.startswith("=") else t
        if re.fullmatch(r"[+-]?(\d+\.?\d*|\.\d+)(e[+-]?\d+)?", t.strip(), re.I): return float(t)
        
        stripped = t.strip().upper()
        is_func_call = any(stripped.startswith(k + "(") for k in F.keys())
        
        if is_func_call or t.startswith("=") or (AUTO.fullmatch(t) and re.search(r"[-+*/^A-Za-z]", t)):
            try: return s.evalf(f_str)
            except Exception:
                if is_func_call or t.startswith("="):
                    raise
                return t
        return t

    def evalf(s, f):
        def C(a):
            v = s.value(ref(a))
            if isinstance(v, Err): raise v
            return 0 if v is None else v
        def R(a, b):
            (r1, c1), (r2, c2) = ref(a), ref(b)
            if (abs(r2 - r1) + 1) * (abs(c2 - c1) + 1) > 50000: raise Err("#REF!")
            out = Rng(); out.nr = abs(r2 - r1) + 1; out.nc = abs(c2 - c1) + 1
            for r in range(min(r1, r2), max(r1, r2) + 1):
                for c in range(min(c1, c2), max(c1, c2) + 1):
                    v = s.value((r, c))
                    if isinstance(v, Err): raise v
                    out.append(v)
            return out
        out = []
        for i, p in enumerate(re.split(r'("[^"]*")', f)):
            if i % 2: out.append(p); continue
            p = p.upper()
            p = re.sub(r"\$?([A-Z]{1,3})\$?(\d+):\$?([A-Z]{1,3})\$?(\d+)", r"R('\1\2','\3\4')", p)
            p = re.sub(r"(?<![A-Z0-9_.'])\$?([A-Z]{1,3})\$?(\d+)(?![\w('])", r"C('\1\2')", p)
            p = p.replace("<>", "!=")
            p = re.sub(r"(?<![<>!=])=(?!=)", "==", p).replace("^", "**")
            out.append(p)
        expr = "".join(out)
        if "__" in expr: raise Err("#NAME?")
        r = eval(lazy(expr), dict(F, C=C, R=R, __builtins__={}))
        if isinstance(r, list): raise Err("#VALUE!")
        if isinstance(r, complex): raise Err("#NUM!")
        return r

class Model(QAbstractTableModel):
    ROWS, COLS = 200, 26
    def __init__(s): super().__init__(); s.sh = Sheet(); s.hl = None; s.hl_color = QColor(0, 0, 0, 0)
    def rowCount(s, p=QModelIndex()): return s.ROWS
    def columnCount(s, p=QModelIndex()): return s.COLS
    def headerData(s, sec, o, role=Qt.ItemDataRole.DisplayRole):
        if role == Qt.ItemDataRole.DisplayRole:
            return col_name(sec) if o == Qt.Orientation.Horizontal else str(sec + 1)
    def data(s, i, role=Qt.ItemDataRole.DisplayRole):
        k = (i.row(), i.column())
        if role == Qt.ItemDataRole.DisplayRole:
            v = s.sh.value(k); return v.code if isinstance(v, Err) else fmt(v)
        if role == Qt.ItemDataRole.EditRole: return s.sh.raw.get(k, "")
        if role == Qt.ItemDataRole.BackgroundRole:
            if s.hl:
                r0, c0, r1, c1 = s.hl
                if r0 <= k[0] <= r1 and c0 <= k[1] <= c1: return QBrush(s.hl_color)
            if k in s.sh.style and "bg" in s.sh.style[k]:
                return QBrush(QColor(s.sh.style[k]["bg"]))
        if role == Qt.ItemDataRole.ForegroundRole:
            if k in s.sh.style and "fg" in s.sh.style[k]:
                return QBrush(QColor(s.sh.style[k]["fg"]))
        if role == Qt.ItemDataRole.TextAlignmentRole:
            a = Qt.AlignmentFlag.AlignRight if isnum(s.sh.value(k)) else Qt.AlignmentFlag.AlignLeft
            return (a | Qt.AlignmentFlag.AlignVCenter).value
    def flags(s, i):
        return Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEditable
    def setData(s, i, v, role=Qt.ItemDataRole.EditRole):
        if role != Qt.ItemDataRole.EditRole: return False
        s.put({(i.row(), i.column()): str(v)}); return True
    def put(s, d):
        for k, v in d.items():
            if str(v).strip() == "": s.sh.raw.pop(k, None)
            else: s.sh.raw[k] = str(v)
        s.sh.cache.clear()
        s.dataChanged.emit(s.index(0, 0), s.index(s.ROWS - 1, s.COLS - 1))
    def set_cell_style(s, keys, bg=None, fg=None):
        for k in keys:
            if k not in s.sh.style: s.sh.style[k] = {}
            if bg is not None: s.sh.style[k]["bg"] = bg
            if fg is not None: s.sh.style[k]["fg"] = fg
        s.dataChanged.emit(s.index(0, 0), s.index(s.ROWS - 1, s.COLS - 1))
    def reset_cell_style(s, keys):
        for k in keys:
            s.sh.style.pop(k, None)
        s.dataChanged.emit(s.index(0, 0), s.index(s.ROWS - 1, s.COLS - 1))
    def reset_data(s, d, styles=None):
        s.beginResetModel()
        s.sh.raw = dict(d)
        s.sh.style = dict(styles or {})
        s.sh.cache.clear()
        s.endResetModel()

OPS = "=(,+-*/^<>&;"

class Deleg(QStyledItemDelegate):
    def eventFilter(s, ed, ev):
        if ev.type() == QEvent.Type.FocusOut:
            v = s.parent()
            if QApplication.focusWidget() is v and v._ref_ctx(ed) is not None: return False
        if ev.type() == QEvent.Type.KeyPress and ev.key() == Qt.Key.Key_F4:
            if isinstance(ed, QLineEdit):
                t, p = ed.text(), ed.cursorPosition()
                nt, np = toggle_f4_ref(t, p)
                ed.setText(nt)
                ed.setCursorPosition(np)
                return True
        return super().eventFilter(ed, ev)

class Table(QTableView):
    formula_bar = None

    def __init__(s):
        super().__init__()
        s._drag = s._prev = s._last = None
        s._fill_drag = None
        s._dg = Deleg(s); s.setItemDelegate(s._dg)
        QApplication.instance().focusChanged.connect(s._focus)
        s.setMouseTracking(True)
        s.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        s.customContextMenuRequested.connect(s.show_context_menu)

    def show_context_menu(s, pos):
        menu = QMenu(s)
        a_bg = menu.addAction(tr("cell_bg_btn"))
        a_fg = menu.addAction(tr("cell_fg_btn"))
        a_rst = menu.addAction(tr("cell_color_reset"))
        action = menu.exec(s.viewport().mapToGlobal(pos))
        
        main_win = s.window()
        if action == a_bg and hasattr(main_win, "change_cell_bg"): main_win.change_cell_bg()
        elif action == a_fg and hasattr(main_win, "change_cell_fg"): main_win.change_cell_fg()
        elif action == a_rst and hasattr(main_win, "reset_cell_color"): main_win.reset_cell_color()

    def _focus(s, old, new):
        if new is s and old is not None:
            s._prev = old; QTimer.singleShot(0, s._clr)
    def _clr(s): s._prev = None

    def _ok(s, w):
        return isinstance(w, QLineEdit) and (w is s.formula_bar or w.parent() is s.viewport())

    def _ref_ctx(s, w):
        try: t, p = w.text(), w.cursorPosition()
        except RuntimeError: return None
        if not is_formula(t) and not t.startswith("="):
            if t == "":
                return p, 0
            return None
        L = s._last
        if L and L[0] is w and L[3] == t and p == L[1] + L[2]: return L[1], L[2]
        h = t[:p].rstrip()
        if h and h[-1] in OPS and h.count('"') % 2 == 0: return p, 0
        if not h:
            return p, 0
        return None

    def _set_ref(s, w, a, b, st, n):
        r0, r1 = sorted((a[0], b[0])); c0, c1 = sorted((a[1], b[1]))
        tl = f"{col_name(c0)}{r0 + 1}"
        r = tl if (r0, c0) == (r1, c1) else f"{tl}:{col_name(c1)}{r1 + 1}"
        t = w.text(); t = t[:st] + r + t[st + n:]
        w.setText(t); w.setCursorPosition(st + len(r))
        s._last = (w, st, len(r), t, a)
        s.model().hl = (r0, c0, r1, c1); s.viewport().update()

    def _clear_hl(s):
        if s.model() and s.model().hl: s.model().hl = None; s.viewport().update()

    def _ref_press(s, e):
        w, s._prev = s._prev, None
        i = s.indexAt(e.position().toPoint())
        if w is None or not i.isValid() or e.button() != Qt.MouseButton.LeftButton or not s._ok(w): return False
        c = s._ref_ctx(w)
        if c is None: return False
        st, n = c; b = (i.row(), i.column()); a = b; L = s._last
        if n and L and L[0] is w and e.modifiers() & Qt.KeyboardModifier.ShiftModifier: a = L[4]
        s._set_ref(w, a, b, st, n); s._drag = (w, a, st)
        w.setFocus(); return True

    def get_selection_bounds(s):
        indexes = s.selectedIndexes()
        if not indexes:
            cur = s.currentIndex()
            if cur.isValid():
                return cur.row(), cur.column(), cur.row(), cur.column()
            return None
        r_min = min(idx.row() for idx in indexes)
        r_max = max(idx.row() for idx in indexes)
        c_min = min(idx.column() for idx in indexes)
        c_max = max(idx.column() for idx in indexes)
        return r_min, c_min, r_max, c_max

    def mousePressEvent(s, e):
        if s._ref_press(e): return
        pos = e.position().toPoint()
        i = s.indexAt(pos)
        if i.isValid() and e.button() == Qt.MouseButton.LeftButton:
            bounds = s.get_selection_bounds()
            if bounds:
                r_min, c_min, r_max, c_max = bounds
                br_rect = s.visualRect(s.model().index(r_max, c_max))
                if abs(pos.x() - br_rect.right()) <= 6 and abs(pos.y() - br_rect.bottom()) <= 6:
                    s._fill_drag = (r_min, c_min, r_max, c_max)
                    return
            
        super().mousePressEvent(e)

    def mouseMoveEvent(s, e):
        if s._drag:
            w, a, st = s._drag; i = s.indexAt(e.position().toPoint())
            if i.isValid(): s._set_ref(w, a, (i.row(), i.column()), st, s._last[2])
            return
        
        pos = e.position().toPoint()
        i = s.indexAt(pos)
        if i.isValid():
            br_rect = s.visualRect(i)
            if abs(pos.x() - br_rect.right()) <= 6 and abs(pos.y() - br_rect.bottom()) <= 6:
                s.setCursor(Qt.CursorShape.CrossCursor)
            else:
                s.unsetCursor()
                
            if s._fill_drag and e.buttons() == Qt.MouseButton.LeftButton:
                r_min, c_min, r_max, c_max = s._fill_drag
                tr, tc = i.row(), i.column()
                s.model().hl = (min(r_min, tr), min(c_min, tc), max(r_max, tr), max(c_max, tc))
                s.viewport().update()
                return

        super().mouseMoveEvent(e)

    def mouseReleaseEvent(s, e):
        s.unsetCursor()
        if s._drag: w = s._drag[0]; s._drag = None; w.setFocus(); return
        
        pos = e.position().toPoint()
        i = s.indexAt(pos)
        
        if s._fill_drag and i.isValid():
            r_min, c_min, r_max, c_max = s._fill_drag
            tr, tc = i.row(), i.column()
            s._fill_drag = None
            s._clear_hl()
            s.perform_fill(r_min, c_min, r_max, c_max, tr, tc)
            return
            
        super().mouseReleaseEvent(e)

    def get_fill_sequence_1d(s, src_vals, source_start_pos, target_start_pos, target_end_pos, is_vertical=True):
        n = len(src_vals)
        if n == 0: return []
            
        numeric_vals = []
        all_numeric = True
        for v in src_vals:
            try:
                if v is None or str(v).strip() == "" or str(v).startswith("="):
                    all_numeric = False
                    break
                numeric_vals.append(float(v))
            except (ValueError, TypeError):
                all_numeric = False
                break
                
        if all_numeric and n >= 2:
            diffs = [numeric_vals[i+1] - numeric_vals[i] for i in range(n-1)]
            base_d = diffs[0]
            is_arithmetic = all(abs(d - base_d) < 1e-9 for d in diffs)
            if is_arithmetic:
                out = []
                for pos in range(target_start_pos, target_end_pos + 1):
                    offset = pos - source_start_pos
                    val = numeric_vals[0] + offset * base_d
                    if abs(val - round(val)) < 1e-9:
                        out.append(str(int(round(val))))
                    else:
                        out.append(str(val))
                return out

        all_formulas = all(str(v).startswith("=") for v in src_vals)
        if all_formulas:
            out = []
            for pos in range(target_start_pos, target_end_pos + 1):
                offset = pos - source_start_pos
                src_idx = (offset % n + n) % n
                formula = src_vals[src_idx]
                dr = offset if is_vertical else 0
                dc = 0 if is_vertical else offset
                out.append(shift_formula(formula, dr, dc))
            return out

        out = []
        for pos in range(target_start_pos, target_end_pos + 1):
            offset = pos - source_start_pos
            src_idx = (offset % n + n) % n
            out.append(src_vals[src_idx])
        return out

    def perform_fill(s, r_min, c_min, r_max, c_max, tr, tc):
        raw = s.model().sh.raw
        updates = {}
        
        d_row = tr - r_max if tr > r_max else (tr - r_min if tr < r_min else 0)
        d_col = tc - c_max if tc > c_max else (tc - c_min if tc < c_min else 0)
        
        if abs(d_row) >= abs(d_col) and d_row != 0:
            target_r_min = min(r_min, tr)
            target_r_max = max(r_max, tr)
            for c in range(c_min, c_max + 1):
                src_vals = [raw.get((r, c), "") for r in range(r_min, r_max + 1)]
                fill_vals = s.get_fill_sequence_1d(src_vals, r_min, target_r_min, target_r_max, is_vertical=True)
                for idx, r in enumerate(range(target_r_min, target_r_max + 1)):
                    if idx < len(fill_vals):
                        v = fill_vals[idx]
                        if v != "": updates[(r, c)] = v
                        else: updates.pop((r, c), None)
                            
        elif abs(d_col) > abs(d_row) and d_col != 0:
            target_c_min = min(c_min, tc)
            target_c_max = max(c_max, tc)
            for r in range(r_min, r_max + 1):
                src_vals = [raw.get((r, c), "") for c in range(c_min, c_max + 1)]
                fill_vals = s.get_fill_sequence_1d(src_vals, c_min, target_c_min, target_c_max, is_vertical=False)
                for idx, c in enumerate(range(target_c_min, target_c_max + 1)):
                    if idx < len(fill_vals):
                        v = fill_vals[idx]
                        if v != "": updates[(r, c)] = v
                        else: updates.pop((r, c), None)
                            
        if updates:
            s.model().put(updates)

    def mouseDoubleClickEvent(s, e):
        if not s._ref_press(e): super().mouseDoubleClickEvent(e)

    def closeEditor(s, ed, hint):
        s._clear_hl(); s._last = None; super().closeEditor(ed, hint)

    def keyPressEvent(s, e):
        if e.key() in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
            if s.state() != QTableView.State.EditingState:
                cur = s.currentIndex()
                if cur.isValid() and cur.row() + 1 < s.model().rowCount():
                    s.setCurrentIndex(s.model().index(cur.row() + 1, cur.column()))
                e.accept()
                return
        if e.matches(QKeySequence.StandardKey.Copy): s.copy_sel()
        elif e.matches(QKeySequence.StandardKey.Paste): s.paste()
        elif e.key() == Qt.Key.Key_Delete:
            s.model().put({(i.row(), i.column()): "" for i in s.selectedIndexes()})
        else: super().keyPressEvent(e)

    def copy_sel(s):
        idx = s.selectedIndexes()
        if not idx: return
        r0, r1 = min(i.row() for i in idx), max(i.row() for i in idx)
        c0, c1 = min(i.column() for i in idx), max(i.column() for i in idx)
        raw = s.model().sh.raw
        QApplication.clipboard().setText("\n".join(
            "\t".join(raw.get((r, c), "") for c in range(c0, c1 + 1)) for r in range(r0, r1 + 1)))

    def paste(s):
        text = QApplication.clipboard().text().replace("\r", "")
        if not text: return
        grid = [l.split("\t") for l in text.rstrip("\n").split("\n")]
        sel = s.selectedIndexes()
        if len(grid) == 1 and len(grid[0]) == 1 and len(sel) > 1:
            s.model().put({(i.row(), i.column()): grid[0][0] for i in sel}); return
        r0 = min((i.row() for i in sel), default=0); c0 = min((i.column() for i in sel), default=0)
        m = s.model()
        m.put({(r0 + r, c0 + c): v for r, row in enumerate(grid) for c, v in enumerate(row)
               if r0 + r < m.ROWS and c0 + c < m.COLS})

class Dlg(QDialog):
    def __init__(s, main):
        super().__init__(main); s.m = main; s.body = None
        s.colors = dict(main.colors); s.fonts = {k: list(v) for k, v in main.fonts.items()}
        s.resize(460, 680); s.lay = QVBoxLayout(s); s.build()

    def build(s):
        if s.body: s.lay.removeWidget(s.body); s.body.deleteLater()
        s.body = QWidget(); s.lay.addWidget(s.body); s.setWindowTitle(tr("settings"))
        lay = QVBoxLayout(s.body); lay.setContentsMargins(0, 0, 0, 0)
        
        io_bar = QHBoxLayout()
        b_exp = QPushButton(tr("export_cfg")); b_exp.clicked.connect(s.export_cfg)
        b_imp = QPushButton(tr("import_cfg")); b_imp.clicked.connect(s.import_cfg)
        io_bar.addWidget(b_exp); io_bar.addWidget(b_imp)
        lay.addLayout(io_bar)

        gen = QFormLayout(); s.theme_cb, s.lang_cb = QComboBox(), QComboBox()
        for k in THEMES: s.theme_cb.addItem(tr("default") if k == "default" else k, k)
        s.theme_cb.setCurrentIndex(s.theme_cb.findData(s.m.theme))
        for n, c in LANGS: s.lang_cb.addItem(n, c)
        s.lang_cb.setCurrentIndex(s.lang_cb.findData(LANG))
        s.theme_cb.currentIndexChanged.connect(s.on_theme); s.lang_cb.currentIndexChanged.connect(s.on_lang)
        gen.addRow(tr("theme"), s.theme_cb); gen.addRow(tr("lang"), s.lang_cb); lay.addLayout(gen)
        sc = QScrollArea(); sc.setWidgetResizable(True)
        box = QWidget(); form = QFormLayout(box); sc.setWidget(box); lay.addWidget(sc, 1)
        form.addRow(QLabel("<b>" + tr("colors") + "</b>")); s.btns = {}
        for k in COLOR_KEYS:
            b = QPushButton(); b.clicked.connect(lambda _, k=k: s.pick(k)); s.btns[k] = b
            s.paint(k); form.addRow(tr("c_" + k), b)
        form.addRow(QLabel("<b>" + tr("fonts") + "</b>"))
        for k in FONT_KEYS:
            fc = QFontComboBox(); fc.setCurrentFont(QFont(s.fonts[k][0]) if s.fonts[k][0] else QFont("맑은 고딕"))
            sp = QSpinBox(); sp.setRange(6, 48); sp.setValue(s.fonts[k][1])
            fc.currentFontChanged.connect(lambda f, k=k: s.fonts[k].__setitem__(0, f.family()))
            sp.valueChanged.connect(lambda v, k=k: s.fonts[k].__setitem__(1, v))
            row = QHBoxLayout(); row.addWidget(fc, 1); row.addWidget(sp); form.addRow(tr("f_" + k), row)
        bar = QHBoxLayout()
        for key, fn in (("reset", s.reset), ("apply", s.apply), ("close", s.accept)):
            b = QPushButton(tr(key)); b.clicked.connect(fn); bar.addWidget(b)
        lay.addLayout(bar)

    def export_cfg(s):
        cfg = dict(lang=LANG, theme=s.m.theme, colors=s.colors, fonts=s.fonts)
        p, _ = QFileDialog.getSaveFileName(s, tr("export_cfg"), "pysheet_config.json", "JSON (*.json)")
        if p:
            Path(p).write_text(json.dumps(cfg, indent=2, ensure_ascii=False), encoding="utf-8")

    def import_cfg(s):
        p, _ = QFileDialog.getOpenFileName(s, tr("import_cfg"), "", "JSON (*.json)")
        if p:
            try:
                cfg = json.loads(Path(p).read_text(encoding="utf-8"))
                s.colors = cfg.get("colors", s.colors)
                s.fonts = cfg.get("fonts", s.fonts)
                if "theme" in cfg and cfg["theme"] in THEMES:
                    s.m.theme = cfg["theme"]
                if "lang" in cfg:
                    global LANG
                    LANG = cfg["lang"]
                s.build()
            except Exception as e:
                QMessageBox.warning(s, "Error", f"Failed to load settings: {e}")

    def on_theme(s, i):
        s.m.set_theme(s.theme_cb.itemData(i)); s.colors = dict(s.m.colors)
        for k in COLOR_KEYS: s.paint(k)
    def on_lang(s, i):
        s.m.set_lang(s.lang_cb.itemData(i)); QTimer.singleShot(0, s.build)

    def paint(s, k):
        c = s.colors[k]; s.btns[k].setText(c)
        s.btns[k].setStyleSheet(f"QPushButton{{background:{c};color:{contrast(c)};}}")
    def pick(s, k):
        c = QColorDialog.getColor(QColor(s.colors[k]), s, tr("c_" + k))
        if c.isValid(): s.colors[k] = c.name(); s.paint(k)
    def reset(s):
        s.colors = make_colors(s.m.theme)
        s.fonts = {k: list(v) for k, v in DEF_FONTS.items()}
        for k in COLOR_KEYS: s.paint(k)
        s.build()
    def apply(s):
        s.m.colors = dict(s.colors); s.m.fonts = {k: list(v) for k, v in s.fonts.items()}
        s.m.apply_style(); s.m.save()
        s.accept()

class Main(QMainWindow):
    def __init__(s):
        global LANG
        super().__init__()
        s.current_memo_index = 0
        s.last_bg_color = "#ffffcc"
        s.last_fg_color = "#000000"
        s.st = QSettings("PySheetCalc", "App")
        try: cfg = json.loads(s.st.value("cfg", "{}") or "{}")
        except ValueError: cfg = {}
        LANG = cfg.get("lang", "ko")
        s.theme = cfg.get("theme", "default") if cfg.get("theme") in THEMES else "default"
        s.colors = {**make_colors(s.theme), **cfg.get("colors", {})}
        s.fonts = {**{k: list(v) for k, v in DEF_FONTS.items()}, **cfg.get("fonts", {})}

        s.model = Model(); s.table = Table(); s.table.setModel(s.model)
        s.table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)

        s.b_cell_bg = QToolButton()
        s.b_cell_bg.setPopupMode(QToolButton.ToolButtonPopupMode.MenuButtonPopup)
        s.b_cell_bg.setMenu(s.create_color_menu(True))
        s.b_cell_bg.clicked.connect(s.apply_last_cell_bg)
        
        s.b_cell_fg = QToolButton()
        s.b_cell_fg.setPopupMode(QToolButton.ToolButtonPopupMode.MenuButtonPopup)
        s.b_cell_fg.setMenu(s.create_color_menu(False))
        s.b_cell_fg.clicked.connect(s.apply_last_cell_fg)

        s.b_settings = QPushButton(); s.b_settings.clicked.connect(s.open_settings)

        s.name = QLabel(); s.name.setFixedWidth(60)
        s.formula = QLineEdit(); s.formula.setObjectName("formula")
        s.formula.returnPressed.connect(lambda: s.commit(False))
        s.formula.installEventFilter(s)
        s.b_fill = QPushButton(); s.b_fill.clicked.connect(lambda: s.commit(True))
        
        fb = QHBoxLayout()
        for w, st in ((s.name, 0), (s.formula, 1), (s.b_fill, 0), (s.b_cell_bg, 0), (s.b_cell_fg, 0), (s.b_settings, 0)):
            fb.addWidget(w, st)

        s.status = QLabel(); s.status.setObjectName("status")
        
        sheet_widget = QWidget()
        v = QVBoxLayout(sheet_widget)
        v.setContentsMargins(0, 0, 0, 0)
        v.addLayout(fb); v.addWidget(s.table, 1); v.addWidget(s.status)

        s.memo_container = QWidget()
        memo_lay = QVBoxLayout(s.memo_container)
        memo_lay.setContentsMargins(0, 0, 0, 0)
        
        s.mlist = QListWidget()
        s.mlist.currentRowChanged.connect(s.on_memo_selection_changed)
        memo_lay.addWidget(s.mlist, 1)
        
        s.mbtns = []
        for key, fn in (("m_save", s.m_save), ("m_over", s.m_over),
                        ("m_ren", s.m_ren), ("m_del", s.m_del), ("m_clear", s.m_clear)):
            b = QPushButton(); b.clicked.connect(fn); memo_lay.addWidget(b); s.mbtns.append((b, key))

        s.splitter = QSplitter(Qt.Orientation.Horizontal)
        s.splitter.addWidget(sheet_widget)
        s.splitter.addWidget(s.memo_container)
        
        sw_sizes = s.st.value("splitter_sizes")
        if sw_sizes:
            try: s.splitter.setSizes([int(x) for x in sw_sizes])
            except: s.splitter.setSizes([850, 250])
        else:
            s.splitter.setSizes([850, 250])

        central_w = QWidget()
        cent_lay = QVBoxLayout(central_w)
        cent_lay.setContentsMargins(0, 0, 0, 0)
        cent_lay.addWidget(s.splitter)
        s.setCentralWidget(central_w)
        s.resize(1100, 650)

        s.m_file = s.menuBar().addMenu("")
        s.acts = []
        for key, fn in (("new", s.new), ("open", s.open_csv), ("save", s.save_csv), ("quit", s.close)):
            a = QAction(s); a.triggered.connect(fn); s.m_file.addAction(a); s.acts.append((a, key))

        s.table.selectionModel().currentChanged.connect(s.on_cur)
        s.table.selectionModel().selectionChanged.connect(lambda *_: s.update_status())
        s.model.dataChanged.connect(lambda *_: (s.on_cur(), s.update_status(), s.auto_save_memo()))
        s.table.formula_bar = s.formula
        s._alt_clean = False; s._alt_t = 0.0
        
        s.load_memo_data()
        s.apply_style()
        s.retr()
        
        QApplication.instance().installEventFilter(s)
        
        if s.current_work_memo and s.current_work_memo.get("raw"):
            raw = {tuple(map(int, k.split(","))): v for k, v in s.current_work_memo.get("raw", {}).items()}
            style = {tuple(map(int, k.split(","))): v for k, v in s.current_work_memo.get("style", {}).items()}
            s.model.reset_data(raw, style)
        
        g = s.st.value("geo")
        if g: s.restoreGeometry(g)

    def create_color_menu(s, is_bg):
        menu = QMenu(s)
        action_pick = menu.addAction(tr("custom"))
        action_pick.triggered.connect(s.pick_new_color_bg if is_bg else s.pick_new_color_fg)
        
        action_reset = menu.addAction(tr("cell_color_reset"))
        action_reset.triggered.connect(s.reset_cell_color)
        return menu

    def apply_last_cell_bg(s):
        idx = s.table.selectedIndexes()
        if not idx: return
        keys = [(i.row(), i.column()) for i in idx]
        s.model.set_cell_style(keys, bg=s.last_bg_color)
        s.retr()

    def apply_last_cell_fg(s):
        idx = s.table.selectedIndexes()
        if not idx: return
        keys = [(i.row(), i.column()) for i in idx]
        s.model.set_cell_style(keys, fg=s.last_fg_color)
        s.retr()

    def pick_new_color_bg(s):
        idx = s.table.selectedIndexes()
        c = QColorDialog.getColor(QColor(s.last_bg_color), s, tr("cell_bg_btn"))
        if c.isValid():
            s.last_bg_color = c.name()
            if idx:
                keys = [(i.row(), i.column()) for i in idx]
                s.model.set_cell_style(keys, bg=s.last_bg_color)
            s.retr()

    def pick_new_color_fg(s):
        idx = s.table.selectedIndexes()
        c = QColorDialog.getColor(QColor(s.last_fg_color), s, tr("cell_fg_btn"))
        if c.isValid():
            s.last_fg_color = c.name()
            if idx:
                keys = [(i.row(), i.column()) for i in idx]
                s.model.set_cell_style(keys, fg=s.last_fg_color)
            s.retr()

    def reset_cell_color(s):
        idx = s.table.selectedIndexes()
        if not idx: return
        keys = [(i.row(), i.column()) for i in idx]
        s.model.reset_cell_style(keys)
        s.retr()

    def apply_style(s):
        def fc(k):
            fam, sz = s.fonts[k]
            return (f'font-family:"{fam}";' if fam else "") + f"font-size:{sz}pt;"
        d = dict(s.colors, **{"f_" + k: fc(k) for k in FONT_KEYS})
        s.setStyleSheet(QSS % d)
        c = QColor(s.colors["select_bg"]); c.setAlpha(90); s.model.hl_color = c
        cs = s.fonts["cell"][1]
        s.table.verticalHeader().setDefaultSectionSize(int(cs * 2.4))
        s.table.horizontalHeader().setDefaultSectionSize(int(cs * 8.5))

    def retr(s):
        s.setWindowTitle(tr("title")); s.m_file.setTitle(tr("file"))
        for a, k in s.acts: a.setText(tr(k))
        s.b_settings.setText("⚙ " + tr("settings"))
        
        s.b_cell_bg.setIcon(make_color_icon(s.last_bg_color))
        s.b_cell_bg.setText(tr("cell_bg_btn"))
        s.b_cell_bg.setMenu(s.create_color_menu(True))
        
        s.b_cell_fg.setIcon(make_color_icon(s.last_fg_color))
        s.b_cell_fg.setText(tr("cell_fg_btn"))
        s.b_cell_fg.setMenu(s.create_color_menu(False))
        
        s.b_fill.setText(tr("fill")); s.formula.setPlaceholderText(tr("ph")); s.update_status()
        
        if hasattr(s, "m_view"): s.m_view.setTitle(tr("view"))
        for b, k in s.mbtns: b.setText(tr(k))
        s.refresh_memo_list_titles()

    def open_settings(s): Dlg(s).exec()
    def set_theme(s, k):
        s.theme = k; s.colors = make_colors(k); s.apply_style(); s.save()
    def set_lang(s, c):
        global LANG
        LANG = c; s.retr(); s.save()

    def on_cur(s, *_):
        s.table._clear_hl()
        i = s.table.currentIndex()
        if i.isValid():
            s.name.setText(f"{col_name(i.column())}{i.row() + 1}")
            s.formula.setText(s.model.sh.raw.get((i.row(), i.column()), ""))
        s.retr()

    def commit(s, fill):
        cur = s.table.currentIndex(); sel = s.table.selectedIndexes()
        targets = sel if fill and sel else [cur]
        s.table._clear_hl(); s.table._last = None
        s.model.put({(i.row(), i.column()): s.formula.text() for i in targets if i.isValid()})
        s.table.setFocus()

    def update_status(s):
        vals = [s.model.sh.value((i.row(), i.column())) for i in s.table.selectedIndexes()]
        ne = [v for v in vals if v not in (None, "")]; ns = [v for v in vals if isnum(v)]
        sm = sum(ns)
        s.status.setText(f"{tr('count')}: {len(ne)}     {tr('sum')}: {fmt(float(sm))}     "
                         f"{tr('avg')}: {fmt(sm / len(ns)) if ns else '-'}")

    def load_memo_data(s):
        try:
            data = json.loads(MEMO_FILE.read_text(encoding="utf-8"))
            s.current_work_memo = data.get("current_work", dict(title=tr("current_work"), time=s.now(), raw={}, style={}))
            s.memos = data.get("memos", [])
        except Exception:
            s.current_work_memo = dict(title=tr("current_work"), time=s.now(), raw={}, style={})
            s.memos = []
        
        s.m_view = s.menuBar().addMenu("")
        toggle_action = QAction(tr("memo") + "  (Alt ×2)", s)
        toggle_action.triggered.connect(s.toggle_memo)
        s.m_view.addAction(toggle_action)
        
        s.refresh_memo_list(0)

    def refresh_memo_list(s, sel_idx=0):
        s.mlist.blockSignals(True)
        s.mlist.clear()
        
        curr_title = tr("current_work")
        s.current_work_memo["title"] = curr_title
        s.mlist.addItem(f"{curr_title}\n{s.current_work_memo.get('time', s.now())}")
        
        for m in s.memos:
            s.mlist.addItem(f"{m['title']}\n{m['time']}")
            
        s.current_memo_index = sel_idx
        s.mlist.setCurrentRow(sel_idx)
        s.mlist.blockSignals(False)
        s.persist_memos_to_disk()

    def refresh_memo_list_titles(s):
        s.mlist.blockSignals(True)
        for idx in range(s.mlist.count()):
            item = s.mlist.item(idx)
            if idx == 0:
                item.setText(f"{tr('current_work')}\n{s.current_work_memo.get('time', s.now())}")
            else:
                m = s.memos[idx - 1]
                item.setText(f"{m['title']}\n{m['time']}")
        s.mlist.blockSignals(False)

    def persist_memos_to_disk(s):
        try:
            payload = {"current_work": s.current_work_memo, "memos": s.memos}
            MEMO_FILE.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
        except OSError: pass

    def auto_save_memo(s):
        idx = s.current_memo_index
        now_time = s.now()
        raw_snap = s.snapshot()
        style_snap = s.snapshot_style()

        if idx == 0:
            s.current_work_memo.update(raw=raw_snap, style=style_snap, time=now_time)
            s.mlist.blockSignals(True)
            item = s.mlist.item(0)
            if item: item.setText(f"{tr('current_work')}\n{now_time}")
            s.mlist.blockSignals(False)
        elif 0 < idx <= len(s.memos):
            memo_idx = idx - 1
            s.memos[memo_idx].update(raw=raw_snap, style=style_snap, time=now_time)
            s.mlist.blockSignals(True)
            item = s.mlist.item(idx)
            if item: item.setText(f"{s.memos[memo_idx]['title']}\n{now_time}")
            s.mlist.blockSignals(False)
        
        s.persist_memos_to_disk()

    def on_memo_selection_changed(s, new_idx):
        if new_idx < 0: return
        
        s.auto_save_memo()
        
        s.current_memo_index = new_idx
        target_memo = s.current_work_memo if new_idx == 0 else s.memos[new_idx - 1]
        
        raw = {tuple(map(int, k.split(","))): v for k, v in target_memo.get("raw", {}).items()}
        style = {tuple(map(int, k.split(","))): v for k, v in target_memo.get("style", {}).items()}
        
        s.model.reset_data(raw, style)
        s.table.clearSelection()
        s.on_cur()
        s.update_status()

    def toggle_memo(s):
        vis = s.memo_container.isVisible()
        s.memo_container.setVisible(not vis)

    def eventFilter(s, o, e):
        t = e.type()
        if t == QEvent.Type.KeyPress:
            if e.key() == Qt.Key.Key_Alt:
                if not e.isAutoRepeat(): s._alt_clean = True
                return True
            elif e.key() == Qt.Key.Key_F4 and o is s.formula:
                txt, pos = s.formula.text(), s.formula.cursorPosition()
                nt, np = toggle_f4_ref(txt, pos)
                s.formula.setText(nt)
                s.formula.setCursorPosition(np)
                return True
            s._alt_clean = False; s._alt_t = 0.0
        elif t == QEvent.Type.KeyRelease and e.key() == Qt.Key.Key_Alt:
            if not e.isAutoRepeat():
                if s._alt_clean and s.isActiveWindow():
                    n = time.monotonic()
                    if n - s._alt_t < 0.45: s._alt_t = 0.0; s.toggle_memo()
                    else: s._alt_t = n
                s._alt_clean = False;
            return True
        elif t == QEvent.Type.MouseButtonPress: s._alt_clean = False
        return False

    def snapshot(s):
        return {f"{r},{c}": v for (r, c), v in s.model.sh.raw.items()}
        
    def snapshot_style(s):
        return {f"{r},{c}": v for (r, c), v in s.model.sh.style.items()}
        
    def now(s): return datetime.datetime.now().strftime("%Y-%m-%d %H:%M")

    def m_save(s):
        t, ok = QInputDialog.getText(s, tr("memo"), tr("m_ask"))
        if ok and t.strip():
            s.memos.append(dict(title=t.strip(), time=s.now(), raw=s.snapshot(), style=s.snapshot_style()))
            s.refresh_memo_list(len(s.memos))
            
    def m_over(s):
        i = s.mlist.currentRow()
        if i >= 0:
            s.auto_save_memo()
            
    def m_ren(s):
        i = s.mlist.currentRow()
        if i <= 0: return
        memo_idx = i - 1
        if memo_idx < len(s.memos):
            t, ok = QInputDialog.getText(s, tr("memo"), tr("m_ask"), QLineEdit.EchoMode.Normal, s.memos[memo_idx]["title"])
            if ok and t.strip():
                s.memos[memo_idx]["title"] = t.strip()
                s.refresh_memo_list(i)
                
    def m_del(s):
        i = s.mlist.currentRow()
        if i <= 0: return
        memo_idx = i - 1
        if memo_idx < len(s.memos):
            s.memos.pop(memo_idx)
            s.refresh_memo_list(0)
            
    def m_clear(s):
        if not s.memos: return
        b = QMessageBox(s); b.setWindowTitle(tr("memo")); b.setText(tr("m_clear_ask"))
        yes = b.addButton(tr("m_clear"), QMessageBox.ButtonRole.DestructiveRole)
        b.addButton(tr("cancel"), QMessageBox.ButtonRole.RejectRole); b.exec()
        if b.clickedButton() is yes:
            s.memos.clear()
            s.refresh_memo_list(0)

    def new(s):
        s.model.reset_data({}, {})
        s.table.clearSelection()
        s.auto_save_memo()
    
    def open_csv(s):
        p, _ = QFileDialog.getOpenFileName(s, "", "", "CSV (*.csv)")
        if not p: return
        with open(p, newline="", encoding="utf-8-sig") as f: rows = list(csv.reader(f))
        s.model.reset_data({(r, c): v for r, row in enumerate(rows) for c, v in enumerate(row)
                            if v and r < Model.ROWS and c < Model.COLS}, {})
        s.table.clearSelection()
        s.auto_save_memo()
        
    def save_csv(s):
        p, _ = QFileDialog.getSaveFileName(s, "", "sheet.csv", "CSV (*.csv)")
        if not p: return
        raw = s.model.sh.raw
        R = max((k[0] for k in raw), default=0) + 1; C = max((k[1] for k in raw), default=0) + 1
        with open(p, "w", newline="", encoding="utf-8-sig") as f:
            csv.writer(f).writerows([[raw.get((r, c), "") for c in range(C)] for r in range(R)])

    def save(s):
        s.st.setValue("cfg", json.dumps(dict(lang=LANG, theme=s.theme, colors=s.colors, fonts=s.fonts),
                                        ensure_ascii=False))
        s.st.setValue("splitter_sizes", s.splitter.sizes())
                                        
    def closeEvent(s, e):
        s.auto_save_memo()
        s.st.setValue("geo", s.saveGeometry())
        s.st.setValue("splitter_sizes", s.splitter.sizes())
        s.save()
        super().closeEvent(e)

if __name__ == "__main__":
    app = QApplication(sys.argv); app.setStyle("Fusion")
    w = Main(); w.show(); sys.exit(app.exec())