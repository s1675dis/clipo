"""
clipo - クリップボード履歴マネージャー
システムトレイに常駐し、クリップボードの変更を監視・履歴として保存する。
Ctrl キーを素早く2回押すと履歴ポップアップが開く。
"""

import re
import sys
import winreg
import ctypes
import ctypes.wintypes
import time
import json
import threading
import tkinter as tk
import tkinter.messagebox
import tkinter.filedialog
from pathlib import Path
from datetime import datetime

import keyboard
import pyperclip
import pystray
from PIL import Image, ImageDraw

# 設定
MAX_HISTORY = 50          # 最大履歴件数
POLL_INTERVAL = 0.5       # クリップボード監視間隔（秒）
# PyInstaller --onefile では __file__ が一時展開ディレクトリを指すため、
# exe 実行時は sys.executable の親ディレクトリを基準にする
_BASE_DIR = (
    Path(sys.executable).parent
    if getattr(sys, "frozen", False)
    else Path(__file__).parent
)
HISTORY_FILE   = _BASE_DIR / "history.json"
CONFIG_FILE    = _BASE_DIR / "config.json"
TEMPLATES_FILE = _BASE_DIR / "templates.json"
PINS_FILE      = _BASE_DIR / "pins.json"
DOUBLE_CTRL_INTERVAL = 0.4  # ダブルCtrl判定間隔（秒）
_CF_HDROP        = 15   # Windows クリップボード形式: ファイルドロップ
_CF_UNICODETEXT  = 13   # Windows クリップボード形式: Unicode テキスト
POPUP_INIT_WIDTH  = 236   # ポップアップ初期幅（px）
POPUP_INIT_HEIGHT = 380   # ポップアップ初期高さ（px）
POPUP_MAX_ROWS = 15       # ポップアップに表示する最大行数
PAGE_JUMP      = 5        # PageUp/PageDown で移動する件数

history: list[dict] = []  # {"text": str, "time": str}
history_lock = threading.Lock()


# ---------- 設定ファイル ----------

def _load_config() -> dict:
    if CONFIG_FILE.exists():
        try:
            return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}


def _save_config(data: dict) -> None:
    try:
        current = _load_config()
        current.update(data)
        CONFIG_FILE.write_text(json.dumps(current, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception as e:
        print(f"[clipo] 設定保存エラー: {e}", file=sys.stderr)


# ---------- テンプレート ----------

def load_templates() -> list[dict]:
    """templates.json から定型文リストを読み込む。形式: [{"name": str, "text": str}]"""
    if TEMPLATES_FILE.exists():
        try:
            return json.loads(TEMPLATES_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass
    return []


def save_templates(templates: list[dict]) -> None:
    try:
        TEMPLATES_FILE.write_text(
            json.dumps(templates, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    except Exception as e:
        print(f"[clipo] テンプレート保存エラー: {e}", file=sys.stderr)


# ---------- ピンファイル ----------

def load_pins() -> list[dict]:
    """pins.json からピンリストを読み込む。形式: [{"text": str, "time": str}]"""
    if PINS_FILE.exists():
        try:
            return json.loads(PINS_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass
    return []


def save_pins(pins: list[dict]) -> None:
    try:
        PINS_FILE.write_text(
            json.dumps(pins, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    except Exception as e:
        print(f"[clipo] ピン保存エラー: {e}", file=sys.stderr)


# ---------- 履歴ファイル ----------

def load_history() -> None:
    if HISTORY_FILE.exists():
        try:
            data = json.loads(HISTORY_FILE.read_text(encoding="utf-8"))
            with history_lock:
                history.extend(data[-MAX_HISTORY:])
        except Exception:
            pass


def save_history() -> None:
    try:
        with history_lock:
            HISTORY_FILE.write_text(
                json.dumps(history, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
    except Exception as e:
        print(f"[clipo] 履歴保存エラー: {e}", file=sys.stderr)


# ---------- クリップボード監視 ----------


# GetClipboardData / DragQueryFileW の戻り値を 64 ビットポインタとして正しく受け取る
# （ctypes.windll のデフォルト restype は c_long=32bit のため 64bit 環境でハンドルが壊れる）
_GetClipboardData = ctypes.windll.user32.GetClipboardData
_GetClipboardData.restype = ctypes.c_void_p          # HANDLE は void* サイズ

_DragQueryFileW = ctypes.windll.shell32.DragQueryFileW
_DragQueryFileW.restype  = ctypes.c_uint              # 戻り値は UINT（ファイル数 or 長さ）
_DragQueryFileW.argtypes = [
    ctypes.c_void_p,   # HDROP hDrop
    ctypes.c_uint,     # UINT  iFile  (0xFFFFFFFF でファイル数取得)
    ctypes.c_wchar_p,  # LPWSTR lpszFile
    ctypes.c_uint,     # UINT  cch
]


def _read_clipboard() -> str:
    """クリップボードの内容を文字列として返す。
    CF_HDROP（ファイルコピー）を優先し、なければ pyperclip でテキストを取得する。
    IsClipboardFormatAvailable でフォーマットを先に判定することで
    クリップボードの二重オープンを避ける。
    """
    if ctypes.windll.user32.IsClipboardFormatAvailable(_CF_HDROP):
        # ファイルコピー: CF_HDROP からファイル名のみ抽出
        for _ in range(3):
            if not ctypes.windll.user32.OpenClipboard(None):
                time.sleep(0.02)
                continue
            try:
                h = _GetClipboardData(_CF_HDROP)
                if h:
                    count = _DragQueryFileW(h, 0xFFFFFFFF, None, 0)
                    names = []
                    for i in range(count):
                        length = _DragQueryFileW(h, i, None, 0)
                        if length:
                            buf = ctypes.create_unicode_buffer(length + 1)
                            _DragQueryFileW(h, i, buf, length + 1)
                            names.append(Path(buf.value).name)
                    if names:
                        return "\n".join(names)
            except Exception:
                pass
            finally:
                ctypes.windll.user32.CloseClipboard()
            break
        return ""

    # テキストコピー: pyperclip に委譲（CF_UNICODETEXT の読み取りを担当）
    try:
        return pyperclip.paste()
    except Exception:
        return ""


def watch_clipboard(icon: pystray.Icon) -> None:
    prev_seq = ctypes.windll.user32.GetClipboardSequenceNumber()

    while not getattr(icon, "_stop_event", threading.Event()).is_set():
        time.sleep(POLL_INTERVAL)

        try:
            seq = ctypes.windll.user32.GetClipboardSequenceNumber()
        except Exception:
            continue

        if seq == prev_seq:
            continue

        # 変更検出後に少し待ち、Windows がすべてのフォーマットを書き終えた状態を読む
        time.sleep(0.1)
        try:
            prev_seq = ctypes.windll.user32.GetClipboardSequenceNumber()
        except Exception:
            prev_seq = seq

        current = _read_clipboard()
        if not current:
            continue

        with history_lock:
            if history and history[0]["text"] == current:
                continue
            history[:] = [h for h in history if h["text"] != current]
            history.insert(0, {"text": current, "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S")})
            if len(history) > MAX_HISTORY:
                history.pop()
        save_history()
        icon.update_menu()


# ---------- トレイアイコン画像 ----------

def create_icon_image(size: int = 64) -> Image.Image:
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    # 青いクリップボードアイコン風
    margin = size // 8
    draw.rectangle(
        [margin, margin * 2, size - margin, size - margin],
        fill=(70, 130, 180),
        outline=(30, 80, 130),
        width=max(1, size // 32),
    )
    # クリップ部分
    clip_w = size // 3
    clip_h = size // 5
    cx = size // 2
    draw.rectangle(
        [cx - clip_w // 2, margin - clip_h // 2, cx + clip_w // 2, margin + clip_h // 2],
        fill=(200, 220, 240),
        outline=(30, 80, 130),
        width=max(1, size // 32),
    )
    # 罫線
    line_x0 = margin * 2
    line_x1 = size - margin * 2
    for i in range(1, 4):
        y = margin * 2 + (size - margin * 3) * i // 4
        draw.line([(line_x0, y), (line_x1, y)], fill=(200, 220, 240), width=max(1, size // 32))
    return img


# ---------- メニュー構築 ----------

def build_menu(icon: pystray.Icon) -> pystray.Menu:
    return pystray.Menu(
        pystray.MenuItem("設定", lambda i, _: on_settings(i)),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("終了", on_quit),
    )


# ---------- イベントハンドラ ----------

def on_quit(icon: pystray.Icon, item) -> None:
    icon._stop_event.set()
    icon.stop()


def on_settings(icon: pystray.Icon) -> None:
    _settings_icon_ref[0] = icon
    _settings_trigger.set()


def show_settings_window(icon: pystray.Icon, parent: tk.Misc | None = None) -> None:
    """設定ウィンドウを表示する。"""
    BG  = "#2b2b2b"
    FG  = "#dcdcdc"
    ENT = "#3c3c3c"
    ACC = "#0078d4"
    LABEL_W = 20

    root = tk.Toplevel(parent) if parent is not None else tk.Tk()
    root.title("clipo - 設定")
    root.configure(bg=BG)
    root.resizable(True, False)
    root.attributes("-topmost", True)

    cfg = _load_config()

    _tk_vars: list[tk.Variable] = []

    def _close_settings(*_):
        for _v in _tk_vars:
            try:
                root.tk.globalunsetvar(_v._name)
            except Exception:
                pass
        _tk_vars.clear()
        root.destroy()

    def section_sep():
        tk.Frame(root, bg="#444444", height=1).pack(fill=tk.X, padx=16, pady=(8, 0))

    # ================================================================
    # OK / キャンセル（先に BOTTOM 確保）
    # ================================================================
    btn_frame = tk.Frame(root, bg=BG)
    btn_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=16, pady=12)

    # ================================================================
    # スタートアップ登録ボタン
    # ================================================================
    startup_frame = tk.Frame(root, bg=BG)
    startup_frame.pack(fill=tk.X, padx=16, pady=(14, 4))

    registered = is_startup_registered()
    startup_label_var = tk.StringVar(
        value="スタートアップから削除する" if registered else "スタートアップに登録する"
    )
    _tk_vars.append(startup_label_var)
    startup_btn_bg = "#c0392b" if registered else "#27ae60"

    startup_btn = tk.Button(
        startup_frame,
        textvariable=startup_label_var,
        bg=startup_btn_bg, fg="#ffffff", relief=tk.FLAT,
        activeforeground="#ffffff",
        font=("Yu Gothic UI", 9), padx=10,
    )
    startup_btn.pack(side=tk.LEFT)

    def toggle_startup():
        if is_startup_registered():
            unregister_startup()
            startup_label_var.set("スタートアップに登録する")
            startup_btn.config(bg="#27ae60", activebackground="#219150")
        else:
            register_startup()
            startup_label_var.set("スタートアップから削除する")
            startup_btn.config(bg="#c0392b", activebackground="#a93226")

    startup_btn.config(command=toggle_startup,
                       activebackground="#a93226" if registered else "#219150")

    section_sep()

    # ================================================================
    # 最大履歴件数 / PageJump件数
    # ================================================================
    def num_row(label: str, default_val: int) -> tk.StringVar:
        f = tk.Frame(root, bg=BG)
        f.pack(fill=tk.X, padx=16, pady=(8, 4))
        tk.Label(f, text=label, bg=BG, fg=FG,
                 font=("Yu Gothic UI", 9), anchor="w", width=LABEL_W).pack(side=tk.LEFT)
        var = tk.StringVar(value=str(default_val))
        _tk_vars.append(var)
        tk.Entry(f, textvariable=var, width=8,
                 bg=ENT, fg=FG, insertbackground=FG, relief=tk.FLAT,
                 font=("Yu Gothic UI", 9)).pack(side=tk.LEFT, ipady=3)
        return var

    max_hist_var  = num_row("最大履歴件数:",           cfg.get("max_history", MAX_HISTORY))
    page_jump_var = num_row("PageUp/Down の移動件数:", cfg.get("page_jump",   PAGE_JUMP))

    section_sep()

    # ================================================================
    # 履歴ファイルの保存先
    # ================================================================
    def path_row(label: str, default_path: Path):
        f = tk.Frame(root, bg=BG)
        f.pack(fill=tk.X, padx=16, pady=(8, 2))
        tk.Label(f, text=label, bg=BG, fg=FG,
                 font=("Yu Gothic UI", 9), anchor="w", width=LABEL_W).pack(side=tk.LEFT)
        var = tk.StringVar(value=str(default_path))
        _tk_vars.append(var)
        entry = tk.Entry(f, textvariable=var, bg=ENT, fg=FG,
                         insertbackground=FG, relief=tk.FLAT,
                         font=("Yu Gothic UI", 9))
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=3)

        def browse():
            cur = Path(var.get())
            chosen = tkinter.filedialog.askopenfilename(
                parent=root,
                title=f"{label}の選択（既存ファイルを選ぶと読み込みます）",
                initialdir=str(cur.parent) if cur.parent.exists() else str(default_path.parent),
                initialfile=cur.name,
                defaultextension=".json",
                filetypes=[("JSON ファイル", "*.json"), ("すべてのファイル", "*.*")],
            )
            if chosen:
                var.set(chosen)

        tk.Button(f, text="参照…", command=browse,
                  bg="#444444", fg=FG, relief=tk.FLAT,
                  activebackground="#555555", activeforeground=FG,
                  font=("Yu Gothic UI", 8), padx=6).pack(side=tk.LEFT, padx=(4, 0))
        return var

    hist_file_var = path_row(
        "履歴の保存先:",
        Path(cfg.get("history_file", str(HISTORY_FILE)))
    )
    tmpl_file_var = path_row(
        "定型文の保存先:",
        Path(cfg.get("templates_file", str(TEMPLATES_FILE)))
    )
    pins_file_var = path_row(
        "ピンの保存先:",
        Path(cfg.get("pins_file", str(PINS_FILE)))
    )

    def save():
        global MAX_HISTORY, PAGE_JUMP, HISTORY_FILE, TEMPLATES_FILE, PINS_FILE
        try:
            new_max  = int(max_hist_var.get())
            new_jump = int(page_jump_var.get())
            if new_max < 1 or new_jump < 1:
                raise ValueError
        except ValueError:
            tkinter.messagebox.showerror(
                "エラー", "件数は1以上の整数を入力してください。", parent=root)
            return

        # 変更前のパスを保持（既存ファイル読み込み判定に使用）
        old_history_file   = HISTORY_FILE
        old_templates_file = TEMPLATES_FILE
        old_pins_file      = PINS_FILE

        MAX_HISTORY    = new_max
        PAGE_JUMP      = new_jump
        HISTORY_FILE   = Path(hist_file_var.get())
        TEMPLATES_FILE = Path(tmpl_file_var.get())
        PINS_FILE      = Path(pins_file_var.get())

        # パスが変わり既存ファイルがある場合は上書きせず既存データを読み込む
        if HISTORY_FILE != old_history_file and HISTORY_FILE.exists():
            try:
                data = json.loads(HISTORY_FILE.read_text(encoding="utf-8"))
                with history_lock:
                    history.clear()
                    history.extend(data[-MAX_HISTORY:])
            except Exception as e:
                print(f"[clipo] 履歴の再読み込みエラー: {e}", file=sys.stderr)

        # 定型文・ピンはポップアップ起動時に新パスから自動読み込みされるため
        # ここでは変更通知のみ（既存ファイルを誤って上書きしないよう即時保存しない）
        if TEMPLATES_FILE != old_templates_file:
            print(f"[clipo] 定型文の保存先を変更: {TEMPLATES_FILE}", file=sys.stderr)
        if PINS_FILE != old_pins_file:
            print(f"[clipo] ピンの保存先を変更: {PINS_FILE}", file=sys.stderr)

        _save_config({
            "max_history":    new_max,
            "page_jump":      new_jump,
            "history_file":   str(HISTORY_FILE),
            "templates_file": str(TEMPLATES_FILE),
            "pins_file":      str(PINS_FILE),
        })
        _close_settings()

    tk.Button(btn_frame, text="キャンセル", command=_close_settings,
              bg="#444444", fg=FG, relief=tk.FLAT,
              activebackground="#555555", activeforeground=FG,
              font=("Yu Gothic UI", 9), padx=8).pack(side=tk.RIGHT)
    tk.Button(btn_frame, text="OK", command=save,
              bg=ACC, fg="#ffffff", relief=tk.FLAT,
              activebackground="#005fa3", activeforeground="#ffffff",
              font=("Yu Gothic UI", 9), padx=12).pack(side=tk.RIGHT, padx=4)

    # ウィンドウサイズをコンテンツに合わせて確定
    root.update_idletasks()
    sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
    dw = 480
    dh = root.winfo_reqheight()
    root.geometry(f"{dw}x{dh}+{(sw - dw)//2}+{(sh - dh)//2}")

    root.bind("<Escape>", _close_settings)
    if parent is not None:
        root.wait_window(root)
    else:
        root.mainloop()


# ---------- スタートアップ登録 ----------

_STARTUP_REG_KEY  = r"Software\Microsoft\Windows\CurrentVersion\Run"
_STARTUP_REG_NAME = "clipo"


def _startup_command() -> str:
    """スタートアップ登録に使うコマンド文字列を返す。
    exe としてコンパイル済みの場合は exe パスのみ、
    .py スクリプト実行中は pythonw.exe + スクリプトパスを使う。
    """
    if getattr(sys, "frozen", False):
        # PyInstaller 等でコンパイルされた exe
        return f'"{sys.executable}"'
    else:
        # .py スクリプトとして実行中
        pythonw = Path(sys.executable).with_name("pythonw.exe")
        script  = Path(__file__).resolve()
        return f'"{pythonw}" "{script}"'


def is_startup_registered() -> bool:
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, _STARTUP_REG_KEY) as key:
            winreg.QueryValueEx(key, _STARTUP_REG_NAME)
            return True
    except OSError:
        return False


def register_startup() -> None:
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, _STARTUP_REG_KEY,
                        0, winreg.KEY_SET_VALUE) as key:
        winreg.SetValueEx(key, _STARTUP_REG_NAME, 0, winreg.REG_SZ, _startup_command())


def unregister_startup() -> None:
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, _STARTUP_REG_KEY,
                            0, winreg.KEY_SET_VALUE) as key:
            winreg.DeleteValue(key, _STARTUP_REG_NAME)
    except OSError:
        pass


# ---------- マルチモニター対応 ----------

class _MONITORINFO(ctypes.Structure):
    _fields_ = [
        ("cbSize",    ctypes.c_ulong),
        ("rcMonitor", ctypes.wintypes.RECT),
        ("rcWork",    ctypes.wintypes.RECT),
        ("dwFlags",   ctypes.c_ulong),
    ]


def _get_monitor_work_area(x: int, y: int) -> tuple[int, int, int, int]:
    """(x, y) を含むモニターの作業領域 (left, top, right, bottom) を返す。"""
    MONITOR_DEFAULTTONEAREST = 2
    pt   = ctypes.wintypes.POINT(x, y)
    hmon = ctypes.windll.user32.MonitorFromPoint(pt, MONITOR_DEFAULTTONEAREST)
    info = _MONITORINFO()
    info.cbSize = ctypes.sizeof(_MONITORINFO)
    ctypes.windll.user32.GetMonitorInfoW(hmon, ctypes.byref(info))
    r = info.rcWork
    return r.left, r.top, r.right, r.bottom


# ---------- テンプレート展開 ----------

# ---- <clipo:N> — 履歴参照 ----
# タグ構文: <clipo:N>  N=インデックス（0=最新）。<clipo> は <clipo:0> の省略形。
_CLIPO_TAG = re.compile(r'<clipo(?::(\d+))?>', re.IGNORECASE)

# ---- <clipo_DATE>format</clipo_DATE> — 日時フォーマット ----
_DATE_TAG = re.compile(r'<clipo_DATE>(.*?)</clipo_DATE>', re.IGNORECASE | re.DOTALL)

# ---- <MStatus> — 月中間/月末自動判別 ----
_MSTATUS_TAG = re.compile(r'<MStatus>', re.IGNORECASE)

# フォーマットトークン（長いものを先に並べて二重マッチを防ぐ）
_DATE_TOKENS = re.compile(r'yyyy|yy|mm|m|dd|d|hh|h|nn|n|ss|s', re.IGNORECASE)


def _apply_date_format(fmt: str, dt: datetime) -> str:
    """日時フォーマット文字列を展開して返す。
    トークン一覧:
      yyyy/yy  年(4桁/下2桁)   mm/m  月(ゼロ埋め/そのまま)
      dd/d     日(ゼロ埋め/そのまま)   hh/h  時(ゼロ埋め/そのまま)
      nn/n     分(ゼロ埋め/そのまま)   ss/s  秒(ゼロ埋め/そのまま)
    """
    def repl(m: re.Match) -> str:
        t = m.group(0).lower()
        if t == "yyyy": return f"{dt.year:04d}"
        if t == "yy":   return f"{dt.year % 100:02d}"
        if t == "mm":   return f"{dt.month:02d}"
        if t == "m":    return str(dt.month)
        if t == "dd":   return f"{dt.day:02d}"
        if t == "d":    return str(dt.day)
        if t == "hh":   return f"{dt.hour:02d}"
        if t == "h":    return str(dt.hour)
        if t == "nn":   return f"{dt.minute:02d}"
        if t == "n":    return str(dt.minute)
        if t == "ss":   return f"{dt.second:02d}"
        if t == "s":    return str(dt.second)
        return m.group(0)
    return _DATE_TOKENS.sub(repl, fmt)


def _get_mstatus(dt: datetime) -> str:
    """日付から月次ステータスを返す。
    第1週(1〜7日)・第4週以降(22日〜): 月末
    第2〜3週(8〜21日): 中間
    """
    return "中間" if 8 <= dt.day <= 21 else "月末"


def _interpolate(text: str, hist: list[dict]) -> str:
    """定型文タグを展開して返す。"""
    # <clipo:N> — 履歴参照
    def replace_clipo(m: re.Match) -> str:
        idx = int(m.group(1)) if m.group(1) is not None else 0
        if 0 <= idx < len(hist):
            return hist[idx]["text"]
        return m.group(0)
    text = _CLIPO_TAG.sub(replace_clipo, text)

    # <clipo_DATE>format</clipo_DATE> — 日時
    now = datetime.now()
    text = _DATE_TAG.sub(lambda m: _apply_date_format(m.group(1), now), text)

    # <MStatus> — 月中間/月末自動判別
    text = _MSTATUS_TAG.sub(lambda _: _get_mstatus(now), text)

    return text


# ---------- タグ挿入ピッカー ----------

_TAG_ITEMS: list[tuple[str, str | None]] = [
    # (表示テキスト, 挿入タグ)  ※ タグが None のものはセクションヘッダー
    ("― 履歴参照 ―",                          None),
    ("最新の履歴",                              "<clipo>"),
    ("2番目の履歴",                             "<clipo:1>"),
    ("3番目の履歴",                             "<clipo:2>"),
    ("― 日時 ―",                              None),
    ("年/月/日  →  2026/04/07",               "<clipo_DATE>yyyy/mm/dd</clipo_DATE>"),
    ("年/月/日 時:分:秒  →  2026/04/07 15:22:40", "<clipo_DATE>yyyy/mm/dd hh:nn:ss</clipo_DATE>"),
    ("年月日（漢字）  →  2026年4月7日",          "<clipo_DATE>yyyy年m月d日</clipo_DATE>"),
    ("時:分  →  15:22",                        "<clipo_DATE>hh:nn</clipo_DATE>"),
    ("時:分:秒  →  15:22:40",                  "<clipo_DATE>hh:nn:ss</clipo_DATE>"),
    ("yymmdd  →  260407",                      "<clipo_DATE>yymmdd</clipo_DATE>"),
    ("下2桁年  →  26",                         "<clipo_DATE>yy</clipo_DATE>"),
    ("月（ゼロ埋めなし）  →  4",               "<clipo_DATE>m</clipo_DATE>"),
    ("― 月次ステータス ―",                     None),
    ("中間/月末 自動判別  (1〜7日・22日〜:月末 / 8〜21日:中間)", "<MStatus>"),
]


def _show_tag_picker(parent: tk.Misc, text_box: tk.Text) -> None:
    """タグ候補ウィンドウを表示し、選択されたタグを text_box のカーソル位置に挿入する。"""
    BG      = "#1e1e1e"
    FG      = "#dcdcdc"
    HDR_FG  = "#888888"
    SEL_BG  = "#0078d4"
    BORDER  = "#555555"

    win = tk.Toplevel(parent)
    win.overrideredirect(True)
    win.attributes("-topmost", True)
    win.configure(bg=BORDER)

    inner = tk.Frame(win, bg=BG)
    inner.pack(padx=1, pady=1, fill=tk.BOTH, expand=True)

    lb = tk.Listbox(
        inner,
        bg=BG, fg=FG,
        selectbackground=SEL_BG, selectforeground="#ffffff",
        activestyle="none",
        font=("Yu Gothic UI", 9),
        bd=0, highlightthickness=0,
        width=46,
    )
    lb.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)

    # ヘッダー行のインデックスを記録（選択不可）
    header_indices: set[int] = set()
    for i, (label, tag) in enumerate(_TAG_ITEMS):
        lb.insert(tk.END, f"  {label}" if tag else label)
        if tag is None:
            lb.itemconfig(i, fg=HDR_FG, selectbackground=BG, selectforeground=HDR_FG)
            header_indices.add(i)

    def insert_tag(idx: int) -> None:
        if idx in header_indices:
            return
        tag = _TAG_ITEMS[idx][1]
        if tag is None:
            return
        text_box.insert(tk.INSERT, tag)
        text_box.focus_set()
        win.destroy()

    lb.bind("<ButtonRelease-1>", lambda ev: insert_tag(lb.nearest(ev.y)))
    lb.bind("<Return>",          lambda _: insert_tag(lb.curselection()[0] if lb.curselection() else -1))
    lb.bind("<Escape>",          lambda _: win.destroy())
    win.bind("<FocusOut>",       lambda _: win.after(100, lambda: win.destroy() if win.winfo_exists() else None))

    # ウィンドウ位置：parent の中央下寄り、画面端補正
    win.update_idletasks()
    pw = parent.winfo_rootx()
    py = parent.winfo_rooty()
    ph = parent.winfo_height()
    ww = win.winfo_reqwidth()
    wh = win.winfo_reqheight()
    sw = win.winfo_screenwidth()
    sh = win.winfo_screenheight()
    x = max(0, min(pw + 8, sw - ww - 4))
    y = min(py + ph - wh - 8, sh - wh - 4)
    win.geometry(f"+{x}+{y}")
    lb.focus_set()


# ---------- 編集ダイアログ ----------

def _edit_dialog(
    parent: tk.Misc,
    title: str,
    initial_text: str,
    on_save,                      # callable(name: str, text: str)
    initial_name: str = "",
    has_name: bool = False,
    show_tag_hint: bool = False,  # True のとき <clipo:N> 構文ヒントを表示
) -> None:
    """テキスト（＋オプションで名前）を編集するモーダルダイアログ。"""
    BG   = "#2b2b2b"
    FG   = "#dcdcdc"
    ENT  = "#3c3c3c"
    ACC  = "#0078d4"

    dlg = tk.Toplevel(parent)
    dlg.title(title)
    dlg.configure(bg=BG)
    dlg.attributes("-topmost", True)
    dlg.resizable(True, True)
    dlg.grab_set()

    # 画面中央に配置
    dlg.update_idletasks()
    sw, sh = dlg.winfo_screenwidth(), dlg.winfo_screenheight()
    dw, dh = 480, 300
    dlg.geometry(f"{dw}x{dh}+{(sw-dw)//2}+{(sh-dh)//2}")

    pad = {"padx": 8, "pady": 4}

    name_var = tk.StringVar(value=initial_name)
    if has_name:
        tk.Label(dlg, text="名前:", bg=BG, fg=FG, font=("Yu Gothic UI", 9)).pack(anchor="w", **pad)
        name_entry = tk.Entry(dlg, textvariable=name_var,
                              bg=ENT, fg=FG, insertbackground=FG,
                              relief=tk.FLAT, font=("Yu Gothic UI", 10))
        name_entry.pack(fill=tk.X, **pad, ipady=3)

    def save():
        on_save(name_var.get().strip(), text_box.get("1.0", "end-1c"))
        dlg.destroy()

    # --- BOTTOM から順に確保: ボタン → テキストボックス ---

    # ① ボタン（最下部）
    btn_frame = tk.Frame(dlg, bg=BG)
    btn_frame.pack(side=tk.BOTTOM, fill=tk.X, **pad)
    tk.Button(
        btn_frame, text="キャンセル", command=dlg.destroy,
        bg="#444444", fg=FG, relief=tk.FLAT,
        activebackground="#555555", activeforeground=FG,
        font=("Yu Gothic UI", 9), padx=8,
    ).pack(side=tk.RIGHT, pady=2)
    tk.Button(
        btn_frame, text="OK", command=save,
        bg=ACC, fg="#ffffff", relief=tk.FLAT,
        activebackground="#005fa3", activeforeground="#ffffff",
        font=("Yu Gothic UI", 9), padx=12,
    ).pack(side=tk.RIGHT, padx=4, pady=2)

    # タグ挿入ボタン（定型文編集時のみ）
    if show_tag_hint:
        def open_tag_picker():
            _show_tag_picker(dlg, text_box)
        tk.Button(
            btn_frame, text="タグを挿入 ▾", command=open_tag_picker,
            bg="#444444", fg=FG, relief=tk.FLAT,
            activebackground="#555555", activeforeground=FG,
            font=("Yu Gothic UI", 9), padx=8,
        ).pack(side=tk.LEFT, pady=2)

    # ② テキストボックス（残り領域をすべて使う）
    text_frame = tk.Frame(dlg, bg=ENT)
    text_frame.pack(fill=tk.BOTH, expand=True, **pad)
    sb = tk.Scrollbar(text_frame)
    sb.pack(side=tk.RIGHT, fill=tk.Y)
    text_box = tk.Text(
        text_frame,
        bg=ENT, fg=FG, insertbackground=FG,
        relief=tk.FLAT, font=("Yu Gothic UI", 10),
        wrap=tk.WORD, undo=True,
        yscrollcommand=sb.set,
    )
    sb.config(command=text_box.yview)
    text_box.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
    text_box.insert("1.0", initial_text)
    text_box.focus_set()

    dlg.bind("<Control-Return>", lambda _: save())
    dlg.bind("<Escape>", lambda _: dlg.destroy())
    dlg.wait_window()


# ---------- 履歴ポップアップ ----------

TOOLTIP_DELAY = 500   # ツールチップ表示までのディレイ（ms）
TOOLTIP_WRAP  = 600   # ツールチップの折り返し幅（px）

def _bind_dnd_reorder(
    listbox: tk.Listbox,
    data: list,
    filtered: list,
    dnd_active: list,   # [bool] — copy_template と共有するフラグ
    on_reorder,         # callable()
    normal_bg: str = "#2b2b2b",
    target_bg: str = "#3d3d3d",
) -> None:
    """リストボックスの行をドラッグ&ドロップで並び替える。
    filtered が空でない場合は filtered 経由で data を操作する。
    """
    THRESHOLD = 5   # ドラッグ開始とみなす移動距離 (px)
    state: dict = {"start_y": 0, "src": -1, "last_tgt": -1, "dragging": False}

    def _reset_bg(idx: int) -> None:
        if 0 <= idx < listbox.size():
            listbox.itemconfig(idx, bg=normal_bg)

    def on_press(event: tk.Event) -> None:
        state["src"] = listbox.nearest(event.y)
        state["start_y"] = event.y
        state["last_tgt"] = state["src"]
        state["dragging"] = False
        dnd_active[0] = False

    def on_motion(event: tk.Event) -> None:
        if state["src"] < 0:
            return
        if not state["dragging"]:
            if abs(event.y - state["start_y"]) > THRESHOLD:
                state["dragging"] = True
                dnd_active[0] = True
                listbox.config(cursor="sb_v_double_arrow")
        if state["dragging"]:
            tgt = listbox.nearest(event.y)
            if tgt != state["last_tgt"]:
                _reset_bg(state["last_tgt"])
                if 0 <= tgt < listbox.size():
                    listbox.itemconfig(tgt, bg=target_bg)
                state["last_tgt"] = tgt

    def on_release(event: tk.Event) -> None:
        if not state["dragging"]:
            state["src"] = -1
            return
        tgt = listbox.nearest(event.y)
        _reset_bg(state["last_tgt"])
        listbox.config(cursor="")
        src = state["src"]
        state["src"] = -1
        state["dragging"] = False
        # フラグをわずかに遅らせてリセット（copy_template の後に走るため）
        listbox.after(50, lambda: dnd_active.__setitem__(0, False))

        if src == tgt or src < 0:
            return
        # filtered 経由で data の実インデックスを取得
        fi = filtered if filtered else list(range(len(data)))
        if not (0 <= src < len(fi) and 0 <= tgt < len(fi)):
            return
        src_real, tgt_real = fi[src], fi[tgt]
        item = data.pop(src_real)
        data.insert(tgt_real, item)
        on_reorder()

    listbox.bind("<ButtonPress-1>",   on_press,    add=True)
    listbox.bind("<B1-Motion>",       on_motion,   add=True)
    listbox.bind("<ButtonRelease-1>", on_release,  add=True)


def _bind_hover_highlight(
    listbox: tk.Listbox,
    normal_fg: str = "#666666",
    hover_fg:  str = "#dcdcdc",
) -> None:
    """マウスが乗った行だけ hover_fg に、それ以外は normal_fg にする。"""
    last: list[int] = [-1]

    def _set(idx: int, fg: str) -> None:
        if 0 <= idx < listbox.size():
            listbox.itemconfig(idx, fg=fg)

    def on_motion(event: tk.Event) -> None:
        idx = listbox.nearest(event.y)
        if idx == last[0]:
            return
        _set(last[0], normal_fg)
        _set(idx,      hover_fg)
        last[0] = idx

    def on_leave(*_) -> None:
        _set(last[0], normal_fg)
        last[0] = -1

    listbox.bind("<Motion>", on_motion, add=True)
    listbox.bind("<Leave>",  on_leave,  add=True)


def _normalize_for_display(text: str) -> str:
    """ツールチップ表示用にタブ・CR等を正規化する。"""
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = text.replace("\t", "    ")
    return text


def _bind_listbox_tooltip(
    root: tk.Tk,
    listbox: tk.Listbox,
    filtered: list,
    history_data: list,
) -> None:
    """リストボックスの各行にホバーすると全文ツールチップを表示する。"""
    tip: tk.Toplevel | None = None
    after_id: str | None = None
    last_idx: int = -1

    def hide() -> None:
        nonlocal tip, after_id, last_idx
        if after_id:
            root.after_cancel(after_id)
            after_id = None
        if tip:
            tip.destroy()
            tip = None
        last_idx = -1

    def show(idx: int, rx: int, ry: int) -> None:
        nonlocal tip
        if tip:
            tip.destroy()
            tip = None
        if idx < 0 or idx >= len(filtered):
            return
        text = _normalize_for_display(history_data[filtered[idx]]["text"])

        tip = tk.Toplevel(root)
        tip.overrideredirect(True)
        tip.attributes("-topmost", True)

        lbl = tk.Label(
            tip, text=text,
            justify=tk.LEFT,
            wraplength=TOOLTIP_WRAP,
            font=("Yu Gothic UI", 9),
            bg="#1e1e1e", fg="#dcdcdc",
            padx=8, pady=6,
            relief=tk.FLAT,
        )
        lbl.pack()

        # 外枠
        tip.configure(bg="#555555")
        lbl.pack(padx=1, pady=1)

        tip.update_idletasks()
        tw = tip.winfo_width()
        th = tip.winfo_height()
        sw = root.winfo_screenwidth()
        sh = root.winfo_screenheight()
        # カーソルの右下、画面端を超えないよう調整
        tx = min(rx + 16, sw - tw - 4)
        ty = min(ry + 16, sh - th - 4)
        tip.geometry(f"+{tx}+{ty}")

    def schedule_show(idx: int, rx: int, ry: int) -> None:
        nonlocal after_id
        if after_id:
            root.after_cancel(after_id)
        after_id = root.after(TOOLTIP_DELAY, lambda: show(idx, rx, ry))

    def on_motion(event: tk.Event) -> None:
        nonlocal last_idx
        idx = listbox.nearest(event.y)
        if idx == last_idx:
            return
        last_idx = idx
        schedule_show(idx, event.x_root, event.y_root)

    def on_leave(*_) -> None:
        hide()

    listbox.bind("<Motion>",   on_motion,          add=True)
    listbox.bind("<Leave>",    on_leave,            add=True)
    listbox.bind("<Button-3>", lambda *_: hide(),   add=True)
    return hide


def _bind_template_tooltip(
    root: tk.Tk,
    listbox: tk.Listbox,
    filtered: list,
    templates: list,
    hist: list,
) -> None:
    """定型文リストのツールチップ。タグ展開後のプレビューも合わせて表示する。"""
    tip: tk.Toplevel | None = None
    after_id: str | None = None
    last_idx: int = -1

    def hide() -> None:
        nonlocal tip, after_id, last_idx
        if after_id:
            root.after_cancel(after_id)
            after_id = None
        if tip:
            tip.destroy()
            tip = None
        last_idx = -1

    def show(idx: int, rx: int, ry: int) -> None:
        nonlocal tip
        if tip:
            tip.destroy()
            tip = None
        if idx < 0 or idx >= len(filtered):
            return
        tmpl = templates[filtered[idx]]
        raw  = _normalize_for_display(tmpl.get("text", ""))
        interpolated = _normalize_for_display(_interpolate(tmpl.get("text", ""), hist))
        has_tags = interpolated != raw

        tip = tk.Toplevel(root)
        tip.overrideredirect(True)
        tip.attributes("-topmost", True)
        tip.configure(bg="#555555")

        inner = tk.Frame(tip, bg="#1e1e1e", padx=8, pady=6)
        inner.pack(padx=1, pady=1, fill=tk.BOTH, expand=True)

        # 原文
        tk.Label(inner, text=raw, justify=tk.LEFT,
                 wraplength=TOOLTIP_WRAP,
                 font=("Yu Gothic UI", 9),
                 bg="#1e1e1e", fg="#dcdcdc").pack(anchor="w")

        # タグが含まれる場合は展開後プレビューを追加
        if has_tags:
            tk.Frame(inner, bg="#444444", height=1).pack(fill=tk.X, pady=(6, 0))
            tk.Label(inner, text="展開後:", justify=tk.LEFT,
                     font=("Yu Gothic UI", 8),
                     bg="#1e1e1e", fg="#888888").pack(anchor="w", pady=(4, 0))
            tk.Label(inner, text=interpolated, justify=tk.LEFT,
                     wraplength=TOOLTIP_WRAP,
                     font=("Yu Gothic UI", 9),
                     bg="#1e1e1e", fg="#7ec8e3").pack(anchor="w")

        tip.update_idletasks()
        sw = root.winfo_screenwidth()
        sh = root.winfo_screenheight()
        tx = min(rx + 16, sw - tip.winfo_width()  - 4)
        ty = min(ry + 16, sh - tip.winfo_height() - 4)
        tip.geometry(f"+{tx}+{ty}")

    def schedule_show(idx: int, rx: int, ry: int) -> None:
        nonlocal after_id
        if after_id:
            root.after_cancel(after_id)
        after_id = root.after(TOOLTIP_DELAY, lambda: show(idx, rx, ry))

    def on_motion(event: tk.Event) -> None:
        nonlocal last_idx
        idx = listbox.nearest(event.y)
        if idx == last_idx:
            return
        last_idx = idx
        schedule_show(idx, event.x_root, event.y_root)

    def on_leave(*_) -> None:
        hide()

    listbox.bind("<Motion>",   on_motion,          add=True)
    listbox.bind("<Leave>",    on_leave,            add=True)
    listbox.bind("<Button-3>", lambda *_: hide(),   add=True)
    return hide


def _setup_resize(root: tk.Tk, min_w: int = 200, min_h: int = 120) -> None:
    """overrideredirect ウィンドウの縁をドラッグしてリサイズできるようにする。"""
    BORDER = 6
    state: dict = {"dragging": False, "mode": None}

    CURSOR_MAP = {
        "e": "size_we",  "w": "size_we",
        "s": "size_ns",  "n": "size_ns",
        "se": "size_nw_se", "nw": "size_nw_se",
        "ne": "size_ne_sw", "sw": "size_ne_sw",
    }

    def hit_test(rx: int, ry: int) -> str | None:
        wx, wy = root.winfo_x(), root.winfo_y()
        ww, wh = root.winfo_width(), root.winfo_height()
        x, y = rx - wx, ry - wy
        if not (0 <= x <= ww and 0 <= y <= wh):
            return None
        on_e = x >= ww - BORDER
        on_w = x <= BORDER
        on_s = y >= wh - BORDER
        on_n = y <= BORDER
        if on_n and on_w: return "nw"
        if on_n and on_e: return "ne"
        if on_s and on_w: return "sw"
        if on_s and on_e: return "se"
        if on_e: return "e"
        if on_w: return "w"
        if on_s: return "s"
        if on_n: return "n"
        return None

    def on_motion(event: tk.Event) -> None:
        if state["dragging"]:
            return
        mode = hit_test(event.x_root, event.y_root)
        root.config(cursor=CURSOR_MAP.get(mode or "", ""))

    def on_press(event: tk.Event) -> None:
        mode = hit_test(event.x_root, event.y_root)
        if not mode:
            return
        state.update({
            "dragging": True, "mode": mode,
            "sx": event.x_root, "sy": event.y_root,
            "sw": root.winfo_width(), "sh": root.winfo_height(),
            "wx": root.winfo_x(),    "wy": root.winfo_y(),
        })
        root.grab_set()  # ドラッグ中はすべてのイベントをこのウィンドウで受け取る

    def on_drag(event: tk.Event) -> None:
        if not state["dragging"]:
            return
        dx = event.x_root - state["sx"]
        dy = event.y_root - state["sy"]
        mode = state["mode"]
        nw, nh = state["sw"], state["sh"]
        nx, ny = state["wx"], state["wy"]
        if "e" in mode: nw = max(min_w, nw + dx)
        if "w" in mode:
            nw = max(min_w, nw - dx)
            nx = state["wx"] + state["sw"] - nw
        if "s" in mode: nh = max(min_h, nh + dy)
        if "n" in mode:
            nh = max(min_h, nh - dy)
            ny = state["wy"] + state["sh"] - nh
        root.geometry(f"{nw}x{nh}+{nx}+{ny}")

    def on_release(event: tk.Event) -> None:
        if state["dragging"]:
            state["dragging"] = False
            state["mode"] = None
            root.grab_release()
            root.config(cursor=CURSOR_MAP.get(hit_test(event.x_root, event.y_root) or "", ""))

    root.bind_all("<Motion>",          on_motion,  add=True)
    root.bind_all("<ButtonPress-1>",   on_press,   add=True)
    root.bind_all("<B1-Motion>",       on_drag,    add=True)
    root.bind_all("<ButtonRelease-1>", on_release, add=True)

# ---------- ポップアップ事前初期化 ----------
_popup_trigger    = threading.Event()  # ホットキー検出 → ワーカーへのシグナル
_popup_active     = threading.Event()  # ポップアップ実行中フラグ（多重起動防止）
_popup_icon_ref: list   = [None]       # icon の参照渡し用
_settings_trigger = threading.Event()  # 設定ウィンドウ表示トリガー
_settings_icon_ref: list = [None]      # 設定ウィンドウ用 icon 参照


def show_popup(icon, _root: tk.Tk | None = None) -> None:
    """コンテキストメニュー風ポップアップ。タブで履歴／定型文を切り替える。"""
    with history_lock:
        current_history = list(history)
    templates = load_templates()
    pins = load_pins()

    # ---- カーソル位置を取得 ----
    root = _root or tk.Tk()
    cx = root.winfo_pointerx()
    cy = root.winfo_pointery()

    # ---- ウィンドウ装飾を除去 ----
    root.overrideredirect(True)
    root.attributes("-topmost", True)
    root.attributes("-toolwindow", True)
    root.configure(bg="#2b2b2b")

    # ---- 外枠 ----
    outer = tk.Frame(root, bg="#555555", bd=1, relief=tk.FLAT)
    outer.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)

    # ================================================================
    # タブバー
    # ================================================================
    TAB_BG_ACTIVE   = "#2b2b2b"
    TAB_BG_INACTIVE = "#1e1e1e"
    TAB_FG_ACTIVE   = "#dcdcdc"
    TAB_FG_INACTIVE = "#888888"
    ACCENT          = "#0078d4"

    tab_bar = tk.Frame(outer, bg=TAB_BG_INACTIVE)
    tab_bar.pack(fill=tk.X, side=tk.TOP)

    # アクティブタブ下線用キャンバス
    tab_indicator = tk.Canvas(outer, height=2, bg="#555555", highlightthickness=0)
    tab_indicator.pack(fill=tk.X, side=tk.TOP)

    # ---- コンテンツ領域（タブごとのページを重ねる） ----
    content_area = tk.Frame(outer, bg="#2b2b2b")
    content_area.pack(fill=tk.BOTH, expand=True)

    # show_popup 内で生成する tk.Variable を追跡し、close 時に明示解放する
    # (別スレッド内で mainloop 終了後に GC が __del__ を呼ぶと tcl エラーになるため)
    _tk_vars: list[tk.Variable] = []

    # 編集ダイアログが開いている間は外クリックでポップアップを閉じないためのカウンター
    _dialog_open = [0]

    def close(*_):
        for _v in _tk_vars:
            try:
                root.tk.globalunsetvar(_v._name)
            except Exception:
                pass
        _tk_vars.clear()
        root.destroy()

    # ================================================================
    # 履歴ページ
    # ================================================================
    history_page = tk.Frame(content_area, bg="#2b2b2b")

    search_var = tk.StringVar()
    _tk_vars.append(search_var)
    search_entry = tk.Entry(
        history_page, textvariable=search_var,
        font=("Yu Gothic UI", 10),
        bg="#3c3c3c", fg="#dcdcdc", insertbackground="#dcdcdc",
        relief=tk.FLAT, bd=0,
    )
    search_entry.pack(fill=tk.X, padx=6, pady=(6, 3), ipady=4)
    tk.Frame(history_page, bg="#555555", height=1).pack(fill=tk.X)

    h_listbox = tk.Listbox(
        history_page,
        font=("Yu Gothic UI", 10),
        selectmode=tk.SINGLE, activestyle="none",
        bd=0, highlightthickness=0,
        bg="#2b2b2b", fg="#666666",
        selectbackground=ACCENT, selectforeground="#ffffff",
    )
    h_listbox.pack(fill=tk.BOTH, expand=True)

    h_status_var = tk.StringVar()
    _tk_vars.append(h_status_var)
    tk.Label(
        history_page, textvariable=h_status_var, anchor="w",
        font=("Yu Gothic UI", 8), fg="#888888", bg="#2b2b2b",
    ).pack(fill=tk.X, padx=6, pady=(1, 4))

    filtered: list[int] = []

    def refresh_history(*_):
        query = search_var.get().lower()
        h_listbox.delete(0, tk.END)
        filtered.clear()
        pinned_texts = {p["text"] for p in pins}  # ピン済みテキストをセットで高速参照
        for i, entry in enumerate(current_history):
            text = entry["text"]
            if query and query not in text.lower():
                continue
            filtered.append(i)
            display = text.replace("\n", " ").replace("\r", "")
            if len(display) > 72:
                display = display[:72] + "…"
            prefix = "📌 " if text in pinned_texts else "  "
            h_listbox.insert(tk.END, f"{prefix}{display}")
        if filtered:
            h_listbox.selection_set(0)
            h_listbox.activate(0)
        h_status_var.set(f"  {len(filtered)}/{len(current_history)} 件  クリック/Enter:コピー  Esc:閉じる")

    def copy_history(*_):
        sel = h_listbox.curselection()
        if not sel:
            return
        text = current_history[filtered[sel[0]]]["text"]
        root.destroy()
        pyperclip.copy(text)
        icon.update_menu()

    def on_history_key(event):
        if event.keysym == "Return":
            copy_history()
        elif event.keysym == "Escape":
            close()
        elif event.keysym in ("Left", "Right"):
            _switch_tab_relative(+1 if event.keysym == "Right" else -1)
        elif event.keysym == "BackSpace" or (len(event.char) == 1 and event.char.isprintable()):
            search_entry.focus_set()
            if event.keysym == "BackSpace":
                search_var.set(search_var.get()[:-1])
            else:
                search_entry.insert(tk.END, event.char)

    def _move_h_selection(delta: int) -> None:
        """h_listbox の選択を delta 分移動（フォーカスは検索バーのまま）。"""
        size = h_listbox.size()
        if size == 0:
            return
        sel = h_listbox.curselection()
        cur = sel[0] if sel else (-1 if delta > 0 else size)
        new = max(0, min(size - 1, cur + delta))
        h_listbox.selection_clear(0, tk.END)
        h_listbox.selection_set(new)
        h_listbox.activate(new)
        h_listbox.see(new)

    def on_search_key(event):
        if event.keysym == "Return":
            copy_history()
        elif event.keysym == "Escape":
            close()
        elif event.keysym == "Down":
            _move_h_selection(+1)
        elif event.keysym == "Up":
            _move_h_selection(-1)
        elif event.keysym == "Next":   # PageDown
            _move_h_selection(+PAGE_JUMP)
        elif event.keysym == "Prior":  # PageUp
            _move_h_selection(-PAGE_JUMP)
        elif event.keysym in ("Left", "Right"):
            _switch_tab_relative(+1 if event.keysym == "Right" else -1)
            return "break"  # Entry のカーソル移動を抑制

    # ---- 履歴: 右クリックメニュー ----
    def _history_edit(hist_idx: int) -> None:
        def on_save(_, text: str) -> None:
            with history_lock:
                history[hist_idx]["text"] = text
                current_history[hist_idx]["text"] = text
            save_history()
            refresh_history()
        _dialog_open[0] += 1
        try:
            _edit_dialog(root, "履歴を編集", current_history[hist_idx]["text"], on_save)
        finally:
            _dialog_open[0] -= 1

    def _history_delete(hist_idx: int) -> None:
        with history_lock:
            history.pop(hist_idx)
            current_history.pop(hist_idx)
        save_history()
        refresh_history()

    def _history_format(hist_idx: int, mode: str) -> None:
        src = current_history[hist_idx]["text"]
        if mode == "angle":
            result = f"<{src}>"
        elif mode == "dquote":
            result = f'"{src}"'
        else:  # blockquote
            result = "\n".join(f"> {line}" for line in src.splitlines())
        with history_lock:
            history[hist_idx]["text"] = result
            current_history[hist_idx]["text"] = result
        save_history()
        refresh_history()

    # ツールチップの hide 関数を後から登録するリスト（コンテキストメニューと共有）
    _tooltip_hiders: list = []

    def _hide_all_tooltips() -> None:
        for _h in _tooltip_hiders:
            _h()

    def _history_context_menu(event: tk.Event) -> None:
        _hide_all_tooltips()
        idx = h_listbox.nearest(event.y)
        has_item = 0 <= idx < len(filtered)
        if has_item:
            h_listbox.selection_clear(0, tk.END)
            h_listbox.selection_set(idx)
            h_listbox.activate(idx)
        hist_idx = filtered[idx] if has_item else -1

        menu = tk.Menu(root, tearoff=0,
                       bg="#2b2b2b", fg="#dcdcdc",
                       activebackground="#0078d4", activeforeground="#ffffff",
                       font=("Yu Gothic UI", 9))
        if has_item:
            is_pinned = any(p["text"] == current_history[hist_idx]["text"] for p in pins)
            pin_label = "ピン留めを解除" if is_pinned else "ピン留めする"
            menu.add_command(label=pin_label,
                             command=lambda: _toggle_pin_from_history(hist_idx))
            menu.add_separator()
            menu.add_command(label="編集…",
                             command=lambda: _history_edit(hist_idx))
            menu.add_command(label="削除",
                             command=lambda: _history_delete(hist_idx))
            menu.add_separator()
            fmt = tk.Menu(menu, tearoff=0,
                          bg="#2b2b2b", fg="#dcdcdc",
                          activebackground="#0078d4", activeforeground="#ffffff",
                          font=("Yu Gothic UI", 9))
            fmt.add_command(label="< > でくくる",
                            command=lambda: _history_format(hist_idx, "angle"))
            fmt.add_command(label='" " でくくる',
                            command=lambda: _history_format(hist_idx, "dquote"))
            fmt.add_command(label="引用文として整形 (> …)",
                            command=lambda: _history_format(hist_idx, "blockquote"))
            menu.add_cascade(label="整形", menu=fmt)
        else:
            menu.add_command(label="（項目を右クリックしてください）", state=tk.DISABLED)

        menu.tk_popup(event.x_root, event.y_root)

    h_listbox.bind("<Button-3>", _history_context_menu)

    search_var.trace_add("write", refresh_history)
    h_listbox.bind("<ButtonRelease-1>", copy_history)
    h_listbox.bind("<Return>", copy_history)
    h_listbox.bind("<Key>", on_history_key)
    search_entry.bind("<Key>", on_search_key)
    _bind_hover_highlight(h_listbox)
    _tooltip_hiders.append(_bind_listbox_tooltip(root, h_listbox, filtered, current_history))

    # ================================================================
    # 定型文ページ
    # ================================================================
    template_page = tk.Frame(content_area, bg="#2b2b2b")

    t_search_var = tk.StringVar()
    _tk_vars.append(t_search_var)
    t_search_entry = tk.Entry(
        template_page, textvariable=t_search_var,
        font=("Yu Gothic UI", 10),
        bg="#3c3c3c", fg="#dcdcdc", insertbackground="#dcdcdc",
        relief=tk.FLAT, bd=0,
    )
    t_search_entry.pack(fill=tk.X, padx=6, pady=(6, 3), ipady=4)
    tk.Frame(template_page, bg="#555555", height=1).pack(fill=tk.X)

    t_listbox = tk.Listbox(
        template_page,
        font=("Yu Gothic UI", 10),
        selectmode=tk.SINGLE, activestyle="none",
        bd=0, highlightthickness=0,
        bg="#2b2b2b", fg="#666666",
        selectbackground=ACCENT, selectforeground="#ffffff",
    )
    t_listbox.pack(fill=tk.BOTH, expand=True)

    t_status_var = tk.StringVar()
    _tk_vars.append(t_status_var)
    tk.Label(
        template_page, textvariable=t_status_var, anchor="w",
        font=("Yu Gothic UI", 8), fg="#888888", bg="#2b2b2b",
    ).pack(fill=tk.X, padx=6, pady=(1, 4))

    t_filtered: list[int] = []

    def refresh_templates(*_):
        _hide_all_tooltips()  # リスト再構築前にツールチップ状態(last_idx含む)をリセット
        query = t_search_var.get().lower()
        t_listbox.delete(0, tk.END)
        t_filtered.clear()
        for i, tmpl in enumerate(templates):
            name = tmpl.get("name", "")
            text = tmpl.get("text", "")
            if query and query not in name.lower() and query not in text.lower():
                continue
            t_filtered.append(i)
            display = name if name else text.replace("\n", " ")[:72]
            t_listbox.insert(tk.END, f"  {display}")
        if t_filtered:
            t_listbox.selection_set(0)
            t_listbox.activate(0)
        t_status_var.set(f"  {len(t_filtered)}/{len(templates)} 件  クリック/Enter:コピー  Esc:閉じる")

    t_dnd_active = [False]  # DnD中はコピーをスキップするフラグ

    def copy_template(*_):
        if t_dnd_active[0]:
            return
        sel = t_listbox.curselection()
        if not sel:
            return
        text = _interpolate(templates[t_filtered[sel[0]]]["text"], current_history)
        root.destroy()
        pyperclip.copy(text)
        icon.update_menu()

    def on_template_key(event):
        if event.keysym == "Return":
            copy_template()
        elif event.keysym == "Escape":
            close()
        elif event.keysym in ("Left", "Right"):
            _switch_tab_relative(+1 if event.keysym == "Right" else -1)
        elif event.keysym == "BackSpace" or (len(event.char) == 1 and event.char.isprintable()):
            t_search_entry.focus_set()
            if event.keysym == "BackSpace":
                t_search_var.set(t_search_var.get()[:-1])
            else:
                t_search_entry.insert(tk.END, event.char)

    def _move_t_selection(delta: int) -> None:
        """t_listbox の選択を delta 分移動（フォーカスは検索バーのまま）。"""
        size = t_listbox.size()
        if size == 0:
            return
        sel = t_listbox.curselection()
        cur = sel[0] if sel else (-1 if delta > 0 else size)
        new = max(0, min(size - 1, cur + delta))
        t_listbox.selection_clear(0, tk.END)
        t_listbox.selection_set(new)
        t_listbox.activate(new)
        t_listbox.see(new)

    def on_t_search_key(event):
        if event.keysym == "Return":
            copy_template()
        elif event.keysym == "Escape":
            close()
        elif event.keysym == "Down":
            _move_t_selection(+1)
        elif event.keysym == "Up":
            _move_t_selection(-1)
        elif event.keysym == "Next":   # PageDown
            _move_t_selection(+PAGE_JUMP)
        elif event.keysym == "Prior":  # PageUp
            _move_t_selection(-PAGE_JUMP)
        elif event.keysym in ("Left", "Right"):
            _switch_tab_relative(+1 if event.keysym == "Right" else -1)
            return "break"  # Entry のカーソル移動を抑制

    # ---- 定型文: 右クリックメニュー ----
    def _template_new() -> None:
        def on_save(_: str, text: str) -> None:
            templates.append({"name": "", "text": text})
            save_templates(templates)
            refresh_templates()
        _dialog_open[0] += 1
        try:
            _edit_dialog(root, "定型文を新規作成", "", on_save, show_tag_hint=True)
        finally:
            _dialog_open[0] -= 1

    def _template_edit(t_idx: int) -> None:
        tmpl = templates[t_idx]
        def on_save(_: str, text: str) -> None:
            templates[t_idx] = {"name": "", "text": text}
            save_templates(templates)
            refresh_templates()
        _dialog_open[0] += 1
        try:
            _edit_dialog(root, "定型文を編集", tmpl.get("text", ""), on_save, show_tag_hint=True)
        finally:
            _dialog_open[0] -= 1

    def _template_delete(t_idx: int) -> None:
        templates.pop(t_idx)
        save_templates(templates)
        refresh_templates()

    def _template_context_menu(event: tk.Event) -> None:
        _hide_all_tooltips()
        idx = t_listbox.nearest(event.y)
        # nearest() はリストが空でも 0 を返すことがあるので件数チェック
        has_item = t_listbox.size() > 0 and 0 <= idx < len(t_filtered)
        if has_item:
            t_listbox.selection_clear(0, tk.END)
            t_listbox.selection_set(idx)
            t_listbox.activate(idx)
        t_idx = t_filtered[idx] if has_item else -1

        menu = tk.Menu(root, tearoff=0,
                       bg="#2b2b2b", fg="#dcdcdc",
                       activebackground="#0078d4", activeforeground="#ffffff",
                       font=("Yu Gothic UI", 9))
        menu.add_command(label="新規作成…", command=_template_new)
        if has_item:
            menu.add_separator()
            menu.add_command(label="編集…",
                             command=lambda: _template_edit(t_idx))
            menu.add_command(label="削除",
                             command=lambda: _template_delete(t_idx))

        menu.tk_popup(event.x_root, event.y_root)

    t_listbox.bind("<Button-3>", _template_context_menu)
    template_page.bind("<Button-3>", _template_context_menu)

    t_search_var.trace_add("write", refresh_templates)
    t_listbox.bind("<ButtonRelease-1>", copy_template)
    t_listbox.bind("<Return>", copy_template)
    t_listbox.bind("<Key>", on_template_key)
    t_search_entry.bind("<Key>", on_t_search_key)
    _bind_dnd_reorder(t_listbox, templates, t_filtered, t_dnd_active,
                      lambda: (save_templates(templates), refresh_templates()))
    _bind_hover_highlight(t_listbox)
    _tooltip_hiders.append(_bind_template_tooltip(root, t_listbox, t_filtered, templates, current_history))

    # ================================================================
    # ピンページ
    # ================================================================
    pin_page = tk.Frame(content_area, bg="#2b2b2b")

    p_search_var = tk.StringVar()
    _tk_vars.append(p_search_var)
    p_search_entry = tk.Entry(
        pin_page, textvariable=p_search_var,
        font=("Yu Gothic UI", 10),
        bg="#3c3c3c", fg="#dcdcdc", insertbackground="#dcdcdc",
        relief=tk.FLAT, bd=0,
    )
    p_search_entry.pack(fill=tk.X, padx=6, pady=(6, 3), ipady=4)
    tk.Frame(pin_page, bg="#555555", height=1).pack(fill=tk.X)

    p_listbox = tk.Listbox(
        pin_page,
        font=("Yu Gothic UI", 10),
        selectmode=tk.SINGLE, activestyle="none",
        bd=0, highlightthickness=0,
        bg="#2b2b2b", fg="#666666",
        selectbackground=ACCENT, selectforeground="#ffffff",
    )
    p_listbox.pack(fill=tk.BOTH, expand=True)

    p_status_var = tk.StringVar()
    _tk_vars.append(p_status_var)
    tk.Label(
        pin_page, textvariable=p_status_var, anchor="w",
        font=("Yu Gothic UI", 8), fg="#888888", bg="#2b2b2b",
    ).pack(fill=tk.X, padx=6, pady=(1, 4))

    p_filtered: list[int] = []

    def refresh_pins(*_):
        _hide_all_tooltips()
        query = p_search_var.get().lower()
        p_listbox.delete(0, tk.END)
        p_filtered.clear()
        for i, entry in enumerate(pins):
            text = entry["text"]
            if query and query not in text.lower():
                continue
            p_filtered.append(i)
            display = text.replace("\n", " ").replace("\r", "")
            if len(display) > 72:
                display = display[:72] + "…"
            p_listbox.insert(tk.END, f"  {display}")
        if p_filtered:
            p_listbox.selection_set(0)
            p_listbox.activate(0)
        p_status_var.set(f"  {len(p_filtered)}/{len(pins)} 件  クリック/Enter:コピー  Esc:閉じる")

    def _toggle_pin_from_history(hist_idx: int) -> None:
        """履歴のピン留めをトグルする（ピン済みなら解除、未ピンなら追加）。"""
        text = current_history[hist_idx]["text"]
        existing = next((i for i, p in enumerate(pins) if p["text"] == text), None)
        if existing is not None:
            pins.pop(existing)
        else:
            pins.append({
                "text": text,
                "time": current_history[hist_idx]["time"],
            })
        save_pins(pins)
        refresh_pins()
        refresh_history()  # 履歴側の📌アイコンを更新

    def copy_pin(*_):
        sel = p_listbox.curselection()
        if not sel:
            return
        text = pins[p_filtered[sel[0]]]["text"]
        root.destroy()
        pyperclip.copy(text)
        icon.update_menu()

    def on_pin_key(event):
        if event.keysym == "Return":
            copy_pin()
        elif event.keysym == "Escape":
            close()
        elif event.keysym in ("Left", "Right"):
            _switch_tab_relative(+1 if event.keysym == "Right" else -1)
        elif event.keysym == "BackSpace" or (len(event.char) == 1 and event.char.isprintable()):
            p_search_entry.focus_force()
            if event.keysym == "BackSpace":
                p_search_var.set(p_search_var.get()[:-1])
            else:
                p_search_entry.insert(tk.END, event.char)

    def _move_p_selection(delta: int) -> None:
        size = p_listbox.size()
        if size == 0:
            return
        sel = p_listbox.curselection()
        cur = sel[0] if sel else (-1 if delta > 0 else size)
        new = max(0, min(size - 1, cur + delta))
        p_listbox.selection_clear(0, tk.END)
        p_listbox.selection_set(new)
        p_listbox.activate(new)
        p_listbox.see(new)

    def on_p_search_key(event):
        if event.keysym == "Return":
            copy_pin()
        elif event.keysym == "Escape":
            close()
        elif event.keysym == "Down":
            _move_p_selection(+1)
        elif event.keysym == "Up":
            _move_p_selection(-1)
        elif event.keysym == "Next":
            _move_p_selection(+PAGE_JUMP)
        elif event.keysym == "Prior":
            _move_p_selection(-PAGE_JUMP)
        elif event.keysym in ("Left", "Right"):
            _switch_tab_relative(+1 if event.keysym == "Right" else -1)
            return "break"  # Entry のカーソル移動を抑制

    def _pin_edit(p_idx: int) -> None:
        def on_save(_, text: str) -> None:
            pins[p_idx]["text"] = text
            save_pins(pins)
            refresh_pins()
        _dialog_open[0] += 1
        try:
            _edit_dialog(root, "ピンを編集", pins[p_idx]["text"], on_save)
        finally:
            _dialog_open[0] -= 1

    def _pin_delete(p_idx: int) -> None:
        pins.pop(p_idx)
        save_pins(pins)
        refresh_pins()
        refresh_history()  # 履歴側の📌アイコンを更新

    def _pin_context_menu(event: tk.Event) -> None:
        _hide_all_tooltips()
        idx = p_listbox.nearest(event.y)
        has_item = p_listbox.size() > 0 and 0 <= idx < len(p_filtered)
        if has_item:
            p_listbox.selection_clear(0, tk.END)
            p_listbox.selection_set(idx)
            p_listbox.activate(idx)
        p_idx = p_filtered[idx] if has_item else -1

        menu = tk.Menu(root, tearoff=0,
                       bg="#2b2b2b", fg="#dcdcdc",
                       activebackground="#0078d4", activeforeground="#ffffff",
                       font=("Yu Gothic UI", 9))
        if has_item:
            menu.add_command(label="編集…",
                             command=lambda: _pin_edit(p_idx))
            menu.add_command(label="削除",
                             command=lambda: _pin_delete(p_idx))
        else:
            menu.add_command(label="（項目を右クリックしてください）", state=tk.DISABLED)

        menu.tk_popup(event.x_root, event.y_root)

    p_listbox.bind("<Button-3>", _pin_context_menu)
    p_search_var.trace_add("write", refresh_pins)
    p_listbox.bind("<ButtonRelease-1>", copy_pin)
    p_listbox.bind("<Return>", copy_pin)
    p_listbox.bind("<Key>", on_pin_key)
    p_search_entry.bind("<Key>", on_p_search_key)
    _bind_hover_highlight(p_listbox)
    _tooltip_hiders.append(_bind_listbox_tooltip(root, p_listbox, p_filtered, pins))

    # ================================================================
    # タブ切り替えロジック
    # ================================================================
    btn_refs: dict = {}

    def switch_tab(name: str) -> None:
        if name == "history":
            history_page.pack(fill=tk.BOTH, expand=True)
            template_page.pack_forget()
            pin_page.pack_forget()
            search_entry.focus_force()
        elif name == "template":
            template_page.pack(fill=tk.BOTH, expand=True)
            history_page.pack_forget()
            pin_page.pack_forget()
            t_search_entry.focus_force()
        else:  # pins
            pin_page.pack(fill=tk.BOTH, expand=True)
            history_page.pack_forget()
            template_page.pack_forget()
            p_search_entry.focus_force()

        for tab_name, btn in btn_refs.items():
            active = tab_name == name
            btn.config(
                bg=TAB_BG_ACTIVE if active else TAB_BG_INACTIVE,
                fg=TAB_FG_ACTIVE if active else TAB_FG_INACTIVE,
            )

        # アクティブタブの下線を描画
        root.update_idletasks()
        active_btn = btn_refs[name]
        bx = active_btn.winfo_x()
        bw = active_btn.winfo_width()
        tab_indicator.delete("all")
        tab_indicator.create_rectangle(bx, 0, bx + bw, 2, fill=ACCENT, outline="")

    for tab_name, label in [("history", "  履歴  "), ("template", "  定型文  "), ("pins", "  ピン  ")]:
        btn = tk.Button(
            tab_bar, text=label,
            font=("Yu Gothic UI", 9),
            bg=TAB_BG_INACTIVE, fg=TAB_FG_INACTIVE,
            relief=tk.FLAT, bd=0,
            activebackground="#3c3c3c", activeforeground="#dcdcdc",
            cursor="hand2",
            command=lambda n=tab_name: switch_tab(n),
        )
        btn.pack(side=tk.LEFT, ipady=5, ipadx=4)
        btn.bind("<Enter>", lambda _, n=tab_name: switch_tab(n))
        btn_refs[tab_name] = btn

    _TAB_ORDER = ["history", "template", "pins"]

    def _switch_tab_relative(delta: int) -> None:
        """現在アクティブなタブから delta 分移動したタブへ切り替える（ループ）。"""
        if history_page.winfo_ismapped():
            cur = "history"
        elif template_page.winfo_ismapped():
            cur = "template"
        else:
            cur = "pins"
        switch_tab(_TAB_ORDER[(_TAB_ORDER.index(cur) + delta) % len(_TAB_ORDER)])

    root.bind("<Escape>", close)

    # ---- ドラッグリサイズ ----
    _setup_resize(root)

    # ---- ウィンドウ位置・サイズ（カーソルのあるモニター上に表示） ----
    root.update_idletasks()
    ml, mt, mr, mb = _get_monitor_work_area(cx, cy)
    cfg = _load_config()
    w = cfg.get("popup_width",  POPUP_INIT_WIDTH)
    h = cfg.get("popup_height", POPUP_INIT_HEIGHT)
    x = max(ml, min(cx, mr - w - 4))
    y = max(mt, min(cy, mb - h - 4))
    root.geometry(f"{w}x{h}+{x}+{y}")
    _watch_click_outside(root, block_fn=lambda: _dialog_open[0] > 0)

    def _on_destroy(*_) -> None:
        try:
            _save_config({"popup_width": root.winfo_width(), "popup_height": root.winfo_height()})
        except tk.TclError:
            pass
    root.bind("<Destroy>", _on_destroy)

    # 初期表示
    refresh_history()
    refresh_templates()
    refresh_pins()
    switch_tab("history")
    root.deiconify()  # 事前初期化時に withdraw されている場合に表示
    search_entry.focus_force()  # ウィンドウの OS フォーカス取得 + 検索バーへのフォーカスを同時に行う
    root.mainloop()


def _watch_click_outside(root: tk.Tk, block_fn=None, interval_ms: int = 50) -> None:
    """マウスボタンが押されたとき、カーソルがウィンドウ外ならば root を閉じる。
    overrideredirect ウィンドウはフォアグラウンドを奪わないため
    GetAsyncKeyState でクリックを検出しカーソル座標と比較する。
    block_fn が指定されかつ True を返す間は、外クリックによる閉じる処理をスキップする。
    """
    VK_LBUTTON = 0x01
    VK_RBUTTON = 0x02

    class _POINT(ctypes.Structure):
        _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]

    _was_pressed = [False]

    def _check() -> None:
        try:
            ldown = bool(ctypes.windll.user32.GetAsyncKeyState(VK_LBUTTON) & 0x8000)
            rdown = bool(ctypes.windll.user32.GetAsyncKeyState(VK_RBUTTON) & 0x8000)
            pressed = ldown or rdown

            if pressed and not _was_pressed[0]:
                # ボタンが押された瞬間 — カーソルがウィンドウ内か判定
                pt = _POINT()
                ctypes.windll.user32.GetCursorPos(ctypes.byref(pt))
                wx = root.winfo_rootx()
                wy = root.winfo_rooty()
                ww = root.winfo_width()
                wh = root.winfo_height()
                if not (wx <= pt.x <= wx + ww and wy <= pt.y <= wy + wh):
                    # 編集ダイアログ等が開いている間は閉じない
                    if block_fn and block_fn():
                        _was_pressed[0] = pressed
                        root.after(interval_ms, _check)
                        return
                    root.destroy()
                    return

            _was_pressed[0] = pressed
            root.after(interval_ms, _check)
        except tk.TclError:
            pass  # root が既に破棄済み

    # 起動直後の自分自身のクリックを誤検知しないよう少し待ってから開始
    root.after(200, _check)


# ---------- ダブルCtrl 検出 ----------

def _popup_prewarm_loop() -> None:
    """Tk を事前初期化し、ホットキー／設定トリガーを待って処理するループ。
    ポップアップが閉じている間に次回用の Tk インタープリターを作成しておくことで
    ポップアップ表示までの遅延を削減する。
    設定ウィンドウも同スレッドで処理することで複数 Tk インスタンスの競合を防ぐ。
    """
    while True:
        root = tk.Tk()
        root.withdraw()

        def _check(_root=root):
            if _settings_trigger.is_set():
                _settings_trigger.clear()
                show_settings_window(_settings_icon_ref[0], _root)
            if _popup_trigger.is_set():
                _root.quit()
                return
            _root.after(50, _check)

        root.after(50, _check)
        root.mainloop()  # _check が quit() を呼ぶまでここで待機

        _popup_trigger.clear()
        icon = _popup_icon_ref[0]
        _popup_active.set()
        try:
            show_popup(icon, root)
        finally:
            _popup_active.clear()


def start_hotkey_listener(icon) -> None:
    """Ctrl キーの2連打を検出してポップアップを呼び出す。
    キーリピートを除外するため、キーアップ → キーダウンの遷移だけをカウントする。
    """
    last_ctrl_time = 0.0
    _ctrl_held = False          # Ctrl が現在押しっぱなしかどうか

    def on_ctrl_down(_):
        nonlocal last_ctrl_time, _ctrl_held
        if _ctrl_held:
            return  # キーリピートは無視
        _ctrl_held = True

        now = time.monotonic()
        if now - last_ctrl_time <= DOUBLE_CTRL_INTERVAL:
            last_ctrl_time = 0.0  # リセット（3連打対策）
            if not _popup_active.is_set():
                _popup_icon_ref[0] = icon
                _popup_trigger.set()
        else:
            last_ctrl_time = now

    def on_ctrl_up(_):
        nonlocal _ctrl_held
        _ctrl_held = False

    keyboard.on_press_key("ctrl", on_ctrl_down, suppress=False)
    keyboard.on_release_key("ctrl", on_ctrl_up, suppress=False)


# ---------- 多重起動防止 ----------

_MUTEX_NAME = "Global\\clipo_single_instance"
_mutex_handle = None  # プロセス終了まで保持する必要があるため module-level で保持

def _ensure_single_instance() -> None:
    """既に clipo が起動している場合はダイアログを表示して終了する。"""
    global _mutex_handle
    _mutex_handle = ctypes.windll.kernel32.CreateMutexW(None, True, _MUTEX_NAME)
    last_error = ctypes.windll.kernel32.GetLastError()
    ERROR_ALREADY_EXISTS = 183
    if last_error == ERROR_ALREADY_EXISTS:
        import tkinter as _tk
        import tkinter.messagebox as _mb
        _r = _tk.Tk()
        _r.withdraw()
        _mb.showwarning("clipo", "clipo はすでに起動しています。", parent=_r)
        _r.destroy()
        sys.exit(0)


# ---------- エントリポイント ----------

def main() -> None:
    _ensure_single_instance()
    global HISTORY_FILE, TEMPLATES_FILE, PINS_FILE, MAX_HISTORY, PAGE_JUMP
    cfg = _load_config()
    if "history_file" in cfg:
        HISTORY_FILE = Path(cfg["history_file"])
    if "templates_file" in cfg:
        TEMPLATES_FILE = Path(cfg["templates_file"])
    if "pins_file" in cfg:
        PINS_FILE = Path(cfg["pins_file"])
    else:
        # 未設定の場合は history_file と同じディレクトリに配置（書き込み権限を確保）
        PINS_FILE = HISTORY_FILE.parent / "pins.json"
    if "max_history" in cfg:
        MAX_HISTORY = int(cfg["max_history"])
    if "page_jump" in cfg:
        PAGE_JUMP = int(cfg["page_jump"])
    load_history()

    image = create_icon_image(64)

    icon = pystray.Icon(
        name="clipo",
        icon=image,
        title="clipo - クリップボード履歴",
    )
    icon._stop_event = threading.Event()
    icon.menu = pystray.Menu(lambda: build_menu(icon))

    # クリップボード監視スレッド
    watcher = threading.Thread(target=watch_clipboard, args=(icon,), daemon=True)
    watcher.start()

    # ポップアップ事前初期化ワーカー（Tk を先行作成して起動遅延を削減）
    threading.Thread(target=_popup_prewarm_loop, daemon=True).start()

    # ダブルCtrl ホットキーリスナー
    start_hotkey_listener(icon)

    icon.run()


if __name__ == "__main__":
    main()
