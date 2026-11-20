"""菜品銷售紀錄系統 — Windows 桌面版（pywebview + WebView2，資料存 %APPDATA%\\DishSales）"""
import base64
import ctypes
import datetime as dt
import functools
import json
import os
import sys
import threading
import time
import webbrowser
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

import webview

APP_TITLE = "菜品銷售紀錄系統"

if getattr(sys, "frozen", False):
    WEB_DIR = Path(sys._MEIPASS) / "web"
else:
    WEB_DIR = Path(__file__).resolve().parent.parent / "web"

DATA_DIR = Path(os.environ.get("DISHSALES_DATA") or Path(os.environ.get("APPDATA", Path.home())) / "DishSales")
DATA_FILE = DATA_DIR / "data.json"
BACKUP_DIR = DATA_DIR / "backups"
KEEP_BACKUPS = 30
FILE_TYPES = {
    ".xlsx": "Excel 活頁簿 (*.xlsx)",
    ".json": "JSON 備份 (*.json)",
}

_window = None


class Api:
    """給網頁呼叫的介面（window.pywebview.api.*）"""

    def load(self):
        if not DATA_FILE.exists():
            return ""
        return DATA_FILE.read_text(encoding="utf-8")

    def save(self, text):
        try:
            json.loads(text)  # 不寫入損壞的內容
            DATA_DIR.mkdir(parents=True, exist_ok=True)
            tmp = DATA_FILE.with_suffix(".tmp")
            tmp.write_text(text, encoding="utf-8")
            os.replace(tmp, DATA_FILE)
            _daily_backup(text)
            return True
        except Exception as e:  # noqa: BLE001 - 錯誤訊息回傳給網頁顯示
            return f"{type(e).__name__}: {e}"

    def data_path(self):
        return str(DATA_FILE)

    def open_url(self, url):
        """用預設瀏覽器開啟 Google 試算表（只允許 Google 網址）"""
        if not str(url).startswith("https://docs.google.com/"):
            return False
        return webbrowser.open(url)

    def export_file(self, b64, filename):
        """另存新檔（Excel 報表或 JSON 備份）；回傳儲存路徑，取消時回傳空字串"""
        ext = Path(filename).suffix.lower()
        docs = Path.home() / "Documents"
        out = os.environ.get("DISHSALES_EXPORT_DIR")  # 自我測試時不跳對話框
        if out:
            res = str(Path(out) / filename)
        else:
            res = _window.create_file_dialog(
                webview.FileDialog.SAVE,
                directory=str(docs if docs.exists() else Path.home()),
                save_filename=filename,
                file_types=(FILE_TYPES.get(ext, "所有檔案 (*.*)"),),
            )
        if not res:
            return ""
        path = Path(res[0] if isinstance(res, (list, tuple)) else res)
        if ext and path.suffix.lower() != ext:
            path = path.with_suffix(ext)
        path.write_bytes(base64.b64decode(b64))
        return str(path)


def _daily_backup(text):
    """每天保留一份資料快照，最多 KEEP_BACKUPS 份"""
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    (BACKUP_DIR / f"data_{dt.date.today():%Y%m%d}.json").write_text(text, encoding="utf-8")
    for old in sorted(BACKUP_DIR.glob("data_*.json"))[:-KEEP_BACKUPS]:
        old.unlink(missing_ok=True)


class _QuietHandler(SimpleHTTPRequestHandler):
    extensions_map = {**SimpleHTTPRequestHandler.extensions_map, ".js": "text/javascript", ".webmanifest": "application/manifest+json"}

    def log_message(self, *args):  # --windowed 時沒有 stderr
        pass


def _serve_web():
    """只綁本機的靜態伺服器（pywebview 會把 file://...?desktop=1 的問號編碼掉，所以不用 file://）"""
    httpd = ThreadingHTTPServer(("127.0.0.1", 0), functools.partial(_QuietHandler, directory=str(WEB_DIR)))
    threading.Thread(target=httpd.serve_forever, daemon=True).start()
    return f"http://127.0.0.1:{httpd.server_address[1]}/index.html?desktop=1"


def _single_instance():
    if os.environ.get("DISHSALES_DATA"):
        return True  # 測試用的獨立資料夾不受限制
    ctypes.windll.kernel32.CreateMutexW(None, False, "DishSales_SingleInstance")
    if ctypes.windll.kernel32.GetLastError() == 183:  # ERROR_ALREADY_EXISTS
        ctypes.windll.user32.MessageBoxW(None, "程式已經在執行中。", APP_TITLE, 0x40)
        return False
    return True


def _selftest(out_path):
    """DISHSALES_SELFTEST=<結果檔>：啟動後自動操作畫面、寫出結果並關閉"""
    result = {"ok": False}
    js = _window.evaluate_js
    try:
        _window.events.loaded.wait(30)
        n = -1
        for _ in range(100):
            n = js("window.__dish && window.__dish.db ? window.__dish.db.dishes.length : -1")
            if n is not None and n >= 0:
                break
            time.sleep(0.2)
        result["dishes"] = n
        time.sleep(0.5)
        result["day_rows"] = js("document.querySelectorAll('.day-groups .row').length")
        # 輸入一筆份數（模擬按 ＋ 三次）
        js("for (let i = 0; i < 3; i++) document.querySelector('.row [data-act=inc]').click()")
        time.sleep(0.2)
        result["day_kpi"] = js("document.querySelector('#daykpi').innerText")
        js("window.__dish.flush()")
        time.sleep(0.8)
        result["saved_today"] = json.loads(DATA_FILE.read_text(encoding="utf-8"))["sales"].get(dt.date.today().isoformat())
        js("window.__dish.demoData(60); window.__dish.flush(); location.hash = '#/month'")
        time.sleep(1)
        result["month_kpi"] = js("[...document.querySelectorAll('.kpi .v')].map(e => e.innerText)")
        result["cal_days"] = js("document.querySelectorAll('.cal-g .cd.has').length")
        result["rank_rows"] = js("document.querySelectorAll('.rank tbody tr').length")
        js("document.querySelector('[data-act=xlsx]').click()")
        time.sleep(1.5)
        xl = list(Path(os.environ.get("DISHSALES_EXPORT_DIR", "")).glob("*.xlsx")) if os.environ.get("DISHSALES_EXPORT_DIR") else []
        result["xlsx"] = [str(p) for p in xl]
        js("location.hash = '#/settings'")
        time.sleep(0.6)
        result["dish_list"] = js("document.querySelectorAll('.dlist li').length")
        result["data_path_shown"] = js("document.querySelector('#dpath') && document.querySelector('#dpath').textContent")
        result["window"] = [_window.x, _window.y, _window.width, _window.height]
        result["ok"] = bool(n and n > 0 and result["saved_today"] and result["cal_days"] and result["rank_rows"] and result["xlsx"] and result["dish_list"])
    except Exception as e:  # noqa: BLE001
        result["error"] = repr(e)
    Path(out_path).write_text(json.dumps(result, ensure_ascii=False, indent=1), encoding="utf-8")
    _window.destroy()


def _fit_screen(margin=24):
    """依螢幕工作區（扣掉工作列）決定視窗大小，並置中"""
    try:
        wa = webview.screens[0].frame
        w, h = max(420, wa.Width - margin * 2), max(560, wa.Height - margin * 2)
        return dict(width=w, height=h, x=wa.X + (wa.Width - w) // 2, y=wa.Y + (wa.Height - h) // 2)
    except Exception:  # noqa: BLE001
        return dict(width=1200, height=760)


def main():
    global _window
    if not _single_instance():
        return
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    url = _serve_web()
    _window = webview.create_window(
        APP_TITLE, url, js_api=Api(), min_size=(420, 560), text_select=True, maximized=True, **_fit_screen(),
    )
    selftest = os.environ.get("DISHSALES_SELFTEST")
    kwargs = dict(private_mode=False, storage_path=str(DATA_DIR / "webview"), gui="edgechromium")
    if selftest:
        webview.start(_selftest, (selftest,), **kwargs)
    else:
        webview.start(**kwargs)


if __name__ == "__main__":
    main()
