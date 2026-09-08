# -*- coding: utf-8 -*-
"""
Streamlit app: bulk lookup of Kuwait MOI residence fines and PACI card status /
card renewal for a list of Civil IDs supplied in an Excel file.

Hardened for Streamlit Community Cloud (Debian container, no root, ~1 GB RAM):
  * On first start, `ensure_chromium()` runs `python -m playwright install
    chromium`, then verifies launch with container-safe flags. Result is
    cached for the life of the process via `@st.cache_resource`.
  * OS shared libraries come from packages.txt (apt). We never call
    `playwright install-deps` / `--with-deps` (needs root).
  * Playwright is pinned in requirements.txt to a Bullseye-compatible build.
  * Chromium uses --no-sandbox / --disable-dev-shm-usage; images/media/fonts
    are blocked; the browser is recycled every N rows to keep memory flat.
  * All Playwright work runs in a dedicated worker thread (avoids the
    "Sync API inside asyncio loop" error) and survives Streamlit reruns.

Selectors (verified on the live pages, Sept 2026):

MOI  https://www.moi.gov.kw/main/eservices/residence/fines-enquiry?civilId=<ID>
     - the page auto-fills #civilId and clicks #btnEnquireFines when ?civilId= is present
     - result HTML is injected into #responseInfo (class "d-none" removed when ready)
     - client-side validation errors appear in label#civilId-error

PACI https://services.paci.gov.kw/
     - service tabs: li.srv1 = "تجديد البطاقة" (InquiryType=1)
                     li.srv2 = "حالة البطاقة"  (InquiryType=2)
     - input #txtCivilId, button #btnSearch, AJAX POST /card/search
     - result message rendered into #inquiryForm .div-message .alert
"""

from __future__ import annotations

import io
import os
import platform
import queue
import re
import subprocess
import sys
import threading
import time
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

# --------------------------------------------------------------------------- #
# Constants
# --------------------------------------------------------------------------- #
COL_CID = "الرقم المدني"
COL_MOI = "الداخليه"
COL_RENEW = "تجديد البطاقه"
COL_STATUS = "حاله البطاقه"
REQUIRED_COLS = [COL_CID, COL_MOI, COL_RENEW, COL_STATUS]

MOI_URL = "https://www.moi.gov.kw/main/eservices/residence/fines-enquiry?civilId="
PACI_URL = "https://services.paci.gov.kw/"

ERR_TIMEOUT = "⏱ انتهت المهلة - لم تظهر النتيجة"
ERR_GENERIC = "⚠ خطأ أثناء الاستعلام"
ERR_BAD_ID = "⚠ رقم مدني غير صالح"
ERR_PREFIXES = ("⏱", "⚠")

IS_LINUX = platform.system() == "Linux"
HAS_DISPLAY = bool(os.environ.get("DISPLAY")) or not IS_LINUX
ON_CLOUD = IS_LINUX and not os.environ.get("DISPLAY")

# Browsers live under a writable cache. Streamlit Cloud's HOME works; honour an
# explicit PLAYWRIGHT_BROWSERS_PATH if the platform already set one.
_BROWSERS_DIR = Path(
    os.environ.get("PLAYWRIGHT_BROWSERS_PATH")
    or (Path.home() / ".cache" / "ms-playwright")
)
if str(_BROWSERS_DIR) != "0":
    _BROWSERS_DIR.mkdir(parents=True, exist_ok=True)
    os.environ["PLAYWRIGHT_BROWSERS_PATH"] = str(_BROWSERS_DIR)

CHROMIUM_ARGS = [
    "--no-sandbox",
    "--disable-setuid-sandbox",
    "--disable-dev-shm-usage",      # /dev/shm is tiny in containers
    "--disable-gpu",
    "--no-zygote",
    "--no-first-run",
    "--no-default-browser-check",
    "--disable-extensions",
    "--disable-background-networking",
    "--disable-background-timer-throttling",
    "--disable-renderer-backgrounding",
    "--disable-features=TranslateUI,AudioServiceOutOfProcess",
    "--mute-audio",
    "--renderer-process-limit=2",
    "--disable-blink-features=AutomationControlled",
    "--lang=ar-KW,ar",
]
BLOCKED_RESOURCE_TYPES = {"image", "media", "font"}


# --------------------------------------------------------------------------- #
# Playwright bootstrap (runs once per container)
# --------------------------------------------------------------------------- #
def _in_thread(fn):
    """Run fn in a plain thread (no asyncio loop) and return its result."""
    box = {}

    def target():
        try:
            box["v"] = fn()
        except Exception as exc:  # noqa: BLE001
            box["e"] = exc

    t = threading.Thread(target=target, daemon=True)
    t.start()
    t.join()
    if "e" in box:
        raise box["e"]
    return box.get("v")


def _launch_probe() -> str:
    """Start Chromium briefly; return its executable path on success.

    Checking `executable_path` alone is not enough – Playwright may point at a
    full browser while headless mode needs chromium-headless-shell. A real
    launch is the only reliable readiness check.
    """
    from playwright.sync_api import sync_playwright

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True, args=CHROMIUM_ARGS)
        try:
            page = browser.new_page()
            page.set_content("<html><body>ok</body></html>")
            return p.chromium.executable_path or "chromium"
        finally:
            browser.close()


def _chromium_works() -> str | None:
    """Return executable path if a headless launch succeeds, else None."""
    try:
        return _in_thread(_launch_probe)
    except Exception:
        return None


def _install_chromium() -> str:
    """Download Chromium binaries (no apt — packages.txt owns system libs)."""
    cmd = [sys.executable, "-m", "playwright", "install", "chromium"]
    proc = subprocess.run(cmd, capture_output=True, text=True, timeout=600)
    out = (proc.stdout or "")[-2000:]
    err = (proc.stderr or "")[-2000:]
    summary = f"$ {' '.join(cmd[2:])}\n{out}\n{err}".strip()
    if proc.returncode != 0:
        raise RuntimeError(
            "playwright install chromium failed "
            f"(exit {proc.returncode}):\n{summary}"
        )
    return summary


@st.cache_resource(show_spinner=False)
def ensure_chromium() -> str:
    """Install + verify Chromium once; cached for the life of the process."""
    exe = _chromium_works()
    if exe:
        return f"Chromium ready: {exe}"

    install_log = _install_chromium()
    exe = _chromium_works()
    if exe:
        return f"Chromium installed: {exe}\n{install_log}"

    raise RuntimeError(
        "Chromium was downloaded but could not launch. "
        "Check packages.txt for missing shared libraries "
        f"(libnss3, libgbm1, libgtk-3-0, …).\n{install_log}"
    )


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #
def normalize_civil_id(value) -> str:
    """Turn whatever pandas read (int, float, str, NaN) into a 12-digit string."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    s = str(value).strip()
    if s.lower() in ("nan", "none", ""):
        return ""
    try:  # Excel often stores big numbers as floats -> "293010112345.0" / "2.93e+11"
        if re.fullmatch(r"[0-9.eE+]+", s) and ("." in s or "e" in s.lower()):
            s = str(int(float(s)))
    except (ValueError, OverflowError):
        pass
    s = s.translate(str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789"))
    s = re.sub(r"\D", "", s)
    return s  # a valid Kuwaiti Civil ID is exactly 12 digits (starts with 2 or 3)


def clean_text(text: str) -> str:
    if not text:
        return ""
    lines = [re.sub(r"\s+", " ", ln).strip() for ln in text.splitlines()]
    return " | ".join(ln for ln in lines if ln)


def is_error(value) -> bool:
    return isinstance(value, str) and value.startswith(ERR_PREFIXES)


# --------------------------------------------------------------------------- #
# Scraper (must be created and used from ONE thread)
# --------------------------------------------------------------------------- #
class GovScraper:
    def __init__(self, headless: bool = True, timeout_s: int = 40):
        from playwright.sync_api import sync_playwright

        self.timeout_ms = timeout_s * 1000
        self._pw = sync_playwright().start()
        self.browser = self._pw.chromium.launch(headless=headless, args=CHROMIUM_ARGS)
        self.context = self.browser.new_context(
            locale="ar-KW",
            viewport={"width": 1280, "height": 800},
            java_script_enabled=True,
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"
            ),
        )
        self.context.set_default_timeout(self.timeout_ms)
        self.context.set_default_navigation_timeout(self.timeout_ms)
        self.context.route("**/*", self._route)
        self.moi_page = self.context.new_page()
        self.paci_page = self.context.new_page()
        self._paci_loaded = False

    @staticmethod
    def _route(route, request):
        if request.resource_type in BLOCKED_RESOURCE_TYPES:
            return route.abort()
        return route.continue_()

    # ---------------------------------------------------------------- MOI --- #
    def moi_fines(self, civil_id: str) -> str:
        page = self.moi_page
        page.goto(MOI_URL + civil_id, wait_until="load")

        js_ready = """
        () => {
            const r = document.querySelector('#responseInfo');
            const e = document.querySelector('#civilId-error');
            const resReady = r && !r.classList.contains('d-none') && r.innerText.trim().length > 0;
            const errTxt = e ? e.innerText.trim() : '';
            const errVisible = e && e.offsetParent !== null && errTxt && errTxt !== 'حقل مطلوب';
            return resReady || errVisible;
        }
        """
        try:
            page.wait_for_function(js_ready, timeout=self.timeout_ms)
        except Exception:
            return ERR_TIMEOUT

        result = page.evaluate(
            """() => {
                const r = document.querySelector('#responseInfo');
                if (r && !r.classList.contains('d-none') && r.innerText.trim())
                    return r.innerText;
                const e = document.querySelector('#civilId-error');
                return e ? e.innerText : '';
            }"""
        )
        return clean_text(result) or ERR_GENERIC

    # --------------------------------------------------------------- PACI --- #
    def _ensure_paci(self, force: bool = False):
        page = self.paci_page
        if force or not self._paci_loaded or "services.paci.gov.kw" not in page.url:
            page.goto(PACI_URL, wait_until="load")
            page.wait_for_selector("#txtCivilId", state="visible")
            page.wait_for_selector("#btnSearch", state="visible")
            try:  # give reCAPTCHA v3 a moment to populate the hidden token
                page.wait_for_function(
                    "() => ((document.querySelector('#Token')||{}).value||'').length > 20",
                    timeout=15000,
                )
            except Exception:
                pass
            self._paci_loaded = True

    def _paci_query(self, civil_id: str, service_li: str, inquiry_type: str) -> str:
        page = self.paci_page
        self._ensure_paci()

        page.click(f"li.{service_li} > a")
        page.wait_for_function(
            "(t) => (document.querySelector('#InquiryType')||{}).value === t",
            arg=inquiry_type,
            timeout=5000,
        )
        page.evaluate(
            """() => {
                const a = document.querySelector('#inquiryForm .alert');
                if (a) a.innerHTML = '';
                const d = document.querySelector('#inquiryForm .div-message');
                if (d) d.style.display = 'none';
            }"""
        )
        page.fill("#txtCivilId", "")
        page.fill("#txtCivilId", civil_id)
        page.click("#btnSearch")

        js_ready = """
        () => {
            const d = document.querySelector('#inquiryForm .div-message');
            const a = document.querySelector('#inquiryForm .alert');
            return d && getComputedStyle(d).display !== 'none' && a && a.innerText.trim().length > 0;
        }
        """
        try:
            page.wait_for_function(js_ready, timeout=self.timeout_ms)
        except Exception:
            self._paci_loaded = False
            return ERR_TIMEOUT

        text = clean_text(page.evaluate("() => document.querySelector('#inquiryForm .alert').innerText"))
        if "حدث خطأ" in text or "error" in text.lower():
            self._paci_loaded = False  # site's error path – reload next time
        return text or ERR_GENERIC

    def paci_card_status(self, civil_id: str) -> str:
        return self._paci_query(civil_id, "srv2", "2")

    def paci_card_renewal(self, civil_id: str) -> str:
        return self._paci_query(civil_id, "srv1", "1")

    # ------------------------------------------------------------ generic --- #
    def safe(self, fn, civil_id: str, retries: int = 1) -> str:
        last = ERR_GENERIC
        for attempt in range(retries + 1):
            try:
                out = fn(civil_id)
                if out != ERR_TIMEOUT or attempt == retries:
                    return out
                last = out
            except Exception as exc:  # noqa: BLE001
                last = f"{ERR_GENERIC}: {type(exc).__name__}"
                self._paci_loaded = False
            time.sleep(1.5)
        return last

    def close(self):
        for fn in (self.context.close, self.browser.close, self._pw.stop):
            try:
                fn()
            except Exception:
                pass


# --------------------------------------------------------------------------- #
# Background job (thread) – survives Streamlit reruns
# --------------------------------------------------------------------------- #
class ScrapeJob:
    def __init__(self, df: pd.DataFrame, opts: dict):
        self.df = df
        self.opts = opts
        self.total = len(df)
        self.processed = 0
        self.current = ""
        self.logs: list[str] = []
        self.error: str | None = None
        self.lock = threading.Lock()
        self.stop_event = threading.Event()
        self.done = threading.Event()
        self.thread = threading.Thread(target=self._run, name="scraper", daemon=True)

    # -- API used by the UI thread --
    def start(self):
        self.thread.start()

    def snapshot(self) -> pd.DataFrame:
        with self.lock:
            return self.df.copy()

    def log(self, msg: str):
        with self.lock:
            self.logs.append(f"[{datetime.now():%H:%M:%S}] {msg}")
            self.logs = self.logs[-300:]

    def _set(self, idx, col, value):
        with self.lock:
            self.df.at[idx, col] = value

    def _needs(self, idx, col) -> bool:
        with self.lock:
            cur = self.df.at[idx, col]
        if not self.opts["skip_filled"]:
            return True
        return not (isinstance(cur, str) and cur.strip() and not is_error(cur))

    # -- worker --
    def _run(self):
        o = self.opts
        scraper = None
        try:
            scraper = GovScraper(headless=o["headless"], timeout_s=o["timeout_s"])
            self.log("تم تشغيل Chromium.")
            rows_since_restart = 0

            for i, (idx, row) in enumerate(self.df.iterrows(), start=1):
                if self.stop_event.is_set():
                    self.log("تم الإيقاف بواسطة المستخدم.")
                    break

                cid = row[COL_CID]
                self.current = cid or "—"

                if not cid or len(cid) != 12:
                    for col, enabled in ((COL_MOI, o["do_moi"]), (COL_STATUS, o["do_status"]), (COL_RENEW, o["do_renew"])):
                        if enabled:
                            self._set(idx, col, ERR_BAD_ID)
                    self.log(f"{cid!r}: رقم مدني غير صالح - تم التخطي")
                    self.processed = i
                    continue

                if rows_since_restart >= o["restart_every"]:
                    scraper.close()
                    scraper = GovScraper(headless=o["headless"], timeout_s=o["timeout_s"])
                    rows_since_restart = 0
                    self.log("تمت إعادة تشغيل المتصفح لتحرير الذاكرة.")

                if o["do_moi"] and self._needs(idx, COL_MOI):
                    val = scraper.safe(scraper.moi_fines, cid, o["retries"])
                    self._set(idx, COL_MOI, val)
                    self.log(f"{cid} | الداخلية: {val}")

                if o["do_status"] and self._needs(idx, COL_STATUS):
                    val = scraper.safe(scraper.paci_card_status, cid, o["retries"])
                    self._set(idx, COL_STATUS, val)
                    self.log(f"{cid} | حالة البطاقة: {val}")

                if o["do_renew"] and self._needs(idx, COL_RENEW):
                    val = scraper.safe(scraper.paci_card_renewal, cid, o["retries"])
                    self._set(idx, COL_RENEW, val)
                    self.log(f"{cid} | تجديد البطاقة: {val}")

                self.processed = i
                rows_since_restart += 1
                if o["delay_s"]:
                    time.sleep(o["delay_s"])

            self.log("انتهى التنفيذ.")
        except Exception as exc:  # noqa: BLE001
            self.error = f"{type(exc).__name__}: {exc}"
            self.log(f"خطأ عام: {self.error}")
        finally:
            if scraper:
                scraper.close()
                self.log("تم إغلاق المتصفح.")
            self.done.set()


# --------------------------------------------------------------------------- #
# Excel export
# --------------------------------------------------------------------------- #
def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        df.to_excel(xw, index=False, sheet_name="Sheet1")
        ws = xw.sheets["Sheet1"]
        ws.sheet_view.rightToLeft = True
        for col_cells in ws.columns:
            width = max(len(str(c.value)) if c.value is not None else 0 for c in col_cells)
            ws.column_dimensions[col_cells[0].column_letter].width = min(max(14, width + 2), 80)
    return buf.getvalue()


# --------------------------------------------------------------------------- #
# Streamlit UI
# --------------------------------------------------------------------------- #
st.set_page_config(page_title="استعلام الداخلية و الهيئة المدنية", page_icon="🇰🇼", layout="wide")
st.markdown(
    "<style>.stApp{direction:rtl} .stDataFrame,.stMarkdown,.stButton,.stDownloadButton{direction:rtl}</style>",
    unsafe_allow_html=True,
)
st.title("🇰🇼 استعلام آلي: غرامات الإقامة (الداخلية) + حالة و تجديد البطاقة (PACI)")
st.caption("ارفع ملف Excel يحتوي الأعمدة: " + " ، ".join(f"`{c}`" for c in REQUIRED_COLS)
           + " — سيتم تعبئة الأعمدة الثلاثة الأخيرة تلقائياً.")

# ---- one-time browser bootstrap ------------------------------------------- #
try:
    with st.spinner("جاري تجهيز المتصفح (يحدث مرة واحدة عند أول تشغيل)…"):
        boot_msg = ensure_chromium()
except Exception as exc:  # noqa: BLE001
    st.error("تعذر تثبيت Chromium. راجع packages.txt / requirements.txt.")
    st.code(str(exc))
    st.stop()

with st.sidebar:
    st.header("⚙️ الإعدادات")
    if ON_CLOUD:
        st.caption("🖥️ بيئة خادم بدون شاشة — المتصفح يعمل Headless إجبارياً.")
        headless = True
    else:
        headless = st.toggle("تشغيل المتصفح في الخلفية (Headless)", value=True)
    timeout_s = st.slider("مهلة انتظار النتيجة (ثانية)", 10, 120, 45)
    delay_s = st.slider("فاصل زمني بين كل رقم مدني (ثانية)", 0.0, 10.0, 2.0, 0.5)
    retries = st.slider("عدد إعادة المحاولة عند انتهاء المهلة", 0, 3, 1)
    restart_every = st.slider("إعادة تشغيل المتصفح كل N صف (لتوفير الذاكرة)", 10, 200, 40, 10)
    skip_filled = st.checkbox("تخطي الخلايا المعبأة مسبقاً", value=True)
    st.markdown("---")
    st.markdown("**الخدمات:**")
    do_moi = st.checkbox("الداخلية - غرامات الإقامة", value=True)
    do_status = st.checkbox("PACI - حالة البطاقة", value=True)
    do_renew = st.checkbox("PACI - تجديد البطاقة", value=True)
    with st.expander("معلومات النظام"):
        st.code(f"python {platform.python_version()} / {platform.system()}\n{boot_msg.splitlines()[0]}")

ss = st.session_state
ss.setdefault("job", None)
ss.setdefault("result_df", None)
ss.setdefault("input_df", None)
ss.setdefault("input_name", None)

uploaded = st.file_uploader("📂 اختر ملف Excel (test.xlsx)", type=["xlsx", "xlsm", "xls"])

# ---- read & validate the upload (only when the file changes) -------------- #
if uploaded is not None and uploaded.name != ss.input_name:
    try:
        df_in = pd.read_excel(uploaded, dtype={COL_CID: str})
    except Exception as exc:  # noqa: BLE001
        st.error(f"تعذر قراءة الملف: {exc}")
        st.stop()
    df_in.columns = [str(c).strip() for c in df_in.columns]
    missing = [c for c in REQUIRED_COLS if c not in df_in.columns]
    if missing:
        st.error("الأعمدة التالية غير موجودة في الملف: " + " ، ".join(missing))
        st.write("الأعمدة الموجودة:", list(df_in.columns))
        st.stop()
    for c in (COL_MOI, COL_RENEW, COL_STATUS):
        df_in[c] = df_in[c].astype("object")
    df_in[COL_CID] = df_in[COL_CID].apply(normalize_civil_id)
    ss.input_df, ss.input_name, ss.result_df = df_in, uploaded.name, None

job: ScrapeJob | None = ss.job
running = job is not None and not job.done.is_set()

if ss.input_df is not None:
    df = ss.input_df
    st.subheader("📋 معاينة الملف")
    st.dataframe(df, use_container_width=True, height=240)
    st.info(f"عدد الصفوف: {len(df)}  —  أرقام مدنية صالحة (12 رقم): {(df[COL_CID].str.len() == 12).sum()}")

    c1, c2 = st.columns(2)
    start_clicked = c1.button("🚀 بدء الاستعلام", type="primary", use_container_width=True, disabled=running)
    stop_clicked = c2.button("⏹ إيقاف", use_container_width=True, disabled=not running)

    if stop_clicked and job is not None:
        job.stop_event.set()

    if start_clicked and not running:
        job = ScrapeJob(
            df.copy(),
            dict(headless=headless, timeout_s=timeout_s, delay_s=delay_s, retries=retries,
                 restart_every=restart_every, skip_filled=skip_filled,
                 do_moi=do_moi, do_status=do_status, do_renew=do_renew),
        )
        job.start()
        ss.job = job
        running = True

# ---- live progress (polls the worker; safe across reruns) ----------------- #
if job is not None:
    progress = st.progress(0.0)
    status_box = st.empty()
    table_box = st.empty()
    log_box = st.expander("سجل التنفيذ", expanded=False).empty()

    def render():
        pct = job.processed / job.total if job.total else 1.0
        progress.progress(min(pct, 1.0), text=f"الصف {job.processed} من {job.total} — {job.current}")
        table_box.dataframe(job.snapshot(), use_container_width=True, height=320)
        with job.lock:
            log_box.code("\n".join(job.logs[-200:]) or "…", language="text")

    while not job.done.is_set():
        status_box.info(f"🔎 جاري الاستعلام… {job.current}")
        render()
        time.sleep(1.0)

    render()
    ss.result_df = job.snapshot()
    if job.error:
        status_box.error(f"توقف التنفيذ بسبب خطأ: {job.error}")
    elif job.stop_event.is_set():
        status_box.warning("تم إيقاف التنفيذ. يمكنك تحميل النتائج الجزئية أدناه.")
    else:
        status_box.success("انتهى الاستعلام لجميع الصفوف ✅")

# ---- download ------------------------------------------------------------- #
if ss.result_df is not None:
    st.subheader("⬇️ تحميل النتائج")
    st.download_button(
        "📥 تحميل ملف Excel المحدث",
        data=to_excel_bytes(ss.result_df),
        file_name=f"results_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )
