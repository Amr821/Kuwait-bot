# -*- coding: utf-8 -*-
import os
os.system("playwright install chromium")
"""
Streamlit app: bulk lookup of Kuwait MOI residence fines and PACI card status /
card renewal for a list of Civil IDs supplied in an Excel file.

Selectors below were taken from the live pages (Sept 2026):

MOI  https://www.moi.gov.kw/main/eservices/residence/fines-enquiry?civilId=<ID>
     - the page auto-fills #civilId and clicks #btnEnquireFines when ?civilId= is present
     - result HTML is injected into #responseInfo (class "d-none" removed when ready)
     - client-side validation errors appear in label#civilId-error

PACI https://services.paci.gov.kw/
     - service tabs: li.srv1 = "تجديد البطاقة" (InquiryType=1)
                     li.srv2 = "حالة البطاقة"  (InquiryType=2)
     - input #txtCivilId, button #btnSearch, AJAX POST /card/search
     - result message rendered into .div-message .alert
"""

import io
import re
import time
from datetime import datetime

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
    # Excel often stores big numbers as floats -> "293010112345.0" or "2.9301e+11"
    try:
        if re.fullmatch(r"[0-9.eE+]+", s) and ("." in s or "e" in s.lower()):
            s = str(int(float(s)))
    except (ValueError, OverflowError):
        pass
    # Convert Arabic-Indic digits to ASCII digits
    s = s.translate(str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789"))
    s = re.sub(r"\D", "", s)
    if 0 < len(s) < 12:
        s = s.zfill(12)
    return s


def clean_text(text: str) -> str:
    if not text:
        return ""
    lines = [re.sub(r"\s+", " ", ln).strip() for ln in text.splitlines()]
    lines = [ln for ln in lines if ln]
    return " | ".join(lines)


# --------------------------------------------------------------------------- #
# Scraper
# --------------------------------------------------------------------------- #
class GovScraper:
    """Owns one Chromium instance with two tabs (MOI + PACI) reused across IDs."""

    def __init__(self, headless: bool = True, timeout_s: int = 40):
        from playwright.sync_api import sync_playwright

        self.timeout_ms = timeout_s * 1000
        self._pw = sync_playwright().start()
        self.browser = self._pw.chromium.launch(
            headless=headless,
            args=["--disable-blink-features=AutomationControlled", "--lang=ar-KW,ar"],
        )
        self.context = self.browser.new_context(
            locale="ar-KW",
            viewport={"width": 1366, "height": 900},
            user_agent=(
                "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
                "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"
            ),
        )
        self.context.set_default_timeout(self.timeout_ms)
        self.moi_page = self.context.new_page()
        self.paci_page = self.context.new_page()
        self._paci_loaded = False

    # ---------------------------------------------------------------- MOI --- #
    def moi_fines(self, civil_id: str) -> str:
        page = self.moi_page
        page.goto(MOI_URL + civil_id, wait_until="domcontentloaded")

        # Wait until either the result box is populated or a validation error shows
        js_ready = """
        () => {
            const r = document.querySelector('#responseInfo');
            const e = document.querySelector('#civilId-error');
            const errVisible = e && e.offsetParent !== null && e.innerText.trim().length > 0;
            const resReady = r && !r.classList.contains('d-none') && r.innerText.trim().length > 0;
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
            page.goto(PACI_URL, wait_until="domcontentloaded")
            page.wait_for_selector("#txtCivilId", state="visible")
            page.wait_for_selector("#btnSearch", state="visible")
            # give reCAPTCHA v3 a moment to populate the hidden token
            try:
                page.wait_for_function(
                    "() => (document.querySelector('#Token')||{}).value?.length > 20",
                    timeout=15000,
                )
            except Exception:
                pass
            self._paci_loaded = True

    def _paci_query(self, civil_id: str, service_li: str, inquiry_type: str) -> str:
        """service_li: 'srv1' (renewal) or 'srv2' (status)."""
        page = self.paci_page
        self._ensure_paci()

        # select the service tab (this also clears the input) and double-check the hidden field
        page.click(f"li.{service_li} > a")
        page.wait_for_function(
            f"() => document.querySelector('#InquiryType') && "
            f"document.querySelector('#InquiryType').value === '{inquiry_type}'",
            timeout=5000,
        )

        # clear any previous message so we never read a stale result
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
            const vis = d && getComputedStyle(d).display !== 'none';
            return vis && a && a.innerText.trim().length > 0;
        }
        """
        try:
            page.wait_for_function(js_ready, timeout=self.timeout_ms)
        except Exception:
            self._paci_loaded = False  # reload next time
            return ERR_TIMEOUT

        text = page.evaluate(
            "() => document.querySelector('#inquiryForm .alert').innerText"
        )
        text = clean_text(text)
        # the site's own error path also calls getRecaptcha() – reload to be safe
        if "حدث خطأ" in text or "error" in text.lower():
            self._paci_loaded = False
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
                # a hard error on PACI usually means the page state is broken
                if fn in (self.paci_card_status, self.paci_card_renewal):
                    self._paci_loaded = False
            time.sleep(1.5)
        return last

    def close(self):
        for obj in (self.context, self.browser):
            try:
                obj.close()
            except Exception:
                pass
        try:
            self._pw.stop()
        except Exception:
            pass


# --------------------------------------------------------------------------- #
# Streamlit UI
# --------------------------------------------------------------------------- #
st.set_page_config(page_title="استعلام الداخلية و الهيئة المدنية", page_icon="🇰🇼", layout="wide")

st.markdown(
    """
    <style>
    .stApp { direction: rtl; }
    .stDataFrame, .stMarkdown, .stButton, .stDownloadButton { direction: rtl; }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("🇰🇼 استعلام آلي: غرامات الإقامة (الداخلية) + حالة و تجديد البطاقة (PACI)")
st.caption(
    "ارفع ملف Excel يحتوي الأعمدة: "
    + " ، ".join(f"`{c}`" for c in REQUIRED_COLS)
    + " — سيتم تعبئة الأعمدة الثلاثة الأخيرة تلقائياً."
)

with st.sidebar:
    st.header("⚙️ الإعدادات")
    headless = st.toggle("تشغيل المتصفح في الخلفية (Headless)", value=True)
    timeout_s = st.slider("مهلة انتظار النتيجة (ثانية)", 10, 120, 40)
    delay_s = st.slider("فاصل زمني بين كل رقم مدني (ثانية)", 0.0, 10.0, 1.5, 0.5)
    retries = st.slider("عدد إعادة المحاولة عند انتهاء المهلة", 0, 3, 1)
    skip_filled = st.checkbox("تخطي الصفوف المعبأة مسبقاً", value=True)
    st.markdown("---")
    st.markdown("**الخدمات:**")
    do_moi = st.checkbox("الداخلية - غرامات الإقامة", value=True)
    do_status = st.checkbox("PACI - حالة البطاقة", value=True)
    do_renew = st.checkbox("PACI - تجديد البطاقة", value=True)

uploaded = st.file_uploader("📂 اختر ملف Excel (test.xlsx)", type=["xlsx", "xlsm", "xls"])

if "result_df" not in st.session_state:
    st.session_state.result_df = None
if "stop" not in st.session_state:
    st.session_state.stop = False


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


if uploaded is not None:
    try:
        df = pd.read_excel(uploaded, dtype={COL_CID: str})
    except Exception as exc:  # noqa: BLE001
        st.error(f"تعذر قراءة الملف: {exc}")
        st.stop()

    df.columns = [str(c).strip() for c in df.columns]
    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing:
        st.error("الأعمدة التالية غير موجودة في الملف: " + " ، ".join(missing))
        st.write("الأعمدة الموجودة:", list(df.columns))
        st.stop()

    for c in (COL_MOI, COL_RENEW, COL_STATUS):
        df[c] = df[c].astype("object")
    df[COL_CID] = df[COL_CID].apply(normalize_civil_id)

    st.subheader("📋 معاينة الملف")
    st.dataframe(df, use_container_width=True, height=260)
    st.info(f"عدد الصفوف: {len(df)}  —  أرقام مدنية صالحة (12 رقم): {(df[COL_CID].str.len() == 12).sum()}")

    col_run, col_stop = st.columns([1, 1])
    run = col_run.button("🚀 بدء الاستعلام", type="primary", use_container_width=True)
    if col_stop.button("⏹ إيقاف", use_container_width=True):
        st.session_state.stop = True

    if run:
        st.session_state.stop = False
        progress = st.progress(0.0, text="جاري تشغيل المتصفح…")
        status_box = st.empty()
        table_box = st.empty()
        log_box = st.expander("سجل التنفيذ", expanded=False)
        logs = []

        def log(msg: str):
            logs.append(f"[{datetime.now():%H:%M:%S}] {msg}")
            log_box.code("\n".join(logs[-200:]), language="text")

        scraper = None
        try:
            scraper = GovScraper(headless=headless, timeout_s=timeout_s)
            log("تم تشغيل Chromium.")
            total = len(df)
            for i, (idx, row) in enumerate(df.iterrows(), start=1):
                if st.session_state.stop:
                    log("تم الإيقاف بواسطة المستخدم.")
                    break

                cid = row[COL_CID]
                progress.progress(i / total, text=f"الصف {i} من {total} — الرقم المدني {cid or '—'}")

                if not cid or len(cid) != 12:
                    for c, enabled in ((COL_MOI, do_moi), (COL_STATUS, do_status), (COL_RENEW, do_renew)):
                        if enabled:
                            df.at[idx, c] = ERR_BAD_ID
                    log(f"{cid!r}: رقم مدني غير صالح - تم التخطي")
                    table_box.dataframe(df, use_container_width=True, height=320)
                    continue

                def needs(col: str) -> bool:
                    cur = df.at[idx, col]
                    return not (skip_filled and isinstance(cur, str) and cur.strip() and not cur.startswith(("⏱", "⚠")))

                if do_moi and needs(COL_MOI):
                    status_box.info(f"🔎 الداخلية — {cid}")
                    df.at[idx, COL_MOI] = scraper.safe(scraper.moi_fines, cid, retries)
                    log(f"{cid} | الداخلية: {df.at[idx, COL_MOI]}")

                if do_status and needs(COL_STATUS):
                    status_box.info(f"🔎 PACI حالة البطاقة — {cid}")
                    df.at[idx, COL_STATUS] = scraper.safe(scraper.paci_card_status, cid, retries)
                    log(f"{cid} | حالة البطاقة: {df.at[idx, COL_STATUS]}")

                if do_renew and needs(COL_RENEW):
                    status_box.info(f"🔎 PACI تجديد البطاقة — {cid}")
                    df.at[idx, COL_RENEW] = scraper.safe(scraper.paci_card_renewal, cid, retries)
                    log(f"{cid} | تجديد البطاقة: {df.at[idx, COL_RENEW]}")

                table_box.dataframe(df, use_container_width=True, height=320)
                st.session_state.result_df = df.copy()
                if delay_s:
                    time.sleep(delay_s)

            progress.progress(1.0, text="اكتمل ✅")
            status_box.success("انتهى الاستعلام لجميع الصفوف.")
        except Exception as exc:  # noqa: BLE001
            status_box.error(f"توقف التنفيذ بسبب خطأ: {exc}")
            log(f"خطأ عام: {exc!r}")
        finally:
            if scraper:
                scraper.close()
                log("تم إغلاق المتصفح.")
            st.session_state.result_df = df.copy()

# ---------------------------------------------------------------- download -- #
if st.session_state.result_df is not None:
    st.subheader("⬇️ تحميل النتائج")
    st.dataframe(st.session_state.result_df, use_container_width=True, height=320)
    fname = f"results_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    st.download_button(
        "📥 تحميل ملف Excel المحدث",
        data=to_excel_bytes(st.session_state.result_df),
        file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )
