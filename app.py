# -*- coding: utf-8 -*-
"""
Streamlit app: bulk lookup of Kuwait MOI residence fines and PACI card status /
card renewal for a list of Civil IDs supplied in an Excel file.

Hardened for Streamlit Community Cloud (Debian container, no root, ~1 GB RAM):
  * Do NOT use packages.txt — Cloud's apt hits an expired bullseye-security
    Release file and aborts the whole deploy whenever packages.txt exists.
  * Instead, on Linux, `ensure_chromium()` vendors Playwright's Chromium
    shared libs by downloading Debian .debs into ~/.cache/pw-syslibs and
    setting LD_LIBRARY_PATH (no root). Then it runs
    `playwright install chromium` and verifies a real headless launch.
  * Playwright is pinned in requirements.txt for Python 3.14 wheels.
  * Chromium uses --no-sandbox / --disable-dev-shm-usage; images/media/fonts
    are blocked; the browser is recycled every N rows to keep memory flat.
  * TargetClosedError hardening: no --no-zygote / VizDisplayCompositor
    tweaks (they crash modern headless Chromium in containers), renderer
    heap is capped, launches retry with back-off, and the scraper detects a
    dead browser and relaunches it instead of aborting the whole job.
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

# Launch flags tuned for a 1 GB, no-root, no-/dev/shm container.
#
# Deliberately NOT used (they are the usual cause of "TargetClosedError:
# Target page, context or browser has been closed" on Streamlit Cloud):
#   --no-zygote                 renderer spawn fails without --single-process
#   --single-process            one crash takes the whole browser down
#   --disable-features=VizDisplayCompositor
#                               Viz is mandatory in current Chromium; disabling
#                               it kills the GPU/compositor process at start-up
CHROMIUM_ARGS = [
    "--no-sandbox",
    "--disable-setuid-sandbox",
    "--disable-dev-shm-usage",      # /dev/shm is tiny in containers
    "--disable-gpu",
    "--disable-software-rasterizer",
    "--disable-accelerated-2d-canvas",
    "--no-first-run",
    "--no-default-browser-check",
    "--no-service-autorun",
    "--disable-extensions",
    "--disable-component-update",
    "--disable-background-networking",
    "--disable-background-timer-throttling",
    "--disable-backgrounding-occluded-windows",
    "--disable-renderer-backgrounding",
    "--disable-ipc-flooding-protection",
    "--disable-hang-monitor",
    "--disable-breakpad",
    "--disable-crash-reporter",
    "--disable-sync",
    "--disable-domain-reliability",
    "--disable-client-side-phishing-detection",
    "--disable-popup-blocking",
    "--disable-prompt-on-repost",
    "--metrics-recording-only",
    "--password-store=basic",
    "--use-mock-keychain",
    "--force-color-profile=srgb",
    "--hide-scrollbars",
    "--mute-audio",
    "--disable-features=TranslateUI,AudioServiceOutOfProcess,MediaRouter,"
    "OptimizationHints,InterestFeedContentSuggestions",
    "--renderer-process-limit=2",
    "--js-flags=--max-old-space-size=256",  # cap renderer heap → no OOM-kill
    "--disable-blink-features=AutomationControlled",
    "--lang=ar-KW,ar",
]
LAUNCH_TIMEOUT_MS = 120_000   # first launch on Cloud can be slow (cold disk)
LAUNCH_ATTEMPTS = 3
# Substrings that identify a dead browser/page (Playwright wording varies).
_CLOSED_MARKERS = (
    "Target page, context or browser has been closed",
    "Target closed",
    "browser has been closed",
    "Browser closed",
    "Connection closed",
    "has been closed",
)


def launch_kwargs() -> dict:
    """Common `chromium.launch(**kwargs)` options for probe + scraper."""
    return dict(
        headless=True,
        args=CHROMIUM_ARGS,
        chromium_sandbox=False,
        timeout=LAUNCH_TIMEOUT_MS,
        # Streamlit owns the process signals; don't let the driver tear the
        # browser down on a SIGHUP/SIGINT meant for the web server.
        handle_sigint=False,
        handle_sigterm=False,
        handle_sighup=False,
    )


def is_closed_error(exc: BaseException) -> bool:
    """True if `exc` means the Playwright target/browser died."""
    if type(exc).__name__ == "TargetClosedError":
        return True
    msg = str(exc)
    return any(m in msg for m in _CLOSED_MARKERS)
BLOCKED_RESOURCE_TYPES = {"image", "media", "font"}

# Extra shared libs chrome-headless-shell needs beyond Playwright's minimal
# debian13 chromium list (Cloud reported libXrender.so.1 missing).
_CHROMIUM_APT_PACKAGES = (
    "libasound2t64",
    "libatk-bridge2.0-0t64",
    "libatk1.0-0t64",
    "libatspi2.0-0t64",
    "libcairo2",
    "libcups2t64",
    "libdbus-1-3",
    "libdrm2",
    "libgbm1",
    "libglib2.0-0t64",
    "libnspr4",
    "libnss3",
    "libpango-1.0-0",
    "libx11-6",
    "libx11-xcb1",
    "libxcb1",
    "libxcb-dri3-0",
    "libxcomposite1",
    "libxcursor1",
    "libxdamage1",
    "libxext6",
    "libxfixes3",
    "libxi6",
    "libxkbcommon0",
    "libxrandr2",
    "libxrender1",
    "libxtst6",
    "libexpat1",
    "libfontconfig1",
    "libfreetype6",
    "libharfbuzz0b",
    "libpng16-16t64",
    "zlib1g",
)
_SYSLIB_MARKER = ".ok.v3"  # bump to force re-vendor when the package set grows
# Fallbacks if the host is older than trixie (no t64 package names).
_CHROMIUM_APT_FALLBACKS = {
    "libasound2t64": "libasound2",
    "libatk-bridge2.0-0t64": "libatk-bridge2.0-0",
    "libatk1.0-0t64": "libatk1.0-0",
    "libatspi2.0-0t64": "libatspi2.0-0",
    "libcups2t64": "libcups2",
    "libglib2.0-0t64": "libglib2.0-0",
    "libpng16-16t64": "libpng16-16",
}
# SONAME -> Debian package when the naive libfoo.so.N -> libfooN guess fails.
_SO_TO_PKG = {
    "libXrender.so.1": "libxrender1",
    "libXi.so.6": "libxi6",
    "libXcursor.so.1": "libxcursor1",
    "libXtst.so.6": "libxtst6",
    "libX11.so.6": "libx11-6",
    "libX11-xcb.so.1": "libx11-xcb1",
    "libXext.so.6": "libxext6",
    "libXfixes.so.3": "libxfixes3",
    "libXdamage.so.1": "libxdamage1",
    "libXcomposite.so.1": "libxcomposite1",
    "libXrandr.so.2": "libxrandr2",
    "libxcb-dri3.so.0": "libxcb-dri3-0",
    "libpng16.so.16": "libpng16-16t64",
    "libharfbuzz.so.0": "libharfbuzz0b",
    "libfreetype.so.6": "libfreetype6",
    "libfontconfig.so.1": "libfontconfig1",
    "libexpat.so.1": "libexpat1",
    "libz.so.1": "zlib1g",
    "libasound.so.2": "libasound2t64",
    "libatk-1.0.so.0": "libatk1.0-0t64",
    "libatk-bridge-2.0.so.0": "libatk-bridge2.0-0t64",
    "libatspi.so.0": "libatspi2.0-0t64",
    "libcups.so.2": "libcups2t64",
    "libdbus-1.so.3": "libdbus-1-3",
    "libdrm.so.2": "libdrm2",
    "libgbm.so.1": "libgbm1",
    "libglib-2.0.so.0": "libglib2.0-0t64",
    "libgobject-2.0.so.0": "libglib2.0-0t64",
    "libgio-2.0.so.0": "libglib2.0-0t64",
    "libnspr4.so": "libnspr4",
    "libnss3.so": "libnss3",
    "libnssutil3.so": "libnss3",
    "libsmime3.so": "libnss3",
    "libssl3.so": "libnss3",
    "libpango-1.0.so.0": "libpango-1.0-0",
    "libpangocairo-1.0.so.0": "libpangocairo-1.0-0",
    "libcairo.so.2": "libcairo2",
    "libxkbcommon.so.0": "libxkbcommon0",
}


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


def _run(cmd: list[str], timeout: int = 600) -> tuple[int, str]:
    proc = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)
    out = ((proc.stdout or "") + "\n" + (proc.stderr or "")).strip()
    return proc.returncode, out[-3000:]


def _try_install_system_deps() -> str:
    """Best-effort `playwright install-deps` (needs root — usually no-op on Cloud)."""
    lines = []
    attempts = (
        [sys.executable, "-m", "playwright", "install-deps", "chromium"],
        ["sudo", "-n", sys.executable, "-m", "playwright", "install-deps", "chromium"],
    )
    for cmd in attempts:
        try:
            code, out = _run(cmd, timeout=300)
        except FileNotFoundError:
            continue
        except subprocess.TimeoutExpired:
            lines.append(f"$ {' '.join(cmd[2:])} -> TIMEOUT")
            continue
        lines.append(f"$ {' '.join(cmd)} -> exit {code}\n{out}")
        if code == 0:
            break
    return "\n".join(lines) or "install-deps skipped (no root / unavailable)"


def _resolve_deb_url(pkg: str, suite: str = "trixie") -> str | None:
    """Return a direct .deb URL for `pkg` from packages.debian.org, or None."""
    import urllib.request
    from html.parser import HTMLParser

    class _HrefParser(HTMLParser):
        def __init__(self):
            super().__init__()
            self.hrefs: list[str] = []

        def handle_starttag(self, tag, attrs):
            if tag != "a":
                return
            href = dict(attrs).get("href")
            if href and ".deb" in href:
                self.hrefs.append(href)

    page = f"https://packages.debian.org/{suite}/amd64/{pkg}/download"
    try:
        html = urllib.request.urlopen(page, timeout=60).read().decode("utf-8", "replace")
    except Exception:
        return None
    parser = _HrefParser()
    parser.feed(html)
    hrefs = []
    for h in parser.hrefs:
        if h.startswith("//"):
            h = "https:" + h
        if h.startswith("http") and h.endswith(".deb"):
            hrefs.append(h)
    preferred = [h for h in hrefs if "debian.org" in h]
    return (preferred or hrefs or [None])[0]


def _extract_deb(deb_path: Path, root: Path) -> bool:
    code, _ = _run(["dpkg-deb", "-x", str(deb_path), str(root)], timeout=120)
    if code == 0:
        return True
    code2, _ = _run(
        ["bash", "-lc", f"cd {deb_path.parent} && ar x {deb_path.name} && "
         f"tar -xf data.tar.* -C {root}"],
        timeout=120,
    )
    return code2 == 0


def _syslib_root() -> Path:
    return Path.home() / ".cache" / "pw-syslibs"


def _pkg_for_soname(soname: str) -> str | None:
    if soname in _SO_TO_PKG:
        return _SO_TO_PKG[soname]
    m = re.match(r"(lib[A-Za-z0-9+._-]+)\.so(?:\.(\d+))?", soname)
    if not m:
        return None
    base = m.group(1).lower().replace("_", "-")
    maj = m.group(2)
    return f"{base}{maj}" if maj else base


def _missing_soname(err: str) -> str | None:
    m = re.search(
        r"error while loading shared libraries:\s*([^\s:]+):\s*cannot open shared object file",
        err,
    )
    return m.group(1) if m else None


def _download_and_extract_pkg(pkg: str, root: Path, log: list[str]) -> bool:
    import urllib.request

    deb_url = _resolve_deb_url(pkg, "trixie") or _resolve_deb_url(pkg, "bookworm")
    used = pkg
    if deb_url is None and pkg in _CHROMIUM_APT_FALLBACKS:
        alt = _CHROMIUM_APT_FALLBACKS[pkg]
        deb_url = _resolve_deb_url(alt, "bookworm") or _resolve_deb_url(alt, "trixie")
        used = alt
    if deb_url is None:
        log.append(f"{pkg}: no .deb URL")
        return False
    deb_dir = root / "debs"
    deb_dir.mkdir(parents=True, exist_ok=True)
    deb_path = deb_dir / f"{used}.deb"
    try:
        urllib.request.urlretrieve(deb_url, deb_path)
    except Exception as exc:  # noqa: BLE001
        log.append(f"{used}: download failed: {exc}")
        return False
    if not _extract_deb(deb_path, root):
        log.append(f"{used}: extract failed")
        return False
    log.append(f"{used}: ok")
    return True


def _vendor_syslibs_from_debs(extra_pkgs: tuple[str, ...] = ()) -> str:
    """Download Chromium .debs from Debian and extract into a user-writable tree.

    Streamlit Cloud cannot use packages.txt right now (apt update fails on an
    expired bullseye-security Release file). Vendoring .debs needs no root.
    """
    if not IS_LINUX:
        return "syslib vendor skipped (not Linux)"

    root = _syslib_root()
    marker = root / _SYSLIB_MARKER
    lib_ok = (root / "usr" / "lib" / "x86_64-linux-gnu").exists() or (
        root / "lib" / "x86_64-linux-gnu"
    ).exists()
    pkgs = tuple(dict.fromkeys((*_CHROMIUM_APT_PACKAGES, *extra_pkgs)))

    # Reuse cache only when marker matches current package-set generation and
    # no extra packages were requested.
    if marker.exists() and lib_ok and not extra_pkgs:
        _prepend_ld_path(root)
        return f"syslibs ready: {root}"

    root.mkdir(parents=True, exist_ok=True)
    log: list[str] = []

    try:
        for pkg in pkgs:
            # Skip if this exact .deb was already extracted in a prior pass.
            if (root / "debs" / f"{pkg}.deb").exists() and lib_ok and extra_pkgs:
                # Still extract again for safety when filling gaps.
                pass
            _download_and_extract_pkg(pkg, root, log)

        if not (root / "usr" / "lib" / "x86_64-linux-gnu").exists() and not (
            root / "lib" / "x86_64-linux-gnu"
        ).exists():
            return "syslib vendor failed — no lib dir\n" + "\n".join(log)

        _prepend_ld_path(root)
        # Drop stale markers from older package sets.
        for old in root.glob(".ok*"):
            old.unlink(missing_ok=True)
        marker.write_text("ok\n", encoding="utf-8")
        return f"syslibs vendored -> {root}\n" + "\n".join(log)
    except Exception as exc:  # noqa: BLE001
        return f"syslib vendor error: {exc}\n" + "\n".join(log)


def _fix_missing_so(err: str) -> str:
    """Parse a missing .so from a launch error and vendor its Debian package."""
    soname = _missing_soname(err)
    if not soname:
        return "no missing .so parsed from launch error"
    pkg = _pkg_for_soname(soname)
    if not pkg:
        return f"unmapped soname: {soname}"
    # Invalidate ready marker so the new package is merged in.
    root = _syslib_root()
    for old in root.glob(".ok*"):
        old.unlink(missing_ok=True)
    return _vendor_syslibs_from_debs(extra_pkgs=(pkg,))



def _prepend_ld_path(root: Path) -> None:
    candidates = [
        root / "usr" / "lib" / "x86_64-linux-gnu",
        root / "lib" / "x86_64-linux-gnu",
        root / "usr" / "lib",
        root / "lib",
    ]
    existing = [str(p) for p in candidates if p.is_dir()]
    if not existing:
        return
    cur = os.environ.get("LD_LIBRARY_PATH", "")
    parts = existing + ([cur] if cur else [])
    os.environ["LD_LIBRARY_PATH"] = ":".join(parts)


def _launch_probe() -> str:
    """Start Chromium briefly; return its executable path on success.

    Checking `executable_path` alone is not enough – Playwright may point at a
    full browser while headless mode needs chromium-headless-shell. A real
    launch is the only reliable readiness check.
    """
    from playwright.sync_api import sync_playwright

    with sync_playwright() as p:
        browser = p.chromium.launch(**launch_kwargs())
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


def _last_launch_error() -> str:
    try:
        _in_thread(_launch_probe)
        return ""
    except Exception as exc:  # noqa: BLE001
        return f"{type(exc).__name__}: {exc}"


def _install_chromium() -> str:
    """Download Chromium browser binaries into PLAYWRIGHT_BROWSERS_PATH."""
    cmd = [sys.executable, "-m", "playwright", "install", "chromium"]
    code, out = _run(cmd, timeout=600)
    summary = f"$ {' '.join(cmd[2:])}\n{out}".strip()
    if code != 0:
        raise RuntimeError(
            "playwright install chromium failed "
            f"(exit {code}):\n{summary}"
        )
    return summary


@st.cache_resource(show_spinner=False)
def ensure_chromium() -> str:
    """Vendor syslibs (Linux) + install Chromium; verify headless launch."""
    notes: list[str] = []

    # On Streamlit Cloud / Linux, pull shared libs BEFORE the first launch try.
    # packages.txt is deliberately not used (Cloud apt currently hard-fails).
    if IS_LINUX:
        notes.append(_vendor_syslibs_from_debs())
    else:
        root = _syslib_root()
        if any(root.glob(".ok*")):
            _prepend_ld_path(root)

    notes.append(_try_install_system_deps())

    if not _chromium_works():
        notes.append(_install_chromium())

    # Iteratively vendor whatever .so the launcher still complains about
    # (e.g. libXrender.so.1 → libxrender1).
    for attempt in range(8):
        exe = _chromium_works()
        if exe:
            return f"Chromium ready: {exe}\n" + "\n".join(notes)
        err = _last_launch_error()
        if not IS_LINUX or not _missing_soname(err):
            notes.append(f"launch still failing (attempt {attempt + 1}): {err}")
            break
        notes.append(_fix_missing_so(err))

    launch_err = _last_launch_error()
    raise RuntimeError(
        "Chromium downloaded but headless launch failed (usually missing "
        "shared libraries). Syslib vendoring from Debian .debs should supply "
        "them without packages.txt.\n"
        f"Launch error: {launch_err}\n" + "\n".join(notes)
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
    UA = (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"
    )

    def __init__(self, headless: bool = True, timeout_s: int = 40, log=None):
        self.headless = headless
        self.timeout_ms = timeout_s * 1000
        self._log = log or (lambda _m: None)
        self.relaunches = 0
        self._pw = None
        self.browser = None
        self.context = None
        self.moi_page = None
        self.paci_page = None
        self._paci_loaded = False
        self._launch()

    # ------------------------------------------------------------ launch --- #
    def _launch(self):
        """Start Playwright + Chromium with retries (Cloud launches can flake)."""
        from playwright.sync_api import sync_playwright

        last: Exception | None = None
        for attempt in range(1, LAUNCH_ATTEMPTS + 1):
            try:
                self._pw = sync_playwright().start()
                kw = launch_kwargs()
                kw["headless"] = self.headless
                self.browser = self._pw.chromium.launch(**kw)
                self.context = self.browser.new_context(
                    locale="ar-KW",
                    viewport={"width": 1024, "height": 768},
                    java_script_enabled=True,
                    user_agent=self.UA,
                )
                self.context.set_default_timeout(self.timeout_ms)
                self.context.set_default_navigation_timeout(self.timeout_ms)
                self.context.route("**/*", self._route)
                self.moi_page = self.context.new_page()
                self.paci_page = self.context.new_page()
                self._paci_loaded = False
                return
            except Exception as exc:  # noqa: BLE001
                last = exc
                self._teardown()
                self._log(f"فشل تشغيل المتصفح (محاولة {attempt}/{LAUNCH_ATTEMPTS}): "
                          f"{type(exc).__name__}: {exc}")
                time.sleep(2 * attempt)
        raise RuntimeError(f"Chromium launch failed after {LAUNCH_ATTEMPTS} attempts: {last}")

    def _teardown(self):
        for obj, fn in (
            (self.context, "close"),
            (self.browser, "close"),
            (self._pw, "stop"),
        ):
            if obj is None:
                continue
            try:
                getattr(obj, fn)()
            except Exception:
                pass
        self._pw = self.browser = self.context = None
        self.moi_page = self.paci_page = None
        self._paci_loaded = False

    def alive(self) -> bool:
        try:
            return (
                self.browser is not None
                and self.browser.is_connected()
                and self.moi_page is not None and not self.moi_page.is_closed()
                and self.paci_page is not None and not self.paci_page.is_closed()
            )
        except Exception:
            return False

    def relaunch(self, reason: str = ""):
        """Kill whatever is left and bring up a fresh browser."""
        self.relaunches += 1
        self._log(f"إعادة تشغيل المتصفح ({self.relaunches}) {reason}".rstrip())
        self._teardown()
        self._launch()

    def ensure_alive(self):
        if not self.alive():
            self.relaunch("— المتصفح أُغلق أو تعطل")

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
                self.ensure_alive()
                out = fn(civil_id)
                if out != ERR_TIMEOUT or attempt == retries:
                    return out
                last = out
            except Exception as exc:  # noqa: BLE001
                last = f"{ERR_GENERIC}: {type(exc).__name__}"
                self._paci_loaded = False
                if is_closed_error(exc):
                    # Browser/renderer died (OOM, crash). Relaunch and give the
                    # same row one more chance instead of failing the job.
                    try:
                        self.relaunch(f"— {type(exc).__name__}")
                    except Exception as launch_exc:  # noqa: BLE001
                        return f"{ERR_GENERIC}: {type(launch_exc).__name__}"
                    if attempt == retries:
                        try:
                            return fn(civil_id)
                        except Exception as exc2:  # noqa: BLE001
                            return f"{ERR_GENERIC}: {type(exc2).__name__}"
            time.sleep(1.5)
        return last

    def close(self):
        self._teardown()


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
            scraper = GovScraper(headless=o["headless"], timeout_s=o["timeout_s"], log=self.log)
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
                    scraper.relaunch("— لتحرير الذاكرة")
                    rows_since_restart = 0
                elif not scraper.alive():
                    scraper.relaunch("— المتصفح أُغلق أو تعطل")

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
    st.error("تعذر تثبيت/تشغيل Chromium. راجع سجل الخطأ أدناه و requirements.txt.")
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
    restart_every = st.slider("إعادة تشغيل المتصفح كل N صف (لتوفير الذاكرة)", 10, 200,
                              20 if ON_CLOUD else 40, 10)
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
