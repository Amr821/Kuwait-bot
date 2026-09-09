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
  * Long-run (700+ rows) hardening:
      - a watchdog thread kills Chromium if any single lookup blocks past
        2×timeout+30 s, so a wedged renderer/driver can never hang the job;
      - wait_for_function uses interval polling (rAF polling stalls in
        background/occluded headless windows → every row "times out");
      - MOI serves a WAF/error page → detected in ≤15 s, not a full timeout;
      - circuit breaker: 5 consecutive all-error rows → relaunch + back-off,
        40 → abort with a clear message instead of "hanging" for hours;
      - the job lives in a process-wide registry (st.cache_resource), so a
        reload / dropped websocket re-attaches instead of starting a second
        browser (two Chromiums on 1 GB = OOM = TargetClosedError);
      - progress UI is a non-blocking st.fragment that ships only a window
        of rows, not all 735 every second.
  * Cloud-crash hardening (renderer dies on first real page):
      - fonts + fontconfig are vendored (no fonts → renderer CHECK on first
        text layout), the boot probe types/clicks/paints like a real row;
      - ONE page for both sites and one renderer (no site isolation) so a
        1 GB container is not OOM-killing the target;
      - launch-profile ladder: headless-shell → full Chromium new headless →
        --single-process; chosen at boot by the probe, advanced at runtime
        after two consecutive dead targets. Everything is logged with RAM.
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
    # No site isolation: the reCAPTCHA iframes on both sites would otherwise
    # get their own renderer processes (3–4 renderers on a 1 GB box → OOM).
    "--disable-site-isolation-trials",
    "--disable-features=TranslateUI,AudioServiceOutOfProcess,MediaRouter,"
    "OptimizationHints,InterestFeedContentSuggestions,IsolateOrigins,site-per-process",
    "--renderer-process-limit=1",
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
    "Page crashed",
    "crashed",
)


# Launch profiles, tried in order. The boot probe (which types, clicks and
# paints like a real row) picks the first that survives; the scraper moves
# to the next one if a target still dies twice in a row at runtime.
#   0  chromium-headless-shell (Playwright default, smallest)
#   1  full Chromium in new headless mode (different binary/code path)
#   2  headless-shell with --single-process (no renderer spawn at all – for
#      containers where forking the renderer fails or RAM is very tight)
LAUNCH_PROFILES = (
    {"name": "headless-shell"},
    {"name": "chromium-new-headless", "channel": "chromium"},
    {"name": "single-process", "extra_args": ["--single-process", "--no-zygote"]},
)
_PROFILE = {"idx": 0}   # chosen by ensure_chromium(); bumped by GovScraper


def launch_kwargs(profile_idx: int | None = None) -> dict:
    """`chromium.launch(**kwargs)` options for the given (or current) profile."""
    idx = _PROFILE["idx"] if profile_idx is None else profile_idx
    prof = LAUNCH_PROFILES[min(idx, len(LAUNCH_PROFILES) - 1)]
    kw = dict(
        headless=True,
        args=[*CHROMIUM_ARGS, *prof.get("extra_args", [])],
        chromium_sandbox=False,
        timeout=LAUNCH_TIMEOUT_MS,
        # Streamlit owns the process signals; don't let the driver tear the
        # browser down on a SIGHUP/SIGINT meant for the web server.
        handle_sigint=False,
        handle_sigterm=False,
        handle_sighup=False,
    )
    if prof.get("channel"):
        kw["channel"] = prof["channel"]
    return kw


def profile_name(idx: int | None = None) -> str:
    idx = _PROFILE["idx"] if idx is None else idx
    return LAUNCH_PROFILES[min(idx, len(LAUNCH_PROFILES) - 1)]["name"]


def is_closed_error(exc: BaseException) -> bool:
    """True if `exc` means the Playwright target/browser died."""
    if type(exc).__name__ == "TargetClosedError":
        return True
    msg = str(exc)
    return any(m in msg for m in _CLOSED_MARKERS)
BLOCKED_RESOURCE_TYPES = {"image", "media", "font"}
# Third-party analytics/ads have no part in the lookup; every request the page
# skips is RAM and CPU the renderer never spends. reCAPTCHA must NOT be here.
BLOCKED_URL_PARTS = (
    "google-analytics.com", "googletagmanager.com", "doubleclick.net",
    "facebook.net", "facebook.com/tr", "hotjar.com", "clarity.ms",
    "googlesyndication.com", "twitter.com/i/adsct", "linkedin.com/px",
)

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
# Chromium's renderer CHECK-fails on its first text layout when fontconfig
# finds *zero* fonts (slim Cloud image). The launch probe with a blank page
# passes, and the very first goto/click dies with TargetClosedError. So we
# vendor a Latin + Arabic font set and point fontconfig at it (no root).
_FONT_APT_PACKAGES = (
    "fonts-dejavu-core",   # Latin/Greek/Cyrillic (arch: all)
    "fonts-noto-core",     # Noto Sans/Naskh Arabic + many scripts (arch: all)
)
_SYSLIB_MARKER = ".ok.v4"  # bump to force re-vendor when the package set grows
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


def _extract_deb_python(deb_path: Path, root: Path) -> bool:
    """Pure-Python .deb extraction (ar archive → data.tar.{xz,gz,bz2,zst})."""
    import tarfile

    data = deb_path.read_bytes()
    if not data.startswith(b"!<arch>\n"):
        return False
    pos = 8
    while pos + 60 <= len(data):
        name = data[pos:pos + 16].decode("ascii", "replace").strip()
        size = int(data[pos + 48:pos + 58].decode("ascii").strip() or "0")
        body = data[pos + 60:pos + 60 + size]
        pos += 60 + size + (size & 1)
        if not name.startswith("data.tar"):
            continue
        if name.endswith(".zst"):
            try:
                from compression import zstd  # Python ≥ 3.14
            except ImportError:
                return False
            body = zstd.decompress(body)
            mode = "r:"
        else:
            mode = "r:*"
        with tarfile.open(fileobj=io.BytesIO(body), mode=mode) as tar:
            members = [m for m in tar.getmembers()
                       if not (m.name.startswith("/") or ".." in m.name.split("/"))]
            try:
                tar.extractall(root, members=members, filter="fully_trusted")
            except TypeError:  # Python < 3.12: no `filter`
                tar.extractall(root, members=members)
        return True
    return False


def _extract_deb(deb_path: Path, root: Path) -> bool:
    for cmd in (
        ["dpkg-deb", "-x", str(deb_path), str(root)],
        ["bash", "-lc", f"cd {deb_path.parent} && ar x {deb_path.name} && tar -xf data.tar.* -C {root}"],
    ):
        try:
            code, _ = _run(cmd, timeout=120)
        except (FileNotFoundError, subprocess.TimeoutExpired):
            continue
        if code == 0:
            return True
    try:
        return _extract_deb_python(deb_path, root)
    except Exception:
        return False


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
    pkgs = tuple(dict.fromkeys((*_CHROMIUM_APT_PACKAGES, *_FONT_APT_PACKAGES, *extra_pkgs)))

    # Reuse cache only when marker matches current package-set generation and
    # no extra packages were requested.
    if marker.exists() and lib_ok and not extra_pkgs:
        _prepend_ld_path(root)
        _setup_fontconfig(root)
        return f"syslibs ready: {root} ({_font_count(root)} fonts)"

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
        _setup_fontconfig(root)
        log.append(f"fonts available: {_font_count(root)}")
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


def _font_count(root: Path) -> int:
    dirs = [root / "usr" / "share" / "fonts", Path("/usr/share/fonts"), Path("/usr/local/share/fonts")]
    n = 0
    for d in dirs:
        if d.is_dir():
            n += sum(1 for f in d.rglob("*") if f.suffix.lower() in (".ttf", ".otf", ".ttc"))
    return n


def _setup_fontconfig(root: Path) -> None:
    """Write a self-contained fonts.conf and export FONTCONFIG_* for Chromium.

    The vendored libfontconfig looks for /etc/fonts/fonts.conf, which slim
    images lack ("Cannot load default config file") → no fonts → renderer
    crash on first layout. Our own config lists the vendored font dir plus
    the usual system dirs and keeps the cache under the writable root.
    """
    if not IS_LINUX:
        return
    conf_dir = root / "etc" / "fonts"
    cache_dir = root / "fc-cache"
    conf_dir.mkdir(parents=True, exist_ok=True)
    cache_dir.mkdir(parents=True, exist_ok=True)
    conf = conf_dir / "fonts.conf"
    conf.write_text(
        f"""<?xml version="1.0"?>
<!DOCTYPE fontconfig SYSTEM "fonts.dtd">
<fontconfig>
  <dir>{root / "usr" / "share" / "fonts"}</dir>
  <dir>/usr/share/fonts</dir>
  <dir>/usr/local/share/fonts</dir>
  <dir prefix="xdg">fonts</dir>
  <cachedir>{cache_dir}</cachedir>
  <alias><family>sans-serif</family><prefer>
    <family>DejaVu Sans</family><family>Noto Sans</family><family>Noto Sans Arabic</family>
  </prefer></alias>
  <alias><family>serif</family><prefer>
    <family>DejaVu Serif</family><family>Noto Serif</family><family>Noto Naskh Arabic</family>
  </prefer></alias>
  <alias><family>monospace</family><prefer><family>DejaVu Sans Mono</family></prefer></alias>
  <alias><family>Arial</family><prefer><family>DejaVu Sans</family></prefer></alias>
  <alias><family>Tahoma</family><prefer><family>DejaVu Sans</family><family>Noto Sans Arabic</family></prefer></alias>
</fontconfig>
""",
        encoding="utf-8",
    )
    os.environ["FONTCONFIG_FILE"] = str(conf)
    os.environ["FONTCONFIG_PATH"] = str(conf_dir)


_PROBE_HTML = (
    "<html><body><p>probe — اختبار المتصفح</p>"
    "<input id='civ' placeholder='الرقم المدني'><button id='go'>استعلام</button>"
    "</body></html>"
)


def _executable_path() -> str:
    from playwright.sync_api import sync_playwright

    with sync_playwright() as p:
        return p.chromium.executable_path or ""


def _raw_chromium_smoke() -> str:
    """Run the Chromium binary directly and return its own stderr.

    Playwright only reports "Target closed"; Chromium's stderr carries the
    real reason (fontconfig errors, CHECK failures, missing GL, OOM…).
    """
    try:
        exe = _in_thread(_executable_path)
    except Exception as exc:  # noqa: BLE001
        return f"executable lookup failed: {exc}"
    if not exe or not Path(exe).exists():
        return f"executable missing: {exe!r}"
    cmd = [exe, "--headless", *[a for a in CHROMIUM_ARGS if not a.startswith("--lang")],
           "--enable-logging=stderr", "--v=0", "--dump-dom",
           "data:text/html;charset=utf-8,<p>probe%20%D8%A7%D8%AE%D8%AA%D8%A8%D8%A7%D8%B1</p>"]
    try:
        code, out = _run(cmd, timeout=90)
    except subprocess.TimeoutExpired:
        return "raw chromium smoke: TIMEOUT"
    except Exception as exc:  # noqa: BLE001
        return f"raw chromium smoke failed to start: {exc}"
    return f"raw chromium smoke: exit {code}\n{out[-1500:]}"


def _launch_probe(profile_idx: int | None = None) -> str:
    """Start Chromium briefly; return its executable path on success.

    Checking `executable_path` alone is not enough – Playwright may point at a
    full browser while headless mode needs chromium-headless-shell. A real
    launch is the only reliable readiness check.
    """
    from playwright.sync_api import sync_playwright

    with sync_playwright() as p:
        browser = p.chromium.launch(**launch_kwargs(profile_idx))
        try:
            # Do what the first real row does: layout (innerText), visibility
            # checks, typing into an input, a click and a paint (screenshot).
            # A renderer that dies here would die on row 1 with TargetClosed.
            page = browser.new_page(locale="ar-KW")
            page.set_content(_PROBE_HTML)
            page.evaluate("() => document.body.innerText")
            page.wait_for_selector("#civ", state="visible", timeout=10_000)
            page.fill("#civ", "200000000000")
            page.click("#go")
            page.screenshot(type="png")
            return p.chromium.executable_path or "chromium"
        finally:
            browser.close()


def _chromium_works(profile_idx: int | None = None) -> str | None:
    """Return executable path if a headless launch succeeds, else None."""
    try:
        return _in_thread(lambda: _launch_probe(profile_idx))
    except Exception:
        return None


def _last_launch_error(profile_idx: int | None = None) -> str:
    try:
        _in_thread(lambda: _launch_probe(profile_idx))
        return ""
    except Exception as exc:  # noqa: BLE001
        return f"{type(exc).__name__}: {exc}"


def _pick_profile(notes: list[str]) -> str | None:
    """Try launch profiles in order; return exe of the first that passes."""
    for idx in range(len(LAUNCH_PROFILES)):
        exe = _chromium_works(idx)
        if exe:
            _PROFILE["idx"] = idx
            if idx:
                notes.append(f"launch profile #{idx} '{profile_name(idx)}' selected "
                             "(earlier profiles crashed in the probe)")
            return exe
        notes.append(f"profile '{profile_name(idx)}' failed: {_last_launch_error(idx)[:300]}")
    return None


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
        exe = _chromium_works(0)
        if exe:
            _PROFILE["idx"] = 0
            return f"Chromium ready [{profile_name()}]: {exe}\n" + "\n".join(notes)
        err = _last_launch_error(0)
        if not IS_LINUX or not _missing_soname(err):
            notes.append(f"launch still failing (attempt {attempt + 1}): {err}")
            break
        notes.append(_fix_missing_so(err))

    # Libraries are fine but the default profile crashes when it renders /
    # types – try the other browser configurations before giving up.
    exe = _pick_profile(notes)
    if exe:
        return f"Chromium ready [{profile_name()}]: {exe}\n" + "\n".join(notes)

    launch_err = _last_launch_error()
    raise RuntimeError(
        "Chromium downloaded but the headless probe failed (missing shared "
        "libraries, no fonts, or a renderer crash).\n"
        f"Launch error: {launch_err}\n"
        f"{_raw_chromium_smoke()}\n"
        f"FONTCONFIG_FILE={os.environ.get('FONTCONFIG_FILE', '')}\n"
        + "\n".join(notes)
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
ERR_BLOCKED = "⚠ الموقع لم يستجب (صفحة غير متوقعة)"
RETRYABLE = (ERR_TIMEOUT, ERR_BLOCKED)
POLL_MS = 250          # wait_for_function interval; rAF polling stalls in background windows
FORM_WAIT_MS = 15_000  # how long the MOI form itself may take to appear


def _descendant_pids(root_pid: int) -> list[int]:
    """PIDs of every process under `root_pid` (Linux via /proc, else `ps`)."""
    children: dict[int, list[int]] = {}
    try:
        if IS_LINUX:
            for d in Path("/proc").iterdir():
                if not d.name.isdigit():
                    continue
                try:
                    stat = (d / "stat").read_text()
                    ppid = int(stat.rsplit(")", 1)[1].split()[1])
                except Exception:
                    continue
                children.setdefault(ppid, []).append(int(d.name))
        else:
            out = subprocess.run(["ps", "-axo", "pid=,ppid="], capture_output=True, text=True, timeout=10).stdout
            for line in out.splitlines():
                parts = line.split()
                if len(parts) == 2:
                    children.setdefault(int(parts[1]), []).append(int(parts[0]))
    except Exception:
        return []
    found, stack = [], [root_pid]
    while stack:
        p = stack.pop()
        for c in children.get(p, []):
            found.append(c)
            stack.append(c)
    return found


def _cmdline(pid: int) -> str:
    try:
        if IS_LINUX:
            return (Path(f"/proc/{pid}/cmdline").read_bytes().replace(b"\0", b" ")).decode("utf-8", "replace")
        return subprocess.run(["ps", "-o", "command=", "-p", str(pid)], capture_output=True, text=True, timeout=10).stdout
    except Exception:
        return ""


def kill_chromium_processes() -> int:
    """SIGKILL every Chromium descended from this process. Returns the count.

    Used by the watchdog when a Playwright call exceeds its hard deadline: the
    pending call then fails with TargetClosedError and the scraper relaunches.
    Only Chromium is killed – the Node driver stays up and reports the loss.
    """
    import signal

    killed = 0
    for pid in _descendant_pids(os.getpid()):
        # Browser binaries: chrome, chrome-headless-shell, headless_shell under
        # .../chromium_headless_shell-NNNN/. The Node driver never matches.
        cmd = _cmdline(pid).lower()
        if "chrom" in cmd or "headless_shell" in cmd:
            try:
                os.kill(pid, signal.SIGKILL)
                killed += 1
            except Exception:
                pass
    return killed


def memory_report() -> str:
    """RSS of this process and of every Chromium/driver child (Linux only)."""
    if not IS_LINUX:
        return ""

    def rss_mb(pid: int) -> float:
        try:
            for line in Path(f"/proc/{pid}/status").read_text().splitlines():
                if line.startswith("VmRSS:"):
                    return int(line.split()[1]) / 1024
        except Exception:
            pass
        return 0.0

    own = rss_mb(os.getpid())
    kids = _descendant_pids(os.getpid())
    browser = sum(rss_mb(pid) for pid in kids)
    limit = ""
    for f in ("/sys/fs/cgroup/memory.max", "/sys/fs/cgroup/memory/memory.limit_in_bytes"):
        try:
            v = Path(f).read_text().strip()
            if v.isdigit() and int(v) < 1 << 50:
                limit = f" / limit {int(v) / 1024 / 1024:.0f} MB"
                break
        except Exception:
            continue
    return f"RAM: app {own:.0f} MB + browser {browser:.0f} MB ({len(kids)} procs){limit}"


class GovScraper:
    UA = (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"
    )

    def __init__(self, headless: bool = True, timeout_s: int = 40, log=None):
        self.headless = headless
        self.timeout_ms = timeout_s * 1000
        # One lookup = at most goto + wait (each ≤ timeout) + slack. If Playwright
        # is still blocked past this, the watchdog kills Chromium.
        self.hard_limit_s = timeout_s * 2 + 30
        self._log = log or (lambda _m: None)
        self.relaunches = 0
        self.closed_errors_in_a_row = 0
        self.op_deadline: float | None = None   # monotonic; read by the watchdog
        self._pw = None
        self.browser = None
        self.context = None
        self.page = None
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
            self.op_deadline = time.monotonic() + LAUNCH_TIMEOUT_MS / 1000 + 30
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
                # ONE page for both sites. A second page meant a second
                # renderer (+ reCAPTCHA frames) – on a 1 GB container that is
                # the difference between working and an OOM-killed target.
                # The job orders PACI status → renewal → MOI so the PACI page is
                # loaded once per row and MOI navigates away afterwards.
                self.page = self.context.new_page()
                self.moi_page = self.paci_page = self.page
                self.crashes = 0
                self.page.on("crash", lambda _pg: self._on_crash("main"))
                self._paci_loaded = False
                self.op_deadline = None
                return
            except Exception as exc:  # noqa: BLE001
                last = exc
                self._teardown()
                self._log(f"فشل تشغيل المتصفح [{profile_name()}] (محاولة {attempt}/{LAUNCH_ATTEMPTS}): "
                          f"{type(exc).__name__}: {exc}")
                if attempt == LAUNCH_ATTEMPTS - 1:
                    self.next_profile("فشل التشغيل")
                time.sleep(2 * attempt)
        self.op_deadline = None
        raise RuntimeError(f"Chromium launch failed after {LAUNCH_ATTEMPTS} attempts: {last}")

    def _on_crash(self, name: str):
        self.crashes = getattr(self, "crashes", 0) + 1
        self._log(f"💥 تعطل محرك العرض لصفحة {name} (renderer crash) — سيُعاد تشغيل المتصفح.")

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
        self.page = self.moi_page = self.paci_page = None
        self._paci_loaded = False

    def alive(self) -> bool:
        try:
            return (
                self.browser is not None
                and self.browser.is_connected()
                and self.page is not None and not self.page.is_closed()
            )
        except Exception:
            return False

    def relaunch(self, reason: str = ""):
        """Kill whatever is left and bring up a fresh browser."""
        self.relaunches += 1
        self._log(f"إعادة تشغيل المتصفح ({self.relaunches}) [{profile_name()}] {reason}".rstrip())
        self._teardown()
        kill_chromium_processes()  # nothing survives a relaunch (no zombie RAM)
        self._launch()

    def ensure_alive(self):
        if not self.alive():
            self.relaunch("— المتصفح أُغلق أو تعطل")

    def next_profile(self, why: str) -> bool:
        """Switch to the next launch profile (if any). Process-wide."""
        if _PROFILE["idx"] + 1 >= len(LAUNCH_PROFILES):
            return False
        _PROFILE["idx"] += 1
        self._log(f"🔁 تبديل إعدادات تشغيل المتصفح إلى '{profile_name()}' ({why}).")
        return True

    @staticmethod
    def _route(route, request):
        if request.resource_type in BLOCKED_RESOURCE_TYPES:
            return route.abort()
        url = request.url
        if any(part in url for part in BLOCKED_URL_PARTS):
            return route.abort()
        return route.continue_()

    def state(self) -> str:
        """One-line diagnostic used in logs after a failure."""
        try:
            connected = self.browser is not None and self.browser.is_connected()
        except Exception:
            connected = False
        try:
            closed = self.page is None or self.page.is_closed()
            url = self.page.url if not closed else "-"
        except Exception:
            closed, url = True, "-"
        return (f"browser connected={connected} page closed={closed} crashes={self.crashes} "
                f"url={url[:80]} {memory_report()}").strip()

    # ---------------------------------------------------------------- MOI --- #
    def moi_fines(self, civil_id: str) -> str:
        page = self.moi_page
        # domcontentloaded: don't wait for every third-party script/tracker to
        # finish; the readiness check below is what actually matters.
        page.goto(MOI_URL + civil_id, wait_until="domcontentloaded")

        # Fail fast when the site serves a WAF/maintenance/error page instead of
        # the enquiry form – otherwise every row burns the full timeout.
        try:
            page.wait_for_selector("#civilId", state="attached",
                                   timeout=min(FORM_WAIT_MS, self.timeout_ms))
        except Exception:
            title = ""
            try:
                title = page.title()
            except Exception:
                pass
            return f"{ERR_BLOCKED}: {title[:60]}" if title else ERR_BLOCKED

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
            page.wait_for_function(js_ready, timeout=self.timeout_ms, polling=POLL_MS)
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
            self._paci_loaded = False
            page.goto(PACI_URL, wait_until="domcontentloaded")
            page.wait_for_selector("#txtCivilId", state="visible")
            page.wait_for_selector("#btnSearch", state="visible")
            try:  # give reCAPTCHA v3 a moment to populate the hidden token
                page.wait_for_function(
                    "() => ((document.querySelector('#Token')||{}).value||'').length > 20",
                    timeout=15000, polling=POLL_MS,
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
            timeout=5000, polling=POLL_MS,
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
            page.wait_for_function(js_ready, timeout=self.timeout_ms, polling=POLL_MS)
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
    def _guarded(self, fn, civil_id: str) -> str:
        """Run one lookup under the watchdog deadline."""
        self.op_deadline = time.monotonic() + self.hard_limit_s
        try:
            return fn(civil_id)
        finally:
            self.op_deadline = None

    def safe(self, fn, civil_id: str, retries: int = 1) -> str:
        last = ERR_GENERIC
        for attempt in range(retries + 1):
            try:
                self.ensure_alive()
                out = self._guarded(fn, civil_id)
                self.closed_errors_in_a_row = 0
                if not out.startswith(RETRYABLE) or attempt == retries:
                    return out
                last = out
            except Exception as exc:  # noqa: BLE001
                last = f"{ERR_GENERIC}: {type(exc).__name__}"
                self._paci_loaded = False
                self._log(f"❌ {civil_id} {getattr(fn, '__name__', '?')}: {type(exc).__name__}: "
                          f"{str(exc).splitlines()[0][:160]} | {self.state()}")
                if is_closed_error(exc):
                    # Browser/renderer died (OOM, crash, watchdog kill). Relaunch
                    # and give the same row one more chance instead of failing.
                    # Twice in a row with no success in between → this browser
                    # configuration cannot render these pages here: switch.
                    self.closed_errors_in_a_row += 1
                    if self.closed_errors_in_a_row >= 2:
                        if self.next_profile(f"{type(exc).__name__} ×{self.closed_errors_in_a_row}"):
                            self.closed_errors_in_a_row = 0
                    try:
                        self.relaunch(f"— {type(exc).__name__}")
                    except Exception as launch_exc:  # noqa: BLE001
                        return f"{ERR_GENERIC}: {type(launch_exc).__name__}"
                    if attempt == retries:
                        try:
                            return self._guarded(fn, civil_id)
                        except Exception as exc2:  # noqa: BLE001
                            return f"{ERR_GENERIC}: {type(exc2).__name__}"
            time.sleep(1.5)
        return last

    def close(self):
        self._teardown()
        kill_chromium_processes()


# --------------------------------------------------------------------------- #
# Background job (thread) – survives Streamlit reruns and page reloads
# --------------------------------------------------------------------------- #
FAIL_ROWS_PER_COOLDOWN = 5     # consecutive all-error rows before a relaunch + pause
MAX_FAIL_ROWS = 40             # give up after this many consecutive all-error rows
COOLDOWN_BASE_S = 30
COOLDOWN_MAX_S = 300
WATCHDOG_TICK_S = 5


class ScrapeJob:
    def __init__(self, df: pd.DataFrame, opts: dict):
        self.df = df
        self.opts = opts
        self.total = len(df)
        self.processed = 0
        self.current = ""
        self.started_at = datetime.now()
        self.finished_at: datetime | None = None
        self.logs: list[str] = []
        self.error: str | None = None
        self.consecutive_fail_rows = 0
        self.cooldowns = 0
        self.watchdog_kills = 0
        self.lock = threading.Lock()
        self.stop_event = threading.Event()
        self.done = threading.Event()
        self.scraper: GovScraper | None = None
        self._excel_cache: tuple[int, bytes] | None = None
        self.thread = threading.Thread(target=self._run, name="scraper", daemon=True)
        self.watchdog = threading.Thread(target=self._watchdog, name="scraper-watchdog", daemon=True)

    # -- API used by the UI thread --
    def start(self):
        self.thread.start()
        self.watchdog.start()

    @property
    def running(self) -> bool:
        return not self.done.is_set()

    def snapshot(self) -> pd.DataFrame:
        with self.lock:
            return self.df.copy()

    def excel_bytes(self) -> bytes:
        """Excel of the current state; rebuilt only when progress changed."""
        key = self.processed if self.running else -1
        cache = self._excel_cache
        if cache is not None and cache[0] == key:
            return cache[1]
        data = to_excel_bytes(self.snapshot())
        self._excel_cache = (key, data)
        return data

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

    # -- watchdog: no Playwright call may block past its hard deadline --
    def _watchdog(self):
        while not self.done.wait(WATCHDOG_TICK_S):
            s = self.scraper
            deadline = s.op_deadline if s is not None else None
            if deadline is None or time.monotonic() < deadline:
                continue
            s.op_deadline = None
            n = kill_chromium_processes()
            self.watchdog_kills += 1
            self.log(f"⏱ المراقب: تجاوز الحد الأقصى للعملية — تم إنهاء Chromium ({n} عملية) وسيُعاد تشغيله.")

    def _cooldown(self, scraper: GovScraper):
        self.cooldowns += 1
        wait = min(COOLDOWN_BASE_S * (2 ** (self.cooldowns - 1)), COOLDOWN_MAX_S)
        self.log(f"⚠ {self.consecutive_fail_rows} صفوف متتالية بلا نتيجة — إعادة تشغيل المتصفح والانتظار {wait} ث.")
        try:
            scraper.relaunch("— بعد فشل متكرر")
        except Exception as exc:  # noqa: BLE001
            self.log(f"تعذر إعادة التشغيل: {type(exc).__name__}: {exc}")
        self.stop_event.wait(wait)

    # -- worker --
    def _run(self):
        o = self.opts
        # PACI first (both queries on one page load), then MOI navigates away.
        services = [
            (COL_STATUS, o["do_status"], "paci_card_status", "حالة البطاقة"),
            (COL_RENEW, o["do_renew"], "paci_card_renewal", "تجديد البطاقة"),
            (COL_MOI, o["do_moi"], "moi_fines", "الداخلية"),
        ]
        scraper = None
        try:
            scraper = GovScraper(headless=o["headless"], timeout_s=o["timeout_s"], log=self.log)
            self.scraper = scraper
            self.log(f"تم تشغيل Chromium. {memory_report()}".strip())
            rows_since_restart = 0

            for i, (idx, row) in enumerate(self.df.iterrows(), start=1):
                if self.stop_event.is_set():
                    self.log("تم الإيقاف بواسطة المستخدم.")
                    break

                cid = row[COL_CID]
                self.current = cid or "—"

                if not cid or len(cid) != 12:
                    for col, enabled, _fn, _lbl in services:
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

                t0 = time.monotonic()
                attempted = errors = 0
                for col, enabled, fn_name, label in services:
                    if self.stop_event.is_set():
                        break
                    if not enabled or not self._needs(idx, col):
                        continue
                    val = scraper.safe(getattr(scraper, fn_name), cid, o["retries"])
                    self._set(idx, col, val)
                    attempted += 1
                    errors += is_error(val)
                    self.log(f"{cid} | {label}: {val}")

                self.processed = i
                rows_since_restart += 1
                mem = f" | {memory_report()}" if IS_LINUX and i % 10 == 0 else ""
                self.log(f"{cid} | الصف {i}/{self.total} في {time.monotonic() - t0:.1f} ث{mem}")

                # Circuit breaker: a site that stopped answering should not cost
                # 735 × timeout. Pause with back-off, then give up with a clear
                # message instead of looking "hung" for hours.
                if attempted and errors == attempted:
                    self.consecutive_fail_rows += 1
                    if self.consecutive_fail_rows >= MAX_FAIL_ROWS:
                        raise RuntimeError(
                            f"{MAX_FAIL_ROWS} صفاً متتالياً بلا أي نتيجة — يبدو أن الموقع يحظر "
                            "الطلبات أو متوقف. حمّل النتائج الجزئية وأعد المحاولة لاحقاً "
                            "(مع تفعيل 'تخطي الخلايا المعبأة')."
                        )
                    if self.consecutive_fail_rows % FAIL_ROWS_PER_COOLDOWN == 0:
                        self._cooldown(scraper)
                        rows_since_restart = 0
                elif attempted:
                    self.consecutive_fail_rows = 0
                    self.cooldowns = 0

                if o["delay_s"] and not self.stop_event.is_set():
                    self.stop_event.wait(o["delay_s"])

            if not self.stop_event.is_set():
                self.log("انتهى التنفيذ.")
        except Exception as exc:  # noqa: BLE001
            self.error = f"{type(exc).__name__}: {exc}"
            self.log(f"خطأ عام: {self.error}")
        finally:
            if scraper:
                scraper.close()
                self.log("تم إغلاق المتصفح.")
            self.scraper = None
            self.finished_at = datetime.now()
            self.done.set()


# --------------------------------------------------------------------------- #
# One-ID diagnostic (sidebar) – shows the real exception instead of a cell
# --------------------------------------------------------------------------- #
def run_quick_test(civil_id: str, timeout_s: int) -> str:
    import traceback

    lines: list[str] = []
    logs: list[str] = []
    t_all = time.monotonic()
    try:
        s = GovScraper(headless=True, timeout_s=timeout_s, log=logs.append)
    except Exception:  # noqa: BLE001
        return "فشل تشغيل المتصفح:\n" + traceback.format_exc() + "\n" + "\n".join(logs)
    try:
        lines.append(memory_report())
        for label, fn in (("حالة البطاقة", s.paci_card_status), ("تجديد البطاقة", s.paci_card_renewal),
                          ("الداخلية", s.moi_fines)):
            t0 = time.monotonic()
            try:
                out = fn(civil_id)
                lines.append(f"✅ {label} ({time.monotonic() - t0:.1f}s): {out}")
            except Exception as exc:  # noqa: BLE001
                lines.append(f"❌ {label} ({time.monotonic() - t0:.1f}s): {type(exc).__name__}: {exc}")
                lines.append(traceback.format_exc()[-1500:])
                lines.append(s.state())
                if is_closed_error(exc):
                    s.relaunch("— quick test")
        lines.append(f"المجموع: {time.monotonic() - t_all:.1f}s | relaunches={s.relaunches} | profile={profile_name()}")
    finally:
        s.close()
    if IS_LINUX:
        lines.append(f"fonts: {_font_count(_syslib_root())} | FONTCONFIG_FILE={os.environ.get('FONTCONFIG_FILE', '-')}")
    return "\n".join(lines + ["", *logs])


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
# Process-wide job registry
# --------------------------------------------------------------------------- #
# One container = one Chromium budget. The job lives here (not in
# session_state) so a page reload / dropped websocket / second tab re-attaches
# to the running job instead of silently starting a second browser and
# OOM-killing both.
@st.cache_resource(show_spinner=False)
def _job_registry() -> dict:
    return {"job": None}


def current_job() -> ScrapeJob | None:
    return _job_registry()["job"]


def set_job(job: ScrapeJob | None) -> None:
    _job_registry()["job"] = job


# --------------------------------------------------------------------------- #
# i18n
# --------------------------------------------------------------------------- #
LANGS = {"ar": "العربية", "en": "English"}
T: dict[str, dict[str, str]] = {
    "ar": {
        "page_title": "استعلام الداخلية و الهيئة المدنية",
        "app_title": "الاستعلام الآلي عن الغرامات والبطاقة المدنية",
        "app_subtitle": "غرامات الإقامة (وزارة الداخلية) • حالة البطاقة وتجديدها (الهيئة العامة للمعلومات المدنية)",
        "hero_badge": "معالجة دفعات Excel",
        "boot_spinner": "جاري تجهيز المتصفح (يحدث مرة واحدة عند أول تشغيل)…",
        "boot_error": "تعذر تثبيت/تشغيل Chromium. راجع سجل الخطأ أدناه و requirements.txt.",
        "settings": "الإعدادات",
        "cloud_note": "بيئة خادم بدون شاشة — المتصفح يعمل في الخلفية إجبارياً.",
        "headless": "تشغيل المتصفح في الخلفية (Headless)",
        "timeout": "مهلة انتظار النتيجة (ثانية)",
        "delay": "فاصل زمني بين كل رقم مدني (ثانية)",
        "retries": "عدد إعادة المحاولة عند انتهاء المهلة",
        "restart_every": "إعادة تشغيل المتصفح كل N صف (لتوفير الذاكرة)",
        "skip_filled": "تخطي الخلايا المعبأة مسبقاً",
        "skip_filled_help": "لاستكمال ملف سبق تحميله جزئياً: ارفع ملف النتائج وسيُكمل من حيث توقف.",
        "services": "الخدمات",
        "svc_moi": "الداخلية — غرامات الإقامة",
        "svc_status": "الهيئة المدنية — حالة البطاقة",
        "svc_renew": "الهيئة المدنية — تجديد البطاقة",
        "sysinfo": "معلومات النظام",
        "started": "بدأ",
        "relaunches": "إعادة تشغيل المتصفح",
        "watchdog": "تدخلات المراقب",
        "cooldowns": "فترات تهدئة",
        "boot_log": "سجل تجهيز المتصفح الكامل:",
        "quick_test": "اختبار سريع لرقم واحد",
        "quick_test_help": "يستعلم عن رقم واحد ويعرض الخطأ الكامل إن حدث (مفيد للتشخيص).",
        "test_cid": "الرقم المدني للاختبار",
        "run_test": "تشغيل الاختبار",
        "testing": "جاري الاختبار…",
        "upload_title": "رفع الملف",
        "upload_label": "اختر ملف Excel",
        "upload_help": "يجب أن يحتوي الملف الأعمدة: {cols} — تُعبأ الأعمدة الثلاثة الأخيرة تلقائياً.",
        "read_error": "تعذر قراءة الملف: {err}",
        "missing_cols": "الأعمدة التالية غير موجودة في الملف: {cols}",
        "existing_cols": "الأعمدة الموجودة:",
        "file_loaded": "تم تحميل الملف: {name}",
        "running_banner": "يوجد استعلام قيد التنفيذ منذ {time} — يمكنك إغلاق الصفحة والعودة لاحقاً؛ التقدم محفوظ على الخادم.",
        "preview_title": "معاينة الملف",
        "stat_rows": "إجمالي الصفوف",
        "stat_valid": "أرقام مدنية صالحة",
        "stat_processed": "تمت معالجتها",
        "stat_errors": "خلايا بها خطأ",
        "start": "بدء الاستعلام",
        "stop": "إيقاف",
        "reset": "إعادة ضبط / مسح العملية",
        "reset_help": "يمسح نتائج العملية الحالية والملف المرفوع ويبدأ من جديد دون إعادة تشغيل الخادم.",
        "reset_done": "تم مسح العملية. يمكنك رفع ملف جديد.",
        "progress_title": "تقدم التنفيذ",
        "progress_text": "الصف {done} من {total} — {cid}",
        "eta_text": "جاري الاستعلام… {cid} — ~{per_row:.0f} ث/صف — الوقت المتبقي التقريبي: {eta:.0f} دقيقة",
        "err_stopped": "توقف التنفيذ بسبب خطأ: {err}",
        "stopped_warning": "تم إيقاف التنفيذ. يمكنك تحميل النتائج الجزئية أدناه.",
        "finished_ok": "انتهى الاستعلام لجميع الصفوف ✅",
        "preview_window": "عرض الصفوف {lo}–{hi} (النتائج الكاملة في ملف التحميل)",
        "log_title": "سجل التنفيذ",
        "download_partial": "تحميل Excel (النتائج الجزئية حتى الآن)",
        "download_final": "تحميل ملف Excel المحدث",
        "status_idle": "جاهز",
        "status_running": "قيد التنفيذ",
        "status_finished": "مكتمل",
        "status_stopped": "متوقف",
        "status_failed": "فشل",
        "empty_hint": "ارفع ملف Excel للبدء.",
        "lang": "اللغة",
    },
    "en": {
        "page_title": "MOI & PACI Bulk Lookup",
        "app_title": "Automated Fines & Civil ID Card Lookup",
        "app_subtitle": "Residence fines (Ministry of Interior) • Card status and renewal (Public Authority for Civil Information)",
        "hero_badge": "Excel batch processing",
        "boot_spinner": "Preparing the browser (one-time setup on first run)…",
        "boot_error": "Could not install/launch Chromium. See the error log below and requirements.txt.",
        "settings": "Settings",
        "cloud_note": "Headless server environment — the browser always runs in the background.",
        "headless": "Run browser in background (headless)",
        "timeout": "Result timeout (seconds)",
        "delay": "Delay between civil IDs (seconds)",
        "retries": "Retries on timeout",
        "restart_every": "Restart browser every N rows (memory)",
        "skip_filled": "Skip already-filled cells",
        "skip_filled_help": "To resume a partially completed file: upload the results file and it continues where it stopped.",
        "services": "Services",
        "svc_moi": "MOI — Residence fines",
        "svc_status": "PACI — Card status",
        "svc_renew": "PACI — Card renewal",
        "sysinfo": "System info",
        "started": "Started",
        "relaunches": "Browser relaunches",
        "watchdog": "Watchdog interventions",
        "cooldowns": "Cool-down pauses",
        "boot_log": "Full browser bootstrap log:",
        "quick_test": "Quick test (single ID)",
        "quick_test_help": "Looks up one civil ID and shows the full error if any (useful for diagnosis).",
        "test_cid": "Civil ID to test",
        "run_test": "Run test",
        "testing": "Testing…",
        "upload_title": "Upload file",
        "upload_label": "Choose an Excel file",
        "upload_help": "The file must contain the columns: {cols} — the last three are filled automatically.",
        "read_error": "Could not read the file: {err}",
        "missing_cols": "These columns are missing from the file: {cols}",
        "existing_cols": "Columns found:",
        "file_loaded": "File loaded: {name}",
        "running_banner": "A lookup has been running since {time} — you can close this page and come back later; progress is kept on the server.",
        "preview_title": "File preview",
        "stat_rows": "Total rows",
        "stat_valid": "Valid civil IDs",
        "stat_processed": "Processed",
        "stat_errors": "Cells with errors",
        "start": "Start lookup",
        "stop": "Stop",
        "reset": "Reset / Clear job",
        "reset_help": "Clears the current job results and the uploaded file so you can start fresh without restarting the server.",
        "reset_done": "Job cleared. You can upload a new file.",
        "progress_title": "Progress",
        "progress_text": "Row {done} of {total} — {cid}",
        "eta_text": "Looking up… {cid} — ~{per_row:.0f} s/row — estimated time left: {eta:.0f} min",
        "err_stopped": "Stopped because of an error: {err}",
        "stopped_warning": "Stopped. You can download the partial results below.",
        "finished_ok": "Lookup finished for all rows ✅",
        "preview_window": "Showing rows {lo}–{hi} (full results are in the download)",
        "log_title": "Execution log",
        "download_partial": "Download Excel (partial results so far)",
        "download_final": "Download updated Excel file",
        "status_idle": "Ready",
        "status_running": "Running",
        "status_finished": "Finished",
        "status_stopped": "Stopped",
        "status_failed": "Failed",
        "empty_hint": "Upload an Excel file to begin.",
        "lang": "Language",
    },
}


def _resolve_lang() -> str:
    ss = st.session_state
    if "lang" not in ss:
        q = st.query_params.get("lang", "ar")
        ss.lang = q if q in LANGS else "ar"
    return ss.lang


def t(key: str, **kw) -> str:
    text = T.get(LANG, T["ar"]).get(key) or T["ar"].get(key) or key
    return text.format(**kw) if kw else text


# --------------------------------------------------------------------------- #
# Styling
# --------------------------------------------------------------------------- #
def inject_css(rtl: bool) -> None:
    direction = "rtl" if rtl else "ltr"
    align = "right" if rtl else "left"
    font = "'Tajawal', 'Inter', system-ui, sans-serif" if rtl else "'Inter', 'Tajawal', system-ui, sans-serif"
    st.markdown(
        f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=Tajawal:wght@400;500;700;800&display=swap');

html, body, .stApp, [data-testid="stSidebar"] {{
  font-family: {font};
}}
.stApp {{ direction: {direction}; }}
[data-testid="stSidebar"] {{ direction: {direction}; }}
.stMarkdown, .stButton, .stDownloadButton, .stDataFrame, .stAlert, label, .stCaption {{
  direction: {direction}; text-align: {align};
}}
.stApp {{
  background:
    radial-gradient(1200px 500px at {'100%' if rtl else '0%'} -10%, rgba(0,120,120,.12), transparent 60%),
    radial-gradient(900px 400px at {'0%' if rtl else '100%'} 0%, rgba(0,80,200,.10), transparent 60%),
    var(--background-color);
}}
/* Hide the default footer / decorations */
footer {{ visibility: hidden; }}
[data-testid="stDecoration"] {{ display: none; }}
.block-container {{ padding-top: 4.4rem; padding-bottom: 3rem; max-width: 1200px; }}

/* Hero */
.kb-hero {{
  position: relative; overflow: hidden;
  border-radius: 22px; padding: 26px 30px; margin-bottom: 18px;
  background: linear-gradient(135deg, #0f766e 0%, #0e7490 45%, #1d4ed8 100%);
  color: #fff; box-shadow: 0 18px 45px -22px rgba(2, 44, 80, .55);
}}
.kb-hero::after {{
  content: ""; position: absolute; inset: 0;
  background: radial-gradient(600px 220px at {'0%' if rtl else '100%'} 0%, rgba(255,255,255,.18), transparent 60%);
  pointer-events: none;
}}
.kb-hero h1 {{ font-size: 1.85rem; margin: 0 0 6px 0; font-weight: 800; letter-spacing: -.01em; color: #fff; }}
.kb-hero p  {{ margin: 0; opacity: .92; font-size: 1rem; }}
.kb-badge {{
  display: inline-block; font-size: .78rem; font-weight: 600; letter-spacing: .04em;
  padding: 4px 10px; border-radius: 999px; background: rgba(255,255,255,.18); margin-bottom: 10px;
}}

/* Cards (bordered containers) */
[data-testid="stVerticalBlockBorderWrapper"] > div:first-child {{
  border-radius: 18px !important;
  border: 1px solid rgba(128,128,128,.18) !important;
  background: var(--secondary-background-color);
  box-shadow: 0 10px 30px -22px rgba(0,0,0,.35);
  padding: 1.1rem 1.2rem !important;
}}
.kb-card-title {{
  font-weight: 700; font-size: 1.05rem; margin: 0 0 .6rem 0; display: flex; align-items: center; gap: .5rem;
}}

/* Stat tiles */
.kb-stats {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: 12px; margin: 6px 0 14px; }}
.kb-stat {{
  border-radius: 14px; padding: 12px 14px;
  background: var(--background-color); border: 1px solid rgba(128,128,128,.16);
}}
.kb-stat .v {{ font-size: 1.6rem; font-weight: 800; line-height: 1.1; }}
.kb-stat .l {{ font-size: .8rem; opacity: .7; margin-top: 2px; }}
.kb-stat.ok .v {{ color: #059669; }}
.kb-stat.bad .v {{ color: #dc2626; }}
.kb-stat.info .v {{ color: #2563eb; }}

/* Status pill */
.kb-pill {{
  display: inline-flex; align-items: center; gap: 6px; font-size: .8rem; font-weight: 700;
  padding: 4px 12px; border-radius: 999px; border: 1px solid transparent;
}}
.kb-pill .dot {{ width: 8px; height: 8px; border-radius: 50%; background: currentColor; }}
.kb-pill.idle     {{ color: #475569; background: rgba(71,85,105,.12); }}
.kb-pill.running  {{ color: #2563eb; background: rgba(37,99,235,.12); }}
.kb-pill.running .dot {{ animation: kb-blink 1.2s infinite; }}
.kb-pill.finished {{ color: #059669; background: rgba(5,150,105,.12); }}
.kb-pill.stopped  {{ color: #d97706; background: rgba(217,119,6,.12); }}
.kb-pill.failed   {{ color: #dc2626; background: rgba(220,38,38,.12); }}
@keyframes kb-blink {{ 0%,100% {{ opacity: 1 }} 50% {{ opacity: .25 }} }}

/* Buttons */
.stButton > button, .stDownloadButton > button {{
  border-radius: 12px !important; font-weight: 700 !important; padding: .6rem 1rem !important;
  transition: transform .08s ease, box-shadow .15s ease;
}}
.stButton > button:hover, .stDownloadButton > button:hover {{ transform: translateY(-1px); }}
.stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"] {{
  background: linear-gradient(135deg, #0f766e, #1d4ed8) !important; border: none !important;
  box-shadow: 0 10px 24px -12px rgba(29,78,216,.6);
}}
[data-testid="stFileUploaderDropzone"] {{ border-radius: 14px; }}
[data-testid="stProgressBar"] > div > div {{ border-radius: 999px; }}
[data-testid="stProgressBar"] > div > div > div {{ background: linear-gradient(90deg, #0f766e, #1d4ed8); border-radius: 999px; }}
[data-testid="stSegmentedControl"] {{ margin-top: 14px; }}
[data-testid="stSegmentedControl"] button {{ border-radius: 10px !important; font-weight: 600; }}
[data-testid="stSidebar"] [data-testid="stExpander"] details {{ border-radius: 12px; }}
/* Sliders/segmented controls are LTR widgets: keep their geometry LTR, labels stay {direction} */
[data-testid="stSlider"] > div, [data-testid="stSegmentedControl"] > div {{ direction: ltr; }}
[data-testid="stSlider"] label, [data-testid="stSlider"] [data-testid="stWidgetLabel"] {{ direction: {direction}; text-align: {align}; }}
</style>
""",
        unsafe_allow_html=True,
    )


def hero(title: str, subtitle: str, badge: str) -> None:
    st.markdown(
        f'<div class="kb-hero"><span class="kb-badge">{badge}</span>'
        f"<h1>🇰🇼 {title}</h1><p>{subtitle}</p></div>",
        unsafe_allow_html=True,
    )


def card_title(icon: str, text: str, pill_html: str = "") -> None:
    st.markdown(
        f'<div class="kb-card-title"><span>{icon}</span><span>{text}</span>'
        f'<span style="flex:1"></span>{pill_html}</div>',
        unsafe_allow_html=True,
    )


def status_pill(state: str) -> str:
    return f'<span class="kb-pill {state}"><span class="dot"></span>{t("status_" + state)}</span>'


def stat_tiles(items: list[tuple[str, object, str]]) -> None:
    tiles = "".join(f'<div class="kb-stat {cls}"><div class="v">{val}</div><div class="l">{label}</div></div>'
                    for label, val, cls in items)
    st.markdown(f'<div class="kb-stats">{tiles}</div>', unsafe_allow_html=True)


def job_tiles(df: pd.DataFrame, j: ScrapeJob | None) -> None:
    valid = int((df[COL_CID].str.len() == 12).sum())
    snap = j.snapshot() if j is not None else df
    errors = int(sum(snap[c].apply(is_error).sum() for c in (COL_MOI, COL_STATUS, COL_RENEW)))
    stat_tiles([
        (t("stat_rows"), len(df), "info"),
        (t("stat_valid"), valid, "ok" if valid == len(df) else ""),
        (t("stat_processed"), j.processed if j is not None else 0, "info"),
        (t("stat_errors"), errors, "bad" if errors else "ok"),
    ])


def job_state(j: ScrapeJob | None) -> str:
    if j is None:
        return "idle"
    if j.running:
        return "running"
    if j.error:
        return "failed"
    if j.stop_event.is_set():
        return "stopped"
    return "finished"


def reset_everything() -> None:
    """Forget the finished/stopped job and the uploaded file; keep Chromium bootstrap."""
    j = current_job()
    if j is not None and j.running:
        return
    set_job(None)
    kill_chromium_processes()
    ss = st.session_state
    ss.input_df = None
    ss.input_name = None
    ss.uploader_key = ss.get("uploader_key", 0) + 1
    ss.flash = "reset_done"


# --------------------------------------------------------------------------- #
# Streamlit UI
# --------------------------------------------------------------------------- #
LANG = _resolve_lang()
RTL = LANG == "ar"
st.set_page_config(page_title=t("page_title"), page_icon="🇰🇼", layout="wide")
inject_css(RTL)

ss = st.session_state
ss.setdefault("input_df", None)
ss.setdefault("input_name", None)
ss.setdefault("uploader_key", 0)
ss.setdefault("flash", None)

# ---- header + language switcher ------------------------------------------ #
h_main, h_lang = st.columns([5, 1.2])
with h_lang:
    choice = st.segmented_control(
        t("lang"), options=list(LANGS), format_func=lambda k: LANGS[k],
        default=LANG, key="lang_ctl", label_visibility="collapsed", width="stretch",
    )
    if choice and choice != LANG:
        ss.lang = choice
        st.query_params["lang"] = choice
        st.rerun()
with h_main:
    hero(t("app_title"), t("app_subtitle"), t("hero_badge"))

# ---- one-time browser bootstrap ------------------------------------------- #
try:
    with st.spinner(t("boot_spinner")):
        boot_msg = ensure_chromium()
except Exception as exc:  # noqa: BLE001
    st.error(t("boot_error"))
    st.code(str(exc))
    st.stop()

job = current_job()
running = job is not None and job.running

# ---- sidebar -------------------------------------------------------------- #
with st.sidebar:
    st.header("⚙️ " + t("settings"))
    if ON_CLOUD:
        st.caption("🖥️ " + t("cloud_note"))
        headless = True
    else:
        headless = st.toggle(t("headless"), value=True)
    timeout_s = st.slider(t("timeout"), 10, 120, 45)
    delay_s = st.slider(t("delay"), 0.0, 10.0, 2.0, 0.5)
    retries = st.slider(t("retries"), 0, 3, 1)
    restart_every = st.slider(t("restart_every"), 10, 200, 20 if ON_CLOUD else 40, 10)
    skip_filled = st.checkbox(t("skip_filled"), value=True, help=t("skip_filled_help"))
    st.markdown("---")
    st.markdown(f"**{t('services')}**")
    do_moi = st.checkbox(t("svc_moi"), value=True)
    do_status = st.checkbox(t("svc_status"), value=True)
    do_renew = st.checkbox(t("svc_renew"), value=True)
    with st.expander("🧭 " + t("sysinfo")):
        st.code(f"python {platform.python_version()} / {platform.system()}\n{boot_msg.splitlines()[0]}")
        if job is not None:
            st.code(
                f"{t('started')}: {job.started_at:%H:%M:%S}\n"
                f"{t('relaunches')}: {job.scraper.relaunches if job.scraper else '-'}\n"
                f"{t('watchdog')}: {job.watchdog_kills}\n"
                f"{t('cooldowns')}: {job.cooldowns}"
            )
        st.caption(t("boot_log"))
        st.code(boot_msg, language="text")
    with st.expander("🔬 " + t("quick_test")):
        st.caption(t("quick_test_help"))
        test_cid = normalize_civil_id(st.text_input(t("test_cid"), value=""))
        if st.button("▶ " + t("run_test"), disabled=running or len(test_cid) != 12):
            with st.spinner(t("testing")):
                try:
                    report = _in_thread(lambda: run_quick_test(test_cid, timeout_s))
                except Exception as exc:  # noqa: BLE001
                    report = f"{type(exc).__name__}: {exc}"
            st.code(report, language="text")

# ---- flash message -------------------------------------------------------- #
if ss.flash:
    st.success(t(ss.flash))
    ss.flash = None

if running:
    st.info("⏳ " + t("running_banner", time=f"{job.started_at:%H:%M:%S}"))

# ---- upload card ---------------------------------------------------------- #
with st.container(border=True):
    card_title("📂", t("upload_title"))
    st.caption(t("upload_help", cols=" ، ".join(f"`{c}`" for c in REQUIRED_COLS)))
    uploaded = st.file_uploader(t("upload_label"), type=["xlsx", "xlsm", "xls"],
                                key=f"uploader_{ss.uploader_key}", label_visibility="collapsed")

    if uploaded is not None and uploaded.name != ss.input_name:
        try:
            df_in = pd.read_excel(uploaded, dtype={COL_CID: str})
        except Exception as exc:  # noqa: BLE001
            st.error(t("read_error", err=exc))
            st.stop()
        df_in.columns = [str(c).strip() for c in df_in.columns]
        missing = [c for c in REQUIRED_COLS if c not in df_in.columns]
        if missing:
            st.error(t("missing_cols", cols=" ، ".join(missing)))
            st.write(t("existing_cols"), list(df_in.columns))
            st.stop()
        for c in (COL_MOI, COL_RENEW, COL_STATUS):
            df_in[c] = df_in[c].astype("object")
        df_in[COL_CID] = df_in[COL_CID].apply(normalize_civil_id)
        ss.input_df, ss.input_name = df_in, uploaded.name
    if ss.input_df is not None:
        st.caption("✅ " + t("file_loaded", name=ss.input_name))

# ---- preview + actions card ---------------------------------------------- #
start_clicked = stop_clicked = reset_clicked = False
if ss.input_df is not None or job is not None:
    with st.container(border=True):
        card_title("📋", t("preview_title"), status_pill(job_state(job)))
        if job is None and ss.input_df is not None:
            # Live tiles move into the auto-refreshing progress card once a job exists.
            job_tiles(ss.input_df, None)
            st.dataframe(ss.input_df, width="stretch", height=240)

        b1, b2, b3 = st.columns(3)
        can_start = ss.input_df is not None and not running
        start_clicked = b1.button("🚀 " + t("start"), type="primary", width="stretch", disabled=not can_start)
        stop_clicked = b2.button("⏹ " + t("stop"), width="stretch", disabled=not running)
        reset_clicked = b3.button("🧹 " + t("reset"), width="stretch", disabled=running or (job is None and ss.input_df is None),
                                  help=t("reset_help"))
else:
    st.caption("💡 " + t("empty_hint"))

if reset_clicked:
    reset_everything()
    st.rerun()

if stop_clicked and job is not None:
    job.stop_event.set()

if start_clicked and not running:
    job = ScrapeJob(
        ss.input_df.copy(),
        dict(headless=headless, timeout_s=timeout_s, delay_s=delay_s, retries=retries,
             restart_every=restart_every, skip_filled=skip_filled,
             do_moi=do_moi, do_status=do_status, do_renew=do_renew),
    )
    set_job(job)
    job.start()
    st.rerun()  # redraw the whole page with the job present (status pill, buttons)

if job is None:
    st.stop()


# ---- live progress ------------------------------------------------------- #
PREVIEW_ROWS = 40


def render_job(j: ScrapeJob) -> None:
    card_title("📈", t("progress_title"), status_pill(job_state(j)))
    job_tiles(j.df, j)
    pct = j.processed / j.total if j.total else 1.0
    st.progress(min(pct, 1.0), text=t("progress_text", done=j.processed, total=j.total, cid=j.current))
    if j.running:
        elapsed = (datetime.now() - j.started_at).total_seconds()
        per_row = elapsed / j.processed if j.processed else 0.0
        eta = per_row * (j.total - j.processed) / 60
        st.info("🔎 " + t("eta_text", cid=j.current, per_row=per_row, eta=eta))
    elif j.error:
        st.error(t("err_stopped", err=j.error))
    elif j.stop_event.is_set():
        st.warning(t("stopped_warning"))
    else:
        st.success(t("finished_ok"))

    snap = j.snapshot()
    if j.running and j.total > PREVIEW_ROWS:
        # Only ship a window around the current row every tick: sending all
        # 735 rows over the websocket every second is what made the UI crawl.
        lo = max(0, j.processed - PREVIEW_ROWS + 5)
        st.caption(t("preview_window", lo=lo + 1, hi=min(lo + PREVIEW_ROWS, j.total)))
        st.dataframe(snap.iloc[lo:lo + PREVIEW_ROWS], width="stretch", height=320)
    else:
        st.dataframe(snap, width="stretch", height=320)

    with st.expander("🧾 " + t("log_title"), expanded=False):
        with j.lock:
            text = "\n".join(j.logs[-200:])
        st.code(text or "…", language="text")

    st.download_button(
        "📥 " + (t("download_partial") if j.running else t("download_final")),
        data=j.excel_bytes(),
        file_name=f"results_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        width="stretch",
        key=f"dl_{j.processed}_{j.running}",
    )


with st.container(border=True):
    if hasattr(st, "fragment"):
        # Non-blocking: the script run finishes immediately and only this block
        # re-executes every 2 s. Streamlit stays responsive (stop button,
        # reconnects) and nothing is held open for the hours a 735-row run takes.
        was_running = running

        @st.fragment(run_every=2.0 if was_running else None)
        def _progress():
            j = current_job()
            if j is None:
                return
            render_job(j)
            if was_running and not j.running:
                st.rerun(scope="app")  # switch to the final (static) view

        _progress()
    else:  # very old Streamlit: blocking poll, but cheap
        box = st.empty()
        while job.running:
            with box.container():
                render_job(job)
            time.sleep(2.0)
        with box.container():
            render_job(job)
