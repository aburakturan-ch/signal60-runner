from __future__ import annotations

import hashlib
import json
import os
import platform
import re
import shlex
import shutil
import subprocess
import sys
import tempfile
import urllib.request
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

REPO = "aburakturan-ch/signal60-runner"
RELEASES_API = f"https://api.github.com/repos/{REPO}/releases?per_page=30"
TAG_PREFIX = "gulsen-v"


@dataclass
class UpdateInfo:
    version: str
    tag: str
    asset_name: str
    download_url: str
    sha_asset_url: Optional[str] = None
    notes: str = ""


def _version_tuple(v: str):
    nums = re.findall(r"\d+", v)
    return tuple(int(x) for x in nums[:4]) + (0,) * max(0, 4 - len(nums[:4]))


def _get_json(url: str):
    req = urllib.request.Request(
        url,
        headers={"User-Agent": "GulsenWordTools-Updater", "Accept": "application/vnd.github+json"},
    )
    with urllib.request.urlopen(req, timeout=12) as r:
        return json.loads(r.read().decode("utf-8"))


def check_for_update(current_version: str) -> Optional[UpdateInfo]:
    releases = _get_json(RELEASES_API)
    best = None
    for rel in releases:
        tag = str(rel.get("tag_name", ""))
        if not tag.startswith(TAG_PREFIX) or rel.get("draft"):
            continue
        version = tag[len(TAG_PREFIX):]
        if _version_tuple(version) <= _version_tuple(current_version):
            continue
        if best is None or _version_tuple(version) > _version_tuple(best[0]):
            best = (version, rel)
    if not best:
        return None

    version, rel = best
    assets = rel.get("assets") or []
    wanted = None
    sysname = platform.system()
    machine = platform.machine().lower()

    if sysname == "Windows":
        for a in assets:
            n = a.get("name", "")
            if n.endswith(".exe") and "Windows10_11_x64" in n and "Compatibility" not in n:
                wanted = a
                break
    elif sysname == "Darwin" and machine in {"arm64", "aarch64"}:
        for a in assets:
            n = a.get("name", "")
            if n.endswith(".zip") and "Mac_AppleSilicon" in n:
                wanted = a
                break

    if not wanted:
        return None

    sha_url = None
    for a in assets:
        if a.get("name") == "SHA256SUMS.txt":
            sha_url = a.get("browser_download_url")
            break

    return UpdateInfo(
        version=version,
        tag=rel.get("tag_name", ""),
        asset_name=wanted.get("name", ""),
        download_url=wanted.get("browser_download_url", ""),
        sha_asset_url=sha_url,
        notes=rel.get("body") or "",
    )


def _download(url: str, dest: Path):
    req = urllib.request.Request(url, headers={"User-Agent": "GulsenWordTools-Updater"})
    with urllib.request.urlopen(req, timeout=90) as r, dest.open("wb") as f:
        shutil.copyfileobj(r, f, length=1024 * 1024)


def _sha256(path: Path):
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest().lower()


def _verify_if_available(path: Path, info: UpdateInfo):
    if not info.sha_asset_url:
        return
    req = urllib.request.Request(info.sha_asset_url, headers={"User-Agent": "GulsenWordTools-Updater"})
    with urllib.request.urlopen(req, timeout=30) as r:
        txt = r.read().decode("utf-8", errors="replace")
    expected = None
    for line in txt.splitlines():
        if info.asset_name in line:
            m = re.search(r"\b([0-9a-fA-F]{64})\b", line)
            if m:
                expected = m.group(1).lower()
                break
    if expected and _sha256(path) != expected:
        raise RuntimeError("Güncelleme dosyasının SHA-256 doğrulaması başarısız oldu.")


def _current_app_bundle() -> Optional[Path]:
    p = Path(sys.executable).resolve()
    for parent in [p] + list(p.parents):
        if parent.suffix == ".app":
            return parent
    return None


def download_and_stage(info: UpdateInfo) -> Path:
    td = Path(tempfile.mkdtemp(prefix="gulsen_update_"))
    asset = td / info.asset_name
    _download(info.download_url, asset)
    _verify_if_available(asset, info)
    return asset


def install_staged(asset: Path) -> None:
    if not getattr(sys, "frozen", False):
        raise RuntimeError("Otomatik kurulum yalnızca standalone uygulamada çalışır.")
    if platform.system() == "Windows":
        _install_windows(asset)
    elif platform.system() == "Darwin":
        _install_macos(asset)
    else:
        raise RuntimeError("Bu işletim sistemi için otomatik güncelleme desteklenmiyor.")


def _install_windows(asset: Path):
    current = Path(sys.executable).resolve()
    if asset.suffix.lower() != ".exe":
        raise RuntimeError("Windows güncelleme paketi EXE değil.")

    td = asset.parent
    script = td / "gulsen_update.cmd"
    script.write_text(
        "@echo off\r\n"
        "chcp 65001 >nul\r\n"
        "timeout /t 2 /nobreak >nul\r\n"
        "set TRY=0\r\n"
        ":RETRY\r\n"
        f'copy /Y "{asset}" "{current}" >nul 2>nul\r\n'
        "if %errorlevel%==0 goto STARTAPP\r\n"
        "set /a TRY+=1\r\n"
        "if %TRY% GEQ 12 goto FAIL\r\n"
        "timeout /t 1 /nobreak >nul\r\n"
        "goto RETRY\r\n"
        ":STARTAPP\r\n"
        f'start "" "{current}"\r\n'
        f'del /Q "{asset}" >nul 2>nul\r\n'
        'del /Q "%~f0" >nul 2>nul\r\n'
        "exit /b 0\r\n"
        ":FAIL\r\n"
        f'start "" "{current}"\r\n'
        "exit /b 1\r\n",
        encoding="utf-8",
    )
    flags = getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0) | getattr(subprocess, "DETACHED_PROCESS", 0)
    subprocess.Popen(["cmd.exe", "/c", str(script)], creationflags=flags, close_fds=True)


def _install_macos(asset: Path):
    current = _current_app_bundle()
    if not current:
        raise RuntimeError("Mevcut .app paketi bulunamadı.")
    if asset.suffix.lower() != ".zip":
        raise RuntimeError("Mac güncelleme paketi ZIP değil.")

    td = asset.parent
    extract = td / "new"
    extract.mkdir(exist_ok=True)
    with zipfile.ZipFile(asset, "r") as z:
        z.extractall(extract)
    apps = list(extract.rglob("*.app"))
    if not apps:
        raise RuntimeError("Güncelleme ZIP'i içinde .app bulunamadı.")
    new_app = apps[0]

    backup = current.with_name(current.name + ".old")
    script = td / "gulsen_update.sh"
    script.write_text(
        "#!/bin/sh\n"
        "sleep 2\n"
        f'rm -rf {shlex.quote(str(backup))}\n'
        f'mv {shlex.quote(str(current))} {shlex.quote(str(backup))} || exit 1\n'
        f'if ditto {shlex.quote(str(new_app))} {shlex.quote(str(current))}; then\n'
        f'  xattr -dr com.apple.quarantine {shlex.quote(str(current))} 2>/dev/null || true\n'
        f'  open {shlex.quote(str(current))}\n'
        f'  rm -rf {shlex.quote(str(backup))}\n'
        "else\n"
        f'  rm -rf {shlex.quote(str(current))}\n'
        f'  mv {shlex.quote(str(backup))} {shlex.quote(str(current))}\n'
        f'  open {shlex.quote(str(current))}\n'
        "fi\n"
        f'rm -rf {shlex.quote(str(td))}\n',
        encoding="utf-8",
    )
    script.chmod(0o755)

    if os.access(current.parent, os.W_OK):
        subprocess.Popen(
            ["/bin/sh", str(script)],
            start_new_session=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    else:
        apple = td / "gulsen_update.applescript"
        cmd = f"/bin/sh {shlex.quote(str(script))}"
        safe = cmd.replace("\\", "\\\\").replace('"', '\\"')
        apple.write_text(f'do shell script "{safe}" with administrator privileges\n', encoding="utf-8")
        subprocess.Popen(
            ["/usr/bin/osascript", str(apple)],
            start_new_session=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
