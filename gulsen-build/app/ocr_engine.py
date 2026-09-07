from __future__ import annotations

import os
import re
import subprocess
import sys
import tempfile
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Optional, Tuple

import cv2
import numpy as np
from PIL import Image, ImageOps


@dataclass
class OCRResult:
    text: str
    rotation: float = 0.0
    engine: str = "tesseract"
    warning: str = ""
    cancelled: bool = False


def _bundle_root() -> Path:
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parent


def _find_tesseract() -> Tuple[str, Optional[str]]:
    env_exe = os.environ.get("GULSEN_TESSERACT_EXE")
    env_data = os.environ.get("GULSEN_TESSDATA")
    if env_exe and Path(env_exe).exists():
        return env_exe, env_data
    root = _bundle_root()
    candidates = [
        root / "tesseract" / "tesseract.exe",
        root / "tesseract" / "tesseract",
        root.parent / "Resources" / "tesseract" / "tesseract",
    ]
    for exe in candidates:
        if exe.exists():
            data = exe.parent / "tessdata"
            return str(exe), str(data) if data.exists() else None
    from shutil import which
    found = which("tesseract")
    if found:
        return found, env_data
    raise RuntimeError("OCR motoru bulunamadı.")


def _cancelled(cancel_event) -> bool:
    try:
        return bool(cancel_event and cancel_event.is_set())
    except Exception:
        return False


def _progress(callback, value: float, message: str = ""):
    if callback is None:
        return
    try:
        callback(max(0.0, min(100.0, float(value))), message)
    except Exception:
        pass


def deskew_image(img: Image.Image) -> Tuple[Image.Image, float]:
    img = ImageOps.exif_transpose(img).convert("RGB")
    analysis = img.copy()
    analysis.thumbnail((2200, 2200), Image.Resampling.LANCZOS)
    arr = np.asarray(analysis)
    gray = cv2.cvtColor(arr, cv2.COLOR_RGB2GRAY)
    blur = cv2.GaussianBlur(gray, (3, 3), 0)
    thresh = cv2.threshold(blur, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (25, 3))
    merged = cv2.dilate(thresh, kernel, iterations=1)
    coords = np.column_stack(np.where(merged > 0))
    if len(coords) < 120:
        return img, 0.0
    angle = cv2.minAreaRect(coords[:, ::-1].astype(np.float32))[-1]
    if angle < -45:
        angle = 90 + angle
    correction = -float(angle)
    if abs(correction) < 0.35 or abs(correction) > 15:
        return img, 0.0
    return img.rotate(correction, resample=Image.Resampling.BICUBIC, expand=True, fillcolor="white"), correction


def _normalize_ocr_text(text: str) -> str:
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    lines = [re.sub(r"[ \t]+", " ", line).strip() for line in text.split("\n")]
    paras, buf = [], []
    for line in lines:
        if not line:
            if buf:
                paras.append(" ".join(buf).strip())
                buf = []
            continue
        if buf and buf[-1].endswith("-") and line and line[0].islower():
            buf[-1] = buf[-1][:-1] + line
        else:
            buf.append(line)
    if buf:
        paras.append(" ".join(buf).strip())
    return "\n\n".join(p for p in paras if p)


def _run_tesseract(cmd, *, cancel_event=None, timeout=90, progress_callback=None):
    creationflags = 0
    if os.name == "nt":
        creationflags |= getattr(subprocess, "CREATE_NO_WINDOW", 0)
    proc = subprocess.Popen(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding="utf-8",
        errors="replace",
        creationflags=creationflags,
    )
    started = time.monotonic()
    while proc.poll() is None:
        elapsed = time.monotonic() - started
        _progress(progress_callback, 35.0 + min(45.0, (elapsed / max(8.0, timeout * 0.55)) * 45.0), "OCR okunuyor…")
        if _cancelled(cancel_event):
            proc.kill()
            try:
                proc.communicate(timeout=2)
            except Exception:
                pass
            return None, "OCR iptal edildi."
        if elapsed > timeout:
            proc.kill()
            try:
                proc.communicate(timeout=2)
            except Exception:
                pass
            raise TimeoutError(f"OCR {timeout} saniye içinde tamamlanamadı.")
        time.sleep(0.12)
    stdout, stderr = proc.communicate()
    return proc.returncode, stderr or stdout or ""


def ocr_image(
    image: Image.Image | str | Path,
    *,
    language: str = "tur+eng",
    deskew: bool = True,
    cancel_event=None,
    timeout: int = 90,
    progress_callback=None,
) -> OCRResult:
    _progress(progress_callback, 2, "Görsel hazırlanıyor…")
    if _cancelled(cancel_event):
        return OCRResult("", cancelled=True, warning="OCR iptal edildi.")

    if isinstance(image, (str, Path)):
        with Image.open(image) as im:
            src = ImageOps.exif_transpose(im).convert("RGB")
    else:
        src = ImageOps.exif_transpose(image).convert("RGB")
    _progress(progress_callback, 10, "Görsel yüklendi")

    rotation = 0.0
    if deskew and not _cancelled(cancel_event):
        try:
            src, rotation = deskew_image(src)
        except Exception:
            rotation = 0.0
    _progress(progress_callback, 23, "Tarama açısı düzeltildi")

    if _cancelled(cancel_event):
        return OCRResult("", rotation=rotation, cancelled=True, warning="OCR iptal edildi.")

    max_side = max(src.size)
    if max_side > 3400:
        ratio = 3400.0 / max_side
        src = src.resize((max(1, int(src.width * ratio)), max(1, int(src.height * ratio))), Image.Resampling.LANCZOS)
    elif src.width < 1300:
        ratio = min(2.0, 1300.0 / max(1, src.width))
        src = src.resize((int(src.width * ratio), int(src.height * ratio)), Image.Resampling.LANCZOS)
    _progress(progress_callback, 30, "OCR hazırlanıyor…")

    exe, tessdata = _find_tesseract()
    with tempfile.TemporaryDirectory(prefix="gulsen_ocr_") as td:
        inp = Path(td) / "page.png"
        outbase = Path(td) / "ocr"
        src.save(inp, "PNG")
        cmd = [exe, str(inp), str(outbase), "-l", language, "--oem", "1", "--psm", "6", "-c", "preserve_interword_spaces=1"]
        if tessdata:
            cmd.extend(["--tessdata-dir", tessdata])

        _progress(progress_callback, 35, "OCR başladı")
        code, diagnostic = _run_tesseract(cmd, cancel_event=cancel_event, timeout=timeout, progress_callback=progress_callback)
        if code is None:
            return OCRResult("", rotation=rotation, cancelled=True, warning=diagnostic)
        _progress(progress_callback, 86, "OCR tamamlandı, metin düzenleniyor…")

        txt_path = outbase.with_suffix(".txt")
        if code != 0 or not txt_path.exists():
            if language != "eng" and not _cancelled(cancel_event):
                return ocr_image(src, language="eng", deskew=False, cancel_event=cancel_event, timeout=timeout, progress_callback=progress_callback)
            raise RuntimeError((diagnostic or "OCR çalıştırılamadı.").strip())
        text = txt_path.read_text(encoding="utf-8", errors="replace")

    _progress(progress_callback, 100, "OCR tamamlandı")
    return OCRResult(_normalize_ocr_text(text), rotation=rotation)
