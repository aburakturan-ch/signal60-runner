from __future__ import annotations

import os
import re
import subprocess
import sys
import tempfile
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

def _bundle_root() -> Path:
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parent

def _find_tesseract() -> Tuple[str, Optional[str]]:
    env_exe=os.environ.get("GULSEN_TESSERACT_EXE"); env_data=os.environ.get("GULSEN_TESSDATA")
    if env_exe and Path(env_exe).exists(): return env_exe,env_data
    root=_bundle_root(); candidates=[root/"tesseract"/"tesseract.exe",root/"tesseract"/"tesseract",root.parent/"Resources"/"tesseract"/"tesseract"]
    for exe in candidates:
        if exe.exists():
            data=exe.parent/"tessdata"; return str(exe),str(data) if data.exists() else None
    from shutil import which
    found=which("tesseract")
    if found: return found,env_data
    raise RuntimeError("OCR motoru bulunamadı.")

def _pil_to_bgr(img: Image.Image) -> np.ndarray:
    rgb=np.asarray(img.convert("RGB")); return cv2.cvtColor(rgb,cv2.COLOR_RGB2BGR)

def _bgr_to_pil(arr: np.ndarray) -> Image.Image:
    return Image.fromarray(cv2.cvtColor(arr,cv2.COLOR_BGR2RGB))

def _rotate_bound(image: np.ndarray,angle: float,border_value=255) -> np.ndarray:
    h,w=image.shape[:2]; center=(w/2.0,h/2.0); m=cv2.getRotationMatrix2D(center,angle,1.0); cos=abs(m[0,0]); sin=abs(m[0,1]); nw=int((h*sin)+(w*cos)); nh=int((h*cos)+(w*sin)); m[0,2]+=(nw/2)-center[0]; m[1,2]+=(nh/2)-center[1]; return cv2.warpAffine(image,m,(nw,nh),flags=cv2.INTER_CUBIC,borderMode=cv2.BORDER_CONSTANT,borderValue=border_value)

def deskew_image(img: Image.Image) -> Tuple[Image.Image,float]:
    img=ImageOps.exif_transpose(img).convert("RGB"); bgr=_pil_to_bgr(img); gray=cv2.cvtColor(bgr,cv2.COLOR_BGR2GRAY)
    if max(gray.shape)>2200:
        scale=2200.0/max(gray.shape); gray_small=cv2.resize(gray,None,fx=scale,fy=scale,interpolation=cv2.INTER_AREA)
    else: gray_small=gray
    blur=cv2.GaussianBlur(gray_small,(3,3),0); thresh=cv2.threshold(blur,0,255,cv2.THRESH_BINARY_INV+cv2.THRESH_OTSU)[1]; kernel=cv2.getStructuringElement(cv2.MORPH_RECT,(25,3)); merged=cv2.dilate(thresh,kernel,iterations=1); coords=np.column_stack(np.where(merged>0))
    if len(coords)<120: return img,0.0
    angle=cv2.minAreaRect(coords[:,::-1].astype(np.float32))[-1]
    if angle < -45: angle=90+angle
    correction=-float(angle)
    if abs(correction)<0.35 or abs(correction)>15: return img,0.0
    return _bgr_to_pil(_rotate_bound(bgr,correction,border_value=(255,255,255))),correction

def _normalize_ocr_text(text: str) -> str:
    text=text.replace("\r\n","\n").replace("\r","\n"); lines=[re.sub(r"[ \t]+"," ",line).strip() for line in text.split("\n")]; paras=[]; buf=[]
    for line in lines:
        if not line:
            if buf: paras.append(" ".join(buf).strip()); buf=[]
            continue
        if buf and buf[-1].endswith("-") and line and line[0].islower(): buf[-1]=buf[-1][:-1]+line
        else: buf.append(line)
    if buf: paras.append(" ".join(buf).strip())
    return "\n\n".join(p for p in paras if p)

def ocr_image(image: Image.Image|str|Path,*,language: str="tur+eng",deskew: bool=True) -> OCRResult:
    if isinstance(image,(str,Path)):
        with Image.open(image) as im: src=ImageOps.exif_transpose(im).convert("RGB")
    else: src=ImageOps.exif_transpose(image).convert("RGB")
    rotation=0.0
    if deskew:
        try: src,rotation=deskew_image(src)
        except Exception: rotation=0.0
    if src.width<1300:
        ratio=min(2.0,1300.0/max(1,src.width)); src=src.resize((int(src.width*ratio),int(src.height*ratio)),Image.Resampling.LANCZOS)
    exe,tessdata=_find_tesseract()
    with tempfile.TemporaryDirectory(prefix="gulsen_ocr_") as td:
        inp=Path(td)/"page.png"; outbase=Path(td)/"ocr"; src.save(inp,"PNG"); cmd=[exe,str(inp),str(outbase),"-l",language,"--oem","1","--psm","6","-c","preserve_interword_spaces=1"]
        if tessdata: cmd.extend(["--tessdata-dir",tessdata])
        cp=subprocess.run(cmd,stdout=subprocess.PIPE,stderr=subprocess.PIPE,text=True,encoding="utf-8",errors="replace",timeout=180); txt_path=outbase.with_suffix(".txt")
        if cp.returncode!=0 or not txt_path.exists():
            if language!="eng": return ocr_image(src,language="eng",deskew=False)
            raise RuntimeError((cp.stderr or "OCR çalıştırılamadı.").strip())
        text=txt_path.read_text(encoding="utf-8",errors="replace")
    return OCRResult(_normalize_ocr_text(text),rotation=rotation)
