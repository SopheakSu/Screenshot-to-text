"""
OCR core for ScreenToExcel.
- preprocess(): cleans a screenshot of a coloured grid (black / magenta cells,
  white / red text, grey grid lines) into black text on white.
- run_ocr(): returns a list of words with boxes, using Windows built-in OCR
  (no extra install) or Tesseract as a fallback.
- parse_rows(): groups words into table rows and extracts the report fields.
"""
import io
import re
import statistics
from dataclasses import dataclass, field

import numpy as np
from PIL import Image

SCALE = 3  # upscale factor before OCR (small screen fonts read much better)
PAD = 20   # white border added around the image (in upscaled pixels)


# --------------------------------------------------------------------------
# Image preprocessing
# --------------------------------------------------------------------------
def preprocess(img: Image.Image, scale: int = SCALE) -> Image.Image:
    """Return a clean black-on-white image, upscaled."""
    img = img.convert("RGB")
    a = np.asarray(img).astype(np.int16)
    h, w, _ = a.shape

    # 1. Find background colours: the most frequent (quantised) colours.
    q = (a // 32).reshape(-1, 3)
    keys = q[:, 0] * 64 + q[:, 1] * 8 + q[:, 2]
    counts = np.bincount(keys, minlength=512)
    order = np.argsort(counts)[::-1]
    total = keys.size
    bgs = []
    for k in order[:6]:
        if counts[k] / total < 0.06:
            break
        mask = keys == k
        bgs.append(a.reshape(-1, 3)[mask].mean(axis=0))
    if not bgs:
        bgs = [np.array([0, 0, 0])]

    # 2. Text-ness = distance from the nearest background colour.
    flat = a.reshape(-1, 1, 3)
    bg = np.array(bgs).reshape(1, -1, 3)
    dist = np.sqrt(((flat - bg) ** 2).sum(axis=2)).min(axis=1).reshape(h, w)
    ink = np.clip(dist * 2.2, 0, 255)  # 0 = background, 255 = text

    # 3. Remove grid lines: only long *unbroken* straight runs are lines.
    #    (Stacked characters like "1" in every row have gaps, so they survive.)
    on = ink > 90
    for axis, length in ((0, h), (1, w)):
        m = on if axis == 0 else on.T
        best = np.zeros(m.shape[1], dtype=np.int32)
        run = np.zeros(m.shape[1], dtype=np.int32)
        for row in m:
            run = np.where(row, run + 1, 0)
            best = np.maximum(best, run)
        # vertical lines span (almost) the whole height; horizontal the width
        lines = best > (max(16, int(length * 0.85)) if axis == 0 else max(16, int(length * 0.6)))
        if axis == 0:
            ink[:, lines] = 0
        else:
            ink[lines, :] = 0
    ink[ink < 60] = 0  # drop faint noise / dotted lines

    out = Image.fromarray((255 - ink).astype(np.uint8), "L")
    if scale != 1:
        out = out.resize((w * scale, h * scale), Image.LANCZOS)
    # white margin so text touching the edge is still read; coords shifted back later
    pad = Image.new("L", (out.width + 2 * PAD, out.height + 2 * PAD), 255)
    pad.paste(out, (PAD, PAD))
    return pad


# --------------------------------------------------------------------------
# OCR engines -> list of Word
# --------------------------------------------------------------------------
@dataclass
class Word:
    text: str
    x: float
    y: float
    w: float
    h: float

    @property
    def cx(self):
        return self.x + self.w / 2

    @property
    def cy(self):
        return self.y + self.h / 2

    @property
    def right(self):
        return self.x + self.w


def _ocr_windows(img: Image.Image):
    """Windows 10/11 built-in OCR (Windows.Media.Ocr)."""
    import asyncio
    from winrt.windows.media.ocr import OcrEngine
    from winrt.windows.globalization import Language
    from winrt.windows.graphics.imaging import BitmapDecoder
    from winrt.windows.storage.streams import InMemoryRandomAccessStream, DataWriter

    max_dim = OcrEngine.max_image_dimension
    if max(img.size) > max_dim:
        r = max_dim / max(img.size)
        img = img.resize((int(img.width * r), int(img.height * r)), Image.LANCZOS)
        back = 1 / r
    else:
        back = 1.0

    buf = io.BytesIO()
    img.convert("RGB").save(buf, format="PNG")
    data = buf.getvalue()

    async def go():
        stream = InMemoryRandomAccessStream()
        writer = DataWriter(stream)
        try:
            writer.write_bytes(data)
        except TypeError:
            writer.write_bytes(list(data))
        await writer.store_async()
        await writer.flush_async()
        writer.detach_stream()
        stream.seek(0)
        decoder = await BitmapDecoder.create_async(stream)
        bmp = await decoder.get_software_bitmap_async()
        engine = None
        for tag in ("en-US", "en-GB", "en"):
            try:
                lang = Language(tag)
                if OcrEngine.is_language_supported(lang):
                    engine = OcrEngine.try_create_from_language(lang)
                    if engine:
                        break
            except Exception:
                pass
        if engine is None:
            engine = OcrEngine.try_create_from_user_profile_languages()
        if engine is None:
            raise RuntimeError(
                "Windows OCR English language pack not found.\n"
                "Settings > Time & Language > Language > add English (United States)."
            )
        return await engine.recognize_async(bmp)

    result = asyncio.run(go())
    words = []
    for line in result.lines:
        for wd in line.words:
            r = wd.bounding_rect
            words.append(Word(wd.text, r.x * back, r.y * back, r.width * back, r.height * back))
    return words


def _ocr_tesseract(img: Image.Image):
    import os
    import pytesseract
    for p in (r"C:\Program Files\Tesseract-OCR\tesseract.exe",
              r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"):
        if os.path.exists(p):
            pytesseract.pytesseract.tesseract_cmd = p
            break
    d = pytesseract.image_to_data(img, lang="eng", config="--psm 6",
                                  output_type=pytesseract.Output.DICT)
    words = []
    for i, t in enumerate(d["text"]):
        t = (t or "").strip()
        if t and float(d["conf"][i]) > -1:
            words.append(Word(t, d["left"][i], d["top"][i], d["width"][i], d["height"][i]))
    return words


def run_ocr(img: Image.Image, engine: str = "auto"):
    """engine: 'auto' | 'windows' | 'tesseract'. Returns (words, engine_used)."""
    clean = preprocess(img)
    errors = []
    order = {"auto": ["windows", "tesseract"], "windows": ["windows"],
             "tesseract": ["tesseract"]}[engine]
    for eng in order:
        try:
            fn = _ocr_windows if eng == "windows" else _ocr_tesseract
            words = fn(clean)
            for wd in words:  # back to original-image coordinates
                wd.x = (wd.x - PAD) / SCALE; wd.y = (wd.y - PAD) / SCALE; wd.w /= SCALE; wd.h /= SCALE
            return words, eng
        except Exception as e:  # try next engine
            errors.append(f"{eng}: {e}")
    raise RuntimeError("No OCR engine worked:\n" + "\n".join(errors))


# --------------------------------------------------------------------------
# Row grouping and field extraction
# --------------------------------------------------------------------------
FIELDS = ["Time", "Code", "Name", "Contact", "Area", "Zone", "Type"]

_CODE = re.compile(r"[0OC][1Il]\s?[-—–~]\s?((?:[0-9OoIlSBGDQC]\s?){5}[0-9OoIlSBGDQC])")
_SEP = r"\s?[-~—–=_]+\s?"
_TCODE = re.compile(r"T[0O][1Il]" + _SEP + r"[0O]*([0-9Il])" + _SEP + r"[0-9Il]?")
_TIME = re.compile(r"(\d{1,2})\s?/\s?(\d{1,2})\s+(\d{2})\s?[:.]?\s?(\d{2})")
_DIGIT_FIX = str.maketrans({"O": "0", "o": "0", "I": "1", "l": "1", "S": "5", "B": "8", "G": "0", "D": "0", "Q": "0", "C": "0"})


@dataclass
class Row:
    words: list
    fields: dict = field(default_factory=dict)
    raw: str = ""


def group_rows(words):
    if not words:
        return []
    hs = [w.h for w in words if w.h > 0]
    med_h = statistics.median(hs) if hs else 10
    words = sorted(words, key=lambda w: w.cy)
    rows, cur = [], [words[0]]
    for w in words[1:]:
        ref = statistics.mean(x.cy for x in cur)
        if abs(w.cy - ref) <= med_h * 0.55:
            cur.append(w)
        else:
            rows.append(cur)
            cur = [w]
    rows.append(cur)
    return [sorted(r, key=lambda w: w.x) for r in rows], med_h


def _row_text(ws, med_h):
    """Join words; insert ' | ' where there is a big horizontal gap (new cell)."""
    parts = []
    for i, w in enumerate(ws):
        if i:
            gap = w.x - ws[i - 1].right
            parts.append(" | " if gap > med_h * 0.9 else " ")
        parts.append(w.text)
    return "".join(parts)


def _clean(s):
    s = s.replace("|", " ")
    s = re.sub(r"\s+", " ", s)
    return s.strip(" -_.,~—–'\"`")


def parse_row_text(text, type_text="Intruder"):
    f = {k: "" for k in FIELDS}
    f["Type"] = type_text
    m_code = _CODE.search(text)
    if not m_code:
        return None
    f["Code"] = "01-" + m_code.group(1).replace(" ", "").translate(_DIGIT_FIX)

    m_time = _TIME.search(text[:m_code.start()].translate(_DIGIT_FIX))
    if m_time:
        mo, d, hh, mm = m_time.groups()
        f["Time"] = f"{int(mo):02d}/{int(d):02d} {int(hh):02d}:{mm}"

    rest = text[m_code.end():]
    # drop the "-____" tail of the code (sometimes clipped as "-00C")
    rest = re.sub(r"^\s?[-—–~]+[_0OoCc.\-—–]{0,4}", "", rest)

    m_t = _TCODE.search(rest)
    if m_t:
        f["Name"] = _clean(rest[:m_t.start()])
        f["Area"] = m_t.group(1).translate(_DIGIT_FIX)
        after = rest[m_t.end():]
        # zone = text up to the next cell boundary / dash column / lone number
        zone = re.split(r"\s\|\s|\s[-–—~_]+(?:\s|$)|\s\d+(?:\s|$)", after, maxsplit=1)[0]
        zone = re.sub(r"[~—–]", "-", zone)
        zone = re.sub(r"^(\d) ([A-Z])", r"\1\2", zone)  # "1 F-Show" -> "1F-Show"
        zone = re.sub(r"([A-Z]) (\d)", r"\1\2", zone)                    # "DP 1-3" -> "DP1-3"
        zone = re.sub(r"(SH|DP|M|F)[IiltT|](?=[-,\s]|$)", r"\g<1>1", zone)  # "SHI" -> "SH1"
        f["Zone"] = _clean(zone).rstrip("/\\|")
    else:
        f["Name"] = _clean(rest.split(" | ")[0])
    return f


def parse_rows(words, type_text="Intruder"):
    grouped = group_rows(words)
    if not grouped:
        return []
    rows, med_h = grouped
    out = []
    for ws in rows:
        text = _row_text(ws, med_h)
        f = parse_row_text(text, type_text)
        if f:
            out.append(Row(ws, f, text))
    return out


def image_to_rows(img, engine="auto", type_text="Intruder"):
    words, used = run_ocr(img, engine)
    return parse_rows(words, type_text), used
