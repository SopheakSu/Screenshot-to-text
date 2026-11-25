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
def trim_edge_strips(img: Image.Image) -> Image.Image:
    """Cut off thin strips at the left/right edge whose background colour differs
    from the main area (a sliver of the neighbouring column caught in the crop)."""
    a = np.asarray(img.convert("RGB")).astype(np.int32) // 48
    key = a[:, :, 0] * 36 + a[:, :, 1] * 6 + a[:, :, 2]
    col_mode = np.array([np.bincount(key[:, x]).argmax() for x in range(key.shape[1])])
    main = np.bincount(col_mode).argmax()
    w = len(col_mode)
    limit = max(4, int(w * 0.08))
    left = 0
    while left < w and col_mode[left] != main:
        left += 1
    right = w
    while right > 0 and col_mode[right - 1] != main:
        right -= 1
    left = left if left <= limit else 0
    right = right if w - right <= limit else w
    if right - left > w * 0.5 and (left or right < w):
        return img.crop((left, 0, right, img.height))
    return img


def preprocess(img: Image.Image, scale: int = SCALE, mode: str = "solid") -> Image.Image:
    """Return a clean black-on-white image, upscaled.
    mode: "solid" (backgrounds = large solid colour areas), "frequent"
    (backgrounds = colours covering >6%), "main" (only the most common colour)."""
    img = trim_edge_strips(img.convert("RGB"))
    a = np.asarray(img).astype(np.int16)
    h, w, _ = a.shape

    # 1. Find background colours: frequent colours that form large SOLID areas.
    #    (Text is thin: almost none of its pixels are surrounded by the same colour.)
    q = (a // 32).reshape(-1, 3)
    keys = q[:, 0] * 64 + q[:, 1] * 8 + q[:, 2]
    counts = np.bincount(keys, minlength=512)
    order = np.argsort(counts)[::-1]
    total = keys.size
    k2d = keys.reshape(h, w)
    r = 2 if min(h, w) > 12 else 1
    bgs = []
    for k in order[:8]:
        frac = counts[k] / total
        if mode == "main" and bgs:
            break
        if frac < (0.06 if mode == "frequent" else 0.03):
            break
        if mode != "solid":
            bgs.append(a.reshape(-1, 3)[keys == k].mean(axis=0))
            continue
        m = k2d == k
        solid = m.copy()
        for dy in range(-r, r + 1):
            for dx in range(-r, r + 1):
                solid &= np.roll(np.roll(m, dy, axis=0), dx, axis=1)
        if solid.sum() / max(m.sum(), 1) > 0.25:   # real background
            bgs.append(a.reshape(-1, 3)[keys == k].mean(axis=0))
    if not bgs:
        bgs = [a.reshape(-1, 3)[keys == order[0]].mean(axis=0)]

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
    conf: float = 50.0

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


def find_tesseract():
    """Bundled copy (inside the exe) first, then a normal installation."""
    import os, shutil, sys
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    for p in (os.path.join(base, "tesseract", "tesseract.exe"),
              r"C:\Program Files\Tesseract-OCR\tesseract.exe",
              r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"):
        if os.path.exists(p):
            return p
    return shutil.which("tesseract")


def _ocr_tesseract(img: Image.Image):
    import os, subprocess, sys, tempfile
    exe = find_tesseract()
    if not exe:
        raise RuntimeError("Tesseract not found")
    env = dict(os.environ)
    tessdata = os.path.join(os.path.dirname(exe), "tessdata")
    if os.path.isdir(tessdata):
        env["TESSDATA_PREFIX"] = tessdata
    fd, path = tempfile.mkstemp(suffix=".png")
    os.close(fd)
    try:
        img.save(path)
        flags = 0x08000000 if sys.platform == "win32" else 0  # CREATE_NO_WINDOW
        r = subprocess.run([exe, path, "stdout", "-l", "eng", "--psm", "6", "tsv"],
                           capture_output=True, stdin=subprocess.DEVNULL,
                           env=env, creationflags=flags, timeout=60)
    finally:
        try:
            os.remove(path)
        except OSError:
            pass
    if r.returncode != 0:
        raise RuntimeError(r.stderr.decode("utf-8", "replace")[-500:])
    words = []
    lines = r.stdout.decode("utf-8", "replace").splitlines()
    for line in lines[1:]:
        p = line.split("\t")
        if len(p) >= 12 and p[11].strip():
            words.append(Word(p[11].strip(), float(p[6]), float(p[7]), float(p[8]), float(p[9]), float(p[10])))
    return words


def _score(words):
    """How much confident text was found (longer + more confident = better)."""
    return sum(len(re.sub(r"[^A-Za-z0-9]", "", w.text)) * max(w.conf, 0) / 100 for w in words)


def run_ocr(img: Image.Image, engine: str = "auto"):
    """engine: 'auto' | 'tesseract' | 'windows'. Returns (words, engine_used).
    Tries several clean-up methods and keeps the most confident result."""
    order = {"auto": ["tesseract", "windows"], "windows": ["windows"],
             "tesseract": ["tesseract"]}[engine]
    errors = []
    for eng in order:
        fn = _ocr_windows if eng == "windows" else _ocr_tesseract
        best, best_score = None, -1
        for mode in ("frequent", "solid", "main"):
            try:
                words = fn(preprocess(img, mode=mode))
            except Exception as e:
                errors.append(f"{eng}/{mode}: {e}")
                break  # engine itself is broken -> next engine
            sc = _score(words) if eng == "tesseract" else len("".join(w.text for w in words))
            if sc > best_score:
                best, best_score = words, sc
        if best:
            for wd in best:  # back to original-image coordinates
                wd.x = (wd.x - PAD) / SCALE; wd.y = (wd.y - PAD) / SCALE
                wd.w /= SCALE; wd.h /= SCALE
            return best, eng
        if best is not None:
            errors.append(f"{eng}: no text found")
    raise RuntimeError("No text found / OCR failed:\n" + "\n".join(errors))


# --------------------------------------------------------------------------
# Rows / cells -> text exactly as in the picture
# --------------------------------------------------------------------------
def group_rows(words):
    """Group words into lines by vertical position. Returns (rows, median_height)."""
    if not words:
        return [], 10
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


def _fix(text):
    """Small, safe fixes for typical OCR mistakes on screen fonts."""
    t = text.replace("€", "E").replace("—", "-").replace("–", "-").replace("~", "-")
    t = re.sub(r"!(?=\d)", "", t)                 # "SH!1" -> "SH1"
    t = re.sub(r"(?<=[A-Za-z0-9])!", "1", t)      # "Beam!-3" -> "Beam1-3"
    t = re.sub(r"(?<=[A-Z])[|](?=[-,\s]|$)", "1", t)
    t = re.sub(r"\b([A-Z]{1,3}) (\d+-)", r"\1\2", t)   # "DP 1-3" -> "DP1-3"
    t = re.sub(r"\b(SH|DP|M)[Il](?=[-,\s]|$)", r"\g<1>1", t)  # "SHI" -> "SH1"
    t = re.sub(r"Bea(?:rn|r)(?=[\dlIs]|\b)", "Beam", t)             # "Bearn" -> "Beam"
    t = re.sub(r"Beam[lI|](?=[-,\s\d]|$)", "Beam1", t)              # "Beaml" -> "Beam1"
    t = re.sub(r"Beam[sSéB](?=[,\s]|$)", "Beam8", t)                # "Beams" -> "Beam8"
    return t.strip(" _|.'`\"")


def image_to_table(img, engine="auto", split_cells=True):
    """Returns (rows, engine_used). rows = list of rows, each a list of cell strings.
    Cells are split where there is a wide horizontal gap (a new column)."""
    words, used = run_ocr(img, engine)
    rows, med_h = group_rows(words)
    out = []
    for ws in rows:
        cells, cur = [], [ws[0].text]
        for prev, w in zip(ws, ws[1:]):
            gap = w.x - prev.right
            if split_cells and gap > med_h * 1.6:
                cells.append(" ".join(cur))
                cur = [w.text]
            else:
                cur.append(w.text)
        cells.append(" ".join(cur))
        cells = [_fix(c) for c in cells]
        cells = [c for c in cells if re.search(r"[A-Za-z0-9].*[A-Za-z0-9]|[A-Za-z0-9]{1}", c)]
        # drop lines that are only noise (e.g. bits of grid lines)
        if cells and sum(len(re.sub(r"[^A-Za-z0-9]", "", c)) for c in cells) >= 2:
            out.append(cells)
    return out, used
