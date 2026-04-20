#!/usr/bin/env python3
"""
RainCo Quote Formatter v2
=========================
Run: python3 RainCo_Quote_Formatter.py
Opens automatically in your browser at http://localhost:8742
"""

import sys, os, io, json, threading, webbrowser, re, traceback, tempfile, urllib.request, urllib.parse
from http.server import HTTPServer, BaseHTTPRequestHandler

PORT = 8742

# ── Auto-install dependencies ─────────────────────────────────────────────────
def ensure(pkg, import_as=None):
    import importlib, subprocess
    try:
        return importlib.import_module(import_as or pkg)
    except ImportError:
        print(f"  Installing {pkg}…", flush=True)
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', pkg,
                               '--quiet', '--break-system-packages'], stderr=subprocess.DEVNULL)
        return importlib.import_module(import_as or pkg)

print("Checking dependencies…", flush=True)
ensure('pdfplumber')
ensure('reportlab')
print("Ready.\n", flush=True)

import pdfplumber
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.utils import ImageReader
from reportlab.platypus import (BaseDocTemplate, PageTemplate, Frame,
                                 SimpleDocTemplate, Table, TableStyle,
                                 Paragraph, Spacer, HRFlowable, KeepTogether,
                                 NextPageTemplate, PageBreak)
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_RIGHT, TA_CENTER, TA_LEFT
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# ── Brand colours — matching RainCo "Dark Slate Green" palette ────────────────
TEAL        = colors.HexColor('#1c3b35')   # primary brand dark forest green
TEAL_L      = colors.HexColor('#eaf1ef')   # light green wash
TEAL_MID    = colors.HexColor('#254f47')   # mid green
WARM_BG     = colors.HexColor('#faf8f5')   # warm off-white (page / card bg)
WARM_BORDER = colors.HexColor('#e2ddd8')   # warm light border
LITE_GREY   = WARM_BORDER                  # alias kept for compatibility
MID_GREY    = colors.HexColor('#8a8a8a')
DARK        = colors.HexColor('#1a1a1a')   # near-black body text
WHITE       = colors.white

# ── Brand font setup (Montserrat — downloaded at first run) ───────────────────
_FONT_DIR = os.path.join(tempfile.gettempdir(), 'rainco_fonts')
_FONTS = {
    'RC':     ('Montserrat',         'latin-400-normal.ttf'),
    'RC-Bd':  ('Montserrat-Bold',    'latin-700-normal.ttf'),
    'RC-Sb':  ('Montserrat-SemiBold','latin-600-normal.ttf'),
    'RC-Lt':  ('Montserrat-Light',   'latin-300-normal.ttf'),
}
_FBASE = 'https://cdn.jsdelivr.net/fontsource/fonts/montserrat@latest/'
_fonts_ok = False

def _init_fonts():
    global _fonts_ok
    if _fonts_ok:
        return
    os.makedirs(_FONT_DIR, exist_ok=True)
    hdrs = {'User-Agent': 'Mozilla/5.0'}
    for alias, (face, fname) in _FONTS.items():
        path = os.path.join(_FONT_DIR, fname)
        if not os.path.exists(path):
            try:
                req = urllib.request.Request(_FBASE + fname, headers=hdrs)
                with urllib.request.urlopen(req, timeout=20) as r:
                    open(path, 'wb').write(r.read())
            except Exception as e:
                print(f'  Font download failed ({face}): {e}', flush=True)
                return
        try:
            pdfmetrics.registerFont(TTFont(alias, path))
        except Exception as e:
            print(f'  Font register failed ({face}): {e}', flush=True)
            return
    _fonts_ok = True
    print('  Fonts: Montserrat ✓', flush=True)

def BF():  return 'RC'    if _fonts_ok else 'Helvetica'
def BFB(): return 'RC-Bd' if _fonts_ok else 'Helvetica-Bold'
def BFSB():return 'RC-Sb' if _fonts_ok else 'Helvetica-Bold'
def BFL(): return 'RC-Lt' if _fonts_ok else 'Helvetica'

# ── Brand asset cache (logo + cover image) ────────────────────────────────────
_ASSET_DIR = os.path.join(tempfile.gettempdir(), 'rainco_assets')
_ASSETS = {
    'logo_dark':  'https://rainco.com.au/cdn/shop/files/Dark_Slate_Green_Logo.png?v=1755137128&width=600',
    'logo_white': 'https://rainco.com.au/cdn/shop/files/Whisper_White_Logo.png?v=1755137493&width=600',
}
_asset_paths = {}

def _init_assets():
    os.makedirs(_ASSET_DIR, exist_ok=True)
    # Clean up any stale cached assets that are no longer in _ASSETS
    _expected = {k + ('.png' if 'png' in v.lower() else '.jpg') for k, v in _ASSETS.items()}
    for fname in os.listdir(_ASSET_DIR):
        if fname not in _expected:
            try:
                os.remove(os.path.join(_ASSET_DIR, fname))
                print(f'  Removed stale asset: {fname}', flush=True)
            except Exception:
                pass
    hdrs = {
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36',
        'Referer': 'https://rainco.com.au/'
    }
    for key, url in _ASSETS.items():
        fname = key + ('.png' if 'png' in url.lower() else '.jpg')
        path  = os.path.join(_ASSET_DIR, fname)
        if not os.path.exists(path):
            try:
                req = urllib.request.Request(url, headers=hdrs)
                with urllib.request.urlopen(req, timeout=20) as r:
                    open(path, 'wb').write(r.read())
                print(f'  Asset: {key} ✓', flush=True)
            except Exception as e:
                print(f'  Asset download failed ({key}): {e}', flush=True)
                path = None
        _asset_paths[key] = path if (path and os.path.exists(path)) else None

def S(name, **kw):
    defaults = dict(fontName='Helvetica', fontSize=9, leading=13, textColor=DARK)
    defaults.update(kw)
    return ParagraphStyle(name, **defaults)

def P(text, style=None, **kw):
    st = style or S('_', **kw)
    return Paragraph(str(text), st)

def SP(h=4): return Spacer(1, h*mm)

# ── Image cache ───────────────────────────────────────────────────────────────
_img_cache = {}   # sku → local temp file path or None

def get_product_image(name, sku):
    """Try to fetch product image from rainco.com.au via Shopify APIs."""
    if sku in _img_cache:
        return _img_cache[sku]

    hdrs = {
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',
        'Accept-Language': 'en-AU,en;q=0.9',
    }
    img_url = None

    # ── Method 1: Shopify predictive search by SKU ──
    try:
        q = urllib.parse.quote(sku)
        url = f'https://rainco.com.au/search/suggest.json?q={q}&resources[type]=product&resources[limit]=3'
        req = urllib.request.Request(url, headers=hdrs)
        with urllib.request.urlopen(req, timeout=10) as r:
            data = json.loads(r.read())
            prods = data.get('resources', {}).get('results', {}).get('products', [])
            if prods and prods[0].get('image'):
                img_url = prods[0]['image']
                print(f'    {sku}: image via search ✓', flush=True)
    except Exception as e:
        print(f'    {sku}: search failed ({e})', flush=True)

    # ── Method 2: Shopify product JSON by handle (derived from name) ──
    if not img_url:
        try:
            # Strip finish variant (e.g. " - Antique Brass") for cleaner handle
            base_name = re.sub(r'\s*[-–]\s*\w[\w\s]*$', '', name).strip()
            handle = re.sub(r'[^a-z0-9]+', '-', base_name.lower()).strip('-')
            url = f'https://rainco.com.au/products/{handle}.json'
            req = urllib.request.Request(url, headers=hdrs)
            with urllib.request.urlopen(req, timeout=10) as r:
                data = json.loads(r.read())
                images = data.get('product', {}).get('images', [])
                if images:
                    img_url = images[0].get('src', '')
                    print(f'    {sku}: image via handle "{handle}" ✓', flush=True)
        except Exception:
            # Also try with full name as handle
            try:
                handle = re.sub(r'[^a-z0-9]+', '-', name.lower()).strip('-')
                url = f'https://rainco.com.au/products/{handle}.json'
                req = urllib.request.Request(url, headers=hdrs)
                with urllib.request.urlopen(req, timeout=10) as r:
                    data = json.loads(r.read())
                    images = data.get('product', {}).get('images', [])
                    if images:
                        img_url = images[0].get('src', '')
                        print(f'    {sku}: image via full handle ✓', flush=True)
            except Exception:
                pass

    # ── Method 3: Shopify product search by name ──
    if not img_url:
        try:
            q = urllib.parse.quote(name)
            url = f'https://rainco.com.au/search/suggest.json?q={q}&resources[type]=product&resources[limit]=3'
            req = urllib.request.Request(url, headers=hdrs)
            with urllib.request.urlopen(req, timeout=10) as r:
                data = json.loads(r.read())
                prods = data.get('resources', {}).get('results', {}).get('products', [])
                if prods and prods[0].get('image'):
                    img_url = prods[0]['image']
                    print(f'    {sku}: image via name search ✓', flush=True)
        except Exception:
            pass

    # ── Method 4: Search results HTML page ──
    if not img_url:
        try:
            q = urllib.parse.quote(sku)
            url = f'https://rainco.com.au/search?q={q}&type=product'
            req = urllib.request.Request(url, headers=hdrs)
            with urllib.request.urlopen(req, timeout=12) as r:
                html = r.read().decode('utf-8', errors='ignore')
            m = re.search(r'(https:)?//cdn\.shopify\.com/s/files/[^\s"\'?]+\.(jpg|png|webp)', html)
            if m:
                img_url = ('https:' if m.group(0).startswith('//') else '') + m.group(0)
                print(f'    {sku}: image via HTML search ✓', flush=True)
            else:
                print(f'    {sku}: no image found in HTML search', flush=True)
        except Exception as e:
            print(f'    {sku}: HTML search failed ({e})', flush=True)

    # ── Download image ──
    if img_url:
        try:
            if img_url.startswith('//'):
                img_url = 'https:' + img_url
            # Request a medium size (Shopify supports _400x400 suffix)
            img_url_sized = re.sub(r'\.(jpg|png|webp)(\?.*)?$', r'_400x400.\1', img_url, flags=re.I)
            req = urllib.request.Request(img_url_sized, headers=hdrs)
            with urllib.request.urlopen(req, timeout=12) as r:
                img_data = r.read()
            suffix = '.png' if 'png' in img_url.lower() else '.jpg'
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
            tmp.write(img_data)
            tmp.close()
            _img_cache[sku] = tmp.name
            return tmp.name
        except Exception as e:
            print(f'    {sku}: image download failed ({e})', flush=True)

    _img_cache[sku] = None
    return None


# ── PDF PARSING ───────────────────────────────────────────────────────────────
def parse_rainco_quote(pdf_bytes):
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        all_lines = []
        for page in pdf.pages:
            all_lines.extend((page.extract_text() or '').split('\n'))

    lines = [l.strip() for l in all_lines if l.strip()]
    result = {
        'quoteNum': '', 'quoteDate': '', 'customerName': '',
        'customerCompany': '', 'items': [],
        'totals': {'rrp': 0, 'discount': 0, 'subtotal': 0, 'shipping': 0, 'gst': 0, 'total': 0},
        'qrCodePath': None
    }

    # QR code extraction disabled

    for l in lines:
        m = re.match(r'Quote\s*#(\S+)', l)
        if m: result['quoteNum'] = m.group(1)
        dm = re.match(r'(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\w*\s+\d+,?\s*\d{4}', l)
        if dm: result['quoteDate'] = l

    def _dedup_line(line):
        """The PDF duplicates column data side-by-side — extract just one occurrence.
        Left column = Shipping Address, Right column = Customer. We want the right (customer) side."""
        words = line.split()
        n = len(words)
        if n == 0:
            return line
        # Try exact-half duplication — if both halves match, names are the same, return either
        for half in range(1, n // 2 + 1):
            if words[:half] == words[half:half*2]:
                return ' '.join(words[:half])
        # No duplication — short line is a single value, return as-is
        if n <= 3:
            return line
        # Two different columns: shipping (left) vs customer (right). Take right (customer) side.
        return ' '.join(words[n // 2:])

    def _extract_customer(start_i):
        """Extract customer name + company from lines starting at start_i."""
        if start_i + 1 < len(lines):
            result['customerName'] = _dedup_line(lines[start_i + 1])
        for j in range(start_i + 2, min(start_i + 10, len(lines))):
            ln = lines[j]
            if re.match(r'ITEMS\b.*PRICE\b', ln, re.I):
                break
            if re.match(r'SHIPPING METHOD|PAYMENT|^\d+\s|^Australia\b|^Standard\b|Tel\.|^\+61', ln, re.I):
                continue
            candidate = _dedup_line(ln)
            if not candidate or not re.match(r'^[A-Za-z]', candidate):
                continue
            if result['customerName'] and candidate.lower() == result['customerName'].lower():
                continue
            result['customerCompany'] = candidate
            break

    # Strategy 1: headers on same line ("SHIPPING ADDRESS ... CUSTOMER ...")
    for i, l in enumerate(lines):
        if 'SHIPPING ADDRESS' in l and 'CUSTOMER' in l:
            _extract_customer(i)
            break

    # Strategy 2: "CUSTOMER" as its own line (some quote formats)
    if not result['customerName']:
        for i, l in enumerate(lines):
            if re.match(r'^CUSTOMER\s*$', l.strip(), re.I):
                _extract_customer(i)
                break

    # Strategy 3: name appears just after the date line
    if not result['customerName']:
        date_idx = next((i for i, l in enumerate(lines)
                         if re.match(r'(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\w*\s+\d+,?\s*\d{4}', l)), -1)
        if date_idx >= 0:
            for j in range(date_idx + 1, min(date_idx + 6, len(lines))):
                candidate = lines[j].strip()
                if candidate and re.match(r'^[A-Z][a-z]', candidate) and '$' not in candidate:
                    result['customerName'] = candidate
                    break

    hi = next((i for i,l in enumerate(lines) if 'ITEMS' in l and 'ITEM TOTAL' in l), -1)
    fi = next((i for i,l in enumerate(lines) if re.match(r'^Subtotal', l, re.I)), -1)
    if hi >= 0 and fi > hi:
        result['items'] = _parse_items(lines[hi+1:fi])

    # Robust totals extraction: join lines from the totals region into one string
    # so that multi-line splits like "Subtotal (incl.\nGST)\n$2,091" are handled.
    # Go back a few lines from fi so we catch RRP / Trade Discount which appear
    # just before the Subtotal line in some quote formats.
    totals_start = max(0, fi - 8) if fi >= 0 else max(0, len(lines) - 30)
    totals_text  = ' '.join(lines[totals_start:])
    full_text    = ' '.join(lines)   # for fields that can appear anywhere

    def search_amount(pattern, text=totals_text):
        """Find the first $amount after pattern in text."""
        m = re.search(pattern, text, re.I)
        if not m: return 0
        ms = re.findall(r'\$([\d,]+(?:\.\d+)?)', text[m.end():m.end()+80])
        return float(ms[0].replace(',','')) if ms else 0

    result['totals']['rrp']      = search_amount(r'\bRRP\b', full_text)
    result['totals']['subtotal'] = search_amount(r'Subtotal\b')
    result['totals']['shipping'] = search_amount(r'\bShipping\b')
    result['totals']['gst']      = search_amount(r'of which GST|GST\s*\(1[0-9]')
    result['totals']['total']    = search_amount(r'GRAND\s*TOTAL|TOTAL\s*\(AUD\)')
    # Trade discount: look for "Trade Discount -$369" (the minus is before the $)
    dm = re.search(r'Trade\s*Discount\D{0,5}([\d,]+(?:\.\d+)?)', full_text, re.I)
    result['totals']['discount'] = float(dm.group(1).replace(',','')) if dm else 0

    return result


def _parse_items(table_lines):
    """
    Parse RainCo quote items. Handles two PDF formats:

    Normal (name fits on one line):
      "Product Name $OrigPrice"
      "$Tax Qty $Total"
      "SKU: XXXXX $SalePrice"

    Wrapped (long name splits across lines):
      "Product Name Part 1"
      "$OrigPrice"
      "Name Part 2 $Tax Qty $Total"   ← name suffix merged with B values
      "$SalePrice"
      "SKU: XXXXX"                     ← no price on SKU line
    """
    items = []
    # Pattern that identifies the qty/total portion of a B line
    # Matches: optional-$ digits  space  integer  space  $digits  at end of line
    B_PAT = re.compile(r'\$?([\d,]+(?:\.\d+)?)\s+(\d+)\s+\$([\d,]+(?:\.\d+)?)\s*$')

    sku_idxs = [i for i,l in enumerate(table_lines) if re.match(r'^SKU:\s*\S+', l, re.I)]

    for ki, sku_i in enumerate(sku_idxs):
        c_line = table_lines[sku_i]

        # SKU + optional sale price from the SKU line itself
        mc  = re.match(r'^SKU:\s*(\S+)\s+\$([\d,]+(?:\.\d+)?)', c_line, re.I)
        mc0 = re.match(r'^SKU:\s*(\S+)', c_line, re.I)
        sku = mc.group(1) if mc else (mc0.group(1) if mc0 else '')
        sku_sale_p = float(mc.group(2).replace(',','')) if mc else None

        # Collect lines belonging to this item (between previous SKU and this one)
        prev_end = sku_idxs[ki-1] + 1 if ki > 0 else 0
        item_lines = table_lines[prev_end:sku_i]

        if not item_lines:
            continue

        # Find the B line — the one containing the "tax qty total" pattern
        b_idx, b_match = None, None
        for j, ln in enumerate(item_lines):
            m = B_PAT.search(ln)
            if m:
                b_idx = j
                b_match = m
                break

        if b_idx is None or b_match is None:
            continue

        b_line   = item_lines[b_idx]
        pre_b    = item_lines[:b_idx]   # lines before B
        post_b   = item_lines[b_idx+1:] # lines after B (before SKU)

        # Any text on the B line before the tax/qty/total is a name continuation
        b_prefix = b_line[:b_match.start()].strip()

        # Determine orig price and name lines from pre-B lines
        orig_p     = None
        name_parts = []

        if pre_b:
            last = pre_b[-1]
            # Is the last pre-B line a bare price? e.g. "$330"
            lone_price = re.match(r'^\$([\d,]+(?:\.\d+)?)\s*$', last)
            if lone_price:
                orig_p     = float(lone_price.group(1).replace(',',''))
                name_parts = list(pre_b[:-1])
            else:
                # Normal case: last pre-B line ends with the orig price
                mp = re.match(r'^(.+?)\s+\$([\d,]+(?:\.\d+)?)\s*$', last)
                if mp:
                    name_parts = list(pre_b[:-1]) + [mp.group(1)]
                    orig_p     = float(mp.group(2).replace(',',''))
                else:
                    name_parts = list(pre_b)

        # Append the B-line name prefix if present
        if b_prefix:
            name_parts.append(b_prefix)

        name = ' '.join(name_parts).strip()

        # Tax / qty / total from B pattern groups
        tax   = float(b_match.group(1).replace(',',''))
        qty   = int(b_match.group(2))
        total = float(b_match.group(3).replace(',',''))

        # Sale price: prefer SKU line, then first post-B line if it's a bare price
        if sku_sale_p is not None:
            sale_p = sku_sale_p
        elif post_b:
            sp_m = re.match(r'^\$([\d,]+(?:\.\d+)?)\s*$', post_b[0])
            sale_p = float(sp_m.group(1).replace(',','')) if sp_m else orig_p
        else:
            sale_p = orig_p

        # Extract finish from name (part after last " - " or " – ")
        finish_m = re.search(r'[-–]\s*(.+)$', name)
        finish   = finish_m.group(1).strip() if finish_m else ''

        if name and sku:
            items.append({'name': name, 'sku': sku, 'finish': finish,
                          'origPrice': orig_p, 'salePrice': sale_p,
                          'tax': tax, 'qty': qty, 'total': total})
    return items


# ── PDF GENERATION ────────────────────────────────────────────────────────────
from reportlab.platypus import Flowable
from reportlab.lib.utils import ImageReader


class RoomHeader(Flowable):
    """Editorial room header: teal accent strip + Montserrat room name + thin rule."""
    H = 11 * mm

    def __init__(self, name, width):
        Flowable.__init__(self)
        self.room_name = name
        self.width = width
        self.height = self.H

    def draw(self):
        c = self.canv
        # Room name — dark, no green bar
        c.setFillColor(DARK)
        c.setFont(BFB(), 9)
        label = self.room_name.upper()
        c.drawString(0, 3*mm, label)
        # Thin grey rule extending to the right
        lw = c.stringWidth(label, BFB(), 9)
        c.setStrokeColor(colors.HexColor('#cccccc'))
        c.setLineWidth(0.4)
        c.line(lw + 4*mm, self.H / 2, self.width, self.H / 2)


class ProductCard(Flowable):
    """Canvas-drawn product card — adapts to any width (single or 2-col)."""
    CARD_H  = 42 * mm
    IMG_SZ  = 28 * mm
    QTY_W   = 20 * mm

    def __init__(self, item, qty, unit_total, img_path, card_w):
        Flowable.__init__(self)
        self.item       = item
        self.qty        = qty
        self.unit_total = unit_total
        self.img_path   = img_path
        self.width      = card_w
        self.height     = self.CARD_H

    # Estimate max chars for the available text column width
    def _max_chars(self):
        text_w = self.width - self.IMG_SZ - 2*mm - 2*mm - 3*mm - self.QTY_W - 2*mm
        return max(16, int(text_w / 1.6))   # ~1.6mm per char at 9pt regular

    def _wrap(self, name):
        """Word-aware wrap into up to 3 lines."""
        n = self._max_chars()
        words = name.split()
        lines, cur, cur_len = [], [], 0
        for word in words:
            extra = len(word) + (1 if cur else 0)
            if cur_len + extra <= n:
                cur.append(word); cur_len += extra
            else:
                if cur: lines.append(' '.join(cur))
                cur, cur_len = [word], len(word)
        if cur: lines.append(' '.join(cur))
        return lines[:3]

    def draw(self):
        c    = self.canv
        item = self.item
        W, H = self.width, self.CARD_H

        # ── Background + thin border ──
        c.setFillColor(WHITE)
        c.setStrokeColor(WARM_BORDER)
        c.setLineWidth(0.4)
        c.roundRect(0, 0, W, H, 1.5*mm, fill=1, stroke=1)

        # ── Image box — inset 2mm from card edge so it stays within rounded corners ──
        IMG = self.IMG_SZ
        ix  = 2*mm
        iy  = (H - IMG) / 2
        c.setFillColor(WHITE)
        c.rect(ix, iy, IMG, IMG, fill=1, stroke=0)
        if self.img_path:
            try:
                ir = ImageReader(self.img_path)
                c.drawImage(ir, ix, iy, IMG, IMG,
                            preserveAspectRatio=True, anchor='c', mask='auto')
            except Exception:
                self._placeholder(c, ix, iy, IMG, IMG)
        else:
            self._placeholder(c, ix, iy, IMG, IMG)

        # Divider — sits right after the image box
        divL = IMG + 2*mm
        c.setStrokeColor(WARM_BORDER);  c.setLineWidth(0.4)
        c.line(divL, 4*mm, divL, H - 4*mm)

        # ── Text column ──
        tx = divL + 3*mm
        y  = H - 8.5*mm

        c.setFillColor(DARK);  c.setFont(BF(), 9)
        for ln in self._wrap(item['name']):
            c.drawString(tx, y, ln);  y -= 4*mm

        if item.get('finish'):
            c.setFillColor(MID_GREY);  c.setFont(BF(), 8)
            c.drawString(tx, y, item['finish']);  y -= 4*mm

        c.setFillColor(colors.HexColor('#aaaaaa'));  c.setFont(BF(), 7)
        c.drawString(tx, y, f'SKU: {item["sku"]}');  y -= 6*mm

        # Price — strikethrough orig + dark sale
        orig = item.get('origPrice');  sale = item.get('salePrice') or orig or 0
        if orig and sale and abs(orig - sale) > 0.01:
            s  = f'${orig:,.2f}'
            c.setFillColor(colors.HexColor('#aaaaaa'));  c.setFont(BF(), 8)
            sw = c.stringWidth(s, BF(), 8)
            c.drawString(tx, y, s)
            c.setStrokeColor(colors.HexColor('#aaaaaa'));  c.setLineWidth(0.5)
            c.line(tx, y + 3.5, tx + sw, y + 3.5)
            c.setFillColor(DARK);  c.setFont(BFB(), 10)
            c.drawString(tx + sw + 2*mm, y, f'${sale:,.2f}')
        else:
            c.setFillColor(DARK);  c.setFont(BFB(), 10)
            c.drawString(tx, y, f'${sale:,.2f}')

        c.setFillColor(colors.HexColor('#aaaaaa'));  c.setFont(BF(), 6.5)
        c.drawString(tx, y - 3.5*mm, 'per unit, ex. GST')

        # ── Qty badge (right column) ──
        divR = W - self.QTY_W - 3*mm
        c.setStrokeColor(WARM_BORDER);  c.setLineWidth(0.4)
        c.line(divR, 4*mm, divR, H - 4*mm)

        cx = divR + self.QTY_W / 2
        cy = H * 0.5

        # Plain number — no circle, minimal
        c.setFillColor(DARK);  c.setFont(BFB(), 13)
        c.drawCentredString(cx, cy - 2*mm, str(self.qty))
        c.setFillColor(MID_GREY);  c.setFont(BF(), 6)
        c.drawCentredString(cx, cy - 6*mm, 'qty')

        c.setFillColor(DARK);  c.setFont(BFB(), 9)
        c.drawCentredString(cx, 9*mm, f'${self.unit_total:,.2f}')
        c.setFillColor(colors.HexColor('#aaaaaa'));  c.setFont(BF(), 6.5)
        c.drawCentredString(cx, 5.5*mm, 'inc. GST')

    def _placeholder(self, c, x, y, w, h):
        c.setFillColor(colors.HexColor('#f2efe9'))
        c.setStrokeColor(WARM_BORDER);  c.setLineWidth(0.4)
        c.rect(x, y, w, h, fill=1, stroke=1)
        c.setFillColor(MID_GREY);  c.setFont(BF(), 7)
        c.drawCentredString(x + w/2, y + h/2, 'No image')


class RoomSubtotal(Flowable):
    """Warm-tinted subtotal bar with teal typography."""
    H = 9 * mm

    def __init__(self, room_name, subtotal, width):
        Flowable.__init__(self)
        self.room_name = room_name
        self.subtotal  = subtotal
        self.width     = width
        self.height    = self.H

    def draw(self):
        c = self.canv
        # Minimal: just a thin top border line, no fill
        c.setStrokeColor(colors.HexColor('#cccccc'))
        c.setLineWidth(0.4)
        c.line(0, self.height, self.width, self.height)
        # Labels in dark — right aligned
        c.setFillColor(DARK)
        c.setFont(BFSB(), 8.5)
        c.drawRightString(self.width - 28*mm, 2*mm, f'{self.room_name} Subtotal')
        c.drawRightString(self.width - 1*mm,  2*mm, f'${self.subtotal:,.2f}')


def generate_formatted_pdf(quote, rooms, room_qtys):
    """Landscape A4 presentation PDF with cover page and 2-column product grid."""
    # Ensure fonts + assets are ready (no-op if background thread already loaded them)
    _init_fonts()
    _init_assets()

    buf    = io.BytesIO()
    W, H   = landscape(A4)          # 841.9 × 595.3 pt  (297 × 210 mm)
    LM = RM = 14*mm
    TM      = 22*mm                 # below header banner
    BM      = 16*mm                 # above footer banner
    CW      = W - LM - RM
    GAP     = 4*mm                  # gap between 2-col cards
    CARD_W  = (CW - GAP) / 2

    # ── Image fetch ──
    all_items = quote['items']
    qty_map   = {e['itemIndex']: e['qtys'] for e in room_qtys}
    print('  Fetching product images…', flush=True)
    img_paths = {}
    for item in all_items:
        path = get_product_image(item['name'], item['sku'])
        img_paths[item['sku']] = path
        print(f'    {item["sku"]}: {"✓" if path else "✗"}', flush=True)

    def room_entries(room):
        out = []
        for idx, item in enumerate(all_items):
            q = qty_map.get(idx, {}).get(room, 0)
            if q > 0:
                unit_total = (item['total'] / item['qty']) * q if item['qty'] else 0
                out.append((item, q, round(unit_total, 2)))
        return out

    # ── Page templates ──
    cover_frame   = Frame(0, 0, W, H, leftPadding=0, rightPadding=0,
                          topPadding=0, bottomPadding=0, id='cover')
    content_frame = Frame(LM, BM, CW, H - TM - BM, leftPadding=0, rightPadding=0,
                          topPadding=0, bottomPadding=0, id='body')

    def on_cover(canvas, doc):
        _draw_cover_page(canvas, W, H, quote)

    def on_content(canvas, doc):
        _draw_page_header(canvas, W, H, quote)
        _draw_page_footer(canvas, W, H)

    templates = [
        PageTemplate(id='cover',   frames=[cover_frame],   onPage=on_cover),
        PageTemplate(id='content', frames=[content_frame], onPage=on_content),
    ]
    doc = BaseDocTemplate(buf, pagesize=landscape(A4), pageTemplates=templates,
                          leftMargin=LM, rightMargin=RM,
                          topMargin=TM, bottomMargin=BM)

    # ── Story — starts with cover + forced page break ──
    story = [NextPageTemplate('content'), PageBreak()]

    # Customer banner
    cust = quote['customerName']
    if quote.get('customerCompany'):
        cust += f'  ·  {quote["customerCompany"]}'
    story.append(SP(2))
    story.append(P(f'Prepared for: <b>{cust}</b>',
                   S('cust', fontName=BF(), fontSize=9, textColor=MID_GREY)))
    story.append(SP(4))
    story.append(HRFlowable(width='100%', color=WARM_BORDER, thickness=0.5))
    story.append(SP(5))

    # ── Room sections ──
    assigned_rooms = [r for r in rooms if room_entries(r)]
    unassigned = [(item, 1, item['total']) for idx, item in enumerate(all_items)
                  if all(qty_map.get(idx, {}).get(r, 0) == 0 for r in rooms)]
    sections = [(r, room_entries(r)) for r in assigned_rooms]
    if unassigned:
        sections.append(('Other Items', unassigned))

    for room_idx, (room_name, entries) in enumerate(sections):
        room_story = []
        if room_idx > 0:
            room_story.append(PageBreak())
        room_story.append(RoomHeader(room_name, CW))
        room_story.append(SP(3))

        room_subtotal = 0
        # Pair items into 2-column rows
        pairs = [(entries[i], entries[i+1] if i+1 < len(entries) else None)
                 for i in range(0, len(entries), 2)]
        for left_e, right_e in pairs:
            def make_card(e):
                return ProductCard(e[0], e[1], e[2], img_paths.get(e[0]['sku']), CARD_W)
            left_card  = make_card(left_e)
            room_subtotal += left_e[2]
            if right_e:
                right_card = make_card(right_e)
                room_subtotal += right_e[2]
            else:
                right_card = Spacer(CARD_W, ProductCard.CARD_H)
            row = Table([[left_card, Spacer(GAP, 1), right_card]],
                        colWidths=[CARD_W, GAP, CARD_W])
            row.setStyle(TableStyle([
                ('VALIGN',        (0,0), (-1,-1), 'TOP'),
                ('TOPPADDING',    (0,0), (-1,-1), 0),
                ('BOTTOMPADDING', (0,0), (-1,-1), 0),
                ('LEFTPADDING',   (0,0), (-1,-1), 0),
                ('RIGHTPADDING',  (0,0), (-1,-1), 0),
            ]))
            room_story.append(row)
            room_story.append(SP(3))

        room_story.append(RoomSubtotal(room_name, room_subtotal, CW))
        room_story.append(SP(7))

        story.extend(room_story)

    # ── Totals ──
    story.append(HRFlowable(width='100%', color=LITE_GREY, thickness=0.5))
    story.append(SP(5))
    totals = quote['totals']

    def tot_row(label, val, bold=False, discount=False):
        if bold:
            sl = S('tbl', fontSize=10, textColor=DARK, alignment=TA_RIGHT, fontName=BFB())
            sv = S('tbv', fontSize=10, textColor=DARK, alignment=TA_RIGHT, fontName=BFB())
        elif discount:
            sl = S('tdl', fontSize=9, textColor=TEAL_MID, alignment=TA_RIGHT, fontName=BFSB())
            sv = S('tdv', fontSize=9, textColor=TEAL_MID, alignment=TA_RIGHT, fontName=BFSB())
        else:
            sl = S('tl', fontSize=9, textColor=MID_GREY, alignment=TA_RIGHT)
            sv = S('tv', fontSize=9, textColor=DARK, alignment=TA_RIGHT)
        val_str = f'${val:,.2f}' if not discount else f'\u2212${val:,.2f}'
        return [P(''), P(label, sl), P(val_str, sv)]

    tot_rows = []
    if totals.get('rrp', 0) > 0:
        tot_rows.append(tot_row('RRP',             totals['rrp']))
    if totals.get('discount', 0) > 0:
        tot_rows.append(tot_row('Trade Discount',  totals['discount'], discount=True))
    tot_rows.append(tot_row('Subtotal',            totals['subtotal']))
    tot_rows.append(tot_row('Shipping',            totals['shipping']))
    tot_rows.append(tot_row('GST (10%)',           totals['gst']))
    tot_rows.append(tot_row('TOTAL (AUD)',         totals['total'], bold=True))

    total_row_idx = len(tot_rows) - 1
    tot_table = Table(tot_rows, colWidths=[CW*0.50, CW*0.32, CW*0.18])
    tot_table.setStyle(TableStyle([
        ('ALIGN',         (0,0), (-1,-1), 'RIGHT'),
        ('TOPPADDING',    (0,0), (-1,-1), 3),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3),
        ('LINEABOVE',     (0,total_row_idx), (-1,total_row_idx), 0.5, LITE_GREY),
        ('TOPPADDING',    (0,total_row_idx), (-1,total_row_idx), 7),
        ('BOTTOMPADDING', (0,total_row_idx), (-1,total_row_idx), 7),
    ]))
    story.append(tot_table)
    story.append(SP(10))

    # ── End-of-document contact block ──
    story.append(HRFlowable(width='100%', color=WARM_BORDER, thickness=0.4))
    story.append(SP(4))
    contact_style = S('contact', fontName=BF(), fontSize=7, textColor=MID_GREY,
                      alignment=1)  # centred
    story.append(P('hello@rainco.com.au  ·  rainco.com.au  ·  ABN: 90 662 536 441  ·  '
                   'Factory 5, 7-11 Lindon Court, Tullamarine VIC 3043', contact_style))
    story.append(SP(2))
    story.append(P('Direct Debit  BSB: 083-004  ·  ACC: 71-870-4963  ·  '
                   'Please use your Quote number as payment reference',
                   S('contact2', fontName=BF(), fontSize=6, textColor=colors.HexColor('#bbbbbb'),
                     alignment=1)))

    doc.build(story)
    return buf.getvalue()



def _draw_cover_page(canvas, W, H, quote):
    """Landscape cover: dark brand panel left + editorial details right."""
    canvas.saveState()
    mid = W / 2
    PAD = 18*mm

    # Full page white base
    canvas.setFillColor(WHITE)
    canvas.rect(0, 0, W, H, fill=1, stroke=0)

    # ══ LEFT — hero image full bleed ══════════════════════════════════════════
    # Fallback dark background if image missing
    canvas.setFillColor(TEAL)
    canvas.rect(0, 0, mid, H, fill=1, stroke=0)

    hero_path = os.path.join(
        os.path.dirname(os.path.abspath(__file__)), 'cover_hero.jpg')
    if os.path.exists(hero_path):
        try:
            from PIL import Image as PILImage
            with PILImage.open(hero_path) as im:
                iw, ih = im.size
        except Exception:
            iw, ih = 2500, 3333

        # Scale to cover the left panel with centre-crop
        scale_w = mid / iw
        scale_h = H / ih
        scale   = max(scale_w, scale_h)
        draw_w  = iw * scale
        draw_h  = ih * scale
        draw_x  = (mid - draw_w) / 2
        draw_y  = (H - draw_h) / 2

        p = canvas.beginPath()
        p.rect(0, 0, mid, H)
        canvas.clipPath(p, stroke=0)
        canvas.drawImage(ImageReader(hero_path),
                         draw_x, draw_y, draw_w, draw_h,
                         preserveAspectRatio=False)
        canvas.restoreState()
        canvas.saveState()

    # ══ RIGHT — clean white editorial panel ══════════════════════════════════
    canvas.setFillColor(WHITE)
    canvas.rect(mid, 0, mid, H, fill=1, stroke=0)

    rx = mid + PAD

    # Dark logo (top of right panel)
    logo_d = _asset_paths.get('logo_dark')
    DLOGO_H = 10*mm
    if logo_d:
        try:
            canvas.drawImage(ImageReader(logo_d), rx, H - PAD - DLOGO_H,
                             DLOGO_H * 4.8, DLOGO_H, preserveAspectRatio=True,
                             anchor='sw', mask='auto')
        except Exception:
            logo_d = None
    if not logo_d:
        canvas.setFillColor(TEAL)
        canvas.setFont(BFB(), 18)
        canvas.drawString(rx, H - PAD - 7*mm, 'RainCo')

    # Thin dark green rule
    canvas.setStrokeColor(TEAL)
    canvas.setLineWidth(0.7)
    canvas.line(rx, H - PAD - 15*mm, W - PAD, H - PAD - 15*mm)

    # 'PRODUCT SPECIFICATION'
    canvas.setFillColor(MID_GREY)
    canvas.setFont(BF(), 7)
    canvas.drawString(rx, H - PAD - 22*mm, 'P R O D U C T   S P E C I F I C A T I O N')

    # Client name — large, vertically centred in right panel
    cy = H / 2 + 12*mm
    cname = quote.get('customerName', '')
    canvas.setFillColor(DARK)
    canvas.setFont(BFB(), 26)
    canvas.drawString(rx, cy, cname)

    if quote.get('customerCompany'):
        canvas.setFillColor(MID_GREY)
        canvas.setFont(BF(), 11)
        canvas.drawString(rx, cy - 10*mm, quote['customerCompany'])

    # Thin warm rule
    canvas.setStrokeColor(WARM_BORDER)
    canvas.setLineWidth(0.5)
    canvas.line(rx, cy - 17*mm, W - PAD, cy - 17*mm)

    # Quote detail block
    canvas.setFillColor(MID_GREY)
    canvas.setFont(BF(), 6.5)
    canvas.drawString(rx,         34*mm, 'QUOTE NUMBER')
    canvas.drawString(rx + 44*mm, 34*mm, 'DATE PREPARED')

    canvas.setFillColor(DARK)
    canvas.setFont(BFB(), 10)
    canvas.drawString(rx,         26*mm, f'#{quote["quoteNum"]}')
    canvas.drawString(rx + 44*mm, 26*mm, quote.get('quoteDate', ''))

    # ── Cart URL button — below quote details ──
    cart_url = quote.get('cartUrl')
    if cart_url:
        btn_text  = 'Pay your quote online \u2192'
        BTN_FONT  = 8
        BTN_PAD_X = 4 * mm   # horizontal padding each side
        BTN_PAD_Y = 2 * mm   # vertical padding each side
        # Measure text to size the box tightly around it
        canvas.setFont(BFSB(), BTN_FONT)
        txt_w   = canvas.stringWidth(btn_text, BFSB(), BTN_FONT)
        font_h  = BTN_FONT * (25.4 / 72)   # pt → mm
        btn_w   = txt_w + 2 * BTN_PAD_X
        BTN_H   = font_h * mm + 2 * BTN_PAD_Y
        btn_x   = rx
        btn_y   = 12 * mm   # lower on the page, below quote number block
        # White fill, teal border, rounded corners
        canvas.setFillColor(WHITE)
        canvas.setStrokeColor(TEAL)
        canvas.setLineWidth(0.8)
        canvas.roundRect(btn_x, btn_y, btn_w, BTN_H, 1.5 * mm, fill=1, stroke=1)
        # Teal label — baseline centred vertically
        # cap-height ≈ 68% of point size; centre = btn_y + (BTN_H - cap) / 2
        cap_h = BTN_FONT * 0.68          # points
        baseline = btn_y + (BTN_H - cap_h) / 2
        canvas.setFillColor(TEAL)
        canvas.drawString(btn_x + BTN_PAD_X, baseline, btn_text)
        # Clickable hyperlink
        canvas.linkURL(cart_url, (btn_x, btn_y, btn_x + btn_w, btn_y + BTN_H), relative=0)

    # ── QR code — bottom right of right panel ──
    qr_path = quote.get('qrCodePath')
    if qr_path and os.path.exists(qr_path):
        QR_SZ = 28*mm
        qr_x  = W - PAD - QR_SZ
        qr_y  = PAD
        try:
            canvas.drawImage(ImageReader(qr_path), qr_x, qr_y, QR_SZ, QR_SZ,
                             preserveAspectRatio=True, anchor='sw', mask='auto')
            canvas.setFillColor(MID_GREY)
            canvas.setFont(BF(), 6.5)
            canvas.drawCentredString(qr_x + QR_SZ / 2, qr_y - 4*mm, 'Scan to pay online')
        except Exception:
            pass

    canvas.restoreState()


def _draw_page_header(canvas, W, H, quote):
    """Clean white header — dark logo left, dark quote info right, thin bottom rule."""
    canvas.saveState()
    BAR = 14*mm

    # White background with subtle bottom rule
    canvas.setFillColor(WHITE)
    canvas.rect(0, H - BAR, W, BAR, fill=1, stroke=0)
    canvas.setStrokeColor(WARM_BORDER)
    canvas.setLineWidth(0.4)
    canvas.line(0, H - BAR, W, H - BAR)

    # Dark logo
    logo = _asset_paths.get('logo_dark')
    if logo:
        try:
            LOGO_H = 6*mm
            canvas.drawImage(ImageReader(logo), 14*mm, H - BAR/2 - LOGO_H/2,
                             LOGO_H * 4.5, LOGO_H, preserveAspectRatio=True,
                             anchor='sw', mask='auto')
        except Exception:
            logo = None
    if not logo:
        canvas.setFillColor(TEAL)
        canvas.setFont(BFB(), 12)
        canvas.drawString(14*mm, H - 9*mm, 'RainCo')

    # Quote info right-aligned in dark green
    canvas.setFillColor(TEAL)
    canvas.setFont(BFSB(), 8)
    canvas.drawRightString(W - 14*mm, H - 7*mm, f'Quote #{quote["quoteNum"]}')
    canvas.setFont(BF(), 6.5)
    canvas.setFillColor(MID_GREY)
    canvas.drawRightString(W - 14*mm, H - 12*mm, quote['quoteDate'])
    canvas.restoreState()


def _draw_page_footer(canvas, W, H):
    """Minimal footer — thin top rule + page number only."""
    canvas.saveState()
    BAR = 8*mm
    canvas.setStrokeColor(WARM_BORDER)
    canvas.setLineWidth(0.4)
    canvas.line(0, BAR, W, BAR)
    canvas.setFillColor(MID_GREY)
    canvas.setFont(BF(), 6.5)
    canvas.drawRightString(W - 14*mm, 3*mm, f'Page {canvas.getPageNumber() - 1}')
    canvas.restoreState()


# ── Embedded HTML frontend ────────────────────────────────────────────────────

# ── HTTP Server (Render-compatible: 0.0.0.0, PORT from env, CORS) ─────────────
class Handler(BaseHTTPRequestHandler):
    def log_message(self, *a): pass

    def _cors(self):
        self.send_header('Access-Control-Allow-Origin',  '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')

    def send_json(self, data, code=200):
        body = json.dumps(data).encode()
        self.send_response(code)
        self.send_header('Content-Type', 'application/json')
        self.send_header('Content-Length', len(body))
        self._cors()
        self.end_headers()
        self.wfile.write(body)

    def do_OPTIONS(self):
        self.send_response(204)
        self._cors()
        self.end_headers()

    def do_GET(self):
        if self.path in ('/', '/health'):
            body = b'{"status":"ok"}'
            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.send_header('Content-Length', len(body))
            self._cors()
            self.end_headers()
            self.wfile.write(body)
        else:
            self.send_error(404)

    def do_POST(self):
        length = int(self.headers.get('Content-Length', 0))
        body   = self.rfile.read(length)

        if self.path == '/parse':
            try:
                pdf = _extract_multipart(body, self.headers['Content-Type'])
                data = parse_rainco_quote(pdf)
                if not data['items']:
                    self.send_json({'error': 'No items found — is this a RainCo quote PDF?'}, 400)
                else:
                    self.send_json(data)
            except Exception as e:
                traceback.print_exc()
                self.send_json({'error': str(e)}, 500)

        elif self.path == '/generate':
            try:
                payload  = json.loads(body)
                pdf      = generate_formatted_pdf(
                    payload['quote'], payload['rooms'], payload['roomQtys'])
                self.send_response(200)
                self.send_header('Content-Type', 'application/pdf')
                self.send_header('Content-Length', len(pdf))
                self.send_header('Content-Disposition', 'attachment; filename=quote.pdf')
                self._cors()
                self.end_headers()
                self.wfile.write(pdf)
            except Exception as e:
                traceback.print_exc()
                self.send_json({'error': str(e)}, 500)
        else:
            self.send_error(404)


def _extract_multipart(body, content_type):
    import email
    raw = f'Content-Type: {content_type}\r\n\r\n'.encode() + body
    msg = email.message_from_bytes(raw)
    for part in msg.walk():
        if part.get_content_disposition() == 'form-data' and 'filename=' in part.get('Content-Disposition',''):
            return part.get_payload(decode=True)
    raise ValueError('No file found in upload')


if __name__ == '__main__':
    PORT = int(os.environ.get('PORT', 8742))
    print(f'RainCo Quote Formatter API', flush=True)
    print(f'Listening on 0.0.0.0:{PORT}', flush=True)
    server = HTTPServer(('0.0.0.0', PORT), Handler)
    def _preload():
        print('  Pre-loading fonts & assets…', flush=True)
        _init_fonts()
        _init_assets()
        print('  Ready.', flush=True)
    threading.Thread(target=_preload, daemon=True).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print('\nStopped.')
