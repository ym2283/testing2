# ============================================================
#  PROFESSIONAL PDF GENERATOR — v7.2f
#  - Adds TABLE2 (same as TABLE, but 2 items share 1 page: top + bottom)
#  - Tightens <br> spacing (collapse multiple to single)
#  - Ensures TABLE2 never renders via format-2/3/4 path
#  - Subcategory header at top as usual
# ============================================================

import gspread
import pandas as pd
import os, re, shutil, tempfile, requests
from datetime import datetime
from typing import Optional, Tuple, List
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

from reportlab.lib.pagesizes import A4
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image,
    PageBreak, Flowable, KeepInFrame
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_LEFT
import html
from reportlab.lib import colors
from reportlab.platypus import Paragraph
from reportlab.lib.styles import ParagraphStyle
import re, html
from googleapiclient.http import MediaFileUpload
# ---------- helpers ----------
def _mm(v): return v * mm
PAGE_W_MM, PAGE_H_MM = 210, 297

def _norm(fmt: Optional[str]) -> str:
    if not fmt: return '2'
    f = str(fmt).strip().upper()
    if f in ('TABLE','TABLE2','1WP','2','3','4'): return f
    if f == '1B': return 'TABLE2'
    if f in ('2A','2B'): return '2'
    if f == '3A': return '3'
    if f == '4A': return '4'
    return '2'

def detect_shape_from_span(text: str) -> str:
    if not text:
        return ""

    raw = str(text).strip()

    if not raw.lower().startswith("<span>"):
        return str(text).strip()

    cleaned = raw[6:].strip()

    icon_map = {
        "flat": "/workspaces/testing2/Flat.png",
        "half round": "/workspaces/testing2/Half Round.png",
        "round": "/workspaces/testing2/Round.png",
        "triangle": "/workspaces/testing2/Triangle.png",
        "square": "/workspaces/testing2/Square.png",
    }

    low = cleaned.lower()
    for key, path in icon_map.items():
        if key in low:
            # IMPORTANT: ReportLab uses <image>, not <img>.
            icon_tag = f'<image file="{path}" width="12"/>'
            return f"{icon_tag} {cleaned}"

    return cleaned

def preprocess_size_data(size_text):
    """Pre-process size data to make it more table-friendly by breaking long lists into 2 lines."""
    if not size_text:
        return size_text
    
    size_text = str(size_text)
    
    # Remove extra spaces
    size_text = re.sub(r'\s*,\s*', ', ', size_text)
    
    # For very long lists, consider breaking into 2 lines
    if len(size_text) > 25 and ',' in size_text:
        parts = [p.strip() for p in size_text.split(',')]
        if len(parts) > 4:
            # Split into roughly equal parts
            mid = len(parts) // 2
            line1 = ', '.join(parts[:mid])
            line2 = ', '.join(parts[mid:])
            return f"{line1},\n{line2}"
    
    return size_text

def _s(v: Optional[object]) -> str:
    if v is None: return ''
    try:
        if isinstance(v, float) and pd.isna(v): return ''
    except Exception:
        pass
    return str(v).strip()

# ---------- custom flowables ----------
class ItemNameTrailingLine(Flowable):
    def __init__(self, text, fontName, fontSize, lineColor=colors.black, gap_mm=4):
        super().__init__()
        self.raw_text = str(text or "")
        self.fontName = fontName; self.fontSize = fontSize
        self.lineColor = lineColor; self.gap = gap_mm * mm
        self.height = fontSize + 2
    def wrap(self, availWidth, availHeight):
        self.availWidth = availWidth
        clean = re.sub(r'\([^>]*\)', '', self.raw_text).strip().upper()
        c = pdfmetrics.stringWidth; text = clean; ellipsis = "…"
        while c(text, self.fontName, self.fontSize) + self.gap > availWidth and len(text) > 1:
            text = text[:-1]
        if text != clean:
            while c(text + ellipsis, self.fontName, self.fontSize) + self.gap > availWidth and len(text) > 1:
                text = text[:-1]
            text += ellipsis
        self.text = text; self.textWidth = c(text, self.fontName, self.fontSize)
        return (availWidth, self.height)
    def draw(self):
        canv = self.canv; y = 0
        canv.setFont(self.fontName, self.fontSize); canv.setFillColor(colors.black)
        canv.drawString(0, y, self.text)
        start_x = self.textWidth + self.gap; end_x = self.availWidth + 13  # extend line to page edge feel
        if end_x > start_x:
            canv.setStrokeColor(self.lineColor); canv.setLineWidth(1.0)
            canv.line(start_x, y + self.fontSize * 0.35, end_x, y + self.fontSize * 0.35)

class SetSubcategoryForFooter(Flowable):
    def __init__(self, subcategory_text: str):
        super().__init__(); self.sub = (subcategory_text or '').upper()
    def wrap(self, w, h): return (0, 0)
    def draw(self): setattr(self.canv, "_current_subcategory", self.sub)

class FooterCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        self.generator = kwargs.pop('generator', None)
        self.raw_data = kwargs.pop('raw_data', None)
        self.doc_ref = kwargs.pop('doc_ref', None)
        super().__init__(*args, **kwargs)
    def draw_footer_now(self):
        sub = getattr(self, "_current_subcategory", "")
        if self.generator and self.doc_ref:
            self.generator._draw_footer(self, self.doc_ref, self.raw_data, sub)
    def showPage(self): self.draw_footer_now(); super().showPage()
    def save(self): self.draw_footer_now(); super().save()

class EllipsizedTextBox(Flowable):
    """Fixed-width text; manual wrap + in-place ellipsis; vertical centering."""
    def __init__(self, text, fontName, fontSize, max_width_pt, max_lines=1,
                 leading=None, align='LEFT', v_align='MIDDLE', textColor=colors.black):
        super().__init__()
        self.text = str(text or "")
        self.fontName = fontName
        self.fontSize = fontSize
        self.max_w = max_width_pt
        self.max_lines = max(1, int(max_lines))
        self.leading = leading if leading else fontSize
        self.align = align
        self.v_align = v_align
        self.textColor = textColor
        self.lines = None
        self.width = max_width_pt
        self.height = self.leading * self.max_lines

    def _stringWidth(self, s):
        return pdfmetrics.stringWidth(s, self.fontName, self.fontSize)

    def _wrap_lines(self):
        """Improved wrapping that tries to show all content without aggressive truncation."""
        if not self.text:
            self.lines = [""]
            return
        
        blocks = self.text.split('\n')
        lines = []
        
        # Helper function
        def fits(s): 
            return self._stringWidth(s) <= self.max_w
        
        for block in blocks:
            words = block.split()
            if not words:
                if len(lines) < self.max_lines:
                    lines.append("")
                continue
                
            current_line = ""
            
            for word in words:
                test_line = (current_line + " " + word).strip() if current_line else word
                
                if fits(test_line):
                    current_line = test_line
                else:
                    # Current line is full, add it
                    if current_line:
                        lines.append(current_line)
                        if len(lines) >= self.max_lines:
                            # We've hit our line limit
                            # Instead of truncating, see if we can append a little more
                            if len(lines[-1]) < 30:  # If last line is short
                                # Try to add part of the word
                                part = word
                                while part and not fits(lines[-1] + " " + part + "…"):
                                    part = part[:-1]
                                if part:
                                    lines[-1] = lines[-1] + " " + part + "…"
                            self.lines = lines
                            return
                    
                    # Start new line with current word
                    if fits(word):
                        current_line = word
                    else:
                        # Word is too long for one line - break it
                        part = word
                        while part and not fits(part + "…"):
                            part = part[:-1]
                        if part:
                            lines.append(part + "…")
                        else:
                            lines.append("…")
                        
                        if len(lines) >= self.max_lines:
                            self.lines = lines
                            return
                        current_line = ""
            
            # Add the last line of the block
            if current_line and len(lines) < self.max_lines:
                lines.append(current_line)
            
            if len(lines) >= self.max_lines:
                break
        
        # Ensure we don't exceed max_lines
        if len(lines) > self.max_lines:
            lines = lines[:self.max_lines]
            # If we had to cut, indicate with ellipsis on last line
            if lines[-1] and not lines[-1].endswith("…"):
                # Try to keep as much as possible
                while lines[-1] and not fits(lines[-1] + "…"):
                    lines[-1] = lines[-1][:-1]
                if lines[-1]:
                    lines[-1] = lines[-1] + "…"
        
        self.lines = lines

    def wrap(self, availWidth, availHeight):
        if self.lines is None: self._wrap_lines()
        return (self.width, self.height)

    def draw(self):
        if self.lines is None: self._wrap_lines()
        c = self.canv; c.saveState()
        c.setFont(self.fontName, self.fontSize)
        c.setFillColor(self.textColor)  # Use the specified text color
        total_h = self.leading * len(self.lines)
        if self.v_align == 'TOP':
            y = self.height - self.leading
        elif self.v_align == 'BOTTOM':
            y = 0
        else:
            y = (self.height + total_h)/2.0 - self.leading
        for line in self.lines:
            if self.align == 'CENTER':
                x = (self.width - pdfmetrics.stringWidth(line, self.fontName, self.fontSize)) / 2.0
            elif self.align == 'RIGHT':
                x = self.width - pdfmetrics.stringWidth(line, self.fontName, self.fontSize)
            else:
                x = 0
            c.drawString(max(0, x), max(0, y), line)
            y -= self.leading
        c.restoreState()

class PaddedBox(Flowable):
    """Fixed-size box with padding that draws a single child flowable."""
    def __init__(self, width, height, child, pad_l=3, pad_r=3, pad_t=0, pad_b=0, valign='MIDDLE'):
        super().__init__()
        self.w, self.h = width, height
        self.child = child
        self.pad_l, self.pad_r, self.pad_t, self.pad_b = pad_l, pad_r, pad_t, pad_b
        self.valign = valign
    def wrap(self, availW, availH):
        inner_w = max(1, self.w - (self.pad_l + self.pad_r))
        inner_h = max(1, self.h - (self.pad_t + self.pad_b))
        cw, ch = self.child.wrap(inner_w, inner_h)
        self.cw, self.ch = min(cw, inner_w), min(ch, inner_h)
        return (self.w, self.h)
    def draw(self):
        if self.valign == 'TOP':
            y = self.h - self.pad_t - self.ch
        elif self.valign == 'BOTTOM':
            y = self.pad_b
        else:
            y = (self.h - self.ch) / 2.0
        self.child.drawOn(self.canv, self.pad_l, y)

# -------- INLINE ICON + TEXT FLOWABLE (for TABLE format) --------
from reportlab.platypus import Flowable, Image, Paragraph

class InlineImageText(Flowable):
    def __init__(self, img_path, txt, style, img_width=12):
        Flowable.__init__(self)
        self.img = Image(img_path, width=img_width, height=img_width)
        self.txt = Paragraph(txt, style)
        self.img_width = img_width

        # Pre-measure sizes
        txt_w, txt_h = self.txt.wrap(1000, 1000)
        self.width = img_width + 2 + txt_w
        self.height = max(img_width, txt_h)

    def wrap(self, w, h):
        return (self.width, self.height)

    def draw(self):
        # Draw icon
        self.img.drawOn(self.canv, 0, 0)
        # Draw text beside it
        self.txt.drawOn(self.canv, self.img_width + 2, 0)



# ----------------------------- main class -----------------------------
class ProfessionalPDFGenerator:
    def upload_to_drive(self, file_path, folder_id):
        file_metadata = {
            'name': os.path.basename(file_path),
            'parents': [folder_id]
        }
        media = MediaFileUpload(file_path, mimetype='application/pdf')
        uploaded = self.drive_service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id'
        ).execute()
        print("Uploaded PDF to Google Drive with File ID:", uploaded.get('id'))


    def __init__(self, credentials_path: str, spreadsheet_id: str):
        self._left_margin_mm, self._right_margin_mm = 5, 5
        self._top_margin_mm,  self._bottom_margin_mm = 0.5, 14.5
        self._header_band_mm  = 14
        self.downloaded_images = {}

        self.detail_row_caps = {'1WP': 16, '2': 16, '3': 9, '4': 6, 'TABLE': 9999}
        self.page_offset_mm = {'2': 0, '3': 0, '4': 0, '1WP': 0, 'TABLE': 0}

        self.setup_google_services(credentials_path)
        self.spreadsheet_id = spreadsheet_id
        self.setup_directories()
        self.setup_custom_fonts()
        self.setup_pdf_styles()

        # per-SKU exclusions for spec table (unchanged)
        self._exclude_dim_weight_skus = {"DL241025", "DL241041", "DL241057", "DL3565"}

        # one-shot flag to hide footer for cover pages
        self._hide_footer_for_page = False

    # ---------- setup ----------
    def setup_google_services(self, credentials_path: str):
        scope = ['https://www.googleapis.com/auth/spreadsheets',
                 'https://www.googleapis.com/auth/drive']
        creds = Credentials.from_service_account_file(credentials_path, scopes=scope)
        self.gs_client = gspread.authorize(creds)
        self.drive_service = build('drive', 'v3', credentials=creds)

    def setup_directories(self):
        self.temp_dir = tempfile.mkdtemp()
        self.output_dir = os.path.join(self.temp_dir, 'output')
        self.images_dir = os.path.join(self.temp_dir, 'images')
        os.makedirs(self.output_dir, exist_ok=True)
        os.makedirs(self.images_dir, exist_ok=True)

    def setup_custom_fonts(self):
        self.fonts_available = {'Avenir-Black': False, 'Avenir-Book': False}
        paths = {
            'Avenir-Black': "/workspaces/testing2/Avenir-Black.ttf",
            'Avenir-Book':  "/workspaces/testing2/Avenir-Book.ttf"
        }
        for fam, p in paths.items():
            if os.path.exists(p):
                try:
                    pdfmetrics.registerFont(TTFont(fam, p))
                    self.fonts_available[fam] = True
                except Exception:
                    pass

    def get_font_name(self, preferred, fallback):
        return preferred if self.fonts_available.get(preferred) else fallback

    def setup_pdf_styles(self):
        self.styles = getSampleStyleSheet()
        avenir_black = self.get_font_name('Avenir-Black', 'Helvetica-Bold')
        avenir_book  = self.get_font_name('Avenir-Book',  'Helvetica')

        self.styles.add(ParagraphStyle('SubcategoryHeader', fontName=avenir_black, fontSize=30, leading=32, alignment=TA_LEFT, leftIndent=0, firstLineIndent=0, spaceBefore=0, spaceAfter=0))
        self.styles.add(ParagraphStyle('ItemName',          fontName=avenir_black, fontSize=20, leading=18))
        self.styles.add(ParagraphStyle('DetailKey',         fontName=avenir_black, fontSize=11, leading=11.0))
        self.styles.add(ParagraphStyle('DetailVal',         fontName=avenir_book,  fontSize=11, leading=11.0))
        self.styles.add(ParagraphStyle('DetailValBold',     fontName=avenir_black, fontSize=11, leading=11.0))
        self.styles.add(ParagraphStyle('DetailValRedBold', parent=self.styles['Normal'], fontName=avenir_black, fontSize=11, leading=11.0,textColor=colors.red))
        self.styles.add(ParagraphStyle('Footer',            fontName=avenir_black, fontSize=9,  leading=11))

    # ---------- layout helpers ----------
    def fixed_box(self, flowable, w, h):
        return KeepInFrame(w, h, [flowable], mode='truncate', vAlign='TOP')

    # ---------- utils ----------
    def clean_html_css(self, text):
        if text is None:
            return ""
        try:
            if isinstance(text, float) and pd.isna(text):
                return ""
        except Exception:
            pass

        s = str(text)

        # --- Step 1️⃣ Normalize all newline variants ---
        # Convert HTML or escaped newlines to real '\n'
        s = (
            s.replace("\\n", "\n")       # literal backslash-n
            .replace("\\r", "\n")       # literal backslash-r
            .replace("&#10;", "\n")     # HTML line feed
            .replace("&#13;", "\n")     # HTML carriage return
            .replace("&#xa;", "\n")     # another newline entity
            .replace("\r\n", "\n")      # Windows style
            .replace("\r", "\n")        # classic Mac style
        )

        # --- Step 2️⃣ Remove HTML tags ---
        s = re.sub(r'</?(?!image\b)[^>]+>', '', s, flags=re.IGNORECASE)


        # --- Step 3️⃣ Strip invisible junk (symbols / control chars) ---
        s = re.sub(r'[\u25A0-\u25FF\u2610-\u2613\uFFFD\uF0A7]', '', s)
        s = re.sub(r'[\x00-\x08\x0B-\x1F\x7F]', '', s)  # safe control chars only

        # --- Step 4️⃣ Collapse excessive newlines ---
        s = re.sub(r'\n{3,}', '\n\n', s)

        # --- Step 5️⃣ Trim each line individually ---
        s = "\n".join(line.strip() for line in s.split("\n"))

        # --- Step 6️⃣ Normalize spaces ---
        s = re.sub(r'[ \t]+', ' ', s)

        return s.strip()

    def safe_paragraph(self, text, style):
        cleaned = self.clean_html_css(text)

        # Split text where <br> appears (case-insensitive)
        cleaned = cleaned.replace('\n', '<br/>')

        # Split text where <br> appears (case-insensitive)
        parts = re.split(r'(?i)<br\s*/?>', cleaned)

        # Rebuild into Flowables separated by small Spacers
        flows = []
        for idx, part in enumerate(parts):
            part = part.strip()
            if part:
                tight_style = ParagraphStyle(
                    name=f"{style.name}_tight",
                    parent=style,
                    leading=style.fontSize + 1   # reduce from 11→9 or 12→10 etc
                )
                flows.append(Paragraph(part, tight_style))
            if idx < len(parts) - 1:
                # Insert 2-3pt space between lines (adjust as needed)
                flows.append(Spacer(1, 1))  # 3pt ≈ 1mm
        return KeepInFrame(9999, 9999, flows, mode='shrink')

    def extract_file_id(self, url: str) -> Optional[str]:
        if not url: return None
        for p in [r'/d/([a-zA-Z0-9_-]+)', r'id=([a-zA-Z0-9_-]+)']:
            m = re.search(p, url)
            if m: return m.group(1)
        return None

    def download_image(self, image_url: str) -> Optional[str]:
        if not image_url or _s(image_url) == '': return None
        if not hasattr(self, "_img_cache"): self._img_cache = {}
        if image_url in self._img_cache: return self._img_cache[image_url]
        try:
            if 'drive.google.com' in image_url:
                fid = self.extract_file_id(image_url)
                if fid:
                    path = self.download_drive_image(fid)
                    if path: self._img_cache[image_url] = path; return path
            elif str(image_url).startswith('http'):
                resp = requests.get(image_url, timeout=10); resp.raise_for_status()
                path = os.path.join(self.images_dir, f"img_{len(self._img_cache)}.jpg")
                with open(path, 'wb') as f: f.write(resp.content)
                self._img_cache[image_url] = path; return path
        except Exception:
            pass
        return None

    def download_drive_image(self, file_id: str) -> Optional[str]:
        try:
            request = self.drive_service.files().get_media(fileId=file_id)
            path = os.path.join(self.images_dir, f"{file_id}.jpg")
            with open(path, 'wb') as f:
                downloader = MediaIoBaseDownload(f, request)
                done = False
                while not done:
                    _, done = downloader.next_chunk()
            return path
        except Exception:
            return None

    def create_safe_image_box(self, path, max_w, max_h, height_cap=0.75, empty_placeholder=False):
        max_h = max_h * height_cap
        if path and os.path.exists(path):
            img = Image(path); img.hAlign = 'CENTER'
            return KeepInFrame(max_w, max_h, [img], mode='shrink', vAlign='TOP')
        if empty_placeholder:
            return Table([['']], colWidths=[max_w], rowHeights=[max_h],
                         style=TableStyle([('LEFTPADDING',(0,0),(-1,-1),0),
                                           ('RIGHTPADDING',(0,0),(-1,-1),0),
                                           ('TOPPADDING',  (0,0),(-1,-1),0),
                                           ('BOTTOMPADDING',(0,0),(-1,-1),0)]))
        ph = Paragraph("No Image", self.styles['DetailVal'])
        return Table([[ph]], colWidths=[max_w], rowHeights=[max_h],
                     style=TableStyle([('VALIGN',(0,0),(-1,-1),'MIDDLE'),
                                       ('ALIGN',(0,0),(-1,-1),'CENTER')]))

    def get_first_non_empty(self, row, keys):
        for k in keys:
            v = row.get(k)
            try:
                if v is None or (isinstance(v, float) and pd.isna(v)): continue
            except Exception:
                pass
            if _s(v) != '': return _s(v)
        return ""

    def get_best_image(self, row):
        return self.get_first_non_empty(row, ['Image URL','Image','IMAGE','image','Main Image','Photo'])

    def get_graph_image(self, row):
        return self.get_first_non_empty(row, ['Image URL Graph','Graph URL','Graph'])

    # ---------- header/footer ----------
    def create_item_name_with_line(self, product_name):
        s = self.styles['ItemName']
        return ItemNameTrailingLine(product_name, s.fontName, s.fontSize, colors.black)

    def create_subcategory_header(self, subcategory_name):
        """
        Auto-shrink text but keep TOTAL HEIGHT FIXED so layout does NOT shift.
        """
        if not subcategory_name:
            return []

        text = self.clean_html_css(subcategory_name).upper()

        base = self.styles['SubcategoryHeader']
        max_w = self._content_width_pts()

        fn = base.fontName
        original_fs = base.fontSize          # 30
        original_leading = base.leading      # 32
        min_fs = original_fs * 0.70          # shrink allowed

        # Shrink-only logic
        fs = original_fs
        def width(size):
            return pdfmetrics.stringWidth(text, fn, size)

        w = width(fs)
        while w > max_w and fs > min_fs:
            fs -= 0.5
            w = width(fs)

        # --- FIX: Keep the SAME height as original ---
        style = ParagraphStyle(
            'SubcategoryHeaderDynamic',
            parent=base,
            fontSize=fs,
            leading=original_leading,    # Keep height same as before
            leftIndent=0,
            firstLineIndent=0,
            alignment=TA_LEFT,
            spaceBefore=0,
            spaceAfter=4
        )

        para = Paragraph(text, style)
        table = Table([[para]], colWidths=[max_w])
        table.setStyle(TableStyle([
            ('LEFTPADDING',(0,0),(-1,-1),0),
            ('RIGHTPADDING',(0,0),(-1,-1),0),
            ('TOPPADDING',(0,0),(-1,-1),0),
            ('BOTTOMPADDING',(0,0),(-1,-1),0),
            ('ALIGN',(0,0),(-1,-1),'LEFT'),
        ]))

        return [table]

    # ---------- footer drawing ----------
    def _draw_footer(self, canv, doc, raw_data, subcategory_upper: str):
        if getattr(self, "_hide_footer_for_page", False):
            self._hide_footer_for_page = False
            return
        canv.saveState()
        footer_text_y = 3.3 * mm
        line_y        = 9.0  * mm
        canv.setStrokeColor(colors.black); canv.setLineWidth(1.0)
        canv.line(10 * mm, line_y, 200 * mm, line_y)
        canv.setFont(self.styles['Footer'].fontName, 9)
        left_text = self.get_first_non_empty(raw_data, ['Footer Left','Company','Brand']) or 'LSK Hardware Trading Sdn Bhd'
        canv.drawString(15 * mm, footer_text_y, left_text)
        canv.drawCentredString(105 * mm, footer_text_y, (subcategory_upper or ''))
        canv.drawRightString(195 * mm, footer_text_y, str(doc.page))
        canv.restoreState()

    # ---------- size math ----------
    def _frame_height_mm(self):
        return PAGE_H_MM - (self._top_margin_mm + self._bottom_margin_mm)

    def _compute_layout(self, per_page: int, desired_gap_mm: float, boost_mm: float) -> Tuple[float,float,float]:
        """Return (container_h_pts, row_h_pts, actual_inter_gap_mm)."""
        frame_h_mm = self._frame_height_mm()
        usable_mm  = frame_h_mm - self._header_band_mm
        SAFETY_MM = 1.5

        gaps_mm = max(0, per_page - 1) * desired_gap_mm
        base_container_mm = max(40, (usable_mm - gaps_mm - SAFETY_MM) / max(1, per_page))
        container_mm = base_container_mm + max(0, boost_mm)

        total_desired = per_page * container_mm + gaps_mm + SAFETY_MM
        if total_desired <= usable_mm or per_page == 1:
            inter_gap_mm = desired_gap_mm
        else:
            inter_gap_mm = max(0.0, (usable_mm - SAFETY_MM - per_page*container_mm) / max(1, per_page-1))
            total_now = per_page*container_mm + (per_page-1)*inter_gap_mm + SAFETY_MM
            if total_now > usable_mm:
                deficit = total_now - usable_mm
                container_mm -= deficit / per_page

        headline_mm = 10
        row_mm = max(22, container_mm - headline_mm)
        return _mm(container_mm), _mm(row_mm), inter_gap_mm

    def _content_width_pts(self):
        content_w_mm = PAGE_W_MM - (self._left_margin_mm + self._right_margin_mm)
        return content_w_mm * mm

    # ---------- table utilities ----------
    def _auto_col_widths_generic(self, headers, rows_text, total_w_pts,
                                base_min_mm=16, pad_pt=6, max_col_ratio=0.55):
        """
        Improved: auto-detect optimal widths, prevent large blank gaps,
        and ensure most values stay in one line without wrapping.
        """
        def sw(txt, fn, fs):
            return pdfmetrics.stringWidth(str(txt or ""), fn, fs)

        fn_h = self.styles['DetailKey'].fontName
        fn_v = self.styles['DetailVal'].fontName
        fs   = 9
        min_pt = base_min_mm * mm

        # --- 1️⃣ Measure maximum text width per column ---
        raw_widths = []
        for j, h in enumerate(headers):
            m = sw(h, fn_h, fs)
            for r in rows_text:
                if j < len(r):
                    m = max(m, sw(r[j], fn_v, fs))
            raw_widths.append(m + pad_pt)

        # --- 2️⃣ Special handling for Size and Height columns ---
        for j, h in enumerate(headers):
            header_clean = str(h).strip().lower()
            if header_clean == "height":
                # Cap Height column width to 130pt
                raw_widths[j] = min(raw_widths[j], 130)
            # 🔴 FIX: Also cap Size column width to prevent it from dominating
            elif header_clean == "size":
                # Cap Size column to reasonable width (30% of total or 100pt max)
                size_max = min(total_w_pts * 0.30, 100)
                raw_widths[j] = min(raw_widths[j], size_max)

        # --- 3️⃣ Normalize extremely large columns ---
        max_allowed = max_col_ratio * total_w_pts
        raw_widths = [min(w, max_allowed) for w in raw_widths]

        # --- 4️⃣ Calculate total and compression ratio ---
        total_raw = sum(raw_widths)
        if total_raw <= total_w_pts:
            # If total is smaller than page width, distribute leftover proportionally
            leftover = total_w_pts - total_raw
            # Give more leftover to longer columns (not short ones)
            avg = total_raw / len(raw_widths)
            weights = [min(2.0, w / avg) for w in raw_widths]
            total_weight = sum(weights)
            adj = [w + leftover * (weights[i] / total_weight) for i, w in enumerate(raw_widths)]
        else:
            # If total too wide, compress large columns more aggressively
            compress_factor = total_w_pts / total_raw
            # Columns > average shrink more, smaller shrink less
            avg = total_raw / len(raw_widths)
            adj = []
            for w in raw_widths:
                ratio = 0.7 + 0.6 * min(1.0, w / avg)
                adj.append(w * compress_factor * ratio)

        # --- 5️⃣ Enforce minimum width & normalize total exactly ---
        adj = [max(min_pt, w) for w in adj]
        factor = total_w_pts / sum(adj)
        final = [w * factor for w in adj]

        return final

    # ---------- fixed-size cell helper ----------
    def _clip_cell(self, text, style, col_width_pt, max_lines=1, pad_lr_pt=3, pad_tb_pt=0, valign='MIDDLE'):
        # --- INLINE ICON HANDLING FOR TABLE FORMAT ---
        # Only convert if input begins with "<span>"
        t_raw = str(text).strip()
        if t_raw.lower().startswith("<span>"):
            # remove <span>
            t_clean = t_raw[6:].strip()

            icon_map = {
                "flat": "/workspaces/testing2/Flat.png",
                "half round": "/workspaces/testing2/Half Round.png",
                "round": "/workspaces/testing2/Round.png",
                "triangle": "/workspaces/testing2/Triangle.png",
                "square": "/workspaces/testing2/Square.png",
            }
            low = t_clean.lower()
            for key, path in icon_map.items():
                if key in low:
                    # RETURN icon+text flowable immediately
                    return (
                        InlineImageText(
                            img_path=path,
                            txt=t_clean,
                            style=style,
                            img_width=12
                        ),
                        style.leading + 6  # row height
                    )

        # --- ICON HANDLING FOR TABLE FORMAT ---
        # Only convert if column value starts with "<span>"
        t_raw = str(text).strip()
        if t_raw.lower().startswith("<span>"):
            # remove <span>
            t_clean = t_raw[6:].strip()

            # mapping
            icon_map = {
                "flat": "/workspaces/testing2/Flat.png",
                "half round": "/workspaces/testing2/Half Round.png",
                "round": "/workspaces/testing2/Round.png",
                "triangle": "/workspaces/testing2/Triangle.png",
                "square": "/workspaces/testing2/Square.png",
            }

            low = t_clean.lower()
            for key, path in icon_map.items():
                if key in low:
                    # IMPORTANT: ReportLab wants <image>
                    text = f'<image file="{path}" width="12"/> {t_clean}'
                    break

        text_str = self.clean_html_css(text or "")

        # 🔴 SMART LINE DETECTION: Auto-calculate needed lines based on content
        # Count actual newlines already in the text
        actual_newlines = text_str.count('\n')
        
        # Calculate how many lines are needed based on text length and column width
        avg_chars_per_line = max(1, int((col_width_pt - 2*pad_lr_pt) / (style.fontSize * 0.6)))
        # Rough estimate: each character is about 60% of font size in width
        
        text_length = len(text_str)
        estimated_lines_needed = max(1, int(text_length / avg_chars_per_line) + 1)
        
        # Use whichever is larger: actual newlines or estimated lines
        lines_needed = max(actual_newlines + 1, estimated_lines_needed)
        
        # But respect the max_lines parameter as an upper bound
        eff_max_lines = min(lines_needed, max_lines)
        
        # For key fields (not value fields), we're more restrictive
        # If this is being called with max_lines=1 (for keys), keep it at 1
        if max_lines == 1:
            eff_max_lines = 1
        elif eff_max_lines < 2 and text_length > avg_chars_per_line:
            # If text is long but we calculated only 1 line, give it at least 2
            eff_max_lines = 2

        effective_leading = style.leading * 1.15
        fixed_h = max(1, eff_max_lines * effective_leading + 3.5 * pad_tb_pt)
    
        # Check if this is an Item Code that should be red
        text_color = getattr(style, 'textColor', colors.black)
    
        inner = EllipsizedTextBox(
            text=text_str,
            fontName=style.fontName,
            fontSize=style.fontSize,
            max_width_pt=max(1, col_width_pt - 2*pad_lr_pt),
            max_lines=eff_max_lines,  # Use calculated max_lines
            leading=style.leading * 1.25,
            align='LEFT',
            v_align='MIDDLE',
            textColor=text_color  
        )
        cell = PaddedBox(
            width=col_width_pt,
            height=fixed_h,
            child=inner,
            pad_l=pad_lr_pt, pad_r=pad_lr_pt, pad_t=pad_tb_pt, pad_b=pad_tb_pt,
            valign=valign
        )
        return cell, fixed_h

    # ---------- spec block ----------
    def build_specifications_card(
        self,
        raw_data,
        resolved_data,
        detail_limit,
        col_w,
        row_h,
        key_w_override=None
    ):
        from math import ceil
        from reportlab.platypus import Table, TableStyle, Paragraph
        from reportlab.lib.styles import ParagraphStyle
        from reportlab.lib import colors
        from reportlab.pdfbase import pdfmetrics

        sw = pdfmetrics.stringWidth

        key_style = self.styles['DetailKey']
        val_style = self.styles['DetailVal']
        val_bold  = self.styles['DetailValBold']
        val_red   = self.styles['DetailValRedBold']

        # 🔴 ALL multiline fields must use Paragraph (never EllipsizedTextBox)
        MULTILINE_KEYS = {
            'size',
            'package include',
            'package includes',
            'specification',
            'specifications',
            'features',
            'description',
            'material list'
        }

        rows = []

        # ----- Item Code (red) -----
        item_code = _s(resolved_data.get('Item Code') or resolved_data.get('Code'))
        if item_code:
            red_style = ParagraphStyle(
                'ItemCodeRed',
                parent=val_bold,
                textColor=colors.red
            )
            rows.append(('Item Code', item_code, red_style))

        exact_packing, exact_price, others = [], [], []

        _skip_dim_weight = item_code in self._exclude_dim_weight_skus


        # ----- Collect parameters -----
        for i in range(1, 21):
            k = f'Parameter{i}'
            key_label = _s(raw_data.get(k, ''))
            raw_val   = resolved_data.get(k, '')

            if not key_label or not raw_val:
                continue

            val = detect_shape_from_span(raw_val)
            val = self.clean_html_css(val)

            low = key_label.lower().strip()

            if _skip_dim_weight and low in ('weight', 'packing dimension'):
                continue

            if low == 'packing':
                exact_packing.append((key_label, val, val_bold))
            elif low == 'price':
                exact_price.append((key_label, val, val_bold))
            else:
                others.append((key_label, val, val_style))

        rows.extend(others + exact_packing + exact_price)

        if detail_limit:
            rows = rows[:detail_limit]

        if not rows:
            rows = [("—", "—", val_style)]

        # ----- Column widths -----
        STANDARD_KEY_RATIO = 0.40
        key_w = col_w * STANDARD_KEY_RATIO
        val_w = col_w - key_w

        # ----- Build table rows -----
        table_rows = []
        row_heights = []

        for key, val, style in rows:
            lowk = key.lower().strip()

            # --- Key cell (always clipped) ---
            key_cell, h1 = self._clip_cell(
                key,
                key_style,
                key_w,
                max_lines=2,
                pad_lr_pt=3,
                pad_tb_pt=0
            )

            # --- Value cell ---
            if lowk in MULTILINE_KEYS:
                # 🔥 REAL wrapping paragraph (NO clipping)
                effective_leading = style.leading * 1.25

                p = Paragraph(
                    self.clean_html_css(val).replace('\n', '<br/>'),
                    ParagraphStyle(
                        f'{lowk}_value',
                        parent=style,
                        leading=effective_leading,
                        wordWrap='LTR',
                        spaceBefore=0,
                        spaceAfter=0
                    )
                )

                pad_tb = 3
                pw, ph = p.wrap(val_w - 10, row_h)

                val_cell = PaddedBox(
                    width=val_w,
                    height=ph + pad_tb * 2,
                    child=p,
                    pad_l=5,
                    pad_r=5,
                    pad_t=pad_tb,
                    pad_b=pad_tb,
                    valign='MIDDLE'
                )

                h2 = ph + pad_tb * 2

            else:
                # 🔒 Single-line or short fields → clipped
                val_cell, h2 = self._clip_cell(
                    val,
                    style,
                    val_w,
                    max_lines=1 if lowk == 'item code' else 5,
                    pad_lr_pt=5,
                    pad_tb_pt=3,
                    valign='MIDDLE'
                )

            rh = max(h1, h2)
            table_rows.append([key_cell, val_cell])
            row_heights.append(rh)

        # ----- Inner table -----
        inner = Table(
            table_rows,
            colWidths=[key_w, val_w],
            rowHeights=row_heights
        )

        ts = [
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('LEFTPADDING', (0, 0), (-1, -1), 0),
            ('RIGHTPADDING', (0, 0), (-1, -1), 0),
            ('TOPPADDING', (0, 0), (-1, -1), -5),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
        ]

        if len(table_rows) > 1:
            ts.insert(0, ('LINEBELOW', (0, 0), (-1, -2), 0.5, colors.HexColor('#999999')))

        inner.setStyle(TableStyle(ts))

        # ----- Outer wrapper -----
        outer = Table([[inner]], colWidths=[col_w])
        outer.setStyle(TableStyle([
            ('LEFTPADDING', (0, 0), (-1, -1), 0),
            ('RIGHTPADDING', (0, 0), (-1, -1), 0),
            ('TOPPADDING', (0, 0), (-1, -1), 0),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ]))

        return outer


    # ---------- product block for 2/3/4 ----------
    def _product_block(self, product_data, container_h, row_h, img_w, spec_w, detail_limit,
                   gutter=0, left_pad=0, img_height_cap=0.75, key_w_override=None):
        elems = []
        raw = product_data['raw']; res = product_data['resolved']; imgs = product_data['images']
        name = self.get_first_non_empty(res, ['Item Name','Title','Product Name']) or 'UNKNOWN PRODUCT'

        elems.append(self.create_item_name_with_line(name))
        elems.append(Spacer(1, 8 * mm))

        main_url = self.get_best_image(imgs)
        main_path = self.download_image(main_url) if main_url else None
        image_box = self.create_safe_image_box(main_path, img_w, row_h, height_cap=img_height_cap)

        spec_card = self.build_specifications_card(raw, res, detail_limit, spec_w, row_h, key_w_override=key_w_override)
        spec_card_box = self.fixed_box(spec_card, spec_w, row_h)

        spec_cell = Table([[spec_card_box]], colWidths=[spec_w], rowHeights=[row_h])
        spec_cell.setStyle(TableStyle([
            ('LEFTPADDING',   (0,0), (-1,-1), 0),
            ('RIGHTPADDING',  (0,0), (-1,-1), 0),
            ('TOPPADDING',    (0,0), (-1,-1), -5),
            ('BOTTOMPADDING', (0,0), (-1,-1), 0),
            ('VALIGN',        (0,0), (-1,-1), 'TOP'),
        ]))

        row_cells, col_w = [], []
        if left_pad > 0:
            row_cells.append(''); col_w.append(left_pad)
        row_cells.append(image_box); col_w.append(img_w)
        row_cells.append('');        col_w.append(gutter)
        row_cells.append(spec_cell); col_w.append(spec_w)

        row = Table([row_cells], colWidths=col_w, rowHeights=[row_h])
        row.setStyle(TableStyle([
            ('VALIGN', (0,0), (-1,-1), 'TOP'),
            ('LEFTPADDING',   (0,0), (-1,-1), 0),
            ('RIGHTPADDING',  (0,0), (-1,-1), 0),
            ('TOPPADDING',    (0,0), (-1,-1), 0),
            ('BOTTOMPADDING', (0,0), (-1,-1), 0),
        ]))
        elems.append(row)

        wrapper = Table([[elems]], colWidths=[sum(col_w)], rowHeights=[container_h])
        wrapper.setStyle(TableStyle([
            ('VALIGN', (0,0), (-1,-1), 'TOP'),
            ('LEFTPADDING',   (0,0), (-1,-1), 0),
            ('RIGHTPADDING',  (0,0), (-1,-1), 0),
            ('TOPPADDING',    (0,0), (-1,-1), 0),
            ('BOTTOMPADDING', (0,0), (-1,-1), 0),
        ]))
        return [wrapper]
    
    def create_standard_spec_block(self, product_data, container_h, row_h,
                                   img_w, spec_w, detail_limit, gutter=0,
                                   left_pad=0, img_height_cap=0.9,
                                   key_w_override=None):
        elems = []
        raw = product_data['raw']; res = product_data['resolved']; imgs = product_data['images']
        name = self.get_first_non_empty(res, ['Item Name','Title','Product Name']) or 'UNKNOWN PRODUCT'

        elems.append(self.create_item_name_with_line(name))
        elems.append(Spacer(1, 8 * mm))

        main_url = self.get_best_image(imgs)
        main_path = self.download_image(main_url) if main_url else None
        image_box = self.create_safe_image_box(main_path, img_w, row_h, height_cap=img_height_cap)

        spec_card = self.build_specifications_card(raw, res, detail_limit, spec_w, row_h,
                                               key_w_override=key_w_override)
        spec_cell = Table([[spec_card]], colWidths=[spec_w], rowHeights=[row_h])
        spec_cell.setStyle(TableStyle([
            ('VALIGN',(0,0),(-1,-1),'TOP'),
            ('LEFTPADDING',(0,0),(-1,-1),0),
            ('RIGHTPADDING',(0,0),(-1,-1),0),
            ('TOPPADDING',(0,0),(-1,-1),-12),
            ('BOTTOMPADDING',(0,0),(-1,-1),0),
        ]))

        row_cells = []
        col_w = []
        if left_pad > 0:
            row_cells.append(''); col_w.append(left_pad)
        row_cells.append(image_box); col_w.append(img_w)
        row_cells.append(''); col_w.append(gutter)
        row_cells.append(spec_cell); col_w.append(spec_w)

        row = Table([row_cells], colWidths=col_w, rowHeights=[row_h])
        row.setStyle(TableStyle([
            ('VALIGN',(0,0),(-1,-1),'TOP'),
            ('LEFTPADDING',(0,0),(-1,-1),0),
            ('RIGHTPADDING',(0,0),(-1,-1),0),
            ('RIGHTPADDING', (1,0), (1,0), 20),
            ('TOPPADDING',(0,0),(-1,-1),0),
            ('BOTTOMPADDING',(0,0),(-1,-1),0),
        ]))
        elems.append(row)
        wrapper = Table([[elems]], colWidths=[sum(col_w)], rowHeights=[container_h])
        wrapper.setStyle(TableStyle([
            ('VALIGN',(0,0),(-1,-1),'TOP'),
            ('LEFTPADDING',(0,0),(-1,-1),0),
            ('RIGHTPADDING',(0,0),(-1,-1),0),
            ('TOPPADDING',(0,0),(-1,-1),0),
            ('BOTTOMPADDING',(0,0),(-1,-1),0),
        ]))
        return [wrapper]

    def create_format_2_layout(self, product_data, container_h, row_h, key_w_override=None):
        total_w = self._content_width_pts()
        left_pad = 6 * mm; gutter = 5 * mm; img_w = 85 * mm
        spec_w   = total_w - (left_pad + img_w + gutter)
        return self.create_standard_spec_block(product_data, container_h, row_h, img_w, spec_w, detail_limit=16,
                                   gutter=gutter, left_pad=left_pad, img_height_cap=0.90, key_w_override=key_w_override)

    def create_product_block_3(self, product_data, container_h, row_h, key_w_override=None):
        total_w = self._content_width_pts()
        left_pad = 23 * mm; gutter = 22 * mm; img_w = 50 * mm
        spec_w   = total_w - (left_pad + img_w + gutter)
        return self.create_standard_spec_block(product_data, container_h, row_h, img_w, spec_w, detail_limit=10,
                                   gutter=gutter, left_pad=left_pad, img_height_cap=0.85)

    def create_product_block_4(self, product_data, container_h, row_h, key_w_override=None):
        total_w = self._content_width_pts()
        left_pad = 22 * mm; gutter = 20 * mm; img_w = 53 * mm
        spec_w   = total_w - (left_pad + img_w + gutter)
        return self.create_standard_spec_block(product_data, container_h, row_h, img_w, spec_w, detail_limit=6,
                                   gutter=gutter, left_pad=left_pad, img_height_cap=0.75)

    # ---------- 1WP ----------
    def create_1wp_format(self, product_data):
        container_h, row_h, _gap = self._compute_layout(1, desired_gap_mm=0, boost_mm=4)
        total_w = self._content_width_pts()

        graph_url  = self.get_graph_image(product_data['images'])
        graph_path = self.download_image(graph_url) if graph_url else None
        graph_h = row_h * 0.52
        graph_w = total_w * 0.85

        graph_inner = self.create_safe_image_box(graph_path, graph_w, graph_h, height_cap=1.0, empty_placeholder=True)
        graph_row = Table([[graph_inner]], colWidths=[total_w])
        graph_row.setStyle(TableStyle([
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('LEFTPADDING', (0,0), (-1,-1), 0),
            ('RIGHTPADDING',(0,0), (-1,-1), 0),
            ('TOPPADDING',  (0,0), (-1,-1), 0),
            ('BOTTOMPADDING',(0,0), (-1,-1), 0),
        ]))

        left_pad = 1 * mm; gutter = 1 * mm; img_w = 100 * mm
        spec_w   = total_w - (left_pad + img_w + gutter)

        raw = product_data['raw']; res = product_data['resolved']
        name = self.get_first_non_empty(res, ['Item Name','Title','Product Name']) or 'UNKNOWN PRODUCT'

        elems = []
        item_name = self.create_item_name_with_line(name)
        item_name_wrapper = Table([[item_name]], colWidths=[self._content_width_pts()])
        item_name_wrapper.setStyle(TableStyle([
            ('LEFTPADDING',(0,0),(-1,-1), 0),
            ('RIGHTPADDING',(0,0),(-1,-1), 0),
            ('TOPPADDING',(0,0),(-1,-1), 0),
            ('BOTTOMPADDING',(0,0),(-1,-1), 0),
        ]))
        elems.append(item_name_wrapper)
        elems.append(Spacer(1, 4 * mm))
        elems.append(graph_row)
        elems.append(Spacer(1, 6 * mm))

        main_url  = self.get_best_image(product_data['images'])
        main_path = self.download_image(main_url) if main_url else None
        # Shift the image downward (adjust TOPPADDING as needed)
        image_box = Table(
            [[self.create_safe_image_box(main_path, img_w, row_h, height_cap=0.40)]],
            colWidths=[img_w],
            style=[
                ('TOPPADDING', (0,0), (-1,-1), 8),  
                ('LEFTPADDING', (0,0), (-1,-1), 0),
                ('RIGHTPADDING',(0,0),(-1,-1), 0),
                ('BOTTOMPADDING',(0,0),(-1,-1), 0),
            ]
        )

        spec_card = self.build_specifications_card(raw, res, detail_limit=16, col_w=spec_w, row_h=row_h)

        bottom_cells = []; col_w = []
        if left_pad > 0:
            bottom_cells.append(''); col_w.append(left_pad)
        bottom_cells.append(image_box); col_w.append(img_w)
        bottom_cells.append('');       col_w.append(gutter)
        bottom_cells.append(spec_card); col_w.append(spec_w)

        row = Table([bottom_cells], colWidths=col_w)
        row.setStyle(TableStyle([
            ('VALIGN', (0,0), (-1,-1), 'TOP'),
            ('LEFTPADDING', (0,0), (-1,-1), 0),
            ('RIGHTPADDING',(0,0),(-1,-1), 0),
            ('TOPPADDING',  (0,0), (-1,-1), 0),
            ('BOTTOMPADDING',(0,0),(-1,-1), 0),
        ]))
        elems.append(row)
        return elems

    # ---------- TABLE (one per page) ----------
    def create_table_format(self, group_data):
        elements = []
        if not group_data:
            return elements

        first = group_data[0]
        res0 = first['resolved']
        imgs0 = first['images']
        name = self.get_first_non_empty(res0, ['Item Name', 'Title', 'Product Name']) or 'Unknown Product'

        total_w_pts = self._content_width_pts()

        # Item name with trailing line
        item_name_flow = self.create_item_name_with_line(name)
        item_name_wrapper = Table([[item_name_flow]], colWidths=[total_w_pts])
        item_name_wrapper.setStyle(TableStyle([
            ('LEFTPADDING', (0,0), (-1,-1), 0),
            ('RIGHTPADDING',(0,0),(-1,-1), 0),
            ('TOPPADDING',  (0,0), (-1,-1), 0),
            ('BOTTOMPADDING',(0,0),(-1,-1), 0),
        ]))
        elements.append(item_name_wrapper)
        elements.append(Spacer(1, 10 * mm))

        # Header image (top banner)
        header_img_url  = self.get_best_image(imgs0)
        header_img_path = self.download_image(header_img_url) if header_img_url else None
        header_img_h = 60 * mm
        elements.append(self.create_safe_image_box(header_img_path, total_w_pts, header_img_h, height_cap=1.0, empty_placeholder=True))
        elements.append(Spacer(1, 10 * mm))

        # Build columns: regular + packing + price (skip weight/packing dimension)
        packing, price, regular = set(), set(), set()
        for product in group_data:
            for i in range(1, 21):
                k = f'Parameter{i}'
                t = self.clean_html_css(product['raw'].get(k, '') or '')
                if not t:
                    continue
                low = t.lower().strip()
                if low in ('weight', 'packing dimension'):
                    continue
                if low == 'packing':
                    packing.add(t)
                elif low == 'price':
                    price.add(t)
                else:
                    regular.add(t)

        # Build columns: keep Parameter1–20 order, but skip unwanted ones
        cols = []
        for i in range(1, 21):
            k = f'Parameter{i}'
            t = self.clean_html_css(group_data[0]['raw'].get(k, '') or '')
            if not t:
                continue
            low = t.lower().strip()
            if low in ('weight', 'packing dimension'):
                continue
            if t not in cols:  # avoid duplicates
                cols.append(t)
        headers = ['Item Code'] + cols

        if len(cols) > 6:
            cols = cols[:6]

        headers = ['Item Code'] + cols
        data = [[self.safe_paragraph(h, self.styles['DetailKey']) for h in headers]]

        # 🔴 Define a per-paragraph red bold style ONCE (guaranteed red)
        red_style = ParagraphStyle(
            'ItemCodeRed',
            parent=self.styles['DetailValBold'],
            textColor=colors.red
        )

        rows_text = []
        for product in group_data:
            raw = product['raw']
            res = product['resolved']
            code = _s(res.get('Item Code') or res.get('Code'))

            row_vals = [code]
            # Force red using the inline style (cannot be overridden by table-wide FONT rules)
            row_cells = [Paragraph(code or "", red_style)]

            for c in cols:
                val = ''
                for i in range(1, 21):
                    k = f'Parameter{i}'
                    param_label = self.clean_html_css(raw.get(k, '') or '')
                    if not param_label:
                        continue
                    low_param = param_label.lower().strip()
                    if low_param in ('weight', 'packing dimension'):
                        continue
                    if param_label == c:
                        val = _s(res.get(k, ''))
                        # 🔴 FIX: Pre-process Size column data to break long lists
                        if c.lower().strip() == 'size':
                            val = preprocess_size_data(val)
                        break
                style = self.styles['DetailValBold'] if c.lower().strip() in ('packing','price') else self.styles['DetailVal']
                row_cells.append(self.safe_paragraph(val, style))
                row_vals.append(val)

            data.append(row_cells)
            rows_text.append(row_vals)

        col_widths = self._auto_col_widths_generic(headers, rows_text, total_w_pts,
                                                base_min_mm=16, pad_pt=6, max_col_ratio=0.55)

        t = Table(data, colWidths=col_widths, repeatRows=1)
        n_rows = len(data)
        style_cmds = [
            ('FONT', (0,0), (-1,-1), self.styles['DetailKey'].fontName),
            ('FONTSIZE', (0,0), (-1,-1), 9),
            ('BACKGROUND', (0,0), (-1,0), colors.white),
            ('LEFTPADDING',  (0,0), (-1,-1), 2),
            ('RIGHTPADDING', (0,0), (-1,-1), 2),
            ('TOPPADDING',   (0,1), (-1,-1), 6),
            ('BOTTOMPADDING',(0,1), (-1,-1), 10),
            ('TOPPADDING',   (0,0), (-1,0), 4),
            ('BOTTOMPADDING',(0,0), (-1,0), 4),
            ('ALIGN',        (0,0), (-1,0), 'LEFT'),
            ('VALIGN',       (0,0), (-1,-1), 'MIDDLE'),
        ]
        if n_rows > 2:
            style_cmds.insert(3, ('LINEBELOW', (0,1), (-1,-2), 0.4, colors.HexColor('#999999')))

        style_cmds += [
            ('WORDWRAP', (0,0), (-1,-1), 'None'),   # disable wrapping
            ('TRUNCATE', (0,0), (-1,-1), True),     # clip long text instead of wrapping
            ('ALIGN', (0,0), (-1,-1), 'LEFT'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ]  
          
        t.setStyle(TableStyle(style_cmds))

        for row_idx in range(1, len(data)):  # skip header row
            t.setStyle(TableStyle([
                ('TEXTCOLOR', (0, row_idx), (0, row_idx), colors.red),
            ]))

        table_wrapper = Table([[t]], colWidths=[total_w_pts])
        table_wrapper.setStyle(TableStyle([
            ('LEFTPADDING', (0, 0), (-1, -1), 0),
            ('RIGHTPADDING', (0, 0), (-1, -1), 0),
            ('TOPPADDING', (0, 0), (-1, -1), 0),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
        ]))
        elements.append(table_wrapper)
        return elements


    # ---------- TABLE2 (two TABLE blocks per page: top + bottom) ----------
    def _table_block(self, group_data):
        """Builds a single TABLE-style block (item name + image + spec table)."""
        # Reuse create_table_format but avoid repeated returns – we need the elements
        return self.create_table_format(group_data)

    def create_full_page_cover(self, image_url: str):
        path = self.download_image(image_url)
        if not path or not os.path.exists(path): return []
        generator = self
        class FullPageImage(Flowable):
            def __init__(self, img_path):
                super().__init__(); self.img_path = img_path; self.width, self.height = A4
            def wrap(self, availWidth, availHeight): return (0, 0)
            def drawOn(self, canv, x, y, _sW=0):
                generator._hide_footer_for_page = True
                canv.saveState()
                canv.drawImage(self.img_path, 0, 0, width=A4[0], height=A4[1])
                canv.restoreState()
        return [FullPageImage(path), PageBreak()]

    # ---------- grouping helper (GroupID-aware pagination for 2/3/4) ----------
    # ---------- grouping helper (GroupID-aware pagination for 2/3/4) ----------
    def _paginate_groups(self, items: List[dict], per_page: int) -> List[List[dict]]:
        clusters: List[List[dict]] = []
        cur: List[dict] = []
        last_gid = None
    
        for it in items:
            gid = _s(it['raw'].get('Group ID',''))
            # Handle null/empty Group IDs - treat them as "no grouping"
            if gid == '':
                gid = None
            
            if last_gid is None or gid == last_gid:
                cur.append(it)
                last_gid = gid
            else:
                clusters.append(cur)
                cur = [it]
                last_gid = gid
            
        if cur:
            clusters.append(cur)

        pages: List[List[dict]] = []
    
        for cl in clusters:
            need = len(cl)
            if need > per_page:
                # Split large cluster across multiple pages
                for i in range(0, len(cl), per_page):
                    pages.append(cl[i:i + per_page])
            else:
                # Try to add to current page if it fits
                if pages and (len(pages[-1]) + need <= per_page):
                    pages[-1].extend(cl)
                else:
                    pages.append(cl)

        # Debug print
        print(f"DEBUG: Paginating {len(items)} items into {len(pages)} pages with {per_page} per page")
        for i, p in enumerate(pages):
            group_ids = set(_s(item['raw'].get('Group ID', '')) for item in p)
            print(f"  Page {i+1}: {len(p)} items, Group IDs: {group_ids}")

        return pages

    # ---------- build ----------
    def generate_professional_pdf(self, output_path: str = None) -> str:
        if not output_path:
            ts = datetime.now().strftime('%Y%m%d_%H%M%S')
            output_path = os.path.join(self.output_dir, f'professional_catalog_{ts}.pdf')

        master_df   = self.get_sheet_data('Master')
        resolved_df = self.get_sheet_data('Master_Resolved')
        images_df   = self.get_sheet_data('Master_With_Images')
        remark_df   = self.get_sheet_data('Remark')
        if master_df.empty: raise Exception("Master sheet is empty")

        remark_lookup = {}
        if not remark_df.empty:
            for _, row in remark_df.iterrows():
                cat = _s(row.get('Category'))
                url = _s(row.get('Cover Page URL') or row.get('URL') or row.get('Cover URL'))
                if cat: remark_lookup[cat.upper()] = url

        merged, total = [], min(len(master_df), len(resolved_df), len(images_df))
        for i in range(total):
            m = master_df.iloc[i].to_dict()
            r = resolved_df.iloc[i].to_dict()
            im = images_df.iloc[i].to_dict()
            fmt = _norm(m.get('Format', '2'))
            m['Format'] = r['Format'] = fmt
            merged.append({'raw': m, 'resolved': r, 'images': im})

        # Build (Category, Format, SubCategory[, Group ID for TABLE/TABLE2]) groups
        groups, current, prev = [], [], (None, None, None)
        for item in merged:
            fmt = _norm(item['raw'].get('Format','2'))
            category = (item['raw'].get('Category','') or '').strip()
            subcategory = item['raw'].get('SubCategory','')
            if fmt in ('TABLE','TABLE2'):
                gid = _s(item['raw'].get('Group ID', '')).strip()
                if gid == '':
                    # generate a synthetic unique ID for blank group IDs
                    # this ensures each consecutive blank section under same subcategory forms its own group
                    gid = f"_blank_{len(groups)}_{len(current)}"
                key = (category, fmt, subcategory, gid)
            else:
                key = (category, fmt, subcategory)
            if key != prev and current:
                groups.append({'category': prev[0], 'format': prev[1], 'subcategory': prev[2], 'rows': current})
                current = []
            current.append(item); prev = key
        if current:
            groups.append({'category': prev[0], 'format': prev[1], 'subcategory': prev[2], 'rows': current})

        doc = SimpleDocTemplate(
            output_path,
            pagesize=A4,
            topMargin=_mm(self._top_margin_mm -1),
            bottomMargin=_mm(self._bottom_margin_mm -2),
            leftMargin=_mm(self._left_margin_mm),
            rightMargin=_mm(self._right_margin_mm),
        )

        story = []
        base_raw = merged[0]['raw'] if merged else {}

        # MAIN COVER (footer suppressed by flowable)
        main_cover_url = remark_lookup.get('DELI CATALOGUE COVER')
        if main_cover_url:
            story.extend(self.create_full_page_cover(main_cover_url))

        prev_category = None
        first_group = True  # Track if this is the very first group

        # Use indexed loop to allow TABLE2 to consume next group
        i = 0
        n = len(groups)
        while i < n:
            g = groups[i]
        
            fmt = _norm(g['format'])
            sub = g['subcategory'] or ''
            cat = _s(g['category']).upper()

            # ONLY add page break when changing categories (not for first group after cover)
            if not first_group and cat != prev_category:
                story.append(PageBreak())
        
            cover_added = False
            if cat != prev_category:
                cover_url = remark_lookup.get(cat)
                if cover_url:
                    story.extend(self.create_full_page_cover(cover_url))
                    cover_added = True
                prev_category = cat

            # If we added a category cover, we're already on a new page
            # Otherwise, only add page break if this isn't the first group
            if story and not cover_added and not first_group:
                story.append(PageBreak())

            items = g['rows']

            # Pure TABLE (1 per page)
            if fmt == 'TABLE':
                story.append(SetSubcategoryForFooter(sub))
                story.extend(self.create_subcategory_header(sub))
                story.append(Spacer(1, _mm(self._header_band_mm - 11.5)))
                story.extend(self.create_table_format(items))
                i += 1
                first_group = False
                continue

            # TABLE2 (two Group-ID tables per page: top + bottom)
            if fmt == 'TABLE2':
                story.append(SetSubcategoryForFooter(sub))
                story.extend(self.create_subcategory_header(sub))
                story.append(Spacer(1, _mm(self._header_band_mm - 11.5)))

                def _same_section(g1, g2):
                    return (
                        _norm(g2['format']) == 'TABLE2' and
                        _s(g1['category']).upper() == _s(g2['category']).upper() and
                        (g1['subcategory'] or '') == (g2['subcategory'] or '')
                    )

                # Render TOP table (current Group ID)
                story.extend(self.create_table_format(items))

                used = 1
                # Try to render BOTTOM table (next Group ID) on the same page
                if i + 1 < n and _same_section(g, groups[i + 1]):
                    story.append(Spacer(1, 14 * mm))
                    story.extend(self.create_table_format(groups[i + 1]['rows']))
                    used = 2

                i += used
                first_group = False
                continue

            # 1WP (one per page)
            if fmt == '1WP':
                for idx, prod in enumerate(items):
                    if idx > 0:
                        story.append(PageBreak())
                    story.append(SetSubcategoryForFooter(sub))
                    story.extend(self.create_subcategory_header(sub))
                    story.append(Spacer(1, _mm(self._header_band_mm - 12)))
                    story.extend(self.create_1wp_format(prod))
                i += 1
                first_group = False
                continue  

            # Formats 2/3/4 with GroupID-aware pagination
            if fmt in ('2','3','4'):
                per_page = {'2':2, '3':3, '4':4}[fmt]
                desired_gap = {'2':15, '3':12, '4':2}[fmt]
                boost = {'2':5, '3':4, '4':0}[fmt]
                cont_h, row_h, inter_gap_mm = self._compute_layout(per_page, desired_gap, boost)

                pages = self._paginate_groups(items, per_page)

                # Process each page
                for page_idx, page_items in enumerate(pages):
                    # only start new page if not the first page of this section
                    if page_idx > 0:
                        story.append(PageBreak())

                    # Add subcategory header on EVERY page for multi-page subcategories
                    story.append(SetSubcategoryForFooter(sub))
                    header_blocks = self.create_subcategory_header(sub)
                    if header_blocks:
                        story.extend(header_blocks)
                    story.append(Spacer(1, _mm(self._header_band_mm - 11.5)))

                    # Add products for this page
                    for j, prod in enumerate(page_items):
                        if j > 0 and inter_gap_mm > 0:
                            story.append(Spacer(1, _mm(inter_gap_mm)))
                        if fmt == '2':
                            story.extend(self.create_format_2_layout(prod, cont_h, row_h))
                        elif fmt == '3':
                            story.extend(self.create_product_block_3(prod, cont_h, row_h))
                        elif fmt == '4':
                            story.extend(self.create_product_block_4(prod, cont_h, row_h))
    
                i += 1
                first_group = False
                continue

            # Fallback treat as format 2
            per_page = 2
            desired_gap = 15
            boost = 5
            cont_h, row_h, inter_gap_mm = self._compute_layout(per_page, desired_gap, boost)
            pages = self._paginate_groups(items, per_page)

            for page_idx, page_items in enumerate(pages):
                if page_idx > 0:
                    story.append(PageBreak())
                story.append(SetSubcategoryForFooter(sub))
                story.extend(self.create_subcategory_header(sub))
                story.append(Spacer(1, _mm(self._header_band_mm - 12)))
                for j, prod in enumerate(page_items):
                    if j > 0 and inter_gap_mm > 0:
                        story.append(Spacer(1, _mm(inter_gap_mm)))
                    story.extend(self.create_format_2_layout(prod, cont_h, row_h))
            i += 1
            first_group = False

        doc.build(story, canvasmaker=lambda *a, **k: FooterCanvas(*a, **k, generator=self, raw_data=base_raw, doc_ref=doc))        
        return output_path

    # ---------- data ----------
    def get_sheet_data(self, sheet_name: str) -> pd.DataFrame:
        try:
            spreadsheet = self.gs_client.open_by_key(self.spreadsheet_id)
            worksheet = spreadsheet.worksheet(sheet_name)
            data = worksheet.get_all_records()
            return pd.DataFrame(data)
        except Exception as e:
            print(f"❌ Error loading {sheet_name}: {e}")
            return pd.DataFrame()

# ----------------------------- runner -----------------------------
def main():
    credentials_json = os.getenv('GOOGLE_CREDENTIALS_JSON')
    spreadsheet_id   = os.getenv('SPREADSHEET_ID')
    if not credentials_json or not spreadsheet_id:
        print("❌ Missing GOOGLE_CREDENTIALS_JSON or SPREADSHEET_ID")
        return
    temp_creds = tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False)
    temp_creds.write(credentials_json); temp_creds.close()
    try:
        gen = ProfessionalPDFGenerator(temp_creds.name, spreadsheet_id)
        out = gen.generate_professional_pdf()
        if out and os.path.exists(out):
            shutil.copy2(out, "./PROFESSIONAL_CATALOG.pdf")
            print("DONE: PDF GENERATED")
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback; traceback.print_exc()
    finally:
        os.unlink(temp_creds.name)

if __name__ == "__main__":
    main()