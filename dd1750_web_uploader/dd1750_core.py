import math, os, re, csv
from collections import OrderedDict

import fitz  # PyMuPDF
from PIL import Image
import pytesseract
import openpyxl
import yaml

from pypdf import PdfReader, PdfWriter, PageObject
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch

BADWORDS = [
    "END ITEM","PUB","PAGE","DATE","IMAGE","LIN","DESC","ACOEI",
    "COMPONENT OF END","SLOC","UIC","NIIN","AUTH","QTYOH",
    "COEI","BOM","LOCATION","LOCID","WTY","ARC","CIIC","UI","SCMC"
]

def load_cfg(path: str) -> dict:
    with open(path, "r") as f:
        return yaml.safe_load(f)

def clean_mat(token: str) -> str:
    token = str(token).split("-C")[0]
    # FIXED regex: dash safely escaped
    return re.sub(r"[^0-9A-Za-z/\-]", "", token)

def looks_like_mat(tok: str) -> bool:
    return bool(re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9/\-]*", tok or "")) and len(tok) >= 3

def group_words_to_lines(words, y_tol=3.0):
    lines = []
    for w in sorted(words, key=lambda w: (round((w[1]+w[3])/2.0,1), w[0])):
        x0,y0,x1,y1,t,*_ = w
        y = (y0+y1)/2.0
        if not lines or abs(lines[-1]["y"]-y) > y_tol:
            lines.append({"y": y, "w": [(x0,t)]})
        else:
            lines[-1]["w"].append((x0,t))
    return lines

def extract_pdf_text_rows(pdf_path: str):
    doc = fitz.open(pdf_path)
    items = []
    for page in doc:
        words = page.get_text("words")
        if not words:
            continue
        W = page.rect.width
        cols = {
            "MAT": (0.06*W, 0.32*W),
            "DESC": (0.33*W, 0.82*W),
            "QTY": (0.83*W, 0.98*W)
        }
        lines = group_words_to_lines(words, y_tol=3.0)
        for L in lines:
            toks = sorted(L["w"], key=lambda z:z[0])
            mat = " ".join([t for x,t in toks if cols["MAT"][0] <= x <= cols["MAT"][1]]).strip()
            desc = " ".join([t for x,t in toks if cols["DESC"][0] <= x <= cols["DESC"][1]]).strip()
            qtys = " ".join([t for x,t in toks if cols["QTY"][0] <= x <= cols["QTY"][1]]).strip()

            if not mat or not desc or not qtys:
                continue

            up = (mat + " " + desc).upper()
            if any(b in up for b in BADWORDS):
                continue

            mat_tok = clean_mat(mat.split()[0])
            m = re.findall(r"\d+", qtys)
            if not m:
                continue
            qty = int(m[-1])
            if qty <= 0:
                continue

            items.append({
                "mat": mat_tok,
                "desc": re.sub(r"\s{2,}", " ", desc).strip(),
                "qty": qty
            })

    doc.close()
    return items

def ocr_page_items(page, dpi=250):
    pix = page.get_pixmap(dpi=dpi)
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    text = pytesseract.image_to_string(img, config="--psm 6")

    items = []
    for raw in text.splitlines():
        line = raw.strip()
        if not line or len(line) < 6:
            continue
        if not re.search(r"\d+\s*$", line):
            continue

        qty = int(re.search(r"(\d+)\s*$", line).group(1))
        if qty <= 0:
            continue

        first = line.split()[0]
        if not looks_like_mat(first):
            continue

        desc = re.sub(r"\s+\d+\s*$", "", line[len(first):]).strip(" ,;:-")
        up = (first + " " + desc).upper()
        if any(b in up for b in BADWORDS):
            continue

        mat_tok = clean_mat(first)
        if not mat_tok or not desc:
            continue

        items.append({
            "mat": mat_tok,
            "desc": re.sub(r"\s{2,}", " ", desc).strip(),
            "qty": qty
        })

    return items

def extract_pdf_ocr_rows(pdf_path: str, dpi=250, page_start=0, page_end=None):
    doc = fitz.open(pdf_path)
    items = []
    end = page_end if page_end is not None else len(doc)
    for i in range(page_start, min(end, len(doc))):
        items.extend(ocr_page_items(doc[i], dpi=dpi))
    doc.close()
    return items

def extract_excel_rows(excel_path: str, sheet=None,
                       col_desc="Description",
                       col_mat="Material",
                       col_qty="OH QTY"):
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb[sheet] if sheet else wb.active

    headers = {}
    for c in ws[1]:
        if c.value:
            headers[str(c.value).strip().lower()] = c.column

    def idx(name): return headers.get(name.lower())

    i_desc, i_mat, i_qty = idx(col_desc), idx(col_mat), idx(col_qty)
    if not (i_desc and i_mat and i_qty):
        raise ValueError(f"Missing columns. Found: {list(headers.keys())}")

    items = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        desc, mat, qty = r[i_desc-1], r[i_mat-1], r[i_qty-1]
        if desc is None or mat is None or qty is None:
            continue
        try:
            qty = int(qty)
        except:
            continue
        if qty <= 0:
            continue

        items.append({
            "mat": clean_mat(mat),
            "desc": str(desc).strip(),
            "qty": qty
        })

    return items

def aggregate(items):
    agg = OrderedDict()
    for it in items:
        key = (it["mat"], it["desc"])
        agg[key] = agg.get(key, 0) + int(it["qty"])
    return [{"mat":k[0], "desc":k[1], "qty":v} for k,v in agg.items()]

def write_audit(items, out_csv_path, qty_max):
    with open(out_csv_path, "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["line","mat","desc","qty","flag"])
        for i,it in enumerate(items, start=1):
            flag = "SUSPICIOUS_QTY" if it["qty"] > qty_max else ""
            w.writerow([i, it["mat"], it["desc"], it["qty"], flag])
