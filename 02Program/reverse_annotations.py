# -*- coding: utf-8 -*-
"""
reverse_annotations.py

Reverse-generate the 42-column Annotations mapping file from the manually
annotated aCRF PDF.

Sources of truth:
  * aCRF.pdf          - annotated CRF; SDTM annotations are PDF /FreeText objects
                        (text in /Contents, geometry in /Rect, color in /C)
  * als.xlsx          - SDS export: Form / Field / LabTableTitle /
                        DataDictionaryEntry sheets provide fieldOID<->label
                        lookups and dictionary entry lists.
  * Annotations.xlsx  - template ONLY: row 1 (the 42 column headers) is read
                        byte-exact; data rows are read solely for the old/new
                        diff section of the QC report.

Outputs:
  * Annotations_reverse.xlsx
  * QC_report_reverse.txt

Usage:
  python reverse_annotations.py [--acrf ..] [--als ..] [--template ..] [--out ..] [--qc ..]
"""

import argparse
import os
import re
import sys
from collections import defaultdict

import openpyxl
import pdfplumber
from pypdf import PdfReader
from pypdf.generic import ArrayObject, DictionaryObject, IndirectObject

# ---------------------------------------------------------------------------
# constants
# ---------------------------------------------------------------------------

# Helvetica AFM character widths (units per 1000 em); at 12pt width = u/1000*12.
AFM = {
    ' ': 278, '!': 278, '"': 355, '#': 556, '$': 556, '%': 889, '&': 667,
    "'": 191, '(': 333, ')': 333, '*': 389, '+': 584, ',': 278, '-': 333,
    '.': 278, '/': 278, '0': 556, '1': 556, '2': 556, '3': 556, '4': 556,
    '5': 556, '6': 556, '7': 556, '8': 556, '9': 556, ':': 278, ';': 278,
    '<': 584, '=': 584, '>': 584, '?': 556, '@': 1015,
    'A': 667, 'B': 667, 'C': 722, 'D': 722, 'E': 667, 'F': 611, 'G': 778,
    'H': 722, 'I': 278, 'J': 500, 'K': 667, 'L': 556, 'M': 833, 'N': 722,
    'O': 778, 'P': 667, 'Q': 778, 'R': 722, 'S': 667, 'T': 611, 'U': 722,
    'V': 667, 'W': 944, 'X': 667, 'Y': 667, 'Z': 611, '[': 278, '\\': 278,
    ']': 278, '^': 469, '_': 556, '`': 333,
    'a': 556, 'b': 556, 'c': 500, 'd': 556, 'e': 556, 'f': 278, 'g': 556,
    'h': 556, 'i': 222, 'j': 222, 'k': 500, 'l': 222, 'm': 833, 'n': 556,
    'o': 556, 'p': 556, 'q': 556, 'r': 333, 's': 500, 't': 278, 'u': 556,
    'v': 500, 'w': 722, 'x': 500, 'y': 500, 'z': 500, '{': 334, '|': 260,
    '}': 334, '~': 584,
}
CJK_WIDTH = 12.0


def text_width(s):
    """Helvetica 12pt text width; CJK chars get the 12.0 fallback."""
    w = 0.0
    for ch in s:
        if ch == '\n':
            w += 3.336  # treated as a space for width purposes
        elif ch in AFM:
            w += AFM[ch] / 1000.0 * 12.0
        else:
            w += CJK_WIDTH
    return w


#: page -> formOID fallback for pages whose header regex fails (filled in after
#: a first run via the QC report when needed)
PAGE_FORM_FALLBACK = {
    4: 'SV', 7: 'IE', 10: 'VS3', 14: 'QS', 20: 'EG', 22: 'LBHEMA',
    37: 'AT', 40: 'FU', 44: 'VS2', 51: 'AE', 60: 'LBOTH', 61: 'PCU',
}

#: forms that may carry a 表格行 (continuation table header band) row
TABLE_ROW_FORMS = {'VS2', 'VS1', 'PC', 'AE', 'CM'}

#: TESTCD code-value aliases (annot text value -> als fieldOID)
ALIASES = {
    'EGHRMN': 'HR', 'PRAG': 'PR', 'RRAG': 'RR', 'QRSAG': 'QRS',
    'QTAG': 'QT', 'QTCFAG': 'QTCF', 'PROT': 'TP', 'KETONEBD': 'KETONES',
    'UWBCMIC': 'UWBC1', 'URBCMIC': 'URBC',
}

#: value patterns -> target fieldOID template (only applied when the resulting
#: fieldOID exists in the page form's als field set)
CODE_PATTERNS = [
    (re.compile(r'^INCL\d+$'), 'IETEST'),
    (re.compile(r'^EXCL\d+$'), 'IETEST1'),
    (re.compile(r'^SCINT(\d+)$'), 'QSRES'),
    (re.compile(r'^ADINT(\d+)$'), 'QS1RES'),
]

#: per-form special VAR mappings.  A str value is the target fieldOID;
#: a dict value maps the annot value to the target fieldOID; None means
#: "known, deliberately left for the grid/QC pass".  Only applied when the
#: target fieldOID exists in the page form's als field set.
SPECIAL = {
    'QS': {
        'SUCAT': {'白酒': 'QSRES15CA1', '啤酒': 'QSRES15CA2',
                  '葡萄酒': 'QSRES15CA3', '其他': 'QSRES15CA4'},
        'SUDUR': 'QSRES15DUR', 'SUDOSE': 'QSRES15', 'SUDOSU': 'QSRES15',
        'SUTRT': 'QSRES17', 'SUEXDUR': 'QSRES17DUR', 'SUEXDOS': 'QSRES17C',
        'FASTDUR': None, 'FASTCAT': None,
    },
    'IE': {'IEORRES': '__GRID__'},
    'AT': {'LBTESTCD': 'ATTESTCD', 'LBTEST': 'ATTEST', 'LBORRES': 'ATORRES'},
    'PR7': {'MBTEST': 'PR7TEST', 'MBTESTCD': 'PR7TESTCD', 'MBORRES': 'PR7ORRES'},
    'PC': {'PCSTAT': 'PCND'},
    'PCU': {'PCTEST': 'PCUTY'},
}

#: LabTableTitle column VARS
TITLE_VARS = ['LBTEST', 'LBORRES', 'LBORRESU', 'LBORNRLO', 'LBORNRHI',
              'LBSIG', 'LBDNLK', 'LBCOM', 'LBSTNRC']

#: formOID -> SDTM domain prefix used for the Domain column when the page
#: carries no domain annotation (default: first two letters of formOID)
DOMAIN_BY_FORM = {
    'VS1': 'VS', 'VS2': 'VS', 'VS3': 'VS', 'LBHEMA': 'LB', 'LBCHEM': 'LB',
    'LBURI': 'LB', 'LBCOAG': 'LB', 'LBHCG': 'LB', 'LBOTH': 'LB',
    'QS1': 'QS', 'PCU': 'PC', 'DSSCR': 'DS', 'DSEOT': 'DS', 'DS1': 'DS',
}

DOMAIN_NAME_RE = re.compile(r'^([A-Z]{2})[（(].+[）)]$')
DOMAIN_CAT_RE = re.compile(r'^([A-Z]{2,4})CAT\s*=')
EQ_RE = re.compile(r'^([A-Z][A-Z0-9]*)\s*=\s*(.*)$')
SUPP_RE = re.compile(r'^([A-Z][A-Z0-9]*)\s+in\s+SUPP', re.I)
TESTCD_RE = re.compile(r'^[A-Z]+TESTCD$')
DATE_HINT_RE = re.compile(r'[（(]\s*\d{4}\s*[-/.]\s*\d{1,2}\s*[-/.]\s*\d{1,2}\s*[）)]')

HEADER = [
    'Page', 'FormOID', 'Form', 'FieldOID', 'Field', 'Item', 'Domain',
    'annotation', 'BACKGROUND', 'X1', 'Y1', 'X2', 'Y2', 'X1_Field',
    'Y1_Field', 'Y2_Field', 'X1_Item', 'Y1_Item', 'Y2_Item', 'Number',
    'Width_annotation', 'Item_All', 'Item_C', 'Item_E', 'Width_itemc',
    'Width_iteme', 'Width_item', 'Width_item_max', 'Item_up', 'Anno_low',
    'Multi', 'Gap', 'i', 'Annotation_', 'Width_annotation_a',
    'Width_annotation_', 'X2_Field', 'Multi_top', 'Highlight',
    'DASHED_BORDER', 'FieldOID_Lag', 'item',
]


def norm_text(s):
    """Normalize label text: full-width punct -> half-width, no whitespace,
    no (yyyy-mm-dd) date hints."""
    t = str(s or '')
    for fw, hw in [('（', '('), ('）', ')'), ('：', ':'), ('？', '?'),
                   ('，', ','), ('。', '.'), ('\u3000', ' ')]:
        t = t.replace(fw, hw)
    t = DATE_HINT_RE.sub('', t)
    return re.sub(r'\s+', '', t)


# ---------------------------------------------------------------------------
# A. extraction
# ---------------------------------------------------------------------------

def resolve(obj):
    if isinstance(obj, IndirectObject):
        return resolve(obj.get_object())
    if isinstance(obj, ArrayObject):
        out = []
        for x in obj:
            r = resolve(x)
            if isinstance(r, list):
                out.extend(r)
            elif r is not None:
                out.append(r)
        return out
    if isinstance(obj, DictionaryObject):
        return obj
    return None


def norm_color(c):
    """/C array -> '0.75 1 1' style string (5 decimals, trailing zeros cut)."""
    c = resolve(c)
    if not isinstance(c, ArrayObject) or not len(c):
        return '0.75 1 1'
    parts = []
    for v in c:
        s = ('%.5f' % float(v)).rstrip('0').rstrip('.')
        parts.append(s if s else '0')
    return ' '.join(parts)


def extract_annots(pdf_path):
    """-> dict page_no -> list of annot dicts (in PDF /Annots array order)."""
    reader = PdfReader(pdf_path)
    pages = {}
    for i, page in enumerate(reader.pages, 1):
        a = page.get('/Annots')
        if a is None:
            continue
        items = resolve(a)
        annots = []
        for d in items:
            if d is None or d.get('/Subtype') != '/FreeText':
                continue
            rect = [float(x) for x in d.get('/Rect')]
            text = str(d.get('/Contents', ''))
            text = text.replace('\r\n', '\n').replace('\r', '\n')
            bs = d.get('/BS')
            dashed = '1' if isinstance(bs, DictionaryObject) and \
                bs.get('/D') is not None else None
            annots.append({
                'page': i, 'idx': len(annots), 'x1': rect[0], 'y1': rect[1],
                'x2': rect[2], 'y2': rect[3], 'text': text,
                'flat': re.sub(r'\s+', ' ', text).strip(),
                'color': norm_color(d.get('/C')), 'dashed': dashed,
            })
        if annots:
            pages[i] = annots
    return pages


def extract_page_text(pdf_path):
    """-> dict page_no -> {'lines': [...], 'chars': [...]} in RAW y-up coords."""
    out = {}
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages, 1):
            H = float(page.height)
            chars = [c for c in page.chars if c.get('text') and
                     str(c.get('text')).strip()]
            lines = []
            # cluster by top (tolerance ~0.3pt)
            chars_sorted = sorted(chars, key=lambda c: (round(c['top'] * 2), c['x0']))
            groups = []
            for c in chars_sorted:
                key = round(c['top'] * 2)
                if groups and groups[-1][0] == key:
                    groups[-1][1].append(c)
                else:
                    groups.append([key, [c]])
            for _, grp in groups:
                grp.sort(key=lambda c: c['x0'])
                sub = []
                for c in grp:
                    if sub and c['x0'] - sub[-1]['x1'] > 20:
                        lines.append(sub)
                        sub = []
                    sub.append(c)
                if sub:
                    lines.append(sub)
            line_objs = []
            for sub in lines:
                text = ''.join(str(c['text']) for c in sub)
                x0 = min(c['x0'] for c in sub)
                x1 = max(c['x1'] for c in sub)
                top = min(c['top'] for c in sub)
                bottom = max(c['bottom'] for c in sub)
                line_objs.append({
                    'text': text, 'x0': x0, 'x1': x1,
                    'top': top, 'bottom': bottom,
                    'y1': H - top, 'y0': H - bottom,
                })
            line_objs.sort(key=lambda l: -l['y1'])
            char_objs = [{
                'text': str(c['text']), 'x0': c['x0'], 'x1': c['x1'],
                'top': c['top'], 'bottom': c['bottom'],
                'y1': H - c['top'], 'y0': H - c['bottom'],
            } for c in chars]
            out[i] = {'lines': line_objs, 'chars': char_objs, 'H': H}
    return out


# ---------------------------------------------------------------------------
# A2. als loading
# ---------------------------------------------------------------------------

def col_index(hdr, key):
    for i, h in enumerate(hdr):
        if h and key.lower() in str(h).lower():
            return i
    return None


def load_als(path):
    wb = openpyxl.load_workbook(path, data_only=True)

    # Form sheet
    forms = {}
    if 'Form' in wb.sheetnames:
        ws = wb['Form']
        hdr = [c.value for c in ws[1]]
        i_form = col_index(hdr, 'formOID')
        i_name = col_index(hdr, 'formName') or col_index(hdr, 'name')
        i_style = col_index(hdr, 'formStyle') or col_index(hdr, 'style')
        for row in ws.iter_rows(min_row=2, values_only=True):
            oid = row[i_form] if i_form is not None else None
            if oid is None:
                continue
            forms[str(oid)] = {
                'name': str(row[i_name]) if i_name is not None and row[i_name] else str(oid),
                'style': str(row[i_style]) if i_style is not None and row[i_style] else '',
            }

    # Field sheet
    fields = defaultdict(dict)          # formOID -> {fieldOID: info}
    norm_labels = defaultdict(dict)     # formOID -> {norm_label: [fieldOID]}
    ws = wb['Field']
    hdr = [c.value for c in ws[1]]
    i_form = col_index(hdr, 'formOID')
    i_field = col_index(hdr, 'fieldOID')
    i_label = col_index(hdr, 'label')
    i_dict = col_index(hdr, 'dataDictionaryOID')
    i_unit = col_index(hdr, 'unit')
    for row in ws.iter_rows(min_row=2, values_only=True):
        form = str(row[i_form]) if i_form is not None and row[i_form] else None
        oid = str(row[i_field]) if i_field is not None and row[i_field] else None
        if form is None or oid is None:
            continue
        label = str(row[i_label]) if i_label is not None and row[i_label] else ''
        info = {
            'fieldOID': oid, 'label': label, 'label_n': norm_text(label),
            'dict': str(row[i_dict]) if i_dict is not None and row[i_dict] else None,
            'unit': str(row[i_unit]) if i_unit is not None and row[i_unit] else '',
        }
        fields[form][oid] = info
        norm_labels[form].setdefault(info['label_n'], []).append(oid)

    # DataDictionaryEntry sheet
    dicts = defaultdict(list)
    if 'DataDictionaryEntry' in wb.sheetnames:
        ws = wb['DataDictionaryEntry']
        hdr = [c.value for c in ws[1]]
        i_dict = col_index(hdr, 'dataDictionaryOID')
        i_ord = col_index(hdr, 'ordinal')
        i_item = col_index(hdr, 'itemDataString')
        for row in ws.iter_rows(min_row=2, values_only=True):
            oid = row[i_dict] if i_dict is not None else None
            if oid is None:
                continue
            item = str(row[i_item]) if i_item is not None and row[i_item] else ''
            ord_ = row[i_ord] if i_ord is not None else 0
            dicts[str(oid)].append((ord_ or 0, item))
    for oid in dicts:
        dicts[oid] = [it for _, it in sorted(dicts[oid])]

    # LabTableTitle sheet: formOID -> {VAR: 中文标题}
    titles = defaultdict(dict)
    if 'LabTableTitle' in wb.sheetnames:
        ws = wb['LabTableTitle']
        rows = list(ws.iter_rows(min_row=2, values_only=True))
        for row in rows:
            vals = [v for v in row]
            form = None
            start = 0
            if vals and vals[0] and str(vals[0]) in forms:
                form = str(vals[0])
                start = 1
            # single pair per row (VAR, title)
            if len(vals) > start + 1 and vals[start] in TITLE_VARS and vals[start + 1]:
                titles[form][str(vals[start])] = str(vals[start + 1])
            # pairs embedded in the row
            i = start
            while i < len(vals):
                if vals[i] in TITLE_VARS and i + 1 < len(vals) and vals[i + 1]:
                    titles[form][str(vals[i])] = str(vals[i + 1])
                    i += 2
                else:
                    i += 1
    # global fallback: first title seen per VAR
    global_titles = {}
    for form in titles:
        for var, t in titles[form].items():
            global_titles.setdefault(var, t)

    wb.close()
    return forms, fields, norm_labels, dicts, titles, global_titles


# ---------------------------------------------------------------------------
# A3. page form detection
# ---------------------------------------------------------------------------

def detect_forms(pages_annots, pages_text):
    """-> dict page_no -> formOID (None when the page has no annots)."""
    out = {}
    for pg in pages_text:
        if pg not in pages_annots:
            out[pg] = None
            continue
        form = None
        band = [l for l in pages_text[pg]['lines'] if 790 <= l['y1'] <= 820]
        for l in sorted(band, key=lambda x: -x['y1']):
            m = re.search(r'(?:HS-10542-113\s*)?(.+?)\s+([A-Z][A-Z0-9]*)\s*$',
                          l['text'])
            if m and re.search(r'[一-鿿]', m.group(1)):
                form = m.group(2)
                break
        if form is None:
            form = PAGE_FORM_FALLBACK.get(pg)
        out[pg] = form
    return out


# ---------------------------------------------------------------------------
# B. matching
# ---------------------------------------------------------------------------

def vertical_overlap(a, b):
    """a/b have y1 (top, raw) and y2-or-y0 (bottom, raw). -> overlap length."""
    ab = a.get('y2', a.get('y0'))
    bb = b.get('y2', b.get('y0'))
    return min(a['y1'], b['y1']) - max(ab, bb)


def find_label_bands(page_text, fields, norm_labels):
    """Match page lines to als field labels.
    -> {fieldOID: [band dicts (y1,y0,x0,x1)]}"""
    lines = page_text['lines']
    # label candidates: single lines + up-to-3-line joins with same x0
    cands = []
    for l in lines:
        cands.append({'text': l['text'], 'y1': l['y1'], 'y0': l['y0'],
                      'x0': l['x0'], 'x1': l['x1']})
    for i in range(len(lines)):
        base = lines[i]
        joined, n = base['text'], 1
        y1, y0 = base['y1'], base['y0']
        x1 = base['x1']
        for j in range(i + 1, len(lines)):
            nxt = lines[j]
            if abs(nxt['x0'] - base['x0']) > 3 or (y0 - nxt['y1']) > 8 or n >= 3:
                break
            joined += nxt['text']
            y1 = nxt['y1']
            x1 = max(x1, nxt['x1'])
            n += 1
            cands.append({'text': joined, 'y1': y1, 'y0': y0,
                          'x0': base['x0'], 'x1': x1})
    # per band, pick best (fieldOID, score); exact=2, contains=1; longer label
    # breaks ties
    band_hits = defaultdict(list)
    for cand in cands:
        cn = norm_text(cand['text'])
        if not cn:
            continue
        for oid, info in fields.items():
            ln = info['label_n']
            if not ln:
                continue
            score = 0
            if cn == ln:
                score = 2
            elif len(ln) >= 4 and (ln in cn or cn in ln):
                score = 1
            if score:
                band_hits[(round(cand['y1'], 1), round(cand['x0'], 1))].append(
                    (score, len(ln), oid, cand))
    bands = defaultdict(list)
    for (_, _), hits in band_hits.items():
        best = max(hits, key=lambda h: (h[0], h[1]))
        bands[best[2]].append(best[3])
    return bands


def find_option_chars(page_text, yn_field_bands):
    """Associate 是/否 chars with YN fields (nearest label band above).
    -> {fieldOID: [char dicts]}"""
    chars = page_text['chars']
    opts = defaultdict(list)
    for c in chars:
        if c['text'] not in ('是', '否'):
            continue
        best, best_d = None, 1e9
        for oid, bands in yn_field_bands.items():
            for b in bands:
                # char must be below the label band, x not left of label
                if c['y1'] >= b['y1'] - 6 or c['x0'] < b['x0'] - 10:
                    continue
                d = b['y1'] - c['y1']
                if d <= 80 and d < best_d:
                    best, best_d = oid, d
        if best is not None:
            opts[best].append(c)
    return opts


def option_bands(opts):
    """-> {fieldOID: [{'y1','y0','option','chars':[...]}]}  stacked options get
    their own band; same-line (side-by-side) options share one band."""
    out = defaultdict(list)
    for oid, chars in opts.items():
        # group by top
        chars_sorted = sorted(chars, key=lambda c: c['y1'])
        groups = []
        for c in chars_sorted:
            if groups and abs(groups[-1][0] - c['y1']) < 2:
                groups[-1][1].append(c)
            else:
                groups.append([c['y1'], [c]])
        for top, g in groups:
            g_sorted = sorted(g, key=lambda c: c['x0'])
            out[oid].append({
                'y1': max(c['y1'] for c in g), 'y0': min(c['y0'] for c in g),
                'chars': g_sorted,
            })
    return out


def assign_annots(page, annots, page_text, form, fields, norm_labels, dicts,
                  titles, global_titles, qc, aliases_used):
    """-> list of assignment dicts: {annot, kind, fieldOID, item, value,
    title}"""
    bands = find_label_bands(page_text, fields, norm_labels)

    yn_fields = {oid: info for oid, info in fields.items()
                 if info['dict'] == 'YN'}
    yn_bands = {oid: bands.get(oid, []) for oid in yn_fields}
    opts = find_option_chars(page_text, yn_bands)
    obands = option_bands(opts)

    # pass 1: domain
    domain_annots = []
    for a in annots:
        flat = a['flat']
        if DOMAIN_NAME_RE.match(flat):
            domain_annots.append(a)
            continue
        m = DOMAIN_CAT_RE.match(flat)
        if m and a['y1'] > 700:
            domain_annots.append(a)
            continue
        if re.match(r'^[A-Z]{2}$', flat) and a['y1'] > 700:
            domain_annots.append(a)

    assignments = []
    notsubmitted = []   # annots with [NOT SUBMITTED] / boilerplate
    for a in annots:
        if a in domain_annots:
            continue
        flat = a['flat']
        asg = {'annot': a, 'kind': 'qc', 'fieldOID': None, 'item': None,
               'value': None, 'title': None}

        # 1) option-band overlap (any YN field)
        if not asg['fieldOID']:
            for oid, bds in obands.items():
                for b in bds:
                    if vertical_overlap(a, b) >= 2:
                        if len(b['chars']) == 1:
                            asg['fieldOID'] = oid
                            asg['item'] = b['chars'][0]['text']
                            asg['kind'] = 'field'
                        else:
                            asg['fieldOID'] = oid
                            asg['kind'] = 'field'
                        break
                if asg['fieldOID']:
                    break

        # 2) VAR = value / VAR in SUPPXX / pure VAR
        if not asg['fieldOID']:
            supp = SUPP_RE.match(flat)
            var, value = None, None
            if supp:
                var = supp.group(1)
            else:
                m = EQ_RE.match(flat)
                if m:
                    var, value = m.group(1), m.group(2).strip()
                elif re.match(r'^[A-Z][A-Z0-9]*$', flat):
                    var = flat

            if var:
                fids = set(fields.keys())
                # title-band check first: VAR is a lab column title and the
                # Chinese title line sits in the annot band
                title_var = var
                title_line = None
                if title_var in TITLE_VARS or title_var == 'LBCLSIG':
                    tvar = title_var if title_var in TITLE_VARS else 'LBSIG'
                    tname = titles.get(form, {}).get(tvar) or \
                        global_titles.get(tvar)
                    if tname:
                        for l in page_text['lines']:
                            if norm_text(l['text']) == norm_text(tname) and \
                                    abs(l['y1'] - a['y1']) <= 8:
                                title_line = l
                                break
                if title_line is not None and value is None:
                    asg['kind'] = 'title'
                    asg['title'] = title_line['text']
                elif var in fids:
                    asg['fieldOID'] = var
                    asg['kind'] = 'field'
                elif value in fids:
                    asg['fieldOID'] = value
                    asg['kind'] = 'field'
                elif value in ALIASES and ALIASES[value] in fids:
                    asg['fieldOID'] = ALIASES[value]
                    asg['kind'] = 'field'
                    aliases_used.append((page, flat, ALIASES[value]))
                elif form in SPECIAL:
                    spec = SPECIAL[form].get(var)
                    if isinstance(spec, dict):
                        tgt = spec.get(value)
                        if tgt in fids:
                            asg['fieldOID'] = tgt
                            asg['kind'] = 'field'
                    elif isinstance(spec, str) and spec == '__GRID__':
                        asg['kind'] = 'grid_pending'
                    elif isinstance(spec, str) and spec in fids:
                        asg['fieldOID'] = spec
                        asg['kind'] = 'field'
                if not asg['fieldOID'] and value:
                    for pat, tpl in CODE_PATTERNS:
                        m = pat.match(value)
                        if m:
                            n = m.group(1) if m.groups() else ''
                            tgt = tpl + n
                            if tgt in fids:
                                asg['fieldOID'] = tgt
                                asg['kind'] = 'field'
                            break

            if not asg['fieldOID'] and not asg['title'] and asg['kind'] != 'grid_pending':
                if flat == '[NOT SUBMITTED]':
                    asg['kind'] = 'notsubmitted'
                    notsubmitted.append(asg)
                elif re.match(r'If this box', flat, re.I):
                    asg['kind'] = 'boilerplate'

        assignments.append(asg)

    # pass 2: grid rows (same band as a TESTCD anchor) + geometry + NOT SUBMITTED
    anchors = {}
    for asg in assignments:
        if asg['fieldOID'] and asg['kind'] == 'field':
            m = EQ_RE.match(asg['annot']['flat'])
            if m and TESTCD_RE.match(m.group(1)):
                anchors[id(asg['annot'])] = (asg['annot'], asg['fieldOID'])
    anchor_annots = [v[0] for v in anchors.values()]

    for asg in assignments:
        if asg['fieldOID'] or asg['title']:
            continue
        a = asg['annot']
        if asg['kind'] == 'grid_pending' or (asg['kind'] == 'qc' and
                                             not asg['fieldOID']):
            hit = None
            for t in anchor_annots:
                if abs(t['y1'] - a['y1']) <= 6:
                    hit = anchors[id(t)][1]
                    break
            if hit:
                asg['fieldOID'] = hit
                asg['kind'] = 'field'
                continue
        if asg['kind'] in ('notsubmitted', 'boilerplate', 'qc',
                           'grid_pending'):
            # geometry: overlap with a label band or nearest label above
            best, best_d = None, 1e9
            for oid, bds in bands.items():
                for b in bds:
                    ov = vertical_overlap(a, b)
                    d_above = b['y1'] - a['y1']   # label above annot
                    limit = 30 if asg['kind'] in ('notsubmitted',
                                                  'boilerplate') else 15
                    if ov >= 2:
                        d = -ov   # prefer real overlap
                        cand_d = 0
                    elif 0 <= d_above <= limit:
                        cand_d = d_above
                        d = d_above
                    else:
                        continue
                    if d < best_d:
                        best_d, best = d, oid
                    # keep first at equal distance
            if best:
                asg['fieldOID'] = best
                asg['kind'] = 'field'
                continue

    # boilerplate sentences attach to the nearest [NOT SUBMITTED] field
    for asg in assignments:
        if asg['kind'] != 'boilerplate':
            continue
        ns = [x for x in assignments if x['kind'] == 'notsubmitted'
              and x['fieldOID']]
        if ns:
            a = asg['annot']
            ref = min(ns, key=lambda x: abs(x['annot']['y1'] - a['y1']))
            asg['fieldOID'] = ref['fieldOID']
            asg['item'] = ref['item']
            asg['kind'] = 'field'

    return assignments, bands, domain_annots, obands


# ---------------------------------------------------------------------------
# C. rows / blocks
# ---------------------------------------------------------------------------

def dict_count(fieldOID, fields, dicts):
    """fields = the per-form {fieldOID: info} dict."""
    info = fields.get(fieldOID)
    if info and info['dict']:
        return len(dicts.get(info['dict'], []))
    return 1


def cjk_ascii_split(s):
    """leading CJK run -> Item_C, trailing ASCII run -> Item_E"""
    lead, tail = '', ''
    for ch in s:
        if ch not in AFM:
            lead += ch
        else:
            tail = s[s.index(ch):]
            break
    return lead, tail


def build_rows(page, form, form_name, assignments, bands, domain_annots,
               obands, fields, dicts, page_text, page_domain, qc):
    """-> ordered list of rows for one page (each row a 42-value dict)."""
    rows = []

    def new_row():
        return {h: None for h in HEADER}

    # --- domain block ------------------------------------------------------
    dom_rows = []
    for k, a in enumerate(sorted(domain_annots,
                                 key=lambda x: (-x['y1'], -x['x1'], x['idx'])), 1):
        r = new_row()
        r['annotation'] = a['text']
        r['BACKGROUND'] = a['color']
        r['X1'], r['Y1'], r['X2'], r['Y2'] = a['x1'], a['y1'], a['x2'], a['y2']
        r['DASHED_BORDER'] = a['dashed']
        r['Number'] = 1
        r['Width_annotation'] = round(text_width(a['text']), 12)
        r['Item_up'] = 195
        r['kind'] = 'domain'
        r['sort'] = (a['y1'], a['x1'], a['idx'])
        r['i'] = k
        r['annot'] = a
        dom_rows.append(r)

    # --- annot assignments -------------------------------------------------
    fld_rows = defaultdict(list)   # (fieldOID, item) -> [rows]
    for asg in assignments:
        a = asg['annot']
        r = new_row()
        r['annotation'] = a['text']
        r['BACKGROUND'] = a['color']
        r['X1'], r['Y1'], r['X2'], r['Y2'] = a['x1'], a['y1'], a['x2'], a['y2']
        r['DASHED_BORDER'] = a['dashed']
        r['Width_annotation'] = round(text_width(a['text']), 12)
        r['annot'] = a
        r['sort'] = (a['y1'], a['x1'], a['idx'])
        if asg['kind'] == 'title':
            r['kind'] = 'title'
            r['Field'] = asg['title']
            r['Number'] = 1
            r['Item_up'] = 195
            r['Highlight'] = '1'
            fld_rows[('__title__', a['y1'] // 8)].append(r)
        elif asg['kind'] == 'field':
            oid = asg['fieldOID']
            info = fields.get(oid)
            r['kind'] = 'field'
            r['FieldOID'] = oid
            r['Field'] = info['label'] if info else ''
            r['Number'] = dict_count(oid, fields, dicts)
            r['_dict'] = info['dict'] if info else None
            r['_unit'] = info['unit'] if info else ''
            if asg['item']:
                r['Item'] = asg['item']
                r['Item_All'] = asg['item']
                r['Item_C'] = asg['item']
                r['Item_E'] = ''
                r['Width_itemc'] = len(asg['item']) * 12
                r['Width_iteme'] = 0
                r['Width_item'] = r['Width_itemc']
                r['Width_item_max'] = r['Width_itemc']
                r['Item_up'] = 207
            else:
                r['Item_up'] = 195
                r['Width_itemc'] = 0
                r['Width_iteme'] = 0
                r['Width_item'] = 0
                r['Width_item_max'] = 0
            # dict-entry item: VAR=value where value is a dict entry
            if not asg['item']:
                m = EQ_RE.match(a['flat'])
                if m and info and info['dict'] and \
                        m.group(2).strip() in dicts.get(info['dict'], []):
                    val = m.group(2).strip()
                    r['Item'] = val
                    r['Item_All'] = val
                    lead, tail = cjk_ascii_split(val)
                    r['Item_C'] = lead
                    r['Item_E'] = tail
                    r['Width_itemc'] = len(lead) * 12
                    r['Width_iteme'] = round(text_width(tail), 12)
                    r['Width_item'] = r['Width_itemc'] + r['Width_iteme']
                    r['Width_item_max'] = r['Width_item']
                    r['Item_up'] = 207
            fld_rows[(oid, asg['item'])].append(r)
        else:
            # qc / unresolved
            r['kind'] = 'qc'
            r['Number'] = 1
            r['Item_up'] = 195
            r['Width_itemc'] = 0
            r['Width_iteme'] = 0
            r['Width_item'] = 0
            r['Width_item_max'] = 0
            fld_rows[('__qc__', a['y1'] // 8)].append(r)
            qc['unmatched'].append(
                (page, a['flat'], (a['x1'], a['y1'], a['x2'], a['y2'])))

    # --- blocks: cluster each (field, item) group by band ------------------
    blocks = []
    for key, rs in fld_rows.items():
        rs_sorted = sorted(rs, key=lambda r: (-r['sort'][0], r['sort'][1]))
        cur = [rs_sorted[0]]
        for r in rs_sorted[1:]:
            if abs(cur[-1]['sort'][0] - r['sort'][0]) <= 6:
                cur.append(r)
            else:
                blocks.append(cur)
                cur = [r]
        blocks.append(cur)

    # --- unit rows ---------------------------------------------------------
    assigned_fields = {a['fieldOID'] for a in assignments}
    unit_rows = []
    for oid, info in fields.items():
        if not info['unit'] or not info['label'].endswith('单位'):
            continue
        if oid in assigned_fields:
            continue
        bds = bands.get(oid, [])
        if not bds:
            continue
        b = min(bds, key=lambda x: -x['y1'])
        # unit text line in the label band
        unit_line = None
        for l in page_text['lines']:
            if abs(l['y1'] - b['y1']) <= 8 and \
                    norm_text(info['unit']) in norm_text(l['text']):
                unit_line = l
                break
        if unit_line is None:
            qc['unit_missing'].append((page, oid, info['label'], info['unit']))
            continue
        r = new_row()
        r['kind'] = 'unit'
        r['FieldOID'] = oid
        r['Field'] = info['label']
        r['BACKGROUND'] = '0.75 1 1'
        r['X1'], r['Y1'], r['X2'], r['Y2'] = \
            unit_line['x0'], unit_line['y1'], unit_line['x1'], unit_line['y0']
        r['Width_annotation'] = 0
        r['Number'] = 1
        unit = info['unit']
        r['Item'] = '&_' + unit
        r['Item_All'] = '&_' + unit
        r['Item_C'] = '&_'
        r['Item_E'] = unit
        r['Width_itemc'] = 24
        r['Width_iteme'] = round(text_width(unit), 12)
        r['Width_item'] = r['Width_itemc'] + r['Width_iteme']
        r['Width_item_max'] = r['Width_item']
        r['Item_up'] = 231
        r['X1_Item'] = unit_line['x0']
        r['Y1_Item'] = unit_line['y1']
        r['Y2_Item'] = round(unit_line['y1'] + 13, 12)
        r['sort'] = (unit_line['y1'] - 0.01, 0, 0)
        r['_band'] = (unit_line['y1'], unit_line['y0'], unit_line['x0'],
                      unit_line['x1'])
        unit_rows.append(r)

    # --- 表格行 row ----------------------------------------------------------
    tbl_row = None
    first_non_dom = None
    for a in sorted(annots_of(assignments) + domain_annots,
                    key=lambda x: -x['y1']):
        if a not in domain_annots:
            first_non_dom = a
            break
    if form in TABLE_ROW_FORMS and first_non_dom is not None:
        lines = page_text['lines']
        groups = {}
        for l in lines:
            key = round(l['y1'], 1)
            groups.setdefault(key, []).append(l)
        best_g, best_y = None, None
        for key, g in groups.items():
            if len(g) < 3:
                continue
            span = max(x['x1'] for x in g) - min(x['x0'] for x in g)
            y = g[0]['y1']
            if span <= 150 or y > 780:
                continue
            if y < first_non_dom['y1'] - 1 or y > first_non_dom['y1'] + 200:
                continue
            if best_y is None or y > best_y:
                best_g, best_y = g, y
        if best_g is not None:
            left = min(best_g, key=lambda x: x['x0'])
            r = new_row()
            r['kind'] = 'tablerow'
            r['Form'] = '表格行'
            r['BACKGROUND'] = '0.75 1 1'
            r['X1'], r['Y1'], r['X2'], r['Y2'] = \
                left['x0'], left['y1'], left['x1'], left['y0']
            r['Width_annotation'] = 0
            r['Number'] = 1
            r['Item_up'] = 195
            r['Width_itemc'] = 0
            r['Width_iteme'] = 0
            r['Width_item'] = 0
            r['Width_item_max'] = 0
            r['Highlight'] = '1'
            r['Y1_Item'] = left['y1']
            r['Y2_Item'] = round(left['y1'] + 13, 12)
            r['sort'] = (left['y1'], 0, 0)
            r['_band'] = (left['y1'], left['y0'], left['x0'], left['x1'])
            tbl_row = r

    # --- order blocks --------------------------------------------------------
    def block_sort_key(block):
        r0 = block[0]
        y = r0['sort'][0]
        opt_rank = {'是': 0, '否': 1}.get(r0.get('Item'), 0)
        return (-y, opt_rank, -r0['sort'][1], r0['sort'][2])

    ordered = []
    for block in blocks:
        block.sort(key=lambda r: (-r['sort'][0], -r['sort'][1], r['sort'][2]))
        ordered.append(block)
    ordered.sort(key=block_sort_key)
    if dom_rows:
        ordered.insert(0, dom_rows)
    for ur in unit_rows:
        ordered.append([ur])
    if tbl_row is not None:
        ordered.append([tbl_row])
    ordered.sort(key=block_sort_key)

    # --- fill per-row columns ------------------------------------------------
    prev = None
    lag = None
    out_rows = []
    for block in ordered:
        n = len(block)
        # i = rank of the annotation in the PDF /Annots array order
        if block[0]['kind'] in ('field', 'title', 'qc'):
            i_map = {id(r['annot']): k for k, r in
                     enumerate(sorted(block, key=lambda r: r['annot']['idx']), 1)}
            for r in block:
                r['i'] = i_map[id(r['annot'])]
        # block-level Anno_low
        wsum = sum((r['Width_annotation'] or 0) for r in block)
        anno_low = 559.0 - (wsum + 3.12 * max(0, n - 1))
        if n == 1 and block[0]['annotation'] == '[NOT SUBMITTED]':
            anno_low += 2.0
        for i, r in enumerate(block):
            kind = r['kind']
            a = r.get('annot')
            # width carry
            if prev is not None and kind in ('unit', 'tablerow'):
                r['Width_annotation_a'] = round(
                    (prev['Width_annotation_a'] or 0) +
                    (prev['Width_annotation'] or 0) + 10, 12)
                r['Width_annotation_'] = prev['Width_annotation']
            elif kind in ('domain', 'field', 'title', 'qc'):
                if i == 0:
                    r['Width_annotation_a'] = 0
                    r['Width_annotation_'] = 0
                else:
                    r['Width_annotation_a'] = round(
                        sum((b['Width_annotation'] or 0) for b in block[:i]) +
                        10 * i, 12)
                    r['Width_annotation_'] = block[i - 1]['Width_annotation']
            # X1_Field / X2_Field
            if kind == 'tablerow':
                r['X1_Field'] = r['X2'] + 10
            elif kind == 'unit':
                r['X1_Field'] = r['X2'] + 10
            elif n == 1:
                r['X1_Field'] = r['X2'] + 10
            else:
                if i == 0:
                    r['X1_Field'] = 43
                else:
                    p = block[i - 1]
                    r['X1_Field'] = round((p['X1_Field'] or 0) +
                                          (p['Width_annotation'] or 0) + 10, 12)
                r['X2_Field'] = round((r['X1_Field'] or 0) +
                                      (r['Width_annotation'] or 0), 12)
            # Y1_Field / Y2_Field
            if kind == 'tablerow':
                r['Y1_Field'] = r['Y1']
            elif kind == 'unit':
                r['Y1_Field'] = r['Y1']
            else:
                bd = None
                if kind == 'field' and r.get('FieldOID'):
                    bd = min(bands.get(r['FieldOID'], []), key=lambda b: -b['y1']) \
                        if bands.get(r['FieldOID']) else None
                if bd is not None and abs(bd['y1'] - r['Y1']) <= 30:
                    r['Y1_Field'] = bd['y1']
                else:
                    r['Y1_Field'] = r['Y1']
            r['Y2_Field'] = round((r['Y1_Field'] or 0) + 13, 12)
            # page/form columns (set before item-coordinate lookup)
            r['Page'] = page
            r['Form'] = form_name
            if kind == 'tablerow':
                r['FormOID'] = None
            else:
                r['FormOID'] = form
                r['Domain'] = page_domain
            # item coordinates
            has_item_coords = (kind == 'field' and r.get('_dict')) or \
                kind in ('domain', 'tablerow', 'unit')
            if kind == 'unit':
                pass  # already set
            elif kind == 'tablerow':
                pass  # already set
            elif kind == 'domain':
                r['Y1_Item'] = r['Y1_Field']
                r['Y2_Item'] = r['Y2_Field']
            elif has_item_coords:
                band = None
                if r.get('FieldOID'):
                    band = find_entry_band(r, fields, dicts, obands, page_text,
                                           bands)
                if band is None and r.get('_dict') == 'YN' and \
                        r.get('FieldOID') in obands:
                    b = obands[r['FieldOID']][0]
                    band = (b['y1'], min(c['x0'] for c in b['chars']))
                if band is None and r.get('FieldOID'):
                    bds = bands.get(r['FieldOID'], [])
                    if bds:
                        b = min(bds, key=lambda x: -x['y1'])
                        band = (b['y1'], b['x0'])
                if band is None:
                    band = (r['Y1'], r['X1'])
                r['Y1_Item'] = band[0]
                r['X1_Item'] = band[1]
                r['Y2_Item'] = round(band[0] + 13, 12)
            r['Anno_low'] = round(anno_low, 12)
            # FieldOID_Lag
            r['FieldOID_Lag'] = lag
            out_rows.append(r)
            prev = r
        if block[0]['kind'] == 'tablerow':
            lag = None
        else:
            lag = block[0].get('FieldOID')
    return out_rows


def annots_of(assignments):
    return [a['annot'] for a in assignments]


def find_entry_band(r, fields, dicts, obands, page_text, bands):
    """locate the dict-entry / option text band for a dict field row.
    For YN option rows only the row's own option text is searched.
    fields = the per-form {fieldOID: info} dict."""
    info = fields.get(r['FieldOID'])
    if not info or not info['dict']:
        return None
    if r.get('Item'):
        entries = [r['Item']]
    else:
        entries = dicts.get(info['dict'], [])
    lab = r['Y1_Field']
    best, best_y = None, None
    for l in page_text['lines']:
        ln = norm_text(l['text'])
        if not ln:
            continue
        for e in entries:
            en = norm_text(e)
            if ln == en and l['y1'] <= lab + 4 and l['x0'] >= (r['X1'] - 15):
                if best_y is None or l['y1'] > best_y:
                    best_y, best = l['y1'], (l['y1'], l['x0'])
    return best


# ---------------------------------------------------------------------------
# E. main
# ---------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser()
    base = r'D:\Program_Project\aCRF_Mapping\01Document'
    ap.add_argument('--acrf', default=os.path.join(base, 'aCRF.pdf'))
    ap.add_argument('--als', default=os.path.join(base, 'als.xlsx'))
    ap.add_argument('--template', default=os.path.join(base, 'Annotations.xlsx'))
    ap.add_argument('--out', default=os.path.join(base, 'Annotations_reverse.xlsx'))
    ap.add_argument('--qc', default=os.path.join(base, 'QC_report_reverse.txt'))
    args = ap.parse_args()

    # width assertions
    assert abs(text_width('IETEST') - 42.012) < 1e-9, text_width('IETEST')
    assert abs(text_width('|') - 3.12) < 1e-9, text_width('|')
    assert abs(text_width('是') - 12.0) < 1e-9, text_width('是')

    # template header
    tpl = openpyxl.load_workbook(args.template, data_only=True)
    tpl_hdr = [c.value for c in tpl['Annotations'][1]]
    assert tpl_hdr == HEADER, 'template header mismatch'
    old_rows = list(tpl['Annotations'].iter_rows(min_row=2, values_only=True))
    old_by_page = defaultdict(list)
    for row in old_rows:
        if row[0] is not None:
            try:
                pg = int(row[0])
            except (TypeError, ValueError):
                continue
            old_by_page[pg].append(row)
    tpl.close()

    qc = {'unmatched': [], 'unit_missing': [], 'aliases': [],
          'special_used': [], 'fallback_pages': []}

    # extract
    pages_annots = extract_annots(args.acrf)
    pages_text = extract_page_text(args.acrf)
    forms, fields, norm_labels, dicts, titles, global_titles = load_als(args.als)

    total_annots = sum(len(v) for v in pages_annots.values())
    page_forms = detect_forms(pages_annots, pages_text)

    qc_lines = []
    qc_lines.append('=' * 72)
    qc_lines.append('QC REPORT - Annotations_reverse.xlsx (reverse-generated from aCRF.pdf)')
    qc_lines.append('=' * 72)
    qc_lines.append('')
    qc_lines.append('Total FreeText annotations extracted: %d' % total_annots)
    offpage = []
    for pg in pages_annots:
        for a in pages_annots[pg]:
            if a['y2'] < 0 or a['y1'] > 842:
                offpage.append((pg, a['flat']))
    qc_lines.append('Off-page annotations (skipped, no row emitted): %s'
                    % (offpage if offpage else 'none'))

    all_rows = []
    per_page_stats = []
    for pg in sorted(page_forms):
        form = page_forms[pg]
        if pg not in pages_annots:
            per_page_stats.append((pg, '-', '-', 0, 0))
            continue
        if form is None:
            per_page_stats.append((pg, '??', '??', len(pages_annots[pg]), 0))
            qc['fallback_pages'].append(pg)
            continue
        annots = [a for a in pages_annots[pg] if not (a['y2'] < 0 or a['y1'] > 842)]
        # page domain
        page_domain = None
        for a in annots:
            m = DOMAIN_NAME_RE.match(a['flat'])
            if m:
                page_domain = m.group(1)
                break
        if page_domain is None:
            for a in annots:
                m = DOMAIN_CAT_RE.match(a['flat'])
                if m:
                    page_domain = m.group(1)
                    break
        if page_domain is None:
            page_domain = DOMAIN_BY_FORM.get(form, form[:2])
        fids_form = fields.get(form)
        if fids_form is None:
            qc_lines.append('WARNING: form %s (page %d) not found in als Field sheet'
                            % (form, pg))
            per_page_stats.append((pg, form, '?', len(annots), 0))
            continue
        assignments, bands, domain_annots, obands = assign_annots(
            pg, annots, pages_text[pg], form, fids_form, norm_labels.get(form, {}),
            dicts, titles, global_titles, qc, qc['aliases'])
        rows = build_rows(pg, form, forms.get(form, {}).get('name', form),
                          assignments, bands, domain_annots, obands, fids_form,
                          dicts, pages_text[pg], page_domain, qc)
        all_rows.extend(rows)
        unmatched_n = sum(1 for r in rows if r['kind'] == 'qc')
        per_page_stats.append((pg, form, len(annots), len(rows), unmatched_n))
        # unannotated fields
        assigned = {a['fieldOID'] for a in assignments if a['fieldOID']}
        for oid in sorted(set(fids_form) - assigned):
            if not fids_form[oid]['label'].endswith('单位'):
                qc_lines.append('  [no annot] p%d %s.%s (%s)' %
                                (pg, form, oid, fids_form[oid]['label']))

    # write xlsx
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Annotations'
    ws.append(list(HEADER))
    for r in all_rows:
        ws.append([r.get(h) for h in HEADER])
    wb.save(args.out)

    annot_rows = sum(1 for r in all_rows if r.get('annotation') is not None)
    unit_n = sum(1 for r in all_rows if r['kind'] == 'unit')
    tbl_n = sum(1 for r in all_rows if r['kind'] == 'tablerow')
    title_n = sum(1 for r in all_rows if r['kind'] == 'title')
    qc_n = sum(1 for r in all_rows if r['kind'] == 'qc')

    # verification
    errors = []
    for r in all_rows:
        if r['kind'] == 'field' and r.get('FieldOID'):
            if r['FieldOID'] not in fields.get(r['FormOID'], {}):
                errors.append('p%d FieldOID %s not in form %s' %
                              (r['Page'], r['FieldOID'], r['FormOID']))
            else:
                info = fields[r['FormOID']][r['FieldOID']]
                if info['dict'] and r.get('Number') != len(dicts.get(info['dict'], [])):
                    errors.append('p%d %s Number=%s dict=%s(%d)' %
                                  (r['Page'], r['FieldOID'], r.get('Number'),
                                   info['dict'], len(dicts.get(info['dict'], []))))
    if annot_rows != 788:
        errors.append('annotation row count %d != 788' % annot_rows)

    # spot checks
    spot = []
    for pg, oid in [(4, 'SVDAT'), (7, 'IETEST'), (7, 'IETEST1'),
                    (51, 'AEOUT'), (60, None), (61, None)]:
        for r in all_rows:
            if r['Page'] == pg and (oid is None or r.get('FieldOID') == oid):
                spot.append('%s p%s: %s' % (oid or 'row', pg, {k: r[k] for k in
                            ['FieldOID', 'Field', 'Item', 'annotation', 'Number',
                             'X1_Field', 'Y1_Field', 'Item_up', 'Anno_low', 'i',
                             'Annotation_']}))
                break

    # QC report
    qc_lines.append('')
    qc_lines.append('--- summary ---')
    qc_lines.append('annotation rows: %d (expect 788)' % annot_rows)
    qc_lines.append('unit rows: %d   table-header rows: %d   title rows: %d   unmatched rows: %d'
                    % (unit_n, tbl_n, title_n, qc_n))
    qc_lines.append('total rows: %d' % len(all_rows))
    qc_lines.append('')
    qc_lines.append('--- per page (page, form, annots, rows, unmatched) ---')
    for row in per_page_stats:
        qc_lines.append('  p%-3s %-7s annots=%-3s rows=%-3s unmatched=%s'
                        % (row[0], row[1] or '-', row[2], row[3], row[4]))
    qc_lines.append('')
    qc_lines.append('--- unmatched annotations (page, text, rect) ---')
    for pg, txt, rect in qc['unmatched']:
        qc_lines.append('  p%d %r rect=%s' % (pg, txt, tuple(round(x, 1) for x in rect)))
    qc_lines.append('')
    qc_lines.append('--- aliases used ---')
    for pg, txt, tgt in qc['aliases']:
        qc_lines.append('  p%d %r -> %s' % (pg, txt, tgt))
    qc_lines.append('')
    qc_lines.append('--- unit rows without unit text on page ---')
    for item in qc['unit_missing']:
        qc_lines.append('  p%d %s (%s) unit=%s' % item)
    qc_lines.append('')
    qc_lines.append('--- old vs new row counts ---')
    for row in per_page_stats:
        pg = row[0]
        old_n = len(old_by_page.get(pg, []))
        qc_lines.append('  p%-3s old=%-3s new=%-3s' % (pg, old_n, row[3]))
    qc_lines.append('')
    qc_lines.append('--- verification errors ---')
    qc_lines.extend('  ' + e for e in errors) if errors else qc_lines.append('  none')
    qc_lines.append('')
    qc_lines.append('--- spot checks ---')
    for s in spot:
        qc_lines.append('  ' + s)
    qc_lines.append('')
    qc_lines.append('--- inference rules / documented deviations from the old file ---')
    for line in INFERENCE_NOTES:
        qc_lines.append('  ' + line)

    with open(args.qc, 'w', encoding='utf-8') as f:
        f.write('\n'.join(qc_lines) + '\n')

    print('rows: %d (annot %d, unit %d, table %d, title %d, unmatched %d)'
          % (len(all_rows), annot_rows, unit_n, tbl_n, title_n, qc_n))
    print('verification errors: %d' % len(errors))
    for e in errors[:20]:
        print('  !!', e)
    print('QC written to', args.qc)


INFERENCE_NOTES = [
    '1. X1_Field: single-annotation block -> X2+10; multi-annotation block -> first row 43, following rows prev X1_Field + prev Width_annotation + 10; unit/table rows -> X2+10.',
    '2. Y1_Field/Y2_Field: matched label band (within 30pt of the annotation) else the annotation band; Y2 = Y1 + 13.',
    '3. Item columns: only YN option rows (Item=yes/no, Width_itemc=len*12), dict-entry rows (Item=entry text) and unit rows (Item="&_"+unit, Item_up=231). Item_up = 195 otherwise, 207 for YN/dict-entry rows.',
    '4. X1_Item/Y1_Item/Y2_Item: present only for dict-field rows and non-field rows (domain/table/unit); band = dict entry text on page, else label band, else annotation band; Y2 = Y1 + 13.',
    '5. Number = count of DataDictionaryEntry rows for the field dataDictionaryOID (1 for non-dict and non-field rows). Verified against the old file (IETEST=5, IETEST1=33, IECAT=2, AESEV=3, AEREL=5, AEACN=4, AEOUT=6).',
    '6. Anno_low = 559 - (sum of block Width_annotation + 3.12*(n-1)); a single-[NOT SUBMITTED] block gets +2.0 (old-tool quirk, reproduced). No-annotation rows: 559.',
    '7. Width_annotation uses the standard Helvetica AFM table at 12pt (IETEST=42.012, "|"=3.12, CJK=12.0). The old tool used a slightly wider near-AFM table (IETEST=42.584), so width-derived values differ by <2%.',
    '8. i = position of the annotation in the PDF /Annots array order (creation order); the old file i-order came from the OLD PDF array and cannot be reproduced for re-drawn annotations. Row order within a band = y desc, x desc.',
    '9. Domain block: all domain annotations of the page (name regex or *CAT = ... with y1>700) form one block; each annotation keeps its own real /Rect (the old file stacked all block rows on one synthetic rect).',
    '10. YN option rows: an annotation whose band overlaps one option band (>=2pt) is attributed to that field with Item=option; overlapping a shared side-by-side band -> field-level row with empty Item.',
    '11. Grid rows: annotations sharing a band (<=6pt) with a TESTCD = code anchor inherit the anchor field (e.g. EGORRES/EGTEST on the HR row -> HR; IEORRES = N/Y -> the IETEST/IETEST1 row).',
    '12. VAR = value: VAR in the form field set wins; else the value (VSTESTCD = HEIGHT -> HEIGHT), else the alias table, else the per-form special table, else the code patterns INCLxx/EXCLxx/SCINTxx/ADINTxx.',
    '13. --DTC style annotations (SVSTDTC/AESTDTC/EGDTC/LBDTC) are pure VARS not present in als and are matched geometrically to their label band (SVDAT/AESTDAT or AESTTIM/EGDAT/LBDAT).',
    '14. LabTableTitle rows: an annotation whose text is one of the 9 column VARS (LBORRES/LBORRESU/...) with the matching Chinese column title line in its band -> Field=title, FieldOID empty, Highlight=1. The old file instead emitted one synthetic 单位 row per lab page; not reproduced.',
    '15. 表格行 rows: forms VS2/VS1/PC/AE/CM only; a group of >=3 same-top text lines spanning >150pt between the domain annotations and the first field annotation -> one row (Form=表格行, FormOID empty).',
    '16. Unit rows: fields whose label ends with 单位, have an als unit value, and carry no annotation on the page -> one synthetic row at the unit text band (X1..Y2 = unit text bbox, X1_Field = X2+10, Item="&_"+unit, Item_up=231). Width_itemc uses len*12 (old-tool convention); Width_iteme uses AFM (old tool used 12 flat for "ms" -> 14.916 now).',
    '17. FieldOID_Lag = FieldOID of the previous block (all rows of a block share it; empty after a 表格行 block). The old tool had a self-lag bug on AESEV/AEREL/AEACN/AEREL1/AEACN1 (FieldOID_Lag = own FieldOID); not reproduced.',
    '18. Highlight = 1 only for table-header and LabTableTitle rows; the old file set it inconsistently. DASHED_BORDER = 1 iff the annotation /BS contains /D. BACKGROUND = annotation /C (0.75 1 1 or 1 1 0.58824 for the yellow DS/SU/EX groups); synthetic rows get 0.75 1 1.',
    '19. Domain column (never filled in the old file): page domain from the domain annotation, else the *CAT prefix, else the formOID table (LB*/VS*/QS1/PCU/DSSCR/DSEOT handled).',
    '20. Pages 8 and 45 carry no annotations and emit no rows. Pages 60/61 (LBOTH/PCU) are new relative to the old file. The p15 "SUSMOKN in SUPPSU" annotation lies off the page (y2<0) and is skipped.',
    '21. Rows with annotations that could not be mapped keep an empty FieldOID/Field and are listed in the unmatched section of this report.',
]

if __name__ == '__main__':
    main()
