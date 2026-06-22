"""
PIPELINE DE RECOMENDACIONES — Módulo independiente
====================================================
Extrae, clasifica y exporta las recomendaciones de guías clínicas en PDF.
Compatible con formatos ACC/AHA (2014+) y ESC (2015+).

No modifica ni depende de pipeline.py.

ESTRATEGIAS DE EXTRACCIÓN:
  A  — Tablas (pdfplumber.extract_tables con múltiples configuraciones):
       Detecta tablas con columnas Recomendación / COR|Clase / LOE|Level.
       Prueba 3 estrategias: lines/lines → text/lines → text/text.
  B  — Fin de línea (formato ACC/AHA y ESC):
       Busca líneas que terminan en CLASS LOE [refs], que es el formato
       dominante en tablas sin rejilla explícita.
  C  — Patrones inline (fallback general):
       "(Class I; Level of Evidence: A)" y variantes.
"""

import re
import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ─────────────────────────────────────────────
# NORMALIZACIÓN
# ─────────────────────────────────────────────

def _normalize_class(raw: str) -> str:
    """
    Normaliza clase de recomendación → I / IIa / IIb / III.

    Acepta variantes ACC/AHA:
      "III: Harm", "III: No Benefit", "III: Potentially Harmful" → III
      "IIb: May be Considered" → IIb
    Acepta prefijos:  CLASS / COR / CLASE / CLASSE
    Acepta numerales: 1 / 2A / 2B / 3
    """
    t = str(raw).strip()
    # Extraer la parte antes de ":" si hay descripción adicional
    t = re.split(r'[:\n]', t)[0].strip()
    t = re.sub(r'\s+', '', t.upper())
    t = re.sub(r'^(CLASS|COR|CLASE|CLASSE|KLASSE)[:\s]*', '', t)
    if re.match(r'^III', t): return 'III'
    if re.match(r'^IIB', t): return 'IIb'
    if re.match(r'^IIA', t): return 'IIa'
    if t in ('II',):         return 'IIa'
    if re.match(r'^I$', t):  return 'I'
    if t in ('1',):          return 'I'
    if t in ('2A',):         return 'IIa'
    if t in ('2B',):         return 'IIb'
    if t in ('3',):          return 'III'
    return raw.strip() or 'Desconocida'


def _normalize_loe(raw: str) -> str:
    """
    Normaliza nivel de evidencia → A / B / C.
    Acepta: B-R, B-NR → B   |   C-LD, C-EO → C   |   N/A → N/A
    """
    t = str(raw).strip().upper()
    if t in ('N/A', 'NA', ''):
        return 'N/A'
    t = re.sub(r'^(LEVEL\s+OF\s+EVIDENCE|LEVEL|LOE|EVIDENCE)[:\s]*', '', t).strip()
    if t.startswith('A'): return 'A'
    if t.startswith('B'): return 'B'
    if t.startswith('C'): return 'C'
    return raw.strip() or 'Desconocida'


def _extract_cited_refs(text: str) -> str:
    """
    Extrae números de referencias citadas.
    Detecta: superíndices Unicode, [1,2,3], (1-3), superíndices pegados a texto.
    """
    SUPERSCRIPT = str.maketrans('⁰¹²³⁴⁵⁶⁷⁸⁹', '0123456789')
    t = text.translate(SUPERSCRIPT)
    numbers: set[int] = set()

    # [1,2,3] o (3-5) o (21,64,67-71)
    for m in re.finditer(r'[\[\(](\d[\d,\s\-–]+\d|\d)[\]\)]', t):
        for part in re.split(r'[,\s]+', m.group(1)):
            part = part.strip()
            rng = re.match(r'^(\d+)[–\-](\d+)$', part)
            if rng:
                try:
                    numbers.update(range(int(rng.group(1)), int(rng.group(2)) + 1))
                except ValueError:
                    pass
            elif part.isdigit():
                numbers.add(int(part))

    # Superíndices pegados: "aspirin5,6" o "hs-cTn.3,10–13"
    for m in re.finditer(r'(?<=[a-zA-Z\.\)])(\d{1,3}(?:[,–\-]\d{1,3})*)', t):
        for part in re.split(r'[,–\-]', m.group(1)):
            if part.strip().isdigit():
                numbers.add(int(part.strip()))

    if not numbers:
        return ''
    return ', '.join(str(n) for n in sorted(numbers))


# ─────────────────────────────────────────────
# PATRONES DE DETECCIÓN DE COLUMNAS
# ─────────────────────────────────────────────

# Cabeceras de columna — permisivos para cubrir variantes con superíndice
# pdfplumber puede devolver "Classa" (superíndice pegado) o "Class a" o "Class^a"
_REC_COL_RE = re.compile(
    r'recommendation|recomendaci[oó]n|recommandation|raccomandazione', re.I
)
_CLASS_COL_RE = re.compile(
    r'^(class[a-z^\s]*|cor|clase|classe|crc|recommendation\s+class)$', re.I
)
_LOE_COL_RE = re.compile(
    r'^(level[a-z^\s]*|loe|evidence|loe\s*\.|niveau|livello|beweis)$', re.I
)

# Valores válidos en celdas de clase (permisivo: admite "III: Harm" etc.)
_CLASS_VAL_RE = re.compile(
    r'^(iii(?:[:\s].+)?|iia|iib|ii(?![aibIAIB])|i(?![IiAaBbLl]))$', re.I
)
_LOE_VAL_RE = re.compile(r'^[abc](?:-[a-z]+)?$', re.I)

# Patrones textuales inline (Estrategia C)
_INLINE_PATTERNS = [
    # "(Level of Evidence: A) I"
    re.compile(
        r'(?P<text>[A-Z][^(]{15,500}?)'
        r'\(Level\s+of\s+Evidence\s*[:\s]+(?P<loe>[ABC](?:-[A-Z]+)?)\)'
        r'\s*(?P<cls>I{1,3}|IIa|IIb|III)\b',
        re.I | re.DOTALL
    ),
    # "(Class I; Level of Evidence: A)"
    re.compile(
        r'(?P<text>[A-Z][^(]{15,500}?)'
        r'\((?:Class\s+)?(?P<cls>I{1,3}|IIa|IIb|III)'
        r'\s*[;,/]\s*(?:Level\s+of\s+Evidence[:\s]+|LOE\s*[:\s]*)?'
        r'(?P<loe>[ABC](?:-[A-Z]+)?)\)',
        re.I | re.DOTALL
    ),
]

# Patrón de fin de línea — formato ACC/AHA y ESC
# Las tablas sin rejilla producen texto donde cada recomendación termina en:
# "...texto... I  A  (21)" o "...texto... IIa  B  (42-44,75-81)" o "...texto... III: Harm  B"
_EOL_PAT = re.compile(
    r'^(?P<lead>.+?)\s+'
    r'(?P<cls>III(?:[:\s]+\S+(?:\s+\S+)?)?|IIa|IIb|I)\s+'
    r'(?P<loe>[ABC])\b'
    r'(?:\s+(?:N/?A|\([\d,\s\-–]+\)|\d[\d,\-–,\s]*))?'
    r'\s*$',
    re.IGNORECASE
)


# ─────────────────────────────────────────────
# EXTRACCIÓN POR TABLAS (ESTRATEGIA A)
# ─────────────────────────────────────────────

_TABLE_SETTINGS = [
    # Estrategia 1: líneas explícitas (tablas con rejilla visible)
    {"vertical_strategy": "lines", "horizontal_strategy": "lines"},
    # Estrategia 2: columnas por texto, filas por líneas (ACC/AHA sin rejilla vertical)
    {"vertical_strategy": "text",  "horizontal_strategy": "lines",
     "intersection_tolerance": 10, "snap_tolerance": 3},
    # Estrategia 3: todo por texto (ESC y otros sin rejilla)
    {"vertical_strategy": "text",  "horizontal_strategy": "text",
     "intersection_tolerance": 10, "snap_tolerance": 3},
]


def _try_extract_tables(page) -> list:
    """Prueba varias configuraciones de pdfplumber y devuelve la primera con tablas."""
    for settings in _TABLE_SETTINGS:
        try:
            tables = page.extract_tables(table_settings=settings) or []
            if tables:
                return tables
        except Exception:
            continue
    return []


def _parse_table(table, page_num: int) -> list[dict]:
    """
    Analiza una tabla pdfplumber e intenta extraer recomendaciones.
    Devuelve lista vacía si la tabla no parece de recomendaciones.
    """
    if not table or len(table) < 2:
        return []

    header = [(str(c or '')).replace('\n', ' ').strip() for c in (table[0] or [])]

    rec_ci = class_ci = loe_ci = -1

    for ci, h in enumerate(header):
        h_clean = re.sub(r'\s+', ' ', h).strip()
        if _REC_COL_RE.search(h_clean)  and rec_ci   < 0: rec_ci   = ci
        if _CLASS_COL_RE.match(h_clean) and class_ci < 0: class_ci = ci
        if _LOE_COL_RE.match(h_clean)   and loe_ci   < 0: loe_ci   = ci

    # Fallback: detectar columnas por valores de celdas
    if class_ci < 0:
        class_votes: dict[int, int] = {}
        loe_votes:   dict[int, int] = {}
        for row in table[1:]:
            for ci, cell in enumerate(row or []):
                v = str(cell or '').strip().split('\n')[0]   # primera línea de celda
                if _CLASS_VAL_RE.match(v): class_votes[ci] = class_votes.get(ci, 0) + 1
                if _LOE_VAL_RE.match(v):   loe_votes[ci]   = loe_votes.get(ci, 0) + 1
        if class_votes: class_ci = max(class_votes, key=class_votes.get)
        if loe_votes:   loe_ci   = max(loe_votes,   key=loe_votes.get)

    if class_ci < 0:
        return []   # no parece tabla de recomendaciones

    # Columna de texto: la más ancha excluidas clase y LOE
    if rec_ci < 0:
        lengths: dict[int, int] = {}
        for row in table[1:]:
            for ci, cell in enumerate(row or []):
                if ci not in (class_ci, loe_ci):
                    lengths[ci] = lengths.get(ci, 0) + len(str(cell or ''))
        if lengths:
            rec_ci = max(lengths, key=lengths.get)

    if rec_ci < 0:
        return []

    recs = []
    for row in table[1:]:
        def cv(idx: int) -> str:
            if idx < 0 or not row or idx >= len(row):
                return ''
            return str(row[idx] or '').replace('\n', ' ').strip()

        text = cv(rec_ci)
        cls  = cv(class_ci)
        loe  = cv(loe_ci) if loe_ci >= 0 else ''

        # Ignorar filas demasiado cortas (cabeceras de sección como "Oxygen", "Nitrates")
        if len(text) < 25:
            continue
        # Ignorar filas sin clase válida (también son sub-cabeceras)
        cls_norm = _normalize_class(cls)
        if cls_norm == 'Desconocida' and not cls:
            continue

        recs.append({
            'text':             text,
            'rec_class':        cls_norm,
            'loe':              _normalize_loe(loe),
            'references_cited': _extract_cited_refs(text),
            'page':             page_num,
            'source':           'tabla',
        })

    return recs


# ─────────────────────────────────────────────
# EXTRACCIÓN POR FIN DE LÍNEA (ESTRATEGIA B)
# ─────────────────────────────────────────────

# Líneas a ignorar en la acumulación de texto
_NOISE_LINE_RE = re.compile(
    r'^(TABLE|FIGURE|Recommendations?|COR|LOE|References?|'
    r'ACS\s+indicates|CCB\s+indicates|LOE\s+indicates|'
    r'\*See\s+Section|\*Short|N/A,\s+not|\d{4}\s+(ACC|AHA|ESC))',
    re.I
)


def _extract_eol(page_text: str, page_num: int) -> list[dict]:
    """
    Extrae recomendaciones del texto de una página buscando el patrón
    'texto largo ... CLASS  LOE  [refs]' al final de cada línea.

    Acumula líneas intermedias para recomendaciones multilinea.
    """
    recs = []
    lines = page_text.split('\n')
    buffer: list[str] = []

    for line in lines:
        stripped = line.strip()
        m = _EOL_PAT.match(stripped)

        if m:
            lead = m.group('lead').strip()
            cls  = m.group('cls').strip()
            loe  = m.group('loe').strip()

            # Texto completo = buffer acumulado + parte inicial de esta línea
            full_parts = buffer + ([lead] if lead else [])
            full_text  = re.sub(r'\s+', ' ', ' '.join(full_parts)).strip()

            if len(full_text) >= 25:
                recs.append({
                    'text':             full_text,
                    'rec_class':        _normalize_class(cls),
                    'loe':              _normalize_loe(loe),
                    'references_cited': _extract_cited_refs(stripped),
                    'page':             page_num,
                    'source':           'texto',
                })
            buffer = []

        else:
            # Acumular si parece texto de recomendación (no ruido)
            if (stripped
                    and len(stripped) >= 15
                    and not _NOISE_LINE_RE.match(stripped)):
                buffer.append(stripped)
            elif not stripped or len(stripped) < 5:
                buffer = []   # línea vacía → reset

    return recs


# ─────────────────────────────────────────────
# EXTRACCIÓN INLINE (ESTRATEGIA C)
# ─────────────────────────────────────────────

def _extract_inline(page_text: str, page_num: int) -> list[dict]:
    """Busca patrones (Class I; Level of Evidence: A) en texto corrido."""
    recs = []
    for pat in _INLINE_PATTERNS:
        for m in pat.finditer(page_text):
            raw_text = m.group('text').strip()
            if len(raw_text) > 400:
                last_stop = raw_text.rfind('. ', len(raw_text) - 380)
                if last_stop > 50:
                    raw_text = raw_text[last_stop + 2:].strip()
            if len(raw_text) < 25:
                continue
            recs.append({
                'text':             raw_text,
                'rec_class':        _normalize_class(m.group('cls')),
                'loe':              _normalize_loe(m.group('loe')),
                'references_cited': _extract_cited_refs(m.group('text')),
                'page':             page_num,
                'source':           'texto',
            })
    return recs


# ─────────────────────────────────────────────
# FUNCIÓN PRINCIPAL
# ─────────────────────────────────────────────

def extract_recommendations_from_pdf(pdf_path: str, debug: bool = False) -> list[dict]:
    """
    Extrae recomendaciones de una guía clínica en PDF.

    Orden de estrategias (se detiene al obtener ≥ 5 resultados):
      A — Tablas pdfplumber (3 configuraciones)
      B — Fin de línea: 'texto CLASS LOE [refs]' (ACC/AHA y ESC sin rejilla)
      C — Patrones inline: '(Class I; Level of Evidence: A)' y variantes

    Returns:
        Lista de dicts con: rec_number, text, rec_class, loe,
        references_cited, page (1-indexed), source.
    """
    log = []
    def dbg(msg):
        log.append(msg)
        if debug:
            print(msg)

    all_table_recs: list[dict] = []
    all_eol_recs:   list[dict] = []
    all_inline_recs: list[dict] = []

    with pdfplumber.open(pdf_path) as pdf:
        dbg(f"[REC] PDF: {len(pdf.pages)} páginas")

        for page_idx, page in enumerate(pdf.pages):
            page_num = page_idx + 1

            # ── A: tablas ────────────────────────────────────────────
            tables = _try_extract_tables(page)
            for table in tables:
                recs = _parse_table(table, page_num)
                all_table_recs.extend(recs)
                if recs:
                    dbg(f"[REC-A] pág {page_num}: {len(recs)} recs en tabla")

            # ── B y C: texto ─────────────────────────────────────────
            try:
                page_text = page.extract_text() or ''
            except Exception:
                page_text = ''

            if page_text:
                eol_recs = _extract_eol(page_text, page_num)
                all_eol_recs.extend(eol_recs)
                if eol_recs:
                    dbg(f"[REC-B] pág {page_num}: {len(eol_recs)} recs fin-de-línea")

                inline_recs = _extract_inline(page_text, page_num)
                all_inline_recs.extend(inline_recs)
                if inline_recs:
                    dbg(f"[REC-C] pág {page_num}: {len(inline_recs)} recs inline")

    dbg(f"[REC] A(tablas)={len(all_table_recs)} "
        f"B(eol)={len(all_eol_recs)} "
        f"C(inline)={len(all_inline_recs)}")

    # Elegir la estrategia con más resultados (≥ 5 para ser fiable)
    candidates = [
        (all_table_recs,  'A-tablas'),
        (all_eol_recs,    'B-eol'),
        (all_inline_recs, 'C-inline'),
    ]
    best = max(candidates, key=lambda x: len(x[0]))
    recommendations, strategy = best

    if len(recommendations) < 5:
        # Ninguna estrategia fue convincente: combinar B + C
        combined = all_eol_recs + all_inline_recs
        if len(combined) > len(recommendations):
            recommendations = combined
            strategy = 'B+C combinadas'

    dbg(f"[REC] Estrategia elegida: {strategy} → {len(recommendations)} recomendaciones")

    # _Deduplicar_ por texto (puede haber solapamiento entre estrategias)
    seen: set[str] = set()
    unique: list[dict] = []
    for r in recommendations:
        key = r['text'][:80].lower()
        if key not in seen:
            seen.add(key)
            unique.append(r)

    # Renumeración secuencial final
    for i, r in enumerate(unique, 1):
        r['rec_number'] = i

    dbg(f"[REC] Total final (sin duplicados): {len(unique)}")
    if not debug:
        for line in log:
            print(line)

    return unique


# ─────────────────────────────────────────────
# EXPORTACIÓN A EXCEL
# ─────────────────────────────────────────────

_CLASS_COLORS = {'I': 'E2EFDA', 'IIa': 'FFF2CC', 'IIb': 'FCE4D6', 'III': 'F8CECC'}
_LOE_COLORS   = {'A': 'DAE8FC', 'B': 'E1D5E7', 'C': 'F5F5F5'}
_HEADER_BG    = '1F4E79'


def _style_header(cell, bg: str = _HEADER_BG, fg: str = 'FFFFFF'):
    cell.font      = Font(bold=True, color=fg, name='Arial', size=10)
    cell.fill      = PatternFill('solid', start_color=bg)
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)


def _thin_border(ws, row: int, n_cols: int):
    thin = Side(style='thin', color='CCCCCC')
    for col in range(1, n_cols + 1):
        ws.cell(row=row, column=col).border = Border(bottom=thin)


def _write_rec_sheet(ws, recommendations: list[dict]):
    ws.title = 'Recomendaciones'
    cols = [
        ('Nº',                        5),
        ('Texto de la recomendación', 75),
        ('Clase',                     10),
        ('LOE',                       12),
        ('Referencias citadas',        25),
        ('Página',                      8),
        ('Fuente',                     10),
        ('Clase (manual)',             12),
        ('LOE (manual)',               12),
        ('Notas',                      30),
    ]
    ws.row_dimensions[1].height = 30
    for ci, (name, width) in enumerate(cols, 1):
        c = ws.cell(row=1, column=ci, value=name)
        _style_header(c)
        ws.column_dimensions[get_column_letter(ci)].width = width
    ws.freeze_panes = 'A2'

    for offset, rec in enumerate(recommendations):
        rn  = offset + 2
        cls = rec.get('rec_class', '')
        loe = rec.get('loe', '')
        values = [
            rec.get('rec_number', ''), rec.get('text', ''),
            cls, loe,
            rec.get('references_cited', ''), rec.get('page', ''),
            rec.get('source', ''), '', '', '',
        ]
        for ci, val in enumerate(values, 1):
            cell = ws.cell(row=rn, column=ci, value=val)
            cell.font      = Font(name='Arial', size=9)
            cell.alignment = Alignment(vertical='top', wrap_text=(ci == 2))
            if ci == 2:
                cell.fill = PatternFill('solid', start_color=_CLASS_COLORS.get(cls, 'F5F5F5'))
            elif ci == 3:
                cell.fill      = PatternFill('solid', start_color=_CLASS_COLORS.get(val, 'F5F5F5'))
                cell.alignment = Alignment(horizontal='center', vertical='top')
                cell.font      = Font(name='Arial', size=9, bold=True)
            elif ci == 4:
                cell.fill      = PatternFill('solid', start_color=_LOE_COLORS.get(val, 'F5F5F5'))
                cell.alignment = Alignment(horizontal='center', vertical='top')
                cell.font      = Font(name='Arial', size=9, bold=True)
        ws.row_dimensions[rn].height = 55
        _thin_border(ws, rn, len(cols))


def _write_summary_sheet(ws, recommendations: list[dict]):
    ws.title = 'Resumen recomendaciones'
    CLASS_ORDER = ['I', 'IIa', 'IIb', 'III', 'Desconocida']
    LOE_ORDER   = ['A', 'B', 'C', 'N/A', 'Desconocida']

    matrix:       dict[str, dict[str, int]] = {}
    class_totals: dict[str, int]            = {}
    loe_totals:   dict[str, int]            = {}
    for rec in recommendations:
        cls = rec.get('rec_class', 'Desconocida')
        loe = rec.get('loe',       'Desconocida')
        matrix.setdefault(cls, {})
        matrix[cls][loe]  = matrix[cls].get(loe, 0) + 1
        class_totals[cls] = class_totals.get(cls, 0) + 1
        loe_totals[loe]   = loe_totals.get(loe, 0)  + 1

    present_cls = [c for c in CLASS_ORDER if c in class_totals]
    present_loe = [l for l in LOE_ORDER   if l in loe_totals]
    ws.column_dimensions['A'].width = 20

    _style_header(ws.cell(row=1, column=1, value='Clase / LOE'))
    for ci, loe in enumerate(present_loe, 2):
        _style_header(ws.cell(row=1, column=ci, value=loe))
        ws.column_dimensions[get_column_letter(ci)].width = 10
    total_col = len(present_loe) + 2
    _style_header(ws.cell(row=1, column=total_col, value='TOTAL'))
    ws.column_dimensions[get_column_letter(total_col)].width = 10

    for ri, cls in enumerate(present_cls, 2):
        c = ws.cell(row=ri, column=1, value=cls)
        c.font = Font(bold=True, name='Arial', size=10)
        c.fill = PatternFill('solid', start_color=_CLASS_COLORS.get(cls, 'F5F5F5'))
        for ci, loe in enumerate(present_loe, 2):
            val = matrix.get(cls, {}).get(loe, 0)
            cell = ws.cell(row=ri, column=ci, value=val if val else '')
            cell.font      = Font(name='Arial', size=10)
            cell.alignment = Alignment(horizontal='center')
        t = ws.cell(row=ri, column=total_col, value=class_totals[cls])
        t.font = Font(bold=True, name='Arial', size=10)
        t.alignment = Alignment(horizontal='center')

    tr = len(present_cls) + 2
    ws.cell(row=tr, column=1, value='TOTAL').font = Font(bold=True, name='Arial', size=10)
    for ci, loe in enumerate(present_loe, 2):
        c = ws.cell(row=tr, column=ci, value=loe_totals.get(loe, 0))
        c.font = Font(bold=True, name='Arial', size=10)
        c.alignment = Alignment(horizontal='center')
    g = ws.cell(row=tr, column=total_col, value=len(recommendations))
    g.font = Font(bold=True, name='Arial', size=10)
    g.alignment = Alignment(horizontal='center')

    leyenda = [
        ('CLASE I',   'Beneficio >> riesgo. Indicado / recomendado.',         'E2EFDA'),
        ('CLASE IIa', 'Beneficio > riesgo. Es razonable realizarlo.',          'FFF2CC'),
        ('CLASE IIb', 'Beneficio ≥ riesgo. Puede considerarse.',               'FCE4D6'),
        ('CLASE III', 'Sin beneficio o dañino. No está recomendado.',          'F8CECC'),
        ('LOE A',     'Múltiples ECAs o meta-análisis de alta calidad.',       'DAE8FC'),
        ('LOE B',     'Un ECA o estudios no aleatorizados de gran tamaño.',    'E1D5E7'),
        ('LOE C',     'Consenso de expertos, estudios pequeños o registros.',  'F5F5F5'),
    ]
    for offset, (label, desc, color) in enumerate(leyenda, tr + 2):
        c = ws.cell(row=offset, column=1, value=label)
        c.font = Font(name='Arial', size=9)
        c.fill = PatternFill('solid', start_color=color)
        ws.cell(row=offset, column=2, value=desc).font = Font(name='Arial', size=9)
    ws.column_dimensions['B'].width = 55


def export_recommendations_to_excel(recommendations: list[dict], output_path: str):
    """Exporta recomendaciones a Excel (hojas: Recomendaciones + Resumen)."""
    wb = Workbook()
    _write_rec_sheet(wb.active, recommendations)
    _write_summary_sheet(wb.create_sheet(), recommendations)
    wb.save(output_path)
    print(f'[OK] Excel guardado: {output_path}')
