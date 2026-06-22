"""
PIPELINE DE RECOMENDACIONES — Módulo independiente
====================================================
Extrae, clasifica y exporta las recomendaciones de guías clínicas en PDF.

No modifica ni depende del pipeline de referencias (pipeline.py).
Se integra en app.py con un único import adicional.

FUNCIONES PÚBLICAS:
  extract_recommendations_from_pdf(pdf_path)  → list[dict]
  export_recommendations_to_excel(recs, path) → None

CAMPOS DE CADA RECOMENDACIÓN:
  rec_number        int   — número secuencial
  text              str   — texto de la recomendación
  rec_class         str   — I / IIa / IIb / III / Desconocida
  loe               str   — A / B / C / Desconocida
  references_cited  str   — números de refs citadas (p.ej. "1, 3, 7")
  page              int   — página del PDF (1-indexed)
  source            str   — 'tabla' o 'texto'

ESTRATEGIAS DE EXTRACCIÓN:
  A — Por tablas   : pdfplumber.extract_tables(), detecta col Rec/Clase/LOE
  B — Por texto    : patrones regex de clase+LOE en texto corrido (fallback)
"""

import re
import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ─────────────────────────────────────────────
# HELPERS DE NORMALIZACIÓN
# ─────────────────────────────────────────────

def _normalize_class(raw: str) -> str:
    """
    Normaliza cualquier variante de clase de recomendación → I / IIa / IIb / III.

    Acepta:
      - Texto: "Class I", "COR IIa", "class iii", "IIB"
      - Numerales arábigos (algunos documentos): 1, 2A, 2B, 3
      - Prefijos en varios idiomas: CLASS / COR / CLASE / CLASSE
    """
    t = re.sub(r'\s+', '', str(raw).strip().upper())
    t = re.sub(r'^(CLASS|COR|CLASE|CLASSE|KLASSE)[:\s]*', '', t)
    if re.match(r'^III', t): return 'III'
    if re.match(r'^IIB', t): return 'IIb'
    if re.match(r'^IIA', t): return 'IIa'
    if t in ('II',):         return 'IIa'   # sin especificar → IIa (conservador)
    if re.match(r'^I$', t):  return 'I'
    if t in ('1',):          return 'I'
    if t in ('2A',):         return 'IIa'
    if t in ('2B',):         return 'IIb'
    if t in ('3',):          return 'III'
    return raw.strip() or 'Desconocida'


def _normalize_loe(raw: str) -> str:
    """
    Normaliza nivel de evidencia → A / B / C.

    Acepta el sistema clásico (A/B/C) y el sistema ACC/AHA 2019+:
      B-R (RCTs), B-NR (no aleatorizados) → B
      C-LD (datos limitados), C-EO (expert opinion) → C
    """
    t = str(raw).strip().upper()
    t = re.sub(r'^(LEVEL\s+OF\s+EVIDENCE|LEVEL|LOE|EVIDENCE)[:\s]*', '', t).strip()
    if t.startswith('A'): return 'A'
    if t.startswith('B'): return 'B'
    if t.startswith('C'): return 'C'
    return raw.strip() or 'Desconocida'


def _extract_cited_refs(text: str) -> str:
    """
    Extrae números de referencias citadas del texto de una recomendación.

    Detecta:
      - Superíndices Unicode: ¹²³ → 123
      - Grupos entre corchetes: [1,2,3] / [1-3]
      - Grupos entre paréntesis: (1,2) / (3-5)
      - Superíndices pegados a letras: "aspirin5,6" (PDFs sin espacio)
    """
    SUPERSCRIPT = str.maketrans('⁰¹²³⁴⁵⁶⁷⁸⁹', '0123456789')
    t = text.translate(SUPERSCRIPT)

    numbers: set[int] = set()

    # Grupos entre corchetes/paréntesis
    for m in re.finditer(r'[\[\(](\d[\d,\s\-–]+\d|\d)[\]\)]', t):
        for part in re.split(r'[,\s]+', m.group(1)):
            part = part.strip()
            if re.match(r'^\d+[–\-]\d+$', part):
                try:
                    a, b = re.split(r'[–\-]', part)
                    numbers.update(range(int(a), int(b) + 1))
                except ValueError:
                    pass
            elif part.isdigit():
                numbers.add(int(part))

    # Superíndices pegados a letras: "aspirin5,6"
    for m in re.finditer(r'(?<=[a-zA-Z\.\)])(\d{1,3}(?:,\d{1,3})*)', t):
        for n in m.group(1).split(','):
            if n.strip().isdigit():
                numbers.add(int(n.strip()))

    if not numbers:
        return ''
    return ', '.join(str(n) for n in sorted(numbers))


# ─────────────────────────────────────────────
# PATRONES DE DETECCIÓN
# ─────────────────────────────────────────────

# Patrones para identificar columnas en tablas
_REC_COL_RE   = re.compile(r'recommendation|recomendaci[oó]n|recommandation|raccomandazione', re.I)
_CLASS_COL_RE = re.compile(r'^class(e)?$|^cor$|class\s+of\s+rec|classe\s+de|klasse', re.I)
_LOE_COL_RE   = re.compile(r'^level|^loe$|evidence|niveau|livello|beweis', re.I)

# Valores válidos en celdas
_CLASS_VAL_RE = re.compile(r'^(i{1,3}[ab]?|iia|iib|iii)$', re.I)
_LOE_VAL_RE   = re.compile(r'^[abc](?:-[a-z]+)?$', re.I)

# Patrones inline para texto corrido (Estrategia B)
# Cubren los tres formatos más frecuentes:
#   1. "... (Level of Evidence: A) I"          — ACC/AHA clásico post-2005
#   2. "... (Class I; Level of Evidence: A)"   — ACC/AHA compacto
#   3. "I  ... (B)"                            — ESC columna izquierda
_INLINE_PATTERNS = [
    re.compile(
        r'(?P<text>[A-Z][^(]{15,500}?)'
        r'\(Level\s+of\s+Evidence\s*[:\s]+(?P<loe>[ABC](?:-[A-Z]+)?)\)'
        r'\s*(?P<cls>I{1,3}|IIa|IIb|III)\b',
        re.I | re.DOTALL
    ),
    re.compile(
        r'(?P<text>[A-Z][^(]{15,500}?)'
        r'\((?:Class\s+)?(?P<cls>I{1,3}|IIa|IIb|III)'
        r'\s*[;,/]\s*(?:Level\s+of\s+Evidence[:\s]+|LOE\s*[:\s]*)?'
        r'(?P<loe>[ABC](?:-[A-Z]+)?)\)',
        re.I | re.DOTALL
    ),
    re.compile(
        r'(?:^|\n)\s*(?P<cls>I{1,3}|IIa|IIb|III)\s+'
        r'(?P<text>[A-Z][^\n]{15,400}?)\s*'
        r'\((?:LOE[:\s]*)?(?P<loe>[ABC](?:-[A-Z]+)?)\)',
        re.I | re.DOTALL | re.MULTILINE
    ),
]


# ─────────────────────────────────────────────
# EXTRACCIÓN PRINCIPAL
# ─────────────────────────────────────────────

def extract_recommendations_from_pdf(pdf_path: str, debug: bool = False) -> list[dict]:
    """
    Extrae recomendaciones de una guía clínica en PDF.

    ESTRATEGIA A — Por tablas (pdfplumber.extract_tables):
      Identifica tablas con columnas Recomendación / Clase / LOE.
      Detecta las columnas por cabecera Y por contenido de celdas (fallback).
      Funciona con guías ESC y ACC/AHA en formato tabular.

    ESTRATEGIA B — Por patrones textuales (se activa si A < 5 resultados):
      Busca patrones inline en texto corrido:
        "(Level of Evidence: A) I"
        "(Class I; Level of Evidence: A)"
      Funciona con guías ACC/AHA en formato de caja de texto.

    Args:
        pdf_path: ruta al PDF de la guía clínica
        debug:    si True, imprime mensajes de diagnóstico

    Returns:
        Lista de dicts con: rec_number, text, rec_class, loe,
        references_cited, page (1-indexed), source ('tabla'|'texto')
    """
    log = []
    def dbg(msg):
        log.append(msg)
        if debug:
            print(msg)

    table_recs: list[dict] = []
    text_recs:  list[dict] = []

    # ── ESTRATEGIA A: tablas ─────────────────────────────────────────
    with pdfplumber.open(pdf_path) as pdf:
        dbg(f"[REC] PDF: {len(pdf.pages)} páginas")

        for page_idx, page in enumerate(pdf.pages):
            try:
                tables = page.extract_tables() or []
            except Exception:
                tables = []

            for table in tables:
                if not table or len(table) < 2:
                    continue

                header = [(str(c or '')).replace('\n', ' ').strip()
                          for c in (table[0] or [])]

                rec_ci = class_ci = loe_ci = -1

                for ci, h in enumerate(header):
                    if _REC_COL_RE.search(h)   and rec_ci   < 0: rec_ci   = ci
                    if _CLASS_COL_RE.match(h)  and class_ci < 0: class_ci = ci
                    if _LOE_COL_RE.match(h)    and loe_ci   < 0: loe_ci   = ci

                # Fallback: detectar columnas por valor si las cabeceras no coinciden
                if class_ci < 0:
                    class_votes: dict[int, int] = {}
                    loe_votes:   dict[int, int] = {}
                    for row in table[1:]:
                        for ci, cell in enumerate(row or []):
                            v = str(cell or '').strip()
                            if _CLASS_VAL_RE.match(v): class_votes[ci] = class_votes.get(ci, 0) + 1
                            if _LOE_VAL_RE.match(v):   loe_votes[ci]   = loe_votes.get(ci, 0) + 1
                    if class_votes: class_ci = max(class_votes, key=class_votes.get)
                    if loe_votes:   loe_ci   = max(loe_votes,   key=loe_votes.get)

                # Sin columna de clase → no es tabla de recomendaciones
                if class_ci < 0:
                    continue

                # Sin columna de texto explícita → la columna más larga
                if rec_ci < 0:
                    lengths: dict[int, int] = {}
                    for row in table[1:]:
                        for ci, cell in enumerate(row or []):
                            if ci not in (class_ci, loe_ci):
                                lengths[ci] = lengths.get(ci, 0) + len(str(cell or ''))
                    if lengths:
                        rec_ci = max(lengths, key=lengths.get)

                if rec_ci < 0:
                    continue

                dbg(f"[REC-A] Tabla en pág {page_idx+1}: "
                    f"rec={rec_ci} class={class_ci} loe={loe_ci}")

                for row in table[1:]:
                    def cv(idx: int) -> str:
                        if idx < 0 or not row or idx >= len(row):
                            return ''
                        return str(row[idx] or '').replace('\n', ' ').strip()

                    text  = cv(rec_ci)
                    cls   = cv(class_ci)
                    loe   = cv(loe_ci) if loe_ci >= 0 else ''

                    if len(text) < 20:
                        continue

                    table_recs.append({
                        'rec_number':       len(table_recs) + 1,
                        'text':             text,
                        'rec_class':        _normalize_class(cls),
                        'loe':              _normalize_loe(loe),
                        'references_cited': _extract_cited_refs(text),
                        'page':             page_idx + 1,
                        'source':           'tabla',
                    })

    dbg(f"[REC-A] {len(table_recs)} recomendaciones por tablas")

    # ── ESTRATEGIA B: patrones textuales ─────────────────────────────
    if len(table_recs) < 5:
        dbg("[REC-B] Activando estrategia textual (fallback)...")

        with pdfplumber.open(pdf_path) as pdf:
            for page_idx, page in enumerate(pdf.pages):
                try:
                    page_text = page.extract_text() or ''
                except Exception:
                    page_text = ''

                for pat in _INLINE_PATTERNS:
                    for m in pat.finditer(page_text):
                        raw_text = m.group('text').strip()
                        # Si el texto capturado es muy largo, tomar la última oración
                        if len(raw_text) > 400:
                            last_stop = raw_text.rfind('. ', len(raw_text) - 380)
                            if last_stop > 50:
                                raw_text = raw_text[last_stop + 2:].strip()
                        if len(raw_text) < 20:
                            continue
                        text_recs.append({
                            'rec_number':       len(text_recs) + 1,
                            'text':             raw_text,
                            'rec_class':        _normalize_class(m.group('cls')),
                            'loe':              _normalize_loe(m.group('loe')),
                            'references_cited': _extract_cited_refs(m.group('text')),
                            'page':             page_idx + 1,
                            'source':           'texto',
                        })

        dbg(f"[REC-B] {len(text_recs)} recomendaciones por texto")

    # Elegir el conjunto con más resultados
    recommendations = table_recs if len(table_recs) >= len(text_recs) else text_recs

    # Renumeración secuencial final
    for i, r in enumerate(recommendations, 1):
        r['rec_number'] = i

    dbg(f"[REC] Total: {len(recommendations)} recomendaciones")
    if not debug:
        for line in log:
            print(line)

    return recommendations


# ─────────────────────────────────────────────
# EXPORTACIÓN A EXCEL
# ─────────────────────────────────────────────

# Paleta de colores
_CLASS_COLORS = {'I': 'E2EFDA', 'IIa': 'FFF2CC', 'IIb': 'FCE4D6', 'III': 'F8CECC'}
_LOE_COLORS   = {'A': 'DAE8FC', 'B':   'E1D5E7', 'C':   'F5F5F5'}
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
    """Hoja principal: una fila por recomendación, código de colores por clase."""
    ws.title = 'Recomendaciones'

    cols = [
        ('Nº',                        5),
        ('Texto de la recomendación', 75),
        ('Clase',                     10),
        ('Nivel evidencia (LOE)',      16),
        ('Referencias citadas',        25),
        ('Página',                      8),
        ('Fuente extracción',          14),
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
            rec.get('rec_number', ''),
            rec.get('text', ''),
            cls,
            loe,
            rec.get('references_cited', ''),
            rec.get('page', ''),
            rec.get('source', ''),
            '',   # clase manual
            '',   # loe manual
            '',   # notas
        ]

        for ci, val in enumerate(values, 1):
            cell = ws.cell(row=rn, column=ci, value=val)
            cell.font      = Font(name='Arial', size=9)
            cell.alignment = Alignment(vertical='top', wrap_text=(ci == 2))

            if ci == 2:   # texto → color de fila según clase
                cell.fill = PatternFill('solid', start_color=_CLASS_COLORS.get(cls, 'F5F5F5'))
            elif ci == 3: # clase → badge coloreado
                cell.fill      = PatternFill('solid', start_color=_CLASS_COLORS.get(val, 'F5F5F5'))
                cell.alignment = Alignment(horizontal='center', vertical='top')
                cell.font      = Font(name='Arial', size=9, bold=True)
            elif ci == 4: # LOE → badge coloreado
                cell.fill      = PatternFill('solid', start_color=_LOE_COLORS.get(val, 'F5F5F5'))
                cell.alignment = Alignment(horizontal='center', vertical='top')
                cell.font      = Font(name='Arial', size=9, bold=True)

        ws.row_dimensions[rn].height = 55
        _thin_border(ws, rn, len(cols))


def _write_summary_sheet(ws, recommendations: list[dict]):
    """Hoja de resumen: tabla de contingencia Clase × LOE + leyenda."""
    ws.title = 'Resumen recomendaciones'

    CLASS_ORDER = ['I', 'IIa', 'IIb', 'III', 'Desconocida']
    LOE_ORDER   = ['A', 'B', 'C', 'Desconocida']

    # Construir matriz de frecuencias
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

    # Cabecera
    _style_header(ws.cell(row=1, column=1, value='Clase / LOE'))
    for ci, loe in enumerate(present_loe, 2):
        _style_header(ws.cell(row=1, column=ci, value=loe))
        ws.column_dimensions[get_column_letter(ci)].width = 10
    total_col = len(present_loe) + 2
    _style_header(ws.cell(row=1, column=total_col, value='TOTAL'))
    ws.column_dimensions[get_column_letter(total_col)].width = 10

    # Filas por clase
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
        t.font      = Font(bold=True, name='Arial', size=10)
        t.alignment = Alignment(horizontal='center')

    # Fila TOTAL
    tr = len(present_cls) + 2
    ws.cell(row=tr, column=1, value='TOTAL').font = Font(bold=True, name='Arial', size=10)
    for ci, loe in enumerate(present_loe, 2):
        c = ws.cell(row=tr, column=ci, value=loe_totals.get(loe, 0))
        c.font = Font(bold=True, name='Arial', size=10)
        c.alignment = Alignment(horizontal='center')
    g = ws.cell(row=tr, column=total_col, value=len(recommendations))
    g.font = Font(bold=True, name='Arial', size=10)
    g.alignment = Alignment(horizontal='center')

    # Leyenda
    leyenda = [
        ('', ''),
        ('CLASE I',   'Beneficio >> riesgo. Está indicado / recomendado.',         'E2EFDA'),
        ('CLASE IIa', 'Beneficio > riesgo. Es razonable realizarlo.',               'FFF2CC'),
        ('CLASE IIb', 'Beneficio ≥ riesgo. Puede considerarse.',                    'FCE4D6'),
        ('CLASE III', 'Sin beneficio o dañino. No está recomendado.',               'F8CECC'),
        ('', ''),
        ('LOE A',     'Múltiples ECAs o meta-análisis de alta calidad.',            'DAE8FC'),
        ('LOE B',     'Un ECA o estudios no aleatorizados de gran tamaño.',         'E1D5E7'),
        ('LOE C',     'Consenso de expertos, estudios pequeños o registros.',       'F5F5F5'),
    ]
    for offset, row_data in enumerate(leyenda, tr + 2):
        if len(row_data) == 2:
            continue
        label, desc, color = row_data
        c = ws.cell(row=offset, column=1, value=label)
        c.font = Font(name='Arial', size=9)
        c.fill = PatternFill('solid', start_color=color)
        ws.cell(row=offset, column=2, value=desc).font = Font(name='Arial', size=9)
        ws.column_dimensions['B'].width = 55


def export_recommendations_to_excel(recommendations: list[dict], output_path: str):
    """
    Exporta las recomendaciones a un Excel con dos hojas:
      - 'Recomendaciones'         : una fila por recomendación
      - 'Resumen recomendaciones' : tabla Clase × LOE + leyenda
    """
    wb = Workbook()
    ws_rec = wb.active
    _write_rec_sheet(ws_rec, recommendations)
    ws_sum = wb.create_sheet('Resumen recomendaciones')
    _write_summary_sheet(ws_sum, recommendations)
    wb.save(output_path)
    print(f'[OK] Excel de recomendaciones guardado: {output_path}')
