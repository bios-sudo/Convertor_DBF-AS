# -*- coding: utf-8 -*-
"""
Genereaza fisele de teren FOLOSIND EXACT formularul tipizat FISA_DESCR.docx:
- fata (fisa descriptiva) e completata cu datele din A.xlsx / S.xlsx
- versoul (elemente taxatorice) ramane necompletat, gata de teren
- 2 fise pe pagina (fata+fata, apoi verso+verso), pentru economie de hartie

Functia principala: generate_fise(a_bytes, s_bytes, template_bytes, isj, os) -> bytes (docx)
"""
import io
import copy
from collections import defaultdict
import openpyxl
import docx
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import biosilv_lib as bl


def raw_tcs(row):
    """Lista (tc_element, gridSpan) NEEXPANDATA pentru un rand de tabel."""
    out = []
    for tc in row._tr.findall(qn('w:tc')):
        tcPr = tc.find(qn('w:tcPr'))
        span = 1
        if tcPr is not None:
            gs = tcPr.find(qn('w:gridSpan'))
            if gs is not None:
                span = int(gs.get(qn('w:val')))
        out.append((tc, span))
    return out


def set_tc_text(tc, text):
    """Inlocuieste continutul unui <w:tc> cu un singur paragraf/run continand 'text'."""
    p = tc.find(qn('w:p'))
    if p is None:
        p = OxmlElement('w:p')
        tc.append(p)
    for r in p.findall(qn('w:r')):
        p.remove(r)
    if text:
        r = OxmlElement('w:r')
        rPr = OxmlElement('w:rPr')
        sz = OxmlElement('w:sz')
        sz.set(qn('w:val'), '16')
        rPr.append(sz)
        r.append(rPr)
        t = OxmlElement('w:t')
        t.set(qn('xml:space'), 'preserve')
        t.text = str(text)
        r.append(t)
        p.append(r)


def fill_row_pair(header_tcs, box_tcs, values):
    bidx = 0
    for (htc, hspan), val in zip(header_tcs, values):
        remaining = hspan
        chars = list(str(val)) if val not in (None, '') else []
        while remaining > 0 and bidx < len(box_tcs):
            btc, bspan = box_tcs[bidx]
            ch = chars.pop(0) if chars else ''
            set_tc_text(btc, ch)
            remaining -= bspan
            bidx += 1


def split_pairs(cel1, pct1, block, n=5):
    out = [(cel1 or '', str(pct1) if pct1 not in (None, '') else '')]
    s = (block or '').ljust(n * 4)
    for i in range(n):
        chunk = s[i * 4:i * 4 + 4]
        cod = chunk[0:3].strip()
        pct = chunk[3:4].strip()
        out.append((cod, pct))
    return out


def fmt_dec(v, width):
    """Campuri zecimale (SUPRAF, ALT MIN, CNS, CRS): fara virgula explicita -
    ultima caseta din dreapta e deja rezervata ca zecimala in formular."""
    if v in (None, ''):
        return ' ' * width
    try:
        f = float(v)
    except (TypeError, ValueError):
        return str(v).rjust(width)
    s = f"{f:.1f}"
    intpart, dec = s.split('.')
    return intpart.rjust(width - 1)[-(width - 1):] + dec[:1]


def fmt_blank_zero(v):
    if v in (None, '', 0, '0'):
        return ''
    return str(v)


def fill_front_table(tbl, a, isj='01', os='01'):
    rows = tbl.rows

    h1 = raw_tcs(rows[1])
    up_str = f"{int(a['up']):02d}" if a.get('up') not in (None, '') else ''
    ua_num = f"{int(a['ua1']):>3d}" if a.get('ua1') is not None else '   '
    ua_raw = f"{ua_num}{str(a.get('ua2') or '').ljust(2)}"[:5]
    fct = f"{a.get('fct1') or ''}{a.get('fct2') or ''}{a.get('fct3') or ''}"
    values1 = [str(isj or ''), str(os or ''), str(a.get('adm') or ''), up_str, ua_raw,
               '', '', '', '', '', str(a.get('sup') or ''), str(a.get('ff') or ''),
               fmt_dec(a.get('spr'), 3), str(a.get('fls') or ''), str(a.get('gf') or ''),
               fct, '', '', '']
    fill_row_pair(h1, raw_tcs(rows[3]), values1)

    h2 = raw_tcs(rows[5])
    values2 = [
        str(a.get('rlf') or ''), str(a.get('cnf') or ''), str(a.get('exp') or ''),
        str(a.get('inc') or ''), fmt_dec(a.get('alt1'), 5), str(a.get('alt2') or ''),
        str(a.get('sol') or ''), str(a.get('erz') or ''), str(a.get('flr') or ''),
        str(a.get('ts') or ''), str(a.get('inv') or ''), '',
        str(a.get('lp1') or ''), str(a.get('lp2') or ''), str(a.get('lp3') or ''), '', '',
    ]
    fill_row_pair(h2, raw_tcs(rows[6]), values2)

    h3 = raw_tcs(rows[7])
    pol = f"{fmt_blank_zero(a.get('pol1'))}{fmt_blank_zero(a.get('pol2'))}{fmt_blank_zero(a.get('pol3'))}"
    values3 = [
        str(a.get('tp') or ''), str(a.get('crti') or ''), pol, str(a.get('lit') or ''),
        str(a.get('drm') or ''), str(a.get('dst') or ''), str(a.get('str') or ''),
        fmt_dec(a.get('cns'), 2), str(a.get('clp') or ''),
        str(a.get('ta') or ''), str(a.get('reg') or ''),
        '', '', '', '', '', '', '', '', '', '', '', '', '',
    ]
    fill_row_pair(h3, raw_tcs(rows[8]), values3)

    h4 = raw_tcs(rows[10])
    values4 = [
        str(a.get('te') or ''), str(a.get('ex') or ''), str(a.get('urg') or ''),
        str(a.get('prm') or ''), str(a.get('nin') or ''), str(a.get('nid') or ''),
        str(a.get('lx1') or ''), str(a.get('lx2') or ''),
        '', '', '',
        str(a.get('dc1') or ''), str(a.get('dc2') or ''), str(a.get('dc3') or ''), str(a.get('dc4') or ''),
        '', '', '', '', '',
    ]
    fill_row_pair(h4, raw_tcs(rows[11]), values4)

    h5 = raw_tcs(rows[13])
    cel_pairs = split_pairs(a.get('cel1'), a.get('cpr1'), a.get('cel'))
    values5 = []
    for cod, pct in cel_pairs:
        values5.append(cod)
        values5.append(pct)
    values5 = values5[:12] + [
        str(a.get('sba') or ''), str(a.get('so') or ''), str(a.get('mr') or ''), str(a.get('ds') or ''),
        '', '', '', '', '',
    ]
    fill_row_pair(h5, raw_tcs(rows[14]), values5)

    h6 = raw_tcs(rows[16])
    mel_pairs = split_pairs(a.get('mel1'), a.get('mpr1'), a.get('mel'))
    values6 = [
        str(a.get('vs') or ''),
        mel_pairs[0][0], mel_pairs[0][1],
        mel_pairs[1][0], '',
        mel_pairs[2][0], '',
        mel_pairs[3][0], '',
        mel_pairs[4][0], '',
        mel_pairs[5][0], '',
        str(a.get('soc') or ''), str(a.get('rs') or ''), '',
        '', '', '', '', '', '', '', '',
    ]
    fill_row_pair(h6, raw_tcs(rows[17]), values6)

    h7 = raw_tcs(rows[18])
    species = a.get('_species', [])
    for i in range(8):
        row_idx = 20 + i
        if row_idx >= len(rows):
            break
        b7 = raw_tcs(rows[row_idx])
        if i < len(species):
            e = species[i]
            pex = f"{e.get('pex1') or ''}{e.get('pex2') or ''}{e.get('pex3') or ''}"
            values7 = [
                str(e.get('elm') or ''), str(e.get('mrg') or ''), str(e.get('vrt') or ''),
                str(e.get('prp') or ''), str(e.get('dm') or ''), str(e.get('hm') or ''),
                str(e.get('m') or ''), str(e.get('cp') or ''), str(e.get('ams') or ''),
                str(e.get('elg') or ''), str(e.get('vit') or ''), str(e.get('tel') or ''),
                str(e.get('cal') or ''), str(e.get('vol') or ''), fmt_dec(e.get('crsc'), 2),
                pex, str(e.get('prov') or ''),
            ]
        else:
            values7 = ['' for _ in h7]
        fill_row_pair(h7, b7, values7)


def _load_xlsx_stream(file_obj, header):
    wb = openpyxl.load_workbook(file_obj, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    h = [str(x).strip() for x in rows[0]]
    idx = {c: h.index(c) for c in header if c in h}
    out = []
    for r in rows[1:]:
        if r is None or all(v is None for v in r):
            continue
        out.append({c: r[idx[c]] for c in header})
    return out


def generate_fise(a_bytes, s_bytes, template_bytes, isj='01', os_code='01', progress_cb=None):
    """a_bytes/s_bytes/template_bytes = continut binar (bytes) al fisierelor A.xlsx,
    S.xlsx si al sablonului FISA_DESCR.docx. Returneaza bytes-ii documentului .docx
    generat (fata completata + verso necompletat, 2 fise/pagina)."""
    A = _load_xlsx_stream(io.BytesIO(a_bytes), bl.A_HEADER)
    S = _load_xlsx_stream(io.BytesIO(s_bytes), bl.S_HEADER)

    def key(r):
        return (r['adm'], r['up'], r['ua1'], str(r['ua2']))

    s_by = defaultdict(list)
    for r in S:
        s_by[key(r)].append(r)
    for a in A:
        a['_species'] = s_by.get(key(a), [])

    master = docx.Document(io.BytesIO(template_bytes))
    front_template = master.tables[0]._tbl
    verso_template = master.tables[2]._tbl
    body = master.element.body

    sectPr = body.find(qn('w:sectPr'))
    for child in list(body):
        if child is not sectPr:
            body.remove(child)

    def page_break_p():
        p_break = OxmlElement('w:p')
        r = OxmlElement('w:r')
        br = OxmlElement('w:br')
        br.set(qn('w:type'), 'page')
        r.append(br)
        p_break.append(r)
        return p_break

    def spacer_p():
        return OxmlElement('w:p')

    def insert_before_sect(el):
        if sectPr is not None:
            body.insert(list(body).index(sectPr), el)
        else:
            body.append(el)

    n = len(A)
    pair_count = (n + 1) // 2
    for pi in range(pair_count):
        pair = A[pi * 2: pi * 2 + 2]

        for a in pair:
            front = copy.deepcopy(front_template)
            front_table = docx.table.Table(front, master)
            fill_front_table(front_table, a, isj=isj, os=os_code)
            insert_before_sect(front)
            insert_before_sect(spacer_p())

        insert_before_sect(page_break_p())

        for _ in pair:
            verso = copy.deepcopy(verso_template)
            insert_before_sect(verso)
            insert_before_sect(spacer_p())

        if pi < pair_count - 1:
            insert_before_sect(page_break_p())

        if progress_cb:
            progress_cb((pi + 1) / pair_count)

    out = io.BytesIO()
    master.save(out)
    return out.getvalue(), len(A)
