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
    """Inlocuieste continutul unui <w:tc> cu un singur paragraf/run continand 'text',
    centrat si cu font 12pt (marimea folosita si de exportul original BIOSILV).
    Sterge intotdeauna continutul anterior (unele casete din sablon au text
    static rezidual, ex. 'Cod') - dar sare peste crearea pPr/jc cand textul
    e gol, ca sa reduca oarecum consumul de memorie fara sa piarda corectitudinea."""
    p = tc.find(qn('w:p'))
    if p is None:
        if not text:
            return
        p = OxmlElement('w:p')
        tc.append(p)
    else:
        for r in p.findall(qn('w:r')):
            p.remove(r)
    if not text:
        return
    pPr = p.find(qn('w:pPr'))
    if pPr is None:
        pPr = OxmlElement('w:pPr')
        p.insert(0, pPr)
    jc = pPr.find(qn('w:jc'))
    if jc is None:
        jc = OxmlElement('w:jc')
        pPr.append(jc)
    jc.set(qn('w:val'), 'center')
    r = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    sz = OxmlElement('w:sz')
    sz.set(qn('w:val'), '24')  # 24 half-points = 12pt
    rPr.append(sz)
    szCs = OxmlElement('w:szCs')
    szCs.set(qn('w:val'), '24')
    rPr.append(szCs)
    r.append(rPr)
    t = OxmlElement('w:t')
    t.set(qn('xml:space'), 'preserve')
    t.text = str(text)
    r.append(t)
    p.append(r)


def fill_row_pair(header_tcs, box_tcs, values):
    """header_tcs/box_tcs = liste (tc,span). values = lista de siruri, aceeasi
    lungime ca header_tcs; fiecare caracter merge intr-o caseta box succesiva."""
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


def width_to_boxindex(box_tcs, target_width):
    """Gaseste indexul in box_tcs a carui pozitie cumulativa (in unitati de
    latime) corespunde lui target_width - folosit pt. a localiza precis un
    camp zecimal in randul de casete, indiferent de eventuale casete mai
    late (span>1) aparute mai devreme in rand."""
    cum = 0
    for i, (tc, span) in enumerate(box_tcs):
        if cum >= target_width:
            return i
        cum += span
    return len(box_tcs)


def fmt_dec_boxes(v, nboxes, leading_space=True):
    """Reproduce exact modul in care programul vechi afiseaza valorile zecimale:
    ' ' + parte_intreaga + ',' + zecimala, apoi aliniat la dreapta pe 'nboxes'
    casete - daca sirul e mai lung decat nr. de casete, ultimele caractere in
    exces (de regula ',X') se comprima in ULTIMA caseta. Returneaza o lista de
    'nboxes' siruri, cate unul per caseta."""
    if v in (None, ''):
        return [''] * nboxes
    try:
        f = float(v)
    except (TypeError, ValueError):
        return list(str(v).rjust(nboxes))[:nboxes]
    s = f"{f:.1f}"
    intpart, dec = s.split('.')
    prefix = ' ' if leading_space else ''
    value_str = f"{prefix}{intpart},{dec}"
    if len(value_str) <= nboxes:
        padded = value_str.rjust(nboxes)
        return list(padded)
    overflow = len(value_str) - nboxes + 1
    head = list(value_str[:nboxes - 1])
    tail = value_str[nboxes - 1:]
    return head + [tail]


def fill_decimal_field(box_tcs, start_idx, nboxes, value, leading_space=True):
    """Completeaza direct 'nboxes' casete incepand de la start_idx, folosind
    fmt_dec_boxes. Foloseste indexare directa in box_tcs (nu prin fill_row_pair),
    ca sa evitam driftul provocat de casetele mai late din alte campuri."""
    parts = fmt_dec_boxes(value, nboxes, leading_space=leading_space)
    for i, part in enumerate(parts):
        idx = start_idx + i
        if idx < len(box_tcs):
            set_tc_text(box_tcs[idx][0], part)


def sv(v):
    """Pastreaza 0 ca '0' (spre deosebire de 'v or ""' care il transforma gresit in gol)."""
    if v is None or v == '':
        return ''
    return str(v)


def sv0blank(v):
    """Pentru sub-campurile din tabelul de specii (M, AMS, ELG, VIT, TEL, CP):
    valoarea 0 se afiseaza gol pe fisa de teren (confirmat din fisierul real)."""
    if v in (None, '', 0):
        return ''
    return str(v)


def fmt_blank_zero(v):
    if v in (None, '', 0, '0'):
        return ''
    return str(v)


def fill_front_table(tbl, a, isj='01', os='01', varsta_offset=0):
    rows = tbl.rows

    h1 = raw_tcs(rows[1])
    up_str = f"{int(a['up']):02d}" if a.get('up') not in (None, '') else ''
    ua_num = f"{int(a['ua1']):>3d}" if a.get('ua1') is not None else '   '
    ua_raw = f"{ua_num}{str(a.get('ua2') or '').ljust(2)}"[:5]
    fct = f"{sv(a.get('fct1'))}{sv(a.get('fct2'))}{sv(a.get('fct3'))}"
    b1 = raw_tcs(rows[3])
    values1 = [str(isj or ''), str(os or ''), sv(a.get('adm')), up_str, ua_raw,
               '', '', '', '', '', sv(a.get('sup')), sv(a.get('ff')),
               '', sv(a.get('fls')), sv(a.get('gf')),
               fct, '', '', '']
    fill_row_pair(h1, b1, values1)
    fill_decimal_field(b1, width_to_boxindex(b1, 24), 3, a.get('spr'))  # SUPRAF (HA)

    h2 = raw_tcs(rows[5])
    ta_val = a.get('ta')
    ta_val = (ta_val or 0) + varsta_offset if ta_val not in (None, '') else ta_val
    b2 = raw_tcs(rows[6])
    values2 = [
        sv(a.get('rlf')), sv(a.get('cnf')), sv(a.get('exp')),
        sv(a.get('inc')), '', '',
        sv(a.get('sol')), sv(a.get('erz')), sv(a.get('flr')),
        sv(a.get('ts')), sv(a.get('inv')), '',
        sv(a.get('lp1')), sv(a.get('lp2')), sv(a.get('lp3')), '', '',
    ]
    fill_row_pair(h2, b2, values2)
    fill_decimal_field(b2, width_to_boxindex(b2, 7), 5, a.get('alt1'))  # ALTITUDINE MIN(MED)
    alt2_str = sv(a.get('alt2')).rjust(4)
    alt2_start = width_to_boxindex(b2, 12)
    for i, ch in enumerate(alt2_str):
        idx = alt2_start + i
        if idx < len(b2):
            set_tc_text(b2[idx][0], ch)

    h3 = raw_tcs(rows[7])
    pol = f"{fmt_blank_zero(a.get('pol1'))}{fmt_blank_zero(a.get('pol2'))}{fmt_blank_zero(a.get('pol3'))}"
    b3 = raw_tcs(rows[8])
    values3 = [
        sv(a.get('tp')), sv(a.get('crti')), pol, sv(a.get('lit')),
        sv(a.get('drm')), '', '',
        '', '',
        '', '',
        '', '', '', '', '', '', '', '', '', '', '', '', '',
    ]
    fill_row_pair(h3, b3, values3)
    fill_decimal_field(b3, width_to_boxindex(b3, 18), 2, a.get('cns'), leading_space=False)  # CNS
    # campurile de dupa DRUM se pozitioneaza direct dupa latime (nu prin fill_row_pair
    # in cascada), ca sa nu fie afectate de caseta lata ascunsa in interiorul DRUM
    dst_str = sv(a.get('dst')).rjust(2)
    for i, ch in enumerate(dst_str):
        idx = width_to_boxindex(b3, 15) + i
        if idx < len(b3):
            set_tc_text(b3[idx][0], ch)
    str_idx = width_to_boxindex(b3, 17)
    if str_idx < len(b3):
        set_tc_text(b3[str_idx][0], sv(a.get('str')))
    clp_idx = width_to_boxindex(b3, 20)
    if clp_idx < len(b3):
        set_tc_text(b3[clp_idx][0], sv(a.get('clp')))
    ta_str = sv(ta_val).rjust(3)
    for i, ch in enumerate(ta_str):
        idx = width_to_boxindex(b3, 21) + i
        if idx < len(b3):
            set_tc_text(b3[idx][0], ch)
    reg_idx = width_to_boxindex(b3, 24)
    if reg_idx < len(b3):
        set_tc_text(b3[reg_idx][0], sv(a.get('reg')))

    h4 = raw_tcs(rows[10])
    values4 = [
        sv0blank(a.get('te')), sv0blank(a.get('ex')), sv0blank(a.get('urg')),
        sv0blank(a.get('prm')), sv0blank(a.get('nin')), sv(a.get('nid')),
        sv(a.get('lx1')), sv(a.get('lx2')),
        '', '', '',
        sv(a.get('dc1')), sv(a.get('dc2')), sv(a.get('dc3')), sv(a.get('dc4')),
        '', '', '', '', '',
    ]
    fill_row_pair(h4, raw_tcs(rows[11]), values4)

    h5 = raw_tcs(rows[13])
    cel_pairs = split_pairs(a.get('cel1'), a.get('cpr1'), a.get('cel'))
    values5 = []
    for cod, pct in cel_pairs:
        values5.append(cod)
        pct_str = str(pct) if cod else ''
        if len(pct_str) == 1:
            pct_str = ' ' + pct_str  # procent aliniat la dreapta cand e o cifra
        values5.append(pct_str)
    values5 = values5[:12] + [
        sv(a.get('sba')), sv(a.get('so')), sv(a.get('mr')), sv(a.get('ds')),
        '', '', '', '', '',
    ]
    fill_row_pair(h5, raw_tcs(rows[14]), values5)

    h6 = raw_tcs(rows[16])
    mel_pairs = split_pairs(a.get('mel1'), a.get('mpr1'), a.get('mel'))
    vs_str = sv0blank(a.get('vs'))
    if 0 < len(vs_str) < 2:
        vs_str = vs_str.rjust(2)  # VS aliniat la dreapta (2 casete)
    elif len(vs_str) == 0:
        vs_str = ''
    values6 = [
        vs_str,
        mel_pairs[0][0], mel_pairs[0][1],
        mel_pairs[1][0], '',
        mel_pairs[2][0], '',
        mel_pairs[3][0], '',
        mel_pairs[4][0], '',
        mel_pairs[5][0], '',
        sv0blank(a.get('soc')), sv0blank(a.get('rs')), '',
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
            vrt_val = e.get('vrt')
            vrt_val = (vrt_val or 0) + varsta_offset if vrt_val not in (None, '') else vrt_val
            vrt_str = sv(vrt_val)
            if len(vrt_str) < 3:
                vrt_str = vrt_str.rjust(3)  # VRT aliniat la dreapta
            prp_str = sv(e.get('prp'))
            if len(prp_str) == 1:
                prp_str = ' ' + prp_str  # PRP aliniat la dreapta cand e o cifra
            values7 = [
                sv(e.get('elm')), sv(e.get('mrg')), vrt_str,
                prp_str, sv(e.get('dm')), sv(e.get('hm')),
                '', '', '', '', '', '', '',
                '', '',  # VOL, CRS - lasate goale intentionat (se calculeaza dupa masuratorile din teren)
                '',      # PEX - la fel, gol pe fisa de teren
                sv(e.get('prov')),
            ]
            fill_row_pair(h7, b7, values7)
            # M/CP/AMS/ELG/VIT/TEL/CAL pozitionate direct dupa latime, ca sa nu
            # fie afectate de caseta lata ascunsa in interiorul HM
            for w, val in [(14, sv0blank(e.get('m'))), (15, sv0blank(e.get('cp'))),
                           (16, sv0blank(e.get('ams'))), (17, sv0blank(e.get('elg'))),
                           (18, sv0blank(e.get('vit'))), (19, sv0blank(e.get('tel')))]:
                idx = width_to_boxindex(b7, w)
                if idx < len(b7):
                    set_tc_text(b7[idx][0], val)
            cal_str = sv(e.get('cal'))
            cal_start = width_to_boxindex(b7, 20)
            for i2, ch in enumerate(cal_str[:2]):
                idx = cal_start + i2
                if idx < len(b7):
                    set_tc_text(b7[idx][0], ch)
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


def generate_fise(a_bytes, s_bytes, template_bytes, isj='01', os_code='01', varsta_offset=0, progress_cb=None):
    """a_bytes/s_bytes/template_bytes = continut binar (bytes) al fisierelor A.xlsx,
    S.xlsx si al sablonului FISA_DESCR.docx. varsta_offset = numar de ani adaugati
    la varsta actuala (TA) si la varsta fiecarui element de arboret (VRT), util
    pentru a genera fise valabile la o data ulterioara fara sa reconverstesti DBF-ul.
    Returneaza bytes-ii documentului .docx generat (fata completata + verso
    necompletat, 2 fise/pagina).

    NOTA IMPORTANTA (memorie): fiecare fisa e construita si serializata separat,
    ca text (bytes), apoi obiectele XML sunt eliberate imediat - documentul NU e
    tinut intreg ca un singur arbore XML urias in memorie in timpul generarii, ca
    sa suporte fisiere cu multe UA-uri fara sa depaseasca limita de RAM (ex. pe
    Streamlit Community Cloud, care ofera doar ~1GB)."""
    import gc
    import zipfile
    from lxml import etree
    from docx.oxml import parse_xml

    A = _load_xlsx_stream(io.BytesIO(a_bytes), bl.A_HEADER)
    S = _load_xlsx_stream(io.BytesIO(s_bytes), bl.S_HEADER)

    def key(r):
        return (r['adm'], r['up'], r['ua1'], str(r['ua2']))

    s_by = defaultdict(list)
    for r in S:
        s_by[key(r)].append(r)
    for a in A:
        a['_species'] = s_by.get(key(a), [])

    # citim sablonul o singura data: extragem sirurile XML ale celor doua tabele
    # (fata + verso), fara sa pastram documentul mare "viu" in memorie ulterior
    master = docx.Document(io.BytesIO(template_bytes))
    front_xml_bytes = etree.tostring(master.tables[0]._tbl)
    verso_xml_bytes = etree.tostring(master.tables[2]._tbl)
    del master
    gc.collect()

    # extragem direct din arhiva .docx originala structura document.xml (text),
    # ca sa reconstruim la final fara sa trecem prin serializarea unui arbore urias
    with zipfile.ZipFile(io.BytesIO(template_bytes)) as zin:
        doc_xml = zin.read('word/document.xml').decode('utf-8')
        other_files = {name: zin.read(name) for name in zin.namelist() if name != 'word/document.xml'}

    body_open_end = doc_xml.index('<w:body>') + len('<w:body>')
    sect_start = doc_xml.index('<w:sectPr')
    head = doc_xml[:body_open_end]
    tail = doc_xml[sect_start:]

    PAGE_BREAK_XML = b'<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
    SPACER_XML = b'<w:p/>'

    n = len(A)
    parts = []  # bucati de XML (bytes) ale corpului documentului, unite la final
    for i, a in enumerate(A):
        front_el = parse_xml(front_xml_bytes)
        front_table = docx.table.Table(front_el, None)
        fill_front_table(front_table, a, isj=isj, os=os_code, varsta_offset=varsta_offset)
        parts.append(etree.tostring(front_el))
        parts.append(SPACER_XML)
        del front_el, front_table

        if i % 2 == 1 or i == n - 1:
            parts.append(PAGE_BREAK_XML)
            # cate un verso pentru fiecare fisa din perechea curenta (1 sau 2)
            pair_size = 2 if (i % 2 == 1) else 1
            for _ in range(pair_size):
                parts.append(verso_xml_bytes)
                parts.append(SPACER_XML)
            if i < n - 1:
                parts.append(PAGE_BREAK_XML)

        if i % 10 == 0:
            gc.collect()
        if progress_cb:
            progress_cb((i + 1) / n)

    body_content = b''.join(parts)
    del parts
    gc.collect()

    full_doc_xml = head.encode('utf-8') + body_content + tail.encode('utf-8')
    del body_content

    out = io.BytesIO()
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name, data in other_files.items():
            zout.writestr(name, data)
        zout.writestr('word/document.xml', full_doc_xml)

    return out.getvalue(), len(A)
