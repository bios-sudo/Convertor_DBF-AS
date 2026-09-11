# -*- coding: utf-8 -*-
"""
Bibliotecă de conversie BIOSILV (.DBF) <-> A.xlsx / S.xlsx
Reguli deduse prin compararea baza_date_Biosilv.DBF cu A.xlsx si S.xlsx (exemplu).
"""
import struct
import datetime
import openpyxl

ENCODING = 'cp1250'

# ---------------------------------------------------------------------------
# Citire DBF
# ---------------------------------------------------------------------------
def read_dbf(path, encoding=ENCODING):
    with open(path, 'rb') as f:
        data = f.read()
    numrec = struct.unpack('<I', data[4:8])[0]
    lenheader = struct.unpack('<H', data[8:10])[0]
    recsize = struct.unpack('<H', data[10:12])[0]
    # Parcurgem descriptorii de camp unul cate unul pana la terminatorul 0x0D.
    # NU folosim formula (lenheader-33)//32, pentru ca tabelele Visual FoxPro
    # legate de o baza de date (.DBC) au dupa ultimul camp un bloc suplimentar
    # de 263 octeti ("backlink" catre .DBC), inclus in lenheader dar care NU
    # sunt descriptori de camp -> formula veche numara "campuri" fantoma.
    fields = []
    pos = 32
    while data[pos] != 0x0D:
        fr = data[pos:pos + 32]
        pos += 32
        name = fr[:11].split(b'\x00')[0].decode('latin1')
        if not name:
            continue
        ftype = fr[11:12].decode('latin1')
        flen = fr[16]
        fdec = fr[17]
        fields.append((name, ftype, flen, fdec))
    header_extra = data[pos + 1:lenheader]  # bloc backlink VFP/DBC (poate fi gol)
    records = []
    pos = lenheader
    for r in range(numrec):
        rec = data[pos:pos + recsize]
        pos += recsize
        if len(rec) < recsize:
            break
        if rec[0:1] == b'*':
            continue  # rand sters
        offset = 1
        row = {}
        for (name, ftype, flen, fdec) in fields:
            raw = rec[offset:offset + flen]
            offset += flen
            try:
                text = raw.decode(encoding, errors='replace')
            except Exception:
                text = raw.decode('latin1', errors='replace')
            val = text.strip()
            if ftype in ('N', 'F'):
                if val in ('', '.', '-'):
                    v = 0
                else:
                    try:
                        v = float(val) if fdec > 0 else int(float(val))
                    except ValueError:
                        v = 0
            else:
                v = val
            row[name] = v
            row['_raw_' + name] = text  # varianta cu spatii pastrate (pt. campuri gen UA)
        records.append(row)
    return fields, records, header_extra


def _fmt_dbf_value(val, ftype, flen, fdec):
    if ftype in ('N', 'F'):
        if val in (None, ''):
            val = 0
        if fdec > 0:
            s = f"{float(val):.{fdec}f}"
        else:
            s = str(int(val))
        return s.rjust(flen)[:flen].encode('latin1')
    else:
        s = '' if val is None else str(val)
        return s.encode('latin1', errors='replace').ljust(flen)[:flen]


def write_dbf(path, fields, rows, encoding=ENCODING, header_extra=b''):
    """fields = lista (name, type, len, dec) EXACT ca in DBF-ul original (BIOSILV).
    rows = lista de dict-uri {nume_camp: valoare}.
    header_extra = octetii dintre terminatorul 0x0D al descriptorilor de camp
    si inceputul datelor (blocul "backlink" catre .DBC, pentru tabele Visual
    FoxPro legate de o baza de date) - se copiaza EXACT din fisierul-sablon
    original, ca sa ramana identic structural cu ce asteapta BIOSILV."""
    recsize = 1 + sum(f[2] for f in fields)
    numrec = len(rows)
    today = datetime.date.today()
    header = bytearray(32)
    header[0] = 0x30  # Visual FoxPro
    header[1] = today.year - 1900
    header[2] = today.month
    header[3] = today.day
    header[4:8] = struct.pack('<I', numrec)
    lenheader = 32 + 32 * len(fields) + 1 + len(header_extra)
    header[8:10] = struct.pack('<H', lenheader)
    header[10:12] = struct.pack('<H', recsize)
    out = bytearray()
    out += header
    for (name, ftype, flen, fdec) in fields:
        fr = bytearray(32)
        nb = name.encode('latin1')[:10]
        fr[0:len(nb)] = nb
        fr[11:12] = ftype.encode('latin1')
        fr[16] = flen
        fr[17] = fdec
        out += fr
    out += b'\x0d'  # terminator descriptori de camp
    out += header_extra  # backlink VFP/DBC, copiat din sablon (poate fi gol)
    for row in rows:
        out += b' '  # not deleted
        for (name, ftype, flen, fdec) in fields:
            out += _fmt_dbf_value(row.get(name, ''), ftype, flen, fdec)
    out += b'\x1a'  # EOF marker
    with open(path, 'wb') as f:
        f.write(out)


# ---------------------------------------------------------------------------
# Campuri A / S (ordinea exacta din fisierele-exemplu A.xlsx / S.xlsx)
# ---------------------------------------------------------------------------
A_HEADER = ['adm', 'up', 'ua1', 'ua2', 'dec1', 'dec2', 'dec3', 'sup', 'ff', 'spr', 'fls', 'gf',
            'fct1', 'fct2', 'fct3', 'rlf', 'cnf', 'exp', 'inc', 'alt1', 'alt2', 'sol', 'erz', 'flr',
            'ts', 'inv', 'tp', 'crti', 'pol1', 'pol2', 'pol3', 'lit', 'drm', 'dst', 'str', 'cns',
            'clp', 'ta', 'reg', 'te', 'ex', 'urg', 'prm', 'nin', 'nid', 'lx1', 'lx2', 'lp1', 'lp2',
            'lp3', 'dc1', 'dc2', 'dc3', 'dc4', 'sba', 'so', 'mr', 'ds', 'cel1', 'cpr1', 'cel', 'vs',
            'soc', 'rs', 'mel1', 'mpr1', 'mel']

S_HEADER = ['adm', 'up', 'ua1', 'ua2', 'elm', 'mrg', 'prov', 'prp', 'spf', 'vrt', 'ams', 'elg',
            'vit', 'tel', 'cal', 'pex1', 'pex2', 'pex3', 'dm', 'dmc', 'hm', 'hmc', 'm', 'hsi', 'cp',
            'cpc', 'vol', 'crs', 'crsc', 'vexpr', 'vexra', 'vexcu', 'vexig', 'vextc']


def _fmt_pair_block(codes_pcts):
    """codes_pcts = lista de (cod, procent) -> sir de 4 caractere per pereche (cod ljust 3 + o cifra)."""
    out = ''
    for cod, pct in codes_pcts:
        cod = (cod or '')[:3].ljust(3)
        p = int(pct or 0)
        out += f"{cod}{p:1d}"
    return out


def _numcode(rec, name):
    """Pentru campuri caracter cu coduri numerice (RLF, POL1, POL2): '' -> 0, altfel int."""
    raw = rec.get('_raw_' + name, '').strip()
    if raw == '':
        return 0
    try:
        return int(raw)
    except ValueError:
        return raw


def _blankable_num(rec, name):
    """SOL: cod afisat ca text; daca valoarea e 0 (necompletat), ramane gol ('')."""
    v = rec.get(name, 0) or 0
    if v == 0:
        return ''
    return str(int(v))


def dbf_record_to_A_row(rec, adm, up):
    # LX1/LXA1 si LX2/LXA2 = "lucrari executate anterior" (cod lucrare + cifra).
    # In multe exporturi A/S aceste doua coloane (lx1, lx2) nu sunt completate,
    # desi DBF-ul original poate avea valori acolo -> la reconversia A+S -> DBF
    # ele vor iesi goale daca lipsesc din sursa (nu exista alta cale de a le
    # reconstitui doar din A si S). Validat pe caz real (UP Composesorat
    # Brazesti): 248/249 campuri identice cu DBF-ul original BIOSILV, singura
    # diferenta fiind exact lx1/lx2 nefolosite in A.xlsx sursa.
    ua_raw = rec['_raw_UA']  # 5 caractere: NNN + 2 caractere subparcela
    ua1 = int(ua_raw[0:3])
    ua2 = ua_raw[3:5]
    sba = ''.join(rec.get(f'SBA{i}', '') for i in range(1, 6))
    lx1 = (rec.get('LX1', '') + rec.get('LXA1', '')).strip()
    lx2 = (rec.get('LX2', '') + rec.get('LXA2', '')).strip()
    cel = _fmt_pair_block([(rec.get(f'SPT{i}', ''), rec.get(f'P{i}', 0)) for i in range(2, 7)])
    mel = _fmt_pair_block([(rec.get(f'SU{i}', ''), rec.get(f'PS{i}', 0)) for i in range(2, 7)])
    row = dict(
        adm=adm, up=up, ua1=ua1, ua2=ua2,
        dec1=rec.get('DEC1', 0), dec2=rec.get('DEC2', 0), dec3=rec.get('DEC3', 0),
        sup=rec.get('SUP', ''), ff=rec.get('F', 0), spr=rec.get('SPR', 0), fls=rec.get('FLS', 0),
        gf=rec.get('GF', 0), fct1=rec.get('FCT1', ''), fct2=rec.get('FCT2', ''), fct3=rec.get('FCT3', ''),
        rlf=_numcode(rec, 'RLF'), cnf=rec.get('CNF', ''), exp=rec.get('EXP', ''), inc=rec.get('INC', 0),
        alt1=rec.get('ALT1', 0), alt2=rec.get('ALT2', 0), sol=_blankable_num(rec, 'SOL'), erz=rec.get('ERZ', ''),
        flr=rec.get('FLR', 0), ts=rec.get('TS', 0), inv=rec.get('INV', 0), tp=rec.get('TP', 0),
        crti=rec.get('CRT', ''), pol1=_numcode(rec, 'POL1'), pol2=_numcode(rec, 'POL2'), pol3=rec.get('POL3', ''),
        lit=rec.get('LIT', 0), drm=rec.get('DRM', ''), dst=rec.get('DST', 0), str=rec.get('STR', 0),
        cns=rec.get('CNS', 0), clp=rec.get('CLP', 0), ta=rec.get('TA', 0), reg=rec.get('REG', 0),
        te=rec.get('TE', 0), ex=rec.get('EX', 0), urg=rec.get('URG', 0), prm=rec.get('PRM', 0),
        nin=rec.get('NIN', 0), nid=rec.get('NID', 0), lx1=lx1, lx2=lx2, lp1=rec.get('LP1', ''),
        lp2=rec.get('LP2', ''), lp3=rec.get('LP3', ''), dc1=rec.get('DC1', ''), dc2=rec.get('DC2', ''),
        dc3=rec.get('DC3', ''), dc4=rec.get('DC4', ''), sba=sba, so=rec.get('SO', 0), mr=rec.get('MR', 0),
        ds=rec.get('DS', 0), cel1=rec.get('SPT1', ''), cpr1=rec.get('P1', 0), cel=cel, vs=rec.get('VS', 0),
        soc=rec.get('SOC', 0), rs=rec.get('RS', 0), mel1=rec.get('SU1', ''), mpr1=rec.get('PS1', 0), mel=mel,
    )
    return [row[h] for h in A_HEADER]


def dbf_record_to_S_rows(rec, adm, up):
    ua_raw = rec['_raw_UA']
    ua1 = int(ua_raw[0:3])
    ua2 = ua_raw[3:5]
    spr = rec.get('SPR', 0) or 0
    out = []
    for i in range(1, 9):
        elm = rec.get(f'ELM{i}', '')
        if not elm:
            continue
        prp = rec.get(f'PRP{i}', 0) or 0
        row = dict(
            adm=adm, up=up, ua1=ua1, ua2=ua2, elm=elm, mrg=rec.get(f'MRG{i}', 0),
            prov=rec.get(f'PROV{i}', ''), prp=prp, spf=round(spr * prp / 10.0, 2),
            vrt=rec.get(f'VRT{i}', 0), ams=rec.get(f'AMS{i}', 0), elg=rec.get(f'ELG{i}', 0),
            vit=rec.get(f'VIT{i}', 0), tel=rec.get(f'TEL{i}', 0), cal=rec.get(f'CAL{i}', 0),
            pex1=rec.get(f'PEX1{i}', 0), pex2=rec.get(f'PEX2{i}', 0), pex3=rec.get(f'PEX3{i}', 0),
            dm=rec.get(f'DM{i}', 0), dmc=rec.get(f'DM{i}', 0), hm=rec.get(f'HM{i}', 0),
            hmc=rec.get(f'HM{i}', 0), m=rec.get(f'M{i}', 0), hsi=0, cp=rec.get(f'CP{i}', 0),
            cpc=rec.get(f'CP{i}', 0), vol=rec.get(f'VOL{i}', 0), crs=0, crsc=rec.get(f'CREST{i}', 0),
            vexpr=0, vexra=0, vexcu=0, vexig=0, vextc=0,
        )
        out.append([row[h] for h in S_HEADER])
    return out


def dbf_to_A_S(dbf_path, adm, up, out_a_path, out_s_path):
    fields, records, _ = read_dbf(dbf_path)
    a_rows = [dbf_record_to_A_row(r, adm, up) for r in records]
    s_rows = []
    for r in records:
        s_rows.extend(dbf_record_to_S_rows(r, adm, up))
    _write_xlsx(out_a_path, 'A', A_HEADER, a_rows)
    _write_xlsx(out_s_path, 'S', S_HEADER, s_rows)
    return len(a_rows), len(s_rows)


def _write_xlsx(path, sheetname, header, rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheetname
    ws.append(header)
    for r in rows:
        ws.append(r)
    wb.save(path)


# ---------------------------------------------------------------------------
# A.xlsx + S.xlsx  ->  BIOSILV .DBF
# ---------------------------------------------------------------------------
def _read_xlsx_rows(path, expected_header):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    header = [str(h).strip() for h in rows[0]]
    data = rows[1:]
    idx = {h: header.index(h) for h in expected_header if h in header}
    out = []
    for row in data:
        if row is None or all(v is None for v in row):
            continue
        d = {h: row[idx[h]] if idx[h] < len(row) else None for h in expected_header}
        out.append(d)
    return out


def _s(v):
    return '' if v is None else str(v).strip()


def _n(v):
    if v is None or v == '':
        return 0
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0


def _split_pair_block(block, n_pairs=5):
    """Inversul lui _fmt_pair_block: sir de n_pairs*4 caractere -> lista (cod,procent)."""
    block = (block or '')
    block = block.ljust(n_pairs * 4)
    out = []
    for i in range(n_pairs):
        chunk = block[i*4:(i+1)*4]
        cod = chunk[0:3].strip()
        try:
            pct = int(chunk[3:4]) if chunk[3:4].strip() != '' else 0
        except ValueError:
            pct = 0
        out.append((cod, pct))
    return out


def a_s_to_dbf(a_path, s_path, template_dbf_path, out_dbf_path):
    """Construieste un DBF BIOSILV nou, folosind structura de campuri din template_dbf_path
    (necesar ca soft-ul BIOSILV sa recunoasca fisierul) si datele din A.xlsx / S.xlsx."""
    fields, _, header_extra = read_dbf(template_dbf_path)

    a_rows = _read_xlsx_rows(a_path, A_HEADER)
    s_rows = _read_xlsx_rows(s_path, S_HEADER)

    # grupeaza randurile S dupa (adm,up,ua1,ua2)
    from collections import defaultdict
    s_by_ua = defaultdict(list)
    for r in s_rows:
        key = (r['adm'], r['up'], r['ua1'], r['ua2'])
        s_by_ua[key].append(r)

    out_records = []
    for a in a_rows:
        key = (a['adm'], a['up'], a['ua1'], a['ua2'])
        elems = s_by_ua.get(key, [])

        ua_num = int(_n(a['ua1']))
        ua2_raw = '' if a['ua2'] is None else str(a['ua2'])
        if len(ua2_raw) < 2:
            ua2_raw = ua2_raw.ljust(2)  # subparcela pe 1 caracter -> completat cu spatiu la dreapta
        ua_text = f"{ua_num:03d}{ua2_raw[:2]}"

        rec = {}
        rec['UA'] = ua_text
        rec['SUP'] = _s(a['sup'])
        rec['F'] = int(_n(a['ff']))
        rec['SPR'] = _n(a['spr'])
        rec['FLS'] = int(_n(a['fls']))
        rec['GF'] = int(_n(a['gf']))
        rec['FCT1'] = _s(a['fct1']); rec['FCT2'] = _s(a['fct2']); rec['FCT3'] = _s(a['fct3'])
        rlf_v = a['rlf']
        rec['RLF'] = str(int(rlf_v)).zfill(2) if isinstance(rlf_v, (int, float)) else _s(rlf_v)
        rec['CNF'] = _s(a['cnf']); rec['EXP'] = _s(a['exp'])
        rec['INC'] = int(_n(a['inc']))
        rec['ALT1'] = _n(a['alt1']); rec['ALT2'] = _n(a['alt2'])
        rec['SOL'] = _s(a['sol']); rec['ERZ'] = _s(a['erz'])
        rec['FLR'] = int(_n(a['flr'])); rec['TS'] = int(_n(a['ts'])); rec['INV'] = int(_n(a['inv']))
        rec['TP'] = int(_n(a['tp'])); rec['CRT'] = _s(a['crti'])
        rec['POL1'] = _s(a['pol1']) or '0'; rec['POL2'] = _s(a['pol2']) or '0'; rec['POL3'] = _s(a['pol3'])
        rec['LIT'] = int(_n(a['lit'])); rec['DRM'] = _s(a['drm'])
        rec['DST'] = int(_n(a['dst'])); rec['STR'] = int(_n(a['str']))
        rec['CNS'] = _n(a['cns']); rec['CLP'] = int(_n(a['clp']))
        rec['TA'] = int(_n(a['ta'])); rec['REG'] = int(_n(a['reg'])); rec['TE'] = int(_n(a['te']))
        rec['EX'] = int(_n(a['ex'])); rec['URG'] = int(_n(a['urg'])); rec['PRM'] = int(_n(a['prm']))
        rec['NIN'] = int(_n(a['nin'])); rec['NID'] = int(_n(a['nid']))
        lx1 = _s(a['lx1']); rec['LX1'] = lx1[:2]; rec['LXA1'] = lx1[2:3]
        lx2 = _s(a['lx2']); rec['LX2'] = lx2[:2]; rec['LXA2'] = lx2[2:3]
        rec['LP1'] = _s(a['lp1']); rec['LP2'] = _s(a['lp2']); rec['LP3'] = _s(a['lp3'])
        rec['DC1'] = _s(a['dc1']); rec['DC2'] = _s(a['dc2']); rec['DC3'] = _s(a['dc3']); rec['DC4'] = _s(a['dc4'])
        sba = _s(a['sba'])
        for i in range(1, 6):
            rec[f'SBA{i}'] = sba[i-1] if i-1 < len(sba) else ''
        rec['SO'] = int(_n(a['so'])); rec['MR'] = int(_n(a['mr'])); rec['DS'] = int(_n(a['ds']))
        rec['SPT1'] = _s(a['cel1']); rec['P1'] = int(_n(a['cpr1']))
        for i, (cod, pct) in enumerate(_split_pair_block(a['cel']), start=2):
            rec[f'SPT{i}'] = cod; rec[f'P{i}'] = pct
        rec['VS'] = int(_n(a['vs']))
        rec['SOC'] = int(_n(a['soc'])); rec['RS'] = int(_n(a['rs']))
        rec['SU1'] = _s(a['mel1']); rec['PS1'] = int(_n(a['mpr1']))
        for i, (cod, pct) in enumerate(_split_pair_block(a['mel']), start=2):
            rec[f'SU{i}'] = cod; rec[f'PS{i}'] = pct
        rec['DEC1'] = int(_n(a['dec1'])); rec['DEC2'] = int(_n(a['dec2'])); rec['DEC3'] = int(_n(a['dec3']))

        # elemente de arboret (din S) -> ELM1..ELM8
        for i in range(1, 9):
            if i <= len(elems):
                e = elems[i-1]
                rec[f'ELM{i}'] = _s(e['elm'])
                rec[f'MRG{i}'] = int(_n(e['mrg']))
                rec[f'VRT{i}'] = int(_n(e['vrt']))
                rec[f'PRP{i}'] = int(_n(e['prp']))
                rec[f'DM{i}'] = int(_n(e['dm']))
                rec[f'HM{i}'] = int(_n(e['hm']))
                rec[f'M{i}'] = int(_n(e['m']))
                rec[f'CP{i}'] = int(_n(e['cp']))
                rec[f'AMS{i}'] = int(_n(e['ams']))
                rec[f'ELG{i}'] = int(_n(e['elg']))
                rec[f'VIT{i}'] = int(_n(e['vit']))
                rec[f'TEL{i}'] = int(_n(e['tel']))
                rec[f'CAL{i}'] = int(_n(e['cal']))
                rec[f'VOL{i}'] = int(_n(e['vol']))
                rec[f'CRS{i}'] = _n(e['crs'])
                rec[f'PEX1{i}'] = int(_n(e['pex1']))
                rec[f'PEX2{i}'] = int(_n(e['pex2']))
                rec[f'PEX3{i}'] = int(_n(e['pex3']))
                rec[f'PROV{i}'] = _s(e['prov'])
                rec[f'CREST{i}'] = _n(e['crsc'])
            else:
                rec[f'ELM{i}'] = ''
                for fn in ['MRG','VRT','PRP','DM','HM','M','CP','AMS','ELG','VIT','TEL','CAL','VOL','PEX1','PEX2','PEX3']:
                    rec[f'{fn}{i}'] = 0
                rec[f'CRS{i}'] = 0.0
                rec[f'PROV{i}'] = ''
                rec[f'CREST{i}'] = 0.0

        out_records.append(rec)

    write_dbf(out_dbf_path, fields, out_records, header_extra=header_extra)
    return len(out_records)
