# -*- coding: utf-8 -*-
"""보고서 제작기 엔진 — AI 없이 프로그램만으로 동작한다.

  1) 새 보고서 만들기   : 원본 파일들 (+ 샘플 보고서) → 보고서
  2) 기존 보고서에 추가 : 이미 만든 보고서 + 추가 원고 → 같은 양식으로 이어 붙인 새 파일

샘플 방식
  샘플(기존) 보고서에서 반복되는 묶음(「□ 11호 (2026.11.09)」 같은 제목 + 그 아래 표·문단)을 찾아,
  마지막 묶음을 복제한 뒤 새 원본과 '줄 단위'로 맞춰 본다.
    - 글자가 같은 줄(대상수·건수 같은 항목명)은 그대로 둔다.
    - 값이 들어가는 줄은 원본 값을 칸 순서대로 그대로 넣는다.
    - 원본 줄이 더 많으면 같은 모양의 줄을 복제하고, 적으면 남는 줄을 지운다.
  제목의 호 번호는 하나 올리고(원본에 Vol.12 같은 번호가 있으면 그 번호), 날짜는 원본에서 찾은 첫 날짜로.
  샘플이 없으면 원본의 표·글을 그대로 옮긴 기본 워드 보고서를 만든다.

★ 원칙: 원고의 글자·숫자는 절대 바꾸지 않는다.
  - 값은 원본 셀·문단을 그대로 복사한다(다시 타이핑 없음). 엑셀 숫자는 셀 표시 형식대로(91.4%, 1,234) 옮긴다.
  - 다 만든 뒤 verify() 가 ① 원본 줄이 순서대로 그대로 들어갔는지 ② 칸·숫자 개수 ③ 기존 내용 보존을 대조한다.
  - 원본 파일·기존 보고서는 덮어쓰지 않는다(항상 새 파일).
"""
import copy
import difflib
import html
import io
import os
import re
import shutil
import subprocess
import sys
import zipfile
from collections import Counter
from datetime import datetime, date, time

HERE = os.path.dirname(os.path.abspath(__file__))
JOBS = os.path.join(HERE, '작업기록')
SRC_EXT = ('.xlsx', '.xlsm', '.docx', '.pptx', '.hwpx', '.hwp', '.pdf', '.txt')


def log(msg):
    print(msg)
    try:
        sys.stdout.flush()
    except Exception:
        pass


def clean(s):
    return re.sub(r'\s+', ' ', str(s)).strip()


def _check_readable(path):
    try:
        with open(path, 'rb') as f:
            f.read(1)
    except PermissionError:
        raise RuntimeError(f"'{os.path.basename(path)}' 이(가) 다른 프로그램에서 열려 있어 읽을 수 없습니다. 그 창을 닫고 다시 하세요.")


# ───────────────────────── 엑셀 셀 → 화면에 보이는 글자 ─────────────────────────

def cell_text(c):
    """엑셀에서 보이는 그대로의 글자(표시 형식 반영). 91.4% 가 0.914 로 바뀌는 일이 없게."""
    v = c.value
    if v is None:
        return ''
    if isinstance(v, str):
        return v
    if isinstance(v, bool):
        return 'TRUE' if v else 'FALSE'
    fmt = (c.number_format or 'General')
    if isinstance(v, (datetime, date)):
        f = fmt.lower()
        if isinstance(v, datetime) and ('h' in f) and (v.hour or v.minute or v.second):
            return v.strftime('%Y-%m-%d %H:%M')
        sep = '.' if 'yyyy.mm' in f or 'yy.mm' in f else ('/' if '/' in f else '-')
        return v.strftime(f'%Y{sep}%m{sep}%d')
    if isinstance(v, time):
        return v.strftime('%H:%M')
    if isinstance(v, (int, float)):
        f = fmt.split(';')[0]
        f = re.sub(r'\[[^\]]*\]|"[^"]*"|\\.|_.|\*.', '', f)
        dec = len(re.search(r'\.(0+)', f).group(1)) if re.search(r'\.(0+)', f) else 0
        if f == 'General' or not re.search(r'[0#]', f):
            if isinstance(v, float):
                return f'{v:.10g}' if v != int(v) else str(int(v))
            return str(v)
        if '%' in f:
            return f'{v * 100:,.{dec}f}%' if ',' in f else f'{v * 100:.{dec}f}%'
        if ',' in f:
            return f'{v:,.{dec}f}'
        return f'{v:.{dec}f}'
    return str(v)


# ───────────────────────── 원본 → 줄(행) 목록 ─────────────────────────
# 줄 = 칸 글자 튜플. 엑셀 행(빈 칸 제외) / 워드·PPT·한글 문단(1칸) / 표 행(글 있는 칸) / 텍스트 줄(1칸)

def _hwp_text(path):
    import tempfile
    fd, tmp = tempfile.mkstemp(suffix='.txt')
    os.close(fd)
    try:
        r = subprocess.run(['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File',
                            os.path.join(HERE, 'hwp-text.ps1'), path, tmp], capture_output=True,
                           creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
        if r.returncode != 0:
            err = (r.stderr or b'').decode('cp949', 'replace') + (r.stdout or b'').decode('cp949', 'replace')
            if '0x80030020' in err or '0x80030021' in err:
                raise RuntimeError(f"'{os.path.basename(path)}' 이(가) 한글에서 열려 있어 읽을 수 없습니다. 창을 닫고 다시 하세요.")
            first = next((ln.strip() for ln in err.splitlines() if ln.strip()), '알 수 없는 오류')
            raise RuntimeError(f'hwp 글자 추출 실패: {first}')
        txt = io.open(tmp, encoding='utf-8', errors='replace').read()
        return '\n'.join(ln for ln in txt.splitlines() if not ln.startswith('### ')).replace('{{OBJ}}', '')
    finally:
        try:
            os.remove(tmp)
        except Exception:
            pass


W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
HP = '{http://www.hancom.co.kr/hwpml/2011/paragraph}'


def _el_text(el, tag):
    return ''.join(t.text or '' for t in el.iter(tag))


def _xml_rows(root, p_tag, t_tag, tbl_tag, tr_tag, tc_tag):
    """XML 트리 → 줄 목록. 표 밖 문단은 1칸 줄, 표는 행마다 글 있는 칸들."""
    rows = []

    def walk(el):
        for ch in el:
            if ch.tag == tbl_tag:
                for tr in ch.iter(tr_tag):
                    cells = []
                    for tc in tr:
                        if tc.tag == tc_tag:
                            t = clean(' '.join(_el_text(p, t_tag) for p in tc.iter(p_tag)))
                            if t:
                                cells.append(t)
                    if cells:
                        rows.append(tuple(cells))
            elif ch.tag == p_tag:
                inner_tbl = next(ch.iter(tbl_tag), None)
                if inner_tbl is not None:           # 한글: 표가 문단 안에 들어 있다
                    walk(ch)
                    continue
                t = clean(_el_text(ch, t_tag))
                if t:
                    rows.append((t,))
            else:
                walk(ch)
    walk(root)
    return rows


def rows_of(path):
    """파일 → 줄 목록(빈 줄 제외)."""
    _check_readable(path)
    ext = os.path.splitext(path)[1].lower()
    rows = []
    if ext in ('.xlsx', '.xlsm'):
        import openpyxl
        wb = openpyxl.load_workbook(path, data_only=True)
        for ws in wb:
            if ws.sheet_state != 'visible':
                continue
            for r in ws.iter_rows():
                vals = tuple(t for t in (clean(cell_text(c)) for c in r) if t)
                if vals:
                    rows.append(vals)
    elif ext == '.docx':
        from lxml import etree
        with zipfile.ZipFile(path) as z:
            root = etree.fromstring(z.read('word/document.xml'))
        rows = _xml_rows(root, W + 'p', W + 't', W + 'tbl', W + 'tr', W + 'tc')
    elif ext == '.pptx':
        from lxml import etree
        with zipfile.ZipFile(path) as z:
            names = sorted((n for n in z.namelist() if re.match(r'ppt/slides/slide\d+\.xml$', n)),
                           key=lambda n: int(re.findall(r'\d+', n)[-1]))
            for n in names:
                root = etree.fromstring(z.read(n))
                rows += _xml_rows(root, A + 'p', A + 't', A + 'tbl', A + 'tr', A + 'tc')
    elif ext == '.hwpx':
        from lxml import etree
        with zipfile.ZipFile(path) as z:
            for n in sorted(z.namelist()):
                if re.search(r'section\d+\.xml$', n, re.I):
                    root = etree.fromstring(z.read(n))
                    rows += _xml_rows(root, HP + 'p', HP + 't', HP + 'tbl', HP + 'tr', HP + 'tc')
    elif ext == '.hwp':
        rows = [(clean(x),) for x in _hwp_text(path).splitlines() if clean(x)]
    elif ext == '.pdf':
        import fitz
        for pg in fitz.open(path):
            rows += [(clean(x),) for x in pg.get_text().splitlines() if clean(x)]
    else:
        txt = None
        for enc in ('utf-8-sig', 'cp949'):
            try:
                txt = io.open(path, encoding=enc).read()
                break
            except UnicodeDecodeError:
                continue
        if txt is None:
            txt = io.open(path, encoding='utf-8', errors='replace').read()
        rows = [(clean(x),) for x in txt.splitlines() if clean(x)]
    return rows


def units_of(path):
    return [c for r in rows_of(path) for c in r]


NUM = re.compile(r'(?<![\w.])[-+]?\d[\d,]*(?:[.:/\-]\d+)*%?')


def numbers(units):
    out = []
    for u in units:
        out += NUM.findall(re.sub(r'https?://\S+', ' ', u))
    return out


# ───────────────────────── 워드(docx) 모델 ─────────────────────────

def _txt(el):
    return ''.join(x.text or '' for x in el.iter(W + 't'))


def _settext(el, t):
    ts = list(el.iter(W + 't'))
    if not ts:
        return False
    ts[0].text = t
    ts[0].set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    for x in ts[1:]:
        x.text = ''
    return True


class Row:
    """보고서 안의 한 줄: 문단 하나(칸 1개) 또는 표 행 하나(글 있는 칸들)."""
    __slots__ = ('el', 'kind', 'slots')

    def __init__(self, el, kind, slots):
        self.el, self.kind, self.slots = el, kind, slots

    def texts(self):
        return tuple(clean(_txt(s)) for s in self.slots)


def _tr_slots(tr):
    return [tc for tc in tr.findall(W + 'tc') if clean(_txt(tc))]


def doc_rows(elements):
    rows = []
    for e in elements:
        if e.tag == W + 'p':
            if clean(_txt(e)):
                rows.append(Row(e, 'p', [e]))
        elif e.tag == W + 'tbl':
            for tr in e.findall(W + 'tr'):
                sl = _tr_slots(tr)
                if sl:
                    rows.append(Row(tr, 'tr', sl))
    return rows


# ───────────────────────── 묶음(섹션) 찾기 ─────────────────────────

DATE = re.compile(r'(\d{4})([./-])\s?(\d{1,2})\2\s?(\d{1,2})')


def _shape(t):
    return re.sub(r'\d+', '#', clean(t))


def find_sections(body_items):
    """body 요소 목록 → (제목 문단 인덱스 목록). 숫자만 다르고 모양이 같은 짧은 문단이 2번 이상 나오면 그게 묶음 제목."""
    cands = {}
    for i, e in enumerate(body_items):
        if e.tag != W + 'p':
            continue
        t = clean(_txt(e))
        if not t or len(t) > 60 or not re.search(r'\d', t):
            continue
        cands.setdefault(_shape(t), []).append(i)
    best = [v for v in cands.values() if len(v) >= 2]
    if not best:
        return []
    return max(best, key=lambda v: (len(v), -v[0]))


KDATE = re.compile(r'(\d{4})\s*년\s*(\d{1,2})\s*월(?:\s*(\d{1,2})\s*일)?')


def source_date(src_rows):
    """원본에서 처음 나오는 날짜 (년, 월, 일 또는 None)."""
    for r in src_rows:
        for c in r:
            m = DATE.search(c)
            if m:
                return int(m.group(1)), int(m.group(3)), int(m.group(4))
            m = KDATE.search(c)
            if m:
                return int(m.group(1)), int(m.group(2)), int(m.group(3)) if m.group(3) else None
    return None


def heading_for(tmpl_heading, prev_headings, src_rows, fallback_no):
    """새 묶음 제목: 템플릿(마지막 묶음) 제목에서 숫자만 바꾼다. 반환 (제목, 호번호, 안내문 목록)
      - 날짜(2026.11.09 / 2026년 11월 (9일)) → 원본에서 찾은 첫 날짜, 쓰는 모양은 템플릿대로
      - 묶음마다 1씩 늘어나는 숫자 중 첫째 = 호 번호 → 원본의 Vol.12/제12호, 없으면 앞 번호+1
      - 그 밖에 1씩 늘어나는 숫자 → +1"""
    notes = []
    t = clean(tmpl_heading)
    flat = ' '.join(c for r in src_rows for c in r)
    m = re.search(r'Vol\.?\s*0*(\d+)', flat, re.I) or re.search(r'제\s*(\d+)\s*호', flat)
    no = int(m.group(1)) if m else fallback_no
    sd = source_date(src_rows)
    edits, date_spans = [], []
    for dm in DATE.finditer(t):
        date_spans.append(dm.span())
        if sd and sd[2]:
            # 샘플의 모양(구분자·띄어쓰기·두 자리)을 그대로 두고 년·월·일 숫자만 바꾼다
            txt, base = dm.group(0), dm.start()
            for g, val in ((4, sd[2]), (3, sd[1]), (1, sd[0])):
                a, b = dm.span(g)
                v = f'{val:02d}' if (g != 1 and len(dm.group(g)) == 2) else str(val)
                txt = txt[:a - base] + v + txt[b - base:]
            edits.append((dm.span(), txt))
        else:
            notes.append(f'제목 날짜를 원본에서 찾지 못했습니다 — 「{t}」의 날짜를 직접 고쳐 주세요.')
    for km in KDATE.finditer(t):
        if any(a <= km.start() < b for a, b in date_spans):
            continue
        date_spans.append(km.span())
        if sd:
            txt = km.group(0)
            txt = re.sub(r'^\d{4}', str(sd[0]), txt)
            txt = re.sub(r'(?<=년)(\s*)\d{1,2}(?=\s*월)', lambda x: x.group(1) + str(sd[1]), txt)
            if km.group(3) and sd[2]:
                txt = re.sub(r'(?<=월)(\s*)\d{1,2}(?=\s*일)', lambda x: x.group(1) + str(sd[2]), txt)
            edits.append((km.span(), txt))
        else:
            notes.append(f'제목 날짜를 원본에서 찾지 못했습니다 — 「{t}」의 날짜를 직접 고쳐 주세요.')
    # 날짜 밖 숫자 중 묶음마다 1씩 늘어나는 자리
    free = [x for x in re.finditer(r'\d+', t) if not any(a <= x.start() < b for a, b in date_spans)]
    seqs = []
    for h in prev_headings:
        hh = clean(h)
        ds = [x.span() for x in DATE.finditer(hh)] + [x.span() for x in KDATE.finditer(hh)]
        seqs.append([int(x.group()) for x in re.finditer(r'\d+', hh) if not any(a <= x.start() < b for a, b in ds)])
    first = True
    found_any = False
    for k, x in enumerate(free):
        seq = [q[k] for q in seqs if len(q) > k]
        inc = len(seq) >= 2 and all(b - a == 1 for a, b in zip(seq, seq[1:]))
        if not inc and not (len(prev_headings) == 1 and k == 0):
            continue
        found_any = True
        if first:
            edits.append((x.span(), str(no)))
            first = False
        else:
            edits.append((x.span(), str(int(x.group()) + 1)))
    if not found_any and free:
        notes.append(f'제목의 번호 자리를 알 수 없어 「{t}」의 번호는 그대로 두었습니다 — 직접 고쳐 주세요.')
    out = t
    for (a, b), txt in sorted(edits, key=lambda e: -e[0][0]):
        out = out[:a] + txt + out[b:]
    return out, no, notes


# ───────────────────────── 줄 맞추기 + 채우기 ─────────────────────────

def _is_value(c):
    return bool(re.search(r'\d', c)) or c.startswith('http') or len(c) > 40


def _key(cells):
    return tuple('#' if _is_value(c) else c for c in cells)


JOIN = '  '      # 칸이 모자라 한 문단에 이어 적을 때 칸 사이


def fill_section(clones, src_rows, boilerplate):
    """복제한 묶음(제목 제외 요소들)을 원본 줄로 채운다. 반환: 안내문 목록."""
    notes = []
    T = doc_rows(clones)
    S = list(src_rows)
    fixed = {id(r) for r in T if r.kind == 'p' and clean(' '.join(r.texts())) in boilerplate}
    kt = [_key(r.texts()) for r in T]
    ks = [_key(r) for r in S]
    sm = difflib.SequenceMatcher(None, kt, ks, autojunk=False)
    KEEP = object()
    pairs = []          # (Row 또는 None, 원본 줄 / None(지움) / KEEP(고정 문구 그대로))
    for op, i1, i2, j1, j2 in sm.get_opcodes():
        if op == 'equal':
            pairs += [(T[i1 + k], S[j1 + k]) for k in range(i2 - i1)]
            continue
        tb, used, last = T[i1:i2], set(), -1
        for s in S[j1:j2]:
            hit = next((k for k in range(last + 1, len(tb)) if k not in used and id(tb[k]) not in fixed
                        and len(tb[k].slots) == len(s)), None)
            if hit is not None:
                for k in range(last + 1, hit):
                    if k not in used:
                        pairs.append((tb[k], KEEP if id(tb[k]) in fixed else None))
                        used.add(k)
                pairs.append((tb[hit], s))
                used.add(hit)
                last = hit
            else:
                pairs.append((None, s))
        for k in range(len(tb)):
            if k not in used:
                pairs.append((tb[k], KEEP if id(tb[k]) in fixed else None))

    # 고정 문구(▼ 전송 오류 분석 같은 소제목)라도, 원본에 없고 그 아래 표 줄이 전부 빠지면 같이 지운다
    paired = {id(t) for t, s in pairs if t is not None and s is not None and s is not KEEP}
    orphan = set()
    for i, r in enumerate(T):
        if r.kind != 'p' or id(r) in paired:
            continue
        under = []
        for x in T[i + 1:]:
            if id(x) in paired or (x.kind == 'p' and under and under[-1].kind == 'tr'):
                break
            under.append(x)
        if under and any(x.kind == 'tr' for x in under):
            orphan.add(id(r))
    row_of = {id(r.el): r for r in T}

    def proto_for(n, cursor):
        """새 줄의 본보기: ① 바로 앞 줄이 같은 칸 수면 그 줄 ② 같은 표의 같은 칸 수 줄 ③ 아무 데나 같은 칸 수 줄
        ④ 앞 줄이 표 안이면 그 표의 더 넓은 줄 ⑤ 문단(칸을 이어 적음)"""
        cur = row_of.get(id(cursor)) if cursor is not None else None
        if cur is not None and len(cur.slots) == n and id(cur) not in fixed:
            return cur, False
        same = [r for r in T if len(r.slots) == n and id(r) not in fixed]
        if cursor is not None and cursor.tag == W + 'tr':
            inner = [r for r in same if r.kind == 'tr' and r.el.getparent() is cursor.getparent()]
            if inner:
                return inner[-1], False
        if same:
            ps = [r for r in same if r.kind == 'p'] if n == 1 else []
            return (ps or same)[-1], False
        if cursor is not None and cursor.tag == W + 'tr':
            wider = [r for r in T if r.kind == 'tr' and r.el.getparent() is cursor.getparent() and len(r.slots) > n]
            if wider:
                return max(wider, key=lambda r: len(r.slots)), False
        allp = [r for r in T if r.kind == 'p']
        if cursor is None and allp:              # 묶음 맨 앞 → 첫 문단 모양(보통 본문)
            return allp[0], True
        ps = [r for r in allp if id(r) not in fixed] or allp
        return (ps[-1] if ps else T[-1]), True

    cursor = None
    first_el = clones[0] if clones else None
    for t, s in pairs:
        if t is not None and s is KEEP:
            if id(t) in orphan:                  # 아래 내용이 다 빠진 소제목 → 같이 지움
                if t.el.getparent() is not None:
                    t.el.getparent().remove(t.el)
            else:
                cursor = t.el
        elif t is not None and s is not None:
            for slot, v in zip(t.slots, s):
                _settext(slot, v)
            cursor = t.el
        elif t is None and s is not None:
            if not T:
                notes.append('샘플 묶음에 글 자리가 없어 원본을 넣지 못했습니다.')
                break
            proto, squash = proto_for(len(s), cursor)
            new = copy.deepcopy(proto.el)
            if new.tag == W + 'tr':
                if cursor is not None and cursor.tag == W + 'tr':
                    cursor.addnext(new)
                else:
                    nxt = cursor.getnext() if cursor is not None else None
                    if nxt is not None and nxt.tag == W + 'tbl' and nxt.find(W + 'tr') is not None:
                        nxt.find(W + 'tr').addprevious(new)
                    else:
                        proto.el.getparent().append(new)
            else:
                if cursor is not None and cursor.tag == W + 'tr':
                    cursor.getparent().addnext(new)
                elif cursor is not None:
                    cursor.addnext(new)
                elif first_el is not None:
                    first_el.addprevious(new)
                else:
                    proto.el.addprevious(new)
            slots = _tr_slots(new) if new.tag == W + 'tr' else [new]
            if squash or len(slots) < len(s):
                _settext(slots[0], JOIN.join(s))
                for x in slots[1:]:
                    _settext(x, '')
                if len(s) > 1:
                    notes.append('샘플에 맞는 칸이 없어 한 줄로 이어 적은 원본: ' + JOIN.join(s)[:80])
            else:
                for slot, v in zip(slots, s):
                    _settext(slot, v)
                for x in slots[len(s):]:
                    _settext(x, '')
            cursor = new
        elif t is not None and s is None:
            if t.kind == 'tr' or id(t) not in fixed or id(t) in orphan:
                if t.el.getparent() is not None:
                    t.el.getparent().remove(t.el)
    for e in list(clones):
        if e.tag == W + 'tbl' and e.find(W + 'tr') is None and e.getparent() is not None:
            e.getparent().remove(e)
    return notes


# ───────────────────────── 매호 이미지 ─────────────────────────
# 이미지 폴더에서 이름의 「N호」(예: 26_12호.png, 26_12호_2.jpg)로 그 호의 그림을 찾아 넣는다.
#  - '호' 바로 앞 숫자 = 호 번호, 그 앞 숫자(26_) = 연도, '호' 뒤 숫자(_2) = 같은 호 안의 순서
#  - 가로 2000px 보다 크면 줄인다. PNG 는 PNG 그대로(용량 큰 사진 PNG 만 JPG 로)
#  - 샘플 묶음에 그림 자리가 있으면 그 자리·폭 그대로 그림만 바꾸고, 없으면 제목 바로 아래에 본문 폭으로 넣는다

IMG_EXT = ('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp', '.tif', '.tiff')
MAX_PX = 2000
BIG_PNG = 1.5 * 1024 * 1024
WP = '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'
RNS = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'


def image_key(name):
    """파일 이름 → (호, 연도 두 자리 또는 None, 순서). 「N호」가 없으면 None."""
    stem = os.path.splitext(os.path.basename(name))[0]
    m = re.search(r'(?<!\d)(\d+)\s*호', stem)
    if not m:
        return None
    ym = re.search(r'(\d{4}|\d{2})\D*$', stem[:m.start()])
    om = re.search(r'\d+', stem[m.end():])
    return int(m.group(1)), (int(ym.group(1)) % 100 if ym else None), (int(om.group()) if om else 0)


def find_images(img_dir, no, year=None):
    """img_dir 에서 no 호 이미지들(순서대로). 반환 (경로 목록, 안내문 목록)."""
    if not img_dir or not os.path.isdir(img_dir):
        return [], []
    cand = []
    for f in os.listdir(img_dir):
        if f.lower().endswith(IMG_EXT):
            k = image_key(f)
            if k and k[0] == no:
                cand.append((k, f))
    notes = []
    years = {k[1] for k, _ in cand}
    if len(years) > 1:
        known = sorted(y for y in years if y is not None)
        pick = year % 100 if year is not None and year % 100 in years else (known[-1] if known else None)
        notes.append(f'{no}호 이미지가 여러 해 것이 있어({", ".join(f for _, f in sorted(cand))}) '
                     f'{pick}년 것을 넣었습니다.')
        cand = [c for c in cand if c[0][1] == pick]
    cand.sort(key=lambda c: (c[0][2], c[1]))
    return [os.path.join(img_dir, f) for _, f in cand], notes


def prep_image(path, out_dir):
    """넣을 그림 준비: 크면 줄이고, 워드가 못 읽는 형식은 PNG 로. 반환 (경로, 가로, 세로, 안내문 또는 '')."""
    from PIL import Image
    im = Image.open(path)
    im.load()
    w0, h0 = im.size
    fmt = (im.format or '').upper()
    name = os.path.basename(path)
    alpha = im.mode in ('RGBA', 'LA', 'PA') or 'transparency' in im.info
    if alpha:
        try:
            alpha = im.convert('RGBA').getchannel('A').getextrema()[0] < 255
        except Exception:
            pass
    need = w0 > MAX_PX or fmt not in ('PNG', 'JPEG', 'GIF') or (fmt == 'JPEG' and im.mode == 'CMYK')
    big_png = fmt == 'PNG' and not alpha and os.path.getsize(path) > BIG_PNG
    if not need and not big_png:
        return path, w0, h0, ''
    if w0 > MAX_PX:
        im = im.resize((MAX_PX, max(1, round(h0 * MAX_PX / w0))), Image.LANCZOS)
    w, h = im.size
    os.makedirs(out_dir, exist_ok=True)
    stem = os.path.splitext(name)[0]
    what = [f'{w0}×{h0} → {w}×{h}'] if (w, h) != (w0, h0) else []
    if fmt == 'JPEG':
        out = os.path.join(out_dir, stem + '.jpg')
        im.convert('RGB').save(out, 'JPEG', quality=88, optimize=True)
    else:
        out = os.path.join(out_dir, stem + '.png')
        im.save(out, 'PNG', optimize=True)
        if fmt not in ('PNG', 'GIF'):
            what.append(f'{fmt or "이 형식"} → PNG')
        if not alpha and os.path.getsize(out) > BIG_PNG:
            jp = os.path.join(out_dir, stem + '.jpg')
            im.convert('RGB').save(jp, 'JPEG', quality=90, optimize=True)
            what.append(f'용량이 커서({os.path.getsize(out) / 1048576:.1f}MB) JPG로')
            out = jp
    return out, w, h, (f'{name}: ' + ', '.join(what)) if what else ''


def _pics(els):
    """요소들 안의 그림(w:drawing) 목록, 문서 순서대로."""
    return [dr for e in els for dr in e.iter(W + 'drawing') if next(dr.iter(A + 'blip'), None) is not None]


def _page_box(d):
    from docx.shared import Cm
    s = d.sections[-1]
    try:
        cw = int(s.page_width - s.left_margin - s.right_margin)
        ch = int(s.page_height - s.top_margin - s.bottom_margin - Cm(3))
    except Exception:
        cw, ch = int(Cm(16)), int(Cm(22))
    return cw, ch


def _fit(cx, w, h, max_h):
    cy = round(cx * h / w)
    if cy > max_h:
        cy = max_h
        cx = round(cy * w / h)
    return cx, cy


def _new_pic_par(d, path, w, h, like=None, width=None):
    """그림 한 장짜리 문단(width 없으면 본문 폭). like 문단이 있으면 그 문단 모양(가운데 정렬 등)을 따른다."""
    from docx.shared import Emu
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    cw, ch = _page_box(d)
    cx, cy = _fit(width or cw, w, h, ch)
    p = d.add_paragraph()
    p.add_run().add_picture(path, width=Emu(cx), height=Emu(cy))
    el = p._p
    el.getparent().remove(el)
    ppr = like.find(W + 'pPr') if like is not None else None
    if ppr is not None:
        old = el.find(W + 'pPr')
        if old is not None:
            el.remove(old)
        el.insert(0, copy.deepcopy(ppr))
    else:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    return el


def _top_of(el, els):
    """el 을 품은, els(묶음 맨 위 요소들) 중 하나."""
    ids = {id(x) for x in els}
    x = el
    while x is not None and id(x) not in ids:
        x = x.getparent()
    return x


def _unique_docpr(d, els):
    used = {int(x.get('id')) for x in d.element.body.iter(WP + 'docPr') if (x.get('id') or '').isdigit()}
    nxt = max(used or {0}) + 1
    seen = set()
    for e in els:
        for x in e.iter(WP + 'docPr'):
            if id(x) in seen:
                continue
            seen.add(id(x))
            x.set('id', str(nxt))
            nxt += 1


def _par_of(el):
    while el is not None and el.tag != W + 'p':
        el = el.getparent()
    return el


def _drop_pic(dr, body):
    """그림 하나를 뺀다. 글 없이 그림만 있던 문단은 문단째(표 칸의 마지막 문단은 남김)."""
    par = _par_of(dr)
    run = dr.getparent()
    run.remove(dr)
    if run.tag == W + 'r' and not len([c for c in run if c.tag != W + 'rPr']) and run.getparent() is not None:
        run.getparent().remove(run)
    if par is None or clean(_txt(par)) or _pics([par]) or par.getparent() is None:
        return
    holder = par.getparent()
    if holder.tag == W + 'tc' and len(holder.findall(W + 'p')) <= 1:
        return
    if par in body:
        body.remove(par)
    holder.remove(par)


def place_images(d, heading, body, imgs, label):
    """새 묶음(heading + body 요소들)에 그림을 넣는다. body 는 새 문단이 생기면 늘어난다.
    imgs: [(경로, 가로, 세로)]. 반환 안내문 목록."""
    notes = []
    pics = _pics(body)
    cw, ch = _page_box(d)
    if not imgs:
        if pics:
            for dr in pics:
                _drop_pic(dr, body)
            notes.append(f'{label} 이미지가 이미지 폴더에 없어, 샘플에 있던 그림을 빼 두었습니다 — 넣을 그림이 있으면 '
                         f'이름을 「…{label}.png」로 저장해 다시 만드세요.')
        return notes
    for dr, (path, w, h) in zip(pics, imgs):
        rid, _ = d.part.get_or_add_image(path)
        for b in dr.iter(A + 'blip'):
            b.set(RNS + 'embed', rid)
        ext = next(dr.iter(WP + 'extent'), None)
        cx0 = int(ext.get('cx')) if ext is not None else cw
        cx, cy = _fit(cx0, w, h, ch)
        if ext is not None:
            ext.set('cx', str(cx)); ext.set('cy', str(cy))
        for x in dr.iter(A + 'xfrm'):
            e2 = x.find(A + 'ext')
            if e2 is not None:
                e2.set('cx', str(cx)); e2.set('cy', str(cy))
    if len(pics) > len(imgs):
        for dr in pics[len(imgs):]:
            _drop_pic(dr, body)
    extra = imgs[len(pics):]
    if extra:
        if pics:
            last = pics[-1]                     # 샘플 마지막 그림 바로 뒤에, 같은 폭·같은 문단 모양으로
            anchor, like = _top_of(last, body), _par_of(last)
            ext = next(last.iter(WP + 'extent'), None)
            width = int(ext.get('cx')) if ext is not None else None
        else:
            anchor, like, width = heading, None, None
        for path, w, h in extra:
            el = _new_pic_par(d, path, w, h, like, width)
            if anchor is not None:
                anchor.addnext(el)
                body.insert(body.index(anchor) + 1 if anchor in body else 0, el)
            elif body:
                body[0].addprevious(el)
                body.insert(0, el)
            else:
                notes.append(f'{label} 이미지를 넣을 자리가 없어 넣지 못했습니다.')
                break
            anchor = el
    _unique_docpr(d, ([heading] if heading is not None else []) + body)
    return notes


def images_for(img_dir, no, year, job, label=None):
    """no 호 그림 찾기 + 준비. 반환 ([(경로, 가로, 세로)], 안내문 목록)."""
    paths, notes = find_images(img_dir, no, year)
    out = []
    for p in paths:
        try:
            q, w, h, n = prep_image(p, os.path.join(job, '이미지'))
            out.append((q, w, h))
            notes.append(f'이미지 {label or f"{no}호"} ← {os.path.basename(p)}' + (f' ({n.split(": ", 1)[1]})' if n else ''))
        except Exception as e:
            notes.append(f'이미지 {os.path.basename(p)} 을(를) 열 수 없어 넣지 못했습니다: {e}')
    return out, notes


def _issue_no(src_rows, path):
    flat = ' '.join(c for r in src_rows for c in r)
    m = re.search(r'Vol\.?\s*0*(\d+)', flat, re.I) or re.search(r'제\s*(\d+)\s*호', flat)
    if m:
        return int(m.group(1))
    k = image_key(path)
    return k[0] if k else None


# ───────────────────────── 워드: 샘플 기반 ─────────────────────────

def _body_items(d):
    return [e for e in d.element.body if e.tag != W + 'sectPr']


def _sections(d):
    items = _body_items(d)
    heads = find_sections(items)
    secs = []
    for n, h in enumerate(heads):
        end = heads[n + 1] if n + 1 < len(heads) else len(items)
        # 문서 맨 끝 빈 문단은 묶음에 넣지 않는다
        while end - 1 > h and items[end - 1].tag == W + 'p' and not clean(_txt(items[end - 1])) \
                and n + 1 == len(heads):
            end -= 1
        secs.append((h, end))
    return items, secs


def _boilerplate(items, secs):
    """묶음 2곳 이상에 똑같이 나오는 문단 = 고정 문구(지우지 않음)."""
    cnt = Counter()
    for h, e in secs:
        seen = set()
        for x in items[h + 1:e]:
            if x.tag == W + 'p':
                t = clean(_txt(x))
                if t:
                    seen.add(t)
        cnt.update(seen)
    return {t for t, c in cnt.items() if c >= 2}


def docx_add_sections(sample, sources, out, mode='add', img_dir='', job=None):
    """sample(docx)의 마지막 묶음 모양으로 sources 마다 묶음을 하나씩 만든다.
    mode='add' : 기존 묶음 뒤에 이어 붙임 / mode='new' : 기존 묶음은 지우고 새로.
    img_dir 가 있으면 묶음마다 그 호의 그림(…N호.png)을 넣는다."""
    import docx
    d = docx.Document(sample)
    items, secs = _sections(d)
    notes = []
    if secs:
        h, e = secs[-1]
        heading_el, body_els = items[h], items[h + 1:e]
        prev_heads = [clean(_txt(items[a])) for a, _ in secs]
        boiler = _boilerplate(items, secs)
        last_no = None
        for k in range(len(re.findall(r'\d+', prev_heads[-1]))):
            seq = [int(re.findall(r'\d+', x)[k]) for x in prev_heads if len(re.findall(r'\d+', x)) > k]
            if len(seq) >= 2 and all(b - a == 1 for a, b in zip(seq, seq[1:])):
                last_no = seq[-1]
                break
        anchor = items[e - 1]
    else:
        # 묶음 제목을 못 찾으면 문서 전체를 한 묶음으로 본다(표지 없이)
        notes.append('샘플에서 반복되는 묶음 제목(예: □ 1호)을 찾지 못해, 샘플 전체를 한 묶음 양식으로 썼습니다.')
        heading_el, body_els, prev_heads, boiler, last_no = None, items, [], set(), None
        anchor = items[-1] if items else None
    proto_heading = copy.deepcopy(heading_el) if heading_el is not None else None
    proto_body = [copy.deepcopy(x) for x in body_els]

    if mode == 'new':
        first = secs[0][0] if secs else 0
        for x in items[first:]:
            x.getparent().remove(x)
        rest = [x for x in d.element.body if x.tag != W + 'sectPr']
        anchor = rest[-1] if rest else None
        if secs:
            notes.append('표지·사업개요 등 첫 묶음 앞부분은 샘플 그대로입니다. 이번 보고서에 맞게 확인하세요.')

    srcs = []
    for s in sources:
        r = rows_of(s)
        if not r:
            notes.append(f'{os.path.basename(s)}: 읽을 글자가 없습니다(스캔 이미지일 수 있음).')
            continue
        srcs.append((s, r))
    added, heads_made = [], []
    no = (last_no or 0) if mode == 'add' else 0
    for s, r in srcs:
        no += 1
        new_els = []
        if proto_heading is not None:
            hd = copy.deepcopy(proto_heading)
            text, used_no, hn = heading_for(_txt(proto_heading), prev_heads or [_txt(proto_heading)], r, no)
            no = used_no
            _settext(hd, text)
            heads_made.append(text)
            notes += hn
            new_els.append(hd)
        body = [copy.deepcopy(x) for x in proto_body]
        holder = d.element.body.makeelement(W + 'tmp', {})     # 채우는 동안 잠시 담아 둘 그릇
        for x in body:
            holder.append(x)
        notes += [f'{os.path.basename(s)}: {n}' for n in fill_section(list(holder), r, boiler)]
        new_els += list(holder)
        for x in new_els:
            if anchor is None:
                d.element.body.insert(len(d.element.body) - 1, x)
            else:
                anchor.addnext(x)
            anchor = x
        if img_dir:
            sd = source_date(r)
            imgs, inotes = images_for(img_dir, no, sd[0] if sd else None, job or JOBS)
            notes += inotes
            hd_el = new_els[0] if proto_heading is not None else None
            body_list = new_els[1:] if proto_heading is not None else list(new_els)
            notes += place_images(d, hd_el, body_list, imgs, f'{no}호')
            anchor = body_list[-1] if body_list else anchor
        added.append((s, no))
    d.save(out)
    return added, heads_made, notes


# ───────────────────────── 워드: 샘플 없이 기본 보고서 ─────────────────────────

def docx_plain(sources, out, title='', img_dir='', job=None, notes=None):
    import docx
    from docx.shared import Pt
    from docx.oxml.ns import qn
    d = docx.Document()
    st = d.styles['Normal']
    st.font.name = '맑은 고딕'
    st.font.size = Pt(10)
    st.element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')
    heads = []
    if title:
        d.add_heading(title, level=0)
        heads.append(title)
    for s in sources:
        name = os.path.splitext(os.path.basename(s))[0]
        if not s.endswith('붙여넣은글.txt'):
            hp = d.add_heading(name, level=1)
            heads.append(name)
            if img_dir:
                r = rows_of(s)
                no = _issue_no(r, s)
                if no is None:
                    if notes is not None:
                        notes.append(f'{os.path.basename(s)}: 몇 호인지 알 수 없어 이미지를 넣지 못했습니다 '
                                     f'(원고에 Vol.N·제N호, 또는 파일 이름에 N호가 있어야 합니다).')
                else:
                    sd = source_date(r)
                    imgs, inotes = images_for(img_dir, no, sd[0] if sd else None, job or JOBS)
                    if notes is not None:
                        notes += inotes
                        if not imgs:
                            notes.append(f'{no}호 이미지가 이미지 폴더에 없습니다.')
                    anchor = hp._p
                    for path, w, h in imgs:
                        el = _new_pic_par(d, path, w, h)
                        anchor.addnext(el)
                        anchor = el
        ext = os.path.splitext(s)[1].lower()
        if ext in ('.xlsx', '.xlsm'):
            import openpyxl
            wb = openpyxl.load_workbook(s, data_only=True)
            vis = [ws for ws in wb if ws.sheet_state == 'visible']
            for ws in vis:
                grid = [[clean(cell_text(c)) for c in r] for r in ws.iter_rows()]
                grid = [r for r in grid if any(r)]
                if not grid:
                    continue
                cols = [j for j in range(max(len(r) for r in grid)) if any(j < len(r) and r[j] for r in grid)]
                if len(vis) > 1:
                    d.add_heading(ws.title, level=2)
                    heads.append(ws.title)
                t = d.add_table(rows=len(grid), cols=len(cols))
                t.style = 'Table Grid'
                for i, r in enumerate(grid):
                    for k, j in enumerate(cols):
                        t.cell(i, k).text = r[j] if j < len(r) else ''
        elif ext == '.docx':
            src = docx.Document(s)
            for el in src.element.body:
                if el.tag in (W + 'p', W + 'tbl'):
                    d.element.body.insert(len(d.element.body) - 1, copy.deepcopy(el))
        else:
            for r in rows_of(s):
                d.add_paragraph('  '.join(r))
    d.save(out)
    return heads


# ───────────────────────── 엑셀: 시트 통째 복사 ─────────────────────────

def _copy_sheet(src_ws, wb, name, index):
    ws = wb.create_sheet(name, index=index)
    for row in src_ws.iter_rows():
        for c in row:
            n = ws.cell(row=c.row, column=c.column, value=c.value)
            if c.has_style:
                n.font = copy.copy(c.font); n.border = copy.copy(c.border); n.fill = copy.copy(c.fill)
                n.number_format = c.number_format; n.protection = copy.copy(c.protection)
                n.alignment = copy.copy(c.alignment)
    for m in src_ws.merged_cells.ranges:
        ws.merge_cells(str(m))
    for k, dim in src_ws.column_dimensions.items():
        ws.column_dimensions[k].width = dim.width
    for k, dim in src_ws.row_dimensions.items():
        if dim.height:
            ws.row_dimensions[k].height = dim.height
    ws.sheet_view.showGridLines = src_ws.sheet_view.showGridLines
    return ws


XL_IMG_PX = 800      # 엑셀 시트에 넣는 그림의 보이는 가로(px)


def xlsx_add_sheets(sample, sources, out, mode='add', img_dir='', job=None):
    """샘플 엑셀(호마다 시트) 뒤에 원본 엑셀의 시트를 그대로 복사해 붙인다. 시트 이름은 샘플 마지막 시트 이름의 번호만 바꿈.
    img_dir 가 있으면 그 호 첫 시트의 표 아래에 그림을 넣는다."""
    import openpyxl
    wb = openpyxl.load_workbook(sample)
    notes = []
    # 이름 모양(숫자만 다른)이 가장 많이 반복되는 시트들 = 호별 시트. 'Sheet12' 같은 빈 시트에 속지 않게.
    shapes = Counter(_shape(n) for n in wb.sheetnames if re.search(r'\d', n))
    top = shapes.most_common(1)[0][0] if shapes else None
    numbered = [(n, int(re.search(r'\d+', n).group())) for n in wb.sheetnames
                if re.search(r'\d', n) and _shape(n) == top]
    last_name, last_no = (numbered[-1] if numbered else (None, 0))
    if mode == 'new':
        notes.append('새 보고서(엑셀): 샘플의 시트들은 그대로 두고 원본 시트를 뒤에 붙였습니다.')
    after = wb.sheetnames.index(last_name) if last_name else len(wb.sheetnames) - 1
    added = []
    no = last_no
    for s in sources:
        if not s.lower().endswith(('.xlsx', '.xlsm')):
            notes.append(f'{os.path.basename(s)}: 엑셀 샘플에는 엑셀 원본만 붙일 수 있어 건너뛰었습니다.')
            continue
        r = rows_of(s)
        flat = ' '.join(c for x in r for c in x)
        m = re.search(r'Vol\.?\s*0*(\d+)', flat, re.I) or re.search(r'제\s*(\d+)\s*호', flat)
        no = int(m.group(1)) if m else no + 1
        src_wb = openpyxl.load_workbook(s)
        first_ws = None
        for ws in [w for w in src_wb if w.sheet_state == 'visible']:
            name = re.sub(r'\d+', str(no), last_name, count=1) if last_name else ws.title
            base_name, i = name, 2
            while name in wb.sheetnames:
                if mode == 'add' and base_name == name:
                    raise ValueError(f'「{name}」 시트가 이미 있습니다. 이미 넣은 원고인지 확인하세요.')
                name = f'{base_name}({i})'
                i += 1
            after += 1
            nw = _copy_sheet(ws, wb, name, after)
            first_ws = first_ws or nw
        if img_dir and first_ws is not None:
            from openpyxl.drawing.image import Image as XLImage
            sd = source_date(r)
            imgs, inotes = images_for(img_dir, no, sd[0] if sd else None, job or JOBS)
            notes += inotes
            if not imgs:
                notes.append(f'{no}호 이미지가 이미지 폴더에 없습니다.')
            row = first_ws.max_row + 2
            for path, w, h in imgs:
                im = XLImage(path)
                im.width = min(w, XL_IMG_PX)
                im.height = round(h * im.width / w)
                first_ws.add_image(im, f'A{row}')
                row += int(im.height / 20) + 2          # 기본 행 높이 20px 기준으로 다음 그림 자리
        added.append((s, no))
    wb.save(out)
    return added, [], notes


# ───────────────────────── 대조 검사 ─────────────────────────

def _strict(u):
    """짧거나 숫자·기호뿐인 칸('-', '0', '3.2%', '건수')은 '다른 글에 포함'으로 넘기지 않는다."""
    return len(u) < 6 or not re.search(r'[가-힣A-Za-z]{2}', u)


def _subseq_missing(src_rows, out_rows):
    """원본 줄들이 결과에 같은 순서로(칸 글자 그대로) 들어 있는지. 못 찾은 원본 줄 목록.
    결과에서 한 문단에 이어 적은 줄(칸 사이 두 칸 띄움)도 같은 줄로 본다."""
    miss, j = [], 0
    for r in src_rows:
        alt = (clean(JOIN.join(r)),)
        k = j
        while k < len(out_rows) and out_rows[k] != r and out_rows[k] != alt:
            k += 1
        if k < len(out_rows):
            j = k + 1
        else:
            miss.append(r)
    return miss


def verify(sources, result, base=None, extra_text='', order_rows=None):
    """원본 ↔ 결과 대조.
    order_rows: (원본 줄 목록, 결과 줄 목록) — 주면 '원본 줄이 순서대로 그대로' 검사도 한다."""
    res_units = units_of(result)
    base_units = units_of(base) if base else []
    res_cnt, base_cnt = Counter(res_units), Counter(base_units)
    avail = res_cnt - base_cnt
    hay = ' \n '.join(avail.elements())
    hay_ns = re.sub(r'\s+', '', hay)
    rows, src_all = [], []
    for s in sources:
        for row in rows_of(s):
            joined = clean(JOIN.join(row))
            if len(row) > 1 and avail[joined] > 0:          # 한 문단에 이어 적은 줄
                avail[joined] -= 1
                for u in row:
                    src_all.append(u)
                    rows.append([os.path.basename(s), u, '일치'])
                continue
            for u in row:
                src_all.append(u)
                rows.append([os.path.basename(s), u, None])
    for r in rows:
        if r[2]:
            continue
        if avail[r[1]] > 0:
            avail[r[1]] -= 1
            r[2] = '일치'
    for r in rows:
        if r[2]:
            continue
        u = r[1]
        if not _strict(u) and u in hay:
            r[2] = '포함'
        elif not _strict(u) and re.sub(r'\s+', '', u) in hay_ns:
            r[2] = '띄어쓰기 다름'
        else:
            r[2] = '없음'
    rows = [tuple(r) for r in rows]
    src_num = Counter(numbers(src_all))
    res_num = Counter(numbers(res_units)) - Counter(numbers(base_units))
    missing_num = src_num - res_num
    ok_pool = set(numbers([extra_text])) if extra_text else set()
    new_num = sorted(n for n in (res_num - src_num).elements() if n not in ok_pool)
    lost = list((base_cnt - res_cnt).elements()) if base else []
    order_miss = _subseq_missing(*order_rows) if order_rows else []
    bad = [r for r in rows if r[2] in ('없음', '띄어쓰기 다름')]
    return {
        'result': result, 'base': base, 'sources': list(sources), 'rows': rows, 'bad': bad,
        'missing_num': sorted(missing_num.elements()), 'new_num': new_num, 'lost': lost,
        'order_miss': order_miss,
        'ok': not bad and not missing_num and not lost and not new_num and not order_miss,
        'count': Counter(r[2] for r in rows),
    }


def verify_html(v, out, notes=()):
    c = v['count']
    esc = html.escape

    def badge(label, n, bad):
        col = '#C0392B' if (bad and n) else ('#2E7D4F' if not bad else '#7A808A')
        return f'<span class="b" style="background:{col}">{label} {n}</span>'

    head = (badge('일치', c.get('일치', 0), False) + badge('합쳐져 포함', c.get('포함', 0), False) +
            badge('띄어쓰기 다름', c.get('띄어쓰기 다름', 0), True) + badge('없음', c.get('없음', 0), True) +
            badge('순서·칸 어긋남', len(v['order_miss']), True) +
            badge('빠진 숫자', len(v['missing_num']), True) + badge('원본에 없는 숫자', len(v['new_num']), True) +
            (badge('기존 내용 사라짐', len(v['lost']), True) if v['base'] else ''))
    verdict = ('<p class="ok">✔ 원본의 글자·숫자가 결과 파일에 그대로 들어갔습니다.</p>' if v['ok'] else
               '<p class="ng">✘ 확인이 필요한 곳이 있습니다. 아래 빨간 항목을 결과 파일에서 직접 확인하세요.</p>')
    parts = [f'<h1>보고서 제작기 — 원본 대조 결과</h1><div class="m">{datetime.now():%Y-%m-%d %H:%M}'
             f' · 결과: {esc(v["result"])}</div>', head, verdict]
    if notes:
        parts.append('<h2>안내</h2><ul>' + ''.join(f'<li>{esc(n)}</li>' for n in notes) + '</ul>')

    def table(title, items, cols):
        if not items:
            return ''
        trs = ''.join('<tr>' + ''.join(f'<td>{esc(str(x))}</td>' for x in it) + '</tr>' for it in items)
        return (f'<h2>{esc(title)} ({len(items)})</h2><table><tr>' +
                ''.join(f'<th>{esc(h)}</th>' for h in cols) + f'</tr>{trs}</table>')

    parts.append(table('결과에 없는 원고', [r for r in v['rows'] if r[2] == '없음'], ['원본 파일', '원고 내용', '상태']))
    parts.append(table('띄어쓰기가 달라진 원고', [r for r in v['rows'] if r[2] == '띄어쓰기 다름'], ['원본 파일', '원고 내용', '상태']))
    parts.append(table('순서대로·칸 그대로 들어가지 않은 원본 줄', [(' | '.join(r),) for r in v['order_miss']], ['원본 줄(칸은 | 로 구분)']))
    parts.append(table('결과에 모자란 숫자', [(n,) for n in v['missing_num']], ['숫자']))
    parts.append(table('원본에 없는데 결과에 생긴 숫자', [(n,) for n in v['new_num']], ['숫자']))
    if v['base']:
        parts.append(table('기존 보고서에서 사라진 내용', [(x,) for x in v['lost']], ['내용']))
    parts.append(table('원고 전체 대조', v['rows'], ['원본 파일', '원고 내용', '상태']))
    css = ('body{font-family:"맑은 고딕";margin:28px;color:#1E2430;background:#F4F6F9}'
           'h1{font-size:20px}h2{font-size:15px;margin-top:26px}.m{color:#7A808A;font-size:12px;margin-bottom:12px}'
           '.b{display:inline-block;color:#fff;border-radius:12px;padding:3px 12px;margin:0 6px 6px 0;font-size:13px}'
           '.ok{color:#2E7D4F;font-weight:bold;font-size:16px}.ng{color:#C0392B;font-weight:bold;font-size:16px}'
           'li{color:#3B6EA5;margin:3px 0}'
           'table{border-collapse:collapse;background:#fff;width:100%;font-size:12px}'
           'td,th{border:1px solid #DDE2E8;padding:4px 8px;text-align:left;word-break:break-all}th{background:#EAF1F8}')
    with io.open(out, 'w', encoding='utf-8') as f:
        f.write(f'<!doctype html><meta charset="utf-8"><title>원본 대조 결과</title><style>{css}</style>' + ''.join(parts))
    return out


# ───────────────────────── 파일 이름 ─────────────────────────

def next_name(base, added_no=None, ext=None):
    """'…_1-11호' → '…_1-12호', 아니면 '…_추가본'."""
    d, f = os.path.split(base)
    stem, e = os.path.splitext(f)
    e = ext or e
    m = re.search(r'(\d+)\s*[-~]\s*(\d+)호', stem)
    if added_no and m:
        stem = stem[:m.start()] + f'{m.group(1)}-{added_no}호' + stem[m.end():]
    else:
        stem += '_추가본'
    return os.path.join(d, stem + e)


def unique(path):
    stem, e = os.path.splitext(path)
    i, out = 2, path
    while os.path.exists(out):
        out = f'{stem}({i}){e}'
        i += 1
    return out


# ───────────────────────── 작업 ─────────────────────────

def _job_dir():
    d = os.path.join(JOBS, datetime.now().strftime('%Y%m%d_%H%M%S'))
    os.makedirs(d, exist_ok=True)
    return d


def _paste_file(job, text):
    if not text or not text.strip():
        return None
    p = os.path.join(job, '붙여넣은글.txt')
    with io.open(p, 'w', encoding='utf-8') as f:
        f.write(text)
    return p


def _out_rows_after(out, base):
    """결과 줄 목록에서 기존 보고서 몫(앞부분)을 빼고 새로 붙은 쪽만."""
    o = rows_of(out)
    if not base:
        return o
    b = rows_of(base)
    sm = difflib.SequenceMatcher(None, b, o, autojunk=False)
    keep = [False] * len(o)
    for blk in sm.get_matching_blocks():
        for k in range(blk.size):
            keep[blk.b + k] = True
    return [r for r, k in zip(o, keep) if not k]


def _finish(job, sources, out, base, notes, extra_text):
    src_rows = [r for s in sources for r in rows_of(s)]
    v = verify(sources, out, base=base, extra_text=extra_text,
               order_rows=(src_rows, _out_rows_after(out, base)) if out.lower().endswith('.docx') else None)
    rep = verify_html(v, os.path.join(job, '원본대조결과.html'), notes)
    rep2 = os.path.splitext(out)[0] + '_원본대조.html'
    shutil.copyfile(rep, rep2)
    return v, rep2


def _order_sources(sources):
    """원본 여러 개면 호 번호(Vol.N) → 날짜 → 넣은 순서."""
    def key(s):
        try:
            flat = ' '.join(c for r in rows_of(s)[:12] for c in r)
        except Exception:
            return (0, '', 0)
        m = re.search(r'Vol\.?\s*0*(\d+)', flat, re.I)
        dm = DATE.search(flat)
        return (int(m.group(1)) if m else 0, dm.group(0) if dm else '')
    idx = {s: i for i, s in enumerate(sources)}
    if all(key(s) == (0, '') for s in sources):
        return list(sources)
    return sorted(sources, key=lambda s: (key(s), idx[s]))


def _make(sample, sources, out, mode, img_dir='', job=None):
    ext = os.path.splitext(sample)[1].lower()
    if img_dir and not os.path.isdir(img_dir):
        raise ValueError(f'이미지 폴더를 찾을 수 없습니다: {img_dir}')
    if ext == '.docx':
        return docx_add_sections(sample, sources, out, mode, img_dir, job)
    if ext in ('.xlsx', '.xlsm'):
        return xlsx_add_sheets(sample, sources, out, mode, img_dir, job)
    if ext == '.hwp':
        raise ValueError('구형 한글(.hwp) 파일은 프로그램이 고칠 수 없습니다.\n'
                         '한글에서 [파일 → 다른 이름으로 저장]을 눌러 형식을 「워드 문서(*.docx)」로 바꿔 저장한 뒤,\n'
                         '그 .docx 파일을 넣어 주세요.')
    if ext in ('.hwpx', '.pptx', '.pdf'):
        kind = {'.hwpx': '한글(.hwpx)', '.pptx': '파워포인트', '.pdf': 'PDF'}[ext]
        raise ValueError(f'{kind} 샘플은 아직 지원하지 않습니다. 샘플은 워드(.docx)나 엑셀로 넣어 주세요.\n'
                         '(한글은 [다른 이름으로 저장 → 워드 문서(*.docx)]로 바꿀 수 있습니다)')
    raise ValueError('샘플(기존 보고서)은 워드(.docx)나 엑셀(.xlsx) 파일이어야 합니다.')


def add(base, sources, out_dir=None, memo='', paste='', img_dir=''):
    """기존 보고서 + 추가 원고 → 새 파일. 반환 (결과 파일, 대조 리포트, 통과 여부).
    img_dir: 매호 이미지 폴더(…12호.png) — 주면 새 호마다 그 호 그림을 넣는다."""
    job = _job_dir()
    p = _paste_file(job, paste)
    sources = _order_sources(list(sources)) + ([p] if p else [])
    if not sources:
        raise ValueError('추가할 원고를 넣어 주세요.')
    ext = os.path.splitext(base)[1].lower()
    tmp = os.path.join(job, 'result' + ext)
    log('== 기존 보고서의 마지막 묶음 모양을 읽는 중')
    added, heads, notes = _make(base, sources, tmp, 'add', img_dir, job)
    if not added:
        raise ValueError('넣을 수 있는 원고가 없었습니다.\n' + '\n'.join(notes))
    log('== ' + ', '.join(f'{n}호' for _, n in added) + ' 추가 — 원본 값 그대로 복사')
    out = next_name(base, max(n for _, n in added))
    out = unique(os.path.join(out_dir, os.path.basename(out)) if out_dir else out)
    shutil.copyfile(tmp, out)
    log('== 원본과 대조하는 중')
    v, rep = _finish(job, sources, out, base, notes, ' '.join(heads))
    return out, rep, v['ok']


def _heading_nos(heads):
    """묶음 제목들 → 각 묶음의 호 번호. 묶음마다 1씩 늘어나는 숫자 자리(날짜 밖)를 호 번호로 본다."""
    seqs = []
    for h in heads:
        ds = [x.span() for x in DATE.finditer(h)] + [x.span() for x in KDATE.finditer(h)]
        seqs.append([int(x.group()) for x in re.finditer(r'\d+', h) if not any(a <= x.start() < b for a, b in ds)])
    width = max((len(s) for s in seqs), default=0)
    for k in range(width):
        col = [s[k] if len(s) > k else None for s in seqs]
        vals = [v for v in col if v is not None]
        if len(vals) >= 2 and all(b > a for a, b in zip(vals, vals[1:])):
            return col
    return [s[0] if s else None for s in seqs]


def add_images(base, out_dir=None, img_dir=''):
    """원고 없이 이미지만: 기존 보고서의 호마다 그 호 그림(…N호.png)을 넣은 새 파일.
    그림이 없는 호는 그대로 둔다. 글자는 하나도 바꾸지 않는다(대조로 확인)."""
    if not img_dir or not os.path.isdir(img_dir):
        raise ValueError('매호 이미지 폴더를 골라 주세요.')
    job = _job_dir()
    ext = os.path.splitext(base)[1].lower()
    tmp = os.path.join(job, 'result' + ext)
    notes, done = [], []
    log('== 기존 보고서의 호 제목을 읽는 중')
    if ext == '.docx':
        import docx
        d = docx.Document(base)
        items, secs = _sections(d)
        if not secs:
            raise ValueError('보고서에서 반복되는 호 제목(예: □ 1호)을 찾지 못해 이미지를 넣을 자리를 모르겠습니다.')
        heads = [clean(_txt(items[h])) for h, _ in secs]
        for (h, e), no, ht in zip(secs, _heading_nos(heads), heads):
            if no is None:
                notes.append(f'「{ht}」: 호 번호를 몰라 건너뛰었습니다.')
                continue
            sd = source_date([(ht,)])                 # 제목에 날짜가 있으면 그 해 그림을 고른다
            imgs, inotes = images_for(img_dir, no, sd[0] if sd else None, job)
            if not imgs:
                continue
            notes += inotes
            notes += place_images(d, items[h], list(items[h + 1:e]), imgs, f'{no}호')
            done.append(no)
        d.save(tmp)
    elif ext in ('.xlsx', '.xlsm'):
        import openpyxl
        from openpyxl.drawing.image import Image as XLImage
        wb = openpyxl.load_workbook(base)
        for ws in wb:
            k = image_key(ws.title)
            if not k or ws.sheet_state != 'visible':
                continue
            imgs, inotes = images_for(img_dir, k[0], None, job)
            if not imgs:
                continue
            if ws._images:
                notes.append(f'「{ws.title}」 시트에 이미 그림이 있어 건너뛰었습니다.')
                continue
            notes += inotes
            row = ws.max_row + 2
            for path, w, h in imgs:
                im = XLImage(path)
                im.width = min(w, XL_IMG_PX)
                im.height = round(h * im.width / w)
                ws.add_image(im, f'A{row}')
                row += int(im.height / 20) + 2
            done.append(k[0])
        wb.save(tmp)
    else:
        _make(base, [], tmp, 'add')          # 지원 안 하는 형식 안내를 그대로 띄운다
    if not done:
        raise ValueError('이미지 폴더에서 보고서의 호와 맞는 그림을 하나도 찾지 못했습니다.\n'
                         '파일 이름에 호 번호가 있는지 확인하세요 (예: 26_12호.png).')
    log('== ' + ', '.join(f'{n}호' for n in done) + ' 이미지 넣음')
    stem, e = os.path.splitext(os.path.basename(base))
    out = unique(os.path.join(out_dir or os.path.dirname(base), stem + '_이미지' + e))
    shutil.copyfile(tmp, out)
    log('== 기존 글자가 그대로인지 대조하는 중')
    v, rep = _finish(job, [], out, base, notes, '')
    return out, rep, v['ok']


def new(sources, out_dir, title='', template='', memo='', paste='', fmt='.docx', img_dir=''):
    """원본 → 새 보고서. template(샘플)이 있으면 그 모양으로, 없으면 기본 워드 보고서."""
    job = _job_dir()
    p = _paste_file(job, paste)
    sources = _order_sources(list(sources)) + ([p] if p else [])
    if not sources:
        raise ValueError('원본 파일을 넣거나 글을 붙여넣어 주세요.')
    name = re.sub(r'[\\/:*?"<>|]', '_', title.strip()) or '보고서_' + datetime.now().strftime('%Y%m%d')
    if template:
        ext = os.path.splitext(template)[1].lower()
        tmp = os.path.join(job, 'result' + ext)
        log('== 샘플 보고서 모양을 읽는 중')
        added, heads, notes = _make(template, sources, tmp, 'new', img_dir, job)
        out = unique(os.path.join(out_dir, name + ext))
        extra = ' '.join(heads) + ' ' + title + ' ' + ' '.join(units_of(tmp)[:0])
        # 샘플 앞부분(표지 등)에 있던 숫자는 '원본에 없는 숫자'로 치지 않는다
        try:
            import docx as _d
            if ext == '.docx':
                dd = _d.Document(template)
                items, secs = _sections(dd)
                first = secs[0][0] if secs else 0
                extra += ' ' + ' '.join(clean(_txt(x)) for x in items[:first])
        except Exception:
            pass
        base_for_check = None
    else:
        out = unique(os.path.join(out_dir, name + '.docx'))
        tmp = os.path.join(job, 'result.docx')
        log('== 원본의 표·글을 그대로 옮겨 기본 보고서를 만드는 중')
        if img_dir and not os.path.isdir(img_dir):
            raise ValueError(f'이미지 폴더를 찾을 수 없습니다: {img_dir}')
        notes = ['샘플 없이 만든 기본 보고서입니다. 원하는 모양이 있으면 예전 보고서를 샘플로 넣어 주세요.']
        heads = docx_plain(sources, tmp, title, img_dir, job, notes)
        extra = ' '.join(heads)
        base_for_check = None
    shutil.copyfile(tmp, out)
    log('== 원본과 대조하는 중')
    v, rep = _finish(job, sources, out, base_for_check, notes, extra)
    return out, rep, v['ok']


# ───────────────────────── 명령줄 ─────────────────────────

def main():
    import argparse
    ap = argparse.ArgumentParser(description='보고서 제작기')
    sub = ap.add_subparsers(dest='cmd', required=True)
    a = sub.add_parser('add'); a.add_argument('base'); a.add_argument('src', nargs='+'); a.add_argument('--out-dir')
    a.add_argument('--img-dir', default='')
    n = sub.add_parser('new'); n.add_argument('src', nargs='+'); n.add_argument('--out-dir', required=True)
    n.add_argument('--title', default=''); n.add_argument('--template', default='')
    n.add_argument('--img-dir', default='')
    c = sub.add_parser('verify'); c.add_argument('result'); c.add_argument('src', nargs='+'); c.add_argument('--base')
    c.add_argument('--out', default='')
    o = ap.parse_args()
    if o.cmd == 'add':
        print(add(o.base, o.src, o.out_dir, img_dir=o.img_dir))
    elif o.cmd == 'new':
        print(new(o.src, o.out_dir, o.title, o.template, img_dir=o.img_dir))
    else:
        src_rows = [r for s in o.src for r in rows_of(s)]
        v = verify(o.src, o.result, base=o.base, order_rows=(src_rows, _out_rows_after(o.result, o.base)))
        print('OK' if v['ok'] else 'NG', dict(v['count']), 'missing_num', v['missing_num'][:20],
              'new_num', v['new_num'][:20], 'lost', v['lost'][:10], 'order_miss', v['order_miss'][:10])
        if o.out:
            verify_html(v, o.out)


if __name__ == '__main__':
    main()
