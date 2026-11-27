# -*- coding: utf-8 -*-
"""보고서 제작기 시험 (가상 데이터만 사용).   python test_engine.py
실제 보고서로도 돌리려면: python test_engine.py <기존보고서.docx> <추가원고.xlsx> <정답.docx>"""
import datetime
import io
import os
import shutil
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import docx                     # noqa: E402
import openpyxl                 # noqa: E402
import engine                   # noqa: E402
from engine import W, _txt, clean   # noqa: E402

TMP = tempfile.mkdtemp(prefix='rm_test_')
engine.JOBS = os.path.join(TMP, 'jobs')
FAILS = []


def check(name, cond, info=''):
    print(('  ✔ ' if cond else '  ✘ ') + name + ('' if cond else f'   {info}'))
    if not cond:
        FAILS.append(name)


def section_dump(path, start):
    d = docx.Document(path)
    on, out = False, []
    for e in d.element.body:
        if e.tag == W + 'p':
            t = clean(_txt(e))
            if t.startswith(start):
                on = True
            st = e.find(W + 'pPr/' + W + 'pStyle')
            if on and t:
                out.append(('P', t, st.get(W + 'val') if st is not None else ''))
        elif e.tag == W + 'tbl' and on:
            for tr in e.findall(W + 'tr'):
                out.append(('T',) + tuple(clean(_txt(tc)) for tc in tr.findall(W + 'tc')))
    return out


def make_sample(path, months):
    d = docx.Document()
    d.add_heading('2026년 운영 보고서', 0)
    d.add_paragraph('작성: 운영팀')
    for i, (mon, rows, bul) in enumerate(months, 1):
        d.add_heading(f'{i}. 2026년 {mon}월 운영 실적', 1)
        d.add_paragraph('가. 방문자 현황')
        t = d.add_table(rows=1 + len(rows), cols=3)
        t.style = 'Table Grid'
        for k, h in enumerate(['구분', '방문자수', '전월대비']):
            t.cell(0, k).text = h
        for r, row in enumerate(rows, 1):
            for k, v in enumerate(row):
                t.cell(r, k).text = v
        d.add_paragraph('나. 주요 내용')
        for b in bul:
            d.add_paragraph(b, style='List Bullet')
        d.add_paragraph('※ 방문자수는 GA 기준')
    d.save(path)


SAMPLE = os.path.join(TMP, 'sample.docx')
make_sample(SAMPLE, [
    (9, [('홈페이지', '12,345', '5.2%'), ('블로그', '2,345', '-1.0%')], ['9월 이벤트 참여 120명', '신규 콘텐츠 8건 발행']),
    (10, [('홈페이지', '13,001', '5.3%'), ('블로그', '2,400', '2.3%')], ['10월 이벤트 참여 98명', '신규 콘텐츠 6건 발행']),
])


def xlsx_src(path, when, data, bullets, extra_first=True):
    wb = openpyxl.Workbook()
    ws = wb.active
    r = 1
    if extra_first:
        ws.cell(1, 1, '기준일')
        c = ws.cell(1, 2, when)
        c.number_format = 'yyyy-mm-dd'
        r = 3
    ws.cell(r, 1, '가. 방문자 현황')
    r += 1
    for k, h in enumerate(['구분', '방문자수', '전월대비'], 1):
        ws.cell(r, k, h)
    for row in data:
        r += 1
        for k, v in enumerate(row, 1):
            c = ws.cell(r, k, v)
            if k == 2:
                c.number_format = '#,##0'
            if k == 3:
                c.number_format = '0.0%'
    r += 2
    ws.cell(r, 1, '나. 주요 내용')
    for b in bullets:
        r += 1
        ws.cell(r, 1, b)
    wb.save(path)


print('\n[1] 셀 표시 형식 그대로 읽기')
p = os.path.join(TMP, 'fmt.xlsx')
wb = openpyxl.Workbook()
ws = wb.active
vals = [(0.914, '0.0%', '91.4%'), (1234567, '#,##0', '1,234,567'), (3.14159, '0.00', '3.14'), (0.5, '0%', '50%'),
        (datetime.datetime(2026, 11, 23), 'yyyy-mm-dd', '2026-11-23'), (42, 'General', '42'), (2.5, 'General', '2.5'),
        (-0.0792, '0.0%', '-7.9%'), ('텍스트 0.914', '@', '텍스트 0.914')]
for i, (v, f, _) in enumerate(vals, 1):
    c = ws.cell(i, 1, v)
    c.number_format = f
wb.save(p)
got = [r[0] for r in engine.rows_of(p)]
for (v, f, want), g in zip(vals, got):
    check(f'{v!r} [{f}] → {want}', g == want, f'실제 {g}')

print('\n[2] 줄이 늘어난 원고(표 +1줄, 글머리 +1줄, 앞에 기준일) 추가')
s11 = os.path.join(TMP, '11월.xlsx')
xlsx_src(s11, datetime.datetime(2026, 11, 30),
         [('홈페이지', 14210, 0.093), ('블로그', 2210, -0.0792), ('인스타그램', 1520, 0.5)],
         ['11월 이벤트 참여 143명', '신규 콘텐츠 9건 발행', '연말 캠페인 준비 착수'])
out, rep, ok = engine.add(SAMPLE, [s11], out_dir=TMP)
check('대조 통과', ok, rep)
sec = section_dump(out, '3.')
check('제목 3. 2026년 11월', sec[0][:2] == ('P', '3. 2026년 11월 운영 실적'), sec[0])
check('표 4줄(머리+3)', [x for x in sec if x[0] == 'T'] == [
    ('T', '구분', '방문자수', '전월대비'), ('T', '홈페이지', '14,210', '9.3%'),
    ('T', '블로그', '2,210', '-7.9%'), ('T', '인스타그램', '1,520', '50.0%')], sec)
check('새 글머리 줄도 글머리 모양', ('P', '연말 캠페인 준비 착수', 'ListBullet') in sec, sec)
check('고정 문구 ※ 유지', any(x[1] == '※ 방문자수는 GA 기준' for x in sec if x[0] == 'P'), sec)
check('기존 9·10월 그대로', section_dump(out, '1.')[:len(section_dump(SAMPLE, '1.'))] == section_dump(SAMPLE, '1.'))

print('\n[3] 줄이 줄어든 원고(표 1줄, 글머리 1줄)')
s12 = os.path.join(TMP, '12월.xlsx')
xlsx_src(s12, datetime.datetime(2026, 12, 31), [('홈페이지', 9000, -0.1)], ['12월 휴무 안내'])
out, rep, ok = engine.add(SAMPLE, [s12], out_dir=TMP)
check('대조 통과', ok, rep)
sec = section_dump(out, '3.')
check('표 머리+1줄', [x for x in sec if x[0] == 'T'] == [('T', '구분', '방문자수', '전월대비'), ('T', '홈페이지', '9,000', '-10.0%')], sec)
check('옛 값(13,001·2,400·98명) 안 남음', not any(v in str(sec) for v in ('13,001', '2,400', '98명', '6건')), sec)
check('제목 12월', sec[0][1] == '3. 2026년 12월 운영 실적', sec[0])

print('\n[4] 원고 여러 개를 한꺼번에(순서 섞어서) 새 보고서로')
out, rep, ok = engine.new([s12, s11], TMP, '하반기', SAMPLE)
check('대조 통과', ok, rep)
d = [clean(p.text) for p in docx.Document(out).paragraphs if clean(p.text)]
heads = [t for t in d if t[:2] in ('1.', '2.', '3.')]
check('묶음 2개, 날짜순, 번호 1·2', heads == ['1. 2026년 11월 운영 실적', '2. 2026년 12월 운영 실적'], heads)
check('표지 유지', d[0] == '2026년 운영 보고서')

print('\n[5] 변조 잡아내기')
out, rep, ok = engine.add(SAMPLE, [s11], out_dir=TMP)
dd = docx.Document(out)
dd.tables[-1].rows[1].cells[1].paragraphs[0].runs[0].text = '14,201'     # 14,210 → 14,201
bad = os.path.join(TMP, 'bad.docx')
dd.save(bad)
v = engine.verify([s11], bad, base=SAMPLE, extra_text='3. 2026년 11월 운영 실적')
check('숫자 한 자리 바뀐 것 불합격', not v['ok'] and '14,210' in v['missing_num'] and '14,201' in v['new_num'], v['missing_num'])
dd = docx.Document(out)
dd.paragraphs[1].runs[0].text = '작성: 기획팀'
dd.save(bad)
v = engine.verify([s11], bad, base=SAMPLE, extra_text='3. 2026년 11월 운영 실적')
check('기존 내용 바뀐 것 불합격', not v['ok'] and v['lost'], v['lost'])

print('\n[6] 워드 원고·붙여넣은 글')
wsrc = os.path.join(TMP, '원고.docx')
w = docx.Document()
w.add_paragraph('기준일 2026년 11월 30일')
w.add_paragraph('가. 방문자 현황')
t = w.add_table(rows=3, cols=3)
for r, row in enumerate([('구분', '방문자수', '전월대비'), ('홈페이지', '14,210', '9.3%'), ('블로그', '2,210', '-7.9%')]):
    for k, v in enumerate(row):
        t.cell(r, k).text = v
w.add_paragraph('나. 주요 내용')
w.add_paragraph('11월 이벤트 참여 143명')
w.save(wsrc)
out, rep, ok = engine.add(SAMPLE, [wsrc], out_dir=TMP)
check('워드 원고 대조 통과', ok, rep)
check('제목 11월', section_dump(out, '3.')[0][1] == '3. 2026년 11월 운영 실적')
out, rep, ok = engine.add(SAMPLE, [], out_dir=TMP, paste='2026.11.30 기준\n가. 방문자 현황\n11월 이벤트 참여 143명\n')
check('붙여넣은 글 대조 통과', ok, rep)

print('\n[7] 샘플 없이 기본 보고서')
out, rep, ok = engine.new([s11, wsrc], TMP, '기본')
check('대조 통과', ok, rep)

print('\n[8] 지원 안 하는 샘플은 알아듣게 거절')
try:
    engine.add(os.path.join(TMP, 'x.hwp'), [s11], out_dir=TMP)
    check('hwp 거절', False)
except ValueError as e:
    check('hwp 거절 + docx 로 저장 안내', 'docx' in str(e), str(e))

print('\n[9] 매호 이미지')
import hashlib                  # noqa: E402
from PIL import Image           # noqa: E402
from docx.shared import Cm      # noqa: E402

check('이름 26_12호.png → 12호·26년·순서0', engine.image_key('26_12호.png') == (12, 26, 0), engine.image_key('26_12호.png'))
check('이름 26_1호_2.jpg → 1호·순서2', engine.image_key('26_1호_2.jpg') == (1, 26, 2))
check('이름 12호.png → 연도 없음', engine.image_key('12호.png') == (12, None, 0))
check('호 없는 이름은 무시', engine.image_key('cover.png') is None)


def img(path, w, h, color, noise=False):
    im = Image.new('RGB', (w, h), color)
    if noise:
        im = Image.frombytes('RGB', (w, h), os.urandom(w * h * 3))
    im.save(path)


def pics_of(path, start):
    """start 로 시작하는 묶음 제목부터 다음 같은 모양 제목 전까지의 그림 [(가로 EMU, 그림 해시)]."""
    d = docx.Document(path)
    on, out = False, []
    shape = None
    for e in d.element.body:
        t = clean(_txt(e)) if e.tag == W + 'p' else ''
        if t.startswith(start):
            on, shape = True, engine._shape(t)
            continue
        if on and t and engine._shape(t) == shape:
            break
        if on:
            for dr in e.iter(W + 'drawing'):
                ext = next(dr.iter(engine.WP + 'extent'))
                rid = next(dr.iter(engine.A + 'blip')).get(engine.RNS + 'embed')
                out.append((int(ext.get('cx')), hashlib.md5(d.part.related_parts[rid].blob).hexdigest()))
    return out


IMG = os.path.join(TMP, '이미지')
os.makedirs(IMG)
img(os.path.join(IMG, '26_3호.png'), 3000, 1500, (200, 60, 60))       # 큰 그림 → 2000px 로
img(os.path.join(IMG, '26_3호_2.png'), 800, 800, (60, 120, 200))
img(os.path.join(IMG, '25_3호.png'), 800, 400, (0, 0, 0))             # 작년 3호 — 넣으면 안 됨
img(os.path.join(IMG, '26_2호.png'), 800, 400, (0, 200, 0))

out, rep, ok = engine.add(SAMPLE, [s11], out_dir=TMP, img_dir=IMG)
check('그림 자리 없는 샘플: 대조 통과', ok, rep)
ps = pics_of(out, '3.')
check('3호에 그림 2장(26년 것만, 순서대로)', len(ps) == 2, ps)
d = docx.Document(out)
cw = d.sections[-1].page_width - d.sections[-1].left_margin - d.sections[-1].right_margin
check('본문 폭으로', ps and ps[0][0] == cw, (ps, cw))
body = [e for e in d.element.body]
hi = next(i for i, e in enumerate(body) if clean(_txt(e)).startswith('3.'))
check('제목 바로 아래', next(body[hi + 1].iter(W + 'drawing'), None) is not None)
rid = next(body[hi + 1].iter(engine.A + 'blip')).get(engine.RNS + 'embed')
check('큰 그림은 2000px 로 줄여 넣음', Image.open(io.BytesIO(d.part.related_parts[rid].blob)).size == (2000, 1000))
check('기존 1·2호엔 그림 없음', not pics_of(out, '1.') and not pics_of(out, '2.'))
check('새 3호 글자는 그대로', [x for x in section_dump(out, '3.') if x[0] == 'T'][1] == ('T', '홈페이지', '14,210', '9.3%'))

# 묶음마다 그림이 있는 샘플 → 그 자리·폭으로 바꿔 넣기
SAMP2 = os.path.join(TMP, 'sample_img.docx')
make_sample(SAMP2, [
    (9, [('홈페이지', '12,345', '5.2%')], ['9월 이벤트']),
    (10, [('홈페이지', '13,001', '5.3%')], ['10월 이벤트']),
])
old = os.path.join(TMP, 'old.png')
img(old, 600, 300, (120, 120, 120))
dd = docx.Document(SAMP2)
heads_ = [p for p in dd.paragraphs if clean(p.text)[:2] in ('1.', '2.')]
for h_ in heads_:
    np_ = dd.add_paragraph()
    np_.add_run().add_picture(old, width=Cm(10))
    h_._p.addnext(np_._p)
dd.save(SAMP2)
out, rep, ok = engine.add(SAMP2, [s11], out_dir=TMP, img_dir=IMG)
check('그림 있는 샘플: 대조 통과', ok, rep)
ps = pics_of(out, '3.')
oldh = pics_of(SAMP2, '2.')[0][1]
check('3호 그림 2장, 옛 그림 아님', len(ps) == 2 and all(h != oldh for _, h in ps), ps)
check('두 장 모두 샘플 그림 폭(10cm) 그대로', len(ps) == 2 and all(cx == Cm(10) for cx, _ in ps), ps)
check('기존 2호 그림은 그대로', pics_of(out, '2.') == pics_of(SAMP2, '2.'))
ids = [x.get('id') for x in docx.Document(out).element.body.iter(engine.WP + 'docPr')]
check('그림 번호(docPr id) 겹치지 않음', len(ids) == len(set(ids)), ids)

# 엑셀 샘플
XS = os.path.join(TMP, 'xs.xlsx')
wb = openpyxl.Workbook()
wb.active.title = '1호'
wb.active['A1'] = '1호'
wb.create_sheet('2호')['A1'] = '2호'
wb.save(XS)
out, rep, ok = engine.add(XS, [s11], out_dir=TMP, img_dir=IMG)
wo = openpyxl.load_workbook(out)
check('엑셀: 3호 시트에 그림 2장', '3호' in wo.sheetnames and len(wo['3호']._images) == 2,
      (wo.sheetnames, [len(ws._images) for ws in wo]))

shutil.move(os.path.join(IMG, '26_3호.png'), os.path.join(TMP, 'x.png'))
shutil.move(os.path.join(IMG, '26_3호_2.png'), os.path.join(TMP, 'y.png'))
shutil.move(os.path.join(IMG, '25_3호.png'), os.path.join(TMP, 'z.png'))
out, rep, ok = engine.add(SAMP2, [s11], out_dir=TMP, img_dir=IMG)
check('그 호 그림이 없으면 샘플 그림을 빼 둠', pics_of(out, '3.') == [], pics_of(out, '3.'))
out, rep, ok = engine.add(SAMP2, [s11], out_dir=TMP)
check('이미지 폴더 안 쓰면 예전처럼(샘플 그림 복제)', len(pics_of(out, '3.')) == 1)

q, w_, h_, n_ = engine.prep_image(os.path.join(TMP, 'x.png'), os.path.join(TMP, 'prep'))
check('3000px → 2000px', (w_, h_) == (2000, 1000) and Image.open(q).size == (2000, 1000), (w_, h_, n_))
img(os.path.join(TMP, 'photo.png'), 1800, 1200, None, noise=True)
q, w_, h_, n_ = engine.prep_image(os.path.join(TMP, 'photo.png'), os.path.join(TMP, 'prep'))
check('용량 큰 사진 PNG → JPG', q.endswith('.jpg') and 'JPG' in n_, (q, n_))
q, w_, h_, n_ = engine.prep_image(os.path.join(TMP, 'y.png'), os.path.join(TMP, 'prep'))
check('작은 PNG 는 그대로', q == os.path.join(TMP, 'y.png') and not n_, (q, n_))

print('\n[10] 원고 없이 이미지만 넣기')
check('제목 → 호 번호', engine._heading_nos(['□ 10호 (2026.10.09)', '□ 11호 (2026.11.09)']) == [10, 11])
check('제목 → 호 번호(번호. 연월)', engine._heading_nos(['1. 2026년 9월 실적', '2. 2026년 10월 실적']) == [1, 2])
IMG2 = os.path.join(TMP, '이미지2')
os.makedirs(IMG2)
img(os.path.join(IMG2, '26_1호.png'), 800, 400, (10, 10, 10))
img(os.path.join(IMG2, '26_2호_1.png'), 800, 400, (20, 20, 20))
img(os.path.join(IMG2, '26_2호_2.png'), 400, 800, (30, 30, 30))
out, rep, ok = engine.add_images(SAMPLE, TMP, IMG2)
check('글자 그대로(대조 통과)', ok, rep)
check('글자 줄 전부 같음', engine.rows_of(out) == engine.rows_of(SAMPLE))
check('1호 1장 · 2호 2장', len(pics_of(out, '1.')) == 1 and len(pics_of(out, '2.')) == 2, (pics_of(out, '1.'), pics_of(out, '2.')))
check('파일 이름 …_이미지.docx', out.endswith('_이미지.docx'), out)
out, rep, ok = engine.add_images(SAMP2, TMP, IMG2)
ps1 = pics_of(out, '1.')
check('그림 있던 호는 그 자리·폭으로 교체', len(ps1) == 1 and ps1[0][1] != pics_of(SAMP2, '1.')[0][1] and ps1[0][0] == Cm(10), ps1)
try:
    engine.add_images(SAMPLE, TMP, os.path.join(TMP, 'prep'))
    check('맞는 그림이 없으면 알려 줌', False)
except ValueError as e:
    check('맞는 그림이 없으면 알려 줌', '찾지 못했' in str(e), str(e))

if len(sys.argv) == 4:
    print('\n[실제] 기존 보고서 + 추가 원고 → 정답과 비교')
    out, rep, ok = engine.add(sys.argv[1], [sys.argv[2]], out_dir=TMP)
    check('대조 통과', ok, rep)
    check('정답과 글자 동일', engine.rows_of(out) == engine.rows_of(sys.argv[3]))

print('\n' + ('모두 통과' if not FAILS else f'실패 {len(FAILS)}건: {FAILS}'))
shutil.rmtree(TMP, ignore_errors=True)
sys.exit(1 if FAILS else 0)
