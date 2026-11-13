# -*- coding: utf-8 -*-
"""상담원장 정리 스크립트

따르는 결정 기록: my-wiki/decisions/d-005-상담원장-정리-기준.md
  d-005-1  처리상태 «종결»은 «완료»로, «대기»는 «보류»로 합친다 (CS기획팀 답)
  d-005-2  처리시간 음수는 지우지 않는다. 값은 그대로 두고 맨 끝 _표시 칸에 이유를 적는다
  d-005-3  만족도 6~10점은 지우지 않는다. 값은 그대로 두고 _표시 칸에 이유를 적는다
           (평균 만족도를 낼 때는 _표시한 줄을 뺀다 — 이 스크립트는 값을 내지 않는다)
  d-005-4  칸 이름은 원본 그대로 둔다. 새 칸은 맨 끝 _표시 하나만 더한다
  d-005-5  고객번호 앞뒤 공백을 없앤다
  d-007    처리상태 «처리완료»는 «완료»로, «처리중»은 «진행중»으로 합친다

그 밖의 처리는 원본 «작성기준» 시트를 따르거나, 2026-11-12 정리 때 묻지 않고 정한 것이다.
아직 결정 기록에 없는 것은 my-wiki/log/2026-11-12.md «묻지 않고 정했던 것» 참고.
  - 작성기준: 테스트 접수(AG999) 제거 · 같은 접수번호는 최종수정일이 늦은 줄만 ·
              7/1 이후 전화 채널 상담시간 초→분 · 4월까지 만족도 0/-/미응답 → 빈칸
  - 짐작(미결정): 전화 채널 표기 묶기 · 7/1 당일 포함 ·
              모든 칸 앞뒤 공백 제거 · 접수일 형식 · 지역/팀/월정액/재문의 표기 · 만족도 «N점»

원본은 읽기만 한다.
사용법: python 정리/정리_상담원장.py [원본.xlsx] [출력 폴더]
  기본값: ../messy/상담원장_2025.xlsx → 정리본/
"""
import csv
import collections
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

sys.stdout.reconfigure(encoding='utf-8')   # 화면 출력 한글 깨짐 방지 (결과 파일과 무관)

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
SRC = sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, '..', 'messy', '상담원장_2025.xlsx')
OUT = sys.argv[2] if len(sys.argv) > 2 else os.path.join(ROOT, '정리본')

# ---------- 엑셀 읽기 (openpyxl 없이) ----------
NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
RNS = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'


def col_index(ref):
    n = 0
    for ch in re.match(r'[A-Z]+', ref).group():
        n = n * 26 + ord(ch) - 64
    return n - 1


def read_sheet(path, name):
    with zipfile.ZipFile(path) as z:
        shared = []
        if 'xl/sharedStrings.xml' in z.namelist():
            for si in ET.fromstring(z.read('xl/sharedStrings.xml')).iter(NS + 'si'):
                shared.append(''.join(t.text or '' for t in si.iter(NS + 't')))
        wb = ET.fromstring(z.read('xl/workbook.xml'))
        rels = {r.get('Id'): r.get('Target') for r in ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
        target = None
        for s in wb.iter(NS + 'sheet'):
            if s.get('name') == name:
                target = rels[s.get(RNS)].lstrip('/')
                target = target if target.startswith('xl/') else 'xl/' + target
        rows = []
        for r in ET.fromstring(z.read(target)).iter(NS + 'row'):
            cells = {}
            for c in r.iter(NS + 'c'):
                t, v = c.get('t'), c.find(NS + 'v')
                if t == 's':
                    val = shared[int(v.text)]
                elif t == 'inlineStr':
                    val = ''.join(x.text or '' for x in c.iter(NS + 't'))
                elif t == 'b':
                    val = 'TRUE' if v is not None and v.text == '1' else 'FALSE'
                else:
                    val = v.text if v is not None else ''
                cells[col_index(c.get('r'))] = val
            if cells and any(x != '' for x in cells.values()):
                rows.append((int(r.get('r')), [cells.get(i, '') for i in range(max(cells) + 1)]))
        return rows


# ---------- 표기 묶음 ----------
PHONE = {'전화', '유선', '전화상담', 'TEL'}
CHANNEL = {}
for k in PHONE:
    CHANNEL[k] = '전화'
for k in ['챗봇', 'chatbot', '챗 봇', 'Chatbot']:
    CHANNEL[k] = '챗봇'
for k in ['앱', '모바일앱', 'app', 'APP']:
    CHANNEL[k] = '앱'
for k in ['이메일', 'e-mail', '메일', 'E-mail']:
    CHANNEL[k] = '이메일'
REGION = {'경기도': '경기', '서울시': '서울', '서울특별시': '서울'}
STATUS = {'종결': '완료', '대기': '보류',                       # d-005-1
          '처리완료': '완료', '처리중': '진행중'}               # d-007
YES, NO = {'Y', 'y', '예', 'TRUE'}, {'N', 'n', '아니오', 'FALSE'}
SAT_NO_ANSWER = {'0', '-', '미응답'}

FLAG_NEG_TIME = '처리시간 음수: 틀린 값이지만 진짜 값을 몰라 그대로 둠 (d-005-2)'
FLAG_SAT_RANGE = '만족도 1~5 밖: NPS를 잘못 넣은 것으로 보이나 확인 못함, 평균에서 뺌 (d-005-3)'


def clean_row(raw, H):
    """한 줄을 정리한다. (정리된 줄, 바꾼 것 목록, _표시 문구) 를 돌려준다."""
    o = list(raw)
    changes = []

    def put(col, value, why):
        i = H[col]
        if o[i] != value:
            o[i] = value
            changes.append((col, why))

    for col, i in H.items():                                   # d-005-5 고객번호 포함
        if o[i] != o[i].strip():
            put(col, o[i].strip(), '앞뒤 공백 제거')
    g = lambda c: o[H[c]]

    put('접수일', re.sub(r'\D', '-', g('접수일')), '접수일 형식 YYYY-MM-DD로')
    put('지역', REGION.get(g('지역'), g('지역')), '지역 표기 묶음')
    if g('월정액'):
        put('월정액', re.sub(r'\D', '', g('월정액')), '월정액 숫자만')
    put('접수채널', CHANNEL[g('접수채널')], '채널 표기 묶음')
    put('상담원팀', g('상담원팀').replace(' ', ''), '팀 이름 띄어쓰기 제거')
    if g('접수채널') == '전화' and g('접수일') >= '2025-07-01':
        minutes = float(g('상담시간')) / 60
        put('상담시간', ('%.2f' % minutes).rstrip('0').rstrip('.'), '7/1 이후 전화: 초 → 분')
    q = g('재문의')
    if q in YES:
        put('재문의', 'Y', '재문의 표기 Y/N으로')
    elif q in NO:
        put('재문의', 'N', '재문의 표기 Y/N으로')
    elif q != '':
        raise ValueError('모르는 재문의 값: %r' % q)
    s = g('만족도')
    if s in SAT_NO_ANSWER:
        put('만족도', '', '만족도 미응답 표기(0 · - · 미응답) → 빈칸')
    elif s.endswith('점'):
        put('만족도', s[:-1], '만족도 «N점» → N')
    put('처리상태', STATUS.get(g('처리상태'), g('처리상태')),
        '처리상태 %s → %s' % (g('처리상태'), STATUS.get(g('처리상태'), g('처리상태'))))

    flags = []
    if g('처리시간') and float(g('처리시간')) < 0:
        flags.append(FLAG_NEG_TIME)
    if g('만족도') and not 1 <= int(g('만족도')) <= 5:
        flags.append(FLAG_SAT_RANGE)
    return o, changes, ' / '.join(flags)


def normalize_for_compare(r, H):
    return tuple(re.sub(r'\D', '-', x.strip()) if i == H['접수일'] else x.strip() for i, x in enumerate(r))


def main():
    rows = read_sheet(SRC, '원장')
    header = rows[0][1]
    H = {h: i for i, h in enumerate(header)}
    raw = [(n, r + [''] * (len(header) - len(r))) for n, r in rows[1:]]

    deleted = []          # (원본행, 원본 값, 이유)
    cleaned = []          # (원본행, 원본 값, 정리된 값, 바꾼 것, 표시)
    for n, r in raw:
        if r[H['상담원ID']] == 'AG999':
            deleted.append((n, r, '테스트 접수 (상담원ID AG999)'))
            continue
        o, ch, flag = clean_row(r, H)
        cleaned.append((n, r, o, ch, flag))

    # 같은 접수번호: 최종수정일이 가장 늦은 줄 하나만 남긴다
    groups = collections.OrderedDict()
    for item in cleaned:
        groups.setdefault(item[1][H['접수번호']], []).append(item)
    keep = set()
    for key, items in groups.items():
        latest = max(it[1][H['최종수정일']] for it in items)
        top = [it for it in items if it[1][H['최종수정일']] == latest]
        if len({normalize_for_compare(it[1], H) for it in top}) != 1:
            raise SystemExit('최종수정일이 같은데 내용이 다른 줄: ' + key)
        kn, kr = top[0][0], top[0][1]
        keep.add(kn)
        for n, r, *_ in items:
            if n == kn:
                continue
            if r == kr:
                why = '완전 중복 (모든 칸이 같음)'
            elif normalize_for_compare(r, H) == normalize_for_compare(kr, H):
                why = '중복 (공백·날짜 표기만 다름)'
            else:
                why = '옛 상태 (같은 접수번호, 최종수정일이 더 이름)'
            deleted.append((n, r, why))

    final = [it for it in cleaned if it[0] in keep]
    deleted.sort(key=lambda d: d[0])
    assert len(final) + len(deleted) == len(raw)

    os.makedirs(OUT, exist_ok=True)
    with open(os.path.join(OUT, '상담원장_정리.csv'), 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.writer(f, lineterminator='\r\n')
        w.writerow(header + ['_표시'])                         # d-005-4
        for n, r, o, ch, flag in final:
            w.writerow(o + [flag])
    with open(os.path.join(OUT, '지운줄.csv'), 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.writer(f, lineterminator='\r\n')
        w.writerow(['원본행'] + header + ['지운 이유'])
        for n, r, why in deleted:
            w.writerow([n] + r + [why])

    # ---------- 변경요약 ----------
    del_count = collections.Counter(why for _, _, why in deleted)
    cell = collections.Counter()
    rows_by_why = collections.defaultdict(set)
    for n, r, o, ch, flag in final:
        for col, why in ch:
            cell[(col, why)] += 1
            rows_by_why[why].add(n)
    flag_count = collections.Counter()
    for n, r, o, ch, flag in final:
        for part in flag.split(' / ') if flag else []:
            flag_count[part] += 1
    phone_after = sum(1 for _, _, o, _, _ in final if o[H['접수채널']] == '전화' and o[H['접수일']] >= '2025-07-01')

    L = []
    L.append('# 상담원장 변경요약')
    L.append('')
    L.append('- 원본: `%s` «원장» 시트 (읽기만 함)' % os.path.basename(SRC))
    L.append('- 스크립트: `정리/정리_상담원장.py`')
    L.append('- 따르는 결정 기록: d-005 · d-007 (my-wiki/decisions/)')
    L.append('')
    L.append('## 줄 수')
    L.append('')
    L.append('| 단계 | 줄 수 |')
    L.append('|---|---|')
    L.append('| 원본 | {:,} |'.format(len(raw)))
    for why in ['테스트 접수 (상담원ID AG999)', '옛 상태 (같은 접수번호, 최종수정일이 더 이름)',
                '완전 중복 (모든 칸이 같음)', '중복 (공백·날짜 표기만 다름)']:
        L.append('| − {} | −{:,} |'.format(why, del_count[why]))
    L.append('| = 정리본 | {:,} |'.format(len(final)))
    L.append('')
    L.append('지운 줄 {:,}줄은 `지운줄.csv`에 원본 값 그대로, 맨 끝 칸에 이유를 붙여 두었다.'.format(len(deleted)))
    L.append('')
    L.append('## 상담시간 단위')
    L.append('')
    L.append('**상담시간 칸의 2025-07-01 이후 전화 줄은 초 단위였던 값을 60으로 나눠 분으로 바꿨다.** '
             '정리본에서 {:,}줄이다. 원본 «작성기준» 시트: «2025-07-01 전화 시스템 교체 뒤로 전화 채널은 초 단위로 기록됨». '
             '전화 채널에는 전화 · 유선 · 전화상담 · TEL(앞뒤 공백 제거 후)을 넣었고, 7/1 당일도 포함했다(둘 다 짐작, 미결정).'
             .format(phone_after))
    L.append('')
    L.append('## 바꾼 칸 (정리본에 남은 줄 기준)')
    L.append('')
    L.append('| 칸 | 무엇을 | 바꾼 칸 수 | 근거 |')
    L.append('|---|---|---|---|')
    basis = {
        '처리상태 종결 → 완료': 'd-005-1', '처리상태 대기 → 보류': 'd-005-1',
        '처리상태 처리완료 → 완료': 'd-007', '처리상태 처리중 → 진행중': 'd-007',
        '7/1 이후 전화: 초 → 분': '작성기준', '만족도 미응답 표기(0 · - · 미응답) → 빈칸': '작성기준',
    }
    for (col, why), c in sorted(cell.items(), key=lambda x: (header.index(x[0][0]), x[0][1])):
        b = basis.get(why, 'd-005-5' if (col == '고객번호' and why == '앞뒤 공백 제거') else '짐작 (미결정)')
        L.append('| {} | {} | {:,} | {} |'.format(col, why, c, b))
    L.append('')
    L.append('## _표시 (값은 그대로 두고 이유만 적음)')
    L.append('')
    L.append('| 이유 | 줄 수 |')
    L.append('|---|---|')
    for part in [FLAG_NEG_TIME, FLAG_SAT_RANGE]:
        L.append('| {} | {:,} |'.format(part, flag_count[part]))
    L.append('')
    L.append('## 바꾸지 않은 것')
    L.append('')
    L.append('- 칸 이름 24개는 원본 그대로다 (d-005-4). 맨 끝에 `_표시` 칸 하나만 더했다.')
    L.append('- 처리시간 음수와 만족도 6~10점의 값은 그대로 두었다 (위 _표시).')
    with open(os.path.join(OUT, '변경요약.md'), 'w', encoding='utf-8', newline='\n') as f:
        f.write('\n'.join(L) + '\n')

    print('원본 {:,} → 정리본 {:,} (지운 줄 {:,})'.format(len(raw), len(final), len(deleted)))


if __name__ == '__main__':
    main()
