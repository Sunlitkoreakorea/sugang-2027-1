# 사용: python -I build.py raw/2027-1_1차_수정.xlsx "1차 수정"
#   두 번째 인자는 제목 옆에 붙는 목록 이름(없으면 "1차 목록").
# 엑셀(년도/학기/학년/블록/학점/과목명/대표교사명/공동교사명)을 읽어
# 수정본처럼 머리 위에 '변경 있음'/'신규 추가' 색 범례가 있고 칸에 "D1 → D2"처럼 적혀 있으면
# 바뀐 뒤 값으로 넣고, 타일에 "D1에서 옮김", "3→4학점", "새 과목" 태그를 단다.
# template.html 의 /*DATA*/…/*END*/ 자리에 넣고
#   artifact.html (아티팩트용 조각)  /  index.html (그냥 열어보는 완전한 문서) 를 만든다.
import json, re, sys
from pathlib import Path
import openpyxl

HERE = Path(__file__).resolve().parent
src = Path(sys.argv[1])
label = sys.argv[2] if len(sys.argv) > 2 else '1차 목록'

EXPECT = ('년도', '학기', '학년', '블록', '학점', '과목명', '대표교사명', '공동교사명')

# 괄호 속 메모 → 짧은 태그
NOTE_RULES = [
    (r'수요 많은 (\d)개 블록 개설', lambda m: f'{m.group(1)}개 블록만 개설'),
    (r'수요조사 후 ([\d,]+)/([A-Z]\d?)\S* 변경될 수 있음', lambda m: f'{m.group(1)}학년 {m.group(2)} 블록으로 바뀔 수 있음'),
    (r'([A-Z]\d?), ([A-Z]\d?), ([A-Z]\d?) 블록 중 수강자 수', lambda m: f'{m.group(1)}, {m.group(2)}, {m.group(3)} 중 1개 블록만 개설'),
    (r'^2022$', lambda m: '2022 교육과정'),
    (r'^가칭$', lambda m: '가칭'),
    (r'기초 수강생만', lambda m: 'R&E 기초 수강생만'),
]

# 과목 분야 (위에서부터 먼저 맞는 것)
OVERRIDE = {'사회과학 방법론': 'social'}
AREA_RULES = [
    ('foreign',  ['한문', '스페인어', '중국어', '프랑스어']),
    ('info',     ['컴퓨터', '프로그래밍', '인공지능']),
    ('art',      ['체육', '미술', '음악', '스포츠']),
    ('research', ['R&E', '과제연구']),
    ('english',  ['영어', '영문학', '영미', '영국 문학', '미국 문학']),
    ('math',     ['수학', '미적분', '대수', '통계', '기하', '방정식']),
    ('science',  ['과학', '생물', 'Biology', 'Chemistry', 'Physics', '물리', '생명', '화학', '세포',
                  '천문', '전자기', '양자', '분자', '탄소', '은하', '역학']),
    ('social',   ['사회', '한국사', '경제', '역사', '윤리', 'History', '세계사', '미국사', '심리',
                  '철학', '지리', '정치', '유럽사', '동아시아']),
    ('korean',   ['국어', '문학', '희곡', '독서', '언어', '시 이해', '소설', '매체']),
]
DEPTS = {'과학과', '사회과', '수학과', '예술체육과', '국어과', '외국어과', '지도교사'}


def clean(s):
    return re.sub(r'\s+', ' ', str(s or '')).strip()


def split_name(raw):
    name, notes = clean(raw), []
    while True:
        m = re.search(r'\s*\(([^()]*)\)\s*$', name)
        if not m:
            break
        notes.insert(0, m.group(1).strip())
        name = name[:m.start()].strip()
    name = re.sub(r'([가-힣])(I{1,3})$', r'\1 \2', name)  # 미적분I → 미적분 I
    tags = []
    for note in notes:
        for pat, fn in NOTE_RULES:
            m = re.search(pat, note)
            if m:
                tags.append([fn(m), note])
                break
        else:
            print('  ! 모르는 메모, 원문 그대로 둠:', note)
            tags.append([note, note])
    return name, tags


def area_of(name):
    if name in OVERRIDE:
        return OVERRIDE[name]
    for key, words in AREA_RULES:
        if any(w in name for w in words):
            return key
    return None


def people(main, co):
    out = []
    for part in [clean(main)] + [clean(x) for x in clean(co).split(',')]:
        if part:
            out.append(part)
    return out


wb = openpyxl.load_workbook(src, data_only=True)
ws = wb.worksheets[0]
cells = [row[:8] for row in ws.iter_rows()]


def fill_key(c):
    f = c.fill
    if f is None or f.fill_type is None:
        return None
    col = f.fgColor
    return (col.type, col.rgb if col.type == 'rgb' else (col.theme if col.type == 'theme' else col.indexed), round(col.tint or 0, 2))


# 머리 행 찾기 + 그 위의 색 범례
hi = next(i for i, row in enumerate(cells[:10]) if tuple(clean(c.value) for c in row) == EXPECT)
legend = {}
for row in cells[:hi]:
    for c in row:
        k, v = fill_key(c), clean(c.value)
        if k and v:
            legend[k] = 'new' if '신규' in v else 'chg' if ('변경' in v or '수정' in v) else None


def split_arrow(v):
    v = clean(v)
    if '→' in v:
        a, b = [x.strip() for x in v.split('→', 1)]
        return b, a
    return v, None


data, missing, changes = [], [], []
year = term = None
for row in cells[hi + 1:]:
    if not any(c.value is not None for c in row):
        continue
    vals, olds = [], []
    for c in row:
        v, o = split_arrow(c.value) if isinstance(c.value, str) else (c.value, None)
        vals.append(v); olds.append(o)
    y, t, g, b, c, n, t1, t2 = vals
    year, term = year or y, term or t
    name, tags = split_name(n)
    area = area_of(name)
    if not area:
        missing.append(name)
    status = next((legend.get(fill_key(x)) for x in row if legend.get(fill_key(x))), None)
    chg = []
    if olds[3]:
        chg.append(f'{olds[3]}에서 옮김')
    if olds[4]:
        chg.append(f'{olds[4]}→{c}학점')
    for i, nm in ((0, '년도'), (1, '학기'), (2, '학년'), (5, '과목명'), (6, '교사'), (7, '교사')):
        if olds[i]:
            chg.append(f'{nm} 바뀜')
    if status == 'new':
        chg = ['새 과목']
    elif status == 'chg' and not chg:
        chg = ['바뀜']
    if chg:
        changes.append(f'{g}/{clean(b)} {name}: {", ".join(chg)}')
    grades = [int(x) for x in re.findall(r'\d', str(g))]
    data.append({
        'g': grades,
        'b': clean(b),
        'c': int(c),
        'n': name,
        't': people(t1, t2),
        'x': tags,
        'a': area or 'research',
        **({'d': chg} if chg else {}),
    })

for line in changes:
    print('  바뀐 과목:', line)

if missing:
    print('  ! 분야를 못 정한 과목:', sorted(set(missing)))

payload = {'year': year, 'term': term, 'label': label, 'rows': data}
js = json.dumps(payload, ensure_ascii=False, separators=(',', ':'))

tpl = (HERE / 'template.html').read_text(encoding='utf-8')
page, n = re.subn(r'/\*DATA\*/.*?/\*END\*/', lambda m: '/*DATA*/' + js + '/*END*/', tpl, flags=re.S)
assert n == 1, 'template.html 에 /*DATA*/…/*END*/ 자리가 없음'

(HERE / 'artifact.html').write_text(page, encoding='utf-8')
head = ('<!doctype html>\n<html lang="ko">\n<meta charset="utf-8">\n'
        '<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">\n')
(HERE / 'index.html').write_text(head + page + '\n</html>\n', encoding='utf-8')

by_grade = {g: sum(1 for d in data if g in d['g']) for g in (1, 2, 3)}
print(f'{len(data)}행  학년별 {by_grade}  분야 {sorted(set(d["a"] for d in data))}')
