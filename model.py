import json, io
import numpy as np
import openpyxl
book = None

def emit(x):
    return json.dumps(x, ensure_ascii=False, allow_nan=False)

def find_sheet(key):
    """시트 이름이 정확히 같으면 그것을, 아니면 key를 포함하는 첫 시트를 반환."""
    if key in book:
        return key
    for s in book:
        if key in s:
            return s
    raise ValueError('필요한 시트를 찾지 못했습니다: ' + key + ' / 이 파일의 시트: ' + ', '.join(book))

def is_sector_code(c):
    """부문 코드 여부. 대분류처럼 A, C01 등 영문 포함 코드도 허용하고,
    9000 이상의 순수 숫자 코드(9090, 9590 등 합계·분석지표)만 제외한다."""
    c = str(c).strip()
    if c in ('', 'None'):
        return False
    if c.isdigit():
        return int(c) < 9000
    return True

def initialize(raw):
    global book, codes, names, L, M, V, sheets, valid
    wb = openpyxl.load_workbook(io.BytesIO(bytes(raw)), read_only=True, data_only=True)
    book = {s: list(wb[s].values) for s in wb.sheetnames}
    wb.close()
    sheets = list(book)

    s_ad = find_sheet('국산투입계수')
    base = book[s_ad]
    # 행 코드(첫 열)와 열 코드(5번째 행)가 모두 있는 부문만 사용
    row_codes = {str(r[0]).strip() for r in base[6:] if r and r[0] is not None}
    codes = [str(c).strip() for c in base[4][2:]
             if c is not None and is_sector_code(c) and str(c).strip() in row_codes]
    if len(codes) < 2 or len(set(codes)) != len(codes):
        raise ValueError('부문 코드를 읽지 못했습니다. (읽은 부문 수: %d) 시트 구조를 확인하세요.' % len(codes))
    nm = {str(r[0]).strip(): str(r[1]) for r in base[6:] if r and r[0] is not None and len(r) > 1}
    names = [nm.get(c, c) for c in codes]

    def matrix(sheet_key):
        rows = book[find_sheet(sheet_key)]
        ri = {str(r[0]).strip(): i for i, r in enumerate(rows) if r and r[0] is not None}
        ci = {str(c).strip(): i for i, c in enumerate(rows[4]) if c is not None}
        a = np.array([[rows[ri[c]][ci[k]] for k in codes] for c in codes], dtype=float)
        if not np.isfinite(a).all():
            raise ValueError(sheet_key + ': 비수치 또는 결측 데이터')
        return a

    L = matrix('생산유발계수'); M = matrix('수입유발계수'); V = matrix('부가가치유발계수')
    ad = matrix('국산투입계수'); am = matrix('수입투입계수')

    # 부가가치계수 벡터: 총투입계수 시트의 부가가치 행 (9690 우선, 없으면 이름으로 탐색,
    # 그래도 없으면 1 - 중간투입계수 합)
    rows = book[find_sheet('총투입계수')]
    ci = {str(c).strip(): i for i, c in enumerate(rows[4]) if c is not None}
    vr = next((r for r in rows[6:] if r and r[0] is not None and str(r[0]).strip() == '9690'), None)
    if vr is None:
        vr = next((r for r in rows[6:] if r and len(r) > 1 and r[1] is not None and '부가가치' in str(r[1])), None)
    if vr is not None:
        v = np.array([vr[ci[c]] for c in codes], dtype=float)
    else:
        v = 1.0 - (ad + am).sum(axis=0)

    n = len(codes)
    residuals = {'생산': float(np.max(np.abs((np.eye(n) - ad) @ L - np.eye(n)))),
                 '수입': float(np.max(np.abs(am @ L - M))),
                 '부가가치': float(np.max(np.abs(v[:, None] * L - V)))}
    valid = max(residuals.values()) < 1e-8
    if not valid:
        raise ValueError('계수 검증 실패: ' + str(residuals))
    return emit({'sectors': [{'code': c, 'name': nme} for c, nme in zip(codes, names)],
                 'sheets': sheets, 'residuals': residuals})

def explore(sheet, rowcode):
    rows = book[sheet]
    choices = [{'code': str(r[0]), 'name': str(r[1])} for r in rows[6:] if len(r) > 1 and r[0] is not None and r[1] is not None]
    chosen = next((r for r in rows[6:] if r and str(r[0]) == rowcode), None)
    if chosen is None:
        chosen = next(r for r in rows[6:] if r and r[0] is not None and len(r) > 1 and r[1] is not None)
    cells = []
    for i in range(2, len(rows[4])):
        if rows[4][i] is None:
            continue
        val = chosen[i] if i < len(chosen) else None
        if isinstance(val, (float, int)):
            val = float(val)
        elif val is not None:
            val = str(val)
        cells.append({'code': str(rows[4][i]), 'name': str(rows[5][i] or ''), 'value': val})
    return emit({'rows': choices, 'selected': str(chosen[0]), 'name': str(chosen[1]),
                 'unit': str(rows[3][0] or '係数').replace('係数', '계수 / 비율: 행·열 명칭 확인'), 'cells': cells})

def simulate(payload):
    if not valid:
        raise ValueError('검증된 데이터가 없습니다.')
    shocks = json.loads(payload); f = np.zeros(len(codes))
    for s in shocks:
        a = float(s['amount'])
        if not np.isfinite(a) or a < 0:
            raise ValueError('수요 증가액은 유한한 0 이상의 수여야 합니다.')
        f[codes.index(s['code'])] += a
    if not f.sum() > 0:
        raise ValueError('한 산업 이상의 양의 수요 증가액을 입력하세요.')
    x = L @ f; m = M @ f; v = V @ f
    result = [{'code': c, 'name': n, 'production': float(a), 'imports': float(b), 'valueAdded': float(d)}
              for c, n, a, b, d in zip(codes, names, x, m, v)]
    result.sort(key=lambda r: r['production'], reverse=True)
    return emit({'demand': float(f.sum()), 'production': float(x.sum()), 'imports': float(m.sum()),
                 'valueAdded': float(v.sum()), 'multiplier': float(x.sum() / f.sum()), 'rows': result})


def linkage():
    """영향력계수(후방연쇄)·감응도계수(전방연쇄)와 유발계수 합계를 부문별로 계산."""
    if not valid:
        raise ValueError('검증된 데이터가 없습니다.')
    n = len(codes)
    col = L.sum(axis=0)          # 열합: 부문 j의 최종수요 1단위가 전 산업에 유발하는 생산
    row = L.sum(axis=1)          # 행합: 전 산업 최종수요 1단위씩이 부문 i에 유발하는 생산
    infl = col / col.mean()      # 영향력계수 (후방연쇄효과)
    sens = row / row.mean()      # 감응도계수 (전방연쇄효과)
    va = V.sum(axis=0); im = M.sum(axis=0)
    def kind(a, b):
        if a >= 1 and b >= 1: return '중추 산업 (전·후방 모두 강함)'
        if a >= 1: return '최종수요 지향 (후방연쇄 강함)'
        if b >= 1: return '중간수요 지향 (전방연쇄 강함)'
        return '독립 산업 (전·후방 모두 약함)'
    rows = [{'code': c, 'name': nm, 'influence': float(a), 'sensitivity': float(b),
             'production': float(p), 'valueAdded': float(v), 'imports': float(i), 'type': kind(a, b)}
            for c, nm, a, b, p, v, i in zip(codes, names, infl, sens, col, va, im)]
    return emit({'rows': rows})
