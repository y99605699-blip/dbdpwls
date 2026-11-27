BUILD = '2026-11-27-r2'
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
    global book, codes, names, L, M, V, sheets, valid, income, AD, AM, vcoef, fdom
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

    # 부가가치 구성: 피용자보수(9610) 영업잉여(9621) 고정자본소모(9622) 생산세(9630)
    income = None
    try:
        byc = {str(r[0]).strip(): r for r in rows[6:] if r and r[0] is not None}
        want = {'comp': '9610', 'surplus': '9621', 'dep': '9622', 'tax': '9630'}
        if all(c in byc for c in want.values()):
            tmp = {k: np.array([byc[c][ci[s]] for s in codes], dtype=float) for k, c in want.items()}
            if np.allclose(sum(tmp.values()), v, atol=1e-8):
                income = tmp
    except Exception:
        income = None

    AD = ad; AM = am; vcoef = v

    # 국산 최종수요(국산거래표의 최종수요계, 코드 9190): 수입대체 시뮬레이션의 기준 수요
    fdom = None
    try:
        r2 = book[find_sheet('국산거래표')]
        c2 = {str(c).strip(): i for i, c in enumerate(r2[4]) if c is not None}
        b2 = {str(r[0]).strip(): r for r in r2[6:] if r and r[0] is not None}
        if '9190' in c2:
            fdom = np.array([float(b2[c][c2['9190']]) for c in codes], dtype=float)
    except Exception:
        fdom = None

    n = len(codes)
    residuals = {'생산': float(np.max(np.abs((np.eye(n) - ad) @ L - np.eye(n)))),
                 '수입': float(np.max(np.abs(am @ L - M))),
                 '부가가치': float(np.max(np.abs(v[:, None] * L - V)))}
    valid = max(residuals.values()) < 1e-8
    if not valid:
        raise ValueError('계수 검증 실패: ' + str(residuals))
    return emit({'sectors': [{'code': c, 'name': nme} for c, nme in zip(codes, names)],
                 'sheets': sheets, 'residuals': residuals, 'build': BUILD})

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
    out = {'demand': float(f.sum()), 'production': float(x.sum()), 'imports': float(m.sum()),
           'valueAdded': float(v.sum()), 'multiplier': float(x.sum() / f.sum()), 'rows': result,
           'income': None}
    if income is not None:
        out['income'] = {'compensation': float((income['comp'] * x).sum()),
                         'surplus': float((income['surplus'] * x).sum()),
                         'depreciation': float((income['dep'] * x).sum()),
                         'tax': float((income['tax'] * x).sum())}
    return emit(out)


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
    cvb = np.std(L, axis=0, ddof=1) / L.mean(axis=0)   # 후방 변동계수: 파급이 특정 산업에 몰린 정도
    cvf = np.std(L, axis=1, ddof=1) / L.mean(axis=1)   # 전방 변동계수
    imp = AM.sum(axis=0)                               # 직접 수입의존도(수입 중간투입 비율)
    ext_o = ext_v = reach = None
    if fdom is not None:
        xb = L @ fdom
        dg = np.diag(L)
        ext_o = xb * col / dg / xb.sum() * 100            # k가 사라질 때 총산출 손실(%)
        ext_v = xb * va / dg / (vcoef * xb).sum() * 100   # k가 사라질 때 부가가치 손실(%)
        reach = (L @ (fdom / fdom.sum())) / dg           # k 가격이 100% 오를 때 최종수요 물가 상승분(비율)
    rows = [{'code': c, 'name': nm, 'influence': float(a), 'sensitivity': float(b),
             'production': float(p), 'valueAdded': float(vv), 'imports': float(i), 'type': kind(a, b),
             'influenceCV': float(cb), 'sensitivityCV': float(cf),
             'vaRatio': float(vr), 'importRatio': float(ir)}
            for c, nm, a, b, p, vv, i, cb, cf, vr, ir in zip(codes, names, infl, sens, col, va, im, cvb, cvf, vcoef, imp)]
    for j, r in enumerate(rows):
        r['extOutput'] = None if ext_o is None else float(ext_o[j])
        r['extVA'] = None if ext_v is None else float(ext_v[j])
        r['priceReach'] = None if reach is None else float(reach[j])
    return emit({'rows': rows})


def demand_structure():
    """최종수요 항목(민간소비·정부소비·투자·수출 등)별 유발액 구조."""
    def read(key):
        rows = book[find_sheet(key)]
        hdr = rows[4]
        cols = [(i, str(hdr[i]).strip(), str(rows[5][i] or '').replace('\n', '').strip())
                for i in range(2, len(hdr)) if hdr[i] is not None]
        byrow = {str(r[0]).strip(): r for r in rows[6:] if r and r[0] is not None}
        return cols, byrow

    cols, byrow = read('최종수요 항목별 생산유발액')
    items = [{'code': c, 'name': nm} for i, c, nm in cols if c != '9190']
    idx = {c: i for i, c, nm in cols}
    def val(r, c):
        x = r[idx[c]] if idx[c] < len(r) else None
        return float(x) if isinstance(x, (int, float)) else 0.0
    rows_out = []
    for c, nm in zip(codes, names):
        r = byrow.get(c)
        if r is None:
            continue
        by = {it['code']: val(r, it['code']) for it in items}
        total = val(r, '9190') if '9190' in idx else sum(by.values())
        rows_out.append({'code': c, 'name': nm, 'total': total, 'by': by})

    economy = {}
    for label, key in (('production', '최종수요 항목별 생산유발액'),
                       ('imports', '최종수요 항목별 수입유발액'),
                       ('valueAdded', '최종수요 항목별 부가가치유발액')):
        try:
            cc, br = read(key)
            ix = {c: i for i, c, nm in cc}
            economy[label] = {it['code']: float(sum(
                (br[s][ix[it['code']]] if s in br and isinstance(br[s][ix[it['code']]], (int, float)) else 0.0)
                for s in codes)) for it in items if it['code'] in ix}
        except Exception:
            pass
    return emit({'items': items, 'rows': rows_out, 'economy': economy})


def paths(payload):
    """한 산업의 수요가 단계별(0차·1차·2차·3차·4차 이상)로 퍼지는 경로 분해: x = f + Ad f + Ad^2 f + ..."""
    p = json.loads(payload); a = float(p['amount'])
    if not np.isfinite(a) or a <= 0:
        raise ValueError('양의 수요 증가액을 입력하세요.')
    k = codes.index(p['code']); n = len(codes)
    f = np.zeros(n); f[k] = a
    total = L @ f
    layers = []; cur = f.copy(); acc = np.zeros(n)
    for _ in range(4):
        layers.append(cur.copy()); acc += cur; cur = AD @ cur
    rest = total - acc
    labels = ['0차 · 최초 수요 충족', '1차 · 직접 투입 산업', '2차 · 투입의 투입', '3차', '4차 이상 (나머지)']
    amounts = [float(l.sum()) for l in layers] + [float(rest.sum())]
    def top(v, kk=8):
        idx = np.argsort(-v)[:kk]
        return [{'code': codes[i], 'name': names[i], 'amount': float(v[i])} for i in idx if v[i] > 0]
    return emit({'demand': a, 'total': float(total.sum()),
                 'rounds': [{'label': lb, 'amount': am_} for lb, am_ in zip(labels, amounts)],
                 'top1': top(layers[1]), 'top2': top(layers[2]), 'topAll': top(total, 10)})

def import_substitution(payload):
    """수입 중간재의 r%를 같은 산업의 국산품으로 대체했을 때(가격·품질 동일 가정) 기준 국산 최종수요에서의 효과."""
    if fdom is None:
        raise ValueError('이 파일에서 국산 최종수요(국산거래표의 최종수요계)를 읽지 못해 계산할 수 없습니다.')
    p = json.loads(payload); r = float(p['ratio']) / 100.0
    if not (0 < r <= 1):
        raise ValueError('대체 비율은 0 초과 100 이하로 입력하세요.')
    n = len(codes)
    targets = list(range(n)) if p['code'] == 'ALL' else [codes.index(p['code'])]
    Ad2 = AD.copy(); Am2 = AM.copy()
    for k in targets:
        Ad2[:, k] += r * AM[:, k]
        Am2[:, k] = (1 - r) * AM[:, k]
    L2 = np.linalg.inv(np.eye(n) - Ad2)
    x0 = L @ fdom; x1 = L2 @ fdom
    imp0 = float((AM * x0[None, :]).sum()); imp1 = float((Am2 * x1[None, :]).sum())
    va0 = float((vcoef * x0).sum()); va1 = float((vcoef * x1).sum())
    replaceable = float(sum((AM[:, k] * x0[k]).sum() for k in targets))
    d = x1 - x0
    idx = np.argsort(-d)[:12]
    return emit({'ratio': r * 100, 'targets': len(targets), 'replaceable': replaceable,
                 'production0': float(x0.sum()), 'production1': float(x1.sum()),
                 'valueAdded0': va0, 'valueAdded1': va1,
                 'imports0': imp0, 'imports1': imp1,
                 'rows': [{'code': codes[i], 'name': names[i], 'delta': float(d[i]), 'base': float(x0[i])}
                          for i in idx if d[i] > 0]})

def evaluate(payload):
    """사업별 경제효과 비교(예산 1조원 = 1,000,000백만원 환산 포함)."""
    items = json.loads(payload); n = len(codes); out = []
    for it in items:
        a = float(it['amount'])
        if not np.isfinite(a) or a <= 0:
            raise ValueError('예산은 양수여야 합니다.')
        f = np.zeros(n); f[codes.index(it['code'])] = a
        x = L @ f
        pr = float(x.sum()); va = float((V @ f).sum()); im = float((M @ f).sum())
        cp = float((income['comp'] * x).sum()) if income is not None else None
        s = 1000000.0 / a
        out.append({'code': it['code'], 'name': it['name'], 'amount': a,
                    'production': pr, 'valueAdded': va, 'imports': im, 'compensation': cp,
                    'multiplier': pr / a, 'vaMultiplier': va / a,
                    'per': {'production': pr * s, 'valueAdded': va * s, 'imports': im * s,
                            'compensation': (cp * s if cp is not None else None)}})
    return emit({'rows': out})


def price_shock(payload):
    """가격 파급(Leontief 가격 모형, 완전 전가 가정).
    domestic: 산업 k의 가격이 a만큼 오를 때 dp_j = a * L[k,j] / L[k,k]
    import  : 수입재 i의 가격이 d_i 오를 때 비용상승 c_j = sum_i Am[i,j] d_i,  dp = L^T c"""
    p = json.loads(payload); a = float(p['pct']) / 100.0
    if not np.isfinite(a) or a <= 0:
        raise ValueError('양의 상승률을 입력하세요.')
    n = len(codes)
    if p['mode'] == 'domestic':
        if p['code'] == 'ALL':
            raise ValueError('국내 가격 인상은 특정 산업을 선택하세요.')
        k = codes.index(p['code'])
        dp = a * L[k, :] / L[k, k]
        direct = a * AD[k, :].copy(); direct[k] = 0.0    # 1차 비용상승(타 산업이 투입하는 k 비용 증가)
    else:
        delta = np.full(n, a) if p['code'] == 'ALL' else np.where(np.arange(n) == codes.index(p['code']), a, 0.0)
        c = (AM * delta[:, None]).sum(axis=0)
        dp = L.T @ c
        direct = c
    idx_all = np.argsort(-dp)
    surplus = income['surplus'] if income is not None else None
    rows = []
    for i in idx_all[:15]:
        sq = None
        if surplus is not None and surplus[i] > 1e-12:
            sq = float(direct[i] / surplus[i] * 100)
        rows.append({'code': codes[i], 'name': names[i], 'dp': float(dp[i] * 100),
                     'direct': float(direct[i] * 100), 'squeeze': sq})
    index = None
    if fdom is not None:
        index = float((fdom / fdom.sum() * dp).sum() * 100)
    return emit({'mode': p['mode'], 'pct': a * 100, 'index': index, 'mean': float(dp.mean() * 100),
                 'max': float(dp.max() * 100), 'rows': rows})

def extraction(payload):
    """가상 제거법: 산업 k가 사라진다면(생산·투입 불가, 해당 최종수요도 소멸) 나머지 경제의 생산·부가가치 손실."""
    if fdom is None:
        raise ValueError('이 파일에서 국산 최종수요를 읽지 못해 계산할 수 없습니다.')
    k = codes.index(json.loads(payload)['code']); n = len(codes)
    keep = [i for i in range(n) if i != k]
    Lm = L[np.ix_(keep, keep)] - np.outer(L[keep, k], L[k, keep]) / L[k, k]
    x = L @ fdom
    xnew = Lm @ fdom[keep]
    d = x[keep] - xnew
    out_loss = float(x[k] + d.sum())
    va_loss = float(vcoef[k] * x[k] + (vcoef[keep] * d).sum())
    order = np.argsort(-d)[:12]
    return emit({'code': codes[k], 'name': names[k], 'own': float(x[k]),
                 'outputLoss': out_loss, 'outputShare': out_loss / float(x.sum()) * 100,
                 'vaLoss': va_loss, 'vaShare': va_loss / float((vcoef * x).sum()) * 100,
                 'rows': [{'code': codes[keep[i]], 'name': names[keep[i]], 'delta': float(d[i]),
                           'share': float(d[i] / x[keep[i]] * 100) if x[keep[i]] else 0.0}
                          for i in order if d[i] > 0]})
