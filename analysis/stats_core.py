"""Loads the de-identified cohort and computes every number used in the manuscript."""
import itertools
import numpy as np
import pandas as pd
import openpyxl
from scipy import stats

import os
# Patient-level data: never commit this file to the repository
XLSX = os.environ.get('ARTHRO_XLSX', '/root/.claude/uploads/cb2515e8-cb13-5a97-91fb-80f392b837dd/7773a8a5-__3._ArthroBrostrom____update_ver2.5.xlsx')
COLS = {1: 'id', 3: 'sex', 4: 'age', 7: 'side', 8: 'opdate', 18: 'bmi', 24: 'beighton',
        25: 'pre_vas', 26: 'pre_aofas', 27: 'pre_kp', 36: 'teg_pre', 37: 'teg_post',
        38: 'vas', 39: 'aofas', 40: 'kp', 41: 'ffi', 42: 'faos', 43: 'faos_pain', 44: 'faos_sym',
        45: 'faos_adl', 46: 'faos_sport', 47: 'faos_qol', 48: 'eq5d', 49: 'faam_adl', 50: 'faam_sport',
        55: 'pre_tt', 56: 'pre_at', 57: 'c_pre_tt', 58: 'c_pre_at', 61: 'tt', 62: 'at',
        63: 'ctratio', 66: 'c_tt', 67: 'c_at'}
GROUPS = ['A', 'S', 'N']
GNAME = {'A': 'Anatomic', 'S': 'Subanatomic', 'N': 'Nonanatomic'}


def _num(x):
    try:
        return float(x)
    except (TypeError, ValueError):
        return np.nan


def _ratio(x):
    try:
        a, b = str(x).split('/')
        return float(a) / float(b) * 100
    except ValueError:
        return np.nan


def load():
    ws = openpyxl.load_workbook(XLSX, data_only=True)['Sheet1']
    rows = list(ws.iter_rows(values_only=True))[1:]
    d = pd.DataFrame([{v: r[k] for k, v in COLS.items()} for r in rows])
    for c in d.columns:
        if c not in ('opdate', 'ctratio', 'sex'):
            d[c] = d[c].map(_num)
    d['sex'] = d['sex'].map(lambda v: {1: 'M', 2: 'F', 'M': 'M', 'F': 'F'}.get(v))
    d['opdate'] = pd.to_datetime(d['opdate'])
    d['ratio'] = d['ctratio'].map(_ratio)
    # Study cohort: postoperative 3D CT ratio available, operated before 2025 (81 ankles)
    s = d[d.ratio.notna() & (d.opdate < '2025-06-01')].copy()
    s['G'] = pd.cut(s.ratio, [-1, 25, 50, 1000], right=False, labels=GROUPS).astype(str)
    s['d_vas'] = s.vas - s.pre_vas
    s['d_kp'] = s.kp - s.pre_kp
    s['d_aofas'] = s.aofas - s.pre_aofas
    s['d_teg'] = s.teg_post - s.teg_pre
    s['ssd_pre_tt'] = s.pre_tt - s.c_pre_tt
    s['ssd_pre_at'] = s.pre_at - s.c_pre_at
    s['ssd_tt'] = s.tt - s.c_tt
    s['ssd_at'] = s['at'] - s.c_at
    s['d_tt'] = s.tt - s.pre_tt
    s['d_at'] = s['at'] - s.pre_at
    return s


S = load()
XR = S[S.tt.notna() & S['at'].notna()].copy()   # ankles with final stress radiographs


def fp(p):
    if p is None or np.isnan(p):
        return '—'
    return '<.001' if p < .001 else ('%.3f' % p).lstrip('0')


def ms(x, dec=2):
    x = pd.Series(x).dropna()
    return f'{x.mean():.{dec}f} ± {x.std():.{dec}f}'


def kw(df, v):
    grs = [df.loc[df.G == g, v].dropna() for g in GROUPS]
    return stats.kruskal(*grs).pvalue


def mwu_pairs(df, v):
    out = {}
    for a, b in itertools.combinations(GROUPS, 2):
        out[a + b] = stats.mannwhitneyu(df.loc[df.G == a, v].dropna(), df.loc[df.G == b, v].dropna()).pvalue
    return out


def wil(df, a, b):
    x = df[[a, b]].dropna()
    return stats.wilcoxon(x[a], x[b]).pvalue


def chi(df, col):
    ct = pd.crosstab(df.G, df[col])
    return stats.chi2_contingency(ct)[1]


def rho(df, v):
    x = df[['ratio', v]].dropna()
    r = stats.spearmanr(x.ratio, x[v])
    return r.statistic, r.pvalue, len(x)
