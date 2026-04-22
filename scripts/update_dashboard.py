#!/usr/bin/env python3
"""Fetch data from Feishu and update bilibili-tl-dashboard index.html"""
import json, os, re, urllib.request, urllib.parse
from datetime import datetime, timedelta

APP_ID     = os.environ['FEISHU_APP_ID']
APP_SECRET = os.environ['FEISHU_APP_SECRET']
WIKI_TOKEN = 'UDziw3KuqiXFpckI9R5cDXd2nsd'
SHEET_ID   = 'd5954c'

# ── Quarter / Model normalisations ──────────────────────────────────────────
QUARTER_MAP = {'Q2': '25年Q2', 'Q3': '25年Q3', 'Q4': '25年Q4'}
MODEL_MAP   = {'Livis': 'Livis眼镜'}

def get_token():
    payload = json.dumps({'app_id': APP_ID, 'app_secret': APP_SECRET}).encode()
    req = urllib.request.Request(
        'https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal',
        data=payload, method='POST', headers={'Content-Type': 'application/json'})
    with urllib.request.urlopen(req) as r:
        return json.loads(r.read())['tenant_access_token']

def get_sheet_token(token):
    url = f'https://open.feishu.cn/open-apis/wiki/v2/spaces/get_node?obj_type=wiki&token={WIKI_TOKEN}'
    req = urllib.request.Request(url, headers={'Authorization': f'Bearer {token}'})
    with urllib.request.urlopen(req) as r:
        return json.loads(r.read())['data']['node']['obj_token']

def fetch_rows(token, sheet_token):
    rng = urllib.parse.quote(SHEET_ID + '!A1:R400')
    url = (f'https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/'
           f'{sheet_token}/values/{rng}')
    req = urllib.request.Request(url, headers={'Authorization': f'Bearer {token}'})
    with urllib.request.urlopen(req) as r:
        return json.loads(r.read())['data']['valueRange']['values']

def excel_date(serial):
    try:
        d = datetime(1899, 12, 30) + timedelta(days=int(serial))
        return d.strftime('%Y-%m-%d')
    except Exception:
        return ''

def safe_float(v):
    if v is None or v == '':
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    # Handle "6009.6/1.06" style expressions
    if re.match(r'^[\d.]+/[\d.]+$', s):
        a, b = s.split('/')
        try:
            return float(a) / float(b)
        except Exception:
            return 0.0
    try:
        return float(s)
    except Exception:
        return 0.0

def get_url(v):
    if isinstance(v, list) and v and isinstance(v[0], dict):
        return v[0].get('link', '')
    return str(v) if isinstance(v, str) else ''

def extract_bvid(v, url_cell):
    """Return BVID string; resolve MID(...) formula by parsing the URL."""
    if isinstance(v, str):
        s = v.strip()
        if s.upper().startswith('MID('):
            url = get_url(url_cell)
            m = re.search(r'bilibili\.com/video/(BV\w+)', url)
            return m.group(1) if m else ''
        if s.startswith('BV') or s.startswith('bv'):
            return s
    return ''

def process_rows(rows):
    result = []
    for row in rows[1:]:           # skip header
        if not row or not row[0]:  # skip empty rows
            continue
        while len(row) < 18:
            row.append(None)

        quarter = QUARTER_MAP.get(str(row[1] or '').strip(), str(row[1] or '').strip())
        model   = MODEL_MAP.get(str(row[5] or '').strip(), str(row[5] or '').strip())

        cost         = safe_float(row[11])
        views        = int(safe_float(row[12]))
        interactions = int(safe_float(row[15]))

        # CPM: if formula like "N2*1000", compute from 播放成本 (col 13)
        cpm_raw = row[14]
        if isinstance(cpm_raw, str) and re.match(r'^[A-Za-z]\d+\*1000$', cpm_raw.strip()):
            cpm = safe_float(row[13]) * 1000
        else:
            cpm = safe_float(cpm_raw)

        cpe  = round(cost / interactions, 2) if interactions > 0 else 0
        bvid = extract_bvid(row[16], row[3])

        result.append({
            'upName':       str(row[4]  or '').strip(),
            'model':        model,
            'upCat':        str(row[17] or '').strip(),
            'targetType':   str(row[7]  or '').strip(),
            'quarter':      quarter,
            'date':         excel_date(row[2]),
            'views':        views,
            'cost':         round(cost, 2),
            'cpe':          cpe,
            'cpm':          round(cpm, 2),
            'interactions': interactions,
            'bvid':         bvid,
        })
    return result

def update_html(data):
    with open('index.html', 'r', encoding='utf-8') as f:
        html = f.read()
    new_embedded = 'const EMBEDDED = ' + json.dumps(data, ensure_ascii=False) + ';'
    html2 = re.sub(r'const EMBEDDED\s*=\s*\[.*?\];', new_embedded, html, flags=re.DOTALL)
    if html2 == html:
        print('WARNING: EMBEDDED pattern not found, index.html not updated')
        return
    with open('index.html', 'w', encoding='utf-8') as f:
        f.write(html2)
    print(f'Updated index.html with {len(data)} records')

if __name__ == '__main__':
    print('Getting access token...')
    token = get_token()
    print('Getting sheet token...')
    sheet_token = get_sheet_token(token)
    print(f'Sheet token: {sheet_token}')
    print('Fetching rows...')
    rows = fetch_rows(token, sheet_token)
    print(f'Raw rows: {len(rows)}')
    data = process_rows(rows)
    print(f'Processed: {len(data)} records')
    update_html(data)
