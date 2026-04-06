#!/usr/bin/env python3
# sheets_helper.py - Google Sheets API helper (gcloud auth token)

import json
import subprocess
import urllib.request
import urllib.parse
import time

_token_cache = {'token': None, 'expires': 0}


def get_access_token():
    """gcloud auth print-access-token でアクセストークンを取得（キャッシュ付き）"""
    now = time.time()
    if _token_cache['token'] and now < _token_cache['expires']:
        return _token_cache['token']
    token = subprocess.check_output(
        ["gcloud", "auth", "print-access-token"], stderr=subprocess.DEVNULL
    ).decode().strip()
    _token_cache['token'] = token
    _token_cache['expires'] = now + 3000  # 50分キャッシュ
    return token


def _api(method, url, body=None, retry=True):
    """Sheets API呼び出し"""
    token = get_access_token()
    data = json.dumps(body).encode() if body else None
    req = urllib.request.Request(url, data=data, headers={
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }, method=method)
    try:
        resp = urllib.request.urlopen(req)
        return json.loads(resp.read())
    except urllib.error.HTTPError as e:
        if e.code == 401 and retry:
            _token_cache['token'] = None
            _token_cache['expires'] = 0
            return _api(method, url, body, retry=False)
        raise Exception(f"Sheets API error {e.code}: {e.read().decode()[:300]}")


def read_sheet(spreadsheet_id, sheet_name):
    """シートを読み込み → list[dict] を返す"""
    encoded = urllib.parse.quote(sheet_name)
    url = f"https://sheets.googleapis.com/v4/spreadsheets/{spreadsheet_id}/values/{encoded}?valueRenderOption=FORMATTED_VALUE"
    result = _api("GET", url)
    rows = result.get('values', [])
    if len(rows) < 2:
        return []
    headers = rows[0]
    data = []
    for row in rows[1:]:
        # 列数が足りない場合は空文字で補完
        padded = row + [''] * (len(headers) - len(row))
        data.append({headers[i]: padded[i] for i in range(len(headers))})
    return data


def write_sheet(spreadsheet_id, sheet_name, headers, rows):
    """シートにデータを書き込み（シートがなければ作成、あればクリアして上書き）"""
    base = f"https://sheets.googleapis.com/v4/spreadsheets/{spreadsheet_id}"

    # シート存在確認
    meta = _api("GET", f"{base}?fields=sheets.properties")
    existing = {s['properties']['title']: s['properties']['sheetId'] for s in meta['sheets']}

    if sheet_name not in existing:
        result = _api("POST", f"{base}:batchUpdate", {
            "requests": [{"addSheet": {"properties": {"title": sheet_name}}}]
        })
        sheet_id = result['replies'][0]['addSheet']['properties']['sheetId']
    else:
        sheet_id = existing[sheet_name]
        encoded = urllib.parse.quote(sheet_name)
        _api("POST", f"{base}/values/{encoded}:clear", {})

    # データ書き込み
    encoded = urllib.parse.quote(sheet_name)
    all_rows = [headers] + rows
    _api("PUT", f"{base}/values/{encoded}!A1?valueInputOption=USER_ENTERED", {
        "values": all_rows
    })

    # ヘッダー太字 + 固定
    _api("POST", f"{base}:batchUpdate", {
        "requests": [
            {
                "repeatCell": {
                    "range": {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": 1},
                    "cell": {"userEnteredFormat": {"textFormat": {"bold": True}}},
                    "fields": "userEnteredFormat.textFormat.bold"
                }
            },
            {
                "updateSheetProperties": {
                    "properties": {"sheetId": sheet_id, "gridProperties": {"frozenRowCount": 1}},
                    "fields": "gridProperties.frozenRowCount"
                }
            },
            {
                "autoResizeDimensions": {
                    "dimensions": {"sheetId": sheet_id, "dimension": "COLUMNS",
                                   "startIndex": 0, "endIndex": len(headers)}
                }
            },
        ]
    })

    return sheet_id
