#!/usr/bin/env python3
# sheets_helper.py - Google Sheets API helper (OAuth2 refresh token)

import json
import os
import urllib.request
import urllib.parse
import time

_token_cache = {'token': None, 'expires': 0}

# Streamlit Cloud: st.secretsから、ローカル: 環境変数 or gcloudから取得
def _get_credentials():
    """OAuth2クレデンシャルを取得"""
    try:
        import streamlit as st
        if hasattr(st, 'secrets') and 'gcp' in st.secrets:
            return st.secrets['gcp']
    except Exception:
        pass
    # ローカルフォールバック: gcloudのcredentials.dbから取得
    import sqlite3
    db_path = os.path.expanduser('~/.config/gcloud/credentials.db')
    if os.path.exists(db_path):
        db = sqlite3.connect(db_path)
        for row in db.execute('SELECT * FROM credentials'):
            return json.loads(row[1])
    raise Exception('Google認証情報が見つかりません。Streamlit secretsまたはgcloud auth loginを設定してください。')


def get_access_token():
    """OAuth2リフレッシュトークンでアクセストークンを取得（キャッシュ付き）"""
    now = time.time()
    if _token_cache['token'] and now < _token_cache['expires']:
        return _token_cache['token']
    creds = _get_credentials()
    data = urllib.parse.urlencode({
        'client_id': creds['client_id'],
        'client_secret': creds['client_secret'],
        'refresh_token': creds['refresh_token'],
        'grant_type': 'refresh_token',
    }).encode()
    req = urllib.request.Request('https://oauth2.googleapis.com/token', data=data)
    resp = urllib.request.urlopen(req)
    result = json.loads(resp.read())
    _token_cache['token'] = result['access_token']
    _token_cache['expires'] = now + 3000  # 50分キャッシュ
    return _token_cache['token']


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
