#!/usr/bin/env python3
# shopify_tag.py - 当選者にShopifyタグを付与するスクリプト

import requests
import time

# ==============================
# 設定
# ==============================
SHOP = 'crewre0826.myshopify.com'
TOKEN = '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'
TAG = 'popup_2026_05'
API_VERSION = '2024-01'

HEADERS = {
    'X-Shopify-Access-Token': TOKEN,
    'Content-Type': 'application/json'
}

# ==============================
# 関数
# ==============================

def get_customer_by_email(email):
    """メールアドレスで顧客を検索"""
    url = f'https://{SHOP}/admin/api/{API_VERSION}/customers/search.json'
    r = requests.get(url, headers=HEADERS, params={'query': f'email:{email}', 'limit': 1})
    r.raise_for_status()
    customers = r.json().get('customers', [])
    return customers[0] if customers else None

def add_tag(customer_id, existing_tags, new_tag):
    """顧客にタグを追加（既存タグを保持）"""
    tags = [t.strip() for t in existing_tags.split(',') if t.strip()]
    if new_tag in tags:
        return False  # すでにタグあり

    tags.append(new_tag)
    url = f'https://{SHOP}/admin/api/{API_VERSION}/customers/{customer_id}.json'
    r = requests.put(url, headers=HEADERS, json={'customer': {'id': customer_id, 'tags': ', '.join(tags)}})
    r.raise_for_status()
    return True

def tag_winners(winner_emails):
    """当選者メールリストにタグを一括付与"""
    print(f'タグ付与開始: {TAG}')
    print(f'対象: {len(winner_emails)}件')
    print('=' * 40)

    success = 0
    already = 0
    not_found = 0
    errors = []

    for i, email in enumerate(winner_emails, 1):
        try:
            customer = get_customer_by_email(email)
            if not customer:
                print(f'[{i:3d}] 未登録  {email}')
                not_found += 1
            else:
                added = add_tag(customer['id'], customer.get('tags', ''), TAG)
                if added:
                    print(f'[{i:3d}] ✓ タグ付与  {email}')
                    success += 1
                else:
                    print(f'[{i:3d}] - 既存タグ  {email}')
                    already += 1
        except Exception as e:
            print(f'[{i:3d}] ERROR  {email}: {e}')
            errors.append(email)

        time.sleep(0.3)  # APIレート制限対策

    print('=' * 40)
    print(f'完了: 付与={success} / 既存={already} / 未登録={not_found} / エラー={len(errors)}')
    if errors:
        print('エラーのメールアドレス:', errors)


# ==============================
# 実行（lottery.pyと組み合わせて使う場合）
# ==============================

if __name__ == '__main__':
    import sys

    if len(sys.argv) > 1 and sys.argv[1] == 'test':
        # テストモード：1件だけ試す
        test_email = sys.argv[2] if len(sys.argv) > 2 else input('テスト用メールアドレス: ')
        tag_winners([test_email])
    else:
        # lottery.pyの当選者データを読み込んで実行
        try:
            from lottery import load_paperform, load_shopify, run_lottery
            print('抽選データを読み込んでタグ付けします...')
            applicants = load_paperform(load_shopify())
            winners, _ = run_lottery(applicants)
            emails = [w['email'] for w in winners]
            tag_winners(emails)
        except Exception as e:
            print(f'エラー: {e}')
            print('使い方: python3 shopify_tag.py test <メールアドレス>')
