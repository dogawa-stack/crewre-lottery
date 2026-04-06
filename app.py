#!/usr/bin/env python3
# app.py - crewre 抽選管理システム

import streamlit as st
import pandas as pd
import json
import os
import time
import tempfile
import sys
import random
import requests
from collections import defaultdict

LOTTERY_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, LOTTERY_DIR)

from lottery import load_shopify, load_applicants, run_lottery, SLOT_DEFS, build_slot_defs, build_time_zone_map, DEFAULT_DAYS, DEFAULT_SLOT_TIMES, SLOT_NAMES
from shopify_tag import get_customer_by_email, add_tag, SHOP, TAG
from email_templates import (
    SUBJECT_WINNER,     BODY_WINNER,
    SUBJECT_LOSER,      BODY_LOSER,
    SUBJECT_WINNER_2ND, BODY_WINNER_2ND,
    SUBJECT_REMINDER,   BODY_REMINDER,
    SUBJECT_THANKS,     BODY_THANKS,
)

RESEND_API_KEY = 're_67fDJotj_H9yrdVfb93TCWfry9Fhvtrm5'
FROM_EMAIL     = 'crewre <crewre@modern-times.co>'
STATE_FILE     = os.path.join(LOTTERY_DIR, 'lottery_state.json')

# ==============================
# Page config
# ==============================

st.set_page_config(page_title='crewre 抽選管理', page_icon='🎯', layout='wide')

# ==============================
# State management
# ==============================

def default_state():
    return {
        'phase': 1,
        'winners': [],
        'losers': [],
        'second_winners': [],
        'sent_modes': [],
        'absent_winner_emails': [],
        'settings': {
            'vip': 3, 'latent': 3, 'new': 4,
            'capacity': 10,
        },
        'url_fields': {
            'attendance_url': '', 'attendance_deadline': '',
            'presale_url': '', 'presale_deadline': '', 'survey_url': '',
        },
        'event_config': {
            'days': DEFAULT_DAYS,
            'slot_times': DEFAULT_SLOT_TIMES,
        },
    }

def load_state():
    if os.path.exists(STATE_FILE):
        try:
            with open(STATE_FILE, encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            pass
    return default_state()

def persist():
    state = {k: st.session_state.get(k, default_state()[k]) for k in default_state()}
    with open(STATE_FILE, 'w', encoding='utf-8') as f:
        json.dump(state, f, ensure_ascii=False, indent=2)

if 'initialized' not in st.session_state:
    for k, v in load_state().items():
        st.session_state[k] = v
    st.session_state.initialized = True

# ==============================
# Email helpers
# ==============================

def fill_template(body, name, slot='', checkin_id='', url_fields=None):
    result = body.replace('{{Name}}', str(name)).replace('{{Slot}}', str(slot)).replace('{{ID}}', str(checkin_id))
    if url_fields:
        result = (result
            .replace('{{出欠フォームURL}}',  url_fields.get('attendance_url', ''))
            .replace('{{回答締切}}',         url_fields.get('attendance_deadline', ''))
            .replace('{{先行販売URL}}',      url_fields.get('presale_url', ''))
            .replace('{{先行販売締切}}',      url_fields.get('presale_deadline', ''))
            .replace('{{アンケートURL}}',    url_fields.get('survey_url', ''))
        )
    return result

def send_one(to_email, subject, body):
    resp = requests.post(
        'https://api.resend.com/emails',
        headers={'Authorization': f'Bearer {RESEND_API_KEY}', 'Content-Type': 'application/json'},
        json={'from': FROM_EMAIL, 'to': [to_email], 'subject': subject, 'text': body},
    )
    return resp.status_code == 200

def send_bulk(recipients, subject, body_tpl, url_fields=None, is_loser=False):
    ok, ng, errors = 0, 0, []
    bar  = st.progress(0)
    info = st.empty()
    total = len(recipients)
    for i, r in enumerate(recipients):
        name  = r.get('name', '')
        email = r.get('email', '')
        slot  = '' if is_loser else r.get('slot', '')
        cid   = '' if is_loser else r.get('checkin_id', '')
        if not email or not name:
            continue
        body = fill_template(body_tpl, name, slot, cid, url_fields)
        if send_one(email, subject, body):
            ok += 1
        else:
            ng += 1
            errors.append(f'{name} <{email}>')
        bar.progress((i + 1) / total)
        info.caption(f'送信中… {i+1}/{total}')
        time.sleep(0.1)
    bar.empty(); info.empty()
    return ok, ng, errors

# ==============================
# Lottery with custom settings
# ==============================

def run_lottery_with_settings(shopify_path, paperform_path, slot_capacity_override=None):
    """抽選実行。slot_capacity_override={slot_id: capacity} で空き枠を指定できる（二次用）"""
    import lottery as _lot

    # 設定を一時的に反映
    s = st.session_state.settings
    ec = st.session_state.get('event_config', default_state()['event_config'])
    orig_alloc     = dict(_lot.ALLOCATION)
    orig_capacity  = _lot.CAPACITY_PER_SLOT
    orig_slot_defs = dict(_lot.SLOT_DEFS)
    orig_tz_map    = dict(_lot.TIME_ZONE_MAP)

    _lot.ALLOCATION    = {'VIP': s['vip'], '潜在': s['latent'], '新規': s['new']}
    _lot.CAPACITY_PER_SLOT = s['capacity']
    _lot.SLOT_DEFS     = build_slot_defs(ec['days'], ec.get('slot_times_per_day', ec.get('slot_times', DEFAULT_SLOT_TIMES)))
    _lot.TIME_ZONE_MAP = build_time_zone_map(ec['days'], ec.get('slot_times_per_day', ec.get('slot_times', DEFAULT_SLOT_TIMES)))

    if slot_capacity_override:
        # 二次抽選: slot_total を空き枠に制限
        orig_slot_defs = dict(_lot.SLOT_DEFS)
        # SLOT_DEFS の中で空き0の枠は抽選対象外にする
        _lot.SLOT_DEFS = {sid: name for sid, name in _lot.SLOT_DEFS.items()
                         if slot_capacity_override.get(sid, 0) > 0}

    shopify    = load_shopify(shopify_path)
    applicants = load_applicants(paperform_path, shopify)
    random.seed()
    winners, losers = run_lottery(applicants)

    # チェックインID振り直し
    by_slot = defaultdict(list)
    for w in winners:
        by_slot[w['slot_id']].append(w)
    new_id = 1
    for sid in sorted(orig_slot_defs.keys() if slot_capacity_override else _lot.SLOT_DEFS.keys()):
        for w in sorted(by_slot.get(sid, []), key=lambda x: x['checkin_id']):
            w['checkin_id'] = new_id
            new_id += 1

    # 元に戻す
    _lot.ALLOCATION        = orig_alloc
    _lot.CAPACITY_PER_SLOT = orig_capacity
    _lot.SLOT_DEFS         = orig_slot_defs
    _lot.TIME_ZONE_MAP     = orig_tz_map
    if slot_capacity_override:
        _lot.SLOT_DEFS = orig_slot_defs

    return winners, losers, applicants

# ==============================
# Phase indicator
# ==============================

st.title('🎯 crewre 抽選管理システム')

phase_labels = ['① 一次抽選', '② 出欠確認', '③ 二次抽選', '④ 直前案内・お礼']
cur = st.session_state.phase
cols = st.columns(4)
for i, (col, label) in enumerate(zip(cols, phase_labels)):
    with col:
        if i + 1 < cur:
            st.success(f'✅ {label}')
        elif i + 1 == cur:
            st.info(f'▶ {label}')
        else:
            st.markdown(f'<div style="color:#bbb;padding:8px">{label}</div>', unsafe_allow_html=True)

st.divider()

# ==============================
# PHASE 1: 一次抽選
# ==============================

if cur == 1:
    st.header('① 一次抽選')

    # イベント設定
    with st.expander('📅 イベント設定（日程・時間帯）', expanded=False):
        if 'event_config' not in st.session_state:
            st.session_state.event_config = default_state()['event_config']
        ec = st.session_state.event_config
        days = ec.get('days', DEFAULT_DAYS)
        slot_times_per_day = ec.get('slot_times_per_day', [DEFAULT_SLOT_TIMES] * len(days))
        if len(slot_times_per_day) != len(days):
            slot_times_per_day = [DEFAULT_SLOT_TIMES] * len(days)

        # --- 開催日数 ---
        num_days = st.number_input('開催日数', min_value=1, max_value=5, value=len(days), step=1, key='num_days')

        day_labels = []
        for i in range(num_days):
            default_label = days[i] if i < len(days) else f'Day{i+1}'
            day_labels.append(st.text_input(f'Day {i+1} の日付', value=default_label, key=f'day_label_{i}'))

        st.divider()

        # --- 共通設定 ---
        st.subheader('⏰ スロット設定')
        gc1, gc2, gc3 = st.columns(3)
        start_h = gc1.number_input('開始時刻（時）', min_value=0, max_value=23, value=10, step=1)
        start_m = gc2.number_input('開始時刻（分）', min_value=0, max_value=59, value=0, step=30)
        interval = gc3.number_input('間隔（分）', min_value=10, max_value=120, value=30, step=5)

        gc4, _ = st.columns(2)
        duration = gc4.number_input('1枠の長さ（分）', min_value=30, max_value=180, value=60, step=30)

        def generate_slots(s_h, s_m, e_h, e_m, intv, dur):
            slots = []
            cur_min = s_h * 60 + s_m
            end_min = e_h * 60 + e_m
            while cur_min + dur <= end_min:
                sh, sm = divmod(cur_min, 60)
                eh, em = divmod(cur_min + dur, 60)
                slots.append(f'{sh}:{sm:02d}〜{eh}:{em:02d}')
                cur_min += intv
            return slots

        st.divider()

        # --- 日ごとの終了時刻 + スロット選択 ---
        times_inputs = []
        for i, day in enumerate(day_labels):
            st.markdown(f'### {day}')
            ec1, ec2 = st.columns(2)
            default_end_h = 19 if i == 0 else 18
            day_end_h = ec1.number_input(f'閉店時刻（時）', min_value=0, max_value=23, value=default_end_h, step=1, key=f'end_h_{i}')
            day_end_m = ec2.number_input(f'閉店時刻（分）', min_value=0, max_value=59, value=0, step=30, key=f'end_m_{i}')

            day_slots = generate_slots(start_h, start_m, day_end_h, day_end_m, interval, duration)

            existing = slot_times_per_day[i] if i < len(slot_times_per_day) else []
            default_slots = existing if existing else day_slots
            selected = []
            slot_cols = st.columns(4)
            for j, slot in enumerate(day_slots):
                checked = slot in default_slots
                col = slot_cols[j % 4]
                if col.checkbox(f'{SLOT_NAMES[j]} {slot}' if j < len(SLOT_NAMES) else slot,
                               value=checked, key=f'slot_{i}_{j}'):
                    selected.append(slot)
            times_inputs.append(selected)
            st.caption(f'→ {len(selected)}枠選択中')
            st.divider()

        # プレビュー
        total_slots = sum(len(t) for t in times_inputs)
        total_people = total_slots * st.session_state.settings['capacity']
        st.info(f'📊 合計 **{total_slots}スロット** × {st.session_state.settings["capacity"]}名 = **{total_people}名**')

        if st.button('✅ 設定を保存', key='save_event', type='primary'):
            st.session_state.event_config = {'days': day_labels, 'slot_times_per_day': times_inputs}
            persist()
            st.success(f'保存しました！（{len(day_labels)}日間・合計{total_slots}スロット）')

    # 設定パネル
    with st.expander('⚙️ 抽選設定', expanded=False):
        s = st.session_state.settings
        st.caption('ステータス別配分（合計が定員と一致するよう設定してください）')
        c1, c2, c3, c4 = st.columns(4)
        s['vip']      = c1.number_input('VIP枠',  min_value=0, value=s['vip'],      step=1)
        s['latent']   = c2.number_input('潜在枠', min_value=0, value=s['latent'],   step=1)
        s['new']      = c3.number_input('新規枠', min_value=0, value=s['new'],       step=1)
        s['capacity'] = c4.number_input('1枠定員', min_value=1, value=s['capacity'], step=1)
        if s['vip'] + s['latent'] + s['new'] != s['capacity']:
            st.warning(f'⚠️ VIP+潜在+新規 = {s["vip"]+s["latent"]+s["new"]} ≠ 定員 {s["capacity"]}')
        st.session_state.settings = s
        persist()

    c1, c2 = st.columns(2)
    shopify_file   = c1.file_uploader('Shopify 顧客CSV',  type='csv', key='s1')
    paperform_file = c2.file_uploader('応募者リストCSV', type='csv', key='p1')

    if shopify_file and paperform_file:
        # アップロードされたCSVのプレビュー
        with st.expander('📋 アップロードCSV確認', expanded=False):
            try:
                shopify_file.seek(0)
                sf_preview = pd.read_csv(shopify_file, nrows=2)
                st.caption('Shopify CSV列:')
                st.code(', '.join(sf_preview.columns.tolist()[:10]) + '...')
                shopify_file.seek(0)
            except Exception as e:
                st.error(f'Shopify CSV読み込みエラー: {e}')
            try:
                paperform_file.seek(0)
                pf_preview = pd.read_csv(paperform_file, nrows=2)
                st.caption('Paperform CSV列:')
                st.code(', '.join(pf_preview.columns.tolist()[:10]) + '...')
                paperform_file.seek(0)
            except Exception as e:
                st.error(f'Paperform CSV読み込みエラー: {e}')

        if st.button('🎲 抽選実行', type='primary'):
            with st.spinner('抽選中…'):
                try:
                    shopify_file.seek(0)
                    paperform_file.seek(0)
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.csv', mode='wb') as sf:
                        sf.write(shopify_file.read()); sp = sf.name
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.csv', mode='wb') as pf:
                        pf.write(paperform_file.read()); pp = pf.name

                    winners, losers, _ = run_lottery_with_settings(sp, pp)
                    st.session_state.winners    = winners
                    st.session_state.losers     = losers
                    st.session_state.sent_modes = []
                    persist(); st.rerun()
                except Exception as e:
                    st.error(f'エラー: {e}')
                    import traceback
                    st.code(traceback.format_exc())
                finally:
                    for p in [sp, pp]:
                        try: os.unlink(p)
                        except: pass

    if st.session_state.winners:
        winners = st.session_state.winners
        losers  = st.session_state.losers

        m1, m2, m3 = st.columns(3)
        m1.metric('総応募者', f'{len(winners)+len(losers)}名')
        m2.metric('当選',     f'{len(winners)}名')
        m3.metric('落選',     f'{len(losers)}名')

        t1, t2 = st.tabs(['当選者一覧', '落選者一覧'])
        with t1:
            st.dataframe(pd.DataFrame([{
                'ID': w['checkin_id'], '氏名': w['name'],
                'メール': w['email'], 'ステータス': w['status'], '当選枠': w['slot'],
            } for w in winners]), use_container_width=True, height=300)
        with t2:
            st.dataframe(pd.DataFrame([{
                '氏名': l['name'], 'メール': l['email'], 'ステータス': l['status'],
            } for l in losers]), use_container_width=True, height=300)

        st.divider()
        st.subheader('Shopify タグ付与')
        st.caption(f'ストア: {SHOP}　タグ: {TAG}')
        sent = st.session_state.sent_modes
        if 'shopify_tag' not in sent:
            if st.button('🏷️ 当選者にタグ付与', type='primary', key='shopify_tag_btn'):
                emails = [w['email'] for w in winners]
                bar  = st.progress(0)
                info = st.empty()
                ok, already, not_found, errors = 0, 0, 0, []
                for i, email in enumerate(emails):
                    try:
                        customer = get_customer_by_email(email)
                        if not customer:
                            not_found += 1
                        else:
                            added = add_tag(customer['id'], customer.get('tags', ''), TAG)
                            if added:
                                ok += 1
                            else:
                                already += 1
                    except Exception as e:
                        errors.append(f'{email}: {e}')
                    bar.progress((i + 1) / len(emails))
                    info.caption(f'処理中… {i+1}/{len(emails)}')
                    time.sleep(0.3)
                bar.empty(); info.empty()
                if not errors:
                    st.success(f'✅ 付与={ok} / 既存={already} / 未登録={not_found}')
                    st.session_state.sent_modes = sent + ['shopify_tag']
                    persist(); st.rerun()
                else:
                    st.warning(f'付与={ok} / 既存={already} / 未登録={not_found} / エラー={len(errors)}')
                    st.error('\n'.join(errors))
        else:
            st.success('✅ Shopifyタグ付与済み')

        st.divider()
        st.subheader('メール送信')
        st.warning('⚠️ メール送信は取り消しできません。本番時のみ実行してください。')
        uf = st.session_state.url_fields
        c1, c2 = st.columns(2)
        uf['attendance_url']      = c1.text_input('出欠フォームURL', value=uf.get('attendance_url', ''))
        uf['attendance_deadline'] = c2.text_input('回答締切',        value=uf.get('attendance_deadline', ''), placeholder='例: 4/10(木) 23:59')
        st.session_state.url_fields = uf; persist()

        # メール送信ロック（チェックボックスで解除しないと押せない）
        mail_unlock = st.checkbox('🔓 メール送信のロックを解除する（本番のみ）', value=False, key='mail_unlock_1')

        sent = st.session_state.sent_modes
        c1, c2 = st.columns(2)
        with c1:
            if 'winner' not in sent:
                if st.button('📨 当選メール送信', type='primary', key='bw1', disabled=not mail_unlock):
                    if not uf.get('attendance_url'):
                        st.warning('出欠フォームURLを入力してください')
                    else:
                        ok, ng, errs = send_bulk(winners, SUBJECT_WINNER, BODY_WINNER, uf)
                        if ng == 0:
                            st.success(f'✅ {ok}件送信完了')
                            st.session_state.sent_modes = sent + ['winner']
                            persist(); st.rerun()
                        else:
                            st.error(f'失敗 {ng}件: ' + ', '.join(errs))
            else:
                st.success('✅ 当選メール送信済み')
        with c2:
            if 'loser' not in sent:
                if st.button('📨 落選メール送信', key='bl1', disabled=not mail_unlock):
                    ok, ng, errs = send_bulk(losers, SUBJECT_LOSER, BODY_LOSER, is_loser=True)
                    if ng == 0:
                        st.success(f'✅ {ok}件送信完了')
                        st.session_state.sent_modes = sent + ['loser']
                        persist(); st.rerun()
                    else:
                        st.error(f'失敗 {ng}件: ' + ', '.join(errs))
            else:
                st.success('✅ 落選メール送信済み')

        sent = st.session_state.sent_modes
        if 'winner' in sent and 'loser' in sent:
            st.divider()
            c1, c2 = st.columns(2)
            with c1:
                if st.button('→ 出欠確認へ（二次抽選あり）', type='primary'):
                    st.session_state.phase = 2; persist(); st.rerun()
            with c2:
                if st.button('→ 直前案内へ（二次抽選なし）'):
                    st.session_state.phase = 4; persist(); st.rerun()

# ==============================
# PHASE 2: 出欠確認
# ==============================

elif cur == 2:
    st.header('② 出欠確認')
    st.info('出欠フォームの締切後、回答CSVをアップロードしてください。')

    attendance_file = st.file_uploader('出欠回答CSV', type='csv')
    if attendance_file:
        df = pd.read_csv(attendance_file)
        st.write('プレビュー:', df.head(3))
        c1, c2, c3 = st.columns(3)
        email_col  = c1.selectbox('メールアドレス列', df.columns.tolist())
        attend_col = c2.selectbox('出欠列',           df.columns.tolist())
        absent_val = c3.text_input('欠席を表す値', value='欠席')

        if st.button('欠席者を確定', type='primary'):
            absent_emails = set(
                df[df[attend_col].astype(str).str.contains(absent_val, na=False)]
                [email_col].str.strip().str.lower().tolist()
            )
            absent_winners = [w for w in st.session_state.winners if w['email'].lower() in absent_emails]
            st.session_state.absent_winner_emails = list(absent_emails)
            persist()
            st.success(f'欠席者: {len(absent_winners)}名 / 空き枠: {len(absent_winners)}枠')
            if absent_winners:
                st.dataframe(pd.DataFrame([{'氏名': w['name'], '枠': w['slot']} for w in absent_winners]),
                             use_container_width=True)

    absent_count = len(st.session_state.get('absent_winner_emails', []))
    c1, c2 = st.columns(2)
    with c1:
        if st.button('← Phase 1に戻る'):
            st.session_state.phase = 1; persist(); st.rerun()
    with c2:
        if st.button(f'→ 二次抽選へ（空き{absent_count}枠）', type='primary'):
            st.session_state.phase = 3; persist(); st.rerun()

# ==============================
# PHASE 3: 二次抽選（新規募集）
# ==============================

elif cur == 3:
    st.header('③ 二次抽選')

    absent_emails  = set(st.session_state.get('absent_winner_emails', []))
    absent_winners = [w for w in st.session_state.winners if w['email'].lower() in absent_emails]

    slot_counts = defaultdict(int)
    for w in absent_winners:
        slot_counts[w['slot_id']] += 1

    c1, c2 = st.columns(2)
    c1.metric('空き枠数', f'{len(absent_winners)}枠')

    if absent_winners:
        with st.expander('空き枠詳細'):
            st.dataframe(pd.DataFrame([{'枠': SLOT_DEFS[sid], '空き人数': cnt}
                                        for sid, cnt in slot_counts.items()]),
                         use_container_width=True)

    st.info('二次募集の応募CSVをアップロードして抽選してください。')

    # 設定パネル（二次用）
    with st.expander('⚙️ 抽選設定', expanded=False):
        s = st.session_state.settings
        c1, c2, c3, c4 = st.columns(4)
        s['vip']      = c1.number_input('VIP枠',  min_value=0, value=s['vip'],      step=1, key='sv2')
        s['latent']   = c2.number_input('潜在枠', min_value=0, value=s['latent'],   step=1, key='sl2')
        s['new']      = c3.number_input('新規枠', min_value=0, value=s['new'],       step=1, key='sn2')
        s['capacity'] = c4.number_input('1枠定員', min_value=1, value=s['capacity'], step=1, key='sc2')
        st.session_state.settings = s; persist()

    c1, c2 = st.columns(2)
    shopify_file2   = c1.file_uploader('Shopify 顧客CSV',      type='csv', key='s2')
    paperform_file2 = c2.file_uploader('Paperform 二次応募CSV', type='csv', key='p2')

    if shopify_file2 and paperform_file2:
        if st.button('🎲 二次抽選実行', type='primary'):
            with st.spinner('二次抽選中…'):
                try:
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.csv', mode='wb') as sf:
                        sf.write(shopify_file2.read()); sp = sf.name
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.csv', mode='wb') as pf:
                        pf.write(paperform_file2.read()); pp = pf.name

                    sw, _, _ = run_lottery_with_settings(sp, pp, slot_capacity_override=slot_counts)
                    st.session_state.second_winners = sw
                    persist(); st.rerun()
                except Exception as e:
                    st.error(f'エラー: {e}')
                finally:
                    for p in [sp, pp]:
                        try: os.unlink(p)
                        except: pass

    if st.session_state.get('second_winners'):
        sw = st.session_state.second_winners
        st.success(f'二次当選者: {len(sw)}名')
        st.dataframe(pd.DataFrame([{
            '氏名': w['name'], 'メール': w['email'],
            'ステータス': w['status'], '当選枠': w['slot'],
        } for w in sw]), use_container_width=True)

        st.warning('⚠️ メール送信は取り消しできません。本番時のみ実行してください。')
        mail_unlock3 = st.checkbox('🔓 メール送信のロックを解除する（本番のみ）', value=False, key='mail_unlock_3')
        sent = st.session_state.sent_modes
        if 'winner2' not in sent:
            if st.button('📨 二次当選メール送信', type='primary', disabled=not mail_unlock3):
                ok, ng, errs = send_bulk(sw, SUBJECT_WINNER_2ND, BODY_WINNER_2ND)
                if ng == 0:
                    st.success(f'✅ {ok}件送信完了')
                    st.session_state.sent_modes = sent + ['winner2']
                    persist(); st.rerun()
                else:
                    st.error(f'失敗 {ng}件: ' + ', '.join(errs))
        else:
            st.success('✅ 二次当選メール送信済み')

    c1, c2 = st.columns(2)
    with c1:
        if st.button('← Phase 2に戻る'):
            st.session_state.phase = 2; persist(); st.rerun()
    with c2:
        if st.button('→ Phase 4（直前案内）へ', type='primary'):
            st.session_state.phase = 4; persist(); st.rerun()

# ==============================
# PHASE 4: 直前案内・お礼
# ==============================

elif cur == 4:
    st.header('④ 直前案内・お礼メール')

    uf = st.session_state.url_fields
    with st.expander('URL設定', expanded=True):
        c1, c2 = st.columns(2)
        uf['presale_url']      = c1.text_input('先行販売URL',  value=uf.get('presale_url', ''))
        uf['presale_deadline'] = c2.text_input('先行販売締切', value=uf.get('presale_deadline', ''))
        uf['survey_url']       = st.text_input('アンケートURL', value=uf.get('survey_url', ''))
        st.session_state.url_fields = uf; persist()

    all_winners = st.session_state.winners + st.session_state.get('second_winners', [])
    st.metric('送信対象（当選者全員）', f'{len(all_winners)}名')
    sent = st.session_state.sent_modes

    st.warning('⚠️ メール送信は取り消しできません。本番時のみ実行してください。')
    mail_unlock4 = st.checkbox('🔓 メール送信のロックを解除する（本番のみ）', value=False, key='mail_unlock_4')

    st.subheader('直前案内メール（前日）')
    if 'reminder' not in sent:
        if st.button('📨 直前案内メール送信', type='primary', disabled=not mail_unlock4):
            if not uf.get('presale_url'):
                st.warning('先行販売URLを入力してください')
            else:
                ok, ng, errs = send_bulk(all_winners, SUBJECT_REMINDER, BODY_REMINDER, uf)
                if ng == 0:
                    st.success(f'✅ {ok}件送信完了')
                    st.session_state.sent_modes = sent + ['reminder']
                    persist(); st.rerun()
                else:
                    st.error(f'失敗 {ng}件: ' + ', '.join(errs))
    else:
        st.success('✅ 直前案内メール送信済み')

    st.subheader('お礼・アンケートメール（翌日）')
    sent = st.session_state.sent_modes
    if 'thanks' not in sent:
        if st.button('📨 お礼メール送信', disabled=not mail_unlock4):
            if not uf.get('survey_url'):
                st.warning('アンケートURLを入力してください')
            else:
                ok, ng, errs = send_bulk(all_winners, SUBJECT_THANKS, BODY_THANKS, uf)
                if ng == 0:
                    st.success(f'✅ {ok}件送信完了')
                    st.session_state.sent_modes = sent + ['thanks']
                    persist(); st.rerun()
                else:
                    st.error(f'失敗 {ng}件: ' + ', '.join(errs))
    else:
        st.success('✅ お礼メール送信済み')

    if st.button('← Phase 3に戻る'):
        st.session_state.phase = 3; persist(); st.rerun()

# ==============================
# Sidebar
# ==============================

with st.sidebar:
    st.markdown('### 管理')
    st.caption(f'現在: Phase {st.session_state.phase}')
    st.caption(f'一次当選: {len(st.session_state.winners)}名')
    st.caption(f'二次当選: {len(st.session_state.get("second_winners", []))}名')
    st.caption(f'送信済み: {", ".join(st.session_state.sent_modes) or "なし"}')
    st.divider()
    if st.button('🔄 新しいイベントでリセット', type='secondary'):
        if os.path.exists(STATE_FILE):
            os.remove(STATE_FILE)
        for k in list(st.session_state.keys()):
            del st.session_state[k]
        st.rerun()
