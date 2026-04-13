#!/usr/bin/env python3
"""
GitHub Actions 버전 - 스마트 동기화
Google Sheets ↔ Google Calendar 비교 후 변경분만 처리
자동 추가된 이벤트에 [자동동기화] 태그를 붙여서 수동 이벤트와 구분
"""

import os
import pickle
import base64
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from datetime import datetime, timedelta, timezone
import re
import requests
import json
import urllib.parse

# ===== 한국 시간대 =====
KST = timezone(timedelta(hours=9))
DAY_NAMES = ['월', '화', '수', '목', '금', '토', '일']
SYNC_TAG = "[자동동기화]"

# ===== 설정 =====
SHARED_SHEET_ID = "1KYTCcWQ_Ctfy72H7w-aVgJMdFGBWjS2HrOVcLgKjOOQ"
TARGET_NAME = "재희"
YEAR = 2026

# ===== 텔레그램 설정 =====
TELEGRAM_BOT_TOKEN = "8731400386:AAEFXJypKgBoVL2AZyugMKSPk9k80_D3AJ8"
TELEGRAM_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID", "31443525")

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets.readonly',
    'https://www.googleapis.com/auth/calendar'
]

# ==========================================================
# 인증
# ==========================================================

def authenticate():
    creds_json = os.environ.get("CREDENTIALS_JSON", "")
    if creds_json:
        with open("credentials.json", "w") as f:
            f.write(creds_json)

    token_b64 = os.environ.get("TOKEN_PICKLE_BASE64", "")
    if token_b64:
        token_bytes = base64.b64decode(token_b64)
        with open("token.pickle", "wb") as f:
            f.write(token_bytes)

    creds = None
    if os.path.exists("token.pickle"):
        with open("token.pickle", "rb") as token:
            creds = pickle.load(token)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
            print("✅ 토큰 자동 갱신 완료")
        else:
            raise Exception("토큰 갱신 불가. PC에서 다시 인증 후 token.pickle을 GitHub Secret에 업데이트해주세요.")

    return creds

# ==========================================================
# 날짜 파싱
# ==========================================================

def parse_date_staff(date_str, year=2026):
    match = re.match(r'(\d{1,2})월\s*(\d{1,2})일', date_str.strip())
    if match:
        try:
            return datetime(year, int(match.group(1)), int(match.group(2)))
        except ValueError:
            return None
    return None

def parse_datetime_cpr(date_str):
    date_str = date_str.strip()
    date_match = re.match(r'(\d{4})\.?\s*(\d{1,2})\.?\s*(\d{1,2})', date_str)
    if not date_match:
        return None
    year, month, day = int(date_match.group(1)), int(date_match.group(2)), int(date_match.group(3))
    time_match = re.search(r'(\d{1,2}):(\d{2})', date_str)
    hour, minute = (int(time_match.group(1)), int(time_match.group(2))) if time_match else (0, 0)
    try:
        return datetime(year, month, day, hour, minute)
    except ValueError:
        return None

def parse_datetime_academic(date_str):
    match = re.match(r'(\d{4})\.\s*(\d{1,2})\.\s*(\d{1,2})', date_str.strip())
    if match:
        try:
            return datetime(int(match.group(1)), int(match.group(2)), int(match.group(3)))
        except ValueError:
            return None
    return None

# ==========================================================
# 1.스탭 탭
# ==========================================================

def find_in_staff_sheet(service, sheet_id, target_name):
    schedules = []
    try:
        sheet = service.spreadsheets().values().get(
            spreadsheetId=sheet_id, range="'1.스탭'!A:I"
        ).execute()
        values = sheet.get('values', [])
        if not values:
            return schedules
        print(f"📊 1.스탭 탭: {len(values)}행 읽음")

        date_rows = {}
        for row_idx, row in enumerate(values):
            row_dates = {}
            for col_idx in range(2, min(len(row), 9)):
                cell = row[col_idx].strip() if col_idx < len(row) and row[col_idx] else ""
                parsed = parse_date_staff(cell, YEAR)
                if parsed:
                    row_dates[col_idx] = parsed
            if len(row_dates) >= 3:
                date_rows[row_idx] = row_dates

        time_slots = {}
        for row_idx, row in enumerate(values):
            if len(row) > 1:
                b_cell = row[1].strip() if row[1] else ""
                if re.match(r'\d+[ap]-\d+[ap]', b_cell, re.IGNORECASE):
                    time_slots[row_idx] = b_cell

        for row_idx, row in enumerate(values):
            for col_idx in range(2, min(len(row), 9)):
                cell_str = row[col_idx].strip() if col_idx < len(row) and row[col_idx] else ""
                if target_name not in cell_str:
                    continue
                found_date = None
                for dr_idx in sorted(date_rows.keys(), reverse=True):
                    if dr_idx < row_idx and col_idx in date_rows[dr_idx]:
                        found_date = date_rows[dr_idx][col_idx]
                        break
                found_slot = "8a-6p"
                for ts_idx in sorted(time_slots.keys(), reverse=True):
                    if ts_idx <= row_idx:
                        found_slot = time_slots[ts_idx]
                        break
                if found_date:
                    start_hour = 8 if '8a' in found_slot.lower().split('-')[0] else 18
                    event_dt = found_date.replace(hour=start_hour, minute=0)
                    schedules.append({
                        'sheet': '1.스탭', 'name': f"근무 ({found_slot})",
                        'datetime': event_dt, 'slot': found_slot, 'location': '',
                        'details': f"근무: {found_slot}, {found_date.strftime('%m월 %d일')}, 함께: {cell_str}"
                    })
    except Exception as e:
        print(f"⚠️  1.스탭 오류: {e}")
    return schedules

# ==========================================================
# CPR교육일정 탭
# ==========================================================

def find_in_cpr_sheet(service, sheet_id, target_name):
    schedules = []
    try:
        sheet = service.spreadsheets().values().get(
            spreadsheetId=sheet_id, range="'CPR교육일정'!A:E"
        ).execute()
        values = sheet.get('values', [])
        if not values:
            return schedules
        print(f"📊 CPR교육일정 탭: {len(values)}행 읽음")

        for row_idx, row in enumerate(values):
            if row_idx < 3:
                continue
            instructor = row[4].strip() if len(row) > 4 and row[4] else ""
            if target_name not in instructor:
                continue
            date_str = row[1].strip() if len(row) > 1 and row[1] else ""
            region = row[0].strip() if len(row) > 0 and row[0] else ""
            venue = row[2].strip() if len(row) > 2 and row[2] else ""
            parsed_dt = parse_datetime_cpr(date_str)
            if parsed_dt:
                schedules.append({
                    'sheet': 'CPR교육일정', 'name': f"CPR교육 - {region}",
                    'datetime': parsed_dt, 'location': venue,
                    'details': f"지역: {region}, 장소: {venue}, 주강사: {instructor}"
                })
    except Exception as e:
        print(f"⚠️  CPR교육일정 오류: {e}")
    return schedules

# ==========================================================
# 3.학술 탭
# ==========================================================

def find_in_academic_sheet(service, sheet_id, target_name):
    schedules = []
    try:
        sheet = service.spreadsheets().values().get(
            spreadsheetId=sheet_id, range="'3.학술'!A:H"
        ).execute()
        values = sheet.get('values', [])
        if not values:
            return schedules
        print(f"📊 3.학술 탭: {len(values)}행 읽음")

        date_rows = {}
        for row_idx, row in enumerate(values):
            row_dates = {}
            for col_idx, cell in enumerate(row):
                parsed = parse_datetime_academic(cell.strip() if cell else "")
                if parsed:
                    row_dates[col_idx] = parsed
            if row_dates:
                date_rows[row_idx] = row_dates

        for row_idx, row in enumerate(values):
            for col_idx, cell in enumerate(row):
                cell_str = cell.strip() if cell else ""
                if target_name not in cell_str:
                    continue
                found_date = None
                for dr_idx in sorted(date_rows.keys(), reverse=True):
                    if dr_idx < row_idx and col_idx in date_rows[dr_idx]:
                        found_date = date_rows[dr_idx][col_idx]
                        break
                if found_date:
                    schedules.append({
                        'sheet': '3.학술', 'name': f"{cell_str} (3.학술)",
                        'datetime': found_date, 'location': '',
                        'details': f"활동: {cell_str}, {found_date.strftime('%Y.%m.%d')}"
                    })
    except Exception as e:
        print(f"⚠️  3.학술 오류: {e}")
    return schedules

# ==========================================================
# 시트 전체 검색
# ==========================================================

def find_all_schedules(service, sheet_id, target_name):
    all_s = []
    print(f"\n🔍 '{target_name}' 검색 중...\n")
    all_s.extend(find_in_staff_sheet(service, sheet_id, target_name))
    all_s.extend(find_in_academic_sheet(service, sheet_id, target_name))
    all_s.extend(find_in_cpr_sheet(service, sheet_id, target_name))
    return all_s

# ==========================================================
# ★ 캘린더에서 자동동기화 이벤트 가져오기
# ==========================================================

def get_auto_synced_events(cal_service):
    """캘린더에서 [자동동기화] 태그가 있는 미래 이벤트 가져오기"""
    try:
        now_kst = datetime.now(KST)
        today_start = now_kst.replace(hour=0, minute=0, second=0, microsecond=0)

        events_result = cal_service.events().list(
            calendarId='primary',
            timeMin=today_start.isoformat(),
            maxResults=500,
            singleEvents=True,
            orderBy='startTime'
        ).execute()

        auto_events = {}
        for event in events_result.get('items', []):
            desc = event.get('description', '')
            if SYNC_TAG not in desc:
                continue

            summary = event.get('summary', '')
            start = event.get('start', {})

            if 'dateTime' in start:
                dt_str = start['dateTime']
                if '+' in dt_str[10:] or dt_str.endswith('Z'):
                    dt = datetime.fromisoformat(dt_str.replace('Z', '+00:00')).astimezone(KST)
                else:
                    dt = datetime.fromisoformat(dt_str)
                date_key = dt.strftime('%Y-%m-%d')
            elif 'date' in start:
                date_key = start['date']
            else:
                continue

            key = (summary, date_key)
            auto_events[key] = {
                'event_id': event['id'],
                'summary': summary,
                'date_key': date_key,
                'location': event.get('location', '')
            }

        print(f"📅 캘린더에서 자동동기화 이벤트 {len(auto_events)}개 조회됨")
        return auto_events

    except Exception as e:
        print(f"⚠️ 캘린더 조회 오류: {e}")
        return {}

# ==========================================================
# 캘린더 추가/삭제/업데이트
# ==========================================================

def add_event_to_calendar(service, name, dt, details, duration_hours=1, location=""):
    try:
        event = {
            'summary': name,
            'start': {'dateTime': dt.isoformat(), 'timeZone': 'Asia/Seoul'},
            'end': {'dateTime': (dt + timedelta(hours=duration_hours)).isoformat(), 'timeZone': 'Asia/Seoul'},
            'description': f'{SYNC_TAG}\n{details}'
        }
        if location:
            event['location'] = location
        return service.events().insert(calendarId='primary', body=event).execute()
    except Exception as e:
        print(f"❌ 추가 오류: {e}")
        return None

def delete_event_from_calendar(service, event_id):
    try:
        service.events().delete(calendarId='primary', eventId=event_id).execute()
        return True
    except Exception as e:
        print(f"❌ 삭제 오류: {e}")
        return False

def update_event_location(service, event_id, location):
    try:
        event = service.events().get(calendarId='primary', eventId=event_id).execute()
        event['location'] = location
        if SYNC_TAG not in event.get('description', ''):
            event['description'] = f"{SYNC_TAG}\n{event.get('description', '')}"
        service.events().update(calendarId='primary', eventId=event_id, body=event).execute()
        return True
    except Exception as e:
        print(f"❌ 위치 업데이트 오류: {e}")
        return False

# ==========================================================
# 텔레그램
# ==========================================================

def send_telegram_message(message):
    try:
        url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
        payload = {
            'chat_id': TELEGRAM_CHAT_ID, 'text': message,
            'parse_mode': 'HTML', 'disable_web_page_preview': 'true'
        }
        resp = requests.post(url, data=payload)
        if resp.status_code == 200:
            print("📱 텔레그램 발송 완료!")
        else:
            print(f"❌ 텔레그램 실패: {resp.status_code} {resp.text}")
    except Exception as e:
        print(f"❌ 텔레그램 오류: {e}")

# ==========================================================
# ★ 캘린더에서 3일간 일정 조회
# ==========================================================

def get_upcoming_3days_report(cal_service):
    try:
        now_kst = datetime.now(KST)
        today_start = now_kst.replace(hour=0, minute=0, second=0, microsecond=0)
        range_end = today_start + timedelta(days=3, hours=23, minutes=59)

        events_result = cal_service.events().list(
            calendarId='primary',
            timeMin=today_start.isoformat(),
            timeMax=range_end.isoformat(),
            singleEvents=True, orderBy='startTime'
        ).execute()

        events = events_result.get('items', [])
        if not events:
            return "📋 3일 내 예정된 일정 없음\n\n"

        report = "📋 <b>앞으로 3일간 일정 (캘린더)</b>\n"
        report += "-" * 30 + "\n"

        current_date = None
        for event in events:
            start = event.get('start', {})
            if 'dateTime' in start:
                dt_str = start['dateTime']
                if '+' in dt_str[10:] or dt_str.endswith('Z'):
                    dt = datetime.fromisoformat(dt_str.replace('Z', '+00:00')).astimezone(KST)
                else:
                    dt = datetime.fromisoformat(dt_str)
                time_str = dt.strftime('%H:%M')
            elif 'date' in start:
                dt = datetime.strptime(start['date'], '%Y-%m-%d')
                time_str = "종일"
            else:
                continue

            date_key = dt.strftime('%Y-%m-%d')
            if date_key != current_date:
                current_date = date_key
                day_name = DAY_NAMES[dt.weekday()]
                report += f"\n<b>📆 {dt.strftime('%m/%d')}({day_name})</b>\n"

            summary = event.get('summary', '(제목 없음)')
            location = event.get('location', '')
            report += f"  ▪️ {time_str} {summary}\n"
            if location:
                encoded = urllib.parse.quote(location)
                report += f"      📍 <a href=\"https://map.naver.com/v5/search/{encoded}\">{location}</a>\n"

        report += "\n"
        return report
    except Exception as e:
        print(f"⚠️ 캘린더 조회 오류: {e}")
        return "📋 캘린더 일정 조회 실패\n\n"

# ==========================================================
# ★ 메인 - 스마트 동기화
# ==========================================================

def main():
    print("=" * 60)
    print("📅 교육 일정 스마트 동기화")
    print("=" * 60)

    try:
        creds = authenticate()
        sheets_svc = build('sheets', 'v4', credentials=creds)
        cal_svc = build('calendar', 'v3', credentials=creds)
        print("✅ 인증 완료")
    except Exception as e:
        print(f"❌ 인증 실패: {e}")
        send_telegram_message(f"❌ 인증 실패: {e}")
        return

    # 1) 시트에서 일정 가져오기
    schedules = find_all_schedules(sheets_svc, SHARED_SHEET_ID, TARGET_NAME)

    now_kst = datetime.now(KST)
    today = now_kst.replace(hour=0, minute=0, second=0, microsecond=0, tzinfo=None)

    unique = []
    seen = set()
    for s in schedules:
        key = (s['datetime'], s['name'])
        if key not in seen and s['datetime'] >= today:
            unique.append(s)
            seen.add(key)

    # 시트 일정 → (이름, 날짜) 매핑
    sheet_events = {}
    for s in unique:
        date_key = s['datetime'].strftime('%Y-%m-%d')
        key = (s['name'], date_key)
        sheet_events[key] = s

    print(f"📊 시트: {len(sheet_events)}개 미래 일정")

    # 2) 캘린더에서 [자동동기화] 이벤트 가져오기
    cal_events = get_auto_synced_events(cal_svc)

    # 3) ★ 비교
    sheet_keys = set(sheet_events.keys())
    cal_keys = set(cal_events.keys())

    to_add = sheet_keys - cal_keys
    to_delete = cal_keys - sheet_keys
    unchanged = sheet_keys & cal_keys

    # 일정 변경 감지 (같은 이름이 삭제+추가에 동시 존재 = 날짜 이동)
    changes = []
    add_names = {}
    for name, date in to_add:
        add_names.setdefault(name, []).append(date)
    del_names = {}
    for name, date in to_delete:
        del_names.setdefault(name, []).append(date)

    changed_event_names = set(add_names.keys()) & set(del_names.keys())
    for name in changed_event_names:
        for old_d in del_names[name]:
            for new_d in add_names[name]:
                changes.append({'name': name, 'old_date': old_d, 'new_date': new_d})

    print(f"\n📋 비교 결과: 추가 {len(to_add)}, 삭제 {len(to_delete)}, 유지 {len(unchanged)}, 변경 {len(changes)}")

    # 4) 삭제 실행
    deleted_count = 0
    deleted_items = []
    for key in to_delete:
        event_info = cal_events[key]
        name, date_key = key
        print(f"  ➖ 삭제: {date_key} | {name}")
        if delete_event_from_calendar(cal_svc, event_info['event_id']):
            deleted_count += 1
            deleted_items.append(f"{date_key} {name}")

    # 5) 추가 실행
    added_count = 0
    added_items = []
    for key in to_add:
        s = sheet_events[key]
        name = s['name']
        loc = s.get('location', '')

        if s['sheet'] == '1.스탭':
            dur = 10 if '8a' in s.get('slot', '8a-6p').lower().split('-')[0] else 14
        elif s['sheet'] == 'CPR교육일정':
            dur = 2
        else:
            dur = 1

        print(f"  ➕ 추가: {s['datetime'].strftime('%Y-%m-%d')} | {name}")
        if add_event_to_calendar(cal_svc, name, s['datetime'], s['details'], dur, loc):
            added_count += 1
            added_items.append(f"{s['datetime'].strftime('%Y-%m-%d')} {name}")

    # 6) 유지 이벤트 위치 업데이트
    loc_updated = 0
    for key in unchanged:
        s = sheet_events[key]
        cal_ev = cal_events[key]
        loc = s.get('location', '')
        if loc and not cal_ev.get('location'):
            if update_event_location(cal_svc, cal_ev['event_id'], loc):
                loc_updated += 1

    # 7) ★ 텔레그램 보고
    now_kst = datetime.now(KST)
    now_day = DAY_NAMES[now_kst.weekday()]

    report = f"📅 <b>교육 일정 동기화 보고</b>\n"
    report += f"⏰ {now_kst.strftime('%Y-%m-%d')} ({now_day}) {now_kst.strftime('%H:%M')}\n\n"

    report += f"➕ 새로 추가: <b>{added_count}</b>개\n"
    report += f"✅ 변동 없음: <b>{len(unchanged)}</b>개\n"
    report += f"➖ 삭제됨: <b>{deleted_count}</b>개\n"

    if changes:
        report += f"🔄 일정 변경: <b>{len(changes)}</b>개\n\n"
        for c in changes[:5]:
            old_dt = datetime.strptime(c['old_date'], '%Y-%m-%d')
            new_dt = datetime.strptime(c['new_date'], '%Y-%m-%d')
            old_day = DAY_NAMES[old_dt.weekday()]
            new_day = DAY_NAMES[new_dt.weekday()]
            report += f"🔄 <b>{c['name']}</b>\n"
            report += f"   {c['old_date']}({old_day}) → {c['new_date']}({new_day})\n"
        if len(changes) > 5:
            report += f"   ... 외 {len(changes) - 5}건\n"
    elif deleted_count > 0 or added_count > 0:
        report += "\n"
        if deleted_items:
            report += "➖ <b>삭제된 일정</b>\n"
            for item in deleted_items[:5]:
                report += f"   ❌ {item}\n"
            if len(deleted_items) > 5:
                report += f"   ... 외 {len(deleted_items) - 5}개\n"
        if added_items:
            report += "➕ <b>추가된 일정</b>\n"
            for item in added_items[:5]:
                report += f"   🆕 {item}\n"
            if len(added_items) > 5:
                report += f"   ... 외 {len(added_items) - 5}개\n"

    if loc_updated > 0:
        report += f"\n📍 위치 업데이트: {loc_updated}개\n"

    report += "\n"
    report += get_upcoming_3days_report(cal_svc)
    report += "✅ 동기화 완료!"
    send_telegram_message(report)

    print(f"\n✅ 완료: 추가 {added_count}, 삭제 {deleted_count}, 유지 {len(unchanged)}, 변경 {len(changes)}")


if __name__ == "__main__":
    main()
