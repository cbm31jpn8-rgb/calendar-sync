#!/usr/bin/env python3
"""
GitHub Actions 버전
Google Sheets (교육 스케줄) → Google Calendar 자동 동기화 + 텔레그램 알림
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

DAY_NAMES = ['월', '화', '수', '목', '금', '토', '일']

# ==========================================================
# 인증 (GitHub Actions 버전)
# ==========================================================

def authenticate():
    """환경변수에서 토큰 로드 → 갱신 → 사용"""
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
            raise Exception(
                "토큰이 없거나 갱신할 수 없습니다.\n"
                "PC에서 다시 인증 후 token.pickle을 GitHub Secret에 업데이트해주세요."
            )

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
                        'datetime': event_dt, 'slot': found_slot,
                        'details': f"근무: {found_slot}, {found_date.strftime('%m월 %d일')}, 함께: {cell_str}"
                    })
                    print(f"  ✅ {found_date.strftime('%m월 %d일')} {found_slot} - {cell_str}")
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
                    'datetime': parsed_dt,
                    'location': venue,
                    'details': f"지역: {region}, 장소: {venue}, 주강사: {instructor}"
                })
                print(f"  ✅ {date_str} - {region} ({venue})")
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
                        'datetime': found_date,
                        'details': f"활동: {cell_str}, {found_date.strftime('%Y.%m.%d')}"
                    })
                    print(f"  ✅ {found_date.strftime('%Y.%m.%d')} - {cell_str}")
    except Exception as e:
        print(f"⚠️  3.학술 오류: {e}")
    return schedules

# ==========================================================
# 유틸리티
# ==========================================================

def find_all_schedules(service, sheet_id, target_name):
    all_s = []
    print(f"\n🔍 '{target_name}' 검색 중...\n")
    all_s.extend(find_in_staff_sheet(service, sheet_id, target_name))
    all_s.extend(find_in_academic_sheet(service, sheet_id, target_name))
    all_s.extend(find_in_cpr_sheet(service, sheet_id, target_name))
    return all_s


def event_exists(cal_service, event_name, event_datetime):
    try:
        day_start = event_datetime.replace(hour=0, minute=0, second=0)
        day_end = day_start + timedelta(days=1)
        result = cal_service.events().list(
            calendarId='primary',
            timeMin=day_start.isoformat() + '+09:00',
            timeMax=day_end.isoformat() + '+09:00',
            q=event_name, singleEvents=True
        ).execute()
        return any(e.get('summary', '') == event_name for e in result.get('items', []))
    except Exception as e:
        print(f"  ⚠️  중복 체크 오류: {e}")
        return False


def add_event_to_calendar(service, name, dt, details, duration_hours=1, location=""):
    try:
        event = {
            'summary': name,
            'start': {'dateTime': dt.isoformat(), 'timeZone': 'Asia/Seoul'},
            'end': {'dateTime': (dt + timedelta(hours=duration_hours)).isoformat(), 'timeZone': 'Asia/Seoul'},
            'description': f'시트에서 자동 추가됨\n{details}'
        }
        if location:
            event['location'] = location
        return service.events().insert(calendarId='primary', body=event).execute()
    except Exception as e:
        print(f"❌ 캘린더 추가 오류: {e}")
        return None


def send_telegram_message(message):
    try:
        url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
        payload = {
            'chat_id': TELEGRAM_CHAT_ID,
            'text': message,
            'parse_mode': 'HTML',
            'disable_web_page_preview': 'true'
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
    """Google Calendar에서 오늘(KST)~3일 후까지 모든 일정 가져오기"""
    try:
        now_kst = datetime.now(KST)
        today_start = now_kst.replace(hour=0, minute=0, second=0, microsecond=0)
        three_days = today_start + timedelta(days=3)

        time_min = today_start.isoformat()
        time_max = three_days.isoformat()

        print(f"📋 캘린더 조회: {time_min} ~ {time_max}")

        events_result = cal_service.events().list(
            calendarId='primary',
            timeMin=time_min,
            timeMax=time_max,
            singleEvents=True,
            orderBy='startTime'
        ).execute()

        events = events_result.get('items', [])
        print(f"📋 캘린더에서 {len(events)}개 일정 조회됨")

        if not events:
            return "📋 3일 내 예정된 일정 없음\n\n"

        report = "📋 <b>앞으로 3일간 일정 (캘린더)</b>\n"
        report += "-" * 30 + "\n"

        current_date = None
        for event in events:
            start = event.get('start', {})
            if 'dateTime' in start:
                dt_str = start['dateTime']
                # +09:00 또는 Z 형식 처리
                if '+' in dt_str[10:] or dt_str.endswith('Z'):
                    dt = datetime.fromisoformat(dt_str.replace('Z', '+00:00'))
                    dt = dt.astimezone(KST)
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
                naver_url = f"https://map.naver.com/v5/search/{encoded}"
                report += f"      📍 <a href=\"{naver_url}\">{location}</a>\n"

        report += "\n"
        return report

    except Exception as e:
        print(f"⚠️ 캘린더 조회 오류: {e}")
        import traceback
        traceback.print_exc()
        return f"📋 캘린더 일정 조회 실패: {str(e)}\n\n"

# ==========================================================
# 메인
# ==========================================================

def main():
    print("=" * 60)
    print("📅 교육 일정 동기화 (GitHub Actions)")
    print("=" * 60)

    # 인증
    try:
        creds = authenticate()
        sheets_svc = build('sheets', 'v4', credentials=creds)
        cal_svc = build('calendar', 'v3', credentials=creds)
        print("✅ 인증 완료")
    except Exception as e:
        print(f"❌ 인증 실패: {e}")
        send_telegram_message(f"❌ 인증 실패: {e}")
        return

    # 검색
    schedules = find_all_schedules(sheets_svc, SHARED_SHEET_ID, TARGET_NAME)
    if not schedules:
        msg = f"⚠️ '{TARGET_NAME}' 일정 없음"
        print(msg)
        send_telegram_message(msg)
        return

    # 중복 제거 + 정렬
    unique = []
    seen = set()
    for s in schedules:
        key = (s['datetime'], s['name'])
        if key not in seen:
            unique.append(s)
            seen.add(key)
    unique.sort(key=lambda x: x['datetime'])

    # 오늘(KST) 이전 제외
    now_kst = datetime.now(KST)
    today = now_kst.replace(hour=0, minute=0, second=0, microsecond=0, tzinfo=None)
    filtered = len(unique)
    unique = [s for s in unique if s['datetime'] >= today]
    if filtered - len(unique) > 0:
        print(f"⏩ 지난 일정 {filtered - len(unique)}개 건너뜀")

    if not unique:
        msg = "✅ 오늘 이후 새 일정 없음"
        print(msg)
        send_telegram_message(msg)
        return

    # 캘린더 추가
    print(f"\n📅 {len(unique)}개 처리 중...\n")
    added = skipped = 0

    for s in unique:
        name = s['name']
        if event_exists(cal_svc, name, s['datetime']):
            print(f"⏭️ 건너뜀: {s['datetime'].strftime('%Y-%m-%d')} | {name}")
            skipped += 1
            continue

        if s['sheet'] == '1.스탭':
            dur = 10 if '8a' in s.get('slot', '8a-6p').lower().split('-')[0] else 14
        elif s['sheet'] == 'CPR교육일정':
            dur = 2
        else:
            dur = 1

        if add_event_to_calendar(cal_svc, name, s['datetime'], s['details'], dur, s.get('location', '')):
            print(f"  ✅ 추가: {name}")
            added += 1
        else:
            print(f"  ❌ 실패: {name}")

    # ★ 텔레그램 보고 (KST 시간 사용)
    now_kst = datetime.now(KST)
    now_day = DAY_NAMES[now_kst.weekday()]
    report = f"📅 <b>교육 일정 동기화 보고</b>\n"
    report += f"⏰ {now_kst.strftime('%Y-%m-%d')} ({now_day}) {now_kst.strftime('%H:%M')}\n"
    report += f"📊 새로 추가: <b>{added}</b>개\n"
    report += f"⏭️ 중복 건너뜀: <b>{skipped}</b>개\n\n"

    # ★ 캘린더에서 앞으로 3일간 일정 가져오기
    report += get_upcoming_3days_report(cal_svc)

    report += "✅ 동기화 완료!"
    send_telegram_message(report)

    print(f"\n✅ 완료: 추가 {added}, 건너뜀 {skipped}")


if __name__ == "__main__":
    main()
