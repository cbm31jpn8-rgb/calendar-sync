#!/usr/bin/env python3
"""
GitHub Actions 버전 - 스마트 동기화

Google Sheets ↔ Google Calendar 비교 후 변경분만 처리
자동 추가된 이벤트에 [자동동기화] 태그를 붙여서 수동 이벤트와 구분

수정 내용:
1) 1.스탭 탭을 A:I가 아니라 A:AZ까지 읽음
2) 날짜가 1~2개만 있는 날짜행도 인식
3) 병합 셀/왼쪽 날짜 셀 구조 대응
4) 8a-6p, 6p-8a 근무 슬롯 안정적으로 파싱
5) 같은 날짜/같은 제목의 기존 캘린더 이벤트도 description/location 업데이트
6) 1.스탭 월별 근무 개수 보고를 텔레그램에 포함
"""

import os
import pickle
import base64
from collections import Counter
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from datetime import datetime, timedelta, timezone
import re
import requests
import urllib.parse

# ===== 한국 시간대 =====
KST = timezone(timedelta(hours=9))
DAY_NAMES = ['월', '화', '수', '목', '금', '토', '일']
SYNC_TAG = "[자동동기화]"

# ===== 설정 =====
SHARED_SHEET_ID = os.environ.get(
    "SHARED_SHEET_ID",
    "1KYTCcWQ_Ctfy72H7w-aVgJMdFGBWjS2HrOVcLgKjOOQ"
)
TARGET_NAME = os.environ.get("TARGET_NAME", "재희")
YEAR = int(os.environ.get("YEAR", "2026"))

# False: YEAR 전체 근무 개수
# True: 오늘 이후 남은 근무 개수만 계산
STAFF_MONTHLY_REPORT_FUTURE_ONLY = os.environ.get(
    "STAFF_MONTHLY_REPORT_FUTURE_ONLY",
    "false"
).lower() == "true"

# ===== 텔레그램 설정 =====
# 토큰은 코드에 직접 넣지 말고 GitHub Secrets에 넣으세요.
TELEGRAM_BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "")
TELEGRAM_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID", "")

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
# 날짜/시간 파싱
# ==========================================================

def parse_date_staff(date_str, year=2026):
    """
    1.스탭 날짜 파싱

    지원 형식:
    - 6월 1일
    - 6월1일
    - 2026.6.1
    - 2026-6-1
    - 2026/6/1
    - 6.1
    - 6/1
    """
    s = str(date_str or "").replace("\u00a0", " ").strip()

    # 2026.6.1 / 2026-6-1 / 2026/6/1
    match = re.search(r'(\d{4})[.\-/]\s*(\d{1,2})[.\-/]\s*(\d{1,2})', s)
    if match:
        try:
            return datetime(
                int(match.group(1)),
                int(match.group(2)),
                int(match.group(3))
            )
        except ValueError:
            return None

    # 6월 1일 / 6월1일
    match = re.search(r'(\d{1,2})\s*월\s*(\d{1,2})\s*일', s)
    if match:
        try:
            return datetime(year, int(match.group(1)), int(match.group(2)))
        except ValueError:
            return None

    # 6.1 / 6/1
    match = re.search(r'(?<!\d)(\d{1,2})[./]\s*(\d{1,2})(?!\d)', s)
    if match:
        try:
            return datetime(year, int(match.group(1)), int(match.group(2)))
        except ValueError:
            return None

    return None


def parse_staff_slot(slot_text):
    """
    1.스탭 근무 슬롯 파싱

    예:
    - 8a-6p   -> 08:00 시작, 10시간
    - 6p-8a   -> 18:00 시작, 14시간
    - 8A - 6P -> 08:00 시작, 10시간
    """
    raw = str(slot_text or "8a-6p").strip()
    normalized = raw.lower().replace(" ", "")

    match = re.search(r'(\d{1,2})([ap])-(\d{1,2})([ap])', normalized)
    if not match:
        return 8, 10, "8a-6p"

    start_h = int(match.group(1))
    start_ap = match.group(2)
    end_h = int(match.group(3))
    end_ap = match.group(4)

    def to_24h(hour, ap):
        if ap == 'a':
            return 0 if hour == 12 else hour
        return 12 if hour == 12 else hour + 12

    start_hour = to_24h(start_h, start_ap)
    end_hour = to_24h(end_h, end_ap)

    if end_hour <= start_hour:
        end_hour += 24

    duration_hours = end_hour - start_hour
    clean_slot = f"{start_h}{start_ap}-{end_h}{end_ap}"

    return start_hour, duration_hours, clean_slot


def parse_datetime_cpr(date_str):
    date_str = str(date_str or "").strip()
    date_match = re.match(r'(\d{4})\.?\s*(\d{1,2})\.?\s*(\d{1,2})', date_str)

    if not date_match:
        return None

    year = int(date_match.group(1))
    month = int(date_match.group(2))
    day = int(date_match.group(3))

    time_match = re.search(r'(\d{1,2}):(\d{2})', date_str)

    if time_match:
        hour = int(time_match.group(1))
        minute = int(time_match.group(2))
    else:
        hour = 0
        minute = 0

    try:
        return datetime(year, month, day, hour, minute)
    except ValueError:
        return None


def parse_datetime_academic(date_str):
    s = str(date_str or "").strip()
    match = re.match(r'(\d{4})\.\s*(\d{1,2})\.\s*(\d{1,2})', s)

    if match:
        try:
            return datetime(
                int(match.group(1)),
                int(match.group(2)),
                int(match.group(3))
            )
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
            spreadsheetId=sheet_id,
            range="'1.스탭'!A:AZ"
        ).execute()

        values = sheet.get('values', [])
        if not values:
            return schedules

        max_cols = max((len(row) for row in values), default=0)
        print(f"📊 1.스탭 탭: {len(values)}행 x 최대 {max_cols}열 읽음")

        # 날짜가 들어 있는 행 찾기
        date_rows = {}

        for row_idx, row in enumerate(values):
            row_dates = {}

            # C열부터 마지막 열까지 전체 검색
            for col_idx in range(2, len(row)):
                cell = str(row[col_idx] or "").strip()
                parsed = parse_date_staff(cell, YEAR)

                if parsed:
                    row_dates[col_idx] = parsed

            # 날짜가 1개만 있는 행도 날짜행으로 인정
            if row_dates:
                date_rows[row_idx] = row_dates

        # 근무 시간 슬롯이 들어 있는 행 찾기
        time_slots = {}

        for row_idx, row in enumerate(values):
            if len(row) > 1:
                b_cell = str(row[1] or "").strip()

                if re.search(
                    r'\d{1,2}\s*[ap]\s*-\s*\d{1,2}\s*[ap]',
                    b_cell,
                    re.IGNORECASE
                ):
                    time_slots[row_idx] = b_cell

        def find_date_for_cell(target_row_idx, target_col_idx):
            """
            target 셀 위쪽에서 가장 가까운 날짜행을 찾습니다.

            1순위: 같은 열에 있는 날짜
            2순위: 같은 날짜행에서 target_col_idx 왼쪽의 가장 가까운 날짜
            """
            for dr_idx in sorted(date_rows.keys(), reverse=True):
                if dr_idx >= target_row_idx:
                    continue

                row_dates = date_rows[dr_idx]

                if target_col_idx in row_dates:
                    return row_dates[target_col_idx]

                left_cols = [c for c in row_dates.keys() if c <= target_col_idx]
                if left_cols:
                    nearest_left_col = max(left_cols)
                    return row_dates[nearest_left_col]

            return None

        def find_slot_for_row(target_row_idx):
            """
            target 행 위쪽에서 가장 가까운 근무 슬롯을 찾습니다.
            """
            found_slot = "8a-6p"

            for ts_idx in sorted(time_slots.keys(), reverse=True):
                if ts_idx <= target_row_idx:
                    found_slot = time_slots[ts_idx]
                    break

            return found_slot

        # target_name 검색
        for row_idx, row in enumerate(values):
            for col_idx in range(2, len(row)):
                cell_str = str(row[col_idx] or "").strip()

                if target_name not in cell_str:
                    continue

                found_date = find_date_for_cell(row_idx, col_idx)

                if not found_date:
                    print(
                        f"⚠️  1.스탭 날짜 못 찾음: "
                        f"row={row_idx + 1}, col={col_idx + 1}, value={cell_str}"
                    )
                    continue

                raw_slot = find_slot_for_row(row_idx)
                start_hour, duration_hours, clean_slot = parse_staff_slot(raw_slot)

                event_dt = found_date.replace(hour=start_hour, minute=0)

                schedules.append({
                    'sheet': '1.스탭',
                    'name': f"근무 ({clean_slot})",
                    'datetime': event_dt,
                    'slot': clean_slot,
                    'duration_hours': duration_hours,
                    'location': '',
                    'details': (
                        f"근무: {clean_slot}, "
                        f"{found_date.strftime('%m월 %d일')}, "
                        f"함께: {cell_str}"
                    )
                })

        print(f"✅ 1.스탭에서 {target_name} 일정 {len(schedules)}개 발견")

        for s in schedules[:20]:
            print(
                f"  - {s['datetime'].strftime('%Y-%m-%d %H:%M')} "
                f"{s['name']} | {s['details']}"
            )

        if len(schedules) > 20:
            print(f"  ... 외 {len(schedules) - 20}개")

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
            spreadsheetId=sheet_id,
            range="'CPR교육일정'!A:E"
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
                    'sheet': 'CPR교육일정',
                    'name': f"CPR교육 - {region}",
                    'datetime': parsed_dt,
                    'location': venue,
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
            spreadsheetId=sheet_id,
            range="'3.학술'!A:H"
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
                        'sheet': '3.학술',
                        'name': f"{cell_str} (3.학술)",
                        'datetime': found_date,
                        'location': '',
                        'details': f"활동: {cell_str}, {found_date.strftime('%Y.%m.%d')}"
                    })

    except Exception as e:
        print(f"⚠️  3.학술 오류: {e}")

    return schedules

# ==========================================================
# 시트 전체 검색
# ==========================================================

def find_all_schedules(service, sheet_id, target_name):
    all_schedules = []

    print(f"\n🔍 '{target_name}' 검색 중...\n")

    all_schedules.extend(find_in_staff_sheet(service, sheet_id, target_name))
    all_schedules.extend(find_in_academic_sheet(service, sheet_id, target_name))
    all_schedules.extend(find_in_cpr_sheet(service, sheet_id, target_name))

    return all_schedules

# ==========================================================
# 1.스탭 월별 근무 개수 보고
# ==========================================================

def get_staff_monthly_work_report(schedules, future_only=False):
    """
    1.스탭 탭에서 찾은 TARGET_NAME의 월별 근무 개수 보고서 생성

    계산 기준:
    - sheet == '1.스탭'인 일정만 계산
    - 같은 날짜 + 같은 근무 슬롯은 1개로 계산
    - 같은 날짜에 8a-6p, 6p-8a처럼 슬롯이 다르면 각각 1개로 계산
    """
    now_kst = datetime.now(KST)
    today = now_kst.replace(
        hour=0,
        minute=0,
        second=0,
        microsecond=0,
        tzinfo=None
    )

    monthly_counts = Counter()
    seen = set()

    for schedule in schedules:
        if schedule.get('sheet') != '1.스탭':
            continue

        dt = schedule.get('datetime')
        if not dt:
            continue

        if future_only and dt < today:
            continue

        slot = schedule.get('slot', schedule.get('name', ''))
        unique_key = (
            dt.strftime('%Y-%m-%d'),
            slot
        )

        if unique_key in seen:
            continue

        seen.add(unique_key)
        monthly_counts[dt.month] += 1

    if future_only:
        title_scope = "오늘 이후 남은 근무"
    else:
        title_scope = f"{YEAR}년 전체 근무"

    report = f"👩‍⚕️ <b>{TARGET_NAME} 1.스탭 월별 근무 개수</b>\n"
    report += f"📌 기준: {title_scope}\n"
    report += "-" * 30 + "\n"

    total = 0

    for month in range(1, 13):
        count = monthly_counts.get(month, 0)
        total += count
        report += f"{month:02d}월: <b>{count}</b>개\n"

    report += f"\n합계: <b>{total}</b>개\n\n"

    return report

# ==========================================================
# 캘린더에서 자동동기화 이벤트 가져오기
# ==========================================================

def get_auto_synced_events(cal_service):
    """
    캘린더에서 [자동동기화] 태그가 있는 미래 이벤트 가져오기
    """
    try:
        now_kst = datetime.now(KST)
        today_start = now_kst.replace(
            hour=0,
            minute=0,
            second=0,
            microsecond=0
        )

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
                    dt = datetime.fromisoformat(
                        dt_str.replace('Z', '+00:00')
                    ).astimezone(KST)
                else:
                    dt = datetime.fromisoformat(dt_str)

                date_key = dt.strftime('%Y-%m-%d %H:%M')

            elif 'date' in start:
                date_key = start['date'] + " 00:00"

            else:
                continue

            key = (summary, date_key)

            auto_events[key] = {
                'event_id': event['id'],
                'summary': summary,
                'date_key': date_key,
                'location': event.get('location', ''),
                'description': desc
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
            'start': {
                'dateTime': dt.isoformat(),
                'timeZone': 'Asia/Seoul'
            },
            'end': {
                'dateTime': (dt + timedelta(hours=duration_hours)).isoformat(),
                'timeZone': 'Asia/Seoul'
            },
            'description': f'{SYNC_TAG}\n{details}'
        }

        if location:
            event['location'] = location

        return service.events().insert(
            calendarId='primary',
            body=event
        ).execute()

    except Exception as e:
        print(f"❌ 추가 오류: {e}")
        return None


def delete_event_from_calendar(service, event_id):
    try:
        service.events().delete(
            calendarId='primary',
            eventId=event_id
        ).execute()

        return True

    except Exception as e:
        print(f"❌ 삭제 오류: {e}")
        return False


def update_event_description_and_location(service, event_id, details, location=""):
    """
    캘린더에 이미 있는 자동동기화 이벤트의 description/location 업데이트.

    같은 날짜, 같은 제목이라도
    1.스탭의 '함께 근무자' 내용이 바뀌면 반영되도록 합니다.
    """
    try:
        event = service.events().get(
            calendarId='primary',
            eventId=event_id
        ).execute()

        changed = False
        details_changed = False
        location_changed = False

        new_description = f"{SYNC_TAG}\n{details}"
        old_description = event.get('description', '')

        if old_description != new_description:
            event['description'] = new_description
            changed = True
            details_changed = True

        old_location = event.get('location', '') or ""
        new_location = location or ""

        if old_location != new_location:
            if new_location:
                event['location'] = new_location
            else:
                event.pop('location', None)

            changed = True
            location_changed = True

        if changed:
            service.events().update(
                calendarId='primary',
                eventId=event_id,
                body=event
            ).execute()

        return details_changed, location_changed

    except Exception as e:
        print(f"❌ 내용/위치 업데이트 오류: {e}")
        return False, False

# ==========================================================
# 텔레그램
# ==========================================================

def send_telegram_message(message):
    try:
        if not TELEGRAM_BOT_TOKEN:
            print("❌ TELEGRAM_BOT_TOKEN이 설정되어 있지 않습니다.")
            return

        if not TELEGRAM_CHAT_ID:
            print("❌ TELEGRAM_CHAT_ID가 설정되어 있지 않습니다.")
            return

        url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"

        payload = {
            'chat_id': TELEGRAM_CHAT_ID,
            'text': message,
            'parse_mode': 'HTML',
            'disable_web_page_preview': 'true'
        }

        resp = requests.post(url, data=payload, timeout=20)

        if resp.status_code == 200:
            print("📱 텔레그램 발송 완료!")
        else:
            print(f"❌ 텔레그램 실패: {resp.status_code} {resp.text}")

    except Exception as e:
        print(f"❌ 텔레그램 오류: {e}")

# ==========================================================
# 캘린더에서 3일간 일정 조회
# ==========================================================

def get_upcoming_3days_report(cal_service):
    try:
        now_kst = datetime.now(KST)
        today_start = now_kst.replace(
            hour=0,
            minute=0,
            second=0,
            microsecond=0
        )

        range_end = today_start + timedelta(days=3, hours=23, minutes=59)

        events_result = cal_service.events().list(
            calendarId='primary',
            timeMin=today_start.isoformat(),
            timeMax=range_end.isoformat(),
            singleEvents=True,
            orderBy='startTime'
        ).execute()

        events = events_result.get('items', [])

        events_by_date = {}

        for event in events:
            start = event.get('start', {})

            if 'dateTime' in start:
                dt_str = start['dateTime']

                if '+' in dt_str[10:] or dt_str.endswith('Z'):
                    dt = datetime.fromisoformat(
                        dt_str.replace('Z', '+00:00')
                    ).astimezone(KST)
                else:
                    dt = datetime.fromisoformat(dt_str)

                time_str = dt.strftime('%H:%M')

            elif 'date' in start:
                dt = datetime.strptime(start['date'], '%Y-%m-%d')
                time_str = "종일"

            else:
                continue

            date_key = dt.strftime('%Y-%m-%d')

            if date_key not in events_by_date:
                events_by_date[date_key] = []

            events_by_date[date_key].append({
                'time_str': time_str,
                'summary': event.get('summary', '(제목 없음)'),
                'location': event.get('location', '')
            })

        report = "📋 <b>앞으로 3일간 일정 (캘린더)</b>\n"
        report += "-" * 30 + "\n"

        # 오늘부터 4일간 표시
        # 기존 코드 로직 유지: 오늘 + 3일 = 총 4일
        for day_offset in range(4):
            day_dt = today_start + timedelta(days=day_offset)
            date_key = day_dt.strftime('%Y-%m-%d')
            day_name = DAY_NAMES[day_dt.weekday()]

            report += f"\n<b>📆 {day_dt.strftime('%m/%d')}({day_name})</b>\n"

            if date_key in events_by_date:
                for ev in events_by_date[date_key]:
                    report += f"  ▪️ {ev['time_str']} {ev['summary']}\n"

                    if ev['location']:
                        encoded = urllib.parse.quote(ev['location'])
                        report += (
                            f"      📍 "
                            f"<a href=\"https://map.naver.com/v5/search/{encoded}\">"
                            f"{ev['location']}</a>\n"
                        )
            else:
                report += "  ▫️ 일정 없음\n"

        report += "\n"
        return report

    except Exception as e:
        print(f"⚠️ 캘린더 조회 오류: {e}")
        return "📋 캘린더 일정 조회 실패\n\n"

# ==========================================================
# 메인 - 스마트 동기화
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
    schedules = find_all_schedules(
        sheets_svc,
        SHARED_SHEET_ID,
        TARGET_NAME
    )

    # 1-1) 1.스탭 월별 근무 개수 보고서 생성
    # 전체 월별 개수를 보려면 미래 일정 필터링 전에 계산해야 합니다.
    staff_monthly_report = get_staff_monthly_work_report(
        schedules,
        future_only=STAFF_MONTHLY_REPORT_FUTURE_ONLY
    )

    print("\n" + staff_monthly_report.replace("<b>", "").replace("</b>", ""))

    now_kst = datetime.now(KST)
    today = now_kst.replace(
        hour=0,
        minute=0,
        second=0,
        microsecond=0,
        tzinfo=None
    )

    # 캘린더 동기화용 미래 일정만 추출
    unique = []
    seen = set()

    for s in schedules:
        key = (
            s['datetime'],
            s['name']
        )

        if key not in seen and s['datetime'] >= today:
            unique.append(s)
            seen.add(key)

    # 시트 일정 → (이름, 날짜+시간) 매핑
    sheet_events = {}

    for s in unique:
        date_key = s['datetime'].strftime('%Y-%m-%d %H:%M')
        key = (
            s['name'],
            date_key
        )
        sheet_events[key] = s

    print(f"📊 시트: {len(sheet_events)}개 미래 일정")

    # 2) 캘린더에서 [자동동기화] 이벤트 가져오기
    cal_events = get_auto_synced_events(cal_svc)

    # 3) 비교
    sheet_keys = set(sheet_events.keys())
    cal_keys = set(cal_events.keys())

    to_add = sheet_keys - cal_keys
    to_delete = cal_keys - sheet_keys
    unchanged = sheet_keys & cal_keys

    # 일정 변경 감지
    # 같은 이름이 삭제 + 추가에 동시에 있으면 날짜 이동 가능성이 있다고 보고 표시
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
                changes.append({
                    'name': name,
                    'old_date': old_d,
                    'new_date': new_d
                })

    print(
        f"\n📋 비교 결과: "
        f"추가 {len(to_add)}, "
        f"삭제 {len(to_delete)}, "
        f"유지 {len(unchanged)}, "
        f"변경 {len(changes)}"
    )

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
            dur = s.get('duration_hours', 10)

        elif s['sheet'] == 'CPR교육일정':
            dur = 2

        else:
            dur = 1

        print(f"  ➕ 추가: {s['datetime'].strftime('%Y-%m-%d')} | {name}")

        if add_event_to_calendar(
            cal_svc,
            name,
            s['datetime'],
            s['details'],
            dur,
            loc
        ):
            added_count += 1
            added_items.append(f"{s['datetime'].strftime('%Y-%m-%d')} {name}")

    # 6) 유지 이벤트 내용/위치 업데이트
    details_updated = 0
    loc_updated = 0

    for key in unchanged:
        s = sheet_events[key]
        cal_ev = cal_events[key]
        loc = s.get('location', '')

        details_changed, location_changed = update_event_description_and_location(
            cal_svc,
            cal_ev['event_id'],
            s['details'],
            loc
        )

        if details_changed:
            details_updated += 1

        if location_changed:
            loc_updated += 1

    # 7) 텔레그램 보고
    now_kst = datetime.now(KST)
    now_day = DAY_NAMES[now_kst.weekday()]

    report = "📅 <b>교육 일정 동기화 보고</b>\n"
    report += f"⏰ {now_kst.strftime('%Y-%m-%d')} ({now_day}) {now_kst.strftime('%H:%M')}\n\n"

    report += f"➕ 새로 추가: <b>{added_count}</b>개\n"
    report += f"✅ 변동 없음: <b>{len(unchanged)}</b>개\n"
    report += f"➖ 삭제됨: <b>{deleted_count}</b>개\n"

    if details_updated > 0:
        report += f"📝 내용 업데이트: <b>{details_updated}</b>개\n"

    if loc_updated > 0:
        report += f"📍 위치 업데이트: <b>{loc_updated}</b>개\n"

    if changes:
        report += f"\n🔄 일정 변경: <b>{len(changes)}</b>개\n\n"

        for c in changes[:5]:
            old_d = c['old_date'].split(' ')[0]
            new_d = c['new_date'].split(' ')[0]

            old_dt = datetime.strptime(old_d, '%Y-%m-%d')
            new_dt = datetime.strptime(new_d, '%Y-%m-%d')

            old_day = DAY_NAMES[old_dt.weekday()]
            new_day = DAY_NAMES[new_dt.weekday()]

            report += f"🔄 <b>{c['name']}</b>\n"
            report += f"   {old_d}({old_day}) → {new_d}({new_day})\n"

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

    # 월별 근무 개수 보고 추가
    report += "\n"
    report += staff_monthly_report

    # 앞으로 3일간 캘린더 일정 추가
    report += get_upcoming_3days_report(cal_svc)

    report += "✅ 동기화 완료!"

    send_telegram_message(report)

    print(
        f"\n✅ 완료: "
        f"추가 {added_count}, "
        f"삭제 {deleted_count}, "
        f"유지 {len(unchanged)}, "
        f"내용업데이트 {details_updated}, "
        f"위치업데이트 {loc_updated}, "
        f"변경 {len(changes)}"
    )


if __name__ == "__main__":
    main()
