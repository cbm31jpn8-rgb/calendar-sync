#!/usr/bin/env python3
"""
Google Sheets (교육 스케줄) → Google Calendar 자동 동기화 + 텔레그램 알림
GitHub Actions 환경에서 동작하도록 작성됨.

필요한 환경변수 (GitHub Secrets):
- CREDENTIALS_JSON       : credentials.json 파일 내용 (전체 텍스트)
- TOKEN_PICKLE_BASE64    : token.pickle을 base64 인코딩한 문자열
- TELEGRAM_BOT_TOKEN     : 텔레그램 봇 토큰
- TELEGRAM_CHAT_ID       : 텔레그램 채팅 ID
"""

import os
import sys
import pickle
import base64
import json
import re
import requests
from datetime import datetime
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# ===== 설정 =====
SHARED_SHEET_ID = "1KYTCcWQ_Ctfy72H7w-aVgJMdFGBWjS2HrOVcLgKjOOQ"
TARGET_NAME = "재희"
LOG_FILE = "sync_log.json"

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets.readonly',
    'https://www.googleapis.com/auth/calendar'
]

# ===== 환경변수 로드 =====
TELEGRAM_BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "")
TELEGRAM_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID", "")
CREDENTIALS_JSON = os.environ.get("CREDENTIALS_JSON", "")
TOKEN_PICKLE_BASE64 = os.environ.get("TOKEN_PICKLE_BASE64", "")


# ===== 텔레그램 발송 (먼저 정의해서 에러 시에도 알림 가능) =====
def send_telegram_message(message):
    """텔레그램으로 메시지 발송"""
    if not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID:
        print("⚠️ 텔레그램 환경변수가 없어 알림을 건너뜁니다.")
        return False
    try:
        url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
        payload = {
            'chat_id': TELEGRAM_CHAT_ID,
            'text': message,
            'parse_mode': 'HTML'
        }
        response = requests.post(url, data=payload, timeout=15)
        if response.status_code == 200:
            print("📱 텔레그램 메시지 발송 완료")
            return True
        print(f"❌ 텔레그램 발송 실패: {response.text}")
        return False
    except Exception as e:
        print(f"❌ 텔레그램 발송 오류: {e}")
        return False


# ===== 인증 (GitHub Actions용) =====
def authenticate():
    """
    환경변수에서 token.pickle을 base64 디코딩해 로드.
    토큰이 만료된 경우 refresh_token으로 갱신만 시도하고,
    그것도 실패하면 텔레그램으로 알림 후 종료한다.
    (Actions에는 브라우저가 없으므로 새로 인증할 수 없음)
    """
    if not CREDENTIALS_JSON or not TOKEN_PICKLE_BASE64:
        msg = "❌ CREDENTIALS_JSON 또는 TOKEN_PICKLE_BASE64 Secret이 설정되지 않았습니다."
        print(msg)
        send_telegram_message(msg)
        sys.exit(1)

    # credentials.json을 임시 파일로 복원 (일부 라이브러리가 파일 경로를 요구할 때 대비)
    with open("credentials.json", "w", encoding="utf-8") as f:
        f.write(CREDENTIALS_JSON)

    # token.pickle 복원
    try:
        token_bytes = base64.b64decode(TOKEN_PICKLE_BASE64)
        creds = pickle.loads(token_bytes)
    except Exception as e:
        msg = f"❌ token.pickle 디코딩 실패: {e}\n\n로컬에서 token.pickle을 새로 만든 뒤 base64로 다시 등록해주세요."
        print(msg)
        send_telegram_message(msg)
        sys.exit(1)

    # 토큰 갱신
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
                print("✅ 토큰 자동 갱신 완료")
            except Exception as e:
                msg = (
                    f"❌ 토큰 갱신 실패: {e}\n\n"
                    f"로컬에서 token.pickle을 새로 생성한 뒤\n"
                    f"GitHub Secret <code>TOKEN_PICKLE_BASE64</code>를 업데이트해주세요."
                )
                print(msg)
                send_telegram_message(msg)
                sys.exit(1)
        else:
            msg = (
                "❌ refresh_token이 없거나 토큰이 무효합니다.\n\n"
                "로컬에서 token.pickle을 새로 만들어 Secret을 다시 등록해주세요."
            )
            print(msg)
            send_telegram_message(msg)
            sys.exit(1)

    return creds


# ===== 날짜/시간 파서 =====
def parse_datetime_staff(s):
    """'2026. 4. 1(수) 14:00' 형식"""
    s = s.strip()
    m = re.match(r'(\d{4})\.\s*(\d{1,2})\.\s*(\d{1,2})\([^)]+\)\s*(\d{1,2}):(\d{2})', s)
    if m:
        try:
            return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)),
                            int(m.group(4)), int(m.group(5)))
        except ValueError:
            return None
    return None


def parse_datetime_academic(s):
    """'2026. 3. 25 수' 형식"""
    s = s.strip()
    m = re.match(r'(\d{4})\.\s*(\d{1,2})\.\s*(\d{1,2})', s)
    if m:
        try:
            return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)), 0, 0)
        except ValueError:
            return None
    return None


# ===== 시트별 검색 =====
def find_in_staff_sheet(service, sheet_id, target_name):
    schedules = []
    try:
        sheet = service.spreadsheets().values().get(
            spreadsheetId=sheet_id, range="1.스탭!A:F"
        ).execute()
        values = sheet.get('values', [])
        print(f"📊 1.스탭 탭: {len(values)}행")

        for row_idx, row in enumerate(values):
            if row_idx < 4 or len(row) < 2:
                continue
            datetime_str = row[1].strip() if len(row) > 1 else ""
            location = row[2].strip() if len(row) > 2 else "교육"
            instructor = row[4].strip() if len(row) > 4 else ""

            if target_name in instructor and datetime_str:
                parsed = parse_datetime_staff(datetime_str)
                if parsed:
                    schedules.append({
                        'sheet': '1.스탭',
                        'name': f"{target_name} - {location}",
                        'datetime': parsed,
                        'details': f"주강사: {instructor}, 장소: {location}"
                    })
                    print(f"  ✅ {datetime_str}")
    except Exception as e:
        print(f"⚠️ 1.스탭 읽기 오류: {e}")
    return schedules


def find_in_academic_sheet(service, sheet_id, sheet_name, target_name):
    schedules = []
    try:
        sheet = service.spreadsheets().values().get(
            spreadsheetId=sheet_id, range=f"{sheet_name}!A:H"
        ).execute()
        values = sheet.get('values', [])
        print(f"📊 {sheet_name} 탭: {len(values)}행")

        for row_idx, row in enumerate(values):
            if row_idx < 2:
                continue
            row_text = " ".join(row).strip()
            if target_name not in row_text:
                continue
            for cell in row:
                cell_str = cell.strip() if cell else ""
                if not cell_str:
                    continue
                parsed = parse_datetime_academic(cell_str)
                if parsed:
                    schedules.append({
                        'sheet': sheet_name,
                        'name': f"{target_name} ({sheet_name})",
                        'datetime': parsed,
                        'details': row_text[:100]
                    })
                    print(f"  ✅ {cell_str}")
                    break
    except Exception as e:
        print(f"⚠️ {sheet_name} 읽기 오류: {e}")
    return schedules


def find_all_schedules(service, sheet_id, target_name):
    print(f"\n🔍 '{target_name}' 검색 중...\n")
    out = []
    out.extend(find_in_staff_sheet(service, sheet_id, target_name))
    out.extend(find_in_academic_sheet(service, sheet_id, "2.학술", target_name))
    out.extend(find_in_academic_sheet(service, sheet_id, "3.CPR교육", target_name))
    return out


# ===== 캘린더 =====
def event_already_exists(service, event_name, event_datetime):
    """동일 시간/제목 이벤트가 이미 있는지 확인"""
    try:
        time_min = event_datetime.isoformat() + "+09:00"
        end = event_datetime.replace(hour=(event_datetime.hour + 1) % 24)
        time_max = end.isoformat() + "+09:00"
        events = service.events().list(
            calendarId='primary',
            timeMin=time_min,
            timeMax=time_max,
            q=event_name,
            singleEvents=True
        ).execute()
        return len(events.get('items', [])) > 0
    except Exception:
        return False


def add_event_to_calendar(service, event_name, event_datetime, details):
    try:
        end = event_datetime.replace(hour=(event_datetime.hour + 1) % 24)
        event = {
            'summary': event_name,
            'start': {'dateTime': event_datetime.isoformat(), 'timeZone': 'Asia/Seoul'},
            'end': {'dateTime': end.isoformat(), 'timeZone': 'Asia/Seoul'},
            'description': f'시트에서 자동 추가됨\n{details}'
        }
        return service.events().insert(calendarId='primary', body=event).execute()
    except Exception as e:
        print(f"❌ 캘린더 추가 오류: {e}")
        return None


# ===== 로그 =====
def load_previous_log():
    if os.path.exists(LOG_FILE):
        try:
            with open(LOG_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            pass
    return {'added': []}


def save_log(schedules):
    data = {
        'added': [{'name': s['name'],
                   'datetime': s['datetime'].isoformat(),
                   'sheet': s['sheet']} for s in schedules],
        'last_sync': datetime.now().isoformat()
    }
    with open(LOG_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


# ===== 보고서 =====
def create_telegram_report(added_count, unique_schedules, previous_log):
    previous_keys = {(item['name'], item['datetime'])
                     for item in previous_log.get('added', [])}
    new_schedules = [s for s in unique_schedules
                     if (s['name'], s['datetime'].isoformat()) not in previous_keys]

    r = "📅 <b>교육 일정 동기화 보고</b>\n"
    r += "=" * 30 + "\n\n"
    r += f"⏰ 시각: {datetime.now().strftime('%Y-%m-%d %H:%M')}\n"
    r += f"👤 대상: {TARGET_NAME}\n"
    r += f"📊 캘린더 추가: <b>{added_count}</b>개\n\n"

    if new_schedules:
        r += "🆕 <b>새로 발견된 일정</b>\n"
        r += "-" * 30 + "\n"
        for s in new_schedules[:15]:
            d = s['datetime'].strftime('%Y.%m.%d %H:%M')
            r += f"📌 {d}\n   {s['name']} ({s['sheet']})\n\n"
        if len(new_schedules) > 15:
            r += f"... 외 {len(new_schedules) - 15}개\n"
    else:
        r += "ℹ️ 신규 일정 없음 (이전과 동일)\n"

    return r


# ===== 메인 =====
def main():
    print("=" * 50)
    print("📅 교육 일정 동기화 (GitHub Actions)")
    print("=" * 50)

    creds = authenticate()
    sheets_service = build('sheets', 'v4', credentials=creds)
    calendar_service = build('calendar', 'v3', credentials=creds)

    previous_log = load_previous_log()
    schedules = find_all_schedules(sheets_service, SHARED_SHEET_ID, TARGET_NAME)

    if not schedules:
        msg = f"⚠️ '{TARGET_NAME}'가 포함된 일정을 찾을 수 없습니다."
        print(msg)
        send_telegram_message(msg)
        return

    # 중복 제거
    unique_schedules = []
    seen = set()
    for s in schedules:
        key = (s['datetime'], s['name'])
        if key not in seen:
            unique_schedules.append(s)
            seen.add(key)
    unique_schedules.sort(key=lambda x: x['datetime'])

    print(f"\n📅 {len(unique_schedules)}개 일정 처리 중...\n")
    added_count = 0
    for s in unique_schedules:
        d = s['datetime'].strftime('%Y-%m-%d %H:%M')
        # 캘린더에 이미 있는지 확인
        if event_already_exists(calendar_service, s['name'], s['datetime']):
            print(f"⏭️  스킵 (이미 존재): {d} | {s['name']}")
            continue
        print(f"➕ 추가: {d} | {s['name']}")
        if add_event_to_calendar(calendar_service, s['name'], s['datetime'], s['details']):
            added_count += 1

    print("=" * 50)
    print(f"✅ 추가됨: {added_count}개 / 발견: {len(unique_schedules)}개")
    print("=" * 50)

    save_log(unique_schedules)
    send_telegram_message(create_telegram_report(added_count, unique_schedules, previous_log))


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        raise
    except Exception as e:
        import traceback
        tb = traceback.format_exc()
        print(tb)
        send_telegram_message(f"❌ 동기화 중 예외 발생:\n<pre>{str(e)[:500]}</pre>")
        sys.exit(1)
