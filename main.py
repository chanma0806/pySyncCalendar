from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import googleapiclient.discovery
from datetime import datetime, timedelta
import win32com.client
import pickle
import os
from os.path import join, dirname
from dotenv import load_dotenv

load_dotenv(verbose=True)

dotenv_path = join(dirname(__file__), '.env')
load_dotenv(dotenv_path)

# 必要なスコープ
SCOPES = ['https://www.googleapis.com/auth/calendar']

CALENDAR_ID = 'primary' # googleアカウントのメインカレンダーID
KEY_FILE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), os.getenv('KEY_FILE_PATH'))
DATE_FORMAT = '%Y-%m-%dT%H:%M:%S'
OUTLOOK_DARE_FORMAT = "%m/%d/%Y %I:%M %p"
CRED_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "token.pickle")

SYNC_TARGET_DAYS = 7

# --util
def get_past_cred():
    if os.path.exists(CRED_PATH):
        with open(CRED_PATH, 'rb') as token:
            return pickle.load(token)
    else:
        return None

def append_prefix(outlook_event_subject):
    if "work" in outlook_event_subject or "備忘録" in outlook_event_subject:
        return outlook_event_subject
    else:
        return "[会議]"+outlook_event_subject

def get_google_api_cred():
    creds = get_past_cred()
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                KEY_FILE_PATH, SCOPES)
            creds = flow.run_local_server(port=0)
            # 新たに取得した認証情報を保存する
            with open(CRED_PATH, 'wb') as token:
                pickle.dump(creds, token)
    return creds

# --usecase
def get_outlook_calendar_events():
    # winPcのローカルにあるOutlookのカレンダーを指定期間分取得
    outlook = win32com.client.Dispatch("Outlook.Application")
    mapi = outlook.GetNamespace("MAPI")
    calendar = mapi.GetDefaultFolder(9)  # 9はカレンダーを指す
    start = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    end = start + timedelta(days=SYNC_TARGET_DAYS)
    outlook_events = calendar.Items
    outlook_events.IncludeRecurrences = True
    outlook_events.Sort("[Start]")
    outlook_events = outlook_events.Restrict("[Start] >= '" + start.strftime(OUTLOOK_DARE_FORMAT) + "' AND [END] <= '" + end.strftime(OUTLOOK_DARE_FORMAT) + "'")
    return outlook_events

def convert_event_outloook_to_google(events):
    # Outlookのイベント達をGoogleのイベントに変換
    ## descriptionにはoutlookeventのIDを入れる
    google_events = []
    for event in events:
        if "キャンセル済" in event.Subject or "Canceled" in event.Subject:
            # スキップ
            continue
        google_event = {
            'summary': append_prefix(event.Subject),
            'description': event.EntryID,
            'start': {
                'dateTime': event.Start.strftime(DATE_FORMAT),
                'timeZone': 'Asia/Tokyo',
            },
            'end': {
                'dateTime': event.End.strftime(DATE_FORMAT),
                'timeZone': 'Asia/Tokyo',
            },
            "reminders": {
                "useDefault": False,
                "overrides": [
                    {"method": "popup", "minutes": 5}
                ]
            }
        }
        google_events.append(google_event)

    return google_events
def write_to_google_calendar(target_events):

    creds = get_google_api_cred()
    service = googleapiclient.discovery.build('calendar', 'v3', credentials=creds)

    # 期間
    start_time = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    end_time = start_time + timedelta(hours=SYNC_TARGET_DAYS*24)  # 1時間後

    # 既存のイベントを検索
    events_result = service.events().list(
        calendarId=CALENDAR_ID,
        timeMin=start_time.isoformat()+"+09:00",
        timeMax=end_time.isoformat()+"+09:00"
    ).execute()
    existing_events = events_result.get('items', [])

    # 既存のイベントがある場合は時刻を更新、ない場合は新規作成
    insert_events = []
    updated_events = []
    for te in target_events:
        is_hit_same_description = False
        for e in existing_events:
            if te["description"] == e["description"]:
                te["id"] = e["id"]
                updated_events.append(te)
                existing_events.remove(e)
                is_hit_same_description = True
                break
        if not is_hit_same_description:
            insert_events.append(te)
        
    # イベントを更新
    for e in updated_events:
        service.events() \
            .update(
                calendarId=CALENDAR_ID,
                eventId=e["id"],
                body=e
            ) \
            .execute()
    # イベントの新規登録
    for e in insert_events:
        service.events() \
            .insert(
                calendarId=CALENDAR_ID,
                body=e
            ) \
            .execute()

if __name__ == "__main__":
    try:
        # creds = get_google_api_cred()
        outlook_events = get_outlook_calendar_events()
        google_events = convert_event_outloook_to_google(outlook_events)
        write_to_google_calendar(google_events)
    except Exception as e:
        print("Error occurred.")
        print(e)

