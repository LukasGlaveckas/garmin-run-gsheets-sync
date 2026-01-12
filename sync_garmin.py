import os
import json
from garminconnect import Garmin
from google.oauth2.service_account import Credentials
import gspread
from datetime import datetime

def format_duration(seconds):
    return round(seconds / 60, 2) if seconds else 0

def format_pace(distance_meters, duration_seconds):
    if not distance_meters or not duration_seconds:
        return 0
    distance_km = distance_meters / 1000
    pace_seconds = duration_seconds / distance_km
    minutes = int(pace_seconds // 60)
    seconds = int(pace_seconds % 60)
    return f"{minutes}:{seconds:02d}"

def main():
    print("Starting Deep Sync (Splits, Watts, TE)...")
    garmin_email = os.environ.get('GARMIN_EMAIL')
    garmin_password = os.environ.get('GARMIN_PASSWORD')
    google_creds_json = os.environ.get('GOOGLE_CREDENTIALS')
    sheet_id = os.environ.get('SHEET_ID')

    if not all([garmin_email, garmin_password, google_creds_json, sheet_id]):
        print("❌ Missing environment variables")
        return

    # 1. Connect to Garmin
    try:
        garmin = Garmin(garmin_email, garmin_password)
        garmin.login()
        print("✅ Connected to Garmin")
    except Exception as e:
        print(f"❌ Garmin Login Failed: {e}")
        return

    # 2. Connect to Google Sheets
    try:
        creds_dict = json.loads(google_creds_json)
        creds = Credentials.from_service_account_info(creds_dict, scopes=['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
        client = gspread.authorize(creds)
        sheet = client.open("Garmin Data").sheet1
        print("✅ Connected to Google Sheets")
    except Exception as e:
        print(f"❌ Google Sheets Connection Failed: {e}")
        return

    # 3. Get Activities
    activities = garmin.get_activities(0, 10) # Pulling last 10 to be fast
    running_activities = [a for a in activities if a.get('activityType', {}).get('typeKey', '').lower() in ['running', 'treadmill_running', 'trail_running']]
    
    existing_data = sheet.get_all_values()
    existing_dates = {row[0] for row in existing_data if row}

    for activity in running_activities:
        activity_date = activity.get('startTimeLocal', '')[:10]
        if activity_date in existing_dates:
            print(f"Skipping {activity_date} - already exists")
            continue

        activity_id = activity.get('activityId')
        print(f"Processing {activity_date} (ID: {activity_id})...")

        # FETCH SPLITS
        try:
            splits_data = garmin.get_activity_splits(activity_id)
            lap_list = splits_data.get('lapSummaries', [])
            split_paces = []
            for lap in lap_list:
                m_s = lap.get('averageSpeed', 0)
                if m_s > 0:
                    p_sec = 1000 / m_s
                    split_paces.append(f"{int(p_sec//60)}:{int(p_sec%60):02d}")
            splits_string = " | ".join(split_paces)
        except:
            splits_string = "N/A"

        # ASSEMBLE ROW
        row = [
            activity_date,
            activity.get('activityName', 'Run'),
            round(activity.get('distance', 0) / 1000, 2),
            format_duration(activity.get('duration', 0)),
            format_pace(activity.get('distance', 0), activity.get('duration', 0)),
            activity.get('averageHR', 0) or 0,
            activity.get('maxHR', 0) or 0,
            activity.get('calories', 0) or 0,
            activity.get('averageRunningCadenceInStepsPerMinute', 0) or 0,
            round(activity.get('elevationGain', 0), 1) if activity.get('elevationGain') else 0,
            activity.get('aerobicTrainingEffect', 0),
            activity.get('anaerobicTrainingEffect', 0),
            activity.get('vO2MaxValue', 0),
            round(activity.get('averageStrideLength', 0), 1) if activity.get('averageStrideLength') else 0,
            activity.get('avgPower', 0) or 0,
            splits_string
        ]

        sheet.append_row(row)
        print(f"✅ Added {activity_date}")

if __name__ == "__main__":
    main()
