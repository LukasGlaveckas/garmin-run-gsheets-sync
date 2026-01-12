import os
import json
from garminconnect import Garmin
from google.oauth2.service_account import Credentials
import gspread
import time

def format_duration(seconds):
    return round(seconds / 60, 2) if seconds else 0

def format_pace(distance_meters, duration_seconds):
    if not distance_meters or not duration_seconds:
        return "0:00"
    distance_km = distance_meters / 1000
    pace_seconds = duration_seconds / distance_km
    minutes = int(pace_seconds // 60)
    seconds = int(pace_seconds % 60)
    return f"{minutes}:{seconds:02d}"

def main():
    print("🚀 Starting Bulk Sync (Last 100 Runs)...")
    
    garmin_email = os.environ.get('GARMIN_EMAIL')
    garmin_password = os.environ.get('GARMIN_PASSWORD')
    google_creds_json = os.environ.get('GOOGLE_CREDENTIALS')
    sheet_id = os.environ.get('SHEET_ID')
    
    if not all([garmin_email, garmin_password, google_creds_json, sheet_id]):
        print("❌ Missing environment variables")
        return

    # 1. Connect
    try:
        garmin = Garmin(garmin_email, garmin_password)
        garmin.login()
        print("✅ Connected to Garmin")
        
        creds_dict = json.loads(google_creds_json)
        creds = Credentials.from_service_account_info(creds_dict, scopes=['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
        client = gspread.authorize(creds)
        sheet = client.open("Garmin Data").sheet1
        print("✅ Connected to Google Sheets")
    except Exception as e:
        print(f"❌ Connection failed: {e}")
        return

    # 2. Get Activities (Increased to 100)
    print("📥 Fetching last 100 activities...")
    try:
        activities = garmin.get_activities(0, 100)
    except Exception as e:
        print(f"❌ Failed to fetch activities: {e}")
        return

    existing_data = sheet.get_all_values()
    existing_dates = {row[0] for row in existing_data if row}
    print(f"found {len(existing_dates)} existing entries.")

    # 3. Process Rows
    for activity in activities:
        # Filter for running
        if 'running' not in activity.get('activityType', {}).get('typeKey', '').lower():
            continue

        date = activity.get('startTimeLocal', '')[:10]
        if date in existing_dates:
            # Optional: Skip duplicates (Comment out if you want to overwrite)
            # print(f"Skipping {date} - already exists")
            # continue
            pass

        a_id = activity.get('activityId')
        print(f"   🔹 Processing {date}...")

        # FETCH SPLITS
        splits_string = "N/A"
        try:
            splits_data = garmin.get_activity_splits(a_id)
            lap_list = splits_data.get('lapSummaries', [])
            split_paces = []
            for lap in lap_list:
                m_s = lap.get('averageSpeed', 0)
                if m_s > 0 and lap.get('distance', 0) > 400: # Filter short/error laps
                    p_sec = 1000 / m_s
                    split_paces.append(f"{int(p_sec//60)}:{int(p_sec%60):02d}")
            if split_paces:
                splits_string = " | ".join(split_paces)
        except:
            pass

        # PREPARE ROW
        row = [
            date,
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
            round(activity.get('avgPower', 0), 0) if activity.get('avgPower') else 0,
            splits_string
        ]

        # Safe Save (Append one by one)
        try:
            sheet.append_row(row)
            # Add to existing dates so we don't add it twice if loop restarts
            existing_dates.add(date) 
        except Exception as e:
            print(f"❌ Error saving row for {date}: {e}")
            time.sleep(5) # Wait if Google complains

    print("✅ Sync complete!")

if __name__ == "__main__":
    main()
