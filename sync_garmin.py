import os
import json
from garminconnect import Garmin
from google.oauth2.service_account import Credentials
import gspread
import time

def format_pace(seconds_per_km):
    if not seconds_per_km: return "0:00"
    minutes = int(seconds_per_km // 60)
    seconds = int(seconds_per_km % 60)
    return f"{minutes}:{seconds:02d}"

def main():
    print("🚀 Restarting Lap Analysis (Safe-Save Mode)...")
    
    garmin_email = os.environ.get('GARMIN_EMAIL')
    garmin_password = os.environ.get('GARMIN_PASSWORD')
    google_creds_json = os.environ.get('GOOGLE_CREDENTIALS')
    
    try:
        # 1. Connect
        garmin = Garmin(garmin_email, garmin_password)
        garmin.login()
        
        creds_dict = json.loads(google_creds_json)
        creds = Credentials.from_service_account_info(creds_dict, scopes=['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
        client = gspread.authorize(creds)
        sheet = client.open("Garmin Data").sheet1
        
        # 2. Reset Sheet (Only once at the start)
        sheet.clear()
        headers = [
            "Date", "Activity Name", "Lap #", "Distance (km)", 
            "Time (min)", "Pace (min/km)", "Avg HR", "Max HR", 
            "Watts", "Cadence", "Stride (m)"
        ]
        sheet.append_row(headers)
        
        # 3. Get Activities (Let's start with 10 to be safe)
        activities = garmin.get_activities(0, 10)
        print(f"📥 Found {len(activities)} activities. Processing one by one...")
        
        for activity in activities:
            # Filter for running
            if 'running' not in activity.get('activityType', {}).get('typeKey', '').lower():
                continue

            a_id = activity.get('activityId')
            date = activity.get('startTimeLocal', '')[:10]
            name = activity.get('activityName', 'Run')
            
            print(f"   🔹 Syncing {date}...")
            
            rows_buffer = []
            try:
                # Get Splits
                splits = garmin.get_activity_splits(a_id)
                laps = splits.get('lapSummaries', [])
                
                if not laps:
                    print(f"      ⚠️ No laps found for {date}")
                    continue

                for lap in laps:
                    # Metrics
                    avg_speed = lap.get('averageSpeed', 0)
                    pace = format_pace(1000/avg_speed) if avg_speed > 0 else "0:00"
                    
                    rows_buffer.append([
                        date,
                        name,
                        lap.get('lapIndex', 0),
                        round(lap.get('distance', 0) / 1000, 2),
                        round(lap.get('duration', 0) / 60, 2),
                        pace,
                        lap.get('averageHR', 0),
                        lap.get('maxHR', 0),
                        round(lap.get('avgPower', 0)) if lap.get('avgPower') else 0,
                        lap.get('averageRunningCadenceInStepsPerMinute', 0),
                        round(lap.get('averageStrideLength', 0) / 100, 2) if lap.get('averageStrideLength') else 0
                    ])
                
                # SAVE IMMEDIATELY
                if rows_buffer:
                    sheet.append_rows(rows_buffer)
                    print(f"      ✅ Saved {len(rows_buffer)} laps.")
                
                # Pause to prevent timeout
                time.sleep(2)

            except Exception as e:
                print(f"      ❌ Failed to sync {date}: {e}")
                # Continue to next run even if this one fails!
                continue

    except Exception as e:
        print(f"❌ Fatal connection error: {e}")

if __name__ == "__main__":
    main()
