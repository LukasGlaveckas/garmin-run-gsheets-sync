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
    print("🚀 Starting Lap-by-Lap Deep Dive (Last 20 Runs)...")
    
    # 1. Setup & Login
    garmin_email = os.environ.get('GARMIN_EMAIL')
    garmin_password = os.environ.get('GARMIN_PASSWORD')
    google_creds_json = os.environ.get('GOOGLE_CREDENTIALS')
    
    try:
        garmin = Garmin(garmin_email, garmin_password)
        garmin.login()
        
        creds_dict = json.loads(google_creds_json)
        creds = Credentials.from_service_account_info(creds_dict, scopes=['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
        client = gspread.authorize(creds)
        sheet = client.open("Garmin Data").sheet1
        
        # 2. CLEAR SHEET for fresh start
        print("🧹 Clearing old data to make room for lap details...")
        sheet.clear()
        
        # 3. Create Headers (Matches your screenshot style)
        headers = [
            "Date", "Activity Name", "Lap #", "Distance (km)", 
            "Time (min)", "Pace (min/km)", "Avg HR", "Max HR", 
            "Watts", "Cadence", "Stride (m)"
        ]
        sheet.append_row(headers)

        # 4. Fetch Last 20 Runs
        activities = garmin.get_activities(0, 20)
        
        print(f"📥 Found {len(activities)} activities. Extracting laps...")
        
        # 5. Loop through runs and extract LAPS
        rows_to_add = []
        for activity in activities:
            activity_type = activity.get('activityType', {}).get('typeKey', 'other')
            if activity_type not in ['running', 'treadmill_running', 'trail_running']:
                continue

            a_id = activity.get('activityId')
            date = activity.get('startTimeLocal', '')[:10]
            name = activity.get('activityName', 'Run')
            
            try:
                # Request specific splits for this run
                splits = garmin.get_activity_splits(a_id)
                laps = splits.get('lapSummaries', [])
                
                print(f"   🔹 Processing {date}: Found {len(laps)} laps")
                
                for lap in laps:
                    # Extract lap data
                    lap_index = lap.get('lapIndex', 0)
                    dist_km = round(lap.get('distance', 0) / 1000, 2)
                    duration_s = lap.get('duration', 0)
                    
                    # Calculate Pace
                    avg_speed = lap.get('averageSpeed', 0) # m/s
                    pace_str = "0:00"
                    if avg_speed > 0:
                        sec_km = 1000 / avg_speed
                        pace_str = format_pace(sec_km)
                    
                    # Metrics
                    avg_hr = lap.get('averageHR', 0)
                    max_hr = lap.get('maxHR', 0)
                    watts = round(lap.get('avgPower', 0)) if lap.get('avgPower') else 0
                    cadence = lap.get('averageRunningCadenceInStepsPerMinute', 0)
                    stride = round(lap.get('averageStrideLength', 0) / 100, 2) if lap.get('averageStrideLength') else 0

                    # Create Row for this single LAP
                    rows_to_add.append([
                        date, name, lap_index, dist_km, 
                        round(duration_s/60, 2), pace_str, avg_hr, max_hr, 
                        watts, cadence, stride
                    ])
                    
                # Small pause to be nice to Garmin API
                time.sleep(1)

            except Exception as e:
                print(f"   ⚠️ Error getting splits for {date}: {e}")
        
        # 6. Write all data at once (Faster)
        if rows_to_add:
            sheet.append_rows(rows_to_add)
            print(f"✅ Success! Added {len(rows_to_add)} lap rows to your sheet.")
        else:
            print("❌ No lap data found.")

    except Exception as e:
        print(f"❌ Fatal Error: {e}")

if __name__ == "__main__":
    main()
