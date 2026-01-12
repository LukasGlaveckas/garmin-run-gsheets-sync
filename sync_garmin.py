import os
import json
from garminconnect import Garmin
from google.oauth2.service_account import Credentials
import gspread

def main():
    print("📊 Analyst Mode: Extracting Deep Performance Metrics...")
    garmin_email = os.environ.get('GARMIN_EMAIL')
    garmin_password = os.environ.get('GARMIN_PASSWORD')
    google_creds_json = os.environ.get('GOOGLE_CREDENTIALS')
    sheet_id = os.environ.get('SHEET_ID')

    try:
        # 1. Initialize Connections
        garmin = Garmin(garmin_email, garmin_password)
        garmin.login()
        creds_dict = json.loads(google_creds_json)
        creds = Credentials.from_service_account_info(creds_dict, scopes=['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
        client = gspread.authorize(creds)
        sheet = client.open("Garmin Data").sheet1
        
        # 2. Get the VERY LAST activity (Focusing on one at a time for deep analysis)
        activities = garmin.get_activities(0, 1)
        if not activities:
            print("No activities found.")
            return

        last_run = activities[0]
        a_id = last_run.get('activityId')
        date = last_run.get('startTimeLocal', '')[:10]
        
        # Check if this specific run is already in the sheet
        existing_dates = {row[0] for row in sheet.get_all_values() if row}
        if date in existing_dates:
            print(f"✅ Training for {date} already analyzed. No new data to sync.")
            return

        print(f"🏃‍♂️ New training found: {date}. Running Deep Analysis...")

        # 3. Coach's Split Extraction (km by km)
        split_list = []
        try:
            # We use the specific splits endpoint for high accuracy
            splits_data = garmin.get_activity_splits(a_id)
            laps = splits_data.get('lapSummaries', [])
            for lap in laps:
                # Filter out 'messy' short laps (e.g., stopping the watch late)
                if lap.get('distance', 0) > 400:
                    speed = lap.get('averageSpeed', 0) # m/s
                    if speed > 0:
                        pace_raw = 1000 / speed
                        split_list.append(f"{int(pace_raw//60)}:{int(pace_raw%60):02d}")
        except Exception as e:
            print(f"Warning: Splits unavailable via standard API: {e}")

        # 4. Assemble Coach's Row
        # These are the metrics needed to track your 3:09:59 Kaunas progress
        row = [
            date,                                     # A: Date
            last_run.get('activityName', 'Run'),      # B: Name
            round(last_run.get('distance', 0)/1000, 2),# C: Distance (km)
            round(last_run.get('duration', 0)/60, 2),  # D: Time (min)
            last_run.get('averageHR', 0),             # E: Avg HR
            last_run.get('maxHR', 0),                 # F: Max HR
            round(last_run.get('avgPower', 0), 0) if last_run.get('avgPower') else 0, # G: Watts
            last_run.get('vO2MaxValue', 0),           # H: VO2 Max
            last_run.get('aerobicTrainingEffect', 0), # I: Aerobic TE
            last_run.get('anaerobicTrainingEffect', 0),# J: Anaerobic TE
            round(last_run.get('averageStrideLength', 0), 1) if last_run.get('averageStrideLength') else 0, # K: Stride (cm)
            last_run.get('averageRunningCadenceInStepsPerMinute', 0), # L: Cadence
            " | ".join(split_list) if split_list else "Manual Lap Required" # M: Splits
        ]

        sheet.append_row(row)
        print(f"✅ Deep Analysis Complete for {date}. Check your sheet!")

    except Exception as e:
        print(f"❌ Coach's Error: {e}")

if __name__ == "__main__":
    main()
