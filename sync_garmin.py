import os
import json
from garminconnect import Garmin
from google.oauth2.service_account import Credentials
import gspread

def main():
    print("🚀 Analyst Mode: Deep Data Extraction Starting...")
    garmin_email = os.environ.get('GARMIN_EMAIL')
    garmin_password = os.environ.get('GARMIN_PASSWORD')
    google_creds_json = os.environ.get('GOOGLE_CREDENTIALS')
    sheet_id = os.environ.get('SHEET_ID')

    try:
        # Connect to Services
        garmin = Garmin(garmin_email, garmin_password)
        garmin.login()
        creds_dict = json.loads(google_creds_json)
        creds = Credentials.from_service_account_info(creds_dict, scopes=['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
        client = gspread.authorize(creds)
        sheet = client.open("Garmin Data").sheet1
        
        # Get last 5 activities to ensure deep processing doesn't timeout
        activities = garmin.get_activities(0, 5)
        existing_dates = {row[0] for row in sheet.get_all_values() if row}

        for activity in activities:
            a_type = activity.get('activityType', {}).get('typeKey', '').lower()
            if 'running' not in a_type: continue
            
            date = activity.get('startTimeLocal', '')[:10]
            if date in existing_dates: continue

            a_id = activity.get('activityId')
            print(f"📊 Analyzing Run: {date}...")

            # DEEP METRIC EXTRACTION
            # Get Splits via the 'Splits' API specifically
            split_paces = []
            try:
                splits_data = garmin.get_activity_splits(a_id)
                for s in splits_data.get('lapSummaries', []):
                    # Filter out small 'leftover' laps less than 500m
                    if s.get('distance', 0) > 500:
                        ms = s.get('averageSpeed', 0)
                        if ms > 0:
                            p = 1000/ms
                            split_paces.append(f"{int(p//60)}:{int(p%60):02d}")
            except: pass
            
            # COACH'S DATA POINTS
            row = [
                date,
                activity.get('activityName'),
                round(activity.get('distance', 0)/1000, 2),
                round(activity.get('duration', 0)/60, 2),
                activity.get('averageHR', 0),
                activity.get('maxHR', 0),
                round(activity.get('avgPower', 0), 0) if activity.get('avgPower') else 0,
                activity.get('vO2MaxValue', 0),
                activity.get('aerobicTrainingEffect', 0),
                activity.get('anaerobicTrainingEffect', 0),
                round(activity.get('averageStrideLength', 0), 1) if activity.get('averageStrideLength') else 0,
                activity.get('averageRunningCadenceInStepsPerMinute', 0),
                activity.get('trainingLoad', 0), # How hard this specific run was
                activity.get('recoveryTime', 0), # Hours needed to recover
                " | ".join(split_paces) if split_paces else "Check Garmin Settings"
            ]
            sheet.append_row(row)
            print(f"✅ Run Analyzed: {date}")

    except Exception as e:
        print(f"❌ Analysis Error: {e}")

if __name__ == "__main__":
    main()
