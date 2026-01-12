import os
import json
from garminconnect import Garmin
from google.oauth2.service_account import Credentials
import gspread
import time

# Load environment variables from .env file if it exists (for local testing)
if os.path.exists('.env'):
    try:
        from dotenv import load_dotenv
        load_dotenv()
    except ImportError:
        pass

def format_duration(seconds):
    """Convert seconds to minutes (rounded to 2 decimals)"""
    return round(seconds / 60, 2) if seconds else 0

def format_pace(distance_meters, duration_seconds):
    """Calculate pace in min/km (Format M:SS)"""
    if not distance_meters or not duration_seconds:
        return "0:00"
    distance_km = distance_meters / 1000
    pace_seconds = duration_seconds / distance_km
    minutes = int(pace_seconds // 60)
    seconds = int(pace_seconds % 60)
    return f"{minutes}:{seconds:02d}"

def main():
    print("🚀 Starting Bulk Sync (Last 100 Runs)...")
    
    # Get credentials
    garmin_email = os.environ.get('GARMIN_EMAIL')
    garmin_password = os.environ.get('GARMIN_PASSWORD')
    google_creds_json = os.environ.get('GOOGLE_CREDENTIALS')
    sheet_id = os.environ.get('SHEET_ID')
    
    if not all([garmin_email, garmin_password, google_creds_json, sheet_id]):
        print("❌ Missing required environment variables")
        return
    
    # Connect to Garmin
    try:
        garmin = Garmin(garmin_email, garmin_password)
        garmin.login()
        print("✅ Connected to Garmin")
    except Exception as e:
        print(f"❌ Failed to connect to Garmin: {e}")
        return
    
    # Connect to Google Sheets
    try:
        creds_dict = json.loads(google_creds_json)
        creds = Credentials.from_service_account_info(
            creds_dict,
            scopes=['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
        )
        client = gspread.authorize(creds)
        sheet = client.open("Garmin Data").sheet1
        print("✅ Connected to Google Sheets")
    except Exception as e:
        print(f"❌ Failed to connect to Google Sheets: {e}")
        return
    
    # Get activities (Last 100)
    print("📥 Fetching last 100 activities...")
    try:
        activities = garmin.get_activities(0, 100)
    except Exception as e:
        print(f"❌ Failed to fetch activities: {e}")
        return
    
    existing_data = sheet.get_all_values()
    existing_dates = {row[0] for row in existing_data if row}
    
    # Process each running activity
    new_entries = 0
    
    for activity in activities:
        # Filter for running activities only
        if activity.get('activityType', {}).get('typeKey', '').lower() not in ['running', 'treadmill_running', 'trail_running']:
            continue
            
        activity_date = activity.get('startTimeLocal', '')[:10]
        
        # Skip if already in sheet (Optional: Comment out to force update)
        if activity_date in existing_dates:
            # print(f"Skipping {activity_date} - already exists")
            continue
            
        print(f"   🔹 Processing {activity_date}...")

        # EXTRACT SPLITS (Km by Km)
        splits_string = "N/A"
        try:
            # We attempt to get the splits for detailed analysis
            activity_id = activity.get('activityId')
            splits_data = garmin.get_activity_splits(activity_id)
            lap_list = splits_data.get('lapSummaries', [])
            split_paces = []
            for lap in lap_list:
                # Only count laps that are substantial (e.g. > 500m) to avoid auto-pause noise
                if lap.get('distance', 0) > 400:
                    m_s = lap.get('averageSpeed', 0)
                    if m_s > 0:
                        p_sec = 1000 / m_s
                        split_paces.append(f"{int(p_sec//60)}:{int(p_sec%60):02d}")
            if split_paces:
                splits_string = " | ".join(split_paces)
        except:
            pass
            
        # EXTRACT METRICS (Including Watts)
        try:
            distance_meters = activity.get('distance', 0)
            duration_seconds = activity.get('duration', 0)
            
            row = [
                activity_date,
                activity.get('activityName', 'Run'),
                round(distance_meters / 1000, 2) if distance_meters else 0,
                format_duration(duration_seconds),
                format_pace(distance_meters, duration_seconds),
                activity.get('averageHR', 0) or 0,
                activity.get('maxHR', 0) or 0,
                activity.get('calories', 0) or 0,
                activity.get('averageRunningCadenceInStepsPerMinute', 0
