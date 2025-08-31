#!/usr/bin/env python3
"""
Online Calendar Event Archiver for Tax Purposes
Works with online ICS URLs and specific month/year filtering
"""

import os
import sys
import requests
from datetime import datetime, date
from icalendar import Calendar
import hashlib
import json
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

def load_env_file():
    """Load environment variables from .env file"""
    env_vars = {}
    env_file = '.env'
    
    if os.path.exists(env_file):
        try:
            with open(env_file, 'r') as f:
                for line in f:
                    line = line.strip()
                    if line and not line.startswith('#') and '=' in line:
                        key, value = line.split('=', 1)
                        env_vars[key.strip()] = value.strip()
            print(f"✓ Loaded configuration from {env_file}")
        except Exception as e:
            print(f"Warning: Could not load {env_file}: {e}")
    
    return env_vars

def download_ics_file(url):
    """Download ICS file from URL with robust session handling and progressive timeouts"""
    print(f"Downloading ICS file from: {url}")
    
    # Create a session with retry strategy
    session = requests.Session()
    
    # Configure retry strategy
    retry_strategy = Retry(
        total=3,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["HEAD", "GET", "OPTIONS"],
        backoff_factor=1
    )
    
    adapter = HTTPAdapter(max_retries=retry_strategy)
    session.mount("http://", adapter)
    session.mount("https://", adapter)
    
    # Set headers that work well with Outlook/Office365
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
        'Accept': 'text/calendar,application/calendar,text/plain,*/*',
        'Accept-Language': 'en-US,en;q=0.9',
        'Accept-Encoding': 'gzip, deflate, br',
        'Connection': 'keep-alive',
        'Cache-Control': 'no-cache',
        'Pragma': 'no-cache'
    }
    
    # Try with increasing timeouts
    timeouts = [30, 60, 120]  # 30s, 60s, 120s
    
    for timeout in timeouts:
        try:
            print(f"Trying with {timeout} second timeout...")
            response = session.get(url, headers=headers, timeout=timeout, stream=True)
            response.raise_for_status()
            
            # Download content in chunks to handle large files better
            content = b''
            content_length = response.headers.get('content-length')
            if content_length:
                print(f"Expected file size: {int(content_length)} bytes")
            
            chunk_size = 8192  # 8KB chunks
            downloaded = 0
            
            for chunk in response.iter_content(chunk_size=chunk_size):
                if chunk:
                    content += chunk
                    downloaded += len(chunk)
                    if content_length and downloaded % (chunk_size * 100) == 0:  # Progress every ~800KB
                        progress = (downloaded / int(content_length)) * 100
                        print(f"  Downloaded {downloaded} bytes ({progress:.1f}%)")
            
            print(f"✓ Downloaded {len(content)} bytes successfully")
            return content
            
        except requests.exceptions.Timeout:
            print(f"✗ Timeout with {timeout} seconds")
            if timeout == timeouts[-1]:  # Last timeout attempt
                print(f"✗ Failed to download after {timeout} seconds")
        except requests.exceptions.ConnectionError as e:
            print(f"✗ Connection error: {e}")
            if timeout == timeouts[-1]:
                print("✗ All connection attempts failed")
        except requests.RequestException as e:
            print(f"✗ Error downloading ICS file: {e}")
            if timeout == timeouts[-1]:
                print("✗ All download attempts failed")
    
    # Final fallback with no timeout and simplified headers
    try:
        print("Final attempt with no timeout and basic headers...")
        simple_headers = {'User-Agent': 'Calendar-Archiver/1.0'}
        response = session.get(url, headers=simple_headers)
        response.raise_for_status()
        print(f"✓ Downloaded {len(response.content)} bytes")
        return response.content
    except requests.RequestException as e:
        print(f"✗ Final download attempt failed: {e}")
        sys.exit(1)
    finally:
        session.close()

def parse_ics_data(ics_data):
    """Parse ICS data and return list of events"""
    print("Parsing calendar events...")
    events = []
    
    try:
        cal = Calendar.from_ical(ics_data)
    except Exception as e:
        print(f"✗ Error parsing ICS data: {e}")
        sys.exit(1)
    
    event_count = 0
    for component in cal.walk():
        if component.name == "VEVENT":
            event_count += 1
            if event_count % 100 == 0:
                print(f"  Parsed {event_count} events so far...")
            
            try:
                # Extract event data
                title = str(component.get('summary', 'Untitled'))
                start_dt = component.get('dtstart').dt
                end_dt = component.get('dtend').dt if component.get('dtend') else start_dt
                uid = str(component.get('uid', f'event_{event_count}'))
                
                # Calculate duration
                if hasattr(start_dt, 'time'):
                    duration = end_dt - start_dt
                    duration_str = str(duration)
                else:
                    duration_str = "All day"
                
                location = str(component.get('location', ''))
                description = str(component.get('description', ''))
                
                # Get participants
                attendees = component.get('attendee', [])
                if not isinstance(attendees, list):
                    attendees = [attendees]
                participants = [str(att).replace('MAILTO:', '') for att in attendees]
                
                # Get timestamps
                created = str(component.get('created', ''))
                last_modified = str(component.get('last-modified', ''))
                
                event = {
                    'uid': uid,
                    'title': title,
                    'start': start_dt,
                    'end': end_dt,
                    'duration': duration_str,
                    'location': location,
                    'description': description,
                    'participants': participants,
                    'created': created,
                    'last_modified': last_modified
                }
                
                events.append(event)
                
            except Exception as e:
                print(f"  Warning: Skipping malformed event: {e}")
                continue
    
    print(f"✓ Parsed {len(events)} total events")
    return events

def filter_events_by_month(events, target_year, target_month):
    """Filter events to only include those from the specified month/year"""
    print(f"Filtering events for {target_year}-{target_month:02d}...")
    filtered_events = []
    
    for event in events:
        event_date = event['start']
        if hasattr(event_date, 'date'):
            event_date = event_date.date()
        
        if event_date.year == target_year and event_date.month == target_month:
            filtered_events.append(event)
    
    print(f"✓ Found {len(filtered_events)} events for {target_year}-{target_month:02d}")
    return filtered_events

def get_event_hash(event):
    """Create a hash of the event for comparison"""
    event_data = f"{event['title']}|{event['start']}|{event['end']}|{event['location']}|{event['description']}"
    return hashlib.md5(event_data.encode('utf-8')).hexdigest()

def save_daily_events(events, target_year, target_month):
    """Save events organized by year/month/day folders"""
    print(f"Saving events to ./{target_year}/{target_month:02d}/...")
    
    # Create metadata directory in the year folder
    metadata_dir = os.path.join(str(target_year), '.metadata')
    os.makedirs(metadata_dir, exist_ok=True)
    
    # Load existing event hashes for this month
    metadata_file = os.path.join(metadata_dir, f"{target_year}_{target_month:02d}_events.json")
    existing_hashes = {}
    if os.path.exists(metadata_file):
        try:
            with open(metadata_file, 'r') as f:
                existing_hashes = json.load(f)
            print(f"✓ Loaded {len(existing_hashes)} existing event hashes")
        except Exception as e:
            print(f"  Warning: Could not load existing metadata: {e}")
    
    updated_count = 0
    new_count = 0
    
    # Group events by date
    events_by_date = {}
    for event in events:
        event_date = event['start']
        if hasattr(event_date, 'date'):
            date_key = event_date.date()
        else:
            date_key = event_date
            
        if date_key not in events_by_date:
            events_by_date[date_key] = []
        events_by_date[date_key].append(event)
    
    # Process each day
    for event_date, date_events in events_by_date.items():
        year = event_date.year
        month = event_date.month
        day = event_date.day
        
        print(f"  Processing {year}-{month:02d}-{day:02d} ({len(date_events)} events)")
        
        # Create folder structure directly from current directory
        folder_path = os.path.join(str(year), f"{month:02d}", f"{day:02d}")
        os.makedirs(folder_path, exist_ok=True)
        
        # Save events to markdown file
        file_path = os.path.join(folder_path, f"{year}-{month:02d}-{day:02d}_events.md")
        
        # Read existing file content if it exists
        existing_events = {}
        if os.path.exists(file_path):
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                # Simple check if file exists (we'll regenerate it completely for simplicity)
                print(f"    Found existing file, will update...")
            except Exception as e:
                print(f"    Warning: Could not read existing file: {e}")
        
        # Sort events by start time (handle both datetime and date objects, timezone-aware and naive)
        def get_sort_key(event):
            start = event['start']
            if hasattr(start, 'time'):
                # datetime object - normalize timezone info for comparison
                if start.tzinfo is not None:
                    # Convert to UTC for consistent comparison
                    return start.utctimetuple()
                else:
                    # Naive datetime
                    return start.timetuple()
            else:
                # date object
                return datetime.combine(start, datetime.min.time()).timetuple()
        
        sorted_events = sorted(date_events, key=get_sort_key)
        
        # Generate markdown content
        markdown_content = f"# Calendar Events - {year}-{month:02d}-{day:02d}\n\n"
        markdown_content += f"**Date:** {event_date.strftime('%A, %B %d, %Y')}\n"
        markdown_content += f"**Total Events:** {len(sorted_events)}\n\n"
        markdown_content += "---\n\n"
        
        for i, event in enumerate(sorted_events, 1):
            event_hash = get_event_hash(event)
            
            # Check if this is an update
            if event['uid'] in existing_hashes:
                if existing_hashes.get(event['uid']) != event_hash:
                    print(f"    Updated: {event['title']}")
                    updated_count += 1
                else:
                    print(f"    Unchanged: {event['title']}")
            else:
                print(f"    New: {event['title']}")
                new_count += 1
            
            # Format start and end times
            start_time = event['start']
            end_time = event['end']
            
            if hasattr(start_time, 'strftime'):
                start_str = start_time.strftime('%H:%M')
                if hasattr(end_time, 'strftime'):
                    end_str = end_time.strftime('%H:%M')
                    time_str = f"{start_str} - {end_str}"
                else:
                    time_str = start_str
            else:
                time_str = "All day"
            
            markdown_content += f"## {i}. {event['title']}\n\n"
            markdown_content += f"**Time:** {time_str}  \n"
            markdown_content += f"**Duration:** {event['duration']}  \n"
            
            if event['location']:
                markdown_content += f"**Location:** {event['location']}  \n"
            
            if event['participants']:
                participants_clean = [p for p in event['participants'] if p.strip()]
                if participants_clean:
                    markdown_content += f"**Participants:** {', '.join(participants_clean)}  \n"
            
            if event['description']:
                # Clean up description
                desc = event['description'].replace('\n', ' ').replace('\r', ' ')
                # Limit description length but show more than before
                if len(desc) > 300:
                    desc = desc[:300] + "..."
                markdown_content += f"**Description:** {desc}  \n"
            
            markdown_content += f"**Event ID:** `{event['uid']}`\n\n"
            markdown_content += "---\n\n"
        
        # Write markdown file
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(markdown_content)
        
        print(f"    Saved to: {file_path}")
    
    # Save metadata with current event hashes
    try:
        with open(metadata_file, 'w') as f:
            json.dump({event['uid']: get_event_hash(event) for event in events}, f, indent=2)
        print(f"  Updated metadata file: {metadata_file}")
    except Exception as e:
        print(f"  Warning: Could not save metadata: {e}")
    
    print(f"✓ Saved events: {new_count} new, {updated_count} updated")

def archive_month(ics_url, target_year, target_month, all_events=None):
    """Archive events for a specific month"""
    print(f"\n=== Archiving {target_year}-{target_month:02d} ===")
    
    # Download and parse ICS if not already provided
    if all_events is None:
        ics_data = download_ics_file(ics_url)
        all_events = parse_ics_data(ics_data)
    
    # Filter for target month
    filtered_events = filter_events_by_month(all_events, target_year, target_month)
    
    if not filtered_events:
        print(f"No events found for {target_year}-{target_month:02d}")
        return 0
    
    # Save events directly to year/month structure
    save_daily_events(filtered_events, target_year, target_month)
    
    print(f"✓ Processed {len(filtered_events)} events for {target_year}-{target_month:02d}")
    return len(filtered_events)

def get_all_event_months(events):
    """Get all unique year/month combinations from events"""
    months = set()
    for event in events:
        event_date = event['start']
        if hasattr(event_date, 'date'):
            event_date = event_date.date()
        months.add((event_date.year, event_date.month))
    return sorted(months)

def main():
    # Load environment variables
    env_vars = load_env_file()
    
    print(f"=== Calendar Archiver Started ===")
    print(f"Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    # Parse command line arguments
    if len(sys.argv) == 1:
        # No arguments - use .env file and archive all months
        ics_url = env_vars.get('ICS_URL')
        if not ics_url:
            print("✗ No ICS_URL found in .env file")
            print("Please create a .env file with your ICS_URL or provide arguments")
            print("Usage: python calendar_archiver.py [ics_url] [year] [month]")
            print("Example: python calendar_archiver.py https://example.com/calendar.ics 2025 08")
            sys.exit(1)
        
        target_year = env_vars.get('DEFAULT_YEAR')
        if target_year:
            target_year = int(target_year)
        else:
            target_year = datetime.now().year
        
        target_month = env_vars.get('DEFAULT_MONTH')
        if target_month:
            target_month = int(target_month)
        else:
            target_month = None  # Archive all months
        
        print(f"URL: {ics_url}")
        if target_month:
            print(f"Target: {target_year}-{target_month:02d}")
        else:
            print(f"Target: All months in {target_year}")
        
        # Download and parse ICS once
        ics_data = download_ics_file(ics_url)
        all_events = parse_ics_data(ics_data)
        
        if target_month:
            # Archive specific month
            total_events = archive_month(ics_url, target_year, target_month, all_events)
        else:
            # Archive all months that have events
            available_months = get_all_event_months(all_events)
            year_months = [(year, month) for year, month in available_months if year == target_year]
            
            if not year_months:
                print(f"No events found for year {target_year}")
                sys.exit(0)
            
            print(f"Found events in {len(year_months)} months: {', '.join([f'{y}-{m:02d}' for y, m in year_months])}")
            
            total_events = 0
            for year, month in year_months:
                events_count = archive_month(ics_url, year, month, all_events)
                total_events += events_count
        
    elif len(sys.argv) == 3:
        # Two arguments: year and month (use .env for URL)
        ics_url = env_vars.get('ICS_URL')
        if not ics_url:
            print("✗ No ICS_URL found in .env file")
            print("Please create a .env file with your ICS_URL or provide the URL as first argument")
            sys.exit(1)
        
        target_year = int(sys.argv[1])
        target_month = int(sys.argv[2])
        
        print(f"URL: {ics_url}")
        print(f"Target: {target_year}-{target_month:02d}")
        
        total_events = archive_month(ics_url, target_year, target_month)
        
    elif len(sys.argv) == 4:
        # Three arguments: URL, year, month (original behavior)
        ics_url = sys.argv[1]
        target_year = int(sys.argv[2])
        target_month = int(sys.argv[3])
        
        print(f"URL: {ics_url}")
        print(f"Target: {target_year}-{target_month:02d}")
        
        total_events = archive_month(ics_url, target_year, target_month)
        
    else:
        print("Usage:")
        print("  python calendar_archiver.py                    # Archive all months using .env")
        print("  python calendar_archiver.py <year> <month>     # Archive specific month using .env URL")
        print("  python calendar_archiver.py <url> <year> <month>  # Archive specific month with URL")
        print("")
        print("Examples:")
        print("  python calendar_archiver.py                    # Archive all months from .env")
        print("  python calendar_archiver.py 2025 08            # Archive Aug 2025 using .env URL")
        print("  python calendar_archiver.py https://example.com/cal.ics 2025 08")
        sys.exit(1)
    
    print(f"\n=== Calendar Archiver Complete ===")
    if 'total_events' in locals():
        print(f"Total events processed: {total_events}")

if __name__ == "__main__":
    main()
