# ==========================================
# Microsoft 365 Room Booking PoC Backend
# ==========================================
import os
import httpx
from fastapi import FastAPI, HTTPException, Header
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from datetime import datetime
from typing import List, Optional
from dotenv import load_dotenv
from urllib.parse import quote
import asyncio
from fastapi import BackgroundTasks
# --- 1. Setup & Configuration ---
load_dotenv()

app = FastAPI(title="Vinci Energies Room Booking API", version="2.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"], 
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Azure AD Credentials
TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"

# --- 2. Data Models ---
class AvailabilityRequest(BaseModel):
    room_email: str
    start_time: datetime
    end_time: datetime
    time_zone: str = "UTC"

class BookingRequest(BaseModel):
    subject: str
    room_email: str
    start_time: datetime
    end_time: datetime
    organizer_email: str # Used for reference only
    attendees: List[str] = []
    description: str = ""
    filiale: str = ""

# --- 3. Auth Helper (System Token) ---
async def get_app_token():
    """Gets the System's 'Master Key' to check room status."""
    if not all([TENANT_ID, CLIENT_ID, CLIENT_SECRET]):
        raise HTTPException(status_code=500, detail="Missing Azure AD credentials.")

    token_url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = {
        "client_id": CLIENT_ID,
        "scope": "https://graph.microsoft.com/.default",
        "client_secret": CLIENT_SECRET,
        "grant_type": "client_credentials",
    }

    async with httpx.AsyncClient() as client:
        response = await client.post(token_url, data=data)
        
    if response.status_code != 200:
        raise HTTPException(status_code=401, detail=f"Azure Auth Failed: {response.text}")
        
    return response.json().get("access_token")

# --- 4. API Endpoints ---

@app.get("/rooms")
async def get_rooms():
    static_rooms = [
        {
            "displayName": "Conference Room A",
            "emailAddress": "ConferenceRoomA@AxiansPoc611.onmicrosoft.com",
            "floor": "Floor 3",
            "department": "Axians"
        },
        {
            "displayName": "Conference Room C",
            "emailAddress": "ConferenceRoomC@AxiansPoc611.onmicrosoft.com",
            "floor": "Floor 4",
            "department": "QHSE"
        }
    ]
    return {"value": static_rooms}

@app.post("/availability")
async def check_availability(req: AvailabilityRequest):
    token = await get_app_token()
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
        "Prefer": f'outlook.timezone="{req.time_zone}"'
    }
    
    payload = {
        "schedules": [req.room_email],
        "startTime": {"dateTime": req.start_time.isoformat(), "timeZone": req.time_zone},
        "endTime": {"dateTime": req.end_time.isoformat(), "timeZone": req.time_zone},
        "availabilityViewInterval": 15
    }

    async with httpx.AsyncClient() as client:
        url = f"{GRAPH_BASE_URL}/users/{req.room_email}/calendar/getSchedule"
        resp = await client.post(url, headers=headers, json=payload)
    return resp.json()

# ... (Previous imports and setup) ...

# ==========================================
# 👻 THE GHOST BUSTER (Background Task)
# ==========================================
async def remove_ghost_meetings():
    """
    Runs continuously. 
    1. Checks all rooms.
    2. Finds meetings started > 5 mins ago.
    3. If NOT 'Checked-In', DELETE them.
    """
    while True:
        try:
            token = await get_app_token()
            headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
            
            # Define the "Ghost Window" (Started 5 to 20 mins ago)
            now = datetime.utcnow()
            five_mins_ago = (now - timedelta(minutes=5)).isoformat() + "Z"
            twenty_mins_ago = (now - timedelta(minutes=20)).isoformat() + "Z"

            rooms = await get_rooms() # We reuse your existing get_rooms function
            
            for room in rooms['value']:
                email = room['emailAddress']
                
                # Find active meetings in the ghost window
                url = (
                    f"{GRAPH_BASE_URL}/users/{email}/calendarView"
                    f"?startDateTime={twenty_mins_ago}"
                    f"&endDateTime={five_mins_ago}"
                    f"&$select=id,subject,categories"
                )
                
                async with httpx.AsyncClient() as client:
                    resp = await client.get(url, headers=headers)
                    if resp.status_code == 200:
                        events = resp.json().get('value', [])
                        for event in events:
                            # 🔍 CHECK: Does it have the 'Checked-In' tag?
                            categories = event.get('categories', [])
                            if "Checked-In" not in categories:
                                print(f"👻 GHOST BUSTED: Deleting {event['subject']} in {email}")
                                # DELETE THE MEETING
                                await client.delete(
                                    f"{GRAPH_BASE_URL}/users/{email}/events/{event['id']}", 
                                    headers=headers
                                )
                                
        except Exception as e:
            print(f"Ghost Buster Error: {e}")
            
        # Sleep for 60 seconds before checking again
        await asyncio.sleep(60)

# Start the Ghost Buster when App starts
@app.on_event("startup")
async def startup_event():
    # Run in background
    asyncio.create_task(remove_ghost_meetings())


# ==========================================
# ✅ NEW ENDPOINT: CHECK-IN
# ==========================================
class CheckInRequest(BaseModel):
    room_email: str
    event_id: str

@app.post("/checkin")
async def check_in_meeting(req: CheckInRequest):
    token = await get_app_token()
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    
    # We PATCH the event to add the "Checked-In" category
    payload = {
        "categories": ["Checked-In"]
    }
    
    url = f"{GRAPH_BASE_URL}/users/{req.room_email}/events/{req.event_id}"
    
    async with httpx.AsyncClient() as client:
        resp = await client.patch(url, headers=headers, json=payload)
        
    if resp.status_code != 200:
        raise HTTPException(status_code=resp.status_code, detail="Check-in failed")
        
    return {"status": "checked-in", "msg": "You have successfully confirmed your attendance."}

# ==========================================
# 🔍 NEW ENDPOINT: GET CURRENT ACTIVE MEETING
# ==========================================
@app.get("/active-meeting")
async def get_active_meeting(room_email: str):
    """Finds the meeting happening RIGHT NOW so we can show the Check-In button."""
    token = await get_app_token()
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    
    now = datetime.utcnow()
    # Check a tight window (Now +/- 1 minute)
    start_window = (now - timedelta(minutes=15)).isoformat() + "Z" # Can check in 15 mins late max
    end_window = (now + timedelta(minutes=15)).isoformat() + "Z"   # Can check in 15 mins early
    
    url = (
        f"{GRAPH_BASE_URL}/users/{room_email}/calendarView"
        f"?startDateTime={start_window}"
        f"&endDateTime={end_window}"
        f"&$select=id,subject,categories,start,end,organizer"
        f"&$top=1"
    )
    
    async with httpx.AsyncClient() as client:
        resp = await client.get(url, headers=headers)
        if resp.status_code == 200:
            events = resp.json().get('value', [])
            if events:
                return events[0] # Return the active event
    return None # No meeting right now
@app.post("/book")
async def create_booking(req: BookingRequest, authorization: Optional[str] = Header(None)):
    """
    1. Checks Room Availability using SYSTEM Token.
    2. Books Meeting using USER Token (Logged-in User becomes Organizer).
    """
    
    if not authorization:
        raise HTTPException(status_code=401, detail="Missing User Token in Header")
    
    # Extract "Bearer <token>"
    user_token = authorization.split(" ")[1]
    system_token = await get_app_token()
    
    # ==================================================================
    # 🛑 STEP 1: ROBUST CONFLICT CHECK (Using System Token)
    # ==================================================================
    # We still check the ROOM's calendar directly to ensure it is free.
    start_dt = req.start_time.replace(tzinfo=None)
    end_dt = req.end_time.replace(tzinfo=None)
    
    start_str = quote(start_dt.isoformat() + "Z")
    end_str = quote(end_dt.isoformat() + "Z")
    
    check_url = (
        f"{GRAPH_BASE_URL}/users/{req.room_email}/calendarView"
        f"?startDateTime={start_str}"
        f"&endDateTime={end_str}"
        f"&$select=subject"
    )

    async with httpx.AsyncClient() as client:
        # Use SYSTEM TOKEN to check the Room
        check_headers = {"Authorization": f"Bearer {system_token}", "Content-Type": "application/json"}
        check_resp = await client.get(check_url, headers=check_headers)
        
        if check_resp.status_code != 200:
            raise HTTPException(status_code=500, detail="System Error: Could not verify room availability.")
            
        events = check_resp.json().get("value", [])
        if len(events) > 0:
            existing_subject = events[0].get('subject', 'Existing Booking')
            raise HTTPException(status_code=409, detail=f"Conflict! Room is already booked for: '{existing_subject}'")

    # ==================================================================
    # 🚀 STEP 2: CREATE BOOKING (Using USER Token)
    # ==================================================================
    
    # Build Attendee List
    # 1. The Room (Resource) - Crucial for Auto-Accept
    all_attendees = [
        {
            "emailAddress": {"address": req.room_email},
            "type": "resource" 
        }
    ]
    
    # 2. The Colleagues (Required)
    for email in req.attendees:
        if email.strip():
            all_attendees.append({"emailAddress": {"address": email.strip()}, "type": "required"})

    meeting_body = f"""
    <html>
    <body>
        <p><strong>Filiale:</strong> {req.filiale}</p>
        <p><strong>Reason:</strong> {req.description}</p>
    </body>
    </html>
    """

    event_payload = {
        "subject": f"{req.subject} ({req.filiale})",
        "body": {
            "contentType": "HTML",
            "content": meeting_body
        },
        "start": {"dateTime": start_dt.isoformat() + "Z", "timeZone": "UTC"},
        "end": {"dateTime": end_dt.isoformat() + "Z", "timeZone": "UTC"},
        "location": {
            "displayName": "Conference Room", 
            "locationEmailAddress": req.room_email
        },
        "attendees": all_attendees
    }

    async with httpx.AsyncClient() as client:
        # 🔴 DYNAMIC: POST TO /me/events
        # 'user_token' belongs to whoever is logged in.
        # So '/me' becomes that specific user automatically.
        user_headers = {"Authorization": f"Bearer {user_token}", "Content-Type": "application/json"}
        url = f"{GRAPH_BASE_URL}/me/events"
        resp = await client.post(url, headers=user_headers, json=event_payload)
        
    if resp.status_code != 201:
        raise HTTPException(status_code=resp.status_code, detail=f"User Booking Failed: {resp.text}")
        
    return {"status": "success", "data": resp.json()}

