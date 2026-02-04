# ==========================================
# Microsoft 365 Room Booking Backend (Final)
# ==========================================
import os
import httpx
import asyncio
from fastapi import FastAPI, HTTPException, Header
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from datetime import datetime, timedelta
from typing import List, Optional
from dotenv import load_dotenv
from urllib.parse import quote

# --- 1. SETUP & CONFIGURATION ---
load_dotenv()

app = FastAPI(title="Vinci Energies Room Booking API", version="3.0.0")

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

# --- 2. DATA MODELS ---
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
    organizer_email: str 
    attendees: List[str] = []
    description: str = ""
    filiale: str = ""

class CheckInRequest(BaseModel):
    room_email: str
    event_id: str

# --- 3. AUTH HELPER (SYSTEM TOKEN) ---
async def get_app_token():
    """Gets the System's 'Master Key' to check/delete meetings."""
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

# --- 4. GHOST BUSTER (BACKGROUND SERVICE) ---
async def remove_ghost_meetings():
    """
    Runs every 60 seconds.
    Checks for meetings started >5 mins ago without 'Checked-In' tag.
    Deletes them automatically.
    """
    print("👻 Ghost Buster Service Started...")
    while True:
        try:
            token = await get_app_token()
            headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
            
            # Kill Zone: Started between 5 and 20 mins ago
            now = datetime.utcnow()
            five_mins_ago = (now - timedelta(minutes=5)).isoformat() + "Z"
            twenty_mins_ago = (now - timedelta(minutes=20)).isoformat() + "Z"

            rooms = await get_rooms() 
            
            for room in rooms['value']:
                email = room['emailAddress']
                
                # Query: Find active/past meetings in the Kill Zone
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
                            # THE CHECK: Is "Checked-In" tag missing?
                            categories = event.get('categories', [])
                            
                            if "Checked-In" not in categories:
                                print(f"❌ DELETING GHOST MEETING: '{event['subject']}' in {room['displayName']}")
                                await client.delete(
                                    f"{GRAPH_BASE_URL}/users/{email}/events/{event['id']}", 
                                    headers=headers
                                )
                                
        except Exception as e:
            print(f"⚠️ Ghost Buster Error: {e}")
            
        await asyncio.sleep(60) # Wait 1 minute

@app.on_event("startup")
async def startup_event():
    asyncio.create_task(remove_ghost_meetings())

# --- 5. API ENDPOINTS ---

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

@app.post("/book")
async def create_booking(req: BookingRequest, authorization: Optional[str] = Header(None)):
    """User-Centric Booking with Hard Conflict Check."""
    if not authorization:
        raise HTTPException(status_code=401, detail="Missing User Token in Header")
    
    user_token = authorization.split(" ")[1]
    system_token = await get_app_token()
    
    # STEP 1: Hard Check (System Token)
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
        check_resp = await client.get(check_url, headers={"Authorization": f"Bearer {system_token}"})
        if check_resp.status_code != 200:
            raise HTTPException(status_code=500, detail="System Error: Could not verify room.")
        if len(check_resp.json().get("value", [])) > 0:
            raise HTTPException(status_code=409, detail="Conflict! Room is already booked.")

    # STEP 2: Book as User (User Token)
    all_attendees = [{"emailAddress": {"address": req.room_email}, "type": "resource"}]
    for email in req.attendees:
        if email.strip():
            all_attendees.append({"emailAddress": {"address": email.strip()}, "type": "required"})

    meeting_body = f"""<html><body><p><strong>Filiale:</strong> {req.filiale}</p><p><strong>Reason:</strong> {req.description}</p></body></html>"""

    event_payload = {
        "subject": f"{req.subject} ({req.filiale})",
        "body": {"contentType": "HTML", "content": meeting_body},
        "start": {"dateTime": start_dt.isoformat() + "Z", "timeZone": "UTC"},
        "end": {"dateTime": end_dt.isoformat() + "Z", "timeZone": "UTC"},
        "location": {"displayName": "Conference Room", "locationEmailAddress": req.room_email},
        "attendees": all_attendees
    }

    async with httpx.AsyncClient() as client:
        url = f"{GRAPH_BASE_URL}/me/events"
        resp = await client.post(url, headers={"Authorization": f"Bearer {user_token}", "Content-Type": "application/json"}, json=event_payload)
        
    if resp.status_code != 201:
        raise HTTPException(status_code=resp.status_code, detail=f"Booking Failed: {resp.text}")
        
    return {"status": "success", "data": resp.json()}

@app.get("/active-meeting")
async def get_active_meeting(room_email: str):
    """Finds the current meeting for Check-In."""
    token = await get_app_token()
    now = datetime.utcnow()
    # Look for meetings starting +/- 15 mins from now
    start_window = (now - timedelta(minutes=15)).isoformat() + "Z"
    end_window = (now + timedelta(minutes=15)).isoformat() + "Z"
    
    url = (
        f"{GRAPH_BASE_URL}/users/{room_email}/calendarView"
        f"?startDateTime={start_window}"
        f"&endDateTime={end_window}"
        f"&$select=id,subject,categories,start,end"
        f"&$top=1"
    )
    
    async with httpx.AsyncClient() as client:
        resp = await client.get(url, headers={"Authorization": f"Bearer {token}"})
        if resp.status_code == 200:
            events = resp.json().get('value', [])
            if events: return events[0]
    return None

@app.post("/checkin")
async def check_in_meeting(req: CheckInRequest):
    """Stamps 'Checked-In' on the meeting."""
    token = await get_app_token()
    payload = {"categories": ["Checked-In"]}
    url = f"{GRAPH_BASE_URL}/users/{req.room_email}/events/{req.event_id}"
    
    async with httpx.AsyncClient() as client:
        resp = await client.patch(url, headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"}, json=payload)
        
    if resp.status_code != 200:
        raise HTTPException(status_code=resp.status_code, detail="Check-in failed")
        
    return {"status": "checked-in"}
