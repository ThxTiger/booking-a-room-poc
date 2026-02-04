# ==========================================
# Microsoft 365 Room Booking PoC Backend
# ==========================================
import os
import httpx
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from datetime import datetime
from typing import List
from dotenv import load_dotenv

load_dotenv()

app = FastAPI(title="Axians Room Booking", version="1.3.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"], 
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"

# --- Data Models ---
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
    # 🆕 NEW FIELDS
    description: str = ""   # e.g. "For purchasing team..."
    filiale: str = ""       # e.g. "Cegelec"

# --- Auth Helper (No Changes) ---
async def get_graph_token():
    if not all([TENANT_ID, CLIENT_ID, CLIENT_SECRET]):
        raise HTTPException(status_code=500, detail="Missing Azure credentials.")
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
        raise HTTPException(status_code=401, detail=f"Auth Failed: {response.text}")
    return response.json().get("access_token")

# --- API Endpoints ---

@app.get("/rooms")
async def get_rooms():
    """
    Returns rooms with STATIC metadata (Floor, Department).
    """
    static_rooms = [
        {
            "displayName": "Conference Room A",
            "emailAddress": "ConferenceRoomA@AxiansPoc611.onmicrosoft.com",
            # 🆕 STATIC DETAILS
            "floor": "Floor 3",
            "department": "Axians",
            "capacity": "10 Seats"
        },
        {
            "displayName": "Conference Room C",
            "emailAddress": "ConferenceRoomC@AxiansPoc611.onmicrosoft.com",
            # 🆕 STATIC DETAILS
            "floor": "Floor 4",
            "department": "QHSE",
            "capacity": "6 Seats"
        }
    ]
    return {"value": static_rooms}

@app.post("/availability")
async def check_availability(req: AvailabilityRequest):
    # (Same as before - no changes needed here)
    token = await get_graph_token()
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
async def create_booking(req: BookingRequest):
    token = await get_graph_token()
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    
    # ==================================================================
    # 🛑 1. THE "HARD" CONFLICT CHECK (CalendarView)
    # ==================================================================
    # We ask for ALL events that overlap with our requested time.
    # Logic: Existing Start < Request End AND Existing End > Request Start
    start_str = req.start_time.isoformat()
    end_str = req.end_time.isoformat()
    
    # URL to search the actual calendar folder
    check_url = (
        f"{GRAPH_BASE_URL}/users/{req.room_email}/calendarView"
        f"?startDateTime={start_str}"
        f"&endDateTime={end_str}"
        f"&$select=subject,start,end"
    )

    async with httpx.AsyncClient() as client:
        check_resp = await client.get(check_url, headers=headers)
        
        if check_resp.status_code == 200:
            events = check_resp.json().get("value", [])
            if len(events) > 0:
                # ⛔ CONFLICT DETECTED
                existing_subject = events[0].get('subject', 'Unknown Meeting')
                raise HTTPException(
                    status_code=409, 
                    detail=f"Conflict! Room is already booked for: '{existing_subject}'"
                )

    # ==================================================================
    # 🚀 2. CREATE THE BOOKING (Mark as BUSY)
    # ==================================================================
    attendee_list = [{"emailAddress": {"address": req.organizer_email}, "type": "required"}]
    for email in req.attendees:
        if email.strip():
            attendee_list.append({"emailAddress": {"address": email.strip()}, "type": "required"})

    # HTML Body
    meeting_body = f"""
    <html>
    <body>
        <h3>Meeting Details</h3>
        <p><strong>Filiale:</strong> {req.filiale}</p>
        <p><strong>Description:</strong> {req.description}</p>
        <hr/>
        <p>Booked via Axians Kiosk</p>
    </body>
    </html>
    """

    event_payload = {
        "subject": f"{req.subject} ({req.filiale})",
        "body": {
            "contentType": "HTML",
            "content": meeting_body
        },
        "start": {"dateTime": start_str, "timeZone": "UTC"},
        "end": {"dateTime": end_str, "timeZone": "UTC"},
        "showAs": "busy",  # 🔴 CRITICAL: Forces the room to appear 'Busy' immediately
        "attendees": attendee_list
    }

    async with httpx.AsyncClient() as client:
        url = f"{GRAPH_BASE_URL}/users/{req.room_email}/events"
        resp = await client.post(url, headers=headers, json=event_payload)
        
    if resp.status_code != 201:
        raise HTTPException(status_code=resp.status_code, detail=resp.json())
        
    return {"status": "success", "data": resp.json()}

    # 3. 🆕 Construct Meeting Body (HTML)
    # This puts the "Reason" and "Filiale" inside the Outlook event description
    meeting_body = f"""
    <html>
    <body>
        <h3>Meeting Details</h3>
        <p><strong>Filiale:</strong> {req.filiale}</p>
        <p><strong>Description:</strong> {req.description}</p>
        <hr/>
        <p>Booked via Axians Kiosk</p>
    </body>
    </html>
    """

    event_payload = {
        "subject": f"{req.subject} ({req.filiale})", # Add Filiale to title too
        "body": {
            "contentType": "HTML",
            "content": meeting_body
        },
        "start": {"dateTime": req.start_time.isoformat(), "timeZone": "UTC"},
        "end": {"dateTime": req.end_time.isoformat(), "timeZone": "UTC"},
        "attendees": attendee_list
    }

    async with httpx.AsyncClient() as client:
        resp = await client.post(f"{GRAPH_BASE_URL}/users/{req.room_email}/events", headers=headers, json=event_payload)
        
    if resp.status_code != 201:
        raise HTTPException(status_code=resp.status_code, detail=resp.json())
        
    return {"status": "success", "data": resp.json()}

