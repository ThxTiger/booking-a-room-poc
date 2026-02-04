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
from urllib.parse import quote  # 🔴 CRITICAL: Used to safely encode dates in URLs

# --- 1. Setup & Configuration ---
load_dotenv()

app = FastAPI(
    title="Vinci Energies Room Booking API", 
    version="1.4.0"
)

# Enable CORS for Vercel
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
    organizer_email: str
    attendees: List[str] = []
    description: str = ""   # Meeting reason
    filiale: str = ""       # Vinci Business Unit

# --- 3. Auth Helper ---
async def get_graph_token():
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
    """Returns static rooms with metadata (Floor, Department) for the frontend."""
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
    # 🛑 STEP 1: ROBUST CONFLICT CHECK (Direct Calendar Query)
    # ==================================================================
    # Use 'Z' to force UTC and 'quote' to make the URL safe.
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
        check_resp = await client.get(check_url, headers=headers)
        
        # Stop if Microsoft returns an error instead of assuming room is free
        if check_resp.status_code != 200:
            raise HTTPException(status_code=500, detail="System Error: Could not verify availability.")
            
        events = check_resp.json().get("value", [])
        
        # Block booking if any events overlap with the requested time
        if len(events) > 0:
            existing_subject = events[0].get('subject', 'Existing Booking')
            raise HTTPException(
                status_code=409, 
                detail=f"Conflict! Room is already booked for: '{existing_subject}'"
            )

    # ==================================================================
    # 🚀 STEP 2: CREATE THE BOOKING
    # ==================================================================
    # Build attendee list (Organizer + Invitees)
    attendee_list = [{"emailAddress": {"address": req.organizer_email}, "type": "required"}]
    for email in req.attendees:
        if email.strip():
            attendee_list.append({"emailAddress": {"address": email.strip()}, "type": "required"})

    # Dynamic HTML Body for Outlook
    meeting_body = f"""
    <html>
    <body>
        <h3>Meeting Information</h3>
        <p><strong>Filiale:</strong> {req.filiale}</p>
        <p><strong>Reason:</strong> {req.description}</p>
        <hr/>
        <p style='color: #888;'>Booked via Axians Room Booking App</p>
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
        "showAs": "busy",  # Ensures the slot is blocked immediately
        "attendees": attendee_list
    }

    async with httpx.AsyncClient() as client:
        url = f"{GRAPH_BASE_URL}/users/{req.room_email}/events"
        resp = await client.post(url, headers=headers, json=event_payload)
        
    if resp.status_code != 201:
        raise HTTPException(status_code=resp.status_code, detail=resp.json())
        
    return {"status": "success", "data": resp.json()}
