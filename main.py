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

# --- 1. Setup & Configuration ---
load_dotenv()

app = FastAPI(
    title="Microsoft 365 Room Booking PoC", 
    description="Backend for real-time room reservation via Microsoft Graph",
    version="1.2.0"
)

# Enable CORS so your Vercel frontend can talk to this Render backend
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Replace with your Vercel URL in production for better security
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Azure AD / Microsoft Entra ID Credentials
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
    # 🆕 NEW: Accept a list of invitee emails
    attendees: List[str] = [] 

# --- 3. Auth Helper ---
async def get_graph_token():
    """Obtains an App-Only Access Token using Client Credentials Flow."""
    if not all([TENANT_ID, CLIENT_ID, CLIENT_SECRET]):
        raise HTTPException(status_code=500, detail="Missing Azure AD credentials in environment.")

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

@app.get("/")
async def health_check():
    """Verifies the service is live."""
    return {"status": "online", "service": "Axians Room Booking API"}

@app.get("/rooms")
async def get_rooms():
    """
    Returns a list of available rooms.
    Uses a static list for immediate reliability.
    """
    static_rooms = [
        {
            "displayName": "Conference Room A",
            "emailAddress": "ConferenceRoomA@AxiansPoc611.onmicrosoft.com"
        },
        {
            "displayName": "Conference Room C",
            "emailAddress": "ConferenceRoomC@AxiansPoc611.onmicrosoft.com"
        }
    ]
    return {"value": static_rooms}

@app.post("/availability")
async def check_availability(req: AvailabilityRequest):
    """
    Fetches real-time busy/free status from Microsoft Graph.
    """
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

    if resp.status_code != 200:
        raise HTTPException(status_code=resp.status_code, detail=resp.json())
        
    return resp.json()

@app.post("/book")
async def create_booking(req: BookingRequest):
    """
    Creates a meeting with attendees, but ONLY if the room is currently free.
    Returns 409 Conflict if the room is already reserved.
    """
    token = await get_graph_token()
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    
    # --- STEP 1: PRE-CHECK CONFLICTS ---
    check_payload = {
        "schedules": [req.room_email],
        "startTime": {"dateTime": req.start_time.isoformat(), "timeZone": "UTC"},
        "endTime": {"dateTime": req.end_time.isoformat(), "timeZone": "UTC"},
        "availabilityViewInterval": 15
    }
    
    async with httpx.AsyncClient() as client:
        check_url = f"{GRAPH_BASE_URL}/users/{req.room_email}/calendar/getSchedule"
        check_resp = await client.post(check_url, headers=headers, json=check_payload)
        
        if check_resp.status_code == 200:
            data = check_resp.json()
            # availabilityView is a string where '0' is free and other numbers (2, 3) are busy
            view_str = data["value"][0].get("availabilityView", "")
            
            # Check if any part of the requested time is busy
            if any(char in view_str for char in ["1", "2", "3", "4"]):
                raise HTTPException(
                    status_code=409, 
                    detail="Sorry, this room is already reserved for that time."
                )

    # --- STEP 2: PREPARE ATTENDEE LIST ---
    # Always include the Organizer (You)
    attendee_list = [
        {
            "emailAddress": {"address": req.organizer_email},
            "type": "required"
        }
    ]

    # 🆕 Loop through any invitees provided in the request
    for email in req.attendees:
        if email.strip():  # Avoid adding empty strings
            attendee_list.append({
                "emailAddress": {"address": email.strip()},
                "type": "required"
            })

    # --- STEP 3: CREATE THE EVENT ---
    event_payload = {
        "subject": req.subject,
        "start": {"dateTime": req.start_time.isoformat(), "timeZone": "UTC"},
        "end": {"dateTime": req.end_time.isoformat(), "timeZone": "UTC"},
        "attendees": attendee_list  # 🆕 Send the full list to Outlook
    }

    async with httpx.AsyncClient() as client:
        url = f"{GRAPH_BASE_URL}/users/{req.room_email}/events"
        resp = await client.post(url, headers=headers, json=event_payload)
        
    if resp.status_code != 201:
        raise HTTPException(status_code=resp.status_code, detail=resp.json())
        
    return {"status": "success", "data": resp.json()}
