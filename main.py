import os
import httpx
from fastapi import FastAPI, HTTPException, Depends
from pydantic import BaseModel
from datetime import datetime, timedelta
from typing import List, Optional
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

app = FastAPI(title="Microsoft 365 Room Booking PoC")

# --- Configuration ---
# You must set these in your environment or a .env file
TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"

# --- Pydantic Models ---

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
    organizer_email: str  # For the PoC, we might pass this manually or via token

# --- Helper: Get Microsoft Graph Token ---
async def get_graph_token():
    """
    Obtains an App-Only Access Token (Client Credentials Flow).
    Permissions required: Calendars.ReadWrite, Place.Read.All
    """
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
        raise HTTPException(status_code=401, detail=f"Failed to auth with Azure: {response.text}")
        
    return response.json().get("access_token")

# --- Endpoints ---

@app.get("/")
async def health_check():
    return {"status": "running", "service": "Room Booking Backend"}

@app.get("/rooms")
async def get_rooms():
    """
    Fetches all Room resources from the tenant.
    Graph API: GET /places/microsoft.graph.room
    """
    token = await get_graph_token()
    headers = {"Authorization": f"Bearer {token}"}
    
    async with httpx.AsyncClient() as client:
        # Note: 'places' endpoint is stable v1.0 but sometimes requires specific directory setups.
        # Alternatively use /users if rooms are just user mailboxes, but /places is the correct architectural choice.
        resp = await client.get(f"{GRAPH_BASE_URL}/places/microsoft.graph.room", headers=headers)
        
    if resp.status_code != 200:
        raise HTTPException(status_code=resp.status_code, detail=resp.text)
        
    return resp.json()

@app.post("/availability")
async def check_availability(req: AvailabilityRequest):
    """
    Checks free/busy status for a specific room.
    Graph API: POST /users/{id}/calendar/getSchedule
    """
    token = await get_graph_token()
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
        "Prefer": f'outlook.timezone="{req.time_zone}"'
    }
    
    # Graph expects a list of schedules
    payload = {
        "schedules": [req.room_email],
        "startTime": {
            "dateTime": req.start_time.isoformat(),
            "timeZone": req.time_zone
        },
        "endTime": {
            "dateTime": req.end_time.isoformat(),
            "timeZone": req.time_zone
        },
        "availabilityViewInterval": 15 # 15 minute slots
    }

    async with httpx.AsyncClient() as client:
        # Using the /users/{id} endpoint because rooms are user objects in Exchange
        url = f"{GRAPH_BASE_URL}/users/{req.room_email}/calendar/getSchedule"
        resp = await client.post(url, headers=headers, json=payload)

    if resp.status_code != 200:
        raise HTTPException(status_code=resp.status_code, detail=resp.text)
        
    return resp.json()

@app.post("/book")
async def create_booking(req: BookingRequest):
    """
    Creates a meeting in the room's calendar.
    Graph API: POST /users/{id}/events
    """
    token = await get_graph_token()
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    
    # Basic event structure
    event_payload = {
        "subject": req.subject,
        "start": {
            "dateTime": req.start_time.isoformat(),
            "timeZone": "UTC"
        },
        "end": {
            "dateTime": req.end_time.isoformat(),
            "timeZone": "UTC"
        },
        "attendees": [
            {
                "emailAddress": {
                    "address": req.organizer_email,
                    "name": "Organizer"
                },
                "type": "required"
            }
        ]
    }

    async with httpx.AsyncClient() as client:
        url = f"{GRAPH_BASE_URL}/users/{req.room_email}/events"
        resp = await client.post(url, headers=headers, json=event_payload)
        
    if resp.status_code != 201:
        raise HTTPException(status_code=resp.status_code, detail=resp.text)
        
    return {"status": "success", "data": resp.json()}
