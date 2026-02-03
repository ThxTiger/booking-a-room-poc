# ==========================================
# Microsoft 365 Room Booking PoC Backend
# ==========================================
import os
import httpx
from fastapi import FastAPI, HTTPException, Depends
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from datetime import datetime
from dotenv import load_dotenv

# --- 1. Setup & Configuration ---
load_dotenv()

app = FastAPI(title="Microsoft 365 Room Booking PoC", version="1.0.0")

# Enable CORS for frontend access
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"], # In production, replace "*" with your frontend domain
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Load credentials from environment variables
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

# --- 3. Auth Helper ---
async def get_graph_token():
    """Obtains App-Only Access Token from Azure AD."""
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
    return {"status": "online", "service": "Room Booking Backend API"}

@app.get("/rooms")
async def get_rooms():
    """
    Returns a list of available rooms.
    Currently uses a STATIC list for immediate availability.
    Uncomment the Graph API code below to switch to dynamic fetching later.
    """
    # --- STATIC LIST (Temporary fallback) ---
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
    # Wrapping in a 'value' key to match Graph API response structure
    return {"value": static_rooms}

    # --- DYNAMIC GRAPH API CALL (Uncomment later) ---
    # token = await get_graph_token()
    # headers = {"Authorization": f"Bearer {token}"}
    # async with httpx.AsyncClient() as client:
    #     resp = await client.get(f"{GRAPH_BASE_URL}/places/microsoft.graph.room", headers=headers)
    # if resp.status_code != 200:
    #     raise HTTPException(status_code=resp.status_code, detail=resp.text)
    # return resp.json()


@app.post("/availability")
async def check_availability(req: AvailabilityRequest):
    """
    Checks free/busy status for a specific room using Microsoft Graph.
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
        "availabilityViewInterval": 15 # Check in 15-minute slots
    }

    async with httpx.AsyncClient() as client:
        # We use the /users endpoint as rooms are technically user objects in Exchange
        url = f"{GRAPH_BASE_URL}/users/{req.room_email}/calendar/getSchedule"
        resp = await client.post(url, headers=headers, json=payload)

    if resp.status_code != 200:
        # Pass through Graph API errors for easier debugging
        raise HTTPException(status_code=resp.status_code, detail=resp.json())
        
    return resp.json()


@app.post("/book")
async def create_booking(req: BookingRequest):
    """
    Creates a meeting event directly in the room's calendar.
    """
    token = await get_graph_token()
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    
    event_payload = {
        "subject": req.subject,
        "start": {"dateTime": req.start_time.isoformat(), "timeZone": "UTC"},
        "end": {"dateTime": req.end_time.isoformat(), "timeZone": "UTC"},
        # Add the organizer as an attendee so they get the invite/confirmation
        "attendees": [
            {
                "emailAddress": {"address": req.organizer_email},
                "type": "required"
            }
        ]
    }

    async with httpx.AsyncClient() as client:
        url = f"{GRAPH_BASE_URL}/users/{req.room_email}/events"
        resp = await client.post(url, headers=headers, json=event_payload)
        
    if resp.status_code != 201:
        raise HTTPException(status_code=resp.status_code, detail=resp.json())
        
    return {"status": "success", "data": resp.json()}
