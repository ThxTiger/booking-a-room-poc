# ==========================================
# Microsoft 365 Room Booking Backend (Final v15 - Secure Check-In)
# ==========================================
import os
import httpx
import asyncio
from fastapi import FastAPI, HTTPException, Header, Depends, status
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from datetime import datetime, timedelta
from typing import List, Optional
from dotenv import load_dotenv
from urllib.parse import quote
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials

load_dotenv()
app = FastAPI(title="Vinci Energies Room Booking API", version="15.0.0")

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

# --- MODELS ---
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

# --- AUTH HELPERS ---
security = HTTPBearer()

# 1. STRICT VERIFICATION (For Booking/Check-in)
def verify_user(credentials: HTTPAuthorizationCredentials = Depends(security)):
    token = credentials.credentials
    if not token:
         raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Invalid authentication credentials",
        )
    return token

# 2. 🆕 OPTIONAL VERIFICATION (For Passive Display/Kiosk)
async def optional_verify_user(authorization: Optional[str] = Header(None)):
    """
    Checks for a token but returns None if missing (instead of crashing).
    Allows 'Kiosk Mode' to work.
    """
    if not authorization:
        return None
    
    parts = authorization.split()
    if len(parts) == 2 and parts[0].lower() == "bearer":
        return parts[1]
    return None

# --- GRAPH HELPER ---
async def get_app_token():
    if not all([TENANT_ID, CLIENT_ID, CLIENT_SECRET]):
        raise HTTPException(status_code=500, detail="Missing Azure AD credentials.")
    token_url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = {
        "client_id": CLIENT_ID, "scope": "https://graph.microsoft.com/.default",
        "client_secret": CLIENT_SECRET, "grant_type": "client_credentials",
    }
    async with httpx.AsyncClient() as client:
        response = await client.post(token_url, data=data)
    return response.json().get("access_token")

# --- BACKGROUND SERVICE ---
async def remove_ghost_meetings():
    print("👻 Ghost Buster Service Started...")
    while True:
        try:
            token = await get_app_token()
            headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
            now = datetime.utcnow()
            five_mins_ago = (now - timedelta(minutes=5)).isoformat() + "Z"
            twenty_mins_ago = (now - timedelta(minutes=20)).isoformat() + "Z"
            rooms = await get_rooms() 
            
            for room in rooms['value']:
                email = room['emailAddress']
                url = f"{GRAPH_BASE_URL}/users/{email}/calendarView?startDateTime={twenty_mins_ago}&endDateTime={five_mins_ago}&$select=id,subject,categories"
                async with httpx.AsyncClient() as client:
                    resp = await client.get(url, headers=headers)
                    if resp.status_code == 200:
                        events = resp.json().get('value', [])
                        for event in events:
                            if "Checked-In" not in event.get('categories', []):
                                print(f"❌ DELETING GHOST MEETING: {event['subject']}")
                                await client.delete(f"{GRAPH_BASE_URL}/users/{email}/events/{event['id']}", headers=headers)
        except Exception as e:
            print(f"⚠️ Ghost Buster Error: {e}")
        await asyncio.sleep(60)

@app.on_event("startup")
async def startup_event():
    asyncio.create_task(remove_ghost_meetings())

# --- ROUTES ---

@app.get("/rooms")
async def get_rooms():
    return {"value": [
        {"displayName": "Conference Room A", "emailAddress": "ConferenceRoomA@AxiansPoc611.onmicrosoft.com", "floor": "Floor 3", "department": "Axians"},
        {"displayName": "Conference Room C", "emailAddress": "ConferenceRoomC@AxiansPoc611.onmicrosoft.com", "floor": "Floor 4", "department": "QHSE"}
    ]}

@app.post("/availability")
async def check_availability(req: AvailabilityRequest):
    token = await get_app_token()
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json", "Prefer": f'outlook.timezone="{req.time_zone}"'}
    payload = {"schedules": [req.room_email], "startTime": {"dateTime": req.start_time.isoformat(), "timeZone": req.time_zone}, "endTime": {"dateTime": req.end_time.isoformat(), "timeZone": req.time_zone}, "availabilityViewInterval": 15}
    async with httpx.AsyncClient() as client:
        resp = await client.post(f"{GRAPH_BASE_URL}/users/{req.room_email}/calendar/getSchedule", headers=headers, json=payload)
    return resp.json()

@app.post("/book")
async def create_booking(req: BookingRequest, authorization: Optional[str] = Header(None)):
    if not authorization: raise HTTPException(status_code=401, detail="Missing User Token")
    user_token = authorization.split(" ")[1]
    system_token = await get_app_token()
    
    start_str = quote(req.start_time.replace(tzinfo=None).isoformat() + "Z")
    end_str = quote(req.end_time.replace(tzinfo=None).isoformat() + "Z")
    check_url = f"{GRAPH_BASE_URL}/users/{req.room_email}/calendarView?startDateTime={start_str}&endDateTime={end_str}&$select=subject"
    
    async with httpx.AsyncClient() as client:
        check_resp = await client.get(check_url, headers={"Authorization": f"Bearer {system_token}"})
        if len(check_resp.json().get("value", [])) > 0:
            raise HTTPException(status_code=409, detail="Conflict! Room is already booked.")

    all_attendees = [{"emailAddress": {"address": req.room_email}, "type": "resource"}]
    for email in req.attendees:
        if email.strip(): all_attendees.append({"emailAddress": {"address": email.strip()}, "type": "required"})
    
    final_subject = f"{req.filiale} : {req.description}" if req.description else f"{req.filiale} : {req.subject}"
    if not final_subject: final_subject = "Meeting"

    event_payload = {
        "subject": final_subject, 
        "body": {"contentType": "HTML", "content": f"Filiale: {req.filiale}<br>Reason: {req.description}"},
        "start": {"dateTime": req.start_time.replace(tzinfo=None).isoformat() + "Z", "timeZone": "UTC"},
        "end": {"dateTime": req.end_time.replace(tzinfo=None).isoformat() + "Z", "timeZone": "UTC"},
        "location": {"displayName": "Conference Room", "locationEmailAddress": req.room_email},
        "attendees": all_attendees
    }

    async with httpx.AsyncClient() as client:
        resp = await client.post(f"{GRAPH_BASE_URL}/me/events", headers={"Authorization": f"Bearer {user_token}", "Content-Type": "application/json"}, json=event_payload)
    if resp.status_code != 201: raise HTTPException(status_code=resp.status_code, detail=f"Booking Failed: {resp.text}")
    return {"status": "success", "data": resp.json()}

# --- 🆕 UPDATED: ACTIVE MEETING WITH ATTENDEES & PRIVACY ---
@app.get("/active-meeting")
async def get_active_meeting(
    room_email: str, 
    user_token: Optional[str] = Depends(optional_verify_user) # <--- ALLOWS ANONYMOUS ACCESS
):
    token = await get_app_token()
    now = datetime.utcnow()
    
    # 1. FETCH MEETING (Logic consolidated)
    found_event = None
    
    # Try Future (Look ahead 12 hours)
    start_win = now.isoformat() + "Z"
    end_win = (now + timedelta(hours=12)).isoformat() + "Z"
    
    # 🛠️ UPDATE: ADDED 'attendees' TO SELECT LIST
    url_future = f"{GRAPH_BASE_URL}/users/{room_email}/calendarView?startDateTime={start_win}&endDateTime={end_win}&$select=id,subject,bodyPreview,categories,start,end,organizer,attendees&$orderby=start/dateTime&$top=1"
    
    async with httpx.AsyncClient() as client:
        resp = await client.get(url_future, headers={"Authorization": f"Bearer {token}"})
        if resp.status_code == 200 and resp.json().get('value'):
            found_event = resp.json().get('value')[0]
    
    # If no future meeting, check Past (Active Now but started earlier)
    if not found_event:
        past_start = (now - timedelta(minutes=60)).isoformat() + "Z"
        # 🛠️ UPDATE: ADDED 'attendees' TO SELECT LIST
        url_past = f"{GRAPH_BASE_URL}/users/{room_email}/calendarView?startDateTime={past_start}&endDateTime={start_win}&$select=id,subject,bodyPreview,categories,start,end,organizer,attendees&$orderby=start/dateTime desc&$top=1"
        async with httpx.AsyncClient() as client:
            resp = await client.get(url_past, headers={"Authorization": f"Bearer {token}"})
            if resp.status_code == 200 and resp.json().get('value'):
                found_event = resp.json().get('value')[0]

    # 2. IF NO MEETING, RETURN NONE
    if not found_event:
        return None

    # 3. 🔒 PRIVACY FILTER (SERVER-SIDE SANITIZATION) 🔒
    # If the user is NOT logged in (user_token is None), HIDE the secrets.
    if not user_token:
        # print(f"🔒 Masking data for {room_email} (Kiosk Mode)")
        found_event["subject"] = "Busy"
        found_event["bodyPreview"] = "Details hidden for privacy."
        
        # Hide Organizer
        if "organizer" in found_event:
            found_event["organizer"]["emailAddress"]["name"] = "Occupied"
            found_event["organizer"]["emailAddress"]["address"] = ""
        
        # Hide Attendees list so hackers can't see who is invited
        if "attendees" in found_event:
             found_event["attendees"] = []
    
    return found_event

# 🔒 SECURE CHECK-IN: Now requires 'verify_user' (Strict Mode)
@app.post("/checkin")
async def check_in_meeting(
    req: CheckInRequest, 
    user_token: str = Depends(verify_user) # <--- THIS ENFORCES AUTH
):
    token = await get_app_token()
    async with httpx.AsyncClient() as client:
        await client.patch(f"{GRAPH_BASE_URL}/users/{req.room_email}/events/{req.event_id}", headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"}, json={"categories": ["Checked-In"]})
    return {"status": "checked-in"}

@app.post("/end-meeting")
async def end_meeting(req: CheckInRequest):
    token = await get_app_token()
    now = datetime.utcnow().isoformat() + "Z"
    payload = { "end": { "dateTime": now, "timeZone": "UTC" } }
    url = f"{GRAPH_BASE_URL}/users/{req.room_email}/events/{req.event_id}"
    async with httpx.AsyncClient() as client:
        resp = await client.patch(url, headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"}, json=payload)
    if resp.status_code != 200: raise HTTPException(status_code=resp.status_code, detail="Failed to end meeting")
    return {"status": "ended"}
