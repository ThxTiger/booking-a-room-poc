# ==========================================
# Microsoft 365 Room Booking Backend (Final v12)
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

load_dotenv()
app = FastAPI(title="Vinci Energies Room Booking API", version="12.0.0")

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

@app.get("/active-meeting")
async def get_active_meeting(room_email: str):
    token = await get_app_token()
    now = datetime.utcnow()
    start_win = now.isoformat() + "Z"
    end_win = (now + timedelta(hours=12)).isoformat() + "Z"
    
    url = f"{GRAPH_BASE_URL}/users/{room_email}/calendarView?startDateTime={start_win}&endDateTime={end_win}&$select=id,subject,categories,start,end,organizer&$orderby=start/dateTime&$top=1"
    async with httpx.AsyncClient() as client:
        resp = await client.get(url, headers={"Authorization": f"Bearer {token}"})
        if resp.status_code == 200:
            events = resp.json().get('value', [])
            if events: return events[0]
            
    past_start = (now - timedelta(minutes=60)).isoformat() + "Z"
    url_past = f"{GRAPH_BASE_URL}/users/{room_email}/calendarView?startDateTime={past_start}&endDateTime={start_win}&$select=id,subject,categories,start,end,organizer&$orderby=start/dateTime desc&$top=1"
    async with httpx.AsyncClient() as client:
        resp = await client.get(url_past, headers={"Authorization": f"Bearer {token}"})
        if resp.status_code == 200:
            events = resp.json().get('value', [])
            if events: return events[0]
    return None

@app.post("/checkin")
async def check_in_meeting(req: CheckInRequest):
    token = await get_app_token()
    async with httpx.AsyncClient() as client:
        await client.patch(f"{GRAPH_BASE_URL}/users/{req.room_email}/events/{req.event_id}", headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"}, json={"categories": ["Checked-In"]})
    return {"status": "checked-in"}

# 🔴 NEW ENDPOINT: END MEETING
@app.post("/end-meeting")
async def end_meeting(req: CheckInRequest):
    token = await get_app_token()
    # Set the meeting End Time to NOW
    now = datetime.utcnow().isoformat() + "Z"
    
    payload = {
        "end": {
            "dateTime": now,
            "timeZone": "UTC"
        }
    }
    
    url = f"{GRAPH_BASE_URL}/users/{req.room_email}/events/{req.event_id}"
    
    async with httpx.AsyncClient() as client:
        resp = await client.patch(url, headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"}, json=payload)
        
    if resp.status_code != 200:
        raise HTTPException(status_code=resp.status_code, detail="Failed to end meeting")
        
    return {"status": "ended"}
