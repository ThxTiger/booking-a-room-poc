# ==========================================
# Microsoft 365 Room Booking Backend (Final v16 - BFF & Secure Cookies)
# ==========================================
import os
import httpx
import asyncio
import secrets
import urllib.parse
from fastapi import FastAPI, HTTPException, Request, Response, status
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import RedirectResponse
from pydantic import BaseModel
from datetime import datetime, timedelta
from typing import List, Optional
from dotenv import load_dotenv

load_dotenv()
app = FastAPI(title="Vinci Energies Room Booking API", version="16.0.0")

# --- CONFIGURATION ---
TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"

# Hardcoded for now, can be moved to .env later:
# FRONTEND_URL = os.getenv("FRONTEND_URL")
# BACKEND_URL = os.getenv("BACKEND_URL")
FRONTEND_URL = "https://booking-frontend-three-flax.vercel.app"
BACKEND_URL = "https://booking-a-room-poc.onrender.com"

# CORS must have explicit origins when allow_credentials=True
app.add_middleware(
    CORSMiddleware,
    allow_origins=[FRONTEND_URL, "http://localhost:5500", "http://127.0.0.1:5500"], 
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

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

class ExtendRequest(BaseModel):
    room_email: str
    event_id: str
    extend_minutes: int = 15

# --- SESSION MANAGEMENT (BFF / HttpOnly Cookies) ---
COOKIE_NAME = "room_session"
_sessions: dict = {}   # { sid -> { access_token, username, expires_at } }

def _get_session(request: Request):
    sid = request.cookies.get(COOKIE_NAME)
    if not sid: return None
    sess = _sessions.get(sid)
    if not sess: return None
    if datetime.utcnow() > sess["expires_at"]:
        _sessions.pop(sid, None)
        return None
    return sess

def _require_session(request: Request):
    sess = _get_session(request)
    if not sess: 
        raise HTTPException(status_code=401, detail="Not authenticated.")
    return sess

def _create_session(token: str, username: str, expires_in: int = 3600):
    sid = secrets.token_urlsafe(32)
    _sessions[sid] = {
        "access_token": token, 
        "username": username,
        "expires_at": datetime.utcnow() + timedelta(seconds=expires_in)
    }
    return sid

# --- AUTH ROUTES ---
@app.get("/auth/login")
async def auth_login():
    state = secrets.token_urlsafe(16)
    _sessions[f"_state_{state}"] = {"ts": datetime.utcnow()}
    params = urllib.parse.urlencode({
        "client_id": CLIENT_ID, 
        "response_type": "code",
        "redirect_uri": f"{BACKEND_URL}/auth/callback",
        "response_mode": "query",
        "scope": "openid profile User.Read Calendars.ReadWrite",
        "state": state, 
        "prompt": "select_account",
    })
    return RedirectResponse(f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/authorize?{params}")

@app.get("/auth/callback")
async def auth_callback(code: str = None, state: str = None, error: str = None, error_description: str = None):
    if error: 
        return RedirectResponse(f"{FRONTEND_URL}?auth_error={urllib.parse.quote(error_description or error)}")
    if not code: 
        return RedirectResponse(f"{FRONTEND_URL}?auth_error=no_code")
    if f"_state_{state}" not in _sessions: 
        return RedirectResponse(f"{FRONTEND_URL}?auth_error=invalid_state")
    
    _sessions.pop(f"_state_{state}")
    
    async with httpx.AsyncClient() as c:
        tr = await c.post(
            f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token",
            data={
                "client_id": CLIENT_ID, 
                "client_secret": CLIENT_SECRET, 
                "code": code,
                "redirect_uri": f"{BACKEND_URL}/auth/callback", 
                "grant_type": "authorization_code",
                "scope": "openid profile User.Read Calendars.ReadWrite"
            }
        )
    
    td = tr.json()
    token = td.get("access_token")
    exp = int(td.get("expires_in", 3600))
    
    if not token: 
        return RedirectResponse(f"{FRONTEND_URL}?auth_error=token_exchange_failed")
    
    async with httpx.AsyncClient() as c:
        me = await c.get(
            f"{GRAPH_BASE_URL}/me?$select=userPrincipalName,mail",
            headers={"Authorization": f"Bearer {token}"}
        )
    
    me_data = me.json()
    username = me_data.get("userPrincipalName") or me_data.get("mail") or "unknown"
    sid = _create_session(token, username, exp)
    
    resp = RedirectResponse(url=f"{FRONTEND_URL}?auth=success", status_code=302)
    resp.set_cookie(
        key=COOKIE_NAME, 
        value=sid, 
        httponly=True, 
        secure=True,
        samesite="none", # Required for cross-site cookie sending
        max_age=exp, 
        path="/"
    )
    return resp

@app.get("/auth/me")
async def auth_me(request: Request):
    sess = _get_session(request)
    if not sess: 
        raise HTTPException(401, "No active session")
    return {"authenticated": True, "username": sess["username"]}

@app.post("/auth/logout")
async def auth_logout(request: Request, response: Response):
    sid = request.cookies.get(COOKIE_NAME)
    if sid: 
        _sessions.pop(sid, None)
    response.delete_cookie(COOKIE_NAME, path="/", samesite="none", secure=True)
    return {"status": "logged_out"}

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
            
            # Fetch rooms logic to iterate
            rooms_resp = await get_rooms() 
            
            for room in rooms_resp['value']:
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
@app.post("/extend-meeting")
async def extend_meeting(req: ExtendRequest):
    token = await get_app_token()
    now = datetime.utcnow()
    new_end = (now + timedelta(minutes=req.extend_minutes)).isoformat() + "Z"
    
    async with httpx.AsyncClient() as client:
        ev = await client.get(f"{GRAPH_BASE_URL}/users/{req.room_email}/events/{req.event_id}?$select=end", headers={"Authorization": f"Bearer {token}"})
        current_end = datetime.fromisoformat(ev.json()["end"]["dateTime"].replace("Z",""))
        new_end_dt = current_end + timedelta(minutes=req.extend_minutes)
        resp = await client.patch(
            f"{GRAPH_BASE_URL}/users/{req.room_email}/events/{req.event_id}", 
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"}, 
            json={"end": {"dateTime": new_end_dt.isoformat() + "Z", "timeZone": "UTC"}}
        )
    if resp.status_code != 200:
        raise HTTPException(status_code=resp.status_code, detail="Failed to extend")
    return {"status": "extended", "new_end": new_end_dt.isoformat()}

@app.get("/rooms")
async def get_rooms():
    return {"value": [
        {"displayName": "Conference Room A", "emailAddress": "ConferenceRoomA@AxiansPoc611.onmicrosoft.com",
         "floor": "Floor 3", "department": "Axians", "capacity": 8, "location": "Casablanca HQ"},
        {"displayName": "Conference Room C", "emailAddress": "ConferenceRoomC@AxiansPoc611.onmicrosoft.com",
         "floor": "Floor 4", "department": "QHSE", "capacity": 8, "location": "Casablanca HQ"}
    ]}

@app.post("/availability")
async def check_availability(req: AvailabilityRequest):
    token = await get_app_token()
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json", "Prefer": f'outlook.timezone="{req.time_zone}"'}
    payload = {"schedules": [req.room_email], "startTime": {"dateTime": req.start_time.isoformat(), "timeZone": req.time_zone}, "endTime": {"dateTime": req.end_time.isoformat(), "timeZone": req.time_zone}, "availabilityViewInterval": 15}
    async with httpx.AsyncClient() as client:
        resp = await client.post(f"{GRAPH_BASE_URL}/users/{req.room_email}/calendar/getSchedule", headers=headers, json=payload)
    return resp.json()

@app.get("/active-meeting")
async def get_active_meeting(room_email: str):
    token = await get_app_token()
    now = datetime.utcnow()
    found_event = None

    start_win = now.isoformat() + "Z"
    end_win = (now + timedelta(hours=12)).isoformat() + "Z"
    url_future = f"{GRAPH_BASE_URL}/users/{room_email}/calendarView?startDateTime={start_win}&endDateTime={end_win}&$select=id,subject,bodyPreview,categories,start,end,organizer,attendees&$orderby=start/dateTime&$top=1"

    async with httpx.AsyncClient() as client:
        resp = await client.get(url_future, headers={"Authorization": f"Bearer {token}"})
        if resp.status_code == 200 and resp.json().get('value'):
            found_event = resp.json().get('value')[0]

    if not found_event:
        past_start = (now - timedelta(minutes=60)).isoformat() + "Z"
        url_past = f"{GRAPH_BASE_URL}/users/{room_email}/calendarView?startDateTime={past_start}&endDateTime={start_win}&$select=id,subject,bodyPreview,categories,start,end,organizer,attendees&$orderby=start/dateTime desc&$top=1"
        async with httpx.AsyncClient() as client:
            resp = await client.get(url_past, headers={"Authorization": f"Bearer {token}"})
            if resp.status_code == 200 and resp.json().get('value'):
                found_event = resp.json().get('value')[0]

    return found_event 

@app.post("/checkin")
async def check_in_meeting(req: CheckInRequest):  
    token = await get_app_token()
    async with httpx.AsyncClient() as client:
        await client.patch(f"{GRAPH_BASE_URL}/users/{req.room_email}/events/{req.event_id}",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json={"categories": ["Checked-In"]})
    return {"status": "checked-in"}

# 🔒 SECURE BOOKING: Now requires the HttpOnly Cookie session
@app.post("/book")
async def create_booking(req: BookingRequest, request: Request):
    # This automatically blocks requests without a valid cookie
    sess = _require_session(request)
    user_token = sess["access_token"]
    username = sess["username"]
    system_token = await get_app_token()
    
    start_str = urllib.parse.quote(req.start_time.replace(tzinfo=None).isoformat() + "Z")
    end_str = urllib.parse.quote(req.end_time.replace(tzinfo=None).isoformat() + "Z")
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
        # Use the delegated user token from the secure session cookie
        resp = await client.post(
            f"{GRAPH_BASE_URL}/me/events", 
            headers={"Authorization": f"Bearer {user_token}", "Content-Type": "application/json"}, 
            json=event_payload
        )
    if resp.status_code != 201: 
        raise HTTPException(status_code=resp.status_code, detail=f"Booking Failed: {resp.text}")
    return {"status": "success", "data": resp.json()}

# 🔒 SECURE END MEETING: Now requires the HttpOnly Cookie session
@app.post("/end-meeting")
async def end_meeting(req: CheckInRequest, request: Request):
    # Protect route using the cookie
    sess = _require_session(request)
    
    token = await get_app_token()
    now = datetime.utcnow().isoformat() + "Z"
    
    payload = { "end": { "dateTime": now, "timeZone": "UTC" } }
    url = f"{GRAPH_BASE_URL}/users/{req.room_email}/events/{req.event_id}"
    
    async with httpx.AsyncClient() as client:
        resp = await client.patch(url, headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"}, json=payload)
    
    if resp.status_code != 200: 
        raise HTTPException(status_code=resp.status_code, detail="Failed to end meeting")
    
    return {"status": "ended"}
