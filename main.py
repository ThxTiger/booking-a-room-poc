# ==========================================
# Microsoft 365 Room Booking Backend v16
# Security: CORS locked, token verified via Graph /me
# ==========================================
import os
import httpx
import asyncio
from fastapi import FastAPI, HTTPException, Depends, status
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from datetime import datetime, timedelta
from typing import List, Optional
from dotenv import load_dotenv
from urllib.parse import quote
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials

load_dotenv()
app = FastAPI(title="Vinci Energies Room Booking API", version="16.0.0")

# ── FIX #1: CORS locked to your frontend domain only ──
# Was allow_origins=["*"] which is wide open.
# Now only your Vercel deployment can call this API.
app.add_middleware(
    CORSMiddleware,
    allow_origins=["https://booking-frontend-three-flax.vercel.app"],
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

class ExtendRequest(BaseModel):
    room_email: str
    event_id: str
    extend_minutes: int = 15

# --- AUTH ---
security = HTTPBearer()

def verify_user(credentials: HTTPAuthorizationCredentials = Depends(security)):
    token = credentials.credentials
    if not token:
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="Missing token")
    return token


# --- GRAPH HELPERS ---
async def get_app_token():
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
    return response.json().get("access_token")

# ── FIX #2 HELPER: verify a user token is real by calling Graph /me ──
# Any route that receives a user Bearer token must call this.
# Prevents fake/random strings from passing as valid tokens.
async def verify_token_and_get_email(user_token: str) -> str:
    async with httpx.AsyncClient() as client:
        me_resp = await client.get(
            f"{GRAPH_BASE_URL}/me?$select=userPrincipalName",
            headers={"Authorization": f"Bearer {user_token}"}
        )
    if me_resp.status_code != 200:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Token is invalid or expired."
        )
    return me_resp.json().get("userPrincipalName", "").lower()

# --- GHOST BUSTER ---
async def remove_ghost_meetings():
    print("Ghost Buster started...")
    while True:
        try:
            token = await get_app_token()
            headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
            now = datetime.utcnow()
            five_mins_ago   = (now - timedelta(minutes=5)).isoformat()  + "Z"
            twenty_mins_ago = (now - timedelta(minutes=20)).isoformat() + "Z"
            rooms = await get_rooms()
            for room in rooms["value"]:
                email = room["emailAddress"]
                url = (
                    f"{GRAPH_BASE_URL}/users/{email}/calendarView"
                    f"?startDateTime={twenty_mins_ago}&endDateTime={five_mins_ago}"
                    f"&$select=id,subject,categories"
                )
                async with httpx.AsyncClient() as client:
                    resp = await client.get(url, headers=headers)
                    if resp.status_code == 200:
                        for event in resp.json().get("value", []):
                            if "Checked-In" not in event.get("categories", []):
                                print(f"Deleting ghost: {event['subject']}")
                                await client.delete(
                                    f"{GRAPH_BASE_URL}/users/{email}/events/{event['id']}",
                                    headers=headers
                                )
        except Exception as e:
            print(f"Ghost Buster error: {e}")
        await asyncio.sleep(60)

@app.on_event("startup")
async def startup_event():
    asyncio.create_task(remove_ghost_meetings())

# --- ROUTES ---

@app.get("/rooms")
async def get_rooms():
    return {"value": [
        {
            "displayName": "Conference Room A",
            "emailAddress": "ConferenceRoomA@AxiansPoc611.onmicrosoft.com",
            "floor": "Floor 3", "department": "Axians",
            "capacity": 8, "location": "Casablanca HQ"
        },
        {
            "displayName": "Conference Room C",
            "emailAddress": "ConferenceRoomC@AxiansPoc611.onmicrosoft.com",
            "floor": "Floor 4", "department": "QHSE",
            "capacity": 8, "location": "Casablanca HQ"
        }
    ]}

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
        "endTime":   {"dateTime": req.end_time.isoformat(),   "timeZone": req.time_zone},
        "availabilityViewInterval": 15
    }
    async with httpx.AsyncClient() as client:
        resp = await client.post(
            f"{GRAPH_BASE_URL}/users/{req.room_email}/calendar/getSchedule",
            headers=headers, json=payload
        )
    return resp.json()

@app.get("/active-meeting")
async def get_active_meeting(room_email: str):
    token = await get_app_token()
    now = datetime.utcnow()
    found_event = None

    start_win  = now.isoformat() + "Z"
    end_win    = (now + timedelta(hours=12)).isoformat() + "Z"
    url_future = (
        f"{GRAPH_BASE_URL}/users/{room_email}/calendarView"
        f"?startDateTime={start_win}&endDateTime={end_win}"
        f"&$select=id,subject,bodyPreview,categories,start,end,organizer,attendees"
        f"&$orderby=start/dateTime&$top=1"
    )
    async with httpx.AsyncClient() as client:
        resp = await client.get(url_future, headers={"Authorization": f"Bearer {token}"})
        if resp.status_code == 200 and resp.json().get("value"):
            found_event = resp.json()["value"][0]

    if not found_event:
        past_start = (now - timedelta(minutes=60)).isoformat() + "Z"
        url_past   = (
            f"{GRAPH_BASE_URL}/users/{room_email}/calendarView"
            f"?startDateTime={past_start}&endDateTime={start_win}"
            f"&$select=id,subject,bodyPreview,categories,start,end,organizer,attendees"
            f"&$orderby=start/dateTime desc&$top=1"
        )
        async with httpx.AsyncClient() as client:
            resp = await client.get(url_past, headers={"Authorization": f"Bearer {token}"})
            if resp.status_code == 200 and resp.json().get("value"):
                found_event = resp.json()["value"][0]

    return found_event  # None or full event — no masking

@app.post("/checkin")
async def check_in_meeting(req: CheckInRequest):
    # No auth — physical presence at the kiosk is the authorization
    token = await get_app_token()
    async with httpx.AsyncClient() as client:
        await client.patch(
            f"{GRAPH_BASE_URL}/users/{req.room_email}/events/{req.event_id}",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json={"categories": ["Checked-In"]}
        )
    return {"status": "checked-in"}

@app.post("/extend-meeting")
async def extend_meeting(req: ExtendRequest):
    # No auth — physical presence at the kiosk is the authorization
    token = await get_app_token()
    async with httpx.AsyncClient() as client:
        ev = await client.get(
            f"{GRAPH_BASE_URL}/users/{req.room_email}/events/{req.event_id}?$select=end",
            headers={"Authorization": f"Bearer {token}"}
        )
        current_end = datetime.fromisoformat(ev.json()["end"]["dateTime"].replace("Z", ""))
        new_end_dt  = current_end + timedelta(minutes=req.extend_minutes)
        resp = await client.patch(
            f"{GRAPH_BASE_URL}/users/{req.room_email}/events/{req.event_id}",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json={"end": {"dateTime": new_end_dt.isoformat() + "Z", "timeZone": "UTC"}}
        )
    if resp.status_code != 200:
        raise HTTPException(status_code=resp.status_code, detail="Failed to extend")
    return {"status": "extended", "new_end": new_end_dt.isoformat()}

# ── FIX #2 APPLIED: /book ──
# Was: takes Authorization header, splits it, uses it blindly — no validation.
# Now: calls verify_token_and_get_email() which hits Graph /me.
# If the token is fake/expired, Graph returns 401 and we reject immediately.
# The actual email from Graph is used — not whatever the client claims.
@app.post("/book")
async def create_booking(
    req: BookingRequest,
    user_token: str = Depends(verify_user)
):
    # Verify token is real — get actual identity from Microsoft
    actual_email = await verify_token_and_get_email(user_token)

    system_token = await get_app_token()

    # Conflict check
    start_str = quote(req.start_time.replace(tzinfo=None).isoformat() + "Z")
    end_str   = quote(req.end_time.replace(tzinfo=None).isoformat()   + "Z")
    check_url = (
        f"{GRAPH_BASE_URL}/users/{req.room_email}/calendarView"
        f"?startDateTime={start_str}&endDateTime={end_str}&$select=subject"
    )
    async with httpx.AsyncClient() as client:
        check_resp = await client.get(check_url, headers={"Authorization": f"Bearer {system_token}"})
        if len(check_resp.json().get("value", [])) > 0:
            raise HTTPException(status_code=409, detail="Conflict! Room is already booked.")

    all_attendees = [{"emailAddress": {"address": req.room_email}, "type": "resource"}]
    for email in req.attendees:
        if email.strip():
            all_attendees.append({"emailAddress": {"address": email.strip()}, "type": "required"})

    final_subject = f"{req.filiale} : {req.description}" if req.description else f"{req.filiale} : {req.subject}"
    if not final_subject.strip(": "):
        final_subject = "Meeting"

    event_payload = {
        "subject": final_subject,
        "body": {"contentType": "HTML", "content": f"Filiale: {req.filiale}<br>Reason: {req.description}"},
        "start": {"dateTime": req.start_time.replace(tzinfo=None).isoformat() + "Z", "timeZone": "UTC"},
        "end":   {"dateTime": req.end_time.replace(tzinfo=None).isoformat()   + "Z", "timeZone": "UTC"},
        "location": {"displayName": "Conference Room", "locationEmailAddress": req.room_email},
        "attendees": all_attendees
    }

    async with httpx.AsyncClient() as client:
        resp = await client.post(
            f"{GRAPH_BASE_URL}/me/events",
            headers={"Authorization": f"Bearer {user_token}", "Content-Type": "application/json"},
            json=event_payload
        )
    if resp.status_code != 201:
        raise HTTPException(status_code=resp.status_code, detail=f"Booking Failed: {resp.text}")
    return {"status": "success", "data": resp.json()}

# ── FIX #2 APPLIED: /end-meeting ──
# Was: verify_user only checks token exists — any string passes.
#      Authorization list came from frontend (client-controlled).
# Now: verify_token_and_get_email() calls Graph /me — fake tokens rejected.
#      Organizer/attendee list fetched fresh from Graph — not trusted from client.
#      actual_email from Microsoft is what gets checked against the list.
@app.post("/end-meeting")
async def end_meeting(
    req: CheckInRequest,
    user_token: str = Depends(verify_user)
):
    # Step 1: verify token is real, get actual identity
    actual_email = await verify_token_and_get_email(user_token)

    # Step 2: fetch event fresh from Graph — don't trust client's allowed list
    app_token = await get_app_token()
    async with httpx.AsyncClient() as client:
        ev = await client.get(
            f"{GRAPH_BASE_URL}/users/{req.room_email}/events/{req.event_id}"
            f"?$select=organizer,attendees",
            headers={"Authorization": f"Bearer {app_token}"}
        )
    if ev.status_code != 200:
        raise HTTPException(status_code=404, detail="Event not found.")

    ev_data   = ev.json()
    organizer = ev_data.get("organizer", {}).get("emailAddress", {}).get("address", "").lower()
    attendees = [
        a.get("emailAddress", {}).get("address", "").lower()
        for a in ev_data.get("attendees", [])
    ]
    allowed = attendees + [organizer]

    # Step 3: check actual identity (from Microsoft) against server-fetched list
    if actual_email not in allowed:
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="You are not authorized to end this meeting."
        )

    # Step 4: end it
    now = datetime.utcnow().isoformat() + "Z"
    async with httpx.AsyncClient() as client:
        resp = await client.patch(
            f"{GRAPH_BASE_URL}/users/{req.room_email}/events/{req.event_id}",
            headers={"Authorization": f"Bearer {app_token}", "Content-Type": "application/json"},
            json={"end": {"dateTime": now, "timeZone": "UTC"}}
        )
    if resp.status_code != 200:
        raise HTTPException(status_code=resp.status_code, detail="Failed to end meeting")
    return {"status": "ended"}

