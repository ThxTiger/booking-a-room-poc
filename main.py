# ==========================================
# Microsoft 365 Room Booking Backend v18
# Security fixes applied:
#   FIX-01 (v16): CORS locked
#   FIX-02 (v16): Token verified via Graph /me
#   FIX-03 (v17): slowapi rate limiting
#   FIX-04 (v17): Input validation (Field constraints)
#   FIX-05 (v17): httpx timeouts (10 s on all calls)
#   FIX-06 (v17): Swagger UI disabled in production
#   FIX-07 (v17): Security response-headers middleware
#   FIX-08 (v17): Structured logging (no meeting subjects)
#   FIX-09 (v17): Proper 404/422 instead of 500
#   FIX-10 (v18): Rate limit added to /rooms, /checkin, /extend-meeting
#   FIX-11 (v18): Query param length validation on GET /active-meeting
# ==========================================
import os
import logging
import httpx
import asyncio
from fastapi import FastAPI, HTTPException, Depends, status, Request
from fastapi.middleware.cors import CORSMiddleware
from starlette.middleware.base import BaseHTTPMiddleware
from pydantic import BaseModel, Field
from datetime import datetime, timedelta
from typing import List, Optional
from dotenv import load_dotenv
from urllib.parse import quote
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials

# FIX-03: rate limiting
from slowapi import Limiter, _rate_limit_exceeded_handler
from slowapi.util import get_remote_address
from slowapi.errors import RateLimitExceeded

load_dotenv()

# FIX-08: structured logging — never log meeting content
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s"
)
logger = logging.getLogger("room-booking")

# FIX-06: disable docs/schema when ENV=production
_is_prod = os.getenv("ENV") == "production"

app = FastAPI(
    title="Vinci Energies Room Booking API",
    version="17.0.0",
    docs_url    = None if _is_prod else "/docs",
    redoc_url   = None if _is_prod else "/redoc",
    openapi_url = None if _is_prod else "/openapi.json",
)

# FIX-03: attach limiter
limiter = Limiter(key_func=get_remote_address)
app.state.limiter = limiter
app.add_exception_handler(RateLimitExceeded, _rate_limit_exceeded_handler)

# FIX-07: security response headers added to every reply
class SecurityHeadersMiddleware(BaseHTTPMiddleware):
    async def dispatch(self, request: Request, call_next):
        response = await call_next(request)
        response.headers["Strict-Transport-Security"] = "max-age=31536000; includeSubDomains"
        response.headers["X-Content-Type-Options"]    = "nosniff"
        response.headers["X-Frame-Options"]           = "DENY"
        response.headers["Referrer-Policy"]           = "no-referrer"
        response.headers["Permissions-Policy"]        = "camera=(), microphone=(), geolocation=()"
        return response

app.add_middleware(SecurityHeadersMiddleware)

# FIX-01 (v16): CORS locked to frontend domain only
app.add_middleware(
    CORSMiddleware,
    allow_origins=["https://booking-frontend-three-flax.vercel.app"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

TENANT_ID      = os.getenv("TENANT_ID")
CLIENT_ID      = os.getenv("CLIENT_ID")
CLIENT_SECRET  = os.getenv("CLIENT_SECRET")
GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"

# FIX-05: single timeout constant used everywhere
HTTPX_TIMEOUT = 10.0

def _timeout_error():
    raise HTTPException(
        status_code=504,
        detail="External service timeout. Please try again."
    )


# ─── MODELS ───────────────────────────────────────────────────
# FIX-04: Field constraints on every user-controlled input

class AvailabilityRequest(BaseModel):
    room_email : str      = Field(..., max_length=200)
    start_time : datetime
    end_time   : datetime
    time_zone  : str      = Field("UTC", max_length=100)

class BookingRequest(BaseModel):
    subject         : str       = Field(..., min_length=1, max_length=200)
    room_email      : str       = Field(..., max_length=200)
    start_time      : datetime
    end_time        : datetime
    organizer_email : str       = Field(..., max_length=200)
    attendees       : List[str] = Field(default_factory=list)
    description     : str       = Field("",  max_length=500)
    filiale         : str       = Field("",  max_length=100)

class CheckInRequest(BaseModel):
    room_email : str = Field(..., max_length=200)
    event_id   : str = Field(..., max_length=500)

class ExtendRequest(BaseModel):
    room_email     : str = Field(..., max_length=200)
    event_id       : str = Field(..., max_length=500)
    extend_minutes : int = Field(15, ge=1, le=120)   # 1 to 120 minutes only


# ─── AUTH ─────────────────────────────────────────────────────
security = HTTPBearer()

def verify_user(credentials: HTTPAuthorizationCredentials = Depends(security)):
    token = credentials.credentials
    if not token:
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="Missing token")
    return token


# ─── GRAPH HELPERS ────────────────────────────────────────────
async def get_app_token():
    if not all([TENANT_ID, CLIENT_ID, CLIENT_SECRET]):
        raise HTTPException(status_code=500, detail="Missing Azure AD credentials.")
    token_url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = {
        "client_id"    : CLIENT_ID,
        "scope"        : "https://graph.microsoft.com/.default",
        "client_secret": CLIENT_SECRET,
        "grant_type"   : "client_credentials",
    }
    try:
        async with httpx.AsyncClient(timeout=HTTPX_TIMEOUT) as client:  # FIX-05
            response = await client.post(token_url, data=data)
    except httpx.TimeoutException:
        _timeout_error()
    return response.json().get("access_token")


# FIX-02 (v16): verify token against Graph /me — rejects fake tokens
async def verify_token_and_get_email(user_token: str) -> str:
    try:
        async with httpx.AsyncClient(timeout=HTTPX_TIMEOUT) as client:  # FIX-05
            me_resp = await client.get(
                f"{GRAPH_BASE_URL}/me?$select=userPrincipalName",
                headers={"Authorization": f"Bearer {user_token}"}
            )
    except httpx.TimeoutException:
        _timeout_error()
    if me_resp.status_code != 200:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Token is invalid or expired."
        )
    return me_resp.json().get("userPrincipalName", "").lower()


# ─── GHOST BUSTER ─────────────────────────────────────────────
async def remove_ghost_meetings():
    logger.info("Ghost Buster started.")
    while True:
        try:
            token   = await get_app_token()
            headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
            now             = datetime.utcnow()
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
                async with httpx.AsyncClient(timeout=HTTPX_TIMEOUT) as client:  # FIX-05
                    resp = await client.get(url, headers=headers)
                    if resp.status_code == 200:
                        for event in resp.json().get("value", []):
                            if "Checked-In" not in event.get("categories", []):
                                # FIX-08: log event ID only, never the subject
                                logger.info(
                                    "Ghost Buster: removing unchecked-in event id=%s room=%s",
                                    event["id"][:12], email
                                )
                                await client.delete(
                                    f"{GRAPH_BASE_URL}/users/{email}/events/{event['id']}",
                                    headers=headers
                                )
        except Exception as e:
            logger.error("Ghost Buster error: %s", str(e))
        await asyncio.sleep(60)

@app.on_event("startup")
async def startup_event():
    asyncio.create_task(remove_ghost_meetings())


# ─── ROUTES ───────────────────────────────────────────────────

@app.get("/rooms")
@limiter.limit("60/minute")
async def get_rooms(request: Request):
    return {"value": [
        {
            "displayName" : "Conference Room A",
            "emailAddress": "ConferenceRoomA@AxiansPoc611.onmicrosoft.com",
            "floor"       : "Floor 3", "department": "Axians",
            "capacity"    : 8, "location": "Casablanca HQ"
        },
        {
            "displayName" : "Conference Room C",
            "emailAddress": "ConferenceRoomC@AxiansPoc611.onmicrosoft.com",
            "floor"       : "Floor 4", "department": "QHSE",
            "capacity"    : 8, "location": "Casablanca HQ"
        }
    ]}


# FIX-03: 30 req/min — each request triggers an outbound Graph call
@app.post("/availability")
@limiter.limit("30/minute")
async def check_availability(request: Request, req: AvailabilityRequest):
    token   = await get_app_token()
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type" : "application/json",
        "Prefer"       : f'outlook.timezone="{req.time_zone}"'
    }
    payload = {
        "schedules"               : [req.room_email],
        "startTime"               : {"dateTime": req.start_time.isoformat(), "timeZone": req.time_zone},
        "endTime"                 : {"dateTime": req.end_time.isoformat(),   "timeZone": req.time_zone},
        "availabilityViewInterval": 15
    }
    try:
        async with httpx.AsyncClient(timeout=HTTPX_TIMEOUT) as client:  # FIX-05
            resp = await client.post(
                f"{GRAPH_BASE_URL}/users/{req.room_email}/calendar/getSchedule",
                headers=headers, json=payload
            )
    except httpx.TimeoutException:
        _timeout_error()
    return resp.json()


# FIX-03: 60 req/min — polled every 5 s by kiosks, but one per kiosk
@app.get("/active-meeting")
@limiter.limit("60/minute")
async def get_active_meeting(request: Request, room_email: str):
    # FIX-04: validate query param length (Pydantic Field only applies to body models)
    if not room_email or len(room_email) > 200:
        raise HTTPException(status_code=422, detail="Invalid room_email.")
    token       = await get_app_token()
    now         = datetime.utcnow()
    found_event = None

    start_win  = now.isoformat() + "Z"
    end_win    = (now + timedelta(hours=12)).isoformat() + "Z"
    url_future = (
        f"{GRAPH_BASE_URL}/users/{room_email}/calendarView"
        f"?startDateTime={start_win}&endDateTime={end_win}"
        f"&$select=id,subject,bodyPreview,categories,start,end,organizer,attendees"
        f"&$orderby=start/dateTime&$top=1"
    )
    async with httpx.AsyncClient(timeout=HTTPX_TIMEOUT) as client:  # FIX-05
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
        async with httpx.AsyncClient(timeout=HTTPX_TIMEOUT) as client:  # FIX-05
            resp = await client.get(url_past, headers={"Authorization": f"Bearer {token}"})
            if resp.status_code == 200 and resp.json().get("value"):
                found_event = resp.json()["value"][0]

    return found_event


@app.post("/checkin")
@limiter.limit("30/minute")
async def check_in_meeting(request: Request, req: CheckInRequest):
    # No auth — physical kiosk presence is the authorization
    token = await get_app_token()
    try:
        async with httpx.AsyncClient(timeout=HTTPX_TIMEOUT) as client:  # FIX-05
            resp = await client.patch(
                f"{GRAPH_BASE_URL}/users/{req.room_email}/events/{req.event_id}",
                headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
                json={"categories": ["Checked-In"]}
            )
    except httpx.TimeoutException:
        _timeout_error()
    # FIX-09: structured errors instead of 500
    if resp.status_code == 404:
        raise HTTPException(status_code=404, detail="Event not found.")
    if resp.status_code not in (200, 201):
        raise HTTPException(status_code=422, detail="Check-in failed.")
    return {"status": "checked-in"}


@app.post("/extend-meeting")
@limiter.limit("30/minute")
async def extend_meeting(request: Request, req: ExtendRequest):
    # No auth — physical kiosk presence is the authorization
    token = await get_app_token()
    async with httpx.AsyncClient(timeout=HTTPX_TIMEOUT) as client:  # FIX-05
        ev = await client.get(
            f"{GRAPH_BASE_URL}/users/{req.room_email}/events/{req.event_id}?$select=end",
            headers={"Authorization": f"Bearer {token}"}
        )
        # FIX-09: structured errors
        if ev.status_code == 404:
            raise HTTPException(status_code=404, detail="Event not found.")
        if ev.status_code != 200:
            raise HTTPException(status_code=422, detail="Could not retrieve event.")

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


# FIX-02 (v16): token verified + FIX-03: rate limited
@app.post("/book")
@limiter.limit("10/minute")
async def create_booking(
    request    : Request,
    req        : BookingRequest,
    user_token : str = Depends(verify_user)
):
    # FIX-02 (v16): verify token is real — get actual identity from Microsoft
    actual_email = await verify_token_and_get_email(user_token)

    system_token = await get_app_token()

    # Conflict check
    start_str = quote(req.start_time.replace(tzinfo=None).isoformat() + "Z")
    end_str   = quote(req.end_time.replace(tzinfo=None).isoformat()   + "Z")
    check_url = (
        f"{GRAPH_BASE_URL}/users/{req.room_email}/calendarView"
        f"?startDateTime={start_str}&endDateTime={end_str}&$select=subject"
    )
    try:
        async with httpx.AsyncClient(timeout=HTTPX_TIMEOUT) as client:  # FIX-05
            check_resp = await client.get(check_url, headers={"Authorization": f"Bearer {system_token}"})
    except httpx.TimeoutException:
        _timeout_error()
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
        "subject" : final_subject,
        "body"    : {"contentType": "HTML", "content": f"Filiale: {req.filiale}<br>Reason: {req.description}"},
        "start"   : {"dateTime": req.start_time.replace(tzinfo=None).isoformat() + "Z", "timeZone": "UTC"},
        "end"     : {"dateTime": req.end_time.replace(tzinfo=None).isoformat()   + "Z", "timeZone": "UTC"},
        "location": {"displayName": "Conference Room", "locationEmailAddress": req.room_email},
        "attendees": all_attendees
    }

    try:
        async with httpx.AsyncClient(timeout=HTTPX_TIMEOUT) as client:  # FIX-05
            resp = await client.post(
                f"{GRAPH_BASE_URL}/me/events",
                headers={"Authorization": f"Bearer {user_token}", "Content-Type": "application/json"},
                json=event_payload
            )
    except httpx.TimeoutException:
        _timeout_error()
    if resp.status_code != 201:
        raise HTTPException(status_code=resp.status_code, detail=f"Booking Failed: {resp.text}")
    return {"status": "success", "data": resp.json()}


# FIX-02 (v16): token verified + server-side attendee fetch + FIX-03: rate limited
@app.post("/end-meeting")
@limiter.limit("10/minute")
async def end_meeting(
    request    : Request,
    req        : CheckInRequest,
    user_token : str = Depends(verify_user)
):
    # Step 1: verify token is real, get actual identity from Microsoft
    actual_email = await verify_token_and_get_email(user_token)

    # Step 2: fetch event fresh from Graph — never trust the client's allowed list
    app_token = await get_app_token()
    try:
        async with httpx.AsyncClient(timeout=HTTPX_TIMEOUT) as client:  # FIX-05
            ev = await client.get(
                f"{GRAPH_BASE_URL}/users/{req.room_email}/events/{req.event_id}"
                f"?$select=organizer,attendees",
                headers={"Authorization": f"Bearer {app_token}"}
            )
    except httpx.TimeoutException:
        _timeout_error()
    # FIX-09: structured errors
    if ev.status_code == 404:
        raise HTTPException(status_code=404, detail="Event not found.")
    if ev.status_code != 200:
        raise HTTPException(status_code=422, detail="Could not retrieve event.")

    ev_data   = ev.json()
    organizer = ev_data.get("organizer", {}).get("emailAddress", {}).get("address", "").lower()
    attendees = [
        a.get("emailAddress", {}).get("address", "").lower()
        for a in ev_data.get("attendees", [])
    ]
    allowed = attendees + [organizer]

    # Step 3: check Microsoft-verified identity against server-fetched list
    if actual_email not in allowed:
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="You are not authorized to end this meeting."
        )

    # Step 4: end the meeting
    now = datetime.utcnow().isoformat() + "Z"
    try:
        async with httpx.AsyncClient(timeout=HTTPX_TIMEOUT) as client:  # FIX-05
            resp = await client.patch(
                f"{GRAPH_BASE_URL}/users/{req.room_email}/events/{req.event_id}",
                headers={"Authorization": f"Bearer {app_token}", "Content-Type": "application/json"},
                json={"end": {"dateTime": now, "timeZone": "UTC"}}
            )
    except httpx.TimeoutException:
        _timeout_error()
    if resp.status_code != 200:
        raise HTTPException(status_code=resp.status_code, detail="Failed to end meeting")
    return {"status": "ended"}
