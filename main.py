# ==========================================
# Microsoft 365 Room Booking Backend v21
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
#   FIX-12 (v19): Added CSP, COEP, COOP, CORP headers to middleware
#   FIX-13 (v20): /active-meeting returns up to 5 upcoming events
#   FIX-14 (v21): App token cached (1 hour) — no longer fetched per request
#   FIX-15 (v21): Email + event_id format validated before Graph URL injection
#   FIX-16 (v21): /checkin enforces 5-min window + already-checked-in guard
#   FIX-17 (v21): /extend-meeting enforces active + checked-in + server-side conflict check
#   FIX-18 (v21): Plain text body with \r\n so bodyPreview is parseable on kiosk
#   FIX-19 (v22): Ghost Buster no longer calls the rate-limited get_rooms() route.
#                 Calling a @limiter.limit() decorated route without a real Request
#                 object raises AttributeError in get_remote_address(), which was
#                 silently caught — meaning the ghost buster never deleted anything.
#                 Room data is now served from _rooms_data() (no decorator) and
#                 get_rooms() delegates to it.
# ==========================================
import os
import re
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

# ─── LOGGING ──────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s"
)
logger = logging.getLogger("room-booking")

# ─── APP ──────────────────────────────────────────────────────
_is_prod = os.getenv("ENV") == "production"

app = FastAPI(
    title="Vinci Energies Room Booking API",
    version="22.0.0",
    docs_url    = None if _is_prod else "/docs",
    redoc_url   = None if _is_prod else "/redoc",
    openapi_url = None if _is_prod else "/openapi.json",
)

limiter = Limiter(key_func=get_remote_address)
app.state.limiter = limiter
app.add_exception_handler(RateLimitExceeded, _rate_limit_exceeded_handler)

# ─── SECURITY HEADERS ─────────────────────────────────────────
class SecurityHeadersMiddleware(BaseHTTPMiddleware):
    async def dispatch(self, request: Request, call_next):
        response = await call_next(request)
        response.headers["Strict-Transport-Security"]    = "max-age=31536000; includeSubDomains"
        response.headers["X-Content-Type-Options"]       = "nosniff"
        response.headers["X-Frame-Options"]              = "DENY"
        response.headers["Referrer-Policy"]              = "no-referrer"
        response.headers["Permissions-Policy"]           = "camera=(), microphone=(), geolocation=()"
        response.headers["Content-Security-Policy"]      = "default-src 'none'"
        response.headers["Cross-Origin-Embedder-Policy"] = "require-corp"
        response.headers["Cross-Origin-Opener-Policy"]   = "same-origin"
        response.headers["Cross-Origin-Resource-Policy"] = "same-origin"
        return response

app.add_middleware(SecurityHeadersMiddleware)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["https://booking-frontend-three-flax.vercel.app"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ─── CONFIG ───────────────────────────────────────────────────
TENANT_ID      = os.getenv("TENANT_ID")
CLIENT_ID      = os.getenv("CLIENT_ID")
CLIENT_SECRET  = os.getenv("CLIENT_SECRET")
GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
HTTPX_TIMEOUT  = 10.0

# ─── FIX-15: INPUT FORMAT VALIDATORS ─────────────────────────
EMAIL_RE   = re.compile(r'^[\w._%+\-]+@[\w.\-]+\.[a-zA-Z]{2,}$')
EVENTID_RE = re.compile(r'^[A-Za-z0-9\-_=+/]+$')

def validate_email(v: str, label: str = "email"):
    if not v or not EMAIL_RE.match(v):
        raise HTTPException(status_code=422, detail=f"Invalid {label} format.")

def validate_event_id(v: str):
    if not v or not EVENTID_RE.match(v):
        raise HTTPException(status_code=422, detail="Invalid event ID format.")

# ─── MODELS ───────────────────────────────────────────────────
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
    extend_minutes : int = Field(15, ge=1, le=120)

# ─── AUTH ─────────────────────────────────────────────────────
security = HTTPBearer()

def verify_user(credentials: HTTPAuthorizationCredentials = Depends(security)):
    token = credentials.credentials
    if not token:
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="Missing token")
    return token

def _timeout_error():
    raise HTTPException(status_code=504, detail="External service timeout. Please try again.")

# ─── TOKEN CACHE ──────────────────────────────────────────────
_token_cache: dict = {"token": None, "expires_at": datetime.min}

async def get_app_token() -> str:
    global _token_cache
    if _token_cache["token"] and datetime.utcnow() < _token_cache["expires_at"]:
        return _token_cache["token"]

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
        async with httpx.AsyncClient(timeout=HTTPX_TIMEOUT) as client:
            response = await client.post(token_url, data=data)
    except httpx.TimeoutException:
        _timeout_error()

    result = response.json()
    if "access_token" not in result:
        raise HTTPException(status_code=500, detail="Failed to obtain app token.")

    _token_cache = {
        "token"     : result["access_token"],
        "expires_at": datetime.utcnow() + timedelta(seconds=result.get("expires_in", 3600) - 60)
    }
    logger.info("App token refreshed. Expires in ~%ds", result.get("expires_in", 3600) - 60)
    return _token_cache["token"]

async def verify_token_and_get_email(user_token: str) -> str:
    try:
        async with httpx.AsyncClient(timeout=HTTPX_TIMEOUT) as client:
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

# ─── FIX-19: INTERNAL ROOM LIST ───────────────────────────────
# Pure data helper — no FastAPI decorators, no rate limiter.
# Always call this from internal code (ghost buster, etc.).
# The public /rooms route delegates to this so the data stays
# in one place.
def _rooms_data() -> list:
    return [
        {
            "displayName" : "Conference Room A",
            "emailAddress": "ConferenceRoomA@VINCIEnergies1.onmicrosoft.com",
            "floor"       : "Floor 3",
            "department"  : "Axians",
            "capacity"    : 8,
            "location"    : "Casablanca HQ"
        },
        {
            "displayName" : "Conference Room C",
            "emailAddress": "ConferenceRoomC@VINCIEnergies1.onmicrosoft.com",
            "floor"       : "Floor 3",
            "department"  : "Axians",
            "capacity"    : 8,
            "location"    : "Casablanca HQ"
        }
    ]

# ─── GHOST BUSTER ─────────────────────────────────────────────
# Removes meetings that were booked but never checked in within 5 minutes.
# FIX-19: now uses _rooms_data() instead of await get_rooms().
# Previously: get_rooms() is decorated with @limiter.limit(), which calls
# get_remote_address(request) with request=None → AttributeError → exception
# silently swallowed → ghost buster never deleted anything.
async def remove_ghost_meetings():
    logger.info("Ghost Buster started.")
    while True:
        try:
            token   = await get_app_token()
            headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
            now             = datetime.utcnow()
            five_mins_ago   = (now - timedelta(minutes=5)).isoformat()  + "Z"
            twenty_mins_ago = (now - timedelta(minutes=20)).isoformat() + "Z"

            # FIX-19: use the plain internal helper, not the rate-limited route
            for room in _rooms_data():
                email = room["emailAddress"]
                url = (
                    f"{GRAPH_BASE_URL}/users/{email}/calendarView"
                    f"?startDateTime={twenty_mins_ago}&endDateTime={five_mins_ago}"
                    f"&$select=id,subject,categories"
                )
                async with httpx.AsyncClient(timeout=HTTPX_TIMEOUT) as client:
                    resp = await client.get(url, headers=headers)
                    if resp.status_code == 200:
                        for event in resp.json().get("value", []):
                            if "Checked-In" not in event.get("categories", []):
                                logger.info(
                                    "Ghost Buster: removing unchecked-in event id=%s room=%s",
                                    event["id"][:12], email
                                )
                                await client.delete(
                                    f"{GRAPH_BASE_URL}/users/{email}/events/{event['id']}",
                                    headers=headers
                                )
                    else:
                        logger.warning(
                            "Ghost Buster: calendarView returned %d for room=%s",
                            resp.status_code, email
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
    # FIX-19: delegates to _rooms_data() — single source of truth
    return {"value": _rooms_data()}


@app.post("/availability")
@limiter.limit("30/minute")
async def check_availability(request: Request, req: AvailabilityRequest):
    validate_email(req.room_email, "room_email")

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
        async with httpx.AsyncClient(timeout=HTTPX_TIMEOUT) as client:
            resp = await client.post(
                f"{GRAPH_BASE_URL}/users/{req.room_email}/calendar/getSchedule",
                headers=headers, json=payload
            )
    except httpx.TimeoutException:
        _timeout_error()
    return resp.json()


@app.get("/active-meeting")
@limiter.limit("60/minute")
async def get_active_meeting(request: Request, room_email: str):
    if not room_email or len(room_email) > 200:
        raise HTTPException(status_code=422, detail="Invalid room_email.")
    validate_email(room_email, "room_email")

    token = await get_app_token()
    now   = datetime.utcnow()

    past_start = (now - timedelta(minutes=60)).isoformat() + "Z"
    now_str    = now.isoformat() + "Z"

    url_active = (
        f"{GRAPH_BASE_URL}/users/{room_email}/calendarView"
        f"?startDateTime={past_start}&endDateTime={now_str}"
        f"&$select=id,subject,bodyPreview,categories,start,end,organizer,attendees"
        f"&$orderby=start/dateTime desc&$top=1"
    )
    try:
        async with httpx.AsyncClient(timeout=HTTPX_TIMEOUT) as client:
            resp = await client.get(url_active, headers={"Authorization": f"Bearer {token}"})
    except httpx.TimeoutException:
        _timeout_error()

    if resp.status_code == 200:
        active_events = resp.json().get("value", [])
        if active_events:
            event     = active_events[0]
            event_end = datetime.fromisoformat(event["end"]["dateTime"].replace("Z", ""))
            if event_end > now:
                return event

    future_end = (now + timedelta(hours=12)).isoformat() + "Z"
    url_future = (
        f"{GRAPH_BASE_URL}/users/{room_email}/calendarView"
        f"?startDateTime={now_str}&endDateTime={future_end}"
        f"&$select=id,subject,bodyPreview,categories,start,end,organizer,attendees"
        f"&$orderby=start/dateTime"
        f"&$top=5"
    )
    try:
        async with httpx.AsyncClient(timeout=HTTPX_TIMEOUT) as client:
            resp = await client.get(url_future, headers={"Authorization": f"Bearer {token}"})
    except httpx.TimeoutException:
        _timeout_error()

    if resp.status_code == 200:
        upcoming = resp.json().get("value", [])
        if upcoming:
            return upcoming

    return None


@app.post("/checkin")
@limiter.limit("30/minute")
async def check_in_meeting(request: Request, req: CheckInRequest):
    validate_email(req.room_email, "room_email")
    validate_event_id(req.event_id)

    token = await get_app_token()

    try:
        async with httpx.AsyncClient(timeout=HTTPX_TIMEOUT) as client:
            ev = await client.get(
                f"{GRAPH_BASE_URL}/users/{req.room_email}/events/{req.event_id}"
                f"?$select=start,end,categories",
                headers={"Authorization": f"Bearer {token}"}
            )
    except httpx.TimeoutException:
        _timeout_error()

    if ev.status_code == 404:
        raise HTTPException(status_code=404, detail="Event not found.")
    if ev.status_code != 200:
        raise HTTPException(status_code=422, detail="Could not retrieve event.")

    ev_data = ev.json()
    now     = datetime.utcnow()
    start   = datetime.fromisoformat(ev_data["start"]["dateTime"].replace("Z", ""))
    end     = datetime.fromisoformat(ev_data["end"]["dateTime"].replace("Z", ""))

    if now >= end:
        raise HTTPException(status_code=403, detail="Meeting has already ended.")

    window_open  = start - timedelta(minutes=1)
    window_close = start + timedelta(minutes=5)
    if not (window_open <= now <= window_close):
        raise HTTPException(
            status_code=403,
            detail="Check-in only allowed within 5 minutes of meeting start."
        )

    if "Checked-In" in ev_data.get("categories", []):
        raise HTTPException(status_code=409, detail="Already checked in.")

    try:
        async with httpx.AsyncClient(timeout=HTTPX_TIMEOUT) as client:
            resp = await client.patch(
                f"{GRAPH_BASE_URL}/users/{req.room_email}/events/{req.event_id}",
                headers={
                    "Authorization": f"Bearer {token}",
                    "Content-Type" : "application/json"
                },
                json={"categories": ["Checked-In"]}
            )
    except httpx.TimeoutException:
        _timeout_error()

    if resp.status_code == 404:
        raise HTTPException(status_code=404, detail="Event not found.")
    if resp.status_code not in (200, 201):
        raise HTTPException(status_code=422, detail="Check-in failed.")

    logger.info("Check-in successful: event=%s room=%s", req.event_id[:12], req.room_email)
    return {"status": "checked-in"}


@app.post("/extend-meeting")
@limiter.limit("30/minute")
async def extend_meeting(request: Request, req: ExtendRequest):
    validate_email(req.room_email, "room_email")
    validate_event_id(req.event_id)

    token = await get_app_token()

    try:
        async with httpx.AsyncClient(timeout=HTTPX_TIMEOUT) as client:
            ev = await client.get(
                f"{GRAPH_BASE_URL}/users/{req.room_email}/events/{req.event_id}"
                f"?$select=start,end,categories",
                headers={"Authorization": f"Bearer {token}"}
            )
    except httpx.TimeoutException:
        _timeout_error()

    if ev.status_code == 404:
        raise HTTPException(status_code=404, detail="Event not found.")
    if ev.status_code != 200:
        raise HTTPException(status_code=422, detail="Could not retrieve event.")

    ev_data = ev.json()
    now     = datetime.utcnow()
    start   = datetime.fromisoformat(ev_data["start"]["dateTime"].replace("Z", ""))
    end     = datetime.fromisoformat(ev_data["end"]["dateTime"].replace("Z", ""))

    if now < start:
        raise HTTPException(status_code=403, detail="Meeting has not started yet.")
    if now >= end:
        raise HTTPException(status_code=403, detail="Meeting has already ended.")

    if "Checked-In" not in ev_data.get("categories", []):
        raise HTTPException(
            status_code=403,
            detail="Meeting must be checked in before it can be extended."
        )

    new_end_dt = end + timedelta(minutes=req.extend_minutes)

    try:
        async with httpx.AsyncClient(timeout=HTTPX_TIMEOUT) as client:
            conflict_resp = await client.get(
                f"{GRAPH_BASE_URL}/users/{req.room_email}/calendarView"
                f"?startDateTime={end.isoformat()}Z&endDateTime={new_end_dt.isoformat()}Z"
                f"&$select=id,start,end&$top=5",
                headers={"Authorization": f"Bearer {token}"}
            )
    except httpx.TimeoutException:
        _timeout_error()

    conflicts = [
        e for e in conflict_resp.json().get("value", [])
        if e["id"] != req.event_id
    ]
    if conflicts:
        raise HTTPException(
            status_code=409,
            detail="Cannot extend — another meeting follows immediately."
        )

    try:
        async with httpx.AsyncClient(timeout=HTTPX_TIMEOUT) as client:
            resp = await client.patch(
                f"{GRAPH_BASE_URL}/users/{req.room_email}/events/{req.event_id}",
                headers={
                    "Authorization": f"Bearer {token}",
                    "Content-Type" : "application/json"
                },
                json={"end": {"dateTime": new_end_dt.isoformat() + "Z", "timeZone": "UTC"}}
            )
    except httpx.TimeoutException:
        _timeout_error()

    if resp.status_code != 200:
        raise HTTPException(status_code=resp.status_code, detail="Failed to extend meeting.")

    logger.info("Meeting extended: event=%s room=%s new_end=%s",
                req.event_id[:12], req.room_email, new_end_dt.isoformat())
    return {"status": "extended", "new_end": new_end_dt.isoformat()}


@app.post("/book")
@limiter.limit("10/minute")
async def create_booking(
    request    : Request,
    req        : BookingRequest,
    user_token : str = Depends(verify_user)
):
    validate_email(req.room_email,      "room_email")
    validate_email(req.organizer_email, "organizer_email")
    for att in req.attendees:
        validate_email(att.strip(), "attendee email")

    actual_email = await verify_token_and_get_email(user_token)

    system_token = await get_app_token()

    start_str = quote(req.start_time.replace(tzinfo=None).isoformat() + "Z")
    end_str   = quote(req.end_time.replace(tzinfo=None).isoformat()   + "Z")
    check_url = (
        f"{GRAPH_BASE_URL}/users/{req.room_email}/calendarView"
        f"?startDateTime={start_str}&endDateTime={end_str}&$select=subject"
    )
    try:
        async with httpx.AsyncClient(timeout=HTTPX_TIMEOUT) as client:
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
        "body"    : {
            "contentType": "Text",
            "content"    : f"Filiale: {req.filiale}\r\nReason: {req.description}"
        },
        "start"   : {"dateTime": req.start_time.replace(tzinfo=None).isoformat() + "Z", "timeZone": "UTC"},
        "end"     : {"dateTime": req.end_time.replace(tzinfo=None).isoformat()   + "Z", "timeZone": "UTC"},
        "location": {"displayName": "Conference Room", "locationEmailAddress": req.room_email},
        "attendees": all_attendees
    }

    try:
        async with httpx.AsyncClient(timeout=HTTPX_TIMEOUT) as client:
            resp = await client.post(
                f"{GRAPH_BASE_URL}/me/events",
                headers={"Authorization": f"Bearer {user_token}", "Content-Type": "application/json"},
                json=event_payload
            )
    except httpx.TimeoutException:
        _timeout_error()

    if resp.status_code != 201:
        raise HTTPException(status_code=resp.status_code, detail=f"Booking Failed: {resp.text}")

    logger.info("Booking created by %s for room %s", actual_email, req.room_email)
    return {"status": "success", "data": resp.json()}


@app.post("/end-meeting")
@limiter.limit("10/minute")
async def end_meeting(
    request    : Request,
    req        : CheckInRequest,
    user_token : str = Depends(verify_user)
):
    validate_email(req.room_email, "room_email")
    validate_event_id(req.event_id)

    actual_email = await verify_token_and_get_email(user_token)

    app_token = await get_app_token()
    try:
        async with httpx.AsyncClient(timeout=HTTPX_TIMEOUT) as client:
            ev = await client.get(
                f"{GRAPH_BASE_URL}/users/{req.room_email}/events/{req.event_id}"
                f"?$select=organizer,attendees",
                headers={"Authorization": f"Bearer {app_token}"}
            )
    except httpx.TimeoutException:
        _timeout_error()

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

    if actual_email not in allowed:
        logger.warning("Unauthorized end-meeting attempt by %s for event %s",
                       actual_email, req.event_id[:12])
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="You are not authorized to end this meeting."
        )

    now = datetime.utcnow().isoformat() + "Z"
    try:
        async with httpx.AsyncClient(timeout=HTTPX_TIMEOUT) as client:
            resp = await client.patch(
                f"{GRAPH_BASE_URL}/users/{req.room_email}/events/{req.event_id}",
                headers={"Authorization": f"Bearer {app_token}", "Content-Type": "application/json"},
                json={"end": {"dateTime": now, "timeZone": "UTC"}}
            )
    except httpx.TimeoutException:
        _timeout_error()

    if resp.status_code != 200:
        raise HTTPException(status_code=resp.status_code, detail="Failed to end meeting")

    logger.info("Meeting ended by %s: event=%s room=%s", actual_email, req.event_id[:12], req.room_email)
    return {"status": "ended"}
