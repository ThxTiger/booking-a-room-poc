import os
import requests
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")

TOKEN_URL = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
GRAPH = "https://graph.microsoft.com/v1.0"


# ---------------------------
# Microsoft token
# ---------------------------

def get_token():
    r = requests.post(
        TOKEN_URL,
        data={
            "client_id": CLIENT_ID,
            "client_secret": CLIENT_SECRET,
            "scope": "https://graph.microsoft.com/.default",
            "grant_type": "client_credentials"
        },
    )

    r.raise_for_status()
    return r.json()["access_token"]


# ---------------------------
# List rooms
# ---------------------------

@app.get("/rooms")
def get_rooms():
    token = get_token()

    r = requests.get(
        f"{GRAPH}/places/microsoft.graph.room",
        headers={"Authorization": f"Bearer {token}"}
    )

    r.raise_for_status()
    return r.json()["value"]


# ---------------------------
# Room availability timeline
# ---------------------------

@app.post("/availability")
def availability(data: dict):
    token = get_token()

    body = {
        "schedules": [data["roomEmail"]],
        "startTime": {
            "dateTime": data["start"],
            "timeZone": "UTC"
        },
        "endTime": {
            "dateTime": data["end"],
            "timeZone": "UTC"
        },
        "availabilityViewInterval": 30
    }

    r = requests.post(
        f"{GRAPH}/me/calendar/getSchedule",
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        },
        json=body
    )

    r.raise_for_status()
    return r.json()


# ---------------------------
# Book room
# ---------------------------

@app.post("/book")
def book_room(data: dict):
    token = get_token()

    meeting = {
        "subject": "Room Reservation",
        "start": {
            "dateTime": data["start"],
            "timeZone": "UTC"
        },
        "end": {
            "dateTime": data["end"],
            "timeZone": "UTC"
        },
        "attendees": [
            {
                "emailAddress": {"address": data["roomEmail"]},
                "type": "resource"
            }
        ],
        "isOnlineMeeting": True,
        "onlineMeetingProvider": "teamsForBusiness"
    }

    r = requests.post(
        f"{GRAPH}/users/{data['userEmail']}/events",
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        },
        json=meeting
    )

    r.raise_for_status()
    return r.json()
