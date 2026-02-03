import os
import requests
from fastapi import FastAPI, HTTPException
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

GRAPH = "https://graph.microsoft.com/v1.0"
TOKEN_URL = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"


# -------------------------
# Get Microsoft token
# -------------------------

def get_token():
    r = requests.post(
        TOKEN_URL,
        data={
            "client_id": CLIENT_ID,
            "client_secret": CLIENT_SECRET,
            "scope": "https://graph.microsoft.com/.default",
            "grant_type": "client_credentials"
        }
    )

    if r.status_code != 200:
        raise HTTPException(status_code=500, detail=r.text)

    return r.json()["access_token"]


# -------------------------
# List rooms (CORRECT)
# -------------------------

@app.get("/rooms")
def get_rooms():
    token = get_token()

    r = requests.get(
        f"{GRAPH}/places/microsoft.graph.room",
        headers={"Authorization": f"Bearer {token}"}
    )

    if r.status_code != 200:
        return {"error": r.status_code, "details": r.text}

    return r.json().get("value", [])


# -------------------------
# Room availability
# -------------------------

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
        f"{GRAPH}/users/{data['roomEmail']}/calendar/getSchedule",
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        },
        json=body
    )

    if r.status_code != 200:
        return {"error": r.status_code, "details": r.text}

    return r.json()
