@app.post("/book")
async def create_booking(req: BookingRequest):
    """
    Creates a meeting, but ONLY if the room is free.
    """
    token = await get_graph_token()
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    
    # 1️⃣ STEP 1: CHECK AVAILABILITY FIRST (The "Guard")
    check_payload = {
        "schedules": [req.room_email],
        "startTime": {"dateTime": req.start_time.isoformat(), "timeZone": "UTC"},
        "endTime": {"dateTime": req.end_time.isoformat(), "timeZone": "UTC"},
        "availabilityViewInterval": 15
    }
    
    async with httpx.AsyncClient() as client:
        # Check schedule
        check_url = f"{GRAPH_BASE_URL}/users/{req.room_email}/calendar/getSchedule"
        check_resp = await client.post(check_url, headers=headers, json=check_payload)
        
        if check_resp.status_code == 200:
            data = check_resp.json()
            # The "availabilityView" string (e.g. "0000" is free, "0220" is mixed)
            # '0'=Free, '2'=Busy, '3'=Busy, '1'=Tentative
            view_str = data["value"][0]["availabilityView"]
            
            if "2" in view_str or "3" in view_str or "1" in view_str:
                # ⛔ STOP! Room is busy.
                raise HTTPException(status_code=409, detail="Sorry, this room is already reserved for that time.")

    # 2️⃣ STEP 2: IF FREE, PROCEED TO BOOK
    event_payload = {
        "subject": req.subject,
        "start": {"dateTime": req.start_time.isoformat(), "timeZone": "UTC"},
        "end": {"dateTime": req.end_time.isoformat(), "timeZone": "UTC"},
        "attendees": [{"emailAddress": {"address": req.organizer_email}, "type": "required"}]
    }

    async with httpx.AsyncClient() as client:
        url = f"{GRAPH_BASE_URL}/users/{req.room_email}/events"
        resp = await client.post(url, headers=headers, json=event_payload)
        
    if resp.status_code != 201:
        raise HTTPException(status_code=resp.status_code, detail=resp.json())
        
    return {"status": "success", "data": resp.json()}
