// 🔴 REPLACE THIS WITH YOUR RENDER URL
const API_URL = "https://booking-a-room-poc.onrender.com"; 

// --- Initialization ---
document.addEventListener("DOMContentLoaded", () => {
    initModalTimes();
    fetchRooms();
});

// Set default times in the modal (Next 30 min block)
function initModalTimes() {
    const now = new Date();
    now.setMinutes(now.getMinutes() - now.getTimezoneOffset()); // Adjust to local string
    document.getElementById('startTime').value = now.toISOString().slice(0,16);
    
    now.setMinutes(now.getMinutes() + 30);
    document.getElementById('endTime').value = now.toISOString().slice(0,16);
}

// --- 1. Fetch Rooms ---
async function fetchRooms() {
    try {
        const res = await fetch(`${API_URL}/rooms`);
        const data = await res.json();
        const select = document.getElementById('roomSelect');
        
        select.innerHTML = '<option value="" disabled selected>Select a room...</option>';
        
        // Handle the "value" list (works for both Static and Graph responses)
        if (data.value && data.value.length > 0) {
            data.value.forEach(room => {
                const opt = document.createElement('option');
                opt.value = room.emailAddress;
                opt.textContent = room.displayName;
                select.appendChild(opt);
            });
        } else {
            select.innerHTML = '<option disabled>No rooms found</option>';
        }
    } catch (err) {
        console.error("Error fetching rooms:", err);
        alert("Backend connection failed. Is the Render server running?");
    }
}

// --- 2. Load Availability (Timeline) ---
async function loadAvailability() {
    const roomEmail = document.getElementById('roomSelect').value;
    if (!roomEmail) return;

    document.getElementById('loadingSpinner').style.display = "inline";
    const timeline = document.getElementById('timeline');
    
    // Range: Now to +8 Hours
    const start = new Date();
    const end = new Date(start.getTime() + 8 * 60 * 60 * 1000);

    try {
        const res = await fetch(`${API_URL}/availability`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                room_email: roomEmail,
                start_time: start.toISOString(),
                end_time: end.toISOString(),
                time_zone: Intl.DateTimeFormat().resolvedOptions().timeZone
            })
        });

        const data = await res.json();
        renderTimeline(data, start, end);
    } catch (err) {
        timeline.innerHTML = `<p class="text-danger p-3">Error loading data: ${err.message}</p>`;
    } finally {
        document.getElementById('loadingSpinner').style.display = "none";
    }
}

// --- 3. Render Visual Timeline ---
function renderTimeline(data, startObj, endObj) {
    const timeline = document.getElementById('timeline');
    let html = `<div class="timeline-wrapper">`;
    
    // Draw 8 Hour Markers
    for (let i=0; i<8; i++) {
        let hour = new Date(startObj.getTime() + i * 60 * 60 * 1000);
        let timeStr = hour.toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'});
        html += `<div class="hour-slot">${timeStr}</div>`;
    }

    // Draw Busy Blocks
    const totalMs = endObj - startObj;
    // Graph returns data inside: value[0].scheduleItems
    const schedule = (data.value && data.value[0]) ? data.value[0] : null; 
    
    if (schedule && schedule.scheduleItems) {
        schedule.scheduleItems.forEach(item => {
            if (item.status === 'busy') {
                const eventStart = new Date(item.start.dateTime + 'Z'); // Assume UTC
                const eventEnd = new Date(item.end.dateTime + 'Z');

                // Math to position the red block
                const offsetMs = eventStart - startObj;
                const durationMs = eventEnd - eventStart;
                
                // Only draw if inside the visible 8-hour window
                if (offsetMs < totalMs && (offsetMs + durationMs) > 0) {
                    const leftPct = Math.max(0, (offsetMs / totalMs) * 100);
                    const widthPct = Math.min(100 - leftPct, (durationMs / totalMs) * 100);
                    
                    html += `<div class="event-block" 
                                  style="left:${leftPct}%; width:${widthPct}%;" 
                                  title="${item.subject || 'Busy'}">
                                  ${item.subject || 'Busy'}
                             </div>`;
                }
            }
        });
    }
    
    html += `</div>`;
    timeline.innerHTML = html;
}

// --- 4. Create Booking ---
async function createBooking() {
    const roomEmail = document.getElementById('roomSelect').value;
    const organizer = document.getElementById('organizerEmail').value;
    const start = new Date(document.getElementById('startTime').value);
    const end = new Date(document.getElementById('endTime').value);
    const subject = document.getElementById('subject').value;

    if (!roomEmail) return alert("Please select a room first.");
    if (!organizer) return alert("Please enter your email.");

    try {
        const res = await fetch(`${API_URL}/book`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                subject: subject,
                room_email: roomEmail,
                start_time: start.toISOString(),
                end_time: end.toISOString(),
                organizer_email: organizer
            })
        });
        
        if (res.ok) {
            alert("✅ Booking Successful! Check your Outlook.");
            // Close Modal
            const modalEl = document.getElementById('bookingModal');
            const modal = bootstrap.Modal.getInstance(modalEl);
            modal.hide();
            // Refresh Timeline
            loadAvailability(); 
        } else {
            const err = await res.json();
            alert("❌ Booking Failed: " + JSON.stringify(err));
        }
    } catch (e) {
        alert("Error: " + e.message);
    }
}
