import express from "express";
import axios from "axios";
import cors from "cors";

const app = express();
app.use(cors());
app.use(express.json());

const TENANT_ID = "YOUR_TENANT_ID";
const CLIENT_ID = "YOUR_CLIENT_ID";
const CLIENT_SECRET = "YOUR_CLIENT_SECRET";

async function getToken() {
  const res = await axios.post(
    `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,
    new URLSearchParams({
      client_id: CLIENT_ID,
      client_secret: CLIENT_SECRET,
      scope: "https://graph.microsoft.com/.default",
      grant_type: "client_credentials"
    })
  );
  return res.data.access_token;
}

app.get("/rooms", async (req, res) => {
  const token = await getToken();
  const rooms = await axios.get(
    "https://graph.microsoft.com/v1.0/places/microsoft.graph.room",
    { headers: { Authorization: `Bearer ${token}` } }
  );
  res.json(rooms.data.value);
});

app.post("/book", async (req, res) => {
  const { userEmail, roomEmail, start, end } = req.body;
  const token = await getToken();

  const meeting = await axios.post(
    `https://graph.microsoft.com/v1.0/users/${userEmail}/events`,
    {
      subject: "Room Reservation",
      start: { dateTime: start, timeZone: "UTC" },
      end: { dateTime: end, timeZone: "UTC" },
      attendees: [
        { emailAddress: { address: roomEmail }, type: "resource" }
      ],
      isOnlineMeeting: true,
      onlineMeetingProvider: "teamsForBusiness"
    },
    { headers: { Authorization: `Bearer ${token}` } }
  );

  res.json(meeting.data);
});

app.listen(3000, () => console.log("Backend running on 3000"));
