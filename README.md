# 🩸 Blood Life Line — Tech Titans Hackathon

> A smart blood-connect platform that links **patients, hospitals, donors and delivery volunteers** through a simple web workflow.

## 🚀 What was improved in this version

The original project was upgraded from a collection of forms/pages into a more interactive user journey.

### User-facing improvements
- **Interactive home dashboard** with live hospital, donor, blood-unit and request counts.
- **Live inventory search** by blood group and hospital/city/locality.
- **One-click blood request flow** from a selected hospital.
- **Booking ID tracking** directly from the home page.
- **Quick Assistant** that guides users to the right action.
- **Interactive blood-group guide** for basic educational compatibility information.
- **Multi-step donor registration** with progress indicator, validation and success screen.
- Responsive design for desktop and mobile.
- Clear loading, success and error states.
- Safer rendering of data received from the backend.
- Improved navigation and calls-to-action.
- Hospital, donor, inventory, delivery and emergency workflows remain available.

### Backend improvements
- Added `GET /api/health` for server health checking.
- Added `GET /api/stats` for dashboard statistics.
- Blood requests now reject a hospital that is not present in the inventory.
- Invalid blood types are rejected by the request endpoint.
- Existing Excel-based storage is preserved for easy hackathon demonstration.

---

## 🧩 Main features

| Feature | Page / API |
|---|---|
| Smart home dashboard | `bloodbridge/index.html` |
| Find blood | `bloodbridge/blood-availability.html` |
| Donor registration | `bloodbridge/registration.html` |
| Hospital portal | `bloodbridge/hospital-portal.html` |
| Hospital inventory | `bloodbridge/inventory.html` |
| Blood-bike registration | `bloodbridge/driver-registration.html` |
| Emergency contact | `bloodbridge/emergency-contact.html` |
| Inventory API | `GET /api/inventory-list` |
| Donor API | `GET /api/donors` |
| Blood request API | `POST /api/request-blood` |
| Booking tracking | `GET /api/blood-bookings/:bookingId` |
| Booking status | `POST /api/blood-bookings/status` |
| Dashboard stats | `GET /api/stats` |
| Server health | `GET /api/health` |

---

## 🏗️ Technology stack

**Frontend**
- HTML5
- CSS3
- Vanilla JavaScript
- Responsive CSS Grid/Flexbox
- LocalStorage for passing a selected hospital between pages

**Backend**
- Node.js
- Express.js
- CORS
- XLSX

**Data storage**
- Excel `.xlsx` files inside `bloodbridge/resources/`

No frontend framework is required.

---

## 📁 Project structure

```text
Hackathon_Tech_Titans/
├── README.md
├── bloodbridge/
│   ├── index.html
│   ├── registration.html
│   ├── blood-availability.html
│   ├── inventory.html
│   ├── hospital-portal.html
│   ├── driver-registration.html
│   ├── emergency-contact.html
│   ├── server.js
│   ├── package.json
│   └── resources/
│       ├── Login_Details.xlsx
│       ├── Inventory_Details.xlsx
│       ├── Blood_Requests.xlsx
│       ├── Blood_Bookings.xlsx
│       ├── Donor_Bookings.xlsx
│       └── Driver_Registrations.xlsx
└── resources/
    └── legacy/demo Excel files
```

`node_modules` is intentionally not included in the updated submission archive. It should be installed with `npm install`.

---


## ⚠️ npm installation note

The `bloodbridge/package-lock.json` in this submission uses the public npm registry:
`https://registry.npmjs.org/`

Run npm from the **bloodbridge** folder:

```powershell
cd bloodbridge
npm install
npm start
```

Do **not** run `npm install` from the outer `Hackathon_Tech_Titans_Interactive_v2` folder because the application's `package.json` is inside `bloodbridge`.

If npm still reports `ENOTFOUND` or references `npme.walmart.com`, check:

```powershell
npm config get registry
```

It should return:

```text
https://registry.npmjs.org/
```


## ⚙️ How to run

### 1. Install Node.js

Install a current LTS version of Node.js.

Check installation:

```bash
node -v
npm -v
```

### 2. Open the project

```bash
cd bloodbridge
```

### 3. Install dependencies

```bash
npm install
```

### 4. Start the server

```bash
npm start
```

You should see:

```text
Server: http://localhost:3000
```

### 5. Open the application

Open:

```text
http://localhost:3000
```

**Important:** Do not open `index.html` directly with `file://`. The application uses backend APIs, so it should be opened through the Node/Express server.

---

## 🧪 Suggested hackathon demo flow

### Demo 1 — Patient requests blood

1. Open the home page.
2. Show the live dashboard counters.
3. Scroll to **Find available blood**.
4. Select a blood group such as `O+`.
5. Search for a hospital.
6. Click **Request from this hospital**.
7. Complete the request form.
8. Submit the request.
9. Note the generated **Booking ID**.
10. Return to Home.
11. Click **Track Request**.
12. Enter the Booking ID.
13. Demonstrate the current request status.

### Demo 2 — Donor registration

1. Click **Become a Donor**.
2. Enter basic details.
3. Select Blood or Organ donation.
4. Continue through the steps.
5. Review the confirmation summary.
6. Submit.
7. Show the success screen.
8. Open the Hospital Portal to demonstrate that the donor is available in the backend data.

### Demo 3 — Hospital workflow

1. Open Hospital Portal.
2. View donor information.
3. View blood inventory.
4. Update hospital inventory.
5. Refresh the home page.
6. Show the updated availability.

---

## 🔌 API reference

### `GET /api/health`

Checks whether the backend is running.

Example response:

```json
{
  "success": true,
  "service": "Blood Life Line",
  "status": "online"
}
```

### `GET /api/stats`

Returns dashboard totals.

Example:

```json
{
  "success": true,
  "hospitals": 4,
  "donors": 5,
  "requests": 4,
  "bookings": 0,
  "bloodUnits": 120
}
```

### `GET /api/inventory-list`

Returns hospital inventory.

### `GET /api/donors`

Returns registered donor records.

### `POST /api/register`

Stores a donor registration.

Expected fields:

```json
{
  "firstName": "Example",
  "lastName": "User",
  "phone": "9876543210",
  "email": "example@email.com",
  "donorType": "blood",
  "bloodGroup": "O+",
  "organType": "",
  "address": "Chennai"
}
```

### `POST /api/request-blood`

Creates a blood request and generates:
- Booking ID
- Pickup OTP
- Request status
- Optional delivery-driver assignment

The request also reduces the selected hospital's inventory when sufficient stock is available.

### `GET /api/blood-bookings/:bookingId`

Retrieves one booking by its Booking ID.

### `POST /api/blood-bookings/status`

Updates a booking status. Completing a booking requires the generated pickup OTP.

---

## 📊 Excel data model

### Inventory

The inventory workbook contains:

```text
Update Date
Hospital Name
Tech_Titans_Code
City
Locality
Contact
A+
A-
B+
B-
O+
O-
AB+
AB-
```

### Donors

```text
Date
First Name
Last Name
Phone
Email
Type
Blood Group
Organ Type
Address
```

### Blood requests

```text
Request Date
Source Hospital
Source City
Source Locality
Blood Type
Units
Delivery Method
Requester Name
Requester Contact
Requester Hospital
Doctor Name
Address
Cause
Driver Name
Driver Phone
Status
```

### Blood bookings

```text
Booking ID
Created At
Hospital Name
City
Locality
Hospital Contact
Blood Type
Units
Delivery Method
Requester Name
Requester Contact
Requester Hospital
Doctor Name
Requester Address
Cause
Driver Name
Driver Phone
Pickup OTP
Status
```

---

## 🔐 Demo access

The existing hospital entry flow uses the hackathon demo code:

```text
Tech_Titans
```

This is **not production authentication**. It is only intended for the hackathon demonstration.

---

## 🛠️ Troubleshooting

### `npm` or `node` is not recognized

Install Node.js and reopen the terminal/VS Code.

### Page loads but live data does not appear

Make sure the server is running:

```bash
npm start
```

Then open:

```text
http://localhost:3000
```

### `EADDRINUSE: address already in use`

Another process is already using port `3000`.

Stop the existing Node process or change the `PORT` value in `server.js`.

### Inventory does not update

Check that:
- The hospital exists in `bloodbridge/resources/Inventory_Details.xlsx`.
- The blood group has enough units.
- The Node process has permission to write to the resources folder.

### Excel data looks unchanged

Refresh the page after the request/update. The backend writes changes to the Excel files.

---

## ⚠️ Important project limitation

This is a **hackathon prototype**, not a production blood-bank system.

For real deployment, it would need:
- Secure authentication and role-based access
- Database instead of Excel files
- HTTPS
- Input validation and rate limiting
- Proper audit logs
- Verified hospital/donor identities
- Secure handling of personal information
- Real notification/SMS/email integrations
- Medical and blood-bank compliance workflows
- Clinician/blood-bank compatibility validation
- Reliable location and delivery tracking

The blood-group guide in the UI is educational only. Actual transfusion decisions must follow qualified medical professionals and blood-bank testing/protocols.

---

## 💡 Future scope

- 📱 Progressive Web App / mobile application
- 🔔 SMS/WhatsApp/email alerts
- 🗺️ Real-time delivery tracking
- 🤖 AI-assisted request prioritization
- 📍 Location-based hospital discovery
- 📈 Hospital analytics dashboard
- 🏥 Verified hospital accounts
- 🔐 Role-based authentication
- 🗄️ MySQL/PostgreSQL/MongoDB database
- ☁️ Cloud deployment
- 📦 Automated inventory alerts when stock becomes low
- 🧑‍🤝‍🧑 Donor availability notifications

---

## 👥 Hackathon presentation idea

### Problem
During urgent blood requirements, patients and hospitals can lose valuable time searching for compatible blood availability and coordinating delivery.

### Solution
Blood Life Line creates a connected workflow where users can:

**Search → Request → Connect → Track**

### Key differentiator
Instead of presenting only a static blood inventory, the prototype connects **inventory + request creation + booking ID + delivery assignment + tracking + donor registration** into one user journey.

---

## 📌 Project status

**Version:** 2.0 — Interactive Hackathon Upgrade  
**Project:** Blood Life Line  
**Team:** Tech Titans  
**Purpose:** Hackathon prototype / academic demonstration
