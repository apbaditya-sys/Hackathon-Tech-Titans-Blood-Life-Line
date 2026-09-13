const express = require('express');
const cors = require('cors');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = 3000;

app.use(cors());
app.use(express.json({ limit: '10mb' })); // Allow larger payloads for photo uploads
app.use(express.static('.')); 

const RESOURCES_DIR = path.join(__dirname, 'resources');
const DONOR_FILE = path.join(RESOURCES_DIR, 'Login_Details.xlsx');
const INVENTORY_FILE = path.join(RESOURCES_DIR, 'Inventory_Details.xlsx');
const BOOKINGS_FILE = path.join(RESOURCES_DIR, 'Donor_Bookings.xlsx');
const BLOOD_REQUESTS_FILE = path.join(RESOURCES_DIR, 'Blood_Requests.xlsx');
const BLOOD_BOOKINGS_FILE = path.join(RESOURCES_DIR, 'Blood_Bookings.xlsx');
const DRIVER_FILE = path.join(RESOURCES_DIR, 'Driver_Registrations.xlsx');

if (!fs.existsSync(RESOURCES_DIR)) fs.mkdirSync(RESOURCES_DIR);

function initFiles() {
    // Donor File
    if (!fs.existsSync(DONOR_FILE)) {
        const wb = XLSX.utils.book_new();
        const headers = [['Date', 'First Name', 'Last Name', 'Phone', 'Email', 'Type', 'Blood Group', 'Organ Type', 'Address']];
        XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(headers), 'Donors');
        XLSX.writeFile(wb, DONOR_FILE);
    }
    // Inventory File
    if (!fs.existsSync(INVENTORY_FILE)) {
        const wb = XLSX.utils.book_new();
        const headers = [['Update Date', 'Hospital Name', 'Tech_Titans_Code', 'City', 'Locality', 'Contact', 'A+', 'A-', 'B+', 'B-', 'O+', 'O-', 'AB+', 'AB-']];
        XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(headers), 'Inventory');
        XLSX.writeFile(wb, INVENTORY_FILE);
    }
    // Bookings File
    if (!fs.existsSync(BOOKINGS_FILE)) {
        const wb = XLSX.utils.book_new();
        const headers = [['Booking Date', 'Donor Name', 'Receiver Name', 'Hospital', 'Status']];
        XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(headers), 'Bookings');
        XLSX.writeFile(wb, BOOKINGS_FILE);
    }
    // Blood Requests File
    if (!fs.existsSync(BLOOD_REQUESTS_FILE)) {
        const wb = XLSX.utils.book_new();
        const headers = [['Request Date', 'Source Hospital', 'Source City', 'Source Locality', 'Blood Type', 'Units', 'Delivery Method', 'Requester Name', 'Requester Contact', 'Requester Hospital', 'Doctor Name', 'Address', 'Cause', 'Driver Name', 'Driver Phone', 'Status']];
        XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(headers), 'Requests');
        XLSX.writeFile(wb, BLOOD_REQUESTS_FILE);
    }

    // Blood Bookings File (for user-side booking tracking)
    if (!fs.existsSync(BLOOD_BOOKINGS_FILE)) {
        const wb = XLSX.utils.book_new();
        const headers = [[
            'Booking ID',
            'Created At',
            'Hospital Name',
            'City',
            'Locality',
            'Hospital Contact',
            'Blood Type',
            'Units',
            'Delivery Method',
            'Requester Name',
            'Requester Contact',
            'Requester Hospital',
            'Doctor Name',
            'Requester Address',
            'Cause',
            'Driver Name',
            'Driver Phone',
            'Pickup OTP',
            'Status'
        ]];
        XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(headers), 'Bookings');
        XLSX.writeFile(wb, BLOOD_BOOKINGS_FILE);
    }
    // Driver Registrations File
    if (!fs.existsSync(DRIVER_FILE)) {
        const wb = XLSX.utils.book_new();
        const headers = [['Registration Date', 'First Name', 'Last Name', 'Phone', 'Age', 'License Number', 'Address', 'Photo Name']];
        XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(headers), 'Drivers');
        XLSX.writeFile(wb, DRIVER_FILE);
    }
}
initFiles();

// Helpers
function generateBookingId() {
    // Example: BLD-1700000000000-482
    return `BLD-${Date.now()}-${Math.floor(100 + Math.random() * 900)}`;
}

function generateOtp() {
    // 6-digit OTP
    return String(Math.floor(100000 + Math.random() * 900000));
}

function safeReadSheetAsAoA(filePath, sheetName) {
    const workbook = XLSX.readFile(filePath);
    const sheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    return { workbook, data };
}

// Routes
app.post('/api/register', (req, res) => {
    try {
        const f = req.body;
        const workbook = XLSX.readFile(DONOR_FILE);
        const data = XLSX.utils.sheet_to_json(workbook.Sheets['Donors'], { header: 1 });
        data.push([new Date().toLocaleString(), f.firstName, f.lastName, f.phone, f.email, f.donorType, f.bloodGroup, f.organType, f.address]);
        workbook.Sheets['Donors'] = XLSX.utils.aoa_to_sheet(data);
        XLSX.writeFile(workbook, DONOR_FILE);
        res.json({ success: true });
    } catch (e) { res.status(500).json({ success: false }); }
});

app.post('/api/inventory', (req, res) => {
    try {
        const f = req.body;
        const workbook = XLSX.readFile(INVENTORY_FILE);
        const data = XLSX.utils.sheet_to_json(workbook.Sheets['Inventory'], { header: 1 });
        
        // Check if this is an update or new registration
        if (f.mode === 'update') {
            // Find the hospital by name and update it
            let hospitalFound = false;
            for (let i = 1; i < data.length; i++) { // Start from 1 to skip headers
                if (data[i][1] === f.hospitalName) { // Column 1 is Hospital Name
                    // Update the existing row
                    data[i] = [
                        new Date().toLocaleString(), // Update Date
                        f.hospitalName,
                        f.techTitansCode,
                        f.city,
                        f.locality,
                        f.contactNumber,
                        f.inventory.a_plus,
                        f.inventory.a_minus,
                        f.inventory.b_plus,
                        f.inventory.b_minus,
                        f.inventory.o_plus,
                        f.inventory.o_minus,
                        f.inventory.ab_plus,
                        f.inventory.ab_minus
                    ];
                    hospitalFound = true;
                    break;
                }
            }
            
            if (!hospitalFound) {
                return res.status(404).json({ success: false, message: 'Hospital not found' });
            }
        } else {
            // Register mode - add new hospital
            // Check if hospital already exists
            for (let i = 1; i < data.length; i++) {
                if (data[i][1] === f.hospitalName) {
                    return res.status(400).json({ success: false, message: 'Hospital already exists. Please use update mode.' });
                }
            }
            
            // Add new row
            data.push([
                new Date().toLocaleString(),
                f.hospitalName,
                f.techTitansCode,
                f.city,
                f.locality,
                f.contactNumber,
                f.inventory.a_plus || 0,
                f.inventory.a_minus || 0,
                f.inventory.b_plus || 0,
                f.inventory.b_minus || 0,
                f.inventory.o_plus || 0,
                f.inventory.o_minus || 0,
                f.inventory.ab_plus || 0,
                f.inventory.ab_minus || 0
            ]);
        }
        
        workbook.Sheets['Inventory'] = XLSX.utils.aoa_to_sheet(data);
        XLSX.writeFile(workbook, INVENTORY_FILE);
        res.json({ success: true });
    } catch (e) { 
        console.error('Error in /api/inventory:', e);
        res.status(500).json({ success: false, message: e.message }); 
    }
});

app.get('/api/inventory-list', (req, res) => {
    try {
        const workbook = XLSX.readFile(INVENTORY_FILE);
        res.json({ success: true, inventory: XLSX.utils.sheet_to_json(workbook.Sheets['Inventory']) });
    } catch (e) { res.json({ success: false, inventory: [] }); }
});

app.get('/api/donors', (req, res) => {
    try {
        const workbook = XLSX.readFile(DONOR_FILE);
        res.json({ success: true, donors: XLSX.utils.sheet_to_json(workbook.Sheets['Donors']) });
    } catch (e) { res.json({ success: false, donors: [] }); }
});

app.post('/api/book-donor', (req, res) => {
    try {
        const f = req.body;
        const workbook = XLSX.readFile(BOOKINGS_FILE);
        const data = XLSX.utils.sheet_to_json(workbook.Sheets['Bookings'], { header: 1 });
        data.push([new Date().toLocaleString(), f.donorName, f.receiverName, f.hospital, 'Paid Successfully']);
        workbook.Sheets['Bookings'] = XLSX.utils.aoa_to_sheet(data);
        XLSX.writeFile(workbook, BOOKINGS_FILE);
        res.json({ success: true });
    } catch (e) { res.status(500).json({ success: false }); }
});

app.post('/api/request-blood', (req, res) => {
    try {
        const f = req.body;
        console.log('Received blood request data:', f); // Debug log

        // Create a booking reference for the user
        const bookingId = generateBookingId();
        const pickupOtp = generateOtp();
        
        // Blood type mapping
        const typeMapping = {
            'a_plus': 'A+', 'a_minus': 'A-',
            'b_plus': 'B+', 'b_minus': 'B-',
            'o_plus': 'O+', 'o_minus': 'O-',
            'ab_plus': 'AB+', 'ab_minus': 'AB-'
        };
        
        // Update inventory - subtract units from source hospital
        const inventoryWorkbook = XLSX.readFile(INVENTORY_FILE);
        const inventoryData = XLSX.utils.sheet_to_json(inventoryWorkbook.Sheets['Inventory']);
        
        const hospital = inventoryData.find(h => 
            h['Hospital Name'] === f.sourceHospital &&
            h.City === f.sourceCity &&
            h.Locality === f.sourceLocality
        );
        
        if (!hospital) {
            return res.status(404).json({
                success: false,
                message: 'Selected hospital was not found in the current inventory.'
            });
        }

        const bloodColumn = typeMapping[f.bloodType] || f.bloodType;
        if (!typeMapping[f.bloodType]) {
            return res.status(400).json({ success: false, message: 'Invalid blood type.' });
        }

        if (hospital) {
            const bloodColumn = typeMapping[f.bloodType] || f.bloodType;
            const currentUnits = parseInt(hospital[bloodColumn] || 0);
            const requestedUnits = parseInt(f.units);
            
            if (currentUnits >= requestedUnits) {
                hospital[bloodColumn] = currentUnits - requestedUnits;
                
                // Write updated inventory back to file
                const inventorySheet = XLSX.utils.json_to_sheet(inventoryData);
                inventoryWorkbook.Sheets['Inventory'] = inventorySheet;
                XLSX.writeFile(inventoryWorkbook, INVENTORY_FILE);
                console.log(`Updated inventory: ${f.sourceHospital} ${bloodColumn} reduced by ${requestedUnits} units`);
            } else {
                return res.status(400).json({ 
                    success: false, 
                    message: `Insufficient units. Only ${currentUnits} units available.` 
                });
            }
        }
        
        // Assign driver if delivery method is 'delivery'
        let driverName = 'N/A';
        let driverPhone = 'N/A';
        
        if (f.deliveryMethod === 'delivery') {
            try {
                const driverWorkbook = XLSX.readFile(DRIVER_FILE);
                const drivers = XLSX.utils.sheet_to_json(driverWorkbook.Sheets['Drivers']);
                
                if (drivers && drivers.length > 0) {
                    // Assign drivers in round-robin fashion based on current number of requests
                    const requestWorkbook = XLSX.readFile(BLOOD_REQUESTS_FILE);
                    const requests = XLSX.utils.sheet_to_json(requestWorkbook.Sheets['Requests']);
                    const driverIndex = requests.length % drivers.length;
                    const assignedDriver = drivers[driverIndex];
                    
                    driverName = `${assignedDriver['First Name']} ${assignedDriver['Last Name']}`;
                    driverPhone = assignedDriver['Phone'];
                    console.log(`Assigned driver: ${driverName} (${driverPhone})`);
                }
            } catch (driverError) {
                console.log('No drivers available for assignment:', driverError.message);
            }
        }
        
        // Save blood request
        const workbook = XLSX.readFile(BLOOD_REQUESTS_FILE);
        const data = XLSX.utils.sheet_to_json(workbook.Sheets['Requests'], { header: 1 });
        
        const newRow = [
            new Date().toLocaleString(),
            f.sourceHospital,
            f.sourceCity,
            f.sourceLocality,
            f.bloodType,
            f.units,
            f.deliveryMethod,
            f.requesterName,
            f.requesterContact,
            f.requesterHospital,
            f.doctorName,
            f.requesterAddress,
            f.cause,
            driverName,
            driverPhone,
            'Pending'
        ];
        
        console.log('New row to be added:', newRow); // Debug log
        data.push(newRow);
        workbook.Sheets['Requests'] = XLSX.utils.aoa_to_sheet(data);
        XLSX.writeFile(workbook, BLOOD_REQUESTS_FILE);

        // Also store a user-visible booking record
        const bookingWb = XLSX.readFile(BLOOD_BOOKINGS_FILE);
        const bookingData = XLSX.utils.sheet_to_json(bookingWb.Sheets['Bookings'], { header: 1 });
        bookingData.push([
            bookingId,
            new Date().toLocaleString(),
            f.sourceHospital,
            f.sourceCity,
            f.sourceLocality,
            f.sourceContact || 'N/A',
            f.bloodType,
            f.units,
            f.deliveryMethod,
            f.requesterName,
            f.requesterContact,
            f.requesterHospital,
            f.doctorName,
            f.requesterAddress,
            f.cause,
            driverName,
            driverPhone,
            pickupOtp,
            'Pending'
        ]);
        bookingWb.Sheets['Bookings'] = XLSX.utils.aoa_to_sheet(bookingData);
        XLSX.writeFile(bookingWb, BLOOD_BOOKINGS_FILE);

        res.json({
            success: true,
            bookingId,
            pickupOtp,
            driverAssigned: driverName,
            status: 'Pending'
        });
    } catch (e) { 
        console.error('Error in /api/request-blood:', e);
        res.status(500).json({ success: false, message: e.message }); 
    }
});

app.get('/api/blood-requests', (req, res) => {
    try {
        const workbook = XLSX.readFile(BLOOD_REQUESTS_FILE);
        const requests = XLSX.utils.sheet_to_json(workbook.Sheets['Requests']);
        console.log('All requests:', requests); // Debug log
        
        // Temporarily show ALL requests to debug
        res.json({ success: true, requests: requests });
        
        // Original filter (commented out for debugging)
        // const deliveryRequests = requests.filter(req => {
        //     const deliveryMethod = req['Delivery Method'];
        //     console.log('Delivery Method:', deliveryMethod); // Debug log
        //     return deliveryMethod && deliveryMethod.toLowerCase() === 'delivery';
        // });
        // console.log('Filtered delivery requests:', deliveryRequests); // Debug log
        // res.json({ success: true, requests: deliveryRequests });
    } catch (e) { 
        console.error('Error in /api/blood-requests:', e);
        res.json({ success: false, requests: [], error: e.message }); 
    }
});

// Blood Booking APIs (simple tracking + status updates)
app.get('/api/blood-bookings', (req, res) => {
    try {
        const wb = XLSX.readFile(BLOOD_BOOKINGS_FILE);
        const bookings = XLSX.utils.sheet_to_json(wb.Sheets['Bookings']);
        res.json({ success: true, bookings });
    } catch (e) {
        res.json({ success: false, bookings: [], error: e.message });
    }
});

app.get('/api/blood-bookings/:bookingId', (req, res) => {
    try {
        const { bookingId } = req.params;
        const wb = XLSX.readFile(BLOOD_BOOKINGS_FILE);
        const bookings = XLSX.utils.sheet_to_json(wb.Sheets['Bookings']);
        const booking = bookings.find(b => String(b['Booking ID']) === String(bookingId));
        if (!booking) return res.status(404).json({ success: false, message: 'Booking not found' });
        res.json({ success: true, booking });
    } catch (e) {
        res.status(500).json({ success: false, message: e.message });
    }
});

app.post('/api/blood-bookings/status', (req, res) => {
    try {
        const { bookingId, status, otp } = req.body;
        if (!bookingId || !status) {
            return res.status(400).json({ success: false, message: 'bookingId and status are required' });
        }

        const wb = XLSX.readFile(BLOOD_BOOKINGS_FILE);
        const aoa = XLSX.utils.sheet_to_json(wb.Sheets['Bookings'], { header: 1 });
        if (!aoa || aoa.length < 2) {
            return res.status(404).json({ success: false, message: 'No bookings found' });
        }

        // Find row by Booking ID (col 0)
        let updated = false;
        for (let i = 1; i < aoa.length; i++) {
            if (String(aoa[i][0]) === String(bookingId)) {
                const expectedOtp = String(aoa[i][17] || '');

                // If marking as Completed, OTP is required (for pickup handover)
                if (String(status).toLowerCase() === 'completed') {
                    if (!otp || String(otp) !== expectedOtp) {
                        return res.status(400).json({ success: false, message: 'Invalid OTP' });
                    }
                }

                aoa[i][18] = status; // Status column
                updated = true;
                break;
            }
        }

        if (!updated) return res.status(404).json({ success: false, message: 'Booking not found' });

        wb.Sheets['Bookings'] = XLSX.utils.aoa_to_sheet(aoa);
        XLSX.writeFile(wb, BLOOD_BOOKINGS_FILE);
        res.json({ success: true });
    } catch (e) {
        res.status(500).json({ success: false, message: e.message });
    }
});

app.post('/api/register-driver', (req, res) => {
    try {
        const f = req.body;
        const workbook = XLSX.readFile(DRIVER_FILE);
        const data = XLSX.utils.sheet_to_json(workbook.Sheets['Drivers'], { header: 1 });
        
        const newRow = [
            new Date().toLocaleString(),
            f.firstName,
            f.lastName,
            f.phone,
            f.age,
            f.licenseNumber,
            f.address,
            f.photoName || ''
        ];
        
        data.push(newRow);
        workbook.Sheets['Drivers'] = XLSX.utils.aoa_to_sheet(data);
        XLSX.writeFile(workbook, DRIVER_FILE);
        res.json({ success: true });
    } catch (e) { 
        console.error('Error in /api/register-driver:', e);
        res.status(500).json({ success: false, message: e.message }); 
    }
});


// Health + dashboard summary APIs
app.get('/api/health', (req, res) => {
    res.json({ success: true, service: 'Blood Life Line', status: 'online', time: new Date().toISOString() });
});

app.get('/api/stats', (req, res) => {
    try {
        const inventory = XLSX.utils.sheet_to_json(XLSX.readFile(INVENTORY_FILE).Sheets['Inventory']);
        const donors = XLSX.utils.sheet_to_json(XLSX.readFile(DONOR_FILE).Sheets['Donors']);
        const requests = XLSX.utils.sheet_to_json(XLSX.readFile(BLOOD_REQUESTS_FILE).Sheets['Requests']);
        const bookings = XLSX.utils.sheet_to_json(XLSX.readFile(BLOOD_BOOKINGS_FILE).Sheets['Bookings']);
        const groups = ['A+','A-','B+','B-','O+','O-','AB+','AB-'];
        const units = inventory.reduce((sum, h) => sum + groups.reduce((x, g) => x + (parseInt(h[g]) || 0), 0), 0);
        res.json({
            success: true,
            hospitals: inventory.length,
            donors: donors.length,
            requests: requests.length,
            bookings: bookings.length,
            bloodUnits: units,
            generatedAt: new Date().toISOString()
        });
    } catch (e) {
        res.status(500).json({ success: false, message: e.message });
    }
});

app.listen(PORT, () => console.log(`Server: http://localhost:${PORT}`));