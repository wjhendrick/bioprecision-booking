/**
 * BioPrecision Booking System — Backend Server
 * ─────────────────────────────────────────────
 * Handles:
 *   - Square payment processing
 *   - Monday.com New Event Request auto-submission
 *   - BioPrecision Excel roster update + email
 *   - Booking confirmation emails to client and owner
 *
 * SETUP:
 *   1. npm install express square dotenv nodemailer cors axios xlsx
 *   2. Create a .env file (see .env.example below)
 *   3. node server.js
 *
 * ⚠️  IMPORTANT: Regenerate your Square Access Token after setup.
 *     Never commit .env to GitHub.
 */

require('dotenv').config();
const express    = require('express');
const cors       = require('cors');
const { SquareClient, SquareEnvironment } = require('square');
const { Resend } = require('resend');
const resend = new Resend(process.env.RESEND_API_KEY);
const axios      = require('axios');
const XLSX       = require('xlsx');
const crypto     = require('crypto');
const path       = require('path');
const fs         = require('fs');

const app = express();
app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// ─────────────────────────────────────────────────────────────────────────────
// SQUARE CLIENT
// ─────────────────────────────────────────────────────────────────────────────
const squareClient = new SquareClient({
  environment: SquareEnvironment.Production,
  token: process.env.SQUARE_ACCESS_TOKEN,
});

// ─────────────────────────────────────────────────────────────────────────────
// EMAIL TRANSPORTER (Nodemailer)
// ─────────────────────────────────────────────────────────────────────────────
// Email sending via Resend
async function sendEmail({ from, to, subject, html }) {
  const { data, error } = await resend.emails.send({ from, to, subject, html });
  if (error) throw new Error(error.message);
  return data;
}

// ─────────────────────────────────────────────────────────────────────────────
// PATHS
// ─────────────────────────────────────────────────────────────────────────────
const ROSTER_PATH = path.join(__dirname, 'BioPrecision_Roster.xlsx');
const AVAILABILITY_PATH = path.join(__dirname, 'availability.json');

// ─────────────────────────────────────────────────────────────────────────────
// AVAILABILITY STORE
// Persists blocked slots and closed dates to a JSON file on the server.
// Structure: { closedDates: [...], blockedSlots: { 'YYYY-MM-DD': { 'HH:MM AM': { hit: bool, pitch: bool } } } }
// ─────────────────────────────────────────────────────────────────────────────
function loadAvailability() {
  try {
    if (fs.existsSync(AVAILABILITY_PATH)) {
      return JSON.parse(fs.readFileSync(AVAILABILITY_PATH, 'utf8'));
    }
  } catch(e) { console.error('Error loading availability:', e.message); }
  return { closedDates: [], blockedSlots: {} };
}

function saveAvailability(data) {
  try {
    fs.writeFileSync(AVAILABILITY_PATH, JSON.stringify(data, null, 2));
  } catch(e) { console.error('Error saving availability:', e.message); }
}

// ─────────────────────────────────────────────────────────────────────────────
// GET /api/availability
// Returns current blocked slots and closed dates — called by booking form
// ─────────────────────────────────────────────────────────────────────────────
app.get('/api/availability', (req, res) => {
  const data = loadAvailability();
  res.json({ success: true, ...data });
});

// ─────────────────────────────────────────────────────────────────────────────
// POST /api/availability
// Admin saves blocked slots and closed dates — called by admin dashboard
// Requires admin password in request body
// ─────────────────────────────────────────────────────────────────────────────
app.post('/api/availability', (req, res) => {
  const { password, closedDates, blockedSlots } = req.body;
  if (password !== process.env.ADMIN_PASSWORD) {
    return res.status(401).json({ success: false, error: 'Unauthorized.' });
  }
  saveAvailability({ closedDates, blockedSlots });
  console.log('✅ Availability updated by admin');
  res.json({ success: true });
});

// ─────────────────────────────────────────────────────────────────────────────
// POST /api/square/charge
// Called by the frontend after the Square card form is tokenized.
// ─────────────────────────────────────────────────────────────────────────────
app.post('/api/square/charge', async (req, res) => {
  const { sourceId, amountMoney, booking } = req.body;

  if (!sourceId || !amountMoney || !booking) {
    return res.status(400).json({ success: false, error: 'Missing required fields.' });
  }

  try {
    // 1. Charge the card via Square
    const { result } = await squareClient.payments.createPayment({
      sourceId,
      idempotencyKey:    crypto.randomUUID(),
      amountMoney: {
        amount:   BigInt(amountMoney.amount), // Square uses cents as BigInt
        currency: 'USD',
      },
      locationId:        process.env.SQUARE_LOCATION_ID,
      note:              `BioPrecision: ${booking.session}`,
      buyerEmailAddress: booking.email,
    });

    const payment = result.payment;
    console.log('✅ Square payment successful:', payment.id);

    // 2. Run all post-payment integrations in parallel
    await Promise.all([
      submitMondayForm(booking),
      updateRosterExcel(booking),
      sendOwnerNotification(booking, payment, amountMoney.amount),
      sendClientConfirmation(booking, payment, amountMoney.amount),
    ]);

    res.json({
      success:    true,
      paymentId:  payment.id,
      receiptUrl: payment.receiptUrl,
    });

  } catch (error) {
    console.error('❌ Charge error:', error);
    const msg = error?.errors?.[0]?.detail || error?.message || 'Payment processing failed.';
    res.status(500).json({ success: false, error: msg });
  }
});

// ─────────────────────────────────────────────────────────────────────────────
// MONDAY.COM — New Event Request Form
// Submits automatically after payment. Two-Way sessions submit twice
// (once for Batting Practice / hitting, once for Bullpen / pitching).
// ─────────────────────────────────────────────────────────────────────────────
function getMondaySessionTypes(serviceType, side) {
  // Two-way individual or team with both sides = two submissions
  if (serviceType === 'both' || side === 'both') {
    return ['Batting Practice', 'Bullpen'];
  }
  if (side === 'hitting' || serviceType === 'hitting') return ['Batting Practice'];
  if (side === 'pitching' || serviceType === 'pitching') return ['Bullpen'];
  return ['Workout'];
}

function getMondayDuration(sessionType) {
  if (sessionType === 'team')       return '240 minutes';
  if (sessionType === 'assessment') return '90 minutes';
  return '60 minutes';
}

function getMondayLevel(level) {
  // Map BioPrecision levels to Monday.com dropdown options
  const map = {
    'High school':             'High School',
    'College':                 'College',
    'Amateur / independent':   'Amateur',
    'Professional':            'MLB',
  };
  return map[level] || 'Other';
}

async function submitMondayForm(booking) {
  const sessionTypes = getMondaySessionTypes(booking.serviceType, booking.side);
  const duration     = getMondayDuration(booking.sessionType);
  const level        = getMondayLevel(booking.level);

  for (const sessionType of sessionTypes) {
    const eventName = booking.clientType === 'team'
      ? `BioPrecision — ${booking.session} · ${booking.teamName}`
      : `BioPrecision — ${booking.session} · ${booking.name}`;

    // Build column values for Monday.com
    // Column IDs below must match your actual Monday.com board column IDs.
    // Find them in: Board Settings → Columns → click a column → ID shown in URL.
    const columnValues = JSON.stringify({
      name:            eventName,               // Event Name
      date4:           { date: booking.date },  // Event Date  (use your column ID)
      hour:            { hour: booking.time },  // Event Time EST
      dropdown:        { labels: ['KinaTrax'] },// Product
      dropdown1:       { labels: ['Tier 2 (System Check + Live Support)'] }, // Support Tier
      text:            duration,                // Expected Duration
      dropdown2:       { labels: [level] },     // Level
      dropdown3:       { labels: [sessionType] },// Session Type
      text1:           booking.name,            // Requestor Name
      email:           { email: booking.email, text: booking.email }, // Requestor Email
    });

    const mutation = `
      mutation {
        create_item(
          board_id: ${process.env.MONDAY_BOARD_ID},
          item_name: "${eventName.replace(/"/g, '\\"')}",
          column_values: "${columnValues.replace(/"/g, '\\"')}"
        ) {
          id
        }
      }
    `;

    const response = await axios.post(
      'https://api.monday.com/v2',
      { query: mutation },
      {
        headers: {
          'Authorization': process.env.MONDAY_API_TOKEN,
          'Content-Type':  'application/json',
          'API-Version':   '2024-01',
        },
      }
    );

    if (response.data.errors) {
      console.error('Monday.com error:', response.data.errors);
    } else {
      console.log(`✅ Monday.com event created (${sessionType}):`, response.data.data?.create_item?.id);
    }
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// EXCEL ROSTER UPDATE
// Individual → append row to "Individual Clients" tab
// Team       → create new tab named after the team
// Then email the updated file to BioPrecision.
// ─────────────────────────────────────────────────────────────────────────────
async function updateRosterExcel(booking) {
  if (!fs.existsSync(ROSTER_PATH)) {
    console.warn('⚠️  Roster file not found at', ROSTER_PATH);
    return;
  }

  const wb = XLSX.readFile(ROSTER_PATH);

  if (booking.clientType === 'individual') {
    // ── Add to Individual Clients tab ──────────────────────────────────────
    const sheetName = 'Individual Clients';
    const ws        = wb.Sheets[sheetName];
    const data      = XLSX.utils.sheet_to_json(ws, { defval: '' });

    data.push({
      ID:            '',
      FirstName:     booking.firstName,
      LastName:      booking.lastName,
      UniformNumber: '',
      Weight:        booking.weight || '',
    });

    wb.Sheets[sheetName] = XLSX.utils.json_to_sheet(data);

    // Style the header row navy (basic — full styling requires xlsx-style)
    console.log(`✅ Excel: added ${booking.name} to Individual Clients tab`);

  } else {
    // ── Create new tab for team ────────────────────────────────────────────
    const teamName = booking.teamName || 'New Team';

    // If tab already exists (re-booking same team), append; otherwise create
    const existingSheet = wb.Sheets[teamName];
    let data = existingSheet
      ? XLSX.utils.sheet_to_json(existingSheet, { defval: '' })
      : [];

    booking.roster.forEach(p => {
      data.push({
        ID:            '',
        FirstName:     p.firstName  || p.name?.split(' ')[0] || '',
        LastName:      p.lastName   || p.name?.split(' ')[1] || '',
        UniformNumber: p.uniform    || '',
        Weight:        p.weight     || '',
      });
    });

    const ws = XLSX.utils.json_to_sheet(data);
    if (existingSheet) {
      wb.Sheets[teamName] = ws;
    } else {
      XLSX.utils.book_append_sheet(wb, ws, teamName);
    }

    console.log(`✅ Excel: created/updated tab "${teamName}" with ${booking.roster.length} players`);
  }

  // Save updated file
  XLSX.writeFile(wb, ROSTER_PATH);

  // Email updated file to BioPrecision
  const subject = booking.clientType === 'team'
    ? `Roster updated — New team: ${booking.teamName}`
    : `Roster updated — ${booking.name} added to Individual Clients`;

  await sendEmail({
    from:    `"BioPrecision System" <bookings@bioprecision.com>`,
    to:      process.env.NOTIFY_EMAIL,
    subject,
    text:    'The BioPrecision roster has been updated after a confirmed payment. See the attached file.',
    // Note: Resend attachment support requires base64 encoding
    // Roster file update is logged server-side
  });

  console.log('✅ Updated roster emailed to BioPrecision.');
}

// ─────────────────────────────────────────────────────────────────────────────
// EMAIL — Owner notification
// ─────────────────────────────────────────────────────────────────────────────
async function sendOwnerNotification(booking, payment, amountCents) {
  const amount = (Number(amountCents) / 100).toFixed(2);

  await sendEmail({
    from:    `"BioPrecision Booking" <bookings@bioprecision.com>`,
    to:      process.env.NOTIFY_EMAIL,
    subject: `New Booking — ${booking.name} | ${booking.session}`,
    html: `
      <div style="font-family:sans-serif;max-width:560px">
        <div style="background:#011244;color:#fff;padding:16px 20px;border-radius:8px 8px 0 0">
          <h2 style="margin:0;font-size:18px">New BioPrecision Session Booked</h2>
        </div>
        <div style="border:1px solid #E5E7EB;border-top:none;border-radius:0 0 8px 8px;padding:20px">
          <table style="font-size:14px;border-collapse:collapse;width:100%">
            <tr><td style="padding:6px 0;color:#6B7280;width:140px">Athlete / Coach</td><td><strong>${booking.name}</strong></td></tr>
            <tr><td style="padding:6px 0;color:#6B7280">Email</td><td>${booking.email}</td></tr>
            <tr><td style="padding:6px 0;color:#6B7280">Phone</td><td>${booking.phone || '—'}</td></tr>
            <tr><td style="padding:6px 0;color:#6B7280">Level</td><td>${booking.level || '—'}</td></tr>
            <tr><td style="padding:6px 0;color:#6B7280">Session</td><td>${booking.session}</td></tr>
            <tr><td style="padding:6px 0;color:#6B7280">Lab side</td><td>${booking.side || '—'}</td></tr>
            <tr><td style="padding:6px 0;color:#6B7280">Date & Time</td><td>${booking.dateTime || 'Remote / Team TBD'}</td></tr>
            <tr><td style="padding:6px 0;color:#6B7280">Referred by</td><td>${booking.referral || '—'}</td></tr>
            <tr><td style="padding:6px 0;color:#6B7280">Amount Paid</td><td><strong style="color:#011244">$${amount}</strong></td></tr>
            <tr><td style="padding:6px 0;color:#6B7280">Square Payment ID</td><td style="font-size:12px;color:#9CA3AF">${payment.id}</td></tr>
            <tr><td style="padding:6px 0;color:#6B7280">Receipt</td><td><a href="${payment.receiptUrl}" style="color:#011244">View Square receipt</a></td></tr>
          </table>
          <hr style="border:none;border-top:1px solid #E5E7EB;margin:16px 0">
          <p style="font-size:12px;color:#9CA3AF;margin:0">Monday.com event request and Excel roster have been updated automatically.</p>
        </div>
      </div>
    `,
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// EMAIL — Client confirmation
// ─────────────────────────────────────────────────────────────────────────────
async function sendClientConfirmation(booking, payment, amountCents) {
  const amount    = (Number(amountCents) / 100).toFixed(2);
  const firstName = booking.firstName || booking.name?.split(' ')[0] || 'Athlete';
  const isTeam    = booking.clientType === 'team';
  const isRemote  = booking.sessionType === 'remote';

  const dateTimeRow = (isTeam || isRemote)
    ? `<tr><td style="padding:6px 0;color:#6B7280;width:130px">Scheduling</td><td>${isTeam ? 'BioPrecision will contact you within 24 hrs' : 'Calendar invite & video call link incoming'}</td></tr>`
    : `<tr><td style="padding:6px 0;color:#6B7280">Date & Time</td><td><strong>${booking.dateTime}</strong></td></tr>`;

  await sendEmail({
    from:    `"BioPrecision" <bookings@bioprecision.com>`,
    to:      booking.email,
    subject: `You're booked! — BioPrecision Session Confirmation`,
    html: `
      <div style="font-family:sans-serif;max-width:520px">
        <div style="background:#011244;color:#fff;padding:16px 20px;border-radius:8px 8px 0 0">
          <h2 style="margin:0;font-size:18px">You're booked at BioPrecision!</h2>
        </div>
        <div style="border:1px solid #E5E7EB;border-top:none;border-radius:0 0 8px 8px;padding:20px">
          <p style="font-size:14px;color:#374151;margin-bottom:16px">Hi ${firstName}, your session has been confirmed and payment received. Here's your booking summary:</p>
          <table style="font-size:14px;border-collapse:collapse;width:100%">
            <tr><td style="padding:6px 0;color:#6B7280;width:130px">Session</td><td><strong>${booking.session}</strong></td></tr>
            ${dateTimeRow}
            <tr><td style="padding:6px 0;color:#6B7280">Amount Paid</td><td><strong style="color:#011244">$${amount}</strong></td></tr>
            <tr><td style="padding:6px 0;color:#6B7280">Receipt</td><td><a href="${payment.receiptUrl}" style="color:#011244">View your Square receipt</a></td></tr>
          </table>
          <hr style="border:none;border-top:1px solid #E5E7EB;margin:16px 0">
          <p style="font-size:13px;color:#6B7280;margin:0">
            BioPrecision LLC · WVU Baseball Biomechanics and Performance Center<br>
            2040 Gyorko Dr, Morgantown, WV 26534 · bioprecision.com
          </p>
        </div>
      </div>
    `,
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// POST /api/test/charge
// Simulates a successful payment without charging a real card.
// Only works when TEST_MODE=true in environment variables.
// ─────────────────────────────────────────────────────────────────────────────
app.post('/api/test/charge', async (req, res) => {
  if (process.env.TEST_MODE !== 'true') {
    return res.status(403).json({ success: false, error: 'Test mode is not enabled.' });
  }

  const { amountMoney, booking } = req.body;
  if (!booking) return res.status(400).json({ success: false, error: 'Missing booking data.' });

  try {
    const fakePayment = {
      id: 'TEST_' + crypto.randomUUID().substring(0, 8).toUpperCase(),
      receiptUrl: 'https://squareup.com/receipt/preview/TEST',
    };

    await Promise.all([
      sendOwnerNotification(booking, fakePayment, amountMoney?.amount || 0),
      sendClientConfirmation(booking, fakePayment, amountMoney?.amount || 0),
    ]);

    console.log('✅ TEST payment simulation successful for:', booking.name);
    res.json({ success: true, paymentId: fakePayment.id, receiptUrl: fakePayment.receiptUrl, testMode: true });

  } catch (error) {
    console.error('❌ Test charge error:', error.message);
    res.status(500).json({ success: false, error: error.message });
  }
});

// ─────────────────────────────────────────────────────────────────────────────
// POST /api/team/request
// Fires when a team client submits their booking request (no payment yet).
// Sends a notification email to BioPrecision and a confirmation to the coach.
// ─────────────────────────────────────────────────────────────────────────────
app.post('/api/team/request', async (req, res) => {
  const { booking } = req.body;
  if (!booking) return res.status(400).json({ success: false, error: 'Missing booking data.' });

  try {
    // Email to BioPrecision owner
    await sendEmail({
      from:    `"BioPrecision Booking" <bookings@bioprecision.com>`,
      to:      process.env.NOTIFY_EMAIL,
      subject: `New Team Booking Request — ${booking.teamName || booking.name}`,
      html: `
        <div style="font-family:sans-serif;max-width:560px">
          <div style="background:#011244;color:#fff;padding:16px 20px;border-radius:8px 8px 0 0">
            <h2 style="margin:0;font-size:18px">⚡ New Team Booking Request</h2>
          </div>
          <div style="border:1px solid #E5E7EB;border-top:none;border-radius:0 0 8px 8px;padding:20px">
            <p style="font-size:14px;color:#374151;margin-bottom:16px">A team has submitted a booking request. Contact them to confirm a date and schedule their session.</p>
            <table style="font-size:14px;border-collapse:collapse;width:100%">
              <tr><td style="padding:6px 0;color:#6B7280;width:160px">Coach / Contact</td><td><strong>${booking.name}</strong></td></tr>
              <tr><td style="padding:6px 0;color:#6B7280">Email</td><td><a href="mailto:${booking.email}" style="color:#011244">${booking.email}</a></td></tr>
              <tr><td style="padding:6px 0;color:#6B7280">Phone</td><td>${booking.phone || '—'}</td></tr>
              <tr><td style="padding:6px 0;color:#6B7280">Team / Organization</td><td>${booking.teamName || booking.teamSchool || '—'}</td></tr>
              <tr><td style="padding:6px 0;color:#6B7280">Level</td><td>${booking.level || '—'}</td></tr>
              <tr><td style="padding:6px 0;color:#6B7280">Service requested</td><td>${booking.session || '—'}</td></tr>
              <tr><td style="padding:6px 0;color:#6B7280">Service type</td><td>${booking.serviceType || '—'}</td></tr>
              <tr><td style="padding:6px 0;color:#6B7280">Referred by</td><td>${booking.referral || '—'}</td></tr>
              <tr><td style="padding:6px 0;color:#6B7280">Goals / notes</td><td>${booking.goals || '—'}</td></tr>
            </table>
            <hr style="border:none;border-top:1px solid #E5E7EB;margin:16px 0">
            <p style="font-size:12px;color:#9CA3AF;margin:0">Reply directly to <a href="mailto:${booking.email}" style="color:#011244">${booking.email}</a> to follow up with this team.</p>
          </div>
        </div>
      `,
    });

    // Confirmation email to coach
    await sendEmail({
      from:    `"BioPrecision" <bookings@bioprecision.com>`,
      to:      booking.email,
      subject: `Team Booking Request Received — BioPrecision`,
      html: `
        <div style="font-family:sans-serif;max-width:520px">
          <div style="background:#011244;color:#fff;padding:16px 20px;border-radius:8px 8px 0 0">
            <h2 style="margin:0;font-size:18px">Team Booking Request Received</h2>
          </div>
          <div style="border:1px solid #E5E7EB;border-top:none;border-radius:0 0 8px 8px;padding:20px">
            <p style="font-size:14px;color:#374151;margin-bottom:16px">Hi ${booking.name.split(' ')[0]}, we've received your team booking request and will be in touch within 24 hours to confirm your date, time, and session details.</p>
            <table style="font-size:14px;border-collapse:collapse;width:100%">
              <tr><td style="padding:6px 0;color:#6B7280;width:130px">Service</td><td><strong>${booking.session || '—'}</strong></td></tr>
              <tr><td style="padding:6px 0;color:#6B7280">Type</td><td>${booking.serviceType || '—'}</td></tr>
            </table>
            <hr style="border:none;border-top:1px solid #E5E7EB;margin:16px 0">
            <p style="font-size:13px;color:#6B7280;margin:0">
              BioPrecision LLC · WVU Baseball Biomechanics and Performance Center<br>
              2040 Gyorko Dr, Morgantown, WV 26534 · bioprecision.com
            </p>
          </div>
        </div>
      `,
    });

    console.log('✅ Team request emails sent for:', booking.name);
    res.json({ success: true });

  } catch (error) {
    console.error('❌ Team request email error:', error.message);
    console.error('   Full error:', JSON.stringify(error, null, 2));
    res.status(500).json({ success: false, error: error.message });
  }
});

// ─────────────────────────────────────────────────────────────────────────────
// START SERVER
// ─────────────────────────────────────────────────────────────────────────────
const PORT = process.env.PORT || 8080;

// Health check for Railway
app.get('/health', (req, res) => res.status(200).send('OK'));

// Serve booking form
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Serve admin dashboard
app.get('/admin', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'admin.html'));
});

app.listen(PORT, '0.0.0.0', () => {
  console.log(`\n🚀 BioPrecision server running on port ${PORT}`);
  console.log(`   Square location: ${process.env.SQUARE_LOCATION_ID}`);
  console.log(`   SMTP host: ${process.env.SMTP_HOST}`);
  console.log(`   SMTP user: ${process.env.SMTP_USER}`);
  console.log(`   SMTP pass set: ${process.env.SMTP_PASS ? 'YES (' + process.env.SMTP_PASS.length + ' chars)' : 'NO - MISSING'}`);
  console.log(`   Notifications → ${process.env.NOTIFY_EMAIL}\n`);

  console.log('✅ Resend email service configured');
});
