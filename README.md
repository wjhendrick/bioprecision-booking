# BioPrecision Booking System — Deployment Guide

## What's in this package

```
bioprecision-booking/
├── public/
│   └── index.html              ← Full booking form (copy from Claude conversation)
├── server.js                   ← Node.js backend (Square + Monday.com + Excel)
├── BioPrecision_Roster.xlsx    ← Your roster file (copy from Claude outputs)
├── .env                        ← Your secrets (rename from .env.example)
├── .env.example                ← Template — fill in and rename to .env
├── package.json                ← Dependencies
└── .gitignore                  ← Protects your .env
```

---

## Step 1 — Fill in your .env file

Rename `.env.example` to `.env` and fill in:

| Variable | Where to find it |
|----------|-----------------|
| `SQUARE_ACCESS_TOKEN` | developer.squareup.com → your app → Credentials → **Regenerate first!** |
| `SQUARE_LOCATION_ID` | `L23VG08ZFF9G3` (already set) |
| `SQUARE_APP_ID` | `sq0idp-K-gpGbesdSUK9ZyAEBSaQw` (already set) |
| `MONDAY_API_TOKEN` | monday.com → Profile → Developers → API |
| `MONDAY_BOARD_ID` | Open your board → number in the URL |
| `SMTP_USER` | Your Gmail address |
| `SMTP_PASS` | Gmail App Password (16 chars, not your login password) |
| `NOTIFY_EMAIL` | Your BioPrecision inbox |

---

## Step 2 — Get your Monday.com column IDs

The server code uses placeholder column IDs. You need to replace them with your actual board column IDs:

1. Open your Monday.com board
2. Click a column header → **Customize this column**
3. The column ID appears in the URL or in the column settings
4. Replace the column IDs in `server.js` in the `submitMondayForm` function

---

## Step 3 — Install and run locally first

```bash
npm install
node server.js
```

Visit `http://localhost:3000` — your booking form should load.

Test with Square sandbox card: `4111 1111 1111 1111` · any future date · any CVV · any ZIP

---

## Step 4 — Deploy to Railway (recommended)

1. Go to [railway.app](https://railway.app) and sign up with GitHub
2. Click **"New Project"** → **"Deploy from GitHub repo"**
3. Connect your repo (push your code to a private GitHub repo first)
4. In Railway, go to **Variables** and add all your `.env` values
5. Railway auto-deploys — your booking system gets a live URL
6. Point your bioprecision.com domain to it (Railway → Settings → Custom Domain)

**⚠️ Never push your `.env` file to GitHub.** Add it to `.gitignore`:
```
.env
node_modules/
```

---

## Step 5 — Update your booking form HTML

In `public/index.html`, find the Square initialization section and update:

```javascript
const SQUARE_APP_ID     = 'sq0idp-K-gpGbesdSUK9ZyAEBSaQw';
const SQUARE_LOCATION_ID = 'L23VG08ZFF9G3';
```

And swap the Square SDK script tag to production:
```html
<!-- Remove sandbox URL, use this: -->
<script src="https://web.squarecdn.com/v1/square.js"></script>
```

---

## Step 6 — Regenerate your Square Access Token

Since this token was shared during setup:

1. Go to developer.squareup.com → your app → Credentials
2. Click **Regenerate** next to Production Access Token
3. Update your `.env` and Railway environment variables with the new token

---

## Testing checklist

- [ ] Square sandbox payment processes successfully
- [ ] Owner notification email arrives with booking details
- [ ] Client confirmation email arrives with receipt link
- [ ] Monday.com event appears on your board
- [ ] BioPrecision_Roster.xlsx arrives by email with new row/tab
- [ ] Two-Way sessions create TWO Monday.com events
- [ ] Team sessions create a new Excel tab named after the team

---

## Support

Built by Claude (Anthropic) for BioPrecision LLC.
Questions? Continue the conversation at claude.ai.
