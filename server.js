require('dotenv').config();
process.on('uncaughtException', e => console.error('uncaughtException:', e.message));
process.on('unhandledRejection', e => console.error('unhandledRejection:', e?.message || e));
const express = require('express');
const QRCode = require('qrcode');
const { Client, LocalAuth } = require('whatsapp-web.js');
const qrcode = require('qrcode-terminal');
const Anthropic = require('@anthropic-ai/sdk');
const cron = require('node-cron');
const XLSX = require('xlsx');
const pdfParse = require('pdf-parse');
const multer = require('multer');
const ical = require('node-ical');
const https = require('https');
const fs = require('fs');
const path = require('path');
const { google } = require('googleapis');
const axios = require('axios');
const crypto = require('crypto');

const app = express();
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

const upload = multer({ storage: multer.memoryStorage() });
const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

const GROUPS = ['Out en nieuw', 'Le Pompadour☕🥂', 'De Petjes (met kleinkinderen)'];
const matchGroup = name => GROUPS.find(g => name === g || name.startsWith(g));
const MY_NAME = 'Nico';
const SCHEDULE_FILE = path.join(__dirname, 'schedule.json');
const SUMMARIES_FILE = path.join(__dirname, 'summaries.json');
const HVA_FILE = path.join(__dirname, 'hva.json');
const MESSAGES_FILE = path.join(__dirname, 'messages.json');

let waStatus = 'disconnected';
let waQR = null;
const webChatHistory = [];

// ─── Message store (real-time accumulation for group summaries) ──────────────

function loadMessages() {
    if (fs.existsSync(MESSAGES_FILE)) return JSON.parse(fs.readFileSync(MESSAGES_FILE, 'utf8'));
    return {};
}

function saveMessages(data) {
    fs.writeFileSync(MESSAGES_FILE, JSON.stringify(data, null, 2));
}

function storeGroupMessage(groupName, senderName, body) {
    const data = loadMessages();
    if (!data[groupName]) data[groupName] = [];
    data[groupName].push({ sender: senderName, body, ts: Date.now() });
    // Keep only last 7 days
    const cutoff = Date.now() - 7 * 24 * 60 * 60 * 1000;
    data[groupName] = data[groupName].filter(m => m.ts > cutoff);
    saveMessages(data);
}

// ─── Helpers ────────────────────────────────────────────────────────────────

function loadSchedule() {
    if (fs.existsSync(SCHEDULE_FILE)) return JSON.parse(fs.readFileSync(SCHEDULE_FILE, 'utf8'));
    return null;
}

function saveSchedule(data) {
    fs.writeFileSync(SCHEDULE_FILE, JSON.stringify(data, null, 2));
}

// Normalize any stored format to array of rowsets
function getRowsets() {
    const data = loadSchedule();
    if (!data) return [];
    if (data.rowsets) return data.rowsets;
    if (data.rows) return [{ rows: data.rows, filename: 'rooster.xlsx', date: Date.now() }];
    return [];
}

function loadSummaries() {
    if (fs.existsSync(SUMMARIES_FILE)) return JSON.parse(fs.readFileSync(SUMMARIES_FILE, 'utf8'));
    return { groups: {}, lastUpdated: null };
}

function saveSummaries(data) {
    fs.writeFileSync(SUMMARIES_FILE, JSON.stringify(data, null, 2));
}

const RECURRING_FILE = path.join(__dirname, 'recurring.json');
const GMAIL_TOKEN_FILE = path.join(__dirname, 'gmail_token.json');
const MAIL_CACHE_FILE  = path.join(__dirname, 'mail_cache.json');

// ─── Gmail OAuth ─────────────────────────────────────────────────────────────

const GMAIL_SCOPES = ['https://www.googleapis.com/auth/gmail.readonly'];

function getOAuth2Client() {
    const redirect = process.env.GMAIL_REDIRECT_URI || 'http://localhost:3000/api/gmail/callback';
    return new google.auth.OAuth2(
        process.env.GMAIL_CLIENT_ID,
        process.env.GMAIL_CLIENT_SECRET,
        redirect
    );
}

function loadGmailToken() {
    if (fs.existsSync(GMAIL_TOKEN_FILE)) return JSON.parse(fs.readFileSync(GMAIL_TOKEN_FILE, 'utf8'));
    return null;
}

function saveGmailToken(token) {
    fs.writeFileSync(GMAIL_TOKEN_FILE, JSON.stringify(token, null, 2));
}

function getAuthedClient() {
    const token = loadGmailToken();
    if (!token) return null;
    const auth = getOAuth2Client();
    auth.setCredentials(token);
    auth.on('tokens', t => { if (t.refresh_token) saveGmailToken({ ...token, ...t }); });
    return auth;
}

// ─── Gmail fetch & classify ───────────────────────────────────────────────────

function loadMailCache() {
    if (fs.existsSync(MAIL_CACHE_FILE)) return JSON.parse(fs.readFileSync(MAIL_CACHE_FILE, 'utf8'));
    return { mails: [], lastFetch: null };
}

function saveMailCache(data) {
    fs.writeFileSync(MAIL_CACHE_FILE, JSON.stringify(data, null, 2));
}

function decodeBody(part) {
    if (!part) return '';
    if (part.body?.data) return Buffer.from(part.body.data, 'base64').toString('utf8');
    if (part.parts) {
        for (const p of part.parts) { const t = decodeBody(p); if (t) return t; }
    }
    return '';
}

function stripHtml(html) {
    return html.replace(/<style[\s\S]*?<\/style>/gi, '')
               .replace(/<[^>]+>/g, ' ')
               .replace(/\s{2,}/g, ' ')
               .trim()
               .slice(0, 3000);
}

async function fetchAndClassifyMails() {
    const auth = getAuthedClient();
    if (!auth) return null;

    const gmail = google.gmail({ version: 'v1', auth });

    // Fetch last 30 unread + recent threads from inbox
    const listRes = await gmail.users.messages.list({
        userId: 'me',
        labelIds: ['INBOX'],
        maxResults: 40,
        q: 'newer_than:7d'
    });

    const messages = listRes.data.messages || [];
    if (!messages.length) return { mails: [], lastFetch: new Date().toISOString() };

    const mails = [];
    for (const { id, threadId } of messages) {
        try {
            const msg = await gmail.users.messages.get({ userId: 'me', id, format: 'full' });
            const headers = msg.data.payload.headers || [];
            const get = name => headers.find(h => h.name.toLowerCase() === name.toLowerCase())?.value || '';

            const subject = get('Subject') || '(geen onderwerp)';
            const from    = get('From');
            const to      = get('To');
            const date    = get('Date');
            const rawBody = stripHtml(decodeBody(msg.data.payload));
            const isRead  = !msg.data.labelIds?.includes('UNREAD');
            const snippet = msg.data.snippet || '';

            mails.push({ id, threadId, subject, from, to, date, snippet, body: rawBody, isRead, labels: msg.data.labelIds || [] });
        } catch {}
    }

    // Group by thread and classify with Claude
    const classified = await classifyMails(mails);
    const result = { mails: classified, lastFetch: new Date().toISOString() };
    saveMailCache(result);
    return result;
}

async function classifyMails(mails) {
    if (!mails.length) return [];

    // Build compact list for Claude
    const mailList = mails.map((m, i) =>
        `[${i}] Van: ${m.from}\nOnderwerp: ${m.subject}\nDatum: ${m.date}\nInhoud: ${m.snippet}`
    ).join('\n\n');

    const res = await anthropic.messages.create({
        model: 'claude-sonnet-4-6',
        max_tokens: 4096,
        system: `Je bent een slimme e-mail assistent voor Nico. Analyseer deze inbox en classificeer elke mail.
Geef voor elke mail terug:
- priority: "urgent" (actie vereist vandaag), "waiting" (kan wachten), "info" (ter info/nieuwsbrief)
- summary: 1-2 zinnen, direct en concreet, geen fluff
- action: beschrijf wat Nico moet doen (of null als niks nodig)
- deadline: eventuele deadline die wordt genoemd (of null)
- needsReply: true als iemand een vraag stelt of op antwoord wacht

Output: JSON array met exact ${mails.length} objecten in dezelfde volgorde als de input, elk met velden: index, priority, summary, action, deadline, needsReply`,
        messages: [{ role: 'user', content: mailList }]
    });

    let classifications = [];
    try {
        const jsonStr = res.content[0].text.match(/\[[\s\S]*\]/)?.[0];
        classifications = JSON.parse(jsonStr);
    } catch { return mails.map(m => ({ ...m, priority: 'info', summary: m.snippet, action: null, deadline: null, needsReply: false })); }

    return mails.map((m, i) => {
        const c = classifications.find(c => c.index === i) || {};
        return { ...m, priority: c.priority || 'info', summary: c.summary || m.snippet, action: c.action || null, deadline: c.deadline || null, needsReply: !!c.needsReply };
    });
}

function loadRecurring() {
    if (fs.existsSync(RECURRING_FILE)) return JSON.parse(fs.readFileSync(RECURRING_FILE, 'utf8'));
    return [];
}

function saveRecurring(data) {
    fs.writeFileSync(RECURRING_FILE, JSON.stringify(data, null, 2));
}

function generateRecurringDates() {
    const rules = loadRecurring();
    const result = {};
    const now = new Date();
    const end = new Date(now.getFullYear() + 1, 11, 31);

    for (const rule of rules) {
        const exceptions = new Set((rule.exceptions || []).map(d => {
            // Get ISO week of exception date, mark that Wednesday as exception
            const date = new Date(d);
            const wed = new Date(date);
            wed.setDate(date.getDate() + ((3 - date.getDay() + 7) % 7));
            return wed.toISOString().split('T')[0];
        }));

        // Find first occurrence of the weekday after today
        const DAY = { maandag:1,dinsdag:2,woensdag:3,donderdag:4,vrijdag:5,zaterdag:6,zondag:0 };
        const targetDay = DAY[rule.day.toLowerCase()] ?? 3;

        const cur = new Date(now);
        cur.setHours(0,0,0,0);
        const diff = (targetDay - cur.getDay() + 7) % 7;
        cur.setDate(cur.getDate() + diff);

        while (cur <= end) {
            const key = cur.toISOString().split('T')[0];
            if (!exceptions.has(key)) {
                if (!result[key]) result[key] = [];
                result[key].push({ title: rule.label || 'Vaste dag', time: rule.time, type: 'recurring' });
            }
            cur.setDate(cur.getDate() + 7);
        }
    }
    return result;
}

function loadHva() {
    if (fs.existsSync(HVA_FILE)) return JSON.parse(fs.readFileSync(HVA_FILE, 'utf8'));
    return { events: {}, lastUpdated: null };
}

function saveHva(data) {
    fs.writeFileSync(HVA_FILE, JSON.stringify(data, null, 2));
}

async function fetchHvaSchedule() {
    const url = process.env.HVA_ICAL_URL;
    if (!url) return;
    try {
        const events = await ical.async.fromURL(url);
        const byDate = {};
        const now = new Date();
        const cutoff = new Date(now.getFullYear(), now.getMonth() - 1, 1);

        for (const key of Object.keys(events)) {
            const ev = events[key];
            if (ev.type !== 'VEVENT' || !ev.start) continue;
            if (new Date(ev.start) < cutoff) continue;

            const dateKey = new Date(ev.start).toISOString().split('T')[0];
            if (!byDate[dateKey]) byDate[dateKey] = [];

            const startTime = new Date(ev.start).toLocaleTimeString('nl-NL', { hour: '2-digit', minute: '2-digit', timeZone: 'Europe/Amsterdam' });
            const endTime = new Date(ev.end).toLocaleTimeString('nl-NL', { hour: '2-digit', minute: '2-digit', timeZone: 'Europe/Amsterdam' });

            const title = ev.summary || 'Les';
            const isFree = /^[A-Z0-9\s\-]+$/.test(title.trim()) ||
                /vrij|vakantie|studievrij|onderwijsvrij|pinkster|pasen|kerst|oud\s*jaar/i.test(title);
            byDate[dateKey].push({
                title,
                location: ev.location || '',
                time: `${startTime}-${endTime}`,
                ...(isFree ? { type: 'free' } : {})
            });
        }

        // Detect vacation: whole weeks (Mon-Fri) with zero class events inside school period
        const classDates = Object.keys(byDate).filter(k => byDate[k].some(e => !e.type));
        if (classDates.length >= 2) {
            classDates.sort();
            const firstClass = new Date(classDates[0]);
            const lastClass  = new Date(classDates[classDates.length - 1]);

            // Get Monday of first week through Monday of last week
            const mondayOf = d => { const m = new Date(d); m.setDate(m.getDate() - ((m.getDay() + 6) % 7)); m.setHours(0,0,0,0); return m; };
            const cur = mondayOf(firstClass);
            while (cur <= lastClass) {
                // Check if this Mon-Fri week has any class events
                let hasClass = false;
                for (let i = 0; i < 5; i++) {
                    const day = new Date(cur); day.setDate(cur.getDate() + i);
                    const key = day.toISOString().split('T')[0];
                    if (byDate[key] && byDate[key].some(e => !e.type)) { hasClass = true; break; }
                }
                if (!hasClass) {
                    for (let i = 0; i < 5; i++) {
                        const day = new Date(cur); day.setDate(cur.getDate() + i);
                        const key = day.toISOString().split('T')[0];
                        if (!byDate[key]) byDate[key] = [{ title: 'Schoolvakantie', time: '', location: '', type: 'free' }];
                    }
                }
                cur.setDate(cur.getDate() + 7);
            }
        }

        saveHva({ events: byDate, lastUpdated: new Date().toISOString() });
        console.log(`✅ HvA rooster opgehaald: ${Object.keys(byDate).length} dagen`);
    } catch (e) {
        console.error('HvA fetch fout:', e.message);
    }
}

function parseExcelRows(buffer) {
    const workbook = XLSX.read(buffer, { type: 'buffer', cellDates: true });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    return XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
}

function excelSerialToDate(serial) {
    const utcDays = Math.floor(serial - 25569);
    return new Date(utcDays * 86400 * 1000);
}

function tryParseDate(val) {
    if (!val) return null;
    if (val instanceof Date && !isNaN(val)) return val;
    if (typeof val === 'number' && val > 40000 && val < 60000) {
        return excelSerialToDate(val);
    }
    if (typeof val === 'string') {
        const cleaned = val.trim();
        // Try Dutch formats: "29 jun", "do 2", "29-6-2025", "2025-06-29"
        const isoMatch = cleaned.match(/^(\d{4})-(\d{2})-(\d{2})$/);
        if (isoMatch) return new Date(cleaned);
        const dmyMatch = cleaned.match(/^(\d{1,2})[-/](\d{1,2})[-/]?(\d{2,4})?$/);
        if (dmyMatch) {
            const y = dmyMatch[3] ? parseInt(dmyMatch[3]) : new Date().getFullYear();
            const year = y < 100 ? 2000 + y : y;
            return new Date(year, parseInt(dmyMatch[2]) - 1, parseInt(dmyMatch[1]));
        }
    }
    return null;
}

function toDateKey(d) {
    return d.toISOString().split('T')[0];
}

function getISOWeek(d) {
    const date = new Date(d);
    date.setHours(0, 0, 0, 0);
    date.setDate(date.getDate() + 3 - ((date.getDay() + 6) % 7));
    const week1 = new Date(date.getFullYear(), 0, 4);
    return 1 + Math.round(((date - week1) / 86400000 - 3 + ((week1.getDay() + 6) % 7)) / 7);
}

function getWeekDates(year, week) {
    const jan4 = new Date(year, 0, 4);
    const dayOfWeek = (jan4.getDay() + 6) % 7;
    const monday = new Date(jan4);
    monday.setDate(jan4.getDate() - dayOfWeek + (week - 1) * 7);
    return Array.from({ length: 7 }, (_, i) => {
        const d = new Date(monday);
        d.setDate(monday.getDate() + i);
        return d;
    });
}

function shiftColor(shiftStr) {
    const m = String(shiftStr).match(/(\d{1,2})[:.h](\d{0,2})/);
    if (!m) return '#5c6bc0';
    const hour = parseInt(m[1]);
    if (hour < 9)  return '#ef6c00';
    if (hour < 12) return '#f9a825';
    if (hour < 15) return '#43a047';
    if (hour < 17) return '#00897b';
    return '#00695c';
}

function parseRowset(rows) {
    const result = {};

    // Find header row (most date-like values in the top 4 rows)
    let headerIdx = 0;
    let maxDates = 0;
    for (let i = 0; i < Math.min(4, rows.length); i++) {
        const count = rows[i].filter(c => tryParseDate(c) !== null).length;
        if (count > maxDates) { maxDates = count; headerIdx = i; }
    }

    const header = rows[headerIdx];

    // Find all employee rows (name in first col, after header)
    const empRows = [];
    for (let i = headerIdx + 1; i < rows.length; i++) {
        const name = String(rows[i][0] || '').trim();
        if (name && name.length > 1) empRows.push({ name, row: rows[i] });
    }

    const nicoRows = empRows.filter(e => e.name.toLowerCase().includes(MY_NAME.toLowerCase()));
    if (!nicoRows.length) return {};

    for (const nicoEntry of nicoRows) {
        for (let j = 1; j < header.length; j++) {
            const dateVal = header[j];
            const shiftVal = String(nicoEntry.row[j] || '').trim();
            if (!shiftVal) continue;

            const date = tryParseDate(dateVal);
            if (!date) continue;

            const key = toDateKey(date);
            const colleagues = empRows
                .filter(e => !e.name.toLowerCase().includes(MY_NAME.toLowerCase()))
                .map(e => ({ name: e.name, shift: String(e.row[j] || '').trim() }))
                .filter(e => e.shift);

            result[key] = { shift: shiftVal, color: shiftColor(shiftVal), colleagues };
        }
    }

    return result;
}

function parseScheduleData() {
    const data = loadSchedule();
    if (!data) return {};
    const merged = {};
    if (data.parsedSchedules) Object.assign(merged, data.parsedSchedules);
    const rowsets = data.rowsets || (data.rows ? [{ rows: data.rows }] : []);
    for (const rs of rowsets) Object.assign(merged, parseRowset(rs.rows));
    for (const key of Object.keys(merged)) {
        if (merged[key] && !merged[key].color) merged[key].color = shiftColor(merged[key].shift || '');
    }
    return merged;
}

// ─── WhatsApp ───────────────────────────────────────────────────────────────

// Verwijder Chromium lock-bestanden zodat herstart na crash werkt
try {
    const { execSync } = require('child_process');
    const authDir = path.join(__dirname, '.wwebjs_auth');
    execSync(`find "${authDir}" -name "Singleton*" -exec rm -f {} \\; 2>/dev/null || true`);
    execSync(`find "${authDir}" -name "lockfile" -exec rm -f {} \\; 2>/dev/null || true`);
    execSync(`find "${authDir}" -name "*.lock" -exec rm -f {} \\; 2>/dev/null || true`);
} catch {}

const WWEBJS_AUTH_DIR = path.join(__dirname, '.wwebjs_auth');
const { execSync } = require('child_process');

function clearWaSession() {
    try { execSync(`pkill -f chromium 2>/dev/null; pkill -f chrome 2>/dev/null; true`); } catch {}
    try { execSync(`rm -rf "${WWEBJS_AUTH_DIR}/session" 2>/dev/null || true`); } catch {}
    console.log('WhatsApp sessie en Chrome processen gewist');
}

async function initWhatsApp() {
    console.log('WhatsApp client starten...');

    // Verwijder bestaande Chrome processen voor een schone start
    try { execSync(`pkill -f chromium 2>/dev/null; pkill -f chrome 2>/dev/null; true`); } catch {}

    // Timeout: als na 90s geen enkel event vuurt, sessie wissen en opnieuw proberen
    const initTimeout = setTimeout(async () => {
        console.error('WhatsApp init timeout — opnieuw starten');
        try { await client.destroy(); } catch {}
        clearWaSession();
        setTimeout(() => initWhatsApp(), 3000);
    }, 90000);

    client.once('qr', () => clearTimeout(initTimeout));
    client.once('ready', () => clearTimeout(initTimeout));
    client.once('auth_failure', () => clearTimeout(initTimeout));

    try {
        await client.initialize();
    } catch(e) {
        clearTimeout(initTimeout);
        console.error('WhatsApp initialize fout:', e.message);
        try { await client.destroy(); } catch {}
        clearWaSession();
        setTimeout(() => initWhatsApp(), 5000);
    }
}

const client = new Client({
    authStrategy: new LocalAuth(),
    puppeteer: {
        executablePath: process.env.CHROME_PATH || '/usr/bin/google-chrome-stable',
        args: [
            '--no-sandbox',
            '--disable-setuid-sandbox',
            '--disable-dev-shm-usage',
            '--disable-gpu',
            '--no-zygote',
            '--disable-extensions',
            '--no-first-run'
        ]
    }
});

client.on('qr', (qr) => {
    waQR = qr;
    waStatus = 'qr';
    qrcode.generate(qr, { small: true });
    console.log('QR code klaar om te scannen via /api/qr');
});

client.on('auth_failure', (msg) => {
    console.error('WhatsApp auth mislukt:', msg);
    waStatus = 'disconnected';
    waQR = null;
    clearWaSession();
    setTimeout(() => initWhatsApp(), 3000);
});

client.on('ready', () => {
    waStatus = 'connected';
    waQR = null;
    console.log('WhatsApp verbonden!');

    cron.schedule('0 18 * * *', () => sendDailySummary(), { timezone: 'Europe/Amsterdam' });
    cron.schedule('0 6 * * *', () => fetchHvaSchedule(), { timezone: 'Europe/Amsterdam' });
    cron.schedule('0 7,12 * * *', () => fetchAndClassifyMails().catch(() => {}), { timezone: 'Europe/Amsterdam' });
    fetchHvaSchedule();
    fetchRoosterFromKevin().catch(e => console.error('Rooster ophalen fout:', e.message));
});

client.on('disconnected', () => { waStatus = 'disconnected'; });

async function sendDailySummary() {
    const stored = loadMessages();
    const myNumber = client.info.wid._serialized;
    const summaryData = { groups: {}, lastUpdated: new Date().toISOString() };
    const since = Date.now() - 24 * 60 * 60 * 1000;

    for (const groupName of GROUPS) {
        const recent = (stored[groupName] || []).filter(m => m.ts > since);
        if (!recent.length) { summaryData.groups[groupName] = 'Geen nieuwe berichten vandaag.'; continue; }

        const msgText = recent.map(m => `${m.sender}: ${m.body}`).join('\n');
        const res = await anthropic.messages.create({
            model: 'claude-sonnet-4-6',
            max_tokens: 512,
            system: 'Vat de volgende WhatsApp-berichten kort samen in het Nederlands. Gebruik maximaal 5 korte bulletpoints.',
            messages: [{ role: 'user', content: msgText }]
        });
        summaryData.groups[groupName] = res.content[0].text;
    }

    saveSummaries(summaryData);

    const text = GROUPS.map(g => `*${g}*\n${summaryData.groups[g]}`).join('\n\n');
    await client.sendMessage(myNumber, `📋 *Dagelijkse samenvatting*\n\n${text}`);
}

const waHistories = {};

function buildScheduleFromRow(dates, shiftRows, nicoIdx) {
    const schedule = {};
    const nicoRow = shiftRows[nicoIdx];
    if (!nicoRow) return schedule;

    dates.forEach((dateStr, i) => {
        const shift = (nicoRow[i] || '').trim();
        if (!shift || !shift.match(/\d{1,2}:\d{2}/)) return;
        const colleagues = shiftRows
            .filter((_, idx) => idx !== nicoIdx)
            .map((row, idx) => ({ name: `Collega ${idx + 1}`, shift: (row[i] || '').trim() }))
            .filter(c => c.shift.match(/\d{1,2}:\d{2}/));
        schedule[dateStr] = { shift, colleagues };
    });
    return schedule;
}

async function parsePDFSchedule(buffer) {
    const base64 = buffer.toString('base64');

    const response = await anthropic.messages.create({
        model: 'claude-sonnet-4-6',
        max_tokens: 2048,
        messages: [{
            role: 'user',
            content: [
                {
                    type: 'document',
                    source: { type: 'base64', media_type: 'application/pdf', data: base64 }
                },
                {
                    type: 'text',
                    text: `Lees dit werkrooster zorgvuldig. Elke kolom is één dag. Koppel elke dienst aan de exacte datum die bovenaan die kolom staat.

Extraheer ALLEEN de diensten van "Nico". Output ALLEEN geldige JSON:
{
  "2026-04-27": {"shift": "12:00-sluit", "colleagues": [{"name": "Lisa", "shift": "9:30-17:00"}]},
  "2026-04-29": {"shift": "17:00-sluit", "colleagues": [{"name": "Tom", "shift": "12:00-17:00"}]}
}

Regels:
- Kijk per kolom welke datum erboven staat en welke dienst Nico in die kolom heeft
- Voeg alleen dagen toe waarop Nico écht een dienst heeft
- Datums in YYYY-MM-DD formaat
- Collega's zijn andere mensen die op DEZELFDE dag (kolom) werken
- Gebruik tijden exact zoals ze in het rooster staan`
                }
            ]
        }]
    });

    const jsonStr = response.content[0].text.match(/\{[\s\S]*\}/)?.[0];
    if (!jsonStr) throw new Error('Claude kon het rooster niet lezen');
    const parsed = JSON.parse(jsonStr);
    console.log('✅ Claude las rooster:', JSON.stringify(parsed, null, 2));
    return parsed;
}

async function handleExcelFromKevin(msg) {
    if (!msg.hasMedia) return false;
    const fname = (msg._data?.filename || '').toLowerCase();
    const mime = msg._data?.mimetype || '';
    const isExcel = fname.match(/\.(xlsx|xls|ods)$/) || mime.includes('spreadsheet') || mime.includes('excel') || mime.includes('officedocument');
    const isPDF = fname.endsWith('.pdf') || mime.includes('pdf');
    if (!isExcel && !isPDF) return false;

    try {
        const media = await msg.downloadMedia();
        if (!media) return false;
        const buffer = Buffer.from(media.data, 'base64');
        const filename = msg._data?.filename || `rooster_${new Date().toLocaleDateString('nl-NL')}`;

        if (isPDF) {
            const parsed = await parsePDFSchedule(buffer);
            if (!parsed) throw new Error('Geen roosterdata gevonden');
            const existing = loadSchedule() || {};
            const parsedSchedules = existing.parsedSchedules || {};
            Object.assign(parsedSchedules, parsed);
            saveSchedule({ ...existing, parsedSchedules, updatedAt: new Date().toISOString() });
        } else {
            const rows = parseExcelRows(buffer);
            const existing = getRowsets();
            const updated = [...existing.filter(r => r.filename !== filename), { rows, filename, date: Date.now() }].slice(-3);
            saveSchedule({ rowsets: updated, updatedAt: new Date().toISOString() });
        }

        console.log(`✅ Rooster opgeslagen: ${filename}`);
        return true;
    } catch (e) {
        console.error('Rooster fout:', e.message);
        return false;
    }
}

client.on('message_create', async (msg) => {
    // Store incoming group messages for summaries
    if (!msg.fromMe && msg.from.includes('@g.us') && msg.body) {
        try {
            const chat = await msg.getChat();
            const matched = matchGroup(chat.name || '');
            if (matched) {
                const contact = await msg.getContact();
                const senderName = contact.name || contact.pushname || 'Onbekend';
                storeGroupMessage(matched, senderName, msg.body);
                console.log(`[opgeslagen] ${matched} — ${senderName}: ${msg.body.slice(0, 50)}`);
            }
        } catch (e) { console.error('[message_create groep fout]', e.message); }
    }

    if (!msg.fromMe || !msg.hasMedia) return;
    const fname = (msg._data?.filename || '').toLowerCase();
    const mime = msg._data?.mimetype || '';
    const isRooster = fname.match(/\.(xlsx|xls|ods|pdf)$/) || mime.includes('spreadsheet') || mime.includes('excel') || mime.includes('officedocument') || mime.includes('pdf');
    if (!isRooster) return;
    const saved = await handleExcelFromKevin(msg);
    if (saved) console.log('✅ Rooster ontvangen via eigen bericht');
});

client.on('message', async (msg) => {
    if (msg.from === 'status@broadcast') return;

    const contact = await msg.getContact();
    const contactName = (contact.name || contact.pushname || '').toLowerCase();
    const isKevin = contactName.includes('kevin');

    // Store group messages for summaries
    if (msg.from.includes('@g.us') && msg.body) {
        try {
            const chat = await msg.getChat();
            const matched = matchGroup(chat.name || '');
            console.log(`[groep] "${chat.name}" — opslaan: ${!!matched}`);
            if (matched) {
                const senderName = contact.name || contact.pushname || 'Onbekend';
                storeGroupMessage(matched, senderName, msg.body);
            }
        } catch (e) { console.error('[groep opslaan fout]', e.message); }
    }

    // Catch Excel from Kevin in any chat (private or group)
    if (isKevin && msg.hasMedia) {
        await handleExcelFromKevin(msg);
        if (msg.from.includes('@g.us')) return; // don't reply in groups
        return;
    }

    // Ignore group messages for the assistant
    if (msg.from.includes('@g.us')) return;

    // Skip empty messages
    if (!msg.body || !msg.body.trim()) return;

    const chatId = msg.from;
    if (!waHistories[chatId]) waHistories[chatId] = [];
    waHistories[chatId].push({ role: 'user', content: msg.body });
    if (waHistories[chatId].length > 20) waHistories[chatId] = waHistories[chatId].slice(-20);

    const schedule = parseScheduleData();
    const scheduleContext = Object.keys(schedule).length
        ? `\n\nWerkrooster van Nico:\n${JSON.stringify(schedule, null, 2)}\nVandaag: ${new Date().toLocaleDateString('nl-NL', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' })}`
        : '';

    try {
        const res = await anthropic.messages.create({
            model: 'claude-sonnet-4-6',
            max_tokens: 1024,
            system: `Je bent de persoonlijke assistent van Nico. Spreek altijd Nederlands. Wees kort en vriendelijk.${scheduleContext}`,
            messages: waHistories[chatId]
        });
        const reply = res.content[0].text;
        waHistories[chatId].push({ role: 'assistant', content: reply });
        await msg.reply(reply);
    } catch (e) {
        console.error('WhatsApp Claude fout:', e.message);
    }
});

// ─── API routes ─────────────────────────────────────────────────────────────

app.get('/api/recurring', (req, res) => {
    res.json(generateRecurringDates());
});

app.get('/api/hva', (req, res) => {
    res.json(loadHva());
});

app.post('/api/hva/refresh', async (req, res) => {
    await fetchHvaSchedule();
    res.json(loadHva());
});

app.get('/api/pdf-rows', (req, res) => {
    const data = loadSchedule();
    if (!data?.pdfRaw) return res.json({ rows: null });
    res.json({ rows: data.pdfRaw.shiftRows, dates: data.pdfRaw.dates, nicoRowIndex: data.nicoRowIndex });
});

app.post('/api/pdf-rows/select', (req, res) => {
    const { rowIndex } = req.body;
    if (rowIndex === undefined) return res.status(400).json({ error: 'rowIndex vereist' });
    const data = loadSchedule() || {};
    data.nicoRowIndex = rowIndex;
    saveSchedule(data);
    res.json({ ok: true });
});

app.get('/api/status', (req, res) => {
    res.json({ status: waStatus, qr: waQR });
});

app.get('/api/qr', async (req, res) => {
    if (!waQR) return res.status(404).json({ error: 'Geen QR beschikbaar' });
    const dataUrl = await QRCode.toDataURL(waQR, { width: 300, margin: 2 });
    const base64 = dataUrl.replace(/^data:image\/png;base64,/, '');
    res.setHeader('Content-Type', 'image/png');
    res.send(Buffer.from(base64, 'base64'));
});

app.get('/api/whatsapp/groups', async (req, res) => {
    try {
        const chats = await client.getChats();
        const groups = chats.filter(c => c.isGroup).map(c => c.name).sort();
        res.json(groups);
    } catch (e) { res.status(500).json({ error: e.message }); }
});

app.get('/api/schedule', (req, res) => {
    const schedule = parseScheduleData();
    const keys = Object.keys(schedule).sort();

    // Determine available weeks
    const weeks = [...new Set(keys.map(k => {
        const d = new Date(k);
        return `${d.getFullYear()}-W${String(getISOWeek(d)).padStart(2, '0')}`;
    }))];

    const reqWeek = req.query.week;
    let targetWeek = reqWeek;

    if (!targetWeek) {
        const now = new Date();
        const curWeekStr = `${now.getFullYear()}-W${String(getISOWeek(now)).padStart(2, '0')}`;
        // Use current week if it has data, else first available
        targetWeek = weeks.includes(curWeekStr) ? curWeekStr : (weeks[0] || curWeekStr);
    }

    const [yr, wk] = targetWeek.split('-W');
    const weekDates = getWeekDates(parseInt(yr), parseInt(wk));
    const days = weekDates.map(d => {
        const key = toDateKey(d);
        return { date: key, dayName: d.toLocaleDateString('nl-NL', { weekday: 'short' }), dayNum: d.getDate(), month: d.getMonth() + 1, ...(schedule[key] || {}) };
    });

    res.json({ week: targetWeek, weeks, days });
});

app.get('/api/summaries', (req, res) => {
    res.json(loadSummaries());
});

app.post('/api/summaries/refresh', async (req, res) => {
    if (waStatus !== 'connected') return res.status(503).json({ error: 'WhatsApp niet verbonden' });
    try {
        const stored = loadMessages();
        const since = Date.now() - 24 * 60 * 60 * 1000;
        const summaryData = { groups: {}, lastUpdated: new Date().toISOString() };

        for (const groupName of GROUPS) {
            const recent = (stored[groupName] || []).filter(m => m.ts > since);
            if (!recent.length) { summaryData.groups[groupName] = 'Geen nieuwe berichten vandaag.'; continue; }

            const msgText = recent.map(m => `${m.sender}: ${m.body}`).join('\n');
            const resp = await anthropic.messages.create({
                model: 'claude-sonnet-4-6',
                max_tokens: 512,
                system: 'Vat de volgende WhatsApp-berichten kort samen in het Nederlands. Gebruik maximaal 5 korte bulletpoints.',
                messages: [{ role: 'user', content: msgText }]
            });
            summaryData.groups[groupName] = resp.content[0].text;
        }

        saveSummaries(summaryData);
        res.json(summaryData);
    } catch (e) {
        console.error('[summaries/refresh fout]', e.stack || e.message);
        res.status(500).json({ error: e.message });
    }
});

app.post('/api/schedule/upload', upload.single('file'), (req, res) => {
    if (!req.file) return res.status(400).json({ error: 'Geen bestand' });
    try {
        const rows = parseExcelRows(req.file.buffer);
        const existing = getRowsets();
        const filename = req.file.originalname || 'upload.xlsx';
        const updated = [...existing.filter(r => r.filename !== filename), { rows, filename, date: Date.now() }].slice(-3);
        saveSchedule({ rowsets: updated, updatedAt: new Date().toISOString() });
        res.json({ ok: true, message: 'Rooster opgeslagen!' });
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

async function getChatsWithRetry() {
    for (let i = 0; i < 3; i++) {
        try { return await client.getChats(); }
        catch (e) {
            if (i === 2) throw e;
            await new Promise(r => setTimeout(r, 3000));
        }
    }
}

app.post('/api/schedule/request-from-kevin', async (req, res) => {
    if (waStatus !== 'connected') return res.status(503).json({ error: 'WhatsApp niet verbonden' });
    try {
        const chats = await getChatsWithRetry();

        let kevinChat = null;
        for (const chat of chats.filter(c => !c.isGroup)) {
            try {
                const contact = await chat.getContact();
                const name = (contact.name || contact.pushname || chat.name || '').toLowerCase();
                if (name.includes('kevin')) { kevinChat = chat; break; }
            } catch {}
        }
        if (!kevinChat) kevinChat = chats.find(c => !c.isGroup && (c.name || '').toLowerCase().includes('kevin'));
        if (!kevinChat) return res.status(404).json({ error: 'Geen chat met Kevin gevonden' });

        await client.sendMessage(kevinChat.id._serialized, 'Hey Kevin, kun je de laatste roosters nog een keer sturen? 🙏');
        res.json({ ok: true });
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

async function fetchRoosterFromKevin() {
    const chats = await getChatsWithRetry();
    const privateChats = chats.filter(c => !c.isGroup);

    let kevinChat = null;
    for (const chat of privateChats) {
        try {
            const contact = await chat.getContact();
            const name = (contact.name || contact.pushname || chat.name || '').toLowerCase();
            if (name.includes('kevin')) { kevinChat = chat; break; }
        } catch {}
    }
    if (!kevinChat) kevinChat = privateChats.find(c => (c.name || '').toLowerCase().includes('kevin')) || null;
    if (!kevinChat) throw new Error('Geen chat met Kevin gevonden');

    let messages;
    for (let attempt = 0; attempt < 4; attempt++) {
        try {
            if (attempt > 0) await new Promise(r => setTimeout(r, attempt * 3000));
            messages = await kevinChat.fetchMessages({ limit: 100 });
            break;
        } catch (e) {
            if (attempt === 3) throw new Error('WhatsApp kon de berichten niet laden');
        }
    }

    const roosterMsgs = messages.filter(m => {
        if (!m.hasMedia) return false;
        const fname = (m._data?.filename || '').toLowerCase();
        const mime = m._data?.mimetype || '';
        return fname.match(/\.(xlsx|xls|ods|pdf)$/) ||
            mime.includes('spreadsheet') || mime.includes('excel') ||
            mime.includes('officedocument') || mime.includes('pdf');
    });

    if (!roosterMsgs.length) throw new Error('Geen roosterbestanden gevonden in chat met Kevin');

    const lastThree = roosterMsgs.slice(-3);
    const rowsets = [];
    const parsedSchedules = {};

    for (const msg of lastThree) {
        try {
            const media = await msg.downloadMedia();
            if (!media) continue;
            const buffer = Buffer.from(media.data, 'base64');
            const fname = (msg._data?.filename || '').toLowerCase();
            const mime = msg._data?.mimetype || '';
            const isPDF = fname.endsWith('.pdf') || mime.includes('pdf');
            const filename = msg._data?.filename || `rooster_${new Date(msg.timestamp * 1000).toLocaleDateString('nl-NL')}`;

            if (isPDF) {
                const parsed = await parsePDFSchedule(buffer);
                if (parsed) Object.assign(parsedSchedules, parsed);
                console.log(`✅ PDF verwerkt: ${filename}`);
            } else {
                const rows = parseExcelRows(buffer);
                rowsets.push({ rows, filename, date: msg.timestamp * 1000 });
                console.log(`✅ Excel verwerkt: ${filename}`);
            }
        } catch (e) {
            console.error('Download/parse fout:', e.message);
        }
    }

    if (!rowsets.length && !Object.keys(parsedSchedules).length) throw new Error('Kon bestanden niet downloaden of verwerken');

    const existing = loadSchedule() || {};
    saveSchedule({ ...existing, rowsets, parsedSchedules: { ...(existing.parsedSchedules || {}), ...parsedSchedules }, updatedAt: new Date().toISOString() });
    console.log(`✅ Roosters opgehaald van Kevin: ${lastThree.length} bestand(en)`);
    return { count: lastThree.length, files: lastThree.map(m => m._data?.filename || 'onbekend') };
}

app.post('/api/schedule/fetch-from-kevin', async (req, res) => {
    if (waStatus !== 'connected') return res.status(503).json({ error: 'WhatsApp niet verbonden' });
    try {
        const result = await fetchRoosterFromKevin();
        res.json({ ok: true, ...result });
    } catch (e) {
        console.error(e);
        res.status(500).json({ error: e.message });
    }
});

app.post('/api/chat', async (req, res) => {
    const { message } = req.body;
    if (!message) return res.status(400).json({ error: 'Geen bericht' });

    webChatHistory.push({ role: 'user', content: message });
    if (webChatHistory.length > 30) webChatHistory.splice(0, webChatHistory.length - 30);

    const schedule = parseScheduleData();
    const scheduleContext = Object.keys(schedule).length
        ? `\n\nWerkrooster van Nico:\n${JSON.stringify(schedule, null, 2)}\nVandaag: ${new Date().toLocaleDateString('nl-NL', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' })}`
        : '';

    const response = await anthropic.messages.create({
        model: 'claude-sonnet-4-6',
        max_tokens: 1024,
        system: `Je bent de persoonlijke assistent van Nico. Spreek altijd Nederlands. Wees kort en vriendelijk.${scheduleContext}`,
        messages: webChatHistory
    });

    const reply = response.content[0].text;
    webChatHistory.push({ role: 'assistant', content: reply });
    res.json({ reply });
});

// ─── Bunq ────────────────────────────────────────────────────────────────────

const BUNQ_BASE = 'https://api.bunq.com';
const BUNQ_CONTEXT_FILE  = path.join(__dirname, 'bunq_context.json');
const BUNQ_CACHE_FILE    = path.join(__dirname, 'bunq_cache.json');
const INVESTMENTS_FILE   = path.join(__dirname, 'investments.json');

function loadBunqContext() {
    if (fs.existsSync(BUNQ_CONTEXT_FILE)) return JSON.parse(fs.readFileSync(BUNQ_CONTEXT_FILE, 'utf8'));
    return null;
}
function saveBunqContext(d) { fs.writeFileSync(BUNQ_CONTEXT_FILE, JSON.stringify(d, null, 2)); }
function loadBunqCache() {
    if (fs.existsSync(BUNQ_CACHE_FILE)) return JSON.parse(fs.readFileSync(BUNQ_CACHE_FILE, 'utf8'));
    return null;
}
function saveBunqCache(d) { fs.writeFileSync(BUNQ_CACHE_FILE, JSON.stringify(d, null, 2)); }

function bunqSign(body, privateKey) {
    const sign = crypto.createSign('SHA256');
    sign.update(typeof body === 'string' ? body : JSON.stringify(body));
    return sign.sign(privateKey, 'base64');
}

function bunqHeaders(token, privateKey, body = '') {
    const reqId = crypto.randomUUID();
    const headers = {
        'Content-Type': 'application/json',
        'User-Agent': 'NicoAssistent/1.0',
        'X-Bunq-Language': 'nl_NL',
        'X-Bunq-Region': 'nl_NL',
        'X-Bunq-Geolocation': '0 0 0 0 000',
        'X-Bunq-Client-Request-Id': reqId,
    };
    if (token) headers['X-Bunq-Client-Authentication'] = token;
    if (privateKey && body) headers['X-Bunq-Client-Signature'] = bunqSign(body, privateKey);
    return headers;
}

async function bunqSetup() {
    const apiKey = process.env.BUNQ_API_KEY;
    if (!apiKey) throw new Error('BUNQ_API_KEY niet ingesteld');

    // Generate RSA key pair
    const { privateKey, publicKey } = crypto.generateKeyPairSync('rsa', {
        modulusLength: 2048,
        publicKeyEncoding: { type: 'spki', format: 'pem' },
        privateKeyEncoding: { type: 'pkcs8', format: 'pem' },
    });

    // 1. Installation
    const installBody = JSON.stringify({ client_public_key: publicKey });
    const installRes = await axios.post(`${BUNQ_BASE}/v1/installation`, installBody, {
        headers: bunqHeaders(null, null, installBody)
    });
    const installToken = installRes.data.Response.find(r => r.Token)?.Token?.token;
    const serverPublicKey = installRes.data.Response.find(r => r.ServerPublicKey)?.ServerPublicKey?.server_public_key;

    // 2. Device server
    const deviceBody = JSON.stringify({ description: 'NicoAssistent', secret: apiKey, permitted_ips: ['*'] });
    await axios.post(`${BUNQ_BASE}/v1/device-server`, deviceBody, {
        headers: bunqHeaders(installToken, privateKey, deviceBody)
    });

    // 3. Session server
    const sessionBody = JSON.stringify({ secret: apiKey });
    const sessionRes = await axios.post(`${BUNQ_BASE}/v1/session-server`, sessionBody, {
        headers: bunqHeaders(installToken, privateKey, sessionBody)
    });
    const sessionToken = sessionRes.data.Response.find(r => r.Token)?.Token?.token;
    const userId = sessionRes.data.Response.find(r => r.UserPerson || r.UserCompany || r.UserApiKey);
    const userIdVal = userId?.UserPerson?.id || userId?.UserCompany?.id || userId?.UserApiKey?.id;

    const ctx = { privateKey, publicKey, serverPublicKey, installToken, sessionToken, userId: userIdVal, createdAt: Date.now() };
    saveBunqContext(ctx);
    console.log('✅ Bunq gekoppeld, userId:', userIdVal);
    return ctx;
}

async function getBunqContext() {
    let ctx = loadBunqContext();
    // Session tokens expire after ~3 hours
    if (!ctx || Date.now() - ctx.createdAt > 3 * 60 * 60 * 1000) {
        ctx = await bunqSetup();
    }
    return ctx;
}

async function bunqGet(path) {
    const ctx = await getBunqContext();
    const res = await axios.get(`${BUNQ_BASE}${path}`, {
        headers: bunqHeaders(ctx.sessionToken, ctx.privateKey, '')
    });
    return res.data.Response;
}

async function fetchBunqData() {
    const ctx = await getBunqContext();

    // Get monetary accounts
    const accountsRes = await bunqGet(`/v1/user/${ctx.userId}/monetary-account`);
    const accounts = accountsRes
        .map(r => r.MonetaryAccountBank || r.MonetaryAccountSavings || r.MonetaryAccountJoint)
        .filter(Boolean)
        .filter(a => a.status === 'ACTIVE')
        .map(a => ({
            id: a.id,
            description: a.description,
            balance: parseFloat(a.balance?.value || 0),
            currency: a.balance?.currency || 'EUR',
            type: a.MonetaryAccountSavings ? 'savings' : 'checking'
        }));

    // Get transactions for each account (last 30 days)
    const allTx = [];
    for (const acc of accounts) {
        try {
            const txRes = await bunqGet(`/v1/user/${ctx.userId}/monetary-account/${acc.id}/payment?count=50`);
            const txs = txRes.map(r => r.Payment).filter(Boolean).map(p => ({
                id: p.id,
                accountId: acc.id,
                amount: parseFloat(p.amount?.value || 0),
                currency: p.amount?.currency || 'EUR',
                description: p.description,
                counterparty: p.counterpart_alias?.display_name || p.counterpart_alias?.iban || '?',
                date: p.created?.slice(0, 10),
                type: parseFloat(p.amount?.value) >= 0 ? 'in' : 'out'
            }));
            allTx.push(...txs);
        } catch {}
    }

    // Analyse with Claude
    const insights = await analyzeBunqData(accounts, allTx);
    const result = { accounts, transactions: allTx.slice(0, 100), insights, lastFetch: new Date().toISOString() };
    saveBunqCache(result);
    return result;
}

async function analyzeBunqData(accounts, transactions) {
    if (!transactions.length) return null;
    const totalBalance = accounts.reduce((s, a) => s + a.balance, 0);
    const out = transactions.filter(t => t.type === 'out');
    const inn = transactions.filter(t => t.type === 'in');
    const totalOut = out.reduce((s, t) => s + Math.abs(t.amount), 0);
    const totalIn  = inn.reduce((s, t) => s + t.amount, 0);

    const txSample = transactions.slice(0, 60).map(t =>
        `${t.date} ${t.type === 'out' ? '-' : '+'}€${Math.abs(t.amount).toFixed(2)} ${t.counterparty}: ${t.description}`
    ).join('\n');

    const res = await anthropic.messages.create({
        model: 'claude-sonnet-4-6',
        max_tokens: 1024,
        system: 'Je bent een financieel assistent voor Nico. Analyseer de transacties en geef praktisch inzicht in het Nederlands. Wees concreet en bondig.',
        messages: [{ role: 'user', content: `Totaal saldo: €${totalBalance.toFixed(2)}\nInkomsten: €${totalIn.toFixed(2)}\nUitgaven: €${totalOut.toFixed(2)}\n\nTransacties:\n${txSample}\n\nGeef:\n1. Top 3-5 uitgavencategorieën met bedrag\n2. Opvallende of grote transacties\n3. Korte tip of observatie` }]
    });
    return res.content[0].text;
}

// ─── Gmail API routes ────────────────────────────────────────────────────────

app.get('/api/gmail/status', (req, res) => {
    const token = loadGmailToken();
    res.json({ connected: !!token });
});

app.get('/api/gmail/auth', (req, res) => {
    if (!process.env.GMAIL_CLIENT_ID) return res.status(500).json({ error: 'GMAIL_CLIENT_ID niet ingesteld in .env' });
    const auth = getOAuth2Client();
    const url = auth.generateAuthUrl({ access_type: 'offline', scope: GMAIL_SCOPES, prompt: 'consent' });
    res.json({ url });
});

app.get('/api/gmail/callback', async (req, res) => {
    const { code } = req.query;
    if (!code) return res.status(400).send('Geen code ontvangen');
    try {
        const auth = getOAuth2Client();
        const { tokens } = await auth.getToken(code);
        saveGmailToken(tokens);
        res.send('<script>window.close();</script><p>✅ Gmail verbonden! Je kunt dit venster sluiten.</p>');
    } catch (e) {
        res.status(500).send('❌ Authenticatie mislukt: ' + e.message);
    }
});

app.get('/api/gmail/mails', async (req, res) => {
    const cache = loadMailCache();
    res.json(cache);
});

app.post('/api/gmail/fetch', async (req, res) => {
    if (!loadGmailToken()) return res.status(401).json({ error: 'Niet ingelogd bij Gmail' });
    try {
        const result = await fetchAndClassifyMails();
        res.json(result);
    } catch (e) {
        console.error('Gmail fetch fout:', e.message);
        res.status(500).json({ error: e.message });
    }
});

app.post('/api/gmail/daily', async (req, res) => {
    if (!loadGmailToken()) return res.json({ skipped: true });
    try {
        const result = await fetchAndClassifyMails();
        res.json(result);
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

// ─── Investment tracker ───────────────────────────────────────────────────────

function loadInvestments() {
    if (fs.existsSync(INVESTMENTS_FILE)) return JSON.parse(fs.readFileSync(INVESTMENTS_FILE, 'utf8'));
    return { positions: [], deposits: [], lastIndex: null, lastIndexDate: null };
}
function saveInvestments(d) { fs.writeFileSync(INVESTMENTS_FILE, JSON.stringify(d, null, 2)); }

async function fetchSP500() {
    try {
        const res = await axios.get('https://query1.finance.yahoo.com/v8/finance/chart/%5EGSPC?interval=1d&range=5d', {
            headers: { 'User-Agent': 'Mozilla/5.0' }
        });
        const closes = res.data.chart.result[0].indicators.quote[0].close.filter(Boolean);
        const timestamps = res.data.chart.result[0].timestamp;
        const price = closes[closes.length - 1];
        const date  = new Date(timestamps[timestamps.length - 1] * 1000).toISOString().split('T')[0];
        const prevClose = closes.length > 1 ? closes[closes.length - 2] : price;
        const dayChange = ((price - prevClose) / prevClose) * 100;
        return { price, date, dayChange };
    } catch (e) {
        console.error('S&P 500 fetch fout:', e.message);
        return null;
    }
}

async function getInvestmentStatus() {
    const data = loadInvestments();
    const sp500 = await fetchSP500();

    // Detect NEW round-up deposits (excluding IDs already baked into snapshot)
    const excludedIds = new Set(data.positions[0]?.excludedIds || []);
    const knownIds = new Set([...excludedIds, ...data.deposits.map(d => d.id)]);
    const bunq = loadBunqCache();
    if (bunq?.transactions) {
        const roundups = bunq.transactions.filter(t =>
            t.type === 'out' && Math.abs(t.amount) < 1.0 &&
            !t.counterparty?.trim().replace('?','') &&
            !knownIds.has(t.id)
        );
        for (const t of roundups) {
            data.deposits.push({ id: t.id, amount: Math.abs(t.amount), date: t.date });
            knownIds.add(t.id);
        }
    }

    // Snapshot value + only NEW deposits on top
    const snapshotValue  = data.positions.reduce((s, p) => s + p.invested, 0);
    const newDeposits    = data.deposits.reduce((s, d) => s + d.amount, 0);
    const totalDeposited = snapshotValue + newDeposits;

    let currentValue = totalDeposited;
    let gainEur = 0, gainPct = 0;

    // If we have an index reference, calculate gain
    if (data.lastIndex && sp500) {
        const indexGain = (sp500.price - data.lastIndex) / data.lastIndex;
        currentValue = totalDeposited * (1 + indexGain);
        gainEur = currentValue - totalDeposited;
        gainPct = indexGain * 100;
    }

    // Update stored index if we got fresh data
    if (sp500 && (!data.lastIndex || sp500.date > (data.lastIndexDate || ''))) {
        if (!data.lastIndex) data.lastIndex = sp500.price; // first time: set as baseline
        data.lastIndexDate = sp500.date;
    }

    saveInvestments(data);

    return {
        totalDeposited,
        currentValue,
        gainEur,
        gainPct,
        sp500,
        deposits: data.deposits.slice().sort((a,b) => b.date?.localeCompare(a.date || '') || 0),
        positions: data.positions
    };
}

// ─── Bunq routes ─────────────────────────────────────────────────────────────

app.get('/api/bunq/status', (req, res) => {
    const ctx = loadBunqContext();
    res.json({ connected: !!ctx });
});

app.post('/api/bunq/setup', async (req, res) => {
    try { await bunqSetup(); res.json({ ok: true }); }
    catch (e) { res.status(500).json({ error: e.message }); }
});

app.get('/api/bunq/data', (req, res) => {
    const cache = loadBunqCache();
    if (!cache) return res.json({ accounts: [], transactions: [], insights: null, lastFetch: null });
    res.json(cache);
});

app.post('/api/bunq/fetch', async (req, res) => {
    try { const data = await fetchBunqData(); res.json(data); }
    catch (e) { console.error('Bunq fout:', e.response?.data || e.message); res.status(500).json({ error: e.message }); }
});

app.get('/api/investments', async (req, res) => {
    try { res.json(await getInvestmentStatus()); }
    catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/investments/set', (req, res) => {
    const { invested, indexAtPurchase } = req.body;
    if (!invested) return res.status(400).json({ error: 'invested vereist' });
    const data = loadInvestments();
    // Set or replace manual position
    data.positions = [{ invested: parseFloat(invested), date: new Date().toISOString().split('T')[0] }];
    if (indexAtPurchase) data.lastIndex = parseFloat(indexAtPurchase);
    saveInvestments(data);
    res.json({ ok: true });
});

// ─── Start ───────────────────────────────────────────────────────────────────

app.listen(3000, () => console.log('App beschikbaar op http://localhost:3000'));
initWhatsApp();
