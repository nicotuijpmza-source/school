require('dotenv').config();
const { Client, LocalAuth, MessageMedia } = require('whatsapp-web.js');
const qrcode = require('qrcode-terminal');
const Anthropic = require('@anthropic-ai/sdk');
const cron = require('node-cron');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

const GROUPS = ['Out en nieuw', 'Le Pompadour☕🥂', 'De Petjes (met kleinkinderen)'];
const MY_NAME = 'Nico';
const SCHEDULE_FILE = path.join(__dirname, 'schedule.json');

// Conversation history per chat
const histories = {};

// Load stored schedule
function loadSchedule() {
    if (fs.existsSync(SCHEDULE_FILE)) {
        return JSON.parse(fs.readFileSync(SCHEDULE_FILE, 'utf8'));
    }
    return null;
}

function saveSchedule(data) {
    fs.writeFileSync(SCHEDULE_FILE, JSON.stringify(data, null, 2));
}

// Parse Excel and extract full schedule data
function parseExcel(buffer) {
    const workbook = XLSX.read(buffer, { type: 'buffer', cellDates: true });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    return rows;
}

// Find Nico's row and extract schedule
function extractNicoSchedule(rows) {
    const schedule = [];
    for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        const nicoIdx = row.findIndex(cell =>
            typeof cell === 'string' && cell.toLowerCase().includes(MY_NAME.toLowerCase())
        );
        if (nicoIdx !== -1) {
            // Get header row (first row with dates)
            const headerRow = rows[0];
            for (let j = 0; j < row.length; j++) {
                if (row[j] && j !== nicoIdx) {
                    schedule.push({
                        date: String(headerRow[j] || ''),
                        time: String(row[j]),
                        rowIndex: i,
                        colIndex: j
                    });
                }
            }
        }
    }
    return schedule;
}

const client = new Client({
    authStrategy: new LocalAuth(),
    puppeteer: {
        executablePath: '/usr/bin/google-chrome-stable',
        args: ['--no-sandbox', '--disable-setuid-sandbox']
    }
});

client.on('qr', (qr) => {
    console.log('Scan this QR code with WhatsApp on your phone:');
    qrcode.generate(qr, { small: true });
});

client.on('ready', async () => {
    console.log('WhatsApp connected! Persoonlijke assistent actief.');

    // Daily summary at 18:00
    cron.schedule('0 18 * * *', async () => {
        await sendDailySummary();
    }, { timezone: 'Europe/Amsterdam' });
});

async function sendDailySummary() {
    const chats = await client.getChats();
    const myNumber = client.info.wid._serialized;
    let fullSummary = '📋 *Dagelijkse samenvatting*\n\n';

    for (const groupName of GROUPS) {
        const group = chats.find(c => c.name === groupName);
        if (!group) {
            fullSummary += `*${groupName}*: groep niet gevonden\n\n`;
            continue;
        }

        const messages = await group.fetchMessages({ limit: 100 });
        const since = Date.now() - 24 * 60 * 60 * 1000;
        const recent = messages.filter(m => m.timestamp * 1000 > since && m.body);

        if (recent.length === 0) {
            fullSummary += `*${groupName}*: geen nieuwe berichten\n\n`;
            continue;
        }

        const msgText = recent.map(m => `${m._data.notifyName || 'Onbekend'}: ${m.body}`).join('\n');

        const response = await anthropic.messages.create({
            model: 'claude-sonnet-4-6',
            max_tokens: 512,
            system: 'Vat de volgende WhatsApp-berichten kort samen in het Nederlands. Max 5 bulletpoints.',
            messages: [{ role: 'user', content: msgText }]
        });

        fullSummary += `*${groupName}*\n${response.content[0].text}\n\n`;
    }

    await client.sendMessage(myNumber, fullSummary);
}

client.on('message', async (msg) => {
    if (msg.from === 'status@broadcast') return;

    const contact = await msg.getContact();
    const contactName = contact.name || contact.pushname || '';
    const isKevin = contactName.toLowerCase().includes('kevin');
    const isGroup = msg.from.includes('@g.us');

    // Handle Excel file from Kevin
    if (isKevin && msg.hasMedia) {
        const media = await msg.downloadMedia();
        if (media && (msg.type === 'document' || media.mimetype.includes('spreadsheet') || media.mimetype.includes('excel'))) {
            try {
                const buffer = Buffer.from(media.data, 'base64');
                const rows = parseExcel(buffer);
                const nicoSchedule = extractNicoSchedule(rows);
                saveSchedule({ rows, nicoSchedule, updatedAt: new Date().toISOString() });
                await msg.reply('✅ Rooster opgeslagen! Vraag me wanneer je werkt of met wie.');
            } catch (e) {
                console.error('Excel parse error:', e.message);
                await msg.reply('❌ Kon het bestand niet lezen. Stuur het opnieuw?');
            }
        }
        return;
    }

    // Ignore group messages (only handle personal chats)
    if (isGroup) return;

    // Alleen reageren in de eigen zelfchat (Nico stuurt bericht naar zichzelf)
    const myNumber = client.info.wid._serialized;
    if (msg.from !== myNumber) return;

    const chatId = msg.from;
    if (!histories[chatId]) histories[chatId] = [];

    const lowerBody = msg.body.toLowerCase();
    const isScheduleQuestion = lowerBody.includes('werk') || lowerBody.includes('dienst') ||
        lowerBody.includes('rooster') || lowerBody.includes('tijd') || lowerBody.includes('collega') ||
        lowerBody.includes('met wie') || lowerBody.includes('wanneer');

    let systemPrompt = 'Je bent een persoonlijke assistent van Nico. Spreek altijd Nederlands. Wees kort en vriendelijk.';

    if (isScheduleQuestion) {
        const schedule = loadSchedule();
        if (schedule) {
            const scheduleText = JSON.stringify(schedule.rows);
            systemPrompt += `\n\nHier is het werkrooster (Excel data):\n${scheduleText}\n\nZoek de rij met "Nico" en beantwoord vragen over werktijden en collega's. Vandaag is ${new Date().toLocaleDateString('nl-NL', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' })}.`;
        } else {
            systemPrompt += '\n\nEr is nog geen rooster beschikbaar. Vraag de gebruiker om het rooster van Kevin door te sturen.';
        }
    }

    histories[chatId].push({ role: 'user', content: msg.body });
    if (histories[chatId].length > 20) histories[chatId] = histories[chatId].slice(-20);

    try {
        const response = await anthropic.messages.create({
            model: 'claude-sonnet-4-6',
            max_tokens: 1024,
            system: systemPrompt,
            messages: histories[chatId],
        });

        const reply = response.content[0].text;
        histories[chatId].push({ role: 'assistant', content: reply });
        await msg.reply(reply);
    } catch (err) {
        console.error('Claude error:', err.message);
        await msg.reply('Er ging iets mis, probeer het opnieuw.');
    }
});

client.initialize();
