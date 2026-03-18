require('dotenv').config();
const express = require('express');
const session = require('express-session');
const cors = require('cors');
const path = require('path');
const { db, run, get, all, syncConfig } = require('./database');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// FIXED: serve static files from root (not /public prefix)
app.use(express.static(path.join(__dirname, 'public')));

app.use(session({
    secret: process.env.SESSION_SECRET || 'fallback_secret',
    resave: false,
    saveUninitialized: true,
    cookie: {
        secure: false,
        httpOnly: true,
        maxAge: 24 * 60 * 60 * 1000
    }
}));

const fs = require('fs');
const bcrypt = require('bcryptjs');
const nodemailer = require('nodemailer');
const { buildQuoteHTML } = require('./templates/quoteHTML');

const transporter = nodemailer.createTransport({
    service: 'gmail',
    auth: {
        user: process.env.GMAIL_USER,
        pass: process.env.GMAIL_PASS
    }
});

// ROUTES
app.get('/', (req, res) => res.redirect('/form'));
app.get('/form', (req, res) => res.sendFile(path.join(__dirname, 'public', 'form.html')));
app.get('/dashboard', (req, res) => res.sendFile(path.join(__dirname, 'public', 'dashboard.html')));
app.get('/admin', (req, res) => res.sendFile(path.join(__dirname, 'public', 'admin.html')));

// ADVISOR LOGIN
app.post('/api/login', async (req, res) => {
    const { email, password } = req.body;

    try {
        const advisor = await get('SELECT * FROM advisors WHERE email = ?', [email]);

        if (!advisor) {
            return res.status(401).json({ success: false, message: "Invalid email or password" });
        }

        const match = await bcrypt.compare(password, advisor.password);
        if (!match) {
            return res.status(401).json({ success: false, message: "Invalid email or password" });
        }

        if (advisor.is_active !== 1) {
            return res.status(403).json({ success: false, message: "Account deactivated" });
        }

        req.session.isAuthenticated = true;
        req.session.advisor = { id: advisor.id, name: advisor.name, email: advisor.email };
        res.json({ success: true, advisor: req.session.advisor });

    } catch (err) {
        console.error('Login error:', err);
        res.status(500).json({ success: false, message: "Internal server error" });
    }
});

// DASHBOARD LOGIN (separate from advisor login)
app.post('/api/dashboard-login', (req, res) => {
    const { password } = req.body;
    // Set your desired dashboard password here
    const DASHBOARD_PASSWORD = process.env.DASHBOARD_PASSWORD || 'admin123'; 

    if (password === DASHBOARD_PASSWORD) {
        req.session.isAdmin = true;
        req.session.save((err) => {
            if (err) {
                console.error('Session save error:', err);
                return res.status(500).json({ success: false, message: 'Session save failed' });
            }
            res.json({ success: true });
        });
        return;
    }
    res.status(401).json({ success: false, message: 'Invalid Password' });
});


app.get('/api/dashboard-logout', (req, res) => {
    req.session.destroy();
    res.json({ success: true });
});

app.post('/api/logout', (req, res) => {
    req.session.destroy();
    res.json({ success: true });
});

app.get('/api/me', (req, res) => {
    if (req.session && req.session.advisor) {
        res.json({ loggedIn: true, ...req.session.advisor });
    } else {
        res.json({ loggedIn: false });
    }
});

app.get('/api/config', (req, res) => {
    const configPath = path.join(__dirname, 'config', 'config.json');
    if (fs.existsSync(configPath)) {
        res.sendFile(configPath);
    } else {
        res.status(404).json({ error: "Config not found" });
    }
});

app.get('/ping', (req, res) => res.json({ status: 'alive', ts: Date.now() }));

app.post('/api/save-config', (req, res) => {
    if (!req.session.isAuthenticated && !req.session.isAdmin) return res.status(403).json({ error: "Unauthorized" });
    const configPath = path.join(__dirname, 'config', 'config.json');
    try {
        fs.writeFileSync(configPath, JSON.stringify(req.body, null, 2));
        syncConfig('health', req.body);
        res.json({ success: true });
    } catch (err) {
        res.status(500).json({ success: false, message: "Failed to save configuration" });
    }
});

app.get('/api/quotes', async (req, res) => {
    const apiKey = req.headers['x-rm-key'];
    const validKey = process.env.RM_API_KEY || 'incremint-rm-2026';
    if (!req.session.isAdmin && !req.session.isAuthenticated && apiKey !== validKey) {
        return res.status(401).json({ error: "Unauthorized" });
    }
    try {
        const quotes = await all('SELECT * FROM quotes ORDER BY submitted_at DESC');
        res.json(quotes);
    } catch (err) {
        res.status(500).json({ error: "Failed to fetch quotes" });
    }
});

app.delete('/api/quotes/:id', async (req, res) => {
    if (!req.session.isAdmin) {
        console.warn(`Unauthorized delete attempt on quote ${req.params.id} from ${req.ip}`);
        return res.status(403).json({ error: "Admin access required" });
    }

    try {
        const quote = await get('SELECT quote_id FROM quotes WHERE quote_id = ?', [req.params.id]);
        if (quote) {
            // Delete both .docx and .pdf if they exist
            for (const ext of ['.docx', '.pdf']) {
                const f = path.join(__dirname, 'quotes', `${quote.quote_id}${ext}`);
                if (fs.existsSync(f)) fs.unlinkSync(f);
            }
            await run('DELETE FROM quotes WHERE quote_id = ?', [req.params.id]);
            
            console.log(`[ADMIN ACTION] Quote ${req.params.id} deleted by Admin at ${new Date().toISOString()}`);
            
            res.json({ success: true });
        } else {
            res.status(404).json({ success: false, message: "Quote not found" });
        }
    } catch (err) {
        console.error('Delete error:', err);
        res.status(500).json({ success: false, message: "Failed to delete quote" });
    }
});

app.get('/api/quotes/:id/html', async (req, res) => {
    if (!req.session.isAuthenticated && !req.session.isAdmin) return res.status(401).json({ error: "Unauthorized" });
    try {
        const id = req.params.id;
        // Always regenerate from DB — never serve stale cached file
        const quote = await get('SELECT * FROM quotes WHERE quote_id = ?', [id]);
        if (!quote) return res.status(404).send('Quote not found');
        const dateObj = new Date(quote.submitted_at);
        const dd = String(dateObj.getDate()).padStart(2, '0');
        const mm = String(dateObj.getMonth() + 1).padStart(2, '0');
        const yyyy = dateObj.getFullYear();
        const quoteData = {
            quote_id: quote.quote_id, date: `${dd}-${mm}-${yyyy}`,
            customer_name: quote.customer_name, customer_phone: quote.customer_phone,
            advisor_note: quote.advisor_note,
            members: JSON.parse(quote.members_json || '[]'),
            insurers: JSON.parse(quote.insurers_json || '[]'),
        };
        const html = buildQuoteHTML(quoteData);
        const htmlPath = path.join(__dirname, 'quotes', `${id}.html`);
        fs.writeFileSync(htmlPath, html);
        res.redirect(`/quotes/${id}.html`);
    } catch (err) {
        res.status(500).send('Failed to fetch quote: ' + err.message);
    }
});

// PREVIEW DATA — returns parsed quote JSON for HTML preview modal
app.get('/api/quotes/:id/data', async (req, res) => {
    if (!req.session.isAuthenticated && !req.session.isAdmin) return res.status(401).json({ error: "Unauthorized" });
    try {
        const quote = await get('SELECT * FROM quotes WHERE quote_id = ?', [req.params.id]);
        if (!quote) return res.status(404).json({ error: "Not found" });
        res.json({
            quote_id:       quote.quote_id,
            customer_name:  quote.customer_name,
            customer_phone: quote.customer_phone,
            customer_email: quote.customer_email,
            advisor_name:   quote.advisor_name,
            advisor_note:   quote.advisor_note,
            submitted_at:   quote.submitted_at,
            reviewed:       quote.reviewed,
            members:        JSON.parse(quote.members_json  || '[]'),
            insurers:       JSON.parse(quote.insurers_json || '[]'),
        });
    } catch (err) {
        res.status(500).json({ error: "Failed to fetch quote data" });
    }
});

app.patch('/api/quotes/:id/reviewed', async (req, res) => {
    if (!req.session.isAdmin && !req.session.isAuthenticated) return res.status(401).json({ error: "Unauthorized" });
    try {
        const quote = await get('SELECT reviewed FROM quotes WHERE quote_id = ?', [req.params.id]);
        if (quote) {
            const newStatus = quote.reviewed === 1 ? 0 : 1;
            const statusLabel = newStatus === 1 ? 'Reviewed' : 'New';
            await run('UPDATE quotes SET reviewed = ?, status = ? WHERE quote_id = ?', [newStatus, statusLabel, req.params.id]);
            
            const user = req.session.isAdmin ? 'Admin' : (req.session.advisor?.name || 'Advisor');
            console.log(`[STATUS SYNC] Quote ${req.params.id} marked as ${statusLabel} by ${user}`);
            
            res.json({ success: true, reviewed: newStatus, status: statusLabel });
        } else {
            res.status(404).json({ success: false });
        }
    } catch (err) {
        res.status(500).json({ success: false });
    }
});

app.patch('/api/quotes/:id/status', async (req, res) => {
    if (!req.session.isAdmin && !req.session.isAuthenticated) return res.status(401).json({ error: "Unauthorized" });
    const { status } = req.body;
    const allowedStatuses = ['New', 'Reviewed', 'Flagged', 'In Progress'];
    
    if (!allowedStatuses.includes(status)) {
        return res.status(400).json({ error: "Invalid status" });
    }

    try {
        const reviewed = status === 'Reviewed' ? 1 : 0;
        await run('UPDATE quotes SET status = ?, reviewed = ? WHERE quote_id = ?', [status, reviewed, req.params.id]);
        
        const user = req.session.isAdmin ? 'Admin' : (req.session.advisor?.name || 'Advisor');
        console.log(`[STATUS UPDATE] Quote ${req.params.id} set to "${status}" by ${user}`);
        
        res.json({ success: true, status });
    } catch (err) {
        console.error('Status update error:', err);
        res.status(500).json({ success: false });
    }
});

// EDIT QUOTE + REGENERATE PDF
app.patch('/api/quotes/:id', async (req, res) => {
    if (!req.session.isAuthenticated && !req.session.isAdmin) return res.status(401).json({ error: "Unauthorized" });
    try {
        const quoteId = req.params.id;
        const body = req.body;
        const existing = await get('SELECT * FROM quotes WHERE quote_id = ?', [quoteId]);
        if (!existing) return res.status(404).json({ success: false, message: "Quote not found" });

        const membersJson = JSON.stringify(body.members || []);
        const insurersJson = JSON.stringify(body.insurers || []);

        await run(`UPDATE quotes SET 
            customer_name = ?, customer_phone = ?, customer_email = ?,
            members_json = ?, insurers_json = ?, advisor_note = ?
            WHERE quote_id = ?`,
            [body.customer_name, body.customer_phone, body.customer_email,
             membersJson, insurersJson, body.advisor_note, quoteId]);

        const dateObj = new Date(existing.submitted_at);
        const dd = String(dateObj.getDate()).padStart(2, '0');
        const mm = String(dateObj.getMonth() + 1).padStart(2, '0');
        const yyyy = dateObj.getFullYear();

        // Regenerate HTML file
        const quoteData = { ...body, quote_id: quoteId, date: `${dd}-${mm}-${yyyy}` };
        try {
            const html = buildQuoteHTML(quoteData);
            fs.writeFileSync(path.join(__dirname, 'quotes', `${quoteId}.html`), html);
            console.log(`[HTML] Regenerated for ${quoteId}`);
        } catch (htmlErr) {
            console.error(`[HTML] Failed for ${quoteId}:`, htmlErr.message);
        }

        res.json({ success: true });
    } catch (err) {
        console.error('Edit quote error:', err);
        res.status(500).json({ success: false, message: "Failed to update quote" });
    }
});

// ADVISORS
app.get('/api/advisors', async (req, res) => {
    if (!req.session.isAuthenticated) return res.status(401).json({ error: "Unauthorized" });
    try {
        const advisors = await all('SELECT id, name, email, created_at, is_active FROM advisors ORDER BY name');
        res.json(advisors);
    } catch (err) {
        res.status(500).json({ error: "Failed fetching advisors" });
    }
});

app.post('/api/advisors', async (req, res) => {
    if (!req.session.isAuthenticated) return res.status(401).json({ error: "Unauthorized" });
    try {
        const { name, email, password } = req.body;
        const hash = await bcrypt.hash(password, 10);
        await run('INSERT INTO advisors (name, email, password, is_active, created_at) VALUES (?, ?, ?, 1, ?)',
            [name, email, hash, new Date().toISOString()]);
        res.json({ success: true });
    } catch (err) {
        if (err.message && err.message.includes('UNIQUE'))
            res.status(400).json({ success: false, message: "Email already exists" });
        else
            res.status(500).json({ success: false, message: "Failed adding advisor" });
    }
});

app.patch('/api/advisors/:id/toggle', async (req, res) => {
    if (!req.session.isAuthenticated) return res.status(401).json({ error: "Unauthorized" });
    try {
        const advisor = await get('SELECT is_active FROM advisors WHERE id = ?', [req.params.id]);
        if (advisor) {
            const newStatus = advisor.is_active === 1 ? 0 : 1;
            await run('UPDATE advisors SET is_active = ? WHERE id = ?', [newStatus, req.params.id]);
            res.json({ success: true });
        } else {
            res.status(404).json({ success: false });
        }
    } catch (err) { res.status(500).json({ success: false }); }
});

app.patch('/api/advisors/:id/password', async (req, res) => {
    if (!req.session.isAuthenticated) return res.status(401).json({ error: "Unauthorized" });
    try {
        const hash = await bcrypt.hash(req.body.password, 10);
        await run('UPDATE advisors SET password = ? WHERE id = ?', [hash, req.params.id]);
        res.json({ success: true });
    } catch (err) { res.status(500).json({ success: false }); }
});

app.delete('/api/advisors/:id', async (req, res) => {
    if (!req.session.isAuthenticated) return res.status(401).json({ error: "Unauthorized" });
    try {
        await run('DELETE FROM advisors WHERE id = ?', [req.params.id]);
        res.json({ success: true });
    } catch (err) { res.status(500).json({ success: false }); }
});

app.post('/api/save-settings', (req, res) => {
    if (!req.session.isAuthenticated) return res.status(401).json({ error: "Unauthorized" });
    try {
        const { email, company, password } = req.body;
        if (email) process.env.MASTER_EMAIL = email;
        if (company) process.env.COMPANY_NAME = company;
        if (password) process.env.DASHBOARD_PASSWORD = password;

        const envPath = path.join(__dirname, '.env');
        if (fs.existsSync(envPath)) {
            let content = fs.readFileSync(envPath, 'utf-8');
            if (email) content = content.replace(/MASTER_EMAIL=.*/, `MASTER_EMAIL=${email}`);
            if (company) content = content.replace(/COMPANY_NAME=.*/, `COMPANY_NAME=${company}`);
            if (password) content = content.replace(/DASHBOARD_PASSWORD=.*/, `DASHBOARD_PASSWORD=${password}`);
            fs.writeFileSync(envPath, content);
        }
        res.json({ success: true });
    } catch (err) {
        res.status(500).json({ success: false });
    }
});

// SUBMIT QUOTE
app.post('/api/submit-quote', async (req, res) => {
    try {
        const body = req.body;
        const dateObj = new Date();
        const yyyy = dateObj.getFullYear();
        const mm = String(dateObj.getMonth() + 1).padStart(2, '0');
        const dd = String(dateObj.getDate()).padStart(2, '0');
        const random4 = Math.floor(1000 + Math.random() * 9000);
        const quoteId = `TK${yyyy}${mm}${dd}${random4}`;
        const formattedDate = `${dd}-${mm}-${yyyy}`;

        const advisor = (req.session && req.session.advisor) ? req.session.advisor : null;
        const advisorName = advisor ? advisor.name : null;
        const submittedBy = advisor ? advisor.name : "Customer";

        await run(`INSERT INTO quotes 
            (quote_id, customer_name, customer_phone, customer_email, advisor_id, advisor_name, submitted_by, members_json, insurers_json, advisor_note, submitted_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
            [quoteId, body.customer_name || 'Unknown', body.customer_phone || null,
             body.customer_email || null, advisor?.id || null, advisorName, submittedBy,
             JSON.stringify(body.members || []), JSON.stringify(body.insurers || []),
             body.advisor_note || null, new Date().toISOString()]);

        const quoteData = { ...body, quote_id: quoteId, date: formattedDate };
        const htmlDir = path.join(__dirname, 'quotes');
        if (!fs.existsSync(htmlDir)) fs.mkdirSync(htmlDir, { recursive: true });
        const quotePath = path.join(htmlDir, `${quoteId}.html`);
        try {
            const html = buildQuoteHTML(quoteData);
            fs.writeFileSync(quotePath, html);
        } catch (htmlErr) {
            console.error("HTML generation failed:", htmlErr);
        }

        const members = body.members || [];
        let membersText = members.map(m => {
            const policy = m.policy_type === 'port' ? `Port from ${m.port_details?.current_insurer || '?'}` : 'Fresh';
            const conds = (m.diseases || []).map(d => d.name).join(', ') || 'None';
            const surgs = (m.surgeries || []).map(s => `${s.name} (${s.year})`).join(', ') || 'None';
            return `  • ${m.name} | ${m.relation} | Age ${m.age} | ${m.city} | SA: ${m.sum_assured}\n    Policy: ${policy}\n    Conditions: ${conds}\n    Surgeries: ${surgs}`;
        }).join('\n\n');

        try {
            await transporter.sendMail({
                from: process.env.GMAIL_USER,
                to: process.env.MASTER_EMAIL,
                subject: `New Quote – ${quoteId} – ${body.customer_name || 'N/A'} – ${formattedDate}`,
                text: `New health insurance quote submitted.\n\nQuote ID: ${quoteId}\nDate: ${formattedDate}\nSubmitted by: ${submittedBy}\nCustomer: ${body.customer_name}\nPhone: ${body.customer_phone}\n\nMEMBERS:\n${membersText}\n\nQuote HTML attached.`,
                attachments: [{ filename: `${quoteId}.html`, path: quotePath }]
            });
        } catch (emailErr) {
            console.error("Email sending failed:", emailErr);
        }

        res.json({ success: true, quoteId });

    } catch (err) {
        console.error('Submit quote error:', err);
        res.status(500).json({ success: false, message: 'Failed to process quote' });
    }
});

// GENERATE QUOTE HTML (called by React dashboard quote tool)
app.post('/generate', async (req, res) => {
    try {
        const body = req.body;
        const dateObj = new Date();
        const yyyy = dateObj.getFullYear();
        const mm = String(dateObj.getMonth() + 1).padStart(2, '0');
        const dd = String(dateObj.getDate()).padStart(2, '0');
        const quoteId = `TK${yyyy}${mm}${dd}${Math.floor(1000 + Math.random() * 9000)}`;
        const quoteData = { ...body, quote_id: quoteId, date: `${dd}-${mm}-${yyyy}` };

        // Save to DB so it appears in admin quote list
        try {
            await run(`INSERT INTO quotes
                (quote_id, customer_name, customer_phone, customer_email, advisor_id, advisor_name, submitted_by, members_json, insurers_json, advisor_note, submitted_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
                [quoteId, body.customer_name || 'Unknown', body.customer_phone || null,
                 null, null, body.advisor_name || body.agent_name || 'RM Dashboard', 'RM Dashboard',
                 JSON.stringify(body.members || []), JSON.stringify(body.insurers || []),
                 null, new Date().toISOString()]);
        } catch (dbErr) {
            console.error('DB save error (non-fatal):', dbErr.message);
        }

        // Save HTML file (like term quotes)
        const html = buildQuoteHTML(quoteData);
        const htmlDir = path.join(__dirname, 'quotes');
        if (!fs.existsSync(htmlDir)) fs.mkdirSync(htmlDir, { recursive: true });
        fs.writeFileSync(path.join(htmlDir, `${quoteId}.html`), html);

        res.json({ success: true, quoteId, url: `/quotes/${quoteId}.html` });
    } catch (err) {
        console.error('Generate error:', err);
        res.status(500).json({ success: false, message: 'Failed to generate quote: ' + err.message });
    }
});

// ── TERM QUOTE ROUTES ────────────────────────────────────────────────────────

// Term config
app.get('/api/term-config', (req, res) => {
    try {
        const configPath = path.join(__dirname, 'config', 'term-config.json');
        const config = JSON.parse(fs.readFileSync(configPath, 'utf8'));
        res.json(config);
    } catch (err) {
        console.error('Term config error:', err);
        res.status(500).json({ error: 'Failed to load term config' });
    }
});

// Save term config (admin)
app.post('/api/save-term-config', (req, res) => {
    if (!req.session.isAuthenticated && !req.session.isAdmin) {
        return res.status(403).json({ success: false, message: 'Not authorized' });
    }
    try {
        const configPath = path.join(__dirname, 'config', 'term-config.json');
        fs.writeFileSync(configPath, JSON.stringify(req.body, null, 2));
        syncConfig('term', req.body);
        res.json({ success: true });
    } catch (err) {
        console.error('Save term config error:', err);
        res.status(500).json({ success: false, message: 'Failed to save' });
    }
});

// Generate term quote → returns HTML URL (opens in new tab, user prints to PDF)
app.post('/generate-term', async (req, res) => {
    try {
        const body = req.body;
        const dateObj = new Date();
        const yyyy = dateObj.getFullYear();
        const mm = String(dateObj.getMonth() + 1).padStart(2, '0');
        const dd = String(dateObj.getDate()).padStart(2, '0');
        const quoteId = `TM${yyyy}${mm}${dd}${Math.floor(1000 + Math.random() * 9000)}`;
        const quoteData = { ...body, quote_id: quoteId, date: `${dd}-${mm}-${yyyy}` };

        // Save to DB
        try {
            await run(`INSERT INTO quotes
                (quote_id, customer_name, customer_phone, customer_email, advisor_id, advisor_name, submitted_by, members_json, insurers_json, advisor_note, submitted_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
                [quoteId, body.customer_name || 'Unknown', body.customer_phone || null,
                 null, null, body.advisor_name || body.agent_name || 'RM Dashboard', 'RM Dashboard',
                 JSON.stringify([{ name: body.customer_name, age: body.age, dob: body.dob, city: body.city, sum_assured: body.sum_assured }]),
                 JSON.stringify(body.insurers || []),
                 body.advisor_note || null, new Date().toISOString()]);
        } catch (dbErr) {
            console.error('Term DB save error (non-fatal):', dbErr.message);
        }

        // Save HTML to file for viewing in browser
        const { buildTermQuoteHTML } = require('./templates/termQuoteHTML');
        const html = buildTermQuoteHTML(quoteData);
        const htmlDir = path.join(__dirname, 'quotes');
        if (!fs.existsSync(htmlDir)) fs.mkdirSync(htmlDir, { recursive: true });
        fs.writeFileSync(path.join(htmlDir, `${quoteId}.html`), html);

        res.json({ success: true, quoteId, url: `/quotes/${quoteId}.html` });
    } catch (err) {
        console.error('Term generate error:', err);
        res.status(500).json({ success: false, message: 'Failed to generate term quote: ' + err.message });
    }
});

// Serve generated quote HTML files
app.use('/quotes', express.static(path.join(__dirname, 'quotes')));

app.listen(PORT, () => {
    console.log(`Server running on http://localhost:${PORT}`);

    // ── Keep-alive: self-ping every 13 min, only 9:30 AM - 10:30 PM IST ──
    const RENDER_URL = process.env.RENDER_EXTERNAL_URL || `http://localhost:${PORT}`;
    setInterval(() => {
        const istHour = new Date().toLocaleString('en-US', { timeZone: 'Asia/Kolkata', hour: 'numeric', hour12: false });
        const hr = parseInt(istHour);
        if (hr >= 9 && hr <= 22) { // 9:30 AM buffer to 10:30 PM IST
            fetch(`${RENDER_URL}/ping`).catch(() => {});
            console.log(`[keep-alive] pinged at IST hour ${hr}`);
        }
    }, 13 * 60 * 1000); // 13 minutes
});
