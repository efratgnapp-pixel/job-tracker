#!/usr/bin/env node
'use strict';

require('dotenv').config();

const http  = require('http');
const https = require('https');
const fs    = require('fs');
const path  = require('path');

const { Document, Packer, Paragraph, TextRun, AlignmentType, BorderStyle } = require('docx');

const PORT    = 3001;
const HTML    = path.join(__dirname, 'job-tracker.html');
const API_KEY = process.env.ANTHROPIC_API_KEY;

// ── CV text parser (shared by docx builder) ────────────────────────────────
function parseCvText(text) {
  const lines = text.split('\n');
  let idx = 0;
  const nextNonEmpty = () => { while (idx < lines.length && !lines[idx].trim()) idx++; return lines[idx++]?.trim() || ''; };

  const name     = nextNonEmpty();
  const subtitle = nextNonEmpty();  // "Project Manager | Business Operations Manager"
  const contact  = nextNonEmpty();  // "City · phone · email · LinkedIn"

  const sections = [];
  let cur = null;
  let profileLines = [];
  let pastProfile = false;

  while (idx < lines.length) {
    const line = lines[idx++].trim();
    if (!line) continue;

    // Section heading: ALL CAPS, ≤6 words, no lowercase, 3–60 chars
    const words = line.split(/\s+/);
    if (words.length <= 6 && line.length >= 3 && line.length <= 60
        && /^[A-Z0-9][A-Z0-9\s&\/\-\.]+$/.test(line) && !/[a-z]/.test(line)) {
      pastProfile = true;
      cur = { heading: line, blocks: [] };
      sections.push(cur);
      continue;
    }
    if (!pastProfile) { profileLines.push(line); continue; }
    if (!cur) continue;

    // Role: starts with //
    if (/^\/\//.test(line)) {
      cur.blocks.push({ type: 'role', text: line }); continue;
    }
    // Company: has | and year, doesn't start with //
    if (/\|/.test(line) && /\d{4}/.test(line)) {
      cur.blocks.push({ type: 'company', text: line }); continue;
    }
    // Bullet
    if (/^[-•*]\s/.test(line)) {
      const t = line.replace(/^[-•*]\s*/, '');
      const last = cur.blocks[cur.blocks.length - 1];
      if (last?.type === 'bullets') last.items.push(t);
      else cur.blocks.push({ type: 'bullets', items: [t] });
      continue;
    }
    cur.blocks.push({ type: 'paragraph', text: line });
  }
  return { name, subtitle, contact, profile: profileLines.join(' '), sections };
}

// ── Word document builder ─────────────────────────────────────────────────
function buildDocx(cvText) {
  const { name, subtitle, contact, profile, sections } = parseCvText(cvText);
  // Exact colours from Efrat's real CV
  const NAVY   = '2F5496';  // name
  const BLUE   = '0070C0';  // subtitle + // prefix
  const DARK   = '262626';  // body text
  const children = [];

  // Name
  children.push(new Paragraph({
    spacing: { after: 60 },
    children: [new TextRun({ text: name, bold: true, size: 52, color: NAVY, allCaps: true })],
  }));
  // Subtitle
  children.push(new Paragraph({
    spacing: { after: 60 },
    children: [new TextRun({ text: subtitle, bold: true, size: 22, color: BLUE })],
  }));
  // Contact
  children.push(new Paragraph({
    spacing: { after: 200 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: 'cccccc', space: 6 } },
    children: [new TextRun({ text: contact, size: 18, color: DARK })],
  }));
  // Profile
  if (profile) {
    children.push(new Paragraph({
      spacing: { before: 120, after: 160 },
      children: [new TextRun({ text: profile, size: 18, color: DARK })],
    }));
  }

  for (const section of sections) {
    children.push(new Paragraph({
      spacing: { before: 240, after: 80 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: 'cccccc', space: 4 } },
      children: [new TextRun({ text: section.heading, bold: true, size: 20, allCaps: true, color: DARK })],
    }));

    for (const block of section.blocks) {
      if (block.type === 'company') {
        children.push(new Paragraph({
          spacing: { before: 140, after: 20 },
          children: [new TextRun({ text: block.text, bold: true, size: 22, color: DARK })],
        }));
      } else if (block.type === 'role') {
        const slashes = block.text.match(/^(\/\/\s*)/)?.[1] || '';
        const rest    = block.text.slice(slashes.length);
        children.push(new Paragraph({
          spacing: { before: 100, after: 30 },
          children: [
            new TextRun({ text: slashes, bold: true, size: 20, color: BLUE }),
            new TextRun({ text: rest,    bold: true, size: 20, color: DARK }),
          ],
        }));
      } else if (block.type === 'bullets') {
        for (const item of block.items) {
          children.push(new Paragraph({
            spacing: { before: 20, after: 20 },
            indent: { left: 320, hanging: 200 },
            children: [new TextRun({ text: `• ${item}`, size: 18, color: DARK })],
          }));
        }
      } else {
        // paragraph — could be bold key-achievement line (first para after role)
        children.push(new Paragraph({
          spacing: { before: 40, after: 40 },
          children: [new TextRun({ text: block.text, size: 18, color: DARK, bold: true })],
        }));
      }
    }
  }

  return new Document({
    sections: [{
      properties: { page: { margin: { top: 900, right: 1080, bottom: 900, left: 1080 } } },
      children,
    }],
  });
}

// ── Cover letter Word document builder ───────────────────────────────────────
function buildCoverLetterDocx(text) {
  const lines = text.split('\n');
  let idx = 0;
  const nextNonEmpty = () => { while (idx < lines.length && !lines[idx].trim()) idx++; return lines[idx++]?.trim() || ''; };

  const name    = nextNonEmpty();
  const contact = nextNonEmpty();
  const date    = nextNonEmpty();

  const NAVY = '2F5496';
  const DARK = '262626';
  const children = [];

  children.push(new Paragraph({
    spacing: { after: 60 },
    children: [new TextRun({ text: name, bold: true, size: 52, color: NAVY, allCaps: true })],
  }));
  children.push(new Paragraph({
    spacing: { after: 200 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: 'cccccc', space: 6 } },
    children: [new TextRun({ text: contact, size: 18, color: DARK })],
  }));
  children.push(new Paragraph({
    spacing: { before: 220, after: 260 },
    children: [new TextRun({ text: date, size: 18, color: DARK })],
  }));

  // Collect body paragraphs (blank-line delimited)
  const paras = [];
  let cur = [];
  while (idx < lines.length) {
    const line = lines[idx++].trim();
    if (!line) { if (cur.length) { paras.push(cur.join(' ')); cur = []; } }
    else cur.push(line);
  }
  if (cur.length) paras.push(cur.join(' '));

  for (const para of paras) {
    children.push(new Paragraph({
      spacing: { before: 80, after: 160 },
      children: [new TextRun({ text: para, size: 20, color: DARK })],
    }));
  }

  return new Document({
    sections: [{
      properties: { page: { margin: { top: 900, right: 1080, bottom: 900, left: 1080 } } },
      children,
    }],
  });
}

// ── Indeed jobs (fetched via Indeed MCP, refreshed when server restarts) ──────
const INDEED_JOBS = [
  { title: 'Engineering Project Manager',                          company: 'The Bread Factory',                 location: 'London',         posted: 'April 13, 2026',  url: 'https://to.indeed.com/aavs64fzyr49', searchQuery: 'Project Manager' },
  { title: 'Project Manager M365 (English and German Speaking)',   company: 'Huxley',                            location: 'London',         posted: 'March 30, 2026',  url: 'https://to.indeed.com/aavhjbypd9js', searchQuery: 'Project Manager' },
  { title: 'Project Delivery Manager',                             company: 'The Guinness Partnership',          location: 'London',         posted: 'May 6, 2026',     url: 'https://to.indeed.com/aajhqrqtcy8x', searchQuery: 'Project Manager' },
  { title: 'Digital Project Manager (AI-enabled delivery)',        company: 'Blackbridge',                       location: 'Spitalfields',   posted: 'March 27, 2026', url: 'https://to.indeed.com/aa69d6npw6hj', searchQuery: 'Project Manager' },
  { title: 'Project Manager',                                      company: 'Imperial College London',           location: 'White City',     posted: 'April 22, 2026',  url: 'https://to.indeed.com/aayhky97wsks', searchQuery: 'Project Manager' },
  { title: 'Project Delivery Manager',                             company: 'Xantura',                           location: 'London',         posted: 'April 13, 2026',  url: 'https://to.indeed.com/aadpyvf6bqqf', searchQuery: 'Project Manager' },
  { title: 'Technical Project Manager - Operational Support Unit', company: 'Metropolitan Police',               location: 'Sydenham',       posted: 'May 11, 2026',    url: 'https://to.indeed.com/aaq6vmjrljxx', searchQuery: 'Project Manager' },
  { title: 'PMO Manager',                                          company: 'Huxley',                            location: 'London',         posted: 'April 15, 2026',  url: 'https://to.indeed.com/aaqcxg64dynf', searchQuery: 'Project Manager' },
  { title: 'Business (finance) Project Manager',                   company: 'Sinomax International Ltd',         location: 'London',         posted: 'April 16, 2026',  url: 'https://to.indeed.com/aacyrths4744', searchQuery: 'Project Manager' },
  { title: 'Senior Project Manager',                               company: 'CGI',                               location: 'London',         posted: 'April 10, 2026',  url: 'https://to.indeed.com/aa4c7blsh68q', searchQuery: 'Project Manager' },
  { title: 'General Manager Service & Operations',                 company: 'Fortnum & Mason',                   location: 'London',         posted: 'April 30, 2026',  url: 'https://to.indeed.com/aa47wcfxw28r', searchQuery: 'Operations Manager' },
  { title: 'Settlement Operations Manager',                        company: 'SEACHANGE FINANCIAL SERVICES LIMITED', location: 'London',      posted: 'May 5, 2026',     url: 'https://to.indeed.com/aa2ttv7z74cn', searchQuery: 'Operations Manager' },
  { title: 'Head of Operations (Tech Startup)',                    company: 'NextStep',                          location: 'London',         posted: 'February 12, 2026', url: 'https://to.indeed.com/aawkkdx6kfmb', searchQuery: 'Operations Manager' },
  { title: 'Digital Operations Manager',                           company: 'Michael Kors',                      location: 'London',         posted: 'May 10, 2026',    url: 'https://to.indeed.com/aaqklzh9cjl8', searchQuery: 'Operations Manager' },
  { title: 'Operations Manager',                                   company: "Megan's",                           location: 'London',         posted: 'April 29, 2026',  url: 'https://to.indeed.com/aavrmc9fvq94', searchQuery: 'Operations Manager' },
  { title: 'Operations Manager',                                   company: 'Mammafiore',                        location: 'West London',    posted: 'April 24, 2026',  url: 'https://to.indeed.com/aaxkqwzgyzfk', searchQuery: 'Operations Manager' },
  { title: 'Property Operations Manager',                          company: 'Ziser London',                      location: 'West Hampstead', posted: 'May 7, 2026',     url: 'https://to.indeed.com/aapjztzl2xyp', searchQuery: 'Operations Manager' },
  { title: 'Operations Manager',                                   company: 'Chevron Traffic Management',        location: 'London',         posted: 'April 17, 2026',  url: 'https://to.indeed.com/aambdwrgmtcw', searchQuery: 'Operations Manager' },
  { title: 'Chief Finance & Operations Officer (CFOO)',            company: 'ANZUK Group',                       location: 'Lambeth',        posted: 'April 21, 2026',  url: 'https://to.indeed.com/aaxjc4ktw6gl', searchQuery: 'Operations Manager' },
  { title: 'Trading Operations Manager',                           company: 'bp',                                location: 'London',         posted: 'April 22, 2026',  url: 'https://to.indeed.com/aag4lsm8vmrd', searchQuery: 'Operations Manager' },
];

function readBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    req.on('data', c => { chunks.push(c); });
    req.on('end',  () => resolve(Buffer.concat(chunks).toString('utf8')));
    req.on('error', reject);
  });
}

function httpsPost(options, body) {
  return new Promise((resolve, reject) => {
    const req = https.request(options, res => {
      const chunks = [];
      res.on('data', c => chunks.push(c));
      res.on('end', () => resolve({ status: res.statusCode, body: Buffer.concat(chunks).toString('utf8') }));
    });
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

function send(res, status, obj) {
  const body = JSON.stringify(obj);
  res.writeHead(status, {
    'Content-Type': 'application/json',
    'Access-Control-Allow-Origin': 'null',   // file:// origin
    'Access-Control-Allow-Headers': 'Content-Type',
  });
  res.end(body);
}

// ── server ─────────────────────────────────────────────────────────────────

const server = http.createServer(async (req, res) => {
  // Permissive CORS so the page works when opened as a local file
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET,POST,OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') { res.writeHead(204); res.end(); return; }

  const url      = new URL(req.url, `http://localhost:${PORT}`);
  const pathname = url.pathname;

  // ── GET / or /job-tracker.html → serve the app ──────────────────────────
  if (req.method === 'GET' && (pathname === '/' || pathname === '/job-tracker.html')) {
    try {
      const html = fs.readFileSync(HTML);
      res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
      res.end(html);
    } catch {
      res.writeHead(404); res.end('job-tracker.html not found next to server.js');
    }
    return;
  }

  // ── GET /api/health ──────────────────────────────────────────────────────
  if (req.method === 'GET' && pathname === '/api/health') {
    send(res, 200, { ok: true, hasKey: !!API_KEY });
    return;
  }

  // ── GET /api/indeed-jobs ─────────────────────────────────────────────────
  if (req.method === 'GET' && pathname === '/api/indeed-jobs') {
    send(res, 200, { jobs: INDEED_JOBS });
    return;
  }

  // ── POST /api/match-score ─────────────────────────────────────────────────
  if (req.method === 'POST' && pathname === '/api/match-score') {
    if (!API_KEY) { send(res, 500, { error: 'ANTHROPIC_API_KEY is not set' }); return; }

    let raw;
    try { raw = await readBody(req); } catch { send(res, 400, { error: 'Failed to read request body' }); return; }

    let payload;
    try { payload = JSON.parse(raw); } catch { send(res, 400, { error: 'Request body is not valid JSON' }); return; }

    const { role, company, notes, jobDescription } = payload;
    if (!role || !company) { send(res, 400, { error: 'Request must include role and company' }); return; }

    const jobContext = jobDescription
      ? `Job Description:\n${jobDescription}${notes ? `\n\nAdditional context: ${notes}` : ''}`
      : `Additional context: ${notes || 'None provided'}`;

    const prompt = `You are a career coach evaluating a candidate against a job listing.

CANDIDATE PROFILE:
Name: Efrat Gnapp
Title: Project Manager / Business Operations Manager
Location: NW8 7BS, London — UK citizen, right to work
Languages: English (fluent), Hebrew (native)

Profile summary: Project Manager with 5+ years managing complex technical implementations and onboarding projects in a global technology organisation. Proven track record delivering SaaS-style implementations end-to-end, managing cross-functional teams, and ensuring successful adoption by 100+ users. Skilled at translating client requirements into executable technical plans while maintaining strong stakeholder relationships at all levels.

Key achievements:
- Led cross-functional delivery of three product programs simultaneously (total budget ~$40M), coordinating R&D, QA, Marketing, and Operations; all milestones delivered on time and within budget including Drupa 2024 launches
- Improved programme planning and delivery efficiency by introducing Lean practices and a global forecasting model across EMEA, APJ, and North America
- Delivered organisation-wide Jira implementation across Operations — end-to-end planning, 5-year legacy data migration, adoption by 100+ engineers

Work experience — HP Indigo, Feb 2019 – Aug 2024:
- Project Manager R&D (Sep 2022 – Aug 2024): led three concurrent product programs (~$40M annual budget), cross-functional delivery, Lean methodology, executive reporting, risk management, technical integrations
- Business Planner (Sep 2021 – Sep 2022): global demand forecasting, Tableau dashboards, supply chain planning across EMEA/APJ/North America, Lean practices
- Process Specialist – Operational Excellence (Feb 2019 – Sep 2021): Jira implementation for 100+ users, API-based legacy data migration, process improvement, budget monitoring

Key skills: Project Management (R&D and Operations), Agile and Lean delivery, stakeholder engagement, resource and capacity planning, Tableau, Jira, Advanced Excel, cross-functional team coordination, budget monitoring

Education: BSc Industrial Engineering and Management (Digital Systems Management), Shenkar College of Engineering and Design, 2016–2021

Job preferences:
- Target roles: Project Manager or Business Operations Manager
- Prefers tech companies or Israeli startups in London; open to any industry
- STRONGLY dislikes forecasting-heavy, financial-modelling, or pure planning desk roles — score these 1–3
- Wants proactive, creative, hands-on delivery work
- Available for hybrid or on-site anywhere in London and Greater London

JOB:
Role: ${role}
Company: ${company}
${jobContext}

Return ONLY a valid JSON object — no markdown, no code fences, no extra text:
{"score":<integer 1-10>,"verdict":"<one concise sentence>","fit":"<one sentence>","gaps":"<one sentence>","position":"<one sentence>"}

Score guide: 9–10 = near-perfect PM/ops delivery role, tech or startup context, London; 7–8 = strong match with minor gaps; 5–6 = partial match, worth considering; 3–4 = weak fit; 1–2 = poor fit or a forecasting/financial-modelling-heavy role she dislikes.`;

    const reqBody = JSON.stringify({
      model: 'claude-sonnet-4-5',
      max_tokens: 512,
      messages: [{ role: 'user', content: prompt }],
    });

    try {
      const result = await httpsPost({
        hostname: 'api.anthropic.com',
        path: '/v1/messages',
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'x-api-key': API_KEY,
          'anthropic-version': '2023-06-01',
          'Content-Length': Buffer.byteLength(reqBody),
        },
      }, reqBody);

      const data = JSON.parse(result.body);
      if (result.status !== 200) {
        send(res, result.status, { error: data.error?.message || `Anthropic returned ${result.status}` });
        return;
      }

      const text  = (data.content?.[0]?.text || '').trim();
      const match = text.match(/\{[\s\S]*\}/);
      if (!match) { send(res, 500, { error: 'No JSON in score response' }); return; }
      send(res, 200, JSON.parse(match[0]));
    } catch (err) {
      send(res, 502, { error: 'Could not reach Anthropic: ' + (err.message || err) });
    }
    return;
  }

  // ── POST /api/tailor-cv ──────────────────────────────────────────────────
  if (req.method === 'POST' && pathname === '/api/tailor-cv') {
    if (!API_KEY) {
      send(res, 500, { error: 'ANTHROPIC_API_KEY is not set' });
      return;
    }

    let raw;
    try { raw = await readBody(req); }
    catch { send(res, 400, { error: 'Failed to read request body' }); return; }

    let payload;
    try { payload = JSON.parse(raw); }
    catch { send(res, 400, { error: 'Request body is not valid JSON' }); return; }

    const { role, company, notes, jobDescription } = payload;
    if (!role || !company) {
      send(res, 400, { error: 'Request must include role and company' });
      return;
    }

    let cvBase;
    try {
      cvBase = fs.readFileSync(path.join(__dirname, 'cv-base.txt'), 'utf8');
    } catch {
      send(res, 500, { error: 'cv-base.txt not found in ~/job-tracker/' });
      return;
    }

    const jobContext = jobDescription
      ? `Job Description:\n${jobDescription}${notes ? `\n\nAdditional context: ${notes}` : ''}`
      : `Additional context: ${notes || 'None provided'}`;

    const prompt = `You are an expert CV writer. Tailor the CV below for the specific job application, preserving the EXACT format and structure of the original.

JOB DETAILS:
Role: ${role}
Company: ${company}
${jobContext}

BASE CV:
${cvBase}

STRICT FORMAT RULES — follow exactly:
1. Line 1: Name in ALL CAPS (e.g. EFRAT GNAPP)
2. Line 2: Role title (e.g. Project Manager | Business Operations Manager) — adjust to match the target role
3. Line 3: Contact info exactly as given (do not change)
4. Blank line, then profile paragraph (no heading) — rewrite to speak directly to this role
5. Section headings in ALL CAPS on their own line: KEY ACHIEVEMENTS, KEY EXPERIENCE & SKILLS, WORK EXPERIENCE, EDUCATION, OTHER INFORMATION
6. KEY ACHIEVEMENTS: 3 bullet points starting with "- ", written in ALL CAPS for the key verb phrase, then normal case for the rest
7. KEY EXPERIENCE & SKILLS: lines in format "CATEGORY IN CAPS: description sentence." — reorder to lead with the most relevant skills
8. WORK EXPERIENCE: company lines as "COMPANY NAME | Date – Date", role lines starting with "// ROLE TITLE | Date – Date", then a bold key-achievement sentence (no bullet), then bullet points starting with "- "
9. Do not add or remove sections; do not add markdown, bold markers, or code fences
10. Write in clear, professional British English
11. Output ONLY the tailored CV text — nothing else`;

    const reqBody = JSON.stringify({
      model: 'claude-sonnet-4-5',
      max_tokens: 2048,
      messages: [{ role: 'user', content: prompt }],
    });

    try {
      const result = await httpsPost({
        hostname: 'api.anthropic.com',
        path: '/v1/messages',
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'x-api-key': API_KEY,
          'anthropic-version': '2023-06-01',
          'Content-Length': Buffer.byteLength(reqBody),
        },
      }, reqBody);

      const data = JSON.parse(result.body);
      if (result.status !== 200) {
        send(res, result.status, { error: data.error?.message || `Anthropic returned ${result.status}` });
        return;
      }

      send(res, 200, { cv: (data.content?.[0]?.text || '').trim() });
    } catch (err) {
      send(res, 502, { error: 'Could not reach Anthropic: ' + (err.message || err) });
    }
    return;
  }

  // ── POST /api/tailor-cv-docx ─────────────────────────────────────────────
  if (req.method === 'POST' && pathname === '/api/tailor-cv-docx') {
    let raw;
    try { raw = await readBody(req); } catch { send(res, 400, { error: 'Failed to read body' }); return; }
    let payload;
    try { payload = JSON.parse(raw); } catch { send(res, 400, { error: 'Invalid JSON' }); return; }

    const { cvText, role = 'CV', company = '' } = payload;
    if (!cvText) { send(res, 400, { error: 'cvText is required' }); return; }

    try {
      const doc    = buildDocx(cvText);
      const buffer = await Packer.toBuffer(doc);
      const safe   = (role + (company ? ' ' + company : '')).replace(/[^a-zA-Z0-9 ]/g, '').trim().replace(/\s+/g, '_').slice(0, 60);
      const fname  = `CV_Efrat_Gnapp_${safe}.docx`;
      res.writeHead(200, {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'Content-Disposition': `attachment; filename="${fname}"`,
        'Content-Length': buffer.length,
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Headers': 'Content-Type',
      });
      res.end(buffer);
    } catch (err) {
      send(res, 500, { error: 'Failed to build Word document: ' + (err.message || err) });
    }
    return;
  }

  // ── POST /api/cover-letter ───────────────────────────────────────────────
  if (req.method === 'POST' && pathname === '/api/cover-letter') {
    if (!API_KEY) { send(res, 500, { error: 'ANTHROPIC_API_KEY is not set' }); return; }

    let raw;
    try { raw = await readBody(req); } catch { send(res, 400, { error: 'Failed to read request body' }); return; }
    let payload;
    try { payload = JSON.parse(raw); } catch { send(res, 400, { error: 'Request body is not valid JSON' }); return; }

    const { role, company, notes, jobDescription } = payload;
    if (!role || !company) { send(res, 400, { error: 'Request must include role and company' }); return; }

    let cvBase;
    try { cvBase = fs.readFileSync(path.join(__dirname, 'cv-base.txt'), 'utf8'); }
    catch { send(res, 500, { error: 'cv-base.txt not found in ~/job-tracker/' }); return; }

    const todayStr = new Date().toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' });

    const jobContext = jobDescription
      ? `Job Description:\n${jobDescription}${notes ? `\n\nAdditional context: ${notes}` : ''}`
      : `Additional context: ${notes || 'None provided'}`;

    const prompt = `You are writing a cover letter on behalf of Efrat Gnapp for a job application. Use the CV below to ground every claim in real experience. Be specific, warm, and direct — not generic or sycophantic.

JOB DETAILS:
Role: ${role}
Company: ${company}
${jobContext}

CV FOR REFERENCE:
${cvBase}

OUTPUT EXACTLY the following format — nothing before, nothing after:
EFRAT GNAPP
NW8 7BS, London · (+44) 7846 676635 · efrat.gnapp@gmail.com · LinkedIn
${todayStr}

Dear Hiring Manager,

[Opening paragraph, 2-3 sentences: express genuine specific interest in this role at this company. Reference something concrete about what draws her to this opportunity — connect the company's work or context to her background.]

[Second paragraph, 3-4 sentences: pick 1-2 achievements from HP Indigo that directly map to the key requirements of this role. Be concrete — name the actual projects, numbers, and outcomes from the CV.]

[Third paragraph, 3-4 sentences: highlight a second angle — cross-functional leadership, stakeholder management, or a specific technical/operational skill this role needs. Draw directly from the CV.]

[Closing paragraph, 2-3 sentences: confident, warm close. Express enthusiasm and openness to discuss further. No empty phrases.]

Yours sincerely,
Efrat Gnapp

RULES:
- Write in first person as Efrat
- Tone: professional and direct — not stiff, not fawning
- British English spelling
- No markdown, no asterisks, no bold markers, no section headers, no bullet points
- Output ONLY the cover letter text in the exact format above`;

    const reqBody = JSON.stringify({
      model: 'claude-sonnet-4-5',
      max_tokens: 1200,
      messages: [{ role: 'user', content: prompt }],
    });

    try {
      const result = await httpsPost({
        hostname: 'api.anthropic.com',
        path: '/v1/messages',
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'x-api-key': API_KEY,
          'anthropic-version': '2023-06-01',
          'Content-Length': Buffer.byteLength(reqBody),
        },
      }, reqBody);

      const data = JSON.parse(result.body);
      if (result.status !== 200) {
        send(res, result.status, { error: data.error?.message || `Anthropic returned ${result.status}` });
        return;
      }
      send(res, 200, { letter: (data.content?.[0]?.text || '').trim() });
    } catch (err) {
      send(res, 502, { error: 'Could not reach Anthropic: ' + (err.message || err) });
    }
    return;
  }

  // ── POST /api/cover-letter-docx ──────────────────────────────────────────
  if (req.method === 'POST' && pathname === '/api/cover-letter-docx') {
    let raw;
    try { raw = await readBody(req); } catch { send(res, 400, { error: 'Failed to read body' }); return; }
    let payload;
    try { payload = JSON.parse(raw); } catch { send(res, 400, { error: 'Invalid JSON' }); return; }

    const { letterText, role = 'Role', company = '' } = payload;
    if (!letterText) { send(res, 400, { error: 'letterText is required' }); return; }

    try {
      const doc    = buildCoverLetterDocx(letterText);
      const buffer = await Packer.toBuffer(doc);
      const safe   = (role + (company ? ' ' + company : '')).replace(/[^a-zA-Z0-9 ]/g, '').trim().replace(/\s+/g, '_').slice(0, 60);
      const fname  = `Cover_Letter_Efrat_Gnapp_${safe}.docx`;
      res.writeHead(200, {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'Content-Disposition': `attachment; filename="${fname}"`,
        'Content-Length': buffer.length,
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Headers': 'Content-Type',
      });
      res.end(buffer);
    } catch (err) {
      send(res, 500, { error: 'Failed to build Word document: ' + (err.message || err) });
    }
    return;
  }

  // ── POST /api/extract ────────────────────────────────────────────────────
  if (req.method === 'POST' && pathname === '/api/extract') {
    if (!API_KEY) {
      send(res, 500, { error: 'ANTHROPIC_API_KEY is not set in the environment.\nRun: export ANTHROPIC_API_KEY=sk-ant-... && node server.js' });
      return;
    }

    let raw;
    try { raw = await readBody(req); }
    catch { send(res, 400, { error: 'Failed to read request body' }); return; }

    let payload;
    try { payload = JSON.parse(raw); }
    catch { send(res, 400, { error: 'Request body is not valid JSON' }); return; }

    const { imageData, mediaType } = payload;
    if (!imageData || !mediaType) {
      send(res, 400, { error: 'Request must include imageData (base64) and mediaType' });
      return;
    }

    const reqBody = JSON.stringify({
      model: 'claude-sonnet-4-5',
      max_tokens: 2048,
      messages: [{
        role: 'user',
        content: [
          { type: 'image', source: { type: 'base64', media_type: mediaType, data: imageData } },
          {
            type: 'text',
            text: 'Look at this screenshot carefully. It may contain one or more job listings. Extract ALL job listings visible and return ONLY a valid JSON array — no markdown, no extra text, no code fences:\n[{"company":"company name","role":"job title","notes":"location and brief description if visible"}]\nIf there is only one job, still return a JSON array with one element. Use an empty string for any field you cannot find.'
          }
        ]
      }]
    });

    try {
      const result = await httpsPost({
        hostname: 'api.anthropic.com',
        path: '/v1/messages',
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'x-api-key': API_KEY,
          'anthropic-version': '2023-06-01',
          'Content-Length': Buffer.byteLength(reqBody),
        }
      }, reqBody);

      const data = JSON.parse(result.body);

      if (result.status !== 200) {
        send(res, result.status, { error: data.error?.message || `Anthropic returned ${result.status}` });
        return;
      }

      const text = (data.content?.[0]?.text || '').trim();
      const arrayMatch = text.match(/\[[\s\S]*\]/);
      const objectMatch = text.match(/\{[\s\S]*\}/);
      if (!arrayMatch && !objectMatch) {
        send(res, 500, { error: 'Claude responded but no JSON was found in the output' });
        return;
      }

      let extracted;
      if (arrayMatch) {
        extracted = JSON.parse(arrayMatch[0]);
        if (!Array.isArray(extracted)) extracted = [extracted];
      } else {
        extracted = [JSON.parse(objectMatch[0])];
      }

      send(res, 200, { jobs: extracted });
    } catch (err) {
      send(res, 502, { error: 'Could not reach Anthropic: ' + (err.message || err) });
    }
    return;
  }

  // ── POST /api/draft-outreach ─────────────────────────────────────────────
  if (req.method === 'POST' && pathname === '/api/draft-outreach') {
    if (!API_KEY) { send(res, 500, { error: 'ANTHROPIC_API_KEY is not set' }); return; }

    let raw;
    try { raw = await readBody(req); } catch { send(res, 400, { error: 'Failed to read request body' }); return; }
    let payload;
    try { payload = JSON.parse(raw); } catch { send(res, 400, { error: 'Request body is not valid JSON' }); return; }

    const { role, company, contactName, notes, jobDescription } = payload;
    if (!role || !company) { send(res, 400, { error: 'Request must include role and company' }); return; }

    const jdContext = jobDescription
      ? `\n\nJob description context: ${jobDescription.slice(0, 500)}`
      : (notes ? `\n\nRole context: ${notes}` : '');

    const hasContact = contactName && contactName.trim();
    const addressee = hasContact ? contactName.trim() : null;

    const introLine = hasContact
      ? `Write two outreach messages from Efrat Gnapp to ${addressee}, who she believes works at or has a connection to ${company}. Efrat is applying for the ${role} role there.`
      : `Write two outreach messages from Efrat Gnapp for the ${role} role at ${company}. She does not have a named contact, so both messages should be addressed warmly but generically — the LinkedIn version to "the team at ${company}" or a recruiter she may reach via LinkedIn, and the email version as a cold outreach to the Hiring Manager.`;

    const linkedinOpening = hasContact
      ? `- Address ${addressee} directly and warmly by first name if identifiable, otherwise use a natural opening`
      : `- Open warmly without a named contact — e.g. "Hi [Name]," left as a placeholder, or addressed to "the ${company} team" — keep it natural, not stiff`;

    const emailOpening = hasContact
      ? `- Opening sentence: mention ${addressee} and state the role (${role} at ${company}) and that she is applying`
      : `- Opening sentence: address the Hiring Manager warmly (e.g. "Dear Hiring Manager,") and state the role (${role} at ${company}) and that she is applying`;

    const prompt = `${introLine}${jdContext}

About Efrat: She is a Project Manager / Business Operations Manager with 5 years at HP Indigo leading technical implementations, cross-functional delivery, and operational excellence programmes. She is now job searching in London.

Write TWO versions:

VERSION 1 — LinkedIn message:
- 3–4 sentences MAX, ideally under 300 characters
- Very warm and direct — suitable for a LinkedIn connection request or message
${linkedinOpening}
- One sentence on who she is and what she is looking for, one sentence asking if they would be happy to refer her or have a quick chat
- Conversational, friendly, no jargon

VERSION 2 — Email:
- 150–200 words MAX — concise and direct, not a long formal letter
${emailOpening}
- One short paragraph (2–3 sentences): one specific, concrete achievement from HP Indigo that is directly relevant to this role — no generic claims
- Closing sentence: a single direct ask — a quick call or internal referral
- Professional British English, no filler phrases, no "I hope this finds you well", no sign-off pleasantries beyond "Best, Efrat"

Both versions must:
- Sound like a real person, not a template
- NOT start with "I hope this message finds you well" or similar filler
- Reference the specific company (${company}) and role (${role})

Return ONLY a valid JSON object — no markdown, no code fences:
{"linkedin":"<VERSION 1 text>","email":"<VERSION 2 text>"}`;

    const reqBody = JSON.stringify({
      model: 'claude-sonnet-4-5',
      max_tokens: 1200,
      messages: [{ role: 'user', content: prompt }],
    });

    try {
      const result = await httpsPost({
        hostname: 'api.anthropic.com',
        path: '/v1/messages',
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'x-api-key': API_KEY,
          'anthropic-version': '2023-06-01',
          'Content-Length': Buffer.byteLength(reqBody),
        },
      }, reqBody);

      const data = JSON.parse(result.body);
      if (result.status !== 200) {
        send(res, result.status, { error: data.error?.message || `Anthropic returned ${result.status}` });
        return;
      }
      const text = (data.content?.[0]?.text || '').trim();
      const match = text.match(/\{[\s\S]*\}/);
      if (!match) { send(res, 500, { error: 'No JSON in outreach response' }); return; }
      const parsed = JSON.parse(match[0]);
      send(res, 200, { linkedin: parsed.linkedin || '', email: parsed.email || '' });
    } catch (err) {
      send(res, 502, { error: 'Could not reach Anthropic: ' + (err.message || err) });
    }
    return;
  }

  // ── POST /api/salary ─────────────────────────────────────────────────────
  if (req.method === 'POST' && pathname === '/api/salary') {
    if (!API_KEY) { send(res, 500, { error: 'ANTHROPIC_API_KEY is not set' }); return; }

    let raw;
    try { raw = await readBody(req); } catch { send(res, 400, { error: 'Failed to read request body' }); return; }
    let payload;
    try { payload = JSON.parse(raw); } catch { send(res, 400, { error: 'Request body is not valid JSON' }); return; }

    const { role, company, jobDescription } = payload;
    if (!role || !company) { send(res, 400, { error: 'Request must include role and company' }); return; }

    const jdSnippet = jobDescription ? `\n\nJob description snippet: ${jobDescription.slice(0, 600)}` : '';
    const prompt = `Search for current 2025 salary benchmarks for a "${role}" role in London, UK. The company is ${company}.${jdSnippet}

Return ONLY a valid JSON object — no markdown, no code fences, no explanation before or after:
{"low":<annual GBP integer>,"mid":<annual GBP integer>,"high":<annual GBP integer>,"seniority":"<seniority level this range applies to>","context":"<1-2 sentences on key factors affecting pay for this role in London>","sources":["<source or data point 1>","<source or data point 2>","<source or data point 3>"]}`;

    const reqBody = JSON.stringify({
      model: 'claude-sonnet-4-6',
      max_tokens: 1024,
      tools: [{ type: 'web_search_20250305', name: 'web_search', max_uses: 5 }],
      messages: [{ role: 'user', content: prompt }],
    });

    try {
      const result = await httpsPost({
        hostname: 'api.anthropic.com',
        path: '/v1/messages',
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'x-api-key': API_KEY,
          'anthropic-version': '2023-06-01',
          'anthropic-beta': 'web-search-2025-03-05',
          'Content-Length': Buffer.byteLength(reqBody),
        },
      }, reqBody);

      const data = JSON.parse(result.body);
      if (result.status !== 200) {
        send(res, result.status, { error: data.error?.message || `Anthropic returned ${result.status}` });
        return;
      }

      const text = (data.content || []).filter(b => b.type === 'text').map(b => b.text).join('\n').trim();
      const match = text.match(/\{[\s\S]*\}/);
      if (!match) { send(res, 500, { error: 'No JSON in salary response' }); return; }
      send(res, 200, JSON.parse(match[0]));
    } catch (err) {
      send(res, 502, { error: 'Could not reach Anthropic: ' + (err.message || err) });
    }
    return;
  }

  // ── POST /api/company-research ────────────────────────────────────────────
  if (req.method === 'POST' && pathname === '/api/company-research') {
    if (!API_KEY) { send(res, 500, { error: 'ANTHROPIC_API_KEY is not set' }); return; }

    let raw;
    try { raw = await readBody(req); } catch { send(res, 400, { error: 'Failed to read request body' }); return; }
    let payload;
    try { payload = JSON.parse(raw); } catch { send(res, 400, { error: 'Request body is not valid JSON' }); return; }

    const { role, company } = payload;
    if (!role || !company) { send(res, 400, { error: 'Request must include role and company' }); return; }

    const prompt = `Research the company "${company}" in the context of a "${role}" job application in London.

Return ONLY a valid JSON object — no markdown, no code fences, no explanation before or after:
{"what_they_do":"<2-3 sentences describing the company, its product or service, and its market>","size_stage":"<company size, funding stage or public/private status, headcount if known>","london_presence":"<London office location, team size, whether London is HQ or a regional office>","recent_news":"<1-2 notable developments from the past 6 months: funding, product launches, leadership changes, press coverage>","culture":"<2-3 sentences on work culture, values, engineering or operations style based on public information>","fit_for_efrat":"<1-2 sentences on why this company and role could be a strong match for a PM/Business Operations Manager with 5 years at HP Indigo, skilled in delivery, cross-functional teams, Jira, and operational excellence>"}`;

    const reqBody = JSON.stringify({
      model: 'claude-sonnet-4-6',
      max_tokens: 1500,
      tools: [{ type: 'web_search_20250305', name: 'web_search', max_uses: 5 }],
      messages: [{ role: 'user', content: prompt }],
    });

    try {
      const result = await httpsPost({
        hostname: 'api.anthropic.com',
        path: '/v1/messages',
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'x-api-key': API_KEY,
          'anthropic-version': '2023-06-01',
          'anthropic-beta': 'web-search-2025-03-05',
          'Content-Length': Buffer.byteLength(reqBody),
        },
      }, reqBody);

      const data = JSON.parse(result.body);
      if (result.status !== 200) {
        send(res, result.status, { error: data.error?.message || `Anthropic returned ${result.status}` });
        return;
      }

      const text = (data.content || []).filter(b => b.type === 'text').map(b => b.text).join('\n').trim();
      const match = text.match(/\{[\s\S]*\}/);
      if (!match) { send(res, 500, { error: 'No JSON in research response' }); return; }
      send(res, 200, JSON.parse(match[0]));
    } catch (err) {
      send(res, 502, { error: 'Could not reach Anthropic: ' + (err.message || err) });
    }
    return;
  }

  // ── POST /api/linkedin-connections ───────────────────────────────────────
  if (req.method === 'POST' && pathname === '/api/linkedin-connections') {
    if (!API_KEY) { send(res, 500, { error: 'ANTHROPIC_API_KEY is not set' }); return; }

    let raw;
    try { raw = await readBody(req); } catch { send(res, 400, { error: 'Failed to read request body' }); return; }
    let payload;
    try { payload = JSON.parse(raw); } catch { send(res, 400, { error: 'Request body is not valid JSON' }); return; }

    const { company } = payload;
    if (!company) { send(res, 400, { error: 'Request must include company' }); return; }

    const prompt = `Search LinkedIn for people who currently work at "${company}" and are likely to be relevant contacts for a Project Manager or Business Operations professional job seeker in London.

Focus on people in these roles: Project Manager, Product Manager, Programme Manager, Business Operations, Operations Manager, Chief of Staff, Head of Operations, Delivery Manager, Scrum Master, Engineering Manager, or similar tech/ops/PM roles. Prioritise people based in London or the UK.

Search for real, publicly visible LinkedIn profiles. For each person found, provide their name, current job title, and their LinkedIn profile URL.

Return ONLY a valid JSON object — no markdown, no code fences, no explanation:
{"connections":[{"name":"<full name>","title":"<current job title at ${company}>","url":"<linkedin.com/in/... profile URL>"},{"name":"...","title":"...","url":"..."}]}

Rules:
- Return up to 5 people
- Only include people who currently work at ${company} based on their LinkedIn profile
- Only include profiles with a real linkedin.com/in/ URL that you have found via search
- If you cannot find any relevant profiles, return {"connections":[]}
- Do not fabricate names, titles, or URLs`;

    const reqBody = JSON.stringify({
      model: 'claude-sonnet-4-6',
      max_tokens: 1000,
      tools: [{ type: 'web_search_20250305', name: 'web_search', max_uses: 5 }],
      messages: [{ role: 'user', content: prompt }],
    });

    try {
      const result = await httpsPost({
        hostname: 'api.anthropic.com',
        path: '/v1/messages',
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'x-api-key': API_KEY,
          'anthropic-version': '2023-06-01',
          'anthropic-beta': 'web-search-2025-03-05',
          'Content-Length': Buffer.byteLength(reqBody),
        },
      }, reqBody);

      const data = JSON.parse(result.body);
      if (result.status !== 200) {
        send(res, result.status, { error: data.error?.message || `Anthropic returned ${result.status}` });
        return;
      }

      const text = (data.content || []).filter(b => b.type === 'text').map(b => b.text).join('\n').trim();
      const match = text.match(/\{[\s\S]*\}/);
      if (!match) { send(res, 500, { error: 'No JSON in connections response' }); return; }
      const parsed = JSON.parse(match[0]);
      send(res, 200, { connections: Array.isArray(parsed.connections) ? parsed.connections : [] });
    } catch (err) {
      send(res, 502, { error: 'Could not reach Anthropic: ' + (err.message || err) });
    }
    return;
  }

  // ── POST /api/rejection-insight ──────────────────────────────────────────
  if (req.method === 'POST' && pathname === '/api/rejection-insight') {
    if (!API_KEY) { send(res, 500, { error: 'ANTHROPIC_API_KEY is not set' }); return; }

    let raw;
    try { raw = await readBody(req); } catch { send(res, 400, { error: 'Failed to read request body' }); return; }
    let payload;
    try { payload = JSON.parse(raw); } catch { send(res, 400, { error: 'Request body is not valid JSON' }); return; }

    const { rejCounts = {}, total = 0 } = payload;
    const REJ_LABELS = {
      no_response:  'no response from the employer',
      cv_screen:    'rejected after CV screening',
      phone_screen: 'rejected after phone screen',
      interview:    'rejected after interview',
      withdrew:     'withdrew themselves',
      other:        'other',
    };
    const summary = Object.entries(rejCounts)
      .filter(([k]) => k !== '_no_reason')
      .map(([k, v]) => `${v}× ${REJ_LABELS[k] || k}`)
      .join(', ');
    if (!summary) { send(res, 200, { insight: '' }); return; }

    const prompt = `A job seeker has ${total} rejections in their pipeline, broken down as follows: ${summary}.

Write a concise 2–3 sentence insight for this person in British English. Be direct and constructive: identify the most significant pattern and give one specific, actionable recommendation they can act on this week. Do not use bullet points. Do not start with "I" or "Your". Avoid generic career advice platitudes.`;

    const reqBody = JSON.stringify({
      model: 'claude-sonnet-4-5',
      max_tokens: 200,
      messages: [{ role: 'user', content: prompt }],
    });

    try {
      const result = await httpsPost({
        hostname: 'api.anthropic.com',
        path: '/v1/messages',
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'x-api-key': API_KEY,
          'anthropic-version': '2023-06-01',
          'Content-Length': Buffer.byteLength(reqBody),
        },
      }, reqBody);

      const data = JSON.parse(result.body);
      if (result.status !== 200) {
        send(res, result.status, { error: data.error?.message || `Anthropic returned ${result.status}` });
        return;
      }
      const insight = data.content?.[0]?.text?.trim() || '';
      send(res, 200, { insight });
    } catch (err) {
      send(res, 502, { error: 'Could not reach Anthropic: ' + (err.message || err) });
    }
    return;
  }

  // ── POST /api/key-skills ──────────────────────────────────────────────────
  if (req.method === 'POST' && pathname === '/api/key-skills') {
    if (!API_KEY) { send(res, 500, { error: 'ANTHROPIC_API_KEY is not set' }); return; }

    let raw;
    try { raw = await readBody(req); } catch { send(res, 400, { error: 'Failed to read request body' }); return; }
    let payload;
    try { payload = JSON.parse(raw); } catch { send(res, 400, { error: 'Request body is not valid JSON' }); return; }

    const { role, company, jobDescription, notes } = payload;
    if (!role || !company) { send(res, 400, { error: 'Request must include role and company' }); return; }

    const jdContext = jobDescription
      ? `\nJob description:\n${jobDescription.slice(0, 3000)}`
      : notes ? `\nJob notes: ${notes.slice(0, 500)}` : '';

    const prompt = `You are a UK-based career coach helping a candidate prepare for a job application in the British job market. Analyse the job and the candidate's background, then return the most important skills she should highlight.

UK MARKET CONTEXT — apply throughout your response:
- Use British English spelling and terminology throughout: "programme" not "program", "organisation" not "organisation", "behaviour" not "behavior", "analyse" not "analyze", "prioritise" not "prioritize", "CV" not "resume"
- Frame skills using language common in UK job postings and interviews: "stakeholder management", "delivery", "commercial awareness", "line management", "business partnering", "continuous improvement", "cross-functional collaboration"
- Reference Agile terminology as used in the UK tech sector: Scrum, Kanban, sprint ceremonies, backlog refinement, ways of working
- Use UK corporate communication norms: understated, evidence-based, outcome-focused — avoid American-style hyperbole
- Where relevant, note how skills map to UK interview competency frameworks (e.g. STAR method, values-based interviews common in UK public sector and tech firms)

CANDIDATE PROFILE:
Name: Efrat Gnapp
Title: Project Manager / Business Operations Manager
Location: London, UK (NW8 — UK citizen, full right to work)

Profile: Project Manager with 5+ years managing complex technical implementations and onboarding programmes in a global technology organisation. Proven track record delivering SaaS-style implementations end-to-end, managing cross-functional teams, and ensuring successful adoption by 100+ users. Skilled at translating client requirements into executable technical plans whilst maintaining strong stakeholder relationships at all levels.

Key achievements:
- Led cross-functional delivery of three product programmes simultaneously (total budget ~$40M), coordinating R&D, QA, Marketing, and Operations; all milestones delivered on time and within budget including Drupa 2024 launches
- Improved programme planning and delivery efficiency by introducing Lean practices and a global forecasting model across EMEA, APJ, and North America
- Delivered organisation-wide Jira implementation across Operations — end-to-end planning, 5-year legacy data migration, adoption by 100+ engineers

Work experience — HP Indigo, Feb 2019 – Aug 2024:
- Project Manager R&D (Sep 2022 – Aug 2024): led three concurrent product programmes (~$40M annual budget), cross-functional delivery, Lean methodology, executive reporting, risk management, technical integrations
- Business Planner (Sep 2021 – Sep 2022): global demand forecasting, Tableau dashboards, supply chain planning across EMEA/APJ/North America, Lean practices
- Process Specialist – Operational Excellence (Feb 2019 – Sep 2021): Jira implementation for 100+ users, API-based legacy data migration, process improvement, budget monitoring

Key skills: Project Management (R&D and Operations), Agile and Lean delivery, stakeholder engagement, resource and capacity planning, Tableau, Jira, Advanced Excel, cross-functional team coordination, budget monitoring

JOB:
Role: ${role}
Company: ${company}${jdContext}

Return ONLY a valid JSON object — no markdown, no code fences, no explanation:
{"lead":[{"skill":"<skill name>","why":"<one sentence: why it's relevant and how to frame it for this role, using UK market language>"},{"skill":"...","why":"..."},{"skill":"...","why":"..."},{"skill":"...","why":"..."},{"skill":"...","why":"..."}],"also":["<skill>","<skill>","<skill>","<skill>"]}

Rules:
- "lead": exactly 5 skills from Efrat's background that are most relevant to THIS specific role, each with a concise one-sentence framing tip written in UK English and UK market tone
- "also": 3–5 additional supporting skills worth mentioning, named using UK terminology
- Only include skills Efrat actually has; do not invent or pad
- Prioritise skills the job description explicitly calls for
- All text must use British English spelling and UK professional register`;

    const reqBody = JSON.stringify({
      model: 'claude-sonnet-4-5',
      max_tokens: 1000,
      messages: [{ role: 'user', content: prompt }],
    });

    try {
      const result = await httpsPost({
        hostname: 'api.anthropic.com',
        path: '/v1/messages',
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'x-api-key': API_KEY,
          'anthropic-version': '2023-06-01',
          'Content-Length': Buffer.byteLength(reqBody),
        },
      }, reqBody);

      const data = JSON.parse(result.body);
      if (result.status !== 200) {
        send(res, result.status, { error: data.error?.message || `Anthropic returned ${result.status}` });
        return;
      }

      const text = (data.content?.[0]?.text || '').trim();
      const match = text.match(/\{[\s\S]*\}/);
      if (!match) { send(res, 500, { error: 'No JSON in key-skills response' }); return; }
      send(res, 200, JSON.parse(match[0]));
    } catch (err) {
      send(res, 502, { error: 'Could not reach Anthropic: ' + (err.message || err) });
    }
    return;
  }

  if (req.method === 'POST' && pathname === '/api/weekly-recommendation') {
    let body = '';
    req.on('data', c => body += c);
    req.on('end', async () => {
      try {
        const { thisWeekCount, pipeline, followUpCount, topScored } = JSON.parse(body);
        const pip = pipeline || {};
        const pipSummary = Object.entries(pip).map(([k,v]) => `${k}: ${v}`).join(', ');
        const topScoredText = (topScored||[]).length
          ? (topScored).map(j => `${j.company} – ${j.role} (${j.score}/10, status: ${j.status})`).join('; ')
          : 'None scored yet';
        const prompt = `You are a career coach reviewing a job seeker's weekly job search pipeline. Give a focused 2–3 sentence recommendation on what she should prioritise this week.

Pipeline state:
- Applications sent this week: ${thisWeekCount}
- Overall pipeline: ${pipSummary || 'empty'}
- Roles needing follow-up: ${followUpCount}
- Top scored roles not yet applied: ${topScoredText}

Be direct and specific. Focus on the most impactful action for this week. Do not use bullet points — write in plain flowing sentences.`;

        const apiRes = await fetch('https://api.anthropic.com/v1/messages', {
          method: 'POST',
          headers: { 'x-api-key': process.env.ANTHROPIC_API_KEY, 'anthropic-version': '2023-06-01', 'content-type': 'application/json' },
          body: JSON.stringify({
            model: 'claude-sonnet-4-5',
            max_tokens: 200,
            messages: [{ role: 'user', content: prompt }],
          }),
        });
        const apiData = await apiRes.json();
        if (!apiRes.ok) { send(res, 502, { error: apiData.error?.message || 'Anthropic error' }); return; }
        const recommendation = apiData.content?.[0]?.text?.trim() || '';
        send(res, 200, { recommendation });
      } catch (err) {
        send(res, 500, { error: err.message });
      }
    });
    return;
  }

  res.writeHead(404); res.end('Not found');
});

server.listen(PORT, '127.0.0.1', () => {
  console.log('');
  console.log('  ┌─ Job Tracker proxy server ──────────────────────┐');
  console.log(`  │  http://localhost:${PORT}/job-tracker.html         │`);
  console.log('  └─────────────────────────────────────────────────┘');
  if (!API_KEY) {
    console.log('');
    console.log('  ⚠  ANTHROPIC_API_KEY is not set.');
    console.log('     Screenshot extraction will fail until you set it:');
    console.log('     export ANTHROPIC_API_KEY=sk-ant-...');
    console.log('     Then restart this server.');
  } else {
    console.log('  ✓  ANTHROPIC_API_KEY detected — ready to extract screenshots');
  }
  console.log('');
});
