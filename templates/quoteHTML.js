'use strict';
const fs   = require('fs');
const path = require('path');

// ── Brand colors (same as React UI avatarColors) ─────────────────────────────
const INSURER_COLORS = ['#1565C0','#2E7D32','#283593','#B71C1C','#E65100','#880E4F'];
const BRAND_GREEN    = '#4CAF50';
const BRAND_DARK     = '#388E3C';

// ── Hardcoded feature data (fallback when features_override not provided) ────
const MASTER = {
  'ICICI Lombard - Elevate': {
    'Room Rent':'Single Private A/C Room','ICU Limit':'No limit on ICU',
    'Pre-Hospitalization':'90 days','Post-Hospitalization':'180 days',
    'Day Care Procedures':'All procedures up to S.I.','Copay':'0%',
    'Restoration Benefit':'Unlimited (incl. same illness)',
    'NCB Benefit':'Unlimited 100%/yr — No Cap','Non-Consumables':'Covered',
    'Hospitalization @ Home':'Up to Sum Assured','Ambulance Cover':'Up to Sum Assured',
    'Air Ambulance':'Up to Sum Assured','AYUSH Treatment':'Up to Sum Assured',
    'Organ Donor':'Up to Sum Assured','Modern Treatments':'Up to Sum Assured',
    '2-Hour Hospitalization':'Covered','E-Consultation':'Unlimited',
    'Preventive Health Check-up':'Covered (Optional)',
    'Maternity Cover':'Optional Rider (10% SA, Max Rs.1L)','Newborn Cover':'Day 1 (10% SA)',
    'Worldwide Cover':'Not Available','OPD Cover':'Optional Rider',
    'Lock-the-Clock':'Not Applicable','Priority Claim Desk':'Not Available',
    'Cash+ Wallet':'—','Unique Features':'Unlimited NCB | 2-hr hospitalization | Newborn Day-1',
  },
  'Niva Bupa - ReAssure 3.0': {
    'Room Rent':'Any Room incl. Suite','ICU Limit':'No limit on ICU',
    'Pre-Hospitalization':'60 days','Post-Hospitalization':'180 days',
    'Day Care Procedures':'All procedures up to S.I.','Copay':'0%',
    'Restoration Benefit':'Unlimited — Same illness, never ends',
    'NCB Benefit':'N/A — Unlimited Cover Plan','Non-Consumables':'Covered',
    'Hospitalization @ Home':'Up to Sum Insured','Ambulance Cover':'Up to Sum Insured',
    'Air Ambulance':'Up to Rs. 5 Lakh','AYUSH Treatment':'Up to Sum Assured',
    'Organ Donor':'Up to Sum Assured','Modern Treatments':'Up to Sum Assured',
    '2-Hour Hospitalization':'Covered','E-Consultation':'Unlimited',
    'Preventive Health Check-up':'Annual check-up covered',
    'Maternity Cover':'Not Available','Newborn Cover':'Not Available',
    'Worldwide Cover':'Available (with rider)','OPD Cover':'Optional (Rs.1L annual)',
    'Lock-the-Clock':'Premium stays locked','Priority Claim Desk':'Optional (HNI/Prime)',
    'Cash+ Wallet':'Cashback every claim-free year',
    'Unique Features':'Unlimited Sum Insured | Worldwide Cover | Lock-the-Clock | Cash+ Wallet',
  },
  'Niva Bupa - Aspire Platinum': {
    'Room Rent':'Single Private A/C Room','ICU Limit':'No limit on ICU',
    'Pre-Hospitalization':'60 days','Post-Hospitalization':'90 days',
    'Day Care Procedures':'All procedures up to S.I.','Copay':'0%',
    'Restoration Benefit':'Unlimited (incl. same illness)',
    'NCB Benefit':'Up to 5X (300%)','Non-Consumables':'Covered',
    'Hospitalization @ Home':'Up to Sum Assured','Ambulance Cover':'Up to Sum Assured',
    'Air Ambulance':'Up to Sum Assured','AYUSH Treatment':'Up to Sum Assured',
    'Organ Donor':'Up to Sum Assured','Modern Treatments':'Up to Sum Assured',
    '2-Hour Hospitalization':'—','E-Consultation':'Unlimited',
    'Preventive Health Check-up':'Standard',
    'Maternity Cover':'Standard Rs.12k/yr','Newborn Cover':'Available',
    'Worldwide Cover':'Not Available','OPD Cover':'Optional Rider',
    'Lock-the-Clock':'Premium stays locked','Priority Claim Desk':'—',
    'Cash+ Wallet':'—','Unique Features':'Premium lock | Child cover till 60 yrs | NCB up to 5X',
  },
  'Tata AIG - Medicare Select': {
    'Room Rent':'Single Private A/C Room','ICU Limit':'No limit on ICU',
    'Pre-Hospitalization':'60 days','Post-Hospitalization':'90 days',
    'Day Care Procedures':'All procedures up to S.I.','Copay':'0%',
    'Restoration Benefit':'Unlimited (incl. same illness)',
    'NCB Benefit':'100%/yr up to 500% (Super NCB)','Non-Consumables':'Covered',
    'Hospitalization @ Home':'Up to Sum Assured','Ambulance Cover':'Up to Sum Assured',
    'Air Ambulance':'Up to Sum Assured','AYUSH Treatment':'Up to Sum Assured',
    'Organ Donor':'Up to Sum Assured','Modern Treatments':'Up to Sum Assured',
    '2-Hour Hospitalization':'—','E-Consultation':'Unlimited',
    'Preventive Health Check-up':'All covered',
    'Maternity Cover':'Optional Rider (10% SA, Max Rs.1L)','Newborn Cover':'Available with rider',
    'Worldwide Cover':'Not Available','OPD Cover':'Optional Rider',
    'Lock-the-Clock':'Not Applicable','Priority Claim Desk':'—',
    'Cash+ Wallet':'—','Unique Features':'Salary discount 7.5% | Super NCB up to 500%',
  },
  'HDFC ERGO - Optima Secure': {
    'Room Rent':'Single Private A/C Room','ICU Limit':'No limit on ICU',
    'Pre-Hospitalization':'60 days','Post-Hospitalization':'180 days',
    'Day Care Procedures':'All procedures up to S.I.','Copay':'0%',
    'Restoration Benefit':'Unlimited (incl. same illness)',
    'NCB Benefit':'2X from Day 1','Non-Consumables':'Covered',
    'Hospitalization @ Home':'Up to Sum Assured','Ambulance Cover':'Up to Sum Assured',
    'Air Ambulance':'Up to Sum Assured','AYUSH Treatment':'Up to Sum Assured',
    'Organ Donor':'Up to Sum Assured','Modern Treatments':'Up to Sum Assured',
    '2-Hour Hospitalization':'—','E-Consultation':'Unlimited',
    'Preventive Health Check-up':'Standard',
    'Maternity Cover':'Not Available','Newborn Cover':'Not Available',
    'Worldwide Cover':'Not Available','OPD Cover':'Optional Rider',
    'Lock-the-Clock':'Not Applicable','Priority Claim Desk':'—',
    'Cash+ Wallet':'—','Unique Features':'2X Cover from Day 1 | Health check-up included',
  },
  'Care Health - Supreme': {
    'Room Rent':'Single Private A/C Room','ICU Limit':'No limit on ICU',
    'Pre-Hospitalization':'60 days','Post-Hospitalization':'180 days',
    'Day Care Procedures':'All procedures up to S.I.','Copay':'0%',
    'Restoration Benefit':'Unlimited (incl. same illness)',
    'NCB Benefit':'100%/yr up to 600% (6X)','Non-Consumables':'Covered',
    'Hospitalization @ Home':'Up to Sum Assured','Ambulance Cover':'Up to Sum Assured',
    'Air Ambulance':'Up to Sum Assured','AYUSH Treatment':'Up to Sum Assured',
    'Organ Donor':'Up to Sum Assured','Modern Treatments':'Up to Sum Assured',
    '2-Hour Hospitalization':'—','E-Consultation':'Unlimited',
    'Preventive Health Check-up':'All covered',
    'Maternity Cover':'Not Available','Newborn Cover':'Not Available',
    'Worldwide Cover':'Not Available','OPD Cover':'Optional Rider',
    'Lock-the-Clock':'Not Applicable','Priority Claim Desk':'—',
    'Cash+ Wallet':'—','Unique Features':'NCB up to 600% (6X) | All non-consumables covered',
  },
};

const HIGHLIGHTS = {
  'ICICI Lombard - Elevate':    ['Unlimited NCB (no cap)','Newborn Day-1 cover (10% SA)','2-hr hospitalization covered'],
  'Niva Bupa - ReAssure 3.0':   ['Truly Unlimited Sum Insured','Worldwide cover with rider','Lock-the-Clock premium freeze'],
  'Niva Bupa - Aspire Platinum':['Premium lock feature','Child cover till 60 yrs','NCB up to 5X (300%)'],
  'Tata AIG - Medicare Select': ['Salary-linked discount 7.5%','Super NCB up to 500%','Optional maternity/newborn rider'],
  'HDFC ERGO - Optima Secure':  ['2X cover from Day 1','Deductible on 1st claim','Health check-up included'],
  'Care Health - Supreme':      ['NCB up to 600% (6X)','All non-consumables covered','Health check-up rider'],
};

const DEFAULT_KEY_FEATURES = [
  'Room Rent','ICU Limit','Copay','Restoration Benefit','NCB Benefit',
  'Non-Consumables','Pre-Hospitalization','Post-Hospitalization',
  'Maternity Cover','Newborn Cover','Worldwide Cover','Unique Features',
];

// ── Logo helpers ──────────────────────────────────────────────────────────────
const LOGO_DIR  = path.join(__dirname, '../logos');
const PUBLIC_DIR = path.join(__dirname, '../public');
const LOGO_FILES = {
  icici: 'icici_lombard.jpg', niva: 'niva_reassure3.jpg',
  tata:  'tata_aig.jpg',      hdfc: 'hdfc_ergo.jpg',
  care:  'care_health.jpg',
};

function getLogoB64(name) {
  if (!name) return null;
  const n = name.toLowerCase();
  let key = null;
  if (n.includes('icici'))      key = 'icici';
  else if (n.includes('niva'))  key = 'niva';
  else if (n.includes('tata'))  key = 'tata';
  else if (n.includes('hdfc'))  key = 'hdfc';
  else if (n.includes('care'))  key = 'care';
  if (!key) return null;
  const file = path.join(LOGO_DIR, LOGO_FILES[key]);
  if (!fs.existsSync(file)) return null;
  return 'data:image/jpeg;base64,' + fs.readFileSync(file).toString('base64');
}

function getIncremintLogoB64() {
  const file = path.join(PUBLIC_DIR, 'logo.png');
  if (!fs.existsSync(file)) return null;
  return 'data:image/png;base64,' + fs.readFileSync(file).toString('base64');
}

// ── Name parsing ──────────────────────────────────────────────────────────────
function parseName(name) {
  // Split "ICICI Lombard – Elevate" or "ICICI Lombard - Elevate"
  const sep = name.includes('–') ? '–' : (name.includes('-') ? '-' : null);
  if (!sep) return { brand: name, sub: '' };
  const idx = name.indexOf(sep);
  return { brand: name.slice(0, idx).trim(), sub: name.slice(idx + 1).trim() };
}

// ── Master lookup ─────────────────────────────────────────────────────────────
function mapMaster(name) {
  if (!name) return null;
  const n = name.toLowerCase();
  if (n.includes('reassure') || n.includes('3.0')) return 'Niva Bupa - ReAssure 3.0';
  if (n.includes('aspire'))                         return 'Niva Bupa - Aspire Platinum';
  if (n.includes('icici') || n.includes('elevate')) return 'ICICI Lombard - Elevate';
  if (n.includes('tata')  || n.includes('medicare'))return 'Tata AIG - Medicare Select';
  if (n.includes('hdfc')  || n.includes('optima'))  return 'HDFC ERGO - Optima Secure';
  if (n.includes('care')  || n.includes('supreme')) return 'Care Health - Supreme';
  return null;
}

function getMasterData(name) {
  return MASTER[name] || MASTER[mapMaster(name)] || {};
}

// ── Number helpers ────────────────────────────────────────────────────────────
function safeNum(v) {
  if (!v) return 0;
  const n = parseFloat(String(v).replace(/,/g, '').trim());
  return isNaN(n) ? 0 : n;
}
function fmtPrem(v) {
  const n = safeNum(v);
  if (!n) return '—';
  return '₹' + n.toLocaleString('en-IN');
}
function esc(s) {
  if (!s) return '';
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

// ── Plan card (Page 1) ────────────────────────────────────────────────────────
function buildPlanCard(ins, idx) {
  const color    = INSURER_COLORS[idx % INSURER_COLORS.length];
  const { brand, sub } = parseName(ins.name || '');
  const logoB64  = getLogoB64(ins.name);
  const logoHTML = logoB64
    ? `<img src="${logoB64}" style="height:38px;max-width:130px;object-fit:contain;display:block;margin-bottom:8px" />`
    : `<div style="height:38px;display:flex;align-items:center;margin-bottom:8px;font-weight:800;font-size:16px;color:${color}">${esc(brand.charAt(0))}</div>`;

  const p1v  = safeNum(ins.premium_1yr);
  const p2v  = safeNum(ins.premium_2yr);
  const p3v  = safeNum(ins.premium_3yr);
  const sa   = esc(ins.sum_assured || '—');

  // Premium rows
  let premRows = `
    <div style="display:flex;justify-content:space-between;align-items:center;padding:8px 0;border-bottom:1px solid #f0f0f0">
      <span style="color:#666;font-size:12px">Sum Assured</span>
      <span style="font-weight:700;font-size:13px">${sa}</span>
    </div>
    <div style="display:flex;justify-content:space-between;align-items:center;padding:8px 0;border-bottom:1px solid #f0f0f0">
      <span style="color:#666;font-size:12px">1 Year</span>
      <span style="font-weight:700;font-size:15px;color:#1a1a1a">${fmtPrem(p1v)}</span>
    </div>`;

  if (p2v > 0) {
    const saving2 = (p1v * 2) - p2v;
    const pct2    = p1v > 0 ? Math.round(saving2 / (p1v * 2) * 100) : 0;
    const save2   = saving2 > 0 ? `<div style="color:${BRAND_DARK};font-size:11px;font-weight:600">Save ₹${Math.round(saving2).toLocaleString('en-IN')} (${pct2}% off)</div>` : '';
    premRows += `
    <div style="display:flex;justify-content:space-between;align-items:flex-start;padding:8px 0;border-bottom:1px solid #f0f0f0">
      <span style="color:#666;font-size:12px;padding-top:2px">2 Year</span>
      <div style="text-align:right">${save2}<span style="font-weight:700;font-size:15px;color:#1a1a1a">${fmtPrem(p2v)}</span></div>
    </div>`;
  }

  if (p3v > 0) {
    const saving3 = (p1v * 3) - p3v;
    const pct3    = p1v > 0 ? Math.round(saving3 / (p1v * 3) * 100) : 0;
    const save3   = saving3 > 0 ? `<div style="color:${BRAND_DARK};font-size:11px;font-weight:600">Save ₹${Math.round(saving3).toLocaleString('en-IN')} (${pct3}% off)</div>` : '';
    premRows += `
    <div style="display:flex;justify-content:space-between;align-items:flex-start;background:#FEF3C7;border-radius:6px;padding:8px 10px;border:1px solid #D97706">
      <div>
        <span style="color:#92400E;font-size:11.5px;font-weight:700">3 Year  ★ BEST VALUE</span>
      </div>
      <div style="text-align:right">${save3}<span style="font-weight:700;font-size:15px;color:#1a1a1a">${fmtPrem(p3v)}</span></div>
    </div>`;
  }

  return `
  <div style="border-radius:12px;border:2px solid ${color};box-shadow:0 2px 8px rgba(0,0,0,0.10);background:#EAF5EA;flex:1;min-width:0">
    <div style="padding:14px 14px 8px">
      ${logoHTML}
      <div style="font-weight:700;font-size:14px;color:#1a1a1a">${esc(brand)}</div>
      ${sub ? `<div style="font-size:11px;color:#888;margin-top:2px">${esc(sub)}</div>` : ''}
      <div style="height:1px;background:#c8e6c9;margin:10px 0 2px"></div>
    </div>
    <div style="padding:0 14px 12px;background:#EAF5EA">
      ${premRows}
    </div>
  </div>`;
}

// ── Features table (Page 2) ───────────────────────────────────────────────────
function buildFeaturesTable(insurers, featuresOverride) {
  const features = featuresOverride
    ? featuresOverride.map(f => f.name)
    : DEFAULT_KEY_FEATURES;

  const colW = Math.floor(82 / insurers.length);
  // Header row
  let headerCells = `<th style="background:${BRAND_GREEN};color:white;font-size:11px;font-weight:700;text-align:left;padding:10px 10px;border:1px solid rgba(255,255,255,0.3);width:18%">FEATURE</th>`;
  insurers.forEach((ins, idx) => {
    const color = INSURER_COLORS[idx % INSURER_COLORS.length];
    const { brand, sub } = parseName(ins.name || '');
    const logoB64 = getLogoB64(ins.name);
    const logoHTML = logoB64
      ? `<img src="${logoB64}" style="height:26px;max-width:72px;object-fit:contain;display:block;margin:0 auto 5px" />`
      : '';
    headerCells += `<th style="background:#EAF5EA;border:1px solid #c8e6c9;border-bottom:3px solid ${color};padding:8px 6px;text-align:center;vertical-align:bottom;width:${colW}%">
      ${logoHTML}
      <div style="font-weight:700;font-size:11px;color:${color}">${esc(brand)}</div>
      ${sub ? `<div style="font-weight:400;font-size:10px;color:#555;margin-top:2px">${esc(sub)}</div>` : ''}
    </th>`;
  });

  // Helper: render unique features as bullet list
  function renderUniqVal(val) {
    if (!val || val === '—') return '—';
    const parts = val.split(/\s*[|,]\s*/).filter(Boolean);
    if (parts.length <= 1) return esc(val);
    return parts.map(p =>
      `<div style="display:flex;align-items:flex-start;gap:4px;margin:2px 0;text-align:left">
        <span style="color:#92400E;font-weight:700;flex-shrink:0">•</span>
        <span>${esc(p.trim())}</span>
      </div>`
    ).join('');
  }

  // Data rows
  let dataRows = '';
  features.forEach((feat, fi) => {
    const isUniq  = feat === 'Unique Features';
    const rowBg   = isUniq ? '#FEF3C7' : (fi % 2 === 0 ? '#f8f8f8' : 'white');
    const lblClr  = isUniq ? '#92400E' : '#1a1a1a';
    const lblW    = isUniq ? '700' : '600';
    let cells = `<td style="padding:8px 10px;border:1px solid #e0e0e0;font-weight:${lblW};font-size:11px;color:${lblClr};background:${isUniq ? '#FEF3C7' : rowBg}">${esc(feat)}</td>`;
    insurers.forEach((ins, idx) => {
      let val = '—';
      if (featuresOverride && fi < featuresOverride.length) {
        const fvals = featuresOverride[fi].values || {};
        for (const [k, v] of Object.entries(fvals)) {
          const mk = mapMaster(k) || k;
          if (mk === ins.name || k === ins.name || mapMaster(ins.name) === mk) {
            val = v || '—'; break;
          }
        }
      } else {
        const md  = getMasterData(ins.name);
        val = md[feat] || '—';
      }
      const isNA  = val.toLowerCase().includes('not available') || val.toLowerCase().includes('not applicable');
      const vClr  = isNA ? '#D32F2F' : (isUniq ? '#92400E' : '#333');
      const cellBg = isUniq ? '#FEF3C7' : rowBg;
      const cellContent = isUniq ? renderUniqVal(val) : esc(val);
      cells += `<td style="padding:8px 6px;border:1px solid #e0e0e0;text-align:center;font-size:11px;color:${vClr};background:${cellBg}">${cellContent}</td>`;
    });
    dataRows += `<tr>${cells}</tr>`;
  });

  return `
  <table style="width:100%;border-collapse:collapse;font-size:11px;table-layout:fixed">
    <thead><tr>${headerCells}</tr></thead>
    <tbody>${dataRows}</tbody>
  </table>`;
}

// ── Why Choose cards (Page 2) ─────────────────────────────────────────────────
function buildWhyChooseCards(insurers, highlightsOverride) {
  return insurers.map((ins, idx) => {
    const color  = INSURER_COLORS[idx % INSURER_COLORS.length];
    const { brand, sub } = parseName(ins.name || '');
    const logoB64 = getLogoB64(ins.name);
    const logoHTML = logoB64
      ? `<img src="${logoB64}" style="width:36px;height:36px;border-radius:50%;object-fit:contain;border:1.5px solid #e0e0e0;padding:2px;flex-shrink:0" />`
      : `<div style="width:36px;height:36px;border-radius:50%;background:${color};display:flex;align-items:center;justify-content:center;color:white;font-weight:700;font-size:15px;flex-shrink:0">${esc(brand.charAt(0))}</div>`;

    let bullets = [];
    if (highlightsOverride) {
      for (const [k, buls] of Object.entries(highlightsOverride)) {
        const mk = mapMaster(k) || k;
        if (mk === ins.name || k === ins.name || mapMaster(ins.name) === mk) {
          bullets = buls || []; break;
        }
      }
    } else {
      const mk = mapMaster(ins.name) || ins.name;
      bullets = HIGHLIGHTS[ins.name] || HIGHLIGHTS[mk] || [];
    }

    const bulletHTML = bullets.map(b =>
      `<div style="display:flex;align-items:flex-start;gap:6px;margin:4px 0;font-size:11px;color:#333">
        <span style="color:${BRAND_GREEN};font-size:13px;line-height:1.4;flex-shrink:0">•</span>
        <span>${esc(b)}</span>
      </div>`
    ).join('');

    return `
    <div style="flex:1;border:1.5px solid ${color};border-radius:10px;padding:12px 14px;background:#EAF5EA;min-width:0">
      <div style="display:flex;align-items:center;gap:10px;margin-bottom:10px">
        ${logoHTML}
        <div>
          <div style="font-weight:700;font-size:13px;color:${color}">${esc(brand)}</div>
          ${sub ? `<div style="font-size:11px;color:#888;margin-top:1px">${esc(sub)}</div>` : ''}
        </div>
      </div>
      ${bulletHTML}
    </div>`;
  }).join('');
}

// ── Members table ─────────────────────────────────────────────────────────────
function buildMembersTable(members) {
  const rows = (members || []).map((m, i) => {
    const bg = i % 2 === 0 ? '#EAF5EA' : 'white';
    return `<tr style="background:${bg}">
      <td style="padding:7px 10px;border:1px solid #e0e0e0;font-weight:700;color:${BRAND_GREEN}">${i + 1}</td>
      <td style="padding:7px 10px;border:1px solid #e0e0e0;font-weight:600">${esc(m.name || m.relation || '')}</td>
      <td style="padding:7px 10px;border:1px solid #e0e0e0">${esc(m.relation || '')}</td>
      <td style="padding:7px 10px;border:1px solid #e0e0e0;text-align:center">${esc(m.dob || '')}</td>
      <td style="padding:7px 10px;border:1px solid #e0e0e0;text-align:center">${esc(m.age || '')}</td>
      <td style="padding:7px 10px;border:1px solid #e0e0e0;text-align:center">${esc(m.city || '')}</td>
      <td style="padding:7px 10px;border:1px solid #e0e0e0;text-align:center;font-weight:700">${esc(m.sum_assured || '')}</td>
    </tr>`;
  }).join('');

  return `
  <table style="width:100%;border-collapse:collapse;font-size:12px">
    <thead>
      <tr style="background:${BRAND_GREEN};color:white">
        <th style="padding:8px 10px;border:1px solid rgba(255,255,255,0.3);text-align:left;font-size:11px;width:4%">#</th>
        <th style="padding:8px 10px;border:1px solid rgba(255,255,255,0.3);text-align:left;font-size:11px;width:20%">NAME</th>
        <th style="padding:8px 10px;border:1px solid rgba(255,255,255,0.3);text-align:left;font-size:11px;width:10%">RELATION</th>
        <th style="padding:8px 10px;border:1px solid rgba(255,255,255,0.3);text-align:center;font-size:11px;width:12%">DATE OF BIRTH</th>
        <th style="padding:8px 10px;border:1px solid rgba(255,255,255,0.3);text-align:center;font-size:11px;width:6%">AGE</th>
        <th style="padding:8px 10px;border:1px solid rgba(255,255,255,0.3);text-align:center;font-size:11px;width:12%">CITY</th>
        <th style="padding:8px 10px;border:1px solid rgba(255,255,255,0.3);text-align:center;font-size:11px;width:12%">SUM ASSURED</th>
      </tr>
    </thead>
    <tbody>${rows}</tbody>
  </table>`;
}

// ── Section bar ───────────────────────────────────────────────────────────────
function sectionBar(text) {
  return `<div style="background:${BRAND_GREEN};color:white;font-weight:700;font-size:12px;padding:9px 14px;border-radius:8px;margin:14px 0 8px">${esc(text)}</div>`;
}

// ── Header block ──────────────────────────────────────────────────────────────
function buildHeader(data) {
  const customerName  = data.customer_name  ? esc(data.customer_name.toUpperCase()) : '';
  const customerPhone = data.customer_phone ? esc(data.customer_phone) : '';
  const customerCity  = data.city           ? esc(data.city)           : '';

  const customerMeta = [
    customerPhone ? `Mobile: <b>${customerPhone}</b>` : '',
    customerCity  ? esc(customerCity)                 : '',
  ].filter(Boolean).join('&nbsp;&nbsp;|&nbsp;&nbsp;');

  return `
  <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px">
    <div>
      ${customerName ? `<div style="font-weight:800;font-size:15px;color:${BRAND_GREEN}">${customerName}</div>` : ''}
      ${customerMeta ? `<div style="font-size:11px;color:#555;margin-top:2px">${customerMeta}</div>` : ''}
    </div>
    <div style="text-align:right;font-size:11px;color:#666;line-height:1.7">
      <div><b style="color:#1a1a1a">Quote ID: ${esc(data.quote_id || '')}</b></div>
      <div>${esc(data.date || '')}</div>
    </div>
  </div>
  <div style="height:2px;background:${BRAND_GREEN};border-radius:1px;margin-bottom:10px"></div>`;
}

// ── Footer ────────────────────────────────────────────────────────────────────
function buildFooter(data) {
  const agentLine = (data.agent_name || data.agent_mobile)
    ? `<div style="display:flex;justify-content:space-between;align-items:center;padding:5px 10px;background:#f9fef9;border:1px solid #c8e6c9;border-radius:6px;margin-bottom:5px;font-size:9.5px">
        <span style="color:#555">For more details, contact your advisor:</span>
        <span style="display:flex;gap:12px;align-items:center">
          ${data.agent_name   ? `<b style="color:#1a1a1a">${esc(data.agent_name)}</b>`                          : ''}
          ${data.agent_code   ? `<span style="color:#666">Code: <b>${esc(data.agent_code)}</b></span>`           : ''}
          ${data.agent_mobile ? `<span style="color:#666">📞 <b>${esc(data.agent_mobile)}</b></span>`            : ''}
        </span>
      </div>`
    : '';
  return `
  <div style="margin-top:10px">
    ${agentLine}
    <div style="padding-top:5px;border-top:1px solid #e0e0e0;display:flex;justify-content:space-between;font-size:9px;color:#999">
      <span>Quote ID: ${esc(data.quote_id || '')}  |  ${esc(data.customer_name || '')}  |  ${esc(data.date || '')}</span>
      <span>Premiums are indicative. Final premium subject to underwriting. IRDAI Registered Broker.</span>
    </div>
  </div>`;
}

// ── Main export ───────────────────────────────────────────────────────────────
function buildQuoteHTML(data) {
  const insurers          = data.insurers || [];
  const members           = data.members  || [];
  const featuresOverride  = data.features_override  || null;
  const highlightsOverride= data.highlights_override || null;

  const planCards = insurers.map((ins, idx) => buildPlanCard(ins, idx)).join('');

  // Advisor note (shown at bottom of page 2) — premium callout style
  const advisorNote = data.advisor_note
    ? `<div style="margin-top:14px;padding:12px 18px;border-left:5px solid ${BRAND_GREEN};border-radius:0 8px 8px 0;background:white;box-shadow:0 2px 8px rgba(76,175,80,0.15);font-size:12px;display:flex;align-items:flex-start;gap:10px">
        <div style="color:${BRAND_GREEN};font-size:20px;line-height:1;flex-shrink:0">💬</div>
        <div>
          <div style="font-weight:700;color:${BRAND_DARK};font-size:11px;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:3px">Advisor's Recommendation</div>
          <div style="color:#333;font-style:italic;line-height:1.5">${esc(data.advisor_note)}</div>
        </div>
      </div>`
    : '';

  const featTable    = buildFeaturesTable(insurers, featuresOverride);
  const whyCards     = buildWhyChooseCards(insurers, highlightsOverride);
  const membersTable = buildMembersTable(members);
  const headerHTML   = buildHeader(data);
  const footerHTML   = buildFooter(data);

  return `<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
  @page { size: A4 landscape; margin: 0.7cm; }
  * { -webkit-print-color-adjust: exact; print-color-adjust: exact; box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Arial, sans-serif; font-size: 12px; color: #1a1a1a; background: white; overflow-x: hidden; }
  .page { padding: 2px 0; width: 100%; overflow: hidden; }
  .page-break { page-break-after: always; page-break-inside: avoid; }
  td, th { word-wrap: break-word; overflow-wrap: break-word; }
</style>
</head>
<body>

<!-- ══════════════ PAGE 1 ══════════════ -->
<div class="page page-break">
  ${headerHTML}

  ${sectionBar('INSURED MEMBERS')}
  ${membersTable}

  ${sectionBar('PLAN COMPARISON — PREMIUMS')}
  <div style="display:flex;gap:14px;align-items:stretch;justify-content:center">
    ${planCards}
  </div>

  ${footerHTML}
</div>

<!-- ══════════════ PAGE 2 ══════════════ -->
<div class="page">
  ${headerHTML}

  ${sectionBar('KEY BENEFITS COMPARISON')}
  ${featTable}

  ${sectionBar('WHY CHOOSE EACH PLAN')}
  <div style="display:flex;gap:12px;margin-top:4px;page-break-inside:avoid">
    ${whyCards}
  </div>

  ${advisorNote}
  ${footerHTML}
</div>

</body>
</html>`;
}

module.exports = { buildQuoteHTML };
