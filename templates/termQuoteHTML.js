'use strict';
const fs   = require('fs');
const path = require('path');

// ── Brand colors ─────────────────────────────────────────────────────────────
const INSURER_COLORS = ['#1565C0','#2E7D32','#283593','#B71C1C','#E65100','#880E4F'];
const BRAND_GREEN    = '#1F4E27';
const BRAND_LIGHT    = '#4CAF50';
const GREEN_OK       = '#E2EFDA';
const AMBER          = '#FFF2CC';

// ── Logo helpers ─────────────────────────────────────────────────────────────
const LOGO_DIR   = path.join(__dirname, '../logos');
const PUBLIC_DIR = path.join(__dirname, '../public');

const TERM_LOGO_MAP = {
  icici:  'icici_pru.jpg',
  hdfc:   'hdfc_life.jpg',
  tata:   'tata_aia.jpg',
  bajaj:  'bajaj_allianz.jpg',
  max:    'max_life.jpg',
  axis:   'max_life.jpg',
};

function getTermLogoB64(name) {
  if (!name) return null;
  const n = name.toLowerCase();
  let key = null;
  if (n.includes('icici'))               key = 'icici';
  else if (n.includes('hdfc'))           key = 'hdfc';
  else if (n.includes('tata'))           key = 'tata';
  else if (n.includes('bajaj'))          key = 'bajaj';
  else if (n.includes('max') || n.includes('axis')) key = 'max';
  if (!key || !TERM_LOGO_MAP[key]) return null;
  const file = path.join(LOGO_DIR, TERM_LOGO_MAP[key]);
  if (!fs.existsSync(file)) return null;
  return 'data:image/jpeg;base64,' + fs.readFileSync(file).toString('base64');
}

function getIncremintLogoB64() {
  const file = path.join(PUBLIC_DIR, 'logo.png');
  if (!fs.existsSync(file)) return null;
  return 'data:image/png;base64,' + fs.readFileSync(file).toString('base64');
}

// ── Helpers ──────────────────────────────────────────────────────────────────
function esc(s) {
  if (!s) return '';
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}
function safeNum(v) {
  if (!v) return 0;
  const n = parseFloat(String(v).replace(/,/g, '').trim());
  return isNaN(n) ? 0 : n;
}
function money(n) {
  const v = safeNum(n);
  if (!v) return '—';
  return '₹' + v.toLocaleString('en-IN');
}
function parseName(name) {
  const sep = name.includes('–') ? '–' : (name.includes('-') ? '-' : null);
  if (!sep) return { brand: name, sub: '' };
  const idx = name.indexOf(sep);
  return { brand: name.slice(0, idx).trim(), sub: name.slice(idx + 1).trim() };
}

// ── Client details table ─────────────────────────────────────────────────────
function buildClientTable(data) {
  const fields1 = [
    { label: 'Client', value: data.customer_name || '' },
    { label: 'DOB', value: data.dob || '' },
    { label: 'Age', value: data.age || '' },
    { label: 'City', value: data.city || '' },
  ];
  const fields2 = [
    { label: 'Sum Assured', value: data.sum_assured || '' },
    { label: 'Policy Term', value: data.policy_term ? `${data.policy_term} yrs` : '' },
    { label: 'Cover Till Age', value: data.cover_till_age || '' },
    { label: 'PPT Options', value: data.ppt_text || '' },
  ];

  function row(fields) {
    const hdr = fields.map(f =>
      `<th style="background:${BRAND_GREEN};color:white;font-size:11px;font-weight:700;padding:9px 12px;border:1px solid rgba(255,255,255,0.3);text-align:center">${esc(f.label)}</th>`
    ).join('');
    const vals = fields.map(f =>
      `<td style="padding:8px 12px;border:1px solid #e0e0e0;text-align:center;font-size:12px;font-weight:600">${esc(f.value)}</td>`
    ).join('');
    return `<table style="width:100%;border-collapse:collapse;margin-bottom:4px"><tr>${hdr}</tr><tr>${vals}</tr></table>`;
  }

  return row(fields1) + row(fields2);
}

// ── Premium comparison table ─────────────────────────────────────────────────
function buildPremiumTable(data) {
  const insurers = data.insurers || [];
  const labelA = data.option_a || 'Regular';
  const labelB = data.option_b || '10 Pay';
  const yrsA   = data.years_a  || 0;
  const yrsB   = data.years_b  || 0;

  // Header
  const cols = [
    '', 'Company', `${labelA}<br>(Annual)`, `${labelB}<br>(Annual)`,
    `Total ${labelA}<br>(${yrsA} yrs)`, `Total ${labelB}<br>(${yrsB} yrs)`,
    `Savings with ${labelB}`, 'Savings %'
  ];

  let headerHTML = cols.map((c, i) => {
    const w = i === 0 ? '8%' : i === 1 ? '18%' : '10.5%';
    return `<th style="background:${BRAND_GREEN};color:white;font-size:10px;font-weight:700;padding:10px 6px;border:1px solid rgba(255,255,255,0.3);text-align:center;width:${w};vertical-align:middle">${c}</th>`;
  }).join('');

  let rowsHTML = '';
  let hasRows = false;

  insurers.forEach((ins, idx) => {
    const premA = safeNum(ins.premium_a);
    const premB = safeNum(ins.premium_b);
    if (premA <= 0 || premB <= 0) return;
    hasRows = true;

    const { brand, sub } = parseName(ins.name || '');
    const logoB64 = getTermLogoB64(ins.name);
    const totA = premA * yrsA;
    const totB = premB * yrsB;
    const delta = totA - totB;
    const pct = totA > 0 ? (Math.abs(delta) / totA * 100).toFixed(1) : '0.0';

    const bg = idx % 2 === 0 ? '#f8fef8' : 'white';

    const logoCell = logoB64
      ? `<img src="${logoB64}" style="height:36px;max-width:70px;object-fit:contain;display:block;margin:0 auto" />`
      : `<div style="width:36px;height:36px;border-radius:50%;background:${INSURER_COLORS[idx % INSURER_COLORS.length]};display:flex;align-items:center;justify-content:center;color:white;font-weight:700;font-size:15px;margin:0 auto">${esc(brand.charAt(0))}</div>`;

    let savingsCell, pctCell;
    if (delta > 0) {
      savingsCell = `<td style="padding:8px 6px;border:1px solid #e0e0e0;text-align:center;font-weight:700;font-size:11px;background:${GREEN_OK};color:#1F4E27">✅ ${money(delta)}</td>`;
      pctCell = `<td style="padding:8px 6px;border:1px solid #e0e0e0;text-align:center;font-weight:700;font-size:11px;background:${GREEN_OK};color:#1F4E27">${pct}%</td>`;
    } else if (delta < 0) {
      savingsCell = `<td style="padding:8px 6px;border:1px solid #e0e0e0;text-align:center;font-weight:700;font-size:11px;background:${AMBER};color:#92400E">⚠️ Extra ${money(Math.abs(delta))}</td>`;
      pctCell = `<td style="padding:8px 6px;border:1px solid #e0e0e0;text-align:center;font-weight:700;font-size:11px;background:${AMBER};color:#92400E">${pct}%</td>`;
    } else {
      savingsCell = `<td style="padding:8px 6px;border:1px solid #e0e0e0;text-align:center;font-size:11px">—</td>`;
      pctCell = `<td style="padding:8px 6px;border:1px solid #e0e0e0;text-align:center;font-size:11px">0.0%</td>`;
    }

    rowsHTML += `<tr style="background:${bg}">
      <td style="padding:6px;border:1px solid #e0e0e0;text-align:center">${logoCell}</td>
      <td style="padding:8px 8px;border:1px solid #e0e0e0;text-align:left;font-size:11px"><b>${esc(brand)}</b>${sub ? `<br><span style="color:#888;font-size:10px">${esc(sub)}</span>` : ''}</td>
      <td style="padding:8px 6px;border:1px solid #e0e0e0;text-align:center;font-weight:700;font-size:12px">${money(premA)}</td>
      <td style="padding:8px 6px;border:1px solid #e0e0e0;text-align:center;font-weight:700;font-size:12px">${money(premB)}</td>
      <td style="padding:8px 6px;border:1px solid #e0e0e0;text-align:center;font-size:11px">${money(totA)}</td>
      <td style="padding:8px 6px;border:1px solid #e0e0e0;text-align:center;font-size:11px">${money(totB)}</td>
      ${savingsCell}
      ${pctCell}
    </tr>`;
  });

  if (!hasRows) {
    return `<div style="padding:20px;text-align:center;color:#999;font-size:12px">No rows with both premium options filled.</div>`;
  }

  return `
  <table style="width:100%;border-collapse:collapse;font-size:11px;table-layout:fixed">
    <thead><tr>${headerHTML}</tr></thead>
    <tbody>${rowsHTML}</tbody>
  </table>`;
}

// ── Features table ───────────────────────────────────────────────────────────
function buildTermFeaturesTable(insurers, featuresOverride) {
  if (!featuresOverride || featuresOverride.length === 0) return '';

  const features = featuresOverride.map(f => f.name);
  const colW = Math.floor(80 / insurers.length);

  let headerCells = `<th style="background:${BRAND_GREEN};color:white;font-size:11px;font-weight:700;text-align:left;padding:10px 10px;border:1px solid rgba(255,255,255,0.3);width:20%">FEATURE</th>`;
  insurers.forEach((ins, idx) => {
    const color = INSURER_COLORS[idx % INSURER_COLORS.length];
    const { brand, sub } = parseName(ins.name || '');
    const logoB64 = getTermLogoB64(ins.name);
    const logoHTML = logoB64
      ? `<img src="${logoB64}" style="height:24px;max-width:65px;object-fit:contain;display:block;margin:0 auto 4px" />`
      : '';
    headerCells += `<th style="background:#EAF5EA;border:1px solid #c8e6c9;border-bottom:3px solid ${color};padding:8px 6px;text-align:center;vertical-align:bottom;width:${colW}%">
      ${logoHTML}
      <div style="font-weight:700;font-size:10px;color:${color}">${esc(brand)}</div>
      ${sub ? `<div style="font-weight:400;font-size:9px;color:#555;margin-top:1px">${esc(sub)}</div>` : ''}
    </th>`;
  });

  let dataRows = '';
  features.forEach((feat, fi) => {
    const rowBg = fi % 2 === 0 ? '#f8f8f8' : 'white';
    let cells = `<td style="padding:8px 10px;border:1px solid #e0e0e0;font-weight:600;font-size:11px;color:#1a1a1a;background:${rowBg}">${esc(feat)}</td>`;
    insurers.forEach(ins => {
      let val = '—';
      if (fi < featuresOverride.length) {
        const fvals = featuresOverride[fi].values || {};
        for (const [k, v] of Object.entries(fvals)) {
          if (k === ins.name || (ins.name && k.toLowerCase().includes(ins.name.split('–')[0].trim().toLowerCase().split(' ')[0]))) {
            val = v || '—'; break;
          }
        }
      }
      const isNA = val.toLowerCase().includes('not available');
      const vClr = isNA ? '#D32F2F' : '#333';
      cells += `<td style="padding:8px 6px;border:1px solid #e0e0e0;text-align:center;font-size:11px;color:${vClr};background:${rowBg}">${esc(val)}</td>`;
    });
    dataRows += `<tr>${cells}</tr>`;
  });

  return `
  <table style="width:100%;border-collapse:collapse;font-size:11px;table-layout:fixed">
    <thead><tr>${headerCells}</tr></thead>
    <tbody>${dataRows}</tbody>
  </table>`;
}

// ── Why Choose cards ─────────────────────────────────────────────────────────
function buildTermWhyCards(insurers, highlightsOverride) {
  return insurers.map((ins, idx) => {
    const color = INSURER_COLORS[idx % INSURER_COLORS.length];
    const { brand, sub } = parseName(ins.name || '');
    const logoB64 = getTermLogoB64(ins.name);
    const logoHTML = logoB64
      ? `<img src="${logoB64}" style="width:34px;height:34px;border-radius:50%;object-fit:contain;border:1.5px solid #e0e0e0;padding:2px;flex-shrink:0" />`
      : `<div style="width:34px;height:34px;border-radius:50%;background:${color};display:flex;align-items:center;justify-content:center;color:white;font-weight:700;font-size:14px;flex-shrink:0">${esc(brand.charAt(0))}</div>`;

    let bullets = [];
    if (highlightsOverride) {
      for (const [k, buls] of Object.entries(highlightsOverride)) {
        if (k === ins.name) { bullets = buls || []; break; }
      }
    }

    const bulletHTML = bullets.map(b =>
      `<div style="display:flex;align-items:flex-start;gap:5px;margin:3px 0;font-size:10px;color:#333">
        <span style="color:${BRAND_LIGHT};font-size:12px;line-height:1.4;flex-shrink:0">•</span>
        <span>${esc(b)}</span>
      </div>`
    ).join('');

    return `
    <div style="flex:1;border:1.5px solid ${color};border-radius:10px;padding:10px 12px;background:#EAF5EA;min-width:0">
      <div style="display:flex;align-items:center;gap:8px;margin-bottom:8px">
        ${logoHTML}
        <div>
          <div style="font-weight:700;font-size:12px;color:${color}">${esc(brand)}</div>
          ${sub ? `<div style="font-size:10px;color:#888;margin-top:1px">${esc(sub)}</div>` : ''}
        </div>
      </div>
      ${bulletHTML}
    </div>`;
  }).join('');
}

// ── Section bar ──────────────────────────────────────────────────────────────
function sectionBar(text) {
  return `<div style="background:${BRAND_GREEN};color:white;font-weight:700;font-size:12px;padding:9px 14px;border-radius:8px;margin:14px 0 8px">${esc(text)}</div>`;
}

// ── Header ───────────────────────────────────────────────────────────────────
function buildHeader(data) {
  const incLogo = getIncremintLogoB64();
  const logoHTML = incLogo
    ? `<img src="${incLogo}" style="height:32px;object-fit:contain" />`
    : `<span style="font-weight:800;font-size:14px;color:${BRAND_GREEN}">incremint</span>`;

  return `
  <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px">
    <div>
      <div style="font-weight:800;font-size:16px;color:${BRAND_GREEN}">TERM LIFE INSURANCE – QUOTE</div>
      <div style="font-size:11px;color:#555;margin-top:2px">
        ${data.customer_name ? `for <b>${esc(data.customer_name.toUpperCase())}</b>` : ''}
        ${data.customer_phone ? `&nbsp;|&nbsp;📞 ${esc(data.customer_phone)}` : ''}
      </div>
    </div>
    <div style="text-align:right">
      ${logoHTML}
      <div style="font-size:10px;color:#666;margin-top:2px">
        <b>Quote ID: ${esc(data.quote_id || '')}</b> &nbsp;|&nbsp; ${esc(data.date || '')}
      </div>
    </div>
  </div>
  <div style="height:2px;background:${BRAND_GREEN};border-radius:1px;margin-bottom:10px"></div>`;
}

// ── Footer ───────────────────────────────────────────────────────────────────
function buildFooter(data) {
  const agentLine = (data.agent_name || data.agent_mobile)
    ? `<div style="display:flex;justify-content:space-between;align-items:center;padding:5px 10px;background:#f9fef9;border:1px solid #c8e6c9;border-radius:6px;margin-bottom:5px;font-size:9.5px">
        <span style="color:#555">For more details, contact your advisor:</span>
        <span style="display:flex;gap:12px;align-items:center">
          ${data.agent_name   ? `<b style="color:#1a1a1a">${esc(data.agent_name)}</b>`               : ''}
          ${data.agent_code   ? `<span style="color:#666">Code: <b>${esc(data.agent_code)}</b></span>` : ''}
          ${data.agent_mobile ? `<span style="color:#666">📞 <b>${esc(data.agent_mobile)}</b></span>`  : ''}
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

// ── Advisory note ────────────────────────────────────────────────────────────
function buildAdvisoryNote(data) {
  const labelA = data.option_a || 'Regular';
  const labelB = data.option_b || '10 Pay';
  const insurers = data.insurers || [];

  // Auto-generate summary from first comparable row
  let summaryLine = '';
  for (const ins of insurers) {
    const a = safeNum(ins.premium_a);
    const b = safeNum(ins.premium_b);
    if (a > 0 && b > 0) {
      const totA = a * (data.years_a || 0);
      const totB = b * (data.years_b || 0);
      const diff = totA - totB;
      if (diff > 0) {
        summaryLine = `Choosing <b>${esc(labelB)}</b> saves ${money(diff)} vs ${esc(labelA)} overall for ${esc(parseName(ins.name).brand)}.`;
      } else if (diff < 0) {
        summaryLine = `Choosing <b>${esc(labelB)}</b> costs ${money(Math.abs(diff))} more vs ${esc(labelA)} overall for ${esc(parseName(ins.name).brand)}.`;
      } else {
        summaryLine = `${esc(labelA)} and ${esc(labelB)} are equal on lifetime outgo.`;
      }
      break;
    }
  }

  const noteText = data.advisor_note || '';

  return `
  <div style="margin-top:12px;padding:12px 18px;border-left:5px solid ${BRAND_GREEN};border-radius:0 8px 8px 0;background:white;box-shadow:0 2px 8px rgba(31,78,39,0.12);font-size:12px">
    <div style="font-weight:700;color:${BRAND_GREEN};font-size:11px;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:6px">💡 Advisory Note</div>
    ${summaryLine ? `<div style="margin-bottom:6px;color:#333">${summaryLine}</div>` : ''}
    <div style="color:#555;font-size:11px;line-height:1.5">
      • Totals use pay years: Regular → Policy Term; 5/10/15 Pay → 5/10/15 years; Pay till 60 → (60 − Age).<br>
      • Tip: Limited-pay finishes premiums early and may reduce lifetime outgo.
    </div>
    ${noteText ? `<div style="margin-top:8px;color:#333;font-style:italic;border-top:1px solid #e0e0e0;padding-top:6px">"${esc(noteText)}"</div>` : ''}
  </div>`;
}

// ── Main export ──────────────────────────────────────────────────────────────
function buildTermQuoteHTML(data) {
  const insurers          = data.insurers || [];
  const featuresOverride  = data.features_override  || null;
  const highlightsOverride= data.highlights_override || null;

  const headerHTML     = buildHeader(data);
  const footerHTML     = buildFooter(data);
  const clientTable    = buildClientTable(data);
  const premiumTable   = buildPremiumTable(data);
  const featTable      = buildTermFeaturesTable(insurers, featuresOverride);
  const whyCards       = buildTermWhyCards(insurers, highlightsOverride);
  const advisoryNote   = buildAdvisoryNote(data);

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
  ${sectionBar('CLIENT DETAILS')}
  ${clientTable}
  ${sectionBar('PREMIUM COMPARISON — ' + esc(data.option_a || 'Regular') + ' vs ' + esc(data.option_b || '10 Pay'))}
  ${premiumTable}
  ${advisoryNote}
  ${footerHTML}
</div>

<!-- ══════════════ PAGE 2 ══════════════ -->
<div class="page">
  ${headerHTML}
  ${sectionBar('KEY FEATURES COMPARISON')}
  ${featTable}
  ${sectionBar('WHY CHOOSE EACH PLAN')}
  <div style="display:flex;gap:10px;margin-top:4px;page-break-inside:avoid">
    ${whyCards}
  </div>
  ${footerHTML}
</div>

</body>
</html>`;
}

module.exports = { buildTermQuoteHTML };
