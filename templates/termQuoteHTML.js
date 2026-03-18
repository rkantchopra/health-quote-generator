'use strict';
const fs   = require('fs');
const path = require('path');

// ── Brand colors ─────────────────────────────────────────────────────────────
const INSURER_COLORS = ['#1565C0','#2E7D32','#283593','#B71C1C','#E65100','#880E4F'];
const BRAND_GREEN    = '#1F4E27';
const BRAND_LIGHT    = '#4CAF50';
const GREEN_OK       = '#E2EFDA';
const AMBER          = '#FFF2CC';
const HIGHLIGHT_GREEN = '#C8F7C8';

// ── Logo helpers ─────────────────────────────────────────────────────────────
const LOGO_DIR   = path.join(__dirname, '../logos');

const TERM_LOGO_MAP = {
  icici:  'ICICIPru_logo.webp',
  hdfc:   'HDFC_Life_logo.avif',
  tata:   'TATA_AIA_logo.png',
  bajaj:  'BAJAJ_logo.avif',
  max:    'MAX_logo.avif',
  axis:   'MAX_logo.avif',
  birla:  'Birla_Sun_Life_logo.avif',
  kotak:  'kotak_logo.avif',
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
  else if (n.includes('birla') || n.includes('aditya')) key = 'birla';
  else if (n.includes('kotak'))          key = 'kotak';
  if (!key || !TERM_LOGO_MAP[key]) return null;
  const file = path.join(LOGO_DIR, TERM_LOGO_MAP[key]);
  if (!fs.existsSync(file)) return null;
  const ext = path.extname(file).toLowerCase();
  const mime = ext === '.png' ? 'image/png' : ext === '.webp' ? 'image/webp' : ext === '.avif' ? 'image/avif' : 'image/jpeg';
  return `data:${mime};base64,` + fs.readFileSync(file).toString('base64');
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
  const empLabel = data.employment_type === 'salaried' ? 'Salaried' : 'Self-Employed';
  const smokerLabel = data.smoker_status === 'smoker' ? '🚬 Smoker' : 'Non-Smoker';

  const fields1 = [
    { label: 'Client', value: data.customer_name || '' },
    { label: 'DOB', value: data.dob || '' },
    { label: 'Age', value: data.age || '' },
    { label: 'City', value: data.city || '' },
    { label: 'Pincode', value: data.pincode || '' },
  ];
  const fields2 = [
    { label: 'Employment', value: empLabel },
    { label: 'Smoker Status', value: smokerLabel },
    { label: 'Sum Assured', value: data.sum_assured || '', highlight: true },
    { label: 'Cover Till Age', value: data.cover_till_age || '', highlight: true },
    { label: 'Policy Term', value: data.policy_term ? `${data.policy_term} yrs` : '' },
  ];

  function row(fields) {
    const hdr = fields.map(f =>
      `<th style="background:${BRAND_GREEN};color:white;font-size:11px;font-weight:700;padding:9px 12px;border:1px solid rgba(255,255,255,0.3);text-align:center">${esc(f.label)}</th>`
    ).join('');
    const vals = fields.map(f => {
      const hlStyle = f.highlight
        ? `background:#E8F5E9;color:${BRAND_GREEN};font-weight:800;font-size:14px`
        : `font-weight:600;font-size:12px`;
      return `<td style="padding:8px 12px;border:1px solid #e0e0e0;text-align:center;${hlStyle}">${esc(f.value)}</td>`;
    }).join('');
    return `<table style="width:100%;border-collapse:collapse;margin-bottom:4px"><tr>${hdr}</tr><tr>${vals}</tr></table>`;
  }

  return row(fields1) + row(fields2);
}

// ── Compute totals helper (handles salaried yr1 vs yr2+) ─────────────────────
function computeTotal(ins, optKey, years, isSalaried) {
  const prem = safeNum(ins[`premium_${optKey}`]);
  if (prem <= 0) return { annual: 0, yr1: 0, yr2: 0, total: 0 };

  if (isSalaried) {
    const yr1Prem = safeNum(ins[`premium_${optKey}_yr1`]);
    const actualYr1 = yr1Prem > 0 ? yr1Prem : prem;
    const yr2Prem = prem; // 2nd year onwards premium
    const total = years <= 1 ? actualYr1 : actualYr1 + yr2Prem * (years - 1);
    return { annual: yr2Prem, yr1: actualYr1, yr2: yr2Prem, total, hasDiscount: yr1Prem > 0 && yr1Prem !== prem };
  } else {
    return { annual: prem, yr1: prem, yr2: prem, total: prem * years, hasDiscount: false };
  }
}

// ── Premium comparison table ─────────────────────────────────────────────────
function buildPremiumTable(data) {
  const insurers = data.insurers || [];
  const labelA = data.option_a || 'Regular';
  const labelB = data.option_b || '10 Pay';
  const yrsA   = data.years_a  || 0;
  const yrsB   = data.years_b  || 0;
  const isSalaried = data.employment_type === 'salaried';

  // Pre-compute all totals to find cheapest
  const computed = insurers.map(ins => {
    const compA = computeTotal(ins, 'a', yrsA, isSalaried);
    const compB = computeTotal(ins, 'b', yrsB, isSalaried);
    return { ins, compA, compB };
  }).filter(c => c.compA.annual > 0 || c.compB.annual > 0);

  // Find cheapest annual for each option
  const validA = computed.filter(c => c.compA.annual > 0);
  const validB = computed.filter(c => c.compB.annual > 0);
  const cheapestA_annual = validA.length > 0 ? Math.min(...validA.map(c => c.compA.yr2)) : 0;
  const cheapestB_annual = validB.length > 0 ? Math.min(...validB.map(c => c.compB.yr2)) : 0;
  const cheapestA_total  = validA.length > 0 ? Math.min(...validA.map(c => c.compA.total)) : 0;
  const cheapestB_total  = validB.length > 0 ? Math.min(...validB.map(c => c.compB.total)) : 0;

  // Column headers
  let cols;
  if (isSalaried) {
    cols = [
      '', 'Company',
      `${labelA}<br>1st Yr`, `${labelA}<br>2nd Yr+`,
      `${labelB}<br>1st Yr`, `${labelB}<br>2nd Yr+`,
      `Total ${labelA}<br>(${yrsA} yrs)`, `Total ${labelB}<br>(${yrsB} yrs)`,
      `Savings with<br>${labelB}`, '%'
    ];
  } else {
    cols = [
      '', 'Company', `${labelA}<br>(Annual)`, `${labelB}<br>(Annual)`,
      `Total ${labelA}<br>(${yrsA} yrs)`, `Total ${labelB}<br>(${yrsB} yrs)`,
      `Savings with<br>${labelB}`, '%'
    ];
  }

  let headerHTML = cols.map((c, i) => {
    let w;
    if (isSalaried) {
      // 10 cols: logo(8%) + company(14%) + 8 data cols(9.75% each = 78%) = 100%
      w = i === 0 ? '8%' : i === 1 ? '14%' : '9.75%';
    } else {
      w = i === 0 ? '8%' : i === 1 ? '18%' : '10.5%';
    }
    return `<th style="background:${BRAND_GREEN};color:white;font-size:10px;font-weight:700;padding:10px 4px;border:1px solid rgba(255,255,255,0.3);text-align:center;width:${w};vertical-align:middle">${c}</th>`;
  }).join('');

  let rowsHTML = '';
  let hasRows = false;

  computed.forEach(({ ins, compA, compB }, idx) => {
    if (compA.annual <= 0 && compB.annual <= 0) return;
    hasRows = true;

    const { brand, sub } = parseName(ins.name || '');
    const logoB64 = getTermLogoB64(ins.name);

    const totA = compA.total;
    const totB = compB.total;
    const delta = totA - totB;
    const pct = totA > 0 ? (Math.abs(delta) / totA * 100).toFixed(1) : '0.0';

    const bg = idx % 2 === 0 ? '#f8fef8' : 'white';

    const logoCell = logoB64
      ? `<img src="${logoB64}" style="height:32px;max-width:60px;object-fit:contain;display:block;margin:0 auto" />`
      : `<div style="width:36px;height:36px;border-radius:50%;background:${INSURER_COLORS[idx % INSURER_COLORS.length]};display:flex;align-items:center;justify-content:center;color:white;font-weight:700;font-size:15px;margin:0 auto">${esc(brand.charAt(0))}</div>`;

    // Highlight cheapest annual premium cells
    const isChpA = compA.yr2 > 0 && compA.yr2 === cheapestA_annual;
    const isChpB = compB.yr2 > 0 && compB.yr2 === cheapestB_annual;
    const isChpTotA = compA.total > 0 && compA.total === cheapestA_total;
    const isChpTotB = compB.total > 0 && compB.total === cheapestB_total;

    const greenBg = `background:${HIGHLIGHT_GREEN};`;
    const greenTick = '✅ ';

    function premCell(val, isCheapest) {
      const style = `padding:8px 4px;border:1px solid #e0e0e0;text-align:center;font-weight:700;font-size:12px;${isCheapest ? greenBg + `color:${BRAND_GREEN}` : ''}`;
      return `<td style="${style}">${isCheapest ? greenTick : ''}${money(val)}</td>`;
    }
    function totalCell(val, isCheapest) {
      const style = `padding:8px 4px;border:1px solid #e0e0e0;text-align:center;font-size:11px;${isCheapest ? greenBg + `color:${BRAND_GREEN};font-weight:700` : ''}`;
      return `<td style="${style}">${isCheapest ? greenTick : ''}${money(val)}</td>`;
    }

    let savingsCell, pctCell;
    if (delta > 0) {
      savingsCell = `<td style="padding:8px 4px;border:1px solid #e0e0e0;text-align:center;font-weight:700;font-size:11px;background:${GREEN_OK};color:#1F4E27">✅ ${money(delta)}</td>`;
      pctCell = `<td style="padding:8px 4px;border:1px solid #e0e0e0;text-align:center;font-weight:700;font-size:11px;background:${GREEN_OK};color:#1F4E27">${pct}%</td>`;
    } else if (delta < 0) {
      savingsCell = `<td style="padding:8px 4px;border:1px solid #e0e0e0;text-align:center;font-weight:700;font-size:11px;background:${AMBER};color:#92400E">⚠️ +${money(Math.abs(delta))}</td>`;
      pctCell = `<td style="padding:8px 4px;border:1px solid #e0e0e0;text-align:center;font-weight:700;font-size:11px;background:${AMBER};color:#92400E">${pct}%</td>`;
    } else {
      savingsCell = `<td style="padding:8px 4px;border:1px solid #e0e0e0;text-align:center;font-size:11px">—</td>`;
      pctCell = `<td style="padding:8px 4px;border:1px solid #e0e0e0;text-align:center;font-size:11px">0%</td>`;
    }

    if (isSalaried) {
      rowsHTML += `<tr style="background:${bg}">
        <td style="padding:6px;border:1px solid #e0e0e0;text-align:center">${logoCell}</td>
        <td style="padding:8px 6px;border:1px solid #e0e0e0;text-align:left;font-size:11px"><b>${esc(brand)}</b>${sub ? `<br><span style="color:#888;font-size:10px">${esc(sub)}</span>` : ''}</td>
        ${premCell(compA.yr1, false)}
        ${premCell(compA.yr2, isChpA)}
        ${premCell(compB.yr1, false)}
        ${premCell(compB.yr2, isChpB)}
        ${totalCell(totA, isChpTotA)}
        ${totalCell(totB, isChpTotB)}
        ${savingsCell}
        ${pctCell}
      </tr>`;
    } else {
      rowsHTML += `<tr style="background:${bg}">
        <td style="padding:6px;border:1px solid #e0e0e0;text-align:center">${logoCell}</td>
        <td style="padding:8px 8px;border:1px solid #e0e0e0;text-align:left;font-size:11px"><b>${esc(brand)}</b>${sub ? `<br><span style="color:#888;font-size:10px">${esc(sub)}</span>` : ''}</td>
        ${premCell(compA.annual, isChpA)}
        ${premCell(compB.annual, isChpB)}
        ${totalCell(totA, isChpTotA)}
        ${totalCell(totB, isChpTotB)}
        ${savingsCell}
        ${pctCell}
      </tr>`;
    }
  });

  if (!hasRows) {
    return `<div style="padding:20px;text-align:center;color:#999;font-size:12px">No rows with premium data.</div>`;
  }

  // Salaried note
  let salNote = '';
  if (isSalaried) {
    salNote = `<div style="margin-top:6px;padding:6px 12px;background:#EBF5FB;border-left:3px solid #2196F3;border-radius:0 6px 6px 0;font-size:10px;color:#1565C0">
      💼 <b>Salaried Discount Applied:</b> 1st year premium shown is the discounted rate. 2nd year onwards is the standard premium. Total = 1st Yr + (2nd Yr × remaining years).
    </div>`;
  }

  return `
  <table style="width:100%;border-collapse:collapse;font-size:11px;table-layout:fixed">
    <thead><tr>${headerHTML}</tr></thead>
    <tbody>${rowsHTML}</tbody>
  </table>${salNote}`;
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

// ── Header (NO Incremint logo — quote is for agent's customer) ───────────────
function buildHeader(data) {
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
      <div style="font-size:10px;color:#666">
        <b>Quote ID: ${esc(data.quote_id || '')}</b> &nbsp;|&nbsp; ${esc(data.date || '')}
      </div>
    </div>
  </div>
  <div style="height:2px;background:${BRAND_GREEN};border-radius:1px;margin-bottom:10px"></div>`;
}

// ── SA + Cover Highlight Banner ──────────────────────────────────────────────
function buildHighlightBanner(data) {
  const sa = data.sum_assured || '';
  const cover = data.cover_till_age || '';
  const term = data.policy_term || '';
  const empLabel = data.employment_type === 'salaried' ? '💼 Salaried' : '🏢 Self-Employed';
  const smokerLabel = data.smoker_status === 'smoker' ? '🚬 Smoker' : '✅ Non-Smoker';

  return `
  <div style="display:flex;gap:10px;margin:8px 0 12px">
    <div style="flex:1;background:#E8F5E9;border:2px solid ${BRAND_GREEN};border-radius:10px;padding:10px 14px;text-align:center">
      <div style="font-size:9px;font-weight:700;color:${BRAND_GREEN};text-transform:uppercase;letter-spacing:0.5px">Sum Assured</div>
      <div style="font-size:22px;font-weight:900;color:${BRAND_GREEN};margin-top:2px">${esc(sa)}</div>
    </div>
    <div style="flex:1;background:#E8F5E9;border:2px solid ${BRAND_GREEN};border-radius:10px;padding:10px 14px;text-align:center">
      <div style="font-size:9px;font-weight:700;color:${BRAND_GREEN};text-transform:uppercase;letter-spacing:0.5px">Cover Till Age</div>
      <div style="font-size:22px;font-weight:900;color:${BRAND_GREEN};margin-top:2px">${esc(cover)}</div>
    </div>
    <div style="flex:0.7;background:#f5f5f5;border:1px solid #e0e0e0;border-radius:10px;padding:10px 14px;text-align:center">
      <div style="font-size:9px;font-weight:700;color:#666;text-transform:uppercase;letter-spacing:0.5px">Policy Term</div>
      <div style="font-size:18px;font-weight:800;color:#333;margin-top:2px">${esc(term)} yrs</div>
    </div>
    <div style="flex:0.7;background:#f5f5f5;border:1px solid #e0e0e0;border-radius:10px;padding:10px 8px;text-align:center">
      <div style="font-size:10px;font-weight:600;color:#555;margin-bottom:4px">${empLabel}</div>
      <div style="font-size:10px;font-weight:600;color:#555">${smokerLabel}</div>
    </div>
  </div>`;
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

// ── Advisory note (recommends cheapest) ──────────────────────────────────────
function buildAdvisoryNote(data) {
  const labelA = data.option_a || 'Regular';
  const labelB = data.option_b || '10 Pay';
  const insurers = data.insurers || [];
  const yrsA = data.years_a || 0;
  const yrsB = data.years_b || 0;
  const isSalaried = data.employment_type === 'salaried';

  // Find cheapest insurer for Option A and Option B
  let cheapestA = null, cheapestB = null;
  let minTotA = Infinity, minTotB = Infinity;
  let cheapestOverall = null, minOverall = Infinity;

  for (const ins of insurers) {
    const compA = computeTotal(ins, 'a', yrsA, isSalaried);
    const compB = computeTotal(ins, 'b', yrsB, isSalaried);
    const { brand } = parseName(ins.name || '');

    if (compA.total > 0 && compA.total < minTotA) {
      minTotA = compA.total;
      cheapestA = { brand, total: compA.total, annual: compA.yr2, option: labelA };
    }
    if (compB.total > 0 && compB.total < minTotB) {
      minTotB = compB.total;
      cheapestB = { brand, total: compB.total, annual: compB.yr2, option: labelB };
    }
    // Overall cheapest
    if (compA.total > 0 && compA.total < minOverall) { minOverall = compA.total; cheapestOverall = { brand, total: compA.total, option: labelA, annual: compA.yr2 }; }
    if (compB.total > 0 && compB.total < minOverall) { minOverall = compB.total; cheapestOverall = { brand, total: compB.total, option: labelB, annual: compB.yr2 }; }
  }

  let lines = [];

  // Main recommendation
  if (cheapestOverall) {
    lines.push(`<div style="font-size:13px;margin-bottom:8px;color:#1a1a1a"><b>👉 Recommended:</b> <span style="color:${BRAND_GREEN};font-weight:800">${esc(cheapestOverall.brand)}</span> with <b>${esc(cheapestOverall.option)}</b> pay — lowest total outgo of <b>${money(cheapestOverall.total)}</b> (${money(cheapestOverall.annual)}/yr)</div>`);
  }

  // Cheapest per option
  if (cheapestA) {
    lines.push(`<div style="font-size:11px;color:#333">• <b>Cheapest ${esc(labelA)}:</b> ${esc(cheapestA.brand)} at ${money(cheapestA.annual)}/yr → Total: ${money(cheapestA.total)}</div>`);
  }
  if (cheapestB) {
    lines.push(`<div style="font-size:11px;color:#333">• <b>Cheapest ${esc(labelB)}:</b> ${esc(cheapestB.brand)} at ${money(cheapestB.annual)}/yr → Total: ${money(cheapestB.total)}</div>`);
  }

  // Savings comparison for cheapest insurer (A vs B)
  if (cheapestA && cheapestB) {
    const diff = cheapestA.total - cheapestB.total;
    if (diff > 0) {
      lines.push(`<div style="font-size:11px;color:${BRAND_GREEN};margin-top:4px">💰 Switching to <b>${esc(labelB)}</b> for ${esc(cheapestB.brand)} saves <b>${money(diff)}</b> over the policy term.</div>`);
    } else if (diff < 0) {
      lines.push(`<div style="font-size:11px;color:${BRAND_GREEN};margin-top:4px">💰 <b>${esc(labelA)}</b> for ${esc(cheapestA.brand)} is cheaper by <b>${money(Math.abs(diff))}</b> overall.</div>`);
    }
  }

  if (isSalaried) {
    lines.push(`<div style="font-size:10px;color:#1565C0;margin-top:6px">💼 <b>Note:</b> Totals account for salaried 1st year discount. Comparison is based on 2nd year onwards premium.</div>`);
  }

  lines.push(`<div style="font-size:10px;color:#888;margin-top:6px">• Pay years: Regular → Policy Term; 5/10/15 Pay → 5/10/15 years; Pay till 60 → (60 − Age)</div>`);

  const noteText = data.advisor_note || '';

  return `
  <div style="margin-top:12px;padding:14px 18px;border-left:5px solid ${BRAND_GREEN};border-radius:0 8px 8px 0;background:white;box-shadow:0 2px 8px rgba(31,78,39,0.12);font-size:12px">
    <div style="font-weight:700;color:${BRAND_GREEN};font-size:12px;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:8px">💡 ADVISORY NOTE</div>
    ${lines.join('\n')}
    ${noteText ? `<div style="margin-top:10px;color:#333;font-style:italic;border-top:1px solid #e0e0e0;padding-top:8px">"${esc(noteText)}"</div>` : ''}
  </div>`;
}

// ── Main export ──────────────────────────────────────────────────────────────
function buildTermQuoteHTML(data) {
  const insurers          = data.insurers || [];
  const featuresOverride  = data.features_override  || null;
  const highlightsOverride= data.highlights_override || null;

  const headerHTML     = buildHeader(data);
  const footerHTML     = buildFooter(data);
  const highlightBanner= buildHighlightBanner(data);
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
  @page { size: A4 landscape; margin: 1.5cm 2cm; }
  @media print { @page { size: A4 landscape; margin: 1.5cm 2cm; } body { margin: 0; } .page { padding: 0; box-shadow: none; } }
  * { -webkit-print-color-adjust: exact; print-color-adjust: exact; box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Arial, sans-serif; font-size: 12px; color: #1a1a1a; background: #f5f5f5; overflow-x: hidden; }
  .page { max-width: 1100px; margin: 20px auto; padding: 32px 40px; background: white; border-radius: 8px; box-shadow: 0 2px 12px rgba(0,0,0,0.08); overflow: hidden; }
  .page-break { page-break-after: always; page-break-inside: avoid; margin-bottom: 20px; }
  td, th { word-wrap: break-word; overflow-wrap: break-word; }
</style>
</head>
<body>

<!-- ══════════════ PAGE 1 ══════════════ -->
<div class="page page-break">
  ${headerHTML}
  ${highlightBanner}
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
