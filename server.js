const express = require('express');
const cors = require('cors');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');
const crypto = require('crypto');

const app = express();
app.use(cors());
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ limit: '10mb', extended: true }));
app.use(express.static(path.join(__dirname, 'public')));

// ─── AUTH ────────────────────────────────────────────────────────────────────

const AUTH_USER = 'adminscj01';
const AUTH_PASS = 'adminscj02';
let authTokens = new Set();

app.post('/api/login', (req, res) => {
  const { username, password } = req.body;
  if (username === AUTH_USER && password === AUTH_PASS) {
    const token = crypto.randomBytes(32).toString('hex');
    authTokens.add(token);
    return res.json({ ok: true, token });
  }
  // ponytail: fixed credentials, single user — upgrade to DB if multi-user needed
  return res.status(401).json({ ok: false, message: 'Invalid credentials' });
});

app.get('/api/check-auth', (req, res) => {
  const token = req.headers['x-auth-token'] || req.query.token;
  if (token && authTokens.has(token)) return res.json({ ok: true });
  res.json({ ok: false });
});

function requireAuth(req, res, next) {
  const token = req.headers['x-auth-token'] || req.query.token;
  if (token && authTokens.has(token)) return next();
  res.status(401).json({ ok: false, message: 'Unauthorized' });
}

// ─── CONFIG ───────────────────────────────────────────────────────────────────

const EXCEL_PATH = path.join(__dirname, 'data', 'Bookings - By Carrier - 2026.xlsx');
const PORT = process.env.PORT || 2099; // Render provides PORT env

// ─── KPI MAPPING ──────────────────────────────────────────────────────────────

const PREFIX_TO_KPI = {
  AMP: 'Air Freight', AXP: 'Air Freight',
  CCE: 'Clearance Export', CCI: 'Clearance Import',
  DOF: 'Ocean Domestic', DRF: 'Reefer Domestic',
  EXP: 'Ocean Export', IMP: 'Ocean Import',
  DOM: 'Full Truck Load',
  LTL: 'Less-than Truck Load Domestic',
  LCL: 'Less-than Container Load Export-Import',
  LCLIMP: 'Less-than Container Load Export-Import',
  LEXP: 'Less-than Container Load Export-Import',
  LMP: 'Less-than Container Load Export-Import',
  LXP: 'Less-than Container Load Export-Import',
  PRO: 'Project', GEN: 'General'
};

const PREFIX_TO_UOM = {
  AMP: 'KG', AXP: 'KG',
  DOF: 'TEU', DRF: 'TEU', EXP: 'TEU', IMP: 'TEU',
  DOM: 'TRIP',
  LCL: 'CBM', LCLIMP: 'CBM', LEXP: 'CBM', LMP: 'CBM', LXP: 'CBM',
  LTL: 'KG',
  CCE: 'JOB', CCI: 'JOB', PRO: 'JOB', GEN: 'JOB'
};

// ─── TEAM MAPPING ─────────────────────────────────────────────────────────────

const SALES_TO_TEAM = {
  'Dimas Destrianto': 'ALPHA', 'Dimas Yudi': 'ALPHA',
  'Karmila Damayanti': 'ALPHA', 'Laudy': 'ALPHA',
  'Selvy Stevani': 'ALPHA', 'Silvy': 'ALPHA', 'Xylon': 'ALPHA',
  'Jimmy Rantung': 'SENIOR', 'Kezia': 'SENIOR',
  'Yanes': 'SENIOR', 'Vipul Malik': 'SENIOR',
  'Rike': 'CHARLIE', 'Silvia Anggraeini': 'CHARLIE',
  'Semarang Office': 'CHARLIE',
  'Rizka Asna': 'DELTA', 'Sony Agustyawan': 'DELTA',
  'Surabaya Office': 'DELTA',
  'Bientang': 'OPS', 'Ceasarsyah': 'OPS', 'Dito': 'OPS',
  'Ilal Albany': 'OPS', 'Operation': 'OPS', 'Operational': 'OPS',
  'Ario Teguh': 'MANAGEMENT', 'Suchit': 'MANAGEMENT', 'Endang': 'MANAGEMENT',
  'Management': 'HERO'
};

const TEAM_TO_BRANCH = {
  ALPHA: 'Jakarta', SENIOR: 'Jakarta',
  CHARLIE: 'Semarang', DELTA: 'Surabaya',
  OPS: 'OPS', HERO: 'HERO',
  MANAGEMENT: 'Management'
};

// ─── MONTHLY TARGETS (USD) ────────────────────────────────────────────────────

const TEAM_MONTHLY_TARGET = {
  ALPHA: 62500,
  CHARLIE: 62500,
  DELTA: 62500,
  SENIOR: 0,
  OPS: 0,
  HERO: 0,
  MANAGEMENT: 0
};

const MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December'];

// ─── CONTRACT HOLDER CLEANING ────────────────────────────────────────────────

function cleanContractHolder(val) {
  if (!val) return '';
  let s = val.toString().trim();
  s = s.toUpperCase();
  s = s.replace(/\s+/g, ' ');
  s = s.replace(/\bPT\.\s*/g, 'PT ');
  s = s.replace(/^[^A-Z0-9]+/, '');
  s = s.replace(/[^A-Z0-9]+$/, '');
  s = s.replace(/\s+/g, ' ');
  return s.trim();
}

// ─── PREFIX EXTRACTION ────────────────────────────────────────────────────────

function extractPrefix(name) {
  if (!name) return '';
  let clean = name.toString().trim();
  clean = clean.replace(/^\d+\.\s*/, '');
  clean = clean.replace(/^[^A-Za-z]+/, '');
  const m = clean.match(/^([A-Za-z]+)/);
  return m ? m[1].toUpperCase() : '';
}

function parseNum(val) {
  if (val === undefined || val === null || val === '') return 0;
  if (typeof val === 'number') return val;
  const str = String(val).replace(/[,\s]/g, '');
  const n = parseFloat(str);
  return isNaN(n) ? 0 : n;
}

// ─── VOLUME CALCULATION ───────────────────────────────────────────────────────

function calcVolume(uom, row) {
  const pieces = parseNum(row[11]);
  const size = (row[12]||'').toString().trim();
  const teu = parseNum(row[13]);
  const weightLb = parseNum(row[29]);
  const volumeFt = parseNum(row[30]);

  switch (uom) {
    case 'TEU': {
      if (teu > 0) return teu;
      if (size === '20DV' || size === '20RF' || size === '20RH') return pieces;
      if (size === '40DV' || size === '40HC' || size === '40RH') return pieces * 2;
      return pieces;
    }
    case 'KG':
      return (parseNum(row[32]) || parseNum(row[29])) / 2.20462;
    case 'CBM':
      return volumeFt * 0.0283168;
    case 'TRIP':
      return pieces;
    case 'JOB':
      return 1;
    default:
      return 0;
  }
}

// ─── DATA LOADING & PROCESSING ────────────────────────────────────────────────

function loadAndProcessData() {
  const wb = XLSX.readFile(EXCEL_PATH);
  const ws = wb.Sheets[wb.SheetNames[0]];
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

  const header = raw[6];
  const records = [];

  for (let i = 7; i < raw.length; i++) {
    const r = raw[i];
    const name = (r[27]||'').toString().trim();
    if (!name || name === 'Total' || name.startsWith('Grand')) continue;

    const prefix = extractPrefix(name);
    const kpi = PREFIX_TO_KPI[prefix] || '';
    const uom = PREFIX_TO_UOM[prefix] || '';

    if (!kpi) continue;

    const sales = (r[23]||'').toString().trim();
    const team = SALES_TO_TEAM[sales] || '';
    const branch = TEAM_TO_BRANCH[team] || '';

    const contractHolderRaw = (r[16]||'').toString().trim();
    const contractHolderClean = cleanContractHolder(r[16]);

    const dateStr = (r[8]||'').toString().trim();
    let monthIdx = -1;
    if (dateStr) {
      const parts = dateStr.split('/');
      if (parts.length >= 3) {
        const m = parseInt(parts[0], 10);
        const y = parseInt(parts[2], 10);
        if (y === 2026 && m >= 1 && m <= 12) {
          monthIdx = m - 1;
        }
      }
    }

    const gp = parseNum(r[44]);
    const volume = calcVolume(uom, r);

    const carrier = (r[0]||'').toString().trim();
    const origin = (r[9]||'').toString().trim();
    const region = (r[20]||'').toString().trim();

    records.push({
      name, prefix, kpi, uom,
      pieces: parseNum(r[11]),
      size: (r[12]||'').toString().trim(),
      teu: parseNum(r[13]),
      weight: parseNum(r[29]),
      volumeFt: parseNum(r[30]),
      sales, team, branch,
      contractHolder: contractHolderRaw,
      contractHolderClean,
      carrier, origin, region,
      gp, volume, dateStr, monthIdx
    });
  }

  return records;
}

// ─── API: Data ────────────────────────────────────────────────────────────────

function buildApiData(records) {
  const activeMonths = [...new Set(records.filter(r => r.monthIdx >= 0).map(r => r.monthIdx))].sort((a,b) => a-b);

  const kpiOrder = [
    'Gross Profit',
    'Ocean Export','Ocean Import','Ocean Domestic',
    'Full Truck Load','Reefer Domestic','Air Freight',
    'Less-than Container Load Export-Import',
    'Less-than Truck Load Domestic',
    'Clearance Import','Clearance Export','Project','General'
  ];

  const teams = ['ALPHA','SENIOR','CHARLIE','DELTA','OPS','HERO','MANAGEMENT'];
  const kpiMatrix = [];

  for (const kpi of kpiOrder) {
    const row = { kpi };
    if (kpi === 'Gross Profit') {
      for (const team of teams) {
        const actual = records
          .filter(r => r.team === team && r.monthIdx >= 0)
          .reduce((s, r) => s + r.gp, 0);
        const target = TEAM_MONTHLY_TARGET[team] || 0;
        row[`${team}_target`] = target;
        row[`${team}_actual`] = Math.round(actual * 100) / 100;
        row[`${team}_ach`] = target > 0 ? Math.round((actual / target) * 10000) / 100 : 0;
      }
    } else {
      for (const team of teams) {
        const teamRecs = records.filter(r => r.kpi === kpi && r.team === team && r.monthIdx >= 0);
        const actual = teamRecs.reduce((s, r) => s + r.volume, 0);
        const target = 0;
        row[`${team}_target`] = target;
        row[`${team}_actual`] = Math.round(actual * 100) / 100;
        row[`${team}_ach`] = target > 0 ? Math.round((actual / target) * 10000) / 100 : 0;
      }
    }
    kpiMatrix.push(row);
  }

  const gpTrend = activeMonths.map(m => {
    const actual = records
      .filter(r => r.monthIdx === m)
      .reduce((s, r) => s + r.gp, 0);
    const totalTarget = Object.values(TEAM_MONTHLY_TARGET).reduce((a,b) => a+b, 0);
    return {
      month: MONTHS[m],
      monthIdx: m,
      target: totalTarget,
      actual: Math.round(actual * 100) / 100,
      ach: totalTarget > 0 ? Math.round((actual / totalTarget) * 10000) / 100 : 0
    };
  });

  const gpByTeam = teams.map(team => {
    const actual = records
      .filter(r => r.team === team && r.monthIdx >= 0)
      .reduce((s, r) => s + r.gp, 0);
    const target = TEAM_MONTHLY_TARGET[team] || 0;
    return {
      team,
      target,
      actual: Math.round(actual * 100) / 100,
      ach: target > 0 ? Math.round((actual / target) * 10000) / 100 : 0
    };
  }).filter(t => t.actual > 0 || t.target > 0);

  const kpiVolumeByTeam = [];
  for (const kpi of kpiOrder) {
    if (kpi === 'Gross Profit') continue;
    const row = { kpi };
    for (const team of teams) {
      const vol = records
        .filter(r => r.kpi === kpi && r.team === team && r.monthIdx >= 0)
        .reduce((s, r) => s + r.volume, 0);
      row[team] = Math.round(vol * 100) / 100;
    }
    kpiVolumeByTeam.push(row);
  }

  const salesGp = {};
  for (const r of records) {
    if (r.monthIdx < 0) continue;
    if (!salesGp[r.sales]) salesGp[r.sales] = { sales: r.sales, team: r.team, gp: 0 };
    salesGp[r.sales].gp += r.gp;
  }
  const topSales = Object.values(salesGp)
    .map(s => ({ ...s, gp: Math.round(s.gp * 100) / 100 }))
    .sort((a, b) => b.gp - a.gp)
    .slice(0, 25);

  const branches = ['Jakarta','Semarang','Surabaya','OPS','HERO','Management'];
  const branchPerf = branches.map(b => {
    const actual = records
      .filter(r => r.branch === b && r.monthIdx >= 0)
      .reduce((s, r) => s + r.gp, 0);
    return {
      branch: b,
      actual: Math.round(actual * 100) / 100
    };
  }).filter(b => b.actual > 0);

  const teamPerf = gpByTeam.map(t => ({
    ...t,
    branch: TEAM_TO_BRANCH[t.team] || '',
    recordCount: records.filter(r => r.team === t.team && r.monthIdx >= 0).length
  }));

  const activeRecs = records.filter(r => r.monthIdx >= 0);
  const custMap = {};
  for (const r of activeRecs) {
    const c = r.contractHolderClean || '(BLANK)';
    if (!custMap[c]) custMap[c] = { customer: c, gp: 0, volume: 0, shipmentCount: 0, uomCounts: {} };
    custMap[c].gp += r.gp;
    custMap[c].volume += r.volume;
    custMap[c].shipmentCount++;
    if (r.uom) custMap[c].uomCounts[r.uom] = (custMap[c].uomCounts[r.uom] || 0) + 1;
  }
  const custArr = Object.values(custMap).map(c => {
    let topUom = '', topCount = 0;
    for (const [uom, cnt] of Object.entries(c.uomCounts)) {
      if (cnt > topCount) { topCount = cnt; topUom = uom; }
    }
    return {
      customer: c.customer,
      gp: Math.round(c.gp * 100) / 100,
      volume: Math.round(c.volume * 100) / 100,
      shipmentCount: c.shipmentCount,
      uom: topUom
    };
  });

  const uomMap = {};
  for (const r of activeRecs) {
    if (!r.uom) continue;
    if (!uomMap[r.uom]) uomMap[r.uom] = { uom: r.uom, totalVolume: 0, shipmentCount: 0 };
    uomMap[r.uom].totalVolume += r.volume;
    uomMap[r.uom].shipmentCount++;
  }
  const volumeByUom = Object.values(uomMap)
    .map(u => ({ uom: u.uom, totalVolume: Math.round(u.totalVolume * 100) / 100, shipmentCount: u.shipmentCount }))
    .sort((a, b) => b.totalVolume - a.totalVolume);

  const totalGP = activeRecs.reduce((s, r) => s + r.gp, 0);
  const topCustomersGP = [...custArr].sort((a, b) => b.gp - a.gp).slice(0, 50);
  const topCustomersVolume = [...custArr].sort((a, b) => b.volume - a.volume).slice(0, 50);
  const customerContribution = topCustomersGP.slice(0, 10).map(c => ({
    customer: c.customer,
    gp: c.gp,
    contributionPct: totalGP > 0 ? Math.round((c.gp / totalGP) * 10000) / 100 : 0
  }));

  const contractHolders = [...new Set(activeRecs.map(r => r.contractHolderClean).filter(Boolean))].sort();
  const prefixes = [...new Set(activeRecs.map(r => r.prefix).filter(Boolean))].sort();
  const carriers = [...new Set(activeRecs.map(r => r.carrier).filter(Boolean))].sort();
  const origins = [...new Set(activeRecs.map(r => r.origin).filter(Boolean))].sort();
  const regions = [...new Set(activeRecs.map(r => r.region).filter(Boolean))].sort();

  return {
    activeMonths: activeMonths.map(m => ({ idx: m, name: MONTHS[m] })),
    kpiMatrix, gpTrend, gpByTeam, kpiVolumeByTeam,
    topSales, branchPerf, teamPerf, teams,
    contractHolders, topCustomersGP, topCustomersVolume,
    customerContribution, volumeByUom,
    filterOptions: {
      months: activeMonths.map(m => ({ idx: m, name: MONTHS[m] })),
      contractHolders,
      teams: [...new Set(activeRecs.map(r => r.team).filter(Boolean))].sort(),
      branches: [...new Set(activeRecs.map(r => r.branch).filter(Boolean))].sort(),
      prefixes, carriers, origins, regions,
      sales: [...new Set(activeRecs.map(r => r.sales).filter(Boolean))].sort()
    },
    summary: {
      totalGP: Math.round(totalGP * 100) / 100,
      totalRecords: activeRecs.length,
      totalTarget: Object.values(TEAM_MONTHLY_TARGET).reduce((a, b) => a + b, 0) * activeMonths.length
    }
  };
}

// ─── CACHED DATA ──────────────────────────────────────────────────────────────

let cachedRecords = null;
function getRecords() {
  if (!cachedRecords) {
    cachedRecords = loadAndProcessData();
  }
  return cachedRecords;
}

// ─── COMMERCIAL KPI TARGETS ──────────────────────────────────────────────────

const COMMERCIAL_KPIS = [
  { code: 'GP',   label: 'Gross Profit',                    uom: 'USD', prefixes: [],                   targets: { ALPHA: 62500,CHARLIE: 62500,DELTA: 62500} },
  { code: 'EXP',  label: 'Ocean Export',                    uom: 'TEU', prefixes: ['EXP'],             targets: { ALPHA: 275, CHARLIE: 275, DELTA: 275 } },
  { code: 'IMP',  label: 'Ocean Import',                    uom: 'TEU', prefixes: ['IMP'],             targets: { ALPHA: 20,  CHARLIE: 20,  DELTA: 20  } },
  { code: 'DOF',  label: 'Ocean Domestic',                  uom: 'TEU', prefixes: ['DOF'],             targets: { ALPHA: 175, CHARLIE: 125, DELTA: 175 } },
  { code: 'DOM',  label: 'Full Truck Load',                 uom: 'TRIP',prefixes: ['DOM'],              targets: { ALPHA: 260,  CHARLIE: 500,  DELTA: 260  } },
  { code: 'DRF',  label: 'Reefer Domestic',                 uom: 'TEU', prefixes: ['DRF'],             targets: { ALPHA: 25,  CHARLIE: 10,  DELTA: 20  } },
  { code: 'AIR',  label: 'Air Freight',                     uom: 'KG',  prefixes: ['AMP','AXP'],       targets: { ALPHA: 5000, CHARLIE: 2000, DELTA: 4000 } },
  { code: 'LCL',  label: 'Less-than Container Load',        uom: 'CBM', prefixes: ['LCL','LCLIMP','LEXP','LMP','LXP'], targets: { ALPHA: 25, CHARLIE: 25, DELTA: 25 } },
  { code: 'LTL',  label: 'Less-than Truck Load',            uom: 'KG',  prefixes: ['LTL'],             targets: { ALPHA: 7500,CHARLIE: 7500,DELTA: 7500} },
  { code: 'LCD',  label: 'Less-than Container Domestic',    uom: 'CBM', prefixes: ['LCD'],             targets: { ALPHA: 0, CHARLIE: 0, DELTA: 0 } },
  { code: 'CCI',  label: 'Clearance Import',                uom: 'JOB', prefixes: ['CCI'],             targets: { ALPHA: 0, CHARLIE: 0, DELTA: 0 } },
  { code: 'CCE',  label: 'Clearance Export',                uom: 'JOB', prefixes: ['CCE'],             targets: { ALPHA: 0, CHARLIE: 0, DELTA: 0 } },
  { code: 'PRO',  label: 'Project',                         uom: 'JOB', prefixes: ['PRO'],             targets: { ALPHA: 0, CHARLIE: 0, DELTA: 0 } },
  { code: 'GEN',  label: 'General',                         uom: 'JOB', prefixes: ['GEN'],             targets: { ALPHA: 0, CHARLIE: 0, DELTA: 0 } }
];

const COMMERCIAL_TEAMS = ['ALPHA','CHARLIE','DELTA'];

// ─── ROUTES ───────────────────────────────────────────────────────────────────

app.get('/api/commercial-kpi', requireAuth, (req, res) => {
  let base = getRecords().filter(r => r.monthIdx >= 0 && COMMERCIAL_TEAMS.includes(r.team));
  function filterBy(recs, field, queryVal) {
    if (queryVal === undefined || queryVal === '') return recs;
    const vals = queryVal.split(',').map(v => v.trim()).filter(Boolean);
    if (vals.length === 0) return recs;
    return recs.filter(r => vals.includes(String(r[field])));
  }
  base = filterBy(base, 'monthIdx', req.query.month);
  base = filterBy(base, 'contractHolderClean', req.query.contractHolder);
  base = filterBy(base, 'team', req.query.team);
  base = filterBy(base, 'branch', req.query.branch);
  base = filterBy(base, 'sales', req.query.sales);
  base = filterBy(base, 'prefix', req.query.prefix);
  base = filterBy(base, 'carrier', req.query.carrier);
  base = filterBy(base, 'origin', req.query.origin);
  base = filterBy(base, 'region', req.query.region);

  const records = base;
  const teams = COMMERCIAL_TEAMS;

  const monthIdxs = [...new Set(records.map(r => r.monthIdx))].sort((a,b) => a-b);
  const activeMonthCount = monthIdxs.length;

  const kpis = COMMERCIAL_KPIS.map(ck => {
    const teamData = {};

    teams.forEach(t => {
      let actual = 0;
      if (ck.code === 'GP') {
        actual = records.filter(r => r.team === t).reduce((s, r) => s + r.gp, 0);
      } else {
        actual = records.filter(r => r.team === t && ck.prefixes.includes(r.prefix)).reduce((s, r) => s + r.volume, 0);
      }
      actual = Math.round(actual * 100) / 100;
      const target = (ck.targets[t] || 0) * activeMonthCount;
      const ach = target > 0 ? Math.round((actual / target) * 10000) / 100 : 0;
      teamData[t] = { target, actual, ach };
    });

    return {
      code: ck.code,
      label: ck.label,
      uom: ck.uom,
      prefixes: ck.prefixes,
      teams: teamData,
    };
  });

  const totals = {};
  teams.forEach(t => {
    const gpKpi = kpis.find(k => k.code === 'GP');
    const gpData = gpKpi ? gpKpi.teams[t] : null;
    totals[t] = {
      actual: gpData ? gpData.actual : 0,
      pct: gpData ? gpData.ach : 0
    };
  });

  const totalGP = teams.reduce((s, t) => s + (totals[t]?.actual || 0), 0);
  const totalTarget = teams.reduce((s, t) => s + (COMMERCIAL_KPIS.find(k => k.code === 'GP')?.targets[t] || 0) * activeMonthCount, 0);
  const totalRecords = records.length;

  res.json({ kpis, teams, totals, months: allMonths(), summary: { totalGP, totalTarget, totalRecords, activeMonthCount } });
});

function allMonths() {
  const m = getRecords().filter(r => r.monthIdx >= 0).map(r => r.monthIdx);
  return [...new Set(m)].sort().map(i => ({ idx: i, name: MONTHS[i] }));
}

app.get('/api/data', requireAuth, (req, res) => {
  const allRecords = getRecords();
  let records = allRecords;

  if (!req.query.team) {
    records = records.filter(r => COMMERCIAL_TEAMS.includes(r.team));
  }

  function filterBy(recs, field, queryVal) {
    if (queryVal === undefined || queryVal === '') return recs;
    const vals = queryVal.split(',').map(v => v.trim()).filter(Boolean);
    if (vals.length === 0) return recs;
    return recs.filter(r => vals.includes(String(r[field])));
  }

  records = filterBy(records, 'monthIdx', req.query.month);
  records = filterBy(records, 'contractHolderClean', req.query.contractHolder);
  records = filterBy(records, 'team', req.query.team);
  records = filterBy(records, 'branch', req.query.branch);
  records = filterBy(records, 'sales', req.query.sales);
  records = filterBy(records, 'prefix', req.query.prefix);
  records = filterBy(records, 'carrier', req.query.carrier);
  records = filterBy(records, 'origin', req.query.origin);
  records = filterBy(records, 'region', req.query.region);

  res.json(buildApiData(records));
});

app.get('/api/kpi-matrix', requireAuth, (req, res) => {
  res.json(buildApiData(getRecords()).kpiMatrix);
});

app.get('/api/gp-trend', requireAuth, (req, res) => {
  res.json(buildApiData(getRecords()).gpTrend);
});

app.get('/api/gp-by-team', requireAuth, (req, res) => {
  res.json(buildApiData(getRecords()).gpByTeam);
});

app.get('/api/kpi-volume-by-team', requireAuth, (req, res) => {
  res.json(buildApiData(getRecords()).kpiVolumeByTeam);
});

app.get('/api/top-sales', requireAuth, (req, res) => {
  res.json(buildApiData(getRecords()).topSales);
});

app.get('/api/branch-performance', requireAuth, (req, res) => {
  res.json(buildApiData(getRecords()).branchPerf);
});

app.get('/api/team-performance', requireAuth, (req, res) => {
  res.json(buildApiData(getRecords()).teamPerf);
});

app.get('/api/refresh', requireAuth, (req, res) => {
  cachedRecords = null;
  const all = loadAndProcessData();
  cachedRecords = all;
  const totalGP = all.filter(r => r.monthIdx >= 0).reduce((s, r) => s + r.gp, 0);
  res.json({ ok: true, records: all.length, totalGP: Math.round(totalGP * 100) / 100 });
});

// ─── UPLOAD: replace Excel file from dashboard ───────────────────────────────
// ponytail: uses base64 JSON to avoid multer/formidable dependency
app.post('/api/upload', requireAuth, (req, res) => {
  const { file: dataUrl } = req.body;
  if (!dataUrl) return res.status(400).json({ ok: false, message: 'No file provided' });

  const base64 = dataUrl.replace(/^data:.+;base64,/, '');
  fs.writeFileSync(EXCEL_PATH, Buffer.from(base64, 'base64'));

  cachedRecords = null;
  try {
    const all = loadAndProcessData();
    cachedRecords = all;
    const totalGP = all.filter(r => r.monthIdx >= 0).reduce((s, r) => s + r.gp, 0);
    res.json({ ok: true, records: all.length, totalGP: Math.round(totalGP * 100) / 100 });
  } catch (e) {
    res.status(400).json({ ok: false, message: e.message });
  }
});

// ─── FILE WATCHER: auto-refresh when Excel changes ──────────────────────────

console.log(`👀 Watching for changes: ${path.basename(EXCEL_PATH)}`);
fs.watchFile(EXCEL_PATH, { interval: 3000 }, (curr, prev) => {
  if (curr.mtimeMs === prev.mtimeMs) return;
  console.log(`🔄 File changed, refreshing data...`);
  cachedRecords = null;
  try {
    const all = loadAndProcessData();
    cachedRecords = all;
    const totalGP = all.filter(r => r.monthIdx >= 0).reduce((s, r) => s + r.gp, 0);
    console.log(`✅ Reloaded ${all.length} records, GP: $${totalGP.toLocaleString(undefined, {minimumFractionDigits:2, maximumFractionDigits:2})}`);
  } catch (err) {
    console.error(`❌ Failed to reload: ${err.message}`);
  }
});

// ─── START ────────────────────────────────────────────────────────────────────

app.listen(PORT, () => {
  console.log(`🚀 SCJ Dashboard Flow running at http://localhost:${PORT}`);
  const all = getRecords();
  const totalGP = all.filter(r => r.monthIdx >= 0).reduce((s, r) => s + r.gp, 0);
  console.log(`📊 Loaded ${all.length} records, total GP: $${totalGP.toLocaleString(undefined, {minimumFractionDigits:2, maximumFractionDigits:2})}`);
});
