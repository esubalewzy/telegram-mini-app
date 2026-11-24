import express from 'express';
import path from 'path';
import { fileURLToPath } from 'url';
import crypto from 'crypto';
import Groq from 'groq-sdk';
import cron from 'node-cron';
import {
  Document, Packer, Paragraph, TextRun, HeadingLevel,
  Table, TableRow, TableCell, WidthType, ShadingType, AlignmentType
} from 'docx';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
app.use(express.json({ limit: '10mb' }));
app.use(express.static(__dirname));

const PORT = process.env.PORT || 3000;
const groq = new Groq({ apiKey: process.env.GROQ_API_KEY || '' });
const BREVO_API_KEY = process.env.BREVO_API_KEY || '';

const temporaryDocuments = new Map();
const DOWNLOAD_EXPIRY = 15 * 60 * 1000;
const sentMaternalDeathAlerts = new Set();

const REGIONAL_EMAILS = ['esubalewzerihun11@gmail.com', 'messiwode46@gmail.com'];
const ZONAL_EMAILS = {
  'kaffa zone': ['tadessebelay1221@gmail.com', 'esubalewzerihun37@gmail.com'],
  'dawro zone': ['ketesfu7@gmail.com', 'akliluasfaw423@gmail.com', 'yohannesante@gmail.com'],
  'west omo zone': ['getahunzewde37@gmail.com', 'esubalewzerihun37@gmail.com'],
  'bench sheko zone': ['Joannagashahun13@gmail.com', 'beyenesham32@gmail.com', 'habtekelbi78@gmail.com', 'ruphaelw@gmail.com'],
  'sheka zone': ['sifanmesfin@gmail.com', 'esubalewzerihun37@gmail.com'],
  'konta zone': ['tsegabebe2015@gmail.com', 'degemukidist365@gmail.com']
};

const cleanupTimer = setInterval(() => {
  const now = Date.now();
  for (const [token, file] of temporaryDocuments.entries()) {
    if (now >= file.expiresAt) temporaryDocuments.delete(token);
  }
}, 60 * 1000);
cleanupTimer.unref?.();

const SHEET_ID = '1d8_5Q2W3aGV68W_DdjrPrFEJnuRSDprrb1b3kngbTpo';
const FINAL_SHEET_CSV_URL = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:csv&sheet=Final`;

const REPORT_FOOTER = `\n\n---\n*This report is prepared by **SWERHB MCH Team**. Note: This is a preliminary report; for more information and interactive analytics, please visit the **RMNCH Dashboard**.*`;

const REGIONAL_TARGETS = [
  { label: 'Contraceptive Acceptance Rate (CAR)', target: 96, patterns: ['contraceptive accept', 'car'] },
  { label: 'Long-Acting Family Planning (LAFP)', target: 48, patterns: ['long acting', 'long-acting', 'lafp'] },
  { label: 'Early ANC', target: 57, patterns: ['early anc'] },
  { label: 'ANC 4', target: 98, patterns: ['anc 4', 'anc-4', 'anc4'] },
  { label: 'ANC 8', target: 66, patterns: ['anc 8', 'anc-8', 'anc8'] },
  { label: 'Skilled Birth Attendance (SBA)', target: 89, patterns: ['skilled birth', 'sba', 'delivery'] },
  { label: 'Hepatitis BD', target: 87, patterns: ['hepatitis', 'hep b', 'hepb'] },
  { label: 'Penta 1', target: 100, patterns: ['penta 1', 'penta-1', 'pentavalent 1', 'penta1'] },
  { label: 'Penta 3', target: 100, patterns: ['penta 3', 'penta-3', 'pentavalent 3', 'penta3'] }
];

function parseCSVLine(text) {
  const result = [];
  let current = '';
  let inQuotes = false;
  for (let i = 0; i < text.length; i++) {
    const char = text[i];
    if (char === '"') {
      if (inQuotes && text[i + 1] === '"') {
        current += '"';
        i++;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (char === ',' && !inQuotes) {
      result.push(current.trim());
      current = '';
    } else {
      current += char;
    }
  }
  result.push(current.trim());
  return result;
}

async function fetchFinalSheetData() {
  try {
    const response = await fetch(FINAL_SHEET_CSV_URL);
    if (!response.ok) throw new Error(`Fetch failed: ${response.status} ${response.statusText}`);
    const csvText = await response.text();
    const lines = csvText.split(/\r?\n/).filter(line => line.trim() !== '');
    if (!lines.length) return [];
    const headers = parseCSVLine(lines[0]);
    return lines.slice(1).map(line => {
      const values = parseCSVLine(line);
      const obj = {};
      headers.forEach((header, index) => { obj[header] = values[index] ?? ''; });
      return obj;
    });
  } catch (err) {
    console.error('Sheet Fetch Error:', err);
    throw err;
  }
}

function getCurrentReportingWeek(rawData) {
  const baseDate = new Date('2026-11-09T00:00:00');
  const baseWeekNumber = 19;
  const currentDate = new Date();
  const diffDays = Math.floor((currentDate - baseDate) / (1000 * 60 * 60 * 24));
  const calculatedWeekNum = baseWeekNumber + Math.max(0, Math.floor(diffDays / 7));
  const targetWeekQuery = `Week ${calculatedWeekNum}`;
  const availableWeeks = [...new Set(rawData.map(r => r.Week?.toString().trim()).filter(Boolean))];
  return availableWeeks.find(w => w.toLowerCase().includes(targetWeekQuery.toLowerCase()))
    || availableWeeks[availableWeeks.length - 1];
}

function findColumnByPatterns(headers, patterns, excludeSubstrings = ['Eligible', 'Target']) {
  return headers.find(h =>
    !excludeSubstrings.some(ex => h.includes(ex)) &&
    patterns.some(p => h.toLowerCase().includes(p.toLowerCase()))
  );
}

async function checkAndSendMaternalDeathAlerts(rawData, targetWeek = null) {
  if (!rawData?.length || !BREVO_API_KEY) return;
  const activeWeek = targetWeek && targetWeek !== 'All Weeks'
    ? targetWeek
    : getCurrentReportingWeek(rawData);

  const headers = Object.keys(rawData[0] || {});
  const maternalDeathCol = headers.find(h => {
    const low = h.toLowerCase();
    return !low.includes('eligible') && !low.includes('target') &&
      (low.includes('maternal death') || h.startsWith('2.8'));
  });

  if (!maternalDeathCol) return;

  const weekRows = rawData.filter(r =>
    r.Week && r.Week.toString().trim().toLowerCase().includes(activeWeek.toLowerCase())
  );

  const newAlerts = [];
  weekRows.forEach(row => {
    const deaths = parseFloat(row[maternalDeathCol]);
    if (!isNaN(deaths) && deaths > 0) {
      const alertKey = `${row.Week}|${row.Zone}|${row.Woreda}|${deaths}`;
      if (!sentMaternalDeathAlerts.has(alertKey)) {
        sentMaternalDeathAlerts.add(alertKey);
        newAlerts.push({
          week: row.Week,
          zone: row.Zone || 'Unknown Zone',
          woreda: row.Woreda || 'Unknown Woreda',
          deaths
        });
      }
    }
  });

  if (!newAlerts.length) return;

  const alertLines = newAlerts
    .map(a => `- Zone: ${a.zone} | Woreda: ${a.woreda} | Week: ${a.week} | Reported Maternal Deaths: ${a.deaths}`)
    .join('\n');

  const alertBody = `Dear Regional Leadership Team,\n\n` +
    `URGENT NOTIFICATION: Maternal death case(s) have been reported in the Weekly RMNCH Surveillance System.\n\n` +
    `Incident Details:\n${alertLines}\n\n` +
    `Please initiate immediate MDSR (Maternal Death Surveillance and Response) verification and follow-up with the respective Zonal Health Department(s).` +
    REPORT_FOOTER;

  try {
    await sendViaBrevo(
      REGIONAL_EMAILS,
      `[URGENT MATERNAL DEATH ALERT] RMNCH Surveillance - ${activeWeek}`,
      alertBody
    );
    console.log(`[Maternal Death Alert] Dispatched ${newAlerts.length} alert(s) to Regional Leadership.`);
  } catch (err) {
    console.error('[Maternal Death Alert] Failed to send email:', err);
  }
}

function buildAnalyticsPayload(rawData, selectedZone, selectedWeek) {
  const isGrandTotal = !selectedZone || selectedZone === 'Grand Total';
  const effectiveWeek = (!selectedWeek || selectedWeek === 'Current Week')
    ? getCurrentReportingWeek(rawData)
    : selectedWeek;

  const regionalWeekData = effectiveWeek === 'All Weeks'
    ? rawData
    : rawData.filter(r => r.Week && r.Week.toLowerCase().includes(effectiveWeek.toLowerCase()));

  const zoneAllData = isGrandTotal
    ? rawData
    : rawData.filter(r => r.Zone && r.Zone.trim().toLowerCase() === selectedZone.trim().toLowerCase());

  const filteredData = effectiveWeek === 'All Weeks'
    ? zoneAllData
    : zoneAllData.filter(r => r.Week && r.Week.toLowerCase().includes(effectiveWeek.toLowerCase()));

  const totalExpectedWoredas = filteredData.length;
  const reportingWoredas = filteredData.filter(r => r.Report && r.Report.toLowerCase() === 'reported');
  const nonReportingRows = filteredData.filter(r => !r.Report || r.Report.toLowerCase() !== 'reported');
  const nonReportingWoredas = nonReportingRows.map(r => isGrandTotal ? `${r.Woreda} (${r.Zone})` : r.Woreda).filter(Boolean);
  const reportingRate = totalExpectedWoredas > 0
    ? `${((reportingWoredas.length / totalExpectedWoredas) * 100).toFixed(1)}% (${reportingWoredas.length}/${totalExpectedWoredas})`
    : '0%';

  const headers = Object.keys(rawData[0] || {});
  const indicatorPairs = [];

  headers.forEach(header => {
    if (header.includes('Eligible')) {
      const baseName = header.replace('Eligible', '').trim();
      const performedKey = headers.find(h => h.startsWith(baseName) && !h.includes('Eligible') && !h.includes('Target'));
      if (performedKey) {
        let catKey = null;
        if (baseName.startsWith('1.')) catKey = 'cat1_srh';
        else if (baseName.startsWith('2.')) catKey = 'cat2_maternal';
        else if (baseName.startsWith('3.')) catKey = 'cat3_epi';
        else if (baseName.startsWith('4.')) catKey = 'cat4_child';
        else if (baseName.startsWith('5.')) catKey = 'cat5_nutrition';

        if (catKey) {
          indicatorPairs.push({ catKey, baseName, performedKey, eligibleKey: header });
        }
      }
    }
  });

  const computeIndicatorMetric = (rows, perfCol, eligCol = null) => {
    let totalPerf = 0;
    let totalElig = 0;
    rows.forEach(r => {
      totalPerf += parseFloat(r[perfCol]) || 0;
      if (eligCol) totalElig += parseFloat(r[eligCol]) || 0;
    });
    const pct = (eligCol && totalElig > 0) ? ((totalPerf / totalElig) * 100) : null;
    return { totalPerf, totalElig, pct };
  };

  const regionalGrandTotalCov = {};
  indicatorPairs.forEach(({ performedKey, eligibleKey }) => {
    const res = computeIndicatorMetric(regionalWeekData, performedKey, eligibleKey);
    regionalGrandTotalCov[performedKey] = res.pct !== null ? res.pct : 0;
  });

  // 1. EXECUTIVE SUMMARY SELECTED INDICATORS
  const findPairOrCol = (patterns) => {
    const pair = indicatorPairs.find(p => patterns.some(pat => p.performedKey.toLowerCase().includes(pat)));
    if (pair) return pair;
    const col = findColumnByPatterns(headers, patterns);
    return col ? { performedKey: col, eligibleKey: null } : null;
  };

  const execSpecs = [
    { cat: 'Category 01 (SRH)', label: 'Contraceptive Acceptance Rate', patterns: ['contraceptive accept', 'car'], countOnly: false },
    { cat: 'Category 02 (Maternal)', label: 'Maternal Death', patterns: ['maternal death', '2.8'], countOnly: true },
    { cat: 'Category 02 (Maternal)', label: 'Early ANC', patterns: ['early anc'], countOnly: false },
    { cat: 'Category 02 (Maternal)', label: 'ANC 8', patterns: ['anc 8', 'anc-8', 'anc8'], countOnly: false },
    { cat: 'Category 03 (EPI)', label: 'Penta 1', patterns: ['penta 1', 'penta-1', 'pentavalent 1', 'penta1'], countOnly: false },
    { cat: 'Category 03 (EPI)', label: 'Penta 3', patterns: ['penta 3', 'penta-3', 'pentavalent 3', 'penta3'], countOnly: false },
    { cat: 'Category 03 (EPI)', label: 'MCV 1', patterns: ['mcv 1', 'mcv-1', 'measles 1', 'mcv1'], countOnly: false },
    { cat: 'Category 04 (Child Health)', label: 'Pneumonia Treatment', patterns: ['pneumonia'], countOnly: false },
    { cat: 'Category 04 (Child Health)', label: 'Neonatal Death', patterns: ['neonatal death', '4.4'], countOnly: true },
    { cat: 'Category 05 (Nutrition)', label: 'Nutritional Screening', patterns: ['nutritional screen', 'nutrition screen', 'screening'], countOnly: false }
  ];

  const execSummaryBullets = execSpecs.map(spec => {
    const match = findPairOrCol(spec.patterns);
    if (!match) return `- **${spec.cat} — ${spec.label}:** N/A`;
    const zMetric = computeIndicatorMetric(filteredData, match.performedKey, spec.countOnly ? null : match.eligibleKey);
    const rMetric = computeIndicatorMetric(regionalWeekData, match.performedKey, spec.countOnly ? null : match.eligibleKey);

    if (spec.countOnly || zMetric.pct === null) {
      return `- **${spec.cat} — ${match.performedKey}:** ${zMetric.totalPerf} case(s) *(Regional Grand Total: ${rMetric.totalPerf})*`;
    }
    const rPctStr = rMetric.pct !== null ? `${rMetric.pct.toFixed(1)}%` : 'N/A';
    return `- **${spec.cat} — ${match.performedKey}:** **${zMetric.pct.toFixed(1)}%** (${zMetric.totalPerf}/${zMetric.totalElig}) *(Regional Grand Total: ${rPctStr})*`;
  }).join('\n');

  // 2. CATEGORY-BY-CATEGORY PERFORMANCE TABLES (Deterministic Markdown)
  const categoryLabels = {
    cat1_srh: 'Category 01: Sexual & Reproductive Health (SRH)',
    cat2_maternal: 'Category 02: Maternal Health Services',
    cat3_epi: 'Category 03: Expanded Program on Immunization (EPI)',
    cat4_child: 'Category 04: Child Health Services',
    cat5_nutrition: 'Category 05: Nutrition Services'
  };

  const categoryBlocks = {};
  const compactCategoryBriefs = [];

  Object.keys(categoryLabels).forEach(catKey => {
    const pairs = indicatorPairs.filter(p => p.catKey === catKey);
    const tableRows = [
      '| Indicator | Performed | Eligible | Coverage (%) | Regional Grand Total (%) | Difference |',
      '| :--- | :---: | :---: | :---: | :---: | :---: |'
    ];
    const briefItems = [];

    pairs.forEach(({ performedKey, eligibleKey }) => {
      const m = computeIndicatorMetric(filteredData, performedKey, eligibleKey);
      const cov = m.pct !== null ? m.pct : 0;
      const regCov = regionalGrandTotalCov[performedKey] || 0;
      const diffNum = Number((cov - regCov).toFixed(1));
      const diffLabel = diffNum >= 0 ? `+${diffNum.toFixed(1)}%` : `${diffNum.toFixed(1)}%`;

      tableRows.push(`| ${performedKey} | ${m.totalPerf} | ${m.totalElig} | ${cov.toFixed(1)}% | ${regCov.toFixed(1)}% | ${diffLabel} |`);
      briefItems.push(`${performedKey}: ${cov.toFixed(1)}% vs Reg ${regCov.toFixed(1)}% (${diffLabel})`);
    });

    categoryBlocks[catKey] = {
      title: categoryLabels[catKey],
      tableMarkdown: pairs.length ? tableRows.join('\n') : '*No indicators recorded in this category.*'
    };
    compactCategoryBriefs.push(`${categoryLabels[catKey]}: ${briefItems.join('; ')}`);
  });

  // 3. SIX-POINT DATA QUALITY & CONSISTENCY ASSESSMENT
  const missingDataLogs = [];
  const sdOutlierLogs = [];
  const consistencyLogs = [];
  const repeatedWeeksLogs = [];
  const horizontalDupLogs = [];
  const sbaHepDiscrepancyLogs = [];

  const penta1Col = findColumnByPatterns(headers, ['penta 1', 'penta-1', 'pentavalent 1', 'penta1']);
  const penta3Col = findColumnByPatterns(headers, ['penta 3', 'penta-3', 'pentavalent 3', 'penta3']);
  const mcv1Col = findColumnByPatterns(headers, ['mcv 1', 'mcv-1', 'measles 1', 'mcv1']);
  const mcv2Col = findColumnByPatterns(headers, ['mcv 2', 'mcv-2', 'measles 2', 'mcv2']);
  const earlyAncCol = findColumnByPatterns(headers, ['early anc']);
  const anc1Col = findColumnByPatterns(headers, ['anc 1', 'anc-1', 'anc1']);
  const anc4Col = findColumnByPatterns(headers, ['anc 4', 'anc-4', 'anc4']);
  const anc8Col = findColumnByPatterns(headers, ['anc 8', 'anc-8', 'anc8']);
  const sbaCol = findColumnByPatterns(headers, ['skilled birth', 'sba', 'delivery']);
  const hepBdCol = findColumnByPatterns(headers, ['hepatitis', 'hep b', 'hepb']);

  const historyByWoreda = {};
  zoneAllData.forEach(r => {
    const key = `${r.Zone}__${r.Woreda}`;
    if (!historyByWoreda[key]) historyByWoreda[key] = [];
    historyByWoreda[key].push(r);
  });

  reportingWoredas.forEach(row => {
    const woredaLabel = isGrandTotal ? `${row.Woreda} (${row.Zone})` : row.Woreda;
    const key = `${row.Zone}__${row.Woreda}`;
    const woredaHistory = historyByWoreda[key] || [];

    // Check 1: Missing Values
    const missingCols = [];
    indicatorPairs.forEach(({ performedKey }) => {
      const val = row[performedKey];
      if (val === undefined || val === null || val.toString().trim() === '') {
        missingCols.push(performedKey);
      }
    });
    if (missingCols.length > 0) {
      missingDataLogs.push(`**${woredaLabel}:** ${missingCols.join(', ')}`);
    }

    // Check 2: Historical Outliers (>= 3 SD)
    const priorWeeks = woredaHistory.filter(h => h.Report?.toLowerCase() === 'reported' && h.Week !== row.Week);
    if (priorWeeks.length >= 3) {
      const woredaOutliers = [];
      indicatorPairs.forEach(({ performedKey }) => {
        const currVal = parseFloat(row[performedKey]);
        if (isNaN(currVal)) return;
        const prevVals = priorWeeks.map(p => parseFloat(p[performedKey])).filter(v => !isNaN(v));
        if (prevVals.length >= 3) {
          const mean = prevVals.reduce((a, b) => a + b, 0) / prevVals.length;
          const variance = prevVals.reduce((sum, v) => sum + Math.pow(v - mean, 2), 0) / prevVals.length;
          const sd = Math.sqrt(variance);
          if (sd > 0) {
            const zScore = Math.abs(currVal - mean) / sd;
            if (zScore >= 3) {
              woredaOutliers.push(`${performedKey}: ${currVal} (${zScore.toFixed(1)} SD vs mean ${mean.toFixed(1)})`);
            }
          }
        }
      });
      if (woredaOutliers.length > 0) {
        sdOutlierLogs.push(`**${woredaLabel}:** ${woredaOutliers.join('; ')}`);
      }
    }

    // Check 3: Logical Consistency
    const getNum = (col) => (col ? parseFloat(row[col]) : NaN);
    const woredaInconsistencies = [];
    const p1 = getNum(penta1Col);
    const p3 = getNum(penta3Col);
    if (!isNaN(p1) && !isNaN(p3) && p3 > p1) woredaInconsistencies.push(`Penta 3 (${p3}) > Penta 1 (${p1})`);
    const m1 = getNum(mcv1Col);
    const m2 = getNum(mcv2Col);
    if (!isNaN(m1) && !isNaN(m2) && m2 > m1) woredaInconsistencies.push(`MCV 2 (${m2}) > MCV 1 (${m1})`);
    const eAnc = getNum(earlyAncCol);
    const a1 = getNum(anc1Col);
    const a4 = getNum(anc4Col);
    const a8 = getNum(anc8Col);
    if (!isNaN(eAnc) && !isNaN(a1) && eAnc > a1) woredaInconsistencies.push(`Early ANC (${eAnc}) > ANC 1 (${a1})`);
    if (!isNaN(a1) && !isNaN(a4) && a4 > a1) woredaInconsistencies.push(`ANC 4 (${a4}) > ANC 1 (${a1})`);
    if (!isNaN(a1) && !isNaN(a8) && a8 > a1) woredaInconsistencies.push(`ANC 8 (${a8}) > ANC 1 (${a1})`);
    if (!isNaN(a4) && !isNaN(a8) && a8 > a4) woredaInconsistencies.push(`ANC 8 (${a8}) > ANC 4 (${a4})`);

    if (woredaInconsistencies.length > 0) {
      consistencyLogs.push(`**${woredaLabel}:** ${woredaInconsistencies.join('; ')}`);
    }

    // Check 4: Same Value Across Multiple Weeks (3+ consecutive weeks)
    if (priorWeeks.length >= 2) {
      const repeatedCols = [];
      indicatorPairs.forEach(({ performedKey }) => {
        const currVal = parseFloat(row[performedKey]);
        if (isNaN(currVal) || currVal === 0) return;
        const prevVals = priorWeeks.map(p => parseFloat(p[performedKey])).filter(v => !isNaN(v));
        if (prevVals.length >= 2 && prevVals.slice(-2).every(v => v === currVal)) {
          repeatedCols.push(`${performedKey} (${currVal})`);
        }
      });
      if (repeatedCols.length > 0) {
        repeatedWeeksLogs.push(`**${woredaLabel}:** ${repeatedCols.join(', ')}`);
      }
    }

    // Check 5: Same Value Horizontally Across Different Categories
    const valueGroups = {};
    indicatorPairs.forEach(({ catKey, performedKey }) => {
      const val = parseFloat(row[performedKey]);
      if (!isNaN(val) && val > 0) {
        if (!valueGroups[val]) valueGroups[val] = [];
        valueGroups[val].push({ catKey, performedKey });
      }
    });
    Object.entries(valueGroups).forEach(([val, items]) => {
      const distinctCategories = new Set(items.map(i => i.catKey));
      if (items.length >= 3 && distinctCategories.size >= 2) {
        const names = items.map(i => i.performedKey).join(', ');
        horizontalDupLogs.push(`**${woredaLabel}:** Count \`${val}\` repeated across [${names}]`);
      }
    });

    // Check 6: SBA > Hepatitis BD Discrepancy
    const sbaVal = getNum(sbaCol);
    const hepVal = getNum(hepBdCol);
    if (!isNaN(sbaVal) && !isNaN(hepVal) && sbaVal > hepVal) {
      const diff = sbaVal - hepVal;
      sbaHepDiscrepancyLogs.push(`**${woredaLabel}:** SBA (${sbaVal}) vs Hepatitis BD (${hepVal}) — Difference: **+${diff}**`);
    }
  });

  const formatAuditList = (items) => (items.length ? items.map(i => `   - ${i}`).join('\n') : '   - None');

  const qualityAuditMarkdown = [
    `1. **Missing Values**\n${formatAuditList(missingDataLogs)}`,
    `2. **Historical Outliers (>= 3 SD from Previous Weeks)**\n${formatAuditList(sdOutlierLogs)}`,
    `3. **Consistency Between Related Indicators (Penta 1/3, MCV 1/2, Early ANC/ANC 1/4/8)**\n${formatAuditList(consistencyLogs)}`,
    `4. **Same Value Across Multiple Weeks**\n${formatAuditList(repeatedWeeksLogs)}`,
    `5. **Same Value Horizontally Across Different Categories**\n${formatAuditList(horizontalDupLogs)}`,
    `6. **Skilled Birth Attendance vs. Hepatitis BD Discrepancy (SBA > Hepatitis BD)**\n${formatAuditList(sbaHepDiscrepancyLogs)}`
  ].join('\n\n');

  // 4. WOREDA CATEGORIZATION BY DISTANCE FROM REGIONAL TARGETS
  const resolvedTargets = REGIONAL_TARGETS.map(t => {
    const pair = indicatorPairs.find(p => t.patterns.some(pat => p.performedKey.toLowerCase().includes(pat)));
    return pair ? { ...t, performedKey: pair.performedKey, eligibleKey: pair.eligibleKey } : null;
  }).filter(Boolean);

  const woredaTargetEvaluations = [];
  reportingWoredas.forEach(row => {
    let totalDistance = 0;
    let evaluatedCount = 0;
    const indicatorGaps = [];

    resolvedTargets.forEach(({ label, target, performedKey, eligibleKey }) => {
      const perf = parseFloat(row[performedKey]);
      const elig = parseFloat(row[eligibleKey]);
      if (!isNaN(perf) && !isNaN(elig) && elig > 0) {
        const actualPct = (perf / elig) * 100;
        const distance = actualPct - target;
        totalDistance += distance;
        evaluatedCount++;
        indicatorGaps.push({ label, actualPct, target, distance });
      }
    });

    if (evaluatedCount > 0) {
      const avgDistance = Number((totalDistance / evaluatedCount).toFixed(1));
      const woredaLabel = isGrandTotal ? `${row.Woreda} (${row.Zone})` : row.Woreda;
      indicatorGaps.sort((a, b) => a.distance - b.distance);
      const furthestBelow = indicatorGaps
        .filter(g => g.distance < 0)
        .slice(0, 2)
        .map(g => `${g.label} (${g.actualPct.toFixed(1)}% vs ${g.target}%, ${g.distance.toFixed(1)}%)`)
        .join('; ');

      let tier = 'Below Target (Needs Close Follow-Up)';
      if (avgDistance >= 0) tier = 'Meeting / Exceeding Regional Targets';
      else if (avgDistance >= -15) tier = 'Near Target (Within 15% of Regional Target)';

      woredaTargetEvaluations.push({
        woreda: woredaLabel,
        avgDistance,
        tier,
        furthestBelow: furthestBelow || 'All evaluated targets met/exceeded'
      });
    }
  });

  woredaTargetEvaluations.sort((a, b) => b.avgDistance - a.avgDistance);

  const tiersOrder = [
    'Meeting / Exceeding Regional Targets',
    'Near Target (Within 15% of Regional Target)',
    'Below Target (Needs Close Follow-Up)'
  ];

  const targetCategorizationMarkdown = tiersOrder.map(tierName => {
    const members = woredaTargetEvaluations.filter(w => w.tier === tierName);
    if (!members.length) return `### ${tierName}\n- None`;
    const rows = members.map(w => {
      const sign = w.avgDistance >= 0 ? '+' : '';
      return `- **${w.woreda}** — Avg Distance from Target: **${sign}${w.avgDistance}%** *(Key Gaps: ${w.furthestBelow})*`;
    }).join('\n');
    return `### ${tierName}\n${rows}`;
  }).join('\n\n');

  return {
    zoneName: selectedZone || 'Grand Total',
    currentWeek: effectiveWeek,
    reportingRate,
    nonReportingWoredas: nonReportingWoredas.join(', ') || 'None',
    execSummaryBullets,
    categoryBlocks,
    compactCategoryBriefs: compactCategoryBriefs.join('\n'),
    qualityAuditMarkdown,
    targetCategorizationMarkdown,
    auditCounts: {
      missing: missingDataLogs.length,
      outliers: sdOutlierLogs.length,
      inconsistencies: consistencyLogs.length,
      repeatedWeeks: repeatedWeeksLogs.length,
      horizontalDups: horizontalDupLogs.length,
      sbaHep: sbaHepDiscrepancyLogs.length
    },
    tierCounts: {
      meeting: woredaTargetEvaluations.filter(w => w.tier === tiersOrder[0]).length,
      near: woredaTargetEvaluations.filter(w => w.tier === tiersOrder[1]).length,
      below: woredaTargetEvaluations.filter(w => w.tier === tiersOrder[2]).length
    }
  };
}

app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'index.html')));
app.get('/health', (req, res) => res.status(200).send('OK - Server is awake'));

app.get('/api/filters', async (req, res) => {
  try {
    const rawData = await fetchFinalSheetData();
    checkAndSendMaternalDeathAlerts(rawData).catch(() => {});
    const zones = [...new Set(rawData.map(r => r.Zone).filter(Boolean))];
    const weeks = [...new Set(rawData.map(r => r.Week).filter(Boolean))];
    res.json({ zones, weeks });
  } catch (err) {
    console.error('Filter Error:', err);
    res.status(500).json({ zones: [], weeks: [] });
  }
});

/**
 * COMPACT LLM NARRATION + DETERMINISTIC REPORT ASSEMBLY:
 * Sends only a ~600-token statistical digest to Groq so it NEVER exceeds Groq's 8,000 TPM limit,
 * then injects the AI takeaways into the exact standardized 4-section report template.
 */
async function generateAiReport(analytics) {
  const {
    zoneName,
    currentWeek,
    reportingRate,
    nonReportingWoredas,
    execSummaryBullets,
    categoryBlocks,
    compactCategoryBriefs,
    qualityAuditMarkdown,
    targetCategorizationMarkdown,
    auditCounts,
    tierCounts
  } = analytics;

  const compactPrompt = `
Scope: ${zoneName} | Week: ${currentWeek} | Reporting Rate: ${reportingRate}
Category Performance vs Regional Grand Total:
${compactCategoryBriefs}
Quality Audit Counts: Missing=${auditCounts.missing}, Outliers>=3SD=${auditCounts.outliers}, Inconsistencies=${auditCounts.inconsistencies}, RepeatedWeeks=${auditCounts.repeatedWeeks}, HorizontalDups=${auditCounts.horizontalDups}, SBA>HepBD=${auditCounts.sbaHep}
Target Tier Counts: Meeting=${tierCounts.meeting}, Near=${tierCounts.near}, Below=${tierCounts.below}
  `.trim();

  const systemPrompt = `You are an expert RMNCH Health Data Analyst. Based on the compact metrics provided, return ONLY a valid JSON object with 6 concise, analytical sentences (one executive overview sentence and one key takeaway sentence per category). Do not include markdown fences or extra text.
Required JSON keys:
{
  "execOverview": "1-2 sentence executive takeaway on reporting and key indicator performance.",
  "cat1_srh": "1-2 sentence key takeaway for Category 01 SRH compared to Regional Grand Total.",
  "cat2_maternal": "1-2 sentence key takeaway for Category 02 Maternal Health compared to Regional Grand Total.",
  "cat3_epi": "1-2 sentence key takeaway for Category 03 EPI compared to Regional Grand Total.",
  "cat4_child": "1-2 sentence key takeaway for Category 04 Child Health compared to Regional Grand Total.",
  "cat5_nutrition": "1-2 sentence key takeaway for Category 05 Nutrition compared to Regional Grand Total."
}`;

  let takeaways = {
    execOverview: `Reporting rate for ${zoneName} in ${currentWeek} stands at ${reportingRate}.`,
    cat1_srh: 'Review family planning uptake indicators against the Regional Grand Total benchmark.',
    cat2_maternal: 'Prioritize early ANC initiation, ANC 8 completion, and immediate MDSR follow-up where applicable.',
    cat3_epi: 'Monitor antigen dropout rates between Penta 1, Penta 3, and MCV 1 relative to the regional average.',
    cat4_child: 'Ensure consistent case management for childhood pneumonia and timely neonatal surveillance.',
    cat5_nutrition: 'Strengthen routine nutritional screening coverage across all reporting woredas.'
  };

  const modelsToTry = ['openai/gpt-oss-20b', 'llama-3.1-8b-instant'];
  for (const modelName of modelsToTry) {
    try {
      const chatCompletion = await groq.chat.completions.create({
        messages: [
          { role: 'system', content: systemPrompt },
          { role: 'user', content: compactPrompt }
        ],
        model: modelName,
        temperature: 0.2,
        max_completion_tokens: 700
      });

      const rawReply = chatCompletion.choices[0]?.message?.content || '';
      const jsonMatch = rawReply.match(/\{[\s\S]*\}/);
      if (jsonMatch) {
        const parsed = JSON.parse(jsonMatch[0]);
        takeaways = { ...takeaways, ...parsed };
      }
      break;
    } catch (err) {
      console.warn(`[Groq Model ${modelName} Warning]:`, err.message);
    }
  }

  const catKeys = ['cat1_srh', 'cat2_maternal', 'cat3_epi', 'cat4_child', 'cat5_nutrition'];
  const categorySections = catKeys.map(k => {
    const block = categoryBlocks[k];
    return `### ${block.title}\n${block.tableMarkdown}\n\n**Key Takeaway:** ${takeaways[k]}`;
  }).join('\n\n');

  return (
    `Hello, This Is Weekly Essential RMNCH Service Report Feedback for ${zoneName}, ${currentWeek}\n\n` +
    `## 1. Executive Summary and Reporting Overview\n` +
    `- **Reporting Rate:** ${reportingRate}\n` +
    `- **Non-Reporting Woredas:** ${nonReportingWoredas}\n` +
    `- **Overview:** ${takeaways.execOverview}\n\n` +
    `### Selected Priority Indicators Summary\n` +
    `${execSummaryBullets}\n\n` +
    `## 2. Category-by-Category Performance Analysis (vs. Regional Grand Total)\n\n` +
    `${categorySections}\n\n` +
    `## 3. Data Quality and Consistency Assessment\n\n` +
    `${qualityAuditMarkdown}\n\n` +
    `## 4. Woreda Performance Categorization (Distance from Regional Targets)\n` +
    `*Regional Targets Reference: Contraceptive Acceptance (96%), Long-Acting Contraceptive (48%), Early ANC (57%), ANC 4 (98%), ANC 8 (66%), Skilled Birth Attendance (89%), Hepatitis BD (87%), Penta 1 (100%), Penta 3 (100%).*\n\n` +
    `${targetCategorizationMarkdown}` +
    `${REPORT_FOOTER}`
  );
}

app.post('/api/chat', async (req, res) => {
  try {
    const { selectedZone, selectedWeek } = req.body;
    const rawData = await fetchFinalSheetData();
    if (!rawData.length) return res.status(500).json({ reply: 'Could not retrieve data from Google Sheet.' });

    await checkAndSendMaternalDeathAlerts(rawData, selectedWeek);

    const analytics = buildAnalyticsPayload(rawData, selectedZone || 'Grand Total', selectedWeek || 'Current Week');
    const responseText = await generateAiReport(analytics);
    res.json({ reply: responseText });
  } catch (error) {
    console.error('Groq/Sheet Error Details:', error);
    res.status(500).json({ reply: 'Error generating report. Please try again.' });
  }
});

app.post('/api/export-docx', async (req, res) => {
  try {
    const { reportText, title } = req.body;
    if (typeof reportText !== 'string' || !reportText.trim()) {
      return res.status(400).json({ error: 'Report text is required.' });
    }

    const lines = reportText.split('\n');
    const docChildren = [new Paragraph({
      text: title || 'Weekly Essential RMNCH Service Report Feedback',
      heading: HeadingLevel.HEADING_1,
      spacing: { after: 200 }
    })];

    let i = 0;
    while (i < lines.length) {
      const line = lines[i].trim();
      if (line.startsWith('|') && line.endsWith('|')) {
        const tableRows = [];
        while (i < lines.length && lines[i].trim().startsWith('|') && lines[i].trim().endsWith('|')) {
          const rowLine = lines[i].trim();
          if (!/^\|?\s*:?-{3,}/.test(rowLine.replace(/\|/g, '').trim())) {
            const cells = rowLine.split('|').slice(1, -1).map(c => c.trim());
            tableRows.push(cells);
          }
          i++;
        }
        if (tableRows.length) {
          const columnCount = Math.max(...tableRows.map(row => row.length));
          const docTableRows = tableRows.map((rowCells, rowIndex) => {
            const isHeader = rowIndex === 0;
            return new TableRow({
              children: Array.from({ length: columnCount }, (_, colIndex) => {
                const cellText = rowCells[colIndex] || '';
                return new TableCell({
                  children: [new Paragraph({
                    alignment: isHeader ? AlignmentType.CENTER : AlignmentType.LEFT,
                    children: [new TextRun({
                      text: cellText, bold: isHeader,
                      color: isHeader ? 'FFFFFF' : '000000', size: 18
                    })]
                  })],
                  width: { size: Math.floor(10000 / columnCount), type: WidthType.DXA },
                  shading: {
                    fill: isHeader ? '1F4E78' : (rowIndex % 2 === 0 ? 'F2F4F8' : 'FFFFFF'),
                    val: ShadingType.CLEAR, color: 'auto'
                  }
                });
              })
            });
          });
          docChildren.push(new Table({
            rows: docTableRows,
            width: { size: 100, type: WidthType.PERCENTAGE }
          }));
          docChildren.push(new Paragraph({ text: '', spacing: { after: 150 } }));
        }
      } else {
        if (line && line !== '---') {
          const isHeading = line.startsWith('## ') || line.startsWith('### ');
          const cleanText = lines[i].replace(/^#{1,3}\s+/, '').replace(/\*\*/g, '').replace(/\*/g, '');
          docChildren.push(new Paragraph({
            heading: isHeading ? HeadingLevel.HEADING_2 : undefined,
            children: [new TextRun({ text: cleanText, bold: isHeading, size: isHeading ? 26 : 22 })],
            spacing: { before: isHeading ? 160 : 0, after: 100 }
          }));
        }
        i++;
      }
    }

    const doc = new Document({ sections: [{ properties: {}, children: docChildren }] });
    const buffer = await Packer.toBuffer(doc);
    const token = crypto.randomUUID();
    const filename = `RMNCH_Service_Report_${Date.now()}.docx`;

    temporaryDocuments.set(token, {
      buffer,
      filename,
      expiresAt: Date.now() + DOWNLOAD_EXPIRY
    });

    const baseUrl = (process.env.APP_BASE_URL || `${req.protocol}://${req.get('host')}`).replace(/\/$/, '');
    const url = `${baseUrl}/api/download-docx/${token}`;
    res.json({ success: true, url, filename });
  } catch (err) {
    console.error('Docx Export Error:', err);
    res.status(500).json({ error: 'Failed to generate Word document.' });
  }
});

app.get('/api/download-docx/:token', (req, res) => {
  const token = req.params.token;
  const file = temporaryDocuments.get(token);
  if (!file || Date.now() >= file.expiresAt) {
    temporaryDocuments.delete(token);
    return res.status(404).send('This download link has expired. Please generate the report again.');
  }

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
  res.setHeader('Content-Disposition', `attachment; filename="${file.filename}"`);
  res.setHeader('Content-Length', file.buffer.length);
  res.setHeader('Cache-Control', 'no-store');
  return res.send(file.buffer);
});

async function sendViaBrevo(toInput, subject, textBody) {
  const recipients = Array.isArray(toInput) ? toInput : [toInput];
  const toList = recipients.map(email => ({ email: email.trim() }));
  const response = await fetch('https://api.brevo.com/v3/smtp/email', {
    method: 'POST',
    headers: {
      accept: 'application/json',
      'api-key': BREVO_API_KEY,
      'content-type': 'application/json'
    },
    body: JSON.stringify({
      sender: { name: 'Dr. Esubalew Zerihun', email: 'esubalewzerihun11@gmail.com' },
      to: toList, subject, textContent: textBody
    })
  });
  if (!response.ok) {
    const errData = await response.json();
    throw new Error(`Brevo API Error: ${JSON.stringify(errData)}`);
  }
  return response.json();
}

async function runRmnchEmailDispatch(forcedDay = null) {
  console.log('[RMNCH Email Dispatch] Starting execution via Brevo...');
  try {
    const rawData = await fetchFinalSheetData();
    if (!rawData?.length) return;
    const currentWeek = getCurrentReportingWeek(rawData);

    await checkAndSendMaternalDeathAlerts(rawData, currentWeek);

    const zones = [...new Set(rawData.map(r => r.Zone).filter(Boolean))];
    const dayOfWeek = forcedDay !== null ? forcedDay : new Date().getDay();

    for (const zone of zones) {
      const cleanZoneKey = zone.toString().trim().toLowerCase();
      const recipientEmail = ZONAL_EMAILS[cleanZoneKey];
      if (!recipientEmail) continue;

      const zoneData = rawData.filter(r =>
        r.Zone?.toString().trim().toLowerCase() === cleanZoneKey &&
        r.Week?.toString().trim().toLowerCase().includes(currentWeek.toLowerCase())
      );
      const totalExpected = zoneData.length;
      const reported = zoneData.filter(r => r.Report && r.Report.toLowerCase() === 'reported');
      const reportingRateStr = totalExpected > 0
        ? `${((reported.length / totalExpected) * 100).toFixed(1)}%` : '0%';
      const nonReporting = zoneData
        .filter(r => !r.Report || r.Report.toLowerCase() !== 'reported')
        .map(r => r.Woreda).join(', ') || 'None';

      let zonalBody = '';
      let subjectLine = '';

      if (dayOfWeek === 1) {
        subjectLine = `[Weekly Status Update] ${zone} - ${currentWeek}`;
        zonalBody = `Dear ${zone} MCH Team,\n\n` +
          `Greetings!\n\n` +
          `Please be reminded that the weekly RMNCH service report is expected from today until Wednesday at 6:30.\n\n` +
          `Current Reporting Rate: ${reportingRateStr} (${reported.length}/${totalExpected} Woredas reported)\n` +
          `Current Non-Reporting Woredas: ${nonReporting}\n\n` +
          `Kindly submit your reports before the deadline.` +
          REPORT_FOOTER;
      } else if (dayOfWeek === 2) {
        subjectLine = `[Weekly Status Update - Reminder] ${zone} - ${currentWeek}`;
        zonalBody = `Dear ${zone} MCH Team,\n\n` +
          `Greetings!\n\n` +
          `This is your daily update regarding the weekly RMNCH service reporting for ${currentWeek} (Deadline: Wednesday 6:30).\n\n` +
          `Current Reporting Rate: ${reportingRateStr} (${reported.length}/${totalExpected} Woredas reported)\n` +
          `Remaining Non-Reporting Woredas: ${nonReporting}\n\n` +
          `Please ensure timely submission.` +
          REPORT_FOOTER;
      } else if (dayOfWeek === 3) {
        subjectLine = `[Weekly Standardized Feedback & Analysis] ${zone} - ${currentWeek}`;
        const analytics = buildAnalyticsPayload(rawData, zone, currentWeek);
        const standardizedReport = await generateAiReport(analytics);

        zonalBody = `Dear ${zone} MCH Team,\n\n` +
          `Greetings!\n\n` +
          `The reporting window for ${currentWeek} has closed. Below is the list of non-reporting woredas followed by your standardized weekly feedback report:\n\n` +
          `Non-Reporting Woredas: ${nonReporting}\n\n` +
          `---------------------------------------------------\n\n` +
          standardizedReport;
      }

      if (zonalBody) {
        await sendViaBrevo(recipientEmail, subjectLine, zonalBody);
        await new Promise(resolve => setTimeout(resolve, 2500));
      }
    }

    if (dayOfWeek === 3) {
      const regionalAnalytics = buildAnalyticsPayload(rawData, 'Grand Total', currentWeek);
      const regionalStandardizedReport = await generateAiReport(regionalAnalytics);

      const regionalBody = `Dear Regional Leadership Team,\n\n` +
        `Formal Greetings,\n\n` +
        `Please find below the standardized Regional Grand Total RMNCH feedback report for ${currentWeek}:\n\n` +
        `---------------------------------------------------\n\n` +
        regionalStandardizedReport;

      await sendViaBrevo(
        REGIONAL_EMAILS,
        `[Regional Grand Total Standardized Report] RMNCH Feedback - ${currentWeek}`,
        regionalBody
      );
    }

    return true;
  } catch (err) {
    console.error('Error in automated email dispatch via Brevo:', err);
    throw err;
  }
}

app.get('/api/test-email', async (req, res) => {
  try {
    const forcedDay = req.query.day ? parseInt(req.query.day, 10) : null;
    await runRmnchEmailDispatch(forcedDay);
    res.send(`RMNCH email test executed successfully via Brevo (Forced Day: ${forcedDay !== null ? forcedDay : 'Today'})!`);
  } catch (err) {
    res.status(500).send('Error executing email dispatch: ' + err.message);
  }
});

cron.schedule('30 10 * * 1,2,3', async () => {
  await runRmnchEmailDispatch();
});

cron.schedule('*/15 * * * *', async () => {
  try {
    const rawData = await fetchFinalSheetData();
    await checkAndSendMaternalDeathAlerts(rawData);
  } catch (err) {
    console.error('[Maternal Death Cron Error]:', err.message);
  }
});

app.listen(PORT, '0.0.0.0', () => {
  console.log(`Server running on port ${PORT}`);
});import express from 'express';
import path from 'path';
import { fileURLToPath } from 'url';
import crypto from 'crypto';
import Groq from 'groq-sdk';
import cron from 'node-cron';
import {
  Document, Packer, Paragraph, TextRun, HeadingLevel,
  Table, TableRow, TableCell, WidthType, ShadingType, AlignmentType
} from 'docx';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
app.use(express.json({ limit: '10mb' }));
app.use(express.static(__dirname));

const PORT = process.env.PORT || 3000;
const groq = new Groq({ apiKey: process.env.GROQ_API_KEY || '' });
const BREVO_API_KEY = process.env.BREVO_API_KEY || '';

const temporaryDocuments = new Map();
const DOWNLOAD_EXPIRY = 15 * 60 * 1000;

// Deduplication tracker for immediate maternal death alerts (key: "Week|Zone|Woreda|Count")
const sentMaternalDeathAlerts = new Set();

const REGIONAL_EMAILS = ['esubalewzerihun11@gmail.com', 'messiwode46@gmail.com'];
const ZONAL_EMAILS = {
  'kaffa zone': ['tadessebelay1221@gmail.com', 'esubalewzerihun37@gmail.com'],
  'dawro zone': ['ketesfu7@gmail.com', 'akliluasfaw423@gmail.com', 'yohannesante@gmail.com'],
  'west omo zone': ['getahunzewde37@gmail.com', 'esubalewzerihun37@gmail.com'],
  'bench sheko zone': ['Joannagashahun13@gmail.com', 'beyenesham32@gmail.com', 'habtekelbi78@gmail.com', 'ruphaelw@gmail.com'],
  'sheka zone': ['sifanmesfin@gmail.com', 'esubalewzerihun37@gmail.com'],
  'konta zone': ['tsegabebe2015@gmail.com', 'degemukidist365@gmail.com']
};

const cleanupTimer = setInterval(() => {
  const now = Date.now();
  for (const [token, file] of temporaryDocuments.entries()) {
    if (now >= file.expiresAt) temporaryDocuments.delete(token);
  }
}, 60 * 1000);
cleanupTimer.unref?.();

const SHEET_ID = '1d8_5Q2W3aGV68W_DdjrPrFEJnuRSDprrb1b3kngbTpo';
const FINAL_SHEET_CSV_URL = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:csv&sheet=Final`;

const REPORT_FOOTER = `\n\n---\n*This report is prepared by **SWERHB MCH Team**. Note: This is a preliminary report; for more information and interactive analytics, please visit the **RMNCH Dashboard**.*`;

// Regional Indicator Targets (%)
const REGIONAL_TARGETS = [
  { label: 'Contraceptive Acceptance Rate (CAR)', target: 96, patterns: ['contraceptive accept', 'car'] },
  { label: 'Long-Acting Family Planning (LAFP)', target: 48, patterns: ['long acting', 'long-acting', 'lafp'] },
  { label: 'Early ANC', target: 57, patterns: ['early anc'] },
  { label: 'ANC 4', target: 98, patterns: ['anc 4', 'anc-4', 'anc4'] },
  { label: 'ANC 8', target: 66, patterns: ['anc 8', 'anc-8', 'anc8'] },
  { label: 'Skilled Birth Attendance (SBA)', target: 89, patterns: ['skilled birth', 'sba', 'delivery'] },
  { label: 'Hepatitis BD', target: 87, patterns: ['hepatitis', 'hep b', 'hepb'] },
  { label: 'Penta 1', target: 100, patterns: ['penta 1', 'penta-1', 'pentavalent 1', 'penta1'] },
  { label: 'Penta 3', target: 100, patterns: ['penta 3', 'penta-3', 'pentavalent 3', 'penta3'] }
];

function parseCSVLine(text) {
  const result = [];
  let current = '';
  let inQuotes = false;
  for (let i = 0; i < text.length; i++) {
    const char = text[i];
    if (char === '"') {
      if (inQuotes && text[i + 1] === '"') {
        current += '"';
        i++;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (char === ',' && !inQuotes) {
      result.push(current.trim());
      current = '';
    } else {
      current += char;
    }
  }
  result.push(current.trim());
  return result;
}

async function fetchFinalSheetData() {
  try {
    const response = await fetch(FINAL_SHEET_CSV_URL);
    if (!response.ok) throw new Error(`Fetch failed: ${response.status} ${response.statusText}`);
    const csvText = await response.text();
    const lines = csvText.split(/\r?\n/).filter(line => line.trim() !== '');
    if (!lines.length) return [];
    const headers = parseCSVLine(lines[0]);
    return lines.slice(1).map(line => {
      const values = parseCSVLine(line);
      const obj = {};
      headers.forEach((header, index) => { obj[header] = values[index] ?? ''; });
      return obj;
    });
  } catch (err) {
    console.error('Sheet Fetch Error:', err);
    throw err;
  }
}

function getCurrentReportingWeek(rawData) {
  const baseDate = new Date('2026-11-09T00:00:00');
  const baseWeekNumber = 19;
  const currentDate = new Date();
  const diffDays = Math.floor((currentDate - baseDate) / (1000 * 60 * 60 * 24));
  const calculatedWeekNum = baseWeekNumber + Math.max(0, Math.floor(diffDays / 7));
  const targetWeekQuery = `Week ${calculatedWeekNum}`;
  const availableWeeks = [...new Set(rawData.map(r => r.Week?.toString().trim()).filter(Boolean))];
  return availableWeeks.find(w => w.toLowerCase().includes(targetWeekQuery.toLowerCase()))
    || availableWeeks[availableWeeks.length - 1];
}

function findColumnByPatterns(headers, patterns, excludeSubstrings = ['Eligible', 'Target']) {
  return headers.find(h =>
    !excludeSubstrings.some(ex => h.includes(ex)) &&
    patterns.some(p => h.toLowerCase().includes(p.toLowerCase()))
  );
}

/**
 * IMMEDIATE MATERNAL DEATH ALERT DISPATCHER
 * Scans current week data and sends an immediate email to Regional Leadership if Maternal Death > 0.
 */
async function checkAndSendMaternalDeathAlerts(rawData, targetWeek = null) {
  if (!rawData?.length || !BREVO_API_KEY) return;
  const activeWeek = targetWeek && targetWeek !== 'All Weeks'
    ? targetWeek
    : getCurrentReportingWeek(rawData);

  const headers = Object.keys(rawData[0] || {});
  const maternalDeathCol = headers.find(h => {
    const low = h.toLowerCase();
    return !low.includes('eligible') && !low.includes('target') &&
      (low.includes('maternal death') || h.startsWith('2.8'));
  });

  if (!maternalDeathCol) return;

  const weekRows = rawData.filter(r =>
    r.Week && r.Week.toString().trim().toLowerCase().includes(activeWeek.toLowerCase())
  );

  const newAlerts = [];
  weekRows.forEach(row => {
    const deaths = parseFloat(row[maternalDeathCol]);
    if (!isNaN(deaths) && deaths > 0) {
      const alertKey = `${row.Week}|${row.Zone}|${row.Woreda}|${deaths}`;
      if (!sentMaternalDeathAlerts.has(alertKey)) {
        sentMaternalDeathAlerts.add(alertKey);
        newAlerts.push({
          week: row.Week,
          zone: row.Zone || 'Unknown Zone',
          woreda: row.Woreda || 'Unknown Woreda',
          deaths
        });
      }
    }
  });

  if (!newAlerts.length) return;

  const alertLines = newAlerts
    .map(a => `- Zone: ${a.zone} | Woreda: ${a.woreda} | Week: ${a.week} | Reported Maternal Deaths: ${a.deaths}`)
    .join('\n');

  const alertBody = `Dear Regional Leadership Team,\n\n` +
    `URGENT NOTIFICATION: Maternal death case(s) have been reported in the Weekly RMNCH Surveillance System.\n\n` +
    `Incident Details:\n${alertLines}\n\n` +
    `Please initiate immediate MDSR (Maternal Death Surveillance and Response) verification and follow-up with the respective Zonal Health Department(s).` +
    REPORT_FOOTER;

  try {
    await sendViaBrevo(
      REGIONAL_EMAILS,
      `[URGENT MATERNAL DEATH ALERT] RMNCH Surveillance - ${activeWeek}`,
      alertBody
    );
    console.log(`[Maternal Death Alert] Dispatched ${newAlerts.length} alert(s) to Regional Leadership.`);
  } catch (err) {
    console.error('[Maternal Death Alert] Failed to send email:', err);
  }
}

/**
 * PURE CODE CALCULATION ENGINE:
 * Computes uniform executive highlights, Regional Grand Total comparisons,
 * 6-point data quality audits (including >= 3 SD outliers, horizontal duplicates, SBA vs Hep BD),
 * and woreda categorization based on distance from regional targets.
 */
function buildAnalyticsPayload(rawData, selectedZone, selectedWeek) {
  const isGrandTotal = !selectedZone || selectedZone === 'Grand Total';
  const effectiveWeek = (!selectedWeek || selectedWeek === 'Current Week')
    ? getCurrentReportingWeek(rawData)
    : selectedWeek;

  // Regional Grand Total slice for the selected week (used as uniform benchmark across all reports)
  const regionalWeekData = effectiveWeek === 'All Weeks'
    ? rawData
    : rawData.filter(r => r.Week && r.Week.toLowerCase().includes(effectiveWeek.toLowerCase()));

  const zoneAllData = isGrandTotal
    ? rawData
    : rawData.filter(r => r.Zone && r.Zone.trim().toLowerCase() === selectedZone.trim().toLowerCase());

  const filteredData = effectiveWeek === 'All Weeks'
    ? zoneAllData
    : zoneAllData.filter(r => r.Week && r.Week.toLowerCase().includes(effectiveWeek.toLowerCase()));

  const totalExpectedWoredas = filteredData.length;
  const reportingWoredas = filteredData.filter(r => r.Report && r.Report.toLowerCase() === 'reported');
  const nonReportingRows = filteredData.filter(r => !r.Report || r.Report.toLowerCase() !== 'reported');
  const nonReportingWoredas = nonReportingRows.map(r => isGrandTotal ? `${r.Woreda} (${r.Zone})` : r.Woreda).filter(Boolean);
  const reportingRate = totalExpectedWoredas > 0
    ? `${((reportingWoredas.length / totalExpectedWoredas) * 100).toFixed(1)}% (${reportingWoredas.length}/${totalExpectedWoredas})`
    : '0%';

  const headers = Object.keys(rawData[0] || {});
  const indicatorPairs = [];

  headers.forEach(header => {
    if (header.includes('Eligible')) {
      const baseName = header.replace('Eligible', '').trim();
      const performedKey = headers.find(h => h.startsWith(baseName) && !h.includes('Eligible') && !h.includes('Target'));
      if (performedKey) {
        let catKey = null;
        if (baseName.startsWith('1.')) catKey = 'cat1_srh';
        else if (baseName.startsWith('2.')) catKey = 'cat2_maternal';
        else if (baseName.startsWith('3.')) catKey = 'cat3_epi';
        else if (baseName.startsWith('4.')) catKey = 'cat4_child';
        else if (baseName.startsWith('5.')) catKey = 'cat5_nutrition';

        if (catKey) {
          indicatorPairs.push({ catKey, baseName, performedKey, eligibleKey: header });
        }
      }
    }
  });

  // Helper to compute coverage or count for a dataset slice
  const computeIndicatorMetric = (rows, perfCol, eligCol = null) => {
    let totalPerf = 0;
    let totalElig = 0;
    rows.forEach(r => {
      totalPerf += parseFloat(r[perfCol]) || 0;
      if (eligCol) totalElig += parseFloat(r[eligCol]) || 0;
    });
    const pct = (eligCol && totalElig > 0) ? ((totalPerf / totalElig) * 100) : null;
    return { totalPerf, totalElig, pct };
  };

  // Regional Grand Total coverage per indicator
  const regionalGrandTotalCov = {};
  indicatorPairs.forEach(({ performedKey, eligibleKey }) => {
    const res = computeIndicatorMetric(regionalWeekData, performedKey, eligibleKey);
    regionalGrandTotalCov[performedKey] = res.pct !== null ? res.pct : 0;
  });

  // 1. EXECUTIVE SUMMARY SELECTED INDICATORS (Uniform across all reports)
  const findPairOrCol = (patterns) => {
    const pair = indicatorPairs.find(p => patterns.some(pat => p.performedKey.toLowerCase().includes(pat)));
    if (pair) return pair;
    const col = findColumnByPatterns(headers, patterns);
    return col ? { performedKey: col, eligibleKey: null } : null;
  };

  const execSpecs = [
    { cat: 'Category 01 (SRH)', label: 'Contraceptive Acceptance Rate', patterns: ['contraceptive accept', 'car'], countOnly: false },
    { cat: 'Category 02 (Maternal)', label: 'Maternal Death', patterns: ['maternal death', '2.8'], countOnly: true },
    { cat: 'Category 02 (Maternal)', label: 'Early ANC', patterns: ['early anc'], countOnly: false },
    { cat: 'Category 02 (Maternal)', label: 'ANC 8', patterns: ['anc 8', 'anc-8', 'anc8'], countOnly: false },
    { cat: 'Category 03 (EPI)', label: 'Penta 1', patterns: ['penta 1', 'penta-1', 'pentavalent 1', 'penta1'], countOnly: false },
    { cat: 'Category 03 (EPI)', label: 'Penta 3', patterns: ['penta 3', 'penta-3', 'pentavalent 3', 'penta3'], countOnly: false },
    { cat: 'Category 03 (EPI)', label: 'MCV 1', patterns: ['mcv 1', 'mcv-1', 'measles 1', 'mcv1'], countOnly: false },
    { cat: 'Category 04 (Child Health)', label: 'Pneumonia Treatment', patterns: ['pneumonia'], countOnly: false },
    { cat: 'Category 04 (Child Health)', label: 'Neonatal Death', patterns: ['neonatal death', '4.4'], countOnly: true },
    { cat: 'Category 05 (Nutrition)', label: 'Nutritional Screening', patterns: ['nutritional screen', 'nutrition screen', 'screening'], countOnly: false }
  ];

  const execSummaryLines = execSpecs.map(spec => {
    const match = findPairOrCol(spec.patterns);
    if (!match) return `- [${spec.cat}] ${spec.label}: Column not found`;
    const zMetric = computeIndicatorMetric(filteredData, match.performedKey, spec.countOnly ? null : match.eligibleKey);
    const rMetric = computeIndicatorMetric(regionalWeekData, match.performedKey, spec.countOnly ? null : match.eligibleKey);

    if (spec.countOnly || zMetric.pct === null) {
      return `- [${spec.cat}] ${match.performedKey}: ${zMetric.totalPerf} case(s) (Regional Grand Total: ${rMetric.totalPerf} case(s))`;
    }
    const rPctStr = rMetric.pct !== null ? `${rMetric.pct.toFixed(1)}%` : 'N/A';
    return `- [${spec.cat}] ${match.performedKey}: ${zMetric.pct.toFixed(1)}% (${zMetric.totalPerf}/${zMetric.totalElig}) vs Regional Grand Total: ${rPctStr}`;
  }).join('\n');

  // 2. CATEGORY-BY-CATEGORY PERFORMANCE VS REGIONAL GRAND TOTAL
  const categoryLabels = {
    cat1_srh: 'Category 01: Sexual & Reproductive Health (SRH)',
    cat2_maternal: 'Category 02: Maternal Health Services',
    cat3_epi: 'Category 03: Expanded Program on Immunization (EPI)',
    cat4_child: 'Category 04: Child Health Services',
    cat5_nutrition: 'Category 05: Nutrition Services'
  };

  const categorizedSummaryText = [];
  Object.keys(categoryLabels).forEach(catKey => {
    const pairs = indicatorPairs.filter(p => p.catKey === catKey);
    const lines = pairs.map(({ performedKey, eligibleKey }) => {
      const m = computeIndicatorMetric(filteredData, performedKey, eligibleKey);
      const cov = m.pct !== null ? m.pct : 0;
      const regCov = regionalGrandTotalCov[performedKey] || 0;
      const diff = (cov - regCov).toFixed(1);
      const diffLabel = Number(diff) >= 0 ? `+${diff}%` : `${diff}%`;
      return `- ${performedKey}: Performed=${m.totalPerf}, Eligible=${m.totalElig}, Coverage=${cov.toFixed(1)}%, Regional Grand Total=${regCov.toFixed(1)}% (Diff: ${diffLabel})`;
    });
    categorizedSummaryText.push(`\n### ${categoryLabels[catKey]}\n${lines.join('\n') || '- No indicators found'}`);
  });

  // 3. SIX-POINT DATA QUALITY & CONSISTENCY ASSESSMENT
  const missingDataLogs = [];
  const sdOutlierLogs = [];
  const consistencyLogs = [];
  const repeatedWeeksLogs = [];
  const horizontalDupLogs = [];
  const sbaHepDiscrepancyLogs = [];

  const penta1Col = findColumnByPatterns(headers, ['penta 1', 'penta-1', 'pentavalent 1', 'penta1']);
  const penta3Col = findColumnByPatterns(headers, ['penta 3', 'penta-3', 'pentavalent 3', 'penta3']);
  const mcv1Col = findColumnByPatterns(headers, ['mcv 1', 'mcv-1', 'measles 1', 'mcv1']);
  const mcv2Col = findColumnByPatterns(headers, ['mcv 2', 'mcv-2', 'measles 2', 'mcv2']);
  const earlyAncCol = findColumnByPatterns(headers, ['early anc']);
  const anc1Col = findColumnByPatterns(headers, ['anc 1', 'anc-1', 'anc1']);
  const anc4Col = findColumnByPatterns(headers, ['anc 4', 'anc-4', 'anc4']);
  const anc8Col = findColumnByPatterns(headers, ['anc 8', 'anc-8', 'anc8']);
  const sbaCol = findColumnByPatterns(headers, ['skilled birth', 'sba', 'delivery']);
  const hepBdCol = findColumnByPatterns(headers, ['hepatitis', 'hep b', 'hepb']);

  const historyByWoreda = {};
  zoneAllData.forEach(r => {
    const key = `${r.Zone}__${r.Woreda}`;
    if (!historyByWoreda[key]) historyByWoreda[key] = [];
    historyByWoreda[key].push(r);
  });

  reportingWoredas.forEach(row => {
    const woredaLabel = isGrandTotal ? `${row.Woreda} (${row.Zone})` : row.Woreda;
    const key = `${row.Zone}__${row.Woreda}`;
    const woredaHistory = historyByWoreda[key] || [];

    // Check 1: Missing Values
    const missingCols = [];
    indicatorPairs.forEach(({ performedKey }) => {
      const val = row[performedKey];
      if (val === undefined || val === null || val.toString().trim() === '') {
        missingCols.push(performedKey);
      }
    });
    if (missingCols.length > 0) {
      missingDataLogs.push(`${woredaLabel}: Missing [${missingCols.join(', ')}]`);
    }

    // Check 2: Historical Outliers (>= 3 Standard Deviations from prior weeks)
    const priorWeeks = woredaHistory.filter(h => h.Report?.toLowerCase() === 'reported' && h.Week !== row.Week);
    if (priorWeeks.length >= 3) {
      indicatorPairs.forEach(({ performedKey }) => {
        const currVal = parseFloat(row[performedKey]);
        if (isNaN(currVal)) return;
        const prevVals = priorWeeks.map(p => parseFloat(p[performedKey])).filter(v => !isNaN(v));
        if (prevVals.length >= 3) {
          const mean = prevVals.reduce((a, b) => a + b, 0) / prevVals.length;
          const variance = prevVals.reduce((sum, v) => sum + Math.pow(v - mean, 2), 0) / prevVals.length;
          const sd = Math.sqrt(variance);
          if (sd > 0) {
            const zScore = Math.abs(currVal - mean) / sd;
            if (zScore >= 3) {
              const dir = currVal > mean ? 'above' : 'below';
              sdOutlierLogs.push(`${woredaLabel}: [${performedKey}] = ${currVal} (${zScore.toFixed(1)} SD ${dir} historical mean ${mean.toFixed(1)}, SD=${sd.toFixed(1)})`);
            }
          }
        }
      });
    }

    // Check 3: Logical Consistency Between Related Indicators
    const getNum = (col) => (col ? parseFloat(row[col]) : NaN);
    const p1 = getNum(penta1Col);
    const p3 = getNum(penta3Col);
    if (!isNaN(p1) && !isNaN(p3) && p3 > p1) {
      consistencyLogs.push(`${woredaLabel}: Penta 3 (${p3}) > Penta 1 (${p1})`);
    }
    const m1 = getNum(mcv1Col);
    const m2 = getNum(mcv2Col);
    if (!isNaN(m1) && !isNaN(m2) && m2 > m1) {
      consistencyLogs.push(`${woredaLabel}: MCV 2 (${m2}) > MCV 1 (${m1})`);
    }
    const eAnc = getNum(earlyAncCol);
    const a1 = getNum(anc1Col);
    const a4 = getNum(anc4Col);
    const a8 = getNum(anc8Col);
    if (!isNaN(eAnc) && !isNaN(a1) && eAnc > a1) {
      consistencyLogs.push(`${woredaLabel}: Early ANC (${eAnc}) > ANC 1 (${a1})`);
    }
    if (!isNaN(a1) && !isNaN(a4) && a4 > a1) {
      consistencyLogs.push(`${woredaLabel}: ANC 4 (${a4}) > ANC 1 (${a1})`);
    }
    if (!isNaN(a1) && !isNaN(a8) && a8 > a1) {
      consistencyLogs.push(`${woredaLabel}: ANC 8 (${a8}) > ANC 1 (${a1})`);
    }
    if (!isNaN(a4) && !isNaN(a8) && a8 > a4) {
      consistencyLogs.push(`${woredaLabel}: ANC 8 (${a8}) > ANC 4 (${a4})`);
    }

    // Check 4: Same Value Across Multiple Weeks (3+ consecutive weeks)
    if (priorWeeks.length >= 2) {
      indicatorPairs.forEach(({ performedKey }) => {
        const currVal = parseFloat(row[performedKey]);
        if (isNaN(currVal) || currVal === 0) return;
        const prevVals = priorWeeks.map(p => parseFloat(p[performedKey])).filter(v => !isNaN(v));
        if (prevVals.length >= 2 && prevVals.slice(-2).every(v => v === currVal)) {
          repeatedWeeksLogs.push(`${woredaLabel}: Repeated value (${currVal}) in [${performedKey}] across 3+ weeks`);
        }
      });
    }

    // Check 5: Same Value Horizontally Across Different Categories/Indicators in the Same Week
    const valueGroups = {};
    indicatorPairs.forEach(({ catKey, performedKey }) => {
      const val = parseFloat(row[performedKey]);
      if (!isNaN(val) && val > 0) {
        if (!valueGroups[val]) valueGroups[val] = [];
        valueGroups[val].push({ catKey, performedKey });
      }
    });
    Object.entries(valueGroups).forEach(([val, items]) => {
      const distinctCategories = new Set(items.map(i => i.catKey));
      if (items.length >= 3 && distinctCategories.size >= 2) {
        const names = items.map(i => i.performedKey).join(', ');
        horizontalDupLogs.push(`${woredaLabel}: Identical count (${val}) repeated horizontally across [${names}]`);
      }
    });

    // Check 6: Skilled Birth Attendance > Hepatitis BD Discrepancy
    const sbaVal = getNum(sbaCol);
    const hepVal = getNum(hepBdCol);
    if (!isNaN(sbaVal) && !isNaN(hepVal) && sbaVal > hepVal) {
      const diff = sbaVal - hepVal;
      sbaHepDiscrepancyLogs.push(`${woredaLabel}: Skilled Birth Attendance (${sbaVal}) > Hepatitis BD (${hepVal}) — Difference: +${diff}`);
    }
  });

  // 4. WOREDA CATEGORIZATION BY DISTANCE FROM REGIONAL TARGETS
  const resolvedTargets = REGIONAL_TARGETS.map(t => {
    const pair = indicatorPairs.find(p => t.patterns.some(pat => p.performedKey.toLowerCase().includes(pat)));
    return pair ? { ...t, performedKey: pair.performedKey, eligibleKey: pair.eligibleKey } : null;
  }).filter(Boolean);

  const woredaTargetEvaluations = [];
  reportingWoredas.forEach(row => {
    let totalDistance = 0;
    let evaluatedCount = 0;
    const indicatorGaps = [];

    resolvedTargets.forEach(({ label, target, performedKey, eligibleKey }) => {
      const perf = parseFloat(row[performedKey]);
      const elig = parseFloat(row[eligibleKey]);
      if (!isNaN(perf) && !isNaN(elig) && elig > 0) {
        const actualPct = (perf / elig) * 100;
        const distance = actualPct - target; // Positive = exceeded target, Negative = below target
        totalDistance += distance;
        evaluatedCount++;
        indicatorGaps.push({ label, actualPct, target, distance });
      }
    });

    if (evaluatedCount > 0) {
      const avgDistance = totalDistance / evaluatedCount;
      const woredaLabel = isGrandTotal ? `${row.Woreda} (${row.Zone})` : row.Woreda;
      // Sort indicator gaps to highlight strongest and furthest behind target
      indicatorGaps.sort((a, b) => a.distance - b.distance);
      const furthestBelow = indicatorGaps
        .filter(g => g.distance < 0)
        .slice(0, 2)
        .map(g => `${g.label}: ${g.actualPct.toFixed(1)}% vs ${g.target}% target (${g.distance.toFixed(1)}%)`)
        .join('; ');

      let tier = 'Below Target (Needs Close Follow-Up)';
      if (avgDistance >= 0) tier = 'Meeting / Exceeding Regional Targets';
      else if (avgDistance >= -15) tier = 'Near Target (Within 15% of Regional Target)';

      woredaTargetEvaluations.push({
        woreda: woredaLabel,
        avgDistance: Number(avgDistance.toFixed(1)),
        tier,
        furthestBelow: furthestBelow || 'All evaluated indicators met or exceeded target'
      });
    }
  });

  woredaTargetEvaluations.sort((a, b) => b.avgDistance - a.avgDistance);

  const targetCategorizationText = woredaTargetEvaluations.length > 0
    ? woredaTargetEvaluations.map(w => {
        const sign = w.avgDistance >= 0 ? '+' : '';
        return `- ${w.woreda} | Tier: ${w.tier} | Avg Distance from Target: ${sign}${w.avgDistance}% | Key Gaps: ${w.furthestBelow}`;
      }).join('\n')
    : '- No woredas with valid denominator data to evaluate against regional targets.';

  return {
    zoneName: selectedZone || 'Grand Total',
    currentWeek: effectiveWeek,
    reportingRate,
    nonReportingWoredas: nonReportingWoredas.join(', ') || 'None',
    execSummaryLines,
    categorizedSummaryText: categorizedSummaryText.join('\n'),
    qualityAuditText: `
1. Missing Values: ${missingDataLogs.join(' | ') || 'None'}
2. Historical Outliers (>= 3 SD from previous weeks): ${sdOutlierLogs.join(' | ') || 'None'}
3. Indicator Logical Consistency Issues (Penta 1/3, MCV 1/2, Early ANC/ANC 1/4/8): ${consistencyLogs.join(' | ') || 'None'}
4. Same Value Repeated Across Multiple Weeks: ${repeatedWeeksLogs.join(' | ') || 'None'}
5. Same Value Horizontally Across Different Categories: ${horizontalDupLogs.join(' | ') || 'None'}
6. Skilled Birth Attendance vs Hepatitis BD Discrepancy (SBA > Hep BD): ${sbaHepDiscrepancyLogs.join(' | ') || 'None'}
    `.trim(),
    targetCategorizationText
  };
}

app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'index.html')));
app.get('/health', (req, res) => res.status(200).send('OK - Server is awake'));

app.get('/api/filters', async (req, res) => {
  try {
    const rawData = await fetchFinalSheetData();
    // Trigger immediate maternal death alert check whenever data is refreshed
    checkAndSendMaternalDeathAlerts(rawData).catch(() => {});
    const zones = [...new Set(rawData.map(r => r.Zone).filter(Boolean))];
    const weeks = [...new Set(rawData.map(r => r.Week).filter(Boolean))];
    res.json({ zones, weeks });
  } catch (err) {
    console.error('Filter Error:', err);
    res.status(500).json({ zones: [], weeks: [] });
  }
});

async function generateAiReport(analytics) {
  const {
    zoneName,
    currentWeek,
    reportingRate,
    nonReportingWoredas,
    execSummaryLines,
    categorizedSummaryText,
    qualityAuditText,
    targetCategorizationText
  } = analytics;

  const preCalculatedBrief = `
SCOPE: ${zoneName} | WEEK: ${currentWeek}
REPORTING RATE: ${reportingRate}
NON-REPORTING WOREDAS: ${nonReportingWoredas}

[DATA FOR SECTION 1: SELECTED EXECUTIVE SUMMARY INDICATORS]
${execSummaryLines}

[DATA FOR SECTION 2: CATEGORY PERFORMANCE VS REGIONAL GRAND TOTAL]
${categorizedSummaryText}

[DATA FOR SECTION 3: 6-POINT DATA QUALITY & CONSISTENCY AUDIT]
${qualityAuditText}

[DATA FOR SECTION 4: WOREDA CATEGORIZATION BY DISTANCE FROM REGIONAL TARGETS]
Regional Targets Reference: CAR 96%, LAFP 48%, Early ANC 57%, ANC 4 98%, ANC 8 66%, SBA 89%, Hepatitis BD 87%, Penta 1 100%, Penta 3 100%.
${targetCategorizationText}
  `.trim();

  const systemPrompt = `You are an expert RMNCH Health Analytics & Data Auditor. All calculations, standard deviations, consistency checks, and target distance rankings have already been computed in code.
You MUST generate a strictly uniform report that follows the EXACT same 4-section template, headings, and flow for EVERY zone and for Grand Total. Do not alter the section titles or order.

Start your response with exactly:
Hello, This Is Weekly Essential RMNCH Service Report Feedback for ${zoneName}, ${currentWeek}

Then follow this exact 4-section structure:

## 1. Executive Summary and Reporting Overview
- State the reporting rate (${reportingRate}) and list non-reporting woredas (${nonReportingWoredas}).
- Present a clean summary of the 10 mandatory executive indicators across all 5 categories (Contraceptive Acceptance Rate, Maternal Death, Early ANC, ANC 8, Penta 1, Penta 3, MCV 1, Pneumonia Treatment, Neonatal Death, and Nutritional Screening) using the exact pre-calculated figures.

## 2. Category-by-Category Performance Analysis (vs. Regional Grand Total)
For each of the 5 categories (Category 01: SRH, Category 02: Maternal Health, Category 03: EPI, Category 04: Child Health, Category 05: Nutrition):
- Render a Markdown table with columns: | Indicator | Performed | Eligible | Coverage (%) | Regional Grand Total (%) | Difference |
- Immediately below each category table, include a bold **Key Takeaway:** summarizing performance relative to the Regional Grand Total.

## 3. Data Quality and Consistency Assessment
List the exact findings and woredas under each of the 6 numbered audit criteria (state "None" if no issues were found):
1. **Missing Values**
2. **Historical Outliers (>= 3 SD from Previous Weeks)**
3. **Consistency Between Related Indicators (Penta 1 vs 3, MCV 1 vs 2, Early ANC vs ANC 1, ANC 1/4/8)**
4. **Same Value Across Multiple Weeks**
5. **Same Value Horizontally Across Different Categories**
6. **Skilled Birth Attendance vs. Hepatitis BD Discrepancy (SBA > Hepatitis BD)**

## 4. Woreda Performance Categorization (Distance from Regional Targets)
- Group and list all evaluated woredas into their performance tiers (**Meeting / Exceeding Regional Targets**, **Near Target (Within 15%)**, and **Below Target (Needs Close Follow-Up)**) showing each woreda's exact average distance from the regional targets (CAR 96%, LAFP 48%, Early ANC 57%, ANC 4 98%, ANC 8 66%, SBA 89%, Hepatitis BD 87%, Penta 1 100%, Penta 3 100%) and key indicator gaps.`;

  const chatCompletion = await groq.chat.completions.create({
    messages: [
      { role: 'system', content: systemPrompt },
      { role: 'user', content: preCalculatedBrief }
    ],
    model: 'openai/gpt-oss-20b',
    temperature: 0.15,
    max_completion_tokens: 3000
  });

  const body = chatCompletion.choices[0]?.message?.content || 'No analysis generated.';
  return `${body.trim()}${REPORT_FOOTER}`;
}

app.post('/api/chat', async (req, res) => {
  try {
    const { selectedZone, selectedWeek } = req.body;
    const rawData = await fetchFinalSheetData();
    if (!rawData.length) return res.status(500).json({ reply: 'Could not retrieve data from Google Sheet.' });

    // Immediate maternal death alert check on live query
    await checkAndSendMaternalDeathAlerts(rawData, selectedWeek);

    const analytics = buildAnalyticsPayload(rawData, selectedZone || 'Grand Total', selectedWeek || 'Current Week');
    const responseText = await generateAiReport(analytics);
    res.json({ reply: responseText });
  } catch (error) {
    console.error('Groq/Sheet Error Details:', error);
    res.status(500).json({ reply: 'Error generating report. Please try again.' });
  }
});

app.post('/api/export-docx', async (req, res) => {
  try {
    const { reportText, title } = req.body;
    if (typeof reportText !== 'string' || !reportText.trim()) {
      return res.status(400).json({ error: 'Report text is required.' });
    }

    const lines = reportText.split('\n');
    const docChildren = [new Paragraph({
      text: title || 'Weekly Essential RMNCH Service Report Feedback',
      heading: HeadingLevel.HEADING_1,
      spacing: { after: 200 }
    })];

    let i = 0;
    while (i < lines.length) {
      const line = lines[i].trim();
      if (line.startsWith('|') && line.endsWith('|')) {
        const tableRows = [];
        while (i < lines.length && lines[i].trim().startsWith('|') && lines[i].trim().endsWith('|')) {
          const rowLine = lines[i].trim();
          if (!/^\|?\s*:?-{3,}/.test(rowLine.replace(/\|/g, '').trim())) {
            const cells = rowLine.split('|').slice(1, -1).map(c => c.trim());
            tableRows.push(cells);
          }
          i++;
        }
        if (tableRows.length) {
          const columnCount = Math.max(...tableRows.map(row => row.length));
          const docTableRows = tableRows.map((rowCells, rowIndex) => {
            const isHeader = rowIndex === 0;
            return new TableRow({
              children: Array.from({ length: columnCount }, (_, colIndex) => {
                const cellText = rowCells[colIndex] || '';
                return new TableCell({
                  children: [new Paragraph({
                    alignment: isHeader ? AlignmentType.CENTER : AlignmentType.LEFT,
                    children: [new TextRun({
                      text: cellText, bold: isHeader,
                      color: isHeader ? 'FFFFFF' : '000000', size: 18
                    })]
                  })],
                  width: { size: Math.floor(10000 / columnCount), type: WidthType.DXA },
                  shading: {
                    fill: isHeader ? '1F4E78' : (rowIndex % 2 === 0 ? 'F2F4F8' : 'FFFFFF'),
                    val: ShadingType.CLEAR, color: 'auto'
                  }
                });
              })
            });
          });
          docChildren.push(new Table({
            rows: docTableRows,
            width: { size: 100, type: WidthType.PERCENTAGE }
          }));
          docChildren.push(new Paragraph({ text: '', spacing: { after: 150 } }));
        }
      } else {
        if (line && line !== '---') {
          const isHeading = line.startsWith('## ') || line.startsWith('### ');
          const cleanText = lines[i].replace(/^#{1,3}\s+/, '').replace(/\*\*/g, '').replace(/\*/g, '');
          docChildren.push(new Paragraph({
            heading: isHeading ? HeadingLevel.HEADING_2 : undefined,
            children: [new TextRun({ text: cleanText, bold: isHeading, size: isHeading ? 26 : 22 })],
            spacing: { before: isHeading ? 160 : 0, after: 100 }
          }));
        }
        i++;
      }
    }

    const doc = new Document({ sections: [{ properties: {}, children: docChildren }] });
    const buffer = await Packer.toBuffer(doc);
    const token = crypto.randomUUID();
    const filename = `RMNCH_Service_Report_${Date.now()}.docx`;

    temporaryDocuments.set(token, {
      buffer,
      filename,
      expiresAt: Date.now() + DOWNLOAD_EXPIRY
    });

    const baseUrl = (process.env.APP_BASE_URL || `${req.protocol}://${req.get('host')}`).replace(/\/$/, '');
    const url = `${baseUrl}/api/download-docx/${token}`;
    res.json({ success: true, url, filename });
  } catch (err) {
    console.error('Docx Export Error:', err);
    res.status(500).json({ error: 'Failed to generate Word document.' });
  }
});

app.get('/api/download-docx/:token', (req, res) => {
  const token = req.params.token;
  const file = temporaryDocuments.get(token);
  if (!file || Date.now() >= file.expiresAt) {
    temporaryDocuments.delete(token);
    return res.status(404).send('This download link has expired. Please generate the report again.');
  }

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
  res.setHeader('Content-Disposition', `attachment; filename="${file.filename}"`);
  res.setHeader('Content-Length', file.buffer.length);
  res.setHeader('Cache-Control', 'no-store');
  return res.send(file.buffer);
});

async function sendViaBrevo(toInput, subject, textBody) {
  const recipients = Array.isArray(toInput) ? toInput : [toInput];
  const toList = recipients.map(email => ({ email: email.trim() }));
  const response = await fetch('https://api.brevo.com/v3/smtp/email', {
    method: 'POST',
    headers: {
      accept: 'application/json',
      'api-key': BREVO_API_KEY,
      'content-type': 'application/json'
    },
    body: JSON.stringify({
      sender: { name: 'Dr. Esubalew Zerihun', email: 'esubalewzerihun11@gmail.com' },
      to: toList, subject, textContent: textBody
    })
  });
  if (!response.ok) {
    const errData = await response.json();
    throw new Error(`Brevo API Error: ${JSON.stringify(errData)}`);
  }
  return response.json();
}

async function runRmnchEmailDispatch(forcedDay = null) {
  console.log('[RMNCH Email Dispatch] Starting execution via Brevo...');
  try {
    const rawData = await fetchFinalSheetData();
    if (!rawData?.length) return;
    const currentWeek = getCurrentReportingWeek(rawData);

    // Always verify and send immediate maternal death alerts first
    await checkAndSendMaternalDeathAlerts(rawData, currentWeek);

    const zones = [...new Set(rawData.map(r => r.Zone).filter(Boolean))];
    const dayOfWeek = forcedDay !== null ? forcedDay : new Date().getDay();

    // 1. ZONAL DISPATCHES
    for (const zone of zones) {
      const cleanZoneKey = zone.toString().trim().toLowerCase();
      const recipientEmail = ZONAL_EMAILS[cleanZoneKey];
      if (!recipientEmail) continue;

      const zoneData = rawData.filter(r =>
        r.Zone?.toString().trim().toLowerCase() === cleanZoneKey &&
        r.Week?.toString().trim().toLowerCase().includes(currentWeek.toLowerCase())
      );
      const totalExpected = zoneData.length;
      const reported = zoneData.filter(r => r.Report && r.Report.toLowerCase() === 'reported');
      const reportingRateStr = totalExpected > 0
        ? `${((reported.length / totalExpected) * 100).toFixed(1)}%` : '0%';
      const nonReporting = zoneData
        .filter(r => !r.Report || r.Report.toLowerCase() !== 'reported')
        .map(r => r.Woreda).join(', ') || 'None';

      let zonalBody = '';
      let subjectLine = '';

      if (dayOfWeek === 1) {
        subjectLine = `[Weekly Status Update] ${zone} - ${currentWeek}`;
        zonalBody = `Dear ${zone} MCH Team,\n\n` +
          `Greetings!\n\n` +
          `Please be reminded that the weekly RMNCH service report is expected from today until Wednesday at 6:30.\n\n` +
          `Current Reporting Rate: ${reportingRateStr} (${reported.length}/${totalExpected} Woredas reported)\n` +
          `Current Non-Reporting Woredas: ${nonReporting}\n\n` +
          `Kindly submit your reports before the deadline.` +
          REPORT_FOOTER;
      } else if (dayOfWeek === 2) {
        subjectLine = `[Weekly Status Update - Reminder] ${zone} - ${currentWeek}`;
        zonalBody = `Dear ${zone} MCH Team,\n\n` +
          `Greetings!\n\n` +
          `This is your daily update regarding the weekly RMNCH service reporting for ${currentWeek} (Deadline: Wednesday 6:30).\n\n` +
          `Current Reporting Rate: ${reportingRateStr} (${reported.length}/${totalExpected} Woredas reported)\n` +
          `Remaining Non-Reporting Woredas: ${nonReporting}\n\n` +
          `Please ensure timely submission.` +
          REPORT_FOOTER;
      } else if (dayOfWeek === 3) {
        subjectLine = `[Weekly Standardized Feedback & Analysis] ${zone} - ${currentWeek}`;
        const analytics = buildAnalyticsPayload(rawData, zone, currentWeek);
        const standardizedReport = await generateAiReport(analytics);

        zonalBody = `Dear ${zone} MCH Team,\n\n` +
          `Greetings!\n\n` +
          `The reporting window for ${currentWeek} has closed. Below is the list of non-reporting woredas followed by your standardized weekly feedback report:\n\n` +
          `Non-Reporting Woredas: ${nonReporting}\n\n` +
          `---------------------------------------------------\n\n` +
          standardizedReport;
      }

      if (zonalBody) {
        await sendViaBrevo(recipientEmail, subjectLine, zonalBody);
        await new Promise(resolve => setTimeout(resolve, 2500));
      }
    }

    // 2. REGIONAL LEADERSHIP DISPATCH (Wednesday Only)
    if (dayOfWeek === 3) {
      const regionalAnalytics = buildAnalyticsPayload(rawData, 'Grand Total', currentWeek);
      const regionalStandardizedReport = await generateAiReport(regionalAnalytics);

      const regionalBody = `Dear Regional Leadership Team,\n\n` +
        `Formal Greetings,\n\n` +
        `Please find below the standardized Regional Grand Total RMNCH feedback report for ${currentWeek}:\n\n` +
        `---------------------------------------------------\n\n` +
        regionalStandardizedReport;

      await sendViaBrevo(
        REGIONAL_EMAILS,
        `[Regional Grand Total Standardized Report] RMNCH Feedback - ${currentWeek}`,
        regionalBody
      );
    }

    return true;
  } catch (err) {
    console.error('Error in automated email dispatch via Brevo:', err);
    throw err;
  }
}

app.get('/api/test-email', async (req, res) => {
  try {
    const forcedDay = req.query.day ? parseInt(req.query.day, 10) : null;
    await runRmnchEmailDispatch(forcedDay);
    res.send(`RMNCH email test executed successfully via Brevo (Forced Day: ${forcedDay !== null ? forcedDay : 'Today'})!`);
  } catch (err) {
    res.status(500).send('Error executing email dispatch: ' + err.message);
  }
});

// Scheduled Mon/Tue/Wed 10:30 dispatch
cron.schedule('30 10 * * 1,2,3', async () => {
  await runRmnchEmailDispatch();
});

// Immediate maternal death surveillance poll every 15 minutes
cron.schedule('*/15 * * * *', async () => {
  try {
    const rawData = await fetchFinalSheetData();
    await checkAndSendMaternalDeathAlerts(rawData);
  } catch (err) {
    console.error('[Maternal Death Cron Error]:', err.message);
  }
});

app.listen(PORT, '0.0.0.0', () => {
  console.log(`Server running on port ${PORT}`);
});
