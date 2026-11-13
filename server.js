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

// Temporary in-memory storage for generated Word documents (15-min expiry)
const temporaryDocuments = new Map();
const DOWNLOAD_EXPIRY = 15 * 60 * 1000;

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

/**
 * PURE CODE CALCULATION ENGINE:
 * Computes all metrics, aggregations, rankings, and anomaly detection entirely in JS.
 */
function buildAnalyticsPayload(rawData, selectedZone, selectedWeek) {
  const isGrandTotal = !selectedZone || selectedZone === 'Grand Total';

  const zoneAllData = isGrandTotal
    ? rawData
    : rawData.filter(r => r.Zone && r.Zone.trim().toLowerCase() === selectedZone.trim().toLowerCase());

  let filteredData = zoneAllData;
  if (selectedWeek && selectedWeek !== 'All Weeks') {
    filteredData = filteredData.filter(r => r.Week && r.Week.toLowerCase().includes(selectedWeek.toLowerCase()));
  }

  const totalExpectedWoredas = filteredData.length;
  const reportingWoredas = filteredData.filter(r => r.Report && r.Report.toLowerCase() === 'reported');
  const nonReportingRows = filteredData.filter(r => !r.Report || r.Report.toLowerCase() !== 'reported');
  const nonReportingWoredas = nonReportingRows.map(r => isGrandTotal ? `${r.Woreda} (${r.Zone})` : r.Woreda).filter(Boolean);
  const reportingRate = totalExpectedWoredas > 0
    ? `${((reportingWoredas.length / totalExpectedWoredas) * 100).toFixed(1)}% (${reportingWoredas.length}/${totalExpectedWoredas})` : '0%';

  // Multi-week trend summary (Excluding 0% weeks)
  const weeklyTrends = {};
  zoneAllData.forEach(r => {
    const wk = r.Week || 'Unknown Week';
    if (!weeklyTrends[wk]) weeklyTrends[wk] = { totalExpected: 0, reportedCount: 0 };
    weeklyTrends[wk].totalExpected++;
    if (r.Report && r.Report.toLowerCase() === 'reported') weeklyTrends[wk].reportedCount++;
  });
  
  const weeklyTrendSummary = Object.keys(weeklyTrends).map(wk => {
    const total = weeklyTrends[wk].totalExpected;
    const reported = weeklyTrends[wk].reportedCount;
    const rate = total > 0 ? (reported / total) * 100 : 0;
    return { wk, rate, reported, total };
  })
  .filter(item => item.rate > 0)
  .map(item => `${item.wk}: ${item.rate.toFixed(1)}% (${item.reported}/${item.total})`)
  .join(' | ') || 'No active reporting history available.';

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

  // Calculate Indicator Summaries per Category
  const categorizedSummaryText = [];
  const categoryLabels = {
    cat1_srh: 'Category 1: SRH Services',
    cat2_maternal: 'Category 2: Maternal Services',
    cat3_epi: 'Category 3: EPI Services',
    cat4_child: 'Category 4: Child Health Services',
    cat5_nutrition: 'Category 5: Nutrition'
  };

  const categoryRankingsText = [];

  Object.keys(categoryLabels).forEach(catKey => {
    const pairs = indicatorPairs.filter(p => p.catKey === catKey);
    const catIndicators = [];
    const woredaScores = [];

    pairs.forEach(({ baseName, performedKey, eligibleKey }) => {
      let totalPerf = 0;
      let totalElig = 0;
      filteredData.forEach(row => {
        totalPerf += parseFloat(row[performedKey]) || 0;
        totalElig += parseFloat(row[eligibleKey]) || 0;
      });
      const cov = totalElig > 0 ? ((totalPerf / totalElig) * 100).toFixed(1) : '0.0';
      catIndicators.push(`- ${baseName}: Performed=${totalPerf}, Eligible=${totalElig}, Coverage=${cov}%`);
    });

    reportingWoredas.forEach(row => {
      let sumPct = 0;
      let validCount = 0;
      pairs.forEach(({ performedKey, eligibleKey }) => {
        const elig = parseFloat(row[eligibleKey]);
        const perf = parseFloat(row[performedKey]);
        if (!isNaN(elig) && elig > 0 && !isNaN(perf)) {
          sumPct += (perf / elig) * 100;
          validCount++;
        }
      });
      if (validCount > 0) {
        const label = isGrandTotal ? `${row.Woreda} (${row.Zone})` : row.Woreda;
        woredaScores.push({ woreda: label, avgCoverage: Number((sumPct / validCount).toFixed(1)) });
      }
    });

    woredaScores.sort((a, b) => b.avgCoverage - a.avgCoverage);
    const best = woredaScores.slice(0, 3).map(w => `${w.woreda} (${w.avgCoverage}%)`).join(', ') || 'None';
    const poor = woredaScores.length > 0 ? woredaScores.slice(-3).reverse().map(w => `${w.woreda} (${w.avgCoverage}%)`).join(', ') : 'None';

    categorizedSummaryText.push(`\n**${categoryLabels[catKey]}**\n${catIndicators.join('\n')}`);
    categoryRankingsText.push(`- ${categoryLabels[catKey]} -> Best 3: [${best}] | Poor 3: [${poor}]`);
  });

  // Data Quality & Consistency Audits (Pure JS Calculations)
  const missingDataLogs = [];
  const outlierLogs = [];
  const repeatedLogs = [];
  const inconsistencyLogs = [];

  const findCol = (patterns) => headers.find(h =>
    !h.includes('Eligible') && !h.includes('Target') &&
    patterns.some(p => h.toLowerCase().includes(p))
  );

  const penta1Col = findCol(['penta 1', 'penta-1', 'pentavalent 1', 'pentavalent-1', 'penta1']);
  const penta3Col = findCol(['penta 3', 'penta-3', 'pentavalent 3', 'pentavalent-3', 'penta3']);
  const anc1Col = findCol(['anc 1', 'anc-1', 'anc1']);
  const anc4Col = findCol(['anc 4', 'anc-4', 'anc4']);
  const anc8Col = findCol(['anc 8', 'anc-8', 'anc8']);

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

    // Missing Data
    indicatorPairs.forEach(({ performedKey }) => {
      const val = row[performedKey];
      if (val === undefined || val === null || val.toString().trim() === '') {
        missingDataLogs.push(`${woredaLabel} missing [${performedKey}]`);
      }
    });

    // Indicator Inconsistencies (Penta 3 > Penta 1, ANC 4 > ANC 1, ANC 8 > ANC 4)
    const p1 = penta1Col ? parseFloat(row[penta1Col]) : NaN;
    const p3 = penta3Col ? parseFloat(row[penta3Col]) : NaN;
    if (!isNaN(p1) && !isNaN(p3) && p3 > p1) {
      inconsistencyLogs.push(`${woredaLabel}: Penta 3 (${p3}) exceeds Penta 1 (${p1})`);
    }

    const a1 = anc1Col ? parseFloat(row[anc1Col]) : NaN;
    const a4 = anc4Col ? parseFloat(row[anc4Col]) : NaN;
    const a8 = anc8Col ? parseFloat(row[anc8Col]) : NaN;
    if (!isNaN(a1) && !isNaN(a4) && a4 > a1) {
      inconsistencyLogs.push(`${woredaLabel}: ANC 4 (${a4}) exceeds ANC 1 (${a1})`);
    }
    if (!isNaN(a4) && !isNaN(a8) && a8 > a4) {
      inconsistencyLogs.push(`${woredaLabel}: ANC 8 (${a8}) exceeds ANC 4 (${a4})`);
    }

    // Historical Outliers & Repeated Values
    const priorWeeks = woredaHistory.filter(h => h.Report?.toLowerCase() === 'reported' && h.Week !== row.Week);
    if (priorWeeks.length >= 2) {
      indicatorPairs.forEach(({ performedKey }) => {
        const currVal = parseFloat(row[performedKey]);
        if (isNaN(currVal)) return;
        const prevVals = priorWeeks.map(p => parseFloat(p[performedKey])).filter(v => !isNaN(v));

        if (prevVals.length >= 2) {
          const recentTwo = prevVals.slice(-2);
          if (currVal > 0 && recentTwo.every(v => v === currVal)) {
            repeatedLogs.push(`${woredaLabel}: Reported identical figure (${currVal}) for [${performedKey}] across 3+ weeks`);
          }
          const mean = prevVals.reduce((a, b) => a + b, 0) / prevVals.length;
          if (mean >= 10 && Math.abs(currVal - mean) / mean > 0.5) {
            const dir = currVal > mean ? 'spike' : 'drop';
            outlierLogs.push(`${woredaLabel}: Unusual ${dir} in [${performedKey}] (Current: ${currVal} vs Historical Avg: ${mean.toFixed(1)})`);
          }
        }
      });
    }
  });

  // Mortality / Death Notifications
  const mortalityLogs = [];
  filteredData.forEach(row => {
    const woredaName = isGrandTotal ? `${row.Woreda} (${row.Zone})` : (row.Woreda || 'Unknown Woreda');
    headers.forEach(h => {
      const lowerH = h.toLowerCase();
      if ((lowerH.includes('maternal death') || lowerH.includes('neonatal death') || h.includes('2.8') || h.includes('4.4')) && parseFloat(row[h]) > 0) {
        mortalityLogs.push(`${woredaName}: ${row[h]} case(s) under [${h}]`);
      }
    });
  });

  return {
    zoneName: selectedZone || 'Grand Total',
    currentWeek: selectedWeek || 'Current Week',
    reportingRate,
    nonReportingWoredas: nonReportingWoredas.join(', ') || 'None',
    weeklyTrendSummary,
    categorizedSummaryText: categorizedSummaryText.join('\n'),
    categoryRankingsText: categoryRankingsText.join('\n'),
    qualityAuditText: `
- Missing Data Issues: ${missingDataLogs.slice(0, 5).join(' | ') || 'None'}
- Statistical Outliers vs Past Weeks: ${outlierLogs.slice(0, 5).join(' | ') || 'None'}
- Repeated Figures Across Weeks: ${repeatedLogs.slice(0, 5).join(' | ') || 'None'}
- Indicator Logic Inconsistencies: ${inconsistencyLogs.slice(0, 5).join(' | ') || 'None'}
    `.trim(),
    mortalityText: mortalityLogs.join(' | ') || 'No maternal or neonatal deaths reported.'
  };
}

app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'index.html')));
app.get('/health', (req, res) => res.status(200).send('OK - Server is awake'));

app.get('/api/filters', async (req, res) => {
  try {
    const rawData = await fetchFinalSheetData();
    const zones = [...new Set(rawData.map(r => r.Zone).filter(Boolean))];
    const weeks = [...new Set(rawData.map(r => r.Week).filter(Boolean))];
    res.json({ zones, weeks });
  } catch (err) {
    console.error('Filter Error:', err);
    res.status(500).json({ zones: [], weeks: [] });
  }
});

async function generateAiReport(analytics, isEmailSummary = false) {
  const {
    zoneName,
    currentWeek,
    reportingRate,
    nonReportingWoredas,
    weeklyTrendSummary,
    categorizedSummaryText,
    categoryRankingsText,
    qualityAuditText,
    mortalityText
  } = analytics;

  const preCalculatedBrief = `
- Zone: ${zoneName}
- Week: ${currentWeek}
- Reporting Rate: ${reportingRate}
- Non-Reporting Woredas: ${nonReportingWoredas}
- Multi-Week Trends: ${weeklyTrendSummary}
- Indicator Coverage Data:
${categorizedSummaryText}
- Category Rankings (Best & Poor 3 Woredas):
${categoryRankingsText}
- Data Quality & Consistency Audit Results:
${qualityAuditText}
- Death Notifications / Mortality Alerts:
${mortalityText}
  `.trim();

  const formatInstruction = isEmailSummary
    ? `CRITICAL FORMAT: Write a short, executive summary email (under 250 words) using clean bullet points based on the pre-calculated data above. Include Reporting Overview, Service Coverage highlights, Top/Poor 3 woredas, Quality/Consistency alerts, and Death Notifications.`
    : `CRITICAL FORMAT: Write the complete comprehensive 9-section report using the pre-calculated data above. Include professional tables for each category and detailed narrative sections for Quality Audits and Death Notifications.`;

  const systemPrompt = `You are an expert RMNCH Health Analytics & Data Auditor. Your job is strictly linguistic narration and formatting—all calculations have already been performed in code. 
Start your response with exactly:
"Hello, This Is Weekly Essential RMNCH Service Report Feedback for ${zoneName}, ${currentWeek}"

${formatInstruction}`;

  const chatCompletion = await groq.chat.completions.create({
    messages: [
      { role: 'system', content: systemPrompt },
      { role: 'user', content: preCalculatedBrief }
    ],
    model: 'openai/gpt-oss-20b',
    temperature: 0.2,
    max_completion_tokens: isEmailSummary ? 600 : 2500
  });

  const body = chatCompletion.choices[0]?.message?.content || 'No analysis generated.';
  return `${body.trim()}${REPORT_FOOTER}`;
}

app.post('/api/chat', async (req, res) => {
  try {
    const { selectedZone, selectedWeek } = req.body;
    const rawData = await fetchFinalSheetData();
    if (!rawData.length) return res.status(500).json({ reply: 'Could not retrieve data from Google Sheet.' });

    const analytics = buildAnalyticsPayload(rawData, selectedZone || 'Grand Total', selectedWeek || 'Current Week');
    const responseText = await generateAiReport(analytics, false);
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
          docChildren.push(new Paragraph({
            children: [new TextRun({ text: lines[i].replace(/\*\*/g, '').replace(/\*/g, ''), size: 22 })],
            spacing: { after: 100 }
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
  const toList = recipients.map(email => ({ email }));
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
    const zones = [...new Set(rawData.map(r => r.Zone).filter(Boolean))];
    const dayOfWeek = forcedDay !== null ? forcedDay : new Date().getDay();

    const zonalEmails = {
      'kaffa zone': ['tadessebelay1221@gmail.com', 'esubalewzerihun37@gmail.com'],
      'dawro zone': ['ketesfu7@gmail.com', 'esubalewzerihun37@gmail.com'],
      'west omo zone': ['esubalewzerihun11@gmail.com', 'esubalewzerihun37@gmail.com'],
      'bench sheko zone': ['Joannagashahun13@gmail.com', 'habtekelbi78@gmail.com', 'ruphaelw@gmail.com'],
      'sheka zone': ['sifanmesfin@gmail.com', 'esubalewzerihun37@gmail.com'],
      'konta zone': ['tsegabebe2015@gmail.com', 'degemukidist365@gmail.com']
    };
    const regionalEmails = ['esubalewzerihun11@gmail.com', 'messiwode46@gmail.com'];

    for (const zone of zones) {
      const cleanZoneKey = zone.toString().trim().toLowerCase();
      const recipientEmail = zonalEmails[cleanZoneKey];
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
      if (dayOfWeek === 1 || dayOfWeek === 2) {
        zonalBody = `Hello,\n\nThis is your Weekly Essential RMNCH Service Status Update for ${zone} (${currentWeek}).\n\n` +
          `📊 Reporting Rate: ${reportingRateStr} (${reported.length}/${totalExpected} Woredas reported)\n` +
          `⚠️ Non-Reporting Woredas: ${nonReporting}\n\n` +
          'Note: Summary analytical feedback will be dispatched on Wednesday.' +
          REPORT_FOOTER;
      } else {
        const analytics = buildAnalyticsPayload(rawData, zone, currentWeek);
        zonalBody = await generateAiReport(analytics, true);
      }

      await sendViaBrevo(
        recipientEmail,
        `[${dayOfWeek === 1 || dayOfWeek === 2 ? 'Status Update' : 'Weekly Summary Feedback'}] ${zone} - ${currentWeek}`,
        zonalBody
      );

      // Throttling pause between zone emails to ensure absolute stability
      await new Promise(resolve => setTimeout(resolve, 2500));
    }

    if (dayOfWeek === 3) {
      const regionalAnalytics = buildAnalyticsPayload(rawData, 'Grand Total', currentWeek);
      const regionalBody = await generateAiReport(regionalAnalytics, true);

      await sendViaBrevo(
        regionalEmails,
        `[Weekly Summary Feedback] Regional Grand Total RMNCH Report - ${currentWeek}`,
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

app.listen(PORT, '0.0.0.0', () => {
  console.log(`Server running on port ${PORT}`);
});
