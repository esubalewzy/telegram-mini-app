import express from 'express';
import path from 'path';
import { fileURLToPath } from 'url';
import Groq from 'groq-sdk';
import cron from 'node-cron';
import { 
  Document, Packer, Paragraph, TextRun, HeadingLevel, 
  Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, AlignmentType 
} from 'docx';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
app.use(express.json());
app.use(express.static(__dirname));

const PORT = process.env.PORT || 3000;

// Initialize Groq & Brevo keys safely
const groq = new Groq({ apiKey: process.env.GROQ_API_KEY || '' });
const BREVO_API_KEY = process.env.BREVO_API_KEY || '';

const SHEET_ID = '1d8_5Q2W3aGV68W_DdjrPrFEJnuRSDprrb1b3kngbTpo';
const FINAL_SHEET_CSV_URL = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:csv&sheet=Final`;

function parseCSVLine(text) {
  let result = [];
  let current = '';
  let inQuotes = false;
   
  for (let i = 0; i < text.length; i++) {
    let char = text[i];
    if (char === '"') {
      inQuotes = !inQuotes;
    } else if (char === ',' && !inQuotes) {
      result.push(current.trim().replace(/^"(.*)"$/, '$1'));
      current = '';
    } else {
      current += char;
    }
  }
  result.push(current.trim().replace(/^"(.*)"$/, '$1'));
  return result;
}

async function fetchFinalSheetData() {
  try {
    const response = await fetch(FINAL_SHEET_CSV_URL);
    if (!response.ok) throw new Error(`Fetch failed: ${response.statusText}`);
    const csvText = await response.text();
     
    const lines = csvText.split('\n').filter(line => line.trim() !== '');
    if (lines.length === 0) return [];

    const headers = parseCSVLine(lines[0]);

    const rows = lines.slice(1).map(line => {
      const values = parseCSVLine(line);
      let obj = {};
      headers.forEach((h, index) => {
        obj[h] = values[index] !== undefined ? values[index] : '';
      });
      return obj;
    });

    return rows;
  } catch (err) {
    console.error('Sheet Fetch Error:', err);
    throw err;
  }
}

// Automatically determine active reporting week starting from Week 19 today and advancing weekly
function getCurrentReportingWeek(rawData) {
  const baseDate = new Date('2026-11-09'); // Week 19 start anchor
  const baseWeekNumber = 19;
   
  const currentDate = new Date();
  const diffTime = currentDate - baseDate;
  const diffDays = Math.floor(diffTime / (1000 * 60 * 60 * 24));
  const diffWeeks = Math.floor(diffDays / 7);
   
  const calculatedWeekNum = baseWeekNumber + Math.max(0, diffWeeks);
  const targetWeekQuery = `Week ${calculatedWeekNum}`;

  const availableWeeks = [...new Set(rawData.map(r => r.Week ? r.Week.toString().trim() : null).filter(Boolean))];
  const matchedWeek = availableWeeks.find(w => w.toLowerCase().includes(targetWeekQuery.toLowerCase()));

  if (matchedWeek) {
    return matchedWeek;
  }
  return availableWeeks[availableWeeks.length - 1];
}

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

// Health check endpoint for Render & UptimeRobot keep-alive
app.get('/health', (req, res) => {
  res.status(200).send('OK - Server is awake');
});

app.get('/api/filters', async (req, res) => {
  try {
    const rawData = await fetchFinalSheetData();
    const zones = [...new Set(rawData.map(r => r.Zone).filter(Boolean))];
    const weeks = [...new Set(rawData.map(r => r.Week).filter(Boolean))];
    res.json({ zones, weeks });
  } catch (err) {
    res.status(500).json({ zones: [], weeks: [] });
  }
});

// Helper to generate AI comprehensive analysis using Groq
async function generateAiReport(zoneName, filteredData, reportingRate, nonReportingWoredas, weeklyTrendSummary, categorizedIndicators, hstp2OfficialTargets, qualityIssues, hepBDMissedWoredas, mortalityAlerts, currentWeek) {
  const systemPrompt = `You are an expert RMNCH Health Analytics & Data Auditor. Generate a complete, exhaustive analytical report without truncating any categories.

Greeting Header (Must start with exactly):
"Hello, This Is Weekly Essential RMNCH Service Report Feedback for ${zoneName} Zone, Week ${currentWeek}"

Data Context:
- Reporting Rate: ${reportingRate}
- Non-Reporting Woredas: ${nonReportingWoredas.join(', ') || 'None'}
- Multi-Week Reporting Trend: ${weeklyTrendSummary}
- Complete Categorized Indicators: ${JSON.stringify(categorizedIndicators)}
- Official Ethiopian HSTP-II Targets: ${JSON.stringify(hstp2OfficialTargets)}
- Hepatitis BD Missed Audit: ${hepBDMissedWoredas.join(' | ') || 'None'}
- Consistency & Data Quality Issues: ${qualityIssues.join(' | ') || 'None'}
- Mortality Alerts (Maternal/Neonatal Only): ${mortalityAlerts.join(' | ') || 'No maternal or neonatal deaths reported.'}

Required Report Structure:
1. **Reporting Rate Summary**
2. **Multi-Week Reporting Trend Analysis**
3. **Category 1: SRH Services** (Table format: | Indicator | Performed | Eligible | Coverage % |)
4. **Category 2: Maternal Services** (Table format: | Indicator | Performed | Eligible | Coverage % |)
5. **Category 3: EPI Services** (Table format: | Indicator | Performed | Eligible | Coverage % |)
6. **Category 4: Child Health Service** (Table format: | Indicator | Performed | Eligible | Coverage % |)
7. **Category 5: Nutrition** (Table format: | Indicator | Performed | Eligible | Coverage % |)
8. **Hepatitis BD Birth Dose Missed Woredas Audit**
9. **Data Quality & Consistency Audit**
10. **Dedicated Ethiopian HSTP-II Strategic Benchmarking Evaluation**
11. **Mortality Alerts** (Strictly Maternal and Neonatal deaths only)`;

  const chatCompletion = await groq.chat.completions.create({
    messages: [
      { role: "system", content: systemPrompt },
      { role: "user", content: `Generate full individualized report for ${zoneName}.` }
    ],
    model: "openai/gpt-oss-20b",
    temperature: 0.3,
    max_completion_tokens: 3000,
  });

  return chatCompletion.choices[0]?.message?.content || 'No analysis generated.';
}

app.post('/api/chat', async (req, res) => {
  try {
    const { message, selectedZone, selectedWeek } = req.body;
    const rawData = await fetchFinalSheetData();
    if (rawData.length === 0) {
      return res.status(500).json({ reply: 'Could not retrieve data from Google Sheet.' });
    }

    let filteredData = rawData;
    if (selectedZone && selectedZone !== 'Grand Total') {
      filteredData = filteredData.filter(r => r.Zone && r.Zone.toLowerCase() === selectedZone.toLowerCase());
    }
    if (selectedWeek && selectedWeek !== 'All Weeks') {
      filteredData = filteredData.filter(r => r.Week && r.Week.toLowerCase().includes(selectedWeek.toLowerCase()));
    }

    const totalExpectedWoredas = filteredData.length;
    const reportingWoredas = filteredData.filter(r => r.Report && r.Report.toLowerCase() === 'reported');
    const nonReportingRows = filteredData.filter(r => !r.Report || r.Report.toLowerCase() !== 'reported');
    const nonReportingWoredas = nonReportingRows.map(r => r.Woreda).filter(Boolean);
    const reportingRate = totalExpectedWoredas > 0 ? ((reportingWoredas.length / totalExpectedWoredas) * 100).toFixed(1) + '%' : '0%';

    let weeklyTrends = {};
    let zoneAllData = selectedZone && selectedZone !== 'Grand Total' 
      ? rawData.filter(r => r.Zone && r.Zone.toLowerCase() === selectedZone.toLowerCase()) 
      : rawData;

    zoneAllData.forEach(r => {
      const wk = r.Week || 'Unknown Week';
      if (!weeklyTrends[wk]) weeklyTrends[wk] = { totalExpected: 0, reportedCount: 0 };
      weeklyTrends[wk].totalExpected++;
      if (r.Report && r.Report.toLowerCase() === 'reported') weeklyTrends[wk].reportedCount++;
    });

    let weeklyTrendSummary = Object.keys(weeklyTrends).map(wk => {
      const rate = ((weeklyTrends[wk].reportedCount / weeklyTrends[wk].totalExpected) * 100).toFixed(1);
      return `Week ${wk}: ${rate}% (${weeklyTrends[wk].reportedCount}/${weeklyTrends[wk].totalExpected} woredas)`;
    }).join(' | ');

    const headers = Object.keys(rawData[0] || {});
    let categorizedIndicators = { cat1_srh: [], cat2_maternal: [], cat3_epi: [], cat4_child: [], cat5_nutrition: [] };

    headers.forEach(header => {
      if (header.includes('Eligible')) {
        const baseName = header.replace('Eligible', '').trim();
        const performedKey = headers.find(h => h.startsWith(baseName) && !h.includes('Eligible') && !h.includes('Target'));
        if (performedKey) {
          let totalPerformed = 0;
          let totalEligible = 0;
          filteredData.forEach(row => {
            totalPerformed += parseFloat(row[performedKey]) || 0;
            totalEligible += parseFloat(row[header]) || 0;
          });
          const coverageNum = totalEligible > 0 ? parseFloat(((totalPerformed / totalEligible) * 100).toFixed(1)) : 0;
          const indicatorEntry = { name: performedKey, performed: totalPerformed, eligible: totalEligible, coverage: coverageNum + '%' };

          if (baseName.startsWith('1.')) categorizedIndicators.cat1_srh.push(indicatorEntry);
          else if (baseName.startsWith('2.')) categorizedIndicators.cat2_maternal.push(indicatorEntry);
          else if (baseName.startsWith('3.')) categorizedIndicators.cat3_epi.push(indicatorEntry);
          else if (baseName.startsWith('4.')) categorizedIndicators.cat4_child.push(indicatorEntry);
          else if (baseName.startsWith('5.')) categorizedIndicators.cat5_nutrition.push(indicatorEntry);
        }
      }
    });

    const hstp2OfficialTargets = {
      maternalAndReproductive: { "CPR": "50%", "ANC 4 Visits": "81%", "Skilled Delivery": "76%", "Early PNC": "76%" },
      childHealthAndEpi: { "Pentavalent-3": "85%", "Full Vaccination": "75%", "MCV2": "80%" }
    };

    const qualityIssues = [];
    const hepBDMissedWoredas = [];
    const mortalityAlerts = [];

    filteredData.forEach(row => {
      const woredaName = row['Woreda'] || 'Unknown Woreda';
      headers.forEach(h => {
        const lowerH = h.toLowerCase();
        if ((lowerH.includes('maternal death') || lowerH.includes('neonatal death') || h.includes('2.8') || h.includes('4.4')) && parseFloat(row[h]) > 0) {
          mortalityAlerts.push(`[CRITICAL MORTALITY ALERT] ${woredaName} reported ${row[h]} case(s) under [${h}].`);
        }
      });
    });

    const responseText = await generateAiReport(
      selectedZone || 'Grand Total', filteredData, reportingRate, nonReportingWoredas, 
      weeklyTrendSummary, categorizedIndicators, hstp2OfficialTargets, qualityIssues, hepBDMissedWoredas, mortalityAlerts, selectedWeek || 'Current Week'
    );

    res.json({ reply: responseText });
  } catch (error) {
    console.error('Groq/Sheet Error Details:', error);
    res.status(500).json({ reply: 'Error generating report. Please try again.' });
  }
});

app.post('/api/export-docx', async (req, res) => {
  try {
    const { reportText, title } = req.body;
    const lines = reportText.split('\n');
    const docChildren = [
      new Paragraph({
        text: title || "This Is Weekly Essential RMNCH Service Report Feedback",
        heading: HeadingLevel.HEADING_1,
        spacing: { after: 200 }
      })
    ];

    let i = 0;
    while (i < lines.length) {
      let line = lines[i].trim();
      if (line.startsWith('|') && line.endsWith('|')) {
        let tableRows = [];
        while (i < lines.length && lines[i].trim().startsWith('|') && lines[i].trim().endsWith('|')) {
          let rowLine = lines[i].trim();
          if (!rowLine.includes('---')) {
            let cells = rowLine.split('|').filter(c => c !== '').map(c => c.trim());
            tableRows.push(cells);
          }
          i++;
        }
        if (tableRows.length > 0) {
          const docTableRows = tableRows.map((rowCells, rowIndex) => {
            const isHeader = rowIndex === 0;
            return new TableRow({
              children: rowCells.map(cellText => {
                return new TableCell({
                  children: [new Paragraph({
                    alignment: isHeader ? AlignmentType.CENTER : AlignmentType.LEFT,
                    children: [new TextRun({ text: cellText, bold: isHeader, color: isHeader ? "FFFFFF" : "000000", size: 18 })]
                  })],
                  width: { size: 2000, type: WidthType.DXA },
                  shading: isHeader ? { fill: "1F4E78", val: ShadingType.CLEAR, color: "auto" } : { fill: rowIndex % 2 === 0 ? "F2F4F8" : "FFFFFF", val: ShadingType.CLEAR, color: "auto" }
                });
              })
            });
          });
          docChildren.push(new Table({ rows: docTableRows, width: { size: 100, type: WidthType.PERCENTAGE } }));
          docChildren.push(new Paragraph({ text: "", spacing: { after: 150 } }));
        }
      } else {
        if (line !== '') {
          docChildren.push(new Paragraph({ children: [new TextRun({ text: lines[i], size: 22 })], spacing: { after: 100 } }));
        }
        i++;
      }
    }

    const doc = new Document({ sections: [{ properties: {}, children: docChildren }] });
    const buffer = await Packer.toBuffer(doc);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', 'attachment; filename=RMNCH_Service_Report_Feedback.docx');
    res.send(buffer);
  } catch (err) {
    console.error('Docx Export Error:', err);
    res.status(500).json({ error: 'Failed to generate Word document.' });
  }
});

// Helper function to send email via Brevo REST API (fully supports arrays or single strings)
async function sendViaBrevo(toInput, subject, textBody) {
  const recipients = Array.isArray(toInput) ? toInput : [toInput];
  const toList = recipients.map(email => ({ email }));

  const response = await fetch('https://api.brevo.com/v3/smtp/email', {
    method: 'POST',
    headers: {
      'accept': 'application/json',
      'api-key': BREVO_API_KEY,
      'content-type': 'application/json'
    },
    body: JSON.stringify({
      sender: { name: "Dr. Esubalew Zerihun", email: "esubalewzerihun11@gmail.com" },
      to: toList,
      subject: subject,
      textContent: textBody
    })
  });

  if (!response.ok) {
    const errData = await response.json();
    throw new Error(`Brevo API Error: ${JSON.stringify(errData)}`);
  }
  return await response.json();
}

// Main email dispatch function with updated zone keys including "zone"
async function runRmnchEmailDispatch() {
  console.log(`[RMNCH Email Dispatch] Starting execution via Brevo...`);

  try {
    const rawData = await fetchFinalSheetData();
    if (!rawData || rawData.length === 0) {
      console.log('[Dispatch] No data found in Google Sheet.');
      return;
    }

    const currentWeek = getCurrentReportingWeek(rawData);
    console.log(`[Dispatch] Active reporting week targeted: "${currentWeek}"`);

    const zones = [...new Set(rawData.map(r => r.Zone).filter(Boolean))];
    const dayOfWeek = new Date().getDay(); // 1 = Mon, 2 = Tue, 3 = Wed

    // Updated keys to include "zone" to match Google Sheet names precisely
    const zonalEmails = {
      "kaffa zone": ["tadessebelay1221@gmail.com", "esubalewzerihun37@gmail.com"],
      "dawro zone": ["ketesfu7@gmail.com", "esubalewzerihun37@gmail.com"],
      "west omo zone": ["esubalewzerihun11@gmail.com", "esubalewzerihun37@gmail.com"],
      "bench sheko zone": ["Joannagashahun13@gmail.com", "habtekelbi78@gmail.com", "ruphaelw@gmail.com"],
      "sheka zone": ["sifanmesfin@gmail.com", "esubalewzerihun37@gmail.com"],
      "konta zone": ["tsegabebe2015@gmail.com", "esubalewzerihun37@gmail.com"]
    };

    // Regional leadership email(s) for Wednesday Grand Total
    const regionalEmails = ["esubalewzerihun11@gmail.com"];

    // 1. Loop through each zone and send strictly zone-filtered reports to zonal coordinators (Mon, Tue, Wed)
    for (const zone of zones) {
      const cleanZoneKey = zone.toString().trim().toLowerCase();
      const recipientEmail = zonalEmails[cleanZoneKey];

      if (!recipientEmail) {
        console.warn(`⚠️ [Dispatch] No email configuration found for zone: "${zone}". Skipping.`);
        continue;
      }

      const zoneData = rawData.filter(r => 
        r.Zone && r.Zone.toString().trim().toLowerCase() === cleanZoneKey &&
        r.Week && r.Week.toString().trim().toLowerCase().includes(currentWeek.toLowerCase())
      );

      const totalExpected = zoneData.length;
      const reported = zoneData.filter(r => r.Report && r.Report.toLowerCase() === 'reported');
      const reportingRateVal = totalExpected > 0 ? (reported.length / totalExpected) * 100 : 0;
      const reportingRateStr = reportingRateVal.toFixed(1) + '%';
      const nonReporting = zoneData.filter(r => !r.Report || r.Report.toLowerCase() !== 'reported').map(r => r.Woreda).join(', ') || 'None';

      let zonalBody = '';

      if (dayOfWeek === 1 || dayOfWeek === 2) {
        zonalBody = `Hello,\n\nThis is your Weekly Essential RMNCH Service Status Update for ${zone} (${currentWeek}).\n\n` +
                    `📊 Reporting Rate: ${reportingRateStr} (${reported.length}/${totalExpected} Woredas reported)\n` +
                    `⚠️ Non-Reporting Woredas: ${nonReporting}\n\n` +
                    `Note: Complete analytical breakdown and feedback will be dispatched on Wednesday.`;
      } else {
        zonalBody = await generateAiReport(zone, zoneData, reportingRateStr, zoneData.filter(r => !r.Report || r.Report.toLowerCase() !== 'reported').map(r => r.Woreda), '', {}, {}, [], [], [], currentWeek);
      }

      try {
        const result = await sendViaBrevo(
          recipientEmail,
          `[${dayOfWeek === 1 || dayOfWeek === 2 ? 'Status Update' : 'Feedback & Analysis'}] ${zone} - ${currentWeek}`,
          zonalBody
        );
        console.log(`[Brevo] Sent ${currentWeek} report for ${zone}. Message ID: ${result.messageId}`);
      } catch (err) {
        console.error(`❌ [Brevo Error for ${zone}]:`, err.message);
      }
    }

    // 2. Send Grand Total Regional Report ONLY on Wednesday (dayOfWeek === 3) to Regional Leadership
    if (dayOfWeek === 3) {
      console.log(`[Dispatch] Wednesday detected: Generating Regional Grand Total Report for leadership...`);
      
      const regionalData = rawData.filter(r => r.Week && r.Week.toString().trim().toLowerCase().includes(currentWeek.toLowerCase()));
      const regTotalExpected = regionalData.length;
      const regReported = regionalData.filter(r => r.Report && r.Report.toLowerCase() === 'reported');
      const regReportingRateVal = regTotalExpected > 0 ? (regReported.length / regTotalExpected) * 100 : 0;
      const regReportingRateStr = regReportingRateVal.toFixed(1) + '%';
      const regNonReporting = regionalData.filter(r => !r.Report || r.Report.toLowerCase() !== 'reported').map(r => `${r.Zone} (${r.Woreda})`).join(', ') || 'None';

      const regionalBody = await generateAiReport(
        'Grand Total', 
        regionalData, 
        regReportingRateStr, 
        regionalData.filter(r => !r.Report || r.Report.toLowerCase() !== 'reported').map(r => `${r.Zone} (${r.Woreda})`), 
        '', {}, {}, [], [], [], 
        currentWeek
      );

      try {
        const regResult = await sendViaBrevo(
          regionalEmails,
          `[Feedback & Analysis] Regional Grand Total RMNCH Report - ${currentWeek}`,
          regionalBody
        );
        console.log(`[Brevo] Sent Regional Grand Total report to leadership. Message ID: ${regResult.messageId}`);
      } catch (err) {
        console.error(`❌ [Brevo Error for Regional Leadership]:`, err.message);
      }
    }

    return true;
  } catch (err) {
    console.error('Error in automated email dispatch via Brevo:', err);
    throw err;
  }
}

// Instant test endpoint
app.get('/api/test-email', async (req, res) => {
  try {
    await runRmnchEmailDispatch();
    res.send('RMNCH email dispatch executed successfully via Brevo!');
  } catch (err) {
    res.status(500).send('Error executing email dispatch: ' + err.message);
  }
});

// Scheduled Cron Job: Runs Mon, Tue, Wed at 14:35 UTC
cron.schedule('30 7 * * 1,2,3', async () => {
  console.log('[CRON Triggered] Running scheduled RMNCH email dispatch...');
  await runRmnchEmailDispatch();
});

// Start Server - Bound explicitly to '0.0.0.0' for Render
app.listen(PORT, '0.0.0.0', () => {
  console.log(`Server running on port ${PORT}`);
});
