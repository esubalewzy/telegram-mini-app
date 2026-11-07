import express from 'express';
import path from 'path';
import { fileURLToPath } from 'url';
import Groq from 'groq-sdk';
import { Document, Packer, Paragraph, TextRun, HeadingLevel } from 'docx';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
app.use(express.json());
app.use(express.static(__dirname));

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

const groq = new Groq({ apiKey: process.env.GROQ_API_KEY });

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

app.post('/api/chat', async (req, res) => {
  try {
    const { message, selectedZone, selectedWeek, user } = req.body;

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

    const headers = Object.keys(rawData[0] || {});
    let categorizedIndicators = {
      cat1_srh: [],
      cat2_maternal: [],
      cat3_epi: [],
      cat4_child: [],
      cat5_nutrition: []
    };

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

          const coverage = totalEligible > 0 ? ((totalPerformed / totalEligible) * 100).toFixed(1) + '%' : 'N/A';
          const indicatorEntry = { name: performedKey, performed: totalPerformed, eligible: totalEligible, coverage };

          if (baseName.startsWith('1.')) categorizedIndicators.cat1_srh.push(indicatorEntry);
          else if (baseName.startsWith('2.')) categorizedIndicators.cat2_maternal.push(indicatorEntry);
          else if (baseName.startsWith('3.')) categorizedIndicators.cat3_epi.push(indicatorEntry);
          else if (baseName.startsWith('4.')) categorizedIndicators.cat4_child.push(indicatorEntry);
          else if (baseName.startsWith('5.')) categorizedIndicators.cat5_nutrition.push(indicatorEntry);
        }
      }
    });

    // Anomaly Audit: HepBD Gap + Performed > Eligible + Mortality
    const qualityIssues = [];
    const hepBDMissedWoredas = [];

    filteredData.forEach(row => {
      const woredaName = row['Woreda'] || 'Unknown Woreda';

      // HepBD Gap Check: Skilled Delivery (2.7) vs Hepatitis BD (3.3)
      const skilledKey = headers.find(h => h.includes('2.7') && !h.includes('Eligible') && !h.includes('Target'));
      const hepBDKey = headers.find(h => h.includes('3.3') && !h.includes('Eligible') && !h.includes('Target'));

      if (skilledKey && hepBDKey) {
        const skilledVal = parseFloat(row[skilledKey]) || 0;
        const hepBDVal = parseFloat(row[hepBDKey]) || 0;
        if (skilledVal > hepBDVal) {
          const missed = skilledVal - hepBDVal;
          hepBDMissedWoredas.push(`${woredaName}: ${skilledVal} skilled births but only ${hepBDVal} Hepatitis BD doses (${missed} babies missed birth dose)`);
        }
      }

      // Performed > Eligible Check
      headers.forEach(header => {
        if (header.includes('Eligible')) {
          const baseName = header.replace('Eligible', '').trim();
          const performedKey = headers.find(h => h.startsWith(baseName) && !h.includes('Eligible') && !h.includes('Target'));
          
          if (performedKey) {
            const perfVal = parseFloat(row[performedKey]);
            const eligVal = parseFloat(row[header]);
            if (!isNaN(perfVal) && !isNaN(eligVal) && eligVal > 0 && perfVal > eligVal && qualityIssues.length < 10) {
              qualityIssues.push(`[Consistency Error] ${woredaName} - ${performedKey}: Performed (${perfVal}) > Eligible (${eligVal}).`);
            }
          }
        }
      });

      // Mortality Check
      headers.forEach(h => {
        if ((h.includes('2.8') || h.toLowerCase().includes('maternal death') || h.includes('4.4') || h.toLowerCase().includes('neonatal')) && !h.includes('Eligible') && !h.includes('Target')) {
          const deathVal = parseFloat(row[h]) || 0;
          if (deathVal > 0) {
            qualityIssues.push(`[Mortality Alert] ${woredaName} reported: ${h} = ${deathVal}`);
          }
        }
      });
    });

    // Comparative Analysis Data Pre-computation
    let comparisonContext = '';
    if (!selectedZone || selectedZone === 'Grand Total') {
      // Zone vs Zone breakdown
      let zoneStats = {};
      rawData.forEach(r => {
        const z = r.Zone || 'Unknown Zone';
        if (!zoneStats[z]) zoneStats[z] = { expected: 0, reported: 0 };
        zoneStats[z].expected++;
        if (r.Report && r.Report.toLowerCase() === 'reported') zoneStats[z].reported++;
      });
      comparisonContext = `Comparative Analysis (Zone vs Zone Performance):\n${JSON.stringify(zoneStats, null, 2)}`;
    } else {
      // Woreda vs Woreda breakdown inside selected zone
      let woredaStats = filteredData.map(r => ({
        woreda: r.Woreda,
        reported: r.Report === 'Reported' ? 'Yes' : 'No'
      }));
      comparisonContext = `Comparative Analysis (Woreda vs Woreda Performance in ${selectedZone}):\n${JSON.stringify(woredaStats, null, 2)}`;
    }

    const systemPrompt = `You are an expert RMNCH Health Analytics & Data Auditor. Generate a complete, highly structured analytical report adhering strictly to this sequence:

Greeting Header: "Hello, This is ${selectedZone || 'Grand Total'} Zone Week ${selectedWeek || 'All Weeks'} report summary"

Data Context:
- Reporting Rate: ${reportingRate} (${reportingWoredas.length}/${totalExpectedWoredas})
- Non-Reporting Woredas: ${nonReportingWoredas.join(', ') || 'None'}
- Categorized Indicators: ${JSON.stringify(categorizedIndicators)}
- Hepatitis BD Audit (Missed Birth Doses): ${hepBDMissedWoredas.length > 0 ? hepBDMissedWoredas.join(' | ') : 'None (All skilled birth babies vaccinated)'}
- Inconsistency Anomalies & Mortalities: ${qualityIssues.length > 0 ? qualityIssues.join(' | ') : 'None'}
- ${comparisonContext}

Required Report Structure:
1. **Reporting Rate Summary**: State submission rates and list non-reporting woredas.
2. **Category 1: SRH Services**: List all indicators with performed, eligible, coverage %.
3. **Category 2: Maternal Services**: List all indicators with performed, eligible, coverage %.
4. **Category 3: EPI Services**: List all indicators with performed, eligible, coverage %.
5. **Category 4: Child Health Service**: List all indicators with performed, eligible, coverage %.
6. **Category 5: Nutrition**: List all indicators with performed, eligible, coverage %.
7. **Hepatitis BD Birth Dose Missed Woredas Audit**: Explicitly list all Woredas failing to vaccinate all skilled birth attendant babies with HepBD.
8. **Comparative Analysis & Benchmark**:
   - If Grand Total was selected: Compare Zone vs. Zone performance and rank them.
   - If a specific Zone was selected: Compare Woreda vs. Woreda performance, highlighting top and poor performer woredas.
9. **Data Quality, Anomalies & Mortalities**: Highlight consistency errors (performed > eligible) and reported maternal/neonatal deaths.`;

    const chatCompletion = await groq.chat.completions.create({
      messages: [
        { role: "system", content: systemPrompt },
        { role: "user", content: message || `Generate full report for ${selectedZone || 'Grand Total'}.` }
      ],
      model: "openai/gpt-oss-20b",
      temperature: 0.3,
      max_completion_tokens: 1500,
    });

    const responseText = chatCompletion.choices[0]?.message?.content || 'No response generated.';
    res.json({ reply: responseText });

  } catch (error) {
    console.error('Groq/Sheet Error Details:', error);
    res.status(500).json({ reply: 'Error generating report. Please try again.' });
  }
});

// Endpoint to generate downloadable Word Document (.docx)
app.post('/api/export-docx', async (req, res) => {
  try {
    const { reportText, title } = req.body;

    const doc = new Document({
      sections: [{
        properties: {},
        children: [
          new Paragraph({
            text: title || "SWERHB RMNCH Analytics Report",
            heading: HeadingLevel.HEADING_1,
          }),
          ...reportText.split('\n').map(line => new Paragraph({
            children: [new TextRun(line)]
          }))
        ],
      }],
    });

    const buffer = await Packer.toBuffer(doc);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', 'attachment; filename=RMNCH_Analytics_Report.docx');
    res.send(buffer);
  } catch (err) {
    console.error('Docx Export Error:', err);
    res.status(500).json({ error: 'Failed to generate Word document.' });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
