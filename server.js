import express from 'express';
import path from 'path';
import { fileURLToPath } from 'url';
import Groq from 'groq-sdk';
import { 
  Document, Packer, Paragraph, TextRun, HeadingLevel, 
  Table, TableRow, TableCell, WidthType, BorderStyle 
} from 'docx';

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

    const qualityIssues = [];
    const hepBDMissedWoredas = [];
    const mortalityAlerts = [];

    filteredData.forEach(row => {
      const woredaName = row['Woreda'] || 'Unknown Woreda';

      // HepBD Audit
      const skilledKey = headers.find(h => h.includes('2.7') && !h.includes('Eligible') && !h.includes('Target'));
      const hepBDKey = headers.find(h => h.includes('3.3') && !h.includes('Eligible') && !h.includes('Target'));
      if (skilledKey && hepBDKey) {
        const skilledVal = parseFloat(row[skilledKey]) || 0;
        const hepBDVal = parseFloat(row[hepBDKey]) || 0;
        if (skilledVal > hepBDVal) {
          const missed = skilledVal - hepBDVal;
          hepBDMissedWoredas.push(`${woredaName}: ${skilledVal} skilled births vs ${hepBDVal} HepBD doses (${missed} missed)`);
        }
      }

      // 1. Performed > Eligible consistency check
      headers.forEach(header => {
        if (header.includes('Eligible')) {
          const baseName = header.replace('Eligible', '').trim();
          const performedKey = headers.find(h => h.startsWith(baseName) && !h.includes('Eligible') && !h.includes('Target'));
          
          if (performedKey) {
            const perfVal = parseFloat(row[performedKey]);
            const eligVal = parseFloat(row[header]);
            if (!isNaN(perfVal) && !isNaN(eligVal) && eligVal > 0 && perfVal > eligVal && qualityIssues.length < 6) {
              qualityIssues.push(`[Over-reporting Error] ${woredaName} - ${performedKey}: Performed (${perfVal}) exceeds Eligible pool (${eligVal}).`);
            }
          }
        }
      });

      // 2. Related indicator consistency checks (e.g. ANC1 vs ANC4, or Pentavalent 1 vs Pentavalent 3)
      const anc1Key = headers.find(h => h.includes('2.1') && !h.includes('Eligible') && !h.includes('Target'));
      const anc4Key = headers.find(h => h.includes('2.3') && !h.includes('Eligible') && !h.includes('Target'));
      if (anc1Key && anc4Key) {
        const anc1 = parseFloat(row[anc1Key]) || 0;
        const anc4 = parseFloat(row[anc4Key]) || 0;
        if (anc4 > anc1 && qualityIssues.length < 8) {
          qualityIssues.push(`[Logic Consistency Error] ${woredaName}: ANC4 visits (${anc4}) cannot exceed ANC1 visits (${anc1}).`);
        }
      }

      // 3. Strict Maternal & Neonatal Mortality check (Only list if > 0)
      headers.forEach(h => {
        const lowerH = h.toLowerCase();
        if ((lowerH.includes('maternal death') || lowerH.includes('neonatal death') || h.includes('2.8') || h.includes('4.4')) && !h.includes('Eligible') && !h.includes('Target')) {
          const deathVal = parseFloat(row[h]) || 0;
          if (deathVal > 0) {
            mortalityAlerts.push(`[CRITICAL MORTALITY ALERT] ${woredaName} reported ${deathVal} case(s) under [${h}]. Immediate investigation required.`);
          }
        }
      });
    });

    let comparisonContext = '';
    if (!selectedZone || selectedZone === 'Grand Total') {
      let zoneStats = {};
      rawData.forEach(r => {
        const z = r.Zone || 'Unknown Zone';
        if (!zoneStats[z]) zoneStats[z] = { total: 0, reported: 0 };
        zoneStats[z].total++;
        if (r.Report && r.Report.toLowerCase() === 'reported') zoneStats[z].reported++;
      });
      comparisonContext = `Zone vs Zone Comparison against HSTP-II targets: ${JSON.stringify(zoneStats)}`;
    } else {
      let woredaStats = filteredData.map(r => ({ woreda: r.Woreda, status: r.Report || 'Not Reported' }));
      comparisonContext = `Woreda Performance in ${selectedZone} against HSTP-II targets: ${JSON.stringify(woredaStats)}`;
    }

    const systemPrompt = `You are an expert RMNCH Health Analytics & Data Auditor. Generate a complete, highly structured analytical report.

Greeting Header (Must start with):
"Hello, This Is Weekly Essential RMNCH Service Report Feedback for ${selectedZone || 'Grand Total'} Zone, Week ${selectedWeek || 'All Weeks'}"

Data Context:
- Reporting Rate: ${reportingRate} (${reportingWoredas.length}/${totalExpectedWoredas})
- Non-Reporting Woredas: ${nonReportingWoredas.join(', ') || 'None'}
- Categorized Indicators: ${JSON.stringify(categorizedIndicators)}
- Hepatitis BD Missed Audit: ${hepBDMissedWoredas.join(' | ') || 'None'}
- Consistency & Over-reporting Quality Issues: ${qualityIssues.join(' | ') || 'None'}
- Mortality Alerts (Maternal/Neonatal Only): ${mortalityAlerts.join(' | ') || 'No maternal or neonatal deaths reported in this selection.'}
- ${comparisonContext}

Required Report Structure & Instructions:
1. **Reporting Rate Summary**
2. **Category 1: SRH Services** (Markdown Table format: | Indicator | Performed | Eligible | Coverage % |)
3. **Category 2: Maternal Services** (Markdown Table format)
4. **Category 3: EPI Services** (Markdown Table format)
5. **Category 4: Child Health Service** (Markdown Table format)
6. **Category 5: Nutrition** (Markdown Table format)
7. **Hepatitis BD Birth Dose Missed Woredas Audit**
8. **Comparative Benchmarking vs. Ethiopian HSTP-II Guidelines**: Evaluate coverage against national HSTP-II strategic targets (e.g., ANC4, Skilled Delivery, Pentavalent 3 coverage benchmarks). Compare Zones vs. Zones or Woredas vs. Woredas accordingly.
9. **Data Quality & Consistency Audit**: Highlight over-reporting (performed > eligible) and related indicator logic errors (e.g., ANC4 > ANC1).
10. **Mortality Alerts**: List strictly maternal and neonatal deaths if reported. If none, explicitly state that zero maternal or neonatal deaths were recorded.`;

    const chatCompletion = await groq.chat.completions.create({
      messages: [
        { role: "system", content: systemPrompt },
        { role: "
