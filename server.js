import express from 'express';
import path from 'path';
import { fileURLToPath } from 'url';
import Groq from 'groq-sdk';

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

    // 1. Filter rows by selected Zone & Week
    let filteredData = rawData;
    if (selectedZone && selectedZone !== 'Grand Total') {
      filteredData = filteredData.filter(r => r.Zone && r.Zone.toLowerCase() === selectedZone.toLowerCase());
    }
    if (selectedWeek && selectedWeek !== 'All Weeks') {
      filteredData = filteredData.filter(r => r.Week && r.Week.toLowerCase().includes(selectedWeek.toLowerCase()));
    }

    // 2. Reporting Rate Breakdown based on 'Report' column
    const totalExpectedWoredas = filteredData.length;
    const reportingWoredas = filteredData.filter(r => r.Report && r.Report.toLowerCase() === 'reported');
    const nonReportingRows = filteredData.filter(r => !r.Report || r.Report.toLowerCase() !== 'reported');
    const nonReportingWoredas = nonReportingRows.map(r => r.Woreda).filter(Boolean);
    const reportingRate = totalExpectedWoredas > 0 ? ((reportingWoredas.length / totalExpectedWoredas) * 100).toFixed(1) + '%' : '0%';

    // 3. Pre-aggregate performed vs eligible with Full Indicator Names/Headers
    const headers = Object.keys(rawData[0] || {});
    const indicatorSummary = {};

    headers.forEach(header => {
      if (header.includes('Eligible')) {
        const baseName = header.replace('Eligible', '').trim();
        const performedKey = headers.find(h => h.startsWith(baseName) && !h.includes('Eligible') && !h.includes('Target'));
        
        if (performedKey) {
          let totalPerformed = 0;
          let totalEligible = 0;

          filteredData.forEach(row => {
            const perfVal = parseFloat(row[performedKey]) || 0;
            const eligVal = parseFloat(row[header]) || 0;
            totalPerformed += perfVal;
            totalEligible += eligVal;
          });

          const coverage = totalEligible > 0 ? ((totalPerformed / totalEligible) * 100).toFixed(1) + '%' : 'N/A';
          indicatorSummary[performedKey] = {
            fullName: performedKey,
            performed: totalPerformed,
            eligible: totalEligible,
            coverage: coverage
          };
        }
      }
    });

    // 4. Data Quality Audit (Performed > Eligible with indicator full names)
    const qualityIssues = [];
    filteredData.forEach(row => {
      const woredaName = row['Woreda'] || 'Unknown Woreda';
      const zoneName = row['Zone'] || 'Unknown Zone';

      headers.forEach(header => {
        if (header.includes('Eligible')) {
          const baseName = header.replace('Eligible', '').trim();
          const performedKey = headers.find(h => h.startsWith(baseName) && !h.includes('Eligible') && !h.includes('Target'));
          
          if (performedKey) {
            const perfVal = parseFloat(row[performedKey]);
            const eligVal = parseFloat(row[header]);

            if (!isNaN(perfVal) && !isNaN(eligVal) && eligVal > 0 && perfVal > eligVal) {
              qualityIssues.push({
                woreda: woredaName,
                zone: zoneName,
                indicator: performedKey,
                performed: perfVal,
                eligible: eligVal,
                discrepancy: `Reported performed value (${perfVal}) exceeds eligible pool (${eligVal})`
              });
            }
          }
        }
      });
    });

    // Structured Prompt enforcing exact sequence
    const systemPrompt = `You are an expert Health & RMNCH Analytics and Data Quality Auditor. Analyze the provided structured data from the SWERHB report.

Scope Filters:
- Zone: ${selectedZone || 'Grand Total'}
- Week: ${selectedWeek || 'All Weeks'}

Reporting Rate Statistics:
- Total Expected Woredas / Submissions: ${totalExpectedWoredas}
- Submitted / Reported Count: ${reportingWoredas.length}
- Reporting Rate: ${reportingRate}
- Non-Reporting / Missing Woredas: ${nonReportingWoredas.length > 0 ? nonReportingWoredas.join(', ') : 'None (100% reporting)'}

Indicator Summaries (Performed, Eligible, Coverage %):
${JSON.stringify(indicatorSummary, null, 2)}

Data Quality Anomalies (Performed > Eligible):
${JSON.stringify(qualityIssues, null, 2)}

**Strict Output Generation Order & Instructions:**
1. **Reporting Rate Analysis**: Start by evaluating the reporting rate. State the submission numbers, the exact reporting percentage, and explicitly list the non-reporting woredas.
2. **Good Performance Side**: Detail the indicators showing strong performance and high coverage percentages. Include the full indicator names (e.g., family planning, antenatal care, vaccinations) along with their numbers.
3. **Poor Performance Side**: Highlight lagging indicators, low coverage rates, or areas needing immediate intervention. Include full indicator names.
4. **Data Quality Issues**: Point out specific data quality discrepancies found in reported woredas where performed numbers exceed eligible pools. You MUST include the exact **Woreda name** and the **full indicator name** for every issue detected.

Keep the output professional, cleanly structured with bold headings, and formatted with bullet points.`;

    const chatCompletion = await groq.chat.completions.create({
      messages: [
        { role: "system", content: systemPrompt },
        { role: "user", content: message || `Analyze reporting rate, performance highlights, and data quality issues for ${selectedZone || 'Grand Total'}.` }
      ],
      model: "openai/gpt-oss-20b",
      temperature: 0.4,
      max_completion_tokens: 1500,
    });

    const responseText = chatCompletion.choices[0]?.message?.content || 'No response generated.';
    res.json({ reply: responseText });

  } catch (error) {
    console.error('Groq/Sheet Error Details:', error);
    res.status(500).json({ reply: 'Error generating structured report. Please try again.' });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
