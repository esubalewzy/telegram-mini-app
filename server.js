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

    // 2. Extract missing/non-reporting woredas
    const nonReportingWoredas = filteredData
      .filter(r => r.Report && r.Report.toLowerCase() !== 'reported')
      .map(r => r.Woreda)
      .filter(Boolean);

    // 3. Pre-aggregate performed vs. eligible numeric values in Node.js
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
          indicatorSummary[baseName] = {
            performed: totalPerformed,
            eligible: totalEligible,
            coverage: coverage
          };
        }
      }
    });

    // Compact system prompt sent to Groq
    const systemPrompt = `You are an expert Health & RMNCH Business Analytics Assistant analyzing live SWERHB reporting data.

Scope Filters:
- Zone: ${selectedZone || 'Grand Total'}
- Week: ${selectedWeek || 'All Weeks'}
- Total Reporting Woredas: ${filteredData.length}
- Non-Reporting / Missing Woredas: ${nonReportingWoredas.length > 0 ? nonReportingWoredas.join(', ') : 'None (All Reported)'}

Pre-Aggregated Indicator Summary (Performed / Eligible = Coverage %):
${JSON.stringify(indicatorSummary, null, 2)}

Instructions:
1. Provide a clean summary of indicator coverage (Performed vs Eligible and calculated Coverage %).
2. Explicitly list any missing/non-reporting Woredas.
3. Keep the output clean, professional, concise, and structured with bullet points.`;

    const chatCompletion = await groq.chat.completions.create({
      messages: [
        { role: "system", content: systemPrompt },
        { role: "user", content: message || `Generate coverage summary and list missing woredas for ${selectedZone || 'Grand Total'}.` }
      ],
      model: "openai/gpt-oss-20b",
      temperature: 0.5,
      max_completion_tokens: 1024,
    });

    const responseText = chatCompletion.choices[0]?.message?.content || 'No response generated.';
    res.json({ reply: responseText });

  } catch (error) {
    console.error('Groq/Sheet Error Details:', error);
    res.status(500).json({ reply: 'Error generating report. Please try again.' });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
