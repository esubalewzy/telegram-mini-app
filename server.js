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

async function fetchFinalSheetData() {
  try {
    const response = await fetch(FINAL_SHEET_CSV_URL);
    const csvText = await response.text();
    
    const lines = csvText.split('\n').filter(line => line.trim() !== '');
    if (lines.length === 0) return [];

    const headers = lines[0].split(',').map(h => h.replace(/^"(.*)"$/, '$1').trim());

    const rows = lines.slice(1).map(line => {
      const values = line.match(/(".*?"|[^",\s]+)(?=\s*,|\s*$)/g) || line.split(',');
      let obj = {};
      headers.forEach((h, index) => {
        let val = values[index] ? values[index].replace(/^"(.*)"$/, '$1').trim() : '';
        obj[h] = val;
      });
      return obj;
    });

    return rows;
  } catch (err) {
    console.error('Error fetching Google Sheet:', err);
    return [];
  }
}

// Endpoint to dynamically provide available filter options to the frontend
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

    let filteredData = rawData;
    if (selectedZone && selectedZone !== 'Grand Total') {
      filteredData = filteredData.filter(r => r.Zone && r.Zone.toLowerCase() === selectedZone.toLowerCase());
    }
    if (selectedWeek && selectedWeek !== 'All Weeks') {
      filteredData = filteredData.filter(r => r.Week && r.Week.toLowerCase().includes(selectedWeek.toLowerCase()));
    }

    const nonReportingWoredas = filteredData
      .filter(r => r.Report !== 'Reported' || !r.Report)
      .map(r => r.Woreda)
      .filter(Boolean);

    const systemPrompt = `You are an expert Health & RMNCH Business Analytics Assistant analyzing live data from the "Final" tab of the SWERHB spreadsheet.

Selected Filter Scope:
- Zone Selection: ${selectedZone || 'Grand Total (All Zones)'}
- Week Selection: ${selectedWeek || 'All Weeks'}

Filtered Row Count: ${filteredData.length}
Non-Reporting / No-Response Woredas: ${nonReportingWoredas.length > 0 ? nonReportingWoredas.join(', ') : 'All woredas reported'}

Dataset Sample Context:
${JSON.stringify(filteredData.slice(0, 30), null, 2)}

Instructions for AI Analysis:
1. **Indicator Coverage Summary**: Calculate coverage percentages (Performed / Eligible * 100) for key indicators based on the matching columns in the dataset.
2. **Missing/No-Response Woredas**: Explicitly list woredas that haven't reported.
3. Keep the output clean, structured, professional, and clear.`;

    const chatCompletion = await groq.chat.completions.create({
      messages: [
        { role: "system", content: systemPrompt },
        { role: "user", content: message || `Generate full coverage summary and list missing woredas for ${selectedZone || 'Grand Total'}.` }
      ],
      model: "openai/gpt-oss-20b",
      temperature: 0.5,
      max_completion_tokens: 1024,
    });

    const responseText = chatCompletion.choices[0]?.message?.content || 'No response generated.';
    res.json({ reply: responseText });

  } catch (error) {
    console.error('Groq/Sheet Error Details:', error);
    res.status(500).json({ reply: 'Sorry, I encountered an issue analyzing the sheet data.' });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
