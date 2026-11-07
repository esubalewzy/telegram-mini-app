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

// Google Sheet ID from your link
const SHEET_ID = '1d8_5Q2W3aGV68W_DdjrPrFEJnuRSDprrb1b3kngbTpo';
const FINAL_SHEET_CSV_URL = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:csv&sheet=Final`;

// Simple CSV parser for fetched Google Sheets export
async function fetchFinalSheetData() {
  try {
    const response = await fetch(FINAL_SHEET_CSV_URL);
    const csvText = await response.text();
    
    // Parse CSV lines into clean row objects
    const lines = csvText.split('\n').filter(line => line.trim() !== '');
    if (lines.length === 0) return [];

    const headers = lines[0].split(',').map(h => h.replace(/^"(.*)"$/, '$1').trim());

    const rows = lines.slice(1).map(line => {
      // Split by comma ignoring commas inside quotes
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

app.post('/api/chat', async (req, res) => {
  try {
    const { message, selectedZone, selectedWeek, user } = req.body;

    // 1. Fetch live data from Final sheet
    const rawData = await fetchFinalSheetData();

    // 2. Filter data by Zone and Week
    let filteredData = rawData;
    if (selectedZone && selectedZone !== 'Grand Total') {
      filteredData = filteredData.filter(r => r.Zone && r.Zone.toLowerCase() === selectedZone.toLowerCase());
    }
    if (selectedWeek && selectedWeek !== 'All Weeks') {
      filteredData = filteredData.filter(r => r.Week && r.Week.toLowerCase().includes(selectedWeek.toLowerCase()));
    }

    // Identify non-reporting / missing response woredas
    const nonReportingWoredas = filteredData
      .filter(r => r.Report !== 'Reported' || !r.Report)
      .map(r => r.Woreda)
      .filter(Boolean);

    const systemPrompt = `You are an expert Health & RMNCH Business Analytics Assistant.
You are analyzing live data from the SWERHB Essential RMNCH Service Weekly Report ("Final" tab).

Selected Filter Scope:
- Zone Selection: ${selectedZone || 'Grand Total (All Zones)'}
- Week Selection: ${selectedWeek || 'All Weeks'}

Filtered Row Count: ${filteredData.length}
Non-Reporting / No-Response Woredas for this selection: ${nonReportingWoredas.length > 0 ? nonReportingWoredas.join(', ') : 'All woredas reported'}

Dataset Context (Filtered Rows):
${JSON.stringify(filteredData.slice(0, 50), null, 2)}

Instructions for AI Analysis:
1. **Indicator Coverage Summary**:
   - For every key indicator (e.g. 1.1 FP, 1.2 Long-acting FP, 2.2 ANC early, 2.7 Skilled Delivery, 3.1 Penta 1, 3.2 Penta 3, 3.5 Measles 1, 5.1 Growth Monitoring, etc.), aggregate total Performed vs total Eligible across the selected rows.
   - Calculate Coverage Percentage: (Total Performed / Total Eligible) * 100.
2. **No-Response / Missing Woredas**:
   - Explicitly list any Woredas that have NOT submitted data or are marked as un-reported for the selected week.
3. **Zonal vs Grand Total**:
   - Provide localized insights for the requested Zone, or present aggregated overall performance if "Grand Total" is selected.
4. Keep the output clean, professional, and well-structured with clear section headers, bullet points, and percentage metrics.`;

    const chatCompletion = await groq.chat.completions.create({
      messages: [
        { role: "system", content: systemPrompt },
        { role: "user", content: message || `Provide indicator coverage summary and list missing woredas for ${selectedZone || 'Grand Total'} during ${selectedWeek || 'All Weeks'}.` }
      ],
      model: "openai/gpt-oss-20b",
      temperature: 0.5,
      max_completion_tokens: 1024,
    });

    const responseText = chatCompletion.choices[0]?.message?.content || 'No response generated.';
    res.json({ reply: responseText });

  } catch (error) {
    console.error('Groq/Sheet Error Details:', error);
    res.status(500).json({ reply: 'Sorry, I encountered an issue fetching or analyzing the sheet data.' });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
