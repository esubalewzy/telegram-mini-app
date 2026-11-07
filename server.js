import express from 'express';
import path from 'path';
import { fileURLToPath } from 'url';
import Groq from 'groq-sdk';
import { 
  Document, Packer, Paragraph, TextRun, HeadingLevel, 
  Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, AlignmentType 
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

    let weeklyTrends = {};
    let zoneAllData = selectedZone && selectedZone !== 'Grand Total' 
      ? rawData.filter(r => r.Zone && r.Zone.toLowerCase() === selectedZone.toLowerCase()) 
      : rawData;

    zoneAllData.forEach(r => {
      const wk = r.Week || 'Unknown Week';
      if (!weeklyTrends[wk]) {
        weeklyTrends[wk] = { totalExpected: 0, reportedCount: 0 };
      }
      weeklyTrends[wk].totalExpected++;
      if (r.Report && r.Report.toLowerCase() === 'reported') {
        weeklyTrends[wk].reportedCount++;
      }
    });

    let weeklyTrendSummary = Object.keys(weeklyTrends).map(wk => {
      const rate = ((weeklyTrends[wk].reportedCount / weeklyTrends[wk].totalExpected) * 100).toFixed(1);
      return `Week ${wk}: ${rate}% (${weeklyTrends[wk].reportedCount}/${weeklyTrends[wk].totalExpected})`;
    }).join(' | ');

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

          const coverageNum = totalEligible > 0 ? parseFloat(((totalPerformed / totalEligible) * 100).toFixed(1)) : 0;
          const coverage = coverageNum + '%';
          const indicatorEntry = { name: performedKey, performed: totalPerformed, eligible: totalEligible, coverage: coverageNum };

          if (baseName.startsWith('1.')) categorizedIndicators.cat1_srh.push(indicatorEntry);
          else if (baseName.startsWith('2.')) categorizedIndicators.cat2_maternal.push(indicatorEntry);
          else if (baseName.startsWith('3.')) categorizedIndicators.cat3_epi.push(indicatorEntry);
          else if (baseName.startsWith('4.')) categorizedIndicators.cat4_child.push(indicatorEntry);
          else if (baseName.startsWith('5.')) categorizedIndicators.cat5_nutrition.push(indicatorEntry);
        }
      }
    });

    // HSTP-II National Benchmark Targets mapping
    const hstp2Benchmarks = {
      "Category 1: SRH Services": { target: 75, desc: "Contraceptive Acceptance Rate / Family Planning Target (~75%)" },
      "Category 2: Maternal Services": { target: 81, desc: "ANC4 Coverage (~81%) & Skilled Delivery (~80%)" },
      "Category 3: EPI Services": { target: 90, desc: "Pentavalent-3 / Fully Immunized Target (~90%)" },
      "Category 4: Child Health Service": { target: 85, desc: "Pneumonia/Diarrhea Treatment & Care Seeking (~85%)" },
      "Category 5: Nutrition": { target: 80, desc: "SAM Screening and Growth Monitoring (~80%)" }
    };

    const qualityIssues = [];
    const hepBDMissedWoredas = [];
    const mortalityAlerts = [];

    filteredData.forEach(row => {
      const woredaName = row['Woreda'] || 'Unknown Woreda';

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

      const anc1Key = headers.find(h => h.includes('2.1') && !h.includes('Eligible') && !h.includes('Target'));
      const anc4Key = headers.find(h => h.includes('2.3') && !h.includes('Eligible') && !h.includes('Target'));
      if (anc1Key && anc4Key) {
        const anc1 = parseFloat(row[anc1Key]) || 0;
        const anc4 = parseFloat(row[anc4Key]) || 0;
        if (anc4 > anc1 && qualityIssues.length < 8) {
          qualityIssues.push(`[Logic Consistency Error] ${woredaName}: ANC4 visits (${anc4}) cannot exceed ANC1 visits (${anc1}).`);
        }
      }

      headers.forEach(h => {
        const lowerH = h.toLowerCase();
        if ((lowerH.includes('maternal death') || lowerH.includes('neonatal death') || h.includes('2.8') || h.includes('4.4')) && !h.includes('Eligible') && !h.includes('Target')) {
          const deathVal = parseFloat(row[h]) || 0;
          if (deathVal > 0) {
            mortalityAlerts.push(`[CRITICAL MORTALITY ALERT] ${woredaName} reported ${deathVal} case(s) under [${h}].`);
          }
        }
      });
    });

    const systemPrompt = `You are an expert RMNCH Health Analytics & Data Auditor. Generate a complete, highly structured analytical report.

Greeting Header (Must start with):
"Hello, This Is Weekly Essential RMNCH Service Report Feedback for ${selectedZone || 'Grand Total'} Zone, Week ${selectedWeek || 'All Weeks'}"

Data Context:
- Reporting Rate: ${reportingRate} (${reportingWoredas.length}/${totalExpectedWoredas})
- Non-Reporting Woredas: ${nonReportingWoredas.join(', ') || 'None'}
- Multi-Week Trend Data: ${weeklyTrendSummary}
- Categorized Indicators: ${JSON.stringify(categorizedIndicators)}
- Ethiopian HSTP-II Strategic Benchmarks: ${JSON.stringify(hstp2Benchmarks)}
- Hepatitis BD Missed Audit: ${hepBDMissedWoredas.join(' | ') || 'None'}
- Quality Issues: ${qualityIssues.join(' | ') || 'None'}
- Mortality Alerts: ${mortalityAlerts.join(' | ') || 'No maternal or neonatal deaths reported.'}

Required Report Structure:
1. **Reporting Rate Summary**
2. **Multi-Week Reporting Trend Analysis**
3. **Category 1: SRH Services** (Markdown Table format: | Indicator | Performed | Eligible | Coverage % | vs HSTP-II Target |)
4. **Category 2: Maternal Services** (Markdown Table format with vs HSTP-II Target)
5. **Category 3: EPI Services** (Markdown Table format with vs HSTP-II Target)
6. **Category 4: Child Health Service** (Markdown Table format with vs HSTP-II Target)
7. **Category 5: Nutrition** (Markdown Table format with vs HSTP-II Target)
8. **Hepatitis BD Birth Dose Missed Woredas Audit**
9. **Comprehensive HSTP-II Benchmarking Evaluation**: Explicitly compare the category coverages against the national Ethiopian HSTP-II targets (SRH 75%, Maternal 81%, EPI 90%, Child Health 85%, Nutrition 80%), highlighting gaps, underperforming domains, and successes across all selected categories.
10. **Data Quality & Consistency Audit**
11. **Mortality Alerts** (Strictly Maternal and Neonatal only)`;

    const chatCompletion = await groq.chat.completions.create({
      messages: [
        { role: "system", content: systemPrompt },
        { role: "user", content: message || `Generate full report with category-by-category HSTP-II benchmarking.` }
      ],
      model: "openai/gpt-oss-20b",
      temperature: 0.3,
      max_completion_tokens: 2500,
    });

    const responseText = chatCompletion.choices[0]?.message?.content || 'No response generated.';
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
                    children: [new TextRun({ 
                      text: cellText, 
                      bold: isHeader,
                      color: isHeader ? "FFFFFF" : "000000",
                      size: 18 
                    })]
                  })],
                  width: { size: 1800, type: WidthType.DXA },
                  shading: isHeader 
                    ? { fill: "1F4E78", val: ShadingType.CLEAR, color: "auto" } 
                    : (rowIndex % 2 === 0 ? { fill: "F2F4F8", val: ShadingType.CLEAR, color: "auto" } : { fill: "FFFFFF", val: ShadingType.CLEAR, color: "auto" })
                });
              })
            });
          });

          docChildren.push(new Table({
            rows: docTableRows,
            width: { size: 100, type: WidthType.PERCENTAGE },
            borders: {
              top: { style: BorderStyle.SINGLE, size: 6, color: "1F4E78" },
              bottom: { style: BorderStyle.SINGLE, size: 6, color: "1F4E78" },
              left: { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC" },
              right: { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC" },
              insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: "D3D3D3" },
              insideVertical: { style: BorderStyle.SINGLE, size: 4, color: "D3D3D3" },
            }
          }));
          docChildren.push(new Paragraph({ text: "", spacing: { after: 150 } }));
        }
      } else {
        if (line !== '') {
          docChildren.push(new Paragraph({
            children: [new TextRun({ text: lines[i], size: 22 })],
            spacing: { after: 100 }
          }));
        }
        i++;
      }
    }

    const doc = new Document({
      sections: [{ properties: {}, children: docChildren }],
    });

    const buffer = await Packer.toBuffer(doc);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', 'attachment; filename=RMNCH_Service_Report_Feedback.docx');
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
