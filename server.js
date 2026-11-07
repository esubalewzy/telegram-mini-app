import express from 'express';
import path from 'path';
import { fileURLToPath } from 'url';
import Groq from 'groq-sdk';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
app.use(express.json());

// Serve static assets
app.use(express.static(__dirname));

// Explicit route for the root URL
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

// Initialize Groq client securely using Render Environment Variables
const groq = new Groq({ apiKey: process.env.GROQ_API_KEY });

app.post('/api/chat', async (req, res) => {
  try {
    const { message, user } = req.body;

    if (!message) {
      return res.status(400).json({ reply: 'Please provide a message.' });
    }

    // Live or baseline zone dataset for detailed performance breakdown
    const liveZoneData = [
      { zone: "North", sales: 54000, target: 50000, status: "Exceeding" },
      { zone: "South", sales: 31000, target: 40000, status: "Underperforming" },
      { zone: "East", sales: 42000, target: 40000, status: "On Track" },
      { zone: "West", sales: 38000, target: 45000, status: "Needs Attention" }
    ];

    const systemPrompt = `You are an expert business analytics assistant for ${user?.first_name || 'the user'}. 
    Here is the live dataset by zone:
    ${JSON.stringify(liveZoneData, null, 2)}
    
    Provide a detailed, professional zone-by-zone analysis and feedback report based strictly on this data. 
    Highlight strengths, weaknesses, and actionable recommendations for each zone clearly and concisely.`;

    // Call Groq's active model API
    const chatCompletion = await groq.chat.completions.create({
      messages: [
        { role: "system", content: systemPrompt },
        { role: "user", content: message }
      ],
      model: "openai/gpt-oss-20b", // Updated to a currently supported active model ID
      temperature: 0.6,
      max_completion_tokens: 1024,
    });

    const responseText = chatCompletion.choices[0]?.message?.content || 'No response generated.';
    res.json({ reply: responseText });

  } catch (error) {
    console.error('Groq AI Error Details:', error);
    res.status(500).json({ reply: 'Sorry, I encountered an issue generating the report from Groq AI.' });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
