import express from 'express';
import path from 'path';
import { fileURLToPath } from 'url';
import { GoogleGenerativeAI } from '@google/generative-ai';

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

// Initialize Gemini API securely using Render Environment Variables
const genAI = new GoogleGenerativeAI(process.env.GEMINI_API_KEY);

app.post('/api/chat', async (req, res) => {
  try {
    const { message, user } = req.body;

    if (!message) {
      return res.status(400).json({ reply: 'Please provide a message.' });
    }

    // Initialize Gemini 3.8 Flash with your AI Studio configurations
    const model = genAI.getGenerativeModel({
      model: "gemini-3.8-flash",
      tools: [{ googleSearch: {} }],
      generationConfig: {
        maxOutputTokens: 8192,
        // thinkingLevel configuration matches your studio settings
        thinkingConfig: { thinkingLevel: "medium" }
      }
    });

    const systemPrompt = `You are a helpful business analytics assistant for ${user?.first_name || 'the user'}. 
    Help them understand their Looker Studio dashboard performance metrics, sales trends, and key performance indicators clearly and concisely.`;

    const result = await model.generateContent(`${systemPrompt}\nUser Question: ${message}`);
    const responseText = result.response.text();

    res.json({ reply: responseText });
  } catch (error) {
    console.error('AI Error Details:', error);
    res.status(500).json({ reply: 'Sorry, I encountered an issue retrieving an answer.' });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
