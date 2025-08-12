const { GoogleGenerativeAI } = require("@google/generative-ai");

// 1. Initialize the Gemini Client
const genAI = new GoogleGenerativeAI(process.env.GEMINI_API_KEY);

// 2. Define the Instructions for the AI
const CHART_NARRATIVE_PROMPT = `You are a crisp analytical chart caption writer. Given the data series and categories, write 3-5 bullet points highlighting: the key trend, any anomalies or outliers, a significant comparison, and a practical takeaway. Return only the bullet points as a single block of text.`;

// 3. Define the Serverless Function
module.exports = async (req, res) => {
  // 4. Set Headers for Security (CORS)
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  // Handle browser's pre-flight request
  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  // 5. Get Data from the Request
  const { data } = req.body;
  if (!data) {
    return res.status(400).json({ error: 'Data is required' });
  }

  try {
    // 6. Prepare and Send the Request to Gemini
    const model = genAI.getGenerativeModel({ model: "gemini-1.5-flash" });
    const dataAsString = data.map(row => row.join('\t')).join('\n');
    const fullPrompt = `${CHART_NARRATIVE_PROMPT}\n\nHere is the data:\n${dataAsString}`;

    const result = await model.generateContent(fullPrompt);
    const response = await result.response;
    const text = response.text();

    // 7. Send the AI's Response Back to Excel
    res.status(200).json({ narrative: text });
  } catch (error) {
    console.error('Error with Gemini API:', error);
    res.status(500).json({ error: 'Failed to generate narrative from Gemini API' });
  }
};