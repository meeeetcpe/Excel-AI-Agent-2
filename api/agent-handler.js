const { GoogleGenerativeAI } = require("@google/generative-ai");

const genAI = new GoogleGenerativeAI(process.env.GEMINI_APIS_KEY);

const DATA_ANALYZER_PROMPT = `You are an expert Excel data analyst. Your task is to analyze a given JSON dataset based on a user's plain-English request and provide a direct result.

You MUST return your response as a single, valid JSON object with two properties:
1. "summary": A one-sentence summary of the action you took (e.g., "I have calculated the sum of the Sales column.").
2. "result": The result of the analysis. This can be a single value (like a number), or an array of arrays if you are returning a new table of data.

If you cannot fulfill the request, the "summary" should explain why, and the "result" should be null.

DATASET:
{{DATA}}

USER REQUEST:
"{{PROMPT}}"
`;

module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') { return res.status(200).end(); }

  const { data, prompt } = req.body;

  if (!prompt) {
    return res.status(400).json({ error: "A prompt is required." });
  }
  if (!data) {
    return res.status(400).json({ error: "A data selection is required for analysis." });
  }

  try {
    const model = genAI.getGenerativeModel({ model: "gemini-1.5-flash" });
    
    let finalPrompt = DATA_ANALYZER_PROMPT.replace('{{DATA}}', JSON.stringify(data));
    finalPrompt = finalPrompt.replace('{{PROMPT}}', prompt);

    const result = await model.generateContent(finalPrompt);
    const response = await result.response;
    const responseText = response.text();
    
    let jsonResponse;
    try {
      jsonResponse = JSON.parse(responseText);
    } catch(e) {
      // If the AI fails to return JSON, wrap its text in a summary.
      console.error("AI did not return valid JSON. Returning raw text.");
      jsonResponse = { summary: responseText, result: null };
    }

    res.status(200).json(jsonResponse);

  } catch (error) {
    console.error('Error in data analyzer handler:', error);
    res.status(500).json({ summary: "An unexpected error occurred with the AI.", result: null });
  }
};
