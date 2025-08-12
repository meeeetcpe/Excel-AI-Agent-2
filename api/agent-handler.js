const { GoogleGenerativeAI } = require("@google/generative-ai");

const genAI = new GoogleGenerativeAI(process.env.GEMINI_API_KEY);

// The agent prompt now only has one tool
const AGENT_PROMPT = `You are an AI agent inside Microsoft Excel. Your goal is to help the user with their request. The only tool you have is a "data_analyzer".

If the user's request involves manipulating, calculating, filtering, or analyzing a selected range of data, use the "data_analyzer" tool.

If the user asks a question that cannot be answered by analyzing data (like a general knowledge question), you must respond with a JSON object where the "summary" explains that you cannot search the internet.

- User Request: "find the total of the 'Sales' column"
- Your Response: {"tool": "data_analyzer", "parameters": {"prompt": "find the total of the sales column"}}

- User Request: "what is the capital of France?"
- Your Response: {"tool": "none", "result": {"summary": "I am an Excel data agent and cannot search the internet for that information."}}
`;

async function runDataAnalyzer(data, prompt) {
  const model = genAI.getGenerativeModel({ model: "gemini-1.5-flash" });
  const DATA_ANALYZER_PROMPT = `You are an expert Excel data analyst. Your task is to analyze a given JSON dataset based on a user's plain-English request. You must return your response as a single, valid JSON object with two properties:
1. "summary": A one-sentence summary of the action you took.
2. "result": The result of the analysis. This can be a single value, an array, or an array of arrays.`;
  const fullPrompt = `${DATA_ANALYZER_PROMPT}\n\nDATASET:\n${JSON.stringify(data)}\n\nUSER REQUEST:\n"${prompt}"`;
  const result = await model.generateContent(fullPrompt);
  const response = await result.response;
  return JSON.parse(response.text());
}


module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') { return res.status(200).end(); }

  const { data, prompt } = req.body;

  try {
    const agentModel = genAI.
