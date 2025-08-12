const { GoogleGenerativeAI } = require("@google/generative-ai");

const genAI = new GoogleGenerativeAI(process.env.GEMINI_API_KEY);

const AGENT_PROMPT = `You are an AI agent inside Microsoft Excel. Your goal is to help the user with their request. The only tool you have is a "data_analyzer".

If the user's request involves manipulating, calculating, filtering, or analyzing a selected range of data, use the "data_analyzer" tool.

If the user asks a question that cannot be answered by analyzing data (like a general knowledge question), you must respond with a JSON object where the "summary" explains that you cannot search the internet.

- User Request: "find the total of the 'Sales' column"
- Your Response: {"tool": "data_analyzer", "parameters": {"prompt": "find the total of the sales column"}}

- User Request: "what is the capital of France?"
- Your Response: {"tool": "none", "result": {"summary": "I am an Excel data agent and cannot search the internet for that information."}}
`;

// THIS HELPER FUNCTION NOW HAS ITS OWN SAFETY CHECK
async function runDataAnalyzer(data, prompt) {
  const model = genAI.getGenerativeModel({ model: "gemini-1.5-flash" });
  const DATA_ANALYZER_PROMPT = `You are an expert Excel data analyst. Your task is to analyze a given JSON dataset based on a user's plain-English request. You must return your response as a single, valid JSON object with two properties:
1. "summary": A one-sentence summary of the action you took.
2. "result": The result of the analysis. This can be a single value, an array, or an array of arrays.`;
  const fullPrompt = `${DATA_ANALYZER_PROMPT}\n\nDATASET:\n${JSON.stringify(data)}\n\nUSER REQUEST:\n"${prompt}"`;
  
  const result = await model.generateContent(fullPrompt);
  const response = await result.response;
  const responseText = response.text();

  // THE FIX IS HERE: Add a try/catch block inside the data analyzer as well.
  try {
    return JSON.parse(responseText);
  } catch (e) {
    // If the data analysis result isn't valid JSON, return it as a plain text summary.
    console.error("Data analyzer failed to return valid JSON. Returning plain text.");
    return { summary: responseText, result: "" };
  }
}

// --- Main Handler ---
module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') { return res.status(200).end(); }

  const { data, prompt } = req.body;

  try {
    const agentModel = genAI.getGenerativeModel({ model: "gemini-1.5-flash" });
    const decisionPrompt = `${AGENT_PROMPT}\n\nUser Prompt: "${prompt}"\nSelected Data Snippet: ${data ? JSON.stringify(data.slice(0, 3)) : "None"}`;
    const decisionResult = await agentModel.generateContent(decisionPrompt);
    const decisionResponse = await decisionResult.response;
    const responseText = decisionResponse.text();

    let toolChoice;
    try {
      toolChoice = JSON.parse(responseText);
    } catch (e) {
      toolChoice = { tool: 'none', result: { summary: responseText } };
    }

    let finalResult;
    if (toolChoice.tool === 'data_analyzer') {
      if (!data) throw new Error("Data analysis requires a selected data range in Excel.");
      finalResult = await runDataAnalyzer(data, toolChoice.parameters.prompt);
    } else {
      finalResult = toolChoice.result;
    }
    
    res.status(200).json(finalResult);

  } catch (error) {
    console.error('Error in agent handler:', error);
    res.status(500).json({ error: error.message });
  }
};
