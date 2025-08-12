const axios = require('axios');
const { GoogleGenerativeAI } = require("@google/generative-ai");

const genAI = new GoogleGenerativeAI(process.env.GEMINI_API_KEY);

const AGENT_PROMPT = `You are an AI agent inside Microsoft Excel. Your goal is to help the user with their request by choosing the correct tool. You must respond with a JSON object containing the name of the tool to use and any necessary parameters.

You have two tools available:

1.  **"data_analyzer"**: Use this tool when the user's request involves manipulating, calculating, filtering, or analyzing a selected range of data within Excel.
    * **Parameters**: "prompt" (the user's specific instruction).
    * **Example**: User says "find the total of the sales column". You respond with: {"tool": "data_analyzer", "parameters": {"prompt": "find the total of the sales column"}}

2.  **"internet_search"**: Use this tool when the user asks a general knowledge question or a question about current events that cannot be answered by the data in the spreadsheet.
    * **Parameters**: "query" (a search query generated from the user's request).
    * **Example**: User says "what were the latest earnings for Microsoft?". You respond with: {"tool": "internet_search", "parameters": {"query": "latest quarterly earnings for Microsoft"}}

Analyze the user's prompt and the available data to make your decision.`;

async function runDataAnalyzer(data, prompt) {
  const model = genAI.getGenerativeModel({ model: "gemini-1.5-flash" });
  const DATA_ANALYZER_PROMPT = `You are an expert Excel data analyst. Your task is to analyze a given JSON dataset based on a user's plain-English request. You must return your response as a single, valid JSON object with two properties:
1. "summary": A one-sentence summary of the action you took.
2. "result": The result of the analysis. This can be a single value (a number or text), or an array of values, or an array of arrays (if you are returning a new table or a filtered version of the original data).`;
  const fullPrompt = `${DATA_ANALYZER_PROMPT}\n\nDATASET:\n${JSON.stringify(data)}\n\nUSER REQUEST:\n"${prompt}"`;
  const result = await model.generateContent(fullPrompt);
  const response = await result.response;
  return JSON.parse(response.text());
}

async function runInternetSearch(query) {
  const model = genAI.getGenerativeModel({ model: "gemini-1.5-flash" });
  const braveResponse = await axios.get(`https://api.search.brave.com/res/v1/web/search`, {
    headers: { 'X-Subscription-Token': process.env.BRAVE_API_KEY },
    params: { q: query },
  });
  const searchResults = braveResponse.data.web.results.map(r => r.description).join('\n\n');
  const SUMMARIZATION_PROMPT = `You are an expert research analyst. Your task is to analyze the provided web search results and create a concise summary. You must return your response as a single, valid JSON object with two properties:
1. "summary": A detailed, multi-paragraph summary of the key findings from the search results.
2. "keyPoints": An array of 3-5 distinct bullet points highlighting the most important facts or figures.`;
  const fullPrompt = `${SUMMARIZATION_PROMPT}\n\nSEARCH RESULTS TO ANALYZE:\n${searchResults}`;
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
    const agentModel = genAI.getGenerativeModel({ model: "gemini-1.5-flash" });
    const decisionPrompt = `${AGENT_PROMPT}\n\nUser Prompt: "${prompt}"\nSelected Data Snippet: ${data ? JSON.stringify(data.slice(0, 3)) : "None"}`;
    const decisionResult = await agentModel.generateContent(decisionPrompt);
    const decisionResponse = await decisionResult.response;
    const toolChoice = JSON.parse(decisionResponse.text());

    let finalResult;
    if (toolChoice.tool === 'data_analyzer') {
      if (!data) throw new Error("Data analysis requires a selected data range.");
      finalResult = await runDataAnalyzer(data, toolChoice.parameters.prompt);
    } else if (toolChoice.tool === 'internet_search') {
      finalResult = await runInternetSearch(toolChoice.parameters.query);
    } else {
      throw new Error("Agent chose an invalid tool.");
    }
    
    finalResult.toolUsed = toolChoice.tool;
    res.status(200).json(finalResult);

  } catch (error) {
    console.error('Error in agent handler:', error);
    res.status(500).json({ error: error.message });
  }
};
