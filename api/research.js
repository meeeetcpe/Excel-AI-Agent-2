const axios = require('axios');
const { GoogleGenerativeAI } = require("@google/generative-ai");

const genAI = new GoogleGenerativeAI(process.env.GEMINI_API_KEY);
const BING_API_KEY = process.env.BING_SEARCH_V7_SUBSCRIPTION_KEY;
const BING_ENDPOINT = 'https://api.bing.microsoft.com/v7.0/search';

const RESEARCH_SUMMARIZATION_PROMPT = `You are a world-class financial and market analyst. Your task is to produce precise, sourced insights for an Excel report. Given the following text chunks from web pages, perform these actions:
1.  Generate a concise summary of 3-5 bullet points covering the main topics.
2.  Extract a table of key facts, each with its source URL.
3.  Return the output as a single, valid JSON object with two fields: "summaryBullets" (an array of strings), and "factsTable" (an array of objects, where each object has 'Fact' and 'Source' keys).
Do not include any preamble, explanation, or markdown backticks like \`\`\`json.`;

module.exports = async (req, res) => {
  // CORS Headers... (same as before)
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') { return res.status(200).end(); }
  
  const { query } = req.body;
  if (!query) { return res.status(400).json({ error: 'Search query is required' }); }

  try {
    // === Part 1: Call Bing Search API ===
    const searchResponse = await axios.get(BING_ENDPOINT, {
        headers: { 'Ocp-Apim-Subscription-Key': BING_API_KEY },
        params: { q: query, count: 5, responseFilter: 'Webpages' },
    });
    const searchResults = searchResponse.data.webPages.value;
    if (!searchResults || searchResults.length === 0) {
        return res.status(404).json({ error: 'No web results found.' });
    }

    // === Part 2: Fetch Content from Top URLs ===
    const pagePromises = searchResults.slice(0, 3).map(result => 
        axios.get(result.url, { timeout: 4000 })
            .then(pageResponse => {
                const textContent = pageResponse.data.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim();
                return { source: result.url, content: textContent.substring(0, 8000) };
            })
            .catch(fetchError => {
                console.warn(`Could not fetch ${result.url}: ${fetchError.message}`);
                return null;
            })
    );
    const pageContents = (await Promise.all(pagePromises)).filter(p => p !== null);
    if (pageContents.length === 0) {
        return res.status(500).json({ error: 'Failed to fetch content from any source.' });
    }

    // === Part 3: Call Gemini to Summarize ===
    const model = genAI.getGenerativeModel({ model: "gemini-1.5-flash" });
    const userContent = pageContents.map(p => `--- START OF ${p.source} ---\n${p.content}\n--- END OF ${p.source} ---`).join('\n\n');
    const fullPrompt = `${RESEARCH_SUMMARIZATION_PROMPT}\n\nHere is the source text:\n${userContent}`;

    const result = await model.generateContent(fullPrompt);
    const response = await result.response;
    const text = response.text();
    
    // Parse the JSON string from Gemini into a real object
    const summaryData = JSON.parse(text);

    res.status(200).json(summaryData);

  } catch (error) {
    console.error('Error during research process:', error);
    res.status(500).json({ error: 'Failed to complete research task' });
  }
};