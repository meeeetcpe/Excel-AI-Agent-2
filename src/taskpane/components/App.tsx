import * as React from 'react';
import { Button, Spinner, Label, makeStyles, shorthands, tokens, Textarea } from "@fluentui/react-components";
import { SendRegular } from "@fluentui/react-icons";

const useStyles = makeStyles({ /* ... existing styles are correct ... */ });

interface Message {
  sender: 'user' | 'agent';
  content: string | React.ReactNode;
}

const App = () => {
  const styles = useStyles();
  const [isLoading, setIsLoading] = React.useState(false);
  const [prompt, setPrompt] = React.useState("");
  const [history, setHistory] = React.useState<Message[]>([
    { sender: 'agent', content: "Hello! Select some data or ask me a question." }
  ]);
  const chatEndRef = React.useRef<HTMLDivElement>(null);

  React.useEffect(() => {
    chatEndRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [history, isLoading]);
  
  const handleSubmit = async () => {
    if (!prompt) return;

    const newUserMessage: Message = { sender: 'user', content: prompt };
    setHistory(prev => [...prev, newUserMessage]);
    setIsLoading(true);
    setPrompt("");

    try {
      await Excel.run(async (context) => {
        let selectedData = null;
        try {
            const range = context.workbook.getSelectedRange();
            range.load("values");
            await context.sync();
            selectedData = range.values;
        } catch (error) {
            console.log("No range selected or readable, proceeding without data.");
        }
        
        const response = await fetch("/api/agent-handler", {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ data: selectedData, prompt: prompt })
        });

        if (!response.ok) {
          const err = await response.json();
          throw new Error(err.error || "Server error");
        }
        const result = await response.json();

        // THIS IS THE CORRECTED DISPLAY LOGIC
        const agentThought: Message = { sender: 'agent', content: <em>Thinking... Used tool: {result.toolUsed || 'none'}</em> };
        
        let finalContent = result.summary || "I have completed the task.";
        
        const agentResponse: Message = { sender: 'agent', content: finalContent };
        
        // Add the thought and the final response to the chat history
        setHistory(prev => [...prev, agentThought, agentResponse]);
        
        // If there's a result to place in the sheet, do it now
        if (result.result) {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const resultRange = range ? range.getOffsetRange(range.rowCount, 0) : sheet.getRange("A1").getOffsetRange(1,0);
            
            if (Array.isArray(result.result) && Array.isArray(result.result[0])) {
              resultRange.getResizedRange(result.result.length - 1, result.result[0].length - 1).values = result.result;
            } else {
              resultRange.getCell(0, 0).values = [[result.result]];
            }
            await context.sync();
        }
      });
    } catch (error) {
      const errorMessage: Message = { sender: 'agent', content: `Error: ${error.message}` };
      setHistory(prev => [...prev, errorMessage]);
    } finally {
      setIsLoading(false);
    }
  };

  return (
    <div className={styles.root}>
        {/* ... existing JSX for chat history and input area is correct ... */}
    </div>
  );
};

export default App;
