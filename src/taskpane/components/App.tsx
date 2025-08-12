import * as React from 'react';
import { Button, Spinner, Label, makeStyles, shorthands, tokens, Textarea } from "@fluentui/react-components";
import { SendRegular } from "@fluentui/react-icons";

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    height: "100vh",
  },
  chatHistory: {
    flexGrow: 1,
    overflowY: "auto",
    ...shorthands.padding("10px"),
    display: "flex",
    flexDirection: "column",
    ...shorthands.gap("15px"),
  },
  chatBubble: {
    maxWidth: "85%",
    ...shorthands.padding("8px", "12px"),
    ...shorthands.borderRadius(tokens.borderRadiusMedium),
    wordWrap: "break-word",
  },
  userBubble: {
    backgroundColor: tokens.colorBrandBackground,
    color: tokens.colorNeutralForegroundOnBrand,
    alignSelf: "flex-end",
  },
  agentBubble: {
    backgroundColor: tokens.colorNeutralBackground3,
    alignSelf: "flex-start",
  },
  inputArea: {
    display: "flex",
    ...shorthands.gap("10px"),
    ...shorthands.padding("10px"),
    ...shorthands.borderTop("1px", "solid", tokens.colorNeutralStroke2),
  },
  thinking: {
    ...shorthands.padding("10px"),
    display: "flex",
    alignItems: "center",
    ...shorthands.gap("10px"),
    color: tokens.colorNeutralForeground2,
  }
});

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

        const agentThought: Message = { sender: 'agent', content: <em>Thinking... Used tool: {result.toolUsed}</em> };
        
        let finalContent = `Summary: ${result.summary}`;
        if (result.keyPoints) {
            finalContent += "\n\nKey Points:\n" + result.keyPoints.map((p: string) => `• ${p}`).join('\n');
        } else if (result.result) {
            finalContent += `\n\nResult: ${JSON.stringify(result.result)}`;
        }
        const agentResponse: Message = { sender: 'agent', content: finalContent };
        
        setHistory(prev => [...prev, agentThought, agentResponse]);
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
      <div className={styles.chatHistory}>
        {history.map((msg, index) => (
          <div key={index} className={`${styles.chatBubble} ${msg.sender === 'user' ? styles.userBubble : styles.agentBubble}`}>
            <div style={{ whiteSpace: "pre-wrap" }}>{msg.content}</div>
          </div>
        ))}
        {isLoading && (
          <div className={styles.thinking}>
            <Spinner size="tiny"/> Thinking...
          </div>
        )}
        <div ref={chatEndRef} />
      </div>
      <div className={styles.inputArea}>
        <Textarea 
          value={prompt}
          onChange={(_, data) => setPrompt(data.value)}
          placeholder="Type your request here..."
          onKeyDown={(e) => { if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); handleSubmit(); } }}
          style={{resize: "none"}}
        />
        <Button icon={<SendRegular />} appearance="primary" onClick={handleSubmit} disabled={isLoading || !prompt} />
      </div>
    </div>
  );
};

export default App;
