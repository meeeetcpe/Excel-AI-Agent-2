import * as React from 'react';
import { Button, Input, Label, makeStyles, shorthands } from "@fluentui/react-components";

const useStyles = makeStyles({
  section: {
    display: "flex",
    flexDirection: "column",
    ...shorthands.gap("10px"),
  },
  centered: {
    alignSelf: "center",
    textAlign: "center"
  }
});

// FIX 1: Add 'isLoading: boolean' to the properties interface
interface ResearchComponentProps {
    isLoading: boolean;
    setIsLoading: (loading: boolean) => void;
    setStatus: (status: string) => void;
}

// FIX 2: Receive 'isLoading' from the props
export const ResearchComponent: React.FC<ResearchComponentProps> = ({ isLoading, setIsLoading, setStatus }) => {
    const styles = useStyles();
    const [query, setQuery] = React.useState("EV sector India 2025");
    
    const handleResearch = async () => {
        if (!query) {
            setStatus("Please enter a research topic.");
            return;
        }
        setIsLoading(true);
        setStatus("1/4: Starting web search...");

        try {
            const response = await fetch("/api/research", {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ query })
            });

            if (!response.ok) {
                const err = await response.json();
                throw new Error(err.error || "Server request failed.");
            }

            setStatus("2/4: Summarizing findings...");
            const result = await response.json();

            await Excel.run(async (context) => {
                setStatus("3/4: Creating new worksheet...");
                const sheetName = `Research_${new Date().getTime()}`;
                const sheet = context.workbook.worksheets.add(sheetName);
                sheet.activate();
                
                let currentRow = 1;

                sheet.getRange(`A${currentRow}`).values = [["AI Summary"]];
                sheet.getRange(`A${currentRow}`).format.font.bold = true;
                currentRow++;
                result.summaryBullets.forEach((bullet: string) => {
                    sheet.getRange(`A${currentRow++}`).values = [[`- ${bullet}`]];
                });

                currentRow += 2;

                sheet.getRange(`A${currentRow}`).values = [["Key Facts"]];
                sheet.getRange(`A${currentRow}`).format.font.bold = true;
                currentRow++;
                
                const tableHeader = [["Fact", "Source"]];
                const tableRange = sheet.getRange(`A${currentRow}`).getResizedRange(0, 1);
                tableRange.values = tableHeader;
                tableRange.format.font.bold = true;
                currentRow++;
                
                if (result.factsTable && result.factsTable.length > 0) {
                    const tableData = result.factsTable.map((fact: { Fact: string, Source: string }) => [fact.Fact, fact.Source]);
                    const dataRange = sheet.getRange(`A${currentRow}`).getResizedRange(tableData.length - 1, 1);
                    dataRange.values = tableData;
                }
                
                sheet.getUsedRange().format.autofitColumns();
                await context.sync();
                setStatus("Success! Research report created.");
            });

        } catch (error) {
            console.error(error);
            setStatus(`Error: ${error.message}`);
        } finally {
            setIsLoading(false);
        }
    };
    
    return (
        <div className={styles.section}>
            <Label size="large" weight="semibold" className={styles.centered}>Web Research</Label>
            <Input 
              value={query} 
              onChange={(e, data) => setQuery(data.value)}
              placeholder="e.g., Q2 earnings for major tech companies" 
            />
            <Button appearance="primary" disabled={!query || isLoading} onClick={handleResearch}>
              Research and Summarize
            </Button>
        </div>
    );
};