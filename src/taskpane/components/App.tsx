import * as React from 'react';
import { Button, Spinner, Label, makeStyles, shorthands, tokens } from "@fluentui/react-components";

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    ...shorthands.gap("20px"),
    ...shorthands.padding("16px"),
  },
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

const App = () => {
  const styles = useStyles();
  const [isLoading, setIsLoading] = React.useState(false);
  const [status, setStatus] = React.useState("Ready.");

  const handleGenerateChart = async () => {
    setIsLoading(true);
    setStatus("1/3: Reading selected data...");
    try {
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load(["values", "address", "columnCount"]);
        await context.sync();
        
        setStatus("2/3: Generating chart narrative...");
        const response = await fetch("/api/generate-chart-narrative", {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ data: range.values })
        });

        if (!response.ok) {
          const err = await response.json();
          throw new Error(err.error || "Server error");
        }
        const result = await response.json();

        setStatus("3/3: Creating chart and inserting text...");
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        
        const chart = sheet.charts.add(Excel.ChartType.columnClustered, range, Excel.ChartSeriesBy.auto);
        chart.name = "AI_Generated_Chart";
        chart.title.text = "AI Generated Chart";

        const narrativeRange = range.getOffsetRange(0, range.columnCount + 1).getResizedRange(0, 1);
        narrativeRange.format.autofitColumns();
        narrativeRange.format.wrapText = true;
        narrativeRange.values = [[result.narrative]];

        await context.sync();
        setStatus("Success! Chart and narrative added.");
      });
    } catch (error) {
      console.error(error);
      setStatus(`Error: ${error.message}`);
    } finally {
      setIsLoading(false);
    }
  };

  return (
    <div className={styles.root}>
      <div className={styles.section}>
        <Label size="large" weight="semibold" className={styles.centered}>Chart Generation</Label>
        <p>1. Select a range of data.<br/>2. Click the button to generate a chart and an AI-powered summary.</p>
        <Button appearance="primary" disabled={isLoading} onClick={handleGenerateChart}>
          Generate Chart from Selection
        </Button>
      </div>

      {isLoading && (
        <div className={styles.section}>
          <Spinner labelPosition='below' label={status} />
        </div>
      )}
      {!isLoading && <p className={styles.centered} style={{ minHeight: "20px" }}>{status}</p>}
    </div>
  );
};

export default App;
