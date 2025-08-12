import * as React from 'react';
import { Button, Input, Label, makeStyles, shorthands } from "@fluentui/react-components";

// ... useStyles and interface definitions remain the same ...
interface ResearchComponentProps {
    isLoading: boolean;
    setIsLoading: (loading: boolean) => void;
    setStatus: (status: string) => void;
}


export const ResearchComponent: React.FC<ResearchComponentProps> = ({ isLoading, setIsLoading, setStatus }) => {
    // ... all other code inside the component remains the same ...
    const styles = useStyles();
    const [query, setQuery] = React.useState("EV sector India 2025");
    const handleResearch = async () => { /* ... existing handleResearch code ... */ };

    return (
        <div className={styles.section}>
            <Label size="large" weight="semibold" className={styles.centered}>Web Research</Label>
            <Input 
              value={query} 
              // THE FIX IS ON THE LINE BELOW: 'e' is replaced with '_'
              onChange={(_, data) => setQuery(data.value)}
              placeholder="e.g., Q2 earnings for major tech companies" 
            />
            <Button appearance="primary" disabled={!query || isLoading} onClick={handleResearch}>
              Research and Summarize
            </Button>
        </div>
    );
};
