import App from "./components/App";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import * as React from "react";
import * as ReactDOM from "react-dom";

/* global document, Office, module, require */

const title = "Excel AI Agent";

const render = (Component) => {
  ReactDOM.render(
    <FluentProvider theme={webLightTheme}>
      {/* THE FIX IS ON THE LINE BELOW: Remove the 'title' property */}
      <Component />
    </FluentProvider>,
    document.getElementById("container")
  );
};

/* Render application after Office initializes */
Office.onReady(() => {
  render(App);
});

if ((module as any).hot) {
  (module as any).hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    render(NextApp);
  });
}
