import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import { FluentProvider } from "@fluentui/react-provider";
import { webLightTheme } from "@fluentui/react-theme";
import App from "./App";

createRoot(document.getElementById("root")!).render(
  <StrictMode>
    <FluentProvider theme={webLightTheme}>
      <App />
    </FluentProvider>
  </StrictMode>
);
