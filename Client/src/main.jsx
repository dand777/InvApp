import React from "react";
import ReactDOM from "react-dom/client";
import { MsalProvider } from "@azure/msal-react";
import App from "./App";
import { ThemeProvider, createTheme } from "@mui/material/styles";
import CssBaseline from "@mui/material/CssBaseline";
import { msalInstance } from "./authConfig";

const darkTheme = createTheme({
  palette: {
    mode: "dark",
    primary: { main: "#1e88e5" },
    secondary: { main: "#f50057" },
    background: {
      default: "#121212",
      paper: "#1e1e2f",
    },
    text: {
      primary: "#ffffff",
      secondary: "#cfcfcf",
    },
  },
  typography: {
    fontFamily: '"Roboto", "Helvetica", "Arial", sans-serif',
    h4: { fontWeight: 700 },
  },
  components: {
    MuiButton: {
      styleOverrides: {
        root: {
          borderRadius: 10,
          textTransform: "none",
          padding: "8px 20px",
          fontWeight: 600,
        },
      },
    },
    MuiCard: {
      styleOverrides: {
        root: {
          borderRadius: 15,
          padding: "16px",
          backgroundColor: "#1e1e2f",
        },
      },
    },
    MuiDataGrid: {
      styleOverrides: {
        root: {
          border: "none",
          color: "#fff",
          backgroundColor: "#1e1e2f",
          borderRadius: 10,
        },
      },
    },
  },
});

const container = document.getElementById("root");

function renderApp() {
  if (!container) return;

  ReactDOM.createRoot(container).render(
    <MsalProvider instance={msalInstance}>
      <ThemeProvider theme={darkTheme}>
        <CssBaseline />
        <App />
      </ThemeProvider>
    </MsalProvider>
  );
}

msalInstance
  .initialize()
  .then(renderApp)
  .catch((error) => {
    if (import.meta.env.DEV) {
      console.error("MSAL initialization failed", error);
    }
    renderApp();
  });
