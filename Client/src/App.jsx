import React from "react";
import InvoiceDashboard from "./components/InvoiceDashboard";
import AuthGate from "./components/AuthGate";

function App() {
  return (
    <AuthGate>
      <InvoiceDashboard />
    </AuthGate>
  );
}

export default App;
