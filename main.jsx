import React, { Suspense, lazy, useState } from "react";
import ReactDOM from "react-dom/client";
import { firebaseReady, parseFirebaseConfig } from "./firebase.js";
import "./index.css";

// La app se carga solo si Firebase está configurado (si no, importarla fallaría antes de poder avisar).
const App = lazy(() => import("./BarberApp.jsx"));

const CAMPOS = ["apiKey", "authDomain", "projectId", "storageBucket", "messagingSenderId", "appId"];

/** Pantalla de ayuda: aparece si a esta instalación aún no se le pegó la configuración de Firebase en Vercel. */
function Setup() {
  const [txt, setTxt] = useState("");
  const found = parseFirebaseConfig(txt);
  const box = { maxWidth: 440, margin: "0 auto", background: "#fff", border: "1px solid #ddd", borderRadius: 10, padding: 20 };
  const li = { marginBottom: 6, lineHeight: 1.4 };
  return (
    <div style={{ minHeight: "100vh", background: "#f3f1ec", padding: "24px 14px", fontFamily: "system-ui, sans-serif", color: "#1b1b1b" }}>
      <div style={box}>
        <h1 style={{ fontSize: 20, margin: "0 0 6px" }}>Falta conectar la base de datos</h1>
        <p style={{ fontSize: 14, color: "#555", margin: "0 0 14px" }}>
          Esta instalación todavía no sabe a qué base de datos de Firebase conectarse. Se arregla en 2 minutos desde Vercel:
        </p>
        <ol style={{ fontSize: 14, paddingLeft: 20, margin: 0 }}>
          <li style={li}>En <b>Firebase</b>: Configuración del proyecto → Tus apps → app web → copia el bloque que empieza con <code>const firebaseConfig = {"{"}</code></li>
          <li style={li}>En <b>Vercel</b>: abre este proyecto → <b>Settings → Environment Variables</b>.</li>
          <li style={li}>En <b>Key</b> escribe exactamente: <code>VITE_FIREBASE_CONFIG</code></li>
          <li style={li}>En <b>Value</b> pega todo el bloque que copiaste. Toca <b>Save</b>.</li>
          <li style={li}>Ve a <b>Deployments</b> → toca los tres puntos del último → <b>Redeploy</b>.</li>
        </ol>
        <p style={{ fontSize: 13, color: "#555", margin: "16px 0 6px" }}>
          ¿Quieres comprobar lo que copiaste antes de pegarlo en Vercel? Pégalo aquí y te digo si está completo:
        </p>
        <textarea value={txt} onChange={(e) => setTxt(e.target.value)} rows={5}
                  style={{ width: "100%", fontSize: 13, padding: 8, border: "1px solid #ccc", borderRadius: 6, boxSizing: "border-box" }}
                  placeholder="const firebaseConfig = { apiKey: ..." />
        {txt.trim() && (
          <ul style={{ listStyle: "none", padding: 0, margin: "8px 0 0", fontSize: 13 }}>
            {CAMPOS.map((c) => (
              <li key={c} style={{ color: found[c] ? "#1d7a46" : "#b3261e" }}>{found[c] ? "✓" : "✗ falta"} {c}</li>
            ))}
          </ul>
        )}
      </div>
    </div>
  );
}

ReactDOM.createRoot(document.getElementById("root")).render(
  <React.StrictMode>
    {firebaseReady ? <Suspense fallback={null}><App /></Suspense> : <Setup />}
  </React.StrictMode>
);
