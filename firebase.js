/**
 * Conexión con Firebase — UNA instalación por negocio, un solo código para todos.
 *
 * La configuración de Firebase de cada negocio NO está escrita aquí: se guarda
 * en Vercel, en una variable llamada VITE_FIREBASE_CONFIG (Settings →
 * Environment Variables). Se pega TODO el bloque que Firebase muestra al
 * registrar la app web (el que empieza con "const firebaseConfig = {").
 *
 * Así el mismo repositorio sirve a todos tus clientes, cada uno con su propia
 * base de datos, y una mejora que subas llega a todos a la vez.
 *
 * Estas llaves no son secretas: lo que protege los datos son las Reglas de
 * Firestore y el inicio de sesión.
 */
import { initializeApp } from "firebase/app";
import { initializeFirestore } from "firebase/firestore";
import { getAuth } from "firebase/auth";

/** Saca las 6 llaves de lo que pegaste: acepta el bloque de Firebase tal cual, comillas simples o dobles, o JSON. */
export function parseFirebaseConfig(text) {
  const out = {};
  const re = /["']?(apiKey|authDomain|projectId|storageBucket|messagingSenderId|appId|measurementId)["']?\s*:\s*["']([^"']+)["']/g;
  let m;
  while ((m = re.exec(String(text || "")))) out[m[1]] = m[2].trim();
  return out;
}

const config = parseFirebaseConfig(import.meta.env.VITE_FIREBASE_CONFIG);
/** ¿Está esta instalación conectada a una base de datos? */
export const firebaseReady = !!(config.apiKey && config.projectId && config.appId);

let db = null;
let auth = null;
if (firebaseReady) {
  const app = initializeApp(config);
  // ignoreUndefinedProperties: un campo vacío nunca debe tumbar un guardado.
  db = initializeFirestore(app, { ignoreUndefinedProperties: true });
  auth = getAuth(app);
}
export { db, auth };
