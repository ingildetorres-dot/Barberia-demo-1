/**
 * Conexión con Firebase — proyecto "barberia-demo-1" (Demo Barbería).
 *
 * Estas llaves no son secretas: Firebase las diseñó para vivir en el código
 * que corre en el navegador de cualquiera. Lo que de verdad protege los datos
 * son las Reglas de Firestore (archivo firestore.rules), no este archivo.
 */
import { initializeApp } from "firebase/app";
import { getFirestore } from "firebase/firestore";

const firebaseConfig = {
  apiKey: "AIzaSyCCPQqPEFJSKZu40QztbTZpGa6HcAmg2fw",
  authDomain: "barberia-demo-1.firebaseapp.com",
  projectId: "barberia-demo-1",
  storageBucket: "barberia-demo-1.firebasestorage.app",
  messagingSenderId: "1044900108840",
  appId: "1:1044900108840:web:407191cd9fc1e4959ef31d",
};

export const db = getFirestore(initializeApp(firebaseConfig));
