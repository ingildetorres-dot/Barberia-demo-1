/**
 * Conexión con Firebase — completa esto con las llaves de TU proyecto.
 *
 * Cómo obtenerlas: console.firebase.google.com → crea un proyecto (p. ej.
 * "demo-barberia") → Compilación → Firestore Database → Crear base de datos
 * → luego Configuración del proyecto → Tus apps → </> (Web) → Registrar app.
 * Firebase te muestra un bloque "firebaseConfig": copia cada valor abajo,
 * entre las comillas.
 *
 * Estas llaves no son secretas: Firebase las diseñó para vivir en el código
 * que corre en el navegador de cualquiera. Lo que de verdad protege los datos
 * son las Reglas de Firestore (archivo firestore.rules), no este archivo.
 */
import { initializeApp } from "firebase/app";
import { getFirestore } from "firebase/firestore";

const firebaseConfig = {
  apiKey: "PEGA_AQUI_TU_apiKey",
  authDomain: "PEGA_AQUI_TU_authDomain",
  projectId: "PEGA_AQUI_TU_projectId",
  storageBucket: "PEGA_AQUI_TU_storageBucket",
  messagingSenderId: "PEGA_AQUI_TU_messagingSenderId",
  appId: "PEGA_AQUI_TU_appId",
};

export const db = getFirestore(initializeApp(firebaseConfig));
