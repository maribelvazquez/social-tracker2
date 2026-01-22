import { initializeApp } from "firebase/app";
import { getFirestore, collection, getDocs, addDoc, updateDoc, deleteDoc, doc, query, orderBy } from "firebase/firestore";

const firebaseConfig = {
  apiKey: "AIzaSyAUI6vhKXXO7Hpxjxx72xH_iLsdh-1LNQo",
  authDomain: "social-tracker360.firebaseapp.com",
  projectId: "social-tracker360",
  storageBucket: "social-tracker360.firebasestorage.app",
  messagingSenderId: "182243011028",
  appId: "1:182243011028:web:a410604502b81dfa6b8e41"
};

const app = initializeApp(firebaseConfig);
const db = getFirestore(app);

// Funciones para interactuar con Firestore
export const getRegistros = async () => {
  const q = query(collection(db, "registros"), orderBy("fecha", "asc"));
  const snapshot = await getDocs(q);
  return snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
};

export const addRegistro = async (data) => {
  // Usar la fecha como ID del documento
  const docRef = doc(db, "registros", data.fecha);
  const { setDoc } = await import("firebase/firestore");
  await setDoc(docRef, data);
  return { id: data.fecha, ...data };
};

export const updateRegistro = async (id, data) => {
  const docRef = doc(db, "registros", id);
  await updateDoc(docRef, data);
  return { id, ...data };
};

export const deleteRegistro = async (id) => {
  const docRef = doc(db, "registros", id);
  await deleteDoc(docRef);
};

export { db };
