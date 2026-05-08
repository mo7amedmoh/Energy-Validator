import { initializeApp } from "firebase/app";
import { getFirestore } from "firebase/firestore";
import { getAuth } from "firebase/auth";
import { getAnalytics } from "firebase/analytics";

// Your web app's Firebase configuration
const firebaseConfig = {
  apiKey: "AIzaSyCfHFm3fSNlqE9c05IZYdq3LMeuuxsyWRA",
  authDomain: "energy-d8005.firebaseapp.com",
  projectId: "energy-d8005",
  storageBucket: "energy-d8005.firebasestorage.app",
  messagingSenderId: "1072538105599",
  appId: "1:1072538105599:web:6d091605131079757afeca",
  measurementId: "G-278R57BJXS",
};

// Initialize Firebase
const app = initializeApp(firebaseConfig);
export const db = getFirestore(app);
export const auth = getAuth(app);
export const analytics = getAnalytics(app);

export default app;
