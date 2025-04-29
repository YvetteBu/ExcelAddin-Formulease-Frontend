import { getAuth, createUserWithEmailAndPassword, signInWithEmailAndPassword, signOut } from "firebase/auth";
import { getFirestore, doc, setDoc } from "firebase/firestore";
import { initializeApp } from "firebase/app";

const firebaseConfig = {
  apiKey: "AIzaSyAd1UWFx31XeoG5JRedsQxuh2ZS1k6NzWo",
  authDomain: "excel-project-eeg.firebaseapp.com",
  projectId: "excel-project-eeg",
  storageBucket: "excel-project-eeg.firebasestorage.app",
};

const app = initializeApp(firebaseConfig);
const auth = getAuth(app);
const db = getFirestore(app);

const firebaseAuth = {
  async signUp(email, password, userData) {
    console.log("[firebaseAuth] signUp() called with:", email, userData);
    try {
      const userCredential = await createUserWithEmailAndPassword(auth, email, password);
      const user = userCredential.user;
      console.log("[firebaseAuth] signUp successful:", user.uid);

      await setDoc(doc(db, "users", user.uid), {
        name: userData.name,
        credits: 20
      });

      return { data: user, error: null };
    } catch (error) {
      console.error("[firebaseAuth] signUp error:", error);
      return { data: null, error: error.message };
    }
  },

  async initializeUserCredits(userId) {
    console.log("[firebaseAuth] initializeUserCredits() called for userId:", userId);
    try {
      await setDoc(doc(db, "credits", userId), {
        credits: 20
      });
      console.log("[firebaseAuth] Credits initialized");
      return { error: null };
    } catch (error) {
      console.error("[firebaseAuth] initializeUserCredits error:", error);
      return { error: error.message };
    }
  },

  async signIn(email, password) {
    console.log("[firebaseAuth] signIn() called with:", email);
    try {
      const userCredential = await signInWithEmailAndPassword(auth, email, password);
      const user = userCredential.user;
      console.log("[firebaseAuth] signIn successful:", user.uid);
      return { data: user, error: null };
    } catch (error) {
      console.error("[firebaseAuth] signIn error:", error);
      return { data: null, error: error.message };
    }
  },

  async signOut() {
    console.log("[firebaseAuth] signOut() called");
    try {
      await signOut(auth);
      console.log("[firebaseAuth] signOut successful");
      return { error: null };
    } catch (error) {
      console.error("[firebaseAuth] signOut error:", error);
      return { error: error.message };
    }
  }
};

export default firebaseAuth;