import { useState, useEffect, useRef, useCallback } from "react";

// ─── jsPDF LOADER ────────────────────────────────────────────────────────────

function loadJsPDF() {
  return new Promise((resolve) => {
    if (window.jspdf) { resolve(window.jspdf.jsPDF); return; }
    const script = document.createElement("script");
    script.src = "https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js";
    script.onload = () => resolve(window.jspdf.jsPDF);
    document.head.appendChild(script);
  });
}

function loadXLSX() {
  return new Promise((resolve) => {
    if (window.XLSX) { resolve(window.XLSX); return; }
    const script = document.createElement("script");
    script.src = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
    script.onload = () => resolve(window.XLSX);
    document.head.appendChild(script);
  });
}

async function exportResultsToExcel(results, config) {
  const XLSX = await loadXLSX();
  const feedbackText = (r) => {
    const passed = r.percent >= config.passingScore;
    return passed
      ? `Félicitations ! Vous avez réussi avec ${r.percent}%.`
      : `Score de ${r.percent}% inférieur au seuil de ${config.passingScore}%. Réviser les notions non maîtrisées.`;
  };
  const rows = results.map(r => ({
    "Nom": r.student.nom,
    "Prénom": r.student.prenom,
    "Code Apogée": r.student.apogee,
    "Date": new Date(r.date).toLocaleString("fr-FR"),
    "Note obtenue": `${r.score}/${r.total} pts (${r.percent}%)`,
    "Feedback": feedbackText(r),
  }));
  const ws = XLSX.utils.json_to_sheet(rows);
  ws["!cols"] = [{ wch: 20 }, { wch: 20 }, { wch: 16 }, { wch: 22 }, { wch: 22 }, { wch: 60 }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Résultats");
  // Blob + anchor for reliable cross-browser download
  const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  const blob = new Blob([wbout], { type: "application/octet-stream" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "Resultats_EduExam.xlsx";
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

// ─── LOCAL DB (localStorage-based) ──────────────────────────────────────────

const DB = {
  getResults: () => JSON.parse(localStorage.getItem("examResults") || "[]"),
  saveResults: (arr) => localStorage.setItem("examResults", JSON.stringify(arr)),
  addResult: (result) => {
    const arr = DB.getResults();
    arr.push(result);
    DB.saveResults(arr);
  },
  getAttemptedApogees: () => {
    const results = DB.getResults();
    return new Set(results.map(r => r.student.apogee));
  },
  hasAttempted: (apogee) => DB.getAttemptedApogees().has(String(apogee).trim()),
  deleteResult: (apogee) => {
    const arr = DB.getResults().filter(r => r.student.apogee !== apogee);
    DB.saveResults(arr);
  },
  getPDFs: () => JSON.parse(localStorage.getItem("examPDFs") || "{}"),
  savePDF: (apogee, pdfDataUrl) => {
    const pdfs = DB.getPDFs();
    pdfs[apogee] = pdfDataUrl;
    localStorage.setItem("examPDFs", JSON.stringify(pdfs));
  },
  deletePDF: (apogee) => {
    const pdfs = DB.getPDFs();
    delete pdfs[apogee];
    localStorage.setItem("examPDFs", JSON.stringify(pdfs));
  },
};

// ─── PDF GENERATOR ──────────────────────────────────────────────────────────

async function generateAndStorePDF(result) {
  const jsPDF = await loadJsPDF();
  const { student, score, total, percent, timeUsed, date, questions, answers, config } = result;
  const passed = percent >= config.passingScore;
  const doc = new jsPDF({ orientation: "portrait", unit: "mm", format: "a4" });

  const W = 210, margin = 18;
  let y = 0;

  // Header band
  doc.setFillColor(15, 23, 42);
  doc.rect(0, 0, W, 42, "F");
  doc.setFillColor(37, 99, 235);
  doc.rect(0, 42, W, 3, "F");

  doc.setTextColor(255, 255, 255);
  doc.setFontSize(18);
  doc.setFont("helvetica", "bold");
  doc.text("EduExam Pro", margin, 16);
  doc.setFontSize(11);
  doc.setFont("helvetica", "normal");
  doc.setTextColor(148, 163, 184);
  doc.text(config.title || "Évaluation", margin, 25);
  doc.text(config.subtitle || "Université Hassan II — Casablanca", margin, 32);

  // Result badge (top right)
  doc.setFillColor(passed ? 21 : 185, passed ? 128 : 28, passed ? 61 : 28);
  doc.roundedRect(W - margin - 38, 10, 38, 22, 4, 4, "F");
  doc.setTextColor(255, 255, 255);
  doc.setFontSize(10);
  doc.setFont("helvetica", "bold");
  doc.text(passed ? "ADMIS(E)" : "AJOURNÉ(E)", W - margin - 19, 20, { align: "center" });
  doc.setFontSize(14);
  doc.text(`${percent}%`, W - margin - 19, 29, { align: "center" });

  y = 56;

  // Student info block
  doc.setFillColor(248, 250, 255);
  doc.roundedRect(margin, y, W - margin * 2, 38, 4, 4, "F");
  doc.setDrawColor(232, 237, 245);
  doc.roundedRect(margin, y, W - margin * 2, 38, 4, 4, "S");

  doc.setTextColor(100, 116, 139);
  doc.setFontSize(8);
  doc.setFont("helvetica", "bold");
  doc.text("INFORMATIONS ÉTUDIANT", margin + 6, y + 8);

  doc.setTextColor(30, 41, 59);
  doc.setFontSize(10);
  doc.setFont("helvetica", "normal");
  const col2 = margin + 90;
  doc.setFont("helvetica", "bold");
  doc.text("Nom :", margin + 6, y + 17);
  doc.text("Prénom :", margin + 6, y + 25);
  doc.text("Code Apogée :", col2, y + 17);
  doc.text("Date de soumission :", col2, y + 25);

  doc.setFont("helvetica", "normal");
  doc.text(student.nom, margin + 26, y + 17);
  doc.text(student.prenom, margin + 26, y + 25);
  doc.text(student.apogee, col2 + 32, y + 17);
  doc.text(new Date(date).toLocaleString("fr-FR"), col2 + 46, y + 25);

  // Temps
  doc.setFont("helvetica", "bold");
  doc.text("Temps utilisé :", margin + 6, y + 33);
  doc.setFont("helvetica", "normal");
  doc.text(formatTime(timeUsed), margin + 36, y + 33);

  y += 48;

  // Score section
  doc.setFillColor(passed ? 240 : 254, passed ? 253 : 242, passed ? 244 : 242);
  doc.roundedRect(margin, y, W - margin * 2, 32, 4, 4, "F");
  doc.setDrawColor(passed ? 74 : 248, passed ? 222 : 113, passed ? 128 : 113);
  doc.roundedRect(margin, y, W - margin * 2, 32, 4, 4, "S");

  doc.setTextColor(100, 116, 139);
  doc.setFontSize(8);
  doc.setFont("helvetica", "bold");
  doc.text("RÉSULTAT FINAL", margin + 6, y + 8);

  doc.setFontSize(22);
  doc.setFont("helvetica", "bold");
  doc.setTextColor(passed ? 21 : 185, passed ? 128 : 28, passed ? 61 : 28);
  doc.text(`${score} / ${total} pts`, margin + 6, y + 24);

  doc.setFontSize(10);
  doc.setTextColor(100, 116, 139);
  doc.setFont("helvetica", "normal");
  const correct = questions.filter((q, i) => answers[i] === q.answer).length;
  doc.text(`${correct} réponses correctes sur ${questions.length} questions`, W - margin - 6, y + 16, { align: "right" });
  doc.text(`Seuil de réussite : ${config.passingScore}%`, W - margin - 6, y + 24, { align: "right" });

  y += 42;

  // Feedback
  if (config.showFeedback) {
    const feedbackText = passed
      ? `Félicitations ! Vous avez réussi cette évaluation avec un score de ${percent}%. Votre maîtrise des concepts est validée.`
      : `Votre score de ${percent}% est inférieur au seuil de ${config.passingScore}% requis. Nous vous encourageons à revoir les notions non maîtrisées et à retenter l'évaluation.`;

    doc.setFillColor(239, 246, 255);
    doc.roundedRect(margin, y, W - margin * 2, 20, 4, 4, "F");
    doc.setDrawColor(147, 197, 253);
    doc.roundedRect(margin, y, W - margin * 2, 20, 4, 4, "S");
    doc.setTextColor(29, 78, 216);
    doc.setFontSize(8);
    doc.setFont("helvetica", "bold");
    doc.text("💡 FEEDBACK", margin + 6, y + 8);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(9);
    doc.setTextColor(30, 41, 59);
    const lines = doc.splitTextToSize(feedbackText, W - margin * 2 - 12);
    doc.text(lines, margin + 6, y + 15);
    y += 28;
  }

  // Detail header
  doc.setFontSize(11);
  doc.setFont("helvetica", "bold");
  doc.setTextColor(30, 41, 59);
  doc.text("Détail des réponses", margin, y + 8);
  doc.setDrawColor(226, 232, 240);
  doc.line(margin, y + 11, W - margin, y + 11);
  y += 16;

  // Question rows
  questions.forEach((q, i) => {
    if (y > 260) {
      doc.addPage();
      y = 20;
    }
    const isCorrect = answers[i] === q.answer;
    const studentAns = answers[i] !== undefined ? q.options[answers[i]] : "Sans réponse";
    const correctAns = q.options[q.answer];

    doc.setFillColor(isCorrect ? 240 : 255, isCorrect ? 253 : 245, isCorrect ? 244 : 245);
    doc.roundedRect(margin, y, W - margin * 2, 24, 3, 3, "F");
    doc.setDrawColor(isCorrect ? 74 : 248, isCorrect ? 222 : 113, isCorrect ? 128 : 113);
    doc.roundedRect(margin, y, W - margin * 2, 24, 3, 3, "S");

    // Q number badge
    doc.setFillColor(isCorrect ? 21 : 185, isCorrect ? 128 : 28, isCorrect ? 61 : 28);
    doc.roundedRect(margin + 3, y + 4, 14, 16, 2, 2, "F");
    doc.setTextColor(255, 255, 255);
    doc.setFontSize(9);
    doc.setFont("helvetica", "bold");
    doc.text(`Q${i + 1}`, margin + 10, y + 14, { align: "center" });

    // Category & score
    doc.setTextColor(100, 116, 139);
    doc.setFontSize(8);
    doc.setFont("helvetica", "normal");
    doc.text(q.category, margin + 22, y + 9);
    doc.setFont("helvetica", "bold");
    doc.setTextColor(isCorrect ? 21 : 185, isCorrect ? 128 : 28, isCorrect ? 61 : 28);
    doc.text(`${isCorrect ? q.score : 0}/${q.score} pts`, W - margin - 6, y + 9, { align: "right" });

    // Question text (truncated)
    const qText = q.text.split("\n")[0].replace(/\$[^$]+\$/g, "[formule]").substring(0, 70) + "…";
    doc.setTextColor(30, 41, 59);
    doc.setFontSize(8);
    doc.setFont("helvetica", "normal");
    doc.text(qText, margin + 22, y + 16);

    // Answers
    doc.setFontSize(7.5);
    doc.setTextColor(100, 116, 139);
    if (!isCorrect) {
      doc.text(`Votre réponse : ${studentAns}`, margin + 22, y + 22);
      doc.setTextColor(21, 128, 61);
      doc.text(`Bonne réponse : ${correctAns}`, margin + 95, y + 22);
    } else {
      doc.setTextColor(21, 128, 61);
      doc.text(`✓ ${correctAns}`, margin + 22, y + 22);
    }

    y += 28;
  });

  // Footer
  if (y > 260) { doc.addPage(); y = 20; }
  y += 6;
  doc.setDrawColor(226, 232, 240);
  doc.line(margin, y, W - margin, y);
  y += 6;
  doc.setFontSize(8);
  doc.setTextColor(148, 163, 184);
  doc.setFont("helvetica", "normal");
  doc.text(`Généré automatiquement par EduExam Pro — ${new Date().toLocaleString("fr-FR")}`, margin, y);
  doc.text("Document confidentiel", W - margin, y, { align: "right" });

  // Save to localStorage only (no auto-download)
  const pdfDataUrl = doc.output("datauristring");
  DB.savePDF(student.apogee, pdfDataUrl);

  return pdfDataUrl;
}

// ─── DATA ───────────────────────────────────────────────────────────────────

const DEFAULT_CONFIG = {
  title: "Évaluation TP — Master Big Data & IA",
  subtitle: "Université Hassan II — Casablanca",
  duration: 60,
  questionCount: 10,
  categories: ["Python", "Machine Learning", "Big Data", "Statistiques", "Deep Learning"],
  shuffleQuestions: true,
  shuffleAnswers: true,
  scorePerQuestion: 2,
  passingScore: 50,
  showFeedback: true,
  tabSwitchWarnings: 3,
};

const QUESTION_BANK = [
  {
    id: 1, category: "Python", difficulty: "Facile", score: 2,
    type: "qcm",
    text: "Quelle est la complexité temporelle d'un accès à un élément dans un dictionnaire Python ?",
    options: ["O(n)", "O(log n)", "O(1)", "O(n²)"],
    answer: 2,
    explanation: "Les dictionnaires Python utilisent une table de hachage, garantissant un accès moyen en O(1).",
  },
  {
    id: 2, category: "Python", difficulty: "Moyen", score: 2,
    type: "qcm",
    text: "Quel est le résultat de l'expression suivante en Python ?\n```python\nx = [1, 2, 3]\ny = x\ny.append(4)\nprint(x)\n```",
    options: ["[1, 2, 3]", "[1, 2, 3, 4]", "Erreur", "None"],
    answer: 1,
    explanation: "y est une référence vers x. Modifier y modifie également x car ils pointent vers le même objet.",
  },
  {
    id: 3, category: "Machine Learning", difficulty: "Moyen", score: 2,
    type: "qcm",
    text: "Dans le contexte de la régression logistique, quelle fonction est utilisée comme fonction d'activation ?",
    options: ["ReLU", "Tanh", "Sigmoïde", "Softmax"],
    answer: 2,
    explanation: "La régression logistique utilise la fonction sigmoïde σ(x) = 1/(1+e^{-x}) pour mapper les valeurs vers [0,1].",
  },
  {
    id: 4, category: "Machine Learning", difficulty: "Difficile", score: 2,
    type: "math",
    text: "Soit un modèle de régression linéaire. L'erreur quadratique moyenne (MSE) est définie par :\n$$MSE = \\frac{1}{n}\\sum_{i=1}^{n}(y_i - \\hat{y}_i)^2$$\nSi on a les valeurs réelles [2, 4, 6] et prédites [3, 3, 5], quel est le MSE ?",
    options: ["1.0", "0.67", "1.33", "2.0"],
    answer: 2,
    explanation: "MSE = ((2-3)² + (4-3)² + (6-5)²) / 3 = (1+1+1)/3 = 1.0. Calcul : (1+1+1)/3 = 1.0. Réponse correcte : 1.0",
  },
  {
    id: 5, category: "Big Data", difficulty: "Facile", score: 2,
    type: "qcm",
    text: "Quel est le rôle du 'NameNode' dans l'architecture HDFS ?",
    options: [
      "Stocker les données réelles",
      "Gérer les métadonnées du système de fichiers",
      "Effectuer les calculs MapReduce",
      "Gérer la sécurité du cluster",
    ],
    answer: 1,
    explanation: "Le NameNode gère les métadonnées HDFS (structure des répertoires, localisation des blocs). Les données sont stockées sur les DataNodes.",
  },
  {
    id: 6, category: "Big Data", difficulty: "Moyen", score: 2,
    type: "qcm",
    text: "Dans Apache Spark, quelle est la différence fondamentale entre un RDD et un DataFrame ?",
    options: [
      "Un DataFrame est distribué, un RDD ne l'est pas",
      "Un RDD est typé, un DataFrame ne l'est pas",
      "Un DataFrame possède un schéma structuré, un RDD est non structuré",
      "Il n'y a pas de différence",
    ],
    answer: 2,
    explanation: "Les DataFrames ont un schéma défini (colonnes typées), permettant des optimisations via Catalyst. Les RDDs sont des collections distribuées non structurées.",
  },
  {
    id: 7, category: "Statistiques", difficulty: "Moyen", score: 2,
    type: "math",
    text: "Pour une distribution normale $X \\sim \\mathcal{N}(\\mu, \\sigma^2)$, quelle proportion des données se trouve dans l'intervalle $[\\mu - 2\\sigma, \\mu + 2\\sigma]$ ?",
    options: ["68.3%", "95.4%", "99.7%", "90.0%"],
    answer: 1,
    explanation: "La règle empirique 68-95-99.7 : 68.3% dans ±1σ, 95.4% dans ±2σ, 99.7% dans ±3σ.",
  },
  {
    id: 8, category: "Deep Learning", difficulty: "Difficile", score: 2,
    type: "qcm",
    text: "Dans un réseau de neurones convolutif (CNN), à quoi sert la couche de 'pooling' ?",
    options: [
      "Augmenter la résolution de l'image",
      "Ajouter de la non-linéarité",
      "Réduire les dimensions spatiales et contrôler l'overfitting",
      "Normaliser les activations",
    ],
    answer: 2,
    explanation: "Le pooling réduit les dimensions spatiales (sous-échantillonnage), diminue le nombre de paramètres et apporte une invariance aux translations.",
  },
  {
    id: 9, category: "Python", difficulty: "Moyen", score: 2,
    type: "code",
    text: "Analysez ce code et identifiez sa sortie :\n```python\nimport numpy as np\nA = np.array([[1, 2], [3, 4]])\nB = np.array([[5, 6], [7, 8]])\nC = A @ B\nprint(C[0][1])\n```",
    options: ["19", "28", "12", "22"],
    answer: 1,
    explanation: "A @ B est le produit matriciel. C[0][1] = 1×6 + 2×8 = 6 + 16 = 22. Correction : C[0] = [1×5+2×7, 1×6+2×8] = [19, 22], donc C[0][1] = 22.",
  },
  {
    id: 10, category: "Machine Learning", difficulty: "Facile", score: 2,
    type: "qcm",
    text: "Quelle technique permet d'éviter l'overfitting en ajoutant une pénalité sur les poids dans la fonction de coût ?",
    options: ["Dropout", "Batch Normalization", "Régularisation L2 (Ridge)", "Data Augmentation"],
    answer: 2,
    explanation: "La régularisation L2 (Ridge) ajoute λ||w||² à la fonction de coût, pénalisant les poids élevés et réduisant l'overfitting.",
  },
  {
    id: 11, category: "Deep Learning", difficulty: "Difficile", score: 2,
    type: "math",
    text: "La fonction d'activation ReLU est définie par $f(x) = \\max(0, x)$. Quelle est sa dérivée pour $x > 0$ ?",
    options: ["0", "1", "x", "e^x"],
    answer: 1,
    explanation: "Pour x > 0, ReLU(x) = x, donc f'(x) = 1. Pour x < 0, f'(x) = 0. C'est un avantage car le gradient ne disparaît pas pour les valeurs positives.",
  },
  {
    id: 12, category: "Big Data", difficulty: "Difficile", score: 2,
    type: "qcm",
    text: "Dans Apache Kafka, qu'est-ce qu'un 'consumer group' ?",
    options: [
      "Un groupe de producteurs partageant le même topic",
      "Un ensemble de consommateurs qui se répartissent les partitions d'un topic",
      "Un cluster de brokers Kafka",
      "Un mécanisme de réplication des données",
    ],
    answer: 1,
    explanation: "Un consumer group permet à plusieurs consommateurs de lire en parallèle depuis différentes partitions d'un topic, assurant un traitement distribué et équilibré.",
  },
  {
    id: 13, category: "Statistiques", difficulty: "Facile", score: 2,
    type: "qcm",
    text: "Quelle mesure statistique est résistante aux valeurs aberrantes (outliers) ?",
    options: ["La moyenne arithmétique", "La variance", "La médiane", "L'écart-type"],
    answer: 2,
    explanation: "La médiane est robuste aux outliers car elle représente la valeur centrale, indépendamment des valeurs extrêmes.",
  },
  {
    id: 14, category: "Python", difficulty: "Difficile", score: 2,
    type: "code",
    text: "Quelle est la complexité de ce code ?\n```python\ndef find_duplicates(lst):\n    seen = set()\n    duplicates = []\n    for item in lst:\n        if item in seen:\n            duplicates.append(item)\n        else:\n            seen.add(item)\n    return duplicates\n```",
    options: ["O(n²)", "O(n log n)", "O(n)", "O(1)"],
    answer: 2,
    explanation: "La vérification 'item in seen' dans un set est O(1) en moyenne. La boucle itère n fois → complexité globale O(n).",
  },
  {
    id: 15, category: "Machine Learning", difficulty: "Difficile", score: 2,
    type: "math",
    text: "L'algorithme du gradient descent met à jour les poids selon :\n$$w_{t+1} = w_t - \\eta \\nabla L(w_t)$$\nSi $\\eta = 0.01$, $w_t = 3.0$ et $\\nabla L = 2.5$, quelle est la valeur de $w_{t+1}$ ?",
    options: ["2.975", "3.025", "2.95", "3.0"],
    answer: 0,
    explanation: "w_{t+1} = 3.0 - 0.01 × 2.5 = 3.0 - 0.025 = 2.975",
  },
];

const ADMIN_CREDENTIALS = { username: "tayebi", password: "tayebi@1992" };

// ─── UTILS ──────────────────────────────────────────────────────────────────

function shuffle(arr) {
  const a = [...arr];
  for (let i = a.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [a[i], a[j]] = [a[j], a[i]];
  }
  return a;
}

function formatTime(seconds) {
  const m = Math.floor(seconds / 60).toString().padStart(2, "0");
  const s = (seconds % 60).toString().padStart(2, "0");
  return `${m}:${s}`;
}

function useKaTeX() {
  const [ready, setReady] = useState(typeof window !== "undefined" && !!window.katex);
  useEffect(() => {
    if (window.katex) { setReady(true); return; }
    const check = setInterval(() => { if (window.katex) { setReady(true); clearInterval(check); } }, 100);
    return () => clearInterval(check);
  }, []);
  return ready;
}

function KaTeXInline({ latex }) {
  const ref = useRef(null);
  const ready = useKaTeX();
  useEffect(() => {
    if (ref.current && ready && window.katex) {
      try { window.katex.render(latex, ref.current, { throwOnError: false, displayMode: false }); }
      catch (e) { ref.current.textContent = latex; }
    }
  }, [latex, ready]);
  return <span ref={ref} style={{ fontFamily: "KaTeX_Main, serif" }} />;
}

function KaTeXBlock({ latex }) {
  const ref = useRef(null);
  const ready = useKaTeX();
  useEffect(() => {
    if (ref.current && ready && window.katex) {
      try { window.katex.render(latex, ref.current, { throwOnError: false, displayMode: true }); }
      catch (e) { ref.current.textContent = latex; }
    }
  }, [latex, ready]);
  return <div ref={ref} style={{ margin: "14px 0", overflowX: "auto" }} />;
}

function renderText(text) {
  if (!text) return "";
  const parts = text.split(/(```[\s\S]*?```|\$\$[\s\S]*?\$\$|\$[^$\n]+\$)/g);
  return parts.map((part, i) => {
    if (part.startsWith("```")) {
      const code = part.replace(/```\w*\n?/, "").replace(/```$/, "");
      return (
        <div key={i} style={{
          background: "#0f172a", borderRadius: 8, padding: "12px 16px",
          fontFamily: "'Fira Code', 'Courier New', monospace", fontSize: 13,
          color: "#e2e8f0", margin: "10px 0", overflowX: "auto",
          border: "1px solid rgba(255,255,255,0.08)", lineHeight: 1.6,
        }}>
          {code.trim().split("\n").map((line, li) => <div key={li}>{line}</div>)}
        </div>
      );
    }
    if (part.startsWith("$$")) return <KaTeXBlock key={i} latex={part.slice(2, -2)} />;
    if (part.startsWith("$")) return <KaTeXInline key={i} latex={part.slice(1, -1)} />;
    return <span key={i}>{part}</span>;
  });
}

// ─── STYLES ─────────────────────────────────────────────────────────────────

const CSS = `
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700&family=Fira+Code:wght@400;500&display=swap');

*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

body {
  font-family: 'Plus Jakarta Sans', sans-serif;
  background: #f8faff;
  color: #1e293b;
  min-height: 100vh;
}

::-webkit-scrollbar { width: 6px; }
::-webkit-scrollbar-track { background: #f1f5f9; }
::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 99px; }

.fade-in { animation: fadeIn 0.5s ease; }
@keyframes fadeIn { from { opacity: 0; transform: translateY(16px); } to { opacity: 1; transform: none; } }

.slide-in { animation: slideIn 0.4s cubic-bezier(0.34,1.56,0.64,1); }
@keyframes slideIn { from { opacity: 0; transform: scale(0.96); } to { opacity: 1; transform: scale(1); } }

.pulse { animation: pulse 2s infinite; }
@keyframes pulse { 0%,100% { opacity:1; } 50% { opacity:0.6; } }

.btn {
  display: inline-flex; align-items: center; gap: 8px;
  padding: 10px 20px; border-radius: 10px; border: none;
  font-family: inherit; font-size: 14px; font-weight: 600;
  cursor: pointer; transition: all 0.2s;
  text-decoration: none; white-space: nowrap;
}
.btn:active { transform: scale(0.97); }
.btn-primary { background: #2563eb; color: #fff; }
.btn-primary:hover { background: #1d4ed8; box-shadow: 0 4px 14px rgba(37,99,235,0.35); }
.btn-secondary { background: #f1f5f9; color: #475569; border: 1px solid #e2e8f0; }
.btn-secondary:hover { background: #e2e8f0; }
.btn-danger { background: #fee2e2; color: #dc2626; }
.btn-danger:hover { background: #fecaca; }
.btn-success { background: #dcfce7; color: #16a34a; }
.btn-success:hover { background: #bbf7d0; }
.btn-lg { padding: 14px 28px; font-size: 16px; border-radius: 12px; }
.btn-sm { padding: 7px 14px; font-size: 13px; border-radius: 8px; }

.card {
  background: #fff; border-radius: 16px;
  border: 1px solid #e8edf5; padding: 24px;
  box-shadow: 0 1px 3px rgba(0,0,0,0.04);
}

.badge {
  display: inline-flex; align-items: center; gap: 4px;
  padding: 3px 10px; border-radius: 99px; font-size: 12px; font-weight: 600;
}
.badge-blue { background: #dbeafe; color: #1d4ed8; }
.badge-green { background: #dcfce7; color: #15803d; }
.badge-amber { background: #fef3c7; color: #b45309; }
.badge-red { background: #fee2e2; color: #b91c1c; }
.badge-purple { background: #f3e8ff; color: #7c3aed; }

input, select, textarea {
  font-family: inherit; font-size: 14px;
  border: 1.5px solid #e2e8f0; border-radius: 10px;
  padding: 10px 14px; width: 100%; background: #fff;
  color: #1e293b; transition: border-color 0.2s, box-shadow 0.2s;
  outline: none;
}
input:focus, select:focus, textarea:focus {
  border-color: #2563eb;
  box-shadow: 0 0 0 3px rgba(37,99,235,0.12);
}
label { font-size: 13px; font-weight: 600; color: #475569; margin-bottom: 6px; display: block; }

.progress-bar { height: 6px; background: #e2e8f0; border-radius: 99px; overflow: hidden; }
.progress-fill { height: 100%; background: linear-gradient(90deg, #2563eb, #7c3aed); border-radius: 99px; transition: width 0.5s ease; }

.sidebar { width: 280px; min-width: 280px; background: #fff; border-right: 1px solid #e8edf5; height: 100vh; overflow-y: auto; position: sticky; top: 0; }

.q-bubble {
  width: 36px; height: 36px; border-radius: 8px;
  display: flex; align-items: center; justify-content: center;
  font-size: 12px; font-weight: 700; cursor: pointer;
  transition: all 0.15s; border: 2px solid transparent;
  flex-shrink: 0;
}
.q-bubble.not-visited { background: #f1f5f9; color: #94a3b8; }
.q-bubble.visited { background: #fef3c7; color: #b45309; border-color: #fbbf24; }
.q-bubble.answered { background: #dcfce7; color: #15803d; border-color: #4ade80; }
.q-bubble.current { background: #dbeafe; color: #1d4ed8; border-color: #3b82f6; }
.q-bubble:hover { transform: scale(1.08); }

.timer-danger { color: #dc2626 !important; animation: timerPulse 1s infinite; }
@keyframes timerPulse { 0%,100%{opacity:1} 50%{opacity:0.5} }

.stat-card {
  background: #fff; border-radius: 14px; padding: 20px;
  border: 1px solid #e8edf5; display: flex; flex-direction: column; gap: 6px;
}
.stat-number { font-size: 28px; font-weight: 700; color: #1e293b; }
.stat-label { font-size: 13px; color: #94a3b8; font-weight: 500; }

.tab { padding: 8px 16px; border-radius: 8px; font-size: 14px; font-weight: 600; cursor: pointer; border: none; background: transparent; color: #64748b; transition: all 0.2s; }
.tab.active { background: #2563eb; color: #fff; }
.tab:hover:not(.active) { background: #f1f5f9; color: #1e293b; }

.modal-backdrop {
  position: fixed; inset: 0; background: rgba(15,23,42,0.5);
  display: flex; align-items: center; justify-content: center; z-index: 1000;
  backdrop-filter: blur(4px);
}
.modal { background: #fff; border-radius: 20px; padding: 32px; max-width: 520px; width: 90%; box-shadow: 0 24px 64px rgba(0,0,0,0.15); }

.table-container { overflow-x: auto; }
table { width: 100%; border-collapse: collapse; }
th { text-align: left; font-size: 12px; font-weight: 700; color: #94a3b8; text-transform: uppercase; letter-spacing: 0.05em; padding: 10px 12px; background: #f8faff; }
td { padding: 12px; font-size: 14px; color: #334155; border-top: 1px solid #f1f5f9; }
tr:hover td { background: #f8faff; }

.hero-bg {
  background: linear-gradient(135deg, #0f172a 0%, #1e3a6e 50%, #0f172a 100%);
  position: relative; overflow: hidden;
}
.hero-bg::before {
  content: ''; position: absolute; inset: 0;
  background: radial-gradient(ellipse at 30% 50%, rgba(37,99,235,0.3) 0%, transparent 60%),
              radial-gradient(ellipse at 70% 30%, rgba(124,58,237,0.2) 0%, transparent 50%);
}
.hero-grid {
  background-image: linear-gradient(rgba(255,255,255,0.03) 1px, transparent 1px),
                    linear-gradient(90deg, rgba(255,255,255,0.03) 1px, transparent 1px);
  background-size: 60px 60px;
  position: absolute; inset: 0;
}

.option-card {
  border: 2px solid #e2e8f0; border-radius: 12px; padding: 14px 18px;
  cursor: pointer; transition: all 0.2s; display: flex; align-items: center; gap: 14px;
  background: #fff; font-size: 15px;
}
.option-card:hover { border-color: #93c5fd; background: #eff6ff; }
.option-card.selected { border-color: #2563eb; background: #eff6ff; }
.option-card.correct { border-color: #4ade80; background: #f0fdf4; }
.option-card.incorrect { border-color: #f87171; background: #fef2f2; }

.nav-admin { background: #fff; border-bottom: 1px solid #e8edf5; padding: 0 24px; height: 64px; display: flex; align-items: center; justify-content: space-between; position: sticky; top: 0; z-index: 100; }

.notification {
  position: fixed; top: 20px; right: 20px; z-index: 9999;
  background: #1e293b; color: #fff; padding: 12px 20px;
  border-radius: 12px; font-size: 14px; font-weight: 500;
  box-shadow: 0 8px 24px rgba(0,0,0,0.2);
  animation: slideInRight 0.3s ease;
  max-width: 320px;
}
@keyframes slideInRight { from { opacity:0; transform: translateX(20px); } to { opacity:1; transform: none; } }

.search-input {
  background: #f8faff;
  border: 1.5px solid #e2e8f0;
  border-radius: 10px;
  padding: 9px 14px 9px 38px;
  font-size: 14px;
  width: 260px;
  transition: all 0.2s;
}
.search-input:focus {
  border-color: #2563eb;
  box-shadow: 0 0 0 3px rgba(37,99,235,0.10);
  background: #fff;
  outline: none;
}
`;

// ─── NOTIFICATION ────────────────────────────────────────────────────────────

function Notification({ message, type, onClose }) {
  useEffect(() => {
    const t = setTimeout(onClose, 3500);
    return () => clearTimeout(t);
  }, []);
  const colors = { success: "#16a34a", error: "#dc2626", info: "#2563eb", warning: "#d97706" };
  return (
    <div className="notification" style={{ borderLeft: `4px solid ${colors[type] || colors.info}` }}>
      {message}
    </div>
  );
}

// ─── HERO / LANDING ─────────────────────────────────────────────────────────

function LandingPage({ onStart, onAdmin }) {
  const [step, setStep] = useState("hero");
  const [form, setForm] = useState({ nom: "", prenom: "", apogee: "" });
  const [errors, setErrors] = useState({});
  const [checking, setChecking] = useState(false);
  const [blocked, setBlocked] = useState(false);
  const [config] = useState(() => JSON.parse(localStorage.getItem("examConfig") || JSON.stringify(DEFAULT_CONFIG)));

  const validate = () => {
    const e = {};
    if (!form.nom.trim()) e.nom = "Requis";
    if (!form.prenom.trim()) e.prenom = "Requis";
    if (!form.apogee.trim()) e.apogee = "Requis";
    else if (!/^\d{5,10}$/.test(form.apogee.trim())) e.apogee = "Code Apogée invalide (5-10 chiffres)";
    setErrors(e);
    return Object.keys(e).length === 0;
  };

  const handleSubmit = () => {
    if (!validate()) return;
    setChecking(true);
    // Check if apogee already attempted (frontend + "backend" via DB)
    const alreadyAttempted = DB.hasAttempted(form.apogee.trim());
    setChecking(false);
    if (alreadyAttempted) {
      setBlocked(true);
      return;
    }
    onStart(form);
  };

  const heroSection = (
    <div className="hero-bg" style={{ minHeight: "100vh", display: "flex", flexDirection: "column" }}>
      <div className="hero-grid" />
      <div style={{ position: "relative", zIndex: 1 }}>
        <nav style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "20px 40px" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
            <div style={{ width: 36, height: 36, borderRadius: 10, background: "rgba(255,255,255,0.15)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 18 }}>🎓</div>
            <span style={{ color: "#fff", fontWeight: 700, fontSize: 16 }}>EduExam Pro</span>
          </div>
          <button onClick={onAdmin} className="btn btn-secondary btn-sm" style={{ background: "rgba(255,255,255,0.1)", border: "1px solid rgba(255,255,255,0.2)", color: "#fff" }}>
            ⚙️ Administration
          </button>
        </nav>

        <div style={{ maxWidth: 800, margin: "60px auto", padding: "0 24px", textAlign: "center" }} className="fade-in">
          <div style={{ display: "inline-flex", alignItems: "center", gap: 8, background: "rgba(37,99,235,0.3)", border: "1px solid rgba(59,130,246,0.4)", borderRadius: 99, padding: "6px 16px", marginBottom: 28 }}>
            <span style={{ width: 8, height: 8, background: "#4ade80", borderRadius: "50%", display: "inline-block" }} className="pulse" />
            <span style={{ color: "#93c5fd", fontSize: 13, fontWeight: 600 }}>Session d'évaluation active</span>
          </div>
          <h1 style={{ fontSize: "clamp(28px, 5vw, 52px)", fontWeight: 800, color: "#fff", lineHeight: 1.15, marginBottom: 20 }}>{config.title}</h1>
          <p style={{ fontSize: 18, color: "#94a3b8", lineHeight: 1.7, marginBottom: 40, maxWidth: 560, margin: "0 auto 40px" }}>
            {config.subtitle || "Plateforme d'évaluation professionnelle pour le Master Big Data & Intelligence Artificielle"}
          </p>
          <div style={{ display: "flex", gap: 12, justifyContent: "center", flexWrap: "wrap", marginBottom: 48 }}>
            {[
              { icon: "⏱", val: `${config.duration} min`, label: "Durée" },
              { icon: "📝", val: config.questionCount, label: "Questions" },
              { icon: "🎯", val: `${config.passingScore}%`, label: "Seuil réussite" },
              { icon: "📊", val: config.questionCount * config.scorePerQuestion, label: "Points total" },
            ].map(({ icon, val, label }) => (
              <div key={label} style={{ background: "rgba(255,255,255,0.07)", border: "1px solid rgba(255,255,255,0.12)", borderRadius: 14, padding: "16px 24px", minWidth: 120 }}>
                <div style={{ fontSize: 22, marginBottom: 4 }}>{icon}</div>
                <div style={{ color: "#fff", fontWeight: 700, fontSize: 22 }}>{val}</div>
                <div style={{ color: "#64748b", fontSize: 12, marginTop: 2 }}>{label}</div>
              </div>
            ))}
          </div>
          <button onClick={() => setStep("form")} className="btn btn-primary btn-lg" style={{ fontSize: 17, padding: "16px 40px", borderRadius: 14, boxShadow: "0 8px 32px rgba(37,99,235,0.5)" }}>
            Commencer l'évaluation →
          </button>
          <p style={{ color: "#475569", fontSize: 13, marginTop: 20 }}>Assurez-vous d'être dans un environnement calme avant de commencer</p>
        </div>

        <div style={{ maxWidth: 900, margin: "60px auto", padding: "0 24px", display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(240px, 1fr))", gap: 20 }}>
          {[
            { icon: "🔒", title: "Session sécurisée", desc: "Détection changement d'onglet, sauvegarde automatique" },
            { icon: "⚡", title: "Rendu instantané", desc: "Formules LaTeX, code colorisé, navigation fluide" },
            { icon: "📄", title: "Certificat PDF", desc: "Généré automatiquement à la fin du test" },
            { icon: "🔁", title: "Tentative unique", desc: "Chaque étudiant ne peut passer le test qu'une seule fois" },
          ].map(({ icon, title, desc }) => (
            <div key={title} style={{ background: "rgba(255,255,255,0.04)", border: "1px solid rgba(255,255,255,0.08)", borderRadius: 14, padding: "20px" }}>
              <div style={{ fontSize: 28, marginBottom: 10 }}>{icon}</div>
              <h3 style={{ color: "#e2e8f0", fontWeight: 700, fontSize: 15, marginBottom: 6 }}>{title}</h3>
              <p style={{ color: "#64748b", fontSize: 13, lineHeight: 1.6 }}>{desc}</p>
            </div>
          ))}
        </div>
      </div>
    </div>
  );

  if (step === "hero") return heroSection;

  // Blocked screen
  if (blocked) return (
    <div style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", background: "#f8faff", padding: 24 }}>
      <style>{CSS}</style>
      <div className="card slide-in" style={{ maxWidth: 480, width: "100%", padding: 40, textAlign: "center" }}>
        <div style={{ width: 72, height: 72, borderRadius: "50%", background: "#fee2e2", display: "flex", alignItems: "center", justifyContent: "center", margin: "0 auto 20px", fontSize: 32 }}>🚫</div>
        <h2 style={{ fontSize: 22, fontWeight: 800, color: "#b91c1c", marginBottom: 10 }}>Accès refusé</h2>
        <p style={{ fontSize: 15, color: "#64748b", lineHeight: 1.7, marginBottom: 8 }}>
          Le Code Apogée <strong style={{ color: "#1e293b", fontFamily: "monospace" }}>{form.apogee}</strong> a déjà été utilisé pour passer cette évaluation.
        </p>
        <p style={{ fontSize: 14, color: "#94a3b8", marginBottom: 28 }}>
          Chaque étudiant ne dispose que d'une seule tentative. Si vous pensez qu'il s'agit d'une erreur, veuillez contacter l'administration.
        </p>
        <div style={{ background: "#fef2f2", border: "1px solid #fecaca", borderRadius: 10, padding: 14, marginBottom: 24, textAlign: "left" }}>
          <p style={{ fontSize: 13, color: "#b91c1c", fontWeight: 600 }}>⚠️ Tentative bloquée</p>
          <p style={{ fontSize: 12, color: "#9f1239", marginTop: 4 }}>Ce code Apogée est enregistré dans le système comme ayant déjà soumis le test. Aucune nouvelle soumission ne sera acceptée.</p>
        </div>
        <button onClick={() => { setBlocked(false); setForm({ nom: "", prenom: "", apogee: "" }); setStep("hero"); }} className="btn btn-secondary" style={{ width: "100%", justifyContent: "center" }}>
          ← Retour à l'accueil
        </button>
      </div>
    </div>
  );

  // Form
  return (
    <div style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", background: "#f8faff", padding: 24 }}>
      <style>{CSS}</style>
      <div className="card slide-in" style={{ maxWidth: 440, width: "100%", padding: 40 }}>
        <button onClick={() => setStep("hero")} className="btn btn-secondary btn-sm" style={{ marginBottom: 24 }}>← Retour</button>
        <div style={{ textAlign: "centre", marginBottom: 32 }}>
          <div style={{ width: 56, height: 56, borderRadius: 14, background: "linear-gradient(135deg, #2563eb, #7c3aed)", display: "flex", alignItems: "center", justifyContent: "center", margin: "0 auto 16px", fontSize: 24 }}>🎓</div>
          <h2 style={{ fontSize: 22, fontWeight: 700, color: "#1e293b", marginBottom: 6 }}>Identification</h2>
          <p style={{ fontSize: 14, color: "#64748b" }}>Veuillez compléter vos informations pour commencer</p>
        </div>

        {[
          { key: "nom", label: "Nom", placeholder: "Votre nom de famille", icon: "👤" },
          { key: "prenom", label: "Prénom", placeholder: "Votre prénom", icon: "✏️" },
          { key: "apogee", label: "Code Apogée", placeholder: "Ex: 12345678", icon: "🔑" },
        ].map(({ key, label, placeholder, icon }) => (
          <div key={key} style={{ textAlign: "left",marginBottom: 20 }}>
            <label>{icon} {label}</label>
            <input
              value={form[key]} placeholder={placeholder}
              onChange={e => setForm(f => ({ ...f, [key]: e.target.value }))}
              onKeyDown={e => e.key === "Enter" && handleSubmit()}
              style={errors[key] ? { borderColor: "#ef4444" } : {}}
            />
            {errors[key] && <p style={{ color: "#ef4444", fontSize: 12, marginTop: 4 }}>⚠ {errors[key]}</p>}
          </div>
        ))}

        <button onClick={handleSubmit} disabled={checking} className="btn btn-primary btn-lg" style={{ width: "100%", justifyContent: "center", marginTop: 8 }}>
          {checking ? "⏳ Vérification..." : "🚀 Commencer le test"}
        </button>

        <div style={{ marginTop: 16, padding: 12, background: "#fffbeb", borderRadius: 10, border: "1px solid #fde68a" }}>
          <p style={{ fontSize: 12, color: "#92400e", lineHeight: 1.6 }}>
            ⚠️ <strong>Tentative unique :</strong> Chaque Code Apogée ne peut être utilisé qu'une seule fois. Assurez-vous d'être prêt avant de commencer.
          </p>
        </div>

        <div style={{ marginTop: 12, padding: 14, background: "#f8faff", borderRadius: 10, border: "1px solid #e2e8f0" }}>
          <p style={{ fontSize: 12, color: "#64748b", lineHeight: 1.6 }}>
            📋 <strong>Test :</strong> {config.title}<br />
            ⏱ <strong>Durée :</strong> {config.duration} minutes<br />
            📝 <strong>Questions :</strong> {config.questionCount}<br />
            🎯 <strong>Note minimale :</strong> {config.passingScore}%
          </p>
        </div>
      </div>
    </div>
  );
}

// ─── EXAM INTERFACE ──────────────────────────────────────────────────────────

function ExamInterface({ student, onFinish }) {
  const config = JSON.parse(localStorage.getItem("examConfig") || JSON.stringify(DEFAULT_CONFIG));
  const allQ = JSON.parse(localStorage.getItem("examQuestions") || JSON.stringify(QUESTION_BANK));

  const [questions] = useState(() => {
    let pool = [...allQ];
    if (config.shuffleQuestions) pool = shuffle(pool);
    return pool.slice(0, config.questionCount);
  });

  const [current, setCurrent] = useState(0);
  const [answers, setAnswers] = useState({});
  const [visited, setVisited] = useState({ 0: true });
  const [timeLeft, setTimeLeft] = useState(config.duration * 60);
  const [finished, setFinished] = useState(false);
  const [showConfirm, setShowConfirm] = useState(false);
  const [tabWarnings, setTabWarnings] = useState(0);
  const [showTabAlert, setShowTabAlert] = useState(false);

  useEffect(() => {
    if (finished) return;
    const t = setInterval(() => {
      setTimeLeft(prev => {
        if (prev <= 1) { clearInterval(t); handleFinish(true); return 0; }
        return prev - 1;
      });
    }, 1000);
    return () => clearInterval(t);
  }, [finished]);

  useEffect(() => {
    localStorage.setItem("examProgress", JSON.stringify({ answers, current, timeLeft }));
  }, [answers, current]);

  useEffect(() => {
    const handle = () => {
      if (document.hidden && !finished) {
        setTabWarnings(w => {
          const next = w + 1;
          setShowTabAlert(true);
          setTimeout(() => setShowTabAlert(false), 4000);
          if (next >= config.tabSwitchWarnings) handleFinish(true);
          return next;
        });
      }
    };
    document.addEventListener("visibilitychange", handle);
    return () => document.removeEventListener("visibilitychange", handle);
  }, [finished]);

  const handleFinish = useCallback((auto = false) => {
    if (finished) return;
    setFinished(true);
    setShowConfirm(false);
    const score = calculateScore();
    const result = {
      student, answers, questions,
      score: score.obtained, total: score.total, percent: score.percent,
      timeUsed: config.duration * 60 - timeLeft,
      date: new Date().toISOString(),
      config,
    };
    DB.addResult(result);
    localStorage.removeItem("examProgress");
    onFinish(result);
  }, [finished, answers, questions, timeLeft]);

  const calculateScore = () => {
    let obtained = 0;
    questions.forEach((q, i) => { if (answers[i] === q.answer) obtained += q.score; });
    const total = questions.reduce((s, q) => s + q.score, 0);
    return { obtained, total, percent: Math.round((obtained / total) * 100) };
  };

  const goTo = (idx) => {
    setCurrent(idx);
    setVisited(v => ({ ...v, [idx]: true }));
  };

  const q = questions[current];
  const isAnswered = (i) => answers[i] !== undefined;
  const getState = (i) => {
    if (i === current) return "current";
    if (isAnswered(i)) return "answered";
    if (visited[i]) return "visited";
    return "not-visited";
  };

  const timerDanger = timeLeft < 300;
  const answeredCount = Object.keys(answers).length;

  return (
    <div style={{ display: "flex", height: "100vh", overflow: "hidden", background: "#f8faff" }}>
      <style>{CSS}</style>

      {showTabAlert && (
        <div style={{ position: "fixed", top: 0, left: 0, right: 0, zIndex: 9999, background: "#dc2626", color: "#fff", padding: "12px 24px", textAlign: "center", fontWeight: 600, fontSize: 14 }}>
          ⚠️ Changement d'onglet détecté ! Avertissement {tabWarnings}/{config.tabSwitchWarnings}. Le test sera soumis automatiquement.
        </div>
      )}

      {/* Sidebar */}
      <div className="sidebar" style={{ padding: 20 }}>
        <div style={{ marginBottom: 20 }}>
          <div style={{ fontWeight: 800, fontSize: 15, color: "#1e293b", marginBottom: 4 }}>🎓 EduExam Pro</div>
          <div style={{ fontSize: 12, color: "#94a3b8" }}>{student.prenom} {student.nom}</div>
          <div style={{ fontSize: 11, color: "#cbd5e1", fontFamily: "monospace" }}>#{student.apogee}</div>
        </div>

        <div style={{ background: timerDanger ? "#fef2f2" : "#f0f9ff", borderRadius: 12, padding: "14px 16px", marginBottom: 20, border: `1px solid ${timerDanger ? "#fecaca" : "#bae6fd"}` }}>
          <div style={{ fontSize: 11, fontWeight: 700, color: timerDanger ? "#dc2626" : "#0369a1", textTransform: "uppercase", letterSpacing: 1, marginBottom: 4 }}>⏱ Temps restant</div>
          <div style={{ fontSize: 28, fontWeight: 800, fontFamily: "monospace", color: timerDanger ? "#dc2626" : "#0f172a" }} className={timerDanger ? "timer-danger" : ""}>
            {formatTime(timeLeft)}
          </div>
        </div>

        <div style={{ marginBottom: 20 }}>
          <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12, color: "#64748b", marginBottom: 8 }}>
            <span>Progression</span>
            <span>{answeredCount}/{questions.length}</span>
          </div>
          <div className="progress-bar">
            <div className="progress-fill" style={{ width: `${(answeredCount / questions.length) * 100}%` }} />
          </div>
        </div>

        <div style={{ fontSize: 11, fontWeight: 700, color: "#94a3b8", textTransform: "uppercase", letterSpacing: 1, marginBottom: 12 }}>Navigation</div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(5, 1fr)", gap: 6, marginBottom: 24 }}>
          {questions.map((_, i) => (
            <div key={i} className={`q-bubble ${getState(i)}`} onClick={() => goTo(i)}>{i + 1}</div>
          ))}
        </div>

        <div style={{ display: "flex", flexDirection: "column", gap: 6, fontSize: 11 }}>
          {[
            { color: "#dcfce7", border: "#4ade80", label: "Répondu" },
            { color: "#dbeafe", border: "#3b82f6", label: "En cours" },
            { color: "#fef3c7", border: "#fbbf24", label: "Visité" },
            { color: "#f1f5f9", border: "transparent", label: "Non visité" },
          ].map(({ color, border, label }) => (
            <div key={label} style={{ display: "flex", alignItems: "center", gap: 8 }}>
              <div style={{ width: 14, height: 14, borderRadius: 4, background: color, border: `2px solid ${border}`, flexShrink: 0 }} />
              <span style={{ color: "#64748b" }}>{label}</span>
            </div>
          ))}
        </div>
      </div>

      {/* Main */}
      <div style={{ flex: 1, overflowY: "auto", padding: 32, paddingBottom: 88 }}>
        <div style={{ maxWidth: 720, margin: "0 auto" }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 28 }}>
            <div>
              <span className="badge badge-purple" style={{ marginRight: 8 }}>{q.category}</span>
              <span className={`badge ${q.difficulty === "Facile" ? "badge-green" : q.difficulty === "Moyen" ? "badge-amber" : "badge-red"}`}>{q.difficulty}</span>
            </div>
            <div style={{ display: "flex", alignItems: "center", gap: 16 }}>
              <span style={{ fontSize: 13, color: "#94a3b8" }}>Q{current + 1}/{questions.length}</span>
              <button onClick={() => setShowConfirm(true)} className="btn btn-primary btn-sm">Terminer ✓</button>
            </div>
          </div>

          <div className="card" style={{ marginBottom: 24 }}>
            <div style={{ fontSize: 13, color: "#94a3b8", marginBottom: 14, fontWeight: 600 }}>Question {current + 1} — {q.score} point{q.score > 1 ? "s" : ""}</div>
            <div style={{ fontSize: 16, lineHeight: 1.7, color: "#1e293b" }}>{renderText(q.text)}</div>
          </div>

          <div style={{ display: "flex", flexDirection: "column", gap: 12, marginBottom: 32 }}>
            {q.options.map((opt, oi) => (
              <div
                key={oi}
                className={`option-card ${answers[current] === oi ? "selected" : ""}`}
                onClick={() => setAnswers(a => ({ ...a, [current]: oi }))}
              >
                <div style={{ width: 28, height: 28, borderRadius: 8, background: answers[current] === oi ? "#2563eb" : "#f1f5f9", color: answers[current] === oi ? "#fff" : "#64748b", display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 700, fontSize: 12, flexShrink: 0 }}>
                  {String.fromCharCode(65 + oi)}
                </div>
                <span>{renderText(opt)}</span>
              </div>
            ))}
          </div>

        </div>
      </div>

      {/* Sticky nav buttons fixed at bottom */}
      <div style={{ position: "fixed", bottom: 0, left: 280, right: 0, background: "#fff", borderTop: "1px solid #e8edf5", padding: "14px 32px", display: "flex", justifyContent: "space-between", zIndex: 200, boxShadow: "0 -2px 12px rgba(0,0,0,0.06)" }}>
        <button onClick={() => current > 0 && goTo(current - 1)} className="btn btn-secondary" disabled={current === 0}>← Précédent</button>
        <button onClick={() => current < questions.length - 1 ? goTo(current + 1) : setShowConfirm(true)} className="btn btn-primary">
          {current < questions.length - 1 ? "Suivant →" : "Terminer ✓"}
        </button>
      </div>

      {showConfirm && (
        <div className="modal-backdrop">
          <div className="modal slide-in">
            <h2 style={{ fontWeight: 800, fontSize: 20, marginBottom: 16 }}>Soumettre le test ?</h2>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginBottom: 20 }}>
              <div style={{ background: "#f0fdf4", borderRadius: 10, padding: 14, textAlign: "center" }}>
                <div style={{ fontSize: 24, fontWeight: 800, color: "#15803d" }}>{answeredCount}</div>
                <div style={{ fontSize: 12, color: "#64748b" }}>Répondues</div>
              </div>
              <div style={{ background: "#fff5f5", borderRadius: 10, padding: 14, textAlign: "center" }}>
                <div style={{ fontSize: 24, fontWeight: 800, color: "#dc2626" }}>{questions.length - answeredCount}</div>
                <div style={{ fontSize: 12, color: "#64748b" }}>Non répondues</div>
              </div>
            </div>
            <p style={{ color: "#64748b", fontSize: 14, marginBottom: 24 }}>
              Cette action est définitive. Votre résultat sera enregistré et un PDF sera généré automatiquement.
            </p>
            <div style={{ display: "flex", gap: 12, justifyContent: "flex-end" }}>
              <button onClick={() => setShowConfirm(false)} className="btn btn-secondary">Annuler</button>
              <button onClick={() => handleFinish(false)} className="btn btn-primary">Confirmer et terminer</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ─── RESULTS PAGE ────────────────────────────────────────────────────────────

function ResultsPage({ result, onHome }) {
  // Generate and store PDF for admin, without downloading
  useEffect(() => {
    generateAndStorePDF(result).catch(() => {});
  }, []);

  return (
    <div style={{ minHeight: "100vh", background: "#f8faff", display: "flex", alignItems: "center", justifyContent: "center", padding: 24 }}>
      <style>{CSS}</style>
      <div className="card slide-in" style={{ maxWidth: 480, width: "100%", padding: 48, textAlign: "center" }}>
        <div style={{ fontSize: 64, marginBottom: 20 }}>✅</div>
        <h1 style={{ fontSize: 24, fontWeight: 800, color: "#1e293b", marginBottom: 16 }}>
          Vous avez complété le test. Merci.
        </h1>
        <button onClick={onHome} className="btn btn-secondary" style={{ marginTop: 8 }}>🏠 Accueil</button>
      </div>
    </div>
  );
}

// ─── ADMIN LOGIN ─────────────────────────────────────────────────────────────

function AdminLogin({ onLogin, onBack }) {
  const [form, setForm] = useState({ username: "", password: "" });
  const [error, setError] = useState("");

  const handle = () => {
    if (form.username === ADMIN_CREDENTIALS.username && form.password === ADMIN_CREDENTIALS.password) {
      onLogin();
    } else {
      setError("Identifiants incorrects");
    }
  };

  return (
    <div style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", background: "linear-gradient(135deg, #0f172a, #1e3a6e)" }}>
      <style>{CSS}</style>
      <div className="card slide-in" style={{ maxWidth: 380, width: "90%", padding: 40 }}>
        <div style={{ textAlign: "center", marginBottom: 32 }}>
          <div style={{ fontSize: 40, marginBottom: 12 }}>⚙️</div>
          <h2 style={{ fontWeight: 800, fontSize: 22 }}>Administration</h2>
          <p style={{ color: "#64748b", fontSize: 14, marginTop: 4 }}>Connectez-vous pour accéder au dashboard</p>
        </div>
        <label style={{textAlign: "left"}}>Identifiant</label>
        <input value={form.username} onChange={e => setForm(f => ({ ...f, username: e.target.value }))} placeholder="admin" style={{ marginBottom: 16 }} />
        <label style={{textAlign: "left"}}>Mot de passe</label>
        <input type="password" value={form.password} onChange={e => setForm(f => ({ ...f, password: e.target.value }))} placeholder="••••••••" onKeyDown={e => e.key === "Enter" && handle()} style={{ marginBottom: error ? 8 : 20 }} />
        {error && <p style={{color: "#ef4444", fontSize: 13, marginBottom: 16 }}>⚠ {error}</p>}
        <button onClick={handle} className="btn btn-primary" style={{ width: "100%", justifyContent: "center" }}>Se connecter</button>
        <button onClick={onBack} className="btn btn-secondary" style={{ width: "100%", justifyContent: "center", marginTop: 10 }}>← Retour</button>
        <p style={{ fontSize: 12, color: "#94a3b8", textAlign: "center", marginTop: 16 }}></p>
      </div>
    </div>
  );
}

// ─── ADMIN DASHBOARD ─────────────────────────────────────────────────────────

function AdminDashboard({ onLogout }) {
  const [tab, setTab] = useState("dashboard");
  const [config, setConfig] = useState(() => JSON.parse(localStorage.getItem("examConfig") || JSON.stringify(DEFAULT_CONFIG)));
  const [questions, setQuestions] = useState(() => JSON.parse(localStorage.getItem("examQuestions") || JSON.stringify(QUESTION_BANK)));
  const [results, setResults] = useState(() => DB.getResults());
  const [notification, setNotification] = useState(null);
  const [editingQ, setEditingQ] = useState(null);
  const [showQModal, setShowQModal] = useState(false);
  const [searchTerm, setSearchTerm] = useState("");
  const [selectedResult, setSelectedResult] = useState(null);

  const notify = (message, type = "success") => {
    setNotification({ message, type });
    setTimeout(() => setNotification(null), 4000);
  };

  const saveConfig = () => {
    localStorage.setItem("examConfig", JSON.stringify(config));
    notify("Configuration sauvegardée ✅");
  };

  const saveQuestions = (qs) => {
    setQuestions(qs);
    localStorage.setItem("examQuestions", JSON.stringify(qs));
  };

  const deleteQuestion = (id) => {
    if (confirm("Supprimer cette question ?")) {
      saveQuestions(questions.filter(q => q.id !== id));
      notify("Question supprimée", "info");
    }
  };

  const saveQuestion = (q) => {
    if (q.id && questions.find(x => x.id === q.id)) {
      saveQuestions(questions.map(x => x.id === q.id ? q : x));
      notify("Question modifiée ✅");
    } else {
      saveQuestions([...questions, { ...q, id: Date.now() }]);
      notify("Question ajoutée ✅");
    }
    setShowQModal(false);
    setEditingQ(null);
  };

  const deleteStudent = (apogee) => {
    if (!confirm(`Supprimer l'étudiant avec le code Apogée ${apogee} ? Cette action est irréversible.`)) return;
    DB.deleteResult(apogee);
    DB.deletePDF(apogee);
    setResults(DB.getResults());
    if (selectedResult?.student?.apogee === apogee) setSelectedResult(null);
    notify(`Étudiant ${apogee} supprimé`, "info");
  };

  const downloadStudentPDF = async (r) => {
    notify("Génération du PDF...", "info");
    try {
      const pdfDataUrl = await generateAndStorePDF(r);
      const link = document.createElement("a");
      link.href = pdfDataUrl;
      link.download = `Evaluation_${r.student.nom}_${r.student.prenom}_${r.student.apogee}.pdf`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      notify("PDF téléchargé ✅");
    } catch (e) {
      notify("Erreur lors de la génération du PDF", "error");
    }
  };

  const filteredResults = results.filter(r => {
    if (!searchTerm.trim()) return true;
    const s = searchTerm.toLowerCase();
    return (
      r.student.nom.toLowerCase().includes(s) ||
      r.student.prenom.toLowerCase().includes(s) ||
      r.student.apogee.includes(s)
    );
  });

  const avgScore = results.length ? Math.round(results.reduce((s, r) => s + r.percent, 0) / results.length) : 0;
  const passRate = results.length ? Math.round((results.filter(r => r.percent >= config.passingScore).length / results.length) * 100) : 0;
  const avgTime = results.length ? Math.round(results.reduce((s, r) => s + r.timeUsed, 0) / results.length / 60) : 0;

  const TABS = [
    { id: "dashboard", label: "📊 Dashboard" },
    { id: "students", label: "👥 Étudiants" },
    { id: "config", label: "⚙️ Configuration" },
    { id: "questions", label: "📝 Questions" },
    { id: "results", label: "📋 Résultats" },
  ];

  return (
    <div style={{ minHeight: "100vh", background: "#f8faff" }}>
      <style>{CSS}</style>
      {notification && <Notification {...notification} onClose={() => setNotification(null)} />}

      <div className="nav-admin">
        <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
          <div style={{ fontWeight: 800, fontSize: 16, color: "#1e293b", marginRight: 8 }}>🎓 EduExam <span style={{ color: "#2563eb" }}>Admin</span></div>
          <div style={{ display: "flex", gap: 4, flexWrap: "wrap" }}>
            {TABS.map(t => (
              <button key={t.id} className={`tab ${tab === t.id ? "active" : ""}`} onClick={() => setTab(t.id)}>{t.label}</button>
            ))}
          </div>
        </div>
        <button onClick={onLogout} className="btn btn-secondary btn-sm">🚪 Déconnexion</button>
      </div>

      <div style={{ padding: 28, maxWidth: 1200, margin: "0 auto" }}>

        {/* DASHBOARD */}
        {tab === "dashboard" && (
          <div className="fade-in">
            <h2 style={{ fontSize: 22, fontWeight: 800, marginBottom: 24 }}>Tableau de bord</h2>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))", gap: 16, marginBottom: 28 }}>
              {[
                { icon: "👥", label: "Étudiants", val: results.length, color: "#2563eb" },
                { icon: "📋", label: "Évaluations", val: results.length, color: "#7c3aed" },
                { icon: "📊", label: "Moyenne", val: `${avgScore}%`, color: "#059669" },
                { icon: "⏱", label: "Temps moyen", val: `${avgTime} min`, color: "#d97706" },
                { icon: "✅", label: "Taux réussite", val: `${passRate}%`, color: results.length && passRate >= 50 ? "#059669" : "#dc2626" },
                { icon: "📝", label: "Questions", val: questions.length, color: "#0891b2" },
              ].map(({ icon, label, val, color }) => (
                <div className="stat-card" key={label}>
                  <div style={{ fontSize: 24 }}>{icon}</div>
                  <div className="stat-number" style={{ color }}>{val}</div>
                  <div className="stat-label">{label}</div>
                </div>
              ))}
            </div>

            {results.length === 0 ? (
              <div className="card" style={{ textAlign: "center", padding: 48 }}>
                <div style={{ fontSize: 48, marginBottom: 12 }}>📭</div>
                <p style={{ color: "#94a3b8", fontSize: 16 }}>Aucune évaluation effectuée pour l'instant</p>
              </div>
            ) : (
              <div className="card">
                <h3 style={{ fontWeight: 700, marginBottom: 16 }}>📈 Répartition des scores</h3>
                <div style={{ display: "flex", gap: 4, alignItems: "flex-end", height: 120 }}>
                  {[0, 10, 20, 30, 40, 50, 60, 70, 80, 90].map(range => {
                    const count = results.filter(r => r.percent >= range && r.percent < range + 10).length;
                    const maxCount = Math.max(...[0, 10, 20, 30, 40, 50, 60, 70, 80, 90].map(r => results.filter(x => x.percent >= r && x.percent < r + 10).length), 1);
                    return (
                      <div key={range} style={{ flex: 1, display: "flex", flexDirection: "column", alignItems: "center", gap: 4 }}>
                        <span style={{ fontSize: 10, color: "#94a3b8" }}>{count}</span>
                        <div style={{ width: "100%", height: `${(count / maxCount) * 90}%`, background: range >= config.passingScore ? "#4ade80" : "#f87171", borderRadius: "4px 4px 0 0", minHeight: count > 0 ? 4 : 0 }} />
                        <span style={{ fontSize: 9, color: "#94a3b8" }}>{range}</span>
                      </div>
                    );
                  })}
                </div>
              </div>
            )}
          </div>
        )}

        {/* STUDENTS */}
        {tab === "students" && (
          <div className="fade-in">
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 24, flexWrap: "wrap", gap: 12 }}>
              <div>
                <h2 style={{ fontSize: 22, fontWeight: 800 }}>Gestion des étudiants</h2>
                <p style={{ color: "#64748b", fontSize: 14, marginTop: 2 }}>{results.length} étudiant{results.length !== 1 ? "s" : ""} enregistré{results.length !== 1 ? "s" : ""}</p>
              </div>
              <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
                <div style={{ position: "relative" }}>
                  <span style={{ position: "absolute", left: 12, top: "50%", transform: "translateY(-50%)", color: "#94a3b8", fontSize: 14 }}>🔍</span>
                  <input
                    className="search-input"
                    placeholder="Rechercher par nom ou Apogée..."
                    value={searchTerm}
                    onChange={e => setSearchTerm(e.target.value)}
                  />
                </div>
                <button onClick={() => exportResultsToExcel(results, config).catch(() => notify("Erreur export Excel", "error"))} className="btn btn-success btn-sm" title="Exporter en Excel">
                  📊 Exporter Excel
                </button>
              </div>
            </div>

            {filteredResults.length === 0 ? (
              <div className="card" style={{ textAlign: "center", padding: 48 }}>
                <div style={{ fontSize: 48, marginBottom: 12 }}>🔍</div>
                <p style={{ color: "#94a3b8" }}>{searchTerm ? "Aucun étudiant trouvé pour cette recherche" : "Aucun étudiant enregistré"}</p>
              </div>
            ) : (
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(340px, 1fr))", gap: 16 }}>
                {filteredResults.slice().reverse().map((r, i) => {
                  const passed = r.percent >= config.passingScore;
                  const pdfs = DB.getPDFs();
                  const hasPDF = !!pdfs[r.student.apogee];
                  return (
                    <div key={i} className="card" style={{ padding: 20 }}>
                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 14 }}>
                        <div>
                          <p style={{ fontWeight: 700, fontSize: 16, color: "#1e293b" }}>{r.student.prenom} {r.student.nom}</p>
                          <p style={{ fontSize: 12, color: "#94a3b8", fontFamily: "monospace", marginTop: 2 }}>#{r.student.apogee}</p>
                        </div>
                        <span className={`badge ${passed ? "badge-green" : "badge-red"}`}>
                          {passed ? "✅ Admis" : "❌ Ajourné"}
                        </span>
                      </div>

                      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 8, marginBottom: 14 }}>
                        {[
                          { label: "Score", val: `${r.score}/${r.total}` },
                          { label: "Pourcentage", val: `${r.percent}%`, color: passed ? "#15803d" : "#dc2626" },
                          { label: "Temps", val: formatTime(r.timeUsed) },
                        ].map(({ label, val, color }) => (
                          <div key={label} style={{ background: "#f8faff", borderRadius: 8, padding: "8px 10px" }}>
                            <p style={{ fontSize: 10, color: "#94a3b8", fontWeight: 600, textTransform: "uppercase" }}>{label}</p>
                            <p style={{ fontSize: 13, fontWeight: 700, color: color || "#1e293b" }}>{val}</p>
                          </div>
                        ))}
                      </div>

                      <p style={{ fontSize: 11, color: "#94a3b8", marginBottom: 14 }}>
                        📅 {new Date(r.date).toLocaleString("fr-FR")}
                      </p>

                      <div style={{ display: "flex", gap: 8 }}>
                        <button onClick={() => setSelectedResult(r)} className="btn btn-secondary btn-sm" style={{ flex: 1, justifyContent: "center" }}>
                          👁 Détails
                        </button>
                        <button onClick={() => downloadStudentPDF(r)} className="btn btn-primary btn-sm" style={{ flex: 1, justifyContent: "center" }}>
                          📄 {hasPDF ? "PDF" : "Générer PDF"}
                        </button>
                        <button onClick={() => deleteStudent(r.student.apogee)} className="btn btn-danger btn-sm" style={{ justifyContent: "center" }}>
                          🗑
                        </button>
                      </div>
                    </div>
                  );
                })}
              </div>
            )}
          </div>
        )}

        {/* CONFIG */}
        {tab === "config" && (
          <div className="fade-in">
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 24 }}>
              <h2 style={{ fontSize: 22, fontWeight: 800 }}>Configuration du test</h2>
              <button onClick={saveConfig} className="btn btn-primary">💾 Sauvegarder</button>
            </div>

            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 20 }}>
              <div className="card">
                <h3 style={{ fontWeight: 700, marginBottom: 20, color: "#1e293b" }}>📋 Informations générales</h3>
                {[
                  { key: "title", label: "Titre du test", type: "text" },
                  { key: "subtitle", label: "Sous-titre", type: "text" },
                ].map(({ key, label, type }) => (
                  <div key={key} style={{ marginBottom: 16 }}>
                    <label>{label}</label>
                    <input type={type} value={config[key] || ""} onChange={e => setConfig(c => ({ ...c, [key]: e.target.value }))} />
                  </div>
                ))}
              </div>

              <div className="card">
                <h3 style={{ fontWeight: 700, marginBottom: 20 }}>⚙️ Paramètres</h3>
                {[
                  { key: "duration", label: "Durée (minutes)", type: "number" },
                  { key: "questionCount", label: "Nombre de questions", type: "number" },
                  { key: "scorePerQuestion", label: "Points par question", type: "number" },
                  { key: "passingScore", label: "Seuil de réussite (%)", type: "number" },
                  { key: "tabSwitchWarnings", label: "Avertissements onglet", type: "number" },
                ].map(({ key, label, type }) => (
                  <div key={key} style={{ marginBottom: 16 }}>
                    <label>{label}</label>
                    <input type={type} value={config[key]} onChange={e => setConfig(c => ({ ...c, [key]: Number(e.target.value) }))} />
                  </div>
                ))}
              </div>

              <div className="card">
                <h3 style={{ fontWeight: 700, marginBottom: 20 }}>🎲 Options avancées</h3>
                {[
                  { key: "shuffleQuestions", label: "Mélanger les questions" },
                  { key: "shuffleAnswers", label: "Mélanger les réponses" },
                  { key: "showFeedback", label: "Afficher les explications" },
                ].map(({ key, label }) => (
                  <div key={key} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
                    <label style={{ marginBottom: 0 }}>{label}</label>
                    <div onClick={() => setConfig(c => ({ ...c, [key]: !c[key] }))}
                      style={{ width: 44, height: 24, borderRadius: 99, cursor: "pointer", background: config[key] ? "#2563eb" : "#e2e8f0", position: "relative", transition: "background 0.2s" }}>
                      <div style={{ width: 18, height: 18, borderRadius: "50%", background: "#fff", position: "absolute", top: 3, left: config[key] ? 23 : 3, transition: "left 0.2s", boxShadow: "0 1px 3px rgba(0,0,0,0.2)" }} />
                    </div>
                  </div>
                ))}
              </div>

              <div className="card">
                <h3 style={{ fontWeight: 700, marginBottom: 20 }}>📂 Catégories actives</h3>
                {["Python", "Machine Learning", "Big Data", "Statistiques", "Deep Learning"].map(cat => (
                  <div key={cat} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
                    <span style={{ fontSize: 14, color: "#334155" }}>{cat}</span>
                    <div onClick={() => setConfig(c => ({ ...c, categories: c.categories.includes(cat) ? c.categories.filter(x => x !== cat) : [...c.categories, cat] }))}
                      style={{ width: 44, height: 24, borderRadius: 99, cursor: "pointer", background: config.categories.includes(cat) ? "#2563eb" : "#e2e8f0", position: "relative", transition: "background 0.2s" }}>
                      <div style={{ width: 18, height: 18, borderRadius: "50%", background: "#fff", position: "absolute", top: 3, left: config.categories.includes(cat) ? 23 : 3, transition: "left 0.2s", boxShadow: "0 1px 3px rgba(0,0,0,0.2)" }} />
                    </div>
                  </div>
                ))}
              </div>
            </div>
          </div>
        )}

        {/* QUESTIONS */}
        {tab === "questions" && (
          <div className="fade-in">
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 24 }}>
              <h2 style={{ fontSize: 22, fontWeight: 800 }}>Banque de questions ({questions.length})</h2>
              <button onClick={() => { setEditingQ(null); setShowQModal(true); }} className="btn btn-primary">+ Ajouter</button>
            </div>

            <div className="card table-container">
              <table>
                <thead>
                  <tr>
                    <th>#</th>
                    <th>Question</th>
                    <th>Catégorie</th>
                    <th>Type</th>
                    <th>Difficulté</th>
                    <th>Points</th>
                    <th>Actions</th>
                  </tr>
                </thead>
                <tbody>
                  {questions.map((q, i) => (
                    <tr key={q.id}>
                      <td style={{ fontWeight: 700, color: "#94a3b8" }}>{i + 1}</td>
                      <td style={{ maxWidth: 300 }}>
                        <p style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", fontSize: 13 }}>
                          {q.text.split("\n")[0].substring(0, 80)}...
                        </p>
                      </td>
                      <td><span className="badge badge-purple">{q.category}</span></td>
                      <td><span className="badge badge-blue">{q.type}</span></td>
                      <td>
                        <span className={`badge ${q.difficulty === "Facile" ? "badge-green" : q.difficulty === "Moyen" ? "badge-amber" : "badge-red"}`}>
                          {q.difficulty}
                        </span>
                      </td>
                      <td style={{ fontWeight: 700 }}>{q.score}</td>
                      <td>
                        <div style={{ display: "flex", gap: 6 }}>
                          <button onClick={() => { setEditingQ(q); setShowQModal(true); }} className="btn btn-secondary btn-sm">✏️</button>
                          <button onClick={() => deleteQuestion(q.id)} className="btn btn-danger btn-sm">🗑</button>
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* RESULTS */}
        {tab === "results" && (
          <div className="fade-in">
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 24, flexWrap: "wrap", gap: 12 }}>
              <h2 style={{ fontSize: 22, fontWeight: 800 }}>Résultats des étudiants</h2>
              <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
                <div style={{ position: "relative" }}>
                  <span style={{ position: "absolute", left: 12, top: "50%", transform: "translateY(-50%)", color: "#94a3b8", fontSize: 14 }}>🔍</span>
                  <input
                    className="search-input"
                    placeholder="Rechercher par nom ou Apogée..."
                    value={searchTerm}
                    onChange={e => setSearchTerm(e.target.value)}
                  />
                </div>
                <button onClick={() => exportResultsToExcel(results, config).catch(() => notify("Erreur export Excel", "error"))} className="btn btn-success btn-sm" title="Exporter en Excel">
                  📊 Exporter Excel
                </button>
              </div>
            </div>
            {filteredResults.length === 0 ? (
              <div className="card" style={{ textAlign: "center", padding: 48 }}>
                <div style={{ fontSize: 48, marginBottom: 12 }}>📭</div>
                <p style={{ color: "#94a3b8" }}>{searchTerm ? "Aucun résultat correspondant" : "Aucun résultat disponible"}</p>
              </div>
            ) : (
              <div className="card table-container">
                <table>
                  <thead>
                    <tr>
                      <th>Étudiant</th>
                      <th>Apogée</th>
                      <th>Score</th>
                      <th>%</th>
                      <th>Temps</th>
                      <th>Date</th>
                      <th>Résultat</th>
                      <th>PDF</th>
                      <th>Action</th>
                    </tr>
                  </thead>
                  <tbody>
                    {filteredResults.slice().reverse().map((r, i) => {
                      const pdfs = DB.getPDFs();
                      const hasPDF = !!pdfs[r.student.apogee];
                      return (
                        <tr key={i}>
                          <td style={{ fontWeight: 600 }}>{r.student.prenom} {r.student.nom}</td>
                          <td style={{ fontFamily: "monospace", color: "#64748b" }}>{r.student.apogee}</td>
                          <td>{r.score}/{r.total}</td>
                          <td style={{ fontWeight: 700, color: r.percent >= config.passingScore ? "#15803d" : "#dc2626" }}>{r.percent}%</td>
                          <td>{formatTime(r.timeUsed)}</td>
                          <td style={{ color: "#94a3b8", fontSize: 13 }}>{new Date(r.date).toLocaleDateString("fr-FR")}</td>
                          <td>
                            <span className={`badge ${r.percent >= config.passingScore ? "badge-green" : "badge-red"}`}>
                              {r.percent >= config.passingScore ? "✅ Admis" : "❌ Ajourné"}
                            </span>
                          </td>
                          <td>
                            <button onClick={() => downloadStudentPDF(r)} className="btn btn-secondary btn-sm" title={hasPDF ? "Télécharger PDF" : "Générer PDF"}>
                              {hasPDF ? "📄" : "⚙️"}
                            </button>
                          </td>
                          <td>
                            <button onClick={() => deleteStudent(r.student.apogee)} className="btn btn-danger btn-sm" title="Supprimer">
                              🗑 Supprimer
                            </button>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        )}
      </div>

      {/* Student Detail Modal */}
      {selectedResult && (
        <div className="modal-backdrop" onClick={e => e.target === e.currentTarget && setSelectedResult(null)}>
          <div className="modal slide-in" style={{ maxWidth: 600, maxHeight: "85vh", overflowY: "auto" }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
              <h2 style={{ fontWeight: 800, fontSize: 20 }}>Détails — {selectedResult.student.prenom} {selectedResult.student.nom}</h2>
              <button onClick={() => setSelectedResult(null)} className="btn btn-secondary btn-sm">✕</button>
            </div>

            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginBottom: 20 }}>
              {[
                { label: "Nom", val: selectedResult.student.nom },
                { label: "Prénom", val: selectedResult.student.prenom },
                { label: "Code Apogée", val: selectedResult.student.apogee },
                { label: "Date", val: new Date(selectedResult.date).toLocaleString("fr-FR") },
                { label: "Score", val: `${selectedResult.score}/${selectedResult.total} pts` },
                { label: "Pourcentage", val: `${selectedResult.percent}%` },
                { label: "Temps utilisé", val: formatTime(selectedResult.timeUsed) },
                { label: "Résultat", val: selectedResult.percent >= config.passingScore ? "✅ Admis(e)" : "❌ Ajourné(e)" },
              ].map(({ label, val }) => (
                <div key={label} style={{ background: "#f8faff", borderRadius: 8, padding: "10px 14px" }}>
                  <p style={{ fontSize: 11, color: "#94a3b8", fontWeight: 600, textTransform: "uppercase" }}>{label}</p>
                  <p style={{ fontSize: 14, fontWeight: 600, color: "#1e293b", marginTop: 2 }}>{val}</p>
                </div>
              ))}
            </div>

            <h3 style={{ fontWeight: 700, fontSize: 14, marginBottom: 12, color: "#475569" }}>DÉTAIL DES RÉPONSES</h3>
            {selectedResult.questions.map((q, i) => {
              const correct = selectedResult.answers[i] === q.answer;
              return (
                <div key={i} style={{ border: `1px solid ${correct ? "#bbf7d0" : "#fecaca"}`, borderRadius: 10, padding: 12, marginBottom: 8, background: correct ? "#f0fdf4" : "#fff5f5" }}>
                  <div style={{ display: "flex", justifyContent: "space-between" }}>
                    <span style={{ fontSize: 13, fontWeight: 600 }}>Q{i + 1} — {q.category}</span>
                    <span style={{ fontSize: 12, fontWeight: 700, color: correct ? "#15803d" : "#dc2626" }}>
                      {correct ? `+${q.score} pts` : "0 pt"} {correct ? "✅" : "❌"}
                    </span>
                  </div>
                  {!correct && (
                    <div style={{ marginTop: 6, fontSize: 12 }}>
                      <p style={{ color: "#dc2626" }}>Réponse donnée : {selectedResult.answers[i] !== undefined ? q.options[selectedResult.answers[i]] : "—"}</p>
                      <p style={{ color: "#15803d" }}>Bonne réponse : {q.options[q.answer]}</p>
                    </div>
                  )}
                </div>
              );
            })}

            <div style={{ display: "flex", gap: 10, marginTop: 20 }}>
              <button onClick={() => downloadStudentPDF(selectedResult)} className="btn btn-primary" style={{ flex: 1, justifyContent: "center" }}>📄 Télécharger PDF</button>
              <button onClick={() => { deleteStudent(selectedResult.student.apogee); setSelectedResult(null); }} className="btn btn-danger" style={{ flex: 1, justifyContent: "center" }}>🗑 Supprimer</button>
            </div>
          </div>
        </div>
      )}

      {showQModal && (
        <QuestionModal
          question={editingQ}
          onSave={saveQuestion}
          onClose={() => { setShowQModal(false); setEditingQ(null); }}
        />
      )}
    </div>
  );
}

// ─── QUESTION MODAL ──────────────────────────────────────────────────────────

function QuestionModal({ question, onSave, onClose }) {
  const [form, setForm] = useState(question || {
    text: "", type: "qcm", category: "Python", difficulty: "Facile", score: 2,
    options: ["", "", "", ""], answer: 0, explanation: "",
  });

  const update = (key, val) => setForm(f => ({ ...f, [key]: val }));
  const updateOpt = (i, val) => setForm(f => ({ ...f, options: f.options.map((o, oi) => oi === i ? val : o) }));

  return (
    <div className="modal-backdrop" onClick={e => e.target === e.currentTarget && onClose()}>
      <div className="modal slide-in" style={{ maxWidth: 640, maxHeight: "85vh", overflowY: "auto" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 24 }}>
          <h2 style={{ fontWeight: 800, fontSize: 20 }}>{question ? "Modifier" : "Ajouter"} une question</h2>
          <button onClick={onClose} className="btn btn-secondary btn-sm">✕</button>
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 12, marginBottom: 16 }}>
          <div>
            <label>Catégorie</label>
            <select value={form.category} onChange={e => update("category", e.target.value)}>
              {["Python", "Machine Learning", "Big Data", "Statistiques", "Deep Learning"].map(c => <option key={c}>{c}</option>)}
            </select>
          </div>
          <div>
            <label>Difficulté</label>
            <select value={form.difficulty} onChange={e => update("difficulty", e.target.value)}>
              {["Facile", "Moyen", "Difficile"].map(d => <option key={d}>{d}</option>)}
            </select>
          </div>
          <div>
            <label>Type</label>
            <select value={form.type} onChange={e => update("type", e.target.value)}>
              {["qcm", "math", "code"].map(t => <option key={t}>{t}</option>)}
            </select>
          </div>
        </div>

        <div style={{ marginBottom: 16 }}>
          <label>Énoncé (supporte LaTeX avec $...$ et code avec ```python)</label>
          <textarea value={form.text} onChange={e => update("text", e.target.value)} rows={4} placeholder="Saisissez l'énoncé..." />
        </div>

        <div style={{ marginBottom: 16 }}>
          <label>Options de réponse</label>
          {(form.options || ["", "", "", ""]).map((opt, i) => (
            <div key={i} style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 8 }}>
              <input type="radio" name="answer" checked={form.answer === i} onChange={() => update("answer", i)} style={{ width: "auto", flexShrink: 0 }} />
              <input value={opt} onChange={e => updateOpt(i, e.target.value)} placeholder={`Option ${String.fromCharCode(65 + i)}`} />
            </div>
          ))}
          <p style={{ fontSize: 12, color: "#94a3b8" }}>🔘 Sélectionnez la bonne réponse</p>
        </div>

        <div style={{ marginBottom: 16 }}>
          <label>Explication (affichée après le test)</label>
          <textarea value={form.explanation} onChange={e => update("explanation", e.target.value)} rows={2} placeholder="Explication de la réponse..." />
        </div>

        <div style={{ marginBottom: 20 }}>
          <label>Points</label>
          <input type="number" value={form.score} onChange={e => update("score", Number(e.target.value))} style={{ maxWidth: 100 }} />
        </div>

        <div style={{ display: "flex", gap: 12, justifyContent: "flex-end" }}>
          <button onClick={onClose} className="btn btn-secondary">Annuler</button>
          <button onClick={() => onSave(form)} className="btn btn-primary">
            {question ? "Modifier" : "Ajouter"} la question
          </button>
        </div>
      </div>
    </div>
  );
}

// ─── APP ROOT ────────────────────────────────────────────────────────────────

export default function App() {
  const [screen, setScreen] = useState("landing");
  const [student, setStudent] = useState(null);
  const [result, setResult] = useState(null);

  useEffect(() => {
    if (!window.katex) {
      const link = document.createElement("link");
      link.rel = "stylesheet";
      link.href = "https://cdnjs.cloudflare.com/ajax/libs/KaTeX/0.16.9/katex.min.css";
      document.head.appendChild(link);
      const script = document.createElement("script");
      script.src = "https://cdnjs.cloudflare.com/ajax/libs/KaTeX/0.16.9/katex.min.js";
      script.async = true;
      document.head.appendChild(script);
    }
  }, []);

  const handleStart = (studentData) => {
    setStudent(studentData);
    setScreen("exam");
  };

  const handleFinish = (res) => {
    setResult(res);
    setScreen("results");
  };

  return (
    <>
      <style>{CSS}</style>
      {screen === "landing" && (
        <LandingPage onStart={handleStart} onAdmin={() => setScreen("adminLogin")} />
      )}
      {screen === "exam" && student && (
        <ExamInterface student={student} onFinish={handleFinish} />
      )}
      {screen === "results" && result && (
        <ResultsPage result={result} onHome={() => setScreen("landing")} />
      )}
      {screen === "adminLogin" && (
        <AdminLogin onLogin={() => setScreen("admin")} onBack={() => setScreen("landing")} />
      )}
      {screen === "admin" && (
        <AdminDashboard onLogout={() => setScreen("landing")} />
      )}
    </>
  );
}
