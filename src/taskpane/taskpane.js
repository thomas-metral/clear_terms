import { apiKey } from "./config.js";

Office.onReady(() => {
  console.log("Add-in prêt.");
});

// 🔄 Rafraîchir l’add-in avec message temporaire
function refreshPane() {
  const messageDiv = document.getElementById("refreshMessage");
  messageDiv.style.display = "block";
  setTimeout(() => {
    messageDiv.style.display = "none";
    location.reload();
  }, 2000);
}

// 📝 Réécriture du contenu d’une cellule sélectionnée
async function rewriteSelectedCell() {
  Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, async (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const originalText = result.value.trim();

      if (!originalText) {
        document.getElementById("output").innerText = "⚠️ Sélectionnez une cellule contenant du texte.";
        return;
      }

      document.getElementById("output").innerText = "Réécriture en cours...";

      const rewrittenText = await callChatGPT(originalText);

      Office.context.document.setSelectedDataAsync(rewrittenText, {
        coercionType: Office.CoercionType.Text
      });

      document.getElementById("output").innerText = "✅ Réécriture insérée.";
    } else {
      document.getElementById("output").innerText = "⚠️ Impossible de lire la cellule sélectionnée.";
    }
  });
}

// 🤖 Appel à l’API OpenAI (instructions personnalisées)
async function callChatGPT(inputText) {
  const endpoint = "https://api.openai.com/v1/chat/completions";

  const response = await fetch(endpoint, {
    method: "POST",
    headers: {
      "Authorization": `Bearer ${apiKey}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      model: "gpt-4", // ou "gpt-3.5-turbo"
      messages: [
        {
          role: "system",
          content: `Tu es un expert du langage clair (norme ISO 24495-1:2023) et de la méthode FALC. Tu aides des compagnies d'assurance à réécrire leurs contrats pour les rendre plus clairs. Conserve les notions juridiques et renvois. Analyse le texte et explique les manquements aux règles (en citant les règles), puis propose une réécriture simplifiée qui respecte :
- le langage clair (guide du gouvernement français)
- la méthode FALC
- les règles du projet In Clear Terms :
  * < 5% de mots juridiques
  * Paragraphes < 15 mots
  * 75% du texte = conseils pratiques
  * Utilise le présent, évite les adverbes en -ment
  * Donne des exemples quand c’est utile.`
        },
        {
          role: "user",
          content: `Réécris ce texte : ${inputText}`
        }
      ]
    })
  });

  const data = await response.json();

  if (data.choices && data.choices.length > 0) {
    return data.choices[0].message.content;
  } else if (data.error) {
    return `❌ Erreur API : ${data.error.message}`;
  } else {
    return "❌ Réponse inattendue de l'API.";
  }
}

// Rendre les fonctions accessibles depuis le HTML
window.rewriteSelectedCell = rewriteSelectedCell;
window.refreshPane = refreshPane;
