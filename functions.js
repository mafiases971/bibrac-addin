// BIBRAC Add-in — Fonctions personnalisées Excel
// Les fonctions sont disponibles dans Excel sous la forme =BIBRAC.NOM_FONCTION()

/**
 * Compte le nombre de cellules d'une plage ayant la même couleur de fond que la cellule cible.
 * @customfunction
 * @param {string} adresseCible Adresse de la cellule de référence (ex: "A1")
 * @param {string} adressePlage Adresse de la plage à analyser (ex: "B1:D10")
 * @returns {Promise<number>} Nombre de cellules ayant la même couleur de fond
 */
async function COMPTECOULEURS(adresseCible, adressePlage) {
  return Excel.run(async (context) => {
    const feuille = context.workbook.worksheets.getActiveWorksheet();

    // Chargement de la couleur de la cellule cible
    const celluleCible = feuille.getRange(adresseCible);
    celluleCible.load("format/fill/color");
    await context.sync();

    const couleurCible = celluleCible.format.fill.color;

    // Chargement des dimensions de la plage
    const plage = feuille.getRange(adressePlage);
    plage.load(["rowCount", "columnCount"]);
    await context.sync();

    // Chargement groupé des couleurs de toutes les cellules
    const cellules = [];
    for (let row = 0; row < plage.rowCount; row++) {
      for (let col = 0; col < plage.columnCount; col++) {
        const cellule = plage.getCell(row, col);
        cellule.load("format/fill/color");
        cellules.push(cellule);
      }
    }
    await context.sync();

    // Comptage
    let compteur = 0;
    for (const cellule of cellules) {
      if (cellule.format.fill.color === couleurCible) {
        compteur++;
      }
    }

    return compteur;
  });
}

CustomFunctions.associate("COMPTECOULEURS", COMPTECOULEURS);

/**
 * Envoie un texte à un modèle d'IA avec une instruction et retourne la réponse.
 * Les clés API sont gérées côté serveur (Cloudflare Worker) — aucune clé à saisir dans Excel.
 *
 * Fournisseurs disponibles :
 *   - claude  : Anthropic Claude Haiku        ✅ clé configurée
 *   - gemini  : Google Gemini 2.0 Flash       ✅ clé configurée
 *   - openai  : OpenAI GPT-4o Mini            ⏳ clé à ajouter dans Cloudflare (OPENAI_KEY)
 *   - grok    : xAI Grok 3                    ⏳ clé à ajouter dans Cloudflare (GROK_KEY)
 *   - llama   : Meta Llama 3.1 70B via Groq   ⏳ clé à ajouter dans Cloudflare (LLAMA_KEY)
 *
 * @customfunction
 * @param {string} texte Le texte à traiter (contenu de la cellule)
 * @param {string} instruction L'instruction à donner au modèle (ex: "Traduis en anglais")
 * @param {string} [provider] Fournisseur IA : claude, openai, gemini, grok, llama (défaut: gemini)
 * @returns {Promise<string>} La réponse générée par le modèle
 */
async function AI(texte, instruction, provider) {
  // URL du proxy Cloudflare Worker — les clés API sont stockées dans Cloudflare
  const PROXY_URL = "https://fonction-excel.mafiases97-1.workers.dev";

  const response = await fetch(PROXY_URL, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      provider: provider || "gemini",
      texte: texte,
      instruction: instruction
    })
  });

  if (!response.ok) {
    const erreur = await response.json().catch(() => ({}));
    throw new Error(`Erreur proxy (${response.status}): ${JSON.stringify(erreur)}`);
  }

  const data = await response.json();
  return data.reponse;
}

CustomFunctions.associate("AI", AI);
