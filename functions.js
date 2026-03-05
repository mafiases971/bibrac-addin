// BIBRAC Add-in â€” Fonctions personnalisÃ©es Excel
// Les fonctions sont disponibles dans Excel sous la forme =BIBRAC.NOM_FONCTION()

/**
 * Compte le nombre de cellules d'une plage ayant la mÃªme couleur de fond que la cellule cible.
 * @customfunction
 * @param {string} adresseCible Adresse de la cellule de rÃ©fÃ©rence (ex: "A1")
 * @param {string} adressePlage Adresse de la plage Ã  analyser (ex: "B1:D10")
 * @returns {Promise<number>} Nombre de cellules ayant la mÃªme couleur de fond
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

    // Chargement groupÃ© des couleurs de toutes les cellules
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
 * Envoie un texte Ã  un modÃ¨le d'IA avec une instruction et retourne la rÃ©ponse.
 * Les clÃ©s API sont gÃ©rÃ©es cÃ´tÃ© serveur (Cloudflare Worker) â€” aucune clÃ© Ã  saisir dans Excel.
 *
 * Fournisseurs disponibles :
 *   - claude  : Anthropic Claude Haiku        âœ… clÃ© configurÃ©e
 *   - gemini  : Google Gemini 2.0 Flash       âœ… clÃ© configurÃ©e
 *   - openai  : OpenAI GPT-4o Mini            â³ clÃ© Ã  ajouter dans Cloudflare (OPENAI_KEY)
 *   - grok    : xAI Grok 3                    â³ clÃ© Ã  ajouter dans Cloudflare (GROK_KEY)
 *   - llama   : Meta Llama 3.1 70B via Groq   â³ clÃ© Ã  ajouter dans Cloudflare (LLAMA_KEY)
 *
 * @customfunction
 * @param {string} texte Le texte Ã  traiter (contenu de la cellule)
 * @param {string} instruction L'instruction Ã  donner au modÃ¨le (ex: "Traduis en anglais")
 * @param {string} [provider] Fournisseur IA : claude, openai, gemini, grok, llama (dÃ©faut: gemini)
 * @returns {Promise<string>} La rÃ©ponse gÃ©nÃ©rÃ©e par le modÃ¨le
 */
async function AI(texte, instruction, provider) {
  // URL du proxy Cloudflare Worker â€” les clÃ©s API sont stockÃ©es dans Cloudflare
  const PROXY_URL = "https://fonction-excel.mafiases97-1.workers.dev";

  const response = await fetch(PROXY_URL, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "X-Bibrac-Token": "GjcZs7OJ0TEOk9BcwqXSmb8CintRoGz6zi0Sx7wsYC4"
    },
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

/**
 * Envoie une instruction a un modele d'IA et retourne le resultat sous forme de matrice Excel.
 * L'IA est automatiquement guidee pour repondre en tableau JSON 2D.
 *
 * @customfunction
 * @param {string} texte Le texte ou contexte a traiter
 * @param {string} instruction L'instruction decrivant la matrice souhaitee
 * @param {string} [provider] Fournisseur IA : claude, openai, gemini, grok, llama (defaut: gemini)
 * @returns {Promise<string[][]>} Tableau 2D retourne dans les cellules Excel
 */
async function AIMATRICE(texte, instruction, provider) {
  const PROXY_URL = "https://fonction-excel.mafiases97-1.workers.dev";
  const instructionMatrice = instruction +
    ". Reponds UNIQUEMENT avec un tableau JSON 2D valide, sans texte ni explication autour. " +
    "Exemple de format attendu : [[\"Colonne1\",\"Colonne2\"],[\"valeur1\",\"valeur2\"]]";
  const response = await fetch(PROXY_URL, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "X-Bibrac-Token": "GjcZs7OJ0TEOk9BcwqXSmb8CintRoGz6zi0Sx7wsYC4"
    },
    body: JSON.stringify({
      provider: provider || "gemini",
      texte: texte,
      instruction: instructionMatrice
    })
  });
  if (!response.ok) {
    const erreur = await response.json().catch(() => ({}));
    throw new Error("Erreur proxy (" + response.status + "): " + JSON.stringify(erreur));
  }
  const data = await response.json();
  try {
    let reponse = data.reponse.trim();
    reponse = reponse.replace(/```json\n?/g, "").replace(/```\n?/g, "").trim();
    const matrice = JSON.parse(reponse);
    if (!Array.isArray(matrice) || !Array.isArray(matrice[0])) {
      throw new Error("Format invalide");
    }
    return matrice;
  } catch (e) {
    throw new Error("La reponse n'est pas une matrice valide. Precisez davantage votre instruction.");
  }
}

CustomFunctions.associate("AIMATRICE", AIMATRICE);
