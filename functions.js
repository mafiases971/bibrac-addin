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
 * La clé API doit être stockée dans une plage nommée Excel (ex: BIBRAC_OPENAI_KEY).
 * @customfunction
 * @param {string} texte Le texte à traiter (contenu de la cellule)
 * @param {string} instruction L'instruction à donner au modèle (ex: "Traduis en anglais")
 * @param {string} [provider] Fournisseur IA : claude, openai, gemini, grok, llama (défaut: openai)
 * @returns {Promise<string>} La réponse générée par le modèle
 */
async function AI(texte, instruction, provider) {
  const providerNorm = (provider || "openai").toLowerCase().trim();

  // Récupération de la clé API depuis une plage nommée Excel
  // Ex: créer une plage nommée "BIBRAC_OPENAI_KEY" avec votre clé API
  const nomCle = `BIBRAC_${providerNorm.toUpperCase()}_KEY`;
  let cleApi = "";

  try {
    await Excel.run(async (context) => {
      const namedItem = context.workbook.names.getItem(nomCle);
      namedItem.load("value");
      await context.sync();
      cleApi = String(namedItem.value).trim();
    });
  } catch (e) {
    throw new Error(`Clé API manquante. Créez une plage nommée "${nomCle}" dans Excel avec votre clé API.`);
  }

  if (!cleApi) {
    throw new Error(`La plage nommée "${nomCle}" est vide.`);
  }

  // Configuration de chaque fournisseur
  let url, headers, body, extraireReponse;

  switch (providerNorm) {

    case "claude":
      url = "https://api.anthropic.com/v1/messages";
      headers = {
        "Content-Type": "application/json",
        "x-api-key": cleApi,
        "anthropic-version": "2023-06-01"
      };
      body = {
        model: "claude-haiku-4-5-20251001",
        max_tokens: 1024,
        messages: [{ role: "user", content: `${instruction}\n\n${texte}` }]
      };
      extraireReponse = (data) => data.content[0].text;
      break;

    case "gemini":
      url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${cleApi}`;
      headers = { "Content-Type": "application/json" };
      body = {
        contents: [{ parts: [{ text: `${instruction}\n\n${texte}` }] }]
      };
      extraireReponse = (data) => data.candidates[0].content.parts[0].text;
      break;

    case "grok":
      url = "https://api.x.ai/v1/chat/completions";
      headers = {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${cleApi}`
      };
      body = {
        model: "grok-3",
        messages: [{ role: "user", content: `${instruction}\n\n${texte}` }]
      };
      extraireReponse = (data) => data.choices[0].message.content;
      break;

    case "llama":
      url = "https://api.groq.com/openai/v1/chat/completions";
      headers = {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${cleApi}`
      };
      body = {
        model: "llama-3.1-70b-versatile",
        messages: [{ role: "user", content: `${instruction}\n\n${texte}` }]
      };
      extraireReponse = (data) => data.choices[0].message.content;
      break;

    case "openai":
    default:
      url = "https://api.openai.com/v1/chat/completions";
      headers = {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${cleApi}`
      };
      body = {
        model: "gpt-4o-mini",
        messages: [{ role: "user", content: `${instruction}\n\n${texte}` }]
      };
      extraireReponse = (data) => data.choices[0].message.content;
      break;
  }

  // Appel API
  const response = await fetch(url, {
    method: "POST",
    headers: headers,
    body: JSON.stringify(body)
  });

  if (!response.ok) {
    const erreur = await response.json().catch(() => ({}));
    throw new Error(`Erreur ${providerNorm} (${response.status}): ${JSON.stringify(erreur)}`);
  }

  const data = await response.json();
  return extraireReponse(data).trim();
}

CustomFunctions.associate("AI", AI);
