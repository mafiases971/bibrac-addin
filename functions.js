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
