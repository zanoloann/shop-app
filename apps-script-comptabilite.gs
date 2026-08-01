/**
 * apps-script-comptabilite.gs — Onglets Paramètres et Virements (règlements
 * artisans). Fait partie du même projet Apps Script UNIQUE que
 * apps-script-classeur-unique.gs (point d'entrée doGet/doPost) et
 * apps-script-planning.gs. Ne PAS déployer séparément.
 *
 * - "Ventes" gagne 4 colonnes : ID_artisan | Taux_frais | Frais | Montant_net
 *   (ajoutées automatiquement en fin de ligne si absentes — voir
 *   ensureVentesExtraColumns_). Le taux est figé à l'écriture de la vente,
 *   jamais recalculé depuis Paramètres après coup.
 * - "Paramètres" : Champ | Valeur — Taux_frais_actuel (ex 0.015 = 1,5%),
 *   Jour_versement (ex 5). Créé automatiquement avec ces valeurs par défaut.
 * - "Virements" : une ligne par artisan × mois calendaire, créée par
 *   generateOrUpdateVirements_ si absente. Les colonnes d'agrégation
 *   (Date_echeance, CA_carte, Frais, Net_a_virer, CA_especes, CA_cheque,
 *   Nb_ventes) sont des FORMULES SUMIFS/COUNTIFS posées directement dans le
 *   Sheet (pointant vers Ventes et Paramètres) — aucun calcul n'est fait en
 *   JS, elles se recalculent seules à l'ouverture ou à toute modification de
 *   Ventes. Les statuts (Viré/Donné/Remis) et leurs dates sont éditables à
 *   la main ou via l'app, et jamais retouchés une fois la ligne créée.
 *
 * MIGRATION UNE FOIS (après avoir collé ce fichier) : ouvrez l'éditeur Apps
 * Script, sélectionnez la fonction "migrateAddArtisanIdAndFrais" dans le
 * menu déroulant en haut, et cliquez ▶ Exécuter. Elle ajoute les 4 colonnes
 * à "Ventes" si besoin et recalcule ID_artisan/Frais/Montant_net pour
 * toutes les ventes déjà existantes (le Taux_frais historique est approximé
 * avec le taux actuel de Paramètres, faute d'historique).
 */

const PARAMS_SHEET_ = 'Paramètres';
const PARAMS_DEFAULTS_ = { Taux_frais_actuel: 0.015, Jour_versement: 5 };

function paramsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(PARAMS_SHEET_);
  if (!sheet) {
    sheet = ss.insertSheet(PARAMS_SHEET_);
    sheet.appendRow(['Champ', 'Valeur']);
    Object.keys(PARAMS_DEFAULTS_).forEach(k => sheet.appendRow([k, PARAMS_DEFAULTS_[k]]));
  }
  return sheet;
}
function getParam_(champ) {
  const rows = paramsSheet_().getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === champ) return rows[i][1];
  }
  return PARAMS_DEFAULTS_[champ];
}
function getTauxFraisActuel_() {
  const v = parseFloat(getParam_('Taux_frais_actuel'));
  return isNaN(v) ? PARAMS_DEFAULTS_.Taux_frais_actuel : v;
}
function getJourVersement_() {
  const v = parseInt(getParam_('Jour_versement'), 10);
  return isNaN(v) ? PARAMS_DEFAULTS_.Jour_versement : v;
}

function round2_(n) { return Math.round((Number(n) || 0) * 100) / 100; }

function findArtisanIdByName_(fullName) {
  if (!fullName) return '';
  const target = String(fullName).trim().toLowerCase();
  const rows = sheet_('Artisans').getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const [id, prenom, nom] = rows[i];
    const full = (String(prenom || '') + ' ' + String(nom || '')).trim().toLowerCase();
    if (full === target) return String(id);
  }
  return '';
}
function getArtisanById_(id) {
  if (!id) return null;
  const rows = sheet_('Artisans').getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) return { id: String(rows[i][0]), prenom: rows[i][1], nom: rows[i][2] };
  }
  return null;
}

const VENTES_EXTRA_HEADERS_ = ['ID_artisan', 'Taux_frais', 'Frais', 'Montant_net'];
function ensureVentesExtraColumns_(sheet) {
  const lastCol = sheet.getLastColumn();
  const header = sheet.getRange(1, 1, 1, Math.max(lastCol, 7)).getValues()[0];
  if (header.length >= 11 && header[7] === 'ID_artisan') return;
  sheet.getRange(1, 8, 1, 4).setValues([VENTES_EXTRA_HEADERS_]);
}

function migrateAddArtisanIdAndFrais() {
  const sheet = sheet_('Ventes');
  ensureVentesExtraColumns_(sheet);
  const rows = sheet.getDataRange().getValues();
  const taux = getTauxFraisActuel_();
  let updated = 0;
  for (let i = 1; i < rows.length; i++) {
    const [id, , , vendeur, montant, paiement] = rows[i];
    if (!id) continue;
    const artisanId = findArtisanIdByName_(vendeur);
    const frais = String(paiement).trim().toLowerCase() === 'carte' ? round2_(Number(montant) * taux) : 0;
    const montantNet = round2_(Number(montant) - frais);
    sheet.getRange(i + 1, 8, 1, 4).setValues([[artisanId, taux, frais, montantNet]]);
    updated++;
  }
  Logger.log('Migration terminée : ' + updated + ' vente(s) mise(s) à jour.');
  return updated;
}

const VIREMENTS_SHEET_ = 'Virements';
const VIREMENTS_HEADERS_ = [
  'ID', 'ID_artisan', 'Nom', 'Periode', 'Date_echeance', 'CA_carte', 'Frais', 'Net_a_virer',
  'CA_especes', 'CA_cheque', 'Nb_ventes', 'Statut_virement', 'Date_virement_reelle', 'Statut_cash',
  'Date_remise_cash_reelle', 'Statut_cheque', 'Date_remise_cheque_reelle', 'Modifie_par'
];
function virementsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(VIREMENTS_SHEET_);
  if (!sheet) {
    sheet = ss.insertSheet(VIREMENTS_SHEET_);
    sheet.appendRow(VIREMENTS_HEADERS_);
  }
  return sheet;
}

function periodeOfDate_(dateCell) {
  let d;
  if (dateCell instanceof Date) d = dateCell;
  else {
    const p = String(dateCell).split('/');
    d = p.length === 3 ? new Date(Number(p[2]), Number(p[1]) - 1, Number(p[0])) : new Date(dateCell);
  }
  if (isNaN(d)) return null;
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
}

/**
 * Crée, si besoin, une ligne par artisan pour la période donnée (y compris
 * 0€). Les colonnes d'agrégation (Date_echeance, CA_carte, Frais,
 * Net_a_virer, CA_especes, CA_cheque, Nb_ventes) sont des FORMULES
 * SUMIFS/COUNTIFS pointant directement vers l'onglet Ventes et
 * Paramètres — aucun calcul n'est fait ici en JS, elles se recalculent
 * seules à chaque ouverture du Sheet. Les lignes déjà créées ne sont pas
 * retouchées (statuts et dates réelles préservés).
 */
/**
 * Crée, si besoin, une ligne par artisan pour la période donnée (y compris
 * 0€), et RAFRAÎCHIT les formules (colonnes E-K) des lignes déjà présentes
 * — sans jamais toucher aux statuts/dates réelles (colonnes L-R). Le
 * script pose lui-même les formules SUMIFS/COUNTIFS (aucune ligne modèle
 * à remplir à la main). Si plusieurs lignes existent déjà pour un même
 * artisan sur la même période (doublon), seule la première est conservée
 * et rafraîchie, les suivantes sont supprimées.
 */
function generateOrUpdateVirements_(periode) {
  // Verrou : deux appels quasi simultanés (double-clic, requêtes concurrentes)
  // liraient chacun le Sheet avant que l'autre n'ait écrit ses lignes,
  // créant des doublons. Le lock sérialise les exécutions.
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    generateOrUpdateVirementsLocked_(periode);
  } finally {
    lock.releaseLock();
  }
}

function virementFormulasFor_(rowNum) {
  const R = rowNum;
  const monthStart = `DATE(VALUE(LEFT(D${R};4));VALUE(MID(D${R};6;2));1)`;
  const monthEnd = `DATE(VALUE(LEFT(D${R};4));VALUE(MID(D${R};6;2))+1;1)`;
  return [
    `=DATE(VALUE(LEFT(D${R};4));VALUE(MID(D${R};6;2))+1;VLOOKUP("Jour_versement";Paramètres!A:B;2;FALSE))`,
    `=SUMIFS(Ventes!E:E;Ventes!H:H;B${R};Ventes!F:F;"Carte";Ventes!B:B;">="&${monthStart};Ventes!B:B;"<"&${monthEnd})`,
    `=SUMIFS(Ventes!J:J;Ventes!H:H;B${R};Ventes!F:F;"Carte";Ventes!B:B;">="&${monthStart};Ventes!B:B;"<"&${monthEnd})`,
    `=F${R}-G${R}`,
    `=SUMIFS(Ventes!E:E;Ventes!H:H;B${R};Ventes!F:F;"Espèces";Ventes!B:B;">="&${monthStart};Ventes!B:B;"<"&${monthEnd})`,
    `=SUMIFS(Ventes!E:E;Ventes!H:H;B${R};Ventes!F:F;"Chèque";Ventes!B:B;">="&${monthStart};Ventes!B:B;"<"&${monthEnd})`,
    `=COUNTIFS(Ventes!H:H;B${R};Ventes!B:B;">="&${monthStart};Ventes!B:B;"<"&${monthEnd})`
  ];
}

function generateOrUpdateVirementsLocked_(periode) {
  const artisans = readArtisans_(); // masque déjà la ligne Administrateur
  const sheet = virementsSheet_();

  // 1) Dédoublonnage : pour un même artisan sur cette période, ne garder
  // que la première ligne rencontrée, supprimer les autres (de bas en
  // haut pour ne pas décaler les numéros de ligne en cours de suppression).
  let rows = sheet.getDataRange().getValues();
  const firstRowByArtisan = {};
  const rowsToDelete = [];
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][3]) !== periode) continue;
    const idArtisan = String(rows[i][1]);
    if (firstRowByArtisan[idArtisan] == null) firstRowByArtisan[idArtisan] = i + 1;
    else rowsToDelete.push(i + 1);
  }
  rowsToDelete.sort((a, b) => b - a).forEach(r => sheet.deleteRow(r));

  // 2) Relit l'état à jour après dédoublonnage.
  rows = sheet.getDataRange().getValues();
  const existingRowByArtisan = {};
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][3]) === periode) existingRowByArtisan[String(rows[i][1])] = i + 1;
  }

  artisans.forEach(a => {
    const existingRow = existingRowByArtisan[a.id];
    if (existingRow) {
      // Ligne déjà là : on rafraîchit seulement les formules, statuts et
      // dates réelles restent inchangés.
      sheet.getRange(existingRow, 5, 1, 7).setFormulas([virementFormulasFor_(existingRow)]);
      return;
    }
    const rowNum = sheet.getLastRow() + 1;
    const nom = (a.prenom + ' ' + a.nom).trim();
    sheet.getRange(rowNum, 1, 1, 4).setValues([[Utilities.getUuid(), a.id, nom, periode]]);
    sheet.getRange(rowNum, 5, 1, 7).setFormulas([virementFormulasFor_(rowNum)]);
    sheet.getRange(rowNum, 12, 1, 7).setValues([['À virer', '', 'À donner', '', 'À remettre', '', '']]);
  });
  SpreadsheetApp.flush();
}

function readVirements_(periode) {
  // Ne crée plus les lignes automatiquement — voir generateOrUpdateVirements_,
  // appelée uniquement sur demande explicite (action "generate").
  const rows = virementsSheet_().getDataRange().getValues();
  const list = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0] || String(r[3]) !== periode) continue;
    list.push({
      id: String(r[0]), idArtisan: String(r[1]), nom: r[2], periode: r[3], dateEcheance: r[4],
      caCarte: Number(r[5]) || 0, frais: Number(r[6]) || 0, netAVirer: Number(r[7]) || 0,
      caEspeces: Number(r[8]) || 0, caCheque: Number(r[9]) || 0, nbVentes: Number(r[10]) || 0,
      statutVirement: r[11] || 'À virer', dateVirementReelle: r[12] || '',
      statutCash: r[13] || 'À donner', dateRemiseCashReelle: r[14] || '',
      statutCheque: r[15] || 'À remettre', dateRemiseChequeReelle: r[16] || '', modifiePar: r[17] || ''
    });
  }
  return list;
}

function updateVirementStatut_(id, type, statut, dateReelle, modifiePar) {
  const sheet = virementsSheet_();
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) {
      if (type === 'cash') sheet.getRange(i + 1, 14, 1, 2).setValues([[statut, dateReelle || '']]);
      else if (type === 'cheque') sheet.getRange(i + 1, 16, 1, 2).setValues([[statut, dateReelle || '']]);
      else sheet.getRange(i + 1, 12, 1, 2).setValues([[statut, dateReelle || '']]);
      sheet.getRange(i + 1, 18).setValue(modifiePar || '');
      break;
    }
  }
  return true;
}
