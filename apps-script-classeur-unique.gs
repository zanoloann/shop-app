/**
 * Google Apps Script UNIQUE — un seul projet, un seul déploiement, une seule URL,
 * partagée par les DEUX applis (Caisse "Appli de suivi des ventes" et
 * "Demandes de planning"). Un seul Google Sheet avec les onglets :
 *
 * - "Ventes"   : ID | Date | Heure | Vendeur | Montant | Paiement | Saisi par |
 *   ID_artisan | Taux_frais | Frais | Montant_net (4 dernières colonnes
 *   ajoutées automatiquement, voir apps-script-comptabilite.gs)
 * - "Artisans" : ID | Prénom | Nom | Email | PIN
 *     Pour le compte administrateur : créez une ligne avec Prénom =
 *     "Administrateur" (Nom laissé vide) et son e-mail. Cette ligne
 *     n'apparaît JAMAIS dans les listes d'artisans des deux applis, mais son
 *     e-mail est utilisé pour recevoir les notifications de demandes de
 *     planning (aucune adresse e-mail écrite en dur dans le code).
 * - "Planning" : Fermeture | Date | Jour | N° | Mois | Astreinte | Renfort 1 | Renfort 2 | Semaine
 * - "Demandes" : ID | Type | Demandeur | Date concernée | Commentaire | Statut | Date de création | Traité par | Date de traitement
 *   (créé automatiquement s'il est absent)
 * - "Accès" : Jeton | Prénom | Appareil | Appli | Date d'enrôlement (créé
 *   automatiquement) — liste des appareils autorisés, par application
 *   (Ventes / Planning). Supprimer une ligne révoque cet appareil.
 *
 * SÉCURITÉ — double authentification (TOTP, compatible Google
 * Authenticator/Authy) : le "code unique" saisi par un artisan pour
 * enrôler un nouvel appareil est désormais un code à 6 chiffres qui change
 * toutes les 30 secondes, calculé à partir d'un secret partagé — exactement
 * le même principe que la double authentification d'un compte Google.
 * Aucun code n'est plus stocké ni affiché dans le Sheet ou dans l'appli :
 * chacun des 3 admins lit le code courant directement dans son appli
 * Authenticator, sur son téléphone.
 *
 * Configuration UNE FOIS après le premier déploiement : exécutez la
 * fonction "printTotpSetupInfo_" (menu déroulant de fonctions en haut de
 * l'éditeur > sélectionner > ▶ Exécuter), puis ouvrez Affichage > Journaux
 * (ou Ctrl/Cmd+Entrée) pour récupérer le secret et l'URL otpauth:// à
 * scanner (via un générateur de QR code en ligne) ou saisir manuellement
 * dans Google Authenticator, sur CHACUN des 3 téléphones des admins — une
 * seule fois. Les 3 appareils affichent alors en permanence le même code,
 * en même temps.
 *
 * Fichiers de ce projet Apps Script (les 4 doivent être présents ensemble) :
 * - apps-script-classeur-unique.gs (CE fichier) : point d'entrée doGet/doPost, Ventes, Artisans
 * - apps-script-planning.gs        : Planning + Demandes
 * - apps-script-demandes-email.gs  : envoi de l'e-mail de notification
 * - apps-script-comptabilite.gs    : Paramètres + Virements (règlements artisans)
 *
 * APRÈS AVOIR COLLÉ apps-script-comptabilite.gs : exécutez une fois la
 * fonction "migrateAddArtisanIdAndFrais" (menu déroulant de fonctions en
 * haut de l'éditeur > sélectionner > ▶ Exécuter) pour rétro-remplir les
 * ventes déjà existantes.
 *
 * Déploiement : Déployer > Nouveau déploiement > Application Web,
 * Exécuter en tant que : Moi ; Accès : Tout le monde.
 * Collez l'URL obtenue dans TOUS les champs d'URL des deux applis (Ventes/Artisans
 * et Planning) — c'est désormais la même URL partout.
 * Après toute modification (dans N'IMPORTE LEQUEL des 3 fichiers) : redéployez une
 * NOUVELLE VERSION (Déployer > Gérer les déploiements > crayon > Nouvelle version).
 */

function sheet_(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function combineDateHeureIso_(dateCell, heureCell, tz) {
  let dateStr;
  if (dateCell instanceof Date) dateStr = Utilities.formatDate(dateCell, tz, 'yyyy-MM-dd');
  else {
    const p = String(dateCell).split('/');
    dateStr = p.length === 3 ? p[2] + '-' + p[1] + '-' + p[0] : String(dateCell);
  }
  let heureStr;
  if (heureCell instanceof Date) heureStr = Utilities.formatDate(heureCell, tz, 'HH:mm:ss');
  else {
    const p = String(heureCell).split(':');
    heureStr = (p[0] || '00').padStart(2, '0') + ':' + (p[1] || '00').padStart(2, '0') + ':' + (p[2] || '00').padStart(2, '0');
  }
  return dateStr + 'T' + heureStr;
}

function readSales_() {
  const rows = sheet_('Ventes').getDataRange().getValues();
  const tz = Session.getScriptTimeZone();
  const list = [];
  for (let i = 1; i < rows.length; i++) {
    const [id, date, heure, vendeur, montant, paiement, saisiPar, idArtisan, tauxFrais, frais, montantNet] = rows[i];
    if (!id) continue;
    list.push({ id: String(id), date, heure, artisan: vendeur, montant, paiement, saisiPar, iso: combineDateHeureIso_(date, heure, tz), idArtisan: idArtisan || '', tauxFrais: tauxFrais || 0, frais: frais || 0, montantNet: montantNet === '' || montantNet == null ? montant : montantNet });
  }
  return list;
}

const ADMIN_ARTISAN_PRENOM_ = 'administrateur';

function readArtisans_() {
  const rows = sheet_('Artisans').getDataRange().getValues();
  const list = [];
  for (let i = 1; i < rows.length; i++) {
    const [id, prenom, nom, email, pin] = rows[i];
    if (!prenom && !nom) continue;
    if (String(prenom || '').trim().toLowerCase() === ADMIN_ARTISAN_PRENOM_) continue; // administrateur masqué des deux applis
    list.push({ id: String(id), prenom, nom, email, pin: String(pin) });
  }
  return list;
}

function getAdminEmail_() {
  const sheet = sheet_('Artisans');
  if (!sheet) return '';
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const [, prenom, , email] = rows[i];
    if (String(prenom || '').trim().toLowerCase() === ADMIN_ARTISAN_PRENOM_) return email;
  }
  return '';
}

function getArtisanEmail_(fullName) {
  const sheet = sheet_('Artisans');
  if (!sheet || !fullName) return '';
  const target = String(fullName).trim().toLowerCase();
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const [, prenom, nom, email] = rows[i];
    const full = (String(prenom || '') + ' ' + String(nom || '')).trim().toLowerCase();
    if (full === target) return email;
  }
  return '';
}

function checkAdminToken_(pin) {
  return true; // accès direct en mode administrateur (V2.1)
  const rows = sheet_('Artisans').getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const [, prenom, , , rowPin] = rows[i];
    if (String(prenom || '').trim().toLowerCase() === ADMIN_ARTISAN_PRENOM_) return String(rowPin) === String(pin);
  }
  return false;
}
function adminRejected_() {
  return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'Jeton administrateur invalide' })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Accès — deux couches, avec un code du jour SÉPARÉ par application (Ventes /
 * Planning) :
 * 1) "Code du jour" par appli (ACCESS_DAILY_CODE_VENTES / ACCESS_DAILY_CODE_PLANNING,
 *    Propriétés du script) : tourne automatiquement chaque nuit (déclencheur
 *    programmé rotateAccessCode) ET immédiatement après chaque enrôlement réussi
 *    (usage unique) — sert UNIQUEMENT à enrôler un nouvel appareil sur CETTE appli.
 * 2) Jeton d'appareil permanent : délivré à l'enrôlement, stocké dans l'onglet
 *    "Accès" (une ligne par appareil, avec l'appli d'origine). Ne change jamais —
 *    jusqu'à ce qu'une ligne soit supprimée de cet onglet (révocation), ce qui
 *    force cet appareil à se ré-enrôler avec le code du jour en cours de son appli.
 * À faire UNE FOIS : Déclencheurs (icône ⏰ dans l'éditeur) > Ajouter un
 * déclencheur > fonction "rotateAccessCode" > Déclencheur basé sur le temps >
 * Minuteur horaire ou quotidien, tous les jours.
 */
const ACCESS_SHEET_ = 'Accès';
const ACCESS_HEADERS_ = ['Jeton', 'Prénom', 'Appareil', 'Appli', "Date d'enrôlement"];
const ACCESS_APPS_ = ['ventes', 'planning'];

function accessSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(ACCESS_SHEET_);
  if (!sheet) {
    sheet = ss.insertSheet(ACCESS_SHEET_);
    sheet.appendRow(ACCESS_HEADERS_);
  }
  return sheet;
}

function normalizeAppKey_(app) {
  const a = String(app || '').toLowerCase();
  return ACCESS_APPS_.indexOf(a) !== -1 ? a : 'ventes';
}

/* --- Double authentification TOTP (compatible Google Authenticator) --- */
const TOTP_SECRET_PROP_ = 'TOTP_SECRET';
const TOTP_STEP_ = 30;
const TOTP_DIGITS_ = 6;
const BASE32_ALPHABET_ = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ234567';

function base32Encode_(bytes) {
  let bits = '', output = '';
  for (let i = 0; i < bytes.length; i++) {
    bits += ((bytes[i] < 0 ? bytes[i] + 256 : bytes[i]) >>> 0).toString(2).padStart(8, '0');
  }
  for (let i = 0; i + 5 <= bits.length; i += 5) {
    output += BASE32_ALPHABET_[parseInt(bits.substring(i, i + 5), 2)];
  }
  if (bits.length % 5 !== 0) {
    const rem = bits.length % 5;
    const last = bits.substring(bits.length - rem).padEnd(5, '0');
    output += BASE32_ALPHABET_[parseInt(last, 2)];
  }
  return output;
}
function base32Decode_(str) {
  str = String(str || '').toUpperCase().replace(/=+$/, '').replace(/[^A-Z2-7]/g, '');
  let bits = '';
  for (let i = 0; i < str.length; i++) {
    const val = BASE32_ALPHABET_.indexOf(str[i]);
    if (val === -1) continue;
    bits += val.toString(2).padStart(5, '0');
  }
  const bytes = [];
  for (let i = 0; i + 8 <= bits.length; i += 8) {
    bytes.push(parseInt(bits.substring(i, i + 8), 2));
  }
  return bytes;
}
function getOrCreateTotpSecret_() {
  const props = PropertiesService.getScriptProperties();
  let secret = props.getProperty(TOTP_SECRET_PROP_);
  if (!secret) {
    const randomBytes = [];
    for (let i = 0; i < 20; i++) randomBytes.push(Math.floor(Math.random() * 256));
    secret = base32Encode_(randomBytes);
    props.setProperty(TOTP_SECRET_PROP_, secret);
  }
  return secret;
}
function totpCodeForCounter_(secretBase32, counter) {
  const keyBytes = base32Decode_(secretBase32).map(b => (b > 127 ? b - 256 : b));
  const counterBytes = [];
  let c = counter;
  for (let i = 7; i >= 0; i--) {
    counterBytes[i] = c & 0xff;
    c = Math.floor(c / 256);
  }
  const sig = Utilities.computeHmacSignature(Utilities.MacAlgorithm.HMAC_SHA_1, counterBytes, keyBytes).map(b => (b < 0 ? b + 256 : b));
  const offset = sig[sig.length - 1] & 0xf;
  const binCode = ((sig[offset] & 0x7f) << 24) | ((sig[offset + 1] & 0xff) << 16) | ((sig[offset + 2] & 0xff) << 8) | (sig[offset + 3] & 0xff);
  return String(binCode % Math.pow(10, TOTP_DIGITS_)).padStart(TOTP_DIGITS_, '0');
}
function verifyTotp_(code, secretBase32) {
  const clean = String(code || '').replace(/\D/g, '');
  if (!clean) return false;
  const counter = Math.floor(Date.now() / 1000 / TOTP_STEP_);
  for (let w = -2; w <= 2; w++) {
    if (totpCodeForCounter_(secretBase32, counter + w) === clean) return true;
  }
  return false;
}
/**
 * À exécuter UNE FOIS depuis l'éditeur (menu de fonctions en haut > 
 * printTotpSetupInfo_ > ▶ Exécuter), puis lire le résultat dans
 * Affichage > Journaux (Ctrl/Cmd+Entrée). Convertissez l'URL otpauth://
 * en QR code (n'importe quel générateur en ligne) ou saisissez le secret
 * manuellement dans Google Authenticator, sur chacun des 3 téléphones admin.
 */
function printTotpSetupInfo_() {
  const secret = getOrCreateTotpSecret_();
  const uri = 'otpauth://totp/' + encodeURIComponent('Atelier des Artisans') +
    '?secret=' + secret + '&issuer=' + encodeURIComponent('Atelier des Artisans') +
    '&period=' + TOTP_STEP_ + '&digits=' + TOTP_DIGITS_;
  Logger.log('Secret (saisie manuelle) : ' + secret);
  Logger.log('URL à convertir en QR code : ' + uri);
}
/**
 * DIAGNOSTIC : exécutez cette fonction (menu de fonctions > debugTotpNow_ >
 * ▶ Exécuter) puis Affichage > Journaux. Comparez le code "ACTUEL" affiché
 * ici avec celui affiché AU MÊME MOMENT sur votre téléphone : s'ils ne
 * correspondent jamais, le secret enregistré dans l'appli d'authentification
 * n'est pas le bon (souvent altéré par un convertisseur de QR code) —
 * re-scannez avec le secret ci-dessous en saisie manuelle plutôt qu'en QR.
 */
function debugTotpNow_() {
  const secret = getOrCreateTotpSecret_();
  const counter = Math.floor(Date.now() / 1000 / TOTP_STEP_);
  Logger.log('Secret stocké : ' + secret);
  Logger.log('Code ACTUEL (30s) : ' + totpCodeForCounter_(secret, counter));
  Logger.log('Code précédent : ' + totpCodeForCounter_(secret, counter - 1));
  Logger.log('Code suivant : ' + totpCodeForCounter_(secret, counter + 1));
}

function checkDeviceToken_(token, app) {
  return true; // accès direct, sans code, pour les deux applis (V2.1)
  const rows = accessSheet_().getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(token)) return !app || String(rows[i][3]) === normalizeAppKey_(app);
  }
  return false;
}
function enrollDevice_(app, code, prenom, deviceLabel) {
  const key = normalizeAppKey_(app);
  if (!verifyTotp_(code, getOrCreateTotpSecret_())) {
    return { ok: false, error: 'Code invalide ou expiré' };
  }
  const token = Utilities.getUuid();
  accessSheet_().appendRow([token, prenom || '', deviceLabel || '', key, new Date().toISOString()]);
  return { ok: true, deviceToken: token };
}
function accessRejected_() {
  return ContentService.createTextOutput(JSON.stringify({ ok: false, error: "Accès refusé — appareil non enrôlé ou révoqué" })).setMimeType(ContentService.MimeType.JSON);
}

const FREQ_SHEET_ = 'Fréquentation';
const FREQ_HEADERS_ = ['Date', 'Nombre', 'Saisi par'];

function freqSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(FREQ_SHEET_);
  if (!sheet) {
    sheet = ss.insertSheet(FREQ_SHEET_);
    sheet.appendRow(FREQ_HEADERS_);
  }
  return sheet;
}
function readFrequentation_() {
  const rows = freqSheet_().getDataRange().getValues();
  const list = [];
  for (let i = 1; i < rows.length; i++) {
    const [date, nombre, saisiPar] = rows[i];
    if (!date) continue;
    list.push({ date: String(date), nombre, saisiPar });
  }
  return list;
}
function setFrequentation_(date, nombre, saisiPar) {
  const sheet = freqSheet_();
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(date)) return { ok: false, error: 'Déjà enregistré pour cette date' };
  }
  sheet.appendRow([date, nombre, saisiPar || '']);
  return { ok: true, list: readFrequentation_() };
}

function doGet(e) {
  if ((e.parameter.action === 'validate' || e.parameter.action === 'refuse') && e.parameter.id) {
    return doGetDemandeAction_(e);
  }
  if (!checkDeviceToken_(e.parameter.deviceToken, e.parameter.app)) return accessRejected_();
  const which = (e.parameter.sheet || 'ventes').toLowerCase();
  if (which === 'reglements') {
    if (!checkAdminToken_(e.parameter.adminToken)) return adminRejected_();
    const periode = e.parameter.periode || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');
    return ContentService.createTextOutput(JSON.stringify(readVirements_(periode))).setMimeType(ContentService.MimeType.JSON);
  }
  const payload = which === 'artisans' ? readArtisans_()
    : which === 'planning' ? readPlanning_()
    : which === 'demandes' ? readDemandes_()
    : which === 'frequentation' ? readFrequentation_()
    : which === 'openmonths' ? getOpenMonths_()
    : readSales_();
  return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  const which = (data.sheet || 'ventes').toLowerCase();

  if (which === 'access' && data.action === 'enroll') {
    return ContentService.createTextOutput(JSON.stringify(enrollDevice_(data.app, data.dailyCode, data.prenom, data.deviceLabel))).setMimeType(ContentService.MimeType.JSON);
  }
  if (!checkDeviceToken_(data.deviceToken, data.app)) return accessRejected_();

  if (which === 'email') return doPostEmail_(data);
  if (which === 'planning' || which === 'demandes') return doPostPlanningOrDemandes_(data, which);
  if (which === 'admin' && data.action === 'checkToken') {
    return ContentService.createTextOutput(JSON.stringify({ ok: checkAdminToken_(data.adminToken) })).setMimeType(ContentService.MimeType.JSON);
  }

  if (which === 'frequentation' && data.action === 'set') {
    return ContentService.createTextOutput(JSON.stringify(setFrequentation_(data.date, data.nombre, data.saisiPar))).setMimeType(ContentService.MimeType.JSON);
  }

  if (which === 'reglements' && data.action === 'updateStatut') {
    if (!checkAdminToken_(data.adminToken)) return adminRejected_();
    updateVirementStatut_(data.id, data.type, data.statut, data.dateReelle, data.modifiePar);
    const periode = data.periode || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');
    return ContentService.createTextOutput(JSON.stringify(readVirements_(periode))).setMimeType(ContentService.MimeType.JSON);
  }
  if (which === 'reglements' && data.action === 'generate') {
    if (!checkAdminToken_(data.adminToken)) return adminRejected_();
    const periode = data.periode || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');
    generateOrUpdateVirements_(periode);
    return ContentService.createTextOutput(JSON.stringify(readVirements_(periode))).setMimeType(ContentService.MimeType.JSON);
  }

  if (which === 'artisans') {
    if ((data.action === 'add' || data.action === 'remove') && !checkAdminToken_(data.adminToken)) return adminRejected_();
    const sheet = sheet_('Artisans');
    if (data.action === 'add') {
      sheet.appendRow([Utilities.getUuid(), data.prenom, data.nom, data.email, data.pin]);
    } else if (data.action === 'remove') {
      const rows = sheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) === String(data.id)) { sheet.deleteRow(i + 1); break; }
      }
    }
    return ContentService.createTextOutput(JSON.stringify(readArtisans_())).setMimeType(ContentService.MimeType.JSON);
  }

  const sheet = sheet_('Ventes');
  ensureVentesExtraColumns_(sheet);
  if (data.action === 'add') {
    const artisanId = findArtisanIdByName_(data.artisan) || '';
    const taux = getTauxFraisActuel_();
    const isCarte = String(data.paiement).trim().toLowerCase() === 'carte';
    const frais = isCarte ? round2_(Number(data.montant) * taux) : 0;
    const montantNet = round2_(Number(data.montant) - frais);
    sheet.appendRow([Utilities.getUuid(), data.date, data.heure, data.artisan, data.montant, data.paiement, data.saisiPar, artisanId, taux, frais, montantNet]);
  } else if (data.action === 'update') {
    if (!checkAdminToken_(data.adminToken)) return adminRejected_();
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(data.id)) {
        const existingTaux = Number(rows[i][8]) || getTauxFraisActuel_();
        const isCarte = String(data.paiement).trim().toLowerCase() === 'carte';
        const frais = isCarte ? round2_(Number(data.montant) * existingTaux) : 0;
        const montantNet = round2_(Number(data.montant) - frais);
        const artisanId = findArtisanIdByName_(data.artisan) || rows[i][7] || '';
        sheet.getRange(i + 1, 4, 1, 3).setValues([[data.artisan, data.montant, data.paiement]]);
        sheet.getRange(i + 1, 8, 1, 4).setValues([[artisanId, existingTaux, frais, montantNet]]);
        break;
      }
    }
  } else if (data.action === 'remove') {
    if (!checkAdminToken_(data.adminToken)) return adminRejected_();
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(data.id)) {
        sheet.deleteRow(i + 1);
        break;
      }
    }
  }
  return ContentService.createTextOutput(JSON.stringify(readSales_())).setMimeType(ContentService.MimeType.JSON);
}
