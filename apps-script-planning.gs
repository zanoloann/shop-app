/**
 * Planning + Demandes — fonctions appelées par le dispatcheur doGet/doPost de
 * apps-script-classeur-unique.gs (même projet, même classeur, même déploiement).
 * Onglet "Planning" : Fermeture | Date | Jour | N° | Mois | Astreinte | Renfort 1 | Renfort 2 | Semaine
 * Onglet "Demandes" (créé automatiquement si absent) :
 *   ID | Type | Demandeur | Date concernée | Commentaire | Statut | Date de création | Traité par | Date de traitement
 */

const DEMANDES_HEADERS_ = ['ID', 'Type', 'Demandeur', 'Date concernée', 'Commentaire', 'Statut', 'Date de création', 'Traité par', 'Date de traitement'];

function demandesSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Demandes');
  if (!sheet) {
    sheet = ss.insertSheet('Demandes');
    sheet.appendRow(DEMANDES_HEADERS_);
  }
  return sheet;
}

function readPlanning_() {
  const sheet = sheet_('Planning');
  const rows = sheet ? sheet.getDataRange().getValues() : [];
  const tz = Session.getScriptTimeZone();
  const list = [];
  for (let i = 1; i < rows.length; i++) {
    const [fermeture, date, jour, numero, mois, astreinte, renfort1, renfort2, semaine] = rows[i];
    if (!date) continue;
    const dateStr = date instanceof Date ? Utilities.formatDate(date, tz, 'dd/MM/yyyy') : String(date);
    list.push({
      fermeture: !!fermeture && String(fermeture).trim() !== '',
      date: dateStr, jour, numero: String(numero || ''), mois,
      astreinte, renfort1, renfort2, semaine: String(semaine || '')
    });
  }
  return list;
}

function readDemandes_() {
  const sheet = demandesSheet_();
  const rows = sheet.getDataRange().getValues();
  const tz = Session.getScriptTimeZone();
  const list = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0]) continue;
    const dateConcernee = r[3] instanceof Date ? Utilities.formatDate(r[3], tz, 'dd/MM/yyyy') : String(r[3]);
    list.push({
      id: r[0], type: r[1], demandeur: r[2], dateConcernee,
      commentaire: r[4], statut: r[5], dateCreation: r[6], traitePar: r[7], dateTraitement: r[8]
    });
  }
  return list;
}

function planningColForRole_(role) {
  return { astreinte: 6, renfort1: 7, renfort2: 8 }[role] || 0;
}

const OPEN_MONTHS_PROP_ = 'OPEN_MONTHS';
function getOpenMonths_() {
  const raw = PropertiesService.getScriptProperties().getProperty(OPEN_MONTHS_PROP_);
  try { return raw ? JSON.parse(raw) : []; } catch (e) { return []; }
}
function setMonthOpen_(monthKey, open) {
  const months = getOpenMonths_();
  const idx = months.indexOf(monthKey);
  if (open && idx === -1) months.push(monthKey);
  if (!open && idx !== -1) months.splice(idx, 1);
  PropertiesService.getScriptProperties().setProperty(OPEN_MONTHS_PROP_, JSON.stringify(months));
  return months;
}

// true si aucune astreinte/renfort n'est assigné ce jour-là (et boutique pas fermée).
function isDayFullyEmpty_(dateStr) {
  const sheet = sheet_('Planning');
  if (!sheet) return false;
  const rows = sheet.getDataRange().getValues();
  const tz = Session.getScriptTimeZone();
  for (let i = 1; i < rows.length; i++) {
    const dateCell = rows[i][1];
    const ds = dateCell instanceof Date ? Utilities.formatDate(dateCell, tz, 'dd/MM/yyyy') : String(dateCell);
    if (ds === dateStr) {
      const fermeture = rows[i][0];
      if (fermeture && String(fermeture).trim() !== '') return false;
      return !rows[i][5] && !rows[i][6] && !rows[i][7];
    }
  }
  return false;
}

// Synthèse texte du planning pour une date donnée (dd/MM/yyyy) — utilisée dans l'e-mail de notification.
function buildDaySummary_(dateStr) {
  const sheet = sheet_('Planning');
  if (!sheet) return 'Planning introuvable.';
  const rows = sheet.getDataRange().getValues();
  const tz = Session.getScriptTimeZone();
  for (let i = 1; i < rows.length; i++) {
    const dateCell = rows[i][1];
    const ds = dateCell instanceof Date ? Utilities.formatDate(dateCell, tz, 'dd/MM/yyyy') : String(dateCell);
    if (ds === dateStr) {
      const fermeture = rows[i][0];
      if (fermeture && String(fermeture).trim() !== '') return 'Boutique fermée ce jour-là.';
      return 'Astreinte : ' + (rows[i][5] || '(personne)') + '\nRenfort 1 : ' + (rows[i][6] || '(personne)') + '\nRenfort 2 : ' + (rows[i][7] || '(personne)');
    }
  }
  return 'Aucune information trouvée pour cette date dans le Planning.';
}

function setPlanningRole_(dateStr, role, value) {
  const sheet = sheet_('Planning');
  const col = planningColForRole_(role);
  if (!sheet || !col) return;
  const rows = sheet.getDataRange().getValues();
  const tz = Session.getScriptTimeZone();
  for (let i = 1; i < rows.length; i++) {
    const dateCell = rows[i][1];
    const ds = dateCell instanceof Date ? Utilities.formatDate(dateCell, tz, 'dd/MM/yyyy') : String(dateCell);
    if (ds === dateStr) { sheet.getRange(i + 1, col).setValue(value || ''); break; }
  }
}

function normalizeName_(v) { return String(v || '').trim().toLowerCase(); }

function getPlanningRoleValue_(dateStr, role) {
  const sheet = sheet_('Planning');
  const col = planningColForRole_(role);
  if (!sheet || !col) return '';
  const rows = sheet.getDataRange().getValues();
  const tz = Session.getScriptTimeZone();
  for (let i = 1; i < rows.length; i++) {
    const dateCell = rows[i][1];
    const ds = dateCell instanceof Date ? Utilities.formatDate(dateCell, tz, 'dd/MM/yyyy') : String(dateCell);
    if (ds === dateStr) return rows[i][col - 1] || '';
  }
  return '';
}

// Met à jour le statut d'une demande et, si validée, applique le changement au Planning.
// Utilisé à la fois par l'appli (bouton Valider/Refuser) et par les liens de l'e-mail.
// Si la validation d'une prise de poste tombe sur un créneau déjà occupé par
// quelqu'un d'autre (deux demandes simultanées pour le même jour), la demande
// reste "En attente" et un conflit est retourné à l'appelant.
function applyDemandeStatus_(id, statut, traitePar) {
  const sheet = demandesSheet_();
  const rows = sheet.getDataRange().getValues();
  let conflict = null;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) {
      const type = rows[i][1];
      const dateCell = rows[i][3];
      const dateConcernee = dateCell instanceof Date ? Utilities.formatDate(dateCell, Session.getScriptTimeZone(), 'dd/MM/yyyy') : String(dateCell);
      const demandeur = rows[i][2];
      if (statut === 'Validée') {
        const isLiberation = String(type || '').startsWith('liberation_');
        const role = isLiberation ? type.replace('liberation_', '') : type;
        if (['astreinte', 'renfort1', 'renfort2'].indexOf(role) !== -1) {
          if (!isLiberation) {
            const current = getPlanningRoleValue_(dateConcernee, role);
            if (current && normalizeName_(current) !== normalizeName_(demandeur)) {
              conflict = { occupiedBy: current, role, dateConcernee, demandeur };
            }
          }
          if (!conflict) setPlanningRole_(dateConcernee, role, isLiberation ? '' : demandeur);
        }
      }
      if (!conflict) {
        sheet.getRange(i + 1, 6, 1, 4).setValues([[statut, rows[i][6], traitePar || '', new Date().toISOString()]]);
        notifyDemandeurStatus_(demandeur, type, dateConcernee, statut);
      }
      break;
    }
  }
  return { list: readDemandes_(), conflict };
}

function doGetDemandeAction_(e) {
  const result = applyDemandeStatus_(e.parameter.id, e.parameter.action === 'validate' ? 'Validée' : 'Refusée', 'Administrateur (e-mail)');
  const msg = result.conflict
    ? 'Conflit ⚠ — ce créneau (' + result.conflict.role + ' du ' + result.conflict.dateConcernee + ') est déjà occupé par ' + result.conflict.occupiedBy + '. La demande de ' + result.conflict.demandeur + ' reste en attente ; traitez-la manuellement depuis l\'appli.'
    : (e.parameter.action === 'validate' ? 'Demande validée ✓' : 'Demande refusée ✕');
  return HtmlService.createHtmlOutput(
    '<div style="font-family:sans-serif;padding:40px;text-align:center;font-size:18px;">' + msg + '<br><span style="font-size:13px;color:#888;">Vous pouvez fermer cette page.</span></div>'
  );
}

function notifyNewDemande_(id, type, demandeur, dateConcernee, commentaire) {
  const base = ScriptApp.getService().getUrl();
  sendNewDemandeEmail_({
    demandeur, type, dateConcernee, commentaire,
    daySummary: buildDaySummary_(dateConcernee || ''),
    validateUrl: base + '?action=validate&id=' + encodeURIComponent(id),
    refuseUrl: base + '?action=refuse&id=' + encodeURIComponent(id)
  });
}

function doPostPlanningOrDemandes_(data, which) {
  if (which === 'planning' && data.action === 'setRole') {
    if (!checkAdminToken_(data.adminToken)) return adminRejected_();
    setPlanningRole_(data.date, data.role, data.value);
    return ContentService.createTextOutput(JSON.stringify(readPlanning_())).setMimeType(ContentService.MimeType.JSON);
  }
  if (which === 'planning' && data.action === 'toggleMonth') {
    if (!checkAdminToken_(data.adminToken)) return adminRejected_();
    const months = setMonthOpen_(data.month, !!data.open);
    return ContentService.createTextOutput(JSON.stringify(months)).setMimeType(ContentService.MimeType.JSON);
  }
  const sheet = demandesSheet_();
  if (data.action === 'add') {
    const id = Utilities.getUuid();
    const isLiberation = String(data.type || '').startsWith('liberation_');
    if (!isLiberation && isDayFullyEmpty_(data.dateConcernee)) {
      sheet.appendRow([id, data.type, data.demandeur, data.dateConcernee || '', data.commentaire || '', 'Validée', new Date().toISOString(), 'Auto (jour libre)', new Date().toISOString()]);
      setPlanningRole_(data.dateConcernee, data.type, data.demandeur);
      sendAutoRegisteredEmail_(data.demandeur, data.type, data.dateConcernee);
    } else {
      sheet.appendRow([id, data.type, data.demandeur, data.dateConcernee || '', data.commentaire || '', 'En attente', new Date().toISOString(), '', '']);
      notifyNewDemande_(id, data.type, data.demandeur, data.dateConcernee, data.commentaire);
    }
  } else if (data.action === 'status') {
    if (!checkAdminToken_(data.adminToken)) return adminRejected_();
    return ContentService.createTextOutput(JSON.stringify(applyDemandeStatus_(data.id, data.statut, data.traitePar))).setMimeType(ContentService.MimeType.JSON);
  } else if (data.action === 'remove') {
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(data.id)) { sheet.deleteRow(i + 1); break; }
    }
  } else if (data.action === 'clearAll') {
    if (!checkAdminToken_(data.adminToken)) return adminRejected_();
    const last = sheet.getLastRow();
    if (last > 1) sheet.getRange(2, 1, last - 1, sheet.getLastColumn()).clearContent();
  }
  return ContentService.createTextOutput(JSON.stringify(readDemandes_())).setMimeType(ContentService.MimeType.JSON);
}
