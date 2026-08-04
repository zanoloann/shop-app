/**
 * Envoi des e-mails de notification — appelé directement depuis
 * apps-script-planning.gs (notifyNewDemande_ / notifyDemandeurStatus_) et
 * depuis le dispatcheur doGet/doPost de apps-script-classeur-unique.gs pour
 * le test manuel (data.test). Même projet, même classeur, même déploiement.
 *
 * Trois e-mails automatiques :
 * 1) À la création d'une demande : l'administrateur reçoit un mail avec les
 *    liens Valider/Refuser ; le demandeur reçoit une confirmation de réception.
 * 2) À la décision de l'administrateur (via l'appli ou les liens de l'e-mail) :
 *    le demandeur reçoit le résultat (validée/refusée).
 *
 * Les adresses sont retrouvées dans l'onglet "Artisans" (getAdminEmail_ /
 * getArtisanEmail_ dans apps-script-classeur-unique.gs) — aucune adresse
 * e-mail écrite en dur dans le code.
 */

function demandeTypeLabel_(type) {
  const labels = {
    astreinte: 'une Astreinte',
    renfort1: 'un Renfort 1',
    renfort2: 'un Renfort 2',
    liberation_astreinte: 'la suppression de son Astreinte',
    liberation_renfort1: 'la suppression de son Renfort 1',
    liberation_renfort2: 'la suppression de son Renfort 2'
  };
  return labels[type] || type;
}

// À exécuter UNE FOIS manuellement depuis l'éditeur (sélectionnez cette
// fonction dans la barre d'exécution puis ▶) pour accorder le scope d'envoi
// de mail — nécessaire une seule fois après le premier déploiement.
function autoriserEnvoiMail() {
  const email = getAdminEmail_() || Session.getActiveUser().getEmail();
  MailApp.sendEmail({ to: email, subject: 'Autorisation — Demandes de planning', body: 'Autorisation accordée avec succès.' });
}

// Appelée depuis apps-script-planning.gs à la création d'une demande.
function sendNewDemandeEmail_(data) {
  const adminEmail = getAdminEmail_();
  const summary = data.daySummary || '';
  const comment = data.commentaire ? '\nCommentaire : ' + data.commentaire + '\n' : '';

  if (adminEmail) {
    const subject = 'Nouvelle demande de planning — ' + data.demandeur;
    const textBody = data.demandeur + ' demande ' + demandeTypeLabel_(data.type) + ' le ' + (data.dateConcernee || '?') + '.\n' +
      comment + '\nSituation actuelle de cette journée :\n' + summary + '\n\n' +
      'Valider : ' + (data.validateUrl || '') + '\nRefuser : ' + (data.refuseUrl || '') + '\n';
    const htmlBody =
      '<div style="font-family:sans-serif;color:#2B241C;max-width:480px;">' +
      '<h2 style="font-weight:600;">Nouvelle demande de planning</h2>' +
      '<p><b>' + data.demandeur + '</b> demande ' + demandeTypeLabel_(data.type) + ' le <b>' + (data.dateConcernee || '?') + '</b>.</p>' +
      (data.commentaire ? '<p style="color:#6B5E48;font-style:italic;">« ' + data.commentaire + ' »</p>' : '') +
      '<div style="background:#FFFCF6;border:1px solid #E7DAC4;border-radius:10px;padding:12px 16px;margin:16px 0;white-space:pre-line;font-size:13px;">' +
      '<b>Situation actuelle de cette journée</b><br>' + summary + '</div>' +
      '<div style="margin-top:20px;">' +
      '<a href="' + (data.validateUrl || '#') + '" style="display:inline-block;padding:10px 22px;background:#3F6B45;color:#fff;text-decoration:none;border-radius:8px;font-weight:600;margin-right:10px;">Valider</a>' +
      '<a href="' + (data.refuseUrl || '#') + '" style="display:inline-block;padding:10px 22px;background:#B5502E;color:#fff;text-decoration:none;border-radius:8px;font-weight:600;">Refuser</a>' +
      '</div>' +
      '<p style="font-size:11px;color:#9C8F79;margin-top:16px;">Vous pouvez aussi traiter cette demande depuis l\'appli "Demandes de planning".</p></div>';
    try { MailApp.sendEmail({ to: adminEmail, subject: subject, body: textBody, htmlBody: htmlBody }); } catch (err) {}
  }

  const demandeurEmail = getArtisanEmail_(data.demandeur);
  if (demandeurEmail) {
    try {
      MailApp.sendEmail({
        to: demandeurEmail,
        subject: 'Votre demande de planning a bien été reçue',
        body: 'Bonjour ' + data.demandeur + ',\n\nVotre demande (' + demandeTypeLabel_(data.type) + ') pour le ' + (data.dateConcernee || '?') + ' a bien été reçue et est en attente de validation par l\'administrateur.' +
          comment + '\nVous recevrez un e-mail dès qu\'elle sera traitée.'
      });
    } catch (err) {}
  }
}

// Appelée depuis apps-script-planning.gs quand un artisan s'enregistre
// directement (jour totalement vide) — validation immédiate, sans admin.
function sendAutoRegisteredEmail_(demandeur, type, dateConcernee) {
  const email = getArtisanEmail_(demandeur);
  if (!email) return;
  try {
    MailApp.sendEmail({
      to: email,
      subject: 'Vous êtes enregistré(e) pour le ' + (dateConcernee || '?'),
      body: 'Bonjour ' + demandeur + ',\n\nVous avez été enregistré(e) automatiquement pour ' + demandeTypeLabel_(type) + ' le ' + (dateConcernee || '?') + ' (personne d\'autre n\'était présent(e) ce jour-là, validation immédiate sans passer par l\'administrateur).\n\nMerci !'
    });
  } catch (err) {}
}

// Appelée depuis apps-script-planning.gs (applyDemandeStatus_) une fois la décision prise.
function notifyDemandeurStatus_(demandeur, type, dateConcernee, statut) {
  const email = getArtisanEmail_(demandeur);
  if (!email) return;
  const verb = statut === 'Validée' ? 'validée' : 'refusée';
  try {
    MailApp.sendEmail({
      to: email,
      subject: 'Votre demande de planning a été ' + verb,
      body: 'Bonjour ' + demandeur + ',\n\nVotre demande (' + demandeTypeLabel_(type) + ') pour le ' + (dateConcernee || '?') + ' a été ' + verb + '.'
    });
  } catch (err) {}
}

// Test manuel depuis l'appli (Réglages > Envoyer un mail de test).
function doPostEmail_(data) {
  const adminEmail = getAdminEmail_();
  if (!adminEmail) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: "Aucun artisan nommé 'Administrateur' trouvé dans l'onglet Artisans" })).setMimeType(ContentService.MimeType.JSON);
  }
  try {
    MailApp.sendEmail({ to: adminEmail, subject: 'Test — Demandes de planning', body: 'Ceci est un e-mail de test envoyé depuis apps-script-demandes-email.gs.' });
    return ContentService.createTextOutput(JSON.stringify({ ok: true, sentTo: adminEmail })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) })).setMimeType(ContentService.MimeType.JSON);
  }
}
