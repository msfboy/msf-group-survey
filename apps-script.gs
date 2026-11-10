/**
 * Collecte centralisée dans Google Sheets + e-mail de remerciement automatique.
 * 1. Créez un Google Sheet > Extensions > Apps Script, collez ce code.
 * 2. Renseignez GUIDE_URL et ADMIN_EMAIL ci-dessous.
 * 3. Déployer > Nouveau déploiement > Application Web
 *    (Exécuter en tant que : moi / Accès : tout le monde). Autorisez l'envoi d'e-mails.
 * 4. Copiez l'URL /exec dans la constante ENDPOINT de index.html.
 *
 * Limite : ~100 e-mails/jour avec un Gmail gratuit, ~1500 avec Google Workspace.
 */
const GUIDE_URL   = "https://drive.google.com/file/d/1gDF2tKOUl9DMkYBRmD8ant2CDgmfMAPE/view?usp=sharing"; // lien de téléchargement du guide (Google Drive)
const ADMIN_EMAIL = ""; // <-- optionnel : votre e-mail pour être alerté à chaque réponse
const SENDER_NAME = "MSF GROUP";

function doPost(e) {
  // Verrou : évite les conflits si plusieurs personnes envoient en même temps.
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  let d;
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    d = JSON.parse(e.postData.contents);
    let headers = sheet.getLastColumn()
      ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0] : [];
    Object.keys(d).forEach(k => { if (headers.indexOf(k) < 0) headers.push(k); });
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    const row = headers.map(h => Array.isArray(d[h]) ? d[h].join(' | ') : (d[h] || ''));
    sheet.appendRow(row);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  // E-mails : une erreur d'envoi ne doit jamais faire perdre la réponse déjà enregistrée.
  try { if (d.email && /^\S+@\S+\.\S+$/.test(d.email)) sendThanks(d); } catch (err) { console.error(err); }
  try { if (ADMIN_EMAIL) sendAlert(d); } catch (err) { console.error(err); }
  return ContentService.createTextOutput('ok');
}

function esc(s) {
  return String(s || '').replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
}

function sendThanks(d) {
  const prenom = esc((d.nom || '').split(' ')[0]);
  const html = `
  <div style="font-family:Segoe UI,Arial,sans-serif;background:#eef1f6;padding:24px">
   <div style="max-width:560px;margin:auto;background:#fff;border-radius:10px;overflow:hidden;border:1px solid #dadce0">
    <div style="background:#1a3a6b;color:#fff;padding:22px 26px">
      <div style="font-size:12px;letter-spacing:1px;color:#f0c75e">MSF GROUP · ACADÉMIE LANGUES & IA APPLIQUÉE</div>
      <div style="font-size:22px;font-weight:600;margin-top:6px">Merci pour votre participation !</div>
    </div>
    <div style="padding:26px;color:#202124;font-size:15px;line-height:1.6">
      <p>Bonjour ${prenom},</p>
      <p>Merci d'avoir pris le temps de répondre à notre enquête. Vos réponses nous aideront à concevoir des formations et des outils parfaitement adaptés à vos besoins.</p>
      <p>Comme promis, voici votre cadeau :</p>
      <div style="background:#fff8e6;border-left:4px solid #c8962e;padding:14px 16px;border-radius:4px">
        🎁 <b>Guide Pratique MSF des expressions clés pour booster votre parcours</b>
      </div>
      <p style="text-align:center;margin:26px 0">
        <a href="${GUIDE_URL}" style="background:#1a3a6b;color:#fff;text-decoration:none;font-weight:600;padding:13px 30px;border-radius:6px;display:inline-block">Télécharger mon guide</a>
      </p>
      <p style="font-size:13px;color:#5f6368">Si le bouton ne fonctionne pas, copiez ce lien dans votre navigateur :<br><a href="${GUIDE_URL}">${GUIDE_URL}</a></p>
      <p>Merci pour votre précieuse collaboration !<br><i>— L'équipe du GROUPE MSF</i></p>
    </div>
   </div></div>`;
  MailApp.sendEmail({
    to: d.email,
    subject: "🎁 Merci ! Voici votre Guide Pratique MSF",
    htmlBody: html,
    body: "Bonjour,\n\nMerci pour votre participation. Téléchargez votre guide ici : " + GUIDE_URL + "\n\n— L'équipe du GROUPE MSF",
    name: SENDER_NAME
  });
}

function sendAlert(d) {
  MailApp.sendEmail({
    to: ADMIN_EMAIL,
    subject: "Nouvelle réponse à l'enquête : " + (d.nom || ''),
    body: "Nom : " + (d.nom || '') + "\nE-mail : " + (d.email || '') + "\nWhatsApp : " + (d.tel || '') +
          "\nLieu : " + (d.lieu || '') + "\nStatut : " + (d.statut || '') + "\n\nToutes les réponses sont dans votre Google Sheet."
  });
}
