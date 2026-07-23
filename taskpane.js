/* =====================================================================
   CHRONO COURRIERS — Complément Word
   Logique : création d'une ligne dans la liste SharePoint (via Microsoft
   Graph), attente du numéro attribué par le flux Power Automate, puis
   insertion de la référence dans le document.
   =====================================================================

   ------------------- CONFIGURATION (à adapter) ----------------------- */
const CONFIG = {
  // Inscription d'application Entra ID (voir guide, étape 1)
  clientId: "60187235-dc22-4982-8a7b-f441a9e32675",
  tenantId: "3da5f599-13bb-44dc-a761-53f3a0dbfcda",

  // Emplacement du registre (phase 1)
  siteHostname: "lgdj.sharepoint.com",   // sans https://
  sitePath: "/sites/Direction",                  // chemin du site ("" si site racine)
  listName: "Chrono Courriers",                  // nom exact de la liste

  // Émetteurs proposés (doivent exister dans la colonne Choix "Emetteur")
  emetteurs: ["DG", "DSSI", "DSI"],

  // Attente du flux : intervalle et durée maximale (millisecondes)
  pollIntervalMs: 1500,
  pollTimeoutMs: 45000,
};
/* ------------------- fin de la configuration ------------------------ */

const GRAPH = "https://graph.microsoft.com/v1.0";
const SCOPES = ["https://graph.microsoft.com/Sites.ReadWrite.All"];

let msalApp = null;
let siteId = null;
let listId = null;

const $ = (id) => document.getElementById(id);
const statut = (msg, cls = "") => { const s = $("statut"); s.textContent = msg; s.className = cls; };

/* ---------------------- Démarrage ----------------------------------- */
Office.onReady(async () => {
  // Émetteurs
  const sel = $("emetteur");
  CONFIG.emetteurs.forEach((e) => {
    const o = document.createElement("option");
    o.value = e; o.textContent = e;
    sel.appendChild(o);
  });

  // Lien vers le registre
  $("registre").addEventListener("click", () => {
    const url = `https://${CONFIG.siteHostname}${CONFIG.sitePath}/Lists/${encodeURIComponent(CONFIG.listName)}`;
    window.open(url, "_blank");
  });

  $("generer").addEventListener("click", genererReference);

  // MSAL
  msalApp = new msal.PublicClientApplication({
    auth: {
      clientId: CONFIG.clientId,
      authority: `https://login.microsoftonline.com/${CONFIG.tenantId}`,
      redirectUri: window.location.origin + window.location.pathname,
    },
    cache: { cacheLocation: "localStorage" },
  });
  await msalApp.initialize();

  const comptes = msalApp.getAllAccounts();
  if (comptes.length > 0) $("compte").textContent = comptes[0].username;
});

/* ---------------------- Authentification ---------------------------- */
async function obtenirJeton() {
  const comptes = msalApp.getAllAccounts();
  if (comptes.length > 0) {
    try {
      const r = await msalApp.acquireTokenSilent({ scopes: SCOPES, account: comptes[0] });
      $("compte").textContent = comptes[0].username;
      return r.accessToken;
    } catch (e) { /* jeton expiré : on retombe sur la fenêtre de connexion */ }
  }
  statut("Connexion à Microsoft 365…");
  const r = await msalApp.loginPopup({ scopes: SCOPES });
  $("compte").textContent = r.account.username;
  return r.accessToken;
}

/* ---------------------- Appels Graph -------------------------------- */
async function graph(jeton, methode, chemin, corps) {
  const rep = await fetch(GRAPH + chemin, {
    method: methode,
    headers: {
      Authorization: "Bearer " + jeton,
      "Content-Type": "application/json",
    },
    body: corps ? JSON.stringify(corps) : undefined,
  });
  if (!rep.ok) {
    let detail = "";
    try { detail = (await rep.json())?.error?.message || ""; } catch (e) {}
    throw new Error(`Graph ${rep.status} — ${detail || rep.statusText}`);
  }
  return rep.status === 204 ? null : rep.json();
}

async function resoudreListe(jeton) {
  if (siteId && listId) return;
  const cheminSite = CONFIG.sitePath
    ? `/sites/${CONFIG.siteHostname}:${CONFIG.sitePath}`
    : `/sites/${CONFIG.siteHostname}`;
  const site = await graph(jeton, "GET", cheminSite);
  siteId = site.id;

  const listes = await graph(jeton, "GET", `/sites/${siteId}/lists?$select=id,displayName&$top=200`);
  const liste = (listes.value || []).find(
    (l) => l.displayName.toLowerCase() === CONFIG.listName.toLowerCase()
  );
  if (!liste) throw new Error(`Liste « ${CONFIG.listName} » introuvable sur ${CONFIG.sitePath || "le site racine"}. Vérifiez la phase 1.`);
  listId = liste.id;
}

/* ---------------------- Scénario principal --------------------------- */
async function genererReference() {
  const btn = $("generer");
  const emetteur = $("emetteur").value;
  const objet = $("objet").value.trim();
  $("tampon").style.display = "none";

  btn.disabled = true;
  try {
    const jeton = await obtenirJeton();

    statut("Vérification du registre…");
    await resoudreListe(jeton);

    statut("Demande du prochain numéro…");
    const item = await graph(jeton, "POST", `/sites/${siteId}/lists/${listId}/items`, {
      fields: { Title: objet || "(sans objet)", Emetteur: emetteur },
    });

    statut("Attribution en cours (file d'attente centrale)…");
    const reference = await attendreReference(jeton, item.id);

    await insererDansWord(reference);
    afficherTampon(reference);
    statut("Référence insérée dans le courrier.", "ok");
    $("objet").value = "";
  } catch (e) {
    statut(messageErreur(e), "alerte");
  } finally {
    btn.disabled = false;
  }
}

async function attendreReference(jeton, itemId) {
  const debut = Date.now();
  while (Date.now() - debut < CONFIG.pollTimeoutMs) {
    const item = await graph(
      jeton, "GET",
      `/sites/${siteId}/lists/${listId}/items/${itemId}?$expand=fields($select=Reference,Numero)`
    );
    const ref = item?.fields?.Reference;
    if (ref) return ref;
    await new Promise((r) => setTimeout(r, CONFIG.pollIntervalMs));
  }
  throw new Error(
    "Le flux n'a pas répondu à temps. La demande est enregistrée au registre : " +
    "le numéro y apparaîtra dès le passage du flux (vérifiez qu'il est activé)."
  );
}

/* ---------------------- Insertion dans Word -------------------------- */
async function insererDansWord(reference) {
  await Word.run(async (ctx) => {
    // 1) Contrôle de contenu balisé "REF" (modèles équipés)
    const ccs = ctx.document.contentControls.getByTag("REF");
    ccs.load("items");
    await ctx.sync();

    if (ccs.items.length > 0) {
      ccs.items[0].insertText(reference, Word.InsertLocation.replace);
    } else {
      // 2) Sinon : au point d'insertion (remplace la sélection éventuelle)
      ctx.document.getSelection().insertText(reference, Word.InsertLocation.replace);
    }
    await ctx.sync();
  });
}

/* ---------------------- Interface ------------------------------------ */
function afficherTampon(reference) {
  $("tamponRef").textContent = reference;
  $("tamponDate").textContent = new Date().toLocaleDateString("fr-FR", {
    day: "2-digit", month: "long", year: "numeric",
  });
  const t = $("tampon");
  t.style.display = "block";
  t.classList.remove("pose");
  void t.offsetWidth; // relance l'animation
  t.classList.add("pose");
}

function messageErreur(e) {
  const m = String(e && e.message ? e.message : e);
  if (m.includes("403") || m.includes("401"))
    return "Accès refusé : vérifiez le consentement administrateur de l'application (guide, étape 1) et vos droits sur la liste.";
  if (m.includes("user_cancelled") || m.includes("popup_window_error"))
    return "Connexion annulée ou fenêtre bloquée : autorisez les fenêtres contextuelles puis réessayez.";
  if (m.includes("400") && m.includes("Emetteur"))
    return "La colonne « Emetteur » est absente ou ne contient pas cette valeur : vérifiez la liste (phase 1).";
  return m;
}
