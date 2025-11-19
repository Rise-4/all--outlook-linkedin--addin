/* global Office, OfficeRuntime, fetch */
Office.initialize = () => {};

const GRAPH = "https://graph.microsoft.com/v1.0";

function open(url) {
  Office.context.ui.openBrowserWindow(url);
}

function encode(str) {
  return encodeURIComponent(str || "").replace(/%20/g, "+");
}

function normalizeName(details) {
  if (details && details.displayName) return details.displayName;
  if (details && details.emailAddress) {
    const local = details.emailAddress.split("@")[0].replace(/[._\-]+/g, " ");
    return local;
  }
  return "";
}

function looksLikeLinkedIn(s) {
  if (!s) return false;
  const t = String(s).toLowerCase();
  return t.includes("linkedin.com") || t.startsWith("/in/") || /^[a-z0-9\-_/]+$/.test(t);
}

function toLinkedInUrl(s) {
  const t = String(s).trim();
  if (t.startsWith("http://") || t.startsWith("https://")) return t;
  if (t.includes("/in/")) return "https://www.linkedin.com" + (t.startsWith("/") ? "" : "/") + t.replace(/^\/?/, "");
  return "https://www.linkedin.com/in/" + t.replace(/^\/|\/$/g, "") + "/";
}

async function getToken() {
  // SSO token (si l’IT a configuré l’App + scopes). Sinon, une invite d’auth pourra s’afficher.
  return OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true });
}

async function graphFetch(path, token) {
  const res = await fetch(GRAPH + path, { headers: { Authorization: "Bearer " + token } });
  if (!res.ok) throw new Error("Graph error " + res.status);
  return res.json();
}

async function tryFindContactByEmail(email, token) {
  const filter = encodeURIComponent(`emailAddresses/any(a:a/address eq '${email}')`);
  const select = encodeURIComponent("displayName,givenName,surname,imAddresses");
  const data = await graphFetch(`/me/contacts?$filter=${filter}&$select=${select}&$top=1`, token);
  if (data.value && data.value.length) return data.value[0];
  return null;
}

async function openLinkedIn(event) {
  try {
    const item = Office.context.mailbox.item;
    if (!item) return event.completed();

    // Cible expéditeur (réception) ou premier destinataire (envoyé)
    let target = item.from || (item.to && item.to.length ? item.to[0] : null);
    const email = target && target.emailAddress;
    const displayName = normalizeName(target);

    // 1) Si Graph dispo → tente IM LinkedIn
    try {
      if (email) {
        const token = await getToken();
        const contact = await tryFindContactByEmail(email, token);
        if (contact && contact.imAddresses && contact.imAddresses.length) {
          const im = contact.imAddresses.find(looksLikeLinkedIn);
          if (im) {
            open(toLinkedInUrl(im));
            return;
          }
        }
        // Pas d'IM LinkedIn → tente nom "propre"
        const name = contact?.givenName || contact?.surname ? `${contact?.givenName || ""} ${contact?.surname || ""}`.trim() : displayName;
        open("https://www.linkedin.com/search/results/people/?keywords=" + encode(name));
        return;
      }
    } catch (e) {
      // Pas de consentement Graph / erreur → on retombe en mode léger
    }

    // 2) Mode LÉGER : recherche People basée sur l’item
    const name = displayName || (item.subject || "");
    open("https://www.linkedin.com/search/results/people/?keywords=" + encode(name));
  } finally {
    event.completed();
  }
}

// Expose
window.openLinkedIn = openLinkedIn;