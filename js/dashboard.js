const CLIENT_ID = CONFIG.CLIENT_ID;
const TENANT_ID = CONFIG.TENANT_ID;
const GROUPS    = CONFIG.GROUPS;

const msalConfig = {
  auth: {
    clientId:    CLIENT_ID,
    authority:   `https://login.microsoftonline.com/${TENANT_ID}`,
    redirectUri: window.location.origin
  },
  cache: { cacheLocation: "sessionStorage" }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);

msalInstance.initialize().then(() => {

  // Check if user is logged in
  const accounts = msalInstance.getAllAccounts();

  // If not logged in send them back to login page
  if (accounts.length === 0) {
    window.location.href = "index.html";
    return;
  }

  // User is logged in — load their info
  const account = accounts[0];
  const claims  = account.idTokenClaims;

  populateDashboard(account, claims);

}).catch(err => console.error("MSAL init error:", err));

function getRole(claims) {
  const userGroups = claims?.groups || [];

  if (userGroups.includes(GROUPS.SALES_ADMIN))    return { label: "Sales Admin",     badge: "badge-sales-admin"     };
  if (userGroups.includes(GROUPS.SALES_ASSOCIATE)) return { label: "Sales Associate", badge: "badge-sales-associate" };
  if (userGroups.includes(GROUPS.IT_ADMIN))        return { label: "IT Admin",        badge: "badge-it-admin"        };
  if (userGroups.includes(GROUPS.IT_HELPDESK))     return { label: "IT Help Desk",    badge: "badge-it-helpdesk"     };

  return { label: "No role assigned", badge: "badge-default" };
}

function populateDashboard(account, claims) {
  const name    = account.name || "";
  const parts   = name.split(" ");
  const first   = parts[0] || "";
  const last    = parts[1] || "";
  const email   = account.username || "";
  const initials = (first[0] || "") + (last[0] || "");
  const role    = getRole(claims);

  // Header
  document.getElementById("dash-initials").textContent = initials.toUpperCase();
  document.getElementById("dash-name").textContent     = name;
  document.getElementById("dash-email").textContent    = email;

  // Role badge
  const roleEl = document.getElementById("dash-role");
  roleEl.textContent = role.label;
  roleEl.className   = `badge ${role.badge}`;

  // Info grid
  document.getElementById("info-first").textContent = first;
  document.getElementById("info-last").textContent  = last;
  document.getElementById("info-email").textContent = email;
  document.getElementById("info-role").textContent  = role.label;
}

function signOut() {
  msalInstance.logoutRedirect({
    postLogoutRedirectUri: window.location.origin + "/index.html"
  });
}