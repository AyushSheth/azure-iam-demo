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

const loginRequest = {
  scopes: ["openid", "profile", "email", "User.Read"]
};

// Initialize MSAL first, then run everything else
const msalInstance = new msal.PublicClientApplication(msalConfig);

msalInstance.initialize().then(() => {

  // Handle redirect back from Microsoft
  msalInstance.handleRedirectPromise().then(response => {
    if (response) {
      showDashboard(response.account, response.idTokenClaims);
    } else {
      const accounts = msalInstance.getAllAccounts();
      if (accounts.length > 0) {
        showDashboard(accounts[0], accounts[0].idTokenClaims);
      }
    }
  }).catch(err => console.error("Auth error:", err));

}).catch(err => console.error("MSAL init error:", err));

function signIn() {
  msalInstance.loginRedirect(loginRequest);
}

function signOut() {
  msalInstance.logoutRedirect();
}

function getRole(claims) {
  const userGroups = claims?.groups || [];

  if (userGroups.includes(GROUPS.SALES_ADMIN))     return { label: "Sales Admin",     badge: "badge-sales-admin"     };
  if (userGroups.includes(GROUPS.SALES_ASSOCIATE))  return { label: "Sales Associate", badge: "badge-sales-associate" };
  if (userGroups.includes(GROUPS.IT_ADMIN))         return { label: "IT Admin",        badge: "badge-it-admin"        };
  if (userGroups.includes(GROUPS.IT_HELPDESK))      return { label: "IT Help Desk",    badge: "badge-it-helpdesk"     };

  return { label: "No role assigned", badge: "badge-default" };
}

function showDashboard(account, claims) {
  // Redirect to the secure dashboard page
  window.location.href = "dashboard.html";
}
