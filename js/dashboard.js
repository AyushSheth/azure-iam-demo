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
  renderRoleContent(role.label);
}

function signOut() {
  msalInstance.logoutRedirect({
    postLogoutRedirectUri: window.location.origin + "/index.html"
  });
}
function renderRoleContent(roleLabel) {
  const container = document.getElementById("role-content");

  if (roleLabel === "Sales Admin") {
    container.innerHTML = `
      <div class="role-section">
        <h2>Sales Overview</h2>
        <div class="role-card">
          <span class="role-card-label">Team Size</span>
          <span class="role-card-value">12 reps</span>
        </div>
        <div class="role-card">
          <span class="role-card-label">Quarterly Revenue</span>
          <span class="role-card-value">$482,000</span>
        </div>
        <div class="role-card">
          <span class="role-card-label">Open Deals</span>
          <span class="role-card-value">37</span>
        </div>
        <div class="role-link-list">
          <a href="#" class="role-link">Manage Sales Team</a>
          <a href="#" class="role-link">View All Reports</a>
        </div>
      </div>
    `;
  }

  else if (roleLabel === "Sales Associate") {
    container.innerHTML = `
      <div class="role-section">
        <h2>My Sales</h2>
        <div class="role-card">
          <span class="role-card-label">My Quota</span>
          <span class="role-card-value">$40,000</span>
        </div>
        <div class="role-card">
          <span class="role-card-label">My Open Deals</span>
          <span class="role-card-value">5</span>
        </div>
        <div class="role-link-list">
          <a href="#" class="role-link">View My Pipeline</a>
        </div>
      </div>
    `;
  }

  else if (roleLabel === "IT Admin") {
    container.innerHTML = `
      <div class="role-section">
        <h2>System Administration</h2>
        <div class="role-card">
          <span class="role-card-label">Active Users</span>
          <span class="role-card-value">247</span>
        </div>
        <div class="role-card">
          <span class="role-card-label">Open Tickets</span>
          <span class="role-card-value">9</span>
        </div>
        <div class="role-card">
          <span class="role-card-label">System Status</span>
          <span class="role-card-value">All Systems Operational</span>
        </div>
        <div class="role-link-list">
          <a href="#" class="role-link">Manage Users</a>
          <a href="#" class="role-link">View System Logs</a>
        </div>
      </div>
    `;
  }

  else if (roleLabel === "IT Help Desk") {
    container.innerHTML = `
      <div class="role-section">
        <h2>Help Desk Queue</h2>
        <div class="role-card">
          <span class="role-card-label">My Open Tickets</span>
          <span class="role-card-value">4</span>
        </div>
        <div class="role-card">
          <span class="role-card-label">Avg Response Time</span>
          <span class="role-card-value">12 min</span>
        </div>
        <div class="role-link-list">
          <a href="#" class="role-link">View Ticket Queue</a>
        </div>
      </div>
    `;
  }

  else {
    container.innerHTML = `
      <div class="no-role-message">
        Your account has not been assigned a role yet. Please contact your administrator to request access.
      </div>
    `;
  }
}