<script>
// ============================================================
//  DIGIFYCE DASHBOARD — FRONTEND JAVASCRIPT
// ============================================================

var currentUser = null;
var clientList  = [];
var currentPage = 'dashboard';

window.onload = function() {
  var saved = sessionStorage.getItem('digifyce_user');
  if (saved) { currentUser = JSON.parse(saved); initApp(); }
};

// ── AUTH ─────────────────────────────────────────────────────
function doLogin() {
  var username = document.getElementById('loginUsername').value.trim();
  var password = document.getElementById('loginPassword').value;
  var btn      = document.getElementById('loginBtn');
  var errEl    = document.getElementById('loginError');
  if (!username || !password) { showLoginError('Please enter username and password'); return; }
  btn.disabled  = true;
  btn.innerHTML = '<span class="spinner"></span> Signing in...';
  errEl.style.display = 'none';
  google.script.run
    .withSuccessHandler(function(res) {
      btn.disabled  = false;
      btn.innerHTML = '<span>Sign In</span>';
      if (res.success) {
        currentUser = res.user;
        sessionStorage.setItem('digifyce_user', JSON.stringify(currentUser));
        initApp();
      } else { showLoginError(res.error || 'Login failed'); }
    })
    .withFailureHandler(function(err) {
      btn.disabled  = false;
      btn.innerHTML = '<span>Sign In</span>';
      showLoginError('Server error: ' + err.message);
    })
    .serverLogin(username, password);
}

function showLoginError(msg) {
  var el = document.getElementById('loginError');
  el.textContent = msg; el.style.display = 'block';
}

function doLogout() {
  sessionStorage.removeItem('digifyce_user');
  currentUser = null;
  document.getElementById('mainApp').style.display    = 'none';
  document.getElementById('loginScreen').style.display = 'flex';
  document.getElementById('loginUsername').value = '';
  document.getElementById('loginPassword').value = '';
}

// ── APP INIT ─────────────────────────────────────────────────
function initApp() {
  document.getElementById('loginScreen').style.display = 'none';
  document.getElementById('mainApp').style.display     = 'flex';
  document.getElementById('userName').textContent   = currentUser.name;
  document.getElementById('userRole').textContent   = currentUser.role;
  document.getElementById('userAvatar').textContent = currentUser.name.charAt(0).toUpperCase();
  if (currentUser.role === 'owner') {
    document.querySelectorAll('.owner-only').forEach(function(el) { el.style.display = 'flex'; });
  }
  var now    = new Date();
  var months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  var rMon   = now.getMonth() === 0 ? 11 : now.getMonth() - 1;
  var rYear  = now.getMonth() === 0 ? now.getFullYear() - 1 : now.getFullYear();
  document.getElementById('currentMonthChip').textContent = months[rMon] + ' ' + rYear + ' (report month)';
  google.script.run.initSystemSheets();
  loadDashboard();
  if (currentUser.role === 'owner') loadApprovalsBadge();
}

// ── NAVIGATION ───────────────────────────────────────────────
function showPage(page) {
  document.querySelectorAll('.page').forEach(function(p) { p.classList.remove('active'); });
  document.querySelectorAll('.nav-item').forEach(function(n) { n.classList.remove('active'); });
  var el  = document.getElementById('page-' + page);
  var nav = document.getElementById('nav-' + page);
  if (el)  el.classList.add('active');
  if (nav) nav.classList.add('active');
  currentPage = page;
  var titles = { dashboard:'Dashboard', clients:'Clients', clientform:'Client Form',
    generate:'Generate Report', logs:'Report Logs', users:'Users', approvals:'Pending Approvals' };
  document.getElementById('topbarTitle').textContent = titles[page] || page;
  document.getElementById('sidebar').classList.remove('open');
  if (page === 'dashboard') loadDashboard();
  if (page === 'clients')   loadClientsTable();
  if (page === 'generate')  loadGenerateDropdown();
  if (page === 'logs')      loadLogs();
  if (page === 'users')     loadUsersTable();
  if (page === 'approvals') loadApprovals();
}

function toggleSidebar() { document.getElementById('sidebar').classList.toggle('open'); }

// ── DASHBOARD ────────────────────────────────────────────────
function loadDashboard() {
  google.script.run
    .withSuccessHandler(function(clients) {
      if (!Array.isArray(clients)) return;
      clientList = clients;
      document.getElementById('statClients').textContent = clients.length;
      renderClientCards(clients);
    })
    .serverGetClients(currentUser.role, currentUser.clients);
  google.script.run
    .withSuccessHandler(function(logs) {
      if (!Array.isArray(logs)) return;
      document.getElementById('statReports').textContent = logs.length;
      document.getElementById('statSuccess').textContent = logs.filter(function(l){ return l.status==='success'; }).length;
      document.getElementById('statFailed').textContent  = logs.filter(function(l){ return l.status==='failed';  }).length;
    })
    .serverGetReportLogs(currentUser.role, currentUser.clients, 200);
}

function renderClientCards(clients) {
  var el = document.getElementById('clientCards');
  if (!clients.length) {
    el.innerHTML = '<div class="empty-state"><div class="empty-state-icon">&#9672;</div>'
      + '<div class="empty-state-text">No clients yet. <a href="#" onclick="showPage(\'clients\')">Add your first client</a></div></div>';
    return;
  }
  var html = '';
  clients.forEach(function(c) {
    var hasSheet = c.sheet_id && c.sheet_id.trim() !== '';
    var hasGAds  = c.google_ads_enabled === 'true' || c.google_ads_enabled === true;
    var sheetTag = hasSheet
      ? '<span class="client-tag green">Sheet OK</span>'
      : '<span class="client-tag" style="background:#FFEBEE;color:#C62828">No Sheet</span>';
    var sheetBtn = hasSheet
      ? '<button class="btn btn-ghost sm" onclick="openClientSheet(\'' + esc(c.client_key) + '\')">Open Sheet</button>'
      : '<button class="btn btn-ghost sm" style="border-color:#2E7D32;color:#2E7D32" onclick="createClientSheet(\'' + esc(c.client_key) + '\',\'' + esc(c.name) + '\')">+ Create Sheet</button>';
    html += '<div class="client-card">'
      + '<div class="client-card-name">' + esc(c.name) + '</div>'
      + '<div class="client-card-key">' + esc(c.client_key) + '</div>'
      + '<div class="client-card-meta"><span class="client-tag">Meta</span><span class="client-tag">GA4</span><span class="client-tag">GSC</span>'
      + (hasGAds ? '<span class="client-tag green">Google Ads</span>' : '') + sheetTag + '</div>'
      + '<div class="client-card-actions" style="flex-wrap:wrap;gap:6px">'
      + '<button class="btn btn-primary sm" onclick="quickGenerate(\'' + esc(c.client_key) + '\',\'' + esc(c.name) + '\')">Generate</button>'
      + sheetBtn
      + '<button class="btn btn-ghost sm" onclick="editClient(\'' + esc(c.client_key) + '\')">Edit</button>'
      + '</div></div>';
  });
  el.innerHTML = html;
}

function quickGenerate(clientKey, clientName) {
  showPage('generate');
  setTimeout(function() { var s = document.getElementById('gen_client'); if (s) s.value = clientKey; }, 150);
}

// ── CLIENTS TABLE ────────────────────────────────────────────
function loadClientsTable() {
  var el = document.getElementById('clientsTable');
  el.innerHTML = '<div class="loading-state">Loading...</div>';
  google.script.run
    .withSuccessHandler(function(clients) {
      if (!Array.isArray(clients)) { el.innerHTML = '<div class="loading-state">Error loading clients</div>'; return; }
      clientList = clients;
      if (!clients.length) { el.innerHTML = '<div class="empty-state"><div class="empty-state-icon">&#9672;</div><div class="empty-state-text">No clients yet.</div></div>'; return; }
      var rows = '';
      clients.forEach(function(c) {
        var hasSheet = c.sheet_id && c.sheet_id.trim() !== '';
        var hasGAds  = c.google_ads_enabled === 'true' || c.google_ads_enabled === true;
        var sheetBtn = hasSheet
          ? '<button class="btn-icon" onclick="openClientSheet(\'' + esc(c.client_key) + '\')">Open Sheet</button>'
          : '<button class="btn-icon" style="color:var(--green);border-color:var(--green)" onclick="createClientSheet(\'' + esc(c.client_key) + '\',\'' + esc(c.name) + '\')">+ Create Sheet</button>';
        var delBtn = currentUser.role === 'owner'
          ? '<button class="btn-icon" style="color:var(--red)" onclick="deleteClient(\'' + esc(c.client_key) + '\',\'' + esc(c.name) + '\')">X</button>' : '';
        rows += '<tr>'
          + '<td class="code-cell">' + esc(c.client_key) + '</td>'
          + '<td><strong>' + esc(c.name) + '</strong></td>'
          + '<td>' + esc(c.agent_email) + '</td>'
          + '<td class="code-cell">' + esc(c.ga4_property_id) + '</td>'
          + '<td><span class="status-badge ' + (hasSheet ? 'success' : 'failed') + '">' + (hasSheet ? 'Linked' : 'No Sheet') + '</span></td>'
          + '<td><span class="status-badge ' + (hasGAds ? 'success' : 'failed') + '">' + (hasGAds ? 'ON' : 'OFF') + '</span></td>'
          + '<td><div style="display:flex;gap:6px;flex-wrap:wrap">' + sheetBtn
          + '<button class="btn-icon" onclick="editClient(\'' + esc(c.client_key) + '\')">Edit</button>'
          + '<button class="btn-icon" onclick="quickGenerate(\'' + esc(c.client_key) + '\',\'' + esc(c.name) + '\')">Run</button>'
          + delBtn + '</div></td></tr>';
      });
      el.innerHTML = '<table><thead><tr><th>Key</th><th>Brand</th><th>Agent</th><th>GA4</th><th>Sheet</th><th>G.Ads</th><th>Actions</th></tr></thead><tbody>' + rows + '</tbody></table>';
    })
    .serverGetClients(currentUser.role, currentUser.clients);
}

// ── CLIENT FORM ──────────────────────────────────────────────
function showAddClient() {
  document.getElementById('clientFormTitle').textContent = 'Add Client';
  document.getElementById('clientFormBtn').textContent   = 'Save Client';
  document.getElementById('editMode').value      = 'add';
  document.getElementById('editClientKey').value  = '';
  document.getElementById('fc_client_key').disabled = false;
  clearClientForm();
  showPage('clientform');
}

function editClient(clientKey) {
  document.getElementById('clientFormTitle').textContent = 'Edit Client';
  document.getElementById('clientFormBtn').textContent   = currentUser.role === 'owner' ? 'Save Changes' : 'Request Edit';
  document.getElementById('editMode').value      = 'edit';
  document.getElementById('editClientKey').value  = clientKey;
  document.getElementById('fc_client_key').disabled = true;
  google.script.run
    .withSuccessHandler(function(res) {
      if (!res.success) { showFlash('error', res.error); return; }
      var c = res.client;
      document.getElementById('fc_client_key').value             = c.client_key;
      document.getElementById('fc_name').value                   = c.name;
      document.getElementById('fc_agent_email').value            = c.agent_email;
      document.getElementById('fc_client_logo_id').value         = c.client_logo_id || '';
      document.getElementById('fc_sheet_id').value               = c.sheet_id || '';
      document.getElementById('fc_meta_access_token').value      = c.meta_access_token;
      document.getElementById('fc_meta_ad_account_ids').value    = c.meta_ad_account_ids;
      document.getElementById('fc_ga4_property_id').value        = c.ga4_property_id;
      document.getElementById('fc_gsc_site_url').value           = c.gsc_site_url;
      document.getElementById('fc_google_ads_enabled').value     = String(c.google_ads_enabled);
      document.getElementById('fc_google_ads_dev_token').value   = c.google_ads_dev_token || '';
      document.getElementById('fc_google_ads_customer_id').value = c.google_ads_customer_id || '';
      document.getElementById('fc_google_ads_mcc_id').value      = c.google_ads_mcc_id || '';
      showPage('clientform');
    })
    .serverGetClientFull(clientKey);
}

function clearClientForm() {
  ['fc_client_key','fc_name','fc_agent_email','fc_client_logo_id','fc_sheet_id',
   'fc_meta_access_token','fc_meta_ad_account_ids','fc_ga4_property_id',
   'fc_gsc_site_url','fc_google_ads_dev_token','fc_google_ads_customer_id','fc_google_ads_mcc_id']
    .forEach(function(id) { document.getElementById(id).value = ''; });
  document.getElementById('fc_google_ads_enabled').value = 'false';
}

function submitClientForm() {
  var mode = document.getElementById('editMode').value;
  var btn  = document.getElementById('clientFormBtn');
  var data = {
    client_key:             document.getElementById('fc_client_key').value.trim().replace(/\s+/g,''),
    name:                   document.getElementById('fc_name').value.trim(),
    agent_email:            document.getElementById('fc_agent_email').value.trim(),
    client_logo_id:         document.getElementById('fc_client_logo_id').value.trim(),
    sheet_id:               document.getElementById('fc_sheet_id').value.trim(),
    meta_access_token:      document.getElementById('fc_meta_access_token').value.trim(),
    meta_ad_account_ids:    document.getElementById('fc_meta_ad_account_ids').value.trim(),
    ga4_property_id:        document.getElementById('fc_ga4_property_id').value.trim(),
    gsc_site_url:           document.getElementById('fc_gsc_site_url').value.trim(),
    google_ads_enabled:     document.getElementById('fc_google_ads_enabled').value === 'true',
    google_ads_dev_token:   document.getElementById('fc_google_ads_dev_token').value.trim(),
    google_ads_customer_id: document.getElementById('fc_google_ads_customer_id').value.trim(),
    google_ads_mcc_id:      document.getElementById('fc_google_ads_mcc_id').value.trim(),
  };
  var required = ['client_key','name','agent_email','meta_access_token','meta_ad_account_ids','ga4_property_id','gsc_site_url'];
  for (var i = 0; i < required.length; i++) {
    if (!data[required[i]]) { showFlash('error', 'Please fill in all required fields (marked with *)'); return; }
  }
  btn.disabled  = true;
  btn.innerHTML = '<span class="spinner"></span> Saving...';
  if (mode === 'add') {
    google.script.run
      .withSuccessHandler(function(res) {
        btn.disabled = false; btn.textContent = 'Save Client';
        if (res.success) { showFlash('success', 'Client added! Now click Create Sheet to set up their data sheet.'); showPage('clients'); }
        else showFlash('error', res.error);
      })
      .serverSaveClient(data, currentUser.role);
  } else {
    var changes = {};
    Object.keys(data).forEach(function(k) { if (k !== 'client_key') changes[k] = data[k]; });
    if (currentUser.role === 'owner') {
      google.script.run
        .withSuccessHandler(function(reqRes) {
          if (!reqRes.success) { btn.disabled = false; btn.textContent = 'Save Changes'; showFlash('error', reqRes.error); return; }
          google.script.run
            .withSuccessHandler(function(appRes) {
              btn.disabled = false; btn.textContent = 'Save Changes';
              if (appRes.success) { showFlash('success', 'Client updated!'); showPage('clients'); }
              else showFlash('error', appRes.error);
            })
            .serverApproveEdit(reqRes.requestId, currentUser.username);
        })
        .serverRequestClientEdit(data.client_key, changes, currentUser.username);
    } else {
      google.script.run
        .withSuccessHandler(function(res) {
          btn.disabled = false; btn.textContent = 'Request Edit';
          if (res.success) { showFlash('success', 'Edit request submitted! Awaiting owner approval.'); showPage('clients'); }
          else showFlash('error', res.error);
        })
        .serverRequestClientEdit(data.client_key, changes, currentUser.username);
    }
  }
}

function deleteClient(clientKey, clientName) {
  if (!confirm('Delete client "' + clientName + '"?\n\nThis cannot be undone.')) return;
  google.script.run
    .withSuccessHandler(function(res) {
      if (res.success) { showFlash('success', clientName + ' deleted.'); loadClientsTable(); loadDashboard(); }
      else showFlash('error', res.error);
    })
    .serverDeleteClient(clientKey, currentUser.role);
}

// ── CREATE SHEET ─────────────────────────────────────────────
function createClientSheet(clientKey, clientName) {
  if (!confirm('Create a new Google Sheet for "' + clientName + '"?\n\nThis will create a fresh sheet with all required tabs and auto-link it to this client.')) return;
  showFlash('info', 'Creating sheet for ' + clientName + '... please wait.');
  google.script.run
    .withSuccessHandler(function(res) {
      if (!res.success) { showFlash('error', 'Failed: ' + res.error); return; }
      var msg = document.getElementById('flashMsg');
      msg.className = 'flash-msg alert alert-success';
      msg.innerHTML = 'Sheet created for ' + esc(clientName) + '! '
        + '<a href="' + res.sheetUrl + '" target="_blank" style="color:var(--green);font-weight:700;text-decoration:underline">Open Sheet</a>';
      msg.style.display = 'block';
      setTimeout(function() { msg.style.display = 'none'; }, 15000);
      if (currentPage === 'clients')   loadClientsTable();
      if (currentPage === 'dashboard') loadDashboard();
    })
    .withFailureHandler(function(err) { showFlash('error', 'Error: ' + err.message); })
    .serverCreateClientSheet(clientKey, clientName, currentUser.role, currentUser.clients);
}

function openClientSheet(clientKey) {
  google.script.run
    .withSuccessHandler(function(res) {
      if (!res.success) { showFlash('error', res.error); return; }
      window.open(res.sheetUrl, '_blank');
    })
    .withFailureHandler(function(err) { showFlash('error', 'Error: ' + err.message); })
    .serverGetClientSheetUrl(clientKey);
}

// ── GENERATE REPORT ──────────────────────────────────────────
function loadGenerateDropdown() {
  google.script.run
    .withSuccessHandler(function(clients) {
      if (!Array.isArray(clients)) return;
      var sel = document.getElementById('gen_client');
      while (sel.options.length > 2) sel.remove(2);
      clients.forEach(function(c) {
        var opt = document.createElement('option');
        opt.value = c.client_key; opt.textContent = c.name;
        sel.appendChild(opt);
      });
    })
    .serverGetClients(currentUser.role, currentUser.clients);
}

function toggleMonthInput() {
  var mode = document.getElementById('gen_mode').value;
  document.getElementById('gen_month_wrap').style.display = mode === 'manual' ? 'block' : 'none';
}

function generateReport() {
  var clientKey = document.getElementById('gen_client').value;
  var mode      = document.getElementById('gen_mode').value;
  var monthRaw  = document.getElementById('gen_month').value;
  var btn       = document.getElementById('genBtn');
  if (!clientKey) { showFlash('error', 'Please select a client'); return; }
  if (mode === 'manual' && !monthRaw) { showFlash('error', 'Please pick a month'); return; }
  var monthStr = '';
  if (mode === 'manual' && monthRaw) {
    var parts = monthRaw.split('-');
    var names = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    monthStr  = names[parseInt(parts[1])-1] + ' ' + parts[0];
  }
  btn.disabled  = true;
  btn.innerHTML = '<span class="spinner"></span> Generating... (this may take 1-2 min)';
  var resultsEl = document.getElementById('genResults');
  var innerEl   = document.getElementById('genResultsInner');
  resultsEl.style.display = 'none';

  function onFail(err) {
    btn.disabled = false; btn.innerHTML = 'Generate Report';
    showFlash('error', 'Error: ' + err.message);
  }

  if (clientKey === 'ALL') {
    google.script.run
      .withSuccessHandler(function(res) {
        btn.disabled = false; btn.innerHTML = 'Generate Report';
        resultsEl.style.display = 'block';
        if (!res.success) { innerEl.innerHTML = '<div class="alert alert-error">' + esc(res.error) + '</div>'; return; }
        var html = '';
        res.results.forEach(function(r) {
          html += '<div class="result-item ' + (r.success ? 'success' : 'failed') + '">'
            + '<div class="result-name">' + esc(r.client) + '</div>';
          if (r.success) {
            html += '<div class="result-links">'
              + '<a class="result-link" href="' + r.slideUrl + '" target="_blank">Open Slides</a>'
              + '<a class="result-link pptx" href="' + r.pptxUrl + '" target="_blank">Download PPTX</a>'
              + '</div>';
          } else { html += '<div class="result-error">Failed: ' + esc(r.error) + '</div>'; }
          html += '</div>';
        });
        innerEl.innerHTML = html;
      })
      .withFailureHandler(onFail)
      .serverGenerateAllReports(monthStr, currentUser.username, currentUser.role, currentUser.clients);
  } else {
    google.script.run
      .withSuccessHandler(function(res) {
        btn.disabled = false; btn.innerHTML = 'Generate Report';
        resultsEl.style.display = 'block';
        if (!res.success) { innerEl.innerHTML = '<div class="alert alert-error">' + esc(res.error) + '</div>'; return; }
        innerEl.innerHTML = '<div class="result-item success">'
          + '<div class="result-name">Report ready - ' + esc(res.month) + '</div>'
          + '<div class="result-links">'
          + '<a class="result-link" href="' + res.slideUrl + '" target="_blank">Open in Google Slides</a>'
          + '<a class="result-link pptx" href="' + res.pptxUrl + '" target="_blank">Download as PPTX</a>'
          + '</div></div>';
        loadApprovalsBadge();
      })
      .withFailureHandler(onFail)
      .serverGenerateReport(clientKey, monthStr, currentUser.username, currentUser.role, currentUser.clients);
  }
}

// ── REPORT LOGS ──────────────────────────────────────────────
function loadLogs() {
  var el = document.getElementById('logsTable');
  el.innerHTML = '<div class="loading-state">Loading logs...</div>';
  google.script.run
    .withSuccessHandler(function(logs) {
      if (!Array.isArray(logs) || !logs.length) {
        el.innerHTML = '<div class="empty-state"><div class="empty-state-icon">&#9636;</div><div class="empty-state-text">No reports generated yet.</div></div>';
        return;
      }
      var rows = '';
      logs.forEach(function(l) {
        rows += '<tr>'
          + '<td><strong>' + esc(l.client_name) + '</strong><br><span style="font-size:11px;color:var(--muted)">' + esc(l.client_key) + '</span></td>'
          + '<td>' + esc(l.month_label) + '</td>'
          + '<td><span class="status-badge ' + l.status + '">' + l.status + '</span></td>'
          + '<td>' + esc(l.generated_by) + '</td>'
          + '<td style="font-size:12px;color:var(--muted)">' + formatDate(l.generated_at) + '</td>'
          + '<td>'
          + (l.slides_url ? '<a class="result-link" href="' + l.slides_url + '" target="_blank" style="font-size:11px">Slides</a> ' : '')
          + (l.pptx_url   ? '<a class="result-link pptx" href="' + l.pptx_url + '" target="_blank" style="font-size:11px">PPTX</a>' : '')
          + (l.error_msg  ? '<span style="color:var(--red);font-size:11px">' + esc(l.error_msg.slice(0,60)) + '</span>' : '')
          + '</td></tr>';
      });
      el.innerHTML = '<table><thead><tr><th>Client</th><th>Month</th><th>Status</th><th>Generated By</th><th>Date</th><th>Links</th></tr></thead><tbody>' + rows + '</tbody></table>';
    })
    .serverGetReportLogs(currentUser.role, currentUser.clients, 100);
}

// ── USERS ────────────────────────────────────────────────────
function loadUsersTable() {
  if (currentUser.role !== 'owner') return;
  var el = document.getElementById('usersTable');
  el.innerHTML = '<div class="loading-state">Loading...</div>';
  google.script.run
    .withSuccessHandler(function(users) {
      if (!Array.isArray(users) || !users.length) { el.innerHTML = '<div class="loading-state">No users found</div>'; return; }
      var rows = '';
      users.forEach(function(u) {
        var isActive = u.is_active === true || u.is_active === 'TRUE';
        rows += '<tr>'
          + '<td class="code-cell">' + esc(u.username) + '</td>'
          + '<td><strong>' + esc(u.name) + '</strong></td>'
          + '<td><span class="status-badge ' + (u.role === 'owner' ? 'approved' : 'pending') + '">' + esc(u.role) + '</span></td>'
          + '<td style="font-size:12px">' + esc(u.assigned_clients) + '</td>'
          + '<td><span class="status-badge ' + (isActive ? 'success' : 'failed') + '">' + (isActive ? 'Active' : 'Inactive') + '</span></td>'
          + '<td>' + (u.username !== 'owner' ? '<button class="btn-icon" style="color:var(--red)" onclick="deleteUser(\'' + esc(u.username) + '\')">Remove</button>' : '<span style="color:var(--muted);font-size:12px">Protected</span>') + '</td>'
          + '</tr>';
      });
      el.innerHTML = '<table><thead><tr><th>Username</th><th>Name</th><th>Role</th><th>Assigned Clients</th><th>Status</th><th>Actions</th></tr></thead><tbody>' + rows + '</tbody></table>';
    })
    .serverGetUsers();
}

function showAddUser() {
  document.getElementById('userModalTitle').textContent = 'Add User';
  ['um_username','um_name','um_password','um_clients'].forEach(function(id) { document.getElementById(id).value = ''; });
  document.getElementById('um_role').value = 'manager';
  document.getElementById('userModal').style.display = 'flex';
}

function closeUserModal() { document.getElementById('userModal').style.display = 'none'; }

function saveUser() {
  var data = {
    username:         document.getElementById('um_username').value.trim(),
    name:             document.getElementById('um_name').value.trim(),
    password:         document.getElementById('um_password').value,
    role:             document.getElementById('um_role').value,
    assigned_clients: document.getElementById('um_clients').value.trim(),
  };
  if (!data.username || !data.name || !data.password) { alert('Please fill in all fields'); return; }
  google.script.run
    .withSuccessHandler(function(res) {
      if (res.success) { closeUserModal(); showFlash('success','User added!'); loadUsersTable(); }
      else alert('Error: ' + res.error);
    })
    .serverSaveUser(data, currentUser.role);
}

function deleteUser(username) {
  if (!confirm('Remove user "' + username + '"?')) return;
  google.script.run
    .withSuccessHandler(function(res) {
      if (res.success) { showFlash('success','User removed.'); loadUsersTable(); }
      else showFlash('error', res.error);
    })
    .serverDeleteUser(username, currentUser.role);
}

// ── APPROVALS ────────────────────────────────────────────────
function loadApprovals() {
  if (currentUser.role !== 'owner') return;
  var el = document.getElementById('approvalsTable');
  el.innerHTML = '<div class="loading-state">Loading...</div>';
  google.script.run
    .withSuccessHandler(function(reqs) {
      if (!Array.isArray(reqs) || !reqs.length) {
        el.innerHTML = '<div class="empty-state"><div class="empty-state-icon">*</div><div class="empty-state-text">No pending approvals.</div></div>';
        updateApprovalBadge(0); return;
      }
      updateApprovalBadge(reqs.length);
      var rows = '';
      reqs.forEach(function(r) {
        var changes = {};
        try { changes = JSON.parse(r.field_changes); } catch(e) {}
        var changeList = Object.keys(changes).map(function(k) {
          return '<span style="font-size:11px"><b>' + esc(k) + '</b>: ' + esc(String(changes[k]).slice(0,40)) + '</span>';
        }).join('<br>');
        rows += '<tr>'
          + '<td class="code-cell">' + esc(r.client_key) + '</td>'
          + '<td>' + esc(r.requested_by) + '</td>'
          + '<td>' + changeList + '</td>'
          + '<td style="font-size:12px;color:var(--muted)">' + formatDate(r.requested_at) + '</td>'
          + '<td><div style="display:flex;gap:6px">'
          + '<button class="btn btn-success sm" onclick="approveEdit(\'' + esc(r.id) + '\')">Approve</button>'
          + '<button class="btn btn-danger sm" onclick="rejectEdit(\'' + esc(r.id) + '\')">Reject</button>'
          + '</div></td></tr>';
      });
      el.innerHTML = '<table><thead><tr><th>Client</th><th>Requested By</th><th>Changes</th><th>Date</th><th>Actions</th></tr></thead><tbody>' + rows + '</tbody></table>';
    })
    .serverGetEditRequests();
}

function loadApprovalsBadge() {
  if (currentUser.role !== 'owner') return;
  google.script.run
    .withSuccessHandler(function(reqs) { if (Array.isArray(reqs)) updateApprovalBadge(reqs.length); })
    .serverGetEditRequests();
}

function updateApprovalBadge(count) {
  var badge = document.getElementById('approvalBadge');
  if (count > 0) { badge.textContent = count; badge.style.display = 'inline-block'; }
  else badge.style.display = 'none';
}

function approveEdit(requestId) {
  google.script.run
    .withSuccessHandler(function(res) {
      if (res.success) { showFlash('success','Edit approved!'); loadApprovals(); }
      else showFlash('error', res.error);
    })
    .serverApproveEdit(requestId, currentUser.username);
}

function rejectEdit(requestId) {
  if (!confirm('Reject this edit request?')) return;
  google.script.run
    .withSuccessHandler(function(res) {
      if (res.success) { showFlash('success','Edit rejected.'); loadApprovals(); }
      else showFlash('error', res.error);
    })
    .serverRejectEdit(requestId, currentUser.username);
}

// ── UTILITIES ────────────────────────────────────────────────
function showFlash(type, msg) {
  var el = document.getElementById('flashMsg');
  el.className = 'flash-msg alert alert-' + type;
  el.textContent = msg;
  el.style.display = 'block';
  setTimeout(function() { el.style.display = 'none'; }, 5000);
}

function togglePwd(inputId, btn) {
  var inp = document.getElementById(inputId);
  if (inp.type === 'password') { inp.type = 'text'; btn.textContent = 'Hide'; }
  else { inp.type = 'password'; btn.textContent = 'Show'; }
}

function esc(str) {
  return String(str || '')
    .replace(/&/g,'&amp;').replace(/</g,'&lt;')
    .replace(/>/g,'&gt;').replace(/"/g,'&quot;')
    .replace(/'/g,'&#39;');
}

function formatDate(str) {
  if (!str) return '-';
  try {
    var d = new Date(str);
    return d.toLocaleDateString('en-IN', { day:'2-digit', month:'short', year:'numeric' })
      + ' ' + d.toLocaleTimeString('en-IN', { hour:'2-digit', minute:'2-digit' });
  } catch(e) { return str; }
}

document.addEventListener('keydown', function(e) {
  if (e.key === 'Enter' && document.getElementById('loginScreen').style.display !== 'none') { doLogin(); }
});
</script>