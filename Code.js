// ============================================================
//  DIGIFYCE — WEB APP BACKEND (Code.gs)
//  Google Apps Script Web App
// ============================================================

const SYSTEM_SHEET_ID  = "1obVFp64gSgwubUPU389Us4i7RN6NmgRookIf4MlTdro";
const DIGIFYCE_LOGO_ID = "127Uj5j2xnyTd5-Uz_YaeNpueycNRjYho";
const OWNER_USERNAME   = "owner";
const OWNER_PASSWORD   = "Digifyce@Owner2025";
const OPENAI_API_KEY   = "YOUR_OPENAI_API_KEY"; // ← REPLACE THIS WITH YOUR ACTUAL API KEY
const OPENAI_MODEL          = "gpt-4o";
const REPORTS_FOLDER_NAME  = "Digifyce Reports";
const REPORTS_FOLDER_ID    = "1FIOOazU78Rt-SnmOkoJtebE_m4RLiuHS";
const SHARE_EMAIL          = "digifycecbe@gmail.com";


// ============================================================
// SECTION 1 — WEB APP ENTRY POINTS
// ============================================================

function doGet(e) {
  return HtmlService
    .createTemplateFromFile("Index")
    .evaluate()
    .setTitle("Digifyce Report Dashboard")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


// ============================================================
// SECTION 2 — SYSTEM SHEET SETUP
// ============================================================

function getSystemSheet() {
  return SpreadsheetApp.openById(SYSTEM_SHEET_ID);
}

function getOrCreateTab(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground("#0D2B6E").setFontColor("#FFFFFF").setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function initSystemSheets() {
  try {
    const ss = getSystemSheet();
    getOrCreateTab(ss, "_Users",        ["username","password_hash","name","role","assigned_clients","created_at","is_active"]);
    getOrCreateTab(ss, "_Clients",      ["client_key","name","agent_email","client_logo_id","sheet_id","meta_access_token","meta_ad_account_ids","google_ads_enabled","google_ads_dev_token","google_ads_customer_id","google_ads_mcc_id","ga4_property_id","gsc_site_url","created_at","updated_at","is_active"]);
    getOrCreateTab(ss, "_ReportLogs",   ["id","client_key","client_name","month_label","generated_by","status","slides_url","pptx_url","error_msg","generated_at"]);
    getOrCreateTab(ss, "_EditRequests", ["id","client_key","requested_by","field_changes","status","requested_at","reviewed_at","reviewed_by"]);
    const usersSheet = ss.getSheetByName("_Users");
    if (usersSheet.getLastRow() <= 1) {
      usersSheet.appendRow([OWNER_USERNAME, hashPassword(OWNER_PASSWORD), "Owner", "owner", "ALL", new Date().toISOString(), true]);
    }
    return { success: true };
  } catch(e) {
    Logger.log("initSystemSheets error: " + e.message);
    return { success: false, error: e.message };
  }
}


// ============================================================
// SECTION 3 — AUTH
// ============================================================

function hashPassword(password) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password + "digifyce_salt_2025");
  return bytes.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

function serverLogin(username, password) {
  try {
    initSystemSheets();
    const ss = getSystemSheet(), sheet = ss.getSheetByName("_Users");
    const data = sheet.getDataRange().getValues(), hdrs = data[0];
    const idx = {
      username: hdrs.indexOf("username"),
      hash:     hdrs.indexOf("password_hash"),
      name:     hdrs.indexOf("name"),
      role:     hdrs.indexOf("role"),
      clients:  hdrs.indexOf("assigned_clients"),
      active:   hdrs.indexOf("is_active")
    };
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (String(row[idx.username]).trim().toLowerCase() !== username.toLowerCase()) continue;
      if (!row[idx.active]) return { success: false, error: "Account is inactive" };
      if (hashPassword(password) !== String(row[idx.hash]).trim()) return { success: false, error: "Invalid password" };
      return { success: true, user: { username: row[idx.username], name: row[idx.name], role: row[idx.role], clients: String(row[idx.clients]) } };
    }
    return { success: false, error: "User not found" };
  } catch(e) { return { success: false, error: e.message }; }
}


// ============================================================
// SECTION 4 — USER MANAGEMENT
// ============================================================

function serverGetUsers() {
  try {
    const sheet = getSystemSheet().getSheetByName("_Users");
    const data = sheet.getDataRange().getValues(), hdrs = data[0];
    return data.slice(1).map(row => { const obj = {}; hdrs.forEach((h,i) => obj[h] = row[i]); obj.password_hash = ""; return obj; });
  } catch(e) { return { error: e.message }; }
}

function serverSaveUser(userData, currentUserRole) {
  try {
    if (currentUserRole !== "owner") return { success: false, error: "Owner only" };
    initSystemSheets();
    const sheet = getSystemSheet().getSheetByName("_Users");
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).toLowerCase() === userData.username.toLowerCase()) return { success: false, error: "Username already exists" };
    }
    sheet.appendRow([userData.username, hashPassword(userData.password), userData.name, userData.role, userData.assigned_clients || "NONE", new Date().toISOString(), true]);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

function serverUpdateUser(userData, currentUserRole) {
  try {
    if (currentUserRole !== "owner") return { success: false, error: "Owner only" };
    const sheet = getSystemSheet().getSheetByName("_Users");
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).toLowerCase() === userData.username.toLowerCase()) {
        sheet.getRange(i+1, 4).setValue(userData.role);
        sheet.getRange(i+1, 5).setValue(userData.assigned_clients || "NONE");
        sheet.getRange(i+1, 7).setValue(userData.is_active !== false);
        if (userData.new_password && userData.new_password.trim() !== "") sheet.getRange(i+1, 2).setValue(hashPassword(userData.new_password));
        return { success: true };
      }
    }
    return { success: false, error: "User not found" };
  } catch(e) { return { success: false, error: e.message }; }
}

function serverDeleteUser(username, currentUserRole) {
  try {
    if (currentUserRole !== "owner") return { success: false, error: "Owner only" };
    if (username === OWNER_USERNAME) return { success: false, error: "Cannot delete owner" };
    const sheet = getSystemSheet().getSheetByName("_Users");
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).toLowerCase() === username.toLowerCase()) { sheet.deleteRow(i+1); return { success: true }; }
    }
    return { success: false, error: "User not found" };
  } catch(e) { return { success: false, error: e.message }; }
}


// ============================================================
// SECTION 5 — CLIENT MANAGEMENT
// ============================================================

function serverGetClients(userRole, userClients) {
  try {
    const sheet = getSystemSheet().getSheetByName("_Clients");
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    const hdrs = data[0];
    let rows = data.slice(1).map(row => {
      const obj = {}; hdrs.forEach((h,i) => obj[h] = row[i]);
      obj.meta_access_token = obj.meta_access_token ? obj.meta_access_token.slice(0,8) + "••••••••" : "";
      return obj;
    }).filter(r => r.is_active !== false && r.is_active !== "FALSE" && r.is_active !== false);
    if (userRole !== "owner") {
      const allowed = String(userClients).split(",").map(s => s.trim().toLowerCase());
      if (!allowed.includes("all")) rows = rows.filter(r => allowed.includes(String(r.client_key).toLowerCase()));
    }
    return rows;
  } catch(e) { return { error: e.message }; }
}

function serverGetClientFull(clientKey) {
  try {
    const sheet = getSystemSheet().getSheetByName("_Clients");
    if (!sheet) return { success: false, error: "Clients sheet not found" };
    const data = sheet.getDataRange().getValues(), hdrs = data[0];
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === clientKey.trim()) {
        const obj = {}; hdrs.forEach((h,j) => obj[h] = data[i][j]); return { success: true, client: obj };
      }
    }
    return { success: false, error: "Client not found: " + clientKey };
  } catch(e) { return { success: false, error: e.message }; }
}

function serverSaveClient(clientData, currentUserRole) {
  try {
    initSystemSheets();
    const sheet = getSystemSheet().getSheetByName("_Clients");
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === clientData.client_key.trim().toLowerCase()) return { success: false, error: "Client key already exists." };
    }
    const now = new Date().toISOString();
    sheet.appendRow([
      clientData.client_key.trim(), clientData.name, clientData.agent_email,
      clientData.client_logo_id || "", clientData.sheet_id || "",
      clientData.meta_access_token, clientData.meta_ad_account_ids,
      clientData.google_ads_enabled || false, clientData.google_ads_dev_token || "",
      clientData.google_ads_customer_id || "", clientData.google_ads_mcc_id || "",
      clientData.ga4_property_id, clientData.gsc_site_url, now, now, true
    ]);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

function serverRequestClientEdit(clientKey, changes, requestedBy) {
  try {
    initSystemSheets();
    const sheet = getSystemSheet().getSheetByName("_EditRequests");
    const id = "ER_" + Date.now();
    sheet.appendRow([id, clientKey, requestedBy, JSON.stringify(changes), "pending", new Date().toISOString(), "", ""]);
    return { success: true, requestId: id, message: "Edit request submitted." };
  } catch(e) { return { success: false, error: e.message }; }
}

function serverGetEditRequests() {
  try {
    const sheet = getSystemSheet().getSheetByName("_EditRequests");
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues(), hdrs = data[0];
    return data.slice(1).filter(r => r[4] === "pending").map(row => { const obj = {}; hdrs.forEach((h,i) => obj[h] = row[i]); return obj; });
  } catch(e) { return { error: e.message }; }
}

function serverApproveEdit(requestId, reviewedBy) {
  try {
    const reqSheet = getSystemSheet().getSheetByName("_EditRequests");
    const data = reqSheet.getDataRange().getValues(), hdrs = data[0];
    const idIdx = hdrs.indexOf("id");
    for (let i = 1; i < data.length; i++) {
      if (data[i][idIdx] !== requestId) continue;
      const changes = JSON.parse(data[i][hdrs.indexOf("field_changes")]);
      const clientKey = data[i][hdrs.indexOf("client_key")];
      const clientSheet = getSystemSheet().getSheetByName("_Clients");
      const cData = clientSheet.getDataRange().getValues(), cHdrs = cData[0];
      for (let j = 1; j < cData.length; j++) {
        if (String(cData[j][0]).trim() === clientKey) {
          Object.entries(changes).forEach(([field, value]) => {
            const colIdx = cHdrs.indexOf(field);
            if (colIdx >= 0) clientSheet.getRange(j+1, colIdx+1).setValue(value);
          });
          const updIdx = cHdrs.indexOf("updated_at");
          if (updIdx >= 0) clientSheet.getRange(j+1, updIdx+1).setValue(new Date().toISOString());
          break;
        }
      }
      reqSheet.getRange(i+1, hdrs.indexOf("status")+1).setValue("approved");
      reqSheet.getRange(i+1, hdrs.indexOf("reviewed_at")+1).setValue(new Date().toISOString());
      reqSheet.getRange(i+1, hdrs.indexOf("reviewed_by")+1).setValue(reviewedBy);
      return { success: true };
    }
    return { success: false, error: "Request not found" };
  } catch(e) { return { success: false, error: e.message }; }
}

function serverRejectEdit(requestId, reviewedBy) {
  try {
    const sheet = getSystemSheet().getSheetByName("_EditRequests");
    const data = sheet.getDataRange().getValues(), hdrs = data[0];
    for (let i = 1; i < data.length; i++) {
      if (data[i][hdrs.indexOf("id")] !== requestId) continue;
      sheet.getRange(i+1, hdrs.indexOf("status")+1).setValue("rejected");
      sheet.getRange(i+1, hdrs.indexOf("reviewed_at")+1).setValue(new Date().toISOString());
      sheet.getRange(i+1, hdrs.indexOf("reviewed_by")+1).setValue(reviewedBy);
      return { success: true };
    }
    return { success: false, error: "Request not found" };
  } catch(e) { return { success: false, error: e.message }; }
}

function serverDeleteClient(clientKey, currentUserRole) {
  try {
    if (currentUserRole !== "owner") return { success: false, error: "Only owner can delete clients" };
    const sheet = getSystemSheet().getSheetByName("_Clients");
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === clientKey) {
        sheet.getRange(i+1, data[0].indexOf("is_active")+1).setValue(false);
        return { success: true };
      }
    }
    return { success: false, error: "Client not found" };
  } catch(e) { return { success: false, error: e.message }; }
}


// ============================================================
// SECTION 6 — REPORT GENERATION
// ============================================================

function serverGenerateReport(clientKey, month, generatedBy, userRole, userClients) {
  try {
    // Access check
    if (userRole !== "owner") {
      const allowed = String(userClients).split(",").map(s => s.trim().toLowerCase());
      if (!allowed.includes("all") && !allowed.includes(clientKey.toLowerCase())) return { success: false, error: "Access denied" };
    }

    // Get client
    const result = serverGetClientFull(clientKey);
    if (!result.success) return { success: false, error: result.error };
    const c = result.client;

    // Validate sheet is linked
    if (!c.sheet_id || String(c.sheet_id).trim() === "") {
      return { success: false, error: "No Google Sheet linked to client '" + c.name + "'. Please create a sheet first from the Clients page." };
    }

    // Validate Meta token exists
    if (!c.meta_access_token || String(c.meta_access_token).trim() === "") {
      return { success: false, error: "Meta Access Token is missing for client '" + c.name + "'. Please update the client settings." };
    }

    // Validate GA4 property
    if (!c.ga4_property_id || String(c.ga4_property_id).trim() === "") {
      return { success: false, error: "GA4 Property ID is missing for client '" + c.name + "'." };
    }

    const CLIENT = {
      clientKey:          clientKey,
      name:               c.name,
      agentEmail:         c.agent_email,
      digifyceLogoFileId: DIGIFYCE_LOGO_ID,
      clientLogoFileId:   c.client_logo_id || "",
      sheetId:            c.sheet_id,
      meta: {
        accessToken:  c.meta_access_token,
        adAccountIds: String(c.meta_ad_account_ids).split(",").map(s => s.trim()).filter(s => s)
      },
      googleAds: {
        enabled:         c.google_ads_enabled === true || c.google_ads_enabled === "TRUE",
        developerToken:  c.google_ads_dev_token || "",
        customerId:      c.google_ads_customer_id || "",
        managerCustomerId: c.google_ads_mcc_id || ""
      },
      ga4: { propertyId: c.ga4_property_id },
      gsc: { siteUrl:    c.gsc_site_url },
    };

    // Validate at least one ad account
    if (!CLIENT.meta.adAccountIds.length) {
      return { success: false, error: "No Meta Ad Account IDs configured for client '" + c.name + "'." };
    }

    const M = getMonthConfig(month || "");
    const slideUrl = _runReportForClient(CLIENT, M);
    const fileId = slideUrl.split("/d/")[1].split("/")[0];
    moveFileToReportsFolder(fileId);
    const pptxUrl = "https://docs.google.com/presentation/d/" + fileId + "/export/pptx";
    _logReport(clientKey, c.name, M.currLabel, generatedBy, "success", slideUrl, pptxUrl, "");
    return { success: true, slideUrl, pptxUrl, month: M.currFull };

  } catch(e) {
    Logger.log("serverGenerateReport ERROR: " + e.message + "\n" + e.stack);
    _logReport(clientKey, clientKey, month || "auto", generatedBy, "failed", "", "", e.message);
    return { success: false, error: e.message };
  }
}

function serverGenerateAllReports(month, generatedBy, userRole, userClients) {
  try {
    const clients = serverGetClients(userRole, userClients);
    if (!Array.isArray(clients)) return { success: false, error: "Could not load clients" };
    const results = [];
    clients.forEach(c => {
      try {
        const r = serverGenerateReport(c.client_key, month, generatedBy, userRole, userClients);
        results.push({ client: c.name, key: c.client_key, success: r.success, slideUrl: r.slideUrl || "", pptxUrl: r.pptxUrl || "", error: r.error || "" });
      } catch(e) {
        results.push({ client: c.name, key: c.client_key, success: false, slideUrl: "", pptxUrl: "", error: e.message });
      }
    });
    return { success: true, results };
  } catch(e) { return { success: false, error: e.message }; }
}

function _logReport(clientKey, clientName, monthLabel, generatedBy, status, slideUrl, pptxUrl, errorMsg) {
  try {
    getSystemSheet().getSheetByName("_ReportLogs")
      .appendRow(["RPT_" + Date.now(), clientKey, clientName, monthLabel, generatedBy, status, slideUrl, pptxUrl, errorMsg, new Date().toISOString()]);
  } catch(e) { Logger.log("Log error: " + e.message); }
}

function getReportsFolder() {
  try {
    return DriveApp.getFolderById(REPORTS_FOLDER_ID);
  } catch(e) {
    Logger.log("  ⚠️ Could not open reports folder: " + e.message);
    return null;
  }
}

function moveFileToReportsFolder(fileId) {
  // DriveApp.getFileById is restricted on this account — silently skip
  Logger.log("  ℹ️ Folder move skipped (Drive API restricted).");
}

// Run this to verify the folder is accessible
function testReportsFolder() {
  try {
    const folder = getReportsFolder();
    if (!folder) { Logger.log("❌ Folder not found"); return; }
    Logger.log("✅ Folder found: " + folder.getName());
    Logger.log("✅ Folder ID: " + folder.getId());
    Logger.log("✅ URL: https://drive.google.com/drive/folders/" + folder.getId());
  } catch(e) {
    Logger.log("❌ " + e.message);
  }
}

function serverGetReportLogs(userRole, userClients, limit) {
  try {
    const sheet = getSystemSheet().getSheetByName("_ReportLogs");
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues(), hdrs = data[0];
    let rows = data.slice(1).reverse().map(row => { const obj = {}; hdrs.forEach((h,i) => obj[h] = row[i]); return obj; });
    if (userRole !== "owner") {
      const allowed = String(userClients).split(",").map(s => s.trim().toLowerCase());
      if (!allowed.includes("all")) rows = rows.filter(r => allowed.includes(String(r.client_key).toLowerCase()));
    }
    return rows.slice(0, limit || 50);
  } catch(e) { return { error: e.message }; }
}


// ============================================================
// SECTION 7A — TAB STRUCTURE DEFINITION
// ============================================================

const TAB_STRUCTURE = {
  "Retention": {
    titleRow:  ["DIGIFYCE — RETENTION DATA"],
    noteRow:   ["One row per channel per month. Channel must be exactly: Email, WhatsApp, Push"],
    headerRow: ["Month","Channel","Sent","Open / Read Rate (%)","Click Rate (%)","Revenue (₹)","ROAS","Note"],
    sampleRows:[
      ["Jan 2026","Email","","","","","",""],
      ["Jan 2026","WhatsApp","","","","","",""],
      ["Jan 2026","Push","","","","","",""],
    ],
  },
  "Retention_Summary": {
    titleRow:  ["DIGIFYCE — RETENTION SUMMARY"],
    noteRow:   ["One row per month. Fill BOTH current and previous month rows for returning rate MoM change."],
    headerRow: ["Month","Total Retention Revenue (₹)","Returning Rate (%)","Notes","Avg Repeat AOV (₹)","Total Retention Orders"],
    sampleRows:[
      ["Jan 2026","","","","",""],
      ["Feb 2026","","","","",""],
    ],
  },
  "Shopify_Summary": {
    titleRow:  ["DIGIFYCE — SHOPIFY SUMMARY"],
    noteRow:   ["Fill BOTH current AND previous month rows. Multiple rows per month are summed automatically."],
    headerRow: ["Month","Gross Sales (₹)","Total Orders","Total New Customers","Avg Order Value (₹)"],
    sampleRows:[
      ["Jan 2026","","","",""],
      ["Feb 2026","","","",""],
    ],
  },
  "Shopify_Products": {
    titleRow:  ["DIGIFYCE — SHOPIFY PRODUCTS"],
    noteRow:   ["One row per product per month. Only current month needed."],
    headerRow: ["Month","Product Name","Net Orders","Revenue (₹)"],
    sampleRows:[
      ["Jan 2026","Product 1","",""],
      ["Jan 2026","Product 2","",""],
    ],
  },
  "Shopify_Locations": {
    titleRow:  ["DIGIFYCE — SHOPIFY LOCATIONS"],
    noteRow:   ["One row per city per month. Only current month needed."],
    headerRow: ["Month","Location (City, State)","Orders","Revenue (₹)"],
    sampleRows:[
      ["Jan 2026","Chennai, Tamil Nadu","",""],
      ["Jan 2026","Mumbai, Maharashtra","",""],
    ],
  },
  "Strategy": {
    titleRow:  ["DIGIFYCE — STRATEGY & PROJECTIONS"],
    noteRow:   ["Sections: APPROACH_1, APPROACH_2, APPROACH_3, PROJECTION. Only current month needed."],
    headerRow: ["Month","Section","Title / Label","Point 1 Heading","Point 1 Body","Point 2 Heading","Point 2 Body","Point 3 Heading","Point 3 Body","Point 4 Heading","Point 5 Heading","Point 6 Heading"],
    sampleRows:[
      ["Jan 2026","APPROACH_1","Approach Title","Heading 1","Body 1","Heading 2","Body 2","","","","",""],
      ["Jan 2026","APPROACH_2","Approach Title","Heading 1","Body 1","","","","","","",""],
      ["Jan 2026","APPROACH_3","Approach Title","Heading 1","Body 1","","","","","","",""],
      ["Jan 2026","PROJECTION","","500000","","50000","","100","3.5","Growth note","ROAS note",""],
    ],
  },
};

const TAB_ORDER = ["Retention","Retention_Summary","Shopify_Summary","Shopify_Products","Shopify_Locations","Strategy"];


// ============================================================
// SECTION 7B — TAB BUILDER
// ============================================================

function _buildTab(sheet, def) {
  const numCols = def.headerRow.length;
  sheet.getRange(1,1,1,numCols)
    .setValues([def.titleRow.concat(Array(numCols-1).fill(""))])
    .setBackground("#0D2B6E").setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(11);
  sheet.getRange(2,1,1,numCols)
    .setValues([def.noteRow.concat(Array(numCols-1).fill(""))])
    .setBackground("#E3F2FD").setFontColor("#546E7A").setFontSize(9).setFontStyle("italic");
  sheet.getRange(3,1,1,numCols)
    .setValues([def.headerRow])
    .setBackground("#0D2B6E").setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(10);
  if (def.sampleRows && def.sampleRows.length > 0) {
    sheet.getRange(4,1,def.sampleRows.length,numCols)
      .setValues(def.sampleRows)
      .setBackground("#F4F7FB").setFontColor("#90A4AE").setFontSize(9).setFontStyle("italic");
  }
  sheet.setFrozenRows(3);
  sheet.autoResizeColumns(1, numCols);
  sheet.setColumnWidth(1, 100);
  sheet.setRowHeight(1, 36);
  sheet.setRowHeight(2, 28);
  sheet.setRowHeight(3, 30);
}


// ============================================================
// SECTION 7C — CREATE NEW SHEET
// ============================================================

function serverCreateClientSheet(clientKey, clientName, currentUserRole, currentUserClients) {
  try {
    if (currentUserRole !== "owner") {
      const allowed = String(currentUserClients).split(",").map(s => s.trim().toLowerCase());
      if (!allowed.includes("all") && !allowed.includes(clientKey.toLowerCase())) return { success: false, error: "Access denied" };
    }
    const ss = SpreadsheetApp.create("Digifyce — " + clientName + " — Data Sheet");
    const ssId = ss.getId();
    const ssUrl = "https://docs.google.com/spreadsheets/d/" + ssId;
    const defaultSheet = ss.getSheets()[0];
    TAB_ORDER.forEach((tabName, idx) => {
      const sheet = idx === 0 ? (defaultSheet.setName(tabName), defaultSheet) : ss.insertSheet(tabName);
      _buildTab(sheet, TAB_STRUCTURE[tabName]);
    });
    const updateResult = _updateClientSheetId(clientKey, ssId);
    if (!updateResult.success) return { success: true, sheetId: ssId, sheetUrl: ssUrl, message: "Sheet created but could not auto-save ID. Copy manually: " + ssId, warning: true };
    return { success: true, sheetId: ssId, sheetUrl: ssUrl, message: "Sheet created and linked successfully!" };
  } catch(e) { return { success: false, error: e.message }; }
}


// ============================================================
// SECTION 7D — RESET EXISTING SHEET
// ============================================================

function serverResetClientSheet(clientKey, currentUserRole, currentUserClients) {
  try {
    if (currentUserRole !== "owner") {
      const allowed = String(currentUserClients).split(",").map(s => s.trim().toLowerCase());
      if (!allowed.includes("all") && !allowed.includes(clientKey.toLowerCase())) return { success: false, error: "Access denied" };
    }
    const result = serverGetClientFull(clientKey);
    if (!result.success) return { success: false, error: result.error };
    const sheetId = result.client.sheet_id;
    if (!sheetId || sheetId.trim() === "") return { success: false, error: "No sheet linked to this client. Use Create Sheet instead." };

    const ss = SpreadsheetApp.openById(sheetId);
    const ssUrl = "https://docs.google.com/spreadsheets/d/" + sheetId;

    const existing = ss.getSheets();
    for (let i = existing.length - 1; i >= 1; i--) ss.deleteSheet(existing[i]);
    const survivor = ss.getSheets()[0];
    survivor.setName("_temp_");

    TAB_ORDER.forEach((tabName, idx) => {
      const sheet = ss.insertSheet(tabName, idx);
      _buildTab(sheet, TAB_STRUCTURE[tabName]);
    });

    ss.deleteSheet(ss.getSheetByName("_temp_"));

    return {
      success:  true,
      sheetUrl: ssUrl,
      message:  "Sheet reset successfully! All 6 tabs rebuilt. ⚠️ All previous data has been cleared.",
    };
  } catch(e) { return { success: false, error: e.message }; }
}


// ============================================================
// SECTION 7E — SHEET URL HELPERS
// ============================================================

function _updateClientSheetId(clientKey, newSheetId) {
  try {
    const sheet = getSystemSheet().getSheetByName("_Clients");
    const data = sheet.getDataRange().getValues(), hdrs = data[0];
    const keyIdx = hdrs.indexOf("client_key"), sheetIdx = hdrs.indexOf("sheet_id"), updIdx = hdrs.indexOf("updated_at");
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][keyIdx]).trim() === clientKey.trim()) {
        sheet.getRange(i+1, sheetIdx+1).setValue(newSheetId);
        if (updIdx >= 0) sheet.getRange(i+1, updIdx+1).setValue(new Date().toISOString());
        return { success: true };
      }
    }
    return { success: false, error: "Client not found: " + clientKey };
  } catch(e) { return { success: false, error: e.message }; }
}

function serverGetClientSheetUrl(clientKey) {
  try {
    const result = serverGetClientFull(clientKey);
    if (!result.success) return { success: false, error: result.error };
    const sheetId = result.client.sheet_id;
    if (!sheetId || sheetId.trim() === "") return { success: false, error: "No sheet linked yet." };
    return { success: true, sheetId, sheetUrl: "https://docs.google.com/spreadsheets/d/" + sheetId };
  } catch(e) { return { success: false, error: e.message }; }
}


// ============================================================
// SECTION C — MONTH CONFIG
// ============================================================

function getMonthConfig(manualMonth) {
  if (manualMonth && manualMonth.trim() !== "") {
    const months = { jan:0,feb:1,mar:2,apr:3,may:4,jun:5,jul:6,aug:7,sep:8,oct:9,nov:10,dec:11,january:0,february:1,march:2,april:3,june:5,july:6,august:7,september:8,october:9,november:10,december:11 };
    const parts = manualMonth.trim().split(/[\s,]+/);
    const rMon = months[parts[0].toLowerCase().slice(0,3)], rYear = parseInt(parts[1]);
    if (rMon === undefined || isNaN(rYear)) throw new Error("Invalid month format: '" + manualMonth + "'. Use 'Jan 2026'");
    const pMon = rMon === 0 ? 11 : rMon - 1, pYear = rMon === 0 ? rYear - 1 : rYear;
    return buildMonthConfig(rYear, rMon, pYear, pMon);
  }
  const now = new Date();
  const rYear = now.getMonth() === 0 ? now.getFullYear() - 1 : now.getFullYear();
  const rMon  = now.getMonth() === 0 ? 11 : now.getMonth() - 1;
  const pYear = rMon === 0 ? rYear - 1 : rYear, pMon = rMon === 0 ? 11 : rMon - 1;
  return buildMonthConfig(rYear, rMon, pYear, pMon);
}

function buildMonthConfig(rYear, rMon, pYear, pMon) {
  const pad = n => n < 10 ? "0" + n : String(n);
  const lastDay = (y,m) => new Date(y, m+1, 0).getDate();
  const SH = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  const LG = ["January","February","March","April","May","June","July","August","September","October","November","December"];
  return {
    currStart:  rYear + "-" + pad(rMon+1) + "-01",
    currEnd:    rYear + "-" + pad(rMon+1) + "-" + pad(lastDay(rYear, rMon)),
    currLabel:  SH[rMon] + " " + rYear,
    currFull:   LG[rMon] + " " + rYear,
    currPeriod: SH[rMon] + " 1 – " + SH[rMon] + " " + lastDay(rYear, rMon) + ", " + rYear,
    prevStart:  pYear + "-" + pad(pMon+1) + "-01",
    prevEnd:    pYear + "-" + pad(pMon+1) + "-" + pad(lastDay(pYear, pMon)),
    prevLabel:  SH[pMon] + " " + pYear,
    nextFull:   LG[rMon === 11 ? 0 : rMon+1] + " " + (rMon === 11 ? rYear+1 : rYear),
  };
}


// ============================================================
// SECTION D — INTERNAL REPORT RUNNER
// ============================================================

function _findExistingReport(clientKey, monthLabel) {
  try {
    // Use _ReportLogs to find existing slide — reliable, no Drive name search
    const sheet = getSystemSheet().getSheetByName("_ReportLogs");
    if (!sheet) return null;
    const data = sheet.getDataRange().getValues(), hdrs = data[0];
    const keyIdx  = hdrs.indexOf("client_key");
    const monIdx  = hdrs.indexOf("month_label");
    const urlIdx  = hdrs.indexOf("slides_url");
    const statIdx = hdrs.indexOf("status");
    // Search from bottom up — most recent first
    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];
      if (String(row[keyIdx]).trim() !== clientKey.trim()) continue;
      if (String(row[monIdx]).trim() !== monthLabel.trim()) continue;
      if (String(row[statIdx]).trim() !== "success") continue;
      const url = String(row[urlIdx] || "").trim();
      if (!url) continue;
      const match = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
      if (!match) continue;
      const fileId = match[1];
      // Skip DriveApp.getFileById — restricted on this account
      // SlidesApp.openById handles missing files gracefully
      Logger.log("  ✓ Found existing report in logs: " + fileId);
      return fileId;
    }
    return null;
  } catch(e) {
    Logger.log("  \u26a0\ufe0f _findExistingReport error: " + e.message);
    return null;
  }
}

function _runReportForClient(CLIENT, M) {
  Logger.log("  Month: " + M.currFull);

  Logger.log("  Reading Sheet...");
  const sheetData = readSheet(M, CLIENT);

  Logger.log("  Fetching Meta (current)...");
  const metaCurr = fetchMeta(M.currStart, M.currEnd, CLIENT);

  Logger.log("  Fetching Meta (previous)...");
  const metaPrev = fetchMeta(M.prevStart, M.prevEnd, CLIENT);

  Logger.log("  Fetching Google Ads (current)...");
  const gAdsCurr = CLIENT.googleAds.enabled ? fetchGoogleAds(M.currStart, M.currEnd, CLIENT) : { campaigns: [], totals: { spend:0, revenue:0, orders:0, impressions:0, clicks:0 } };

  Logger.log("  Fetching Google Ads (previous)...");
  const gAdsPrev = CLIENT.googleAds.enabled ? fetchGoogleAds(M.prevStart, M.prevEnd, CLIENT) : { campaigns: [], totals: { spend:0, revenue:0, orders:0, impressions:0, clicks:0 } };

  Logger.log("  Fetching GA4 (current)...");
  const ga4Curr = fetchGA4(M.currStart, M.currEnd, CLIENT);

  Logger.log("  Fetching GA4 (previous)...");
  const ga4Prev = fetchGA4(M.prevStart, M.prevEnd, CLIENT);

  Logger.log("  Fetching Search Console...");
  let gsc = { topPages: [], topQueries: [] };
  try { gsc = fetchGSC(M.currStart, M.currEnd, CLIENT); } catch(e) { Logger.log("  ⚠️ GSC skipped: " + e.message); }

  Logger.log("  Computing metrics...");
  const computed = buildMetrics(metaCurr, metaPrev, gAdsCurr, gAdsPrev, ga4Curr, ga4Prev, sheetData, M);

  Logger.log("  Running AI analysis...");
  let inferences, recs;
  try {
    inferences = buildAIInferences(computed, sheetData, metaCurr, M, CLIENT.name);
    recs       = buildAIRecommendations(computed, sheetData, metaCurr, M, CLIENT.name);
    Logger.log("  ✓ AI done");
  } catch(e) {
    Logger.log("  ⚠️ AI failed, using rule-based: " + e.message);
    inferences = buildRuleInferences(computed, sheetData, M);
    recs       = buildRuleRecommendations(computed, sheetData, M);
  }

  Logger.log("  Building slides...");
  const reportName = "Digifyce – " + CLIENT.name + " – " + M.currFull + " Report";
  const existingId = _findExistingReport(CLIENT.clientKey, M.currLabel);
  let deck;
  if (existingId) {
    Logger.log("  Updating existing report...");
    deck = SlidesApp.openById(existingId);
    deck.getSlides().forEach(slide => slide.remove());
    // Rename using SlidesApp only — avoids DriveApp.getFileById restriction
    deck.setName(reportName);
  } else {
    Logger.log("  Creating new report...");
    deck = SlidesApp.create(reportName);
  }

  const ctx = { deck, M, CLIENT, metaCurr, metaPrev, gAdsCurr, gAdsPrev, ga4Curr, ga4Prev, gsc, sheetData, computed, inferences, recs };
  buildAllSlides(ctx);

  const slides = deck.getSlides();
  if (!existingId && slides.length > 1) {
    const first = slides[0];
    if (first.getShapes().length === 0 && first.getImages().length === 0) first.remove();
  }

  const fileId = deck.getId();

  // Move to reports folder (only on first creation, not on updates)
  if (!existingId) {
    moveFileToReportsFolder(fileId);
  }

  Logger.log("  ✓ Done: https://docs.google.com/presentation/d/" + fileId);
  return "https://docs.google.com/presentation/d/" + fileId;
}


// ============================================================
// SECTION E — COLOR PALETTE
// ============================================================

const C = {
  navy:"0D2B6E", blue:"1565C0", blueMid:"1976D2", cyan:"29B6F6",
  cyanLight:"B3E5FC", sky:"42A5F5", white:"FFFFFF", offWhite:"F8FAFD",
  blueLight:"E3F2FD", bluePale:"BBDEFB", dark:"0D2B6E", mid:"546E7A",
  muted:"90A4AE", line:"E0E7EF", lineDark:"CFD8DC",
  green:"2E7D32", greenLight:"E8F5E9", greenMid:"43A047",
  red:"C62828", redLight:"FFEBEE", orange:"E65100", orangeLight:"FFF3E0",
  waGreen:"25D366", waLight:"E8F5E9", pushOr:"FB8C00", pushLight:"FFF3E0",
  editFlag:"CC0000",
};


// ============================================================
// SECTION F — SHEET READER
// ============================================================

function readSheet(M, CLIENT) {
  let ss;
  try {
    ss = SpreadsheetApp.openById(CLIENT.sheetId);
  } catch(e) {
    throw new Error("Cannot open client sheet (ID: " + CLIENT.sheetId + "). Make sure the sheet exists and is shared with the script. Error: " + e.message);
  }

  const curr = M.currLabel, prev = M.prevLabel;

  // Row 1=title, 2=note, 3=headers, 4+=data
  function getRows(tabName, monthLabel) {
    const sheet = ss.getSheetByName(tabName);
    if (!sheet) throw new Error("Tab '" + tabName + "' not found in client sheet. Please reset the sheet structure from the dashboard.");
    const data = sheet.getDataRange().getValues();
    if (data.length < 3) return [];
    const headers = data[2];

    // Parse target month label "Mar 2026" → { month: 2, year: 2026 }
    const MONTHS = {jan:0,feb:1,mar:2,apr:3,may:4,jun:5,jul:6,aug:7,sep:8,oct:9,nov:10,dec:11};
    const parts = monthLabel.trim().split(" ");
    const targetMonth = MONTHS[parts[0].toLowerCase().slice(0,3)];
    const targetYear  = parseInt(parts[1]);

    return data.slice(3)
      .filter(r => {
        const v = r[0];
        if (v === null || v === undefined || v === "") return false;
        // Handle Date objects (when cell is formatted as Date in Google Sheets)
        if (v instanceof Date) {
          return v.getFullYear() === targetYear && v.getMonth() === targetMonth;
        }
        // Handle plain text strings like "Mar 2026"
        return String(v).trim() === monthLabel;
      })
      .map(r => { const obj = {}; headers.forEach((h,i) => { obj[String(h).trim()] = r[i]; }); return obj; });
  }

  // num() handles: numbers, "13,92,972", "₹2,28,132.20", percentages like 0.1995
  function num(v) {
    if (v === null || v === undefined || v === "") return 0;
    if (typeof v === "number") return v;
    // Remove ₹, commas, spaces, then parse
    const cleaned = String(v).replace(/[₹,\s]/g, "");
    return parseFloat(cleaned) || 0;
  }
  function rv(row, col) { const v = row[col]; if (v === null || v === undefined || v === "") return "Nil"; return String(v); }

  function readShopifySummary(monthLabel) {
    const rows = getRows("Shopify_Summary", monthLabel);
    if (!rows.length) return { grossSales: 0, totalOrders: 0, newCustomers: 0, aov: 0 };
    const grossSales   = rows.reduce((s,r) => s + num(r["Gross Sales (₹)"]), 0);
    const totalOrders  = rows.reduce((s,r) => s + num(r["Total Orders"]), 0);
    const newCustomers = rows.reduce((s,r) => s + num(r["Total New Customers"]), 0);
    const aov = totalOrders > 0 ? Math.round(grossSales / totalOrders) : 0;
    return { grossSales: Math.round(grossSales), totalOrders: Math.round(totalOrders), newCustomers: Math.round(newCustomers), aov };
  }

  const shopCurr = readShopifySummary(curr);
  const shopPrev = readShopifySummary(prev);

  const shopifyProducts = getRows("Shopify_Products", curr)
    .map(r => ({ name: String(r["Product Name"] || "").trim(), orders: num(r["Net Orders"]), revenue: num(r["Revenue (₹)"]) }))
    .filter(p => p.name && p.revenue > 0).sort((a,b) => b.revenue - a.revenue).slice(0, 8);

  const shopifyLocations = getRows("Shopify_Locations", curr)
    .map(r => ({ location: String(r["Location (City, State)"] || "").trim(), orders: num(r["Orders"]), revenue: num(r["Revenue (₹)"]) }))
    .filter(l => l.location && l.revenue > 0).sort((a,b) => b.revenue - a.revenue).slice(0, 8);

  const retRows = getRows("Retention", curr);
  function retRow(channel) { return retRows.find(r => String(r["Channel"]).toLowerCase() === channel.toLowerCase()) || {}; }

  const retSummCurr = getRows("Retention_Summary", curr)[0] || {};
  const retSummPrev = getRows("Retention_Summary", prev)[0] || {};
  const returningRateCurr = num(retSummCurr["Returning Rate (%)"]);
  const returningRatePrev = num(retSummPrev["Returning Rate (%)"]);
  const retChange = returningRatePrev > 0
    ? (returningRateCurr >= returningRatePrev ? "↑ " : "↓ ") + Math.abs(returningRateCurr - returningRatePrev).toFixed(1) + " pts vs " + prev
    : "";

  const retention = {
    email:    { sent: rv(retRow("Email"),"Sent"), openRate: rv(retRow("Email"),"Open / Read Rate (%)"), clickRate: rv(retRow("Email"),"Click Rate (%)"), revenue: rv(retRow("Email"),"Revenue (₹)"), roas: rv(retRow("Email"),"ROAS"), note: rv(retRow("Email"),"Note") },
    whatsapp: { sent: rv(retRow("WhatsApp"),"Sent"), readRate: rv(retRow("WhatsApp"),"Open / Read Rate (%)"), clickRate: rv(retRow("WhatsApp"),"Click Rate (%)"), revenue: rv(retRow("WhatsApp"),"Revenue (₹)"), roas: rv(retRow("WhatsApp"),"ROAS"), note: rv(retRow("WhatsApp"),"Note") },
    push:     { sent: rv(retRow("Push"),"Sent"), openRate: rv(retRow("Push"),"Open / Read Rate (%)"), clickRate: rv(retRow("Push"),"Click Rate (%)"), revenue: rv(retRow("Push"),"Revenue (₹)"), roas: rv(retRow("Push"),"ROAS"), note: rv(retRow("Push"),"Note") },
    totalRevenue:  rv(retSummCurr,"Total Retention Revenue (₹)"),
    returningRate: returningRateCurr > 0 ? returningRateCurr.toFixed(1) + "%" : "Nil",
    avgRepeatAOV:  rv(retSummCurr,"Avg Repeat AOV (₹)"),
    totalOrders:   rv(retSummCurr,"Total Retention Orders"),
    retChange,
  };

  const stratRows = getRows("Strategy", curr);
  function stratRow(s) { return stratRows.find(r => String(r["Section"]).trim() === s) || {}; }
  function parseApproach(section) {
    const r = stratRow(section);
    if (!r["Title / Label"]) return null;
    const points = [];
    [["Point 1 Heading","Point 1 Body"],["Point 2 Heading","Point 2 Body"],["Point 3 Heading","Point 3 Body"]].forEach(([hk,bk]) => {
      const h = String(r[hk] || "").trim(), b = String(r[bk] || "").trim();
      if (h) points.push({ heading: h, body: b });
    });
    return { title: String(r["Title / Label"]).trim(), points };
  }
  const approaches = ["APPROACH_1","APPROACH_2","APPROACH_3"].map(parseApproach).filter(Boolean);

  const projRow = stratRow("PROJECTION");
  function fmtMoney(n) { n = parseFloat(n) || 0; if (n >= 100000) return "₹" + (n/100000).toFixed(0) + "L"; if (n >= 1000) return "₹" + (n/1000).toFixed(0) + "K"; return "₹" + Math.round(n); }
  const projections = {
    targetSalesRaw:  num(projRow["Point 1 Heading"]),
    targetSales:     fmtMoney(projRow["Point 1 Heading"]),
    targetBudget:    fmtMoney(projRow["Point 2 Heading"]),
    targetCustomers: String(projRow["Point 3 Heading"] || "").trim(),
    targetROAS:      (num(projRow["Point 4 Heading"]) || 0).toFixed(1) + "x",
    growthNote:      String(projRow["Point 5 Heading"] || "").trim(),
    roasNote:        String(projRow["Point 6 Heading"] || "").trim(),
  };

  return {
    shopCurr, shopPrev, retention,
    shopify: { grossSales: shopCurr.grossSales, totalOrders: shopCurr.totalOrders, newCustomers: shopCurr.newCustomers, aov: shopCurr.aov, salesByProduct: shopifyProducts, salesByLocation: shopifyLocations },
    approaches, projections,
  };
}


// ============================================================
// SECTION G — META ADS API
// ============================================================

function fetchMeta(dateStart, dateEnd, CLIENT) {
  const token = CLIENT.meta.accessToken, v = "v19.0";
  const tr = encodeURIComponent(JSON.stringify({ since: dateStart, until: dateEnd }));
  let totSpend = 0, totRev = 0, totOrders = 0, totImpr = 0, totClicks = 0;
  const allCampaigns = [], allCreatives = [];

  CLIENT.meta.adAccountIds.forEach(acct => {
    if (!acct) return;
    try {
      Logger.log("    Meta account: " + acct);
      const cf = encodeURIComponent("campaign_name,spend,impressions,clicks,actions,action_values");
      const cr = safeGet("https://graph.facebook.com/" + v + "/" + acct + "/insights?fields=" + cf + "&level=campaign&time_range=" + tr + "&limit=100&access_token=" + token, "Meta campaigns [" + acct + "]");
      (cr.data || []).forEach(row => {
        const orders = mVal(row.actions,"purchase"), revenue = mVal(row.action_values,"purchase"), spend = parseFloat(row.spend || 0);
        allCampaigns.push({ channel:"META", name: row.campaign_name || "—", objective:"SALES", orders: Math.round(orders), spent: Math.round(spend), sales: Math.round(revenue), roas: spend > 0 ? (revenue/spend).toFixed(1) + "x" : "—", account: acct });
      });

      const af = encodeURIComponent("spend,impressions,clicks,actions,action_values");
      const ar = safeGet("https://graph.facebook.com/" + v + "/" + acct + "/insights?fields=" + af + "&time_range=" + tr + "&access_token=" + token, "Meta totals [" + acct + "]");
      const a = ar.data && ar.data[0] ? ar.data[0] : {};
      totSpend += parseFloat(a.spend || 0);
      totRev   += mVal(a.action_values,"purchase");
      totOrders += mVal(a.actions,"purchase");
      totImpr  += parseInt(a.impressions || 0);
      totClicks += parseInt(a.clicks || 0);

      const adf = encodeURIComponent("ad_name,spend,ctr,actions,action_values");
      const adr = safeGet("https://graph.facebook.com/" + v + "/" + acct + "/insights?fields=" + adf + "&level=ad&limit=50&time_range=" + tr + "&access_token=" + token, "Meta ads [" + acct + "]");
      (adr.data || []).forEach(ad => {
        const o = Math.round(mVal(ad.actions,"purchase")), r = mVal(ad.action_values,"purchase"), s = parseFloat(ad.spend || 0);
        if (o > 0) allCreatives.push({ name: ad.ad_name || "—", ctr: parseFloat(ad.ctr || 0).toFixed(1) + "%", cpa: o > 0 ? "₹" + Math.round(s/o) : "—", orders: o, revenue: r });
      });
    } catch(e) {
      Logger.log("    ⚠️ Meta error for " + acct + ": " + e.message);
    }
  });

  totSpend = Math.round(totSpend); totRev = Math.round(totRev); totOrders = Math.round(totOrders);
  return {
    campaigns:    allCampaigns,
    totals:       { spend: totSpend, revenue: totRev, orders: totOrders, impressions: totImpr, clicks: totClicks },
    topCreatives: allCreatives.sort((a,b) => b.revenue - a.revenue).slice(0, 3)
  };
}

function mVal(arr, type) {
  if (!arr || !Array.isArray(arr)) return 0;
  const m = arr.find(a => a.action_type === type);
  return m ? parseFloat(m.value) : 0;
}


// ============================================================
// SECTION H — GOOGLE ADS API
// ============================================================

function fetchGoogleAds(dateStart, dateEnd, CLIENT) {
  if (!CLIENT.googleAds.enabled) return { campaigns: [], totals: { spend:0, revenue:0, orders:0, impressions:0, clicks:0 } };
  try {
    const custId = CLIENT.googleAds.customerId.replace(/-/g,""), devToken = CLIENT.googleAds.developerToken;
    const mccId  = CLIENT.googleAds.managerCustomerId ? CLIENT.googleAds.managerCustomerId.replace(/-/g,"") : custId;
    const headers = { "Authorization": "Bearer " + ScriptApp.getOAuthToken(), "developer-token": devToken, "login-customer-id": mccId, "Content-Type": "application/json" };
    const query = `SELECT campaign.name,campaign.status,metrics.cost_micros,metrics.conversions,metrics.conversions_value,metrics.impressions,metrics.clicks FROM campaign WHERE segments.date BETWEEN '${dateStart}' AND '${dateEnd}' AND campaign.status = 'ENABLED' ORDER BY metrics.cost_micros DESC LIMIT 50`;
    const resp = UrlFetchApp.fetch("https://googleads.googleapis.com/v19/customers/" + custId + "/googleAds:search", { method:"POST", contentType:"application/json", headers, payload: JSON.stringify({ query }), muteHttpExceptions: true });
    const json = JSON.parse(resp.getContentText());
    if (json.error) throw new Error("Google Ads [" + CLIENT.name + "]: " + json.error.message);
    const results = json.results || [], allCampaigns = [];
    let totSpend = 0, totRev = 0, totOrders = 0, totImpr = 0, totClicks = 0;
    results.forEach(row => {
      const m = row.metrics || {}, spend = parseInt(m.cost_micros || 0) / 1e6, revenue = parseFloat(m.conversions_value || 0), orders = Math.round(parseFloat(m.conversions || 0));
      allCampaigns.push({ channel:"GOOGLE", name: row.campaign?.name || "—", objective:"SALES", orders, spent: Math.round(spend), sales: Math.round(revenue), roas: spend > 0 ? (revenue/spend).toFixed(1) + "x" : "—" });
      totSpend += spend; totRev += revenue; totOrders += orders; totImpr += parseInt(m.impressions || 0); totClicks += parseInt(m.clicks || 0);
    });
    return { campaigns: allCampaigns, totals: { spend: Math.round(totSpend), revenue: Math.round(totRev), orders: totOrders, impressions: totImpr, clicks: totClicks } };
  } catch(e) {
    Logger.log("  ⚠️ Google Ads error: " + e.message);
    return { campaigns: [], totals: { spend:0, revenue:0, orders:0, impressions:0, clicks:0 } };
  }
}


// ============================================================
// SECTION I — GA4 API
// ============================================================

function fetchGA4(dateStart, dateEnd, CLIENT) {
  try {
    const pid = CLIENT.ga4.propertyId;
    const chResp = ga4Report(pid, {
      dateRanges: [{ startDate: dateStart, endDate: dateEnd }],
      dimensions: [{ name: "sessionDefaultChannelGrouping" }],
      metrics: [{ name:"sessions" },{ name:"conversions" },{ name:"totalRevenue" },{ name:"newUsers" }]
    });
    const lmap = { "Organic Search":"SEO / Search", "Direct":"Direct", "Organic Social":"Social Organic", "Referral":"Referral" };
    const channels = {};
    (chResp.rows || []).forEach(row => {
      const lbl = lmap[row.dimensionValues[0].value]; if (!lbl) return;
      const s = parseInt(row.metricValues[0].value) || 0, c = parseInt(row.metricValues[1].value) || 0, r = parseFloat(row.metricValues[2].value) || 0;
      channels[lbl] = { sessions: s, convRate: s > 0 ? parseFloat(((c/s)*100).toFixed(2)) : 0, revenue: Math.round(r) };
    });
    ["SEO / Search","Direct","Social Organic","Referral"].forEach(ch => { if (!channels[ch]) channels[ch] = { sessions:0, convRate:0, revenue:0 }; });
    const ovResp = ga4Report(pid, {
      dateRanges: [{ startDate: dateStart, endDate: dateEnd }],
      metrics: [{ name:"sessions" },{ name:"newUsers" },{ name:"conversions" },{ name:"totalRevenue" }]
    });
    const ov = ovResp.rows && ovResp.rows[0] ? ovResp.rows[0].metricValues : [];
    const s = parseInt((ov[0] || {}).value || 0), c = parseInt((ov[2] || {}).value || 0);
    return { channels, overall: { sessions: s, newUsers: parseInt((ov[1] || {}).value || 0), convRate: s > 0 ? parseFloat(((c/s)*100).toFixed(2)) : 0, revenue: Math.round(parseFloat((ov[3] || {}).value || 0)) } };
  } catch(e) {
    Logger.log("  ⚠️ GA4 error: " + e.message);
    return { channels: { "SEO / Search":{sessions:0,convRate:0,revenue:0}, "Direct":{sessions:0,convRate:0,revenue:0}, "Social Organic":{sessions:0,convRate:0,revenue:0}, "Referral":{sessions:0,convRate:0,revenue:0} }, overall: { sessions:0, newUsers:0, convRate:0, revenue:0 } };
  }
}

function ga4Report(pid, body) {
  const resp = UrlFetchApp.fetch("https://analyticsdata.googleapis.com/v1beta/properties/" + pid + ":runReport", {
    method: "POST", contentType: "application/json",
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
    payload: JSON.stringify(body), muteHttpExceptions: true
  });
  const json = JSON.parse(resp.getContentText());
  if (json.error) throw new Error("GA4: " + json.error.message);
  return json;
}


// ============================================================
// SECTION J — SEARCH CONSOLE API
// ============================================================

function fetchGSC(dateStart, dateEnd, CLIENT) {
  const token = ScriptApp.getOAuthToken();
  function q(dim, lim) {
    const resp = UrlFetchApp.fetch(
      "https://www.googleapis.com/webmasters/v3/sites/" + encodeURIComponent(CLIENT.gsc.siteUrl) + "/searchAnalytics/query",
      { method:"POST", contentType:"application/json", headers:{ Authorization:"Bearer " + token }, payload: JSON.stringify({ startDate: dateStart, endDate: dateEnd, dimensions: [dim], rowLimit: lim }), muteHttpExceptions: true }
    );
    const json = JSON.parse(resp.getContentText());
    if (json.error) throw new Error("GSC: " + json.error.message);
    return json.rows || [];
  }
  return {
    topPages:   q("page",5).map(r => ({ page: r.keys[0], clicks: r.clicks, impressions: r.impressions, ctr: (r.ctr*100).toFixed(1)+"%", position: r.position.toFixed(1) })),
    topQueries: q("query",5).map(r => ({ query: r.keys[0], clicks: r.clicks, impressions: r.impressions, ctr: (r.ctr*100).toFixed(1)+"%", position: r.position.toFixed(1) })),
  };
}


// ============================================================
// SECTION K — COMPUTED METRICS
// ============================================================

function buildMetrics(metaCurr, metaPrev, gAdsCurr, gAdsPrev, ga4Curr, ga4Prev, sheetData, M) {
  const MetaCu = metaCurr.totals, MetaPr = metaPrev.totals;
  const GAcu   = gAdsCurr.totals, GAPr   = gAdsPrev.totals;

  function addTotals(a, b) {
    const spend = (a.spend + b.spend), revenue = (a.revenue + b.revenue), orders = (a.orders + b.orders);
    const impressions = (a.impressions || 0) + (b.impressions || 0), clicks = (a.clicks || 0) + (b.clicks || 0);
    const roas = spend > 0 ? parseFloat((revenue/spend).toFixed(2)) : 0;
    const ctr  = impressions > 0 ? parseFloat(((clicks/impressions)*100).toFixed(2)) : 0;
    const cpc  = clicks > 0 ? parseFloat((spend/clicks).toFixed(1)) : 0;
    const cpa  = orders > 0 ? Math.round(spend/orders) : 0;
    return { spend, revenue, orders, impressions, clicks, roas,
      roasStr: roas > 0 ? roas.toFixed(1) + "x" : "—",
      ctr, ctrStr: ctr > 0 ? ctr.toFixed(2) + "%" : "—",
      cpc, cpcStr: cpc > 0 ? "₹" + cpc.toFixed(1) : "—",
      cpa, cpaStr: cpa > 0 ? "₹" + Math.round(cpa) : "—" };
  }

  const Cu = addTotals(MetaCu, GAcu);
  const Pr = addTotals(MetaPr, GAPr);

  const shopCurr = sheetData.shopCurr, shopPrev = sheetData.shopPrev;
  const currRev    = shopCurr.grossSales   || 0;
  const currOrders = shopCurr.totalOrders  || 0;
  const currCust   = shopCurr.newCustomers || 0;
  const currAOV    = shopCurr.aov || (currOrders > 0 ? Math.round(currRev/currOrders) : 0);
  const currSpend  = Cu.spend;
  const currSess   = ga4Curr.overall.sessions;

  const prevRev    = shopPrev.grossSales   || 0;
  const prevOrders = shopPrev.totalOrders  || 0;
  const prevCust   = shopPrev.newCustomers || 0;
  const prevAOV    = shopPrev.aov || (prevOrders > 0 ? Math.round(prevRev/prevOrders) : 0);
  const prevSpend  = Pr.spend;
  const prevSess   = ga4Prev.overall.sessions;
  const prevCamps  = metaPrev.campaigns.length + (gAdsPrev.campaigns || []).length;
  const currCamps  = metaCurr.campaigns.length + (gAdsCurr.campaigns || []).length;

  const currRoasNum = currSpend > 0 ? parseFloat((currRev/currSpend).toFixed(2)) : 0;
  const prevRoasNum = prevSpend > 0 ? parseFloat((prevRev/prevSpend).toFixed(2)) : 0;
  const currRoasStr = currRoasNum > 0 ? currRoasNum.toFixed(1) + "x" : "—";
  const prevRoasStr = prevRoasNum > 0 ? prevRoasNum.toFixed(1) + "x" : "—";

  function delta(curr, prev, type) {
    if (!prev || prev === 0) return { pct:"—", arrow:"", pos: true, note:"No prev month data" };
    const p = ((curr - prev) / prev) * 100, pos = p >= 0;
    const str = Math.abs(p) < 0.1 ? "0%" : (pos ? "+" : "") + p.toFixed(1) + "%";
    const notes = { revenue: pos ? "Strong overall growth" : "Opportunity to drive higher sales", spend: pos ? "Controlled budget expansion" : "Leaner spend — efficiency focus", roas: pos ? "Major efficiency improvement" : "Optimization opportunity ahead", campaigns: pos ? "Scaling active campaigns" : "Better consolidation & focus", customers: pos ? "Strong acquisition performance" : "Acquisition needs attention", aov: pos ? "Higher order value per customer" : "Opportunity to increase AOV", sessions: pos ? "Traffic growing well" : "Traffic opportunity to explore", orders: pos ? "Strong order volume" : "Conversion focus needed" };
    return { pct: str, arrow: pos ? "↑ " : "↓ ", pos, note: notes[type] || (pos ? "Improving" : "Needs attention") };
  }

  const dRev  = delta(currRev,currRev!==0?prevRev:0,"revenue"),   dSpend = delta(currSpend, prevSpend,"spend");
  const dRoas = delta(currRoasNum, prevRoasNum,"roas"),             dCamps = delta(currCamps, prevCamps,"campaigns");
  const dCust = delta(currCust, prevCust,"customers"),              dAOV   = delta(currAOV, prevAOV,"aov");
  const dSess = delta(currSess, prevSess,"sessions"),               dOrd   = delta(currOrders, prevOrders,"orders");
  const dCTR  = delta(Cu.ctr, Pr.ctr,"roas"),                      dCPC   = delta(Cu.cpc, Pr.cpc,"spend");
  const dCPA  = delta(Cu.cpa, Pr.cpa,"spend"),                     dCR    = delta(ga4Curr.overall.convRate, ga4Prev.overall.convRate,"roas");
  const retChange = sheetData.retention.retChange || "";

  const allPos  = [dRev,dRoas,dCust,dAOV].every(d => d.pos);
  const mostPos = [dRev,dRoas,dCust,dAOV].filter(d => d.pos).length >= 3;
  const banner  = allPos ? "Every single metric improved this month — a strong signal that our strategy is working." : mostPos ? "Strong performance across key metrics this month — with targeted opportunities ahead." : "Mixed results this month — key wins to build on and clear areas to optimize.";

  const fmtR = v => v > 0 ? "₹" + fN(v) : "—";
  const fmtS = v => v > 0 ? fK(v) : "—";
  const fmtN = v => v > 0 ? String(v) : "—";
  const fmtA = v => v > 0 ? "₹" + fN(v) : "—";
  const sfx  = pct => pct !== "—" ? " vs " + M.prevLabel : "";

  return {
    grossSales: "₹" + fN(currRev), grossSalesNum: currRev,
    totalSpend: fK(currSpend),     totalSpendNum: currSpend,
    roas: currRoasStr, roasNum: currRoasNum,
    newCustomers: currCust > 0 ? String(currCust) : "—", newCustomersNum: currCust,
    sessions: fN(currSess), totalOrders: currOrders,
    aov: currAOV > 0 ? "₹" + fN(currAOV) : "—", aovNum: currAOV,
    ctr: Cu.ctrStr, cpc: Cu.cpcStr, cpa: Cu.cpaStr,
    convRate: ga4Curr.overall.convRate.toFixed(1) + "%",
    returningRate: sheetData.retention.returningRate,
    revChange:   dRev.arrow  + dRev.pct.replace("+","")  + sfx(dRev.pct),
    custChange:  dCust.arrow + dCust.pct.replace("+","") + sfx(dCust.pct),
    sessChange:  dSess.arrow + dSess.pct.replace("+","") + sfx(dSess.pct),
    ordChange:   dOrd.arrow  + dOrd.pct.replace("+","")  + sfx(dOrd.pct),
    aovChange:   dAOV.arrow  + dAOV.pct.replace("+","")  + sfx(dAOV.pct),
    roasChange:  dRoas.arrow + dRoas.pct.replace("+","") + sfx(dRoas.pct),
    spendChange: dSpend.arrow + dSpend.pct.replace("+","") + sfx(dSpend.pct),
    ctrChange:   dCTR.arrow  + dCTR.pct.replace("+","")  + sfx(dCTR.pct),
    cpcChange:   dCPC.arrow  + dCPC.pct.replace("+","")  + sfx(dCPC.pct),
    cpaChange:   dCPA.arrow  + dCPA.pct.replace("+","")  + sfx(dCPA.pct),
    crChange:    dCR.arrow   + dCR.pct.replace("+","")   + sfx(dCR.pct),
    retChange,
    comp: {
      revenue:   { prev: fmtR(prevRev),   curr: "₹" + fN(currRev),  ...dRev },
      spend:     { prev: fmtS(prevSpend),  curr: fK(currSpend),      ...dSpend },
      roas:      { prev: prevRoasStr,      curr: currRoasStr,         ...dRoas },
      campaigns: { prev: fmtN(prevCamps),  curr: String(currCamps),  ...dCamps },
      customers: { prev: fmtN(prevCust),   curr: fmtN(currCust),     ...dCust },
      aov:       { prev: fmtA(prevAOV),    curr: fmtA(currAOV),      ...dAOV },
      orders:    { prev: fmtN(prevOrders), curr: fmtN(currOrders),   ...dOrd },
    },
    banner, prevRev, prevSpend, prevRoasNum,
  };
}


// ============================================================
// SECTION L — AI INFERENCES
// ============================================================

function buildAIInferences(computed, sheetData, metaCurr, M, brandName) {
  const topCampaigns = metaCurr.campaigns.slice(0,3).map(c => c.name + " (ROAS:" + c.roas + ", Spend:₹" + fN(c.spent) + ")").join(" | ");
  const topProducts  = sheetData.shopify.salesByProduct.slice(0,3).map(p => p.name + " ₹" + fN(p.revenue)).join(" | ");
  const topLocations = sheetData.shopify.salesByLocation.slice(0,3).map(l => l.location + " ₹" + fN(l.revenue)).join(" | ");
  const prompt = `You are a senior performance marketing analyst writing a monthly report for ${brandName}, an Indian D2C brand. Write exactly 4 sharp inferences based on the data below.\n\nPERFORMANCE DATA — ${M.currFull}:\n- Gross Sales: ${computed.grossSales} (${computed.revChange})\n- Ad Spend: ${computed.totalSpend} (${computed.spendChange})\n- ROAS: ${computed.roas} (${computed.roasChange})\n- Orders: ${computed.totalOrders} (${computed.ordChange})\n- New Customers: ${computed.newCustomers} (${computed.custChange})\n- AOV: ${computed.aov} (${computed.aovChange})\n- Returning Rate: ${computed.returningRate} (${computed.retChange})\n- CTR: ${computed.ctr} (${computed.ctrChange})\n- CPC: ${computed.cpc} (${computed.cpcChange})\n- CPA: ${computed.cpa} (${computed.cpaChange})\n- Conv Rate: ${computed.convRate} (${computed.crChange})\n- Sessions: ${computed.sessions} (${computed.sessChange})\n- Top Campaigns: ${topCampaigns||"N/A"}\n- Top Products: ${topProducts||"N/A"}\n- Top Locations: ${topLocations||"N/A"}\n- Email ROAS: ${sheetData.retention.email.roas} | WA ROAS: ${sheetData.retention.whatsapp.roas}\n\nSTRICT RULES:\n- Each inference must reference specific numbers from the data\n- Title: max 8 words, punchy, direct\n- Body: max 2 sentences, max 160 characters total, no fluff\n- Agency tone — direct, confident, consultant-style\n- No generic statements. Every sentence must be data-backed\n- Respond ONLY with valid JSON, no markdown:\n{"inferences":[{"title":"...","body":"..."},{"title":"...","body":"..."},{"title":"...","body":"..."},{"title":"...","body":"..."}]}`;

  const resp = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", {
    method: "POST",
    headers: { "Authorization": "Bearer " + OPENAI_API_KEY, "Content-Type": "application/json" },
    payload: JSON.stringify({ model: OPENAI_MODEL, temperature: 0.4, max_tokens: 600, messages: [{ role:"user", content: prompt }] }),
    muteHttpExceptions: true
  });
  const json = JSON.parse(resp.getContentText());
  if (json.error) throw new Error("OpenAI: " + json.error.message);

  let parsed;
  try {
    const raw = json.choices[0].message.content.trim().replace(/```json|```/g,"").trim();
    parsed = JSON.parse(raw);
  } catch(e) {
    throw new Error("OpenAI response parse error: " + e.message);
  }

  const colors = [C.blue, C.navy, "00695C", "4527A0"];
  const bgs    = [C.blueLight, C.blueLight, "E0F2F1", "EDE7F6"];
  return parsed.inferences.slice(0,4).map((inf,i) => ({ num: String(i+1), color: colors[i] || C.blue, bg: bgs[i] || C.blueLight, title: inf.title || " ", body: inf.body || " ", isEdit: false }));
}


// ============================================================
// SECTION M — AI RECOMMENDATIONS
// ============================================================

function buildAIRecommendations(computed, sheetData, metaCurr, M, brandName) {
  const topCampaigns = metaCurr.campaigns.slice(0,3).map(c => c.name + " (ROAS:" + c.roas + ")").join(" | ");
  const topProducts  = sheetData.shopify.salesByProduct.slice(0,3).map(p => p.name + " ₹" + fN(p.revenue)).join(" | ");
  const prompt = `You are a senior media buyer and growth strategist writing next month's action plan for ${brandName}, an Indian D2C brand. Write exactly 4 recommendations for ${M.nextFull} based on ${M.currFull} performance.\n\nPERFORMANCE DATA — ${M.currFull}:\n- ROAS: ${computed.roas} (${computed.roasChange})\n- CPA: ${computed.cpa} (${computed.cpaChange})\n- CTR: ${computed.ctr} (${computed.ctrChange})\n- Conv Rate: ${computed.convRate} (${computed.crChange})\n- AOV: ${computed.aov} (${computed.aovChange})\n- New Customers: ${computed.newCustomers} (${computed.custChange})\n- Returning Rate: ${computed.returningRate} (${computed.retChange})\n- Ad Spend: ${computed.totalSpend} (${computed.spendChange})\n- Top Campaigns: ${topCampaigns||"N/A"}\n- Top Products: ${topProducts||"N/A"}\n- Email ROAS: ${sheetData.retention.email.roas} | WA ROAS: ${sheetData.retention.whatsapp.roas}\n- Top Locations: ${sheetData.shopify.salesByLocation.slice(0,3).map(l=>l.location).join(", ")||"N/A"}\n\nSTRICT RULES:\n- Title: max 8 words, action-oriented\n- Exactly 2 bullet points per recommendation\n- Each bullet: max 90 characters, specific action with data reference\n- Direct agency language — tell them exactly what to do and why\n- Cover: budget/scaling, creative, CRO/landing page, retention\n- Respond ONLY with valid JSON, no markdown:\n{"recommendations":[{"title":"...","bullets":["...","..."]},{"title":"...","bullets":["...","..."]},{"title":"...","bullets":["...","..."]},{"title":"...","bullets":["...","..."]}]}`;

  const resp = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", {
    method: "POST",
    headers: { "Authorization": "Bearer " + OPENAI_API_KEY, "Content-Type": "application/json" },
    payload: JSON.stringify({ model: OPENAI_MODEL, temperature: 0.4, max_tokens: 700, messages: [{ role:"user", content: prompt }] }),
    muteHttpExceptions: true
  });
  const json = JSON.parse(resp.getContentText());
  if (json.error) throw new Error("OpenAI: " + json.error.message);

  let parsed;
  try {
    const raw = json.choices[0].message.content.trim().replace(/```json|```/g,"").trim();
    parsed = JSON.parse(raw);
  } catch(e) {
    throw new Error("OpenAI response parse error: " + e.message);
  }

  const colors = [C.blue, "00897B", "4527A0", C.navy];
  const bgs    = [C.blueLight, "E0F2F1", "EDE7F6", C.blueLight];
  return parsed.recommendations.slice(0,4).map((rec,i) => ({ num: String(i+1).padStart(2,"0"), color: colors[i] || C.blue, bg: bgs[i] || C.blueLight, title: rec.title || " ", bullets: (rec.bullets || [" "," "]).slice(0,2).map(b => b || " "), isEdit: false }));
}


// ============================================================
// RULE-BASED FALLBACK
// ============================================================

function buildRuleInferences(computed, sheetData, M) {
  const comp = computed.comp;
  return [
    { num:"1", color:C.blue,    bg:C.blueLight, isEdit:false, title: comp.roas.pos ? "ROAS Improved — Strategy Working" : "ROAS Under Pressure",          body: comp.roas.pos ? "Campaigns consolidated from " + comp.campaigns.prev + " to " + comp.campaigns.curr + " drove ROAS from " + comp.roas.prev + " to " + comp.roas.curr + "." : "ROAS moved from " + comp.roas.prev + " to " + comp.roas.curr + " — creative mix and targeting need review." },
    { num:"2", color:C.navy,    bg:C.blueLight, isEdit:false, title: comp.customers.pos ? "Acquisition Strategy Delivered" : "Acquisition Needs Attention", body: comp.customers.pos ? "New customers grew " + comp.customers.pct + " (" + comp.customers.prev + " → " + comp.customers.curr + "). Creative and audience resonated." : "New customers moved from " + comp.customers.prev + " to " + comp.customers.curr + ". Review targeting and creative angles." },
    { num:"3", color:"00695C",  bg:"E0F2F1",    isEdit:false, title: "Spend vs Revenue Efficiency",                                                         body: "Revenue " + comp.revenue.pct + " vs spend " + comp.spend.pct + " — " + (comp.revenue.pos && !comp.spend.pos ? "growth was organic, not spend-driven." : "review spend efficiency for next month.") },
    { num:"4", color:"4527A0",  bg:"EDE7F6",    isEdit:false, title: computed.returningRate !== "Nil" ? "Retention Rate: " + computed.returningRate : "Retention Data Not Available", body: computed.retChange ? "Returning customer rate " + computed.retChange + ". AOV at " + computed.aov + "." : "Fill retention data in the sheet for detailed analysis." },
  ];
}

function buildRuleRecommendations(computed, sheetData, M) {
  const comp = computed.comp;
  return [
    { num:"01", color:C.blue,   bg:C.blueLight, isEdit:false, title: comp.roas.pos ? "Scale Budget on Winning Creatives" : "Restructure Budget First",     bullets: comp.roas.pos ? ["ROAS at " + computed.roas + " — increase budget 10–15% only on top-performing creatives.", "Pause underperformers. Consolidate spend around what's working."] : ["ROAS at " + computed.roas + " — pause low-performers before scaling any budget.", "Reallocate to creatives with strongest conversion signals."] },
    { num:"02", color:"00897B", bg:"E0F2F1",    isEdit:false, title: "Strengthen Creative Hook Strategy",                                                   bullets: ["Add problem-first hooks at the start of ad creatives to capture attention immediately.", "Include direct CTA overlays and testimonial snippets to build trust and drive action."] },
    { num:"03", color:"4527A0", bg:"EDE7F6",    isEdit:false, title: "Align Landing Page with Ad Messaging",                                                bullets: ["Ensure landing page mirrors the visual style and value proposition shown in ads.", "Add social proof and clear CTA above the fold to reduce drop-off."] },
    { num:"04", color:C.navy,   bg:C.blueLight, isEdit:false, title: "Push Retention Channels Harder",                                                      bullets: ["Email and WhatsApp ROAS often outperform paid — increase broadcast frequency.", "Segment repeat buyers and target with exclusive offers to lift returning rate."] },
  ];
}


// ============================================================
// SECTION N — SLIDE ENGINE
// ============================================================

const SW = 720, SH = 540, PAD = 12;

function newSlide(ctx) { const s = ctx.deck.appendSlide(SlidesApp.PredefinedLayout.BLANK); s.getBackground().setSolidFill("#FFFFFF"); return s; }

function bx(s, x, y, w, h, hex) {
  if (!hex) return null;
  x = Math.max(0, x); y = Math.max(0, y); w = Math.min(w, SW-x); h = Math.min(h, SH-y);
  if (w <= 0 || h <= 0) return null;
  const sh = s.insertShape(SlidesApp.ShapeType.RECTANGLE, x, y, w, h);
  sh.getFill().setSolidFill("#" + hex.replace(/^#+/,"").slice(0,6)); sh.getBorder().setTransparent(); return sh;
}

function tx(s, text, x, y, w, h, size, hex, bold, align) {
  x = Math.max(0, x); y = Math.max(0, y); w = Math.min(w, SW-x); h = Math.min(h, SH-y);
  if (w <= 0 || h <= 0) return null;
  let t = String(text == null ? "" : text).trim(); if (t === "") t = " ";
  const box = s.insertTextBox(t, x, y, w, h);
  const st = box.getText().getTextStyle();
  st.setFontSize(size || 10); st.setForegroundColor("#" + ((hex || C.dark)).replace(/^#+/,"").slice(0,6)); st.setBold(bold === true);
  st.setFontFamily(bold === true ? "Montserrat" : "Open Sans");
  box.getText().getParagraphStyle().setParagraphAlignment(align === "center" ? SlidesApp.ParagraphAlignment.CENTER : align === "right" ? SlidesApp.ParagraphAlignment.END : SlidesApp.ParagraphAlignment.START);
  box.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE); box.getFill().setTransparent(); box.getBorder().setTransparent(); return box;
}

function changeBadge(s, x, y, w, h, changeStr, isPositive) {
  if (!changeStr || changeStr.trim() === "" || changeStr.trim() === "—") return;
  const bg  = isPositive === false ? C.redLight  : isPositive === true ? C.greenLight : C.blueLight;
  const clr = isPositive === false ? C.red       : isPositive === true ? C.green      : C.blue;
  bx(s, x, y, w, h, bg); bx(s, x, y, 2, h, clr);
  const clean = changeStr.replace("↑ ","").replace("↓ ","");
  const arrow = isPositive === false ? "↓ " : isPositive === true ? "↑ " : "";
  tx(s, arrow + clean, x+4, y, w-4, h, 7.5, clr, true, "left");
}

function hdr(s, title, chipText, chipBg) {
  bx(s, 0, 0, SW, 46, C.navy); bx(s, 0, 44, SW, 2, C.cyan);
  tx(s, title, 16, 0, 490, 46, 11, C.white, true, "left");
  if (chipText) { const cw=140, ch=26, cx=SW-cw-12, cy=10; bx(s,cx,cy,cw,ch,chipBg||C.blue); bx(s,cx,cy,cw,2,C.cyan); tx(s,chipText,cx,cy,cw,ch,8,C.white,true,"center"); }
}

function ftr(s) { bx(s, 0, 525, SW, 1, C.line); bx(s, 0, 526, SW, 14, "F8FAFD"); tx(s, "Confidential @ Digifyce", 0, 526, SW, 14, 7, C.muted, false, "center"); }

function statCard(s, x, y, w, h, label, value, change, accentHex, isPositive) {
  const acc = accentHex || C.blue;
  bx(s,x,y,w,h,C.white); bx(s,x,y,w,1,C.line); bx(s,x,y+h-1,w,1,C.line); bx(s,x,y,1,h,C.line); bx(s,x+w-1,y,1,h,C.line);
  bx(s,x,y,w,4,acc); bx(s,x,y+4,3,h-4,acc);
  tx(s,label,x+10,y+8,w-14,13,7.5,C.muted,false,"left");
  tx(s,value,x+10,y+22,w-14,30,17,C.dark,true,"left");
  if (change && change.trim() !== "" && change.trim() !== "—") changeBadge(s,x+10,y+h-18,w-20,14,change,isPositive);
}

function tblFixed(s, x, y, cols, rows, hh, rh, accentHex, totRow) {
  const acc = accentHex || C.navy;
  const maxW = SW - x - PAD;
  let totalW = cols.reduce((s,c) => s + c.w, 0);
  if (totalW > maxW) { const sc = maxW/totalW; cols = cols.map(c => ({...c, w: Math.floor(c.w*sc)})); }
  let cx = x; const tw = cols.reduce((s,c) => s + c.w, 0);
  bx(s, x, y, tw, hh, acc);
  cols.forEach(c => { tx(s,c.label,cx+8,y+2,c.w-10,hh-4,8,C.white,true,"left"); if (cx+c.w < x+tw) bx(s,cx+c.w-1,y+4,1,hh-8,"1976D2"); cx += c.w; });
  rows.forEach((row, ri) => {
    cx = x; const ry = y+hh+ri*rh; if (ry+rh > SH-16) return;
    const bg = ri%2===0 ? C.white : C.blueLight;
    bx(s,x,ry,tw,rh,bg); bx(s,x,ry+rh-1,tw,1,C.line);
    cols.forEach((c,ci) => { let val = String(row[ci] || "—"); if (val.length>34&&c.w<180) val = val.slice(0,32)+"…"; tx(s,val,cx+8,ry+3,c.w-10,rh-4,8,C.dark,false,"left"); if (cx+c.w<x+tw) bx(s,cx+c.w-1,ry,1,rh,C.line); cx += c.w; });
  });
  if (totRow) { cx = x; const ty = y+hh+rows.length*rh; if (ty+hh <= SH-16) { bx(s,x,ty,tw,hh,C.blueLight); bx(s,x,ty,tw,2,acc); cols.forEach((c,ci) => { tx(s,String(totRow[ci]||""),cx+8,ty+2,c.w-10,hh-4,8.5,acc,true,"left"); cx += c.w; }); } }
}

function buildAllSlides(ctx) {
  slide_Cover(ctx); slide_HighLevel(ctx); slide_Comparison(ctx);
  slide_CampaignTable(ctx, ctx.metaCurr, "META ADS");
  if (ctx.CLIENT.googleAds.enabled) slide_CampaignTable(ctx, ctx.gAdsCurr, "GOOGLE ADS");
  slide_OrganicChannels(ctx); slide_SEOPages(ctx); slide_SEOQueries(ctx); slide_BestCreatives(ctx);
  if (ctx.sheetData.approaches.length > 0) { slide_ApproachesIntro(ctx); ctx.sheetData.approaches.forEach((ap,i) => slide_Approach(ctx, i+1, ap)); }
  slide_SalesByProduct(ctx); slide_SalesByLocation(ctx); slide_Retention(ctx);
  slide_Inferences(ctx); slide_Recommendations(ctx); slide_Projections(ctx); slide_ThankYou(ctx);
}

function slide_Cover(ctx) {
  const s = newSlide(ctx), M = ctx.M, CLIENT = ctx.CLIENT;
  bx(s,0,0,310,SH,C.navy); bx(s,0,0,310,4,C.cyan);
  const c1 = s.insertShape(SlidesApp.ShapeType.ELLIPSE,-80,300,320,320); c1.getFill().setSolidFill("#1565C0"); c1.getBorder().setTransparent(); c1.sendToBack();
  const c2 = s.insertShape(SlidesApp.ShapeType.ELLIPSE,60,-80,240,240);  c2.getFill().setSolidFill("#0D3080"); c2.getBorder().setTransparent(); c2.sendToBack();
  insertLogo(s,CLIENT.digifyceLogoFileId,22,22,140,36,C.navy,"DIGIFYCE");
  bx(s,22,76,120,20,C.blue); bx(s,22,76,120,2,C.cyan); tx(s,"MONTHLY REPORT",22,76,120,20,6.5,C.white,true,"center");
  tx(s,"MARKETING",   22,108,274,44,28,C.white,   true,"left");
  tx(s,"PERFORMANCE", 22,154,274,44,28,C.cyan,    true,"left");
  tx(s,"ANALYSIS",    22,200,274,36,18,C.bluePale,false,"left");
  bx(s,22,248,180,2,C.cyan);
  tx(s,M.currFull.toUpperCase(),22,258,274,18,9,C.bluePale,false,"left");
  tx(s,M.currPeriod,            22,280,274,16,7.5,C.muted,  false,"left");
  bx(s,310,0,2,SH,C.cyan); bx(s,312,0,408,SH,C.white);
  tx(s,"PREPARED FOR",330,48,200,14,7.5,C.muted,true,"left"); bx(s,330,64,52,2,C.cyan);
  tx(s,CLIENT.name,330,74,378,60,24,C.navy,true,"left"); bx(s,330,142,360,1,C.line);
  [["Report Period",M.currPeriod],["Type","Sales & Marketing Analysis"],["By","Digifyce"]].forEach(([l,v],i) => { tx(s,l,330,154+i*40,100,14,8,C.muted,false,"left"); tx(s,v,436,154+i*40,270,14,9,C.dark,true,"left"); });
  bx(s,330,290,378,148,"F8FAFD"); bx(s,330,290,378,3,C.navy); bx(s,330,290,3,148,C.cyan);
  insertLogo(s,CLIENT.clientLogoFileId,366,300,306,128,C.line,CLIENT.name);
  ftr(s);
}

function slide_HighLevel(ctx) {
  const s = newSlide(ctx), M = ctx.M, cm = ctx.computed;
  hdr(s,"HIGH LEVEL OVERVIEW",M.currFull,C.blue);
  const CW=173, CH=80, GAP=3, startX=PAD, startY=52;
  [{l:"Gross Sales",v:cm.grossSales,ch:cm.revChange,acc:C.navy,pos:cm.comp.revenue.pos},{l:"Total Ad Spend",v:cm.totalSpend,ch:cm.spendChange,acc:C.red,pos:cm.comp.spend.pos},{l:"Overall ROAS",v:cm.roas,ch:cm.roasChange,acc:C.blue,pos:cm.comp.roas.pos},{l:"New Customers",v:cm.newCustomers,ch:cm.custChange,acc:C.cyan,pos:cm.comp.customers.pos}]
    .forEach((c,i) => statCard(s,startX+i*(CW+GAP),startY,CW,CH,c.l,c.v,c.ch,c.acc,c.pos));
  [{l:"Website Sessions",v:cm.sessions,ch:cm.sessChange,acc:C.blue,pos:null},{l:"Total Orders",v:String(cm.totalOrders),ch:cm.ordChange,acc:C.navy,pos:cm.comp.orders.pos},{l:"Avg Order Value",v:cm.aov,ch:cm.aovChange,acc:C.blue,pos:cm.comp.aov.pos},{l:"Returning Rate",v:cm.returningRate,ch:cm.retChange,acc:C.cyan,pos:null}]
    .forEach((c,i) => statCard(s,startX+i*(CW+GAP),startY+CH+GAP,CW,CH,c.l,c.v,c.ch,c.acc,c.pos));
  [{l:"CTR (Combined)",v:cm.ctr,ch:cm.ctrChange,acc:"42A5F5",pos:null},{l:"Conv Rate",v:cm.convRate,ch:cm.crChange,acc:"42A5F5",pos:null},{l:"CPC (Combined)",v:cm.cpc,ch:cm.cpcChange,acc:"42A5F5",pos:null},{l:"CPA (Combined)",v:cm.cpa,ch:cm.cpaChange,acc:"42A5F5",pos:null}]
    .forEach((c,i) => statCard(s,startX+i*(CW+GAP),startY+(CH+GAP)*2,CW,CH,c.l,c.v,c.ch,c.acc,c.pos));
  tx(s,"Sources: Shopify (Sales) · Meta Ads (Spend/CTR/CPC/CPA) · GA4 (Sessions/CR) · " + M.currPeriod,PAD,startY+(CH+GAP)*3+4,SW-PAD*2,12,7,C.muted,false,"center");
  ftr(s);
}

function slide_Comparison(ctx) {
  const s = newSlide(ctx), M = ctx.M, cm = ctx.computed;
  hdr(s,"MOM COMPARISON",M.prevLabel + " → " + M.currLabel,C.blue);
  bx(s,PAD,50,SW-PAD*2,22,C.blueLight); bx(s,PAD,50,3,22,C.cyan);
  tx(s,cm.banner,PAD+8,50,SW-PAD*2-10,22,8.5,C.navy,false,"left");
  const comp = cm.comp;
  tblFixed(s,PAD,76,[{label:"Metric",w:180},{label:M.prevLabel.toUpperCase(),w:90},{label:M.currLabel.toUpperCase(),w:90},{label:"Change",w:70},{label:"Notes",w:210}],[
    ["Total Revenue (₹)",  comp.revenue.prev,   comp.revenue.curr,   comp.revenue.pct,   comp.revenue.note],
    ["Total Ad Spend",      comp.spend.prev,     comp.spend.curr,     comp.spend.pct,     comp.spend.note],
    ["Overall ROAS",        comp.roas.prev,      comp.roas.curr,      comp.roas.pct,      comp.roas.note],
    ["Total Orders",        comp.orders.prev,    comp.orders.curr,    comp.orders.pct,    comp.orders.note],
    ["New Customers",       comp.customers.prev, comp.customers.curr, comp.customers.pct, comp.customers.note],
    ["Avg Order Value (₹)", comp.aov.prev,       comp.aov.curr,       comp.aov.pct,       comp.aov.note],
    ["Active Campaigns",    comp.campaigns.prev, comp.campaigns.curr, comp.campaigns.pct, comp.campaigns.note],
  ],28,34,C.navy);
  tx(s,"ROAS = Shopify Revenue ÷ Ad Spend  |  '—' = no sheet data for " + M.prevLabel,PAD,SH-28,SW-PAD*2,14,7,C.muted,false,"center");
  ftr(s);
}

function slide_CampaignTable(ctx, adData, channelLabel) {
  const s = newSlide(ctx);
  hdr(s,"PAID ADS — CAMPAIGN OVERVIEW",channelLabel,C.blue);
  const allCamps = (adData.campaigns || []).slice(0, 6);
  const tot = (adData.campaigns || []).reduce((a,c) => ({ orders: a.orders+c.orders, spent: a.spent+c.spent, sales: a.sales+c.sales }), { orders:0, spent:0, sales:0 });
  const totRoas = tot.spent > 0 ? (tot.sales/tot.spent).toFixed(1) + "x" : "—";
  const totRoasNum = tot.spent > 0 ? tot.sales/tot.spent : 0;
  const cols = [{label:"Campaign Name",w:254},{label:"Orders",w:56},{label:"Spent (₹)",w:84},{label:"Sales (₹)",w:84},{label:"ROAS",w:56},{label:"Channel",w:56}];
  const rows = allCamps.map(c => [c.name, String(c.orders), fN(c.spent), fN(c.sales), c.roas, c.channel]);
  tblFixed(s,PAD,52,cols,rows,28,27,C.navy,["TOTAL",String(tot.orders),fN(tot.spent),fN(tot.sales),totRoas,""]);
  const cardY = Math.min(52+28+rows.length*27+28+8, 400);
  bx(s,506,cardY,200,110,C.blueLight); bx(s,506,cardY,200,3,C.navy); bx(s,506,cardY,3,110,C.cyan);
  bx(s,516,cardY+10,108,20,C.navy); bx(s,516,cardY+10,108,2,C.cyan);
  tx(s,"Overall ROAS",516,cardY+10,108,20,7.5,C.white,true,"center");
  const arrow = totRoasNum >= 2 ? "↗" : "↘";
  const roasColor = totRoasNum >= 2 ? C.green : C.red;
  tx(s,totRoas,516,cardY+32,130,40,26,C.navy,true,"left");
  tx(s,arrow,648,cardY+36,42,32,20,roasColor,true,"left");
  tx(s,"Spend: " + fK((adData.totals||{}).spend||0),516,cardY+78,180,16,8,C.mid,false,"left");
  ftr(s);
}

function slide_OrganicChannels(ctx) {
  const s = newSlide(ctx);
  hdr(s,"ORGANIC ACQUISITION BREAKDOWN","",C.blue);
  const chs = [{key:"SEO / Search",acc:C.navy,bg:"EDE7F6",label:"SEO / Search"},{key:"Direct",acc:C.blue,bg:C.blueLight,label:"Direct"},{key:"Social Organic",acc:C.cyan,bg:"B3E5FC",label:"Social Organic"}];
  const CARD_W=226, CARD_H=452, GAP=5;
  chs.forEach((ch,i) => {
    const d = ctx.ga4Curr.channels[ch.key] || { sessions:0, convRate:0, revenue:0 };
    const x = PAD + i*(CARD_W+GAP);
    bx(s,x,50,CARD_W,CARD_H,ch.bg); bx(s,x,50,CARD_W,4,ch.acc);
    bx(s,x,50,1,CARD_H,C.line); bx(s,x+CARD_W-1,50,1,CARD_H,C.line); bx(s,x,50+CARD_H-1,CARD_W,1,C.line);
    bx(s,x+12,62,CARD_W-24,22,ch.acc); bx(s,x+12,62,CARD_W-24,2,C.cyan);
    tx(s,ch.label.toUpperCase(),x+12,62,CARD_W-24,22,8,C.white,true,"center");
    bx(s,x+12,88,CARD_W-24,1,C.line);
    [["Sessions",fN(d.sessions)],["Conv Rate",d.convRate.toFixed(2)+"%"],["Revenue","₹"+fN(d.revenue)]].forEach(([lbl,val],j) => {
      const gy = 96+j*112;
      bx(s,x+12,gy,CARD_W-24,106,C.white); bx(s,x+12,gy,CARD_W-24,3,ch.acc);
      tx(s,lbl,x+20,gy+8,CARD_W-40,14,8,C.muted,false,"left");
      tx(s,val,x+20,gy+26,CARD_W-40,40,22,ch.acc,true,"left");
    });
  });
  ftr(s);
}

function slide_SEOPages(ctx) {
  const s = newSlide(ctx);
  hdr(s,"SEO — TOP PERFORMING PAGES",ctx.M.currFull,C.navy);
  const rows = (ctx.gsc.topPages || []).slice(0,8).map(r => [r.page, String(r.clicks), String(r.impressions), r.ctr, r.position]);
  if (!rows.length) { bx(s,PAD,70,SW-PAD*2,60,C.blueLight); tx(s,"No Search Console data available for this period.",PAD+8,70,SW-PAD*2,60,10,C.muted,false,"center"); }
  else tblFixed(s,PAD,52,[{label:"Page URL",w:312},{label:"Clicks",w:78},{label:"Impressions",w:92},{label:"CTR",w:72},{label:"Position",w:72}],rows,28,28,C.navy);
  ftr(s);
}

function slide_SEOQueries(ctx) {
  const s = newSlide(ctx);
  hdr(s,"SEO — TOP SEARCH QUERIES",ctx.M.currFull,C.navy);
  const rows = (ctx.gsc.topQueries || []).slice(0,8).map(r => [r.query, String(r.clicks), String(r.impressions), r.ctr, r.position]);
  if (!rows.length) { bx(s,PAD,70,SW-PAD*2,60,C.blueLight); tx(s,"No Search Console data available for this period.",PAD+8,70,SW-PAD*2,60,10,C.muted,false,"center"); }
  else tblFixed(s,PAD,52,[{label:"Query",w:312},{label:"Clicks",w:78},{label:"Impressions",w:92},{label:"CTR",w:72},{label:"Position",w:72}],rows,28,28,C.navy);
  ftr(s);
}

function slide_BestCreatives(ctx) {
  const s = newSlide(ctx), mc = ctx.metaCurr;
  hdr(s,"BEST PERFORMING CREATIVES","",C.blue);
  tx(s,"Top ad creatives ranked by revenue — replace placeholders with screenshots from Ads Manager",PAD,50,SW-PAD*2,14,8,C.muted,false,"left");
  const accs=[ C.navy,C.blue,C.cyan], bgs=[C.blueLight,"EDE7F6","B3E5FC"], ranks=["#1  TOP PERFORMER","#2  RUNNER UP","#3  RISING STAR"];
  for (let i = 0; i < 3; i++) {
    const cr = mc.topCreatives[i] || { name:"—", ctr:"—", cpa:"—", orders:0 };
    const x = PAD+i*232, CARD_W=226;
    bx(s,x,66,CARD_W,452,bgs[i]); bx(s,x,66,CARD_W,4,accs[i]);
    bx(s,x,66,1,452,C.line); bx(s,x+CARD_W-1,66,1,452,C.line); bx(s,x,66+452-1,CARD_W,1,C.line);
    bx(s,x+8,78,CARD_W-16,22,accs[i]); bx(s,x+8,78,CARD_W-16,2,C.cyan);
    tx(s,ranks[i],x+8,78,CARD_W-16,22,7.5,C.white,true,"left");
    bx(s,x+8,108,CARD_W-16,168,C.white); bx(s,x+8,108,CARD_W-16,3,C.editFlag);
    bx(s,x+8,108,1,168,C.line); bx(s,x+CARD_W-9,108,1,168,C.line); bx(s,x+8,108+168-1,CARD_W-16,1,C.line);
    tx(s,"[ Add screenshot ]",x+8,108,CARD_W-16,168,9,C.muted,false,"center");
    bx(s,x+8,282,CARD_W-16,1,C.line);
    tx(s,cr.name,x+8,288,CARD_W-16,28,8.5,C.dark,true,"left");
    [["CTR",cr.ctr],["CPA",cr.cpa],["Orders",String(cr.orders)]].forEach(([lbl,val],j) => {
      const sx = x+8+j*72;
      bx(s,sx,322,68,46,C.white); bx(s,sx,322,68,3,accs[i]);
      tx(s,lbl,sx+4,326,60,14,7,C.muted,false,"left");
      tx(s,val, sx+4,340,60,24,10,C.dark,true,"left");
    });
  }
  ftr(s);
}

function slide_ApproachesIntro(ctx) {
  const s = newSlide(ctx);
  bx(s,0,0,300,SH,C.navy); bx(s,0,0,300,4,C.cyan);
  const c1 = s.insertShape(SlidesApp.ShapeType.ELLIPSE,-60,300,280,280); c1.getFill().setSolidFill("#1565C0"); c1.getBorder().setTransparent(); c1.sendToBack();
  tx(s,"Approaches\nTaken",24,120,260,140,34,C.white,true,"left");
  tx(s,"Strategic initiatives\nexecuted this month",24,270,260,60,11,C.bluePale,false,"left");
  bx(s,300,0,2,SH,C.cyan); bx(s,302,0,418,SH,C.white);
  if (ctx.sheetData.approaches.length > 0) {
    const startY=60, cardH=120, gap=12;
    ctx.sheetData.approaches.forEach((ap,i) => {
      const y = startY+i*(cardH+gap);
      bx(s,320,y,382,cardH,C.blueLight); bx(s,320,y,382,4,C.navy); bx(s,320,y,4,cardH,C.cyan);
      bx(s,332,y+12,28,28,C.navy); tx(s,"0"+(i+1),332,y+12,28,28,11,C.white,true,"center");
      tx(s,ap.title,368,y+14,326,22,11,C.navy,true,"left");
      if (ap.points.length > 0) tx(s,ap.points[0].heading,368,y+40,326,18,9,C.mid,false,"left");
      if (ap.points.length > 1) tx(s,ap.points[1].heading,368,y+62,326,18,9,C.mid,false,"left");
    });
  }
  ftr(s);
}

function slide_Approach(ctx, num, approach) {
  const s = newSlide(ctx);
  bx(s,0,0,SW,46,C.navy); bx(s,0,44,SW,2,C.cyan); bx(s,0,0,4,46,C.cyan);
  tx(s,"APPROACH 0"+num,14,0,120,46,8.5,C.bluePale,true,"left");
  tx(s,approach.title,140,0,566,46,13,C.white,true,"left");
  bx(s,PAD,50,294,464,C.blueLight); bx(s,PAD,50,294,4,C.editFlag); bx(s,PAD,50,3,464,C.cyan);
  tx(s,"[ Replace with campaign image ]",PAD,50,294,464,9,C.muted,false,"center");
  const cardH = approach.points.length > 2 ? 142 : 184;
  approach.points.forEach((pt,i) => {
    const y = 54+i*(cardH+6); if (y+cardH > SH-16) return;
    bx(s,318,y,388,cardH,C.white); bx(s,318,y,388,1,C.line); bx(s,318,y,4,cardH,C.blue);
    bx(s,706,y,1,cardH,C.line); bx(s,318,y+cardH-1,388,1,C.line);
    const nb = s.insertShape(SlidesApp.ShapeType.ELLIPSE,330,y+10,28,28); nb.getFill().setSolidFill("#"+C.navy); nb.getBorder().setTransparent();
    tx(s,String(i+1),330,y+10,28,28,10,C.white,true,"center");
    tx(s,pt.heading,366,y+10,332,20,10,C.dark,true,"left");
    bx(s,366,y+34,330,1,C.line);
    tx(s,pt.body,366,y+40,332,cardH-50,8.5,C.mid,false,"left");
  });
  ftr(s);
}

function slide_SalesByProduct(ctx) {
  const s = newSlide(ctx);
  hdr(s,"SALES BY PRODUCT",ctx.M.currFull,C.blue);
  const products = ctx.sheetData.shopify.salesByProduct;
  if (!products || !products.length) {
    bx(s,PAD,70,SW-PAD*2,60,C.blueLight); bx(s,PAD,70,3,60,C.cyan);
    tx(s,"No product data for " + ctx.M.currLabel + ". Fill the Shopify_Products tab.",PAD+8,70,SW-PAD*2-10,60,10,C.muted,false,"center");
    ftr(s); return;
  }
  const totR = products.reduce((s,p) => s + p.revenue, 0);
  const rows = products.slice(0,8).map(p => [p.name, String(p.orders), fN(p.revenue), totR > 0 ? ((p.revenue/totR)*100).toFixed(1)+"%" : "—"]);
  tblFixed(s,PAD,52,[{label:"Product",w:328},{label:"Orders",w:96},{label:"Revenue (₹)",w:136},{label:"% of Total",w:94}],rows,28,30,C.navy);
  ftr(s);
}

function slide_SalesByLocation(ctx) {
  const s = newSlide(ctx);
  hdr(s,"SALES BY LOCATION",ctx.M.currFull,C.blue);
  const locs = ctx.sheetData.shopify.salesByLocation;
  if (!locs || !locs.length) {
    bx(s,PAD,70,SW-PAD*2,60,C.blueLight); bx(s,PAD,70,3,60,C.cyan);
    tx(s,"No location data for " + ctx.M.currLabel + ". Fill the Shopify_Locations tab.",PAD+8,70,SW-PAD*2-10,60,10,C.muted,false,"center");
    ftr(s); return;
  }
  const totR = locs.reduce((s,l) => s + l.revenue, 0);
  const rows = locs.slice(0,8).map(l => [l.location, String(l.orders), fN(l.revenue), totR > 0 ? ((l.revenue/totR)*100).toFixed(1)+"%" : "—"]);
  tblFixed(s,PAD,52,[{label:"Location",w:302},{label:"Orders",w:106},{label:"Revenue (₹)",w:148},{label:"% of Total",w:108}],rows,28,30,C.navy);
  ftr(s);
}

function slide_Retention(ctx) {
  const s = newSlide(ctx), ret = ctx.sheetData.retention;
  hdr(s,"RETENTION CHANNEL PERFORMANCE","",C.blue);
  const channels = [
    { name:"Email Marketing",    acc:"00897B", bg:"E0F2F1", rows:[["Sent",ret.email.sent],["Open Rate",ret.email.openRate],["Click Rate",ret.email.clickRate],["Revenue",ret.email.revenue],["ROAS",ret.email.roas]], note:ret.email.note },
    { name:"WhatsApp Campaigns", acc:"25D366", bg:"E8F5E9", rows:[["Sent",ret.whatsapp.sent],["Read Rate",ret.whatsapp.readRate],["Click Rate",ret.whatsapp.clickRate],["Revenue",ret.whatsapp.revenue],["ROAS",ret.whatsapp.roas]], note:ret.whatsapp.note },
    { name:"Push Notifications",  acc:"FB8C00", bg:"FFF3E0", rows:[["Sent",ret.push.sent],["Open Rate",ret.push.openRate],["Click Rate",ret.push.clickRate],["Revenue",ret.push.revenue],["ROAS",ret.push.roas]], note:ret.push.note },
  ];
  const CARD_W=226, CARD_H=408, GAP=5;
  channels.forEach((ch,i) => {
    const x = PAD+i*(CARD_W+GAP);
    bx(s,x,50,CARD_W,CARD_H,ch.bg); bx(s,x,50,CARD_W,4,ch.acc);
    bx(s,x,50,1,CARD_H,C.line); bx(s,x+CARD_W-1,50,1,CARD_H,C.line);
    bx(s,x+10,62,CARD_W-20,22,ch.acc); bx(s,x+10,62,CARD_W-20,2,C.white);
    tx(s,ch.name.toUpperCase(),x+10,62,CARD_W-20,22,7.5,C.white,true,"center");
    bx(s,x+10,88,CARD_W-20,1,C.line);
    const halfW = Math.floor((CARD_W-20)/2);
    ch.rows.slice(0,4).forEach(([lbl,val],j) => {
      const gx=x+10+(j%2)*halfW, gy=96+Math.floor(j/2)*82;
      bx(s,gx,gy,halfW-4,78,C.white); bx(s,gx,gy,halfW-4,3,ch.acc);
      tx(s,lbl,gx+6,gy+8,halfW-12,14,7.5,C.muted,false,"left");
      tx(s,val, gx+6,gy+24,halfW-12,32,16,ch.acc,true,"left");
    });
    if (ch.rows[4]) {
      const ry = 96+2*82+6;
      bx(s,x+10,ry-2,CARD_W-20,1,C.line);
      bx(s,x+10,ry+2,CARD_W-20,24,C.white); bx(s,x+10,ry+2,CARD_W-20,3,ch.acc);
      tx(s,ch.rows[4][0],x+14,ry+6,50,16,7.5,C.muted,false,"left");
      tx(s,ch.rows[4][1],x+68,ry+4,CARD_W-80,22,12,ch.acc,true,"left");
    }
    bx(s,x+10,342,CARD_W-20,1,C.line);
    tx(s,ch.note||"—",x+12,346,CARD_W-18,54,7.5,C.mid,false,"left");
  });
  bx(s,PAD,462,SW-PAD*2,52,C.blueLight); bx(s,PAD,462,SW-PAD*2,3,C.navy); bx(s,PAD,462,3,52,C.cyan);
  tx(s,"Total Retention Revenue: "+ret.totalRevenue+"   |   Returning Rate: "+ret.returningRate+"   |   Avg Repeat AOV: "+ret.avgRepeatAOV,PAD+8,462,SW-PAD*2-8,52,9,C.navy,true,"center");
  ftr(s);
}

function slide_Inferences(ctx) {
  const s = newSlide(ctx);
  hdr(s,"INFERENCES",ctx.M.currFull,C.blue);
  const CW=348, CH=218, GAP=8, startX=PAD, startY=52;
  ctx.inferences.forEach((inf,i) => {
    const x=startX+(i%2)*(CW+GAP), y=startY+Math.floor(i/2)*(CH+GAP);
    if (y+CH > SH-16) return;
    bx(s,x,y,CW,CH,C.white); bx(s,x,y,CW,4,inf.color);
    bx(s,x,y,1,CH,C.line); bx(s,x+CW-1,y,1,CH,C.line); bx(s,x,y+CH-1,CW,1,C.line);
    bx(s,x,y+4,3,CH-4,inf.color);
    const nb = s.insertShape(SlidesApp.ShapeType.ELLIPSE,x+12,y+12,28,28); nb.getFill().setSolidFill("#"+inf.color.replace(/^#+/,"").slice(0,6)); nb.getBorder().setTransparent();
    tx(s,inf.num,x+12,y+12,28,28,10,C.white,true,"center");
    tx(s,inf.title,x+50,y+14,CW-62,22,10,C.dark,true,"left");
    bx(s,x+10,y+44,CW-20,1,C.line);
    tx(s,inf.body,x+14,y+50,CW-26,CH-58,8.5,C.mid,false,"left");
  });
  ftr(s);
}

function slide_Recommendations(ctx) {
  const s = newSlide(ctx), M = ctx.M;
  hdr(s,"RECOMMENDATIONS — "+M.nextFull.toUpperCase(),"",C.blue);
  const allPos = [ctx.computed.comp.revenue,ctx.computed.comp.roas,ctx.computed.comp.customers].every(d => d.pos);
  bx(s,PAD,50,SW-PAD*2,20,C.blueLight); bx(s,PAD,50,3,20,C.cyan);
  tx(s,allPos ? "Compounding the momentum from "+M.currFull+" — focus, scale, and convert." : "Targeted actions to fix gaps and amplify wins from "+M.currFull+".",PAD+8,50,SW-PAD*2-10,20,8.5,C.navy,false,"left");
  const CW=348, CH=202, GAP=8, startX=PAD, startY=74;
  ctx.recs.forEach((rec,i) => {
    const x=startX+(i%2)*(CW+GAP), y=startY+Math.floor(i/2)*(CH+GAP);
    if (y+CH > SH-16) return;
    bx(s,x,y,CW,CH,C.white); bx(s,x,y,CW,1,C.line); bx(s,x+CW-1,y,1,CH,C.line); bx(s,x,y+CH-1,CW,1,C.line);
    bx(s,x,y,4,CH,rec.color);
    bx(s,x+12,y+10,36,36,rec.color); bx(s,x+12,y+10,36,3,C.cyan);
    tx(s,rec.num,x+12,y+10,36,36,16,C.white,true,"center");
    tx(s,rec.title,x+56,y+12,CW-68,26,9.5,C.dark,true,"left");
    bx(s,x+8,y+46,CW-16,1,C.line);
    (rec.bullets || []).slice(0,2).forEach((b,bi) => {
      bx(s,x+14,y+54+bi*52,4,4,rec.color);
      tx(s,b,x+24,y+50+bi*52,CW-36,46,8,C.mid,false,"left");
    });
  });
  ftr(s);
}

function slide_Projections(ctx) {
  const s = newSlide(ctx), M = ctx.M, pr = ctx.sheetData.projections, cm = ctx.computed;
  hdr(s,M.nextFull.toUpperCase()+" — FORECAST & BUDGET PLAN","",C.blue);
  const actualCust = cm.newCustomersNum || 0, targetCust = parseInt(pr.targetCustomers) || 0;
  const custNote = (actualCust > 0 && targetCust > 0) ? (targetCust >= actualCust ? "+" : "") + Math.round(((targetCust-actualCust)/actualCust)*100) + "% vs " + M.currLabel : "";
  const cards = [
    { l:"Projected Sales",       v:pr.targetSales,                      n:pr.growthNote, c:C.navy },
    { l:"Recommended Ad Budget", v:pr.targetBudget,                     n:"",            c:C.blue },
    { l:"Target New Customers",  v:targetCust > 0 ? String(targetCust) : "—", n:custNote, c:C.cyan },
    { l:"Target ROAS",           v:pr.targetROAS,                       n:pr.roasNote,   c:"4527A0" },
  ];
  cards.forEach((k,i) => {
    const y = 50+i*109;
    bx(s,448,y,260,106,C.blueLight); bx(s,448,y,260,4,k.c); bx(s,448,y,4,106,C.cyan);
    bx(s,448,y,260,1,C.line); bx(s,708,y,1,106,C.line); bx(s,448,y+106-1,260,1,C.line);
    tx(s,k.l,458,y+8,248,14,8,C.muted,false,"left");
    tx(s,k.v,458,y+26,248,38,22,C.dark,true,"left");
    if (k.n) tx(s,k.n,458,y+70,248,14,8,k.c,false,"left");
  });
  tx(s,"Revenue Trend",PAD,52,360,18,11,C.dark,true,"left");
  const chartData = [
    { label:M.prevLabel, value:cm.prevRev||100000, proj:false },
    { label:M.currLabel, value:cm.grossSalesNum||200000, proj:false },
    { label:M.nextFull.split(" ")[0]+"*", value:pr.targetSalesRaw||300000, proj:true },
  ];
  const maxV = Math.max(...chartData.map(d => d.value)) * 1.1;
  const cX=PAD+28, cY=80, cW=388, cH=400;
  bx(s,cX,cY,cW,cH,"F8FAFD"); bx(s,cX,cY,1,cH,C.line); bx(s,cX,cY+cH,cW,1,C.line);
  [0.25,0.5,0.75].forEach(pct => {
    const gy = cY+cH-Math.round(pct*cH);
    bx(s,cX+1,gy,cW-2,1,C.line);
    tx(s,fL(maxV*pct),cX-28,gy-8,26,16,7,C.muted,false,"right");
  });
  const pts = chartData.map((d,i) => ({
    px: cX+50+i*Math.floor((cW-100)/Math.max(1,chartData.length-1)),
    py: cY+cH-Math.round((d.value/maxV)*cH),
    label:d.label, value:d.value, proj:d.proj,
  }));
  pts.forEach((pt,i) => {
    if (i < pts.length-1) {
      const np = pts[i+1];
      const line = s.insertLine(SlidesApp.LineCategory.STRAIGHT,pt.px,pt.py,np.px,np.py);
      line.getLineFill().setSolidFill(np.proj ? "#90A4AE" : "#"+C.navy); line.setWeight(2.5);
      if (np.proj) line.setDashStyle(SlidesApp.DashStyle.DASH);
    }
    const dot = s.insertShape(SlidesApp.ShapeType.ELLIPSE,pt.px-6,pt.py-6,12,12);
    dot.getFill().setSolidFill(pt.proj ? "#90A4AE" : "#"+C.navy); dot.getBorder().setTransparent();
    tx(s,fL(pt.value),pt.px-32,pt.py-22,64,16,7.5,pt.proj?C.muted:C.navy,true,"center");
    tx(s,pt.label,pt.px-32,cY+cH+6,64,14,7.5,C.muted,false,"center");
  });
  tx(s,"* Projected",cX,cY+cH+22,80,12,7,C.muted,false,"left");
  ftr(s);
}

function slide_ThankYou(ctx) {
  const s = newSlide(ctx), CLIENT = ctx.CLIENT;
  bx(s,0,0,320,SH,C.navy); bx(s,0,0,320,4,C.cyan);
  const c1 = s.insertShape(SlidesApp.ShapeType.ELLIPSE,-60,280,260,260); c1.getFill().setSolidFill("#1565C0"); c1.getBorder().setTransparent(); c1.sendToBack();
  const c2 = s.insertShape(SlidesApp.ShapeType.ELLIPSE,60,-60,200,200);  c2.getFill().setSolidFill("#0D3080"); c2.getBorder().setTransparent(); c2.sendToBack();
  insertLogo(s,CLIENT.digifyceLogoFileId,24,24,140,38,C.navy,"DIGIFYCE");
  tx(s,"Thank",30,140,270,74,52,C.white,true,"left");
  tx(s,"You.", 30,216,270,74,52,C.cyan, true,"left");
  bx(s,30,300,200,2,C.cyan);
  const allGood = [ctx.computed.comp.revenue,ctx.computed.comp.roas].every(d => d.pos);
  tx(s,allGood ? "Let's build on this momentum." : "Let's act on these insights.",30,312,270,30,11,C.bluePale,false,"left");
  tx(s,"See you next month.",30,344,270,22,10,C.muted,false,"left");
  bx(s,320,0,2,SH,C.cyan); bx(s,322,0,398,SH,C.white); bx(s,322,0,398,4,C.navy);
  bx(s,340,80,360,80,C.blueLight); bx(s,340,80,360,4,C.navy); bx(s,340,80,3,80,C.cyan);
  tx(s,"GET IN TOUCH",350,88,200,14,8,C.muted,true,"left");
  tx(s,CLIENT.agentEmail,350,106,340,18,9,C.navy,true,"left");
  bx(s,340,172,360,80,C.blueLight); bx(s,340,172,360,4,C.navy); bx(s,340,172,3,80,C.cyan);
  tx(s,"WEBSITE",350,180,200,14,8,C.muted,true,"left");
  tx(s,"www.digifyce.com",350,198,340,18,9,C.navy,true,"left");
  insertLogo(s,CLIENT.clientLogoFileId,340,280,360,150,C.blueLight,CLIENT.name);
  bx(s,0,526,SW,14,C.navy); bx(s,0,526,SW,1,C.cyan);
  tx(s,"Confidential @ Digifyce",0,527,SW,13,7,C.bluePale,false,"center");
}


// ============================================================
// SECTION O — FORMATTERS & UTILITIES
// ============================================================

function insertLogo(slide, fileId, x, y, w, h, fallbackColor, fallbackText) {
  if (!fileId || fileId.startsWith("YOUR_") || fileId.trim() === "") {
    if (fallbackText) tx(slide, fallbackText, x, y+h/4, w, h/2, 11, C.white, true, "center");
    return;
  }
  try {
    // Use export URL to fetch image — avoids DriveApp.getFileById restriction
    const exportUrl = "https://drive.google.com/uc?export=download&id=" + fileId;
    const response  = UrlFetchApp.fetch(exportUrl, { muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) throw new Error("Image fetch failed: " + response.getResponseCode());
    const blob = response.getBlob();
    const img  = slide.insertImage(blob);
    img.setLeft(x).setTop(y).setWidth(w).setHeight(h);
  } catch(e) {
    Logger.log("WARN logo [" + fileId + "]: " + e.message);
    if (fallbackText) tx(slide, fallbackText, x, y+h/4, w, h/2, 11, C.white, true, "center");
  }
}

function fN(n) {
  if (n === null || n === undefined) return "—";
  n = parseFloat(String(n).replace(/,/g,""));
  if (isNaN(n)) return "—";
  return Math.round(n).toLocaleString("en-IN");
}

function fK(n) {
  n = parseFloat(n) || 0;
  if (n >= 100000) return "₹" + (n/100000).toFixed(1) + "L";
  if (n >= 1000)   return "₹" + (n/1000).toFixed(0) + "K";
  return "₹" + Math.round(n);
}

function fL(n) {
  n = parseFloat(n) || 0;
  if (n >= 10000000) return (n/10000000).toFixed(1) + "Cr";
  if (n >= 100000)   return (n/100000).toFixed(0) + "L";
  if (n >= 1000)     return (n/1000).toFixed(0) + "K";
  return String(Math.round(n));
}

function safeGet(url, label) {
  try {
    const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const text = resp.getContentText();
    let json;
    try { json = JSON.parse(text); } catch(e) { throw new Error(label + ": Invalid JSON response — " + text.slice(0,200)); }
    if (json.error) throw new Error(label + ": " + (json.error.message || JSON.stringify(json.error)));
    return json;
  } catch(e) {
    Logger.log("safeGet error [" + label + "]: " + e.message);
    throw e;
  }
}


// ============================================================
// DEBUG FUNCTION — Run this first to diagnose errors
// ============================================================

function debugClientAPIs() {
  const CLIENT_KEY = "CoreEat"; // ← change this to your client key
  Logger.log("=== DEBUG START: " + CLIENT_KEY + " ===");

  const result = serverGetClientFull(CLIENT_KEY);
  if (!result.success) { Logger.log("❌ Client not found: " + CLIENT_KEY); return; }
  const c = result.client;

  Logger.log("Client: " + c.name);
  Logger.log("Sheet ID: " + (c.sheet_id || "MISSING"));
  Logger.log("Meta Token: " + (c.meta_access_token ? c.meta_access_token.slice(0,12)+"..." : "MISSING"));
  Logger.log("Meta Ad Accounts: " + (c.meta_ad_account_ids || "MISSING"));
  Logger.log("GA4 Property: " + (c.ga4_property_id || "MISSING"));
  Logger.log("GSC Site: " + (c.gsc_site_url || "MISSING"));
  Logger.log("Google Ads: " + (c.google_ads_enabled === true || c.google_ads_enabled === "TRUE" ? "ENABLED" : "DISABLED"));

  const CLIENT = {
    name: c.name, agentEmail: c.agent_email,
    digifyceLogoFileId: DIGIFYCE_LOGO_ID, clientLogoFileId: c.client_logo_id || "",
    sheetId: c.sheet_id,
    meta: { accessToken: c.meta_access_token, adAccountIds: String(c.meta_ad_account_ids).split(",").map(s => s.trim()) },
    googleAds: { enabled: c.google_ads_enabled === true || c.google_ads_enabled === "TRUE", developerToken: c.google_ads_dev_token || "", customerId: c.google_ads_customer_id || "", managerCustomerId: c.google_ads_mcc_id || "" },
    ga4: { propertyId: c.ga4_property_id },
    gsc: { siteUrl: c.gsc_site_url },
  };

  const M = getMonthConfig("");

  Logger.log("\n[1] Sheet tabs...");
  try {
    const ss = SpreadsheetApp.openById(c.sheet_id);
    const tabs = ss.getSheets().map(s => s.getName());
    Logger.log("Found tabs: " + tabs.join(", "));
    const required = ["Retention","Retention_Summary","Shopify_Summary","Shopify_Products","Shopify_Locations","Strategy"];
    const missing = required.filter(t => !tabs.includes(t));
    if (missing.length) Logger.log("❌ MISSING TABS: " + missing.join(", "));
    else Logger.log("✅ All required tabs present");
    const d = readSheet(M, CLIENT);
    Logger.log("✅ Sheet read OK — Curr Sales:₹" + d.shopCurr.grossSales + " Orders:" + d.shopCurr.totalOrders + " Custs:" + d.shopCurr.newCustomers);
    Logger.log("   Prev Sales:₹" + d.shopPrev.grossSales + " Orders:" + d.shopPrev.totalOrders);
  } catch(e) { Logger.log("❌ Sheet error: " + e.message); }

  Logger.log("\n[2] Meta Ads...");
  try {
    const m = fetchMeta(M.currStart, M.currEnd, CLIENT);
    Logger.log("✅ Spend:₹" + m.totals.spend + " Rev:₹" + m.totals.revenue + " Impr:" + m.totals.impressions + " Campaigns:" + m.campaigns.length);
  } catch(e) { Logger.log("❌ Meta error: " + e.message); }

  Logger.log("\n[3] GA4...");
  try {
    const g = fetchGA4(M.currStart, M.currEnd, CLIENT);
    Logger.log("✅ Sessions:" + g.overall.sessions + " CR:" + g.overall.convRate + "%");
  } catch(e) { Logger.log("❌ GA4 error: " + e.message); }

  Logger.log("\n[4] Search Console...");
  try {
    const g = fetchGSC(M.currStart, M.currEnd, CLIENT);
    Logger.log("✅ Pages:" + g.topPages.length + " Queries:" + g.topQueries.length);
  } catch(e) { Logger.log("❌ GSC error: " + e.message); }

  Logger.log("\n[5] Google Ads...");
  if (!CLIENT.googleAds.enabled) {
    Logger.log("⏭️ Disabled");
  } else {
    try {
      const g = fetchGoogleAds(M.currStart, M.currEnd, CLIENT);
      Logger.log("✅ Spend:₹" + g.totals.spend + " Campaigns:" + g.campaigns.length);
    } catch(e) { Logger.log("❌ Google Ads error: " + e.message); }
  }

  Logger.log("\n[6] OpenAI...");
  try {
    const resp = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", {
      method: "POST",
      headers: { "Authorization": "Bearer " + OPENAI_API_KEY, "Content-Type": "application/json" },
      payload: JSON.stringify({ model: OPENAI_MODEL, max_tokens: 10, messages: [{ role:"user", content:"Say OK" }] }),
      muteHttpExceptions: true
    });
    const j = JSON.parse(resp.getContentText());
    if (j.error) Logger.log("❌ OpenAI: " + j.error.message);
    else Logger.log("✅ OpenAI: " + j.choices[0].message.content);
  } catch(e) { Logger.log("❌ OpenAI: " + e.message); }

  Logger.log("\n=== DEBUG DONE ===");
}

function resetAllClientSheets() {
  const clients = serverGetClients("owner", "ALL");
  clients.forEach(c => {
    if (c.sheet_id && c.sheet_id.trim() !== "") {
      Logger.log("Resetting: " + c.name);
      const result = serverResetClientSheet(c.client_key, "owner", "ALL");
      Logger.log(result.success ? "✅ " + result.message : "❌ " + result.error);
    }
  });
}