/* global Office, document */
"use strict";

const SETTINGS_KEY = "companySignatureSettings";

const DEFAULTS = {
  displayName:  "",
  jobTitle:     "",
  department:   "",
  companyName:  "",
  phone:        "",
  mobilePhone:  "",
  emailAddress: "",
  website:      "",
  address:      "",
  logoUrl:      "",
  accentColor:  "#0078d4",
  disclaimer:   "",
  autoInsert:   false,
};

function readSettings() {
  const rs  = Office.context.roamingSettings;
  const raw = rs.get(SETTINGS_KEY);
  return raw ? Object.assign({}, DEFAULTS, raw) : Object.assign({}, DEFAULTS);
}

function persistSettings(data, callback) {
  const rs = Office.context.roamingSettings;
  rs.set(SETTINGS_KEY, data);
  rs.saveAsync(callback);
}

function escHtml(s) {
  return String(s)
    .replace(/&/g,  "&amp;")
    .replace(/</g,  "&lt;")
    .replace(/>/g,  "&gt;")
    .replace(/"/g,  "&quot;");
}

function buildSignatureHtml(cfg) {
  const color   = escHtml(cfg.accentColor || "#0078d4");
  const name    = escHtml(cfg.displayName  || "");
  const company = escHtml(cfg.companyName  || "");
  const titleParts = [cfg.jobTitle, cfg.department].filter(Boolean);
  const titleLine  = escHtml(titleParts.join(" \u2022 "));

  let logoCell = "";
  if (cfg.logoUrl) {
    logoCell =
      '<td style="vertical-align:middle;padding-right:16px;border-right:2px solid ' + color + ';">' +
        '<img src="' + escHtml(cfg.logoUrl) + '"' +
             ' alt="' + escHtml(cfg.companyName) + ' logo"' +
             ' style="max-height:56px;max-width:140px;display:block;"/>' +
      '</td>';
  }

  const contacts = [];
  if (cfg.phone)        contacts.push('<span>&#128222;&nbsp;' + escHtml(cfg.phone) + '</span>');
  if (cfg.mobilePhone)  contacts.push('<span>&#128241;&nbsp;' + escHtml(cfg.mobilePhone) + '</span>');
  if (cfg.emailAddress) {
    contacts.push(
      '<a href="mailto:' + escHtml(cfg.emailAddress) + '"' +
         ' style="color:' + color + ';text-decoration:none;">' +
         escHtml(cfg.emailAddress) + '</a>'
    );
  }
  if (cfg.website) {
    contacts.push(
      '<a href="' + escHtml(cfg.website) + '"' +
         ' style="color:' + color + ';text-decoration:none;">' +
         escHtml(cfg.website) + '</a>'
    );
  }
  if (cfg.address) contacts.push('<span>&#127968;&nbsp;' + escHtml(cfg.address) + '</span>');

  const contactHtml = contacts.length
    ? '<div style="margin-top:6px;font-size:11px;color:#555555;line-height:1.8;">' +
        contacts.join("&nbsp;&nbsp;|&nbsp;&nbsp;") +
      '</div>'
    : "";

  const disclaimerHtml = cfg.disclaimer
    ? '<tr><td colspan="2" style="padding-top:10px;border-top:1px solid #e0e0e0;' +
        'font-size:9px;color:#999999;font-style:italic;line-height:1.5;">' +
        escHtml(cfg.disclaimer) + '</td></tr>'
    : "";

  let infoCell =
    '<div style="font-weight:700;font-size:15px;color:' + color + ';">' + name + '</div>';
  if (titleLine) {
    infoCell += '<div style="font-size:12px;color:#444444;margin-top:1px;">' + titleLine + '</div>';
  }
  if (company) {
    infoCell += '<div style="font-size:11px;font-weight:600;color:#222222;margin-top:2px;">' + company + '</div>';
  }
  infoCell += contactHtml;
  const padLeft = cfg.logoUrl ? "padding-left:14px;" : "";

  return (
    '<table cellpadding="0" cellspacing="0" border="0"' +
          ' style="font-family:Calibri,Arial,sans-serif;margin-top:18px;' +
                  'padding-top:12px;border-top:3px solid ' + color + ';">' +
      '<tr>' + logoCell +
        '<td style="vertical-align:middle;' + padLeft + '">' + infoCell + '</td>' +
      '</tr>' + disclaimerHtml +
    '</table>'
  );
}

// ============================================================
// Office.Body API wrappers
// ============================================================

/**
 * Office.Body.setSignatureAsync (Mailbox 1.13+)
 * Preferred: Outlook manages the signature area separately,
 * preventing duplication on reply/forward.
 * Falls back to prependAsync on older clients.
 */
function apiSetSignature(html, callback) {
  const item = Office.context.mailbox.item;
  if (!item || !item.body) {
    callback(new Error("No compose item is open."));
    return;
  }
  if (typeof item.body.setSignatureAsync === "function") {
    item.body.setSignatureAsync(
      html,
      { coercionType: Office.CoercionType.Html },
      function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          callback(null, "Signature set via setSignatureAsync.");
        } else {
          apiPrepend(html, callback);
        }
      }
    );
  } else {
    apiPrepend(html, callback);
  }
}

/**
 * Office.Body.prependAsync (Mailbox 1.1+)
 * Inserts HTML at the very start of the compose body.
 */
function apiPrepend(html, callback) {
  Office.context.mailbox.item.body.prependAsync(
    html,
    { coercionType: Office.CoercionType.Html },
    function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        callback(null, "Signature prepended via prependAsync.");
      } else {
        callback(new Error(result.error.message));
      }
    }
  );
}

/**
 * Office.Body.appendAsync (Mailbox 1.1+)
 * Appends HTML at the end of the compose body.
 */
function apiAppend(html, callback) {
  Office.context.mailbox.item.body.appendAsync(
    html,
    { coercionType: Office.CoercionType.Html },
    function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        callback(null, "Signature appended via appendAsync.");
      } else {
        callback(new Error(result.error.message));
      }
    }
  );
}

// ============================================================
// UI helpers
// ============================================================

let currentSettings = {};

function showStatus(msg, type) {
  const el = document.getElementById("statusBanner");
  if (!el) return;
  el.textContent = msg;
  el.className   = "status-banner status-" + type;
  el.classList.remove("hidden");
  clearTimeout(el._t);
  el._t = setTimeout(function () { el.classList.add("hidden"); }, 3500);
}

function showPanel(name) {
  const sp = document.getElementById("settingsPanel");
  const mp = document.getElementById("mainPanel");
  if (name === "settings") {
    populateForm(currentSettings);
    sp.classList.remove("hidden");
    mp.classList.add("hidden");
  } else {
    sp.classList.add("hidden");
    mp.classList.remove("hidden");
  }
}

function renderPreview() {
  const el = document.getElementById("signaturePreview");
  if (el) el.innerHTML = buildSignatureHtml(currentSettings);
}

function populateForm(s) {
  function g(id) { return document.getElementById(id); }
  g("displayName").value    = s.displayName  || "";
  g("jobTitle").value       = s.jobTitle      || "";
  g("department").value     = s.department    || "";
  g("companyName").value    = s.companyName   || "";
  g("phone").value          = s.phone         || "";
  g("mobilePhone").value    = s.mobilePhone   || "";
  g("emailAddress").value   = s.emailAddress  || "";
  g("website").value        = s.website       || "";
  g("address").value        = s.address       || "";
  g("logoUrl").value        = s.logoUrl       || "";
  g("accentColor").value    = s.accentColor   || "#0078d4";
  g("colorHex").textContent = s.accentColor   || "#0078d4";
  g("disclaimer").value     = s.disclaimer    || "";
  g("autoInsert").checked   = !!s.autoInsert;
}

function collectForm() {
  function g(id) { return document.getElementById(id); }
  return {
    displayName:  g("displayName").value.trim(),
    jobTitle:     g("jobTitle").value.trim(),
    department:   g("department").value.trim(),
    companyName:  g("companyName").value.trim(),
    phone:        g("phone").value.trim(),
    mobilePhone:  g("mobilePhone").value.trim(),
    emailAddress: g("emailAddress").value.trim(),
    website:      g("website").value.trim(),
    address:      g("address").value.trim(),
    logoUrl:      g("logoUrl").value.trim(),
    accentColor:  g("accentColor").value,
    disclaimer:   g("disclaimer").value.trim(),
    autoInsert:   g("autoInsert").checked,
  };
}

// ============================================================
// Event bindings
// ============================================================

function bindEvents() {
  document.getElementById("settingsToggle").addEventListener("click", function () { showPanel("settings"); });
  document.getElementById("cancelSettings").addEventListener("click", function () { showPanel("main"); });
  document.getElementById("refreshPreview").addEventListener("click", renderPreview);

  document.getElementById("accentColor").addEventListener("input", function () {
    document.getElementById("colorHex").textContent = this.value;
  });

  document.getElementById("saveSettings").addEventListener("click", function () {
    currentSettings = collectForm();
    persistSettings(currentSettings, function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        renderPreview();
        showPanel("main");
        showStatus("Settings saved.", "success");
      } else {
        showStatus("Save failed: " + result.error.message, "error");
      }
    });
  });

  document.getElementById("btnSetSignature").addEventListener("click", function () {
    apiSetSignature(buildSignatureHtml(currentSettings), function (err, msg) {
      if (err) showStatus(err.message, "error");
      else     showStatus(msg, "success");
    });
  });

  document.getElementById("btnPrepend").addEventListener("click", function () {
    apiPrepend(buildSignatureHtml(currentSettings), function (err, msg) {
      if (err) showStatus(err.message, "error");
      else     showStatus(msg, "success");
    });
  });

  document.getElementById("btnAppend").addEventListener("click", function () {
    apiAppend(buildSignatureHtml(currentSettings), function (err, msg) {
      if (err) showStatus(err.message, "error");
      else     showStatus(msg, "success");
    });
  });
}

// ============================================================
// Initialisation
// ============================================================

Office.onReady(function (info) {
  if (info.host !== Office.HostType.Outlook) return;
  currentSettings = readSettings();
  renderPreview();
  bindEvents();
  if (currentSettings.autoInsert) {
    apiSetSignature(buildSignatureHtml(currentSettings), function () { /* silent */ });
  }
});
