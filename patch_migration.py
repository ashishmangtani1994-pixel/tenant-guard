#!/usr/bin/env python3
"""
Run from your tenant-guard folder:
  python3 patch_migration.py
It reads index.html, applies the Migration tab, writes it back.
"""

import sys, os, shutil
from datetime import datetime

SRC = "index.html"
BAK = f"index.html.bak.{datetime.now().strftime('%Y%m%d_%H%M%S')}"

if not os.path.exists(SRC):
    print(f"ERROR: {SRC} not found. Run this from /Users/ashish/Desktop/tenant-guard/"); sys.exit(1)

shutil.copy2(SRC, BAK)
print(f"Backup written: {BAK}")

with open(SRC, "r", encoding="utf-8") as f:
    html = f.read()

changes = 0

# ── 1. INFO_TABS: add "migration" ──────────────────────────────────────────
old = '"devcompliance","privileged","securescore","jailbroken","legacyauth","appcreds","oauth"];'
new = '"devcompliance","privileged","securescore","jailbroken","legacyauth","appcreds","oauth","migration"];'
if old in html: html = html.replace(old, new, 1); changes += 1; print("✓ Patch 1: INFO_TABS")
else: print("✗ Patch 1 not found — already applied or source mismatch")

# ── 2. syncFilterBar: hide infoCards for migration ─────────────────────────
old = 'el("infoCards").classList.toggle("hide", !info || dash);'
new = 'el("infoCards").classList.toggle("hide", !info || dash || tab==="migration");'
if old in html: html = html.replace(old, new, 1); changes += 1; print("✓ Patch 2: syncFilterBar infoCards")
else: print("✗ Patch 2 not found")

# ── 3. render() form-tab block ─────────────────────────────────────────────
old = (
  '  const isTeams = tab==="teams", isUsers = tab==="users", isGuests = tab==="guests";\n'
  '  const isFormTab = isTeams || isUsers || isGuests;\n'
  '  el("teamsPanel").classList.toggle("hide", !isTeams);\n'
  '  el("usersPanel").classList.toggle("hide", !isUsers);\n'
  '  el("guestsPanel").classList.toggle("hide", !isGuests);\n'
  '  document.querySelector(".tablewrap").classList.toggle("hide", isFormTab);\n'
  '  el("filterbar").classList.toggle("hide", isFormTab);\n'
  '  el("actionbar").classList.toggle("hide", isFormTab);\n'
  '  document.querySelector(".controls").classList.toggle("hide", isFormTab);\n'
  '  if(isFormTab) return;'
)
new = (
  '  const isTeams = tab==="teams", isUsers = tab==="users", isGuests = tab==="guests", isMig = tab==="migration";\n'
  '  const isFormTab = isTeams || isUsers || isGuests || isMig;\n'
  '  el("teamsPanel").classList.toggle("hide", !isTeams);\n'
  '  el("usersPanel").classList.toggle("hide", !isUsers);\n'
  '  el("guestsPanel").classList.toggle("hide", !isGuests);\n'
  '  el("migrationPanel").classList.toggle("hide", !isMig);\n'
  '  document.querySelector(".tablewrap").classList.toggle("hide", isFormTab);\n'
  '  el("filterbar").classList.toggle("hide", isFormTab);\n'
  '  el("actionbar").classList.toggle("hide", isFormTab);\n'
  '  document.querySelector(".controls").classList.toggle("hide", isFormTab);\n'
  '  if(isFormTab){ if(isMig) renderMigration(); return; }'
)
if old in html: html = html.replace(old, new, 1); changes += 1; print("✓ Patch 3: render() form-tab block")
else: print("✗ Patch 3 not found")

# ── 4. Sidenav: Migration navgroup before </aside> ─────────────────────────
old = '\n    </aside>'
new = (
  '\n'
  '      <div class="navgroup" data-group="migration">\n'
  '        <button class="navhead" data-grouphead="migration"><span class="caret">▾</span> Migration</button>\n'
  '        <div class="navitems">\n'
  '          <button class="tab" data-tab="migration">Tenant Migration</button>\n'
  '        </div>\n'
  '      </div>\n'
  '\n'
  '    </aside>'
)
if old in html: html = html.replace(old, new, 1); changes += 1; print("✓ Patch 4: sidenav Migration tab")
else: print("✗ Patch 4 not found")

# ── 5. HTML: migrationPanel div before tablewrap ───────────────────────────
old = '      <div class="tablewrap">'
new = (
  '      <!-- ==== MIGRATION PANEL ==== -->\n'
  '      <div class="card hide" id="migrationPanel" style="margin:0 0 20px;padding:24px"></div>\n'
  '\n'
  '      <div class="tablewrap">'
)
if old in html: html = html.replace(old, new, 1); changes += 1; print("✓ Patch 5: migrationPanel HTML")
else: print("✗ Patch 5 not found")

# ── 6. Migration JS block before boot comment ──────────────────────────────
MIG_JS = r"""
/* ====================================================================
   MIGRATION TAB — dual-tenant MSAL + scan
   Two independent PublicClientApplication instances so each tenant
   gets its own session-storage token cache and sign-in account.
   Required delegated scopes: User.Read, Reports.Read.All, Sites.Read.All
   ==================================================================== */
let msalSrc = null, msalTgt = null;
let srcAccount = null, tgtAccount = null;
let srcData = null, tgtData = null;
let srcScanning = false, tgtScanning = false;

const MIG_SCOPES = ["User.Read", "Reports.Read.All", "Sites.Read.All"];

async function migGetApp(side) {
  if (side === "src" && msalSrc) return msalSrc;
  if (side === "tgt" && msalTgt) return msalTgt;
  const app = new msal.PublicClientApplication({
    auth: { clientId: clientId(), authority: "https://login.microsoftonline.com/organizations", redirectUri: REDIRECT },
    cache: { cacheLocation: "sessionStorage", storeAuthStateInCookie: false }
  });
  await app.initialize();
  if (side === "src") msalSrc = app; else msalTgt = app;
  return app;
}

async function migConnect(side) {
  const btn = el(side + "ConnectBtn");
  if (btn) { btn.disabled = true; btn.textContent = "Signing in\u2026"; }
  try {
    const app = await migGetApp(side);
    const result = await app.loginPopup({ scopes: MIG_SCOPES });
    if (side === "src") srcAccount = result.account; else tgtAccount = result.account;
    renderMigration();
    await migScan(side);
  } catch(e) {
    toast("Sign-in failed: " + (e.message || String(e)).slice(0, 120), true);
    if (btn) { btn.disabled = false; btn.textContent = "Connect"; }
    renderMigration();
  }
}

async function migGetToken(side) {
  const app  = side === "src" ? msalSrc  : msalTgt;
  const acct = side === "src" ? srcAccount : tgtAccount;
  try {
    const r = await app.acquireTokenSilent({ scopes: MIG_SCOPES, account: acct });
    return r.accessToken;
  } catch(e) {
    const r = await app.acquireTokenPopup({ scopes: MIG_SCOPES, account: acct });
    return r.accessToken;
  }
}

async function migScan(side) {
  if (side === "src") srcScanning = true; else tgtScanning = true;
  renderMigration();
  const result = { mailboxes: null, onedrive: null, sharepoint: null };

  try {
    const tok = await migGetToken(side);

    /* ---- Mailboxes via usage report ---- */
    try {
      let url = GRAPH_BETA + "/reports/getMailboxUsageDetail(period='D90')?$format=application/json&$top=500";
      const rows = []; let pages = 0;
      while (url && pages++ < 8) {
        const r = await fetch(url, { headers: { Authorization: "Bearer " + tok } });
        if (!r.ok) throw new Error("HTTP " + r.status);
        const j = await r.json(); rows.push(...(j.value || [])); url = j["@odata.nextLink"] || null;
      }
      const usedBytes = rows.reduce((a, m) => a + (parseFloat(m.storageUsedInBytes) || 0), 0);
      const shared = rows.filter(m => {
        const rt = String(m.recipientType || m.recipientTypeDetails || "").toLowerCase();
        if (rt.includes("shared")) return true;
        const hasLic = Array.isArray(m.assignedLicenses) ? m.assignedLicenses.length > 0 : !!m.hasLicense;
        return !hasLic && parseFloat(m.storageUsedInBytes || 0) > 0;
      }).length;
      result.mailboxes = { total: rows.length, user: rows.length - shared, shared,
        usedGB: (usedBytes / Math.pow(2, 30)).toFixed(2) };
    } catch(e) { result.mailboxes = { error: String(e.message || e).slice(0, 80) }; }

    /* ---- OneDrive via usage report ---- */
    try {
      let url = GRAPH_BETA + "/reports/getOneDriveUsageAccountDetail(period='D90')?$format=application/json&$top=500";
      const rows = []; let pages = 0;
      while (url && pages++ < 8) {
        const r = await fetch(url, { headers: { Authorization: "Bearer " + tok } });
        if (!r.ok) throw new Error("HTTP " + r.status);
        const j = await r.json(); rows.push(...(j.value || [])); url = j["@odata.nextLink"] || null;
      }
      const usedBytes  = rows.reduce((a, o) => a + (parseFloat(o.storageUsedInBytes)     || 0), 0);
      const allocBytes = rows.reduce((a, o) => a + (parseFloat(o.storageAllocatedInBytes) || 0), 0);
      result.onedrive = { total: rows.length,
        usedGB:  (usedBytes  / Math.pow(2, 30)).toFixed(2),
        allocGB: (allocBytes / Math.pow(2, 30)).toFixed(2) };
    } catch(e) { result.onedrive = { error: String(e.message || e).slice(0, 80) }; }

    /* ---- SharePoint sites + storage ---- */
    try {
      let siteCount = 0, url = GRAPH + "/sites?search=*&$select=id,displayName&$top=200", pages = 0;
      while (url && pages++ < 5) {
        const r = await fetch(url, { headers: { Authorization: "Bearer " + tok, ConsistencyLevel: "eventual" } });
        if (!r.ok) throw new Error("HTTP " + r.status);
        const j = await r.json(); siteCount += (j.value || []).length; url = j["@odata.nextLink"] || null;
      }
      let spUsedGB = null;
      try {
        const sr = await fetch(GRAPH_BETA + "/reports/getSharePointSiteUsageDetail(period='D90')?$format=application/json&$top=500",
          { headers: { Authorization: "Bearer " + tok } });
        if (sr.ok) {
          const sj = await sr.json();
          const b = (sj.value || []).reduce((a, s) => a + (parseFloat(s.storageUsedInBytes) || 0), 0);
          spUsedGB = (b / Math.pow(2, 30)).toFixed(2);
        }
      } catch(_e) { /* storage report optional */ }
      result.sharepoint = { total: siteCount, usedGB: spUsedGB };
    } catch(e) { result.sharepoint = { error: String(e.message || e).slice(0, 80) }; }

  } catch(e) {
    toast("Scan error: " + (e.message || String(e)).slice(0, 120), true);
  }

  if (side === "src") { srcData = result; srcScanning = false; }
  else               { tgtData = result; tgtScanning = false; }
  renderMigration();
}

/* ---- UI helpers ---- */
function _migStatRow(icon, lbl, val, sub, col) {
  return '<div style="display:flex;align-items:center;gap:12px;padding:12px 0;border-bottom:1px solid var(--border)">'
    + '<div style="width:30px;text-align:center;font-size:20px">' + icon + '</div>'
    + '<div style="flex:1"><div style="font-size:13px;font-weight:600;color:var(--text)">' + lbl + '</div>'
    + (sub ? '<div style="font-size:11.5px;color:var(--muted);margin-top:2px">' + sub + '</div>' : '')
    + '</div>'
    + '<div style="font-family:\'Space Grotesk\';font-weight:700;font-size:20px;color:' + col + '">' + val + '</div>'
    + '</div>';
}
function _migErrRow(msg) {
  return '<div style="font-size:12px;color:var(--warn);padding:8px 0;border-bottom:1px solid var(--border)">\u26a0 ' + escapeHtml(msg) + '</div>';
}

function migTenantCard(side) {
  const label   = side === "src" ? "Source Tenant" : "Target Tenant";
  const acct    = side === "src" ? srcAccount      : tgtAccount;
  const data    = side === "src" ? srcData         : tgtData;
  const loading = side === "src" ? srcScanning     : tgtScanning;
  const col     = side === "src" ? "var(--accent)"  : "var(--accent-2)";

  const hdr =
    '<div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:18px">'
    + '<div>'
    + '<div style="font-family:\'Space Grotesk\';font-weight:700;font-size:15px;color:' + col + '">' + label + '</div>'
    + (acct
        ? '<div style="font-size:11.5px;color:var(--muted);margin-top:2px;font-family:var(--mono)">' + escapeHtml(acct.username) + '</div>'
        : '<div style="font-size:11.5px;color:var(--faint);margin-top:2px">Not connected</div>')
    + '</div>'
    + (acct
        ? '<button class="btn" onclick="migScan(\'' + side + '\')" '
          + (loading ? 'disabled' : '') + ' style="font-size:12px;padding:6px 12px">'
          + (loading ? '<span class="spin"></span>&nbsp;Scanning\u2026' : '\u21bb Rescan') + '</button>'
        : '<button class="btn primary" id="' + side + 'ConnectBtn" onclick="migConnect(\'' + side + '\')"'
          + ' style="font-size:13px;padding:9px 18px">Connect</button>')
    + '</div>';

  if (!acct) return '<div class="card" style="padding:22px;min-height:290px">' + hdr
    + '<div style="text-align:center;padding:36px 0 26px">'
    + '<div style="font-size:34px;margin-bottom:10px">\ud83d\udd17</div>'
    + '<div style="font-size:13px;color:var(--muted)">Sign in to scan this tenant\u2019s data</div>'
    + '<div style="font-size:11px;color:var(--faint);margin-top:7px;line-height:1.6">'
    + 'Needs <code style="color:var(--accent)">Reports.Read.All</code> &amp; '
    + '<code style="color:var(--accent)">Sites.Read.All</code><br>(admin consent required)</div>'
    + '</div></div>';

  if (loading && !data) return '<div class="card" style="padding:22px;min-height:290px">' + hdr
    + '<div class="skelwrap"><div class="skel"></div><div class="skel"></div><div class="skel"></div></div>'
    + '</div>';

  const mbx = data && data.mailboxes;
  const od  = data && data.onedrive;
  const sp  = data && data.sharepoint;

  return '<div class="card" style="padding:22px;min-height:290px">' + hdr
    + (!data ? '<div style="text-align:center;color:var(--muted);font-size:13px;padding:20px 0">Click \u21bb Rescan to load data</div>' : '')
    + (mbx ? (mbx.error ? _migErrRow('Mailboxes: ' + mbx.error)
        : _migStatRow('\ud83d\udcec', 'Mailboxes', mbx.total,
            mbx.user + ' user \u00b7 ' + mbx.shared + ' shared \u00b7 ' + mbx.usedGB + ' GB used', col)) : '')
    + (od  ? (od.error  ? _migErrRow('OneDrive: ' + od.error)
        : _migStatRow('\u2601\ufe0f', 'OneDrive accounts', od.total,
            od.usedGB + ' GB used \u00b7 ' + od.allocGB + ' GB allocated', col)) : '')
    + (sp  ? (sp.error  ? _migErrRow('SharePoint: ' + sp.error)
        : _migStatRow('\ud83c\udfe2', 'SharePoint sites', sp.total,
            sp.usedGB ? sp.usedGB + ' GB used' : 'Storage data unavailable', col)) : '')
    + '</div>';
}

function migCompareTable() {
  if (!srcData || !tgtData) return '';
  function _get(d, key, sub) {
    if (!d || !d[key] || d[key].error) return '\u2014';
    var v = d[key][sub]; return v != null ? String(v) : '\u2014';
  }
  function _gb(d, key, sub) { var v = _get(d,key,sub); return v==='\u2014' ? '\u2014' : v+' GB'; }
  function _row(lbl, src, tgt) {
    return '<tr>'
      + '<td style="padding:10px 14px;font-size:13px;color:var(--muted);font-weight:600">' + lbl + '</td>'
      + '<td style="padding:10px 14px;font-family:\'Space Grotesk\';font-weight:700;font-size:15px;color:var(--accent)">' + src + '</td>'
      + '<td style="padding:10px 14px;font-family:\'Space Grotesk\';font-weight:700;font-size:15px;color:var(--accent-2)">' + tgt + '</td>'
      + '</tr>';
  }
  return '<div class="card" style="margin-top:18px;padding:0;overflow:hidden">'
    + '<div style="padding:16px 20px;border-bottom:1px solid var(--border);font-family:\'Space Grotesk\';font-weight:700;font-size:14px">Migration Scope Comparison</div>'
    + '<table style="width:100%;border-collapse:collapse">'
    + '<thead><tr style="background:var(--panel-2)">'
    + '<th style="padding:10px 14px;text-align:left;font-size:11px;text-transform:uppercase;letter-spacing:.05em;color:var(--faint);font-weight:600;width:40%">Metric</th>'
    + '<th style="padding:10px 14px;text-align:left;font-size:11px;text-transform:uppercase;letter-spacing:.05em;color:var(--accent);font-weight:600">Source</th>'
    + '<th style="padding:10px 14px;text-align:left;font-size:11px;text-transform:uppercase;letter-spacing:.05em;color:var(--accent-2);font-weight:600">Target</th>'
    + '</tr></thead><tbody>'
    + _row('Total mailboxes',    _get(srcData,'mailboxes','total'),   _get(tgtData,'mailboxes','total'))
    + _row('User mailboxes',     _get(srcData,'mailboxes','user'),    _get(tgtData,'mailboxes','user'))
    + _row('Shared mailboxes',   _get(srcData,'mailboxes','shared'),  _get(tgtData,'mailboxes','shared'))
    + _row('Mailbox storage',    _gb(srcData,'mailboxes','usedGB'),   _gb(tgtData,'mailboxes','usedGB'))
    + _row('OneDrive accounts',  _get(srcData,'onedrive','total'),    _get(tgtData,'onedrive','total'))
    + _row('OneDrive used',      _gb(srcData,'onedrive','usedGB'),    _gb(tgtData,'onedrive','usedGB'))
    + _row('OneDrive allocated', _gb(srcData,'onedrive','allocGB'),   _gb(tgtData,'onedrive','allocGB'))
    + _row('SharePoint sites',   _get(srcData,'sharepoint','total'),  _get(tgtData,'sharepoint','total'))
    + _row('SharePoint storage', _gb(srcData,'sharepoint','usedGB'),  _gb(tgtData,'sharepoint','usedGB'))
    + '</tbody></table></div>';
}

function renderMigration() {
  var panel = el('migrationPanel');
  if (!panel) return;
  panel.innerHTML =
    '<h2 style="font-family:\'Space Grotesk\';font-size:20px;margin:0 0 6px">Tenant Migration Scanner</h2>'
    + '<p style="color:var(--muted);font-size:13px;margin:0 0 20px">Sign in to each tenant independently to scan mailboxes, OneDrive, and SharePoint data side by side.</p>'
    + '<div style="display:grid;grid-template-columns:1fr 1fr;gap:16px">'
    + migTenantCard('src')
    + migTenantCard('tgt')
    + '</div>'
    + migCompareTable();
}

"""

old = '/* ---------- boot ---------- */'
if old in html:
    html = html.replace(old, MIG_JS + old, 1)
    changes += 1; print("✓ Patch 6: migration JS block")
else:
    print("✗ Patch 6 not found")

with open(SRC, "w", encoding="utf-8") as f:
    f.write(html)

print(f"\nDone — {changes}/6 patches applied. File size: {len(html):,} bytes")
print(f"Now run:  ./deploy.sh tenant-guard")
