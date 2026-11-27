/**
 * Renata PLC — Staging Log Sheet -> GitHub Auto-Sync & Live REST API
 *
 * Setup:
 * 1. Run `setupConfig()` once (or use menu: GitHub Sync -> Configure GitHub)
 *    with your GitHub Personal Access Token (repo scope).
 * 2. Run `installTriggers()` once (or use menu: GitHub Sync -> Enable Auto-Sync)
 *    so any new entry or modification automatically commits to GitHub.
 * 3. (Optional) Deploy -> New deployment -> Web app (Execute as: Me, Who has access: Anyone)
 *    to enable the live Apps Script REST API endpoints (`doGet` / `doPost`).
 */

const DEFAULT_OWNER = 'arif0850';
const DEFAULT_REPO = 'staging-log-sheet';
const DEFAULT_BRANCH = 'main';
const DEFAULT_SPREADSHEET_ID = '1HVJLz9_5P0B8x1KnxnLdx6QdNjohN6xcKOVEB5ECDhg';
const SHEET_NOT_EXITED = 'Not Exited';
const SHEET_SUMMARY = 'Summary';

function getSpreadsheet() {
  var active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID') || DEFAULT_SPREADSHEET_ID;
  return SpreadsheetApp.openById(id);
}

const HEADERS = [
  'sl',
  'stagingRoom',
  'pageNo',
  'productName',
  'batchNo',
  'itemDescription',
  'containerIds',
  'noOfContainers',
  'stagingDate',
  'stagingValidity',
  'validityStatus',
  'daysInStaging',
  'exitDate',
  'remarks',
  'note',
  'sourcePhoto'
];

// ============================================================================
// MENU & TRIGGERS
// ============================================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🔄 GitHub Sync')
    .addItem('Sync Now to GitHub', 'manualSync')
    .addItem('Enable Auto-Sync Triggers', 'installTriggers')
    .addItem('Configure GitHub Settings', 'promptConfig')
    .addItem('Check Sync Status', 'checkStatus')
    .addToUi();
}

/**
 * Installable trigger handler for both onEdit and onChange events.
 * Simple onEdit() cannot call UrlFetchApp, so we use installable triggers.
 */
function onSheetChange(e) {
  try {
    if (e && e.range) {
      const sheetName = e.range.getSheet().getName();
      if (sheetName !== SHEET_NOT_EXITED && sheetName !== SHEET_SUMMARY) {
        return;
      }
    }
    const triggerType = e && e.changeType ? 'onChange:' + e.changeType : 'onEdit';
    syncToGitHub(triggerType);
  } catch (err) {
    console.error('Auto-sync error:', err);
  }
}

function manualSync() {
  const result = syncToGitHub('manual');
  SpreadsheetApp.getUi().alert(
    '✅ Synced to GitHub!\n\n' +
    'Entries: ' + result.totalEntries + '\n' +
    'Containers: ' + result.totalContainers + '\n' +
    'Commit: ' + result.commitSha.substring(0, 7)
  );
}

function installTriggers() {
  const ss = getSpreadsheet();
  const existing = ScriptApp.getProjectTriggers();
  existing.forEach(function (t) {
    if (t.getHandlerFunction() === 'onSheetChange') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('onSheetChange')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  ScriptApp.newTrigger('onSheetChange')
    .forSpreadsheet(ss)
    .onChange()
    .create();

  if (SpreadsheetApp.getActiveSpreadsheet()) {
    SpreadsheetApp.getUi().alert(
      '✅ Auto-Sync Triggers Enabled!\n\n' +
      'Any new entry or modification in "' + SHEET_NOT_EXITED + '" or "' + SHEET_SUMMARY + '" will automatically push to GitHub.'
    );
  }
  return 'Auto-Sync Triggers Enabled!';
}

function setupConfig(token, owner, repo, branch) {
  const props = PropertiesService.getScriptProperties();
  props.setProperties({
    GITHUB_TOKEN: token || props.getProperty('GITHUB_TOKEN') || '',
    GITHUB_OWNER: owner || DEFAULT_OWNER,
    GITHUB_REPO: repo || DEFAULT_REPO,
    GITHUB_BRANCH: branch || DEFAULT_BRANCH
  });
}

function promptConfig() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const currentRepo = (props.getProperty('GITHUB_OWNER') || DEFAULT_OWNER) + '/' + (props.getProperty('GITHUB_REPO') || DEFAULT_REPO);

  const tokenResp = ui.prompt(
    'GitHub Configuration',
    'Target Repo: ' + currentRepo + '\nEnter GitHub Personal Access Token (repo scope):',
    ui.ButtonSet.OK_CANCEL
  );
  if (tokenResp.getSelectedButton() !== ui.Button.OK) return;
  const token = tokenResp.getResponseText().trim();
  if (!token) {
    ui.alert('Token cannot be empty.');
    return;
  }
  setupConfig(token, DEFAULT_OWNER, DEFAULT_REPO, DEFAULT_BRANCH);
  ui.alert('✅ GitHub token saved in Script Properties.');
}

function checkStatus() {
  const props = PropertiesService.getScriptProperties();
  const hasToken = !!props.getProperty('GITHUB_TOKEN');
  const owner = props.getProperty('GITHUB_OWNER') || DEFAULT_OWNER;
  const repo = props.getProperty('GITHUB_REPO') || DEFAULT_REPO;
  const branch = props.getProperty('GITHUB_BRANCH') || DEFAULT_BRANCH;
  const lastSync = props.getProperty('LAST_SYNC_AT') || 'Never';
  const lastCommit = props.getProperty('LAST_COMMIT_SHA') || 'None';
  const triggers = ScriptApp.getProjectTriggers().filter(function (t) {
    return t.getHandlerFunction() === 'onSheetChange';
  }).length;

  SpreadsheetApp.getUi().alert(
    'GitHub Sync Status\n\n' +
    'Repo: ' + owner + '/' + repo + ' (' + branch + ')\n' +
    'Token configured: ' + (hasToken ? 'Yes' : 'No') + '\n' +
    'Active auto-sync triggers: ' + triggers + '\n' +
    'Last synced: ' + lastSync + '\n' +
    'Last commit: ' + lastCommit
  );
}

// ============================================================================
// DATA EXTRACTION
// ============================================================================

function formatCellValue(val) {
  if (val === '' || val === null || val === undefined) return null;
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'dd.MM.yy');
  }
  return val;
}

function buildDataset(sourceLabel) {
  const ss = getSpreadsheet();
  const ws = ss.getSheetByName(SHEET_NOT_EXITED);
  if (!ws) throw new Error('Sheet "' + SHEET_NOT_EXITED + '" not found.');

  const lastRow = ws.getLastRow();
  const lastCol = Math.max(ws.getLastColumn(), HEADERS.length);
  const allValues = ws.getRange(1, 1, lastRow, lastCol).getValues();

  const title = allValues[0] && allValues[0][0] ? String(allValues[0][0]) : 'Log Sheet for Staging Dispensed Material and Intermediate Product';
  const subtitle = allValues[1] && allValues[1][0] ? String(allValues[1][0]) : '';
  const rawHeaders = allValues[3] ? allValues[3].slice(0, HEADERS.length).map(function (h) { return String(h || ''); }) : HEADERS;

  const entries = [];
  for (var r = 4; r < allValues.length; r++) {
    var row = allValues[r];
    var sl = row[0];
    var productName = row[3];
    if ((sl === '' || sl === null) && (productName === '' || productName === null)) {
      continue;
    }
    var obj = {};
    for (var c = 0; c < HEADERS.length; c++) {
      var val = formatCellValue(row[c]);
      if (HEADERS[c] === 'stagingRoom' && val !== null) {
        val = String(val);
      }
      obj[HEADERS[c]] = val;
    }
    var cids = obj.containerIds ? String(obj.containerIds) : '';
    obj.containerIdList = cids
      ? cids.split(',').map(function (s) { return s.trim(); }).filter(Boolean)
      : [];
    entries.push(obj);
  }

  // Compute product summary, room summary & validity breakdown dynamically from entries
  var productMap = {};
  var roomMap = {};
  var validityBreakdown = { Expired: 0, Valid: 0, 'N/A': 0 };
  var totalContainers = 0;

  entries.forEach(function (e) {
    var p = e.productName || 'Unknown';
    var rm = String(e.stagingRoom || '115');
    if (!productMap[p]) {
      productMap[p] = {
        productName: p,
        entryCount: 0,
        containerCount: 0,
        expiredCount: 0,
        entries: []
      };
    }
    if (!roomMap[rm]) {
      roomMap[rm] = {
        stagingRoom: rm,
        entryCount: 0,
        containerCount: 0,
        validCount: 0,
        expiredCount: 0,
        naCount: 0,
        entries: []
      };
    }

    var containers = Number(e.noOfContainers) || 0;
    totalContainers += containers;
    productMap[p].entryCount += 1;
    productMap[p].containerCount += containers;

    roomMap[rm].entryCount += 1;
    roomMap[rm].containerCount += containers;

    var status = String(e.validityStatus || '').trim();
    if (status.toLowerCase() === 'expired') {
      productMap[p].expiredCount += 1;
      roomMap[rm].expiredCount += 1;
      validityBreakdown.Expired += 1;
    } else if (status.toLowerCase() === 'valid') {
      roomMap[rm].validCount += 1;
      validityBreakdown.Valid += 1;
    } else if (status.toUpperCase() === 'N/A') {
      roomMap[rm].naCount += 1;
      validityBreakdown['N/A'] += 1;
    } else if (status) {
      validityBreakdown[status] = (validityBreakdown[status] || 0) + 1;
    }
    productMap[p].entries.push(e);
    roomMap[rm].entries.push(e);
  });

  var byProduct = Object.keys(productMap)
    .map(function (k) {
      return {
        product: productMap[k].productName,
        entries: productMap[k].entryCount,
        containers: productMap[k].containerCount,
        expired: productMap[k].expiredCount
      };
    })
    .sort(function (a, b) {
      return b.entries - a.entries || a.product.localeCompare(b.product);
    });

  var sortedRooms = Object.keys(roomMap).sort();
  var byRoom = sortedRooms.map(function (rm) {
    return {
      stagingRoom: rm,
      entries: roomMap[rm].entryCount,
      containers: roomMap[rm].containerCount,
      valid: roomMap[rm].validCount,
      expired: roomMap[rm].expiredCount,
      na: roomMap[rm].naCount
    };
  });

  var totalExpired = byProduct.reduce(function (acc, x) { return acc + x.expired; }, 0);
  var nowIso = new Date().toISOString().replace(/\.\d{3}Z$/, 'Z');

  var endpointsMap = {
    all: 'data/all.json',
    notExited: 'data/not-exited.json',
    summary: 'data/summary.json',
    expired: 'data/expired.json',
    valid: 'data/valid.json',
    na: 'data/na.json',
    products: 'data/products.json',
    rooms: 'data/rooms.json',
    meta: 'data/meta.json',
    notExitedCsv: 'data/not-exited.csv',
    summaryCsv: 'data/summary.csv'
  };
  sortedRooms.forEach(function (rm) {
    endpointsMap['room' + rm] = 'data/room-' + rm + '.json';
  });

  var meta = {
    title: title,
    subtitle: subtitle,
    facility: 'Renata PLC, Rajendrapur Potent Product Facility',
    lastSyncedAt: nowIso,
    source: sourceLabel || 'google-sheets-apps-script',
    spreadsheetId: ss.getId(),
    spreadsheetUrl: ss.getUrl(),
    rooms: sortedRooms,
    totalEntries: entries.length,
    totalContainers: totalContainers,
    validityBreakdown: validityBreakdown,
    headers: rawHeaders,
    endpoints: endpointsMap
  };

  var expired = entries.filter(function (e) {
    return String(e.validityStatus || '').toLowerCase() === 'expired';
  });
  var valid = entries.filter(function (e) {
    return String(e.validityStatus || '').toLowerCase() === 'valid';
  });
  var naEntries = entries.filter(function (e) {
    return String(e.validityStatus || '').toUpperCase() === 'N/A';
  });

  var summaryPayload = {
    meta: {
      title: meta.title,
      subtitle: meta.subtitle,
      lastSyncedAt: nowIso,
      totalEntries: entries.length,
      totalContainers: totalContainers
    },
    byRoom: byRoom,
    byProduct: byProduct,
    total: {
      product: 'TOTAL',
      entries: entries.length,
      containers: totalContainers,
      expired: totalExpired
    },
    validityBreakdown: validityBreakdown
  };

  var allPayload = {
    meta: meta,
    summary: summaryPayload,
    entries: entries
  };

  var roomFiles = {};
  sortedRooms.forEach(function (rm) {
    roomFiles['data/room-' + rm + '.json'] = {
      lastSyncedAt: nowIso,
      stagingRoom: rm,
      count: roomMap[rm].entryCount,
      containerCount: roomMap[rm].containerCount,
      entries: roomMap[rm].entries
    };
  });

  return {
    meta: meta,
    summary: summaryPayload,
    all: allPayload,
    notExited: { meta: meta, count: entries.length, entries: entries },
    expired: { lastSyncedAt: nowIso, status: 'Expired', count: expired.length, entries: expired },
    valid: { lastSyncedAt: nowIso, status: 'Valid', count: valid.length, entries: valid },
    na: { lastSyncedAt: nowIso, status: 'N/A', count: naEntries.length, entries: naEntries },
    products: {
      lastSyncedAt: nowIso,
      productCount: Object.keys(productMap).length,
      products: Object.keys(productMap).map(function (k) { return productMap[k]; })
    },
    rooms: {
      lastSyncedAt: nowIso,
      roomCount: sortedRooms.length,
      rooms: sortedRooms.map(function (rm) { return roomMap[rm]; })
    },
    roomFiles: roomFiles,
    notExitedCsv: toCsv(
      [rawHeaders.map(function (h) { return h.replace(/\r?\n/g, ' '); })].concat(
        entries.map(function (e) {
          return HEADERS.map(function (k) { return e[k] === null ? '' : e[k]; });
        })
      )
    ),
    summaryCsv: toCsv(
      [['Product', 'Entries', 'Containers', 'Expired (past validity)']]
        .concat(byProduct.map(function (p) { return [p.product, p.entries, p.containers, p.expired]; }))
        .concat([['TOTAL', entries.length, totalContainers, totalExpired]])
    )
  };
}

function toCsv(rows) {
  return rows.map(function (row) {
    return row.map(function (cell) {
      var s = cell === null || cell === undefined ? '' : String(cell);
      if (/[",\n\r]/.test(s)) {
        return '"' + s.replace(/"/g, '""') + '"';
      }
      return s;
    }).join(',');
  }).join('\n') + '\n';
}

// ============================================================================
// GITHUB ATOMIC COMMIT SYNC (Git Data API)
// ============================================================================

function syncToGitHub(triggerSource) {
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);

  try {
    var props = PropertiesService.getScriptProperties();
    var token = props.getProperty('GITHUB_TOKEN');
    if (!token) {
      throw new Error('GITHUB_TOKEN is not set in Script Properties.');
    }
    var owner = props.getProperty('GITHUB_OWNER') || DEFAULT_OWNER;
    var repo = props.getProperty('GITHUB_REPO') || DEFAULT_REPO;
    var branch = props.getProperty('GITHUB_BRANCH') || DEFAULT_BRANCH;

    var dataset = buildDataset('google-sheets:' + (triggerSource || 'sync'));

    var files = [
      { path: 'data/meta.json', content: JSON.stringify(dataset.meta, null, 2) },
      { path: 'data/not-exited.json', content: JSON.stringify(dataset.notExited, null, 2) },
      { path: 'data/summary.json', content: JSON.stringify(dataset.summary, null, 2) },
      { path: 'data/all.json', content: JSON.stringify(dataset.all, null, 2) },
      { path: 'data/expired.json', content: JSON.stringify(dataset.expired, null, 2) },
      { path: 'data/valid.json', content: JSON.stringify(dataset.valid, null, 2) },
      { path: 'data/na.json', content: JSON.stringify(dataset.na, null, 2) },
      { path: 'data/products.json', content: JSON.stringify(dataset.products, null, 2) },
      { path: 'data/rooms.json', content: JSON.stringify(dataset.rooms, null, 2) },
      { path: 'data/not-exited.csv', content: dataset.notExitedCsv },
      { path: 'data/summary.csv', content: dataset.summaryCsv }
    ];
    Object.keys(dataset.roomFiles || {}).forEach(function (rmPath) {
      files.push({ path: rmPath, content: JSON.stringify(dataset.roomFiles[rmPath], null, 2) });
    });

    var baseApi = 'https://api.github.com/repos/' + owner + '/' + repo;

    // 1. Get branch ref
    var refData = ghRequest('GET', baseApi + '/git/ref/heads/' + branch, token);
    var headSha = refData.object.sha;

    // 2. Get head commit tree
    var commitData = ghRequest('GET', baseApi + '/git/commits/' + headSha, token);
    var baseTreeSha = commitData.tree.sha;

    // 3. Create new tree with all updated files
    var treeEntries = files.map(function (f) {
      return {
        path: f.path,
        mode: '100644',
        type: 'blob',
        content: f.content
      };
    });

    var newTree = ghRequest('POST', baseApi + '/git/trees', token, {
      base_tree: baseTreeSha,
      tree: treeEntries
    });

    // 4. Create commit
    var commitMsg = 'chore(sync): update staging data from Google Sheets (' +
      dataset.meta.totalEntries + ' entries, ' + (triggerSource || 'sync') + ')';
    var newCommit = ghRequest('POST', baseApi + '/git/commits', token, {
      message: commitMsg,
      tree: newTree.sha,
      parents: [headSha]
    });

    // 5. Update branch ref
    ghRequest('PATCH', baseApi + '/git/refs/heads/' + branch, token, {
      sha: newCommit.sha,
      force: false
    });

    props.setProperties({
      LAST_SYNC_AT: dataset.meta.lastSyncedAt,
      LAST_COMMIT_SHA: newCommit.sha
    });

    return {
      ok: true,
      commitSha: newCommit.sha,
      totalEntries: dataset.meta.totalEntries,
      totalContainers: dataset.meta.totalContainers
    };
  } finally {
    lock.releaseLock();
  }
}

function ghRequest(method, url, token, payload) {
  var options = {
    method: method,
    muteHttpExceptions: true,
    headers: {
      Authorization: 'Bearer ' + token,
      Accept: 'application/vnd.github+json',
      'X-GitHub-Api-Version': '2022-11-28',
      'User-Agent': 'Google-Apps-Script-Staging-Sync'
    }
  };
  if (payload !== undefined) {
    options.contentType = 'application/json';
    options.payload = JSON.stringify(payload);
  }
  var resp = UrlFetchApp.fetch(url, options);
  var code = resp.getResponseCode();
  var body = resp.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error('GitHub API ' + method + ' ' + url + ' failed (' + code + '): ' + body);
  }
  return JSON.parse(body);
}

// ============================================================================
// OPTIONAL LIVE WEB APP REST ENDPOINTS (doGet / doPost)
// ============================================================================

function doGet(e) {
  var params = (e && e.parameter) || {};
  var endpoint = (params.endpoint || params.action || 'all').toLowerCase();
  var format = (params.format || 'json').toLowerCase();
  var dataset = buildDataset('apps-script-webapp');

  if (format === 'csv') {
    var csvText = endpoint === 'summary' ? dataset.summaryCsv : dataset.notExitedCsv;
    return ContentService.createTextOutput(csvText).setMimeType(ContentService.MimeType.CSV);
  }

  var payload;
  switch (endpoint) {
    case 'meta':
      payload = dataset.meta;
      break;
    case 'summary':
      payload = dataset.summary;
      break;
    case 'expired':
      payload = dataset.expired;
      break;
    case 'valid':
      payload = dataset.valid;
      break;
    case 'na':
      payload = dataset.na;
      break;
    case 'products':
      payload = dataset.products;
      break;
    case 'rooms':
      payload = dataset.rooms;
      break;
    case 'room':
      var rmKey = String(params.room || params.q || '').trim();
      var rmObj = (dataset.rooms.rooms || []).filter(function (r) {
        return String(r.stagingRoom) === rmKey;
      })[0];
      payload = rmObj || { stagingRoom: rmKey, entryCount: 0, containerCount: 0, entries: [] };
      break;
    case 'not-exited':
    case 'not_exited':
      var list = dataset.notExited.entries;
      if (params.room) {
        var rmFilter = String(params.room).trim();
        list = list.filter(function (x) { return String(x.stagingRoom || '') === rmFilter; });
      }
      if (params.status) {
        var st = params.status.toLowerCase();
        list = list.filter(function (x) {
          return String(x.validityStatus || '').toLowerCase() === st;
        });
      }
      if (params.room || params.status) {
        payload = { count: list.length, room: params.room || null, status: params.status || null, entries: list };
      } else {
        payload = dataset.notExited;
      }
      break;
    case 'search':
      var q = (params.q || '').toLowerCase().trim();
      var matches = dataset.notExited.entries.filter(function (x) {
        return JSON.stringify(x).toLowerCase().indexOf(q) !== -1;
      });
      payload = { query: params.q || '', count: matches.length, entries: matches };
      break;
    case 'batch':
      var batch = (params.batch || params.q || '').toLowerCase().trim();
      var batchMatches = dataset.notExited.entries.filter(function (x) {
        return String(x.batchNo || '').toLowerCase() === batch;
      });
      payload = { batch: params.batch || '', count: batchMatches.length, entries: batchMatches };
      break;
    case 'container':
      var cid = (params.id || params.q || '').toLowerCase().trim();
      var cidMatches = dataset.notExited.entries.filter(function (x) {
        return (x.containerIdList || []).some(function (c) {
          return c.toLowerCase() === cid;
        });
      });
      payload = { containerId: params.id || '', count: cidMatches.length, entries: cidMatches };
      break;
    case 'all':
    default:
      payload = dataset.all;
      break;
  }

  return ContentService.createTextOutput(JSON.stringify(payload, null, 2))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var action = (body.action || 'sync').toLowerCase();

    if (action === 'sync') {
      var res = syncToGitHub('webapp-post');
      return jsonOut(res);
    }

    var ss = getSpreadsheet();
    var ws = ss.getSheetByName(SHEET_NOT_EXITED);

    if (action === 'add' && body.entry) {
      var lastRow = ws.getLastRow();
      var nextSl = Number(ws.getRange(lastRow, 1).getValue()) + 1 || (lastRow - 3);
      var entry = body.entry;
      entry.sl = entry.sl || nextSl;
      var row = HEADERS.map(function (k) { return entry[k] !== undefined && entry[k] !== null ? entry[k] : ''; });
      ws.appendRow(row);
      var syncRes = syncToGitHub('webapp-add');
      return jsonOut({ ok: true, added: entry, sync: syncRes });
    }

    return jsonOut({ ok: false, error: 'Unsupported action: ' + action });
  } catch (err) {
    return jsonOut({ ok: false, error: err.message });
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj, null, 2))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * One-time helper: populates the active Google Sheet from GitHub's data/all.json
 * with both "Not Exited" and "Summary" tabs, headers, styling, and formulas.
 */
function bootstrapSheetFromGitHub() {
  var owner = PropertiesService.getScriptProperties().getProperty('GITHUB_OWNER') || DEFAULT_OWNER;
  var repo = PropertiesService.getScriptProperties().getProperty('GITHUB_REPO') || DEFAULT_REPO;
  var branch = PropertiesService.getScriptProperties().getProperty('GITHUB_BRANCH') || DEFAULT_BRANCH;
  var url = 'https://raw.githubusercontent.com/' + owner + '/' + repo + '/' + branch + '/data/all.json';

  var resp = UrlFetchApp.fetch(url);
  var data = JSON.parse(resp.getContentText());
  var ss = getSpreadsheet();
  ss.rename('Staging_Not_Exited_26-11-2026');

  // 1. Setup "Not Exited" sheet
  var ws = ss.getSheetByName(SHEET_NOT_EXITED) || ss.getSheets()[0];
  ws.setName(SHEET_NOT_EXITED);
  ws.clear();

  var title = data.meta.title || 'Log Sheet for Staging Dispensed Material and Intermediate Product';
  var subtitle = data.meta.subtitle || '';
  var rawHeaders = data.meta.headers || HEADERS;

  ws.getRange(1, 1).setValue(title).setFontWeight('bold').setFontSize(13);
  ws.getRange(2, 1).setValue(subtitle).setFontColor('#555555').setFontSize(10);

  var colCount = HEADERS.length;
  var headerRange = ws.getRange(4, 1, 1, colCount);
  headerRange.setValues([rawHeaders.slice(0, colCount)])
    .setFontWeight('bold')
    .setBackground('#1e293b')
    .setFontColor('#ffffff')
    .setWrap(true);

  var rows = (data.entries || []).map(function (e) {
    return HEADERS.map(function (k) {
      return e[k] === null || e[k] === undefined ? '' : e[k];
    });
  });

  if (rows.length > 0) {
    ws.getRange(5, 1, rows.length, colCount).setValues(rows);
    // Color-code Staging Room (col 2) and Validity Status (col 11)
    for (var i = 0; i < rows.length; i++) {
      var rm = String(rows[i][1] || '');
      var rmCell = ws.getRange(5 + i, 2);
      if (rm === '126') {
        rmCell.setBackground('#ede9fe').setFontColor('#5b21b6').setFontWeight('bold');
      } else if (rm === '119') {
        rmCell.setBackground('#ccfbf1').setFontColor('#115e59').setFontWeight('bold');
      } else if (rm === '115') {
        rmCell.setBackground('#e0f2fe').setFontColor('#075985').setFontWeight('bold');
      }

      var st = String(rows[i][10] || '').toLowerCase();
      var cell = ws.getRange(5 + i, 11);
      if (st === 'expired') {
        cell.setBackground('#fee2e2').setFontColor('#991b1b').setFontWeight('bold');
      } else if (st === 'valid') {
        cell.setBackground('#dcfce7').setFontColor('#166534').setFontWeight('bold');
      } else if (st === 'n/a') {
        cell.setBackground('#fef3c7').setFontColor('#92400e').setFontWeight('bold');
      }
    }
  }
  ws.setFrozenRows(4);
  ws.autoResizeColumns(1, colCount);

  // 2. Setup "Summary" sheet
  var wsSum = ss.getSheetByName(SHEET_SUMMARY) || ss.insertSheet(SHEET_SUMMARY);
  wsSum.clear();
  var roomLabel = ((data.meta && data.meta.rooms) || ['115', '119', '126']).join(', ');
  wsSum.getRange(1, 1).setValue('Summary — Entries Not Yet Exited (Rooms ' + roomLabel + ')').setFontWeight('bold').setFontSize(13);
  wsSum.getRange(3, 1, 1, 4)
    .setValues([['Product', 'Entries', 'Containers', 'Expired (past validity)']])
    .setFontWeight('bold')
    .setBackground('#1e293b')
    .setFontColor('#ffffff');

  var byProd = (data.summary && data.summary.byProduct) || [];
  var sumRows = byProd.map(function (p) {
    return [p.product, p.entries, p.containers, p.expired];
  });
  if (sumRows.length > 0) {
    wsSum.getRange(4, 1, sumRows.length, 4).setValues(sumRows);
  }
  var totalRowIdx = 4 + sumRows.length;
  var total = (data.summary && data.summary.total) || { product: 'TOTAL', entries: 0, containers: 0, expired: 0 };
  wsSum.getRange(totalRowIdx, 1, 1, 4)
    .setValues([[total.product, total.entries, total.containers, total.expired]])
    .setFontWeight('bold')
    .setBackground('#e2e8f0');

  var vb = (data.meta && data.meta.validityBreakdown) || {};
  var vbStart = totalRowIdx + 2;
  wsSum.getRange(vbStart, 1).setValue('Validity Status').setFontWeight('bold');
  wsSum.getRange(vbStart + 1, 1, 3, 2).setValues([
    ['Expired', vb.Expired || 0],
    ['Valid', vb.Valid || 0],
    ['N/A', vb['N/A'] || 0]
  ]);

  var byRoom = (data.summary && data.summary.byRoom) || [];
  if (byRoom.length > 0) {
    var rmStart = vbStart + 5;
    wsSum.getRange(rmStart, 1, 1, 6)
      .setValues([['Staging Room', 'Entries', 'Containers', 'Valid', 'Expired', 'N/A']])
      .setFontWeight('bold')
      .setBackground('#1e293b')
      .setFontColor('#ffffff');
    var rmRows = byRoom.map(function (r) {
      return ['Room ' + r.stagingRoom, r.entries, r.containers, r.valid, r.expired, r.na];
    });
    wsSum.getRange(rmStart + 1, 1, rmRows.length, 6).setValues(rmRows);
  }
  wsSum.autoResizeColumns(1, 6);

  return 'Bootstrapped ' + rows.length + ' entries into ' + ss.getName();
}
