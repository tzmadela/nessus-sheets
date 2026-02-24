require('dotenv').config();
var axios = require('axios');
var google = require('googleapis').google;
var https = require('https');
var readline = require('readline');

var CONFIG = {
  nessusUrl: process.env.NESSUS_URL,
  accessKey: process.env.NESSUS_ACCESS_KEY,
  secretKey: process.env.NESSUS_SECRET_KEY,
  spreadsheetId: process.env.SPREADSHEET_ID,
  credentialsPath: process.env.GOOGLE_CREDENTIALS_PATH || './credentials.json'
};

function checkConfig() {
  var missing = [];
  if (!CONFIG.nessusUrl) missing.push('NESSUS_URL');
  if (!CONFIG.accessKey) missing.push('NESSUS_ACCESS_KEY');
  if (!CONFIG.secretKey) missing.push('NESSUS_SECRET_KEY');
  if (!CONFIG.spreadsheetId) missing.push('SPREADSHEET_ID');
  if (missing.length > 0) {
    console.log('\nMissing in .env:');
    missing.forEach(function(m) { console.log('  - ' + m); });
    process.exit(1);
  }
}

var nessusClient = axios.create({
  baseURL: CONFIG.nessusUrl,
  headers: {
    'X-ApiKeys': 'accessKey=' + CONFIG.accessKey + '; secretKey=' + CONFIG.secretKey,
    'Content-Type': 'application/json'
  },
  httpsAgent: new https.Agent({ rejectUnauthorized: false }),
  timeout: 30000
});

async function nessusGet(endpoint) {
  try {
    var r = await nessusClient.get(endpoint);
    return r.data;
  } catch (e) {
    console.error('Nessus Error (' + endpoint + '):', e.message);
    return null;
  }
}

async function getSheets() {
  var auth = new google.auth.GoogleAuth({
    keyFile: CONFIG.credentialsPath,
    scopes: ['https://www.googleapis.com/auth/spreadsheets']
  });
  return google.sheets({ version: 'v4', auth });
}

async function getExistingSheets(sheets) {
  var sp = await sheets.spreadsheets.get({ spreadsheetId: CONFIG.spreadsheetId });
  return sp.data.sheets;
}

async function clearDataSheets(sheets) {
  console.log('Clearing data sheets (keeping Run Log)...');
  try {
    var sheetList = await getExistingSheets(sheets);
    var requests = [];
    var hasRunLog = false;
    var hasAtLeastOne = false;

    for (var i = 0; i < sheetList.length; i++) {
      var title = sheetList[i].properties.title;
      var sid = sheetList[i].properties.sheetId;

      if (title === 'Run Log') {
        hasRunLog = true;
        hasAtLeastOne = true;
        continue;
      }

      if (!hasAtLeastOne) {
        requests.push({
          updateSheetProperties: {
            properties: {
              sheetId: sid,
              title: 'Sheet1',
              gridProperties: { rowCount: 1000, columnCount: 15 }
            },
            fields: 'title,gridProperties(rowCount,columnCount)'
          }
        });
        requests.push({
          updateCells: { range: { sheetId: sid }, fields: 'userEnteredValue' }
        });
        hasAtLeastOne = true;
      } else {
        requests.push({ deleteSheet: { sheetId: sid } });
      }
    }

    if (requests.length > 0) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: CONFIG.spreadsheetId,
        resource: { requests: requests }
      });
    }

    console.log('  Cleared data sheets' + (hasRunLog ? ' (Run Log preserved)' : ''));
  } catch (e) {
    console.error('  Error clearing sheets:', e.message);
  }
}

async function createSheet(sheets, sheetName, rows) {
  var rowCount = Math.min(Math.max(Math.ceil(rows * 1.5), 1000), 50000);
  try {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: CONFIG.spreadsheetId,
      resource: {
        requests: [{
          addSheet: {
            properties: {
              title: sheetName,
              gridProperties: { rowCount: rowCount, columnCount: 15 }
            }
          }
        }]
      }
    });
    console.log('  Created "' + sheetName + '" (' + rowCount + ' rows)');
  } catch (e) {
    if (e.message.includes('already exists')) {
      try {
        var sheetList = await getExistingSheets(sheets);
        var sheet = sheetList.find(function(s) { return s.properties.title === sheetName; });
        if (sheet) {
          var sid = sheet.properties.sheetId;
          var reqs = [
            { updateCells: { range: { sheetId: sid }, fields: 'userEnteredValue' } }
          ];
          if (rowCount > sheet.properties.gridProperties.rowCount) {
            reqs.push({
              updateSheetProperties: {
                properties: {
                  sheetId: sid,
                  gridProperties: { rowCount: rowCount, columnCount: 15 }
                },
                fields: 'gridProperties(rowCount,columnCount)'
              }
            });
          }
          await sheets.spreadsheets.batchUpdate({
            spreadsheetId: CONFIG.spreadsheetId,
            resource: { requests: reqs }
          });
          console.log('  Cleared "' + sheetName + '"');
        }
      } catch (e2) { console.error('  Error resizing:', e2.message); }
    } else {
      console.error('  Error creating sheet:', e.message);
    }
  }
}

async function expandIfNeeded(sheets, sheetName, neededRows) {
  var target = Math.min(Math.max(Math.ceil(neededRows * 1.5), 1000), 200000);
  try {
    var sheetList = await getExistingSheets(sheets);
    var sheet = sheetList.find(function(s) { return s.properties.title === sheetName; });
    if (!sheet) return;
    if (target > sheet.properties.gridProperties.rowCount) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: CONFIG.spreadsheetId,
        resource: {
          requests: [{
            updateSheetProperties: {
              properties: {
                sheetId: sheet.properties.sheetId,
                gridProperties: { rowCount: target, columnCount: 15 }
              },
              fields: 'gridProperties(rowCount,columnCount)'
            }
          }]
        }
      });
      console.log('  Expanded "' + sheetName + '" to ' + target + ' rows');
    }
  } catch (e) { console.error('  Expand error:', e.message); }
}

async function writeToSheet(sheets, sheetName, data) {
  await expandIfNeeded(sheets, sheetName, data.length + 100);
  var batchSize = 10000;
  for (var i = 0; i < data.length; i += batchSize) {
    var chunk = data.slice(i, i + batchSize);
    var startRow = i + 1;
    await sheets.spreadsheets.values.update({
      spreadsheetId: CONFIG.spreadsheetId,
      range: sheetName + '!A' + startRow,
      valueInputOption: 'RAW',
      resource: { values: chunk }
    });
    if (data.length > batchSize) {
      console.log('  Batch: rows ' + startRow + '-' + (startRow + chunk.length - 1));
    }
  }
  console.log('  Written ' + data.length + ' rows to "' + sheetName + '"');
}

async function addFilter(sheets, sheetName) {
  try {
    var sheetList = await getExistingSheets(sheets);
    var sheet = sheetList.find(function(s) { return s.properties.title === sheetName; });
    if (!sheet) return;
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: CONFIG.spreadsheetId,
      resource: {
        requests: [{
          setBasicFilter: {
            filter: {
              range: {
                sheetId: sheet.properties.sheetId,
                startRowIndex: 0,
                startColumnIndex: 0
              }
            }
          }
        }]
      }
    });
  } catch (e) {}
}

async function testConnection() {
  console.log('\nTesting Nessus connection...');
  var data = await nessusGet('/server/status');
  if (data) {
    console.log('Nessus connection successful!');
    console.log('Status:', JSON.stringify(data));
    return true;
  }
  console.log('Could not connect to Nessus');
  return false;
}

async function listScans() {
  console.log('\nFetching scans...');
  var data = await nessusGet('/scans');
  if (!data || !data.scans) { console.log('No scans found'); return; }

  console.log('\nAvailable Scans:');
  console.log('-'.repeat(70));
  data.scans.forEach(function(s) {
    console.log('  ID: ' + s.id + '  Name: ' + s.name + '  Status: ' + s.status);
  });
  console.log('-'.repeat(70));
  console.log('Total: ' + data.scans.length + ' scans');

  var sheets = await getSheets();
  await createSheet(sheets, 'Scans', data.scans.length + 5);
  var rows = [['Scan ID', 'Name', 'Status', 'Creation Date', 'Last Modified']];
  data.scans.forEach(function(s) {
    rows.push([
      s.id, s.name, s.status,
      s.creation_date ? new Date(s.creation_date * 1000).toISOString() : '',
      s.last_modification_date ? new Date(s.last_modification_date * 1000).toISOString() : ''
    ]);
  });
  await writeToSheet(sheets, 'Scans', rows);
}

async function listScanHistory(scanId) {
  console.log('\nFetching history for scan ' + scanId + '...');
  var data = await nessusGet('/scans/' + scanId);
  if (!data || !data.history) { console.log('No history found'); return null; }

  console.log('\nScan History:');
  console.log('-'.repeat(80));
  data.history.forEach(function(h) {
    var date = h.creation_date ? new Date(h.creation_date * 1000).toLocaleString() : 'N/A';
    console.log('  ' + h.history_id + '    ' + h.status + '    ' + date);
  });
  console.log('-'.repeat(80));
  return data.history;
}

async function pullHosts(scanId, historyId) {
  var endpoint = '/scans/' + scanId;
  if (historyId) {
    endpoint += '?history_id=' + historyId;
    console.log('\nPulling scan ' + scanId + ' (history: ' + historyId + ')...');
  } else {
    console.log('\nPulling scan ' + scanId + ' (latest)...');
  }

  var data = await nessusGet(endpoint);
  if (!data) { console.log('Could not fetch scan'); return; }

  var info = data.info || {};
  console.log('Scan: ' + (info.name || 'N/A'));
  console.log('Status: ' + (info.status || 'N/A'));
  console.log('Hosts: ' + (data.hosts ? data.hosts.length : 0));
  if (!data.hosts || data.hosts.length === 0) { console.log('No hosts.'); return; }

  var sheets = await getSheets();
  var startTime = Date.now();
  var totalHosts = data.hosts.length;

  // Clear data sheets but keep Run Log
  await clearDataSheets(sheets);

  // ---- HOSTS ----
  console.log('\nWriting hosts...');
  await createSheet(sheets, 'Hosts', totalHosts + 10);
  var hostData = [['Hostname / IP', 'Host ID', 'Critical', 'High', 'Medium', 'Low', 'Total', 'Risk Score']];

  data.hosts.forEach(function(host) {
    var c = host.critical || 0;
    var h = host.high || 0;
    var m = host.medium || 0;
    var l = host.low || 0;
    var score = c * 10 + h * 5 + m * 2 + l;
    hostData.push([
      host.hostname || host.host_ip || 'Unknown',
      host.host_id, c, h, m, l, c + h + m + l, score
    ]);
  });

  var hdr = hostData.shift();
  hostData.sort(function(a, b) { return b[7] - a[7]; });
  hostData.unshift(hdr);
  await writeToSheet(sheets, 'Hosts', hostData);
  await addFilter(sheets, 'Hosts');

  // ---- PREPARE VULN SHEETS ----
  var estVulns = totalHosts * 8;
  await createSheet(sheets, 'All Vulnerabilities', estVulns);
  await createSheet(sheets, 'Critical and High', Math.ceil(estVulns * 0.15));

  // ---- PULL VULNS ----
  console.log('\nPulling vulnerabilities for ' + totalHosts + ' hosts (excluding Info)...');
  console.log('Estimated time: ~' + Math.round(totalHosts * 0.6 / 60) + ' minutes\n');

  var sevNames = { 1: 'Low', 2: 'Medium', 3: 'High', 4: 'Critical' };
  var allVulns = [['Hostname', 'Host ID', 'OS', 'Plugin ID', 'Plugin Name', 'Severity', 'Severity Text', 'Plugin Family', 'Count']];
  var critHigh = [['Hostname', 'Host ID', 'OS', 'Plugin ID', 'Plugin Name', 'Severity Text', 'Plugin Family', 'Count']];
  var processed = 0;
  var failed = 0;
  var totalCrit = 0;
  var totalHighV = 0;
  var totalMed = 0;
  var totalLow = 0;
  var totalInfo = 0;

  for (var idx = 0; idx < data.hosts.length; idx++) {
    var host = data.hosts[idx];
    var hostId = host.host_id;
    var hostname = host.hostname || host.host_ip || 'Unknown';
    processed++;

    var pct = Math.round((processed / totalHosts) * 100);
    var elapsed = Math.round((Date.now() - startTime) / 1000);
    var rate = processed / (elapsed || 1);
    var remaining = Math.round((totalHosts - processed) / rate / 60);

    process.stdout.write(
      '\r  [' + pct + '%] ' + processed + '/' + totalHosts +
      ' | ' + hostname +
      ' | ~' + remaining + 'min left' +
      ' | vulns: ' + (allVulns.length - 1) +
      '                              '
    );

    var hEndpoint = '/scans/' + scanId + '/hosts/' + hostId;
    if (historyId) hEndpoint += '?history_id=' + historyId;

    var hResult = await nessusGet(hEndpoint);
    if (!hResult || !hResult.vulnerabilities) { failed++; continue; }

    var os = (hResult.info || {})['operating-system'] || 'Unknown';

    hResult.vulnerabilities.forEach(function(v) {
      var count = v.count || 1;

      // Count info but skip adding to sheet
      if (v.severity === 0) {
        totalInfo += count;
        return;
      }

      // Track counts
      if (v.severity === 4) totalCrit += count;
      if (v.severity === 3) totalHighV += count;
      if (v.severity === 2) totalMed += count;
      if (v.severity === 1) totalLow += count;

      // Add to all vulns sheet
      allVulns.push([
        hostname, hostId, os,
        v.plugin_id, v.plugin_name, v.severity,
        sevNames[v.severity] || 'Unknown',
        v.plugin_family || 'N/A', count
      ]);

      // Also add critical and high
      if (v.severity >= 3) {
        critHigh.push([
          hostname, hostId, os,
          v.plugin_id, v.plugin_name,
          sevNames[v.severity] || 'Unknown',
          v.plugin_family || 'N/A', count
        ]);
      }
    });

    await new Promise(function(r) { setTimeout(r, 500); });

    // Save progress every 100 hosts
    if (processed % 100 === 0) {
      console.log('\n  Saving progress... (' + allVulns.length + ' vulns)');
      await expandIfNeeded(sheets, 'All Vulnerabilities', allVulns.length + (totalHosts - processed) * 8);
      await writeToSheet(sheets, 'All Vulnerabilities', allVulns);
    }
  }

  // ---- WRITE ALL VULNS ----
  console.log('\n\nWriting all vulnerabilities (' + allVulns.length + ' rows)...');
  var vHdr = allVulns.shift();
  allVulns.sort(function(a, b) { return (b[5] || 0) - (a[5] || 0); });
  allVulns.unshift(vHdr);
  await writeToSheet(sheets, 'All Vulnerabilities', allVulns);
  await addFilter(sheets, 'All Vulnerabilities');

  // ---- WRITE CRIT+HIGH ----
  console.log('Writing Critical & High (' + critHigh.length + ' rows)...');
  var cHdr = critHigh.shift();
  critHigh.sort(function(a, b) {
    var o = { 'Critical': 0, 'High': 1 };
    return (o[a[5]] || 99) - (o[b[5]] || 99);
  });
  critHigh.unshift(cHdr);
  await writeToSheet(sheets, 'Critical and High', critHigh);
  await addFilter(sheets, 'Critical and High');

  // ---- DASHBOARD ----
  console.log('Creating dashboard...');
  var totalElapsed = Math.round((Date.now() - startTime) / 1000 / 60);
  var totalVulns = allVulns.length - 1;
  var totalAllCounts = totalCrit + totalHighV + totalMed + totalLow;

  await createSheet(sheets, 'Dashboard', 100);
  var dash = [
    ['NESSUS VULNERABILITY DASHBOARD'],
    [],
    ['Scan Name', info.name || 'N/A'],
    ['Scan ID', scanId],
    ['History ID', historyId || 'Latest'],
    ['Status', info.status || 'N/A'],
    ['Scan Completed', info.scan_end ? new Date(info.scan_end * 1000).toISOString() : 'N/A'],
    ['Data Pulled', new Date().toLocaleString()],
    ['Processing Time', totalElapsed + ' minutes'],
    ['Total Hosts', totalHosts],
    ['Failed Hosts', failed],
    ['Total Vulns (excl Info)', totalVulns],
    ['Critical + High', critHigh.length - 1],
    [],
    ['SEVERITY SUMMARY'],
    ['Severity', 'Count', 'Percentage'],
    ['Critical', totalCrit, totalAllCounts > 0 ? Math.round(totalCrit / totalAllCounts * 100) + '%' : '0%'],
    ['High', totalHighV, totalAllCounts > 0 ? Math.round(totalHighV / totalAllCounts * 100) + '%' : '0%'],
    ['Medium', totalMed, totalAllCounts > 0 ? Math.round(totalMed / totalAllCounts * 100) + '%' : '0%'],
    ['Low', totalLow, totalAllCounts > 0 ? Math.round(totalLow / totalAllCounts * 100) + '%' : '0%'],
    ['Info (excluded)', totalInfo, 'Not included in sheets'],
    [],
    ['TOP 30 HOSTS BY RISK'],
    ['Rank', 'Host', 'Critical', 'High', 'Medium', 'Low', 'Total', 'Risk Score']
  ];

  var topCount = Math.min(30, hostData.length - 1);
  for (var i = 1; i <= topCount; i++) {
    dash.push([i, hostData[i][0], hostData[i][2], hostData[i][3], hostData[i][4], hostData[i][5], hostData[i][6], hostData[i][7]]);
  }
  await writeToSheet(sheets, 'Dashboard', dash);

  // ---- RUN LOG (preserved across runs) ----
  console.log('Updating run log...');
  await createSheet(sheets, 'Run Log', 500);
  var now = new Date();
  var runType = process.argv.includes('--auto') ? 'Scheduled' : 'Manual';

  try {
    var ex = await sheets.spreadsheets.values.get({
      spreadsheetId: CONFIG.spreadsheetId,
      range: 'Run Log!A1:A1'
    });
    if (!ex.data.values || ex.data.values.length === 0) {
      await sheets.spreadsheets.values.update({
        spreadsheetId: CONFIG.spreadsheetId,
        range: 'Run Log!A1',
        valueInputOption: 'RAW',
        resource: {
          values: [['Date', 'Time', 'Scan ID', 'History', 'Hosts', 'Failed', 'Vulns', 'Crit', 'High', 'Med', 'Low', 'Info Skipped', 'Minutes', 'Type']]
        }
      });
    }
  } catch (e) {
    await sheets.spreadsheets.values.update({
      spreadsheetId: CONFIG.spreadsheetId,
      range: 'Run Log!A1',
      valueInputOption: 'RAW',
      resource: {
        values: [['Date', 'Time', 'Scan ID', 'History', 'Hosts', 'Failed', 'Vulns', 'Crit', 'High', 'Med', 'Low', 'Info Skipped', 'Minutes', 'Type']]
      }
    });
  }

  await sheets.spreadsheets.values.append({
    spreadsheetId: CONFIG.spreadsheetId,
    range: 'Run Log!A:N',
    valueInputOption: 'RAW',
    resource: {
      values: [[
        now.toLocaleDateString(),
        now.toLocaleTimeString(),
        scanId,
        historyId || 'Latest',
        totalHosts, failed, totalVulns,
        totalCrit, totalHighV, totalMed, totalLow, totalInfo,
        totalElapsed, runType
      ]]
    }
  });
  console.log('  Run logged');

  // ---- DONE ----
  console.log('\n' + '='.repeat(50));
  console.log('  ALL DONE!');
  console.log('='.repeat(50));
  console.log('  Time:            ' + totalElapsed + ' minutes');
  console.log('  Hosts:           ' + totalHosts + ' (' + failed + ' failed)');
  console.log('  Vulns (no info): ' + totalVulns);
  console.log('  Critical:        ' + totalCrit);
  console.log('  High:            ' + totalHighV);
  console.log('  Medium:          ' + totalMed);
  console.log('  Low:             ' + totalLow);
  console.log('  Info (excluded): ' + totalInfo);
  console.log('='.repeat(50));
  console.log('\nCheck your Google Sheet!');
}

async function autoRun() {
  checkConfig();
  var scanIds = process.argv.slice(3);
  if (scanIds.length === 0) {
    console.log('Usage: node nessus-to-sheets.js --auto 5 39 42');
    process.exit(1);
  }

  console.log('='.repeat(50));
  console.log('  NESSUS AUTO-RUN | ' + new Date().toLocaleString());
  console.log('  Scans: ' + scanIds.join(', '));
  console.log('='.repeat(50));

  var ok = await testConnection();
  if (!ok) { console.log('Cannot connect. Aborting.'); process.exit(1); }

  for (var i = 0; i < scanIds.length; i++) {
    console.log('\n' + '='.repeat(50));
    console.log('  Scan ' + scanIds[i] + ' (' + (i + 1) + '/' + scanIds.length + ')');
    console.log('='.repeat(50));
    try { await pullHosts(scanIds[i].trim(), null); }
    catch (e) { console.error('Error:', e.message); }
    if (i < scanIds.length - 1) {
      await new Promise(function(r) { setTimeout(r, 5000); });
    }
  }

  console.log('\n' + '='.repeat(50));
  console.log('  AUTO-RUN COMPLETE | ' + new Date().toLocaleString());
  console.log('='.repeat(50));
  process.exit(0);
}

var rl = readline.createInterface({ input: process.stdin, output: process.stdout });
function ask(q) { return new Promise(function(r) { rl.question(q, r); }); }

async function main() {
  checkConfig();
  if (process.argv.includes('--auto')) { await autoRun(); return; }

  console.log('='.repeat(50));
  console.log('  NESSUS TO GOOGLE SHEETS');
  console.log('='.repeat(50));
  console.log('  Nessus: ' + CONFIG.nessusUrl);
  console.log('  Sheet:  ' + CONFIG.spreadsheetId);
  console.log('  Info:   EXCLUDED');
  console.log('='.repeat(50));

  while (true) {
    console.log('\nOptions:');
    console.log('  1. Test Connection');
    console.log('  2. List Scans');
    console.log('  3. Pull Hosts & Vulns (Latest)');
    console.log('  4. Pull Hosts & Vulns (Historical)');
    console.log('  5. Exit');
    var c = await ask('\nChoose (1-5): ');
    switch (c.trim()) {
      case '1': await testConnection(); break;
      case '2': await listScans(); break;
      case '3':
        var sid = await ask('Scan ID: ');
        await pullHosts(sid.trim(), null);
        break;
      case '4':
        var hsid = await ask('Scan ID: ');
        var hist = await listScanHistory(hsid.trim());
        if (hist) {
          var hid = await ask('History ID: ');
          await pullHosts(hsid.trim(), hid.trim());
        }
        break;
      case '5':
        console.log('Bye!');
        rl.close();
        process.exit(0);
      default: console.log('Invalid.');
    }
  }
}

main().catch(console.error);
