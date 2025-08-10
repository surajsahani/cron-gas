/**
 * Push Sync (Sheet -> Backend) integrated with TriggerScheduler
 * Edit CONFIG.BACKEND_URL and SHEET_NAME for your sheet/backend.
 */

var CONFIG = {
  BACKEND_URL: 'https://ng-campus-pulse.onrender.com/api/import-data',
  TIMEOUT_MS: 30000,
  MAX_RETRIES: 3,
  RETRY_DELAY_MS: 2000
};

var SHEET_NAME = 'Form Responses 1';
var LOG_SHEET_NAME = 'Sync Logs';

/**
 * Main syncing function — reads sheet rows, validates, injects timestamp and posts to backend
 */
function pushDataToBackend() {
  var startTime = new Date();
  log('INFO', 'Starting sync at ' + startTime.toISOString());

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      log('ERROR', 'Sheet not found: ' + SHEET_NAME);
      return { success: false, message: 'Sheet not found: ' + SHEET_NAME };
    }

    var values = sheet.getDataRange().getValues();
    if (!values || values.length <= 1) {
      log('WARN', 'No data rows found');
      return { success: false, message: 'No data rows found' };
    }

    var headers = values[0];
    var rows = values.slice(1);

    // Map header → index for stable access
    var headerIndex = {};
    headers.forEach(function (h, i) { headerIndex[String(h).trim()] = i; });

    // Required columns (loose matching)
    var required = ['Choose the campus you are referring to', 'Name', 'Email Address'];
    var missing = [];
    required.forEach(function (r) {
      var found = Object.keys(headerIndex).some(function (h) { return h.indexOf(r.substring(0, 10)) !== -1; });
      if (!found) missing.push(r);
    });
    if (missing.length) {
      log('ERROR', 'Missing required columns: ' + missing.join(', '));
      return { success: false, message: 'Missing required columns: ' + missing.join(', ') };
    }

    var data = [];
    var valid = 0, skipped = 0;

    rows.forEach(function (row, idx) {
      // build object using headers
      var obj = {};
      headers.forEach(function (h, j) {
        obj[String(h).trim()] = row[j] === undefined ? '' : row[j];
      });

      var campus = obj['Choose the campus you are referring to'] || obj['Choose the campus you are referring to '] || '';
      var name = obj['Name'] || obj['Name '] || '';

      if (!String(campus).trim() || !String(name).trim()) {
        skipped++;
        return;
      }

      // add sync timestamp so backend recognizes it's new
      obj._sync_timestamp = new Date().toISOString();

      data.push(obj);
      valid++;
    });

    log('INFO', 'Rows processed: ' + (rows.length) + ' — valid: ' + valid + ' skipped: ' + skipped);

    if (data.length === 0) {
      log('WARN', 'No valid data to send');
      return { success: false, message: 'No valid data to send' };
    }

    // send data
    var result = sendDataWithRetry(data);
    var endTime = new Date();
    log('INFO', 'Sync finished in ' + ((endTime - startTime) / 1000) + 's — success: ' + !!result.success);
    return result;

  } catch (e) {
    log('ERROR', 'pushDataToBackend exception: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

/**
 * Wrapper used by scheduler when you want hybrid timing.
 * This is the function you point TriggerScheduler at as 'targetFunc'
 */
function scheduledPush() {
  // simple wrapper to call pushDataToBackend and log
  var res = pushDataToBackend();
  if (res && res.success) {
    log('INFO', 'Push successful');
  } else {
    log('ERROR', 'Push failed: ' + (res && res.message ? res.message : 'unknown'));
  }
}

/**
 * Send to backend with retry
 */
function sendDataWithRetry(payload) {
  for (var attempt = 1; attempt <= CONFIG.MAX_RETRIES; attempt++) {
    try {
      log('INFO', 'Sending data to backend (attempt ' + attempt + ') - ' + payload.length + ' records');
      var response = UrlFetchApp.fetch(CONFIG.BACKEND_URL, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({ records: payload }),
        muteHttpExceptions: true
      });
      var code = response.getResponseCode();
      var text = response.getContentText();
      log('INFO', 'Backend responded ' + code);

      if (code >= 200 && code < 300) {
        // try parse JSON if present
        try {
          var body = JSON.parse(text || '{}');
          log('INFO', 'Backend response parsed');
          return { success: true, data: body };
        } catch (parseErr) {
          return { success: true, message: 'Data sent (non-JSON response)' };
        }
      } else {
        throw new Error('HTTP ' + code + ': ' + text);
      }
    } catch (e) {
      log('ERROR', 'Attempt ' + attempt + ' failed: ' + e.toString());
      if (attempt < CONFIG.MAX_RETRIES) {
        Utilities.sleep(CONFIG.RETRY_DELAY_MS);
      } else {
        return { success: false, message: e.toString() };
      }
    }
  }
}

/**
 * Logging helper: append to a 'Sync Logs' sheet
 */
function log(level, message) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var logSheet = ss.getSheetByName(LOG_SHEET_NAME);
    if (!logSheet) {
      logSheet = ss.insertSheet(LOG_SHEET_NAME);
      logSheet.appendRow(['timestamp', 'level', 'message']);
    }
    logSheet.appendRow([new Date(), level, message]);
  } catch (e) {
    // if logging to sheet fails, fallback to console
    console.log(level + ': ' + message);
  }
}

/**
 * Convenience: setup triggers (create hybrid schedule)
 * Call this once after deploying to create triggers.
 */
function setupTriggers() {
  TriggerScheduler.createHybridTrigger({
    targetFunc: 'scheduledPush',
    peakStart: 6,
    peakEnd: 23, // exclusive, so runs until 22:59
    peakIntervalMinutes: 10,
    offPeakIntervalMinutes: 60
  });
  log('INFO', 'setupTriggers executed');
}

/**
 * Remove triggers for scheduledPush
 */
function removeTriggers() {
  TriggerScheduler.deleteHybridTrigger('scheduledPush');
  log('INFO', 'removeTriggers executed');
}
