/**
 * ========================================
 * ESATTO CUSTOM GIVEAWAY - GOOGLE APPS SCRIPT
 * ========================================
 * Sheet Name: Esatto Giveaway
 * Spreadsheet ID: 1h5lTyYvmkbnUbKj77yyQb6RgRYzxBHpXTGvaSRMEw2s
 * Headers: Full Name, Mobile Number, Who Referred You?
 *
 * Store the reCAPTCHA values in Script Properties using these names:
 * RECAPTCHA_SITE_KEY, RECAPTCHA_SECRET_KEY, RECAPTCHA_ALLOWED_HOSTNAMES,
 * RECAPTCHA_SCORE_THRESHOLD.
 */

function doPost(e) {
  return handleRequest(e);
}

function doGet(e) {
  return jsonResponse({
    result: 'error',
    message: 'Use POST for giveaway submissions.'
  });
}

function handleRequest(e) {
  var lock = LockService.getScriptLock();
  var acquired = lock.tryLock(10000);

  if (!acquired) {
    Logger.log('Unable to acquire script lock for giveaway submission.');
    return jsonResponse({
      result: 'error',
      message: 'Temporary server lock. Please try again.'
    });
  }

  try {
    var p = e.parameter || {};
    
    // Attempt to parse JSON body if parameters are empty (fallback)
    if (Object.keys(p).length === 0 && e.postData && e.postData.contents) {
      try {
        p = JSON.parse(e.postData.contents);
      } catch (err) { }
    }

    const data = {
      name: p.name || '',
      phone: p.phone || '',
      referredBy: p.referredBy || '',
      shoeColor: p.shoeColor || p.shoePreference || '',
      recaptchaToken: p.recaptchaToken || ''
    };

    var recaptchaResult = verifyRecaptchaToken_(data.recaptchaToken, e);
    if (!recaptchaResult.ok) {
      Logger.log('Rejected giveaway submission after reCAPTCHA verification failed: ' + recaptchaResult.message);
      return jsonResponse({
        result: 'error',
        message: 'reCAPTCHA verification failed.'
      });
    }

    const SPREADSHEET_ID = '1h5lTyYvmkbnUbKj77yyQb6RgRYzxBHpXTGvaSRMEw2s';
    const TARGET_SHEET_NAME = 'Esatto Giveaway';

    var doc = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = doc.getSheetByName(TARGET_SHEET_NAME);

    if (!sheet) {
      sheet = doc.insertSheet(TARGET_SHEET_NAME);
      var headers = [
        "Full Name",
        "Mobile Number",
        "Who Referred You?",
        "Shoe Color"
      ];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }

    var rowData = [
      data.name,
      data.phone,
      data.referredBy,
      data.shoeColor
    ];

    sheet.appendRow(rowData);

    // Send Email Notification (skipping SMS for high-volume giveaway entries)
    try {
      var emailBody = "New Custom Shoe Giveaway Entry Details:\n\n" +
                      "Name: " + data.name + "\n" +
                      "Phone: " + data.phone + "\n" +
                      "Referred By: " + data.referredBy + "\n" +
                      "Shoe Color: " + data.shoeColor;

      MailApp.sendEmail("brandon@fdbespoke.com", "New Giveaway Entry: " + data.name, emailBody);
    } catch (err) {
      Logger.log("Error sending giveaway notification: " + err.toString());
    }

    return jsonResponse({ 'result': 'success' });

  } catch (error) {
    Logger.log('Giveaway submission error: ' + error.toString());
    return jsonResponse({ 'result': 'error', 'message': 'Unable to process submission.' });
  } finally {
    lock.releaseLock();
  }
}

function verifyRecaptchaToken_(token, e) {
  if (!token) {
    return {
      ok: false,
      message: 'Missing reCAPTCHA token.'
    };
  }

  var props = PropertiesService.getScriptProperties();
  var secretKey = props.getProperty('RECAPTCHA_SECRET_KEY');
  var siteKey = props.getProperty('RECAPTCHA_SITE_KEY');
  var allowedHostnames = getAllowedHostnames_();
  var scoreThreshold = Number(props.getProperty('RECAPTCHA_SCORE_THRESHOLD') || 0.5);

  if (!siteKey || !secretKey) {
    Logger.log('reCAPTCHA configuration is missing. SITE_KEY present: ' + !!siteKey + ', SECRET_KEY present: ' + !!secretKey);
    return {
      ok: false,
      message: 'reCAPTCHA configuration is incomplete.'
    };
  }

  try {
    var response = UrlFetchApp.fetch('https://www.google.com/recaptcha/api/siteverify', {
      method: 'post',
      payload: {
        secret: secretKey,
        response: token
      },
      muteHttpExceptions: true
    });

    var responseCode = response.getResponseCode();
    var responseText = response.getContentText() || '{}';

    if (responseCode !== 200) {
      Logger.log('reCAPTCHA siteverify returned HTTP ' + responseCode + ': ' + responseText);
      return {
        ok: false,
        message: 'Verification service error.'
      };
    }

    var verification = JSON.parse(responseText);
    var hostname = verification.hostname || '';
    var score = typeof verification.score === 'number' ? verification.score : null;
    var action = verification.action || '';

    if (!verification.success) {
      Logger.log('reCAPTCHA rejected token. Errors: ' + JSON.stringify(verification['error-codes'] || []));
      return {
        ok: false,
        message: 'Token verification failed.'
      };
    }

    if (score === null) {
      Logger.log('Expected a reCAPTCHA v3 score but none was returned: ' + responseText);
      return {
        ok: false,
        message: 'Unexpected reCAPTCHA response.'
      };
    }

    if (action !== 'submit') {
      Logger.log('Unexpected reCAPTCHA action: ' + action);
      return {
        ok: false,
        message: 'Unexpected reCAPTCHA action.'
      };
    }

    if (score < scoreThreshold) {
      Logger.log('Low reCAPTCHA score: ' + score + ' threshold: ' + scoreThreshold + ' hostname: ' + hostname);
      return {
        ok: false,
        message: 'Low reCAPTCHA score.'
      };
    }

    if (allowedHostnames.length && allowedHostnames.indexOf(hostname) === -1) {
      Logger.log('reCAPTCHA hostname mismatch. Received: ' + hostname + ' Allowed: ' + allowedHostnames.join(', '));
      return {
        ok: false,
        message: 'Hostname mismatch.'
      };
    }

    return {
      ok: true,
      message: 'Verified successfully.'
    };
  } catch (error) {
    Logger.log('reCAPTCHA verification request failed: ' + error.toString());
    return {
      ok: false,
      message: 'Verification request failed.'
    };
  }
}

function getAllowedHostnames_() {
  var props = PropertiesService.getScriptProperties();
  var rawHostnames = props.getProperty('RECAPTCHA_ALLOWED_HOSTNAMES') || 'esattocustom.com,www.esattocustom.com';

  return rawHostnames
    .split(',')
    .map(function(hostname) {
      return hostname.trim();
    })
    .filter(function(hostname) {
      return hostname.length > 0;
    });
}

function jsonResponse(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

// --- SETUP FUNCTION ---
// Run this once to initialize the sheet and headers
function setup() {
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = "Esatto Giveaway";
  var sheet = doc.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = doc.insertSheet(sheetName);
  }
  
  var headers = [
    "Full Name",
    "Mobile Number",
    "Who Referred You?",
    "Shoe Color"
  ];
  
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
}
