/**
 * ========================================
 * ESATTO CUSTOM CONTACT - GOOGLE APPS SCRIPT
 * ========================================
 * Sheet Name: Esatto
 * Headers: Date and Time, Full Name, Email, Contact Number, Source, Other
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
    message: 'Use POST for contact submissions.'
  });
}

function handleRequest(e) {
  var lock = LockService.getScriptLock();
  var acquired = lock.tryLock(10000);

  if (!acquired) {
    Logger.log('Unable to acquire script lock for contact submission.');
    return jsonResponse({
      result: 'error',
      message: 'Temporary server lock. Please try again.'
    });
  }

  try {
    // --- STEP 1: Parse request data ---
    // The browser sends the form fields, including the reCAPTCHA token,
    // to this endpoint. The token must be verified here before any lead is saved.
    var p = e.parameter || {};
    
    // Attempt to parse JSON body if parameters are empty (fallback)
    if (Object.keys(p).length === 0 && e.postData && e.postData.contents) {
      try {
        p = JSON.parse(e.postData.contents);
      } catch (err) { }
    }

    // Original Contact Submission routing
    const data = {
      name: p.name || '',
      email: p.email || '',
      phone: p.phone || '',
      source: p.source || '',
      other: p.otherSourceText || '',
      recaptchaToken: p.recaptchaToken || ''
    };

    var recaptchaResult = verifyRecaptchaToken_(data.recaptchaToken, e);
    if (!recaptchaResult.ok) {
      Logger.log('Rejected contact submission after reCAPTCHA verification failed: ' + recaptchaResult.message);
      return jsonResponse({
        result: 'error',
        message: 'reCAPTCHA verification failed.'
      });
    }

    const SPREADSHEET_ID = '1h5lTyYvmkbnUbKj77yyQb6RgRYzxBHpXTGvaSRMEw2s';
    const TARGET_SHEET_NAME = 'Esatto';

    var doc = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = doc.getSheetByName(TARGET_SHEET_NAME);

    if (!sheet) {
      throw new Error(`Sheet "${TARGET_SHEET_NAME}" not found. Please run setup() or create the tab.`);
    }

    var rowData = [
      new Date(),        // Column A: Date and Time
      data.name,         // Column B: Full Name
      data.email,        // Column C: Email
      data.phone,        // Column D: Contact Number
      data.source,       // Column E: Source
      data.other         // Column F: Other
    ];

    sheet.appendRow(rowData);

    try {
      // SMS Notification (Short)
      MailApp.sendEmail("9496162100@vtext.com", "New Lead", "New Lead: " + data.name + ". Check email for details.");

      // Email Notification (Detailed)
      var emailBody = "New Lead Details:\n\n" +
                      "Name: " + data.name + "\n" +
                      "Email: " + data.email + "\n" +
                      "Phone: " + data.phone + "\n" +
                      "Source: " + data.source + "\n" +
                      "Other: " + data.other;

      MailApp.sendEmail("brandon@fdbespoke.com,9496162100@vtext.com", "New Lead on Esatto Site: " + data.name, emailBody);
    } catch (e) {
      Logger.log("Error sending notification: " + e.toString());
    }

    return jsonResponse({ 'result': 'success' });

  } catch (error) {
    Logger.log('Contact submission error: ' + error.toString());
    return jsonResponse({ 'result': 'error', 'message': 'Unable to process submission.' });
  } finally {
    lock.releaseLock();
  }
}

function verifyRecaptchaToken_(token, e) {
  // Security note: the site key is public, but the secret key must stay server-side.
  // This verification blocks direct POSTs that bypass the frontend.
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

    // This site is using reCAPTCHA v3, so the score/action checks are required.
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
// Run this once to create the "Esatto" tab and headers
function setup() {
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = "Esatto";
  var sheet = doc.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = doc.insertSheet(sheetName);
  }
  
  var headers = [
    "Date and Time",
    "Full Name", 
    "Email", 
    "Contact Number", 
    "Source",
    "Other"
  ];
  
  // Only set headers if the first row is empty
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
}
