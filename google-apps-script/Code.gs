const APP_CONFIG = Object.freeze({
  agencyName: 'Bill Layne Insurance Agency',
  officePhone: '(336) 835-1993',
  officeWebsite: 'https://www.billlayneinsurance.com',
  defaultTimeZone: 'America/New_York',
  maxEmailAttachmentBytes: 18 * 1024 * 1024,
  maxSingleAttachmentBytes: 7 * 1024 * 1024
});

function doGet(e) {
  const payload = {
    ok: true,
    app: 'sendbilldocs',
    version: '2026-04-01',
    timestamp: new Date().toISOString()
  };

  const callback = e && e.parameter ? e.parameter.callback : '';
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + JSON.stringify(payload) + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return jsonResponse_(payload);
}

function doPost(e) {
  try {
    const runtime = getRuntimeConfig_();
    const submission = normalizeSubmission_(parseRequestBody_(e));
    const folder = createSubmissionFolder_(runtime.rootFolder, submission);
    const savedFiles = saveFiles_(folder, submission.files);

    const officeEmail = sendOfficeNotification_(runtime, submission, folder, savedFiles);
    const customerEmail = sendCustomerConfirmation_(runtime, submission, folder, savedFiles);

    return jsonResponse_({
      ok: true,
      confirmationNumber: submission.confirmationNumber,
      driveFolderUrl: folder.getUrl(),
      fileCount: savedFiles.length,
      officeEmailSent: officeEmail.sent,
      customerEmailSent: customerEmail.sent
    });
  } catch (error) {
    console.error('SendBillDocs backend failure: ' + (error && error.stack ? error.stack : error));
    return jsonResponse_({
      ok: false,
      error: error && error.message ? error.message : String(error)
    });
  }
}

function parseRequestBody_(e) {
  const contents = e && e.postData && e.postData.contents ? e.postData.contents : '';
  if (!contents) {
    throw new Error('Missing POST body.');
  }

  try {
    return JSON.parse(contents);
  } catch (error) {
    throw new Error('Could not parse JSON payload.');
  }
}

function normalizeSubmission_(payload) {
  if (!payload || typeof payload !== 'object') {
    throw new Error('Payload must be an object.');
  }

  const files = Array.isArray(payload.files) ? payload.files : [];
  if (!files.length) {
    throw new Error('At least one file is required.');
  }

  const submission = {
    name: cleanText_(payload.name) || 'Unknown Customer',
    email: cleanText_(payload.email),
    phone: cleanText_(payload.phone),
    docType: cleanText_(payload.docType) || 'Unspecified',
    notes: cleanText_(payload.notes),
    confirmationNumber: cleanText_(payload.confirmationNumber) || generateFallbackConfirmation_(),
    timestamp: cleanText_(payload.timestamp) || new Date().toISOString(),
    files: []
  };

  if (!submission.email) {
    throw new Error('Customer email address is required.');
  }

  submission.files = files.map(function(file, index) {
    if (!file || !file.data) {
      throw new Error('File #' + (index + 1) + ' is missing base64 data.');
    }

    const fileName = sanitizeFileName_(cleanText_(file.name) || ('upload-' + (index + 1)));
    const mimeType = inferMimeType_(cleanText_(file.type), fileName);
    const byteSize = Number(file.size) > 0 ? Number(file.size) : estimateByteSize_(file.data);

    return {
      name: fileName,
      mimeType: mimeType,
      sizeBytes: byteSize,
      base64Data: String(file.data)
    };
  });

  return submission;
}

function getRuntimeConfig_() {
  const props = PropertiesService.getScriptProperties();
  const rootFolderId = cleanText_(props.getProperty('UPLOADS_ROOT_FOLDER_ID'));
  if (!rootFolderId) {
    throw new Error('Missing script property: UPLOADS_ROOT_FOLDER_ID');
  }

  const officeEmails = String(props.getProperty('OFFICE_EMAILS') || 'billlayneinsurance@gmail.com')
    .split(',')
    .map(function(value) { return cleanText_(value); })
    .filter(Boolean);

  if (!officeEmails.length) {
    throw new Error('Missing script property: OFFICE_EMAILS');
  }

  return {
    rootFolder: DriveApp.getFolderById(rootFolderId),
    officeEmails: officeEmails,
    fromName: cleanText_(props.getProperty('FROM_NAME')) || APP_CONFIG.agencyName,
    customerReplyTo: cleanText_(props.getProperty('CUSTOMER_REPLY_TO')) || officeEmails[0],
    timeZone: cleanText_(props.getProperty('TIMEZONE')) || Session.getScriptTimeZone() || APP_CONFIG.defaultTimeZone
  };
}

function createSubmissionFolder_(rootFolder, submission) {
  const submissionDate = new Date(submission.timestamp);
  const safeDate = isNaN(submissionDate.getTime()) ? new Date() : submissionDate;
  const yearMonth = Utilities.formatDate(safeDate, APP_CONFIG.defaultTimeZone, 'yyyy-MM');
  const day = Utilities.formatDate(safeDate, APP_CONFIG.defaultTimeZone, 'yyyy-MM-dd');

  const monthFolder = findOrCreateFolder_(rootFolder, yearMonth);
  const dayFolder = findOrCreateFolder_(monthFolder, day);

  const folderName = [
    submission.confirmationNumber,
    sanitizeFileName_(submission.name).replace(/\.[^.]+$/, '')
  ].join(' - ');

  return dayFolder.createFolder(folderName);
}

function saveFiles_(folder, files) {
  return files.map(function(file) {
    const bytes = Utilities.base64Decode(file.base64Data);
    const blob = Utilities.newBlob(bytes, file.mimeType, file.name);
    const driveFile = folder.createFile(blob);

    return {
      name: file.name,
      mimeType: file.mimeType,
      sizeBytes: file.sizeBytes || bytes.length,
      blob: blob,
      driveFile: driveFile,
      url: driveFile.getUrl()
    };
  });
}

function sendOfficeNotification_(runtime, submission, folder, savedFiles) {
  const attachmentPlan = buildAttachmentPlan_(savedFiles);
  const subject = '[SendBillDocs] ' + submission.docType + ' from ' + submission.name + ' (' + submission.confirmationNumber + ')';
  const htmlBody = buildOfficeHtmlBody_(runtime, submission, folder, savedFiles, attachmentPlan.skipped);
  const plainBody = buildOfficePlainBody_(submission, folder, savedFiles, attachmentPlan.skipped);

  MailApp.sendEmail({
    to: runtime.officeEmails[0],
    cc: runtime.officeEmails.slice(1).join(','),
    subject: subject,
    name: runtime.fromName,
    replyTo: submission.email,
    htmlBody: htmlBody,
    body: plainBody,
    attachments: attachmentPlan.attachments
  });

  return { sent: true };
}

function sendCustomerConfirmation_(runtime, submission, folder, savedFiles) {
  const subject = '✅ Photos Received — ' + submission.confirmationNumber + ' | Bill Layne Insurance';
  const htmlBody = buildCustomerHtmlBody_(runtime, submission, savedFiles);
  const plainBody = buildCustomerPlainBody_(runtime, submission, savedFiles);

  MailApp.sendEmail({
    to: submission.email,
    subject: subject,
    name: runtime.fromName,
    replyTo: runtime.customerReplyTo,
    htmlBody: htmlBody,
    body: plainBody
  });

  return { sent: true };
}

function buildAttachmentPlan_(savedFiles) {
  let runningBytes = 0;
  const attachments = [];
  const skipped = [];

  savedFiles.forEach(function(file) {
    const tooLargeForSingleAttachment = file.sizeBytes > APP_CONFIG.maxSingleAttachmentBytes;
    const tooLargeForRunningTotal = runningBytes + file.sizeBytes > APP_CONFIG.maxEmailAttachmentBytes;

    if (tooLargeForSingleAttachment || tooLargeForRunningTotal) {
      skipped.push(file);
      return;
    }

    attachments.push(file.blob.copyBlob().setName(file.name));
    runningBytes += file.sizeBytes;
  });

  return {
    attachments: attachments,
    skipped: skipped
  };
}

function buildOfficeHtmlBody_(runtime, submission, folder, savedFiles, skippedFiles) {
  const fileItems = savedFiles.map(function(file) {
    return '<li><a href="' + htmlEscape_(file.url) + '">' + htmlEscape_(file.name) + '</a> (' + htmlEscape_(formatFileSize_(file.sizeBytes)) + ')</li>';
  }).join('');

  const skippedNote = skippedFiles.length
    ? '<p><strong>Note:</strong> ' + skippedFiles.length + ' file(s) were too large to attach to the email, but all uploads are saved in Drive.</p>'
    : '';

  return [
    '<div style="font-family:Arial,sans-serif;color:#1f2937;line-height:1.6;">',
    '<h2 style="margin:0 0 12px;">New SendBillDocs Upload</h2>',
    '<p><strong>Confirmation #:</strong> ' + htmlEscape_(submission.confirmationNumber) + '<br>',
    '<strong>Name:</strong> ' + htmlEscape_(submission.name) + '<br>',
    '<strong>Email:</strong> <a href="mailto:' + htmlEscape_(submission.email) + '">' + htmlEscape_(submission.email) + '</a><br>',
    '<strong>Phone:</strong> ' + htmlEscape_(submission.phone || 'Not provided') + '<br>',
    '<strong>Document Type:</strong> ' + htmlEscape_(submission.docType) + '<br>',
    '<strong>Submitted:</strong> ' + htmlEscape_(submission.timestamp) + '</p>',
    '<p><strong>Notes:</strong><br>' + htmlEscape_(submission.notes || 'None') + '</p>',
    '<p><strong>Drive Folder:</strong> <a href="' + htmlEscape_(folder.getUrl()) + '">' + htmlEscape_(folder.getUrl()) + '</a></p>',
    skippedNote,
    '<p><strong>Files:</strong></p>',
    '<ul>' + fileItems + '</ul>',
    '</div>'
  ].join('');
}

function buildOfficePlainBody_(submission, folder, savedFiles, skippedFiles) {
  const fileLines = savedFiles.map(function(file) {
    return '- ' + file.name + ' (' + formatFileSize_(file.sizeBytes) + '): ' + file.url;
  }).join('\n');

  const skippedNote = skippedFiles.length
    ? '\n\nNote: ' + skippedFiles.length + ' file(s) were too large to attach to the email, but all uploads are saved in Drive.'
    : '';

  return [
    'New SendBillDocs upload',
    '',
    'Confirmation #: ' + submission.confirmationNumber,
    'Name: ' + submission.name,
    'Email: ' + submission.email,
    'Phone: ' + (submission.phone || 'Not provided'),
    'Document Type: ' + submission.docType,
    'Submitted: ' + submission.timestamp,
    '',
    'Notes:',
    submission.notes || 'None',
    '',
    'Drive Folder:',
    folder.getUrl(),
    '',
    'Files:',
    fileLines,
    skippedNote
  ].join('\n');
}

function buildCustomerHtmlBody_(runtime, submission, savedFiles) {
  var fileRows = savedFiles.map(function(file) {
    var icon = '&#128196;';
    var lower = file.name.toLowerCase();
    if (lower.match(/\.(jpg|jpeg|png|gif|webp|heic|heif|bmp|tiff)$/)) icon = '&#128247;';
    else if (lower.match(/\.pdf$/)) icon = '&#128203;';
    return '<tr><td style="padding:8px 12px;border-bottom:1px solid #f1f5f9;font-family:Arial,sans-serif;font-size:14px;color:#334155;">' + icon + '&nbsp;&nbsp;' + htmlEscape_(file.name) + '</td><td style="padding:8px 12px;border-bottom:1px solid #f1f5f9;font-family:Arial,sans-serif;font-size:13px;color:#94a3b8;text-align:right;white-space:nowrap;">' + htmlEscape_(formatFileSize_(file.sizeBytes)) + '</td></tr>';
  }).join('');

  var localTime = '';
  try {
    var d = new Date(submission.timestamp);
    localTime = Utilities.formatDate(d, runtime.timeZone || APP_CONFIG.defaultTimeZone, "MMMM d, yyyy 'at' h:mm a");
  } catch(e) {
    localTime = submission.timestamp;
  }

  return [
    '<!DOCTYPE html>',
    '<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1.0"></head>',
    '<body style="margin:0;padding:0;background-color:#f1f5f9;-webkit-text-size-adjust:100%;-ms-text-size-adjust:100%;">',
    '<!--[if mso]><table role="presentation" width="600" align="center" cellpadding="0" cellspacing="0" border="0"><tr><td><![endif]-->',
    '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="max-width:600px;margin:0 auto;">',

    '<tr><td style="padding:0;">',
    '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="background:linear-gradient(135deg,#003f87 0%,#0076d3 100%);border-radius:0 0 16px 16px;">',
    '<tr><td style="padding:36px 30px 28px;text-align:center;">',
    '<img src="https://i.imgur.com/lxu9nfT.png" alt="Bill Layne Insurance" width="180" style="display:block;margin:0 auto 16px;max-width:180px;height:auto;">',
    '<table role="presentation" cellpadding="0" cellspacing="0" border="0" style="margin:0 auto;"><tr><td style="background-color:rgba(255,255,255,0.15);border-radius:20px;padding:6px 16px;"><span style="font-family:Arial,sans-serif;font-size:13px;color:rgba(255,255,255,0.9);letter-spacing:0.3px;">&#10003; Photos Received Successfully</span></td></tr></table>',
    '</td></tr>',
    '</table>',
    '</td></tr>',

    '<tr><td style="padding:20px 16px 0;">',
    '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color:#ffffff;border-radius:16px;box-shadow:0 1px 3px rgba(0,0,0,0.08);">',

    '<tr><td style="padding:28px 28px 0;">',
    '<p style="margin:0 0 4px;font-family:Arial,sans-serif;font-size:22px;font-weight:700;color:#0f2744;">Hi ' + htmlEscape_(submission.name.split(' ')[0]) + ',</p>',
    '<p style="margin:0;font-family:Arial,sans-serif;font-size:15px;color:#64748b;line-height:1.5;">We have your photos. Our team will review them and reach out if we need anything else.</p>',
    '</td></tr>',

    '<tr><td style="padding:20px 28px 0;">',
    '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color:#f0f9ff;border-radius:12px;border:1px solid #bae6fd;">',
    '<tr><td style="padding:16px 20px;">',
    '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">',
    '<tr><td style="font-family:Arial,sans-serif;font-size:11px;font-weight:700;color:#0369a1;text-transform:uppercase;letter-spacing:0.5px;padding-bottom:8px;">Confirmation Details</td></tr>',
    '<tr><td style="padding-bottom:6px;"><table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%"><tr><td style="font-family:Arial,sans-serif;font-size:14px;color:#64748b;">Confirmation #</td><td style="font-family:Arial,sans-serif;font-size:14px;font-weight:700;color:#0f2744;text-align:right;">' + htmlEscape_(submission.confirmationNumber) + '</td></tr></table></td></tr>',
    '<tr><td style="padding-bottom:6px;"><table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%"><tr><td style="font-family:Arial,sans-serif;font-size:14px;color:#64748b;">Document Type</td><td style="font-family:Arial,sans-serif;font-size:14px;color:#0f2744;text-align:right;">' + htmlEscape_(submission.docType) + '</td></tr></table></td></tr>',
    '<tr><td><table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%"><tr><td style="font-family:Arial,sans-serif;font-size:14px;color:#64748b;">Received</td><td style="font-family:Arial,sans-serif;font-size:14px;color:#0f2744;text-align:right;">' + htmlEscape_(localTime) + '</td></tr></table></td></tr>',
    '</table>',
    '</td></tr>',
    '</table>',
    '</td></tr>',

    '<tr><td style="padding:20px 28px 0;">',
    '<p style="margin:0 0 10px;font-family:Arial,sans-serif;font-size:13px;font-weight:700;color:#0f2744;text-transform:uppercase;letter-spacing:0.5px;">Files Received (' + savedFiles.length + ')</p>',
    '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color:#f8fafc;border-radius:10px;overflow:hidden;">',
    fileRows,
    '</table>',
    '</td></tr>',

    '<tr><td style="padding:24px 28px 0;">',
    '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color:#f0fdf4;border-radius:12px;border:1px solid #bbf7d0;">',
    '<tr><td style="padding:16px 20px;">',
    '<p style="margin:0 0 8px;font-family:Arial,sans-serif;font-size:13px;font-weight:700;color:#166534;text-transform:uppercase;letter-spacing:0.5px;">What Happens Next</p>',
    '<table role="presentation" cellpadding="0" cellspacing="0" border="0">',
    '<tr><td style="padding:3px 0;font-family:Arial,sans-serif;font-size:14px;color:#334155;line-height:1.5;">&#10004;&#65039; Your photos are securely stored</td></tr>',
    '<tr><td style="padding:3px 0;font-family:Arial,sans-serif;font-size:14px;color:#334155;line-height:1.5;">&#128269; Our team will review your submission</td></tr>',
    '<tr><td style="padding:3px 0;font-family:Arial,sans-serif;font-size:14px;color:#334155;line-height:1.5;">&#128222; We will contact you if anything else is needed</td></tr>',
    '</table>',
    '</td></tr>',
    '</table>',
    '</td></tr>',

    '<tr><td style="padding:24px 28px;">',
    '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">',
    '<tr><td style="text-align:center;">',
    '<a href="' + htmlEscape_(APP_CONFIG.officeWebsite) + '" target="_blank" style="display:inline-block;background-color:#0076d3;color:#ffffff;font-family:Arial,sans-serif;font-size:15px;font-weight:700;text-decoration:none;padding:14px 36px;border-radius:12px;">Visit Our Website</a>',
    '</td></tr>',
    '</table>',
    '</td></tr>',

    '</table>',
    '</td></tr>',

    '<tr><td style="padding:20px 16px 0;">',
    '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color:#0f172a;border-radius:16px;">',
    '<tr><td style="padding:28px 28px 20px;text-align:center;">',
    '<img src="https://i.imgur.com/lxu9nfT.png" alt="Bill Layne Insurance" width="140" style="display:block;margin:0 auto 12px;max-width:140px;height:auto;opacity:0.9;">',
    '<p style="margin:0 0 4px;font-family:Arial,sans-serif;font-size:14px;color:#e2e8f0;">Bill Layne Insurance Agency</p>',
    '<p style="margin:0 0 4px;font-family:Arial,sans-serif;font-size:13px;color:#94a3b8;">1283 N Bridge St, Elkin, NC 28621</p>',
    '<p style="margin:0 0 12px;font-family:Arial,sans-serif;font-size:13px;color:#94a3b8;"><a href="tel:3368351993" style="color:#60a5fa;text-decoration:none;">(336) 835-1993</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href="mailto:docs@billlayneinsurance.com" style="color:#60a5fa;text-decoration:none;">docs@billlayneinsurance.com</a></p>',
    '<table role="presentation" cellpadding="0" cellspacing="0" border="0" style="margin:0 auto;"><tr>',
    '<td style="padding:0 6px;"><a href="https://www.billlayneinsurance.com" style="color:#60a5fa;font-family:Arial,sans-serif;font-size:12px;text-decoration:none;">Website</a></td>',
    '<td style="color:#475569;font-size:12px;">|</td>',
    '<td style="padding:0 6px;"><a href="https://www.facebook.com/dollarbillagency" style="color:#60a5fa;font-family:Arial,sans-serif;font-size:12px;text-decoration:none;">Facebook</a></td>',
    '<td style="color:#475569;font-size:12px;">|</td>',
    '<td style="padding:0 6px;"><a href="https://billlayneinsurance.com/get-quote" style="color:#60a5fa;font-family:Arial,sans-serif;font-size:12px;text-decoration:none;">Get a Quote</a></td>',
    '</tr></table>',
    '</td></tr>',
    '<tr><td style="padding:0 28px 20px;text-align:center;"><p style="margin:0;font-family:Arial,sans-serif;font-size:11px;color:#475569;">&copy; 2026 Bill Layne Insurance Agency. All rights reserved.</p></td></tr>',
    '</table>',
    '</td></tr>',

    '<tr><td style="padding:20px 0;">&nbsp;</td></tr>',

    '</table>',
    '<!--[if mso]></td></tr></table><![endif]-->',
    '</body></html>'
  ].join('');
}

function buildCustomerPlainBody_(runtime, submission, savedFiles) {
  var fileLines = savedFiles.map(function(file) {
    return '  - ' + file.name + ' (' + formatFileSize_(file.sizeBytes) + ')';
  }).join('\n');

  var localTime = '';
  try {
    var d = new Date(submission.timestamp);
    localTime = Utilities.formatDate(d, runtime.timeZone || APP_CONFIG.defaultTimeZone, "MMMM d, yyyy 'at' h:mm a");
  } catch(e) {
    localTime = submission.timestamp;
  }

  return [
    'Hi ' + submission.name.split(' ')[0] + ',',
    '',
    'We received your photos! Our team will review them and reach out if we need anything else.',
    '',
    '--- CONFIRMATION DETAILS ---',
    'Confirmation #: ' + submission.confirmationNumber,
    'Document Type: ' + submission.docType,
    'Received: ' + localTime,
    '',
    '--- FILES RECEIVED (' + savedFiles.length + ') ---',
    fileLines,
    '',
    '--- WHAT HAPPENS NEXT ---',
    '  * Your photos are securely stored',
    '  * Our team will review your submission',
    '  * We will contact you if anything else is needed',
    '',
    'Questions? Call us at ' + APP_CONFIG.officePhone,
    'Or visit ' + APP_CONFIG.officeWebsite,
    '',
    'Bill Layne Insurance Agency',
    '1283 N Bridge St, Elkin, NC 28621',
    APP_CONFIG.officePhone,
    'docs@billlayneinsurance.com'
  ].join('\n');
}

function findOrCreateFolder_(parentFolder, childName) {
  const matches = parentFolder.getFoldersByName(childName);
  return matches.hasNext() ? matches.next() : parentFolder.createFolder(childName);
}

function sanitizeFileName_(name) {
  return String(name || 'upload')
    .replace(/[\\/:*?"<>|#%&{}$!'@+=`]/g, '-')
    .replace(/\s+/g, ' ')
    .trim();
}

function inferMimeType_(declaredType, fileName) {
  if (declaredType) {
    return declaredType;
  }

  const lowerName = String(fileName || '').toLowerCase();
  const extension = lowerName.indexOf('.') > -1 ? lowerName.split('.').pop() : '';
  const mimeTypes = {
    pdf: 'application/pdf',
    jpg: 'image/jpeg',
    jpeg: 'image/jpeg',
    png: 'image/png',
    gif: 'image/gif',
    bmp: 'image/bmp',
    tif: 'image/tiff',
    tiff: 'image/tiff',
    webp: 'image/webp',
    heic: 'image/heic',
    heif: 'image/heif',
    doc: 'application/msword',
    docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
  };

  return mimeTypes[extension] || 'application/octet-stream';
}

function estimateByteSize_(base64Data) {
  const normalized = String(base64Data || '').replace(/\s/g, '');
  const padding = normalized.endsWith('==') ? 2 : normalized.endsWith('=') ? 1 : 0;
  return Math.max(0, Math.floor((normalized.length * 3) / 4) - padding);
}

function formatFileSize_(bytes) {
  if (!bytes || bytes < 1024) return bytes + ' B';
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
  return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
}

function cleanText_(value) {
  return String(value || '').trim();
}

function generateFallbackConfirmation_() {
  const now = new Date();
  const datePart = Utilities.formatDate(now, APP_CONFIG.defaultTimeZone, 'yyMMdd');
  const randomPart = Math.random().toString(36).slice(2, 6).toUpperCase();
  return 'BL-' + datePart + '-' + randomPart;
}

function htmlEscape_(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function jsonResponse_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
