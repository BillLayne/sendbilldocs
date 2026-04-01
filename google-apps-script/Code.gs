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
  const subject = 'We received your documents (' + submission.confirmationNumber + ')';
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
  const fileList = savedFiles.map(function(file) {
    return '<li>' + htmlEscape_(file.name) + ' (' + htmlEscape_(formatFileSize_(file.sizeBytes)) + ')</li>';
  }).join('');

  return [
    '<div style="font-family:Arial,sans-serif;color:#1f2937;line-height:1.6;">',
    '<h2 style="margin:0 0 12px;">We Received Your Documents</h2>',
    '<p>Thank you for sending your files to ' + htmlEscape_(runtime.fromName) + '.</p>',
    '<p><strong>Confirmation #:</strong> ' + htmlEscape_(submission.confirmationNumber) + '<br>',
    '<strong>Document Type:</strong> ' + htmlEscape_(submission.docType) + '<br>',
    '<strong>Files Received:</strong> ' + savedFiles.length + '</p>',
    '<ul>' + fileList + '</ul>',
    '<p>We will review your documents and reach out if we need anything else.</p>',
    '<p>If you have questions, call us at ' + htmlEscape_(APP_CONFIG.officePhone) + ' or visit <a href="' + htmlEscape_(APP_CONFIG.officeWebsite) + '">' + htmlEscape_(APP_CONFIG.officeWebsite) + '</a>.</p>',
    '</div>'
  ].join('');
}

function buildCustomerPlainBody_(runtime, submission, savedFiles) {
  const fileLines = savedFiles.map(function(file) {
    return '- ' + file.name + ' (' + formatFileSize_(file.sizeBytes) + ')';
  }).join('\n');

  return [
    'We received your documents.',
    '',
    'Confirmation #: ' + submission.confirmationNumber,
    'Document Type: ' + submission.docType,
    'Files Received: ' + savedFiles.length,
    '',
    fileLines,
    '',
    'Thank you for sending your files to ' + runtime.fromName + '.',
    'If you have questions, call us at ' + APP_CONFIG.officePhone + '.'
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
