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
  var subject = 'Photos Received - ' + submission.confirmationNumber + ' | Bill Layne Insurance';
  var htmlBody = buildCustomerHtmlBody_(runtime, submission, savedFiles);
  var plainBody = buildCustomerPlainBody_(runtime, submission, savedFiles);

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

// SENDBILLDOCS — ELITE GMAIL TEMPLATE
// buildCustomerPlainBody_ functions. Replace these two functions

function buildCustomerHtmlBody_(runtime, submission, savedFiles) {
  var fileRows = savedFiles.map(function(file) {
    var icon = '&#128196;';
    var lower = file.name.toLowerCase();
    if (lower.match(/\.(jpg|jpeg|png|gif|webp|heic|heif|bmp|tiff)$/)) icon = '&#128248;';
    else if (lower.match(/\.pdf$/)) icon = '&#128203;';
    return '<tr><td style="padding:14px 20px;border-bottom:1px solid #e2e8f0;"><table cellpadding="0" cellspacing="0" border="0" width="100%"><tr><td style="font-size:14px;color:#334155;font-family:\'Inter\',Arial,\'Helvetica Neue\',Helvetica,sans-serif;line-height:1.5;">' + icon + '&nbsp;&nbsp;' + htmlEscape_(file.name) + '</td><td align="right" style="font-size:13px;color:#64748b;font-family:\'Inter\',Arial,\'Helvetica Neue\',Helvetica,sans-serif;white-space:nowrap;">' + htmlEscape_(formatFileSize_(file.sizeBytes)) + '</td></tr></table></td></tr>';
  }).join('');
  // Remove border from last file row
  if (savedFiles.length > 0) {
    fileRows = fileRows.replace(/border-bottom:1px solid #e2e8f0;([^"]*$)/, '$1');
  }

  var localTime = '';
  try {
    var d = new Date(submission.timestamp);
    localTime = Utilities.formatDate(d, runtime.timeZone || APP_CONFIG.defaultTimeZone, "MMMM d, yyyy 'at' h:mm a");
  } catch(e) {
    localTime = submission.timestamp;
  }

  var firstName = (submission.name || 'there').split(' ')[0];
  var ff = "font-family:'Inter',Arial,'Helvetica Neue',Helvetica,sans-serif;";

  return [
    '<!DOCTYPE html><html lang="en" xmlns="http://www.w3.org/1999/xhtml" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><meta http-equiv="X-UA-Compatible" content="IE=edge"><meta name="x-apple-disable-message-reformatting"><meta name="format-detection" content="telephone=no"><title>Photos Received - ' + htmlEscape_(firstName) + ' - Bill Layne Insurance</title>',
    '<!--[if mso]><noscript><xml><o:OfficeDocumentSettings><o:AllowPNG/><o:PixelsPerInch>96</o:PixelsPerInch></o:OfficeDocumentSettings></xml></noscript><![endif]-->',
    '<style>body,table,td,p,a{-webkit-text-size-adjust:100%;-ms-text-size-adjust:100%}table,td{mso-table-lspace:0pt;mso-table-rspace:0pt}img{-ms-interpolation-mode:bicubic;border:0;outline:none;text-decoration:none;display:block}body{margin:0!important;padding:0!important;background-color:#f1f5f9;width:100%!important}.card-pad{padding:28px 32px!important}.hero-pad{padding:36px 28px!important}@media only screen and (max-width:620px){.email-container{width:100%!important}.card-pad{padding:20px 16px!important}.hero-pad{padding:28px 16px!important}.cta-btn{width:100%!important}}</style>',
    '</head><body style="margin:0;padding:0;background-color:#f1f5f9;">',

    '<div style="display:none;white-space:nowrap;font:15px courier;color:#f1f5f9;line-height:0;width:600px!important;min-width:600px!important;max-width:600px!important;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</div>',

    '<div style="display:none;max-height:0;overflow:hidden;mso-hide:all;font-size:1px;color:#f1f5f9;line-height:1px;">We received your photos &#8212; confirmation ' + htmlEscape_(submission.confirmationNumber) + '. All saved securely.&nbsp;&#847;&nbsp;&#847;&nbsp;&#847;&nbsp;&#847;&nbsp;&#847;&nbsp;&#847;&nbsp;&#847;&nbsp;&#847;&nbsp;&#847;&nbsp;&#847;&nbsp;&#847;&nbsp;&#847;&nbsp;&#847;&nbsp;&#847;&nbsp;&#847;&nbsp;&#847;&nbsp;&#847;</div>',

    '<table cellpadding="0" cellspacing="0" border="0" width="100%" bgcolor="#f1f5f9" style="background-color:#f1f5f9;"><tr><td align="center" style="padding:20px 10px;">',
    '<!--[if mso]><table align="center" border="0" cellspacing="0" cellpadding="0" width="600"><tr><td width="600"><![endif]-->',
    '<table cellpadding="0" cellspacing="0" border="0" width="600" class="email-container" style="max-width:600px;">',

    // CARD 1: HEADER
    '<tr><td style="padding-bottom:4px;">',
    '<table cellpadding="0" cellspacing="0" border="0" width="100%" bgcolor="#fafafa" style="background-color:#fafafa;border-radius:16px 16px 0 0;border:1px solid #e2e8f0;">',
    '<tr><td style="padding:28px 24px;text-align:center;" class="card-pad">',
    '<table cellpadding="0" cellspacing="0" border="0" style="margin:0 auto 16px auto;"><tr><td bgcolor="#ffffff" style="background-color:#ffffff;border-radius:8px;padding:8px 14px;">',
    '<img src="https://i.imgur.com/lxu9nfT.png" width="180" alt="Bill Layne Insurance Agency" style="display:block;width:180px;height:auto;">',
    '</td></tr></table>',
    '<table cellpadding="0" cellspacing="0" border="0" style="margin:0 auto;"><tr><td style="background-color:#ecfdf5;border-radius:20px;padding:8px 18px;">',
    '<span style="font-size:14px;color:#059669;font-weight:600;' + ff + '">&#9989; Photos Received Successfully</span>',
    '</td></tr></table>',
    '</td></tr></table></td></tr>',

    // CARD 2: CONFIRMATION DETAILS
    '<tr><td style="padding-bottom:4px;">',
    '<table cellpadding="0" cellspacing="0" border="0" width="100%" bgcolor="#ffffff" style="background-color:#ffffff;border-radius:16px;border:1px solid #e2e8f0;">',
    '<tr><td style="padding:28px 24px;" class="card-pad">',

    '<p style="margin:0 0 6px 0;font-size:12px;color:#64748b;' + ff + 'letter-spacing:1.5px;text-transform:uppercase;">CONFIRMATION</p>',
    '<p style="margin:0 0 12px 0;font-size:24px;font-weight:800;color:#0f172a;' + ff + 'line-height:1.3;">Hi ' + htmlEscape_(firstName) + ', we got your photos!</p>',
    '<p style="margin:0 0 24px 0;font-size:15px;color:#334155;' + ff + 'line-height:1.6;">We have your photos and they\'re saved securely. Our team will review them and reach out if we need anything else.</p>',

    // Details Card
    '<table cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color:#f8fafc;border-radius:12px;border:1px solid #e2e8f0;margin-bottom:20px;">',
    '<tr><td style="padding:20px;">',
    '<p style="margin:0 0 16px 0;font-size:12px;color:#64748b;' + ff + 'letter-spacing:1.5px;text-transform:uppercase;">DETAILS</p>',
    '<table cellpadding="0" cellspacing="0" border="0" width="100%">',
    '<tr><td style="padding:0 0 12px 0;"><table cellpadding="0" cellspacing="0" border="0" width="100%"><tr><td style="font-size:14px;color:#64748b;' + ff + '">Confirmation&nbsp;#</td><td align="right" style="font-size:14px;font-weight:700;color:#0f172a;' + ff + '">' + htmlEscape_(submission.confirmationNumber) + '</td></tr></table></td></tr>',
    '<tr><td style="padding:0 0 12px 0;border-top:1px solid #e2e8f0;padding-top:12px;"><table cellpadding="0" cellspacing="0" border="0" width="100%"><tr><td style="font-size:14px;color:#64748b;' + ff + '">Document Type</td><td align="right" style="font-size:14px;color:#0f172a;' + ff + '">' + htmlEscape_(submission.docType) + '</td></tr></table></td></tr>',
    '<tr><td style="border-top:1px solid #e2e8f0;padding-top:12px;"><table cellpadding="0" cellspacing="0" border="0" width="100%"><tr><td style="font-size:14px;color:#64748b;' + ff + '">Received</td><td align="right" style="font-size:14px;color:#0f172a;' + ff + '">' + htmlEscape_(localTime) + '</td></tr></table></td></tr>',
    '</table></td></tr></table>',

    // Files
    '<p style="margin:0 0 12px 0;font-size:12px;color:#64748b;' + ff + 'letter-spacing:1.5px;text-transform:uppercase;">FILES RECEIVED (' + savedFiles.length + ')</p>',
    '<table cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color:#f8fafc;border-radius:12px;border:1px solid #e2e8f0;">',
    fileRows,
    '</table>',

    '</td></tr></table></td></tr>',

    // CARD 3: NEXT STEPS
    '<tr><td style="padding-bottom:4px;">',
    '<table cellpadding="0" cellspacing="0" border="0" width="100%" bgcolor="#ffffff" style="background-color:#ffffff;border-radius:16px;border:1px solid #e2e8f0;border-left:4px solid #10b981;">',
    '<tr><td style="padding:28px 24px;" class="card-pad">',
    '<p style="margin:0 0 6px 0;font-size:12px;color:#64748b;' + ff + 'letter-spacing:1.5px;text-transform:uppercase;">NEXT STEPS</p>',
    '<p style="margin:0 0 24px 0;font-size:22px;font-weight:800;color:#0f172a;' + ff + '">What Happens Next</p>',

    buildStep_(ff, '1', '&#128274; Securely stored', 'Your photos are saved in our system'),
    buildStep_(ff, '2', '&#128203; Team review', 'Our team will review your submission'),
    buildStepLast_(ff, '3', '&#128222; We\'ll be in touch', 'We\'ll contact you if anything else is needed'),

    '</td></tr></table></td></tr>',

    // CARD 4: CTA
    '<tr><td style="padding-bottom:4px;">',
    '<table cellpadding="0" cellspacing="0" border="0" width="100%" bgcolor="#003f87" style="background-color:#003f87;border-radius:16px;border:1px solid #e2e8f0;">',
    '<tr><td style="padding:32px 24px;text-align:center;" class="card-pad">',
    '<p style="margin:0 0 20px 0;font-size:20px;font-weight:700;color:#ffffff;' + ff + '">Need help? We\'re here for you.</p>',
    '<table cellpadding="0" cellspacing="0" border="0" style="margin:0 auto 16px auto;"><tr><td bgcolor="#10b981" style="background-color:#10b981;border-radius:8px;"><a href="tel:3368351993" style="display:inline-block;padding:16px 40px;color:#ffffff;text-decoration:none;font-weight:700;font-size:16px;' + ff + '" class="cta-btn">&#128222;&nbsp;&nbsp;Call (336) 835-1993</a></td></tr></table>',
    '<table cellpadding="0" cellspacing="0" border="0" style="margin:0 auto;"><tr><td bgcolor="#ffffff" style="background-color:#ffffff;border-radius:8px;"><a href="https://www.BillLayneInsurance.com?utm_source=email&amp;utm_medium=photo_confirm&amp;utm_content=website_btn" style="display:inline-block;padding:14px 32px;color:#003f87;text-decoration:none;font-weight:700;font-size:14px;' + ff + '">Visit Our Website</a></td></tr></table>',
    '</td></tr></table></td></tr>',

    // FOOTER
    '<tr><td>',
    '<table cellpadding="0" cellspacing="0" border="0" width="100%" bgcolor="#fafafa" style="background-color:#fafafa;border-radius:0 0 16px 16px;border:1px solid #e2e8f0;border-top:none;">',
    '<tr><td style="padding:28px 24px;text-align:center;" class="card-pad">',

    '<table cellpadding="0" cellspacing="0" border="0" width="60" style="margin:0 auto 20px auto;"><tr><td style="height:3px;background:linear-gradient(90deg,#003f87,#C8A84E);font-size:0;line-height:0;">&nbsp;</td></tr></table>',

    // Agent Signature Chip
    '<table cellpadding="0" cellspacing="0" border="0" style="margin:0 auto 20px auto;background-color:#f1f5f9;border-radius:12px;border:1px solid #e2e8f0;"><tr>',
    '<td style="padding:14px 0 14px 16px;vertical-align:middle;"><table cellpadding="0" cellspacing="0" border="0" width="64" height="64" style="width:64px;height:64px;"><tr><td width="64" height="64" bgcolor="#ffffff" style="background-color:#ffffff;border-radius:10px;width:64px;height:64px;padding:0;"><img src="https://i.imgur.com/XacnUW4.jpeg" width="64" height="64" alt="Bill Layne" style="display:block;width:64px;height:64px;border-radius:10px;"></td></tr></table></td>',
    '<td style="padding:14px 18px 14px 14px;vertical-align:middle;text-align:left;">',
    '<p style="margin:0 0 1px 0;font-size:15px;font-weight:700;color:#0f172a;' + ff + '">Bill Layne</p>',
    '<p style="margin:0 0 5px 0;font-size:12px;color:#64748b;' + ff + '">Licensed Insurance Agent &bull; Since 2005</p>',
    '<table cellpadding="0" cellspacing="0" border="0"><tr><td style="padding-right:8px;"><a href="tel:3368351993" style="font-size:12px;font-weight:600;color:#003f87;text-decoration:none;' + ff + '">&#128222; (336)&nbsp;835&#8209;1993</a></td><td style="color:#cbd5e1;font-size:12px;padding-right:8px;">|</td><td><a href="mailto:Save@BillLayneInsurance.com" style="font-size:12px;font-weight:600;color:#003f87;text-decoration:none;' + ff + '">Send Email</a></td></tr></table>',
    '</td></tr></table>',

    // Agency Logo
    '<table cellpadding="0" cellspacing="0" border="0" style="margin:0 auto 16px auto;"><tr><td bgcolor="#ffffff" style="background-color:#ffffff;border-radius:8px;padding:8px 14px;"><img src="https://i.imgur.com/lxu9nfT.png" width="140" alt="Bill Layne Insurance Agency" style="display:block;width:140px;height:auto;"></td></tr></table>',

    '<p style="margin:0 0 4px 0;font-size:14px;font-weight:700;color:#0f172a;' + ff + '">Bill Layne Insurance Agency</p>',
    '<p style="margin:0 0 4px 0;font-size:12px;color:#64748b;' + ff + '">1283 N Bridge St, Elkin, NC 28621</p>',
    '<p style="margin:0 0 4px 0;font-size:12px;color:#64748b;' + ff + '"><a href="tel:3368351993" style="color:#003f87;text-decoration:none;' + ff + '">(336)&nbsp;835&#8209;1993</a>&nbsp;&bull;&nbsp;<a href="mailto:Save@BillLayneInsurance.com" style="color:#003f87;text-decoration:none;' + ff + '">Save@BillLayneInsurance.com</a></p>',
    '<p style="margin:0 0 16px 0;font-size:12px;color:#64748b;' + ff + '"><a href="https://www.BillLayneInsurance.com" style="color:#003f87;text-decoration:none;' + ff + '">BillLayneInsurance.com</a>&nbsp;&bull;&nbsp;Est. 2005</p>',

    // Social Row
    '<table cellpadding="0" cellspacing="0" border="0" style="margin:0 auto 16px auto;"><tr>',
    '<td style="padding:0 6px;"><a href="https://facebook.com/dollarbillagency" style="font-size:11px;color:#64748b;text-decoration:none;' + ff + '">Facebook</a></td>',
    '<td style="color:#cbd5e1;font-size:11px;">|</td>',
    '<td style="padding:0 6px;"><a href="https://youtube.com/@ncautoandhome" style="font-size:11px;color:#64748b;text-decoration:none;' + ff + '">YouTube</a></td>',
    '<td style="color:#cbd5e1;font-size:11px;">|</td>',
    '<td style="padding:0 6px;"><a href="https://instagram.com/ncautoandhome" style="font-size:11px;color:#64748b;text-decoration:none;' + ff + '">Instagram</a></td>',
    '<td style="color:#cbd5e1;font-size:11px;">|</td>',
    '<td style="padding:0 6px;"><a href="https://x.com/shopsavecompare" style="font-size:11px;color:#64748b;text-decoration:none;' + ff + '">X</a></td>',
    '</tr></table>',

    // Google Review Badge
    '<table cellpadding="0" cellspacing="0" border="0" style="margin:0 auto 14px auto;"><tr><td style="background-color:#f8fafc;border-radius:8px;padding:8px 14px;border:1px solid #e2e8f0;"><table cellpadding="0" cellspacing="0" border="0"><tr><td valign="middle" style="padding-right:6px;"><img src="https://i.imgur.com/nDFmjxh.png" width="18" height="18" alt="Google" style="display:block;width:18px;height:18px;"></td><td valign="middle"><p style="margin:0;font-size:12px;font-weight:700;color:#0f172a;' + ff + '">4.9 &#11088;&#11088;&#11088;&#11088;&#11088; <span style="font-weight:400;color:#64748b;">100+ Google Reviews</span></p></td></tr></table></td></tr></table>',

    '<p style="margin:0 0 10px 0;font-size:11px;color:#64748b;' + ff + 'text-align:center;">Follow us on Facebook for tips, reminders &amp; updates &nbsp;&rarr;&nbsp;<a href="https://facebook.com/dollarbillagency" target="_blank" style="color:#003f87;font-weight:700;text-decoration:none;' + ff + '">facebook.com/dollarbillagency</a></p>',

    '<p style="margin:0 0 10px 0;font-size:11px;color:#64748b;' + ff + 'text-align:center;">&#128242; Want policy updates in Messenger? <a href="https://m.me/dollarbillagency?text=Yes%2C%20please%20send%20my%20policy%20updates%20via%20Messenger" target="_blank" style="color:#003f87;font-weight:700;text-decoration:none;' + ff + '">Tap here to connect &rarr;</a></p>',

    '<p style="margin:0;font-size:11px;color:#94a3b8;' + ff + '">To unsubscribe from agency communications, reply with UNSUBSCRIBE.</p>',

    '</td></tr></table></td></tr>',

    '</table>',
    '<!--[if mso]></td></tr></table><![endif]-->',
    '</td></tr></table>',
    '</body></html>'
  ].join('');
}

function buildStep_(ff, num, title, desc) {
  return '<table cellpadding="0" cellspacing="0" border="0" width="100%" style="margin-bottom:20px;"><tr><td width="44" valign="top" style="padding-right:16px;"><table cellpadding="0" cellspacing="0" border="0" width="36" height="36"><tr><td width="36" height="36" align="center" valign="middle" bgcolor="#10b981" style="background-color:#10b981;border-radius:8px;font-size:15px;font-weight:700;color:#ffffff;' + ff + 'line-height:36px;">' + num + '</td></tr></table></td><td valign="top"><p style="margin:0 0 4px 0;font-size:16px;font-weight:700;color:#0f172a;' + ff + '">' + title + '</p><p style="margin:0;font-size:14px;color:#64748b;' + ff + 'line-height:1.5;">' + desc + '</p></td></tr></table>';
}

function buildStepLast_(ff, num, title, desc) {
  return '<table cellpadding="0" cellspacing="0" border="0" width="100%"><tr><td width="44" valign="top" style="padding-right:16px;"><table cellpadding="0" cellspacing="0" border="0" width="36" height="36"><tr><td width="36" height="36" align="center" valign="middle" bgcolor="#10b981" style="background-color:#10b981;border-radius:8px;font-size:15px;font-weight:700;color:#ffffff;' + ff + 'line-height:36px;">' + num + '</td></tr></table></td><td valign="top"><p style="margin:0 0 4px 0;font-size:16px;font-weight:700;color:#0f172a;' + ff + '">' + title + '</p><p style="margin:0;font-size:14px;color:#64748b;' + ff + 'line-height:1.5;">' + desc + '</p></td></tr></table>';
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
    '  1. Your photos are securely stored',
    '  2. Our team will review your submission',
    '  3. We\'ll contact you if anything else is needed',
    '',
    'Questions? Call us at (336) 835-1993',
    'Or visit https://www.billlayneinsurance.com',
    '',
    'Bill Layne Insurance Agency',
    '1283 N Bridge St, Elkin, NC 28621',
    '(336) 835-1993',
    'Save@BillLayneInsurance.com'
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
