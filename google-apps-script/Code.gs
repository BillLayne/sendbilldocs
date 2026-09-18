const APP_CONFIG = Object.freeze({
  agencyName: 'Bill Layne Insurance Agency',
  officePhone: '(336) 835-1993',
  officeWebsite: 'https://www.billlayneinsurance.com',
  emailLogoUrl: 'https://www.sendbilldocs.com/assets/bill-layne-logo.png',
  emailHeroImageUrl: 'https://www.sendbilldocs.com/assets/email-hero-2026.jpg',
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

  const officeEmails = String(props.getProperty('OFFICE_EMAILS') || 'docs@billlayneinsurance.com')
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
  var subject = buildCustomerSubject_(submission);
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

// ============================================================================
// CUSTOMER CONFIRMATION EMAIL - GOLD ELITE v2 SHELL (2026-09-18)
// Same structure as every Bill Layne Insurance Gmail email (the v2 template Bill
// confirmed in real Gmail on a phone): spacer div first, 1px 600-wide image second,
// preheader, Outlook ghost, fluid container, seamed cards, black header with the
// 190px logo and a gold badge, black CTA band, the fixed black footer.
// Wording follows the document type: "tow bill" (the /towing/ page), "photos", or
// "documents". Non-ASCII stays entity-encoded.
// ============================================================================
var GE_FONT = "font-family:Arial,Helvetica,sans-serif;";
var GE_SERIF = "font-family:Georgia,'Times New Roman',serif;";
var GE_TSA = "-webkit-text-size-adjust:100%;-ms-text-size-adjust:100%;";
var GE_LOGO = 'https://img.billlayneinsurance.com/cdn-cgi/image/width=380,format=png/i/2026/08/bli-agency-logo-4jqj2j.png';

function submissionNoun_(submission) {
  var docType = String(submission.docType || '');
  if (/tow/i.test(docType)) return { noun: 'tow bill', plural: false, label: 'Tow bill' };
  if (/photo/i.test(docType)) return { noun: 'photos', plural: true, label: 'Photos' };
  return { noun: 'documents', plural: true, label: 'Documents' };
}

function goldEliteLink_(href, text, color) {
  return '<a href="' + href + '" style="color:' + (color || '#D4A843') + ';text-decoration:none;font-weight:bold;">' + text + '</a>';
}

function goldEliteHeader_(badgeText) {
  return '<!-- HEADER -->' +
    '<tr><td style="padding-bottom:4px;">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="#000000" style="background-color:#000000;border-radius:16px 16px 0 0;">' +
    '<tr><td style="background-color:#000000;padding:18px 24px;border-bottom:4px solid #D4A843;border-radius:16px 16px 0 0;">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"><tr>' +
    '<td align="left" valign="middle" style="padding:0;"><img src="' + GE_LOGO + '" alt="Bill Layne Insurance Agency" width="190" style="display:block;width:190px;max-width:190px;height:auto;border:0;"></td>' +
    '<td align="right" valign="middle" style="padding:0 0 0 12px;"><table role="presentation" cellpadding="0" cellspacing="0" border="0" align="right"><tr><td style="' + GE_TSA + 'border:1px solid #D4A843;padding:7px 13px;' + GE_FONT + 'font-size:11px;line-height:14px;font-weight:bold;letter-spacing:1.5px;text-transform:uppercase;color:#D4A843;">' + badgeText + '</td></tr></table></td>' +
    '</tr></table>' +
    '</td></tr></table>' +
    '</td></tr>';
}

function goldEliteFooter_(trailing) {
  return '<!-- FOOTER -->' +
    '<tr><td>' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="#000000" style="background-color:#000000;border-radius:0 0 16px 16px;">' +
    '<tr><td style="padding:26px 20px 30px 20px;text-align:center;border-radius:0 0 16px 16px;">' +
    '<table role="presentation" cellpadding="0" cellspacing="0" border="0" align="center" style="margin:0 auto 14px auto;"><tr><td bgcolor="#ffffff" style="background-color:#ffffff;padding:8px 14px;border-radius:6px;"><img src="' + GE_LOGO + '" alt="Bill Layne Insurance Agency" width="150" style="display:block;width:150px;max-width:150px;height:auto;border:0;"></td></tr></table>' +
    '<div style="' + GE_TSA + GE_SERIF + 'font-size:15px;line-height:20px;color:#ffffff;font-weight:bold;">Bill Layne Insurance Agency</div>' +
    '<div style="' + GE_TSA + GE_FONT + 'font-size:11px;line-height:16px;color:#D4A843;font-weight:bold;letter-spacing:.5px;padding-top:3px;">Your Neighbor. Your Agent.</div>' +
    '<div style="' + GE_TSA + GE_FONT + 'font-size:12.5px;line-height:20px;color:#c9c9c9;padding-top:10px;">' +
    '1283 N Bridge St, PO Box 827, Elkin, NC 28621<br>' +
    goldEliteLink_('tel:+13368351993', '(336) 835-1993') + ' &middot; ' +
    goldEliteLink_('mailto:Save&#64;BillLayneInsurance&#46;com', 'Save&#64;BillLayneInsurance&#46;com') + '<br>' +
    goldEliteLink_('https://www.BillLayneInsurance.com', 'www.BillLayneInsurance.com') +
    '</div>' +
    '<div style="' + GE_TSA + GE_FONT + 'font-size:12.5px;line-height:20px;color:#ffffff;padding-top:14px;">' +
    goldEliteLink_('https://www.billlayneinsurance.com/clients/', 'Client Hub', '#ffffff') + ' &middot; ' +
    goldEliteLink_('https://www.billlayneinsurance.com/service-center#payment-panel', 'Pay', '#ffffff') + ' &middot; ' +
    goldEliteLink_('https://www.billlayneinsurance.com/service-center#idcard-panel', 'ID Cards', '#ffffff') + ' &middot; ' +
    goldEliteLink_('https://www.billlayneinsurance.com/claims-center/', 'Claims', '#ffffff') +
    '</div>' +
    '<div style="' + GE_TSA + GE_FONT + 'font-size:12.5px;line-height:20px;padding-top:6px;">' + goldEliteLink_('https://g.page/r/CXGq9B7-jzu7EBM/review', '&#9733; Review us on Google') + '</div>' +
    '<div style="' + GE_TSA + GE_FONT + 'font-size:12.5px;line-height:20px;color:#D4A843;padding-top:9px;">' +
    goldEliteLink_('https://www.facebook.com/dollarbillagency', 'Facebook') + ' &middot; ' +
    goldEliteLink_('https://www.instagram.com/ncautoandhome', 'Instagram') + ' &middot; ' +
    goldEliteLink_('https://www.tiktok.com/@ncautoandhome', 'TikTok') + ' &middot; ' +
    goldEliteLink_('https://www.youtube.com/@ncautoandhome', 'YouTube') + ' &middot; ' +
    goldEliteLink_('https://x.com/shopsavecompare', 'X') +
    '</div>' +
    '<div style="' + GE_TSA + GE_FONT + 'font-size:10.5px;line-height:15px;color:#8a8a8a;padding-top:14px;">NC License #6571216 &middot; Serving North Carolina since 2005</div>' +
    (trailing || '') +
    '</td></tr></table>' +
    '</td></tr>';
}

function goldEliteDetailRow_(label, value, last) {
  return '<tr><td valign="top" style="padding:0 0 14px 0;">' +
    '<div style="' + GE_TSA + GE_FONT + 'font-size:15px;line-height:21px;color:#000000;font-weight:bold;">' + label + '</div>' +
    '<div style="' + GE_TSA + GE_FONT + 'font-size:14px;line-height:21px;color:#555555;padding-top:4px;">' + value + '</div>' +
    '</td></tr>' + (last ? '' : '<tr><td style="padding:0 0 14px 0;border-top:1px solid #ece3d2;"></td></tr>');
}

function goldEliteStep_(num, title, desc, last) {
  return '<tr><td valign="top" width="34" style="width:34px;padding:0 12px ' + (last ? '0' : '14px') + ' 0;">' +
    '<table role="presentation" cellpadding="0" cellspacing="0" border="0"><tr><td width="28" height="28" align="center" valign="middle" bgcolor="#000000" style="' + GE_TSA + 'width:28px;height:28px;background-color:#000000;border-radius:6px;' + GE_FONT + 'font-size:13px;font-weight:bold;color:#D4A843;">' + num + '</td></tr></table></td>' +
    '<td valign="top" style="padding:0 0 ' + (last ? '0' : '14px') + ' 0;">' +
    '<div style="' + GE_TSA + GE_FONT + 'font-size:15px;line-height:21px;color:#000000;font-weight:bold;">' + title + '</div>' +
    '<div style="' + GE_TSA + GE_FONT + 'font-size:14px;line-height:21px;color:#555555;padding-top:2px;">' + desc + '</div>' +
    '</td></tr>';
}

function buildCustomerHtmlBody_(runtime, submission, savedFiles) {
  var kind = submissionNoun_(submission);
  var firstName = htmlEscape_((submission.name || 'there').split(' ')[0]);
  var ref = htmlEscape_(submission.confirmationNumber);
  var localTime = '';
  try {
    var d = new Date(submission.timestamp);
    localTime = Utilities.formatDate(d, runtime.timeZone || APP_CONFIG.defaultTimeZone, "MMMM d, yyyy 'at' h:mm a");
  } catch (e) {
    localTime = submission.timestamp;
  }

  var fileRows = savedFiles.map(function(file, index) {
    var lower = String(file.name || '').toLowerCase();
    var icon = lower.match(/\.(jpg|jpeg|png|gif|webp|heic|heif|bmp|tiff)$/) ? '&#128248;' : lower.match(/\.pdf$/) ? '&#128203;' : '&#128196;';
    return '<tr><td style="padding:9px 0;' + (index ? 'border-top:1px solid #ece3d2;' : '') + '">' +
      '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"><tr>' +
      '<td valign="top" style="' + GE_TSA + GE_FONT + 'font-size:14px;line-height:20px;color:#000000;padding-right:12px;">' + icon + ' ' + htmlEscape_(file.name) + '</td>' +
      '<td valign="top" align="right" style="' + GE_TSA + GE_FONT + 'font-size:12.5px;line-height:20px;color:#8a8a8a;">' + htmlEscape_(formatFileSize_(file.sizeBytes)) + '</td>' +
      '</tr></table></td></tr>';
  }).join('');

  var got = kind.plural ? 'we got your ' + kind.noun : 'we got your ' + kind.noun;
  var preheader = 'We received your ' + kind.noun + ' &mdash; confirmation ' + ref + '. Saved securely; we follow up if anything else is needed.';
  var heroBody = kind.noun === 'tow bill'
    ? 'Your tow bill is saved securely with your file. We will review it and follow up, usually within 1 business day, if anything else is needed.'
    : 'Your ' + kind.noun + ' ' + (kind.plural ? 'are' : 'is') + ' saved securely with your file. Our team will review ' + (kind.plural ? 'them' : 'it') + ' and follow up, usually within 1 business day, if anything else is needed.';

  var rows = [
    goldEliteHeader_(kind.label + ' received'),

    // HERO
    '<tr><td style="padding-bottom:4px;">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="#ffffff" style="background-color:#ffffff;border:1px solid #e3dccd;">' +
    '<tr><td class="hero-pad" style="padding:30px 20px 26px 20px;text-align:center;">' +
    '<div style="' + GE_TSA + GE_FONT + 'font-size:12px;line-height:16px;font-weight:bold;letter-spacing:1.6px;text-transform:uppercase;color:#8a6d2f;">&#9989; Received</div>' +
    '<div style="' + GE_TSA + GE_SERIF + 'font-size:27px;line-height:33px;color:#000000;padding-top:12px;">Thank you, ' + firstName + ' &mdash; ' + got + '.</div>' +
    '<div style="' + GE_TSA + GE_FONT + 'font-size:15px;line-height:23px;color:#374151;padding-top:14px;">' + heroBody + '</div>' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"><tr><td align="center" style="padding-top:16px;">' +
    '<table role="presentation" cellpadding="0" cellspacing="0" border="0"><tr><td style="background-color:#fdfbf6;border:1px solid #e3dccd;border-radius:4px;padding:9px 16px;text-align:center;">' +
    '<div style="' + GE_TSA + GE_FONT + 'font-size:10px;line-height:14px;font-weight:bold;letter-spacing:1px;text-transform:uppercase;color:#8a6d2f;">Confirmation</div>' +
    '<div style="' + GE_TSA + GE_FONT + 'font-size:15px;line-height:20px;font-weight:bold;color:#000000;padding-top:2px;">' + ref + '</div>' +
    '</td></tr></table>' +
    '</td></tr></table>' +
    '</td></tr></table>' +
    '</td></tr>',

    // DETAILS + FILES
    '<tr><td style="padding-bottom:4px;">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="#fdfbf6" style="background-color:#fdfbf6;border:1px solid #e3dccd;">' +
    '<tr><td class="card-pad" style="padding:24px 20px 10px 20px;">' +
    '<div style="' + GE_TSA + GE_FONT + 'font-size:11px;line-height:15px;font-weight:bold;letter-spacing:1.6px;text-transform:uppercase;color:#8a6d2f;padding-bottom:14px;">Submission details</div>' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">' +
    goldEliteDetailRow_('Confirmation #', ref, false) +
    goldEliteDetailRow_('What you sent', htmlEscape_(submission.docType || kind.label), false) +
    goldEliteDetailRow_('Received', htmlEscape_(localTime), false) +
    goldEliteDetailRow_('Name on file', htmlEscape_(submission.name || ''), true) +
    '</table>' +
    '<div style="' + GE_TSA + GE_FONT + 'font-size:11px;line-height:15px;font-weight:bold;letter-spacing:1.6px;text-transform:uppercase;color:#8a6d2f;padding:8px 0 6px 0;">Files received (' + savedFiles.length + ')</div>' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="#ffffff" style="background-color:#ffffff;border:1px solid #e3dccd;"><tr><td style="padding:4px 12px;"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">' + fileRows + '</table></td></tr></table>' +
    '<div style="height:14px;line-height:14px;font-size:0;">&nbsp;</div>' +
    '</td></tr></table>' +
    '</td></tr>',

    // WHAT HAPPENS NEXT
    '<tr><td style="padding-bottom:4px;">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="#ffffff" style="background-color:#ffffff;border:1px solid #e3dccd;">' +
    '<tr><td class="card-pad" style="padding:24px 20px;">' +
    '<div style="' + GE_TSA + GE_FONT + 'font-size:11px;line-height:15px;font-weight:bold;letter-spacing:1.6px;text-transform:uppercase;color:#8a6d2f;padding-bottom:14px;">What happens next</div>' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">' +
    goldEliteStep_('1', 'Saved securely', 'Your ' + kind.noun + ' ' + (kind.plural ? 'are' : 'is') + ' filed with your account.', false) +
    goldEliteStep_('2', 'We take it from here', 'Our team reviews what you sent' + (kind.noun === 'tow bill' ? ' and handles the tow reimbursement with your carrier' : '') + '.', false) +
    goldEliteStep_('3', 'We follow up', 'Usually within 1 business day, and only if anything else is needed.', true) +
    '</table>' +
    '</td></tr></table>' +
    '</td></tr>',

    // CTA
    '<tr><td style="padding-bottom:4px;">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="#000000" style="background-color:#000000;border-left:5px solid #D4A843;">' +
    '<tr><td class="card-pad" style="padding:22px 20px;">' +
    '<div style="' + GE_TSA + GE_SERIF + 'font-size:19px;line-height:25px;color:#ffffff;">Questions? We&rsquo;re here.</div>' +
    '<div style="' + GE_TSA + GE_FONT + 'font-size:14px;line-height:22px;color:#e5e7eb;padding-top:7px;">No action is needed. If you have something to add, reply to this email or reach us below.</div>' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top:16px;">' +
    '<tr><td class="full-btn" align="center" style="background-color:#D4A843;border-radius:6px;"><a href="tel:+13368351993" style="' + GE_TSA + 'display:block;padding:14px 18px;' + GE_FONT + 'font-size:15px;font-weight:bold;color:#000000;text-decoration:none;">Call (336) 835-1993</a></td></tr>' +
    '<tr><td height="8" style="font-size:0;line-height:8px;">&nbsp;</td></tr>' +
    '<tr><td class="full-btn" align="center" style="border:1px solid #3a3a3a;border-radius:6px;"><a href="mailto:Save&#64;BillLayneInsurance&#46;com?subject=' + encodeURIComponent('Question about ' + submission.confirmationNumber) + '" style="' + GE_TSA + 'display:block;padding:14px 18px;' + GE_FONT + 'font-size:15px;font-weight:bold;color:#ffffff;text-decoration:none;">Email Us</a></td></tr>' +
    '<tr><td align="center" style="' + GE_TSA + 'padding-top:10px;' + GE_FONT + 'font-size:12.5px;line-height:18px;color:#c9c9c9;">Prefer to text? (336) 835-1993</td></tr>' +
    '</table>' +
    '</td></tr></table>' +
    '</td></tr>',

    goldEliteFooter_('<div style="' + GE_TSA + GE_FONT + 'font-size:10.5px;line-height:16px;color:#8a8a8a;padding-top:12px;">This confirms we received your upload; it does not change your policy or confirm coverage. To unsubscribe from agency communications, reply with UNSUBSCRIBE.</div>')
  ].join('');

  return '<!DOCTYPE html>' +
    '<html lang="en" xmlns="http://www.w3.org/1999/xhtml" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">' +
    '<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><meta name="x-apple-disable-message-reformatting"><meta name="color-scheme" content="light only"><meta name="supported-color-schemes" content="light only">' +
    '<title>' + htmlEscape_(kind.label) + ' received &mdash; ' + ref + ' | Bill Layne Insurance</title>' +
    '<!--[if mso]><xml><o:OfficeDocumentSettings><o:PixelsPerInch>96</o:PixelsPerInch></o:OfficeDocumentSettings></xml><style>table{border-collapse:collapse}</style><![endif]-->' +
    '<style>body,table,td,a{-webkit-text-size-adjust:100%;-ms-text-size-adjust:100%}table,td{mso-table-lspace:0;mso-table-rspace:0}img{-ms-interpolation-mode:bicubic;border:0;height:auto;line-height:100%;outline:none;text-decoration:none}body{margin:0;padding:0;width:100%!important;background-color:#f1efe9}@media only screen and (max-width:620px){.email-container{width:100%!important;padding:0 8px!important}.card-pad{padding:22px 18px!important}.hero-pad{padding:26px 18px!important}.full-btn{width:100%!important;text-align:center!important}}</style>' +
    '<script type="application/ld+json">{"@context":"http://schema.org","@type":"EmailMessage","description":"' + htmlEscape_(kind.label) + ' received confirmation ' + ref + ' for ' + firstName + ' - Bill Layne Insurance"}</script>' +
    '</head>' +
    '<body style="margin:0;padding:0;width:100%!important;background-color:#f1efe9;">' +
    '<div style="display:none;white-space:nowrap;font:15px courier;color:#f1efe9;line-height:0;width:600px!important;min-width:600px!important;max-width:600px!important;">' + new Array(31).join('&nbsp;') + '</div>' +
    '<img src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=" width="600" height="1" alt="" style="display:block;width:600px!important;min-width:600px!important;max-width:600px!important;height:1px!important;line-height:1px;font-size:0;border:0;">' +
    '<div style="display:none;font-size:1px;line-height:1px;max-height:0;max-width:0;opacity:0;overflow:hidden;mso-hide:all;">' + preheader + new Array(41).join('&zwnj;&nbsp;') + '</div>' +
    '<!--[if mso]><table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr><td align="center"><table role="presentation" width="600" cellpadding="0" cellspacing="0"><tr><td><![endif]-->' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color:#f1efe9;"><tr><td align="center" style="padding:20px 0;">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" class="email-container" style="max-width:600px;margin:0 auto;">' +
    rows +
    '</table>' +
    '</td></tr></table>' +
    '<!--[if mso]></td></tr></table></td></tr></table><![endif]-->' +
    '</body></html>';
}

function buildCustomerSubject_(submission) {
  var kind = submissionNoun_(submission);
  var label = kind.label.split(' ').map(function(w) { return w.charAt(0).toUpperCase() + w.slice(1); }).join(' ');
  return label + ' Received - ' + submission.confirmationNumber + ' | Bill Layne Insurance';
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
    'We received your ' + submissionNoun_(submission).noun + '. Our team will review and follow up, usually within 1 business day, if anything else is needed.',
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
    '  1. Your ' + submissionNoun_(submission).noun + ' ' + (submissionNoun_(submission).plural ? 'are' : 'is') + ' securely stored',
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
