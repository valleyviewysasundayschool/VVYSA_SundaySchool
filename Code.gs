/**
 * Valley View Sunday School — Schedule Automation & Mobile Teacher Web App
 * Combines Google Calendar Sync, RSVP Monitoring, App-Styled HTML Reminders,
 * and the Mobile Web App (Home, Profile, Calendar, Attendance QR, & Sub Requests).
 */

const CONFIG = {
  CALENDAR_ID: '62eb2017c3517a34a455ffe06eba47705d2fbcb7526ea3bf2742efe018abcab2@group.calendar.google.com',
  SHEET_NAME: 'Schedule',
  CONTACT_SHEET_NAME: 'Contact',
  MONTHLY_TRAINING_SHEET_NAME: 'Monthly Training',
  TEACHING_AIDS_SHEET_NAME: 'Teaching Aids',
  TIME_ZONE: 'America/Denver', // Utah time, including daylight-saving changes
  DEFAULT_START_HOUR: 11,
  DEFAULT_START_MINUTE: 30,
  DURATION_MINUTES: 30,
  DAYS_AHEAD: 15,
  REMINDER_WINDOW_MINUTES: 60,
  EVENT_MARKER_PREFIX: '[SUNDAY_SCHOOL_AUTOMATION_V1]',
  REMINDER_PROPERTY_PREFIX: 'SUNCMF_REMINDER_SENT_V1:',
  SUBSTITUTE_ALERT_PROPERTY_PREFIX: 'SUNCMF_SUB_ALERT_SENT_V1:',
  INVITE_RETRY_PROPERTY_PREFIX: 'SUNCMF_INVITE_RETRY_V1:',
  MAIL_SENDER_NAME: 'Valley View Sunday School',
  // This Apps Script deployment (/exec). It powers the signed RSVP links in emails and the app's sign-in + data API.
  SUNDAY_SCHOOL_WEB_APP_URL: 'https://script.google.com/macros/s/AKfycbzBaM05VE2DM7TF2mZVSMM29DfTGQJkWGK84z_baeGSUisNamgZP2E2SZGKVoo3t_XS0w/exec',
  // After you publish the app on GitHub Pages (e.g. 'https://yourname.github.io/sunday-school/'), paste that address here
  // so the "Open App" buttons in emails open the installable app. Until then they point at a small landing page.
  GITHUB_PAGES_APP_URL: '',
  SESSION_DAYS: 90,          // how long a teacher stays signed in on their phone
  LOGIN_CODE_TTL_SEC: 600,   // emailed sign-in codes expire after 10 minutes
  BOOT_CACHE_KEY: 'SUNCMF_BOOT_DATA_V4',
  BOOT_CACHE_TTL_SECONDS: 90,
  // PUBLIC check-in address only (no admin key). The app adds ?class=Sunday%20School / Gospel%20Foundation / Temple%20Prep.
  // Never put the admin address (the one with key=...) here or anywhere in this project.
  ATTENDANCE_WEB_APP_URL: 'https://script.google.com/macros/s/AKfycbzqklYpvUI9-MxgHskE-X3zvIctPcxUtYHlkQoWT7jLA1XeYV1_aCQ-HCWOx4SJro-ahA/exec',
  APP_ICON_DRIVE_ID: '115TNM-K951a6qK5K5yPPs5ESKk-Qe4WM',
  FALLBACK_IMAGE_DRIVE_ID: '1CjvczesK34E4Jv4N0RX11GALvQ1Gj5_N',
  TEMPLE_ICON_DARK_DRIVE_ID: '13VY5W5lMc8pKWCv-dsudf_FT3tAV2tSU',
  TEMPLE_ICON_WHITE_DRIVE_ID: '1Il4a6rnbJI932z7VmTEEpa74muj1tngp',
  TEMPLE_ICON_DARK_URL: 'https://drive.google.com/thumbnail?id=13VY5W5lMc8pKWCv-dsudf_FT3tAV2tSU&sz=w256',
  TEMPLE_ICON_WHITE_URL: 'https://drive.google.com/thumbnail?id=1Il4a6rnbJI932z7VmTEEpa74muj1tngp&sz=w256',
  EMAIL_THEME: 'light' // 'light' or 'dark'
};

const MANAGED_FUNCTIONS = [
  'syncTeachingEventsToCalendar',
  'syncTeacherRSVPStatus',
  'sendScheduledReminders'
];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Sunday School Tools')
    .addItem('🚀 Enable Background Automation (Run Once)', 'installTriggers')
    .addSeparator()
    .addItem('📅 Sync Schedule to Google Calendar', 'syncTeachingEventsToCalendar')
    .addItem('✅ Update Acceptance Status from Calendar', 'syncTeacherRSVPStatus')
    .addItem('⏰ Check & Send Upcoming Reminders', 'sendScheduledReminders')
    .addSeparator()
    .addItem('🔗 Check Web App Setup', 'checkWebAppSetup')
    .addItem('🔒 Sign Out All Devices', 'signOutAllDevices')
    .addItem('🖼️ Upload PNG/SVG Icon to Selected Cell', 'openIconUploaderDialog')
    .addToUi();
}

/**
 * Opens a modal dialog inside Google Sheets allowing the user to upload any .png or .svg file
 * from their computer directly into the currently selected cell (e.g., Teaching Aids -> Logo / Icon URL).
 */
function openIconUploaderDialog() {
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body { font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif; background: #232323; color: #fff; margin: 0; padding: 20px; }
        .box { border: 2px dashed #5c8d97; border-radius: 12px; padding: 22px; text-align: center; background: #2e2e2e; cursor: pointer; }
        .box:hover { background: #363636; }
        .preview { width: 56px; height: 56px; border-radius: 10px; background: #d9d9d9; margin: 14px auto 6px; display: none; align-items: center; justify-content: center; padding: 4px; overflow: hidden; }
        .preview img { max-width: 100%; max-height: 100%; object-fit: contain; }
        .btn { background: #5c8d97; color: #fff; border: none; border-radius: 999px; padding: 10px 22px; font-weight: 700; font-size: 13px; cursor: pointer; margin-top: 14px; width: 100%; }
        .btn:disabled { opacity: 0.5; cursor: not-allowed; }
        .hint { font-size: 11.5px; color: #b0b0b0; margin-top: 10px; line-height: 1.4; }
        #status { font-size: 12px; margin-top: 10px; font-weight: 600; color: #8fd19e; text-align: center; }
      </style>
    </head>
    <body>
      <div class="box" onclick="document.getElementById('fileInput').click()">
        <div style="font-weight:700;font-size:14px;">Click to Choose a PNG or SVG File</div>
        <div style="font-size:11.5px;color:#b8b8b8;margin-top:4px;">Supports .png and .svg icons for Teaching Aids</div>
        <input type="file" id="fileInput" accept=".png,.svg,image/png,image/svg+xml" style="display:none" onchange="onFilePicked(this)">
        <div id="previewBox" class="preview"><img id="previewImg" alt="Preview"></div>
        <div id="fileName" style="font-size:11px;color:#d0d0d0;margin-top:4px;"></div>
      </div>
      <button id="saveBtn" class="btn" disabled onclick="uploadToCell()">Insert Icon into Selected Cell</button>
      <div id="status"></div>
      <div class="hint">
        Tip: You can also paste a Google Drive link, a direct .png/.svg URL, or raw <code>&lt;svg&gt;...&lt;/svg&gt;</code> code from Figma directly into the <b>Logo / Icon URL</b> column.
      </div>
      <script>
        let currentDataUri = '';
        let currentName = '';
        function onFilePicked(input) {
          const file = input.files && input.files[0];
          if (!file) return;
          currentName = file.name;
          document.getElementById('fileName').textContent = file.name;
          const reader = new FileReader();
          reader.onload = function(e) {
            const rawDataUri = e.target.result;
            if (file.type === 'image/svg+xml' || file.name.toLowerCase().endsWith('.svg') || rawDataUri.length <= 42000) {
              currentDataUri = rawDataUri;
              showPreview(currentDataUri);
            } else {
              // Resize large PNG to crisp 128x128 icon so it fits cleanly in a Sheet cell
              const img = new Image();
              img.onload = function() {
                const canvas = document.createElement('canvas');
                const maxDim = 128;
                let w = img.width || 128;
                let h = img.height || 128;
                if (w > maxDim || h > maxDim) {
                  if (w > h) { h = Math.round(h * maxDim / w); w = maxDim; }
                  else { w = Math.round(w * maxDim / h); h = maxDim; }
                }
                canvas.width = w;
                canvas.height = h;
                const ctx = canvas.getContext('2d');
                ctx.drawImage(img, 0, 0, w, h);
                currentDataUri = canvas.toDataURL('image/png');
                showPreview(currentDataUri);
              };
              img.src = rawDataUri;
            }
          };
          reader.readAsDataURL(file);
        }
        function showPreview(uri) {
          const box = document.getElementById('previewBox');
          const img = document.getElementById('previewImg');
          img.src = uri;
          box.style.display = 'flex';
          document.getElementById('saveBtn').disabled = false;
        }
        function uploadToCell() {
          if (!currentDataUri) return;
          const btn = document.getElementById('saveBtn');
          btn.disabled = true;
          btn.textContent = 'Saving to cell...';
          google.script.run
            .withSuccessHandler(function(msg) {
              document.getElementById('status').textContent = msg || 'Saved!';
              setTimeout(function() { google.script.host.close(); }, 900);
            })
            .withFailureHandler(function(err) {
              btn.disabled = false;
              btn.textContent = 'Insert Icon into Selected Cell';
              document.getElementById('status').style.color = '#ff8a80';
              document.getElementById('status').textContent = err.message || 'Upload failed';
            })
            .saveUploadedIconToActiveCell(currentDataUri, currentName);
        }
      </script>
    </body>
    </html>
  `).setWidth(380).setHeight(330);
  SpreadsheetApp.getUi().showModalDialog(html, 'Upload PNG or SVG Teaching Aid Icon');
}

/**
 * Writes an uploaded PNG/SVG data URI (or saves to Drive if > 48,000 chars) into the active cell.
 */
function saveUploadedIconToActiveCell(dataUri, fileName) {
  const cell = SpreadsheetApp.getActiveRange();
  if (!cell) throw new Error('Please select a cell in the sheet first.');
  const raw = String(dataUri || '').trim();
  if (!raw.startsWith('data:image/')) throw new Error('Invalid PNG or SVG file.');

  if (raw.length <= 48000) {
    cell.setValue(raw);
    invalidateBootCache_();
    return 'Icon saved to ' + cell.getA1Notation() + '!';
  }

  // If the SVG/PNG exceeds cell character limits, save to Google Drive and store the Drive URL
  const match = raw.match(/^data:(image\/[a-zA-Z0-9+.-]+);base64,(.+)$/);
  if (!match) throw new Error('Could not parse base64 image.');
  const bytes = Utilities.base64Decode(match[2]);
  const blob = Utilities.newBlob(bytes, match[1], fileName || 'teaching_aid_icon.png');
  const file = DriveApp.createFile(blob);
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {}
  cell.setValue(file.getUrl());
  invalidateBootCache_();
  return 'Saved to Drive & linked in ' + cell.getA1Notation() + '!';
}

/**
 * Automatically clears the 90-second boot cache whenever someone edits the spreadsheet
 * so new teachers, RSVP changes, or Teaching Aids appear immediately.
 */
function onEdit() {
  invalidateBootCache_();
}

function invalidateBootCache_() {
  try {
    CacheService.getScriptCache().remove(CONFIG.BOOT_CACHE_KEY);
    CacheService.getScriptCache().remove('allowed_users_v1');
  } catch (e) {}
}

/**
 * Serves the Valley View Sunday School Mobile Web App.
 */
function doGet(e) {
  // Signed 1-click RSVP links from calendar-invite emails.
  // (No data is served from GET requests anymore: the app talks to doPost after the teacher signs in.)
  if (e && e.parameter) {
    if (e.parameter.rsvp && e.parameter.a) {
      return handleSignedRsvp_(e.parameter);
    }
    if (e.parameter.rsvp && e.parameter.row) {
      // Legacy row-number links are no longer honored (they could change the wrong person's status).
      return renderRsvpMessagePage_('This link has expired',
        'Please open the Sunday School app to respond to your assignment, or use the calendar invitation.');
    }
  }
  return renderLandingPage_();
}


/** Replaces only this script's triggers; unrelated project triggers are preserved. */
function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (MANAGED_FUNCTIONS.includes(trigger.getHandlerFunction())) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('syncTeachingEventsToCalendar')
    .timeBased().everyDays(1).atHour(1).inTimezone(CONFIG.TIME_ZONE).create();

  // Check RSVP declines every 15 minutes so substitute alerts are timely.
  ScriptApp.newTrigger('syncTeacherRSVPStatus')
    .timeBased().everyMinutes(15).create();

  ScriptApp.newTrigger('sendScheduledReminders')
    .timeBased().everyMinutes(15).create();

  const message =
    'Automation enabled. Calendar sync runs daily; RSVP and reminder checks run every 15 minutes.\n\n' +
    'The first run may ask for Calendar, Sheets, and email permissions.' +
    (getSafeWebAppUrl_() ? '' : '\n\nNOTE: No Web App URL is set yet, so emails will not include RSVP/app buttons. See "Check Web App Setup".');
  try {
    SpreadsheetApp.getUi().alert(message);
  } catch (e) {
    console.log(message);
  }
}

function syncTeachingEventsToCalendar() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    console.log('Calendar sync skipped because another automation task is running.');
    return;
  }

  try {
    const sheet = getScheduleSheet_();
    const calendar = getCalendar_();
    const contacts = getContactsConfig_();
    if (!sheet || !calendar || !contacts || !checkCalendarTimeZone_(calendar)) return;

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    const values = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
    const richTexts = sheet.getRange(2, 6, lastRow - 1, 1).getRichTextValues();
    let createdCount = 0;
    let removedManagedCount = 0;
    let errorCount = 0;

    for (let i = 0; i < values.length; i++) {
      try {
        const row = values[i];
        const room = String(row[0] || '').trim();
        const teacher = String(row[2] || '').trim();
        const conductor = String(row[3] || '').trim();
        const rawDate = row[4];
        const lessonTitle = String(row[5] || '').trim();
        const teacherEmail = String(row[6] || '').trim().toLowerCase();
        if (!isValidAssignment_(teacher, teacherEmail, rawDate) || !room) continue;

        const dateKey = sheetDateKey_(rawDate, sheet);
        const daysUntilEvent = dayOrdinalFromKey_(dateKey) - dayOrdinal_(new Date(), CONFIG.TIME_ZONE);
        if (daysUntilEvent < 0 || daysUntilEvent > CONFIG.DAYS_AHEAD) continue;

        const assignmentKey = assignmentKey_(dateKey, room, teacherEmail);
        const marker = `${CONFIG.EVENT_MARKER_PREFIX}:${assignmentKey}`;
        const startTime = localDateTime_(dateKey, CONFIG.DEFAULT_START_HOUR, CONFIG.DEFAULT_START_MINUTE);
        const durationMs = (CONFIG.DURATION_MINUTES || 30) * 60 * 1000;
        const endTime = new Date(startTime.getTime() + durationMs);
        const eventTitle = `Sunday School Teaching: ${teacher} (${room})`;
        const dayStart = localDateTime_(dateKey, 0, 0);
        const dayEnd = localDateTime_(dateKey, 23, 59);
        const existingEvents = calendar.getEvents(dayStart, dayEnd);
        let matchingManagedEvent = false;

        for (const event of existingEvents) {
          const title = event.getTitle();
          const description = event.getDescription() || '';
          const eventIsManaged = description.includes(CONFIG.EVENT_MARKER_PREFIX + ':');
          const sameRoom = title.endsWith(`(${room})`) && title.startsWith('Sunday School Teaching: ');
          if (!sameRoom) continue;

          if (eventIsManaged && description.includes(marker)) {
            matchingManagedEvent = true;
            continue;
          }

          if (eventIsManaged && !description.includes(marker)) {
            event.deleteEvent();
            removedManagedCount++;
            sheet.getRange(i + 2, 2).clearContent();
          } else if (!eventIsManaged && title === eventTitle && guestIsInvited_(event, teacherEmail)) {
            event.setTime(startTime, endTime);
            event.setDescription(`${marker}\n\n${description}`);
            matchingManagedEvent = true;
          }
        }

        const lessonUrl = extractRichTextUrl_(richTexts[i][0]);
        const webAppUrl = getAppUrl_();
        const aidsRaw = String(row[8] || '').trim();
        const formattedDate = Utilities.formatDate(startTime, CONFIG.TIME_ZONE, 'EEEE, MMMM d, yyyy');
        const shortDate = Utilities.formatDate(startTime, CONFIG.TIME_ZONE, 'MMMM d');
        const inviteCtx = {
          rowNumber: i + 2, dateKey: dateKey, assignmentKey: assignmentKey,
          teacher: teacher, teacherEmail: teacherEmail, room: room, conductor: conductor,
          lessonTitle: lessonTitle, lessonUrl: lessonUrl, aidsRaw: aidsRaw,
          formattedDate: formattedDate, shortDate: shortDate, webAppUrl: webAppUrl, contacts: contacts
        };

        if (matchingManagedEvent) {
          retryPendingInvite_(inviteCtx); // only does anything if an earlier invite email failed
          continue;
        }

        const description =
          `${CONFIG.EVENT_MARKER_PREFIX}:${assignmentKey}\n\n` +
          'Valley View Sunday School Teaching Assignment\n\n' +
          `• Teacher: ${teacher}\n` +
          `• Conductor: ${conductor || 'N/A'}\n` +
          `• Room / Side: ${room}\n` +
          '• Time: 11:30 AM – 12:00 PM (Utah time)\n' +
          `• Date: ${formattedDate}\n\n` +
          `📖 Come, Follow Me Lesson:\n${lessonTitle}\n` +
          (lessonUrl ? `Lesson Link: ${lessonUrl}\n\n` : '\n') +
          (webAppUrl ? `📱 Sunday School App:\n${webAppUrl}\n\n` : '') +
          '----------------------------------------\n' +
          'RSVP INSTRUCTIONS:\n' +
          '• If you can teach, tap YES to accept this calendar invitation.\n' +
          '• If you cannot teach or need a substitute, tap NO and request a sub in the Sunday School App or contact:\n' +
          `  ${contacts.secName} (Sunday School Secretary)\n` +
          `  Phone: ${contacts.secPhone} | Email: ${contacts.secEmail}\n` +
          '----------------------------------------';

        calendar.createEvent(eventTitle, startTime, endTime, {
          description: description,
          guests: teacherEmail,
          sendInvites: false
        });
        createdCount++;

        sendInviteEmail_(inviteCtx);
      } catch (error) {
        errorCount++;
        console.error(`Calendar sync failed for sheet row ${i + 2}: ${error}`);
      }
    }

    cleanupOldProperties_();
    console.log(`Calendar sync complete. Created ${createdCount} event(s); removed ${removedManagedCount} outdated automation-managed event(s); ${errorCount} row error(s).`);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Checks RSVP status and alerts on every newly observed decline.
 * A manual "SUB" entry in column B always wins over the calendar status,
 * so it prevents reminders and triggers (once) the substitute alert.
 */
function syncTeacherRSVPStatus() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) {
    console.log('RSVP sync skipped because another automation task is running.');
    return;
  }

  try {
    const sheet = getScheduleSheet_();
    const calendar = getCalendar_();
    const contacts = getContactsConfig_();
    if (!sheet || !calendar || !contacts || !checkCalendarTimeZone_(calendar)) return;

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    const values = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
    const properties = PropertiesService.getScriptProperties();
    const todayOrdinal = dayOrdinal_(new Date(), CONFIG.TIME_ZONE);
    let updatedCount = 0;
    let alertCount = 0;

    for (let i = 0; i < values.length; i++) {
      const room = String(values[i][0] || '').trim();
      const currentSheetStatus = String(values[i][1] || '').trim().toUpperCase();
      const teacher = String(values[i][2] || '').trim();
      const rawDate = values[i][4];
      const lessonTitle = String(values[i][5] || '').trim();
      const teacherEmail = String(values[i][6] || '').trim().toLowerCase();

      if (!isValidAssignment_(teacher, teacherEmail, rawDate) || !room) continue;

      const dateKey = sheetDateKey_(rawDate, sheet);
      const daysUntil = dayOrdinalFromKey_(dateKey) - todayOrdinal;
      if (daysUntil < -1 || daysUntil > CONFIG.DAYS_AHEAD) continue;

      let statusText = null;

      if (currentSheetStatus === 'SUB') {
        statusText = 'SUB';
      } else {
        const dayStart = localDateTime_(dateKey, 0, 0);
        const dayEnd = localDateTime_(dateKey, 23, 59);
        const events = calendar.getEvents(dayStart, dayEnd);

        for (const event of events) {
          if (event.getTitle() !== `Sunday School Teaching: ${teacher} (${room})`) continue;
          const guest = event.getGuestList().find(g => g.getEmail().trim().toLowerCase() === teacherEmail);
          if (!guest) continue;
          const calStatus = guestStatusText_(guest.getGuestStatus());
          // If the teacher already responded via the styled email/app (ACCEPTED or TENTATIVE) and calendar is still PENDING, preserve sheet status
          if (calStatus === 'PENDING' && (currentSheetStatus === 'ACCEPTED' || currentSheetStatus === 'TENTATIVE')) {
            statusText = currentSheetStatus;
          } else {
            statusText = calStatus;
          }
          break;
        }
      }

      if (statusText === null) continue;

      if (currentSheetStatus !== statusText) {
        sheet.getRange(i + 2, 2).setValue(statusText);
        updatedCount++;
      }

      const assignmentKey = assignmentKey_(dateKey, room, teacherEmail);
      const alertKey = CONFIG.SUBSTITUTE_ALERT_PROPERTY_PREFIX + assignmentKey;

      if (statusText === 'SUB') {
        if (properties.getProperty(alertKey) !== 'SENT') {
          if (sendSubAlertEmail_(contacts, teacher, dateKey, room, lessonTitle)) {
            properties.setProperty(alertKey, 'SENT');
            alertCount++;
          }
        }
      } else {
        properties.deleteProperty(alertKey);
      }
    }

    console.log(`RSVP sync complete. Updated ${updatedCount} status(es); sent ${alertCount} substitute alert(s).`);
  } finally {
    lock.releaseLock();
  }
}

function sendScheduledReminders() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return;

  try {
    const sheet = getScheduleSheet_();
    const calendar = getCalendar_();
    const contacts = getContactsConfig_();
    if (!sheet || !calendar || !contacts || !checkCalendarTimeZone_(calendar)) return;

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    const values = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
    const richTexts = sheet.getRange(2, 6, lastRow - 1, 1).getRichTextValues();
    const now = new Date();
    const todayOrdinal = dayOrdinal_(now, CONFIG.TIME_ZONE);
    const properties = PropertiesService.getScriptProperties();

    for (let i = 0; i < values.length; i++) {
      const room = String(values[i][0] || '').trim();
      const currentSheetStatus = String(values[i][1] || '').trim().toUpperCase();
      const teacher = String(values[i][2] || '').trim();
      const conductor = String(values[i][3] || '').trim();
      const rawDate = values[i][4];
      const lessonTitle = String(values[i][5] || '').trim();
      const teacherEmail = String(values[i][6] || '').trim().toLowerCase();

      if (!isValidAssignment_(teacher, teacherEmail, rawDate) || !room) continue;

      const dateKey = sheetDateKey_(rawDate, sheet);
      const daysUntil = dayOrdinalFromKey_(dateKey) - todayOrdinal;
      if (daysUntil < -1 || daysUntil > 8) continue;

      const teachingDate = localDateTime_(dateKey, 12, 0);
      const assignmentKey = assignmentKey_(dateKey, room, teacherEmail);
      const reminders = [
        { type: 'ONE_WEEK', target: reminderTarget_(dateKey, -7, 16, 0), label: 'Upcoming Sunday Assignment' },
        { type: 'FINAL_DAY', target: reminderTarget_(dateKey, -1, 9, 0), label: 'Tomorrow’s Teaching Assignment' }
      ];

      for (const reminder of reminders) {
        const diffMinutes = (now.getTime() - reminder.target.getTime()) / 60000;
        if (diffMinutes < 0 || diffMinutes >= CONFIG.REMINDER_WINDOW_MINUTES) continue;

        const propertyKey = CONFIG.REMINDER_PROPERTY_PREFIX + assignmentKey + ':' + reminder.type;
        if (properties.getProperty(propertyKey)) continue;

        if (currentSheetStatus === 'SUB') continue;
        const calendarStatus = getMatchingEventStatus_(calendar, dateKey, teacher, room, teacherEmail);
        if (calendarStatus === 'SUB') continue;

        const lessonUrl = extractRichTextUrl_(richTexts[i][0]);
        const formattedDate = Utilities.formatDate(teachingDate, CONFIG.TIME_ZONE, 'EEEE, MMMM d, yyyy');
        const shortDate = Utilities.formatDate(teachingDate, CONFIG.TIME_ZONE, 'MMMM d');
        const subject = `Reminder: Sunday School Teaching Assignment on ${shortDate}`;
        const plainBody =
          `Hi ${teacher},\n\n` +
          'This is a friendly reminder that you are scheduled to teach Sunday School:\n\n' +
          `• Date: ${formattedDate}\n` +
          `• Room: ${room}\n` +
          `• Conductor: ${conductor || 'N/A'}\n` +
          `• Lesson: ${lessonTitle}\n` +
          (lessonUrl ? `• Lesson Link: ${lessonUrl}\n\n` : '\n') +
          `If you need to reschedule or request a sub, please use the Sunday School app or contact ${contacts.secName} (${contacts.secPhone}, ${contacts.secEmail}) as soon as possible.\n\n` +
          'Thanks for all you do, and feel free to reach out if you need help.\n' +
          '-The Sunday School Presidency';

        const aidsRaw = String(values[i][8] || '').trim();
        const htmlBody = buildStyledReminderEmailHtml_({
          dateKey: dateKey,
          teacher: teacher,
          badgeText: reminder.label,
          formattedDate: formattedDate,
          room: room,
          conductor: conductor || 'N/A',
          lessonRaw: lessonTitle,
          lessonUrl: lessonUrl,
          secName: contacts.secName,
          secPhone: contacts.secPhone,
          secEmail: contacts.secEmail,
          aidsRaw: aidsRaw
        });

        try {
          MailApp.sendEmail(teacherEmail, subject, plainBody, mailOptions_({ htmlBody: htmlBody }, contacts));
          properties.setProperty(propertyKey, new Date().toISOString());
        } catch (error) {
          console.error(`Reminder failed for ${teacher} (${teacherEmail}) on ${dateKey}. Error: ${error}`);
        }
      }
    }
  } finally {
    lock.releaseLock();
  }
}

/**
 * Builds an app-styled HTML reminder email matching the Figma dark UI theme.
 */
const STANDARD_TEACHING_AIDS = [
  {
    name: "Gospel Library — Come, Follow Me",
    description: "Official Sunday School manual and scriptures",
    url: "https://www.churchofjesuschrist.org/study/come-follow-me"
  },
  {
    name: "Follow Him Podcast",
    description: "Weekly podcast episodes and show notes with Hank Smith & John Bytheway",
    url: "https://followhim.co"
  },
  {
    name: "Teaching in the Savior’s Way",
    description: "Guiding principles for Christ-centered gospel teaching",
    url: "https://www.churchofjesuschrist.org/study/manual/teaching-in-the-saviors-way"
  },
  {
    name: "Scripture Central",
    description: "Contextual commentary, charts, historical insights, and videos",
    url: "https://scripturecentral.org"
  },
  {
    name: "Church Media Library",
    description: "Sacred artwork, historical images, and video clips for lessons",
    url: "https://www.churchofjesuschrist.org/media"
  }
];

/**
 * Builds a high-resolution (128px) favicon URL for any website link.
 * Uses fallback_opts=TYPE,SIZE (without URL) so if a domain has NO favicon,
 * Google returns HTTP 404 and the app automatically falls back to the Book SVG icon.
 */
function buildAutoFaviconUrl_(url) {
  const s = String(url || '').trim();
  if (!s || !/^https?:\/\//i.test(s)) return '';
  try {
    const m = s.match(/^(https?:\/\/[^\/?#]+)/i);
    const origin = m ? m[1] : s;
    return 'https://t2.gstatic.com/faviconV2?client=SOCIAL&type=FAVICON&fallback_opts=TYPE,SIZE&url=' +
      encodeURIComponent(origin) + '&size=128';
  } catch (e) {
    return '';
  }
}

/**
 * Infers a clean title from a URL if the Title cell was left blank.
 */
function inferAidTitleFromUrl_(url) {
  const s = String(url || '').trim();
  if (!s) return 'Lesson Study Resource';
  const m = s.match(/^https?:\/\/(?:www\.)?([^\/?#]+)/i);
  const host = m ? m[1].toLowerCase() : '';
  if (host.indexOf('youtube.com') !== -1 || host.indexOf('youtu.be') !== -1) return 'YouTube Video Resource';
  if (host.indexOf('fairlatterdaysaints.org') !== -1) return 'FAIR Study Resource';
  if (host.indexOf('followhim.co') !== -1) return 'Follow Him Podcast';
  if (host.indexOf('scripturecentral.org') !== -1) return 'Scripture Central Resource';
  if (host.indexOf('churchofjesuschrist.org') !== -1) return 'Church Study Resource';
  return host || 'Lesson Study Resource';
}

/**
 * Reads the unified "Teaching Aids" tab (Layout Option A):
 *   Column A: Lesson Date (or "Standard")
 *   Column B: Title / Source Name
 *   Column C: Description
 *   Column D: Link URL
 *   Column E: Custom Icon (Optional) — when left blank, the app shows the link's favicon (falling back to the book icon).
 * Also remains backward-compatible if the sheet still has the legacy column order.
 */
function readTeachingAidsSheet_(ss) {
  const teachingAids = [];
  const aidsSheet = (ss || SpreadsheetApp.getActiveSpreadsheet()).getSheetByName(CONFIG.TEACHING_AIDS_SHEET_NAME);
  if (!aidsSheet || aidsSheet.getLastRow() < 2) {
    return STANDARD_TEACHING_AIDS.map(function(a) {
      return {
        dateKey: '',
        name: a.name,
        description: a.description,
        url: a.url,
        icon: '',
        faviconUrl: buildAutoFaviconUrl_(a.url),
        isStandard: true
      };
    });
  }

  const numAidRows = aidsSheet.getLastRow() - 1;
  const headerRow = aidsSheet.getRange(1, 1, 1, 5).getValues()[0].map(function(h) { return String(h || '').trim().toLowerCase(); });
  // Detect Layout Option A (Column A header mentions "date", "week", or "lesson") vs legacy layout
  const isOptionALayout = headerRow[0].indexOf('date') !== -1 || headerRow[0].indexOf('week') !== -1 || headerRow[0].indexOf('standard') !== -1 || headerRow[0].indexOf('lesson') !== -1;

  const aVals = aidsSheet.getRange(2, 1, numAidRows, 5).getValues();
  const aRichColB = aidsSheet.getRange(2, 2, numAidRows, 1).getRichTextValues();
  const aRichColC = aidsSheet.getRange(2, 3, numAidRows, 1).getRichTextValues();
  const aRichColD = aidsSheet.getRange(2, 4, numAidRows, 1).getRichTextValues();
  const aRichColE = aidsSheet.getRange(2, 5, numAidRows, 1).getRichTextValues();
  const aFormulasColD = aidsSheet.getRange(2, 4, numAidRows, 1).getFormulas();
  const aFormulasColE = aidsSheet.getRange(2, 5, numAidRows, 1).getFormulas();

  // Map any floating OverGridImage anchored on the Teaching Aids sheet by row number
  const overGridByRow = {};
  try {
    const sheetImages = aidsSheet.getImages ? aidsSheet.getImages() : [];
    for (let k = 0; k < sheetImages.length; k++) {
      const imgObj = sheetImages[k];
      const anchor = imgObj.getAnchorCell ? imgObj.getAnchorCell() : null;
      if (anchor) {
        overGridByRow[anchor.getRow()] = imgObj;
      }
    }
  } catch (e) {}

  for (let i = 0; i < aVals.length; i++) {
    const rowNum = i + 2;
    let dateKey = '';
    let isStandard = true;
    let sName = '';
    let desc = '';
    let finalUrl = '';
    let iconResolved = '';

    if (isOptionALayout) {
      const rawDateOrStd = aVals[i][0];
      const rawStr = String(rawDateOrStd || '').trim();
      const upper = rawStr.toUpperCase();
      if (rawDateOrStd instanceof Date && !isNaN(rawDateOrStd.getTime())) {
        dateKey = sheetDateKey_(rawDateOrStd, aidsSheet);
        isStandard = false;
      } else if (/^\d{4}-\d{2}-\d{2}$/.test(rawStr)) {
        dateKey = rawStr;
        isStandard = false;
      } else if (/^\d{1,2}\/\d{1,2}\/\d{2,4}$/.test(rawStr)) {
        const parsedD = new Date(rawStr);
        if (!isNaN(parsedD.getTime())) {
          dateKey = sheetDateKey_(parsedD, aidsSheet);
          isStandard = false;
        }
      } else if (upper === 'NO' || upper === 'WEEKLY') {
        isStandard = false;
      } else {
        dateKey = '';
        isStandard = true;
      }

      sName = String(aVals[i][1] || '').trim();
      desc = String(aVals[i][2] || '').trim();
      const urlText = String(aVals[i][3] || '').trim();
      const richUrlD = extractRichTextUrl_(aRichColD[i] ? aRichColD[i][0] : null);
      const richUrlB = extractRichTextUrl_(aRichColB[i] ? aRichColB[i][0] : null);
      finalUrl = richUrlD || urlText || richUrlB;

      iconResolved = resolveTeachingAidIcon_(
        aVals[i][4],
        aRichColE[i] ? aRichColE[i][0] : null,
        aFormulasColE[i] ? aFormulasColE[i][0] : '',
        overGridByRow[rowNum] || null
      );
    } else {
      // Legacy 5-column layout fallback (Source Name | Description | URL | Logo / Icon URL | Standard Aid)
      sName = String(aVals[i][0] || '').trim();
      desc = String(aVals[i][1] || '').trim();
      const urlText = String(aVals[i][2] || '').trim();
      const richUrlC = extractRichTextUrl_(aRichColC[i] ? aRichColC[i][0] : null);
      finalUrl = richUrlC || urlText;
      iconResolved = resolveTeachingAidIcon_(
        aVals[i][3],
        aRichColD[i] ? aRichColD[i][0] : null,
        aFormulasColD[i] ? aFormulasColD[i][0] : '',
        overGridByRow[rowNum] || null
      );
      const std = String(aVals[i][4] || 'YES').trim().toUpperCase();
      if (/^\d{4}-\d{2}-\d{2}$/.test(std)) {
        dateKey = std;
        isStandard = false;
      } else {
        isStandard = std !== 'NO';
      }
    }

    if (!sName && !finalUrl) continue;
    if (!sName && finalUrl) {
      sName = inferAidTitleFromUrl_(finalUrl);
    }
    if (!desc) {
      desc = isStandard ? 'Standard study resource for teachers' : 'Assigned study resource for this Sunday’s lesson';
    }

    teachingAids.push({
      dateKey: dateKey,
      name: sName,
      description: desc,
      url: finalUrl || 'https://www.churchofjesuschrist.org/study/come-follow-me',
      icon: iconResolved || '', // Blank means: automatically display favicon (fallback to book icon)
      faviconUrl: buildAutoFaviconUrl_(finalUrl),
      isStandard: isStandard
    });
  }

  return teachingAids;
}

/**
 * Resolves the teaching aids to display for a specific Sunday dateKey (in emails and calendar invites).
 * Uses assigned rows for that Sunday from the "Teaching Aids" tab (plus any legacy Column I entries),
 * or falls back to the Standard rows on the "Teaching Aids" tab.
 */
function parseTeachingAidsList_(aidsRaw, dateKey) {
  let allSheetAids = [];
  try {
    allSheetAids = readTeachingAidsSheet_(SpreadsheetApp.getActiveSpreadsheet());
  } catch (e) {}

  const targetDate = String(dateKey || '').trim();
  const weeklyFromTab = targetDate
    ? allSheetAids.filter(function(a) { return a.dateKey === targetDate; })
    : [];

  const legacyWeekly = [];
  if (aidsRaw) {
    const parts = String(aidsRaw).split(/,|\n/).map(function(s) { return s.trim(); }).filter(Boolean);
    parts.forEach(function(entry) {
      if (entry.indexOf('|') !== -1) {
        const p = entry.split('|').map(function(s) { return s.trim(); });
        const u = p[2] || p[0];
        legacyWeekly.push({
          name: p[0] || inferAidTitleFromUrl_(u),
          description: p[1] || 'Assigned study resource',
          url: u,
          icon: p[3] || '',
          faviconUrl: buildAutoFaviconUrl_(u)
        });
      } else {
        const u = entry.indexOf('http') === 0 ? entry : 'https://www.churchofjesuschrist.org/study/come-follow-me';
        const alreadyInTab = weeklyFromTab.some(function(w) { return w.url === u; });
        if (!alreadyInTab) {
          legacyWeekly.push({
            name: entry.indexOf('http') === 0 ? inferAidTitleFromUrl_(u) : entry,
            description: 'Specially assigned lesson resource',
            url: u,
            icon: '',
            faviconUrl: buildAutoFaviconUrl_(u)
          });
        }
      }
    });
  }

  const combinedWeekly = weeklyFromTab.concat(legacyWeekly);
  if (combinedWeekly.length > 0) return combinedWeekly;

  const standardFromTab = allSheetAids.filter(function(a) { return a.isStandard !== false; });
  if (standardFromTab.length > 0) return standardFromTab;

  return STANDARD_TEACHING_AIDS.map(function(a) {
    return {
      name: a.name,
      description: a.description,
      url: a.url,
      icon: '',
      faviconUrl: buildAutoFaviconUrl_(a.url)
    };
  });
}

function buildGoogleCalendarAddUrl_(data) {
  if (!data.dateKey) return 'https://calendar.google.com';
  try {
    const start = localDateTime_(data.dateKey, CONFIG.DEFAULT_START_HOUR, CONFIG.DEFAULT_START_MINUTE);
    const end = new Date(start.getTime() + (CONFIG.DURATION_MINUTES || 30) * 60 * 1000);
    const fmt = function(d) { return Utilities.formatDate(d, 'UTC', "yyyyMMdd'T'HHmmss'Z'"); };
    const title = encodeURIComponent('Sunday School Teaching: ' + data.teacher + ' (' + data.room + ')');
    const details = encodeURIComponent(
      'Valley View Sunday School Teaching Assignment\n' +
      'Lesson: ' + (data.lessonRaw || '') + '\n' +
      (data.lessonUrl ? 'Lesson Link: ' + data.lessonUrl + '\n' : '') +
      'Room: ' + (data.room || '') + '\n' +
      'Conductor: ' + (data.conductor || '')
    );
    const location = encodeURIComponent(data.room || 'Valley View Ward');
    return 'https://calendar.google.com/calendar/render?action=TEMPLATE&text=' + title +
      '&dates=' + fmt(start) + '/' + fmt(end) +
      '&details=' + details +
      '&location=' + location;
  } catch (e) {
    return 'https://calendar.google.com';
  }
}

function syncSingleAssignmentToCalendarAndRsvp_(rowNumber, statusToSet) {
  let gcalTemplateUrl = 'https://calendar.google.com';
  let formattedDate = '';
  let room = '';
  let teacher = '';
  let lessonTitle = '';

  try {
    const sheet = getScheduleSheet_();
    if (!sheet) return { gcalTemplateUrl: gcalTemplateUrl };
    const r = Number(rowNumber);
    if (!r || r < 2 || r > sheet.getLastRow()) return { gcalTemplateUrl: gcalTemplateUrl };

    const rowVals = sheet.getRange(r, 1, 1, 7).getValues()[0];
    const richVal = sheet.getRange(r, 6, 1, 1).getRichTextValue();
    room = String(rowVals[0] || '').trim();
    teacher = String(rowVals[2] || '').trim();
    const conductor = String(rowVals[3] || '').trim();
    const rawDate = rowVals[4];
    lessonTitle = String(rowVals[5] || '').trim();
    const teacherEmail = String(rowVals[6] || '').trim().toLowerCase();
    const lessonUrl = extractRichTextUrl_(richVal);

    if (!(rawDate instanceof Date && !isNaN(rawDate.getTime()))) {
      return { gcalTemplateUrl: gcalTemplateUrl };
    }

    const dateKey = sheetDateKey_(rawDate, sheet);
    const startTime = localDateTime_(dateKey, CONFIG.DEFAULT_START_HOUR, CONFIG.DEFAULT_START_MINUTE);
    const durationMs = (CONFIG.DURATION_MINUTES || 30) * 60 * 1000;
    const endTime = new Date(startTime.getTime() + durationMs);
    formattedDate = Utilities.formatDate(startTime, CONFIG.TIME_ZONE, 'EEEE, MMMM d, yyyy');

    gcalTemplateUrl = buildGoogleCalendarAddUrl_({
      dateKey: dateKey,
      teacher: teacher,
      room: room,
      conductor: conductor,
      lessonRaw: lessonTitle,
      lessonUrl: lessonUrl
    });

    const calendar = getCalendar_();
    if (calendar && teacherEmail && room) {
      const dayStart = localDateTime_(dateKey, 0, 0);
      const dayEnd = localDateTime_(dateKey, 23, 59);
      const eventTitle = `Sunday School Teaching: ${teacher} (${room})`;
      const assignmentKey = assignmentKey_(dateKey, room, teacherEmail);
      const marker = `${CONFIG.EVENT_MARKER_PREFIX}:${assignmentKey}`;
      const existingEvents = calendar.getEvents(dayStart, dayEnd);
      let targetEvent = null;

      for (let i = 0; i < existingEvents.length; i++) {
        const ev = existingEvents[i];
        const desc = ev.getDescription() || '';
        if (ev.getTitle() === eventTitle || desc.indexOf(marker) !== -1) {
          targetEvent = ev;
          break;
        }
      }

      // If the event isn't on the ward Google Calendar yet, create it immediately and add the teacher as a guest
      if (!targetEvent && statusToSet !== 'SUB') {
        const contacts = getContactsConfig_() || { secName: '', secPhone: '', secEmail: '' };
        const webAppUrl = getAppUrl_();
        const description =
          `${marker}\n\n` +
          'Valley View Sunday School Teaching Assignment\n\n' +
          `• Teacher: ${teacher}\n` +
          `• Conductor: ${conductor || 'N/A'}\n` +
          `• Room / Side: ${room}\n` +
          '• Time: 11:30 AM – 12:00 PM (Utah time)\n' +
          `• Date: ${formattedDate}\n\n` +
          `📖 Come, Follow Me Lesson:\n${lessonTitle}\n` +
          (lessonUrl ? `Lesson Link: ${lessonUrl}\n\n` : '\n') +
          (webAppUrl ? `📱 Sunday School App:\n${webAppUrl}\n\n` : '') +
          '----------------------------------------\n' +
          'RSVP Status: ' + statusToSet + '\n' +
          '----------------------------------------';

        targetEvent = calendar.createEvent(eventTitle, startTime, endTime, {
          description: description,
          guests: teacherEmail,
          sendInvites: false
        });
      }

      if (targetEvent) {
        // Ensure teacher is on the guest list so the event appears on their Google Calendar
        if (!guestIsInvited_(targetEvent, teacherEmail) && statusToSet !== 'SUB') {
          try {
            targetEvent.addGuest(teacherEmail);
          } catch (e) {}
        }

        // Try to sync guest status via Advanced Calendar API v3 if enabled in the project
        try {
          if (typeof Calendar !== 'undefined' && Calendar.Events && Calendar.Events.get && Calendar.Events.patch) {
            const rawCalId = targetEvent.getId().replace(/@google\.com$/, '');
            const apiEv = Calendar.Events.get(CONFIG.CALENDAR_ID, rawCalId);
            if (apiEv && apiEv.attendees) {
              const respMap = { 'ACCEPTED': 'accepted', 'TENTATIVE': 'tentative', 'SUB': 'declined' };
              let changed = false;
              for (let k = 0; k < apiEv.attendees.length; k++) {
                if (String(apiEv.attendees[k].email || '').toLowerCase() === teacherEmail) {
                  apiEv.attendees[k].responseStatus = respMap[statusToSet] || 'accepted';
                  changed = true;
                }
              }
              if (changed) {
                Calendar.Events.patch({ attendees: apiEv.attendees }, CONFIG.CALENDAR_ID, rawCalId, { sendUpdates: 'none' });
              }
            }
          }
        } catch (calApiErr) {
          // Advanced Calendar service is optional; ignore if not enabled
        }
      }
    }
  } catch (err) {
    console.error('syncSingleAssignmentToCalendarAndRsvp_ error: ' + err);
  }

  return {
    gcalTemplateUrl: gcalTemplateUrl,
    formattedDate: formattedDate,
    room: room,
    teacher: teacher,
    lessonTitle: lessonTitle
  };
}

function handleEmailRsvpAction_(rowParam, rsvpParam) {
  const rowNum = Number(rowParam);
  const rsvp = String(rsvpParam || '').trim().toUpperCase();
  let statusToSet = 'ACCEPTED';
  let statusHeading = 'RSVP Confirmed: Accepted!';
  let statusSub = 'Thank you! Your teaching assignment is marked as ACCEPTED on the schedule and synced to your Google Calendar for 11:30 AM – 12:00 PM.';
  let badgeColor = '#2e6f40';

  if (rsvp === 'NO' || rsvp === 'SUB') {
    statusToSet = 'SUB';
    statusHeading = 'Substitute Request Logged';
    statusSub = 'We have marked your assignment as needing a substitute (SUB) and notified the Sunday School Presidency & Secretary.';
    badgeColor = '#9E4242';
  } else if (rsvp === 'MAYBE' || rsvp === 'TENTATIVE') {
    statusToSet = 'TENTATIVE';
    statusHeading = 'RSVP Recorded: Maybe';
    statusSub = 'Your response has been recorded as TENTATIVE on the schedule and synced to your Google Calendar.';
    badgeColor = '#555888';
  }

  let calInfo = null;
  try {
    if (statusToSet === 'SUB') {
      requestSubFromApp(rowNum);
      calInfo = syncSingleAssignmentToCalendarAndRsvp_(rowNum, statusToSet);
    } else {
      // updateTeacherRsvpFromApp already syncs the calendar (inside the lock) and returns the details.
      calInfo = updateTeacherRsvpFromApp(rowNum, statusToSet).calInfo;
    }
  } catch (err) {
    console.error('Error handling email RSVP: ' + err);
    return renderRsvpMessagePage_('We could not record your response',
      'Please open the Sunday School app to respond, or contact the Sunday School Secretary.');
  }
  calInfo = calInfo || { gcalTemplateUrl: 'https://calendar.google.com' };
  const webAppUrl = getAppUrl_();
  const baseAppUrl = webAppUrl ? (webAppUrl + (webAppUrl.indexOf('?') === -1 ? '?page=calendar' : '&page=calendar')) : '';
  const showCalendarSaveBtn = (statusToSet === 'ACCEPTED' || statusToSet === 'TENTATIVE') && calInfo.gcalTemplateUrl;

  const html = `<!DOCTYPE html>
  <html>
  <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Valley View Sunday School — RSVP</title>
  </head>
  <body style="margin:0;padding:32px 16px;background-color:#f0f2f5;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif;color:#1e242e;text-align:center;">
    <div style="max-width:420px;margin:32px auto;background:#ffffff;border-radius:20px;padding:28px 24px;box-shadow:0 8px 24px rgba(0,0,0,0.08);border:1px solid #dcdfe6;">
      <img src="${CONFIG.TEMPLE_ICON_DARK_URL || 'https://drive.google.com/thumbnail?id=13VY5W5lMc8pKWCv-dsudf_FT3tAV2tSU&sz=w256'}" width="22" height="38" alt="Temple" style="display:block;margin:0 auto 12px;" />
      <div style="font-size:13px;font-weight:800;color:#495085;text-transform:uppercase;letter-spacing:0.4px;margin-bottom:8px;">Valley View Sunday School</div>
      <div style="display:inline-block;background-color:${badgeColor};color:#ffffff;font-size:15px;font-weight:800;padding:8px 18px;border-radius:999px;margin-bottom:14px;">
        ${statusHeading}
      </div>
      <p style="font-size:13.5px;color:#4b5563;line-height:1.6;margin:0 0 18px 0;">
        ${statusSub}
      </p>
      ${calInfo.formattedDate ? `
      <div style="background-color:#9E4242;border-radius:14px;padding:12px;margin-bottom:18px;text-align:left;">
        <div style="font-size:11.5px;font-weight:700;color:#ffffff;margin-bottom:6px;">Synced Assignment (11:30 AM – 12:00 PM)</div>
        <div style="background-color:#F4F4F4;border-radius:10px;padding:10px 12px;color:#232323;font-size:12px;line-height:1.65;">
          <div><strong>Date:</strong> ${escapeHtml_(calInfo.formattedDate)}</div>
          <div><strong>Room:</strong> ${escapeHtml_(calInfo.room)}</div>
          ${calInfo.lessonTitle ? `<div><strong>Lesson:</strong> ${escapeHtml_(calInfo.lessonTitle)}</div>` : ''}
        </div>
      </div>` : ''}
      ${showCalendarSaveBtn ? `
      <div style="margin-bottom:12px;">
        <a href="${calInfo.gcalTemplateUrl}" target="_blank" style="display:block;background-color:#2e6f40;color:#ffffff;text-decoration:none;font-weight:700;font-size:13px;padding:11px 20px;border-radius:999px;box-shadow:0 3px 10px rgba(46,111,64,0.25);">
          📅 View / Save in My Google Calendar
        </a>
      </div>` : ''}
      ${baseAppUrl ? `
      <div>
        <a href="${baseAppUrl}" style="display:inline-block;background-color:#555888;color:#ffffff;text-decoration:none;font-weight:700;font-size:12.5px;padding:9px 20px;border-radius:999px;">
          Open Sunday School Calendar
        </a>
      </div>` : ''}
    </div>
  </body>
  </html>`;

  return HtmlService.createHtmlOutput(html)
    .setTitle('Valley View Sunday School — RSVP')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function buildStyledReminderEmailHtml_(data) {
  const parsed = parseLessonTitleAndScripture_(data.lessonRaw);
  const webAppUrl = getSafeWebAppUrl_();
  const appUrl = getAppUrl_();
  const isLight = (CONFIG.EMAIL_THEME || 'light').toLowerCase() === 'light';
  const isInvite = Boolean(data.isInvite);

  const lessonButton = data.lessonUrl
    ? `<a href="${data.lessonUrl}" target="_blank" style="display:inline-block;background-color:#282a3e;color:#ffffff;text-decoration:none;font-weight:700;font-size:12px;padding:7px 18px;border-radius:999px;box-shadow:0 2px 6px rgba(0,0,0,0.25);white-space:nowrap;text-align:center;">View Lesson</a>`
    : '';
  const appButton = appUrl
    ? `<a href="${appUrl}" style="display:inline-block;background-color:#9e4340;color:#ffffff;text-decoration:none;font-weight:700;font-size:11px;padding:7px 16px;border-radius:999px;">Open Web App</a>`
    : '';

  const sep = (webAppUrl && webAppUrl.indexOf('?') !== -1) ? '&' : '?';
  const rsvpId = data.assignmentKey ? rsvpId_(data.assignmentKey) : '';
  const rsvpLink = function(action) {
    return (webAppUrl && rsvpId)
      ? `${webAppUrl}${sep}rsvp=${action}&a=${rsvpId}&s=${signRsvp_(rsvpId, action)}`
      : '#';
  };
  const yesUrl = rsvpLink('YES');
  const maybeUrl = rsvpLink('MAYBE');
  const noUrl = rsvpLink('NO');
  const gcalAddUrl = isInvite ? buildGoogleCalendarAddUrl_(data) : 'https://calendar.google.com';

  const rsvpSectionLight = (isInvite && webAppUrl && data.assignmentKey) ? `
          <!-- Interactive Calendar Invite RSVP Card -->
          <div style="background-color:#f8f9fc;border:1px solid #dce1f0;border-radius:16px;padding:16px;margin-bottom:16px;text-align:center;">
            <div style="font-size:12.5px;font-weight:800;color:#1f242e;margin-bottom:4px;">
              Can you teach this Sunday School lesson?
            </div>
            <div style="font-size:11.5px;color:#6b7280;margin-bottom:12px;">
              Tap below to RSVP directly to the Sunday School schedule:
            </div>
            <table cellpadding="0" cellspacing="0" border="0" align="center" style="margin:0 auto 12px;">
              <tr>
                <td style="padding:0 4px;">
                  <a href="${yesUrl}" target="_blank" style="display:inline-block;background-color:#2e6f40;color:#ffffff;text-decoration:none;font-weight:700;font-size:12px;padding:8px 16px;border-radius:999px;">Yes, I’ll Teach</a>
                </td>
                <td style="padding:0 4px;">
                  <a href="${maybeUrl}" target="_blank" style="display:inline-block;background-color:#555888;color:#ffffff;text-decoration:none;font-weight:700;font-size:12px;padding:8px 14px;border-radius:999px;">Maybe</a>
                </td>
                <td style="padding:0 4px;">
                  <a href="${noUrl}" target="_blank" style="display:inline-block;background-color:#9E4242;color:#ffffff;text-decoration:none;font-weight:700;font-size:12px;padding:8px 14px;border-radius:999px;">Need a Sub</a>
                </td>
              </tr>
            </table>
            <div>
              <a href="${gcalAddUrl}" target="_blank" style="font-size:11.5px;font-weight:700;color:#495085;text-decoration:underline;">
                📅 Add Event to My Google Calendar (11:30 AM – 12:00 PM)
              </a>
            </div>
          </div>` : '';

  const rsvpSectionDark = (isInvite && webAppUrl && data.assignmentKey) ? `
        <!-- Interactive Calendar Invite RSVP Card -->
        <div style="background-color:#2c2c2c;border:1px solid #444654;border-radius:16px;padding:16px;margin-bottom:16px;text-align:center;">
          <div style="font-size:12.5px;font-weight:800;color:#ffffff;margin-bottom:4px;">
            Can you teach this Sunday School lesson?
          </div>
          <div style="font-size:11.5px;color:#a3a3a3;margin-bottom:12px;">
            Tap below to RSVP directly to the Sunday School schedule:
          </div>
          <table cellpadding="0" cellspacing="0" border="0" align="center" style="margin:0 auto 12px;">
            <tr>
              <td style="padding:0 4px;">
                <a href="${yesUrl}" target="_blank" style="display:inline-block;background-color:#2e6f40;color:#ffffff;text-decoration:none;font-weight:700;font-size:12px;padding:8px 16px;border-radius:999px;">Yes, I’ll Teach</a>
              </td>
              <td style="padding:0 4px;">
                <a href="${maybeUrl}" target="_blank" style="display:inline-block;background-color:#555888;color:#ffffff;text-decoration:none;font-weight:700;font-size:12px;padding:8px 14px;border-radius:999px;">Maybe</a>
              </td>
              <td style="padding:0 4px;">
                <a href="${noUrl}" target="_blank" style="display:inline-block;background-color:#9E4242;color:#ffffff;text-decoration:none;font-weight:700;font-size:12px;padding:8px 14px;border-radius:999px;">Need a Sub</a>
              </td>
            </tr>
          </table>
          <div>
            <a href="${gcalAddUrl}" target="_blank" style="font-size:11.5px;font-weight:700;color:#9eb0ff;text-decoration:underline;">
              📅 Add Event to My Google Calendar (11:30 AM – 12:00 PM)
            </a>
          </div>
        </div>` : '';

  const aidsList = parseTeachingAidsList_(data.aidsRaw, data.dateKey);

  const resolveEmailAidIconUrl = function(aid) {
    const raw = String((aid && aid.icon) || '').trim();
    if (raw && (raw.indexOf('http://') === 0 || raw.indexOf('https://') === 0)) return raw;
    return (aid && aid.faviconUrl) ? aid.faviconUrl : buildAutoFaviconUrl_(aid && aid.url);
  };

  const aidsHtmlLight = aidsList.map(function(aid, idx) {
    const isLast = idx === aidsList.length - 1;
    const iconUrl = resolveEmailAidIconUrl(aid);
    const iconCell = iconUrl ? `
      <td valign="middle" width="32" style="padding-right:10px;">
        <div style="width:28px;height:28px;border-radius:6px;background-color:#f0f2f5;text-align:center;line-height:28px;">
          <img src="${escapeHtml_(iconUrl)}" width="18" height="18" alt="" style="vertical-align:middle;border:0;border-radius:3px;" />
        </div>
      </td>` : '';
    return `
      <div style="padding:8px 0;${isLast ? '' : 'border-bottom:1px solid #f0f2f5;'}">
        <table cellpadding="0" cellspacing="0" border="0" width="100%">
          <tr>
            ${iconCell}
            <td valign="middle">
              <div style="margin-bottom:2px;">
                <a href="${aid.url}" target="_blank" style="font-size:12.5px;font-weight:700;color:#495085;text-decoration:underline;">
                  ${escapeHtml_(aid.name)}
                </a>
              </div>
              <div style="font-size:11.5px;color:#6b7280;line-height:1.4;">
                ${escapeHtml_(aid.description)}
              </div>
            </td>
          </tr>
        </table>
      </div>`;
  }).join('');

  const aidsHtmlDark = aidsList.map(function(aid, idx) {
    const isLast = idx === aidsList.length - 1;
    const iconUrl = resolveEmailAidIconUrl(aid);
    const iconCell = iconUrl ? `
      <td valign="middle" width="32" style="padding-right:10px;">
        <div style="width:28px;height:28px;border-radius:6px;background-color:#333333;text-align:center;line-height:28px;">
          <img src="${escapeHtml_(iconUrl)}" width="18" height="18" alt="" style="vertical-align:middle;border:0;border-radius:3px;" />
        </div>
      </td>` : '';
    return `
      <div style="padding:8px 0;${isLast ? '' : 'border-bottom:1px solid #383838;'}">
        <table cellpadding="0" cellspacing="0" border="0" width="100%">
          <tr>
            ${iconCell}
            <td valign="middle">
              <div style="margin-bottom:2px;">
                <a href="${aid.url}" target="_blank" style="font-size:12.5px;font-weight:700;color:#9eb0ff;text-decoration:underline;">
                  ${escapeHtml_(aid.name)}
                </a>
              </div>
              <div style="font-size:11.5px;color:#a3a3a3;line-height:1.4;">
                ${escapeHtml_(aid.description)}
              </div>
            </td>
          </tr>
        </table>
      </div>`;
  }).join('');

  if (isLight) {
    return `
    <div style="margin:0;padding:28px 12px;background-color:#f0f2f5;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif;color:#1e242e;">
      <div style="max-width:440px;margin:0 auto;background-color:#ffffff;border-radius:20px;overflow:hidden;border:1px solid #dcdfe6;box-shadow:0 8px 24px rgba(0,0,0,0.08);">
        
        <!-- Top Header Bar -->
        <table cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color:#ffffff;padding:16px 20px;border-bottom:1px solid #eaecef;">
          <tr>
            <td valign="middle" style="padding:0;">
              <table cellpadding="0" cellspacing="0" border="0">
                <tr>
                  <td valign="middle" style="padding-right:16px;">
                    <img src="${CONFIG.TEMPLE_ICON_DARK_URL || 'https://drive.google.com/thumbnail?id=13VY5W5lMc8pKWCv-dsudf_FT3tAV2tSU&sz=w256'}" width="18" height="31" alt="Temple" style="display:block;border:0;outline:none;" />
                  </td>
                  <td valign="middle" style="font-size:15px;font-weight:800;letter-spacing:-0.2px;color:#1f242e;line-height:1.2;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;">
                    Valley View<br>Sunday School
                  </td>
                </tr>
              </table>
            </td>
            <td align="right" valign="middle" style="padding:0;">
              ${appButton}
            </td>
          </tr>
        </table>

        <div style="padding:22px;">
          <p style="margin:0 0 6px 0;font-size:12px;font-weight:700;color:#495085;letter-spacing:0.3px;text-transform:uppercase;text-align:center;">${data.badgeText}</p>
          <h2 style="margin:0 0 16px 0;font-size:22px;font-weight:800;color:#111827;letter-spacing:-0.3px;text-align:center;">Hey ${data.teacher},</h2>

          <!-- Lesson Card (Signature Slate Purple Card #555888) -->
          <div style="background-color:#555888;border-radius:16px;padding:18px;margin-bottom:16px;box-shadow:0 4px 14px rgba(85,88,136,0.25);">
            <div style="font-size:11.5px;color:#e0e2f5;margin-bottom:5px;font-weight:600;">${isInvite ? 'You Are Invited to Teach' : 'Your Upcoming Lesson'}</div>
            <div style="font-family:'Source Serif Pro',Georgia,'Times New Roman',serif;font-size:20px;font-weight:900;font-style:italic;color:#ffffff;line-height:1.25;margin-bottom:12px;">
              ${parsed.title}
            </div>
            <table cellpadding="0" cellspacing="0" border="0" width="100%">
              <tr>
                <td valign="middle" style="font-size:12px;color:#d5d7eb;font-weight:500;padding-right:12px;line-height:1.35;">
                  ${parsed.scripture || data.lessonRaw}
                </td>
                <td align="right" valign="middle" style="width:1%;white-space:nowrap;">
                  ${lessonButton}
                </td>
              </tr>
            </table>
          </div>

          ${rsvpSectionLight}

          <!-- Schedule Details Card (#9E4242 Red Container with #F4F4F4 Inside Box) -->
          <div style="background-color:#9E4242;border-radius:16px;padding:16px;margin-bottom:16px;box-shadow:0 4px 14px rgba(158,66,66,0.22);">
            <div style="font-size:13px;font-weight:700;color:#ffffff;letter-spacing:0.2px;margin-bottom:10px;">Teaching Assignment Details</div>
            <div style="background-color:#F4F4F4;border-radius:12px;padding:14px;color:#232323;font-size:12.5px;line-height:1.8;">
              <div><span style="color:#555555;font-weight:600;">Date:</span> <strong>${data.formattedDate}</strong></div>
              <div><span style="color:#555555;font-weight:600;">Time:</span> <strong>11:30 AM – 12:00 PM</strong> (Utah time)</div>
              <div><span style="color:#555555;font-weight:600;">Room / Side:</span> <strong>${data.room}</strong></div>
              <div><span style="color:#555555;font-weight:600;">Conductor:</span> <strong>${data.conductor}</strong></div>
            </div>
          </div>

          <!-- Teaching Aids Section -->
          <div style="background-color:#ffffff;border:1px solid #e5e7eb;border-radius:16px;padding:16px;margin-bottom:16px;box-shadow:0 2px 8px rgba(0,0,0,0.04);">
            <div style="font-size:13px;font-weight:700;color:#374151;margin-bottom:10px;letter-spacing:-0.1px;">
              Want some lesson prep help? Check out these teaching aids
            </div>
            ${aidsHtmlLight}
          </div>

          <!-- Need a Sub Footer Box -->
          <div style="background-color:#f1f4fa;border:1px solid #dce3f0;border-radius:14px;padding:14px;font-size:12px;color:#4b5563;line-height:1.55;">
            <strong style="color:#1f2937;display:block;margin-bottom:4px;">Need to Reschedule or Request a Sub?</strong>
            You can request a substitute directly with one tap in the <a href="${appUrl}" style="color:#2563eb;font-weight:600;text-decoration:underline;">Sunday School App</a>, or contact Secretary <strong>${data.secName}</strong> at ${data.secPhone} (<a href="mailto:${data.secEmail}" style="color:#2563eb;font-weight:600;text-decoration:none;">${data.secEmail}</a>).
          </div>

          <p style="margin:18px 0 0 0;font-size:11.5px;color:#9ca3af;text-align:center;">
            Thank you for your devoted service in Valley View Sunday School!
          </p>
        </div>
      </div>
    </div>`;
  }

  // Dark Mode Template
  return `
  <div style="margin:0;padding:24px 12px;background-color:#141414;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif;color:#ffffff;">
    <div style="max-width:440px;margin:0 auto;background-color:#232323;border-radius:20px;overflow:hidden;border:1px solid #383838;box-shadow:0 12px 32px rgba(0,0,0,0.5);">
      
      <!-- Top App Header -->
      <table cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color:#2c2c2c;padding:16px 20px;border-bottom:1px solid #3a3a3a;">
        <tr>
          <td valign="middle" style="padding:0;">
            <table cellpadding="0" cellspacing="0" border="0">
              <tr>
                <td valign="middle" style="padding-right:16px;">
                  <img src="${CONFIG.TEMPLE_ICON_WHITE_URL || 'https://drive.google.com/thumbnail?id=1Il4a6rnbJI932z7VmTEEpa74muj1tngp&sz=w256'}" width="18" height="31" alt="Temple" style="display:block;border:0;outline:none;" />
                </td>
                <td valign="middle" style="font-size:15px;font-weight:800;letter-spacing:-0.2px;color:#ffffff;line-height:1.2;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;">
                  Valley View<br>Sunday School
                </td>
              </tr>
            </table>
          </td>
          <td align="right" valign="middle" style="padding:0;">
            ${appButton}
          </td>
        </tr>
      </table>

      <div style="padding:22px;">
        <p style="margin:0 0 6px 0;font-size:12px;font-weight:600;color:#9eb0ff;letter-spacing:0.2px;text-align:center;">${data.badgeText}</p>
        <h2 style="margin:0 0 16px 0;font-size:22px;font-weight:800;color:#ffffff;letter-spacing:-0.3px;text-align:center;">Hey ${data.teacher},</h2>

        <!-- Lesson Card -->
        <div style="background-color:#555888;border-radius:16px;padding:18px;margin-bottom:16px;box-shadow:0 6px 16px rgba(0,0,0,0.25);">
          <div style="font-size:11.5px;color:#e0e2f5;margin-bottom:5px;font-weight:500;">${isInvite ? 'You Are Invited to Teach' : 'Your Upcoming Lesson'}</div>
          <div style="font-family:'Source Serif Pro',Georgia,'Times New Roman',serif;font-size:20px;font-weight:900;font-style:italic;color:#ffffff;line-height:1.25;margin-bottom:12px;">
            ${parsed.title}
          </div>
          <table cellpadding="0" cellspacing="0" border="0" width="100%">
            <tr>
              <td valign="middle" style="font-size:12px;color:#d5d7eb;font-weight:500;padding-right:12px;line-height:1.35;">
                ${parsed.scripture || data.lessonRaw}
              </td>
              <td align="right" valign="middle" style="width:1%;white-space:nowrap;">
                ${lessonButton}
              </td>
            </tr>
          </table>
        </div>

        ${rsvpSectionDark}

        <!-- Schedule Details Card (#9E4242 Red Container with #F4F4F4 Inside Box) -->
        <div style="background-color:#9E4242;border-radius:16px;padding:16px;margin-bottom:16px;box-shadow:0 4px 14px rgba(0,0,0,0.35);">
          <div style="font-size:13px;font-weight:700;color:#ffffff;letter-spacing:0.2px;margin-bottom:10px;">Teaching Assignment Details</div>
          <div style="background-color:#F4F4F4;border-radius:12px;padding:14px;color:#232323;font-size:12.5px;line-height:1.8;">
            <div><span style="color:#555555;font-weight:600;">Date:</span> <strong>${data.formattedDate}</strong></div>
            <div><span style="color:#555555;font-weight:600;">Time:</span> <strong>11:30 AM – 12:00 PM</strong> (Utah time)</div>
            <div><span style="color:#555555;font-weight:600;">Room / Side:</span> <strong>${data.room}</strong></div>
            <div><span style="color:#555555;font-weight:600;">Conductor:</span> <strong>${data.conductor}</strong></div>
          </div>
        </div>

        <!-- Teaching Aids Section -->
        <div style="background-color:#2c2c2c;border:1px solid #3d3d3d;border-radius:16px;padding:16px;margin-bottom:16px;">
          <div style="font-size:13px;font-weight:700;color:#ffffff;margin-bottom:10px;letter-spacing:-0.1px;">
            Want some lesson prep help? Check out these teaching aids
          </div>
          ${aidsHtmlDark}
        </div>

        <!-- Need a Sub Footer Box -->
        <div style="background-color:#2c2f36;border:1px solid #3d424d;border-radius:14px;padding:14px;font-size:11.5px;color:#cccccc;line-height:1.55;">
          <strong style="color:#ffffff;display:block;margin-bottom:4px;">Need to Reschedule or Request a Sub?</strong>
          You can request a substitute directly with one tap in the <a href="${appUrl}" style="color:#7da5ff;text-decoration:underline;">Sunday School App</a>, or contact Secretary <strong>${data.secName}</strong> at ${data.secPhone} (<a href="mailto:${data.secEmail}" style="color:#7da5ff;text-decoration:none;">${data.secEmail}</a>).
        </div>

        <p style="margin:16px 0 0 0;font-size:11px;color:#7a7a7a;text-align:center;">
          Thank you for your devoted service in Valley View Sunday School!
        </p>
      </div>
    </div>
  </div>`;
}

function sendSubAlertEmail_(contacts, teacher, dateKey, room, lessonTitle) {
  const dateForDisplay = localDateTime_(dateKey, 12, 0);
  const formattedDate = Utilities.formatDate(dateForDisplay, CONFIG.TIME_ZONE, 'EEEE, MMMM d, yyyy');
  const subject = `🚨 ACTION REQUIRED: Sunday School Substitute Needed — ${room}`;
  const plainBody =
    'Hello,\n\n' +
    'A scheduled Sunday School teacher has requested a substitute or declined their assignment.\n\n' +
    `• Teacher: ${teacher}\n` +
    `• Date: ${formattedDate}\n` +
    `• Room: ${room}\n` +
    (lessonTitle ? `• Lesson: ${lessonTitle}\n\n` : '\n') +
    'Please coordinate a substitute and update the Google Sheet with the replacement teacher so a new calendar invitation can be generated.\n\n' +
    'Thank you.';

  const sheetUrl = SpreadsheetApp.getActiveSpreadsheet() ? SpreadsheetApp.getActiveSpreadsheet().getUrl() : '';
  const htmlBody = `
  <div style="margin:0;padding:24px 12px;background-color:#1a1a1a;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif;color:#ffffff;">
    <div style="max-width:440px;margin:0 auto;background-color:#232323;border-radius:20px;overflow:hidden;border:1px solid #383838;">
      <div style="background-color:#9e4340;padding:18px 22px;">
        <div style="font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:0.6px;color:#f9d6d5;">Substitute Needed</div>
        <div style="font-size:20px;font-weight:800;color:#ffffff;margin-top:4px;">Valley View Sunday School</div>
      </div>
      <div style="padding:22px;">
        <p style="margin:0 0 14px 0;font-size:14px;color:#e0e0e0;line-height:1.5;">
          <strong>${teacher}</strong> has requested a substitute (or declined the calendar invitation) for their upcoming teaching assignment:
        </p>
        <div style="background-color:#3d3d3d;border-radius:14px;padding:16px;margin-bottom:18px;">
          <div style="background-color:#f5f5f5;border-radius:10px;padding:14px;color:#232323;font-size:13px;line-height:1.7;">
            <div><strong>Teacher:</strong> ${teacher}</div>
            <div><strong>Date:</strong> ${formattedDate}</div>
            <div><strong>Room:</strong> ${room}</div>
            ${lessonTitle ? `<div><strong>Lesson:</strong> ${lessonTitle}</div>` : ''}
            <div><strong>Sheet Status:</strong> SUB</div>
          </div>
        </div>
        ${sheetUrl ? `<a href="${sheetUrl}" style="display:block;text-align:center;background-color:#555888;color:#ffffff;text-decoration:none;font-weight:700;font-size:14px;padding:12px 20px;border-radius:999px;">Open Schedule Google Sheet</a>` : ''}
      </div>
    </div>
  </div>`;

  try {
    if (contacts && contacts.alertEmails && contacts.alertEmails.length > 0) {
      MailApp.sendEmail(contacts.alertEmails, subject, plainBody, mailOptions_({ htmlBody: htmlBody }, contacts));
    }
    return true;
  } catch (error) {
    console.error(`Substitute alert failed for ${teacher} on ${dateKey}, ${room}. Error: ${error}`);
    return false;
  }
}

/* ============================================================================
 * WEB APP BACKEND API (Called via doGet and google.script.run)
 * ============================================================================ */

function getWebAppBootData(forceRefresh) {
  const cache = CacheService.getScriptCache();
  if (!forceRefresh) {
    try {
      const cachedJson = cache.get(CONFIG.BOOT_CACHE_KEY);
      if (cachedJson) {
        return JSON.parse(cachedJson);
      }
    } catch (e) {}
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const schedSheet = getScheduleSheet_();
  if (!schedSheet) throw new Error('Schedule sheet not found.');

  const upcomingSundayKey = getUpcomingSundayDateKey_();
  const lastRow = schedSheet.getLastRow();
  const maxCols = Math.max(9, schedSheet.getLastColumn());
  const values = lastRow >= 2 ? schedSheet.getRange(2, 1, lastRow - 1, maxCols).getValues() : [];
  const richTexts = lastRow >= 2 ? schedSheet.getRange(2, 6, lastRow - 1, 1).getRichTextValues() : [];

  const schedule = [];
  const teachersSet = {};

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const room = String(row[0] || '').trim();
    const status = String(row[1] || '').trim().toUpperCase();
    const rawTeacher = String(row[2] || '').trim();
    const teacher = (rawTeacher.toUpperCase() === 'N/A') ? '' : rawTeacher;
    const conductor = String(row[3] || '').trim();
    const rawDate = row[4];
    const lessonRaw = String(row[5] || '').trim();
    const teacherEmail = String(row[6] || '').trim().toLowerCase();
    const aidsRaw = String(row[8] || '').trim();

    if (!(rawDate instanceof Date && !isNaN(rawDate.getTime())) && !lessonRaw) continue;

    let dateKey = '';
    let dateLabel = '';
    if (rawDate instanceof Date && !isNaN(rawDate.getTime())) {
      dateKey = sheetDateKey_(rawDate, schedSheet);
      dateLabel = formatShortDisplayDate_(rawDate, schedSheet.getParent().getSpreadsheetTimeZone());
    } else {
      dateKey = String(rawDate || '');
      dateLabel = dateKey;
    }

    const lessonUrl = extractRichTextUrl_(richTexts[i] ? richTexts[i][0] : null) || 'https://www.churchofjesuschrist.org/study/come-follow-me?lang=eng';
    const parsed = parseLessonTitleAndScripture_(lessonRaw);
    const roomShort = room.toLowerCase().includes('elders') ? 'EQ' : (room.toLowerCase().includes('relief') ? 'RS' : room);

    if (teacher && teacher.toUpperCase() !== 'N/A') {
      teachersSet[teacher] = true;
    }

    schedule.push({
      rowNumber: i + 2,
      room: room,
      roomShort: roomShort,
      status: status,
      teacher: teacher,
      conductor: conductor,
      dateKey: dateKey,
      dateLabel: dateLabel,
      lessonRaw: lessonRaw,
      lessonTitle: parsed.title,
      scriptureBlock: parsed.scripture || parsed.dateRange,
      lessonUrl: lessonUrl,
      lessonImage: '',
      aidsRaw: aidsRaw
    });
  }

  // Contacts & Teachers
  const contactsList = [];
  const contactSheet = ss.getSheetByName(CONFIG.CONTACT_SHEET_NAME);
  if (contactSheet && contactSheet.getLastRow() >= 2) {
    const cVals = contactSheet.getRange(2, 1, contactSheet.getLastRow() - 1, 4).getValues();
    for (const r of cVals) {
      const calling = String(r[0] || '').trim();
      const name = String(r[1] || '').trim();
      const phone = String(r[2] || '').trim();
      const email = String(r[3] || '').trim();
      if (!name) continue;
      contactsList.push({ calling: calling, name: name }); // phone/email intentionally not sent to the browser
      if (!calling.toLowerCase().includes('bishopric')) {
        teachersSet[name] = true;
      }
    }
  }

  // Monthly Training
  const monthlyTrainings = [];
  const mtSheet = ss.getSheetByName(CONFIG.MONTHLY_TRAINING_SHEET_NAME);
  if (mtSheet && mtSheet.getLastRow() >= 2) {
    const mtVals = mtSheet.getRange(2, 1, mtSheet.getLastRow() - 1, 5).getValues();
    const mtRich = mtSheet.getRange(2, 4, mtSheet.getLastRow() - 1, 1).getRichTextValues();
    for (let i = 0; i < mtVals.length; i++) {
      const mVal = mtVals[i][0];
      const tName = String(mtVals[i][1] || '').trim();
      const sub = String(mtVals[i][2] || "Teaching in the Savior's way").trim();
      const urlText = String(mtVals[i][3] || '').trim();
      const richUrl = extractRichTextUrl_(mtRich[i] ? mtRich[i][0] : null);
      const studyUrl = richUrl || urlText || 'https://www.churchofjesuschrist.org/study/manual/teaching-in-the-saviors-way';
      const imgUrl = String(mtVals[i][4] || '').trim();
      if (!tName) continue;

      let monthKey = '';
      let monthLabel = '';
      if (mVal instanceof Date && !isNaN(mVal.getTime())) {
        monthKey = Utilities.formatDate(mVal, CONFIG.TIME_ZONE, 'yyyy-MM');
        monthLabel = Utilities.formatDate(mVal, CONFIG.TIME_ZONE, 'MMMM yyyy');
      } else {
        monthKey = String(mVal || '');
        monthLabel = monthKey;
      }

      const formattedTitle = (tName.startsWith('“') || tName.startsWith('"')) ? tName : `“${tName}”`;
      monthlyTrainings.push({
        monthKey: monthKey,
        monthLabel: monthLabel,
        title: formattedTitle,
        subtitle: sub,
        url: studyUrl,
        imageUrl: imgUrl || fetchLinkPreviewImage_(studyUrl)
      });
    }
  }

  // Unified Teaching Aids tab (Layout Option A: Lesson Date or "Standard" | Title | Description | Link URL | Custom Icon)
  const teachingAids = readTeachingAidsSheet_(ss);

  // Fetch OpenGraph preview image for the current upcoming lesson and next few lessons (cached)
  const seenUrls = {};
  for (let i = 0; i < schedule.length; i++) {
    const item = schedule[i];
    if (item.dateKey >= upcomingSundayKey && Object.keys(seenUrls).length < 6) {
      if (!seenUrls[item.lessonUrl]) {
        seenUrls[item.lessonUrl] = fetchLinkPreviewImage_(item.lessonUrl);
      }
      item.lessonImage = seenUrls[item.lessonUrl] || '';
    }
  }

  const bootPayload = {
    upcomingSundayKey: upcomingSundayKey,
    schedule: schedule,
    teachers: Object.keys(teachersSet),
    contacts: contactsList,
    monthlyTrainings: monthlyTrainings,
    teachingAids: teachingAids,
    attendanceBaseUrl: getAttendanceCheckInBaseUrl_(),
    defaultLessonImageDriveId: CONFIG.FALLBACK_IMAGE_DRIVE_ID,
    appIconDriveId: CONFIG.APP_ICON_DRIVE_ID
  };

  try {
    const serialized = JSON.stringify(bootPayload);
    if (serialized.length < 95000) {
      cache.put(CONFIG.BOOT_CACHE_KEY, serialized, CONFIG.BOOT_CACHE_TTL_SECONDS || 90);
    }
  } catch (e) {}

  return bootPayload;
}

/**
 * Reads and validates one schedule row. Throws if the row is not an active assignment, or if the
 * caller's expected date/room no longer match (stale page, sorted sheet, etc.).
 */
function validateAssignmentRow_(sheet, rowNumber, expectedDateKey, expectedRoom) {
  const r = Number(rowNumber);
  if (!r || r < 2 || r > sheet.getLastRow()) throw new Error('Invalid schedule row.');
  const v = sheet.getRange(r, 1, 1, 7).getValues()[0];
  const room = String(v[0] || '').trim();
  const teacher = String(v[2] || '').trim();
  const rawDate = v[4];
  const email = String(v[6] || '').trim().toLowerCase();
  if (!isValidAssignment_(teacher, email, rawDate) || !room) {
    throw new Error('That row is not an active teaching assignment.');
  }
  const dateKey = sheetDateKey_(rawDate, sheet);
  if (expectedDateKey && String(expectedDateKey) !== dateKey) {
    throw new Error('The schedule changed. Please refresh the app and try again.');
  }
  if (expectedRoom && String(expectedRoom).trim() !== room) {
    throw new Error('The schedule changed. Please refresh the app and try again.');
  }
  return { row: r, room: room, teacher: teacher, email: email, dateKey: dateKey, lessonTitle: String(v[5] || '').trim() };
}

/**
 * Handles a teacher's "Request a Sub" submission from the Profile or Calendar page.
 * Writes "SUB" to Column B in the Schedule sheet and dispatches the substitute alert email.
 */
function requestSubFromApp(rowNumber, expectedDateKey, expectedRoom) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  let a = null;
  try {
    const sheet = getScheduleSheet_();
    if (!sheet) throw new Error('Schedule sheet not found.');
    a = validateAssignmentRow_(sheet, rowNumber, expectedDateKey, expectedRoom);
    sheet.getRange(a.row, 2).setValue('SUB');
    SpreadsheetApp.flush();
    invalidateBootCache_();
  } finally {
    lock.releaseLock();
  }

  // Send substitute alert outside the sheet lock so concurrent app users never wait on MailApp
  try {
    const contacts = getContactsConfig_();
    const properties = PropertiesService.getScriptProperties();
    const alertKey = CONFIG.SUBSTITUTE_ALERT_PROPERTY_PREFIX + assignmentKey_(a.dateKey, a.room, a.email);

    if (contacts && properties.getProperty(alertKey) !== 'SENT') {
      if (sendSubAlertEmail_(contacts, a.teacher, a.dateKey, a.room, a.lessonTitle)) {
        properties.setProperty(alertKey, 'SENT');
      }
    }
  } catch (e) {
    console.error('Sub alert email error: ' + e);
  }

  return { success: true, rowNumber: a.row, status: 'SUB', teacher: a.teacher, dateKey: a.dateKey, room: a.room };
}

/**
 * Allows a teacher to set their RSVP (ACCEPTED / TENTATIVE / PENDING).
 * Releases the sheet lock immediately after writing Column B so 10+ simultaneous users never block on CalendarApp.
 */
function updateTeacherRsvpFromApp(rowNumber, newStatus, expectedDateKey, expectedRoom) {
  const normalized = String(newStatus || 'ACCEPTED').trim().toUpperCase();
  if (['ACCEPTED', 'TENTATIVE', 'PENDING'].indexOf(normalized) === -1) throw new Error('Invalid RSVP status.');

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  let a = null;
  try {
    const sheet = getScheduleSheet_();
    if (!sheet) throw new Error('Schedule sheet not found.');
    a = validateAssignmentRow_(sheet, rowNumber, expectedDateKey, expectedRoom);
    sheet.getRange(a.row, 2).setValue(normalized);
    SpreadsheetApp.flush();
    invalidateBootCache_();
  } finally {
    lock.releaseLock();
  }

  // Sync the assignment to Google Calendar after releasing the spreadsheet lock
  let calInfo = null;
  try {
    calInfo = syncSingleAssignmentToCalendarAndRsvp_(a.row, normalized);
  } catch (e) {
    console.error('Calendar sync after RSVP failed: ' + e);
  }

  return { success: true, rowNumber: a.row, status: normalized, calInfo: calInfo };
}

/**
 * Allows any teacher to contribute a weekly Teaching Aid link directly from the "+" button in the app.
 * Appends a new row to the "Teaching Aids" tab (Layout Option A) for the chosen Sunday dateKey.
 */
function addTeachingAidFromApp(payload) {
  if (!payload || typeof payload !== 'object') {
    throw new Error('Invalid teaching aid submission.');
  }
  const rawDateKey = String(payload.dateKey || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(rawDateKey)) {
    throw new Error('Please select a valid Sunday lesson date.');
  }

  let rawUrl = String(payload.url || '').trim();
  if (!rawUrl) {
    throw new Error('Please paste a link URL for the teaching aid.');
  }
  if (!/^https?:\/\//i.test(rawUrl) && rawUrl.indexOf('.') !== -1) {
    rawUrl = 'https://' + rawUrl;
  }
  if (!/^https?:\/\/[^\s]+$/i.test(rawUrl)) {
    throw new Error('Please enter a valid http:// or https:// web link.');
  }

  let title = String(payload.name || '').trim().slice(0, 140);
  if (!title) {
    title = inferAidTitleFromUrl_(rawUrl);
  }

  const sharedBy = String(payload.sharedBy || '').trim().slice(0, 60);
  let description = String(payload.description || '').trim().slice(0, 220);
  if (!description) {
    description = sharedBy ? ('Shared by ' + sharedBy) : 'Community study aid for this Sunday’s lesson';
  }

  const customIcon = String(payload.icon || '').trim().slice(0, 500);
  // Prevent formula injection when writing user strings to the sheet
  const safeCellText = function(str) {
    const s = String(str || '').trim();
    return /^[=+\-@]/.test(s) ? ("'" + s) : s;
  };

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aidsSheet = ss.getSheetByName(CONFIG.TEACHING_AIDS_SHEET_NAME);
    if (!aidsSheet) throw new Error('Teaching Aids sheet tab not found.');

    const parts = rawDateKey.split('-');
    const formattedSheetDate = Number(parts[1]) + '/' + Number(parts[2]) + '/' + parts[0];

    aidsSheet.appendRow([
      formattedSheetDate,
      safeCellText(title),
      safeCellText(description),
      rawUrl,
      safeCellText(customIcon)
    ]);
    SpreadsheetApp.flush();
    invalidateBootCache_();

    return {
      success: true,
      aid: {
        dateKey: rawDateKey,
        name: title,
        description: description,
        url: rawUrl,
        icon: customIcon,
        faviconUrl: buildAutoFaviconUrl_(rawUrl),
        isStandard: false
      }
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Determines the current upcoming Sunday date key ('yyyy-MM-dd') in America/Denver time.
 * Rule: On Sundays before 12:30 PM, today's Sunday is the upcoming lesson.
 * After 12:30 PM on Sunday (or Mon–Sat), the following Sunday is the upcoming lesson.
 */
function getUpcomingSundayDateKey_() {
  const now = new Date();
  const dayOfWeek = Number(Utilities.formatDate(now, CONFIG.TIME_ZONE, 'u')); // 1 = Mon ... 7 = Sun
  const hour = Number(Utilities.formatDate(now, CONFIG.TIME_ZONE, 'H'));
  const minute = Number(Utilities.formatDate(now, CONFIG.TIME_ZONE, 'm'));
  const todayKey = Utilities.formatDate(now, CONFIG.TIME_ZONE, 'yyyy-MM-dd');

  let daysToAdd = 0;
  if (dayOfWeek === 7) {
    // Sunday: switch after 12:30 PM
    if (hour > 12 || (hour === 12 && minute >= 30)) {
      daysToAdd = 7;
    } else {
      daysToAdd = 0;
    }
  } else {
    daysToAdd = 7 - dayOfWeek;
  }

  const parts = todayKey.split('-').map(Number);
  const target = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2] + daysToAdd));
  return Utilities.formatDate(target, 'UTC', 'yyyy-MM-dd');
}

const KNOWN_LESSON_PREVIEW_IMAGES = {
  "/study/manual/come-follow-me-for-home-and-church-old-testament-2026/14": "https://www.churchofjesuschrist.org/imgs/1c2413340f1911ef92c5eeeeac1eb3d4e4ffbb9d/full/%21800%2C/0/default",
  "/study/manual/come-follow-me-for-home-and-church-old-testament-2026/41": "https://www.churchofjesuschrist.org/imgs/d17747ea13ad11efb4bbeeeeac1ea2a0b67749b9/full/%21800%2C/0/default",
  "/study/manual/come-follow-me-for-home-and-church-old-testament-2026/42": "https://www.churchofjesuschrist.org/imgs/f7d8b44313b411efaa94eeeeac1e64314ec0298d/full/%21800%2C/0/default",
  "/study/manual/come-follow-me-for-home-and-church-old-testament-2026/43": "https://www.churchofjesuschrist.org/imgs/687a7dba13ae11efbf2feeeeac1e59432305a0b9/full/%21800%2C/0/default",
  "/study/manual/come-follow-me-for-home-and-church-old-testament-2026/44": "https://www.churchofjesuschrist.org/imgs/41295c2416c611ef867beeeeac1ed83083908adf/full/%21800%2C/0/default",
  "/study/manual/come-follow-me-for-home-and-church-old-testament-2026/45": "https://www.churchofjesuschrist.org/imgs/06873993130e11efa101eeeeac1e8344e4f8072c/full/%21800%2C/0/default",
  "/study/manual/come-follow-me-for-home-and-church-old-testament-2026/46": "https://www.churchofjesuschrist.org/imgs/09a7b376130e11efa101eeeeac1e8344d5564be6/full/%21800%2C/0/default",
  "/study/manual/come-follow-me-for-home-and-church-old-testament-2026/47": "https://www.churchofjesuschrist.org/imgs/04c11f93130e11ef8512eeeeac1ef92bdebe188c/full/%21800%2C/0/default",
  "/study/manual/come-follow-me-for-home-and-church-old-testament-2026/48": "https://www.churchofjesuschrist.org/imgs/8fd9786212fb11efb34deeeeac1e1bdcbdf3f714/full/%21800%2C/0/default",
  "/study/manual/come-follow-me-for-home-and-church-old-testament-2026/49": "https://www.churchofjesuschrist.org/imgs/122c24163de211eeb167eeeeac1e3d73bdc661c5/full/%21800%2C/0/default",
  "/study/manual/come-follow-me-for-home-and-church-old-testament-2026/50": "https://www.churchofjesuschrist.org/imgs/04f85c32130e11ef9843eeeeac1ef6279de3dd7a/full/%21800%2C/0/default",
  "/study/manual/come-follow-me-for-home-and-church-old-testament-2026/51": "https://www.churchofjesuschrist.org/imgs/a73e29f0174c6084fdfac76c89a6f538b7ae84f6/full/%21800%2C/0/default",
  "/study/manual/come-follow-me-for-home-and-church-old-testament-2026/52": "https://www.churchofjesuschrist.org/imgs/bcb38e7212fb11efa003eeeeac1e5954cb4a32f2/full/%21800%2C/0/default",
  "/study/manual/come-follow-me-for-home-and-church-new-testament-2027/01": "https://www.churchofjesuschrist.org/imgs/9l6bp4fq6rpinignvzypvk14ighnp4zq5i4ukid8/full/%21800%2C/0/default",
  "/study/manual/come-follow-me-for-home-and-church-new-testament-2027/02": "https://www.churchofjesuschrist.org/imgs/ysaa6kruu6q6uh4inoa154y716r44j3d7uz63gap/full/%21800%2C/0/default",
  "/study/manual/come-follow-me-for-home-and-church-new-testament-2027/03": "https://www.churchofjesuschrist.org/imgs/35a13198247703463bbb13c73e179aff6f2da9d2/full/%21800%2C/0/default",
  "/study/manual/come-follow-me-for-home-and-church-new-testament-2027/04": "https://www.churchofjesuschrist.org/imgs/d22c692e4be94ca356c65d2dde291a6cffae07d9/full/%21800%2C/0/default",
  "/study/manual/come-follow-me-for-home-and-church-new-testament-2027/05": "https://www.churchofjesuschrist.org/imgs/de6e55828c8fe331fab06bbd22ea9237c1fb4f8b/full/%21800%2C/0/default",
  "/study/manual/come-follow-me-for-home-and-church-new-testament-2027/06": "https://www.churchofjesuschrist.org/imgs/d55a1cc7a2b071ccacbab4ca4692b48b11eb73fa/full/%21800%2C/0/default",
  "/study/manual/come-follow-me-for-home-and-church-new-testament-2027/07": "https://www.churchofjesuschrist.org/imgs/4746527d5d73a1af740f250d0c1d2f405529c7c0/full/%21800%2C/0/default",
  "/study/manual/come-follow-me-for-home-and-church-new-testament-2027/08": "https://www.churchofjesuschrist.org/imgs/07e9e52036df11eea64deeeeac1e890de985aabd/full/%21800%2C/0/default",
  "/study/manual/come-follow-me-for-home-and-church-new-testament-2027/09": "https://www.churchofjesuschrist.org/imgs/9dr72oj7t1s4ubx91mp7pwkvjq2v1sufgennx8za/full/%21800%2C/0/default",
  "/study/manual/come-follow-me-for-home-and-church-new-testament-2027/10": "https://www.churchofjesuschrist.org/imgs/5953f6f7aed711eda61beeeeac1eaa47906f2781/full/%21800%2C/0/default",
  "/study/manual/come-follow-me-for-home-and-church-new-testament-2027/11": "https://www.churchofjesuschrist.org/imgs/ltsi1p4vegujabrwpo0k1s8lwr28id5yz38eygy3/full/%21800%2C/0/default",
  "/study/manual/come-follow-me-for-home-and-church-new-testament-2027/12": "https://www.churchofjesuschrist.org/imgs/o5l7pztdqt808trssmxftw1nsbaz14xbyjagw821/full/%21800%2C/0/default",
  "/study/manual/come-follow-me-for-home-and-church-new-testament-2027/13": "https://www.churchofjesuschrist.org/imgs/e712432bcf39035169f353cd513558b816b50bd8/full/%21800%2C/0/default",
  "/study/manual/come-follow-me-for-home-and-church-new-testament-2027/14": "https://www.churchofjesuschrist.org/imgs/1c2413340f1911ef92c5eeeeac1eb3d4e4ffbb9d/full/%21800%2C/0/default",
  "/study/manual/teaching-in-the-saviors-way": "https://www.churchofjesuschrist.org/imgs/510f01577b0a11ec9315eeeeac1e5d4d35398f79/full/%21800%2C/0/default"
};

/**
 * Public Web App helper to resolve the preview image for any lesson/study URL.
 */
function getLinkPreviewImageFromApp(url) {
  return fetchLinkPreviewImage_(url);
}

/**
 * Resolves the preview image for a lesson or training link:
 * 1. Checks known Come, Follow Me 2026-2027 lesson image map
 * 2. Checks CacheService
 * 3. For churchofjesuschrist.org/study/... links, queries the Church Content API for the lead artwork
 * 4. Scans full HTML for og:image, twitter:image, or lead churchofjesuschrist.org/imgs/ artwork
 */
function fetchLinkPreviewImage_(url) {
  if (!url || !url.startsWith('http')) return '';
  for (const pathKey in KNOWN_LESSON_PREVIEW_IMAGES) {
    if (url.indexOf(pathKey) !== -1) {
      return KNOWN_LESSON_PREVIEW_IMAGES[pathKey];
    }
  }
  try {
    const cache = CacheService.getScriptCache();
    const cacheKey = 'OG_IMG_V2_' + Utilities.base64EncodeWebSafe(url).slice(0, 80);
    const cached = cache.get(cacheKey);
    if (cached) return cached;

    // Try Church Content API first if this is a churchofjesuschrist.org/study/ URL
    const studyMatch = url.match(/churchofjesuschrist\.org\/study(\/[^?#]+)/i);
    if (studyMatch && studyMatch[1]) {
      const apiUri = 'https://www.churchofjesuschrist.org/study/api/v3/language-pages/type/content?lang=eng&uri=' + encodeURIComponent(studyMatch[1]);
      const apiResp = UrlFetchApp.fetch(apiUri, { muteHttpExceptions: true, followRedirects: true });
      if (apiResp.getResponseCode() === 200) {
        const apiText = apiResp.getContentText();
        const bodyImg = apiText.match(/https:\/\/www\.churchofjesuschrist\.org\/imgs\/[a-zA-Z0-9]+\/full\/[^"'\\\s]+/i);
        if (bodyImg && bodyImg[0] && !bodyImg[0].endsWith('.svg')) {
          const cleanUrl = bodyImg[0].replace(/&amp;/g, '&');
          cache.put(cacheKey, cleanUrl, 21600);
          return cleanUrl;
        }
      }
    }

    const resp = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      followRedirects: true,
      headers: { 'User-Agent': 'Mozilla/5.0 (compatible; GoogleAppsScript)' }
    });
    if (resp.getResponseCode() === 200) {
      const html = resp.getContentText();
      const ogMatch = html.match(/<meta[^>]+(?:property|name)=["'](?:og:image|twitter:image)["'][^>]+content=["']([^"']+)["']/i)
                   || html.match(/<meta[^>]+content=["']([^"']+)["'][^>]+(?:property|name)=["'](?:og:image|twitter:image)["']/i);
      if (ogMatch && ogMatch[1] && !ogMatch[1].endsWith('.svg') && ogMatch[1].indexOf('740e66cd0e91a92ab56b9086b632901cf2cd2c9e') === -1) {
        const imgUrl = ogMatch[1].replace(/&amp;/g, '&').trim();
        cache.put(cacheKey, imgUrl, 21600);
        return imgUrl;
      }
      const churchImgMatch = html.match(/https:\/\/www\.churchofjesuschrist\.org\/imgs\/[a-zA-Z0-9]+\/full\/[^"'\s>]+/i);
      if (churchImgMatch && churchImgMatch[0] && !churchImgMatch[0].endsWith('.svg') && churchImgMatch[0].indexOf('740e66cd0e91a92ab56b9086b632901cf2cd2c9e') === -1) {
        const imgUrl = churchImgMatch[0].replace(/&amp;/g, '&').trim();
        cache.put(cacheKey, imgUrl, 21600);
        return imgUrl;
      }
    }
  } catch (e) {
    // Ignore fetch errors and fall back to default image on client
  }
  return '';
}

/**
 * Resolves any PNG or SVG icon from the Teaching Aids sheet (Column D):
 * - In-cell CellImage or OverGridImage uploaded directly into the sheet
 * - =IMAGE("...") formula or hyperlinked cell
 * - Google Drive file link or File ID (reads PNG/SVG blob via DriveApp -> base64 data URI, or thumbnail fallback)
 * - Raw <svg>...</svg> XML pasted from Figma/Illustrator
 * - Direct .png / .svg URL or data:image/... URI
 */
function resolveTeachingAidIcon_(cellVal, richTextVal, formulaVal, overGridImg) {
  if (overGridImg) {
    try {
      const u = overGridImg.getUrl ? overGridImg.getUrl() : '';
      if (u) return u;
    } catch (e) {}
  }

  if (cellVal && typeof cellVal === 'object') {
    try {
      if (typeof cellVal.getUrl === 'function') {
        const cellImgUrl = cellVal.getUrl();
        if (cellImgUrl) return resolveDriveOrDirectIconUrl_(cellImgUrl);
      }
      if (typeof cellVal.getContentUrl === 'function') {
        const contentUrl = cellVal.getContentUrl();
        if (contentUrl) return contentUrl;
      }
    } catch (e) {}
  }

  const fStr = String(formulaVal || '').trim();
  if (fStr) {
    const m = fStr.match(/^=IMAGE\(\s*["']([^"']+)["']/i);
    if (m && m[1]) return resolveDriveOrDirectIconUrl_(m[1].trim());
  }

  const richLink = extractRichTextUrl_(richTextVal);
  const rawStr = String(cellVal || '').trim();
  const candidate = (richLink && !rawStr.startsWith('<svg') && !rawStr.startsWith('data:image/')) ? richLink : rawStr;
  if (!candidate || candidate === 'CellImage') return '';

  if (candidate.startsWith('<svg') || (candidate.indexOf('<svg') !== -1 && candidate.indexOf('</svg>') !== -1)) {
    const svgXml = candidate.slice(candidate.indexOf('<svg'));
    return 'data:image/svg+xml;base64,' + Utilities.base64Encode(Utilities.newBlob(svgXml, 'image/svg+xml').getBytes());
  }

  return resolveDriveOrDirectIconUrl_(candidate);
}

function resolveDriveOrDirectIconUrl_(str) {
  const s = String(str || '').trim();
  if (!s) return '';
  if (s.startsWith('data:image/')) return s;

  let driveId = '';
  const m = s.match(/drive\.google\.com\/(?:file\/d\/|open\?id=|uc\?(?:export=\w+&)?id=|thumbnail\?id=)([a-zA-Z0-9_-]{15,})/i);
  if (m && m[1]) {
    driveId = m[1];
  } else if (/^[a-zA-Z0-9_-]{25,45}$/.test(s)) {
    driveId = s;
  }

  if (driveId) {
    try {
      const cache = CacheService.getScriptCache();
      const cKey = 'AID_ICON_V1_' + driveId;
      const cached = cache.get(cKey);
      if (cached) return cached;

      const file = DriveApp.getFileById(driveId);
      const blob = file.getBlob();
      const mime = blob.getContentType() || (file.getName().toLowerCase().endsWith('.svg') ? 'image/svg+xml' : 'image/png');
      const bytes = blob.getBytes();
      if (bytes.length <= 70000) {
        const dataUri = 'data:' + mime + ';base64,' + Utilities.base64Encode(bytes);
        cache.put(cKey, dataUri, 21600);
        return dataUri;
      }
    } catch (e) {}
    return 'https://drive.google.com/thumbnail?id=' + driveId + '&sz=w256';
  }

  return s;
}

function extractRichTextUrl_(richText) {
  if (!richText) return '';
  const direct = richText.getLinkUrl();
  if (direct) return direct;
  const runs = richText.getRuns ? richText.getRuns() : [];
  for (let i = 0; i < runs.length; i++) {
    const u = runs[i].getLinkUrl();
    if (u) return u;
  }
  return '';
}

function parseLessonTitleAndScripture_(rawText) {
  const raw = String(rawText || '').trim();
  if (!raw) return { dateRange: '', title: '“Come, Follow Me”', scripture: '' };
  let dateRange = '';
  let rest = raw;
  const dotIdx = raw.indexOf('. ');
  if (dotIdx !== -1) {
    dateRange = raw.slice(0, dotIdx).trim();
    rest = raw.slice(dotIdx + 2).trim();
  }
  let title = rest;
  let scripture = '';
  const quoteColonIdx = Math.max(rest.indexOf('”:'), rest.indexOf('":'));
  if (quoteColonIdx !== -1) {
    title = rest.slice(0, quoteColonIdx + 1).trim();
    scripture = rest.slice(quoteColonIdx + 2).trim();
  } else if (rest.includes(':')) {
    const parts = rest.split(':');
    title = parts[0].trim();
    scripture = parts.slice(1).join(':').trim();
  }
  return { dateRange: dateRange, title: title, scripture: scripture };
}

function formatShortDisplayDate_(date, timeZone) {
  const monthIdx = Number(Utilities.formatDate(date, timeZone, 'M')) - 1;
  const day = Number(Utilities.formatDate(date, timeZone, 'd'));
  const year = Utilities.formatDate(date, timeZone, 'yyyy');
  const months = ['Jan.', 'Feb.', 'Mar.', 'Apr.', 'May', 'June', 'July', 'Aug.', 'Sept.', 'Oct.', 'Nov.', 'Dec.'];
  return `${months[monthIdx]} ${day}, ${year}`;
}

function getSafeWebAppUrl_() {
  const override = String(CONFIG.SUNDAY_SCHOOL_WEB_APP_URL || '').trim();
  if (override) return override;
  try {
    const url = ScriptApp.getService().getUrl();
    // Only trust the public /exec URL. The /dev test URL only works for the script owner.
    if (url && /\/exec$/.test(url)) return url;
  } catch (e) {}
  return '';
}

/** Menu helper: shows the addresses the emails and app will use. */
function checkWebAppSetup() {
  const api = getSafeWebAppUrl_();
  const pwa = String(CONFIG.GITHUB_PAGES_APP_URL || '').trim();
  const msg =
    'Backend / RSVP link address (Apps Script):\n' + (api || '(not set)') + '\n\n' +
    'App address (GitHub Pages):\n' + (pwa || '(not set yet - paste it into CONFIG.GITHUB_PAGES_APP_URL after you publish the site)') + '\n\n' +
    (pwa ? 'Email "Open App" buttons go to the GitHub Pages app.' : 'Until GITHUB_PAGES_APP_URL is set, email "Open App" buttons open a small landing page.');
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { console.log(msg); }
}

/** Menu helper: signs every phone out. People request a new emailed code the next time they open the app. */
function signOutAllDevices() {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('SESSION_EPOCH', String(Number(props.getProperty('SESSION_EPOCH') || '1') + 1));
  try { SpreadsheetApp.getUi().alert('Everyone has been signed out. They will sign in again with an emailed code.'); } catch (e) {}
}

function safeJsonForHtml_(obj) {
  return JSON.stringify(obj).replace(/</g, '\\u003c').replace(/>/g, '\\u003e').replace(/&/g, '\\u0026');
}

function escapeHtml_(v) {
  return String(v == null ? '' : v)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

/** MailApp options: friendly sender name and replies go to the Sunday School Secretary. */
function mailOptions_(opts, contacts) {
  const o = Object.assign({}, opts || {});
  o.name = CONFIG.MAIL_SENDER_NAME;
  const reply = contacts && String(contacts.secEmail || '').trim();
  if (reply && /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(reply)) o.replyTo = reply;
  return o;
}

/* ---------------- Signed RSVP links (emailed buttons) ---------------- */

function getRsvpSecret_() {
  const props = PropertiesService.getScriptProperties();
  let secret = props.getProperty('RSVP_SECRET');
  if (!secret) {
    secret = Utilities.getUuid() + Utilities.getUuid();
    props.setProperty('RSVP_SECRET', secret);
  }
  return secret;
}

function hmacWebSafe_(message) {
  const sig = Utilities.computeHmacSha256Signature(message, getRsvpSecret_());
  return Utilities.base64EncodeWebSafe(sig).replace(/=+$/g, '');
}

/** Opaque public id for an assignment (does not reveal the teacher's email). */
function rsvpId_(assignmentKey) {
  return hmacWebSafe_('id|' + assignmentKey).slice(0, 22);
}

function signRsvp_(id, action) {
  return hmacWebSafe_('sig|' + id + '|' + action).slice(0, 32);
}

function safeEqual_(a, b) {
  a = String(a || ''); b = String(b || '');
  if (a.length !== b.length) return false;
  let diff = 0;
  for (let i = 0; i < a.length; i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diff === 0;
}

function normalizeRsvpAction_(raw) {
  const a = String(raw || '').trim().toUpperCase();
  if (a === 'YES') return 'YES';
  if (a === 'MAYBE' || a === 'TENTATIVE') return 'MAYBE';
  if (a === 'NO' || a === 'SUB') return 'NO';
  return '';
}

function findAssignmentByRsvpId_(sheet, id) {
  if (sheet.getLastRow() < 2) return null;
  const vals = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
  for (let i = 0; i < vals.length; i++) {
    const room = String(vals[i][0] || '').trim();
    const teacher = String(vals[i][2] || '').trim();
    const rawDate = vals[i][4];
    const email = String(vals[i][6] || '').trim().toLowerCase();
    if (!isValidAssignment_(teacher, email, rawDate) || !room) continue;
    const dateKey = sheetDateKey_(rawDate, sheet);
    if (safeEqual_(rsvpId_(assignmentKey_(dateKey, room, email)), id)) {
      return { row: i + 2, room: room, teacher: teacher, dateKey: dateKey, lessonTitle: String(vals[i][5] || '').trim() };
    }
  }
  return null;
}

/**
 * Two-step handler for emailed RSVP buttons:
 * 1) Verify signature + find the assignment by id (never by row number).
 * 2) Show a confirmation page; only change the sheet when the person taps Confirm (confirm=1).
 * This keeps email link scanners from changing anyone's status.
 */
function handleSignedRsvp_(p) {
  const action = normalizeRsvpAction_(p.rsvp);
  const id = String(p.a || '');
  if (!action || !id || !safeEqual_(String(p.s || ''), signRsvp_(id, action))) {
    return renderRsvpMessagePage_('This link is not valid',
      'It may have been altered or copied incorrectly. Please open the Sunday School app to respond.');
  }

  const sheet = getScheduleSheet_();
  const found = sheet ? findAssignmentByRsvpId_(sheet, id) : null;
  if (!found) {
    return renderRsvpMessagePage_('This assignment has changed',
      'The schedule was updated after this email was sent. Please open the Sunday School app to see your current assignments.');
  }

  const daysUntil = dayOrdinalFromKey_(found.dateKey) - dayOrdinal_(new Date(), CONFIG.TIME_ZONE);
  if (daysUntil < -1) {
    return renderRsvpMessagePage_('This assignment has already passed', 'No response is needed. Thank you for your service!');
  }

  if (String(p.confirm || '') !== '1') {
    return renderRsvpConfirmPage_(action, found, p);
  }
  return handleEmailRsvpAction_(found.row, action);
}

function renderRsvpShell_(innerHtml) {
  const html = `<!DOCTYPE html><html><head><meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Valley View Sunday School — RSVP</title></head>
  <body style="margin:0;padding:32px 16px;background-color:#f0f2f5;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif;color:#1e242e;text-align:center;">
    <div style="max-width:420px;margin:32px auto;background:#ffffff;border-radius:20px;padding:28px 24px;box-shadow:0 8px 24px rgba(0,0,0,0.08);border:1px solid #dcdfe6;">
      <div style="font-size:13px;font-weight:800;color:#495085;text-transform:uppercase;letter-spacing:0.4px;margin-bottom:12px;">Valley View Sunday School</div>
      ${innerHtml}
    </div>
  </body></html>`;
  return HtmlService.createHtmlOutput(html)
    .setTitle('Valley View Sunday School — RSVP')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function renderRsvpMessagePage_(heading, body) {
  return renderRsvpShell_(
    `<div style="font-size:17px;font-weight:800;margin-bottom:10px;">${escapeHtml_(heading)}</div>
     <p style="font-size:13.5px;color:#4b5563;line-height:1.6;margin:0;">${escapeHtml_(body)}</p>`);
}

function renderRsvpConfirmPage_(action, found, p) {
  const labels = {
    YES: { title: 'Confirm: I can teach', btn: 'Yes, I’ll Teach', color: '#2e6f40' },
    MAYBE: { title: 'Confirm: Maybe', btn: 'Confirm Maybe', color: '#555888' },
    NO: { title: 'Confirm: I need a substitute', btn: 'Request a Sub', color: '#9E4242' }
  }[action];
  const base = getSafeWebAppUrl_();
  const sep = base.indexOf('?') === -1 ? '?' : '&';
  const confirmUrl = `${base}${sep}rsvp=${encodeURIComponent(action)}&a=${encodeURIComponent(p.a)}&s=${encodeURIComponent(p.s)}&confirm=1`;
  const formattedDate = Utilities.formatDate(localDateTime_(found.dateKey, 12, 0), CONFIG.TIME_ZONE, 'EEEE, MMMM d, yyyy');
  return renderRsvpShell_(
    `<div style="font-size:17px;font-weight:800;margin-bottom:12px;">${labels.title}</div>
     <div style="background:#F4F4F4;border-radius:12px;padding:12px 14px;text-align:left;font-size:12.5px;line-height:1.7;margin-bottom:16px;">
       <div><strong>Teacher:</strong> ${escapeHtml_(found.teacher)}</div>
       <div><strong>Date:</strong> ${escapeHtml_(formattedDate)}</div>
       <div><strong>Room:</strong> ${escapeHtml_(found.room)}</div>
       ${found.lessonTitle ? `<div><strong>Lesson:</strong> ${escapeHtml_(found.lessonTitle)}</div>` : ''}
     </div>
     <a href="${escapeHtml_(confirmUrl)}" target="_top" style="display:block;background-color:${labels.color};color:#ffffff;text-decoration:none;font-weight:700;font-size:14px;padding:12px 20px;border-radius:999px;">${labels.btn}</a>
     <p style="font-size:11.5px;color:#6b7280;margin:12px 0 0 0;">Nothing is recorded until you tap the button above.</p>`);
}

/* ---------------- Invite email (with retry flag) ---------------- */

function sendInviteEmail_(ctx) {
  const props = PropertiesService.getScriptProperties();
  const retryKey = CONFIG.INVITE_RETRY_PROPERTY_PREFIX + ctx.assignmentKey;
  try {
    const contacts = ctx.contacts || {};
    const subject = `Invitation: Sunday School Teaching Assignment on ${ctx.shortDate} (11:30 AM – 12:00 PM)`;
    const plainBody =
      `Hi ${ctx.teacher},\n\n` +
      'You have been invited to teach Sunday School in Valley View Ward:\n\n' +
      `• Date: ${ctx.formattedDate}\n` +
      '• Time: 11:30 AM – 12:00 PM (Utah time)\n' +
      `• Room / Side: ${ctx.room}\n` +
      `• Conductor: ${ctx.conductor || 'N/A'}\n` +
      `• Lesson: ${ctx.lessonTitle}\n` +
      (ctx.lessonUrl ? `• Lesson Link: ${ctx.lessonUrl}\n\n` : '\n') +
      (ctx.webAppUrl ? `Please RSVP or manage your assignment in the Sunday School App (${ctx.webAppUrl}) or contact ` : 'Please RSVP using the calendar invitation or contact ') +
      `${contacts.secName} (${contacts.secPhone}, ${contacts.secEmail}).\n\n` +
      'Thank you for your service!\n' +
      '-The Sunday School Presidency';

    const htmlBody = buildStyledReminderEmailHtml_({
      isInvite: true,
      rowNumber: ctx.rowNumber,
      assignmentKey: ctx.assignmentKey,
      dateKey: ctx.dateKey,
      teacher: ctx.teacher,
      badgeText: 'Calendar Invitation',
      formattedDate: ctx.formattedDate,
      room: ctx.room,
      conductor: ctx.conductor || 'N/A',
      lessonRaw: ctx.lessonTitle,
      lessonUrl: ctx.lessonUrl,
      secName: contacts.secName,
      secPhone: contacts.secPhone,
      secEmail: contacts.secEmail,
      aidsRaw: ctx.aidsRaw
    });

    MailApp.sendEmail(ctx.teacherEmail, subject, plainBody, mailOptions_({ htmlBody: htmlBody }, contacts));
    props.deleteProperty(retryKey);
    return true;
  } catch (err) {
    console.error(`Styled calendar invite email failed for ${ctx.teacher} (${ctx.teacherEmail}): ${err}`);
    try { props.setProperty(retryKey, new Date().toISOString()); } catch (e) {}
    return false;
  }
}

/** If an invite email failed earlier, try again on the next sync. */
function retryPendingInvite_(ctx) {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty(CONFIG.INVITE_RETRY_PROPERTY_PREFIX + ctx.assignmentKey)) return;
  sendInviteEmail_(ctx);
}

/**
 * Dynamically fetches the secretary's contact info and alert emails
 * directly from the "Contact" tab based on callings in Column A.
 */
function getContactsConfig_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet ? spreadsheet.getSheetByName(CONFIG.CONTACT_SHEET_NAME) : null;
  if (!sheet) {
    console.error(`Required sheet not found: ${CONFIG.CONTACT_SHEET_NAME}.`);
    return null;
  }

  const lastContactRow = sheet.getLastRow();
  if (lastContactRow < 2) return { secName: '', secPhone: '', secEmail: '', alertEmails: '' };
  const values = sheet.getRange(2, 1, lastContactRow - 1, 4).getValues();
  let secName = '', secPhone = '', secEmail = '';
  const alertEmailList = [];

  for (const row of values) {
    const calling = String(row[0] || '').toLowerCase();
    const name = String(row[1] || '').trim();
    const phone = String(row[2] || '').trim();
    const email = String(row[3] || '').trim();

    const isSundaySchool = calling.includes('sunday school');

    if (isSundaySchool && calling.includes('secretary')) {
      secName = name;
      secPhone = phone;
      secEmail = email;
    }

    if (isSundaySchool && (calling.includes('president') || calling.includes('counselor') || calling.includes('secretary')) && email) {
      alertEmailList.push(email);
    }
  }

  return {
    secName: secName,
    secPhone: secPhone,
    secEmail: secEmail,
    alertEmails: alertEmailList.join(', ')
  };
}

/**
 * Deletes stored reminder/alert flags for assignments whose date has passed,
 * so Script Properties never fill up. Called at the end of the daily sync.
 */
function cleanupOldProperties_() {
  const props = PropertiesService.getScriptProperties();
  const todayKey = Utilities.formatDate(new Date(), CONFIG.TIME_ZONE, 'yyyy-MM-dd');
  const prefixes = [CONFIG.REMINDER_PROPERTY_PREFIX, CONFIG.SUBSTITUTE_ALERT_PROPERTY_PREFIX, CONFIG.INVITE_RETRY_PROPERTY_PREFIX];
  let removed = 0;

  Object.keys(props.getProperties()).forEach(key => {
    const prefix = prefixes.find(p => key.startsWith(p));
    if (!prefix) return;
    let encoded = key.slice(prefix.length).split(':')[0];
    while (encoded.length % 4) encoded += '=';
    try {
      const raw = Utilities.newBlob(Utilities.base64DecodeWebSafe(encoded)).getDataAsString();
      if (raw.split('|')[0] < todayKey) {
        props.deleteProperty(key);
        removed++;
      }
    } catch (e) {
      // Ignore keys that can't be parsed.
    }
  });

  if (removed) console.log(`Cleaned up ${removed} expired reminder/alert flag(s).`);
}

function getScheduleSheet_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet ? spreadsheet.getSheetByName(CONFIG.SHEET_NAME) : null;
  if (!sheet) console.error(`Required sheet not found: ${CONFIG.SHEET_NAME}`);
  return sheet;
}

function getCalendar_() {
  const calendar = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
  if (!calendar) console.error('Could not access the configured calendar. Check its ID and sharing permissions.');
  return calendar;
}

function checkCalendarTimeZone_(calendar) {
  const actual = calendar.getTimeZone();
  if (actual !== CONFIG.TIME_ZONE) {
    console.error(`TIME ZONE MISMATCH: Calendar is set to ${actual}; script expects ${CONFIG.TIME_ZONE} (Utah time). Update the Google Calendar timezone or CONFIG.TIME_ZONE before running automation.`);
    return false;
  }
  return true;
}

function isValidAssignment_(teacher, email, date) {
  return Boolean(teacher) && teacher.toUpperCase() !== 'N/A' &&
    Boolean(email) && date instanceof Date && !isNaN(date.getTime());
}

function sheetDateKey_(date, sheet) {
  return Utilities.formatDate(date, sheet.getParent().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
}

function dayOrdinal_(date, timeZone) {
  return dayOrdinalFromKey_(Utilities.formatDate(date, timeZone, 'yyyy-MM-dd'));
}

function dayOrdinalFromKey_(dateKey) {
  const parts = dateKey.split('-').map(Number);
  return Math.floor(Date.UTC(parts[0], parts[1] - 1, parts[2]) / 86400000);
}

function localDateTime_(dateKey, hour, minute) {
  const parts = dateKey.split('-').map(Number);
  const targetAsUtc = Date.UTC(parts[0], parts[1] - 1, parts[2], hour, minute, 0, 0);
  let instant = targetAsUtc;
  for (let i = 0; i < 3; i++) {
    const local = Utilities.formatDate(new Date(instant), CONFIG.TIME_ZONE, "yyyy-MM-dd'T'HH:mm:ss");
    const localParts = local.match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})$/);
    if (!localParts) throw new Error(`Could not convert local date/time: ${dateKey} ${hour}:${minute}`);
    const localAsUtc = Date.UTC(
      Number(localParts[1]), Number(localParts[2]) - 1, Number(localParts[3]),
      Number(localParts[4]), Number(localParts[5]), Number(localParts[6])
    );
    const offset = localAsUtc - instant;
    const corrected = targetAsUtc - offset;
    if (corrected === instant) break;
    instant = corrected;
  }
  return new Date(instant);
}

function assignmentKey_(dateKey, room, email) {
  const raw = `${dateKey}|${String(room).trim().toLowerCase()}|${String(email).trim().toLowerCase()}`;
  return Utilities.base64EncodeWebSafe(raw).replace(/=+$/g, '');
}

function guestIsInvited_(event, email) {
  const normalized = String(email).trim().toLowerCase();
  return event.getGuestList().some(guest => guest.getEmail().trim().toLowerCase() === normalized);
}

function guestStatusText_(status) {
  if (status === CalendarApp.GuestStatus.YES) return 'ACCEPTED';
  if (status === CalendarApp.GuestStatus.NO) return 'SUB';
  if (status === CalendarApp.GuestStatus.MAYBE) return 'TENTATIVE';
  return 'PENDING';
}

function getMatchingEventStatus_(calendar, dateKey, teacher, room, email) {
  const dayStart = localDateTime_(dateKey, 0, 0);
  const dayEnd = localDateTime_(dateKey, 23, 59);
  const expectedTitle = `Sunday School Teaching: ${teacher} (${room})`;
  for (const event of calendar.getEvents(dayStart, dayEnd)) {
    if (event.getTitle() !== expectedTitle) continue;
    const guest = event.getGuestList().find(g => g.getEmail().trim().toLowerCase() === email);
    if (guest) return guestStatusText_(guest.getGuestStatus());
  }
  return null;
}

function reminderTarget_(teachingDateKey, dayOffset, hour, minute) {
  const parts = teachingDateKey.split('-').map(Number);
  const shifted = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2] + dayOffset));
  const shiftedKey = Utilities.formatDate(shifted, 'UTC', 'yyyy-MM-dd');
  return localDateTime_(shiftedKey, hour, minute);
}

/* ============================================================================
 * APP API (used by the installable app on GitHub Pages)
 *  - Every call is a POST with a text/plain JSON body (avoids browser preflight).
 *  - Sign-in = emailed 6-digit code, sent ONLY to addresses already on the schedule or in a Sunday School calling.
 *  - Sessions are signed tokens. Removing someone from the sheet locks them out within ~2 minutes.
 *  - Teachers can change only their own assignments; the Sunday School presidency can change any.
 *  - Only the public attendance CHECK-IN address (no admin key) is returned, and only to signed-in people.
 *    This project never reads or writes attendance data.
 * ========================================================================== */

/**
 * The attendance app's PUBLIC check-in address, with any query string (such as an admin key) removed.
 * This is the only attendance address this project ever hands out.
 */
function getAttendanceCheckInBaseUrl_() {
  const raw = String(CONFIG.ATTENDANCE_WEB_APP_URL || '').trim();
  if (!/^https:\/\/script\.google\.com\/macros\/s\/[A-Za-z0-9_-]+\/exec/.test(raw)) return '';
  return raw.split('?')[0].split('#')[0];
}

function getAppUrl_() {
  return String(CONFIG.GITHUB_PAGES_APP_URL || '').trim() || getSafeWebAppUrl_();
}

function renderLandingPage_() {
  const pwa = String(CONFIG.GITHUB_PAGES_APP_URL || '').trim();
  const inner = pwa
    ? `<div style="font-size:17px;font-weight:800;margin-bottom:10px;">Open the Sunday School app</div>
       <p style="font-size:13.5px;color:#4b5563;line-height:1.6;margin:0 0 16px 0;">The app has its own address. Open it, sign in with your email, and add it to your home screen.</p>
       <a href="${escapeHtml_(pwa)}" target="_top" style="display:block;background-color:#9e4340;color:#ffffff;text-decoration:none;font-weight:700;font-size:14px;padding:12px 20px;border-radius:999px;">Open App</a>`
    : `<div style="font-size:17px;font-weight:800;margin-bottom:10px;">Sunday School</div>
       <p style="font-size:13.5px;color:#4b5563;line-height:1.6;margin:0;">This address powers the Sunday School app. Please use the app link you were given.</p>`;
  return renderRsvpShell_(inner);
}

function authError_() { const e = new Error('auth'); e.name = 'AuthError'; return e; }
function clientError_(msg) { const e = new Error(msg); e.client = true; return e; }

function doPost(e) {
  let out;
  try {
    const raw = (e && e.postData && e.postData.contents) ? e.postData.contents : '{}';
    out = routeApi_(JSON.parse(raw));
  } catch (err) {
    if (err && err.name === 'AuthError') {
      out = { ok: false, code: 'auth', message: 'Please sign in again.' };
    } else {
      const msg = String((err && err.message) || '');
      console.error('API error: ' + msg);
      const safe = err && (err.client || /^(Invalid schedule row|That row is not|The schedule changed|Invalid RSVP|Please |Teaching Aids sheet|Invalid teaching aid)/.test(msg));
      out = { ok: false, code: 'error', message: safe ? msg : 'Something went wrong. Please try again.' };
    }
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function routeApi_(req) {
  switch (String(req.action || '')) {
    case 'requestCode': return apiRequestCode_(req);
    case 'verifyCode': return apiVerifyCode_(req);
    case 'boot': return apiBoot_(req);
    case 'rsvp': return apiRsvp_(req);
    case 'requestSub': return apiRequestSub_(req);
    case 'addTeachingAid': return apiAddTeachingAid_(req);
    case 'linkPreview': return apiLinkPreview_(req);
    default: return { ok: false, code: 'bad_request', message: 'Unknown request.' };
  }
}

/** People allowed in: anyone with an email on the Schedule sheet, plus Sunday School callings on the Contact sheet. */
function getAllowedUsers_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('allowed_users_v1');
  if (hit) { try { return JSON.parse(hit); } catch (e) {} }

  const users = {};
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  const sched = getScheduleSheet_();
  if (sched && sched.getLastRow() >= 2) {
    sched.getRange(2, 3, sched.getLastRow() - 1, 5).getValues().forEach(function(r) { // C..G
      const name = String(r[0] || '').trim();
      const email = String(r[4] || '').trim().toLowerCase();
      if (name && name.toUpperCase() !== 'N/A' && /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email) && !users[email]) {
        users[email] = { email: email, name: name, role: 'teacher' };
      }
    });
  }

  const contacts = spreadsheet ? spreadsheet.getSheetByName(CONFIG.CONTACT_SHEET_NAME) : null;
  if (contacts && contacts.getLastRow() >= 2) {
    contacts.getRange(2, 1, contacts.getLastRow() - 1, 4).getValues().forEach(function(r) {
      const calling = String(r[0] || '').toLowerCase();
      const name = String(r[1] || '').trim();
      const email = String(r[3] || '').trim().toLowerCase();
      if (!email || !calling.includes('sunday school')) return;
      if (!(calling.includes('president') || calling.includes('counselor') || calling.includes('secretary'))) return;
      users[email] = { email: email, name: (users[email] && users[email].name) || name, role: 'presidency' };
    });
  }

  try { cache.put('allowed_users_v1', JSON.stringify(users), 120); } catch (e) {}
  return users;
}

function emailHash_(email) { return hmacWebSafe_('em|' + email).slice(0, 24); }

function apiRequestCode_(req) {
  const email = String(req.email || '').trim().toLowerCase();
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) {
    return { ok: false, code: 'bad_email', message: 'Please enter a valid email address.' };
  }
  const cache = CacheService.getScriptCache();
  const eh = emailHash_(email);
  const rlKey = 'rl:' + eh;
  const used = Number(cache.get(rlKey) || 0);
  if (used >= 5) {
    return { ok: false, code: 'rate', message: 'Too many codes requested. Please wait an hour and try again.' };
  }
  cache.put(rlKey, String(used + 1), 3600);

  const user = getAllowedUsers_()[email];
  if (user) {
    const num = parseInt(Utilities.getUuid().replace(/-/g, '').slice(0, 8), 16) % 1000000;
    const code = String(num).padStart(6, '0');
    cache.put('code:' + eh, JSON.stringify({ h: hmacWebSafe_('code|' + email + '|' + code), t: 0 }), CONFIG.LOGIN_CODE_TTL_SEC);
    try {
      let contacts = null; try { contacts = getContactsConfig_(); } catch (e) {}
      MailApp.sendEmail(
        email,
        'Your Sunday School app code: ' + code,
        'Hi ' + user.name + ',\n\nYour sign-in code is ' + code + '\n\nIt expires in ' + Math.round(CONFIG.LOGIN_CODE_TTL_SEC / 60) +
        ' minutes. If you did not ask for this, you can ignore this email.\n\n- Valley View Sunday School',
        mailOptions_({}, contacts)
      );
    } catch (err) {
      console.error('Sign-in code email failed: ' + err);
      return { ok: false, code: 'mail', message: 'We could not send the email right now. Please try again in a few minutes.' };
    }
  }
  return { ok: true, message: 'If that address is on the teaching schedule, a code is on its way.' };
}

function signSession_(email) {
  const props = PropertiesService.getScriptProperties();
  const exp = Math.floor(Date.now() / 1000) + CONFIG.SESSION_DAYS * 86400;
  const payload = Utilities.base64EncodeWebSafe(email + '|' + exp + '|' + (props.getProperty('SESSION_EPOCH') || '1'));
  return payload + '.' + hmacWebSafe_('session|' + payload).slice(0, 32);
}

function requireUser_(req) {
  const parts = String(req.token || '').split('.');
  if (parts.length !== 2) throw authError_();
  if (!safeEqual_(parts[1], hmacWebSafe_('session|' + parts[0]).slice(0, 32))) throw authError_();
  let decoded;
  try { decoded = Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString(); } catch (e) { throw authError_(); }
  const f = decoded.split('|');
  const epoch = PropertiesService.getScriptProperties().getProperty('SESSION_EPOCH') || '1';
  if (f.length !== 3 || Number(f[1]) < Date.now() / 1000 || f[2] !== epoch) throw authError_();
  const user = getAllowedUsers_()[f[0]];
  if (!user) throw authError_();
  return user;
}

function apiVerifyCode_(req) {
  const email = String(req.email || '').trim().toLowerCase();
  const code = String(req.code || '').replace(/\D/g, '');
  const cache = CacheService.getScriptCache();
  const key = 'code:' + emailHash_(email);
  const raw = cache.get(key);
  if (!raw || code.length !== 6) {
    return { ok: false, code: 'bad_code', message: 'That code has expired or is not right. Request a new one.' };
  }
  const rec = JSON.parse(raw);
  if (rec.t >= 5) {
    cache.remove(key);
    return { ok: false, code: 'bad_code', message: 'Too many tries. Please request a new code.' };
  }
  if (!safeEqual_(rec.h, hmacWebSafe_('code|' + email + '|' + code))) {
    rec.t++;
    cache.put(key, JSON.stringify(rec), CONFIG.LOGIN_CODE_TTL_SEC);
    return { ok: false, code: 'bad_code', message: 'That code is not right. Please check it and try again.' };
  }
  cache.remove(key);
  const user = getAllowedUsers_()[email];
  if (!user) return { ok: false, code: 'bad_code', message: 'That email is no longer on the schedule.' };
  return { ok: true, token: signSession_(email), me: { email: user.email, name: user.name, role: user.role } };
}

function apiBoot_(req) {
  const user = requireUser_(req);
  const boot = getWebAppBootData(req.refresh === true);
  boot.attendanceBaseUrl = getAttendanceCheckInBaseUrl_(); // always the public check-in address, never an admin link
  return { ok: true, me: { email: user.email, name: user.name, role: user.role }, boot: boot };
}

function apiAuthorizedRow_(req, user) {
  const sheet = getScheduleSheet_();
  if (!sheet) throw clientError_('Schedule sheet not found.');
  const a = validateAssignmentRow_(sheet, req.row, req.dateKey, req.room);
  if (user.role !== 'presidency' && a.email !== user.email) {
    throw clientError_('You can only change your own assignments.');
  }
  return a;
}

function apiRsvp_(req) {
  const user = requireUser_(req);
  const a = apiAuthorizedRow_(req, user);
  const res = updateTeacherRsvpFromApp(a.row, String(req.status || '').toUpperCase(), a.dateKey, a.room);
  return { ok: true, status: res.status };
}

function apiRequestSub_(req) {
  const user = requireUser_(req);
  const a = apiAuthorizedRow_(req, user);
  requestSubFromApp(a.row, a.dateKey, a.room);
  return { ok: true, status: 'SUB' };
}

function apiAddTeachingAid_(req) {
  const user = requireUser_(req);
  req.sharedBy = user.name; // never trust a name sent by the browser
  const res = addTeachingAidFromApp(req);
  return Object.assign({ ok: true }, res);
}

function apiLinkPreview_(req) {
  requireUser_(req);
  return { ok: true, imageUrl: getLinkPreviewImageFromApp(String(req.url || '')) };
}
