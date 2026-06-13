// ============================================================
// SSB MENTORSHIP BOOKING PORTAL — Code.gs
// Google Sheet ID: 1zEUDXBmlj1L27VLXLcjUdg9SStLsIx1QsD6zvyErtUA
// ============================================================

var SHEET_ID = '1zEUDXBmlj1L27VLXLcjUdg9SStLsIx1QsD6zvyErtUA';
var SHEET_MENTORS = 'Mentor Details';
var SHEET_SESSIONS = 'Master Session Details';
var SHEET_STUDENTS = 'Student Data';
var APP_TITLE = 'SSB Mentorship Booking Portal';

// ============================================================
// WEB APP ENTRY POINT
// ============================================================

function doGet(e) {
  if (e && e.parameter && e.parameter.token) {
    return handleMentorFeedbackPage(e.parameter.token);
  }
  var template = HtmlService.createTemplateFromFile('index');
  var html = template.evaluate()
    .setTitle(APP_TITLE)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  return html;
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================================
// MENTOR FEEDBACK PAGE (Token-based, no auth required)
// ============================================================

function handleMentorFeedbackPage(token) {
  var sessionData = getSessionByToken(token);
  if (!sessionData) {
    return HtmlService.createHtmlOutput(
      '<html><body style="font-family:sans-serif;text-align:center;padding:60px;">' +
      '<h2 style="color:#e53e3e;">Invalid or Expired Link</h2>' +
      '<p>This feedback link is invalid or has already been used.</p>' +
      '</body></html>'
    );
  }
  var template = HtmlService.createTemplateFromFile('mentor_feedback');
  template.sessionData = JSON.stringify(sessionData);
  template.token = token;
  return template.evaluate()
    .setTitle('Mentor Feedback — ' + APP_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getSessionByToken(token) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(SHEET_SESSIONS);
    var data = sheet.getDataRange().getValues();
    var decoded = Utilities.newBlob(Utilities.base64Decode(token)).getDataAsString();
    var parts = decoded.split('|');
    if (parts.length < 3) return null;
    var rowIndex = parseInt(parts[0]);
    var studentEmail = parts[1];
    var mentorName = parts[2];
    if (rowIndex < 1 || rowIndex >= data.length) return null;
    var row = data[rowIndex];
    if (row[0] !== studentEmail || row[6] !== mentorName) return null;
    if (row[11] && row[12]) return null; // already submitted
    return {
      rowIndex: rowIndex,
      studentEmail: row[0],
      studentName: row[1],
      rollNo: row[2],
      batch: row[3],
      msnSession: row[5],
      mentorName: row[6],
      dateTime: row[7] ? Utilities.formatDate(new Date(row[7]), Session.getScriptTimeZone(), 'dd MMM yyyy, hh:mm a') : '',
      token: token
    };
  } catch (e) {
    Logger.log('getSessionByToken error: ' + e.toString());
    return null;
  }
}

function submitMentorFeedback(token, feedback, rating) {
  try {
    var sessionData = getSessionByToken(token);
    if (!sessionData) return { success: false, message: 'Invalid or expired token.' };
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(SHEET_SESSIONS);
    var sheetRow = sessionData.rowIndex + 1; // 1-indexed
    sheet.getRange(sheetRow, 12).setValue(feedback);
    sheet.getRange(sheetRow, 13).setValue(rating);
    return { success: true, message: 'Thank you! Your feedback has been recorded.' };
  } catch (e) {
    Logger.log('submitMentorFeedback error: ' + e.toString());
    return { success: false, message: 'Error submitting feedback: ' + e.toString() };
  }
}

// ============================================================
// TEST HELPER — Run this manually to get a feedback URL for testing
// ============================================================

function testMentorFeedbackLink() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_SESSIONS);
  var data = sheet.getDataRange().getValues();

  // Loop through all data rows and log a feedback URL for each
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue; // skip empty rows
    var studentEmail = row[0];
    var mentorName = row[6];
    if (!mentorName) continue;
    var token = generateFeedbackToken(i, studentEmail, mentorName);
    var url = ScriptApp.getService().getUrl() + '?token=' + token;
    Logger.log('Row ' + i + ' | Session: ' + row[5] + ' | Student: ' + row[1] + ' | Mentor: ' + mentorName);
    Logger.log('Feedback URL: ' + url);
    Logger.log('---');
  }
}

// ============================================================
// AUTHENTICATION & STUDENT LOOKUP
// ============================================================

function getStudentInfo() {
  try {
    var email = Session.getActiveUser().getEmail();
    if (!email) return { success: false, reason: 'no_email' };
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(SHEET_STUDENTS);
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().toLowerCase().trim() === email.toLowerCase().trim()) {
        return {
          success: true,
          email: data[i][0],
          name: data[i][1],
          rollNo: data[i][2],
          batch: data[i][3]
        };
      }
    }
    return { success: false, reason: 'not_found', email: email };
  } catch (e) {
    Logger.log('getStudentInfo error: ' + e.toString());
    return { success: false, reason: 'error', message: e.toString() };
  }
}

// ============================================================
// MENTOR DATA
// ============================================================

function getMentors() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(SHEET_MENTORS);
    var data = sheet.getDataRange().getValues();
    var mentors = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue;
      mentors.push({
        id: row[0],
        name: row[1],
        email: row[2],
        linkedin: row[3],
        university: row[4],
        educationStream: row[5],
        currentRole: row[6],
        currentOrg: row[7],
        prevOrgRole: row[8],
        prevCompanies: row[9],
        prevRoles: row[10],
        skills: row[11],
        photoLink: row[12]
        // schedulistaLink (col 13) intentionally unused
      });
    }
    return { success: true, mentors: mentors };
  } catch (e) {
    Logger.log('getMentors error: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

// ============================================================
// SESSION DATA
// ============================================================

function getStudentSessions(studentEmail) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(SHEET_SESSIONS);
    var data = sheet.getDataRange().getValues();
    var sessions = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue;
      if (row[0].toString().toLowerCase().trim() !== studentEmail.toLowerCase().trim()) continue;
      sessions.push({
        rowIndex: i,
        studentEmail: row[0],
        fullName: row[1],
        rollNo: row[2],
        batch: row[3],
        meetLink: row[4],
        msnSession: row[5],
        mentor: row[6],
        dateTime: row[7] ? new Date(row[7]).toISOString() : '',
        recording: row[8],
        studentFeedback: row[9],
        studentRating: row[10],
        mentorFeedback: row[11],
        mentorRating: row[12]
      });
    }
    sessions.sort(function(a, b) { return new Date(b.dateTime) - new Date(a.dateTime); });
    return { success: true, sessions: sessions };
  } catch (e) {
    Logger.log('getStudentSessions error: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

function getDashboardStats(studentEmail) {
  try {
    var result = getStudentSessions(studentEmail);
    if (!result.success) return result;
    var sessions = result.sessions;
    var now = new Date();
    var total = sessions.length;
    var completed = 0;
    var pendingFeedback = 0;
    for (var i = 0; i < sessions.length; i++) {
      var s = sessions[i];
      var sessionDate = s.dateTime ? new Date(s.dateTime) : null;
      if (sessionDate && sessionDate < now) {
        completed++;
        if (!s.studentFeedback || !s.studentRating) pendingFeedback++;
      }
    }
    return { success: true, total: total, completed: completed, pendingFeedback: pendingFeedback, sessions: sessions };
  } catch (e) {
    Logger.log('getDashboardStats error: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

// ============================================================
// FEEDBACK VALIDATION
// ============================================================

function checkFeedbackRequired(studentEmail) {
  try {
    var result = getStudentSessions(studentEmail);
    if (!result.success) return result;
    var now = new Date();
    var pendingSessions = [];
    for (var i = 0; i < result.sessions.length; i++) {
      var s = result.sessions[i];
      var sessionDate = s.dateTime ? new Date(s.dateTime) : null;
      if (sessionDate && sessionDate < now) {
        if (!s.studentFeedback || s.studentFeedback.toString().trim() === '' ||
            !s.studentRating || s.studentRating.toString().trim() === '') {
          pendingSessions.push(s);
        }
      }
    }
    return { success: true, feedbackRequired: pendingSessions.length > 0, pendingSessions: pendingSessions };
  } catch (e) {
    Logger.log('checkFeedbackRequired error: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

function submitStudentFeedback(studentEmail, rowIndex, feedback, rating) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(SHEET_SESSIONS);
    var data = sheet.getDataRange().getValues();
    var idx = parseInt(rowIndex);
    if (data[idx][0].toString().toLowerCase().trim() !== studentEmail.toLowerCase().trim()) {
      return { success: false, message: 'Unauthorized.' };
    }
    sheet.getRange(idx + 1, 10).setValue(feedback); // MSN-StudentFeedback
    sheet.getRange(idx + 1, 11).setValue(rating);   // MSN-StudentRating
    return { success: true, message: 'Feedback submitted successfully!' };
  } catch (e) {
    Logger.log('submitStudentFeedback error: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

// ============================================================
// BOOKING FLOW
// ============================================================

function bookSession(studentEmail, studentName, rollNo, batch, mentorId, mentorName, mentorEmail, dateTimeStr) {
  try {
    var sessionDateTime = new Date(dateTimeStr);
    var now = new Date();

    if (sessionDateTime <= now) {
      return { success: false, message: 'Cannot book a session in the past.' };
    }

    var overlapCheck = checkOverlapping(studentEmail, mentorName, sessionDateTime);
    if (!overlapCheck.success) return overlapCheck;
    if (overlapCheck.hasOverlap) return { success: false, message: overlapCheck.message };

    var calResult = createCalendarEvent(studentEmail, studentName, mentorEmail, mentorName, sessionDateTime);
    if (!calResult.success) return calResult;

    var sessionNumber = getNextSessionNumber(studentEmail);

    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(SHEET_SESSIONS);

    sheet.appendRow([
      studentEmail, studentName, rollNo, batch,
      calResult.meetLink, sessionNumber, mentorName,
      sessionDateTime,
      '', '', '', '', ''
    ]);

    var newRowIndex = sheet.getLastRow() - 1; // 0-indexed

    sendStudentConfirmationEmail(studentEmail, studentName, mentorName, sessionDateTime, calResult.meetLink);
    sendMentorNotificationEmail(mentorEmail, mentorName, studentName, studentEmail, rollNo, batch, sessionDateTime, calResult.meetLink);

    // Schedule mentor feedback email 1 hour after session ends
    scheduleMentorFeedbackTrigger(newRowIndex, sessionDateTime);

    return {
      success: true,
      message: 'Session booked successfully!',
      sessionNumber: sessionNumber,
      meetLink: calResult.meetLink,
      dateTime: sessionDateTime.toISOString()
    };
  } catch (e) {
    Logger.log('bookSession error: ' + e.toString());
    return { success: false, message: 'Booking failed: ' + e.toString() };
  }
}

function checkOverlapping(studentEmail, mentorName, sessionDateTime) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(SHEET_SESSIONS);
    var data = sheet.getDataRange().getValues();
    var sessionEnd = new Date(sessionDateTime.getTime() + 60 * 60 * 1000);

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[7]) continue;
      var existingStart = new Date(row[7]);
      var existingEnd = new Date(existingStart.getTime() + 60 * 60 * 1000);
      var overlaps = sessionDateTime < existingEnd && sessionEnd > existingStart;

      if (overlaps && row[0].toString().toLowerCase().trim() === studentEmail.toLowerCase().trim()) {
        return { success: true, hasOverlap: true, message: 'You already have a session overlapping this time slot.' };
      }
      if (overlaps && row[6] && row[6].toString().trim() === mentorName.trim()) {
        return { success: true, hasOverlap: true, message: 'This mentor is not available at the selected time.' };
      }
    }
    return { success: true, hasOverlap: false };
  } catch (e) {
    Logger.log('checkOverlapping error: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

function getNextSessionNumber(studentEmail) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_SESSIONS);
  var data = sheet.getDataRange().getValues();
  var count = 0;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().toLowerCase().trim() === studentEmail.toLowerCase().trim()) count++;
  }
  return 'MSN ' + (count + 1);
}

// ============================================================
// CALENDAR & MEET
// ============================================================

function createCalendarEvent(studentEmail, studentName, mentorEmail, mentorName, sessionDateTime) {
  try {
    var endTime = new Date(sessionDateTime.getTime() + 60 * 60 * 1000);
    var requestId = Utilities.getUuid(); // unique per request — required by Google

    // Use Advanced Calendar API directly to create event WITH Meet link in one call
    var eventResource = {
      summary: 'SSB Mentorship Session - ' + studentName + ' with ' + mentorName,
      description: 'SSB Mentorship Portal session.\n\nStudent: ' + studentName + ' (' + studentEmail + ')\nMentor: ' + mentorName + ' (' + mentorEmail + ')',
      start: {
        dateTime: sessionDateTime.toISOString(),
        timeZone: Session.getScriptTimeZone()
      },
      end: {
        dateTime: endTime.toISOString(),
        timeZone: Session.getScriptTimeZone()
      },
      attendees: [
        { email: studentEmail },
        { email: mentorEmail }
      ],
      conferenceData: {
        createRequest: {
          requestId: requestId,
          conferenceSolutionKey: { type: 'hangoutsMeet' }
        }
      },
      reminders: {
        useDefault: false,
        overrides: [{ method: 'popup', minutes: 60 }]
      }
    };

    // conferenceDataVersion=1 is required to generate Meet link
    var createdEvent = Calendar.Events.insert(eventResource, 'primary', {
      conferenceDataVersion: 1,
      sendUpdates: 'all'
    });

    var meetLink = '';
    if (createdEvent.conferenceData && createdEvent.conferenceData.entryPoints) {
      for (var i = 0; i < createdEvent.conferenceData.entryPoints.length; i++) {
        if (createdEvent.conferenceData.entryPoints[i].entryPointType === 'video') {
          meetLink = createdEvent.conferenceData.entryPoints[i].uri;
          break;
        }
      }
    }

    // Fallback only if API didn't return a Meet link
    if (!meetLink) {
      Logger.log('Meet link not returned by API, using fallback');
      meetLink = generateMeetLink();
    }

    return { success: true, eventId: createdEvent.id, meetLink: meetLink };
  } catch (e) {
    Logger.log('createCalendarEvent error: ' + e.toString());
    return { success: false, message: 'Calendar event creation failed: ' + e.toString() };
  }
}

function generateMeetLink() {
  var chars = 'abcdefghijklmnopqrstuvwxyz';
  function rand(n) {
    var s = '';
    for (var i = 0; i < n; i++) s += chars[Math.floor(Math.random() * chars.length)];
    return s;
  }
  return 'https://meet.google.com/' + rand(3) + '-' + rand(4) + '-' + rand(3);
}

// ============================================================
// EMAIL NOTIFICATIONS
// ============================================================

function sendStudentConfirmationEmail(studentEmail, studentName, mentorName, sessionDateTime, meetLink) {
  try {
    var formattedDate = Utilities.formatDate(sessionDateTime, Session.getScriptTimeZone(), 'EEEE, dd MMMM yyyy');
    var formattedTime = Utilities.formatDate(sessionDateTime, Session.getScriptTimeZone(), 'hh:mm a z');
    var subject = '✅ Mentorship Session Confirmed — ' + mentorName;
    var body = '<!DOCTYPE html><html><body style="font-family:\'Segoe UI\',Arial,sans-serif;max-width:600px;margin:0 auto;background:#f8f9fa;padding:20px;">' +
      '<div style="background:linear-gradient(135deg,#1a1a2e 0%,#16213e 100%);padding:30px;border-radius:12px 12px 0 0;text-align:center;">' +
      '<h1 style="color:#00d4aa;margin:0;font-size:24px;">SSB Mentorship Portal</h1>' +
      '<p style="color:#a0aec0;margin:8px 0 0;">Scaler School of Business</p></div>' +
      '<div style="background:white;padding:32px;border-radius:0 0 12px 12px;box-shadow:0 4px 20px rgba(0,0,0,0.08);">' +
      '<h2 style="color:#1a1a2e;margin-top:0;">Your session is confirmed! 🎉</h2>' +
      '<p style="color:#4a5568;">Hi <strong>' + studentName + '</strong>,</p>' +
      '<p style="color:#4a5568;">Your mentorship session has been successfully booked.</p>' +
      '<div style="background:#f0fdf9;border-left:4px solid #00d4aa;padding:20px;border-radius:8px;margin:24px 0;">' +
      '<table style="width:100%;border-collapse:collapse;">' +
      '<tr><td style="padding:8px 0;color:#718096;width:140px;">Mentor</td><td style="padding:8px 0;font-weight:600;color:#1a1a2e;">' + mentorName + '</td></tr>' +
      '<tr><td style="padding:8px 0;color:#718096;">Date</td><td style="padding:8px 0;font-weight:600;color:#1a1a2e;">' + formattedDate + '</td></tr>' +
      '<tr><td style="padding:8px 0;color:#718096;">Time</td><td style="padding:8px 0;font-weight:600;color:#1a1a2e;">' + formattedTime + '</td></tr>' +
      '<tr><td style="padding:8px 0;color:#718096;">Duration</td><td style="padding:8px 0;font-weight:600;color:#1a1a2e;">60 Minutes</td></tr>' +
      '</table></div>' +
      '<div style="text-align:center;margin:24px 0;">' +
      '<a href="' + meetLink + '" style="background:#00d4aa;color:white;padding:14px 32px;border-radius:8px;text-decoration:none;font-weight:600;font-size:16px;display:inline-block;">Join Google Meet</a></div>' +
      '<p style="color:#718096;font-size:14px;">A calendar invite has been sent to your email. Please be ready 5 minutes before the session.</p>' +
      '<hr style="border:none;border-top:1px solid #e2e8f0;margin:24px 0;">' +
      '<p style="color:#a0aec0;font-size:12px;text-align:center;">SSB Mentorship Portal · Scaler School of Business · Bengaluru</p>' +
      '</div></body></html>';
    GmailApp.sendEmail(studentEmail, subject,
      'Your mentorship session with ' + mentorName + ' is confirmed on ' + formattedDate + ' at ' + formattedTime + '. Join: ' + meetLink,
      { htmlBody: body, name: 'SSB Mentorship Portal' });
  } catch (e) { Logger.log('sendStudentConfirmationEmail error: ' + e.toString()); }
}

function sendMentorNotificationEmail(mentorEmail, mentorName, studentName, studentEmail, rollNo, batch, sessionDateTime, meetLink) {
  try {
    var formattedDate = Utilities.formatDate(sessionDateTime, Session.getScriptTimeZone(), 'EEEE, dd MMMM yyyy');
    var formattedTime = Utilities.formatDate(sessionDateTime, Session.getScriptTimeZone(), 'hh:mm a z');
    var subject = '📅 New Mentorship Session Booked — ' + studentName;
    var body = '<!DOCTYPE html><html><body style="font-family:\'Segoe UI\',Arial,sans-serif;max-width:600px;margin:0 auto;background:#f8f9fa;padding:20px;">' +
      '<div style="background:linear-gradient(135deg,#1a1a2e 0%,#16213e 100%);padding:30px;border-radius:12px 12px 0 0;text-align:center;">' +
      '<h1 style="color:#00d4aa;margin:0;font-size:24px;">SSB Mentorship Portal</h1>' +
      '<p style="color:#a0aec0;margin:8px 0 0;">Scaler School of Business</p></div>' +
      '<div style="background:white;padding:32px;border-radius:0 0 12px 12px;box-shadow:0 4px 20px rgba(0,0,0,0.08);">' +
      '<h2 style="color:#1a1a2e;margin-top:0;">New session booked with you</h2>' +
      '<p style="color:#4a5568;">Hi <strong>' + mentorName + '</strong>,</p>' +
      '<p style="color:#4a5568;">A student has booked a mentorship session with you.</p>' +
      '<div style="background:#f0fdf9;border-left:4px solid #00d4aa;padding:20px;border-radius:8px;margin:24px 0;">' +
      '<table style="width:100%;border-collapse:collapse;">' +
      '<tr><td style="padding:8px 0;color:#718096;width:140px;">Student</td><td style="padding:8px 0;font-weight:600;color:#1a1a2e;">' + studentName + '</td></tr>' +
      '<tr><td style="padding:8px 0;color:#718096;">Email</td><td style="padding:8px 0;color:#1a1a2e;">' + studentEmail + '</td></tr>' +
      '<tr><td style="padding:8px 0;color:#718096;">Roll No</td><td style="padding:8px 0;color:#1a1a2e;">' + rollNo + '</td></tr>' +
      '<tr><td style="padding:8px 0;color:#718096;">Batch</td><td style="padding:8px 0;color:#1a1a2e;">' + batch + '</td></tr>' +
      '<tr><td style="padding:8px 0;color:#718096;">Date</td><td style="padding:8px 0;font-weight:600;color:#1a1a2e;">' + formattedDate + '</td></tr>' +
      '<tr><td style="padding:8px 0;color:#718096;">Time</td><td style="padding:8px 0;font-weight:600;color:#1a1a2e;">' + formattedTime + '</td></tr>' +
      '<tr><td style="padding:8px 0;color:#718096;">Duration</td><td style="padding:8px 0;font-weight:600;color:#1a1a2e;">60 Minutes</td></tr>' +
      '</table></div>' +
      '<div style="text-align:center;margin:24px 0;">' +
      '<a href="' + meetLink + '" style="background:#00d4aa;color:white;padding:14px 32px;border-radius:8px;text-decoration:none;font-weight:600;font-size:16px;display:inline-block;">Join Google Meet</a></div>' +
      '<p style="color:#718096;font-size:14px;">A calendar invite has been sent to your email.</p>' +
      '<hr style="border:none;border-top:1px solid #e2e8f0;margin:24px 0;">' +
      '<p style="color:#a0aec0;font-size:12px;text-align:center;">SSB Mentorship Portal · Scaler School of Business · Bengaluru</p>' +
      '</div></body></html>';
    GmailApp.sendEmail(mentorEmail, subject,
      'New session booked with ' + studentName + ' on ' + formattedDate + ' at ' + formattedTime + '. Join: ' + meetLink,
      { htmlBody: body, name: 'SSB Mentorship Portal' });
  } catch (e) { Logger.log('sendMentorNotificationEmail error: ' + e.toString()); }
}

// ============================================================
// TRIGGER-BASED MENTOR FEEDBACK — fires 1 hour after session ends
// ============================================================

function scheduleMentorFeedbackTrigger(rowIndex, sessionDateTime) {
  try {
    // Fire 2 minutes after session start (for testing)
    var triggerTime = new Date(sessionDateTime.getTime() + 2 * 60 * 1000);

    // Store the rowIndex in Script Properties so the trigger knows which row to process
    var props = PropertiesService.getScriptProperties();
    var pending = props.getProperty('PENDING_FEEDBACK_TRIGGERS');
    var pendingList = pending ? JSON.parse(pending) : [];
    pendingList.push({ rowIndex: rowIndex, triggerTime: triggerTime.toISOString() });
    props.setProperty('PENDING_FEEDBACK_TRIGGERS', JSON.stringify(pendingList));

    // Create a one-time time-based trigger
    ScriptApp.newTrigger('sendScheduledMentorFeedbackEmail')
      .timeBased()
      .at(triggerTime)
      .create();

    Logger.log('Scheduled mentor feedback trigger for row ' + rowIndex + ' at ' + triggerTime);
  } catch (e) {
    Logger.log('scheduleMentorFeedbackTrigger error: ' + e.toString());
  }
}

function sendScheduledMentorFeedbackEmail() {
  try {
    var props = PropertiesService.getScriptProperties();
    var pending = props.getProperty('PENDING_FEEDBACK_TRIGGERS');
    if (!pending) return;

    var pendingList = JSON.parse(pending);
    var now = new Date();
    var remaining = [];

    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(SHEET_SESSIONS);
    var data = sheet.getDataRange().getValues();

    for (var i = 0; i < pendingList.length; i++) {
      var item = pendingList[i];
      var triggerTime = new Date(item.triggerTime);

      // Process items whose trigger time has passed (within a 10-min window)
      if (now >= triggerTime && (now - triggerTime) < 10 * 60 * 1000) {
        var rowIndex = item.rowIndex;
        if (rowIndex < 1 || rowIndex >= data.length) continue;

        var row = data[rowIndex];
        if (!row[0]) continue;

        // Skip if mentor feedback already submitted
        if (row[11] && row[12]) {
          Logger.log('Mentor feedback already submitted for row ' + rowIndex + ', skipping.');
          continue;
        }

        var mentorName = row[6];
        var mentorEmail = getMentorEmailByName(mentorName);
        if (!mentorEmail) {
          Logger.log('No mentor email found for: ' + mentorName);
          continue;
        }

        var token = generateFeedbackToken(rowIndex, row[0], mentorName);
        var feedbackUrl = ScriptApp.getService().getUrl() + '?token=' + token;

        sendMentorFeedbackEmail(mentorEmail, mentorName, row[1], row[0], row[7], feedbackUrl, row[5]);
        Logger.log('Sent mentor feedback email for row ' + rowIndex + ' to ' + mentorEmail);

      } else if (now < triggerTime) {
        // Not yet due — keep in list
        remaining.push(item);
      }
      // Items older than 10 min window are dropped (already processed or missed)
    }

    props.setProperty('PENDING_FEEDBACK_TRIGGERS', JSON.stringify(remaining));

    // Clean up past triggers for this function to avoid accumulation
    cleanupTriggers('sendScheduledMentorFeedbackEmail');

  } catch (e) {
    Logger.log('sendScheduledMentorFeedbackEmail error: ' + e.toString());
  }
}

function cleanupTriggers(functionName) {
  try {
    var triggers = ScriptApp.getProjectTriggers();
    var now = new Date();
    for (var i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === functionName) {
        // Delete triggers that are in the past
        ScriptApp.deleteTrigger(triggers[i]);
      }
    }
  } catch (e) {
    Logger.log('cleanupTriggers error: ' + e.toString());
  }
}

function sendMentorFeedbackRequestIfNeeded(studentEmail, currentMentorEmail, currentMentorName, newRowIndex) {
  // This function is now a no-op — feedback emails are sent via scheduled triggers
  // (scheduleMentorFeedbackTrigger is called directly from bookSession)
  Logger.log('sendMentorFeedbackRequestIfNeeded called — handled by trigger scheduler instead.');
}

function getMentorEmailByName(mentorName) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(SHEET_MENTORS);
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] && data[i][1].toString().trim() === mentorName.toString().trim()) {
        return data[i][2];
      }
    }
    return null;
  } catch (e) { return null; }
}

function generateFeedbackToken(rowIndex, studentEmail, mentorName) {
  var raw = rowIndex + '|' + studentEmail + '|' + mentorName;
  return Utilities.base64Encode(raw);
}

function sendMentorFeedbackEmail(mentorEmail, mentorName, studentName, studentEmail, sessionDate, feedbackUrl, msnSession) {
  try {
    var formattedDate = sessionDate ? Utilities.formatDate(new Date(sessionDate), Session.getScriptTimeZone(), 'dd MMM yyyy') : 'N/A';
    var subject = '📝 Feedback Requested — Session with ' + studentName + ' (' + msnSession + ')';
    var body = '<!DOCTYPE html><html><body style="font-family:\'Segoe UI\',Arial,sans-serif;max-width:600px;margin:0 auto;background:#f8f9fa;padding:20px;">' +
      '<div style="background:linear-gradient(135deg,#1a1a2e 0%,#16213e 100%);padding:30px;border-radius:12px 12px 0 0;text-align:center;">' +
      '<h1 style="color:#00d4aa;margin:0;font-size:24px;">SSB Mentorship Portal</h1>' +
      '<p style="color:#a0aec0;margin:8px 0 0;">Scaler School of Business</p></div>' +
      '<div style="background:white;padding:32px;border-radius:0 0 12px 12px;box-shadow:0 4px 20px rgba(0,0,0,0.08);">' +
      '<h2 style="color:#1a1a2e;margin-top:0;">Please submit your session feedback</h2>' +
      '<p style="color:#4a5568;">Hi <strong>' + mentorName + '</strong>,</p>' +
      '<p style="color:#4a5568;">Your feedback for the completed mentorship session is requested:</p>' +
      '<div style="background:#f0fdf9;border-left:4px solid #00d4aa;padding:20px;border-radius:8px;margin:24px 0;">' +
      '<table style="width:100%;border-collapse:collapse;">' +
      '<tr><td style="padding:8px 0;color:#718096;width:140px;">Session</td><td style="padding:8px 0;font-weight:600;color:#1a1a2e;">' + msnSession + '</td></tr>' +
      '<tr><td style="padding:8px 0;color:#718096;">Student</td><td style="padding:8px 0;font-weight:600;color:#1a1a2e;">' + studentName + '</td></tr>' +
      '<tr><td style="padding:8px 0;color:#718096;">Student Email</td><td style="padding:8px 0;color:#1a1a2e;">' + studentEmail + '</td></tr>' +
      '<tr><td style="padding:8px 0;color:#718096;">Session Date</td><td style="padding:8px 0;font-weight:600;color:#1a1a2e;">' + formattedDate + '</td></tr>' +
      '</table></div>' +
      '<div style="text-align:center;margin:24px 0;">' +
      '<a href="' + feedbackUrl + '" style="background:#00d4aa;color:white;padding:14px 32px;border-radius:8px;text-decoration:none;font-weight:600;font-size:16px;display:inline-block;">Submit Feedback</a></div>' +
      '<p style="color:#718096;font-size:14px;">This link is unique to this session. Please do not share it.</p>' +
      '<hr style="border:none;border-top:1px solid #e2e8f0;margin:24px 0;">' +
      '<p style="color:#a0aec0;font-size:12px;text-align:center;">SSB Mentorship Portal · Scaler School of Business · Bengaluru</p>' +
      '</div></body></html>';
    GmailApp.sendEmail(mentorEmail, subject,
      'Please submit feedback for your session with ' + studentName + ' on ' + formattedDate + ': ' + feedbackUrl,
      { htmlBody: body, name: 'SSB Mentorship Portal' });
  } catch (e) { Logger.log('sendMentorFeedbackEmail error: ' + e.toString()); }
}

// ============================================================
// UTILITY
// ============================================================

function getAppUrl() {
  return ScriptApp.getService().getUrl();
}
