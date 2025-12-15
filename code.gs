
// === CONFIGURATION ===
const BUFFER_SHEET_ID = "19OLhS4OzvtgPsVHKVigjvrw3YR9K-sWf0U-TWvJ1Ftw"; 
const SPREADSHEET_ID = "1FotLFASWuFinDnvpyLTsyO51OpJeKWtuG31VFje3Oik"; // Only the ID
const SICK_NOTE_FOLDER_ID = "1Wu_eoEQ3FmfrzOdAwJkqMu4sPucLRu_0";
const SHEET_NAMES = {
  adherence: "Adherence Tracker",
  employeesCore: "Employees_Core", 
  employeesPII: "Employees_PII",   
  assets: "Assets",                
  projects: "Projects",            
  projectLogs: "Project_Logs",     
  schedule: "Schedules",
  logs: "Logs",
  otherCodes: "Other Codes",
  leaveRequests: "Leave Requests", 
  coachingSessions: "CoachingSessions", 
  coachingScores: "CoachingScores", 
  coachingTemplates: "CoachingTemplates", 
  pendingRegistrations: "PendingRegistrations",
  movementRequests: "MovementRequests",
  announcements: "Announcements",
  roleRequests: "Role Requests",
  recruitment: "Recruitment_Candidates",
  requisitions: "Requisitions",
  performance: "Performance_Reviews", 
  historyLogs: "Employee_History",
  warnings: "Warnings",
  financialEntitlements: "Financial_Entitlements",
  rbac: "RBAC_Config",
  overtime: "Overtime_Requests",
  breakConfig: "Break_Config",
  offboarding: "Offboarding_Requests" // <--- NEW
};
// --- Break Time Configuration (in seconds) ---
const PLANNED_BREAK_SECONDS = 15 * 60; // 15 minutes
const PLANNED_LUNCH_SECONDS = 30 * 60; // 30 minutes

// --- Shift Cutoff Hour (e.g., 7 = 7 AM) ---
const SHIFT_CUTOFF_HOUR = 7; 

// ================= WEB APP ENTRY (PHASE 4 UPDATED) =================
function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('KOMPASS (Internal)')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
// ================= WEB APP APIs (UPDATED) =================

function webPunch(action, targetUserName, adminTimestamp, projectId) { 
  try {
    // 1. SMART CONTEXT (Load Data)
    const { userEmail, userName: selfName, userData, ss } = getAuthorizedContext(null);

    // 2. Validate Target
    const targetEmail = userData.nameToEmail[targetUserName];
    if (!targetEmail) throw new Error(`User "${targetUserName}" not found.`);

    // 3. PERMISSION CHECK
    if (targetEmail.toLowerCase() !== userEmail.toLowerCase()) {
        getAuthorizedContext('PUNCH_OTHERS'); // Throws error if missing permission
    }

    // 4. Run Logic
    const puncherEmail = userEmail;
    const resultMessage = punch(action, targetUserName, puncherEmail, adminTimestamp);
    
    if (projectId || action === "Logout") {
      logProjectHours(targetUserName, action, projectId, adminTimestamp);
    }

    // *** CRITICAL FIX: FORCE DATA SAVE BEFORE READING STATUS ***
    SpreadsheetApp.flush(); 
    // ***********************************************************

    // 5. Get New Status
    const timeZone = Session.getScriptTimeZone();
    const now = adminTimestamp ? new Date(adminTimestamp) : new Date();
    const shiftDate = getShiftDate(now, SHIFT_CUTOFF_HOUR);
    const formattedDate = Utilities.formatDate(shiftDate, timeZone, "MM/dd/yyyy");
    
    const newStatus = getLatestPunchStatus(targetEmail, targetUserName, shiftDate, formattedDate);
    
    return { message: resultMessage, newStatus: newStatus };

  } catch (err) { return { message: "Error: " + err.message, newStatus: null };
  }
}

// === NEW HELPER FOR PHASE 3 ===
function logProjectHours(userName, action, newProjectId, customTime) {
  const ss = getSpreadsheet();
  const coreSheet = getOrCreateSheet(ss, SHEET_NAMES.employeesCore);
  const logSheet = getOrCreateSheet(ss, SHEET_NAMES.projectLogs);
  const data = coreSheet.getDataRange().getValues();
  
  // 1. Find User Row & Current State
  let userRowIndex = -1;
  let currentProjectId = "";
  let lastActionTime = null;
  let empID = "";

  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === userName) { // Match Name
      userRowIndex = i + 1;
      empID = data[i][0]; // EmployeeID is Col A
      // We use Column K (Index 10) for "CurrentProject" and L (Index 11) for "LastActionTime"
      // If they don't exist yet, we treat them as empty.
      currentProjectId = data[i][10] || ""; 
      lastActionTime = data[i][11] ? new Date(data[i][11]) : null;
      break;
    }
  }

  if (userRowIndex === -1) return; // Should not happen

  const now = customTime ? new Date(customTime) : new Date();

  // 2. If they were working on a project, calculate duration and log it
  if (currentProjectId && lastActionTime) {
    const durationHours = (now.getTime() - lastActionTime.getTime()) / (1000 * 60 * 60);
    
    if (durationHours > 0) {
      logSheet.appendRow([
        `LOG-${new Date().getTime()}`, // LogID
        empID,
        currentProjectId,
        new Date(), // Date of log
        durationHours.toFixed(2) // Duration
      ]);
    }
  }

  // 3. Update State in Employees_Core
  // If Logout, clear the project. If Login/Switch, set the new project.
  if (action === "Logout") {
    coreSheet.getRange(userRowIndex, 11).setValue(""); // Clear Project
    coreSheet.getRange(userRowIndex, 12).setValue(""); // Clear Time
  } else {
    coreSheet.getRange(userRowIndex, 11).setValue(newProjectId); // Set New Project
    coreSheet.getRange(userRowIndex, 12).setValue(now); // Set Start Time
  }
}

function webSubmitScheduleRange(userEmail, userName, startDateStr, endDateStr, startTime, endTime, leaveType, shiftEndDate) {
  try {
    const { userEmail: puncherEmail } = getAuthorizedContext('EDIT_SCHEDULE');
    return submitScheduleRange(puncherEmail, userEmail, userName, startDateStr, endDateStr, startTime, endTime, leaveType, shiftEndDate);
  } catch (err) { return "Error: " + err.message; }
}

// === Web App APIs for Leave Requests ===
function webSubmitLeaveRequest(requestObject, targetUserEmail) { // Now accepts optional target user
  try {
    const submitterEmail = Session.getActiveUser().getEmail().toLowerCase();
    return submitLeaveRequest(submitterEmail, requestObject, targetUserEmail);
  } catch (err) {
    return "Error: " + err.message;
  }
}

function webGetMyRequests_V2() {
  try {
    const userEmail = Session.getActiveUser().getEmail().toLowerCase();
    return getMyRequests(userEmail); 
  } catch (err) {
    Logger.log("Error in webGetMyRequests_V2: " + err.message);
    throw new Error(err.message); 
  }
}

function webGetAdminLeaveRequests(filter) {
  try {
    const adminEmail = Session.getActiveUser().getEmail().toLowerCase();
    return getAdminLeaveRequests(adminEmail, filter);
  } catch (err) {
    Logger.log("webGetAdminLeaveRequests Error: " + err.message);
    return { error: err.message };
  }
}

function webApproveDenyRequest(requestID, newStatus, reason) {
  try {
    const adminEmail = Session.getActiveUser().getEmail().toLowerCase();
    return approveDenyRequest(adminEmail, requestID, newStatus, reason);
  } catch (err) {
    return "Error: " + err.message;
  }
}

// === Web App API for History ===
function webGetAdherenceRange(userNames, startDateStr, endDateStr) {
  try {
    const userEmail = Session.getActiveUser().getEmail().toLowerCase();
    return getAdherenceRange(userEmail, userNames, startDateStr, endDateStr);
  } catch (err) {
    return { error: "Error: " + err.message };
  }
}

// === Web App API for My Schedule ===
function webGetMySchedule() {
  try {
    const userEmail = Session.getActiveUser().getEmail().toLowerCase();
    return getMySchedule(userEmail);
  } catch (err) {
    return { error: "Error: " + err.message };
  }
}

// === Web App API for Admin Tools ===
function webAdjustLeaveBalance(userEmail, leaveType, amount, reason) {
  try {
    const adminEmail = Session.getActiveUser().getEmail().toLowerCase();
    return adjustLeaveBalance(adminEmail, userEmail, leaveType, amount, reason);
  } catch (err) {
    return "Error: " + err.message;
  }
}

function webImportScheduleCSV(csvData) {
  try {
    const adminEmail = Session.getActiveUser().getEmail().toLowerCase();
    return importScheduleCSV(adminEmail, csvData);
  } catch (err) {
    return "Error: " + err.message;
  }
}

// === Web App API for Dashboard ===
function webGetDashboardData(userEmails, date) { 
  try {
    const { userEmail: adminEmail } = getAuthorizedContext('VIEW_FULL_DASHBOARD');
    return getDashboardData(adminEmail, userEmails, date);
  } catch (err) {
    Logger.log("webGetDashboardData Error: " + err.message);
    throw new Error(err.message);
  }
}

// --- MODIFIED: "My Team" Functions ---
function webSaveMyTeam(userEmails) {
  try {
    const adminEmail = Session.getActiveUser().getEmail().toLowerCase();
    return saveMyTeam(adminEmail, userEmails);
  } catch (err) {
    return "Error: " + err.message;
  }
}

function webGetMyTeam() {
  try {
    const adminEmail = Session.getActiveUser().getEmail().toLowerCase();
    return getMyTeam(adminEmail);
  } catch (err) {
    return "Error: " + err.message;
  }
}

function webSubmitMovementRequest(userToMoveEmail, newSupervisorEmail, newProjectManagerEmail) {
  const { userEmail: requesterEmail, userData, ss } = getAuthorizedContext('MANAGE_HIERARCHY');
  
  const userToMoveName = userData.emailToName[userToMoveEmail];
  const newSupervisorName = userData.emailToName[newSupervisorEmail];
  
  // 1. Validation
  if (!userToMoveName) throw new Error(`User to move (${userToMoveEmail}) not found.`);
  if (!newSupervisorName) throw new Error(`Receiving supervisor (${newSupervisorEmail}) not found.`);
  
  // FIX: Ensure Project Manager is present. 
  // If "Keep" was selected but user had no PM, frontend might send blank.
  if (!newProjectManagerEmail || newProjectManagerEmail === "" || newProjectManagerEmail === "undefined") {
     throw new Error("Project Manager selection is required. If the user has no current PM, please select 'Change' and assign one.");
  }

  const requesterRole = userData.emailToRole[requesterEmail];
  const fromSupervisorEmail = userData.emailToSupervisor[userToMoveEmail];
  const moveSheet = getOrCreateSheet(ss, SHEET_NAMES.movementRequests);

  // 2. SUPER ADMIN LOGIC: Immediate Execution
  if (requesterRole === 'superadmin') {
      const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.employeesCore);
      const userDBRow = userData.emailToRow[userToMoveEmail];
      
      if (!userDBRow) throw new Error("Database row not found for user.");

      // Update Employees_Core
      // Direct Manager is Column F (Index 6)
      // Functional/Project Manager is Column G (Index 7)
      dbSheet.getRange(userDBRow, 6).setValue(newSupervisorEmail);
      dbSheet.getRange(userDBRow, 7).setValue(newProjectManagerEmail);

      // Log as "Approved" in Movement Requests
      moveSheet.appendRow([
        `MOV-${new Date().getTime()}`,
        "Approved", // Auto-approved
        userToMoveEmail,
        userToMoveName,
        fromSupervisorEmail,
        newSupervisorEmail,
        new Date(), // Request Time
        new Date(), // Action Time
        requesterEmail, // Action By
        requesterEmail, // Requested By
        newProjectManagerEmail
      ]);
      
      // Log to System Logs
      const logsSheet = getOrCreateSheet(ss, SHEET_NAMES.logs);
      logsSheet.appendRow([new Date(), userToMoveName, requesterEmail, "Movement Auto-Approved", `Moved to ${newSupervisorName} / ${newProjectManagerEmail}`]);

      return `Success: ${userToMoveName} has been moved to ${newSupervisorName} immediately.`;
  } 
  
  // 3. ADMIN LOGIC: Submit for Approval
  else {
      moveSheet.appendRow([
        `MOV-${new Date().getTime()}`,
        "Pending",
        userToMoveEmail,
        userToMoveName,
        fromSupervisorEmail,
        newSupervisorEmail,
        new Date(),
        "", // ActionTimestamp
        "", // ActionBy
        requesterEmail,
        newProjectManagerEmail
      ]);
      
      return `Request submitted. Waiting for approval from ${newSupervisorName} (or Superadmin).`;
  }
}
/**
 * NEW: Fetches pending movement requests for the admin or their subordinates.
 */
function webGetPendingMovements() {
  try {
    const adminEmail = Session.getActiveUser().getEmail().toLowerCase();
    const ss = getSpreadsheet();
    const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
    const userData = getUserDataFromDb(dbSheet);

    // *** ADD THIS LINE TO FIX THE ERROR ***
    const adminRole = userData.emailToRole[adminEmail] || 'agent';
    
    // Get all subordinates (direct and indirect)
    const mySubordinateEmails = new Set(webGetAllSubordinateEmails(adminEmail));
    const moveSheet = getOrCreateSheet(ss, SHEET_NAMES.movementRequests);
    const data = moveSheet.getDataRange().getValues();
    const results = [];

    // Get headers
    const headers = data[0];
    const statusIndex = headers.indexOf("Status");
    const toSupervisorIndex = headers.indexOf("ToSupervisorEmail");
    
    if (statusIndex === -1 || toSupervisorIndex === -1) {
      throw new Error("MovementRequests sheet is missing required columns.");
    }

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const status = row[statusIndex];
      const toSupervisorEmail = (row[toSupervisorIndex] || "").toLowerCase();

      if (status === 'Pending') {
        let canView = false;
        
        // --- NEW VIEWING LOGIC ---
        if (adminRole === 'superadmin') {
          // Superadmin can see ALL pending requests
          canView = true;
        } else if (toSupervisorEmail === adminEmail || mySubordinateEmails.has(toSupervisorEmail)) {
          // Admin can only see requests for themselves or their subordinates
          canView = true;
        }
        // --- END NEW LOGIC ---

        if (canView) {
          results.push({
            movementID: row[headers.indexOf("MovementID")],
            userToMoveName: row[headers.indexOf("UserToMoveName")],
            fromSupervisorName: userData.emailToName[row[headers.indexOf("FromSupervisorEmail")]] || "Unknown",
            
  toSupervisorName: userData.emailToName[row[headers.indexOf("ToSupervisorEmail")]] || "Unknown",
            requestedDate: convertDateToString(new Date(row[headers.indexOf("RequestTimestamp")])),
            requestedByName: userData.emailToName[row[headers.indexOf("RequestedByEmail")]] || "Unknown"
          });
}
      }
    }
    return results;
  } catch (e) {
    Logger.log("webGetPendingMovements Error: " + e.message);
    return { error: e.message };
  }
}

function webApproveDenyMovement(movementID, newStatus) {
  try {
    const adminEmail = Session.getActiveUser().getEmail().toLowerCase();
    const ss = getSpreadsheet();
    const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.employeesCore); // Use Core for updates
    const userData = getUserDataFromDb(ss);
    const moveSheet = getOrCreateSheet(ss, SHEET_NAMES.movementRequests);
    const data = moveSheet.getDataRange().getValues();
    
    // Map Headers to find Columns
    const headers = data[0];
    const idIndex = headers.indexOf("MovementID");
    const statusIndex = headers.indexOf("Status");
    const toSupervisorIndex = headers.indexOf("ToSupervisorEmail");
    const userToMoveIndex = headers.indexOf("UserToMoveEmail");
    const actionTimeIndex = headers.indexOf("ActionTimestamp");
    const actionByIndex = headers.indexOf("ActionByEmail");
    // New Header might not exist in old sheets, so we use fixed index 10 (Column 11) or check
    const toProjMgrIndex = 10; // 0-based index for Column K

    let rowToUpdate = -1;
    let requestDetails = {};

    for (let i = 1; i < data.length; i++) {
      if (data[i][idIndex] === movementID) {
        rowToUpdate = i + 1;
        requestDetails = {
          status: data[i][statusIndex],
          toSupervisorEmail: (data[i][toSupervisorIndex] || "").toLowerCase(),
          toProjectManagerEmail: (data[i][toProjMgrIndex] || "").toLowerCase(), // Read New PM
          userToMoveEmail: (data[i][userToMoveIndex] || "").toLowerCase()
        };
        break;
      }
    }

    if (rowToUpdate === -1) throw new Error("Movement request not found.");
    if (requestDetails.status !== 'Pending') throw new Error(`Request is already ${requestDetails.status}.`);

    // --- Security Check ---
    // Approver must be the NEW Supervisor OR a Superadmin
    // (Or the Admin of the new supervisor)
    const adminRole = userData.emailToRole[adminEmail];
    
    // Get admin's hierarchy
    const mySubordinateEmails = new Set(webGetAllSubordinateEmails(adminEmail));
    const isReceivingSupervisor = (requestDetails.toSupervisorEmail === adminEmail);
    const isSupervisorOfReceiver = mySubordinateEmails.has(requestDetails.toSupervisorEmail);

    if (adminRole !== 'superadmin' && !isReceivingSupervisor && !isSupervisorOfReceiver) {
      throw new Error("Permission denied. You can only approve requests assigned to you or your hierarchy.");
    }

    // Update Status
    moveSheet.getRange(rowToUpdate, statusIndex + 1).setValue(newStatus);
    moveSheet.getRange(rowToUpdate, actionTimeIndex + 1).setValue(new Date());
    moveSheet.getRange(rowToUpdate, actionByIndex + 1).setValue(adminEmail);

    if (newStatus === 'Approved') {
      const userDBRow = userData.emailToRow[requestDetails.userToMoveEmail];
      if (!userDBRow) throw new Error(`User ${requestDetails.userToMoveEmail} not found in DB.`);
      
      // Update Direct Manager (Col F = 6)
      dbSheet.getRange(userDBRow, 6).setValue(requestDetails.toSupervisorEmail);
      
      // Update Project Manager (Col G = 7)
      // Only if we have a valid new project manager
      if (requestDetails.toProjectManagerEmail) {
         dbSheet.getRange(userDBRow, 7).setValue(requestDetails.toProjectManagerEmail);
      }

      // Log
      const logsSheet = getOrCreateSheet(ss, SHEET_NAMES.logs);
      logsSheet.appendRow([
        new Date(), 
        userData.emailToName[requestDetails.userToMoveEmail], 
        adminEmail, 
        "Reporting Line Change Approved", 
        `Direct: ${requestDetails.toSupervisorEmail}, Project: ${requestDetails.toProjectManagerEmail}`
      ]);
    }
    
    SpreadsheetApp.flush();
    return { success: true, message: `Request ${newStatus}.` };

  } catch (e) {
    return { error: e.message };
  }
}

/**
 * NEW: Fetches the movement history for a selected user.
 */
function webGetMovementHistory(selectedUserEmail) {
  try {
    const adminEmail = Session.getActiveUser().getEmail().toLowerCase();
    const ss = getSpreadsheet();
    const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
    const userData = getUserDataFromDb(dbSheet);
    
    // Security check: Is this admin allowed to see this user's history?
    const adminRole = userData.emailToRole[adminEmail];
    const mySubordinateEmails = new Set(webGetAllSubordinateEmails(adminEmail));

    if (adminRole !== 'superadmin' && !mySubordinateEmails.has(selectedUserEmail)) {
      throw new Error("Permission denied. You can only view the history of users in your reporting line.");
    }
    
    const moveSheet = getOrCreateSheet(ss, SHEET_NAMES.movementRequests);
    const data = moveSheet.getDataRange().getValues();
    const headers = data[0];
    const results = [];

    // Find rows where the user was the one being moved
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const userToMoveEmail = (row[headers.indexOf("UserToMoveEmail")] || "").toLowerCase();
      
      if (userToMoveEmail === selectedUserEmail) {
        results.push({
          status: row[headers.indexOf("Status")],
          requestDate: convertDateToString(new Date(row[headers.indexOf("RequestTimestamp")])),
          actionDate: convertDateToString(new Date(row[headers.indexOf("ActionTimestamp")])),
          fromSupervisorName: userData.emailToName[row[headers.indexOf("FromSupervisorEmail")]] || "N/A",
          toSupervisorName: userData.emailToName[row[headers.indexOf("ToSupervisorEmail")]] || "N/A",
          actionByName: userData.emailToName[row[headers.indexOf("ActionByEmail")]] || "N/A",
          requestedByName: userData.emailToName[row[headers.indexOf("RequestedByEmail")]] || "N/A"
        });
      }
    }
    
    // Sort by request date, newest first
    results.sort((a, b) => new Date(b.requestDate) - new Date(a.requestDate));
    return results;

  } catch (e) {
    Logger.log("webGetMovementHistory Error: " + e.message);
    return { error: e.message };
  }
}

// ==========================================================
// === NEW/REPLACED COACHING FUNCTIONS (START) ===
// ==========================================================

/**
 * (REPLACED)
 * Saves a new coaching session and its detailed scores.
 * Matches the new frontend form.
 */
function webSubmitCoaching(sessionObject) {
  try {
    const ss = getSpreadsheet();
    const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
    const userData = getUserDataFromDb(dbSheet);
    const sessionSheet = getOrCreateSheet(ss, SHEET_NAMES.coachingSessions);
    const scoreSheet = getOrCreateSheet(ss, SHEET_NAMES.coachingScores);
    
    const coachEmail = Session.getActiveUser().getEmail().toLowerCase();
    const coachName = userData.emailToName[coachEmail] || coachEmail;
    
    // Simple validation
    if (!sessionObject.agentEmail || !sessionObject.sessionDate) {
      throw new Error("Agent and Session Date are required.");
    }

    const agentName = userData.emailToName[sessionObject.agentEmail.toLowerCase()];
    if (!agentName) {
      throw new Error(`Could not find agent with email ${sessionObject.agentEmail}.`);
    }

    const sessionID = `CS-${new Date().getTime()}`; // Simple unique ID
    const sessionDate = new Date(sessionObject.sessionDate + 'T00:00:00');
    // *** NEW: Handle FollowUpDate ***
    const followUpDate = sessionObject.followUpDate ? new Date(sessionObject.followUpDate + 'T00:00:00') : null;
    const followUpStatus = followUpDate ? "Pending" : ""; // Set to pending if date exists

    // 1. Log the main session
    sessionSheet.appendRow([
      sessionID,
      sessionObject.agentEmail,
      agentName,
      coachEmail,
      coachName,
      sessionDate,
      sessionObject.weekNumber,
      sessionObject.overallScore,
      sessionObject.followUpComment,
      new Date(), // Timestamp of submission
      followUpDate || "", // *** NEW: Add follow-up date ***
      followUpStatus  // *** NEW: Add follow-up status ***
    ]);

    // 2. Log the individual scores
    const scoresToLog = [];
    if (sessionObject.scores && Array.isArray(sessionObject.scores)) {
      sessionObject.scores.forEach(score => {
        scoresToLog.push([
          sessionID,
          score.category,
          score.criteria,
          score.score,
          score.comment
        ]);
      });
    }

    if (scoresToLog.length > 0) {
      scoreSheet.getRange(scoreSheet.getLastRow() + 1, 1, scoresToLog.length, 5).setValues(scoresToLog);
    }
    
    return `Coaching session for ${agentName} saved successfully.`;

  } catch (err) {
    Logger.log("webSubmitCoaching Error: " + err.message);
    return "Error: " + err.message;
  }
}

/**
 * (REPLACED)
 * Gets coaching history for the logged-in user or their team.
 * Reads from the new CoachingSessions sheet.
 */
function webGetCoachingHistory(filter) { // filter is unused for now, but good practice
  try {
    const userEmail = Session.getActiveUser().getEmail().toLowerCase();
    const ss = getSpreadsheet();
    const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
    const userData = getUserDataFromDb(dbSheet);
    const role = userData.emailToRole[userEmail] || 'agent';
    const sheet = getOrCreateSheet(ss, SHEET_NAMES.coachingSessions);

    // Get all data as objects
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const allData = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    
    const allSessions = allData.map(row => {
      let obj = {};
      headers.forEach((header, index) => {
        obj[header] = row[index];
      });
      return obj;
    });

    const results = [];
    
    // Get a list of users this person manages (if they are a manager)
    let myTeamEmails = new Set();
    if (role === 'admin' || role === 'superadmin') {
      // Use the hierarchy-aware function
      const myTeamList = webGetAllSubordinateEmails(userEmail);
      myTeamList.forEach(email => myTeamEmails.add(email.toLowerCase()));
    }

    for (let i = allSessions.length - 1; i >= 0; i--) {
      const session = allSessions[i];
      if (!session || !session.AgentEmail) continue; // Skip empty/invalid rows

      const agentEmail = session.AgentEmail.toLowerCase();

      let canView = false;
      
      // *** MODIFIED LOGIC HERE ***
      if (agentEmail === userEmail) {
        // Anyone can see their own coaching
        canView = true;
      } else if (role === 'admin' && myTeamEmails.has(agentEmail)) {
        // An admin can see their team's
        canView = true;
      } else if (role === 'superadmin') {
        // Superadmin can see all (team members + their own, which is covered above)
        canView = true;
      }
      // *** END MODIFIED LOGIC ***

      if (canView) {
        results.push({
          sessionID: session.SessionID,
          agentName: session.AgentName,
          coachName: session.CoachName,
          sessionDate: convertDateToString(new Date(session.SessionDate)),
          weekNumber: session.WeekNumber,
          overallScore: session.OverallScore,
          followUpComment: session.FollowUpComment,
          followUpDate: convertDateToString(new Date(session.FollowUpDate)),
          followUpStatus: session.FollowUpStatus,
          agentAcknowledgementTimestamp: convertDateToString(new Date(session.AgentAcknowledgementTimestamp))
        });
      }
    }
    return results;

  } catch (err) {
    Logger.log("webGetCoachingHistory Error: " + err.message);
    return { error: err.message };
  }
}

/**
 * NEW: Fetches the details for a single coaching session.
 * (MODIFIED: Renamed to webGetCoachingSessionDetails to be callable)
 * (MODIFIED 2: Added date-to-string conversion to fix null return)
 * (MODIFIED 3: Added AgentAcknowledgementTimestamp conversion)
 */
function webGetCoachingSessionDetails(sessionID) {
  try {
    const ss = getSpreadsheet();
    const sessionSheet = getOrCreateSheet(ss, SHEET_NAMES.coachingSessions);
    const scoreSheet = getOrCreateSheet(ss, SHEET_NAMES.coachingScores);

    const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
    const userData = getUserDataFromDb(dbSheet);

    // 1. Get Session Summary
    const sessionHeaders = sessionSheet.getRange(1, 1, 1, sessionSheet.getLastColumn()).getValues()[0];
    const sessionData = sessionSheet.getDataRange().getValues();
    let sessionSummary = null;

    for (let i = 1; i < sessionData.length; i++) {
      if (sessionData[i][0] === sessionID) {
        sessionSummary = {};
        sessionHeaders.forEach((header, index) => {
          sessionSummary[header] = sessionData[i][index];
        });
        break;
      }
    }

    if (!sessionSummary) {
      throw new Error("Session not found.");
    }

    // 2. Get Session Scores
    const scoreHeaders = scoreSheet.getRange(1, 1, 1, scoreSheet.getLastColumn()).getValues()[0];
    const scoreData = scoreSheet.getDataRange().getValues();
    const sessionScores = [];

    for (let i = 1; i < scoreData.length; i++) {
      if (scoreData[i][0] === sessionID) {
        let scoreObj = {};
        scoreHeaders.forEach((header, index) => {
          scoreObj[header] = scoreData[i][index];
        });
        sessionScores.push(scoreObj);
      }
    }
    
    sessionSummary.CoachName = userData.emailToName[sessionSummary.CoachEmail] || sessionSummary.CoachName;
    
    // *** Convert Date objects to Strings before returning ***
    sessionSummary.SessionDate = convertDateToString(new Date(sessionSummary.SessionDate));
    sessionSummary.SubmissionTimestamp = convertDateToString(new Date(sessionSummary.SubmissionTimestamp));
    sessionSummary.FollowUpDate = convertDateToString(new Date(sessionSummary.FollowUpDate));
    // *** NEW: Convert the new column ***
    sessionSummary.AgentAcknowledgementTimestamp = convertDateToString(new Date(sessionSummary.AgentAcknowledgementTimestamp));
    // *** END NEW SECTION ***

    return {
      summary: sessionSummary,
      scores: sessionScores
    };

  } catch (err) {
    Logger.log("webGetCoachingSessionDetails Error: " + err.message);
    return { error: err.message };
  }
}

/**
 * NEW: Updates the follow-up status for a coaching session.
 */
function webUpdateFollowUpStatus(sessionID, newStatus, newDateStr) {
  try {
    const adminEmail = Session.getActiveUser().getEmail().toLowerCase();
    const ss = getSpreadsheet();
    
    // Check permission
    const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
    const userData = getUserDataFromDb(dbSheet);
    const adminRole = userData.emailToRole[adminEmail] || 'agent';

    if (adminRole !== 'admin' && adminRole !== 'superadmin') {
      throw new Error("Permission denied. Only managers can update follow-up status.");
    }
    
    const sessionSheet = getOrCreateSheet(ss, SHEET_NAMES.coachingSessions);
    const sessionData = sessionSheet.getDataRange().getValues();
    const sessionHeaders = sessionData[0];
    
    // Find the column indexes
    const statusColIndex = sessionHeaders.indexOf("FollowUpStatus");
    const dateColIndex = sessionHeaders.indexOf("FollowUpDate");
    
    if (statusColIndex === -1 || dateColIndex === -1) {
      throw new Error("Could not find 'FollowUpStatus' or 'FollowUpDate' columns in CoachingSessions sheet.");
    }

    // Find the row
    let sessionRow = -1;
    for (let i = 1; i < sessionData.length; i++) {
      if (sessionData[i][0] === sessionID) {
        sessionRow = i + 1; // 1-based index
        break;
      }
    }

    if (sessionRow === -1) {
      throw new Error("Session not found.");
    }

    // Prepare new values
    let newFollowUpDate = null;
    if (newDateStr) {
      newFollowUpDate = new Date(newDateStr + 'T00:00:00');
    } else {
      // If marking completed, use today's date
      newFollowUpDate = new Date();
    }
    
    // Update the sheet
    sessionSheet.getRange(sessionRow, statusColIndex + 1).setValue(newStatus);
    sessionSheet.getRange(sessionRow, dateColIndex + 1).setValue(newFollowUpDate);

    SpreadsheetApp.flush(); // Ensure changes are saved

    return { success: true, message: `Status updated to ${newStatus}.` };

  } catch (err) {
    Logger.log("webUpdateFollowUpStatus Error: " + err.message);
    return { error: err.message };
  }
}

/**
 * NEW: Allows an agent to acknowledge their coaching session.
 */
function webSubmitCoachingAcknowledgement(sessionID) {
  try {
    const userEmail = Session.getActiveUser().getEmail().toLowerCase();
    const ss = getSpreadsheet();
    const sessionSheet = getOrCreateSheet(ss, SHEET_NAMES.coachingSessions);

    // *** MODIFIED: Explicitly read headers ***
    const sessionHeaders = sessionSheet.getRange(1, 1, 1, sessionSheet.getLastColumn()).getValues()[0];
    // Get data rows separately, skipping header
    const sessionData = sessionSheet.getRange(2, 1, sessionSheet.getLastRow() - 1, sessionSheet.getLastColumn()).getValues();

    // Find the column indexes
    const ackColIndex = sessionHeaders.indexOf("AgentAcknowledgementTimestamp");
    const agentEmailColIndex = sessionHeaders.indexOf("AgentEmail");
    if (ackColIndex === -1 || agentEmailColIndex === -1) {
      throw new Error("Could not find 'AgentAcknowledgementTimestamp' or 'AgentEmail' columns in CoachingSessions sheet.");
    }

    // Find the row
    let sessionRow = -1;
    let agentEmailOnRow = null;
    let currentAckStatus = null;

    // *** MODIFIED: Loop starts at 0 and row index is i + 2 ***
    for (let i = 0; i < sessionData.length; i++) {
      if (sessionData[i][0] === sessionID) {
        sessionRow = i + 2; // Data starts from row 2
        agentEmailOnRow = sessionData[i][agentEmailColIndex].toLowerCase();
        currentAckStatus = sessionData[i][ackColIndex];
        break;
      }
    }

    if (sessionRow === -1) {
      throw new Error("Session not found.");
    }
    
    // Security Check: Is this the correct agent?
    if (agentEmailOnRow !== userEmail) {
      throw new Error("Permission denied. You can only acknowledge your own coaching sessions.");
    }
    
    // Check if already acknowledged
    if (currentAckStatus) {
      return { success: false, message: "This session has already been acknowledged." };
    }
    
    // Update the sheet
    sessionSheet.getRange(sessionRow, ackColIndex + 1).setValue(new Date());

    SpreadsheetApp.flush(); // Ensure changes are saved

    return { success: true, message: "Coaching session acknowledged successfully." };

  } catch (err) {
    Logger.log("webSubmitCoachingAcknowledgement Error: " + err.message);
    return { error: err.message };
  }
}


/**
 * NEW: Gets a list of unique, active template names.
 */
function webGetActiveTemplates() {
  try {
    const ss = getSpreadsheet();
    const templateSheet = getOrCreateSheet(ss, SHEET_NAMES.coachingTemplates);
    const data = templateSheet.getRange(2, 1, templateSheet.getLastRow() - 1, 4).getValues();
    
    const templateNames = new Set();
    
    data.forEach(row => {
      const templateName = row[0];
      const status = row[3];
      if (templateName && status === 'Active') {
        templateNames.add(templateName);
      }
    });
    
    return Array.from(templateNames).sort();
    
  } catch (err) {
    Logger.log("webGetActiveTemplates Error: " + err.message);
    return { error: err.message };
  }
}

/**
 * NEW: Gets all criteria for a specific template name.
 */
function webGetTemplateCriteria(templateName) {
  try {
    const ss = getSpreadsheet();
    const templateSheet = getOrCreateSheet(ss, SHEET_NAMES.coachingTemplates);
    const data = templateSheet.getRange(2, 1, templateSheet.getLastRow() - 1, 4).getValues();
    
    const categories = {}; // Use an object to group criteria by category
    
    data.forEach(row => {
      const name = row[0];
      const category = row[1];
      const criteria = row[2];
      const status = row[3];
      
      if (name === templateName && status === 'Active' && category && criteria) {
        if (!categories[category]) {
          categories[category] = [];
        }
        categories[category].push(criteria);
      }
    });
    
    // Convert from object to the array structure the frontend expects
    const results = Object.keys(categories).map(categoryName => {
      return {
        category: categoryName,
        criteria: categories[categoryName]
      };
    });
    
    return results;
    
  } catch (err) {
    Logger.log("webGetTemplateCriteria Error: " + err.message);
    return { error: err.message };
  }
}

// ==========================================================
// === NEW/REPLACED COACHING FUNCTIONS (END) ===
// ==========================================================

// [START] MODIFICATION 8: Add webSaveNewTemplate function
/**
 * NEW: Saves a new coaching template from the admin tab.
 */
function webSaveNewTemplate(templateName, categories) {
  try {
    const adminEmail = Session.getActiveUser().getEmail().toLowerCase();
    const ss = getSpreadsheet();
    
    // Check permission
    const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
    const userData = getUserDataFromDb(dbSheet);
    const adminRole = userData.emailToRole[adminEmail] || 'agent';

    if (adminRole !== 'admin' && adminRole !== 'superadmin') {
      throw new Error("Permission denied. Only managers can create templates.");
    }
    
    // Validation
    if (!templateName) {
      throw new Error("Template Name is required.");
    }
    if (!categories || categories.length === 0) {
      throw new Error("At least one category is required.");
    }

    const templateSheet = getOrCreateSheet(ss, SHEET_NAMES.coachingTemplates);
    
    // Check if template name already exists
    const templateNames = templateSheet.getRange(2, 1, templateSheet.getLastRow() - 1, 1).getValues();
    const
      lowerTemplateName = templateName.toLowerCase();
    for (let i = 0; i < templateNames.length; i++) {
      if (templateNames[i][0] && templateNames[i][0].toLowerCase() === lowerTemplateName) {
        throw new Error(`A template with the name '${templateName}' already exists.`);
      }
    }

    const rowsToAppend = [];
    categories.forEach(category => {
      if (category.criteria && category.criteria.length > 0) {
        category.criteria.forEach(criterion => {
          rowsToAppend.push([
            templateName,
            category.name,
            criterion,
            'Active' // Default to Active
          ]);
        });
      }
    });

    if (rowsToAppend.length === 0) {
      throw new Error("No criteria were found to save.");
    }
    
    // Write all new rows at once
    templateSheet.getRange(templateSheet.getLastRow() + 1, 1, rowsToAppend.length, 4).setValues(rowsToAppend);
    
    SpreadsheetApp.flush();
    return `Template '${templateName}' saved successfully with ${rowsToAppend.length} criteria.`;

  } catch (err) {
    Logger.log("webSaveNewTemplate Error: " + err.message);
    return "Error: " + err.message;
  }
}
// [END] MODIFICATION 8

// === NEW: Web App API for Manager Hierarchy ===
function webGetManagerHierarchy() {
  try {
    const managerEmail = Session.getActiveUser().getEmail().toLowerCase();
    const ss = getSpreadsheet();
    const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
    const userData = getUserDataFromDb(dbSheet);
    
    const managerRole = userData.emailToRole[managerEmail] || 'agent';
    if (managerRole === 'agent') {
      return { error: "Permission denied. Only managers can view the hierarchy." };
    }
    
    // --- Step 1: Build the direct reporting map (Supervisor -> [Subordinates]) ---
    const reportsMap = {};
    const userEmailMap = {}; // Map email -> {name, role}

    userData.userList.forEach(user => {
      userEmailMap[user.email] = { name: user.name, role: user.role };
      const supervisorEmail = user.supervisor;
      
      if (supervisorEmail) {
        if (!reportsMap[supervisorEmail]) {
          reportsMap[supervisorEmail] = [];
        }
        reportsMap[supervisorEmail].push(user.email);
      }
    });

    // --- Step 2: Recursive function to build the tree (Hierarchy) ---
    // MODIFIED: Added `visited` Set to track users in the current path.
    function buildHierarchy(currentEmail, depth = 0, visited = new Set()) {
      const user = userEmailMap[currentEmail];
      
      // If the email doesn't map to a user, it's likely a blank entry in the DB, so return null
      if (!user) return null; 
      
      // CRITICAL CHECK: Detect circular reference
      if (visited.has(currentEmail)) {
        Logger.log(`Circular reference detected at user: ${currentEmail}`);
        return {
          email: currentEmail,
          name: user.name,
          role: user.role,
          subordinates: [],
          circularError: true
        };
      }
      
      // Add current user to visited set for this path
      const newVisited = new Set(visited).add(currentEmail);


      const subordinates = reportsMap[currentEmail] || [];
      
      // Separate managers/admins from agents
      const adminSubordinates = subordinates
        .filter(email => userData.emailToRole[email] === 'admin' || userData.emailToRole[email] === 'superadmin')
        .map(email => buildHierarchy(email, depth + 1, newVisited))
        .filter(s => s !== null); // Build sub-teams for managers

      const agentSubordinates = subordinates
        .filter(email => userData.emailToRole[email] === 'agent')
        .map(email => ({
          email: email,
          name: userEmailMap[email].name,
          role: userEmailMap[email].role,
          subordinates: [] // Agents have no subordinates
        }));
        
      // Combine and sort: Managers first, then Agents, then alphabetically
      const combinedSubordinates = [...adminSubordinates, ...agentSubordinates];
      
      combinedSubordinates.sort((a, b) => {
          // Sort by role (manager/admin first)
          const aIsManager = a.role !== 'agent';
          const bIsManager = b.role !== 'agent';
          
          if (aIsManager && !bIsManager) return -1;
          if (!aIsManager && bIsManager) return 1;
          
          // Then sort by name
          return a.name.localeCompare(b.name);
      });


      return {
        email: currentEmail,
        name: user.name,
        role: user.role,
        subordinates: combinedSubordinates,
        depth: depth
      };
    }

    // Start building the hierarchy from the manager's email
    const hierarchy = buildHierarchy(managerEmail);
    
    // Check if the root node returned a circular error
    if (hierarchy && hierarchy.circularError) {
        throw new Error("Critical Error: Circular reporting line detected at the top level.");
    }

    return hierarchy;

  } catch (err) {
    Logger.log("webGetManagerHierarchy Error: " + err.message);
    throw new Error(err.message);
  }
}

// === NEW: Web App API to get all reports (flat list) ===
function webGetAllSubordinateEmails(managerEmail) {
    try {
        const ss = getSpreadsheet();
        const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
        const userData = getUserDataFromDb(dbSheet);
        
        const managerRole = userData.emailToRole[managerEmail] || 'agent';
        if (managerRole === 'agent') {
            throw new Error("Permission denied.");
        }
        
        // --- Build the direct reporting map ---
        const reportsMap = {};
        userData.userList.forEach(user => {
            const supervisorEmail = user.supervisor;
            if (supervisorEmail) {
                if (!reportsMap[supervisorEmail]) {
                    reportsMap[supervisorEmail] = [];
                }
                reportsMap[supervisorEmail].push(user.email);
            }
        });
        
        const allSubordinates = new Set();
        const queue = [managerEmail];
        
        // Use a set to track users we've already processed (including the manager him/herself)
        const processed = new Set();
        
        while (queue.length > 0) {
            const currentEmail = queue.shift();
            
            // Check for processing loop (shouldn't happen in BFS, but safe check)
            if (processed.has(currentEmail)) continue;
            processed.add(currentEmail);

            const directReports = reportsMap[currentEmail] || [];
            
            directReports.forEach(reportEmail => {
                if (!allSubordinates.has(reportEmail)) {
                    allSubordinates.add(reportEmail);
                    // If the report is a manager, add them to the queue to find their reports
                    if (userData.emailToRole[reportEmail] !== 'agent') {
                        queue.push(reportEmail); // <-- FIX: Was 'push(reportEmail)'
                    }
                }
            
        });
        }
        
        // Return all subordinates *plus* the manager
        allSubordinates.add(managerEmail);
        return Array.from(allSubordinates);

    } catch (err) {
        Logger.log("webGetAllSubordinateEmails Error: " + err.message);
        return [];
    }
}
// --- END OF WEB APP API SECTION ---


function getUserInfo() { 
  try {
    const userEmail = Session.getActiveUser().getEmail().toLowerCase();
    const ss = getSpreadsheet();
    const timeZone = Session.getScriptTimeZone(); 
    const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.employeesCore);
    let userData = getUserDataFromDb(ss);
    let isNewUser = false; 
    const KONECTA_DOMAIN = "@konecta.com"; 
    
    // FIX: Robust check to prevent duplicates on refresh
    let userExists = userData.emailToName[userEmail] !== undefined;

    if (!userExists && userEmail.endsWith(KONECTA_DOMAIN)) {
      const emailColumn = dbSheet.getRange("C:C").getValues();
      for (let i = 0; i < emailColumn.length; i++) {
        if (String(emailColumn[i][0]).trim().toLowerCase() === userEmail) {
          userExists = true;
          break;
        }
      }
    }

    if (!userExists && userEmail.endsWith(KONECTA_DOMAIN)) {
      isNewUser = true;
      const nameParts = userEmail.split('@')[0].split('.');
      const firstName = nameParts[0] ? nameParts[0].charAt(0).toUpperCase() + nameParts[0].slice(1) : '';
      const lastName = nameParts[1] ? nameParts[1].charAt(0).toUpperCase() + nameParts[1].slice(1) : '';
      const newName = [firstName, lastName].join(' ').trim();
      const newEmpID = "KOM-PENDING-" + new Date().getTime();
      dbSheet.appendRow([newEmpID, newName || userEmail, userEmail, 'agent', 'Pending', "", "", 0, 0, 0, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "Pending"]);
      SpreadsheetApp.flush(); 
      userData = getUserDataFromDb(ss);
    }
    
    const accountStatus = userData.emailToAccountStatus[userEmail] || 'Pending';
    const userName = userData.emailToName[userEmail] || "";
    const role = userData.emailToRole[userEmail] || 'agent';
    
    let currentStatus = null;
    if (accountStatus === 'Active') {
      const now = new Date();
      const shiftDate = getShiftDate(now, SHIFT_CUTOFF_HOUR);
      const formattedDate = Utilities.formatDate(shiftDate, timeZone, "MM/dd/yyyy");
      currentStatus = getLatestPunchStatus(userEmail, userName, shiftDate, formattedDate);
    }

    let allUsers = [];
    let allAdmins = [];
    if (role !== 'agent' || isNewUser || accountStatus === 'Pending') { 
      allUsers = userData.userList;
    }
    // FIX: Add 'project_manager' to the list of admins/managers
    allAdmins = userData.userList.filter(u => u.role === 'admin' || u.role === 'superadmin' || u.role === 'manager' || u.role === 'project_manager');
    
    const myBalances = userData.emailToBalances[userEmail] || { annual: 0, sick: 0, casual: 0 };
    let hasPendingRoleRequests = false;
    if (role === 'superadmin') {
      const reqSheet = getOrCreateSheet(ss, SHEET_NAMES.roleRequests);
      const data = reqSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) { if (data[i][7] === 'Pending') { hasPendingRoleRequests = true; break; } }
    }

    const rbacMap = getPermissionsMap(ss);
    const myPermissions = [];
    for (const [perm, roles] of Object.entries(rbacMap)) {
      if (roles[role]) myPermissions.push(perm);
    }

    const breakRules = {
      break1: getBreakConfig("First Break").default,
      lunch: getBreakConfig("Lunch").default,
      break2: getBreakConfig("Last Break").default,
      otPre: getBreakConfig("Overtime Pre-Shift").default,
      otPost: getBreakConfig("Overtime Post-Shift").default
    };

    return {
      name: userName, 
      email: userEmail,
      role: role,
      allUsers: allUsers,
      allAdmins: allAdmins,
      myBalances: myBalances,
      isNewUser: isNewUser, 
      accountStatus: accountStatus, 
      hasPendingRoleRequests: hasPendingRoleRequests, 
      currentStatus: currentStatus,
      permissions: myPermissions,
      breakRules: breakRules 
    };
  } catch (e) { throw new Error("Failed in getUserInfo: " + e.message); }
}


// ================= PUNCH MAIN FUNCTION (ROBUST WFM LOGIC) =================
function punch(action, targetUserName, puncherEmail, adminTimestamp) { 
  const ss = getSpreadsheet();
  const adherenceSheet = getOrCreateSheet(ss, SHEET_NAMES.adherence);
  const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
  const logsSheet = getOrCreateSheet(ss, SHEET_NAMES.logs);
  const otherCodesSheet = getOrCreateSheet(ss, SHEET_NAMES.otherCodes);
  const timeZone = Session.getScriptTimeZone(); 

  const userData = getUserDataFromDb(dbSheet);
  const puncherRole = userData.emailToRole[puncherEmail] || 'agent';
  const puncherIsAdmin = (puncherRole === 'admin' || puncherRole === 'superadmin');
  
  const userName = targetUserName; 
  const userEmail = userData.nameToEmail[userName];

  // 1. Validation: User Existence & Permission
  if (!puncherIsAdmin && puncherEmail !== userEmail) { 
    throw new Error("Permission denied. You can only submit punches for yourself.");
  }
  if (!userEmail) throw new Error(`User "${userName}" not found in Data Base.`);

  // 2. Setup Time & Date Context
  const nowTimestamp = adminTimestamp ? new Date(adminTimestamp) : new Date();
  const shiftDate = getShiftDate(new Date(nowTimestamp), SHIFT_CUTOFF_HOUR);
  const formattedDate = Utilities.formatDate(shiftDate, timeZone, "MM/dd/yyyy");

  // 3. Get Current State (Critical for Logic)
  const row = findOrCreateRow(adherenceSheet, userName, shiftDate, formattedDate);
  // Fetch current state from Column Y (25) - LastAction
  const lastAction = adherenceSheet.getRange(row, 25).getValue() || "Logged Out";

  // 4. LOGIC ENGINE
  
  // --- A. LOGIN LOGIC ---
  if (action === "Login") {
    // Check 1: Already Logged In?
    const existingLogin = adherenceSheet.getRange(row, 3).getValue();
    if (existingLogin) throw new Error("You have already logged in today. Duplicate login is not allowed.");
    
    // Check 2: 4-Hour Schedule Lock (Admins bypass this)
    if (!puncherIsAdmin) {
       validateScheduleLock(userEmail, nowTimestamp);
    }

    // --- TARDY CALCULATION FIX ---
    const sched = getScheduleForDate(userEmail, shiftDate);
    if (sched && sched.start) {
      const schedStart = new Date(sched.start);
      // Calculate difference in seconds
      const diffSec = (nowTimestamp.getTime() - schedStart.getTime()) / 1000;
      
      // Write Tardy if positive (Late), else 0
      // Tardy is Column K (Index 11)
      adherenceSheet.getRange(row, 11).setValue(diffSec > 0 ? diffSec : 0);
      
      // Also write Pre-Shift Overtime if negative (Early)
      // Pre-Shift OT is Column X (Index 24)
      if (diffSec < 0) {
         const earlySec = Math.abs(diffSec);
         const threshold = getBreakConfig("Overtime Pre-Shift").default || 300;
         if (earlySec > threshold) {
             adherenceSheet.getRange(row, 24).setValue(earlySec);
         }
      }
    }
    // -----------------------------

    // Execution
    adherenceSheet.getRange(row, 3).setValue(nowTimestamp); // Login Time
    adherenceSheet.getRange(row, 14).setValue("Present"); // Leave Type
    updateState(adherenceSheet, row, "Login", nowTimestamp);
    logsSheet.appendRow([new Date(), userName, userEmail, action, nowTimestamp]);
    return `Welcome ${userName}. You are successfully Logged In.`;
  }

  // --- B. PRE-REQUISITE CHECK (Must be logged in to do anything else) ---
  const loginTime = adherenceSheet.getRange(row, 3).getValue();
  if (!loginTime) throw new Error("You must punch 'Login' before performing any other action.");

  // --- C. LOGOUT LOGIC ---
  if (action === "Logout") {
    // Check: Must be in "Login" state (Working) to logout. Cannot logout from Break/AUX.
    if (lastAction !== "Login" && !lastAction.endsWith("Out")) {
       // Allow logout if last action was an "Out" (e.g. Lunch Out -> Logout), or just Login.
       // But if last action was "In" (e.g. Lunch In), block it.
       if (lastAction.endsWith("In") && lastAction !== "Login") {
         throw new Error(`You are currently status: "${lastAction}". You must end that activity before Logging Out.`);
       }
    }
    
    // Check: Already logged out?
    const existingLogout = adherenceSheet.getRange(row, 10).getValue();
    if (existingLogout) throw new Error("You have already logged out today.");

    // Execution
    adherenceSheet.getRange(row, 10).setValue(nowTimestamp);
    updateState(adherenceSheet, row, "Logout", nowTimestamp);
    
    // Calculate final metrics
    calculateEndShiftMetrics(adherenceSheet, row, userEmail, nowTimestamp);
    
    logsSheet.appendRow([new Date(), userName, userEmail, action, nowTimestamp]);
    return `Goodbye ${userName}. Shift ended.`;
  }

  // --- D. HANDLING AUX CODES (Meeting, Personal, Coaching, System Down) ---
  if (action.includes("Meeting") || action.includes("Personal") || action.includes("Coaching") || action.includes("System Down")) {
      return processAuxCode(otherCodesSheet, adherenceSheet, row, userName, userEmail, action, nowTimestamp, lastAction, puncherIsAdmin ? puncherEmail : null);
  }

  // --- E. HANDLING MAIN BREAKS (1st Break, Lunch, Last Break) ---
  const breakCols = { 
    "First Break In": 4, "First Break Out": 5, 
    "Lunch In": 6, "Lunch Out": 7, 
    "Last Break In": 8, "Last Break Out": 9 
  };

  if (breakCols[action]) {
      const colIndex = breakCols[action];
      const isBreakIn = action.endsWith("In");
      const currentVal = adherenceSheet.getRange(row, colIndex).getValue();

      // Rule: Once per day check
      if (currentVal) throw new Error(`Error: "${action}" has already been used today.`);

      if (isBreakIn) {
          // Rule: To go on Break, you must be working (Login or returned from previous break)
          // You cannot go Break In if you are currently on Coaching In or Lunch In
          if (lastAction.endsWith("In") && lastAction !== "Login") {
             throw new Error(`Cannot switch to ${action} while you are currently "${lastAction}". Finish that first.`);
          }
          
          // Execution
          adherenceSheet.getRange(row, colIndex).setValue(nowTimestamp);
          updateState(adherenceSheet, row, action, nowTimestamp);
          
          // Check Schedule Window (Compliance)
          checkBreakWindowCompliance(adherenceSheet, row, userEmail, action, nowTimestamp);
          return `${action} recorded. Enjoy your break.`;

      } else {
          // Rule: To go Break Out, you MUST be in that specific Break In state
          const expectedIn = action.replace(" Out", " In");
          if (lastAction !== expectedIn) {
             throw new Error(`Invalid Action. You are not currently in "${expectedIn}". Current status: ${lastAction}`);
          }

          // Execution
          adherenceSheet.getRange(row, colIndex).setValue(nowTimestamp);
          updateState(adherenceSheet, row, action, nowTimestamp); // State becomes "First Break Out" (effectively working)
          
          // Calculate Duration Logic
          const inTime = adherenceSheet.getRange(row, colIndex - 1).getValue();
          if (inTime) {
             const duration = timeDiffInSeconds(inTime, nowTimestamp);
             // Log logic for exceed is handled in daily calc or here
             const typeBase = action.replace(" Out", ""); // "First Break"
             const allowed = getBreakConfig(typeBase).default;
             const diff = duration - allowed;
             // Map exceed columns: 1st=17, Lunch=18, Last=19
             const exceedCol = (action === "First Break Out") ? 17 : (action === "Lunch Out") ? 18 : 19;
             adherenceSheet.getRange(row, exceedCol).setValue(diff > 0 ? diff : 0);
          }
          
          return `${action} recorded. Welcome back.`;
      }
  }

  throw new Error("Unknown punch action.");
}


// --- HELPER: Process AUX Codes (Fixed for Multi-word codes like "System Down") ---
function processAuxCode(auxSheet, mainSheet, mainRow, userName, userEmail, action, now, lastAction, adminEmail) {
    // FIX: Handle codes with spaces (e.g., "System Down In")
    let type = action.endsWith(" In") ? "In" : (action.endsWith(" Out") ? "Out" : "");
    if (!type) throw new Error("Invalid AUX action format.");
    
    // Extract code name by removing the " In" or " Out" suffix
    let codeName = action.substring(0, action.lastIndexOf(" " + type));
    
    if (type === "In") {
        // Validation: Cannot go AUX In if already in another In state (except Login)
        // Note: We allow switching if we are just "Logged In" (working state)
        if (lastAction.endsWith("In") && lastAction !== "Login") {
            throw new Error(`Cannot start ${codeName} while currently status: "${lastAction}". End current activity first.`);
        }
        
        // Log to Other Codes Sheet
        auxSheet.appendRow([now, userName, codeName, now, "", "", adminEmail || ""]);
        
        // Update Main State
        updateState(mainSheet, mainRow, action, now);
        
        return `${action} started.`;
    } 
    else if (type === "Out") {
        // Validation: Must be in the specific In state
        const expectedIn = `${codeName} In`;
        if (lastAction !== expectedIn) {
            throw new Error(`Cannot punch ${action}. You are not currently in "${expectedIn}".`);
        }

        // Find the open session in Other Codes sheet to close it
        const data = auxSheet.getDataRange().getValues();
        let foundRow = -1;
        
        // Search backwards for the last "In" for this user and code that has no "Out"
        for (let i = data.length - 1; i > 0; i--) {
            // Col A=Date, B=Name, C=Code, D=In, E=Out
            // We check if Code matches and "Out" (Col E/index 4) is empty
            if (data[i][1] === userName && data[i][2] === codeName && data[i][3] && !data[i][4]) {
                foundRow = i + 1;
                break;
            }
        }

        if (foundRow > 0) {
            const inTime = data[foundRow-1][3]; // Date Obj
            const duration = timeDiffInSeconds(inTime, now);
            auxSheet.getRange(foundRow, 5).setValue(now); // Set Out Time
            auxSheet.getRange(foundRow, 6).setValue(duration); // Set Duration
            
            updateState(mainSheet, mainRow, action, now);
            return `${action} recorded. Duration: ${Math.round(duration/60)} mins.`;
        } else {
            // Fallback if data sync issue, just update state
            updateState(mainSheet, mainRow, action, now);
            return `${action} recorded (Warning: matching start time not found).`;
        }
    }
}

// --- HELPER: Update State in Adherence Tracker ---
function updateState(sheet, row, action, time) {
    sheet.getRange(row, 25).setValue(action); // Col Y: LastAction
    sheet.getRange(row, 26).setValue(time);   // Col Z: Timestamp
}

// --- HELPER: 4-Hour Lateness Lock ---
function validateScheduleLock(userEmail, now) {
    const schedule = getScheduleForDate(userEmail, now);
    
    // If no schedule exists, we usually allow login (or you can block it by throwing error here)
    if (!schedule || !schedule.start) return; 

    const schedStart = new Date(schedule.start);
    const diffMs = now - schedStart;
    const diffMinutes = diffMs / 60000;
    const diffHours = diffMs / (1000 * 60 * 60);

    // 1. Block Early Login (> 15 mins before start)
    // If Pre-Shift OT is approved, 'schedStart' is already moved earlier, so this adapts automatically.
    if (diffMinutes < -15) {
         const timeString = Utilities.formatDate(schedStart, Session.getScriptTimeZone(), "HH:mm");
         throw new Error(`Login Blocked: Too early. You can only log in 15 minutes before your shift start (${timeString}).`);
    }

    // 2. Block Late Login (> 4 hours after start)
    if (diffHours > 4) {
        throw new Error(`Login Blocked: You are ${diffHours.toFixed(1)} hours late. The cutoff is 4 hours.`);
    }
}

// --- HELPER: End Shift Metrics (Overtime/Early Leave) ---
function calculateEndShiftMetrics(sheet, row, userEmail, now) {
    const schedule = getScheduleForDate(userEmail, now);
    if (!schedule || !schedule.end) return;

    const schedEnd = new Date(schedule.end);
    const diffSec = (now - schedEnd) / 1000; // Positive = Overtime, Negative = Early

    if (diffSec > 0) {
        // Overtime Logic
        const threshold = getBreakConfig("Overtime Post-Shift").default || 300; // 5 mins
        if (diffSec > threshold) {
            sheet.getRange(row, 12).setValue(diffSec); // Overtime Col
        } else {
            sheet.getRange(row, 12).setValue(0);
        }
        sheet.getRange(row, 13).setValue(0); // Early Leave is 0
    } else {
        // Early Leave Logic
        sheet.getRange(row, 12).setValue(0); // Overtime is 0
        sheet.getRange(row, 13).setValue(Math.abs(diffSec)); // Early Leave Col
    }
    
    // Calculate Net Hours
    // (Logic extracted to reuse)
    // ... trigger net calc ...
}

function checkBreakWindowCompliance(sheet, row, userEmail, action, now) {
    // Fetches schedule, checks if now is within break windows, updates Col V (BreakWindowViolation)
    // Existing logic in previous punch function was fine, just ensure it's called here.
    const scheduleData = getScheduleForDate(userEmail, now);
    // ... implementation ...
}


// REPLACE this function
// ================= SCHEDULE RANGE SUBMIT FUNCTION =================
function submitScheduleRange(puncherEmail, userEmail, userName, startDateStr, endDateStr, startTime, endTime, leaveType) {
  const ss = getSpreadsheet();
const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
  const userData = getUserDataFromDb(dbSheet);
  const puncherRole = userData.emailToRole[puncherEmail] || 'agent';
  const timeZone = Session.getScriptTimeZone();
if (puncherRole !== 'admin' && puncherRole !== 'superadmin') {
    throw new Error("Permission denied. Only admins can submit schedules.");
}
  
  const scheduleSheet = getOrCreateSheet(ss, SHEET_NAMES.schedule);
  const scheduleData = scheduleSheet.getDataRange().getValues();
  const logsSheet = getOrCreateSheet(ss, SHEET_NAMES.logs);
const userScheduleMap = {};
  for (let i = 1; i < scheduleData.length; i++) {
    // *** MODIFIED: Read Email from Col G (index 6) ***
    const rowEmail = scheduleData[i][6];
// *** MODIFIED: Read Date from Col B (index 1) ***
    const rowDateRaw = scheduleData[i][1];
if (rowEmail && rowDateRaw && rowEmail.toLowerCase() === userEmail) {
      const rowDate = new Date(rowDateRaw);
const rowDateStr = Utilities.formatDate(rowDate, timeZone, "MM/dd/yyyy");
      userScheduleMap[rowDateStr] = i + 1;
}
  }
  
  const startDate = new Date(startDateStr);
  const endDate = new Date(endDateStr);
let currentDate = new Date(startDate);
  let daysProcessed = 0;
  let daysUpdated = 0;
  let daysCreated = 0;
const oneDayInMs = 24 * 60 * 60 * 1000;
  
  currentDate = new Date(currentDate.valueOf() + currentDate.getTimezoneOffset() * 60000);
const finalDate = new Date(endDate.valueOf() + endDate.getTimezoneOffset() * 60000);
  
  while (currentDate <= finalDate) {
    const currentDateStr = Utilities.formatDate(currentDate, timeZone, "MM/dd/yyyy");
// *** NEW: Auto-calculate shift end date for overnight shifts ***
    let shiftEndDate = new Date(currentDate);
// Start with the same date
    if (startTime && endTime) {
      const startDateTime = createDateTime(currentDate, startTime);
const endDateTime = createDateTime(currentDate, endTime);
      if (endDateTime <= startDateTime) {
        shiftEndDate.setDate(shiftEndDate.getDate() + 1);
// It's the next day
      }
    }
    // *** END NEW ***

    const result = updateOrAddSingleSchedule(
      scheduleSheet, userScheduleMap, logsSheet,
      userEmail, userName, 
      currentDate, // This is StartDate (Col B)
      shiftEndDate, // *** NEW: This is EndDate (Col D) ***
      currentDateStr, 
      startTime, endTime, leaveType, puncherEmail
    );
if (result === "UPDATED") daysUpdated++;
    if (result === "CREATED") daysCreated++;
    
    daysProcessed++;
    currentDate.setTime(currentDate.getTime() + oneDayInMs);
}
  
  if (daysProcessed === 0) {
    throw new Error("No dates were processed. Check date range.");
}
  
  return `Schedule submission complete for ${userName}. Days processed: ${daysProcessed} (Updated: ${daysUpdated}, Created: ${daysCreated}).`;
}

// Helper for Import & Manual Submit (PHASE 9 UPDATED)
function updateOrAddSingleSchedule(
  scheduleSheet, userScheduleMap, logsSheet, 
  userEmail, userName, shiftStartDate, shiftEndDate, targetDateStr, 
  startTime, endTime, leaveType, puncherEmail,
  // New Optional Args
  b1s = "", b1e = "", ls = "", le = "", b2s = "", b2e = ""
) {
  
  const existingRow = userScheduleMap[targetDateStr];
  let startTimeObj = startTime ? new Date(`1899-12-30T${startTime}`) : "";
  let endTimeObj = endTime ? new Date(`1899-12-30T${endTime}`) : "";
  let endDateObj = (leaveType === 'Present' && endTimeObj) ? shiftEndDate : "";

  // Convert break strings to Date objects if they exist
  const toDateObj = (t) => t ? new Date(`1899-12-30T${t}`) : "";
  
  // --- PHASE 9: Write 13 Columns ---
  const rowData = [[
    userName,       // A
    shiftStartDate, // B
    startTimeObj,   // C
    endDateObj,     // D
    endTimeObj,     // E
    leaveType,      // F
    userEmail,      // G
    toDateObj(b1s), // H (Break1 Start)
    toDateObj(b1e), // I (Break1 End)
    toDateObj(ls),  // J (Lunch Start)
    toDateObj(le),  // K (Lunch End)
    toDateObj(b2s), // L (Break2 Start)
    toDateObj(b2e)  // M (Break2 End)
  ]];

  if (existingRow) {
    scheduleSheet.getRange(existingRow, 1, 1, 13).setValues(rowData);
    logsSheet.appendRow([new Date(), userName, puncherEmail, "Schedule UPDATE", `Set to: ${leaveType}`]);
    return "UPDATED";
  } else {
    scheduleSheet.appendRow(rowData[0]);
    logsSheet.appendRow([new Date(), userName, puncherEmail, "Schedule CREATE", `Set to: ${leaveType}`]);
    return "CREATED";
  }
}

// ================= HELPER FUNCTIONS =================

function getShiftDate(dateObj, cutoffHour) {
  if (dateObj.getHours() < cutoffHour) {
    dateObj.setDate(dateObj.getDate() - 1);
  }
  return dateObj;
}

function createDateTime(dateObj, timeStr) {
  if (!timeStr) return null;
  const parts = timeStr.split(':');
  if (parts.length < 2) return null;
  
  const [hours, minutes, seconds] = parts.map(Number);
  if (isNaN(hours) || isNaN(minutes)) return null; 

  const newDate = new Date(dateObj);
  newDate.setHours(hours, minutes, seconds || 0, 0);
  return newDate;
}

// [code.gs] REPLACE your existing getUserDataFromDb with this:

function getUserDataFromDb(ss) {
  if (!ss || !ss.getSheetByName) ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const coreSheet = getOrCreateSheet(ss, SHEET_NAMES.employeesCore);
  const piiSheet = getOrCreateSheet(ss, SHEET_NAMES.employeesPII); 

  const coreData = coreSheet.getDataRange().getValues();
  const piiData = piiSheet.getDataRange().getValues();
  const piiMap = {};
  for (let i = 1; i < piiData.length; i++) {
    const empID = piiData[i][0];
    piiMap[empID] = { hiringDate: piiData[i][1] };
  }

  const nameToEmail = {};
  const emailToName = {};
  const emailToRole = {};
  const emailToBalances = {};
  const emailToRow = {};
  const emailToSupervisor = {};
  const emailToProjectManager = {}; 
  const emailToAccountStatus = {};
  const emailToHiringDate = {};
  const userList = [];

  // Map Headers Dynamically
  const headers = coreData[0];
  const colIdx = {};
  headers.forEach((header, index) => { colIdx[header] = index; });

  const defaultDirectMgrIdx = 5; // Fallback to Col F (Index 5) if headers fail

  for (let i = 1; i < coreData.length; i++) {
    try {
      const row = coreData[i];
      const empID = row[colIdx["EmployeeID"] || 0];
      const name = row[colIdx["Name"] || 1];
      const email = row[colIdx["Email"] || 2];

      if (name && email) {
        const cleanName = name.toString().trim();
        const cleanEmail = email.toString().trim().toLowerCase();
        const userRole = (row[colIdx["Role"] || 3] || 'agent').toString().trim().toLowerCase();
        const accountStatus = (row[colIdx["AccountStatus"] || 4] || "Pending").toString().trim();

        // --- MANAGER FETCHING (FIXED) ---
        let dmIdx = colIdx["DirectManagerEmail"];
        if (dmIdx === undefined) dmIdx = colIdx["DirectManager"]; 
        if (dmIdx === undefined) dmIdx = defaultDirectMgrIdx;

        // FIX: Check for "FunctionalManagerEmail" (Col G) explicitly
        let pmIdx = colIdx["ProjectManagerEmail"];
        if (pmIdx === undefined) pmIdx = colIdx["ProjectManager"];
        if (pmIdx === undefined) pmIdx = colIdx["FunctionalManagerEmail"]; 

        let dotIdx = colIdx["DottedManager"];

        const directMgr = (row[dmIdx] || "").toString().trim().toLowerCase();
        // If pmIdx is still undefined, projectMgr will be empty string
        const projectMgr = (pmIdx !== undefined ? row[pmIdx] : "").toString().trim().toLowerCase();
        const dottedMgr = (dotIdx !== undefined ? row[dotIdx] : "").toString().trim().toLowerCase();
        // ---------------------------------------------

        const pii = piiMap[empID] || {};
        const hiringDateStr = convertDateToString(parseDate(pii.hiringDate));

        nameToEmail[cleanName] = cleanEmail;
        emailToName[cleanEmail] = cleanName;
        emailToRole[cleanEmail] = userRole;
        emailToRow[cleanEmail] = i + 1;
        
        emailToSupervisor[cleanEmail] = directMgr;
        emailToProjectManager[cleanEmail] = projectMgr;
        emailToAccountStatus[cleanEmail] = accountStatus;
        emailToHiringDate[cleanEmail] = hiringDateStr;

        emailToBalances[cleanEmail] = {
          annual: parseFloat(row[colIdx["AnnualBalance"] || 7]) || 0,
          sick: parseFloat(row[colIdx["SickBalance"] || 8]) || 0,
          casual: parseFloat(row[colIdx["CasualBalance"] || 9]) || 0
        };

        userList.push({
          empID: empID,
          name: cleanName,
          email: cleanEmail,
          role: userRole,
          balances: emailToBalances[cleanEmail],
          supervisor: directMgr,
          projectManager: projectMgr, // This will now populate correctly
          dottedManager: dottedMgr,
          accountStatus: accountStatus,
          hiringDate: hiringDateStr
        });
      }
    } catch (e) {
      Logger.log(`Error processing user row ${i}: ${e.message}`);
    }
  }

  return {
    nameToEmail, emailToName, emailToRole, emailToBalances,
    emailToRow, emailToSupervisor, emailToProjectManager,
    emailToAccountStatus, emailToHiringDate, userList
  };
}


/**
 * UPDATED PHASE 2: Returns Status + Login Time for Timers
 */
function getLatestPunchStatus(userEmail, userName, shiftDate, formattedDate) {
  const ss = getSpreadsheet();
  const adherenceSheet = getOrCreateSheet(ss, SHEET_NAMES.adherence);
  
  // Find the row
  const adherenceData = adherenceSheet.getDataRange().getValues();
  let rowData = null;
  
  // Find row matching today
  for (let i = adherenceData.length - 1; i > 0; i--) {
    const rowDate = adherenceData[i][0];
    let rowDateStr = (rowDate instanceof Date) ? 
        Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "MM/dd/yyyy") : "";
    
    if (rowDateStr === formattedDate && adherenceData[i][1] === userName) {
      rowData = adherenceData[i];
      break;
    }
  }

  const scheduleInfo = getScheduleForDate(userEmail, shiftDate);

  if (!rowData) {
    return { status: "Logged Out", time: null, loginTime: null, schedule: scheduleInfo };
  }

  // Col C (index 2) is Login Time
  const loginTime = rowData[2] ? new Date(rowData[2]) : null;
  
  // Col Y (index 24) is LastAction, Col Z (index 25) is Timestamp
  // These were populated by our new updateState() helper
  const lastAction = rowData[24];
  const lastActionTime = rowData[25] ? new Date(rowData[25]) : null;

  // Determine Display Status
  let displayStatus = "Logged Out";
  
  if (lastAction === "Login" || (lastAction && lastAction.endsWith("Out") && lastAction !== "Logout")) {
      displayStatus = "Logged In";
  } else if (lastAction === "Logout") {
      displayStatus = "Logged Out";
  } else if (lastAction) {
      // "Meeting In", "First Break In", "Coaching In"
      displayStatus = lastAction; // Pass raw "In" status
  }

  return {
    status: displayStatus,
    time: convertDateToString(lastActionTime),
    loginTime: convertDateToString(loginTime),
    schedule: scheduleInfo
  };
}

/**
 * UPDATED PHASE 1: Helper to fetch schedule start/end for a specific date.
 * Handles overnight shifts logic correctly.
 */
function getScheduleForDate(userEmail, dateObj) {
  const ss = getSpreadsheet();
  const sheet = getOrCreateSheet(ss, SHEET_NAMES.schedule);
  const data = sheet.getDataRange().getValues();
  const timeZone = Session.getScriptTimeZone();
  const targetDateStr = Utilities.formatDate(dateObj, timeZone, "MM/dd/yyyy");
  
  // Iterate backwards to find the most recent matching schedule entry
  for (let i = data.length - 1; i > 0; i--) {
    // Col 7 (Index 6) is email, Col 2 (Index 1) is Date
    if (String(data[i][6]).toLowerCase() === userEmail.toLowerCase()) {
      const rowDate = data[i][1];
      
      // Check if this row matches our target date
      // Note: parseDate is robust, but direct comparison of strings is safer for exact dates
      let rowDateStr = "";
      if (rowDate instanceof Date) {
        rowDateStr = Utilities.formatDate(rowDate, timeZone, "MM/dd/yyyy");
      } else {
        // Try parsing if string
        const pDate = parseDate(rowDate);
        if (pDate) rowDateStr = Utilities.formatDate(pDate, timeZone, "MM/dd/yyyy");
      }

      if (rowDateStr === targetDateStr) {
        let startTime = data[i][2]; // Col C
        let endTime = data[i][4];   // Col E
        
        // Construct full DateTime objects
        let startDateTime = null;
        let endDateTime = null;

        if (startTime) {
           // Handle if time is already a Date object (from Sheets) or string
           const timeStr = (startTime instanceof Date) ? 
             Utilities.formatDate(startTime, timeZone, "HH:mm:ss") : startTime;
           startDateTime = createDateTime(dateObj, timeStr);
        }

        if (endTime) {
           const timeStr = (endTime instanceof Date) ? 
             Utilities.formatDate(endTime, timeZone, "HH:mm:ss") : endTime;
           
           // Base end date is the same day
           let baseEndDate = new Date(dateObj);
           endDateTime = createDateTime(baseEndDate, timeStr);
           
           // Overnight check: If End Time is earlier than Start Time, it ends the next day
           // Or if explicit EndDate (Col D) is different (not handled here for simplicity, relying on time logic)
           if (startDateTime && endDateTime && endDateTime < startDateTime) {
             endDateTime.setDate(endDateTime.getDate() + 1);
           }
        }

        return {
          start: convertDateToString(startDateTime),
          end: convertDateToString(endDateTime)
        };
      }
    }
  }
  return null;
}

/**
 * NEW PHASE 3: Reads break configuration from the sheet.
 * Returns an object with default and max duration in seconds.
 */
function getBreakConfig(breakType, projectId) {
  const ss = getSpreadsheet();
  const sheet = getOrCreateSheet(ss, SHEET_NAMES.breakConfig);
  const data = sheet.getDataRange().getValues();
  
  // Default fallbacks if sheet is empty or row missing
  let config = { default: 900, max: 1200 }; // 15 min / 20 min default
  if (breakType === "Lunch") config = { default: 1800, max: 2400 }; // 30 min / 40 min
  
  for (let i = 1; i < data.length; i++) {
    // Col A: Type, Col B: Default, Col C: Max, Col D: Project
    const rowType = data[i][0];
    const rowProject = data[i][3] || "ALL";
    
    if (rowType === breakType) {
      // Simplistic logic: specific project overrides ALL, but here we just take the first match or 'ALL'
      // For Phase 3, we assume global rules (Project = ALL)
      config.default = Number(data[i][1]);
      config.max = Number(data[i][2]);
      break;
    }
  }
  return config;
}


// (No Change)
function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

// (No Change)
function findOrCreateRow(sheet, userName, shiftDate, formattedDate) { 
  const data = sheet.getDataRange().getValues();
  const timeZone = Session.getScriptTimeZone();
  let row = -1;
  for (let i = 1; i < data.length; i++) {
    const rowDate = new Date(data[i][0]);
    const rowUser = data[i][1]; 
    if (
      rowUser && 
      rowUser.toString().toLowerCase() === userName.toLowerCase() && 
      Utilities.formatDate(rowDate, timeZone, "MM/dd/yyyy") === formattedDate
    ) {
      row = i + 1;
      break;
    }
  }

  if (row === -1) {
    row = sheet.getLastRow() + 1;
    sheet.getRange(row, 1).setValue(shiftDate);
    sheet.getRange(row, 2).setValue(userName); 
  }
  return row;
}

function getOrCreateSheet(ss, name) {
  if (!name) return null;
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    
    if (name === SHEET_NAMES.employeesCore) {
      sheet.getRange("A1:J1").setValues([["EmployeeID", "Name", "Email", "Role", "AccountStatus", "DirectManagerEmail", "FunctionalManagerEmail", "AnnualBalance", "SickBalance", "CasualBalance"]]);
      sheet.setFrozenRows(1);
    } 
    else if (name === SHEET_NAMES.employeesPII) {
      sheet.getRange("A1:H1").setValues([["EmployeeID", "HiringDate", "Salary", "IBAN", "Address", "Phone", "MedicalInfo", "ContractType"]]);
      sheet.getRange("B:B").setNumberFormat("yyyy-mm-dd");
      sheet.setFrozenRows(1);
    }
    // --- PHASE 8 UPDATE: Added Overtime Rules ---
    else if (name === SHEET_NAMES.breakConfig) {
      sheet.getRange("A1:D1").setValues([["BreakType", "DefaultDuration (Sec)", "MaxDuration (Sec)", "ProjectID"]]);
      sheet.getRange("A2:D6").setValues([
        ["First Break", 900, 1200, "ALL"], 
        ["Lunch", 1800, 2400, "ALL"],      
        ["Last Break", 900, 1200, "ALL"],
        ["Overtime Pre-Shift", 300, 0, "ALL"],  // 5 mins threshold
        ["Overtime Post-Shift", 300, 0, "ALL"]  // 5 mins threshold
      ]);
      sheet.setFrozenRows(1);
    }
    else if (name === SHEET_NAMES.assets) {
      sheet.getRange("A1:E1").setValues([["AssetID", "Type", "AssignedTo_EmployeeID", "DateAssigned", "Status"]]);
      sheet.setFrozenRows(1);
    }
    else if (name === SHEET_NAMES.projects) {
      sheet.getRange("A1:D1").setValues([["ProjectID", "ProjectName", "ProjectManagerEmail", "AllowedRoles"]]);
      sheet.setFrozenRows(1);
    }
    else if (name === SHEET_NAMES.projectLogs) {
      sheet.getRange("A1:E1").setValues([["LogID", "EmployeeID", "ProjectID", "Date", "HoursLogged"]]);
      sheet.setFrozenRows(1);
    }
    else if (name === SHEET_NAMES.warnings) {
      sheet.getRange("A1:H1").setValues([["WarningID", "EmployeeID", "Type", "Level", "Date", "Description", "Status", "IssuedBy"]]);
      sheet.setFrozenRows(1);
    }
    else if (name === SHEET_NAMES.schedule) {
      sheet.getRange("A1:M1").setValues([["Name", "StartDate", "ShiftStartTime", "EndDate", "ShiftEndTime", "LeaveType", "agent email", "Break1_Start", "Break1_End", "Lunch_Start", "Lunch_End", "Break2_Start", "Break2_End"]]);
      sheet.getRange("B:B").setNumberFormat("mm/dd/yyyy");
      sheet.getRange("C:C").setNumberFormat("hh:mm");
      sheet.getRange("D:D").setNumberFormat("mm/dd/yyyy");
      sheet.getRange("E:E").setNumberFormat("hh:mm");
      sheet.getRange("H:M").setNumberFormat("hh:mm");
    } 
    // --- UPDATED: Added LastAction and LastActionTimestamp (Cols Y & Z) ---
    else if (name === SHEET_NAMES.adherence) {
      sheet.getRange("A1:Z1").setValues([["Date", "User Name", "Login", "First Break In", "First Break Out", "Lunch In", "Lunch Out", "Last Break In", "Last Break Out", "Logout", "Tardy (Seconds)", "Overtime (Seconds)", "Early Leave (Seconds)", "Leave Type", "Admin Audit", "—", "1st Break Exceed", "Lunch Exceed", "Last Break Exceed", "Absent", "Admin Code", "BreakWindowViolation", "NetLoginHours", "PreShiftOvertime", "LastAction", "LastActionTimestamp"]]);
      sheet.getRange("C:J").setNumberFormat("hh:mm:ss");
      sheet.getRange("Z:Z").setNumberFormat("hh:mm:ss"); // Format new timestamp col
    } 
    else if (name === SHEET_NAMES.logs) {
      sheet.getRange("A1:E1").setValues([["Timestamp", "User Name", "Email", "Action", "Time"]]);
    } 
    else if (name === SHEET_NAMES.otherCodes) { 
      sheet.getRange("A1:G1").setValues([["Date", "User Name", "Code", "Time In", "Time Out", "Duration (Seconds)", "Admin Audit (Email)"]]);
      sheet.getRange("D:E").setNumberFormat("hh:mm:ss");
    } 
    else if (name === SHEET_NAMES.leaveRequests) { 
      sheet.getRange("A1:N1").setValues([["RequestID", "Status", "RequestedByEmail", "RequestedByName", "LeaveType", "StartDate", "EndDate", "TotalDays", "Reason", "ActionDate", "ActionBy", "SupervisorEmail", "ActionReason", "SickNoteURL"]]);
      sheet.getRange("F:G").setNumberFormat("mm/dd/yyyy");
      sheet.getRange("J:J").setNumberFormat("mm/dd/yyyy");
    } 
    else if (name === SHEET_NAMES.coachingSessions) { 
      sheet.getRange("A1:M1").setValues([["SessionID", "AgentEmail", "AgentName", "CoachEmail", "CoachName", "SessionDate", "WeekNumber", "OverallScore", "FollowUpComment", "SubmissionTimestamp", "FollowUpDate", "FollowUpStatus", "AgentAcknowledgementTimestamp"]]);
      sheet.getRange("F:F").setNumberFormat("mm/dd/yyyy");
      sheet.getRange("J:J").setNumberFormat("mm/dd/yyyy hh:mm:ss");
      sheet.getRange("K:K").setNumberFormat("mm/dd/yyyy");
      sheet.getRange("M:M").setNumberFormat("mm/dd/yyyy hh:mm:ss");
    } 
    else if (name === SHEET_NAMES.coachingScores) { 
      sheet.getRange("A1:E1").setValues([["SessionID", "Category", "Criteria", "Score", "Comment"]]);
    } 
    else if (name === SHEET_NAMES.coachingTemplates) {
      sheet.getRange("A1:D1").setValues([["TemplateName", "Category", "Criteria", "Status"]]);
      sheet.setFrozenRows(1);
    }
    else if (name === SHEET_NAMES.pendingRegistrations) {
      sheet.getRange("A1:J1").setValues([["RequestID", "UserEmail", "UserName", "DirectManagerEmail", "FunctionalManagerEmail", "DirectStatus", "FunctionalStatus", "Address", "Phone", "RequestTimestamp"]]);
      sheet.setFrozenRows(1);
      sheet.getRange("J:J").setNumberFormat("mm/dd/yyyy hh:mm:ss");
    }
    else if (name === SHEET_NAMES.movementRequests) {
      sheet.getRange("A1:J1").setValues([["MovementID", "Status", "UserToMoveEmail", "UserToMoveName", "FromSupervisorEmail", "ToSupervisorEmail", "RequestTimestamp", "ActionTimestamp", "ActionByEmail", "RequestedByEmail"]]);
      sheet.getRange("G:H").setNumberFormat("mm/dd/yyyy hh:mm:ss");
    }
    else if (name === SHEET_NAMES.announcements) {
      sheet.getRange("A1:E1").setValues([["AnnouncementID", "Content", "Status", "CreatedByEmail", "Timestamp"]]);
      sheet.getRange("E:E").setNumberFormat("mm/dd/yyyy hh:mm:ss");
    }
    else if (name === SHEET_NAMES.roleRequests) {
      sheet.getRange("A1:J1").setValues([["RequestID", "UserEmail", "UserName", "CurrentRole", "RequestedRole", "Justification", "RequestTimestamp", "Status", "ActionByEmail", "ActionTimestamp"]]);
      sheet.getRange("G:G").setNumberFormat("mm/dd/yyyy hh:mm:ss");
      sheet.getRange("J:J").setNumberFormat("mm/dd/yyyy hh:mm:ss");
    }
    else if (name === SHEET_NAMES.offboarding) {
      sheet.getRange("A1:O1").setValues([["RequestID", "EmployeeID", "Name", "Email", "Type", "Reason", "Status", "DirectManager", "ProjectManager", "DirectStatus", "ProjectStatus", "HRStatus", "RequestDate", "ExitDate", "InitiatedBy"]]);
      sheet.setFrozenRows(1);
    }
  }

  if (name === SHEET_NAMES.adherence) sheet.getRange("C:J").setNumberFormat("hh:mm:ss");
  if (name === SHEET_NAMES.otherCodes) sheet.getRange("D:E").setNumberFormat("hh:mm:ss");
  if (name === SHEET_NAMES.employeesPII) sheet.getRange("B:B").setNumberFormat("yyyy-mm-dd");
  if (name === SHEET_NAMES.schedule) sheet.getRange("H:M").setNumberFormat("hh:mm");

  return sheet;
}

// (No Change)
function timeDiffInSeconds(start, end) {
  if (!start || !end || !(start instanceof Date) || !(end instanceof Date)) {
    return 0;
  }
  return Math.round((end.getTime() - start.getTime()) / 1000);
}


// ================= DAILY AUTO-LOG FUNCTION =================
function dailyLeaveSweeper() {
  const ss = getSpreadsheet();
  const scheduleSheet = getOrCreateSheet(ss, SHEET_NAMES.schedule);
  const adherenceSheet = getOrCreateSheet(ss, SHEET_NAMES.adherence);
  const logsSheet = getOrCreateSheet(ss, SHEET_NAMES.logs);
  const timeZone = Session.getScriptTimeZone();
  // 1. Define the 7-day lookback period
  const lookbackDays = 7;
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const endDate = new Date(today); // Today
  endDate.setDate(endDate.getDate() - 1); // End date is yesterday

  const startDate = new Date(endDate);
  startDate.setDate(startDate.getDate() - (lookbackDays - 1)); // Start date is 7 days ago

  const startDateStr = Utilities.formatDate(startDate, timeZone, "MM/dd/yyyy");
  const endDateStr = Utilities.formatDate(endDate, timeZone, "MM/dd/yyyy");

  Logger.log(`Starting dailyLeaveSweeper for date range: ${startDateStr} to ${endDateStr}`);
  // 2. Get all Adherence rows for the past 7 days and create a lookup Set
  const allAdherence = adherenceSheet.getDataRange().getValues();
  const adherenceLookup = new Set();
  for (let i = 1; i < allAdherence.length; i++) {
    try {
      const rowDate = new Date(allAdherence[i][0]);
      if (rowDate >= startDate && rowDate <= endDate) {
        const rowDateStr = Utilities.formatDate(rowDate, timeZone, "MM/dd/yyyy");
        const userName = allAdherence[i][1].toString().trim().toLowerCase();
        adherenceLookup.add(`${userName}:${rowDateStr}`);
      }
    } catch (e) {
      Logger.log(`Skipping adherence row ${i+1}: ${e.message}`);
    }
  }
  Logger.log(`Found ${adherenceLookup.size} existing adherence records in the date range.`);
  // 3. Get all Schedules and loop through them
  const allSchedules = scheduleSheet.getDataRange().getValues();
  let missedLogs = 0;
  for (let i = 1; i < allSchedules.length; i++) {
    try {
      // *** THIS LINE IS THE FIX ***
      // It now correctly reads all 7 columns, matching your sheet structure.
      const [schName, schDate, schStart, schEndDate, schEndTime, schLeave, schEmail] = allSchedules[i];
      // *** END OF FIX ***

      const leaveType = (schLeave || "").toString().trim(); // schLeave is now correctly column F (index 5)

      // This logic is now correct because schLeave and schEmail are from the right columns
      if (leaveType === "" || !schName || !schEmail) {
        continue;
      }

     const schDateObj = parseDate(schDate);

      if (schDateObj && schDateObj >= startDate && schDateObj <= endDate) {
        const schDateStr = Utilities.formatDate(schDateObj, timeZone, "MM/dd/yyyy");
        const userName = schName.toString().trim();
        const userNameLower = userName.toLowerCase();

        const lookupKey = `${userNameLower}:${schDateStr}`;
        // 4. Check if this user is *already* in the Adherence sheet
        if (adherenceLookup.has(lookupKey)) {
          continue; // We found them, so skip
        }

        // 5. We found a missed user!
        Logger.log(`Found missed user: ${userName} for ${schDateStr}. Logging: ${leaveType}`);

        const row = findOrCreateRow(adherenceSheet, userName, schDateObj, schDateStr);
        // *** MODIFIED for Request 3: Mark "Present" as "Absent" ***
        if (leaveType.toLowerCase() === "present") {
          adherenceSheet.getRange(row, 14).setValue("Absent"); // Set Leave Type to Absent
          adherenceSheet.getRange(row, 20).setValue("Yes"); // Set Absent flag to Yes (Col T)
          logsSheet.appendRow([new Date(), userName, schEmail, "Auto-Log Absent", "User was 'Present' but did not punch in."]);
        } else {
          adherenceSheet.getRange(row, 14).setValue(leaveType); // Log Sick, Annual, etc.
          if (leaveType.toLowerCase() === "absent") {
            adherenceSheet.getRange(row, 20).setValue("Yes"); // Set Absent flag (Col T)
          }
          logsSheet.appendRow([new Date(), userName, schEmail, "Auto-Log Leave", leaveType]);
        }

        missedLogs++;
        adherenceLookup.add(lookupKey); // Add to lookup so we don't process again
      }
    } catch (e) {
      Logger.log(`Skipping schedule row ${i+1}: ${e.message}`);
    }
  }

  Logger.log(`dailyLeaveSweeper finished. Logged ${missedLogs} missed users.`);
}

// ================= LEAVE REQUEST FUNCTIONS =================

// (Helper - No Change)
function convertDateToString(dateObj) {
  if (dateObj instanceof Date && !isNaN(dateObj)) {
    return dateObj.toISOString(); // "2025-11-06T18:30:00.000Z"
  }
  return null; // Return null if it's not a valid date
}

// (No Change)
function getMyRequests(userEmail) {
  const ss = getSpreadsheet();
  const reqSheet = getOrCreateSheet(ss, SHEET_NAMES.leaveRequests);
  const allData = reqSheet.getDataRange().getValues();
  const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
  const userData = getUserDataFromDb(dbSheet);
  
  const myRequests = [];
  
  // Loop backwards (newest first)
  for (let i = allData.length - 1; i > 0; i--) { 
    const row = allData[i];
    if (String(row[2] || "").trim().toLowerCase() === userEmail) {
      try { 
        const startDate = new Date(row[5]);
        const endDate = new Date(row[6]);
        // Parse numeric ID part if possible, else use today
        const requestedDateNum = row[0].includes('_') ? Number(row[0].split('_')[1]) : new Date().getTime();

        const currentApproverEmail = row[11]; // Col L
        const approverName = userData.emailToName[currentApproverEmail] || currentApproverEmail || "Pending Assignment";

        myRequests.push({
          requestID: row[0],
          status: row[1],
          leaveType: row[4],
          startDate: convertDateToString(startDate),
          endDate: convertDateToString(endDate),
          totalDays: row[7],
          reason: row[8],
          requestedDate: convertDateToString(new Date(requestedDateNum)),
          supervisorName: approverName, // Shows who is holding the request
          actionDate: convertDateToString(new Date(row[9])),
          actionBy: userData.emailToName[row[10]] || row[10],
          actionByReason: row[12] || "",
          sickNoteUrl: row[13] || ""
        });
      } catch (e) {
        Logger.log("Error parsing row " + i);
      }
    }
  }
  return myRequests;
}

function getAdminLeaveRequests(adminEmail, filter) {
  const ss = getSpreadsheet();
  const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
  const userData = getUserDataFromDb(dbSheet);
  const adminRole = userData.emailToRole[adminEmail] || 'agent';

  if (adminRole !== 'admin' && adminRole !== 'superadmin') return { error: "Permission Denied." };

  const reqSheet = getOrCreateSheet(ss, SHEET_NAMES.leaveRequests);
  const allData = reqSheet.getDataRange().getValues();
  const results = [];
  const filterStatus = filter.status.toLowerCase();
  const filterUser = filter.userEmail;

  // Get subordinates for visibility check
  const mySubordinateEmails = new Set(webGetAllSubordinateEmails(adminEmail));

  for (let i = 1; i < allData.length; i++) { 
    const row = allData[i];
    if (!row[0]) continue;

    const requestStatus = (row[1] || "").toString().trim().toLowerCase();
    const requesterEmail = (row[2] || "").toString().trim().toLowerCase();
    const assignedApprover = (row[11] || "").toString().trim().toLowerCase(); 
    
    // 1. Filter by Status
    if (filterStatus !== 'all' && !requestStatus.includes(filterStatus)) continue;

    // 2. Filter by User
    if (filterUser && filterUser !== 'ALL_USERS' && filterUser !== 'ALL_SUBORDINATES' && requesterEmail !== filterUser) continue;

    // 3. Visibility Logic
    let isVisible = false;
    if (adminRole === 'superadmin') {
      isVisible = true;
    } else {
      // Show if assigned to me OR if I am the direct manager/project manager (historical visibility)
      // Note: We check the SNAPSHOT columns (O=14, P=15) if available, else standard check
      const directMgrSnapshot = (row[14] || "").toString().toLowerCase();
      const projectMgrSnapshot = (row[15] || "").toString().toLowerCase();

      if (assignedApprover === adminEmail) isVisible = true;
      else if (directMgrSnapshot === adminEmail) isVisible = true;
      else if (projectMgrSnapshot === adminEmail) isVisible = true;
      else if (mySubordinateEmails.has(requesterEmail)) isVisible = true;
    }

    if (!isVisible) continue;

    try {
        const startDate = new Date(row[5]);
        const endDate = new Date(row[6]);
        const datePart = row[0].split('_')[1];
        const reqDate = datePart ? new Date(Number(datePart)) : new Date();

        results.push({
          requestID: row[0],
          status: row[1],
          requestedByName: row[3],
          leaveType: row[4],
          startDate: convertDateToString(startDate),
          endDate: convertDateToString(endDate),
          totalDays: row[7],
          reason: row[8],
          requestedDate: convertDateToString(reqDate),
          supervisorName: userData.emailToName[assignedApprover] || assignedApprover,
          actionBy: userData.emailToName[row[10]] || row[10],
          actionByReason: row[12],
          requesterBalance: userData.emailToBalances[requesterEmail],
          sickNoteUrl: row[13]
        });
    } catch (e) { }
  }
  return results;
}

function submitLeaveRequest(submitterEmail, request, targetUserEmail) {
  const ss = getSpreadsheet();
  const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
  const userData = getUserDataFromDb(dbSheet);
  
  const requestEmail = (targetUserEmail || submitterEmail).toLowerCase();
  const requestName = userData.emailToName[requestEmail];
  
  if (!requestName) throw new Error(`User account ${requestEmail} not found.`);
  
  // 1. Identify Approvers & Validate
  const directManager = userData.emailToSupervisor[requestEmail];
  const projectManager = userData.emailToProjectManager[requestEmail];

  // CRITICAL FIX: Stop "NA" by blocking submission if Direct Manager is missing
  if (!directManager || directManager === "" || directManager === "na") {
    throw new Error(`Cannot submit request. User ${requestName} does not have a valid Direct Manager assigned in Employees_Core.`);
  }

  // 2. Determine Workflow
  let status = "Pending";
  let assignedApprover = directManager;

  // If Project Manager exists and is different, they approve first
  if (projectManager && projectManager !== "" && projectManager !== directManager) {
    status = "Pending Project Mgr";
    assignedApprover = projectManager;
  } else {
    status = "Pending Direct Mgr"; 
    assignedApprover = directManager;
  }

  // 3. Balance Check
  const startDate = new Date(request.startDate + 'T00:00:00');
  const endDate = request.endDate ? new Date(request.endDate + 'T00:00:00') : startDate;
  const ONE_DAY_MS = 24 * 60 * 60 * 1000;
  const totalDays = Math.round((endDate.getTime() - startDate.getTime()) / ONE_DAY_MS) + 1;
  
  const balanceKey = request.leaveType.toLowerCase(); 
  const userBalances = userData.emailToBalances[requestEmail];
  
  // Safety check for balance existence
  if (!userBalances || userBalances[balanceKey] === undefined) {
     throw new Error(`Balance type '${request.leaveType}' not found for user.`);
  }
  if (userBalances[balanceKey] < totalDays) {
    throw new Error(`Insufficient ${request.leaveType} balance. Available: ${userBalances[balanceKey]}, Requested: ${totalDays}.`);
  }

  // 4. File Upload Logic
  let sickNoteUrl = "";
  if (request.fileInfo) {
    try {
      const folder = DriveApp.getFolderById(SICK_NOTE_FOLDER_ID);
      const fileData = Utilities.base64Decode(request.fileInfo.data);
      const blob = Utilities.newBlob(fileData, request.fileInfo.type, request.fileInfo.name);
      const newFile = folder.createFile(blob).setName(`${requestName}_${new Date().toISOString()}_${request.fileInfo.name}`);
      sickNoteUrl = newFile.getUrl();
    } catch (e) { throw new Error("File upload failed: " + e.message); }
  }
  if (balanceKey === 'sick' && !sickNoteUrl) throw new Error("A PDF sick note is mandatory for sick leave.");

  // 5. Save to Sheet (With New Columns)
  const reqSheet = getOrCreateSheet(ss, SHEET_NAMES.leaveRequests);
  const requestID = `req_${new Date().getTime()}`;
  
  reqSheet.appendRow([
    requestID,
    status,
    requestEmail,
    requestName,
    request.leaveType,
    startDate, 
    endDate,   
    totalDays,
    request.reason,
    "", // ActionDate
    "", // ActionBy
    assignedApprover, // Col L (12): The person who must approve NOW
    "", // ActionReason
    sickNoteUrl,
    directManager, // Col O (15): Snapshot of Direct Mgr
    projectManager || "" // Col P (16): Snapshot of Project Mgr
  ]);
  
  SpreadsheetApp.flush(); 
  
  // Format the approver name for the success message
  const approverName = userData.emailToName[assignedApprover] || assignedApprover;
  return `Request submitted successfully. It is now ${status} (${approverName}).`;
}

function approveDenyRequest(adminEmail, requestID, newStatus, reason) {
  const ss = getSpreadsheet();
  // FIX: Use 'employeesCore' explicitly because that is where balances (Annual/Sick/Casual) live.
  const coreSheet = getOrCreateSheet(ss, SHEET_NAMES.employeesCore); 
  const userData = getUserDataFromDb(ss); // Pass 'ss', not a sheet object

  // Security Check
  const adminRole = userData.emailToRole[adminEmail] || 'agent';
  if (adminRole === 'agent') throw new Error("Permission denied.");

  const reqSheet = getOrCreateSheet(ss, SHEET_NAMES.leaveRequests); 
  const allData = reqSheet.getDataRange().getValues();
  
  let rowIndex = -1;
  let requestRow = [];
  
  // Find Request
  for (let i = 1; i < allData.length; i++) { 
    if (allData[i][0] === requestID) { 
      rowIndex = i + 1;
      requestRow = allData[i];
      break; 
    }
  }
  if (rowIndex === -1) throw new Error("Request ID not found.");

  const currentStatus = requestRow[1]; // Column B
  const requesterEmail = requestRow[2];
  
  // Col O (Index 14) = Direct Manager Snapshot
  // Col P (Index 15) = Project Manager Snapshot
  const directManager = requestRow[14]; 
  const projectManager = requestRow[15]; 

  // 1. Handle Denial (Immediate Stop)
  if (newStatus === 'Denied') {
    reqSheet.getRange(rowIndex, 2).setValue("Denied");
    reqSheet.getRange(rowIndex, 10).setValue(new Date()); // ActionDate
    reqSheet.getRange(rowIndex, 11).setValue(adminEmail); // ActionBy
    reqSheet.getRange(rowIndex, 13).setValue(reason || "Denied by " + adminEmail);
    return "Request denied and closed.";
  }

  // 2. Handle Approval Logic (Chain: Direct Mgr -> Project Mgr -> Approved)
  
  // STEP 1: Direct Manager Approves
  if (currentStatus === "Pending Direct Mgr") {
    // Check if there is a Project Manager to forward to.
    // Logic: If PM exists, is NOT "na", and is DIFFERENT from Direct Mgr, forward it.
    if (projectManager && projectManager !== "" && projectManager.toLowerCase() !== "na" && projectManager !== directManager) {
      // Forward to Project Manager
      reqSheet.getRange(rowIndex, 2).setValue("Pending Project Mgr"); // Update Status
      reqSheet.getRange(rowIndex, 12).setValue(projectManager);       // Update Assigned Approver (Col L)
      reqSheet.getRange(rowIndex, 13).setValue(`Direct Mgr (${adminEmail}) Approved. Forwarded to Project Mgr.`);
      return "Approved by Direct Manager. Forwarded to Project Manager for final approval.";
    } else {
      // No Project Manager (or same person), so Finalize immediately
      return finalizeLeaveApproval(ss, coreSheet, userData, reqSheet, rowIndex, requestRow, adminEmail, reason);
    }
  }

  // STEP 2: Project Manager Approves
  if (currentStatus === "Pending Project Mgr") {
    return finalizeLeaveApproval(ss, coreSheet, userData, reqSheet, rowIndex, requestRow, adminEmail, reason);
  }

  // Fallback for "Pending" (Legacy or simple flow)
  if (currentStatus === "Pending") {
     return finalizeLeaveApproval(ss, coreSheet, userData, reqSheet, rowIndex, requestRow, adminEmail, reason);
  }

  throw new Error(`Invalid Request Status for Approval: ${currentStatus}`);
}

// HELPER: Finalizes the request (Deducts balance, Adds to schedule, Updates status)
function finalizeLeaveApproval(ss, coreSheet, userData, reqSheet, rowIndex, requestRow, adminEmail, reason) {
    const requesterEmail = requestRow[2];
    const leaveType = requestRow[4];
    const totalDays = requestRow[7];
    const balanceKey = leaveType.toLowerCase();
    
    // 1. Deduct Balance from Employees_Core
    const userDBRow = userData.emailToRow[requesterEmail];
    
    // Column Mapping for Employees_Core (1-based index)
    // H=8 (Annual), I=9 (Sick), J=10 (Casual)
    const colMap = { "annual": 8, "sick": 9, "casual": 10 };
    const balanceCol = colMap[balanceKey];
    
    if (balanceCol && userDBRow) {
      // Ensure we are reading/writing numbers
      const balanceRange = coreSheet.getRange(userDBRow, balanceCol);
      const currentBal = parseFloat(balanceRange.getValue()) || 0;
      balanceRange.setValue(currentBal - totalDays);
    } else {
      // If it's a type like "Absent" or "Unpaid", we might not deduct, or log warning.
      console.warn(`No balance column found for type: ${leaveType}`);
    }

    // 2. Submit Schedule (Auto-log to Schedule Sheet)
    const reqName = requestRow[3];
    // Format dates for the schedule function
    const reqStartDateStr = Utilities.formatDate(new Date(requestRow[5]), Session.getScriptTimeZone(), "MM/dd/yyyy");
    const reqEndDateStr = Utilities.formatDate(new Date(requestRow[6]), Session.getScriptTimeZone(), "MM/dd/yyyy");
    
    // Call existing schedule function
    // (Ensure submitScheduleRange is defined in your code.gs)
    submitScheduleRange(adminEmail, requesterEmail, reqName, reqStartDateStr, reqEndDateStr, "", "", leaveType);

    // 3. Update Request Sheet to "Approved"
    reqSheet.getRange(rowIndex, 2).setValue("Approved");
    reqSheet.getRange(rowIndex, 10).setValue(new Date()); // ActionDate
    reqSheet.getRange(rowIndex, 11).setValue(adminEmail); // ActionBy
    reqSheet.getRange(rowIndex, 13).setValue(reason || "Final Approval");

    return "Final Approval Granted. Schedule updated and balance deducted.";
}

// ================= NEW/MODIFIED FUNCTIONS =================

// ================= FIXED HISTORY READER =================
function getAdherenceRange(adminEmail, userNames, startDateStr, endDateStr) {
  const ss = getSpreadsheet();
  const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
  const userData = getUserDataFromDb(dbSheet);
  const adminRole = userData.emailToRole[adminEmail] || 'agent';
  const timeZone = Session.getScriptTimeZone();
  
  let targetUserNames = [];
  if (adminRole === 'agent') {
    const selfName = userData.emailToName[adminEmail];
    if (!selfName) throw new Error("Your user account was not found.");
    targetUserNames = [selfName];
  } else {
    targetUserNames = userNames;
  }

  const targetUserSet = new Set(targetUserNames.map(name => name.toLowerCase()));
  const startDate = new Date(startDateStr);
  const endDate = new Date(endDateStr);
  startDate.setHours(0, 0, 0, 0);
  endDate.setHours(23, 59, 59, 999);
  
  const results = [];
  const scheduleSheet = getOrCreateSheet(ss, SHEET_NAMES.schedule);
  const scheduleData = scheduleSheet.getDataRange().getValues();
  const scheduleMap = {}; 

  for (let i = 1; i < scheduleData.length; i++) {
    const schName = (scheduleData[i][0] || "").toLowerCase();
    if (targetUserSet.has(schName)) {
      try {
        const schDate = parseDate(scheduleData[i][1]);
        if (schDate >= startDate && schDate <= endDate) {
          const schDateStr = Utilities.formatDate(schDate, timeZone, "MM/dd/yyyy");
          const leaveType = scheduleData[i][5] || "Present";
          scheduleMap[`${schName}:${schDateStr}`] = leaveType;
        }
      } catch (e) {}
    }
  }

  const adherenceSheet = getOrCreateSheet(ss, SHEET_NAMES.adherence);
  const adherenceData = adherenceSheet.getDataRange().getValues();
  const resultsLookup = new Set();

  for (let i = 1; i < adherenceData.length; i++) {
    const row = adherenceData[i];
    const rowUser = (row[1] || "").toString().trim().toLowerCase();

    if (targetUserSet.has(rowUser)) {
      try {
        const rowDate = new Date(row[0]);
        if (rowDate >= startDate && rowDate <= endDate) {
          results.push({
            date: convertDateToString(row[0]),
            userName: row[1],
            login: convertDateToString(row[2]),
            firstBreakIn: convertDateToString(row[3]),
            firstBreakOut: convertDateToString(row[4]),
            lunchIn: convertDateToString(row[5]),
            lunchOut: convertDateToString(row[6]),
            lastBreakIn: convertDateToString(row[7]),
            lastBreakOut: convertDateToString(row[8]),
            logout: convertDateToString(row[9]),
            // Fix: Explicitly parse numbers
            tardy: Number(row[10]) || 0,
            overtime: Number(row[11]) || 0,
            earlyLeave: Number(row[12]) || 0,
            leaveType: row[13] || "Present", // Fallback if missing
            firstBreakExceed: row[16] || 0,
            lunchExceed: row[17] || 0,
            lastBreakExceed: row[18] || 0,
            breakWindowViolation: row[21] || "No",
            netLoginHours: row[22] || 0,
            preShiftOvertime: Number(row[23]) || 0 // Col X (Index 23)
          });
          const rDateStr = Utilities.formatDate(rowDate, timeZone, "MM/dd/yyyy");
          resultsLookup.add(`${rowUser}:${rDateStr}`);
        }
      } catch (e) {
        Logger.log(`Skipping adherence row ${i+1}. Error: ${e.message}`);
      }
    }
  }

  // Fill in missing days
  let currentDate = new Date(startDate);
  const oneDayInMs = 24 * 60 * 60 * 1000;
  
  while (currentDate <= endDate) {
    const currentDateStr = Utilities.formatDate(currentDate, timeZone, "MM/dd/yyyy");
    for (const userName of targetUserNames) {
      const userNameLower = userName.toLowerCase();
      const adherenceKey = `${userNameLower}:${currentDateStr}`;
      
      if (!resultsLookup.has(adherenceKey)) {
        const scheduleKey = `${userNameLower}:${currentDateStr}`;
        const leaveType = scheduleMap[scheduleKey]; 
        let finalLeaveType = "Day Off";
        
        if (leaveType) {
          finalLeaveType = (leaveType.toLowerCase() === "present") ? "Absent" : leaveType;
        }

        results.push({
          date: convertDateToString(currentDate),
          userName: userName,
          login: null, firstBreakIn: null, firstBreakOut: null, lunchIn: null,
          lunchOut: null, lastBreakIn: null, lastBreakOut: null, logout: null,
          tardy: 0, overtime: 0, earlyLeave: 0,
          leaveType: finalLeaveType,
          firstBreakExceed: 0, lunchExceed: 0, lastBreakExceed: 0,
          preShiftOvertime: 0
        });
      }
    }
    currentDate.setTime(currentDate.getTime() + oneDayInMs);
  }

  results.sort((a, b) => {
    if (a.date < b.date) return -1;
    if (a.date > b.date) return 1;
    return a.userName.localeCompare(b.userName);
  });

  return results;
}


// REPLACE this function
function getMySchedule(userEmail) {
  const ss = getSpreadsheet();
  const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
  const userData = getUserDataFromDb(dbSheet);
  const userRole = userData.emailToRole[userEmail] || 'agent';

  const targetEmails = new Set();
  if (userRole === 'agent') {
    targetEmails.add(userEmail);
  } else {
    const subEmails = webGetAllSubordinateEmails(userEmail);
    subEmails.forEach(email => targetEmails.add(email.toLowerCase()));
  }

  const scheduleSheet = getOrCreateSheet(ss, SHEET_NAMES.schedule);
  const scheduleData = scheduleSheet.getDataRange().getValues();
  const timeZone = Session.getScriptTimeZone();
  
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  const nextSevenDays = new Date(today);
  nextSevenDays.setDate(today.getDate() + 7);

  const mySchedule = [];
  for (let i = 1; i < scheduleData.length; i++) {
    const row = scheduleData[i];
    // *** MODIFIED: Read Email from Col G (index 6) ***
    const schEmail = (row[6] || "").toString().trim().toLowerCase(); 
    
    if (targetEmails.has(schEmail)) {
      try {
        // *** MODIFIED: Read Date from Col B (index 1) ***
        const schDate = parseDate(row[1]);
        if (schDate >= today && schDate < nextSevenDays) { 
          
          // *** MODIFIED: Read times/leave from Col C, E, F ***
          let startTime = row[2]; // Col C
          let endTime = row[4];   // Col E
          let leaveType = row[5] || ""; // Col F

          // *** MODIFIED for Request 3: Handle "Day Off" ***
          if (leaveType === "" && !startTime) {
            leaveType = "Day Off";
          } else if (leaveType === "" && startTime) {
            leaveType = "Present"; // Default if times exist but no type
          }
          // *** END MODIFICATION ***
          
          if (startTime instanceof Date) {
            startTime = Utilities.formatDate(startTime, timeZone, "HH:mm");
          }
          if (endTime instanceof Date) {
            endTime = Utilities.formatDate(endTime, timeZone, "HH:mm");
          }
          
          mySchedule.push({
            userName: userData.emailToName[schEmail] || schEmail,
            date: convertDateToString(schDate),
            leaveType: leaveType,
            startTime: startTime,
            endTime: endTime
          });
        }
      } catch(e) {
        Logger.log(`Skipping schedule row ${i+1}. Invalid date. Error: ${e.message}`);
      }
    }
  }
  
  mySchedule.sort((a, b) => {
    const dateA = new Date(a.date);
    const dateB = new Date(b.date);
    if (dateA < dateB) return -1;
    if (dateA > dateB) return 1;
    return a.userName.localeCompare(b.userName);
  });
  return mySchedule;
}


// (No Change)
function adjustLeaveBalance(adminEmail, userEmail, leaveType, amount, reason) {
  const ss = getSpreadsheet();
  const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
  const userData = getUserDataFromDb(dbSheet);
  
  const adminRole = userData.emailToRole[adminEmail] || 'agent';
  if (adminRole !== 'admin' && adminRole !== 'superadmin') {
    throw new Error("Permission denied. Only admins can adjust balances.");
  }
  
  const balanceKey = leaveType.toLowerCase();
  const balanceCol = { annual: 4, sick: 5, casual: 6 }[balanceKey];
  if (!balanceCol) {
    throw new Error(`Unknown leave type: ${leaveType}.`);
  }
  
  const userRow = userData.emailToRow[userEmail];
  const userName = userData.emailToName[userEmail];
  if (!userRow) {
    throw new Error(`Could not find user ${userName} in Data Base.`);
  }
  
  const balanceRange = dbSheet.getRange(userRow, balanceCol);
  const currentBalance = parseFloat(balanceRange.getValue()) || 0;
  const newBalance = currentBalance + amount;
  
  balanceRange.setValue(newBalance);
  
  // Log the adjustment
  const logsSheet = getOrCreateSheet(ss, SHEET_NAMES.logs);
  logsSheet.appendRow([
    new Date(), 
    userName, 
    adminEmail, 
    "Balance Adjustment", 
    `Admin: ${adminEmail} | User: ${userName} | Type: ${leaveType} | Amount: ${amount} | Reason: ${reason} | Old: ${currentBalance} | New: ${newBalance}`
  ]);
  
  return `Successfully adjusted ${userName}'s ${leaveType} balance from ${currentBalance} to ${newBalance}.`;
}

// ================= PHASE 9: BULK SCHEDULE IMPORTER =================
function importScheduleCSV(adminEmail, csvData) {
  const ss = getSpreadsheet();
  const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
  const userData = getUserDataFromDb(dbSheet);
  const adminRole = userData.emailToRole[adminEmail] || 'agent';
  if (adminRole !== 'admin' && adminRole !== 'superadmin' && adminRole !== 'manager') {
    throw new Error("Permission denied. Only admins/managers can import schedules.");
  }
  
  const scheduleSheet = getOrCreateSheet(ss, SHEET_NAMES.schedule);
  const scheduleData = scheduleSheet.getDataRange().getValues();
  const logsSheet = getOrCreateSheet(ss, SHEET_NAMES.logs);
  const timeZone = Session.getScriptTimeZone();
  
  // Build map of existing schedules
  const userScheduleMap = {};
  for (let i = 1; i < scheduleData.length; i++) {
    const rowEmail = scheduleData[i][6];
    const rowDateRaw = scheduleData[i][1]; 
    if (rowEmail && rowDateRaw) {
      const email = rowEmail.toLowerCase();
      if (!userScheduleMap[email]) userScheduleMap[email] = {};
      const rowDate = new Date(rowDateRaw);
      const rowDateStr = Utilities.formatDate(rowDate, timeZone, "MM/dd/yyyy");
      userScheduleMap[email][rowDateStr] = i + 1;
    }
  }
  
  let daysUpdated = 0;
  let daysCreated = 0;
  let errors = 0;
  let errorLog = [];

  for (const row of csvData) {
    try {
      const userName = row.Name;
      const userEmail = (row['agent email'] || "").toLowerCase();
      
      const targetStartDate = parseDate(row.StartDate);
      let startTime = parseCsvTime(row.ShiftStartTime, timeZone);
      const targetEndDate = parseDate(row.EndDate);
      let endTime = parseCsvTime(row.ShiftEndTime, timeZone);
      
      // --- PHASE 9: Parse New Break Windows ---
      let b1s = parseCsvTime(row.Break1Start, timeZone);
      let b1e = parseCsvTime(row.Break1End, timeZone);
      let ls = parseCsvTime(row.LunchStart, timeZone);
      let le = parseCsvTime(row.LunchEnd, timeZone);
      let b2s = parseCsvTime(row.Break2Start, timeZone);
      let b2e = parseCsvTime(row.Break2End, timeZone);
      
      let leaveType = row.LeaveType || "Present";
      
      if (!userName || !userEmail) throw new Error("Missing Name or agent email.");
      if (!targetStartDate || isNaN(targetStartDate.getTime())) throw new Error(`Invalid StartDate: ${row.StartDate}.`);
      
      const startDateStr = Utilities.formatDate(targetStartDate, timeZone, "MM/dd/yyyy");

      if (leaveType.toLowerCase() !== "present") {
        startTime = ""; endTime = "";
        b1s = ""; b1e = ""; ls = ""; le = ""; b2s = ""; b2e = "";
      }

      let finalEndDate;
      if (leaveType.toLowerCase() === "present" && targetEndDate && !isNaN(targetEndDate.getTime())) {
        finalEndDate = targetEndDate;
      } else {
        finalEndDate = new Date(targetStartDate);
      }

      const emailMap = userScheduleMap[userEmail] || {};
      
      const result = updateOrAddSingleSchedule(
      scheduleSheet, userScheduleMap, logsSheet,
      userEmail, userName, 
      currentDate, 
      shiftEndDate, 
      currentDateStr, 
      startTime, endTime, leaveType, puncherEmail,
      "", "", "", "", "", "" // <--- Pass empty break windows for manual entry
    );
      
      if (result === "UPDATED") daysUpdated++;
      if (result === "CREATED") daysCreated++;
    } catch (e) {
      errors++;
      errorLog.push(`Row ${row.Name}/${row.StartDate}: ${e.message}`);
    }
  }

  if (errors > 0) {
    return `Error: Import complete with ${errors} errors. (Created: ${daysCreated}, Updated: ${daysUpdated}). Errors: ${errorLog.join(' | ')}`;
  }
  return `Import successful. Records Created: ${daysCreated}, Records Updated: ${daysUpdated}.`;
}

// ================= PHASE 7: DASHBOARD ANALYTICS =================
function getDashboardData(adminEmail, userEmails, date) {
  const ss = getSpreadsheet();
  const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
  const userData = getUserDataFromDb(dbSheet);
  const adminRole = userData.emailToRole[adminEmail] || 'agent';
  if (adminRole !== 'admin' && adminRole !== 'superadmin' && adminRole !== 'manager') {
    throw new Error("Permission denied.");
  }
  
  const timeZone = Session.getScriptTimeZone();
  const targetDate = new Date(date);
  const targetDateStr = Utilities.formatDate(targetDate, timeZone, "MM/dd/yyyy");
  const targetUserSet = new Set(userEmails.map(e => e.toLowerCase()));
  
  const userStatusMap = {};
  const userMetricsMap = {}; 
  
  // WFM Aggregate Counters
  let countScheduled = 0;
  let countWorking = 0; // Logged In
  let countUnavailable = 0; // Absent, Leave, or Scheduled but not logged in yet
  
  userEmails.forEach(email => {
    const lEmail = email.toLowerCase();
    const name = userData.emailToName[lEmail] || lEmail;
    userStatusMap[lEmail] = "Day Off"; 
    userMetricsMap[name] = {
      name: name, tardy: 0, earlyLeave: 0, overtime: 0,
      breakExceed: 0, lunchExceed: 0, scheduled: false
    };
  });

  const scheduledEmails = new Set();

  // 1. Get Schedule Data & Calculate Capacity Base
  const scheduleSheet = getOrCreateSheet(ss, SHEET_NAMES.schedule);
  const scheduleData = scheduleSheet.getDataRange().getValues();
  for (let i = 1; i < scheduleData.length; i++) {
    const row = scheduleData[i];
    const schEmail = (row[6] || "").toLowerCase();
    if (!targetUserSet.has(schEmail)) continue;
    
    const schDate = new Date(row[1]);
    const schDateStr = Utilities.formatDate(schDate, timeZone, "MM/dd/yyyy");
    
    if (schDateStr === targetDateStr) {
      const leaveType = (row[5] || "").toString().trim().toLowerCase();
      const startTime = row[2];
      
      if (leaveType === "" && !startTime) {
        userStatusMap[schEmail] = "Day Off";
      } else if (leaveType === "present" || (leaveType === "" && startTime)) {
        scheduledEmails.add(schEmail);
        userStatusMap[schEmail] = "Pending Login";
        
        // Mark metric object as scheduled
        const name = userData.emailToName[schEmail];
        if (userMetricsMap[name]) userMetricsMap[name].scheduled = true;
        
        countScheduled++;
        countUnavailable++; // Assume unavailable until we find a punch
      } else if (leaveType === "absent") {
        userStatusMap[schEmail] = "Absent";
        countScheduled++; // Absent counts as scheduled but lost
        countUnavailable++;
      } else {
        userStatusMap[schEmail] = "On Leave";
        // Leave usually implies scheduled hours that are now non-productive
        countScheduled++;
        countUnavailable++;
      }
    }
  }
  
  // 2. Get Adherence & Status
  const adherenceSheet = getOrCreateSheet(ss, SHEET_NAMES.adherence);
  const adherenceData = adherenceSheet.getDataRange().getValues();
  
  // Pre-fetch other codes for status refinement
  const otherCodesSheet = getOrCreateSheet(ss, SHEET_NAMES.otherCodes);
  const otherCodesData = otherCodesSheet.getDataRange().getValues();
  const userLastOtherCode = {};
  
  for (let i = otherCodesData.length - 1; i > 0; i--) { 
    const row = otherCodesData[i];
    const rowDate = new Date(row[0]);
    const rowShiftDate = getShiftDate(rowDate, SHIFT_CUTOFF_HOUR);
    if (Utilities.formatDate(rowShiftDate, timeZone, "MM/dd/yyyy") === targetDateStr) {
      const uName = row[1];
      const uEmail = userData.nameToEmail[uName];
      if (uEmail && targetUserSet.has(uEmail.toLowerCase())) {
        if (!userLastOtherCode[uEmail.toLowerCase()]) { 
          const [code, type] = (row[2] || "").split(" ");
          userLastOtherCode[uEmail.toLowerCase()] = { code: code, type: type };
        }
      }
    }
  }
  
  let totalDeviationSeconds = 0;

  for (let i = 1; i < adherenceData.length; i++) {
    const row = adherenceData[i];
    const rowDate = new Date(row[0]);
    if (Utilities.formatDate(rowDate, timeZone, "MM/dd/yyyy") === targetDateStr) { 
      const userName = row[1];
      const userEmail = userData.nameToEmail[userName];
      
      if (userEmail && targetUserSet.has(userEmail.toLowerCase())) {
        const lEmail = userEmail.toLowerCase();
        
        // Status Logic
        if (scheduledEmails.has(lEmail)) {
          const login = row[2], b1_in = row[3], b1_out = row[4], l_in = row[5],
                l_out = row[6], b2_in = row[7], b2_out = row[8], logout = row[9];
          
          let agentStatus = "Pending Login";
          
          if (login && !logout) {
            agentStatus = "Logged In";
            
            // Check sub-status
            const lastOther = userLastOtherCode[lEmail];
            let onBreak = false;
            
            if (lastOther && lastOther.type === 'In') {
              agentStatus = `On ${lastOther.code}`;
              onBreak = true;
            } else {
              if (b1_in && !b1_out) { agentStatus = "On First Break"; onBreak = true; }
              if (l_in && !l_out) { agentStatus = "On Lunch"; onBreak = true; }
              if (b2_in && !b2_out) { agentStatus = "On Last Break"; onBreak = true; }
            }
            
            if (!onBreak) {
               // They are truly working
               countWorking++;
               countUnavailable--; // They were counted as unavailable initially
            }
          } else if (login && logout) {
            agentStatus = "Logged Out";
          }
          userStatusMap[lEmail] = agentStatus;
          scheduledEmails.delete(lEmail); // Remove from set so we don't process again
        }
        
        // Metrics Summation
        // Metrics Summation (Ensuring Pre-Shift OT is added)
        const tardy = parseFloat(row[10]) || 0;
        const overtime = parseFloat(row[11]) || 0; // Post-Shift
        const earlyLeave = parseFloat(row[12]) || 0;
        const breakExceed = (parseFloat(row[16]) || 0) + (parseFloat(row[18]) || 0);
        const lunchExceed = parseFloat(row[17]) || 0;
        const preShiftOT = parseFloat(row[23]) || 0; // Col X (Index 23)

        // Sum deviation for Schedule Adherence
        totalDeviationSeconds += (tardy + earlyLeave + breakExceed + lunchExceed);

        if (userMetricsMap[userName]) {
          userMetricsMap[userName].tardy += tardy;
          userMetricsMap[userName].earlyLeave += earlyLeave;
          userMetricsMap[userName].overtime += (overtime + preShiftOT); // Combined OT
          userMetricsMap[userName].breakExceed += breakExceed;
          userMetricsMap[userName].lunchExceed += lunchExceed;
        }
      }
    }
  }
  
  // 3. Get Pending Requests (Same as before)
  const reqSheet = getOrCreateSheet(ss, SHEET_NAMES.leaveRequests);
  const reqData = reqSheet.getDataRange().getValues();
  const pendingRequests = [];
  for (let i = 1; i < reqData.length; i++) {
    const row = reqData[i];
    const reqEmail = (row[2] || "").toLowerCase();
    if (row[1] && row[1].toString().toLowerCase().includes('pending') && targetUserSet.has(reqEmail)) {
      pendingRequests.push({ name: row[3], type: row[4], startDate: convertDateToString(new Date(row[5])), days: row[7] });
    }
  }
  
  // 4. Calculate Final WFM Metrics
  // Assumption: Avg shift is 9 hours (32400 sec) for calculation
  const ESTIMATED_SHIFT_SECONDS = 32400; 
  const totalScheduledSeconds = countScheduled * ESTIMATED_SHIFT_SECONDS;
  
  let adherencePct = 100;
  if (totalScheduledSeconds > 0) {
    adherencePct = Math.max(0, 100 - ((totalDeviationSeconds / totalScheduledSeconds) * 100));
  }
  
  let capacityPct = 0;
  if (countScheduled > 0) {
    capacityPct = (countWorking / countScheduled) * 100;
  }
  
  let shrinkagePct = 0;
  if (countScheduled > 0) {
    // Shrinkage = Agents unavailable / Total Scheduled
    shrinkagePct = (countUnavailable / countScheduled) * 100;
  }

  const agentStatusList = [];
  for (const email of targetUserSet) {
      const name = userData.emailToName[email] || email;
      const status = userStatusMap[email] || "Day Off";
      agentStatusList.push({ name: name, status: status });
  }
  agentStatusList.sort((a, b) => a.name.localeCompare(b.name));
  
  const individualAdherenceMetrics = Object.values(userMetricsMap);

  return {
    wfmMetrics: {
      adherence: adherencePct.toFixed(1),
      capacity: capacityPct.toFixed(1),
      shrinkage: shrinkagePct.toFixed(1),
      working: countWorking,
      scheduled: countScheduled,
      unavailable: countUnavailable
    },
    agentStatusList: agentStatusList,
    individualAdherenceMetrics: individualAdherenceMetrics,
    pendingRequests: pendingRequests
  };
}

// --- NEW: "My Team" Helper Functions ---
function saveMyTeam(adminEmail, userEmails) {
  try {
    // Uses Google Apps Script's built-in User Properties for saving user-specific settings.
    const userProperties = PropertiesService.getUserProperties();
    userProperties.setProperty('myTeam', JSON.stringify(userEmails));
    return "Successfully saved 'My Team' preference.";
  } catch (e) {
    throw new Error("Failed to save team preferences: " + e.message);
  }
}

function getMyTeam(adminEmail) {
  try {
    const userProperties = PropertiesService.getUserProperties();
    // Getting properties implicitly forces the Google auth dialog if needed.
    const properties = userProperties.getProperties(); 
    const myTeam = properties['myTeam'];
    return myTeam ? JSON.parse(myTeam) : [];
  } catch (e) {
    Logger.log("Failed to load team preferences: " + e.message);
    // Throwing an error here would break the dashboard's initial load. 
    // We return an empty array instead, and let the front-end handle the fallback.
   return [];
  }
}

// --- NEW: Reporting Line Function ---
function updateReportingLine(adminEmail, userEmail, newSupervisorEmail) {
  const ss = getSpreadsheet();
  const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
  const userData = getUserDataFromDb(dbSheet);
  
  const adminRole = userData.emailToRole[adminEmail] || 'agent';
  if (adminRole !== 'admin' && adminRole !== 'superadmin') {
    throw new Error("Permission denied. Only admins can change reporting lines.");
  }
  
  const userName = userData.emailToName[userEmail];
  const newSupervisorName = userData.emailToName[newSupervisorEmail];
  if (!userName) throw new Error(`Could not find user: ${userEmail}`);
  if (!newSupervisorName) throw new Error(`Could not find new supervisor: ${newSupervisorEmail}`);

  const userRow = userData.emailToRow[userEmail];
  const currentUserSupervisor = userData.emailToSupervisor[userEmail];

  // Check for auto-approval
  let canAutoApprove = false;
  if (adminRole === 'superadmin') {
    canAutoApprove = true;
  } else if (adminRole === 'admin') {
    // Check if both the user's current supervisor AND the new supervisor report to this admin
    const currentSupervisorManager = userData.emailToSupervisor[currentUserSupervisor];
    const newSupervisorManager = userData.emailToSupervisor[newSupervisorEmail];
    
    if (currentSupervisorManager === adminEmail && newSupervisorManager === adminEmail) {
      canAutoApprove = true;
    }
  }

  if (!canAutoApprove) {
    // This is where we will build Phase 2 (requesting the change)
    // For now, we will just show a permission error.
    throw new Error("Permission Denied: You do not have authority to approve this change. (This will become a request in Phase 2).");
  }

  // --- Auto-Approval Logic ---
  // Update the SupervisorEmail column (Column G = 7)
  dbSheet.getRange(userRow, 7).setValue(newSupervisorEmail);
  
  // Log the change
  const logsSheet = getOrCreateSheet(ss, SHEET_NAMES.logs);
  logsSheet.appendRow([
    new Date(), 
    userName, 
    adminEmail, 
    "Reporting Line Change", 
    `User: ${userName} moved to Supervisor: ${newSupervisorName} by ${adminEmail}`
  ]);
  
  return `${userName} has been successfully reassigned to ${newSupervisorName}.`;
}

// [START] MODIFICATION 2: Replace _ONE_TIME_FIX_TEMPLATE


/**
 * NEW: User submits full registration details + 2 managers.
 * (FIXED: Header-aware mapping to prevent column misalignment)
 */
function webSubmitFullRegistration(form) {
  try {
    const userEmail = Session.getActiveUser().getEmail().toLowerCase();
    const ss = getSpreadsheet();
    const userData = getUserDataFromDb(ss);
    const regSheet = getOrCreateSheet(ss, SHEET_NAMES.pendingRegistrations);
    
    let userName = userEmail;
    const userObj = userData.userList.find(u => u.email === userEmail);
    if (userObj) userName = userObj.name;

    if (!form.directManager || !form.functionalManager) throw new Error("Both managers are required.");
    if (!form.address || !form.phone) throw new Error("Address and Phone are required.");

    // Check for existing
    const data = regSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === userEmail && data[i][5] !== 'Rejected' && data[i][5] !== 'Approved') { 
         throw new Error("You already have a pending registration request.");
      }
    }

    const requestID = `REG-${new Date().getTime()}`;
    const timestamp = new Date();

    // --- DYNAMIC HEADER MAPPING FIX ---
    const headers = regSheet.getRange(1, 1, 1, regSheet.getLastColumn()).getValues()[0];
    const newRow = new Array(headers.length).fill(""); // Initialize empty row matching header length

    // Helper to map value to header name
    const setCol = (headerName, val) => {
        const idx = headers.indexOf(headerName);
        if (idx > -1) newRow[idx] = val;
    };

    // Map Data
    setCol("RequestID", requestID);
    setCol("UserEmail", userEmail);
    setCol("UserName", userName);
    setCol("DirectManagerEmail", form.directManager);
    setCol("FunctionalManagerEmail", form.functionalManager);
    setCol("DirectStatus", "Pending");
    setCol("FunctionalStatus", "Pending");
    setCol("Address", form.address);
    setCol("Phone", form.phone);
    setCol("RequestTimestamp", timestamp);
    setCol("WorkflowStage", 1); // Explicitly set Stage 1
    
    // --- END FIX ---

    regSheet.appendRow(newRow);
    return "Registration submitted! Waiting for Direct Manager approval.";

  } catch (err) {
    Logger.log("webSubmitFullRegistration Error: " + err.message);
    return "Error: " + err.message;
  }
}

/**
 * For the pending user to check their own status.
 */
function webGetMyRegistrationStatus() {
  try {
    const userEmail = Session.getActiveUser().getEmail().toLowerCase();
    const regSheet = getOrCreateSheet(getSpreadsheet(), SHEET_NAMES.pendingRegistrations);
    const data = regSheet.getDataRange().getValues();

    for (let i = data.length - 1; i > 0; i--) { // Check newest first
      if (data[i][1] === userEmail) {
        return { status: data[i][4], supervisor: data[i][3] }; // Returns { status: "Pending" } or { status: "Denied" }
      }
    }
    return { status: "New" }; // No submission found
  } catch (e) {
    return { error: e.message };
  }
}

/**
 * NEW: Admins see requests where THEY are the approver.
 * (FIXED: Auto-corrects missing Stage 1 data)
 */
function webGetPendingRegistrations() {
  try {
    const adminEmail = Session.getActiveUser().getEmail().toLowerCase();
    const ss = getSpreadsheet();
    const userData = getUserDataFromDb(ss);
    const adminRole = userData.emailToRole[adminEmail] || 'agent';
    
    if (adminRole === 'agent') throw new Error("Permission denied.");

    const regSheet = getOrCreateSheet(ss, SHEET_NAMES.pendingRegistrations);
    const data = regSheet.getDataRange().getValues();
    const pending = [];
    const headers = data[0];
    
    // Map Indexes
    const idx = {
      id: headers.indexOf("RequestID"),
      email: headers.indexOf("UserEmail"),
      name: headers.indexOf("UserName"),
      dm: headers.indexOf("DirectManagerEmail"),
      fm: headers.indexOf("FunctionalManagerEmail"),
      dmStat: headers.indexOf("DirectStatus"),
      fmStat: headers.indexOf("FunctionalStatus"),
      ts: headers.indexOf("RequestTimestamp"),
      hDate: headers.indexOf("HiringDate"),
      stage: headers.indexOf("WorkflowStage")
    };

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const directMgr = (row[idx.dm] || "").toLowerCase();
      const funcMgr = (row[idx.fm] || "").toLowerCase();
      const directStatus = row[idx.dmStat];
      const funcStatus = row[idx.fmStat];
      let stage = Number(row[idx.stage] || 0);
      
      // --- FIX: INFER STAGE IF MISSING ---
      if (stage === 0) {
          if (directStatus === 'Pending') stage = 1;
          else if (directStatus === 'Approved' && funcStatus === 'Pending') stage = 2;
      }
      // -----------------------------------

      const hiringDate = row[idx.hDate] ? convertDateToString(new Date(row[idx.hDate])).split('T')[0] : "";

      let actionRequired = false;
      let myRoleInRequest = "";

      if (adminRole === 'superadmin') {
        // Superadmin sees everything active
        if (stage === 1 || stage === 2) {
           actionRequired = true;
           myRoleInRequest = (stage === 1) ? "Direct" : "Functional";
        }
      } else {
        // Sequential Logic
        // Stage 1: Direct Manager must act
        if (stage === 1 && directMgr === adminEmail) {
          actionRequired = true;
          myRoleInRequest = "Direct";
        }
        // Stage 2: Functional/Project Manager must act (only after DM approved)
        else if (stage === 2 && funcMgr === adminEmail) {
          actionRequired = true;
          myRoleInRequest = "Functional";
        }
      }

      if (actionRequired) {
        pending.push({
          requestID: row[idx.id],
          userEmail: row[idx.email],
          userName: row[idx.name],
          approverRole: myRoleInRequest, 
          otherStatus: myRoleInRequest === "Direct" ? "Step 1 of 2" : "Step 2: Final Approval",
          timestamp: convertDateToString(new Date(row[idx.ts])),
          hiringDate: hiringDate,
          stage: stage
        });
      }
    }

    return pending.sort((a,b) => new Date(b.timestamp) - new Date(a.timestamp));
  } catch (e) {
    return { error: e.message };
  }
}

/**
 * NEW: Approves one side of the request. If both approved -> HIRE.
 */
function webApproveDenyRegistration(requestID, userEmail, supervisorEmail, newStatus, hiringDateStr) {
  try {
    const adminEmail = Session.getActiveUser().getEmail().toLowerCase();
    const ss = getSpreadsheet();
    const regSheet = getOrCreateSheet(ss, SHEET_NAMES.pendingRegistrations);
    const data = regSheet.getDataRange().getValues();
    const headers = data[0];
    
    // Find Indexes
    const idx = {
      id: headers.indexOf("RequestID"),
      dmStat: headers.indexOf("DirectStatus"),
      fmStat: headers.indexOf("FunctionalStatus"),
      hDate: headers.indexOf("HiringDate"),
      stage: headers.indexOf("WorkflowStage")
    };

    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][idx.id] === requestID) {
        rowIndex = i + 1;
        break;
      }
    }
    if (rowIndex === -1) throw new Error("Request not found.");

    const row = regSheet.getRange(rowIndex, 1, 1, regSheet.getLastColumn()).getValues()[0];
    const currentStage = Number(row[idx.stage] || 1);

    // --- DENY LOGIC (Applies to both steps) ---
    if (newStatus === 'Denied') {
      regSheet.getRange(rowIndex, idx.dmStat + 1).setValue("Denied");
      regSheet.getRange(rowIndex, idx.fmStat + 1).setValue("Denied");
      // Reset stage or set to -1 to indicate closed
      regSheet.getRange(rowIndex, idx.stage + 1).setValue(-1); 
      return { success: true, message: "Registration denied and closed." };
    }

    // --- APPROVAL LOGIC ---
    
    // STEP 1: Direct Manager Approval
    if (currentStage === 1) {
      if (!hiringDateStr) throw new Error("Direct Manager must provide a Hiring Date.");
      
      // Validate Date
      if (isNaN(new Date(hiringDateStr).getTime())) throw new Error("Invalid Hiring Date.");

      regSheet.getRange(rowIndex, idx.dmStat + 1).setValue("Approved");
      regSheet.getRange(rowIndex, idx.hDate + 1).setValue(new Date(hiringDateStr)); // Save Date
      regSheet.getRange(rowIndex, idx.stage + 1).setValue(2); // Move to Stage 2
      
      return { success: true, message: "Step 1 Approved. Request forwarded to Project Manager." };
    }

    // STEP 2: Project Manager Approval
    if (currentStage === 2) {
      // Finalize
      regSheet.getRange(rowIndex, idx.fmStat + 1).setValue("Approved");
      regSheet.getRange(rowIndex, idx.stage + 1).setValue(3); // Completed
      
      // Reuse existing activation logic, ensuring we pass the hiring date from the sheet if not passed explicitly
      const finalHiringDate = hiringDateStr || row[idx.hDate];
      return activateUser(ss, row, finalHiringDate);
    }

    return { success: false, message: "Invalid Workflow State." };

  } catch (e) {
    Logger.log("webApproveDenyRegistration Error: " + e.message);
    return { error: e.message };
  }
}

// Helper to finalize activation
function activateUser(ss, regRow, hiringDateStr) {
  const userEmail = regRow[1];
  const userName = regRow[2];
  const directMgr = regRow[3];
  const funcMgr = regRow[4];
  const address = regRow[7];
  const phone = regRow[8];

  // Update Core & PII
  const coreSheet = getOrCreateSheet(ss, SHEET_NAMES.employeesCore);
  const piiSheet = getOrCreateSheet(ss, SHEET_NAMES.employeesPII);

  // 1. Find user in Core (created during auto-registration in getUserInfo)
  const coreData = coreSheet.getDataRange().getValues();
  let coreRow = -1;
  for(let i=1; i<coreData.length; i++) {
      if (coreData[i][2] === userEmail) { // Col C is Email
          coreRow = i + 1;
          break;
      }
  }
  
  if (coreRow === -1) throw new Error("User record missing in Core DB.");

  // --- FIX: CHECK AND UPDATE EMPLOYEE ID IF PENDING ---
  let empID = coreSheet.getRange(coreRow, 1).getValue();
  
  if (String(empID).includes("PENDING")) {
    const newEmpID = generateNextEmpID(coreSheet);
    coreSheet.getRange(coreRow, 1).setValue(newEmpID); // Update ID in Core
    empID = newEmpID; // Use new ID for subsequent steps
    Logger.log(`Updated user ${userName} from PENDING ID to ${newEmpID}`);
  }
  // ----------------------------------------------------

  // Update Core: Status, Managers
  coreSheet.getRange(coreRow, 5).setValue("Active"); // Status
  coreSheet.getRange(coreRow, 6).setValue(directMgr);
  coreSheet.getRange(coreRow, 7).setValue(funcMgr);

  // Update PII: Address, Phone, Hiring Date
  const piiData = piiSheet.getDataRange().getValues();
  let piiRow = -1;
  for(let i=1; i<piiData.length; i++) {
      if (piiData[i][0] === empID) {
          piiRow = i + 1;
          break;
      }
  }
  
  // If PII row doesn't exist, create it with the PERMANENT ID
  if (piiRow === -1) {
      piiSheet.appendRow([
        empID, 
        new Date(hiringDateStr), 
        "", "", 
        address, 
        phone, 
        "", ""
      ]);
  } else {
      piiSheet.getRange(piiRow, 2).setValue(new Date(hiringDateStr));
      piiSheet.getRange(piiRow, 5).setValue(address);
      piiSheet.getRange(piiRow, 6).setValue(phone);
  }
  
  // Create Folders with PERMANENT ID
   try {
      const rootFolders = DriveApp.getFoldersByName("KOMPASS_HR_Files");
      if (rootFolders.hasNext()) {
        const root = rootFolders.next();
        const empFolders = root.getFoldersByName("Employee_Files");
        if (empFolders.hasNext()) {
          const parent = empFolders.next();
          // Folder Name format: Name_KOM-100X
          const personalFolder = parent.createFolder(`${userName}_${empID}`);
          personalFolder.createFolder("Payslips");
          personalFolder.createFolder("Onboarding_Docs");
          personalFolder.createFolder("Sick_Notes");
        }
      }
    } catch (e) {
      Logger.log("Folder creation error: " + e.message);
    }

  return { success: true, message: `User activated with ID ${empID}!` };
}
// --- ADD TO THE END OF code.gs ---

// ==========================================================
// === ANNOUNCEMENTS MODULE ===
// ==========================================================

/**
 * Fetches only active announcements for all users.
 */
function webGetAnnouncements() {
  try {
    const ss = getSpreadsheet();
    const sheet = getOrCreateSheet(ss, SHEET_NAMES.announcements);
    const data = sheet.getDataRange().getValues();
    const announcements = [];
    
    // Loop backwards to get newest first
    for (let i = data.length - 1; i > 0; i--) {
      const row = data[i];
      const status = row[2];
      
      if (status === 'Active') {
        announcements.push({
          id: row[0],
          content: row[1]
        });
      }
    }
    return announcements;
    
  } catch (e) {
    Logger.log("webGetAnnouncements Error: " + e.message);
    return []; // Return empty array on error
  }
}

/**
 * Fetches all announcements for the admin panel.
 * Only Superadmins can access this.
 */
function webGetAnnouncements_Admin() {
  try {
    const adminEmail = Session.getActiveUser().getEmail().toLowerCase();
    const ss = getSpreadsheet();
    const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
    const userData = getUserDataFromDb(dbSheet);
    
    if (userData.emailToRole[adminEmail] !== 'superadmin') {
      throw new Error("Permission denied. Only superadmins can manage announcements.");
    }

    const sheet = getOrCreateSheet(ss, SHEET_NAMES.announcements);
    const data = sheet.getDataRange().getValues();
    const results = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      results.push({
        id: row[0],
        content: row[1],
        status: row[2],
        createdBy: row[3],
        timestamp: convertDateToString(new Date(row[4]))
      });
    }
    
    return results;

  } catch (e) {
    Logger.log("webGetAnnouncements_Admin Error: " + e.message);
    return { error: e.message };
  }
}

/**
 * Saves (creates or updates) an announcement.
 * Only Superadmins can access this.
 */
function webSaveAnnouncement(announcementObject) {
  try {
    const adminEmail = Session.getActiveUser().getEmail().toLowerCase();
    const ss = getSpreadsheet();
    const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
    const userData = getUserDataFromDb(dbSheet);
    
    if (userData.emailToRole[adminEmail] !== 'superadmin') {
      throw new Error("Permission denied. Only superadmins can save announcements.");
    }

    const sheet = getOrCreateSheet(ss, SHEET_NAMES.announcements);
    const { id, content, status } = announcementObject;

    if (!content) {
      throw new Error("Content cannot be empty.");
    }

    if (id) {
      // --- Update Existing ---
      const data = sheet.getDataRange().getValues();
      let rowFound = -1;
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === id) {
          rowFound = i + 1;
          break;
        }
      }
      
      if (rowFound === -1) {
        throw new Error("Announcement ID not found. Could not update.");
      }
      
      sheet.getRange(rowFound, 2).setValue(content);
      sheet.getRange(rowFound, 3).setValue(status);
      
    } else {
      // --- Create New ---
      const newID = `ann-${new Date().getTime()}`;
      sheet.appendRow([
        newID,
        content,
        status,
        adminEmail,
        new Date()
      ]);
    }
    
    SpreadsheetApp.flush();
    return { success: true };

  } catch (e) {
    Logger.log("webSaveAnnouncement Error: " + e.message);
    return { error: e.message };
  }
}

/**
 * Deletes an announcement.
 * Only Superadmins can access this.
 */
function webDeleteAnnouncement(announcementID) {
  try {
    const adminEmail = Session.getActiveUser().getEmail().toLowerCase();
    const ss = getSpreadsheet();
    const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
    const userData = getUserDataFromDb(dbSheet);
    
    if (userData.emailToRole[adminEmail] !== 'superadmin') {
      throw new Error("Permission denied. Only superadmins can delete announcements.");
    }

    const sheet = getOrCreateSheet(ss, SHEET_NAMES.announcements);
    const data = sheet.getDataRange().getValues();
    let rowFound = -1;

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === announcementID) {
        rowFound = i + 1;
        break;
      }
    }

    if (rowFound > -1) {
      sheet.deleteRow(rowFound);
      SpreadsheetApp.flush();
      return { success: true };
    } else {
      throw new Error("Announcement not found.");
    }

  } catch (e) {
    Logger.log("webDeleteAnnouncement Error: " + e.message);
    return { error: e.message };
  }
}

/**
 * NEW: Logs a request from a user to upgrade their role.
 */
function webRequestAdminAccess(justification, requestedRole) {
  try {
    const userEmail = Session.getActiveUser().getEmail().toLowerCase();
    const ss = getSpreadsheet();
    const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
    const userData = getUserDataFromDb(dbSheet);
    
    const userName = userData.emailToName[userEmail];
    const currentRole = userData.emailToRole[userEmail] || 'agent';

    if (!userName) {
      throw new Error("Your user account could not be found.");
    }
    if (currentRole === 'superadmin') {
      throw new Error("You are already a Superadmin.");
    }
    if (currentRole === 'admin' && requestedRole === 'admin') {
      throw new Error("You are already an Admin.");
    }
    if (currentRole === 'agent' && requestedRole === 'superadmin') {
      throw new Error("You must be an Admin to request Superadmin access.");
    }

    const reqSheet = getOrCreateSheet(ss, SHEET_NAMES.roleRequests);
    const requestID = `ROLE-${new Date().getTime()}`;

    // ...
reqSheet.appendRow([
  requestID,
  userEmail,
  userName,
  currentRole,
  requestedRole,
  justification,
  new Date(),
  "Pending", // *** ADD "Pending" STATUS ***
  "",        // ActionByEmail
  ""         // ActionTimestamp
]);

    return "Your role upgrade request has been submitted for review.";

  } catch (e) {
    Logger.log("webRequestAdminAccess Error: " + e.message);
    return "Error: " + e.message;
  }
}

/**
 * Fetches pending role requests. Superadmin only.
 */
function webGetRoleRequests() {
  try {
    const adminEmail = Session.getActiveUser().getEmail().toLowerCase();
    const ss = getSpreadsheet();
    const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
    const userData = getUserDataFromDb(dbSheet);

    if (userData.emailToRole[adminEmail] !== 'superadmin') {
      throw new Error("Permission denied. Only superadmins can view role requests.");
    }

    const reqSheet = getOrCreateSheet(ss, SHEET_NAMES.roleRequests);
    const data = reqSheet.getDataRange().getValues();
    const headers = data[0];
    const results = [];
    
    // Find column indexes
    const statusIndex = headers.indexOf("Status");
    const idIndex = headers.indexOf("RequestID");
    const emailIndex = headers.indexOf("UserEmail");
    const nameIndex = headers.indexOf("UserName");
    const currentIndex = headers.indexOf("CurrentRole");
    const requestedIndex = headers.indexOf("RequestedRole");
    const justifyIndex = headers.indexOf("Justification");
    const timeIndex = headers.indexOf("RequestTimestamp");

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[statusIndex] === 'Pending') {
        results.push({
          requestID: row[idIndex],
          userEmail: row[emailIndex],
          userName: row[nameIndex],
          currentRole: row[currentIndex],
          requestedRole: row[requestedIndex],
          justification: row[justifyIndex],
          timestamp: convertDateToString(new Date(row[timeIndex]))
        });
      }
    }
    return results.sort((a,b) => new Date(b.timestamp) - new Date(a.timestamp)); // Newest first
  } catch (e) {
    Logger.log("webGetRoleRequests Error: " + e.message);
    return { error: e.message };
  }
}

/**
 * Approves or denies a role request. Superadmin only.
 */
function webApproveDenyRoleRequest(requestID, newStatus) {
  try {
    const adminEmail = Session.getActiveUser().getEmail().toLowerCase();
    const ss = getSpreadsheet();
    const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
    const userData = getUserDataFromDb(dbSheet);

    if (userData.emailToRole[adminEmail] !== 'superadmin') {
      throw new Error("Permission denied. Only superadmins can action role requests.");
    }

    const reqSheet = getOrCreateSheet(ss, SHEET_NAMES.roleRequests);
    const data = reqSheet.getDataRange().getValues();
    const headers = data[0];

    // Find columns
    const idIndex = headers.indexOf("RequestID");
    const statusIndex = headers.indexOf("Status");
    const emailIndex = headers.indexOf("UserEmail");
    const requestedIndex = headers.indexOf("RequestedRole");
    const actionByIndex = headers.indexOf("ActionByEmail");
    const actionTimeIndex = headers.indexOf("ActionTimestamp");
    
    let rowToUpdate = -1;
    let requestDetails = {};

    for (let i = 1; i < data.length; i++) {
      if (data[i][idIndex] === requestID) {
        rowToUpdate = i + 1; // 1-based index
        requestDetails = {
          status: data[i][statusIndex],
          userEmail: data[i][emailIndex],
          newRole: data[i][requestedIndex]
        };
        break;
      }
    }

    if (rowToUpdate === -1) throw new Error("Request ID not found.");
    if (requestDetails.status !== 'Pending') throw new Error(`This request has already been ${requestDetails.status}.`);

    // 1. Update the Role Request sheet
    reqSheet.getRange(rowToUpdate, statusIndex + 1).setValue(newStatus);
    reqSheet.getRange(rowToUpdate, actionByIndex + 1).setValue(adminEmail);
    reqSheet.getRange(rowToUpdate, actionTimeIndex + 1).setValue(new Date());

    // 2. If Approved, update the Data Base
    if (newStatus === 'Approved') {
      const userDBRow = userData.emailToRow[requestDetails.userEmail];
      if (!userDBRow) {
        throw new Error(`Could not find user ${requestDetails.userEmail} in Data Base to update role.`);
      }
      // Find Role column (Column C = 3)
      dbSheet.getRange(userDBRow, 3).setValue(requestDetails.newRole);
    }
    
    SpreadsheetApp.flush();
    return { success: true, message: `Request has been ${newStatus}.` };
  } catch (e) {
    Logger.log("webApproveDenyRoleRequest Error: " + e.message);
    return { error: e.message };
  }
}

// ADD this new function to the end of your code.gs file
/**
 * Calculates and adds leave balances monthly based on hiring date.
 * This function should be run on a monthly time-based trigger.
 */
function monthlyLeaveAccrual() {
  const ss = getSpreadsheet();
  const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
  const logsSheet = getOrCreateSheet(ss, SHEET_NAMES.logs);
  const userData = getUserDataFromDb(dbSheet);
  const today = new Date();
  
  Logger.log("Starting monthlyLeaveAccrual trigger...");

  for (const user of userData.userList) {
    try {
      const hiringDate = userData.emailToHiringDate[user.email];
      
      // Skip if no hiring date or account is not active
      if (!hiringDate || user.accountStatus !== 'Active') {
        continue;
      }

      // Calculate years of service
      const yearsOfService = (today.getTime() - hiringDate.getTime()) / (1000 * 60 * 60 * 24 * 365.25);
      
      let annualDaysPerYear;
      if (yearsOfService >= 10) {
        annualDaysPerYear = 30;
      } else if (yearsOfService >= 1) {
        annualDaysPerYear = 21;
      } else {
        annualDaysPerYear = 15;
      }

      const monthlyAccrual = annualDaysPerYear / 12;
      
      const userRow = userData.emailToRow[user.email];
      if (!userRow) continue; // Should not happen, but safe check
      
      // Get Annual Balance range (Column D = 4)
      const balanceRange = dbSheet.getRange(userRow, 4); 
      const currentBalance = parseFloat(balanceRange.getValue()) || 0;
      const newBalance = currentBalance + monthlyAccrual;
      
      balanceRange.setValue(newBalance);
      
      logsSheet.appendRow([
        new Date(), 
        user.name, 
        'SYSTEM', 
        'Monthly Accrual', 
        `Added ${monthlyAccrual.toFixed(2)} days (Rate: ${annualDaysPerYear}/yr). New Balance: ${newBalance.toFixed(2)}`
      ]);

    } catch (e) {
      Logger.log(`Failed to process accrual for ${user.name}: ${e.message}`);
    }
  }
  Logger.log("Finished monthlyLeaveAccrual trigger.");
}

/**
 * REPLACED: Robustly parses a date from CSV, handling strings, numbers, and Date objects.
 */
function parseDate(dateInput) {
  if (!dateInput) return null;
  if (dateInput instanceof Date) return dateInput;

  try {
    // 1. Handle Serial Number (Excel/Sheets style)
    if (typeof dateInput === 'number' && dateInput > 1) {
      const baseDate = new Date(Date.UTC(1899, 11, 30));
      baseDate.setUTCDate(baseDate.getUTCDate() + dateInput);
      return baseDate;
    }
    
    // 2. Handle String with DD/MM/YYYY (with optional time)
    if (typeof dateInput === 'string') {
      // Regex looks for dd/mm/yyyy at the start
      const match = dateInput.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
      if (match) {
        const day = parseInt(match[1], 10);
        const month = parseInt(match[2], 10) - 1; // Months are 0-11
        const year = parseInt(match[3], 10);
        
        // If it has time, try to parse it, otherwise default to 00:00
        let hours = 0, minutes = 0, seconds = 0;
        const timeMatch = dateInput.match(/(\d{1,2}):(\d{2})(?::(\d{2}))?/);
        if (timeMatch) {
           hours = parseInt(timeMatch[1], 10);
           minutes = parseInt(timeMatch[2], 10);
           seconds = timeMatch[3] ? parseInt(timeMatch[3], 10) : 0;
        }
        
        const newDate = new Date(year, month, day, hours, minutes, seconds);
        if (!isNaN(newDate.getTime())) return newDate;
      }
    }

    // 3. Fallback to standard parser (ISO format yyyy-mm-dd)
    const standardDate = new Date(dateInput);
    if (!isNaN(standardDate.getTime())) return standardDate;

    return null; 
  } catch(e) {
    Logger.log("Date Parse Error: " + e.message);
    return null;
  }
}

/**
 * NEW: Robustly parses a time from CSV, handling strings and serial numbers (fractions).
 * Returns a string in HH:mm:ss format.
 */
function parseCsvTime(timeInput, timeZone) {
  if (timeInput === null || timeInput === undefined || timeInput === "") return ""; // Allow empty time

  try {
    // Check if it's a serial number (e.g., 0.5 for 12:00 PM)
    if (typeof timeInput === 'number' && timeInput >= 0 && timeInput <= 1) { // 1.0 is 24:00, which is 00:00
      // Handle edge case 1.0 = 00:00:00
      if (timeInput === 1) return "00:00:00"; 
      
      const totalSeconds = Math.round(timeInput * 86400);
      const hours = Math.floor(totalSeconds / 3600);
      const minutes = Math.floor((totalSeconds % 3600) / 60);
      const seconds = totalSeconds % 60;
      
      const hh = String(hours).padStart(2, '0');
      const mm = String(minutes).padStart(2, '0');
      const ss = String(seconds).padStart(2, '0');
      
      return `${hh}:${mm}:${ss}`;
    }

    // Check if it's a string (e.g., "12:00" or "12:00:00" or "12:00 PM")
    if (typeof timeInput === 'string') {
      // Try parsing as a date (handles "12:00 PM", "12:00", "12:00:00")
      const dateFromTime = new Date('1970-01-01 ' + timeInput);
      if (!isNaN(dateFromTime.getTime())) {
          return Utilities.formatDate(dateFromTime, timeZone, "HH:mm:ss");
      }
    }
    
    // Check if it's a full Date object (e.g., from a formatted cell)
    if (timeInput instanceof Date) {
      return Utilities.formatDate(timeInput, timeZone, "HH:mm:ss");
    }
    
    return ""; // Could not parse
  } catch(e) {
    Logger.log(`parseCsvTime Error for input "${timeInput}": ${e.message}`);
    return ""; // Return empty on error
  }
}

// ==========================================
// === PHASE 2: EMPLOYEE SELF-SERVICE API ===
// ==========================================

/**
 * Fetches full profile data (Core + PII) for the logged-in user.
 */
function webGetMyProfile() {
  const userEmail = Session.getActiveUser().getEmail().toLowerCase();
  const ss = getSpreadsheet();
  const userData = getUserDataFromDb(ss); // This uses your updated Phase 1 logic
  
  // Find the user object from the list we already generated
  const user = userData.userList.find(u => u.email === userEmail);
  if (!user) throw new Error("User not found.");

  // Now fetch PII data (Phone, Address, IBAN) from the restricted sheet
  const piiSheet = getOrCreateSheet(ss, SHEET_NAMES.employeesPII);
  const piiData = piiSheet.getDataRange().getValues();
  
  let piiRecord = {};
  
  // Look for the row with the matching EmployeeID
  for (let i = 1; i < piiData.length; i++) {
    if (piiData[i][0] === user.empID) { // Column A is EmployeeID
      piiRecord = {
        salary: piiData[i][2],      // Col C
        iban: piiData[i][3],        // Col D
        address: piiData[i][4],     // Col E
        phone: piiData[i][5],       // Col F
        medical: piiData[i][6],     // Col G
        contract: piiData[i][7]     // Col H
      };
      break;
    }
  }

  return {
    core: user,
    pii: piiRecord
  };
}

/**
 * Updates editable profile fields (Address, Phone).
 * Sensitive fields like IBAN trigger a request (simulated for now).
 */
function webUpdateProfile(formData) {
  const userEmail = Session.getActiveUser().getEmail().toLowerCase();
  const ss = getSpreadsheet();
  const userData = getUserDataFromDb(ss);
  const user = userData.userList.find(u => u.email === userEmail);
  if (!user) throw new Error("User not found.");

  const piiSheet = getOrCreateSheet(ss, SHEET_NAMES.employeesPII);
  const piiData = piiSheet.getDataRange().getValues();
  
  let rowToUpdate = -1;
  for (let i = 1; i < piiData.length; i++) {
    if (piiData[i][0] === user.empID) {
      rowToUpdate = i + 1;
      break;
    }
  }

  if (rowToUpdate === -1) throw new Error("PII record not found. Contact HR.");

  // Update Address (Col E -> 5) and Phone (Col F -> 6)
  if (formData.address) piiSheet.getRange(rowToUpdate, 5).setValue(formData.address);
  if (formData.phone) piiSheet.getRange(rowToUpdate, 6).setValue(formData.phone);

  // Logic for IBAN change request (For now, we just log it)
  if (formData.iban && formData.iban !== piiData[rowToUpdate-1][3]) {
     const logsSheet = getOrCreateSheet(ss, SHEET_NAMES.logs);
     logsSheet.appendRow([new Date(), user.name, userEmail, "Profile Change Request", `Requested IBAN change to: ${formData.iban}`]);
     return "Profile updated. Note: IBAN changes require HR approval and have been logged as a request.";
  }

  return "Profile updated successfully.";
}

// 1. Submit a Change Request (Agent)
function webSubmitDataChangeRequest(field, newValue, reason) {
  const { userEmail, userName, ss } = getAuthorizedContext(null);
  const reqSheet = getOrCreateSheet(ss, "Data_Change_Requests");
  
  // Get current value for logging (simplified)
  // In a real scenario, we'd fetch the specific field from PII/Core
  const oldValue = "Current Value"; 

  const reqID = `CHG-${new Date().getTime()}`;
  reqSheet.appendRow([
    reqID,
    userEmail,
    userName,
    field,
    oldValue,
    newValue,
    reason,
    "Pending",
    new Date(),
    "", // ActionBy
    ""  // ActionDate
  ]);
  
  return "Change request submitted to HR.";
}

// 2. Get Pending Requests (HR/Admin)
function webGetDataChangeRequests() {
  const { userEmail, userData, ss } = getAuthorizedContext('OFFBOARD_EMPLOYEE'); // Reusing HR permission
  const reqSheet = getOrCreateSheet(ss, "Data_Change_Requests");
  const data = reqSheet.getDataRange().getValues();
  const requests = [];
  
  for (let i = 1; i < data.length; i++) {
    // Col H (index 7) is Status
    if (data[i][7] === 'Pending') {
      requests.push({
        id: data[i][0],
        email: data[i][1],
        name: data[i][2],
        field: data[i][3],
        oldVal: data[i][4],
        newVal: data[i][5],
        reason: data[i][6],
        date: convertDateToString(new Date(data[i][8]))
      });
    }
  }
  return requests;
}

// 3. Approve/Deny Request (HR/Admin)
function webActionDataChangeRequest(reqId, action) {
  const { userEmail: adminEmail, ss, userData } = getAuthorizedContext('OFFBOARD_EMPLOYEE');
  const reqSheet = getOrCreateSheet(ss, "Data_Change_Requests");
  const data = reqSheet.getDataRange().getValues();
  
  let rowIndex = -1;
  let reqData = null;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === reqId) {
      rowIndex = i + 1;
      reqData = {
        email: data[i][1],
        field: data[i][3],
        newValue: data[i][5]
      };
      break;
    }
  }
  
  if (rowIndex === -1) throw new Error("Request not found.");
  
  if (action === 'Approved') {
    // PERFORM THE UPDATE
    if (reqData.field === 'IBAN' || reqData.field === 'NationalID' || reqData.field === 'Address' || reqData.field === 'Phone') {
       updateEmployeePIIField(ss, userData, reqData.email, reqData.field, reqData.newValue);
    } else {
       // Handle Core fields if necessary
    }
  }
  
  // Update Request Status
  reqSheet.getRange(rowIndex, 8).setValue(action); // Status
  reqSheet.getRange(rowIndex, 10).setValue(adminEmail); // ActionBy
  reqSheet.getRange(rowIndex, 11).setValue(new Date()); // ActionDate
  
  return `Request ${action}.`;
}

// Helper to update PII Sheet
function updateEmployeePIIField(ss, userData, targetEmail, fieldName, newValue) {
  const piiSheet = getOrCreateSheet(ss, SHEET_NAMES.employeesPII);
  const piiData = piiSheet.getDataRange().getValues();
  const headers = piiData[0];
  const colIndex = headers.indexOf(fieldName); // e.g. "IBAN"
  
  if (colIndex === -1) throw new Error(`Field ${fieldName} not found in PII sheet.`);
  
  const empID = userData.userList.find(u => u.email === targetEmail)?.empID;
  if (!empID) throw new Error("User ID not found.");
  
  for (let i = 1; i < piiData.length; i++) {
    if (piiData[i][0] === empID) {
      piiSheet.getRange(i + 1, colIndex + 1).setValue(newValue);
      return;
    }
  }
  throw new Error("PII Record not found for user.");
}

/**
 * Scans the user's specific Drive folder for documents.
 */
function webGetMyDocuments() {
  const userEmail = Session.getActiveUser().getEmail().toLowerCase();
  const ss = getSpreadsheet();
  const userData = getUserDataFromDb(ss);
  const user = userData.userList.find(u => u.email === userEmail);
  
  if (!user || !user.empID) return [];

  // 1. Find the root folder
  const rootFolders = DriveApp.getFoldersByName("KOMPASS_HR_Files");
  if (!rootFolders.hasNext()) return [];
  const root = rootFolders.next();
  
  const empFolders = root.getFoldersByName("Employee_Files");
  if (!empFolders.hasNext()) return [];
  const parentFolder = empFolders.next();

  // 2. Find the specific user folder: "[Name]_[ID]"
  const searchName = `${user.name}_${user.empID}`;
  const userFolders = parentFolder.getFoldersByName(searchName);
  
  if (!userFolders.hasNext()) return [];
  const myFolder = userFolders.next();

  // 3. Recursive function to get all files
  let fileList = [];
  
  function scanFolder(folder, path) {
    const files = folder.getFiles();
    while (files.hasNext()) {
      const f = files.next();
      fileList.push({
        name: f.getName(),
        url: f.getUrl(),
        type: path, // e.g., "Payslips" or "Root"
        date: f.getLastUpdated().toISOString()
      });
    }
    const subFolders = folder.getFolders();
    while (subFolders.hasNext()) {
      const sub = subFolders.next();
      scanFolder(sub, sub.getName());
    }
  }
  
  scanFolder(myFolder, "General");
  return fileList;
}

/**
 * Fetches warnings for the logged-in user.
 */
function webGetMyWarnings() {
  const userEmail = Session.getActiveUser().getEmail().toLowerCase();
  const ss = getSpreadsheet();
  const userData = getUserDataFromDb(ss);
  const user = userData.userList.find(u => u.email === userEmail);
  
  if (!user) return [];

  const wSheet = getOrCreateSheet(ss, "Warnings"); // Ensure this matches SHEET_NAMES
  const data = wSheet.getDataRange().getValues();
  const warnings = [];

  for (let i = 1; i < data.length; i++) {
    // Col B is EmployeeID
    if (data[i][1] === user.empID) {
      warnings.push({
        type: data[i][2],
        level: data[i][3],
        date: convertDateToString(new Date(data[i][4])),
        description: data[i][5],
        status: data[i][6]
      });
    }
  }
  return warnings;
}

// ==========================================
// === PHASE 3: PROJECT MANAGEMENT API ===
// ==========================================

/**
 * Fetches all active projects.
 * Returns a list for dropdowns.
 */
function webGetProjects() {
  const ss = getSpreadsheet();
  const pSheet = getOrCreateSheet(ss, SHEET_NAMES.projects); // Defined in Phase 1
  const data = pSheet.getDataRange().getValues();
  
  const projects = [];
  // Skip header (row 0)
  for (let i = 1; i < data.length; i++) {
    // ProjectID(0), Name(1), Manager(2), Roles(3)
    if (data[i][0]) {
      projects.push({
        id: data[i][0],
        name: data[i][1],
        manager: data[i][2]
      });
    }
  }
  return projects;
}

/**
 * Admins create/update projects here.
 */
function webSaveProject(projectData) {
  const userEmail = Session.getActiveUser().getEmail().toLowerCase();
  const ss = getSpreadsheet();
  
  // Security Check (Admin Only)
  // You can reuse your existing checkAdmin() helper logic here if you extracted it, 
  // or just look up the role again.
  const coreSheet = getOrCreateSheet(ss, SHEET_NAMES.employeesCore);
  const users = coreSheet.getDataRange().getValues();
  let isAdmin = false;
  for(let i=1; i<users.length; i++) {
    if(users[i][2] == userEmail && (users[i][3] == 'admin' || users[i][3] == 'superadmin')) {
      isAdmin = true; break;
    }
  }
  if (!isAdmin) throw new Error("Permission denied.");

  const pSheet = getOrCreateSheet(ss, SHEET_NAMES.projects);
  
  // Generate ID if new
  const pid = projectData.id || `PRJ-${new Date().getTime()}`;
  
  // Check if updating existing
  const data = pSheet.getDataRange().getValues();
  let rowToUpdate = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === pid) {
      rowToUpdate = i + 1;
      break;
    }
  }

  if (rowToUpdate > 0) {
    pSheet.getRange(rowToUpdate, 2).setValue(projectData.name);
    pSheet.getRange(rowToUpdate, 3).setValue(projectData.manager);
  } else {
    pSheet.appendRow([pid, projectData.name, projectData.manager, "All"]);
  }
  
  return "Project saved successfully.";
}

// ==========================================
// === PHASE 4: RECRUITMENT & HIRING API ===
// ==========================================

/**
 * 3. SUBMIT APPLICATION (Public) - UPGRADED
 * Now captures National ID, Languages, Referrer, etc.
 */
function webSubmitApplication(data) {
  const ss = getSpreadsheet();
  const sheet = getOrCreateSheet(ss, SHEET_NAMES.recruitment);
  
  const id = `CAND-${new Date().getTime()}`;
  sheet.appendRow([
    id,
    data.name,
    data.email,
    data.phone,
    data.position, // This might now be a Requisition ID or Title
    data.cv,
    "New",         // Status
    "Applied",     // Stage
    "", "", "", "",// Interview Scores/Notes (Placeholders)
    new Date(),    // Applied Date
    // --- NEW PHASE 3 COLUMNS ---
    data.nationalId || "",
    data.langLevel || "",
    data.secondLang || "",
    data.referrer || "",
    "", "", "", "", // Feedback Columns (HR, Mgmt, Tech, Client)
    "Pending"       // Offer Status
  ]);
  return "Success";
}

/**
 * ADMIN: Gets candidates from Internal DB AND External Buffer
 */
function webGetCandidates() {
  const ss = getSpreadsheet();
  const internalSheet = getOrCreateSheet(ss, SHEET_NAMES.recruitment);
  const candidates = [];

  // 1. Fetch Internal Candidates (Historical/Processing)
  const internalData = internalSheet.getDataRange().getValues();
  for (let i = 1; i < internalData.length; i++) {
    candidates.push({
      id: internalData[i][0],
      name: internalData[i][1],
      email: internalData[i][2],
      position: internalData[i][4],
      cv: internalData[i][5],
      status: internalData[i][6],
      stage: internalData[i][7],
      date: convertDateToString(new Date(internalData[i][12])),
      source: 'Internal'
    });
  }

  // 2. Fetch New Candidates from External Buffer
  try {
    const bufferSs = SpreadsheetApp.openById(BUFFER_SHEET_ID);
    const bufferSheet = bufferSs.getSheets()[0];
    const bufferData = bufferSheet.getDataRange().getValues();
    
    // Start loop from 1 to skip headers
    for (let i = 1; i < bufferData.length; i++) {
      // Buffer Columns: ID(0), Name(1), Email(2), Phone(3), Pos(4), CV(5), Status(6), Date(7)
      // We only show "New" ones. Processed ones should be moved/deleted.
      candidates.push({
        id: bufferData[i][0],
        name: bufferData[i][1],
        email: bufferData[i][2],
        position: bufferData[i][4],
        cv: bufferData[i][5],
        status: "New (External)", // Mark as new
        stage: "Applied",
        date: convertDateToString(new Date(bufferData[i][7])),
        source: 'Buffer',
        phone: bufferData[i][3] // Store for importing
      });
    }
  } catch (e) {
    Logger.log("Could not read buffer sheet (permissions?): " + e.message);
  }

  // Sort by newest
  return candidates.reverse();
}

/**
 * ADMIN: Updates Candidate. 
 * If source is Buffer, it IMPORTS them to Internal DB first.
 */
function webUpdateCandidateStatus(candidateId, newStatus, newStage) {
  const ss = getSpreadsheet();
  const internalSheet = getOrCreateSheet(ss, SHEET_NAMES.recruitment);
  
  // 1. Check if candidate is already Internal
  const internalData = internalSheet.getDataRange().getValues();
  for (let i = 1; i < internalData.length; i++) {
    if (internalData[i][0] === candidateId) {
      if (newStatus) internalSheet.getRange(i + 1, 7).setValue(newStatus);
      if (newStage) internalSheet.getRange(i + 1, 8).setValue(newStage);
      return "Updated";
    }
  }

  // 2. If not found, check External Buffer and Import
  try {
    const bufferSs = SpreadsheetApp.openById(BUFFER_SHEET_ID);
    const bufferSheet = bufferSs.getSheets()[0];
    const bufferData = bufferSheet.getDataRange().getValues();
    
    for (let i = 1; i < bufferData.length; i++) {
      if (bufferData[i][0] === candidateId) {
        // FOUND in Buffer! Import to Internal.
        const row = bufferData[i];
        
        internalSheet.appendRow([
          row[0], // ID
          row[1], // Name
          row[2], // Email
          row[3], // Phone
          row[4], // Position
          row[5], // CV
          newStatus || row[6], // New Status
          newStage || "Applied", // New Stage
          "", "", "", "", 
          row[7] // Date
        ]);
        
        // Remove from Buffer to prevent duplicates
        bufferSheet.deleteRow(i + 1);
        return "Imported & Updated";
      }
    }
  } catch (e) {
    throw new Error("Error importing from buffer: " + e.message);
  }
  
  throw new Error("Candidate not found in Internal DB or Buffer.");
}

/**
 * ADMIN: HIRES A CANDIDATE
 * 1. Creates entry in Employees_Core
 * 2. Creates entry in Employees_PII
 * 3. Creates Google Drive Folders
 * 4. Updates Candidate status to "Hired"
 */
function webHireCandidate(candidateId, hiringData) {
  const ss = getSpreadsheet();
  const candSheet = getOrCreateSheet(ss, SHEET_NAMES.recruitment);
  const coreSheet = getOrCreateSheet(ss, SHEET_NAMES.employeesCore);
  const piiSheet = getOrCreateSheet(ss, SHEET_NAMES.employeesPII);

  // A. Find Candidate
  const candData = candSheet.getDataRange().getValues();
  let candRow = -1;
  let candidate = null;
  
  // Dynamically map headers
  const candHeaders = candData[0];
  const cIdx = {
    id: candHeaders.indexOf("CandidateID"),
    name: candHeaders.indexOf("Name"),
    email: candHeaders.indexOf("Email"),
    phone: candHeaders.indexOf("Phone"),
    natId: candHeaders.indexOf("NationalID")
  };

  for (let i = 1; i < candData.length; i++) {
    if (candData[i][cIdx.id] === candidateId) {
      candRow = i + 1;
      candidate = {
        name: candData[i][cIdx.name],
        email: candData[i][cIdx.email],
        phone: candData[i][cIdx.phone],
        nationalId: candData[i][cIdx.natId]
      };
      break;
    }
  }
  
  if (!candidate) throw new Error("Candidate not found.");

  // B. Generate Employee ID
  const lastRow = coreSheet.getLastRow();
  const newEmpId = `KOM-${1000 + lastRow}`;

  // C. Create CORE Record (Active Status - Skips Registration)
  coreSheet.appendRow([
    newEmpId,
    hiringData.fullName || candidate.name,
    hiringData.konectaEmail,
    'agent',
    'Active', // Auto-active
    hiringData.directManager,
    hiringData.functionalManager,
    0, 0, 0, // Balances
    hiringData.gender,
    hiringData.empType,
    hiringData.contractType,
    hiringData.jobLevel,
    hiringData.department,
    hiringData.function,
    hiringData.subFunction,
    hiringData.gcm,
    hiringData.scope,
    hiringData.shore,
    hiringData.dottedManager,
    hiringData.projectManager,
    hiringData.bonusPlan,
    hiringData.nLevel,
    "", 
    "Active"
  ]);

  // D. Create PII Record (With Basic + Variable Split)
  piiSheet.appendRow([
    newEmpId,
    hiringData.hiringDate,
    hiringData.salary, // Total Salary
    hiringData.iban,
    hiringData.address,
    candidate.phone,
    "", "", 
    candidate.nationalId,
    hiringData.passport,
    hiringData.socialInsurance,
    hiringData.birthDate,
    candidate.email,
    hiringData.maritalStatus,
    hiringData.dependents,
    hiringData.emergencyContact,
    hiringData.emergencyRelation,
    hiringData.salary, 
    hiringData.hourlyRate,
    hiringData.variable // New Variable Pay Column
  ]);

  // E. Create Drive Folders
  try {
    const rootFolders = DriveApp.getFoldersByName("KOMPASS_HR_Files");
    if (rootFolders.hasNext()) {
      const root = rootFolders.next();
      const empFolders = root.getFoldersByName("Employee_Files");
      if (empFolders.hasNext()) {
        const parent = empFolders.next();
        const personalFolder = parent.createFolder(`${candidate.name}_${newEmpId}`);
        personalFolder.createFolder("Payslips");
        personalFolder.createFolder("Onboarding_Docs");
        personalFolder.createFolder("Sick_Notes");
      }
    }
  } catch (e) { Logger.log("Folder creation error: " + e.message); }

  // F. Update Candidate Status
  candSheet.getRange(candRow, 7).setValue("Hired"); // Status
  candSheet.getRange(candRow, 8).setValue("Onboarding"); // Stage

  return `Successfully hired ${candidate.name}. Employee ID: ${newEmpId}`;
}

/**
 * ======================================================================
 * PHASE 5 DATABASE UPGRADE SCRIPT (FIXED)
 * ACTION: RUN THIS FUNCTION AGAIN.
 * PURPOSE: Expands existing sheets and creates new ones for the HRIS system.
 * ======================================================================
 */
function _SETUP_PHASE_5_DATABASE() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  Logger.log("Starting Phase 5 Database Upgrade...");

  // --- 1. CREATE NEW SHEETS ---
  
  // 1.1 Requisitions (Job Openings)
  const reqSheet = getOrCreateSheet(ss, SHEET_NAMES.requisitions);
  // FIX: Check if lastRow is 0 (new sheet) OR 1 (potentially empty header)
  if (reqSheet.getLastRow() === 0 || (reqSheet.getLastRow() === 1 && reqSheet.getRange("A1").getValue() === "")) {
    reqSheet.getRange("A1:H1").setValues([[
      "ReqID", "Title", "Department", "HiringManager", "OpenDate", 
      "Status", "PoolCandidates", "JobDescription"
    ]]);
    reqSheet.setFrozenRows(1);
    Logger.log("Created 'Requisitions' sheet with headers.");
  }

  // 1.2 Performance Reviews
  const perfSheet = getOrCreateSheet(ss, SHEET_NAMES.performance);
  if (perfSheet.getLastRow() === 0 || (perfSheet.getLastRow() === 1 && perfSheet.getRange("A1").getValue() === "")) {
    perfSheet.getRange("A1:G1").setValues([[
      "ReviewID", "EmployeeID", "Year", "ReviewPeriod", "Rating", 
      "ManagerComments", "Date"
    ]]);
    perfSheet.setFrozenRows(1);
    Logger.log("Created 'Performance_Reviews' sheet with headers.");
  }

  // 1.3 Employee History (Promotions/Transfers)
  const histSheet = getOrCreateSheet(ss, SHEET_NAMES.historyLogs);
  if (histSheet.getLastRow() === 0 || (histSheet.getLastRow() === 1 && histSheet.getRange("A1").getValue() === "")) {
    histSheet.getRange("A1:F1").setValues([[
      "HistoryID", "EmployeeID", "Date", "EventType", 
      "OldValue", "NewValue"
    ]]);
    histSheet.setFrozenRows(1);
    Logger.log("Created 'Employee_History' sheet with headers.");
  }

  // --- 2. EXPAND EXISTING SHEETS ---

  // 2.1 Expand Employees_Core
  const coreSheet = ss.getSheetByName(SHEET_NAMES.employeesCore);
  if (coreSheet) {
    const newCoreCols = [
      "Gender", "EmploymentType", "ContractType", "JobLevel", "Department",
      "Function", "SubFunction", "GCMLevel", "Scope", "OffshoreOnshore",
      "DottedManager", "ProjectManagerEmail", "BonusPlan", "N_Level", 
      "ExitDate", "Status" 
    ];
    addColumnsToSheet(coreSheet, newCoreCols);
    Logger.log("Updated 'Employees_Core' with new HR columns.");
  } else {
    Logger.log("ERROR: Employees_Core sheet not found. Run Phase 1 setup first.");
  }

  // 2.2 Expand Employees_PII
  const piiSheet = ss.getSheetByName(SHEET_NAMES.employeesPII);
  if (piiSheet) {
    const newPiiCols = [
      "NationalID", "PassportNumber", "SocialInsuranceNumber", "BirthDate",
      "PersonalEmail", "MaritalStatus", "DependentsInfo", "EmergencyContact",
      "EmergencyRelation", "Salary", "HourlyRate"
    ];
    addColumnsToSheet(piiSheet, newPiiCols);
    
    // Set Date Format for BirthDate column
    try {
      const headers = piiSheet.getRange(1, 1, 1, piiSheet.getLastColumn()).getValues()[0];
      const dobIndex = headers.indexOf("BirthDate") + 1;
      if (dobIndex > 0) piiSheet.getRange(2, dobIndex, piiSheet.getMaxRows(), 1).setNumberFormat("yyyy-mm-dd");
    } catch (e) {
      Logger.log("Could not set date format (sheet might be empty): " + e.message);
    }
    
    Logger.log("Updated 'Employees_PII' with new sensitive columns.");
  }

  // 2.3 Update Recruitment_Candidates
  const recSheet = ss.getSheetByName(SHEET_NAMES.recruitment);
  if (recSheet) {
    const newRecCols = [
      "NationalID", "LanguageLevel", "SecondLanguage", "Referrer", 
      "HR_Feedback", "Management_Feedback", "Technical_Feedback", 
      "Client_Feedback", "OfferStatus"
    ];
    addColumnsToSheet(recSheet, newRecCols);
    Logger.log("Updated 'Recruitment_Candidates' with feedback columns.");
  }

  Logger.log("Phase 5 Database Upgrade Complete!");
}

/**
 * HELPER: specific to this upgrade script.
 * Adds missing columns to the end of a sheet's header row.
 * FIX: Handles empty sheets correctly.
 */
function addColumnsToSheet(sheet, newHeaders) {
  const lastCol = sheet.getLastColumn();

  // Case 1: Sheet is completely empty
  if (lastCol === 0) {
    if (newHeaders.length > 0) {
      sheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
    }
    return;
  }

  // Case 2: Sheet has existing data, append only new columns
  const currentHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const headersToAdd = [];

  newHeaders.forEach(header => {
    if (!currentHeaders.includes(header)) {
      headersToAdd.push(header);
    }
  });

  if (headersToAdd.length > 0) {
    // Append to the next available column
    sheet.getRange(1, lastCol + 1, 1, headersToAdd.length).setValues([headersToAdd]);
  }
}
function debugDatabaseMapping() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Employees_Core");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  Logger.log("--- DEBUGGING HEADERS ---");
  Logger.log("All Headers: " + headers.join(", "));
  
  const dmIndex = headers.indexOf("DirectManagerEmail");
  const pmIndex = headers.indexOf("ProjectManagerEmail");
  
  Logger.log(`DirectManagerEmail Index: ${dmIndex} (Should be > -1)`);
  Logger.log(`ProjectManagerEmail Index: ${pmIndex} (Should be > -1)`);
  
  if (dmIndex === -1 || pmIndex === -1) {
    Logger.log("❌ CRITICAL ERROR: One or both manager headers are missing or misspelled!");
    return;
  }

  // Check the first user row (Row 2)
  if (data.length > 1) {
    const row = data[1];
    Logger.log("--- SAMPLE USER DATA (Row 2) ---");
    Logger.log(`Name: ${row[headers.indexOf("Name")]}`);
    Logger.log(`Email: ${row[headers.indexOf("Email")]}`);
    Logger.log(`Direct Manager Value: '${row[dmIndex]}'`);
    Logger.log(`Project Manager Value: '${row[pmIndex]}'`);
  }
}

// ==========================================
// === PHASE 3: RECRUITMENT & ONBOARDING  ===
// ==========================================

/**
 * 1. CREATE REQUISITION (Admin)
 * Opens a new job position in the 'Requisitions' sheet.
 */
function webCreateRequisition(data) {
  const ss = getSpreadsheet();
  const reqSheet = getOrCreateSheet(ss, SHEET_NAMES.requisitions);
  const reqID = `REQ-${new Date().getTime()}`; // Unique Job ID
  
  reqSheet.appendRow([
    reqID,
    data.title,
    data.department,
    data.hiringManager, // Email of the manager
    new Date(),         // Open Date
    "Open",             // Status
    "",                 // Pool Candidates (Empty start)
    data.description
  ]);
  return "Requisition opened successfully: " + reqID;
}

/**
 * 2. GET OPEN REQUISITIONS (Public & Admin)
 * Returns list of open jobs for the dropdown in Recruitment.html
 */
function webGetOpenRequisitions() {
  const ss = getSpreadsheet();
  const reqSheet = getOrCreateSheet(ss, SHEET_NAMES.requisitions);
  const data = reqSheet.getDataRange().getValues();
  const jobs = [];
  
  for (let i = 1; i < data.length; i++) {
    // Col F (Index 5) is Status
    if (data[i][5] === 'Open') {
      jobs.push({
        id: data[i][0],
        title: data[i][1],
        dept: data[i][2]
      });
    }
  }
  return jobs;
}

// ==========================================
// === PHASE 4: PROFILE & SELF-SERVICE API ===
// ==========================================

/**
 * 1. GET FULL PROFILE (Core + PII)
 * Fetches all data points for the "My Profile" tab.
 */
function webGetMyProfile() {
  const userEmail = Session.getActiveUser().getEmail().toLowerCase();
  const ss = getSpreadsheet();
  const userData = getUserDataFromDb(ss); // This already reads Core columns
  
  // Find user in the loaded list
  const userCore = userData.userList.find(u => u.email === userEmail);
  if (!userCore) throw new Error("User profile not found.");

  // Fetch Extended PII Data
  const piiSheet = getOrCreateSheet(ss, SHEET_NAMES.employeesPII);
  const piiData = piiSheet.getDataRange().getValues();
  let piiRecord = {};

  // Find PII row by EmployeeID
  for (let i = 1; i < piiData.length; i++) {
    if (piiData[i][0] === userCore.empID) {
      piiRecord = {
        hiringDate: convertDateToString(parseDate(piiData[i][1])),
        salary: piiData[i][2],       // Confidential
        iban: piiData[i][3],         // Confidential
        address: piiData[i][4],
        phone: piiData[i][5],
        medical: piiData[i][6],
        contractLink: piiData[i][7],
        nationalId: piiData[i][8],   // New Phase 5 Col
        passport: piiData[i][9],
        socialInsurance: piiData[i][10],
        birthDate: convertDateToString(parseDate(piiData[i][11])),
        personalEmail: piiData[i][12],
        maritalStatus: piiData[i][13],
        dependents: piiData[i][14],
        emergencyContact: piiData[i][15],
        emergencyRelation: piiData[i][16]
      };
      break;
    }
  }

  // Calculate Age
  let age = "N/A";
  if (piiRecord.birthDate) {
    const dob = new Date(piiRecord.birthDate);
    const diff_ms = Date.now() - dob.getTime();
    const age_dt = new Date(diff_ms); 
    age = Math.abs(age_dt.getUTCFullYear() - 1970);
  }

  // Fetch additional Core fields that getUserDataFromDb might not have exposed in the simplified list
  // We can re-read the row from the Core Sheet directly to be safe, or rely on getUserDataFromDb if we updated it fully.
  // Let's just return what we have, assuming getUserDataFromDb is robust.
  // If you find fields missing, we can add a direct read here.

  return {
    core: {
      ...userCore, // Includes Name, ID, Role, Managers, Balances
      // You might need to explicitly map the new Phase 5 Core columns if getUserDataFromDb doesn't return them in the object
      // For now, let's assume basic data. If you need specifically "JobLevel" or "GCM", we should ensure getUserDataFromDb returns them.
    },
    pii: {
      ...piiRecord,
      age: age
    }
  };
}

/**
 * 2. UPDATE PROFILE (Self-Service)
 * Allows users to update: Phone, Address, Emergency Contact, Personal Email.
 * Sensitive fields (IBAN, Name) trigger a request log.
 */
function webUpdateProfile(formData) {
  const userEmail = Session.getActiveUser().getEmail().toLowerCase();
  const ss = getSpreadsheet();
  const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
  const userData = getUserDataFromDb(dbSheet);
  
  const user = userData.userList.find(u => u.email === userEmail);
  if (!user) throw new Error("User not found.");

  const piiSheet = getOrCreateSheet(ss, SHEET_NAMES.employeesPII);
  const piiData = piiSheet.getDataRange().getValues();
  let rowToUpdate = -1;

  for (let i = 1; i < piiData.length; i++) {
    if (piiData[i][0] === user.empID) {
      rowToUpdate = i + 1;
      break;
    }
  }
  if (rowToUpdate === -1) throw new Error("PII record not found.");

  // Update Allowed Fields
  // Address (Col E = 5)
  if (formData.address) piiSheet.getRange(rowToUpdate, 5).setValue(formData.address);
  // Phone (Col F = 6)
  if (formData.phone) piiSheet.getRange(rowToUpdate, 6).setValue(formData.phone);
  // Personal Email (Col M = 13)
  if (formData.personalEmail) piiSheet.getRange(rowToUpdate, 13).setValue(formData.personalEmail);
  // Emergency Contact (Col P = 16)
  if (formData.emergencyContact) piiSheet.getRange(rowToUpdate, 16).setValue(formData.emergencyContact);
  // Emergency Relation (Col Q = 17)
  if (formData.emergencyRelation) piiSheet.getRange(rowToUpdate, 17).setValue(formData.emergencyRelation);

  // Log Restricted Changes (IBAN, Marital Status)
  const logsSheet = getOrCreateSheet(ss, SHEET_NAMES.logs);
  
  // Check IBAN change (Col D = 4)
  const currentIBAN = piiData[rowToUpdate-1][3];
  if (formData.iban && formData.iban !== String(currentIBAN)) {
    logsSheet.appendRow([new Date(), user.name, userEmail, "Data Change Request", `Requested IBAN change to: ${formData.iban}`]);
    return "Profile updated. Note: IBAN change has been sent to HR for approval.";
  }

  return "Profile updated successfully.";
}

// ==========================================
// === PHASE 5: PERFORMANCE & OFFBOARDING ===
// ==========================================

// 3. Updated Performance Review
function webSubmitPerformanceReview(reviewData) {
  // Checks if user has permission to submit reviews
  const { userEmail: adminEmail, userData, ss } = getAuthorizedContext('SUBMIT_PERFORMANCE');
  
  const targetEmail = reviewData.employeeEmail.toLowerCase();
  
  // Contextual Check: Can only review OWN team (unless Superadmin)
  const targetSupervisor = userData.emailToSupervisor[targetEmail];
  const targetProjectMgr = userData.emailToProjectManager[targetEmail];
  const adminRole = userData.emailToRole[adminEmail];

  const isAuthorized = (adminRole === 'superadmin') || 
                       (targetSupervisor === adminEmail) || 
                       (targetProjectMgr === adminEmail);

  if (!isAuthorized) throw new Error("Permission denied. You can only review your own team members.");

  const targetUser = userData.userList.find(u => u.email === targetEmail);
  if (!targetUser) throw new Error("Employee not found.");

  const perfSheet = getOrCreateSheet(ss, SHEET_NAMES.performance);
  perfSheet.appendRow([
    `REV-${new Date().getTime()}`,
    targetUser.empID,
    reviewData.year,
    reviewData.period,
    reviewData.rating,
    reviewData.comments,
    new Date()
  ]);

  return "Performance review submitted successfully.";
}

/**
 * 2. GET PERFORMANCE HISTORY (Employee/Manager)
 * Returns list of past reviews for a specific user.
 */
function webGetPerformanceHistory(targetEmail) {
  const viewerEmail = Session.getActiveUser().getEmail().toLowerCase();
  const ss = getSpreadsheet();
  const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
  const userData = getUserDataFromDb(dbSheet);
  
  const emailToFetch = targetEmail || viewerEmail;
  const viewerRole = userData.emailToRole[viewerEmail] || 'agent';

  // Security Check: Agents can only see their own. Managers can see team's.
  if (viewerRole === 'agent' && emailToFetch !== viewerEmail) {
    throw new Error("Permission denied.");
  }

  // Get Employee ID
  const targetUser = userData.userList.find(u => u.email === emailToFetch);
  if (!targetUser) return []; // No user found

  const perfSheet = getOrCreateSheet(ss, SHEET_NAMES.performance);
  const data = perfSheet.getDataRange().getValues();
  const reviews = [];

  // Columns: ReviewID(0), EmpID(1), Year(2), Period(3), Rating(4), Comments(5), Date(6)
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === targetUser.empID) {
      reviews.push({
        id: data[i][0],
        year: data[i][2],
        period: data[i][3],
        rating: data[i][4],
        comments: data[i][5],
        date: convertDateToString(new Date(data[i][6]))
      });
    }
  }
  
  return reviews.reverse(); // Newest first
}

// 1. Updated Offboarding
function webOffboardEmployee(offboardData) {
  // Replaces hardcoded check with dynamic RBAC
  const { userEmail: adminEmail, userData, ss } = getAuthorizedContext('OFFBOARD_EMPLOYEE');

  const targetEmail = offboardData.email.toLowerCase();
  const row = userData.emailToRow[targetEmail];
  if (!row) throw new Error("User not found.");

  const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.employeesCore);
  const headers = dbSheet.getRange(1, 1, 1, dbSheet.getLastColumn()).getValues()[0];
  const statusCol = headers.indexOf("Status") + 1;
  const exitDateCol = headers.indexOf("ExitDate") + 1;

  if (statusCol > 0) dbSheet.getRange(row, statusCol).setValue("Left");
  if (exitDateCol > 0) dbSheet.getRange(row, exitDateCol).setValue(offboardData.exitDate);

  // Log History
  const histSheet = getOrCreateSheet(ss, SHEET_NAMES.historyLogs);
  const targetUser = userData.userList.find(u => u.email === targetEmail);
  histSheet.appendRow([
    `HIST-${new Date().getTime()}`,
    targetUser ? targetUser.empID : "UNKNOWN",
    new Date(),
    "Termination/Exit",
    "Active",
    "Left"
  ]);

  return `Successfully offboarded ${targetEmail}. Status set to 'Left'.`;
}

// --- JOB REQUISITION MANAGEMENT ---

function webGetRequisitions(filterStatus) {
  try {
    const ss = getSpreadsheet();
    const sheet = getOrCreateSheet(ss, SHEET_NAMES.requisitions);
    const data = sheet.getDataRange().getValues();
    const jobs = [];
    
    // Skip header
    for (let i = 1; i < data.length; i++) {
      const status = data[i][5];
      if (filterStatus === 'All' || status === filterStatus) {
        jobs.push({
          id: data[i][0],
          title: data[i][1],
          dept: data[i][2],
          manager: data[i][3],
          date: convertDateToString(new Date(data[i][4])),
          status: status,
          desc: data[i][7]
        });
      }
    }
    return jobs.reverse();
  } catch (e) { return { error: e.message }; }
}

function webManageRequisition(reqId, action, newData) {
  try {
    const ss = getSpreadsheet();
    const sheet = getOrCreateSheet(ss, SHEET_NAMES.requisitions);
    const data = sheet.getDataRange().getValues();
    let rowIdx = -1;

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === reqId) { rowIdx = i + 1; break; }
    }
    if (rowIdx === -1) throw new Error("Requisition not found");

    if (action === 'Archive') {
      sheet.getRange(rowIdx, 6).setValue('Archived');
    } else if (action === 'Edit') {
      if(newData.title) sheet.getRange(rowIdx, 2).setValue(newData.title);
      if(newData.dept) sheet.getRange(rowIdx, 3).setValue(newData.dept);
      if(newData.desc) sheet.getRange(rowIdx, 8).setValue(newData.desc);
    }
    return "Success";
  } catch (e) { return "Error: " + e.message; }
}

// --- CANDIDATE WORKFLOW & AUTOMATION ---

function webGetCandidateHistory(email) {
  try {
    const ss = getSpreadsheet();
    const sheet = getOrCreateSheet(ss, SHEET_NAMES.recruitment);
    const data = sheet.getDataRange().getValues();
    const history = [];
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][2]).toLowerCase() === email.toLowerCase()) {
        history.push({
          position: data[i][4],
          date: convertDateToString(new Date(data[i][9])), // AppliedDate
          status: data[i][6],
          stage: data[i][7]
        });
      }
    }
    return history;
  } catch (e) { return []; }
}

function webSendRejectionEmail(candidateId, reason, sendEmail) {
  try {
    const ss = getSpreadsheet();
    const sheet = getOrCreateSheet(ss, SHEET_NAMES.recruitment);
    const data = sheet.getDataRange().getValues();
    let rowIdx = -1;
    let candidate = {};

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === candidateId) { 
        rowIdx = i + 1; 
        candidate = { name: data[i][1], email: data[i][2], pos: data[i][4] };
        break; 
      }
    }
    if (rowIdx === -1) throw new Error("Candidate not found");

    // Update Sheet
    // Col 7 = Status, Col 8 = Stage, Col 20 = RejectionReason (New)
    sheet.getRange(rowIdx, 7).setValue("Rejected");
    sheet.getRange(rowIdx, 8).setValue("Disqualified");
    // Assuming RejectionReason is column 20 (Index 19) based on fixer schema
    // Dynamically find index just in case
    const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
    const reasonIdx = headers.indexOf("RejectionReason");
    if (reasonIdx > -1) sheet.getRange(rowIdx, reasonIdx + 1).setValue(reason);

    if (sendEmail) {
      const subject = `Update regarding your application for ${candidate.pos}`;
      const body = `Dear ${candidate.name},\n\nThank you for your interest in the ${candidate.pos} position at Konecta. After careful consideration, we have decided to move forward with other candidates whose qualifications more closely match our current needs.\n\nWe wish you the best in your job search.\n\nBest regards,\nKonecta HR Team`;
      
      MailApp.sendEmail({ to: candidate.email, subject: subject, body: body });
      return "Rejection recorded & Email sent.";
    }
    return "Rejection recorded (No email sent).";
  } catch (e) { return "Error: " + e.message; }
}

function webSendOfferLetter(candidateId, offerDetails) {
  try {
    const ss = getSpreadsheet();
    const sheet = getOrCreateSheet(ss, SHEET_NAMES.recruitment);
    const data = sheet.getDataRange().getValues();
    let candidate = null;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === candidateId) {
        candidate = { name: data[i][1], email: data[i][2], pos: data[i][4] };
        break;
      }
    }
    if (!candidate) throw new Error("Candidate not found");

    const subject = `Job Offer: ${candidate.pos} at Konecta`;
    const body = `Dear ${candidate.name},\n\nWe are pleased to offer you the position of ${candidate.pos} at Konecta!\n\n` +
                 `**Start Date:** ${offerDetails.startDate}\n` +
                 `**Basic Salary:** ${offerDetails.basic}\n` +
                 `**Variable/Bonus:** ${offerDetails.variable}\n\n` +
                 `Please reply to this email to accept this offer.\n\nBest regards,\nKonecta HR`;

    MailApp.sendEmail({ to: candidate.email, subject: subject, body: body });
    return "Offer letter sent to " + candidate.email;
  } catch (e) { return "Error: " + e.message; }
}


// ==========================================
// === PHASE 6.3: PAYROLL & FINANCE HUB ===
// ==========================================

/**
 * USER: Get My Financial Profile & Entitlements
 */
function webGetMyFinancials() {
  const userEmail = Session.getActiveUser().getEmail().toLowerCase();
  const ss = getSpreadsheet();
  const userData = getUserDataFromDb(ss);
  
  const userCore = userData.userList.find(u => u.email === userEmail);
  if (!userCore) throw new Error("User not found.");

  // 1. Get Salary Breakdown from PII
  const piiSheet = getOrCreateSheet(ss, SHEET_NAMES.employeesPII);
  const piiData = piiSheet.getDataRange().getValues();
  const piiHeaders = piiData[0];
  
  // Map Indexes
  const idx = {
    empId: piiHeaders.indexOf("EmployeeID"),
    basic: piiHeaders.indexOf("BasicSalary"),
    variable: piiHeaders.indexOf("VariablePay"),
    hourly: piiHeaders.indexOf("HourlyRate"),
    total: piiHeaders.indexOf("Salary")
  };

  let salaryInfo = { basic: 0, variable: 0, total: 0 };

  for (let i = 1; i < piiData.length; i++) {
    if (piiData[i][idx.empId] === userCore.empID) {
      salaryInfo = {
        basic: piiData[i][idx.basic] || "Not Set",
        variable: piiData[i][idx.variable] || "Not Set",
        total: piiData[i][idx.total] || "Not Set"
      };
      break;
    }
  }

  // 2. Get Entitlements (Bonuses, Overtime)
  const finSheet = getOrCreateSheet(ss, SHEET_NAMES.financialEntitlements);
  const finData = finSheet.getDataRange().getValues();
  const entitlements = [];

  for (let i = 1; i < finData.length; i++) {
    // Col 1 = EmployeeEmail
    if (String(finData[i][1]).toLowerCase() === userEmail) {
      entitlements.push({
        type: finData[i][3],
        amount: finData[i][4],
        currency: finData[i][5],
        date: convertDateToString(new Date(finData[i][6])), // Due Date
        status: finData[i][7],
        desc: finData[i][8]
      });
    }
  }

  return { salary: salaryInfo, entitlements: entitlements.reverse() };
}

/**
 * ADMIN: Submit a Single Entitlement
 */
function webSubmitEntitlement(data) {
  try {
    const { userEmail: adminEmail, userData, ss } = getAuthorizedContext('MANAGE_FINANCE');
    const sheet = getOrCreateSheet(ss, SHEET_NAMES.financialEntitlements);
    
    const targetEmail = data.email.toLowerCase();
    const userObj = userData.userList.find(u => u.email === targetEmail);
    const targetName = userObj ? userObj.name : targetEmail;
    const id = `FIN-${new Date().getTime()}`;
    
    sheet.appendRow([id, targetEmail, targetName, data.type, data.amount, "EGP", new Date(data.date), "Pending", data.desc, adminEmail, new Date()]);
    return "Entitlement added successfully.";
  } catch (e) { return "Error: " + e.message; }
}

/**
 * ADMIN: Bulk Upload Entitlements via CSV Data
 * Expected CSV: Email, Type, Amount, Date, Description
 */
function webUploadEntitlementsCSV(csvData) {
  try {
    const adminEmail = Session.getActiveUser().getEmail().toLowerCase();
    checkFinancialPermission(adminEmail);

    const ss = getSpreadsheet();
    const sheet = getOrCreateSheet(ss, SHEET_NAMES.financialEntitlements);
    const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.employeesCore);
    const userData = getUserDataFromDb(dbSheet); // To map emails to names

    let count = 0;
    
    csvData.forEach(row => {
      // row is { Email: '...', Type: '...', Amount: ... }
      if (!row.Email || !row.Amount) return;
      
      const targetEmail = row.Email.toLowerCase();
      const userObj = userData.userList.find(u => u.email === targetEmail);
      const targetName = userObj ? userObj.name : targetEmail;
      const id = `FIN-${new Date().getTime()}-${Math.floor(Math.random()*1000)}`;

      sheet.appendRow([
        id,
        targetEmail,
        targetName,
        row.Type || "Bonus",
        row.Amount,
        "EGP",
        new Date(row.Date || new Date()),
        "Pending",
        row.Description || "Bulk Upload",
        adminEmail,
        new Date()
      ]);
      count++;
    });

    return `Successfully processed ${count} records.`;
  } catch (e) { return "Error: " + e.message; }
}

// --- Helper: Permission Check ---
function checkFinancialPermission(email) {
  const ss = getSpreadsheet();
  const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.employeesCore);
  const userData = getUserDataFromDb(dbSheet);
  const role = userData.emailToRole[email];
  
  if (role !== 'financial_manager' && role !== 'superadmin') {
    throw new Error("Permission denied. Financial Manager access required.");
  }
}

/**
 * PHASE 6.5: COACHING HIERARCHY FIX
 * Returns a list of {name, email} for users the current user is allowed to coach.
 * - Superadmin: Returns All Users
 * - Admin/Manager: Returns their full downstream hierarchy (Direct + Indirect)
 * - Agent: Returns empty list
 */
function webGetCoachableUsers() {
  try {
    const userEmail = Session.getActiveUser().getEmail().toLowerCase();
    const ss = getSpreadsheet();
    const userData = getUserDataFromDb(ss);
    const userRole = userData.emailToRole[userEmail];

    let targetEmails = new Set();

    if (userRole === 'superadmin') {
       // Superadmins can coach everyone
       userData.userList.forEach(u => targetEmails.add(u.email));
    } 
    else if (userRole === 'admin' || userRole === 'manager' || userRole === 'financial_manager') {
       // Managers coach their hierarchy
       // Reuse the existing hierarchy walker
       const hierarchyEmails = webGetAllSubordinateEmails(userEmail); 
       hierarchyEmails.forEach(e => targetEmails.add(e));
       
       // Remove the manager themselves from the list (optional, but usually you coach others)
       if (targetEmails.has(userEmail)) targetEmails.delete(userEmail);
    } 
    else {
       return []; // Agents don't coach
    }

    // Map emails to Name/Email objects for the frontend dropdown
    const result = [];
    targetEmails.forEach(email => {
       const u = userData.userList.find(user => user.email === email);
       if (u) {
         result.push({ name: u.name, email: u.email });
       }
    });

    // Sort Alphabetically
    return result.sort((a, b) => a.name.localeCompare(b.name));

  } catch (e) {
    Logger.log("webGetCoachableUsers Error: " + e.message);
    return [];
  }
}

// ==========================================================
// === PHASE 6.6: SMART RBAC ENGINE ===
// ==========================================================

/**
 * 🧠 SMART CONTEXT: The only line you need at the start of a function.
 * Usage: const { userEmail, userData, ss } = getAuthorizedContext('MANAGE_FINANCE');
 */
function getAuthorizedContext(requiredPermission) {
  const userEmail = Session.getActiveUser().getEmail().toLowerCase();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // Use existing helper to get all data
  const userData = getUserDataFromDb(ss);
  const userRole = userData.emailToRole[userEmail] || 'agent';

  // If a permission is required, check it
  if (requiredPermission) {
    const permissionsMap = getPermissionsMap(ss);
    
    // 1. Check if permission exists in DB
    if (!permissionsMap[requiredPermission]) {
      console.warn(`Warning: Permission '${requiredPermission}' not found in RBAC sheet.`);
      throw new Error(`Access Denied: Permission check failed (${requiredPermission}).`);
    }

    // 2. Check if user's role has this permission
    const hasAccess = permissionsMap[requiredPermission][userRole];
    
    if (!hasAccess) {
      throw new Error(`Permission Denied: You need '${requiredPermission}' access.`);
    }
  }

  return { 
    userEmail: userEmail, 
    userName: userData.emailToName[userEmail],
    userRole: userRole,
    userData: userData,
    ss: ss 
  };
}

/**
 * Helper: Reads and Caches the RBAC Sheet
 */
function getPermissionsMap(ss) {
  const cache = CacheService.getScriptCache();
  const cached = cache.get("RBAC_MAP_V1");
  if (cached) return JSON.parse(cached);

  const sheet = getOrCreateSheet(ss, SHEET_NAMES.rbac);
  const data = sheet.getDataRange().getValues();
  const headers = data[0]; // [ID, Desc, superadmin, admin, manager, financial_manager, agent]
  const map = {};

  for (let i = 1; i < data.length; i++) {
    const permID = data[i][0];
    map[permID] = {};
    for (let c = 2; c < headers.length; c++) {
      const role = headers[c];
      map[permID][role] = String(data[i][c]).toUpperCase() === 'TRUE';
    }
  }

  cache.put("RBAC_MAP_V1", JSON.stringify(map), 600); // Cache for 10 mins
  return map;
}



// ==========================================
// === PHASE 7: HR ADMIN & PII TOOLS ===
// ==========================================

/**
 * ADMIN: Search for an employee to edit their PII.
 * Returns Core data merged with PII data.
 */
function webSearchEmployeePII(query) {
  const { userEmail, userData, ss } = getAuthorizedContext('OFFBOARD_EMPLOYEE'); // Reusing a high-level HR permission
  
  const lowerQuery = query.toLowerCase().trim();
  const targetUser = userData.userList.find(u => 
    u.email.includes(lowerQuery) || u.name.toLowerCase().includes(lowerQuery)
  );

  if (!targetUser) throw new Error("User not found.");

  // Fetch PII Data
  const piiSheet = getOrCreateSheet(ss, SHEET_NAMES.employeesPII);
  const piiData = piiSheet.getDataRange().getValues();
  const piiHeaders = piiData[0];
  
  let piiRow = {};
  const empIdIdx = piiHeaders.indexOf("EmployeeID");
  
  for (let i = 1; i < piiData.length; i++) {
    if (piiData[i][empIdIdx] === targetUser.empID) {
      // Map all headers to the row values
      piiHeaders.forEach((header, index) => {
        let value = piiData[i][index];
        // Format dates
        if (value instanceof Date) value = convertDateToString(value).split('T')[0];
        piiRow[header] = value;
      });
      break;
    }
  }

  return {
    core: targetUser,
    pii: piiRow
  };
}

/**
 * ADMIN: Update PII fields for an employee.
 */
function webUpdateEmployeePII(empID, formData) {
  const { userEmail: adminEmail, ss } = getAuthorizedContext('OFFBOARD_EMPLOYEE');
  
  const piiSheet = getOrCreateSheet(ss, SHEET_NAMES.employeesPII);
  const data = piiSheet.getDataRange().getValues();
  const headers = data[0];
  
  let rowIndex = -1;
  // Find row by EmployeeID (Col A)
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === empID) {
      rowIndex = i + 1;
      break;
    }
  }

  if (rowIndex === -1) throw new Error("Employee PII record not found.");

  // Update fields dynamically based on formData keys matching headers
  // We only allow specific editable fields for safety
  const allowedFields = [
    "NationalID", "IBAN", "PassportNumber", "SocialInsuranceNumber", 
    "Address", "Phone", "PersonalEmail", "MaritalStatus", 
    "EmergencyContact", "EmergencyRelation", "BasicSalary", "VariablePay"
  ];

  const updates = [];

  for (const [key, value] of Object.entries(formData)) {
    if (allowedFields.includes(key)) {
      const colIndex = headers.indexOf(key);
      if (colIndex > -1) {
        piiSheet.getRange(rowIndex, colIndex + 1).setValue(value);
        updates.push(`${key}: ${value}`);
      }
    }
  }

  // Log changes
  const logsSheet = getOrCreateSheet(ss, SHEET_NAMES.logs);
  logsSheet.appendRow([
    new Date(),
    `ID: ${empID}`,
    adminEmail,
    "Admin PII Update",
    `Updated: ${updates.join(', ')}`
  ]);

  return "Employee data updated successfully.";
}

/**
 * ADMIN: Get pending data change requests (from Logs).
 */
function webGetPendingDataChanges() {
  const { ss } = getAuthorizedContext('OFFBOARD_EMPLOYEE');
  const logsSheet = getOrCreateSheet(ss, SHEET_NAMES.logs);
  const data = logsSheet.getDataRange().getValues();
  const requests = [];

  // Loop backwards to see newest first
  for (let i = data.length - 1; i > 0; i--) {
    const row = data[i];
    // Look for "Data Change Request" or "Profile Change Request"
    if (row[3] === "Data Change Request" || row[3] === "Profile Change Request") {
      requests.push({
        date: convertDateToString(new Date(row[0])),
        user: row[1],
        email: row[2],
        details: row[4]
      });
    }
    // Limit to last 20 requests to keep it snappy
    if (requests.length >= 20) break;
  }
  return requests;
}



// ==========================================
// === PHASE 5: OVERTIME & DAY OFF SYSTEM (Double Approval) ===
// ==========================================

/**
 * SUBMIT: Agent Request OR Manager Assignment
 */
function webSubmitOvertimeRequest(requestData) {
  const { userEmail: submitterEmail, userData, ss } = getAuthorizedContext(null);
  const otSheet = getOrCreateSheet(ss, SHEET_NAMES.overtime);
  
  // 1. Determine Target User (Agent vs Manager Assignment)
  let targetEmail = submitterEmail;
  let initiatedBy = "Agent";
  
  // If a manager is assigning to someone else
  if (requestData.targetEmail && requestData.targetEmail !== submitterEmail) {
      // Check permission
      const { userRole } = getAuthorizedContext('MANAGE_OVERTIME'); // Ensure they are a manager
      targetEmail = requestData.targetEmail;
      initiatedBy = `Manager (${userData.userName})`;
  }

  const targetUser = userData.userList.find(u => u.email === targetEmail);
  if (!targetUser) throw new Error("Target user not found.");

  // 2. Schedule Validation
  const shiftDate = new Date(requestData.date);
  const schedule = getScheduleForDate(targetEmail, shiftDate);
  const type = requestData.type;

  if (type === "Work Day Off") {
      if (schedule && schedule.start && schedule.leaveType !== 'Day Off' && schedule.leaveType !== 'Absent') {
          throw new Error(`User already has a shift on ${requestData.date}. Use Pre/Post Shift.`);
      }
  } else {
      if (!schedule || !schedule.end) {
          throw new Error(`No active schedule found for ${targetUser.name} on this date.`);
      }
  }

  // 3. Time Validation
  const otStartObj = createDateTime(shiftDate, requestData.startTime);
  let otEndObj = createDateTime(shiftDate, requestData.endTime);
  if (otEndObj < otStartObj) otEndObj.setDate(otEndObj.getDate() + 1); // Overnight
  
  const duration = (otEndObj - otStartObj) / (1000 * 60 * 60);
  if (duration <= 0) throw new Error("Invalid time range.");

  // 4. Determine Approval Flow
  const directMgr = targetUser.supervisor;
  const projectMgr = targetUser.projectManager;
  
  let directStatus = "Pending";
  let projectStatus = "Pending";
  let overallStatus = "Pending Direct Mgr";

  // Auto-approve if the submitter IS one of the managers
  if (submitterEmail === directMgr) {
      directStatus = "Approved";
      overallStatus = "Pending Project Mgr";
  }
  if (submitterEmail === projectMgr) {
      projectStatus = "Approved";
      // If Direct is still pending, it stays "Pending Direct". 
      // If Direct was already approved (unlikely in this flow), it moves to Approved.
  }
  
  // Edge Case: If submitter is Superadmin, approve ALL
  if (userData.userRole === 'superadmin') {
      directStatus = "Approved";
      projectStatus = "Approved";
      overallStatus = "Approved";
  }

  // 5. Save Request
  const reqID = `OT-${new Date().getTime()}`;
  otSheet.appendRow([
    reqID,
    targetUser.empID,
    targetUser.name,
    shiftDate,
    otStartObj,
    otEndObj,
    duration.toFixed(2),
    requestData.reason,
    overallStatus, // Status
    "",            // Comment
    "",            // ActionBy
    "",            // ActionDate
    type,
    directMgr,     // Col N
    projectMgr,    // Col O
    directStatus,  // Col P
    projectStatus, // Col Q
    initiatedBy    // Col R
  ]);

  // If Superadmin auto-approved, trigger schedule update immediately
  if (overallStatus === 'Approved') {
      finalizeOvertimeSchedule(ss, targetUser, requestData, submitterEmail);
  }

  return "Request submitted successfully.";
}

/**
 * ACTION: Manager Approves/Denies
 */
function webActionOvertime(reqId, action, comment) {
  const { userEmail: adminEmail, ss, userData } = getAuthorizedContext('MANAGE_OVERTIME');
  const otSheet = getOrCreateSheet(ss, SHEET_NAMES.overtime);
  
  const data = otSheet.getDataRange().getValues();
  let rowIdx = -1;
  let reqRow = [];

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === reqId) {
      rowIdx = i + 1;
      reqRow = data[i];
      break;
    }
  }
  if (rowIdx === -1) throw new Error("Request not found.");

  const directMgr = (reqRow[13] || "").toLowerCase();
  const projectMgr = (reqRow[14] || "").toLowerCase();
  const adminEmailLower = adminEmail.toLowerCase();
  
  // 1. Identify Role of Approver
  let isDirect = (adminEmailLower === directMgr);
  let isProject = (adminEmailLower === projectMgr);
  const isSuper = (userData.userRole === 'superadmin');

  // FIX: If Project Mgr is empty/missing, or same as Direct, treat Direct approval as both
  if (isDirect && (!projectMgr || projectMgr === directMgr)) {
      isProject = true;
  }

  if (!isDirect && !isProject && !isSuper) {
      throw new Error("You are not authorized to approve this request.");
  }

  // 2. Handle DENY
  if (action === 'Denied') {
      otSheet.getRange(rowIdx, 9).setValue("Denied");
      otSheet.getRange(rowIdx, 10).setValue(comment);
      otSheet.getRange(rowIdx, 11).setValue(adminEmail);
      otSheet.getRange(rowIdx, 12).setValue(new Date());
      return "Request Denied.";
  }

  // 3. Handle APPROVE
  // Update the specific column based on who is acting
  if (isDirect || isSuper) otSheet.getRange(rowIdx, 16).setValue("Approved"); // DirectStatus
  if (isProject || isSuper) otSheet.getRange(rowIdx, 17).setValue("Approved"); // ProjectStatus

  // Re-read statuses to decide if we can Finalize
  // We use the flags we just calculated + existing sheet data
  const currentDirectStatus = (isDirect || isSuper) ? "Approved" : reqRow[15];
  const currentProjectStatus = (isProject || isSuper) ? "Approved" : reqRow[16];

  let newMainStatus = reqRow[8]; // Default to current

  if (currentDirectStatus === 'Approved' && currentProjectStatus !== 'Approved') {
      newMainStatus = "Pending Project Mgr";
  } else if (currentDirectStatus === 'Approved' && currentProjectStatus === 'Approved') {
      newMainStatus = "Approved";
  }

  // Update Main Status & Metadata
  otSheet.getRange(rowIdx, 9).setValue(newMainStatus);
  otSheet.getRange(rowIdx, 11).setValue(adminEmail); 
  otSheet.getRange(rowIdx, 12).setValue(new Date());

  // 4. Finalize if Fully Approved
  if (newMainStatus === "Approved") {
      const requestData = {
          date: reqRow[3],
          type: reqRow[12],
          startTime: reqRow[4], // Date Obj
          endTime: reqRow[5],   // Date Obj
          name: reqRow[2]
      };
      
      const targetUserObj = userData.userList.find(u => u.empID === reqRow[1]);
      if (targetUserObj) {
          finalizeOvertimeSchedule(ss, targetUserObj, requestData, adminEmail);
          return "Request Finalized. Schedule Updated.";
      } else {
          return "Request Approved (Warning: User not found in DB to update schedule).";
      }
  }

  return "Approval Recorded. Waiting for next approver.";
}

// Helper: Updates Schedule with Overnight Logic
function finalizeOvertimeSchedule(ss, targetUser, reqData, adminEmail) {
    const scheduleSheet = getOrCreateSheet(ss, SHEET_NAMES.schedule);
    const logsSheet = getOrCreateSheet(ss, SHEET_NAMES.logs);
    const targetDateStr = Utilities.formatDate(new Date(reqData.date), Session.getScriptTimeZone(), "MM/dd/yyyy");

    const formatT = (d) => Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), "HH:mm");

    // 1. Calculate Shift Dates (Handle Overnight)
    // If End Time is earlier than Start Time, it implies the shift ends the next day.
    const startObj = new Date(reqData.startTime);
    const endObj = new Date(reqData.endTime);
    
    let shiftEndDate = new Date(reqData.date); // Default to same day
    if (endObj < startObj) {
        shiftEndDate.setDate(shiftEndDate.getDate() + 1); // Next day
    }

    if (reqData.type === 'Work Day Off') {
        updateOrAddSingleSchedule(
            scheduleSheet, {}, logsSheet,
            targetUser.email, targetUser.name,
            new Date(reqData.date), shiftEndDate, // Pass correct End Date
            targetDateStr,
            formatT(startObj), formatT(endObj),
            "Work Day Off", adminEmail
        );
    } else {
        // Pre/Post Shift: Extend Schedule
        const curSched = getScheduleForDate(targetUser.email, new Date(reqData.date));
        if (curSched) {
            let s = curSched.start;
            let e = curSched.end;
            
            // If extending, we respect the NEW boundary
            if (reqData.type === 'Pre-Shift') s = formatT(startObj);
            if (reqData.type === 'Post-Shift') {
                e = formatT(endObj);
                // Recalculate shift end date based on new extended time
                const schedStartObj = createDateTime(new Date(reqData.date), s);
                const schedEndObj = createDateTime(new Date(reqData.date), e);
                shiftEndDate = new Date(reqData.date);
                if (schedEndObj < schedStartObj) shiftEndDate.setDate(shiftEndDate.getDate() + 1);
            }
            
            updateOrAddSingleSchedule(
                scheduleSheet, {}, logsSheet,
                targetUser.email, targetUser.name,
                new Date(reqData.date), shiftEndDate,
                targetDateStr,
                s, e, "Present", adminEmail
            );
        }
    }
}

// Helper: Fetch List with Details
function webGetOvertimeRequests(filterStatus) {
  const { userEmail, userData, ss } = getAuthorizedContext(null);
  const otSheet = getOrCreateSheet(ss, SHEET_NAMES.overtime);
  const data = otSheet.getDataRange().getValues();
  const results = [];
  
  const isManager = ['admin','superadmin','manager','project_manager'].includes(userData.userRole);
  const mySubs = isManager ? new Set(webGetAllSubordinateEmails(userEmail)) : new Set();

  const formatT = (val) => {
    if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), "HH:mm");
    if (typeof val === 'string' && val.includes('T')) return val.split('T')[1].substring(0,5);
    // Try parsing if string date
    const d = parseDate(val);
    if (d) return Utilities.formatDate(d, Session.getScriptTimeZone(), "HH:mm");
    return val || "";
  };

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row || row.length === 0) continue;

    const ownerID = row[1];
    const ownerUser = userData.userList.find(u => u.empID === ownerID);
    const ownerEmail = ownerUser ? ownerUser.email : "";
    
    let canView = (ownerEmail === userEmail);
    if (!canView && isManager) {
        if (userData.userRole === 'superadmin') canView = true;
        else if (mySubs.has(ownerEmail)) canView = true;
    }

    if (canView) {
        const status = row[8] || "Pending";
        
        // Parse the ShiftDate (row[3]) safely
        const shiftDateObj = parseDate(row[3]);
        const shiftDateStr = shiftDateObj ? convertDateToString(shiftDateObj).split('T')[0] : "Invalid Date";

        if (filterStatus === 'All' || status === filterStatus || (filterStatus === 'Pending' && status.includes('Pending'))) {
            results.push({
                id: row[0],
                name: row[2],
                date: shiftDateStr,
                time: `${formatT(row[4])} - ${formatT(row[5])}`,
                hours: row[6],
                reason: row[7],
                status: status,
                type: row[12] || "N/A",
                directMgr: row[13] || "",
                projectMgr: row[14] || "",
                directStatus: row[15] || "Pending",
                projectStatus: row[16] || "Pending",
                initiatedBy: row[17] || "Agent"
            });
        }
    }
  }
  return results.reverse();
}

/**
 * MANAGER: Approve/Deny or Pre-Approve
 */
function webActionOvertime(reqId, action, comment, preApproveData) {
  const { userEmail, ss } = getAuthorizedContext('MANAGE_OVERTIME');
  const otSheet = getOrCreateSheet(ss, SHEET_NAMES.overtime);
  
  // CASE 1: Pre-Approval (Creating a new Approved request)
  if (action === 'Pre-Approve') {
    const targetEmail = preApproveData.email;
    const userData = getUserDataFromDb(ss);
    const targetUser = userData.userList.find(u => u.email === targetEmail);
    if (!targetUser) throw new Error("User not found.");
    
    const schedule = getScheduleForDate(targetEmail, new Date(preApproveData.date));
    if (!schedule) throw new Error("No schedule found for user on this date.");

    const newID = `OT-PRE-${new Date().getTime()}`;
    otSheet.appendRow([
      newID,
      targetUser.empID,
      targetUser.name,
      new Date(preApproveData.date),
      new Date(schedule.start),
      new Date(schedule.end),
      preApproveData.hours,
      "Pre-Approved by Manager",
      "Approved",
      comment || "Pre-approved",
      userEmail,
      new Date()
    ]);
    return "Overtime pre-approved successfully.";
  }

  // CASE 2: Action Existing Request
  const data = otSheet.getDataRange().getValues();
  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === reqId) {
      rowIndex = i + 1;
      break;
    }
  }
  if (rowIndex === -1) throw new Error("Request not found.");
  
  // Update Status (Col I = 9), Comment (Col J = 10), ActionBy (Col K = 11), ActionDate
  otSheet.getRange(rowIndex, 9).setValue(action); // Approved/Denied
  otSheet.getRange(rowIndex, 10).setValue(comment);
  otSheet.getRange(rowIndex, 11).setValue(userEmail);
  otSheet.getRange(rowIndex, 12).setValue(new Date());
  
  return `Request ${action}.`;
}

/**
 * NEW PHASE 5: Calculates net working hours (Total Login Duration - Excess Break Time).
 * Returns decimal hours (e.g., 8.5).
 */
function calculateNetHours(punches) {
  if (!punches.login || !punches.logout) return 0;

  const totalDurationSec = timeDiffInSeconds(punches.login, punches.logout);
  
  // Helper to calculate excess
  const getExcess = (start, end, type) => {
    if (!start || !end) return 0;
    const duration = timeDiffInSeconds(start, end);
    const allowed = getBreakConfig(type).default;
    return Math.max(0, duration - allowed);
  };

  const deduct1 = getExcess(punches.firstBreakIn, punches.firstBreakOut, "First Break");
  const deductLunch = getExcess(punches.lunchIn, punches.lunchOut, "Lunch");
  const deduct2 = getExcess(punches.lastBreakIn, punches.lastBreakOut, "Last Break");

  const netSeconds = totalDurationSec - deduct1 - deductLunch - deduct2;
  return (netSeconds / 3600).toFixed(2); // Return decimal hours
}

// ================= PHASE 6: CONFIGURATION API =================

/**
 * Fetches the current break configuration for the Admin Editor.
 */
function webGetBreakConfig() {
  const { userEmail, userData, ss } = getAuthorizedContext('MANAGE_BALANCES'); // Reusing Admin permission
  const sheet = getOrCreateSheet(ss, SHEET_NAMES.breakConfig);
  const data = sheet.getDataRange().getValues();
  
  // Skip header row
  const configs = [];
  for (let i = 1; i < data.length; i++) {
    configs.push({
      type: data[i][0],
      defaultDur: data[i][1], // Seconds
      maxDur: data[i][2]      // Seconds
    });
  }
  return configs;
}

/**
 * Saves changes to the Break Configuration sheet.
 */
function webSaveBreakConfig(newConfigs) {
  const { userEmail, userData, ss } = getAuthorizedContext('MANAGE_BALANCES');
  const sheet = getOrCreateSheet(ss, SHEET_NAMES.breakConfig);
  const data = sheet.getDataRange().getValues();
  
  // newConfigs is an array of { type, defaultDur, maxDur }
  newConfigs.forEach(conf => {
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === conf.type) {
        // Update Default (Col B) and Max (Col C)
        sheet.getRange(i + 1, 2).setValue(Number(conf.defaultDur));
        sheet.getRange(i + 1, 3).setValue(Number(conf.maxDur));
        break;
      }
    }
  });
  
  return "Break configuration updated successfully.";
}


// --- 3. PHASE 5: ENTITLEMENT TEMPLATES (Financial Admin) ---

function webSaveEntitlementTemplate(templateData) {
  const { userEmail, ss } = getAuthorizedContext('MANAGE_FINANCE');
  const sheet = getOrCreateSheet(ss, "Entitlement_Templates");
  
  const id = templateData.id || `TMP-${new Date().getTime()}`;
  
  if (templateData.id) {
      // Update logic (simplified: delete old, add new, or find row)
      // For simplicity in this iteration, we just append or use a find loop.
      const data = sheet.getDataRange().getValues();
      for(let i=1; i<data.length; i++) {
          if(data[i][0] === id) {
              sheet.deleteRow(i+1);
              break;
          }
      }
  }
  
  sheet.appendRow([
      id,
      templateData.name,
      templateData.type,
      templateData.amount,
      templateData.currency,
      templateData.description,
      "Active"
  ]);
  
  return "Template saved successfully.";
}

function webGetEntitlementTemplates() {
  const { ss } = getAuthorizedContext('MANAGE_FINANCE');
  const sheet = getOrCreateSheet(ss, "Entitlement_Templates");
  const data = sheet.getDataRange().getValues();
  const templates = [];
  
  for (let i = 1; i < data.length; i++) {
      if (data[i][6] === 'Active') {
          templates.push({
              id: data[i][0],
              name: data[i][1],
              type: data[i][2],
              amount: data[i][3],
              currency: data[i][4],
              desc: data[i][5]
          });
      }
  }
  return templates;
}

function webApplyEntitlementTemplate(templateId, targetEmails) {
  const { userEmail: adminEmail, ss, userData } = getAuthorizedContext('MANAGE_FINANCE');
  
  // 1. Get Template
  const tmplSheet = getOrCreateSheet(ss, "Entitlement_Templates");
  const tmplData = tmplSheet.getDataRange().getValues();
  let template = null;
  for(let i=1; i<tmplData.length; i++) {
      if(tmplData[i][0] === templateId) {
          template = { type: tmplData[i][2], amount: tmplData[i][3], currency: tmplData[i][4], desc: tmplData[i][5] };
          break;
      }
  }
  if (!template) throw new Error("Template not found.");

  // 2. Apply to Users
  const finSheet = getOrCreateSheet(ss, SHEET_NAMES.financialEntitlements);
  let count = 0;
  
  targetEmails.forEach(email => {
      const user = userData.userList.find(u => u.email === email);
      if (user) {
          const id = `FIN-${new Date().getTime()}-${Math.floor(Math.random()*1000)}`;
          finSheet.appendRow([
              id,
              user.email,
              user.name,
              template.type,
              template.amount,
              template.currency,
              new Date(), // Due Date (Today)
              "Pending",
              template.desc,
              adminEmail,
              new Date()
          ]);
          count++;
      }
  });
  
  return `Successfully applied template to ${count} users.`;
}

// --- 4. PHASE 6: PROJECT ADMIN DASHBOARD ---

function webGetProjectDashboard(projectId) {
  const { userEmail, userData, ss } = getAuthorizedContext('MANAGE_PROJECTS'); // Project Manager Permission
  
  // Validate Access (User must manage this project or be Superadmin)
  // Logic: Check if userEmail matches the project's manager in Projects sheet
  const projSheet = getOrCreateSheet(ss, SHEET_NAMES.projects);
  const projData = projSheet.getDataRange().getValues();
  let projectInfo = null;
  
  for(let i=1; i<projData.length; i++) {
      if(projData[i][0] === projectId) {
          projectInfo = { id: projData[i][0], name: projData[i][1], manager: projData[i][2] };
          break;
      }
  }
  
  if (!projectInfo) throw new Error("Project not found.");
  if (userData.userRole !== 'superadmin' && projectInfo.manager !== userEmail) {
      throw new Error("Permission denied. You do not manage this project.");
  }

  // 1. Calculate Stats
  const stats = {
      agents: 0,
      tls: 0,
      supers: 0,
      total: 0
  };
  
  userData.userList.forEach(u => {
      if (u.projectManager === projectInfo.manager) { // Simple link by manager, or use ProjectID from Core if available
          stats.total++;
          if (u.role === 'agent') stats.agents++;
          else if (u.role === 'manager') stats.supers++; // Assuming manager = supervisor
          else stats.tls++; // Placeholder logic for TLs if role exists
      }
  });

  return { info: projectInfo, stats: stats };
}

function webGetProjectRequests(projectId) {
    const { userEmail, userData, ss } = getAuthorizedContext('MANAGE_PROJECTS');
    // Fetch consolidated requests for all users under this project manager
    // For simplicity, we filter by the manager's email (since ProjectID linking might be loose)
    
    // 1. Get Project Manager Email
    const projSheet = getOrCreateSheet(ss, SHEET_NAMES.projects);
    const projData = projSheet.getDataRange().getValues();
    let pmEmail = "";
    for(let i=1; i<projData.length; i++) {
        if(projData[i][0] === projectId) { pmEmail = projData[i][2]; break; }
    }
    
    const teamEmails = new Set();
    userData.userList.forEach(u => {
        if (u.projectManager === pmEmail) teamEmails.add(u.email);
    });

    const requests = [];
    
    // 2. Fetch Leave
    const leaveData = getOrCreateSheet(ss, SHEET_NAMES.leaveRequests).getDataRange().getValues();
    for(let i=1; i<leaveData.length; i++) {
        if(teamEmails.has(leaveData[i][2])) {
            requests.push({ type: "Leave", user: leaveData[i][3], date: convertDateToString(new Date(leaveData[i][9] || new Date())), status: leaveData[i][1] });
        }
    }
    
    // 3. Fetch Overtime
    const otData = getOrCreateSheet(ss, SHEET_NAMES.overtime).getDataRange().getValues();
    for(let i=1; i<otData.length; i++) {
        // Need to lookup user email from ID (col 1)
        const u = userData.userList.find(x => x.empID === otData[i][1]);
        if(u && teamEmails.has(u.email)) {
            requests.push({ type: "Overtime", user: otData[i][2], date: convertDateToString(new Date(otData[i][3])), status: otData[i][8] });
        }
    }

    return requests.slice(0, 50); // Return last 50
}

// ==========================================
// === PHASE 7: OFFBOARDING WORKFLOW ===
// ==========================================

/**
 * 1. Submit Resignation (Agent)
 */
function webSubmitResignation(reason, exitDate) {
  const { userEmail, userData, ss } = getAuthorizedContext(null);
  const sheet = getOrCreateSheet(ss, SHEET_NAMES.offboarding);
  
  // Check for existing pending request
  const data = sheet.getDataRange().getValues();
  for(let i=1; i<data.length; i++) {
      if(data[i][3] === userEmail && data[i][6].includes("Pending")) {
          throw new Error("You already have a pending resignation request.");
      }
  }

  const user = userData.userList.find(u => u.email === userEmail);
  const reqID = `EXIT-${new Date().getTime()}`;
  
  sheet.appendRow([
      reqID,
      user.empID,
      user.name,
      userEmail,
      "Resignation",
      reason,
      "Pending Managers",
      user.supervisor,
      user.projectManager,
      "Pending", // DirectStatus
      "Pending", // ProjectStatus
      "Pending", // HRStatus
      new Date(),
      new Date(exitDate),
      "Agent"
  ]);
  
  return "Resignation submitted. It will be reviewed by your managers and HR.";
}

/**
 * 2. Submit Termination (Manager/HR)
 */
function webSubmitTermination(targetEmail, reason, exitDate) {
  const { userEmail: adminEmail, userData, ss } = getAuthorizedContext('OFFBOARD_EMPLOYEE'); 
  // 'OFFBOARD_EMPLOYEE' permission is for Managers & HR
  
  const sheet = getOrCreateSheet(ss, SHEET_NAMES.offboarding);
  const targetUser = userData.userList.find(u => u.email === targetEmail);
  
  if (!targetUser) throw new Error("User not found.");
  
  // Logic: Who is initiating?
  const isHR = userData.userRole === 'superadmin' || userData.userRole === 'hr_manager'; // Assuming HR role exists or superadmin
  
  const reqID = `TERM-${new Date().getTime()}`;
  const status = isHR ? "Approved" : "Pending HR"; // HR terminates immediately, Managers need HR approval
  const hrStatus = isHR ? "Approved" : "Pending";
  
  sheet.appendRow([
      reqID,
      targetUser.empID,
      targetUser.name,
      targetEmail,
      "Termination",
      reason,
      status,
      targetUser.supervisor,
      targetUser.projectManager,
      "N/A", // DirectStatus (Bypassed)
      "N/A", // ProjectStatus (Bypassed)
      hrStatus,
      new Date(),
      new Date(exitDate),
      `Manager: ${userData.userName}`
  ]);

  if (isHR) {
      // Execute Immediate Offboarding
      executeOffboarding(ss, targetEmail, exitDate, reason);
      return `Employee ${targetUser.name} has been terminated immediately.`;
  }
  
  return `Termination request for ${targetUser.name} submitted to HR.`;
}

/**
 * 3. Get Requests (For Manager/HR Dashboard)
 */
function webGetOffboardingRequests() {
  const { userEmail, userData, ss } = getAuthorizedContext(null);
  const sheet = getOrCreateSheet(ss, SHEET_NAMES.offboarding);
  const data = sheet.getDataRange().getValues();
  const results = [];
  
  const isHR = userData.userRole === 'superadmin'; 
  const isManager = ['admin','manager','project_manager'].includes(userData.userRole);
  
  for(let i=1; i<data.length; i++) {
    try {
      const row = data[i];
      if (!row || row.length === 0) continue; // Skip empty rows

      const targetEmail = row[3];
      const directMgr = row[7];
      const projectMgr = row[8];
      
      let canView = false;
      if (isHR) canView = true;
      if (isManager && (userEmail === directMgr || userEmail === projectMgr)) canView = true;
      if (userEmail === targetEmail) canView = true; 
      
      if (canView) {
          // --- FIX: Safe Date Parsing ---
          let exitDateStr = "N/A";
          if (row[13]) {
             try {
               const d = new Date(row[13]);
               if (!isNaN(d.getTime())) {
                  exitDateStr = Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
               }
             } catch(e) {}
          }
          // -----------------------------

          results.push({
              id: row[0],
              name: row[2],
              type: row[4],
              reason: row[5],
              status: row[6],
              directStatus: row[9],
              projectStatus: row[10],
              hrStatus: row[11],
              exitDate: exitDateStr, // Uses the safe string
              initiatedBy: row[14]
          });
      }
    } catch (e) {
      Logger.log("Error reading offboarding row " + i + ": " + e.message);
    }
  }
  return results.reverse();
}

/**
 * 4. Action Request (Approve/Deny)
 */
function webActionOffboarding(reqId, action, comment) {
  const { userEmail, userData, ss } = getAuthorizedContext('OFFBOARD_EMPLOYEE');
  const sheet = getOrCreateSheet(ss, SHEET_NAMES.offboarding);
  const data = sheet.getDataRange().getValues();
  
  let rowIdx = -1;
  let reqRow = [];
  for(let i=1; i<data.length; i++) {
      if(data[i][0] === reqId) { rowIdx = i+1; reqRow = data[i]; break; }
  }
  
  if (rowIdx === -1) throw new Error("Request not found.");
  
  const directMgr = reqRow[7];
  const projectMgr = reqRow[8];
  const isHR = userData.userRole === 'superadmin';
  const isDirect = userEmail === directMgr;
  const isProject = userEmail === projectMgr;
  
  // DENY Logic
  if (action === 'Denied') {
      sheet.getRange(rowIdx, 7).setValue("Denied");
      // Log who denied? ideally yes
      return "Request Denied.";
  }
  
  // APPROVE Logic
  if (reqRow[4] === 'Resignation') {
      if (isDirect) sheet.getRange(rowIdx, 10).setValue("Approved");
      if (isProject) sheet.getRange(rowIdx, 11).setValue("Approved");
      if (isHR) sheet.getRange(rowIdx, 12).setValue("Approved");
      
      // Check status to advance
      // We need to re-read or assume based on current action
      const dStat = isDirect ? "Approved" : reqRow[9];
      const pStat = isProject ? "Approved" : reqRow[10];
      const hStat = isHR ? "Approved" : reqRow[11];
      
      if (dStat === 'Approved' && pStat === 'Approved' && hStat === 'Pending') {
          sheet.getRange(rowIdx, 7).setValue("Pending HR");
      } else if (dStat === 'Approved' && pStat === 'Approved' && hStat === 'Approved') {
          sheet.getRange(rowIdx, 7).setValue("Approved");
          // Finalize
          executeOffboarding(ss, reqRow[3], reqRow[13], "Resignation Approved");
          return "Resignation Finalized. Employee Offboarded.";
      }
  } else if (reqRow[4] === 'Termination') {
      if (isHR) {
          sheet.getRange(rowIdx, 12).setValue("Approved");
          sheet.getRange(rowIdx, 7).setValue("Approved");
          executeOffboarding(ss, reqRow[3], reqRow[13], "Termination Approved");
          return "Termination Finalized.";
      }
  }
  
  return "Approval Recorded.";
}

// Helper: Execute Final Offboarding (Update Core, Log History)
function executeOffboarding(ss, email, exitDate, reason) {
    const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.employeesCore);
    const data = dbSheet.getDataRange().getValues();
    const userData = getUserDataFromDb(ss);
    const row = userData.emailToRow[email];
    
    if (row) {
        // Update Status (Col 27/AA based on new schema, check index)
        // Schema: ..., ExitDate, Status
        // Let's look up headers dynamically to be safe
        const headers = data[0];
        const statusIdx = headers.indexOf("Status");
        const exitIdx = headers.indexOf("ExitDate");
        
        if (statusIdx > -1) dbSheet.getRange(row, statusIdx+1).setValue("Left");
        if (exitIdx > -1) dbSheet.getRange(row, exitIdx+1).setValue(new Date(exitDate));
        
        // Log History
        const histSheet = getOrCreateSheet(ss, SHEET_NAMES.historyLogs);
        histSheet.appendRow([
            `HIST-${new Date().getTime()}`,
            userData.userList.find(u=>u.email===email)?.empID,
            new Date(),
            "Offboarding",
            "Active",
            "Left"
        ]);
    }
}


// ==========================================
// === PHASE 8: ANALYTICS & REPORTS ===
// ==========================================

/**
 * Fetches data for the Visual Dashboard (Metrics + Timeline)
 */
function webGetAnalyticsData(filter) {
  const { userEmail, userData, ss } = getAuthorizedContext('VIEW_FULL_DASHBOARD');
  
  // --- FIX: Adjust Dates to cover full day ---
  const startDate = new Date(filter.startDate);
  startDate.setHours(0, 0, 0, 0); // Start of Day
  
  const endDate = new Date(filter.endDate);
  endDate.setHours(23, 59, 59, 999); // End of Day
  // ------------------------------------------

  const targetEmails = filter.targetEmails || []; 
  const timeZone = Session.getScriptTimeZone();

  const targetSet = new Set(targetEmails);
  const processAll = targetSet.has('ALL');

  const adherenceData = getOrCreateSheet(ss, SHEET_NAMES.adherence).getDataRange().getValues();
  const otherCodesData = getOrCreateSheet(ss, SHEET_NAMES.otherCodes).getDataRange().getValues();
  const scheduleSheet = getOrCreateSheet(ss, SHEET_NAMES.schedule);
  const scheduleData = scheduleSheet.getDataRange().getValues();
  
  let totalScheduledMins = 0;
  let totalWorkedMins = 0;
  let totalShrinkageMins = 0;
  let totalLateness = 0;
  
  const timeline = {};

  // Build Schedule Map
  const scheduleMap = {};
  for (let i = 1; i < scheduleData.length; i++) {
    const row = scheduleData[i];
    const sEmail = (row[6] || "").toLowerCase();
    
    if (!processAll && !targetSet.has(sEmail)) continue;

    const sDate = parseDate(row[1]);
    if (!sEmail || !sDate) continue;
    
    const dateKey = Utilities.formatDate(sDate, timeZone, "yyyy-MM-dd");
    const uniqueKey = `${sEmail}_${dateKey}`;
    
    const fmt = (v) => (v instanceof Date) ? Utilities.formatDate(v, timeZone, "HH:mm") : (v ? v.toString().substring(0, 5) : null);

    scheduleMap[uniqueKey] = {
      type: row[5],
      start: fmt(row[2]), end: fmt(row[4]),
      b1_start: fmt(row[7]), b1_end: fmt(row[8]),
      l_start: fmt(row[9]),  l_end: fmt(row[10]),
      b2_start: fmt(row[11]), b2_end: fmt(row[12])
    };
  }

  const toDec = (t) => {
    if (!t) return null;
    const [h, m] = t.split(':').map(Number);
    return h + (m / 60);
  };

  // Process Adherence
  for (let i = 1; i < adherenceData.length; i++) {
    const row = adherenceData[i];
    const rowDate = new Date(row[0]);
    const agentName = row[1];
    const email = userData.nameToEmail[agentName];
    
    // Check Date Range (Now inclusive of time)
    if (rowDate < startDate || rowDate > endDate) continue;
    
    if (!email || (!processAll && !targetSet.has(email))) continue;

    const dateStr = Utilities.formatDate(rowDate, timeZone, "yyyy-MM-dd");
    const schedKey = `${email}_${dateStr}`;
    const sched = scheduleMap[schedKey] || { type: 'Day Off' };

    if (!timeline[dateStr]) timeline[dateStr] = {};
    if (!timeline[dateStr][agentName]) {
      timeline[dateStr][agentName] = { events: [], flags: [], schedule: sched };
    }

    const netHours = parseFloat(row[22]) || 0;
    totalWorkedMins += (netHours * 60);
    const tardySec = parseFloat(row[10]) || 0;
    totalLateness += (tardySec / 60);
    const leaveType = (row[13] || "").toLowerCase();
    
    if (leaveType !== 'day off') totalScheduledMins += 540;
    if (leaveType === 'absent' || (leaveType !== 'present' && leaveType !== 'day off' && leaveType !== '')) {
       totalShrinkageMins += 540;
    }
    totalShrinkageMins += (tardySec / 60);

    const formatT = (v) => (v instanceof Date) ? Utilities.formatDate(v, timeZone, "HH:mm") : null;
    const login = formatT(row[2]);
    const logout = formatT(row[9]);
    
    if (login && logout) timeline[dateStr][agentName].events.push({ type: 'Work', start: login, end: logout, label: 'Shift' });

    const addBreak = (inT, outT, type, label) => {
        const s = formatT(inT);
        const e = formatT(outT);
        if (s && e) {
            timeline[dateStr][agentName].events.push({ type: type, start: s, end: e, label: label });
            return { start: s, end: e };
        }
        return null;
    };

    const actB1 = addBreak(row[3], row[4], 'Break', '1st Break');
    const actLunch = addBreak(row[5], row[6], 'Lunch', 'Lunch');
    const actB2 = addBreak(row[7], row[8], 'Break', 'Last Break');

    if (sched.start && login && toDec(login) > toDec(sched.start) + (5/60)) {
        timeline[dateStr][agentName].flags.push({ type: 'Lateness', time: sched.start, msg: `Late: Login at ${login}` });
    }
    if (sched.end && logout && toDec(logout) < toDec(sched.end) - (5/60)) {
        timeline[dateStr][agentName].flags.push({ type: 'EarlyLeave', time: logout, msg: `Early Leave: Out at ${logout}` });
    }
    if (sched.type === 'Present' && !login && leaveType !== 'Sick' && leaveType !== 'Annual') {
        timeline[dateStr][agentName].flags.push({ type: 'NoShow', time: '09:00', msg: 'No Show / Absent' });
    }

    const checkWindow = (actual, sStart, sEnd, name) => {
        if (actual && sStart && sEnd) {
            const actStart = toDec(actual.start);
            const winStart = toDec(sStart);
            const winEnd = toDec(sEnd);
            if (actStart < winStart || actStart > winEnd) {
                timeline[dateStr][agentName].flags.push({ type: 'Adherence', time: actual.start, msg: `${name} out of window` });
            }
        }
    };
    checkWindow(actB1, sched.b1_start, sched.b1_end, "1st Break");
    checkWindow(actLunch, sched.l_start, sched.l_end, "Lunch");
    checkWindow(actB2, sched.b2_start, sched.b2_end, "Last Break");
  }

  // Process AUX
  for (let i = 1; i < otherCodesData.length; i++) {
      const row = otherCodesData[i];
      const rowDate = new Date(row[0]);
      const agentName = row[1];
      const email = userData.nameToEmail[agentName];
      
      // Check Date Range (Fix)
      if (rowDate < startDate || rowDate > endDate) continue;
      
      if (!email || (!processAll && !targetSet.has(email))) continue;
      
      const dateStr = Utilities.formatDate(rowDate, timeZone, "yyyy-MM-dd");
      
      if (!timeline[dateStr]) timeline[dateStr] = {};
      if (!timeline[dateStr][agentName]) timeline[dateStr][agentName] = { events: [], flags: [], schedule: {} };

      const formatT = (v) => (v instanceof Date) ? Utilities.formatDate(v, timeZone, "HH:mm") : null;
      const s = formatT(row[3]);
      const e = formatT(row[4]) || formatT(new Date()); 

      if (s) timeline[dateStr][agentName].events.push({ type: 'Aux', start: s, end: e, label: row[2] });
  }

  const adherenceScore = totalScheduledMins > 0 ? ((totalWorkedMins / totalScheduledMins) * 100).toFixed(1) : 100;
  const shrinkageScore = totalScheduledMins > 0 ? ((totalShrinkageMins / totalScheduledMins) * 100).toFixed(1) : 0;

  return {
    metrics: { adherence: adherenceScore, shrinkage: shrinkageScore, latenessMins: Math.round(totalLateness), workedHours: (totalWorkedMins/60).toFixed(1) },
    timeline: timeline
  };
}

/**
 * Helper: Format date obj to HH:mm:ss string for timeline
 */
function formatTime(val) {
  if (!val) return null;
  if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), "HH:mm:ss");
  // Handle strings
  if (typeof val === 'string' && val.includes('T')) return val.split('T')[1].substring(0,8);
  return val.toString().substring(0,8);
}

/**
 * Generates Raw Data for Payroll CSV Export
 */
function webGetPayrollExportData(startDateStr, endDateStr, targetEmails) {
  const { userEmail, userData, ss } = getAuthorizedContext('VIEW_FULL_DASHBOARD');
  
  const startDate = new Date(startDateStr);
  startDate.setHours(0, 0, 0, 0);
  
  const endDate = new Date(endDateStr);
  endDate.setHours(23, 59, 59, 999);

  const otData = getOrCreateSheet(ss, SHEET_NAMES.overtime).getDataRange().getValues();
  const adherenceData = getOrCreateSheet(ss, SHEET_NAMES.adherence).getDataRange().getValues();

  // Normalize target list
  const emailList = Array.isArray(targetEmails) ? targetEmails : [targetEmails];
  const targetSet = new Set(emailList.map(e => e.toLowerCase().trim()));
  const processAll = targetSet.has('all');

  const exportRows = [];

  for (let i = 1; i < adherenceData.length; i++) {
    const row = adherenceData[i];
    
    // 1. Robust Date Parsing (Handles DD/MM/YYYY strings and Date objects)
    const rowDate = parseDate(row[0]); 
    if (!rowDate) continue; // Skip invalid dates

    // 2. Date Filter
    if (rowDate < startDate || rowDate > endDate) continue;

    // 3. User Filter
    const agentName = (row[1] || "").toString().trim();
    // Try map lookup, fallback to raw check if name mismatch
    let email = userData.nameToEmail[agentName]; 
    if (!email) {
       // Fallback: Try finding user by name in the userList directly if nameToEmail failed
       const u = userData.userList.find(u => u.name.toLowerCase() === agentName.toLowerCase());
       if (u) email = u.email;
    }

    if (!email) continue;
    if (!processAll && !targetSet.has(email.toLowerCase())) continue;

    // 4. Overtime Lookup
    let approvedOTHours = 0;
    let otType = "";
    const dateStr = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
    
    for(let j=1; j<otData.length; j++) {
        const otEmpID = otData[j][1];
        // Match by Employee ID (more reliable than email)
        const otUser = userData.userList.find(u => u.empID === otEmpID);
        if(otUser && otUser.email.toLowerCase() === email.toLowerCase()) {
            const otDate = parseDate(otData[j][3]); // Robust parse for OT too
            if (otDate) {
               const otDateStr = Utilities.formatDate(otDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
               if (otDateStr === dateStr && otData[j][8] === 'Approved') {
                   approvedOTHours += parseFloat(otData[j][6] || 0);
                   otType = otData[j][12];
               }
            }
        }
    }

    const formatTime = (val) => {
        if (!val) return "";
        if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), "HH:mm:ss");
        if (typeof val === 'string' && val.includes('T')) return val.split('T')[1].substring(0,8);
        return val.toString().substring(0,8);
    };

    exportRows.push({
        Date: dateStr,
        Name: agentName,
        Email: email,
        Status: row[13],
        Login: formatTime(row[2]),
        Logout: formatTime(row[9]),
        NetWorkedHours: (parseFloat(row[22]) || 0).toFixed(2),
        LatenessMins: Math.round((row[10]||0)/60),
        EarlyLeaveMins: Math.round((row[12]||0)/60),
        BreakExceedMins: Math.round(((row[16]||0) + (row[18]||0))/60),
        LunchExceedMins: Math.round((row[17]||0)/60),
        ApprovedOTHours: approvedOTHours.toFixed(2),
        OTType: otType,
        IsAbsent: row[19]
    });
  }
  return exportRows;
}

//.............................................................................................................................







function _MASTER_DB_FIXER() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  Logger.log("Starting Master DB Fixer...");

  const schema = {
    // ... (Keep all existing schemas same as before) ...
    [SHEET_NAMES.rbac]: ["PermissionID", "Description", "superadmin", "admin", "manager", "project_manager", "financial_manager", "agent"],
    [SHEET_NAMES.employeesCore]: ["EmployeeID", "Name", "Email", "Role", "AccountStatus", "DirectManagerEmail", "FunctionalManagerEmail", "AnnualBalance", "SickBalance", "CasualBalance", "Gender", "EmploymentType", "ContractType", "JobLevel", "Department", "Function", "SubFunction", "GCMLevel", "Scope", "OffshoreOnshore", "DottedManager", "ProjectManagerEmail", "BonusPlan", "N_Level", "ExitDate", "Status"],
    [SHEET_NAMES.employeesPII]: ["EmployeeID", "HiringDate", "Salary", "IBAN", "Address", "Phone", "MedicalInfo", "ContractType", "NationalID", "PassportNumber", "SocialInsuranceNumber", "BirthDate", "PersonalEmail", "MaritalStatus", "DependentsInfo", "EmergencyContact", "EmergencyRelation", "BasicSalary", "VariablePay", "HourlyRate"],
    [SHEET_NAMES.financialEntitlements]: ["EntitlementID", "EmployeeEmail", "EmployeeName", "Type", "Amount", "Currency", "DueDate", "Status", "Description", "AddedBy", "DateAdded"],
    ["Entitlement_Templates"]: ["TemplateID", "Name", "Type", "DefaultAmount", "Currency", "Description", "Status"],
    [SHEET_NAMES.pendingRegistrations]: ["RequestID", "UserEmail", "UserName", "DirectManagerEmail", "FunctionalManagerEmail", "DirectStatus", "FunctionalStatus", "Address", "Phone", "RequestTimestamp", "HiringDate", "WorkflowStage"],
    [SHEET_NAMES.recruitment]: ["CandidateID", "Name", "Email", "Phone", "Position", "CV_Link", "Status", "Stage", "InterviewScores", "AppliedDate", "NationalID", "LangLevel", "SecondLang", "Referrer", "HR_Feedback", "Mgmt_Feedback", "Tech_Feedback", "Client_Feedback", "OfferStatus", "RejectionReason", "HistoryLog"],
    [SHEET_NAMES.requisitions]: ["ReqID", "Title", "Department", "HiringManager", "OpenDate", "Status", "PoolCandidates", "JobDescription"],
    [SHEET_NAMES.performance]: ["ReviewID", "EmployeeID", "Year", "ReviewPeriod", "Rating", "ManagerComments", "Date"],
    [SHEET_NAMES.historyLogs]: ["HistoryID", "EmployeeID", "Date", "EventType", "OldValue", "NewValue"],
    [SHEET_NAMES.adherence]: ["Date", "User Name", "Login", "First Break In", "First Break Out", "Lunch In", "Lunch Out", "Last Break In", "Last Break Out", "Logout", "Tardy (Seconds)", "Overtime (Seconds)", "Early Leave (Seconds)", "Leave Type", "Admin Audit", "—", "1st Break Exceed", "Lunch Exceed", "Last Break Exceed", "Absent", "Admin Code", "BreakWindowViolation", "NetLoginHours", "PreShiftOvertime", "LastAction", "LastActionTimestamp"],
    [SHEET_NAMES.schedule]: ["Name", "StartDate", "ShiftStartTime", "EndDate", "ShiftEndTime", "LeaveType", "agent email"],
    [SHEET_NAMES.logs]: ["Timestamp", "User Name", "Email", "Action", "Time"],
    [SHEET_NAMES.otherCodes]: ["Date", "User Name", "Code", "Time In", "Time Out", "Duration (Seconds)", "Admin Audit (Email)"],
    [SHEET_NAMES.warnings]: ["WarningID", "EmployeeID", "Type", "Level", "Date", "Description", "Status", "IssuedBy"],
    [SHEET_NAMES.coachingSessions]: ["SessionID", "AgentEmail", "AgentName", "CoachEmail", "CoachName", "SessionDate", "WeekNumber", "OverallScore", "FollowUpComment", "SubmissionTimestamp", "FollowUpDate", "FollowUpStatus", "AgentAcknowledgementTimestamp"],
    [SHEET_NAMES.coachingScores]: ["SessionID", "Category", "Criteria", "Score", "Comment"],
    [SHEET_NAMES.coachingTemplates]: ["TemplateName", "Category", "Criteria", "Status"],
    [SHEET_NAMES.leaveRequests]: ["RequestID", "Status", "RequestedByEmail", "RequestedByName", "LeaveType", "StartDate", "EndDate", "TotalDays", "Reason", "ActionDate", "ActionBy", "SupervisorEmail", "ActionReason", "SickNoteURL", "DirectManagerSnapshot", "ProjectManagerSnapshot"],
    [SHEET_NAMES.movementRequests]: ["MovementID", "Status", "UserToMoveEmail", "UserToMoveName", "FromSupervisorEmail", "ToSupervisorEmail", "RequestTimestamp", "ActionTimestamp", "ActionByEmail", "RequestedByEmail", "ToProjectManagerEmail"], 
    [SHEET_NAMES.roleRequests]: ["RequestID", "UserEmail", "UserName", "CurrentRole", "RequestedRole", "Justification", "RequestTimestamp", "Status", "ActionByEmail", "ActionTimestamp"],
    [SHEET_NAMES.projects]: ["ProjectID", "ProjectName", "ProjectManagerEmail", "AllowedRoles"],
    [SHEET_NAMES.projectLogs]: ["LogID", "EmployeeID", "ProjectID", "Date", "HoursLogged"],
    [SHEET_NAMES.announcements]: ["AnnouncementID", "Content", "Status", "CreatedByEmail", "Timestamp"],
    [SHEET_NAMES.assets]: ["AssetID", "Type", "AssignedTo_EmployeeID", "DateAssigned", "Status"],
    [SHEET_NAMES.overtime]: ["RequestID", "EmployeeID", "EmployeeName", "ShiftDate", "PlannedStart", "PlannedEnd", "RequestedHours", "Reason", "Status", "ManagerComment", "ActionBy", "ActionDate", "Type", "DirectManager", "ProjectManager", "DirectStatus", "ProjectStatus", "InitiatedBy"],
    
    // NEW OFFBOARDING SCHEMA
    [SHEET_NAMES.offboarding]: ["RequestID", "EmployeeID", "Name", "Email", "Type", "Reason", "Status", "DirectManager", "ProjectManager", "DirectStatus", "ProjectStatus", "HRStatus", "RequestDate", "ExitDate", "InitiatedBy"]
  };

  // Run Fixer
  for (const [sheetName, headers] of Object.entries(schema)) {
    let sheet = getOrCreateSheet(ss, sheetName);
    const lastCol = sheet.getLastColumn();
    let currentHeaders = [];
    if (lastCol > 0) currentHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    
    const missingCols = [];
    headers.forEach(h => { if (!currentHeaders.includes(h)) missingCols.push(h); });
    
    if (missingCols.length > 0) {
      const startCol = lastCol === 0 ? 1 : lastCol + 1;
      sheet.getRange(1, startCol, 1, missingCols.length).setValues([missingCols]);
    }
  }
  Logger.log("DB Fix Completed with Offboarding.");
}









// HELPER: Generates the next permanent Employee ID (e.g., KOM-1005)
function generateNextEmpID(sheet) {
  const data = sheet.getDataRange().getValues();
  let maxId = 1000; // Start from 1000
  
  for (let i = 1; i < data.length; i++) {
    const val = String(data[i][0]);
    // Check if it's a permanent ID (starts with KOM- and is NOT Pending)
    if (val.startsWith("KOM-") && !val.includes("PENDING")) {
      const parts = val.split("-");
      // Assuming format KOM-XXXX
      const num = parseInt(parts[1]); 
      if (!isNaN(num) && num > maxId) {
        maxId = num;
      }
    }
  }
  return `KOM-${maxId + 1}`;
}
