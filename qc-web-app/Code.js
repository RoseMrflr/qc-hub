/**
 * QC Operations Hub — Google Apps Script Web App
 * Data source: this bound spreadsheet (Master Log / ES Tracker / Employees / Config tabs).
 * Built for: Rosie's QC team.
 *
 * Pages: Tasks (dashboard + inline Add Entry), ES Tracker & Doc Creator,
 * Employees (roster + per-person coaching history), Config (editable dropdowns + roster).
 */

// ---------- CONFIG ----------
const SS = SpreadsheetApp.getActiveSpreadsheet();
const MASTER_SHEET_NAME = 'Master Log';
const ES_SHEET_NAME = 'ES Tracker';
const CONFIG_SHEET_NAME = 'Config';
const EMPLOYEES_SHEET_NAME = 'Employees';
const REPORTS_FOLDER_NAME = 'ES Tracker — Incident Reports';

const MASTER_HEADERS = ['ID','Date','Type','Topic','Person','Details','Issues','GoalsActionItems','DueDate','Status','Priority','Link','CreatedAt','Archived'];
const ES_HEADERS = ['ID','Agent','EmployeeName','Position','OrderID','IncidentDate','ImpactLevel','ErrorClass',
  'ErrorCategory','SanctionType','ViolationCategory','Client','Feedback','RootCause','Status','DocLink','CreatedAt'];
const MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
const EMPLOYEE_HEADERS = ['Name','RealName'].concat(MONTHS);
const CONFIG_COLUMNS = ['Type Options','Status Options','Priority Options','Impact Level','Error Class'];

const DEFAULT_TYPE_OPTIONS = ['Coaching','Observation','Clarification','Project','Calibration','Training','Huddle','Meeting','Admin Task','FGD / Escalation','Client Dispute','Audit','Other'];
const DEFAULT_STATUS_OPTIONS = ['Open','In Progress','Pending','Completed','Done','Closed'];
const DEFAULT_PRIORITY_OPTIONS = ['High','Medium','Low'];
const DEFAULT_IMPACT_OPTIONS = ['High','Medium','Low'];
const DEFAULT_ERRORCLASS_OPTIONS = ['Critical Error','Cleint Dispute','Quality / Operational Error','Minor Error','Major Error'];

const MASTER_LOG_SEED = [["1","2025-07-30","Coaching","Performance","Ron","Agenda/Topics: Check DNF, Quiz, Attendance (Punctuality, Overtime) - Done\nJump in Approvals when ops is in meeting\nInvest on backup devices (ask one by one)\nOT Reminder; be mindful of tagging. - Done\nAsk topic for calib\n\nIssues/Concerns: Why do you think productivity is reduced during auditor\nDebtor name, expiration, Juris --- make the handover accurate and complete\nBe more flexible in handling handovers - NJ\n\nCommendable: Good job on stepping in on QA / Approvals\nProactively sharing screen to support agents\nPunctual\nContinued support on confe calls\n\nGoals/Action Items: Goal: Increase Productivity","","","2026-07-27","Completed","","","2026-08-13T20:58:29.797624",""],["2","2026-07-03","FGD / Escalation","Scoreform Discussion","Rosie","Agenda/Topics: Score Form Discussion.pptx","","","2026-07-03","Completed","","Recap: Team SR 2026 July, 3 | Meeting | Microsoft Teams","2026-08-13T20:58:29.797624",""],["3","2026-07-06","Admin Task","ARE: 132805","Jarence","","","","2026-07-06","Completed","","","2026-08-13T20:58:29.797624",""],["4","2026-07-07","Huddle","Huddle","Auditors","Agenda/Topics: Welcome new L2, Goals\nSheet, Feedback, Prioritization,  Auditor as Approver\n\nGoals/Action Items: FL GS - GS Handover - Done\nFolder for Paid Copies - (Especially in Ficoso) para hindi ulit ulit yung copies - Done | Find in Completed\nBP Counting notification - Done; informed auditor\nReassigned Audit Notification - Done by TM/TL H\nDati natunog pag natapos na si agent Notification - Done by TM/TL H\n\nPrioritization:  Sabay sabay e prio - Bella county results\n                Nabibigyan ng order - From agent \n                Prio Agent > Auditor – nanonotif ang auditor pag may completed - Done by TM/TL H\n","","","2026-07-07","Completed","","L2 Huddles","2026-08-13T20:58:29.797624",""],["5","2026-07-08","Cleint Dispute","129727.0","Janet","Agenda/Topics: b. Major Error - x. Results Listing/s\n\nIssues/Concerns: Client-Reported Error\n\nGoals/Action Items: Final Written Warning","","","","Closed","High","","2026-08-13T20:58:29.797624",""],["6","2026-07-08","Cleint Dispute","129727.0","James","Agenda/Topics: b. Major Error - x. Results Listing/s\n\nIssues/Concerns: Client-Reported Error\n\nGoals/Action Items: Final Written Warning","","","","Closed","High","","2026-08-13T20:58:29.797624",""],["7","2026-07-09","Coaching","","Janelle","Issues/Concerns: Template should be first\nDo not put names on recycled templates\nIn double checking, file name should be equals to Debtor name","","","2026-07-09","Completed","","","2026-08-13T20:58:29.797624",""],["8","2026-07-09","Admin Task","ARE: Mike","Mike","","","","2026-07-09","Completed","","","2026-08-13T20:58:29.797624",""],["9","2026-07-10","Calibration","Through Date Upates Discussion","Janet","Agenda/Topics: Through Date & FC Updates​ Discussion.pptx\n\nGoals/Action Items: Guru / Sheet for State Holiday\nTo Jen: approval if we can use through date for both state and FC - Follow-up Sent\nTD Source in unified listings - Done","","","2026-07-03","Completed","","Recap: Team SR 2026 July, 10 | Meeting | Microsoft Teams","2026-08-13T20:58:29.797624",""],["10","2026-07-10","Admin Task","Montly Report","Ops","","","","2026-07-10","Completed","","","2026-08-13T20:58:29.797624",""],["11","2026-07-10","Admin Task","Through Date Upates Discussion","Team","","","","2026-07-10","Completed","","","2026-08-13T20:58:29.797624",""],["12","2026-07-13","Project","Handover Template","Me","Agenda/Topics: QA/Audit\n\nGoals/Action Items: Latest update: Added TD Source","","","","Completed","High","SR - HO Listing Template","2026-08-13T20:58:29.797624",""],["13","2026-07-15","Quality / Operational Error","132431.0","James","Agenda/Topics: PROFESSIONAL CONDUCT & WORKPLACE STANDARDS\nQuality & Performance Standards\nIncorrect guidance or review decision resulting in a critical error\n\nIssues/Concerns: Error Sanctions\n\nGoals/Action Items: Written Warning","","","","Closed","Medium","","2026-08-13T20:58:29.797624",""],["14","2026-07-17","Admin Task","Order Modification List | Order Modifications","Upper Ops","","","","2026-07-17","Completed","","","2026-08-13T20:58:29.797624",""],["15","2026-07-21","Clarification","134633.0","","Agenda/Topics: If FL Recorder is 10 years","","","","Open","High","","2026-08-13T20:58:29.797624",""],["16","2026-07-21","Clarification","Indexing vs Content","","Agenda/Topics: Indexing vs Content","","","","Open","High","","2026-08-13T20:58:29.797624",""],["17","2026-07-22","Clarification","125445.0","","Agenda/Topics: If hit or note","","","","Open","High","","2026-08-13T20:58:29.797624",""],["18","2026-07-23","Coaching","Will Issue\nCommittment\nPerformance","Florence","Agenda/Topics: I wanted to have this conversation because I believe you deserve to know where you stand. Right now, your performance is at a point where immediate and consistent improvement is needed. If things continue as they are, there could be serious consequences, including the possibility of losing your role. My goal is to help you avoid that, and I'm committed to supporting you.\n\nIssues/Concerns: Speed is unexplainable; started wiht bio breaks (it's not fair)\nAsk if works wiht other company\nOther workmates are depressed too (Pete and Benj)\nRemind her she's in recalibration notice\nViolation: bio break 16min saka pending task 11min - 07/23/2026\nDelayed Submission: NR 134727\nSuper Delayed Submission: Important - 133374","","","2026-07-24","Completed","","","2026-08-13T20:58:29.797624",""],["19","2026-07-24","Coaching","Performance","Florence","Agenda/Topics: Check DNF, Quiz, Attendance (Punctuality, Overtime)\nJump in Approvals when ops is in meeting\nInvest on backup devices (ask one by one)\nOT Reminder; be mindful of tagging. \n\nIssues/Concerns: 133854 TX - tagal\n134024 - GA - tagal\nZoning out\nAttendance: 07/07/2026 - Late 17 minutes\nAttendance: 07/09/2026 - Late 10 minutes; unresponsive\nSleeping\n\nCommendable: Good job on stepping in on QA / Approvals\n\n\nGoals/Action Items: Fix tagging (133854, 134024) & OT tagging\nAvoid late/unresponsive (7/7, 7/9)\nStay focused, avoid zoning out\nComplete DNF/Quiz/Attendance checks\nKeep up good QA/Approvals work\nPrepare backup device","","","2026-07-15","Completed","","","2026-08-13T20:58:29.797624",""],["20","2026-07-24","Calibration","FC, NJ IO Sequence, Ways of Searches","Rosie, Janet","Agenda/Topics: Calibration - July 24, 2026.pptx","","","2026-07-24","Completed","","Recap: Team SR 2026 July, 24 | Meeting | Microsoft Teams","2026-08-13T20:58:29.797624",""],["21","2026-07-24","Admin Task","Guru / Sheet for State Holiday\nTo Jen: approval if we can use through date for both state and FC","Upper Ops","","","","2026-07-24","In Progress","","","2026-08-13T20:58:29.797624",""],["22","2026-07-24","Admin Task","Coding: Guidelines; Training and Development","Team","","","","2026-07-24","Open","","","2026-08-13T20:58:29.797624",""],["23","2026-07-24","Admin Task","UCC Guru Card to mirror what's in the app","Rosie / Janet","","","","2026-07-24","Open","","","2026-08-13T20:58:29.797624",""],["24","2026-07-28","Coaching","Performance\nPrev: Goal Improve Speed in Review and PAS Time. Avoid errors.","Becky","Agenda/Topics: Check DNF, Quiz, Attendance (Punctuality, Overtime)\nJump in Approvals when ops is in meeting\nInvest on backup devices (ask one by one)\nOT Reminder; be mindful of tagging. \n\nCommendable: Good job on stepping in on QA / Approvals - GJ\nOvertime rendered\nPunctual\nMore Exposure sa UF\n\nGoals/Action Items: Goal is to become a leader she is the next in line","","","2026-07-16","Completed","","","2026-08-13T20:58:29.797624",""],["25","2026-07-29","Coaching","Performance\nMusta ang gala\nPrev: 100% error-free reviewed report, 0 disputes, 15-20% increase in productivity in terms of reviewed SFC than in previous month","Jeselle","Agenda/Topics: Check DNF, Quiz, Attendance (Punctuality, Overtime) - Done\nJump in Approvals when ops is in meeting\nInvest on backup devices (ask one by one) - Done\nOT Reminder; be mindful of tagging.  - Done\nAsk topic for calib \n\nIssues/Concerns: Sad, namatayan din\nDecline in Productivity\n\nCommendable: Good obsevation inputs during huddle - Done\nGood job on stepping in on QA / Approvals - Done\nPunctual - Done\n\nGoals/Action Items: Status of old goal:\nPrev: 100% error-free reviewed report - met \n0 disputes,  - met \n15-20% increase in productivity in terms of reviewed SFC than in previous month - goal again\ndapat di mkapasok ang TLs sa top tank gn auditors","","","","Open","","","2026-08-13T20:58:29.797624",""],["26","2026-07-31","Coaching","Performance\n2k accuracy points","Jia","Agenda/Topics: Check DNF, Quiz, Attendance (Punctuality, Overtime) - Done\nJump in Approvals when ops is in meeting - Done\nInvest on backup devices (ask one by one) - Done\nOT Reminder; be mindful of tagging. - Done\nAsk topic for calib\n\nIssues/Concerns: Kumustahin ang well-being\n\nCommendable: Good job on stepping in on QA / Approvals - GJ - Done\nPunctual - Done\nMore Exposure sa UF - Done\nCharm, good approach always sa agent - Done\nImportant - 133374 - excessive - Done\nPerformance: highest DS - Done\n\nGoals/Action Items: Next goal: \n\nExposure sa role ng L2. \n3 approvals per day. \nProductivity. \nTripple reviews\nAbove average ","","","","Completed","","","2026-08-13T20:58:29.797624",""],["27","2026-07-31","Coaching","","Agents","Issues/Concerns: Aways start in PDF converter\nRIght click nsert copied cell","","","","Open","","","2026-08-13T20:58:29.797624",""],["28","2026-07-31","Meeting","","Ops","Issues/Concerns: Auditor: Lower SFC since the agents they handle has lower SFc too sicne they are T2, they do not work on difficult orders yet - Done\nBecky's DNF - with File 132233","","","","Open","","","2026-08-13T20:58:29.797624",""],["29","2026-08-01","Huddle","Huddle","Auditors","Issues/Concerns: In porgress - make it complete\nAsk topic for calib\nRemind the use of DS - in-house filings\n\nCommendable: Agent Approach","","","","Open","","","2026-08-13T20:58:29.797624",""],["30","2026-08-01","Admin Task","Clarify 132805; hindi ginalaw ni flor pero nakta nya ang error\nNJ Courts sequence","Ops","","","","2026-08-01","Open","","","2026-08-13T20:58:29.797624",""],["31","2026-08-07","Huddle","Assure na hindi kasalanan na wala si Tara","Jia","","","","","Open","","","2026-08-13T20:58:29.797624",""],["32","2026-08-15","Coaching","Performance\nDiscuss to do good because Appraisal is nearing\nTarget: 200 DS per day","Florence","","","","","Open","","","2026-08-13T20:58:29.797624",""],["33","2026-08-17","Coaching","","Florence","Issues/Concerns: Just 1 blunder would risk her termination","","","","Open","","","2026-08-13T20:58:29.797624",""],["34","2026-08-19","Coaching","","Becky","Commendable: July third week inJuly -  stepping up","","","","Open","","","2026-08-13T20:58:29.797624",""],["35","","Project","Training Master Sheet","Me","Agenda/Topics: Training","","","","In Progress","Medium","Training Master Sheet\n","2026-08-13T20:58:29.797624",""],["36","2026-07-21","Clarification","If FL Recorder is 10 years","","Question to Client: If FL Recorder is 10 years","","","","Open","","","2026-08-13T20:58:29.797624",""],["37","2026-07-22","Clarification","CT Municipal","","Question to Client: If hit or note","","","","Open","","","2026-08-13T20:58:29.797624",""],["38","2026-07-31","Clarification","Open case with Stipulation of dismissal","","Question to Client: If hit or note or HWAN","","","","Open","","","2026-08-13T20:58:29.797624",""],["39","","Project","Handover Template","Me","Category: QA/Audit | Notes: Modify With Error form","","","","Active","High","SR - HO Listing Template","2026-08-13T20:58:29.797624",""],["40","","Project","Sheet for Approval Guide","","Category: Other","","","","Active","","","2026-08-13T20:58:29.797624",""],["41","2026-07-31","Calibration","TD: onsite, same date; Error Scoring, No need FC, MA District Court, NC Cert, MI UCC","Rosie, Jia, Janet","Materials: Calibration - July 31, 2026.pptx","","","","Done","","Recap: Team SR 2026 July, 31 | Meeting | Microsoft Teams","2026-08-13T20:58:29.797624",""],["42","","Audit","Order 136810 flagged by Mike (ARE)","Florence","","","","","Pending","","","2026-08-13T20:58:29.797624",""]];
const ES_TRACKER_SEED = [[1, "Janet", "Javelin C. Belocura", "Trainer and Workforce Coordinator", 129727, "2026-07-08", "High", "Cleint Dispute", "b. Major Error - x. Results Listing/s", "Final Written Warning", "Client-Reported Error", "GREGORY M. SWEENEY, JR., D.M.D., P.C.", "Missing results:\nB 13-7084445 FS and its UCC3s\nB 14-7361661 FS and its UCC3s\nB 20-7246767 FS and its UCC3", "This incident was caused by not utilizing all available search logic, resulting in missed filings.", "Completed", "https://drive.google.com/file/d/1xpTjj9ZZuXiQGNHEY_QhpgO5A-bC7veL/view?usp=drivesdk", ""], [2, "James", "James Brylle Dagasdas Limbauan", "L3 Auditor / Team Leader", 129727, "2026-07-08", "High", "Cleint Dispute", "b. Major Error - x. Results Listing/s", "Final Written Warning", "Client-Reported Error", "GREGORY M. SWEENEY, JR., D.M.D., P.C.", "Missing results:\nB 13-7084445 FS and its UCC3s\nB 14-7361661 FS and its UCC3s\nB 20-7246767 FS and its UCC3", "This incident was caused by not utilizing all available search logic, resulting in missed filings.", "Completed", "https://drive.google.com/file/d/1zRc4oLV1MC1LLi6Z4z-tLNXmhh1W8_Zj/view?usp=drivesdk", ""], [3, "James", "James Brylle Dagasdas Limbauan", "L3 Auditor / Team Leader", 132431, "2026-07-15", "Medium", "Quality / Operational Error", "PROFESSIONAL CONDUCT & WORKPLACE STANDARDS Quality & Performance Standards Incorrect guidance or review decision resulting in a critical error", "Written Warning", "Error Sanctions", "FT Assets LLC", "Incorrect Results Note: \"See attached\" was entered instead of \"See attached certified search.\" Incorrect Through Date (per indexing office): May 22, 2026 was entered instead of July 10, 2026. Missing Document Attachment (1/1): UCC Certificate Listing for FT Assets LLC was not attached.", "Incorrect guidance was provided to the agent despite the availability of documented guidance in the General channel and the calibration session discussion.", "Completed", "https://drive.google.com/file/d/1hBUomxpn90jlpzOZWmSwgjo6aR6NozzL/view?usp=drivesdk", ""], [4, "Mike", "Mike John Asumbrado", "T3 Agent", 135640, "2026-07-31", "High", "Critical Error", "c. Critical Error - iii. No Results Found", "Final Written Warning", "Error Sanctions", "Andrew P. Kontos", "Lee County, FL County Combo – Recorder. Score & Feedback: 80 - Reported \"No results found\", when there is a Final Judgment result # 2014000161319.", "", "Completed", "https://drive.google.com/file/d/1jfbySGpRJTd83IoTFFbgp1HBwXQ9MKH8/view?usp=drivesdk", ""], [5, "Jarence", "Jarence Jae Taduran Javier", "L1 Auditor", 135488, "2026-08-03", "High", "Cleint Dispute", "a. Minor Error - ii. Formatting (Including Matrix Errors)", "Final Written Warning", "Error Sanctions", "Llano Estacado Properties, LLC", "A lapse date was incorrectly entered for Transmitting Utility UCC No. 20240000009512. The lapse date should be left blank.", "", "Completed", "https://drive.google.com/file/d/1Tt2lAV5vs9UvmnDvXG77CeD-6cSupOto/view?usp=drivesdk", ""], [6, "Jarence", "Jarence Jae Taduran Javier", "L1 Auditor", 132805, "2026-07-06", "High", "Critical Error", "c. Critical Error - i. Debtor/Party Searched", "Final Written Warning", "Error Sanctions", "Milne Aseptics LLC", "Score & Feedback: Score: 80.00 - WA, Dept of Licensing - Incorrect Debtor/Party Searched (CE): Wrote Milne Fruit Products, Inc. instead of Milne Aseptics LLC", "WE Juris, IO: WA, Dept of Licensing. WE SFC: 1. WE SPD: 7", "Completed", "https://drive.google.com/file/d/1M56sEouviG_A2OcnoHPaBfPQLEKLN8NC/view?usp=drivesdk", ""], [7, "Becky", "Quennie Mae Quintana Pamisa", "L2 Auditor", 136354, "2026-08-09", "High", "Cleint Dispute", "b. Major Error - xi. Results Note", "Final Written Warning", "Error Sanctions", "WABILOGIC INC.", "Incorrect results note for NV SOS. It should state, \"See attached certified search – no results found,\" instead of \"No results found.\"", "The certificate was attached and cost included in feedback, but the results note was not updated.", "In Progress", "https://drive.google.com/file/d/128PB2qpebtkzHFF2hD4WRm2aLBPTGB4v/view?usp=drivesdk", ""]];
const EMPLOYEES_SEED = [["Lizzy", "Mary Luz Taño - Rico", "TM", "TM", "TM", "TM", "TM", "TM", "TM", "TM", "TM", "TM", "TM"], ["Keith", "Kenneth Lance Dela Cruz Cataquiz", "L3", "TS", "TS", "TS", "TS", "TS", "TS", "TS", "TS", "TS", "TS"], ["Rosie", "Rose Manalipon Miraflor", "L3", "QC", "QC", "QC", "QC", "QC", "QC", "QC", "QC", "QC", "QC"], ["Janet", "Javelin C. Belocura", "TWC", "TWC", "TWC", "TWC", "TWC", "TWC", "TWC", "TWC", "TWC", "TWC", "TWC"], ["Florence", "Floremy Ann Coral Sambrano", "L2", "L2", "L2", "L2", "L2", "L2", "L2", "T3", "", "", ""], ["Joan", "Joanne Dimaculangan De Villa", "T3", "T3", "L1", "T3", "Trainer", "Trainer", "T3", "T3", "", "", ""], ["Becky", "Quennie Mae Quintana Pamisa", "L2", "L2", "L2", "L2", "L2", "L2", "L2", "L2", "", "", ""], ["Henschel", "Henschel Rous Sunga Perez", "L2", "L3", "L3", "L3", "L3", "L3", "L3", "L3", "L3", "L3", "L3"], ["James", "James Brylle Dagasdas Limbauan", "L1", "L3", "L3", "L3", "L3", "L3", "L3", "L3", "L3", "L3", "L3"], ["Jarence", "Jarence Jae Taduran Javier", "T3", "T3", "L1", "L1", "T3", "T3", "L1", "L1", "", "", ""], ["Jeselle", "Jeselle Indiano", "L1", "L2", "L2", "L2", "L2", "L2", "L2", "L2", "", "", ""], ["Jia", "Benjie S. Arnidoval", "T3", "L1", "T3", "L1", "L1", "L2", "L2", "L2", "", "", ""], ["Pete", "Peter D. Balabagno", "T3", "L1", "T3", "T3", "L1", "T3", "L1", "L1", "", "", ""], ["Ron", "Shin Ray S. Acorda", "T3", "T3", "L1", "L1", "T3", "L1", "L2", "", "", "", ""], ["Mike", "Mike John Asumbrado", "Trainee", "T2", "T2", "T2", "T2", "T2", "T3", "T3", "", "", ""], ["Simon", "Simon Peter Raposas", "Trainee", "T2", "T2", "T2", "T2", "T2", "T3", "", "", "", ""], ["Wenyel", "Wenyel Anne Martinez Vidanes", "T3", "T3", "T3", "L3", "L1", "L1", "L1", "", "", "", ""], ["Annie", "Anne Jhasmine Gonzales Galang", "", "Trainee", "Trainee", "Trainee", "T2", "T2", "T1", "", "", "", ""], ["Janelle", "Janna Lea Rodriguez De Torres", "", "", "", "", "Trainee", "Trainee", "T1", "", "", "", ""], ["Gabe", "Matthew Gabriel Carpio", "", "", "", "", "Trainee", "Trainee", "T1", "", "", "", ""]];

// ---------- SETUP (runs automatically on first load) ----------
function ensureSetup_() {
  let ml = SS.getSheetByName(MASTER_SHEET_NAME);
  if (!ml) {
    ml = SS.insertSheet(MASTER_SHEET_NAME);
    ml.appendRow(MASTER_HEADERS);
    ml.getRange(1, 1, 1, MASTER_HEADERS.length).setFontWeight('bold').setBackground('#2F5597').setFontColor('#ffffff');
    MASTER_LOG_SEED.forEach(row => ml.appendRow(row));
    ml.setFrozenRows(1);
  }
  migrateMasterLogSchema_(ml);
  migrateArchivedColumn_(ml);
  let es = SS.getSheetByName(ES_SHEET_NAME);
  if (!es) {
    es = SS.insertSheet(ES_SHEET_NAME);
    es.appendRow(ES_HEADERS);
    es.getRange(1, 1, 1, ES_HEADERS.length).setFontWeight('bold').setBackground('#C0392B').setFontColor('#ffffff');
    ES_TRACKER_SEED.forEach(row => es.appendRow(row));
    es.setFrozenRows(1);
  }
  let emp = SS.getSheetByName(EMPLOYEES_SHEET_NAME);
  if (!emp) {
    emp = SS.insertSheet(EMPLOYEES_SHEET_NAME);
    emp.appendRow(EMPLOYEE_HEADERS);
    emp.getRange(1, 1, 1, EMPLOYEE_HEADERS.length).setFontWeight('bold').setBackground('#146c43').setFontColor('#ffffff');
    EMPLOYEES_SEED.forEach(row => emp.appendRow(row));
    emp.setFrozenRows(1);
  }
  let cfg = SS.getSheetByName(CONFIG_SHEET_NAME);
  if (!cfg) {
    cfg = SS.insertSheet(CONFIG_SHEET_NAME);
    cfg.appendRow(CONFIG_COLUMNS);
    cfg.getRange(1, 1, 1, CONFIG_COLUMNS.length).setFontWeight('bold');
    const lists = [DEFAULT_TYPE_OPTIONS, DEFAULT_STATUS_OPTIONS, DEFAULT_PRIORITY_OPTIONS, DEFAULT_IMPACT_OPTIONS, DEFAULT_ERRORCLASS_OPTIONS];
    const maxLen = Math.max(...lists.map(l => l.length));
    for (let i = 0; i < maxLen; i++) {
      cfg.appendRow(lists.map(l => l[i] || ''));
    }
  }
  const def = SS.getSheetByName('Sheet1');
  if (def && def.getLastRow() === 0 && SS.getSheets().length > 1) {
    SS.deleteSheet(def);
  }
}

// One-time, idempotent migration: inserts the 'Issues' / 'GoalsActionItems' columns into an
// already-existing Master Log sheet (right after 'Details'), shifting later columns right and
// preserving all existing data. No-ops once the columns are already present.
function migrateMasterLogSchema_(sheet) {
  sheet = sheet || SS.getSheetByName(MASTER_SHEET_NAME);
  if (!sheet) return;
  const lastCol = sheet.getLastColumn();
  const currentHeaders = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  if (currentHeaders.indexOf('Issues') !== -1) return; // already migrated
  const detailsIdx = currentHeaders.indexOf('Details'); // 0-based
  if (detailsIdx > -1) {
    sheet.insertColumnsAfter(detailsIdx + 1, 2);
    const target = sheet.getRange(1, detailsIdx + 2, 1, 2);
    target.setValues([['Issues', 'GoalsActionItems']]);
    target.setFontWeight('bold').setBackground('#2F5597').setFontColor('#ffffff');
  } else {
    const startCol = (lastCol || 0) + 1;
    const target = sheet.getRange(1, startCol, 1, 2);
    target.setValues([['Issues', 'GoalsActionItems']]);
    target.setFontWeight('bold').setBackground('#2F5597').setFontColor('#ffffff');
  }
}

// One-time, idempotent migration: appends an 'Archived' column to an already-existing Master
// Log sheet. Existing rows are left blank (= not archived). No-ops once already present.
function migrateArchivedColumn_(sheet) {
  sheet = sheet || SS.getSheetByName(MASTER_SHEET_NAME);
  if (!sheet) return;
  const lastCol = sheet.getLastColumn();
  const currentHeaders = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  if (currentHeaders.indexOf('Archived') !== -1) return; // already migrated
  const startCol = (lastCol || 0) + 1;
  const target = sheet.getRange(1, startCol, 1, 1);
  target.setValues([['Archived']]);
  target.setFontWeight('bold').setBackground('#2F5597').setFontColor('#ffffff');
}

// ---------- WEB APP ENTRY ----------
function doGet(e) {
  ensureSetup_();
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('QC Operations Hub')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ---------- CONFIG: DROPDOWN OPTIONS ----------
function readConfigColumn_(colIndex) {
  const cfg = SS.getSheetByName(CONFIG_SHEET_NAME);
  const lastRow = cfg.getLastRow();
  if (lastRow < 2) return [];
  return cfg.getRange(2, colIndex + 1, lastRow - 1, 1).getValues().flat().filter(v => v !== '');
}

function getOptions() {
  ensureSetup_();
  return {
    type: readConfigColumn_(0),
    status: readConfigColumn_(1),
    priority: readConfigColumn_(2),
    impact: readConfigColumn_(3),
    errorClass: readConfigColumn_(4)
  };
}

function saveOptions(lists) {
  ensureSetup_();
  const cfg = SS.getSheetByName(CONFIG_SHEET_NAME);
  const lastRow = cfg.getLastRow();
  if (lastRow > 1) cfg.getRange(2, 1, lastRow - 1, CONFIG_COLUMNS.length).clearContent();
  const cols = [lists.type || [], lists.status || [], lists.priority || [], lists.impact || [], lists.errorClass || []];
  const maxLen = Math.max(1, ...cols.map(l => l.length));
  const out = [];
  for (let i = 0; i < maxLen; i++) out.push(cols.map(l => l[i] || ''));
  if (out.length) cfg.getRange(2, 1, out.length, CONFIG_COLUMNS.length).setValues(out);
  return { success: true };
}

// ---------- EMPLOYEES ----------
function currentMonthKey_() {
  return MONTHS[new Date().getMonth()];
}

function getEmployees() {
  ensureSetup_();
  const sheet = SS.getSheetByName(EMPLOYEES_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, EMPLOYEE_HEADERS.length).getValues();
  const curMonth = currentMonthKey_();
  return values
    .filter(row => row[0])
    .map(row => {
      const obj = {};
      EMPLOYEE_HEADERS.forEach((h, i) => obj[h] = row[i]);
      obj.CurrentPosition = obj[curMonth] || '';
      return obj;
    });
}

function getEmployeeNames() {
  return getEmployees().map(e => e.Name);
}

function saveEmployees(rows) {
  ensureSetup_();
  const sheet = SS.getSheetByName(EMPLOYEES_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, EMPLOYEE_HEADERS.length).clearContent();
  const out = rows
    .filter(r => r.Name)
    .map(r => EMPLOYEE_HEADERS.map(h => r[h] || ''));
  if (out.length) sheet.getRange(2, 1, out.length, EMPLOYEE_HEADERS.length).setValues(out);
  return { success: true, count: out.length };
}

function getEmployeeHistory(name) {
  ensureSetup_();
  const employees = getEmployees();
  const employee = employees.find(e => e.Name === name) || null;

  const masterEntries = getMasterLogEntries().filter(e => (e.Person || '').trim() === name);
  const esEntries = getEsTrackerEntries().filter(e => e.Agent === name || e.EmployeeName === name);

  return { employee: employee, masterEntries: masterEntries, esEntries: esEntries };
}

// ---------- MASTER LOG ----------
function sheetToObjects_(sheet, headers) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  return values
    .map((row, i) => {
      const obj = { _row: i + 2 };
      headers.forEach((h, c) => {
        let v = row[c];
        if (v instanceof Date) v = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        obj[h] = v;
      });
      return obj;
    })
    .filter(obj => headers.some(h => obj[h] !== '' && obj[h] !== null && obj[h] !== undefined));
}

function getMasterLogEntries() {
  ensureSetup_();
  const sheet = SS.getSheetByName(MASTER_SHEET_NAME);
  return sheetToObjects_(sheet, MASTER_HEADERS).sort((a, b) => (b.Date || '').localeCompare(a.Date || ''));
}

function addMasterLogEntry(entry) {
  ensureSetup_();
  const sheet = SS.getSheetByName(MASTER_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  const ids = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().filter(v => typeof v === 'number') : [];
  const nextId = ids.length ? Math.max(...ids) + 1 : 1;
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  const row = MASTER_HEADERS.map(h => {
    if (h === 'ID') return nextId;
    if (h === 'CreatedAt') return now;
    return entry[h] || '';
  });
  sheet.appendRow(row);
  return { success: true, id: nextId };
}

function updateMasterLogStatus(id, status) {
  ensureSetup_();
  const sheet = SS.getSheetByName(MASTER_SHEET_NAME);
  const idCol = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat();
  const idx = idCol.indexOf(Number(id));
  if (idx === -1) return { success: false, error: 'Entry not found' };
  const statusCol = MASTER_HEADERS.indexOf('Status') + 1;
  sheet.getRange(idx + 2, statusCol).setValue(status);
  return { success: true };
}

// Full inline-edit save: updates any subset of fields on an existing Master Log row by ID.
// ID and CreatedAt are never overwritten.
function updateMasterLogEntry(entry) {
  ensureSetup_();
  const sheet = SS.getSheetByName(MASTER_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: false, error: 'No entries yet' };
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  const idx = ids.findIndex(v => Number(v) === Number(entry.ID));
  if (idx === -1) return { success: false, error: 'Entry not found' };
  const rowNum = idx + 2;
  const current = sheet.getRange(rowNum, 1, 1, MASTER_HEADERS.length).getValues()[0];
  MASTER_HEADERS.forEach((h, i) => {
    if (h === 'ID' || h === 'CreatedAt') return;
    if (entry[h] !== undefined) current[i] = entry[h];
  });
  sheet.getRange(rowNum, 1, 1, MASTER_HEADERS.length).setValues([current]);
  return { success: true };
}

// Archives (or unarchives) a Master Log entry by ID. Archived entries stay in the sheet
// (nothing is deleted) but are hidden from the main Tasks list and shown in the Archive tab.
function archiveMasterLogEntry(id, archived) {
  ensureSetup_();
  const sheet = SS.getSheetByName(MASTER_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: false, error: 'No entries yet' };
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  const idx = ids.findIndex(v => Number(v) === Number(id));
  if (idx === -1) return { success: false, error: 'Entry not found' };
  const archivedCol = MASTER_HEADERS.indexOf('Archived') + 1;
  sheet.getRange(idx + 2, archivedCol).setValue(archived ? 'Yes' : '');
  return { success: true };
}

// "Mass schedule": generates N future Master Log entries from a template, one per month, on
// (or clamped to the last day of the month for) the given day-of-month. Starts this month if
// that day hasn't passed yet, otherwise starts next month. Returns how many rows were created.
function addRecurringMasterLogEntries(template, dayOfMonth, months) {
  ensureSetup_();
  const sheet = SS.getSheetByName(MASTER_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  const ids = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().filter(v => typeof v === 'number') : [];
  let nextId = ids.length ? Math.max(...ids) + 1 : 1;
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  const today = new Date();
  const day = Math.max(1, Math.min(31, Number(dayOfMonth) || 1));
  const n = Math.max(1, Math.min(60, Number(months) || 12));
  const startOffset = (today.getDate() > day) ? 1 : 0;
  const rows = [];
  const dates = [];
  for (let i = 0; i < n; i++) {
    const targetMonthIndex = today.getMonth() + startOffset + i;
    const targetYear = today.getFullYear() + Math.floor(targetMonthIndex / 12);
    const mIdx = ((targetMonthIndex % 12) + 12) % 12;
    const lastDayOfMonth = new Date(targetYear, mIdx + 1, 0).getDate();
    const useDay = Math.min(day, lastDayOfMonth);
    const dateStr = Utilities.formatDate(new Date(targetYear, mIdx, useDay), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const row = MASTER_HEADERS.map(h => {
      if (h === 'ID') return nextId;
      if (h === 'CreatedAt') return now;
      if (h === 'Date') return dateStr;
      if (h === 'DueDate') return dateStr;
      return template[h] || '';
    });
    rows.push(row);
    dates.push(dateStr);
    nextId++;
  }
  if (rows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, MASTER_HEADERS.length).setValues(rows);
  }
  return { success: true, count: rows.length, dates: dates };
}

// ---------- TASKS PAGE DATA ----------
function getDashboardData() {
  ensureSetup_();
  const entries = getMasterLogEntries();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const weekFromNow = Utilities.formatDate(new Date(Date.now() + 7*24*60*60*1000), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  const openStatuses = ['Open', 'In Progress', 'Pending'];
  const doneStatuses = ['Completed', 'Done', 'Closed'];

  const open = entries.filter(e => openStatuses.includes(e.Status));
  const completed = entries.filter(e => doneStatuses.includes(e.Status));
  const overdue = entries.filter(e => e.DueDate && e.DueDate < today && openStatuses.includes(e.Status));
  const dueThisWeek = entries.filter(e => e.DueDate && e.DueDate >= today && e.DueDate <= weekFromNow && openStatuses.includes(e.Status));

  const byType = {};
  entries.forEach(e => { byType[e.Type] = (byType[e.Type] || 0) + 1; });

  const esEntries = getEsTrackerEntries();
  const esOpen = esEntries.filter(e => !['Completed','Closed'].includes(e.Status));

  return {
    total: entries.length,
    open: open.length,
    completed: completed.length,
    overdue: overdue.length,
    dueThisWeek: dueThisWeek.length,
    byType: byType,
    recent: entries.slice(0, 12),
    overdueList: overdue,
    esTotal: esEntries.length,
    esOpen: esOpen.length
  };
}

// ---------- ES TRACKER ----------
function getEsTrackerEntries() {
  ensureSetup_();
  const sheet = SS.getSheetByName(ES_SHEET_NAME);
  return sheetToObjects_(sheet, ES_HEADERS).sort((a, b) => (b.IncidentDate || '').localeCompare(a.IncidentDate || ''));
}

function addEsTrackerEntry(entry) {
  ensureSetup_();
  const sheet = SS.getSheetByName(ES_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  const ids = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().filter(v => typeof v === 'number') : [];
  const nextId = ids.length ? Math.max(...ids) + 1 : 1;
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  const row = ES_HEADERS.map(h => {
    if (h === 'ID') return nextId;
    if (h === 'CreatedAt') return now;
    if (h === 'DocLink') return '';
    return entry[h] || '';
  });
  sheet.appendRow(row);
  return { success: true, id: nextId };
}

// ---------- DOC CREATOR (Incident Report from ES Tracker) ----------
function getOrCreateReportsFolder_() {
  const folders = DriveApp.getFoldersByName(REPORTS_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(REPORTS_FOLDER_NAME);
}

function createIncidentDoc(esId) {
  ensureSetup_();
  const sheet = SS.getSheetByName(ES_SHEET_NAME);
  const entries = sheetToObjects_(sheet, ES_HEADERS);
  const record = entries.find(e => Number(e.ID) === Number(esId));
  if (!record) return { success: false, error: 'ES Tracker entry not found' };

  const folder = getOrCreateReportsFolder_();
  const fileName = `Incident Report - ${record.EmployeeName || record.Agent} - Order ${record.OrderID} - ${record.IncidentDate}`;
  const doc = DocumentApp.create(fileName);
  const body = doc.getBody();
  body.clear();

  body.appendParagraph('ERROR / SANCTION INCIDENT REPORT').setHeading(DocumentApp.ParagraphHeading.TITLE);
  body.appendParagraph('Search Recorder — Quality Control').setHeading(DocumentApp.ParagraphHeading.SUBTITLE);
  body.appendHorizontalRule();

  const infoTable = [
    ['Employee Name', record.EmployeeName || ''],
    ['Agent / Alias', record.Agent || ''],
    ['Position', record.Position || ''],
    ['Order ID', String(record.OrderID || '')],
    ['Incident Date', record.IncidentDate || ''],
    ['Level of Impact', record.ImpactLevel || ''],
    ['Error Class', record.ErrorClass || ''],
    ['Sanction Type', record.SanctionType || ''],
    ['Violation Category', record.ViolationCategory || ''],
    ['Client / SN', record.Client || ''],
    ['Status', record.Status || '']
  ];
  const table = body.appendTable(infoTable);
  for (let i = 0; i < infoTable.length; i++) {
    table.getRow(i).getCell(0).setBold(true).setWidth(160);
  }

  body.appendParagraph('');
  body.appendParagraph('Error Category').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph(record.ErrorCategory || 'N/A');

  body.appendParagraph('Feedback / Description of Incident').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph(record.Feedback || 'N/A');

  body.appendParagraph('Root Cause').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph(record.RootCause || 'N/A');

  body.appendParagraph('');
  body.appendHorizontalRule();
  body.appendParagraph('Acknowledgement').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('I acknowledge that I have read and understood the contents of this incident report.');
  body.appendParagraph('');
  body.appendParagraph('Employee Signature: ______________________________     Date: ______________');
  body.appendParagraph('');
  body.appendParagraph('QC / TL Signature: ______________________________     Date: ______________');

  doc.saveAndClose();

  const file = DriveApp.getFileById(doc.getId());
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);

  const url = file.getUrl();
  const docCol = ES_HEADERS.indexOf('DocLink') + 1;
  const rowIndex = entries.findIndex(e => Number(e.ID) === Number(esId)) + 2;
  sheet.getRange(rowIndex, docCol).setValue(url);

  return { success: true, url: url, fileName: fileName };
}
