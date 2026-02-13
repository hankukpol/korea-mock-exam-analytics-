// ============================================
// 현장 모의고사 성적 분석 시스템 - 백엔드
// ============================================

var CONFIG = {
  SHEET_STUDENTS_POLICE: 'Students_Police',
  SHEET_STUDENTS_FIRE: 'Students_Fire',
  SHEET_EXAM: 'Exams',
  SHEET_LOG: '로그인기록',
  SHEET_SCORE_POLICE: 'Scores_Police',
  SHEET_SCORE_FIRE: 'Scores_Fire',
  SHEET_SCORE_LEGACY: 'Scores',
  SHEET_QUESTIONS_POLICE: 'ItemAnalysis_Police',
  SHEET_QUESTIONS_FIRE: 'ItemAnalysis_Fire',
  SHEET_QUESTIONS_LEGACY: 'ItemAnalysis',
  SHEET_RESPONSE_CANDIDATES: ['StudentResponses', 'Responses', '문항응답', '답안데이터', 'StudentAnswer'],
  ADMIN_PASSWORD_DEFAULT: 'admin1234',
  ADMIN_SESSION_TTL_SECONDS: 1800
};
var SPLIT_MIGRATION_PROPERTY_KEY = 'SPLIT_SCORE_ITEM_MIGRATION_V1';

var STUDENT_HEADERS = [
  '수험번호', '이름', '직렬', '지원지역', '비밀번호',
  '연락처', '생년월일', '응시분야', '응시계급', '고사실', '시험과목', '비고'
];
var EXAM_HEADERS = ['시험ID', '시험명', '시행일', '과목1', '과목2', '과목3'];
var SCORE_HEADERS = ['시험ID', '수험번호', '과목1점수', '과목2점수', '과목3점수', '총점', '평균'];
var QUESTION_HEADERS = ['시험ID', '문항번호', '과목명', '정답', '배점', '정답률', '난이도'];
var LOG_HEADERS = ['이름', '연락처', '접속시간', '작업', 'IP정보'];

var STUDENT_COL = {
  ID: 0,
  NAME: 1,
  TYPE: 2,
  REGION: 3,
  PASSWORD: 4,
  PHONE: 5,
  BIRTH: 6,
  FIELD: 7,
  RANK: 8,
  LOCATION: 9,
  SUBJECT: 10,
  NOTE: 11
};

var EXAM_COL = {
  ID: 0,
  NAME: 1,
  DATE: 2,
  S1: 3,
  S2: 4,
  S3: 5
};

var SCORE_COL = {
  EXAM_ID: 0,
  STUDENT_ID: 1,
  S1: 2,
  S2: 3,
  S3: 4,
  TOTAL: 5,
  AVG: 6,
  ITEM_START: 7
};

var ITEM_COL = {
  EXAM_ID: 0,
  NUM: 1,
  SUBJECT: 2,
  ANSWER: 3,
  POINTS: 4,
  CORRECT_RATE: 5,
  DIFFICULTY: 6
};

// ============================================
// 라우팅
// ============================================

function doGet(e) {
  var page = e && e.parameter && e.parameter.page ? e.parameter.page : 'index';
  var templateName = page === 'admin' ? 'Admin' : 'Index';
  var title = page === 'admin' ? '현장모의고사 - 관리자' : '현장모의고사 성적 분석 시스템';

  return HtmlService.createTemplateFromFile(templateName)
    .evaluate()
    .setTitle(title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('성적 시스템 관리')
    .addItem('데이터 초기화 (시트 생성 + 목업 데이터)', 'initializeSampleData')
    .addToUi();
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================
// 시트 헬퍼
// ============================================

function getSheet(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    initSheet(sheet, sheetName);
  }
  if (sheetName === CONFIG.SHEET_STUDENTS_POLICE || sheetName === CONFIG.SHEET_STUDENTS_FIRE) {
    ensureStudentsSchema_(sheet, sheetName);
  }
  if (isQuestionSheetName_(sheetName)) {
    migrateItemAnalysisAddSubjectColumn_(sheet);
  }
  return sheet;
}

function initSheet(sheet, sheetName) {
  if (sheetName === CONFIG.SHEET_STUDENTS_POLICE || sheetName === CONFIG.SHEET_STUDENTS_FIRE) {
    sheet.getRange(1, 1, 1, STUDENT_HEADERS.length).setValues([STUDENT_HEADERS]);
    styleHeader_(sheet, STUDENT_HEADERS.length, '#4285f4');
    return;
  }

  if (sheetName === CONFIG.SHEET_EXAM) {
    sheet.getRange(1, 1, 1, EXAM_HEADERS.length).setValues([EXAM_HEADERS]);
    styleHeader_(sheet, EXAM_HEADERS.length, '#34a853');
    return;
  }

  if (isScoreSheetName_(sheetName)) {
    sheet.getRange(1, 1, 1, SCORE_HEADERS.length).setValues([SCORE_HEADERS]);
    styleHeader_(sheet, SCORE_HEADERS.length, '#fbbc04');
    return;
  }

  if (isQuestionSheetName_(sheetName)) {
    sheet.getRange(1, 1, 1, QUESTION_HEADERS.length).setValues([QUESTION_HEADERS]);
    styleHeader_(sheet, QUESTION_HEADERS.length, '#673ab7');
    return;
  }

  if (sheetName === CONFIG.SHEET_LOG) {
    sheet.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS]);
    styleHeader_(sheet, LOG_HEADERS.length, '#ea4335');
  }
}

function styleHeader_(sheet, colCount, bgColor) {
  sheet.getRange(1, 1, 1, colCount)
    .setFontWeight('bold')
    .setBackground(bgColor)
    .setFontColor('#ffffff');
  sheet.setFrozenRows(1);
}

function ensureStudentsSchema_(sheet, sheetName) {
  if (sheet.getLastRow() < 1) {
    initSheet(sheet, sheetName);
    return;
  }

  if (sheet.getLastColumn() < STUDENT_HEADERS.length) {
    sheet.insertColumnsAfter(sheet.getLastColumn(), STUDENT_HEADERS.length - sheet.getLastColumn());
  }
  sheet.getRange(1, 1, 1, STUDENT_HEADERS.length).setValues([STUDENT_HEADERS]);
  styleHeader_(sheet, STUDENT_HEADERS.length, '#4285f4');
}

function migrateItemAnalysisAddSubjectColumn_(sheet) {
  if (!sheet) return false;

  var sheetName = sheet.getName();
  if (!isQuestionSheetName_(sheetName)) return false;

  if (sheet.getLastRow() < 1) {
    initSheet(sheet, sheetName);
    return true;
  }

  if (sheet.getLastColumn() < 2) {
    if (sheet.getLastColumn() < QUESTION_HEADERS.length) {
      sheet.insertColumnsAfter(sheet.getLastColumn(), QUESTION_HEADERS.length - sheet.getLastColumn());
    }
    sheet.getRange(1, 1, 1, QUESTION_HEADERS.length).setValues([QUESTION_HEADERS]);
    styleHeader_(sheet, QUESTION_HEADERS.length, '#673ab7');
    return true;
  }

  var headerCols = Math.max(sheet.getLastColumn(), QUESTION_HEADERS.length);
  var header = sheet.getRange(1, 1, 1, headerCols).getValues()[0];
  var first = String(header[0] || '').trim();
  var second = String(header[1] || '').trim();
  var third = String(header[2] || '').trim();
  var changed = false;

  // 기존 6컬럼 스키마: 시험ID | 문항번호 | 정답 | ...
  // 3번째 컬럼(C)에 과목명 컬럼을 삽입해 하위 호환을 유지한다.
  if (first === '시험ID' && second === '문항번호' && third !== '과목명') {
    sheet.insertColumnBefore(3);
    changed = true;
  }

  if (sheet.getLastColumn() < QUESTION_HEADERS.length) {
    sheet.insertColumnsAfter(sheet.getLastColumn(), QUESTION_HEADERS.length - sheet.getLastColumn());
    changed = true;
  }

  sheet.getRange(1, 1, 1, QUESTION_HEADERS.length).setValues([QUESTION_HEADERS]);
  styleHeader_(sheet, QUESTION_HEADERS.length, '#673ab7');
  return changed;
}

function getStudentSheetName_(type) {
  if (type === '경찰') return CONFIG.SHEET_STUDENTS_POLICE;
  if (type === '소방') return CONFIG.SHEET_STUDENTS_FIRE;
  return null;
}

function getScoreSheetNameByType_(type) {
  if (type === '경찰') return CONFIG.SHEET_SCORE_POLICE;
  if (type === '소방') return CONFIG.SHEET_SCORE_FIRE;
  return '';
}

function getQuestionSheetNameByType_(type) {
  if (type === '경찰') return CONFIG.SHEET_QUESTIONS_POLICE;
  if (type === '소방') return CONFIG.SHEET_QUESTIONS_FIRE;
  return '';
}

function isScoreSheetName_(sheetName) {
  return sheetName === CONFIG.SHEET_SCORE_POLICE ||
    sheetName === CONFIG.SHEET_SCORE_FIRE ||
    sheetName === CONFIG.SHEET_SCORE_LEGACY;
}

function isQuestionSheetName_(sheetName) {
  return sheetName === CONFIG.SHEET_QUESTIONS_POLICE ||
    sheetName === CONFIG.SHEET_QUESTIONS_FIRE ||
    sheetName === CONFIG.SHEET_QUESTIONS_LEGACY;
}

function trimRightEmptyRow_(row) {
  var end = row.length;
  while (end > 0 && String(row[end - 1] || '') === '') end--;
  return row.slice(0, end);
}

function appendRowsWithPadding_(sheet, rows) {
  if (!rows || rows.length === 0) return 0;
  var maxLen = 0;
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].length > maxLen) maxLen = rows[i].length;
  }
  if (maxLen <= 0) return 0;

  if (sheet.getLastColumn() < maxLen) {
    sheet.insertColumnsAfter(sheet.getLastColumn(), maxLen - sheet.getLastColumn());
  }

  var normalized = rows.map(function(row) {
    var out = row.slice(0, maxLen);
    while (out.length < maxLen) out.push('');
    return out;
  });

  sheet.getRange(sheet.getLastRow() + 1, 1, normalized.length, maxLen).setValues(normalized);
  return normalized.length;
}

function buildExamTypeMapForMigration_() {
  var map = {};
  var examSheet = getSheet(CONFIG.SHEET_EXAM);
  var rows = examSheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    var examId = String(rows[i][EXAM_COL.ID] || '').trim();
    if (!examId) continue;
    var examType = inferExamType_(examId, rows[i][EXAM_COL.S1]);
    if (examType) map[examId] = examType;
  }
  return map;
}

function buildStudentTypeMapForMigration_() {
  var map = {};
  var policeRecords = getStudentRecords_(getSheet(CONFIG.SHEET_STUDENTS_POLICE));
  var fireRecords = getStudentRecords_(getSheet(CONFIG.SHEET_STUDENTS_FIRE));
  var all = policeRecords.concat(fireRecords);
  for (var i = 0; i < all.length; i++) {
    var sid = String(all[i].studentId || '').trim();
    var t = String(all[i].type || '').trim();
    if (sid && t && !map[sid]) map[sid] = t;
  }
  return map;
}

function resolveTypeForMigration_(examId, studentId, examTypeMap, studentTypeMap) {
  var eId = String(examId || '').trim();
  if (!eId) return '';
  var byExamMap = String((examTypeMap && examTypeMap[eId]) || '').trim();
  if (byExamMap === '경찰' || byExamMap === '소방') return byExamMap;

  var byExamId = inferExamType_(eId, '');
  if (byExamId === '경찰' || byExamId === '소방') return byExamId;

  var sid = String(studentId || '').trim();
  var byStudent = String((studentTypeMap && studentTypeMap[sid]) || '').trim();
  if (byStudent === '경찰' || byStudent === '소방') return byStudent;

  return '';
}

function migrateLegacyScoresToSplit_(legacySheet, policeSheet, fireSheet, examTypeMap, studentTypeMap) {
  if (!legacySheet || legacySheet.getLastRow() < 2) {
    return { police: 0, fire: 0, unknown: 0 };
  }

  var policeKeys = {};
  var fireKeys = {};

  var pRows = policeSheet.getDataRange().getValues();
  for (var p = 1; p < pRows.length; p++) {
    var pExamId = String(pRows[p][SCORE_COL.EXAM_ID] || '').trim();
    var pSid = String(pRows[p][SCORE_COL.STUDENT_ID] || '').trim();
    if (pExamId && pSid) policeKeys[pExamId + '::' + pSid] = true;
  }
  var fRows = fireSheet.getDataRange().getValues();
  for (var f = 1; f < fRows.length; f++) {
    var fExamId = String(fRows[f][SCORE_COL.EXAM_ID] || '').trim();
    var fSid = String(fRows[f][SCORE_COL.STUDENT_ID] || '').trim();
    if (fExamId && fSid) fireKeys[fExamId + '::' + fSid] = true;
  }

  var policeAppend = [];
  var fireAppend = [];
  var unknownCount = 0;
  var legacyRows = legacySheet.getDataRange().getValues();
  for (var i = 1; i < legacyRows.length; i++) {
    var row = trimRightEmptyRow_(legacyRows[i]);
    var examId = String(row[SCORE_COL.EXAM_ID] || '').trim();
    var sid = String(row[SCORE_COL.STUDENT_ID] || '').trim();
    if (!examId || !sid) continue;

    var key = examId + '::' + sid;
    var resolvedType = resolveTypeForMigration_(examId, sid, examTypeMap, studentTypeMap);
    if (resolvedType === '경찰') {
      if (policeKeys[key]) continue;
      policeKeys[key] = true;
      policeAppend.push(row);
      continue;
    }
    if (resolvedType === '소방') {
      if (fireKeys[key]) continue;
      fireKeys[key] = true;
      fireAppend.push(row);
      continue;
    }
    unknownCount++;
  }

  var addedPolice = appendRowsWithPadding_(policeSheet, policeAppend);
  var addedFire = appendRowsWithPadding_(fireSheet, fireAppend);
  return { police: addedPolice, fire: addedFire, unknown: unknownCount };
}

function migrateLegacyItemsToSplit_(legacySheet, policeSheet, fireSheet, examTypeMap) {
  if (!legacySheet) {
    return { police: 0, fire: 0, unknown: 0 };
  }

  migrateItemAnalysisAddSubjectColumn_(legacySheet);
  migrateItemAnalysisAddSubjectColumn_(policeSheet);
  migrateItemAnalysisAddSubjectColumn_(fireSheet);

  if (legacySheet.getLastRow() < 2) {
    return { police: 0, fire: 0, unknown: 0 };
  }

  var policeKeys = {};
  var fireKeys = {};

  var pRows = policeSheet.getDataRange().getValues();
  for (var p = 1; p < pRows.length; p++) {
    var pExamId = String(pRows[p][ITEM_COL.EXAM_ID] || '').trim();
    var pNum = Number(pRows[p][ITEM_COL.NUM]) || 0;
    if (pExamId && pNum) policeKeys[pExamId + '::' + pNum] = true;
  }
  var fRows = fireSheet.getDataRange().getValues();
  for (var f = 1; f < fRows.length; f++) {
    var fExamId = String(fRows[f][ITEM_COL.EXAM_ID] || '').trim();
    var fNum = Number(fRows[f][ITEM_COL.NUM]) || 0;
    if (fExamId && fNum) fireKeys[fExamId + '::' + fNum] = true;
  }

  var policeAppend = [];
  var fireAppend = [];
  var unknownCount = 0;
  var legacyRows = legacySheet.getDataRange().getValues();
  for (var i = 1; i < legacyRows.length; i++) {
    var row = trimRightEmptyRow_(legacyRows[i]);
    var examId = String(row[ITEM_COL.EXAM_ID] || '').trim();
    var qNum = Number(row[ITEM_COL.NUM]) || 0;
    if (!examId || !qNum) continue;

    var resolvedType = resolveTypeForMigration_(examId, '', examTypeMap, null);
    var key = examId + '::' + qNum;
    if (resolvedType === '경찰') {
      if (policeKeys[key]) continue;
      policeKeys[key] = true;
      policeAppend.push(row);
      continue;
    }
    if (resolvedType === '소방') {
      if (fireKeys[key]) continue;
      fireKeys[key] = true;
      fireAppend.push(row);
      continue;
    }
    unknownCount++;
  }

  var addedPolice = appendRowsWithPadding_(policeSheet, policeAppend);
  var addedFire = appendRowsWithPadding_(fireSheet, fireAppend);
  return { police: addedPolice, fire: addedFire, unknown: unknownCount };
}

function ensureSeparatedExamDataReady_() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty(SPLIT_MIGRATION_PROPERTY_KEY) === '1') return;

  var lock = LockService.getScriptLock();
  var locked = false;
  try {
    lock.waitLock(10000);
    locked = true;
  } catch (lockErr) {
    return;
  }

  try {
    if (props.getProperty(SPLIT_MIGRATION_PROPERTY_KEY) === '1') return;

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var policeScoreSheet = getSheet(CONFIG.SHEET_SCORE_POLICE);
    var fireScoreSheet = getSheet(CONFIG.SHEET_SCORE_FIRE);
    var policeItemSheet = getSheet(CONFIG.SHEET_QUESTIONS_POLICE);
    var fireItemSheet = getSheet(CONFIG.SHEET_QUESTIONS_FIRE);

    var examTypeMap = buildExamTypeMapForMigration_();
    var studentTypeMap = buildStudentTypeMapForMigration_();

    var legacyScoreSheet = ss.getSheetByName(CONFIG.SHEET_SCORE_LEGACY);
    var legacyItemSheet = ss.getSheetByName(CONFIG.SHEET_QUESTIONS_LEGACY);

    migrateLegacyScoresToSplit_(legacyScoreSheet, policeScoreSheet, fireScoreSheet, examTypeMap, studentTypeMap);
    migrateLegacyItemsToSplit_(legacyItemSheet, policeItemSheet, fireItemSheet, examTypeMap);

    props.setProperty(SPLIT_MIGRATION_PROPERTY_KEY, '1');
  } finally {
    if (locked) lock.releaseLock();
  }
}

function rerunSeparatedExamDataMigration(adminToken) {
  requireAdminToken_(adminToken);
  PropertiesService.getScriptProperties().deleteProperty(SPLIT_MIGRATION_PROPERTY_KEY);
  ensureSeparatedExamDataReady_();
  return { success: true, message: '분리 시트 자동 설정/이관을 다시 실행했습니다.' };
}

function getScoreSheetsForType_(type) {
  ensureSeparatedExamDataReady_();
  var names = [];
  var preferred = getScoreSheetNameByType_(String(type || '').trim());
  if (preferred) {
    names.push(preferred);
  } else {
    names.push(CONFIG.SHEET_SCORE_POLICE, CONFIG.SHEET_SCORE_FIRE);
  }

  var sheets = names.map(function(name) { return getSheet(name); });
  var legacy = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_SCORE_LEGACY);
  if (legacy) sheets.push(legacy);
  return sheets;
}

function getQuestionSheetsForType_(type) {
  ensureSeparatedExamDataReady_();
  var names = [];
  var preferred = getQuestionSheetNameByType_(String(type || '').trim());
  if (preferred) {
    names.push(preferred);
  } else {
    names.push(CONFIG.SHEET_QUESTIONS_POLICE, CONFIG.SHEET_QUESTIONS_FIRE);
  }

  var sheets = names.map(function(name) { return getSheet(name); });
  var legacy = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_QUESTIONS_LEGACY);
  if (legacy) {
    migrateItemAnalysisAddSubjectColumn_(legacy);
    sheets.push(legacy);
  }
  return sheets;
}

function collectScoreRowsByExam_(examId, examType) {
  var scoreRows = [SCORE_HEADERS.slice()];
  var seenStudentMap = {};
  var targetExamId = String(examId || '').trim();
  var scoreSheets = getScoreSheetsForType_(examType);

  for (var s = 0; s < scoreSheets.length; s++) {
    var data = scoreSheets[s].getDataRange().getValues();
    for (var r = 1; r < data.length; r++) {
      var rowExamId = String(data[r][SCORE_COL.EXAM_ID] || '').trim();
      if (rowExamId !== targetExamId) continue;

      var sid = String(data[r][SCORE_COL.STUDENT_ID] || '').trim();
      if (!sid || seenStudentMap[sid]) continue;
      seenStudentMap[sid] = true;
      scoreRows.push(data[r]);
    }
  }
  return scoreRows;
}

function collectItemRowsByExam_(examId, examType) {
  var itemRows = [QUESTION_HEADERS.slice()];
  var seenQuestionMap = {};
  var targetExamId = String(examId || '').trim();
  var questionSheets = getQuestionSheetsForType_(examType);

  for (var s = 0; s < questionSheets.length; s++) {
    var data = questionSheets[s].getDataRange().getValues();
    for (var r = 1; r < data.length; r++) {
      var rowExamId = String(data[r][ITEM_COL.EXAM_ID] || '').trim();
      if (rowExamId !== targetExamId) continue;

      var qNum = Number(data[r][ITEM_COL.NUM]) || 0;
      if (!qNum || seenQuestionMap[qNum]) continue;
      seenQuestionMap[qNum] = true;
      itemRows.push(data[r]);
    }
  }

  if (itemRows.length > 2) {
    var sortedBody = itemRows.slice(1).sort(function(a, b) {
      return (Number(a[ITEM_COL.NUM]) || 0) - (Number(b[ITEM_COL.NUM]) || 0);
    });
    itemRows = [itemRows[0]].concat(sortedBody);
  }

  return itemRows;
}

function isStudentSheetName_(sheetName) {
  return sheetName === CONFIG.SHEET_STUDENTS_POLICE || sheetName === CONFIG.SHEET_STUDENTS_FIRE;
}

function resolveStudentSheetAndRow_(sheetName, rowIndex) {
  if (!isStudentSheetName_(sheetName)) {
    throw new Error('허용되지 않은 시트입니다.');
  }

  var rowNum = Number(rowIndex);
  if (!isFinite(rowNum) || rowNum % 1 !== 0 || rowNum < 2) {
    throw new Error('유효하지 않은 행 번호입니다.');
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('지정된 시트를 찾을 수 없습니다.');
  }
  if (rowNum > sheet.getLastRow()) {
    throw new Error('유효하지 않은 행 번호입니다.');
  }

  return {
    sheet: sheet,
    rowIndex: rowNum
  };
}

// ============================================
// 인증 헬퍼
// ============================================

function getAdminPassword_() {
  var fromProperty = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  return fromProperty || CONFIG.ADMIN_PASSWORD_DEFAULT;
}

function createAdminToken_() {
  var token = Utilities.getUuid();
  CacheService.getScriptCache().put('ADMIN_TOKEN_' + token, '1', CONFIG.ADMIN_SESSION_TTL_SECONDS);
  return token;
}

function requireAdminToken_(adminToken) {
  if (!adminToken) {
    throw new Error('관리자 인증이 필요합니다. 다시 로그인해 주세요.');
  }
  var exists = CacheService.getScriptCache().get('ADMIN_TOKEN_' + adminToken);
  if (!exists) {
    throw new Error('관리자 세션이 만료되었습니다. 다시 로그인해 주세요.');
  }
}

// ============================================
// 학생 데이터 헬퍼
// ============================================

function toStudentRecord_(row, rowIndex) {
  var legacyAdminRow =
    row.length >= 5 &&
    String(row[0] || '').trim() !== '' &&
    String(row[3] || '').trim() !== '' &&
    (String(row[4] || '').trim() === '경찰' || String(row[4] || '').trim() === '소방');

  if (legacyAdminRow) {
    var legacyId = String(row[3] || '').trim();
    var legacyPhone = String(row[1] || '').trim();
    return {
      rowIndex: rowIndex,
      studentId: legacyId,
      name: String(row[0] || '').trim(),
      type: String(row[4] || '').trim(),
      region: String(row[5] || '').trim(),
      password: defaultPassword_(legacyPhone, legacyId),
      phone: legacyPhone,
      birthDate: row[2] || '',
      examField: row[6] || '',
      examRank: row[7] || '',
      examLocation: row[8] || '',
      examSubject: row[9] || '',
      note: row[11] || ''
    };
  }

  return {
    rowIndex: rowIndex,
    studentId: String(row[STUDENT_COL.ID] || '').trim(),
    name: String(row[STUDENT_COL.NAME] || '').trim(),
    type: String(row[STUDENT_COL.TYPE] || '').trim(),
    region: String(row[STUDENT_COL.REGION] || '').trim(),
    password: String(row[STUDENT_COL.PASSWORD] || '').trim(),
    phone: String(row[STUDENT_COL.PHONE] || '').trim(),
    birthDate: row[STUDENT_COL.BIRTH] || '',
    examField: row[STUDENT_COL.FIELD] || '',
    examRank: row[STUDENT_COL.RANK] || '',
    examLocation: row[STUDENT_COL.LOCATION] || '',
    examSubject: row[STUDENT_COL.SUBJECT] || '',
    note: row[STUDENT_COL.NOTE] || ''
  };
}

function getStudentRecords_(sheet) {
  var data = sheet.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < data.length; i++) {
    var rec = toStudentRecord_(data[i], i + 1);
    if (!rec.studentId && !rec.name) continue;
    list.push(rec);
  }
  return list;
}

function defaultPassword_(phone, studentId) {
  var digits = String(phone || '').replace(/\D/g, '');
  if (digits.length >= 4) return digits.slice(-4);
  var sid = String(studentId || '').replace(/\D/g, '');
  if (sid.length >= 4) return sid.slice(-4);
  return '0000';
}

function buildStudentRow_(data, existingPassword) {
  var studentId = String(data.examNumber || data.studentId || '').trim();
  var name = String(data.name || '').trim();
  var examType = String(data.examType || data.type || '').trim();
  var region = String(data.examRegion || data.region || '').trim();
  var phone = String(data.phone || '').trim();
  var password = String(data.password || existingPassword || defaultPassword_(phone, studentId)).trim();

  return [
    studentId,
    name,
    examType,
    region,
    password,
    phone,
    data.birthDate || '',
    data.examField || '',
    data.examRank || '',
    data.examLocation || '',
    data.examSubject || '',
    data.note || ''
  ];
}

function inferExamType_(examId, firstSubject) {
  var id = String(examId || '').toUpperCase();
  if (id.indexOf('P') === 0) return '경찰';
  if (id.indexOf('F') === 0) return '소방';

  var s = String(firstSubject || '');
  if (s.indexOf('경찰') >= 0 || s.indexOf('형사') >= 0 || s.indexOf('헌법') >= 0) return '경찰';
  if (s.indexOf('소방') >= 0 || s.indexOf('행정법') >= 0) return '소방';
  return '';
}

function isPlaceholderSubject_(name) {
  return /^과목\d*$/.test(String(name || '').trim());
}

function normalizeSubjectList_(subjects) {
  return (subjects || [])
    .map(function(s) { return String(s || '').trim(); })
    .filter(function(s) { return !!s && !isPlaceholderSubject_(s); });
}

function mergeSubjects_(existing, defaults, length) {
  var result = [];
  for (var i = 0; i < length; i++) {
    var v = String((existing && existing[i]) || (defaults && defaults[i]) || ('과목' + (i + 1))).trim();
    result.push(v || ('과목' + (i + 1)));
  }
  return result;
}

function inferFireTrack_(studentInfo, examInfo) {
  var sourceText = [
    (studentInfo && studentInfo.examField) || '',
    (studentInfo && studentInfo.examRank) || '',
    (studentInfo && studentInfo.examSubject) || '',
    (examInfo && examInfo.name) || '',
    (examInfo && examInfo.subjects ? examInfo.subjects.join(' ') : '')
  ].join(' ');

  if (/구급/.test(sourceText)) return 'fire_paramedic_special';
  if (/구조|학과/.test(sourceText)) return 'fire_rescue_academic_special';
  if (/공채/.test(sourceText)) return 'fire_public';
  if (/경채/.test(sourceText)) return 'fire_special';
  return 'fire_unknown';
}

function resolveItemCountsBySubject_(subjects, track, fallbackCounts) {
  var list = Array.isArray(subjects) ? subjects : [];
  var fallback = Array.isArray(fallbackCounts) ? fallbackCounts : [];
  var counts = [];

  for (var i = 0; i < list.length; i++) {
    var name = String(list[i] || '').replace(/\s+/g, '');
    var mapped = 0;

    if (track === 'police') {
      if (name.indexOf('헌법') >= 0) mapped = 20;
      else if (name.indexOf('형사') >= 0) mapped = 40;
      else if (name.indexOf('경찰학') >= 0 || name.indexOf('경찰') >= 0) mapped = 40;
    } else if (track === 'fire_public') {
      mapped = 25;
    } else if (track === 'fire_paramedic_special') {
      if (name.indexOf('응급') >= 0) mapped = 40;
      else if (name.indexOf('소방학') >= 0) mapped = 25;
    } else if (track === 'fire_rescue_academic_special' || track === 'fire_special') {
      if (name.indexOf('소방학') >= 0) mapped = 25;
      else if (name.indexOf('소방관계법규') >= 0 || name.indexOf('법규') >= 0) mapped = 40;
    }

    if (mapped <= 0) mapped = Number(fallback[i]) || 0;
    counts.push(mapped);
  }

  return counts;
}

function totalItemsFromCounts_(counts, fallbackTotal) {
  var sum = (counts || []).reduce(function(acc, n) {
    return acc + (Number(n) || 0);
  }, 0);
  return sum > 0 ? sum : (Number(fallbackTotal) || 0);
}

function getExamStructure_(studentInfo, examInfo) {
  var studentType = String((studentInfo && studentInfo.type) || (examInfo && examInfo.type) || '').trim();
  var existingSubjects = normalizeSubjectList_(examInfo && examInfo.subjects);

  if (studentType === '경찰') {
    var policeDefaults = ['헌법', '형사법', '경찰학'];
    var policeSubjects = mergeSubjects_(existingSubjects, policeDefaults, 3);
    var policeCounts = resolveItemCountsBySubject_(policeSubjects, 'police', [20, 40, 40]);
    return {
      track: 'police',
      subjects: policeSubjects,
      itemCounts: policeCounts,
      totalItems: totalItemsFromCounts_(policeCounts, 100)
    };
  }

  if (studentType === '소방') {
    var fireTrack = inferFireTrack_(studentInfo, examInfo);
    if (fireTrack === 'fire_public') {
      var firePublicDefaults = ['행정법', '소방학개론', '소방관계법규'];
      var firePublicSubjects = mergeSubjects_(existingSubjects, firePublicDefaults, 3);
      var firePublicCounts = resolveItemCountsBySubject_(firePublicSubjects, fireTrack, [25, 25, 25]);
      return {
        track: fireTrack,
        subjects: firePublicSubjects,
        itemCounts: firePublicCounts,
        totalItems: totalItemsFromCounts_(firePublicCounts, 75)
      };
    }

    if (fireTrack === 'fire_paramedic_special') {
      var fireParamedicDefaults = ['소방학개론', '응급처치학개론'];
      var fireParamedicSubjects = mergeSubjects_(existingSubjects, fireParamedicDefaults, 2);
      var fireParamedicCounts = resolveItemCountsBySubject_(fireParamedicSubjects, fireTrack, [25, 40]);
      return {
        track: fireTrack,
        subjects: fireParamedicSubjects,
        itemCounts: fireParamedicCounts,
        totalItems: totalItemsFromCounts_(fireParamedicCounts, 65)
      };
    }

    if (fireTrack === 'fire_rescue_academic_special' || fireTrack === 'fire_special') {
      var fireSpecialDefaults = ['소방학개론', '소방관계법규'];
      var fireSpecialSubjects = mergeSubjects_(existingSubjects, fireSpecialDefaults, 2);
      var fireSpecialCounts = resolveItemCountsBySubject_(fireSpecialSubjects, fireTrack, [25, 40]);
      return {
        track: fireTrack,
        subjects: fireSpecialSubjects,
        itemCounts: fireSpecialCounts,
        totalItems: totalItemsFromCounts_(fireSpecialCounts, 65)
      };
    }

    // 트랙 불명확 시, 시트 과목 개수 기준으로 추론
    if (existingSubjects.length >= 3) {
      var inferredPublicSubjects = mergeSubjects_(existingSubjects, ['행정법', '소방학개론', '소방관계법규'], 3);
      var inferredPublicCounts = resolveItemCountsBySubject_(inferredPublicSubjects, 'fire_public', [25, 25, 25]);
      return {
        track: fireTrack,
        subjects: inferredPublicSubjects,
        itemCounts: inferredPublicCounts,
        totalItems: totalItemsFromCounts_(inferredPublicCounts, 75)
      };
    }
    if (existingSubjects.length === 2) {
      var inferredSpecialSubjects = mergeSubjects_(existingSubjects, ['소방학개론', '소방관계법규'], 2);
      var inferredTrack = /구급/.test(inferredSpecialSubjects.join(' ')) ? 'fire_paramedic_special' : 'fire_special';
      var inferredSpecialCounts = resolveItemCountsBySubject_(inferredSpecialSubjects, inferredTrack, [25, 40]);
      return {
        track: fireTrack,
        subjects: inferredSpecialSubjects,
        itemCounts: inferredSpecialCounts,
        totalItems: totalItemsFromCounts_(inferredSpecialCounts, 65)
      };
    }
  }

  var fallbackLength = existingSubjects.length > 0 ? existingSubjects.length : 3;
  var fallbackCounts = [];
  for (var f = 0; f < fallbackLength; f++) fallbackCounts.push(0);
  return {
    track: 'unknown',
    subjects: mergeSubjects_(existingSubjects, ['과목1', '과목2', '과목3'], fallbackLength),
    itemCounts: fallbackCounts,
    totalItems: 0
  };
}

function getSubjectByQuestionNo_(num, subjects, itemCounts) {
  var qNum = Number(num) || 0;
  if (qNum <= 0) return '';
  var list = Array.isArray(subjects) ? subjects : [];
  var counts = Array.isArray(itemCounts) ? itemCounts : [];
  if (!list.length || !counts.length || list.length !== counts.length) return '';
  var sumCount = counts.reduce(function(acc, v) { return acc + (Number(v) || 0); }, 0);
  if (sumCount <= 0) return '';

  var end = 0;
  for (var i = 0; i < counts.length; i++) {
    end += Number(counts[i]) || 0;
    if (qNum <= end) return String(list[i] || '').trim();
  }

  return '';
}

function normalizeSubjectKey_(value) {
  return String(value || '').replace(/\s+/g, '').trim();
}

function resolveItemSubject_(itemRow, qNum, subjects, itemCounts) {
  var sheetSubject = String((itemRow && itemRow[ITEM_COL.SUBJECT]) || '').trim();
  if (sheetSubject) {
    var list = Array.isArray(subjects) ? subjects : [];
    var targetKey = normalizeSubjectKey_(sheetSubject);
    if (targetKey && list.length > 0) {
      for (var i = 0; i < list.length; i++) {
        var subjectName = String(list[i] || '').trim();
        if (normalizeSubjectKey_(subjectName) === targetKey) {
          return subjectName;
        }
      }
    }
    return sheetSubject;
  }
  return getSubjectByQuestionNo_(qNum, subjects, itemCounts);
}

function toClientDateValue_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    var tz = Session.getScriptTimeZone() || 'Asia/Seoul';
    return Utilities.formatDate(value, tz, 'yyyy-MM-dd');
  }
  return String(value);
}

// ============================================
// 사용자 기능
// ============================================

function loginUser(studentId, name, password, examType) {
  try {
    if (!studentId || (!name && !password)) {
      return { success: false, message: '수험번호와 이름(또는 비밀번호)을 입력해 주세요.' };
    }

    var targetType = String(examType || '').trim();
    var sheetsToSearch = [];
    if (targetType === '경찰') sheetsToSearch.push(getSheet(CONFIG.SHEET_STUDENTS_POLICE));
    else if (targetType === '소방') sheetsToSearch.push(getSheet(CONFIG.SHEET_STUDENTS_FIRE));
    else {
      sheetsToSearch.push(getSheet(CONFIG.SHEET_STUDENTS_POLICE));
      sheetsToSearch.push(getSheet(CONFIG.SHEET_STUDENTS_FIRE));
    }

    var targetId = String(studentId).trim();
    var targetName = String(name || '').trim();
    var targetPwd = String(password || '').trim();
    var typeMismatch = false;

    for (var sIdx = 0; sIdx < sheetsToSearch.length; sIdx++) {
      var students = getStudentRecords_(sheetsToSearch[sIdx]);
      for (var i = 0; i < students.length; i++) {
        var s = students[i];
        if (s.studentId !== targetId) continue;

        if (targetType && s.type !== targetType) {
          typeMismatch = true;
          continue;
        }

        var nameMatch = targetName && s.name === targetName;
        var pwdMatch = targetPwd && s.password === targetPwd;
        if (!nameMatch && !pwdMatch) continue;

        return {
          success: true,
          data: {
            id: s.studentId,
            studentId: s.studentId,
            name: s.name,
            examType: s.type,
            type: s.type,
            region: s.region,
            phone: s.phone || '',
            birthDate: s.birthDate || '',
            examField: s.examField || '',
            examRank: s.examRank || '',
            examLocation: s.examLocation || '',
            examSubject: s.examSubject || ''
          }
        };
      }
    }

    if (typeMismatch) {
      return { success: false, message: '선택한 직렬과 수험생 정보가 일치하지 않습니다.' };
    }

    return { success: false, message: '일치하는 수험생 정보를 찾을 수 없습니다.' };
  } catch (e) {
    return { success: false, message: '오류가 발생했습니다: ' + e.message };
  }
}

function initializeSampleData(adminToken) {
  var hasSpreadsheetUi = false;
  try {
    SpreadsheetApp.getUi();
    hasSpreadsheetUi = true;
  } catch (uiErr) {}

  if (!hasSpreadsheetUi) {
    requireAdminToken_(adminToken);
  }

  var policeSheet = getSheet(CONFIG.SHEET_STUDENTS_POLICE);
  var fireSheet = getSheet(CONFIG.SHEET_STUDENTS_FIRE);
  
  policeSheet.clearContents();
  fireSheet.clearContents();
  
  initSheet(policeSheet, CONFIG.SHEET_STUDENTS_POLICE);
  initSheet(fireSheet, CONFIG.SHEET_STUDENTS_FIRE);
  
  // 경찰 데이터
  policeSheet.appendRow(['21889', '이동근', '경찰', '대구', '1234', '01011112222', '1997.01.20', '일반', '순경', '대구고사장', '형사법/경찰학/헌법', '샘플']);
  policeSheet.appendRow(['21890', '홍길동', '경찰', '경북', '1111', '01022223333', '1998.08.13', '일반', '순경', '경북고사장', '형사법/경찰학/헌법', '샘플']);
  policeSheet.appendRow(['21893', '박민수', '경찰', '부산', '5555', '01055556666', '1995.12.24', '일반', '순경', '부산고사장', '형사법/경찰학/헌법', '샘플']);

  // 소방 데이터
  fireSheet.appendRow(['21891', '김철수', '소방', '서울', '3333', '01033334444', '1996.02.11', '공채', '소방사', '서울고사장', '소방학/소방법규/행정법', '샘플']);
  fireSheet.appendRow(['21892', '이영희', '소방', '경기', '4444', '01044445555', '1999.04.03', '구급', '소방사', '경기고사장', '소방학/소방법규/행정법', '샘플']);

  var examSheet = getSheet(CONFIG.SHEET_EXAM);
  examSheet.clearContents();
  initSheet(examSheet, CONFIG.SHEET_EXAM);
  examSheet.appendRow(['P2512', '25년 12월 경찰 모의고사', '2025-12-27', '형사법', '경찰학', '헌법']);
  examSheet.appendRow(['F2512', '25년 12월 소방 모의고사', '2025-12-27', '소방학', '소방법규', '행정법']);

  var policeScoreSheet = getSheet(CONFIG.SHEET_SCORE_POLICE);
  var fireScoreSheet = getSheet(CONFIG.SHEET_SCORE_FIRE);
  policeScoreSheet.clearContents();
  fireScoreSheet.clearContents();
  initSheet(policeScoreSheet, CONFIG.SHEET_SCORE_POLICE);
  initSheet(fireScoreSheet, CONFIG.SHEET_SCORE_FIRE);
  policeScoreSheet.appendRow(['P2512', '21889', 77.5, 92.5, 27.5, 197.5, 65.8, 'O', 'X', 'O', 'X', 'O']);
  policeScoreSheet.appendRow(['P2512', '21890', 85.0, 80.0, 40.0, 205.0, 68.3, 'O', 'O', 'X', 'O', 'X']);
  policeScoreSheet.appendRow(['P2512', '21893', 95.0, 90.0, 45.0, 230.0, 76.7, 'O', 'O', 'O', 'X', 'O']);
  fireScoreSheet.appendRow(['F2512', '21891', 90.0, 85.0, 80.0, 255.0, 85.0, 'O', 'X', 'O', 'O', 'X']);
  fireScoreSheet.appendRow(['F2512', '21892', 70.0, 75.0, 65.0, 210.0, 70.0, 'X', 'X', 'O', 'X', 'O']);

  var policeQSheet = getSheet(CONFIG.SHEET_QUESTIONS_POLICE);
  var fireQSheet = getSheet(CONFIG.SHEET_QUESTIONS_FIRE);
  policeQSheet.clearContents();
  fireQSheet.clearContents();
  initSheet(policeQSheet, CONFIG.SHEET_QUESTIONS_POLICE);
  initSheet(fireQSheet, CONFIG.SHEET_QUESTIONS_FIRE);
  policeQSheet.appendRow(['P2512', 1, '형사법', '3', 2.5, 75.5, '상']);
  policeQSheet.appendRow(['P2512', 2, '형사법', '2', 2.5, 62.1, '중']);
  policeQSheet.appendRow(['P2512', 3, '경찰학', '1', 2.5, 48.0, '상']);
  policeQSheet.appendRow(['P2512', 4, '경찰학', '4', 2.5, 35.4, '상']);
  policeQSheet.appendRow(['P2512', 5, '헌법', '2', 2.5, 41.8, '중']);
  fireQSheet.appendRow(['F2512', 1, '소방학', '4', 3, 65.2, '하']);
  fireQSheet.appendRow(['F2512', 2, '소방학', '1', 3, 45.8, '상']);
  fireQSheet.appendRow(['F2512', 3, '소방법규', '2', 3, 39.0, '상']);
  fireQSheet.appendRow(['F2512', 4, '소방법규', '3', 3, 58.3, '중']);
  fireQSheet.appendRow(['F2512', 5, '행정법', '4', 3, 31.6, '상']);

  return '더미 데이터 초기화 완료! 이제 앱을 테스트할 수 있습니다.';
}

function logPrintAction(name, phone) {
  try {
    logAction(name, phone, '수험표 인쇄');
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function logAction(name, phone, action) {
  var sheet = getSheet(CONFIG.SHEET_LOG);
  sheet.appendRow([name || '', phone || '', new Date(), action || '', '']);
}

// ============================================
// 관리자 기능
// ============================================

function verifyAdmin(password) {
  try {
    var ok = String(password || '') === String(getAdminPassword_());
    if (!ok) return { success: false, message: '비밀번호가 올바르지 않습니다.' };
    return { success: true, token: createAdminToken_() };
  } catch (e) {
    return { success: false, message: '관리자 인증 중 오류가 발생했습니다: ' + e.message };
  }
}

function getDashboardStats(adminToken) {
  try {
    requireAdminToken_(adminToken);

    var policeRecords = getStudentRecords_(getSheet(CONFIG.SHEET_STUDENTS_POLICE));
    var fireRecords = getStudentRecords_(getSheet(CONFIG.SHEET_STUDENTS_FIRE));
    var totalRecords = policeRecords.concat(fireRecords);

    var printCount = 0;
    var logData = getSheet(CONFIG.SHEET_LOG).getDataRange().getValues();
    for (var j = 1; j < logData.length; j++) {
      if (String(logData[j][3] || '').indexOf('인쇄') >= 0) {
        printCount++;
      }
    }

    return {
      success: true,
      data: {
        total: totalRecords.length,
        police: policeRecords.length,
        fire: fireRecords.length,
        printCount: printCount
      }
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function registerStudent(data, adminToken) {
  try {
    requireAdminToken_(adminToken);

    if (!data || !data.name || !data.phone || !data.examNumber || !data.examType) {
      return { success: false, message: '필수 항목(성명, 연락처, 응시번호, 시험유형)을 모두 입력해 주세요.' };
    }

    var sheetName = getStudentSheetName_(data.examType);
    if (!sheetName) return { success: false, message: '올바르지 않은 시험 유형입니다.' };

    var sheet = getSheet(sheetName);
    var records = getStudentRecords_(sheet);
    var targetId = String(data.examNumber).trim();
    var exists = records.some(function(s) { return s.studentId === targetId; });
    if (exists) {
      return { success: false, message: '이미 해당 직렬에 등록된 응시번호입니다: ' + targetId };
    }

    sheet.appendRow(buildStudentRow_(data, ''));
    return { success: true, message: data.name + ' 수험생이 등록되었습니다.' };
  } catch (e) {
    return { success: false, message: '등록 중 오류가 발생했습니다: ' + e.message };
  }
}

function bulkRegisterStudents(dataArray, adminToken) {
  try {
    requireAdminToken_(adminToken);

    if (!dataArray || dataArray.length === 0) {
      return { success: false, message: '등록할 데이터가 없습니다.' };
    }

    var policeRows = [];
    var fireRows = [];
    var policeIds = {};
    var fireIds = {};

    // 기존 ID 로드
    getStudentRecords_(getSheet(CONFIG.SHEET_STUDENTS_POLICE)).forEach(function(s){ policeIds[s.studentId] = true; });
    getStudentRecords_(getSheet(CONFIG.SHEET_STUDENTS_FIRE)).forEach(function(s){ fireIds[s.studentId] = true; });

    for (var j = 0; j < dataArray.length; j++) {
      var d = dataArray[j] || {};
      if (!d.name || !d.phone || !d.examNumber || !d.examType) continue;

      var sid = String(d.examNumber).trim();
      if (!sid) continue;

      if (d.examType === '경찰') {
        if (!policeIds[sid]) {
          policeRows.push(buildStudentRow_(d, ''));
          policeIds[sid] = true;
        }
      } else if (d.examType === '소방') {
        if (!fireIds[sid]) {
          fireRows.push(buildStudentRow_(d, ''));
          fireIds[sid] = true;
        }
      }
    }

    var totalRegistered = 0;
    if (policeRows.length > 0) {
      var pSheet = getSheet(CONFIG.SHEET_STUDENTS_POLICE);
      pSheet.getRange(pSheet.getLastRow() + 1, 1, policeRows.length, STUDENT_HEADERS.length).setValues(policeRows);
      totalRegistered += policeRows.length;
    }
    if (fireRows.length > 0) {
      var fSheet = getSheet(CONFIG.SHEET_STUDENTS_FIRE);
      fSheet.getRange(fSheet.getLastRow() + 1, 1, fireRows.length, STUDENT_HEADERS.length).setValues(fireRows);
      totalRegistered += fireRows.length;
    }

    if (totalRegistered === 0) {
      return { success: false, message: '유효한 신규 데이터가 없습니다. 중복/필수값 누락 여부를 확인해 주세요.' };
    }

    return { success: true, message: totalRegistered + '명의 수험생이 등록되었습니다.' };
  } catch (e) {
    return { success: false, message: '일괄 등록 중 오류 발생: ' + e.message };
  }
}

function getStudentList(adminToken) {
  try {
    requireAdminToken_(adminToken);

    var policeRecords = getStudentRecords_(getSheet(CONFIG.SHEET_STUDENTS_POLICE));
    var fireRecords = getStudentRecords_(getSheet(CONFIG.SHEET_STUDENTS_FIRE));

    function mapFn(s, sName) {
      return {
        rowIndex: s.rowIndex,
        sheetName: sName,
        name: s.name,
        phone: s.phone,
        birthDate: s.birthDate,
        examNumber: s.studentId,
        examType: s.type,
        examRegion: s.region,
        examField: s.examField,
        examRank: s.examRank,
        examLocation: s.examLocation,
        examSubject: s.examSubject,
        note: s.note
      };
    }

    var students = policeRecords.map(function(s) { return mapFn(s, CONFIG.SHEET_STUDENTS_POLICE); })
          .concat(fireRecords.map(function(s) { return mapFn(s, CONFIG.SHEET_STUDENTS_FIRE); }));

    return { success: true, data: students };
  } catch (e) {
    return { success: false, message: '목록 조회 중 오류 발생: ' + e.message };
  }
}

function updateStudent(rowIndex, data, adminToken, sheetName) {
  try {
    requireAdminToken_(adminToken);
    if (!rowIndex || !data || !sheetName) return { success: false, message: '수정할 데이터가 올바르지 않습니다.' };

    var resolved = resolveStudentSheetAndRow_(sheetName, rowIndex);
    var sheet = resolved.sheet;
    var safeRowIndex = resolved.rowIndex;

    var existingRow = sheet.getRange(safeRowIndex, 1, 1, STUDENT_HEADERS.length).getValues()[0];
    var existingRecord = toStudentRecord_(existingRow, safeRowIndex);
    var targetId = String(data.examNumber || '').trim();
    if (!targetId) return { success: false, message: '응시번호를 입력해 주세요.' };

    var targetSheetName = getStudentSheetName_(data.examType);
    if (!targetSheetName) return { success: false, message: '올바르지 않은 시험 유형입니다.' };

    // 직렬이 변경된 경우: 새 시트 등록 성공 후 원본 삭제
    if (existingRecord.type !== data.examType) {
      var targetSheet = getSheet(targetSheetName);
      var targetRecords = getStudentRecords_(targetSheet);
      var duplicateInTarget = targetRecords.some(function(s) { return s.studentId === targetId; });
      if (duplicateInTarget) {
        return { success: false, message: '이미 해당 직렬에 등록된 응시번호입니다. ' + targetId };
      }

      targetSheet.appendRow(buildStudentRow_(data, existingRecord.password));
      sheet.deleteRow(safeRowIndex);
      return { success: true, message: '수험생 정보가 수정되었습니다.' };
    }

    var records = getStudentRecords_(sheet);
    var duplicateInSheet = records.some(function(s) {
      return s.rowIndex !== safeRowIndex && s.studentId === targetId;
    });
    if (duplicateInSheet) {
      return { success: false, message: '이미 해당 직렬에 등록된 응시번호입니다. ' + targetId };
    }

    var row = buildStudentRow_(data, existingRecord.password);
    sheet.getRange(safeRowIndex, 1, 1, STUDENT_HEADERS.length).setValues([row]);
    return { success: true, message: '수험생 정보가 수정되었습니다.' };
  } catch (e) {
    return { success: false, message: '수정 중 오류 발생: ' + e.message };
  }
}

function deleteStudent(rowIndex, adminToken, sheetName) {
  try {
    requireAdminToken_(adminToken);
    var resolved = resolveStudentSheetAndRow_(sheetName, rowIndex);
    resolved.sheet.deleteRow(resolved.rowIndex);
    return { success: true, message: '삭제되었습니다.' };
  } catch (e) {
    return { success: false, message: '삭제 중 오류 발생: ' + e.message };
  }
}

function getExamList(studentType) {
  try {
    var examData = getSheet(CONFIG.SHEET_EXAM).getDataRange().getValues();
    var list = [];

    for (var i = 1; i < examData.length; i++) {
      var row = examData[i];
      var examId = String(row[EXAM_COL.ID] || '').trim();
      if (!examId) continue;

      var examType = inferExamType_(examId, row[EXAM_COL.S1]);
      if (studentType && examType && examType !== studentType) continue;

      list.push({
        id: examId,
        name: String(row[EXAM_COL.NAME] || ''),
        date: toClientDateValue_(row[EXAM_COL.DATE]),
        examType: examType
      });
    }

    list.sort(function(a, b) {
      var ad = a.date ? Date.parse(a.date) : 0;
      var bd = b.date ? Date.parse(b.date) : 0;
      ad = isNaN(ad) ? 0 : ad;
      bd = isNaN(bd) ? 0 : bd;
      return bd - ad;
    });

    return { success: true, data: list };
  } catch (e) {
    return { success: false, message: '시험 목록 조회 중 오류 발생: ' + e.message };
  }
}

function getMyExamAccess(studentId, studentType) {
  try {
    var sid = String(studentId || '').trim();
    var sType = String(studentType || '').trim();
    if (!sid) {
      return { success: false, message: '수험번호가 필요합니다.' };
    }

    // 직렬 값이 비어 전달된 경우 수험번호 기준으로 직렬을 보정한다.
    if (!sType) {
      var policeRecords = getStudentRecords_(getSheet(CONFIG.SHEET_STUDENTS_POLICE));
      for (var p = 0; p < policeRecords.length; p++) {
        if (String(policeRecords[p].studentId || '').trim() === sid) {
          sType = String(policeRecords[p].type || '').trim();
          break;
        }
      }
      if (!sType) {
        var fireRecords = getStudentRecords_(getSheet(CONFIG.SHEET_STUDENTS_FIRE));
        for (var f = 0; f < fireRecords.length; f++) {
          if (String(fireRecords[f].studentId || '').trim() === sid) {
            sType = String(fireRecords[f].type || '').trim();
            break;
          }
        }
      }
    }

    var examSheet = getSheet(CONFIG.SHEET_EXAM);
    var examRows = examSheet.getDataRange().getValues();
    var list = [];
    var examMap = {};
    var scoreExamMap = {};

    for (var i = 1; i < examRows.length; i++) {
      var row = examRows[i];
      var examId = String(row[EXAM_COL.ID] || '').trim();
      if (!examId) continue;

      var examType = inferExamType_(examId, row[EXAM_COL.S1]);
      if (sType && examType && examType !== sType) continue;

      var subjects = [
        String(row[EXAM_COL.S1] || '').trim(),
        String(row[EXAM_COL.S2] || '').trim(),
        String(row[EXAM_COL.S3] || '').trim()
      ].filter(function(v) { return !!v; });

      var examInfo = {
        id: examId,
        name: String(row[EXAM_COL.NAME] || examId),
        date: toClientDateValue_(row[EXAM_COL.DATE]),
        examType: examType,
        subjects: subjects,
        hasScore: false,
        canPrint: true,
        canAnalyze: false
      };
      list.push(examInfo);
      examMap[examId] = examInfo;
    }

    var scoreSheets = getScoreSheetsForType_(sType);
    for (var ssIdx = 0; ssIdx < scoreSheets.length; ssIdx++) {
      var scoreSheet = scoreSheets[ssIdx];
      var scoreLastRow = scoreSheet.getLastRow();
      if (scoreLastRow < 2) continue;
      var scoreRows = scoreSheet.getRange(2, 1, scoreLastRow - 1, 2).getValues();
      for (var j = 0; j < scoreRows.length; j++) {
        var rowStudentId = String(scoreRows[j][1] || '').trim();
        if (rowStudentId !== sid) continue;
        var scoreExamId = String(scoreRows[j][0] || '').trim();
        if (!scoreExamId || scoreExamMap[scoreExamId]) continue;
        scoreExamMap[scoreExamId] = true;
      }
    }

    var hasMasterExams = list.length > 0;
    Object.keys(scoreExamMap).forEach(function(examId) {
      if (examMap[examId]) {
        examMap[examId].hasScore = true;
        examMap[examId].canAnalyze = true;
        return;
      }

      // 시험 마스터가 전혀 없을 때만 비상 폴백으로 노출한다.
      if (hasMasterExams) return;

      var fallbackType = inferExamType_(examId, '');
      if (sType && fallbackType && fallbackType !== sType) return;
      list.push({
        id: examId,
        name: examId + ' (마스터 정보 없음)',
        date: '',
        examType: fallbackType,
        subjects: [],
        hasScore: true,
        canPrint: true,
        canAnalyze: true
      });
    });

    list.sort(function(a, b) {
      var ad = a.date ? Date.parse(a.date) : 0;
      var bd = b.date ? Date.parse(b.date) : 0;
      ad = isNaN(ad) ? 0 : ad;
      bd = isNaN(bd) ? 0 : bd;
      return bd - ad;
    });

    return { success: true, data: list };
  } catch (e) {
    return { success: false, message: '시험 접근 정보 조회 중 오류 발생: ' + e.message };
  }
}

function getStudentExamList(studentId, studentType) {
  try {
    var sid = String(studentId || '').trim();
    var sType = String(studentType || '').trim();
    if (!sid) {
      return { success: false, message: '수험번호가 필요합니다.' };
    }

    var myExamMap = {};
    var scoreSheets = getScoreSheetsForType_(sType);
    for (var ssIdx = 0; ssIdx < scoreSheets.length; ssIdx++) {
      var scoreSheet = scoreSheets[ssIdx];
      var scoreLastRow = scoreSheet.getLastRow();
      if (scoreLastRow < 2) continue;

      // 성능 최적화: 시험ID/수험번호 2개 컬럼만 조회
      var scoreRows = scoreSheet.getRange(2, 1, scoreLastRow - 1, 2).getValues();
      for (var i = 0; i < scoreRows.length; i++) {
        var rowStudentId = String(scoreRows[i][1] || '').trim();
        if (rowStudentId !== sid) continue;
        var examId = String(scoreRows[i][0] || '').trim();
        if (examId) myExamMap[examId] = true;
      }
    }

    var myExamIds = Object.keys(myExamMap);
    if (myExamIds.length === 0) {
      return { success: true, data: [] };
    }

    var examSheet = getSheet(CONFIG.SHEET_EXAM);
    var examLastRow = examSheet.getLastRow();
    if (examLastRow < 2) {
      return { success: true, data: [] };
    }

    // 시험마스터는 6개 컬럼만 사용
    var examRows = examSheet.getRange(2, 1, examLastRow - 1, 6).getValues();
    var list = [];
    var foundMap = {};
    for (var j = 0; j < examRows.length; j++) {
      var row = examRows[j];
      var eId = String(row[EXAM_COL.ID] || '').trim();
      if (!eId || !myExamMap[eId]) continue;

      var eType = inferExamType_(eId, row[EXAM_COL.S1]);
      if (sType && eType && eType !== sType) continue;

      list.push({
        id: eId,
        name: String(row[EXAM_COL.NAME] || eId),
        date: toClientDateValue_(row[EXAM_COL.DATE]),
        examType: eType
      });
      foundMap[eId] = true;
    }

    // 성적 시트에는 있는데 Exams 마스터가 비어있는 경우를 대비한 폴백
    for (var k = 0; k < myExamIds.length; k++) {
      var missingId = myExamIds[k];
      if (foundMap[missingId]) continue;
      var missingType = inferExamType_(missingId, '');
      if (sType && missingType && missingType !== sType) continue;
      list.push({
        id: missingId,
        name: missingId + ' (마스터 정보 없음)',
        date: '',
        examType: missingType
      });
    }

    list.sort(function(a, b) {
      var ad = a.date ? Date.parse(a.date) : 0;
      var bd = b.date ? Date.parse(b.date) : 0;
      ad = isNaN(ad) ? 0 : ad;
      bd = isNaN(bd) ? 0 : bd;
      return bd - ad;
    });

    return { success: true, data: list };
  } catch (e) {
    return { success: false, message: '응시 시험 목록 조회 중 오류 발생: ' + e.message };
  }
}

// ============================================
// 성적 분석
// ============================================

function getScoreAnalysis(examId, studentId) {
  try {
    if (!examId || !studentId) {
      return { success: false, message: '시험ID와 수험번호가 필요합니다.' };
    }

    var examSheet = getSheet(CONFIG.SHEET_EXAM);

    var policeRecords = getStudentRecords_(getSheet(CONFIG.SHEET_STUDENTS_POLICE));
    var fireRecords = getStudentRecords_(getSheet(CONFIG.SHEET_STUDENTS_FIRE));
    var students = policeRecords.concat(fireRecords);
    
    var studentMap = {};
    for (var i = 0; i < students.length; i++) {
      studentMap[String(students[i].studentId).trim()] = students[i];
    }

    var student = studentMap[String(studentId).trim()];
    if (!student) return { success: false, message: '학생 정보를 찾을 수 없습니다.' };

    var examRows = examSheet.getDataRange().getValues();
    var exam = null;
    for (var j = 1; j < examRows.length; j++) {
      if (String(examRows[j][EXAM_COL.ID]) === String(examId)) {
        exam = {
          id: String(examRows[j][EXAM_COL.ID]),
          name: String(examRows[j][EXAM_COL.NAME]),
          date: toClientDateValue_(examRows[j][EXAM_COL.DATE]),
          subjects: [
            String(examRows[j][EXAM_COL.S1] || '').trim(),
            String(examRows[j][EXAM_COL.S2] || '').trim(),
            String(examRows[j][EXAM_COL.S3] || '').trim()
          ]
        };
        exam.type = inferExamType_(exam.id, exam.subjects[0]);
        break;
      }
    }
    if (!exam) return { success: false, message: '시험 정보를 찾을 수 없습니다.' };
    if (student.type && exam.type && student.type !== exam.type) {
      return { success: false, message: '선택한 시험은 해당 직렬의 시험이 아닙니다.' };
    }
    var analysisType = exam.type || student.type || '';
    var examStructure = getExamStructure_(student, exam);
    exam.track = examStructure.track;
    exam.subjects = examStructure.subjects;
    exam.subjectItemCounts = examStructure.itemCounts;
    exam.totalItems = examStructure.totalItems;
    var itemRows = collectItemRowsByExam_(exam.id, analysisType);
    var subjectFullScoreMap = {};
    var subjectPointItemCountMap = {};
    exam.subjects.forEach(function(subjName) {
      var key = String(subjName || '').trim();
      if (key) {
        subjectFullScoreMap[key] = 0;
        subjectPointItemCountMap[key] = 0;
      }
    });
    for (var q = 1; q < itemRows.length; q++) {
      if (String(itemRows[q][ITEM_COL.EXAM_ID]) !== String(examId)) continue;
      var qNum = Number(itemRows[q][ITEM_COL.NUM]) || 0;
      if (!qNum) continue;
      var mappedSubject = resolveItemSubject_(itemRows[q], qNum, exam.subjects, exam.subjectItemCounts);
      if (!mappedSubject) continue;
      var point = Number(itemRows[q][ITEM_COL.POINTS]) || 0;
      if (point > 0) {
        if (!subjectFullScoreMap[mappedSubject]) subjectFullScoreMap[mappedSubject] = 0;
        subjectFullScoreMap[mappedSubject] += point;
        subjectPointItemCountMap[mappedSubject] = (subjectPointItemCountMap[mappedSubject] || 0) + 1;
      }
    }
    for (var si = 0; si < exam.subjects.length; si++) {
      var sName = String(exam.subjects[si] || '').trim();
      var expectedItemCount = Number((exam.subjectItemCounts && exam.subjectItemCounts[si]) || 0);
      var actualPointCount = Number(subjectPointItemCountMap[sName] || 0);
      // 문항 배점 입력이 일부만 되어 있으면 full score 기반 계산을 비활성화하고 최고점 기준으로 fallback
      if (expectedItemCount > 0 && actualPointCount > 0 && actualPointCount < expectedItemCount) {
        subjectFullScoreMap[sName] = 0;
      }
    }

    var scoreRows = collectScoreRowsByExam_(exam.id, analysisType);
    var allScores = [];
    var regionalScores = [];
    var myScore = null;

    for (var k = 1; k < scoreRows.length; k++) {
      if (String(scoreRows[k][SCORE_COL.EXAM_ID]) !== String(examId)) continue;

      var sId = String(scoreRows[k][SCORE_COL.STUDENT_ID] || '').trim();
      var scoreObj = {
        studentId: sId,
        s1: Number(scoreRows[k][SCORE_COL.S1]) || 0,
        s2: Number(scoreRows[k][SCORE_COL.S2]) || 0,
        s3: Number(scoreRows[k][SCORE_COL.S3]) || 0,
        total: Number(scoreRows[k][SCORE_COL.TOTAL]) || 0,
        avg: Number(scoreRows[k][SCORE_COL.AVG]) || 0,
        items: scoreRows[k].slice(SCORE_COL.ITEM_START).map(function(v) { return String(v || '').trim().toUpperCase(); })
      };

      allScores.push(scoreObj);
      if (sId === String(studentId)) myScore = scoreObj;

      if (studentMap[sId] && studentMap[sId].region === student.region) {
        regionalScores.push(scoreObj);
      }
    }

    if (!myScore) return { success: false, message: '해당 시험의 성적 데이터가 없습니다.' };
    if (allScores.length === 0) return { success: false, message: '분석 가능한 성적 데이터가 없습니다.' };
    if (regionalScores.length === 0) regionalScores.push(myScore);

    allScores.sort(function(a, b) { return b.total - a.total; });
    regionalScores.sort(function(a, b) { return b.total - a.total; });

    // 전체 석차 리스트 (동점자 동일 석차) 가공
    var overallRanking = [];

    var totalCount = allScores.length;
    var regionalCount = regionalScores.length;

    function rankByValueForStudent_(rows, key, sid) {
      var targetVal = null;
      for (var idx = 0; idx < rows.length; idx++) {
        if (String(rows[idx].studentId) === String(sid)) {
          targetVal = Number(rows[idx][key]) || 0;
          break;
        }
      }
      if (targetVal === null) return 0;
      var higherCount = rows.filter(function(row) {
        return (Number(row[key]) || 0) > targetVal;
      }).length;
      return higherCount + 1;
    }

    function buildRankMapByKey_(rows, key) {
      var rankMap = {};
      var prevVal = null;
      var currentRank = 0;
      for (var rIdx = 0; rIdx < rows.length; rIdx++) {
        var row = rows[rIdx];
        var val = Number(row[key]) || 0;
        if (prevVal === null || val < prevVal) {
          currentRank = rIdx + 1;
          prevVal = val;
        }
        rankMap[String(row.studentId)] = currentRank;
      }
      return rankMap;
    }

    var overallTotalRankMap = buildRankMapByKey_(allScores, 'total');
    var regionalTotalRankMap = buildRankMapByKey_(regionalScores, 'total');
    var overallRank = Number(overallTotalRankMap[String(studentId)]) || rankByValueForStudent_(allScores, 'total', studentId);
    var regionalRank = Number(regionalTotalRankMap[String(studentId)]) || rankByValueForStudent_(regionalScores, 'total', studentId);

    function toTopPercent_(rank, count) {
      if (!count || count <= 0 || !rank || rank <= 0) return 100;
      var top = ((Number(rank) - 1) / Number(count)) * 100;
      if (!isFinite(top)) return 100;
      if (top < 0) top = 0;
      if (top > 100) top = 100;
      return Number(top.toFixed(1));
    }

    overallRanking = allScores.map(function(s) {
      var sInfo = studentMap[s.studentId] || {};
      return {
        rank: Number(overallTotalRankMap[String(s.studentId)]) || 0,
        isMe: String(s.studentId) === String(studentId),
        studentIdMasked: maskStudentId_(s.studentId),
        region: sInfo.region || '-',
        s1: s.s1,
        s2: s.s2,
        s3: s.s3,
        total: s.total,
        avg: s.avg
      };
    });

    // 동점 보정 percentile rank (0~100, 값이 클수록 유리)
    function toPercentileRank_(rows, key, sid) {
      if (!rows || rows.length === 0) return 0;
      var targetVal = null;
      for (var t = 0; t < rows.length; t++) {
        if (String(rows[t].studentId) === String(sid)) {
          targetVal = Number(rows[t][key]) || 0;
          break;
        }
      }
      if (targetVal === null) return 0;

      var lowerCount = 0;
      var equalCount = 0;
      for (var p = 0; p < rows.length; p++) {
        var v = Number(rows[p][key]) || 0;
        if (v < targetVal) lowerCount++;
        else if (v === targetVal) equalCount++;
      }
      return Number((((lowerCount + 0.5 * equalCount) / rows.length) * 100).toFixed(1));
    }

    // topPercent: 상위 누적 비율(작을수록 상위), percentile: 동점 보정 백분위(클수록 상위)
    var percentile = toPercentileRank_(allScores, 'total', studentId);
    var regionalPercentile = toPercentileRank_(regionalScores, 'total', studentId);
    var topPercent = toTopPercent_(overallRank, totalCount);
    var regionalTopPercent = toTopPercent_(regionalRank, regionalCount);

    var totalSum = allScores.reduce(function(acc, s) { return acc + s.total; }, 0);
    var overallAvg = Number((totalSum / totalCount).toFixed(1));
    var maxTotal = Number(allScores[0].total || 0);
    var minTotal = Number(allScores[allScores.length - 1].total || 0);

    var top10Count = Math.max(1, Math.ceil(totalCount * 0.1));
    var top30Count = Math.max(1, Math.ceil(totalCount * 0.3));
    var top10Scores = allScores.slice(0, top10Count);
    var top30Scores = allScores.slice(0, top30Count);

    function avgByKey_(rows, key) {
      if (!rows || rows.length === 0) return 0;
      var sum = rows.reduce(function(acc, row) { return acc + (Number(row[key]) || 0); }, 0);
      return Number((sum / rows.length).toFixed(1));
    }

    function rankByKey_(rows, key, sid) {
      return rankByValueForStudent_(rows, key, sid);
    }

    var subjectStats = exam.subjects.map(function(name, idx) {
      var key = 's' + (idx + 1);
      var values = allScores.map(function(s) { return Number(s[key]) || 0; });
      var subjectMax = Math.max.apply(null, values);
      var subjectFullMark = Number(subjectFullScoreMap[name]) || 0;
      var avg = Number((values.reduce(function(a, b) { return a + b; }, 0) / values.length).toFixed(1));
      var top10Avg = avgByKey_(top10Scores, key);
      var top30Avg = avgByKey_(top30Scores, key);
      var my = Number(myScore[key]) || 0;
      var myRank = rankByKey_(allScores, key, studentId);
      var subjectPercentile = toPercentileRank_(allScores, key, studentId);
      var subjectTopPercent = toTopPercent_(myRank, totalCount);
      var scoreRateBase = subjectFullMark > 0 ? subjectFullMark : subjectMax;
      var scoreRate = scoreRateBase > 0 ? Number((my / scoreRateBase * 100).toFixed(1)) : 0;
      // 과목 평가: 소수 집단에서도 과도한 왜곡을 줄이기 위해
      // 표본이 적을 때는 과목별 성취율(만점 대비), 표본이 충분할 때는 과목 내 상위 백분위 기준으로 판정
      var grade = '취약';
      var gradeBasis = 'percentile';
      if (totalCount < 10) {
        gradeBasis = 'scoreRate';
        if (scoreRate >= 80) grade = '우수';
        else if (scoreRate >= 60) grade = '보통';
      } else {
        if (subjectTopPercent <= 30) grade = '우수';
        else if (subjectTopPercent <= 60) grade = '보통';
      }

      return {
        name: name,
        my: my,
        avg: avg,
        max: subjectMax,
        fullMark: scoreRateBase,
        myRank: myRank,
        percentile: subjectPercentile,
        topPercent: subjectTopPercent,
        scoreRate: scoreRate,
        gradeBasis: gradeBasis,
        totalCount: totalCount,
        top10Avg: top10Avg,
        top30Avg: top30Avg,
        grade: grade
      };
    });

    var externalAnswerMap = getExternalStudentAnswers_(examId, studentId);
    var itemAnalysis = [];
    for (var m = 1; m < itemRows.length; m++) {
      if (String(itemRows[m][ITEM_COL.EXAM_ID]) !== String(examId)) continue;

      var num = Number(itemRows[m][ITEM_COL.NUM]);
      if (!num) continue;

      var answerText = String(itemRows[m][ITEM_COL.ANSWER] || '').trim();
      var externalRaw = externalAnswerMap.hasOwnProperty(num) ? externalAnswerMap[num] : '';
      var rawValue = String(externalRaw || myScore.items[num - 1] || '').trim();
      var upperRaw = rawValue.toUpperCase();
      var mark = '-';
      var studentAnswer = '-';
      if (upperRaw === 'O' || upperRaw === 'X') {
        mark = upperRaw;
        studentAnswer = upperRaw === 'O' ? answerText : '-';
      } else if (rawValue) {
        studentAnswer = rawValue;
        var normalizedAnswer = answerText.toUpperCase();
        if (normalizedAnswer) {
          mark = upperRaw === normalizedAnswer ? 'O' : 'X';
        }
      }

      itemAnalysis.push({
        num: num,
        subject: resolveItemSubject_(itemRows[m], num, exam.subjects, exam.subjectItemCounts),
        answer: answerText || '-',
        points: itemRows[m][ITEM_COL.POINTS],
        correctRate: Number(itemRows[m][ITEM_COL.CORRECT_RATE]) || 0,
        difficulty: itemRows[m][ITEM_COL.DIFFICULTY] || '-',
        my: mark,
        studentAnswer: studentAnswer
      });
    }

    if (itemAnalysis.length === 0 && myScore.items.length > 0) {
      for (var n = 0; n < myScore.items.length; n++) {
        var externalItem = externalAnswerMap.hasOwnProperty(n + 1) ? externalAnswerMap[n + 1] : '';
        var rawItem = String(externalItem || myScore.items[n] || '').trim();
        var upperItem = rawItem.toUpperCase();
        var fallbackMark = (upperItem === 'O' || upperItem === 'X') ? upperItem : '-';
        var fallbackStudentAnswer = rawItem || '-';
        var correctCount = allScores.filter(function(s) { return s.items[n] === 'O'; }).length;
        itemAnalysis.push({
          num: n + 1,
          subject: resolveItemSubject_(null, n + 1, exam.subjects, exam.subjectItemCounts),
          answer: '-',
          points: '-',
          correctRate: Number((correctCount / totalCount * 100).toFixed(1)),
          difficulty: '-',
          my: fallbackMark,
          studentAnswer: fallbackStudentAnswer
        });
      }
    }

    var killerQuestions = itemAnalysis
      .slice()
      .sort(function(a, b) { return a.correctRate - b.correctRate; })
      .slice(0, 5);

    function getDistributionMaxScore_(studentInfo, examInfo, observedMaxScore, myTotal) {
      var studentType = String((studentInfo && studentInfo.type) || '').trim();
      var sourceText = [
        (studentInfo && studentInfo.examField) || '',
        (studentInfo && studentInfo.examRank) || '',
        (studentInfo && studentInfo.examSubject) || '',
        (examInfo && examInfo.name) || ''
      ].join(' ');

      if (studentType === '경찰') return 250;
      if (studentType === '소방') {
        if (/경채|구급|구조|학과/.test(sourceText)) return 200;
        return 300;
      }
      var fallbackObserved = Number(observedMaxScore) || 0;
      return fallbackObserved > 260 ? 300 : 250;
    }

    var normalizedMax = getDistributionMaxScore_(student, exam, maxTotal, myScore.total);
    if (normalizedMax <= 0) normalizedMax = 250;
    var distributionBins = [];
    var decileCount = 10;
    for (var d = 0; d < decileCount; d++) {
      var fromPct = d * 10;
      var toPct = (d + 1) * 10;
      distributionBins.push({
        decile: d + 1,
        shortLabel: (d + 1) + '등급',
        bandLabel: '상위 ' + fromPct + '~' + toPct + '%',
        threshold: 0,
        upper: 0,
        label: '-',
        count: 0,
        ratio: 0
      });
    }

    for (var ds = 0; ds < allScores.length; ds++) {
      var scoreValue = Number(allScores[ds].total) || 0;
      if (scoreValue < 0) scoreValue = 0;
      if (scoreValue > normalizedMax) scoreValue = normalizedMax;
      var binIdx = Math.floor(ds * decileCount / totalCount);
      if (binIdx < 0) binIdx = 0;
      if (binIdx >= distributionBins.length) binIdx = distributionBins.length - 1;
      var targetBin = distributionBins[binIdx];
      targetBin.count++;
      if (targetBin.count === 1) {
        targetBin.upper = scoreValue;
        targetBin.threshold = scoreValue;
      } else {
        if (scoreValue > targetBin.upper) targetBin.upper = scoreValue;
        if (scoreValue < targetBin.threshold) targetBin.threshold = scoreValue;
      }
    }

    var safeTotalCount = Math.max(1, totalCount);
    var safeOverallRank = Math.max(1, overallRank);
    var myDecileIdx = Math.min(decileCount - 1, Math.floor((safeOverallRank - 1) * decileCount / safeTotalCount));

    for (var db = 0; db < distributionBins.length; db++) {
      var bin = distributionBins[db];
      if (bin.count > 0) {
        var low = Number(bin.threshold) || 0;
        var high = Number(bin.upper) || 0;
        if (low > high) {
          var temp = low;
          low = high;
          high = temp;
        }
        bin.threshold = low;
        bin.upper = high;
        bin.label = (low === high) ? (high + '점') : (low + '~' + high + '점');
      } else {
        bin.threshold = 0;
        bin.upper = 0;
        bin.label = '데이터 없음';
      }
      bin.ratio = Number((bin.count / totalCount * 100).toFixed(1));
    }

    var regionalTotalSum = regionalScores.reduce(function(acc, s) { return acc + (Number(s.total) || 0); }, 0);
    var regionalAvgTotal = Number((regionalTotalSum / regionalCount).toFixed(1));
    var regionalTopCount = Math.max(1, Math.ceil(regionalCount * 0.1));
    var regionalTopAvg = Number((regionalScores.slice(0, regionalTopCount).reduce(function(acc, s) {
      return acc + (Number(s.total) || 0);
    }, 0) / regionalTopCount).toFixed(1));

    var itemCorrectCount = itemAnalysis.filter(function(q) { return q.my === 'O'; }).length;
    var itemWrongCount = itemAnalysis.filter(function(q) { return q.my === 'X'; }).length;
    var itemUnknownCount = Math.max(0, itemAnalysis.length - itemCorrectCount - itemWrongCount);
    var itemCorrectRate = itemAnalysis.length > 0 ? Number((itemCorrectCount / itemAnalysis.length * 100).toFixed(1)) : 0;

    // 킬러 문항 (정답률 40% 미만) 정복률 분석
    var killerItems = itemAnalysis.filter(function(q) { return q.correctRate < 40; });
    var killerCorrect = killerItems.filter(function(q) { return q.my === 'O'; }).length;
    var killerConquerRate = killerItems.length > 0 ? Number((killerCorrect / killerItems.length * 100).toFixed(1)) : 0;

    // 나만 틀린 문제 (정답률 80% 이상인데 틀린 경우)
    var easyMissedQuestions = itemAnalysis.filter(function(q) {
      return q.correctRate >= 80 && q.my === 'X';
    }).sort(function(a, b) { 
      return b.correctRate - a.correctRate; 
    }).slice(0, 5);



    // 과목 밸런스 분석 (과목 간 표준편차)
    // 과목별 만점 차이를 보정하기 위해 과목별 최고점 대비 비율(%)로 균형도 계산
    var myScores = subjectStats.map(function(s) {
      var subjectMax = Number(s.max) || 0;
      var subjectMy = Number(s.my) || 0;
      if (subjectMax <= 0) return 0;
      return subjectMy / subjectMax * 100;
    });
    var avgScore = myScores.length > 0
      ? (myScores.reduce(function(a, b) { return a + b; }, 0) / myScores.length)
      : 0;
    var variance = myScores.length > 0
      ? (myScores.reduce(function(sum, score) {
          return sum + Math.pow(score - avgScore, 2);
        }, 0) / myScores.length)
      : 0;
    var stdDev = Math.sqrt(variance);
    var balanceAssessment = stdDev < 7 ? '매우 균형' : (stdDev < 12 ? '균형' : (stdDev < 18 ? '보통' : '불균형'));

    // 경고 및 조언 메시지 생성
    var warnings = [];
    subjectStats.forEach(function(subj) {
      if (subj.grade === '취약') {
        warnings.push(subj.name + ' 과목이 상대적으로 취약합니다. 집중 학습이 필요합니다.');
      }
      var maxPoint = Number(subj.fullMark || subj.max) || 0;
      var relativeGapThreshold = maxPoint > 0 ? (maxPoint * 0.05) : 5;
      if (subj.my < subj.avg - relativeGapThreshold) {
        warnings.push(subj.name + ' 점수가 전체 평균보다 낮습니다.');
      }
    });
    if (balanceAssessment === '불균형') {
      warnings.push('과목 간 점수 편차가 매우 큽니다. 취약 과목 보완이 시급합니다.');
    }

    var top10TotalAvg = Number((top10Scores.reduce(function(acc, s) { return acc + (Number(s.total) || 0); }, 0) / top10Count).toFixed(1));
    var top30TotalAvg = Number((top30Scores.reduce(function(acc, s) { return acc + (Number(s.total) || 0); }, 0) / top30Count).toFixed(1));
    function maskStudentId_(sid) {
      var raw = String(sid || '');
      if (raw.length <= 2) return raw;
      return raw.slice(0, 2) + new Array(raw.length - 1).join('*');
    }

    function inferZoneByTopPercent_(p, sampleCount) {
      if ((Number(sampleCount) || 0) < 30) return '참고치';
      var v = Number(p);
      if (!isFinite(v) || v < 0) v = 100;
      if (v <= 5) return '합격확실권';
      if (v <= 15) return '합격유력권';
      if (v <= 30) return '합격가능권';
      return '합격도전권';
    }

    var competitorStart = Math.max(0, overallRank - 6);
    var competitorEnd = Math.min(allScores.length, competitorStart + 10);
    var competitors = [];
    for (var cp = competitorStart; cp < competitorEnd; cp++) {
      var cpScore = allScores[cp];
      var cpItems = Array.isArray(cpScore.items) ? cpScore.items : [];
      var cpCorrect = cpItems.filter(function(v) { return String(v).toUpperCase() === 'O'; }).length;
      var cpWrong = cpItems.filter(function(v) { return String(v).toUpperCase() === 'X'; }).length;
      var cpChoice = cpItems.filter(function(v) { return /^\d+$/.test(String(v || '').trim()); }).length;
      var cpRank = Number(overallTotalRankMap[String(cpScore.studentId)]) || rankByValueForStudent_(allScores, 'total', cpScore.studentId);
      var cpTopPercent = toTopPercent_(cpRank, totalCount);
      competitors.push({
        rank: cpRank,
        topPercent: cpTopPercent,
        studentIdMasked: maskStudentId_(cpScore.studentId),
        total: Number(cpScore.total) || 0,
        gapFromMe: Number(((Number(myScore.total) || 0) - (Number(cpScore.total) || 0)).toFixed(1)),
        zone: inferZoneByTopPercent_(cpTopPercent, totalCount),
        items: cpItems,
        responseSummary: {
          correct: cpCorrect,
          wrong: cpWrong,
          choice: cpChoice
        }
      });
    }

    return {
      success: true,
      data: {
        student: {
          id: student.studentId,
          name: student.name,
          type: student.type,
          region: student.region,
          examField: student.examField,
          examRank: student.examRank,
          examLocation: student.examLocation,
          examSubject: student.examSubject
        },
        exam: exam,
        myScore: myScore,
        ranks: {
          overall: overallRank,
          totalCount: totalCount,
          regional: regionalRank > 0 ? regionalRank : 1,
          regionalCount: regionalCount,
          percentile: percentile,
          topPercent: topPercent,
          regionalPercentile: regionalPercentile,
          regionalTopPercent: regionalTopPercent
        },
        stats: {
          overallAvg: overallAvg,
          subjects: subjectStats,
          top10TotalAvg: top10TotalAvg,
          top30TotalAvg: top30TotalAvg,
          maxTotal: maxTotal,
          minTotal: minTotal,
          predictionReliability: {
            isReliable: totalCount >= 30,
            minRecommendedCount: 30,
            totalCount: totalCount
          },
          balanceScore: {
            stdDev: Number(stdDev.toFixed(1)),
            assessment: balanceAssessment
          }
        },
        regional: {
          regionName: student.region,
          avgTotal: regionalAvgTotal,
          top10AvgTotal: regionalTopAvg,
          myTotal: Number(myScore.total) || 0,
          rank: regionalRank > 0 ? regionalRank : 1,
          count: regionalCount,
          percentile: regionalPercentile,
          topPercent: regionalTopPercent
        },
        distribution: {
          totalMax: normalizedMax,
          mode: 'decile',
          binCount: decileCount,
          myBinIndex: myDecileIdx,
          bins: distributionBins
        },
        itemStats: {
          totalItems: itemAnalysis.length,
          correctCount: itemCorrectCount,
          wrongCount: itemWrongCount,
          unknownCount: itemUnknownCount,
          myCorrectRate: itemCorrectRate,
          killerTotal: killerItems.length,
          killerCorrect: killerCorrect,
          killerConquerRate: killerConquerRate
        },
        competitors: competitors,

        killerQuestions: killerQuestions,
        easyMissedQuestions: easyMissedQuestions,
        warnings: warnings,
        overallRanking: overallRanking,
        itemAnalysis: itemAnalysis,
        resultStatus: 'UNKNOWN'
      }
    };
  } catch (e) {
    return { success: false, message: '분석 중 오류 발생: ' + e.message };
  }
}

function getStudentReport(examId, studentId) {
  return getScoreAnalysis(examId, studentId);
}

function getExternalStudentAnswers_(examId, studentId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var candidateNames = CONFIG.SHEET_RESPONSE_CANDIDATES || [];
    var answerSheet = null;
    for (var i = 0; i < candidateNames.length; i++) {
      var byName = ss.getSheetByName(candidateNames[i]);
      if (byName) {
        answerSheet = byName;
        break;
      }
    }
    if (!answerSheet || answerSheet.getLastRow() < 2) return {};

    var rows = answerSheet.getDataRange().getValues();
    var headers = rows[0] || [];

    function norm_(v) {
      return String(v || '').replace(/\s+/g, '').toLowerCase();
    }
    function findIdx_(candidates) {
      for (var c = 0; c < headers.length; c++) {
        var hv = norm_(headers[c]);
        for (var p = 0; p < candidates.length; p++) {
          if (hv === candidates[p] || hv.indexOf(candidates[p]) >= 0) return c;
        }
      }
      return -1;
    }

    var examIdx = findIdx_(['시험id', 'examid', 'exam_id']);
    var studentIdx = findIdx_(['수험번호', '응시번호', 'studentid', 'student_id']);
    if (examIdx < 0 || studentIdx < 0) return {};

    var itemIdx = findIdx_(['문항번호', '문항', 'itemno', 'item_num', 'itemnum']);
    var answerIdx = findIdx_(['학생답안', '응답', '선택지', '답안', 'answer', 'choice', 'response']);
    var examKey = String(examId || '').trim();
    var studentKey = String(studentId || '').trim();
    var answerMap = {};

    for (var r = 1; r < rows.length; r++) {
      var row = rows[r];
      if (String(row[examIdx] || '').trim() !== examKey) continue;
      if (String(row[studentIdx] || '').trim() !== studentKey) continue;

      if (itemIdx >= 0 && answerIdx >= 0) {
        var num = Number(row[itemIdx]) || 0;
        var ans = String(row[answerIdx] || '').trim();
        if (num > 0 && ans) answerMap[num] = ans;
        continue;
      }

      for (var col = 0; col < headers.length; col++) {
        var key = norm_(headers[col]);
        var matched = key.match(/^(?:q|문항)?(\d+)$/);
        if (!matched) continue;
        var qNum = Number(matched[1]) || 0;
        if (!qNum) continue;
        var qAnswer = String(row[col] || '').trim();
        if (qAnswer) answerMap[qNum] = qAnswer;
      }
    }

    return answerMap;
  } catch (e) {
    return {};
  }
}

// ===== 추가 유틸리티 함수 =====

/**
 * 수험번호 마스킹 (예: 12345 -> 12***)
 */
function maskStudentId_(id) {
  var str = String(id || '').trim();
  if (str.length <= 2) return str;
  return str.substring(0, 2) + '*'.repeat(str.length - 2);
}

/**
 * 이름 마스킹 (예: 홍길동 -> 홍*동, 김철수 -> 김*수)
 */
function maskName_(name) {
  var str = String(name || '').trim();
  if (str.length <= 1) return str;
  if (str.length === 2) return str[0] + '*';
  return str[0] + '*'.repeat(str.length - 2) + str[str.length - 1];
}
