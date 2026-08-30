function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) ? e.parameter.page.toLowerCase() : '';
  const template = HtmlService.createTemplateFromFile('index');
  template.pageMode = (page === 'student') ? 'student' : 'teacher';
  return template.evaluate()
    .setTitle(page === 'student' ? '한영 번역 마스터 (학생 제출)' : '한영 번역 마스터 (교사 대시보드)')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

const SHEET_NAMES = {
  SETTINGS: '설정',
  LYRICS: '원본정보',
  SUBMISSIONS: '제출현황',
  GRADE_BEST: '학년최종결과',
  AI_FEEDBACK: 'AI피드백',
  PEER_EVAL: '동료평가'
};

function findClassSheet(ss, className) {
  if (!className) return null;
  const strName = String(className).trim();
  const options = [
    `[학급] ${strName}`,
    `[학급]${strName}`,
    strName,
    `${strName}반`
  ];
  for (let opt of options) {
    let sheet = ss.getSheetByName(opt);
    if (sheet) return sheet;
  }
  return null;
}

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. 설정
  let settingsSheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet(SHEET_NAMES.SETTINGS);
  }
  settingsSheet.clearContents();
  settingsSheet.getRange(1, 1, 10, 2).setValues([
    ['설정명', '값'],
    ['반이름 나열 (쉼표구분)', '1반, 2반, 3반'],
    ['반별 모둠 개수', 8],
    ['모둠원 역할 (쉼표구분)', 'A,B,C,D'],
    ['ChatGPT API 키', ''],
    ['Claude API 키', ''],
    ['동료 평가 참여 활성화', 'Y'],
    ['동료 평가 결과 확인 활성화', 'Y'],
    ['현재 활성 진도', '1차시'],
    ['진도 목록 (쉼표구분)', '1차시, 2차시, 3차시']
  ]);
  settingsSheet.getRange("A1:B1").setFontWeight("bold");
  settingsSheet.setColumnWidth(1, 180);
  settingsSheet.setColumnWidth(2, 250);
  
  // 2. 원본정보 ('원본가사'에서 변경됨)
  let lyricsSheet = ss.getSheetByName(SHEET_NAMES.LYRICS) || ss.getSheetByName('원본정보') || ss.getSheetByName('원본가사');
  if (!lyricsSheet) {
    lyricsSheet = ss.insertSheet(SHEET_NAMES.LYRICS);
    lyricsSheet.appendRow(['진도ID', '순번', '담당 역할', '우리말 문장', '영어 문장']);
    lyricsSheet.getRange("A1:E1").setFontWeight("bold");
    
    // Example data
    const sampleLyrics = [
      ['1차시', 1, 'A', '나는 그저 가난한 소년일 뿐입니다.', 'I am just a poor boy'],
      ['1차시', 2, 'B', '내 이야기는 거의 전해지지 않지만,', 'Though my story is seldom told'],
      ['1차시', 3, 'C', '나는 내 저항력을 탕진했습니다.', 'I have squandered my resistance'],
      ['1차시', 4, 'D', '중얼거림 한 주머니를 위해.', 'For a pocketful of mumbles'],
      ['1차시', 5, 'A', '그런 것들이 약속입니다.', 'Such are promises'],
      ['1차시', 6, 'B', '모두 거짓과 농담일 뿐,', 'All lies and jest'],
      ['1차시', 7, 'C', '여전히 사람은 들으려는 것만 듣고,', 'Still a man hears what he wants to hear'],
      ['1차시', 8, 'D', '나머지는 무시합니다.', 'And disregards the rest'],
    ];
    sampleLyrics.forEach(row => lyricsSheet.appendRow(row));
  } else {
    if (lyricsSheet.getName() !== SHEET_NAMES.LYRICS) {
      lyricsSheet.setName(SHEET_NAMES.LYRICS);
    }
    const firstHeader = String(lyricsSheet.getRange(1, 1).getValue()).trim();
    if (firstHeader !== '진도ID' && firstHeader !== '차시' && firstHeader !== 'Lesson') {
      lyricsSheet.insertColumnBefore(1);
      lyricsSheet.getRange(1, 1).setValue('진도ID');
      const lastRow = lyricsSheet.getLastRow();
      if (lastRow > 1) {
        const fillValues = Array(lastRow - 1).fill(['1차시']);
        lyricsSheet.getRange(2, 1, lastRow - 1, 1).setValues(fillValues);
      }
    }
    lyricsSheet.getRange(1, 1, 1, 5).setValues([['진도ID', '순번', '담당 역할', '우리말 문장', '영어 문장']]);
    lyricsSheet.getRange("A1:E1").setFontWeight("bold");
  }
  
  // 3. 제출현황
  let submissionsSheet = ss.getSheetByName(SHEET_NAMES.SUBMISSIONS);
  if (!submissionsSheet) {
    submissionsSheet = ss.insertSheet(SHEET_NAMES.SUBMISSIONS);
    submissionsSheet.appendRow(['타임스탬프', '반', '모둠', '역할', '순번', '영어 문장', '진도ID']);
    submissionsSheet.getRange("A1:G1").setFontWeight("bold");
  } else {
    submissionsSheet.getRange(1, 1, 1, 7).setValues([['타임스탬프', '반', '모둠', '역할', '순번', '영어 문장', '진도ID']]);
    submissionsSheet.getRange("A1:G1").setFontWeight("bold");
  }
  
  // 4. 각 반 시트 ([학급] 반이름)
  const currentClasses = getSettings().classes;
  currentClasses.forEach(className => {
    let classSheet = findClassSheet(ss, className);
    const targetSheetName = `[학급] ${className}`;
    if (!classSheet) {
      classSheet = ss.insertSheet(targetSheetName);
      classSheet.appendRow(['순번', '영어 문장', '베스트 정답', '채택모둠', '진도ID']);
      classSheet.getRange("A1:E1").setFontWeight("bold");
    } else {
      if (classSheet.getName() !== targetSheetName) {
        classSheet.setName(targetSheetName);
      }
      classSheet.getRange(1, 1, 1, 5).setValues([['순번', '영어 문장', '베스트 정답', '채택모둠', '진도ID']]);
      classSheet.getRange("A1:E1").setFontWeight("bold");
    }
  });
  
  // 5. 학년최종결과
  let gradeSheet = ss.getSheetByName(SHEET_NAMES.GRADE_BEST);
  if (!gradeSheet) {
    gradeSheet = ss.insertSheet(SHEET_NAMES.GRADE_BEST);
    gradeSheet.appendRow(['순번', '우리말 문장', '학년 베스트 영어 문장', '출처', '진도ID']);
    gradeSheet.getRange("A1:E1").setFontWeight("bold");
  } else {
    gradeSheet.getRange(1, 1, 1, 5).setValues([['순번', '우리말 문장', '학년 베스트 영어 문장', '출처', '진도ID']]);
    gradeSheet.getRange("A1:E1").setFontWeight("bold");
  }

  // 6. AI피드백
  let aiSheet = ss.getSheetByName(SHEET_NAMES.AI_FEEDBACK);
  if (!aiSheet) {
    aiSheet = ss.insertSheet(SHEET_NAMES.AI_FEEDBACK);
    aiSheet.appendRow(['타임스탬프', '반', '모둠', '역할', '순번', '우리말 원문', '학생 영문 번역', '문법 피드백', '어휘 피드백', '표현 피드백', 'Paraphrase 피드백', '제공 AI', '진도ID']);
    aiSheet.getRange("A1:M1").setFontWeight("bold");
  } else {
    aiSheet.getRange(1, 1, 1, 13).setValues([['타임스탬프', '반', '모둠', '역할', '순번', '우리말 원문', '학생 영문 번역', '문법 피드백', '어휘 피드백', '표현 피드백', 'Paraphrase 피드백', '제공 AI', '진도ID']]);
    aiSheet.getRange("A1:M1").setFontWeight("bold");
  }

  // 7. 동료평가
  let peerSheet = ss.getSheetByName(SHEET_NAMES.PEER_EVAL);
  if (!peerSheet) {
    peerSheet = ss.insertSheet(SHEET_NAMES.PEER_EVAL);
    peerSheet.appendRow(['타임스탬프', '평가자반', '평가자모둠', '평가자역할', '피평가자모둠', '피평가자역할', '피평가자순번', '우리말문장', '피평가자영어문장', '어법점수', '어휘점수', '표현점수', '총점', '객관식설명', '평가유형', '진도ID']);
    peerSheet.getRange("A1:P1").setFontWeight("bold");
  } else {
    peerSheet.getRange(1, 1, 1, 16).setValues([['타임스탬프', '평가자반', '평가자모둠', '평가자역할', '피평가자모둠', '피평가자역할', '피평가자순번', '우리말문장', '피평가자영어문장', '어법점수', '어휘점수', '표현점수', '총점', '객관식설명', '평가유형', '진도ID']]);
    peerSheet.getRange("A1:P1").setFontWeight("bold");
  }
}

function getSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  if (!sheet) {
    setupSheets();
    sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  }
  
  const data = sheet.getDataRange().getValues();
  let settings = {
    classes: ['1반', '2반', '3반'],
    numGroups: 8,
    roles: ['A', 'B', 'C', 'D'],
    chatGptKey: '',
    claudeKey: '',
    enablePeerEvalPart: true,
    enablePeerEvalResult: true,
    activeLesson: '1차시',
    lessonList: ['1차시', '2차시', '3차시']
  };
  
  let foundClasses = false;
  let legacyNumClasses = null;
  
  for (let i = 1; i < data.length; i++) {
    const key = data[i][0] ? data[i][0].toString().trim() : '';
    const val = data[i][1] ? data[i][1].toString().trim() : '';
    
    if (key === '반이름 나열 (쉼표구분)' || key === '반 이름 나열 (쉼표구분)' || key === '반이름 나열') {
      const parsed = val.split(',').map(s => s.trim()).filter(s => s);
      if (parsed.length > 0) {
        settings.classes = parsed;
        foundClasses = true;
      }
    } else if (key === '반 개수') {
      legacyNumClasses = parseInt(val);
    } else if (key === '반별 모둠 개수') {
      settings.numGroups = parseInt(val) || 8;
    } else if (key === '모둠원 역할 (쉼표구분)') {
      const parsedRoles = val.split(',').map(s => s.trim()).filter(s => s);
      if (parsedRoles.length > 0) settings.roles = parsedRoles;
    } else if (key === 'ChatGPT API 키' || key === 'ChatGPT API Key' || key === 'chatGptKey') {
      settings.chatGptKey = val;
    } else if (key === 'Claude API 키' || key === 'Claude API Key' || key === 'claudeKey') {
      settings.claudeKey = val;
    } else if (key === '동료 평가 참여 활성화' || key === 'enablePeerEvalPart') {
      settings.enablePeerEvalPart = (val !== 'N' && val !== 'false' && val !== 'FALSE' && val !== '0');
    } else if (key === '동료 평가 결과 확인 활성화' || key === 'enablePeerEvalResult') {
      settings.enablePeerEvalResult = (val !== 'N' && val !== 'false' && val !== 'FALSE' && val !== '0');
    } else if (key === '현재 활성 진도' || key === 'activeLesson') {
      if (val) settings.activeLesson = val;
    } else if (key === '진도 목록 (쉼표구분)' || key === '진도 목록' || key === 'lessonList') {
      const parsedLessons = val.split(',').map(s => s.trim()).filter(s => s);
      if (parsedLessons.length > 0) settings.lessonList = parsedLessons;
    }
  }
  
  if (!foundClasses && legacyNumClasses && legacyNumClasses > 0) {
    settings.classes = [];
    for (let i = 1; i <= legacyNumClasses; i++) {
      settings.classes.push(`${i}반`);
    }
  }
  
  if (!settings.roles || settings.roles.length === 0) {
    settings.roles = ['A', 'B', 'C', 'D'];
  }
  if (!settings.lessonList || settings.lessonList.length === 0) {
    settings.lessonList = ['1차시', '2차시', '3차시'];
  }
  if (!settings.activeLesson) {
    settings.activeLesson = settings.lessonList[0] || '1차시';
  }
  
  settings.numClasses = settings.classes.length;
  return settings;
}

function savePeerEvalActivation(key, isEnabled) {
  const currentSettings = getSettings();
  if (key === 'enablePeerEvalPart') {
    currentSettings.enablePeerEvalPart = !!isEnabled;
  } else if (key === 'enablePeerEvalResult') {
    currentSettings.enablePeerEvalResult = !!isEnabled;
  }
  saveSettings(currentSettings);
  return { success: true, settings: currentSettings };
}

function getPeerEvalActivationStatus() {
  const settings = getSettings();
  return {
    enablePeerEvalPart: settings.enablePeerEvalPart !== false,
    enablePeerEvalResult: settings.enablePeerEvalResult !== false
  };
}

function saveSettings(newSettings) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  if (!sheet) {
    setupSheets();
    sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  }
  
  const currentSettings = getSettings();
  
  let chatGptKey = (newSettings.chatGptKey !== undefined && newSettings.chatGptKey !== null && newSettings.chatGptKey.trim() !== '')
    ? newSettings.chatGptKey.trim()
    : currentSettings.chatGptKey;
    
  let claudeKey = (newSettings.claudeKey !== undefined && newSettings.claudeKey !== null && newSettings.claudeKey.trim() !== '')
    ? newSettings.claudeKey.trim()
    : currentSettings.claudeKey;

  let classesArr = [];
  if (Array.isArray(newSettings.classes)) {
    classesArr = newSettings.classes;
  } else if (typeof newSettings.classes === 'string') {
    classesArr = newSettings.classes.split(',').map(s => s.trim()).filter(s => s);
  } else if (newSettings.numClasses) {
    for (let i = 1; i <= parseInt(newSettings.numClasses); i++) classesArr.push(`${i}반`);
  }
  if (classesArr.length === 0) classesArr = ['1반', '2반', '3반'];

  let activeLesson = (newSettings.activeLesson !== undefined && newSettings.activeLesson !== null && String(newSettings.activeLesson).trim() !== '')
    ? String(newSettings.activeLesson).trim()
    : (currentSettings.activeLesson || '1차시');

  let lessonListArr = [];
  if (Array.isArray(newSettings.lessonList)) {
    lessonListArr = newSettings.lessonList;
  } else if (typeof newSettings.lessonList === 'string') {
    lessonListArr = newSettings.lessonList.split(',').map(s => s.trim()).filter(s => s);
  }
  if (lessonListArr.length === 0) lessonListArr = currentSettings.lessonList || ['1차시', '2차시', '3차시'];
  if (!lessonListArr.includes(activeLesson)) lessonListArr.unshift(activeLesson);

  const classesStr = classesArr.join(', ');
  const rolesStr = Array.isArray(newSettings.roles) ? newSettings.roles.join(', ') : (newSettings.roles || 'A, B, C, D');
  const lessonListStr = lessonListArr.join(', ');

  const enablePeerEvalPart = (newSettings.enablePeerEvalPart !== false && newSettings.enablePeerEvalPart !== 'N' && newSettings.enablePeerEvalPart !== 'false') ? 'Y' : 'N';
  const enablePeerEvalResult = (newSettings.enablePeerEvalResult !== false && newSettings.enablePeerEvalResult !== 'N' && newSettings.enablePeerEvalResult !== 'false') ? 'Y' : 'N';
  
  sheet.clearContents();
  sheet.getRange(1, 1, 10, 2).setValues([
    ['설정명', '값'],
    ['반이름 나열 (쉼표구분)', classesStr],
    ['반별 모둠 개수', newSettings.numGroups || 8],
    ['모둠원 역할 (쉼표구분)', rolesStr],
    ['ChatGPT API 키', chatGptKey || ''],
    ['Claude API 키', claudeKey || ''],
    ['동료 평가 참여 활성화', enablePeerEvalPart],
    ['동료 평가 결과 확인 활성화', enablePeerEvalResult],
    ['현재 활성 진도', activeLesson],
    ['진도 목록 (쉼표구분)', lessonListStr]
  ]);
  sheet.getRange("A1:B1").setFontWeight("bold");
  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 250);
  
  // 1-1. 원본정보 시트 검사 (단일 원본정보 시트 + 진도ID 열 방식 & 개별 차시 시트 방식 모두 지원)
  let defaultLyricsSheet = ss.getSheetByName(SHEET_NAMES.LYRICS) || ss.getSheetByName('원본정보') || ss.getSheetByName('원본가사');
  if (!defaultLyricsSheet) {
    defaultLyricsSheet = ss.insertSheet(SHEET_NAMES.LYRICS);
    defaultLyricsSheet.appendRow(['진도ID', '순번', '담당 역할', '우리말 문장', '영어 문장']);
    defaultLyricsSheet.getRange("A1:E1").setFontWeight("bold");
  }

  // 1-2. 필요한 반 시트 추가 / 헤더 갱신 / 이름 [학급] 변경
  classesArr.forEach(className => {
    let classSheet = findClassSheet(ss, className);
    const targetSheetName = `[학급] ${className}`;
    if (!classSheet) {
      classSheet = ss.insertSheet(targetSheetName);
      classSheet.appendRow(['순번', '영어 문장', '베스트 정답', '채택모둠', '진도ID']);
      classSheet.getRange("A1:E1").setFontWeight("bold");
    } else {
      if (classSheet.getName() !== targetSheetName) {
        classSheet.setName(targetSheetName);
      }
      classSheet.getRange(1, 1, 1, 5).setValues([['순번', '영어 문장', '베스트 정답', '채택모둠', '진도ID']]);
      classSheet.getRange("A1:E1").setFontWeight("bold");
    }
  });
  
  // 2. 나열된 반 목록에 포함되지 않는 불필요한 학급 시트만 삭제 (원본정보 관련 시트는 보존)
  const validSheetNames = classesArr.flatMap(c => [`[학급] ${c}`, `[학급]${c}`, c, `${c}반`]);
  const systemSheetNames = [
    SHEET_NAMES.SETTINGS,
    SHEET_NAMES.LYRICS,
    '원본정보',
    '원본가사',
    SHEET_NAMES.SUBMISSIONS,
    SHEET_NAMES.GRADE_BEST,
    SHEET_NAMES.AI_FEEDBACK,
    SHEET_NAMES.PEER_EVAL
  ];
  
  const sheets = ss.getSheets();
  sheets.forEach(s => {
    const sName = s.getName();
    const isOriginalSheet = sName.startsWith('원본') || sName.startsWith('[원본]');
    if (!systemSheetNames.includes(sName) && !validSheetNames.includes(sName) && !isOriginalSheet) {
      ss.deleteSheet(s);
    }
  });
  
  return true;
}

function getLyricsData(lessonId) {
  const settings = getSettings();
  const targetLesson = (lessonId && String(lessonId).trim()) ? String(lessonId).trim() : (settings.activeLesson || '1차시');
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const candidateNames = [
    `원본정보-${targetLesson}`,
    `원본가사-${targetLesson}`,
    `원본-${targetLesson}`,
    `원본정보_${targetLesson}`,
    `원본가사_${targetLesson}`,
    `원본_${targetLesson}`,
    `[원본] ${targetLesson}`,
    SHEET_NAMES.LYRICS,
    '원본정보',
    '원본가사'
  ];

  let sheet = null;
  for (let name of candidateNames) {
    sheet = ss.getSheetByName(name);
    if (sheet) break;
  }

  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0].map(h => String(h).trim());
  let lessonColIdx = -1;
  let seqColIdx = 0;
  let roleColIdx = 1;
  let korColIdx = 2;
  let engColIdx = 3;

  if (headers[0] === '진도ID' || headers[0] === '차시' || headers[0] === 'Lesson') {
    lessonColIdx = 0;
    seqColIdx = 1;
    roleColIdx = 2;
    korColIdx = 3;
    engColIdx = 4;
  }

  let lyrics = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (lessonColIdx !== -1) {
      const rowLesson = row[lessonColIdx] ? String(row[lessonColIdx]).trim() : '';
      if (rowLesson !== targetLesson && rowLesson !== '') continue;
    }
    if (row[seqColIdx]) {
      lyrics.push({
        seq: row[seqColIdx],
        role: row[roleColIdx] ? String(row[roleColIdx]).trim() : 'A',
        korean: row[korColIdx] || '',
        lyric: row[engColIdx] || row[korColIdx] || '',
        lessonId: targetLesson
      });
    }
  }
  return lyrics;
}

function getLyricsByRole(role, classNum, groupNum, lessonId) {
  const targetLesson = (lessonId && String(lessonId).trim()) ? String(lessonId).trim() : getSettings().activeLesson;
  const allLyrics = getLyricsData(targetLesson);
  const filtered = allLyrics.filter(item => item.role === role);

  if (classNum && groupNum) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const subSheet = ss.getSheetByName(SHEET_NAMES.SUBMISSIONS);
    if (subSheet && subSheet.getLastRow() > 1) {
      const subCols = Math.max(7, subSheet.getLastColumn());
      const subData = subSheet.getRange(2, 1, subSheet.getLastRow() - 1, subCols).getValues();
      let prevMap = {};
      subData.forEach(r => {
        let t_stamp = r[0];
        let t_class = r[1];
        let t_group = r[2];
        let t_role = r[3];
        let t_seq = r[4];
        let t_trans = r[5];
        let t_lesson = r[6] ? String(r[6]).trim() : '1차시';

        if (t_class == classNum && t_group == groupNum && t_role == role && (t_lesson == targetLesson || (!r[6] && targetLesson == '1차시'))) {
          if (!prevMap[t_seq] || t_stamp > prevMap[t_seq].timestamp) {
            prevMap[t_seq] = { translation: t_trans, timestamp: t_stamp };
          }
        }
      });
      filtered.forEach(item => {
        if (prevMap[item.seq]) {
          item.previousTranslation = prevMap[item.seq].translation;
        }
      });
    }
  }

  return filtered;
}

function submitTranslation(studentData, answers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.SUBMISSIONS);
  if (!sheet) throw new Error('제출현황 시트를 찾을 수 없습니다.');
  
  const targetLesson = studentData.lessonId || getSettings().activeLesson;
  const timestamp = new Date();
  const rows = answers.map(ans => [
    timestamp,
    studentData.classNum,
    studentData.groupNum,
    studentData.role,
    ans.seq,
    ans.translation,
    targetLesson
  ]);
  
  if (rows.length > 0) {
    const lock = LockService.getScriptLock();
    let hasLock = false;
    try {
      hasLock = lock.tryLock(10000);
      if (!hasLock) {
        throw new Error('현재 제출 요청이 몰려 대기시간이 초과되었습니다. 잠시 후 다시 제출해 주세요.');
      }
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    } catch (e) {
      throw new Error('제출 중 오류가 발생했습니다: ' + e.message);
    } finally {
      if (hasLock) {
        lock.releaseLock();
      }
    }
  }
  return true;
}

function getSubmittedAnswers(classNum, lessonId) {
  const targetLesson = (lessonId && String(lessonId).trim()) ? String(lessonId).trim() : getSettings().activeLesson;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetsMap = {};
  ss.getSheets().forEach(s => { sheetsMap[s.getName()] = s; });

  const lyrics = getLyricsData(targetLesson);

  const subSheet = sheetsMap[SHEET_NAMES.SUBMISSIONS];
  let subData = [];
  if (subSheet && subSheet.getLastRow() > 1) {
    const maxCols = Math.max(7, subSheet.getLastColumn());
    subData = subSheet.getRange(2, 1, subSheet.getLastRow() - 1, maxCols).getValues();
  }
  
  let subMap = {};
  subData.forEach(r => {
    let t_stamp = r[0];
    let t_class = r[1];
    let t_group = r[2];
    let t_seq = r[4];
    let t_trans = r[5];
    let t_lesson = r[6] ? String(r[6]).trim() : '1차시';
    
    if (t_class == classNum && (t_lesson == targetLesson || (!r[6] && targetLesson == '1차시'))) {
      if (!subMap[t_seq]) subMap[t_seq] = {};
      if (!subMap[t_seq][t_group] || t_stamp > subMap[t_seq][t_group].timestamp) {
        subMap[t_seq][t_group] = {
          translation: t_trans,
          timestamp: t_stamp
        };
      }
    }
  });
  
  const classSheet = findClassSheet(ss, classNum);
  let bestMap = {};
  if (classSheet && classSheet.getLastRow() > 1) {
    const maxCols = Math.max(5, classSheet.getLastColumn());
    const classData = classSheet.getRange(2, 1, classSheet.getLastRow() - 1, maxCols).getValues();
    classData.forEach(r => {
      let b_seq = r[0];
      let b_best = r[2];
      let b_groups = r[3];
      let b_lesson = r[4] ? String(r[4]).trim() : '1차시';
      if (b_lesson == targetLesson || (!r[4] && targetLesson == '1차시')) {
        bestMap[b_seq] = {
          best: b_best,
          groups: b_groups ? b_groups.toString().split(',').map(g => parseInt(g.trim())).filter(g => !isNaN(g)) : []
        };
      }
    });
  }
  
  return lyrics.map(lyric => {
    let submissionObj = {};
    if (subMap[lyric.seq]) {
      Object.keys(subMap[lyric.seq]).forEach(gNum => {
         submissionObj[gNum] = subMap[lyric.seq][gNum].translation;
      });
    }
    
    return {
      seq: lyric.seq,
      role: lyric.role,
      korean: lyric.korean || lyric.lyric,
      lyric: lyric.lyric || lyric.korean,
      submissions: submissionObj,
      best: bestMap[lyric.seq] ? bestMap[lyric.seq].best : "",
      bestGroups: bestMap[lyric.seq] ? bestMap[lyric.seq].groups : []
    };
  });
}

function saveClassBest(classNum, seq, lyricText, bestAnswer, groupIds, lessonId) {
  const targetLesson = (lessonId && String(lessonId).trim()) ? String(lessonId).trim() : getSettings().activeLesson;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let classSheet = findClassSheet(ss, classNum);
  const targetSheetName = `[학급] ${classNum}`;
  if (!classSheet) {
    classSheet = ss.insertSheet(targetSheetName);
    classSheet.appendRow(['순번', '영어 문장', '베스트 정답', '채택모둠', '진도ID']);
    classSheet.getRange("A1:E1").setFontWeight("bold");
  } else {
    classSheet.getRange(1, 1, 1, 5).setValues([['순번', '영어 문장', '베스트 정답', '채택모둠', '진도ID']]);
    classSheet.getRange("A1:E1").setFontWeight("bold");
  }
  
  const lastRow = classSheet.getLastRow();
  let foundRow = -1;
  if (lastRow > 1) {
    const maxCols = Math.max(5, classSheet.getLastColumn());
    const data = classSheet.getRange(2, 1, lastRow - 1, maxCols).getValues();
    for (let i = 0; i < data.length; i++) {
      let r_seq = data[i][0];
      let r_lesson = data[i][4] ? String(data[i][4]).trim() : '1차시';
      if (r_seq == seq && (r_lesson == targetLesson || (!data[i][4] && targetLesson == '1차시'))) {
        foundRow = i + 2;
        break;
      }
    }
  }
  
  const groupsStr = Array.isArray(groupIds) ? groupIds.join(', ') : (groupIds || '');
  
  if (foundRow > -1) {
    classSheet.getRange(foundRow, 2, 1, 4).setValues([[lyricText, bestAnswer, groupsStr, targetLesson]]);
  } else {
    classSheet.appendRow([seq, lyricText, bestAnswer, groupsStr, targetLesson]);
  }
  return true;
}

function getLeaderboard(classNum, lessonId) {
  const targetLesson = (lessonId && String(lessonId).trim()) ? String(lessonId).trim() : getSettings().activeLesson;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const classSheet = findClassSheet(ss, classNum);
  
  if (!classSheet || classSheet.getLastRow() <= 1) return [];
  
  const maxCols = Math.max(5, classSheet.getLastColumn());
  const data = classSheet.getRange(2, 1, classSheet.getLastRow() - 1, maxCols).getValues();
  let groupCounts = {};
  
  data.forEach(r => {
    let groupsStr = r[3];
    let r_lesson = r[4] ? String(r[4]).trim() : '1차시';
    if ((r_lesson == targetLesson || (!r[4] && targetLesson == '1차시')) && groupsStr) {
      let groups = groupsStr.toString().split(',').map(s => s.trim()).filter(s => s);
      groups.forEach(g => {
        groupCounts[g] = (groupCounts[g] || 0) + 1;
      });
    }
  });
  
  let result = Object.keys(groupCounts).map(g => ({
    group: parseInt(g),
    count: groupCounts[g]
  }));
  
  result.sort((a, b) => b.count - a.count);
  return result;
}

function getAllClassesBest(lessonId) {
  const targetLesson = (lessonId && String(lessonId).trim()) ? String(lessonId).trim() : getSettings().activeLesson;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetsMap = {};
  ss.getSheets().forEach(s => { sheetsMap[s.getName()] = s; });
  
  const lyrics = getLyricsData(targetLesson);
  const settings = getSettings();
  
  const gradeSheet = sheetsMap[SHEET_NAMES.GRADE_BEST];
  let gradeBestMap = {};
  if (gradeSheet && gradeSheet.getLastRow() > 1) {
    const maxCols = Math.max(5, gradeSheet.getLastColumn());
    const gradeData = gradeSheet.getRange(2, 1, gradeSheet.getLastRow() - 1, maxCols).getValues();
    gradeData.forEach(r => {
      let g_lesson = r[4] ? String(r[4]).trim() : '1차시';
      if (g_lesson == targetLesson || (!r[4] && targetLesson == '1차시')) {
        gradeBestMap[r[0]] = {
          best: r[2],
          source: r[3]
        };
      }
    });
  }
  
  let classBestsMap = {};
  settings.classes.forEach(cName => {
    classBestsMap[cName] = {};
    let cSheet = findClassSheet(ss, cName);
    if (cSheet && cSheet.getLastRow() > 1) {
      const maxCols = Math.max(5, cSheet.getLastColumn());
      const cData = cSheet.getRange(2, 1, cSheet.getLastRow() - 1, maxCols).getValues();
      cData.forEach(r => {
        let c_lesson = r[4] ? String(r[4]).trim() : '1차시';
        if (c_lesson == targetLesson || (!r[4] && targetLesson == '1차시')) {
          classBestsMap[cName][r[0]] = {
            best: r[2],
            groups: r[3]
          };
        }
      });
    }
  });
  
  return lyrics.map(lyric => {
    let cb = {};
    settings.classes.forEach(cName => {
      if (classBestsMap[cName] && classBestsMap[cName][lyric.seq]) {
        cb[cName] = classBestsMap[cName][lyric.seq];
      }
    });
    
    return {
      seq: lyric.seq,
      korean: lyric.korean,
      lyric: lyric.lyric,
      classBests: cb,
      gradeBest: gradeBestMap[lyric.seq] ? gradeBestMap[lyric.seq].best : "",
      gradeBestSource: gradeBestMap[lyric.seq] ? gradeBestMap[lyric.seq].source : ""
    };
  });
}

function saveGradeBest(seq, lyricText, bestAnswer, sourceStr, lessonId) {
  const targetLesson = (lessonId && String(lessonId).trim()) ? String(lessonId).trim() : getSettings().activeLesson;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let gradeSheet = ss.getSheetByName(SHEET_NAMES.GRADE_BEST);
  if (!gradeSheet) {
    gradeSheet = ss.insertSheet(SHEET_NAMES.GRADE_BEST);
    gradeSheet.appendRow(['순번', '우리말 문장', '학년 베스트 영어 문장', '출처', '진도ID']);
    gradeSheet.getRange("A1:E1").setFontWeight("bold");
  } else {
    gradeSheet.getRange(1, 1, 1, 5).setValues([['순번', '우리말 문장', '학년 베스트 영어 문장', '출처', '진도ID']]);
    gradeSheet.getRange("A1:E1").setFontWeight("bold");
  }
  
  const lastRow = gradeSheet.getLastRow();
  let foundRow = -1;
  if (lastRow > 1) {
    const maxCols = Math.max(5, gradeSheet.getLastColumn());
    const data = gradeSheet.getRange(2, 1, lastRow - 1, maxCols).getValues();
    for (let i = 0; i < data.length; i++) {
      let r_seq = data[i][0];
      let r_lesson = data[i][4] ? String(data[i][4]).trim() : '1차시';
      if (r_seq == seq && (r_lesson == targetLesson || (!data[i][4] && targetLesson == '1차시'))) {
        foundRow = i + 2;
        break;
      }
    }
  }
  
  if (foundRow > -1) {
    gradeSheet.getRange(foundRow, 2, 1, 4).setValues([[lyricText, bestAnswer, sourceStr, targetLesson]]);
  } else {
    gradeSheet.appendRow([seq, lyricText, bestAnswer, sourceStr, targetLesson]);
  }
  return true;
}

// --- AI Feedback Integration ---

function callChatGPT(apiKey, systemPrompt, userPrompt) {
  const url = 'https://api.openai.com/v1/chat/completions';
  const payload = {
    model: 'gpt-4o',
    messages: [
      { role: 'system', content: systemPrompt },
      { role: 'user', content: userPrompt }
    ],
    response_format: { type: "json_object" },
    temperature: 0.3
  };
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': 'Bearer ' + apiKey
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch(url, options);
  const resCode = response.getResponseCode();
  const resText = response.getContentText();
  if (resCode !== 200) {
    throw new Error('ChatGPT API 호출 실패 (' + resCode + '): ' + resText);
  }
  const json = JSON.parse(resText);
  return json.choices[0].message.content;
}

function callClaude(apiKey, systemPrompt, userPrompt) {
  if (!apiKey || typeof apiKey !== 'string') {
    throw new Error("Claude API 키가 설정되어 있지 않습니다. '설정' 화면 또는 B6셀에 API 키를 등록해 주세요.");
  }

  const cleanKey = apiKey.trim();
  if (cleanKey.startsWith('sk-proj-')) {
    throw new Error("Claude API 키 자리에 OpenAI API 키('sk-proj-...')가 입력되었습니다. '설정' 화면의 B6셀에 Anthropic Claude 전용 API 키('sk-ant-...')를 입력해 주세요.");
  }

  const url = 'https://api.anthropic.com/v1/messages';
  const candidateModels = [
    { id: 'claude-3-7-sonnet-20250219', name: 'Claude 3.7 Sonnet' },
    { id: 'claude-3-5-sonnet-20241022', name: 'Claude 3.5 Sonnet' },
    { id: 'claude-3-5-haiku-20241022', name: 'Claude 3.5 Haiku' },
    { id: 'claude-3-5-sonnet-20240620', name: 'Claude 3.5 Sonnet (20240620)' },
    { id: 'claude-3-haiku-20240307', name: 'Claude 3 Haiku' }
  ];

  let lastErrorText = '';
  for (let modelObj of candidateModels) {
    const payload = {
      model: modelObj.id,
      max_tokens: 1000,
      system: systemPrompt,
      messages: [
        { role: 'user', content: userPrompt }
      ]
    };
    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'x-api-key': cleanKey,
        'anthropic-version': '2023-06-01'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const resCode = response.getResponseCode();
    const resText = response.getContentText();

    if (resCode === 200) {
      const json = JSON.parse(resText);
      return {
        text: json.content[0].text,
        modelName: `Claude (${modelObj.name})`
      };
    }

    lastErrorText = resText;
    if (resCode === 404) {
      continue;
    } else {
      throw new Error('Claude API 호출 실패 (' + resCode + '): ' + resText);
    }
  }

  throw new Error('Claude API 호출 실패 (404 Not Found):\n입력하신 Anthropic API 키에 이용 가능한 Claude 모델 권한이 없거나 결제 크레딧이 부족합니다.\n1. console.anthropic.com에 접속하여 API 키(sk-ant-...)가 올바른지 확인해 주세요.\n2. Plans & Billing 메뉴에서 최소 크레딧($5) 충전 여부를 확인해 주세요.');
}

function parseAiResponseJson(rawText) {
  if (!rawText) return { grammar: '-', vocabulary: '-', expression: '-', overall: '-' };

  let clean = rawText.trim();
  
  const jsonBlockMatch = clean.match(/```(?:json)?\s*([\s\S]*?)\s*```/i);
  if (jsonBlockMatch) {
    clean = jsonBlockMatch[1].trim();
  } else {
    const firstBrace = clean.indexOf('{');
    const lastBrace = clean.lastIndexOf('}');
    if (firstBrace !== -1 && lastBrace > firstBrace) {
      clean = clean.substring(firstBrace, lastBrace + 1).trim();
    }
  }

  if (clean.includes('""')) {
    clean = clean.replace(/""/g, '"');
  }

  let obj = null;

  try {
    obj = JSON.parse(clean);
  } catch (e1) {
    try {
      let fixedJson = clean.replace(/\r?\n/g, function(match, offset, string) {
        let quotes = 0;
        for (let i = 0; i < offset; i++) {
          if (string[i] === '"' && (i === 0 || string[i - 1] !== '\\')) {
            quotes++;
          }
        }
        return (quotes % 2 === 1) ? '\\n' : match;
      });
      obj = JSON.parse(fixedJson);
    } catch (e2) {
      let grammar = '', vocabulary = '', expression = '', overall = '';
      
      const gMatch = clean.match(/"(?:grammar|grammarFeedback|Grammar)"\s*:\s*"([\s\S]*?)"\s*,\s*"/i);
      const vMatch = clean.match(/"(?:vocabulary|vocabFeedback|Vocabulary)"\s*:\s*"([\s\S]*?)"\s*,\s*"/i);
      const eMatch = clean.match(/"(?:expression|exprFeedback|Expression)"\s*:\s*"([\s\S]*?)"\s*,\s*"/i);
      const oMatch = clean.match(/"(?:overall|paraphrase|Overall)"\s*:\s*"([\s\S]*?)"\s*\}?$/i);

      if (gMatch) grammar = gMatch[1].replace(/\\n/g, '\n').trim();
      if (vMatch) vocabulary = vMatch[1].replace(/\\n/g, '\n').trim();
      if (eMatch) expression = eMatch[1].replace(/\\n/g, '\n').trim();
      if (oMatch) overall = oMatch[1].replace(/\\n/g, '\n').trim();

      if (grammar || vocabulary || expression || overall) {
        obj = { grammar, vocabulary, expression, overall };
      }
    }
  }

  if (obj && typeof obj === 'object') {
    const getKey = (...keys) => {
      for (let k of keys) {
        if (obj[k] !== undefined && obj[k] !== null && String(obj[k]).trim() !== '') {
          return String(obj[k]).trim();
        }
      }
      const lowerObj = {};
      Object.keys(obj).forEach(key => { lowerObj[key.toLowerCase()] = obj[key]; });
      for (let k of keys) {
        const lk = k.toLowerCase();
        if (lowerObj[lk] !== undefined && lowerObj[lk] !== null && String(lowerObj[lk]).trim() !== '') {
          return String(lowerObj[lk]).trim();
        }
      }
      return '';
    };

    return {
      grammar: getKey('grammar', 'Grammar', 'grammarFeedback', 'grammar_feedback', '문법') || '-',
      vocabulary: getKey('vocabulary', 'Vocabulary', 'vocabFeedback', 'vocabulary_feedback', '어휘') || '-',
      expression: getKey('expression', 'Expression', 'exprFeedback', 'expression_feedback', '표현') || '-',
      overall: getKey('overall', 'Overall', 'paraphrase', 'Paraphrase', 'paraphrasing', 'overallFeedback') || rawText
    };
  }

  return {
    grammar: "피드백 파싱 오류 (형식이 맞지 않음)",
    vocabulary: "-",
    expression: "-",
    overall: rawText
  };
}

function getAiFeedbackData(classNum, lessonId) {
  const targetLesson = (lessonId && String(lessonId).trim()) ? String(lessonId).trim() : getSettings().activeLesson;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aiSheet = ss.getSheetByName(SHEET_NAMES.AI_FEEDBACK);
  if (!aiSheet || aiSheet.getLastRow() <= 1) return [];

  const maxCols = Math.max(13, aiSheet.getLastColumn());
  const data = aiSheet.getRange(2, 1, aiSheet.getLastRow() - 1, maxCols).getValues();
  let list = [];
  data.forEach(r => {
    let t_stamp = r[0];
    let t_class = r[1];
    let t_group = r[2];
    let t_role = r[3];
    let t_seq = r[4];
    let t_korean = r[5];
    let t_translation = r[6];
    let t_grammar = r[7];
    let t_vocab = r[8];
    let t_expr = r[9];
    let t_overall = r[10];
    let t_provider = r[11];
    let t_lesson = r[12] ? String(r[12]).trim() : '1차시';

    if ((!classNum || t_class == classNum) && (t_lesson == targetLesson || (!r[12] && targetLesson == '1차시'))) {
      list.push({
        timestamp: t_stamp,
        classNum: t_class,
        group: t_group,
        role: t_role,
        seq: t_seq,
        korean: t_korean,
        translation: t_translation,
        grammar: t_grammar,
        vocabulary: t_vocab,
        expression: t_expr,
        overall: t_overall,
        provider: t_provider,
        lessonId: t_lesson
      });
    }
  });
  return list;
}

function generateAiFeedback(classNum, provider, lessonId) {
  const targetLesson = (lessonId && String(lessonId).trim()) ? String(lessonId).trim() : getSettings().activeLesson;
  const settings = getSettings();
  let apiKey = '';
  let providerName = '';

  if (provider === 'claude') {
    apiKey = settings.claudeKey;
    providerName = 'Claude (claude-3-5-sonnet)';
    if (!apiKey) {
      throw new Error("Claude API 키가 설정되어 있지 않습니다. '설정' 화면 또는 B6셀에 API 키를 등록해 주세요.");
    }
  } else {
    apiKey = settings.chatGptKey;
    providerName = 'ChatGPT (gpt-4o)';
    if (!apiKey) {
      throw new Error("ChatGPT API 키가 설정되어 있지 않습니다. '설정' 화면 또는 B5셀에 API 키를 등록해 주세요.");
    }
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const subSheet = ss.getSheetByName(SHEET_NAMES.SUBMISSIONS);
  if (!subSheet || subSheet.getLastRow() <= 1) {
    throw new Error('제출현황 시트에 학생 제출 데이터가 없습니다.');
  }

  const maxSubCols = Math.max(7, subSheet.getLastColumn());
  const subData = subSheet.getRange(2, 1, subSheet.getLastRow() - 1, maxSubCols).getValues();
  
  let latestMap = {};
  subData.forEach(r => {
    let t_stamp = r[0];
    let t_class = r[1];
    let t_group = r[2];
    let t_role = r[3];
    let t_seq = r[4];
    let t_trans = r[5];
    let t_lesson = r[6] ? String(r[6]).trim() : '1차시';

    if (t_class == classNum && (t_lesson == targetLesson || (!r[6] && targetLesson == '1차시')) && t_trans && String(t_trans).trim()) {
      let key = `${t_seq}_${t_group}`;
      if (!latestMap[key] || t_stamp > latestMap[key].timestamp) {
        latestMap[key] = {
          timestamp: t_stamp,
          classNum: t_class,
          group: t_group,
          role: t_role,
          seq: t_seq,
          translation: t_trans
        };
      }
    }
  });

  const targets = Object.values(latestMap);
  if (targets.length === 0) {
    throw new Error(`${classNum}반 학생들의 제출된 영어 번역이 없습니다.`);
  }

  const lyrics = getLyricsData(targetLesson);
  let lyricMap = {};
  lyrics.forEach(l => { lyricMap[l.seq] = l.korean || l.lyric; });

  let aiSheet = ss.getSheetByName(SHEET_NAMES.AI_FEEDBACK);
  if (!aiSheet) {
    setupSheets();
    aiSheet = ss.getSheetByName(SHEET_NAMES.AI_FEEDBACK);
  } else {
    aiSheet.getRange(1, 1, 1, 13).setValues([['타임스탬프', '반', '모둠', '역할', '순번', '우리말 원문', '학생 영문 번역', '문법 피드백', '어휘 피드백', '표현 피드백', 'Paraphrase 피드백', '제공 AI', '진도ID']]);
    aiSheet.getRange("A1:M1").setFontWeight("bold");
  }

  const systemPrompt = `당신은 친절하고 전문적인 중·고등학교 영어 선생님이자 번역 및 Paraphrasing 평가 전문가입니다.
학생들이 우리말 문장을 보고 영어 문장으로 번역한 결과물을 분석하고 피드백을 제공합니다.

[핵심 평가 지침]
1. 잘못된 지식 수정: 학생 번역에서 발생한 문법적, 어휘적 오류를 정확히 파악하여 올바르게 고쳐줍니다.
2. 모르는 정보 제공: 학생이 알지 못했던 유용한 어휘, 구문, 시제/수동태/물주구문 등 다채로운 표현 정보를 제공합니다.
3. Paraphrasing 능력 향상: 학생이 표현력을 다각도로 넓힐 수 있도록 Paraphrasing 기법을 지도합니다.

[작성 규칙]
1. 문법(grammar), 어휘(vocabulary), 표현(expression) 피드백은 반드시 100% 한글(한국어)로 다정하고 알기 쉽게 1-2문장으로 작성하세요.
2. "overall" 필드에는 학생 번역을 바탕으로 한 Paraphrase 문장을 최대 3개까지 작성하여 하나의 텍스트(줄바꿈으로 구분)로 제공하세요.
   Paraphrase 3가지 문장 구성 조건:
   - 1) [주어 변경 문장]: 주어를 변경하여 태(수동태/능동태) 변화, 물주구문(사물주어) 등 다양한 구문을 활용한 문장
   - 2) [부정어 포함 문장]: 부정 표현(do not V, cannot help -ing, never 등)을 활용한 문장
   - 3) [부정어 미포함 문장]: 긍정 및 유의 표현(like, keep from -ing, prevent from -ing 등)을 활용한 문장
   (각 문장은 영문과 함께 어떤 구문 구조가 적용되었는지 간단한 한글 부연 설명을 괄호 안에 첨부해 주세요.)
3. 반드시 아래 JSON 구조로만 응답하세요. 다른 설명이나 텍스트는 절대 금지합니다.
{
  "grammar": "잘못된 문법 교정 및 성수/시제 일치에 대한 한글 피드백 1-2문장",
  "vocabulary": "어휘 선택의 적절성 및 어휘 교정/추천에 대한 한글 피드백 1-2문장",
  "expression": "표현의 자연스러움 및 구문 활용에 대한 한글 피드백 1-2문장",
  "overall": "1. [주어 변경] (Paraphrase 영문) (구문 설명)\n2. [부정어 포함] (Paraphrase 영문) (구문 설명)\n3. [부정어 미포함] (Paraphrase 영문) (구문 설명)"
}`;

  const timestamp = new Date();
  
  // 기존 AI피드백 데이터 맵핑 및 이미 완성된 피드백 판별
  let existingAiMap = {};
  if (aiSheet.getLastRow() > 1) {
    const existingData = aiSheet.getRange(2, 1, aiSheet.getLastRow() - 1, 12).getValues();
    existingData.forEach((r, idx) => {
      let e_class = r[1];
      let e_group = r[2];
      let e_seq = r[4];
      let grammarFeedback = r[7] ? r[7].toString().trim() : '';
      let vocabFeedback = r[8] ? r[8].toString().trim() : '';
      let exprFeedback = r[9] ? r[9].toString().trim() : '';
      let overallFeedback = r[10] ? r[10].toString().trim() : '';

      // 피드백 셀(문법, 어휘, 표현, Paraphrase)이 하나라도 비어있지 않고 기록되어 있는지 검사
      let hasFeedback = Boolean(grammarFeedback || vocabFeedback || exprFeedback || overallFeedback);

      if (e_class == classNum) {
        existingAiMap[`${e_seq}_${e_group}`] = {
          rowIndex: idx + 2,
          hasFeedback: hasFeedback
        };
      }
    });
  }

  targets.forEach(target => {
    const key = `${target.seq}_${target.group}`;
    const existingInfo = existingAiMap[key];

    // 이미 AI 피드백 셀이 채워져 있는 경우 다시 피드백을 요청하지 않고 스킵
    if (existingInfo && existingInfo.hasFeedback) {
      return;
    }

    const koreanText = lyricMap[target.seq] || '';
    const userPrompt = `- 우리말 원문: "${koreanText}"
- 학생 영어 번역: "${target.translation}"

위 학생의 번역을 바탕으로:
1) 잘못된 정보 교정 및 문법(grammar), 어휘(vocabulary), 표현(expression) 피드백을 한글로 작성해 주세요.
2) 주어 변경 문장, 부정어 포함 문장, 부정어 미포함 문장의 3가지 Paraphrase 제안 문장을 생성하여 "overall" 항목에 줄바꿈으로 구분하여 제공해 주세요.`;

    let rawResponse = '';
    let actualProviderName = providerName;
    if (provider === 'claude') {
      const claudeRes = callClaude(apiKey, systemPrompt, userPrompt);
      rawResponse = claudeRes.text;
      actualProviderName = claudeRes.modelName;
    } else {
      rawResponse = callChatGPT(apiKey, systemPrompt, userPrompt);
    }

    const parsed = parseAiResponseJson(rawResponse);
    const rowValues = [
      timestamp,
      target.classNum,
      target.group,
      target.role,
      target.seq,
      koreanText,
      target.translation,
      parsed.grammar || '',
      parsed.vocabulary || '',
      parsed.expression || '',
      parsed.overall || '',
      actualProviderName
    ];

    if (existingInfo && existingInfo.rowIndex) {
      // 피드백이 비어있던 기존 행 업데이트
      aiSheet.getRange(existingInfo.rowIndex, 1, 1, 12).setValues([rowValues]);
    } else {
      // 신규 행 추가
      aiSheet.appendRow(rowValues);
    }
  });

  return getAiFeedbackData(classNum);
}

function getInitialData() {
  let serviceUrl = '';
  try {
    serviceUrl = ScriptApp.getService().getUrl();
  } catch (err) {
    serviceUrl = '';
  }
  return {
    settings: getSettings(),
    webAppUrl: serviceUrl
  };
}

function clearTestData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. 제출현황 시트 초기화 (헤더 남기고 삭제)
  const subSheet = ss.getSheetByName(SHEET_NAMES.SUBMISSIONS);
  if (subSheet && subSheet.getLastRow() > 1) {
    subSheet.getRange(2, 1, subSheet.getLastRow() - 1, subSheet.getLastColumn()).clearContent();
  }
  
  // 2. 시스템 시트를 제외한 모든 반 시트 초기화
  const systemSheetNames = [
    SHEET_NAMES.SETTINGS,
    SHEET_NAMES.LYRICS,
    '원본가사',
    SHEET_NAMES.SUBMISSIONS,
    SHEET_NAMES.GRADE_BEST,
    SHEET_NAMES.AI_FEEDBACK
  ];
  
  const sheets = ss.getSheets();
  sheets.forEach(s => {
    if (!systemSheetNames.includes(s.getName()) && s.getLastRow() > 1) {
      s.getRange(2, 1, s.getLastRow() - 1, s.getLastColumn()).clearContent();
    }
  });
  
  // 3. 학년최종결과 시트 초기화
  const gradeSheet = ss.getSheetByName(SHEET_NAMES.GRADE_BEST);
  if (gradeSheet && gradeSheet.getLastRow() > 1) {
    gradeSheet.getRange(2, 1, gradeSheet.getLastRow() - 1, gradeSheet.getLastColumn()).clearContent();
  }

  // 4. AI피드백 시트 초기화
  const aiSheet = ss.getSheetByName(SHEET_NAMES.AI_FEEDBACK);
  if (aiSheet && aiSheet.getLastRow() > 1) {
    aiSheet.getRange(2, 1, aiSheet.getLastRow() - 1, aiSheet.getLastColumn()).clearContent();
  }
  
  // 5. 동료평가 시트 초기화
  const peerSheet = ss.getSheetByName(SHEET_NAMES.PEER_EVAL);
  if (peerSheet && peerSheet.getLastRow() > 1) {
    peerSheet.getRange(2, 1, peerSheet.getLastRow() - 1, peerSheet.getLastColumn()).clearContent();
  }

  return true;
}

function getComparisonData(lessonId) {
  const targetLesson = (lessonId && String(lessonId).trim()) ? String(lessonId).trim() : getSettings().activeLesson;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetsMap = {};
  ss.getSheets().forEach(s => { sheetsMap[s.getName()] = s; });
  
  const lyrics = getLyricsData(targetLesson);

  const gradeSheet = sheetsMap[SHEET_NAMES.GRADE_BEST];
  let gradeBestMap = {};
  if (gradeSheet && gradeSheet.getLastRow() > 1) {
    const maxCols = Math.max(5, gradeSheet.getLastColumn());
    const gradeData = gradeSheet.getRange(2, 1, gradeSheet.getLastRow() - 1, maxCols).getValues();
    gradeData.forEach(r => {
      let g_lesson = r[4] ? String(r[4]).trim() : '1차시';
      if (g_lesson == targetLesson || (!r[4] && targetLesson == '1차시')) {
        gradeBestMap[r[0]] = {
          best: r[2] || '',
          source: r[3] || ''
        };
      }
    });
  }

  return lyrics.map(lyric => {
    const gBest = gradeBestMap[lyric.seq] || {};
    return {
      seq: lyric.seq,
      role: lyric.role,
      korean: lyric.korean,
      originalEnglish: lyric.lyric,
      gradeBest: gBest.best || '',
      gradeBestSource: gBest.source || ''
    };
  });
}

// --- 동료 평가 (Peer Evaluation) Backend Functions ---

function getPeerEvalSentencesToEvaluate(classNum, groupNum, role) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let peerSheet = ss.getSheetByName(SHEET_NAMES.PEER_EVAL);
    if (!peerSheet) {
      setupSheets();
      peerSheet = ss.getSheetByName(SHEET_NAMES.PEER_EVAL);
    }

    // 1. 이미 이 평가자가 평가를 제출했는지 확인
    if (peerSheet && peerSheet.getLastRow() > 1 && peerSheet.getLastColumn() > 0) {
      const peerData = peerSheet.getRange(2, 1, peerSheet.getLastRow() - 1, peerSheet.getLastColumn()).getValues();
      const evaluatorRecords = peerData.filter(r => r[1] == classNum && r[2] == groupNum && r[3] == role);
      if (evaluatorRecords.length > 0) {
        const prevEvals = evaluatorRecords.map(r => ({
          timestamp: r[0] ? r[0].toString() : '',
          evaluatorClass: r[1],
          evaluatorGroup: r[2],
          evaluatorRole: r[3],
          targetGroup: r[4],
          targetRole: r[5],
          targetSeq: r[6],
          korean: r[7] || '',
          targetTranslation: r[8] || '',
          grammarScore: Number(r[9]) || 0,
          vocabScore: Number(r[10]) || 0,
          exprScore: Number(r[11]) || 0,
          totalScore: Number(r[12]) || 0,
          comment: r[13] || '',
          evalType: r[14] || ''
        }));
        return { alreadyEvaluated: true, evaluations: prevEvals };
      }
    }

    // 2. 누적 평가 횟수 집계 (공평 배분 알고리즘)
    let evalCounts = {};
    if (peerSheet && peerSheet.getLastRow() > 1 && peerSheet.getLastColumn() > 0) {
      const peerData = peerSheet.getRange(2, 1, peerSheet.getLastRow() - 1, peerSheet.getLastColumn()).getValues();
      peerData.forEach(r => {
        let pClass = r[1];
        let pTargetGroup = r[4];
        let pTargetSeq = r[6];
        if (pClass == classNum) {
          let key = `${pTargetSeq}_${pTargetGroup}`;
          evalCounts[key] = (evalCounts[key] || 0) + 1;
        }
      });
    }

    // 원본 우리말 문장 맵핑
    const lyrics = getLyricsData();
    let lyricMap = {};
    lyrics.forEach(l => { lyricMap[l.seq] = l.korean || l.lyric; });

    // 3. 해당 반 학생들이 제출한 영문 번역 수집
    const subSheet = ss.getSheetByName(SHEET_NAMES.SUBMISSIONS);
    let otherGroupMap = {};
    let allSubMap = {};

    if (subSheet && subSheet.getLastRow() > 1 && subSheet.getLastColumn() > 0) {
      const subData = subSheet.getRange(2, 1, subSheet.getLastRow() - 1, subSheet.getLastColumn()).getValues();
      subData.forEach(r => {
        let t_stamp = r[0];
        let t_class = r[1];
        let t_group = r[2];
        let t_role = r[3];
        let t_seq = r[4];
        let t_trans = r[5];

        if (t_class == classNum && t_trans && String(t_trans).trim() !== '') {
          let key = `${t_seq}_${t_group}`;
          let getTimeVal = (ts) => (ts instanceof Date ? ts.getTime() : new Date(ts).getTime() || 0);
          let item = {
            timestamp: t_stamp,
            classNum: t_class,
            group: t_group,
            role: t_role,
            seq: t_seq,
            translation: String(t_trans),
            evalCount: evalCounts[key] || 0,
            korean: lyricMap[t_seq] || ''
          };

          if (!allSubMap[key] || getTimeVal(t_stamp) > getTimeVal(allSubMap[key].timestamp)) {
            allSubMap[key] = item;
          }
          if (t_group != groupNum) {
            if (!otherGroupMap[key] || getTimeVal(t_stamp) > getTimeVal(otherGroupMap[key].timestamp)) {
              otherGroupMap[key] = item;
            }
          }
        }
      });
    }

    let candidates = Object.values(otherGroupMap);
    let allCandidates = Object.values(allSubMap);

    // Fallback 1: 다른 모둠 제출 문장이 전혀 없으면 학급 내 모든 제출 문장 사용 (자주 배부되더라도 진행)
    if (candidates.length === 0) {
      candidates = allCandidates;
    }

    // Fallback 2: 학급 내 제출 문장이 전혀 없는 경우 '원본정보' 시트의 가사 원문 문장으로 생성하여 절대 차단되지 않도록 보장
    if (candidates.length === 0) {
      lyrics.forEach((l, idx) => {
        if (l.lyric || l.korean) {
          candidates.push({
            timestamp: new Date(),
            classNum: classNum,
            group: 0,
            role: l.role || (idx % 2 === 0 ? role : 'B'),
            seq: l.seq,
            translation: l.lyric || l.korean,
            evalCount: 0,
            korean: l.korean || l.lyric
          });
        }
      });
    }

    // Candidate Pool 1: 동일한 역할 (role == evaluatorRole)
    let sameRolePool = candidates.filter(c => c.role == role);
    // Candidate Pool 2: 다른 역할 (role != evaluatorRole)
    let diffRolePool = candidates.filter(c => c.role != role);

    // 역할에 맞는 문장이 부족하면 전체 candidates에서 충원
    if (sameRolePool.length === 0) sameRolePool = candidates;
    if (diffRolePool.length === 0) diffRolePool = candidates;

    function pickMinEvalCandidate(pool, excludeKey) {
      if (!pool || pool.length === 0) return null;
      let filtered = pool;
      if (excludeKey && pool.length > 1) {
        filtered = pool.filter(c => `${c.seq}_${c.group}` !== excludeKey);
      }
      if (filtered.length === 0) filtered = pool;

      let minCount = Math.min(...filtered.map(c => c.evalCount));
      let mins = filtered.filter(c => c.evalCount === minCount);
      let randomIndex = Math.floor(Math.random() * mins.length);
      const chosen = mins[randomIndex];
      return {
        seq: chosen.seq,
        group: chosen.group,
        role: chosen.role,
        korean: chosen.korean,
        translation: chosen.translation
      };
    }

    const selectedSame = pickMinEvalCandidate(sameRolePool);
    const excludeKey = selectedSame ? `${selectedSame.seq}_${selectedSame.group}` : null;
    const selectedDiff = pickMinEvalCandidate(diffRolePool, excludeKey);

    return {
      alreadyEvaluated: false,
      targetSameRole: selectedSame,
      targetDiffRole: selectedDiff
    };
  } catch (err) {
    return { alreadyEvaluated: false, message: '동료 평가 데이터 처리 중 오류가 발생했습니다: ' + err.message };
  }
}

function submitPeerEvaluation(evaluatorData, evaluations) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let peerSheet = ss.getSheetByName(SHEET_NAMES.PEER_EVAL);
  if (!peerSheet) {
    setupSheets();
    peerSheet = ss.getSheetByName(SHEET_NAMES.PEER_EVAL);
  }

  const timestamp = new Date();
  const rows = evaluations.map(ev => [
    timestamp,
    evaluatorData.classNum,
    evaluatorData.groupNum,
    evaluatorData.role,
    ev.targetGroup,
    ev.targetRole,
    ev.targetSeq,
    ev.korean || '',
    ev.targetTranslation || '',
    ev.grammarScore ? 1 : 0,
    ev.vocabScore ? 1 : 0,
    ev.exprScore ? 1 : 0,
    (ev.grammarScore ? 1 : 0) + (ev.vocabScore ? 1 : 0) + (ev.exprScore ? 1 : 0),
    ev.comment || '',
    ev.evalType || ''
  ]);

  if (rows.length > 0) {
    const lock = LockService.getScriptLock();
    let hasLock = false;
    try {
      hasLock = lock.tryLock(10000);
      if (!hasLock) {
        throw new Error('현재 제출 요청이 몰려 대기시간이 초과되었습니다. 잠시 후 다시 제출해 주세요.');
      }
      peerSheet.getRange(peerSheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    } catch (e) {
      throw new Error('동료 평가 제출 중 오류가 발생했습니다: ' + e.message);
    } finally {
      if (hasLock) {
        lock.releaseLock();
      }
    }
  }
  return true;
}

function getPeerEvaluationResults(classNum, groupNum, role) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 해당 역할의 모든 가사 원문
    const lyrics = getLyricsData().filter(l => l.role === role);
    if (lyrics.length === 0) return [];

    // 해당 학생(groupNum, role)이 제출한 최신 영어 번역 가져오기
    const subSheet = ss.getSheetByName(SHEET_NAMES.SUBMISSIONS);
    let subMap = {};
    if (subSheet && subSheet.getLastRow() > 1 && subSheet.getLastColumn() > 0) {
      const subData = subSheet.getRange(2, 1, subSheet.getLastRow() - 1, subSheet.getLastColumn()).getValues();
      subData.forEach(r => {
        let t_stamp = r[0];
        let t_class = r[1];
        let t_group = r[2];
        let t_role = r[3];
        let t_seq = r[4];
        let t_trans = r[5];
        if (t_class == classNum && t_group == groupNum && t_role == role) {
          let getTimeVal = (ts) => (ts instanceof Date ? ts.getTime() : new Date(ts).getTime() || 0);
          if (!subMap[t_seq] || getTimeVal(t_stamp) > getTimeVal(subMap[t_seq].timestamp)) {
            subMap[t_seq] = { translation: String(t_trans), timestamp: t_stamp };
          }
        }
      });
    }

    // 동료 평가 시트에서 이 학생 문장에 대해 제출된 피드백들 추출
    let peerMap = {};
    const peerSheet = ss.getSheetByName(SHEET_NAMES.PEER_EVAL);
    if (peerSheet && peerSheet.getLastRow() > 1 && peerSheet.getLastColumn() > 0) {
      const peerData = peerSheet.getRange(2, 1, peerSheet.getLastRow() - 1, peerSheet.getLastColumn()).getValues();
      peerData.forEach(r => {
        let pClass = r[1];
        let pEvalGroup = r[2];
        let pEvalRole = r[3];
        let pTargetGroup = r[4];
        let pTargetRole = r[5];
        let pTargetSeq = r[6];
        let pGrammar = r[9];
        let pVocab = r[10];
        let pExpr = r[11];
        let pTotal = r[12];
        let pComment = r[13];
        let pType = r[14];

        if (pClass == classNum && pTargetGroup == groupNum && pTargetRole == role) {
          if (!peerMap[pTargetSeq]) peerMap[pTargetSeq] = [];
          peerMap[pTargetSeq].push({
            evaluatorGroup: pEvalGroup,
            evaluatorRole: pEvalRole,
            grammarScore: Number(pGrammar) || 0,
            vocabScore: Number(pVocab) || 0,
            exprScore: Number(pExpr) || 0,
            totalScore: Number(pTotal) || 0,
            comment: String(pComment || ''),
            evalType: String(pType || '')
          });
        }
      });
    }

    return lyrics.map(item => {
      const studentTrans = subMap[item.seq] ? subMap[item.seq].translation : '';
      const reviews = peerMap[item.seq] || [];
      let avgScore = 0;
      if (reviews.length > 0) {
        const sum = reviews.reduce((acc, curr) => acc + (Number(curr.totalScore) || 0), 0);
        avgScore = Math.round((sum / reviews.length) * 10) / 10;
      }
      return {
        seq: item.seq,
        role: item.role,
        korean: item.korean || item.lyric,
        studentTranslation: studentTrans,
        reviews: reviews,
        avgScore: avgScore,
        reviewCount: reviews.length
      };
    });
  } catch (err) {
    return [];
  }
}
