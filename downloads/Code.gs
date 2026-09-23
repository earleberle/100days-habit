const TZ = 'Asia/Seoul';
const SETTINGS_SHEET = '설정';
const LOG_SHEET = '인증기록';

// [배포용] 프로젝트 설정 > 스크립트 속성에 SPREADSHEET_ID를 한 번만 등록하세요.
// 이후 참여자/습관/기간/앱명/카카오키/공유URL/사진폴더명/공유사용 여부는 스프레드시트 '설정'에서만 바꿉니다.
function getSpreadsheetId_() {
  const id = String(PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID') || '').trim();
  if (!id) throw new Error('배포 설정이 필요합니다. Apps Script > 프로젝트 설정 > 스크립트 속성에 SPREADSHEET_ID를 등록해 주세요.');
  return id;
}
function getSpreadsheet_() { return SpreadsheetApp.openById(getSpreadsheetId_()); }

function getAppConfig_(settings) {
  const defaults = {
    appName: '🌱 100일 습관 챌린지',
    kakaoJsKey: '',
    shareUrl: '',
    photoFolderName: '100일 습관 챌린지 인증사진',
    kakaoShareEnabled: true
  };
  if (!settings) return defaults;
  const last = Math.max(settings.getLastRow(), 2);
  const vals = settings.getRange(2, 6, last - 1, 2).getDisplayValues(); // F:G
  const map = {};
  vals.forEach(r => { const k=String(r[0]||'').trim(); if(k) map[k]=String(r[1]||'').trim(); });
  const bool = v => !['FALSE','N','0','아니오','미사용'].includes(String(v||'').trim().toUpperCase());
  return {
    appName: map['앱이름'] || defaults.appName,
    kakaoJsKey: map['카카오JS키'] || '',
    shareUrl: map['공유URL'] || '',
    photoFolderName: map['사진폴더명'] || defaults.photoFolderName,
    kakaoShareEnabled: map['카카오공유사용'] === '' || map['카카오공유사용'] == null ? true : bool(map['카카오공유사용'])
  };
}

// SPREADSHEET_ID 등록 후 1회 실행하면 설정 시트 F:G에 배포용 설정 항목을 만들어 줍니다.
function prepareDeploymentSettings() {
  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName(SETTINGS_SHEET);
  if (!sh) throw new Error('설정 시트를 찾을 수 없습니다.');
  const rows = [
    ['설정키','값'],
    ['앱이름','🌱 100일 습관 챌린지'],
    ['카카오JS키',''],
    ['공유URL',''],
    ['사진폴더명','100일 습관 챌린지 인증사진'],
    ['카카오공유사용','TRUE']
  ];
  const current = sh.getRange(1,6,rows.length,2).getDisplayValues();
  rows.forEach((r,i)=>{
    if (!String(current[i][0]||'').trim()) sh.getRange(i+1,6).setValue(r[0]);
    if (!String(current[i][1]||'').trim() && r[1] !== '') sh.getRange(i+1,7).setValue(r[1]);
  });
  sh.autoResizeColumns(6,2);
  return '배포용 설정 항목을 만들었습니다. 설정!F:G에서 값을 수정하세요.';
}

function doGet() {
  let title = '🌱 100일 습관 챌린지';
  try { const ss=getSpreadsheet_(); title=getAppConfig_(ss.getSheetByName(SETTINGS_SHEET)).appName || title; } catch(e) {}
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle(title)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

function getInitialData() {
  const ss = getSpreadsheet_();
  const settings = ss.getSheetByName(SETTINGS_SHEET);
  if (!settings) throw new Error('설정 시트를 찾을 수 없습니다.');

  const startDate = settings.getRange('B2').getValue();
  const endDate = settings.getRange('D2').getValue();
  const lastRow = Math.max(settings.getLastRow(), 5);
  const rows = settings.getRange(5, 1, lastRow - 4, 4).getValues();

  const participants = rows
    .filter(r => r[0] && r[3] !== false)
    .map(r => ({
      name: String(r[0]).trim(),
      habit1: String(r[1] || '').trim(),
      habit2: String(r[2] || '').trim()
    }));

  const today = stripTime_(new Date());
  const start = stripTime_(startDate);
  const end = stripTime_(endDate);
  const day = Math.max(1, Math.min(100, daysBetween_(start, today) + 1));

  const log = ss.getSheetByName(LOG_SHEET);
  const logRows = log && log.getLastRow() >= 2
    ? log.getRange(2, 1, log.getLastRow() - 1, 10).getValues()
    : [];

  const mindsetSheet = ss.getSheetByName('시작마음가짐');
  const mindsetCompleted = {};
  if (mindsetSheet && mindsetSheet.getLastRow() >= 2) {
    const mindsetNames = mindsetSheet.getRange(2, 1, mindsetSheet.getLastRow() - 1, 1).getDisplayValues();
    mindsetNames.forEach(r => {
      const n = String(r[0] || '').trim();
      if (n) mindsetCompleted[n] = true;
    });
  }
  const mindsetMissing = participants.map(p => p.name).filter(name => !mindsetCompleted[name]);

  const passMap = {};
  const passSheet = ss.getSheetByName('패스권');
  if (passSheet && passSheet.getLastRow() >= 2) {
    const passRows = passSheet.getRange(2,1,passSheet.getLastRow()-1,2).getValues();
    passRows.forEach(r => {
      const n=String(r[1]||'').trim(), d=formatDate_(r[0]);
      if(!n || !d) return;
      if(!passMap[n]) passMap[n]={};
      passMap[n][d]=true;
    });
  }

  const histories = {};
  participants.forEach(p => {
    histories[p.name] = buildParticipantHistoryFromRows_(p, start, today, logRows, passMap[p.name] || {});
  });

  const appConfig = getAppConfig_(settings);
  const deployedUrl = ScriptApp.getService().getUrl() || '';

  return {
    appName: appConfig.appName,
    kakaoJsKey: appConfig.kakaoJsKey,
    shareUrl: appConfig.shareUrl || deployedUrl,
    kakaoShareEnabled: appConfig.kakaoShareEnabled,
    histories: histories,
    mindsetNotice: {
      total: participants.length,
      completedCount: participants.length - mindsetMissing.length,
      missing: mindsetMissing,
      allDone: mindsetMissing.length === 0
    },
    startDate: formatDate_(start),
    endDate: formatDate_(end),
    today: formatDate_(today),
    day: day,
    participants: participants,
    dashboard: buildDashboardFromRows_(participants, today, logRows),
    recent: getRecentRecordsFromRows_(logRows, 30)
  };
}

function buildParticipantHistoryFromRows_(participant, start, today, rows, myPasses) {
  const byDate = {};
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    if (String(r[2] || '').trim() !== participant.name || r[5] !== true) continue;
    const d = formatDate_(r[0]);
    if (!byDate[d]) byDate[d] = {};
    byDate[d][Number(r[3])] = {
      habitName: String(r[4] || ''),
      caption: String(r[6] || ''),
      photoUrl: String(r[7] || ''),
      photoViewUrl: toPhotoViewUrl_(String(r[7] || '')),
      certifiedAt: formatDateTime_(r[8]),
      late: String(r[9] || '') === 'Y'
    };
  }

  const days = [];
  let habit1Count=0, habit2Count=0, bothCount=0, currentStreak=0, bestStreak=0;
  for (let i=0;i<100;i++) {
    const d=new Date(start); d.setDate(d.getDate()+i);
    const key=formatDate_(d), rec=byDate[key]||{};
    const h1=rec[1]||null, h2=rec[2]||null;
    const count=(h1?1:0)+(h2?1:0), future=d>today, pass=!!myPasses[key];
    if(h1) habit1Count++;
    if(h2) habit2Count++;
    if(count===2) bothCount++;
    if(!future && count===2){ currentStreak++; bestStreak=Math.max(bestStreak,currentStreak); }
    else if(!future && pass){ /* pass bridges streak */ }
    else if(!future){ currentStreak=0; }
    days.push({day:i+1,date:key,future,pass,habit1:h1,habit2:h2,count});
  }
  return {
    name:participant.name, habit1:participant.habit1, habit2:participant.habit2,
    habit1Count, habit2Count, bothCount,
    totalRate:Math.round(((habit1Count+habit2Count)/200)*100),
    bestStreak, days
  };
}


function saveCertification(payload) {
  if (!payload) throw new Error('저장할 내용이 없습니다.');

  const name = clean_(payload.name);
  const targetDateText = clean_(payload.targetDate);
  const habits = Array.isArray(payload.habits) ? payload.habits : [];

  if (!name) throw new Error('참여자를 선택해 주세요.');
  if (!targetDateText) throw new Error('인증 날짜를 선택해 주세요.');
  if (!habits.length) throw new Error('인증할 습관을 하나 이상 선택해 주세요.');

  const ss = getSpreadsheet_();
  const settings = ss.getSheetByName(SETTINGS_SHEET);
  const log = ss.getSheetByName(LOG_SHEET);
  if (!settings || !log) throw new Error('필요한 시트를 찾을 수 없습니다.');

  const participant = findParticipant_(settings, name);
  if (!participant) throw new Error('설정 시트에서 참여자를 찾을 수 없습니다.');

  const start = stripTime_(settings.getRange('B2').getValue());
  const end = stripTime_(settings.getRange('D2').getValue());
  const targetDate = parseDate_(targetDateText);
  const today = stripTime_(new Date());

  if (targetDate < start || targetDate > end) throw new Error('100일 챌린지 기간 안의 날짜만 인증할 수 있습니다.');
  if (targetDate > today) throw new Error('미래 날짜는 인증할 수 없습니다.');

  const day = daysBetween_(start, targetDate) + 1;
  const now = new Date();
  const late = targetDate.getTime() !== today.getTime();

  const existing = getExistingMap_(log, name, targetDateText);
  const results = [];

  habits.forEach(item => {
    const habitNo = Number(item.habitNo);
    if (![1, 2].includes(habitNo)) return;

    const habitName = habitNo === 1 ? participant.habit1 : participant.habit2;
    if (!habitName) throw new Error(name + '님의 습관 ' + habitNo + '가 아직 설정되지 않았습니다.');

    const old = existing[habitNo] || null;
    const caption = clean_(item.caption);
    const hasNewPhoto = !!(item.photo && item.photo.data);

    // 화면에서 기존 인증 체크박스가 자동으로 체크되어 있어도,
    // 내용이 바뀌지 않은 습관은 기존 행/기존 인증시각을 그대로 둡니다.
    const contentChanged = !old || caption !== old.caption || hasNewPhoto;

    if (!contentChanged) {
      results.push({ habitNo, habitName, unchanged: true });
      return;
    }

    let photoUrl = old ? old.photoUrl : '';
    if (hasNewPhoto) {
      photoUrl = savePhoto_(item.photo, name, targetDateText, habitNo);
    }

    const certifiedAt = now;
    const rowValues = [
      targetDate,
      day,
      name,
      habitNo,
      habitName,
      true,
      caption,
      photoUrl,
      certifiedAt,
      late ? 'Y' : ''
    ];

    if (old) {
      log.getRange(old.row, 1, 1, 10).setValues([rowValues]);
    } else {
      log.appendRow(rowValues);
    }

    results.push({ habitNo, habitName, unchanged: false });
  });

  formatLogColumns_(log);

  return {
    ok: true,
    message: (results.filter(r => !r.unchanged).length || results.length) + '개 습관 인증을 저장했어요!',
    dashboard: buildDashboardFromSettings_(),
    dayDetail: getDayDetail(name, targetDateText)
  };
}

function getDayDetail(name, dateText) {
  name = clean_(name);
  dateText = clean_(dateText);
  if (!name || !dateText) return { records: [] };

  const ss = getSpreadsheet_();
  const log = ss.getSheetByName(LOG_SHEET);
  if (!log || log.getLastRow() < 2) return { records: [] };

  const values = log.getRange(2, 1, log.getLastRow() - 1, 10).getValues();
  const records = values
    .filter(r => formatDate_(r[0]) === dateText && String(r[2]).trim() === name && r[5] === true)
    .map(r => ({
      habitNo: Number(r[3]),
      habitName: String(r[4] || ''),
      caption: String(r[6] || ''),
      photoUrl: String(r[7] || ''),
      photoViewUrl: toPhotoViewUrl_(String(r[7] || '')),
      certifiedAt: formatDateTime_(r[8]),
      late: String(r[9] || '') === 'Y'
    }))
    .sort((a, b) => a.habitNo - b.habitNo);

  return { records };
}


function getFinaleState() {
  const ss=getSpreadsheet_();
  const settings=ss.getSheetByName(SETTINGS_SHEET);
  const participants=settings.getRange(5,1,Math.max(1,settings.getLastRow()-4),4).getValues()
    .filter(r=>r[0]&&r[3]!==false).map(r=>String(r[0]).trim());
  const end=stripTime_(settings.getRange('D2').getValue());
  const today=stripTime_(new Date());
  const log=ss.getSheetByName(LOG_SHEET);
  const rows=log&&log.getLastRow()>=2?log.getRange(2,1,log.getLastRow()-1,10).getValues():[];
  const endText=formatDate_(end);
  const finalRows=rows.filter(r=>formatDate_(r[0])===endText&&r[5]===true);
  const completed=participants.filter(name=>{
    const mine=finalRows.filter(r=>String(r[2]).trim()===name);
    return mine.some(r=>Number(r[3])===1)&&mine.some(r=>Number(r[3])===2);
  });
  const total=rows.filter(r=>r[5]===true).length;
  return {
    eligible: today>=end,
    complete: today>=end && completed.length===participants.length,
    completedCount: completed.length,
    participantCount: participants.length,
    finalCertCount: finalRows.length,
    requiredFinalCertCount: participants.length*2,
    totalCertifications: total,
    participants: participants,
    startDate: formatDate_(settings.getRange('B2').getValue()),
    endDate: endText
  };
}

function getWeeklyOverview(anchorDateText) {
  const ss = getSpreadsheet_();
  const settings = ss.getSheetByName(SETTINGS_SHEET);
  const log = ss.getSheetByName(LOG_SHEET);
  const participants = settings.getRange(5, 1, Math.max(1, settings.getLastRow() - 4), 4).getValues()
    .filter(r => r[0] && r[3] !== false)
    .map(r => ({name:String(r[0]).trim(), habit1:String(r[1]||'').trim(), habit2:String(r[2]||'').trim()}));

  const start = stripTime_(settings.getRange('B2').getValue());
  const end = stripTime_(settings.getRange('D2').getValue());
  const today = stripTime_(new Date());
  let anchor = anchorDateText ? parseDate_(anchorDateText) : today;
  const dow = anchor.getDay();
  const monday = new Date(anchor);
  monday.setDate(anchor.getDate() - (dow === 0 ? 6 : dow - 1));
  const sunday = new Date(monday); sunday.setDate(monday.getDate() + 6);

  const rows = log && log.getLastRow() >= 2 ? log.getRange(2,1,log.getLastRow()-1,10).getValues() : [];
  const passMap = getPassMap_();

  const dates = [];
  for (let i=0;i<7;i++) {
    const d = new Date(monday); d.setDate(monday.getDate()+i);
    dates.push({
      date: formatDate_(d),
      label: ['월','화','수','목','금','토','일'][i],
      day: daysBetween_(start,d)+1,
      beforeStart: d < start,
      afterEnd: d > end,
      future: d > today
    });
  }

  const people = participants.map(p => {
    const cells = dates.map(info => {
      const pass = !!(passMap[p.name] && passMap[p.name][info.date]);
      const mine = rows.filter(r => String(r[2]).trim()===p.name && formatDate_(r[0])===info.date && r[5]===true);
      const h1 = mine.find(r=>Number(r[3])===1);
      const h2 = mine.find(r=>Number(r[3])===2);
      return {
        date: info.date,
        count: (h1?1:0)+(h2?1:0),
        pass,
        future: info.future,
        beforeStart: info.beforeStart,
        afterEnd: info.afterEnd
      };
    });
    const passCount = cells.filter(c=>c.pass).length;
    const h1Count = dates.filter(info => rows.some(r=>String(r[2]).trim()===p.name && formatDate_(r[0])===info.date && Number(r[3])===1 && r[5]===true)).length;
    const h2Count = dates.filter(info => rows.some(r=>String(r[2]).trim()===p.name && formatDate_(r[0])===info.date && Number(r[3])===2 && r[5]===true)).length;
    return {name:p.name, habit1:p.habit1, habit2:p.habit2, cells, passCount, h1Count, h2Count};
  });

  return {
    weekStart: formatDate_(monday),
    weekEnd: formatDate_(sunday),
    anchorDate: formatDate_(anchor),
    dates, people,
    canNext: sunday < end && monday <= today
  };
}

function useWeeklyPass(name, dateText) {
  name = clean_(name); dateText = clean_(dateText);
  if (!name || !dateText) throw new Error('참여자와 날짜를 확인해 주세요.');

  const ss = getSpreadsheet_();
  const settings = ss.getSheetByName(SETTINGS_SHEET);
  const start = stripTime_(settings.getRange('B2').getValue());
  const end = stripTime_(settings.getRange('D2').getValue());
  const date = parseDate_(dateText);
  const today = stripTime_(new Date());
  if (date < start || date > end || date > today) throw new Error('패스권은 챌린지 기간의 오늘 또는 지난 날짜에만 사용할 수 있어요.');

  let sheet = ss.getSheetByName('패스권');
  if (!sheet) {
    sheet = ss.insertSheet('패스권');
    sheet.appendRow(['날짜','참여자','사용일시']);
    sheet.setFrozenRows(1);
  }

  const dow = date.getDay();
  const monday = new Date(date); monday.setDate(date.getDate() - (dow===0?6:dow-1));
  const sunday = new Date(monday); sunday.setDate(monday.getDate()+6);

  const rows = sheet.getLastRow()>=2 ? sheet.getRange(2,1,sheet.getLastRow()-1,3).getValues() : [];
  const existingIndex = rows.findIndex(r => formatDate_(r[0])===dateText && String(r[1]).trim()===name);
  if (existingIndex >= 0) {
    sheet.deleteRow(existingIndex+2);
    return {ok:true, used:false, message:'패스권 사용을 취소했어요.'};
  }

  const weekly = rows.filter(r => {
    const d=stripTime_(r[0]);
    return String(r[1]).trim()===name && d>=monday && d<=sunday;
  });
  if (weekly.length >= 2) throw new Error('이번 주 패스권 2회를 이미 모두 사용했어요.');

  sheet.appendRow([date,name,new Date()]);
  sheet.getRange(2,1,Math.max(1,sheet.getLastRow()-1),1).setNumberFormat('yyyy.mm.dd');
  sheet.getRange(2,3,Math.max(1,sheet.getLastRow()-1),1).setNumberFormat('yyyy.mm.dd hh:mm');
  return {ok:true, used:true, message:'패스권을 사용했어요. 이번 주 '+(weekly.length+1)+'/2회 사용'};
}

function getPassMap_() {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName('패스권');
  const map = {};
  if (!sheet || sheet.getLastRow()<2) return map;
  sheet.getRange(2,1,sheet.getLastRow()-1,3).getValues().forEach(r=>{
    const name=String(r[1]||'').trim(), d=formatDate_(r[0]);
    if (!name || !d) return;
    if (!map[name]) map[name]={};
    map[name][d]=true;
  });
  return map;
}

function getParticipantHistory(name) {
  name = clean_(name);
  if (!name) throw new Error('참여자를 선택해 주세요.');

  const ss = getSpreadsheet_();
  const settings = ss.getSheetByName(SETTINGS_SHEET);
  const log = ss.getSheetByName(LOG_SHEET);
  const participant = findParticipant_(settings, name);
  if (!participant) throw new Error('참여자를 찾을 수 없습니다.');

  const start = stripTime_(settings.getRange('B2').getValue());
  const today = stripTime_(new Date());
  const rows = log && log.getLastRow() >= 2
    ? log.getRange(2, 1, log.getLastRow() - 1, 10).getValues()
    : [];

  const passMap = getPassMap_();
  const myPasses = passMap[name] || {};
  const byDate = {};
  rows.filter(r => String(r[2]).trim() === name && r[5] === true).forEach(r => {
    const d = formatDate_(r[0]);
    if (!byDate[d]) byDate[d] = {};
    byDate[d][Number(r[3])] = {
      habitName: String(r[4] || ''),
      caption: String(r[6] || ''),
      photoUrl: String(r[7] || ''),
      photoViewUrl: toPhotoViewUrl_(String(r[7] || '')),
      certifiedAt: formatDateTime_(r[8]),
      late: String(r[9] || '') === 'Y'
    };
  });

  const days = [];
  for (let i = 0; i < 100; i++) {
    const d = new Date(start);
    d.setDate(d.getDate() + i);
    const key = formatDate_(d);
    const rec = byDate[key] || {};
    days.push({
      day: i + 1,
      date: key,
      future: d > today,
      pass: !!myPasses[key],
      habit1: rec[1] || null,
      habit2: rec[2] || null,
      count: (rec[1] ? 1 : 0) + (rec[2] ? 1 : 0)
    });
  }

  const habit1Count = days.filter(d => d.habit1).length;
  const habit2Count = days.filter(d => d.habit2).length;
  const bothCount = days.filter(d => d.count === 2).length;

  let currentStreak = 0;
  let bestStreak = 0;
  days.forEach(d => {
    if (!d.future && d.count === 2) {
      currentStreak++;
      bestStreak = Math.max(bestStreak, currentStreak);
    } else if (!d.future && d.pass) {
      // 패스일은 성공일 수에는 포함하지 않지만 연속 기록을 끊지 않습니다.
    } else if (!d.future) {
      currentStreak = 0;
    }
  });

  return {
    name,
    habit1: participant.habit1,
    habit2: participant.habit2,
    habit1Count,
    habit2Count,
    bothCount,
    totalRate: Math.round(((habit1Count + habit2Count) / 200) * 100),
    bestStreak,
    days
  };
}

function saveStartMindset(payload) {
  const name = clean_(payload && payload.name);
  const password = clean_(payload && payload.password);
  if (!name) throw new Error('참여자를 선택해 주세요.');
  if (!password || password.length < 4) throw new Error('비밀번호는 4자 이상으로 설정해 주세요.');

  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName('시작마음가짐');
  if (!sheet) {
    sheet = ss.insertSheet('시작마음가짐');
    sheet.appendRow(['참여자','작성일시','100일 동안 만들고 싶은 변화','이 습관을 정한 이유','100일 뒤 나에게 한마디','비밀번호']);
    sheet.setFrozenRows(1);
  } else if (sheet.getLastColumn() < 6) {
    sheet.getRange(1,6).setValue('비밀번호');
  }

  const lastRow=sheet.getLastRow();
  let existing=null, rowNo=0;
  if(lastRow>=2){
    const vals=sheet.getRange(2,1,lastRow-1,6).getValues();
    const idx=vals.findIndex(r=>String(r[0]).trim()===name);
    if(idx>=0){existing=vals[idx];rowNo=idx+2;}
  }
  if(existing && existing[5] && String(existing[5])!==password) throw new Error('기존 비밀번호가 맞지 않아요.');

  const values=[name,new Date(),clean_(payload.change),clean_(payload.reason),clean_(payload.message),password];
  if(rowNo) sheet.getRange(rowNo,1,1,6).setValues([values]);
  else sheet.appendRow(values);
  return {ok:true,message:rowNo?'시작하는 마음가짐을 수정 저장했어요!':'시작하는 마음가짐을 잠금 저장했어요!'};
}


function getMindsetNoticeStatus() {
  const ss = getSpreadsheet_();
  const settings = ss.getSheetByName(SETTINGS_SHEET);
  const mindset = ss.getSheetByName('시작마음가짐');

  const lastRow = Math.max(settings.getLastRow(), 5);
  const participants = settings.getRange(5, 1, lastRow - 4, 4).getValues()
    .filter(r => r[0] && r[3] !== false)
    .map(r => String(r[0]).trim());
  const completed = {};

  if (mindset && mindset.getLastRow() >= 2) {
    const names = mindset.getRange(2, 1, mindset.getLastRow() - 1, 1).getDisplayValues();
    for (let i = 0; i < names.length; i++) {
      const n = String(names[i][0] || '').trim();
      if (n) completed[n] = true;
    }
  }

  const missing = participants.filter(name => !completed[name]);
  return {
    total: participants.length,
    completedCount: participants.length - missing.length,
    missing: missing,
    allDone: missing.length === 0
  };
}

function getStartMindsetStatus(name) {
  name=clean_(name);
  const ss=getSpreadsheet_(), sheet=ss.getSheetByName('시작마음가짐');
  if(!sheet||sheet.getLastRow()<2) return {exists:false,locked:false};
  const width=Math.max(6,sheet.getLastColumn());
  const vals=sheet.getRange(2,1,sheet.getLastRow()-1,width).getValues();
  const row=vals.find(r=>String(r[0]).trim()===name);
  return row ? {exists:true,locked:!!row[5],writtenAt:formatDateTime_(row[1])} : {exists:false,locked:false};
}

function unlockStartMindset(name,password) {
  name=clean_(name); password=clean_(password);
  const ss=getSpreadsheet_(), sheet=ss.getSheetByName('시작마음가짐');
  if(!sheet||sheet.getLastRow()<2) throw new Error('저장된 시작 마음가짐이 없어요.');
  const width=Math.max(6,sheet.getLastColumn());
  const vals=sheet.getRange(2,1,sheet.getLastRow()-1,width).getValues();
  const row=vals.find(r=>String(r[0]).trim()===name);
  if(!row) throw new Error('저장된 시작 마음가짐이 없어요.');
  if(row[5] && String(row[5])!==password) throw new Error('비밀번호가 맞지 않아요.');
  return {writtenAt:formatDateTime_(row[1]),change:String(row[2]||''),reason:String(row[3]||''),message:String(row[4]||'')};
}



function saveReflection(payload) {
  const name = clean_(payload && payload.name);
  if (!name) throw new Error('참여자를 선택해 주세요.');

  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName('회고응답');
  if (!sheet) {
    sheet = ss.insertSheet('회고응답');
    sheet.appendRow(['참여자', '작성일시', '가장 잘 지킨 습관', '가장 힘들었던 순간', '100일 동안 달라진 점', '앞으로도 계속 가져갈 습관']);
    sheet.setFrozenRows(1);
  }

  const values = [
    name,
    new Date(),
    clean_(payload.bestHabit),
    clean_(payload.hardest),
    clean_(payload.changed),
    clean_(payload.keepHabit)
  ];

  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const names = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues().flat();
    const idx = names.findIndex(v => v.trim() === name);
    if (idx >= 0) {
      sheet.getRange(idx + 2, 1, 1, 6).setValues([values]);
      return { ok: true, message: '100일 후기를 수정 저장했어요!' };
    }
  }
  sheet.appendRow(values);
  return { ok: true, message: '100일 후기를 저장했어요!' };
}

function getReflection(name) {
  name = clean_(name);
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName('회고응답');
  if (!sheet || sheet.getLastRow() < 2) return null;
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  const row = values.find(r => String(r[0]).trim() === name);
  if (!row) return null;
  return {
    bestHabit: String(row[2] || ''),
    hardest: String(row[3] || ''),
    changed: String(row[4] || ''),
    keepHabit: String(row[5] || '')
  };
}

function deleteCertification(name, dateText, habitNo) {
  name = clean_(name);
  dateText = clean_(dateText);
  habitNo = Number(habitNo);

  const ss = getSpreadsheet_();
  const log = ss.getSheetByName(LOG_SHEET);
  if (!log || log.getLastRow() < 2) return { ok: true };

  const values = log.getRange(2, 1, log.getLastRow() - 1, 10).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    const r = values[i];
    if (formatDate_(r[0]) === dateText && String(r[2]).trim() === name && Number(r[3]) === habitNo) {
      log.deleteRow(i + 2);
      break;
    }
  }
  return { ok: true, dashboard: buildDashboardFromSettings_(), dayDetail: getDayDetail(name, dateText) };
}

function buildDashboardFromSettings_() {
  const ss = getSpreadsheet_();
  const settings = ss.getSheetByName(SETTINGS_SHEET);
  const rows = settings.getRange(5, 1, Math.max(1, settings.getLastRow() - 4), 4).getValues();
  const participants = rows
    .filter(r => r[0] && r[3] !== false)
    .map(r => ({ name: String(r[0]).trim(), habit1: String(r[1] || '').trim(), habit2: String(r[2] || '').trim() }));
  return buildDashboard_(participants, stripTime_(new Date()));
}

function buildDashboard_(participants, today) {
  const ss = getSpreadsheet_();
  const log = ss.getSheetByName(LOG_SHEET);
  const rows = log && log.getLastRow() >= 2
    ? log.getRange(2, 1, log.getLastRow() - 1, 10).getValues()
    : [];
  return buildDashboardFromRows_(participants, today, rows);
}

function buildDashboardFromRows_(participants, today, rows) {
  const todayText = formatDate_(today);
  const stats = {};
  participants.forEach(p => stats[p.name] = {h1:0,h2:0,t1:'',t2:''});

  rows.forEach(r => {
    if (r[5] !== true) return;
    const name = String(r[2] || '').trim();
    const s = stats[name];
    if (!s) return;
    const no = Number(r[3]);
    if (no === 1) s.h1++;
    if (no === 2) s.h2++;
    if (formatDate_(r[0]) === todayText) {
      if (no === 1) s.t1 = formatTime_(r[8]);
      if (no === 2) s.t2 = formatTime_(r[8]);
    }
  });

  return participants.map(p => ({
    name:p.name, habit1:p.habit1, habit2:p.habit2,
    habit1Count:stats[p.name].h1, habit2Count:stats[p.name].h2,
    today1:stats[p.name].t1, today2:stats[p.name].t2
  }));
}

function getRecentRecords_(limit) {
  const ss = getSpreadsheet_();
  const log = ss.getSheetByName(LOG_SHEET);
  const rows = log && log.getLastRow() >= 2
    ? log.getRange(2, 1, log.getLastRow() - 1, 10).getValues()
    : [];
  return getRecentRecordsFromRows_(rows, limit);
}

function getRecentRecordsFromRows_(rows, limit) {
  return rows
    .filter(r => r[5] === true)
    .sort((a, b) => new Date(b[8]).getTime() - new Date(a[8]).getTime())
    .slice(0, limit)
    .map(r => ({
      date: formatDate_(r[0]), day: Number(r[1]), name: String(r[2] || ''),
      habitNo: Number(r[3]), habitName: String(r[4] || ''), caption: String(r[6] || ''),
      photoViewUrl: toPhotoViewUrl_(String(r[7] || '')),
      certifiedAt: formatDateTime_(r[8]), late: String(r[9] || '') === 'Y'
    }));
}

function findParticipant_(settings, name) {
  const lastRow = settings.getLastRow();
  if (lastRow < 5) return null;
  const rows = settings.getRange(5, 1, lastRow - 4, 4).getValues();
  const row = rows.find(r => String(r[0]).trim() === name && r[3] !== false);
  return row ? { name, habit1: String(row[1] || '').trim(), habit2: String(row[2] || '').trim() } : null;
}

function getExistingMap_(log, name, dateText) {
  const map = {};
  if (log.getLastRow() < 2) return map;
  const values = log.getRange(2, 1, log.getLastRow() - 1, 10).getValues();
  values.forEach((r, i) => {
    if (formatDate_(r[0]) === dateText && String(r[2]).trim() === name) {
      map[Number(r[3])] = {
        row: i + 2,
        caption: String(r[6] || ''),
        photoUrl: String(r[7] || ''),
        certifiedAt: r[8],
        late: String(r[9] || '')
      };
    }
  });
  return map;
}

function savePhoto_(photo, name, dateText, habitNo) {
  const folder = getOrCreatePhotoFolder_();
  const bytes = Utilities.base64Decode(photo.data);
  const ext = mimeToExt_(photo.mimeType);
  const safeName = name.replace(/[\\/:*?"<>|]/g, '_');
  const filename = dateText.replace(/-/g, '') + '_' + safeName + '_습관' + habitNo + '_' + Utilities.getUuid().slice(0, 8) + '.' + ext;
  const blob = Utilities.newBlob(bytes, photo.mimeType || 'image/jpeg', filename);
  const file = folder.createFile(blob);

  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {
    // 조직 정책상 링크 공유가 막혀 있어도 파일 저장 자체는 유지합니다.
  }
  return file.getUrl();
}

function getOrCreatePhotoFolder_() {
  const props = PropertiesService.getScriptProperties();
  const savedId = props.getProperty('PHOTO_FOLDER_ID');
  if (savedId) {
    try { return DriveApp.getFolderById(savedId); } catch (e) {}
  }
  const ss = getSpreadsheet_();
  const cfg = getAppConfig_(ss.getSheetByName(SETTINGS_SHEET));
  const folderName = cfg.photoFolderName || '100일 습관 챌린지 인증사진';
  const folders = DriveApp.getFoldersByName(folderName);
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
  props.setProperty('PHOTO_FOLDER_ID', folder.getId());
  return folder;
}

function toPhotoViewUrl_(url) {
  if (!url) return '';
  const m = url.match(/[-\w]{25,}/);
  return m ? 'https://drive.google.com/thumbnail?id=' + m[0] + '&sz=w1200' : url;
}

function formatLogColumns_(sheet) {
  if (sheet.getLastRow() < 2) return;
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).setNumberFormat('yyyy.mm.dd');
  sheet.getRange(2, 9, sheet.getLastRow() - 1, 1).setNumberFormat('yyyy.mm.dd hh:mm');
}

function parseDate_(text) {
  const parts = String(text).split('-').map(Number);
  if (parts.length !== 3) throw new Error('날짜 형식이 올바르지 않습니다.');
  return new Date(parts[0], parts[1] - 1, parts[2]);
}

function stripTime_(date) {
  const d = new Date(date);
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function daysBetween_(a, b) {
  const aa = Date.UTC(a.getFullYear(), a.getMonth(), a.getDate());
  const bb = Date.UTC(b.getFullYear(), b.getMonth(), b.getDate());
  return Math.floor((bb - aa) / 86400000);
}

function formatDate_(date) {
  if (!date || Object.prototype.toString.call(date) !== '[object Date]' || isNaN(date)) return '';
  return Utilities.formatDate(date, TZ, 'yyyy-MM-dd');
}

function formatTime_(date) {
  if (!date || Object.prototype.toString.call(date) !== '[object Date]' || isNaN(date)) return '';
  return Utilities.formatDate(date, TZ, 'HH:mm');
}

function formatDateTime_(date) {
  if (!date || Object.prototype.toString.call(date) !== '[object Date]' || isNaN(date)) return '';
  return Utilities.formatDate(date, TZ, 'yyyy.MM.dd HH:mm');
}

function clean_(value) {
  return String(value == null ? '' : value).trim().slice(0, 1000);
}

function mimeToExt_(mime) {
  const map = {
    'image/jpeg': 'jpg',
    'image/png': 'png',
    'image/webp': 'webp',
    'image/gif': 'gif',
    'image/heic': 'heic',
    'image/heif': 'heif'
  };
  return map[mime] || 'jpg';
}
