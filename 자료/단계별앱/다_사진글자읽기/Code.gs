/**
 * [다] 사진 속 글자 읽기 (원본 사진 + 읽은 글)
 * AI·디지털 기반 수업과 평가 연수 · 단계별 앱
 *
 * 학생이 공책을 찍어 올리면 사진은 드라이브에 저장되고, 드라이브 OCR 이 사진 속 글자를 읽어 화면에 돌려줍니다.
 * 학생은 읽은 글을 원본과 견줘 고친 뒤 냅니다. [응답] 시트에는 학생이 확인한 글, 드라이브가 처음 읽은 글, 사진 링크가 남습니다.
 * API 키도 결제도 필요 없습니다.
 *
 * 1. 구글 시트 → 확장 프로그램 → Apps Script → 이 글 전체를 Code.gs 에 붙여 넣기 (있던 글은 지우고)
 * 2. 파일 옆 + → HTML → 이름 Index → Index.html 내용 붙여 넣기
 * 3. 위쪽 함수 칸에서 setup 을 고르고 ▷실행 → 권한 허용. 이어서 testDriveOcr 을 실행해 「연결이 됩니다」를 확인
 * 4. 배포 → 새 배포 → 유형 「웹 앱」 → 실행 사용자 「나」 · 액세스 권한 「모든 사용자」 → 배포 → /exec 주소 복사
 * 코드를 고친 뒤에는 배포 → 배포 관리 → 연필 → 버전 「새 버전」 → 배포. 「새 배포」를 누르면 주소가 바뀝니다.
 */

/* ===== 바꿀 곳은 여기 ===== */
const APP = {
  NAME: '사진 속 글자 읽기',
  CLASS_CODE: '',       // 우리 반 코드 (예: '4817'). 비워 두면 묻지 않습니다
  MAX_TEXT: 5000,
  MAX_PHOTOS: 3,        // 한 번에 올리는 사진 수 (쪽 수)
  MAX_FILE_MB: 10,
  PHOTO_LABEL: '사진'
};

const HEAD = ['낸 시각', '학년', '반', '번호', '이름', '과제', '학생이 확인한 글', '드라이브가 읽은 글', '사진1', '사진2', '사진3'];

const SAMPLE_TASKS = [
  ['연소의 조건 서술형', '공책에 답을 쓰고 찍어 올려요. 촛불을 컵으로 덮으면 꺼지는 까닭을 연소의 조건과 이어서 설명해요.', 'O'],
  ['전기 회로 관찰 기록', '공책에 그린 회로와 설명을 찍어 올려요. 전구에 불이 켜지는 회로의 특징을 적어요.', 'O']
];

/* ================= 공통: 시트 만들기 · 메뉴 · 학생 확인 ================= */

const SHEET_TASK = '과제';
const SHEET_ANSWER = '응답';

/** 시트를 열 때 위쪽 메뉴를 붙입니다 */
function onOpen() {
  const menu = SpreadsheetApp.getUi().createMenu('🧩 ' + APP.NAME)
    .addItem('처음 설정 (시트 · 폴더 만들기)', 'setup')
    .addItem('학생 주소 확인', 'showUrl');
  if (typeof MENU_EXTRA !== 'undefined') {
    menu.addSeparator();
    MENU_EXTRA.forEach(function (m) { menu.addItem(m[0], m[1]); });
  }
  menu.addSeparator()
    .addItem('새 학기: 응답 시트 보관하고 새로 만들기', 'archiveAnswers')
    .addToUi();
}

/** 처음 한 번. 편집기에서 setup 을 골라 실행하면 권한을 묻고 시트를 만듭니다 */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let task = ss.getSheetByName(SHEET_TASK);
  if (!task) {
    task = ss.insertSheet(SHEET_TASK);
    const head = ['과제 이름', '학생에게 보일 질문', '열기 (O / X)'];
    task.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#ffcd29');
    task.getRange(2, 1, SAMPLE_TASKS.length, head.length).setValues(SAMPLE_TASKS);
    task.setFrozenRows(1);
    task.setColumnWidth(1, 180);
    task.setColumnWidth(2, 520);
  }
  let ans = ss.getSheetByName(SHEET_ANSWER);
  if (!ans) ans = makeAnswerSheet_(ss, SHEET_ANSWER);
  if (typeof folder_ === 'function') folder_();

  // 새 시트에 딸려 온 빈 「시트1」은 치웁니다
  ['시트1', 'Sheet1'].forEach(function (n) {
    const s = ss.getSheetByName(n);
    if (s && s.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s);
  });
  say_('처음 설정을 마쳤습니다. [과제] 시트에서 질문을 고친 뒤 배포하세요.');
}

function makeAnswerSheet_(ss, name) {
  const sh = ss.insertSheet(name);
  sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold').setBackground('#dbe8fb');
  sh.setFrozenRows(1);
  return sh;
}

/** 배포한 주소를 보여 줍니다. 배포 전이면 배포하라고 알려 줍니다 */
function showUrl() {
  let url = '';
  try { url = ScriptApp.getService().getUrl() || ''; } catch (e) { url = ''; }
  say_(url
    ? '학생 주소 (끝이 /exec)\n' + url + '\n\n끝이 /dev 인 주소는 선생님만 열립니다.'
    : '아직 배포하지 않았습니다. 편집기 오른쪽 위 [배포] → [새 배포] → 유형 「웹 앱」.');
}

/** 새 학기. 지금 응답 시트의 이름을 「응답_보관_날짜」로 바꾸고 빈 응답 시트를 새로 만듭니다. 지우지 않습니다 */
function archiveAnswers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const old = ss.getSheetByName(SHEET_ANSWER);
  if (!old) { setup(); return; }
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmm');
  old.setName(SHEET_ANSWER + '_보관_' + stamp);
  makeAnswerSheet_(ss, SHEET_ANSWER);
  say_('[' + old.getName() + '] 로 보관하고 빈 [' + SHEET_ANSWER + '] 시트를 만들었습니다. 드라이브 파일은 그대로 있습니다.');
}

/** 학생 화면이 처음 부르는 함수. 열린 과제만 돌려줍니다 */
function getTasks() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TASK);
  const rows = sh && sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues() : [];
  const tasks = rows
    .filter(function (r) { return String(r[0]).trim() && String(r[2]).trim().toUpperCase() !== 'X'; })
    .map(function (r) { return { name: String(r[0]).trim(), question: String(r[1]).trim() }; });
  return { ok: true, app: APP.NAME, needCode: !!String(APP.CLASS_CODE || '').trim(), tasks: tasks };
}

/**
 * 학생이 보낸 학년 · 반 · 번호 · 이름을 서버에서 다시 확인합니다.
 * 화면에서 막았어도 여기서 한 번 더 봅니다. 화면은 누구나 고칠 수 있습니다.
 */
function who_(p) {
  p = p || {};
  const code = String(APP.CLASS_CODE || '').trim();
  if (code && String(p.code || '').trim() !== code) throw oops_('우리 반 코드가 맞지 않아요.');
  const grade = parseInt(p.grade, 10), cls = parseInt(p.cls, 10), num = parseInt(p.num, 10);
  const name = clean_(p.name, 20);
  if (!(grade >= 1 && grade <= 6)) throw oops_('학년을 골라요.');
  if (!(cls >= 0 && cls <= 30)) throw oops_('반을 숫자로 적어요.');
  if (!(num >= 1 && num <= 60)) throw oops_('번호를 숫자로 적어요.');
  if (!name) throw oops_('이름을 적어요.');
  const task = clean_(p.task, 60);
  const open = getTasks().tasks.some(function (t) { return t.name === task; });
  if (!open) throw oops_('닫혔거나 없는 과제예요. 화면을 새로 고쳐요.');
  return { grade: grade, cls: cls, num: num, name: name, task: task };
}

/** 앞뒤 빈칸을 자르고 길이를 넘기지 않게 합니다 */
function clean_(s, max) {
  return String(s === undefined || s === null ? '' : s).replace(/\r\n/g, '\n').trim().slice(0, max);
}

/** =, +, -, @ 로 시작하는 글은 시트가 수식으로 읽습니다. 앞에 ' 를 붙여 글자로 둡니다 */
function cell_(s) {
  s = String(s === undefined || s === null ? '' : s);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function link_(url, label) {
  return url ? '=HYPERLINK("' + String(url).replace(/"/g, '') + '","' + label + '")' : '';
}

/** 응답 시트에 한 줄을 붙입니다. appendRow 는 여러 학생이 동시에 내도 줄이 겹치지 않습니다 */
function addRow_(row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_ANSWER) || makeAnswerSheet_(ss, SHEET_ANSWER);
  sh.appendRow(row);
}

function whoCells_(w) {
  return [new Date(), w.grade, w.cls, w.num, cell_(w.name), cell_(w.task)];
}

function hhmm_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm');
}

/** 시트에서 누른 메뉴면 창으로, 편집기에서 실행했으면 실행 기록으로 알립니다 */
function say_(msg) {
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
}

/** 학생에게 그대로 보여 줄 까닭. 이것으로 던진 글만 화면에 뜹니다 */
function oops_(msg) {
  const e = new Error(msg);
  e.forStudent = true;
  return e;
}

/** 학생에게 보낼 실패 모양. 서버 속사정(구글 오류 글)은 화면에 띄우지 않고 why 에만 담습니다 */
function fail_(err) {
  const msg = String((err && err.message) || err || '');
  return { ok: false, msg: err && err.forStudent ? msg : '저장하지 못했어요. 잠시 뒤 다시 눌러요.', why: msg.slice(0, 200) };
}

/* ================= 학생 화면 열기 ================= */

/** 학생이 /exec 주소를 열면 Index.html 을 보여 줍니다 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(APP.NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ================= 드라이브에 파일 저장 ================= */

/**
 * 제출물 폴더. 시트가 있는 폴더 안에 「앱 이름 제출물」을 한 번 만들고 그 ID 를 기억합니다.
 * 폴더는 선생님만 열립니다. 학생에게 공유하지 않습니다.
 */
function folder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('FOLDER_ID');
  if (id) {
    try {
      const f = DriveApp.getFolderById(id);
      if (!f.isTrashed || !f.isTrashed()) return f;
    } catch (e) { /* 지워졌으면 새로 만듭니다 */ }
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const again = props.getProperty('FOLDER_ID');
    if (again && again !== id) return DriveApp.getFolderById(again);
    const parents = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId()).getParents();
    const parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
    const f = parent.createFolder(APP.NAME + ' 제출물');
    props.setProperty('FOLDER_ID', f.getId());
    return f;
  } finally {
    lock.releaseLock();
  }
}

/** 과제마다 하위 폴더. 두 학생이 같은 순간 처음 내도 폴더가 두 개 생기지 않게 잠급니다 */
function taskFolder_(task) {
  const root = folder_();
  const found = root.getFoldersByName(task);
  if (found.hasNext()) return found.next();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const again = root.getFoldersByName(task);
    return again.hasNext() ? again.next() : root.createFolder(task);
  } finally {
    lock.releaseLock();
  }
}

const EXT = {
  'image/jpeg': 'jpg', 'image/png': 'png', 'image/gif': 'gif', 'image/webp': 'webp', 'image/heic': 'heic',
  'application/pdf': 'pdf', 'audio/webm': 'webm', 'audio/mp4': 'm4a', 'audio/x-m4a': 'm4a', 'audio/mpeg': 'mp3',
  'audio/wav': 'wav', 'audio/ogg': 'ogg', 'audio/aac': 'aac', 'video/webm': 'webm'
};

/**
 * 학생이 보낸 파일 하나를 과제 폴더에 저장합니다.
 * file = { b64, mime, name }. 파일 이름은 「6-1-03 김정호 사진1 0914-103245.jpg」 모양입니다.
 */
function saveFile_(w, file, label) {
  if (!file || !file.b64) throw oops_('파일이 비어 있어요.');
  const bytes = Utilities.base64Decode(String(file.b64));
  if (bytes.length > APP.MAX_FILE_MB * 1024 * 1024) throw oops_('파일이 ' + APP.MAX_FILE_MB + 'MB 보다 커요. 줄여서 올려요.');
  const mime = String(file.mime || 'application/octet-stream').split(';')[0].trim().toLowerCase();
  const fromName = (String(file.name || '').match(/\.([A-Za-z0-9]{1,5})$/) || [])[1];
  const ext = (fromName || EXT[mime] || 'bin').toLowerCase();
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMdd-HHmmss');
  const name = w.grade + '-' + w.cls + '-' + ('0' + w.num).slice(-2) + ' ' + w.name + ' ' + label + ' ' + stamp + '.' + ext;
  const saved = taskFolder_(w.task).createFile(Utilities.newBlob(bytes, mime, name));
  return { id: saved.getId(), url: saved.getUrl(), name: name, bytes: bytes, mime: mime };
}

/** 학생 화면이 돌려준 파일 ID 가 정말 이 앱의 제출물 폴더 안 파일인지 봅니다 */
function ownFile_(id) {
  const f = DriveApp.getFileById(String(id || ''));
  const rootId = folder_().getId();
  const p1 = f.getParents();
  if (!p1.hasNext()) throw oops_('파일을 찾지 못했어요. 다시 올려요.');
  const task = p1.next();
  const p2 = task.getParents();
  if (task.getId() !== rootId && !(p2.hasNext() && p2.next().getId() === rootId)) throw oops_('파일을 찾지 못했어요. 다시 올려요.');
  return f;
}

/* ================= 드라이브 OCR: 키 없이 사진 속 글자 읽기 =================
 *
 * 사진을 「구글 문서로 바꿔 올리기」로 보내면 드라이브가 사진 속 글자를 읽어 문서에 적습니다.
 * 드라이브에서 사진을 오른쪽 클릭 → 「Google 문서로 열기」 하는 것과 같은 기능입니다.
 * 그 문서를 글로 내려받고, 문서는 곧바로 휴지통으로 보냅니다. 선생님 계정 안에서 끝나고 키가 필요 없습니다.
 *
 * 구글 도움말의 조건: 2MB 이하 · jpg png gif · 글자 높이 10픽셀 이상. 그림과 그래프는 글로 옮기지 않습니다.
 * 학생 화면이 사진을 가로 세로 1600픽셀 안으로 줄여 보내므로 보통 2MB 를 넘지 않습니다.
 */
const OCR_MIME = { 'image/jpeg': true, 'image/png': true, 'image/gif': true };

function ocrBytes_(bytes, mime) {
  if (!OCR_MIME[mime]) return { ok: false, text: '', msg: 'jpg · png 사진만 글자를 읽어요.' };
  if (bytes.length > 2 * 1024 * 1024) return { ok: false, text: '', msg: '사진이 2MB 를 넘어 글자를 읽지 못했어요.' };
  const token = ScriptApp.getOAuthToken();
  const boundary = 'ocr' + Date.now();
  const meta = { name: '_글자읽기_' + Date.now(), mimeType: 'application/vnd.google-apps.document', parents: [folder_().getId()] };
  const head = '--' + boundary + '\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n' + JSON.stringify(meta) +
               '\r\n--' + boundary + '\r\nContent-Type: ' + mime + '\r\n\r\n';
  const tail = '\r\n--' + boundary + '--';
  const payload = Utilities.newBlob(head).getBytes().concat(Array.prototype.slice.call(bytes), Utilities.newBlob(tail).getBytes());

  const up = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&ocrLanguage=ko&fields=id', {
    method: 'post', contentType: 'multipart/related; boundary=' + boundary, payload: payload,
    headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true
  });
  if (up.getResponseCode() !== 200) {
    return { ok: false, text: '', msg: '드라이브가 사진을 받지 못했어요 (' + up.getResponseCode() + ').', why: up.getContentText().slice(0, 200) };
  }
  const docId = JSON.parse(up.getContentText()).id;
  try {
    const res = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(docId) + '/export?mimeType=text%2Fplain', {
      headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 200) return { ok: false, text: '', msg: '읽은 글을 받지 못했어요 (' + res.getResponseCode() + ').' };
    const text = String(res.getContentText('UTF-8') || '').replace(/^\uFEFF/, '').replace(/\r\n/g, '\n').trim();
    return text ? { ok: true, text: text } : { ok: false, text: '', msg: '사진에서 글자를 찾지 못했어요.' };
  } finally {
    try { DriveApp.getFileById(docId).setTrashed(true); } catch (e) { /* 못 치운 문서는 제출물 폴더에 「_글자읽기_」로 남습니다 */ }
  }
}

/** 편집기에서 실행해 보는 연결 시험. 글자를 문서로 바꿨다 되받아 봅니다 */
function testDriveOcr() {
  const said = '연결 시험 ' + Math.floor(1000 + Math.random() * 9000);
  const token = ScriptApp.getOAuthToken();
  const boundary = 'ocr' + Date.now();
  const meta = { name: '_연결시험_', mimeType: 'application/vnd.google-apps.document', parents: [folder_().getId()] };
  const body = '--' + boundary + '\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n' + JSON.stringify(meta) +
               '\r\n--' + boundary + '\r\nContent-Type: text/plain\r\n\r\n' + said + '\r\n--' + boundary + '--';
  const up = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id', {
    method: 'post', contentType: 'multipart/related; boundary=' + boundary, payload: Utilities.newBlob(body).getBytes(),
    headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true
  });
  if (up.getResponseCode() !== 200) { say_('드라이브 연결 실패 (' + up.getResponseCode() + ')\n' + up.getContentText().slice(0, 300)); return; }
  const id = JSON.parse(up.getContentText()).id;
  try { DriveApp.getFileById(id).setTrashed(true); } catch (e) { }
  say_('드라이브 연결이 됩니다. 사진 속 글자 읽기를 쓸 수 있습니다.');
}

/* ================= 사진 → 글자 읽기 → 학생이 확인하고 내기 ================= */

/**
 * 1단계. 사진을 제출물 폴더에 저장하고 드라이브 OCR 로 글자를 읽어 돌려줍니다.
 * 학생은 읽은 글을 화면에서 고친 뒤 submitChecked 로 냅니다.
 * p.images = [{ b64, mime, name }]
 */
function readPhotos(p) {
  try {
    const w = who_(p);
    const list = (p.images || []).slice(0, APP.MAX_PHOTOS);
    if (!list.length) throw oops_('사진을 먼저 골라요.');
    const files = [], texts = [], misses = [];
    list.forEach(function (img, i) {
      const s = saveFile_(w, img, APP.PHOTO_LABEL + (list.length > 1 ? (i + 1) : ''));
      files.push({ id: s.id });
      const r = ocrBytes_(s.bytes, s.mime);
      if (r.ok) texts.push(list.length > 1 ? '[' + (i + 1) + '쪽]\n' + r.text : r.text);
      else misses.push((list.length > 1 ? (i + 1) + '쪽: ' : '') + r.msg);
    });
    return { ok: true, files: files, text: texts.join('\n\n'), misses: misses };
  } catch (e) {
    return fail_(e);
  }
}

/**
 * 2단계. 학생이 확인한 글과 드라이브가 처음 읽은 글을 함께 남깁니다.
 * 두 칸을 견주면 학생이 어디를 고쳤는지 보입니다.
 */
function submitChecked(p) {
  try {
    const w = who_(p);
    const ids = (p.fileIds || []).slice(0, APP.MAX_PHOTOS);
    if (!ids.length) throw oops_('사진을 먼저 올려요.');
    const text = clean_(p.text, APP.MAX_TEXT);
    if (!text) throw oops_('글이 비어 있어요. 읽은 글을 확인하거나 직접 적어요.');
    const links = ids.map(function (id, i) {
      return link_(ownFile_(id).getUrl(), APP.PHOTO_LABEL + (ids.length > 1 ? (i + 1) : '') + ' 열기');
    });
    while (links.length < APP.MAX_PHOTOS) links.push('');
    addRow_(whoCells_(w).concat([cell_(text), cell_(clean_(p.ocr, APP.MAX_TEXT))], links));
    return { ok: true, at: hhmm_() };
  } catch (e) {
    return fail_(e);
  }
}
