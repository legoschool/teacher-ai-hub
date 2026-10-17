/**
 * [바] 녹음 파일 올리고 선생님이 나중에 받아쓰기
 * AI·디지털 기반 수업과 평가 연수 · 단계별 앱
 *
 * 학생은 기기 녹음 앱으로 녹음한 파일(또는 화면의 녹음 단추가 여는 녹음기)을 올립니다. [응답] 시트에 「받아쓰기 대기」로 쌓입니다.
 * 수업이 끝난 뒤 선생님이 시트 메뉴 「받아쓰기 대기 모두 받아쓰기」를 누르면 Gemini API 가 녹음을 글로 옮겨 적습니다.
 *
 * 1. 구글 시트 → 확장 프로그램 → Apps Script → 이 글 전체를 Code.gs 에 붙여 넣기 (있던 글은 지우고)
 * 2. 파일 옆 + → HTML → 이름 Index → Index.html 내용 붙여 넣기
 * 3. 위쪽 함수 칸에서 setup 을 고르고 ▷실행 → 권한 허용
 * 4. 시트로 돌아가 새로 고침 → 메뉴 「Gemini API 키 넣기」 → 「Gemini 연결 시험」
 * 5. 배포 → 새 배포 → 유형 「웹 앱」 → 실행 사용자 「나」 · 액세스 권한 「모든 사용자」 → 배포 → /exec 주소 복사
 * 코드를 고친 뒤에는 배포 → 배포 관리 → 연필 → 버전 「새 버전」 → 배포. 「새 배포」를 누르면 주소가 바뀝니다.
 */

/* ===== 바꿀 곳은 여기 ===== */
const APP = {
  NAME: '녹음 파일 받아쓰기',
  CLASS_CODE: '',       // 우리 반 코드 (예: '4817'). 비워 두면 묻지 않습니다
  MAX_TEXT: 1000,       // 학생 메모 글자 수
  MAX_FILE_MB: 10,      // 올리는 녹음 파일 크기 한도
  MAX_STT_MB: 14        // 받아쓰기에 보내는 파일 크기 한도. 한 번 요청에 담아 보내는 방식이라 크게 잡지 않습니다
};

const HEAD = ['낸 시각', '학년', '반', '번호', '이름', '과제', '학생 메모', '녹음 파일', '파일 ID', '상태', '받아쓰기', '받아쓰기 시각'];

const MENU_EXTRA = [
  ['받아쓰기 대기 모두 받아쓰기', 'transcribeWaiting'],
  ['고른 줄 다시 받아쓰기', 'transcribeSelected'],
  ['Gemini API 키 넣기', 'setApiKey'],
  ['Gemini 연결 시험', 'testGemini']
];

const SAMPLE_TASKS = [
  ['과학 발표 녹음', '모둠 발표를 녹음해서 올려요. 파일 이름은 바꾸지 않아도 돼요.', 'O'],
  ['낭독 녹음', '교과서 42쪽 첫 문단을 읽고 녹음해서 올려요.', 'O']
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

/* ================= 선생님이 나중에 받아쓰기: Gemini API =================
 *
 * 녹음 파일을 글로 바꾸는 무료 길은 키 없이는 없습니다. 구글 AI 스튜디오에서 Gemini API 키를 받아
 * [프로젝트 설정] → [스크립트 속성] 에 GEMINI_API_KEY 로 넣습니다. 키는 코드와 시트에 적지 않습니다.
 *
 * 무료 등급 키로 보낸 내용은 구글이 제품 개선에 쓸 수 있다고 가격 안내에 적혀 있습니다.
 * 학생 목소리를 보내기 전에 결제를 연결한 유료 등급으로 쓸지, 학교 개인정보 담당자와 먼저 정합니다.
 */

const STT_PROMPT =
  '이 녹음은 초등학생이 수업 시간에 말한 답변입니다. 말한 내용을 그대로 받아써 주세요.\n' +
  '규칙\n' +
  '1. 말한 그대로 적습니다. 문장을 다듬거나 요약하지 않습니다.\n' +
  '2. "음", "어" 같은 군말은 빼도 됩니다.\n' +
  '3. 문장이 끝나는 곳에서 줄을 바꿉니다.\n' +
  '4. 알아들을 수 없는 부분은 (알아듣기 어려움) 으로 적습니다.\n' +
  '5. 설명 없이 받아쓴 글만 출력합니다.';

/** 모델 이름은 구글이 바꾸고 내립니다. 스크립트 속성 GEMINI_MODEL 에 새 이름을 적으면 그것부터 씁니다 */
const GEMINI_MODELS = ['gemini-flash-latest', 'gemini-2.5-flash'];

/** 드라이브가 적은 형식 이름을 Gemini 가 알아듣는 이름으로. 소리만 든 webm 을 드라이브는 video/webm 으로 적습니다 */
function audioMime_(mime, name) {
  mime = String(mime || '').split(';')[0].trim().toLowerCase();
  const ext = (String(name || '').match(/\.([a-z0-9]+)$/i) || [])[1] || '';
  if (mime === 'video/webm' || (!/^audio\//.test(mime) && /^webm$/i.test(ext))) return 'audio/webm';
  const map = { 'audio/x-m4a': 'audio/mp4', 'audio/m4a': 'audio/mp4', 'video/mp4': 'audio/mp4', 'audio/x-wav': 'audio/wav',
                'audio/wave': 'audio/wav', 'audio/mpeg': 'audio/mp3', 'audio/x-aac': 'audio/aac' };
  if (map[mime]) return map[mime];
  if (!/^audio\//.test(mime)) {
    const byExt = { m4a: 'audio/mp4', mp3: 'audio/mp3', wav: 'audio/wav', ogg: 'audio/ogg', aac: 'audio/aac', flac: 'audio/flac' };
    if (byExt[ext.toLowerCase()]) return byExt[ext.toLowerCase()];
  }
  return mime;
}

/** Gemini 에 한 번 묻습니다. 모델이 없다는 대답(404)일 때만 다음 모델로 넘어갑니다 */
function gemini_(parts) {
  const props = PropertiesService.getScriptProperties();
  const key = String(props.getProperty('GEMINI_API_KEY') || '').trim();
  if (!key) return { ok: false, stop: true, msg: 'GEMINI_API_KEY 가 없습니다. 메뉴 「Gemini API 키 넣기」로 넣으세요.' };
  const models = [props.getProperty('GEMINI_MODEL')].concat(GEMINI_MODELS)
    .filter(function (m, i, a) { return m && a.indexOf(m) === i; });
  let last = '';
  for (let i = 0; i < models.length; i++) {
    const res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + models[i] + ':generateContent', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { 'x-goog-api-key': key },       // 키를 주소에 붙이지 않습니다. 오류 기록에 주소가 남습니다
      payload: JSON.stringify({ contents: [{ parts: parts }], generationConfig: { temperature: 0, maxOutputTokens: 8192 } })
    });
    const code = res.getResponseCode();
    const body = res.getContentText();
    let j = {};
    try { j = JSON.parse(body); } catch (e) { j = {}; }
    const said = (j.error && j.error.message) || body.slice(0, 200);
    if (code === 200) {
      const cand = (j.candidates || [])[0];
      const text = cand && cand.content && cand.content.parts
        ? cand.content.parts.map(function (x) { return x.text || ''; }).join('').trim() : '';
      return text ? { ok: true, text: text, model: models[i] } : { ok: false, msg: '빈 답이 왔습니다 (' + ((cand && cand.finishReason) || '까닭 없음') + ')' };
    }
    if (code === 404) { last = models[i] + ' 모델 없음'; continue; }
    if (code === 429) return { ok: false, stop: true, msg: '한도에 걸렸습니다 (429). 잠시 뒤 다시 누르세요. ' + said };
    if (code === 400 || code === 401 || code === 403) {
      const auth = /API key|API_KEY|PERMISSION|billing/i.test(said);
      return { ok: false, stop: auth, msg: '(' + code + ') ' + said };
    }
    return { ok: false, msg: '(' + code + ') ' + said };
  }
  return { ok: false, stop: true, msg: last + '. 스크립트 속성 GEMINI_MODEL 에 AI 스튜디오에 보이는 모델 이름을 적으세요.' };
}

/** 녹음 파일 하나를 받아씁니다 */
function transcribeFile_(fileId) {
  const f = DriveApp.getFileById(fileId);
  const blob = f.getBlob();
  const bytes = blob.getBytes();
  if (bytes.length > APP.MAX_STT_MB * 1024 * 1024) return { ok: false, msg: '파일이 ' + APP.MAX_STT_MB + 'MB 를 넘습니다. 짧게 나눠 받으세요.' };
  const mime = audioMime_(f.getMimeType(), f.getName());
  return gemini_([{ text: STT_PROMPT }, { inline_data: { mime_type: mime, data: Utilities.base64Encode(bytes) } }]);
}

/** 메뉴: 「받아쓰기 대기」인 줄을 차례로 받아씁니다. 6분 실행 한도 전에 멈추고 남은 수를 알려 줍니다 */
function transcribeWaiting() {
  transcribeRows_(null);
}

/** 메뉴: 시트에서 고른 줄만 다시 받아씁니다 (완료된 줄도) */
function transcribeSelected() {
  const sh = SpreadsheetApp.getActiveSheet();
  if (sh.getName() !== SHEET_ANSWER) { say_('[' + SHEET_ANSWER + '] 시트에서 줄을 고른 뒤 누르세요.'); return; }
  const r = sh.getActiveRange();
  const rows = [];
  for (let i = r.getRow(); i < r.getRow() + r.getNumRows(); i++) if (i > 1) rows.push(i);
  transcribeRows_(rows);
}

function transcribeRows_(only) {
  const t0 = Date.now();
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ANSWER);
  if (!sh || sh.getLastRow() < 2) { say_('받아쓸 줄이 없습니다.'); return; }
  const cId = HEAD.indexOf('파일 ID'), cState = HEAD.indexOf('상태'), cText = HEAD.indexOf('받아쓰기');
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, HEAD.length).getValues();
  const todo = [];
  values.forEach(function (row, i) {
    const rowNo = i + 2;
    if (!row[cId]) return;
    if (only ? only.indexOf(rowNo) >= 0 : String(row[cState]) === '받아쓰기 대기') todo.push({ rowNo: rowNo, id: String(row[cId]) });
  });
  if (!todo.length) { say_(only ? '고른 줄에 녹음 파일이 없습니다.' : '「받아쓰기 대기」인 줄이 없습니다.'); return; }

  let done = 0, failed = 0, stopMsg = '';
  for (let k = 0; k < todo.length; k++) {
    if (Date.now() - t0 > 270000) break;          // 6분 한도 전에 멈춥니다
    let r;
    try { r = transcribeFile_(todo[k].id); } catch (e) { r = { ok: false, msg: String(e.message || e).slice(0, 200) }; }
    const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    if (r.ok) {
      sh.getRange(todo[k].rowNo, cState + 1, 1, 3).setValues([['완료', cell_(r.text), stamp]]);
      done++;
    } else {
      sh.getRange(todo[k].rowNo, cState + 1, 1, 3).setValues([['실패', cell_(r.msg), stamp]]);
      failed++;
      if (r.stop) { stopMsg = r.msg; break; }
    }
    SpreadsheetApp.flush();
  }
  const left = todo.length - done - failed;
  say_('받아쓰기 ' + done + '건 완료, ' + failed + '건 실패' + (left ? ', ' + left + '건 남음 (메뉴를 한 번 더 누르세요)' : '') +
       (stopMsg ? '\n\n멈춘 까닭: ' + stopMsg : '') +
       (failed ? '\n\n실패한 줄은 상태를 「받아쓰기 대기」로 고치거나 줄을 골라 「고른 줄 다시 받아쓰기」를 누릅니다.' : ''));
}

/** 메뉴: 키를 스크립트 속성에 넣습니다. 편집기 [프로젝트 설정] → [스크립트 속성] 에서 넣어도 같습니다 */
function setApiKey() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Gemini API 키 넣기', 'AI 스튜디오(aistudio.google.com)에서 받은 키를 붙여 넣으세요.\n키는 스크립트 속성에만 저장되고 시트에는 남지 않습니다.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const key = String(r.getResponseText() || '').trim();
  if (!key) return;
  PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', key);
  ui.alert('저장했습니다. 끝자리 ' + key.slice(-4) + '. 「연결 시험」을 눌러 보세요.');
}

/** 메뉴: 짧은 글로 키와 모델이 되는지 봅니다 */
function testGemini() {
  const r = gemini_([{ text: '"연결됨" 이라고만 답하세요.' }]);
  say_(r.ok ? 'Gemini 연결이 됩니다. 모델 ' + r.model + '\n답: ' + r.text : 'Gemini 연결이 안 됩니다.\n' + r.msg);
}

/* ================= 이 단계에서 하는 일 ================= */

/** 학생 화면의 「내기」. p.audio = { b64, mime, name } */
function submitAudio(p) {
  try {
    const w = who_(p);
    if (!p.audio || !p.audio.b64) throw oops_('녹음 파일을 먼저 골라요.');
    const mime = String(p.audio.mime || '').toLowerCase();
    if (mime && !/^(audio|video)\//.test(mime)) throw oops_('소리 파일만 올릴 수 있어요.');
    const s = saveFile_(w, p.audio, '녹음');
    addRow_(whoCells_(w).concat([cell_(clean_(p.text, APP.MAX_TEXT)), link_(s.url, '녹음 듣기'), s.id, '받아쓰기 대기', '', '']));
    return { ok: true, at: hhmm_() };
  } catch (e) {
    return fail_(e);
  }
}
