/**
 * [가] 학생 글 모으기
 * AI·디지털 기반 수업과 평가 연수 · 단계별 앱
 *
 * 학생이 친 글을 [응답] 시트에 한 줄씩 모읍니다. 파일도 AI 도 쓰지 않습니다.
 *
 * 1. 구글 시트 → 확장 프로그램 → Apps Script → 이 글 전체를 Code.gs 에 붙여 넣기 (있던 글은 지우고)
 * 2. 파일 옆 + → HTML → 이름 Index → Index.html 내용 붙여 넣기
 * 3. 위쪽 함수 칸에서 setup 을 고르고 ▷실행 → 권한 허용
 * 4. 배포 → 새 배포 → 유형 「웹 앱」 → 실행 사용자 「나」 · 액세스 권한 「모든 사용자」 → 배포 → /exec 주소 복사
 * 코드를 고친 뒤에는 배포 → 배포 관리 → 연필 → 버전 「새 버전」 → 배포. 「새 배포」를 누르면 주소가 바뀝니다.
 */

/* ===== 바꿀 곳은 여기 ===== */
const APP = {
  NAME: '학생 글 모으기',
  CLASS_CODE: '',      // 우리 반 코드 (예: '4817'). 비워 두면 묻지 않습니다
  MAX_TEXT: 5000       // 답 글자 수 한도
};

/** [응답] 시트의 열. 순서를 바꾸면 submitText 의 줄도 같이 바꿉니다 */
const HEAD = ['낸 시각', '학년', '반', '번호', '이름', '과제', '답'];

/** 처음 설정 때 [과제] 시트에 들어가는 예시. 시트에서 고치면 됩니다 */
const SAMPLE_TASKS = [
  ['계절이 바뀌는 까닭', '지구의 자전축이 기울어진 채 공전한다는 말을 넣어, 우리나라에 여름과 겨울이 생기는 까닭을 설명해요.', 'O'],
  ['연소의 조건', '촛불을 컵으로 덮으면 꺼지는 까닭을 연소의 조건과 이어서 써요.', 'O']
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

/* ================= 이 단계에서 하는 일 ================= */

/** 학생 화면의 「내기」. p = { grade, cls, num, name, code, task, text } */
function submitText(p) {
  try {
    const w = who_(p);
    const text = clean_(p.text, APP.MAX_TEXT);
    if (!text) throw oops_('답을 적어요.');
    addRow_(whoCells_(w).concat([cell_(text)]));
    return { ok: true, at: hhmm_() };
  } catch (e) {
    return fail_(e);
  }
}
