/**
 * 세미나 요청서 <-> 구글시트 연동용 Apps Script  (v2)
 *
 * v1에서 달라진 점
 *  1) 요청부서를 폼에서 직접 입력할 수 있게 됨 → 시트 드롭다운(데이터 확인)에 걸리지 않도록 처리
 *  2) '부서 관리' 탭이 있으면 부서 버튼 목록을 시트에서 읽어옴 (없으면 기본 2개 사용 — 시트 안 고쳐도 동작함)
 *  3) 접수번호가 겹치지 않도록 생성 방식 변경 (기존: 행 번호 기준 → 변경: 기존 최대번호 +1)
 *  4) 두 사람이 동시에 제출해도 덮어쓰기가 나지 않도록 잠금(LockService) 추가
 *  5) '기수 관리' 참조 범위를 넉넉하게 확장 (기수를 계속 추가해도 잠금 수식이 계속 동작)
 *
 * [설치 방법]  ※ v1과 동일합니다
 * 1. 구글시트(세미나_기획_프로세스_통합관리시트) 열기
 * 2. 상단 메뉴 확장 프로그램 → Apps Script
 * 3. Code.gs 내용을 전부 지우고 이 파일 내용을 붙여넣기 → 저장(Ctrl+S)
 * 4. 우측 상단 "배포" → "배포 관리" → 연필(수정) → 버전 "새 버전" → "배포"
 *    (처음이라면 "새 배포" → 유형 "웹 앱" → 실행 계정 "나" → 액세스 "모든 사용자" → 배포)
 * 5. 웹 앱 URL을 HTML 파일의 API_URL 에 붙여넣기
 *
 * ※ 코드를 고친 뒤에는 반드시 "배포 → 배포 관리 → 수정 → 새 버전 → 배포"를 다시 눌러야
 *   변경사항이 실제로 반영됩니다. (URL은 그대로 유지됩니다)
 */

const SHEET_NAME = "니즈접수";
const QUARTER_SHEET_NAME = "기수 관리";
const DEPT_SHEET_NAME = "부서 관리";   // 없어도 됩니다. 있으면 자동으로 읽어옵니다.

// 시트를 안 고쳤을 때 사용할 기본 부서 목록
const DEFAULT_DEPTS = ["채널마케팅본부", "프로덕트마케팅본부"];

// 니즈접수 탭의 컬럼 순서 (A=1 기준)
const COL = {
  ID: 1,             // 접수번호
  QTR: 2,            // 대상세미나월(기수)
  DEPT: 3,           // 요청부서
  OWNER: 4,          // 담당자
  PRODUCT_TYPE: 5,   // 제품구분
  PRODUCT_NAME: 6,   // 교재·서비스명
  PRODUCT_DETAIL: 7, // 상세내용
  KIND: 8,           // 세미나종류
  TARGET: 9,         // 목표고객
  BACKGROUND: 10,    // 요청사유(배경)
  EFFECT: 11,        // 기대효과
  REF: 12,           // 참고자료
  EMAIL: 13,         // 작성자 이메일
  SUBMITTED: 14,     // 제출일시
  STATUS: 15,        // 확인상태
  HOLD_REASON: 16,   // 차기 검토 사유
  LOCK: 17,          // 잠금상태 (수식 - 건드리지 않음)
  NOTIF: 18          // 알림발송이력
};
const TOTAL_COLS = 18;

/* =========================================================
   GET : 접수 목록 + 부서 목록을 함께 내려줌
   ========================================================= */
function doGet(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    const depts = getDeptList();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return jsonResponse({ ok: true, needs: [], depts: depts });
    }
    const values = sheet.getRange(2, 1, lastRow - 1, TOTAL_COLS).getValues();
    const needs = [];
    values.forEach(function (row) {
      const id = row[COL.ID - 1];
      if (!id) return; // 접수번호 없는 빈 행은 건너뜀
      needs.push({
        id: id,
        qtr: row[COL.QTR - 1],
        dept: row[COL.DEPT - 1],
        owner: row[COL.OWNER - 1],
        productType: row[COL.PRODUCT_TYPE - 1],
        productName: row[COL.PRODUCT_NAME - 1],
        productDetail: row[COL.PRODUCT_DETAIL - 1],
        kind: row[COL.KIND - 1],
        target: row[COL.TARGET - 1],
        background: row[COL.BACKGROUND - 1],
        effect: row[COL.EFFECT - 1],
        ref: row[COL.REF - 1],
        email: row[COL.EMAIL - 1],
        submitted: formatDate(row[COL.SUBMITTED - 1]),
        status: row[COL.STATUS - 1],
        holdReason: row[COL.HOLD_REASON - 1] || ""
      });
    });
    return jsonResponse({ ok: true, needs: needs, depts: depts });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.toString() });
  }
}

/* =========================================================
   POST : 새 요청서 1건 저장
   ========================================================= */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    // 동시에 제출해도 같은 줄에 겹쳐 쓰지 않도록 최대 20초 대기
    lock.waitLock(20000);

    const data = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    const newRowIndex = sheet.getLastRow() + 1;
    const now = new Date();
    const id = makeNextId(sheet, now);

    const row = new Array(TOTAL_COLS).fill("");
    row[COL.ID - 1] = id;
    row[COL.QTR - 1] = data.qtr || "";
    row[COL.DEPT - 1] = String(data.dept || "").trim();
    row[COL.OWNER - 1] = data.owner || "";
    row[COL.PRODUCT_TYPE - 1] = data.productType || "";
    row[COL.PRODUCT_NAME - 1] = data.productName || "";
    row[COL.PRODUCT_DETAIL - 1] = data.productDetail || "";
    row[COL.KIND - 1] = data.kind || "";
    row[COL.TARGET - 1] = data.target || "";
    row[COL.BACKGROUND - 1] = data.background || "";
    row[COL.EFFECT - 1] = data.effect || "";
    row[COL.REF - 1] = data.ref || "";
    row[COL.EMAIL - 1] = data.email || "";
    row[COL.SUBMITTED - 1] = now;
    row[COL.STATUS - 1] = "접수완료";
    row[COL.HOLD_REASON - 1] = "";
    row[COL.NOTIF - 1] = "";

    const range = sheet.getRange(newRowIndex, 1, 1, TOTAL_COLS);

    // 요청부서 칸에 드롭다운(데이터 확인)이 걸려 있으면 직접 입력한 부서명이 거부되므로,
    // 이 행의 부서 칸에 한해 규칙을 해제한 뒤 값을 기록합니다.
    try {
      sheet.getRange(newRowIndex, COL.DEPT).clearDataValidations();
    } catch (ignore) {}

    range.setValues([row]);

    // 잠금상태(Q열)는 '기수 관리' 탭을 참조하는 수식으로 채움 (범위를 넉넉히 잡아 기수 추가에 대응)
    const lockFormula =
      "=IFERROR(IF(TODAY()>INDEX('" + QUARTER_SHEET_NAME + "'!$B$2:$B$200," +
      "MATCH(B" + newRowIndex + ",'" + QUARTER_SHEET_NAME + "'!$A$2:$A$200,0))," +
      '"마감(수정불가)","입력가능"),"")';
    sheet.getRange(newRowIndex, COL.LOCK).setFormula(lockFormula);

    return jsonResponse({ ok: true, id: id, submitted: formatDate(now) });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.toString() });
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

/* =========================================================
   보조 함수
   ========================================================= */

/**
 * '부서 관리' 탭 A2:A 의 값을 부서 목록으로 사용.
 * 탭이 없거나 비어 있으면 기본 목록을 사용합니다. (시트를 안 고쳐도 동작)
 */
function getDeptList() {
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DEPT_SHEET_NAME);
    if (!sh) return DEFAULT_DEPTS;
    const last = sh.getLastRow();
    if (last < 2) return DEFAULT_DEPTS;
    const vals = sh.getRange(2, 1, last - 1, 1).getValues();
    const list = [];
    vals.forEach(function (r) {
      const v = String(r[0] || "").trim();
      if (v && list.indexOf(v) === -1) list.push(v);
    });
    return list.length ? list : DEFAULT_DEPTS;
  } catch (err) {
    return DEFAULT_DEPTS;
  }
}

/**
 * 접수번호 생성: 올해 기존 번호 중 가장 큰 값 +1
 * (중간 행을 지워도 번호가 겹치지 않습니다)
 */
function makeNextId(sheet, now) {
  const year = now.getFullYear();
  const prefix = "N-" + year + "-";
  let max = 0;
  const last = sheet.getLastRow();
  if (last >= 2) {
    const ids = sheet.getRange(2, COL.ID, last - 1, 1).getValues();
    ids.forEach(function (r) {
      const v = String(r[0] || "");
      if (v.indexOf(prefix) === 0) {
        const n = parseInt(v.substring(prefix.length), 10);
        if (!isNaN(n) && n > max) max = n;
      }
    });
  }
  return prefix + String(max + 1).padStart(3, "0");
}

function formatDate(d) {
  if (!d) return "";
  if (!(d instanceof Date)) return String(d);
  const pad = function (n) { return String(n).padStart(2, "0"); };
  return d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate()) +
         " " + pad(d.getHours()) + ":" + pad(d.getMinutes());
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * [점검용] Apps Script 편집기에서 이 함수를 한 번 실행해보면
 * 시트 연결과 부서 목록이 정상인지 로그로 확인할 수 있습니다.
 * (실행 → checkSetup → 하단 '실행 로그' 확인)
 */
function checkSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME);
  Logger.log("시트 이름: " + (sh ? "정상 (" + SHEET_NAME + ")" : "★ '" + SHEET_NAME + "' 탭을 찾을 수 없습니다"));
  if (sh) Logger.log("현재 접수 건수: " + Math.max(0, sh.getLastRow() - 1) + "건");
  Logger.log("기수 관리 탭: " + (ss.getSheetByName(QUARTER_SHEET_NAME) ? "정상" : "★ 없음"));
  Logger.log("부서 관리 탭: " + (ss.getSheetByName(DEPT_SHEET_NAME) ? "정상" : "없음 → 기본 목록 사용"));
  Logger.log("부서 목록: " + getDeptList().join(", "));
}
