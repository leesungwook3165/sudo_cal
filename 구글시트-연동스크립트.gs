/*
 * 수도요금 정산 — 구글 시트 연동 스크립트 (v5: 구성원 관리 확장)
 *
 * v5 변경점: 구성원 탭에 "구분(관리자/구성원)"·"계좌" 칼럼 추가 (기존 시트 자동 보강),
 *   앱에서 구성원 편집 저장 지원, 정산 메일·가족 화면에 관리자 계좌(입금 안내) 표시
 *
 * v4 변경점: 사진이 드라이브에 "수도정산-고지서 / 2026년 / 06월분" 형태의
 *   월별 폴더로 자동 정리됩니다. (정산 사진 = 고지서 귀속월, 검침 사진 = 검침일의 달)
 * v3 변경점: "중간검침" 탭 — 매월 검침값·계량기 사진만 기록 (정산·메일과 무관)
 *   기존 v2에서 업그레이드하려면 이 파일 전체를 다시 붙여넣고
 *   [배포] → [배포 관리] → 기존 배포 ✏️ → 버전 "새 버전" 으로 재배포하세요. (URL 유지됨)
 *
 * ── 설치 방법 ─────────────────────────────────────────────
 * 1. 새 구글 시트를 만든다 (이름 예: "수도요금 정산")
 * 2. 시트 메뉴 [확장 프로그램] → [Apps Script]
 * 3. 기본 코드를 지우고 이 파일 전체를 붙여넣기 → 저장
 * 4. [배포] → [새 배포] → 유형 "웹 앱"
 *    - 실행 계정: 나 / 액세스: 링크가 있는 모든 사용자
 * 5. 승인 후 발급된 URL(https://script.google.com/macros/s/…/exec)을
 *    정산 계산기 앱의 "구글 시트 연동" 칸에 저장
 *
 * ── 권한 설정 (관리자만 수정, 가족은 열람만) ─────────────────
 * 1. 우측 상단 [공유] → "링크가 있는 모든 사용자" → 역할 "뷰어"
 * 2. 그 링크를 가족방에 공유 — 가족은 모든 기록을 볼 수 있지만 수정은 불가
 * 3. 수정 권한은 시트 소유자(관리자)에게만 있음
 * 4. "구성원" 탭의 가족 이름·이메일·전화번호는 관리자가 직접 입력
 *
 * ── 알림 동작 ────────────────────────────────────────────
 * 정산이 저장되면:
 *  - "구성원" 탭에서 메일수신=Y 인 사람에게 정산 내역 메일 자동 발송
 *  - 전화번호 목록은 앱으로 돌려보내져 "문자 보내기" 버튼에 사용됨
 */

const SHEET_NAME  = '정산기록';
const MEMBER_NAME = '구성원';
const MID_NAME    = '중간검침';   // 매월 검침값·사진만 기록 (정산과 무관)
const FOLDER_NAME = '수도정산-고지서';

/* ★ 반드시 나만 아는 값으로 변경하세요 (앱의 "관리자 키"와 동일해야 함)
 *   키가 틀린 요청은 기록·조회 모두 거부됩니다. 변경 후 재배포 필수. */
const ADMIN_KEY = '여기를-나만아는-키로-변경';

/* 정산 앱 URL — 모든 알림 메일에 첨부됨 */
const APP_URL = 'https://leesungwook3165.github.io/sudo_cal/';

/* ── 자동 백업 설정 ──
 * 정산이 저장될 때마다 시트 전체 사본이 "수도정산-백업" 폴더에 날짜별로 생성됩니다.
 * BACKUP_EMAIL에 메일 주소를 넣으면(예: 본인의 다른 계정) 엑셀 파일도 함께 발송되어
 * 구글 계정 밖에도 사본이 남습니다. 비워두면 메일 백업은 생략. */
const BACKUP_FOLDER = '수도정산-백업';
const BACKUP_KEEP = 12;          // 최근 12개 스냅샷 유지 (약 2년치)
const BACKUP_EMAIL = '';         // 예: 'nauk2@naver.com'

function checkKey_(key) {
  return String(key || '') === ADMIN_KEY && ADMIN_KEY !== '여기를-나만아는-키로-변경';
}

/** 시트 스냅샷 백업 (저장 때마다 자동 호출, 수동 실행도 가능) */
function backupNow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const it = DriveApp.getFoldersByName(BACKUP_FOLDER);
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder(BACKUP_FOLDER);
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HHmm');
  const name = '백업 ' + stamp + ' - ' + ss.getName();
  DriveApp.getFileById(ss.getId()).makeCopy(name, folder);

  // 오래된 스냅샷 정리 (최근 BACKUP_KEEP개만 유지)
  const files = [];
  const fi = folder.getFiles();
  while (fi.hasNext()) files.push(fi.next());
  files.sort(function (a, b) { return b.getDateCreated() - a.getDateCreated(); });
  for (let i = BACKUP_KEEP; i < files.length; i++) files[i].setTrashed(true);

  // 메일 백업: 엑셀 사본을 지정 주소로 발송 (구글 계정 밖 오프사이트 사본)
  if (BACKUP_EMAIL) {
    const url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=xlsx';
    const blob = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }
    }).getBlob().setName(name + '.xlsx');
    MailApp.sendEmail(BACKUP_EMAIL, '[수도정산] 자동 백업 ' + stamp,
      '정산 저장 시점의 시트 엑셀 사본입니다.', { attachments: [blob] });
  }
}

/** 정산기록 탭 컬럼 순서 (인덱스는 1-based)
 * 29~40: 입금 추적 컬럼 (v6 추가 — 구버전 시트 자동 보강)
 */
const PAY_COL_BASE = 29; // 29:1층입금여부 30:1층입금일 31:1층실입금액 32:2층… 35:3층… 38:리마인더발송일 39:자동확인로그
const PAY_HEADERS = ['1층입금여부','1층입금일','1층실입금액','2층입금여부','2층입금일','2층실입금액',
  '3층입금여부','3층입금일','3층실입금액','리마인더발송일','자동확인로그'];

/** 정산기록 탭 (없으면 생성) */
function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow([
      '검침일', '이전검침일', '총고지액(원)',
      '1층이전', '1층이번', '1층사용량(㎥)', '1층부담액(원)',
      '2층이전', '2층이번', '2층사용량(㎥)', '2층부담액(원)',
      '3층이전', '3층이번', '3층사용량(㎥)', '3층부담액(원)',
      '추가비용합계(원)', '추가비용내역', '1층최종부담(원)', '2층최종부담(원)', '3층최종부담(원)',
      '고지서사진', '1층계량기사진', '2층계량기사진', '3층계량기사진', '추가증빙사진', '기록시각', '배분방식', '고지서월'
    ].concat(PAY_HEADERS));
    sh.setFrozenRows(1);
    sh.getRange('1:1').setFontWeight('bold');
  }
  // 구버전 시트: 입금 추적 컬럼(29~39)이 없으면 헤더 자동 추가
  if (sh.getLastColumn() < PAY_COL_BASE + PAY_HEADERS.length - 1) {
    for (let i = 0; i < PAY_HEADERS.length; i++) {
      const col = PAY_COL_BASE + i;
      if (sh.getRange(1, col).getValue() !== PAY_HEADERS[i]) sh.getRange(1, col).setValue(PAY_HEADERS[i]);
    }
    sh.getRange('1:1').setFontWeight('bold');
  }
  return sh;
}

/** 구성원 탭 (없으면 생성) — 가족이 직접 연락처를 관리하는 곳 */
function getMemberSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(MEMBER_NAME);
  if (!sh) {
    sh = ss.insertSheet(MEMBER_NAME);
    sh.appendRow(['이름', '층', '이메일', '전화번호', '메일수신(Y/N)', '구분(관리자/구성원)', '계좌']);
    sh.setFrozenRows(1);
    sh.getRange('1:1').setFontWeight('bold');
    // 작성 예시 한 줄
    sh.appendRow(['(예시) 홍길동', '1층', 'example@gmail.com', '010-1234-5678', 'Y', '관리자', '국민 000-00-000000']);
  }
  // 구버전 시트: 구분·계좌 칼럼이 없으면 헤더만 보강
  if (sh.getLastColumn() < 7) {
    sh.getRange(1, 6).setValue('구분(관리자/구성원)');
    sh.getRange(1, 7).setValue('계좌');
    sh.getRange('1:1').setFontWeight('bold');
  }
  return sh;
}

/** 중간검침 탭 (없으면 생성) — 매월 검침값·사진만 기록 */
function getMidSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(MID_NAME);
  if (!sh) {
    sh = ss.insertSheet(MID_NAME);
    sh.appendRow(['검침일', '1층검침(㎥)', '2층검침(㎥)', '3층검침(㎥)',
      '1층사진', '2층사진', '3층사진', '기록시각']);
    sh.setFrozenRows(1);
    sh.getRange('1:1').setFontWeight('bold');
  }
  return sh;
}

/** 구성원 목록 읽기 */
function getMembers_() {
  const sh = getMemberSheet();
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 7).getValues()
    .map(function (r) {
      return {
        name: String(r[0] || ''), floor: String(r[1] || ''),
        email: String(r[2] || ''), phone: String(r[3] || ''),
        mail: String(r[4] || '').trim().toUpperCase() === 'Y',
        role: String(r[5] || '').indexOf('관리') >= 0 ? '관리자' : '구성원',
        account: String(r[6] || '')
      };
    })
    .filter(function (m) { return m.name && m.name.indexOf('(예시)') < 0; });
}

/** 관리자 행의 계좌 (입금 안내용) */
function adminAccount_() {
  const admins = getMembers_().filter(function (m) { return m.role === '관리자' && m.account; });
  return admins.length ? { account: admins[0].account, owner: admins[0].name } : null;
}

/** 관리자 층 (예: '1층') — 입금 대시보드에서 자동 제외 */
function adminFloor_() {
  const admins = getMembers_().filter(function (m) { return m.role === '관리자' && m.floor; });
  return admins.length ? String(admins[0].floor) : '';
}

/** 층 이름을 1/2/3 으로 변환 */
function floorNum_(s) {
  const m = String(s || '').match(/([123])/);
  return m ? Number(m[1]) : 0;
}

/** 금액 안전 변환 — Date나 이상값을 걸러냄
 * 1억원 초과는 오류로 간주 (실용상 그 이상 청구액은 없음)
 */
function safeAmount_(v) {
  if (v instanceof Date) return 0;
  const n = Number(v);
  if (!isFinite(n) || n < 0 || n > 100000000) return 0;
  return n;
}

/** 특정 행에서 각 층의 입금 상태 배열 반환 */
function readDeposits_(row) {
  const base = PAY_COL_BASE - 1;
  const out = [];
  for (let i = 0; i < 3; i++) {
    const st = String(row[base + i * 3] || '');
    out.push({
      floor: i + 1,
      paid: st.indexOf('Y') === 0 || st.indexOf('자동') >= 0 || st === '수동',
      auto: st.indexOf('자동') >= 0,
      when: formatDate_(row[base + i * 3 + 1]),
      amount: safeAmount_(row[base + i * 3 + 2])
    });
  }
  return out;
}

/** 고지서 사진 폴더 */
function getFolder() {
  const it = DriveApp.getFoldersByName(FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
}

/** 하위 폴더 (없으면 생성) */
function subFolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

/** 월별 사진 폴더: 수도정산-고지서 / 2026년 / 06월분 */
function getMonthFolder_(ym) {
  // ym: 'YYYY-MM' (없거나 형식이 다르면 루트에 저장)
  const m = String(ym || '').match(/^(\d{4})-(\d{2})/);
  const root = getFolder();
  if (!m) return root;
  return subFolder_(subFolder_(root, m[1] + '년'), m[2] + '월분');
}

/** GET: 기록 조회는 누구나 가능(가족 뷰어용) · 전화번호·관리자 권한은 키가 맞을 때만 */
function doGet(e) {
  const isAdmin = checkKey_(e && e.parameter && e.parameter.key);
  const sh = getSheet();
  const last = sh.getLastRow();
  const out = { ok: true, admin: isAdmin, prev: null, phones: [], history: [], mids: [] };

  // 중간검침 기록 (매월 검침값·사진)
  const ms = getMidSheet();
  const mlast = ms.getLastRow();
  if (mlast >= 2) {
    out.mids = ms.getRange(2, 1, mlast - 1, 7).getValues().map(function (r) {
      return {
        date: formatDate_(r[0]),
        readings: [Number(r[1]), Number(r[2]), Number(r[3])],
        photos: { m1: String(r[4] || ''), m2: String(r[5] || ''), m3: String(r[6] || '') }
      };
    }).filter(function (m) { return m.date; })
      .sort(function (a, b) { return a.date < b.date ? 1 : -1; }); // 최신이 앞
  }
  if (last >= 2) {
    const lastCol = Math.max(sh.getLastColumn(), PAY_COL_BASE + PAY_HEADERS.length - 1);
    const rows = sh.getRange(2, 1, last - 1, lastCol).getValues();
    out.history = rows.map(function (r) {
      const amt = [Number(r[6]), Number(r[10]), Number(r[14])];
      let finals = [safeAmount_(r[17]), safeAmount_(r[18]), safeAmount_(r[19])];
      if (!(finals[0] || finals[1] || finals[2])) finals = amt; // 구버전 행 호환
      return {
        date: formatDate_(r[0]), prevDate: formatDate_(r[1]), total: Number(r[2]),
        prev: [Number(r[3]), Number(r[7]), Number(r[11])],
        curr: [Number(r[4]), Number(r[8]), Number(r[12])],
        use:  [Number(r[5]), Number(r[9]), Number(r[13])],
        amt: amt,
        extraSum: Number(r[15]) || 0, extraDesc: String(r[16] || ''),
        finals: finals,
        photos: { bill: String(r[20] || ''), f1: String(r[21] || ''),
                  f2: String(r[22] || ''), f3: String(r[23] || '') },
        extraPhotos: String(r[24] || '').split('\n').filter(function (u) { return u; }),
        savedAt: r[25] instanceof Date ? r[25].toISOString() : '',
        split: String(r[25] || '사용량'),
        billMonth: formatMonth_(r[26]) || formatDate_(r[0]).slice(0, 7),
        deposits: readDeposits_(r)
      };
    }).filter(function (h) { return h.date; })
      .reverse(); // 최신 회차가 앞으로
    if (out.history.length) {
      out.prev = { date: out.history[0].date, readings: out.history[0].curr };
    }
  }
  out.adminFloor = adminFloor_();
  // 입금 안내용 계좌 — 가족 화면에도 표시되도록 항상 포함
  const acct = adminAccount_();
  if (acct) { out.account = acct.account; out.accountOwner = acct.owner; }
  if (isAdmin) {
    out.phones = getMembers_().map(function (m) { return m.phone; })
      .filter(function (p) { return p; });
    out.members = getMembers_(); // 앱에서 구성원 편집용
    out.sheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  }
  return json_(out);
}

/** 디버그 로그 시트에 한 줄 추가 (Cloud 로그 안 뜰 때 대안) */
function debugLog_(tag, msg) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName('디버그로그');
    if (!sh) {
      sh = ss.insertSheet('디버그로그');
      sh.appendRow(['시각', '태그', '내용']);
      sh.setFrozenRows(1);
    }
    sh.appendRow([new Date(), String(tag || ''), String(msg || '').slice(0, 500)]);
    // 오래된 로그 정리 (최근 200줄만 유지)
    const last = sh.getLastRow();
    if (last > 202) sh.deleteRows(2, last - 202);
  } catch (e) { /* 로깅 실패해도 무시 */ }
}

/** POST: 정산 저장 → 시트 기록 + 사진 보관 + 메일 발송 */
function doPost(e) {
  try {
    const rawBody = (e && e.postData ? e.postData.contents : '(no body)');
    debugLog_('doPost', 'body=' + rawBody);
    console.log('[doPost] raw body: ' + rawBody);
    // MacroDroid가 카뱅 알림 여러 줄 텍스트를 그대로 넣으면 JSON이 깨지므로 이스케이프 처리
    let d;
    try {
      d = JSON.parse(rawBody);
    } catch (parseErr) {
      const safeBody = rawBody.replace(/\r/g, '\\r').replace(/\n/g, '\\n').replace(/\t/g, '\\t');
      debugLog_('doPost', 'JSON 재파싱 시도 (개행 이스케이프)');
      d = JSON.parse(safeBody);
    }
    debugLog_('doPost', 'kind=' + d.kind + ' keyTail=' + (d.key ? String(d.key).slice(-2) : '(none)'));
    if (!checkKey_(d.key)) { debugLog_('doPost', 'unauthorized'); return json_({ ok: false, error: 'unauthorized', hint: 'key mismatch' }); }

    // 매월 중간 검침: 검침값·사진만 기록 (메일 발송·백업 없음)
    if (d.kind === 'mid') return saveMid_(d);
    if (d.kind === 'midDelete') return deleteMid_(d);
    if (d.kind === 'members') return saveMembers_(d);
    if (d.kind === 'markPaid') return markPaid_(d, false);
    if (d.kind === 'markUnpaid') return markPaid_(d, true);
    if (d.kind === 'deposit') return depositWebhook_(d);

    // 항목별 사진 저장 (구버전 d.img는 고지서로 취급, x0/x1…은 추가비용 증빙)
    const slotName = { bill: '고지서', f1: '1층계량기', f2: '2층계량기', f3: '3층계량기' };
    const imgSrc = d.imgs || (d.img ? { bill: d.img } : {});
    const urls = { bill: '', f1: '', f2: '', f3: '' };
    const extraUrls = [];
    // 귀속월(billMonth) 기준 월별 폴더에 저장 (없으면 검침일의 달)
    const folder = Object.keys(imgSrc).length
      ? getMonthFolder_(d.billMonth || String(d.date || '').slice(0, 7)) : null;

    Object.keys(imgSrc).forEach(function (s) {
      const raw = imgSrc[s];
      if (!raw) return;
      const label = slotName[s] ||
        ((d.imgLabels && d.imgLabels[s]) ? d.imgLabels[s] : s) + '-증빙';
      const base64 = raw.indexOf(',') >= 0 ? raw.split(',')[1] : raw;
      const blob = Utilities.newBlob(
        Utilities.base64Decode(base64), 'image/jpeg',
        d.date + '-' + label + '.jpg');
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      if (slotName[s]) urls[s] = file.getUrl();
      else extraUrls.push(file.getUrl());
    });

    // 추가 비용 요약
    const modeName = { eq: '균등', use: '사용량비율', ppl: '인원수비율' };
    const extras = d.extras || [];
    const extraSum = d.extraSum || 0;
    const extraDesc = extras.map(function (x) {
      return x.name + ' ' + Number(x.amt).toLocaleString() + '원('
        + (modeName[x.mode] || x.mode) + (x.month ? '·' + x.month : '') + ')';
    }).join(', ');
    const finals = d.finals || d.amt;

    // 수정 모드: 기존 사진 유지분 병합 (새 업로드가 있으면 교체)
    const keep = (d.keep && d.keep.photos) || {};
    const keepExtra = (d.keep && d.keep.extraPhotos) || [];
    const fin = {
      bill: urls.bill || keep.bill || '',
      f1: urls.f1 || keep.f1 || '',
      f2: urls.f2 || keep.f2 || '',
      f3: urls.f3 || keep.f3 || ''
    };
    const finExtra = keepExtra.concat(extraUrls);

    const rowVals = [
      d.date, d.prevDate || '', d.total,
      d.prev[0], d.curr[0], d.use[0], d.amt[0],
      d.prev[1], d.curr[1], d.use[1], d.amt[1],
      d.prev[2], d.curr[2], d.use[2], d.amt[2],
      extraSum, extraDesc, finals[0], finals[1], finals[2],
      fin.bill, fin.f1, fin.f2, fin.f3, finExtra.join('\n'), new Date(), d.split || '사용량', d.billMonth || ''
    ];

    const sh = getSheet();
    let updated = false;
    if (d.mode === 'update' && d.origDate) {
      const last = sh.getLastRow();
      if (last >= 2) {
        const dates = sh.getRange(2, 1, last - 1, 1).getValues();
        for (let i = dates.length - 1; i >= 0; i--) {
          if (formatDate_(dates[i][0]) === d.origDate) {
            sh.getRange(i + 2, 1, 1, rowVals.length).setValues([rowVals]);
            updated = true;
            break;
          }
        }
      }
    }
    if (!updated) sh.appendRow(rowVals);

    // 신규 저장일 때만 가족 알림 발송 (수정은 조용히), 스냅샷 백업은 항상
    const mailed = updated ? 0 : notifyMembers_(d, fin.bill);
    try { backupNow(); } catch (be) { /* 백업 실패해도 저장은 유지 */ }
    const phones = getMembers_().map(function (m) { return m.phone; })
      .filter(function (p) { return p; });

    return json_({ ok: true, updated: updated, urls: fin, mailed: mailed, phones: phones });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/** 중간 검침 저장 — 같은 날짜면 교체(기존 사진은 새 업로드 없으면 유지) */
function saveMid_(d) {
  const labels = { m1: '1층계량기', m2: '2층계량기', m3: '3층계량기' };
  const imgs = d.imgs || {};
  const urls = { m1: '', m2: '', m3: '' };
  const hasImg = Object.keys(imgs).some(function (s) { return imgs[s]; });
  const folder = hasImg ? getMonthFolder_(String(d.date || '').slice(0, 7)) : null;

  Object.keys(labels).forEach(function (s) {
    const raw = imgs[s];
    if (!raw) return;
    const base64 = raw.indexOf(',') >= 0 ? raw.split(',')[1] : raw;
    const blob = Utilities.newBlob(
      Utilities.base64Decode(base64), 'image/jpeg',
      d.date + '-' + labels[s] + '-중간검침.jpg');
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    urls[s] = file.getUrl();
  });

  const sh = getMidSheet();
  const row = [d.date, d.readings[0], d.readings[1], d.readings[2],
    urls.m1, urls.m2, urls.m3, new Date()];

  let updated = false;
  const last = sh.getLastRow();
  if (last >= 2) {
    const dates = sh.getRange(2, 1, last - 1, 1).getValues();
    for (let i = dates.length - 1; i >= 0; i--) {
      if (formatDate_(dates[i][0]) === d.date) {
        const old = sh.getRange(i + 2, 5, 1, 3).getValues()[0]; // 기존 사진 유지
        if (!urls.m1 && old[0]) row[4] = old[0];
        if (!urls.m2 && old[1]) row[5] = old[1];
        if (!urls.m3 && old[2]) row[6] = old[2];
        sh.getRange(i + 2, 1, 1, row.length).setValues([row]);
        updated = true;
        break;
      }
    }
  }
  if (!updated) sh.appendRow(row);
  return json_({ ok: true, updated: updated, urls: { m1: row[4], m2: row[5], m3: row[6] } });
}

/** 중간 검침 삭제 (날짜 기준) */
function deleteMid_(d) {
  const sh = getMidSheet();
  const last = sh.getLastRow();
  if (last >= 2) {
    const dates = sh.getRange(2, 1, last - 1, 1).getValues();
    for (let i = dates.length - 1; i >= 0; i--) {
      if (formatDate_(dates[i][0]) === d.date) {
        sh.deleteRow(i + 2);
        return json_({ ok: true, deleted: true });
      }
    }
  }
  return json_({ ok: true, deleted: false });
}

/** 특정 회차(origDate 기준) 행 번호 반환 (없으면 -1) */
function findRowByDate_(sh, origDate) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const dates = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = dates.length - 1; i >= 0; i--) {
    if (formatDate_(dates[i][0]) === origDate) return i + 2;
  }
  return -1;
}

/** 층별 입금 확인/취소 — 관리자 수동, source: 'manual' 또는 '자동' */
function markPaid_(d, unmark) {
  const floor = Number(d.floor);
  if (!(floor >= 1 && floor <= 3)) return json_({ ok: false, error: 'bad floor' });
  const sh = getSheet();
  const row = findRowByDate_(sh, d.origDate);
  if (row < 0) return json_({ ok: false, error: 'row not found' });
  const base = PAY_COL_BASE + (floor - 1) * 3;
  if (unmark) {
    sh.getRange(row, base, 1, 3).setValues([['', '', '']]);
  } else {
    const src = d.source === 'auto' ? '자동' : '수동';
    const when = d.when ? new Date(d.when) : new Date();
    const amt = safeAmount_(d.amount);
    sh.getRange(row, base, 1, 3).setValues([[src, when, amt]]);
    // 관리자에게 확인 메일 (수동일 때는 조용히)
    if (d.source === 'auto') notifyDepositToAdmin_(d, floor, amt, row);
  }
  return json_({ ok: true, deposits: readDeposits_(sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0]) });
}

/** 관리자에게 입금 확인 알림 메일 */
function notifyDepositToAdmin_(d, floor, amount, row) {
  const admin = getMembers_().find(function (m) { return m.role === '관리자' && m.email; });
  if (!admin) return;
  const sh = getSheet();
  const r = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];
  const billMonth = formatMonth_(r[26]) || formatDate_(r[0]).slice(0, 7);
  const mm = Number(String(billMonth).split('-')[1]);
  const deposits = readDeposits_(r);
  const paidCount = deposits.filter(function (x) { return x.paid; }).length;
  const adminF = floorNum_(adminFloor_());
  const target = 3 - (adminF ? 1 : 0);
  const subject = '[수도정산] ' + mm + '월분 ' + floor + '층 입금 확인 (' + paidCount + '/' + target + ')';
  const body = mm + '월분 ' + floor + '층 입금이 확인되었습니다.\n\n'
    + '입금액: ' + Number(amount).toLocaleString() + '원\n'
    + '확인 방식: ' + (d.source === 'auto' ? '카카오뱅크 알림 자동 매칭' : '관리자 수동 확인') + '\n\n'
    + '현재 진행: ' + paidCount + '/' + target + '\n'
    + deposits.map(function (x) {
      if (adminF === x.floor) return x.floor + '층: (관리자층 · 제외)';
      return x.floor + '층: ' + (x.paid ? '✓ 확인 (' + x.when + ')' : '⏳ 대기');
    }).join('\n') + '\n\n'
    + '📱 상세 확인: ' + APP_URL + '\n'
    + '시트: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl();
  MailApp.sendEmail(admin.email, subject, body);
}

/** 카카오뱅크 알림 파싱 — MacroDroid가 알림 텍스트를 그대로 전달
 * 확인된 카뱅 포맷: "07/09 11:29\n입금 100원\n이성욱 → 입출금통장(1703)\n잔액 24,193원"
 * 층수 힌트도 추출: 보내는 사람 표시명·메모에 "2층" 같은 표기가 있으면 우선 사용
 */
function parseKakaoBank_(text) {
  const t = String(text || '');
  if (t.indexOf('입금') < 0) return null; // 출금·이체 알림은 무시
  const amMain = t.match(/입금\s*([\d,]+)\s*원/);
  const am = amMain || t.match(/([\d,]+)\s*원/);
  if (!am) return null;
  const amount = Number(am[1].replace(/,/g, ''));
  if (!amount) return null;
  let sender = '';
  const p0 = t.match(/입금\s*[\d,]+\s*원\s*\n?\s*([가-힣A-Za-z0-9]{2,15})\s*→/);
  const p1 = t.match(/([가-힣A-Za-z]{2,10})님/);
  const p2 = t.match(/(?:입금)\s*[:.]?\s*[\d,]+\s*원\s*\n?\s*([가-힣A-Za-z0-9]{2,15})/);
  const p3 = t.match(/([가-힣A-Za-z]{2,10})\s*[\d,]+\s*원\s*입금/);
  sender = (p0 && p0[1]) || (p1 && p1[1]) || (p2 && p2[1]) || (p3 && p3[1]) || '';
  // 층수 힌트: 텍스트 어디든 "1층/2층/3층" 있으면 추출 (계좌번호 "1703" 같은 건 뒤에 층이 없으니 무관)
  const floorHint = (t.match(/([123])\s*층/) || [])[1];
  return {
    amount: amount,
    sender: sender,
    floorHint: floorHint ? Number(floorHint) : 0,
    rawText: t.slice(0, 200)
  };
}

/** 원 단위 정확 매칭
 * 실입금액이 부담액과 ±20원 이내이면 매치 (은행 수수료 정도 허용)
 * 반올림 매칭은 하지 않음 — 구성원에게 정확한 원 단위로 이체 요청
 * 반환: {match: bool, roundedTo: 0}
 */
function amountMatches_(actual, owed) {
  if (Math.abs(actual - owed) <= 20) return { match: true, roundedTo: 0 };
  return { match: false, roundedTo: 0 };
}

/** 관리자에게 자동 매칭 실패/애매 알림 */
function notifyAdminConfirmNeeded_(parsed, reasonLabel, candidateSummary, finals, members, adminF) {
  const admin = getMembers_().find(function (m) { return m.role === '관리자' && m.email; });
  if (!admin) return;
  const compareRows = [1, 2, 3].map(function (f) {
    if (adminF === f) return f + '층: (관리자층 · 제외)';
    const owed = finals[f - 1];
    const diff = parsed.amount - owed;
    const m = members.find(function (mm) { return floorNum_(mm.floor) === f; });
    const name = m ? m.name : '';
    const sign = diff === 0 ? '정확' : (diff > 0 ? '초과 ' + diff.toLocaleString() + '원' : '부족 ' + Math.abs(diff).toLocaleString() + '원');
    return f + '층 ' + name + ' 부담액 ' + owed.toLocaleString() + '원 (' + sign + ')';
  }).join('\n');
  const body = '입금 알림을 받았지만 ' + reasonLabel + ' 자동 매칭이 안전하지 않아 보류했습니다.\n\n'
    + '입금자: ' + (parsed.sender || '(파싱 실패)') + '\n'
    + '입금액: ' + parsed.amount.toLocaleString() + '원\n\n'
    + '── 각 층 비교 ──\n' + compareRows + '\n'
    + (candidateSummary ? '\n' + candidateSummary + '\n' : '')
    + '\n원문: ' + parsed.rawText + '\n\n'
    + '앱에서 수동으로 확인해 주세요.\n'
    + '📱 앱: ' + APP_URL + '\n'
    + '시트: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl();
  MailApp.sendEmail(admin.email, '[수도정산] 자동 매칭 실패 — 확인 필요 (' + reasonLabel + ')', body);
}

/** MacroDroid 웹훅: 카뱅 알림 텍스트를 받아 자동 매칭 (안전 우선) */
function depositWebhook_(d) {
  debugLog_('deposit', 'text=' + d.text);
  const parsed = parseKakaoBank_(d.text);
  debugLog_('deposit', 'parsed=' + JSON.stringify(parsed));
  if (!parsed) return json_({ ok: false, error: 'parse failed', text: d.text });
  const sh = getSheet();
  const last = sh.getLastRow();
  if (last < 2) { debugLog_('deposit', 'no settlement'); return json_({ ok: false, error: 'no settlement' }); }
  const row = last;
  const r = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];
  const savedAt = r[25] instanceof Date ? r[25].getTime() : 0;
  if (savedAt && Date.now() < savedAt) { debugLog_('deposit', 'before settlement'); return json_({ ok: false, error: 'before settlement' }); }
  const members = getMembers_();
  const adminF = floorNum_(adminFloor_());
  const deposits = readDeposits_(r);
  const finals = [safeAmount_(r[17]), safeAmount_(r[18]), safeAmount_(r[19])];
  debugLog_('deposit', 'adminFloor=' + adminF + ' finals=' + JSON.stringify(finals) + ' paid=' + JSON.stringify(deposits.map(function(x){return x.paid;})));

  // 0단계: 층수 힌트가 있으면 그 층으로 좁힘 ("2층김철수" 같이 이체 시 표시명·메모에 층 표기)
  const allowedFloors = parsed.floorHint
    ? [parsed.floorHint].filter(function (f) { return f !== adminF && !deposits[f - 1].paid && finals[f - 1]; })
    : [1, 2, 3].filter(function (f) { return f !== adminF && !deposits[f - 1].paid && finals[f - 1]; });

  // 1단계: 금액 매칭 (반올림 허용)
  const amtMatch = [];
  const roundInfo = {};
  allowedFloors.forEach(function (f) {
    const m = amountMatches_(parsed.amount, finals[f - 1]);
    if (m.match) { amtMatch.push(f); roundInfo[f] = m.roundedTo; }
  });
  if (!amtMatch.length) {
    const rsn = parsed.floorHint
      ? '층 힌트 ' + parsed.floorHint + '층인데 금액 불일치'
      : '금액 불일치 (반올림 포함 대조)';
    notifyAdminConfirmNeeded_(parsed, rsn, '', finals, members, adminF);
    return json_({ ok: false, error: 'no amount match', parsed: parsed });
  }

  // 2단계: 이름 매칭 (2자 이상만)
  const nameMatch = amtMatch.filter(function (f) {
    return members.some(function (m) {
      return floorNum_(m.floor) === f
        && m.name && m.name.length >= 2
        && parsed.sender && parsed.sender.indexOf(m.name) >= 0;
    });
  });

  // 3단계: 안전 결정
  let chosen = -1;
  let reason = '';
  if (parsed.floorHint) {
    // 층 힌트가 있으면 그 층에 금액이 맞기만 하면 인정 (강한 신호)
    if (amtMatch.length === 1) chosen = amtMatch[0];
    else reason = '층 힌트 있으나 후보 다중';
  } else if (parsed.sender) {
    if (nameMatch.length === 1) chosen = nameMatch[0];
    else if (nameMatch.length === 0) reason = '이름 불일치';
    else reason = '이름·금액 매칭 후보 다중 (' + nameMatch.join(', ') + '층)';
  } else {
    if (amtMatch.length === 1) chosen = amtMatch[0];
    else reason = '입금자명 없음 + 금액 동일 층 다중 (' + amtMatch.join(', ') + '층)';
  }

  debugLog_('deposit', 'allow=' + JSON.stringify(allowedFloors) + ' amt=' + JSON.stringify(amtMatch) + ' name=' + JSON.stringify(nameMatch) + ' chosen=' + chosen + ' reason=' + reason);
  if (chosen > 0) {
    return markPaid_({
      origDate: formatDate_(r[0]),
      floor: chosen,
      amount: parsed.amount,
      source: 'auto',
      when: new Date()
    }, false);
  }

  // 자동 확정 안 함 — 관리자 판단 요청
  const summary = '자동 매칭 보류 사유: ' + reason;
  notifyAdminConfirmNeeded_(parsed, reason, summary, finals, members, adminF);
  return json_({ ok: false, error: 'need review', reason: reason, parsed: parsed });
}

/** 매일 오전 실행 — 3일 넘은 미납 세대에 리마인더 발송
 * 설치: Apps Script 편집기 → 트리거 → 함수 checkOverdue → 시간 기반 → 오전 9~10시
 */
function checkOverdue() {
  const sh = getSheet();
  const last = sh.getLastRow();
  if (last < 2) return;
  const r = sh.getRange(last, 1, 1, sh.getLastColumn()).getValues()[0];
  const savedAt = r[25] instanceof Date ? r[25].getTime() : 0;
  if (!savedAt) return;
  const days = (Date.now() - savedAt) / 86400000;
  if (days < 3) return;
  // 오늘 이미 발송했으면 스킵 (리마인더발송일 = PAY_COL_BASE + 9)
  const remCol = PAY_COL_BASE + 9;
  const remLast = r[remCol - 1];
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  if (remLast instanceof Date && Utilities.formatDate(remLast, Session.getScriptTimeZone(), 'yyyy-MM-dd') === today) return;
  const adminF = floorNum_(adminFloor_());
  const deposits = readDeposits_(r);
  const overdue = deposits.filter(function (x) { return !x.paid && x.floor !== adminF; });
  if (!overdue.length) return;
  const billMonth = formatMonth_(r[26]) || formatDate_(r[0]).slice(0, 7);
  const mm = Number(String(billMonth).split('-')[1]);
  const acct = adminAccount_();
  const members = getMembers_();
  const finals = [safeAmount_(r[17]), safeAmount_(r[18]), safeAmount_(r[19])];
  // 각 미납 세대에 개별 리마인더
  overdue.forEach(function (o) {
    const m = members.find(function (mm) { return floorNum_(mm.floor) === o.floor && mm.email && mm.mail; });
    if (!m) return;
    const amt = finals[o.floor - 1];
    const subj = '[수도정산] ' + mm + '월분 미납 안내 (D+' + Math.floor(days) + ')';
    let body = m.name + '님, 안녕하세요.\n\n' + mm + '월분 수도요금 정산 후 3일이 지났지만 아직 입금이 확인되지 않았습니다.\n\n'
      + '· 부담액: ' + amt.toLocaleString() + '원\n';
    if (acct) body += '· 입금 계좌: ' + acct.account + ' (' + acct.owner + ')\n';
    body += '\n📱 상세 확인: ' + APP_URL;
    body += '\n\n확인 후 회신 부탁드립니다. 감사합니다.';
    MailApp.sendEmail(m.email, subj, body);
  });
  // 관리자 요약 메일
  const admin = getMembers_().find(function (m) { return m.role === '관리자' && m.email; });
  if (admin) {
    const summary = '⚠ ' + mm + '월분 D+' + Math.floor(days) + ' 미납 요약\n\n'
      + overdue.map(function (o) {
        const m = members.find(function (mm) { return floorNum_(mm.floor) === o.floor; });
        return o.floor + '층 ' + (m ? m.name : '') + ' ' + finals[o.floor - 1].toLocaleString() + '원';
      }).join('\n') + '\n\n📱 앱: ' + APP_URL + '\n시트: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl();
    MailApp.sendEmail(admin.email, '[수도정산] ' + mm + '월분 미납 요약 (D+' + Math.floor(days) + ')', summary);
  }
  // 리마인더 발송일 기록
  sh.getRange(last, remCol).setValue(new Date());
}

/** 구성원 명단 저장 — 앱에서 편집한 전체 목록으로 교체 */
function saveMembers_(d) {
  const list = (d.members || []).filter(function (m) { return m && String(m.name || '').trim(); });
  const sh = getMemberSheet();
  const last = sh.getLastRow();
  if (last >= 2) sh.getRange(2, 1, last - 1, sh.getLastColumn()).clearContent();
  if (list.length) {
    const rows = list.map(function (m) {
      return [String(m.name || ''), String(m.floor || ''), String(m.email || ''),
        String(m.phone || ''), m.mail ? 'Y' : 'N',
        m.role === '관리자' ? '관리자' : '구성원', String(m.account || '')];
    });
    sh.getRange(2, 1, rows.length, 7).setValues(rows);
  }
  return json_({ ok: true, count: list.length, members: getMembers_() });
}

/** 구성원에게 정산 내역 메일 발송 (메일수신=Y 대상) */
function notifyMembers_(d, imgUrl) {
  const members = getMembers_().filter(function (m) { return m.mail && m.email; });
  if (!members.length) return 0;

  const mm = Number(String(d.billMonth || d.date).split('-')[1]);
  const sheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  const extras = d.extras || [];
  const finals = d.finals || d.amt;
  const grand = Number(d.total) + Number(d.extraSum || 0);
  const subject = '[수도요금] ' + mm + '월분 정산 안내 — 총 '
    + grand.toLocaleString() + '원';

  let body = mm + '월분 수도요금 정산 결과입니다.\n'
    + '검침 기간: ' + (d.prevDate || '-') + ' → ' + d.date + '\n'
    + '수도 고지금액: ' + Number(d.total).toLocaleString() + '원\n';
  if (extras.length) {
    body += '추가 비용: ' + extras.map(function (x) {
      const mm = x.month ? Number(String(x.month).split('-')[1]) + '월분 ' : '';
      return mm + x.name + ' ' + Number(x.amt).toLocaleString() + '원';
    }).join(', ') + '\n';
  }
  body += '\n';
  for (let i = 0; i < 3; i++) {
    body += (i + 1) + '층: 검침 ' + d.prev[i] + '→' + d.curr[i]
      + ' / 사용량 ' + d.use[i] + '㎥ / 부담액 '
      + Number(finals[i]).toLocaleString() + '원';
    if (extras.length) {
      body += ' (수도 ' + Number(d.amt[i]).toLocaleString() + '원 + 추가 '
        + (Number(finals[i]) - Number(d.amt[i])).toLocaleString() + '원)';
    }
    body += '\n';
  }
  const acct = adminAccount_();
  if (acct) body += '\n입금 계좌: ' + acct.account + ' (' + acct.owner + ')\n';
  body += '\n※ 위 부담액을 <원 단위까지 정확히> 이체해 주시면 자동으로 확인 처리됩니다.\n'
       +  '  (반올림하지 마시고 정확한 금액으로 부탁드립니다)';
  body += '\n\n📱 상세 확인: ' + APP_URL;
  body += '\n상세 내역(시트): ' + sheetUrl;
  if (imgUrl) body += '\n고지서 사진: ' + imgUrl;

  members.forEach(function (m) {
    MailApp.sendEmail(m.email, subject, body);
  });
  return members.length;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function pad2_(n) { return ('0' + Number(n)).slice(-2); }

function formatMonth_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM');
  }
  // '2026-06', '2026. 6', '2026년 6월' 등 어떤 형식이든 YYYY-MM으로
  const m = String(v || '').match(/(\d{4})\D+(\d{1,2})/);
  return m ? m[1] + '-' + pad2_(m[2]) : '';
}

function formatDate_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  // '2026-06-28', '2026. 6. 28', '2026년 6월 28일' 등 → YYYY-MM-DD, 실패 시 ''
  const m = String(v || '').match(/(\d{4})\D+(\d{1,2})\D+(\d{1,2})/);
  return m ? m[1] + '-' + pad2_(m[2]) + '-' + pad2_(m[3]) : '';
}

