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
    ]);
    sh.setFrozenRows(1);
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
    const rows = sh.getRange(2, 1, last - 1, 27).getValues();
    out.history = rows.map(function (r) {
      const amt = [Number(r[6]), Number(r[10]), Number(r[14])];
      let finals = [Number(r[17]), Number(r[18]), Number(r[19])];
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
        split: String(r[25] || '사용량'),
        billMonth: formatMonth_(r[26]) || formatDate_(r[0]).slice(0, 7)
      };
    }).filter(function (h) { return h.date; })
      .reverse(); // 최신 회차가 앞으로
    if (out.history.length) {
      out.prev = { date: out.history[0].date, readings: out.history[0].curr };
    }
  }
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

/** POST: 정산 저장 → 시트 기록 + 사진 보관 + 메일 발송 */
function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (!checkKey_(d.key)) return json_({ ok: false, error: 'unauthorized' });

    // 매월 중간 검침: 검침값·사진만 기록 (메일 발송·백업 없음)
    if (d.kind === 'mid') return saveMid_(d);
    if (d.kind === 'midDelete') return deleteMid_(d);
    if (d.kind === 'members') return saveMembers_(d);

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
