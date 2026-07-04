// ============================================================
//  플러스프레스 BI/CI 실시간 보고서 시스템
//  설문 수집 + 변리사 의견 + 고객 보고서 API
//
//  [스크립트 속성 설정]
//  SHEET_ID         : Google Sheets ID
//  TELEGRAM_TOKEN   : 텔레그램 봇 토큰 (선택)
//  TELEGRAM_CHAT_ID : 알림받을 chat_id (선택)
// ============================================================

function getConfig() {
  const p = PropertiesService.getScriptProperties();
  return {
    sheetId: p.getProperty('SHEET_ID'),
    tgToken: p.getProperty('TELEGRAM_TOKEN'),
    tgChat: p.getProperty('TELEGRAM_CHAT_ID'),
  };
}

function ss() { return SpreadsheetApp.openById(getConfig().sheetId); }

// ── 시트 초기화 ────────────────────────────────────────────
function initSheets() {
  const book = ss();

  // ① 업체목록 — 이영지님이 직접 관리하는 마스터 탭
  //    상태: 1=수요조사 2=방향서 3=검토·시안 4=납품·출원
  let s = getOrCreate(book, '업체목록');
  s.clearContents();
  s.getRange(1,1,1,9).setValues([[
    '업체코드','업체명','접근키','상태(1~4)','방향서','컬러방향','서체방향','무드요약','메모'
  ]]).setFontWeight('bold').setBackground('#EEEDFE');
  // 샘플 행
  s.getRange(2,1,1,4).setValues([['A01','샘플 브랜드', makeKey(), 1]]);

  // ② 설문응답
  s = getOrCreate(book, '설문응답');
  s.clearContents();
  s.getRange(1,1,1,19).setValues([[
    '시각','업체코드','브랜드명','슬로건','업종','출원형태','로고유형',
    '모티프','스타일','느낌문장','컬러','서체','활용처','요청사항',
    '중점사항','무드수치','경쟁사URL','레퍼런스URL','레퍼런스이미지'
  ]]).setFontWeight('bold').setBackground('#E1F5EE');

  // ③ 변리사의견 — '반영상태' 컬럼을 이영지님이 '반영완료'로 바꾸면 보고서에 배지 표시
  s = getOrCreate(book, '변리사의견');
  s.clearContents();
  s.getRange(1,1,1,6).setValues([[
    '시각','업체코드','작성자','구분','의견','반영상태'
  ]]).setFontWeight('bold').setBackground('#FAEEDA');

  // ④ 시안 — 이미지URL은 공개 이미지 링크 (GitHub raw, Drive 공개링크 등)
  s = getOrCreate(book, '시안');
  s.clearContents();
  s.getRange(1,1,1,5).setValues([[
    '업체코드','시안명','이미지URL','설명','상태(검토중/확정)'
  ]]).setFontWeight('bold').setBackground('#E6F1FB');

  Browser.msgBox('초기화 완료!\n업체목록 탭에 업체 10곳을 등록하세요.\n접근키는 makeKeys() 실행으로 일괄 생성됩니다.');
}

function getOrCreate(book, name) {
  return book.getSheetByName(name) || book.insertSheet(name);
}

// ── 접근키 생성 도우미 ─────────────────────────────────────
function makeKey() {
  return Utilities.getUuid().replace(/-/g,'').slice(0,8);
}

// 업체목록의 접근키 빈 칸을 일괄 채움
function makeKeys() {
  const s = ss().getSheetByName('업체목록');
  const rows = s.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] && !rows[i][2]) {
      s.getRange(i+1, 3).setValue(makeKey());
    }
  }
}

// 업체별 링크 목록 출력 (로그에서 복사)
function printLinks() {
  const BASE_REPORT = 'https://YOUR-GITHUB-PAGES/report.html';
  const BASE_SURVEY = 'https://YOUR-GITHUB-PAGES/index.html';
  const BASE_ATTORNEY = 'https://YOUR-GITHUB-PAGES/attorney.html';
  const rows = ss().getSheetByName('업체목록').getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const [code, name, key] = rows[i];
    if (!code) continue;
    Logger.log(`\n■ ${name} (${code})`);
    // 설문 링크에 k(접근키)를 함께 넣어야 제출 완료 후 "결과 확인하기" 버튼이 만들어집니다
    Logger.log(`  설문:   ${BASE_SURVEY}?c=${code}&k=${key}`);
    Logger.log(`  보고서: ${BASE_REPORT}?c=${code}&k=${key}`);
    Logger.log(`  변리사: ${BASE_ATTORNEY}?c=${code}`);
  }
}

// ── GET: 보고서 데이터 API ─────────────────────────────────
function doGet(e) {
  const code = (e.parameter.c || '').toUpperCase();
  const key = e.parameter.k || '';

  const out = obj => ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);

  if (!code || !key) return out({ error: '잘못된 요청입니다' });

  const book = ss();
  const master = book.getSheetByName('업체목록').getDataRange().getValues();
  let company = null;
  for (let i = 1; i < master.length; i++) {
    if (String(master[i][0]).toUpperCase() === code) {
      if (String(master[i][2]) !== key) return out({ error: '접근키가 올바르지 않습니다' });
      company = {
        code: code,
        name: master[i][1],
        status: master[i][3] || 1,
        direction: master[i][4] || '',
        colorDir: master[i][5] || '',
        fontDir: master[i][6] || '',
        moodDir: master[i][7] || ''
      };
      break;
    }
  }
  if (!company) return out({ error: '등록되지 않은 업체 코드입니다' });

  // 설문응답 (해당 코드의 최신 1건)
  const sv = book.getSheetByName('설문응답').getDataRange().getValues();
  const svHead = sv[0];
  let survey = {};
  for (let i = sv.length - 1; i >= 1; i--) {
    if (String(sv[i][1]).toUpperCase() === code) {
      for (let j = 2; j < svHead.length; j++) {
        if (sv[i][j]) survey[svHead[j]] = sv[i][j];
      }
      break;
    }
  }

  // 변리사 의견 (전체, 최신순)
  const op = book.getSheetByName('변리사의견').getDataRange().getValues();
  const attorney = [];
  for (let i = 1; i < op.length; i++) {
    if (String(op[i][1]).toUpperCase() === code) {
      attorney.push({
        time: formatTime(op[i][0]),
        author: op[i][2], category: op[i][3],
        opinion: op[i][4], applied: op[i][5] || ''
      });
    }
  }
  attorney.reverse();

  // 시안
  const dr = book.getSheetByName('시안').getDataRange().getValues();
  const drafts = [];
  for (let i = 1; i < dr.length; i++) {
    if (String(dr[i][0]).toUpperCase() === code) {
      drafts.push({
        name: dr[i][1], url: dr[i][2] || '',
        desc: dr[i][3] || '', status: dr[i][4] || '검토중'
      });
    }
  }

  return out({ company, survey, attorney, drafts });
}

function formatTime(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, 'Asia/Seoul', 'M/d HH:mm');
  }
  return String(v);
}

// ── POST: 설문 / 변리사 의견 기록 ──────────────────────────
function doPost(e) {
  const out = obj => ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);

  try {
    const d = JSON.parse(e.postData.contents);

    if (d.type === 'survey') {
      const imgLinks = saveRefImages(d.code, d.refImages || []);
      ss().getSheetByName('설문응답').appendRow([
        d.timestamp || new Date(), (d.code||'').toUpperCase(),
        d.brandName||'', d.slogan||'', d.industry||'', d.markType||'',
        d.logoType||'', d.motif||'', d.style||'', d.feeling||'',
        d.color||'', d.font||'', d.usage||'', d.request||'',
        d.care||'', d.mood||'', d.competitors||'', d.refUrl||'',
        imgLinks.join(', ')
      ]);
      notify(`📋 *설문 응답 접수*\n업체: ${d.code} ${d.brandName||''}\n출원: ${d.markType||'-'} / 로고: ${d.logoType||'-'}\n경쟁사: ${d.competitors||'없음'}\n→ 방향서 작성 차례입니다`);
      return out({ status: 'ok' });
    }

    if (d.type === 'attorney') {
      ss().getSheetByName('변리사의견').appendRow([
        d.timestamp || new Date(), (d.code||'').toUpperCase(),
        d.author||'', d.category||'', d.opinion||'', ''
      ]);
      notify(`⚖️ *변리사 의견 등록*\n업체: ${d.code}\n작성: ${d.author}\n구분: ${d.category}\n───\n${d.opinion}\n\n→ 반영 후 '반영상태'를 반영완료로 변경하세요`);
      return out({ status: 'ok' });
    }

    return out({ status: 'error', msg: 'unknown type' });

  } catch(err) {
    return out({ status: 'error', msg: err.message });
  }
}

// ── 레퍼런스 이미지 Drive 저장 ──────────────────────────────
function saveRefImages(code, images) {
  if (!images || !images.length) return [];
  const folder = getRefFolder();
  const codeFolder = getOrCreateFolder(folder, (code || 'UNKNOWN').toUpperCase());
  const links = [];

  images.forEach((img, i) => {
    try {
      const match = /^data:(image\/\w+);base64,(.+)$/.exec(img.dataUrl);
      if (!match) return;
      const contentType = match[1];
      const bytes = Utilities.base64Decode(match[2]);
      const blob = Utilities.newBlob(bytes, contentType, img.name || ('ref_'+i+'.png'));
      const file = codeFolder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      links.push('https://drive.google.com/uc?id=' + file.getId());
    } catch(e) {
      console.error('이미지 저장 실패:', e.message);
    }
  });
  return links;
}

function getRefFolder() {
  const name = 'BICI_레퍼런스이미지';
  const it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}

function getOrCreateFolder(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

// ── 텔레그램 알림 ──────────────────────────────────────────
function notify(text) {
  try {
    const c = getConfig();
    if (!c.tgToken || !c.tgChat) return;
    UrlFetchApp.fetch(`https://api.telegram.org/bot${c.tgToken}/sendMessage`, {
      method: 'post',
      payload: { chat_id: c.tgChat, text: text, parse_mode: 'Markdown' }
    });
  } catch(e) { console.error(e.message); }
}
