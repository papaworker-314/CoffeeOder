// 단체 음료 주문 저장소 — 구글 시트 Apps Script
// ▼ 관리자 PIN (집계 화면에서 삭제·초기화할 때 사용). 원하는 숫자로 바꾸세요.
const ADMIN_PIN = '0000';
const SHEET = '주문';
const HEAD = ['id', 'ts', '주문시간', '이름', '메뉴', 'HOT/ICE', '사이즈', '디카페인', '추가요청', '수정시간'];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(SHEET);
  if (!s) {
    s = ss.insertSheet(SHEET);
    s.appendRow(HEAD);
    s.setFrozenRows(1);
  } else if (s.getRange(1, 9).getValue() !== '추가요청') {
    s.getRange(1, 1, 1, HEAD.length).setValues([HEAD]); // 이전 버전 시트 머리글 보완
  }
  return s;
}
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function rowOf_(s, id) {
  const ids = s.getRange(1, 1, s.getLastRow(), 1).getValues().flat().map(String);
  const i = ids.indexOf(String(id));
  return i > 0 ? i + 1 : -1;
}
function clean_(o) {
  return {
    name: String(o.name || '').trim().slice(0, 30),
    menu: String(o.menu || '').slice(0, 60),
    temp: o.temp === 'HOT' ? 'HOT' : 'ICE',
    size: ['톨', '그란데', '벤티'].indexOf(o.size) > -1 ? o.size : '톨',
    dc: o.dc ? 'Y' : '',
    note: String(o.note || '').trim().slice(0, 100)
  };
}

// 주문 목록 읽기
function doGet() {
  const v = sheet_().getDataRange().getValues();
  v.shift();
  const orders = v.filter(r => r[0]).map(r => ({
    id: String(r[0]), ts: Number(r[1]), name: String(r[3]), menu: String(r[4]),
    temp: String(r[5]), size: String(r[6] || '톨'), dc: r[7] === 'Y' ? 1 : 0, note: String(r[8] || '')
  }));
  return out_({ ok: true, orders });
}

// 주문 추가 / 수정 / 취소 / (관리자) 삭제 / 초기화
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const s = sheet_();

    if (d.action === 'add') {
      const o = d.order || {}, c = clean_(o);
      if (!c.name || !c.menu) return out_({ ok: false, error: 'bad order' });
      if (rowOf_(s, o.id) === -1) {
        s.appendRow([String(o.id), Number(o.ts) || Date.now(), new Date(), c.name, c.menu, c.temp, c.size, c.dc, c.note, '']);
      }
      return out_({ ok: true });
    }
    if (d.action === 'update') {
      const o = d.order || {}, c = clean_(o);
      const r = rowOf_(s, o.id);
      if (r === -1) return out_({ ok: false, error: 'notfound' });
      if (!c.name || !c.menu) return out_({ ok: false, error: 'bad order' });
      s.getRange(r, 4, 1, 7).setValues([[c.name, c.menu, c.temp, c.size, c.dc, c.note, new Date()]]);
      return out_({ ok: true });
    }
    if (d.action === 'cancel') {
      const r = rowOf_(s, d.id);
      if (r > 0) s.deleteRow(r);
      return out_({ ok: true });
    }

    // 아래는 관리자 전용
    if (String(d.pin) !== ADMIN_PIN) return out_({ ok: false, error: 'pin' });
    if (d.action === 'delete') {
      const r = rowOf_(s, d.id);
      if (r > 0) s.deleteRow(r);
      return out_({ ok: true });
    }
    if (d.action === 'reset') {
      if (s.getLastRow() > 1) s.deleteRows(2, s.getLastRow() - 1);
      return out_({ ok: true });
    }
    return out_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
