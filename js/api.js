/* A4 발주 취합 - 공통 API 클라이언트 */
window.A4 = window.A4 || {};
A4.API_URL = 'https://script.google.com/macros/s/AKfycbxwiAdiUkBUyISY9R5fx3SVu2zMrPBM0ahRPpsm-ldRFIEKn_cmb4BH4whM2-vXG1_knw/exec';
A4.SITE_URL = 'https://ahndasom91-ux.github.io/a4/';

/** Apps Script 웹앱 호출. text/plain 으로 보내 CORS preflight 를 피한다. */
A4.api = async function (action, payload) {
  const body = JSON.stringify(Object.assign({ action: action }, payload || {}));
  let res;
  try {
    res = await fetch(A4.API_URL, { method: 'POST', body: body, headers: { 'Content-Type': 'text/plain;charset=utf-8' }, redirect: 'follow' });
  } catch (e) {
    throw new Error('서버에 연결할 수 없습니다. 네트워크 상태를 확인해 주세요.');
  }
  let text = await res.text();
  let json;
  try { json = JSON.parse(text); }
  catch (e) {
    if (/Authorization|권한|승인|Google/i.test(text) && res.status >= 400) throw new Error('서버가 아직 승인되지 않았습니다. 담당자가 Apps Script에서 setup()을 실행해 권한을 승인해야 합니다.');
    throw new Error('서버 응답을 해석할 수 없습니다 (HTTP ' + res.status + ').');
  }
  if (!json.ok) {
    const err = new Error(json.error || '알 수 없는 오류');
    if (json.auth === false) err.auth = false;
    throw err;
  }
  return json;
};

/* ───────── 공통 유틸 ───────── */
A4.q = function (sel, root) { return (root || document).querySelector(sel); };
A4.qa = function (sel, root) { return Array.prototype.slice.call((root || document).querySelectorAll(sel)); };
A4.esc = function (s) { return String(s === null || s === undefined ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); };
A4.fmtDT = function (iso) {
  if (!iso) return '';
  const m = String(iso).match(/^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2}))?/);
  if (!m) return String(iso);
  const d = new Date(+m[1], +m[2] - 1, +m[3]);
  const wd = ['일', '월', '화', '수', '목', '금', '토'][d.getDay()];
  return m[2] + '/' + m[3] + '(' + wd + ')' + (m[4] ? ' ' + m[4] + ':' + m[5] : '');
};
A4.pathOf = function (r) { return [r['본부'], r['사업단'], r['지점']].filter(Boolean); };
A4.norm = function (s) { return String(s || '').toLowerCase().replace(/\s+/g, ''); };
/** '김경순/백지연', '박유경 ,김민경', '송은정2(대직)' → ['김경순','백지연'] 등 */
A4.splitNames = function (s) {
  return String(s || '').split(/[\/,，·]|\s{2,}|\s,\s?/).map(function (x) { return x.replace(/\(.*?\)/g, '').trim(); }).filter(Boolean);
};
A4.toast = function (msg, kind) {
  let el = A4.q('#toast');
  if (!el) { el = document.createElement('div'); el.id = 'toast'; document.body.appendChild(el); }
  el.textContent = msg; el.className = 'show ' + (kind || '');
  clearTimeout(A4._toastT); A4._toastT = setTimeout(function () { el.className = ''; }, 3200);
};
A4.copy = async function (text) {
  try { await navigator.clipboard.writeText(text); A4.toast('복사했습니다'); }
  catch (e) { window.prompt('복사해 주세요', text); }
};
A4.statusLabel = { '진행중': '진행중', '마감': '마감 · 읽기 전용', '발주완료': '발주요청 완료' };
