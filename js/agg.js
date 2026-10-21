/**
 * A4 발주 집계 모듈 - 브라우저(SheetJS 전역 XLSX)와 Node(require('xlsx')) 양쪽에서 동작.
 *
 * 입력: 조직현황 / 신계약정산평가업적 워크북(들)  →  출력: 지점별 신계약건수·산정박스 + 경고
 *
 * 현행 엑셀 수식 재현:
 *   L = IFERROR(SUMIFS(정산!Q:Q, 정산!F:F, E) / COUNTIFS(E:E, E), 0)
 *   M = ROUNDUP(L/30, 0)
 * 개선점: 지점명 앞뒤 공백 제거(trim) 후 매칭, 별칭 테이블, 매칭 실패를 경고로 노출.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.A4Agg = factory();
})(typeof self !== 'undefined' ? self : this, function () {

  var ORG_REQUIRED = ['부문', '본부', '사업단', '지점', '담당주임단', '주소'];
  var ORG_OPTIONAL = ['관리자', '직책', '전화번호', 'FAX'];
  var PERF_BRANCH_COL = '지점/센터';
  var DEFAULT_COUNT_COL = '생명/장기건수(신)';

  function s(v) { return v === null || v === undefined ? '' : String(v).trim(); }

  /** 지점명 정규화: 끝의 분할 글자(A/B/C/D) 제거. 현행 E열 규칙 */
  function normalizeBranch(name) {
    var x = s(name);
    if (!x) return '';
    // 샛별지점B(서부1) → 샛별지점(서부1)
    x = x.replace(/^(.*지점)[A-Z](\(.*\))$/, '$1$2');
    // 다송지점B → 다송지점 / 블링지점(GRP)B → 블링지점(GRP)
    x = x.replace(/^(.*(?:지점|\)))[A-Z]$/, '$1');
    return x;
  }

  /** 원본 지점명에서 분할 표식(B/C/D)만 뽑기 */
  function splitLabel(name) {
    var x = s(name);
    var m = x.match(/지점([A-Z])\(.*\)$/) || x.match(/(?:지점|\))([A-Z])$/);
    return m ? m[1] : '';
  }

  function sheetRows(XLSX, ws) {
    return XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
  }

  function findHeaderRow(rows, required, maxScan) {
    for (var i = 0; i < Math.min(rows.length, maxScan || 15); i++) {
      var cells = (rows[i] || []).map(s);
      var ok = required.every(function (h) { return cells.indexOf(h) >= 0; });
      if (ok) return i;
    }
    return -1;
  }

  function colIndex(headerCells, name) {
    for (var i = 0; i < headerCells.length; i++) if (s(headerCells[i]) === name) return i;
    return -1;
  }

  /**
   * 워크북 배열에서 조직현황 시트와 정산 시트를 찾는다.
   * 조직현황 후보가 여럿이면 이름이 정확히 '조직현황'인 것을 우선, 그 다음 첫 번째.
   */
  function detectSheets(XLSX, workbooks, countCol) {
    countCol = countCol || DEFAULT_COUNT_COL;
    var orgCands = [], perfCands = [];
    workbooks.forEach(function (wb, wi) {
      wb.SheetNames.forEach(function (name) {
        var rows = sheetRows(XLSX, wb.Sheets[name]);
        var h = findHeaderRow(rows, ORG_REQUIRED);
        if (h >= 0) orgCands.push({ wb: wi, name: name, rows: rows, header: h });
        var hp = findHeaderRow(rows, [PERF_BRANCH_COL, countCol], 5);
        if (hp >= 0) perfCands.push({ wb: wi, name: name, rows: rows, header: hp });
      });
    });
    var org = orgCands.filter(function (c) { return c.name === '조직현황'; })[0] || orgCands[0] || null;
    var perf = perfCands.filter(function (c) { return c.name === '정산평가업적'; })[0] || perfCands[0] || null;
    return { org: org, perf: perf, orgCandidates: orgCands.map(function (c) { return c.name; }), perfCandidates: perfCands.map(function (c) { return c.name; }) };
  }

  /**
   * 집계 실행.
   * @param XLSX     SheetJS
   * @param workbooks 워크북 배열 (1개 또는 2개)
   * @param opts     { aliases:{조직현황명:정산명}, divisor:30, countCol:'생명/장기건수(신)', trim:true }
   */
  function aggregate(XLSX, workbooks, opts) {
    opts = opts || {};
    var aliases = opts.aliases || {};
    var divisor = Number(opts.divisor || 30);
    var countCol = opts.countCol || DEFAULT_COUNT_COL;
    var trim = opts.trim !== false;

    var det = detectSheets(XLSX, workbooks, countCol);
    if (!det.org) throw new Error('조직현황 시트를 찾을 수 없습니다. 머리글에 ' + ORG_REQUIRED.join('·') + ' 열이 있어야 합니다.');
    if (!det.perf) throw new Error('신계약정산평가업적 시트를 찾을 수 없습니다. 머리글에 "' + PERF_BRANCH_COL + '"과 "' + countCol + '" 열이 있어야 합니다.');

    // ── 정산: 지점명별 건수 합계
    var ph = det.perf.rows[det.perf.header].map(s);
    var pBranch = colIndex(ph, PERF_BRANCH_COL), pCount = colIndex(ph, countCol);
    var pName = colIndex(ph, '업적대상자');
    var perfSum = {}, perfPeople = {}, perfRowCount = 0;
    for (var i = det.perf.header + 1; i < det.perf.rows.length; i++) {
      var pr = det.perf.rows[i] || [];
      var bn = pr[pBranch];
      if (bn === null || bn === undefined || String(bn) === '') continue;
      var key = trim ? s(bn) : String(bn);
      var n = Number(pr[pCount]); if (!isFinite(n)) n = 0;
      perfSum[key] = (perfSum[key] || 0) + n;
      perfPeople[key] = (perfPeople[key] || 0) + 1;
      perfRowCount++;
    }

    // ── 조직현황: 행 추출
    var oh = det.org.rows[det.org.header].map(s);
    var c = {};
    ORG_REQUIRED.concat(ORG_OPTIONAL).forEach(function (h) { c[h] = colIndex(oh, h); });
    var rows = [];
    for (var r = det.org.header + 1; r < det.org.rows.length; r++) {
      var or = det.org.rows[r] || [];
      var get = function (h) { return c[h] >= 0 ? s(or[c[h]]) : ''; };
      var 본부 = get('본부'), 사업단 = get('사업단'), 지점 = get('지점'), 부문 = get('부문');
      if (!본부 && !사업단 && !지점 && !부문) continue;
      var kind = 지점 ? '지점' : '단위';
      var norm = 지점 ? normalizeBranch(지점) : '';
      var canon = norm && aliases[norm] ? s(aliases[norm]) : norm;
      rows.push({
        seq: rows.length + 1, kind: kind,
        '부문': 부문, '본부': 본부, '사업단': 사업단, '지점': 지점, '지점_정규화': norm, _canon: canon, _split: splitLabel(지점),
        '관리자': get('관리자'), '직책': get('직책'), '담당주임단': get('담당주임단'),
        '전화번호': get('전화번호'), 'FAX': get('FAX'), '주소': get('주소')
      });
    }

    // ── 분할 지점 수 (정규화명 기준, 현행 COUNTIFS(E:E,E) 재현) - 별칭은 정규화명에 적용된 canon 기준으로 센다
    var dup = {};
    rows.forEach(function (x) { if (x.kind === '지점') dup[x._canon] = (dup[x._canon] || 0) + 1; });

    // ── 집계
    var warnings = { unmatched: [], missingManager: [], missingAddr: [], perfOnly: [] };
    var totals = { branchRows: 0, unitRows: 0, boxes: 0, count: 0, matched: 0, unmatched: 0, split: 0 };
    var usedPerf = {};
    rows.forEach(function (x) {
      x.row_id = 'r' + x.seq;
      if (x.kind !== '지점') {
        totals.unitRows++;
        x['신계약건수'] = ''; x['산정박스'] = ''; x['분할수'] = ''; x['매칭상태'] = '';
        if (!x['담당주임단']) warnings.missingManager.push(label(x));
        if (!x['주소']) warnings.missingAddr.push(label(x));
        return;
      }
      totals.branchRows++;
      var matched = Object.prototype.hasOwnProperty.call(perfSum, x._canon);
      var sum = matched ? perfSum[x._canon] : 0;
      var share = sum / dup[x._canon];
      var boxes = share > 0 ? Math.ceil(share / divisor - 1e-9) : 0;
      x['신계약건수'] = Math.round(share * 100) / 100;
      x['산정박스'] = boxes;
      x['분할수'] = dup[x._canon];
      x['매칭상태'] = matched ? '매칭' : '매칭안됨';
      if (matched) { usedPerf[x._canon] = true; totals.matched++; } else { totals.unmatched++; }
      if (dup[x._canon] > 1) totals.split++;
      totals.boxes += boxes; totals.count += share;
      if (!matched) warnings.unmatched.push({ row_id: x.row_id, '지점': x['지점'], '정규화': x['지점_정규화'], '본부': x['본부'], suggestions: suggest(x['지점_정규화'], perfSum) });
      if (!x['담당주임단']) warnings.missingManager.push(label(x));
      if (!x['주소']) warnings.missingAddr.push(label(x));
    });
    Object.keys(perfSum).forEach(function (k) { if (!usedPerf[k]) warnings.perfOnly.push({ name: k, count: perfSum[k], people: perfPeople[k] }); });
    warnings.perfOnly.sort(function (a, b) { return b.count - a.count; });
    totals.count = Math.round(totals.count * 100) / 100;

    rows.forEach(function (x) { delete x._canon; x['분할표식'] = x._split; delete x._split; });

    return {
      rows: rows, warnings: warnings, totals: totals,
      meta: { orgSheet: det.org.name, perfSheet: det.perf.name, perfRows: perfRowCount, divisor: divisor, countCol: countCol, aliases: aliases,
              orgCandidates: det.orgCandidates, perfCandidates: det.perfCandidates }
    };
  }

  function label(x) { return [x['본부'], x['사업단'], x['지점']].filter(Boolean).join(' › '); }

  /** 매칭 실패 지점에 대해 정산 지점명 후보 제안 */
  function suggest(norm, perfSum) {
    var base = norm.replace(/\(.*?\)/g, '').replace(/지점$/, '');
    if (!base) return [];
    var out = [];
    Object.keys(perfSum).forEach(function (k) {
      var kb = k.replace(/\(.*?\)/g, '').replace(/지점$/, '');
      if (kb === base || k.indexOf(base) >= 0 || (base.length >= 2 && kb.indexOf(base) >= 0)) out.push(k);
    });
    return out.slice(0, 5);
  }

  return { normalizeBranch: normalizeBranch, splitLabel: splitLabel, detectSheets: detectSheets, aggregate: aggregate, DEFAULT_COUNT_COL: DEFAULT_COUNT_COL };
});
