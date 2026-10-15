/* =====================================================================
   ShoeBoard.gs — 엘리자벳 슈즈 방송 현황판 (shoeboard.html) 서버
   Code.gs 와 별도 파일. 함수·변수 이름은 전부 sbd 로 시작 (충돌 없음)
   Rhino 문법 (let/const/화살표 금지)

   Code.gs doGet 에 한 줄 필요 (shoe 줄 바로 위):
     if (String(mode).indexOf("sbd") === 0) return sbdRoute_(e, mode);

   모드
     sbdget            : 리스트 수량 + 내가 누른 수 + 주문서 주문 수 + 특가 목록
     sbdtap&k=&d=1|-1  : 누른 수 더하기/빼기  (k = 번호|칼라|사이즈)
     sbdreset&ok=1     : 누른 수 전부 지우기 (현황 탭 비움)
   ===================================================================== */

var SBD = {
  LIST_TAB: "엘리자벳",        // 방송 리스트 (번호·상품명·가격·칼라·225~250)
  TAP_TAB: "엘리자벳현황",     // 내가 누른 수 저장
  ORDER_TAB: "주문접수",
  OID_PREFIX: "S",            // shoes.html 주문번호 접두
  SINCE: "261015"             // 이 날짜(yyMMdd) 이후 주문만 셈 — 방송마다 바꿈
};

function sbdRoute_(e, mode) {
  var p = (e && e.parameter) || {};
  try {
    if (mode === "sbdget") return sbdJson_(sbdGet_());
    if (mode === "sbdtap") return sbdJson_(sbdTap_(String(p.k || ""), Number(p.d || 0)));
    if (mode === "sbdreset") {
      if (String(p.ok) !== "1") return sbdJson_({ ok: false, msg: "확인값 없음" });
      return sbdJson_(sbdReset_());
    }
    return sbdJson_({ ok: false, msg: "unknown mode" });
  } catch (err) {
    return sbdJson_({ ok: false, msg: String(err && err.message || err) });
  }
}

function sbdJson_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function sbdStr_(v) { return String(v == null ? "" : v).replace(/\s+/g, " ").replace(/^\s+|\s+$/g, ""); }
function sbdKey_(v) { return String(v == null ? "" : v).replace(/\s+/g, ""); }

/* ---------- 리스트 (엘리자벳 탭) ---------- */
function sbdList_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SBD.LIST_TAB);
  if (!sh) throw new Error("'" + SBD.LIST_TAB + "' 탭이 없어요");
  var v = sh.getDataRange().getValues();
  if (v.length < 2) return { sizes: [], items: [] };
  var h = v[0], ci = {}, sizes = [], sizeCol = {};
  for (var c = 0; c < h.length; c++) {
    var t = sbdKey_(h[c]);
    if (/^\d{3}$/.test(t)) { sizes.push(t); sizeCol[t] = c; continue; }
    if (t === "번호" || t === "방송번호") ci.num = c;
    else if (t === "상품명") ci.name = c;
    else if (t === "가격" || t === "판매가") ci.price = c;
    else if (t === "칼라" || t === "컬러" || t === "색상") ci.color = c;
    else if (t === "판매상태" || t === "상태") ci.state = c;
  }
  var items = [], byNum = {}, last = { num: "", name: "", price: 0 };
  for (var r = 1; r < v.length; r++) {
    var row = v[r];
    var num = sbdStr_(row[ci.num]).replace(/[^0-9]/g, "");
    var name = ci.name != null ? sbdStr_(row[ci.name]) : "";
    var price = ci.price != null ? Number(String(row[ci.price]).replace(/[^0-9]/g, "")) || 0 : 0;
    var color = ci.color != null ? sbdStr_(row[ci.color]) : "";
    if (!num && !color) continue;
    if (!num) { num = last.num; if (!name) name = last.name; if (!price) price = last.price; }
    if (!num || !color) continue;
    if (ci.state != null && /숨김|삭제/.test(String(row[ci.state]))) continue;
    last = { num: num, name: name, price: price };
    var it = byNum[num];
    if (!it) { it = byNum[num] = { num: Number(num), name: name, price: price, colors: [] }; items.push(it); }
    var q = {};
    for (var s = 0; s < sizes.length; s++) q[sizes[s]] = Number(row[sizeCol[sizes[s]]]) || 0;
    it.colors.push({ color: color, qty: q });
  }
  items.sort(function (a, b) { return a.num - b.num; });
  return { sizes: sizes, items: items };
}

/* ---------- 내가 누른 수 (엘리자벳현황 탭) ---------- */
function sbdTapSheet_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SBD.TAP_TAB);
  if (!sh) {
    sh = ss.insertSheet(SBD.TAP_TAB);
    sh.getRange(1, 1, 1, 6).setValues([["키", "번호", "칼라", "사이즈", "누른수", "마지막"]]);
    sh.setFrozenRows(1);
  }
  return sh;
}
function sbdTaps_(sh) {
  sh = sh || sbdTapSheet_();
  var n = sh.getLastRow(), m = {};
  if (n < 2) return m;
  var v = sh.getRange(2, 1, n - 1, 5).getValues();
  for (var i = 0; i < v.length; i++) {
    var k = String(v[i][0] || "");
    if (k) m[k] = (m[k] || 0) + (Number(v[i][4]) || 0);
  }
  return m;
}
function sbdTap_(k, d) {
  var parts = k.split("|");
  if (parts.length !== 3 || !parts[0] || !parts[2]) return { ok: false, msg: "키 형식 오류" };
  if (d !== 1 && d !== -1) return { ok: false, msg: "d 는 1 또는 -1" };
  k = sbdKey_(parts[0]) + "|" + sbdKey_(parts[1]) + "|" + sbdKey_(parts[2]);
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = sbdTapSheet_();
    var n = sh.getLastRow(), row = 0, cur = 0;
    if (n >= 2) {
      var keys = sh.getRange(2, 1, n - 1, 1).getValues();
      for (var i = 0; i < keys.length; i++) if (String(keys[i][0]) === k) { row = i + 2; break; }
    }
    var now = Utilities.formatDate(new Date(), "Asia/Seoul", "HH:mm:ss");
    if (row) {
      cur = Number(sh.getRange(row, 5).getValue()) || 0;
      cur = Math.max(0, cur + d);
      sh.getRange(row, 5, 1, 2).setValues([[cur, now]]);
    } else if (d > 0) {
      cur = 1;
      sh.appendRow([k, Number(parts[0]) || parts[0], parts[1], String(parts[2]), cur, now]);
    }
    SpreadsheetApp.flush();
    return { ok: true, k: k, n: cur, taps: sbdTaps_(sh) };
  } finally {
    lock.releaseLock();
  }
}
function sbdReset_() {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = sbdTapSheet_();
    var ss = SpreadsheetApp.getActive();
    var name = SBD.TAP_TAB + "_백업" + Utilities.formatDate(new Date(), "Asia/Seoul", "MMdd_HHmm");
    var had = sh.getLastRow() >= 2;
    if (had) {
      sh.copyTo(ss).setName(name);
      sh.getRange(2, 1, sh.getLastRow() - 1, 6).clearContent();
    }
    return { ok: true, backup: had ? name : "" };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 주문서로 들어온 주문 (주문접수 탭) ---------- */
function sbdOrders_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SBD.ORDER_TAB);
  var res = { counts: {}, specials: [], recent: [], orders: 0, pairs: 0 };
  if (!sh || sh.getLastRow() < 2) return res;
  var v = sh.getDataRange().getValues();
  var h = v[0], ci = {};
  for (var c = 0; c < h.length; c++) {
    var t = sbdKey_(h[c]);
    if (ci.items == null && /주문내역|주문상품|상품내역/.test(t)) ci.items = c;
    else if (ci.oid == null && t === "주문번호") ci.oid = c;
    else if (ci.nick == null && /닉네임/.test(t)) ci.nick = c;
    else if (ci.pay == null && t === "결제방법") ci.pay = c;
    else if (ci.paid == null && /입금확인/.test(t)) ci.paid = c;
    else if (ci.time == null && /주문일시|주문시간|타임스탬프|일시/.test(t)) ci.time = c;
  }
  if (ci.items == null || ci.oid == null) throw new Error("주문접수 헤더(주문내역/주문번호)를 못 찾았어요");
  var reShoe = /^\[신발\]\s*(\d+)번\s.*\(([^()\/]+)\/(\d{3})\)x(\d+)/;
  var reOne = /^\[특가\]\s*(.*?)\s*(\d+)원\(([^()\/]*)\/(\d{3})\)x(\d+)/;
  var names = {};
  for (var r = 1; r < v.length; r++) {
    var row = v[r];
    var oid = String(row[ci.oid] || "");
    if (oid.charAt(0) !== SBD.OID_PREFIX || oid.substr(1, 6) < SBD.SINCE) continue;
    if (ci.paid != null && /취소/.test(String(row[ci.paid]))) continue;
    if (ci.pay != null && /대체됨/.test(String(row[ci.pay]))) continue;
    var items = String(row[ci.items] || "").replace(/\s*\[추가주문\]\s*$/, "");
    var nick = ci.nick != null ? sbdStr_(row[ci.nick]) : "";
    var tm = "";
    if (ci.time != null && row[ci.time] instanceof Date) tm = Utilities.formatDate(row[ci.time], "Asia/Seoul", "HH:mm");
    else if (oid.length >= 11) tm = oid.substr(7, 2) + ":" + oid.substr(9, 2);
    var lines = items.split(", "), any = false;
    for (var i = 0; i < lines.length; i++) {
      var L = sbdStr_(lines[i]), m;
      if ((m = reShoe.exec(L))) {
        var k = m[1] + "|" + sbdKey_(m[2]) + "|" + m[3], q = Number(m[4]) || 1;
        res.counts[k] = (res.counts[k] || 0) + q;
        res.pairs += q; any = true;
        res.recent.push({ t: tm, oid: oid, nick: nick, txt: m[1] + "번 " + sbdStr_(m[2]) + "/" + m[3] + " x" + q });
      } else if ((m = reOne.exec(L))) {
        var who = sbdStr_(m[1]);
        var nk = sbdKey_(who);
        names[nk] = (names[nk] || 0) + 1;
        res.specials.push({ t: tm, oid: oid, nick: nick, name: who, price: Number(m[2]) || 0, color: sbdStr_(m[3]), size: m[4] });
        res.pairs += 1; any = true;
        res.recent.push({ t: tm, oid: oid, nick: nick, txt: "특가 " + who + " " + sbdStr_(m[3]) + "/" + m[4] });
      }
    }
    if (any) res.orders++;
  }
  for (var j = 0; j < res.specials.length; j++) res.specials[j].dup = names[sbdKey_(res.specials[j].name)] > 1;
  res.recent = res.recent.slice(-30).reverse();
  return res;
}

function sbdGet_() {
  var L = sbdList_();
  var o = sbdOrders_();
  return {
    ok: true,
    ts: Utilities.formatDate(new Date(), "Asia/Seoul", "HH:mm:ss"),
    since: SBD.SINCE,
    sizes: L.sizes,
    items: L.items,
    taps: sbdTaps_(),
    orders: o.counts,
    specials: o.specials,
    recent: o.recent,
    orderCount: o.orders,
    pairCount: o.pairs
  };
}

/* ---------- 1회용: 10/15 리스트로 엘리자벳 탭 교체 ----------
   편집기에서 sbdImport1015 를 한 번 실행. 기존 탭은 '엘리자벳_백업1015' 로 복사됨.
   다시 실행해도 같은 내용으로 덮어쓰기만 함 (백업은 이미 있으면 건너뜀) */
function sbdImport1015() {
  var H = ["번호", "상품명", "가격", "칼라", "225", "230", "235", "240", "245", "250", "판매상태", "사진"];
  var R = SBD_1015_ROWS;
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SBD.LIST_TAB);
  if (sh) {
    if (!ss.getSheetByName(SBD.LIST_TAB + "_백업1015")) sh.copyTo(ss).setName(SBD.LIST_TAB + "_백업1015");
    sh.clearContents();
  } else {
    sh = ss.insertSheet(SBD.LIST_TAB);
  }
  var out = [H];
  for (var i = 0; i < R.length; i++) out.push(R[i].concat(["판매중", "images/shoes1015/" + R[i][0] + ".jpg"]));
  sh.getRange(1, 1, out.length, H.length).setNumberFormat("@").setValues(out.map(function (r) { return r.map(String); }));
  sh.getRange(2, 1, out.length - 1, 1).setNumberFormat("0").setValues(R.map(function (r) { return [r[0]]; }));
  sh.getRange(2, 3, out.length - 1, 1).setNumberFormat("0").setValues(R.map(function (r) { return [r[2]]; }));
  sh.getRange(2, 5, out.length - 1, 6).setNumberFormat("0").setValues(R.map(function (r) { return r.slice(4, 10); }));
  sh.setFrozenRows(1);
  try { CacheService.getScriptCache().removeAll(["shoes_v1"]); } catch (e) {}
  Logger.log("완료: " + R.length + "행 / 상품 " + R[R.length - 1][0] + "개");
}
var SBD_1015_ROWS = [
  [1, "라운드쉐입 우드솔 롱부츠", 69000, "블랙", 1, 0, 5, 6, 0, 0],
  [1, "라운드쉐입 우드솔 롱부츠", 69000, "브라운", 0, 0, 0, 6, 2, 5],
  [2, "웨스턴 스티치 하프부츠", 59000, "초코", 0, 0, 1, 0, 5, 2],
  [3, "글로시 통굽 로퍼 롱부츠", 69000, "블랙", 1, 1, 1, 0, 4, 0],
  [3, "글로시 통굽 로퍼 롱부츠", 69000, "카키", 2, 4, 4, 8, 7, 5],
  [4, "소프트 퍼 롱부츠", 69000, "블랙", 3, 1, 7, 0, 2, 1],
  [5, "베이직 패치포인트 하프부츠", 59000, "블랙", 0, 4, 6, 7, 4, 0],
  [6, "퍼라인 빈티지벨트 미틀부츠", 59000, "블랙", 2, 8, 6, 7, 10, 4],
  [7, "진주주얼리 앵클부츠 초특가", 49000, "블랙", 3, 1, 13, 4, 0, 0],
  [8, "스웨이드 아몬드 앵글부츠", 59000, "베이지", 2, 4, 6, 7, 4, 2],
  [9, "모아르 스웨이드 앵클부츠", 59000, "브라운", 3, 1, 5, 9, 2, 6],
  [10, "레이스업 레이어드 워커", 59000, "브라운", 1, 4, 6, 8, 7, 3],
  [11, "스웨이드 퍼 어글리 통굽부츠", 59000, "블랙", 0, 3, 3, 8, 9, 6],
  [11, "스웨이드 퍼 어글리 통굽부츠", 59000, "베이지", 4, 5, 7, 9, 9, 8],
  [12, "소가죽 뽀글이 워커부츠", 59000, "블랙", 3, 3, 10, 5, 8, 7],
  [12, "소가죽 뽀글이 워커부츠", 59000, "그레이", 3, 5, 9, 12, 9, 7],
  [13, "소가죽 스퀘어 로퍼 초특가", 39000, "블랙", 1, 2, 4, 0, 0, 0],
  [14, "실버문 포인트 로퍼", 49000, "블랙", 1, 5, 2, 13, 3, 2],
  [14, "실버문 포인트 로퍼", 49000, "베이지", 1, 0, 3, 8, 8, 2],
  [14, "실버문 포인트 로퍼", 49000, "화이트", 3, 1, 8, 12, 6, 3],
  [15, "벨벳 메리제인로퍼 초특가", 39000, "브라운", 3, 1, 2, 0, 0, 3],
  [16, "소가죽 베이직 스판앵클", 69000, "블랙", 0, 0, 2, 7, 2, 5],
  [16, "소가죽 베이직 스판앵클", 69000, "아이보리", 4, 9, 10, 6, 3, 6],
  [17, "스웨이드 베이직 스판앵클", 69000, "스웨이드블랙", 2, 10, 0, 0, 4, 5],
  [17, "스웨이드 베이직 스판앵클", 69000, "스웨이드베이지", 2, 6, 1, 9, 8, 2],
  [18, "와이드벨트 앵클부츠", 49000, "블랙", 5, 6, 3, 4, 5, 0],
  [18, "와이드벨트 앵클부츠", 49000, "아이보리", 3, 2, 4, 6, 6, 1],
  [19, "소가죽 스퀘어 장식 슬립온 초특가", 39000, "블랙", 2, 0, 8, 0, 3, 2],
  [19, "소가죽 스퀘어 장식 슬립온 초특가", 39000, "화이트", 5, 6, 15, 15, 14, 5],
  [20, "소가죽 진주벨트 뮬", 39000, "베이지", 2, 3, 3, 0, 7, 0],
  [20, "소가죽 진주벨트 뮬", 39000, "화이트", 1, 3, 3, 8, 5, 3],
  [21, "소가죽 베이직 핑크포인트 스니커즈 초특가", 39000, "화이트", 4, 9, 11, 17, 15, 8],
  [22, "니트 스판 스니커즈", 39000, "블랙", 3, 2, 6, 2, 0, 0],
  [22, "니트 스판 스니커즈", 39000, "그레이", 1, 3, 5, 3, 0, 1]
];


/* ---------- 1회용: 10/15 추가상품 (23·24번) 엘리자벳 탭 맨 아래에 붙이기 ----------
   이미 있는 번호|칼라 는 건너뜀 (여러 번 실행해도 중복 안 생김) */
function sbdAdd1015b() {
  var R = [
    [23, "쉬크라인힐 아이유 협찬슈즈", 69000, "블랙무광", 4, 0, 28, 40, 5, 2],
    [23, "쉬크라인힐 아이유 협찬슈즈", 69000, "블랙유광", 2, 8, 0, 12, 12, 0],
    [23, "쉬크라인힐 아이유 협찬슈즈", 69000, "크림무광", 1, 5, 4, 8, 16, 0],
    [24, "소가죽 비비안 뮬", 49000, "블랙", 1, 1, 4, 11, 11, 2],
    [24, "소가죽 비비안 뮬", 49000, "카멜", 0, 2, 6, 12, 9, 2]
  ];
  var sh = SpreadsheetApp.getActive().getSheetByName(SBD.LIST_TAB);
  var v = sh.getDataRange().getValues(), h = v[0], have = {}, cN = -1, cC = -1;
  for (var c = 0; c < h.length; c++) { var t = sbdKey_(h[c]); if (t === "번호") cN = c; if (t === "칼라") cC = c; }
  for (var r = 1; r < v.length; r++) have[sbdKey_(v[r][cN]) + "|" + sbdKey_(v[r][cC])] = 1;
  var add = [];
  for (var i = 0; i < R.length; i++) {
    if (have[R[i][0] + "|" + sbdKey_(R[i][3])]) continue;
    add.push(R[i].concat(["판매중", "images/shoes1015/" + R[i][0] + ".jpg"]));
  }
  if (add.length) sh.getRange(sh.getLastRow() + 1, 1, add.length, add[0].length).setValues(add);
  try { CacheService.getScriptCache().remove("shoes_v1"); } catch (e) {}
  Logger.log("추가: " + add.length + "행 (건너뜀 " + (R.length - add.length) + ")");
}
