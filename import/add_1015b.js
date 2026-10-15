

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
