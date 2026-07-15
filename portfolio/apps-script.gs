// ============================================================
// 구글시트 → 웹으로 포트폴리오 데이터를 넘겨주는 백엔드 코드
//
// [사용법]
// 1. 구글시트를 새로 만들고 → 확장 프로그램 → Apps Script 열기
// 2. 이 파일 내용을 통째로 붙여넣기
// 3. 아래 TOKEN을 아무 문자열로 바꾸기 (index.html의 TOKEN과 똑같이!)
// 4. setup() 실행  → 시트가 자동으로 만들어짐 (첫 실행 때 권한 승인 필요)
// 5. test() 실행   → 주가가 잘 조회되는지 로그로 확인
// 6. 배포 → 새 배포 → 유형: 웹 앱
//      - 실행 계정: 나
//      - 액세스 권한: 모든 사용자
// 7. 발급된 URL을 index.html의 API_URL에 붙여넣기
//
// ※ 종목 목록은 이 코드가 아니라 '시트'가 정합니다.
//   시트에 적힌 만큼만 웹에 보이고, 종목을 추가해도 코드는 안 고쳐도 됩니다.
// ============================================================

const SHEET_NAME = 'holdings';
const TOKEN = 'JooJu';  // index.html의 TOKEN과 동일하게 맞출 것

// 웹에서 데이터를 읽어갈 때 (GET)
function doGet(e) {
  if (e.parameter.token !== TOKEN) return json({ error: 'unauthorized' });
  return json({ holdings: getHoldingsWithPrices() });
}

// 웹에서 종목을 추가/수정/삭제할 때 (POST)
function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  if (body.token !== TOKEN) return json({ error: 'unauthorized' });

  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);

  if (body.action === 'add') {
    sheet.appendRow([body.ticker, body.name, body.quantity, body.avgPrice, body.buyDate]);

  } else if (body.action === 'delete') {
    const row = findRow(sheet, body.ticker);
    if (row > 0) sheet.deleteRow(row);

  } else if (body.action === 'update') {
    const row = findRow(sheet, body.ticker);
    if (row > 0) sheet.getRange(row, 3, 1, 2).setValues([[body.quantity, body.avgPrice]]);
  }

  return json({ holdings: getHoldingsWithPrices() });
}

// 티커로 시트에서 몇 번째 줄인지 찾기
function findRow(sheet, ticker) {
  const tickers = sheet.getRange('A:A').getValues().flat();
  return tickers.indexOf(ticker) + 1;  // 못 찾으면 0
}

// 1. 시트에서 보유종목을 읽고
// 2. 종목마다 현재가를 가져와서
// 3. 손익까지 계산해서 합쳐 반환
function getHoldingsWithPrices() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  const rows = sheet.getDataRange().getValues().slice(1);  // 헤더 1행 제외

  return rows.filter(row => row[0]).map(function (row) {
    const ticker = row[0];
    const name = row[1];
    const quantity = Number(row[2]);
    const avgPrice = Number(row[3]);
    const buyDate = row[4];

    const currentPrice = getPrice(ticker);
    const value = currentPrice * quantity;
    const cost = avgPrice * quantity;

    return {
      ticker: ticker,
      name: name,
      quantity: quantity,
      avgPrice: avgPrice,
      buyDate: buyDate instanceof Date ? Utilities.formatDate(buyDate, 'Asia/Seoul', 'yyyy-MM-dd') : buyDate,
      currentPrice: currentPrice,
      value: value,
      profit: value - cost,
      profitRate: cost ? ((value - cost) / cost) * 100 : 0,
    };
  });
}

// 현재가 조회 (야후파이낸스)
// 한국주식 티커: 코스피는 .KS, 코스닥은 .KQ 를 종목코드 뒤에 붙임
//   예) 삼성전자 005930.KS / SK하이닉스 000660.KS / 대한항공 003490.KS
// 나중에 다른 API로 바꾸고 싶으면 이 함수만 고치면 됨
function getPrice(ticker) {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(ticker);
  if (cached) return Number(cached);  // 캐시가 있으면 그대로 사용

  try {
    const url = 'https://query1.finance.yahoo.com/v8/finance/chart/' + ticker;
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const price = JSON.parse(res.getContentText()).chart.result[0].meta.regularMarketPrice;

    cache.put(ticker, String(price), 60);  // 60초간 캐싱 (야후 차단 방지)
    return price;
  } catch (err) {
    return 0;  // 조회 실패하면 0원으로 표시
  }
}

// JSON 형태로 응답 만들기
function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// [1단계] setup() - 시트를 자동으로 만들어 줍니다.
//
// holdings 시트가 없으면 만들고, 헤더 + 예시 3종목을 넣습니다.
// 수량과 평단가는 예시값이니 본인 매수 기록으로 고치세요.
//
// 이미 데이터가 있으면 아무것도 하지 않습니다 (덮어쓰기 방지).
// ============================================================
function setup() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

  // 안전장치: 이미 뭔가 적혀 있으면 손대지 않음
  if (sheet.getLastRow() > 0) {
    Logger.log('이미 데이터가 있어서 그대로 두었습니다. 처음부터 다시 만들려면 시트 내용을 직접 지우고 실행하세요.');
    return;
  }

  const rows = [
    ['ticker', 'name', 'quantity', 'avgPrice', 'buyDate'],
    ['005930.KS', '삼성전자', 10, 240000, '2026-03-14'],
    ['000660.KS', 'SK하이닉스', 2, 2250000, '2026-05-02'],
    ['003490.KS', '대한항공', 50, 23800, '2026-06-11'],
  ];
  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);

  sheet.getRange(1, 1, 1, rows[0].length).setFontWeight('bold');  // 헤더 굵게
  sheet.setFrozenRows(1);                                          // 헤더 고정

  Logger.log('시트 준비 완료! 이제 test()를 실행해보세요.');
}

// ============================================================
// [2단계] test() - 배포 전에 이 함수를 실행해보면
// 시트 읽기 + 주가 조회가 잘 되는지 로그로 확인할 수 있습니다.
// ============================================================
function test() {
  const holdings = getHoldingsWithPrices();

  if (holdings.length === 0) {
    Logger.log('시트가 비어 있습니다. setup()을 먼저 실행하세요.');
    return;
  }

  holdings.forEach(function (h) {
    Logger.log('%s (%s) | 현재가 %s원 | 손익 %s원 (%s%%)',
      h.name, h.ticker, h.currentPrice.toLocaleString(),
      Math.round(h.profit).toLocaleString(), h.profitRate.toFixed(2));
  });
}
