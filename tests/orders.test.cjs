const { test } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const XLSX = require("xlsx");
const { JSDOM } = require("jsdom");
const app = require("../app.js");
const headers = [
  "주문서번호",
  "사과상품선택",
  "사과 수량선택",
  "주문자이름",
  "연락처",
  "배송지주소",
  "입금자/총입금액",
  "주문서내용확인",
  "개인정보수집동의",
];
const rows = [
  headers,
  [
    2,
    "5키로 (배송비 포함 5만원)",
    "1. 1박스",
    "테스트 가",
    1020556927,
    "테스트 주소 가",
    "테스트 / 50000",
    "확인",
    "동의",
  ],
  [],
  [
    10,
    "10키로",
    "2. 2박스",
    "테스트 나",
    "01000000001",
    "테스트 주소 나",
    "테스트 / 100000",
    "확인",
    "동의",
  ],
  [
    15,
    "5키로",
    "1박스",
    "테스트 다",
    "01000000002",
    "테스트 주소 다",
    "테스트 / 50000",
  ],
];
function workbook() {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    wb,
    XLSX.utils.aoa_to_sheet(Array.from({ length: 100 }, (_, i) => ["통계", i])),
    "통계",
  );
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), "설문 답변");
  return wb;
}
test("주문 시트 선택, 빈 행 제외, 데이터 정리", () => {
  const res = app.processWorkbook(workbook());
  assert.equal(res.sheetName, "설문 답변");
  assert.equal(res.orders.length, 3);
  assert.equal(res.orders[0].phone, "01020556927");
  assert.equal(res.orders[0].product, "5키로");
  assert.equal(res.orders[0].boxes, "1박스");
});
test("행 위치가 아닌 원본 번호로 필터, 포함 경계와 번호 공백 처리", () => {
  const orders = app.processWorkbook(workbook()).orders;
  assert.deepEqual(
    app.selectedOrders(orders, 10).map((o) => o.seq),
    [10, 15],
  );
  assert.equal(app.selectedOrders(orders, 1).length, 3);
  for (const start of [16, 0, -1, 1.5, null, NaN])
    assert.equal(app.selectedOrders(orders, start).length, 0);
});
test("실제 XLSX 왕복: 7열만 출력, 선택된 주문은 순번 1부터, 연락처는 문자열", () => {
  const orders = app.processWorkbook(workbook()).orders;
  const wb = app.createExportWorkbook(orders, 10);
  const roundtrip = XLSX.read(
    XLSX.write(wb, { type: "buffer", bookType: "xlsx" }),
    { type: "buffer" },
  );
  assert.deepEqual(roundtrip.SheetNames, ["Sheet1"]);
  assert.equal(roundtrip.Sheets.Sheet1["!autofilter"], undefined);
  const output = XLSX.utils.sheet_to_json(roundtrip.Sheets.Sheet1, {
    header: 1,
  });
  assert.deepEqual(output, [
    app.exportHeaders,
    [
      1,
      "10키로",
      "2박스",
      "테스트 나",
      "01000000001",
      "테스트 주소 나",
      "테스트 / 100000",
    ],
    [
      2,
      "5키로",
      "1박스",
      "테스트 다",
      "01000000002",
      "테스트 주소 다",
      "테스트 / 50000",
    ],
  ]);
  assert.equal(roundtrip.Sheets.Sheet1.E2.t, "s");
  assert.throws(() => app.createExportWorkbook(orders, 100));
});
test("행 선택·키보드 선택, 원본 번호 표기, 엑셀 순번 일치, 리셋·동일 파일 재선택", async () => {
  const dom = new JSDOM(
    fs.readFileSync(path.join(__dirname, "../index.html"), "utf8"),
    { runScripts: "outside-only", url: "http://localhost/" },
  );
  const w = dom.window;
  let saved;
  w.eval(
    fs.readFileSync(require.resolve("xlsx/dist/xlsx.full.min.js"), "utf8"),
  );
  w.XLSX.writeFile = (wb, name) => {
    saved = { wb, name };
  };
  w.eval(fs.readFileSync(path.join(__dirname, "../app.js"), "utf8"));
  const input = w.document.getElementById("file-input");
  async function upload(bytes, name = "sample.xlsx") {
    Object.defineProperty(input, "files", {
      configurable: true,
      value: [new w.File([bytes], name)],
    });
    input.dispatchEvent(new w.Event("change"));
    await new Promise((resolve, reject) => {
      const timeout = setTimeout(() => {
        observer.disconnect();
        reject(new Error("파일 처리 시간 초과"));
      }, 3000);
      const observer = new w.MutationObserver(() => {
        if (!w.document.getElementById("btn-select").disabled) {
          clearTimeout(timeout);
          observer.disconnect();
          resolve();
        }
      });
      observer.observe(w.document.getElementById("btn-select"), {
        attributes: true,
        attributeFilter: ["disabled"],
      });
    });
  }
  await upload(XLSX.write(workbook(), { type: "buffer", bookType: "xlsx" }));
  const doc = w.document;
  assert.equal(
    doc.getElementById("preview-section").hidden,
    false,
    doc.getElementById("file-message").textContent,
  );
  assert.equal(doc.getElementById("revision"), null);
  assert.equal(doc.body.classList.contains("has-preview"), true);
  doc.querySelector('tr[data-original-number="10"] td:nth-child(4)').click();
  assert.equal(doc.querySelectorAll("tr.excluded").length, 1);
  assert.equal(
    doc.getElementById("selection-summary").textContent,
    "원본 10번부터 · 2건 저장",
  );
  assert.equal(
    doc.querySelector('tr[data-original-number="10"] td').textContent,
    "1(원본10)",
  );
  // 제외된 행을 클릭해 시작 지점을 앞으로 되돌립니다.
  doc.querySelector("tr.excluded td:nth-child(4)").click();
  assert.equal(doc.querySelectorAll("tr.excluded").length, 0);
  const lastRow = doc.querySelector('tr[data-original-number="15"]');
  lastRow.focus();
  lastRow.dispatchEvent(new w.KeyboardEvent("keydown", { key: "Enter" }));
  assert.equal(doc.activeElement.dataset.originalNumber, "15");
  assert.equal(doc.querySelectorAll("tr.excluded").length, 2);
  doc.querySelector('tr[data-original-number="10"]').click();
  doc.getElementById("btn-download").click();
  assert.match(saved.name, /_수정본_10번부터\.xlsx$/);
  const visible = Array.from(
    doc.querySelectorAll("tbody tr:not(.excluded)"),
  ).map((tr) =>
    Array.from(tr.cells).map((td, i) =>
      i === 0 ? Number(td.firstChild.textContent) : td.textContent,
    ),
  );
  const exported = XLSX.utils
    .sheet_to_json(saved.wb.Sheets.Sheet1, { header: 1 })
    .slice(1);
  assert.deepEqual(visible, JSON.parse(JSON.stringify(exported)));
  await upload(Buffer.from("invalid"), "broken.xlsx");
  assert.equal(doc.getElementById("file-message").hidden, false);
  doc.getElementById("btn-reset").click();
  assert.equal(doc.getElementById("upload-section").hidden, false);
  assert.equal(doc.body.classList.contains("has-preview"), false);
  await upload(XLSX.write(workbook(), { type: "buffer", bookType: "xlsx" }));
  assert.match(
    doc.getElementById("selection-summary").textContent,
    /^3건 저장/,
  );
  dom.window.close();
});
test("주문 시트 없음, 빈 주문, 잘못된 주문번호 오류", () => {
  assert.throws(
    () => app.processWorkbook({ SheetNames: [], Sheets: {} }),
    /주문 시트/,
  );
  for (const data of [[headers], [headers, ["ABC", "사과", "", "테스트"]]]) {
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(data), "답변");
    assert.throws(() => app.processWorkbook(wb));
  }
});

test("주문번호 열이 없을 때 전화번호와 혼동하지 않고 행 순서 부여", () => {
  const res = app.buildOrdersFromRows([
    ["상품", "이름", "전화번호"],
    ["5키로", "테스트", "01000000000"],
    [],
    ["10키로", "테스트", "01000000001"],
  ]);
  assert.deepEqual(
    res.orders.map((o) => o.seq),
    [1, 2],
  );
});
