/* 신선한 산지직송 사과 주문서 수정기
 * - 네이버폼 응답 엑셀(설문 답변)을 "신선한 산지직송 사과 주문서" 형식으로 변환
 * - 브라우저: window.XLSX (SheetJS CDN) / Node: require("xlsx")
 */
(function () {
  var XLSX = null;
  try {
    if (typeof window !== "undefined" && window.XLSX) {
      XLSX = window.XLSX;
    } else if (typeof require === "function") {
      XLSX = require("xlsx");
    }
  } catch (e) {
    XLSX = null;
  }

  var state = {
    file: null,
    fileName: "",
    sheetName: "",
    orders: [],
    fixes: {
      phone: 0,
      product: 0,
      boxes: 0,
      recipient: 0,
      address: 0,
      deposit: 0,
      confirm: 0,
      privacy: 0,
    },
    revision: 1,
    loading: false,
  };

  /* ---------- 기본 유틸 ---------- */
  function str(v) {
    if (v === null || v === undefined) return "";
    if (typeof v === "number") return String(v);
    return String(v).trim();
  }

  function normKey(s) {
    var t = String(s === null || s === undefined ? "" : s).toLowerCase();
    var i = t.indexOf("(");
    if (i >= 0) t = t.slice(0, i);
    t = t.replace(/[*.]/g, "").replace(/\s+/g, "").trim();
    return t;
  }

  function findCol(headerRow, keys) {
    var norms = headerRow.map(function (h) {
      return normKey(h);
    });
    for (var k = 0; k < keys.length; k++) {
      var nk = normKey(keys[k]);
      var idx = norms.indexOf(nk);
      if (idx === -1 && nk !== "번호")
        idx = norms.findIndex(function (h) {
          return h.indexOf(nk) !== -1;
        });
      if (idx !== -1) return idx;
    }
    return -1;
  }

  function cleanProduct(v) {
    var s = str(v);
    var i = s.indexOf("(");
    if (i >= 0) s = s.slice(0, i);
    return s.replace(/\s+/g, "").trim();
  }

  function cleanBoxes(v) {
    var s = str(v);
    s = s.replace(/^\s*\d+\.\s*/, "").trim();
    return s;
  }

  function normalizePhone(v) {
    var d = str(v).replace(/[^0-9]/g, "");
    if (d.length === 10 && d.charAt(0) === "1") d = "0" + d;
    return d;
  }

  /* ---------- 주문 데이터 추출 / 변환 ---------- */
  function buildOrdersFromRows(rows) {
    var header = ((rows || [])[0] || []).map(function (c) {
      return str(c);
    });
    var col = {
      seq: findCol(header, [
        "주문서번호",
        "주문번호",
        "순번",
        "번호",
        "참여자",
        "응답자",
        "participant",
      ]),
      product: findCol(header, ["사과상품선택", "사과상품", "상품", "product"]),
      boxes: findCol(header, [
        "사과수량선택",
        "수량",
        "박스",
        "수량선택",
        "quantity",
      ]),
      recipient: findCol(header, [
        "주문자이름",
        "받는사람",
        "고객이름",
        "이름",
        "recipient",
      ]),
      phone: findCol(header, ["연락처", "전화번호", "휴대폰", "전화", "phone"]),
      address: findCol(header, ["배송지주소", "배송지", "주소", "address"]),
      deposit: findCol(header, [
        "입금자",
        "총입금액",
        "입금",
        "deposit",
        "결제금액",
      ]),
      confirm: findCol(header, [
        "주문서내용확인",
        "주문서확인",
        "confirm",
        "주문서내용",
      ]),
      privacy: findCol(header, [
        "개인정보",
        "개인정부",
        "동의",
        "privacy",
        "개인정보수집",
      ]),
    };
    var orders = [];
    var fixes = { phone: 0, product: 0, boxes: 0 };
    var rawCount = 0;

    for (var i = 1; i < (rows ? rows.length : 0); i++) {
      var row = rows[i] || [];
      var get = function (idx) {
        return idx >= 0 && idx < row.length ? row[idx] : null;
      };
      var rawProduct = get(col.product);
      var rawBoxes = get(col.boxes);
      var rawPhone = get(col.phone);
      var product = cleanProduct(rawProduct);
      var boxes = cleanBoxes(rawBoxes);
      var phone = normalizePhone(rawPhone);

      if (
        str(get(col.recipient)) === "" &&
        product === "" &&
        boxes === "" &&
        str(get(col.address)) === ""
      )
        continue;
      rawCount++;
      if (
        str(rawPhone).replace(/[^0-9]/g, "") !== "" &&
        str(rawPhone).replace(/[^0-9]/g, "") !== phone
      )
        fixes.phone++;
      if (str(rawProduct) !== product) fixes.product++;
      if (str(rawBoxes) !== boxes) fixes.boxes++;

      var seqStr =
        str(get(col.seq)) !== ""
          ? str(get(col.seq))
          : String(orders.length + 1);
      var seqVal = /^\d+$/.test(seqStr) ? parseInt(seqStr, 10) : seqStr;
      orders.push({
        seq: seqVal,
        product: product,
        boxes: boxes,
        recipient: str(get(col.recipient)),
        phone: phone,
        address: str(get(col.address)),
        deposit: str(get(col.deposit)),
        confirm: str(get(col.confirm)),
        privacy: str(get(col.privacy)),
      });
    }
    return {
      header: header,
      col: col,
      orders: orders,
      fixes: fixes,
      rawCount: rawCount,
    };
  }

  function processWorkbook(wb) {
    var sheetName = "";
    var rows = null;
    var bestScore = -1;
    var sheets = wb && wb.SheetNames ? wb.SheetNames : [];
    for (var i = 0; i < sheets.length; i++) {
      var ws = wb.Sheets[sheets[i]];
      if (!ws) continue;
      var data = XLSX.utils.sheet_to_json(ws, {
        header: 1,
        defval: null,
        raw: true,
      });
      var header = (data[0] || []).map(function (c) {
        return str(c);
      });
      var hasProduct =
        findCol(header, ["사과상품선택", "사과상품", "상품", "product"]) >= 0;
      var hasRecipient =
        findCol(header, [
          "주문자이름",
          "받는사람",
          "고객이름",
          "이름",
          "recipient",
        ]) >= 0;
      if (!hasProduct || !hasRecipient) continue;
      var flat = header.join(" ");
      var hasApple = flat.indexOf("사과") !== -1;
      var hasOrder = /(수량|상품|주문|박스|배송)/.test(flat);
      var score =
        (hasApple ? 2 : 0) +
        (hasOrder ? 1 : 0) +
        Math.min(data ? data.length : 0, 50);
      if (score > bestScore) {
        bestScore = score;
        sheetName = sheets[i];
        rows = data;
      }
    }
    if (!rows)
      throw new Error(
        "상품과 주문자이름 열이 있는 주문 시트를 찾지 못했습니다. 네이버폼 응답 엑셀인지 확인해 주세요.",
      );
    var res = buildOrdersFromRows(rows);
    if (!res.orders.length)
      throw new Error(
        "주문 데이터가 없습니다. 주문이 포함된 파일을 선택해 주세요.",
      );
    if (
      res.orders.some(function (o) {
        return !Number.isSafeInteger(Number(o.seq)) || Number(o.seq) < 1;
      })
    ) {
      throw new Error(
        "주문서번호는 1 이상의 정수여야 합니다. 원본 파일의 번호를 확인해 주세요.",
      );
    }
    return {
      sheetName: sheetName,
      orders: res.orders,
      fixes: res.fixes,
      rawCount: res.rawCount,
      header: res.header,
      col: res.col,
    };
  }

  /* ---------- 미리보기 렌더링 ---------- */
  function esc(s) {
    return String(s === null || s === undefined ? "" : s).replace(
      /[&<>"']/g,
      function (c) {
        if (c === "&") return "&amp;";
        if (c === "<") return "&lt;";
        if (c === ">") return "&gt;";
        if (c === '"') return "&quot;";
        return "&#39;";
      },
    );
  }

  var exportHeaders = [
    "순번",
    "사과상품선택",
    "사과 수량선택",
    "주문자이름",
    "연락처",
    "배송지주소",
    "입금자/총입금액",
  ];

  function isSelected(order, start) {
    return (
      Number.isSafeInteger(start) && start >= 1 && Number(order.seq) >= start
    );
  }

  function selectedOrders(orders, start) {
    return orders.filter(function (o) {
      return isSelected(o, start);
    });
  }

  function exportRow(order, index) {
    return [
      index + 1,
      order.product,
      order.boxes,
      order.recipient,
      order.phone,
      order.address,
      order.deposit,
    ];
  }

  function createExportWorkbook(orders, start) {
    var selected = selectedOrders(orders, start);
    if (!selected.length) throw new Error("다운로드할 주문이 없습니다.");
    var ws = XLSX.utils.aoa_to_sheet(
      [exportHeaders].concat(selected.map(exportRow)),
    );
    ws["!cols"] = [
      { wch: 8 },
      { wch: 18 },
      { wch: 16 },
      { wch: 16 },
      { wch: 18 },
      { wch: 60 },
      { wch: 30 },
    ];
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
    return wb;
  }

  function showMessage(message) {
    var el = document.getElementById("file-message");
    el.textContent = message;
    el.hidden = !message;
  }

  function selectStartRow(seq) {
    if (state.revision === Number(seq)) return;
    var restoreFocus = document
      .getElementById("preview-tbody")
      .contains(document.activeElement);
    state.revision = Number(seq);
    renderPreview();
    if (restoreFocus) {
      var row = document.querySelector(
        'tr[data-original-number="' + state.revision + '"]',
      );
      if (row) row.focus({ preventScroll: true });
    }
  }

  function renderPreview() {
    var selected = selectedOrders(state.orders, state.revision);
    document.getElementById("meta-file").textContent = state.fileName;
    document.getElementById("selection-summary").textContent =
      state.revision === 1
        ? selected.length + "건 저장 · 시작할 행을 클릭하세요"
        : "원본 " + state.revision + "번부터 · " + selected.length + "건 저장";
    document.getElementById("btn-download").disabled = !selected.length;
    var tbody = document.getElementById("preview-tbody");
    tbody.innerHTML = "";
    var sequence = 0;
    state.orders.forEach(function (o) {
      var active = isSelected(o, state.revision);
      var tr = document.createElement("tr");
      if (!active) tr.classList.add("excluded");
      tr.tabIndex = 0;
      tr.setAttribute("aria-label", "원본 " + o.seq + "번부터 저장");
      if (Number(o.seq) === state.revision)
        tr.setAttribute("aria-current", "true");
      tr.addEventListener("keydown", function (event) {
        if (event.key === "Enter" || event.key === " ") {
          event.preventDefault();
          selectStartRow(o.seq);
        }
      });
      tr.dataset.originalNumber = String(o.seq);
      tr.addEventListener("click", function () {
        selectStartRow(o.seq);
      });
      var cells = exportRow(o, active ? sequence++ : -1);
      if (!active) cells[0] = "—";
      cells.forEach(function (cell, c) {
        var td = document.createElement("td");
        td.textContent = cell;
        if (c === 0) {
          var source = document.createElement("span");
          source.className = "source-number";
          source.textContent = "(원본" + o.seq + ")";
          td.appendChild(source);
        }
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });
    document.getElementById("preview-section").hidden = false;
    document.getElementById("upload-section").hidden = true;
    document.body.classList.add("has-preview");
  }

  function handleFile(file) {
    if (state.loading) return;
    if (!file || !/\.(xlsx|xls)$/i.test(file.name)) {
      showMessage(".xlsx 또는 .xls 형식의 엑셀 파일을 선택해 주세요.");
      return;
    }
    if (!XLSX) {
      showMessage(
        "엑셀 처리 도구를 불러오지 못했습니다. 인터넷 연결을 확인한 뒤 페이지를 새로고침해 주세요.",
      );
      return;
    }
    state.loading = true;
    document.getElementById("btn-select").disabled = true;
    showMessage("주문 파일을 읽고 있습니다…");
    var reader = new FileReader();
    function finish() {
      state.loading = false;
      document.getElementById("btn-select").disabled = false;
      document.getElementById("file-input").value = "";
    }
    reader.onload = function (e) {
      try {
        var wb = XLSX.read(e.target.result, { type: "array" });
        var res = processWorkbook(wb);
        state.file = file;
        state.fileName = file.name;
        state.sheetName = res.sheetName;
        state.orders = res.orders;
        state.fixes = res.fixes;
        state.rawCount = res.rawCount;
        state.revision = 1;
        renderPreview();
        showMessage("");
      } catch (error) {
        showMessage("파일을 처리하지 못했습니다. " + error.message);
      } finally {
        finish();
      }
    };
    reader.onerror = function () {
      showMessage("파일을 읽지 못했습니다. 다시 선택해 주세요.");
      finish();
    };
    reader.readAsArrayBuffer(file);
  }

  function downloadModified() {
    renderPreview();
    if (!selectedOrders(state.orders, state.revision).length) return;
    try {
      var wb = createExportWorkbook(state.orders, state.revision);
      var base = state.fileName.replace(/\.[^.]*$/, "");
      XLSX.writeFile(wb, base + "_수정본_" + state.revision + "번부터.xlsx");
      showMessage("");
    } catch (error) {
      showMessage("다운로드하지 못했습니다. " + error.message);
    }
  }

  function resetAll() {
    showMessage("");
    state.revision = 1;
    document.getElementById("file-input").value = "";
    state.file = null;
    state.fileName = "";
    state.sheetName = "";
    state.orders = [];
    state.fixes = { phone: 0, product: 0, boxes: 0 };
    document.body.classList.remove("has-preview");
    var preview = document.getElementById("preview-section");
    if (preview) preview.hidden = true;
    var upload = document.getElementById("upload-section");
    if (upload) upload.hidden = false;
  }

  /* ---------- 브라우저 전용 DOM 와이어링 ---------- */
  if (typeof document !== "undefined") {
    var dropzone = document.getElementById("dropzone");
    var fileInput = document.getElementById("file-input");
    var btnSelect = document.getElementById("btn-select");
    var btnDownload = document.getElementById("btn-download");
    var btnReset = document.getElementById("btn-reset");

    if (btnSelect && fileInput) {
      btnSelect.addEventListener("click", function () {
        fileInput.click();
      });
      fileInput.addEventListener("change", function (e) {
        var f = e.target.files && e.target.files[0];
        if (f) handleFile(f);
      });
    }
    if (dropzone) {
      dropzone.addEventListener("dragover", function (e) {
        e.preventDefault();
        dropzone.classList.add("dragover");
      });
      dropzone.addEventListener("dragleave", function () {
        dropzone.classList.remove("dragover");
      });
      dropzone.addEventListener("drop", function (e) {
        e.preventDefault();
        dropzone.classList.remove("dragover");
        var f = e.dataTransfer
          ? e.dataTransfer.files && e.dataTransfer.files[0]
          : null;
        if (f) handleFile(f);
      });
    }

    if (btnDownload) btnDownload.addEventListener("click", downloadModified);
    if (btnReset) btnReset.addEventListener("click", resetAll);
  }

  /* ---------- 테스트/Node 노출 ---------- */
  if (typeof module !== "undefined" && module.exports) {
    module.exports = {
      state: state,
      selectedOrders: selectedOrders,
      createExportWorkbook: createExportWorkbook,
      exportHeaders: exportHeaders,
      str: str,
      normKey: normKey,
      findCol: findCol,
      cleanProduct: cleanProduct,
      cleanBoxes: cleanBoxes,
      normalizePhone: normalizePhone,
      buildOrdersFromRows: buildOrdersFromRows,
      processWorkbook: processWorkbook,
    };
  }
})();
