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
    fixes: { phone: 0, product: 0, boxes: 0, recipient: 0, address: 0, deposit: 0, confirm: 0, privacy: 0 },
    revision: 1,
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
    var norms = headerRow.map(function (h) { return normKey(h); });
    for (var k = 0; k < keys.length; k++) {
      var nk = normKey(keys[k]);
      var idx = norms.indexOf(nk);
      if (idx === -1) idx = norms.findIndex(function (h) { return h.indexOf(nk) !== -1; });
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
    var header = (rows[0] || []).map(function (c) { return str(c); });
    var col = {
      seq: findCol(header, ["주문서번호", "주문번호", "순번", "번호", "참여자", "응답자", "participant"]),
      product: findCol(header, ["사과상품선택", "사과상품", "상품", "product"]),
      boxes: findCol(header, ["사과수량선택", "수량", "박스", "수량선택", "quantity"]),
      recipient: findCol(header, ["주문자이름", "받는사람", "고객이름", "이름", "recipient"]),
      phone: findCol(header, ["연락처", "전화번호", "휴대폰", "전화", "phone"]),
      address: findCol(header, ["배송지주소", "배송지", "주소", "address"]),
      deposit: findCol(header, ["입금자", "총입금액", "입금", "deposit", "결제금액"]),
      confirm: findCol(header, ["주문서내용확인", "주문서확인", "confirm", "주문서내용"]),
      privacy: findCol(header, ["개인정보", "개인정부", "동의", "privacy", "개인정보수집"]),
    };
    var orders = [];
    var fixes = { phone: 0, product: 0, boxes: 0 };
    var rawCount = 0;

    for (var i = 1; i < (rows ? rows.length : 0); i++) {
      var row = rows[i] || [];
      var get = function (idx) { return idx >= 0 && idx < row.length ? row[idx] : null; };
      var rawProduct = get(col.product);
      var rawBoxes = get(col.boxes);
      var rawPhone = get(col.phone);
      var product = cleanProduct(rawProduct);
      var boxes = cleanBoxes(rawBoxes);
      var phone = normalizePhone(rawPhone);

      if (str(get(col.recipient)) !== "" || product !== "" || boxes !== "" || str(get(col.address)) !== "") {
        rawCount++;
        if (str(rawPhone).replace(/[^0-9]/g, "") !== "" && str(rawPhone).replace(/[^0-9]/g, "") !== phone) fixes.phone++;
        if (String(rawProduct) !== product) fixes.product++;
        if (String(rawBoxes) !== boxes) fixes.boxes++;
      }

      var seqStr = str(get(col.seq)) !== "" ? str(get(col.seq)) : String(i);
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
    return { header: header, col: col, orders: orders, fixes: fixes, rawCount: rawCount };
  }

  function processWorkbook(wb) {
    var sheetName = "";
    var rows = null;
    var bestScore = -1;
    var sheets = (wb && wb.SheetNames) ? wb.SheetNames : [];
    for (var i = 0; i < sheets.length; i++) {
      var ws = wb.Sheets[sheets[i]];
      if (!ws) continue;
      var data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
      var header = (data[0] || []).map(function (c) { return str(c); });
      var flat = header.join(" ");
      var hasApple = flat.indexOf("사과") !== -1;
      var hasOrder = /(수량|상품|주문|박스|배송)/.test(flat);
      var score = (hasApple ? 2 : 0) + (hasOrder ? 1 : 0) + Math.min((data ? data.length : 0), 50);
      if (score > bestScore) {
        bestScore = score;
        sheetName = sheets[i];
        rows = data;
      }
    }
    var res = buildOrdersFromRows(rows);
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
    return String(s === null || s === undefined ? "" : s)
      .replace(/[&<>"']/g, function (c) {
        if (c === "&") return "&amp;";
        if (c === "<") return "&lt;";
        if (c === ">") return "&gt;";
        if (c === '"') return "&quot;";
        return "&#39;";
      });
  }

  function renderPreview() {
    var revInput = document.getElementById("revision");
    var rev = parseInt(revInput ? revInput.value : "1", 10);
    if (isNaN(rev) || rev < 1) rev = 1;
    state.revision = rev;
    var badge = document.getElementById("revision-badge");
    if (badge) badge.textContent = "v" + rev;

    var fileEl = document.getElementById("meta-file");
    if (fileEl) fileEl.textContent = state.fileName;
    var sheetEl = document.getElementById("meta-sheet");
    if (sheetEl) sheetEl.textContent = state.sheetName;
    var countEl = document.getElementById("meta-count");
    if (countEl) countEl.textContent = state.orders.length;

    var fixesEl = document.getElementById("fixes");
    fixesEl.innerHTML = "";
    var fixParts = [];
    if (state.fixes.phone) fixParts.push("전화번호 정규화: " + state.fixes.phone + "건");
    if (state.fixes.product) fixParts.push("상품명 정리: " + state.fixes.product + "건");
    if (state.fixes.boxes) fixParts.push("수량/박스 정리: " + state.fixes.boxes + "건");
    if (fixParts.length) {
      fixesEl.innerHTML =
        '<div class="fixes-title">수정 내역</div><ul><li>' + fixParts.join("</li><li>") + "</li></ul>";
    } else {
      fixesEl.innerHTML = '<div class="fixes-title">수정 내역</div><div>별도의 데이터 수정이 필요하지 않습니다.</div>';
    }

    var tbody = document.getElementById("preview-tbody");
    if (tbody) {
      tbody.innerHTML = "";
      for (var i = 0; i < state.orders.length; i++) {
        var o = state.orders[i];
        var tr = document.createElement("tr");
        var cells = [
          String(i + 1),
          o.seq,
          o.product,
          o.boxes,
          o.recipient,
          o.phone,
          o.address,
          o.deposit,
        ];
        for (var c = 0; c < cells.length; c++) {
          var td = document.createElement("td");
          td.textContent = cells[c] === null || cells[c] === undefined ? "" : String(cells[c]);
          tr.appendChild(td);
        }
        tbody.appendChild(tr);
      }
    }

    var preview = document.getElementById("preview-section");
    if (preview) {
      preview.hidden = false;
      preview.scrollIntoView({ behavior: "smooth", block: "start" });
    }
    var upload = document.getElementById("upload-section");
    if (upload) upload.hidden = true;
  }

  /* ---------- 파일 처리 ---------- */
  function handleFile(file) {
    state.file = file;
    state.fileName = file.name;
    var reader = new FileReader();
    reader.onload = function (e) {
      var data = e.target.result;
      var wb = XLSX.read(data, { type: "array" });
      var res = processWorkbook(wb);
      state.sheetName = res.sheetName;
      state.orders = res.orders;
      state.fixes = res.fixes;
      state.rawCount = res.rawCount;
      renderPreview();
    };
    reader.readAsArrayBuffer(file);
  }

  /* ---------- 수정본 다운로 ---------- */
  function downloadModified() {
    if (!state.orders || !state.orders.length) return;
    var rev = state.revision;
    var header = [
      null,
      "사과상품선택",
      "사과 수량선택",
      "주문자이름",
      "연락처",
      "배송지주소",
      "입금자/총입금액",
      "주문서내용확인",
      null,
      "개인정보수집박의",
    ];
    var data = [];
    for (var i = 0; i < state.orders.length; i++) {
      var o = state.orders[i];
      data.push([
        o.seq,
        o.product,
        o.boxes,
        o.recipient,
        o.phone,
        o.address,
        o.deposit,
        o.confirm,
        null,
        o.privacy,
      ]);
    }
    var ws = XLSX.utils.aoa_to_sheet([header].concat(data));
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
    var base = state.fileName.replace(/\.[^.]*$/, "");
    var outName = base + "_v" + rev + ".xlsx";
    XLSX.writeFile(wb, outName);
  }

  function resetAll() {
    state.file = null;
    state.fileName = "";
    state.sheetName = "";
    state.orders = [];
    state.fixes = { phone: 0, product: 0, boxes: 0 };
    var revInput = document.getElementById("revision");
    if (revInput) revInput.value = "1";
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

    function isValidFile(file) {
      if (!file) return false;
      return /\.xlsx?$/i.test(file.name || "") || /\.xls$/i.test(file.name || "");
    }

    if (btnSelect && fileInput) {
      btnSelect.addEventListener("click", function () { fileInput.click(); });
      fileInput.addEventListener("change", function (e) {
        var f = e.target.files && e.target.files[0];
        if (f && isValidFile(f)) handleFile(f);
      });
    }
    if (dropzone) {
      dropzone.addEventListener("dragover", function (e) { e.preventDefault(); dropzone.classList.add("dragover"); });
      dropzone.addEventListener("dragleave", function () { dropzone.classList.remove("dragover"); });
      dropzone.addEventListener("drop", function (e) {
        e.preventDefault();
        dropzone.classList.remove("dragover");
        var f = e.dataTransfer ? e.dataTransfer.files && e.dataTransfer.files[0] : null;
        if (f && isValidFile(f)) handleFile(f);
      });
    }

    if (btnDownload) btnDownload.addEventListener("click", downloadModified);
    if (btnReset) btnReset.addEventListener("click", resetAll);
  }

  /* ---------- 테스트/Node 노출 ---------- */
  if (typeof module !== "undefined" && module.exports) {
    module.exports = {
      state: state,
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
