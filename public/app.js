const statusEl = document.querySelector("#status");
const logEl = document.querySelector("#log");
const bodyEl = document.querySelector("#ordersBody");
const summaryEl = document.querySelector("#summary");
const downloadLink = document.querySelector("#downloadLink");
const placeOrdersButton = document.querySelector("#placeOrders");
const orderError = document.querySelector("#orderError");
const newOrdersTab = document.querySelector("#newOrdersTab");
const confirmedOrdersTab = document.querySelector("#confirmedOrdersTab");
const newOrdersCount = document.querySelector("#newOrdersCount");
const confirmedOrdersCount = document.querySelector("#confirmedOrdersCount");
const prevOrdersPageButton = document.querySelector("#prevOrdersPage");
const nextOrdersPageButton = document.querySelector("#nextOrdersPage");
const ordersPageInfo = document.querySelector("#ordersPageInfo");
const trackingBody = document.querySelector("#trackingBody");
const shipmentsBody = document.querySelector("#shipmentsBody");
const dispatchSummary = document.querySelector("#dispatchSummary");
const shipmentSummary = document.querySelector("#shipmentSummary");

const ORDERS_PAGE_SIZE = 10;
let currentOrders = [];
let confirmedOrders = [];
let activeOrderView = "new";
let currentOrderPage = 1;
let currentDownload = null;
let currentTrackingRecords = [];

function today() {
  return new Date().toISOString().slice(0, 10);
}

document.querySelector("#dateFrom").value = today();
document.querySelector("#dateTo").value = today();
document.querySelector("#shipDateFrom").value = today();
document.querySelector("#shipDateTo").value = today();

function setStatus(text) {
  statusEl.textContent = text;
}

function log(message, detail) {
  const time = new Date().toLocaleTimeString("ko-KR", { hour12: false });
  const extra = detail ? `\n${typeof detail === "string" ? detail : JSON.stringify(detail, null, 2)}` : "";
  logEl.textContent = `[${time}] ${message}${extra}\n\n${logEl.textContent}`;
}

function clearOrderError() {
  orderError.textContent = "";
  orderError.classList.add("hidden");
}

function showOrderError(message) {
  orderError.textContent = message;
  orderError.classList.remove("hidden");
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function renderRows(rows) {
  currentOrders = rows || [];
  currentOrderPage = 1;
  renderOrdersPage();
}

function visibleOrders() {
  return activeOrderView === "confirmed" ? confirmedOrders : currentOrders;
}

function setOrderView(view) {
  activeOrderView = view;
  currentOrderPage = 1;
  newOrdersTab.classList.toggle("active", view === "new");
  confirmedOrdersTab.classList.toggle("active", view === "confirmed");
  renderOrdersPage();
}

function renderOrdersPage() {
  const orders = visibleOrders();
  newOrdersCount.textContent = currentOrders.length.toLocaleString("ko-KR");
  confirmedOrdersCount.textContent = confirmedOrders.length.toLocaleString("ko-KR");
  if (orders.length === 0) {
    const emptyText = activeOrderView === "confirmed" ? "발주확인 완료된 주문이 없습니다." : "조회된 신규 주문이 없습니다.";
    bodyEl.innerHTML = `<tr><td colspan="6" class="empty">${emptyText}</td></tr>`;
    summaryEl.textContent = `${activeOrderView === "confirmed" ? "발주확인 완료" : "신규 주문"} 0건`;
    ordersPageInfo.textContent = "1 / 1";
    prevOrdersPageButton.disabled = true;
    nextOrdersPageButton.disabled = true;
    return;
  }

  const totalPages = Math.max(1, Math.ceil(orders.length / ORDERS_PAGE_SIZE));
  currentOrderPage = Math.min(Math.max(1, currentOrderPage), totalPages);
  const start = (currentOrderPage - 1) * ORDERS_PAGE_SIZE;
  const pageRows = orders.slice(start, start + ORDERS_PAGE_SIZE);

  bodyEl.innerHTML = pageRows.map((row) => `
    <tr>
      <td>${escapeHtml(row.productOrderNo || row["상품주문번호"] || "")}</td>
      <td>${escapeHtml(row.receiver || row["수취인명"] || "")}</td>
      <td>${escapeHtml(row.phone || row["수취인연락처1"] || "")}</td>
      <td>${escapeHtml(row.address || row["통합배송지"] || "")}</td>
      <td>${escapeHtml(row.product || row["상품명"] || "")}<br><small>${escapeHtml(row.option || row["옵션정보"] || "")}</small></td>
      <td>${escapeHtml(String(row.quantity || row["수량"] || ""))}</td>
    </tr>
  `).join("");

  const label = activeOrderView === "confirmed" ? "발주확인 완료" : "신규 주문";
  summaryEl.textContent = `${label} ${orders.length.toLocaleString("ko-KR")}건 중 ${start + 1}-${start + pageRows.length} 표시`;
  ordersPageInfo.textContent = `${currentOrderPage} / ${totalPages}`;
  prevOrdersPageButton.disabled = currentOrderPage <= 1;
  nextOrdersPageButton.disabled = currentOrderPage >= totalPages;
}

function renderTracking(records) {
  currentTrackingRecords = records || [];
  if (currentTrackingRecords.length === 0) {
    trackingBody.innerHTML = `<tr><td colspan="4" class="empty">읽은 송장번호가 없습니다.</td></tr>`;
    dispatchSummary.textContent = "0건";
    return;
  }
  trackingBody.innerHTML = currentTrackingRecords.map((row) => `
    <tr>
      <td>${escapeHtml(row.productOrderNo || "")}</td>
      <td>${escapeHtml(row.trackingNo || "")}</td>
      <td>${escapeHtml(row.carrier || "CJ대한통운")}</td>
      <td>${escapeHtml(row.status || "READY")}</td>
    </tr>
  `).join("");
  dispatchSummary.textContent = `${currentTrackingRecords.length.toLocaleString("ko-KR")}건 전송대기`;
}

function renderShipments(rows, summary) {
  if (!rows || rows.length === 0) {
    shipmentsBody.innerHTML = `<tr><td colspan="6" class="empty">조건에 맞는 출고 내역이 없습니다.</td></tr>`;
  } else {
    shipmentsBody.innerHTML = rows.slice().reverse().map((row) => `
      <tr>
        <td>${escapeHtml(String(row.dispatchDate || row.sentAt || row.createdAt || "").slice(0, 10))}</td>
        <td>${escapeHtml(row.productOrderNo || "")}</td>
        <td>${escapeHtml(row.trackingNo || "")}</td>
        <td>${escapeHtml(row.carrier || "CJ대한통운")}</td>
        <td>${escapeHtml(row.status || "")}</td>
        <td>${escapeHtml(row.error || "")}</td>
      </tr>
    `).join("");
  }
  const count = summary || { total: rows?.length || 0, ready: 0, sent: 0, failed: 0 };
  shipmentSummary.textContent = `전체 ${count.total}건 / 대기 ${count.ready} / 완료 ${count.sent} / 실패 ${count.failed}`;
}

async function postJson(url, payload) {
  const response = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  const text = await response.text();
  let data;
  try {
    data = text ? JSON.parse(text) : {};
  } catch {
    throw new Error(`서버가 JSON이 아닌 응답을 보냈습니다. 상태 ${response.status}: ${text.slice(0, 500)}`);
  }
  if (!response.ok || !data.ok) {
    const detail = data.results ? `\n${JSON.stringify(data.results, null, 2)}` : "";
    throw new Error(`${data.error || "요청 처리에 실패했습니다."}${detail}`);
  }
  return data;
}

async function postForm(url, form) {
  const response = await fetch(url, { method: "POST", body: new FormData(form) });
  const text = await response.text();
  let data;
  try {
    data = text ? JSON.parse(text) : {};
  } catch {
    throw new Error(`서버가 JSON이 아닌 응답을 보냈습니다. 상태 ${response.status}: ${text.slice(0, 500)}`);
  }
  if (!response.ok || !data.ok) {
    throw new Error(data.error || "요청 처리에 실패했습니다.");
  }
  return data;
}

function credentialPayload() {
  return {};
}

function resetOrderWorkflow() {
  clearOrderError();
  confirmedOrders = [];
  activeOrderView = "new";
  currentDownload = null;
  downloadLink.classList.add("hidden");
  downloadLink.removeAttribute("href");
  downloadLink.removeAttribute("download");
  placeOrdersButton.disabled = true;
}

document.querySelector("#testToken").addEventListener("click", async () => {
  clearOrderError();
  setStatus("연결 중");
  log("스마트스토어 API 연결 상태를 확인합니다.");
  try {
    const data = await postJson("/api/token/test", credentialPayload());
    setStatus("연결됨");
    log("스마트스토어 API 연결 성공", {
      고정IP프록시: data.proxyEnabled ? "사용 중" : "미사용",
      오늘변경주문: data.todayChangedCount,
      응답키: data.rawKeys,
    });
  } catch (error) {
    setStatus("연결 실패");
    log("스마트스토어 API 연결 실패", error.message);
  }
});

document.querySelector("#fetchOrders").addEventListener("click", async () => {
  resetOrderWorkflow();
  setStatus("조회 중");
  log("신규주문(발주 전) 주문 조회를 시작합니다.");
  try {
    const payload = {
      ...credentialPayload(),
      dateFrom: document.querySelector("#dateFrom").value,
      dateTo: document.querySelector("#dateTo").value,
      maxOrders: Number(document.querySelector("#maxOrders").value || 300),
    };
    const data = await postJson("/api/orders/fetch", payload);
    currentOrders = data.preview || [];
    confirmedOrders = data.confirmedPreview || [];
    setOrderView(currentOrders.length > 0 ? "new" : "confirmed");
    currentDownload = { url: data.downloadUrl, name: data.downloadName };
    placeOrdersButton.disabled = currentOrders.length === 0;
    if (currentDownload.url) {
      downloadLink.href = currentDownload.url;
      downloadLink.download = currentDownload.name;
      downloadLink.classList.remove("hidden");
    } else {
      downloadLink.classList.add("hidden");
      downloadLink.removeAttribute("href");
      downloadLink.removeAttribute("download");
    }
    setStatus("완료");
    log(`주문조회 완료: ${data.rows}건`, {
      신규주문: data.newRows ?? currentOrders.length,
      발주확인완료: data.confirmedRows ?? confirmedOrders.length,
      다음단계: currentOrders.length > 0 ? "발주확인 후 CJ 엑셀 다운로드" : "이미 발주확인된 건은 CJ 엑셀 다운로드 가능",
    });
  } catch (error) {
    setStatus("오류");
    log("주문조회 실패", error.message);
  }
});

document.querySelector("#placeOrders").addEventListener("click", async () => {
  clearOrderError();
  if (currentOrders.length === 0) {
    log("발주확인 대기", "먼저 주문조회를 해주세요.");
    return;
  }
  setStatus("발주확인 중");
  log(`스마트스토어 발주확인을 시작합니다: ${currentOrders.length}건`);
  try {
    const data = await postJson("/api/orders/place-order", {
      ...credentialPayload(),
      records: currentOrders,
    });
    if (data.failed > 0) {
      setStatus("확인 필요");
      const message = `발주확인 실패 포함: 성공 ${data.confirmed}건 / 실패 ${data.failed}건`;
      showOrderError(message);
      log(message, data.results);
      return;
    }
    confirmedOrders = confirmedOrders.concat(currentOrders.map((row) => ({
      ...row,
      orderWorkflowStatus: "발주확인 완료",
      confirmedAt: new Date().toISOString(),
    })));
    currentOrders = [];
    setOrderView("confirmed");
    placeOrdersButton.disabled = true;
    currentDownload = { url: data.downloadUrl, name: data.downloadName };
    if (currentDownload.url) {
      downloadLink.href = currentDownload.url;
      downloadLink.download = currentDownload.name;
      downloadLink.classList.remove("hidden");
    }
    setStatus("완료");
    log(`발주확인 완료: 성공 ${data.confirmed}건 / 실패 ${data.failed}건`, data.downloadName || data.results);
  } catch (error) {
    setStatus("오류");
    showOrderError(error.message);
    log("발주확인 실패", error.message);
  }
});

document.querySelector("#downloadOutboundList").addEventListener("click", () => {
  log("출고 리스트 다운로드", "출고 리스트 양식 샘플 등록 후 사용 가능합니다.");
  alert("출고 리스트 양식 샘플 등록 후 사용 가능합니다.");
});

prevOrdersPageButton.addEventListener("click", () => {
  currentOrderPage -= 1;
  renderOrdersPage();
});

nextOrdersPageButton.addEventListener("click", () => {
  currentOrderPage += 1;
  renderOrdersPage();
});

newOrdersTab.addEventListener("click", () => setOrderView("new"));
confirmedOrdersTab.addEventListener("click", () => setOrderView("confirmed"));

document.querySelector("#trackingForm").addEventListener("submit", async (event) => {
  event.preventDefault();
  setStatus("확인 중");
  try {
    const data = await postForm("/api/tracking/upload", event.currentTarget);
    renderTracking(data.records || []);
    await loadShipments();
    setStatus("완료");
    log(`송장번호 ${data.rows}건을 읽었습니다.`, data.records.slice(0, 10));
  } catch (error) {
    setStatus("오류");
    log("송장 파일 읽기 실패", error.message);
  }
});

document.querySelector("#sendDispatch").addEventListener("click", async () => {
  if (currentTrackingRecords.length === 0) {
    log("스마트스토어 입력 대기", "먼저 CJ 송장 결과 파일을 읽어주세요.");
    return;
  }
  setStatus("전송 중");
  log(`스마트스토어 운송장 입력을 시작합니다: ${currentTrackingRecords.length}건`);
  try {
    const data = await postJson("/api/dispatch/send", {
      ...credentialPayload(),
      records: currentTrackingRecords,
    });
    currentTrackingRecords = data.records || currentTrackingRecords;
    renderTracking(currentTrackingRecords);
    await loadShipments();
    setStatus("완료");
    log(`스마트스토어 운송장 입력 완료: 성공 ${data.sent}건 / 실패 ${data.failed}건`, data.results);
  } catch (error) {
    setStatus("오류");
    log("스마트스토어 운송장 입력 실패", error.message);
  }
});

async function loadShipments() {
  const data = await postJson("/api/shipments/list", {
    dateFrom: document.querySelector("#shipDateFrom").value,
    dateTo: document.querySelector("#shipDateTo").value,
    query: document.querySelector("#shipmentQuery").value.trim(),
    status: document.querySelector("#shipmentStatus").value,
  });
  renderShipments(data.rows || [], data.summary);
}

document.querySelector("#loadShipments").addEventListener("click", async () => {
  try {
    await loadShipments();
    log("출고현황을 조회했습니다.");
  } catch (error) {
    log("출고현황 조회 실패", error.message);
  }
});

document.querySelector("#shipmentQuery").addEventListener("keydown", async (event) => {
  if (event.key === "Enter") {
    event.preventDefault();
    await loadShipments();
  }
});

document.querySelectorAll(".tab").forEach((button) => {
  button.addEventListener("click", () => {
    const tab = button.dataset.tab;
    document.querySelectorAll(".tab").forEach((item) => item.classList.toggle("active", item === button));
    document.querySelector("#tabShipments").classList.toggle("active", tab === "shipments");
    document.querySelector("#tabLogs").classList.toggle("active", tab === "logs");
  });
});

loadShipments().catch(() => {});
