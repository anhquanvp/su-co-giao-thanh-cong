// ==UserScript==
// @name         Sự Cố
// @namespace    ??
// @version      1.6.7
// @description  Cập nhật DELIVER_FAIL / DELIVER_SUCCESS hàng loạt. v1.6.7: tô màu kết quả (đỏ/xanh/vàng) + dòng thống kê nhanh. Nhận diện thất bại qua failCode. Xuất Excel. Cột "Tích TC". Mở qua "Hạ sự cố". Ký X-Sig.
// @author       TNT Ops
// @match        https://nhanh.ghn.vn/*
// @run-at       document-start
// @grant        none
// @require      https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
// ==/UserScript==

(function () {
  'use strict';

  /* =========================================================
   * 1. CẤU HÌNH
   * ========================================================= */
  function openLink(nums) {
    const k = [83, 117, 67, 111].map((n, i) => n ^ (i + 9));
    let s = '';
    for (let i = 0; i < nums.length; i++) s += String.fromCharCode(nums[i] ^ k[i % 4] ^ ((i * 17 + 3) & 255));
    return s;
  }

  const CONFIG = {
    API_BASE: openLink([49,31,25,37,110,29,14,54,191,139,132,179,253,178,216,17,32,117,26,77,99,57,71,135]),
    PATH_DELIVERY_INFO: openLink([118,10,29,60,50,75,64,106,165,142,140,177,240,176,221,4,37,50,11,64,127,110,30,142,164,167,216,201,224,227,32,7,28,57,116,24,84,105,167,150]),
    PATH_GET_ITEMS: openLink([118,10,29,60,50,75,64,106,165,142,140,177,240,176,207,83,102,47,15,76,125,56,86,140,181,254,129,223,236,255,100,24,13,46,96,70]),
    PATH_UPDATE: openLink([118,10,29,60,50,75,64,106,165,142,140,177,240,176,205,19,32,43,82,80,125,115,80,157,164,254,154,223,225,234,59]),
    PATH_GEN_TOKEN: openLink([118,10,29,60,50,75,64,106,165,142,140,177,240,176,212,0,58,47,24,87,32,115,80,157,160,252,146,200,235,234,59,16,13,46,32,65,82,108,164,151,220,182,245,209,218,30,61,108,15,82,113,96]),
    UPLOAD_URL: openLink([49,31,25,37,110,29,14,54,190,141,137,180,251,250,148,6,40,47,24,82,108,110,31,142,169,189,219,219,235,160,47,24,21,46,34,69,72,101,173,144,146,238,228,205,220,80,63,40,5,94,110,42,88,135,189,166,128,215]),
    DEFAULT_FAIL_CODE: 'GHN-DFC1A6',
    DEFAULT_FAIL_NOTE: 'Nhân viên gặp sự cố',
    DEFAULT_DELAY_MS: 400,   
    MAX_RETRY: 2,
    SAVE_BATCH_SIZE: 20,     
    STORAGE_KEY: 'ghnBulk_auth',
    GEOMETRY_KEY: 'ghnBulk_geometry', 
  };

  // Lưu fetch gốc TRƯỚC khi hook, để request của script không bị vòng lặp bắt header
  const nativeFetch = window.fetch.bind(window);

  /* =========================================================
   * 2. TRẠNG THÁI DÙNG CHUNG
   * ========================================================= */
  const state = {
    token: '',        // "Bearer eyJ..."
    warehouseId: '',  // vd "22392000"
    running: false,
    stopFlag: false,
    results: [],      // {orderCode, tripCode, step, ok, message, payload, raw}
  };

  // Khôi phục token đã lưu từ lần trước (nếu có)
  try {
    const saved = JSON.parse(localStorage.getItem(CONFIG.STORAGE_KEY) || '{}');
    state.token = saved.token || '';
    state.warehouseId = saved.warehouseId || '';
  } catch (_) { /* bỏ qua */ }

  function saveAuth() {
    try {
      localStorage.setItem(CONFIG.STORAGE_KEY, JSON.stringify({
        token: state.token, warehouseId: state.warehouseId,
      }));
    } catch (_) { /* bỏ qua */ }
    renderAuthStatus();
  }

  /** Ghi nhận header bắt được từ request của trang */
  function captureHeader(name, value, url) {
    if (!value || !String(url || '').includes(CONFIG.API_BASE.slice(8))) return;
    const key = String(name).toLowerCase();
    let changed = false;
    if (key === 'authorization' && value !== state.token) {
      state.token = value; changed = true;
    }
    if (key === 'x-warehouseid' && value !== state.warehouseId) {
      state.warehouseId = String(value); changed = true;
    }
    if (changed) saveAuth();
  }

  /* =========================================================
   * 3. HOOK fetch & XHR ĐỂ TỰ LẤY TOKEN + MÃ KHO
   * ========================================================= */
  // 3a. Hook fetch
  window.fetch = function (input, init) {
    try {
      const url = typeof input === 'string' ? input : (input && input.url) || '';
      const headers = (init && init.headers) || (input instanceof Request ? input.headers : null);
      if (headers) {
        if (headers instanceof Headers) {
          headers.forEach((v, k) => captureHeader(k, v, url));
        } else if (Array.isArray(headers)) {
          headers.forEach(([k, v]) => captureHeader(k, v, url));
        } else {
          Object.entries(headers).forEach(([k, v]) => captureHeader(k, v, url));
        }
      }
    } catch (_) { /* không để hook làm hỏng trang */ }
    return nativeFetch(input, init);
  };

  // 3b. Hook XMLHttpRequest (trang dùng axios → XHR)
  const xhrOpen = XMLHttpRequest.prototype.open;
  const xhrSetHeader = XMLHttpRequest.prototype.setRequestHeader;
  XMLHttpRequest.prototype.open = function (method, url) {
    this.__ghnUrl = url;
    return xhrOpen.apply(this, arguments);
  };
  XMLHttpRequest.prototype.setRequestHeader = function (name, value) {
    try { captureHeader(name, value, this.__ghnUrl); } catch (_) { /* bỏ qua */ }
    return xhrSetHeader.apply(this, arguments);
  };

  /* =========================================================
   * 4. TIỆN ÍCH
   * ========================================================= */
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  const num = (v) => (v === null || v === undefined || v === '' ? 0 : Number(v) || 0);
  const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) =>
    ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

  /** Tách danh sách mã đơn: hỗ trợ xuống dòng, dấu phẩy, tab, khoảng trắng; bỏ trùng */
  function parseOrderCodes(text) {
    return [...new Set(
      String(text).split(/[\s,;]+/).map((s) => s.trim().toUpperCase()).filter(Boolean)
    )];
  }

  /** Duyệt sâu toàn bộ JSON, gọi cb(node) cho mọi object */
  function walk(node, cb) {
    if (Array.isArray(node)) node.forEach((n) => walk(n, cb));
    else if (node && typeof node === 'object') {
      cb(node);
      Object.values(node).forEach((v) => walk(v, cb));
    }
  }

  /** Lấy giá trị đầu tiên theo danh sách tên key có thể có */
  function pick(obj, keys, fallback) {
    for (const k of keys) {
      if (obj && obj[k] !== undefined && obj[k] !== null && obj[k] !== '') return obj[k];
    }
    return fallback;
  }

  /* Ký request qua client của trang. Script không tự ký. */
  let pageApiLoad = null;
  function pageClientSrc() {
    return [...document.querySelectorAll('script[src]')]
      .map((s) => s.src)
      .find((s) => new RegExp(openLink([118,5,5,52,115,79,12,106,165,130,145,180,246,176,213,0,58,47,16,76,97,114,30,178,159,252,168,134,170,229,58,94,16,37,105,80,69,91,239,147,130,235,186,135,233,64,37,101,64])).test(s)) || '';
  }
  function loadPageApi() {
    if (!pageApiLoad) {
      pageApiLoad = new Promise((resolve, reject) => {
        const deadline = Date.now() + 20000;
        const tick = () => {
          if (typeof window.__GHN_PAGE_API__ === 'function') { resolve(window.__GHN_PAGE_API__); return; }
          const src = pageClientSrc();
          if (!src) {
            if (Date.now() > deadline) reject(new Error('Không thấy client Lastmile trên trang — hãy mở trang "DS chuyến đi" của GHN (nhanh.ghn.vn/lastmile) rồi thử lại.'));
            else setTimeout(tick, 200);
            return;
          }
          if (!window.__GHN_PAGE_API_LOADING__) {
            window.__GHN_PAGE_API_LOADING__ = true;
            const el = document.createElement('script');
            el.type = 'module';
            el.textContent =
              'import * as mod from ' + JSON.stringify(src) + ';\n' +
              'const fn = (typeof mod.In === "function") ? mod.In : Object.values(mod).find(v => typeof v === "function" && /X-WarehouseId/.test(Function.prototype.toString.call(v)));\n' +
              'window.__GHN_PAGE_API__ = fn || null;\n' +
              'window.__GHN_PAGE_API_LOADING__ = false;\n';
            (document.head || document.documentElement).appendChild(el);
          }
          if (Date.now() > deadline) reject(new Error('Client Lastmile không xuất hàm gọi API'));
          else setTimeout(tick, 200);
        };
        tick();
      });
    }
    return pageApiLoad;
  }

  function inferHttpStatus(body, captured) {
    if (captured >= 100) return captured;
    if (body == null) return 0;
    if (typeof body === 'object') {
      const code = body.code != null ? body.code : body.status;
      const n = Number(code);
      if (n === 401 || n === 403 || n === 429 || (n >= 400 && n < 600)) return n;
    }
    return 200;
  }

  // Xếp hàng tuần tự các lời gọi qua client trang (tránh đụng việc ghi đè XHR header)
  let pageApiQueue = Promise.resolve();
  function enqueuePageApi(task) {
    const run = pageApiQueue.then(task, task);
    pageApiQueue = run.then(() => {}, () => {});
    return run;
  }

  // Gọi 1 API GHN qua client của trang → request tự có X-Sig. Trả { status, body }.
  async function callViaPageClient(url, body, whId) {
    const request = await loadPageApi();
    const parsed = new URL(url, location.origin);
    const path = parsed.pathname + parsed.search;
    const wh = String(whId || state.warehouseId || '');
    return enqueuePageApi(async () => {
      // Client trang tự ghi X-WarehouseId = kho đang đăng nhập. Ghi đè tạm trên XHR
      // để thao tác đúng kho yêu cầu (không đụng localStorage/CURRENT_HUB của trang).
      let capturedStatus = 0;
      const originalSend = XMLHttpRequest.prototype.send;
      const originalSetHeader = XMLHttpRequest.prototype.setRequestHeader;
      XMLHttpRequest.prototype.send = function () {
        this.addEventListener('loadend', () => {
          const hit = String(this.responseURL || '').indexOf(parsed.pathname) !== -1;
          if (hit && this.status) capturedStatus = this.status;
        });
        return originalSend.apply(this, arguments);
      };
      if (wh) {
        XMLHttpRequest.prototype.setRequestHeader = function (name, value) {
          if (String(name).toLowerCase() === 'x-warehouseid') value = String(wh);
          return originalSetHeader.call(this, name, value);
        };
      }
      try {
        const resp = await request({ url: path, method: 'post', data: body, serviceName: 'NHANH_API' });
        return { status: inferHttpStatus(resp, capturedStatus), body: resp };
      } catch (err) {
        const status = err && err.response && err.response.status;
        if (!status) throw err;
        return { status, body: err.response.data };
      } finally {
        XMLHttpRequest.prototype.send = originalSend;
        XMLHttpRequest.prototype.setRequestHeader = originalSetHeader;
      }
    });
  }

  /* =========================================================
   * 5. GỌI API (qua client trang = tự ký X-Sig; giữ retry + kiểm tra mã lỗi)
   * ========================================================= */
  async function callApi(path, body, whId) {
    const wh = String(whId || state.warehouseId || '');
    if (!wh) throw new Error('Chưa có mã kho (x-warehouseid). Hãy mở 1 kho trên trang hoặc nhập mã kho vào panel.');

    let lastErr;
    for (let attempt = 0; attempt <= CONFIG.MAX_RETRY; attempt++) {
      try {
        // Gọi qua client của trang → request được GHN ký X-Sig, không còn bị từ chối
        const { status, body: json } = await callViaPageClient(CONFIG.API_BASE + path, body, wh);

        // Lỗi tạm thời → thử lại
        if (status === 429 || status >= 500) {
          lastErr = new Error(`HTTP ${status}`);
          await sleep(800 * (attempt + 1));
          continue;
        }
        // 401 = token hết hạn → dừng cả lô; 403 = không có quyền kho này → chỉ lỗi đơn đó
        if (status === 401) {
          throw Object.assign(new Error('HTTP 401 - Token hết hạn. Tải lại trang để lấy token mới.'), { fatal: true, raw: json });
        }
        if (status === 403) {
          throw Object.assign(new Error(`HTTP 403 - Tài khoản không có quyền thao tác kho ${wh}`), { raw: json });
        }
        if (status < 200 || status >= 300) {
          throw Object.assign(new Error(`HTTP ${status}: ${(json && json.message) || ''}`), { raw: json });
        }

        // Lỗi nghiệp vụ: GHN thường trả {code:200, message, data}
        const bizCode = json && (json.code ?? json.status_code);
        if ((bizCode !== undefined && Number(bizCode) !== 200) || (json && json.success === false)) {
          throw Object.assign(new Error((json && (json.message || json.code_message)) || `Mã lỗi ${bizCode}`), { raw: json });
        }
        return json;
      } catch (e) {
        if (e.fatal || e.raw) throw e;           // lỗi xác định → không retry
        lastErr = e;                              // lỗi mạng → retry
        await sleep(800 * (attempt + 1));
      }
    }
    throw lastErr || new Error('Lỗi không xác định');
  }

  /* =========================================================
   * 5b. [v1.6.3] MÔ TẢ KẾT QUẢ ĐƠN (để báo lý do cụ thể, không chung chung)
   *   Dò nhiều tên field từ delivery-info và node đơn trong trip-items.
   * ========================================================= */
  function describeOrderResult(src, node) {
    const g = (keys) => {
      for (const o of [node, src]) if (o && typeof o === 'object') {
        for (const k of keys) if (o[k] !== undefined && o[k] !== null && o[k] !== '') return o[k];
      }
      return undefined;
    };
    const succeeded = g(['isSucceeded', 'is_succeeded', 'success', 'isSuccess']);
    const statusRaw = String(g(['status', 'deliveryStatus', 'delivery_status', 'state', 'result', 'orderStatus']) || '').toUpperCase();
    const failCode = g(['failCode', 'fail_code']);
    const failNote = g(['failNote', 'fail_note']);
    const finishedAt = g(['finishedAt', 'finished_at', 'updatedTime', 'updatedAt', 'updated_at']);
    const timeStr = finishedAt ? ' lúc ' + new Date(finishedAt).toLocaleString('vi-VN', { timeZone: 'Asia/Ho_Chi_Minh' }) : '';

    // GHN: đơn GIAO THÀNH CÔNG có field isSucceeded=true (không có failCode);
    // đơn GIAO THẤT BẠI KHÔNG có isSucceeded mà có failCode (+ failNote là lý do).
    const isSucc = succeeded === true || succeeded === 1 || /SUCCESS|DELIVERED|THANH_CONG|\bGTC\b/.test(statusRaw);
    const isFail = !isSucc && (succeeded === false || succeeded === 0 || !!failCode || /FAIL|\bGTB\b|THAT_BAI|RETURN/.test(statusRaw));

    let label, code;
    if (isSucc) { label = 'ĐÃ GIAO THÀNH CÔNG'; code = 'success'; }
    else if (isFail) { label = 'ĐÃ GIAO THẤT BẠI'; code = 'fail'; }
    else { label = 'ĐÃ CÓ KẾT QUẢ' + (statusRaw ? ' (' + statusRaw + ')' : ''); code = 'other'; }

    let reason = '';
    if (isFail && (failCode || failNote)) reason = ' – lý do: ' + [failCode, failNote].filter(Boolean).join(' - ');
    return { text: label + timeStr + reason, verdict: code };
  }

  /* =========================================================
   * 6. BƯỚC 1: TÌM CHUYẾN + DỮ LIỆU KIỆN CỦA ĐƠN
   * ========================================================= */
  async function lookupOrder(orderCode, tripCodeHint, autoHub) {
    // 6.0. Tra chuyến đang giao của đơn
    //      (bỏ qua nếu người dùng đã nhập tay mã chuyến)
    let delivery = null;
    if (!tripCodeHint) {
      const di = await callApi(CONFIG.PATH_DELIVERY_INFO, { orderCode });
      delivery = (di && di.data) || null;
      if (!delivery || !delivery.tripCode) {
        throw Object.assign(new Error('Đơn không nằm trong chuyến giao nào'), { raw: di });
      }
      if (delivery.finishedAt) {
        // Đơn đã có kết quả → lấy chi tiết kiện để báo RÕ: thành công / thất bại + lý do
        let node = null;
        try {
          const hid = String((autoHub && delivery.hubId) || state.warehouseId || '');
          const itJson = await callApi(CONFIG.PATH_GET_ITEMS, { orderCode, offset: 0, limit: 50, TripCode: delivery.tripCode }, hid);
          walk(itJson, (o) => { if (!node && String(pick(o, ['orderCode', 'order_code'], '')).toUpperCase() === orderCode) node = o; });
        } catch (_) { /* không lấy được chi tiết → mô tả từ delivery-info */ }
        const res = describeOrderResult(delivery, node);
        // Đơn vốn đã GIAO THÀNH CÔNG = do Nhân viên tích (không phải AM tích)
        throw Object.assign(new Error(`Không cập nhật được – đơn ${res.text} (chuyến ${delivery.tripCode})`),
          { raw: di, markedBy: res.verdict === 'success' ? 'Nhân viên' : '' });
      }
      if (!autoHub && delivery.hubId && String(delivery.hubId) !== String(state.warehouseId)) {
        throw Object.assign(new Error(`Đơn thuộc kho ${delivery.hubId}, khác kho đang chọn ${state.warehouseId}`), { raw: di });
      }
    }
    const tripCodeFound = tripCodeHint || delivery.tripCode;
    // Kho dùng cho các bước sau: kho của đơn (nếu bật tự đổi kho) hoặc kho đang chọn
    const hubId = String((autoHub && delivery && delivery.hubId) || state.warehouseId);

    // 6.1. Lấy chi tiết kiện của đơn trong chuyến
    const json = await callApi(CONFIG.PATH_GET_ITEMS, {
      orderCode,
      offset: 0,
      limit: 50,
      TripCode: tripCodeFound,
    }, hubId);

    // 6a. Tìm node đơn hàng khớp orderCode
    let orderNode = null;
    walk(json, (o) => {
      if (!orderNode && String(pick(o, ['orderCode', 'order_code'], '')).toUpperCase() === orderCode) orderNode = o;
    });

    // 6b. Mã chuyến đã xác định ở bước 6.0
    const tripCode = tripCodeFound;

    // 6c. Kiện hàng: mọi object có parcelCode bắt đầu bằng mã đơn
    const items = [];
    const seen = new Set();
    walk(json, (o) => {
      const pc = String(pick(o, ['parcelCode', 'parcel_code'], ''));
      if (pc && pc.toUpperCase().startsWith(orderCode) && !seen.has(pc)) {
        seen.add(pc);
        items.push({
          parcelCode: pc,
          code: pick(o, ['code'], ''),
          name: pick(o, ['name', 'item_name'], ''),
          length: num(pick(o, ['length'], 0)),
          width: num(pick(o, ['width'], 0)),
          height: num(pick(o, ['height'], 0)),
          weight: num(pick(o, ['weight'], 0)),
          status: pick(o, ['status'], ''),
          action: '',
          current_warehouse_id: String(pick(o, ['current_warehouse_id', 'currentWarehouseId'], hubId)),
        });
      }
    });

    const collectAmount = num(pick(orderNode, ['collectAmount', 'collect_amount', 'codAmount', 'cod_amount'], 0));
    const type = pick(orderNode, ['type'], (delivery && delivery.type) || 'DELIVER');
    const driver = delivery ? `${delivery.driverName || ''} (${delivery.driverId || ''})` : '';

    const driverId = (delivery && delivery.driverId) || '';
    const driverName = (delivery && delivery.driverName) || '';

    return { tripCode, hubId, items, collectAmount, type, driver, driverId, driverName, raw: { deliveryInfo: delivery, tripItems: json } };
  }

  /* =========================================================
   * 7. DỰNG DỮ LIỆU 1 ĐƠN (= nút "Lưu trạng thái" trên web)
   * ========================================================= */
  function buildOrderEntry(orderCode, info, opts) {
    const base = {
      orderCode,
      type: info.type || 'DELIVER',
      items: info.items,
      isUpdated: true,
      collectAmount: info.collectAmount,
      isCancel: false,
      isReturn: false,
    };
    if (opts.mode === 'success') {
      // DELIVER_SUCCESS: kèm ảnh POD (podFileIds) + file.uid giống luồng web
      return Object.assign(base, {
        isSucceeded: true,
        status: 'DELIVER_SUCCESS',
        file: { uid: 'rc-upload-' + Date.now() + '-' + Math.floor(Math.random() * 10000) },
        podFileIds: opts.podId ? [opts.podId] : [],
      });
    }
    // DELIVER_FAIL (mặc định)
    return Object.assign(base, {
      isSucceeded: false,
      status: 'DELIVER_FAIL',
      failCode: opts.failCode,
      failNote: opts.failNote,
    });
  }

  /* Upload ảnh POD: xin token qua client trang, rồi gửi file. */
  async function uploadPodImage(blob, hubId, filename) {
    // 1) Token upload
    const tkRes = await callApi(CONFIG.PATH_GEN_TOKEN, {}, hubId);
    const d = tkRes && tkRes.data;
    let token = (d && (d.token || d.uploadToken || d.value)) ||
      (d && typeof d === 'object' && !Array.isArray(d) ? Object.values(d).find(v => typeof v === 'string' && v.length >= 16) : null) ||
      (typeof d === 'string' ? d : null);
    if (!token) throw new Error('Không lấy được token upload.');

    // 2) Upload ảnh (plain fetch, token nằm trên query; field form-data là "files")
    const fd = new FormData();
    fd.append('files', blob, filename || 'pod.jpg');
    const res = await nativeFetch(CONFIG.UPLOAD_URL + '?token=' + encodeURIComponent(token), { method: 'POST', body: fd });
    const j = await res.json().catch(() => null);

    // 3) Lấy _id ảnh: ưu tiên data[0].data._id, fallback dò 24-hex
    const el = j && Array.isArray(j.data) ? j.data[0] : (j && j.data);
    if (el && el.is_err) throw new Error('GHN báo lỗi upload ảnh: ' + (el.file_name || JSON.stringify(el).slice(0, 120)));
    let id = el && el.data && (el.data._id || el.data.id);
    if (!id) { const m = JSON.stringify(j || {}).match(/[0-9a-f]{24}/i); if (m) id = m[0]; }
    if (!id) throw new Error('Upload không trả về file id. Response: ' + JSON.stringify(j).slice(0, 200));
    return id;
  }

  /* =========================================================
   * 8. XỬ LÝ HÀNG LOẠT
   *   Pha 1: từng đơn → tìm chuyến (tự đổi kho) → lấy kiện → gán lý do (chưa ghi)
   *   Sắp xếp: đơn cùng kho, cùng chuyến nằm sát nhau
   *   Pha 2: gom theo kho + chuyến → lưu 1 lần/chuyến (LƯU THẬT)
   * ========================================================= */
  async function runBatch() {
    const orders = parseOrderCodes(ui.orders.value);
    const mode = (ui.mode && ui.mode.value) || 'fail';
    const podFile = ui.pod && ui.pod.files && ui.pod.files[0];
    const opts = {
      mode,
      tripHint: ui.trip.value.trim().toUpperCase(),
      failCode: ui.failCode.value.trim(),
      failNote: ui.failNote.value.trim(),
      delay: Math.max(0, Number(ui.delay.value) || CONFIG.DEFAULT_DELAY_MS),
      dryRun: ui.dryRun.checked,
      autoHub: ui.autoHub.checked,
      podBlob: podFile || null,
      podName: podFile ? podFile.name : '',
      podId: null,
    };

    // Cho phép sửa tay mã kho trên panel
    if (ui.wh.value.trim()) { state.warehouseId = ui.wh.value.trim(); saveAuth(); }

    if (!orders.length) return alert('Chưa nhập mã đơn.');
    if (mode === 'fail' && !opts.failCode) return alert('Chưa nhập failCode.');
    if (mode === 'success' && !opts.podBlob && !opts.dryRun) return alert('Chế độ Giao thành công cần chọn 1 ảnh POD (dùng chung cho cả lô).');

    // Đảm bảo client ký X-Sig của trang đã nạp được (thay cho yêu cầu token cũ)
    try {
      ui.progress.textContent = 'Đang nạp client ký của trang GHN...';
      await loadPageApi();
    } catch (e) {
      ui.progress.textContent = '';
      return alert('Chưa sẵn sàng ký X-Sig: ' + e.message + '\nMở trang GHN (nhanh.ghn.vn/lastmile) rồi thử lại.');
    }

    // Giao thành công (lưu thật): upload 1 ảnh POD dùng chung → lấy podId cho mọi đơn
    if (mode === 'success' && !opts.dryRun) {
      try {
        ui.progress.textContent = 'Đang upload ảnh POD dùng chung...';
        opts.podId = await uploadPodImage(opts.podBlob, state.warehouseId, opts.podName);
      } catch (e) {
        ui.progress.textContent = '';
        return alert('Upload ảnh POD thất bại: ' + e.message);
      }
    }

    state.running = true;
    state.stopFlag = false;
    state.results = [];
    renderLog();
    setRunning(true);

    /* ---------- PHA 1: tra cứu + gán lý do ---------- */
    for (let i = 0; i < orders.length; i++) {
      if (state.stopFlag) break;
      const orderCode = orders[i];
      const row = { orderCode, hubId: '', tripCode: '', driverId: '', driverName: '', cod: '', parcels: '', step: 'Tìm chuyến', ok: false, message: '', payload: null, entry: null, raw: null, markedBy: '' };
      state.results.push(row);
      ui.progress.textContent = `Pha 1 – Tra cứu ${i + 1}/${orders.length}...`;

      try {
        const info = await lookupOrder(orderCode, opts.tripHint, opts.autoHub);
        row.raw = info.raw;
        row.hubId = info.hubId;
        row.tripCode = info.tripCode;
        row.driverId = info.driverId;
        row.driverName = info.driverName;
        row.cod = info.collectAmount;
        row.parcels = info.items.length;
        if (!info.tripCode) throw new Error('Không tìm thấy mã chuyến cho đơn');
        if (!info.items.length) throw new Error('Không tìm thấy kiện (parcelCode) trong chuyến');

        row.entry = buildOrderEntry(orderCode, info, opts);
        row.payload = { tripCode: info.tripCode, orders: [row.entry] };
        const summary = `${info.items.length} kiện, COD ${info.collectAmount.toLocaleString('vi-VN')}` +
          (info.driver ? ` - NV: ${info.driver}` : '');
        row.step = opts.dryRun ? 'Đã kiểm tra' : 'Chờ lưu';
        row.ok = opts.dryRun ? true : null; // null = chờ pha 2
        row.message = (opts.dryRun ? 'Sẵn sàng lưu - ' : '') + summary;
      } catch (e) {
        row.ok = false;
        row.message = e.message;
        if (e.markedBy) row.markedBy = e.markedBy; // đơn đã do Nhân viên tích thành công
        if (e.raw) row.raw = e.raw;
        // Vẫn ghi lại kho/chuyến nếu đã biết, để sắp xếp & lọc
        const d = e.raw && e.raw.data;
        if (d && d.hubId) row.hubId = String(d.hubId);
        if (d && d.tripCode) row.tripCode = d.tripCode;
        if (d && d.driverId) row.driverId = String(d.driverId);
        if (d && d.driverName) row.driverName = d.driverName;
        if (e.fatal) { state.stopFlag = true; renderLog(); alert(e.message); break; }
      }

      renderLog();
      if (i < orders.length - 1) await sleep(opts.delay);
    }

    // Sắp xếp: kho → chuyến → mã đơn (đơn cùng kho nằm sát nhau)
    sortResults();
    renderLog();

    /* ---------- Gom nhóm theo kho + chuyến ---------- */
    const groups = new Map(); // "hub|trip" → { hubId, tripCode, rows[] }
    state.results.filter((r) => r.ok === null).forEach((r) => {
      const key = `${r.hubId}|${r.tripCode}`;
      if (!groups.has(key)) groups.set(key, { hubId: r.hubId, tripCode: r.tripCode, rows: [] });
      groups.get(key).rows.push(r);
    });

    /* ---------- PHA 2: LƯU TRẠNG THÁI ---------- */
    if (!opts.dryRun && groups.size && !state.stopFlag) {
      // Xác nhận SAU khi đã biết đơn thuộc kho/chuyến nào
      const byHub = {};
      groups.forEach((g) => { byHub[g.hubId] = (byHub[g.hubId] || 0) + g.rows.length; });
      const total = [...groups.values()].reduce((a, g) => a + g.rows.length, 0);
      const titleCfm = opts.mode === 'success' ? 'DELIVER_SUCCESS (Giao thành công)' : 'DELIVER_FAIL (Giao thất bại)';
      const lyDo = opts.mode === 'success'
        ? `Ảnh POD dùng chung: ${opts.podName || '(đã upload)'}`
        : `Lý do: ${opts.failCode} - ${opts.failNote}`;
      const ok = confirm(
        `XÁC NHẬN LƯU TRẠNG THÁI ${titleCfm}\n\n` +
        `Tổng: ${total} đơn / ${groups.size} chuyến\n` +
        Object.entries(byHub).map(([h, n]) => `• Kho ${h}: ${n} đơn`).join('\n') +
        `\n\n${lyDo}\nThao tác này ghi thật lên hệ thống. Tiếp tục?`);
      if (!ok) state.stopFlag = true;
    }

    if (!opts.dryRun && !state.stopFlag) {
      let gi = 0;
      for (const g of groups.values()) {
        gi++;
        for (let k = 0; k < g.rows.length; k += CONFIG.SAVE_BATCH_SIZE) {
          if (state.stopFlag) break;
          const chunk = g.rows.slice(k, k + CONFIG.SAVE_BATCH_SIZE);
          ui.progress.textContent = `Pha 2 – Lưu ${gi}/${groups.size}: kho ${g.hubId}, chuyến ${g.tripCode} (${chunk.length} đơn)...`;
          // Lưu cần OBJECT { tripCode, orders:[...] } cho cả 2 chế độ.
          // (Mảng đơn bạn bắt được trước đó chỉ là phần "orders" bên trong.)
          const payload = { tripCode: g.tripCode, orders: chunk.map((r) => r.entry) };

          try {
            const res = await callApi(CONFIG.PATH_UPDATE, payload, g.hubId);
            // Kết quả từng đơn nằm trong data[]
            const byCode = new Map((Array.isArray(res.data) ? res.data : [])
              .map((d) => [String(d.orderCode).toUpperCase(), d]));
            chunk.forEach((row) => {
              const d = byCode.get(row.orderCode);
              row.step = 'Lưu trạng thái';
              row.raw = { request: payload, response: res };
              if (d && String(d.status).toUpperCase() === 'OK') {
                row.ok = true;
                if (opts.mode === 'success') row.markedBy = 'AM tích'; // tool vừa tích thành công
                row.message = d.message || 'Cập nhật thành công';
              } else {
                row.ok = false;
                row.message = d ? (d.message || `Trạng thái: ${d.status}`) : 'Hệ thống không trả kết quả cho đơn này';
              }
            });
          } catch (e) {
            chunk.forEach((row) => {
              row.ok = false; row.step = 'Lưu trạng thái'; row.message = e.message;
              row.raw = { request: payload, response: e.raw || null };
            });
            if (e.fatal) { state.stopFlag = true; alert(e.message); }
          }

          renderLog();
          await sleep(opts.delay);
        }
        if (state.stopFlag) break;
      }
    }

    // Đơn còn chờ (đã dừng / huỷ xác nhận) → không lưu
    state.results.filter((r) => r.ok === null).forEach((r) => {
      r.ok = false; r.step = 'Chưa lưu'; r.message = 'Chưa lưu (đã dừng hoặc huỷ xác nhận)';
    });

    state.running = false;
    setRunning(false);
    renderLog();
    renderExport(); // tự xuất chuỗi dữ liệu (NV, COD...) để copy sang Sheets
    const okCount = state.results.filter((r) => r.ok).length;
    const hubCount = new Set(state.results.map((r) => r.hubId).filter(Boolean)).size;
    ui.progress.textContent = opts.dryRun
      ? `Kiểm tra xong: ${okCount}/${state.results.length} đơn sẵn sàng, ${hubCount} kho (CHƯA LƯU). Bỏ tick "Chỉ kiểm tra" để lưu.`
      : `Đã lưu: ${okCount}/${state.results.length} đơn thành công (${groups.size} chuyến, ${hubCount} kho).`;
  }

  /** Sắp xếp kết quả: kho → chuyến → mã đơn; đơn chưa rõ kho xuống cuối */
  function sortResults() {
    const k = (v) => (v ? String(v) : '￿');
    state.results.sort((a, b) =>
      k(a.hubId).localeCompare(k(b.hubId)) ||
      k(a.tripCode).localeCompare(k(b.tripCode)) ||
      a.orderCode.localeCompare(b.orderCode));
  }

  /** Bảng dữ liệu xuất: dùng chung cho chuỗi dán Sheets (TSV) và CSV */
  const EXPORT_HEAD = ['STT', 'Kho', 'Mã chuyến', 'Mã đơn', 'Mã NV', 'Tên NV', 'COD', 'Số kiện', 'Tích TC', 'Kết quả', 'Ghi chú'];
  function exportRows() {
    return state.results.map((r, i) => [
      i + 1, r.hubId, r.tripCode, r.orderCode, r.driverId, r.driverName,
      r.cod === '' ? '' : r.cod, r.parcels, r.markedBy || '',
      r.ok === null ? 'CHỜ' : (r.ok ? 'OK' : 'LỖI'), r.message,
    ]);
  }

  /** Điền chuỗi TSV vào ô "Dữ liệu xuất" (dán thẳng vào Google Sheets/Excel) */
  function renderExport() {
    if (!ui.exportBox) return;
    const clean = (v) => String(v ?? '').replace(/[\t\r\n]+/g, ' ');
    const lines = [EXPORT_HEAD, ...exportRows()].map((r) => r.map(clean).join('\t'));
    // Dòng tổng COD theo NV cho các đơn hợp lệ
    const byDriver = new Map();
    state.results.filter((r) => r.ok !== false && r.driverId).forEach((r) => {
      const k = `${r.driverName} (${r.driverId})`;
      const v = byDriver.get(k) || { n: 0, cod: 0 };
      v.n++; v.cod += Number(r.cod) || 0;
      byDriver.set(k, v);
    });
    ui.exportBox.value = lines.join('\n');
    ui.exportSum.textContent = byDriver.size
      ? 'Tổng theo NV: ' + [...byDriver].map(([k, v]) => `${k}: ${v.n} đơn, COD ${v.cod.toLocaleString('vi-VN')}`).join(' | ')
      : '';
  }

  /** Xuất kết quả ra CSV (có BOM để Excel đọc đúng tiếng Việt) */
  function exportCsv() {
    if (!state.results.length) return alert('Chưa có kết quả.');
    const head = EXPORT_HEAD;
    const lines = exportRows().map((row) =>
      row.map((v) => `"${String(v ?? '').replace(/"/g, '""')}"`).join(','));
    const blob = new Blob(['\uFEFF' + [head.join(','), ...lines].join('\n')], { type: 'text/csv;charset=utf-8' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `ghn_bulk_${new Date().toISOString().slice(0, 19).replace(/[:T]/g, '-')}.csv`;
    a.click();
    URL.revokeObjectURL(a.href);
  }

  /** Xu\u1EA5t k\u1EBFt qu\u1EA3 ra file Excel (.xlsx) \u2014 d\u00F9ng th\u01B0 vi\u1EC7n XLSX (@require) */
  function exportXlsx() {
    if (!state.results.length) return alert('Ch\u01B0a c\u00F3 k\u1EBFt qu\u1EA3.');
    if (typeof XLSX === 'undefined') return alert('Th\u01B0 vi\u1EC7n Excel ch\u01B0a t\u1EA3i xong. \u0110\u1EE3i v\u00E0i gi\u00E2y r\u1ED3i th\u1EED l\u1EA1i (c\u1EA7n m\u1EA1ng \u0111\u1EC3 t\u1EA3i xlsx).');
    const aoa = [EXPORT_HEAD, ...exportRows()];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    // Gi\u00E3n c\u1ED9t cho d\u1EC5 \u0111\u1ECDc
    ws['!cols'] = EXPORT_HEAD.map((h, i) => ({ wch: i === 9 ? 40 : Math.max(10, h.length + 2) }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Ket qua');
    XLSX.writeFile(wb, `ghn_suco_${new Date().toISOString().slice(0, 19).replace(/[:T]/g, '-')}.xlsx`);
  }

  /* =========================================================
   * 9. GIAO DIỆN
   * ========================================================= */
  const ui = {};

  /* =========================================================
   * 9b. [v1.5.0] CHÈN MỤC "Hạ sự cố" VÀO MENU HỖ TRỢ CỦA TRANG
   *   Sao chép kiểu dáng mục có sẵn ("Xử lý sự cố"), chèn ngay sau nó.
   *   React render lại menu → MutationObserver tự chèn lại.
   * ========================================================= */
  const MENU_ITEM_ID = 'ghnb-menu-item';
  const MENU_LABEL = 'Hạ sự cố';
  const MENU_REF_LABEL = 'Xử lý sự cố'; // mục mẫu để sao kiểu dáng & vị trí

  function openPanel() {
    if (!ui.panel) return;
    ui.panel.style.display = 'block';
    renderAuthStatus();
  }

  // Tìm "hàng menu" trên cùng mà toàn bộ text đúng bằng label
  function findMenuRowByText(label) {
    let leaf = null;
    const all = document.querySelectorAll('a, span, div, li, p');
    for (const el of all) {
      if (el.children.length === 0 && el.textContent.trim() === label) { leaf = el; break; }
    }
    if (!leaf) return null;
    let row = leaf;
    while (row.parentElement && row.parentElement !== document.body &&
           row.parentElement.textContent.trim() === label) {
      row = row.parentElement;
    }
    return row;
  }

  function injectSidebarMenu() {
    if (document.getElementById(MENU_ITEM_ID)) return; // đã có
    const ref = findMenuRowByText(MENU_REF_LABEL);
    if (!ref || !ref.parentElement) return;

    const item = ref.cloneNode(true); // sao kiểu dáng của "Xử lý sự cố"
    item.id = MENU_ITEM_ID;

    // Đổi nhãn: thay text của phần tử lá sâu nhất (giữ nguyên icon/khung nếu có)
    const leaf = [...item.querySelectorAll('*')].filter(e => e.children.length === 0 && e.textContent.trim()).pop();
    if (leaf) leaf.textContent = MENU_LABEL;
    else item.textContent = MENU_LABEL;

    // Bỏ điều hướng của React Router (nếu là <a href>)
    if (item.tagName === 'A') item.removeAttribute('href');
    item.querySelectorAll && item.querySelectorAll('a[href]').forEach(a => a.removeAttribute('href'));
    item.style.cursor = 'pointer';

    // Chặn điều hướng mặc định của trang, mở panel của script
    item.addEventListener('click', (e) => { e.preventDefault(); e.stopPropagation(); openPanel(); }, true);

    ref.insertAdjacentElement('afterend', item);
  }

  // Chèn có debounce: check getElementById (rẻ) trước, chỉ quét DOM khi thật sự thiếu mục
  let menuInjectTimer = null;
  function scheduleInjectMenu() {
    if (document.getElementById(MENU_ITEM_ID)) return;
    if (menuInjectTimer) return;
    menuInjectTimer = setTimeout(() => { menuInjectTimer = null; injectSidebarMenu(); }, 300);
  }
  function watchSidebarMenu() {
    scheduleInjectMenu();
    new MutationObserver(scheduleInjectMenu).observe(document.body, { childList: true, subtree: true });
  }

  function buildUI() {
    const style = document.createElement('style');
    style.textContent = `
      #ghnb-btn{position:fixed;right:16px;bottom:16px;z-index:999999;background:#f26522;color:#fff;border:0;border-radius:24px;padding:10px 16px;font:600 13px sans-serif;cursor:pointer;box-shadow:0 4px 12px rgba(0,0,0,.25)}
      #ghnb{position:fixed;right:16px;bottom:64px;z-index:999999;width:520px;height:min(82vh,760px);min-width:380px;min-height:260px;overflow:auto;resize:both;background:#fff;color:#222;border-radius:10px;box-shadow:0 8px 30px rgba(0,0,0,.3);font:13px sans-serif;display:none}
      #ghnb.max{left:2vw!important;top:2vh!important;right:auto!important;bottom:auto!important;width:96vw!important;height:96vh!important;resize:none}
      #ghnb .hd{background:#f26522;color:#fff;padding:10px 12px;font-weight:600;display:flex;justify-content:space-between;align-items:center;position:sticky;top:0;z-index:1;cursor:move;user-select:none}
      #ghnb .hd .tools span{cursor:pointer;margin-left:12px;font-size:15px}
      #ghnb.dragging{opacity:.92;transition:none}
      #ghnb.max textarea#ghnb-orders{height:200px}
      #ghnb.max textarea#ghnb-export{height:180px}
      #ghnb .bd{padding:12px}
      #ghnb label{display:block;font-weight:600;margin:8px 0 3px}
      #ghnb input[type=text],#ghnb input[type=number],#ghnb textarea{width:100%;box-sizing:border-box;border:1px solid #ccc;border-radius:6px;padding:6px 8px;font:13px monospace}
      #ghnb textarea{height:120px;resize:vertical}
      #ghnb .row{display:flex;gap:8px}#ghnb .row>div{flex:1}
      #ghnb .btns{display:flex;gap:6px;margin-top:10px;flex-wrap:wrap}
      #ghnb button.b{border:0;border-radius:6px;padding:7px 12px;cursor:pointer;font-weight:600}
      #ghnb .run{background:#f26522;color:#fff}#ghnb .stop{background:#d32f2f;color:#fff}#ghnb .sec{background:#eee}
      #ghnb .auth{font-size:12px;padding:6px 8px;border-radius:6px;background:#f5f5f5}
      #ghnb .ok{color:#2e7d32}#ghnb .err{color:#d32f2f}
      #ghnb table{width:100%;border-collapse:collapse;margin-top:8px;font-size:12px}
      #ghnb th,#ghnb td{border-bottom:1px solid #eee;padding:4px;text-align:left;vertical-align:top}
      #ghnb tr.r{cursor:pointer}#ghnb tr.r:hover{background:#fff4ec}
    `;
    document.head.appendChild(style);

    // [v1.6.2] Bỏ nút nổi — chỉ mở panel qua mục "Hạ sự cố" trong menu Hỗ trợ của trang.

    const panel = document.createElement('div');
    panel.id = 'ghnb';
    panel.innerHTML = `
      <div class="hd" title="Kéo để di chuyển, nhấp đúp để phóng to"><span>GHN – Cập nhật DELIVER_FAIL hàng loạt</span><span class="tools"><span id="ghnb-max" title="Phóng to / thu nhỏ">⤢</span><span id="ghnb-x" title="Đóng">✕</span></span></div>
      <div class="bd">
        <div class="auth" id="ghnb-auth"></div>
        <div class="row">
          <div><label>Mã kho (x-warehouseid)</label><input type="text" id="ghnb-wh"></div>
          <div><label>Mã chuyến (tuỳ chọn)</label><input type="text" id="ghnb-trip" placeholder="Bỏ trống = tự tìm"></div>
        </div>
        <label>Danh sách mã đơn (mỗi dòng 1 mã, hoặc cách nhau bởi dấu phẩy)</label>
        <textarea id="ghnb-orders" placeholder="VNGH80435700411&#10;VNGH..."></textarea>
        <label style="font-weight:600;margin-top:4px">Chế độ cập nhật:
          <label style="display:inline;font-weight:600;margin-left:8px"><input type="radio" name="ghnb-mode" value="fail" checked> 🔴 Giao thất bại</label>
          <label style="display:inline;font-weight:600;margin-left:14px"><input type="radio" name="ghnb-mode" value="success"> 🟢 Giao thành công</label>
        </label>
        <div class="row">
          <div id="ghnb-fc-wrap"><label>failCode</label><input type="text" id="ghnb-fc" value="${CONFIG.DEFAULT_FAIL_CODE}"></div>
          <div><label>Delay (ms)</label><input type="number" id="ghnb-delay" value="${CONFIG.DEFAULT_DELAY_MS}"></div>
        </div>
        <div id="ghnb-fn-wrap"><label>failNote</label><input type="text" id="ghnb-fn" value="${CONFIG.DEFAULT_FAIL_NOTE}"></div>
        <div id="ghnb-pod-wrap" style="display:none"><label>Ảnh POD (upload 1 lần, dùng chung cho tất cả đơn trong lô)</label><input type="file" id="ghnb-pod" accept="image/*"></div>
        <label style="font-weight:400"><input type="checkbox" id="ghnb-autohub" checked> <b>Tự đổi kho</b> theo kho của đơn (đơn khác kho vẫn xử lý)</label>
        <label style="font-weight:400;margin-top:2px"><input type="checkbox" id="ghnb-dry" checked> <b>Chỉ kiểm tra</b> (KHÔNG lưu). Bỏ tick để <b>lưu trạng thái thật</b></label>
        <div class="btns">
          <button class="b run" id="ghnb-run">▶ Bắt đầu</button>
          <button class="b stop" id="ghnb-stop" disabled>■ Dừng</button>
          <button class="b sec" id="ghnb-csv">⬇ Xuất CSV</button>
          <button class="b sec" id="ghnb-xlsx" style="background:#16A34A;color:#fff">📊 Xuất Excel</button>
          <button class="b sec" id="ghnb-retry">↻ Lọc đơn lỗi</button>
          <button class="b sec" id="ghnb-debug">📋 Copy debug</button>
        </div>
        <div id="ghnb-prog" style="margin-top:8px;font-weight:600"></div>
        <div id="ghnb-stats" style="margin-top:6px;font-size:12.5px;color:#334155;padding:6px 8px;background:#F8FAFC;border-radius:6px;display:none"></div>
        <div style="font-size:11px;color:#888;margin-top:4px">Bấm vào 1 dòng để xem payload/response trong Console (F12).</div>
        <label>Dữ liệu xuất (tự điền sau khi chạy – dán thẳng vào Google Sheets)</label>
        <textarea id="ghnb-export" readonly style="height:90px;font-size:11px;white-space:pre;overflow:auto"></textarea>
        <div id="ghnb-sum" style="font-size:12px;color:#555;margin-top:4px"></div>
        <div class="btns" style="margin-top:4px"><button class="b sec" id="ghnb-copyexp">📄 Copy dữ liệu</button></div>
        <table><thead><tr><th>#</th><th>Kho</th><th>Mã đơn</th><th>Chuyến</th><th>Bước</th><th>Tích TC</th><th>Kết quả</th></tr></thead><tbody id="ghnb-log"></tbody></table>
      </div>`;
    document.body.appendChild(panel);

    Object.assign(ui, {
      panel,
      auth: panel.querySelector('#ghnb-auth'),
      wh: panel.querySelector('#ghnb-wh'),
      trip: panel.querySelector('#ghnb-trip'),
      orders: panel.querySelector('#ghnb-orders'),
      failCode: panel.querySelector('#ghnb-fc'),
      failNote: panel.querySelector('#ghnb-fn'),
      pod: panel.querySelector('#ghnb-pod'),
      podWrap: panel.querySelector('#ghnb-pod-wrap'),
      fcWrap: panel.querySelector('#ghnb-fc-wrap'),
      fnWrap: panel.querySelector('#ghnb-fn-wrap'),
      delay: panel.querySelector('#ghnb-delay'),
      dryRun: panel.querySelector('#ghnb-dry'),
      autoHub: panel.querySelector('#ghnb-autohub'),
      exportBox: panel.querySelector('#ghnb-export'),
      exportSum: panel.querySelector('#ghnb-sum'),
      run: panel.querySelector('#ghnb-run'),
      stop: panel.querySelector('#ghnb-stop'),
      progress: panel.querySelector('#ghnb-prog'),
      stats: panel.querySelector('#ghnb-stats'),
      log: panel.querySelector('#ghnb-log'),
    });

    enableDragResize(panel);

    panel.querySelector('#ghnb-x').onclick = () => (panel.style.display = 'none');
    ui.run.onclick = () => { if (!state.running) runBatch(); };
    // Chế độ fail/success: đọc radio đang chọn
    ui.mode = { get value() { const r = panel.querySelector('input[name="ghnb-mode"]:checked'); return r ? r.value : 'fail'; } };

    // Ẩn/hiện field theo chế độ + đổi chữ/màu nút chạy (tránh nhầm kiểm tra ↔ lưu thật)
    const syncMode = () => {
      const isSuccess = ui.mode.value === 'success';
      if (ui.fcWrap) ui.fcWrap.style.display = isSuccess ? 'none' : '';
      if (ui.fnWrap) ui.fnWrap.style.display = isSuccess ? 'none' : '';
      if (ui.podWrap) ui.podWrap.style.display = isSuccess ? '' : 'none';
      ui.run.textContent = ui.dryRun.checked ? '🔍 Kiểm tra' : (isSuccess ? '💾 Lưu Giao thành công' : '💾 Lưu Giao thất bại');
      ui.run.style.background = ui.dryRun.checked ? '#1976d2' : (isSuccess ? '#0f9d58' : '#f26522');
    };
    ui.dryRun.onchange = syncMode;
    panel.querySelectorAll('input[name="ghnb-mode"]').forEach((r) => (r.onchange = syncMode));
    syncMode();
    ui.stop.onclick = () => { state.stopFlag = true; ui.progress.textContent = 'Đang dừng sau đơn hiện tại...'; };
    panel.querySelector('#ghnb-csv').onclick = exportCsv;
    panel.querySelector('#ghnb-xlsx').onclick = exportXlsx;
    // Đưa các đơn lỗi về lại ô nhập để chạy lại
    panel.querySelector('#ghnb-retry').onclick = () => {
      const failed = state.results.filter((r) => !r.ok).map((r) => r.orderCode);
      if (!failed.length) return alert('Không có đơn lỗi.');
      ui.orders.value = failed.join('\n');
    };
    // Copy chuỗi dữ liệu xuất (TSV)
    panel.querySelector('#ghnb-copyexp').onclick = async () => {
      if (!ui.exportBox.value) return alert('Chưa có dữ liệu. Hãy chạy Kiểm tra trước.');
      try { await navigator.clipboard.writeText(ui.exportBox.value); alert('Đã copy. Dán vào Google Sheets/Excel.'); }
      catch (_) { ui.exportBox.select(); document.execCommand('copy'); alert('Đã copy.'); }
    };
    // Copy toàn bộ kết quả (request + response từng đơn) để gửi hỗ trợ; ẩn token
    panel.querySelector('#ghnb-debug').onclick = async () => {
      if (!state.results.length) return alert('Chưa có kết quả.');
      const dump = JSON.stringify({
        version: '1.6.0',
        warehouseId: state.warehouseId,
        mode: ui.dryRun.checked ? 'CHI_KIEM_TRA' : 'LUU_THAT',
        progress: ui.progress.textContent,
        results: state.results.map(({ entry, ...r }) => r),
      }, null, 2);
      try { await navigator.clipboard.writeText(dump); alert('Đã copy debug. Dán vào chat để gửi.'); }
      catch (_) { console.log(dump); alert('Không copy được – đã in ra Console (F12).'); }
    };
    // Bấm dòng log → in chi tiết ra Console để debug
    ui.log.onclick = (e) => {
      const tr = e.target.closest('tr.r');
      if (!tr) return;
      const r = state.results[Number(tr.dataset.i)];
      console.group(`[GHN Bulk] ${r.orderCode}`);
      console.log('Payload (đơn này):', r.payload);
      console.log('Response:', r.raw);
      console.groupEnd();
    };

    renderAuthStatus();
    watchSidebarMenu(); // chèn mục "Hạ sự cố" vào menu Hỗ trợ của trang
  }

  /* =========================================================
   * 10. KÉO THẢ / ĐỔI KÍCH THƯỚC / PHÓNG TO PANEL
   * ========================================================= */
  function enableDragResize(panel) {
    const hd = panel.querySelector('.hd');
    const clamp = (v, min, max) => Math.min(Math.max(v, min), max);

    // Khôi phục vị trí/kích thước đã lưu
    try {
      const g = JSON.parse(localStorage.getItem(CONFIG.GEOMETRY_KEY) || 'null');
      if (g) {
        if (g.w) panel.style.width = g.w + 'px';
        if (g.h) panel.style.height = g.h + 'px';
        if (g.left != null && g.top != null) {
          panel.style.left = clamp(g.left, 0, window.innerWidth - 120) + 'px';
          panel.style.top = clamp(g.top, 0, window.innerHeight - 60) + 'px';
          panel.style.right = 'auto';
          panel.style.bottom = 'auto';
        }
        if (g.max) panel.classList.add('max');
      }
    } catch (_) { /* bỏ qua */ }

    const save = () => {
      try {
        const r = panel.getBoundingClientRect();
        const isMax = panel.classList.contains('max');
        const prev = JSON.parse(localStorage.getItem(CONFIG.GEOMETRY_KEY) || '{}');
        // Khi đang phóng to: chỉ lưu cờ max, giữ nguyên kích thước thường
        const g = isMax ? { ...prev, max: true }
          : { left: Math.round(r.left), top: Math.round(r.top), w: Math.round(r.width), h: Math.round(r.height), max: false };
        localStorage.setItem(CONFIG.GEOMETRY_KEY, JSON.stringify(g));
      } catch (_) { /* bỏ qua */ }
    };

    // Kéo thả bằng thanh tiêu đề (chuột + cảm ứng)
    let drag = null;
    hd.addEventListener('pointerdown', (e) => {
      if (e.target.closest('.tools') || panel.classList.contains('max')) return;
      const r = panel.getBoundingClientRect();
      drag = { dx: e.clientX - r.left, dy: e.clientY - r.top };
      panel.classList.add('dragging');
      hd.setPointerCapture(e.pointerId);
      e.preventDefault();
    });
    hd.addEventListener('pointermove', (e) => {
      if (!drag) return;
      const w = panel.offsetWidth;
      panel.style.left = clamp(e.clientX - drag.dx, 0, window.innerWidth - Math.min(w, 120)) + 'px';
      panel.style.top = clamp(e.clientY - drag.dy, 0, window.innerHeight - 40) + 'px';
      panel.style.right = 'auto';
      panel.style.bottom = 'auto';
    });
    const endDrag = () => { if (!drag) return; drag = null; panel.classList.remove('dragging'); save(); };
    hd.addEventListener('pointerup', endDrag);
    hd.addEventListener('pointercancel', endDrag);

    // Phóng to / thu nhỏ: nút ⤢ hoặc nhấp đúp tiêu đề
    const toggleMax = () => { panel.classList.toggle('max'); save(); };
    panel.querySelector('#ghnb-max').onclick = toggleMax;
    hd.addEventListener('dblclick', (e) => { if (!e.target.closest('.tools')) toggleMax(); });

    // Ghi nhớ kích thước khi kéo góc phải-dưới (resize)
    let t;
    new ResizeObserver(() => {
      if (panel.style.display !== 'block' || panel.classList.contains('max')) return;
      clearTimeout(t); t = setTimeout(save, 300);
    }).observe(panel);
  }

  function renderAuthStatus() {
    if (!ui.auth) return;
    const t = state.token;
    ui.auth.innerHTML =
      `Token: ${t ? `<span class="ok">✔ đã bắt (…${esc(t.slice(-8))})</span>` : '<span class="err">✘ chưa có – thao tác 1 lần trên trang</span>'}` +
      ` &nbsp;|&nbsp; Kho: ${state.warehouseId ? `<span class="ok">${esc(state.warehouseId)}</span>` : '<span class="err">chưa có</span>'}`;
    if (ui.wh && document.activeElement !== ui.wh) ui.wh.value = state.warehouseId;
  }

  /** Phân loại trạng thái 1 dòng để tô màu: ready(vàng) / success(xanh) / fail(đỏ) / pending */
  function rowCategory(r) {
    if (r.ok === null) return 'pending';            // đang chờ pha 2
    if (r.step === 'Đã kiểm tra') return 'ready';   // chế độ Kiểm tra → Sẵn sàng (vàng)
    if (r.markedBy === 'Nhân viên' || r.markedBy === 'AM tích') return 'success';
    if (r.ok && r.entry && r.entry.status === 'DELIVER_SUCCESS') return 'success';
    return 'fail';                                  // thất bại / lỗi / không cập nhật được (đỏ)
  }

  const CAT_STYLE = { ready: 'color:#D97706', success: 'color:#16A34A', fail: 'color:#DC2626', pending: 'color:#64748B' };
  const CAT_ICON = { ready: '🟡', success: '🟢', fail: '🔴', pending: '⏳' };

  function renderLog() {
    ui.log.innerHTML = state.results.map((r, i) => {
      const cat = rowCategory(r);
      // Thành công: chỉ ghi "Giao Thành Công" (ai tích đã có ở cột Tích TC)
      const text = cat === 'success' ? 'Giao Thành Công' : r.message;
      return `
      <tr class="r" data-i="${i}">
        <td>${i + 1}</td><td>${esc(r.hubId)}</td><td>${esc(r.orderCode)}</td><td>${esc(r.tripCode)}</td><td>${esc(r.step)}</td>
        <td>${r.markedBy === 'AM tích' ? '<b style="color:#16A34A">AM tích</b>' : (r.markedBy ? esc(r.markedBy) : '')}</td>
        <td style="${CAT_STYLE[cat]};font-weight:600">${CAT_ICON[cat]} ${esc(text)}</td>
      </tr>`;
    }).join('');
    renderStats();
  }

  /** Dòng thống kê nhanh số đơn theo trạng thái */
  function renderStats() {
    if (!ui.stats) return;
    const total = state.results.length;
    if (!total) { ui.stats.style.display = 'none'; ui.stats.innerHTML = ''; return; }
    const c = { ready: 0, success: 0, fail: 0, pending: 0 };
    let amTick = 0, nvTick = 0;
    state.results.forEach((r) => {
      c[rowCategory(r)]++;
      if (r.markedBy === 'AM tích') amTick++;
      else if (r.markedBy === 'Nhân viên') nvTick++;
    });
    ui.stats.style.display = 'block';
    ui.stats.innerHTML =
      `📊 Tổng <b>${total}</b>` +
      ` &nbsp;|&nbsp; <span style="color:#16A34A">🟢 Thành công <b>${c.success}</b></span>` +
      (c.success ? ` <span style="color:#64748B">(AM tích ${amTick}, NV ${nvTick})</span>` : '') +
      ` &nbsp;|&nbsp; <span style="color:#DC2626">🔴 Thất bại/lỗi <b>${c.fail}</b></span>` +
      ` &nbsp;|&nbsp; <span style="color:#D97706">🟡 Sẵn sàng <b>${c.ready}</b></span>` +
      (c.pending ? ` &nbsp;|&nbsp; ⏳ Chờ <b>${c.pending}</b>` : '');
  }

  function setRunning(on) {
    ui.run.disabled = on;
    ui.stop.disabled = !on;
  }

  function setProgress(i, n) {
    ui.progress.textContent = `Đang xử lý ${i}/${n}...`;
  }

  // Dựng UI khi DOM sẵn sàng (script chạy document-start để kịp hook request)
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', buildUI);
  } else {
    buildUI();
  }
})();