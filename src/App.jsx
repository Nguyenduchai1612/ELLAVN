import React, { createContext, useContext, useReducer, useMemo, useState, useRef, useCallback, useEffect } from "react";
import * as XLSX from "xlsx";
import {
  ComposedChart, Bar, Line, PieChart, Pie, Cell, XAxis, YAxis, CartesianGrid,
  Tooltip, Legend, ResponsiveContainer,
} from "recharts";
import {
  Upload, TrendingUp, TrendingDown, DollarSign, Package, Percent, ShoppingBag,
  Settings, FileSpreadsheet, Download, AlertCircle, CheckCircle2, X, Info,
  LayoutDashboard, Store, Trash2, Lock, Unlock, Users, Calendar, Megaphone,
  RefreshCw, ChevronDown, Search, ClipboardList, ShieldCheck,
  ChevronLeft, ChevronRight, Plus, AlertTriangle, Target, Wallet,
} from "lucide-react";

/* ============================================================================
 * 1. DATA MODEL (documented as JSDoc "types")
 * ==========================================================================
 *
 * @typedef {"tiktok"|"shopee"} Platform
 * @typedef {"success"|"cancelled"|"returned"} OrderStatus
 *
 * @typedef SkuConfigEntry
 * @property {string} sku            Seller SKU (mã SKU người bán)
 * @property {"kinh"|"hop"|"khan"|"khac"} type
 * @property {string} name           Tên sản phẩm hiển thị
 * @property {number} cogs           Giá vốn (VNĐ)
 *
 * @typedef OrderLineItem
 * @property {string} sellerSku
 * @property {number} subtotal
 * @property {number} cogs
 * @property {"kinh"|"hop"|"khan"|"khac"} type
 *
 * @typedef Order
 * @property {string} id             Order ID
 * @property {Platform} platform
 * @property {string} date           ISO date (yyyy-mm-dd)
 * @property {OrderStatus} status
 * @property {OrderLineItem[]} items
 * @property {number} gmv            = tổng subtotal các items (giá trị khởi tạo)
 * @property {number} settlementAmount  Số tiền quyết toán thực tế (nếu có)
 * @property {Object} fees           Chi tiết các dòng phí sàn
 * @property {number} feesTotal
 * @property {number} cogsTotal
 * @property {string} [returnReason]
 * // computed at read-time:
 * @property {number} nettRevenue
 * @property {number} packagingFee
 * @property {number} profitBeforeAds
 *
 * @typedef Settings
 * @property {number} packagingFee        Mặc định 6000đ / đơn
 * @property {number} monthlyFixedCost    Mặc định 0
 * @property {number} otherVariableCost   Mặc định 0
 *
 * @typedef AdsSpend
 * @property {number} tiktok
 * @property {number} shopee
 */

/* ============================================================================
 * 2. CONSTANTS & FORMATTERS
 * ========================================================================== */

const PLATFORM = { TIKTOK: "tiktok", SHOPEE: "shopee" };
const STATUS = { SUCCESS: "success", CANCELLED: "cancelled", RETURNED: "returned" };
const DEFAULT_PACKAGING_FEE = 6000;

const COLORS = {
  navy: "#0c2d55",
  navyLight: "#14477f",
  pink: "#d4537e",
  pinkDark: "#993556",
  gray: "#8a8f98",
  grayLight: "#eef1f5",
  green: "#1d9e75",
  red: "#d03b3b",
  amber: "#c98500",
};

function formatVND(n) {
  const v = Math.round(Number(n) || 0);
  return v.toLocaleString("vi-VN") + "đ";
}
function formatNumber(n) {
  return Math.round(Number(n) || 0).toLocaleString("vi-VN");
}
function formatPercent(n) {
  const v = Math.round((Number(n) || 0) * 10) / 10;
  return v.toFixed(1) + "%";
}
function uid(prefix = "id") {
  return prefix + "_" + Math.random().toString(36).slice(2, 10);
}
function todayIso() {
  return new Date().toISOString().slice(0, 10);
}
function platformLabel(p) {
  return p === PLATFORM.TIKTOK ? "TikTok Shop" : "Shopee";
}
function statusLabel(s) {
  if (s === STATUS.SUCCESS) return "Thành công";
  if (s === STATUS.CANCELLED) return "Đã hủy";
  return "Hoàn trả";
}

/* ============================================================================
 * 3. FILE PARSING ENGINE (XLSX / CSV via SheetJS)
 * Real TikTok Shop / Shopee export files vary between accounts, ngôn ngữ và
 * phiên bản. Bộ máy dưới đây dùng khớp cột theo "alias" (không phân biệt hoa
 * thường / dấu), cố gắng khớp tốt nhất và báo lỗi rõ ràng nếu thiếu cột bắt
 * buộc, thay vì crash toàn bộ import.
 * ========================================================================== */

const COLUMN_ALIASES = {
  orderId: ["order id", "mã đơn hàng", "order_id", "id đơn hàng", "ma don hang"],
  // File Đối soát/Income của TikTok gộp chung nhiều LOẠI dòng trong cùng 1
  // sheet (đơn hàng thật, "GMV thanh toán cho Quảng cáo TikTok", điều chỉnh...).
  // Cột này giúp lọc ra CHỈ các dòng "Đơn hàng" thật, tránh đếm nhầm dòng
  // billing quảng cáo thành đơn hàng (làm sai lệch GMV/số đơn/doanh thu).
  transactionType: ["transaction type", "loại giao dịch", "loai giao dich"],
  sellerSku: ["seller sku", "sku người bán", "sku nguoi ban", "sku"],
  skuId: ["sku id", "id sku"],
  productName: ["tên sản phẩm", "ten san pham", "product name"],
  // Ưu tiên "before discount" (giá trị khởi tạo) cho GMV; "after discount" và các biến thể
  // khác được giữ làm phương án dự phòng nếu file không có cột "before discount".
  subtotal: ["sku subtotal before discount", "subtotal before discount", "subtotal", "thành tiền", "thanh tien", "giá trị đơn hàng", "gia tri don hang", "sku subtotal after discount", "giá bán", "gia ban"],
  // "Phí vận chuyển của người bán" là tên cột THẬT trong file Đối soát/Income
  // của TikTok Seller Center (phí logistics trừ vào người bán). Trước đây chỉ có
  // alias "phí vận chuyển" (2 từ) trong khi cột thật dài hơn — vì trường này ở
  // chế độ exact-match (EXACT_ONLY_FIELDS) nên KHÔNG BAO GIỜ khớp được, khiến
  // phí vận chuyển luôn bị tính = 0đ dù file có dữ liệu thật.
  shippingFee: ["shipping fee", "phí vận chuyển", "phi van chuyen", "phí vận chuyển của người bán", "phi van chuyen cua nguoi ban"],
  transactionFee: ["transaction fee", "phí giao dịch", "phi giao dich"],
  commissionFee: ["tiktok shop commission fee", "commission fee", "phí hoa hồng", "phi hoa hong", "phí cố định", "phi co dinh"],
  affiliateCommission: ["affiliate commission", "hoa hồng tiếp thị liên kết", "hoa hong affiliate", "phí affiliate", "phi affiliate"],
  voucherXtra: ["voucher xtra", "phí voucher xtra", "phi voucher xtra"],
  orderProcessingFee: ["order processing fee", "phí xử lý đơn hàng", "phi xu ly don hang"],
  taxWithheld: ["vat/pit withheld", "taxes", "thuế khấu trừ", "thue khau tru", "vat", "pit", "thuế", "thue"],
  serviceFee: ["service fee", "phí dịch vụ", "phi dich vu"],
  paymentFee: ["payment fee", "phí thanh toán", "phi thanh toan"],
  freeshipXtra: ["freeship xtra", "hoàn xu xtra", "hoan xu xtra", "phí freeship xtra", "phi freeship xtra"],
  // "Order Amount" (tổng tiền buyer đã thanh toán) dùng làm giá trị quyết toán dự phòng
  // khi chưa có file Đối soát/Statement of Account chi tiết (chưa trừ hoa hồng sàn).
  settlementAmount: ["net settlement amount", "số tiền quyết toán", "so tien quyet toan", "total settlement amount", "escrow amount", "total release amount", "số tiền thực nhận", "so tien thuc nhan", "order amount"],
  orderStatus: ["order status", "trạng thái đơn hàng", "trang thai don hang", "trạng thái", "trang thai"],
  cancelType: ["cancelation/return type", "cancellation/return type", "cancel/return type", "cancel type"],
  orderDate: ["created time", "order date", "ngày tạo đơn", "ngay tao don", "thời gian tạo đơn", "thoi gian tao don", "ngày đặt hàng", "ngay dat hang"],
  returnStatus: ["return status", "trạng thái hoàn tiền", "trang thai hoan tien", "trạng thái trả hàng", "trang thai tra hang"],
  returnReason: ["return reason", "lý do hoàn", "ly do hoan", "lý do trả hàng", "ly do tra hang"],
  skuType: ["loại sku", "loai sku", "loại", "loai", "type", "phân loại", "phan loai"],
  cogs: ["giá vốn", "gia von", "cogs", "cost of goods", "gia von don vi"],
  // "Giá bán lẻ (Nội tệ)"/"price" trong file Batch Edit Template — giá NIÊM YẾT
  // trên gian hàng, KHÁC với "cogs" (giá vốn nhập hàng). Lưu lại để dùng cho
  // tính năng gợi ý giá bán sau này, không hiển thị thêm cột trong bảng SKU
  // (bảng SKU chỉ giữ đúng 4 cột: SKU / Tên SP / Tồn kho / Giá vốn).
  // "price" là header tiếng Anh THẬT của cột giá bán trong file Batch Edit
  // Template (Tiktoksellercenter_batchedit_...) — trước đây chỉ có các alias
  // dài ("selling price"...) nên không khớp được với header ngắn "price",
  // khiến giá bán luôn bị đọc = 0đ dù file có dữ liệu thật.
  sellingPrice: ["giá bán lẻ", "gia ban le", "retail price", "selling price", "listing price", "price", "giá bán"],
  // Cột "quantity"/"Số lượng" trong file Batch Edit Template (Tiktoksellercenter_batchedit...)
  // của Seller Center — dùng để cảnh báo tồn kho thấp ở trang Sản phẩm.
  stockQty: ["quantity", "số lượng", "so luong", "available stock", "tồn kho", "ton kho", "lượng tồn kho", "luong ton kho"],
  adsCost: ["chi phí", "chi phi", "cost", "spend", "ad spend", "amount spent"],
};

function normalizeHeader(h) {
  return String(h ?? "")
    .trim()
    .toLowerCase()
    .replace(/_/g, " ")
    .replace(/\s+/g, " ")
    .normalize("NFC");
}

/** Các trường CHỈ được khớp khi trùng khớp CHÍNH XÁC — không dùng "chứa cụm
 * từ" — vì nhiều file đơn hàng có các cột phái sinh dễ gây hiểu nhầm, ví dụ
 * "Shipping Fee After Discount" (phí ship buyer trả) chứa sẵn cụm "shipping
 * fee" nhưng KHÔNG phải là phí sàn trừ vào người bán như cột "Shipping Fee"
 * trong file Đối soát thật. Khớp mờ sẽ khiến chi phí bị tính sai/thổi phồng. */
const EXACT_ONLY_FIELDS = new Set(["shippingFee", "transactionType"]);

/**
 * Khớp cột theo 2 vòng: (1) khớp CHÍNH XÁC trước — tránh trường hợp một cột
 * như "Return Order ID" bị nhầm là "Order ID" chỉ vì chứa cụm từ đó;
 * (2) khớp theo "chứa cụm từ" cho các trường còn thiếu (trừ EXACT_ONLY_FIELDS),
 * dùng như phương án dự phòng.
 */
function buildHeaderMap(headers) {
  const map = {};
  const norms = headers.map(normalizeHeader);
  norms.forEach((norm, idx) => {
    if (!norm) return;
    for (const [field, aliases] of Object.entries(COLUMN_ALIASES)) {
      if (field in map) continue;
      if (aliases.includes(norm)) map[field] = idx;
    }
  });
  norms.forEach((norm, idx) => {
    if (!norm) return;
    for (const [field, aliases] of Object.entries(COLUMN_ALIASES)) {
      if (field in map || EXACT_ONLY_FIELDS.has(field)) continue;
      if (aliases.some((a) => norm.includes(a))) map[field] = idx;
    }
  });
  return map;
}

/** VNĐ trong các file xuất từ Seller Center không có phần thập phân, nhưng có
 * thể ở dạng "44.650₫" (dấu chấm phân cách hàng nghìn + ký hiệu tiền tệ).
 * Cách an toàn nhất là loại bỏ mọi ký tự không phải chữ số (giữ lại dấu trừ). */
function toNumber(v) {
  if (v === undefined || v === null || v === "") return 0;
  if (typeof v === "number") return v;
  const s = String(v).trim();
  if (!s) return 0;
  const isNegative = /^-/.test(s) || /^\(.*\)$/.test(s);
  const digits = s.replace(/[^\d]/g, "");
  if (!digits) return 0;
  const n = parseInt(digits, 10);
  return isNegative ? -n : n;
}

/** Nhiều file xuất từ Seller Center có 1-2 dòng "mô tả cột" ngay dưới header
 * (vd: "Platform unique order ID.") chứ không phải dữ liệu thật. Các dòng này
 * luôn có ID không phải là chuỗi số thuần, nên lọc bằng cách kiểm tra ID. */
function looksLikeNumericId(v) {
  const s = String(v ?? "").trim();
  return /^\d+$/.test(s);
}

function toIsoDate(v) {
  if (!v) return todayIso();
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  const s = String(v).trim();
  // try dd/mm/yyyy
  const m1 = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
  if (m1) {
    const [, d, mo, y] = m1;
    const yyyy = y.length === 2 ? "20" + y : y;
    return `${yyyy}-${mo.padStart(2, "0")}-${d.padStart(2, "0")}`;
  }
  const m2 = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m2) {
    const [, y, mo, d] = m2;
    return `${y}-${mo.padStart(2, "0")}-${d.padStart(2, "0")}`;
  }
  const d = new Date(s);
  if (!isNaN(d.getTime())) return d.toISOString().slice(0, 10);
  return todayIso();
}

/**
 * Xác định trạng thái đơn dựa trên Order Status + cột Cancelation/Return Type
 * (nếu có). Lưu ý: substatus "Đã hoàn tất" nghĩa là ĐÃ HOÀN THÀNH đơn hàng
 * (completed), KHÔNG phải "hoàn tiền" — nên không được dùng từ khóa "hoàn"
 * đơn thuần để suy ra RETURNED, dễ nhầm lẫn.
 */
function guessOrderStatus(rawStatus, rawCancelType) {
  const s = normalizeHeader(rawStatus);
  const c = normalizeHeader(rawCancelType);
  // Lưu ý: KHÔNG dò theo chuỗi con "huy" không dấu — từ "chuyển" (vận CHUYển)
  // chứa sẵn "huy" như một chuỗi con và sẽ khiến toàn bộ đơn "Đã vận chuyển"
  // bị nhận nhầm thành "Đã hủy". Chỉ dò "hủy" có dấu (đặc trưng cho hủy đơn).
  if (c.includes("cancel") || s.includes("hủy")) return STATUS.CANCELLED;
  if (c.includes("return") || c.includes("refund") || c.includes("hoàn tiền") || c.includes("hoàn hàng")) return STATUS.RETURNED;
  return STATUS.SUCCESS;
}

/** Một số file xuất từ TikTok Seller Center có "!ref" (phạm vi vùng dữ liệu
 * trong XML) khai báo SAI — thiếu mất các dòng dữ liệu thật ở cuối sheet
 * (quan sát thấy ở các file "Batch Edit Template"). Phải quét lại toàn bộ ô
 * thực tế có trong sheet để tính đúng phạm vi trước khi chuyển sang JSON,
 * nếu không sẽ mất dữ liệu một cách âm thầm (không báo lỗi). */
function fixSheetRange(sheet) {
  let maxRow = -1, maxCol = -1;
  for (const key in sheet) {
    if (key[0] === "!") continue;
    const m = key.match(/^([A-Z]+)(\d+)$/);
    if (!m) continue;
    const col = XLSX.utils.decode_col(m[1]);
    const row = parseInt(m[2], 10) - 1;
    if (row > maxRow) maxRow = row;
    if (col > maxCol) maxCol = col;
  }
  if (maxRow < 0) return;
  const declared = sheet["!ref"] ? XLSX.utils.decode_range(sheet["!ref"]) : null;
  const eRow = declared ? Math.max(declared.e.r, maxRow) : maxRow;
  const eCol = declared ? Math.max(declared.e.c, maxCol) : maxCol;
  if (!declared || declared.e.r < maxRow || declared.e.c < maxCol) {
    sheet["!ref"] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: eRow, c: eCol } });
  }
}

/** Nhiều file xuất từ Seller Center (đặc biệt file Đối soát/Statement of
 * Account/Income) có NHIỀU sheet trong cùng 1 workbook — vd: "Chi tiết đơn
 * hàng", "Báo cáo", "Lịch sử rút tiền", "Giải thích về phí". Chỉ sheet đầu
 * tiên (thường là "Chi tiết đơn hàng"/"Order Detail") chứa dữ liệu từng dòng
 * cần để tính phí sàn; các sheet còn lại là tổng hợp/tài liệu tham khảo và
 * không có cấu trúc dòng-đơn-hàng nên KHÔNG được đọc làm dữ liệu đơn hàng.
 * Trước đây hệ thống luôn lấy `wb.SheetNames[0]` theo VỊ TRÍ — nếu TikTok đổi
 * thứ tự sheet giữa các lần xuất, hệ thống sẽ âm thầm đọc nhầm sheet tổng hợp
 * (vd: "Báo cáo") và fail import mà không rõ lý do. Hàm dưới đây tìm sheet
 * theo TÊN trước, chỉ fallback về sheet đầu tiên nếu không tên nào khớp. */
const DETAIL_SHEET_NAME_HINTS = [
  "chi tiết đơn hàng", "chi tiet don hang", "order detail", "order details", "chi tiết đơn hàng/điều chỉnh",
];
const NON_DETAIL_SHEET_NAME_HINTS = ["báo cáo", "bao cao", "report", "lịch sử rút tiền", "lich su rut tien", "giải thích", "giai thich"];

function pickDetailSheetName(sheetNames) {
  const norms = sheetNames.map(normalizeHeader);
  const hintIdx = norms.findIndex((n) => DETAIL_SHEET_NAME_HINTS.some((h) => n.includes(h)));
  if (hintIdx !== -1) return sheetNames[hintIdx];
  const nonDetailIdx = norms.findIndex((n) => NON_DETAIL_SHEET_NAME_HINTS.some((h) => n.includes(h)));
  const fallbackIdx = nonDetailIdx === 0 && sheetNames.length > 1 ? 1 : 0;
  return sheetNames[fallbackIdx];
}

/** Đọc 1 file xlsx/csv ở client-side, trả về { headers, rows } */
async function readSpreadsheet(file) {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array", cellDates: true });
  const sheetName = pickDetailSheetName(wb.SheetNames);
  const sheet = wb.Sheets[sheetName];
  fixSheetRange(sheet);
  const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: "" });
  if (!raw.length) return { headers: [], rows: [] };
  const headers = raw[0];
  const rows = raw.slice(1).filter((r) => r.some((c) => String(c).trim() !== ""));
  return { headers, rows };
}

/**
 * Parse file Đơn hàng & Quyết toán (TikTok hoặc Shopee).
 * Giả định: mỗi dòng = 1 line item (1 SKU) của 1 đơn hàng. Các cột phí ở
 * cấp đơn hàng có thể lặp lại trên mọi dòng của cùng Order ID, hoặc chỉ có ở
 * dòng đầu — engine gộp theo Order ID và lấy giá trị phí lớn nhất tìm thấy
 * cho mỗi loại phí (tránh cộng dồn phí bị lặp), còn subtotal thì CỘNG DỒN
 * qua các dòng (vì mỗi dòng là 1 sản phẩm trong đơn 3-món).
 */
function parseSettlementRows(headers, rows, platform, skuConfigMap) {
  const map = buildHeaderMap(headers);
  const warnings = [];
  if (map.orderId === undefined) {
    warnings.push("Không tìm thấy cột 'Order ID' — không thể import file này.");
    return { orders: [], warnings, ok: false };
  }

  const grouped = new Map();
  const missingSkuSet = new Set();

  rows.forEach((row) => {
    const orderIdRaw = row[map.orderId];
    const orderId = String(orderIdRaw ?? "").trim();
    if (!orderId) return;
    // Bỏ qua dòng "mô tả cột" (vd: "Platform unique order ID.") nằm ngay dưới
    // header trong nhiều file xuất từ TikTok Seller Center: Order ID thật của
    // TikTok luôn là một chuỗi số thuần.
    if (platform === PLATFORM.TIKTOK && !looksLikeNumericId(orderId)) return;
    // File Đối soát/Income có thể chứa các dòng KHÔNG PHẢI đơn hàng thật (vd:
    // "GMV thanh toán cho Quảng cáo TikTok" — phí quảng cáo bị trừ trực tiếp
    // từ số dư, không phải doanh thu bán hàng) nhưng ID của các dòng này vẫn
    // là chuỗi số thuần nên không bị lọc bởi looksLikeNumericId ở trên. Nếu
    // file có cột "Loại giao dịch", chỉ giữ lại các dòng thực sự là "Đơn hàng".
    if (map.transactionType !== undefined) {
      const txType = normalizeHeader(row[map.transactionType]);
      if (txType && !txType.includes("đơn hàng") && !txType.includes("don hang")) return;
    }

    const sellerSku = map.sellerSku !== undefined ? String(row[map.sellerSku] ?? "").trim() : "";
    const subtotal = map.subtotal !== undefined ? toNumber(row[map.subtotal]) : 0;

    const feeFields = [
      "shippingFee", "transactionFee", "commissionFee", "affiliateCommission",
      "voucherXtra", "orderProcessingFee", "taxWithheld", "serviceFee",
      "paymentFee", "freeshipXtra",
    ];
    const fees = {};
    feeFields.forEach((f) => {
      fees[f] = map[f] !== undefined ? Math.abs(toNumber(row[map[f]])) : 0;
    });
    const settlementAmount = map.settlementAmount !== undefined ? toNumber(row[map.settlementAmount]) : 0;
    const rawStatus = map.orderStatus !== undefined ? row[map.orderStatus] : "";
    const rawCancelType = map.cancelType !== undefined ? row[map.cancelType] : "";
    const rawDate = map.orderDate !== undefined ? row[map.orderDate] : "";

    if (!grouped.has(orderId)) {
      grouped.set(orderId, {
        id: orderId,
        platform,
        date: toIsoDate(rawDate),
        status: guessOrderStatus(rawStatus, rawCancelType),
        items: [],
        gmv: 0,
        settlementAmount: 0,
        fees: Object.fromEntries(feeFields.map((f) => [f, 0])),
      });
    }
    const ord = grouped.get(orderId);
    ord.gmv += subtotal;
    ord.settlementAmount = Math.max(ord.settlementAmount, settlementAmount);
    feeFields.forEach((f) => {
      ord.fees[f] = Math.max(ord.fees[f], fees[f]);
    });

    if (sellerSku) {
      const cfg = skuConfigMap.get(sellerSku.toLowerCase());
      ord.items.push({
        sellerSku,
        subtotal,
        cogs: cfg ? cfg.cogs : 0,
        type: cfg ? cfg.type : inferSkuType(sellerSku + " " + (map.productName !== undefined ? row[map.productName] : "")),
      });
      if (!cfg) missingSkuSet.add(sellerSku);
    }
  });

  if (missingSkuSet.size > 0) {
    warnings.push(`${missingSkuSet.size} SKU chưa có trong Config SKU (tạm tính giá vốn = 0đ): ${Array.from(missingSkuSet).slice(0, 8).join(", ")}${missingSkuSet.size > 8 ? "..." : ""}`);
  }
  if (map.settlementAmount === undefined) {
    warnings.push("File không có cột số tiền quyết toán/Order Amount — Doanh thu thực tế tạm lấy bằng GMV.");
  }
  if (!feeFieldsFoundAny(map)) {
    warnings.push("File này không có cột phí sàn chi tiết (Transaction fee/Commission/...) — Tổng phí sàn sẽ là 0đ cho các đơn này cho tới khi bạn import file Đối soát/Statement of Account.");
  }

  const orders = Array.from(grouped.values()).map((ord) => {
    const feesTotal = Object.values(ord.fees).reduce((s, v) => s + v, 0);
    const cogsTotal = ord.items.reduce((s, it) => s + it.cogs, 0);
    const settlementAmount = ord.settlementAmount > 0 ? ord.settlementAmount : ord.gmv;
    return { ...ord, settlementAmount, feesTotal, cogsTotal };
  });

  return { orders, warnings, ok: true };
}

function feeFieldsFoundAny(map) {
  return ["shippingFee", "transactionFee", "commissionFee", "affiliateCommission", "voucherXtra", "orderProcessingFee", "taxWithheld", "serviceFee", "paymentFee", "freeshipXtra"].some((f) => map[f] !== undefined);
}

/** Nhận diện loại SKU (Kính/Hộp/Khăn) từ chuỗi văn bản (SKU + tên sản phẩm).
 * Kiểm tra "Hộp" trước vì các sản phẩm dạng "Hộp kính..." chứa cả 2 từ khóa. */
function inferSkuType(text) {
  const t = String(text || "");
  if (/h[ộo]p/i.test(t)) return "hop";
  if (/kh[ăa]n/i.test(t)) return "khan";
  if (/k[íi]nh/i.test(t)) return "kinh";
  return "khac";
}

/**
 * Parse file Đơn trả hàng / Hoàn tiền — trả về map orderId -> {reason}.
 * CHỈ đánh dấu đơn là "Hoàn trả" khi Return Status thực sự đã hoàn tất
 * (Completed); các yêu cầu bị từ chối/hủy (Refund rejected, Request Canceled)
 * không được tính là mất doanh thu vì đơn hàng vẫn giữ nguyên trạng thái bán thành công.
 */
function parseReturnsRows(headers, rows) {
  const map = buildHeaderMap(headers);
  const result = new Map();
  const warnings = [];
  if (map.orderId === undefined) {
    warnings.push("Không tìm thấy cột 'Order ID' trong file trả hàng/hoàn tiền.");
    return { returns: result, warnings, ok: false };
  }
  let skippedNotCompleted = 0;
  rows.forEach((row) => {
    const orderId = String(row[map.orderId] ?? "").trim();
    if (!orderId || !looksLikeNumericId(orderId)) return;
    const rsRaw = map.returnStatus !== undefined ? row[map.returnStatus] : "";
    const rs = normalizeHeader(rsRaw);
    const isRejectedOrCancelled = rs.includes("reject") || rs.includes("từ chối") || rs.includes("tu choi") || (rs.includes("cancel") && !rs.includes("complet"));
    if (isRejectedOrCancelled) { skippedNotCompleted++; return; }
    const isCompleted = rs === "" || rs.includes("complet") || rs.includes("hoàn thành") || rs.includes("hoàn tất") || rs.includes("hoan thanh") || rs.includes("hoan tat") || rs.includes("approved") || rs.includes("refunded");
    if (!isCompleted) { skippedNotCompleted++; return; }
    const reason = map.returnReason !== undefined ? String(row[map.returnReason] ?? "") : "";
    result.set(orderId, { reason });
  });
  if (skippedNotCompleted > 0) {
    warnings.push(`Bỏ qua ${skippedNotCompleted} yêu cầu hoàn tiền chưa hoàn tất/bị từ chối — các đơn này vẫn được tính là bán thành công.`);
  }
  return { returns: result, warnings, ok: true };
}

/**
 * Parse file Config SKU (Seller SKU/Loại/Tên/Giá vốn) HOẶC file "Batch Edit
 * Template" xuất trực tiếp từ TikTok Seller Center (product_id, product_name,
 * sku_id, price, quantity, seller_sku) — loại file này KHÔNG có cột giá vốn
 * nên hệ thống tạm đặt = 0đ, người dùng chỉnh sửa trực tiếp trong Config Center.
 */
// Các cụm từ chắc chắn là VĂN BẢN HƯỚNG DẪN/METADATA của file Batch Edit
// Template TikTok Seller Center (dòng 2-5: dòng "V4/All_Information/metric",
// dòng nhãn tiếng Việt, dòng "Bắt buộc/Không bắt buộc", dòng "Không thể
// chỉnh sửa/hướng dẫn nhập liệu") — KHÔNG phải dữ liệu SKU thật.
const SKU_JUNK_MARKERS = [
  "không thể chỉnh sửa", "khong the chinh sua",
  "bắt buộc", "bat buoc",
  "không bắt buộc", "khong bat buoc",
  "metric", "all_information",
  "sku người bán", "sku nguoi ban",
  "mã nhận dạng sản phẩm", "ma nhan dang san pham",
];
function isJunkSkuRow(sellerSkuRaw) {
  const s = String(sellerSkuRaw ?? "").trim();
  if (!s) return true; // seller_sku rỗng -> chắc chắn không phải dòng dữ liệu
  const norm = normalizeHeader(s);
  return SKU_JUNK_MARKERS.some((m) => norm.includes(m));
}

function parseSkuConfigRows(headers, rows) {
  const map = buildHeaderMap(headers);
  const warnings = [];
  if (map.sellerSku === undefined) {
    warnings.push("File cần có cột 'Seller SKU' (SKU người bán) để làm khóa ánh xạ giá vốn.");
    return { entries: [], warnings, ok: false };
  }
  const hasCogsColumn = map.cogs !== undefined;
  if (!hasCogsColumn) {
    warnings.push("File không có cột 'Giá vốn' (có vẻ đây là file xuất Batch Edit Template từ Seller Center, chỉ chứa giá bán) — hệ thống tạm đặt Giá vốn = 0đ cho các SKU này, vui lòng chỉnh trực tiếp trong bảng Config SKU.");
  }
  const entries = [];
  const seen = new Set();
  let skippedJunkRows = 0;
  rows.forEach((row) => {
    const sku = String(row[map.sellerSku] ?? "").trim();
    // Lớp lọc 1: kiểm tra trực tiếp nội dung ô seller_sku — bắt các dòng
    // rác/hướng dẫn của Batch Edit Template ("Không thể chỉnh sửa", "Bắt
    // buộc", "metric"...) ngay cả khi file không có cột sku_id để đối chiếu.
    if (isJunkSkuRow(sku)) { if (sku) skippedJunkRows++; return; }
    // Lớp lọc 2 (dự phòng, khi có cột sku_id): dữ liệu SKU thật luôn có
    // ID SKU là chuỗi số thuần; dòng mô tả/hướng dẫn thì không.
    if (map.skuId !== undefined) {
      const idVal = row[map.skuId];
      if (!looksLikeNumericId(idVal)) { skippedJunkRows++; return; }
    }
    const key = sku.toLowerCase();
    if (seen.has(key)) return; // giữ dòng xuất hiện đầu tiên cho mỗi SKU
    seen.add(key);
    const cogs = hasCogsColumn ? toNumber(row[map.cogs]) : 0;
    const productName = map.productName !== undefined ? String(row[map.productName] ?? "") : sku;
    const rawType = map.skuType !== undefined ? String(row[map.skuType] ?? "") : "";
    const type = inferSkuType(rawType + " " + productName);
    const stockQty = map.stockQty !== undefined ? toNumber(row[map.stockQty]) : 0;
    const sellingPrice = map.sellingPrice !== undefined ? toNumber(row[map.sellingPrice]) : 0;
    entries.push({ sku, type, name: productName, cogs, stockQty, sellingPrice });
  });
  if (map.stockQty === undefined) {
    warnings.push("File không có cột 'Số lượng' (tồn kho) — cảnh báo tồn kho thấp sẽ không khả dụng cho các SKU import từ file này.");
  }
  if (skippedJunkRows > 0) {
    warnings.push(`Đã tự động bỏ qua ${skippedJunkRows} dòng hướng dẫn/metadata của Batch Edit Template (không phải dữ liệu SKU thật).`);
  }
  return { entries, warnings, ok: true };
}

/** Parse file báo cáo quảng cáo -> tổng chi phí ads (VNĐ). Cấu trúc sẵn sàng
 * nhận, mặc định hệ thống vẫn tính Ads = 0đ trừ khi người dùng import. */
function parseAdsRows(headers, rows) {
  const map = buildHeaderMap(headers);
  const warnings = [];
  if (map.adsCost === undefined) {
    warnings.push("Không tìm thấy cột chi phí quảng cáo (Cost/Amount Spent) — Ads sẽ giữ mặc định 0đ.");
    return { total: 0, warnings, ok: false };
  }
  let total = 0;
  rows.forEach((row) => {
    total += toNumber(row[map.adsCost]);
  });
  return { total, warnings, ok: true };
}

/* ============================================================================
 * 4. FINANCIAL ENGINE — công thức tính toán đa tầng
 * ========================================================================== */

/** Tính các chỉ số cấp Đơn hàng */
function computeOrderMetrics(order, settings) {
  const isCancelled = order.status === STATUS.CANCELLED;
  const isReturned = order.status === STATUS.RETURNED;
  const isLoss = isCancelled || isReturned;
  const nettRevenue = isLoss ? 0 : (order.settlementAmount > 0 ? order.settlementAmount : order.gmv);
  // Đơn HỦY: không phát sinh phí sàn lẫn phí đóng gói (chưa từng đóng gói/giao).
  // Đơn HOÀN TRẢ: TikTok vẫn thu phí hoa hồng + phí giao dịch trên đơn hoàn
  // (áp dụng từ 01/04/2026 — theo ELLA_PROFIT_CALCULATOR_v2_1.xlsx), nên GIỮ
  // nguyên phí sàn thật đọc được từ file Quyết toán; đồng thời vẫn tính phí
  // đóng gói 6.000đ vì đây là tổn thất vật tư bao bì đã thực chi.
  // Đơn THÀNH CÔNG: tính đủ cả phí sàn thật + phí đóng gói như bình thường.
  const feesTotal = isCancelled ? 0 : order.feesTotal;
  const packagingFee = isCancelled ? 0 : settings.packagingFee;
  const profitBeforeAds = nettRevenue - order.cogsTotal - feesTotal - packagingFee;
  return { ...order, nettRevenue, feesTotal, packagingFee, profitBeforeAds };
}

/** Gộp danh sách đơn hàng đã compute -> các chỉ số tổng hợp (aggregate) */
function aggregateOrders(computedOrders, adsFee = 0, monthlyFixedCost = 0, otherVariableCost = 0) {
  const base = {
    orderCount: computedOrders.length,
    successCount: 0,
    cancelledCount: 0,
    returnedCount: 0,
    gmv: 0,
    nettRevenue: 0,
    cogsTotal: 0,
    feesTotal: 0,
    packagingTotal: 0,
    profitBeforeAds: 0,
  };
  for (const o of computedOrders) {
    base.gmv += o.gmv;
    base.nettRevenue += o.nettRevenue;
    base.cogsTotal += o.cogsTotal;
    base.feesTotal += o.feesTotal;
    base.packagingTotal += o.packagingFee;
    base.profitBeforeAds += o.profitBeforeAds;
    if (o.status === STATUS.SUCCESS) base.successCount++;
    else if (o.status === STATUS.CANCELLED) base.cancelledCount++;
    else base.returnedCount++;
  }
  const returnRate = base.orderCount > 0 ? ((base.cancelledCount + base.returnedCount) / base.orderCount) * 100 : 0;
  const profitAfterAds = base.profitBeforeAds - adsFee - monthlyFixedCost - otherVariableCost;
  const margin = base.nettRevenue > 0 ? (profitAfterAds / base.nettRevenue) * 100 : 0;
  return { ...base, returnRate, adsFee, fixedCost: monthlyFixedCost, otherVariableCost, profitAfterAds, margin };
}

function aggregateByDate(computedOrders) {
  const map = new Map();
  for (const o of computedOrders) {
    if (!map.has(o.date)) {
      map.set(o.date, { date: o.date, nettRevenue: 0, cost: 0, profit: 0 });
    }
    const d = map.get(o.date);
    d.nettRevenue += o.nettRevenue;
    d.cost += o.cogsTotal + o.feesTotal + o.packagingFee;
    d.profit += o.profitBeforeAds;
  }
  return Array.from(map.values()).sort((a, b) => a.date.localeCompare(b.date));
}

function isWithinRange(dateStr, start, end) {
  if (!start && !end) return true;
  if (start && dateStr < start) return false;
  if (end && dateStr > end) return false;
  return true;
}

/* ============================================================================
 * 5. DEMO / SEED DATA — để dashboard có dữ liệu minh họa trước khi import
 * ========================================================================== */

/*
 * Dữ liệu SKU & đơn hàng TikTok Shop dưới đây được nhập THẬT từ 5 file Seller
 * Center do người dùng cung cấp (batch-edit template x3, Tất cả đơn hàng,
 * Đơn trả hàng hoàn tiền), xử lý bằng đúng engine parsing phía trên. Giá vốn
 * (cogs) đang = 0đ cho toàn bộ SKU vì Seller Center không xuất giá vốn — vào
 * Config Center để nhập tay. Tổng phí sàn đang = 0đ vì file "Tất cả đơn hàng"
 * không chứa cột phí chi tiết — import thêm file Đối soát/Statement of
 * Account để có số liệu đầy đủ.
 */
const REAL_TIKTOK_SKU_CONFIG = [
  {
    "sku": "EL-0210",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mắt Nữ Phân Cực Cao Cấp Chống Tia UV400 Form Hàn Quốc Thời Trang Cá Tính |EL-02",
    "cogs": 0,
    "stockQty": 61
  },
  {
    "sku": "EL-0232",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mắt Nữ Phân Cực Cao Cấp Chống Tia UV400 Form Hàn Quốc Thời Trang Cá Tính |EL-02",
    "cogs": 0,
    "stockQty": 95
  },
  {
    "sku": "EL-0240",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mắt Nữ Phân Cực Cao Cấp Chống Tia UV400 Form Hàn Quốc Thời Trang Cá Tính |EL-02",
    "cogs": 0,
    "stockQty": 67
  },
  {
    "sku": "EL-0199",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mắt Nữ Phân Cực Cao Cấp Chống Tia UV400 Form Hàn Quốc Cá Tính | EL-01",
    "cogs": 0,
    "stockQty": 87
  },
  {
    "sku": "EL-0140",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mắt Nữ Phân Cực Cao Cấp Chống Tia UV400 Form Hàn Quốc Cá Tính | EL-01",
    "cogs": 0,
    "stockQty": 99
  },
  {
    "sku": "EL-0130",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mắt Nữ Phân Cực Cao Cấp Chống Tia UV400 Form Hàn Quốc Cá Tính | EL-01",
    "cogs": 0,
    "stockQty": 98
  },
  {
    "sku": "EL-0114",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mắt Nữ Phân Cực Cao Cấp Chống Tia UV400 Form Hàn Quốc Cá Tính | EL-01",
    "cogs": 0,
    "stockQty": 100
  },
  {
    "sku": "EL - 0310",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mắt Nữ Phân Cực Phong Cách Cổ Điển Chống Tia UV400 Street Style Thời Trang | EL-03",
    "cogs": 0,
    "stockQty": 75
  },
  {
    "sku": "EL - 0340",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mắt Nữ Phân Cực Phong Cách Cổ Điển Chống Tia UV400 Street Style Thời Trang | EL-03",
    "cogs": 0,
    "stockQty": 97
  },
  {
    "sku": "EL - 0352",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mắt Nữ Phân Cực Phong Cách Cổ Điển Chống Tia UV400 Street Style Thời Trang | EL-03",
    "cogs": 0,
    "stockQty": 99
  },
  {
    "sku": "EL - 0332 Nâu - Nâu",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mắt Nữ Phân Cực Phong Cách Cổ Điển Chống Tia UV400 Street Style Thời Trang | EL-03",
    "cogs": 0,
    "stockQty": 0
  },
  {
    "sku": "EL-0410",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mắt Nữ Phân Cực Hiện Đại Chống Tia UV400 Phong Cách Thời Trang Cao Cấp | EL-04",
    "cogs": 0,
    "stockQty": 97
  },
  {
    "sku": "EL-0432",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mắt Nữ Phân Cực Hiện Đại Chống Tia UV400 Phong Cách Thời Trang Cao Cấp | EL-04",
    "cogs": 0,
    "stockQty": 99
  },
  {
    "sku": "EL-0440",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mắt Nữ Phân Cực Hiện Đại Chống Tia UV400 Phong Cách Thời Trang Cao Cấp | EL-04",
    "cogs": 0,
    "stockQty": 100
  },
  {
    "sku": "EL-0414",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mắt Nữ Phân Cực Hiện Đại Chống Tia UV400 Phong Cách Thời Trang Cao Cấp | EL-04",
    "cogs": 0,
    "stockQty": 0
  },
  {
    "sku": "EL-0610",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mắt Nữ Phân Cực Hiện Đại Chống Tia UV400 Thời Trang Xu Hướng | EL-06",
    "cogs": 0,
    "stockQty": 99
  },
  {
    "sku": "EL-0640",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mắt Nữ Phân Cực Hiện Đại Chống Tia UV400 Thời Trang Xu Hướng | EL-06",
    "cogs": 0,
    "stockQty": 100
  },
  {
    "sku": "EL-0661",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mắt Nữ Phân Cực Hiện Đại Chống Tia UV400 Thời Trang Xu Hướng | EL-06",
    "cogs": 0,
    "stockQty": 100
  },
  {
    "sku": "EL-0615",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mắt Nữ Phân Cực Hiện Đại Chống Tia UV400 Thời Trang Xu Hướng | EL-06",
    "cogs": 0,
    "stockQty": 100
  },
  {
    "sku": "EL-0710",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mắt Nữ Phân Cực Form Oval Hiện Đại Chống Tia UV400 Thời Trang Cao Cấp | EL-07",
    "cogs": 0,
    "stockQty": 98
  },
  {
    "sku": "EL-0732",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mắt Nữ Phân Cực Form Oval Hiện Đại Chống Tia UV400 Thời Trang Cao Cấp | EL-07",
    "cogs": 0,
    "stockQty": 100
  },
  {
    "sku": "EL-0740",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mắt Nữ Phân Cực Form Oval Hiện Đại Chống Tia UV400 Thời Trang Cao Cấp | EL-07",
    "cogs": 0,
    "stockQty": 98
  },
  {
    "sku": "EL-0810",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mắt Nữ Phân Cực Form Vuông  Chống Tia UV400 Thời Trang Thanh Lịch | EL-08",
    "cogs": 0,
    "stockQty": 50
  },
  {
    "sku": "EL-0832",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mắt Nữ Phân Cực Form Vuông  Chống Tia UV400 Thời Trang Thanh Lịch | EL-08",
    "cogs": 0,
    "stockQty": 89
  },
  {
    "sku": "EL-0861",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mắt Nữ Phân Cực Form Vuông  Chống Tia UV400 Thời Trang Thanh Lịch | EL-08",
    "cogs": 0,
    "stockQty": 96
  },
  {
    "sku": "EL-0820",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mắt Nữ Phân Cực Form Vuông  Chống Tia UV400 Thời Trang Thanh Lịch | EL-08",
    "cogs": 0,
    "stockQty": 95
  },
  {
    "sku": "El-0910",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mắt Nữ Phân Cực Form Vuông Hiện Đại Chống Tia UV400 Thời Trang Cá Tính | EL-09",
    "cogs": 0,
    "stockQty": 65
  },
  {
    "sku": "EL-0933",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mắt Nữ Phân Cực Form Vuông Hiện Đại Chống Tia UV400 Thời Trang Cá Tính | EL-09",
    "cogs": 0,
    "stockQty": 69
  },
  {
    "sku": "EL-0940",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mắt Nữ Phân Cực Form Vuông Hiện Đại Chống Tia UV400 Thời Trang Cá Tính | EL-09",
    "cogs": 0,
    "stockQty": 77
  },
  {
    "sku": "EL-0930",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mắt Nữ Phân Cực Form Vuông Hiện Đại Chống Tia UV400 Thời Trang Cá Tính | EL-09",
    "cogs": 0,
    "stockQty": 86
  },
  {
    "sku": "EL-1010",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mắt Nữ Phân Cực Form Vuông Cá Tính Chống Tia UV400 Thời Trang Hiện Đại | EL-10",
    "cogs": 0,
    "stockQty": 94
  },
  {
    "sku": "EL-1040",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mắt Nữ Phân Cực Form Vuông Cá Tính Chống Tia UV400 Thời Trang Hiện Đại | EL-10",
    "cogs": 0,
    "stockQty": 97
  },
  {
    "sku": "EL-1014",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mắt Nữ Phân Cực Form Vuông Cá Tính Chống Tia UV400 Thời Trang Hiện Đại | EL-10",
    "cogs": 0,
    "stockQty": 98
  },
  {
    "sku": "EL-1020",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mắt Nữ Phân Cực Form Vuông Cá Tính Chống Tia UV400 Thời Trang Hiện Đại | EL-10",
    "cogs": 0,
    "stockQty": 100
  },
  {
    "sku": "EL-1110",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Trong Gọng Oval Phong Cách Y2K Unisex Thời Trang Hiện Đại | EL-11",
    "cogs": 0,
    "stockQty": 99
  },
  {
    "sku": "EL-1121",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Trong Gọng Oval Phong Cách Y2K Unisex Thời Trang Hiện Đại | EL-11",
    "cogs": 0,
    "stockQty": 100
  },
  {
    "sku": "EL-1261",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Trong Gọng Oval Phong Cách Y2K Unisex Thời Trang Hiện Đại | EL-11",
    "cogs": 0,
    "stockQty": 100
  },
  {
    "sku": "EL-1610",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Râm Nam Nữ Cá Tính Đi Biển Chống Tia UV400 Phong Cách Thời Trang | EL-16",
    "cogs": 0,
    "stockQty": 71
  },
  {
    "sku": "EL-1633",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Râm Nam Nữ Cá Tính Đi Biển Chống Tia UV400 Phong Cách Thời Trang | EL-16",
    "cogs": 0,
    "stockQty": 0
  },
  {
    "sku": "EL-1640",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Râm Nam Nữ Cá Tính Đi Biển Chống Tia UV400 Phong Cách Thời Trang | EL-16",
    "cogs": 0,
    "stockQty": 100
  },
  {
    "sku": "EL-1314",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Trong Unisex Phong Cách Thời Trang Hiện Đại Chống Tia UV400 Cá Tính | EL-13",
    "cogs": 0,
    "stockQty": 99
  },
  {
    "sku": "EL-1310",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Trong Unisex Phong Cách Thời Trang Hiện Đại Chống Tia UV400 Cá Tính | EL-13",
    "cogs": 0,
    "stockQty": 99
  },
  {
    "sku": "EL-1333",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Trong Unisex Phong Cách Thời Trang Hiện Đại Chống Tia UV400 Cá Tính | EL-13",
    "cogs": 0,
    "stockQty": 99
  },
  {
    "sku": "EL-1352",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Trong Unisex Phong Cách Thời Trang Hiện Đại Chống Tia UV400 Cá Tính | EL-13",
    "cogs": 0,
    "stockQty": 100
  },
  {
    "sku": "EL-1340",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Trong Unisex Phong Cách Thời Trang Hiện Đại Chống Tia UV400 Cá Tính | EL-13",
    "cogs": 0,
    "stockQty": 100
  },
  {
    "sku": "EL-1210",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Gọng Kính Cận Unisex Form Vuông Tròn Có Thể Lắp Tròng Cận Phong Cách Hiện Đại | EL-12",
    "cogs": 0,
    "stockQty": 98
  },
  {
    "sku": "EL-1222",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Gọng Kính Cận Unisex Form Vuông Tròn Có Thể Lắp Tròng Cận Phong Cách Hiện Đại | EL-12",
    "cogs": 0,
    "stockQty": 100
  },
  {
    "sku": "EL-1252",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Gọng Kính Cận Unisex Form Vuông Tròn Có Thể Lắp Tròng Cận Phong Cách Hiện Đại | EL-12",
    "cogs": 0,
    "stockQty": 100
  },
  {
    "sku": "EL-1710",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mắt Nữ Gọng Kim Loại Form Mắt Mèo Chống Tia UV400 Thời Trang Sang Trọng | EL-17",
    "cogs": 0,
    "stockQty": 33
  },
  {
    "sku": "EL-1711",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mắt Nữ Gọng Kim Loại Form Mắt Mèo Chống Tia UV400 Thời Trang Sang Trọng | EL-17",
    "cogs": 0,
    "stockQty": 88
  },
  {
    "sku": "EL-1732",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mắt Nữ Gọng Kim Loại Form Mắt Mèo Chống Tia UV400 Thời Trang Sang Trọng | EL-17",
    "cogs": 0,
    "stockQty": 89
  },
  {
    "sku": "EL-1740",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mắt Nữ Gọng Kim Loại Form Mắt Mèo Chống Tia UV400 Thời Trang Sang Trọng | EL-17",
    "cogs": 0,
    "stockQty": 75
  },
  {
    "sku": "EL-1763",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mắt Nữ Gọng Kim Loại Form Mắt Mèo Chống Tia UV400 Thời Trang Sang Trọng | EL-17",
    "cogs": 0,
    "stockQty": 99
  },
  {
    "sku": "EL-1910",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mắt Nữ Tròng Bầu Dục Gọng Kim Loại Chống Tia UV400 Phong Cách Thời Trang Hot Trend | EL-19",
    "cogs": 0,
    "stockQty": 93
  },
  {
    "sku": "EL-1961",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mắt Nữ Tròng Bầu Dục Gọng Kim Loại Chống Tia UV400 Phong Cách Thời Trang Hot Trend | EL-19",
    "cogs": 0,
    "stockQty": 100
  },
  {
    "sku": "EL-1930",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mắt Nữ Tròng Bầu Dục Gọng Kim Loại Chống Tia UV400 Phong Cách Thời Trang Hot Trend | EL-19",
    "cogs": 0,
    "stockQty": 97
  },
  {
    "sku": "EL-1914",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mắt Nữ Tròng Bầu Dục Gọng Kim Loại Chống Tia UV400 Phong Cách Thời Trang Hot Trend | EL-19",
    "cogs": 0,
    "stockQty": 99
  },
  {
    "sku": "EL-2252",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mắt Tròn Không Gọng Phong Cách Âu Mỹ Chống Tia UV400 Hot Trend 2026 | EL-22",
    "cogs": 0,
    "stockQty": 100
  },
  {
    "sku": "EL-2210",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mắt Tròn Không Gọng Phong Cách Âu Mỹ Chống Tia UV400 Hot Trend 2026 | EL-22",
    "cogs": 0,
    "stockQty": 99
  },
  {
    "sku": "EL-2230",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mắt Tròn Không Gọng Phong Cách Âu Mỹ Chống Tia UV400 Hot Trend 2026 | EL-22",
    "cogs": 0,
    "stockQty": 99
  },
  {
    "sku": "EL-2110",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mát Cánh Bướm Không Gọng Siêu Nhẹ Chống Tia UV400 Phong Cách Thời Trang | EL-21",
    "cogs": 0,
    "stockQty": 94
  },
  {
    "sku": "EL-2152",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mát Cánh Bướm Không Gọng Siêu Nhẹ Chống Tia UV400 Phong Cách Thời Trang | EL-21",
    "cogs": 0,
    "stockQty": 95
  },
  {
    "sku": "EL-2132",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mát Cánh Bướm Không Gọng Siêu Nhẹ Chống Tia UV400 Phong Cách Thời Trang | EL-21",
    "cogs": 0,
    "stockQty": 93
  },
  {
    "sku": "EL-2310",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mắt Nữ Form Mắt Mèo Đính Đá Chống Tia UV400 Phong Cách Hiện Đại Sang Trọng | EL-23",
    "cogs": 0,
    "stockQty": 89
  },
  {
    "sku": "EL-2330",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mắt Nữ Form Mắt Mèo Đính Đá Chống Tia UV400 Phong Cách Hiện Đại Sang Trọng | EL-23",
    "cogs": 0,
    "stockQty": 97
  },
  {
    "sku": "EL-2340",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mắt Nữ Form Mắt Mèo Đính Đá Chống Tia UV400 Phong Cách Hiện Đại Sang Trọng | EL-23",
    "cogs": 0,
    "stockQty": 100
  },
  {
    "sku": "EL-2432",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mắt Nữ Gọng Kim Loại Cao Cấp Chống Tia UV400 Phong Cách Hiện Đại Sang Trọng | EL-24",
    "cogs": 0,
    "stockQty": 99
  },
  {
    "sku": "EL-2410",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mắt Nữ Gọng Kim Loại Cao Cấp Chống Tia UV400 Phong Cách Hiện Đại Sang Trọng | EL-24",
    "cogs": 0,
    "stockQty": 98
  },
  {
    "sku": "EL-2431",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mắt Nữ Gọng Kim Loại Cao Cấp Chống Tia UV400 Phong Cách Hiện Đại Sang Trọng | EL-24",
    "cogs": 0,
    "stockQty": 96
  },
  {
    "sku": "EL-2450",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mắt Nữ Gọng Kim Loại Cao Cấp Chống Tia UV400 Phong Cách Hiện Đại Sang Trọng | EL-24",
    "cogs": 0,
    "stockQty": 97
  },
  {
    "sku": "EL-2440",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mắt Nữ Gọng Kim Loại Cao Cấp Chống Tia UV400 Phong Cách Hiện Đại Sang Trọng | EL-24",
    "cogs": 0,
    "stockQty": 99
  },
  {
    "sku": "EL-2470",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mắt Nữ Gọng Kim Loại Cao Cấp Chống Tia UV400 Phong Cách Hiện Đại Sang Trọng | EL-24",
    "cogs": 0,
    "stockQty": 98
  },
  {
    "sku": "EL-2550",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mắt Nữ Mắt Mèo Gọng Kim Loại Đổi Màu Chống Tia UV400 Phong Cách Retro Hàn Quốc | EL-25",
    "cogs": 0,
    "stockQty": 93
  },
  {
    "sku": "EL-2533",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mắt Nữ Mắt Mèo Gọng Kim Loại Đổi Màu Chống Tia UV400 Phong Cách Retro Hàn Quốc | EL-25",
    "cogs": 0,
    "stockQty": 96
  },
  {
    "sku": "EL-2530",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mắt Nữ Mắt Mèo Gọng Kim Loại Đổi Màu Chống Tia UV400 Phong Cách Retro Hàn Quốc | EL-25",
    "cogs": 0,
    "stockQty": 93
  },
  {
    "sku": "EL-2510",
    "type": "kinh",
    "name": "ELLA Eyewear Việt Nam | Kính Mắt Nữ Mắt Mèo Gọng Kim Loại Đổi Màu Chống Tia UV400 Phong Cách Retro Hàn Quốc | EL-25",
    "cogs": 0,
    "stockQty": 92
  },
  {
    "sku": "EL-Hop",
    "type": "hop",
    "name": "Hộp kính Premium | ELLA Eyewear Việt Nam",
    "cogs": 0,
    "stockQty": 569
  },
  {
    "sku": "0122 Vàng hồng",
    "type": "hop",
    "name": "ELLA Accents | Dây Chuyền Bạc Thiên Nga Đính Đá Zircon Cao Cấp | Thanh Lịch Nữ Tính [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "0121 Bạc Trắng",
    "type": "hop",
    "name": "ELLA Accents | Dây Chuyền Bạc Thiên Nga Đính Đá Zircon Cao Cấp | Thanh Lịch Nữ Tính [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "0181 Bạc Xanh",
    "type": "hop",
    "name": "ELLA Accents | Dây Chuyền Bạc Thiên Nga Đính Đá Zircon Cao Cấp | Thanh Lịch Nữ Tính [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "0199 Vàng Trắng",
    "type": "hop",
    "name": "ELLA Accents | Dây Chuyền Bạc Thiên Nga Đính Đá Zircon Cao Cấp | Thanh Lịch Nữ Tính [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "0114 Bạc Trắng",
    "type": "hop",
    "name": "ELLA Accents | Dây Chuyền Bạc Thiên Nga Đính Đá Zircon Cao Cấp | Thanh Lịch Nữ Tính [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "0299 Bạc Trắng",
    "type": "hop",
    "name": "ELLA Accents | Dây Chuyền Bạc Thiên Nga Full Đá Zircon Cao Cấp | Sang Trọng Nổi Bật [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "0296 Vàng Trắng",
    "type": "hop",
    "name": "ELLA Accents | Dây Chuyền Bạc Thiên Nga Full Đá Zircon Cao Cấp | Sang Trọng Nổi Bật [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "0210 Vàng Đen",
    "type": "hop",
    "name": "ELLA Accents | Dây Chuyền Bạc Thiên Nga Full Đá Zircon Cao Cấp | Sang Trọng Nổi Bật [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "0399 Bạc Trắng",
    "type": "hop",
    "name": "ELLA Accents | Dây Chuyền Bạc Thiên Nga Đá Zircon Dáng Lông Vũ | Sang Trọng Tinh Tế [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "0322 Vàng Hồng",
    "type": "hop",
    "name": "ELLA Accents | Dây Chuyền Bạc Thiên Nga Đá Zircon Dáng Lông Vũ | Sang Trọng Tinh Tế [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "0310 Vàng Đen",
    "type": "hop",
    "name": "ELLA Accents | Dây Chuyền Bạc Thiên Nga Đá Zircon Dáng Lông Vũ | Sang Trọng Tinh Tế [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "0510 Bạc 4 Lá Đen",
    "type": "hop",
    "name": "ELLA Accents | Set Dây Chuyền Và Bông Tai Cỏ 4 Lá Cao Cấp | Sang Trọng May Mắn [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "05 Tai Bạc Đen",
    "type": "hop",
    "name": "ELLA Accents | Set Dây Chuyền Và Bông Tai Cỏ 4 Lá Cao Cấp | Sang Trọng May Mắn [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "Set Bạc 4 Lá Đen",
    "type": "hop",
    "name": "ELLA Accents | Set Dây Chuyền Và Bông Tai Cỏ 4 Lá Cao Cấp | Sang Trọng May Mắn [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "0596 Vàng 4 Lá Đen",
    "type": "hop",
    "name": "ELLA Accents | Set Dây Chuyền Và Bông Tai Cỏ 4 Lá Cao Cấp | Sang Trọng May Mắn [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "05 Tai Vàng Đen",
    "type": "hop",
    "name": "ELLA Accents | Set Dây Chuyền Và Bông Tai Cỏ 4 Lá Cao Cấp | Sang Trọng May Mắn [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "Set Vàng 4 Lá Đen",
    "type": "hop",
    "name": "ELLA Accents | Set Dây Chuyền Và Bông Tai Cỏ 4 Lá Cao Cấp | Sang Trọng May Mắn [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "0514 Bạc 4 Lá Trắng",
    "type": "hop",
    "name": "ELLA Accents | Set Dây Chuyền Và Bông Tai Cỏ 4 Lá Cao Cấp | Sang Trọng May Mắn [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "05 Tai Bạc Trắng",
    "type": "hop",
    "name": "ELLA Accents | Set Dây Chuyền Và Bông Tai Cỏ 4 Lá Cao Cấp | Sang Trọng May Mắn [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "Set Bạc 4 Lá Trắng",
    "type": "hop",
    "name": "ELLA Accents | Set Dây Chuyền Và Bông Tai Cỏ 4 Lá Cao Cấp | Sang Trọng May Mắn [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "0614 Bạc Nơ",
    "type": "hop",
    "name": "ELLA Accents | Bộ Trang Sức Nữ Đính Đá Zircon Cao Cấp | Thanh Lịch Sang Trọng [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 498
  },
  {
    "sku": "0652 Vàng Nơ",
    "type": "hop",
    "name": "ELLA Accents | Bộ Trang Sức Nữ Đính Đá Zircon Cao Cấp | Thanh Lịch Sang Trọng [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "0692 Cổ Bạc Trái Tim",
    "type": "hop",
    "name": "ELLA Accents | Bộ Trang Sức Nữ Đính Đá Zircon Cao Cấp | Thanh Lịch Sang Trọng [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "0693 Tay Bạc Trái Tim",
    "type": "hop",
    "name": "ELLA Accents | Bộ Trang Sức Nữ Đính Đá Zircon Cao Cấp | Thanh Lịch Sang Trọng [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "0612 Bạc Kim Băng",
    "type": "hop",
    "name": "ELLA Accents | Bộ Trang Sức Nữ Đính Đá Zircon Cao Cấp | Thanh Lịch Sang Trọng [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "0696 Vàng Kim Băng",
    "type": "hop",
    "name": "ELLA Accents | Bộ Trang Sức Nữ Đính Đá Zircon Cao Cấp | Thanh Lịch Sang Trọng [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "0910 Bạc Đen",
    "type": "hop",
    "name": "ELLA Accents | Dây Chuyền Bạc Cỏ 4 Lá Đính Đá Zircon Cao Cấp | Thanh Lịch Nổi Bật [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "0999 Bạc Trắng",
    "type": "hop",
    "name": "ELLA Accents | Dây Chuyền Bạc Cỏ 4 Lá Đính Đá Zircon Cao Cấp | Thanh Lịch Nổi Bật [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 499
  },
  {
    "sku": "0972 Bạc Đỏ",
    "type": "hop",
    "name": "ELLA Accents | Dây Chuyền Bạc Cỏ 4 Lá Đính Đá Zircon Cao Cấp | Thanh Lịch Nổi Bật [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "0950 Vàng Hồng",
    "type": "hop",
    "name": "ELLA Accents | Dây Chuyền Bạc Cỏ 4 Lá Đính Đá Zircon Cao Cấp | Thanh Lịch Nổi Bật [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "0923 Vàng Hồng",
    "type": "hop",
    "name": "ELLA Accents | Dây Chuyền Bạc Cỏ 4 Lá Đính Đá Zircon Cao Cấp | Thanh Lịch Nổi Bật [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "1010 Bạc Đen",
    "type": "hop",
    "name": "ELLA Accents | Vòng Tay Cỏ 4 Lá Đính Đá Zircon Cao Cấp | Nữ Tính Thanh Lịch [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "1040 Bạc trắng",
    "type": "hop",
    "name": "ELLA Accents | Vòng Tay Cỏ 4 Lá Đính Đá Zircon Cao Cấp | Nữ Tính Thanh Lịch [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "1072 Bạc Đỏ",
    "type": "hop",
    "name": "ELLA Accents | Vòng Tay Cỏ 4 Lá Đính Đá Zircon Cao Cấp | Nữ Tính Thanh Lịch [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "1092 Bạc Hồng",
    "type": "hop",
    "name": "ELLA Accents | Vòng Tay Cỏ 4 Lá Đính Đá Zircon Cao Cấp | Nữ Tính Thanh Lịch [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "1090 Vàng Đỏ",
    "type": "hop",
    "name": "ELLA Accents | Vòng Tay Cỏ 4 Lá Đính Đá Zircon Cao Cấp | Nữ Tính Thanh Lịch [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "1240 Bạc Trắng",
    "type": "hop",
    "name": "ELLA Accents | Vòng Tay Cỏ 4 Lá Bạc S925 Đính Đá Zircon Cao Cấp| Tinh Tế Nữ Tính [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "1140 Bạc",
    "type": "hop",
    "name": "ELLA Accents | Vòng Tay Nút Thắt Bạc S925 Đính Đá Zircon | Thanh Lịch Tinh Tế [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "1150 Vàng",
    "type": "hop",
    "name": "ELLA Accents | Vòng Tay Nút Thắt Bạc S925 Đính Đá Zircon | Thanh Lịch Tinh Tế [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "1322 Bạc Hồng",
    "type": "hop",
    "name": "ELLA Accents | Vòng Tay Cỏ 4 Lá Bạc S925 Đính Đá Zircon | Tinh Tế Thanh Lịch [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "1350 Vàng Hồng",
    "type": "hop",
    "name": "ELLA Accents | Vòng Tay Cỏ 4 Lá Bạc S925 Đính Đá Zircon | Tinh Tế Thanh Lịch [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "1372 Đỏ Vàng Hồng",
    "type": "hop",
    "name": "ELLA Accents | Vòng Tay Cỏ 4 Lá Bạc S925 Đính Đá Zircon | Tinh Tế Thanh Lịch [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "1440 Bạc Hồng",
    "type": "hop",
    "name": "ELLA Accents | Vòng Tay Cỏ 4 Lá Bạc S925 Đính Đá Zircon | Tinh Tế Sang Trọng [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 0
  },
  {
    "sku": "1450 Vàng Hồng",
    "type": "hop",
    "name": "ELLA Accents | Vòng Tay Cỏ 4 Lá Bạc S925 Đính Đá Zircon | Tinh Tế Sang Trọng [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 0
  },
  {
    "sku": "1472 Đỏ Vàng Hồng",
    "type": "hop",
    "name": "ELLA Accents | Vòng Tay Cỏ 4 Lá Bạc S925 Đính Đá Zircon | Tinh Tế Sang Trọng [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 0
  },
  {
    "sku": "15 Tay Đen Bạc",
    "type": "hop",
    "name": "ELLA Accents | SET Tay Cổ Khuyên Tai Cỏ 4 Lá Bạc S925 Đính Đá Zircon | Đơn Giản Nữ Tính [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "15 Cổ Đen Bạc",
    "type": "hop",
    "name": "ELLA Accents | SET Tay Cổ Khuyên Tai Cỏ 4 Lá Bạc S925 Đính Đá Zircon | Đơn Giản Nữ Tính [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "15 Set Đen Bạc",
    "type": "hop",
    "name": "ELLA Accents | SET Tay Cổ Khuyên Tai Cỏ 4 Lá Bạc S925 Đính Đá Zircon | Đơn Giản Nữ Tính [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "15 Tay Đen Vàng",
    "type": "hop",
    "name": "ELLA Accents | SET Tay Cổ Khuyên Tai Cỏ 4 Lá Bạc S925 Đính Đá Zircon | Đơn Giản Nữ Tính [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "15 Cổ Đen Vàng",
    "type": "hop",
    "name": "ELLA Accents | SET Tay Cổ Khuyên Tai Cỏ 4 Lá Bạc S925 Đính Đá Zircon | Đơn Giản Nữ Tính [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "15 Set Đen Vàng",
    "type": "hop",
    "name": "ELLA Accents | SET Tay Cổ Khuyên Tai Cỏ 4 Lá Bạc S925 Đính Đá Zircon | Đơn Giản Nữ Tính [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "15 Tay Đỏ Bạc",
    "type": "hop",
    "name": "ELLA Accents | SET Tay Cổ Khuyên Tai Cỏ 4 Lá Bạc S925 Đính Đá Zircon | Đơn Giản Nữ Tính [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "15 Cổ Đỏ Bạc",
    "type": "hop",
    "name": "ELLA Accents | SET Tay Cổ Khuyên Tai Cỏ 4 Lá Bạc S925 Đính Đá Zircon | Đơn Giản Nữ Tính [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "15 Set Đỏ Bạc",
    "type": "hop",
    "name": "ELLA Accents | SET Tay Cổ Khuyên Tai Cỏ 4 Lá Bạc S925 Đính Đá Zircon | Đơn Giản Nữ Tính [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "15 Tay Đỏ Vàng",
    "type": "hop",
    "name": "ELLA Accents | SET Tay Cổ Khuyên Tai Cỏ 4 Lá Bạc S925 Đính Đá Zircon | Đơn Giản Nữ Tính [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "15 Cổ Đỏ Vàng",
    "type": "hop",
    "name": "ELLA Accents | SET Tay Cổ Khuyên Tai Cỏ 4 Lá Bạc S925 Đính Đá Zircon | Đơn Giản Nữ Tính [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "15 Set Đỏ Vàng",
    "type": "hop",
    "name": "ELLA Accents | SET Tay Cổ Khuyên Tai Cỏ 4 Lá Bạc S925 Đính Đá Zircon | Đơn Giản Nữ Tính [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "15 Tay Trắng Bạc",
    "type": "hop",
    "name": "ELLA Accents | SET Tay Cổ Khuyên Tai Cỏ 4 Lá Bạc S925 Đính Đá Zircon | Đơn Giản Nữ Tính [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 498
  },
  {
    "sku": "15 Set Trắng Bạc",
    "type": "hop",
    "name": "ELLA Accents | SET Tay Cổ Khuyên Tai Cỏ 4 Lá Bạc S925 Đính Đá Zircon | Đơn Giản Nữ Tính [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "15 Tay Trắng Vàng",
    "type": "hop",
    "name": "ELLA Accents | SET Tay Cổ Khuyên Tai Cỏ 4 Lá Bạc S925 Đính Đá Zircon | Đơn Giản Nữ Tính [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "15 Cổ Trắng Vàng",
    "type": "hop",
    "name": "ELLA Accents | SET Tay Cổ Khuyên Tai Cỏ 4 Lá Bạc S925 Đính Đá Zircon | Đơn Giản Nữ Tính [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "15 Set Trắng Vàng",
    "type": "hop",
    "name": "ELLA Accents | SET Tay Cổ Khuyên Tai Cỏ 4 Lá Bạc S925 Đính Đá Zircon | Đơn Giản Nữ Tính [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "00 Hop",
    "type": "hop",
    "name": "Hộp kính Premium | ELLA Eyewear Việt Nam",
    "cogs": 0,
    "stockQty": 9987
  },
  {
    "sku": "1640 Bạc 2",
    "type": "hop",
    "name": "ELLA Accents | Vòng Tay Cỏ 4 Lá Đính Đá Zircon Cao Cấp | Thanh Lịch Nữ Tính [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "1650 Vàng 2",
    "type": "hop",
    "name": "ELLA Accents | Vòng Tay Cỏ 4 Lá Đính Đá Zircon Cao Cấp | Thanh Lịch Nữ Tính [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "1641 Bạc 3",
    "type": "hop",
    "name": "ELLA Accents | Vòng Tay Cỏ 4 Lá Đính Đá Zircon Cao Cấp | Thanh Lịch Nữ Tính [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "1651 Vàng 3",
    "type": "hop",
    "name": "ELLA Accents | Vòng Tay Cỏ 4 Lá Đính Đá Zircon Cao Cấp | Thanh Lịch Nữ Tính [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "1642 Bạc 5",
    "type": "hop",
    "name": "ELLA Accents | Vòng Tay Cỏ 4 Lá Đính Đá Zircon Cao Cấp | Thanh Lịch Nữ Tính [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "1642 Vàng 5",
    "type": "hop",
    "name": "ELLA Accents | Vòng Tay Cỏ 4 Lá Đính Đá Zircon Cao Cấp | Thanh Lịch Nữ Tính [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "1740 Bạc",
    "type": "hop",
    "name": "ELLA Accents | Vòng Tay Đính Đá Khóa Tròn | Thanh Lịch Tinh Tế [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "1750 Vàng",
    "type": "hop",
    "name": "ELLA Accents | Vòng Tay Đính Đá Khóa Tròn | Thanh Lịch Tinh Tế [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "21 Bạc",
    "type": "hop",
    "name": "ELLA Accents | Vòng Tay Đính Đá Charm Ngọc Trai Thánh Giá | Sang Trọng Cá Tính [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "21 Vàng",
    "type": "hop",
    "name": "ELLA Accents | Vòng Tay Đính Đá Charm Ngọc Trai Thánh Giá | Sang Trọng Cá Tính [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "22 Vòng cổ",
    "type": "hop",
    "name": "ELLA Accents | Vòng Cổ Trái Tim Mặt Trăng Ngôi Sao Đính Đá | Tinh Tế Cuốn Hút [Kèm Hộp]",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "24 Bạc",
    "type": "khac",
    "name": "ELLA Accents | Dây Chuyền Hai Vòng Tròn Đan Nhau Số La Mã | Tinh Tế Cổ Điển",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "24 Vàng",
    "type": "khac",
    "name": "ELLA Accents | Dây Chuyền Hai Vòng Tròn Đan Nhau Số La Mã | Tinh Tế Cổ Điển",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "23 Kiềng Bạc",
    "type": "khac",
    "name": "ELLA Accents | Set 2 Vòng Tay Số La Mã Đính Đá | Cổ Điển Hiện Đại",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "23 Lắc Bạc",
    "type": "khac",
    "name": "ELLA Accents | Set 2 Vòng Tay Số La Mã Đính Đá | Cổ Điển Hiện Đại",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "23 Set Bạc",
    "type": "khac",
    "name": "ELLA Accents | Set 2 Vòng Tay Số La Mã Đính Đá | Cổ Điển Hiện Đại",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "23 Kiềng Vàng",
    "type": "khac",
    "name": "ELLA Accents | Set 2 Vòng Tay Số La Mã Đính Đá | Cổ Điển Hiện Đại",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "23 Lắc Vàng",
    "type": "khac",
    "name": "ELLA Accents | Set 2 Vòng Tay Số La Mã Đính Đá | Cổ Điển Hiện Đại",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "23 Set Vàng",
    "type": "khac",
    "name": "ELLA Accents | Set 2 Vòng Tay Số La Mã Đính Đá | Cổ Điển Hiện Đại",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "25 Bạc",
    "type": "khac",
    "name": "ELLA Accents | Dây Chuyền Mặt Hình Quạt Đính Đá Zircon | Sang Trọng Nổi Bật",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "25 Vàng",
    "type": "khac",
    "name": "ELLA Accents | Dây Chuyền Mặt Hình Quạt Đính Đá Zircon | Sang Trọng Nổi Bật",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "26 Bạc",
    "type": "khac",
    "name": "ELLA Accents | Vòng Tay Bản Đôi Đính Đá Zircon | Thanh Lịch Cuốn Hút",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "26 Vàng",
    "type": "khac",
    "name": "ELLA Accents | Vòng Tay Bản Đôi Đính Đá Zircon | Thanh Lịch Cuốn Hút",
    "cogs": 0,
    "stockQty": 500
  },
  {
    "sku": "00 Hop Qua",
    "type": "hop",
    "name": "Hộp Quà Đựng Vòng Cổ Vòng Tay Trang Sức | ELLA Eyewear Việt Nam",
    "cogs": 0,
    "stockQty": 9996
  },
  {
    "sku": "EL-Khan",
    "type": "khan",
    "name": "Khăn Lau Kinh Cao Cấp Microfiber | ELLA Eyewear Việt Nam",
    "cogs": 0,
    "stockQty": 569
  },
  {
    "sku": "00 Khan",
    "type": "khan",
    "name": "Khăn Lau Kinh Cao Cấp Microfiber | ELLA Eyewear Việt Nam",
    "cogs": 0,
    "stockQty": 9984
  }
];

const REAL_TIKTOK_ORDERS = [
  {
    "id": "585733419655530088",
    "platform": "tiktok",
    "date": "2026-08-26",
    "status": "success",
    "items": [
      {
        "sellerSku": "EL-1710",
        "subtotal": 225000,
        "cogs": 0,
        "type": "kinh"
      },
      {
        "sellerSku": "00 Khan",
        "subtotal": 0,
        "cogs": 0,
        "type": "khan"
      },
      {
        "sellerSku": "00 Hop",
        "subtotal": 0,
        "cogs": 0,
        "type": "hop"
      }
    ],
    "gmv": 225000,
    "settlementAmount": 74500,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "585725413333042839",
    "platform": "tiktok",
    "date": "2026-08-25",
    "status": "success",
    "items": [
      {
        "sellerSku": "EL-1710",
        "subtotal": 225000,
        "cogs": 0,
        "type": "kinh"
      },
      {
        "sellerSku": "00 Khan",
        "subtotal": 0,
        "cogs": 0,
        "type": "khan"
      },
      {
        "sellerSku": "00 Hop",
        "subtotal": 0,
        "cogs": 0,
        "type": "hop"
      }
    ],
    "gmv": 225000,
    "settlementAmount": 83405,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "585722957185845232",
    "platform": "tiktok",
    "date": "2026-08-25",
    "status": "success",
    "items": [
      {
        "sellerSku": "0999 Bạc Trắng",
        "subtotal": 289000,
        "cogs": 0,
        "type": "hop"
      },
      {
        "sellerSku": "00 Khan",
        "subtotal": 0,
        "cogs": 0,
        "type": "khan"
      },
      {
        "sellerSku": "00 Hop Qua",
        "subtotal": 0,
        "cogs": 0,
        "type": "hop"
      }
    ],
    "gmv": 289000,
    "settlementAmount": 149000,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "585722073937970558",
    "platform": "tiktok",
    "date": "2026-08-25",
    "status": "success",
    "items": [
      {
        "sellerSku": "00 Khan",
        "subtotal": 0,
        "cogs": 0,
        "type": "khan"
      },
      {
        "sellerSku": "00 Hop Qua",
        "subtotal": 0,
        "cogs": 0,
        "type": "hop"
      },
      {
        "sellerSku": "15 Tay Trắng Bạc",
        "subtotal": 378000,
        "cogs": 0,
        "type": "hop"
      }
    ],
    "gmv": 378000,
    "settlementAmount": 303000,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "585720900647946121",
    "platform": "tiktok",
    "date": "2026-08-25",
    "status": "success",
    "items": [
      {
        "sellerSku": "00 Khan",
        "subtotal": 0,
        "cogs": 0,
        "type": "khan"
      },
      {
        "sellerSku": "00 Hop Qua",
        "subtotal": 0,
        "cogs": 0,
        "type": "hop"
      },
      {
        "sellerSku": "0614 Bạc Nơ",
        "subtotal": 299000,
        "cogs": 0,
        "type": "hop"
      }
    ],
    "gmv": 299000,
    "settlementAmount": 183080,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "585714621638804654",
    "platform": "tiktok",
    "date": "2026-08-25",
    "status": "cancelled",
    "items": [
      {
        "sellerSku": "0999 Bạc Trắng",
        "subtotal": 289000,
        "cogs": 0,
        "type": "hop"
      },
      {
        "sellerSku": "00 Khan",
        "subtotal": 0,
        "cogs": 0,
        "type": "khan"
      },
      {
        "sellerSku": "00 Hop Qua",
        "subtotal": 0,
        "cogs": 0,
        "type": "hop"
      }
    ],
    "gmv": 289000,
    "settlementAmount": 149000,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "585710000478193451",
    "platform": "tiktok",
    "date": "2026-08-24",
    "status": "cancelled",
    "items": [
      {
        "sellerSku": "00 Hop",
        "subtotal": 0,
        "cogs": 0,
        "type": "hop"
      },
      {
        "sellerSku": "EL - 0310",
        "subtotal": 228000,
        "cogs": 0,
        "type": "kinh"
      },
      {
        "sellerSku": "00 Khan",
        "subtotal": 0,
        "cogs": 0,
        "type": "khan"
      }
    ],
    "gmv": 228000,
    "settlementAmount": 89000,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "585709998942684971",
    "platform": "tiktok",
    "date": "2026-08-24",
    "status": "cancelled",
    "items": [
      {
        "sellerSku": "00 Hop",
        "subtotal": 0,
        "cogs": 0,
        "type": "hop"
      },
      {
        "sellerSku": "EL-0210",
        "subtotal": 219000,
        "cogs": 0,
        "type": "kinh"
      },
      {
        "sellerSku": "00 Khan",
        "subtotal": 0,
        "cogs": 0,
        "type": "khan"
      }
    ],
    "gmv": 219000,
    "settlementAmount": 84500,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "585708686512064507",
    "platform": "tiktok",
    "date": "2026-08-24",
    "status": "success",
    "items": [
      {
        "sellerSku": "00 Khan",
        "subtotal": 0,
        "cogs": 0,
        "type": "khan"
      },
      {
        "sellerSku": "00 Hop",
        "subtotal": 0,
        "cogs": 0,
        "type": "hop"
      },
      {
        "sellerSku": "EL-1740",
        "subtotal": 225000,
        "cogs": 0,
        "type": "kinh"
      }
    ],
    "gmv": 225000,
    "settlementAmount": 86375,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "585708644237214836",
    "platform": "tiktok",
    "date": "2026-08-24",
    "status": "success",
    "items": [
      {
        "sellerSku": "EL-1732",
        "subtotal": 225000,
        "cogs": 0,
        "type": "kinh"
      },
      {
        "sellerSku": "00 Khan",
        "subtotal": 0,
        "cogs": 0,
        "type": "khan"
      },
      {
        "sellerSku": "00 Hop",
        "subtotal": 0,
        "cogs": 0,
        "type": "hop"
      }
    ],
    "gmv": 225000,
    "settlementAmount": 83405,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "585705399746660251",
    "platform": "tiktok",
    "date": "2026-08-24",
    "status": "success",
    "items": [
      {
        "sellerSku": "00 Hop Qua",
        "subtotal": 54000,
        "cogs": 0,
        "type": "hop"
      }
    ],
    "gmv": 54000,
    "settlementAmount": 29000,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "585691433855386805",
    "platform": "tiktok",
    "date": "2026-08-23",
    "status": "cancelled",
    "items": [
      {
        "sellerSku": "EL - 0310",
        "subtotal": 228000,
        "cogs": 0,
        "type": "kinh"
      },
      {
        "sellerSku": "00 Khan",
        "subtotal": 0,
        "cogs": 0,
        "type": "khan"
      },
      {
        "sellerSku": "00 Hop",
        "subtotal": 0,
        "cogs": 0,
        "type": "hop"
      }
    ],
    "gmv": 228000,
    "settlementAmount": 76540,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "585681459190859756",
    "platform": "tiktok",
    "date": "2026-08-23",
    "status": "success",
    "items": [
      {
        "sellerSku": "EL - 0310",
        "subtotal": 228000,
        "cogs": 0,
        "type": "kinh"
      },
      {
        "sellerSku": "00 Khan",
        "subtotal": 0,
        "cogs": 0,
        "type": "khan"
      },
      {
        "sellerSku": "00 Hop",
        "subtotal": 0,
        "cogs": 0,
        "type": "hop"
      }
    ],
    "gmv": 228000,
    "settlementAmount": 89000,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "585677223626442490",
    "platform": "tiktok",
    "date": "2026-08-22",
    "status": "success",
    "items": [
      {
        "sellerSku": "EL-0210",
        "subtotal": 219000,
        "cogs": 0,
        "type": "kinh"
      },
      {
        "sellerSku": "00 Khan",
        "subtotal": 0,
        "cogs": 0,
        "type": "khan"
      },
      {
        "sellerSku": "00 Hop",
        "subtotal": 0,
        "cogs": 0,
        "type": "hop"
      }
    ],
    "gmv": 219000,
    "settlementAmount": 77740,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "585632693228897836",
    "platform": "tiktok",
    "date": "2026-08-20",
    "status": "success",
    "items": [
      {
        "sellerSku": "00 Hop",
        "subtotal": 0,
        "cogs": 0,
        "type": "hop"
      },
      {
        "sellerSku": "EL-0810",
        "subtotal": 222000,
        "cogs": 0,
        "type": "kinh"
      },
      {
        "sellerSku": "00 Khan",
        "subtotal": 0,
        "cogs": 0,
        "type": "khan"
      }
    ],
    "gmv": 222000,
    "settlementAmount": 56414,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4569,
      "commissionFee": 9216,
      "affiliateCommission": 0,
      "voucherXtra": 4189,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 20974,
    "cogsTotal": 0
  },
  {
    "id": "585632631629973132",
    "platform": "tiktok",
    "date": "2026-08-20",
    "status": "success",
    "items": [
      {
        "sellerSku": "El-0910",
        "subtotal": 224000,
        "cogs": 0,
        "type": "kinh"
      },
      {
        "sellerSku": "00 Khan",
        "subtotal": 0,
        "cogs": 0,
        "type": "khan"
      },
      {
        "sellerSku": "00 Hop",
        "subtotal": 0,
        "cogs": 0,
        "type": "hop"
      }
    ],
    "gmv": 224000,
    "settlementAmount": 70787,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4116,
      "commissionFee": 9054,
      "affiliateCommission": 0,
      "voucherXtra": 4116,
      "orderProcessingFee": 3000,
      "taxWithheld": 3539,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 23825,
    "cogsTotal": 0
  },
  {
    "id": "585629230338836236",
    "platform": "tiktok",
    "date": "2026-08-19",
    "status": "cancelled",
    "items": [
      {
        "sellerSku": "1440 Bạc Hồng",
        "subtotal": 389000,
        "cogs": 0,
        "type": "hop"
      }
    ],
    "gmv": 389000,
    "settlementAmount": 219520,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "585628921378669753",
    "platform": "tiktok",
    "date": "2026-08-19",
    "status": "success",
    "items": [
      {
        "sellerSku": "00 Khan",
        "subtotal": 0,
        "cogs": 0,
        "type": "khan"
      },
      {
        "sellerSku": "00 Hop",
        "subtotal": 0,
        "cogs": 0,
        "type": "hop"
      },
      {
        "sellerSku": "El-0910",
        "subtotal": 224000,
        "cogs": 0,
        "type": "kinh"
      }
    ],
    "gmv": 224000,
    "settlementAmount": 73256,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4939,
      "commissionFee": 9054,
      "affiliateCommission": 0,
      "voucherXtra": 4116,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 21109,
    "cogsTotal": 0
  },
  {
    "id": "585603563024843910",
    "platform": "tiktok",
    "date": "2026-08-18",
    "status": "success",
    "items": [
      {
        "sellerSku": "00 Hop",
        "subtotal": 0,
        "cogs": 0,
        "type": "hop"
      },
      {
        "sellerSku": "EL-0810",
        "subtotal": 222000,
        "cogs": 0,
        "type": "kinh"
      },
      {
        "sellerSku": "00 Khan",
        "subtotal": 0,
        "cogs": 0,
        "type": "khan"
      }
    ],
    "gmv": 222000,
    "settlementAmount": 83780,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 5027,
      "commissionFee": 9216,
      "affiliateCommission": 0,
      "voucherXtra": 4189,
      "orderProcessingFee": 3000,
      "taxWithheld": 4189,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 25621,
    "cogsTotal": 0
  },
  {
    "id": "585599681654654117",
    "platform": "tiktok",
    "date": "2026-08-18",
    "status": "success",
    "items": [
      {
        "sellerSku": "EL-0410",
        "subtotal": 224000,
        "cogs": 0,
        "type": "kinh"
      },
      {
        "sellerSku": "00 Khan",
        "subtotal": 0,
        "cogs": 0,
        "type": "khan"
      },
      {
        "sellerSku": "00 Hop",
        "subtotal": 0,
        "cogs": 0,
        "type": "hop"
      }
    ],
    "gmv": 224000,
    "settlementAmount": 87000,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 5220,
      "commissionFee": 9570,
      "affiliateCommission": 0,
      "voucherXtra": 4350,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22140,
    "cogsTotal": 0
  },
  {
    "id": "585599680981599397",
    "platform": "tiktok",
    "date": "2026-08-18",
    "status": "success",
    "items": [
      {
        "sellerSku": "EL-0861",
        "subtotal": 222000,
        "cogs": 0,
        "type": "kinh"
      },
      {
        "sellerSku": "00 Khan",
        "subtotal": 0,
        "cogs": 0,
        "type": "khan"
      },
      {
        "sellerSku": "00 Hop",
        "subtotal": 0,
        "cogs": 0,
        "type": "hop"
      }
    ],
    "gmv": 222000,
    "settlementAmount": 72051,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 5027,
      "commissionFee": 9216,
      "affiliateCommission": 0,
      "voucherXtra": 4189,
      "orderProcessingFee": 3000,
      "taxWithheld": 3603,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 25035,
    "cogsTotal": 0
  },
  {
    "id": "585597924166239613",
    "platform": "tiktok",
    "date": "2026-08-18",
    "status": "success",
    "items": [
      {
        "sellerSku": "0614 Bạc Nơ",
        "subtotal": 299000,
        "cogs": 0,
        "type": "hop"
      }
    ],
    "gmv": 299000,
    "settlementAmount": 164000,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 10440,
      "commissionFee": 20880,
      "affiliateCommission": 0,
      "voucherXtra": 8700,
      "orderProcessingFee": 3000,
      "taxWithheld": 8200,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 51220,
    "cogsTotal": 0
  },
  {
    "id": "585546603222041732",
    "platform": "tiktok",
    "date": "2026-08-14",
    "status": "success",
    "items": [
      {
        "sellerSku": "EL-1711",
        "subtotal": 225000,
        "cogs": 0,
        "type": "kinh"
      },
      {
        "sellerSku": "EL-2470",
        "subtotal": 289000,
        "cogs": 0,
        "type": "kinh"
      },
      {
        "sellerSku": "EL-Hop",
        "subtotal": 0,
        "cogs": 0,
        "type": "hop"
      },
      {
        "sellerSku": "EL-Khan",
        "subtotal": 0,
        "cogs": 0,
        "type": "khan"
      }
    ],
    "gmv": 514000,
    "settlementAmount": 230875,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 11544,
      "commissionFee": 25396,
      "affiliateCommission": 0,
      "voucherXtra": 11544,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 51484,
    "cogsTotal": 0
  },
  {
    "id": "584967706483263106",
    "platform": "tiktok",
    "date": "2026-07-11",
    "status": "returned",
    "items": [],
    "gmv": 0,
    "settlementAmount": 68440,
    "fees": {
      "shippingFee": 50149,
      "transactionFee": 4700,
      "commissionFee": 10340,
      "affiliateCommission": 0,
      "voucherXtra": 4700,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 72889,
    "cogsTotal": 0,
    "returnReason": "Product doesn't match description"
  },
  {
    "id": "584998579542066372",
    "platform": "tiktok",
    "date": "2026-07-13",
    "status": "returned",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 58192,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 61192,
    "cogsTotal": 0,
    "returnReason": "Poor quality"
  },
  {
    "id": "584999192966236014",
    "platform": "tiktok",
    "date": "2026-07-13",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 57375,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4375,
      "commissionFee": 9625,
      "affiliateCommission": 0,
      "voucherXtra": 4375,
      "orderProcessingFee": 3000,
      "taxWithheld": 6125,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 27500,
    "cogsTotal": 0
  },
  {
    "id": "584922997042938896",
    "platform": "tiktok",
    "date": "2026-07-08",
    "status": "returned",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 20032,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 23032,
    "cogsTotal": 0,
    "returnReason": "Wrong product sent"
  },
  {
    "id": "584916807743472717",
    "platform": "tiktok",
    "date": "2026-07-08",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 15800,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 15800,
    "cogsTotal": 0
  },
  {
    "id": "585013012634436752",
    "platform": "tiktok",
    "date": "2026-07-14",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 55384,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4170,
      "commissionFee": 9175,
      "affiliateCommission": 0,
      "voucherXtra": 4170,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 20515,
    "cogsTotal": 0
  },
  {
    "id": "585011350058009684",
    "platform": "tiktok",
    "date": "2026-07-13",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 52975,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 6150,
      "commissionFee": 9625,
      "affiliateCommission": 0,
      "voucherXtra": 4375,
      "orderProcessingFee": 3000,
      "taxWithheld": 8750,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 31900,
    "cogsTotal": 0
  },
  {
    "id": "585009712164669322",
    "platform": "tiktok",
    "date": "2026-07-13",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 56166,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4319,
      "commissionFee": 9501,
      "affiliateCommission": 0,
      "voucherXtra": 4319,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 21139,
    "cogsTotal": 0
  },
  {
    "id": "585007888246933168",
    "platform": "tiktok",
    "date": "2026-07-13",
    "status": "returned",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 3000,
    "cogsTotal": 0,
    "returnReason": "Change of mind"
  },
  {
    "id": "584993810672748047",
    "platform": "tiktok",
    "date": "2026-07-12",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 67290,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4950,
      "commissionFee": 10890,
      "affiliateCommission": 0,
      "voucherXtra": 4950,
      "orderProcessingFee": 3000,
      "taxWithheld": 4950,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 28740,
    "cogsTotal": 0
  },
  {
    "id": "584971541907605432",
    "platform": "tiktok",
    "date": "2026-07-11",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 72240,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4950,
      "commissionFee": 10890,
      "affiliateCommission": 0,
      "voucherXtra": 4950,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 23790,
    "cogsTotal": 0
  },
  {
    "id": "585016345254528906",
    "platform": "tiktok",
    "date": "2026-07-14",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 56150,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 9295,
      "affiliateCommission": 0,
      "voucherXtra": 4225,
      "orderProcessingFee": 3000,
      "taxWithheld": 5070,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 25815,
    "cogsTotal": 0
  },
  {
    "id": "585014643423741355",
    "platform": "tiktok",
    "date": "2026-07-14",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 55384,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4170,
      "commissionFee": 9175,
      "affiliateCommission": 0,
      "voucherXtra": 4170,
      "orderProcessingFee": 3000,
      "taxWithheld": 5004,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 25519,
    "cogsTotal": 0
  },
  {
    "id": "585006460688696421",
    "platform": "tiktok",
    "date": "2026-07-13",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 60388,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4170,
      "commissionFee": 9175,
      "affiliateCommission": 0,
      "voucherXtra": 4170,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 20515,
    "cogsTotal": 0
  },
  {
    "id": "584985155664512529",
    "platform": "tiktok",
    "date": "2026-07-12",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 52770,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 9295,
      "affiliateCommission": 0,
      "voucherXtra": 4225,
      "orderProcessingFee": 3000,
      "taxWithheld": 8450,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 29195,
    "cogsTotal": 0
  },
  {
    "id": "585014576632661920",
    "platform": "tiktok",
    "date": "2026-07-14",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 78370,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 5975,
      "commissionFee": 13145,
      "affiliateCommission": 0,
      "voucherXtra": 5975,
      "orderProcessingFee": 3000,
      "taxWithheld": 9450,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 37545,
    "cogsTotal": 0
  },
  {
    "id": "585011639365306351",
    "platform": "tiktok",
    "date": "2026-07-13",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 56218,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4170,
      "commissionFee": 9175,
      "affiliateCommission": 0,
      "voucherXtra": 4170,
      "orderProcessingFee": 3000,
      "taxWithheld": 4170,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 24685,
    "cogsTotal": 0
  },
  {
    "id": "585011391783339428",
    "platform": "tiktok",
    "date": "2026-07-13",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 60388,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4170,
      "commissionFee": 9175,
      "affiliateCommission": 0,
      "voucherXtra": 4170,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 20515,
    "cogsTotal": 0
  },
  {
    "id": "585004930905703638",
    "platform": "tiktok",
    "date": "2026-07-13",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 53029,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4244,
      "commissionFee": 9338,
      "affiliateCommission": 0,
      "voucherXtra": 4244,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 20826,
    "cogsTotal": 0
  },
  {
    "id": "585002228942144934",
    "platform": "tiktok",
    "date": "2026-07-13",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 60388,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4170,
      "commissionFee": 9175,
      "affiliateCommission": 0,
      "voucherXtra": 4170,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 20515,
    "cogsTotal": 0
  },
  {
    "id": "584998547703891779",
    "platform": "tiktok",
    "date": "2026-07-13",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 57807,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4319,
      "commissionFee": 9501,
      "affiliateCommission": 0,
      "voucherXtra": 4319,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 21139,
    "cogsTotal": 0
  },
  {
    "id": "584996494703560389",
    "platform": "tiktok",
    "date": "2026-07-13",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 56150,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 9295,
      "affiliateCommission": 0,
      "voucherXtra": 4225,
      "orderProcessingFee": 3000,
      "taxWithheld": 5070,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 25815,
    "cogsTotal": 0
  },
  {
    "id": "584996262720144736",
    "platform": "tiktok",
    "date": "2026-07-13",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 113901,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 8058,
      "commissionFee": 17726,
      "affiliateCommission": 0,
      "voucherXtra": 8058,
      "orderProcessingFee": 3000,
      "taxWithheld": 5572,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 42414,
    "cogsTotal": 0
  },
  {
    "id": "584996144296592870",
    "platform": "tiktok",
    "date": "2026-07-13",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 57375,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4375,
      "commissionFee": 9625,
      "affiliateCommission": 0,
      "voucherXtra": 4375,
      "orderProcessingFee": 3000,
      "taxWithheld": 6125,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 27500,
    "cogsTotal": 0
  },
  {
    "id": "584987155868911044",
    "platform": "tiktok",
    "date": "2026-07-12",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 72240,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4950,
      "commissionFee": 10890,
      "affiliateCommission": 0,
      "voucherXtra": 4950,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 23790,
    "cogsTotal": 0
  },
  {
    "id": "584974363472463097",
    "platform": "tiktok",
    "date": "2026-07-11",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 59379,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4350,
      "commissionFee": 9570,
      "affiliateCommission": 0,
      "voucherXtra": 4350,
      "orderProcessingFee": 3000,
      "taxWithheld": 3741,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 25011,
    "cogsTotal": 0
  },
  {
    "id": "584968474065602497",
    "platform": "tiktok",
    "date": "2026-07-11",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 72240,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4950,
      "commissionFee": 10890,
      "affiliateCommission": 0,
      "voucherXtra": 4950,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 23790,
    "cogsTotal": 0
  },
  {
    "id": "584806249539011757",
    "platform": "tiktok",
    "date": "2026-07-01",
    "status": "returned",
    "items": [],
    "gmv": 0,
    "settlementAmount": 52470,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 8600,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 30950,
    "cogsTotal": 0,
    "returnReason": "Change of mind"
  },
  {
    "id": "584946710787687709",
    "platform": "tiktok",
    "date": "2026-07-10",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 58250,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4375,
      "commissionFee": 9625,
      "affiliateCommission": 0,
      "voucherXtra": 4375,
      "orderProcessingFee": 3000,
      "taxWithheld": 5250,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 26625,
    "cogsTotal": 0
  },
  {
    "id": "584946197678097853",
    "platform": "tiktok",
    "date": "2026-07-09",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 58670,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4375,
      "commissionFee": 9625,
      "affiliateCommission": 0,
      "voucherXtra": 4375,
      "orderProcessingFee": 3000,
      "taxWithheld": 4830,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 26205,
    "cogsTotal": 0
  },
  {
    "id": "584925610408707279",
    "platform": "tiktok",
    "date": "2026-07-08",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 62360,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 9460,
      "affiliateCommission": 0,
      "voucherXtra": 4300,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 21060,
    "cogsTotal": 0
  },
  {
    "id": "585004800206013755",
    "platform": "tiktok",
    "date": "2026-07-13",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584951859530794465",
    "platform": "tiktok",
    "date": "2026-07-10",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 63120,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4350,
      "commissionFee": 9570,
      "affiliateCommission": 0,
      "voucherXtra": 4350,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 21270,
    "cogsTotal": 0
  },
  {
    "id": "584946964702463892",
    "platform": "tiktok",
    "date": "2026-07-10",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 72240,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4950,
      "commissionFee": 10890,
      "affiliateCommission": 0,
      "voucherXtra": 4950,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 23790,
    "cogsTotal": 0
  },
  {
    "id": "584946921069512617",
    "platform": "tiktok",
    "date": "2026-07-10",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 57250,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4375,
      "commissionFee": 9625,
      "affiliateCommission": 0,
      "voucherXtra": 4375,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 21375,
    "cogsTotal": 0
  },
  {
    "id": "584946675303221014",
    "platform": "tiktok",
    "date": "2026-07-10",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 58670,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4375,
      "commissionFee": 9625,
      "affiliateCommission": 0,
      "voucherXtra": 4375,
      "orderProcessingFee": 3000,
      "taxWithheld": 4830,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 26205,
    "cogsTotal": 0
  },
  {
    "id": "584927454558455455",
    "platform": "tiktok",
    "date": "2026-07-08",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 53760,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 9460,
      "affiliateCommission": 0,
      "voucherXtra": 4300,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 21060,
    "cogsTotal": 0
  },
  {
    "id": "585002392841848766",
    "platform": "tiktok",
    "date": "2026-07-13",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584997977733039939",
    "platform": "tiktok",
    "date": "2026-07-13",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584985737076900991",
    "platform": "tiktok",
    "date": "2026-07-12",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584962975275517049",
    "platform": "tiktok",
    "date": "2026-07-10",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584954109588964929",
    "platform": "tiktok",
    "date": "2026-07-10",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 59118,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4350,
      "commissionFee": 9570,
      "affiliateCommission": 0,
      "voucherXtra": 4350,
      "orderProcessingFee": 3000,
      "taxWithheld": 4002,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 25272,
    "cogsTotal": 0
  },
  {
    "id": "584946932066518030",
    "platform": "tiktok",
    "date": "2026-07-10",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 59037,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4375,
      "commissionFee": 9625,
      "affiliateCommission": 0,
      "voucherXtra": 4375,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 21375,
    "cogsTotal": 0
  },
  {
    "id": "584946621956589241",
    "platform": "tiktok",
    "date": "2026-07-10",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 54750,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4375,
      "commissionFee": 9625,
      "affiliateCommission": 0,
      "voucherXtra": 4375,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 21375,
    "cogsTotal": 0
  },
  {
    "id": "584945575417055079",
    "platform": "tiktok",
    "date": "2026-07-09",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 57784,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4375,
      "commissionFee": 9625,
      "affiliateCommission": 0,
      "voucherXtra": 4375,
      "orderProcessingFee": 3000,
      "taxWithheld": 5716,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 27091,
    "cogsTotal": 0
  },
  {
    "id": "584924881988716291",
    "platform": "tiktok",
    "date": "2026-07-08",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 62360,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 9460,
      "affiliateCommission": 0,
      "voucherXtra": 4300,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 21060,
    "cogsTotal": 0
  },
  {
    "id": "584929434496828530",
    "platform": "tiktok",
    "date": "2026-07-08",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 60105,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 9460,
      "affiliateCommission": 0,
      "voucherXtra": 4300,
      "orderProcessingFee": 3000,
      "taxWithheld": 2255,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 23315,
    "cogsTotal": 0
  },
  {
    "id": "584921792390334067",
    "platform": "tiktok",
    "date": "2026-07-08",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 58060,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 9460,
      "affiliateCommission": 0,
      "voucherXtra": 4300,
      "orderProcessingFee": 3000,
      "taxWithheld": 4300,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 25360,
    "cogsTotal": 0
  },
  {
    "id": "584920829624157692",
    "platform": "tiktok",
    "date": "2026-07-08",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 176076,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 11781,
      "commissionFee": 25919,
      "affiliateCommission": 0,
      "voucherXtra": 11781,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 52481,
    "cogsTotal": 0
  },
  {
    "id": "584915514765182248",
    "platform": "tiktok",
    "date": "2026-07-08",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 78065,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 7225,
      "commissionFee": 11990,
      "affiliateCommission": 0,
      "voucherXtra": 5450,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 27665,
    "cogsTotal": 0
  },
  {
    "id": "584901689863014387",
    "platform": "tiktok",
    "date": "2026-07-07",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 71640,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 5550,
      "commissionFee": 10890,
      "affiliateCommission": 0,
      "voucherXtra": 4950,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 24390,
    "cogsTotal": 0
  },
  {
    "id": "584899501241960395",
    "platform": "tiktok",
    "date": "2026-07-07",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 63500,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4375,
      "commissionFee": 9625,
      "affiliateCommission": 0,
      "voucherXtra": 4375,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 21375,
    "cogsTotal": 0
  },
  {
    "id": "584892818053301778",
    "platform": "tiktok",
    "date": "2026-07-06",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 56995,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 9295,
      "affiliateCommission": 0,
      "voucherXtra": 4225,
      "orderProcessingFee": 3000,
      "taxWithheld": 4225,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 24970,
    "cogsTotal": 0
  },
  {
    "id": "584887885002540089",
    "platform": "tiktok",
    "date": "2026-07-06",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 57725,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4375,
      "commissionFee": 9625,
      "affiliateCommission": 0,
      "voucherXtra": 4375,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 21375,
    "cogsTotal": 0
  },
  {
    "id": "584915342472152080",
    "platform": "tiktok",
    "date": "2026-07-08",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 62360,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 9460,
      "affiliateCommission": 0,
      "voucherXtra": 4300,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 21060,
    "cogsTotal": 0
  },
  {
    "id": "584909737740699120",
    "platform": "tiktok",
    "date": "2026-07-07",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 79840,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 5450,
      "commissionFee": 11990,
      "affiliateCommission": 0,
      "voucherXtra": 5450,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 25890,
    "cogsTotal": 0
  },
  {
    "id": "584905373455582775",
    "platform": "tiktok",
    "date": "2026-07-07",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 79840,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 5450,
      "commissionFee": 11990,
      "affiliateCommission": 0,
      "voucherXtra": 5450,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 25890,
    "cogsTotal": 0
  },
  {
    "id": "584881822465623528",
    "platform": "tiktok",
    "date": "2026-07-06",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 54620,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 9460,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 8600,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 28800,
    "cogsTotal": 0
  },
  {
    "id": "584869361397696162",
    "platform": "tiktok",
    "date": "2026-07-05",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 69380,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4700,
      "commissionFee": 10340,
      "affiliateCommission": 0,
      "voucherXtra": 3760,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 21800,
    "cogsTotal": 0
  },
  {
    "id": "584895097893652175",
    "platform": "tiktok",
    "date": "2026-07-07",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 59125,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4375,
      "commissionFee": 9625,
      "affiliateCommission": 0,
      "voucherXtra": 4375,
      "orderProcessingFee": 3000,
      "taxWithheld": 4375,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 25750,
    "cogsTotal": 0
  },
  {
    "id": "584892382934434939",
    "platform": "tiktok",
    "date": "2026-07-06",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61220,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 9295,
      "affiliateCommission": 0,
      "voucherXtra": 4225,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 20745,
    "cogsTotal": 0
  },
  {
    "id": "584878962775721490",
    "platform": "tiktok",
    "date": "2026-07-06",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 58920,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 9460,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 4300,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 24500,
    "cogsTotal": 0
  },
  {
    "id": "584868111187019265",
    "platform": "tiktok",
    "date": "2026-07-05",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 73230,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4950,
      "commissionFee": 10890,
      "affiliateCommission": 0,
      "voucherXtra": 3960,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22800,
    "cogsTotal": 0
  },
  {
    "id": "584946993963369966",
    "platform": "tiktok",
    "date": "2026-07-10",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584946646200648806",
    "platform": "tiktok",
    "date": "2026-07-10",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584946050873984445",
    "platform": "tiktok",
    "date": "2026-07-09",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584939788820514352",
    "platform": "tiktok",
    "date": "2026-07-09",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584870555856963006",
    "platform": "tiktok",
    "date": "2026-07-05",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 63220,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 9460,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 20200,
    "cogsTotal": 0
  },
  {
    "id": "584860012135351851",
    "platform": "tiktok",
    "date": "2026-07-05",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 62065,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 9295,
      "affiliateCommission": 0,
      "voucherXtra": 3380,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 19900,
    "cogsTotal": 0
  },
  {
    "id": "584858509342770385",
    "platform": "tiktok",
    "date": "2026-07-04",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 60184,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 9460,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 3036,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 23236,
    "cogsTotal": 0
  },
  {
    "id": "584845313450476714",
    "platform": "tiktok",
    "date": "2026-07-04",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 58858,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4204,
      "commissionFee": 9625,
      "affiliateCommission": 0,
      "voucherXtra": 3500,
      "orderProcessingFee": 3000,
      "taxWithheld": 5688,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 26017,
    "cogsTotal": 0
  },
  {
    "id": "584836516501554510",
    "platform": "tiktok",
    "date": "2026-07-03",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 63825,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4950,
      "commissionFee": 10890,
      "affiliateCommission": 0,
      "voucherXtra": 3960,
      "orderProcessingFee": 3000,
      "taxWithheld": 9405,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 32205,
    "cogsTotal": 0
  },
  {
    "id": "584925601280459983",
    "platform": "tiktok",
    "date": "2026-07-08",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584920739073852924",
    "platform": "tiktok",
    "date": "2026-07-08",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584920732532966908",
    "platform": "tiktok",
    "date": "2026-07-08",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584850074054723010",
    "platform": "tiktok",
    "date": "2026-07-04",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 89015,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 5975,
      "commissionFee": 13145,
      "affiliateCommission": 0,
      "voucherXtra": 4780,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 26900,
    "cogsTotal": 0
  },
  {
    "id": "584829459118392634",
    "platform": "tiktok",
    "date": "2026-07-03",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 62065,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 9295,
      "affiliateCommission": 0,
      "voucherXtra": 3380,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 19900,
    "cogsTotal": 0
  },
  {
    "id": "584825430178170740",
    "platform": "tiktok",
    "date": "2026-07-03",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 73230,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4950,
      "commissionFee": 10890,
      "affiliateCommission": 0,
      "voucherXtra": 3960,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22800,
    "cogsTotal": 0
  },
  {
    "id": "584824686039369075",
    "platform": "tiktok",
    "date": "2026-07-02",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 59580,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4200,
      "commissionFee": 11340,
      "affiliateCommission": 0,
      "voucherXtra": 3360,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 21900,
    "cogsTotal": 0
  },
  {
    "id": "584822457121277075",
    "platform": "tiktok",
    "date": "2026-07-02",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61070,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584818353666295410",
    "platform": "tiktok",
    "date": "2026-07-02",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61070,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584767225431032883",
    "platform": "tiktok",
    "date": "2026-06-29",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 19700,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 19700,
    "cogsTotal": 0
  },
  {
    "id": "584900722709006303",
    "platform": "tiktok",
    "date": "2026-07-07",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584846423586276794",
    "platform": "tiktok",
    "date": "2026-07-04",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 68280,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4950,
      "commissionFee": 10890,
      "affiliateCommission": 0,
      "voucherXtra": 3960,
      "orderProcessingFee": 3000,
      "taxWithheld": 4950,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 27750,
    "cogsTotal": 0
  },
  {
    "id": "584834245535631183",
    "platform": "tiktok",
    "date": "2026-07-03",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 63175,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4345,
      "commissionFee": 9460,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 20245,
    "cogsTotal": 0
  },
  {
    "id": "584832443203225332",
    "platform": "tiktok",
    "date": "2026-07-03",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 73230,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4950,
      "commissionFee": 10890,
      "affiliateCommission": 0,
      "voucherXtra": 3960,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22800,
    "cogsTotal": 0
  },
  {
    "id": "584832325480645839",
    "platform": "tiktok",
    "date": "2026-07-03",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 56770,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 9460,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 20200,
    "cogsTotal": 0
  },
  {
    "id": "584756346729235605",
    "platform": "tiktok",
    "date": "2026-06-28",
    "status": "returned",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 3000,
    "cogsTotal": 0,
    "returnReason": "Change of mind"
  },
  {
    "id": "584748339570771693",
    "platform": "tiktok",
    "date": "2026-06-28",
    "status": "returned",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 3000,
    "cogsTotal": 0,
    "returnReason": "Item too big or too small"
  },
  {
    "id": "584627477538440854",
    "platform": "tiktok",
    "date": "2026-06-21",
    "status": "returned",
    "items": [],
    "gmv": 0,
    "settlementAmount": 134471,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 9226,
      "commissionFee": 24911,
      "affiliateCommission": 0,
      "voucherXtra": 7381,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 44518,
    "cogsTotal": 0,
    "returnReason": "Change of mind"
  },
  {
    "id": "584886216934786535",
    "platform": "tiktok",
    "date": "2026-07-06",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584822116592616523",
    "platform": "tiktok",
    "date": "2026-07-02",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 54470,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 6600,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 28950,
    "cogsTotal": 0
  },
  {
    "id": "584810153184887905",
    "platform": "tiktok",
    "date": "2026-07-02",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 58163,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4375,
      "commissionFee": 11812,
      "affiliateCommission": 0,
      "voucherXtra": 3500,
      "orderProcessingFee": 3000,
      "taxWithheld": 4025,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 26712,
    "cogsTotal": 0
  },
  {
    "id": "584808409268586304",
    "platform": "tiktok",
    "date": "2026-07-02",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 56150,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4375,
      "commissionFee": 11812,
      "affiliateCommission": 0,
      "voucherXtra": 3500,
      "orderProcessingFee": 3000,
      "taxWithheld": 6038,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 28725,
    "cogsTotal": 0
  },
  {
    "id": "584805557850440995",
    "platform": "tiktok",
    "date": "2026-07-01",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 54405,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4450,
      "commissionFee": 12015,
      "affiliateCommission": 0,
      "voucherXtra": 3560,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 23025,
    "cogsTotal": 0
  },
  {
    "id": "584805530612302904",
    "platform": "tiktok",
    "date": "2026-07-01",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 52470,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 8600,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 30950,
    "cogsTotal": 0
  },
  {
    "id": "584797452944181169",
    "platform": "tiktok",
    "date": "2026-07-01",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 56770,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 4300,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 26650,
    "cogsTotal": 0
  },
  {
    "id": "584796809576350782",
    "platform": "tiktok",
    "date": "2026-07-01",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 62188,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4375,
      "commissionFee": 11812,
      "affiliateCommission": 0,
      "voucherXtra": 3500,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22687,
    "cogsTotal": 0
  },
  {
    "id": "584715464658028201",
    "platform": "tiktok",
    "date": "2026-06-26",
    "status": "returned",
    "items": [],
    "gmv": 0,
    "settlementAmount": 57725,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4375,
      "commissionFee": 11812,
      "affiliateCommission": 0,
      "voucherXtra": 3500,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22687,
    "cogsTotal": 0,
    "returnReason": "Change of mind"
  },
  {
    "id": "584813717067891850",
    "platform": "tiktok",
    "date": "2026-07-02",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 52685,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 11408,
      "affiliateCommission": 0,
      "voucherXtra": 3380,
      "orderProcessingFee": 3000,
      "taxWithheld": 7267,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 29280,
    "cogsTotal": 0
  },
  {
    "id": "584813080424712071",
    "platform": "tiktok",
    "date": "2026-07-02",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61070,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584804980527367951",
    "platform": "tiktok",
    "date": "2026-07-01",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 70755,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4950,
      "commissionFee": 13365,
      "affiliateCommission": 0,
      "voucherXtra": 3960,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 25275,
    "cogsTotal": 0
  },
  {
    "id": "584800650607953406",
    "platform": "tiktok",
    "date": "2026-07-01",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61070,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584701120789514163",
    "platform": "tiktok",
    "date": "2026-06-25",
    "status": "returned",
    "items": [],
    "gmv": 0,
    "settlementAmount": 63305,
    "fees": {
      "shippingFee": 14576,
      "transactionFee": 4450,
      "commissionFee": 12015,
      "affiliateCommission": 0,
      "voucherXtra": 3560,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 37601,
    "cogsTotal": 0,
    "returnReason": "Wrong product sent"
  },
  {
    "id": "584797498598131018",
    "platform": "tiktok",
    "date": "2026-07-01",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 118957,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 8185,
      "commissionFee": 22100,
      "affiliateCommission": 0,
      "voucherXtra": 6548,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 39833,
    "cogsTotal": 0
  },
  {
    "id": "584794600346846566",
    "platform": "tiktok",
    "date": "2026-07-01",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 54405,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4450,
      "commissionFee": 12015,
      "affiliateCommission": 0,
      "voucherXtra": 3560,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 23025,
    "cogsTotal": 0
  },
  {
    "id": "584779455904056947",
    "platform": "tiktok",
    "date": "2026-06-30",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61070,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584766548902970958",
    "platform": "tiktok",
    "date": "2026-06-29",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 78205,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 5450,
      "commissionFee": 14715,
      "affiliateCommission": 0,
      "voucherXtra": 4360,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 27525,
    "cogsTotal": 0
  },
  {
    "id": "584762641806099502",
    "platform": "tiktok",
    "date": "2026-06-29",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 59952,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 11408,
      "affiliateCommission": 0,
      "voucherXtra": 3380,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22013,
    "cogsTotal": 0
  },
  {
    "id": "584759965939107171",
    "platform": "tiktok",
    "date": "2026-06-29",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61070,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584759520965985370",
    "platform": "tiktok",
    "date": "2026-06-29",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61070,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584750795781670602",
    "platform": "tiktok",
    "date": "2026-06-28",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 70755,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4950,
      "commissionFee": 13365,
      "affiliateCommission": 0,
      "voucherXtra": 3960,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 25275,
    "cogsTotal": 0
  },
  {
    "id": "584744560699343995",
    "platform": "tiktok",
    "date": "2026-06-28",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 62188,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4375,
      "commissionFee": 11812,
      "affiliateCommission": 0,
      "voucherXtra": 3500,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22687,
    "cogsTotal": 0
  },
  {
    "id": "584840252997469555",
    "platform": "tiktok",
    "date": "2026-07-03",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584778828100503199",
    "platform": "tiktok",
    "date": "2026-06-30",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61070,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584777278961649561",
    "platform": "tiktok",
    "date": "2026-06-30",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61070,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584770130963957484",
    "platform": "tiktok",
    "date": "2026-06-29",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 59580,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4200,
      "commissionFee": 11340,
      "affiliateCommission": 0,
      "voucherXtra": 3360,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 21900,
    "cogsTotal": 0
  },
  {
    "id": "584766894034224675",
    "platform": "tiktok",
    "date": "2026-06-29",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61070,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584765076407223994",
    "platform": "tiktok",
    "date": "2026-06-29",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 52470,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 8600,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 30950,
    "cogsTotal": 0
  },
  {
    "id": "584762955278944202",
    "platform": "tiktok",
    "date": "2026-06-29",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61070,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584762042742375545",
    "platform": "tiktok",
    "date": "2026-06-29",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 52470,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 8600,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 30950,
    "cogsTotal": 0
  },
  {
    "id": "584736375046571105",
    "platform": "tiktok",
    "date": "2026-06-27",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 58163,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4375,
      "commissionFee": 11812,
      "affiliateCommission": 0,
      "voucherXtra": 3500,
      "orderProcessingFee": 3000,
      "taxWithheld": 4025,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 26712,
    "cogsTotal": 0
  },
  {
    "id": "584821452965381985",
    "platform": "tiktok",
    "date": "2026-07-02",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584820610594669643",
    "platform": "tiktok",
    "date": "2026-07-02",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584806414934246538",
    "platform": "tiktok",
    "date": "2026-07-01",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584761165995869884",
    "platform": "tiktok",
    "date": "2026-06-29",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61070,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584756745378039654",
    "platform": "tiktok",
    "date": "2026-06-28",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61070,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584755444707918863",
    "platform": "tiktok",
    "date": "2026-06-28",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 70755,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4950,
      "commissionFee": 13365,
      "affiliateCommission": 0,
      "voucherXtra": 3960,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 25275,
    "cogsTotal": 0
  },
  {
    "id": "584745946992444873",
    "platform": "tiktok",
    "date": "2026-06-28",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 62188,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4375,
      "commissionFee": 11812,
      "affiliateCommission": 0,
      "voucherXtra": 3500,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22687,
    "cogsTotal": 0
  },
  {
    "id": "584740836458661809",
    "platform": "tiktok",
    "date": "2026-06-28",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 62188,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4375,
      "commissionFee": 11812,
      "affiliateCommission": 0,
      "voucherXtra": 3500,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22687,
    "cogsTotal": 0
  },
  {
    "id": "584735154905777344",
    "platform": "tiktok",
    "date": "2026-06-27",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 57813,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4375,
      "commissionFee": 11812,
      "affiliateCommission": 0,
      "voucherXtra": 3500,
      "orderProcessingFee": 3000,
      "taxWithheld": 4375,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 27062,
    "cogsTotal": 0
  },
  {
    "id": "584727720448853625",
    "platform": "tiktok",
    "date": "2026-06-27",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 63025,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4750,
      "commissionFee": 12825,
      "affiliateCommission": 0,
      "voucherXtra": 3800,
      "orderProcessingFee": 3000,
      "taxWithheld": 4750,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 29125,
    "cogsTotal": 0
  },
  {
    "id": "584800391290520827",
    "platform": "tiktok",
    "date": "2026-07-01",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584791347939345506",
    "platform": "tiktok",
    "date": "2026-07-01",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584787335825228862",
    "platform": "tiktok",
    "date": "2026-06-30",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584731729214342556",
    "platform": "tiktok",
    "date": "2026-06-27",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 62188,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4375,
      "commissionFee": 11812,
      "affiliateCommission": 0,
      "voucherXtra": 3500,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22687,
    "cogsTotal": 0
  },
  {
    "id": "584782677643855482",
    "platform": "tiktok",
    "date": "2026-06-30",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584781133403424334",
    "platform": "tiktok",
    "date": "2026-06-30",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584771773667443953",
    "platform": "tiktok",
    "date": "2026-06-29",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584733554987402243",
    "platform": "tiktok",
    "date": "2026-06-27",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 56063,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4375,
      "commissionFee": 11812,
      "affiliateCommission": 0,
      "voucherXtra": 3500,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22687,
    "cogsTotal": 0
  },
  {
    "id": "584733237862565321",
    "platform": "tiktok",
    "date": "2026-06-27",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61070,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584723622913672316",
    "platform": "tiktok",
    "date": "2026-06-26",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 54970,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584722610278007809",
    "platform": "tiktok",
    "date": "2026-06-26",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 58163,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4375,
      "commissionFee": 11812,
      "affiliateCommission": 0,
      "voucherXtra": 3500,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22687,
    "cogsTotal": 0
  },
  {
    "id": "584682266523960447",
    "platform": "tiktok",
    "date": "2026-06-24",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 56900,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4375,
      "commissionFee": 11812,
      "affiliateCommission": 0,
      "voucherXtra": 3500,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22687,
    "cogsTotal": 0
  },
  {
    "id": "584680307884787228",
    "platform": "tiktok",
    "date": "2026-06-24",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 57630,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4700,
      "commissionFee": 12690,
      "affiliateCommission": 0,
      "voucherXtra": 3760,
      "orderProcessingFee": 3000,
      "taxWithheld": 9400,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 33550,
    "cogsTotal": 0
  },
  {
    "id": "584650654230546165",
    "platform": "tiktok",
    "date": "2026-06-22",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 15800,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 15800,
    "cogsTotal": 0
  },
  {
    "id": "584707302568592827",
    "platform": "tiktok",
    "date": "2026-06-25",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 54470,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 6600,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 28950,
    "cogsTotal": 0
  },
  {
    "id": "584707035286832512",
    "platform": "tiktok",
    "date": "2026-06-25",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61070,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584706876272772636",
    "platform": "tiktok",
    "date": "2026-06-25",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61070,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584697787102627546",
    "platform": "tiktok",
    "date": "2026-06-25",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 66805,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4950,
      "commissionFee": 13365,
      "affiliateCommission": 0,
      "voucherXtra": 3960,
      "orderProcessingFee": 3000,
      "taxWithheld": 3950,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 29225,
    "cogsTotal": 0
  },
  {
    "id": "584697471900419929",
    "platform": "tiktok",
    "date": "2026-06-25",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61070,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584684873760343583",
    "platform": "tiktok",
    "date": "2026-06-24",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 53438,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4375,
      "commissionFee": 11812,
      "affiliateCommission": 0,
      "voucherXtra": 3500,
      "orderProcessingFee": 3000,
      "taxWithheld": 8750,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 31437,
    "cogsTotal": 0
  },
  {
    "id": "584684516392731803",
    "platform": "tiktok",
    "date": "2026-06-24",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 56770,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 4300,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 26650,
    "cogsTotal": 0
  },
  {
    "id": "584679585204504097",
    "platform": "tiktok",
    "date": "2026-06-24",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 58855,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4450,
      "commissionFee": 12015,
      "affiliateCommission": 0,
      "voucherXtra": 3560,
      "orderProcessingFee": 3000,
      "taxWithheld": 4450,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 27475,
    "cogsTotal": 0
  },
  {
    "id": "584673476786357639",
    "platform": "tiktok",
    "date": "2026-06-23",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61070,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584644084026607333",
    "platform": "tiktok",
    "date": "2026-06-22",
    "status": "returned",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 3000,
    "cogsTotal": 0,
    "returnReason": "Item too big or too small"
  },
  {
    "id": "584697818660243150",
    "platform": "tiktok",
    "date": "2026-06-25",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61070,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584694742001551288",
    "platform": "tiktok",
    "date": "2026-06-25",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61070,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584686952078083519",
    "platform": "tiktok",
    "date": "2026-06-24",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61070,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584686587201947435",
    "platform": "tiktok",
    "date": "2026-06-24",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 52470,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584682678938535004",
    "platform": "tiktok",
    "date": "2026-06-24",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 53760,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584676361080112388",
    "platform": "tiktok",
    "date": "2026-06-24",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 81547,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 5975,
      "commissionFee": 16132,
      "affiliateCommission": 0,
      "voucherXtra": 4780,
      "orderProcessingFee": 3000,
      "taxWithheld": 4481,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 34368,
    "cogsTotal": 0
  },
  {
    "id": "584672987535476707",
    "platform": "tiktok",
    "date": "2026-06-23",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61070,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584671708236318332",
    "platform": "tiktok",
    "date": "2026-06-23",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61070,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584671048687781436",
    "platform": "tiktok",
    "date": "2026-06-23",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 111363,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 8228,
      "commissionFee": 22214,
      "affiliateCommission": 0,
      "voucherXtra": 6582,
      "orderProcessingFee": 3000,
      "taxWithheld": 8227,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 48251,
    "cogsTotal": 0
  },
  {
    "id": "584664039912081085",
    "platform": "tiktok",
    "date": "2026-06-23",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 80053,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 5975,
      "commissionFee": 16132,
      "affiliateCommission": 0,
      "voucherXtra": 4780,
      "orderProcessingFee": 3000,
      "taxWithheld": 5975,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 35862,
    "cogsTotal": 0
  },
  {
    "id": "584655163152107399",
    "platform": "tiktok",
    "date": "2026-06-22",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 56872,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4375,
      "commissionFee": 11812,
      "affiliateCommission": 0,
      "voucherXtra": 3500,
      "orderProcessingFee": 3000,
      "taxWithheld": 5316,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 28003,
    "cogsTotal": 0
  },
  {
    "id": "584643979251451748",
    "platform": "tiktok",
    "date": "2026-06-22",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 35100,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 35100,
    "cogsTotal": 0
  },
  {
    "id": "584722772158219706",
    "platform": "tiktok",
    "date": "2026-06-26",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584681753869648940",
    "platform": "tiktok",
    "date": "2026-06-24",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 56875,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4375,
      "commissionFee": 11812,
      "affiliateCommission": 0,
      "voucherXtra": 3500,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22687,
    "cogsTotal": 0
  },
  {
    "id": "584681510164924325",
    "platform": "tiktok",
    "date": "2026-06-24",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 57183,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4375,
      "commissionFee": 11812,
      "affiliateCommission": 0,
      "voucherXtra": 3500,
      "orderProcessingFee": 3000,
      "taxWithheld": 5005,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 27692,
    "cogsTotal": 0
  },
  {
    "id": "584680893524771955",
    "platform": "tiktok",
    "date": "2026-06-24",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 59952,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 11408,
      "affiliateCommission": 0,
      "voucherXtra": 3380,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22013,
    "cogsTotal": 0
  },
  {
    "id": "584677323597973251",
    "platform": "tiktok",
    "date": "2026-06-24",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 54190,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 6880,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 29230,
    "cogsTotal": 0
  },
  {
    "id": "584673468719793975",
    "platform": "tiktok",
    "date": "2026-06-23",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61070,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584669785011881045",
    "platform": "tiktok",
    "date": "2026-06-23",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 56770,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 4300,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 26650,
    "cogsTotal": 0
  },
  {
    "id": "584668758038644454",
    "platform": "tiktok",
    "date": "2026-06-23",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61070,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584667840294847853",
    "platform": "tiktok",
    "date": "2026-06-23",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 70755,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4950,
      "commissionFee": 13365,
      "affiliateCommission": 0,
      "voucherXtra": 3960,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 25275,
    "cogsTotal": 0
  },
  {
    "id": "584657544100153185",
    "platform": "tiktok",
    "date": "2026-06-22",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 78205,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 5450,
      "commissionFee": 14715,
      "affiliateCommission": 0,
      "voucherXtra": 4360,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 27525,
    "cogsTotal": 0
  },
  {
    "id": "584650781984720853",
    "platform": "tiktok",
    "date": "2026-06-22",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 51502,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 11408,
      "affiliateCommission": 0,
      "voucherXtra": 3380,
      "orderProcessingFee": 3000,
      "taxWithheld": 8450,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 30463,
    "cogsTotal": 0
  },
  {
    "id": "584646981210900351",
    "platform": "tiktok",
    "date": "2026-06-22",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 54190,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584640759432382353",
    "platform": "tiktok",
    "date": "2026-06-21",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 57277,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4375,
      "commissionFee": 11812,
      "affiliateCommission": 0,
      "voucherXtra": 3500,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22687,
    "cogsTotal": 0
  },
  {
    "id": "584638475417847036",
    "platform": "tiktok",
    "date": "2026-06-21",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 57760,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4375,
      "commissionFee": 11812,
      "affiliateCommission": 0,
      "voucherXtra": 3500,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22687,
    "cogsTotal": 0
  },
  {
    "id": "584722328558339149",
    "platform": "tiktok",
    "date": "2026-06-26",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584668198896501772",
    "platform": "tiktok",
    "date": "2026-06-23",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 56938,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4375,
      "commissionFee": 11812,
      "affiliateCommission": 0,
      "voucherXtra": 3500,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22687,
    "cogsTotal": 0
  },
  {
    "id": "584668113097098252",
    "platform": "tiktok",
    "date": "2026-06-23",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 56413,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4375,
      "commissionFee": 11812,
      "affiliateCommission": 0,
      "voucherXtra": 3500,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22687,
    "cogsTotal": 0
  },
  {
    "id": "584667799637689629",
    "platform": "tiktok",
    "date": "2026-06-23",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61070,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584664146382783595",
    "platform": "tiktok",
    "date": "2026-06-23",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61070,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584663325294233062",
    "platform": "tiktok",
    "date": "2026-06-23",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61070,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584663224971331294",
    "platform": "tiktok",
    "date": "2026-06-23",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 54758,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4380,
      "commissionFee": 11812,
      "affiliateCommission": 0,
      "voucherXtra": 3500,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22692,
    "cogsTotal": 0
  },
  {
    "id": "584659715167061749",
    "platform": "tiktok",
    "date": "2026-06-23",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 56968,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4375,
      "commissionFee": 11812,
      "affiliateCommission": 0,
      "voucherXtra": 3500,
      "orderProcessingFee": 3000,
      "taxWithheld": 5220,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 27907,
    "cogsTotal": 0
  },
  {
    "id": "584654923158881665",
    "platform": "tiktok",
    "date": "2026-06-22",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61070,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584643280563111427",
    "platform": "tiktok",
    "date": "2026-06-21",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 59952,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 11408,
      "affiliateCommission": 0,
      "voucherXtra": 3380,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22013,
    "cogsTotal": 0
  },
  {
    "id": "584642525821371773",
    "platform": "tiktok",
    "date": "2026-06-21",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 59952,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 11408,
      "affiliateCommission": 0,
      "voucherXtra": 3380,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22013,
    "cogsTotal": 0
  },
  {
    "id": "584634334681073644",
    "platform": "tiktok",
    "date": "2026-06-21",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 56163,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4375,
      "commissionFee": 11812,
      "affiliateCommission": 0,
      "voucherXtra": 3500,
      "orderProcessingFee": 3000,
      "taxWithheld": 6025,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 28712,
    "cogsTotal": 0
  },
  {
    "id": "584632251655947410",
    "platform": "tiktok",
    "date": "2026-06-21",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 50657,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 11408,
      "affiliateCommission": 0,
      "voucherXtra": 3380,
      "orderProcessingFee": 3000,
      "taxWithheld": 9295,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 31308,
    "cogsTotal": 0
  },
  {
    "id": "584622156068258975",
    "platform": "tiktok",
    "date": "2026-06-20",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 57253,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4375,
      "commissionFee": 11812,
      "affiliateCommission": 0,
      "voucherXtra": 3500,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22687,
    "cogsTotal": 0
  },
  {
    "id": "584622120888993477",
    "platform": "tiktok",
    "date": "2026-06-20",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 56124,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4375,
      "commissionFee": 11812,
      "affiliateCommission": 0,
      "voucherXtra": 3500,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22687,
    "cogsTotal": 0
  },
  {
    "id": "584496838639256591",
    "platform": "tiktok",
    "date": "2026-06-12",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 35500,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 35500,
    "cogsTotal": 0
  },
  {
    "id": "584701506413823158",
    "platform": "tiktok",
    "date": "2026-06-25",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584648409988105232",
    "platform": "tiktok",
    "date": "2026-06-22",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 52563,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4375,
      "commissionFee": 11812,
      "affiliateCommission": 0,
      "voucherXtra": 3500,
      "orderProcessingFee": 3000,
      "taxWithheld": 9625,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 32312,
    "cogsTotal": 0
  },
  {
    "id": "584642282865199045",
    "platform": "tiktok",
    "date": "2026-06-21",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 51502,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 11408,
      "affiliateCommission": 0,
      "voucherXtra": 3380,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22013,
    "cogsTotal": 0
  },
  {
    "id": "584641790205724645",
    "platform": "tiktok",
    "date": "2026-06-21",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 86029,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 5974,
      "commissionFee": 16132,
      "affiliateCommission": 0,
      "voucherXtra": 4780,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 29886,
    "cogsTotal": 0
  },
  {
    "id": "584640237649037127",
    "platform": "tiktok",
    "date": "2026-06-21",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 82730,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 5975,
      "commissionFee": 16132,
      "affiliateCommission": 0,
      "voucherXtra": 4780,
      "orderProcessingFee": 3000,
      "taxWithheld": 3298,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 33185,
    "cogsTotal": 0
  },
  {
    "id": "584638617767216373",
    "platform": "tiktok",
    "date": "2026-06-21",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 59952,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 11408,
      "affiliateCommission": 0,
      "voucherXtra": 3380,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22013,
    "cogsTotal": 0
  },
  {
    "id": "584627154291164237",
    "platform": "tiktok",
    "date": "2026-06-21",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 70755,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4950,
      "commissionFee": 13365,
      "affiliateCommission": 0,
      "voucherXtra": 3960,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 25275,
    "cogsTotal": 0
  },
  {
    "id": "584625635341862239",
    "platform": "tiktok",
    "date": "2026-06-20",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 80053,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 5975,
      "commissionFee": 16132,
      "affiliateCommission": 0,
      "voucherXtra": 4780,
      "orderProcessingFee": 3000,
      "taxWithheld": 5975,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 35862,
    "cogsTotal": 0
  },
  {
    "id": "584625634753152303",
    "platform": "tiktok",
    "date": "2026-06-20",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61070,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584603762354390563",
    "platform": "tiktok",
    "date": "2026-06-19",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 78205,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 5450,
      "commissionFee": 14715,
      "affiliateCommission": 0,
      "voucherXtra": 4360,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 27525,
    "cogsTotal": 0
  },
  {
    "id": "584582705930470769",
    "platform": "tiktok",
    "date": "2026-06-18",
    "status": "returned",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 3000,
    "cogsTotal": 0,
    "returnReason": "Change of mind"
  },
  {
    "id": "584575638290465834",
    "platform": "tiktok",
    "date": "2026-06-17",
    "status": "returned",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 3000,
    "cogsTotal": 0,
    "returnReason": "Item too big or too small"
  },
  {
    "id": "584554616711448570",
    "platform": "tiktok",
    "date": "2026-06-16",
    "status": "returned",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 3000,
    "cogsTotal": 0,
    "returnReason": "Change of mind"
  },
  {
    "id": "584682245800494119",
    "platform": "tiktok",
    "date": "2026-06-24",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584678077268396012",
    "platform": "tiktok",
    "date": "2026-06-24",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584634740103415786",
    "platform": "tiktok",
    "date": "2026-06-21",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 59952,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 11408,
      "affiliateCommission": 0,
      "voucherXtra": 3380,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22013,
    "cogsTotal": 0
  },
  {
    "id": "584634719964726538",
    "platform": "tiktok",
    "date": "2026-06-21",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 59952,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 11408,
      "affiliateCommission": 0,
      "voucherXtra": 3380,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22013,
    "cogsTotal": 0
  },
  {
    "id": "584634681530943466",
    "platform": "tiktok",
    "date": "2026-06-21",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61070,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584634622389028309",
    "platform": "tiktok",
    "date": "2026-06-21",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 56479,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4375,
      "commissionFee": 11812,
      "affiliateCommission": 0,
      "voucherXtra": 3500,
      "orderProcessingFee": 3000,
      "taxWithheld": 5709,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 28396,
    "cogsTotal": 0
  },
  {
    "id": "584633990212388708",
    "platform": "tiktok",
    "date": "2026-06-21",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 52178,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 11408,
      "affiliateCommission": 0,
      "voucherXtra": 3380,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22013,
    "cogsTotal": 0
  },
  {
    "id": "584633263456617875",
    "platform": "tiktok",
    "date": "2026-06-21",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 52516,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 11408,
      "affiliateCommission": 0,
      "voucherXtra": 3380,
      "orderProcessingFee": 3000,
      "taxWithheld": 7436,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 29449,
    "cogsTotal": 0
  },
  {
    "id": "584632762813417362",
    "platform": "tiktok",
    "date": "2026-06-21",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 59952,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 11408,
      "affiliateCommission": 0,
      "voucherXtra": 3380,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22013,
    "cogsTotal": 0
  },
  {
    "id": "584627053462062956",
    "platform": "tiktok",
    "date": "2026-06-20",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 60855,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4950,
      "commissionFee": 13365,
      "affiliateCommission": 0,
      "voucherXtra": 3960,
      "orderProcessingFee": 3000,
      "taxWithheld": 9900,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 35175,
    "cogsTotal": 0
  },
  {
    "id": "584625910713714583",
    "platform": "tiktok",
    "date": "2026-06-20",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 52510,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 8560,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 30910,
    "cogsTotal": 0
  },
  {
    "id": "584624911668315894",
    "platform": "tiktok",
    "date": "2026-06-20",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 59952,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 11408,
      "affiliateCommission": 0,
      "voucherXtra": 3380,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22013,
    "cogsTotal": 0
  },
  {
    "id": "584624740298164177",
    "platform": "tiktok",
    "date": "2026-06-20",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 77160,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 6495,
      "commissionFee": 14715,
      "affiliateCommission": 0,
      "voucherXtra": 4360,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 28570,
    "cogsTotal": 0
  },
  {
    "id": "584620047719171090",
    "platform": "tiktok",
    "date": "2026-06-20",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 53438,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4375,
      "commissionFee": 11812,
      "affiliateCommission": 0,
      "voucherXtra": 3500,
      "orderProcessingFee": 3000,
      "taxWithheld": 8750,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 31437,
    "cogsTotal": 0
  },
  {
    "id": "584619878435292264",
    "platform": "tiktok",
    "date": "2026-06-20",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 56500,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4375,
      "commissionFee": 11812,
      "affiliateCommission": 0,
      "voucherXtra": 3500,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22687,
    "cogsTotal": 0
  },
  {
    "id": "584619330759984217",
    "platform": "tiktok",
    "date": "2026-06-20",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 59250,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 1820,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 24170,
    "cogsTotal": 0
  },
  {
    "id": "584618057525265639",
    "platform": "tiktok",
    "date": "2026-06-20",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 62188,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4375,
      "commissionFee": 11812,
      "affiliateCommission": 0,
      "voucherXtra": 3500,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22687,
    "cogsTotal": 0
  },
  {
    "id": "584615301852858301",
    "platform": "tiktok",
    "date": "2026-06-20",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 96830,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 6700,
      "commissionFee": 18090,
      "affiliateCommission": 0,
      "voucherXtra": 5360,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 33150,
    "cogsTotal": 0
  },
  {
    "id": "584603401683895395",
    "platform": "tiktok",
    "date": "2026-06-19",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 51679,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 11408,
      "affiliateCommission": 0,
      "voucherXtra": 3380,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22013,
    "cogsTotal": 0
  },
  {
    "id": "584600743791265110",
    "platform": "tiktok",
    "date": "2026-06-19",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 53810,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584471691706008625",
    "platform": "tiktok",
    "date": "2026-06-11",
    "status": "returned",
    "items": [],
    "gmv": 0,
    "settlementAmount": 49976,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4107,
      "commissionFee": 11088,
      "affiliateCommission": 0,
      "voucherXtra": 3285,
      "orderProcessingFee": 3000,
      "taxWithheld": 8213,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 29693,
    "cogsTotal": 0,
    "returnReason": "Change of mind"
  },
  {
    "id": "584670191393015510",
    "platform": "tiktok",
    "date": "2026-06-23",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584669674307618419",
    "platform": "tiktok",
    "date": "2026-06-23",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584666427463599116",
    "platform": "tiktok",
    "date": "2026-06-23",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584664896112526359",
    "platform": "tiktok",
    "date": "2026-06-23",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584615606787867738",
    "platform": "tiktok",
    "date": "2026-06-20",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61025,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4345,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22395,
    "cogsTotal": 0
  },
  {
    "id": "584614590167483723",
    "platform": "tiktok",
    "date": "2026-06-20",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 86028,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 5975,
      "commissionFee": 16132,
      "affiliateCommission": 0,
      "voucherXtra": 4780,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 29887,
    "cogsTotal": 0
  },
  {
    "id": "584613238018507943",
    "platform": "tiktok",
    "date": "2026-06-20",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 56500,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4375,
      "commissionFee": 11812,
      "affiliateCommission": 0,
      "voucherXtra": 3500,
      "orderProcessingFee": 3000,
      "taxWithheld": 5688,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 28375,
    "cogsTotal": 0
  },
  {
    "id": "584611263950980846",
    "platform": "tiktok",
    "date": "2026-06-19",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 59197,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4980,
      "commissionFee": 11408,
      "affiliateCommission": 0,
      "voucherXtra": 3380,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22768,
    "cogsTotal": 0
  },
  {
    "id": "584610925290358473",
    "platform": "tiktok",
    "date": "2026-06-19",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 62087,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4375,
      "commissionFee": 11812,
      "affiliateCommission": 0,
      "voucherXtra": 3500,
      "orderProcessingFee": 3000,
      "taxWithheld": 101,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22788,
    "cogsTotal": 0
  },
  {
    "id": "584605853212771550",
    "platform": "tiktok",
    "date": "2026-06-19",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 78205,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 5450,
      "commissionFee": 14715,
      "affiliateCommission": 0,
      "voucherXtra": 4360,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 27525,
    "cogsTotal": 0
  },
  {
    "id": "584605799496648094",
    "platform": "tiktok",
    "date": "2026-06-19",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 58197,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 5980,
      "commissionFee": 11408,
      "affiliateCommission": 0,
      "voucherXtra": 3380,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 23768,
    "cogsTotal": 0
  },
  {
    "id": "584604801399031194",
    "platform": "tiktok",
    "date": "2026-06-19",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 55472,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4480,
      "commissionFee": 11408,
      "affiliateCommission": 0,
      "voucherXtra": 3380,
      "orderProcessingFee": 3000,
      "taxWithheld": 4225,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 26493,
    "cogsTotal": 0
  },
  {
    "id": "584601682523752111",
    "platform": "tiktok",
    "date": "2026-06-19",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 59952,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 11408,
      "affiliateCommission": 0,
      "voucherXtra": 3380,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22013,
    "cogsTotal": 0
  },
  {
    "id": "584589748584154410",
    "platform": "tiktok",
    "date": "2026-06-18",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 54195,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4100,
      "commissionFee": 11070,
      "affiliateCommission": 0,
      "voucherXtra": 3280,
      "orderProcessingFee": 3000,
      "taxWithheld": 3895,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 25345,
    "cogsTotal": 0
  },
  {
    "id": "584588520836072805",
    "platform": "tiktok",
    "date": "2026-06-18",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 57200,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 3870,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 26220,
    "cogsTotal": 0
  },
  {
    "id": "584586974754604506",
    "platform": "tiktok",
    "date": "2026-06-18",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 58577,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4375,
      "commissionFee": 11812,
      "affiliateCommission": 0,
      "voucherXtra": 3500,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22687,
    "cogsTotal": 0
  },
  {
    "id": "584659677515712245",
    "platform": "tiktok",
    "date": "2026-06-23",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584657428195804406",
    "platform": "tiktok",
    "date": "2026-06-22",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584657187960882361",
    "platform": "tiktok",
    "date": "2026-06-22",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584655309972276759",
    "platform": "tiktok",
    "date": "2026-06-22",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584648795897169769",
    "platform": "tiktok",
    "date": "2026-06-22",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584642212442179043",
    "platform": "tiktok",
    "date": "2026-06-21",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584603403860019173",
    "platform": "tiktok",
    "date": "2026-06-19",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 59952,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 11408,
      "affiliateCommission": 0,
      "voucherXtra": 3380,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22013,
    "cogsTotal": 0
  },
  {
    "id": "584598859589321853",
    "platform": "tiktok",
    "date": "2026-06-19",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 78205,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 5450,
      "commissionFee": 14715,
      "affiliateCommission": 0,
      "voucherXtra": 4360,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 27525,
    "cogsTotal": 0
  },
  {
    "id": "584598247083312170",
    "platform": "tiktok",
    "date": "2026-06-19",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 59952,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 11408,
      "affiliateCommission": 0,
      "voucherXtra": 3380,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22013,
    "cogsTotal": 0
  },
  {
    "id": "584582265629869424",
    "platform": "tiktok",
    "date": "2026-06-18",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 67030,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4700,
      "commissionFee": 12690,
      "affiliateCommission": 0,
      "voucherXtra": 3760,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 24150,
    "cogsTotal": 0
  },
  {
    "id": "584568322290976566",
    "platform": "tiktok",
    "date": "2026-06-17",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 59952,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 11408,
      "affiliateCommission": 0,
      "voucherXtra": 3380,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22013,
    "cogsTotal": 0
  },
  {
    "id": "584566416406971896",
    "platform": "tiktok",
    "date": "2026-06-17",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 52563,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4125,
      "commissionFee": 11138,
      "affiliateCommission": 0,
      "voucherXtra": 3300,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 21563,
    "cogsTotal": 0
  },
  {
    "id": "584566182851347566",
    "platform": "tiktok",
    "date": "2026-06-17",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 56856,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 4214,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 26564,
    "cogsTotal": 0
  },
  {
    "id": "584565252528506099",
    "platform": "tiktok",
    "date": "2026-06-17",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 52109,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4125,
      "commissionFee": 11138,
      "affiliateCommission": 0,
      "voucherXtra": 3300,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 21563,
    "cogsTotal": 0
  },
  {
    "id": "584563210216768675",
    "platform": "tiktok",
    "date": "2026-06-16",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 51502,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 11408,
      "affiliateCommission": 0,
      "voucherXtra": 3380,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22013,
    "cogsTotal": 0
  },
  {
    "id": "584499633385801236",
    "platform": "tiktok",
    "date": "2026-06-12",
    "status": "returned",
    "items": [],
    "gmv": 0,
    "settlementAmount": 86028,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 5975,
      "commissionFee": 16132,
      "affiliateCommission": 0,
      "voucherXtra": 4780,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 29887,
    "cogsTotal": 0,
    "returnReason": "Change of mind"
  },
  {
    "id": "584626372523427820",
    "platform": "tiktok",
    "date": "2026-06-20",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584621954385479577",
    "platform": "tiktok",
    "date": "2026-06-20",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584582978595882030",
    "platform": "tiktok",
    "date": "2026-06-18",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 63305,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4450,
      "commissionFee": 12015,
      "affiliateCommission": 0,
      "voucherXtra": 3560,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 23025,
    "cogsTotal": 0
  },
  {
    "id": "584582503401424124",
    "platform": "tiktok",
    "date": "2026-06-18",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 52685,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 11408,
      "affiliateCommission": 0,
      "voucherXtra": 3380,
      "orderProcessingFee": 3000,
      "taxWithheld": 7267,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 29280,
    "cogsTotal": 0
  },
  {
    "id": "584577904458761305",
    "platform": "tiktok",
    "date": "2026-06-17",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 55727,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 11408,
      "affiliateCommission": 0,
      "voucherXtra": 3380,
      "orderProcessingFee": 3000,
      "taxWithheld": 4225,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 26238,
    "cogsTotal": 0
  },
  {
    "id": "584577717239056290",
    "platform": "tiktok",
    "date": "2026-06-17",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61070,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584574624187385206",
    "platform": "tiktok",
    "date": "2026-06-17",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61070,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584564890403964382",
    "platform": "tiktok",
    "date": "2026-06-17",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 55727,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 11408,
      "affiliateCommission": 0,
      "voucherXtra": 3380,
      "orderProcessingFee": 3000,
      "taxWithheld": 4225,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 26238,
    "cogsTotal": 0
  },
  {
    "id": "584557952368936372",
    "platform": "tiktok",
    "date": "2026-06-16",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 59952,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 11408,
      "affiliateCommission": 0,
      "voucherXtra": 3380,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22013,
    "cogsTotal": 0
  },
  {
    "id": "584548811920738019",
    "platform": "tiktok",
    "date": "2026-06-16",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 78205,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 5450,
      "commissionFee": 14715,
      "affiliateCommission": 0,
      "voucherXtra": 4360,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 27525,
    "cogsTotal": 0
  },
  {
    "id": "584546278842598505",
    "platform": "tiktok",
    "date": "2026-06-15",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 55727,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 11408,
      "affiliateCommission": 0,
      "voucherXtra": 3380,
      "orderProcessingFee": 3000,
      "taxWithheld": 4225,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 26238,
    "cogsTotal": 0
  },
  {
    "id": "584544333034391373",
    "platform": "tiktok",
    "date": "2026-06-15",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 78205,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 5450,
      "commissionFee": 14715,
      "affiliateCommission": 0,
      "voucherXtra": 4360,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 27525,
    "cogsTotal": 0
  },
  {
    "id": "584533454121568093",
    "platform": "tiktok",
    "date": "2026-06-15",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61442,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4325,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22463,
    "cogsTotal": 0
  },
  {
    "id": "584523562468410705",
    "platform": "tiktok",
    "date": "2026-06-14",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 60412,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4450,
      "commissionFee": 12015,
      "affiliateCommission": 0,
      "voucherXtra": 3560,
      "orderProcessingFee": 3000,
      "taxWithheld": 2893,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 25918,
    "cogsTotal": 0
  },
  {
    "id": "584435881477309503",
    "platform": "tiktok",
    "date": "2026-06-08",
    "status": "returned",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 3000,
    "cogsTotal": 0,
    "returnReason": "Change of mind"
  },
  {
    "id": "584625319945733423",
    "platform": "tiktok",
    "date": "2026-06-20",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584623707962508291",
    "platform": "tiktok",
    "date": "2026-06-20",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584617345141737458",
    "platform": "tiktok",
    "date": "2026-06-20",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584615120099116080",
    "platform": "tiktok",
    "date": "2026-06-20",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584569684243023413",
    "platform": "tiktok",
    "date": "2026-06-17",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 59952,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 11408,
      "affiliateCommission": 0,
      "voucherXtra": 3380,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22013,
    "cogsTotal": 0
  },
  {
    "id": "584569020237252329",
    "platform": "tiktok",
    "date": "2026-06-17",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 55727,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 11408,
      "affiliateCommission": 0,
      "voucherXtra": 3380,
      "orderProcessingFee": 3000,
      "taxWithheld": 4225,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 26238,
    "cogsTotal": 0
  },
  {
    "id": "584568621338231918",
    "platform": "tiktok",
    "date": "2026-06-17",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 59952,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 11408,
      "affiliateCommission": 0,
      "voucherXtra": 3380,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22013,
    "cogsTotal": 0
  },
  {
    "id": "584568391677282278",
    "platform": "tiktok",
    "date": "2026-06-17",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61070,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584565287718585657",
    "platform": "tiktok",
    "date": "2026-06-17",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 52109,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4125,
      "commissionFee": 11138,
      "affiliateCommission": 0,
      "voucherXtra": 3300,
      "orderProcessingFee": 3000,
      "taxWithheld": 6353,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 27916,
    "cogsTotal": 0
  },
  {
    "id": "584565263539471932",
    "platform": "tiktok",
    "date": "2026-06-17",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 118878,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 8610,
      "commissionFee": 23247,
      "affiliateCommission": 0,
      "voucherXtra": 6888,
      "orderProcessingFee": 3000,
      "taxWithheld": 6411,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 48156,
    "cogsTotal": 0
  },
  {
    "id": "584558262689892263",
    "platform": "tiktok",
    "date": "2026-06-16",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 52109,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4125,
      "commissionFee": 11138,
      "affiliateCommission": 0,
      "voucherXtra": 3300,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 21563,
    "cogsTotal": 0
  },
  {
    "id": "584558022073550146",
    "platform": "tiktok",
    "date": "2026-06-16",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 56065,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 11408,
      "affiliateCommission": 0,
      "voucherXtra": 3380,
      "orderProcessingFee": 3000,
      "taxWithheld": 3887,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 25900,
    "cogsTotal": 0
  },
  {
    "id": "584557864346748839",
    "platform": "tiktok",
    "date": "2026-06-16",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 50657,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 11408,
      "affiliateCommission": 0,
      "voucherXtra": 3380,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22013,
    "cogsTotal": 0
  },
  {
    "id": "584557325280118055",
    "platform": "tiktok",
    "date": "2026-06-16",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 53966,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 11408,
      "affiliateCommission": 0,
      "voucherXtra": 3380,
      "orderProcessingFee": 3000,
      "taxWithheld": 5986,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 27999,
    "cogsTotal": 0
  },
  {
    "id": "584555394908718096",
    "platform": "tiktok",
    "date": "2026-06-16",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 56770,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 4300,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 26650,
    "cogsTotal": 0
  },
  {
    "id": "584554421475050940",
    "platform": "tiktok",
    "date": "2026-06-16",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61070,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584553775253521940",
    "platform": "tiktok",
    "date": "2026-06-16",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 55727,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 11408,
      "affiliateCommission": 0,
      "voucherXtra": 3380,
      "orderProcessingFee": 3000,
      "taxWithheld": 4225,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 26238,
    "cogsTotal": 0
  },
  {
    "id": "584552061940369262",
    "platform": "tiktok",
    "date": "2026-06-16",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 59952,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 11408,
      "affiliateCommission": 0,
      "voucherXtra": 3380,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22013,
    "cogsTotal": 0
  },
  {
    "id": "584551558529648322",
    "platform": "tiktok",
    "date": "2026-06-16",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 59952,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 11408,
      "affiliateCommission": 0,
      "voucherXtra": 3380,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22013,
    "cogsTotal": 0
  },
  {
    "id": "584551030711748551",
    "platform": "tiktok",
    "date": "2026-06-16",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 52470,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584544740793812533",
    "platform": "tiktok",
    "date": "2026-06-15",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 160493,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 11003,
      "commissionFee": 31303,
      "affiliateCommission": 0,
      "voucherXtra": 9275,
      "orderProcessingFee": 3000,
      "taxWithheld": 9844,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 64425,
    "cogsTotal": 0
  },
  {
    "id": "584538411374249785",
    "platform": "tiktok",
    "date": "2026-06-15",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 63678,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4475,
      "commissionFee": 12082,
      "affiliateCommission": 0,
      "voucherXtra": 3580,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 23137,
    "cogsTotal": 0
  },
  {
    "id": "584529512916289452",
    "platform": "tiktok",
    "date": "2026-06-14",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 57417,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 11408,
      "affiliateCommission": 0,
      "voucherXtra": 3380,
      "orderProcessingFee": 3000,
      "taxWithheld": 2535,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 24548,
    "cogsTotal": 0
  },
  {
    "id": "584527491362227851",
    "platform": "tiktok",
    "date": "2026-06-14",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 63305,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4450,
      "commissionFee": 12015,
      "affiliateCommission": 0,
      "voucherXtra": 3560,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 23025,
    "cogsTotal": 0
  },
  {
    "id": "584523267490154169",
    "platform": "tiktok",
    "date": "2026-06-14",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61167,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4600,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22738,
    "cogsTotal": 0
  },
  {
    "id": "584508977860281902",
    "platform": "tiktok",
    "date": "2026-06-13",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 52792,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4325,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22463,
    "cogsTotal": 0
  },
  {
    "id": "584601383343064163",
    "platform": "tiktok",
    "date": "2026-06-19",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584557640684635310",
    "platform": "tiktok",
    "date": "2026-06-16",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 56065,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 11408,
      "affiliateCommission": 0,
      "voucherXtra": 3380,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22013,
    "cogsTotal": 0
  },
  {
    "id": "584556983648290110",
    "platform": "tiktok",
    "date": "2026-06-16",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 55727,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 11408,
      "affiliateCommission": 0,
      "voucherXtra": 3380,
      "orderProcessingFee": 3000,
      "taxWithheld": 4225,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 26238,
    "cogsTotal": 0
  },
  {
    "id": "584556717675742457",
    "platform": "tiktok",
    "date": "2026-06-16",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 67305,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 5450,
      "commissionFee": 14715,
      "affiliateCommission": 0,
      "voucherXtra": 4360,
      "orderProcessingFee": 3000,
      "taxWithheld": 10900,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 38425,
    "cogsTotal": 0
  },
  {
    "id": "584555089060136038",
    "platform": "tiktok",
    "date": "2026-06-16",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 114612,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 8461,
      "commissionFee": 22845,
      "affiliateCommission": 0,
      "voucherXtra": 6769,
      "orderProcessingFee": 3000,
      "taxWithheld": 8462,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 49537,
    "cogsTotal": 0
  },
  {
    "id": "584554957696238734",
    "platform": "tiktok",
    "date": "2026-06-16",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 54375,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 11408,
      "affiliateCommission": 0,
      "voucherXtra": 3380,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22013,
    "cogsTotal": 0
  },
  {
    "id": "584554739983680832",
    "platform": "tiktok",
    "date": "2026-06-16",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 52769,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 11408,
      "affiliateCommission": 0,
      "voucherXtra": 3380,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22013,
    "cogsTotal": 0
  },
  {
    "id": "584554751869617672",
    "platform": "tiktok",
    "date": "2026-06-16",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 55858,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 11408,
      "affiliateCommission": 0,
      "voucherXtra": 3380,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22013,
    "cogsTotal": 0
  },
  {
    "id": "584554465068812190",
    "platform": "tiktok",
    "date": "2026-06-16",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 55727,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 11408,
      "affiliateCommission": 0,
      "voucherXtra": 3380,
      "orderProcessingFee": 3000,
      "taxWithheld": 4225,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 26238,
    "cogsTotal": 0
  },
  {
    "id": "584554260438680619",
    "platform": "tiktok",
    "date": "2026-06-16",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61070,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584552761157912361",
    "platform": "tiktok",
    "date": "2026-06-16",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 117056,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 8058,
      "commissionFee": 21755,
      "affiliateCommission": 0,
      "voucherXtra": 6446,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 39259,
    "cogsTotal": 0
  },
  {
    "id": "584550548029212582",
    "platform": "tiktok",
    "date": "2026-06-16",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 56318,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 11408,
      "affiliateCommission": 0,
      "voucherXtra": 3380,
      "orderProcessingFee": 3000,
      "taxWithheld": 3634,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 25647,
    "cogsTotal": 0
  },
  {
    "id": "584548765813146836",
    "platform": "tiktok",
    "date": "2026-06-16",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 59952,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 11408,
      "affiliateCommission": 0,
      "voucherXtra": 3380,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22013,
    "cogsTotal": 0
  },
  {
    "id": "584548705238156352",
    "platform": "tiktok",
    "date": "2026-06-16",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 52117,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 11408,
      "affiliateCommission": 0,
      "voucherXtra": 3380,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22013,
    "cogsTotal": 0
  },
  {
    "id": "584547228870346270",
    "platform": "tiktok",
    "date": "2026-06-15",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 59952,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 11408,
      "affiliateCommission": 0,
      "voucherXtra": 3380,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22013,
    "cogsTotal": 0
  },
  {
    "id": "584543348761527365",
    "platform": "tiktok",
    "date": "2026-06-15",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 59952,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 11408,
      "affiliateCommission": 0,
      "voucherXtra": 3380,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22013,
    "cogsTotal": 0
  },
  {
    "id": "584533538699117685",
    "platform": "tiktok",
    "date": "2026-06-15",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 59055,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4725,
      "commissionFee": 11340,
      "affiliateCommission": 0,
      "voucherXtra": 3360,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22425,
    "cogsTotal": 0
  },
  {
    "id": "584530608425894944",
    "platform": "tiktok",
    "date": "2026-06-14",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 58490,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 2580,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 24930,
    "cogsTotal": 0
  },
  {
    "id": "584506962931975686",
    "platform": "tiktok",
    "date": "2026-06-13",
    "status": "returned",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 3000,
    "cogsTotal": 0,
    "returnReason": "Change of mind"
  },
  {
    "id": "584540520588740019",
    "platform": "tiktok",
    "date": "2026-06-15",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 59952,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 11408,
      "affiliateCommission": 0,
      "voucherXtra": 3380,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22013,
    "cogsTotal": 0
  },
  {
    "id": "584540265799911247",
    "platform": "tiktok",
    "date": "2026-06-15",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 58462,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4125,
      "commissionFee": 11138,
      "affiliateCommission": 0,
      "voucherXtra": 3300,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 21563,
    "cogsTotal": 0
  },
  {
    "id": "584529811145393226",
    "platform": "tiktok",
    "date": "2026-06-14",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 67030,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4700,
      "commissionFee": 12690,
      "affiliateCommission": 0,
      "voucherXtra": 3760,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 24150,
    "cogsTotal": 0
  },
  {
    "id": "584528306781062296",
    "platform": "tiktok",
    "date": "2026-06-14",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 63305,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4450,
      "commissionFee": 12015,
      "affiliateCommission": 0,
      "voucherXtra": 3560,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 23025,
    "cogsTotal": 0
  },
  {
    "id": "584526725504468121",
    "platform": "tiktok",
    "date": "2026-06-14",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 52362,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 8708,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 31058,
    "cogsTotal": 0
  },
  {
    "id": "584525649262314697",
    "platform": "tiktok",
    "date": "2026-06-14",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 52792,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4325,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22463,
    "cogsTotal": 0
  },
  {
    "id": "584525179381385058",
    "platform": "tiktok",
    "date": "2026-06-14",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 56770,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 4300,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 26650,
    "cogsTotal": 0
  },
  {
    "id": "584523267577841389",
    "platform": "tiktok",
    "date": "2026-06-14",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61442,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4325,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22463,
    "cogsTotal": 0
  },
  {
    "id": "584522592373213041",
    "platform": "tiktok",
    "date": "2026-06-14",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 59203,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4475,
      "commissionFee": 12082,
      "affiliateCommission": 0,
      "voucherXtra": 3580,
      "orderProcessingFee": 3000,
      "taxWithheld": 4475,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 27612,
    "cogsTotal": 0
  },
  {
    "id": "584512402165106473",
    "platform": "tiktok",
    "date": "2026-06-13",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61442,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4325,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22463,
    "cogsTotal": 0
  },
  {
    "id": "584569131764974858",
    "platform": "tiktok",
    "date": "2026-06-17",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584569137354933514",
    "platform": "tiktok",
    "date": "2026-06-17",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584567570739070024",
    "platform": "tiktok",
    "date": "2026-06-17",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584502709607957569",
    "platform": "tiktok",
    "date": "2026-06-13",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 52792,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4325,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22463,
    "cogsTotal": 0
  },
  {
    "id": "584497039369275212",
    "platform": "tiktok",
    "date": "2026-06-12",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 55117,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4450,
      "commissionFee": 12015,
      "affiliateCommission": 0,
      "voucherXtra": 3560,
      "orderProcessingFee": 3000,
      "taxWithheld": 8188,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 31213,
    "cogsTotal": 0
  },
  {
    "id": "584494085927765277",
    "platform": "tiktok",
    "date": "2026-06-12",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 59952,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 11408,
      "affiliateCommission": 0,
      "voucherXtra": 3380,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22013,
    "cogsTotal": 0
  },
  {
    "id": "584493205343208601",
    "platform": "tiktok",
    "date": "2026-06-12",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 56770,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 4300,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 26650,
    "cogsTotal": 0
  },
  {
    "id": "584487688022296207",
    "platform": "tiktok",
    "date": "2026-06-12",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 51927,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4325,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 9515,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 31978,
    "cogsTotal": 0
  },
  {
    "id": "584556629202076967",
    "platform": "tiktok",
    "date": "2026-06-16",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584556332500551296",
    "platform": "tiktok",
    "date": "2026-06-16",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584555824628729127",
    "platform": "tiktok",
    "date": "2026-06-16",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584554681349342587",
    "platform": "tiktok",
    "date": "2026-06-16",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584552422908724473",
    "platform": "tiktok",
    "date": "2026-06-16",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584552020182861549",
    "platform": "tiktok",
    "date": "2026-06-16",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584552060868855662",
    "platform": "tiktok",
    "date": "2026-06-16",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584503670122579377",
    "platform": "tiktok",
    "date": "2026-06-13",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 82084,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 5975,
      "commissionFee": 16132,
      "affiliateCommission": 0,
      "voucherXtra": 4780,
      "orderProcessingFee": 3000,
      "taxWithheld": 3944,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 33831,
    "cogsTotal": 0
  },
  {
    "id": "584497138438604397",
    "platform": "tiktok",
    "date": "2026-06-12",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 58855,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4450,
      "commissionFee": 12015,
      "affiliateCommission": 0,
      "voucherXtra": 3560,
      "orderProcessingFee": 3000,
      "taxWithheld": 4450,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 27475,
    "cogsTotal": 0
  },
  {
    "id": "584496748531778777",
    "platform": "tiktok",
    "date": "2026-06-12",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61070,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584495711089100326",
    "platform": "tiktok",
    "date": "2026-06-12",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 57965,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4450,
      "commissionFee": 12015,
      "affiliateCommission": 0,
      "voucherXtra": 3560,
      "orderProcessingFee": 3000,
      "taxWithheld": 5340,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 28365,
    "cogsTotal": 0
  },
  {
    "id": "584495221806696189",
    "platform": "tiktok",
    "date": "2026-06-12",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61070,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584495142434735446",
    "platform": "tiktok",
    "date": "2026-06-12",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 58576,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 2494,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 24844,
    "cogsTotal": 0
  },
  {
    "id": "584495002330826212",
    "platform": "tiktok",
    "date": "2026-06-12",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 55652,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 5418,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 27768,
    "cogsTotal": 0
  },
  {
    "id": "584492941569197281",
    "platform": "tiktok",
    "date": "2026-06-12",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 56318,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 11408,
      "affiliateCommission": 0,
      "voucherXtra": 3380,
      "orderProcessingFee": 3000,
      "taxWithheld": 3634,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 25647,
    "cogsTotal": 0
  },
  {
    "id": "584486271598102081",
    "platform": "tiktok",
    "date": "2026-06-12",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 57413,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4475,
      "commissionFee": 12082,
      "affiliateCommission": 0,
      "voucherXtra": 3580,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 23137,
    "cogsTotal": 0
  },
  {
    "id": "584484716490098255",
    "platform": "tiktok",
    "date": "2026-06-12",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61070,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584443817956705597",
    "platform": "tiktok",
    "date": "2026-06-09",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 67402,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4725,
      "commissionFee": 12758,
      "affiliateCommission": 0,
      "voucherXtra": 3780,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 24263,
    "cogsTotal": 0
  },
  {
    "id": "584427874778514832",
    "platform": "tiktok",
    "date": "2026-06-08",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61070,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584405953909655318",
    "platform": "tiktok",
    "date": "2026-06-07",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 54728,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4475,
      "commissionFee": 12082,
      "affiliateCommission": 0,
      "voucherXtra": 3580,
      "orderProcessingFee": 3000,
      "taxWithheld": 8950,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 32087,
    "cogsTotal": 0
  },
  {
    "id": "584405947546502934",
    "platform": "tiktok",
    "date": "2026-06-07",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 52792,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4325,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 8650,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 31113,
    "cogsTotal": 0
  },
  {
    "id": "584543628831917120",
    "platform": "tiktok",
    "date": "2026-06-15",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584542774059893904",
    "platform": "tiktok",
    "date": "2026-06-15",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584494535720011090",
    "platform": "tiktok",
    "date": "2026-06-12",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 60635,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4450,
      "commissionFee": 12015,
      "affiliateCommission": 0,
      "voucherXtra": 3560,
      "orderProcessingFee": 3000,
      "taxWithheld": 2670,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 25695,
    "cogsTotal": 0
  },
  {
    "id": "584493629573465840",
    "platform": "tiktok",
    "date": "2026-06-12",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61442,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4325,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22463,
    "cogsTotal": 0
  },
  {
    "id": "584491024171435786",
    "platform": "tiktok",
    "date": "2026-06-12",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 62522,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 5631,
      "commissionFee": 12082,
      "affiliateCommission": 0,
      "voucherXtra": 3580,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 24293,
    "cogsTotal": 0
  },
  {
    "id": "584484917130004237",
    "platform": "tiktok",
    "date": "2026-06-12",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61070,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584484796603598441",
    "platform": "tiktok",
    "date": "2026-06-12",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 57117,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4325,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 4325,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 26788,
    "cogsTotal": 0
  },
  {
    "id": "584482864462661347",
    "platform": "tiktok",
    "date": "2026-06-11",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61442,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4325,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22463,
    "cogsTotal": 0
  },
  {
    "id": "584459864335156701",
    "platform": "tiktok",
    "date": "2026-06-10",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61070,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584457519427519563",
    "platform": "tiktok",
    "date": "2026-06-10",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 60545,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4825,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22875,
    "cogsTotal": 0
  },
  {
    "id": "584449078465365254",
    "platform": "tiktok",
    "date": "2026-06-09",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61442,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4325,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22463,
    "cogsTotal": 0
  },
  {
    "id": "584447056070346727",
    "platform": "tiktok",
    "date": "2026-06-09",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 164459,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 12981,
      "commissionFee": 35049,
      "affiliateCommission": 0,
      "voucherXtra": 10385,
      "orderProcessingFee": 3000,
      "taxWithheld": 25963,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 87378,
    "cogsTotal": 0
  },
  {
    "id": "584430074865747929",
    "platform": "tiktok",
    "date": "2026-06-08",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61070,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584423023416804924",
    "platform": "tiktok",
    "date": "2026-06-08",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61070,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584420336539043648",
    "platform": "tiktok",
    "date": "2026-06-07",
    "status": "returned",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 3000,
    "cogsTotal": 0,
    "returnReason": "Item too big or too small"
  },
  {
    "id": "584418997003912999",
    "platform": "tiktok",
    "date": "2026-06-07",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61070,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584417736171160953",
    "platform": "tiktok",
    "date": "2026-06-07",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 56337,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 11408,
      "affiliateCommission": 0,
      "voucherXtra": 3380,
      "orderProcessingFee": 3000,
      "taxWithheld": 3615,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 25628,
    "cogsTotal": 0
  },
  {
    "id": "584401831806731786",
    "platform": "tiktok",
    "date": "2026-06-06",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 57823,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4375,
      "commissionFee": 11812,
      "affiliateCommission": 0,
      "voucherXtra": 3500,
      "orderProcessingFee": 3000,
      "taxWithheld": 4365,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 27052,
    "cogsTotal": 0
  },
  {
    "id": "584400196730913888",
    "platform": "tiktok",
    "date": "2026-06-06",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 55727,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 11408,
      "affiliateCommission": 0,
      "voucherXtra": 3380,
      "orderProcessingFee": 3000,
      "taxWithheld": 4225,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 26238,
    "cogsTotal": 0
  },
  {
    "id": "584396884251543507",
    "platform": "tiktok",
    "date": "2026-06-06",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61070,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584476682329228898",
    "platform": "tiktok",
    "date": "2026-06-11",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61442,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4325,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22463,
    "cogsTotal": 0
  },
  {
    "id": "584476051788302030",
    "platform": "tiktok",
    "date": "2026-06-11",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 57117,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4325,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 4325,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 26788,
    "cogsTotal": 0
  },
  {
    "id": "584475953463985455",
    "platform": "tiktok",
    "date": "2026-06-11",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61442,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4325,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22463,
    "cogsTotal": 0
  },
  {
    "id": "584475059244205071",
    "platform": "tiktok",
    "date": "2026-06-11",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 54728,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4475,
      "commissionFee": 12082,
      "affiliateCommission": 0,
      "voucherXtra": 3580,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 23137,
    "cogsTotal": 0
  },
  {
    "id": "584471950259291294",
    "platform": "tiktok",
    "date": "2026-06-11",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 52792,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4325,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22463,
    "cogsTotal": 0
  },
  {
    "id": "584468934150096203",
    "platform": "tiktok",
    "date": "2026-06-11",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61442,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4325,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22463,
    "cogsTotal": 0
  },
  {
    "id": "584466619178387216",
    "platform": "tiktok",
    "date": "2026-06-10",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 59580,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4200,
      "commissionFee": 11340,
      "affiliateCommission": 0,
      "voucherXtra": 3360,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 21900,
    "cogsTotal": 0
  },
  {
    "id": "584409209701304214",
    "platform": "tiktok",
    "date": "2026-06-07",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 54728,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4475,
      "commissionFee": 12082,
      "affiliateCommission": 0,
      "voucherXtra": 3580,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 23137,
    "cogsTotal": 0
  },
  {
    "id": "584408247621289477",
    "platform": "tiktok",
    "date": "2026-06-07",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61442,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4325,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22463,
    "cogsTotal": 0
  },
  {
    "id": "584388292587915076",
    "platform": "tiktok",
    "date": "2026-06-06",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 51927,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4325,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22463,
    "cogsTotal": 0
  },
  {
    "id": "584382139529725764",
    "platform": "tiktok",
    "date": "2026-06-05",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 52974,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4325,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22463,
    "cogsTotal": 0
  },
  {
    "id": "584368801502037920",
    "platform": "tiktok",
    "date": "2026-06-05",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61070,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584368259014296799",
    "platform": "tiktok",
    "date": "2026-06-05",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61442,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4325,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22463,
    "cogsTotal": 0
  },
  {
    "id": "584368163323086472",
    "platform": "tiktok",
    "date": "2026-06-05",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61025,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4345,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22395,
    "cogsTotal": 0
  },
  {
    "id": "584358826278553114",
    "platform": "tiktok",
    "date": "2026-06-04",
    "status": "returned",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 3000,
    "cogsTotal": 0,
    "returnReason": "Change of mind"
  },
  {
    "id": "584334754431272721",
    "platform": "tiktok",
    "date": "2026-06-03",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 35100,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 35100,
    "cogsTotal": 0
  },
  {
    "id": "584299430474909491",
    "platform": "tiktok",
    "date": "2026-06-01",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 20900,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 20900,
    "cogsTotal": 0
  },
  {
    "id": "584287926775416385",
    "platform": "tiktok",
    "date": "2026-05-31",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 54294,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4450,
      "commissionFee": 12015,
      "affiliateCommission": 0,
      "voucherXtra": 3560,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 23025,
    "cogsTotal": 0
  },
  {
    "id": "584512482670576843",
    "platform": "tiktok",
    "date": "2026-06-13",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584509996120310802",
    "platform": "tiktok",
    "date": "2026-06-13",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584459622085658134",
    "platform": "tiktok",
    "date": "2026-06-10",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 55473,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4450,
      "commissionFee": 12015,
      "affiliateCommission": 0,
      "voucherXtra": 3560,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 23025,
    "cogsTotal": 0
  },
  {
    "id": "584451298155726489",
    "platform": "tiktok",
    "date": "2026-06-09",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 52792,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4325,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22463,
    "cogsTotal": 0
  },
  {
    "id": "584372939750737790",
    "platform": "tiktok",
    "date": "2026-06-05",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 52395,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4805,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 8170,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 31025,
    "cogsTotal": 0
  },
  {
    "id": "584369050949682877",
    "platform": "tiktok",
    "date": "2026-06-05",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 51927,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4325,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22463,
    "cogsTotal": 0
  },
  {
    "id": "584367667634603164",
    "platform": "tiktok",
    "date": "2026-06-05",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61070,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584363637367473721",
    "platform": "tiktok",
    "date": "2026-06-04",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 51927,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4325,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22463,
    "cogsTotal": 0
  },
  {
    "id": "584354437850105411",
    "platform": "tiktok",
    "date": "2026-06-04",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 117056,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 8058,
      "commissionFee": 21755,
      "affiliateCommission": 0,
      "voucherXtra": 6446,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 39259,
    "cogsTotal": 0
  },
  {
    "id": "584353535709447172",
    "platform": "tiktok",
    "date": "2026-06-04",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 63678,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4475,
      "commissionFee": 12082,
      "affiliateCommission": 0,
      "voucherXtra": 3580,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 23137,
    "cogsTotal": 0
  },
  {
    "id": "584347323513079516",
    "platform": "tiktok",
    "date": "2026-06-03",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 80068,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 5575,
      "commissionFee": 15052,
      "affiliateCommission": 0,
      "voucherXtra": 4460,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 28087,
    "cogsTotal": 0
  },
  {
    "id": "584497016958125354",
    "platform": "tiktok",
    "date": "2026-06-12",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584493648752837917",
    "platform": "tiktok",
    "date": "2026-06-12",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584376322807334044",
    "platform": "tiktok",
    "date": "2026-06-05",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61070,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584372504641897865",
    "platform": "tiktok",
    "date": "2026-06-05",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 102698,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 9982,
      "commissionFee": 22214,
      "affiliateCommission": 0,
      "voucherXtra": 6582,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 41778,
    "cogsTotal": 0
  },
  {
    "id": "584372182587639595",
    "platform": "tiktok",
    "date": "2026-06-05",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61070,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584369019282162843",
    "platform": "tiktok",
    "date": "2026-06-05",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61070,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584363237132896219",
    "platform": "tiktok",
    "date": "2026-06-04",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61164,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4700,
      "commissionFee": 12690,
      "affiliateCommission": 0,
      "voucherXtra": 3760,
      "orderProcessingFee": 3000,
      "taxWithheld": 5866,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 30016,
    "cogsTotal": 0
  },
  {
    "id": "584361468217427596",
    "platform": "tiktok",
    "date": "2026-06-04",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61070,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584355124018054270",
    "platform": "tiktok",
    "date": "2026-06-04",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61070,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584353383754859624",
    "platform": "tiktok",
    "date": "2026-06-04",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 58847,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4325,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 2595,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 25058,
    "cogsTotal": 0
  },
  {
    "id": "584352937831794026",
    "platform": "tiktok",
    "date": "2026-06-04",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61070,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584350549062092079",
    "platform": "tiktok",
    "date": "2026-06-04",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 52792,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4325,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22463,
    "cogsTotal": 0
  },
  {
    "id": "584337460876248579",
    "platform": "tiktok",
    "date": "2026-06-03",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 52178,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 11408,
      "affiliateCommission": 0,
      "voucherXtra": 3380,
      "orderProcessingFee": 3000,
      "taxWithheld": 7774,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 29787,
    "cogsTotal": 0
  },
  {
    "id": "584336276012304275",
    "platform": "tiktok",
    "date": "2026-06-03",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61442,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4325,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22463,
    "cogsTotal": 0
  },
  {
    "id": "584217809902339623",
    "platform": "tiktok",
    "date": "2026-05-27",
    "status": "returned",
    "items": [],
    "gmv": 0,
    "settlementAmount": 50348,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4325,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22463,
    "cogsTotal": 0,
    "returnReason": "Change of mind"
  },
  {
    "id": "584476012626216437",
    "platform": "tiktok",
    "date": "2026-06-11",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584358253173704581",
    "platform": "tiktok",
    "date": "2026-06-04",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 52792,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4325,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22463,
    "cogsTotal": 0
  },
  {
    "id": "584354551035823527",
    "platform": "tiktok",
    "date": "2026-06-04",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61070,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584318255903377200",
    "platform": "tiktok",
    "date": "2026-06-02",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 59397,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 6370,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 24508,
    "cogsTotal": 0
  },
  {
    "id": "584312854779233909",
    "platform": "tiktok",
    "date": "2026-06-02",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 59389,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4450,
      "commissionFee": 12015,
      "affiliateCommission": 0,
      "voucherXtra": 3560,
      "orderProcessingFee": 3000,
      "taxWithheld": 3916,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 26941,
    "cogsTotal": 0
  },
  {
    "id": "584467421811541970",
    "platform": "tiktok",
    "date": "2026-06-10",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584446090083337931",
    "platform": "tiktok",
    "date": "2026-06-09",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584333995415799490",
    "platform": "tiktok",
    "date": "2026-06-03",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 63305,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4450,
      "commissionFee": 12015,
      "affiliateCommission": 0,
      "voucherXtra": 3560,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 23025,
    "cogsTotal": 0
  },
  {
    "id": "584322185508193389",
    "platform": "tiktok",
    "date": "2026-06-02",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 53743,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4325,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22463,
    "cogsTotal": 0
  },
  {
    "id": "584278871644866551",
    "platform": "tiktok",
    "date": "2026-05-31",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 60397,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 5370,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 23508,
    "cogsTotal": 0
  },
  {
    "id": "584450738355995825",
    "platform": "tiktok",
    "date": "2026-06-09",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584315576162224008",
    "platform": "tiktok",
    "date": "2026-06-02",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 52792,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4325,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22463,
    "cogsTotal": 0
  },
  {
    "id": "584315322382845806",
    "platform": "tiktok",
    "date": "2026-06-02",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 53973,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 5230,
      "commissionFee": 12082,
      "affiliateCommission": 0,
      "voucherXtra": 3580,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 23892,
    "cogsTotal": 0
  },
  {
    "id": "584312844189009568",
    "platform": "tiktok",
    "date": "2026-06-02",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 55712,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4475,
      "commissionFee": 12082,
      "affiliateCommission": 0,
      "voucherXtra": 3580,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 23137,
    "cogsTotal": 0
  },
  {
    "id": "584312717404898317",
    "platform": "tiktok",
    "date": "2026-06-02",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 52792,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4325,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22463,
    "cogsTotal": 0
  },
  {
    "id": "584305843260851606",
    "platform": "tiktok",
    "date": "2026-06-01",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 52792,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4325,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 8650,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 31113,
    "cogsTotal": 0
  },
  {
    "id": "584302705530471861",
    "platform": "tiktok",
    "date": "2026-06-01",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 52792,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4325,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22463,
    "cogsTotal": 0
  },
  {
    "id": "584299504885990804",
    "platform": "tiktok",
    "date": "2026-06-01",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 57965,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4450,
      "commissionFee": 12015,
      "affiliateCommission": 0,
      "voucherXtra": 3560,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 23025,
    "cogsTotal": 0
  },
  {
    "id": "584289425147462935",
    "platform": "tiktok",
    "date": "2026-05-31",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 59580,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4200,
      "commissionFee": 11340,
      "affiliateCommission": 0,
      "voucherXtra": 3360,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 21900,
    "cogsTotal": 0
  },
  {
    "id": "584436502187639985",
    "platform": "tiktok",
    "date": "2026-06-08",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584427595636114832",
    "platform": "tiktok",
    "date": "2026-06-08",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584424848798877282",
    "platform": "tiktok",
    "date": "2026-06-08",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584305693337093178",
    "platform": "tiktok",
    "date": "2026-06-01",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 53484,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4325,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 7958,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 30421,
    "cogsTotal": 0
  },
  {
    "id": "584305559317677609",
    "platform": "tiktok",
    "date": "2026-06-01",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 57630,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4700,
      "commissionFee": 12690,
      "affiliateCommission": 0,
      "voucherXtra": 3760,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 24150,
    "cogsTotal": 0
  },
  {
    "id": "584300236010850127",
    "platform": "tiktok",
    "date": "2026-06-01",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 56252,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4325,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22463,
    "cogsTotal": 0
  },
  {
    "id": "584291412815218623",
    "platform": "tiktok",
    "date": "2026-05-31",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 63305,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4450,
      "commissionFee": 12015,
      "affiliateCommission": 0,
      "voucherXtra": 3560,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 23025,
    "cogsTotal": 0
  },
  {
    "id": "584285198644578141",
    "platform": "tiktok",
    "date": "2026-05-31",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 55694,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 5376,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 27726,
    "cogsTotal": 0
  },
  {
    "id": "584277873596073416",
    "platform": "tiktok",
    "date": "2026-05-30",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 52792,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4325,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 8650,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 31113,
    "cogsTotal": 0
  },
  {
    "id": "584275493331175044",
    "platform": "tiktok",
    "date": "2026-05-30",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 58781,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4700,
      "commissionFee": 12690,
      "affiliateCommission": 0,
      "voucherXtra": 3760,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 24150,
    "cogsTotal": 0
  },
  {
    "id": "584274149181982127",
    "platform": "tiktok",
    "date": "2026-05-30",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 51290,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4450,
      "commissionFee": 12015,
      "affiliateCommission": 0,
      "voucherXtra": 3560,
      "orderProcessingFee": 3000,
      "taxWithheld": 12015,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 35040,
    "cogsTotal": 0
  },
  {
    "id": "584268010394846392",
    "platform": "tiktok",
    "date": "2026-05-30",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 77557,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 5750,
      "commissionFee": 15525,
      "affiliateCommission": 0,
      "voucherXtra": 4600,
      "orderProcessingFee": 3000,
      "taxWithheld": 5118,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 33993,
    "cogsTotal": 0
  },
  {
    "id": "584261416871430106",
    "platform": "tiktok",
    "date": "2026-05-29",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 52792,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4325,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 8650,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 31113,
    "cogsTotal": 0
  },
  {
    "id": "584241044617266692",
    "platform": "tiktok",
    "date": "2026-05-28",
    "status": "returned",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 3000,
    "cogsTotal": 0,
    "returnReason": "Change of mind"
  },
  {
    "id": "584413530701793029",
    "platform": "tiktok",
    "date": "2026-06-07",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584260776037091136",
    "platform": "tiktok",
    "date": "2026-05-29",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 51201,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584258664178550685",
    "platform": "tiktok",
    "date": "2026-05-29",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 59522,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4450,
      "commissionFee": 12015,
      "affiliateCommission": 0,
      "voucherXtra": 3560,
      "orderProcessingFee": 3000,
      "taxWithheld": 3783,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 26808,
    "cogsTotal": 0
  },
  {
    "id": "584256559093286381",
    "platform": "tiktok",
    "date": "2026-05-29",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 58047,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 5855,
      "commissionFee": 12082,
      "affiliateCommission": 0,
      "voucherXtra": 3580,
      "orderProcessingFee": 3000,
      "taxWithheld": 4251,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 28768,
    "cogsTotal": 0
  },
  {
    "id": "584252338261493156",
    "platform": "tiktok",
    "date": "2026-05-29",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61442,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4325,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22463,
    "cogsTotal": 0
  },
  {
    "id": "584247471314470375",
    "platform": "tiktok",
    "date": "2026-05-29",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 59661,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4450,
      "commissionFee": 12015,
      "affiliateCommission": 0,
      "voucherXtra": 3560,
      "orderProcessingFee": 3000,
      "taxWithheld": 3644,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 26669,
    "cogsTotal": 0
  },
  {
    "id": "584271834226263870",
    "platform": "tiktok",
    "date": "2026-05-30",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 49764,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4325,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 11678,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 34141,
    "cogsTotal": 0
  },
  {
    "id": "584260874406561627",
    "platform": "tiktok",
    "date": "2026-05-29",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 56780,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4900,
      "commissionFee": 13230,
      "affiliateCommission": 0,
      "voucherXtra": 3920,
      "orderProcessingFee": 3000,
      "taxWithheld": 13230,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 38280,
    "cogsTotal": 0
  },
  {
    "id": "584252914215454275",
    "platform": "tiktok",
    "date": "2026-05-29",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 59952,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 11408,
      "affiliateCommission": 0,
      "voucherXtra": 3380,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22013,
    "cogsTotal": 0
  },
  {
    "id": "584252002843198594",
    "platform": "tiktok",
    "date": "2026-05-29",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61442,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4325,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22463,
    "cogsTotal": 0
  },
  {
    "id": "584249582907786362",
    "platform": "tiktok",
    "date": "2026-05-29",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61442,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4325,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22463,
    "cogsTotal": 0
  },
  {
    "id": "584231943358350656",
    "platform": "tiktok",
    "date": "2026-05-28",
    "status": "returned",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 3000,
    "cogsTotal": 0,
    "returnReason": "Item too big or too small"
  },
  {
    "id": "584230348743345257",
    "platform": "tiktok",
    "date": "2026-05-27",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 59952,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 11408,
      "affiliateCommission": 0,
      "voucherXtra": 3380,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22013,
    "cogsTotal": 0
  },
  {
    "id": "584224576880674716",
    "platform": "tiktok",
    "date": "2026-05-27",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61070,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584224444338046825",
    "platform": "tiktok",
    "date": "2026-05-27",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 54435,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4325,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 7007,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 29470,
    "cogsTotal": 0
  },
  {
    "id": "584379571507922750",
    "platform": "tiktok",
    "date": "2026-06-05",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584379274145596691",
    "platform": "tiktok",
    "date": "2026-06-05",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584378971215070483",
    "platform": "tiktok",
    "date": "2026-06-05",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584367743674057736",
    "platform": "tiktok",
    "date": "2026-06-05",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584254456974574830",
    "platform": "tiktok",
    "date": "2026-05-29",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 54801,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 6269,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 28619,
    "cogsTotal": 0
  },
  {
    "id": "584252902946932005",
    "platform": "tiktok",
    "date": "2026-05-29",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 57333,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4325,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 4109,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 26572,
    "cogsTotal": 0
  },
  {
    "id": "584249358255097055",
    "platform": "tiktok",
    "date": "2026-05-29",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 123390,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 8482,
      "commissionFee": 22903,
      "affiliateCommission": 0,
      "voucherXtra": 6786,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 41171,
    "cogsTotal": 0
  },
  {
    "id": "584249017255757129",
    "platform": "tiktok",
    "date": "2026-05-29",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 49764,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4325,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22463,
    "cogsTotal": 0
  },
  {
    "id": "584245646794393443",
    "platform": "tiktok",
    "date": "2026-05-28",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 59570,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 1500,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 23850,
    "cogsTotal": 0
  },
  {
    "id": "584236579005891588",
    "platform": "tiktok",
    "date": "2026-05-28",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61070,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584235629373392045",
    "platform": "tiktok",
    "date": "2026-05-28",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 48989,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 5100,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 23238,
    "cogsTotal": 0
  },
  {
    "id": "584231681465091915",
    "platform": "tiktok",
    "date": "2026-05-28",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 63305,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4450,
      "commissionFee": 12015,
      "affiliateCommission": 0,
      "voucherXtra": 3560,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 23025,
    "cogsTotal": 0
  },
  {
    "id": "584353978234930760",
    "platform": "tiktok",
    "date": "2026-06-04",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584352455463371902",
    "platform": "tiktok",
    "date": "2026-06-04",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584237258501359272",
    "platform": "tiktok",
    "date": "2026-05-28",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 59952,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4225,
      "commissionFee": 11408,
      "affiliateCommission": 0,
      "voucherXtra": 3380,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22013,
    "cogsTotal": 0
  },
  {
    "id": "584237125726471965",
    "platform": "tiktok",
    "date": "2026-05-28",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 50699,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4325,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22463,
    "cogsTotal": 0
  },
  {
    "id": "584236579487320050",
    "platform": "tiktok",
    "date": "2026-05-28",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 57117,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4325,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 4325,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 26788,
    "cogsTotal": 0
  },
  {
    "id": "584225007441839487",
    "platform": "tiktok",
    "date": "2026-05-27",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 49593,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 5080,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 23218,
    "cogsTotal": 0
  },
  {
    "id": "584220246856468348",
    "platform": "tiktok",
    "date": "2026-05-27",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 49764,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4325,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22463,
    "cogsTotal": 0
  },
  {
    "id": "584235916606605144",
    "platform": "tiktok",
    "date": "2026-05-28",
    "status": "returned",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 5100,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 8100,
    "cogsTotal": 0,
    "returnReason": "Change of mind"
  },
  {
    "id": "584224063800575693",
    "platform": "tiktok",
    "date": "2026-05-27",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 54003,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4325,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 7439,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 29902,
    "cogsTotal": 0
  },
  {
    "id": "584223320247207073",
    "platform": "tiktok",
    "date": "2026-05-27",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 119590,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 8228,
      "commissionFee": 22214,
      "affiliateCommission": 0,
      "voucherXtra": 6582,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 40024,
    "cogsTotal": 0
  },
  {
    "id": "584223145355478860",
    "platform": "tiktok",
    "date": "2026-05-27",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 59580,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4200,
      "commissionFee": 11340,
      "affiliateCommission": 0,
      "voucherXtra": 3360,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 21900,
    "cogsTotal": 0
  },
  {
    "id": "584218987353572924",
    "platform": "tiktok",
    "date": "2026-05-27",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 66069,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 5575,
      "commissionFee": 15052,
      "affiliateCommission": 0,
      "voucherXtra": 4460,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 28087,
    "cogsTotal": 0
  },
  {
    "id": "584218638557939438",
    "platform": "tiktok",
    "date": "2026-05-27",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61070,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584218357472003645",
    "platform": "tiktok",
    "date": "2026-05-27",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 49764,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4325,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22463,
    "cogsTotal": 0
  },
  {
    "id": "584218253909001231",
    "platform": "tiktok",
    "date": "2026-05-27",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 54781,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4325,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22463,
    "cogsTotal": 0
  },
  {
    "id": "584217538765620546",
    "platform": "tiktok",
    "date": "2026-05-27",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 53268,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4325,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22463,
    "cogsTotal": 0
  },
  {
    "id": "584204775103628864",
    "platform": "tiktok",
    "date": "2026-05-26",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 63305,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4450,
      "commissionFee": 12015,
      "affiliateCommission": 0,
      "voucherXtra": 3560,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 23025,
    "cogsTotal": 0
  },
  {
    "id": "584196840915765060",
    "platform": "tiktok",
    "date": "2026-05-25",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61070,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4300,
      "commissionFee": 11610,
      "affiliateCommission": 0,
      "voucherXtra": 3440,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22350,
    "cogsTotal": 0
  },
  {
    "id": "584195085803423656",
    "platform": "tiktok",
    "date": "2026-05-25",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 56780,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4325,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 4662,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 27125,
    "cogsTotal": 0
  },
  {
    "id": "584182245043897541",
    "platform": "tiktok",
    "date": "2026-05-24",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 47469,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 5765,
      "commissionFee": 11340,
      "affiliateCommission": 0,
      "voucherXtra": 3360,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 23465,
    "cogsTotal": 0
  },
  {
    "id": "584325467667334829",
    "platform": "tiktok",
    "date": "2026-06-02",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584204398019053552",
    "platform": "tiktok",
    "date": "2026-05-26",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61442,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4325,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22463,
    "cogsTotal": 0
  },
  {
    "id": "584204105354610134",
    "platform": "tiktok",
    "date": "2026-05-26",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 57122,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4325,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 4320,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 26783,
    "cogsTotal": 0
  },
  {
    "id": "584203724114986453",
    "platform": "tiktok",
    "date": "2026-05-26",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61442,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4325,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22463,
    "cogsTotal": 0
  },
  {
    "id": "584203742634214798",
    "platform": "tiktok",
    "date": "2026-05-26",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 51595,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4475,
      "commissionFee": 12082,
      "affiliateCommission": 0,
      "voucherXtra": 3580,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 23137,
    "cogsTotal": 0
  },
  {
    "id": "584203648444368408",
    "platform": "tiktok",
    "date": "2026-05-26",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 51595,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4475,
      "commissionFee": 12082,
      "affiliateCommission": 0,
      "voucherXtra": 3580,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 23137,
    "cogsTotal": 0
  },
  {
    "id": "584203508855244339",
    "platform": "tiktok",
    "date": "2026-05-26",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 61442,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4325,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22463,
    "cogsTotal": 0
  },
  {
    "id": "584203164956394877",
    "platform": "tiktok",
    "date": "2026-05-26",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 51595,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4475,
      "commissionFee": 12082,
      "affiliateCommission": 0,
      "voucherXtra": 3580,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 23137,
    "cogsTotal": 0
  },
  {
    "id": "584202473368093760",
    "platform": "tiktok",
    "date": "2026-05-26",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 58462,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4125,
      "commissionFee": 11138,
      "affiliateCommission": 0,
      "voucherXtra": 3300,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 21563,
    "cogsTotal": 0
  },
  {
    "id": "584189414293538439",
    "platform": "tiktok",
    "date": "2026-05-25",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 78205,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 5450,
      "commissionFee": 14715,
      "affiliateCommission": 0,
      "voucherXtra": 4360,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 27525,
    "cogsTotal": 0
  },
  {
    "id": "584188466423826166",
    "platform": "tiktok",
    "date": "2026-05-25",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 86028,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 5975,
      "commissionFee": 16132,
      "affiliateCommission": 0,
      "voucherXtra": 4780,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 29887,
    "cogsTotal": 0
  },
  {
    "id": "584186189427672244",
    "platform": "tiktok",
    "date": "2026-05-25",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 63305,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4450,
      "commissionFee": 12015,
      "affiliateCommission": 0,
      "voucherXtra": 3560,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 23025,
    "cogsTotal": 0
  },
  {
    "id": "584160304935896366",
    "platform": "tiktok",
    "date": "2026-05-23",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 34800,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 34800,
    "cogsTotal": 0
  },
  {
    "id": "584192309225686271",
    "platform": "tiktok",
    "date": "2026-05-25",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 56922,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4325,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 4520,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 26983,
    "cogsTotal": 0
  },
  {
    "id": "584191402734290740",
    "platform": "tiktok",
    "date": "2026-05-25",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 49764,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4325,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22463,
    "cogsTotal": 0
  },
  {
    "id": "584188567263872212",
    "platform": "tiktok",
    "date": "2026-05-25",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 51927,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4325,
      "commissionFee": 11678,
      "affiliateCommission": 0,
      "voucherXtra": 3460,
      "orderProcessingFee": 3000,
      "taxWithheld": 9515,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 31978,
    "cogsTotal": 0
  },
  {
    "id": "584188570266862995",
    "platform": "tiktok",
    "date": "2026-05-25",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 63635,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 6955,
      "commissionFee": 15052,
      "affiliateCommission": 0,
      "voucherXtra": 4460,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 29467,
    "cogsTotal": 0
  },
  {
    "id": "584181575160530769",
    "platform": "tiktok",
    "date": "2026-05-24",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 63305,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4450,
      "commissionFee": 12015,
      "affiliateCommission": 0,
      "voucherXtra": 3560,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 23025,
    "cogsTotal": 0
  },
  {
    "id": "584179952615457841",
    "platform": "tiktok",
    "date": "2026-05-24",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 62188,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4375,
      "commissionFee": 11812,
      "affiliateCommission": 0,
      "voucherXtra": 3500,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22687,
    "cogsTotal": 0
  },
  {
    "id": "584176233285977304",
    "platform": "tiktok",
    "date": "2026-05-24",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 59580,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4200,
      "commissionFee": 11340,
      "affiliateCommission": 0,
      "voucherXtra": 3360,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 21900,
    "cogsTotal": 0
  },
  {
    "id": "584173927425410997",
    "platform": "tiktok",
    "date": "2026-05-24",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 59580,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4200,
      "commissionFee": 11340,
      "affiliateCommission": 0,
      "voucherXtra": 3360,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 21900,
    "cogsTotal": 0
  },
  {
    "id": "584172978175903465",
    "platform": "tiktok",
    "date": "2026-05-24",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 62188,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4375,
      "commissionFee": 11812,
      "affiliateCommission": 0,
      "voucherXtra": 3500,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22687,
    "cogsTotal": 0
  },
  {
    "id": "584171357070722726",
    "platform": "tiktok",
    "date": "2026-05-24",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 59580,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4200,
      "commissionFee": 11340,
      "affiliateCommission": 0,
      "voucherXtra": 3360,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 21900,
    "cogsTotal": 0
  },
  {
    "id": "584169651161171163",
    "platform": "tiktok",
    "date": "2026-05-24",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 62188,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4375,
      "commissionFee": 11812,
      "affiliateCommission": 0,
      "voucherXtra": 3500,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22687,
    "cogsTotal": 0
  },
  {
    "id": "584169072621815318",
    "platform": "tiktok",
    "date": "2026-05-24",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 58410,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4450,
      "commissionFee": 12015,
      "affiliateCommission": 0,
      "voucherXtra": 3560,
      "orderProcessingFee": 3000,
      "taxWithheld": 4895,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 27920,
    "cogsTotal": 0
  },
  {
    "id": "584168371943409449",
    "platform": "tiktok",
    "date": "2026-05-24",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 59580,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4200,
      "commissionFee": 11340,
      "affiliateCommission": 0,
      "voucherXtra": 3360,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 21900,
    "cogsTotal": 0
  },
  {
    "id": "584161943939417996",
    "platform": "tiktok",
    "date": "2026-05-23",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 62188,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4375,
      "commissionFee": 11812,
      "affiliateCommission": 0,
      "voucherXtra": 3500,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22687,
    "cogsTotal": 0
  },
  {
    "id": "584161905631725443",
    "platform": "tiktok",
    "date": "2026-05-23",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 55059,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4125,
      "commissionFee": 11138,
      "affiliateCommission": 0,
      "voucherXtra": 3300,
      "orderProcessingFee": 3000,
      "taxWithheld": 3403,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 24966,
    "cogsTotal": 0
  },
  {
    "id": "584161414875153739",
    "platform": "tiktok",
    "date": "2026-05-23",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 51642,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4200,
      "commissionFee": 11340,
      "affiliateCommission": 0,
      "voucherXtra": 3360,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 21900,
    "cogsTotal": 0
  },
  {
    "id": "584157957974820764",
    "platform": "tiktok",
    "date": "2026-05-23",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 63305,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4450,
      "commissionFee": 12015,
      "affiliateCommission": 0,
      "voucherXtra": 3560,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 23025,
    "cogsTotal": 0
  },
  {
    "id": "584148929071318641",
    "platform": "tiktok",
    "date": "2026-05-22",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 50022,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4125,
      "commissionFee": 11138,
      "affiliateCommission": 0,
      "voucherXtra": 3300,
      "orderProcessingFee": 3000,
      "taxWithheld": 8440,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 30003,
    "cogsTotal": 0
  },
  {
    "id": "584109729135101538",
    "platform": "tiktok",
    "date": "2026-05-20",
    "status": "returned",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 60000,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 63000,
    "cogsTotal": 0,
    "returnReason": "Product doesn't match description"
  },
  {
    "id": "584095604165871366",
    "platform": "tiktok",
    "date": "2026-05-19",
    "status": "returned",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4100,
      "commissionFee": 11070,
      "affiliateCommission": 0,
      "voucherXtra": 3280,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 18450,
    "cogsTotal": 0,
    "returnReason": "Item too big or too small"
  },
  {
    "id": "584279406139443037",
    "platform": "tiktok",
    "date": "2026-05-31",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584176781728646374",
    "platform": "tiktok",
    "date": "2026-05-24",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 115157,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 7930,
      "commissionFee": 21411,
      "affiliateCommission": 0,
      "voucherXtra": 6344,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 38685,
    "cogsTotal": 0
  },
  {
    "id": "584174083752625635",
    "platform": "tiktok",
    "date": "2026-05-24",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 53116,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4200,
      "commissionFee": 11340,
      "affiliateCommission": 0,
      "voucherXtra": 3360,
      "orderProcessingFee": 3000,
      "taxWithheld": 6464,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 28364,
    "cogsTotal": 0
  },
  {
    "id": "584174108329739568",
    "platform": "tiktok",
    "date": "2026-05-24",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 52776,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4200,
      "commissionFee": 11340,
      "affiliateCommission": 0,
      "voucherXtra": 3360,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 21900,
    "cogsTotal": 0
  },
  {
    "id": "584173390252311597",
    "platform": "tiktok",
    "date": "2026-05-24",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 59389,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4450,
      "commissionFee": 12015,
      "affiliateCommission": 0,
      "voucherXtra": 3560,
      "orderProcessingFee": 3000,
      "taxWithheld": 3916,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 26941,
    "cogsTotal": 0
  },
  {
    "id": "584173314846458885",
    "platform": "tiktok",
    "date": "2026-05-24",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 47324,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4125,
      "commissionFee": 11138,
      "affiliateCommission": 0,
      "voucherXtra": 3300,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 21563,
    "cogsTotal": 0
  },
  {
    "id": "584173230032061526",
    "platform": "tiktok",
    "date": "2026-05-24",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 58462,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4125,
      "commissionFee": 11138,
      "affiliateCommission": 0,
      "voucherXtra": 3300,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 21563,
    "cogsTotal": 0
  },
  {
    "id": "584172779450959740",
    "platform": "tiktok",
    "date": "2026-05-24",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 50109,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4125,
      "commissionFee": 11138,
      "affiliateCommission": 0,
      "voucherXtra": 3300,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 21563,
    "cogsTotal": 0
  },
  {
    "id": "584170925243926465",
    "platform": "tiktok",
    "date": "2026-05-24",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 58462,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4125,
      "commissionFee": 11138,
      "affiliateCommission": 0,
      "voucherXtra": 3300,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 21563,
    "cogsTotal": 0
  },
  {
    "id": "584170983482623862",
    "platform": "tiktok",
    "date": "2026-05-24",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 59580,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4200,
      "commissionFee": 11340,
      "affiliateCommission": 0,
      "voucherXtra": 3360,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 21900,
    "cogsTotal": 0
  },
  {
    "id": "584168325532321223",
    "platform": "tiktok",
    "date": "2026-05-24",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 51320,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4375,
      "commissionFee": 11812,
      "affiliateCommission": 0,
      "voucherXtra": 3500,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 22687,
    "cogsTotal": 0
  },
  {
    "id": "584164496823780817",
    "platform": "tiktok",
    "date": "2026-05-23",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 47020,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4100,
      "commissionFee": 11070,
      "affiliateCommission": 0,
      "voucherXtra": 3280,
      "orderProcessingFee": 3000,
      "taxWithheld": 11070,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 32520,
    "cogsTotal": 0
  },
  {
    "id": "584163005901342220",
    "platform": "tiktok",
    "date": "2026-05-23",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 51779,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4125,
      "commissionFee": 11138,
      "affiliateCommission": 0,
      "voucherXtra": 3300,
      "orderProcessingFee": 3000,
      "taxWithheld": 6683,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 28246,
    "cogsTotal": 0
  },
  {
    "id": "584160825143821401",
    "platform": "tiktok",
    "date": "2026-05-23",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 63305,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4450,
      "commissionFee": 12015,
      "affiliateCommission": 0,
      "voucherXtra": 3560,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 23025,
    "cogsTotal": 0
  },
  {
    "id": "584160587786191887",
    "platform": "tiktok",
    "date": "2026-05-23",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 51290,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4450,
      "commissionFee": 12015,
      "affiliateCommission": 0,
      "voucherXtra": 3560,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 23025,
    "cogsTotal": 0
  },
  {
    "id": "584157564472493978",
    "platform": "tiktok",
    "date": "2026-05-23",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 74308,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 5450,
      "commissionFee": 14715,
      "affiliateCommission": 0,
      "voucherXtra": 4360,
      "orderProcessingFee": 3000,
      "taxWithheld": 3897,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 31422,
    "cogsTotal": 0
  },
  {
    "id": "584155443558254050",
    "platform": "tiktok",
    "date": "2026-05-23",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 53910,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4200,
      "commissionFee": 11340,
      "affiliateCommission": 0,
      "voucherXtra": 3360,
      "orderProcessingFee": 3000,
      "taxWithheld": 5670,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 27570,
    "cogsTotal": 0
  },
  {
    "id": "584154933833729500",
    "platform": "tiktok",
    "date": "2026-05-23",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 49349,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4125,
      "commissionFee": 11138,
      "affiliateCommission": 0,
      "voucherXtra": 3300,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 21563,
    "cogsTotal": 0
  },
  {
    "id": "584151901531833980",
    "platform": "tiktok",
    "date": "2026-05-22",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 58462,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4125,
      "commissionFee": 11138,
      "affiliateCommission": 0,
      "voucherXtra": 3300,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 21563,
    "cogsTotal": 0
  },
  {
    "id": "584149723483243534",
    "platform": "tiktok",
    "date": "2026-05-22",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 52114,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4125,
      "commissionFee": 11138,
      "affiliateCommission": 0,
      "voucherXtra": 3300,
      "orderProcessingFee": 3000,
      "taxWithheld": 6348,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 27911,
    "cogsTotal": 0
  },
  {
    "id": "584142670394131899",
    "platform": "tiktok",
    "date": "2026-05-22",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 58462,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4125,
      "commissionFee": 11138,
      "affiliateCommission": 0,
      "voucherXtra": 3300,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 21563,
    "cogsTotal": 0
  },
  {
    "id": "584141032724268374",
    "platform": "tiktok",
    "date": "2026-05-22",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 53289,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4125,
      "commissionFee": 11138,
      "affiliateCommission": 0,
      "voucherXtra": 3300,
      "orderProcessingFee": 3000,
      "taxWithheld": 5173,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 26736,
    "cogsTotal": 0
  },
  {
    "id": "584139154267538749",
    "platform": "tiktok",
    "date": "2026-05-22",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 51642,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4200,
      "commissionFee": 11340,
      "affiliateCommission": 0,
      "voucherXtra": 3360,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 21900,
    "cogsTotal": 0
  },
  {
    "id": "584108268202526600",
    "platform": "tiktok",
    "date": "2026-05-20",
    "status": "returned",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 3000,
    "cogsTotal": 0,
    "returnReason": "Change of mind"
  },
  {
    "id": "584278557479896756",
    "platform": "tiktok",
    "date": "2026-05-30",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584275477814871684",
    "platform": "tiktok",
    "date": "2026-05-30",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584271966265443425",
    "platform": "tiktok",
    "date": "2026-05-30",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584267803433011005",
    "platform": "tiktok",
    "date": "2026-05-30",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584266054901597656",
    "platform": "tiktok",
    "date": "2026-05-30",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584162543140439976",
    "platform": "tiktok",
    "date": "2026-05-23",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 51779,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4125,
      "commissionFee": 11138,
      "affiliateCommission": 0,
      "voucherXtra": 3300,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 21563,
    "cogsTotal": 0
  },
  {
    "id": "584150228773602880",
    "platform": "tiktok",
    "date": "2026-05-22",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 53924,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4125,
      "commissionFee": 11138,
      "affiliateCommission": 0,
      "voucherXtra": 3300,
      "orderProcessingFee": 3000,
      "taxWithheld": 4538,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 26101,
    "cogsTotal": 0
  },
  {
    "id": "584146661049664756",
    "platform": "tiktok",
    "date": "2026-05-22",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 58462,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4125,
      "commissionFee": 11138,
      "affiliateCommission": 0,
      "voucherXtra": 3300,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 21563,
    "cogsTotal": 0
  },
  {
    "id": "584107326824678826",
    "platform": "tiktok",
    "date": "2026-05-20",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 56900,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4200,
      "commissionFee": 11340,
      "affiliateCommission": 0,
      "voucherXtra": 3360,
      "orderProcessingFee": 3000,
      "taxWithheld": 2680,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 24580,
    "cogsTotal": 0
  },
  {
    "id": "584053956125820231",
    "platform": "tiktok",
    "date": "2026-05-16",
    "status": "returned",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 3000,
    "cogsTotal": 0,
    "returnReason": "Item too big or too small"
  },
  {
    "id": "584255008370361538",
    "platform": "tiktok",
    "date": "2026-05-29",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 0,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 0,
      "commissionFee": 0,
      "affiliateCommission": 0,
      "voucherXtra": 0,
      "orderProcessingFee": 0,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 0,
    "cogsTotal": 0
  },
  {
    "id": "584146046191371640",
    "platform": "tiktok",
    "date": "2026-05-22",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 58997,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4450,
      "commissionFee": 12015,
      "affiliateCommission": 0,
      "voucherXtra": 3560,
      "orderProcessingFee": 3000,
      "taxWithheld": 4308,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 27333,
    "cogsTotal": 0
  },
  {
    "id": "584141221501306082",
    "platform": "tiktok",
    "date": "2026-05-22",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 52209,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4200,
      "commissionFee": 11340,
      "affiliateCommission": 0,
      "voucherXtra": 3360,
      "orderProcessingFee": 3000,
      "taxWithheld": 7371,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 29271,
    "cogsTotal": 0
  },
  {
    "id": "584132109927876319",
    "platform": "tiktok",
    "date": "2026-05-21",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 82675,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 5750,
      "commissionFee": 15525,
      "affiliateCommission": 0,
      "voucherXtra": 4600,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 28875,
    "cogsTotal": 0
  },
  {
    "id": "584109354334520863",
    "platform": "tiktok",
    "date": "2026-05-20",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 78205,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 5450,
      "commissionFee": 14715,
      "affiliateCommission": 0,
      "voucherXtra": 4360,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 27525,
    "cogsTotal": 0
  },
  {
    "id": "584108856826889475",
    "platform": "tiktok",
    "date": "2026-05-20",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 51779,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4125,
      "commissionFee": 11138,
      "affiliateCommission": 0,
      "voucherXtra": 3300,
      "orderProcessingFee": 3000,
      "taxWithheld": 0,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 21563,
    "cogsTotal": 0
  },
  {
    "id": "584106352780347136",
    "platform": "tiktok",
    "date": "2026-05-19",
    "status": "success",
    "items": [],
    "gmv": 0,
    "settlementAmount": 56808,
    "fees": {
      "shippingFee": 0,
      "transactionFee": 4200,
      "commissionFee": 11340,
      "affiliateCommission": 0,
      "voucherXtra": 3360,
      "orderProcessingFee": 3000,
      "taxWithheld": 2772,
      "serviceFee": 0,
      "paymentFee": 0,
      "freeshipXtra": 0
    },
    "feesTotal": 24672,
    "cogsTotal": 0
  }
];


/* ============================================================================
 * 6. GLOBAL STATE — React Context + useReducer (Zustand-style store)
 * ========================================================================== */

/* Bảng phí sàn TikTok Shop — lấy đúng từ file ELLA_PROFIT_CALCULATOR_v2_1.xlsx
 * (sheet DASHBOARD, đối chiếu trực tiếp seller-vn.tiktok.com, cập nhật
 * 01/08/2026, hiệu lực từ 09/05/2026). Đây KHÔNG phải số ước tính — là bảng
 * phí thật người dùng đã tự tổng hợp. Dùng làm dữ liệu khởi tạo cho bảng
 * Config phí sàn; người dùng có thể thêm/sửa/xóa từng dòng. `valueType`:
 * "percent" (áp theo % giá bán) hoặc "fixed" (số tiền cố định/đơn). */
const DEFAULT_FEE_CONFIG = [
  { id: "fee-1", name: "Phí hoa hồng – Marketplace", valueType: "percent", value: 12.5, required: true, note: "Ngành thời trang/phụ kiện – mức mặc định" },
  { id: "fee-2", name: "Phí hoa hồng – Mall", valueType: "percent", value: 15.5, required: true, note: "Nhà bán hàng chính hãng" },
  { id: "fee-3", name: "Phí giao dịch (Transaction Fee)", valueType: "percent", value: 6, required: true, note: "Giảm còn 5% nếu chi GMV Max ≥4% GMV trong 30 ngày (từ ngày 31)" },
  { id: "fee-4", name: "Phí xử lý đơn hàng (Order Processing)", valueType: "fixed", value: 3000, required: true, note: "Cố định mỗi đơn giao thành công" },
  { id: "fee-5", name: "Voucher Extra Program", valueType: "percent", value: 3, required: false, note: "Tối đa 50.000đ/đơn — áp dụng nếu tham gia chương trình" },
  { id: "fee-6", name: "Hoàn phí vận chuyển (SFR)", valueType: "fixed", value: 4000, required: false, note: "Dao động 1.620đ–10.738đ/đơn thực tế — 4.000đ là mức ước tính" },
  { id: "fee-7", name: "Affiliate KOL/KOC – Tự nhiên", valueType: "percent", value: 10, required: false, note: "Mức tiêu chuẩn ngành phụ kiện — tự đàm phán theo từng KOL/KOC" },
  { id: "fee-8", name: "Affiliate KOL/KOC – Quảng cáo", valueType: "percent", value: 5, required: false, note: "Đơn hàng đến từ quảng cáo" },
  { id: "fee-9", name: "Phí hoa hồng ngành Thời trang (Marketplace)", valueType: "percent", value: 12.5, required: false, note: "Tùy ngành hàng cấp 3 cụ thể (10–12.5%) — xem Seller Center" },
  { id: "fee-10", name: "Chương trình tiết kiệm phí GD (GMV Max)", valueType: "percent", value: 5, required: false, note: "Giảm 1% nếu chi GMV Max ≥4% GMV trong 30 ngày liên tục" },
];

const initialState = {
  role: "admin", // 'admin' | 'staff'
  settings: {
    packagingFee: DEFAULT_PACKAGING_FEE,
    monthlyFixedCost: 0,
    otherVariableCost: 0,
  },
  feeConfig: DEFAULT_FEE_CONFIG,
  skuConfig: {
    [PLATFORM.TIKTOK]: REAL_TIKTOK_SKU_CONFIG,
    [PLATFORM.SHOPEE]: [],
  },
  orders: {
    [PLATFORM.TIKTOK]: REAL_TIKTOK_ORDERS,
    [PLATFORM.SHOPEE]: [],
  },
  ads: { [PLATFORM.TIKTOK]: 0, [PLATFORM.SHOPEE]: 0 },
  logs: [
    { id: uid("log"), time: new Date().toLocaleString("vi-VN"), type: "success", message: "Đã nạp dữ liệu thật TikTok Shop từ Seller Center: 604 đơn hàng và 166 SKU sản phẩm (gộp từ 3 file Batch Edit Template + file Đối soát/Income + file Trả hàng hoàn tiền)." },
    { id: uid("log"), time: new Date().toLocaleString("vi-VN"), type: "warning", message: "166 SKU đang có Giá vốn = 0đ (Seller Center không xuất giá vốn) — vào tab Sản phẩm, dùng công cụ 'Nhập nhanh' để cập nhật hàng loạt." },
    { id: uid("log"), time: new Date().toLocaleString("vi-VN"), type: "success", message: "Đã đọc phí sàn chi tiết từ file Đối soát: Tổng phí sàn ≈ 12.657.000đ (Transaction fee, Commission, Voucher Xtra, Order processing, Shipping fee, Thuế khấu trừ)." },
    { id: uid("log"), time: new Date().toLocaleString("vi-VN"), type: "info", message: "File 'Đơn trả hàng hoàn tiền' (42 dòng): khớp 28 đơn với dữ liệu hiện có, 7 yêu cầu chưa hoàn tất/bị từ chối được bỏ qua." },
    { id: uid("log"), time: new Date().toLocaleString("vi-VN"), type: "warning", message: "7 SKU đang tồn kho thấp (dưới 50, gồm 6 SKU đã hết hàng) — xem cảnh báo đỏ ở tab Sản phẩm > Phân tích sản phẩm." },
    { id: uid("log"), time: new Date().toLocaleString("vi-VN"), type: "info", message: "Shopee chưa có dữ liệu — import file đơn hàng Shopee ở tab Import khi sẵn sàng." },
  ],
  dateRange: { start: "", end: "" },
};

/**
 * Gộp 1 đơn hàng đã có trong state (từ lần import trước, vd: file "Tất cả
 * đơn hàng") với bản ghi mới parse được (vd: từ file Đối soát/Statement of
 * Account/Income) khi CÙNG Order ID.
 *
 * TRƯỚC ĐÂY: reducer dùng `existing.set(o.id, o)` — GHI ĐÈ TOÀN BỘ đơn hàng
 * cũ bằng bản ghi mới. File Đối soát không có cột SKU/Subtotal (chỉ có phí +
 * số tiền quyết toán cấp đơn hàng), nên nếu import file Đối soát SAU file
 * "Tất cả đơn hàng", toàn bộ items/GMV/giá vốn của đơn sẽ bị XÓA SẠCH dù phí
 * sàn đúng. Ngược lại nếu import file Đối soát TRƯỚC, sau đó import lại file
 * "Tất cả đơn hàng" (vốn không có cột phí), phí sàn vừa nhập sẽ bị XÓA MẤT về
 * 0đ — đúng hiện tượng "Tổng phí sàn tự nhiên về 0" mà người dùng gặp phải.
 *
 * QUY TẮC GỘP: với từng trường, ưu tiên giá trị "có ý nghĩa" (khác 0/rỗng) từ
 * CẢ 2 bản ghi thay vì để bản ghi mới ghi đè trắng trơn bản ghi cũ.
 */
function mergeOrderRecords(oldOrd, newOrd) {
  if (!oldOrd) return newOrd;
  const items = newOrd.items.length > 0 ? newOrd.items : oldOrd.items;
  const gmv = newOrd.gmv > 0 ? newOrd.gmv : oldOrd.gmv;
  const cogsTotal = items === newOrd.items ? newOrd.cogsTotal : oldOrd.cogsTotal;
  const fees = {};
  const feeFields = new Set([...Object.keys(oldOrd.fees || {}), ...Object.keys(newOrd.fees || {})]);
  feeFields.forEach((f) => {
    fees[f] = Math.max(oldOrd.fees?.[f] || 0, newOrd.fees?.[f] || 0);
  });
  const feesTotal = Object.values(fees).reduce((s, v) => s + v, 0);
  const settlementAmount = Math.max(oldOrd.settlementAmount || 0, newOrd.settlementAmount || 0);
  return {
    ...oldOrd,
    ...newOrd,
    items,
    gmv,
    cogsTotal,
    fees,
    feesTotal,
    settlementAmount,
    // Trạng thái đơn (hủy/hoàn) do file trả hàng cập nhật riêng qua IMPORT_RETURNS
    // — giữ nguyên trạng thái đã biết thay vì để file Đối soát (thường không có
    // cột trạng thái đáng tin cậy) ghi đè về "Thành công" mặc định.
    status: oldOrd.status,
    date: oldOrd.date || newOrd.date,
  };
}

function skuConfigToMap(list) {
  const m = new Map();
  list.forEach((e) => m.set(e.sku.toLowerCase(), e));
  return m;
}

function reducer(state, action) {
  switch (action.type) {
    case "IMPORT_SETTLEMENT": {
      const { platform, orders, warnings, fileName } = action.payload;
      const existing = new Map(state.orders[platform].map((o) => [o.id, o]));
      orders.forEach((o) => existing.set(o.id, mergeOrderRecords(existing.get(o.id), o)));
      const log = {
        id: uid("log"),
        time: new Date().toLocaleString("vi-VN"),
        type: warnings.length ? "warning" : "success",
        message: `[${platformLabel(platform)}] Import "${fileName}": ${orders.length} đơn hàng.${warnings.length ? " " + warnings.length + " cảnh báo." : ""}`,
        details: warnings,
      };
      return {
        ...state,
        orders: { ...state.orders, [platform]: Array.from(existing.values()) },
        logs: [log, ...state.logs].slice(0, 50),
      };
    }
    case "IMPORT_RETURNS": {
      const { platform, returns, fileName, warnings } = action.payload;
      const updated = state.orders[platform].map((o) => {
        if (returns.has(o.id)) {
          const r = returns.get(o.id);
          return { ...o, status: STATUS.RETURNED, returnReason: r.reason };
        }
        return o;
      });
      const matched = Array.from(returns.keys()).filter((id) => state.orders[platform].some((o) => o.id === id)).length;
      const log = {
        id: uid("log"),
        time: new Date().toLocaleString("vi-VN"),
        type: warnings && warnings.length ? "warning" : "success",
        message: `[${platformLabel(platform)}] Import trả hàng/hoàn tiền "${fileName}": khớp ${matched}/${returns.size} đơn.`,
        details: warnings,
      };
      return { ...state, orders: { ...state.orders, [platform]: updated }, logs: [log, ...state.logs].slice(0, 50) };
    }
    case "IMPORT_ADS": {
      const { platform, total, fileName } = action.payload;
      const log = {
        id: uid("log"),
        time: new Date().toLocaleString("vi-VN"),
        type: "success",
        message: `[${platformLabel(platform)}] Import quảng cáo "${fileName}": tổng chi phí ${formatVND(total)}.`,
      };
      return { ...state, ads: { ...state.ads, [platform]: total }, logs: [log, ...state.logs].slice(0, 50) };
    }
    case "IMPORT_SKU_CONFIG": {
      const { platform, entries, fileName } = action.payload;
      // TRƯỚC ĐÂY: ghi đè toàn bộ danh sách SKU mỗi lần import — nếu Seller
      // Center xuất catalog thành NHIỀU file batch-edit (thường xảy ra khi
      // catalog lớn, TikTok chia trang), import file thứ 2 sẽ XÓA MẤT toàn bộ
      // SKU của file thứ 1. Nay GỘP theo SKU (key), file import sau chỉ cập
      // nhật/thêm SKU trùng tên, không xóa các SKU đã có từ file trước.
      //
      // ĐỒNG THỜI: file Batch Edit Template của TikTok KHÔNG có cột Giá vốn
      // (chỉ có Tồn kho/Tên/Giá bán) nên parseSkuConfigRows luôn trả về
      // cogs=0 cho các dòng này. Nếu ghi đè thẳng, mỗi lần import lại file
      // tồn kho mới sẽ XÓA MẤT Giá vốn người dùng đã nhập tay trước đó. Vì
      // vậy khi gộp: nếu bản ghi mới có cogs=0 nhưng SKU đã tồn tại với
      // cogs>0, GIỮ NGUYÊN cogs cũ — chỉ cập nhật tồn kho/tên/giá bán mới.
      const merged = new Map(state.skuConfig[platform].map((e) => [e.sku.toLowerCase(), e]));
      entries.forEach((e) => {
        const key = e.sku.toLowerCase();
        const old = merged.get(key);
        const cogs = e.cogs > 0 ? e.cogs : (old ? old.cogs : 0);
        merged.set(key, { ...e, cogs });
      });
      const list = Array.from(merged.values());
      const log = {
        id: uid("log"),
        time: new Date().toLocaleString("vi-VN"),
        type: "success",
        message: `[${platformLabel(platform)}] Cập nhật Config SKU "${fileName}": ${entries.length} SKU (tổng hiện có: ${list.length} SKU).`,
      };
      return { ...state, skuConfig: { ...state.skuConfig, [platform]: list }, logs: [log, ...state.logs].slice(0, 50) };
    }
    case "UPDATE_SKU_COGS": {
      const { platform, sku, cogs } = action.payload;
      const list = state.skuConfig[platform].map((e) => (e.sku === sku ? { ...e, cogs } : e));
      return { ...state, skuConfig: { ...state.skuConfig, [platform]: list } };
    }
    case "BULK_UPDATE_SKU_COGS": {
      // Công cụ "Nhập nhanh": áp 1 mức giá vốn cho NHIỀU SKU được chọn cùng lúc.
      const { platform, skus, cogs } = action.payload;
      const skuSet = new Set(skus);
      const list = state.skuConfig[platform].map((e) => (skuSet.has(e.sku) ? { ...e, cogs } : e));
      const log = {
        id: uid("log"),
        time: new Date().toLocaleString("vi-VN"),
        type: "success",
        message: `[${platformLabel(platform)}] Nhập nhanh giá vốn ${formatVND(cogs)} cho ${skus.length} SKU.`,
      };
      return { ...state, skuConfig: { ...state.skuConfig, [platform]: list }, logs: [log, ...state.logs].slice(0, 50) };
    }
    case "ADD_LOG":
      return { ...state, logs: [action.payload, ...state.logs].slice(0, 50) };
    case "ADD_FEE_CONFIG_ITEM":
      return { ...state, feeConfig: [...state.feeConfig, action.payload] };
    case "UPDATE_FEE_CONFIG_ITEM": {
      const { id, patch } = action.payload;
      return { ...state, feeConfig: state.feeConfig.map((f) => (f.id === id ? { ...f, ...patch } : f)) };
    }
    case "DELETE_FEE_CONFIG_ITEM":
      return { ...state, feeConfig: state.feeConfig.filter((f) => f.id !== action.payload) };
    case "UPDATE_SETTINGS":
      return { ...state, settings: { ...state.settings, ...action.payload } };
    case "SET_ROLE":
      return { ...state, role: action.payload };
    case "SET_DATE_RANGE":
      return { ...state, dateRange: action.payload };
    case "CLEAR_SKU_CONFIG": {
      // Xóa riêng danh sách SKU & Tồn kho của 1 sàn — không đụng tới đơn
      // hàng/Ads/SKU của sàn còn lại.
      const { platform } = action.payload;
      const log = {
        id: uid("log"),
        time: new Date().toLocaleString("vi-VN"),
        type: "info",
        message: `Đã xóa toàn bộ SKU & tồn kho của ${platformLabel(platform)}.`,
      };
      return { ...state, skuConfig: { ...state.skuConfig, [platform]: [] }, logs: [log, ...state.logs].slice(0, 50) };
    }
    case "CLEAR_ORDERS": {
      // Xóa riêng đơn hàng + dữ liệu Đối soát/Quyết toán của 1 sàn (đơn hàng,
      // phí sàn, trạng thái trả hàng đều nằm chung trong orders[platform]).
      const { platform } = action.payload;
      const log = {
        id: uid("log"),
        time: new Date().toLocaleString("vi-VN"),
        type: "info",
        message: `Đã xóa toàn bộ đơn hàng & dữ liệu Đối soát của ${platformLabel(platform)}.`,
      };
      return { ...state, orders: { ...state.orders, [platform]: [] }, logs: [log, ...state.logs].slice(0, 50) };
    }
    case "CLEAR_ADS": {
      // Xóa dữ liệu chi phí quảng cáo (cả 2 sàn — Marketing hiển thị gộp).
      const log = {
        id: uid("log"),
        time: new Date().toLocaleString("vi-VN"),
        type: "info",
        message: "Đã xóa toàn bộ dữ liệu chi phí quảng cáo (Ads).",
      };
      return { ...state, ads: { [PLATFORM.TIKTOK]: 0, [PLATFORM.SHOPEE]: 0 }, logs: [log, ...state.logs].slice(0, 50) };
    }
    case "RESET_ALL":
      return {
        ...initialState,
        orders: { [PLATFORM.TIKTOK]: [], [PLATFORM.SHOPEE]: [] },
        skuConfig: { [PLATFORM.TIKTOK]: [], [PLATFORM.SHOPEE]: [] },
        logs: [{ id: uid("log"), time: new Date().toLocaleString("vi-VN"), type: "info", message: "Đã xóa toàn bộ dữ liệu (kể cả dữ liệu TikTok đã nhập ban đầu). Sẵn sàng import dữ liệu mới." }],
      };
    case "RESTORE_INITIAL_IMPORT":
      // Khôi phục lại đúng snapshot dữ liệu TikTok Shop đã nhập ban đầu từ 5 file
      // Seller Center (không phải dữ liệu giả lập) — hữu ích nếu người dùng đã
      // thử chỉnh sửa/xóa và muốn quay về điểm khởi đầu.
      return {
        ...state,
        orders: { [PLATFORM.TIKTOK]: REAL_TIKTOK_ORDERS, [PLATFORM.SHOPEE]: state.orders[PLATFORM.SHOPEE] },
        skuConfig: { [PLATFORM.TIKTOK]: REAL_TIKTOK_SKU_CONFIG, [PLATFORM.SHOPEE]: state.skuConfig[PLATFORM.SHOPEE] },
        logs: [{ id: uid("log"), time: new Date().toLocaleString("vi-VN"), type: "info", message: "Đã khôi phục snapshot dữ liệu TikTok Shop ban đầu (604 đơn, 166 SKU)." }, ...state.logs].slice(0, 50),
      };
    default:
      return state;
  }
}

const AppContext = createContext(null);
function useApp() {
  const ctx = useContext(AppContext);
  if (!ctx) throw new Error("useApp must be used within AppProvider");
  return ctx;
}

function AppProvider({ children }) {
  const [state, dispatch] = useReducer(reducer, initialState);

  // Computed orders (metrics applied) — recompute whenever raw orders/settings change
  const computedOrders = useMemo(() => {
    const out = {};
    for (const p of [PLATFORM.TIKTOK, PLATFORM.SHOPEE]) {
      out[p] = state.orders[p].map((o) => computeOrderMetrics(o, state.settings));
    }
    return out;
  }, [state.orders, state.settings]);

  const skuConfigMaps = useMemo(() => ({
    [PLATFORM.TIKTOK]: skuConfigToMap(state.skuConfig[PLATFORM.TIKTOK]),
    [PLATFORM.SHOPEE]: skuConfigToMap(state.skuConfig[PLATFORM.SHOPEE]),
  }), [state.skuConfig]);

  const value = { state, dispatch, computedOrders, skuConfigMaps };
  return <AppContext.Provider value={value}>{children}</AppContext.Provider>;
}

/* ============================================================================
 * 7. EXPORT REPORT (XLSX) — xuất báo cáo tổng hợp
 * ========================================================================== */

function exportReportToExcel(state, computedOrders, aggAll, aggByPlatform) {
  const wb = XLSX.utils.book_new();

  const summaryRows = [
    ["BÁO CÁO LỢI NHUẬN ELLA EYEWEAR", "", ""],
    ["Xuất lúc", new Date().toLocaleString("vi-VN"), ""],
    [],
    ["Chỉ số", "Giá trị", ""],
    ["Doanh thu gộp (GMV)", aggAll.gmv, ""],
    ["Doanh thu thực tế (Nett Revenue)", aggAll.nettRevenue, ""],
    ["Tổng giá vốn (COGS)", aggAll.cogsTotal, ""],
    ["Tổng phí sàn & thuế", aggAll.feesTotal, ""],
    ["Tổng phí đóng gói", aggAll.packagingTotal, ""],
    ["Lợi nhuận trước Ads", aggAll.profitBeforeAds, ""],
    ["Chi phí Ads", aggAll.adsFee, ""],
    ["Chi phí cố định", aggAll.fixedCost, ""],
    ["Lợi nhuận thực sau Ads", aggAll.profitAfterAds, ""],
    ["Biên lợi nhuận ròng (%)", Number(aggAll.margin.toFixed(1)), ""],
    ["Tổng số đơn", aggAll.orderCount, ""],
    ["Tỷ lệ hoàn/hủy (%)", Number(aggAll.returnRate.toFixed(1)), ""],
  ];
  const wsSummary = XLSX.utils.aoa_to_sheet(summaryRows);
  XLSX.utils.book_append_sheet(wb, wsSummary, "Tong Quan");

  const compareRows = [
    ["Sàn", "Số đơn", "Doanh thu thực", "Tổng phí sàn", "Tỷ lệ hoàn (%)", "Lợi nhuận ròng", "Biên LN (%)"],
  ];
  [PLATFORM.TIKTOK, PLATFORM.SHOPEE].forEach((p) => {
    const a = aggByPlatform[p];
    compareRows.push([platformLabel(p), a.orderCount, a.nettRevenue, a.feesTotal, Number(a.returnRate.toFixed(1)), a.profitAfterAds, Number(a.margin.toFixed(1))]);
  });
  const wsCompare = XLSX.utils.aoa_to_sheet(compareRows);
  XLSX.utils.book_append_sheet(wb, wsCompare, "So Sanh San");

  [PLATFORM.TIKTOK, PLATFORM.SHOPEE].forEach((p) => {
    const rows = [
      ["Order ID", "Ngày", "Trạng thái", "GMV", "Doanh thu thực", "COGS", "Phí sàn", "Phí đóng gói", "LN trước Ads"],
    ];
    computedOrders[p].forEach((o) => {
      rows.push([o.id, o.date, statusLabel(o.status), o.gmv, o.nettRevenue, o.cogsTotal, o.feesTotal, o.packagingFee, o.profitBeforeAds]);
    });
    const ws = XLSX.utils.aoa_to_sheet(rows);
    XLSX.utils.book_append_sheet(wb, ws, platformLabel(p).slice(0, 30));
  });

  XLSX.writeFile(wb, `ELLA_Bao_Cao_Loi_Nhuan_${todayIso()}.xlsx`);
}

function downloadSkuConfigTemplate(platform) {
  const wb = XLSX.utils.book_new();
  const rows = [
    ["Seller SKU", "Loại SKU (Kính/Hộp/Khăn)", "Tên sản phẩm", "Giá vốn"],
    ["ELLA-TR01-KINH", "Kính", "Gọng kính Titan TR01", 145000],
    ["ELLA-BOX-STD", "Hộp", "Hộp kính tiêu chuẩn", 18000],
    ["ELLA-CLOTH-STD", "Khăn", "Khăn lau kính microfiber", 4000],
  ];
  const ws = XLSX.utils.aoa_to_sheet(rows);
  XLSX.utils.book_append_sheet(wb, ws, "Config SKU");
  XLSX.writeFile(wb, `Mau_Config_SKU_${platformLabel(platform).replace(" ", "_")}.xlsx`);
}

/* ============================================================================
 * 8. SMALL UI PRIMITIVES
 * ========================================================================== */

function Badge({ children, tone = "gray" }) {
  const tones = {
    gray: "bg-slate-100 text-slate-600",
    green: "bg-emerald-50 text-emerald-700",
    red: "bg-rose-50 text-rose-700",
    amber: "bg-amber-50 text-amber-700",
    pink: "bg-pink-50 text-pink-700",
    navy: "bg-blue-50 text-blue-800",
  };
  return <span className={`inline-flex items-center px-2 py-0.5 rounded-full text-xs font-medium ${tones[tone]}`}>{children}</span>;
}

function Card({ children, className = "" }) {
  return <div className={`bg-white rounded-2xl border border-slate-200 shadow-sm ${className}`}>{children}</div>;
}

function Tooltip2({ text }) {
  const [show, setShow] = useState(false);
  return (
    <span className="relative inline-flex">
      <Info
        size={14}
        className="text-slate-400 cursor-help ml-1"
        onMouseEnter={() => setShow(true)}
        onMouseLeave={() => setShow(false)}
      />
      {show && (
        <span className="absolute z-20 bottom-full left-1/2 -translate-x-1/2 mb-2 w-56 p-2 rounded-lg bg-slate-900 text-white text-xs leading-snug shadow-lg">
          {text}
        </span>
      )}
    </span>
  );
}

/* ============================================================================
 * 9. HEADER + DATE RANGE FILTER
 * ========================================================================== */

const RANGE_PRESETS = [
  { key: "today", label: "Hôm nay" },
  { key: "yesterday", label: "Hôm qua" },
  { key: "7d", label: "7 ngày" },
  { key: "30d", label: "30 ngày" },
  { key: "all", label: "Tất cả" },
  { key: "custom", label: "Tùy chỉnh" },
];

function computePresetRange(key) {
  const fmt = (d) => d.toISOString().slice(0, 10);
  const now = new Date();
  if (key === "today") return { start: fmt(now), end: fmt(now) };
  if (key === "yesterday") {
    const y = new Date(now); y.setDate(y.getDate() - 1);
    return { start: fmt(y), end: fmt(y) };
  }
  if (key === "7d") {
    const s = new Date(now); s.setDate(s.getDate() - 6);
    return { start: fmt(s), end: fmt(now) };
  }
  if (key === "30d") {
    const s = new Date(now); s.setDate(s.getDate() - 29);
    return { start: fmt(s), end: fmt(now) };
  }
  return { start: "", end: "" };
}

/**
 * Sidebar điều hướng bên trái — thay cho menu ngang cũ. Có thể thu gọn
 * (collapse) chỉ còn icon. 4 phân hệ chính theo đúng yêu cầu: Tổng quan,
 * Sản phẩm, Marketing, Import.
 */
const SIDEBAR_SECTIONS = [
  { key: "overview", label: "Tổng quan", icon: LayoutDashboard },
  { key: "product", label: "Sản phẩm", icon: Package },
  { key: "marketing", label: "Marketing", icon: Megaphone },
  { key: "import", label: "Import", icon: FileSpreadsheet },
];

function Sidebar({ activeSection, setActiveSection, collapsed, setCollapsed }) {
  return (
    <aside
      className={`bg-white border-r border-slate-200 h-screen sticky top-0 flex flex-col shrink-0 transition-all duration-200 ${
        collapsed ? "w-[68px]" : "w-60"
      }`}
    >
      <div className="flex items-center gap-2 px-3.5 py-4 border-b border-slate-100">
        {/* Khung logo — thay src bằng file logo thật (SVG/PNG) khi có */}
        <div className="w-9 h-9 rounded-xl bg-blue-950 flex items-center justify-center text-white font-bold text-sm shrink-0 overflow-hidden">
          <img src="/logo.png" alt="ELLA Accents" className="w-full h-full object-cover hidden" onLoad={(e) => e.currentTarget.classList.remove("hidden")} onError={(e) => { e.currentTarget.classList.add("hidden"); }} />
          <span className="pointer-events-none">EA</span>
        </div>
        {!collapsed && (
          <div className="min-w-0">
            <div className="font-semibold text-blue-950 leading-tight truncate">ELLA Accents</div>
            <div className="text-[11px] text-slate-400 leading-tight truncate">Profit Control Center</div>
          </div>
        )}
      </div>

      <nav className="flex-1 py-3 px-2 space-y-1">
        {SIDEBAR_SECTIONS.map((s) => (
          <button
            key={s.key}
            onClick={() => setActiveSection(s.key)}
            title={collapsed ? s.label : undefined}
            className={`w-full flex items-center gap-2.5 px-2.5 py-2.5 rounded-xl text-sm font-medium transition ${
              activeSection === s.key ? "bg-blue-950 text-white shadow-sm" : "text-slate-500 hover:bg-slate-50 hover:text-blue-950"
            }`}
          >
            <s.icon size={17} className="shrink-0" />
            {!collapsed && <span className="truncate">{s.label}</span>}
          </button>
        ))}
      </nav>

      <button
        onClick={() => setCollapsed(!collapsed)}
        className="flex items-center justify-center gap-1.5 border-t border-slate-100 py-3 text-xs text-slate-400 hover:text-blue-950 transition"
      >
        {collapsed ? <ChevronRight size={15} /> : <><ChevronLeft size={15} /> Thu gọn</>}
      </button>
    </aside>
  );
}

/**
 * Thanh bộ lọc thời gian dùng chung — đặt BÊN TRONG từng phân hệ cần lọc
 * (Tổng quan, Sản phẩm, Marketing), KHÔNG đặt ở Import. Nút "Tùy chỉnh" mở
 * popover 2 ô chọn ngày xổ xuống ngay dưới, không đẩy các nút khác.
 * `onGoToImport`: điều hướng sang Mục Import khi bấm "Import dữ liệu".
 */
function DateFilterBar({ title, icon: Icon, onGoToImport }) {
  const { state, dispatch } = useApp();
  const [preset, setPreset] = useState("all");
  const [showCustom, setShowCustom] = useState(false);
  const popoverRef = useRef(null);

  useEffect(() => {
    function onClickOutside(e) {
      if (popoverRef.current && !popoverRef.current.contains(e.target)) setShowCustom(false);
    }
    if (showCustom) document.addEventListener("mousedown", onClickOutside);
    return () => document.removeEventListener("mousedown", onClickOutside);
  }, [showCustom]);

  function applyPreset(key) {
    setPreset(key);
    if (key === "custom") { setShowCustom((v) => !v); return; }
    setShowCustom(false);
    dispatch({ type: "SET_DATE_RANGE", payload: computePresetRange(key) });
  }

  return (
    <div className="flex flex-wrap items-center gap-3 justify-between mb-4">
      <h2 className="text-lg font-semibold text-blue-950 flex items-center gap-2">
        {Icon && <Icon size={19} />} {title}
      </h2>
      <div className="flex items-center gap-2 flex-wrap">
        <div className="flex items-center bg-slate-100 rounded-xl p-1 gap-1 relative">
          {RANGE_PRESETS.map((r) => (
            <div key={r.key} className="relative">
              <button
                onClick={() => applyPreset(r.key)}
                className={`px-3 py-1.5 rounded-lg text-xs font-medium transition whitespace-nowrap ${
                  preset === r.key ? "bg-blue-950 text-white shadow" : "text-slate-500 hover:text-blue-950"
                }`}
              >
                {r.label}
              </button>
              {r.key === "custom" && showCustom && (
                <div
                  ref={popoverRef}
                  className="absolute right-0 top-full mt-2 z-40 bg-white border border-slate-200 rounded-xl shadow-lg p-3 flex items-center gap-1.5 text-xs"
                >
                  <input
                    type="date"
                    className="border border-slate-200 rounded-lg px-2 py-1.5"
                    value={state.dateRange.start}
                    onChange={(e) => dispatch({ type: "SET_DATE_RANGE", payload: { ...state.dateRange, start: e.target.value } })}
                  />
                  <span className="text-slate-400">–</span>
                  <input
                    type="date"
                    className="border border-slate-200 rounded-lg px-2 py-1.5"
                    value={state.dateRange.end}
                    onChange={(e) => dispatch({ type: "SET_DATE_RANGE", payload: { ...state.dateRange, end: e.target.value } })}
                  />
                </div>
              )}
            </div>
          ))}
        </div>
        <button
          onClick={onGoToImport}
          className="flex items-center gap-1.5 bg-pink-600 hover:bg-pink-700 text-white text-sm font-medium px-3.5 py-2 rounded-xl transition shadow-sm"
        >
          <Upload size={15} /> Import dữ liệu
        </button>
      </div>
    </div>
  );
}

/* ============================================================================
 * 10. KPI CARDS
 * ========================================================================== */

function KPICard({ icon: Icon, label, value, tooltip, tone = "default", sub }) {
  const toneClass =
    tone === "positive" ? "text-emerald-600" : tone === "negative" ? "text-rose-600" : "text-blue-950";
  const iconBg =
    tone === "positive" ? "bg-emerald-50 text-emerald-600" : tone === "negative" ? "bg-rose-50 text-rose-600" : "bg-blue-50 text-blue-900";
  return (
    <Card className="p-4 flex flex-col gap-2">
      <div className="flex items-center justify-between">
        <div className="flex items-center text-xs font-medium text-slate-500">
          {label}
          {tooltip && <Tooltip2 text={tooltip} />}
        </div>
        <div className={`w-8 h-8 rounded-lg flex items-center justify-center ${iconBg}`}>
          <Icon size={16} />
        </div>
      </div>
      <div className={`text-xl font-semibold ${toneClass}`}>{value}</div>
      {sub && <div className="text-xs text-slate-400">{sub}</div>}
    </Card>
  );
}

function KPISection({ agg }) {
  const profitTone = agg.profitAfterAds >= 0 ? "positive" : "negative";
  return (
    <div className="grid grid-cols-2 md:grid-cols-4 xl:grid-cols-7 gap-3">
      <KPICard icon={DollarSign} label="Doanh thu gộp (GMV)" value={formatVND(agg.gmv)} tooltip="Tổng giá trị đơn hàng khởi tạo, bao gồm cả đơn hủy/hoàn." />
      <KPICard icon={TrendingUp} label="Doanh thu thực tế" value={formatVND(agg.nettRevenue)} tooltip="Đơn thành công = tiền quyết toán thực nhận. Đơn hủy/hoàn = 0đ." />
      <KPICard icon={Percent} label="Tổng phí sàn & thuế" value={formatVND(agg.feesTotal)} tooltip="Transaction fee + Commission + Affiliate + Voucher Xtra + Order processing + VAT/PIT..." />
      <KPICard icon={Megaphone} label="Tổng chi phí Ads" value={formatVND(agg.adsFee)} tooltip="Mặc định 0đ, cập nhật khi import báo cáo TikTok/Shopee Ads." />
      <KPICard icon={ShoppingBag} label="Lợi nhuận trước Ads" value={formatVND(agg.profitBeforeAds)} tooltip="Doanh thu thực − COGS − Phí sàn − Phí đóng gói (6.000đ/đơn)." />
      <KPICard icon={agg.profitAfterAds >= 0 ? TrendingUp : TrendingDown} tone={profitTone} label="Lợi nhuận thực sau Ads" value={formatVND(agg.profitAfterAds)} tooltip="Lợi nhuận trước Ads − Chi phí Ads − Chi phí cố định." />
      <KPICard icon={Percent} tone={profitTone} label="Biên lợi nhuận ròng" value={formatPercent(agg.margin)} tooltip="(Lợi nhuận thực sau Ads / Doanh thu thực tế) × 100" />
    </div>
  );
}

/* ============================================================================
 * 11. CHARTS
 * ========================================================================== */

function TrendChart({ data }) {
  return (
    <Card className="p-4">
      <div className="flex items-center justify-between mb-3">
        <h3 className="text-sm font-semibold text-blue-950">Xu hướng doanh thu & lợi nhuận</h3>
        <Badge tone="navy">Theo ngày</Badge>
      </div>
      <div className="h-72">
        <ResponsiveContainer width="100%" height="100%">
          <ComposedChart data={data} margin={{ top: 5, right: 10, left: -10, bottom: 0 }}>
            <CartesianGrid strokeDasharray="3 3" stroke="#eef1f5" />
            <XAxis dataKey="date" tick={{ fontSize: 11, fill: "#94a3b8" }} tickFormatter={(d) => d.slice(5)} />
            <YAxis tick={{ fontSize: 11, fill: "#94a3b8" }} tickFormatter={(v) => (v / 1000000).toFixed(1) + "tr"} />
            <Tooltip formatter={(v) => formatVND(v)} labelFormatter={(l) => "Ngày " + l} />
            <Legend wrapperStyle={{ fontSize: 12 }} />
            <Bar dataKey="nettRevenue" name="Doanh thu thực" fill={COLORS.navyLight} radius={[4, 4, 0, 0]} barSize={18} />
            <Bar dataKey="cost" name="Tổng chi phí" fill="#c7ccd4" radius={[4, 4, 0, 0]} barSize={18} />
            <Line type="monotone" dataKey="profit" name="Lợi nhuận ròng" stroke={COLORS.pink} strokeWidth={2.5} dot={{ r: 3 }} />
          </ComposedChart>
        </ResponsiveContainer>
      </div>
    </Card>
  );
}

function CostPieChart({ agg }) {
  const data = [
    { name: "Giá vốn (COGS)", value: agg.cogsTotal, color: COLORS.navy },
    { name: "Phí sàn & thuế", value: agg.feesTotal, color: COLORS.pink },
    { name: "Phí đóng gói", value: agg.packagingTotal, color: COLORS.amber },
    { name: "Phí Ads", value: agg.adsFee, color: COLORS.gray },
  ].filter((d) => d.value > 0);

  const total = data.reduce((s, d) => s + d.value, 0);

  return (
    <Card className="p-4">
      <div className="flex items-center justify-between mb-3">
        <h3 className="text-sm font-semibold text-blue-950">Cơ cấu chi phí</h3>
        <Badge tone="pink">Tổng {formatVND(total)}</Badge>
      </div>
      <div className="h-72">
        <ResponsiveContainer width="100%" height="100%">
          <PieChart>
            <Pie data={data} dataKey="value" nameKey="name" innerRadius={55} outerRadius={90} paddingAngle={2}>
              {data.map((d, i) => <Cell key={i} fill={d.color} />)}
            </Pie>
            <Tooltip formatter={(v) => formatVND(v)} />
            <Legend wrapperStyle={{ fontSize: 12 }} />
          </PieChart>
        </ResponsiveContainer>
      </div>
    </Card>
  );
}

/* ============================================================================
 * 12. PLATFORM COMPARISON TABLE
 * ========================================================================== */

function ComparisonTable({ aggByPlatform }) {
  return (
    <Card className="p-4 overflow-x-auto">
      <h3 className="text-sm font-semibold text-blue-950 mb-3">So sánh TikTok Shop vs Shopee</h3>
      <table className="w-full text-sm min-w-[640px]">
        <thead>
          <tr className="text-left text-xs text-slate-400 border-b border-slate-100">
            <th className="py-2 pr-2">Sàn</th>
            <th className="py-2 pr-2">Số đơn</th>
            <th className="py-2 pr-2">Doanh thu thực</th>
            <th className="py-2 pr-2">Tổng phí sàn</th>
            <th className="py-2 pr-2">Tỷ lệ hoàn (%)</th>
            <th className="py-2 pr-2">Lợi nhuận ròng</th>
            <th className="py-2 pr-2">Biên LN (%)</th>
          </tr>
        </thead>
        <tbody>
          {[PLATFORM.TIKTOK, PLATFORM.SHOPEE].map((p) => {
            const a = aggByPlatform[p];
            return (
              <tr key={p} className="border-b border-slate-50 last:border-0">
                <td className="py-2.5 pr-2 font-medium text-blue-950 flex items-center gap-1.5">
                  {p === PLATFORM.TIKTOK ? <Store size={14} /> : <ShoppingBag size={14} />}
                  {platformLabel(p)}
                </td>
                <td className="py-2.5 pr-2">{formatNumber(a.orderCount)}</td>
                <td className="py-2.5 pr-2">{formatVND(a.nettRevenue)}</td>
                <td className="py-2.5 pr-2">{formatVND(a.feesTotal)}</td>
                <td className="py-2.5 pr-2">
                  <Badge tone={a.returnRate > 15 ? "red" : a.returnRate > 8 ? "amber" : "green"}>{formatPercent(a.returnRate)}</Badge>
                </td>
                <td className={`py-2.5 pr-2 font-medium ${a.profitAfterAds >= 0 ? "text-emerald-600" : "text-rose-600"}`}>{formatVND(a.profitAfterAds)}</td>
                <td className="py-2.5 pr-2">{formatPercent(a.margin)}</td>
              </tr>
            );
          })}
        </tbody>
      </table>
    </Card>
  );
}

/* ============================================================================
 * 13. DASHBOARD TAB
 * ========================================================================== */

function DashboardTab() {
  const { state, computedOrders } = useApp();
  const { start, end } = state.dateRange;

  const filtered = useMemo(() => {
    const out = {};
    for (const p of [PLATFORM.TIKTOK, PLATFORM.SHOPEE]) {
      out[p] = computedOrders[p].filter((o) => isWithinRange(o.date, start, end));
    }
    return out;
  }, [computedOrders, start, end]);

  const allOrders = useMemo(() => [...filtered[PLATFORM.TIKTOK], ...filtered[PLATFORM.SHOPEE]], [filtered]);
  const totalAds = state.ads[PLATFORM.TIKTOK] + state.ads[PLATFORM.SHOPEE];
  const aggAll = useMemo(
    () => aggregateOrders(allOrders, totalAds, state.settings.monthlyFixedCost, state.settings.otherVariableCost),
    [allOrders, totalAds, state.settings]
  );
  const aggByPlatform = useMemo(() => {
    const out = {};
    for (const p of [PLATFORM.TIKTOK, PLATFORM.SHOPEE]) {
      const revShare = aggAll.nettRevenue > 0
        ? aggregateOrders(filtered[p]).nettRevenue / aggAll.nettRevenue
        : 0;
      out[p] = aggregateOrders(
        filtered[p],
        state.ads[p],
        state.settings.monthlyFixedCost * revShare,
        state.settings.otherVariableCost * revShare
      );
    }
    return out;
  }, [filtered, state.ads, state.settings, aggAll.nettRevenue]);

  const trendData = useMemo(() => aggregateByDate(allOrders), [allOrders]);

  const feesLookMissing = allOrders.length > 0 && aggAll.feesTotal === 0;

  return (
    <div className="space-y-5">
      {feesLookMissing && (
        <Card className="p-3.5 bg-amber-50 border-amber-100">
          <div className="flex gap-2 text-xs text-amber-800">
            <AlertCircle size={15} className="shrink-0 mt-0.5" />
            <div>
              Chưa phát hiện dữ liệu phí sàn chi tiết (Transaction fee / Commission / Affiliate...) trong các đơn hiện có — <b>Tổng phí sàn</b> đang tạm hiển thị 0đ. Import thêm file Đối soát/Statement of Account (Payment/Income) từ Seller Center ở tab Import Center để tính chính xác lợi nhuận sau phí.
            </div>
          </div>
        </Card>
      )}
      <KPISection agg={aggAll} />
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
        <TrendChart data={trendData} />
        <CostPieChart agg={aggAll} />
      </div>
      <ComparisonTable aggByPlatform={aggByPlatform} />
      {allOrders.length === 0 && (
        <Card className="p-8 text-center text-slate-400 text-sm">
          Không có đơn hàng nào trong khoảng thời gian đã chọn. Hãy import dữ liệu hoặc đổi bộ lọc thời gian.
        </Card>
      )}
    </div>
  );
}

/* ============================================================================
 * 14. PLATFORM DETAIL TAB (TikTok / Shopee)
 * ========================================================================== */

function PlatformDetailTab({ platform }) {
  const { state, computedOrders } = useApp();
  const { start, end } = state.dateRange;
  const [search, setSearch] = useState("");
  const [statusFilter, setStatusFilter] = useState("all");

  const orders = useMemo(() => {
    return computedOrders[platform]
      .filter((o) => isWithinRange(o.date, start, end))
      .filter((o) => (statusFilter === "all" ? true : o.status === statusFilter))
      .filter((o) => (search ? o.id.toLowerCase().includes(search.toLowerCase()) : true))
      .sort((a, b) => b.date.localeCompare(a.date));
  }, [computedOrders, platform, start, end, statusFilter, search]);

  const agg = useMemo(() => aggregateOrders(orders, state.ads[platform], state.settings.monthlyFixedCost, state.settings.otherVariableCost), [orders, state.ads, platform, state.settings]);

  return (
    <div className="space-y-5">
      <KPISection agg={agg} />
      <Card className="p-4">
        <div className="flex flex-wrap items-center gap-3 mb-3 justify-between">
          <h3 className="text-sm font-semibold text-blue-950 flex items-center gap-2">
            {platform === PLATFORM.TIKTOK ? <Store size={16} /> : <ShoppingBag size={16} />}
            Chi tiết đơn hàng {platformLabel(platform)}
          </h3>
          <div className="flex items-center gap-2">
            <div className="relative">
              <Search size={14} className="absolute left-2.5 top-2.5 text-slate-400" />
              <input
                value={search}
                onChange={(e) => setSearch(e.target.value)}
                placeholder="Tìm Order ID..."
                className="pl-8 pr-3 py-1.5 border border-slate-200 rounded-lg text-sm w-44"
              />
            </div>
            <select
              value={statusFilter}
              onChange={(e) => setStatusFilter(e.target.value)}
              className="border border-slate-200 rounded-lg text-sm px-2 py-1.5"
            >
              <option value="all">Tất cả trạng thái</option>
              <option value={STATUS.SUCCESS}>Thành công</option>
              <option value={STATUS.RETURNED}>Hoàn trả</option>
              <option value={STATUS.CANCELLED}>Đã hủy</option>
            </select>
          </div>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-sm min-w-[900px]">
            <thead>
              <tr className="text-left text-xs text-slate-400 border-b border-slate-100">
                <th className="py-2 pr-2">Order ID</th>
                <th className="py-2 pr-2">Ngày</th>
                <th className="py-2 pr-2">Trạng thái</th>
                <th className="py-2 pr-2">SKU</th>
                <th className="py-2 pr-2">GMV</th>
                <th className="py-2 pr-2">DT thực</th>
                <th className="py-2 pr-2">COGS</th>
                <th className="py-2 pr-2">Phí sàn</th>
                <th className="py-2 pr-2">Đóng gói</th>
                <th className="py-2 pr-2">LN trước Ads</th>
              </tr>
            </thead>
            <tbody>
              {orders.slice(0, 200).map((o) => (
                <tr key={o.id} className="border-b border-slate-50 last:border-0 hover:bg-slate-50">
                  <td className="py-2 pr-2 font-medium text-blue-950">{o.id}</td>
                  <td className="py-2 pr-2 text-slate-500">{o.date}</td>
                  <td className="py-2 pr-2">
                    <Badge tone={o.status === STATUS.SUCCESS ? "green" : o.status === STATUS.RETURNED ? "amber" : "red"}>
                      {statusLabel(o.status)}
                    </Badge>
                  </td>
                  <td className="py-2 pr-2 text-slate-500 max-w-[160px] truncate" title={o.items.map((i) => i.sellerSku).join(", ")}>
                    {o.items.map((i) => i.sellerSku).join(", ") || "—"}
                  </td>
                  <td className="py-2 pr-2">{formatVND(o.gmv)}</td>
                  <td className="py-2 pr-2">{formatVND(o.nettRevenue)}</td>
                  <td className="py-2 pr-2">{formatVND(o.cogsTotal)}</td>
                  <td className="py-2 pr-2">{formatVND(o.feesTotal)}</td>
                  <td className="py-2 pr-2">{formatVND(o.packagingFee)}</td>
                  <td className={`py-2 pr-2 font-medium ${o.profitBeforeAds >= 0 ? "text-emerald-600" : "text-rose-600"}`}>{formatVND(o.profitBeforeAds)}</td>
                </tr>
              ))}
            </tbody>
          </table>
          {orders.length === 0 && <div className="text-center text-slate-400 text-sm py-8">Chưa có đơn hàng phù hợp bộ lọc.</div>}
          {orders.length > 200 && <div className="text-center text-slate-400 text-xs py-3">Hiển thị 200/{orders.length} đơn gần nhất theo bộ lọc.</div>}
        </div>
      </Card>
    </div>
  );
}

/* ============================================================================
 * 15. IMPORT CENTER
 * ========================================================================== */

function DropZone({ title, description, icon: Icon, platformSelector = true, onFile, accentColor = "navy" }) {
  const [platform, setPlatform] = useState(PLATFORM.TIKTOK);
  const [dragOver, setDragOver] = useState(false);
  const [status, setStatus] = useState(null); // {type, message}
  const inputRef = useRef(null);

  async function handleFiles(files) {
    const file = files[0];
    if (!file) return;
    if (!/\.(xlsx|xls|csv)$/i.test(file.name)) {
      setStatus({ type: "error", message: "Chỉ hỗ trợ file .xlsx, .xls hoặc .csv" });
      return;
    }
    try {
      const result = await onFile(file, platform);
      setStatus({ type: result.ok === false ? "warning" : "success", message: result.message });
    } catch (err) {
      setStatus({ type: "error", message: "Lỗi đọc file: " + (err?.message || "không xác định") });
    }
  }

  const ring = accentColor === "pink" ? "hover:border-pink-400" : "hover:border-blue-400";

  return (
    <Card className="p-4 flex flex-col gap-3">
      <div className="flex items-start gap-3">
        <div className="w-9 h-9 rounded-xl bg-blue-50 text-blue-900 flex items-center justify-center shrink-0"><Icon size={17} /></div>
        <div>
          <div className="font-medium text-sm text-blue-950">{title}</div>
          <div className="text-xs text-slate-400">{description}</div>
        </div>
      </div>

      {platformSelector && (
        <div className="flex bg-slate-100 rounded-lg p-1 w-fit text-xs">
          {[PLATFORM.TIKTOK, PLATFORM.SHOPEE].map((p) => (
            <button
              key={p}
              onClick={() => setPlatform(p)}
              className={`px-3 py-1 rounded-md font-medium transition ${platform === p ? "bg-white shadow text-blue-950" : "text-slate-400"}`}
            >
              {platformLabel(p)}
            </button>
          ))}
        </div>
      )}

      <div
        onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
        onDragLeave={() => setDragOver(false)}
        onDrop={(e) => { e.preventDefault(); setDragOver(false); handleFiles(e.dataTransfer.files); }}
        onClick={() => inputRef.current?.click()}
        className={`border-2 border-dashed rounded-xl p-6 flex flex-col items-center justify-center text-center cursor-pointer transition ${
          dragOver ? "border-blue-500 bg-blue-50" : `border-slate-200 ${ring}`
        }`}
      >
        <Upload size={20} className="text-slate-400 mb-2" />
        <div className="text-xs text-slate-500">Kéo thả file vào đây hoặc <span className="text-blue-700 font-medium">chọn file</span></div>
        <div className="text-[11px] text-slate-300 mt-1">.xlsx .xls .csv</div>
        <input ref={inputRef} type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={(e) => handleFiles(e.target.files)} />
      </div>

      {status && (
        <div className={`flex items-start gap-2 text-xs p-2.5 rounded-lg ${
          status.type === "success" ? "bg-emerald-50 text-emerald-700" : status.type === "warning" ? "bg-amber-50 text-amber-700" : "bg-rose-50 text-rose-700"
        }`}>
          {status.type === "success" ? <CheckCircle2 size={14} className="shrink-0 mt-0.5" /> : <AlertCircle size={14} className="shrink-0 mt-0.5" />}
          <span>{status.message}</span>
        </div>
      )}
    </Card>
  );
}

/* ============================================================================
 * 15b. MARKETING TAB (ADS REPORT)
 * ========================================================================== */

/**
 * LƯU Ý QUAN TRỌNG: state.ads hiện chỉ lưu 1 số TỔNG chi phí Ads (đọc từ cột
 * Cost/Amount Spent của file Ads) — Seller Center/Ads Manager thường KHÔNG
 * xuất kèm cột "Doanh thu do quảng cáo mang lại" (GMV attributed) trong cùng
 * file cơ bản. Vì vậy "Doanh thu từ Ads" bên dưới là ƯỚC TÍNH dựa trên tổng
 * doanh thu thực của shop trong cùng khoảng thời gian đang lọc — không phải
 * số do TikTok/Shopee tự attribute cho riêng quảng cáo. Khi có file Ads xuất
 * kèm cột doanh thu/GMV thật, cần cập nhật lại phần này để dùng số chính xác.
 */
function MarketingTab() {
  const { state, computedOrders } = useApp();
  const { start, end } = state.dateRange;

  const revenueInRange = useMemo(() => {
    let sum = 0;
    for (const p of [PLATFORM.TIKTOK, PLATFORM.SHOPEE]) {
      for (const o of computedOrders[p]) {
        if (isWithinRange(o.date, start, end)) sum += o.nettRevenue;
      }
    }
    return sum;
  }, [computedOrders, start, end]);

  const adsCost = state.ads[PLATFORM.TIKTOK] + state.ads[PLATFORM.SHOPEE];
  const hasAdsData = adsCost > 0;
  const roas = adsCost > 0 ? revenueInRange / adsCost : 0;
  const cir = revenueInRange > 0 ? (adsCost / revenueInRange) * 100 : 0;

  let assessment = { tone: "gray", text: "Chưa có dữ liệu quảng cáo — import file Ads (TikTok Ads Manager/Shopee Ads) ở tab Import để xem đánh giá hiệu quả." };
  if (hasAdsData) {
    if (cir <= 15) assessment = { tone: "green", text: `CIR ${formatPercent(cir)} — Hiệu quả TỐT, chi phí quảng cáo đang chiếm tỷ trọng thấp so với doanh thu (ROAS ${roas.toFixed(2)}x).` };
    else if (cir <= 25) assessment = { tone: "amber", text: `CIR ${formatPercent(cir)} — Ở mức TRUNG BÌNH, nên theo dõi thêm và tối ưu targeting/creative (ROAS ${roas.toFixed(2)}x).` };
    else assessment = { tone: "red", text: `CIR ${formatPercent(cir)} — CAO, chi phí Ads đang chiếm tỷ trọng lớn trong doanh thu, cần rà soát lại hiệu quả chiến dịch (ROAS ${roas.toFixed(2)}x).` };
  }

  return (
    <div className="space-y-5">
      {!hasAdsData && (
        <Card className="p-3.5 bg-amber-50 border-amber-100">
          <div className="flex gap-2 text-xs text-amber-800">
            <AlertCircle size={15} className="shrink-0 mt-0.5" />
            Chưa có dữ liệu chi phí quảng cáo. Vào tab Import để tải lên file Ads TikTok/Shopee.
          </div>
        </Card>
      )}
      <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
        <KPICard icon={Megaphone} label="Tổng chi phí Ads" value={formatVND(adsCost)} tooltip="Tổng chi phí quảng cáo TikTok + Shopee trong kỳ đang lọc." />
        <KPICard icon={DollarSign} label="Doanh thu từ Ads (ước tính)" value={formatVND(revenueInRange)} tooltip="Ước tính theo tổng doanh thu thực của shop trong kỳ — chưa phải số attribution riêng cho Ads do file Ads chưa có cột doanh thu quy đổi." />
        <KPICard icon={Target} tone={roas >= 4 ? "positive" : roas > 0 && roas < 2 ? "negative" : "default"} label="ROAS" value={hasAdsData ? `${roas.toFixed(2)}x` : "—"} tooltip="Return on Ad Spend = Doanh thu / Chi phí Ads." />
        <KPICard icon={Percent} tone={cir > 25 ? "negative" : cir > 0 && cir <= 15 ? "positive" : "default"} label="CIR" value={hasAdsData ? formatPercent(cir) : "—"} tooltip="Cost Income Ratio = (Chi phí Ads / Doanh thu) × 100%." />
      </div>
      <Card className={`p-4 border ${assessment.tone === "green" ? "bg-emerald-50 border-emerald-100" : assessment.tone === "amber" ? "bg-amber-50 border-amber-100" : assessment.tone === "red" ? "bg-rose-50 border-rose-100" : "bg-slate-50 border-slate-100"}`}>
        <h3 className="text-sm font-semibold text-blue-950 mb-1.5 flex items-center gap-2"><ClipboardList size={15} /> Đánh giá & mô tả tình trạng</h3>
        <p className="text-xs text-slate-600 leading-relaxed">{assessment.text}</p>
      </Card>
    </div>
  );
}

/* ============================================================================
 * 15c. SECTION WRAPPERS (Sidebar section → sub-tabs + DateFilterBar)
 * ========================================================================== */

const OVERVIEW_SUBTABS = [
  { key: "overview", label: "Tổng quan" },
  { key: "tiktok", label: "TikTok Shop" },
  { key: "shopee", label: "Shopee" },
  { key: "config", label: "Config" },
];

function SubTabNav({ tabs, active, onChange }) {
  return (
    <div className="flex items-center gap-1 bg-white border border-slate-200 rounded-xl p-1 w-fit mb-4 flex-wrap">
      {tabs.map((t) => (
        <button
          key={t.key}
          onClick={() => onChange(t.key)}
          className={`px-3.5 py-1.5 rounded-lg text-xs font-medium transition ${
            active === t.key ? "bg-blue-950 text-white shadow" : "text-slate-500 hover:text-blue-950"
          }`}
        >
          {t.label}
        </button>
      ))}
    </div>
  );
}

function OverviewSection({ onGoToImport }) {
  const [sub, setSub] = useState("overview");
  return (
    <div>
      <DateFilterBar title="Tổng quan" icon={LayoutDashboard} onGoToImport={onGoToImport} />
      <SubTabNav tabs={OVERVIEW_SUBTABS} active={sub} onChange={setSub} />
      {sub === "overview" && <DashboardTab />}
      {sub === "tiktok" && <PlatformDetailTab platform={PLATFORM.TIKTOK} />}
      {sub === "shopee" && <PlatformDetailTab platform={PLATFORM.SHOPEE} />}
      {sub === "config" && <FeeConfigCenterTab />}
    </div>
  );
}

const PRODUCT_SUBTABS = [
  { key: "analysis", label: "Phân tích Sản phẩm" },
  { key: "tiktok", label: "SKU TikTok Shop" },
  { key: "shopee", label: "SKU Shopee" },
];

function ProductSection({ onGoToImport }) {
  const [sub, setSub] = useState("analysis");
  return (
    <div>
      <DateFilterBar title="Sản phẩm" icon={Package} onGoToImport={onGoToImport} />
      <SubTabNav tabs={PRODUCT_SUBTABS} active={sub} onChange={setSub} />
      {sub === "analysis" && <ProductAnalysisTab />}
      {sub === "tiktok" && <SkuManagementTable platform={PLATFORM.TIKTOK} />}
      {sub === "shopee" && <SkuManagementTable platform={PLATFORM.SHOPEE} />}
    </div>
  );
}

function MarketingSection({ onGoToImport }) {
  return (
    <div>
      <DateFilterBar title="Marketing" icon={Megaphone} onGoToImport={onGoToImport} />
      <MarketingTab />
    </div>
  );
}

function ImportCenterTab() {
  const { state, dispatch, skuConfigMaps } = useApp();

  const handleSettlement = useCallback(async (file, platform) => {
    const { headers, rows } = await readSpreadsheet(file);
    const { orders, warnings, ok } = parseSettlementRows(headers, rows, platform, skuConfigMaps[platform]);
    if (!ok) return { ok: false, message: warnings.join(" ") };
    dispatch({ type: "IMPORT_SETTLEMENT", payload: { platform, orders, warnings, fileName: file.name } });
    return { ok: true, message: `Upload thành công ${orders.length} đơn hàng (${platformLabel(platform)}).${warnings.length ? " Có " + warnings.length + " cảnh báo SKU, xem Nhật ký bên dưới." : ""}` };
  }, [dispatch, skuConfigMaps]);

  const handleReturns = useCallback(async (file, platform) => {
    const { headers, rows } = await readSpreadsheet(file);
    const { returns, warnings, ok } = parseReturnsRows(headers, rows);
    if (!ok) return { ok: false, message: warnings.join(" ") };
    dispatch({ type: "IMPORT_RETURNS", payload: { platform, returns, fileName: file.name, warnings } });
    return { ok: true, message: `Đã cập nhật trạng thái hoàn/trả cho ${returns.size} đơn (${platformLabel(platform)}).${warnings.length ? " " + warnings.join(" ") : ""}` };
  }, [dispatch]);

  const handleAds = useCallback(async (file, platform) => {
    const { headers, rows } = await readSpreadsheet(file);
    const { total, warnings, ok } = parseAdsRows(headers, rows);
    dispatch({ type: "IMPORT_ADS", payload: { platform, total, fileName: file.name } });
    if (!ok) return { ok: false, message: `Không nhận diện được cột chi phí — Ads giữ 0đ. ${warnings.join(" ")}` };
    return { ok: true, message: `Upload thành công. Tổng chi phí Ads ${platformLabel(platform)}: ${formatVND(total)}.` };
  }, [dispatch]);

  const handleConfig = useCallback(async (file, platform) => {
    const { headers, rows } = await readSpreadsheet(file);
    const { entries, warnings, ok } = parseSkuConfigRows(headers, rows);
    if (!ok) return { ok: false, message: warnings.join(" ") };
    dispatch({ type: "IMPORT_SKU_CONFIG", payload: { platform, entries, fileName: file.name } });
    return { ok: true, message: `Đã ghi đè Config SKU ${platformLabel(platform)} với ${entries.length} dòng.` };
  }, [dispatch]);

  return (
    <div className="space-y-5">
      <Card className="p-4 bg-blue-50/40 border-blue-100">
        <div className="flex gap-2 text-sm text-blue-900">
          <Info size={16} className="shrink-0 mt-0.5" />
          <div>
            Toàn bộ file được đọc và xử lý <b>ngay trên trình duyệt</b> (client-side), không upload lên máy chủ nào khác.
            Hệ thống tự nhận diện tiêu đề cột theo tên tiếng Anh/tiếng Việt phổ biến của TikTok Shop & Shopee. Nếu file có cấu trúc khác thường, hệ thống sẽ báo cảnh báo cụ thể thay vì import sai dữ liệu.
          </div>
        </div>
      </Card>

      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
        <DropZone
          title="1. File Đơn hàng & Quyết toán"
          description={'VD: "Tất cả đơn hàng" (OrderSKUList) từ TikTok Seller Center, hoặc file Đối soát/Statement of Account nếu có.'}
          icon={FileSpreadsheet}
          onFile={handleSettlement}
        />
        <DropZone
          title="2. File Đơn trả hàng / Hoàn tiền"
          description={'VD: "Đơn trả hàng hoàn tiền" từ Seller Center. Chỉ các yêu cầu đã hoàn tất (Completed) mới được tính là mất doanh thu.'}
          icon={RefreshCw}
          onFile={handleReturns}
        />
        <DropZone
          title="3. File Báo cáo Quảng cáo"
          description="TikTok Ads / Shopee Ads. Mặc định Ads = 0đ cho tới khi có dữ liệu import."
          icon={Megaphone}
          onFile={handleAds}
        />
        <DropZone
          title="4. File Config SKU & Giá vốn mới"
          description={'File Config SKU chuẩn (Seller SKU/Loại/Giá vốn) hoặc file "Batch Edit Template" xuất từ Seller Center (chưa có giá vốn, sẽ cần nhập tay sau).'}
          icon={Package}
          onFile={handleConfig}
          accentColor="pink"
        />
      </div>

      <Card className="p-4">
        <div className="flex items-center justify-between mb-3">
          <h3 className="text-sm font-semibold text-blue-950 flex items-center gap-2"><ClipboardList size={16} /> Nhật ký Import</h3>
        </div>
        <div className="space-y-2 max-h-72 overflow-y-auto">
          {state.logs.map((log) => (
            <div key={log.id} className="flex items-start gap-2 text-xs border-b border-slate-50 pb-2 last:border-0">
              <span className={`w-1.5 h-1.5 rounded-full mt-1.5 shrink-0 ${
                log.type === "success" ? "bg-emerald-500" : log.type === "warning" ? "bg-amber-500" : "bg-blue-400"
              }`} />
              <div className="flex-1">
                <div className="text-slate-600">{log.message}</div>
                {log.details && log.details.length > 0 && (
                  <ul className="mt-1 text-[11px] text-amber-600 list-disc list-inside space-y-0.5">
                    {log.details.slice(0, 5).map((d, i) => <li key={i}>{d}</li>)}
                    {log.details.length > 5 && <li>...và {log.details.length - 5} cảnh báo khác</li>}
                  </ul>
                )}
                <div className="text-[10px] text-slate-300 mt-0.5">{log.time}</div>
              </div>
            </div>
          ))}
          {state.logs.length === 0 && <div className="text-xs text-slate-400 text-center py-4">Chưa có hoạt động import nào.</div>}
        </div>
      </Card>
    </div>
  );
}

/* ============================================================================
 * 16. CONFIG CENTER
 * ========================================================================== */

function CogsInlineInput({ platform, sku, value, isAdmin }) {
  const { dispatch } = useApp();
  const [local, setLocal] = useState(String(value));
  React.useEffect(() => setLocal(String(value)), [value]);

  if (!isAdmin) return <span className="font-medium">{formatVND(value)}</span>;

  return (
    <input
      type="number"
      value={local}
      onChange={(e) => setLocal(e.target.value)}
      onBlur={() => {
        const n = Number(local) || 0;
        if (n !== value) dispatch({ type: "UPDATE_SKU_COGS", payload: { platform, sku, cogs: n } });
      }}
      className="w-24 text-right border border-slate-200 rounded-md px-1.5 py-1 text-xs focus:border-blue-400 focus:outline-none"
    />
  );
}

const LOW_STOCK_THRESHOLD = 50;

/**
 * Bảng quản lý SKU (tab Sản phẩm) — theo đúng yêu cầu:
 * - Bỏ hoàn toàn cột "Loại", chỉ còn 4 cột: SKU | Tên sản phẩm | Tồn kho | Giá vốn.
 * - Cảnh báo đỏ khi Số lượng tồn kho khả dụng < 50 (đọc từ cột "Số lượng"
 *   trong chính file Sản phẩm/Batch Edit Template khi import).
 * - Công cụ nhập giá vốn đồng loạt: chọn nhiều SKU (checkbox) → nhập số tiền
 *   → bấm "Nhập nhanh" → áp dụng cho toàn bộ SKU đã chọn (dispatch
 *   BULK_UPDATE_SKU_COGS).
 */
function SkuManagementTable({ platform }) {
  const { state, dispatch } = useApp();
  const [search, setSearch] = useState("");
  const [selected, setSelected] = useState(new Set());
  const [bulkAmount, setBulkAmount] = useState("");
  const isAdmin = state.role === "admin";

  const list = state.skuConfig[platform].filter((e) =>
    search ? (e.sku + e.name).toLowerCase().includes(search.toLowerCase()) : true
  );
  const missingCogsCount = state.skuConfig[platform].filter((e) => !e.cogs).length;
  const lowStockCount = state.skuConfig[platform].filter((e) => e.stockQty < LOW_STOCK_THRESHOLD).length;

  const allVisibleSelected = list.length > 0 && list.every((e) => selected.has(e.sku));
  function toggleAll() {
    const next = new Set(selected);
    if (allVisibleSelected) list.forEach((e) => next.delete(e.sku));
    else list.forEach((e) => next.add(e.sku));
    setSelected(next);
  }
  function toggleOne(sku) {
    const next = new Set(selected);
    if (next.has(sku)) next.delete(sku); else next.add(sku);
    setSelected(next);
  }
  function applyBulk() {
    const cogs = Number(bulkAmount);
    if (!cogs || selected.size === 0) return;
    dispatch({ type: "BULK_UPDATE_SKU_COGS", payload: { platform, skus: Array.from(selected), cogs } });
    setSelected(new Set());
    setBulkAmount("");
  }

  return (
    <Card className="p-4">
      <div className="flex items-center justify-between mb-3 gap-2 flex-wrap">
        <h3 className="text-sm font-semibold text-blue-950 flex items-center gap-2">
          {platform === PLATFORM.TIKTOK ? <Store size={15} /> : <ShoppingBag size={15} />} SKU {platformLabel(platform)}
        </h3>
        <div className="flex items-center gap-2">
          <div className="relative">
            <Search size={13} className="absolute left-2.5 top-2.5 text-slate-400" />
            <input value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Tìm SKU..." className="pl-7 pr-2 py-1.5 border border-slate-200 rounded-lg text-xs w-36" />
          </div>
          <button
            onClick={() => downloadSkuConfigTemplate(platform)}
            className="flex items-center gap-1 text-xs font-medium border border-slate-200 rounded-lg px-2.5 py-1.5 hover:bg-slate-50"
          >
            <Download size={13} /> File mẫu
          </button>
        </div>
      </div>

      {(missingCogsCount > 0 || lowStockCount > 0) && (
        <div className="flex flex-col gap-1.5 mb-2">
          {missingCogsCount > 0 && (
            <div className="flex items-start gap-1.5 text-[11px] text-amber-700 bg-amber-50 rounded-lg px-2.5 py-2">
              <AlertCircle size={12} className="shrink-0 mt-0.5" />
              {missingCogsCount} SKU đang có giá vốn = 0đ — chọn SKU bên dưới và dùng công cụ "Nhập nhanh", hoặc bấm trực tiếp vào ô giá vốn.
            </div>
          )}
          {lowStockCount > 0 && (
            <div className="flex items-start gap-1.5 text-[11px] text-rose-700 bg-rose-50 rounded-lg px-2.5 py-2">
              <AlertTriangle size={12} className="shrink-0 mt-0.5" />
              {lowStockCount} SKU đang tồn kho dưới {LOW_STOCK_THRESHOLD} — xem các dòng highlight đỏ bên dưới để nhập hàng kịp thời.
            </div>
          )}
        </div>
      )}

      {isAdmin && (
        <div className="flex flex-wrap items-center gap-2 mb-3 bg-slate-50 border border-slate-100 rounded-xl p-2.5">
          <span className="text-[11px] text-slate-400">{selected.size} SKU đã chọn</span>
          <input
            type="number"
            value={bulkAmount}
            onChange={(e) => setBulkAmount(e.target.value)}
            placeholder="Nhập giá vốn (đ)"
            className="border border-slate-200 rounded-lg px-2.5 py-1.5 text-xs w-36"
          />
          <button
            onClick={applyBulk}
            disabled={selected.size === 0 || !bulkAmount}
            className="flex items-center gap-1 bg-blue-950 hover:bg-blue-900 disabled:bg-slate-200 disabled:text-slate-400 text-white text-xs font-medium px-3 py-1.5 rounded-lg transition"
          >
            <CheckCircle2 size={13} /> Nhập nhanh
          </button>
        </div>
      )}

      <div className="overflow-x-auto max-h-96 overflow-y-auto">
        <table className="w-full text-xs min-w-[460px]">
          <thead className="sticky top-0 bg-white">
            <tr className="text-left text-slate-400 border-b border-slate-100">
              {isAdmin && (
                <th className="py-2 pr-2 w-6">
                  <input type="checkbox" checked={allVisibleSelected} onChange={toggleAll} />
                </th>
              )}
              <th className="py-2 pr-2">Mã SKU Người Bán</th>
              <th className="py-2 pr-2">Tên Sản Phẩm</th>
              <th className="py-2 pr-2 text-right">Số Lượng Tồn Kho</th>
              <th className="py-2 pr-2 text-right">Giá Vốn</th>
            </tr>
          </thead>
          <tbody>
            {list.map((e, i) => {
              const lowStock = e.stockQty < LOW_STOCK_THRESHOLD;
              return (
                <tr key={e.sku + i} className={`border-b border-slate-50 last:border-0 ${lowStock ? "bg-rose-50/60" : ""}`}>
                  {isAdmin && (
                    <td className="py-2 pr-2"><input type="checkbox" checked={selected.has(e.sku)} onChange={() => toggleOne(e.sku)} /></td>
                  )}
                  <td className="py-2 pr-2 font-mono text-blue-950">{e.sku}</td>
                  <td className="py-2 pr-2 text-slate-500 max-w-[220px] truncate" title={e.name}>{e.name}</td>
                  <td className={`py-2 pr-2 text-right font-medium ${lowStock ? "text-rose-600" : "text-slate-600"}`}>
                    <span className="inline-flex items-center gap-1 justify-end">
                      {lowStock && <AlertTriangle size={12} />} {formatNumber(e.stockQty)}
                    </span>
                  </td>
                  <td className="py-2 pr-2 text-right font-medium">
                    <CogsInlineInput platform={platform} sku={e.sku} value={e.cogs} isAdmin={isAdmin} />
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
        {list.length === 0 && <div className="text-center text-slate-400 py-6">Không tìm thấy SKU phù hợp.</div>}
      </div>
    </Card>
  );
}

/**
 * Tab "Phân tích Sản phẩm" — gộp doanh thu/lợi nhuận/tồn kho theo từng SKU
 * trên CẢ 2 sàn, dựa trên breakdown items[] của các đơn hàng đã tính toán
 * trong khoảng thời gian đang lọc. Cảnh báo đỏ SKU tồn kho < 50.
 * Lưu ý: Lợi nhuận SKU ở đây là biên gộp (Doanh thu − Giá vốn), CHƯA trừ
 * phí sàn/phí đóng gói vì các khoản phí đó phát sinh ở cấp ĐƠN HÀNG chứ
 * không tách được theo từng SKU trong 1 đơn nhiều sản phẩm.
 */
function ProductAnalysisTab() {
  const { state, computedOrders } = useApp();
  const { start, end } = state.dateRange;
  const [sortBy, setSortBy] = useState("revenue");

  const skuStats = useMemo(() => {
    const map = new Map();
    for (const platform of [PLATFORM.TIKTOK, PLATFORM.SHOPEE]) {
      const cfgMap = skuConfigToMap(state.skuConfig[platform]);
      const orders = computedOrders[platform].filter((o) => isWithinRange(o.date, start, end) && o.status !== STATUS.CANCELLED);
      for (const o of orders) {
        for (const it of o.items) {
          const key = platform + "::" + it.sellerSku;
          if (!map.has(key)) {
            const cfg = cfgMap.get(it.sellerSku.toLowerCase());
            map.set(key, {
              sku: it.sellerSku,
              platform,
              name: cfg ? cfg.name : it.sellerSku,
              stockQty: cfg ? cfg.stockQty : 0,
              orderCount: 0,
              revenue: 0,
              cogs: 0,
            });
          }
          const s = map.get(key);
          s.orderCount += 1;
          s.revenue += it.subtotal;
          s.cogs += it.cogs;
        }
      }
    }
    return Array.from(map.values()).map((s) => ({ ...s, profit: s.revenue - s.cogs }));
  }, [computedOrders, state.skuConfig, start, end]);

  const sorted = [...skuStats].sort((a, b) => b[sortBy] - a[sortBy]);
  const lowStock = [PLATFORM.TIKTOK, PLATFORM.SHOPEE]
    .flatMap((p) => state.skuConfig[p].map((e) => ({ ...e, platform: p })))
    .filter((e) => e.stockQty < LOW_STOCK_THRESHOLD)
    .sort((a, b) => a.stockQty - b.stockQty);

  return (
    <div className="space-y-4">
      {lowStock.length > 0 && (
        <Card className="p-4 bg-rose-50/60 border-rose-100">
          <h3 className="text-sm font-semibold text-rose-700 mb-2 flex items-center gap-2"><AlertTriangle size={15} /> {lowStock.length} SKU sắp hết hàng (tồn kho &lt; {LOW_STOCK_THRESHOLD})</h3>
          <div className="flex flex-wrap gap-1.5">
            {lowStock.slice(0, 30).map((e) => (
              <span key={e.platform + e.sku} className="text-[11px] font-mono bg-white border border-rose-200 text-rose-600 rounded-lg px-2 py-1">
                {e.sku} · {formatNumber(e.stockQty)} còn lại
              </span>
            ))}
          </div>
        </Card>
      )}
      <Card className="p-4">
        <div className="flex items-center justify-between mb-3 flex-wrap gap-2">
          <h3 className="text-sm font-semibold text-blue-950">Doanh thu & lợi nhuận theo SKU</h3>
          <select value={sortBy} onChange={(e) => setSortBy(e.target.value)} className="border border-slate-200 rounded-lg text-xs px-2 py-1.5">
            <option value="revenue">Sắp xếp theo Doanh thu</option>
            <option value="profit">Sắp xếp theo Lợi nhuận</option>
            <option value="orderCount">Sắp xếp theo Số đơn</option>
          </select>
        </div>
        <div className="overflow-x-auto max-h-[480px] overflow-y-auto">
          <table className="w-full text-xs min-w-[620px]">
            <thead className="sticky top-0 bg-white">
              <tr className="text-left text-slate-400 border-b border-slate-100">
                <th className="py-2 pr-2">SKU</th>
                <th className="py-2 pr-2">Tên sản phẩm</th>
                <th className="py-2 pr-2 text-right">Tồn kho</th>
                <th className="py-2 pr-2 text-right">Số đơn</th>
                <th className="py-2 pr-2 text-right">Doanh thu</th>
                <th className="py-2 pr-2 text-right">Lợi nhuận gộp</th>
              </tr>
            </thead>
            <tbody>
              {sorted.slice(0, 200).map((s) => {
                const lowStockRow = s.stockQty < LOW_STOCK_THRESHOLD;
                return (
                  <tr key={s.platform + s.sku} className={`border-b border-slate-50 last:border-0 ${lowStockRow ? "bg-rose-50/60" : ""}`}>
                    <td className="py-2 pr-2 font-mono text-blue-950">{s.sku}</td>
                    <td className="py-2 pr-2 text-slate-500 max-w-[220px] truncate" title={s.name}>{s.name}</td>
                    <td className={`py-2 pr-2 text-right font-medium ${lowStockRow ? "text-rose-600" : "text-slate-600"}`}>{formatNumber(s.stockQty)}</td>
                    <td className="py-2 pr-2 text-right">{formatNumber(s.orderCount)}</td>
                    <td className="py-2 pr-2 text-right">{formatVND(s.revenue)}</td>
                    <td className={`py-2 pr-2 text-right font-medium ${s.profit >= 0 ? "text-emerald-600" : "text-rose-600"}`}>{formatVND(s.profit)}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
          {sorted.length === 0 && <div className="text-center text-slate-400 py-8">Không có dữ liệu SKU trong khoảng thời gian đã chọn.</div>}
        </div>
      </Card>
    </div>
  );
}

/**
 * Bảng cấu hình PHÍ SÀN — KHÔNG dùng để tự động tính lại phí từng đơn (phí
 * thật vẫn luôn đọc từ file Quyết toán/Đối soát khi import). Đây là nơi LƯU
 * TRỮ thông tin các hạng mục phí sàn (tên, mức phí %/số tiền cố định, bắt
 * buộc hay tùy chọn, ghi chú) để:
 *  1) Tham khảo nhanh ngay trong Tổng quan.
 *  2) Làm dữ liệu nguồn cho tính năng "Gợi ý giá bán" sẽ xây dựng sau này
 *     (nhập giá vốn + loại mặt hàng → tự cộng các phí trong bảng này để ra
 *     giá bán đề xuất & biên lợi nhuận). Vì vậy bảng được thiết kế dạng danh
 *     sách THÊM/SỬA/XÓA được, không phải vài ô input cố định.
 */
function FeeConfigTable() {
  const { state, dispatch } = useApp();
  const isAdmin = state.role === "admin";
  const [draft, setDraft] = useState(null); // fee item đang sửa (bản nháp)

  function startEdit(item) { setDraft({ ...item }); }
  function cancelEdit() { setDraft(null); }
  function saveEdit() {
    dispatch({ type: "UPDATE_FEE_CONFIG_ITEM", payload: draft });
    setDraft(null);
  }
  function addNew() {
    dispatch({
      type: "ADD_FEE_CONFIG_ITEM",
      payload: { id: uid("fee"), name: "Phí mới", valueType: "percent", value: 0, required: false, note: "" },
    });
  }

  return (
    <Card className="p-4">
      <div className="flex items-center justify-between mb-1">
        <h3 className="text-sm font-semibold text-blue-950 flex items-center gap-2"><Percent size={15} /> Bảng cấu hình phí sàn</h3>
        {isAdmin && (
          <button onClick={addNew} className="flex items-center gap-1 text-xs font-medium border border-slate-200 rounded-lg px-2.5 py-1.5 hover:bg-slate-50">
            <Plus size={13} /> Thêm hạng mục phí
          </button>
        )}
      </div>
      <p className="text-[11px] text-slate-400 mb-3">
        Bảng lưu thông tin phí sàn (hoa hồng, voucher, phí xử lý đơn...) để tham khảo và làm dữ liệu nguồn cho tính năng gợi ý giá bán sau này. Số phí thật dùng để tính lợi nhuận vẫn luôn lấy từ file Quyết toán/Đối soát bạn import.
      </p>
      <div className="overflow-x-auto max-h-96 overflow-y-auto">
        <table className="w-full text-xs min-w-[560px]">
          <thead className="sticky top-0 bg-white">
            <tr className="text-left text-slate-400 border-b border-slate-100">
              <th className="py-2 pr-2">Loại phí</th>
              <th className="py-2 pr-2">Mức phí</th>
              <th className="py-2 pr-2">Bắt buộc</th>
              <th className="py-2 pr-2">Ghi chú</th>
              {isAdmin && <th className="py-2 pr-2"></th>}
            </tr>
          </thead>
          <tbody>
            {state.feeConfig.map((item) => {
              const editing = draft && draft.id === item.id;
              return (
                <tr key={item.id} className="border-b border-slate-50 last:border-0 align-top">
                  <td className="py-2 pr-2 font-medium text-blue-950">
                    {editing ? (
                      <input value={draft.name} onChange={(e) => setDraft({ ...draft, name: e.target.value })} className="border border-slate-200 rounded px-1.5 py-1 w-40" />
                    ) : item.name}
                  </td>
                  <td className="py-2 pr-2 whitespace-nowrap">
                    {editing ? (
                      <div className="flex items-center gap-1">
                        <input type="number" value={draft.value} onChange={(e) => setDraft({ ...draft, value: Number(e.target.value) })} className="border border-slate-200 rounded px-1.5 py-1 w-20" />
                        <select value={draft.valueType} onChange={(e) => setDraft({ ...draft, valueType: e.target.value })} className="border border-slate-200 rounded px-1 py-1">
                          <option value="percent">%</option>
                          <option value="fixed">đ/đơn</option>
                        </select>
                      </div>
                    ) : (
                      item.valueType === "percent" ? `${item.value}%` : `${formatVND(item.value)}/đơn`
                    )}
                  </td>
                  <td className="py-2 pr-2">
                    {editing ? (
                      <input type="checkbox" checked={draft.required} onChange={(e) => setDraft({ ...draft, required: e.target.checked })} />
                    ) : (
                      <Badge tone={item.required ? "navy" : "gray"}>{item.required ? "Bắt buộc" : "Tùy chọn"}</Badge>
                    )}
                  </td>
                  <td className="py-2 pr-2 text-slate-500 max-w-[220px]">
                    {editing ? (
                      <input value={draft.note} onChange={(e) => setDraft({ ...draft, note: e.target.value })} className="border border-slate-200 rounded px-1.5 py-1 w-full" />
                    ) : item.note}
                  </td>
                  {isAdmin && (
                    <td className="py-2 pr-2">
                      {editing ? (
                        <div className="flex items-center gap-1">
                          <button onClick={saveEdit} className="text-emerald-600 hover:underline">Lưu</button>
                          <button onClick={cancelEdit} className="text-slate-400 hover:underline">Hủy</button>
                        </div>
                      ) : (
                        <div className="flex items-center gap-2">
                          <button onClick={() => startEdit(item)} className="text-blue-700 hover:underline">Sửa</button>
                          <button onClick={() => dispatch({ type: "DELETE_FEE_CONFIG_ITEM", payload: item.id })} className="text-rose-500 hover:underline">Xóa</button>
                        </div>
                      )}
                    </td>
                  )}
                </tr>
              );
            })}
          </tbody>
        </table>
        {state.feeConfig.length === 0 && <div className="text-center text-slate-400 py-6">Chưa có hạng mục phí nào.</div>}
      </div>
    </Card>
  );
}

function SettingsForm() {
  const { state, dispatch } = useApp();
  const isAdmin = state.role === "admin";
  const [form, setForm] = useState(state.settings);

  function save() {
    dispatch({ type: "UPDATE_SETTINGS", payload: form });
  }

  return (
    <Card className="p-4">
      <h3 className="text-sm font-semibold text-blue-950 mb-4 flex items-center gap-2"><Settings size={15} /> Thông số tài chính cố định</h3>
      <div className="space-y-4">
        <div>
          <label className="text-xs font-medium text-slate-500 flex items-center gap-1">
            Phí đóng gói cố định / đơn <Tooltip2 text="Áp dụng cho cả đơn thành công và đơn hoàn (ghi nhận là chi phí tổn thất)." />
          </label>
          <div className="flex items-center gap-2 mt-1.5">
            <input
              type="number"
              disabled={!isAdmin}
              value={form.packagingFee}
              onChange={(e) => setForm({ ...form, packagingFee: Number(e.target.value) })}
              className="border border-slate-200 rounded-lg px-3 py-2 text-sm w-40 disabled:bg-slate-50 disabled:text-slate-400"
            />
            <span className="text-xs text-slate-400">đ / đơn (mặc định 6.000đ)</span>
          </div>
        </div>
        <div>
          <label className="text-xs font-medium text-slate-500">Chi phí cố định hàng tháng</label>
          <div className="flex items-center gap-2 mt-1.5">
            <input
              type="number"
              disabled={!isAdmin}
              value={form.monthlyFixedCost}
              onChange={(e) => setForm({ ...form, monthlyFixedCost: Number(e.target.value) })}
              className="border border-slate-200 rounded-lg px-3 py-2 text-sm w-40 disabled:bg-slate-50 disabled:text-slate-400"
            />
            <span className="text-xs text-slate-400">đ / tháng (mặc định 0đ)</span>
          </div>
        </div>
        <div>
          <label className="text-xs font-medium text-slate-500">Biến phí khác</label>
          <div className="flex items-center gap-2 mt-1.5">
            <input
              type="number"
              disabled={!isAdmin}
              value={form.otherVariableCost}
              onChange={(e) => setForm({ ...form, otherVariableCost: Number(e.target.value) })}
              className="border border-slate-200 rounded-lg px-3 py-2 text-sm w-40 disabled:bg-slate-50 disabled:text-slate-400"
            />
            <span className="text-xs text-slate-400">đ (mặc định 0đ)</span>
          </div>
        </div>
        {isAdmin ? (
          <button onClick={save} className="bg-blue-950 hover:bg-blue-900 text-white text-sm font-medium px-4 py-2 rounded-xl transition">
            Lưu thay đổi
          </button>
        ) : (
          <div className="flex items-center gap-1.5 text-xs text-slate-400"><Lock size={12} /> Chế độ Staff chỉ xem, không thể chỉnh sửa cấu hình.</div>
        )}
      </div>
    </Card>
  );
}

function RolePanel() {
  const { state, dispatch } = useApp();
  return (
    <Card className="p-4">
      <h3 className="text-sm font-semibold text-blue-950 mb-3 flex items-center gap-2"><Users size={15} /> Phân quyền hệ thống</h3>
      <div className="flex items-center gap-2">
        {[
          { key: "admin", label: "Admin", icon: ShieldCheck, desc: "Toàn quyền: import, chỉnh Config." },
          { key: "staff", label: "Staff", icon: Lock, desc: "Chỉ xem & import, không sửa Config." },
        ].map((r) => (
          <button
            key={r.key}
            onClick={() => dispatch({ type: "SET_ROLE", payload: r.key })}
            className={`flex-1 border rounded-xl p-3 text-left transition ${
              state.role === r.key ? "border-blue-950 bg-blue-50/50" : "border-slate-200 hover:border-slate-300"
            }`}
          >
            <div className="flex items-center gap-1.5 font-medium text-sm text-blue-950"><r.icon size={14} /> {r.label}</div>
            <div className="text-xs text-slate-400 mt-1">{r.desc}</div>
          </button>
        ))}
      </div>
      <div className="text-[11px] text-slate-300 mt-3">* Khung phân quyền demo phía client. Khi kết nối backend thật, cần xác thực (auth) phía server để đảm bảo an toàn dữ liệu.</div>
    </Card>
  );
}

function FeeConfigCenterTab() {
  const { state, dispatch, computedOrders } = useApp();
  const allComputed = { [PLATFORM.TIKTOK]: computedOrders[PLATFORM.TIKTOK], [PLATFORM.SHOPEE]: computedOrders[PLATFORM.SHOPEE] };
  const totalAds = state.ads[PLATFORM.TIKTOK] + state.ads[PLATFORM.SHOPEE];
  const allOrders = [...allComputed[PLATFORM.TIKTOK], ...allComputed[PLATFORM.SHOPEE]];
  const aggAll = aggregateOrders(allOrders, totalAds, state.settings.monthlyFixedCost, state.settings.otherVariableCost);
  const aggByPlatform = {
    [PLATFORM.TIKTOK]: aggregateOrders(allComputed[PLATFORM.TIKTOK], state.ads[PLATFORM.TIKTOK]),
    [PLATFORM.SHOPEE]: aggregateOrders(allComputed[PLATFORM.SHOPEE], state.ads[PLATFORM.SHOPEE]),
  };

  return (
    <div className="space-y-5">
      <RolePanel />
      <FeeConfigTable />
      <SettingsForm />
      <Card className="p-4">
        <h3 className="text-sm font-semibold text-blue-950 mb-3 flex items-center gap-2"><Download size={15} /> Xuất báo cáo</h3>
        <div className="flex flex-wrap gap-2">
          <button
            onClick={() => exportReportToExcel(state, allComputed, aggAll, aggByPlatform)}
            className="flex items-center gap-1.5 bg-blue-950 hover:bg-blue-900 text-white text-sm font-medium px-4 py-2 rounded-xl transition"
          >
            <FileSpreadsheet size={15} /> Xuất báo cáo Excel tổng hợp
          </button>
          <button
            onClick={() => dispatch({ type: "RESTORE_INITIAL_IMPORT" })}
            className="flex items-center gap-1.5 border border-slate-200 text-sm font-medium px-4 py-2 rounded-xl hover:bg-slate-50 transition"
          >
            <RefreshCw size={15} /> Khôi phục dữ liệu TikTok ban đầu
          </button>
        </div>
      </Card>
      <DataManagementPanel />
    </div>
  );
}

/**
 * Bảng "Quản lý & Xóa Dữ Liệu" — thay vì chỉ có 1 nút xóa tất cả, tách riêng
 * từng nút theo đúng loại dữ liệu (SKU/Đơn hàng+Đối soát/Ads theo từng sàn)
 * để người dùng xóa đúng phần cần làm sạch mà KHÔNG ảnh hưởng dữ liệu khác.
 * Mỗi nút đều có Confirm Alert trước khi thực thi (không thể hoàn tác).
 */
function DataManagementPanel() {
  const { state, dispatch } = useApp();
  const isAdmin = state.role === "admin";

  const actions = [
    {
      key: "sku-tiktok",
      label: "Xóa SKU TikTok Shop",
      desc: `${state.skuConfig[PLATFORM.TIKTOK].length} SKU hiện có — chỉ xóa danh sách SKU & tồn kho của TikTok.`,
      confirmMsg: "Xóa toàn bộ danh sách SKU & tồn kho của TikTok Shop? Đơn hàng và dữ liệu Shopee sẽ KHÔNG bị ảnh hưởng.",
      run: () => dispatch({ type: "CLEAR_SKU_CONFIG", payload: { platform: PLATFORM.TIKTOK } }),
    },
    {
      key: "sku-shopee",
      label: "Xóa SKU Shopee",
      desc: `${state.skuConfig[PLATFORM.SHOPEE].length} SKU hiện có — chỉ xóa danh sách SKU & tồn kho của Shopee.`,
      confirmMsg: "Xóa toàn bộ danh sách SKU & tồn kho của Shopee? Đơn hàng và dữ liệu TikTok sẽ KHÔNG bị ảnh hưởng.",
      run: () => dispatch({ type: "CLEAR_SKU_CONFIG", payload: { platform: PLATFORM.SHOPEE } }),
    },
    {
      key: "orders-tiktok",
      label: "Xóa Đơn Hàng & Đối Soát TikTok",
      desc: `${state.orders[PLATFORM.TIKTOK].length} đơn hiện có — chỉ xóa đơn hàng + dữ liệu Quyết toán/phí sàn của TikTok.`,
      confirmMsg: "Xóa toàn bộ đơn hàng & dữ liệu Đối soát của TikTok Shop? SKU và dữ liệu Shopee sẽ KHÔNG bị ảnh hưởng.",
      run: () => dispatch({ type: "CLEAR_ORDERS", payload: { platform: PLATFORM.TIKTOK } }),
    },
    {
      key: "orders-shopee",
      label: "Xóa Đơn Hàng & Đối Soát Shopee",
      desc: `${state.orders[PLATFORM.SHOPEE].length} đơn hiện có — chỉ xóa đơn hàng + dữ liệu Quyết toán/phí sàn của Shopee.`,
      confirmMsg: "Xóa toàn bộ đơn hàng & dữ liệu Đối soát của Shopee? SKU và dữ liệu TikTok sẽ KHÔNG bị ảnh hưởng.",
      run: () => dispatch({ type: "CLEAR_ORDERS", payload: { platform: PLATFORM.SHOPEE } }),
    },
    {
      key: "ads",
      label: "Xóa Dữ Liệu Quảng Cáo (Ads)",
      desc: `Tổng hiện có: ${formatVND(state.ads[PLATFORM.TIKTOK] + state.ads[PLATFORM.SHOPEE])} — chỉ xóa dữ liệu chi phí Ads, không ảnh hưởng đơn hàng/SKU.`,
      confirmMsg: "Xóa toàn bộ dữ liệu chi phí quảng cáo (Ads) của cả 2 sàn?",
      run: () => dispatch({ type: "CLEAR_ADS" }),
    },
  ];

  return (
    <Card className="p-4">
      <h3 className="text-sm font-semibold text-blue-950 mb-1 flex items-center gap-2"><Trash2 size={15} /> Quản lý & Xóa Dữ Liệu</h3>
      <p className="text-[11px] text-slate-400 mb-3">Xóa đúng phần dữ liệu cần làm sạch trước khi import lại — không ảnh hưởng các phần dữ liệu còn lại. Mỗi thao tác đều cần xác nhận và không thể hoàn tác.</p>
      {!isAdmin ? (
        <div className="text-xs text-slate-400 flex items-center gap-1.5"><Lock size={13} /> Chỉ tài khoản Admin mới có quyền xóa dữ liệu.</div>
      ) : (
        <div className="grid sm:grid-cols-2 gap-2.5">
          {actions.map((a) => (
            <div key={a.key} className="flex items-center justify-between gap-3 border border-slate-100 rounded-xl px-3 py-2.5">
              <div className="min-w-0">
                <div className="text-xs font-medium text-blue-950">{a.label}</div>
                <div className="text-[11px] text-slate-400 truncate">{a.desc}</div>
              </div>
              <button
                onClick={() => { if (confirm(a.confirmMsg)) a.run(); }}
                className="shrink-0 flex items-center gap-1 text-xs font-medium border border-rose-200 text-rose-600 rounded-lg px-2.5 py-1.5 hover:bg-rose-50 transition"
              >
                <Trash2 size={12} /> Xóa
              </button>
            </div>
          ))}
          <div className="flex items-center justify-between gap-3 border border-rose-200 bg-rose-50/60 rounded-xl px-3 py-2.5 sm:col-span-2">
            <div className="min-w-0">
              <div className="text-xs font-semibold text-rose-700">Reset Toàn Bộ Hệ Thống</div>
              <div className="text-[11px] text-rose-500/80">Xóa sạch TOÀN BỘ dữ liệu (SKU, đơn hàng, Ads của cả 2 sàn) về trạng thái mặc định trống.</div>
            </div>
            <button
              onClick={() => { if (confirm("Xóa TOÀN BỘ dữ liệu hệ thống (SKU, đơn hàng, Ads của cả TikTok lẫn Shopee) và không thể hoàn tác?")) dispatch({ type: "RESET_ALL" }); }}
              className="shrink-0 flex items-center gap-1.5 bg-rose-600 hover:bg-rose-700 text-white text-xs font-semibold rounded-lg px-3 py-1.5 transition"
            >
              <Trash2 size={13} /> Reset toàn bộ
            </button>
          </div>
        </div>
      )}
    </Card>
  );
}

/* ============================================================================
 * 17. ROOT APP
 * ========================================================================== */

function AppShell() {
  const [activeSection, setActiveSection] = useState("overview");
  const [collapsed, setCollapsed] = useState(false);
  const goToImport = () => setActiveSection("import");

  return (
    <div className="min-h-screen bg-slate-50 flex">
      <Sidebar activeSection={activeSection} setActiveSection={setActiveSection} collapsed={collapsed} setCollapsed={setCollapsed} />
      <div className="flex-1 min-w-0 flex flex-col">
        <main className="flex-1 max-w-7xl w-full mx-auto px-4 sm:px-6 py-5">
          {activeSection === "overview" && <OverviewSection onGoToImport={goToImport} />}
          {activeSection === "product" && <ProductSection onGoToImport={goToImport} />}
          {activeSection === "marketing" && <MarketingSection onGoToImport={goToImport} />}
          {activeSection === "import" && <ImportCenterTab />}
        </main>
        <footer className="text-center text-[11px] text-slate-300 py-4">
          ELLA Accents · Profit Control Center · Dữ liệu xử lý cục bộ trên trình duyệt, không lưu trữ máy chủ ngoài.
        </footer>
      </div>
    </div>
  );
}

export default function App() {
  return (
    <AppProvider>
      <AppShell />
    </AppProvider>
  );
}
