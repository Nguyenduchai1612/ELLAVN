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
  ChevronLeft, ChevronRight, Plus, AlertTriangle, Target, Wallet, Calculator,
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
  // Cột "product_id"/"ID sản phẩm" trong file Batch Edit Template — dùng làm
  // lớp lọc dòng rác/hướng dẫn thứ 2 (độc lập với cột seller_sku/sku_id).
  productId: ["product id", "id sản phẩm", "id san pham"],
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
  // Sheet dữ liệu SKU thật trong file "Batch Edit Template" (Tiktoksellercenter_batchedit_...)
  // của TikTok Seller Center — luôn tên "Template", tách biệt các sheet phụ
  // "Instruction/HiddenStyle/HiddenAttr/SpecialProductListingType/TemplateConfig/Brand".
  "template",
];
const NON_DETAIL_SHEET_NAME_HINTS = [
  "báo cáo", "bao cao", "report", "lịch sử rút tiền", "lich su rut tien", "giải thích", "giai thich",
  "instruction", "hướng dẫn", "huong dan", "hiddenstyle", "hiddenattr", "speciallistingtype",
  "templateconfig", "brand",
];

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
  "sku người bán", "sku nguoi ban", "seller_sku",
  "mã nhận dạng sản phẩm", "ma nhan dang san pham",
  "lượng hàng có sẵn", "luong hang co san",
  "nhập giá của sản phẩm", "nhap gia cua san pham",
];
// Các cụm từ chắc chắn là RÁC khi xuất hiện trong cột product_id/ID sản phẩm.
const PRODUCT_ID_JUNK_MARKERS = [
  "product_id", "id sản phẩm", "id san pham", "v4",
  "bắt buộc", "bat buoc", "không thể chỉnh sửa", "khong the chinh sua",
];
function isJunkSkuRow(sellerSkuRaw) {
  const s = String(sellerSkuRaw ?? "").trim();
  if (!s) return true; // seller_sku rỗng -> chắc chắn không phải dòng dữ liệu
  const norm = normalizeHeader(s);
  return SKU_JUNK_MARKERS.some((m) => norm.includes(m));
}
function isJunkProductIdRow(productIdRaw) {
  const s = String(productIdRaw ?? "").trim();
  if (!s) return false; // rỗng không tự nó là dấu hiệu rác cho cột này
  const norm = normalizeHeader(s);
  return PRODUCT_ID_JUNK_MARKERS.some((m) => norm.includes(m));
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
    // Lớp lọc 2: kiểm tra thêm cột product_id (ID sản phẩm) nếu có — dòng dữ
    // liệu rác thường chứa "product_id", "ID sản phẩm", "V4", "Bắt buộc",
    // "Không thể chỉnh sửa" ở CẢ 2 cột cùng lúc, nên kiểm tra thêm cột này
    // giúp bắt chắc chắn hơn, không phụ thuộc hoàn toàn vào 1 cột duy nhất.
    if (map.productId !== undefined && isJunkProductIdRow(row[map.productId])) { skippedJunkRows++; return; }
    // Lớp lọc 3 (dự phòng, khi có cột sku_id): dữ liệu SKU thật luôn có
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
 * 5. INITIAL STATE — TẤT CẢ rỗng khi mới mở app. KHÔNG có dữ liệu mẫu/mock
 * fix cứng. Người dùng chủ động import file thật (Đơn hàng/Quyết toán, Trả
 * hàng, Ads, Config SKU/Batch Edit Template) ở tab Import để có số liệu.
 * ========================================================================== */

// KHÔNG fix cứng dữ liệu mẫu. Toàn bộ dữ liệu (đơn hàng, SKU, phí sàn theo
// đơn, Ads) chỉ được nạp vào state khi người dùng thực sự import file ở tab
// Import — trạng thái khởi tạo (initialState bên dưới) luôn là mảng rỗng, và
// khi có dữ liệu trong localStorage (từ phiên trước) thì ưu tiên khôi phục
// từ đó (xem loadPersistedState()).

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
    [PLATFORM.TIKTOK]: [],
    [PLATFORM.SHOPEE]: [],
  },
  orders: {
    [PLATFORM.TIKTOK]: [],
    [PLATFORM.SHOPEE]: [],
  },
  ads: { [PLATFORM.TIKTOK]: 0, [PLATFORM.SHOPEE]: 0 },
  logs: [],
  dateRange: { start: "", end: "" },
};

// ---- Lưu/khôi phục state qua localStorage (giữ dữ liệu qua các lần F5) ----
const STORAGE_KEY = "ella_accents_state_v1";

/** Đọc state đã lưu ở phiên trước từ localStorage; nếu chưa có / lỗi / dữ
 * liệu hỏng thì trả về initialState (rỗng hoàn toàn) — KHÔNG BAO GIỜ rơi về
 * dữ liệu mẫu giả lập. */
function loadPersistedState() {
  if (typeof window === "undefined" || !window.localStorage) return initialState;
  try {
    const raw = window.localStorage.getItem(STORAGE_KEY);
    if (!raw) return initialState;
    const parsed = JSON.parse(raw);
    // Merge nông với initialState để tránh crash nếu bản lưu cũ thiếu field
    // (vd: người dùng nâng cấp app lên bản có thêm state.feeConfig).
    return {
      ...initialState,
      ...parsed,
      settings: { ...initialState.settings, ...(parsed.settings || {}) },
      skuConfig: { ...initialState.skuConfig, ...(parsed.skuConfig || {}) },
      orders: { ...initialState.orders, ...(parsed.orders || {}) },
      ads: { ...initialState.ads, ...(parsed.ads || {}) },
      dateRange: { ...initialState.dateRange, ...(parsed.dateRange || {}) },
      feeConfig: Array.isArray(parsed.feeConfig) ? parsed.feeConfig : initialState.feeConfig,
      logs: Array.isArray(parsed.logs) ? parsed.logs : [],
    };
  } catch {
    return initialState;
  }
}

function persistState(state) {
  if (typeof window === "undefined" || !window.localStorage) return;
  try {
    window.localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  } catch {
    // localStorage đầy hoặc bị chặn (chế độ ẩn danh...) — bỏ qua, không crash app.
  }
}

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
        logs: [{ id: uid("log"), time: new Date().toLocaleString("vi-VN"), type: "info", message: "Đã xóa toàn bộ dữ liệu hệ thống. Sẵn sàng import dữ liệu mới." }],
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
  const [state, dispatch] = useReducer(reducer, undefined, loadPersistedState);

  // Tự động lưu toàn bộ state vào localStorage mỗi khi thay đổi — giữ dữ liệu
  // qua các lần F5/đóng-mở lại trình duyệt, không cần máy chủ ngoài.
  useEffect(() => {
    persistState(state);
  }, [state]);

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
  { key: "calculator", label: "Tính Lợi Nhuận", icon: Calculator },
  { key: "import", label: "Import", icon: FileSpreadsheet },
];

/** Logo góc trên Sidebar — ưu tiên hiển thị /favicon.png; nếu ảnh chưa tồn
 * tại/tải lỗi thì fallback về chữ "EA" (không để trống hoặc vỡ layout). */
function LogoMark() {
  const [imgFailed, setImgFailed] = useState(false);
  return (
    <div className="w-9 h-9 rounded-xl bg-blue-950 flex items-center justify-center text-white font-bold text-sm shrink-0 overflow-hidden">
      {!imgFailed ? (
        <img
          src="/favicon.png"
          alt="ELLA Accents Logo"
          className="w-8 h-8 rounded-lg object-cover"
          onError={() => setImgFailed(true)}
        />
      ) : (
        <span className="pointer-events-none">EA</span>
      )}
    </div>
  );
}

function Sidebar({ activeSection, setActiveSection, collapsed, setCollapsed }) {
  return (
    <aside
      className={`bg-white border-r border-slate-200 h-screen sticky top-0 flex flex-col shrink-0 transition-all duration-200 ${
        collapsed ? "w-[68px]" : "w-60"
      }`}
    >
      <div className="flex items-center gap-2 px-3.5 py-4 border-b border-slate-100">
        <LogoMark />
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

/* ============================================================================
 * 15c. PROFIT CALCULATOR — "Tính Lợi Nhuận Dự Phóng"
 * ========================================================================== */

const CALC_PLATFORMS = [
  { key: PLATFORM.TIKTOK, label: "TikTok Shop" },
  { key: PLATFORM.SHOPEE, label: "Shopee" },
];

/**
 * Công cụ ước tính lợi nhuận CHO 1 SẢN PHẨM trước khi lên đơn thật — KHÔNG
 * đụng vào dữ liệu đơn hàng/SKU đã import, chỉ là máy tính độc lập. Tỷ lệ
 * phí sàn % mặc định được gợi ý từ chính Bảng cấu hình phí sàn (Tổng quan >
 * Config) — đúng như mục đích ban đầu khi xây bảng đó: làm dữ liệu nguồn cho
 * tính năng gợi ý giá bán.
 *
 * Quy ước tính (nhất quán với cách hệ thống tính lợi nhuận đơn hàng thật):
 * - Phí sàn & phí đóng gói được tính trên MỌI đơn dự kiến (kể cả đơn hoàn),
 *   vì TikTok vẫn thu hoa hồng + phí giao dịch trên đơn hoàn (từ 01/04/2026).
 * - Tỷ lệ hoàn trả làm giảm DOANH THU KỲ VỌNG (theo xác suất), không làm
 *   giảm phí sàn/phí đóng gói đã phát sinh.
 * - Điểm hòa vốn ROAS = Giá bán / Lợi nhuận gộp — ROAS tối thiểu cần đạt để
 *   chi phí Ads không vượt quá phần lợi nhuận gộp còn lại.
 */
function ProfitCalculatorSection({ onGoToImport }) {
  const { state } = useApp();

  const defaultFeeRate = useMemo(() => {
    return state.feeConfig
      .filter((f) => f.required && f.valueType === "percent")
      .reduce((s, f) => s + f.value, 0);
  }, [state.feeConfig]);

  const [form, setForm] = useState({
    productName: "",
    platform: PLATFORM.TIKTOK,
    category: "",
    listingPrice: 0,
    cogs: 0,
    packagingFee: DEFAULT_PACKAGING_FEE,
    platformFeeRate: defaultFeeRate,
    targetRoas: 5,
    returnRate: 0,
  });

  function setField(key, value) {
    setForm((f) => ({ ...f, [key]: value }));
  }

  const result = useMemo(() => {
    const price = Number(form.listingPrice) || 0;
    const cogs = Number(form.cogs) || 0;
    const packaging = Number(form.packagingFee) || 0;
    const feeRate = Number(form.platformFeeRate) || 0;
    const targetRoas = Number(form.targetRoas) || 0;
    const returnRate = Math.min(100, Math.max(0, Number(form.returnRate) || 0));

    const platformFeeAmount = (price * feeRate) / 100;
    const expectedRevenue = price * (1 - returnRate / 100);
    const adsCostRate = targetRoas > 0 ? 100 / targetRoas : 0;
    const adsAmount = (price * adsCostRate) / 100;

    const totalVariableCost = cogs + packaging + platformFeeAmount + adsAmount;

    const grossProfit = expectedRevenue - cogs - packaging - platformFeeAmount;
    const grossMarginPct = price > 0 ? (grossProfit / price) * 100 : 0;

    const netProfit = grossProfit - adsAmount;
    const netMarginPct = price > 0 ? (netProfit / price) * 100 : 0;

    // ROAS hòa vốn: Phí Ads tối đa có thể chi (= LN gộp) tương ứng ROAS = Giá bán / LN gộp
    const breakEvenRoas = grossProfit > 0 ? price / grossProfit : null;

    let tier = { tone: "gray", label: "Nhập đủ Giá bán & Giá vốn để xem đánh giá" };
    if (price > 0) {
      if (netMarginPct > 20) tier = { tone: "green", label: "🟢 Sản phẩm Tiềm Năng Cao" };
      else if (netMarginPct >= 10) tier = { tone: "amber", label: "🟡 Mức Lợi Nhuận Trung Bình" };
      else tier = { tone: "red", label: "🔴 Cảnh Báo Biên Lợi Nhuận Thấp – Rủi Ro Cao" };
    }

    return { platformFeeAmount, expectedRevenue, adsCostRate, adsAmount, totalVariableCost, grossProfit, grossMarginPct, netProfit, netMarginPct, breakEvenRoas, tier };
  }, [form]);

  const tierCardClass = {
    gray: "bg-slate-50 border-slate-100",
    green: "bg-emerald-50 border-emerald-100",
    amber: "bg-amber-50 border-amber-100",
    red: "bg-rose-50 border-rose-100",
  }[result.tier.tone];
  const tierTextClass = {
    gray: "text-slate-500",
    green: "text-emerald-700",
    amber: "text-amber-700",
    red: "text-rose-700",
  }[result.tier.tone];

  return (
    <div className="space-y-4">
      <div>
        <h2 className="text-lg font-semibold text-blue-950 flex items-center gap-2"><Calculator size={18} /> Tính Lợi Nhuận Dự Phóng</h2>
        <p className="text-xs text-slate-400 mt-0.5">Ước tính lợi nhuận cho 1 sản phẩm/lô hàng mới TRƯỚC khi lên đơn thật — không ảnh hưởng dữ liệu đơn hàng đã import.</p>
      </div>

      <div className="grid lg:grid-cols-2 gap-4">
        <Card className="p-4 space-y-3">
          <h3 className="text-sm font-semibold text-blue-950 mb-1">Thông tin đầu vào</h3>

          <div>
            <label className="text-xs text-slate-500 mb-1 block">Tên sản phẩm</label>
            <input value={form.productName} onChange={(e) => setField("productName", e.target.value)} placeholder="Vd: Kính mát ELLA Round" className="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm" />
          </div>

          <div className="grid grid-cols-2 gap-3">
            <div>
              <label className="text-xs text-slate-500 mb-1 block">Sàn</label>
              <select value={form.platform} onChange={(e) => setField("platform", e.target.value)} className="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm">
                {CALC_PLATFORMS.map((p) => <option key={p.key} value={p.key}>{p.label}</option>)}
              </select>
            </div>
            <div>
              <label className="text-xs text-slate-500 mb-1 block">Ngành hàng</label>
              <input value={form.category} onChange={(e) => setField("category", e.target.value)} placeholder="Vd: Phụ kiện thời trang" className="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm" />
            </div>
          </div>

          <div className="grid grid-cols-2 gap-3">
            <div>
              <label className="text-xs text-slate-500 mb-1 block">Giá bán niêm yết (đ)</label>
              <input type="number" value={form.listingPrice} onChange={(e) => setField("listingPrice", e.target.value)} className="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm" />
            </div>
            <div>
              <label className="text-xs text-slate-500 mb-1 block">Giá vốn COGS (đ)</label>
              <input type="number" value={form.cogs} onChange={(e) => setField("cogs", e.target.value)} className="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm" />
            </div>
          </div>

          <div className="grid grid-cols-2 gap-3">
            <div>
              <label className="text-xs text-slate-500 mb-1 block">Phí đóng gói (đ)</label>
              <input type="number" value={form.packagingFee} onChange={(e) => setField("packagingFee", e.target.value)} className="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm" />
            </div>
            <div>
              <label className="text-xs text-slate-500 mb-1 block flex items-center justify-between">
                <span>Tỷ lệ phí sàn (%)</span>
                <button type="button" onClick={() => setField("platformFeeRate", defaultFeeRate)} className="text-blue-700 text-[10px] hover:underline">Dùng mức từ Config</button>
              </label>
              <input type="number" value={form.platformFeeRate} onChange={(e) => setField("platformFeeRate", e.target.value)} className="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm" />
            </div>
          </div>

          <div className="grid grid-cols-2 gap-3">
            <div>
              <label className="text-xs text-slate-500 mb-1 block">ROAS mục tiêu (lần)</label>
              <input type="number" step="0.1" value={form.targetRoas} onChange={(e) => setField("targetRoas", e.target.value)} className="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm" />
              <p className="text-[10px] text-slate-400 mt-1">≈ {formatPercent(result.adsCostRate)} doanh thu chi cho Ads</p>
            </div>
            <div>
              <label className="text-xs text-slate-500 mb-1 block">Tỷ lệ hoàn trả dự kiến (%)</label>
              <input type="number" value={form.returnRate} onChange={(e) => setField("returnRate", e.target.value)} className="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm" />
            </div>
          </div>
        </Card>

        <div className="space-y-4">
          <Card className="p-4">
            <h3 className="text-sm font-semibold text-blue-950 mb-3">Kết quả ước tính</h3>
            <div className="grid grid-cols-2 gap-3 text-sm">
              <div className="bg-slate-50 rounded-xl p-3">
                <div className="text-[11px] text-slate-400 mb-1">Tổng chi phí biến đổi</div>
                <div className="font-semibold text-blue-950">{formatVND(result.totalVariableCost)}</div>
              </div>
              <div className="bg-slate-50 rounded-xl p-3">
                <div className="text-[11px] text-slate-400 mb-1">Doanh thu kỳ vọng</div>
                <div className="font-semibold text-blue-950">{formatVND(result.expectedRevenue)}</div>
              </div>
              <div className="bg-slate-50 rounded-xl p-3">
                <div className="text-[11px] text-slate-400 mb-1">Lợi nhuận gộp</div>
                <div className={`font-semibold ${result.grossProfit >= 0 ? "text-emerald-600" : "text-rose-600"}`}>{formatVND(result.grossProfit)}</div>
                <div className="text-[11px] text-slate-400">Biên gộp {formatPercent(result.grossMarginPct)}</div>
              </div>
              <div className="bg-slate-50 rounded-xl p-3">
                <div className="text-[11px] text-slate-400 mb-1">Lợi nhuận ròng</div>
                <div className={`font-semibold ${result.netProfit >= 0 ? "text-emerald-600" : "text-rose-600"}`}>{formatVND(result.netProfit)}</div>
                <div className="text-[11px] text-slate-400">Biên ròng {formatPercent(result.netMarginPct)}</div>
              </div>
              <div className="col-span-2 bg-slate-50 rounded-xl p-3 flex items-center justify-between">
                <div>
                  <div className="text-[11px] text-slate-400 mb-1">Điểm hòa vốn ROAS</div>
                  <div className="font-semibold text-blue-950">{result.breakEvenRoas ? result.breakEvenRoas.toFixed(2) + "x" : "—"}</div>
                </div>
                <Target size={22} className="text-slate-300" />
              </div>
            </div>
          </Card>

          <Card className={`p-4 border ${tierCardClass}`}>
            <div className={`text-sm font-semibold ${tierTextClass}`}>{result.tier.label}</div>
            {form.listingPrice > 0 && (
              <p className="text-xs text-slate-500 mt-1.5 leading-relaxed">
                Biên lợi nhuận ròng {formatPercent(result.netMarginPct)}
                {result.breakEvenRoas ? ` — cần ROAS thực tế đạt tối thiểu ${result.breakEvenRoas.toFixed(2)}x để không lỗ khi chạy Ads.` : "."}
              </p>
            )}
          </Card>
        </div>
      </div>
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
          {activeSection === "calculator" && <ProfitCalculatorSection onGoToImport={goToImport} />}
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
