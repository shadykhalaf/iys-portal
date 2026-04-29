import { useState, useMemo, useRef, useEffect } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";

const SUPABASE_URL = "https://kbhwmamypvzbyngsmanl.supabase.co";
const SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImtiaHdtYW15cHZ6YnluZ3NtYW5sIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzczMjQyMDYsImV4cCI6MjA5MjkwMDIwNn0.Hv6sTULF6_ob-KELD8eMIYSbRobM_MsGQkDElPsjjoI";
const SUPABASE_CONFIGURED = !SUPABASE_URL.includes("YOUR_PROJECT");
const ORDER_EMAIL = "moemen@inyourshoe.com";
const RECEIPT_EMAIL = "shady@inyourshoe.com";

const CURRENCY_RATES = {
  EGP: { rate: 1, symbol: "EGP", flag: "🇪🇬" },
  USD: { rate: 0.0204, symbol: "USD", flag: "🇺🇸" },
  SAR: { rate: 0.0765, symbol: "SAR", flag: "🇸🇦" },
  AED: { rate: 0.0749, symbol: "AED", flag: "🇦🇪" },
};

// ─── SUPABASE HELPERS ───
async function sbFetch(path, opts = {}) {
  const controller = new AbortController();
  const t = setTimeout(() => controller.abort(), 10000);
  try {
    const res = await fetch(`${SUPABASE_URL}/rest/v1${path}`, {
      signal: controller.signal,
      ...opts,
      headers: {
        apikey: SUPABASE_ANON_KEY,
        Authorization: `Bearer ${SUPABASE_ANON_KEY}`,
        "Content-Type": "application/json",
        Prefer: "return=representation",
        ...(opts.headers || {}),
      },
    });
    clearTimeout(t);
    const text = await res.text();
    try { return { data: JSON.parse(text), ok: res.ok, status: res.status }; }
    catch { return { data: text, ok: res.ok, status: res.status }; }
  } catch (e) {
    clearTimeout(t);
    return { data: e.message, ok: false, status: 0 };
  }
}

async function authFetch(path, body) {
  const controller = new AbortController();
  const t = setTimeout(() => controller.abort(), 10000);
  try {
    const res = await fetch(`${SUPABASE_URL}/auth/v1${path}`, {
      method: "POST", signal: controller.signal,
      headers: { apikey: SUPABASE_ANON_KEY, "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });
    clearTimeout(t);
    return res.json();
  } catch (e) {
    clearTimeout(t);
    return { error: "timeout", msg: e.name === "AbortError" ? "Request timed out." : e.message };
  }
}

// ─── CATALOG: save/load from Supabase ───
// We store the entire catalog as a single JSON row in a `iys_catalog` table
async function saveCatalog(rows, imageCol, token) {
  return sbFetch("/iys_catalog?id=eq.1", {
    method: "PATCH",
    headers: { Authorization: `Bearer ${token}`, Prefer: "resolution=merge-duplicates,return=representation" },
    body: JSON.stringify({ data: rows, image_col: imageCol, updated_at: new Date().toISOString() }),
  }).then(async (r) => {
    if (r.status === 404 || (Array.isArray(r.data) && r.data.length === 0)) {
      // Row doesn't exist yet — insert
      return sbFetch("/iys_catalog", {
        method: "POST",
        headers: { Authorization: `Bearer ${token}`, Prefer: "resolution=merge-duplicates,return=representation" },
        body: JSON.stringify({ id: 1, data: rows, image_col: imageCol, updated_at: new Date().toISOString() }),
      });
    }
    return r;
  });
}

async function loadCatalog() {
  return sbFetch("/iys_catalog?id=eq.1&select=data,image_col");
}

// ─── INVENTORY: load & decrement from Supabase ───
async function loadInventory() {
  return sbFetch("/iys_inventory?select=*");
}

async function decrementInventory(items, token) {
  // items: [{ sku, qty }]
  for (const item of items) {
    await sbFetch(`/iys_inventory?sku=eq.${encodeURIComponent(item.sku)}`, {
      method: "PATCH",
      headers: { Authorization: `Bearer ${token}` },
      body: JSON.stringify({ reserved: item.qty }),
    });
  }
}

async function saveInventorySnapshot(products, token) {
  // Upsert all SKU stock levels
  const rows = [];
  products.forEach(p => p.variants.forEach(v => {
    if (v.sku) rows.push({ sku: v.sku, product_title: p.title, size: v.size, stock: v.rawStock, reserved: 0 });
  }));
  if (!rows.length) return;
  return sbFetch("/iys_inventory", {
    method: "POST",
    headers: { Authorization: `Bearer ${token}`, Prefer: "resolution=merge-duplicates,return=representation" },
    body: JSON.stringify(rows),
  });
}

// ─── DATA PROCESSING ───
function processProducts(raw, imageCol) {
  const grouped = {};
  raw.forEach((r) => {
    const title = (r["Product Title"] || r["Title"] || "").trim();
    if (!title) return;
    const imgUrl = imageCol ? (r[imageCol] || "") : (r["Product Image url"] || r["Product Image URL"] || r["Image Src"] || "");
    if (!grouped[title]) {
      grouped[title] = {
        title, type: r["Product type"] || r["Type"] || "Other",
        image: imgUrl.trim(),
        retailPrice: parseFloat(r["Price"] || r["Variant Price"] || 0) || 0,
        discount: parseFloat(r["Bulk Discount"] || 0.35),
        variants: [],
      };
    }
    const rawQty = parseInt(r["Quantity at location 'Heliopolis Warehouse'"] || 0) || 0;
    grouped[title].variants.push({
      size: r["Variant title"] || r["Option1 Value"] || "Default Title",
      sku: r["SKU"] || r["Variant SKU"] || "",
      barcode: r["Barcode"] || r["Variant Barcode"] || "",
      rawStock: rawQty,
    });
  });
  return Object.values(grouped).map((p) => ({ ...p, bulkPrice: Math.round(p.retailPrice * (1 - p.discount)) }));
}

// Merge live inventory reservations into products
function applyInventory(products, inventory) {
  if (!inventory || !inventory.length) return products.map(p => ({
    ...p, variants: p.variants.map(v => ({ ...v, available: Math.max(0, v.rawStock - 5), inStock: v.rawStock > 5 }))
  }));
  const invMap = {};
  inventory.forEach(i => { invMap[i.sku] = i; });
  return products.map(p => ({
    ...p,
    variants: p.variants.map(v => {
      const inv = invMap[v.sku];
      const reserved = inv ? (inv.reserved || 0) : 0;
      const effective = Math.max(0, v.rawStock - 5 - reserved);
      return { ...v, available: effective, inStock: effective > 0 };
    })
  }));
}

// ─── CSV / XLSX PARSER ───
function parseFile(file, callback) {
  const ext = file.name.split(".").pop().toLowerCase();
  if (ext === "csv") {
    Papa.parse(file, {
      header: true, skipEmptyLines: true,
      complete: (res) => callback(null, [{ name: "Sheet1", rows: res.data }]),
      error: (e) => callback(e.message),
    });
  } else if (ext === "xlsx" || ext === "xls") {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: "array" });
        const sheets = wb.SheetNames.map(name => ({
          name,
          rows: XLSX.utils.sheet_to_json(wb.Sheets[name], { defval: "" }),
        })).filter(s => s.rows.length > 0);
        callback(null, sheets);
      } catch (err) { callback(err.message); }
    };
    reader.readAsArrayBuffer(file);
  } else {
    callback("Unsupported file type. Use .csv, .xlsx, or .xls");
  }
}

// ─── ORDER HELPERS ───
function generateOrderCSV(cart, products, customerInfo, currency) {
  const cur = CURRENCY_RATES[currency];
  let totalEGP = 0;
  const rows = [];
  Object.entries(cart).forEach(([key, qty]) => {
    if (qty <= 0) return;
    const [pTitle, size] = key.split("|||");
    const prod = products.find(p => p.title === pTitle);
    if (!prod) return;
    const v = prod.variants.find(v => v.size === size);
    if (!v) return;
    const lt = prod.bulkPrice * qty;
    totalEGP += lt;
    rows.push({ type: prod.type, title: prod.title, variant: size, sku: v.sku, barcode: v.barcode, retailPrice: prod.retailPrice, bulkPrice: prod.bulkPrice, qty, lt });
  });
  const cur2dp = n => n.toFixed(currency === "EGP" ? 0 : 2);
  const L = ["ORDERED ITEMS",
    `Product Type,Product Title,Variant,SKU,Barcode,Retail (EGP),Bulk (EGP),Bulk (${currency}),Qty,Total (EGP),Total (${currency})`];
  rows.forEach(r => L.push(`"${r.type}","${r.title}","${r.variant}","${r.sku}","${r.barcode}",${r.retailPrice},${r.bulkPrice},${cur2dp(r.bulkPrice * cur.rate)},${r.qty},${r.lt},${cur2dp(r.lt * cur.rate)}`));
  L.push("", "ORDER OVERVIEW",
    `Customer Name,"${customerInfo.name}"`,
    `Contact,"${customerInfo.contact}"`,
    `Currency,${currency}`,
    `Total Items,${rows.reduce((s, r) => s + r.qty, 0)}`,
    `Total (EGP),"${totalEGP.toLocaleString()}"`,
    ...(currency !== "EGP" ? [`Total (${currency}),"${cur2dp(totalEGP * cur.rate)}"`] : []),
    `Order Date,"${new Date().toLocaleDateString("en-GB")}"`);
  return L.join("\r\n");
}

function doDownloadCSV(cart, products, customerInfo, currency) {
  const csv = generateOrderCSV(cart, products, customerInfo, currency);
  const blob = new Blob(["\uFEFF" + csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = `IYS_Order_${customerInfo.name.replace(/[^a-zA-Z0-9]/g, "_")}_${new Date().toISOString().split("T")[0]}.csv`;
  document.body.appendChild(a); a.click(); document.body.removeChild(a); URL.revokeObjectURL(url);
}

async function uploadReceiptToStorage(csvContent, filename) {
  try {
    const blob = new Blob(["\uFEFF" + csvContent], { type: "text/csv;charset=utf-8;" });
    const controller = new AbortController();
    const t = setTimeout(() => controller.abort(), 15000);
    const res = await fetch(`${SUPABASE_URL}/storage/v1/object/iys-receipts/${filename}`, {
      method: "POST",
      signal: controller.signal,
      headers: {
        apikey: SUPABASE_ANON_KEY,
        Authorization: `Bearer ${SUPABASE_ANON_KEY}`,
        "Content-Type": "text/csv",
        "x-upsert": "true",
      },
      body: blob,
    });
    clearTimeout(t);
    if (!res.ok) { const err = await res.text(); console.error("Receipt upload failed:", res.status, err); }
    return res.ok;
  } catch (e) {
    console.error("Receipt upload error:", e);
    return false;
  }
}

async function saveReceiptRecord(filename, customerInfo, totalEGP, currency) {
  return sbFetch("/iys_receipts", {
    method: "POST",
    body: JSON.stringify({
      filename,
      customer_name: customerInfo.name,
      customer_contact: customerInfo.contact,
      currency,
      total_egp: totalEGP,
      receipt_email: RECEIPT_EMAIL,
      created_at: new Date().toISOString(),
    }),
  });
}


async function saveOrder(cart, products, customerInfo, currency, token) {
  const cur = CURRENCY_RATES[currency];
  let totalEGP = 0;
  const items = [];
  Object.entries(cart).forEach(([key, qty]) => {
    if (!qty) return;
    const [pTitle, size] = key.split("|||");
    const prod = products.find(p => p.title === pTitle);
    if (!prod) return;
    const v = prod.variants.find(v => v.size === size);
    const lt = prod.bulkPrice * qty;
    totalEGP += lt;
    items.push({ title: pTitle, type: prod.type, size, sku: v?.sku || "", barcode: v?.barcode || "", bulkPrice: prod.bulkPrice, retailPrice: prod.retailPrice, qty, lineTotalEGP: lt, lineTotalConverted: parseFloat((lt * cur.rate).toFixed(2)) });
  });
  return sbFetch("/iys_orders", {
    method: "POST",
    headers: { Authorization: `Bearer ${token}` },
    body: JSON.stringify({
      customer_name: customerInfo.name,
      customer_contact: customerInfo.contact,
      currency,
      total_egp: totalEGP,
      total_converted: parseFloat((totalEGP * cur.rate).toFixed(2)),
      items,
      order_date: new Date().toISOString(),
      status: "pending",
    }),
  });
}

// ─── ICONS ───
const Ic = {
  Cart: () => <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="9" cy="21" r="1"/><circle cx="20" cy="21" r="1"/><path d="M1 1h4l2.68 13.39a2 2 0 0 0 2 1.61h9.72a2 2 0 0 0 2-1.61L23 6H6"/></svg>,
  Check: () => <svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><polyline points="20 6 9 17 4 12"/></svg>,
  Download: () => <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>,
  X: () => <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>,
  Upload: () => <svg width="44" height="44" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>,
  Trash: () => <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>,
  Shield: () => <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/></svg>,
  User: () => <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/></svg>,
  LogOut: () => <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"/><polyline points="16 17 21 12 16 7"/><line x1="21" y1="12" x2="9" y2="12"/></svg>,
  Plus: () => <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>,
  Refresh: () => <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polyline points="23 4 23 10 17 10"/><path d="M20.49 15a9 9 0 1 1-2.12-9.36L23 10"/></svg>,
};

const fonts = <link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700;800&family=Playfair+Display:wght@700;800;900&display=swap" rel="stylesheet" />;

function Field({ label, value, onChange, type = "text", placeholder, error, disabled }) {
  return (
    <div style={{ marginBottom: 14 }}>
      <label style={{ fontSize: 11, fontWeight: 700, color: "#666", display: "block", marginBottom: 5, textTransform: "uppercase", letterSpacing: "0.5px" }}>{label}</label>
      <input type={type} value={value} onChange={onChange} placeholder={placeholder} disabled={disabled}
        style={{ width: "100%", padding: "10px 13px", border: `1px solid ${error ? "#e25822" : "#e0dbd5"}`, borderRadius: 9, fontSize: 14, outline: "none", fontFamily: "'DM Sans', sans-serif", boxSizing: "border-box", background: disabled ? "#f9f9f9" : error ? "#fef8f5" : "#fff" }} />
      {error && <div style={{ fontSize: 11, color: "#e25822", marginTop: 3 }}>{error}</div>}
    </div>
  );
}

// ─── PRODUCT CARD ───
function ProductCard({ product, cart, setCart, currency }) {
  const cur = CURRENCY_RATES[currency];
  const dp = egp => `${cur.symbol} ${(egp * cur.rate).toFixed(currency === "EGP" ? 0 : 2)}`;
  const [imgErr, setImgErr] = useState(false);
  const cartTotal = product.variants.reduce((s, v) => s + (cart[`${product.title}|||${v.size}`] || 0), 0);
  const hasStock = product.variants.some(v => v.inStock);

  return (
    <div style={{ background: "#fff", borderRadius: 14, overflow: "hidden", border: "1px solid #e8e4df", position: "relative", display: "flex", flexDirection: "column" }}>
      {cartTotal > 0 && <div style={{ position: "absolute", top: 10, right: 10, zIndex: 2, background: "#1a1a1a", color: "#fff", borderRadius: 20, padding: "3px 11px", fontSize: 11, fontWeight: 700 }}>{cartTotal} in cart</div>}
      <div style={{ width: "100%", aspectRatio: "1/1.1", background: "#f5f1ec", overflow: "hidden", position: "relative", display: "flex", alignItems: "center", justifyContent: "center" }}>
        {product.image && !imgErr
          ? <img src={product.image} alt={product.title} onError={() => setImgErr(true)} style={{ width: "100%", height: "100%", objectFit: "cover" }} />
          : <div style={{ color: "#ccc", textAlign: "center", padding: 16, fontSize: 11 }}><div style={{ fontSize: 32, marginBottom: 6 }}>👟</div>{product.title}</div>}
        {!hasStock && <div style={{ position: "absolute", inset: 0, background: "rgba(0,0,0,0.55)", display: "flex", alignItems: "center", justifyContent: "center", color: "#fff", fontWeight: 700, fontSize: 16 }}>OUT OF STOCK</div>}
      </div>
      <div style={{ padding: "14px 16px 16px", flex: 1, display: "flex", flexDirection: "column" }}>
        <div style={{ fontSize: 10, fontWeight: 700, color: "#999", textTransform: "uppercase", letterSpacing: "1px", marginBottom: 3 }}>{product.type}</div>
        <div style={{ fontSize: 14, fontWeight: 700, color: "#1a1a1a", marginBottom: 8, lineHeight: 1.3 }}>{product.title}</div>
        <div style={{ display: "flex", alignItems: "baseline", gap: 7, marginBottom: 10 }}>
          <span style={{ fontSize: 17, fontWeight: 800 }}>{dp(product.bulkPrice)}</span>
          <span style={{ fontSize: 12, color: "#bbb", textDecoration: "line-through" }}>{dp(product.retailPrice)}</span>
          <span style={{ fontSize: 10, fontWeight: 700, color: "#e25822", background: "#fef3ee", padding: "2px 6px", borderRadius: 5 }}>-{Math.round(product.discount * 100)}%</span>
        </div>
        {product.variants.map(v => {
          const key = `${product.title}|||${v.size}`;
          const qty = cart[key] || 0;
          return (
            <div key={v.size} style={{ display: "flex", alignItems: "center", gap: 7, padding: "6px 0", borderTop: "1px solid #f0ece7" }}>
              <div style={{ minWidth: 90, fontSize: 12, fontWeight: 600, color: v.inStock ? "#1a1a1a" : "#ccc" }}>{v.size}</div>
              {v.inStock ? (
                <>
                  <div style={{ fontSize: 11, color: "#aaa", flex: 1 }}>max <b style={{ color: "#555" }}>{v.available}</b></div>
                  <div style={{ display: "flex", alignItems: "center" }}>
                    <button onClick={() => setCart(p => { const n = { ...p }; if ((n[key] || 0) > 0) { n[key]--; if (!n[key]) delete n[key]; } return n; })}
                      style={{ width: 25, height: 25, border: "1px solid #ddd", borderRadius: "7px 0 0 7px", background: qty > 0 ? "#f5f1ec" : "#fafafa", cursor: qty > 0 ? "pointer" : "default", fontSize: 15, fontWeight: 700, color: qty > 0 ? "#1a1a1a" : "#ccc", display: "flex", alignItems: "center", justifyContent: "center" }}>−</button>
                    <div style={{ width: 30, height: 25, border: "1px solid #ddd", borderLeft: "none", borderRight: "none", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 12, fontWeight: 700, background: "#fff" }}>{qty}</div>
                    <button onClick={() => setCart(p => { const n = { ...p }; if ((n[key] || 0) < v.available) n[key] = (n[key] || 0) + 1; return n; })}
                      style={{ width: 25, height: 25, border: "1px solid #ddd", borderRadius: "0 7px 7px 0", background: qty < v.available ? "#1a1a1a" : "#eee", cursor: qty < v.available ? "pointer" : "default", fontSize: 15, fontWeight: 700, color: qty < v.available ? "#fff" : "#ccc", display: "flex", alignItems: "center", justifyContent: "center" }}>+</button>
                  </div>
                </>
              ) : <div style={{ fontSize: 11, color: "#ccc", fontWeight: 600 }}>Out of Stock</div>}
            </div>
          );
        })}
      </div>
    </div>
  );
}

// ══════════════════════════════════════
// MAIN APP
// ══════════════════════════════════════
export default function App() {
  const [session, setSession] = useState(null);
  const [authView, setAuthView] = useState("login");
  const [authForm, setAuthForm] = useState({ email: "", password: "", name: "" });
  const [authLoading, setAuthLoading] = useState(false);
  const [authError, setAuthError] = useState("");

  // Catalog
  const [rawData, setRawData] = useState(null);
  const [imageCol, setImageCol] = useState("");
  const [csvHeaders, setCsvHeaders] = useState([]);
  const [uploadedSheets, setUploadedSheets] = useState([]); // [{ name, rows }]
  const [activeSheets, setActiveSheets] = useState([]); // which sheet names are enabled
  const [uploadFile, setUploadFile] = useState("");
  const [uploadStats, setUploadStats] = useState(null);
  const [parseError, setParseError] = useState("");
  const [saving, setSaving] = useState(false);
  const [catalogLoading, setCatalogLoading] = useState(false);
  const fileRef = useRef(null);

  // Inventory
  const [inventory, setInventory] = useState([]);

  // Admin users
  const [users, setUsers] = useState([]);
  const [adminTab, setAdminTab] = useState("data"); // data | orders | users
  const [orders, setOrders] = useState([]);
  const [ordersLoading, setOrdersLoading] = useState(false);
  const [expandedOrder, setExpandedOrder] = useState(null);

  // Shop
  const [shopPage, setShopPage] = useState("shop");
  const [cart, setCart] = useState({});
  const [currency, setCurrency] = useState("EGP");
  const [search, setSearch] = useState("");
  const [filterType, setFilterType] = useState("All");
  const [customerInfo, setCustomerInfo] = useState({ name: "", contact: "" });
  const [custErrors, setCustErrors] = useState({});

  const isAdmin = session?.user?.role === "admin";

  // Merge all active sheet rows
  const mergedRows = useMemo(() => {
    if (!uploadedSheets.length) return rawData || [];
    return uploadedSheets.filter(s => activeSheets.includes(s.name)).flatMap(s => s.rows);
  }, [uploadedSheets, activeSheets, rawData]);

  const products = useMemo(() => mergedRows.length ? processProducts(mergedRows, imageCol) : [], [mergedRows, imageCol]);
  const liveProducts = useMemo(() => applyInventory(products, inventory), [products, inventory]);
  const productTypes = useMemo(() => ["All", ...[...new Set(liveProducts.map(p => p.type))].sort()], [liveProducts]);
  const filtered = useMemo(() => liveProducts.filter(p =>
    (filterType === "All" || p.type === filterType) &&
    (!search || p.title.toLowerCase().includes(search.toLowerCase()))
  ), [liveProducts, filterType, search]);

  const cur = CURRENCY_RATES[currency];
  const dp = egp => `${cur.symbol} ${(egp * cur.rate).toFixed(currency === "EGP" ? 0 : 2)}`;
  const cartItems = useMemo(() => Object.entries(cart).filter(([, q]) => q > 0).map(([key, qty]) => {
    const [pTitle, size] = key.split("|||");
    const prod = liveProducts.find(p => p.title === pTitle);
    return { key, pTitle, size, qty, prod };
  }), [cart, liveProducts]);
  const totalEGP = cartItems.reduce((s, i) => s + (i.prod?.bulkPrice || 0) * i.qty, 0);
  const totalItems = cartItems.reduce((s, i) => s + i.qty, 0);

  // ── Load session ──
  useEffect(() => {
    const stored = sessionStorage.getItem("iys_session");
    if (stored) { try { setSession(JSON.parse(stored)); } catch {} }
  }, []);

  // ── Load catalog on login ──
  useEffect(() => {
    if (!session) return;
    setCatalogLoading(true);
    Promise.all([loadCatalog(), loadInventory()]).then(([catRes, invRes]) => {
      if (catRes.ok && Array.isArray(catRes.data) && catRes.data.length > 0) {
        const { data, image_col } = catRes.data[0];
        if (data && data.length) {
          setRawData(data);
          const cols = Object.keys(data[0]);
          setCsvHeaders(cols);
          if (image_col) setImageCol(image_col);
        }
      }
      if (invRes.ok && Array.isArray(invRes.data)) setInventory(invRes.data);
      setCatalogLoading(false);
    });
  }, [session]);

  // ── Load users (admin) ──
  useEffect(() => {
    if (!isAdmin || adminTab !== "users") return;
    sbFetch("/iys_users?select=*&order=created_at.desc").then(r => { if (r.ok) setUsers(r.data); });
  }, [isAdmin, adminTab]);

  // ── Load orders (admin) ──
  useEffect(() => {
    if (!isAdmin || adminTab !== "orders") return;
    setOrdersLoading(true);
    sbFetch("/iys_orders?select=*&order=order_date.desc").then(r => {
      if (r.ok) setOrders(r.data);
      setOrdersLoading(false);
    });
  }, [isAdmin, adminTab]);

  // ── AUTH ──
  const upsertProfile = (userId, email, name, token) =>
    sbFetch("/iys_users?on_conflict=id", {
      method: "POST",
      headers: { Authorization: `Bearer ${token}`, Prefer: "resolution=merge-duplicates,return=representation" },
      body: JSON.stringify({ id: userId, email, name: name || email.split("@")[0], role: "user" }),
    });

  const handleLogin = async () => {
    if (!SUPABASE_CONFIGURED) { setAuthError("Supabase not configured."); return; }
    setAuthLoading(true); setAuthError("");
    try {
      const res = await authFetch("/token?grant_type=password", { email: authForm.email, password: authForm.password });
      if (res.error || !res.access_token) { setAuthError(res.msg || res.error_description || "Invalid credentials."); setAuthLoading(false); return; }
      const { access_token: token, user } = res;
      const roleRes = await sbFetch(`/iys_users?id=eq.${user.id}&select=role`);
      let role = "user";
      if (roleRes.ok && roleRes.data?.length > 0) role = roleRes.data[0].role || "user";
      else await upsertProfile(user.id, user.email, user.email.split("@")[0], token);
      const s = { token, user: { id: user.id, email: user.email, role } };
      setSession(s); sessionStorage.setItem("iys_session", JSON.stringify(s));
    } catch (e) { setAuthError("Error: " + e.message); }
    setAuthLoading(false);
  };

  const handleSignup = async () => {
    if (!SUPABASE_CONFIGURED) { setAuthError("Supabase not configured."); return; }
    if (!authForm.name.trim()) { setAuthError("Name is required."); return; }
    setAuthLoading(true); setAuthError("");
    try {
      const res = await authFetch("/signup", { email: authForm.email, password: authForm.password });
      if (res.error_code === "user_already_exists") { setAuthView("login"); setAuthError("Account exists — please sign in."); setAuthLoading(false); return; }
      if (res.error || res.error_code) { setAuthError(res.msg || res.error_description || "Signup failed."); setAuthLoading(false); return; }
      const userId = res.user?.id; const token = res.access_token;
      if (userId && token) await upsertProfile(userId, authForm.email, authForm.name, token);
      setAuthView("login"); setAuthError("✓ Account created! Please sign in.");
    } catch (e) { setAuthError("Error: " + e.message); }
    setAuthLoading(false);
  };

  const handleLogout = () => { setSession(null); sessionStorage.removeItem("iys_session"); setCart({}); setShopPage("shop"); setRawData(null); setInventory([]); };
  const promoteUser = async (id) => {
    await sbFetch(`/iys_users?id=eq.${id}`, { method: "PATCH", headers: { Authorization: `Bearer ${session.token}` }, body: JSON.stringify({ role: "admin" }) });
    setUsers(u => u.map(x => x.id === id ? { ...x, role: "admin" } : x));
  };

  // ── FILE UPLOAD ──
  const handleFile = (file) => {
    if (!file) return;
    setParseError(""); setUploadFile(file.name);
    parseFile(file, (err, sheets) => {
      if (err) { setParseError(err); return; }
      setUploadedSheets(prev => {
        const existingNames = prev.map(s => s.name);
        const newSheets = sheets.filter(s => !existingNames.includes(s.name));
        return [...prev, ...newSheets];
      });
      setActiveSheets(prev => [...new Set([...prev, ...sheets.map(s => s.name)])]);
      const allRows = sheets.flatMap(s => s.rows);
      if (allRows.length) {
        const cols = Object.keys(allRows[0]);
        setCsvHeaders(cols);
        const detected = cols.find(c => /image/i.test(c) || /photo/i.test(c));
        if (detected && !imageCol) setImageCol(detected);
        const prods = processProducts(allRows, detected || imageCol);
        setUploadStats({ rows: allRows.length, products: prods.length, sheets: sheets.length });
      }
    });
  };

  const saveCatalogToSupabase = async () => {
    setSaving(true);
    const allRows = uploadedSheets.filter(s => activeSheets.includes(s.name)).flatMap(s => s.rows);
    setRawData(allRows);
    const prods = processProducts(allRows, imageCol);
    await Promise.all([
      saveCatalog(allRows, imageCol, session.token),
      saveInventorySnapshot(prods, session.token),
    ]);
    const invRes = await loadInventory();
    if (invRes.ok) setInventory(invRes.data);
    setSaving(false);
  };

  const clearCatalog = async () => {
    setRawData(null); setUploadedSheets([]); setActiveSheets([]); setUploadFile(""); setUploadStats(null); setCsvHeaders([]); setImageCol("");
    await saveCatalog([], "", session.token);
  };

  const handleConfirmOrder = async () => {
    const e = {};
    if (!customerInfo.name.trim()) e.name = "Required";
    if (!customerInfo.contact.trim()) e.contact = "Required";
    setCustErrors(e);
    if (Object.keys(e).length) return;

    // Decrement inventory in Supabase
    const itemsToDecrement = cartItems.map(i => ({ sku: i.prod?.variants.find(v => v.size === i.size)?.sku, qty: i.qty })).filter(i => i.sku);
    if (itemsToDecrement.length) await decrementInventory(itemsToDecrement, session.token);

    // Refresh inventory so vendor portal updates immediately
    const invRes = await loadInventory();
    if (invRes.ok) setInventory(invRes.data);

    const csvContent = generateOrderCSV(cart, liveProducts, customerInfo, currency);
    const filename = `IYS_Order_${customerInfo.name.replace(/[^a-zA-Z0-9]/g, "_")}_${new Date().toISOString().split("T")[0]}.csv`;

    // Download to vendor's device
    doDownloadCSV(cart, liveProducts, customerInfo, currency);

    // Upload receipt to Supabase Storage (iys-receipts bucket) — admin gets a copy
    uploadReceiptToStorage(csvContent, filename);

    // Save a receipt record so admin can retrieve/email it
    saveReceiptRecord(filename, customerInfo, totalEGP, currency);

    await saveOrder(cart, liveProducts, customerInfo, currency, session.token);
    setShopPage("confirmed");
  };

  // ═══════════════ AUTH SCREEN ═══════════════
  if (!session) return (
    <div style={{ minHeight: "100vh", background: "#faf8f5", display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "'DM Sans', sans-serif" }}>
      {fonts}
      <div style={{ width: "100%", maxWidth: 400, padding: "0 20px" }}>
        <div style={{ textAlign: "center", marginBottom: 32 }}>
          <div style={{ fontFamily: "'Playfair Display', serif", fontSize: 42, fontWeight: 900, color: "#1a1a1a", marginBottom: 4 }}>IYS</div>
          <div style={{ fontSize: 13, color: "#aaa" }}>Wholesale Vendor Portal</div>
        </div>
        <div style={{ background: "#fff", borderRadius: 20, border: "1px solid #e8e4df", padding: 28, boxShadow: "0 4px 24px rgba(0,0,0,0.06)" }}>
          <div style={{ display: "flex", marginBottom: 22, background: "#f5f1ec", borderRadius: 11, padding: 3 }}>
            {["login", "signup"].map(v => (
              <button key={v} onClick={() => { setAuthView(v); setAuthError(""); }}
                style={{ flex: 1, padding: "8px", border: "none", borderRadius: 9, background: authView === v ? "#fff" : "transparent", color: authView === v ? "#1a1a1a" : "#aaa", fontSize: 13, fontWeight: 700, cursor: "pointer", fontFamily: "'DM Sans', sans-serif", boxShadow: authView === v ? "0 1px 5px rgba(0,0,0,0.07)" : "none" }}>
                {v === "login" ? "Sign In" : "Sign Up"}
              </button>
            ))}
          </div>
          {authView === "signup" && <Field label="Full Name" value={authForm.name} onChange={e => setAuthForm(p => ({ ...p, name: e.target.value }))} placeholder="Your name" />}
          <Field label="Email" type="email" value={authForm.email} onChange={e => setAuthForm(p => ({ ...p, email: e.target.value }))} placeholder="you@example.com" />
          <Field label="Password" type="password" value={authForm.password} onChange={e => setAuthForm(p => ({ ...p, password: e.target.value }))} placeholder="••••••••" />
          {authError && <div style={{ background: authError.includes("✓") ? "#e8f5e9" : "#fef3ee", border: `1px solid ${authError.includes("✓") ? "#a5d6a7" : "#f5c6a8"}`, borderRadius: 9, padding: "9px 13px", color: authError.includes("✓") ? "#2e7d32" : "#e25822", fontSize: 13, fontWeight: 600, marginBottom: 14 }}>{authError}</div>}
          <button onClick={authView === "login" ? handleLogin : handleSignup} disabled={authLoading}
            style={{ width: "100%", padding: 13, background: authLoading ? "#ccc" : "#1a1a1a", color: "#fff", border: "none", borderRadius: 11, fontSize: 15, fontWeight: 800, cursor: authLoading ? "default" : "pointer", fontFamily: "'DM Sans', sans-serif" }}>
            {authLoading ? "Please wait…" : authView === "login" ? "Sign In" : "Create Account"}
          </button>
          <div style={{ marginTop: 16, fontSize: 12, color: "#bbb", textAlign: "center" }}>
            {authView === "login" ? "New? " : "Have an account? "}
            <span onClick={() => { setAuthView(authView === "login" ? "signup" : "login"); setAuthError(""); }} style={{ color: "#e25822", fontWeight: 700, cursor: "pointer" }}>
              {authView === "login" ? "Create account" : "Sign in"}
            </span>
          </div>
        </div>
        <div style={{ textAlign: "center", marginTop: 14, fontSize: 11, color: "#ccc" }}>Admin access is assigned by the IYS team.</div>
      </div>
    </div>
  );

  // ── SHARED NAV ──
  const NavBar = () => (
    <div style={{ background: "#fff", borderBottom: "1px solid #e8e4df", position: "sticky", top: 0, zIndex: 200 }}>
      <div style={{ maxWidth: 1280, margin: "0 auto", padding: "0 20px", display: "flex", alignItems: "center", justifyContent: "space-between", height: 52 }}>
        <div style={{ fontFamily: "'Playfair Display', serif", fontSize: 20, fontWeight: 900, color: "#1a1a1a" }}>IYS</div>
        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
          {!isAdmin && (
            <>
              <div style={{ display: "flex", gap: 2, background: "#f5f1ec", borderRadius: 9, padding: 3 }}>
                {Object.entries(CURRENCY_RATES).map(([code, info]) => (
                  <button key={code} onClick={() => setCurrency(code)}
                    style={{ padding: "4px 9px", border: "none", borderRadius: 7, background: currency === code ? "#1a1a1a" : "transparent", color: currency === code ? "#fff" : "#888", fontSize: 11, fontWeight: 700, cursor: "pointer", fontFamily: "'DM Sans', sans-serif" }}>
                    {info.flag} {code}
                  </button>
                ))}
              </div>
              <button onClick={() => totalItems > 0 && setShopPage("cart")}
                style={{ background: totalItems > 0 ? "#1a1a1a" : "#f5f1ec", color: totalItems > 0 ? "#fff" : "#aaa", border: "none", borderRadius: 9, padding: "6px 13px", cursor: totalItems > 0 ? "pointer" : "default", display: "flex", alignItems: "center", gap: 6, fontFamily: "'DM Sans', sans-serif", fontSize: 13, fontWeight: 700 }}>
                <Ic.Cart /> {totalItems}{totalItems > 0 && <span style={{ fontSize: 11, opacity: 0.7 }}>{dp(totalEGP)}</span>}
              </button>
            </>
          )}
          <div style={{ display: "flex", alignItems: "center", gap: 5, background: isAdmin ? "#1a1a1a" : "#f5f1ec", color: isAdmin ? "#fff" : "#888", borderRadius: 9, padding: "5px 11px", fontSize: 12, fontWeight: 700 }}>
            {isAdmin ? <Ic.Shield /> : <Ic.User />}
            <span style={{ maxWidth: 110, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{session.user.email}</span>
          </div>
          <button onClick={handleLogout} style={{ background: "none", border: "1px solid #e0dbd5", borderRadius: 9, padding: "5px 11px", cursor: "pointer", display: "flex", alignItems: "center", gap: 5, fontSize: 12, fontWeight: 600, color: "#888", fontFamily: "'DM Sans', sans-serif" }}>
            <Ic.LogOut /> Sign out
          </button>
        </div>
      </div>
    </div>
  );

  // ═══════════════ ADMIN VIEW ═══════════════
  if (isAdmin) return (
    <div style={{ minHeight: "100vh", background: "#faf8f5", fontFamily: "'DM Sans', sans-serif" }}>
      {fonts}
      <NavBar />
      <div style={{ maxWidth: 800, margin: "0 auto", padding: "28px 20px" }}>
        <div style={{ display: "flex", gap: 0, marginBottom: 24, background: "#f5f1ec", borderRadius: 12, padding: 3, width: "fit-content" }}>
          {[["data", "📊 Data"], ["orders", "🧾 Orders"], ["users", "👥 Users"], ["portal", "🛍 Portal"]].map(([k, l]) => (
            <button key={k} onClick={() => setAdminTab(k)}
              style={{ padding: "8px 18px", border: "none", borderRadius: 10, background: adminTab === k ? "#fff" : "transparent", color: adminTab === k ? "#1a1a1a" : "#999", fontSize: 13, fontWeight: 700, cursor: "pointer", fontFamily: "'DM Sans', sans-serif", boxShadow: adminTab === k ? "0 1px 5px rgba(0,0,0,0.07)" : "none" }}>
              {l}
            </button>
          ))}
        </div>

        {adminTab === "data" && (
          <>
            <h1 style={{ fontFamily: "'Playfair Display', serif", fontSize: 26, fontWeight: 900, color: "#1a1a1a", marginBottom: 4 }}>Data Manager</h1>
            <p style={{ color: "#888", fontSize: 13, marginBottom: 20 }}>Upload CSV or Excel files. Multiple files and sheets are supported. Save to publish to all vendors.</p>

            {/* Drop zone */}
            <div onDragOver={e => { e.preventDefault(); }} onDrop={e => { e.preventDefault(); [...e.dataTransfer.files].forEach(handleFile); }}
              onClick={() => fileRef.current?.click()}
              style={{ border: "2px dashed #d5d0c9", borderRadius: 16, padding: "36px 20px", cursor: "pointer", background: "#fff", textAlign: "center", marginBottom: 14 }}>
              <input ref={fileRef} type="file" accept=".csv,.xlsx,.xls" multiple style={{ display: "none" }} onChange={e => { [...e.target.files].forEach(handleFile); e.target.value = ""; }} />
              <div style={{ color: "#ccc", display: "flex", justifyContent: "center", marginBottom: 10 }}><Ic.Upload /></div>
              <p style={{ fontSize: 14, fontWeight: 700, color: "#1a1a1a", marginBottom: 3 }}>Drop files here or click to browse</p>
              <p style={{ fontSize: 12, color: "#aaa" }}>Supports .csv, .xlsx, .xls — multiple files &amp; sheets</p>
            </div>

            {parseError && <div style={{ background: "#fef3ee", border: "1px solid #f5c6a8", borderRadius: 9, padding: "9px 13px", color: "#e25822", fontSize: 13, fontWeight: 600, marginBottom: 12 }}>{parseError}</div>}

            {/* Sheet selector */}
            {uploadedSheets.length > 0 && (
              <div style={{ background: "#fff", border: "1px solid #e8e4df", borderRadius: 12, padding: "14px 16px", marginBottom: 12 }}>
                <div style={{ fontSize: 12, fontWeight: 700, color: "#555", marginBottom: 10 }}>📋 Sheets loaded — toggle which ones to include:</div>
                <div style={{ display: "flex", flexWrap: "wrap", gap: 7 }}>
                  {uploadedSheets.map(s => {
                    const on = activeSheets.includes(s.name);
                    return (
                      <button key={s.name} onClick={() => setActiveSheets(prev => on ? prev.filter(n => n !== s.name) : [...prev, s.name])}
                        style={{ padding: "5px 13px", border: `1px solid ${on ? "#1a1a1a" : "#ddd"}`, borderRadius: 8, background: on ? "#1a1a1a" : "#fff", color: on ? "#fff" : "#888", fontSize: 12, fontWeight: 600, cursor: "pointer", fontFamily: "'DM Sans', sans-serif" }}>
                        {on ? "✓ " : ""}{s.name} <span style={{ opacity: 0.5, fontSize: 10 }}>({s.rows.length})</span>
                      </button>
                    );
                  })}
                </div>
              </div>
            )}

            {/* Image col picker */}
            {csvHeaders.length > 0 && (
              <div style={{ background: "#fff", border: "1px solid #e8e4df", borderRadius: 12, padding: "14px 16px", marginBottom: 12 }}>
                <div style={{ fontSize: 12, fontWeight: 700, color: "#555", marginBottom: 7 }}>📸 Image URL column:</div>
                <select value={imageCol} onChange={e => setImageCol(e.target.value)}
                  style={{ width: "100%", padding: "8px 11px", border: "1px solid #e0dbd5", borderRadius: 8, fontSize: 13, fontFamily: "'DM Sans', sans-serif", background: "#faf8f5", outline: "none" }}>
                  <option value="">— None —</option>
                  {csvHeaders.map(h => <option key={h} value={h}>{h}</option>)}
                </select>
                {imageCol && <div style={{ fontSize: 11, color: "#27ae60", marginTop: 5, fontWeight: 600 }}>✓ Using: {imageCol}</div>}
              </div>
            )}

            {uploadStats && (
              <div style={{ background: "#e8f5e9", border: "1px solid #a5d6a7", borderRadius: 12, padding: "14px 18px", marginBottom: 14 }}>
                <div style={{ fontSize: 13, fontWeight: 700, color: "#2e7d32", marginBottom: 8 }}>✓ Ready to publish</div>
                <div style={{ display: "flex", gap: 20, fontSize: 13 }}>
                  <div><span style={{ fontSize: 19, fontWeight: 800, color: "#1a1a1a", display: "block" }}>{uploadStats.sheets}</span>sheets</div>
                  <div><span style={{ fontSize: 19, fontWeight: 800, color: "#1a1a1a", display: "block" }}>{uploadStats.rows}</span>rows</div>
                  <div><span style={{ fontSize: 19, fontWeight: 800, color: "#1a1a1a", display: "block" }}>{uploadStats.products}</span>products</div>
                </div>
              </div>
            )}

            {/* Current live catalog info */}
            {rawData && rawData.length > 0 && !uploadedSheets.length && (
              <div style={{ background: "#fff3e0", border: "1px solid #ffe0b2", borderRadius: 12, padding: "12px 16px", marginBottom: 14, fontSize: 13, color: "#e65100" }}>
                📦 <b>Live catalog:</b> {processProducts(rawData, imageCol).length} products currently published to vendors.
              </div>
            )}

            <div style={{ display: "flex", gap: 9 }}>
              {uploadedSheets.length > 0 && (
                <button onClick={saveCatalogToSupabase} disabled={saving}
                  style={{ flex: 1, padding: 12, background: saving ? "#ccc" : "#e25822", color: "#fff", border: "none", borderRadius: 10, fontSize: 14, fontWeight: 800, cursor: saving ? "default" : "pointer", fontFamily: "'DM Sans', sans-serif" }}>
                  {saving ? "Saving…" : "💾 Save & Publish to Vendors"}
                </button>
              )}
              {(rawData || uploadedSheets.length > 0) && (
                <button onClick={clearCatalog} style={{ padding: "12px 16px", background: "#fff", color: "#e25822", border: "1px solid #f5c6a8", borderRadius: 10, fontSize: 13, fontWeight: 700, cursor: "pointer", fontFamily: "'DM Sans', sans-serif", display: "flex", alignItems: "center", gap: 5 }}>
                  <Ic.Trash /> Clear All
                </button>
              )}
            </div>
          </>
        )}

        {adminTab === "orders" && (
          <>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 20 }}>
              <h1 style={{ fontFamily: "'Playfair Display', serif", fontSize: 26, fontWeight: 900, color: "#1a1a1a" }}>Orders</h1>
              <button onClick={() => { setOrdersLoading(true); sbFetch("/iys_orders?select=*&order=order_date.desc").then(r => { if (r.ok) setOrders(r.data); setOrdersLoading(false); }); }}
                style={{ background: "#f5f1ec", border: "none", borderRadius: 9, padding: "7px 13px", cursor: "pointer", display: "flex", alignItems: "center", gap: 5, fontSize: 12, fontWeight: 700, color: "#666", fontFamily: "'DM Sans', sans-serif" }}>
                <Ic.Refresh /> Refresh
              </button>
            </div>
            {ordersLoading
              ? <div style={{ textAlign: "center", padding: 40, color: "#aaa" }}>Loading orders…</div>
              : orders.length === 0
                ? <div style={{ textAlign: "center", padding: 40, color: "#aaa", background: "#fff", borderRadius: 14, border: "1px solid #e8e4df" }}>No orders yet.</div>
                : <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                  {orders.map(order => {
                    const cur = CURRENCY_RATES[order.currency] || CURRENCY_RATES.EGP;
                    const dp = egp => `${cur.symbol} ${(egp * cur.rate).toFixed(order.currency === "EGP" ? 0 : 2)}`;
                    const isExpanded = expandedOrder === order.id;
                    return (
                      <div key={order.id} style={{ background: "#fff", borderRadius: 13, border: "1px solid #e8e4df", overflow: "hidden" }}>
                        <div onClick={() => setExpandedOrder(isExpanded ? null : order.id)}
                          style={{ display: "flex", alignItems: "center", gap: 14, padding: "14px 18px", cursor: "pointer" }}>
                          <div style={{ flex: 1 }}>
                            <div style={{ fontSize: 14, fontWeight: 700, color: "#1a1a1a" }}>{order.customer_name}</div>
                            <div style={{ fontSize: 12, color: "#aaa", marginTop: 2 }}>{order.customer_contact} · {new Date(order.order_date).toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit" })}</div>
                          </div>
                          <div style={{ textAlign: "right" }}>
                            <div style={{ fontSize: 15, fontWeight: 800, color: "#1a1a1a" }}>{dp(order.total_egp)}</div>
                            {order.currency !== "EGP" && <div style={{ fontSize: 11, color: "#aaa" }}>EGP {order.total_egp?.toLocaleString()}</div>}
                          </div>
                          <div style={{ fontSize: 11, fontWeight: 700, padding: "3px 10px", borderRadius: 7, background: order.status === "pending" ? "#fff3e0" : "#e8f5e9", color: order.status === "pending" ? "#e65100" : "#2e7d32", marginLeft: 4 }}>
                            {order.status || "pending"}
                          </div>
                          <div style={{ fontSize: 14, color: "#ccc", marginLeft: 2 }}>{isExpanded ? "▲" : "▼"}</div>
                        </div>
                        {isExpanded && (
                          <div style={{ borderTop: "1px solid #f0ece7", padding: "14px 18px", background: "#faf8f5" }}>
                            <div style={{ fontSize: 11, fontWeight: 700, color: "#999", textTransform: "uppercase", letterSpacing: "0.5px", marginBottom: 10 }}>Order Items</div>
                            <div style={{ display: "grid", gridTemplateColumns: "2fr 1fr 1fr 1fr 1fr", gap: "6px 12px", fontSize: 11 }}>
                              <div style={{ fontWeight: 700, color: "#aaa" }}>Product</div>
                              <div style={{ fontWeight: 700, color: "#aaa" }}>Size</div>
                              <div style={{ fontWeight: 700, color: "#aaa" }}>SKU</div>
                              <div style={{ fontWeight: 700, color: "#aaa" }}>Qty</div>
                              <div style={{ fontWeight: 700, color: "#aaa", textAlign: "right" }}>Total</div>
                              {(order.items || []).map((item, i) => (
                                <>
                                  <div key={`t${i}`} style={{ fontWeight: 600, color: "#1a1a1a", fontSize: 12 }}>{item.title}</div>
                                  <div key={`s${i}`} style={{ color: "#666" }}>{item.size}</div>
                                  <div key={`k${i}`} style={{ color: "#aaa" }}>{item.sku || "—"}</div>
                                  <div key={`q${i}`} style={{ fontWeight: 700 }}>{item.qty}</div>
                                  <div key={`l${i}`} style={{ fontWeight: 700, textAlign: "right" }}>{dp(item.lineTotalEGP)}</div>
                                </>
                              ))}
                            </div>
                            <div style={{ marginTop: 12, display: "flex", justifyContent: "flex-end", gap: 8 }}>
                              <button onClick={async () => {
                                await sbFetch(`/iys_orders?id=eq.${order.id}`, { method: "PATCH", headers: { Authorization: `Bearer ${session.token}` }, body: JSON.stringify({ status: "fulfilled" }) });
                                setOrders(prev => prev.map(o => o.id === order.id ? { ...o, status: "fulfilled" } : o));
                              }} style={{ fontSize: 12, fontWeight: 700, padding: "5px 14px", borderRadius: 8, background: "#e8f5e9", color: "#2e7d32", border: "1px solid #a5d6a7", cursor: "pointer", fontFamily: "'DM Sans', sans-serif" }}>
                                ✓ Mark Fulfilled
                              </button>
                            </div>
                          </div>
                        )}
                      </div>
                    );
                  })}
                </div>}
          </>
        )}

        {adminTab === "users" && (
          <>
            <h1 style={{ fontFamily: "'Playfair Display', serif", fontSize: 26, fontWeight: 900, color: "#1a1a1a", marginBottom: 20 }}>User Management</h1>
            <div style={{ background: "#fff", borderRadius: 14, border: "1px solid #e8e4df", overflow: "hidden" }}>
              {users.length === 0 && <div style={{ padding: 32, textAlign: "center", color: "#aaa" }}>No users yet.</div>}
              {users.map((u, i) => (
                <div key={u.id} style={{ display: "flex", alignItems: "center", gap: 12, padding: "13px 18px", borderBottom: i < users.length - 1 ? "1px solid #f0ece7" : "none" }}>
                  <div style={{ width: 34, height: 34, borderRadius: "50%", background: u.role === "admin" ? "#1a1a1a" : "#f5f1ec", display: "flex", alignItems: "center", justifyContent: "center" }}>
                    {u.role === "admin" ? <Ic.Shield /> : <Ic.User />}
                  </div>
                  <div style={{ flex: 1 }}>
                    <div style={{ fontSize: 13, fontWeight: 700 }}>{u.name || "—"}</div>
                    <div style={{ fontSize: 11, color: "#aaa" }}>{u.email}</div>
                  </div>
                  <span style={{ fontSize: 11, fontWeight: 700, padding: "3px 9px", borderRadius: 7, background: u.role === "admin" ? "#1a1a1a" : "#f5f1ec", color: u.role === "admin" ? "#fff" : "#888" }}>{u.role === "admin" ? "Admin" : "Vendor"}</span>
                  {u.role !== "admin" && <button onClick={() => promoteUser(u.id)} style={{ fontSize: 11, fontWeight: 700, padding: "4px 11px", borderRadius: 7, background: "#fef3ee", color: "#e25822", border: "1px solid #f5c6a8", cursor: "pointer", fontFamily: "'DM Sans', sans-serif", display: "flex", alignItems: "center", gap: 3 }}><Ic.Plus /> Admin</button>}
                </div>
              ))}
            </div>
          </>
        )}
        {adminTab === "portal" && (
          <>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 20 }}>
              <div>
                <h1 style={{ fontFamily: "'Playfair Display', serif", fontSize: 26, fontWeight: 900, color: "#1a1a1a", marginBottom: 4 }}>Vendor Portal Preview</h1>
                <p style={{ color: "#888", fontSize: 13 }}>This is exactly what vendors see. You can browse and place test orders.</p>
              </div>
            </div>
            <div style={{ background: "#fff", borderRadius: 14, border: "1px solid #e8e4df", padding: 20, marginBottom: 14 }}>
              <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 20 }}>
                <input type="text" placeholder="Search…" value={search} onChange={e => setSearch(e.target.value)}
                  style={{ padding: "8px 13px", border: "1px solid #e0dbd5", borderRadius: 9, fontSize: 13, background: "#faf8f5", width: 200, outline: "none", fontFamily: "'DM Sans', sans-serif" }} />
                {productTypes.map(t => (
                  <button key={t} onClick={() => setFilterType(t)}
                    style={{ padding: "6px 13px", border: "1px solid", borderColor: filterType === t ? "#1a1a1a" : "#e0dbd5", borderRadius: 8, background: filterType === t ? "#1a1a1a" : "#fff", color: filterType === t ? "#fff" : "#666", fontSize: 12, fontWeight: 600, cursor: "pointer", fontFamily: "'DM Sans', sans-serif" }}>
                    {t}
                  </button>
                ))}
              </div>
              {catalogLoading
                ? <div style={{ textAlign: "center", padding: "60px 0", color: "#aaa", fontSize: 14 }}>Loading catalog…</div>
                : liveProducts.length === 0
                  ? <div style={{ textAlign: "center", padding: "60px 0" }}><div style={{ fontSize: 44, marginBottom: 12 }}>⏳</div><div style={{ fontSize: 15, fontWeight: 600, color: "#aaa" }}>No products published yet.</div><div style={{ fontSize: 12, color: "#ccc", marginTop: 4 }}>Upload and publish a CSV from the Data tab.</div></div>
                  : <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill,minmax(240px,1fr))", gap: 16 }}>
                    {filtered.map(p => <ProductCard key={p.title} product={p} cart={cart} setCart={setCart} currency={currency} />)}
                  </div>}
            </div>
            {totalItems > 0 && (
              <div style={{ position: "sticky", bottom: 18, display: "flex", justifyContent: "center" }}>
                <div onClick={() => { setAdminTab("data"); setShopPage("cart"); }} style={{ background: "#1a1a1a", color: "#fff", borderRadius: 14, padding: "12px 24px", display: "inline-flex", alignItems: "center", gap: 16, boxShadow: "0 8px 36px rgba(0,0,0,0.22)", cursor: "pointer", fontFamily: "'DM Sans', sans-serif" }}>
                  <Ic.Cart /><span style={{ fontWeight: 700 }}>{totalItems} item{totalItems > 1 ? "s" : ""}</span>
                  <div style={{ width: 1, height: 16, background: "rgba(255,255,255,0.2)" }} />
                  <span style={{ fontWeight: 800, fontSize: 15 }}>{dp(totalEGP)}</span>
                </div>
              </div>
            )}
          </>
        )}

      </div>
    </div>
  );

  // ═══════════════ VENDOR SHOP ═══════════════
  if (shopPage === "shop") return (
    <div style={{ minHeight: "100vh", background: "#faf8f5", fontFamily: "'DM Sans', sans-serif" }}>
      {fonts}
      <NavBar />
      <div style={{ maxWidth: 1280, margin: "0 auto", padding: "28px 20px 16px" }}>
        {catalogLoading
          ? <div style={{ textAlign: "center", padding: "80px 0", color: "#aaa", fontSize: 14 }}>Loading catalog…</div>
          : liveProducts.length === 0
            ? <div style={{ textAlign: "center", padding: "80px 0" }}><div style={{ fontSize: 44, marginBottom: 12 }}>⏳</div><div style={{ fontSize: 15, fontWeight: 600, color: "#aaa" }}>No products available yet.</div><div style={{ fontSize: 12, color: "#ccc", marginTop: 4 }}>The IYS team will upload the catalog shortly.</div></div>
            : <>
              <h1 style={{ fontFamily: "'Playfair Display', serif", fontSize: "clamp(24px,4vw,36px)", fontWeight: 900, color: "#1a1a1a", marginBottom: 6 }}>Bulk Order</h1>
              <p style={{ color: "#888", fontSize: 13, marginBottom: 18 }}>Select quantities and place your wholesale order.</p>
              <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 20 }}>
                <input type="text" placeholder="Search…" value={search} onChange={e => setSearch(e.target.value)}
                  style={{ padding: "8px 13px", border: "1px solid #e0dbd5", borderRadius: 9, fontSize: 13, background: "#fff", width: 200, outline: "none", fontFamily: "'DM Sans', sans-serif" }} />
                {productTypes.map(t => (
                  <button key={t} onClick={() => setFilterType(t)}
                    style={{ padding: "6px 13px", border: "1px solid", borderColor: filterType === t ? "#1a1a1a" : "#e0dbd5", borderRadius: 8, background: filterType === t ? "#1a1a1a" : "#fff", color: filterType === t ? "#fff" : "#666", fontSize: 12, fontWeight: 600, cursor: "pointer", fontFamily: "'DM Sans', sans-serif" }}>
                    {t}
                  </button>
                ))}
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill,minmax(260px,1fr))", gap: 16 }}>
                {filtered.map(p => <ProductCard key={p.title} product={p} cart={cart} setCart={setCart} currency={currency} />)}
              </div>
            </>}
      </div>
      {totalItems > 0 && (
        <div onClick={() => setShopPage("cart")} style={{ position: "fixed", bottom: 18, left: "50%", transform: "translateX(-50%)", background: "#1a1a1a", color: "#fff", borderRadius: 14, padding: "12px 24px", display: "flex", alignItems: "center", gap: 16, boxShadow: "0 8px 36px rgba(0,0,0,0.22)", zIndex: 100, cursor: "pointer", fontFamily: "'DM Sans', sans-serif", whiteSpace: "nowrap" }}>
          <Ic.Cart /><span style={{ fontWeight: 700 }}>{totalItems} item{totalItems > 1 ? "s" : ""}</span>
          <div style={{ width: 1, height: 16, background: "rgba(255,255,255,0.2)" }} />
          <span style={{ fontWeight: 800, fontSize: 15 }}>{dp(totalEGP)}</span>
          <span style={{ fontSize: 12, opacity: 0.5 }}>→ Review</span>
        </div>
      )}
    </div>
  );

  // ═══════════════ CART ═══════════════
  if (shopPage === "cart") return (
    <div style={{ minHeight: "100vh", background: "#faf8f5", fontFamily: "'DM Sans', sans-serif" }}>
      {fonts}
      <NavBar />
      <div style={{ maxWidth: 800, margin: "0 auto", padding: "24px 20px 80px" }}>
        <button onClick={() => setShopPage("shop")} style={{ background: "none", border: "none", color: "#888", fontSize: 13, cursor: "pointer", marginBottom: 16, fontFamily: "'DM Sans', sans-serif", fontWeight: 600 }}>← Back</button>
        <h1 style={{ fontFamily: "'Playfair Display', serif", fontSize: 26, fontWeight: 900, color: "#1a1a1a", marginBottom: 22 }}>Review Order</h1>
        <div style={{ display: "flex", gap: 22, flexWrap: "wrap" }}>
          <div style={{ flex: "1 1 380px" }}>
            <div style={{ background: "#fff", borderRadius: 13, border: "1px solid #e8e4df", overflow: "hidden" }}>
              {cartItems.map((item, i) => (
                <div key={item.key} style={{ display: "flex", alignItems: "center", gap: 13, padding: "13px 16px", borderBottom: i < cartItems.length - 1 ? "1px solid #f0ece7" : "none" }}>
                  <div style={{ width: 48, height: 48, borderRadius: 8, background: "#f5f1ec", overflow: "hidden", flexShrink: 0 }}>
                    {item.prod?.image && <img src={item.prod.image} alt="" style={{ width: "100%", height: "100%", objectFit: "cover" }} />}
                  </div>
                  <div style={{ flex: 1 }}>
                    <div style={{ fontSize: 13, fontWeight: 700 }}>{item.pTitle}</div>
                    <div style={{ fontSize: 11, color: "#aaa" }}>{item.size} · qty {item.qty}</div>
                  </div>
                  <div style={{ textAlign: "right" }}>
                    <div style={{ fontSize: 13, fontWeight: 800 }}>{dp((item.prod?.bulkPrice || 0) * item.qty)}</div>
                    <div style={{ fontSize: 11, color: "#bbb" }}>{dp(item.prod?.bulkPrice || 0)} ea</div>
                  </div>
                  <button onClick={() => setCart(p => { const n = { ...p }; delete n[item.key]; return n; })} style={{ background: "none", border: "none", cursor: "pointer", color: "#ccc" }}><Ic.X /></button>
                </div>
              ))}
            </div>
            <div style={{ marginTop: 10, padding: "16px 20px", background: "#1a1a1a", borderRadius: 13, color: "#fff", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
              <div>
                <div style={{ fontSize: 11, opacity: 0.5, marginBottom: 2 }}>TOTAL ({currency})</div>
                <div style={{ fontSize: 24, fontWeight: 900, fontFamily: "'Playfair Display', serif" }}>{dp(totalEGP)}</div>
                {currency !== "EGP" && <div style={{ fontSize: 11, opacity: 0.4 }}>EGP {totalEGP.toLocaleString()}</div>}
              </div>
              <div style={{ fontSize: 12, opacity: 0.4 }}>{totalItems} items</div>
            </div>
          </div>
          <div style={{ flex: "1 1 240px" }}>
            <div style={{ background: "#fff", borderRadius: 13, border: "1px solid #e8e4df", padding: 20 }}>
              <h3 style={{ fontFamily: "'Playfair Display', serif", fontSize: 17, fontWeight: 800, marginBottom: 16 }}>Your Details</h3>
              <Field label="Full Name" value={customerInfo.name} onChange={e => { setCustomerInfo(p => ({ ...p, name: e.target.value })); setCustErrors(p => ({ ...p, name: undefined })); }} placeholder="Your name" error={custErrors.name} />
              <Field label="Email or Phone" value={customerInfo.contact} onChange={e => { setCustomerInfo(p => ({ ...p, contact: e.target.value })); setCustErrors(p => ({ ...p, contact: undefined })); }} placeholder="email or +966…" error={custErrors.contact} />
              <div style={{ background: "#f5f1ec", borderRadius: 8, padding: "9px 11px", marginBottom: 14, fontSize: 12, color: "#888", lineHeight: 1.5 }}>Your order will be emailed to the IYS team and a CSV sheet will download.</div>
              <button onClick={handleConfirmOrder} style={{ width: "100%", padding: 12, background: "#e25822", color: "#fff", border: "none", borderRadius: 10, fontSize: 14, fontWeight: 800, cursor: "pointer", fontFamily: "'DM Sans', sans-serif", display: "flex", alignItems: "center", justifyContent: "center", gap: 6 }}>
                <Ic.Check /> Confirm &amp; Send Order
              </button>
            </div>
          </div>
        </div>
      </div>
    </div>
  );

  // ═══════════════ CONFIRMED ═══════════════
  return (
    <div style={{ minHeight: "100vh", background: "#faf8f5", fontFamily: "'DM Sans', sans-serif", display: "flex", alignItems: "center", justifyContent: "center" }}>
      {fonts}
      <div style={{ maxWidth: 500, padding: "0 20px", textAlign: "center" }}>
        <div style={{ width: 64, height: 64, borderRadius: "50%", background: "#e8f5e9", display: "flex", alignItems: "center", justifyContent: "center", margin: "0 auto 18px" }}>
          <svg width="30" height="30" viewBox="0 0 24 24" fill="none" stroke="#2e7d32" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><polyline points="20 6 9 17 4 12"/></svg>
        </div>
        <h1 style={{ fontFamily: "'Playfair Display', serif", fontSize: 30, fontWeight: 900, color: "#1a1a1a", marginBottom: 8 }}>Order Sent!</h1>
        <p style={{ color: "#888", fontSize: 14, marginBottom: 24, lineHeight: 1.6 }}>Your order has been saved and the IYS team will be in touch. A CSV sheet has been downloaded to your device.</p>
        <div style={{ background: "#fff", borderRadius: 13, border: "1px solid #e8e4df", padding: 20, marginBottom: 18, textAlign: "left" }}>
          <div style={{ display: "flex", justifyContent: "space-between", flexWrap: "wrap", gap: 10, marginBottom: 14 }}>
            {[["Customer", customerInfo.name], ["Contact", customerInfo.contact], ["Currency", currency]].map(([l, v]) => (
              <div key={l}><div style={{ fontSize: 10, color: "#aaa", fontWeight: 700, marginBottom: 1 }}>{l}</div><div style={{ fontSize: 13, fontWeight: 700 }}>{v}</div></div>
            ))}
          </div>
          <div style={{ padding: "13px 16px", background: "#1a1a1a", borderRadius: 10, color: "#fff", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            <span style={{ fontSize: 12, opacity: 0.5 }}>{totalItems} items</span>
            <span style={{ fontSize: 19, fontWeight: 900, fontFamily: "'Playfair Display', serif" }}>{dp(totalEGP)}</span>
          </div>
        </div>
        <div style={{ display: "flex", gap: 8, justifyContent: "center", flexWrap: "wrap" }}>
          <button onClick={() => doDownloadCSV(cart, liveProducts, customerInfo, currency)} style={{ padding: "10px 20px", background: "#1a1a1a", color: "#fff", border: "none", borderRadius: 10, fontSize: 13, fontWeight: 700, cursor: "pointer", fontFamily: "'DM Sans', sans-serif", display: "flex", alignItems: "center", gap: 6 }}><Ic.Download /> Re-download</button>
          <button onClick={() => { setShopPage("shop"); setCart({}); setCustomerInfo({ name: "", contact: "" }); }} style={{ padding: "10px 20px", background: "#fff", color: "#666", border: "1px solid #ddd", borderRadius: 10, fontSize: 13, fontWeight: 700, cursor: "pointer", fontFamily: "'DM Sans', sans-serif" }}>New Order</button>
        </div>
        <div style={{ marginTop: 16, padding: "10px 16px", background: "#e8f5e9", borderRadius: 9, fontSize: 12, color: "#2e7d32", fontWeight: 600 }}>✓ Order saved — the IYS team will review it shortly.</div>
      </div>
    </div>
  );
}
