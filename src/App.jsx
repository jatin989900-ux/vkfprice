import { useState, useEffect, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import { initializeApp } from "firebase/app";
import { getFirestore, doc, getDoc, setDoc, onSnapshot } from "firebase/firestore";

const firebaseConfig = {
  apiKey: "AIzaSyCEJlH9aGlr0pneb7hT1sIxy1iDnQV3Y4g",
  authDomain: "vkf-price.firebaseapp.com",
  projectId: "vkf-price",
  storageBucket: "vkf-price.firebasestorage.app",
  messagingSenderId: "199170851796",
  appId: "1:199170851796:web:e2e46e279995a41b890c41"
};
const app = initializeApp(firebaseConfig);
const db  = getFirestore(app);
const FS  = {
  settings:  () => doc(db, "vkf", "settings"),
  brands:    () => doc(db, "vkf", "brands"),
  items:     () => doc(db, "vkf", "items"),
  estimates: () => doc(db, "vkf", "estimates"),
};
async function fsLoad() {
  try {
    const [sS, bS, iS, eS] = await Promise.all([getDoc(FS.settings()), getDoc(FS.brands()), getDoc(FS.items()), getDoc(FS.estimates())]);
    return { settings: sS.exists() ? sS.data().v : null, brands: bS.exists() ? bS.data().v : null, items: iS.exists() ? iS.data().v : null, estimates: eS.exists() ? eS.data().v : [] };
  } catch (e) { console.error(e); return { settings: null, brands: null, items: null, estimates: [] }; }
}
async function fsSave(key, value) {
  await setDoc(FS[key](), { v: value, updatedAt: new Date().toISOString() });
}

// ── CONSTANTS ─────────────────────────────────────────────────────
const VER  = "4.5";
const CATS = ["Bed Sheets","Comforters","Comforter Sets","Towels","Pillows","Dohars","Blankets","Top Sheets","Other Items"];
const GST_OPTS = [{ label:"5% — Default (Textiles)", v:0.05 },{ label:"12%", v:0.12 },{ label:"18%", v:0.18 }];
const ADDON_OPTS = [
  { label:"Not set", v:"" },{ label:"+ ₹100", v:100 },{ label:"+ ₹200", v:200 },
  { label:"+ ₹300", v:300 },{ label:"+ ₹500", v:500 },{ label:"+ ₹1,000", v:1000 },{ label:"Custom ₹", v:"custom" },
];

// ── CALC ENGINE ───────────────────────────────────────────────────
function psychRound(raw) {
  if (!raw || isNaN(raw)) return null;
  const n = parseFloat(raw);
  if (n <= 0) return null;
  const r5 = Math.round(n / 5) * 5;
  if (Math.floor(r5 / 100) > Math.floor(n / 100)) return Math.floor(n / 100) * 100 - 5;
  return r5;
}
function calcIncGST(ex, g) { return +(parseFloat(ex||0) * (1 + parseFloat(g||0.05))).toFixed(2); }
function calcRL(ex, m) { const v = parseFloat(ex||0); return v > 0 ? psychRound(v * (1 + parseFloat(m||0.18))) : null; }
function calcDM(ex, m) { const v = parseFloat(ex||0); if (m == null || m === "" || isNaN(parseFloat(m))) return null; return v > 0 ? psychRound(v * (1 + parseFloat(m))) : null; }
function calcPL(rl, a) { return (rl != null && typeof a === "number" && !isNaN(a)) ? psychRound(rl + a) : null; }
function calcAddon(pla, ca) {
  if (pla === "" || pla == null) return null;
  if (pla === "custom") { const v = parseFloat(ca); return !isNaN(v) && v > 0 ? v : null; }
  return typeof pla === "number" ? pla : null;
}
function resolveRL(item, brand, settings) {
  const v = parseFloat(item.customRL);
  if (item.customRL !== "" && item.customRL != null && !isNaN(v)) return v / 100;
  if (brand && brand.rlMarkup != null) return brand.rlMarkup;
  return parseFloat(settings.defaultRL || 0.18);
}
function resolveDM(item, brand, settings) {
  const v = parseFloat(item.customDM);
  if (item.customDM !== "" && item.customDM != null && !isNaN(v)) return v / 100;
  if (brand && brand.dmMarkup != null) return brand.dmMarkup;
  if (settings.defaultDM != null && settings.defaultDM !== "") return parseFloat(settings.defaultDM);
  return null;
}

// ── COMPUTE ───────────────────────────────────────────────────────
function computeItem(item, brand, settings) {
  const rlM = resolveRL(item, brand, settings);
  const dmM = resolveDM(item, brand, settings);
  const add = calcAddon(item.plAddon, item.customAddon);
  const ex  = parseFloat(item.purchaseEx || 0);
  const gst = parseFloat(item.gst || 0.05);

  const rlCustom = (item.customRLPrice !== "" && item.customRLPrice != null && !isNaN(parseFloat(item.customRLPrice))) ? parseFloat(item.customRLPrice) : null;
  const dmCustom = (item.customDMPrice !== "" && item.customDMPrice != null && !isNaN(parseFloat(item.customDMPrice))) ? parseFloat(item.customDMPrice) : null;

  const rl = rlCustom !== null ? rlCustom : calcRL(ex, rlM);
  const dm = dmCustom !== null ? dmCustom : calcDM(ex, dmM);
  const pl = calcPL(rl, add);

  const sdmAddOnRaw = item.sdmAddon;
  const sdmPctRaw  = item.sdmPct;
  // sdmPct takes priority over sdmAddon if set
  const sdmPctVal  = (sdmPctRaw!==""&&sdmPctRaw!=null&&!isNaN(parseFloat(sdmPctRaw))&&parseFloat(sdmPctRaw)>0) ? parseFloat(sdmPctRaw)/100 : null;
  const sdmAddOn   = sdmPctVal!=null && ex>0
    ? +(ex * sdmPctVal).toFixed(2)
    : (sdmAddOnRaw!==""&&sdmAddOnRaw!=null&&!isNaN(parseFloat(sdmAddOnRaw))&&parseFloat(sdmAddOnRaw)>=0)?parseFloat(sdmAddOnRaw):null;
  const sdm    = (sdmAddOn !== null && ex > 0) ? +(ex + sdmAddOn).toFixed(2) : null;
  const sdmInc = sdm != null ? +(sdm * (1 + gst)).toFixed(2) : null;

  function profitIncGST(price) {
    if (price == null || !ex || ex <= 0) return null;
    const exRev = price / (1 + gst);          // selling price ex-GST
    const amt   = +(exRev - ex).toFixed(2);
    return { amt, pct: +((amt / exRev) * 100).toFixed(1) }; // divide by selling price ex-GST
  }
  function profitExGST(price) {
    if (price == null || !ex || ex <= 0) return null;
    const amt = +(price - ex).toFixed(2);
    return { amt, pct: +((amt / price) * 100).toFixed(1) }; // divide by selling price
  }

  return {
    incGST: calcIncGST(ex, gst),
    rlM, rl, rlProfit: profitIncGST(rl),
    dmM, dm, dmProfit: profitIncGST(dm),
    add, pl, plProfit: profitIncGST(pl),
    sdmAddOn, sdm, sdmInc, sdmProfit: profitExGST(sdm), sdmPctVal,
    rlPriceLocked: rlCustom !== null,
    dmPriceLocked: dmCustom !== null,
    gst,
  };
}

// ── DEFAULTS ──────────────────────────────────────────────────────
// sdmPIN added
const DEF_S = { co:"VK Furnishing", tag:"Wholesale Bedding & Textiles, Delhi NCR", defaultRL:0.18, defaultDM:null, adminPIN:"1234", salesPIN:"0000", sdmPIN:"9999", categories:CATS.slice(), salesmen:["Vikas","Amar Jatav","Pankaj","Harender"], estimatePrefix:"VKF", estimateCounter:1 };
const DEF_B = [
  { id:"b1", code:"TRI", name:"Trident",      rlMarkup:0.22, dmMarkup:null },
  { id:"b2", code:"STH", name:"Story@Home",   rlMarkup:0.18, dmMarkup:null },
  { id:"b3", code:"SWT", name:"Swayam",       rlMarkup:0.20, dmMarkup:null },
  { id:"b4", code:"LOC", name:"Local / Misc", rlMarkup:0.18, dmMarkup:null },
];
const T0 = new Date().toISOString();
const DEF_I = [
  { id:"i1", cat:"Bed Sheets", bId:"b1", name:"King Cotton Bedsheet 200TC",    purchaseEx:380, gst:0.05, customRL:"", customDM:"", customRLPrice:"", customDMPrice:"", sdmAddon:"", plAddon:300,  customAddon:"", purchaseDate:"", purchaseHistory:[], active:true, notes:"", createdAt:T0, updatedAt:T0 },
  { id:"i2", cat:"Comforters", bId:"b2", name:"Winter Hollow Fibre Comforter", purchaseEx:550, gst:0.05, customRL:"", customDM:"", customRLPrice:"", customDMPrice:"", sdmAddon:"", plAddon:500,  customAddon:"", purchaseDate:"", purchaseHistory:[], active:true, notes:"", createdAt:T0, updatedAt:T0 },
  { id:"i3", cat:"Towels",     bId:"b3", name:"Premium Bath Towel 500 GSM",    purchaseEx:180, gst:0.18, customRL:"", customDM:"", customRLPrice:"", customDMPrice:"", sdmAddon:"", plAddon:200,  customAddon:"", purchaseDate:"", purchaseHistory:[], active:true, notes:"", createdAt:T0, updatedAt:T0 },
];

// ── HELPERS ───────────────────────────────────────────────────────
const uid     = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 5);
const fp      = (n) => (n != null && !isNaN(n)) ? "₹" + Number(n).toLocaleString("en-IN") : "—";
const fpct    = (n) => (n != null && n !== "" && !isNaN(parseFloat(n))) ? (parseFloat(n) * 100).toFixed(0) + "%" : "—";
const fpctRaw = (s) => (s !== "" && s != null && !isNaN(parseFloat(s))) ? parseFloat(s).toFixed(1) + "%" : null;
function fmtDate(iso) { if (!iso) return ""; return new Date(iso).toLocaleDateString("en-IN", { day:"2-digit", month:"short", year:"numeric" }); }
function fmtDateTime(iso) { if (!iso) return ""; const d = new Date(iso); return fmtDate(iso) + " " + d.toLocaleTimeString("en-IN", { hour:"2-digit", minute:"2-digit", hour12:true }); }
function dateGroupKey(iso) {
  if (!iso) return "Unknown";
  const d = new Date(iso), today = new Date(), yest = new Date(); yest.setDate(today.getDate() - 1);
  if (d.toDateString() === today.toDateString()) return "Today";
  if (d.toDateString() === yest.toDateString())  return "Yesterday";
  return d.toLocaleDateString("en-IN", { weekday:"long", day:"2-digit", month:"short", year:"numeric" });
}
function sameDay(a, b) { return new Date(a).toDateString() === new Date(b).toDateString(); }

// ── PALETTE ───────────────────────────────────────────────────────
const C = {
  bg:"#F0F2F7", card:"#fff", border:"#E5E8EF", text:"#111827", sec:"#6B7280", mute:"#9CA3AF",
  blue:"#1648D6", navy:"#0B1D5C",
  rl:"#067D62",  rlBg:"#ECFDF5",  rlBr:"#6EE7B7",
  dm:"#1C64F2",  dmBg:"#EFF6FF",  dmBr:"#93C5FD",
  pl:"#6D28D9",  plBg:"#F5F3FF",  plBr:"#C4B5FD",
  sdm:"#C2410C", sdmBg:"#FFF7ED", sdmBr:"#FDBA74",
  profit:"#065F46", profBg:"#F0FDF4",
  amb:"#B45309",  ambBg:"#FEF3C7",
  red:"#DC2626",  redBg:"#FEF2F2",
  new_:"#0E7490", newBg:"#ECFEFF", newBr:"#A5F3FC",
  upd:"#7C3AED",  updBg:"#F5F3FF", updBr:"#DDD6FE",
  lock:"#7C3AED", lockBg:"#F5F3FF",lockBr:"#C4B5FD",
};

// ── TOAST ─────────────────────────────────────────────────────────
let _toast = null;
const toast = (m, t = "ok") => _toast && _toast({ id: uid(), m, t });
function ToastHost() {
  const [list, setList] = useState([]);
  useEffect(() => {
    _toast = (x) => { setList(p => [...p, x]); setTimeout(() => setList(p => p.filter(y => y.id !== x.id)), 2600); };
  }, []);
  const BG = { ok:"#059669", err:C.red, warn:"#D97706" };
  return (
    <div style={{ position:"fixed", top:14, right:14, zIndex:9999, display:"flex", flexDirection:"column", gap:8 }}>
      {list.map(x => <div key={x.id} style={{ padding:"11px 16px", borderRadius:10, fontSize:13, fontWeight:700, color:"#fff", background:BG[x.t]||BG.ok, boxShadow:"0 4px 16px rgba(0,0,0,0.2)" }}>{x.m}</div>)}
    </div>
  );
}

function SyncBadge({ status }) {
  const cfg = {
    synced:  { col:"#059669", bg:"#ECFDF5", br:"#6EE7B7", icon:"☁️", label:"Synced" },
    saving:  { col:"#D97706", bg:"#FEF3C7", br:"#FCD34D", icon:"⏳", label:"Saving…" },
    error:   { col:C.red,     bg:C.redBg,   br:"#FCA5A5", icon:"⚠️", label:"Sync error" },
    offline: { col:C.sec,     bg:"#F3F4F6", br:C.border,  icon:"📵", label:"Offline" },
  }[status] || { col:C.sec, bg:"#F3F4F6", br:C.border, icon:"…", label:status };
  return <span style={{ display:"inline-flex", alignItems:"center", gap:4, padding:"3px 9px", borderRadius:20, fontSize:11, fontWeight:700, color:cfg.col, background:cfg.bg, border:"1px solid "+cfg.br }}>{cfg.icon} {cfg.label}</span>;
}

// ── UI ATOMS ──────────────────────────────────────────────────────
function Chip({ col, bg, br, children }) {
  return <span style={{ display:"inline-block", padding:"2px 9px", borderRadius:20, fontSize:11, fontWeight:700, color:col, background:bg, border:"1px solid "+br }}>{children}</span>;
}
function PBox({ label, val, col, bg, br, locked, large, locked2, onClick }) {
  return (
    <div onClick={onClick} style={{ flex:1, background:bg, border:"1px solid "+(locked?col:br), borderRadius:10, padding:large?"14px 10px":"9px 6px", textAlign:"center", position:"relative", cursor:onClick?"pointer":"default" }}>
      {locked  && <div style={{ position:"absolute", top:3, right:5, fontSize:9, color:col, fontWeight:800 }}>📌</div>}
      {locked2 && <div style={{ position:"absolute", top:3, right:5, fontSize:10 }}>🔒</div>}
      <div style={{ fontSize:9, fontWeight:700, color:C.mute, letterSpacing:"0.5px", textTransform:"uppercase", marginBottom:3 }}>{label}</div>
      <div style={{ fontSize:large?18:15, fontWeight:800, color:col }}>{val}</div>
    </div>
  );
}
function Fld({ label, hint, err, children }) {
  return (
    <div style={{ marginBottom:14 }}>
      <label style={{ display:"block", fontSize:11, fontWeight:700, color:C.sec, letterSpacing:"0.7px", textTransform:"uppercase", marginBottom:5 }}>{label}</label>
      {children}
      {err  && <div style={{ fontSize:11, color:C.red,  marginTop:3 }}>{err}</div>}
      {!err && hint && <div style={{ fontSize:11, color:C.mute, marginTop:3 }}>{hint}</div>}
    </div>
  );
}
const INP = { width:"100%", padding:"11px 13px", borderRadius:9, fontSize:14, border:"1.5px solid #E5E8EF", background:"#FAFBFD", fontFamily:"inherit", outline:"none", boxSizing:"border-box", color:"#111827" };
const SEL = Object.assign({}, INP, { cursor:"pointer" });
function BtnP({ color, onClick, children, style }) {
  return <button onClick={onClick} style={Object.assign({ padding:"13px 20px", borderRadius:10, border:"none", background:color||C.blue, color:"#fff", fontWeight:700, fontSize:14, cursor:"pointer", fontFamily:"inherit", width:"100%" }, style||{})}>{children}</button>;
}
function BtnO({ color, onClick, children, style }) {
  const col = color || C.sec;
  return <button onClick={onClick} style={Object.assign({ padding:"13px 20px", borderRadius:10, border:"1.5px solid "+col, background:"#fff", color:col, fontWeight:600, fontSize:14, cursor:"pointer", fontFamily:"inherit", width:"100%" }, style||{})}>{children}</button>;
}

function MarginOverride({ label, col, bg, br, value, onChange, srcLabel, err }) {
  const has = value !== "" && value != null && !isNaN(parseFloat(value));
  return (
    <div style={{ marginBottom:14 }}>
      <label style={{ display:"block", fontSize:11, fontWeight:700, color:C.sec, letterSpacing:"0.7px", textTransform:"uppercase", marginBottom:5 }}>{label}</label>
      <div style={{ position:"relative" }}>
        <input style={Object.assign({}, INP, { borderColor:err?C.red:has?col:C.border, paddingRight:has?"90px":"13px" })}
          type="number" step="0.5" min="0" max="100"
          placeholder={"Leave blank = " + srcLabel} value={value}
          onChange={e => onChange(e.target.value)} />
        {has && <div style={{ position:"absolute", right:10, top:"50%", transform:"translateY(-50%)", background:bg, border:"1px solid "+br, color:col, fontSize:11, fontWeight:700, padding:"2px 8px", borderRadius:6 }}>{parseFloat(value).toFixed(1) + "%"}</div>}
      </div>
      {err  && <div style={{ fontSize:11, color:C.red, marginTop:3 }}>{err}</div>}
      {!err && <div style={{ fontSize:11, color:has?col:C.mute, fontWeight:has?600:400, marginTop:3 }}>{has ? "⚡ Margin override active" : "Using: " + srcLabel}</div>}
    </div>
  );
}

function PriceOverride({ label, col, bg, br, value, onChange, err, hint }) {
  const has = value !== "" && value != null && !isNaN(parseFloat(value)) && parseFloat(value) > 0;
  return (
    <div style={{ marginBottom:14, background:has?bg:"transparent", border:has?"1.5px solid "+col:"1.5px solid "+C.border, borderRadius:10, padding:has?"10px 12px":"0", transition:"all 0.15s" }}>
      {has && <div style={{ fontSize:9, fontWeight:800, color:col, letterSpacing:"0.7px", textTransform:"uppercase", marginBottom:6 }}>📌 DIRECT PRICE LOCKED — margin % ignored</div>}
      <label style={{ display:"block", fontSize:11, fontWeight:700, color:has?col:C.sec, letterSpacing:"0.7px", textTransform:"uppercase", marginBottom:5 }}>{label}</label>
      <div style={{ position:"relative" }}>
        <input style={Object.assign({}, INP, { borderColor:err?C.red:has?col:C.border, fontWeight:has?800:400, fontSize:has?16:14, color:has?col:C.text, background:"#fff" })}
          type="number" step="1" min="0" placeholder="Leave blank to use margin %"
          value={value} onChange={e => onChange(e.target.value)} />
        {has && <div style={{ position:"absolute", right:10, top:"50%", transform:"translateY(-50%)", background:col, color:"#fff", fontSize:12, fontWeight:800, padding:"3px 10px", borderRadius:7 }}>{fp(parseFloat(value))}</div>}
      </div>
      {err  && <div style={{ fontSize:11, color:C.red, marginTop:3 }}>{err}</div>}
      {!err && <div style={{ fontSize:11, color:has?col:C.mute, fontWeight:has?700:400, marginTop:4 }}>{has ? "Margin % field above is now ignored for this item" : (hint || "Enter a price to skip margin calculation")}</div>}
    </div>
  );
}

function ProfitRow({ label, col, price, profit, locked, exGST }) {
  if (!price) return null;
  return (
    <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", padding:"5px 10px", background:C.profBg, borderRadius:7, marginBottom:4, border:"1px solid #BBF7D0" }}>
      <div style={{ display:"flex", alignItems:"center", gap:8 }}>
        <span style={{ fontSize:11, fontWeight:700, color:col }}>{label}</span>
        {locked && <span style={{ fontSize:9, color:C.lock, fontWeight:800, background:C.lockBg, padding:"1px 5px", borderRadius:4 }}>📌 FIXED</span>}
        {exGST  && <span style={{ fontSize:9, color:C.sdm,  fontWeight:800, background:C.sdmBg,  padding:"1px 5px", borderRadius:4 }}>Ex-GST</span>}
        <span style={{ fontSize:13, fontWeight:800, color:col }}>{fp(price)}</span>
      </div>
      {profit ? (
        <div style={{ display:"flex", alignItems:"center", gap:6 }}>
          <span style={{ fontSize:11, fontWeight:700, color:C.profit }}>{"+"+fp(profit.amt)}</span>
          <span style={{ background:"#D1FAE5", color:C.profit, fontSize:11, fontWeight:800, padding:"2px 8px", borderRadius:20 }}>{profit.pct + "%"}</span>
        </div>
      ) : <span style={{ fontSize:11, color:C.mute }}>—</span>}
    </div>
  );
}

// ── SDM PIN MODAL ─────────────────────────────────────────────────
function SDMPinModal({ pin, onSuccess, onClose }) {
  return (
    <div style={{ position:"fixed", inset:0, zIndex:700, background:"rgba(0,0,0,0.55)", display:"flex", alignItems:"center", justifyContent:"center", padding:20 }}>
      <div style={{ background:"#fff", borderRadius:22, padding:"28px 24px", width:"100%", maxWidth:320, boxShadow:"0 20px 60px rgba(0,0,0,0.3)" }}>
        <div style={{ textAlign:"center", marginBottom:16 }}>
          <div style={{ fontSize:28, marginBottom:8 }}>🟠</div>
          <div style={{ fontSize:16, fontWeight:800, color:C.text }}>SDM Prices</div>
          <div style={{ fontSize:12, color:C.sec, marginTop:4 }}>Enter SDM PIN to view special pricing</div>
        </div>
        <PINPad
          label="Enter SDM PIN"
          pin={pin}
          onSuccess={() => { onSuccess(); }}
        />
        <div style={{ marginTop:16 }}>
          <BtnO color={C.sec} onClick={onClose}>Cancel</BtnO>
        </div>
      </div>
    </div>
  );
}

// ── ITEM CARD ─────────────────────────────────────────────────────
function ItemCard({ it, isAdmin, showDate, priceFilter, sdmUnlocked, onSDMClick, onAddToEstimate }) {
  const filter = priceFilter || "all";
  const single = filter !== "all" && filter !== "highlights";
  const isNew    = it.createdAt && it.updatedAt && sameDay(it.createdAt, it.updatedAt);
  const isEdited = it.createdAt && it.updatedAt && !sameDay(it.createdAt, it.updatedAt);
  const isHighlighted = !!it.highlighted;

  const canSeeSdm = isAdmin || sdmUnlocked;

  return (
    <div style={{ background:C.card, border:"2px solid "+(isHighlighted?"#F59E0B":C.border), borderRadius:12, padding:"13px", marginBottom:8, boxShadow:isHighlighted?"0 2px 12px rgba(245,158,11,0.18)":"0 1px 5px rgba(0,0,0,0.04)" }}>
      {/* Highlight banner */}
      {isHighlighted && (
        <div style={{ background:"#FEF3C7", border:"1px solid #FCD34D", borderRadius:8, padding:"6px 11px", marginBottom:9, display:"flex", alignItems:"center", justifyContent:"space-between" }}>
          <span style={{ fontSize:12, fontWeight:800, color:"#92400E" }}>⭐ Highlighted Product</span>
          {it.commission && <span style={{ fontSize:12, fontWeight:800, color:"#B45309", background:"#fff", border:"1.5px solid #FCD34D", borderRadius:20, padding:"2px 10px" }}>+{fp(parseFloat(it.commission))}/piece</span>}
        </div>
      )}
      {isHighlighted && it.highlightNote && (
        <div style={{ fontSize:11, color:"#92400E", fontStyle:"italic", marginBottom:8, fontWeight:600 }}>📌 {it.highlightNote}</div>
      )}
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", marginBottom:9 }}>
        <div style={{ flex:1, paddingRight:8 }}>
          <div style={{ fontWeight:700, fontSize:14, marginBottom:4 }}>{it.name}</div>
          <div style={{ display:"flex", gap:6, flexWrap:"wrap", alignItems:"center" }}>
            {it._b && <div style={{ background:C.navy, color:"#fff", borderRadius:5, padding:"1px 8px", fontSize:10, fontWeight:800 }}>{it._b.code}</div>}
            <Chip col={C.sec} bg="#F3F4F6" br={C.border}>{it.cat}</Chip>
            {it.pinned && <Chip col={C.blue} bg="#EFF6FF" br="#BFDBFE">📌 Pinned</Chip>}
            {isNew    && <Chip col={C.new_} bg={C.newBg} br={C.newBr}>🆕 New</Chip>}
            {isEdited && <Chip col={C.upd}  bg={C.updBg} br={C.updBr}>✏️ Updated</Chip>}
          </div>
          {showDate && (
            <div style={{ marginTop:5, fontSize:10, color:C.mute }}>
              {isNew ? "Added " + fmtDateTime(it.createdAt) : "Updated " + fmtDateTime(it.updatedAt)}
              {isEdited && <span style={{ marginLeft:6, color:C.updBr }}>· Added {fmtDate(it.createdAt)}</span>}
            </div>
          )}
        </div>
        <Chip col={C.amb} bg={C.ambBg} br="#FCD34D">{"GST " + (it.gst * 100).toFixed(0) + "%"}</Chip>
      </div>

      {isAdmin && (
        <div style={{ background:C.ambBg, border:"1px solid #FCD34D", borderRadius:8, padding:"7px 11px", marginBottom:9 }}>
          <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center" }}>
            <span style={{ fontSize:11, fontWeight:700, color:C.amb, textTransform:"uppercase" }}>Purchase Ex-GST</span>
            <span style={{ fontWeight:800, fontSize:13, color:C.amb }}>{fp(it.purchaseEx) + " + "}<span style={{ fontSize:10, fontWeight:600 }}>{(it.gst * 100).toFixed(0) + "% = " + fp(it.incGST)}</span></span>
          </div>
          {it.purchaseDate && <div style={{ fontSize:10, color:C.amb, marginTop:3, fontWeight:600 }}>📅 Purchased: {fmtDate(it.purchaseDate)}</div>}
        </div>
      )}

      {single ? (
        <div>
          {filter === "rl"  && <PBox label="RL — Wholesale (Inc-GST)"     val={fp(it.rl)}  col={C.rl}  bg={C.rlBg}  br={C.rlBr}  locked={it.rlPriceLocked} large />}
          {filter === "dm"  && <PBox label="DM — Sp. Wholesale (Inc-GST)" val={fp(it.dm)}  col={C.dm}  bg={C.dmBg}  br={C.dmBr}  locked={it.dmPriceLocked} large />}
          {filter === "pl"  && <PBox label="PL — Retail (Inc-GST)"        val={fp(it.pl)}  col={C.pl}  bg={C.plBg}  br={C.plBr}  large />}
          {filter === "sdm" && (
            <PBox
              label="SDM — Ex-GST + total"
              val={canSeeSdm && it.sdm != null ? fp(it.sdm) + " + GST = " + fp(it.sdmInc) : "🔒 Tap to unlock"}
              col={C.sdm} bg={C.sdmBg} br={C.sdmBr}
              locked2={!canSeeSdm}
              large
              onClick={!canSeeSdm ? onSDMClick : undefined}
            />
          )}
        </div>
      ) : (
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:7 }}>
          <PBox label="RL — Wholesale"  val={fp(it.rl)}  col={C.rl}  bg={C.rlBg}  br={C.rlBr}  locked={it.rlPriceLocked} />
          <PBox label="DM — Sp. WHL"    val={fp(it.dm)}  col={C.dm}  bg={C.dmBg}  br={C.dmBr}  locked={it.dmPriceLocked} />
          <PBox label="PL — Retail"     val={fp(it.pl)}  col={C.pl}  bg={C.plBg}  br={C.plBr} />
          <PBox
            label="SDM (tap to unlock)"
            val={canSeeSdm && it.sdm != null ? fp(it.sdm) + " + GST = " + fp(it.sdmInc) : "🔒"}
            col={C.sdm} bg={canSeeSdm ? C.sdmBg : "#F3F4F6"} br={canSeeSdm ? C.sdmBr : C.border}
            locked2={!canSeeSdm}
            onClick={!canSeeSdm ? onSDMClick : undefined}
          />
        </div>
      )}

      {isAdmin && (
        <div style={{ marginTop:9 }}>
          <div style={{ fontSize:10, fontWeight:700, color:C.profit, textTransform:"uppercase", letterSpacing:"0.5px", marginBottom:5 }}>💰 Profit / piece — RL/DM/PL: GST backed out · SDM: Ex-GST direct</div>
          <ProfitRow label="RL (Inc-GST)" col={C.rl}  price={it.rl}  profit={it.rlProfit}  locked={it.rlPriceLocked} />
          <ProfitRow label="DM (Inc-GST)" col={C.dm}  price={it.dm}  profit={it.dmProfit}  locked={it.dmPriceLocked} />
          <ProfitRow label="PL (Inc-GST)" col={C.pl}  price={it.pl}  profit={it.plProfit} />
          <ProfitRow label="SDM (Ex-GST)" col={C.sdm} price={it.sdm} profit={it.sdmProfit} exGST />
        </div>
      )}
      {it.notes && <div style={{ marginTop:7, fontSize:11, color:C.mute, fontStyle:"italic" }}>{"📝 " + it.notes}</div>}
      {onAddToEstimate && (
        <button onClick={() => onAddToEstimate(it)} style={{ marginTop:9, width:"100%", padding:"9px", borderRadius:8, border:"1.5px solid "+C.blue, background:"#EFF6FF", color:C.blue, fontWeight:700, fontSize:13, cursor:"pointer", fontFamily:"inherit" }}>
          + Add to Estimate
        </button>
      )}
    </div>
  );
}

// ── DRAWER ────────────────────────────────────────────────────────
function Drawer({ open, onClose, title, footer, children }) {
  if (!open) return null;
  return (
    <div>
      <div onClick={onClose} style={{ position:"fixed", inset:0, background:"rgba(0,0,0,0.45)", zIndex:500 }} />
      <div style={{ position:"fixed", bottom:0, left:0, right:0, zIndex:501, background:"#fff", borderRadius:"22px 22px 0 0", maxHeight:"92vh", display:"flex", flexDirection:"column", boxShadow:"0 -8px 40px rgba(0,0,0,0.18)" }}>
        <div style={{ display:"flex", justifyContent:"center", padding:"10px 0 0" }}><div style={{ width:36, height:4, borderRadius:2, background:C.border }} /></div>
        <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", padding:"8px 20px 12px" }}>
          <div style={{ fontSize:18, fontWeight:800, color:C.text }}>{title}</div>
          <button onClick={onClose} style={{ width:34, height:34, borderRadius:8, border:"1.5px solid "+C.border, background:"#F9FAFB", cursor:"pointer", fontSize:16, color:C.sec, fontFamily:"inherit" }}>✕</button>
        </div>
        <div style={{ flex:1, overflowY:"auto", padding:"0 20px 6px" }}>{children}</div>
        {footer && <div style={{ padding:"12px 20px", borderTop:"1px solid "+C.border }}>{footer}</div>}
      </div>
    </div>
  );
}

function Confirm({ title, msg, onOk, onNo }) {
  return (
    <div style={{ position:"fixed", inset:0, zIndex:600, background:"rgba(0,0,0,0.5)", display:"flex", alignItems:"center", justifyContent:"center", padding:20 }}>
      <div style={{ background:"#fff", borderRadius:18, padding:"26px 22px", maxWidth:320, width:"100%" }}>
        <div style={{ fontSize:17, fontWeight:800, color:C.text, marginBottom:8 }}>{title}</div>
        <div style={{ fontSize:13, color:C.sec, lineHeight:1.7, marginBottom:24 }}>{msg}</div>
        <div style={{ display:"flex", gap:10 }}><BtnO onClick={onNo}>Cancel</BtnO><BtnP color={C.red} onClick={onOk}>Delete</BtnP></div>
      </div>
    </div>
  );
}

function PINPad({ label, pin, onSuccess }) {
  const [val, setVal] = useState(""); const [err, setErr] = useState(""); const [shk, setShk] = useState(false);
  function press(d) {
    if (val.length >= 4) return;
    const next = val + d; setVal(next);
    if (next.length === 4) {
      if (next === String(pin)) { onSuccess(); }
      else { setShk(true); setErr("Wrong PIN — try again."); setTimeout(() => { setVal(""); setShk(false); setErr(""); }, 700); }
    }
  }
  const keys = ["1","2","3","4","5","6","7","8","9","","0","⌫"];
  return (
    <div style={{ textAlign:"center" }}>
      <div style={{ fontSize:13, fontWeight:700, color:C.sec, marginBottom:16 }}>{label}</div>
      <div style={{ display:"flex", justifyContent:"center", gap:14, marginBottom:6 }}>
        {[0,1,2,3].map(i => <div key={i} style={{ width:15, height:15, borderRadius:"50%", background:val.length>i?C.blue:"transparent", border:"2.5px solid "+(val.length>i?C.blue:C.border), transition:"all 0.12s" }} />)}
      </div>
      <div style={{ minHeight:22, fontSize:12, color:C.red, fontWeight:600, marginBottom:8 }}>{err}</div>
      <div style={{ display:"grid", gridTemplateColumns:"repeat(3,1fr)", gap:10, maxWidth:220, margin:"0 auto" }}>
        {keys.map((k, i) => (
          <button key={i} disabled={k === ""} onClick={() => k === "⌫" ? setVal(p => p.slice(0,-1)) : k ? press(k) : null}
            style={{ height:56, borderRadius:12, fontSize:k==="⌫"?20:22, fontWeight:700, cursor:k===""?"default":"pointer", fontFamily:"inherit", background:k===""?"transparent":"#fff", border:k===""?"none":"1.5px solid "+C.border, color:k==="⌫"?C.red:C.text, boxShadow:k!==""?"0 2px 6px rgba(0,0,0,0.07)":"none" }}>
            {k}
          </button>
        ))}
      </div>
    </div>
  );
}

// ── LOGIN ─────────────────────────────────────────────────────────
function Login({ settings, onLogin }) {
  const [role, setRole] = useState(null);
  const roles = [
    { r:"admin",    icon:"🔐", t:"Admin",    d:"Full access — prices, brands, settings" },
    { r:"salesman", icon:"📋", t:"Salesman", d:"View price list only" },
  ];
  return (
    <div style={{ minHeight:"100vh", background:"linear-gradient(160deg,"+C.navy+" 0%,"+C.blue+" 100%)", display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center", padding:20 }}>
      <div style={{ marginBottom:30, textAlign:"center" }}>
        <div style={{ width:66, height:66, borderRadius:18, background:"rgba(255,255,255,0.13)", border:"2px solid rgba(255,255,255,0.2)", display:"flex", alignItems:"center", justifyContent:"center", fontSize:24, fontWeight:900, color:"#fff", margin:"0 auto 12px" }}>VK</div>
        <div style={{ fontSize:20, fontWeight:800, color:"#fff" }}>{settings.co}</div>
        <div style={{ fontSize:12, color:"rgba(255,255,255,0.4)", marginTop:3 }}>{settings.tag}</div>
      </div>
      <div style={{ background:"#fff", borderRadius:22, padding:"28px 24px", width:"100%", maxWidth:340, boxShadow:"0 30px 80px rgba(0,0,0,0.3)" }}>
        {!role ? (
          <div>
            <div style={{ fontSize:16, fontWeight:800, color:C.text, textAlign:"center", marginBottom:18 }}>Select your role</div>
            {roles.map(o => (
              <button key={o.r} onClick={() => setRole(o.r)} style={{ display:"flex", alignItems:"center", gap:14, padding:"14px 16px", borderRadius:12, border:"1.5px solid "+C.border, background:"#FAFBFD", cursor:"pointer", width:"100%", fontFamily:"inherit", marginBottom:10, textAlign:"left" }}>
                <span style={{ fontSize:24 }}>{o.icon}</span>
                <div style={{ flex:1 }}><div style={{ fontWeight:700, fontSize:14, color:C.text }}>{o.t}</div><div style={{ fontSize:12, color:C.sec, marginTop:1 }}>{o.d}</div></div>
                <span style={{ color:C.mute, fontSize:18 }}>›</span>
              </button>
            ))}
          </div>
        ) : (
          <div>
            <button onClick={() => setRole(null)} style={{ background:"none", border:"none", cursor:"pointer", color:C.sec, fontSize:13, fontWeight:600, marginBottom:16, padding:0, fontFamily:"inherit" }}>← Back</button>
            <PINPad label={"Enter " + (role === "admin" ? "Admin" : "Salesman") + " PIN"} pin={role === "admin" ? settings.adminPIN : settings.salesPIN} onSuccess={() => onLogin(role)} />
          </div>
        )}
      </div>
      <div style={{ marginTop:16, fontSize:11, color:"rgba(255,255,255,0.25)" }}>{"v" + VER}</div>
    </div>
  );
}

// ── BRANDS VIEW ───────────────────────────────────────────────────
function BrandsView({ brands, onBrandsChange }) {
  const EF = { code:"", name:"", rlMarkup:"18", dmMarkup:"" };
  const [form, setForm] = useState(EF); const [eid, setEid] = useState(null);
  const [open, setOpen] = useState(false); const [conf, setConf] = useState(null);
  function close() { setOpen(false); setEid(null); setForm(EF); }
  function startEdit(b) { setForm({ code:b.code, name:b.name, rlMarkup:+(b.rlMarkup*100).toFixed(1), dmMarkup:b.dmMarkup!=null?+(b.dmMarkup*100).toFixed(1):"" }); setEid(b.id); setOpen(true); }
  function save() {
    if (!form.code.trim()) return toast("Brand code required","err");
    if (!form.name.trim()) return toast("Brand name required","err");
    if (isNaN(parseFloat(form.rlMarkup))) return toast("RL markup must be a number","err");
    const entry = { id:eid||uid(), code:form.code.trim().toUpperCase(), name:form.name.trim(), rlMarkup:parseFloat(form.rlMarkup)/100, dmMarkup:form.dmMarkup!==""&&!isNaN(parseFloat(form.dmMarkup))?parseFloat(form.dmMarkup)/100:null };
    onBrandsChange(eid ? brands.map(b => b.id===eid?entry:b) : [...brands, entry]);
    toast(eid ? "Brand updated" : "Brand added"); close();
  }
  function del(id) { onBrandsChange(brands.filter(b => b.id !== id)); setConf(null); toast("Brand removed","warn"); }
  return (
    <div style={{ padding:"16px 16px 80px" }}>
      {conf && <Confirm title="Remove Brand?" msg="Items using this brand will keep prices but lose markup override." onOk={() => del(conf)} onNo={() => setConf(null)} />}
      <div style={{ background:"#EFF6FF", border:"1px solid #BFDBFE", borderRadius:10, padding:"10px 14px", marginBottom:14, fontSize:12, color:"#1E40AF", lineHeight:1.6 }}><strong>Brand markups are defaults.</strong> Override per item in Master Sheet.</div>
      <BtnP onClick={() => { setForm(EF); setEid(null); setOpen(true); }}>+ Add Brand</BtnP>
      <div style={{ height:12 }} />
      {brands.length === 0 && <div style={{ textAlign:"center", padding:"36px", color:C.mute, fontSize:14 }}>No brands yet.</div>}
      {brands.map(b => (
        <div key={b.id} style={{ background:C.card, border:"1px solid "+C.border, borderRadius:13, padding:"14px", marginBottom:10, boxShadow:"0 1px 6px rgba(0,0,0,0.05)" }}>
          <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:8 }}>
            <div style={{ display:"flex", alignItems:"center", gap:9 }}>
              <div style={{ background:C.navy, color:"#fff", borderRadius:7, padding:"4px 10px", fontWeight:800, fontSize:12 }}>{b.code}</div>
              <div style={{ fontWeight:700, fontSize:14 }}>{b.name}</div>
            </div>
            <div style={{ display:"flex", gap:7 }}>
              <button onClick={() => startEdit(b)} style={{ padding:"5px 13px", borderRadius:7, border:"1.5px solid "+C.border, background:"#F9FAFB", cursor:"pointer", fontSize:12, fontWeight:600, color:C.sec, fontFamily:"inherit" }}>Edit</button>
              <button onClick={() => setConf(b.id)} style={{ padding:"5px 10px", borderRadius:7, border:"1.5px solid #FCA5A5", background:C.redBg, cursor:"pointer", fontSize:13, color:C.red, fontFamily:"inherit" }}>✕</button>
            </div>
          </div>
          <div style={{ display:"flex", gap:7 }}>
            <Chip col={C.rl} bg={C.rlBg} br={C.rlBr}>{"RL " + fpct(b.rlMarkup)}</Chip>
            {b.dmMarkup != null ? <Chip col={C.dm} bg={C.dmBg} br={C.dmBr}>{"DM " + fpct(b.dmMarkup)}</Chip> : <Chip col={C.mute} bg="#F3F4F6" br={C.border}>DM not set</Chip>}
          </div>
        </div>
      ))}
      <Drawer open={open} onClose={close} title={eid ? "Edit Brand" : "Add Brand"}
        footer={<div style={{ display:"flex", gap:9 }}><BtnO onClick={close}>Cancel</BtnO><BtnP color={eid?"#F59E0B":C.blue} onClick={save}>{eid ? "Update" : "Add Brand"}</BtnP></div>}>
        <Fld label="Brand Code *" hint="3–6 letters"><input style={INP} maxLength={6} value={form.code} placeholder="TRI" onChange={e => setForm(p => ({...p, code:e.target.value.toUpperCase()}))} /></Fld>
        <Fld label="Brand Name *"><input style={INP} value={form.name} placeholder="e.g. Trident" onChange={e => setForm(p => ({...p, name:e.target.value}))} /></Fld>
        <Fld label="Default RL Markup % *" hint="18 = 18%"><input style={INP} type="number" step="0.5" value={form.rlMarkup} placeholder="18" onChange={e => setForm(p => ({...p, rlMarkup:e.target.value}))} /></Fld>
        <Fld label="Default DM Markup %" hint="Leave blank to decide later."><input style={INP} type="number" step="0.5" value={form.dmMarkup} placeholder="Leave blank" onChange={e => setForm(p => ({...p, dmMarkup:e.target.value}))} /></Fld>
      </Drawer>
    </div>
  );
}

// ── EXPORT MODAL ──────────────────────────────────────────────────
function ExportModal({ items, brands, settings, onClose }) {
  const [fCat, setFCat] = useState("All"); const [fBrand, setFBrand] = useState("All"); const [fSt, setFSt] = useState("Active");
  const filtered = useMemo(() => {
    return items.filter(i => {
      if (fCat !== "All" && i.cat !== fCat) return false;
      if (fBrand !== "All" && i.bId !== fBrand) return false;
      if (fSt === "Active" && !i.active) return false;
      if (fSt === "Inactive" && i.active) return false;
      return true;
    }).map(i => { const b = brands.find(x => x.id===i.bId)||null; return Object.assign({}, i, computeItem(i,b,settings), {_b:b}); });
  }, [items, brands, settings, fCat, fBrand, fSt]);
  const today = new Date().toLocaleDateString("en-IN", { day:"2-digit", month:"short", year:"numeric" });
  function doXLSX(rows, fname) {
    if (!rows.length) return toast("No items","warn");
    try { const ws = XLSX.utils.json_to_sheet(rows); const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, "Price List"); XLSX.writeFile(wb, fname+"_"+today+".xlsx"); toast("Excel downloaded!"); }
    catch { toast("Export failed","err"); }
  }
  function doJSON(rows, fname) {
    if (!rows.length) return toast("No items","warn");
    const blob = new Blob([JSON.stringify({ exportedAt:new Date().toISOString(), data:rows }, null, 2)], { type:"application/json" });
    const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = fname+"_"+today+".json"; a.click(); toast("JSON downloaded!");
  }
  function adminRows() {
    return filtered.map(it => ({
      "Category":it.cat,"Brand Code":it._b?it._b.code:"","Brand Name":it._b?it._b.name:"","Item Name":it.name,
      "Purchase Ex-GST":it.purchaseEx,"GST %":(it.gst*100).toFixed(0)+"%","Purchase Inc-GST":it.incGST,
      "RL Markup":it.rlPriceLocked?"(locked)":fpct(it.rlM),"RL Price (Inc-GST)":it.rl||"","RL Profit Rs":it.rlProfit?it.rlProfit.amt:"","RL Profit %":it.rlProfit?it.rlProfit.pct:"",
      "DM Markup":it.dmPriceLocked?"(locked)":it.dmM!=null?fpct(it.dmM):"","DM Price (Inc-GST)":it.dm||"","DM Profit Rs":it.dmProfit?it.dmProfit.amt:"","DM Profit %":it.dmProfit?it.dmProfit.pct:"",
      "PL Addon":it.add||"","PL Price (Inc-GST)":it.pl||"","PL Profit Rs":it.plProfit?it.plProfit.amt:"","PL Profit %":it.plProfit?it.plProfit.pct:"",
      "SDM Addon":it.sdmAddOn||"","SDM Price (Ex-GST)":it.sdm||"","SDM Inc-GST":it.sdmInc||"","SDM Profit Rs":it.sdmProfit?it.sdmProfit.amt:"","SDM Profit %":it.sdmProfit?it.sdmProfit.pct:"",
      "Status":it.active?"Active":"Inactive","Added On":fmtDate(it.createdAt),"Updated On":fmtDate(it.updatedAt),"Notes":it.notes,
    }));
  }
  function salesRows() {
    return filtered.map(it => ({
      "Category":it.cat,"Brand Code":it._b?it._b.code:"","Brand Name":it._b?it._b.name:"","Item Name":it.name,
      "GST %":(it.gst*100).toFixed(0)+"%","RL Price":it.rl||"","DM Price":it.dm||"","PL Price":it.pl||"",
      "SDM Price (Ex-GST)":it.sdm||"","SDM Inc-GST Total":it.sdmInc||"","Notes":it.notes,
    }));
  }
  return (
    <div style={{ position:"fixed", inset:0, zIndex:600, background:"rgba(0,0,0,0.5)", display:"flex", alignItems:"center", justifyContent:"center", padding:16 }}>
      <div style={{ background:"#fff", borderRadius:20, padding:"22px 18px", maxWidth:400, width:"100%", maxHeight:"88vh", overflowY:"auto" }}>
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:16 }}>
          <div style={{ fontSize:17, fontWeight:800 }}>⬇ Export Price List</div>
          <button onClick={onClose} style={{ width:32, height:32, borderRadius:8, border:"1.5px solid "+C.border, background:"#F9FAFB", cursor:"pointer", fontSize:15, color:C.sec, fontFamily:"inherit" }}>✕</button>
        </div>
        <div style={{ background:"#F8FAFF", border:"1px solid #BFDBFE", borderRadius:10, padding:"13px", marginBottom:14 }}>
          <div style={{ fontSize:11, fontWeight:700, color:C.blue, textTransform:"uppercase", letterSpacing:"0.6px", marginBottom:10 }}>Filter Items</div>
          <Fld label="Category"><select style={SEL} value={fCat} onChange={e => setFCat(e.target.value)}><option value="All">All Categories</option>{(settings.categories||CATS).map(c => <option key={c} value={c}>{c}</option>)}</select></Fld>
          <Fld label="Brand"><select style={SEL} value={fBrand} onChange={e => setFBrand(e.target.value)}><option value="All">All Brands</option>{brands.map(b => <option key={b.id} value={b.id}>{b.code+" — "+b.name}</option>)}</select></Fld>
          <Fld label="Status"><div style={{ display:"flex", border:"1px solid "+C.border, borderRadius:9, overflow:"hidden" }}>{["All","Active","Inactive"].map(s => <button key={s} onClick={() => setFSt(s)} style={{ flex:1, padding:"9px", border:"none", cursor:"pointer", fontWeight:600, fontSize:12, fontFamily:"inherit", background:fSt===s?C.blue:"#fff", color:fSt===s?"#fff":C.sec }}>{s}</button>)}</div></Fld>
          <div style={{ fontSize:12, color:C.mute, fontWeight:600 }}>{filtered.length + " items selected"}</div>
        </div>
        <div style={{ background:C.ambBg, border:"1px solid #FCD34D", borderRadius:10, padding:"13px", marginBottom:10 }}>
          <div style={{ fontSize:12, fontWeight:800, color:C.amb, marginBottom:3 }}>🔐 Admin Export</div>
          <div style={{ fontSize:11, color:C.amb, marginBottom:10 }}>Includes purchase + profit. Do NOT share.</div>
          <div style={{ display:"flex", gap:8 }}><BtnP color={C.amb} onClick={() => doXLSX(adminRows(),"VKF_Admin")} style={{ fontSize:12, padding:"10px" }}>📊 Excel</BtnP><BtnO color={C.amb} onClick={() => doJSON(adminRows(),"VKF_Admin")} style={{ fontSize:12, padding:"10px" }}>JSON</BtnO></div>
        </div>
        <div style={{ background:C.rlBg, border:"1px solid "+C.rlBr, borderRadius:10, padding:"13px" }}>
          <div style={{ fontSize:12, fontWeight:800, color:C.rl, marginBottom:3 }}>📋 Salesman Export</div>
          <div style={{ fontSize:11, color:C.rl, marginBottom:10 }}>Prices only. Safe to share.</div>
          <div style={{ display:"flex", gap:8 }}><BtnP color={C.rl} onClick={() => doXLSX(salesRows(),"VKF_PriceList")} style={{ fontSize:12, padding:"10px" }}>📊 Excel</BtnP><BtnO color={C.rl} onClick={() => doJSON(salesRows(),"VKF_PriceList")} style={{ fontSize:12, padding:"10px" }}>JSON</BtnO></div>
        </div>
      </div>
    </div>
  );
}

// ── PURCHASE HISTORY MODAL ────────────────────────────────────────
function PurchaseHistoryModal({ item, onClose }) {
  const history = (item.purchaseHistory || []).slice().reverse(); // newest first
  return (
    <div style={{ position:"fixed", inset:0, zIndex:600, background:"rgba(0,0,0,0.5)", display:"flex", alignItems:"center", justifyContent:"center", padding:16 }}>
      <div style={{ background:"#fff", borderRadius:20, padding:"22px 18px", maxWidth:380, width:"100%", maxHeight:"80vh", display:"flex", flexDirection:"column" }}>
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:4 }}>
          <div style={{ fontSize:17, fontWeight:800 }}>📅 Purchase History</div>
          <button onClick={onClose} style={{ width:32, height:32, borderRadius:8, border:"1.5px solid "+C.border, background:"#F9FAFB", cursor:"pointer", fontSize:15, color:C.sec, fontFamily:"inherit" }}>✕</button>
        </div>
        <div style={{ fontSize:12, color:C.sec, marginBottom:16, fontWeight:600 }}>{item.name}</div>
        {history.length === 0 ? (
          <div style={{ textAlign:"center", padding:"36px 0", color:C.mute, fontSize:14 }}>
            No history yet. History builds from the next price update.
          </div>
        ) : (
          <div style={{ overflowY:"auto", flex:1 }}>
            {history.map((h, i) => (
              <div key={i} style={{ display:"flex", justifyContent:"space-between", alignItems:"center", padding:"11px 14px", background:i===0?C.ambBg:"#F9FAFB", border:"1px solid "+(i===0?"#FCD34D":C.border), borderRadius:10, marginBottom:8 }}>
                <div>
                  <div style={{ fontSize:13, fontWeight:800, color:i===0?C.amb:C.text }}>{fp(h.price)}</div>
                  <div style={{ fontSize:10, color:C.mute, marginTop:2 }}>Recorded: {fmtDate(h.recordedAt)}</div>
                </div>
                <div style={{ textAlign:"right" }}>
                  <div style={{ fontSize:12, fontWeight:700, color:i===0?C.amb:C.sec }}>{fmtDate(h.date)}</div>
                  {i === 0 && <div style={{ fontSize:10, color:C.amb, marginTop:2 }}>Latest</div>}
                </div>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}

// ── IMPORT MODAL ──────────────────────────────────────────────────
function ImportModal({ brands, items, onItemsChange, onClose, cats }) {
  const [mode,    setMode]    = useState("merge");   // "merge" | "replace"
  const [preview, setPreview] = useState(null);      // { rows, errors }
  const [loading, setLoading] = useState(false);
  const fileRef = useRef(null);

  // Map a raw row (from xlsx or json) → item object
  function rowToItem(row) {
    const name = (row["Item Name"] || row["name"] || "").toString().trim();
    if (!name) return null;
    const pxRaw = parseFloat(row["Purchase Ex-GST"] || row["purchaseEx"] || 0);
    if (!pxRaw || isNaN(pxRaw) || pxRaw <= 0) return null;

    // GST — handle "5%" string or 0.05 number
    let gst = 0.05;
    const gstRaw = row["GST %"] || row["gst"];
    if (gstRaw != null) {
      const g = parseFloat(String(gstRaw).replace("%",""));
      if (!isNaN(g)) gst = g > 1 ? g / 100 : g;
    }

    // Brand — match by code
    const bCode = (row["Brand Code"] || row["bId"] || "").toString().trim().toUpperCase();
    const brand = brands.find(b => b.code === bCode);
    const bId   = brand ? brand.id : "";

    // Category
    const catRaw = (row["Category"] || row["cat"] || "").toString().trim();
    const catList = cats || CATS;
    const cat    = catList.includes(catRaw) ? catRaw : catList[0] || "Other Items";

    // Active
    const statusRaw = (row["Status"] || row["active"] || "Active").toString().trim();
    const active    = statusRaw === "Active" || statusRaw === "true" || statusRaw === true;

    const now = new Date().toISOString();
    return {
      id:           uid(),
      cat, bId, name,
      purchaseEx:   pxRaw,
      gst,
      customRL:     "",
      customDM:     "",
      customRLPrice:"",
      customDMPrice:"",
      sdmAddon:     "",
      plAddon:      "",
      customAddon:  "",
      active,
      notes:        (row["Notes"] || row["notes"] || "").toString().trim(),
      createdAt:    now,
      updatedAt:    now,
    };
  }

  function parseXLSX(file) {
    return new Promise((res, rej) => {
      const reader = new FileReader();
      reader.onload = e => {
        try {
          const wb   = XLSX.read(e.target.result, { type:"array" });
          const ws   = wb.Sheets[wb.SheetNames[0]];
          const rows = XLSX.utils.sheet_to_json(ws);
          res(rows);
        } catch(err) { rej(err); }
      };
      reader.onerror = rej;
      reader.readAsArrayBuffer(file);
    });
  }

  function parseJSON(file) {
    return new Promise((res, rej) => {
      const reader = new FileReader();
      reader.onload = e => {
        try {
          const parsed = JSON.parse(e.target.result);
          // support both raw array and { data: [...] } shape
          res(Array.isArray(parsed) ? parsed : (parsed.data || []));
        } catch(err) { rej(err); }
      };
      reader.onerror = rej;
      reader.readAsText(file);
    });
  }

  async function handleFile(file) {
    if (!file) return;
    setLoading(true);
    try {
      const ext  = file.name.split(".").pop().toLowerCase();
      const rows = ext === "json" ? await parseJSON(file) : await parseXLSX(file);
      const good = []; const errors = [];
      rows.forEach((r, i) => {
        const item = rowToItem(r);
        if (item) good.push(item);
        else errors.push("Row " + (i+2) + ": skipped — missing name or purchase price");
      });
      setPreview({ rows:good, errors });
    } catch(e) {
      toast("Could not read file — check format","err");
    }
    setLoading(false);
  }

  function doImport() {
    if (!preview || !preview.rows.length) return;
    const now = new Date().toISOString();
    if (mode === "replace") {
      onItemsChange(preview.rows);
      toast(preview.rows.length + " items imported (replaced all)");
    } else {
      // merge: match existing by name — update if found, add if new
      const next = [...items];
      let added = 0, updated = 0;
      preview.rows.forEach(newItem => {
        const idx = next.findIndex(x => x.name.toLowerCase() === newItem.name.toLowerCase());
        if (idx >= 0) {
          next[idx] = Object.assign({}, next[idx], newItem, { id:next[idx].id, createdAt:next[idx].createdAt, updatedAt:now });
          updated++;
        } else {
          next.push(newItem);
          added++;
        }
      });
      onItemsChange(next);
      toast(added + " added, " + updated + " updated");
    }
    onClose();
  }

  return (
    <div style={{ position:"fixed", inset:0, zIndex:600, background:"rgba(0,0,0,0.5)", display:"flex", alignItems:"center", justifyContent:"center", padding:16 }}>
      <div style={{ background:"#fff", borderRadius:20, padding:"22px 18px", maxWidth:420, width:"100%", maxHeight:"90vh", overflowY:"auto" }}>
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:16 }}>
          <div style={{ fontSize:17, fontWeight:800 }}>⬆ Import Items</div>
          <button onClick={onClose} style={{ width:32, height:32, borderRadius:8, border:"1.5px solid "+C.border, background:"#F9FAFB", cursor:"pointer", fontSize:15, color:C.sec, fontFamily:"inherit" }}>✕</button>
        </div>

        {/* Format guide */}
        <div style={{ background:"#F0F6FF", border:"1px solid "+C.dmBr, borderRadius:10, padding:"12px", marginBottom:14, fontSize:12, color:"#1E40AF", lineHeight:1.7 }}>
          <strong>Accepted formats:</strong> Excel (.xlsx) or JSON (.json) exported from this app.<br />
          Required columns: <strong>Item Name</strong>, <strong>Purchase Ex-GST</strong>.<br />
          Optional: Category, Brand Code, GST %, Status, Notes.
        </div>

        {/* Import mode */}
        <div style={{ marginBottom:14 }}>
          <div style={{ fontSize:11, fontWeight:700, color:C.sec, textTransform:"uppercase", letterSpacing:"0.6px", marginBottom:8 }}>Import Mode</div>
          <div style={{ display:"flex", gap:8 }}>
            {[
              { v:"merge",   label:"Merge",   hint:"Add new · Update existing by name" },
              { v:"replace", label:"Replace",  hint:"⚠ Delete all current items first" },
            ].map(o => {
              const on = mode === o.v;
              return (
                <button key={o.v} onClick={() => setMode(o.v)} style={{ flex:1, padding:"10px 8px", borderRadius:10, border:"1.5px solid "+(on?(o.v==="replace"?C.red:C.blue):C.border), background:on?(o.v==="replace"?C.redBg:"#EFF6FF"):"#F9FAFB", cursor:"pointer", fontFamily:"inherit", textAlign:"left" }}>
                  <div style={{ fontSize:13, fontWeight:700, color:on?(o.v==="replace"?C.red:C.blue):C.text }}>{o.label}</div>
                  <div style={{ fontSize:10, color:on?(o.v==="replace"?C.red:C.blue):C.mute, marginTop:2 }}>{o.hint}</div>
                </button>
              );
            })}
          </div>
        </div>

        {/* File picker */}
        <div style={{ marginBottom:14 }}>
          <input ref={fileRef} type="file" accept=".xlsx,.xls,.json" style={{ display:"none" }} onChange={e => handleFile(e.target.files[0])} />
          <BtnO color={C.blue} onClick={() => { setPreview(null); fileRef.current && fileRef.current.click(); }}>
            {loading ? "Reading file…" : "📂 Choose File (.xlsx or .json)"}
          </BtnO>
        </div>

        {/* Preview */}
        {preview && (
          <div>
            <div style={{ background:"#F0FDF4", border:"1px solid #BBF7D0", borderRadius:10, padding:"12px", marginBottom:10 }}>
              <div style={{ fontSize:13, fontWeight:800, color:C.profit, marginBottom:6 }}>
                ✅ {preview.rows.length} item{preview.rows.length !== 1 ? "s" : ""} ready to import
              </div>
              {preview.rows.slice(0, 4).map((r, i) => (
                <div key={i} style={{ fontSize:12, color:C.profit, padding:"3px 0", borderBottom:"1px solid #D1FAE5" }}>
                  {r.name} — {fp(r.purchaseEx)} Ex-GST
                </div>
              ))}
              {preview.rows.length > 4 && <div style={{ fontSize:11, color:C.mute, marginTop:4 }}>…and {preview.rows.length - 4} more</div>}
            </div>
            {preview.errors.length > 0 && (
              <div style={{ background:C.redBg, border:"1px solid #FCA5A5", borderRadius:10, padding:"10px", marginBottom:10 }}>
                <div style={{ fontSize:12, fontWeight:700, color:C.red, marginBottom:4 }}>⚠ {preview.errors.length} row{preview.errors.length!==1?"s":""} skipped</div>
                {preview.errors.slice(0,3).map((e,i) => <div key={i} style={{ fontSize:11, color:C.red }}>{e}</div>)}
                {preview.errors.length > 3 && <div style={{ fontSize:11, color:C.mute }}>…and {preview.errors.length-3} more</div>}
              </div>
            )}
            {mode === "replace" && (
              <div style={{ background:C.ambBg, border:"1px solid #FCD34D", borderRadius:10, padding:"10px", marginBottom:12, fontSize:12, color:C.amb, fontWeight:600 }}>
                ⚠ Replace mode will permanently delete all {items.length} existing items and replace with the {preview.rows.length} imported items.
              </div>
            )}
            <BtnP
              color={mode === "replace" ? C.red : C.blue}
              onClick={doImport}
            >
              {mode === "replace" ? "⚠ Replace All & Import" : "⬆ Import " + preview.rows.length + " Items"}
            </BtnP>
          </div>
        )}
      </div>
    </div>
  );
}

// ── MASTER VIEW ───────────────────────────────────────────────────
function todayISO() { const d = new Date(); return d.toISOString().slice(0,10); }
const EI = { cat:"Bed Sheets", bId:"", name:"", purchaseEx:"", gst:0.05, customRL:"", customDM:"", customRLPrice:"", customDMPrice:"", sdmAddon:"", sdmPct:"", plAddon:"", customAddon:"", purchaseDate:"", highlighted:false, commission:"", highlightNote:"", pinned:false, active:true, notes:"" };

function MasterView({ brands, items, onItemsChange, settings }) {
  const [open, setOpen] = useState(false); const [eid, setEid] = useState(null);
  const [form, setForm] = useState(EI); const [errs, setErrs] = useState({});
  const [fCat, setFCat] = useState("All"); const [fSt, setFSt] = useState("Active"); const [fBrand, setFBrand] = useState("All");
  const [search, setSearch] = useState("");
  const [conf, setConf] = useState(null); const [showExp, setShowExp] = useState(false); const [showImp, setShowImp] = useState(false);
  const [historyItem, setHistoryItem] = useState(null);
  const [showScroll, setShowScroll] = useState(false);
  useEffect(() => {
    function onScroll() { setShowScroll(window.scrollY > 150); }
    window.addEventListener("scroll", onScroll);
    return () => window.removeEventListener("scroll", onScroll);
  }, []);

  const brand  = brands.find(b => b.id === form.bId) || null;
  const prevEx = parseFloat(form.purchaseEx) || 0;
  const prev   = prevEx > 0 ? computeItem(Object.assign({}, form, { purchaseEx:prevEx }), brand, settings) : null;

  function close() { setOpen(false); setEid(null); }
  function openAdd() { setForm(Object.assign({},EI,{_sigRL:""})); setEid(null); setErrs({}); setOpen(true); }
  function openEdit(it) {
    setForm(Object.assign({}, EI, it, { customDM:it.customDM||"", customRLPrice:it.customRLPrice||"", customDMPrice:it.customDMPrice||"", sdmAddon:it.sdmAddon!=null?String(it.sdmAddon):"", sdmPct:it.sdmPct||"", purchaseDate:it.purchaseDate||"", highlighted:!!it.highlighted, commission:it.commission||"", highlightNote:it.highlightNote||"", pinned:!!it.pinned, _sigRL:"" }));
    setEid(it.id); setErrs({}); setOpen(true);
  }
  function validate() {
    const e = {};
    if (!form.name.trim()) e.name = "Item name required";
    const px = parseFloat(form.purchaseEx);
    if (!form.purchaseEx || isNaN(px) || px <= 0) e.px = "Enter a valid purchase price";
    if (form.customRL !== "" && form.customRL != null) { const v = parseFloat(form.customRL); if (isNaN(v)||v<0||v>100) e.crl = "Enter 0–100"; }
    if (form.customDM !== "" && form.customDM != null) { const v = parseFloat(form.customDM); if (isNaN(v)||v<0||v>100) e.cdm = "Enter 0–100"; }
    if (form.customRLPrice !== "" && form.customRLPrice != null) { const v = parseFloat(form.customRLPrice); if (isNaN(v)||v<=0) e.crlp = "Enter a valid price > 0"; }
    if (form.customDMPrice !== "" && form.customDMPrice != null) { const v = parseFloat(form.customDMPrice); if (isNaN(v)||v<=0) e.cdmp = "Enter a valid price > 0"; }
    if (form.sdmAddon !== "" && form.sdmAddon != null) { const v = parseFloat(form.sdmAddon); if (isNaN(v)||v<0) e.sdma = "Enter 0 or more"; }
    if (form.plAddon === "custom") { const cv = parseFloat(form.customAddon); if (isNaN(cv)||cv<=0) e.ca = "Enter a valid amount"; }
    setErrs(e);
    return Object.keys(e).length === 0;
  }
  function save() {
    if (isDuplicate) return toast("Item already exists — use a unique name","err");
    if (!validate()) return;
    const now = new Date().toISOString();
    const orig = eid ? items.find(i => i.id === eid) : null;
    const newPx = parseFloat(form.purchaseEx);
    const oldPx = orig ? parseFloat(orig.purchaseEx) : null;
    const pDate = form.purchaseDate || todayISO();
    // Append to purchaseHistory if price changed or new item
    const prevHistory = (orig && orig.purchaseHistory) ? orig.purchaseHistory : [];
    const priceChanged = !orig || oldPx !== newPx;
    const newHistory = priceChanged
      ? [...prevHistory, { price:newPx, date:pDate, recordedAt:now }]
      : prevHistory;
    const entry = Object.assign({}, form, { id:eid||uid(), purchaseEx:newPx, gst:parseFloat(form.gst), purchaseDate:pDate, purchaseHistory:newHistory, createdAt:orig?(orig.createdAt||now):now, updatedAt:now });
    onItemsChange(eid ? items.map(i => i.id===eid?entry:i) : [...items, entry]);
    toast(eid ? "Item updated" : "Item added"); close();
  }
  function del(id) { onItemsChange(items.filter(i => i.id !== id)); setConf(null); toast("Item removed","warn"); }
  function toggle(id) { onItemsChange(items.map(i => i.id===id ? Object.assign({},i,{active:!i.active,updatedAt:new Date().toISOString()}) : i)); }

  const list = useMemo(() => {
    return items.filter(i => {
      if (fCat !== "All" && i.cat !== fCat) return false;
      if (fBrand !== "All" && i.bId !== fBrand) return false;
      if (fSt === "Active" && !i.active) return false;
      if (fSt === "Inactive" && i.active) return false;
      if (search) {
        const q = search.toLowerCase();
        const b = brands.find(x => x.id===i.bId);
        if (!i.name.toLowerCase().includes(q) && !(b&&b.name.toLowerCase().includes(q)) && !(b&&b.code.toLowerCase().includes(q))) return false;
      }
      return true;
    }).map(i => { const b = brands.find(x => x.id===i.bId)||null; return Object.assign({}, i, computeItem(i,b,settings), {_b:b}); })
      .sort((a,b) => new Date(b.updatedAt||0) - new Date(a.updatedAt||0));
  }, [items, brands, settings, fCat, fSt, fBrand, search]);

  // Real-time duplicate check — exclude current item when editing
  const isDuplicate = form.name.trim() !== "" && items.some(i => i.name.trim().toLowerCase() === form.name.trim().toLowerCase() && i.id !== eid);

  function ei(k) { return errs[k] ? Object.assign({}, INP, { borderColor:C.red }) : INP; }
  const cats = (settings.categories && settings.categories.length > 0) ? settings.categories : CATS;
  const rlSrc = fpctRaw(form.customRL) ? "override ("+fpctRaw(form.customRL)+")" : brand ? "brand "+brand.code+" ("+(brand.rlMarkup*100).toFixed(0)+"%)" : "default ("+(settings.defaultRL*100).toFixed(0)+"%)";
  const dmSrc = fpctRaw(form.customDM) ? "override ("+fpctRaw(form.customDM)+")" : brand&&brand.dmMarkup!=null ? "brand "+brand.code+" ("+(brand.dmMarkup*100).toFixed(0)+"%)" : settings.defaultDM ? "default ("+(parseFloat(settings.defaultDM)*100).toFixed(0)+"%)" : "not set";
  const rlPriceLocked = form.customRLPrice !== "" && form.customRLPrice != null && !isNaN(parseFloat(form.customRLPrice)) && parseFloat(form.customRLPrice) > 0;
  const dmPriceLocked = form.customDMPrice !== "" && form.customDMPrice != null && !isNaN(parseFloat(form.customDMPrice)) && parseFloat(form.customDMPrice) > 0;
  const sdmHas = form.sdmAddon !== "" && form.sdmAddon != null && !isNaN(parseFloat(form.sdmAddon));

  return (
    <div style={{ padding:"16px 16px 80px" }}>
      {conf && <Confirm title="Delete Item?" msg="This will permanently remove the item." onOk={() => del(conf)} onNo={() => setConf(null)} />}
      {showExp && <ExportModal items={items} brands={brands} settings={settings} onClose={() => setShowExp(false)} />}
      {showImp && <ImportModal items={items} brands={brands} onItemsChange={onItemsChange} onClose={() => setShowImp(false)} cats={cats} />}
      {historyItem && <PurchaseHistoryModal item={historyItem} onClose={() => setHistoryItem(null)} />}
      <div style={{ background:C.ambBg, border:"1px solid #FCD34D", borderRadius:10, padding:"9px 13px", marginBottom:12, fontSize:12, color:C.amb, fontWeight:600 }}>
        ⚠ Purchase WITHOUT GST. RL/DM/PL are Inc-GST. SDM is Ex-GST (customer pays + GST on top).
      </div>
      {/* Scroll buttons */}
      {showScroll && (
        <div style={{ position:"fixed", right:16, bottom:80, zIndex:300, display:"flex", flexDirection:"column", gap:8 }}>
          <button onClick={() => window.scrollTo({top:0,behavior:"smooth"})} style={{ width:40, height:40, borderRadius:20, background:C.navy, color:"#fff", border:"none", fontSize:18, cursor:"pointer", boxShadow:"0 3px 12px rgba(0,0,0,0.2)", display:"flex", alignItems:"center", justifyContent:"center" }}>↑</button>
          <button onClick={() => window.scrollTo({top:document.body.scrollHeight,behavior:"smooth"})} style={{ width:40, height:40, borderRadius:20, background:C.navy, color:"#fff", border:"none", fontSize:18, cursor:"pointer", boxShadow:"0 3px 12px rgba(0,0,0,0.2)", display:"flex", alignItems:"center", justifyContent:"center" }}>↓</button>
        </div>
      )}
      <input style={Object.assign({},INP,{marginBottom:10})} placeholder="🔍 Search by item or brand..." value={search} onChange={e => setSearch(e.target.value)} />
      <div style={{ display:"flex", gap:8, marginBottom:8, flexWrap:"wrap" }}>
        <select style={Object.assign({},SEL,{flex:1,padding:"9px 11px"})} value={fCat} onChange={e => setFCat(e.target.value)}>
          <option value="All">All Categories</option>{cats.map(c => <option key={c} value={c}>{c}</option>)}
        </select>
        <select style={Object.assign({},SEL,{flex:1,padding:"9px 11px"})} value={fBrand} onChange={e => setFBrand(e.target.value)}>
          <option value="All">All Brands</option>{brands.map(b => <option key={b.id} value={b.id}>{b.code+" — "+b.name}</option>)}
        </select>
      </div>
      <div style={{ display:"flex", border:"1px solid "+C.border, borderRadius:9, overflow:"hidden", marginBottom:12 }}>
        {["All","Active","Inactive"].map(s => <button key={s} onClick={() => setFSt(s)} style={{ flex:1, padding:"9px 11px", border:"none", cursor:"pointer", fontWeight:600, fontSize:12, fontFamily:"inherit", background:fSt===s?C.blue:"#fff", color:fSt===s?"#fff":C.sec }}>{s}</button>)}
      </div>
      <div style={{ display:"grid", gridTemplateColumns:"2fr 1fr 1fr", gap:8, marginBottom:12 }}>
        <BtnP onClick={openAdd}>+ Add Item</BtnP>
        <BtnO color={C.blue} onClick={() => setShowExp(true)}>⬇ Export</BtnO>
        <BtnO color={C.rl}   onClick={() => setShowImp(true)}>⬆ Import</BtnO>
      </div>
      <div style={{ fontSize:12, color:C.mute, marginBottom:10, fontWeight:600 }}>{list.length + " item" + (list.length !== 1 ? "s" : "")}</div>
      {list.length === 0 && <div style={{ textAlign:"center", padding:"36px", color:C.mute, fontSize:14 }}>No items. Add using the button above.</div>}
      {list.map(it => (
        <div key={it.id}>
          <ItemCard it={it} isAdmin showDate sdmUnlocked={true} />
          <div style={{ display:"flex", gap:8, marginTop:-4, marginBottom:8 }}>
            <button onClick={() => openEdit(it)} style={{ flex:1, padding:"9px", borderRadius:8, border:"1.5px solid "+C.border, background:"#F9FAFB", cursor:"pointer", fontSize:13, fontWeight:600, color:C.sec, fontFamily:"inherit" }}>Edit</button>
            <button onClick={() => setHistoryItem(it)} style={{ padding:"9px 11px", borderRadius:8, border:"1.5px solid "+C.dmBr, background:C.dmBg, cursor:"pointer", fontSize:12, fontWeight:600, color:C.dm, fontFamily:"inherit" }}>📅</button>
            <button onClick={() => toggle(it.id)} style={{ flex:1, padding:"9px", borderRadius:8, border:"1.5px solid "+C.border, background:"#F9FAFB", cursor:"pointer", fontSize:12, fontWeight:600, color:it.active?"#991B1B":"#065F46", fontFamily:"inherit" }}>{it.active ? "Set Inactive" : "Set Active"}</button>
            <button onClick={() => setConf(it.id)} style={{ width:40, borderRadius:8, border:"1.5px solid #FCA5A5", background:C.redBg, cursor:"pointer", fontSize:15, color:C.red, fontFamily:"inherit" }}>✕</button>
          </div>
        </div>
      ))}

      <Drawer open={open} onClose={close} title={eid ? "Edit Item" : "Add Item"}
        footer={<div style={{ display:"flex", gap:9 }}><BtnO onClick={close}>Cancel</BtnO><BtnP color={isDuplicate?"#9CA3AF":eid?"#F59E0B":C.blue} onClick={save}>{eid ? "Update Item" : "Save Item"}</BtnP></div>}>
        <Fld label="Category *"><select style={SEL} value={form.cat} onChange={e => setForm(p => ({...p,cat:e.target.value,bId:""}))}>
          {cats.map(c => <option key={c} value={c}>{c}</option>)}
        </select></Fld>
        <Fld label="Brand"><select style={SEL} value={form.bId} onChange={e => setForm(p => ({...p,bId:e.target.value}))}>
          <option value="">— Select Brand (optional) —</option>
          {brands.map(b => <option key={b.id} value={b.id}>{b.code + " — " + b.name}</option>)}
        </select>
        {brand && <div style={{ marginTop:4, display:"flex", gap:6 }}><Chip col={C.rl} bg={C.rlBg} br={C.rlBr}>{"RL " + fpct(brand.rlMarkup)}</Chip>{brand.dmMarkup!=null?<Chip col={C.dm} bg={C.dmBg} br={C.dmBr}>{"DM " + fpct(brand.dmMarkup)}</Chip>:<Chip col={C.mute} bg="#F9FAFB" br={C.border}>DM not set</Chip>}</div>}
        </Fld>
        <Fld label="Item Name *" err={errs.name || (isDuplicate ? "Item already exists — duplicate name not allowed" : "")}>
          <input style={Object.assign({},ei("name"),{borderColor:isDuplicate?C.red:errs.name?C.red:C.border})} value={form.name} placeholder="e.g. King Cotton Bedsheet 200TC" onChange={e => setForm(p => ({...p,name:e.target.value}))} />
        </Fld>
        {/* Signature brand: reverse-calculate purchase price from RL inc-GST */}
        {brand && brand.name.toLowerCase().includes("signature") && (
          <div style={{ background:"#FFF7ED", border:"1px solid #FDBA74", borderRadius:11, padding:"12px", marginBottom:14 }}>
            <div style={{ fontSize:11, fontWeight:800, color:"#C2410C", textTransform:"uppercase", letterSpacing:"0.6px", marginBottom:4 }}>🧮 Signature — Calculate Purchase from RL Price</div>
            <div style={{ fontSize:11, color:"#C2410C", marginBottom:10 }}>Enter the RL price (Inc-GST) → purchase price is auto-filled as RL × 86%</div>
            <Fld label="RL Price Inc-GST (₹)" hint={(() => { const v = parseFloat(form._sigRL); return (!isNaN(v)&&v>0) ? "Purchase Ex-GST = " + fp(+(v*0.86).toFixed(2)) : "e.g. 1000 → purchase = ₹860"; })()}>
              <input
                style={Object.assign({},INP,{fontWeight:700,fontSize:16,color:"#C2410C",borderColor:"#FDBA74"})}
                type="number" step="1" min="0"
                placeholder="e.g. 1000"
                value={form._sigRL||""}
                onChange={e => {
                  const v = parseFloat(e.target.value);
                  const px = (!isNaN(v)&&v>0) ? (+(v*0.86).toFixed(2)).toString() : "";
                  setForm(p => ({...p, _sigRL:e.target.value, purchaseEx:px}));
                }}
              />
            </Fld>
          </div>
        )}
        <Fld label="Purchase Price WITHOUT GST (₹) *" err={errs.px}>
          <input style={Object.assign({},ei("px"),{fontWeight:700,fontSize:16,color:C.blue})} type="number" step="0.01" placeholder="e.g. 380" value={form.purchaseEx} onChange={e => setForm(p => ({...p,purchaseEx:e.target.value, _sigRL:""}))} />
        </Fld>
        <Fld label="GST % *" hint="Default 5% for textiles.">
          <select style={SEL} value={form.gst} onChange={e => setForm(p => ({...p,gst:parseFloat(e.target.value)}))}>
            {GST_OPTS.map(o => <option key={o.v} value={o.v}>{o.label}</option>)}
          </select>
          {prevEx > 0 && <div style={{ marginTop:4, fontSize:12, color:C.amb, fontWeight:600 }}>Purchase Inc-GST = {fp(calcIncGST(prevEx, form.gst))}</div>}
        </Fld>
        <Fld label="Purchase Date" hint={form.purchaseDate ? "" : "Defaults to today if left blank"}>
          <div style={{ display:"flex", gap:8, alignItems:"center" }}>
            <input style={Object.assign({},INP,{flex:1})} type="date" value={form.purchaseDate} onChange={e => setForm(p => ({...p,purchaseDate:e.target.value}))} />
            {form.purchaseDate && <button onClick={() => setForm(p => ({...p,purchaseDate:""}))} style={{ padding:"8px 12px", borderRadius:8, border:"1.5px solid "+C.border, background:"#F9FAFB", cursor:"pointer", fontSize:12, color:C.sec, fontFamily:"inherit", whiteSpace:"nowrap" }}>Reset to today</button>}
          </div>
        </Fld>
        {/* RL */}
        <div style={{ background:"#F8FDF9", border:"1px solid "+C.rlBr, borderRadius:11, padding:"12px", marginBottom:14 }}>
          <div style={{ fontSize:11, fontWeight:800, color:C.rl, textTransform:"uppercase", letterSpacing:"0.6px", marginBottom:10 }}>🟢 RL — Wholesale (Inc-GST)</div>
          <MarginOverride label={rlPriceLocked?"Custom RL Margin % (ignored — price locked)":"Custom RL Margin %"} col={C.rl} bg={C.rlBg} br={C.rlBr} value={form.customRL} onChange={v => setForm(p => ({...p,customRL:v}))} srcLabel={rlSrc} err={errs.crl} />
          <PriceOverride  label="Custom RL Price ₹ (Direct)" col={C.rl} bg={C.rlBg} br={C.rlBr} value={form.customRLPrice} onChange={v => setForm(p => ({...p,customRLPrice:v}))} err={errs.crlp} hint="Enter exact RL price — margin % ignored" />
        </div>
        {/* DM */}
        <div style={{ background:"#F0F6FF", border:"1px solid "+C.dmBr, borderRadius:11, padding:"12px", marginBottom:14 }}>
          <div style={{ fontSize:11, fontWeight:800, color:C.dm, textTransform:"uppercase", letterSpacing:"0.6px", marginBottom:10 }}>🔵 DM — Special Wholesale (Inc-GST)</div>
          <MarginOverride label={dmPriceLocked?"Custom DM Margin % (ignored — price locked)":"Custom DM Margin %"} col={C.dm} bg={C.dmBg} br={C.dmBr} value={form.customDM} onChange={v => setForm(p => ({...p,customDM:v}))} srcLabel={dmSrc} err={errs.cdm} />
          <PriceOverride  label="Custom DM Price ₹ (Direct)" col={C.dm} bg={C.dmBg} br={C.dmBr} value={form.customDMPrice} onChange={v => setForm(p => ({...p,customDMPrice:v}))} err={errs.cdmp} hint="Enter exact DM price — margin % ignored" />
        </div>
        {/* SDM */}
        <div style={{ background:C.sdmBg, border:"1px solid "+C.sdmBr, borderRadius:11, padding:"12px", marginBottom:14 }}>
          <div style={{ fontSize:11, fontWeight:800, color:C.sdm, textTransform:"uppercase", letterSpacing:"0.6px", marginBottom:4 }}>🟠 SDM — Special DM (Ex-GST, + GST extra)</div>
          <div style={{ fontSize:11, color:C.sdm, marginBottom:10 }}>Customer pays this + GST. Hidden behind PIN in salesman view.</div>
          <MarginOverride
            label={form.sdmPct!==""&&form.sdmPct!=null&&!isNaN(parseFloat(form.sdmPct))?"SDM Margin % (active)":"SDM Margin % (Override)"}
            col={C.sdm} bg={C.sdmBg} br={C.sdmBr}
            value={form.sdmPct}
            onChange={v => setForm(p => ({...p, sdmPct:v, sdmAddon:""}))}
            srcLabel={prevEx>0&&form.sdmPct!==""&&!isNaN(parseFloat(form.sdmPct)) ? "SDM = "+fp(+(prevEx*(1+parseFloat(form.sdmPct)/100)).toFixed(2))+" Ex-GST" : "Enter % or use add-on below"}
            err={errs.sdma}
          />
          <Fld label="— OR — SDM Add-on ₹ (fixed amount on purchase Ex-GST)"
            hint={prevEx>0&&form.sdmAddon!==""&&!isNaN(parseFloat(form.sdmAddon))&&form.sdmPct===""
              ? "SDM = "+fp(prevEx+parseFloat(form.sdmAddon))+" Ex-GST · Profit = ₹"+parseFloat(form.sdmAddon).toFixed(2)
              : form.sdmPct!==""?"Clear margin % above first to use add-on":"e.g. 20 → SDM = purchase + ₹20"}>
            <input
              style={Object.assign({},INP,{
                borderColor:errs.sdma?C.red:form.sdmAddon!==""&&form.sdmPct===""?C.sdm:C.border,
                color:form.sdmAddon!==""&&form.sdmPct===""?C.sdm:C.text,
                fontWeight:form.sdmAddon!==""&&form.sdmPct===""?700:400,
                opacity:form.sdmPct!==""?0.5:1
              })}
              type="number" step="1" min="0" placeholder="e.g. 20"
              disabled={form.sdmPct!==""}
              value={form.sdmAddon}
              onChange={e => setForm(p => ({...p, sdmAddon:e.target.value}))}
            />
          </Fld>
        </div>
        {/* PL */}
        <Fld label="PL Add-on ₹ (on top of RL — Inc-GST total)" err={errs.ca}>
          <select style={SEL} value={form.plAddon||""} onChange={e => { const v = e.target.value; setForm(p => ({...p,plAddon:v===""?"":v==="custom"?"custom":parseInt(v)})); }}>
            {ADDON_OPTS.map(o => <option key={String(o.v)} value={o.v}>{o.label}</option>)}
          </select>
          {form.plAddon === "custom" && <input style={Object.assign({},INP,{marginTop:7})} type="number" placeholder="Custom ₹ amount" value={form.customAddon} onChange={e => setForm(p => ({...p,customAddon:e.target.value}))} />}
        </Fld>
        {/* Live Preview */}
        {prev && (
          <div style={{ background:"#F8FAFF", border:"1px solid #BFDBFE", borderRadius:11, padding:"13px", marginBottom:6 }}>
            <div style={{ fontSize:11, fontWeight:700, color:C.blue, textTransform:"uppercase", letterSpacing:"0.5px", marginBottom:9 }}>📊 Live Price Preview</div>
            <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:7, marginBottom:9 }}>
              <div style={{ background:C.ambBg, border:"1px solid #FCD34D", borderRadius:8, padding:"7px 9px", textAlign:"center" }}>
                <div style={{ fontSize:9, color:C.amb, fontWeight:700, textTransform:"uppercase", marginBottom:2 }}>Inc-GST</div>
                <div style={{ fontSize:13, fontWeight:800, color:C.amb }}>{fp(prev.incGST)}</div>
              </div>
              <div style={{ background:"#F3F4F6", border:"1px solid "+C.border, borderRadius:8, padding:"7px 9px", textAlign:"center" }}>
                <div style={{ fontSize:9, color:C.sec, fontWeight:700, textTransform:"uppercase", marginBottom:2 }}>RL Markup</div>
                <div style={{ fontSize:13, fontWeight:800 }}>{prev.rlPriceLocked ? "📌 Fixed" : fpct(prev.rlM)}</div>
              </div>
            </div>
            <div style={{ fontSize:10, color:C.sec, marginBottom:5, fontWeight:600 }}>RL/DM/PL profit = (price ÷ (1+GST)) − purchase</div>
            <ProfitRow label="RL (Inc-GST)" col={C.rl}  price={prev.rl}  profit={prev.rlProfit}  locked={prev.rlPriceLocked} />
            <ProfitRow label="DM (Inc-GST)" col={C.dm}  price={prev.dm}  profit={prev.dmProfit}  locked={prev.dmPriceLocked} />
            <ProfitRow label="PL (Inc-GST)" col={C.pl}  price={prev.pl}  profit={prev.plProfit} />
            {prev.sdm && (
              <div>
                <div style={{ fontSize:10, color:C.sdm, marginTop:5, marginBottom:3, fontWeight:600 }}>SDM = Ex-GST · profit = add-on · customer pays {fp(prev.sdmInc)} inc-GST</div>
                <ProfitRow label="SDM (Ex-GST)" col={C.sdm} price={prev.sdm} profit={prev.sdmProfit} exGST />
              </div>
            )}
          </div>
        )}
        {/* Pin item */}
        <div style={{ background:form.pinned?"#EFF6FF":"#F9FAFB", border:"1.5px solid "+(form.pinned?C.blue:C.border), borderRadius:11, padding:"12px", marginBottom:14 }}>
          <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between" }}>
            <div>
              <div style={{ fontSize:12, fontWeight:800, color:form.pinned?C.blue:C.sec }}>📌 Pin in salesman view</div>
              <div style={{ fontSize:11, color:form.pinned?C.blue:C.mute, marginTop:2 }}>Appears at top of default All Prices view</div>
            </div>
            <button onClick={() => setForm(p => ({...p, pinned:!p.pinned}))}
              style={{ padding:"7px 16px", borderRadius:20, border:"1.5px solid "+(form.pinned?C.blue:C.border), background:form.pinned?C.blue:"#fff", color:form.pinned?"#fff":C.sec, fontWeight:700, fontSize:12, cursor:"pointer", fontFamily:"inherit" }}>
              {form.pinned ? "ON" : "OFF"}
            </button>
          </div>
        </div>
        {/* Highlights */}
        <div style={{ background:form.highlighted?"#FEF3C7":"#F9FAFB", border:"1.5px solid "+(form.highlighted?"#F59E0B":C.border), borderRadius:11, padding:"12px", marginBottom:14 }}>
          <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:form.highlighted?12:0 }}>
            <div>
              <div style={{ fontSize:12, fontWeight:800, color:form.highlighted?"#92400E":C.sec }}>⭐ Highlight this product</div>
              <div style={{ fontSize:11, color:form.highlighted?"#B45309":C.mute, marginTop:2 }}>Appears in salesman Highlights filter</div>
            </div>
            <button onClick={() => setForm(p => ({...p, highlighted:!p.highlighted}))}
              style={{ padding:"7px 16px", borderRadius:20, border:"1.5px solid "+(form.highlighted?"#F59E0B":C.border), background:form.highlighted?"#F59E0B":"#fff", color:form.highlighted?"#fff":C.sec, fontWeight:700, fontSize:12, cursor:"pointer", fontFamily:"inherit" }}>
              {form.highlighted ? "ON" : "OFF"}
            </button>
          </div>
          {form.highlighted && (
            <div>
              <Fld label="Commission ₹ per piece" hint="e.g. 20 → shown as +₹20/piece">
                <input style={INP} type="number" step="1" min="0" placeholder="e.g. 20" value={form.commission||""} onChange={e => setForm(p => ({...p,commission:e.target.value}))} />
              </Fld>
              <Fld label="Highlight Note (optional)" hint="e.g. Push this week · Clearance">
                <input style={INP} value={form.highlightNote||""} placeholder="e.g. Push this week" onChange={e => setForm(p => ({...p,highlightNote:e.target.value}))} />
              </Fld>
            </div>
          )}
        </div>
        <Fld label="Notes (optional)"><input style={INP} value={form.notes} placeholder="e.g. seasonal, imported..." onChange={e => setForm(p => ({...p,notes:e.target.value}))} /></Fld>
        <Fld label="Status">
          <div style={{ display:"flex", gap:8 }}>
            {["Active","Inactive"].map(s => { const on = (s==="Active"&&form.active)||(s==="Inactive"&&!form.active); return <button key={s} onClick={() => setForm(p => ({...p,active:s==="Active"}))} style={{ flex:1, padding:"10px", borderRadius:8, border:"1.5px solid "+(on?C.blue:C.border), background:on?"#EFF6FF":"#F9FAFB", color:on?C.blue:C.sec, fontWeight:700, fontSize:13, cursor:"pointer", fontFamily:"inherit" }}>{s}</button>; })}
          </div>
        </Fld>
      </Drawer>
    </div>
  );
}

// ── PRICE LIST VIEW ───────────────────────────────────────────────
// ── ESTIMATE VIEW ─────────────────────────────────────────────────
const PRICE_TYPES = [
  { id:"rl",  label:"RL",  col:C.rl,  bg:C.rlBg  },
  { id:"dm",  label:"DM",  col:C.dm,  bg:C.dmBg  },
  { id:"pl",  label:"PL",  col:C.pl,  bg:C.plBg  },
  { id:"sdm", label:"SDM", col:C.sdm, bg:C.sdmBg },
];

// ── WHATSAPP HELPER ───────────────────────────────────────────────
const WA_NUMBER = "917532002298"; // kept for backward compat
const WA_TEAM = "911140553488";
function buildWAText(est, co) {
  let msg = `*ESTIMATE — ${est.number}*\n${co}\nDate: ${fmtDate(est.createdAt)}\nSalesman: ${est.salesmanName}`;
  if (est.custName)  msg += `\nCustomer: ${est.custName}`;
  if (est.custPhone) msg += ` | ${est.custPhone}`;
  msg += `\n\n*Items:*`;
  (est.lines||[]).forEach(l => {
    const effP = parseFloat(l.customPrice)>0?parseFloat(l.customPrice):l.unitPrice;
    const disc = parseFloat(l.itemDiscount)||0;
    const amt  = effP * l.qty * (1 - disc/100);
    msg += `\n• ${l.name} × ${l.qty} @ ${fp(effP)} = ${fp(amt)}`;
    if (disc > 0) msg += ` (${disc}% off)`;
    if (l.includeGST) msg += ` +${((l.gstPct||0.05)*100).toFixed(0)}% GST`;
  });
  msg += `\n\nSubtotal: ${fp(est.subtotal)}`;
  if (est.billGST)    msg += `\nGST: ${fp(est.billGSTAmt)}`;
  if (est.otherAmt)   msg += `\n${est.otherLabel||"Other"}: ${fp(est.otherAmt)}`;
  if (est.adjustment) msg += `\nAdjustment: ${est.adjustment>0?"+":""}${fp(est.adjustment)}`;
  if (est.roundOff)    msg += `\nRound Off: ${est.roundOff>0?"+":""}${fp(est.roundOff)}`;
  msg += `\n*TOTAL: ${fp(est.grandTotal)}*`;
  if (est.narration)  msg += `\n\nNote: ${est.narration}`;
  return encodeURIComponent(msg);
}

const PRINT_CSS = `
@page { size: A4 portrait; margin: 8mm; }
@media print {
  body { margin: 0; }
  body * { visibility: hidden !important; }
  #vkf-print-area { display: block !important; position: absolute !important; left: 0 !important; top: 0 !important; width: 100% !important; background: #fff !important; padding: 0 !important; font-size: 11px !important; }
  #vkf-print-area * { visibility: visible !important; }
  #vkf-print-area table { page-break-inside: avoid; }
}
`;

// ── ESTIMATE VIEW ─────────────────────────────────────────────────
function EstimateView({ brands, items, settings, estimates, onEstimatesSave, isAdmin }) {
  const salesmen = settings.salesmen || [];
  const prefix   = settings.estimatePrefix || "VKF";

  const [salesmanName, setSalesmanName] = useState("");
  const [custName,     setCustName]     = useState("");
  const [custPhone,    setCustPhone]    = useState("");
  const [search,       setSearch]       = useState("");
  const [cartLines,    setCartLines]    = useState([]);
  const [defaultPT,    setDefaultPT]    = useState("");
  const [billGST,      setBillGST]      = useState(false);
  const [billGSTPct,   setBillGSTPct]   = useState(0.05);
  const [otherLabel,   setOtherLabel]   = useState("");
  const [otherAmt,     setOtherAmt]     = useState("");
  const [adjustment,   setAdjustment]   = useState("");
  const [narration,    setNarration]    = useState("");
  const [pricePopup,   setPricePopup]   = useState(null);
  const [saved,        setSaved]        = useState(null);
  const [editingEstId, setEditingEstId] = useState(null); // id being edited
  const [showHistory,  setShowHistory]  = useState(false);
  const [editingLine,  setEditingLine]  = useState(null);
  const [customPriceInput, setCustomPriceInput] = useState("");
  const [sdmUnlockedEst, setSdmUnlockedEst] = useState(false);
  const [sdmPinPopup,   setSdmPinPopup]   = useState(false);
  const [roundOff,      setRoundOff]       = useState("");

  useEffect(() => {
    if (!document.getElementById("vkf-print-css")) {
      const s = document.createElement("style");
      s.id = "vkf-print-css"; s.textContent = PRINT_CSS;
      document.head.appendChild(s);
    }
  }, []);

  useEffect(() => {
    if (window._pendingEstimateItem) {
      const { item, priceType } = window._pendingEstimateItem;
      window._pendingEstimateItem = null;
      const prices = { rl:item.rl, dm:item.dm, pl:item.pl, sdm:item.sdmInc||item.sdm };
      const up = prices[priceType];
      if (up) {
        if (!defaultPT) setDefaultPT(priceType);
        setCartLines(p => {
          const ex = p.find(l => l.itemId===item.id && l.priceType===priceType);
          if (ex) return p.map(l => l.itemId===item.id&&l.priceType===priceType?{...l,qty:l.qty+1}:l);
          return [...p, { id:uid(), itemId:item.id, name:item.name, priceType, unitPrice:up, originalPrice:up, customPrice:"", qty:1, itemDiscount:"", includeGST:false, gstPct:item.gst||0.05, _it:item }];
        });
        toast(item.name + " added to estimate");
      }
    }
  }, []);

  const enriched = useMemo(() => {
    return items.filter(i => i.active).map(i => {
      const b = brands.find(x => x.id===i.bId)||null;
      return Object.assign({}, i, computeItem(i,b,settings), {_b:b});
    });
  }, [items, brands, settings]);

  const searchResults = search.trim().length > 0
    ? enriched.filter(i => i.name.toLowerCase().includes(search.toLowerCase()) || (i._b && i._b.name.toLowerCase().includes(search.toLowerCase())))
    : [];

  function addToCart(it, priceType) {
    const prices = { rl:it.rl, dm:it.dm, pl:it.pl, sdm:it.sdmInc||it.sdm };
    const up = prices[priceType];
    if (!up) return toast("No price set for " + priceType.toUpperCase(),"warn");
    if (!defaultPT) setDefaultPT(priceType);
    setCartLines(p => {
      const ex = p.find(l => l.itemId===it.id && l.priceType===priceType);
      if (ex) return p.map(l => l.itemId===it.id&&l.priceType===priceType?{...l,qty:l.qty+1}:l);
      return [...p, { id:uid(), itemId:it.id, name:it.name, priceType, unitPrice:up, originalPrice:up, customPrice:"", qty:1, itemDiscount:"", includeGST:false, gstPct:it.gst||0.05, _it:it }];
    });
    setPricePopup(null); setSearch("");
    toast(it.name + " added");
  }

  function effPrice(l) { const cp=parseFloat(l.customPrice); return (!isNaN(cp)&&cp>0)?cp:l.unitPrice; }

  function lineTotal(l) {
    const base = effPrice(l) * l.qty;
    const disc = parseFloat(l.itemDiscount)||0;
    const afterDisc = base*(1-disc/100);
    return +(afterDisc + (l.includeGST?afterDisc*l.gstPct:0)).toFixed(2);
  }

  function lineProfit(l) {
    if (!l._it||!l._it.purchaseEx) return null;
    const ep = effPrice(l); const disc=parseFloat(l.itemDiscount)||0;
    const sellAfterDisc = ep*(1-disc/100);
    const sellExGST = l.includeGST ? sellAfterDisc/(1+l.gstPct) : sellAfterDisc;
    return +((sellExGST-l._it.purchaseEx)*l.qty).toFixed(2);
  }

  const subtotal    = cartLines.reduce((s,l)=>s+lineTotal(l),0);
  const billGSTAmt  = billGST?+(subtotal*billGSTPct).toFixed(2):0;
  const otherAmtN   = parseFloat(otherAmt)||0;
  const adjN        = parseFloat(adjustment)||0;
  const roundOffN   = parseFloat(roundOff)||0;
  const preRound    = +(subtotal+billGSTAmt+otherAmtN+adjN).toFixed(2);
  const autoRoundSuggestion = roundOffN===0 ? +(Math.round(preRound)-preRound).toFixed(2) : null;
  const grandTotal  = +(preRound+roundOffN).toFixed(2);
  const totalProfit = cartLines.reduce((s,l)=>{ const p=lineProfit(l); return s+(p||0); },0);

  function nextEstNo() {
    const used = (estimates||[]).map(e => parseInt((e.number||"").replace(/[^0-9]/g,""))||0);
    const max  = used.length?Math.max(...used):9999;
    return prefix+"-"+String(max+1);
  }

  function saveEstimate() {
    if (!salesmanName) return toast("Select salesman first","err");
    if (!cartLines.length) return toast("Add at least one item","err");
    const existing = editingEstId ? (estimates||[]).find(e=>e.id===editingEstId) : null;
    const est = {
      id:       existing ? existing.id     : uid(),
      number:   existing ? existing.number : nextEstNo(),
      createdAt:existing ? existing.createdAt : new Date().toISOString(),
      updatedAt:new Date().toISOString(),
      salesmanName, custName, custPhone,
      lines: cartLines.map(l=>{ const {_it,...rest}=l; return rest; }),
      billGST, billGSTPct, billGSTAmt,
      otherLabel, otherAmt:otherAmtN,
      adjustment:adjN, narration,
      subtotal, grandTotal, totalProfit, roundOff:roundOffN,
      status:"active",
    };
    const next = existing
      ? (estimates||[]).map(e=>e.id===existing.id?est:e)
      : [...(estimates||[]),est];
    onEstimatesSave(next);
    setSaved(est);
    setEditingEstId(null);
    toast(existing?"Estimate "+est.number+" updated!":"Estimate "+est.number+" saved!");
  }

  function clearAll() {
    setCartLines([]); setCustName(""); setCustPhone(""); setBillGST(false);
    setOtherLabel(""); setOtherAmt(""); setAdjustment(""); setNarration(""); setSaved(null); setDefaultPT(""); setRoundOff(""); setEditingEstId(null);
  }

  function loadEstimate(est) {
    setSalesmanName(est.salesmanName||"");
    setCustName(est.custName||"");
    setCustPhone(est.custPhone||"");
    setNarration(est.narration||"");
    setOtherLabel(est.otherLabel||"");
    setOtherAmt(est.otherAmt?String(est.otherAmt):"");
    setAdjustment(est.adjustment?String(est.adjustment):"");
    setRoundOff(est.roundOff?String(est.roundOff):"");
    setBillGST(!!est.billGST);
    setBillGSTPct(est.billGSTPct||0.05);
    setDefaultPT("");
    setSaved(null);
    setEditingEstId(est.id);
    const restored = (est.lines||[]).map(l => {
      const it = enriched.find(x=>x.id===l.itemId)||null;
      return {...l, _it:it, customPrice:l.customPrice||"", originalPrice:l.originalPrice||l.unitPrice};
    });
    setCartLines(restored);
    toast("Editing "+est.number+" — save to update");
  }

  function doPrint() {
    const el = document.getElementById("vkf-print-area");
    if (!el) return;
    el.style.display = "block";
    window.onafterprint = () => { el.style.display = "none"; window.onafterprint = null; };
    window.print();
  }

  function doDownloadPDF() {
    // Open a clean print window with just the estimate content
    const el = document.getElementById("vkf-print-area");
    if (!el) return;
    const fname = saved ? saved.number : "Estimate";
    const printWin = window.open("", "_blank", "width=800,height=900");
    if (!printWin) { toast("Allow popups to download PDF","warn"); return; }
    printWin.document.write(
      "<html><head><title>" + fname + "</title>" +
      "<style>" +
      "body{font-family:Arial,sans-serif;font-size:11px;color:#000;margin:0;padding:16px;}" +
      "table{width:100%;border-collapse:collapse;}" +
      "th,td{padding:5px 7px;border:1px solid #ccc;}" +
      "th{background:#f0f0f0;}" +
      "@page{size:A4 portrait;margin:8mm;}" +
      "@media print{body{padding:0;}}" +
      "</style></head><body>" +
      el.innerHTML +
      "<script>window.onload=function(){window.print();window.onafterprint=function(){window.close();};}<\/script>" +
      "</body></html>"
    );
    printWin.document.close();
  }

  const SS = { background:C.card, border:"1px solid "+C.border, borderRadius:12, padding:"14px", marginBottom:10 };

  if (saved) {
    return (
      <div style={{ padding:"16px 16px 80px" }}>
        {/* Hidden clean print area */}
        <div id="vkf-print-area" style={{ display:"none", fontFamily:"Arial,sans-serif", fontSize:11, color:"#000", padding:12 }}>
          <div style={{ textAlign:"center", borderBottom:"2px solid #000", paddingBottom:10, marginBottom:12 }}>
            <div style={{ fontSize:17, fontWeight:900 }}>{settings.co}</div>
            <div style={{ fontSize:11 }}>{settings.tag}</div>
            <div style={{ fontSize:13, fontWeight:700, marginTop:4, letterSpacing:1 }}>ESTIMATE / PERFORMA</div>
          </div>
          <div style={{ display:"flex", justifyContent:"space-between", marginBottom:12, fontSize:12 }}>
            <div><b>Estimate No:</b> {saved.number}<br/><b>Date:</b> {fmtDate(saved.createdAt)}<br/><b>Salesman:</b> {saved.salesmanName}</div>
            <div style={{ textAlign:"right" }}>{saved.custName?(<div><b>Customer:</b> {saved.custName}</div>):null}{saved.custPhone?(<div><b>Phone:</b> {saved.custPhone}</div>):null}</div>
          </div>
          {(()=>{
            const hasDisc = saved.lines.some(l=>parseFloat(l.itemDiscount)>0);
            return (
              <table style={{ width:"100%", borderCollapse:"collapse", fontSize:12, marginBottom:12 }}>
                <thead>
                  <tr style={{ background:"#f0f0f0" }}>
                    <th style={{ padding:"6px 8px", textAlign:"left", border:"1px solid #ccc" }}>Item</th>
                    <th style={{ padding:"6px 4px", textAlign:"center", border:"1px solid #ccc" }}>Qty</th>
                    <th style={{ padding:"6px 4px", textAlign:"right", border:"1px solid #ccc" }}>Rate</th>
                    {hasDisc && <th style={{ padding:"6px 4px", textAlign:"center", border:"1px solid #ccc" }}>Disc%</th>}
                    <th style={{ padding:"6px 4px", textAlign:"right", border:"1px solid #ccc" }}>Amount</th>
                  </tr>
                </thead>
                <tbody>
                  {saved.lines.map((l,i) => {
                    const ep=parseFloat(l.customPrice)>0?parseFloat(l.customPrice):l.unitPrice;
                    const disc=parseFloat(l.itemDiscount)||0;
                    const amt=ep*l.qty*(1-disc/100);
                    return (
                      <tr key={l.id} style={{ borderBottom:"1px solid #ddd" }}>
                        <td style={{ padding:"6px 8px", border:"1px solid #ddd" }}>{l.name}{l.includeGST?" (+"+((l.gstPct||0.05)*100).toFixed(0)+"% GST)":""}</td>
                        <td style={{ padding:"6px 4px", textAlign:"center", border:"1px solid #ddd" }}>{l.qty}</td>
                        <td style={{ padding:"6px 4px", textAlign:"right", border:"1px solid #ddd" }}>{fp(ep)}</td>
                        {hasDisc && <td style={{ padding:"6px 4px", textAlign:"center", border:"1px solid #ddd" }}>{disc>0?disc+"%":"-"}</td>}
                        <td style={{ padding:"6px 4px", textAlign:"right", fontWeight:700, border:"1px solid #ddd" }}>{fp(amt)}</td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            );
          })()}
          <div style={{ width:"55%", marginLeft:"auto", fontSize:13 }}>
            <div style={{ display:"flex", justifyContent:"space-between", padding:"4px 0", borderBottom:"1px solid #eee" }}><span>Subtotal</span><b>{fp(saved.subtotal)}</b></div>
            {saved.billGST&&<div style={{ display:"flex", justifyContent:"space-between", padding:"4px 0", borderBottom:"1px solid #eee" }}><span>GST ({((saved.billGSTPct||0.05)*100).toFixed(0)}%)</span><span>{fp(saved.billGSTAmt)}</span></div>}
            {saved.otherAmt>0&&<div style={{ display:"flex", justifyContent:"space-between", padding:"4px 0", borderBottom:"1px solid #eee" }}><span>{saved.otherLabel||"Other"}</span><span>{fp(saved.otherAmt)}</span></div>}
            {saved.adjustment?<div style={{ display:"flex", justifyContent:"space-between", padding:"4px 0", borderBottom:"1px solid #eee" }}><span>Adjustment</span><span>{saved.adjustment>0?"+":""}{fp(saved.adjustment)}</span></div>:null}
              {saved.roundOff?<div style={{ display:"flex", justifyContent:"space-between", padding:"4px 0", borderBottom:"1px solid #eee" }}><span>Round Off</span><span>{saved.roundOff>0?"+":""}{fp(saved.roundOff)}</span></div>:null}
            <div style={{ display:"flex", justifyContent:"space-between", padding:"6px 0", fontWeight:900, fontSize:16, borderTop:"2px solid #000" }}><span>GRAND TOTAL</span><span>{fp(saved.grandTotal)}</span></div>
          </div>
          {saved.narration&&<div style={{ marginTop:12, fontSize:11, borderTop:"1px dashed #ccc", paddingTop:8 }}><b>Note:</b> {saved.narration}</div>}
          <div style={{ marginTop:16, fontSize:10, textAlign:"center", color:"#888", borderTop:"1px dashed #ccc", paddingTop:8 }}>This is a computer generated estimate — {settings.co}</div>
        </div>

        <div style={{ background:"#F0FDF4", border:"1px solid #BBF7D0", borderRadius:12, padding:"14px", marginBottom:14 }}>
          <div style={{ fontSize:16, fontWeight:800, color:C.profit, marginBottom:2 }}>✅ {saved.number} saved!</div>
          <div style={{ fontSize:12, color:C.profit }}>By {saved.salesmanName}{saved.custName?" · "+saved.custName:""}</div>
          <div style={{ fontSize:18, fontWeight:900, color:C.navy, marginTop:4 }}>{fp(saved.grandTotal)}</div>
        </div>

        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:10, marginBottom:10 }}>
          <BtnP onClick={doPrint}>🖨 Print</BtnP>
          <BtnP color="#7C3AED" onClick={doDownloadPDF}>⬇ Save PDF</BtnP>
        </div>
        <div style={{ display:"grid", gridTemplateColumns:saved.custPhone?"1fr 1fr":"1fr", gap:10, marginBottom:10 }}>
          {saved.custPhone && (
            <a href={"https://wa.me/"+saved.custPhone.replace(/[^0-9]/g,"").replace(/^0/,"91")+"?text="+buildWAText(saved,settings.co)}
              target="_blank" rel="noopener noreferrer"
              style={{ display:"flex", alignItems:"center", justifyContent:"center", gap:6, padding:"13px 16px", borderRadius:10, background:"#25D366", color:"#fff", fontWeight:700, fontSize:13, textDecoration:"none" }}>
              💬 Customer
            </a>
          )}
          <a href={"https://wa.me/"+WA_TEAM+"?text="+buildWAText(saved,settings.co)}
            target="_blank" rel="noopener noreferrer"
            style={{ display:"flex", alignItems:"center", justifyContent:"center", gap:6, padding:"13px 16px", borderRadius:10, background:"#128C7E", color:"#fff", fontWeight:700, fontSize:13, textDecoration:"none" }}>
            💬 Team
          </a>
        </div>
        <BtnO color={C.blue} onClick={clearAll}>+ New Estimate</BtnO>
        <div style={{ height:10 }} />
        <BtnO onClick={()=>setShowHistory(true)}>📋 View All Estimates</BtnO>
        {showHistory&&<EstimateHistory estimates={estimates} onEstimatesSave={onEstimatesSave} isAdmin={isAdmin} onClose={()=>setShowHistory(false)} settings={settings} onLoadEstimate={loadEstimate} />}
      </div>
    );
  }

  return (
    <div style={{ padding:"16px 16px 80px" }}>
      {sdmPinPopup&&(
        <div style={{ position:"fixed", inset:0, zIndex:800, background:"rgba(0,0,0,0.55)", display:"flex", alignItems:"center", justifyContent:"center", padding:20 }}>
          <div style={{ background:"#fff", borderRadius:22, padding:"28px 24px", width:"100%", maxWidth:320, boxShadow:"0 20px 60px rgba(0,0,0,0.3)" }}>
            <div style={{ textAlign:"center", marginBottom:16 }}>
              <div style={{ fontSize:28, marginBottom:8 }}>🟠</div>
              <div style={{ fontSize:16, fontWeight:800, color:C.text }}>SDM Prices</div>
              <div style={{ fontSize:12, color:C.sec, marginTop:4 }}>Enter SDM PIN to use SDM pricing in estimates</div>
            </div>
            <PINPad label="Enter SDM PIN" pin={settings.sdmPIN||"9999"} onSuccess={()=>{ setSdmUnlockedEst(true); setSdmPinPopup(false); toast("SDM unlocked for this estimate 🟠"); }} />
            <div style={{ marginTop:16 }}><BtnO color={C.sec} onClick={()=>setSdmPinPopup(false)}>Cancel</BtnO></div>
          </div>
        </div>
      )}

      {pricePopup&&(
        <div style={{ position:"fixed", inset:0, zIndex:600, background:"rgba(0,0,0,0.5)", display:"flex", alignItems:"center", justifyContent:"center", padding:20 }}>
          <div style={{ background:"#fff", borderRadius:20, padding:"22px 18px", maxWidth:340, width:"100%" }}>
            <div style={{ fontSize:15, fontWeight:800, marginBottom:4 }}>{pricePopup.name}</div>
            <div style={{ fontSize:11, color:C.sec, marginBottom:14 }}>Select price type to add</div>
            {PRICE_TYPES.map(pt => {
              const prices={rl:pricePopup.rl,dm:pricePopup.dm,pl:pricePopup.pl,sdm:pricePopup.sdmInc||pricePopup.sdm};
              const pv=prices[pt.id]; if(!pv)return null;
              // SDM requires PIN unlock
              if(pt.id==="sdm"&&!sdmUnlockedEst){
                return(
                  <button key={pt.id} onClick={()=>setSdmPinPopup(true)}
                    style={{ display:"flex", alignItems:"center", justifyContent:"space-between", width:"100%", padding:"13px 16px", borderRadius:10, border:"1.5px dashed "+pt.col, background:"#F9FAFB", cursor:"pointer", fontFamily:"inherit", marginBottom:8 }}>
                    <span style={{ fontWeight:700, color:pt.col, fontSize:14 }}>SDM 🔒 PIN required</span>
                    <span style={{ fontSize:12, color:pt.col }}>Tap to unlock</span>
                  </button>
                );
              }
              return (
                <button key={pt.id} onClick={()=>addToCart(pricePopup,pt.id)}
                  style={{ display:"flex", alignItems:"center", justifyContent:"space-between", width:"100%", padding:"13px 16px", borderRadius:10, border:"1.5px solid "+pt.col, background:pt.bg, cursor:"pointer", fontFamily:"inherit", marginBottom:8 }}>
                  <span style={{ fontWeight:700, color:pt.col, fontSize:14 }}>{pt.label}{defaultPT===pt.id?" ✓":""}</span>
                  <span style={{ fontWeight:900, color:pt.col, fontSize:16 }}>{fp(pv)}</span>
                </button>
              );
            })}
            <BtnO onClick={()=>setPricePopup(null)} style={{ marginTop:4 }}>Cancel</BtnO>
          </div>
        </div>
      )}

      {editingLine&&(
        <div style={{ position:"fixed", inset:0, zIndex:700, background:"rgba(0,0,0,0.5)", display:"flex", alignItems:"center", justifyContent:"center", padding:20 }}>
          <div style={{ background:"#fff", borderRadius:18, padding:"22px 18px", maxWidth:300, width:"100%" }}>
            <div style={{ fontSize:14, fontWeight:800, marginBottom:4 }}>✏️ Custom Price</div>
            <div style={{ fontSize:11, color:C.sec, marginBottom:12 }}>{cartLines.find(l=>l.id===editingLine)?.name}</div>
            <Fld label="Custom Price ₹" hint={"Original: "+fp(cartLines.find(l=>l.id===editingLine)?.originalPrice)}>
              <input style={INP} type="number" step="1" placeholder="Enter custom price..."
                value={customPriceInput} onChange={e=>setCustomPriceInput(e.target.value)} autoFocus />
            </Fld>
            <div style={{ display:"flex", gap:8 }}>
              <BtnO onClick={()=>{ setCartLines(p=>p.map(l=>l.id===editingLine?{...l,customPrice:""}:l)); setEditingLine(null); setCustomPriceInput(""); }}>Reset</BtnO>
              <BtnP onClick={()=>{ setCartLines(p=>p.map(l=>l.id===editingLine?{...l,customPrice:customPriceInput}:l)); setEditingLine(null); setCustomPriceInput(""); }}>Apply</BtnP>
            </div>
          </div>
        </div>
      )}

      <div style={SS}>
        <div style={{ fontSize:13, fontWeight:800, color:editingEstId?C.amb:C.navy, marginBottom:editingEstId?8:12 }}>
          {editingEstId ? "✏️ Editing Estimate" : "🧾 New Estimate"}
        </div>
        {editingEstId && (
          <div style={{ background:C.ambBg, border:"1px solid #FCD34D", borderRadius:8, padding:"7px 11px", marginBottom:12, fontSize:12, color:C.amb, fontWeight:700 }}>
            ⚠ Editing: {(estimates||[]).find(e=>e.id===editingEstId)?.number||""} — saving will overwrite the original
            <button onClick={clearAll} style={{ float:"right", background:"none", border:"none", color:C.red, cursor:"pointer", fontWeight:700, fontFamily:"inherit", fontSize:12 }}>Cancel Edit</button>
          </div>
        )}
        <Fld label="Salesman *">
          <select style={SEL} value={salesmanName} onChange={e=>setSalesmanName(e.target.value)}>
            <option value="">— Select Salesman —</option>
            {salesmen.map(s=><option key={s} value={s}>{s}</option>)}
          </select>
        </Fld>
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:8 }}>
          <Fld label="Customer Name"><input style={INP} placeholder="Name (optional)" value={custName} onChange={e=>setCustName(e.target.value)} /></Fld>
          <Fld label="Phone"><input style={INP} placeholder="Phone (optional)" value={custPhone} onChange={e=>setCustPhone(e.target.value)} /></Fld>
        </div>
        {defaultPT&&(
          <div style={{ background:"#EFF6FF", border:"1px solid #BFDBFE", borderRadius:8, padding:"7px 11px", fontSize:12, color:C.blue, fontWeight:600 }}>
            🔒 Default: <strong>{defaultPT.toUpperCase()}</strong> &nbsp;
            <button onClick={()=>setDefaultPT("")} style={{ background:"none", border:"none", color:C.blue, cursor:"pointer", fontWeight:700, textDecoration:"underline", fontFamily:"inherit", fontSize:12 }}>Change</button>
          </div>
        )}
      </div>

      <div style={SS}>
        <div style={{ fontSize:12, fontWeight:800, color:C.navy, marginBottom:8 }}>Add Items</div>
        <input style={Object.assign({},INP,{marginBottom:0})} placeholder="🔍 Search item name or brand..." value={search} onChange={e=>setSearch(e.target.value)} />
        {searchResults.length>0&&(
          <div style={{ marginTop:6, maxHeight:220, overflowY:"auto", border:"1px solid "+C.border, borderRadius:9 }}>
            {searchResults.map(it=>(
              <button key={it.id} onClick={()=>{ 
                if(defaultPT==="sdm"&&!sdmUnlockedEst){ setSdmPinPopup(true); return; }
                defaultPT?addToCart(it,defaultPT):setPricePopup(it); 
              }}
                style={{ display:"flex", justifyContent:"space-between", alignItems:"center", width:"100%", padding:"10px 12px", background:"#fff", border:"none", borderBottom:"1px solid "+C.border, cursor:"pointer", fontFamily:"inherit", textAlign:"left" }}>
                <div>
                  <div style={{ fontSize:13, fontWeight:600, color:C.text }}>{it.name}</div>
                  <div style={{ fontSize:10, color:C.mute }}>{it.cat}{it._b?" · "+it._b.code:""}</div>
                </div>
                <span style={{ fontSize:11, color:C.blue, fontWeight:700 }}>{defaultPT?"+ Add ("+defaultPT.toUpperCase()+")":"+ Select Price"}</span>
              </button>
            ))}
          </div>
        )}
      </div>

      {cartLines.length>0&&(
        <div style={SS}>
          <div style={{ fontSize:12, fontWeight:800, color:C.navy, marginBottom:10 }}>Cart ({cartLines.length} item{cartLines.length!==1?"s":""})</div>
          {cartLines.map(l=>{
            const pt=PRICE_TYPES.find(p=>p.id===l.priceType)||PRICE_TYPES[0];
            const ep=effPrice(l); const hasCustom=parseFloat(l.customPrice)>0;
            return (
              <div key={l.id} style={{ background:"#F9FAFB", border:"1px solid "+C.border, borderRadius:10, padding:"11px 12px", marginBottom:8 }}>
                <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", marginBottom:8 }}>
                  <div style={{ flex:1, paddingRight:8 }}>
                    <div style={{ fontSize:13, fontWeight:700 }}>{l.name}</div>
                    <div style={{ display:"flex", gap:5, alignItems:"center", marginTop:3, flexWrap:"wrap" }}>
                      <div style={{ background:pt.bg, color:pt.col, border:"1px solid "+pt.col, borderRadius:20, padding:"1px 8px", fontSize:10, fontWeight:800 }}>{pt.label} · {fp(ep)}{hasCustom?" ✏️":""}</div>
                      {PRICE_TYPES.map(p2=>{
                        const prices2={rl:l._it&&l._it.rl,dm:l._it&&l._it.dm,pl:l._it&&l._it.pl,sdm:l._it&&(l._it.sdmInc||l._it.sdm)};
                        if(!prices2[p2.id]||p2.id===l.priceType)return null;
                        return <button key={p2.id} onClick={()=>{ if(p2.id==="sdm"&&!sdmUnlockedEst){ setSdmPinPopup(true); return; } setCartLines(prev=>prev.map(x=>x.id===l.id?{...x,priceType:p2.id,unitPrice:prices2[p2.id],originalPrice:prices2[p2.id],customPrice:""}:x)); }} style={{ fontSize:9,padding:"1px 5px",borderRadius:4,border:"1px solid "+p2.col,background:p2.bg,color:p2.col,cursor:"pointer",fontFamily:"inherit",fontWeight:700 }}>{p2.id==="sdm"&&!sdmUnlockedEst?"SDM 🔒":p2.label}</button>;
                      })}
                      <button onClick={()=>{ setEditingLine(l.id); setCustomPriceInput(l.customPrice||""); }} style={{ fontSize:12, background:"#F0F9FF", border:"1.5px solid "+C.blue, cursor:"pointer", color:C.blue, padding:"2px 8px", borderRadius:6, fontWeight:700, fontFamily:"inherit" }}>✏️ Price</button>
                    </div>
                  </div>
                  <button onClick={()=>setCartLines(p=>p.filter(x=>x.id!==l.id))} style={{ width:28,height:28,borderRadius:6,border:"1.5px solid #FCA5A5",background:C.redBg,cursor:"pointer",fontSize:13,color:C.red,fontFamily:"inherit" }}>✕</button>
                </div>
                <div style={{ display:"grid", gridTemplateColumns:"auto 1fr 1fr", gap:8, alignItems:"end" }}>
                  <div>
                    <div style={{ fontSize:10,color:C.sec,fontWeight:700,marginBottom:3 }}>QTY</div>
                    <div style={{ display:"flex", alignItems:"center", gap:4 }}>
                      <button onClick={()=>setCartLines(p=>p.map(x=>x.id===l.id?{...x,qty:Math.max(1,x.qty-1)}:x))} style={{ width:28,height:28,borderRadius:6,border:"1.5px solid "+C.border,background:"#fff",cursor:"pointer",fontFamily:"inherit",fontWeight:700,fontSize:14 }}>−</button>
                      <input type="number" min="1" value={l.qty} onChange={e=>setCartLines(p=>p.map(x=>x.id===l.id?{...x,qty:Math.max(1,parseInt(e.target.value)||1)}:x))} style={{ width:40,textAlign:"center",padding:"5px 2px",borderRadius:6,border:"1.5px solid "+C.border,fontSize:13,fontWeight:700,fontFamily:"inherit" }} />
                      <button onClick={()=>setCartLines(p=>p.map(x=>x.id===l.id?{...x,qty:x.qty+1}:x))} style={{ width:28,height:28,borderRadius:6,border:"1.5px solid "+C.border,background:"#fff",cursor:"pointer",fontFamily:"inherit",fontWeight:700,fontSize:14 }}>+</button>
                    </div>
                  </div>
                  <div>
                    <div style={{ fontSize:10,color:C.sec,fontWeight:700,marginBottom:3 }}>DISC %</div>
                    <input type="number" min="0" max="100" placeholder="0" value={l.itemDiscount} onChange={e=>setCartLines(p=>p.map(x=>x.id===l.id?{...x,itemDiscount:e.target.value}:x))} style={{ width:"100%",padding:"7px 8px",borderRadius:6,border:"1.5px solid "+C.border,fontSize:12,fontFamily:"inherit" }} />
                  </div>
                  <div>
                    <div style={{ fontSize:10,color:C.sec,fontWeight:700,marginBottom:3 }}>+GST</div>
                    <div style={{ display:"flex", gap:4 }}>
                      <button onClick={()=>setCartLines(p=>p.map(x=>x.id===l.id?{...x,includeGST:!x.includeGST}:x))}
                        style={{ padding:"5px 8px",borderRadius:6,border:"1.5px solid "+(l.includeGST?C.amb:C.border),background:l.includeGST?C.ambBg:"#fff",color:l.includeGST?C.amb:C.sec,fontWeight:700,fontSize:11,cursor:"pointer",fontFamily:"inherit" }}>
                        {l.includeGST?"ON":"OFF"}
                      </button>
                      {l.includeGST&&(
                        <select value={l.gstPct} onChange={e=>setCartLines(p=>p.map(x=>x.id===l.id?{...x,gstPct:parseFloat(e.target.value)}:x))}
                          style={{ padding:"5px 4px",borderRadius:6,border:"1.5px solid "+C.amb,fontSize:11,fontFamily:"inherit",background:C.ambBg,color:C.amb,fontWeight:700 }}>
                          <option value={0.05}>5%</option><option value={0.12}>12%</option><option value={0.18}>18%</option>
                        </select>
                      )}
                    </div>
                  </div>
                </div>
                <div style={{ marginTop:6, display:"flex", justifyContent:"space-between" }}>
                  <div style={{ fontSize:12,fontWeight:800,color:C.navy }}>Line total: {fp(lineTotal(l))}</div>
                  {isAdmin&&lineProfit(l)!==null&&<div style={{ fontSize:11,fontWeight:700,color:C.profit }}>Profit: {fp(lineProfit(l))}</div>}
                </div>
              </div>
            );
          })}
        </div>
      )}

      {cartLines.length>0&&(
        <div style={SS}>
          <div style={{ fontSize:12,fontWeight:800,color:C.navy,marginBottom:12 }}>Bill Summary</div>
          <div style={{ display:"flex", justifyContent:"space-between", fontSize:13, marginBottom:10 }}><span style={{ color:C.sec }}>Subtotal</span><span style={{ fontWeight:800 }}>{fp(subtotal)}</span></div>
          <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:10 }}>
            <div><div style={{ fontSize:12,fontWeight:700 }}>Add GST on bill</div><div style={{ fontSize:10,color:C.mute }}>Optional</div></div>
            <div style={{ display:"flex", gap:6, alignItems:"center" }}>
              <button onClick={()=>setBillGST(p=>!p)} style={{ padding:"7px 14px",borderRadius:20,border:"1.5px solid "+(billGST?C.amb:C.border),background:billGST?C.ambBg:"#fff",color:billGST?C.amb:C.sec,fontWeight:700,fontSize:12,cursor:"pointer",fontFamily:"inherit" }}>{billGST?"ON":"OFF"}</button>
              {billGST&&(<select value={billGSTPct} onChange={e=>setBillGSTPct(parseFloat(e.target.value))} style={{ padding:"7px 8px",borderRadius:8,border:"1.5px solid "+C.amb,fontSize:12,fontFamily:"inherit",background:C.ambBg,color:C.amb,fontWeight:700 }}><option value={0.05}>5%</option><option value={0.12}>12%</option><option value={0.18}>18%</option></select>)}
              {billGST&&<span style={{ fontSize:12,fontWeight:700,color:C.amb }}>{fp(billGSTAmt)}</span>}
            </div>
          </div>
          <div style={{ display:"grid", gridTemplateColumns:"2fr 1fr", gap:8, marginBottom:10 }}>
            <Fld label="Other Charges"><input style={INP} placeholder="e.g. Packing" value={otherLabel} onChange={e=>setOtherLabel(e.target.value)} /></Fld>
            <Fld label="Amount ₹"><input style={INP} type="number" placeholder="0" value={otherAmt} onChange={e=>setOtherAmt(e.target.value)} /></Fld>
          </div>
          <Fld label="Adjustment ₹" hint="Use − for discount"><input style={INP} type="number" placeholder="e.g. -50" value={adjustment} onChange={e=>setAdjustment(e.target.value)} /></Fld>
          <Fld label="Round Off ₹" hint="e.g. -2 or +3 to make round figure">
            <div style={{ display:"flex", gap:8, alignItems:"flex-start" }}>
              <input style={Object.assign({},INP,{flex:1})} type="number" step="0.01" placeholder="e.g. -2 or +3" value={roundOff} onChange={e=>setRoundOff(e.target.value)} />
              {autoRoundSuggestion!==null&&autoRoundSuggestion!==0&&(
                <button onClick={()=>setRoundOff(String(autoRoundSuggestion))}
                  style={{ padding:"11px 12px", borderRadius:9, border:"1.5px solid "+C.blue, background:"#EFF6FF", color:C.blue, fontWeight:700, fontSize:12, cursor:"pointer", fontFamily:"inherit", whiteSpace:"nowrap" }}>
                  Auto {autoRoundSuggestion>0?"+":""}{autoRoundSuggestion}
                </button>
              )}
            </div>
          </Fld>
          <Fld label="Narration / Note"><input style={INP} placeholder="e.g. Cash payment..." value={narration} onChange={e=>setNarration(e.target.value)} /></Fld>
          <div style={{ display:"flex", justifyContent:"space-between", background:C.navy, borderRadius:10, padding:"13px 16px", marginTop:4 }}>
            <span style={{ color:"#fff",fontSize:15,fontWeight:700 }}>Grand Total</span>
            <span style={{ color:"#fff",fontSize:20,fontWeight:900 }}>{fp(grandTotal)}</span>
          </div>
          {isAdmin&&(<div style={{ background:C.profBg,border:"1px solid #BBF7D0",borderRadius:8,padding:"8px 12px",marginTop:8,display:"flex",justifyContent:"space-between" }}><span style={{ fontSize:12,fontWeight:700,color:C.profit }}>💰 Estimated Profit</span><span style={{ fontSize:14,fontWeight:900,color:C.profit }}>{fp(totalProfit)}</span></div>)}
        </div>
      )}

      {cartLines.length>0&&(<div style={{ display:"grid", gridTemplateColumns:"2fr 1fr", gap:10 }}><BtnP color={C.profit} onClick={saveEstimate}>💾 Save Estimate</BtnP><BtnO color={C.red} onClick={clearAll}>Clear</BtnO></div>)}
      <div style={{ height:12 }} />
      <BtnO onClick={()=>setShowHistory(true)}>📋 View All Estimates</BtnO>
      {showHistory&&<EstimateHistory estimates={estimates} onEstimatesSave={onEstimatesSave} isAdmin={isAdmin} onClose={()=>setShowHistory(false)} settings={settings} onLoadEstimate={loadEstimate} />}
    </div>
  );
}

// ── ESTIMATE HISTORY MODAL ────────────────────────────────────────
function EstimateCard({ e, cancelled, age, canEdit, isAdmin, H24, H48, settings, onLoadEstimate, cancelEst, deleteEst }) {
  const [expanded, setExpanded] = useState(false);
  return (
    <div style={{ background:cancelled?"#FFF5F5":"#F9FAFB", border:"1px solid "+(cancelled?"#FCA5A5":C.border), borderRadius:10, padding:"12px", marginBottom:8, opacity:cancelled?0.75:1 }}>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", marginBottom:6 }}>
        <div style={{ flex:1 }}>
          <div style={{ display:"flex", alignItems:"center", gap:6, flexWrap:"wrap" }}>
            <span style={{ fontWeight:800, fontSize:14, color:C.navy }}>{e.number}</span>
            {cancelled && <span style={{ fontSize:10, color:C.red, fontWeight:800, background:C.redBg, padding:"1px 6px", borderRadius:4 }}>CANCELLED</span>}
            {!isAdmin && age>=H24 && age<H48 && <span style={{ fontSize:10, color:C.mute, background:"#F3F4F6", padding:"1px 6px", borderRadius:4 }}>View only</span>}
          </div>
          <div style={{ fontSize:11, color:C.sec, marginTop:2 }}>{e.salesmanName}{e.custName?" · "+e.custName:""}{e.custPhone?" · "+e.custPhone:""}</div>
          <div style={{ fontSize:10, color:C.mute, marginTop:2 }}>{fmtDateTime(e.createdAt)} · {e.lines.length} item{e.lines.length!==1?"s":""}</div>
        </div>
        <div style={{ textAlign:"right" }}>
          <div style={{ fontSize:16, fontWeight:900, color:C.navy }}>{fp(e.grandTotal)}</div>
          {isAdmin && e.totalProfit!=null && <div style={{ fontSize:11, color:C.profit, fontWeight:700 }}>+{fp(e.totalProfit)} profit</div>}
        </div>
      </div>
      {/* Expand button */}
      <button onClick={()=>setExpanded(p=>!p)}
        style={{ width:"100%", padding:"5px", borderRadius:6, border:"1px solid "+C.border, background:"#F3F4F6", cursor:"pointer", fontFamily:"inherit", fontSize:11, color:C.sec, fontWeight:600, marginBottom:expanded?8:0 }}>
        {expanded?"▲ Hide Details":"▼ View Details"}
      </button>
      {/* Expanded detail */}
      {expanded && (
        <div style={{ background:"#fff", border:"1px solid "+C.border, borderRadius:8, padding:"10px", marginBottom:6 }}>
          {(e.custName||e.custPhone) && <div style={{ fontSize:11, color:C.sec, marginBottom:6 }}>Customer: {e.custName||""}{e.custPhone?" · "+e.custPhone:""}</div>}
          {/* Item breakdown */}
          {(e.lines||[]).map((l,i) => {
            const ep = parseFloat(l.customPrice)>0?parseFloat(l.customPrice):l.unitPrice;
            const disc = parseFloat(l.itemDiscount)||0;
            const amt = ep*l.qty*(1-disc/100);
            const gstAmt = l.includeGST?amt*l.gstPct:0;
            const lineTotal = +(amt+gstAmt).toFixed(2);
            // Profit per line (admin only)
            let lineProfit = null;
            if (isAdmin && l.purchaseEx) {
              const sellExGST = l.includeGST?amt/(1+l.gstPct):amt;
              lineProfit = +((sellExGST/l.qty - l.purchaseEx)*l.qty).toFixed(2);
            }
            const pt = [{id:"rl",col:C.rl},{id:"dm",col:C.dm},{id:"pl",col:C.pl},{id:"sdm",col:C.sdm}].find(p=>p.id===l.priceType);
            return (
              <div key={l.id||i} style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", padding:"6px 0", borderBottom:"1px solid #F3F4F6" }}>
                <div style={{ flex:1, paddingRight:8 }}>
                  <div style={{ fontSize:12, fontWeight:600 }}>{l.name}</div>
                  <div style={{ fontSize:10, color:C.mute }}>
                    {l.qty} × {fp(ep)}
                    {disc>0?" − "+disc+"%":""}
                    {l.includeGST?" + GST":""}
                    {pt&&<span style={{ color:pt.col, fontWeight:700 }}> [{l.priceType.toUpperCase()}]</span>}
                  </div>
                </div>
                <div style={{ textAlign:"right" }}>
                  <div style={{ fontSize:12, fontWeight:700 }}>{fp(lineTotal)}</div>
                  {isAdmin && lineProfit!==null && <div style={{ fontSize:10, color:C.profit, fontWeight:700 }}>+{fp(lineProfit)}</div>}
                </div>
              </div>
            );
          })}
          {/* Totals */}
          <div style={{ marginTop:8, paddingTop:6, borderTop:"1px solid "+C.border }}>
            <div style={{ display:"flex", justifyContent:"space-between", fontSize:11, color:C.sec, marginBottom:2 }}><span>Subtotal</span><span>{fp(e.subtotal)}</span></div>
            {e.billGST && <div style={{ display:"flex", justifyContent:"space-between", fontSize:11, color:C.amb, marginBottom:2 }}><span>GST ({((e.billGSTPct||0.05)*100).toFixed(0)}%)</span><span>{fp(e.billGSTAmt)}</span></div>}
            {e.otherAmt>0 && <div style={{ display:"flex", justifyContent:"space-between", fontSize:11, color:C.sec, marginBottom:2 }}><span>{e.otherLabel||"Other"}</span><span>{fp(e.otherAmt)}</span></div>}
            {e.adjustment ? <div style={{ display:"flex", justifyContent:"space-between", fontSize:11, color:C.sec, marginBottom:2 }}><span>Adjustment</span><span>{e.adjustment>0?"+":""}{fp(e.adjustment)}</span></div> : null}
            {e.roundOff ? <div style={{ display:"flex", justifyContent:"space-between", fontSize:11, color:C.sec, marginBottom:2 }}><span>Round Off</span><span>{e.roundOff>0?"+":""}{fp(e.roundOff)}</span></div> : null}
            <div style={{ display:"flex", justifyContent:"space-between", fontSize:14, fontWeight:900, color:C.navy, borderTop:"1px solid "+C.border, paddingTop:5, marginTop:3 }}><span>TOTAL</span><span>{fp(e.grandTotal)}</span></div>
            {isAdmin && e.totalProfit!=null && <div style={{ display:"flex", justifyContent:"space-between", fontSize:12, fontWeight:700, color:C.profit, marginTop:4, background:C.profBg, padding:"5px 8px", borderRadius:6 }}><span>💰 Total Profit</span><span>{fp(e.totalProfit)}</span></div>}
          </div>
          {e.narration && <div style={{ fontSize:11, color:C.sec, fontStyle:"italic", marginTop:6 }}>Note: {e.narration}</div>}
        </div>
      )}
      {!cancelled && (
        <div style={{ display:"flex", gap:6, flexWrap:"wrap" }}>
          {settings && <a href={"https://wa.me/"+WA_NUMBER+"?text="+buildWAText(e,settings.co)} target="_blank" rel="noopener noreferrer" style={{ padding:"5px 10px",borderRadius:6,border:"1px solid #25D366",background:"#F0FFF4",color:"#15803D",fontWeight:700,fontSize:11,textDecoration:"none" }}>💬 WA</a>}
          {canEdit && onLoadEstimate && <button onClick={()=>{ onLoadEstimate(e); }} style={{ padding:"5px 10px",borderRadius:6,border:"1px solid "+C.blue,background:"#EFF6FF",color:C.blue,fontWeight:700,fontSize:11,cursor:"pointer",fontFamily:"inherit" }}>✏️ Edit</button>}
          {isAdmin && <button onClick={()=>cancelEst(e.id)} style={{ padding:"5px 10px",borderRadius:6,border:"1px solid "+C.amb,background:C.ambBg,color:C.amb,fontWeight:700,fontSize:11,cursor:"pointer",fontFamily:"inherit" }}>Cancel</button>}
          {isAdmin && <button onClick={()=>deleteEst(e.id)} style={{ padding:"5px 10px",borderRadius:6,border:"1px solid #FCA5A5",background:C.redBg,color:C.red,fontWeight:700,fontSize:11,cursor:"pointer",fontFamily:"inherit" }}>Delete</button>}
        </div>
      )}
    </div>
  );
}

function EstimateHistory({ estimates, onEstimatesSave, isAdmin, onClose, settings, onLoadEstimate }) {
  const now = Date.now();
  const H24 = 24*60*60*1000;
  const H48 = 48*60*60*1000;

  const list = (estimates||[]).slice().reverse().filter(e => {
    if (isAdmin) return true;
    return now - new Date(e.createdAt).getTime() < H48;
  });

  const today = new Date().toDateString();
  const todayList = list.filter(e => new Date(e.createdAt).toDateString()===today);

  function cancelEst(id) {
    onEstimatesSave((estimates||[]).map(e=>e.id===id?{...e,status:"cancelled"}:e));
    toast("Cancelled","warn");
  }
  function deleteEst(id) {
    onEstimatesSave((estimates||[]).filter(e=>e.id!==id));
    toast("Deleted","warn");
  }

  return (
    <div style={{ position:"fixed", inset:0, zIndex:600, background:"rgba(0,0,0,0.5)", display:"flex", alignItems:"center", justifyContent:"center", padding:16 }}>
      <div style={{ background:"#fff", borderRadius:20, padding:"20px 16px", maxWidth:440, width:"100%", maxHeight:"90vh", display:"flex", flexDirection:"column" }}>
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:14 }}>
          <div style={{ fontSize:16, fontWeight:800 }}>📋 {isAdmin?"All Estimates":"Your Estimates"}</div>
          <button onClick={onClose} style={{ width:32,height:32,borderRadius:8,border:"1.5px solid "+C.border,background:"#F9FAFB",cursor:"pointer",fontSize:15,color:C.sec,fontFamily:"inherit" }}>✕</button>
        </div>
        {todayList.length>0&&(
          <div style={{ background:C.profBg,border:"1px solid #BBF7D0",borderRadius:8,padding:"8px 12px",marginBottom:10,fontSize:12 }}>
            <span style={{ fontWeight:700,color:C.profit }}>Today: {todayList.length} · {fp(todayList.filter(e=>e.status!=="cancelled").reduce((s,e)=>s+e.grandTotal,0))}</span>
            {isAdmin&&<span style={{ color:C.profit,marginLeft:8 }}>· Profit: {fp(todayList.filter(e=>e.status!=="cancelled").reduce((s,e)=>s+(e.totalProfit||0),0))}</span>}
          </div>
        )}
        {!isAdmin&&<div style={{ fontSize:11,color:C.mute,marginBottom:8 }}>Showing last 48 hours only.</div>}
        <div style={{ flex:1, overflowY:"auto" }}>
          {list.length===0&&<div style={{ textAlign:"center",padding:36,color:C.mute }}>No estimates found.</div>}
          {list.map(e=>{
            const cancelled=e.status==="cancelled";
            const age=now-new Date(e.createdAt).getTime();
            const canEdit=isAdmin||age<H24;
            return <EstimateCard key={e.id} e={e} cancelled={cancelled} age={age} canEdit={canEdit} isAdmin={isAdmin} H24={H24} H48={H48} settings={settings} onLoadEstimate={onLoadEstimate} cancelEst={cancelEst} deleteEst={deleteEst} />;
          })}
        </div>
      </div>
    </div>
  );
}


const PRICE_FILTERS = [
  { id:"all",        label:"All Prices",   col:C.text,    bg:"#F3F4F6", br:C.border },
  { id:"rl",         label:"🟢 RL",        col:C.rl,      bg:C.rlBg,   br:C.rlBr },
  { id:"dm",         label:"🔵 DM",        col:C.dm,      bg:C.dmBg,   br:C.dmBr },
  { id:"pl",         label:"🟣 PL",        col:C.pl,      bg:C.plBg,   br:C.plBr },
  { id:"sdm",        label:"🟠 SDM 🔒",   col:C.sdm,     bg:C.sdmBg,  br:C.sdmBr },
  { id:"highlights", label:"⭐ Highlights", col:"#B45309", bg:"#FEF3C7", br:"#FCD34D" },
];

function PriceListView({ brands, items, settings, isAdmin, onAddToEstimate }) {
  const [viewMode, setViewMode] = useState("date");
  const [fCat,     setFCat]     = useState("All");
  const [fBrand,   setFBrand]   = useState("All");
  const [search,   setSearch]   = useState("");
  const [priceFilter, setPriceFilter] = useState("all");
  const [sdmUnlocked, setSdmUnlocked] = useState(false);
  const [showSDMPin,  setShowSDMPin]  = useState(false);
  const [showScroll,  setShowScroll]  = useState(false);
  useEffect(() => {
    function onScroll() { setShowScroll(window.scrollY > 150); }
    window.addEventListener("scroll", onScroll);
    return () => window.removeEventListener("scroll", onScroll);
  }, []);

  function handleSDMClick() {
    if (sdmUnlocked) return;
    setShowSDMPin(true);
  }
  function handleSDMFilterClick() {
    if (priceFilter === "sdm") { setPriceFilter("all"); return; }
    if (!sdmUnlocked) { setShowSDMPin(true); return; }
    setPriceFilter("sdm");
  }

  const enriched = useMemo(() => {
    return items.filter(i => i.active).map(i => { const b = brands.find(x => x.id===i.bId)||null; return Object.assign({}, i, computeItem(i,b,settings), {_b:b}); });
  }, [items, brands, settings]);

  const filtered = enriched.filter(i => {
    if (fCat !== "All" && i.cat !== fCat) return false;
    if (fBrand !== "All" && i.bId !== fBrand) return false;
    if (search && !i.name.toLowerCase().includes(search.toLowerCase()) && !(i._b && i._b.name.toLowerCase().includes(search.toLowerCase()))) return false;
    if (!isAdmin && priceFilter === "highlights" && !i.highlighted) return false;
    return true;
  });

  const grouped = useMemo(() => { const g = {}; filtered.forEach(i => { if (!g[i.cat]) g[i.cat]=[]; g[i.cat].push(i); }); return g; }, [filtered]);
  const byDate  = useMemo(() => {
    const sorted = filtered.slice().sort((a,b) => new Date(b.updatedAt||0)-new Date(a.updatedAt||0));
    const g = {}; sorted.forEach(i => { const k = dateGroupKey(i.updatedAt); if (!g[k]) g[k]=[]; g[k].push(i); }); return g;
  }, [filtered]);

  const today        = new Date().toLocaleDateString("en-IN", { day:"2-digit", month:"short", year:"numeric" });
  const views        = [{ id:"general", icon:"📂", label:"By Category" }, { id:"date", icon:"📅", label:"By Date" }];
  const activeFilter = isAdmin ? "all" : priceFilter;

  return (
    <div style={{ padding:"16px 16px 80px" }}>
      {/* Scroll buttons — admin only */}
      {isAdmin && showScroll && (
        <div style={{ position:"fixed", right:16, bottom:80, zIndex:300, display:"flex", flexDirection:"column", gap:8 }}>
          <button onClick={() => window.scrollTo({top:0,behavior:"smooth"})} style={{ width:40, height:40, borderRadius:20, background:C.navy, color:"#fff", border:"none", fontSize:18, cursor:"pointer", boxShadow:"0 3px 12px rgba(0,0,0,0.2)", display:"flex", alignItems:"center", justifyContent:"center" }}>↑</button>
          <button onClick={() => window.scrollTo({top:document.body.scrollHeight,behavior:"smooth"})} style={{ width:40, height:40, borderRadius:20, background:C.navy, color:"#fff", border:"none", fontSize:18, cursor:"pointer", boxShadow:"0 3px 12px rgba(0,0,0,0.2)", display:"flex", alignItems:"center", justifyContent:"center" }}>↓</button>
        </div>
      )}
      {/* SDM PIN Modal */}
      {showSDMPin && !isAdmin && (
        <SDMPinModal
          pin={settings.sdmPIN || "9999"}
          onSuccess={() => {
            setSdmUnlocked(true);
            setShowSDMPin(false);
            setPriceFilter("sdm");
            toast("SDM prices unlocked 🟠");
          }}
          onClose={() => setShowSDMPin(false)}
        />
      )}

      <div style={{ background:C.navy, borderRadius:13, padding:"15px", marginBottom:12, color:"#fff" }}>
        <div style={{ fontWeight:800, fontSize:17 }}>{settings.co}</div>
        <div style={{ fontSize:11, color:"rgba(255,255,255,0.5)", marginTop:2 }}>{settings.tag}</div>
        <div style={{ marginTop:8, fontSize:11, color:"rgba(255,255,255,0.35)" }}>{"Price List · " + today}</div>
      </div>

      <div style={{ display:"flex", border:"1px solid "+C.border, borderRadius:10, overflow:"hidden", marginBottom:14 }}>
        {views.map(v => (
          <button key={v.id} onClick={() => setViewMode(v.id)} style={{ flex:1, padding:"11px", border:"none", cursor:"pointer", fontWeight:700, fontSize:13, fontFamily:"inherit", display:"flex", alignItems:"center", justifyContent:"center", gap:6, background:viewMode===v.id?C.blue:"#fff", color:viewMode===v.id?"#fff":C.sec }}>
            <span>{v.icon}</span><span>{v.label}</span>
          </button>
        ))}
      </div>

      {viewMode === "date" && (
        <div style={{ display:"flex", gap:10, marginBottom:12, flexWrap:"wrap" }}>
          <Chip col={C.new_} bg={C.newBg} br={C.newBr}>🆕 New today</Chip>
          <Chip col={C.upd}  bg={C.updBg} br={C.updBr}>✏️ Updated</Chip>
        </div>
      )}

      <input style={Object.assign({},INP,{marginBottom:10})} placeholder="🔍 Search by item or brand..." value={search} onChange={e => setSearch(e.target.value)} />

      <div style={{ display:"flex", gap:8, marginBottom:12 }}>
        <select style={Object.assign({},SEL,{flex:1,padding:"9px 11px"})} value={fCat} onChange={e => setFCat(e.target.value)}>
          <option value="All">All Categories</option>{(settings.categories||CATS).map(c => <option key={c} value={c}>{c}</option>)}
        </select>
        <select style={Object.assign({},SEL,{flex:1,padding:"9px 11px"})} value={fBrand} onChange={e => setFBrand(e.target.value)}>
          <option value="All">All Brands</option>{brands.map(b => <option key={b.id} value={b.id}>{b.code+" — "+b.name}</option>)}
        </select>
      </div>

      {/* Price type filter — salesman: clickable buttons with SDM PIN lock */}
      {!isAdmin ? (
        <div style={{ marginBottom:12 }}>
          <div style={{ fontSize:11, fontWeight:700, color:C.mute, textTransform:"uppercase", letterSpacing:"0.5px", marginBottom:6 }}>Show Price Type</div>
          <div style={{ display:"flex", gap:6, flexWrap:"wrap" }}>
            {PRICE_FILTERS.map(f => {
              const on = priceFilter === f.id;
              const isSdm = f.id === "sdm";
              const label = isSdm ? (sdmUnlocked ? "🟠 SDM 🔓" : "🟠 SDM 🔒") : f.label;
              return (
                <button key={f.id}
                  onClick={isSdm ? handleSDMFilterClick : () => setPriceFilter(f.id)}
                  style={{ padding:"7px 14px", borderRadius:20, border:"1.5px solid "+(on?f.col:C.border), background:on?f.bg:"#fff", color:on?f.col:C.sec, fontWeight:700, fontSize:12, cursor:"pointer", fontFamily:"inherit" }}>
                  {label}
                </button>
              );
            })}
          </div>
          {sdmUnlocked && (
            <div style={{ marginTop:6, display:"flex", alignItems:"center", gap:8 }}>
              <div style={{ fontSize:11, color:C.sdm, fontWeight:700 }}>🔓 SDM prices unlocked for this session</div>
              <button onClick={() => { setSdmUnlocked(false); if (priceFilter === "sdm") setPriceFilter("all"); }} style={{ fontSize:11, color:C.sec, background:"none", border:"none", cursor:"pointer", padding:0, fontFamily:"inherit", textDecoration:"underline" }}>Lock again</button>
            </div>
          )}
          {priceFilter !== "all" && (
            <div style={{ fontSize:11, marginTop:4, color:PRICE_FILTERS.find(f=>f.id===priceFilter).col, fontWeight:600 }}>
              {priceFilter === "sdm" && "Prices shown Ex-GST — customer pays + GST on top"}
            </div>
          )}
        </div>
      ) : null}

      <div style={{ fontSize:12, color:C.mute, marginBottom:10, fontWeight:600 }}>{filtered.length + " items"}</div>

      {viewMode === "general" && (
        <div>
          {Object.keys(grouped).length === 0 && <div style={{ textAlign:"center", padding:"36px", color:C.mute, fontSize:14 }}>No items found.</div>}
          {Object.keys(grouped).map(cat => (
            <div key={cat}>
              <div style={{ fontSize:12, fontWeight:800, color:C.navy, marginBottom:7, marginTop:4, textTransform:"uppercase", letterSpacing:"0.4px", borderLeft:"3px solid "+C.blue, paddingLeft:9 }}>{cat}</div>
              {grouped[cat].map(it => (
                <ItemCard key={it.id} it={it} isAdmin={isAdmin} showDate={false} priceFilter={activeFilter} sdmUnlocked={isAdmin||sdmUnlocked} onSDMClick={handleSDMClick} onAddToEstimate={!isAdmin?onAddToEstimate:undefined} />
              ))}
            </div>
          ))}
        </div>
      )}

      {viewMode === "date" && (
        <div>
          {/* Pinned items — salesman only, only in default all-prices view */}
          {!isAdmin && priceFilter === "all" && (() => {
            const pinned = filtered.filter(i => i.pinned);
            if (!pinned.length) return null;
            return (
              <div style={{ marginBottom:16 }}>
                <div style={{ display:"flex", alignItems:"center", gap:10, marginBottom:10 }}>
                  <div style={{ flex:1, height:1, background:"#BFDBFE" }} />
                  <div style={{ background:C.blue, color:"#fff", borderRadius:20, padding:"4px 14px", fontSize:12, fontWeight:700, whiteSpace:"nowrap" }}>📌 Pinned Items</div>
                  <div style={{ flex:1, height:1, background:"#BFDBFE" }} />
                </div>
                {pinned.map(it => (
                  <ItemCard key={it.id} it={it} isAdmin={false} showDate={false} priceFilter="all" sdmUnlocked={sdmUnlocked} onSDMClick={handleSDMClick} onAddToEstimate={onAddToEstimate} />
                ))}
              </div>
            );
          })()}
          {Object.keys(byDate).length === 0 && <div style={{ textAlign:"center", padding:"36px", color:C.mute, fontSize:14 }}>No items found.</div>}
          {Object.keys(byDate).map(day => (
            <div key={day}>
              <div style={{ display:"flex", alignItems:"center", gap:10, marginBottom:10, marginTop:8 }}>
                <div style={{ flex:1, height:1, background:C.border }} />
                <div style={{ background:C.navy, color:"#fff", borderRadius:20, padding:"4px 14px", fontSize:12, fontWeight:700, whiteSpace:"nowrap" }}>{day}</div>
                <div style={{ flex:1, height:1, background:C.border }} />
              </div>
              {byDate[day].map(it => (
                <ItemCard key={it.id} it={it} isAdmin={isAdmin} showDate priceFilter={activeFilter} sdmUnlocked={isAdmin||sdmUnlocked} onSDMClick={handleSDMClick} onAddToEstimate={!isAdmin?onAddToEstimate:undefined} />
              ))}
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

// ── DASHBOARD VIEW ────────────────────────────────────────────────
function DashboardView({ brands, items, settings, onSettingsChange, estimates, onEstimatesSave }) {
  const [fBrand, setFBrand] = useState("All");
  const [showEst, setShowEst] = useState(false);
  const PTYPES = [
    { id:"rl",  label:"RL",  col:C.rl,  bg:C.rlBg,  br:C.rlBr },
    { id:"dm",  label:"DM",  col:C.dm,  bg:C.dmBg,  br:C.dmBr },
    { id:"pl",  label:"PL",  col:C.pl,  bg:C.plBg,  br:C.plBr },
    { id:"sdm", label:"SDM", col:C.sdm, bg:C.sdmBg, br:C.sdmBr },
  ];

  const enriched = useMemo(() => {
    return items.filter(i => i.active).map(i => {
      const b = brands.find(x => x.id===i.bId)||null;
      return Object.assign({}, i, computeItem(i,b,settings), {_b:b});
    });
  }, [items, brands, settings]);

  const filtered = fBrand === "All" ? enriched : enriched.filter(i => i.bId === fBrand);

  function avgMargin(arr, key) {
    const valid = arr.filter(i => i[key] && i[key+"Profit"] && i[key+"Profit"].pct != null);
    if (!valid.length) return null;
    return +(valid.reduce((s,i) => s + i[key+"Profit"].pct, 0) / valid.length).toFixed(1);
  }
  function topItems(arr, key, n, asc) {
    return arr.filter(i => i[key] && i[key+"Profit"] && i[key+"Profit"].pct != null)
      .sort((a,b) => asc ? a[key+"Profit"].pct - b[key+"Profit"].pct : b[key+"Profit"].pct - a[key+"Profit"].pct)
      .slice(0, n);
  }

  const SS = { background:C.card, border:"1px solid "+C.border, borderRadius:13, padding:"16px", marginBottom:13, boxShadow:"0 1px 6px rgba(0,0,0,0.04)" };

  function MarginBar({ pct, col }) {
    if (pct == null) return <span style={{ fontSize:11, color:C.mute }}>—</span>;
    const w = Math.min(Math.abs(pct), 60);
    return (
      <div style={{ display:"flex", alignItems:"center", gap:8 }}>
        <div style={{ flex:1, background:"#F3F4F6", borderRadius:20, height:8, overflow:"hidden" }}>
          <div style={{ width:w+"%", background:pct<0?C.red:col, height:"100%", borderRadius:20, transition:"width 0.3s" }} />
        </div>
        <span style={{ fontSize:12, fontWeight:800, color:pct<0?C.red:col, minWidth:42, textAlign:"right" }}>{pct}%</span>
      </div>
    );
  }

  return (
    <div style={{ padding:"16px 16px 80px" }}>
      {showEst && <EstimateHistory estimates={estimates} onEstimatesSave={onEstimatesSave} isAdmin={true} onClose={() => setShowEst(false)} settings={settings} />}

      {/* Today's estimates summary */}
      {(() => {
        const todayEst = (estimates||[]).filter(e => new Date(e.createdAt).toDateString()===new Date().toDateString());
        const todayTotal = todayEst.reduce((s,e)=>s+e.grandTotal,0);
        return (
          <div style={{ background:C.dmBg, border:"1px solid "+C.dmBr, borderRadius:12, padding:"14px", marginBottom:14 }}>
            <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:todayEst.length>0?10:0 }}>
              <div>
                <div style={{ fontSize:13, fontWeight:800, color:C.dm }}>📋 Today's Estimates</div>
                <div style={{ fontSize:11, color:C.dm, marginTop:2 }}>{todayEst.length} estimate{todayEst.length!==1?"s":""} · {fp(todayTotal)}</div>
              </div>
              <button onClick={() => setShowEst(true)} style={{ padding:"7px 14px", borderRadius:8, border:"1.5px solid "+C.dm, background:"#fff", color:C.dm, fontWeight:700, fontSize:12, cursor:"pointer", fontFamily:"inherit" }}>View All</button>
            </div>
            {todayEst.length > 0 && (
              <div>
                {todayEst.slice(-3).reverse().map(e => (
                  <div key={e.id} style={{ display:"flex", justifyContent:"space-between", fontSize:12, padding:"5px 0", borderBottom:"1px solid "+C.dmBr }}>
                    <span style={{ fontWeight:700, color:C.dm }}>{e.number}</span>
                    <span style={{ color:C.sec }}>{e.salesmanName}{e.custName?" · "+e.custName:""}</span>
                    <span style={{ fontWeight:800, color:C.navy }}>{fp(e.grandTotal)}</span>
                  </div>
                ))}
              </div>
            )}
          </div>
        );
      })()}

      {/* Brand margin filter */}
      <div style={{ marginBottom:14 }}>
        <select style={Object.assign({},SEL,{marginBottom:0})} value={fBrand} onChange={e => setFBrand(e.target.value)}>
          <option value="All">All Brands</option>
          {brands.map(b => <option key={b.id} value={b.id}>{b.code+" — "+b.name}</option>)}
        </select>
      </div>

      {/* Summary cards */}
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:10, marginBottom:14 }}>
        {PTYPES.map(pt => {
          const key = pt.id === "sdm" ? "sdm" : pt.id;
          const profKey = key + "Profit";
          const avg = avgMargin(filtered, key);
          return (
            <div key={pt.id} style={{ background:pt.bg, border:"1.5px solid "+pt.br, borderRadius:12, padding:"14px 12px", textAlign:"center" }}>
              <div style={{ fontSize:10, fontWeight:800, color:pt.col, textTransform:"uppercase", letterSpacing:"0.5px", marginBottom:6 }}>{pt.label} Avg Margin</div>
              <div style={{ fontSize:26, fontWeight:900, color:pt.col }}>{avg != null ? avg+"%" : "—"}</div>
              <div style={{ fontSize:10, color:pt.col, marginTop:4, opacity:0.7 }}>
                {filtered.filter(i => i[key] && i[profKey]).length} items
              </div>
            </div>
          );
        })}
      </div>

      {/* Per price type details */}
      {PTYPES.map(pt => {
        const key = pt.id;
        const profKey = key + "Profit";
        const best = topItems(filtered, key, 5, false);
        const worst = topItems(filtered, key, 5, true);
        if (!best.length) return null;
        return (
          <div key={pt.id} style={SS}>
            <div style={{ fontSize:13, fontWeight:800, color:pt.col, marginBottom:12 }}>{pt.label} — Margin Breakdown</div>
            <div style={{ fontSize:11, fontWeight:700, color:C.profit, textTransform:"uppercase", letterSpacing:"0.4px", marginBottom:6 }}>Top 5 Highest</div>
            {best.map((it,i) => (
              <div key={it.id} style={{ marginBottom:7 }}>
                <div style={{ fontSize:12, fontWeight:600, color:C.text, marginBottom:3, display:"flex", justifyContent:"space-between" }}>
                  <span>{it.name}</span>
                  <span style={{ fontSize:11, color:C.mute }}>{it._b?it._b.code:""}</span>
                </div>
                <MarginBar pct={it[profKey]?it[profKey].pct:null} col={pt.col} />
              </div>
            ))}
            <div style={{ fontSize:11, fontWeight:700, color:C.red, textTransform:"uppercase", letterSpacing:"0.4px", marginBottom:6, marginTop:14 }}>Bottom 5 — Watch these</div>
            {worst.map((it,i) => (
              <div key={it.id} style={{ marginBottom:7 }}>
                <div style={{ fontSize:12, fontWeight:600, color:C.text, marginBottom:3, display:"flex", justifyContent:"space-between" }}>
                  <span>{it.name}</span>
                  <span style={{ fontSize:11, color:C.mute }}>{it._b?it._b.code:""}</span>
                </div>
                <MarginBar pct={it[profKey]?it[profKey].pct:null} col={pt.col} />
              </div>
            ))}
          </div>
        );
      })}

      {/* Brand-wise summary */}
      <div style={SS}>
        <div style={{ fontSize:13, fontWeight:800, marginBottom:12 }}>Brand-wise Average RL Margin</div>
        {brands.map(b => {
          const bItems = enriched.filter(i => i.bId === b.id && i.rl && i.rlProfit);
          if (!bItems.length) return null;
          const avg = +(bItems.reduce((s,i)=>s+i.rlProfit.pct,0)/bItems.length).toFixed(1);
          return (
            <div key={b.id} style={{ marginBottom:10 }}>
              <div style={{ fontSize:12, fontWeight:700, color:C.text, marginBottom:4, display:"flex", justifyContent:"space-between" }}>
                <span>{b.name}</span><span style={{ fontSize:11, color:C.mute }}>{bItems.length} items</span>
              </div>
              <div style={{ display:"flex", alignItems:"center", gap:8 }}>
                <div style={{ flex:1, background:"#F3F4F6", borderRadius:20, height:8, overflow:"hidden" }}>
                  <div style={{ width:Math.min(Math.abs(avg),60)+"%", background:C.rl, height:"100%", borderRadius:20 }} />
                </div>
                <span style={{ fontSize:12, fontWeight:800, color:C.rl, minWidth:42, textAlign:"right" }}>{avg}%</span>
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

// ── SETTINGS VIEW ─────────────────────────────────────────────────
function SalesmanAdder({ settings, onSettingsChange }) {
  const [val, setVal] = useState("");
  return (
    <div style={{ display:"flex", gap:8, marginTop:8 }}>
      <input style={Object.assign({},INP,{flex:1})} placeholder="Add salesman name..." value={val} onChange={e=>setVal(e.target.value)}
        onKeyDown={e=>{ if(e.key==="Enter"){ const v=val.trim(); if(!v)return; const cur=settings.salesmen||[]; if(cur.map(s=>s.toLowerCase()).includes(v.toLowerCase()))return toast("Already exists","warn"); onSettingsChange({...settings,salesmen:[...cur,v]}); setVal(""); toast("Salesman added"); } }} />
      <button onClick={()=>{ const v=val.trim(); if(!v)return; const cur=settings.salesmen||[]; if(cur.map(s=>s.toLowerCase()).includes(v.toLowerCase()))return toast("Already exists","warn"); onSettingsChange({...settings,salesmen:[...cur,v]}); setVal(""); toast("Salesman added"); }}
        style={{ padding:"11px 16px", borderRadius:9, border:"none", background:C.blue, color:"#fff", fontWeight:700, fontSize:13, cursor:"pointer", fontFamily:"inherit", whiteSpace:"nowrap" }}>+ Add</button>
    </div>
  );
}

function SettingsView({ settings, onSettingsChange, syncStatus }) {
  const [form, setForm] = useState({
    co:settings.co, tag:settings.tag,
    defaultRL:+(settings.defaultRL*100).toFixed(1),
    defaultDM:settings.defaultDM!=null?+(parseFloat(settings.defaultDM)*100).toFixed(1):"",
  });
  useEffect(() => {
    setForm({ co:settings.co, tag:settings.tag, defaultRL:+(settings.defaultRL*100).toFixed(1), defaultDM:settings.defaultDM!=null?+(parseFloat(settings.defaultDM)*100).toFixed(1):"" });
  }, [settings]);
  const [nAP,setNAP]=useState(""); const [cAP,setCAP]=useState("");
  const [nSP,setNSP]=useState(""); const [cSP,setCSP]=useState("");
  const [nSDM,setNSDM]=useState(""); const [cSDM,setCSDM]=useState("");
  const [newCat, setNewCat] = useState("");

  function saveGen() {
    if (!form.co.trim()) return toast("Company name required","err");
    if (isNaN(parseFloat(form.defaultRL))) return toast("RL markup must be a number","err");
    onSettingsChange({ ...settings, co:form.co.trim(), tag:form.tag.trim(), defaultRL:parseFloat(form.defaultRL)/100, defaultDM:form.defaultDM!==""&&!isNaN(parseFloat(form.defaultDM))?parseFloat(form.defaultDM)/100:null });
    toast("Settings saved");
  }
  function changePIN(type) {
    const np = type==="admin"?nAP:type==="salesman"?nSP:nSDM;
    const cp = type==="admin"?cAP:type==="salesman"?cSP:cSDM;
    if (np.length!==4||isNaN(parseInt(np))) return toast("PIN must be exactly 4 digits","err");
    if (np!==cp) return toast("PINs do not match","err");
    const key = type==="admin"?"adminPIN":type==="salesman"?"salesPIN":"sdmPIN";
    onSettingsChange({ ...settings, [key]:np });
    if (type==="admin"){setNAP("");setCAP("");}
    else if (type==="salesman"){setNSP("");setCSP("");}
    else {setNSDM("");setCSDM("");}
    const label = type==="admin"?"Admin":type==="salesman"?"Salesman":"SDM";
    toast(label + " PIN updated");
  }

  const SS = { background:C.card, border:"1px solid "+C.border, borderRadius:13, padding:"17px", marginBottom:13, boxShadow:"0 1px 6px rgba(0,0,0,0.04)" };
  const pinBlocks = [
    { type:"admin",    label:"Admin",    col:C.navy, nv:nAP, setNV:setNAP, cv:cAP, setCV:setCAP },
    { type:"salesman", label:"Salesman", col:C.blue, nv:nSP, setNV:setNSP, cv:cSP, setCV:setCSP },
    { type:"sdm",      label:"SDM View", col:C.sdm,  nv:nSDM,setNV:setNSDM,cv:cSDM,setCV:setCSDM },
  ];

  return (
    <div style={{ padding:"16px 16px 80px" }}>
      <div style={Object.assign({},SS,{display:"flex",alignItems:"center",justifyContent:"space-between"})}>
        <div><div style={{ fontSize:13, fontWeight:800, marginBottom:4 }}>☁️ Cloud Sync</div><div style={{ fontSize:11, color:C.mute }}>All devices share one database</div></div>
        <SyncBadge status={syncStatus} />
      </div>
      <div style={SS}>
        <div style={{ fontSize:14, fontWeight:800, marginBottom:14 }}>General Settings</div>
        <Fld label="Company Name"><input style={INP} value={form.co} onChange={e => setForm(p => ({...p,co:e.target.value}))} /></Fld>
        <Fld label="Tagline"><input style={INP} value={form.tag} onChange={e => setForm(p => ({...p,tag:e.target.value}))} /></Fld>
        <Fld label="Default RL Markup %" hint="18 = 18%"><input style={INP} type="number" step="0.5" value={form.defaultRL} onChange={e => setForm(p => ({...p,defaultRL:e.target.value}))} /></Fld>
        <Fld label="Default DM Markup %" hint="Leave blank if not decided."><input style={INP} type="number" step="0.5" placeholder="Leave blank" value={form.defaultDM} onChange={e => setForm(p => ({...p,defaultDM:e.target.value}))} /></Fld>
        <BtnP onClick={saveGen}>Save Settings</BtnP>
      </div>
      {/* Categories */}
      <div style={SS}>
        <div style={{ fontSize:14, fontWeight:800, marginBottom:4 }}>Categories</div>
        <div style={{ fontSize:11, color:C.mute, marginBottom:14 }}>Add or remove product categories. Used in all item dropdowns.</div>
        {(settings.categories||CATS).map((cat, i) => (
          <div key={cat} style={{ display:"flex", alignItems:"center", justifyContent:"space-between", padding:"9px 12px", background:"#F9FAFB", border:"1px solid "+C.border, borderRadius:8, marginBottom:7 }}>
            <span style={{ fontSize:13, fontWeight:600, color:C.text }}>{cat}</span>
            <button onClick={() => {
              const next = (settings.categories||CATS).filter(c => c !== cat);
              if (next.length === 0) return toast("Must keep at least one category","warn");
              onSettingsChange({ ...settings, categories:next });
              toast("Category removed","warn");
            }} style={{ width:28, height:28, borderRadius:6, border:"1.5px solid #FCA5A5", background:C.redBg, cursor:"pointer", fontSize:13, color:C.red, fontFamily:"inherit" }}>✕</button>
          </div>
        ))}
        <div style={{ display:"flex", gap:8, marginTop:8 }}>
          <input style={Object.assign({},INP,{flex:1})} placeholder="New category name..." value={newCat} onChange={e => setNewCat(e.target.value)} />
          <button onClick={() => {
            const v = newCat.trim();
            if (!v) return;
            const cur = settings.categories||CATS;
            if (cur.map(c=>c.toLowerCase()).includes(v.toLowerCase())) return toast("Category already exists","warn");
            onSettingsChange({ ...settings, categories:[...cur, v] });
            setNewCat("");
            toast("Category added");
          }} style={{ padding:"11px 16px", borderRadius:9, border:"none", background:C.blue, color:"#fff", fontWeight:700, fontSize:13, cursor:"pointer", fontFamily:"inherit", whiteSpace:"nowrap" }}>+ Add</button>
        </div>
      </div>
      {pinBlocks.map(pt => (
        <div key={pt.type} style={SS}>
          <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:14 }}>
            <div style={{ width:10, height:10, borderRadius:"50%", background:pt.col }} />
            <div style={{ fontSize:14, fontWeight:800 }}>{"Change " + pt.label + " PIN"}</div>
          </div>
          {pt.type === "sdm" && (
            <div style={{ background:C.sdmBg, border:"1px solid "+C.sdmBr, borderRadius:8, padding:"8px 12px", marginBottom:12, fontSize:12, color:C.sdm, fontWeight:600 }}>
              🟠 This PIN is required by salesmen to view SDM prices. Keep it confidential.
            </div>
          )}
          <Fld label="New PIN (4 digits)">
            <input style={INP} type="password" maxLength={4} placeholder="4-digit PIN"
              value={pt.nv} onChange={e => { const v=e.target.value.replace(/\D/g,"").slice(0,4); pt.setNV(v); }} />
          </Fld>
          <Fld label="Confirm PIN">
            <input style={INP} type="password" maxLength={4} placeholder="Re-enter to confirm"
              value={pt.cv} onChange={e => { const v=e.target.value.replace(/\D/g,"").slice(0,4); pt.setCV(v); }} />
          </Fld>
          <BtnP color={pt.col} onClick={() => changePIN(pt.type)}>{"Update " + pt.label + " PIN"}</BtnP>
        </div>
      ))}
      {/* Salesman Management */}
      <div style={SS}>
        <div style={{ fontSize:14, fontWeight:800, marginBottom:14 }}>👥 Salesmen</div>
        <div style={{ fontSize:11, color:C.mute, marginBottom:10 }}>Salesmen listed here appear in the estimate dropdown.</div>
        {(settings.salesmen||[]).map(s => (
          <div key={s} style={{ display:"flex", alignItems:"center", justifyContent:"space-between", padding:"9px 12px", background:"#F9FAFB", border:"1px solid "+C.border, borderRadius:8, marginBottom:7 }}>
            <span style={{ fontSize:13, fontWeight:600 }}>{s}</span>
            <button onClick={() => { const next=(settings.salesmen||[]).filter(x=>x!==s); if(next.length===0)return toast("Keep at least one salesman","warn"); onSettingsChange({...settings,salesmen:next}); toast("Removed","warn"); }}
              style={{ width:28, height:28, borderRadius:6, border:"1.5px solid #FCA5A5", background:C.redBg, cursor:"pointer", fontSize:13, color:C.red, fontFamily:"inherit" }}>✕</button>
          </div>
        ))}
        <SalesmanAdder settings={settings} onSettingsChange={onSettingsChange} />
      </div>
      <div style={Object.assign({},SS,{textAlign:"center"})}>
        <div style={{ fontSize:13, color:C.sec, lineHeight:1.7 }}>{"v" + VER}</div>
        <div style={{ fontSize:12, color:C.mute, marginTop:4 }}>Data syncs automatically across all devices</div>
      </div>
    </div>
  );
}

// ── ADMIN APP ─────────────────────────────────────────────────────
function AdminApp({ settings, onSettingsChange, brands, onBrandsChange, items, onItemsChange, onLogout, syncStatus, estimates, onEstimatesSave }) {
  const [tab, setTab] = useState("master");
  const NAV = [{ id:"master",icon:"📋",label:"Master" },{ id:"dashboard",icon:"📊",label:"Dashboard" },{ id:"brands",icon:"🏷",label:"Brands" },{ id:"settings",icon:"⚙️",label:"Settings" }];
  return (
    <div style={{ minHeight:"100vh", background:C.bg }}>
      <div style={{ background:C.navy, padding:"13px 16px", display:"flex", alignItems:"center", justifyContent:"space-between", position:"sticky", top:0, zIndex:100, boxShadow:"0 2px 10px rgba(0,0,0,0.2)" }}>
        <div>
          <div style={{ color:"#fff", fontWeight:800, fontSize:15 }}>{settings.co}</div>
          <div style={{ color:"rgba(255,255,255,0.4)", fontSize:11, display:"flex", alignItems:"center", gap:8 }}>Admin 🔐 <SyncBadge status={syncStatus} /></div>
        </div>
        <button onClick={onLogout} style={{ padding:"7px 13px", borderRadius:8, border:"1.5px solid rgba(255,255,255,0.22)", background:"rgba(255,255,255,0.08)", color:"#fff", fontSize:12, fontWeight:600, cursor:"pointer", fontFamily:"inherit" }}>Logout</button>
      </div>
      {tab === "master"    && <MasterView    brands={brands} items={items} onItemsChange={onItemsChange} settings={settings} />}
      {tab === "dashboard" && <DashboardView brands={brands} items={items} settings={settings} onSettingsChange={onSettingsChange} estimates={estimates} onEstimatesSave={onEstimatesSave} />}
      {tab === "brands"    && <BrandsView    brands={brands} onBrandsChange={onBrandsChange} />}
      {tab === "settings"  && <SettingsView  settings={settings} onSettingsChange={onSettingsChange} syncStatus={syncStatus} />}
      <div style={{ position:"fixed", bottom:0, left:0, right:0, zIndex:200, background:"#fff", borderTop:"1px solid "+C.border, display:"flex", height:62, boxShadow:"0 -3px 16px rgba(0,0,0,0.07)" }}>
        {NAV.map(t => (
          <button key={t.id} onClick={() => setTab(t.id)} style={{ flex:1, display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center", gap:2, border:"none", background:"transparent", cursor:"pointer", fontFamily:"inherit", borderTop:tab===t.id?"2.5px solid "+C.blue:"2.5px solid transparent" }}>
            <span style={{ fontSize:19 }}>{t.icon}</span><span style={{ fontSize:10, fontWeight:700, color:tab===t.id?C.blue:C.mute }}>{t.label}</span>
          </button>
        ))}
      </div>
    </div>
  );
}

// ── SALESMAN APP ──────────────────────────────────────────────────
function SalesmanApp({ settings, brands, items, onLogout, syncStatus, estimates, onEstimatesSave }) {
  const [tab, setTab] = useState("prices");
  const [addPopupItem, setAddPopupItem] = useState(null); // item to add from prices tab

  const NAV = [
    { id:"prices",   icon:"🏷", label:"Prices" },
    { id:"estimate", icon:"🧾", label:"Estimate" },
  ];

  return (
    <div style={{ minHeight:"100vh", background:C.bg }}>
      <div style={{ background:C.navy, padding:"13px 16px", display:"flex", alignItems:"center", justifyContent:"space-between", position:"sticky", top:0, zIndex:100, boxShadow:"0 2px 10px rgba(0,0,0,0.2)" }}>
        <div>
          <div style={{ color:"#fff", fontWeight:800, fontSize:15 }}>{settings.co}</div>
          <div style={{ color:"rgba(255,255,255,0.4)", fontSize:11, display:"flex", alignItems:"center", gap:8 }}>Price List 📋 <SyncBadge status={syncStatus} /></div>
        </div>
        <button onClick={onLogout} style={{ padding:"7px 13px", borderRadius:8, border:"1.5px solid rgba(255,255,255,0.22)", background:"rgba(255,255,255,0.08)", color:"#fff", fontSize:12, fontWeight:600, cursor:"pointer", fontFamily:"inherit" }}>Logout</button>
      </div>
      {/* Price-to-estimate popup */}
      {addPopupItem && (
        <div style={{ position:"fixed", inset:0, zIndex:700, background:"rgba(0,0,0,0.5)", display:"flex", alignItems:"center", justifyContent:"center", padding:20 }}>
          <div style={{ background:"#fff", borderRadius:20, padding:"22px 18px", maxWidth:340, width:"100%" }}>
            <div style={{ fontSize:15, fontWeight:800, marginBottom:4 }}>{addPopupItem.name}</div>
            <div style={{ fontSize:11, color:C.sec, marginBottom:14 }}>Select price to add to estimate</div>
            {PRICE_TYPES.map(pt => {
              const prices = { rl:addPopupItem.rl, dm:addPopupItem.dm, pl:addPopupItem.pl, sdm:addPopupItem.sdmInc||addPopupItem.sdm };
              const pv = prices[pt.id];
              if (!pv) return null;
              return (
                <button key={pt.id} onClick={() => {
                  // pass to estimate tab via localStorage-like state — we navigate to estimate tab
                  setTab("estimate");
                  // We store the pending item in a ref so EstimateView can pick it up
                  window._pendingEstimateItem = { item: addPopupItem, priceType: pt.id };
                  setAddPopupItem(null);
                  toast(addPopupItem.name + " → Estimate");
                }}
                  style={{ display:"flex", alignItems:"center", justifyContent:"space-between", width:"100%", padding:"13px 16px", borderRadius:10, border:"1.5px solid "+pt.col, background:pt.bg, cursor:"pointer", fontFamily:"inherit", marginBottom:8 }}>
                  <span style={{ fontWeight:700, color:pt.col, fontSize:14 }}>{pt.label}</span>
                  <span style={{ fontWeight:900, color:pt.col, fontSize:16 }}>{fp(pv)}</span>
                </button>
              );
            })}
            <BtnO onClick={() => setAddPopupItem(null)} style={{ marginTop:4 }}>Cancel</BtnO>
          </div>
        </div>
      )}
      {tab === "prices"   && <PriceListView brands={brands} items={items} settings={settings} isAdmin={false} onAddToEstimate={it => setAddPopupItem(it)} />}
      {tab === "estimate" && <EstimateView  brands={brands} items={items} settings={settings} estimates={estimates} onEstimatesSave={onEstimatesSave} isAdmin={false} />}
      <div style={{ position:"fixed", bottom:0, left:0, right:0, zIndex:200, background:"#fff", borderTop:"1px solid "+C.border, display:"flex", height:62, boxShadow:"0 -3px 16px rgba(0,0,0,0.07)" }}>
        {NAV.map(t => (
          <button key={t.id} onClick={() => setTab(t.id)} style={{ flex:1, display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center", gap:2, border:"none", background:"transparent", cursor:"pointer", fontFamily:"inherit", borderTop:tab===t.id?"2.5px solid "+C.blue:"2.5px solid transparent" }}>
            <span style={{ fontSize:19 }}>{t.icon}</span><span style={{ fontSize:10, fontWeight:700, color:tab===t.id?C.blue:C.mute }}>{t.label}</span>
          </button>
        ))}
      </div>
    </div>
  );
}

// ── ROOT ──────────────────────────────────────────────────────────
export default function App() {
  const [settings,   setSettings]   = useState(DEF_S);
  const [brands,     setBrands]     = useState(DEF_B);
  const [items,      setItems]      = useState(DEF_I);
  const [estimates,  setEstimates]  = useState([]);
  const [role,       setRole]       = useState(null);
  const [ready,      setReady]      = useState(false);
  const [syncStatus, setSyncStatus] = useState("synced");

  // Inject global styles once
  useEffect(() => {
    if (!document.getElementById("vkf-global-css")) {
      const s = document.createElement("style");
      s.id = "vkf-global-css";
      s.textContent = [
        "*{box-sizing:border-box;-webkit-tap-highlight-color:transparent;}",
        "input:focus,select:focus{outline:none!important;border-color:#1648D6!important;box-shadow:0 0 0 3px rgba(22,72,214,0.1)!important;}",
        "input[type=number]::-webkit-inner-spin-button{-webkit-appearance:none;}",
        "input::placeholder{color:#CBD5E1;}",
        "button:active{opacity:0.75;transform:scale(0.97);}",
        "::-webkit-scrollbar{width:4px;}",
        "::-webkit-scrollbar-thumb{background:#D1D5DB;border-radius:4px;}",
        "@keyframes shk{0%,100%{transform:translateX(0)}25%{transform:translateX(-5px)}75%{transform:translateX(5px)}}",
        "@media print{.no-print{display:none!important;}}",
      ].join("");
      document.head.appendChild(s);
    }
  }, []);

  useEffect(() => {
    fsLoad().then(d => {
      if (d.settings)  setSettings(s => Object.assign({}, DEF_S, d.settings));
      if (d.brands)    setBrands(d.brands);
      if (d.items)     setItems(d.items);
      if (d.estimates) setEstimates(d.estimates);
      setReady(true);
    }).catch(() => { setSyncStatus("offline"); setReady(true); });
  }, []);

  const listenersStarted = useRef(false);
  useEffect(() => {
    if (!ready || listenersStarted.current) return;
    listenersStarted.current = true;
    const unsubS = onSnapshot(FS.settings(),  snap => { if (snap.exists()) setSettings(s => Object.assign({}, DEF_S, snap.data().v)); }, () => setSyncStatus("offline"));
    const unsubB = onSnapshot(FS.brands(),    snap => { if (snap.exists()) setBrands(snap.data().v);    }, () => setSyncStatus("offline"));
    const unsubI = onSnapshot(FS.items(),     snap => { if (snap.exists()) setItems(snap.data().v);     }, () => setSyncStatus("offline"));
    const unsubE = onSnapshot(FS.estimates(), snap => { if (snap.exists()) setEstimates(snap.data().v); }, () => {});
    return () => { unsubS(); unsubB(); unsubI(); unsubE(); };
  }, [ready]);

  async function saveSettings(next)  { setSettings(next);  setSyncStatus("saving"); try { await fsSave("settings",next);  setSyncStatus("synced"); } catch { setSyncStatus("error"); toast("Sync failed","err"); } }
  async function saveBrands(next)    { setBrands(next);    setSyncStatus("saving"); try { await fsSave("brands",next);    setSyncStatus("synced"); } catch { setSyncStatus("error"); toast("Sync failed","err"); } }
  async function saveItems(next)     { setItems(next);     setSyncStatus("saving"); try { await fsSave("items",next);     setSyncStatus("synced"); } catch { setSyncStatus("error"); toast("Sync failed","err"); } }
  async function saveEstimates(next) { setEstimates(next); try { await fsSave("estimates",next); } catch { toast("Estimate sync failed","err"); } }

  if (!ready) return (
    <div style={{ minHeight:"100vh", background:C.bg, display:"flex", alignItems:"center", justifyContent:"center" }}>
      <div style={{ textAlign:"center" }}>
        <div style={{ width:52, height:52, borderRadius:14, background:C.navy, display:"flex", alignItems:"center", justifyContent:"center", color:"#fff", fontWeight:900, fontSize:20, margin:"0 auto 10px" }}>VK</div>
        <div style={{ color:C.sec, fontSize:13 }}>Loading…</div>
      </div>
    </div>
  );

  return (
    <div>

      <ToastHost />
      <div style={{ fontFamily:"system-ui,-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif" }}>
        {!role               && <Login       settings={settings} onLogin={setRole} />}
        {role === "admin"    && <AdminApp    settings={settings} onSettingsChange={saveSettings} brands={brands} onBrandsChange={saveBrands} items={items} onItemsChange={saveItems} onLogout={() => setRole(null)} syncStatus={syncStatus} estimates={estimates} onEstimatesSave={saveEstimates} />}
        {role === "salesman" && <SalesmanApp settings={settings} brands={brands} items={items} onLogout={() => setRole(null)} syncStatus={syncStatus} estimates={estimates} onEstimatesSave={saveEstimates} />}
      </div>
    </div>
  );
}
