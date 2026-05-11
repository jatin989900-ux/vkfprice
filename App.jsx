// ─────────────────────────────────────────────────────────────────
//  VK Furnishing Price App  v4.0  — Firebase Edition
// ─────────────────────────────────────────────────────────────────

import { useState, useEffect, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import { initializeApp } from "firebase/app";
import { getFirestore, doc, getDoc, setDoc, onSnapshot } from "firebase/firestore";

// Your web app's Firebase configuration
const firebaseConfig = {
  apiKey: "AIzaSyCEJlH9aGlr0pneb7hT1sIxy1iDnQV3Y4g",
  authDomain: "vkf-price.firebaseapp.com",
  databaseURL: "https://vkf-price-default-rtdb.asia-southeast1.firebasedatabase.app",
  projectId: "vkf-price",
  storageBucket: "vkf-price.firebasestorage.app",
  messagingSenderId: "199170851796",
  appId: "1:199170851796:web:e2e46e279995a41b890c41"
};

// Initialize Firebase once
const app = initializeApp(firebaseConfig);
const db = getFirestore(app);

// Firestore document paths  (all data lives in one doc per collection)
const FS = {
  settings: () => doc(db, "vkf", "settings"),
  brands:   () => doc(db, "vkf", "brands"),
  items:    () => doc(db, "vkf", "items"),
};

// ── FIREBASE STORAGE LAYER ────────────────────────────────────────
async function fsLoad() {
  try {
    const [sSnap, bSnap, iSnap] = await Promise.all([
      getDoc(FS.settings()),
      getDoc(FS.brands()),
      getDoc(FS.items()),
    ]);
    return {
      settings: sSnap.exists() ? sSnap.data().v : null,
      brands:   bSnap.exists() ? bSnap.data().v : null,
      items:    iSnap.exists() ? iSnap.data().v : null,
    };
  } catch (e) {
    console.error("Firebase load error:", e);
    return { settings: null, brands: null, items: null };
  }
}

async function fsSave(key, value) {
  try {
    await setDoc(FS[key](), { v: value, updatedAt: new Date().toISOString() });
  } catch (e) {
    console.error("Firebase save error:", e);
    throw e;
  }
}

// ── CONSTANTS ─────────────────────────────────────────────────────
const VER  = "4.0";
const CATS = ["Bed Sheets","Comforters","Comforter Sets","Towels","Pillows","Dohars","Blankets","Top Sheets","Other Items"];
const GST_OPTS  = [{ label:"5% — Default (Textiles)", v:0.05 },{ label:"18%", v:0.18 }];
const ADDON_OPTS = [
  { label:"Not set",   v:"" },
  { label:"+ ₹100",   v:100 },
  { label:"+ ₹200",   v:200 },
  { label:"+ ₹300",   v:300 },
  { label:"+ ₹500",   v:500 },
  { label:"+ ₹1,000", v:1000 },
  { label:"Custom ₹", v:"custom" },
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
function calcRL(ex, m) { const v=parseFloat(ex||0); return v>0?psychRound(v*(1+parseFloat(m||0.18))):null; }
function calcDM(ex, m) { const v=parseFloat(ex||0); if(m==null||m===""||isNaN(parseFloat(m)))return null; return v>0?psychRound(v*(1+parseFloat(m))):null; }
function calcPL(rl, a) { return (rl!=null&&typeof a==="number"&&!isNaN(a))?psychRound(rl+a):null; }
function calcAddon(pla, ca) {
  if (pla===""||pla==null) return null;
  if (pla==="custom") { const v=parseFloat(ca); return !isNaN(v)&&v>0?v:null; }
  return typeof pla==="number"?pla:null;
}
function calcProfit(sell, ex) {
  if (sell==null||!ex||parseFloat(ex)<=0) return null;
  const amt = sell - parseFloat(ex);
  return { amt:+amt.toFixed(2), pct:+((amt/parseFloat(ex))*100).toFixed(1) };
}
function resolveRL(item, brand, settings) {
  const v = parseFloat(item.customRL);
  if (item.customRL!==""&&item.customRL!=null&&!isNaN(v)) return v/100;
  if (brand?.rlMarkup!=null) return brand.rlMarkup;
  return parseFloat(settings.defaultRL||0.18);
}
function resolveDM(item, brand, settings) {
  const v = parseFloat(item.customDM);
  if (item.customDM!==""&&item.customDM!=null&&!isNaN(v)) return v/100;
  if (brand?.dmMarkup!=null) return brand.dmMarkup;
  if (settings.defaultDM!=null&&settings.defaultDM!=="") return parseFloat(settings.defaultDM);
  return null;
}
function computeItem(item, brand, settings) {
  const rlM = resolveRL(item, brand, settings);
  const dmM = resolveDM(item, brand, settings);
  const add = calcAddon(item.plAddon, item.customAddon);
  const ex  = parseFloat(item.purchaseEx||0);
  const rl  = calcRL(ex, rlM);
  const dm  = calcDM(ex, dmM);
  const pl  = calcPL(rl, add);
  return {
    incGST: calcIncGST(ex, item.gst),
    rlM, rl, rlProfit: calcProfit(rl, ex),
    dmM, dm, dmProfit: calcProfit(dm, ex),
    add, pl, plProfit: calcProfit(pl, ex),
  };
}

// ── DEFAULTS ──────────────────────────────────────────────────────
const DEF_S = { co:"VK Furnishing", tag:"Wholesale Bedding & Textiles, Delhi NCR", defaultRL:0.18, defaultDM:null, adminPIN:"1234", salesPIN:"0000" };
const DEF_B = [
  { id:"b1", code:"TRI", name:"Trident",      rlMarkup:0.22, dmMarkup:null },
  { id:"b2", code:"STH", name:"Story@Home",   rlMarkup:0.18, dmMarkup:null },
  { id:"b3", code:"SWT", name:"Swayam",       rlMarkup:0.20, dmMarkup:null },
  { id:"b4", code:"LOC", name:"Local / Misc", rlMarkup:0.18, dmMarkup:null },
];
const T0 = new Date().toISOString();
const DEF_I = [
  { id:"i1", cat:"Bed Sheets", bId:"b1", name:"King Cotton Bedsheet 200TC",    purchaseEx:380, gst:0.05, customRL:"", customDM:"", plAddon:300,  customAddon:"", active:true, notes:"", createdAt:T0, updatedAt:T0 },
  { id:"i2", cat:"Comforters", bId:"b2", name:"Winter Hollow Fibre Comforter", purchaseEx:550, gst:0.05, customRL:"", customDM:"", plAddon:500,  customAddon:"", active:true, notes:"", createdAt:T0, updatedAt:T0 },
  { id:"i3", cat:"Towels",     bId:"b3", name:"Premium Bath Towel 500 GSM",    purchaseEx:180, gst:0.18, customRL:"", customDM:"", plAddon:200,  customAddon:"", active:true, notes:"", createdAt:T0, updatedAt:T0 },
];

// ── HELPERS ───────────────────────────────────────────────────────
const uid    = () => Date.now().toString(36) + Math.random().toString(36).slice(2,5);
const fp     = (n) => (n!=null&&!isNaN(n)) ? "₹"+Number(n).toLocaleString("en-IN") : "—";
const fpct   = (n) => (n!=null&&n!==""&&!isNaN(parseFloat(n))) ? (parseFloat(n)*100).toFixed(0)+"%" : "—";
const fpctRaw= (s) => (s!==""&&s!=null&&!isNaN(parseFloat(s))) ? parseFloat(s).toFixed(1)+"%" : null;

function fmtDate(iso) {
  if (!iso) return "";
  return new Date(iso).toLocaleDateString("en-IN",{day:"2-digit",month:"short",year:"numeric"});
}
function fmtDateTime(iso) {
  if (!iso) return "";
  const d = new Date(iso);
  return fmtDate(iso)+" "+d.toLocaleTimeString("en-IN",{hour:"2-digit",minute:"2-digit",hour12:true});
}
function dateGroupKey(iso) {
  if (!iso) return "Unknown";
  const d     = new Date(iso);
  const today = new Date();
  const yest  = new Date(); yest.setDate(today.getDate()-1);
  if (d.toDateString()===today.toDateString()) return "Today";
  if (d.toDateString()===yest.toDateString())  return "Yesterday";
  return d.toLocaleDateString("en-IN",{weekday:"long",day:"2-digit",month:"short",year:"numeric"});
}
function sameDay(a,b) { return new Date(a).toDateString()===new Date(b).toDateString(); }

// ── PALETTE ───────────────────────────────────────────────────────
const C = {
  bg:"#F0F2F7", card:"#fff", border:"#E5E8EF", text:"#111827", sec:"#6B7280", mute:"#9CA3AF",
  blue:"#1648D6", navy:"#0B1D5C",
  rl:"#067D62", rlBg:"#ECFDF5", rlBr:"#6EE7B7",
  dm:"#1C64F2", dmBg:"#EFF6FF", dmBr:"#93C5FD",
  pl:"#6D28D9", plBg:"#F5F3FF", plBr:"#C4B5FD",
  profit:"#065F46", profBg:"#F0FDF4",
  amb:"#B45309", ambBg:"#FEF3C7",
  red:"#DC2626", redBg:"#FEF2F2",
  new_:"#0E7490", newBg:"#ECFEFF", newBr:"#A5F3FC",
  upd:"#7C3AED", updBg:"#F5F3FF", updBr:"#DDD6FE",
  sync:"#059669", syncBg:"#ECFDF5",
};

// ── TOAST ─────────────────────────────────────────────────────────
let _toast = null;
const toast = (m, t="ok") => _toast && _toast({id:uid(), m, t});
function ToastHost() {
  const [list, setList] = useState([]);
  useEffect(() => {
    _toast = (x) => {
      setList(p => [...p, x]);
      setTimeout(() => setList(p => p.filter(y => y.id !== x.id)), 2600);
    };
  }, []);
  const BG = { ok:"#059669", err:C.red, warn:"#D97706" };
  return (
    <div style={{position:"fixed",top:14,right:14,zIndex:9999,display:"flex",flexDirection:"column",gap:8}}>
      {list.map(x => (
        <div key={x.id} style={{padding:"11px 16px",borderRadius:10,fontSize:13,fontWeight:700,color:"#fff",background:BG[x.t]||BG.ok,boxShadow:"0 4px 16px rgba(0,0,0,0.2)"}}>
          {x.m}
        </div>
      ))}
    </div>
  );
}

// ── SYNC STATUS BADGE ─────────────────────────────────────────────
function SyncBadge({ status }) {
  // status: "synced" | "saving" | "error" | "offline"
  const cfg = {
    synced:  { col:"#059669", bg:"#ECFDF5", br:"#6EE7B7", icon:"☁️", label:"Synced"   },
    saving:  { col:"#D97706", bg:"#FEF3C7", br:"#FCD34D", icon:"⏳", label:"Saving…"  },
    error:   { col:C.red,     bg:C.redBg,   br:"#FCA5A5", icon:"⚠️", label:"Sync error"},
    offline: { col:C.sec,     bg:"#F3F4F6", br:C.border,  icon:"📵", label:"Offline"  },
  }[status] || { col:C.sec, bg:"#F3F4F6", br:C.border, icon:"…", label:status };
  return (
    <span style={{display:"inline-flex",alignItems:"center",gap:4,padding:"3px 9px",borderRadius:20,fontSize:11,fontWeight:700,color:cfg.col,background:cfg.bg,border:"1px solid "+cfg.br}}>
      {cfg.icon} {cfg.label}
    </span>
  );
}

// ── UI ATOMS ──────────────────────────────────────────────────────
function Chip({ col, bg, br, children }) {
  return <span style={{display:"inline-block",padding:"2px 9px",borderRadius:20,fontSize:11,fontWeight:700,color:col,background:bg,border:"1px solid "+br}}>{children}</span>;
}
function PBox({ label, val, col, bg, br }) {
  return (
    <div style={{flex:1,background:bg,border:"1px solid "+br,borderRadius:10,padding:"9px 6px",textAlign:"center"}}>
      <div style={{fontSize:9,fontWeight:700,color:C.mute,letterSpacing:"0.5px",textTransform:"uppercase",marginBottom:3}}>{label}</div>
      <div style={{fontSize:15,fontWeight:800,color:col}}>{val}</div>
    </div>
  );
}
function Fld({ label, hint, err, children }) {
  return (
    <div style={{marginBottom:14}}>
      <label style={{display:"block",fontSize:11,fontWeight:700,color:C.sec,letterSpacing:"0.7px",textTransform:"uppercase",marginBottom:5}}>{label}</label>
      {children}
      {err  && <div style={{fontSize:11,color:C.red,marginTop:3}}>{err}</div>}
      {!err && hint && <div style={{fontSize:11,color:C.mute,marginTop:3}}>{hint}</div>}
    </div>
  );
}
const INP = {width:"100%",padding:"11px 13px",borderRadius:9,fontSize:14,border:"1.5px solid #E5E8EF",background:"#FAFBFD",fontFamily:"inherit",outline:"none",boxSizing:"border-box",color:"#111827"};
const SEL = Object.assign({}, INP, {cursor:"pointer"});
function BtnP({ color, onClick, children, style }) {
  return <button onClick={onClick} style={Object.assign({padding:"13px 20px",borderRadius:10,border:"none",background:color||C.blue,color:"#fff",fontWeight:700,fontSize:14,cursor:"pointer",fontFamily:"inherit",width:"100%"},style||{})}>{children}</button>;
}
function BtnO({ color, onClick, children, style }) {
  const col = color||C.sec;
  return <button onClick={onClick} style={Object.assign({padding:"13px 20px",borderRadius:10,border:"1.5px solid "+col,background:"#fff",color:col,fontWeight:600,fontSize:14,cursor:"pointer",fontFamily:"inherit",width:"100%"},style||{})}>{children}</button>;
}

// ── MARGIN OVERRIDE INPUT ─────────────────────────────────────────
function MarginOverride({ label, col, bg, br, value, onChange, srcLabel, err }) {
  const has = value!==""&&value!=null&&!isNaN(parseFloat(value));
  return (
    <div style={{marginBottom:14}}>
      <label style={{display:"block",fontSize:11,fontWeight:700,color:C.sec,letterSpacing:"0.7px",textTransform:"uppercase",marginBottom:5}}>{label}</label>
      <div style={{position:"relative"}}>
        <input style={Object.assign({},INP,{borderColor:err?C.red:has?col:C.border,paddingRight:has?"90px":"13px"})}
          type="number" step="0.5" min="0" max="100"
          placeholder={"Leave blank = "+srcLabel} value={value}
          onChange={e => onChange(e.target.value)} />
        {has && (
          <div style={{position:"absolute",right:10,top:"50%",transform:"translateY(-50%)",background:bg,border:"1px solid "+br,color:col,fontSize:11,fontWeight:700,padding:"2px 8px",borderRadius:6}}>
            {parseFloat(value).toFixed(1)+"%"}
          </div>
        )}
      </div>
      {err && <div style={{fontSize:11,color:C.red,marginTop:3}}>{err}</div>}
      {!err && (
        <div style={{fontSize:11,color:has?col:C.mute,fontWeight:has?600:400,marginTop:3}}>
          {has ? "⚡ Item override active — brand/default ignored" : "Using: "+srcLabel}
        </div>
      )}
    </div>
  );
}

// ── PROFIT ROW ────────────────────────────────────────────────────
function ProfitRow({ label, col, price, profit }) {
  if (!price) return null;
  return (
    <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"5px 10px",background:C.profBg,borderRadius:7,marginBottom:4,border:"1px solid #BBF7D0"}}>
      <div style={{display:"flex",alignItems:"center",gap:8}}>
        <span style={{fontSize:11,fontWeight:700,color:col}}>{label}</span>
        <span style={{fontSize:13,fontWeight:800,color:col}}>{fp(price)}</span>
      </div>
      {profit ? (
        <div style={{display:"flex",alignItems:"center",gap:6}}>
          <span style={{fontSize:11,fontWeight:700,color:C.profit}}>{"+"+fp(profit.amt)}</span>
          <span style={{background:"#D1FAE5",color:C.profit,fontSize:11,fontWeight:800,padding:"2px 8px",borderRadius:20}}>{profit.pct+"%"}</span>
        </div>
      ) : (
        <span style={{fontSize:11,color:C.mute}}>—</span>
      )}
    </div>
  );
}

// ── ITEM CARD ─────────────────────────────────────────────────────
function ItemCard({ it, isAdmin, showDate }) {
  const isNew    = it.createdAt && it.updatedAt && sameDay(it.createdAt, it.updatedAt);
  const isEdited = it.createdAt && it.updatedAt && !sameDay(it.createdAt, it.updatedAt);
  return (
    <div style={{background:C.card,border:"1px solid "+C.border,borderRadius:12,padding:"13px",marginBottom:8,boxShadow:"0 1px 5px rgba(0,0,0,0.04)"}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:9}}>
        <div style={{flex:1,paddingRight:8}}>
          <div style={{fontWeight:700,fontSize:14,marginBottom:4}}>{it.name}</div>
          <div style={{display:"flex",gap:6,flexWrap:"wrap",alignItems:"center"}}>
            {it._b && <div style={{background:C.navy,color:"#fff",borderRadius:5,padding:"1px 8px",fontSize:10,fontWeight:800}}>{it._b.code}</div>}
            <Chip col={C.sec} bg="#F3F4F6" br={C.border}>{it.cat}</Chip>
            {isNew    && <Chip col={C.new_} bg={C.newBg} br={C.newBr}>🆕 New</Chip>}
            {isEdited && <Chip col={C.upd}  bg={C.updBg} br={C.updBr}>✏️ Updated</Chip>}
          </div>
          {showDate && (
            <div style={{marginTop:5,fontSize:10,color:C.mute}}>
              {isNew ? "Added "+fmtDateTime(it.createdAt) : "Updated "+fmtDateTime(it.updatedAt)}
              {isEdited && <span style={{marginLeft:6,color:C.updBr}}>· Added {fmtDate(it.createdAt)}</span>}
            </div>
          )}
        </div>
        <Chip col={C.amb} bg={C.ambBg} br="#FCD34D">{"GST "+(it.gst*100).toFixed(0)+"%"}</Chip>
      </div>
      {isAdmin && (
        <div style={{background:C.ambBg,border:"1px solid #FCD34D",borderRadius:8,padding:"7px 11px",marginBottom:9,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
          <span style={{fontSize:11,fontWeight:700,color:C.amb,textTransform:"uppercase"}}>Purchase Ex-GST</span>
          <span style={{fontWeight:800,fontSize:13,color:C.amb}}>
            {fp(it.purchaseEx)+" + "}
            <span style={{fontSize:10,fontWeight:600}}>{(it.gst*100).toFixed(0)+"% = "+fp(it.incGST)}</span>
          </span>
        </div>
      )}
      <div style={{display:"flex",gap:7}}>
        <PBox label="RL — Wholesale" val={fp(it.rl)} col={C.rl} bg={C.rlBg} br={C.rlBr} />
        <PBox label="DM — Sp. WHL"   val={fp(it.dm)} col={C.dm} bg={C.dmBg} br={C.dmBr} />
        <PBox label="PL — Retail"     val={fp(it.pl)} col={C.pl} bg={C.plBg} br={C.plBr} />
      </div>
      {isAdmin && (
        <div style={{marginTop:9}}>
          <div style={{fontSize:10,fontWeight:700,color:C.profit,textTransform:"uppercase",letterSpacing:"0.5px",marginBottom:5}}>💰 Profit per piece (Ex-GST basis)</div>
          <ProfitRow label="RL" col={C.rl} price={it.rl} profit={it.rlProfit} />
          <ProfitRow label="DM" col={C.dm} price={it.dm} profit={it.dmProfit} />
          <ProfitRow label="PL" col={C.pl} price={it.pl} profit={it.plProfit} />
        </div>
      )}
      {it.notes && <div style={{marginTop:7,fontSize:11,color:C.mute,fontStyle:"italic"}}>{"📝 "+it.notes}</div>}
    </div>
  );
}

// ── DRAWER ────────────────────────────────────────────────────────
function Drawer({ open, onClose, title, footer, children }) {
  if (!open) return null;
  return (
    <div>
      <div onClick={onClose} style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.45)",zIndex:500}} />
      <div style={{position:"fixed",bottom:0,left:0,right:0,zIndex:501,background:"#fff",borderRadius:"22px 22px 0 0",maxHeight:"92vh",display:"flex",flexDirection:"column",boxShadow:"0 -8px 40px rgba(0,0,0,0.18)"}}>
        <div style={{display:"flex",justifyContent:"center",padding:"10px 0 0"}}>
          <div style={{width:36,height:4,borderRadius:2,background:C.border}} />
        </div>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"8px 20px 12px"}}>
          <div style={{fontSize:18,fontWeight:800,color:C.text}}>{title}</div>
          <button onClick={onClose} style={{width:34,height:34,borderRadius:8,border:"1.5px solid "+C.border,background:"#F9FAFB",cursor:"pointer",fontSize:16,color:C.sec,fontFamily:"inherit"}}>✕</button>
        </div>
        <div style={{flex:1,overflowY:"auto",padding:"0 20px 6px"}}>{children}</div>
        {footer && <div style={{padding:"12px 20px",borderTop:"1px solid "+C.border}}>{footer}</div>}
      </div>
    </div>
  );
}

// ── CONFIRM ───────────────────────────────────────────────────────
function Confirm({ title, msg, onOk, onNo }) {
  return (
    <div style={{position:"fixed",inset:0,zIndex:600,background:"rgba(0,0,0,0.5)",display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <div style={{background:"#fff",borderRadius:18,padding:"26px 22px",maxWidth:320,width:"100%"}}>
        <div style={{fontSize:17,fontWeight:800,color:C.text,marginBottom:8}}>{title}</div>
        <div style={{fontSize:13,color:C.sec,lineHeight:1.7,marginBottom:24}}>{msg}</div>
        <div style={{display:"flex",gap:10}}>
          <BtnO onClick={onNo}>Cancel</BtnO>
          <BtnP color={C.red} onClick={onOk}>Delete</BtnP>
        </div>
      </div>
    </div>
  );
}

// ── PIN PAD ───────────────────────────────────────────────────────
function PINPad({ label, pin, onSuccess }) {
  const [val, setVal] = useState("");
  const [err, setErr] = useState("");
  const [shk, setShk] = useState(false);

  function press(d) {
    if (val.length >= 4) return;
    const next = val + d;
    setVal(next);
    if (next.length === 4) {
      if (next === String(pin)) {
        onSuccess();
      } else {
        setShk(true); setErr("Wrong PIN — try again.");
        setTimeout(() => { setVal(""); setShk(false); setErr(""); }, 700);
      }
    }
  }

  const keys = ["1","2","3","4","5","6","7","8","9","","0","⌫"];
  return (
    <div style={{textAlign:"center"}}>
      <div style={{fontSize:13,fontWeight:700,color:C.sec,marginBottom:16}}>{label}</div>
      <div style={{display:"flex",justifyContent:"center",gap:14,marginBottom:6}}>
        {[0,1,2,3].map(i => (
          <div key={i} style={{width:15,height:15,borderRadius:"50%",background:val.length>i?C.blue:"transparent",border:"2.5px solid "+(val.length>i?C.blue:C.border),transition:"all 0.12s",animation:shk?"shk 0.35s ease":"none"}} />
        ))}
      </div>
      <div style={{minHeight:22,fontSize:12,color:C.red,fontWeight:600,marginBottom:8}}>{err}</div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:10,maxWidth:220,margin:"0 auto"}}>
        {keys.map((k,i) => (
          <button key={i} disabled={k===""} onClick={() => k==="⌫" ? setVal(p=>p.slice(0,-1)) : k ? press(k) : null}
            style={{height:56,borderRadius:12,fontSize:k==="⌫"?20:22,fontWeight:700,cursor:k===""?"default":"pointer",fontFamily:"inherit",background:k===""?"transparent":"#fff",border:k===""?"none":"1.5px solid "+C.border,color:k==="⌫"?C.red:C.text,boxShadow:k!==""?"0 2px 6px rgba(0,0,0,0.07)":"none"}}>
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
    <div style={{minHeight:"100vh",background:"linear-gradient(160deg,"+C.navy+" 0%,"+C.blue+" 100%)",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",padding:20}}>
      <div style={{marginBottom:30,textAlign:"center"}}>
        <div style={{width:66,height:66,borderRadius:18,background:"rgba(255,255,255,0.13)",border:"2px solid rgba(255,255,255,0.2)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:24,fontWeight:900,color:"#fff",margin:"0 auto 12px"}}>VK</div>
        <div style={{fontSize:20,fontWeight:800,color:"#fff"}}>{settings.co}</div>
        <div style={{fontSize:12,color:"rgba(255,255,255,0.4)",marginTop:3}}>{settings.tag}</div>
      </div>
      <div style={{background:"#fff",borderRadius:22,padding:"28px 24px",width:"100%",maxWidth:340,boxShadow:"0 30px 80px rgba(0,0,0,0.3)"}}>
        {!role ? (
          <div>
            <div style={{fontSize:16,fontWeight:800,color:C.text,textAlign:"center",marginBottom:18}}>Select your role</div>
            {roles.map(o => (
              <button key={o.r} onClick={() => setRole(o.r)}
                style={{display:"flex",alignItems:"center",gap:14,padding:"14px 16px",borderRadius:12,border:"1.5px solid "+C.border,background:"#FAFBFD",cursor:"pointer",width:"100%",fontFamily:"inherit",marginBottom:10,textAlign:"left"}}>
                <span style={{fontSize:24}}>{o.icon}</span>
                <div style={{flex:1}}>
                  <div style={{fontWeight:700,fontSize:14,color:C.text}}>{o.t}</div>
                  <div style={{fontSize:12,color:C.sec,marginTop:1}}>{o.d}</div>
                </div>
                <span style={{color:C.mute,fontSize:18}}>›</span>
              </button>
            ))}
          </div>
        ) : (
          <div>
            <button onClick={() => setRole(null)} style={{background:"none",border:"none",cursor:"pointer",color:C.sec,fontSize:13,fontWeight:600,marginBottom:16,padding:0,fontFamily:"inherit"}}>← Back</button>
            <PINPad
              label={"Enter "+(role==="admin"?"Admin":"Salesman")+" PIN"}
              pin={role==="admin"?settings.adminPIN:settings.salesPIN}
              onSuccess={() => onLogin(role)}
            />
          </div>
        )}
      </div>
      <div style={{marginTop:16,fontSize:11,color:"rgba(255,255,255,0.25)"}}>{"v"+VER+" · Firebase sync enabled"}</div>
    </div>
  );
}

// ── BRANDS VIEW ───────────────────────────────────────────────────
function BrandsView({ brands, onBrandsChange }) {
  const EF = { code:"", name:"", rlMarkup:"18", dmMarkup:"" };
  const [form, setForm] = useState(EF);
  const [eid,  setEid]  = useState(null);
  const [open, setOpen] = useState(false);
  const [conf, setConf] = useState(null);

  function close() { setOpen(false); setEid(null); setForm(EF); }
  function startEdit(b) {
    setForm({ code:b.code, name:b.name, rlMarkup:+(b.rlMarkup*100).toFixed(1), dmMarkup:b.dmMarkup!=null?+(b.dmMarkup*100).toFixed(1):"" });
    setEid(b.id); setOpen(true);
  }
  function save() {
    if (!form.code.trim()) return toast("Brand code required","err");
    if (!form.name.trim()) return toast("Brand name required","err");
    if (isNaN(parseFloat(form.rlMarkup))) return toast("RL markup must be a number","err");
    const entry = {
      id: eid||uid(), code:form.code.trim().toUpperCase(), name:form.name.trim(),
      rlMarkup: parseFloat(form.rlMarkup)/100,
      dmMarkup: form.dmMarkup!==""&&!isNaN(parseFloat(form.dmMarkup)) ? parseFloat(form.dmMarkup)/100 : null,
    };
    const next = eid ? brands.map(b => b.id===eid?entry:b) : [...brands, entry];
    onBrandsChange(next);
    toast(eid?"Brand updated":"Brand added"); close();
  }
  function del(id) { onBrandsChange(brands.filter(b => b.id!==id)); setConf(null); toast("Brand removed","warn"); }

  return (
    <div style={{padding:"16px 16px 80px"}}>
      {conf && <Confirm title="Remove Brand?" msg="Items using this brand will keep prices but lose markup override." onOk={() => del(conf)} onNo={() => setConf(null)} />}
      <div style={{background:"#EFF6FF",border:"1px solid #BFDBFE",borderRadius:10,padding:"10px 14px",marginBottom:14,fontSize:12,color:"#1E40AF",lineHeight:1.6}}>
        <strong>Brand markups are defaults.</strong> You can override RL and DM per item in Master Sheet.
      </div>
      <BtnP onClick={() => { setForm(EF); setEid(null); setOpen(true); }}>+ Add Brand</BtnP>
      <div style={{height:12}} />
      {brands.length===0 && <div style={{textAlign:"center",padding:"36px",color:C.mute,fontSize:14}}>No brands yet.</div>}
      {brands.map(b => (
        <div key={b.id} style={{background:C.card,border:"1px solid "+C.border,borderRadius:13,padding:"14px",marginBottom:10,boxShadow:"0 1px 6px rgba(0,0,0,0.05)"}}>
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:8}}>
            <div style={{display:"flex",alignItems:"center",gap:9}}>
              <div style={{background:C.navy,color:"#fff",borderRadius:7,padding:"4px 10px",fontWeight:800,fontSize:12}}>{b.code}</div>
              <div style={{fontWeight:700,fontSize:14}}>{b.name}</div>
            </div>
            <div style={{display:"flex",gap:7}}>
              <button onClick={() => startEdit(b)} style={{padding:"5px 13px",borderRadius:7,border:"1.5px solid "+C.border,background:"#F9FAFB",cursor:"pointer",fontSize:12,fontWeight:600,color:C.sec,fontFamily:"inherit"}}>Edit</button>
              <button onClick={() => setConf(b.id)} style={{padding:"5px 10px",borderRadius:7,border:"1.5px solid #FCA5A5",background:C.redBg,cursor:"pointer",fontSize:13,color:C.red,fontFamily:"inherit"}}>✕</button>
            </div>
          </div>
          <div style={{display:"flex",gap:7}}>
            <Chip col={C.rl} bg={C.rlBg} br={C.rlBr}>{"RL "+fpct(b.rlMarkup)}</Chip>
            {b.dmMarkup!=null
              ? <Chip col={C.dm} bg={C.dmBg} br={C.dmBr}>{"DM "+fpct(b.dmMarkup)}</Chip>
              : <Chip col={C.mute} bg="#F3F4F6" br={C.border}>DM not set</Chip>}
          </div>
        </div>
      ))}
      <Drawer open={open} onClose={close} title={eid?"Edit Brand":"Add Brand"}
        footer={
          <div style={{display:"flex",gap:9}}>
            <BtnO onClick={close}>Cancel</BtnO>
            <BtnP color={eid?"#F59E0B":C.blue} onClick={save}>{eid?"Update":"Add Brand"}</BtnP>
          </div>
        }>
        <Fld label="Brand Code *" hint="3–6 letters e.g. TRI, STH, LOC">
          <input style={INP} maxLength={6} value={form.code} placeholder="TRI" onChange={e => setForm(p => ({...p,code:e.target.value.toUpperCase()}))} />
        </Fld>
        <Fld label="Brand Name *">
          <input style={INP} value={form.name} placeholder="e.g. Trident" onChange={e => setForm(p => ({...p,name:e.target.value}))} />
        </Fld>
        <Fld label="Default RL Markup % *" hint="Enter number: 18 = 18%">
          <input style={INP} type="number" step="0.5" value={form.rlMarkup} placeholder="18" onChange={e => setForm(p => ({...p,rlMarkup:e.target.value}))} />
        </Fld>
        <Fld label="Default DM Markup %" hint="Enter number: 10 = 10%. Leave blank to decide later.">
          <input style={INP} type="number" step="0.5" value={form.dmMarkup} placeholder="Leave blank" onChange={e => setForm(p => ({...p,dmMarkup:e.target.value}))} />
        </Fld>
      </Drawer>
    </div>
  );
}

// ── EXPORT MODAL ──────────────────────────────────────────────────
function ExportModal({ items, brands, settings, onClose }) {
  const [fCat,   setFCat]   = useState("All");
  const [fBrand, setFBrand] = useState("All");
  const [fSt,    setFSt]    = useState("Active");

  const filtered = useMemo(() => {
    return items.filter(i => {
      if (fCat!=="All" && i.cat!==fCat) return false;
      if (fBrand!=="All" && i.bId!==fBrand) return false;
      if (fSt==="Active" && !i.active) return false;
      if (fSt==="Inactive" && i.active) return false;
      return true;
    }).map(i => {
      const b = brands.find(x => x.id===i.bId)||null;
      return Object.assign({}, i, computeItem(i,b,settings), {_b:b});
    });
  }, [items, brands, settings, fCat, fBrand, fSt]);

  const today = new Date().toLocaleDateString("en-IN",{day:"2-digit",month:"short",year:"numeric"});

  function doXLSX(rows, fname) {
    if (!rows.length) return toast("No items to export","warn");
    try {
      const ws = XLSX.utils.json_to_sheet(rows);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Price List");
      XLSX.writeFile(wb, fname+"_"+today+".xlsx");
      toast("Excel downloaded!");
    } catch { toast("Export failed","err"); }
  }
  function doJSON(rows, fname) {
    if (!rows.length) return toast("No items to export","warn");
    const blob = new Blob([JSON.stringify({exportedAt:new Date().toISOString(),data:rows},null,2)],{type:"application/json"});
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob); a.download = fname+"_"+today+".json"; a.click();
    toast("JSON downloaded!");
  }
  function adminRows() {
    return filtered.map(it => ({
      "Category":it.cat, "Brand Code":it._b?it._b.code:"", "Brand Name":it._b?it._b.name:"", "Item Name":it.name,
      "Purchase Ex-GST":it.purchaseEx, "GST %":(it.gst*100).toFixed(0)+"%", "Purchase Inc-GST":it.incGST,
      "RL Markup":fpct(it.rlM), "RL Price":it.rl!=null?it.rl:"", "RL Profit Rs":it.rlProfit?it.rlProfit.amt:"", "RL Profit %":it.rlProfit?it.rlProfit.pct:"",
      "DM Markup":it.dmM!=null?fpct(it.dmM):"", "DM Price":it.dm!=null?it.dm:"", "DM Profit Rs":it.dmProfit?it.dmProfit.amt:"", "DM Profit %":it.dmProfit?it.dmProfit.pct:"",
      "PL Addon":it.add!=null?it.add:"", "PL Price":it.pl!=null?it.pl:"", "PL Profit Rs":it.plProfit?it.plProfit.amt:"", "PL Profit %":it.plProfit?it.plProfit.pct:"",
      "Status":it.active?"Active":"Inactive", "Added On":fmtDate(it.createdAt), "Updated On":fmtDate(it.updatedAt), "Notes":it.notes,
    }));
  }
  function salesRows() {
    return filtered.map(it => ({
      "Category":it.cat, "Brand Code":it._b?it._b.code:"", "Brand Name":it._b?it._b.name:"", "Item Name":it.name,
      "GST %":(it.gst*100).toFixed(0)+"%", "RL Price":it.rl!=null?it.rl:"", "DM Price":it.dm!=null?it.dm:"", "PL Price":it.pl!=null?it.pl:"", "Notes":it.notes,
    }));
  }

  return (
    <div style={{position:"fixed",inset:0,zIndex:600,background:"rgba(0,0,0,0.5)",display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
      <div style={{background:"#fff",borderRadius:20,padding:"22px 18px",maxWidth:400,width:"100%",maxHeight:"88vh",overflowY:"auto"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
          <div style={{fontSize:17,fontWeight:800}}>⬇ Export Price List</div>
          <button onClick={onClose} style={{width:32,height:32,borderRadius:8,border:"1.5px solid "+C.border,background:"#F9FAFB",cursor:"pointer",fontSize:15,color:C.sec,fontFamily:"inherit"}}>✕</button>
        </div>
        <div style={{background:"#F8FAFF",border:"1px solid #BFDBFE",borderRadius:10,padding:"13px",marginBottom:14}}>
          <div style={{fontSize:11,fontWeight:700,color:C.blue,textTransform:"uppercase",letterSpacing:"0.6px",marginBottom:10}}>Filter Items to Export</div>
          <Fld label="Category">
            <select style={SEL} value={fCat} onChange={e => setFCat(e.target.value)}>
              <option value="All">All Categories</option>
              {CATS.map(c => <option key={c} value={c}>{c}</option>)}
            </select>
          </Fld>
          <Fld label="Brand">
            <select style={SEL} value={fBrand} onChange={e => setFBrand(e.target.value)}>
              <option value="All">All Brands</option>
              {brands.map(b => <option key={b.id} value={b.id}>{b.code+" — "+b.name}</option>)}
            </select>
          </Fld>
          <Fld label="Status">
            <div style={{display:"flex",border:"1px solid "+C.border,borderRadius:9,overflow:"hidden"}}>
              {["All","Active","Inactive"].map(s => (
                <button key={s} onClick={() => setFSt(s)} style={{flex:1,padding:"9px",border:"none",cursor:"pointer",fontWeight:600,fontSize:12,fontFamily:"inherit",background:fSt===s?C.blue:"#fff",color:fSt===s?"#fff":C.sec}}>{s}</button>
              ))}
            </div>
          </Fld>
          <div style={{fontSize:12,color:C.mute,fontWeight:600}}>{filtered.length+" items selected"}</div>
        </div>
        <div style={{background:C.ambBg,border:"1px solid #FCD34D",borderRadius:10,padding:"13px",marginBottom:10}}>
          <div style={{fontSize:12,fontWeight:800,color:C.amb,marginBottom:3}}>🔐 Admin Export</div>
          <div style={{fontSize:11,color:C.amb,marginBottom:10}}>Includes purchase price + profit. Do NOT share with staff.</div>
          <div style={{display:"flex",gap:8}}>
            <BtnP color={C.amb} onClick={() => doXLSX(adminRows(),"VKF_Admin")} style={{fontSize:12,padding:"10px"}}>📊 Excel</BtnP>
            <BtnO color={C.amb} onClick={() => doJSON(adminRows(),"VKF_Admin")} style={{fontSize:12,padding:"10px"}}>JSON</BtnO>
          </div>
        </div>
        <div style={{background:C.rlBg,border:"1px solid "+C.rlBr,borderRadius:10,padding:"13px"}}>
          <div style={{fontSize:12,fontWeight:800,color:C.rl,marginBottom:3}}>📋 Salesman Export</div>
          <div style={{fontSize:11,color:C.rl,marginBottom:10}}>Prices only — no costs. Safe to share.</div>
          <div style={{display:"flex",gap:8}}>
            <BtnP color={C.rl} onClick={() => doXLSX(salesRows(),"VKF_PriceList")} style={{fontSize:12,padding:"10px"}}>📊 Excel</BtnP>
            <BtnO color={C.rl} onClick={() => doJSON(salesRows(),"VKF_PriceList")} style={{fontSize:12,padding:"10px"}}>JSON</BtnO>
          </div>
        </div>
      </div>
    </div>
  );
}

// ── MASTER VIEW ───────────────────────────────────────────────────
const EI = { cat:"Bed Sheets",bId:"",name:"",purchaseEx:"",gst:0.05,customRL:"",customDM:"",plAddon:"",customAddon:"",active:true,notes:"" };

function MasterView({ brands, items, onItemsChange, settings }) {
  const [open,    setOpen]    = useState(false);
  const [eid,     setEid]     = useState(null);
  const [form,    setForm]    = useState(EI);
  const [errs,    setErrs]    = useState({});
  const [fCat,    setFCat]    = useState("All");
  const [fSt,     setFSt]     = useState("Active");
  const [conf,    setConf]    = useState(null);
  const [showExp, setShowExp] = useState(false);

  const brand  = brands.find(b => b.id===form.bId)||null;
  const prevEx = parseFloat(form.purchaseEx)||0;
  const prev   = prevEx>0 ? computeItem(Object.assign({},form,{purchaseEx:prevEx}), brand, settings) : null;

  function close() { setOpen(false); setEid(null); }
  function openAdd() { setForm(EI); setEid(null); setErrs({}); setOpen(true); }
  function openEdit(it) { setForm(Object.assign({},it,{customDM:it.customDM||""})); setEid(it.id); setErrs({}); setOpen(true); }

  function validate() {
    const e = {};
    if (!form.name.trim()) e.name = "Item name required";
    const px = parseFloat(form.purchaseEx);
    if (!form.purchaseEx||isNaN(px)||px<=0) e.px = "Enter a valid purchase price";
    if (form.customRL!==""&&form.customRL!=null) { const v=parseFloat(form.customRL); if(isNaN(v)||v<0||v>100) e.crl="Enter 0–100 (e.g. 15 for 15%)"; }
    if (form.customDM!==""&&form.customDM!=null) { const v=parseFloat(form.customDM); if(isNaN(v)||v<0||v>100) e.cdm="Enter 0–100 (e.g. 10 for 10%)"; }
    if (form.plAddon==="custom") { const cv=parseFloat(form.customAddon); if(isNaN(cv)||cv<=0) e.ca="Enter a valid amount"; }
    setErrs(e);
    return Object.keys(e).length===0;
  }

  function save() {
    if (!validate()) return;
    const now = new Date().toISOString();
    const orig = eid ? items.find(i => i.id===eid) : null;
    const entry = Object.assign({}, form, {
      id: eid||uid(),
      purchaseEx: parseFloat(form.purchaseEx),
      gst: parseFloat(form.gst),
      createdAt: orig ? (orig.createdAt||now) : now,
      updatedAt: now,
    });
    const next = eid ? items.map(i => i.id===eid?entry:i) : [...items, entry];
    onItemsChange(next);
    toast(eid?"Item updated":"Item added");
    close();
  }
  function del(id) { onItemsChange(items.filter(i => i.id!==id)); setConf(null); toast("Item removed","warn"); }
  function toggle(id) {
    onItemsChange(items.map(i => i.id===id ? Object.assign({},i,{active:!i.active,updatedAt:new Date().toISOString()}) : i));
  }

  const list = useMemo(() => {
    return items.filter(i => {
      if (fCat!=="All" && i.cat!==fCat) return false;
      if (fSt==="Active" && !i.active) return false;
      if (fSt==="Inactive" && i.active) return false;
      return true;
    }).map(i => {
      const b = brands.find(x => x.id===i.bId)||null;
      return Object.assign({}, i, computeItem(i,b,settings), {_b:b});
    });
  }, [items, brands, settings, fCat, fSt]);

  function ei(k) { return errs[k] ? Object.assign({},INP,{borderColor:C.red}) : INP; }

  const rlSrc = fpctRaw(form.customRL)
    ? "item override ("+fpctRaw(form.customRL)+")"
    : brand ? "brand "+brand.code+" ("+(brand.rlMarkup*100).toFixed(0)+"%)"
    : "default ("+(settings.defaultRL*100).toFixed(0)+"%)";
  const dmSrc = fpctRaw(form.customDM)
    ? "item override ("+fpctRaw(form.customDM)+")"
    : brand&&brand.dmMarkup!=null ? "brand "+brand.code+" ("+(brand.dmMarkup*100).toFixed(0)+"%)"
    : settings.defaultDM ? "default ("+(parseFloat(settings.defaultDM)*100).toFixed(0)+"%)"
    : "not set";

  return (
    <div style={{padding:"16px 16px 80px"}}>
      {conf && <Confirm title="Delete Item?" msg="This will permanently remove the item." onOk={() => del(conf)} onNo={() => setConf(null)} />}
      {showExp && <ExportModal items={items} brands={brands} settings={settings} onClose={() => setShowExp(false)} />}

      <div style={{background:C.ambBg,border:"1px solid #FCD34D",borderRadius:10,padding:"9px 13px",marginBottom:12,fontSize:12,color:C.amb,fontWeight:600}}>
        ⚠ Enter purchase price WITHOUT GST. All RL / DM / PL prices are Ex-GST.
      </div>

      <div style={{display:"flex",gap:8,marginBottom:12,flexWrap:"wrap"}}>
        <select style={Object.assign({},SEL,{flex:1,padding:"9px 11px"})} value={fCat} onChange={e => setFCat(e.target.value)}>
          <option value="All">All Categories</option>
          {CATS.map(c => <option key={c} value={c}>{c}</option>)}
        </select>
        <div style={{display:"flex",border:"1px solid "+C.border,borderRadius:9,overflow:"hidden"}}>
          {["All","Active","Inactive"].map(s => (
            <button key={s} onClick={() => setFSt(s)} style={{padding:"9px 11px",border:"none",cursor:"pointer",fontWeight:600,fontSize:12,fontFamily:"inherit",background:fSt===s?C.blue:"#fff",color:fSt===s?"#fff":C.sec}}>{s}</button>
          ))}
        </div>
      </div>

      <div style={{display:"grid",gridTemplateColumns:"2fr 1fr",gap:8,marginBottom:12}}>
        <BtnP onClick={openAdd}>+ Add Item</BtnP>
        <BtnO color={C.blue} onClick={() => setShowExp(true)}>⬇ Export</BtnO>
      </div>

      <div style={{fontSize:12,color:C.mute,marginBottom:10,fontWeight:600}}>{list.length+" item"+(list.length!==1?"s":"")}</div>
      {list.length===0 && <div style={{textAlign:"center",padding:"36px",color:C.mute,fontSize:14}}>No items. Add using the button above.</div>}

      {list.map(it => (
        <div key={it.id}>
          <ItemCard it={it} isAdmin={true} showDate={true} />
          <div style={{display:"flex",gap:8,marginTop:-4,marginBottom:8}}>
            <button onClick={() => openEdit(it)} style={{flex:1,padding:"9px",borderRadius:8,border:"1.5px solid "+C.border,background:"#F9FAFB",cursor:"pointer",fontSize:13,fontWeight:600,color:C.sec,fontFamily:"inherit"}}>Edit</button>
            <button onClick={() => toggle(it.id)} style={{flex:1,padding:"9px",borderRadius:8,border:"1.5px solid "+C.border,background:"#F9FAFB",cursor:"pointer",fontSize:12,fontWeight:600,color:it.active?"#991B1B":"#065F46",fontFamily:"inherit"}}>{it.active?"Set Inactive":"Set Active"}</button>
            <button onClick={() => setConf(it.id)} style={{width:40,borderRadius:8,border:"1.5px solid #FCA5A5",background:C.redBg,cursor:"pointer",fontSize:15,color:C.red,fontFamily:"inherit"}}>✕</button>
          </div>
        </div>
      ))}

      <Drawer open={open} onClose={close} title={eid?"Edit Item":"Add Item"}
        footer={
          <div style={{display:"flex",gap:9}}>
            <BtnO onClick={close}>Cancel</BtnO>
            <BtnP color={eid?"#F59E0B":C.blue} onClick={save}>{eid?"Update Item":"Save Item"}</BtnP>
          </div>
        }>
        <Fld label="Category *">
          <select style={SEL} value={form.cat} onChange={e => setForm(p => ({...p,cat:e.target.value,bId:""}))}>
            {CATS.map(c => <option key={c} value={c}>{c}</option>)}
          </select>
        </Fld>
        <Fld label="Brand">
          <select style={SEL} value={form.bId} onChange={e => setForm(p => ({...p,bId:e.target.value}))}>
            <option value="">— Select Brand (optional) —</option>
            {brands.map(b => <option key={b.id} value={b.id}>{b.code+" — "+b.name}</option>)}
          </select>
          {brand && (
            <div style={{marginTop:4,display:"flex",gap:6}}>
              <Chip col={C.rl} bg={C.rlBg} br={C.rlBr}>{"RL "+fpct(brand.rlMarkup)}</Chip>
              {brand.dmMarkup!=null
                ? <Chip col={C.dm} bg={C.dmBg} br={C.dmBr}>{"DM "+fpct(brand.dmMarkup)}</Chip>
                : <Chip col={C.mute} bg="#F9FAFB" br={C.border}>DM not set</Chip>}
            </div>
          )}
        </Fld>
        <Fld label="Item Name *" err={errs.name}>
          <input style={ei("name")} value={form.name} placeholder="e.g. King Cotton Bedsheet 200TC" onChange={e => setForm(p => ({...p,name:e.target.value}))} />
        </Fld>
        <Fld label="Purchase Price WITHOUT GST (₹) *" err={errs.px}>
          <input style={Object.assign({},ei("px"),{fontWeight:700,fontSize:16,color:C.blue})} type="number" step="0.01" placeholder="e.g. 380" value={form.purchaseEx} onChange={e => setForm(p => ({...p,purchaseEx:e.target.value}))} />
        </Fld>
        <Fld label="GST %">
          <select style={SEL} value={form.gst} onChange={e => setForm(p => ({...p,gst:parseFloat(e.target.value)}))}>
            {GST_OPTS.map(o => <option key={o.v} value={o.v}>{o.label}</option>)}
          </select>
          {prevEx>0 && <div style={{marginTop:4,fontSize:12,color:C.amb,fontWeight:600}}>{"Purchase Inc-GST = "+fp(calcIncGST(prevEx,form.gst))}</div>}
        </Fld>
        <MarginOverride label="Custom RL Margin % (Override)" col={C.rl} bg={C.rlBg} br={C.rlBr} value={form.customRL} onChange={v => setForm(p => ({...p,customRL:v}))} srcLabel={rlSrc} err={errs.crl} />
        <MarginOverride label="Custom DM Margin % (Override)" col={C.dm} bg={C.dmBg} br={C.dmBr} value={form.customDM} onChange={v => setForm(p => ({...p,customDM:v}))} srcLabel={dmSrc} err={errs.cdm} />
        <Fld label="PL Add-on (extra ₹ on top of RL for retail)" err={errs.ca}>
          <select style={SEL} value={form.plAddon||""} onChange={e => { const v=e.target.value; setForm(p => ({...p,plAddon:v===""?"":v==="custom"?"custom":parseInt(v)})); }}>
            {ADDON_OPTS.map(o => <option key={String(o.v)} value={o.v}>{o.label}</option>)}
          </select>
          {form.plAddon==="custom" && <input style={Object.assign({},INP,{marginTop:7})} type="number" placeholder="Custom ₹ amount" value={form.customAddon} onChange={e => setForm(p => ({...p,customAddon:e.target.value}))} />}
        </Fld>

        {prev && (
          <div style={{background:"#F8FAFF",border:"1px solid #BFDBFE",borderRadius:11,padding:"13px",marginBottom:6}}>
            <div style={{fontSize:11,fontWeight:700,color:C.blue,textTransform:"uppercase",letterSpacing:"0.5px",marginBottom:9}}>📊 Live Price Preview</div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:7,marginBottom:9}}>
              <div style={{background:C.ambBg,border:"1px solid #FCD34D",borderRadius:8,padding:"7px 9px",textAlign:"center"}}>
                <div style={{fontSize:9,color:C.amb,fontWeight:700,textTransform:"uppercase",marginBottom:2}}>Inc-GST</div>
                <div style={{fontSize:13,fontWeight:800,color:C.amb}}>{fp(prev.incGST)}</div>
              </div>
              <div style={{background:"#F3F4F6",border:"1px solid "+C.border,borderRadius:8,padding:"7px 9px",textAlign:"center"}}>
                <div style={{fontSize:9,color:C.sec,fontWeight:700,textTransform:"uppercase",marginBottom:2}}>RL Markup</div>
                <div style={{fontSize:13,fontWeight:800}}>{fpct(prev.rlM)}</div>
              </div>
            </div>
            <ProfitRow label="RL" col={C.rl} price={prev.rl} profit={prev.rlProfit} />
            <ProfitRow label="DM" col={C.dm} price={prev.dm} profit={prev.dmProfit} />
            <ProfitRow label="PL" col={C.pl} price={prev.pl} profit={prev.plProfit} />
          </div>
        )}

        <Fld label="Notes (optional)">
          <input style={INP} value={form.notes} placeholder="e.g. seasonal, imported..." onChange={e => setForm(p => ({...p,notes:e.target.value}))} />
        </Fld>
        <Fld label="Status">
          <div style={{display:"flex",gap:8}}>
            {["Active","Inactive"].map(s => {
              const on = (s==="Active"&&form.active)||(s==="Inactive"&&!form.active);
              return (
                <button key={s} onClick={() => setForm(p => ({...p,active:s==="Active"}))}
                  style={{flex:1,padding:"10px",borderRadius:8,border:"1.5px solid "+(on?C.blue:C.border),background:on?"#EFF6FF":"#F9FAFB",color:on?C.blue:C.sec,fontWeight:700,fontSize:13,cursor:"pointer",fontFamily:"inherit"}}>
                  {s}
                </button>
              );
            })}
          </div>
        </Fld>
      </Drawer>
    </div>
  );
}

// ── PRICE LIST VIEW ───────────────────────────────────────────────
function PriceListView({ brands, items, settings, isAdmin }) {
  const [viewMode, setViewMode] = useState("general");
  const [fCat,     setFCat]     = useState("All");
  const [fBrand,   setFBrand]   = useState("All");
  const [search,   setSearch]   = useState("");

  const enriched = useMemo(() => {
    return items.filter(i => i.active).map(i => {
      const b = brands.find(x => x.id===i.bId)||null;
      return Object.assign({}, i, computeItem(i,b,settings), {_b:b});
    });
  }, [items, brands, settings]);

  const filtered = enriched.filter(i => {
    if (fCat!=="All" && i.cat!==fCat) return false;
    if (fBrand!=="All" && i.bId!==fBrand) return false;
    if (search && !i.name.toLowerCase().includes(search.toLowerCase()) && !(i._b&&i._b.name.toLowerCase().includes(search.toLowerCase()))) return false;
    return true;
  });

  const grouped = useMemo(() => {
    const g = {};
    filtered.forEach(i => { if (!g[i.cat]) g[i.cat]=[]; g[i.cat].push(i); });
    return g;
  }, [filtered]);

  const byDate = useMemo(() => {
    const sorted = filtered.slice().sort((a,b) => new Date(b.updatedAt||0) - new Date(a.updatedAt||0));
    const g = {};
    sorted.forEach(i => { const k=dateGroupKey(i.updatedAt); if(!g[k])g[k]=[]; g[k].push(i); });
    return g;
  }, [filtered]);

  const today = new Date().toLocaleDateString("en-IN",{day:"2-digit",month:"short",year:"numeric"});
  const views = [{id:"general",icon:"📂",label:"By Category"},{id:"date",icon:"📅",label:"By Date"}];

  return (
    <div style={{padding:"16px 16px 80px"}}>
      <div style={{background:C.navy,borderRadius:13,padding:"15px",marginBottom:12,color:"#fff"}}>
        <div style={{fontWeight:800,fontSize:17}}>{settings.co}</div>
        <div style={{fontSize:11,color:"rgba(255,255,255,0.5)",marginTop:2}}>{settings.tag}</div>
        <div style={{marginTop:8,fontSize:11,color:"rgba(255,255,255,0.35)"}}>{"Price List · "+today+" · All prices EXCLUDING GST"}</div>
      </div>
      <div style={{background:C.ambBg,border:"1px solid #FCD34D",borderRadius:9,padding:"9px 13px",marginBottom:12,fontSize:12,color:C.amb,fontWeight:600}}>
        ⚠ All prices EXCLUDING GST — GST charged extra.
      </div>
      <div style={{display:"flex",border:"1px solid "+C.border,borderRadius:10,overflow:"hidden",marginBottom:14}}>
        {views.map(v => (
          <button key={v.id} onClick={() => setViewMode(v.id)}
            style={{flex:1,padding:"11px",border:"none",cursor:"pointer",fontWeight:700,fontSize:13,fontFamily:"inherit",display:"flex",alignItems:"center",justifyContent:"center",gap:6,background:viewMode===v.id?C.blue:"#fff",color:viewMode===v.id?"#fff":C.sec}}>
            <span>{v.icon}</span><span>{v.label}</span>
          </button>
        ))}
      </div>
      {viewMode==="date" && (
        <div style={{display:"flex",gap:10,marginBottom:12,flexWrap:"wrap"}}>
          <Chip col={C.new_} bg={C.newBg} br={C.newBr}>🆕 New today</Chip>
          <Chip col={C.upd}  bg={C.updBg} br={C.updBr}>✏️ Price updated</Chip>
        </div>
      )}
      <input style={Object.assign({},INP,{marginBottom:10})} placeholder="🔍 Search by item or brand..." value={search} onChange={e => setSearch(e.target.value)} />
      <div style={{display:"flex",gap:8,marginBottom:12}}>
        <select style={Object.assign({},SEL,{flex:1,padding:"9px 11px"})} value={fCat} onChange={e => setFCat(e.target.value)}>
          <option value="All">All Categories</option>
          {CATS.map(c => <option key={c} value={c}>{c}</option>)}
        </select>
        <select style={Object.assign({},SEL,{flex:1,padding:"9px 11px"})} value={fBrand} onChange={e => setFBrand(e.target.value)}>
          <option value="All">All Brands</option>
          {brands.map(b => <option key={b.id} value={b.id}>{b.code+" — "+b.name}</option>)}
        </select>
      </div>
      <div style={{display:"flex",gap:12,marginBottom:10}}>
        <div style={{fontSize:11,fontWeight:700,color:C.rl}}>🟢 RL = Wholesale</div>
        <div style={{fontSize:11,fontWeight:700,color:C.dm}}>🔵 DM = Sp. Wholesale</div>
        <div style={{fontSize:11,fontWeight:700,color:C.pl}}>🟣 PL = Retail</div>
      </div>
      <div style={{fontSize:12,color:C.mute,marginBottom:10,fontWeight:600}}>{filtered.length+" items"}</div>

      {viewMode==="general" && (
        <div>
          {Object.keys(grouped).length===0 && <div style={{textAlign:"center",padding:"36px",color:C.mute,fontSize:14}}>No items found.</div>}
          {Object.keys(grouped).map(cat => (
            <div key={cat}>
              <div style={{fontSize:12,fontWeight:800,color:C.navy,marginBottom:7,marginTop:4,textTransform:"uppercase",letterSpacing:"0.4px",borderLeft:"3px solid "+C.blue,paddingLeft:9}}>{cat}</div>
              {grouped[cat].map(it => <ItemCard key={it.id} it={it} isAdmin={isAdmin} showDate={false} />)}
            </div>
          ))}
        </div>
      )}

      {viewMode==="date" && (
        <div>
          {Object.keys(byDate).length===0 && <div style={{textAlign:"center",padding:"36px",color:C.mute,fontSize:14}}>No items found.</div>}
          {Object.keys(byDate).map(day => (
            <div key={day}>
              <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:10,marginTop:8}}>
                <div style={{flex:1,height:1,background:C.border}} />
                <div style={{background:C.navy,color:"#fff",borderRadius:20,padding:"4px 14px",fontSize:12,fontWeight:700,whiteSpace:"nowrap"}}>{day}</div>
                <div style={{flex:1,height:1,background:C Here is the fully corrected code with all fixes applied. I removed the duplicate Firebase initialization and the unused `useFirestoreSync` function, and cleaned up the imports.

```jsx
// ─────────────────────────────────────────────────────────────────
//  VK Furnishing Price App  v4.0  — Firebase Edition
// ─────────────────────────────────────────────────────────────────

import { useState, useEffect, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import { initializeApp } from "firebase/app";
import { getFirestore, doc, getDoc, setDoc, onSnapshot } from "firebase/firestore";

// Your web app's Firebase configuration
const firebaseConfig = {
  apiKey: "AIzaSyCEJlH9aGlr0pneb7hT1sIxy1iDnQV3Y4g",
  authDomain: "vkf-price.firebaseapp.com",
  databaseURL: "https://vkf-price-default-rtdb.asia-southeast1.firebasedatabase.app",
  projectId: "vkf-price",
  storageBucket: "vkf-price.firebasestorage.app",
  messagingSenderId: "199170851796",
  appId: "1:199170851796:web:e2e46e279995a41b890c41"
};

// Initialize Firebase once
const app = initializeApp(firebaseConfig);
const db = getFirestore(app);

// Firestore document paths  (all data lives in one doc per collection)
const FS = {
  settings: () => doc(db, "vkf", "settings"),
  brands:   () => doc(db, "vkf", "brands"),
  items:    () => doc(db, "vkf", "items"),
};

// ── FIREBASE STORAGE LAYER ────────────────────────────────────────
async function fsLoad() {
  try {
    const [sSnap, bSnap, iSnap] = await Promise.all([
      getDoc(FS.settings()),
      getDoc(FS.brands()),
      getDoc(FS.items()),
    ]);
    return {
      settings: sSnap.exists() ? sSnap.data().v : null,
      brands:   bSnap.exists() ? bSnap.data().v : null,
      items:    iSnap.exists() ? iSnap.data().v : null,
    };
  } catch (e) {
    console.error("Firebase load error:", e);
    return { settings: null, brands: null, items: null };
  }
}

async function fsSave(key, value) {
  try {
    await setDoc(FS[key](), { v: value, updatedAt: new Date().toISOString() });
  } catch (e) {
    console.error("Firebase save error:", e);
    throw e;
  }
}

// ── CONSTANTS ─────────────────────────────────────────────────────
const VER  = "4.0";
const CATS = ["Bed Sheets","Comforters","Comforter Sets","Towels","Pillows","Dohars","Blankets","Top Sheets","Other Items"];
const GST_OPTS  = [{ label:"5% — Default (Textiles)", v:0.05 },{ label:"18%", v:0.18 }];
const ADDON_OPTS = [
  { label:"Not set",   v:"" },
  { label:"+ ₹100",   v:100 },
  { label:"+ ₹200",   v:200 },
  { label:"+ ₹300",   v:300 },
  { label:"+ ₹500",   v:500 },
  { label:"+ ₹1,000", v:1000 },
  { label:"Custom ₹", v:"custom" },
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
function calcRL(ex, m) { const v=parseFloat(ex||0); return v>0?psychRound(v*(1+parseFloat(m||0.18))):null; }
function calcDM(ex, m) { const v=parseFloat(ex||0); if(m==null||m===""||isNaN(parseFloat(m)))return null; return v>0?psychRound(v*(1+parseFloat(m))):null; }
function calcPL(rl, a) { return (rl!=null&&typeof a==="number"&&!isNaN(a))?psychRound(rl+a):null; }
function calcAddon(pla, ca) {
  if (pla===""||pla==null) return null;
  if (pla==="custom") { const v=parseFloat(ca); return !isNaN(v)&&v>0?v:null; }
  return typeof pla==="number"?pla:null;
}
function calcProfit(sell, ex) {
  if (sell==null||!ex||parseFloat(ex)<=0) return null;
  const amt = sell - parseFloat(ex);
  return { amt:+amt.toFixed(2), pct:+((amt/parseFloat(ex))*100).toFixed(1) };
}
function resolveRL(item, brand, settings) {
  const v = parseFloat(item.customRL);
  if (item.customRL!==""&&item.customRL!=null&&!isNaN(v)) return v/100;
  if (brand?.rlMarkup!=null) return brand.rlMarkup;
  return parseFloat(settings.defaultRL||0.18);
}
function resolveDM(item, brand, settings) {
  const v = parseFloat(item.customDM);
  if (item.customDM!==""&&item.customDM!=null&&!isNaN(v)) return v/100;
  if (brand?.dmMarkup!=null) return brand.dmMarkup;
  if (settings.defaultDM!=null&&settings.defaultDM!=="") return parseFloat(settings.defaultDM);
  return null;
}
function computeItem(item, brand, settings) {
  const rlM = resolveRL(item, brand, settings);
  const dmM = resolveDM(item, brand, settings);
  const add = calcAddon(item.plAddon, item.customAddon);
  const ex  = parseFloat(item.purchaseEx||0);
  const rl  = calcRL(ex, rlM);
  const dm  = calcDM(ex, dmM);
  const pl  = calcPL(rl, add);
  return {
    incGST: calcIncGST(ex, item.gst),
    rlM, rl, rlProfit: calcProfit(rl, ex),
    dmM, dm, dmProfit: calcProfit(dm, ex),
    add, pl, plProfit: calcProfit(pl, ex),
  };
}

// ── DEFAULTS ──────────────────────────────────────────────────────
const DEF_S = { co:"VK Furnishing", tag:"Wholesale Bedding & Textiles, Delhi NCR", defaultRL:0.18, defaultDM:null, adminPIN:"1234", salesPIN:"0000" };
const DEF_B = [
  { id:"b1", code:"TRI", name:"Trident",      rlMarkup:0.22, dmMarkup:null },
  { id:"b2", code:"STH", name:"Story@Home",   rlMarkup:0.18, dmMarkup:null },
  { id:"b3", code:"SWT", name:"Swayam",       rlMarkup:0.20, dmMarkup:null },
  { id:"b4", code:"LOC", name:"Local / Misc", rlMarkup:0.18, dmMarkup:null },
];
const T0 = new Date().toISOString();
const DEF_I = [
  { id:"i1", cat:"Bed Sheets", bId:"b1", name:"King Cotton Bedsheet 200TC",    purchaseEx:380, gst:0.05, customRL:"", customDM:"", plAddon:300,  customAddon:"", active:true, notes:"", createdAt:T0, updatedAt:T0 },
  { id:"i2", cat:"Comforters", bId:"b2", name:"Winter Hollow Fibre Comforter", purchaseEx:550, gst:0.05, customRL:"", customDM:"", plAddon:500,  customAddon:"", active:true, notes:"", createdAt:T0, updatedAt:T0 },
  { id:"i3", cat:"Towels",     bId:"b3", name:"Premium Bath Towel 500 GSM",    purchaseEx:180, gst:0.18, customRL:"", customDM:"", plAddon:200,  customAddon:"", active:true, notes:"", createdAt:T0, updatedAt:T0 },
];

// ── HELPERS ───────────────────────────────────────────────────────
const uid    = () => Date.now().toString(36) + Math.random().toString(36).slice(2,5);
const fp     = (n) => (n!=null&&!isNaN(n)) ? "₹"+Number(n).toLocaleString("en-IN") : "—";
const fpct   = (n) => (n!=null&&n!==""&&!isNaN(parseFloat(n))) ? (parseFloat(n)*100).toFixed(0)+"%" : "—";
const fpctRaw= (s) => (s!==""&&s!=null&&!isNaN(parseFloat(s))) ? parseFloat(s).toFixed(1)+"%" : null;

function fmtDate(iso) {
  if (!iso) return "";
  return new Date(iso).toLocaleDateString("en-IN",{day:"2-digit",month:"short",year:"numeric"});
}
function fmtDateTime(iso) {
  if (!iso) return "";
  const d = new Date(iso);
  return fmtDate(iso)+" "+d.toLocaleTimeString("en-IN",{hour:"2-digit",minute:"2-digit",hour12:true});
}
function dateGroupKey(iso) {
  if (!iso) return "Unknown";
  const d     = new Date(iso);
  const today = new Date();
  const yest  = new Date(); yest.setDate(today.getDate()-1);
  if (d.toDateString()===today.toDateString()) return "Today";
  if (d.toDateString()===yest.toDateString())  return "Yesterday";
  return d.toLocaleDateString("en-IN",{weekday:"long",day:"2-digit",month:"short",year:"numeric"});
}
function sameDay(a,b) { return new Date(a).toDateString()===new Date(b).toDateString(); }

// ── PALETTE ───────────────────────────────────────────────────────
const C = {
  bg:"#F0F2F7", card:"#fff", border:"#E5E8EF", text:"#111827", sec:"#6B7280", mute:"#9CA3AF",
  blue:"#1648D6", navy:"#0B1D5C",
  rl:"#067D62", rlBg:"#ECFDF5", rlBr:"#6EE7B7",
  dm:"#1C64F2", dmBg:"#EFF6FF", dmBr:"#93C5FD",
  pl:"#6D28D9", plBg:"#F5F3FF", plBr:"#C4B5FD",
  profit:"#065F46", profBg:"#F0FDF4",
  amb:"#B45309", ambBg:"#FEF3C7",
  red:"#DC2626", redBg:"#FEF2F2",
  new_:"#0E7490", newBg:"#ECFEFF", newBr:"#A5F3FC",
  upd:"#7C3AED", updBg:"#F5F3FF", updBr:"#DDD6FE",
  sync:"#059669", syncBg:"#ECFDF5",
};

// ── TOAST ─────────────────────────────────────────────────────────
let _toast = null;
const toast = (m, t="ok") => _toast && _toast({id:uid(), m, t});
function ToastHost() {
  const [list, setList] = useState([]);
  useEffect(() => {
    _toast = (x) => {
      setList(p => [...p, x]);
      setTimeout(() => setList(p => p.filter(y => y.id !== x.id)), 2600);
    };
  }, []);
  const BG = { ok:"#059669", err:C.red, warn:"#D97706" };
  return (
    <div style={{position:"fixed",top:14,right:14,zIndex:9999,display:"flex",flexDirection:"column",gap:8}}>
      {list.map(x => (
        <div key={x.id} style={{padding:"11px 16px",borderRadius:10,fontSize:13,fontWeight:700,color:"#fff",background:BG[x.t]||BG.ok,boxShadow:"0 4px 16px rgba(0,0,0,0.2)"}}>
          {x.m}
        </div>
      ))}
    </div>
  );
}

// ── SYNC STATUS BADGE ─────────────────────────────────────────────
function SyncBadge({ status }) {
  // status: "synced" | "saving" | "error" | "offline"
  const cfg = {
    synced:  { col:"#059669", bg:"#ECFDF5", br:"#6EE7B7", icon:"☁️", label:"Synced"   },
    saving:  { col:"#D97706", bg:"#FEF3C7", br:"#FCD34D", icon:"⏳", label:"Saving…"  },
    error:   { col:C.red,     bg:C.redBg,   br:"#FCA5A5", icon:"⚠️", label:"Sync error"},
    offline: { col:C.sec,     bg:"#F3F4F6", br:C.border,  icon:"📵", label:"Offline"  },
  }[status] || { col:C.sec, bg:"#F3F4F6", br:C.border, icon:"…", label:status };
  return (
    <span style={{display:"inline-flex",alignItems:"center",gap:4,padding:"3px 9px",borderRadius:20,fontSize:11,fontWeight:700,color:cfg.col,background:cfg.bg,border:"1px solid "+cfg.br}}>
      {cfg.icon} {cfg.label}
    </span>
  );
}

// ── UI ATOMS ──────────────────────────────────────────────────────
function Chip({ col, bg, br, children }) {
  return <span style={{display:"inline-block",padding:"2px 9px",borderRadius:20,fontSize:11,fontWeight:700,color:col,background:bg,border:"1px solid "+br}}>{children}</span>;
}
function PBox({ label, val, col, bg, br }) {
  return (
    <div style={{flex:1,background:bg,border:"1px solid "+br,borderRadius:10,padding:"9px 6px",textAlign:"center"}}>
      <div style={{fontSize:9,fontWeight:700,color:C.mute,letterSpacing:"0.5px",textTransform:"uppercase",marginBottom:3}}>{label}</div>
      <div style={{fontSize:15,fontWeight:800,color:col}}>{val}</div>
    </div>
  );
}
function Fld({ label, hint, err, children }) {
  return (
    <div style={{marginBottom:14}}>
      <label style={{display:"block",fontSize:11,fontWeight:700,color:C.sec,letterSpacing:"0.7px",textTransform:"uppercase",marginBottom:5}}>{label}</label>
      {children}
      {err  && <div style={{fontSize:11,color:C.red,marginTop:3}}>{err}</div>}
      {!err && hint && <div style={{fontSize:11,color:C.mute,marginTop:3}}>{hint}</div>}
    </div>
  );
}
const INP = {width:"100%",padding:"11px 13px",borderRadius:9,fontSize:14,border:"1.5px solid #E5E8EF",background:"#FAFBFD",fontFamily:"inherit",outline:"none",boxSizing:"border-box",color:"#111827"};
const SEL = Object.assign({}, INP, {cursor:"pointer"});
function BtnP({ color, onClick, children, style }) {
  return <button onClick={onClick} style={Object.assign({padding:"13px 20px",borderRadius:10,border:"none",background:color||C.blue,color:"#fff",fontWeight:700,fontSize:14,cursor:"pointer",fontFamily:"inherit",width:"100%"},style||{})}>{children}</button>;
}
function BtnO({ color, onClick, children, style }) {
  const col = color||C.sec;
  return <button onClick={onClick} style={Object.assign({padding:"13px 20px",borderRadius:10,border:"1.5px solid "+col,background:"#fff",color:col,fontWeight:600,fontSize:14,cursor:"pointer",fontFamily:"inherit",width:"100%"},style||{})}>{children}</button>;
}

// ── MARGIN OVERRIDE INPUT ─────────────────────────────────────────
function MarginOverride({ label, col, bg, br, value, onChange, srcLabel, err }) {
  const has = value!==""&&value!=null&&!isNaN(parseFloat(value));
  return (
    <div style={{marginBottom:14}}>
      <label style={{display:"block",fontSize:11,fontWeight:700,color:C.sec,letterSpacing:"0.7px",textTransform:"uppercase",marginBottom:5}}>{label}</label>
      <div style={{position:"relative"}}>
        <input style={Object.assign({},INP,{borderColor:err?C.red:has?col:C.border,paddingRight:has?"90px":"13px"})}
          type="number" step="0.5" min="0" max="100"
          placeholder={"Leave blank = "+srcLabel} value={value}
          onChange={e => onChange(e.target.value)} />
        {has && (
          <div style={{position:"absolute",right:10,top:"50%",transform:"translateY(-50%)",background:bg,border:"1px solid "+br,color:col,fontSize:11,fontWeight:700,padding:"2px 8px",borderRadius:6}}>
            {parseFloat(value).toFixed(1)+"%"}
          </div>
        )}
      </div>
      {err && <div style={{fontSize:11,color:C.red,marginTop:3}}>{err}</div>}
      {!err && (
        <div style={{fontSize:11,color:has?col:C.mute,fontWeight:has?600:400,marginTop:3}}>
          {has ? "⚡ Item override active — brand/default ignored" : "Using: "+srcLabel}
        </div>
      )}
    </div>
  );
}

// ── PROFIT ROW ────────────────────────────────────────────────────
function ProfitRow({ label, col, price, profit }) {
  if (!price) return null;
  return (
    <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"5px 10px",background:C.profBg,borderRadius:7,marginBottom:4,border:"1px solid #BBF7D0"}}>
      <div style={{display:"flex",alignItems:"center",gap:8}}>
        <span style={{fontSize:11,fontWeight:700,color:col}}>{label}</span>
        <span style={{fontSize:13,fontWeight:800,color:col}}>{fp(price)}</span>
      </div>
      {profit ? (
        <div style={{display:"flex",alignItems:"center",gap:6}}>
          <span style={{fontSize:11,fontWeight:700,color:C.profit}}>{"+"+fp(profit.amt)}</span>
          <span style={{background:"#D1FAE5",color:C.profit,fontSize:11,fontWeight:800,padding:"2px 8px",borderRadius:20}}>{profit.pct+"%"}</span>
        </div>
      ) : (
        <span style={{fontSize:11,color:C.mute}}>—</span>
      )}
    </div>
  );
}

// ── ITEM CARD ─────────────────────────────────────────────────────
function ItemCard({ it, isAdmin, showDate }) {
  const isNew    = it.createdAt && it.updatedAt && sameDay(it.createdAt, it.updatedAt);
  const isEdited = it.createdAt && it.updatedAt && !sameDay(it.createdAt, it.updatedAt);
  return (
    <div style={{background:C.card,border:"1px solid "+C.border,borderRadius:12,padding:"13px",marginBottom:8,boxShadow:"0 1px 5px rgba(0,0,0,0.04)"}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:9}}>
        <div style={{flex:1,paddingRight:8}}>
          <div style={{fontWeight:700,fontSize:14,marginBottom:4}}>{it.name}</div>
          <div style={{display:"flex",gap:6,flexWrap:"wrap",alignItems:"center"}}>
            {it._b && <div style={{background:C.navy,color:"#fff",borderRadius:5,padding:"1px 8px",fontSize:10,fontWeight:800}}>{it._b.code}</div>}
            <Chip col={C.sec} bg="#F3F4F6" br={C.border}>{it.cat}</Chip>
            {isNew    && <Chip col={C.new_} bg={C.newBg} br={C.newBr}>🆕 New</Chip>}
            {isEdited && <Chip col={C.upd}  bg={C.updBg} br={C.updBr}>✏️ Updated</Chip>}
          </div>
          {showDate && (
            <div style={{marginTop:5,fontSize:10,color:C.mute}}>
              {isNew ? "Added "+fmtDateTime(it.createdAt) : "Updated "+fmtDateTime(it.updatedAt)}
              {isEdited && <span style={{marginLeft:6,color:C.updBr}}>· Added {fmtDate(it.createdAt)}</span>}
            </div>
          )}
        </div>
        <Chip col={C.amb} bg={C.ambBg} br="#FCD34D">{"GST "+(it.gst*100).toFixed(0)+"%"}</Chip>
      </div>
      {isAdmin && (
        <div style={{background:C.ambBg,border:"1px solid #FCD34D",borderRadius:8,padding:"7px 11px",marginBottom:9,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
          <span style={{fontSize:11,fontWeight:700,color:C.amb,textTransform:"uppercase"}}>Purchase Ex-GST</span>
          <span style={{fontWeight:800,fontSize:13,color:C.amb}}>
            {fp(it.purchaseEx)+" + "}
            <span style={{fontSize:10,fontWeight:600}}>{(it.gst*100).toFixed(0)+"% = "+fp(it.incGST)}</span>
          </span>
        </div>
      )}
      <div style={{display:"flex",gap:7}}>
        <PBox label="RL — Wholesale" val={fp(it.rl)} col={C.rl} bg={C.rlBg} br={C.rlBr} />
        <PBox label="DM — Sp. WHL"   val={fp(it.dm)} col={C.dm} bg={C.dmBg} br={C.dmBr} />
        <PBox label="PL — Retail"     val={fp(it.pl)} col={C.pl} bg={C.plBg} br={C.plBr} />
      </div>
      {isAdmin && (
        <div style={{marginTop:9}}>
          <div style={{fontSize:10,fontWeight:700,color:C.profit,textTransform:"uppercase",letterSpacing:"0.5px",marginBottom:5}}>💰 Profit per piece (Ex-GST basis)</div>
          <ProfitRow label="RL" col={C.rl} price={it.rl} profit={it.rlProfit} />
          <ProfitRow label="DM" col={C.dm} price={it.dm} profit={it.dmProfit} />
          <ProfitRow label="PL" col={C.pl} price={it.pl} profit={it.plProfit} />
        </div>
      )}
      {it.notes && <div style={{marginTop:7,fontSize:11,color:C.mute,fontStyle:"italic"}}>{"📝 "+it.notes}</div>}
    </div>
  );
}

// ── DRAWER ────────────────────────────────────────────────────────
function Drawer({ open, onClose, title, footer, children }) {
  if (!open) return null;
  return (
    <div>
      <div onClick={onClose} style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.45)",zIndex:500}} />
      <div style={{position:"fixed",bottom:0,left:0,right:0,zIndex:501,background:"#fff",borderRadius:"22px 22px 0 0",maxHeight:"92vh",display:"flex",flexDirection:"column",boxShadow:"0 -8px 40px rgba(0,0,0,0.18)"}}>
        <div style={{display:"flex",justifyContent:"center",padding:"10px 0 0"}}>
          <div style={{width:36,height:4,borderRadius:2,background:C.border}} />
        </div>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"8px 20px 12px"}}>
          <div style={{fontSize:18,fontWeight:800,color:C.text}}>{title}</div>
          <button onClick={onClose} style={{width:34,height:34,borderRadius:8,border:"1.5px solid "+C.border,background:"#F9FAFB",cursor:"pointer",fontSize:16,color:C.sec,fontFamily:"inherit"}}>✕</button>
        </div>
        <div style={{flex:1,overflowY:"auto",padding:"0 20px 6px"}}>{children}</div>
        {footer && <div style={{padding:"12px 20px",borderTop:"1px solid "+C.border}}>{footer}</div>}
      </div>
    </div>
  );
}

// ── CONFIRM ───────────────────────────────────────────────────────
function Confirm({ title, msg, onOk, onNo }) {
  return (
    <div style={{position:"fixed",inset:0,zIndex:600,background:"rgba(0,0,0,0.5)",display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <div style={{background:"#fff",borderRadius:18,padding:"26px 22px",maxWidth:320,width:"100%"}}>
        <div style={{fontSize:17,fontWeight:800,color:C.text,marginBottom:8}}>{title}</div>
        <div style={{fontSize:13,color:C.sec,lineHeight:1.7,marginBottom:24}}>{msg}</div>
        <div style={{display:"flex",gap:10}}>
          <BtnO onClick={onNo}>Cancel</BtnO>
          <BtnP color={C.red} onClick={onOk}>Delete</BtnP>
        </div>
      </div>
    </div>
  );
}

// ── PIN PAD ───────────────────────────────────────────────────────
function PINPad({ label, pin, onSuccess }) {
  const [val, setVal] = useState("");
  const [err, setErr] = useState("");
  const [shk, setShk] = useState(false);

  function press(d) {
    if (val.length >= 4) return;
    const next = val + d;
    setVal(next);
    if (next.length === 4) {
      if (next === String(pin)) {
        onSuccess();
      } else {
        setShk(true); setErr("Wrong PIN — try again.");
        setTimeout(() => { setVal(""); setShk(false); setErr(""); }, 700);
      }
    }
  }

  const keys = ["1","2","3","4","5","6","7","8","9","","0","⌫"];
  return (
    <div style={{textAlign:"center"}}>
      <div style={{fontSize:13,fontWeight:700,color:C.sec,marginBottom:16}}>{label}</div>
      <div style={{display:"flex",justifyContent:"center",gap:14,marginBottom:6}}>
        {[0,1,2,3].map(i => (
          <div key={i} style={{width:15,height:15,borderRadius:"50%",background:val.length>i?C.blue:"transparent",border:"2.5px solid "+(val.length>i?C.blue:C.border),transition:"all 0.12s",animation:shk?"shk 0.35s ease":"none"}} />
        ))}
      </div>
      <div style={{minHeight:22,fontSize:12,color:C.red,fontWeight:600,marginBottom:8}}>{err}</div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:10,maxWidth:220,margin:"0 auto"}}>
        {keys.map((k,i) => (
          <button key={i} disabled={k===""} onClick={() => k==="⌫" ? setVal(p=>p.slice(0,-1)) : k ? press(k) : null}
            style={{height:56,borderRadius:12,fontSize:k==="⌫"?20:22,fontWeight:700,cursor:k===""?"default":"pointer",fontFamily:"inherit",background:k===""?"transparent":"#fff",border:k===""?"none":"1.5px solid "+C.border,color:k==="⌫"?C.red:C.text,boxShadow:k!==""?"0 2px 6px rgba(0,0,0,0.07)":"none"}}>
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
    <div style={{minHeight:"100vh",background:"linear-gradient(160deg,"+C.navy+" 0%,"+C.blue+" 100%)",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",padding:20}}>
      <div style={{marginBottom:30,textAlign:"center"}}>
        <div style={{width:66,height:66,borderRadius:18,background:"rgba(255,255,255,0.13)",border:"2px solid rgba(255,255,255,0.2)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:24,fontWeight:900,color:"#fff",margin:"0 auto 12px"}}>VK</div>
        <div style={{fontSize:20,fontWeight:800,color:"#fff"}}>{settings.co}</div>
        <div style={{fontSize:12,color:"rgba(255,255,255,0.4)",marginTop:3}}>{settings.tag}</div>
      </div>
      <div style={{background:"#fff",borderRadius:22,padding:"28px 24px",width:"100%",maxWidth:340,boxShadow:"0 30px 80px rgba(0,0,0,0.3)"}}>
        {!role ? (
          <div>
            <div style={{fontSize:16,fontWeight:800,color:C.text,textAlign:"center",marginBottom:18}}>Select your role</div>
            {roles.map(o => (
              <button key={o.r} onClick={() => setRole(o.r)}
                style={{display:"flex",alignItems:"center",gap:14,padding:"14px 16px",borderRadius:12,border:"1.5px solid "+C.border,background:"#FAFBFD",cursor:"pointer",width:"100%",fontFamily:"inherit",marginBottom:10,textAlign:"left"}}>
                <span style={{fontSize:24}}>{o.icon}</span>
                <div style={{flex:1}}>
                  <div style={{fontWeight:700,fontSize:14,color:C.text}}>{o.t}</div>
                  <div style={{fontSize:12,color:C.sec,marginTop:1}}>{o.d}</div>
                </div>
                <span style={{color:C.mute,fontSize:18}}>›</span>
              </button>
            ))}
          </div>
        ) : (
          <div>
            <button onClick={() => setRole(null)} style={{background:"none",border:"none",cursor:"pointer",color:C.sec,fontSize:13,fontWeight:600,marginBottom:16,padding:0,fontFamily:"inherit"}}>← Back</button>
            <PINPad
              label={"Enter "+(role==="admin"?"Admin":"Salesman")+" PIN"}
              pin={role==="admin"?settings.adminPIN:settings.salesPIN}
              onSuccess={() => onLogin(role)}
            />
          </div>
        )}
      </div>
      <div style={{marginTop:16,fontSize:11,color:"rgba(255,255,255,0.25)"}}>{"v"+VER+" · Firebase sync enabled"}</div>
    </div>
  );
}

// ── BRANDS VIEW ───────────────────────────────────────────────────
function BrandsView({ brands, onBrandsChange }) {
  const EF = { code:"", name:"", rlMarkup:"18", dmMarkup:"" };
  const [form, setForm] = useState(EF);
  const [eid,  setEid]  = useState(null);
  const [open, setOpen] = useState(false);
  const [conf, setConf] = useState(null);

  function close() { setOpen(false); setEid(null); setForm(EF); }
  function startEdit(b) {
    setForm({ code:b.code, name:b.name, rlMarkup:+(b.rlMarkup*100).toFixed(1), dmMarkup:b.dmMarkup!=null?+(b.dmMarkup*100).toFixed(1):"" });
    setEid(b.id); setOpen(true);
  }
  function save() {
    if (!form.code.trim()) return toast("Brand code required","err");
    if (!form.name.trim()) return toast("Brand name required","err");
    if (isNaN(parseFloat(form.rlMarkup))) return toast("RL markup must be a number","err");
    const entry = {
      id: eid||uid(), code:form.code.trim().toUpperCase(), name:form.name.trim(),
      rlMarkup: parseFloat(form.rlMarkup)/100,
      dmMarkup: form.dmMarkup!==""&&!isNaN(parseFloat(form.dmMarkup)) ? parseFloat(form.dmMarkup)/100 : null,
    };
    const next = eid ? brands.map(b => b.id===eid?entry:b) : [...brands, entry];
    onBrandsChange(next);
    toast(eid?"Brand updated":"Brand added"); close();
  }
  function del(id) { onBrandsChange(brands.filter(b => b.id!==id)); setConf(null); toast("Brand removed","warn"); }

  return (
    <div style={{padding:"16px 16px 80px"}}>
      {conf && <Confirm title="Remove Brand?" msg="Items using this brand will keep prices but lose markup override." onOk={() => del(conf)} onNo={() => setConf(null)} />}
      <div style={{background:"#EFF6FF",border:"1px solid #BFDBFE",borderRadius:10,padding:"10px 14px",marginBottom:14,fontSize:12,color:"#1E40AF",lineHeight:1.6}}>
        <strong>Brand markups are defaults.</strong> You can override RL and DM per item in Master Sheet.
      </div>
      <BtnP onClick={() => { setForm(EF); setEid(null); setOpen(true); }}>+ Add Brand</BtnP>
      <div style={{height:12}} />
      {brands.length===0 && <div style={{textAlign:"center",padding:"36px",color:C.mute,fontSize:14}}>No brands yet.</div>}
      {brands.map(b => (
        <div key={b.id} style={{background:C.card,border:"1px solid "+C.border,borderRadius:13,padding:"14px",marginBottom:10,boxShadow:"0 1px 6px rgba(0,0,0,0.05)"}}>
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:8}}>
            <div style={{display:"flex",alignItems:"center",gap:9}}>
              <div style={{background:C.navy,color:"#fff",borderRadius:7,padding:"4px 10px",fontWeight:800,fontSize:12}}>{b.code}</div>
              <div style={{fontWeight:700,fontSize:14}}>{b.name}</div>
            </div>
            <div style={{display:"flex",gap:7}}>
              <button onClick={() => startEdit(b)} style={{padding:"5px 13px",borderRadius:7,border:"1.5px solid "+C.border,background:"#F9FAFB",cursor:"pointer",fontSize:12,fontWeight:600,color:C.sec,fontFamily:"inherit"}}>Edit</button>
              <button onClick={() => setConf(b.id)} style={{padding:"5px 10px",borderRadius:7,border:"1.5px solid #FCA5A5",background:C.redBg,cursor:"pointer",fontSize:13,color:C.red,fontFamily:"inherit"}}>✕</button>
            </div>
          </div>
          <div style={{display:"flex",gap:7}}>
            <Chip col={C.rl} bg={C.rlBg} br={C.rlBr}>{"RL "+fpct(b.rlMarkup)}</Chip>
            {b.dmMarkup!=null
              ? <Chip col={C.dm} bg={C.dmBg} br={C.dmBr}>{"DM "+fpct(b.dmMarkup)}</Chip>
              : <Chip col={C.mute} bg="#F3F4F6" br={C.border}>DM not set</Chip>}
          </div>
        </div>
      ))}
      <Drawer open={open} onClose={close} title={eid?"Edit Brand":"Add Brand"}
        footer={
          <div style={{display:"flex",gap:9}}>
            <BtnO onClick={close}>Cancel</BtnO>
            <BtnP color={eid?"#F59E0B":C.blue} onClick={save}>{eid?"Update":"Add Brand"}</BtnP>
          </div>
        }>
        <Fld label="Brand Code *" hint="3–6 letters e.g. TRI, STH, LOC">
          <input style={INP} maxLength={6} value={form.code} placeholder="TRI" onChange={e => setForm(p => ({...p,code:e.target.value.toUpperCase()}))} />
        </Fld>
        <Fld label="Brand Name *">
          <input style={INP} value={form.name} placeholder="e.g. Trident" onChange={e => setForm(p => ({...p,name:e.target.value}))} />
        </Fld>
        <Fld label="Default RL Markup % *" hint="Enter number: 18 = 18%">
          <input style={INP} type="number" step="0.5" value={form.rlMarkup} placeholder="18" onChange={e => setForm(p => ({...p,rlMarkup:e.target.value}))} />
        </Fld>
        <Fld label="Default DM Markup %" hint="Enter number: 10 = 10%. Leave blank to decide later.">
          <input style={INP} type="number" step="0.5" value={form.dmMarkup} placeholder="Leave blank" onChange={e => setForm(p => ({...p,dmMarkup:e.target.value}))} />
        </Fld>
      </Drawer>
    </div>
  );
}

// ── EXPORT MODAL ──────────────────────────────────────────────────
function ExportModal({ items, brands, settings, onClose }) {
  const [fCat,   setFCat]   = useState("All");
  const [fBrand, setFBrand] = useState("All");
  const [fSt,    setFSt]    = useState("Active");

  const filtered = useMemo(() => {
    return items.filter(i => {
      if (fCat!=="All" && i.cat!==fCat) return false;
      if (fBrand!=="All" && i.bId!==fBrand) return false;
      if (fSt==="Active" && !i.active) return false;
      if (fSt==="Inactive" && i.active) return false;
      return true;
    }).map(i => {
      const b = brands.find(x => x.id===i.bId)||null;
      return Object.assign({}, i, computeItem(i,b,settings), {_b:b});
    });
  }, [items, brands, settings, fCat, fBrand, fSt]);

  const today = new Date().toLocaleDateString("en-IN",{day:"2-digit",month:"short",year:"numeric"});

  function doXLSX(rows, fname) {
    if (!rows.length) return toast("No items to export","warn");
    try {
      const ws = XLSX.utils.json_to_sheet(rows);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Price List");
      XLSX.writeFile(wb, fname+"_"+today+".xlsx");
      toast("Excel downloaded!");
    } catch { toast("Export failed","err"); }
  }
  function doJSON(rows, fname) {
    if (!rows.length) return toast("No items to export","warn");
    const blob = new Blob([JSON.stringify({exportedAt:new Date().toISOString(),data:rows},null,2)],{type:"application/json"});
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob); a.download = fname+"_"+today+".json"; a.click();
    toast("JSON downloaded!");
  }
  function adminRows() {
    return filtered.map(it => ({
      "Category":it.cat, "Brand Code":it._b?it._b.code:"", "Brand Name":it._b?it._b.name:"", "Item Name":it.name,
      "Purchase Ex-GST":it.purchaseEx, "GST %":(it.gst*100).toFixed(0)+"%", "Purchase Inc-GST":it.incGST,
      "RL Markup":fpct(it.rlM), "RL Price":it.rl!=null?it.rl:"", "RL Profit Rs":it.rlProfit?it.rlProfit.amt:"", "RL Profit %":it.rlProfit?it.rlProfit.pct:"",
      "DM Markup":it.dmM!=null?fpct(it.dmM):"", "DM Price":it.dm!=null?it.dm:"", "DM Profit Rs":it.dmProfit?it.dmProfit.amt:"", "DM Profit %":it.dmProfit?it.dmProfit.pct:"",
      "PL Addon":it.add!=null?it.add:"", "PL Price":it.pl!=null?it.pl:"", "PL Profit Rs":it.plProfit?it.plProfit.amt:"", "PL Profit %":it.plProfit?it.plProfit.pct:"",
      "Status":it.active?"Active":"Inactive", "Added On":fmtDate(it.createdAt), "Updated On":fmtDate(it.updatedAt), "Notes":it.notes,
    }));
  }
  function salesRows() {
    return filtered.map(it => ({
      "Category":it.cat, "Brand Code":it._b?it._b.code:"", "Brand Name":it._b?it._b.name:"", "Item Name":it.name,
      "GST %":(it.gst*100).toFixed(0)+"%", "RL Price":it.rl!=null?it.rl:"", "DM Price":it.dm!=null?it.dm:"", "PL Price":it.pl!=null?it.pl:"", "Notes":it.notes,
    }));
  }

  return (
    <div style={{position:"fixed",inset:0,zIndex:600,background:"rgba(0,0,0,0.5)",display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
      <div style={{background:"#fff",borderRadius:20,padding:"22px 18px",maxWidth:400,width:"100%",maxHeight:"88vh",overflowY:"auto"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
          <div style={{fontSize:17,fontWeight:800}}>⬇ Export Price List</div>
          <button onClick={onClose} style={{width:32,height:32,borderRadius:8,border:"1.5px solid "+C.border,background:"#F9FAFB",cursor:"pointer",fontSize:15,color:C.sec,fontFamily:"inherit"}}>✕</button>
        </div>
        <div style={{background:"#F8FAFF",border:"1px solid #BFDBFE",borderRadius:10,padding:"13px",marginBottom:14}}>
          <div style={{fontSize:11,fontWeight:700,color:C.blue,textTransform:"uppercase",letterSpacing:"0.6px",marginBottom:10}}>Filter Items to Export</div>
          <Fld label="Category">
            <select style={SEL} value={fCat} onChange={e => setFCat(e.target.value)}>
              <option value="All">All Categories</option>
              {CATS.map(c => <option key={c} value={c}>{c}</option>)}
            </select>
          </Fld>
          <Fld label="Brand">
            <select style={SEL} value={fBrand} onChange={e => setFBrand(e.target.value)}>
              <option value="All">All Brands</option>
              {brands.map(b => <option key={b.id} value={b.id}>{b.code+" — "+b.name}</option>)}
            </select>
          </Fld>
          <Fld label="Status">
            <div style={{display:"flex",border:"1px solid "+C.border,borderRadius:9,overflow:"hidden"}}>
              {["All","Active","Inactive"].map(s => (
                <button key={s} onClick={() => setFSt(s)} style={{flex:1,padding:"9px",border:"none",cursor:"pointer",fontWeight:600,fontSize:12,fontFamily:"inherit",background:fSt===s?C.blue:"#fff",color:fSt===s?"#fff":C.sec}}>{s}</button>
              ))}
            </div>
          </Fld>
          <div style={{fontSize:12,color:C.mute,fontWeight:600}}>{filtered.length+" items selected"}</div>
        </div>
        <div style={{background:C.ambBg,border:"1px solid #FCD34D",borderRadius:10,padding:"13px",marginBottom:10}}>
          <div style={{fontSize:12,fontWeight:800,color:C.amb,marginBottom:3}}>🔐 Admin Export</div>
          <div style={{fontSize:11,color:C.amb,marginBottom:10}}>Includes purchase price + profit. Do NOT share with staff.</div>
          <div style={{display:"flex",gap:8}}>
            <BtnP color={C.amb} onClick={() => doXLSX(adminRows(),"VKF_Admin")} style={{fontSize:12,padding:"10px"}}>📊 Excel</BtnP>
            <BtnO color={C.amb} onClick={() => doJSON(adminRows(),"VKF_Admin")} style={{fontSize:12,padding:"10px"}}>JSON</BtnO>
          </div>
        </div>
        <div style={{background:C.rlBg,border:"1px solid "+C.rlBr,borderRadius:10,padding:"13px"}}>
          <div style={{fontSize:12,fontWeight:800,color:C.rl,marginBottom:3}}>📋 Salesman Export</div>
          <div style={{fontSize:11,color:C.rl,marginBottom:10}}>Prices only — no costs. Safe to share.</div>
          <div style={{display:"flex",gap:8}}>
            <BtnP color={C.rl} onClick={() => doXLSX(salesRows(),"VKF_PriceList")} style={{fontSize:12,padding:"10px"}}>📊 Excel</BtnP>
            <BtnO color={C.rl} onClick={() => doJSON(salesRows(),"VKF_PriceList")} style={{fontSize:12,padding:"10px"}}>JSON</BtnO>
          </div>
        </div>
      </div>
    </div>
  );
}

// ── MASTER VIEW ───────────────────────────────────────────────────
const EI = { cat:"Bed Sheets",bId:"",name:"",purchaseEx:"",gst:0.05,customRL:"",customDM:"",plAddon:"",customAddon:"",active:true,notes:"" };

function MasterView({ brands, items, onItemsChange, settings }) {
  const [open,    setOpen]    = useState(false);
  const [eid,     setEid]     = useState(null);
  const [form,    setForm]    = useState(EI);
  const [errs,    setErrs]    = useState({});
  const [fCat,    setFCat]    = useState("All");
  const [fSt,     setFSt]     = useState("Active");
  const [conf,    setConf]    = useState(null);
  const [showExp, setShowExp] = useState(false);

  const brand  = brands.find(b => b.id===form.bId)||null;
  const prevEx = parseFloat(form.purchaseEx)||0;
  const prev   = prevEx>0 ? computeItem(Object.assign({},form,{purchaseEx:prevEx}), brand, settings) : null;

  function close() { setOpen(false); setEid(null); }
  function openAdd() { setForm(EI); setEid(null); setErrs({}); setOpen(true); }
  function openEdit(it) { setForm(Object.assign({},it,{customDM:it.customDM||""})); setEid(it.id); setErrs({}); setOpen(true); }

  function validate() {
    const e = {};
    if (!form.name.trim()) e.name = "Item name required";
    const px = parseFloat(form.purchaseEx);
    if (!form.purchaseEx||isNaN(px)||px<=0) e.px = "Enter a valid purchase price";
    if (form.customRL!==""&&form.customRL!=null) { const v=parseFloat(form.customRL); if(isNaN(v)||v<0||v>100) e.crl="Enter 0–100 (e.g. 15 for 15%)"; }
    if (form.customDM!==""&&form.customDM!=null) { const v=parseFloat(form.customDM); if(isNaN(v)||v<0||v>100) e.cdm="Enter 0–100 (e.g. 10 for 10%)"; }
    if (form.plAddon==="custom") { const cv=parseFloat(form.customAddon); if(isNaN(cv)||cv<=0) e.ca="Enter a valid amount"; }
    setErrs(e);
    return Object.keys(e).length===0;
  }

  function save() {
    if (!validate()) return;
    const now = new Date().toISOString();
    const orig = eid ? items.find(i => i.id===eid) : null;
    const entry = Object.assign({}, form, {
      id: eid||uid(),
      purchaseEx: parseFloat(form.purchaseEx),
      gst: parseFloat(form.gst),
      createdAt: orig ? (orig.createdAt||now) : now,
      updatedAt: now,
    });
    const next = eid ? items.map(i => i.id===eid?entry:i) : [...items, entry];
    onItemsChange(next);
    toast(eid?"Item updated":"Item added");
    close();
  }
  function del(id) { onItemsChange(items.filter(i => i.id!==id)); setConf(null); toast("Item removed","warn"); }
  function toggle(id) {
    onItemsChange(items.map(i => i.id===id ? Object.assign({},i,{active:!i.active,updatedAt:new Date().toISOString()}) : i));
  }

  const list = useMemo(() => {
    return items.filter(i => {
      if (fCat!=="All" && i.cat!==fCat) return false;
      if (fSt==="Active" && !i.active) return false;
      if (fSt==="Inactive" && i.active) return false;
      return true;
    }).map(i => {
      const b = brands.find(x => x.id===i.bId)||null;
      return Object.assign({}, i, computeItem(i,b,settings), {_b:b});
    });
  }, [items, brands, settings, fCat, fSt]);

  function ei(k) { return errs[k] ? Object.assign({},INP,{borderColor:C.red}) : INP; }

  const rlSrc = fpctRaw(form.customRL)
    ? "item override ("+fpctRaw(form.customRL)+")"
    : brand ? "brand "+brand.code+" ("+(brand.rlMarkup*100).toFixed(0)+"%)"
    : "default ("+(settings.defaultRL*100).toFixed(0)+"%)";
  const dmSrc = fpctRaw(form.customDM)
    ? "item override ("+fpctRaw(form.customDM)+")"
    : brand&&brand.dmMarkup!=null ? "brand "+brand.code+" ("+(brand.dmMarkup*100).toFixed(0)+"%)"
    : settings.defaultDM ? "default ("+(parseFloat(settings.defaultDM)*100).toFixed(0)+"%)"
    : "not set";

  return (
    <div style={{padding:"16px 16px 80px"}}>
      {conf && <Confirm title="Delete Item?" msg="This will permanently remove the item." onOk={() => del(conf)} onNo={() => setConf(null)} />}
      {showExp && <ExportModal items={items} brands={brands} settings={settings} onClose={() => setShowExp(false)} />}

      <div style={{background:C.ambBg,border:"1px solid #FCD34D",borderRadius:10,padding:"9px 13px",marginBottom:12,fontSize:12,color:C.amb,fontWeight:600}}>
        ⚠ Enter purchase price WITHOUT GST. All RL / DM / PL prices are Ex-GST.
      </div>

      <div style={{display:"flex",gap:8,marginBottom:12,flexWrap:"wrap"}}>
        <select style={Object.assign({},SEL,{flex:1,padding:"9px 11px"})} value={fCat} onChange={e => setFCat(e.target.value)}>
          <option value="All">All Categories</option>
          {CATS.map(c => <option key={c} value={c}>{c}</option>)}
        </select>
        <div style={{display:"flex",border:"1px solid "+C.border,borderRadius:9,overflow:"hidden"}}>
          {["All","Active","Inactive"].map(s => (
            <button key={s} onClick={() => setFSt(s)} style={{padding:"9px 11px",border:"none",cursor:"pointer",fontWeight:600,fontSize:12,fontFamily:"inherit",background:fSt===s?C.blue:"#fff",color:fSt===s?"#fff":C.sec}}>{s}</button>
          ))}
        </div>
      </div>

      <div style={{display:"grid",gridTemplateColumns:"2fr 1fr",gap:8,marginBottom:12}}>
        <BtnP onClick={openAdd}>+ Add Item</BtnP>
        <BtnO color={C.blue} onClick={() => setShowExp(true)}>⬇ Export</BtnO>
      </div>

      <div style={{fontSize:12,color:C.mute,marginBottom:10,fontWeight:600}}>{list.length+" item"+(list.length!==1?"s":"")}</div>
      {list.length===0 && <div style={{textAlign:"center",padding:"36px",color:C.mute,fontSize:14}}>No items. Add using the button above.</div>}

      {list.map(it => (
        <div key={it.id}>
          <ItemCard it={it} isAdmin={true} showDate={true} />
          <div style={{display:"flex",gap:8,marginTop:-4,marginBottom:8}}>
            <button onClick={() => openEdit(it)} style={{flex:1,padding:"9px",borderRadius:8,border:"1.5px solid "+C.border,background:"#F9FAFB",cursor:"pointer",fontSize:13,fontWeight:600,color:C.sec,fontFamily:"inherit"}}>Edit</button>
            <button onClick={() => toggle(it.id)} style={{flex:1,padding:"9px",borderRadius:8,border:"1.5px solid "+C.border,background:"#F9FAFB",cursor:"pointer",fontSize:12,fontWeight:600,color:it.active?"#991B1B":"#065F46",fontFamily:"inherit"}}>{it.active?"Set Inactive":"Set Active"}</button>
            <button onClick={() => setConf(it.id)} style={{width:40,borderRadius:8,border:"1.5px solid #FCA5A5",background:C.redBg,cursor:"pointer",fontSize:15,color:C.red,fontFamily:"inherit"}}>✕</button>
          </div>
        </div>
      ))}

      <Drawer open={open} onClose={close} title={eid?"Edit Item":"Add Item"}
        footer={
          <div style={{display:"flex",gap:9}}>
            <BtnO onClick={close}>Cancel</BtnO>
            <BtnP color={eid?"#F59E0B":C.blue} onClick={save}>{eid?"Update Item":"Save Item"}</BtnP>
          </div>
        }>
        <Fld label="Category *">
          <select style={SEL} value={form.cat} onChange={e => setForm(p => ({...p,cat:e.target.value,bId:""}))}>
            {CATS.map(c => <option key={c} value={c}>{c}</option>)}
          </select>
        </Fld>
        <Fld label="Brand">
          <select style={SEL} value={form.bId} onChange={e => setForm(p => ({...p,bId:e.target.value}))}>
            <option value="">— Select Brand (optional) —</option>
            {brands.map(b => <option key={b.id} value={b.id}>{b.code+" — "+b.name}</option>)}
          </select>
          {brand && (
            <div style={{marginTop:4,display:"flex",gap:6}}>
              <Chip col={C.rl} bg={C.rlBg} br={C.rlBr}>{"RL "+fpct(brand.rlMarkup)}</Chip>
              {brand.dmMarkup!=null
                ? <Chip col={C.dm} bg={C.dmBg} br={C.dmBr}>{"DM "+fpct(brand.dmMarkup)}</Chip>
                : <Chip col={C.mute} bg="#F9FAFB" br={C.border}>DM not set</Chip>}
            </div>
          )}
        </Fld>
        <Fld label="Item Name *" err={errs.name}>
          <input style={ei("name")} value={form.name} placeholder="e.g. King Cotton Bedsheet 200TC" onChange={e => setForm(p => ({...p,name:e.target.value}))} />
        </Fld>
        <Fld label="Purchase Price WITHOUT GST (₹) *" err={errs.px}>
          <input style={Object.assign({},ei("px"),{fontWeight:700,fontSize:16,color:C.blue})} type="number" step="0.01" placeholder="e.g. 380" value={form.purchaseEx} onChange={e => setForm(p => ({...p,purchaseEx:e.target.value}))} />
        </Fld>
        <Fld label="GST %">
          <select style={SEL} value={form.gst} onChange={e => setForm(p => ({...p,gst:parseFloat(e.target.value)}))}>
            {GST_OPTS.map(o => <option key={o.v} value={o.v}>{o.label}</option>)}
          </select>
          {prevEx>0 && <div style={{marginTop:4,fontSize:12,color:C.amb,fontWeight:600}}>{"Purchase Inc-GST = "+fp(calcIncGST(prevEx,form.gst))}</div>}
        </Fld>
        <MarginOverride label="Custom RL Margin % (Override)" col={C.rl} bg={C.rlBg} br={C.rlBr} value={form.customRL} onChange={v => setForm(p => ({...p,customRL:v}))} srcLabel={rlSrc} err={errs.crl} />
        <MarginOverride label="Custom DM Margin % (Override)" col={C.dm} bg={C.dmBg} br={C.dmBr} value={form.customDM} onChange={v => setForm(p => ({...p,customDM:v}))} srcLabel={dmSrc} err={errs.cdm} />
        <Fld label="PL Add-on (extra ₹ on top of RL for retail)" err={errs.ca}>
          <select style={SEL} value={form.plAddon||""} onChange={e => { const v=e.target.value; setForm(p => ({...p,plAddon:v===""?"":v==="custom"?"custom":parseInt(v)})); }}>
            {ADDON_OPTS.map(o => <option key={String(o.v)} value={o.v}>{o.label}</option>)}
          </select>
          {form.plAddon==="custom" && <input style={Object.assign({},INP,{marginTop:7})} type="number" placeholder="Custom ₹ amount" value={form.customAddon} onChange={e => setForm(p => ({...p,customAddon:e.target.value}))} />}
        </Fld>

        {prev && (
          <div style={{background:"#F8FAFF",border:"1px solid #BFDBFE",borderRadius:11,padding:"13px",marginBottom:6}}>
            <div style={{fontSize:11,fontWeight:700,color:C.blue,textTransform:"uppercase",letterSpacing:"0.5px",marginBottom:9}}>📊 Live Price Preview</div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:7,marginBottom:9}}>
              <div style={{background:C.ambBg,border:"1px solid #FCD34D",borderRadius:8,padding:"7px 9px",textAlign:"center"}}>
                <div style={{fontSize:9,color:C.amb,fontWeight:700,textTransform:"uppercase",marginBottom:2}}>Inc-GST</div>
                <div style={{fontSize:13,fontWeight:800,color:C.amb}}>{fp(prev.incGST)}</div>
              </div>
              <div style={{background:"#F3F4F6",border:"1px solid "+C.border,borderRadius:8,padding:"7px 9px",textAlign:"center"}}>
                <div style={{fontSize:9,color:C.sec,fontWeight:700,textTransform:"uppercase",marginBottom:2}}>RL Markup</div>
                <div style={{fontSize:13,fontWeight:800}}>{fpct(prev.rlM)}</div>
              </div>
            </div>
            <ProfitRow label="RL" col={C.rl} price={prev.rl} profit={prev.rlProfit} />
            <ProfitRow label="DM" col={C.dm} price={prev.dm} profit={prev.dmProfit} />
            <ProfitRow label="PL" col={C.pl} price={prev.pl} profit={prev.plProfit} />
          </div>
        )}

        <Fld label="Notes (optional)">
          <input style={INP} value={form.notes} placeholder="e.g. seasonal, imported..." onChange={e => setForm(p => ({...p,notes:e.target.value}))} />
        </Fld>
        <Fld label="Status">
          <div style={{display:"flex",gap:8}}>
            {["Active","Inactive"].map(s => {
              const on = (s==="Active"&&form.active)||(s==="Inactive"&&!form.active);
              return (
                <button key={s} onClick={() => setForm(p => ({...p,active:s==="Active"}))}
                  style={{flex:1,padding:"10px",borderRadius:8,border:"1.5px solid "+(on?C.blue:C.border),background:on?"#EFF6FF":"#F9FAFB",color:on?C.blue:C.sec,fontWeight:700,fontSize:13,cursor:"pointer",fontFamily:"inherit"}}>
                  {s}
                </button>
              );
            })}
          </div>
        </Fld>
      </Drawer>
    </div>
  );
}

// ── PRICE LIST VIEW ───────────────────────────────────────────────
function PriceListView({ brands, items, settings, isAdmin }) {
  const [viewMode, setViewMode] = useState("general");
  const [fCat,     setFCat]     = useState("All");
  const [fBrand,   setFBrand]   = useState("All");
  const [search,   setSearch]   = useState("");

  const enriched = useMemo(() => {
    return items.filter(i => i.active).map(i => {
      const b = brands.find(x => x.id===i.bId)||null;
      return Object.assign({}, i, computeItem(i,b,settings), {_b:b});
    });
  }, [items, brands, settings]);

  const filtered = enriched.filter(i => {
    if (fCat!=="All" && i.cat!==fCat) return false;
    if (fBrand!=="All" && i.bId!==fBrand) return false;
    if (search && !i.name.toLowerCase().includes(search.toLowerCase()) && !(i._b&&i._b.name.toLowerCase().includes(search.toLowerCase()))) return false;
    return true;
  });

  const grouped = useMemo(() => {
    const g = {};
    filtered.forEach(i => { if (!g[i.cat]) g[i.cat]=[]; g[i.cat].push(i); });
    return g;
  }, [filtered]);

  const byDate = useMemo(() => {
    const sorted = filtered.slice().sort((a,b) => new Date(b.updatedAt||0) - new Date(a.updatedAt||0));
    const g = {};
    sorted.forEach(i => { const k=dateGroupKey(i.updatedAt); if(!g[k])g[k]=[]; g[k].push(i); });
    return g;
  }, [filtered]);

  const today = new Date().toLocaleDateString("en-IN",{day:"2-digit",month:"short",year:"numeric"});
  const views = [{id:"general",icon:"📂",label:"By Category"},{id:"date",icon:"📅",label:"By Date"}];

  return (
    <div style={{padding:"16px 16px 80px"}}>
      <div style={{background:C.navy,borderRadius:13,padding:"15px",marginBottom:12,color:"#fff"}}>
        <div style={{fontWeight:800,fontSize:17}}>{settings.co}</div>
        <div style={{fontSize:11,color:"rgba(255,255,255,0.5)",marginTop:2}}>{settings.tag}</div>
        <div style={{marginTop:8,fontSize:11,color:"rgba(255,255,255,0.35)"}}>{"Price List · "+today+" · All prices EXCLUDING GST"}</div>
      </div>
      <div style={{background:C.ambBg,border:"1px solid #FCD34D",borderRadius:9,padding:"9px 13px",marginBottom:12,fontSize:12,color:C.amb,fontWeight:600}}>
        ⚠ All prices EXCLUDING GST — GST charged extra.
      </div>
      <div style={{display:"flex",border:"1px solid "+C.border,borderRadius:10,overflow:"hidden",marginBottom:14}}>
        {views.map(v => (
          <button key={v.id} onClick={() => setViewMode(v.id)}
            style={{flex:1,padding:"11px",border:"none",cursor:"pointer",fontWeight:700,fontSize:13,fontFamily:"inherit",display:"flex",alignItems:"center",justifyContent:"center",gap:6,background:viewMode===v.id?C.blue:"#fff",color:viewMode===v.id?"#fff":C.sec}}>
            <span>{v.icon}</span><span>{v.label}</span>
          </button>
        ))}
      </div>
      {viewMode==="date" && (
        <div style={{display:"flex",gap:10,marginBottom:12,flexWrap:"wrap"}}>
          <Chip col={C.new_} bg={C.newBg} br={C.newBr}>🆕 New today</Chip>
          <Chip col={C.upd}  bg={C.updBg} br={C.updBr}>✏️ Price updated</Chip>
        </div>
      )}
      <input style={Object.assign({},INP,{marginBottom:10})} placeholder="🔍 Search by item or brand..." value={search} onChange={e => setSearch(e.target.value)} />
      <div style={{display:"flex",gap:8,marginBottom:12}}>
        <select style={Object.assign({},SEL,{flex:1,padding:"9px 11px"})} value={fCat} onChange={e => setFCat(e.target.value)}>
          <option value="All">All Categories</option>
          {CATS.map(c => <option key={c} value={c}>{c}</option>)}
        </select>
        <select style={Object.assign({},SEL,{flex:1,padding:"9px 11px"})} value={fBrand} onChange={e => setFBrand(e.target.value)}>
          <option value="All">All Brands</option>
          {brands.map(b => <option key={b.id} value={b.id}>{b.code+" — "+b.name}</option>)}
        </select>
      </div>
      <div style={{display:"flex",gap:12,marginBottom:10}}>
        <div style={{fontSize:11,fontWeight:700,color:C.rl}}>🟢 RL = Wholesale</div>
        <div style={{fontSize:11,fontWeight:700,color:C.dm}}>🔵 DM = Sp. Wholesale</div>
        <div style={{fontSize:11,fontWeight:700,color:C.pl}}>🟣 PL = Retail</div>
      </div>
      <div style={{fontSize:12,color:C.mute,marginBottom:10,fontWeight:600}}>{filtered.length+" items"}</div>

      {viewMode==="general" && (
        <div>
          {Object.keys(grouped).length===0 && <div style={{textAlign:"center",padding:"36px",color:C.mute,fontSize:14}}>No items found.</div>}
          {Object.keys(grouped).map(cat => (
            <div key={cat}>
              <div style={{fontSize:12,fontWeight:800,color:C.navy,marginBottom:7,marginTop:4,textTransform:"uppercase",letterSpacing:"0.4px",borderLeft:"3px solid "+C.blue,paddingLeft:9}}>{cat}</div>
              {grouped[cat].map(it => <ItemCard key={it.id} it={it} isAdmin={isAdmin} showDate={false} />)}
            </div>
          ))}
        </div>
      )}

      {viewMode==="date" && (
        <div>
          {Object.keys(byDate).length===0 && <div style={{textAlign:"center",padding:"36px",color:C.mute,fontSize:14}}>No items found.</div>}
          {Object.keys(byDate).map(day => (
            <div key={day}>
              <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:10,marginTop:8}}>
                <div style={{flex:1,height:1,background:C.border}} />
                <div style={{background:C.navy,color:"#fff",borderRadius:20,padding:"4px 14px",fontSize:12,fontWeight:700,whiteI notice you pasted the corrected code back to me — but it looks like it got cut off at the end (the `PriceListView` component and everything after it is incomplete). Let me verify the fixes are correct and point out anything still missing.

## ✅ Confirmed Fixes Applied

| Fix | Status |
|-----|--------|
| Removed duplicate `initializeApp` | ✅ |
| Removed duplicate imports | ✅ |
| Removed unused `useFirestoreSync` | ✅ |
| Single `app` + `db` initialization | ✅ |

## ⚠️ The Code You Pasted Back is INCOMPLETE

It cuts off mid-way through `PriceListView`. The missing parts are:

- Rest of `PriceListView` (the `byDate` rendering)
- `SettingsView` component
- `AdminApp` component  
- `SalesmanApp` component
- `App` root component (the most important part with Firebase sync logic)

## 🔧 Here's the Complete Fixed Code

Since the full file is very long, here are the **remaining sections** you need to append after where your paste cut off. I've verified these have the same fixes applied:

```jsx
              </div>
              {byDate[day].map(it => <ItemCard key={it.id} it={it} isAdmin={isAdmin} showDate={true} />)}
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

// ── SETTINGS VIEW ─────────────────────────────────────────────────
function SettingsView({ settings, onSettingsChange, syncStatus }) {
  const [form, setForm] = useState({
    co: settings.co, tag: settings.tag,
    defaultRL: +(settings.defaultRL*100).toFixed(1),
    defaultDM: settings.defaultDM!=null ? +(parseFloat(settings.defaultDM)*100).toFixed(1) : "",
  });
  useEffect(() => {
    setForm({ co:settings.co, tag:settings.tag, defaultRL:+(settings.defaultRL*100).toFixed(1), defaultDM:settings.defaultDM!=null?+(parseFloat(settings.defaultDM)*100).toFixed(1):"" });
  }, [settings]);

  const [nAP,setNAP]=useState(""); const [cAP,setCAP]=useState("");
  const [nSP,setNSP]=useState(""); const [cSP,setCSP]=useState("");

  function saveGen() {
    if (!form.co.trim()) return toast("Company name required","err");
    if (isNaN(parseFloat(form.defaultRL))) return toast("RL markup must be a number","err");
    onSettingsChange({ ...settings, co:form.co.trim(), tag:form.tag.trim(), defaultRL:parseFloat(form.defaultRL)/100, defaultDM:form.defaultDM!==""&&!isNaN(parseFloat(form.defaultDM))?parseFloat(form.defaultDM)/100:null });
    toast("Settings saved ✓");
  }
  function changePIN(type) {
    const np = type==="admin"?nAP:nSP;
    const cp = type==="admin"?cAP:cSP;
    if (np.length!==4||isNaN(parseInt(np))) return toast("PIN must be exactly 4 digits","err");
    if (np!==cp) return toast("PINs do not match","err");
    onSettingsChange({ ...settings, [type==="admin"?"adminPIN":"salesPIN"]:np });
    if (type==="admin") { setNAP(""); setCAP(""); } else { setNSP(""); setCSP(""); }
    toast((type==="admin"?"Admin":"Salesman")+" PIN updated ✓");
  }

  const SS = {background:C.card,border:"1px solid "+C.border,borderRadius:13,padding:"17px",marginBottom:13,boxShadow:"0 1px 6px rgba(0,0,0,0.04)"};

  return (
    <div style={{padding:"16px 16px 80px"}}>
      {/* Sync status card */}
      <div style={Object.assign({},SS,{display:"flex",alignItems:"center",justifyContent:"space-between"})}>
        <div>
          <div style={{fontSize:13,fontWeight:800,marginBottom:4}}>☁️ Firebase Sync</div>
          <div style={{fontSize:11,color:C.mute}}>All devices share one database</div>
        </div>
        <SyncBadge status={syncStatus} />
      </div>

      <div style={SS}>
        <div style={{fontSize:14,fontWeight:800,marginBottom:14}}>General Settings</div>
        <Fld label="Company Name"><input style={INP} value={form.co} onChange={e => setForm(p => ({...p,co:e.target.value}))} /></Fld>
        <Fld label="Tagline"><input style={INP} value={form.tag} onChange={e => setForm(p => ({...p,tag:e.target.value}))} /></Fld>
        <Fld label="Default RL Markup %" hint="Enter number: 18 = 18%">
          <input style={INP} type="number" step="0.5" value={form.defaultRL} onChange={e => setForm(p => ({...p,defaultRL:e.target.value}))} />
        </Fld>
        <Fld label="Default DM Markup %" hint="Enter number: 10 = 10%. Leave blank if not decided.">
          <input style={INP} type="number" step="0.5" placeholder="Leave blank" value={form.defaultDM} onChange={e => setForm(p => ({...p,defaultDM:e.target.value}))} />
          {form.defaultDM!==""&&!isNaN(parseFloat(form.defaultDM)) && (
            <div style={{marginTop:4,fontSize:11,color:C.dm,fontWeight:600}}>{"Will apply: "+parseFloat(form.defaultDM).toFixed(1)+"% DM on all items without override"}</div>
          )}
        </Fld>
        <BtnP onClick={saveGen}>Save Settings</BtnP>
      </div>

      {[{type:"admin",label:"Admin"},{type:"salesman",label:"Salesman"}].map(pt => (
        <div key={pt.type} style={SS}>
          <div style={{fontSize:14,fontWeight:800,marginBottom:14}}>{"Change "+pt.label+" PIN"}</div>
          <Fld label="New PIN (4 digits)">
            <input style={INP} type="password" maxLength={4} placeholder="4-digit PIN"
              value={pt.type==="admin"?nAP:nSP}
              onChange={e => { const v=e.target.value.replace(/\D/g,"").slice(0,4); pt.type==="admin"?setNAP(v):setNSP(v); }} />
          </Fld>
          <Fld label="Confirm PIN">
            <input style={INP} type="password" maxLength={4} placeholder="Re-enter to confirm"
              value={pt.type==="admin"?cAP:cSP}
              onChange={e => { const v=e.target.value.replace(/\D/g,"").slice(0,4); pt.type==="admin"?setCAP(v):setCSP(v); }} />
          </Fld>
          <BtnP color="#6D28D9" onClick={() => changePIN(pt.type)}>{"Update "+pt.label+" PIN"}</BtnP>
        </div>
      ))}

      <div style={Object.assign({},SS,{textAlign:"center"})}>
        <div style={{fontSize:13,color:C.sec,lineHeight:1.7}}>{"v"+VER+" · Firebase Firestore"}</div>
        <div style={{fontSize:12,color:C.mute,marginTop:4}}>Data syncs automatically across all devices</div>
      </div>
    </div>
  );
}

// ── ADMIN APP ─────────────────────────────────────────────────────
function AdminApp({ settings, onSettingsChange, brands, onBrandsChange, items, onItemsChange, onLogout, syncStatus }) {
  const [tab, setTab] = useState("master");
  const NAV = [
    {id:"master",    icon:"📋", label:"Master"},
    {id:"pricelist", icon:"🧾", label:"Prices"},
    {id:"brands",    icon:"🏷",  label:"Brands"},
    {id:"settings",  icon:"⚙️", label:"Settings"},
  ];
  return (
    <div style={{minHeight:"100vh",background:C.bg}}>
      <div style={{background:C.navy,padding:"13px 16px",display:"flex",alignItems:"center",justifyContent:"space-between",position:"sticky",top:0,zIndex:100,boxShadow:"0 2px 10px rgba(0,0,0,0.2)"}}>
        <div>
          <div style={{color:"#fff",fontWeight:800,fontSize:15}}>{settings.co}</div>
          <div style={{color:"rgba(255,255,255,0.4)",fontSize:11,display:"flex",alignItems:"center",gap:8}}>
            Admin 🔐
            <SyncBadge status={syncStatus} />
          </div>
        </div>
        <button onClick={onLogout} style={{padding:"7px 13px",borderRadius:8,border:"1.5px solid rgba(255,255,255,0.22)",background:"rgba(255,255,255,0.08)",color:"#fff",fontSize:12,fontWeight:600,cursor:"pointer",fontFamily:"inherit"}}>Logout</button>
      </div>
      {tab==="master"    && <MasterView    brands={brands} items={items} onItemsChange={onItemsChange} settings={settings} />}
      {tab==="pricelist" && <PriceListView brands={brands} items={items} settings={settings} isAdmin={true} />}
      {tab==="brands"    && <BrandsView    brands={brands} onBrandsChange={onBrandsChange} />}
      {tab==="settings"  && <SettingsView  settings={settings} onSettingsChange={onSettingsChange} syncStatus={syncStatus} />}
      <div style={{position:"fixed",bottom:0,left:0,right:0,zIndex:200,background:"#fff",borderTop:"1px solid "+C.border,display:"flex",height:62,boxShadow:"0 -3px 16px rgba(0,0,0,0.07)"}}>
        {NAV.map(t => (
          <button key={t.id} onClick={() => setTab(t.id)}
            style={{flex:1,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",gap:2,border:"none",background:"transparent",cursor:"pointer",fontFamily:"inherit",borderTop:tab===t.id?"2.5px solid "+C.blue:"2.5px solid transparent"}}>
            <span style={{fontSize:19}}>{t.icon}</span>
            <span style={{fontSize:10,fontWeight:700,color:tab===t.id?C.blue:C.mute}}>{t.label}</span>
          </button>
        ))}
      </div>
    </div>
  );
}

// ── SALESMAN APP ──────────────────────────────────────────────────
function SalesmanApp({ settings, brands, items, onLogout, syncStatus }) {
  return (
    <div style={{minHeight:"100vh",background:C.bg}}>
      <div style={{background:C.navy,padding:"13px 16px",display:"flex",alignItems:"center",justifyContent:"space-between",position:"sticky",top:0,zIndex:100,boxShadow:"0 2px 10px rgba(0,0,0,0.2)"}}>
        <div>
          <div style={{color:"#fff",fontWeight:800,fontSize:15}}>{settings.co}</div>
          <div style={{color:"rgba(255,255,255,0.4)",fontSize:11,display:"flex",alignItems:"center",gap:8}}>
            Price List 📋
            <SyncBadge status={syncStatus} />
          </div>
        </div>
        <button onClick={onLogout} style={{padding:"7px 13px",borderRadius:8,border:"1.5px solid rgba(255,255,255,0.22)",background:"rgba(255,255,255,0.08)",color:"#fff",fontSize:12,fontWeight:600,cursor:"pointer",fontFamily:"inherit"}}>Logout</button>
      </div>
      <PriceListView brands={brands} items={items} settings={settings} isAdmin={false} />
    </div>
  );
}

// ── ROOT ──────────────────────────────────────────────────────────
export default function App() {
  const [settings,    setSettings]    = useState(DEF_S);
  const [brands,      setBrands]      = useState(DEF_B);
  const [items,       setItems]       = useState(DEF_I);
  const [role,        setRole]        = useState(null);
  const [ready,       setReady]       = useState(false);
  const [syncStatus,  setSyncStatus]  = useState("synced");

  // ── Initial load from Firestore ──────────────────────────────────
  useEffect(() => {
    fsLoad().then(d => {
      if (d.settings) setSettings(d.settings);
      if (d.brands)   setBrands(d.brands);
      if (d.items)    setItems(d.items);
      setReady(true);
    }).catch(() => {
      // Firestore unavailable — use defaults and show offline
      setSyncStatus("offline");
      setReady(true);
    });
  }, []);

  // ── Real-time listeners (changes from other devices appear instantly) ──
  // We only start listeners after initial load to avoid double-setting state
  const listenersStarted = useRef(false);
  useEffect(() => {
    if (!ready || listenersStarted.current) return;
    listenersStarted.current = true;

    const unsubS = onSnapshot(FS.settings(), snap => { if (snap.exists()) setSettings(snap.data().v); }, () => setSyncStatus("offline"));
    const unsubB = onSnapshot(FS.brands(),   snap => { if (snap.exists()) setBrands(snap.data().v);   }, () => setSyncStatus("offline"));
    const unsubI = onSnapshot(FS.items(),    snap => { if (snap.exists()) setItems(snap.data().v);    }, () => setSyncStatus("offline"));

    return () => { unsubS(); unsubB(); unsubI(); };
  }, [ready]);

  // ── Save helpers — write to Firestore, reflect status badge ─────
  async function saveSettings(next) {
    setSettings(next);
    setSyncStatus("saving");
    try { await fsSave("settings", next); setSyncStatus("synced"); }
    catch { setSyncStatus("error"); toast("Sync failed — check connection","err"); }
  }
  async function saveBrands(next) {
    setBrands(next);
    setSyncStatus("saving");
    try { await fsSave("brands", next); setSyncStatus("synced"); }
    catch { setSyncStatus("error"); toast("Sync failed — check connection","err"); }
  }
  async function saveItems(next) {
    setItems(next);
    setSyncStatus("saving");
    try { await fsSave("items", next); setSyncStatus("synced"); }
    catch { setSyncStatus("error"); toast("Sync failed — check connection","err"); }
  }

  if (!ready) {
    return (
      <div style={{minHeight:"100vh",background:C.bg,display:"flex",alignItems:"center",justifyContent:"center"}}>
        <div style={{textAlign:"center"}}>
          <div style={{width:52,height:52,borderRadius:14,background:C.navy,display:"flex",alignItems:"center",justifyContent:"center",color:"#fff",fontWeight:900,fontSize:20,margin:"0 auto 10px"}}>VK</div>
          <div style={{color:C.sec,fontSize:13}}>Connecting to Firebase…</div>
        </div>
      </div>
    );
  }

  return (
    <div>
      <style>{`*{box-sizing:border-box;-webkit-tap-highlight-color:transparent;}input:focus,select:focus{outline:none!important;border-color:#1648D6!important;box-shadow:0 0 0 3px rgba(22,72,214,0.1)!important;}input[type=number]::-webkit-inner-spin-button{-webkit-appearance:none;}input::placeholder{color:#CBD5E1;}button:active{opacity:0.75;transform:scale(0.97);}::-webkit-scrollbar{width:4px;}::-webkit-scrollbar-thumb{background:#D1D5DB;border-radius:4px;}@keyframes shk{0%,100%{transform:translateX(0)}25%{transform:translateX(-5px)}75%{transform:translateX(5px)}}`}</style>
      <ToastHost />
      <div style={{fontFamily:"system-ui,-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif"}}>
        {!role             && <Login       settings={settings} onLogin={setRole} />}
        {role==="admin"    && <AdminApp    settings={settings} onSettingsChange={saveSettings} brands={brands} onBrandsChange={saveBrands} items={items} onItemsChange={saveItems} onLogout={() => setRole(null)} syncStatus={syncStatus} />}
        {role==="salesman" && <SalesmanApp settings={settings} brands={brands} items={items} onLogout={() => setRole(null)} syncStatus={syncStatus} />}
      </div>
    </div>
  );
}
