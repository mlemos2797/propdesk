import { useState, useEffect, useCallback, useRef } from "react";
import {
  LineChart, Line, AreaChart, Area, RadarChart, Radar, PolarGrid, PolarAngleAxis,
  XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer,
  PieChart, Pie, Cell, Legend
} from "recharts";
import * as XLSX from "xlsx";
import { createClient } from "@supabase/supabase-js";

// ─── SUPABASE (as tuas chaves) ────────────────────────────────────────────────
const supabaseUrl = "https://hcoigktkhcbcndupnuro.supabase.co";
const supabaseAnonKey = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Imhjb2lna3RraGNiY25kdXBudXJvIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODA5MTA0NTMsImV4cCI6MjA5NjQ4NjQ1M30.oqUwazkGCXQ0Gyz-qfQ6eMDKQBn9pT16nEFoSiT-wno";
const supabase = createClient(supabaseUrl, supabaseAnonKey);

// ─── CONSTANTS ────────────────────────────────────────────────────────────────
const DAYS = ["MON","TUE","WED","THU","FRI"];
const DAY_LABELS = ["Monday","Tuesday","Wednesday","Thursday","Friday"];
const EMOTIONS = ["Disciplined","Calm","Confident","Patient","Focused","Hesitant","Anxious","FOMO","Greedy","Revenge","Impulsive","Overconfident"];
const SETUPS = ["ICT FVG","Order Block","Breaker Block","MSS/CHoCH","Liquidity Sweep","OR Mid Pullback","NWOG/NDOG","PD Array","Silver Bullet","Killzone Fade","Custom"];
const SESSIONS = ["Asian","London","NY AM (08-11)","NY Lunch","NY PM (13-16)","Overnight"];
const MARKETS = ["ES","MES","NQ","MNQ","YM","MYM","RTY","M2K","EUR/USD","GBP/USD","XAU/USD"];
const BIASES = ["Strong Bullish","Bullish","Neutral","Bearish","Strong Bearish"];
const H1_STRUCT = ["Confirms ✓","Partial ~","Against ✗"];
const H1_POS = ["Discount","Equilibrium","Premium"];
const TF_CONFIRM = ["YES ✓","PARTIAL ~","NO ✗"];
const PHASES = ["Challenge","Verification","Funded","Paused","Disabled"];

const mkId = () => `p${Date.now()}`;
const DEFAULT_PROPS = [
  { id:"p1", name:"APEX",     color:"#F59E0B", accountSize:50000,  maxDD:2500, weeklyTarget:2, riskPct:0.5, phase:"Funded", note:"" },
  { id:"p2", name:"FTMO",     color:"#10B981", accountSize:100000, maxDD:5000, weeklyTarget:2, riskPct:0.5, phase:"Funded", note:"" },
  { id:"p3", name:"MFF",      color:"#818CF8", accountSize:50000,  maxDD:2000, weeklyTarget:2, riskPct:0.5, phase:"Funded", note:"" },
  { id:"p4", name:"THE5%ERS", color:"#F43F5E", accountSize:40000,  maxDD:2000, weeklyTarget:2, riskPct:0.5, phase:"Funded", note:"" },
];
const DEFAULT_SCHEDULE = { MON:["p1"], TUE:["p2"], WED:["p3"], THU:["p4"], FRI:["p1"] };

// ─── HOOKS ────────────────────────────────────────────────────────────────────
const useWindowSize = () => {
  const [w, setW] = useState(typeof window !== "undefined" ? window.innerWidth : 1200);
  useEffect(() => { const h=()=>setW(window.innerWidth); window.addEventListener("resize",h); return ()=>window.removeEventListener("resize",h); },[]);
  return w;
};

const useToast = () => {
  const [t,setT]=useState(null);
  const show=useCallback((msg,type="success")=>setT({msg,type,id:Date.now()}),[]);
  const hide=useCallback(()=>setT(null),[]);
  return [t,show,hide];
};

// ─── HELPERS ──────────────────────────────────────────────────────────────────
const getTodayDayKey = () => ["SUN","MON","TUE","WED","THU","FRI","SAT"][new Date().getDay()];
const getWeekDates = () => {
  const today=new Date(), day=today.getDay(), mon=new Date(today);
  mon.setDate(today.getDate()-day+(day===0?-6:1));
  return DAYS.map((_,i)=>{ const d=new Date(mon); d.setDate(mon.getDate()+i); return d.toISOString().split("T")[0]; });
};
const normDay = (schedule, d) => { const v=schedule[d]; return Array.isArray(v)?v:v?[v]:[]; };
const fmtPct = (n,dec=2) => { const v=parseFloat(n); return isNaN(v)?"—":(v>=0?"+":"")+v.toFixed(dec)+"%"; };
const fmtR = (n) => { const v=parseFloat(n); return isNaN(v)?"—":(v>=0?"+":"")+v.toFixed(1)+"R"; };
const rColor = (r) => !r?"#6B7280":r>=4.5?"#10B981":r>=3.5?"#F59E0B":r>=2.5?"#F97316":"#EF4444";

// ─── TOAST ────────────────────────────────────────────────────────────────────
const Toast = ({ message, type, onDone }) => {
  useEffect(()=>{ const t=setTimeout(onDone,2600); return()=>clearTimeout(t); },[]);
  const colors = { success:["#064E3B","#10B981","#6EE7B7"], error:["#7F1D1D","#EF4444","#FCA5A5"], info:["#1E1B4B","#818CF8","#A5B4FC"] };
  const [bg,bd,col] = colors[type]||colors.info;
  return (
    <div style={{ position:"fixed",bottom:24,right:24,zIndex:9998,background:bg,border:`1px solid ${bd}`,
      borderRadius:10,padding:"12px 20px",color:col,fontSize:13,fontFamily:"'IBM Plex Mono',monospace",
      fontWeight:700,boxShadow:"0 8px 32px rgba(0,0,0,0.6)",animation:"slideIn 0.2s ease" }}>
      {message}
    </div>
  );
};

// ─── CONFIRM MODAL ────────────────────────────────────────────────────────────
const ConfirmModal = ({ message, onConfirm, onCancel }) => (
  <div onClick={onCancel} style={{ position:"fixed",inset:0,zIndex:9999,background:"rgba(0,0,0,0.8)",
    backdropFilter:"blur(4px)",display:"flex",alignItems:"center",justifyContent:"center" }}>
    <div onClick={e=>e.stopPropagation()} style={{ background:"#0D1117",border:"1px solid #374151",
      borderRadius:14,padding:"28px 32px",maxWidth:380,width:"90%",boxShadow:"0 20px 60px rgba(0,0,0,0.8)" }}>
      <div style={{ fontSize:13,color:"#D1D5DB",fontFamily:"'IBM Plex Mono',monospace",lineHeight:1.7,marginBottom:24 }}>{message}</div>
      <div style={{ display:"flex",gap:10,justifyContent:"flex-end" }}>
        <button onClick={onCancel} style={{ padding:"9px 22px",borderRadius:8,cursor:"pointer",fontSize:12,
          fontFamily:"'IBM Plex Mono',monospace",border:"1px solid #374151",background:"transparent",color:"#9CA3AF" }}>CANCEL</button>
        <button onClick={onConfirm} style={{ padding:"9px 22px",borderRadius:8,cursor:"pointer",fontSize:12,
          fontFamily:"'IBM Plex Mono',monospace",border:"1px solid #EF4444",background:"#7F1D1D",color:"#FCA5A5",fontWeight:700 }}>CONFIRM</button>
      </div>
    </div>
  </div>
);

// ─── AUTH SCREEN ──────────────────────────────────────────────────────────────
const AuthScreen = ({ onLogin, showToast }) => {
  const [mode, setMode] = useState("login"); // login | signup
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [loading, setLoading] = useState(false);

  const handle = async (e) => {
    e.preventDefault();
    if (!email || !password) { showToast("Preenche email e password", "error"); return; }
    setLoading(true);
    try {
      if (mode === "login") {
        const { data, error } = await supabase.auth.signInWithPassword({ email, password });
        if (error) throw error;
        showToast("Login feito ✓", "success");
        onLogin(data.session);
      } else {
        const { data, error } = await supabase.auth.signUp({ email, password });
        if (error) throw error;
        showToast("Conta criada! Já podes entrar.", "success");
        setMode("login");
      }
    } catch (err) {
      showToast(err.message || "Erro", "error");
    } finally {
      setLoading(false);
    }
  };

  return (
    <div style={{ minHeight:"100vh", background:"#060A0F", display:"flex", alignItems:"center", justifyContent:"center", padding:20 }}>
      <div style={{ background:"#0D1117", border:"1px solid #1F2937", borderRadius:16, padding:"36px 32px", width:"100%", maxWidth:380 }}>
        <div style={{ fontFamily:"'Bebas Neue',sans-serif", fontSize:32, letterSpacing:"0.15em",
          background:"linear-gradient(135deg,#F59E0B,#FDE68A)", WebkitBackgroundClip:"text", WebkitTextFillColor:"transparent",
          textAlign:"center", marginBottom:8 }}>PROPDESK</div>
        <div style={{ fontSize:12, color:"#6B7280", textAlign:"center", marginBottom:28, fontFamily:"'IBM Plex Mono',monospace" }}>
          {mode === "login" ? "Entra na tua conta" : "Cria a tua conta"}
        </div>

        <form onSubmit={handle}>
          <div style={{ marginBottom:14 }}>
            <label style={{ fontSize:10, color:"#6B7280", fontFamily:"'IBM Plex Mono',monospace", letterSpacing:"0.12em", display:"block", marginBottom:6 }}>EMAIL</label>
            <input type="email" value={email} onChange={e=>setEmail(e.target.value)} required
              style={{ width:"100%", background:"#0A0D14", border:"1px solid #1F2937", borderRadius:8, color:"#F9FAFB",
                padding:"12px 14px", fontSize:14, fontFamily:"'IBM Plex Mono',monospace", outline:"none" }}
              placeholder="teu@email.com" />
          </div>
          <div style={{ marginBottom:22 }}>
            <label style={{ fontSize:10, color:"#6B7280", fontFamily:"'IBM Plex Mono',monospace", letterSpacing:"0.12em", display:"block", marginBottom:6 }}>PASSWORD</label>
            <input type="password" value={password} onChange={e=>setPassword(e.target.value)} required minLength={6}
              style={{ width:"100%", background:"#0A0D14", border:"1px solid #1F2937", borderRadius:8, color:"#F9FAFB",
                padding:"12px 14px", fontSize:14, fontFamily:"'IBM Plex Mono',monospace", outline:"none" }}
              placeholder="mínimo 6 caracteres" />
          </div>
          <button type="submit" disabled={loading} style={{
            width:"100%", padding:"14px", borderRadius:10, cursor:loading?"wait":"pointer", border:"none",
            background:"linear-gradient(135deg,#92400E,#F59E0B)", color:"#0F0F14",
            fontSize:15, fontFamily:"'Bebas Neue',sans-serif", letterSpacing:"0.2em" }}>
            {loading ? "A processar..." : mode === "login" ? "ENTRAR" : "CRIAR CONTA"}
          </button>
        </form>

        <div style={{ marginTop:20, textAlign:"center", fontSize:12, color:"#6B7280", fontFamily:"'IBM Plex Mono',monospace" }}>
          {mode === "login" ? (
            <>Não tens conta? <button type="button" onClick={()=>setMode("signup")} style={{ background:"none", border:"none", color:"#F59E0B", cursor:"pointer", fontFamily:"'IBM Plex Mono',monospace" }}>Regista-te</button></>
          ) : (
            <>Já tens conta? <button type="button" onClick={()=>setMode("login")} style={{ background:"none", border:"none", color:"#F59E0B", cursor:"pointer", fontFamily:"'IBM Plex Mono',monospace" }}>Entra</button></>
          )}
        </div>
      </div>
    </div>
  );
};

// ─── EXPORTS ──────────────────────────────────────────────────────────────────
const doExportJSON = (trades,props,schedule) => {
  const blob=new Blob([JSON.stringify({trades,props,schedule,exportedAt:new Date().toISOString()},null,2)],{type:"application/json"});
  const a=document.createElement("a"); a.href=URL.createObjectURL(blob);
  a.download=`propdesk_${new Date().toISOString().split("T")[0]}.json`;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
};
const doExportXLSX = (trades,props) => {
  const wb=XLSX.utils.book_new();
  const rows=trades.map(t=>{
    const p=props.find(x=>x.id===t.propId);
    return { Date:t.date, Prop:p?.name||"", Market:t.market, Session:t.session,
      Direction:t.direction, Setup:t.setup, "Weekly Bias":t.weeklyBias, "Daily Bias":t.bias,
      "H1 Structure":t.h1Structure, "H1 Position":t.h1Position,
      Result:t.result, "Result R":t.resultR, "Account %":t.resultPct,
      Grade:t.grade, "Setup ★":t.setupRating, "Execution ★":t.executionRating,
      "Management ★":t.managementRating, "Rules Followed":t.rulesFollowed===true?"YES":t.rulesFollowed===false?"NO":"",
      "Pre Emotions":(t.preEmotions||[]).join(", "), "During Emotions":(t.duringEmotions||[]).join(", "),
      "Key Lesson":t.lesson, "Went Well":t.wentWell, "To Improve":t.toImprove };
  });
  XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows.length?rows:[{Note:"No trades"}]),"Trades");
  const summary=props.map(p=>{
    const pt=trades.filter(t=>t.propId===p.id);
    const wins=pt.filter(t=>t.result==="WIN").length;
    const totalPct=pt.reduce((s,t)=>s+parseFloat(t.resultPct||0),0);
    const totalR=pt.reduce((s,t)=>s+parseFloat(t.resultR||0),0);
    return { Prop:p.name,"Account $":p.accountSize,"Max DD $":p.maxDD,"Risk/Trade %":p.riskPct,
      "Weekly Target %":p.weeklyTarget,Phase:p.phase,"Total Trades":pt.length,
      "Win Rate %":pt.length>0?Math.round((wins/pt.length)*100):0,
      "Total R":totalR.toFixed(2),"Total P&L %":totalPct.toFixed(2),
      "Total P&L $":((totalPct/100)*p.accountSize).toFixed(2) };
  });
  XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(summary),"Prop Summary");
  XLSX.writeFile(wb,`propdesk_${new Date().toISOString().split("T")[0]}.xlsx`);
};

// ─── STAR RATING ──────────────────────────────────────────────────────────────
const StarRating = ({ value, onChange }) => (
  <div style={{ display:"flex",gap:4 }}>
    {[1,2,3,4,5].map(s=>(
      <button key={s} type="button" onClick={()=>onChange(s)} style={{
        background:"none",border:"none",cursor:"pointer",padding:"2px 3px",fontSize:22,
        color:s<=value?"#F59E0B":"#374151",transition:"color 0.15s" }}>★</button>
    ))}
  </div>
);

// ─── EMOTION PILL ─────────────────────────────────────────────────────────────
const EmotionPill = ({ label, selected, onClick }) => {
  const neg=["Hesitant","Anxious","FOMO","Greedy","Revenge","Impulsive","Overconfident"].includes(label);
  return (
    <button type="button" onClick={onClick} style={{
      padding:"4px 10px",borderRadius:20,fontSize:11,fontWeight:600,cursor:"pointer",
      border:selected?"none":"1px solid #374151",
      background:selected?(neg?"#7F1D1D":"#064E3B"):"transparent",
      color:selected?(neg?"#FCA5A5":"#6EE7B7"):"#9CA3AF",
      transition:"all 0.15s",fontFamily:"'IBM Plex Mono',monospace" }}>
      {label}
    </button>
  );
};

// ─── CHOICE PILLS ─────────────────────────────────────────────────────────────
const ChoicePills = ({ options, value, onChange, color="#818CF8", size=11 }) => (
  <div style={{ display:"flex",gap:6,flexWrap:"wrap" }}>
    {options.map(o=>(
      <button key={o} type="button" onClick={()=>onChange(o===value?null:o)} style={{
        padding:`${size===11?4:6}px ${size===11?10:14}px`,borderRadius:20,cursor:"pointer",
        fontSize:size,fontFamily:"'IBM Plex Mono',monospace",fontWeight:600,transition:"all 0.15s",
        border:`1px solid ${value===o?color:"#1F2937"}`,
        background:value===o?`${color}20`:"transparent",
        color:value===o?color:"#6B7280" }}>{o}</button>
    ))}
  </div>
);

// ─── PROP CARD ────────────────────────────────────────────────────────────────
const PropCard = ({ prop, trades, isToday }) => {
  const week = getWeekDates();
  const allTrades = trades.filter(t=>t.propId===prop.id);
  const weekTrades = allTrades.filter(t=>week.includes(t.date));
  const weekPct = weekTrades.reduce((s,t)=>s+parseFloat(t.resultPct||0),0);
  const totalPct = allTrades.reduce((s,t)=>s+parseFloat(t.resultPct||0),0);
  const weekDollar = (weekPct/100)*prop.accountSize;
  const progress = Math.min(Math.max((weekPct/prop.weeklyTarget)*100,0),100);

  let running=0, maxDD=0;
  allTrades.forEach(t=>{ running+=parseFloat(t.resultPct||0); if(running<maxDD)maxDD=running; });
  const ddUsedPct = Math.abs(maxDD);
  const ddMax = (prop.maxDD/prop.accountSize)*100;
  const ddRatio = Math.min((ddUsedPct/ddMax)*100,100);
  const ddStatus = ddRatio>=80?"#EF4444":ddRatio>=50?"#F59E0B":"#10B981";

  const wins=weekTrades.filter(t=>t.result==="WIN").length;
  const wr=weekTrades.length>0?Math.round((wins/weekTrades.length)*100):null;

  let cum=0;
  const spark = allTrades.slice(-8).map((t,i)=>{cum+=parseFloat(t.resultR||0);return{i,v:parseFloat(cum.toFixed(2))};});

  return (
    <div style={{ background:isToday?"linear-gradient(135deg,#0F172A,#1E1B4B)":"#0D1117",
      border:`1px solid ${isToday?prop.color:"#1F2937"}`,borderRadius:12,padding:"16px 18px",
      position:"relative",overflow:"hidden",boxShadow:isToday?`0 0 24px ${prop.color}33`:"none",
      transition:"all 0.2s" }}>
      {isToday&&<div style={{ position:"absolute",top:9,right:11,fontSize:9,fontFamily:"'IBM Plex Mono',monospace",
        color:prop.color,letterSpacing:"0.15em",fontWeight:700 }}>◆ TODAY</div>}

      <div style={{ display:"flex",alignItems:"center",gap:8,marginBottom:10 }}>
        <div style={{ width:8,height:8,borderRadius:"50%",background:prop.color,flexShrink:0,
          boxShadow:`0 0 6px ${prop.color}` }}/>
        <span style={{ fontFamily:"'Bebas Neue',sans-serif",fontSize:18,letterSpacing:"0.1em",color:"#F9FAFB" }}>{prop.name}</span>
        <span style={{ fontSize:9,color:"#4B5563",fontFamily:"'IBM Plex Mono',monospace",marginLeft:"auto" }}>{prop.phase.toUpperCase()}</span>
      </div>

      <div style={{ marginBottom:8 }}>
        <div style={{ display:"flex",justifyContent:"space-between",marginBottom:3 }}>
          <span style={{ fontSize:9,color:"#6B7280",fontFamily:"'IBM Plex Mono',monospace",letterSpacing:"0.1em" }}>WEEK</span>
          <span style={{ fontSize:11,fontFamily:"'IBM Plex Mono',monospace",fontWeight:700,
            color:weekPct>=0?"#10B981":"#EF4444" }}>{fmtPct(weekPct)} / {prop.weeklyTarget}%</span>
        </div>
        <div style={{ background:"#1F2937",borderRadius:3,height:5,overflow:"hidden" }}>
          <div style={{ height:"100%",width:`${progress}%`,borderRadius:3,transition:"width 0.5s",
            background:progress>=100?"#10B981":progress>=60?prop.color:"#374151" }}/>
        </div>
      </div>

      <div style={{ marginBottom:12 }}>
        <div style={{ display:"flex",justifyContent:"space-between",marginBottom:3 }}>
          <span style={{ fontSize:9,color:"#6B7280",fontFamily:"'IBM Plex Mono',monospace",letterSpacing:"0.1em" }}>DRAWDOWN</span>
          <span style={{ fontSize:10,fontFamily:"'IBM Plex Mono',monospace",fontWeight:700,color:ddStatus }}>
            {ddUsedPct.toFixed(2)}% / {ddMax.toFixed(2)}%
          </span>
        </div>
        <div style={{ background:"#1F2937",borderRadius:3,height:4,overflow:"hidden" }}>
          <div style={{ height:"100%",width:`${ddRatio}%`,borderRadius:3,background:ddStatus,transition:"width 0.5s" }}/>
        </div>
      </div>

      <div style={{ display:"flex",gap:8,marginBottom:10 }}>
        {[
          { l:"WEEK $", v:weekDollar>=0?`+$${Math.abs(weekDollar).toFixed(0)}`:`-$${Math.abs(weekDollar).toFixed(0)}`, c:weekDollar>=0?"#10B981":"#EF4444" },
          { l:"ALL-TIME", v:fmtPct(totalPct), c:totalPct>=0?"#10B981":"#EF4444" },
          { l:"WIN%", v:wr!==null?`${wr}%`:"—", c:wr>=50?"#10B981":wr!=null&&wr<50?"#EF4444":"#6B7280" },
          { l:"TRADES W", v:weekTrades.length, c:"#9CA3AF" },
        ].map(i=>(
          <div key={i.l} style={{ flex:1 }}>
            <div style={{ fontSize:8,color:"#4B5563",fontFamily:"'IBM Plex Mono',monospace",letterSpacing:"0.1em",marginBottom:1 }}>{i.l}</div>
            <div style={{ fontSize:12,fontFamily:"'IBM Plex Mono',monospace",fontWeight:700,color:i.c }}>{i.v}</div>
          </div>
        ))}
      </div>

      {spark.length>1&&(
        <div style={{ marginTop:4 }}>
          <ResponsiveContainer width="100%" height={36}>
            <AreaChart data={spark} margin={{top:2,bottom:2,left:0,right:0}}>
              <defs>
                <linearGradient id={`sg${prop.id}`} x1="0" y1="0" x2="0" y2="1">
                  <stop offset="5%" stopColor={prop.color} stopOpacity={0.3}/>
                  <stop offset="95%" stopColor={prop.color} stopOpacity={0}/>
                </linearGradient>
              </defs>
              <Area type="monotone" dataKey="v" stroke={prop.color} fill={`url(#sg${prop.id})`}
                strokeWidth={1.5} dot={false}/>
            </AreaChart>
          </ResponsiveContainer>
        </div>
      )}
    </div>
  );
};

// ─── TRADE FORM (igual ao original) ───────────────────────────────────────────
const makeBlank = (propId, date) => ({
  id:"", date:date||new Date().toISOString().split("T")[0], propId:propId||"",
  market:"ES", session:"NY AM (08-11)", direction:"", setup:"",
  weeklyBias:"", weeklyNarrative:"", weeklyKeyLevel:"",
  h1Narrative:"", h1Structure:"", h1Position:"", h1KeyLevel:"",
  tf5mVisible:"", tf5mNote:"", tf2mVisible:"", tf2mNote:"", tf1mVisible:"", tf1mNote:"",
  bias:"", confluences:"", entryReason:"",
  entry:"", sl:"", tp:"", plannedR:"3", result:"", resultR:"", resultPct:"",
  preEmotions:[], duringEmotions:[], postEmotions:[],
  executionRating:0, setupRating:0, managementRating:0,
  rulesFollowed:null, wentWell:"", toImprove:"", lesson:"", mistakes:"", screenNote:"", grade:"",
});

const TradeForm = ({ props, schedule, onSave, editTrade, onCancel, screenW, showToast }) => {
  const todayDay = getTodayDayKey();
  const todayIds = normDay(schedule, todayDay);
  const defaultPropId = todayIds[0] || props[0]?.id || "";
  const [form, setForm] = useState(()=>editTrade?{...makeBlank(defaultPropId),...editTrade}:makeBlank(defaultPropId));
  const [errors, setErrors] = useState({});
  const currentProp = props.find(p=>p.id===form.propId);

  useEffect(()=>{
    if(editTrade) setForm({...makeBlank(defaultPropId),...editTrade});
    else setForm(makeBlank(defaultPropId));
    setErrors({});
  },[editTrade?.id]);

  const set = useCallback((k,v)=>{ setForm(f=>({...f,[k]:v})); setErrors(e=>({...e,[k]:false})); },[]);
  const toggleEm = useCallback((ph,em)=>setForm(f=>({...f,[ph]:f[ph].includes(em)?f[ph].filter(x=>x!==em):[...f[ph],em]})),[]);

  const calcR = () => {
    const e=parseFloat(form.entry),s=parseFloat(form.sl),t=parseFloat(form.tp);
    if(!e||!s||!t) return null;
    const risk=Math.abs(e-s),reward=Math.abs(t-e);
    return risk===0?null:(reward/risk).toFixed(2);
  };

  const autoCalcPct = () => {
    if(!form.resultR||!currentProp) return null;
    return (parseFloat(form.resultR)*currentProp.riskPct).toFixed(3);
  };

  const handleSave = () => {
    const errs={};
    if(!form.date) errs.date=true;
    if(!form.propId) errs.propId=true;
    if(!form.result) errs.result=true;
    if(!form.direction) errs.direction=true;
    if(Object.keys(errs).length>0){ setErrors(errs); showToast("Fill required: Date, Prop, Direction, Result","error"); return; }
    const finalForm = {...form};
    if(!finalForm.resultPct && finalForm.resultR && currentProp) {
      finalForm.resultPct = autoCalcPct();
    }
    onSave({...finalForm,id:editTrade?.id||Date.now().toString()});
    showToast(editTrade?"Trade updated ✓":"Trade saved ✓","success");
  };

  const iS = { background:"#0D1117",borderRadius:8,color:"#F9FAFB",padding:"8px 12px",
    fontSize:13,width:"100%",fontFamily:"'IBM Plex Mono',monospace",outline:"none" };
  const fld = (k) => ({...iS,border:`1px solid ${errors[k]?"#EF4444":"#1F2937"}`});
  const lS = { fontSize:10,color:"#6B7280",fontFamily:"'IBM Plex Mono',monospace",
    letterSpacing:"0.12em",fontWeight:700,marginBottom:4,display:"block" };
  const sec = (n,t,col="#4B5563") => (
    <div style={{ fontSize:11,color:col,fontFamily:"'IBM Plex Mono',monospace",letterSpacing:"0.2em",
      fontWeight:700,padding:"14px 0 10px",borderBottom:`1px solid #1F2937`,marginBottom:14 }}>
      {n} — {t}
    </div>
  );

  const c2=screenW<640?"1fr":"1fr 1fr";
  const c3=screenW<640?"1fr":screenW<960?"1fr 1fr":"1fr 1fr 1fr";
  const c4=screenW<640?"1fr 1fr":screenW<960?"1fr 1fr":"1fr 1fr 1fr 1fr";

  return (
    <div style={{ color:"#F9FAFB" }}>
      <div style={{ display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:24,flexWrap:"wrap",gap:10 }}>
        <h2 style={{ fontFamily:"'Bebas Neue',sans-serif",fontSize:28,letterSpacing:"0.1em",
          background:"linear-gradient(90deg,#F59E0B,#FDE68A)",WebkitBackgroundClip:"text",WebkitTextFillColor:"transparent" }}>
          {editTrade?"EDIT TRADE":"NEW TRADE ENTRY"}
        </h2>
        {onCancel&&<button type="button" onClick={onCancel} style={{ background:"none",border:"1px solid #374151",
          color:"#6B7280",borderRadius:8,padding:"6px 14px",cursor:"pointer",fontSize:12,
          fontFamily:"'IBM Plex Mono',monospace" }}>CANCEL</button>}
      </div>

      {sec("01","TRADE META")}
      <div style={{ display:"grid",gridTemplateColumns:c4,gap:12,marginBottom:14 }}>
        <div><label style={lS}>DATE {errors.date&&<span style={{color:"#EF4444"}}>*</span>}</label>
          <input type="date" value={form.date} onChange={e=>set("date",e.target.value)} style={fld("date")}/></div>
        <div><label style={lS}>PROP FIRM {errors.propId&&<span style={{color:"#EF4444"}}>*</span>}</label>
          <select value={form.propId} onChange={e=>set("propId",e.target.value)} style={fld("propId")}>
            <option value="">Select...</option>
            {props.filter(p=>p.phase!=="Disabled").map(p=><option key={p.id} value={p.id}>{p.name}</option>)}
          </select></div>
        <div><label style={lS}>MARKET</label>
          <select value={form.market} onChange={e=>set("market",e.target.value)} style={fld("market")}>
            {MARKETS.map(m=><option key={m}>{m}</option>)}</select></div>
        <div><label style={lS}>SESSION</label>
          <select value={form.session} onChange={e=>set("session",e.target.value)} style={fld("session")}>
            {SESSIONS.map(s=><option key={s}>{s}</option>)}</select></div>
      </div>
      <div style={{ display:"grid",gridTemplateColumns:c3,gap:12,marginBottom:14 }}>
        <div><label style={lS}>DIRECTION {errors.direction&&<span style={{color:"#EF4444"}}>*</span>}</label>
          <div style={{ display:"flex",gap:8 }}>
            {["LONG","SHORT"].map(d=>(
              <button key={d} type="button" onClick={()=>set("direction",d)} style={{
                flex:1,padding:"8px",borderRadius:8,cursor:"pointer",fontSize:12,
                fontFamily:"'IBM Plex Mono',monospace",fontWeight:700,transition:"all 0.15s",
                border:`1px solid ${errors.direction?"#EF4444":form.direction===d?(d==="LONG"?"#10B981":"#EF4444"):"#1F2937"}`,
                background:form.direction===d?(d==="LONG"?"#064E3B":"#7F1D1D"):"transparent",
                color:form.direction===d?(d==="LONG"?"#6EE7B7":"#FCA5A5"):"#6B7280" }}>
                {d==="LONG"?"↑ LONG":"↓ SHORT"}
              </button>))}
          </div></div>
        <div><label style={lS}>SETUP</label>
          <select value={form.setup} onChange={e=>set("setup",e.target.value)} style={fld("setup")}>
            <option value="">Select setup...</option>
            {SETUPS.map(s=><option key={s}>{s}</option>)}</select></div>
        <div><label style={lS}>RESULT {errors.result&&<span style={{color:"#EF4444"}}>*</span>}</label>
          <div style={{ display:"flex",gap:6 }}>
            {["WIN","LOSS","BE"].map(r=>(
              <button key={r} type="button" onClick={()=>set("result",r)} style={{
                flex:1,padding:"8px 4px",borderRadius:8,cursor:"pointer",fontSize:11,
                fontFamily:"'IBM Plex Mono',monospace",fontWeight:700,transition:"all 0.15s",
                border:`1px solid ${errors.result?"#EF4444":form.result===r?(r==="WIN"?"#10B981":r==="LOSS"?"#EF4444":"#F59E0B"):"#1F2937"}`,
                background:form.result===r?(r==="WIN"?"#064E3B":r==="LOSS"?"#7F1D1D":"#78350F"):"transparent",
                color:form.result===r?(r==="WIN"?"#6EE7B7":r==="LOSS"?"#FCA5A5":"#FDE68A"):"#6B7280" }}>{r}</button>
            ))}</div></div>
      </div>

      {sec("02","MULTI-TIMEFRAME ANALYSIS","#818CF8")}

      <div style={{ background:"#0A0D14",border:"1px solid #1a2035",borderRadius:10,padding:"14px 16px",marginBottom:12 }}>
        <div style={{ fontSize:9,color:"#818CF8",fontFamily:"'IBM Plex Mono',monospace",letterSpacing:"0.2em",
          fontWeight:700,marginBottom:10 }}>WEEKLY CONTEXT</div>
        <div style={{ display:"grid",gridTemplateColumns:c2,gap:12 }}>
          <div>
            <label style={lS}>WEEKLY BIAS</label>
            <ChoicePills options={BIASES} value={form.weeklyBias} onChange={v=>set("weeklyBias",v)} color="#818CF8"/>
          </div>
          <div>
            <label style={lS}>WEEKLY KEY LEVEL</label>
            <input value={form.weeklyKeyLevel} onChange={e=>set("weeklyKeyLevel",e.target.value)}
              style={fld("weeklyKeyLevel")} placeholder="e.g. 5220 weekly OB, 5180 NWOG..."/>
          </div>
        </div>
        <div style={{ marginTop:10 }}>
          <label style={lS}>WEEKLY NARRATIVE</label>
          <textarea value={form.weeklyNarrative} onChange={e=>set("weeklyNarrative",e.target.value)} rows={2}
            style={{...fld("weeklyNarrative"),resize:"vertical"}}
            placeholder="Weekly structure, higher-timeframe draw, expected weekly profile..."/>
        </div>
      </div>

      <div style={{ background:"#0A0D14",border:"1px solid #1a2035",borderRadius:10,padding:"14px 16px",marginBottom:12 }}>
        <div style={{ fontSize:9,color:"#F59E0B",fontFamily:"'IBM Plex Mono',monospace",letterSpacing:"0.2em",
          fontWeight:700,marginBottom:10 }}>H1 — 1 HOUR TIMEFRAME</div>
        <div style={{ marginBottom:10 }}>
          <label style={lS}>WHAT IS HAPPENING ON H1</label>
          <textarea value={form.h1Narrative} onChange={e=>set("h1Narrative",e.target.value)} rows={2}
            style={{...fld("h1Narrative"),resize:"vertical"}}
            placeholder="Price respecting H1 OB, displacement into FVG, CHoCH confirmed at 5250..."/>
        </div>
        <div style={{ display:"grid",gridTemplateColumns:c3,gap:12 }}>
          <div>
            <label style={lS}>STRUCTURE CONFIRMS BIAS?</label>
            <ChoicePills options={H1_STRUCT} value={form.h1Structure} onChange={v=>set("h1Structure",v)} color="#F59E0B"/>
          </div>
          <div>
            <label style={lS}>PRICE POSITION ON H1</label>
            <ChoicePills options={H1_POS} value={form.h1Position} onChange={v=>set("h1Position",v)}
              color={form.h1Position==="Discount"?"#10B981":form.h1Position==="Premium"?"#EF4444":"#F59E0B"}/>
          </div>
          <div>
            <label style={lS}>H1 KEY LEVEL</label>
            <input value={form.h1KeyLevel} onChange={e=>set("h1KeyLevel",e.target.value)}
              style={fld("h1KeyLevel")} placeholder="FVG / OB level..."/>
          </div>
        </div>
      </div>

      <div style={{ background:"#0A0D14",border:"1px solid #1a2035",borderRadius:10,padding:"14px 16px",marginBottom:14 }}>
        <div style={{ fontSize:9,color:"#10B981",fontFamily:"'IBM Plex Mono',monospace",letterSpacing:"0.2em",
          fontWeight:700,marginBottom:10 }}>ENTRY TIMEFRAMES</div>
        {[
          {tf:"5m",visK:"tf5mVisible",noteK:"tf5mNote",label:"5M",hint:"Setup visible on 5M? Describe OB/FVG position..."},
          {tf:"2m",visK:"tf2mVisible",noteK:"tf2mNote",label:"2M",hint:"Confirmation on 2M? Price behaviour into level..."},
          {tf:"1m",visK:"tf1mVisible",noteK:"tf1mNote",label:"1M",hint:"1M trigger — displacement, MSS, exact entry signal..."},
        ].map(({visK,noteK,label,hint})=>(
          <div key={label} style={{ display:"grid",gridTemplateColumns:`60px 200px 1fr`,gap:12,alignItems:"start",marginBottom:10 }}>
            <div style={{ paddingTop:8 }}>
              <div style={{ fontSize:13,fontFamily:"'Bebas Neue',sans-serif",letterSpacing:"0.15em",
                color:"#10B981",marginBottom:4 }}>{label}</div>
              <ChoicePills options={TF_CONFIRM} value={form[visK]} onChange={v=>set(visK,v)} color="#10B981" size={9}/>
            </div>
            <div style={{ gridColumn:"2/4" }}>
              <textarea value={form[noteK]} onChange={e=>set(noteK,e.target.value)} rows={2}
                style={{...fld(noteK),resize:"vertical",width:"100%"}} placeholder={hint}/>
            </div>
          </div>
        ))}
      </div>

      {sec("03","PRICE LEVELS & R")}
      <div style={{ display:"grid",gridTemplateColumns:screenW<640?"1fr 1fr":"repeat(6,1fr)",gap:12,marginBottom:8 }}>
        {[{k:"entry",l:"ENTRY"},{k:"sl",l:"STOP LOSS"},{k:"tp",l:"TAKE PROFIT"},
          {k:"plannedR",l:"PLANNED R"},{k:"resultR",l:"RESULT R"},{k:"resultPct",l:"ACCOUNT %"}].map(({k,l})=>(
          <div key={k}><label style={lS}>{l}</label>
            <input type="number" step="any" value={form[k]} onChange={e=>set(k,e.target.value)}
              style={fld(k)} placeholder={k==="plannedR"?"3":"0.00"}/></div>
        ))}
      </div>
      <div style={{ display:"flex",gap:14,marginBottom:14,flexWrap:"wrap" }}>
        {calcR()&&<div style={{ fontSize:11,color:"#6B7280",fontFamily:"'IBM Plex Mono',monospace",
          padding:"6px 12px",background:"#111827",borderRadius:6,border:"1px solid #1F2937" }}>
          AUTO R:R → <span style={{color:"#F59E0B",fontWeight:700}}>{calcR()}R</span>
        </div>}
        {autoCalcPct()&&!form.resultPct&&<div style={{ fontSize:11,color:"#6B7280",fontFamily:"'IBM Plex Mono',monospace",
          padding:"6px 12px",background:"#111827",borderRadius:6,border:"1px solid #1F2937" }}>
          AUTO P&L → <span style={{color:parseFloat(autoCalcPct())>=0?"#10B981":"#EF4444",fontWeight:700}}>
            {fmtPct(autoCalcPct())}
          </span> ({currentProp?.riskPct}% risk/trade)
        </div>}
      </div>

      {sec("04","DAILY BIAS & CONFLUENCES")}
      <div style={{ marginBottom:12 }}>
        <label style={lS}>DAILY BIAS</label>
        <ChoicePills options={BIASES} value={form.bias} onChange={v=>set("bias",v)} color="#A5B4FC" size={12}/>
      </div>
      {[{k:"confluences",l:"CONFLUENCES",p:"FVG + OB alignment, Tuesday profile, Phase A regime..."},
        {k:"entryReason",l:"ENTRY REASON (LTF trigger)",p:"M1 displacement, MSS confirmed, entered at 50% FVG..."}
      ].map(({k,l,p})=>(
        <div key={k} style={{marginBottom:12}}>
          <label style={lS}>{l}</label>
          <textarea value={form[k]} onChange={e=>set(k,e.target.value)} rows={2}
            style={{...fld(k),resize:"vertical"}} placeholder={p}/>
        </div>
      ))}

      {sec("05","PSYCHOLOGY & EMOTIONS")}
      {[{ph:"preEmotions",l:"PRE-TRADE"},{ph:"duringEmotions",l:"DURING TRADE"},{ph:"postEmotions",l:"POST-TRADE"}].map(({ph,l})=>(
        <div key={ph} style={{marginBottom:14}}>
          <label style={lS}>{l}</label>
          <div style={{display:"flex",flexWrap:"wrap",gap:6}}>
            {EMOTIONS.map(em=><EmotionPill key={em} label={em} selected={form[ph].includes(em)} onClick={()=>toggleEm(ph,em)}/>)}
          </div>
        </div>
      ))}

      {sec("06","RATINGS & COMPLIANCE")}
      <div style={{display:"grid",gridTemplateColumns:c3,gap:16,marginBottom:14}}>
        {[{k:"setupRating",l:"SETUP QUALITY"},{k:"executionRating",l:"EXECUTION"},{k:"managementRating",l:"MANAGEMENT"}].map(({k,l})=>(
          <div key={k}><label style={lS}>{l}</label><StarRating value={form[k]} onChange={v=>set(k,v)}/></div>
        ))}
      </div>
      <div style={{display:"flex",gap:20,flexWrap:"wrap",marginBottom:14,alignItems:"flex-start"}}>
        <div>
          <label style={lS}>RULES FOLLOWED?</label>
          <div style={{display:"flex",gap:8}}>
            {[{v:true,l:"✓ YES"},{v:false,l:"✗ NO"}].map(o=>(
              <button key={String(o.v)} type="button" onClick={()=>set("rulesFollowed",o.v)} style={{
                padding:"8px 16px",borderRadius:8,cursor:"pointer",fontSize:12,
                fontFamily:"'IBM Plex Mono',monospace",fontWeight:700,transition:"all 0.15s",
                border:`1px solid ${form.rulesFollowed===o.v?(o.v?"#10B981":"#EF4444"):"#1F2937"}`,
                background:form.rulesFollowed===o.v?(o.v?"#064E3B":"#7F1D1D"):"transparent",
                color:form.rulesFollowed===o.v?(o.v?"#6EE7B7":"#FCA5A5"):"#6B7280" }}>{o.l}</button>
            ))}
          </div>
        </div>
        <div>
          <label style={lS}>TRADE GRADE</label>
          <div style={{display:"flex",gap:6}}>
            {["A+","A","B","C","D","F"].map(g=>(
              <button key={g} type="button" onClick={()=>set("grade",g)} style={{
                width:40,height:34,borderRadius:8,cursor:"pointer",fontSize:14,
                fontFamily:"'Bebas Neue',sans-serif",letterSpacing:"0.1em",transition:"all 0.15s",
                border:`1px solid ${form.grade===g?"#818CF8":"#1F2937"}`,
                background:form.grade===g?"#1E1B4B":"transparent",
                color:form.grade===g?"#A5B4FC":"#6B7280" }}>{g}</button>
            ))}
          </div>
        </div>
      </div>

      {sec("07","POST-TRADE REVIEW")}
      <div style={{display:"grid",gridTemplateColumns:c2,gap:12,marginBottom:12}}>
        <div><label style={lS}>✓ WHAT WENT WELL</label>
          <textarea value={form.wentWell} onChange={e=>set("wentWell",e.target.value)} rows={3}
            style={{...fld("wentWell"),resize:"vertical"}} placeholder="Waited for confirmation, respected SL..."/></div>
        <div><label style={lS}>✗ WHAT TO IMPROVE</label>
          <textarea value={form.toImprove} onChange={e=>set("toImprove",e.target.value)} rows={3}
            style={{...fld("toImprove"),resize:"vertical"}} placeholder="Entered too early, moved SL..."/></div>
      </div>
      <div style={{display:"grid",gridTemplateColumns:c2,gap:12,marginBottom:12}}>
        <div><label style={lS}>★ KEY LESSON</label>
          <textarea value={form.lesson} onChange={e=>set("lesson",e.target.value)} rows={2}
            style={{...fld("lesson"),resize:"vertical"}} placeholder="One concrete takeaway..."/></div>
        <div><label style={lS}>⚠ MISTAKES / RULES BROKEN</label>
          <textarea value={form.mistakes} onChange={e=>set("mistakes",e.target.value)} rows={2}
            style={{...fld("mistakes"),resize:"vertical"}} placeholder="Rule breach and why..."/></div>
      </div>
      <div style={{marginBottom:24}}>
        <label style={lS}>📎 CHART NOTE</label>
        <textarea value={form.screenNote} onChange={e=>set("screenNote",e.target.value)} rows={2}
          style={{...fld("screenNote"),resize:"vertical"}} placeholder="Key levels, annotations, chart context..."/>
      </div>
      <button type="button" onClick={handleSave} style={{
        width:"100%",padding:"15px",borderRadius:10,cursor:"pointer",
        background:"linear-gradient(135deg,#92400E,#F59E0B)",border:"none",
        color:"#0F0F14",fontSize:15,fontFamily:"'Bebas Neue',sans-serif",letterSpacing:"0.2em" }}>
        {editTrade?"UPDATE TRADE":"SAVE TRADE ENTRY"}
      </button>
    </div>
  );
};

// ─── TRADE ROW ────────────────────────────────────────────────────────────────
const TradeRow = ({ trade, props, onEdit, onDeleteRequest }) => {
  const [exp, setExp] = useState(false);
  const prop = props.find(p=>p.id===trade.propId);
  const rNum = parseFloat(trade.resultR);
  const avg = [trade.setupRating,trade.executionRating,trade.managementRating]
    .filter(Boolean).reduce((a,b,_,arr)=>a+b/arr.length,0);

  const mtfBadge = (val) => {
    if(!val) return null;
    const colors = { "Confirms ✓":"#10B981","Partial ~":"#F59E0B","Against ✗":"#EF4444",
      "Discount":"#10B981","Equilibrium":"#F59E0B","Premium":"#EF4444",
      "YES ✓":"#10B981","PARTIAL ~":"#F59E0B","NO ✗":"#EF4444" };
    const c = colors[val]||"#6B7280";
    return <span style={{fontSize:9,padding:"2px 6px",borderRadius:4,background:`${c}20`,
      color:c,fontFamily:"'IBM Plex Mono',monospace",fontWeight:700,marginRight:4}}>{val}</span>;
  };

  return (
    <div style={{marginBottom:8}}>
      <div onClick={()=>setExp(!exp)} style={{
        display:"grid",gridTemplateColumns:"88px 75px 55px 65px 90px 58px 62px 50px 55px 24px",
        gap:6,alignItems:"center",padding:"10px 14px",cursor:"pointer",
        background:exp?"#111827":"#0D1117",
        border:`1px solid ${exp?"#374151":"#1F2937"}`,
        borderRadius:exp?"10px 10px 0 0":10 }}>
        <span style={{fontSize:11,color:"#9CA3AF",fontFamily:"'IBM Plex Mono',monospace",whiteSpace:"nowrap"}}>{trade.date}</span>
        <span style={{fontSize:12,fontFamily:"'IBM Plex Mono',monospace",fontWeight:700,color:prop?.color||"#9CA3AF"}}>{prop?.name||"—"}</span>
        <span style={{fontSize:11,color:"#9CA3AF",fontFamily:"'IBM Plex Mono',monospace"}}>{trade.market}</span>
        <span style={{fontSize:12,fontFamily:"'IBM Plex Mono',monospace",fontWeight:700,
          color:trade.direction==="LONG"?"#6EE7B7":"#FCA5A5",whiteSpace:"nowrap"}}>
          {trade.direction==="LONG"?"↑ L":"↓ S"}
        </span>
        <span style={{fontSize:11,color:"#9CA3AF",fontFamily:"'IBM Plex Mono',monospace",overflow:"hidden",whiteSpace:"nowrap",textOverflow:"ellipsis"}}>{trade.setup||"—"}</span>
        <span style={{fontSize:13,fontFamily:"'IBM Plex Mono',monospace",fontWeight:700,
          color:trade.result==="WIN"?"#10B981":trade.result==="LOSS"?"#EF4444":"#F59E0B"}}>{trade.result||"—"}</span>
        <span style={{fontSize:13,fontFamily:"'IBM Plex Mono',monospace",fontWeight:700,
          color:rNum>0?"#10B981":rNum<0?"#EF4444":"#9CA3AF",whiteSpace:"nowrap"}}>
          {trade.resultR?fmtR(rNum):"—"}
        </span>
        <span style={{fontSize:12,fontFamily:"'IBM Plex Mono',monospace",color:"#A5B4FC"}}>{trade.grade||"—"}</span>
        <span style={{fontSize:11,fontFamily:"'IBM Plex Mono',monospace",color:rColor(avg)}}>{avg>0?`${avg.toFixed(1)}★`:"—"}</span>
        <span style={{fontSize:11,color:"#4B5563"}}>{exp?"▲":"▼"}</span>
      </div>
      {exp&&(
        <div style={{background:"#080E17",border:"1px solid #1F2937",borderTop:"none",
          borderRadius:"0 0 10px 10px",padding:16}}>

          {(trade.weeklyBias||trade.h1Structure||trade.h1Position)&&(
            <div style={{display:"flex",gap:6,marginBottom:12,flexWrap:"wrap",alignItems:"center"}}>
              <span style={{fontSize:9,color:"#4B5563",fontFamily:"'IBM Plex Mono',monospace",letterSpacing:"0.1em",marginRight:4}}>MTF:</span>
              {trade.weeklyBias&&<span style={{fontSize:9,padding:"2px 6px",borderRadius:4,background:"#818CF820",
                color:"#818CF8",fontFamily:"'IBM Plex Mono',monospace",fontWeight:700}}>W: {trade.weeklyBias}</span>}
              {trade.bias&&<span style={{fontSize:9,padding:"2px 6px",borderRadius:4,background:"#A5B4FC20",
                color:"#A5B4FC",fontFamily:"'IBM Plex Mono',monospace",fontWeight:700}}>D: {trade.bias}</span>}
              {mtfBadge(trade.h1Structure)} {mtfBadge(trade.h1Position)}
              {trade.tf5mVisible&&<span style={{fontSize:9,color:"#6B7280",fontFamily:"'IBM Plex Mono',monospace"}}>5M: {trade.tf5mVisible}</span>}
              {trade.tf2mVisible&&<span style={{fontSize:9,color:"#6B7280",fontFamily:"'IBM Plex Mono',monospace"}}>2M: {trade.tf2mVisible}</span>}
              {trade.tf1mVisible&&<span style={{fontSize:9,color:"#6B7280",fontFamily:"'IBM Plex Mono',monospace"}}>1M: {trade.tf1mVisible}</span>}
            </div>
          )}

          {(trade.h1Narrative||trade.weeklyNarrative||trade.confluences||trade.entryReason)&&(
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(200px,1fr))",gap:12,marginBottom:12}}>
              {trade.weeklyNarrative&&<div><div style={{fontSize:9,color:"#818CF8",fontFamily:"'IBM Plex Mono',monospace",letterSpacing:"0.15em",marginBottom:3}}>WEEKLY</div><div style={{fontSize:12,color:"#D1D5DB",lineHeight:1.6}}>{trade.weeklyNarrative}</div></div>}
              {trade.h1Narrative&&<div><div style={{fontSize:9,color:"#F59E0B",fontFamily:"'IBM Plex Mono',monospace",letterSpacing:"0.15em",marginBottom:3}}>H1</div><div style={{fontSize:12,color:"#D1D5DB",lineHeight:1.6}}>{trade.h1Narrative}</div></div>}
              {trade.confluences&&<div><div style={{fontSize:9,color:"#4B5563",fontFamily:"'IBM Plex Mono',monospace",letterSpacing:"0.15em",marginBottom:3}}>CONFLUENCES</div><div style={{fontSize:12,color:"#D1D5DB",lineHeight:1.6}}>{trade.confluences}</div></div>}
              {trade.entryReason&&<div><div style={{fontSize:9,color:"#4B5563",fontFamily:"'IBM Plex Mono',monospace",letterSpacing:"0.15em",marginBottom:3}}>ENTRY</div><div style={{fontSize:12,color:"#D1D5DB",lineHeight:1.6}}>{trade.entryReason}</div></div>}
            </div>
          )}

          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(220px,1fr))",gap:12,marginBottom:12}}>
            {trade.wentWell&&<div style={{padding:"10px 14px",background:"#064E3B22",borderRadius:8,border:"1px solid #064E3B"}}><div style={{fontSize:9,color:"#059669",fontFamily:"'IBM Plex Mono',monospace",letterSpacing:"0.15em",marginBottom:3}}>✓ WENT WELL</div><div style={{fontSize:12,color:"#D1FAE5"}}>{trade.wentWell}</div></div>}
            {trade.toImprove&&<div style={{padding:"10px 14px",background:"#7F1D1D22",borderRadius:8,border:"1px solid #7F1D1D"}}><div style={{fontSize:9,color:"#DC2626",fontFamily:"'IBM Plex Mono',monospace",letterSpacing:"0.15em",marginBottom:3}}>✗ TO IMPROVE</div><div style={{fontSize:12,color:"#FEE2E2"}}>{trade.toImprove}</div></div>}
            {trade.lesson&&<div style={{padding:"10px 14px",background:"#1E1B4B44",borderRadius:8,border:"1px solid #3730A3"}}><div style={{fontSize:9,color:"#818CF8",fontFamily:"'IBM Plex Mono',monospace",letterSpacing:"0.15em",marginBottom:3}}>★ LESSON</div><div style={{fontSize:12,color:"#C7D2FE"}}>{trade.lesson}</div></div>}
            {trade.mistakes&&<div style={{padding:"10px 14px",background:"#78350F22",borderRadius:8,border:"1px solid #92400E"}}><div style={{fontSize:9,color:"#F59E0B",fontFamily:"'IBM Plex Mono',monospace",letterSpacing:"0.15em",marginBottom:3}}>⚠ MISTAKES</div><div style={{fontSize:12,color:"#FDE68A"}}>{trade.mistakes}</div></div>}
          </div>

          {(trade.preEmotions?.length>0||trade.duringEmotions?.length>0||trade.postEmotions?.length>0)&&(
            <div style={{display:"flex",gap:16,marginBottom:12,flexWrap:"wrap"}}>
              {[["PRE",trade.preEmotions],["DURING",trade.duringEmotions],["POST",trade.postEmotions]].map(([lbl,ems])=>
                ems?.length>0?(<div key={lbl}><div style={{fontSize:9,color:"#4B5563",fontFamily:"'IBM Plex Mono',monospace",letterSpacing:"0.15em",marginBottom:4}}>{lbl}</div>
                <div style={{display:"flex",gap:4,flexWrap:"wrap"}}>{ems.map(em=><EmotionPill key={em} label={em} selected/>)}</div></div>):null
              )}
            </div>
          )}

          <div style={{display:"flex",gap:10,justifyContent:"flex-end",marginTop:8}}>
            <button type="button" onClick={e=>{e.stopPropagation();onEdit(trade);}} style={{
              padding:"7px 18px",borderRadius:8,cursor:"pointer",fontSize:12,
              fontFamily:"'IBM Plex Mono',monospace",border:"1px solid #374151",background:"#111827",color:"#D1D5DB",fontWeight:700}}>EDIT</button>
            <button type="button" onClick={e=>{e.stopPropagation();onDeleteRequest(trade.id);}} style={{
              padding:"7px 18px",borderRadius:8,cursor:"pointer",fontSize:12,
              fontFamily:"'IBM Plex Mono',monospace",border:"1px solid #EF4444",background:"#7F1D1D22",color:"#EF4444",fontWeight:700}}>DELETE</button>
          </div>
        </div>
      )}
    </div>
  );
};

// ─── ANALYTICS ────────────────────────────────────────────────────────────────
const Analytics = ({ trades, props, screenW }) => {
  if(!trades.length) return <div style={{textAlign:"center",padding:"60px 0",color:"#4B5563",fontFamily:"'IBM Plex Mono',monospace",fontSize:13}}>No trades yet.</div>;
  const wins=trades.filter(t=>t.result==="WIN").length, losses=trades.filter(t=>t.result==="LOSS").length, bes=trades.filter(t=>t.result==="BE").length;
  const wr=Math.round((wins/trades.length)*100);
  const avgR=trades.filter(t=>t.resultR).reduce((s,t,_,a)=>s+parseFloat(t.resultR||0)/a.length,0);
  const totalR=trades.reduce((s,t)=>s+parseFloat(t.resultR||0),0);
  const totalPct=trades.reduce((s,t)=>s+parseFloat(t.resultPct||0),0);
  const ruleBreaks=trades.filter(t=>t.rulesFollowed===false).length;
  let cum=0;
  const cumR=trades.map((t,i)=>{cum+=parseFloat(t.resultR||0);return{n:i+1,r:parseFloat(cum.toFixed(2))};});
  const propPerf=props.map(p=>{
    const pt=trades.filter(t=>t.propId===p.id);
    const pw=pt.filter(t=>t.result==="WIN").length;
    const ptR=pt.reduce((s,t)=>s+parseFloat(t.resultR||0),0);
    const ptPct=pt.reduce((s,t)=>s+parseFloat(t.resultPct||0),0);
    return{name:p.name,color:p.color,trades:pt.length,wr:pt.length>0?Math.round((pw/pt.length)*100):0,totalR:ptR.toFixed(1),totalPct:ptPct.toFixed(2)};
  }).filter(p=>p.trades>0);

  const h1StructStats = {};
  trades.forEach(t=>{ if(t.h1Structure){ if(!h1StructStats[t.h1Structure])h1StructStats[t.h1Structure]={wins:0,total:0}; h1StructStats[t.h1Structure].total++; if(t.result==="WIN")h1StructStats[t.h1Structure].wins++; } });
  const h1PosStats = {};
  trades.forEach(t=>{ if(t.h1Position){ if(!h1PosStats[t.h1Position])h1PosStats[t.h1Position]={wins:0,total:0}; h1PosStats[t.h1Position].total++; if(t.result==="WIN")h1PosStats[t.h1Position].wins++; } });

  const setupStats={};
  trades.forEach(t=>{ if(!t.setup)return; if(!setupStats[t.setup])setupStats[t.setup]={wins:0,total:0,totalR:0}; setupStats[t.setup].total++; if(t.result==="WIN")setupStats[t.setup].wins++; setupStats[t.setup].totalR+=parseFloat(t.resultR||0); });
  const setupData=Object.entries(setupStats).map(([s,v])=>({setup:s,wr:Math.round((v.wins/v.total)*100),avgR:(v.totalR/v.total).toFixed(2),count:v.total})).sort((a,b)=>parseFloat(b.avgR)-parseFloat(a.avgR));

  const emotionStats={};
  trades.forEach(t=>[...(t.preEmotions||[]),...(t.duringEmotions||[])].forEach(em=>{ if(!emotionStats[em])emotionStats[em]={wins:0,total:0}; emotionStats[em].total++; if(t.result==="WIN")emotionStats[em].wins++; }));
  const emotionData=Object.entries(emotionStats).filter(([,v])=>v.total>=2).map(([em,v])=>({emotion:em,wr:Math.round((v.wins/v.total)*100),count:v.total})).sort((a,b)=>b.wr-a.wr);

  const avgSetup=trades.filter(t=>t.setupRating).reduce((s,t,_,a)=>s+t.setupRating/a.length,0);
  const avgExec=trades.filter(t=>t.executionRating).reduce((s,t,_,a)=>s+t.executionRating/a.length,0);
  const avgMgmt=trades.filter(t=>t.managementRating).reduce((s,t,_,a)=>s+t.managementRating/a.length,0);

  const cS={background:"#0D1117",border:"1px solid #1F2937",borderRadius:12,padding:"16px 20px"};
  const kC=screenW<640?"repeat(3,1fr)":"repeat(6,1fr)";
  const mC=screenW<900?"1fr":"2fr 1fr";
  const bC=screenW<700?"1fr":screenW<1100?"1fr 1fr":"repeat(3,1fr)";

  return (
    <div>
      <h2 style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:28,letterSpacing:"0.1em",
        background:"linear-gradient(90deg,#818CF8,#C4B5FD)",WebkitBackgroundClip:"text",WebkitTextFillColor:"transparent",marginBottom:24}}>PERFORMANCE ANALYTICS</h2>

      <div style={{display:"grid",gridTemplateColumns:kC,gap:10,marginBottom:20}}>
        {[{l:"TRADES",v:trades.length,c:"#F9FAFB"},{l:"WIN RATE",v:`${wr}%`,c:wr>=50?"#10B981":"#EF4444"},
          {l:"AVG R",v:fmtR(avgR),c:avgR>=0?"#10B981":"#EF4444"},{l:"TOTAL R",v:fmtR(totalR),c:totalR>=0?"#10B981":"#EF4444"},
          {l:"TOTAL P&L",v:fmtPct(totalPct),c:totalPct>=0?"#10B981":"#EF4444"},{l:"RULE BREAKS",v:ruleBreaks,c:ruleBreaks>0?"#EF4444":"#10B981"}].map(k=>(
          <div key={k.l} style={{...cS,textAlign:"center"}}>
            <div style={{fontSize:9,color:"#4B5563",fontFamily:"'IBM Plex Mono',monospace",letterSpacing:"0.15em",marginBottom:6}}>{k.l}</div>
            <div style={{fontSize:20,fontFamily:"'Bebas Neue',sans-serif",letterSpacing:"0.1em",color:k.c}}>{k.v}</div>
          </div>
        ))}
      </div>

      <div style={{display:"grid",gridTemplateColumns:mC,gap:16,marginBottom:16}}>
        <div style={cS}>
          <div style={{fontSize:10,color:"#6B7280",fontFamily:"'IBM Plex Mono',monospace",letterSpacing:"0.15em",marginBottom:12}}>CUMULATIVE R</div>
          <ResponsiveContainer width="100%" height={200}>
            <AreaChart data={cumR}>
              <defs><linearGradient id="cumGrad" x1="0" y1="0" x2="0" y2="1">
                <stop offset="5%" stopColor="#F59E0B" stopOpacity={0.3}/>
                <stop offset="95%" stopColor="#F59E0B" stopOpacity={0}/>
              </linearGradient></defs>
              <CartesianGrid stroke="#1F2937" strokeDasharray="3 3"/>
              <XAxis dataKey="n" stroke="#374151" tick={{fill:"#6B7280",fontSize:10}}/>
              <YAxis stroke="#374151" tick={{fill:"#6B7280",fontSize:10}}/>
              <Tooltip contentStyle={{background:"#0D1117",border:"1px solid #374151",color:"#F9FAFB",fontSize:11,fontFamily:"'IBM Plex Mono',monospace"}}/>
              <Area type="monotone" dataKey="r" stroke="#F59E0B" fill="url(#cumGrad)" strokeWidth={2} dot={false}/>
            </AreaChart>
          </ResponsiveContainer>
        </div>
        <div style={cS}>
          <div style={{fontSize:10,color:"#6B7280",fontFamily:"'IBM Plex Mono',monospace",letterSpacing:"0.15em",marginBottom:12}}>AVG RATINGS</div>
          <ResponsiveContainer width="100%" height={200}>
            <RadarChart data={[{s:"Setup",v:avgSetup},{s:"Execution",v:avgExec},{s:"Management",v:avgMgmt}]}>
              <PolarGrid stroke="#1F2937"/>
              <PolarAngleAxis dataKey="s" tick={{fill:"#9CA3AF",fontSize:11}}/>
              <Radar dataKey="v" stroke="#818CF8" fill="#818CF8" fillOpacity={0.3}/>
            </RadarChart>
          </ResponsiveContainer>
        </div>
      </div>

      {(Object.keys(h1StructStats).length>0||Object.keys(h1PosStats).length>0)&&(
        <div style={{...cS,marginBottom:16}}>
          <div style={{fontSize:10,color:"#6B7280",fontFamily:"'IBM Plex Mono',monospace",letterSpacing:"0.15em",marginBottom:14}}>MULTI-TIMEFRAME ANALYSIS STATS</div>
          <div style={{display:"grid",gridTemplateColumns:screenW<700?"1fr":"1fr 1fr",gap:20}}>
            {Object.keys(h1StructStats).length>0&&(
              <div>
                <div style={{fontSize:9,color:"#F59E0B",fontFamily:"'IBM Plex Mono',monospace",letterSpacing:"0.15em",marginBottom:10}}>H1 STRUCTURE → WIN RATE</div>
                {Object.entries(h1StructStats).map(([k,v])=>{
                  const wr=Math.round((v.wins/v.total)*100);
                  const c=k.includes("✓")?"#10B981":k.includes("~")?"#F59E0B":"#EF4444";
                  return(<div key={k} style={{marginBottom:8}}>
                    <div style={{display:"flex",justifyContent:"space-between",marginBottom:2}}>
                      <span style={{fontSize:11,fontFamily:"'IBM Plex Mono',monospace",color:c}}>{k}</span>
                      <span style={{fontSize:11,fontFamily:"'IBM Plex Mono',monospace",color:wr>=50?"#10B981":"#EF4444"}}>{wr}% ({v.total})</span>
                    </div>
                    <div style={{background:"#1F2937",borderRadius:3,height:4}}><div style={{height:"100%",width:`${wr}%`,background:c,borderRadius:3}}/></div>
                  </div>);
                })}
              </div>
            )}
            {Object.keys(h1PosStats).length>0&&(
              <div>
                <div style={{fontSize:9,color:"#818CF8",fontFamily:"'IBM Plex Mono',monospace",letterSpacing:"0.15em",marginBottom:10}}>H1 POSITION → WIN RATE</div>
                {Object.entries(h1PosStats).map(([k,v])=>{
                  const wr=Math.round((v.wins/v.total)*100);
                  const c=k==="Discount"?"#10B981":k==="Premium"?"#EF4444":"#F59E0B";
                  return(<div key={k} style={{marginBottom:8}}>
                    <div style={{display:"flex",justifyContent:"space-between",marginBottom:2}}>
                      <span style={{fontSize:11,fontFamily:"'IBM Plex Mono',monospace",color:c}}>{k}</span>
                      <span style={{fontSize:11,fontFamily:"'IBM Plex Mono',monospace",color:wr>=50?"#10B981":"#EF4444"}}>{wr}% ({v.total})</span>
                    </div>
                    <div style={{background:"#1F2937",borderRadius:3,height:4}}><div style={{height:"100%",width:`${wr}%`,background:c,borderRadius:3}}/></div>
                  </div>);
                })}
              </div>
            )}
          </div>
        </div>
      )}

      <div style={{display:"grid",gridTemplateColumns:bC,gap:16,marginBottom:16}}>
        <div style={cS}>
          <div style={{fontSize:10,color:"#6B7280",fontFamily:"'IBM Plex Mono',monospace",letterSpacing:"0.15em",marginBottom:12}}>RESULTS</div>
          <ResponsiveContainer width="100%" height={160}>
            <PieChart>
              <Pie data={[{name:"WIN",value:wins||0},{name:"LOSS",value:losses||0},{name:"BE",value:bes||0}]} dataKey="value" cx="50%" cy="50%" outerRadius={55}>
                {["#10B981","#EF4444","#F59E0B"].map((c,i)=><Cell key={i} fill={c}/>)}
              </Pie>
              <Tooltip contentStyle={{background:"#0D1117",border:"1px solid #374151",color:"#F9FAFB",fontSize:11,fontFamily:"'IBM Plex Mono',monospace"}}/>
              <Legend iconType="circle" iconSize={8} formatter={v=><span style={{fontSize:11,fontFamily:"'IBM Plex Mono',monospace",color:"#9CA3AF"}}>{v}</span>}/>
            </PieChart>
          </ResponsiveContainer>
        </div>
        <div style={cS}>
          <div style={{fontSize:10,color:"#6B7280",fontFamily:"'IBM Plex Mono',monospace",letterSpacing:"0.15em",marginBottom:12}}>PER PROP</div>
          {propPerf.map(p=>(
            <div key={p.name} style={{marginBottom:10}}>
              <div style={{display:"flex",justifyContent:"space-between",marginBottom:3}}>
                <span style={{fontSize:12,fontFamily:"'IBM Plex Mono',monospace",fontWeight:700,color:p.color}}>{p.name}</span>
                <span style={{fontSize:10,fontFamily:"'IBM Plex Mono',monospace",color:"#6B7280"}}>{p.trades}T · {p.wr}% · {parseFloat(p.totalR)>=0?"+":""}{p.totalR}R · {fmtPct(p.totalPct)}</span>
              </div>
              <div style={{background:"#1F2937",borderRadius:3,height:4}}><div style={{height:"100%",width:`${p.wr}%`,background:p.color,borderRadius:3}}/></div>
            </div>
          ))}
        </div>
        {emotionData.length>0&&(
          <div style={cS}>
            <div style={{fontSize:10,color:"#6B7280",fontFamily:"'IBM Plex Mono',monospace",letterSpacing:"0.15em",marginBottom:12}}>EMOTION WIN RATE</div>
            {emotionData.slice(0,7).map(e=>{
              const neg=["Hesitant","Anxious","FOMO","Greedy","Revenge","Impulsive","Overconfident"].includes(e.emotion);
              return(<div key={e.emotion} style={{marginBottom:8}}>
                <div style={{display:"flex",justifyContent:"space-between",marginBottom:2}}>
                  <span style={{fontSize:10,fontFamily:"'IBM Plex Mono',monospace",color:neg?"#FCA5A5":"#6EE7B7"}}>{e.emotion}</span>
                  <span style={{fontSize:10,fontFamily:"'IBM Plex Mono',monospace",color:e.wr>=50?"#10B981":"#EF4444"}}>{e.wr}% ({e.count})</span>
                </div>
                <div style={{background:"#1F2937",borderRadius:3,height:3}}><div style={{height:"100%",width:`${e.wr}%`,background:e.wr>=50?"#10B981":"#EF4444",borderRadius:3}}/></div>
              </div>);
            })}
          </div>
        )}
      </div>

      {setupData.length>0&&(
        <div style={cS}>
          <div style={{fontSize:10,color:"#6B7280",fontFamily:"'IBM Plex Mono',monospace",letterSpacing:"0.15em",marginBottom:14}}>SETUP PERFORMANCE</div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(160px,1fr))",gap:10}}>
            {setupData.map(s=>(
              <div key={s.setup} style={{padding:"10px 14px",background:"#111827",borderRadius:8,border:"1px solid #1F2937"}}>
                <div style={{fontSize:11,fontFamily:"'IBM Plex Mono',monospace",color:"#D1D5DB",marginBottom:6,fontWeight:700}}>{s.setup}</div>
                <div style={{display:"flex",gap:12}}>
                  {[{l:"WR",v:`${s.wr}%`,c:s.wr>=50?"#10B981":"#EF4444"},{l:"AVG R",v:`${parseFloat(s.avgR)>=0?"+":""}${s.avgR}R`,c:parseFloat(s.avgR)>=0?"#10B981":"#EF4444"},{l:"N",v:s.count,c:"#9CA3AF"}].map(x=>(
                    <div key={x.l}><div style={{fontSize:9,color:"#4B5563",fontFamily:"'IBM Plex Mono',monospace",letterSpacing:"0.1em"}}>{x.l}</div>
                    <div style={{fontSize:14,fontFamily:"'Bebas Neue',sans-serif",color:x.c}}>{x.v}</div></div>
                  ))}
                </div>
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
};

// ─── SETTINGS ─────────────────────────────────────────────────────────────────
const Settings = ({ props, onSaveProps, trades, schedule, showToast, onImport, screenW, onLogout }) => {
  const [local, setLocal] = useState(()=>props.map(p=>({...p})));
  const [confirmRemove, setConfirmRemove] = useState(null);
  const importRef = useRef();
  useEffect(()=>setLocal(props.map(p=>({...p}))),[props]);

  const setF = (id,k,v) => setLocal(lp=>lp.map(p=>p.id===id?{...p,[k]:v}:p));
  const addProp = () => {
    const newProp = { id:mkId(), name:"NEW FIRM", color:"#94A3B8", accountSize:50000,
      maxDD:2500, weeklyTarget:2, riskPct:0.5, phase:"Challenge", note:"" };
    setLocal(lp=>[...lp,newProp]);
  };
  const removeProp = (id) => { setLocal(lp=>lp.filter(p=>p.id!==id)); setConfirmRemove(null); };

  const handleImport = (e) => {
    const file=e.target.files[0]; if(!file) return;
    const reader=new FileReader();
    reader.onload=(ev)=>{ try {
      const data=JSON.parse(ev.target.result);
      if(data.trades&&data.props) { onImport(data); showToast("Imported ✓","success"); }
      else showToast("Invalid file","error");
    } catch { showToast("Could not read file","error"); } };
    reader.readAsText(file); e.target.value="";
  };

  const iS={background:"#0D1117",border:"1px solid #1F2937",borderRadius:8,color:"#F9FAFB",
    padding:"7px 10px",fontSize:12,fontFamily:"'IBM Plex Mono',monospace",outline:"none",width:"100%"};
  const lS={fontSize:9,color:"#4B5563",fontFamily:"'IBM Plex Mono',monospace",letterSpacing:"0.15em",display:"block",marginBottom:3};
  const gC=screenW<640?"repeat(2,1fr)":screenW<960?"repeat(3,1fr)":"2fr 1fr 1fr 1fr 1fr 1fr 80px";
  const btnS=(c)=>({padding:"10px 18px",borderRadius:9,cursor:"pointer",fontSize:12,fontFamily:"'IBM Plex Mono',monospace",
    fontWeight:700,letterSpacing:"0.1em",border:`1px solid ${c}`,background:`${c}18`,color:c,transition:"all 0.15s"});

  return (
    <div>
      {confirmRemove&&<ConfirmModal message={`Remove "${local.find(p=>p.id===confirmRemove)?.name}"? Their trade history will be kept.`}
        onConfirm={()=>removeProp(confirmRemove)} onCancel={()=>setConfirmRemove(null)}/>}

      <h2 style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:28,letterSpacing:"0.1em",
        background:"linear-gradient(90deg,#6EE7B7,#34D399)",WebkitBackgroundClip:"text",WebkitTextFillColor:"transparent",marginBottom:24}}>SETTINGS & DATA</h2>

      <div style={{background:"#0D1117",border:"1px solid #1F2937",borderRadius:12,padding:20,marginBottom:20}}>
        <div style={{fontSize:10,color:"#6B7280",fontFamily:"'IBM Plex Mono',monospace",letterSpacing:"0.2em",marginBottom:14}}>DATA MANAGEMENT</div>
        <div style={{display:"flex",gap:10,flexWrap:"wrap",marginBottom:12}}>
          <button type="button" onClick={()=>{doExportJSON(trades,props,schedule);showToast("JSON exported ✓","success");}} style={btnS("#10B981")}>⬇ EXPORT JSON</button>
          <button type="button" onClick={()=>{doExportXLSX(trades,props);showToast("XLSX exported ✓","success");}} style={btnS("#10B981")}>⬇ EXPORT XLSX</button>
          <button type="button" onClick={()=>importRef.current?.click()} style={btnS("#818CF8")}>⬆ IMPORT JSON</button>
          <input ref={importRef} type="file" accept=".json" onChange={handleImport} style={{display:"none"}}/>
          <button type="button" onClick={onLogout} style={btnS("#EF4444")}>SAIR (LOGOUT)</button>
        </div>
        <div style={{fontSize:11,color:"#4B5563",fontFamily:"'IBM Plex Mono',monospace",lineHeight:1.8}}>
          {trades.length} trades · Dados sincronizados com a cloud · <span style={{color:"#F59E0B"}}>Export regularmente como backup</span>
        </div>
      </div>

      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
        <div style={{fontSize:10,color:"#6B7280",fontFamily:"'IBM Plex Mono',monospace",letterSpacing:"0.2em"}}>PROP FIRM CONFIGURATION</div>
        <button type="button" onClick={addProp} style={btnS("#F59E0B")}>+ ADD FIRM</button>
      </div>

      {local.map(p=>(
        <div key={p.id} style={{background:"#0D1117",border:`1px solid ${p.color}44`,borderRadius:12,
          padding:"18px 20px",marginBottom:14,boxShadow:`0 0 12px ${p.color}11`}}>
          <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:14}}>
            <div style={{width:10,height:10,borderRadius:"50%",background:p.color}}/>
            <span style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:18,letterSpacing:"0.1em",color:"#F9FAFB"}}>{p.name}</span>
            <button type="button" onClick={()=>setConfirmRemove(p.id)} style={{marginLeft:"auto",background:"none",
              border:"1px solid #374151",borderRadius:6,padding:"4px 10px",cursor:"pointer",
              fontSize:10,color:"#6B7280",fontFamily:"'IBM Plex Mono',monospace"}}>REMOVE</button>
          </div>
          <div style={{display:"grid",gridTemplateColumns:gC,gap:12}}>
            <div><label style={lS}>FIRM NAME</label><input value={p.name} onChange={e=>setF(p.id,"name",e.target.value)} style={iS}/></div>
            <div><label style={lS}>ACCOUNT ($)</label><input type="number" value={p.accountSize} onChange={e=>setF(p.id,"accountSize",parseFloat(e.target.value)||0)} style={iS}/></div>
            <div><label style={lS}>MAX DD ($)</label><input type="number" value={p.maxDD} onChange={e=>setF(p.id,"maxDD",parseFloat(e.target.value)||0)} style={iS}/></div>
            <div><label style={lS}>WEEKLY TARGET %</label><input type="number" step="0.1" value={p.weeklyTarget} onChange={e=>setF(p.id,"weeklyTarget",parseFloat(e.target.value)||2)} style={iS}/></div>
            <div><label style={lS}>RISK/TRADE %</label><input type="number" step="0.1" value={p.riskPct} onChange={e=>setF(p.id,"riskPct",parseFloat(e.target.value)||0.5)} style={iS}/></div>
            <div><label style={lS}>PHASE</label>
              <select value={p.phase} onChange={e=>setF(p.id,"phase",e.target.value)} style={iS}>
                {PHASES.map(ph=><option key={ph}>{ph}</option>)}
              </select></div>
            <div><label style={lS}>COLOR</label><input type="color" value={p.color} onChange={e=>setF(p.id,"color",e.target.value)} style={{...iS,height:34,padding:2,cursor:"pointer"}}/></div>
          </div>
          <div style={{marginTop:10}}>
            <label style={lS}>NOTES</label>
            <input value={p.note||""} onChange={e=>setF(p.id,"note",e.target.value)} style={iS} placeholder="Broker, rules, payout info..."/>
          </div>
        </div>
      ))}
      <button type="button" onClick={()=>{onSaveProps(local);showToast("Settings saved ✓","success");}} style={{
        padding:"13px 32px",borderRadius:10,cursor:"pointer",
        background:"linear-gradient(135deg,#064E3B,#10B981)",border:"none",
        color:"#F9FAFB",fontSize:14,fontFamily:"'Bebas Neue',sans-serif",letterSpacing:"0.2em"}}>
        SAVE ALL SETTINGS
      </button>
    </div>
  );
};

// ─── MAIN APP ─────────────────────────────────────────────────────────────────
export default function App() {
  const screenW = useWindowSize();
  const [session, setSession] = useState(null);
  const [loadingAuth, setLoadingAuth] = useState(true);
  const [props, setPropsState] = useState(DEFAULT_PROPS);
  const [schedule, setScheduleState] = useState(DEFAULT_SCHEDULE);
  const [trades, setTradesState] = useState([]);
  const [view, setView] = useState("dashboard");
  const [editTrade, setEditTrade] = useState(null);
  const [filterProp, setFilterProp] = useState("all");
  const [filterResult, setFilterResult] = useState("all");
  const [confirmDelete, setConfirmDelete] = useState(null);
  const [navOpen, setNavOpen] = useState(false);
  const [toast, showToast, hideToast] = useToast();
  const [cloudReady, setCloudReady] = useState(false);

  // Auth listener
  useEffect(() => {
    supabase.auth.getSession().then(({ data: { session } }) => {
      setSession(session);
      setLoadingAuth(false);
    });
    const { data: { subscription } } = supabase.auth.onAuthStateChange((_event, session) => {
      setSession(session);
    });
    return () => subscription.unsubscribe();
  }, []);

  // Load from cloud when logged in
  useEffect(() => {
    if (!session?.user) return;
    const load = async () => {
      const { data, error } = await supabase
        .from("user_data")
        .select("*")
        .eq("user_id", session.user.id)
        .maybeSingle();
      if (error) {
        console.error(error);
        showToast("Erro a carregar dados", "error");
        return;
      }
      if (data) {
        if (data.props) setPropsState(data.props);
        if (data.schedule) setScheduleState(data.schedule);
        if (data.trades) setTradesState(data.trades);
      } else {
        // primeira vez → criar linha
        await supabase.from("user_data").insert({
          user_id: session.user.id,
          props: DEFAULT_PROPS,
          schedule: DEFAULT_SCHEDULE,
          trades: []
        });
      }
      setCloudReady(true);
    };
    load();
  }, [session?.user?.id]);

  // Save to cloud (debounce simples)
  const saveToCloud = useCallback(async (nextProps, nextSchedule, nextTrades) => {
    if (!session?.user) return;
    try {
      await supabase.from("user_data").upsert({
        user_id: session.user.id,
        props: nextProps,
        schedule: nextSchedule,
        trades: nextTrades,
        updated_at: new Date().toISOString()
      }, { onConflict: "user_id" });
    } catch (e) {
      console.error(e);
      showToast("Erro a gravar na cloud", "error");
    }
  }, [session?.user]);

  const setProps = useCallback((updater) => {
    setPropsState(prev => {
      const next = typeof updater === "function" ? updater(prev) : updater;
      saveToCloud(next, schedule, trades);
      return next;
    });
  }, [schedule, trades, saveToCloud]);

  const setSchedule = useCallback((updater) => {
    setScheduleState(prev => {
      const next = typeof updater === "function" ? updater(prev) : updater;
      saveToCloud(props, next, trades);
      return next;
    });
  }, [props, trades, saveToCloud]);

  const setTrades = useCallback((updater) => {
    setTradesState(prev => {
      const next = typeof updater === "function" ? updater(prev) : updater;
      saveToCloud(props, schedule, next);
      return next;
    });
  }, [props, schedule, saveToCloud]);

  const todayDay = getTodayDayKey();
  const weekDates = getWeekDates();
  const nm = useCallback((d)=>normDay(schedule,d),[schedule]);
  const todayPropIds = nm(todayDay);
  const todayProp = props.find(p=>p.id===todayPropIds[0]);
  const weekPnlPct = trades.filter(t=>todayPropIds.includes(t.propId)&&weekDates.includes(t.date)).reduce((s,t)=>s+parseFloat(t.resultPct||0),0);

  const saveTrade = useCallback((trade)=>{
    setTrades(prev=>{ const idx=prev.findIndex(t=>t.id===trade.id); if(idx>=0){const n=[...prev];n[idx]=trade;return n;} return [trade,...prev]; });
    setEditTrade(null); setView("log");
  },[setTrades]);

  const handleDeleteConfirm = useCallback(()=>{
    setTrades(prev=>prev.filter(t=>t.id!==confirmDelete));
    setConfirmDelete(null); showToast("Trade deleted","error");
  },[confirmDelete,setTrades,showToast]);

  const handleImport = useCallback((data)=>{ setProps(data.props); setSchedule(data.schedule||DEFAULT_SCHEDULE); setTrades(data.trades); },[setProps,setSchedule,setTrades]);
  const changeSchedule = useCallback((day,propId)=>{ setSchedule(s=>{ const cur=nm(day); const next=cur.includes(propId)?cur.filter(x=>x!==propId):[...cur,propId]; return{...s,[day]:next}; }); },[setSchedule,nm]);

  const handleLogout = async () => {
    await supabase.auth.signOut();
    setSession(null);
    setPropsState(DEFAULT_PROPS);
    setScheduleState(DEFAULT_SCHEDULE);
    setTradesState([]);
    setCloudReady(false);
  };

  const filteredTrades = trades.filter(t=>(filterProp==="all"||t.propId===filterProp)&&(filterResult==="all"||t.result===filterResult));
  const NAV=[{id:"dashboard",l:"DASHBOARD"},{id:"schedule",l:"SCHEDULE"},{id:"new",l:"＋ NEW"},{id:"log",l:"TRADE LOG"},{id:"analytics",l:"ANALYTICS"},{id:"settings",l:"SETTINGS"}];
  const isMob=screenW<768, isMid=screenW<1100;

  if (loadingAuth) {
    return (
      <div style={{ minHeight:"100vh", background:"#060A0F", display:"flex", alignItems:"center", justifyContent:"center",
        color:"#6B7280", fontFamily:"'IBM Plex Mono',monospace", fontSize:14 }}>
        A carregar...
      </div>
    );
  }

  if (!session) {
    return (
      <>
        <style>{`
          @import url('https://fonts.googleapis.com/css2?family=Bebas+Neue&family=IBM+Plex+Mono:wght@400;600;700&display=swap');
          *{box-sizing:border-box;margin:0;padding:0;}body{background:#060A0F;}
        `}</style>
        {toast&&<Toast key={toast.id} message={toast.msg} type={toast.type} onDone={hideToast}/>}
        <AuthScreen onLogin={setSession} showToast={showToast} />
      </>
    );
  }

  return (
    <>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Bebas+Neue&family=IBM+Plex+Mono:wght@400;600;700&display=swap');
        *{box-sizing:border-box;margin:0;padding:0;}body{background:#060A0F;}
        ::-webkit-scrollbar{width:5px;height:5px;}::-webkit-scrollbar-track{background:#0D1117;}::-webkit-scrollbar-thumb{background:#1F2937;border-radius:3px;}
        input:focus,textarea:focus,select:focus{border-color:#374151!important;outline:none;}
        input[type=date]::-webkit-calendar-picker-indicator{filter:invert(0.5);}
        select option{background:#0D1117;color:#F9FAFB;}
        @keyframes slideIn{from{transform:translateY(16px);opacity:0;}to{transform:translateY(0);opacity:1;}}
      `}</style>

      {toast&&<Toast key={toast.id} message={toast.msg} type={toast.type} onDone={hideToast}/>}
      {confirmDelete&&<ConfirmModal message="Delete this trade? Cannot be undone." onConfirm={handleDeleteConfirm} onCancel={()=>setConfirmDelete(null)}/>}

      <div style={{minHeight:"100vh",background:"#060A0F",color:"#F9FAFB",fontFamily:"'IBM Plex Mono',monospace",display:"flex",flexDirection:"column"}}>

        {/* HEADER */}
        <div style={{borderBottom:"1px solid #111827",padding:`0 ${isMob?14:28}px`,display:"flex",alignItems:"center",gap:12,background:"#060A0F",position:"sticky",top:0,zIndex:100,minHeight:54}}>
          <div style={{paddingRight:isMob?12:20,borderRight:"1px solid #111827",marginRight:isMob?8:16,flexShrink:0}}>
            <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:isMob?17:21,letterSpacing:"0.15em",background:"linear-gradient(135deg,#F59E0B,#FDE68A)",WebkitBackgroundClip:"text",WebkitTextFillColor:"transparent"}}>PROPDESK</div>
          </div>
          {todayProp&&!isMob&&(
            <div style={{padding:"5px 10px",background:`${todayProp.color}15`,border:`1px solid ${todayProp.color}44`,borderRadius:7,flexShrink:0}}>
              <div style={{fontSize:8,color:"#6B7280",letterSpacing:"0.2em",marginBottom:1}}>TODAY</div>
              <div style={{display:"flex",gap:6}}>
                {todayPropIds.map(id=>{ const p=props.find(x=>x.id===id); return p?<span key={id} style={{fontSize:13,fontWeight:700,color:p.color,fontFamily:"'Bebas Neue',sans-serif",letterSpacing:"0.1em"}}>{p.name}</span>:null; })}
              </div>
            </div>
          )}
          {todayProp&&!isMob&&(
            <div style={{padding:"5px 10px",background:"#0D1117",border:"1px solid #1F2937",borderRadius:7,flexShrink:0}}>
              <div style={{fontSize:8,color:"#6B7280",letterSpacing:"0.2em",marginBottom:1}}>WEEK P&L</div>
              <div style={{fontSize:13,fontWeight:700,fontFamily:"'Bebas Neue',sans-serif",letterSpacing:"0.1em",color:weekPnlPct>=0?"#10B981":"#EF4444"}}>{fmtPct(weekPnlPct)} / {todayProp.weeklyTarget}%</div>
            </div>
          )}
          <div style={{flex:1}}/>
          {!isMob&&<nav style={{display:"flex",gap:0}}>
            {NAV.map(n=>(
              <button key={n.id} type="button" onClick={()=>{setView(n.id);if(n.id!=="new")setEditTrade(null);}} style={{
                padding:"16px 14px",background:"none",border:"none",cursor:"pointer",fontSize:11,
                letterSpacing:"0.1em",fontFamily:"'IBM Plex Mono',monospace",fontWeight:700,whiteSpace:"nowrap",
                color:view===n.id?"#F9FAFB":"#4B5563",
                borderBottom:view===n.id?"2px solid #F59E0B":"2px solid transparent"}}>{n.l}</button>
            ))}
          </nav>}
          {isMob&&<button type="button" onClick={()=>setNavOpen(!navOpen)} style={{background:"none",border:"1px solid #1F2937",borderRadius:8,color:"#9CA3AF",padding:"6px 11px",cursor:"pointer",fontSize:16}}>{navOpen?"✕":"☰"}</button>}
        </div>

        {isMob&&navOpen&&(
          <div style={{background:"#0D1117",borderBottom:"1px solid #1F2937",position:"sticky",top:54,zIndex:99}}>
            {NAV.map(n=>(
              <button key={n.id} type="button" onClick={()=>{setView(n.id);if(n.id!=="new")setEditTrade(null);setNavOpen(false);}} style={{
                display:"block",width:"100%",padding:"12px 20px",border:"none",borderBottom:"1px solid #111827",
                cursor:"pointer",textAlign:"left",fontSize:12,letterSpacing:"0.12em",
                fontFamily:"'IBM Plex Mono',monospace",fontWeight:700,
                color:view===n.id?"#F59E0B":"#9CA3AF",background:view===n.id?"#111827":"transparent"}}>{n.l}</button>
            ))}
          </div>
        )}

        <div style={{flex:1,padding:isMob?"14px":isMid?"20px":"28px",maxWidth:1440,width:"100%",margin:"0 auto"}}>
          {!cloudReady && (
            <div style={{ textAlign:"center", padding:"40px 0", color:"#6B7280", fontFamily:"'IBM Plex Mono',monospace", fontSize:13 }}>
              A sincronizar com a cloud...
            </div>
          )}

          {cloudReady && view==="dashboard"&&(
            <div>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"baseline",marginBottom:20,flexWrap:"wrap",gap:8}}>
                <h2 style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:26,letterSpacing:"0.1em",background:"linear-gradient(90deg,#F59E0B,#FDE68A)",WebkitBackgroundClip:"text",WebkitTextFillColor:"transparent"}}>DASHBOARD</h2>
                <span style={{fontSize:10,color:"#4B5563",fontFamily:"'IBM Plex Mono',monospace",letterSpacing:"0.1em"}}>{new Date().toLocaleDateString("en-US",{weekday:"long",month:"long",day:"numeric"}).toUpperCase()}</span>
              </div>
              <div style={{display:"grid",gridTemplateColumns:isMob?"1fr 1fr":isMid?"1fr 1fr":"repeat(4,1fr)",gap:12,marginBottom:20}}>
                {props.filter(p=>p.phase!=="Disabled").map(p=><PropCard key={p.id} prop={p} trades={trades} isToday={nm(todayDay).includes(p.id)}/>)}
              </div>
              <div style={{background:"#0D1117",border:"1px solid #1F2937",borderRadius:12,padding:18,marginBottom:18}}>
                <div style={{fontSize:10,color:"#6B7280",fontFamily:"'IBM Plex Mono',monospace",letterSpacing:"0.15em",marginBottom:14}}>THIS WEEK</div>
                <div style={{display:"grid",gridTemplateColumns:isMob?"1fr 1fr":"repeat(5,1fr)",gap:10}}>
                  {DAYS.map((day,i)=>{
                    const pids=nm(day),isToday=day===todayDay;
                    const dayProps=props.filter(p=>pids.includes(p.id));
                    const dt=trades.filter(t=>t.date===weekDates[i]&&pids.includes(t.propId));
                    const dR=dt.reduce((s,t)=>s+parseFloat(t.resultR||0),0);
                    const dPct=dt.reduce((s,t)=>s+parseFloat(t.resultPct||0),0);
                    return(
                      <div key={day} style={{padding:"12px",borderRadius:10,
                        border:`1px solid ${isToday?(dayProps[0]?.color||"#F59E0B"):"#1F2937"}`,
                        background:isToday?`${dayProps[0]?.color||"#F59E0B"}0A`:"#080E17"}}>
                        <div style={{display:"flex",justifyContent:"space-between",marginBottom:5}}>
                          <span style={{fontSize:11,fontFamily:"'Bebas Neue',sans-serif",letterSpacing:"0.1em",color:isToday?"#F9FAFB":"#6B7280"}}>{DAY_LABELS[i].slice(0,3).toUpperCase()}</span>
                          {isToday&&<span style={{fontSize:8,color:"#F59E0B",letterSpacing:"0.15em"}}>NOW</span>}
                        </div>
                        {dayProps.length>0?(<>
                          <div style={{display:"flex",flexWrap:"wrap",gap:3,marginBottom:3}}>
                            {dayProps.map(p=><span key={p.id} style={{fontSize:11,fontFamily:"'Bebas Neue',sans-serif",letterSpacing:"0.08em",color:p.color}}>{p.name}</span>)}
                          </div>
                          <div style={{fontSize:10,color:"#6B7280",fontFamily:"'IBM Plex Mono',monospace"}}>{dt.length>0?`${dt.length}T · ${fmtR(dR)} · ${fmtPct(dPct)}`:"No trades"}</div>
                        </>):<div style={{fontSize:12,color:"#374151",fontFamily:"'IBM Plex Mono',monospace"}}>—</div>}
                      </div>
                    );
                  })}
                </div>
              </div>
              {trades.length>0&&(
                <div style={{background:"#0D1117",border:"1px solid #1F2937",borderRadius:12,padding:18}}>
                  <div style={{display:"flex",justifyContent:"space-between",marginBottom:14}}>
                    <div style={{fontSize:10,color:"#6B7280",fontFamily:"'IBM Plex Mono',monospace",letterSpacing:"0.15em"}}>RECENT TRADES</div>
                    <button type="button" onClick={()=>setView("log")} style={{fontSize:10,color:"#F59E0B",background:"none",border:"none",cursor:"pointer",fontFamily:"'IBM Plex Mono',monospace"}}>VIEW ALL →</button>
                  </div>
                  {trades.slice(0,5).map(t=><TradeRow key={t.id} trade={t} props={props} onEdit={tr=>{setEditTrade(tr);setView("new");}} onDeleteRequest={id=>setConfirmDelete(id)}/>)}
                </div>
              )}
              {!trades.length&&<div style={{textAlign:"center",padding:"40px 0"}}>
                <div style={{fontSize:13,color:"#374151",fontFamily:"'IBM Plex Mono',monospace",marginBottom:16}}>No trades logged yet.</div>
                <button type="button" onClick={()=>setView("new")} style={{padding:"12px 28px",borderRadius:10,cursor:"pointer",background:"linear-gradient(135deg,#92400E,#F59E0B)",border:"none",color:"#0F0F14",fontSize:13,fontFamily:"'Bebas Neue',sans-serif",letterSpacing:"0.2em"}}>LOG YOUR FIRST TRADE</button>
              </div>}
            </div>
          )}

          {cloudReady && view==="schedule"&&(
            <div>
              <h2 style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:26,letterSpacing:"0.1em",background:"linear-gradient(90deg,#F59E0B,#FDE68A)",WebkitBackgroundClip:"text",WebkitTextFillColor:"transparent",marginBottom:16}}>WEEKLY SCHEDULE</h2>
              <p style={{fontSize:12,color:"#6B7280",fontFamily:"'IBM Plex Mono',monospace",marginBottom:22,lineHeight:1.7}}>Select one or more props per day. Click to toggle. Adjust based on P&L priorities each week.</p>
              <div style={{display:"grid",gridTemplateColumns:isMob?"1fr 1fr":isMid?"repeat(3,1fr)":"repeat(5,1fr)",gap:14,marginBottom:28}}>
                {DAYS.map((day,i)=>{
                  const isToday=day===todayDay;
                  return(
                    <div key={day}>
                      <div style={{fontSize:11,fontFamily:"'Bebas Neue',sans-serif",letterSpacing:"0.2em",marginBottom:10,color:isToday?"#F59E0B":"#9CA3AF"}}>{DAY_LABELS[i].toUpperCase()} {isToday&&"◆"}</div>
                      <div style={{display:"flex",flexDirection:"column",gap:7}}>
                        {props.filter(p=>p.phase!=="Disabled").map(p=>{
                          const sel=nm(day).includes(p.id);
                          return(<button key={p.id} type="button" onClick={()=>changeSchedule(day,p.id)} style={{
                            padding:"9px 12px",borderRadius:9,cursor:"pointer",textAlign:"left",transition:"all 0.15s",
                            border:`1px solid ${sel?p.color:"#1F2937"}`,background:sel?`${p.color}18`:"transparent",
                            color:sel?p.color:"#4B5563",fontSize:12,fontFamily:"'IBM Plex Mono',monospace",fontWeight:700}}>
                            {sel?"◆ ":"○ "}{p.name}
                          </button>);
                        })}
                        <button type="button" onClick={()=>setSchedule(s=>({...s,[day]:[]}))} style={{padding:"5px 12px",borderRadius:9,cursor:"pointer",textAlign:"left",border:"1px solid #1F2937",background:"transparent",color:"#374151",fontSize:10,fontFamily:"'IBM Plex Mono',monospace"}}>✕ clear</button>
                      </div>
                    </div>
                  );
                })}
              </div>
              <div style={{background:"#0D1117",border:"1px solid #1F2937",borderRadius:12,padding:18}}>
                <div style={{fontSize:10,color:"#6B7280",fontFamily:"'IBM Plex Mono',monospace",letterSpacing:"0.15em",marginBottom:12}}>SUMMARY</div>
                <div style={{display:"flex",gap:20,flexWrap:"wrap"}}>
                  {props.map(p=>{ const days=DAYS.filter(d=>nm(d).includes(p.id)); return days.length>0?(
                    <div key={p.id} style={{display:"flex",alignItems:"center",gap:8}}>
                      <div style={{width:7,height:7,borderRadius:"50%",background:p.color}}/>
                      <span style={{fontFamily:"'IBM Plex Mono',monospace",fontSize:12,color:p.color,fontWeight:700}}>{p.name}</span>
                      <span style={{fontFamily:"'IBM Plex Mono',monospace",fontSize:11,color:"#6B7280"}}>→ {days.join(", ")}</span>
                    </div>
                  ):null; })}
                </div>
              </div>
            </div>
          )}

          {cloudReady && view==="new"&&<TradeForm key={editTrade?.id||"new"} props={props} schedule={schedule} onSave={saveTrade} editTrade={editTrade} screenW={screenW} showToast={showToast} onCancel={editTrade?()=>{setEditTrade(null);setView("log");}:null}/>}

          {cloudReady && view==="log"&&(
            <div>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20,flexWrap:"wrap",gap:10}}>
                <h2 style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:26,letterSpacing:"0.1em",background:"linear-gradient(90deg,#F59E0B,#FDE68A)",WebkitBackgroundClip:"text",WebkitTextFillColor:"transparent"}}>TRADE LOG <span style={{fontSize:14,color:"#374151",marginLeft:8}}>({filteredTrades.length})</span></h2>
                <button type="button" onClick={()=>{setEditTrade(null);setView("new");}} style={{padding:"8px 18px",borderRadius:8,cursor:"pointer",background:"linear-gradient(135deg,#92400E,#F59E0B)",border:"none",color:"#0F0F14",fontSize:12,fontFamily:"'Bebas Neue',sans-serif",letterSpacing:"0.15em"}}>+ NEW</button>
              </div>
              <div style={{display:"flex",gap:8,marginBottom:14,flexWrap:"wrap"}}>
                {["all",...props.map(p=>p.id)].map(id=>{const p=props.find(x=>x.id===id),isAll=id==="all"; return(
                  <button key={id} type="button" onClick={()=>setFilterProp(id)} style={{padding:"5px 12px",borderRadius:20,cursor:"pointer",fontSize:11,fontFamily:"'IBM Plex Mono',monospace",transition:"all 0.15s",
                    border:`1px solid ${filterProp===id?(isAll?"#F59E0B":p?.color||"#F59E0B"):"#1F2937"}`,
                    background:filterProp===id?(isAll?"#78350F":`${p?.color}20`||"transparent"):"transparent",
                    color:filterProp===id?(isAll?"#FDE68A":p?.color||"#F9FAFB"):"#6B7280"}}>{isAll?"ALL":p?.name}</button>
                );})}
                <div style={{width:"1px",background:"#1F2937",margin:"0 4px"}}/>
                {["all","WIN","LOSS","BE"].map(r=>(
                  <button key={r} type="button" onClick={()=>setFilterResult(r)} style={{padding:"5px 12px",borderRadius:20,cursor:"pointer",fontSize:11,fontFamily:"'IBM Plex Mono',monospace",transition:"all 0.15s",
                    border:`1px solid ${filterResult===r?(r==="WIN"?"#10B981":r==="LOSS"?"#EF4444":r==="BE"?"#F59E0B":"#818CF8"):"#1F2937"}`,
                    background:filterResult===r?(r==="WIN"?"#064E3B":r==="LOSS"?"#7F1D1D":r==="BE"?"#78350F":"#1E1B4B"):"transparent",
                    color:filterResult===r?(r==="WIN"?"#6EE7B7":r==="LOSS"?"#FCA5A5":r==="BE"?"#FDE68A":"#A5B4FC"):"#6B7280"}}>{r==="all"?"ALL":r}</button>
                ))}
              </div>
              {!isMob&&filteredTrades.length>0&&(
                <div style={{display:"grid",gridTemplateColumns:"88px 75px 55px 65px 90px 58px 62px 50px 55px 24px",gap:6,padding:"6px 14px",marginBottom:4,borderBottom:"1px solid #111827"}}>
                  {["DATE","PROP","MKT","DIR","SETUP","RESULT","R","GRADE","★",""].map(h=><span key={h} style={{fontSize:9,color:"#374151",fontFamily:"'IBM Plex Mono',monospace",letterSpacing:"0.15em"}}>{h}</span>)}
                </div>
              )}
              {filteredTrades.length===0?<div style={{textAlign:"center",padding:"60px 0",color:"#374151",fontFamily:"'IBM Plex Mono',monospace",fontSize:13}}>No trades found. <button type="button" onClick={()=>setView("new")} style={{color:"#F59E0B",background:"none",border:"none",cursor:"pointer",fontFamily:"'IBM Plex Mono',monospace",fontSize:13}}>Log a trade →</button></div>
                :filteredTrades.map(t=><TradeRow key={t.id} trade={t} props={props} onEdit={tr=>{setEditTrade(tr);setView("new");}} onDeleteRequest={id=>setConfirmDelete(id)}/>)}
            </div>
          )}

          {cloudReady && view==="analytics"&&<Analytics trades={trades} props={props} screenW={screenW}/>}
          {cloudReady && view==="settings"&&<Settings props={props} onSaveProps={setProps} trades={trades} schedule={schedule} showToast={showToast} onImport={handleImport} screenW={screenW} onLogout={handleLogout}/>}
        </div>

        <div style={{borderTop:"1px solid #0D1117",padding:"8px 20px",display:"flex",justifyContent:"space-between",flexWrap:"wrap",gap:6}}>
          <span style={{fontSize:9,color:"#1F2937",fontFamily:"'IBM Plex Mono',monospace",letterSpacing:"0.2em"}}>PROPDESK v2.0 · CLOUD</span>
          <span style={{fontSize:9,color:"#1F2937",fontFamily:"'IBM Plex Mono',monospace",letterSpacing:"0.1em"}}>{trades.length} TRADES · {props.filter(p=>p.phase!=="Disabled").length} PROPS ACTIVE</span>
        </div>
      </div>
    </>
  );
}