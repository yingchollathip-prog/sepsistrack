import { useState, useEffect, useCallback, useRef } from "react";
import * as XLSX from "xlsx";

// ─── Constants ────────────────────────────────────────────────────────────────
const DEADLINE_SEC = 60 * 60;
const STORAGE_KEY  = "sepsis_cases_v4";
const AUDIT_KEY    = "sepsis_audit_v4";

// ─── Google Sheets Backend Config ────────────────────────────────────────────
const SHEETS_URL = "https://script.google.com/macros/s/AKfycbx2pTp_HDxIAoEC-fm7QrrwVnuwfJu9A08eenPudasFd3vnG_nwTqxHOn65BUt8INSlfQ/exec";
const API_KEY    = "SepsisTrack-ER-2026";

// Helper: GET request to Sheets backend
const sheetsGet = async (action) => {
  const url = `${SHEETS_URL}?action=${action}&apiKey=${API_KEY}`;
  const res  = await fetch(url);
  const json = await res.json();
  if (!json.ok) throw new Error(json.error || "Sheets GET failed");
  return json.data;
};

// Helper: POST request to Sheets backend
const sheetsPost = async (body) => {
  const res  = await fetch(SHEETS_URL, {
    method:  "POST",
    headers: { "Content-Type": "application/json" },
    body:    JSON.stringify({ ...body, apiKey: API_KEY }),
  });
  const json = await res.json();
  if (!json.ok) throw new Error(json.error || "Sheets POST failed");
  return json.data;
};

const SOURCES = ["Pneumonia","UTI / Urinary","Abdominal / GI","Skin / Soft Tissue",
  "Bacteremia / Unknown","CNS / Meningitis","Endocarditis","Bone / Joint","Other"];
const ANTIBIOTICS = ["Piperacillin-Tazobactam (Tazocin)","Meropenem","Ceftriaxone",
  "Cefazolin","Vancomycin","Metronidazole","Amoxicillin-Clavulanate","Imipenem","Other"];
const DELAY_REASONS = ["Waiting for physician order","Difficult IV access",
  "Waiting for medication from pharmacy","Diagnostic uncertainty",
  "Patient unstable / resuscitation in progress","Allergy clarification needed","Other"];
const STATUS_STEPS = [
  { key:"recognized",  label:"Sepsis Recognized",      short:"Recognized", icon:"🔴" },
  { key:"ordered",     label:"Antibiotic Ordered",      short:"Ordered",    icon:"📋" },
  { key:"prepared",    label:"Antibiotic Prepared",     short:"Prepared",   icon:"💊" },
  { key:"administered",label:"Antibiotic Administered", short:"Given",      icon:"✅" },
];
const VIEWS = ["board","patients","history","monthly","report"];

// ─── Utilities ────────────────────────────────────────────────────────────────
const nowISO  = () => new Date().toISOString();
const toDate  = (s) => s ? new Date(s) : null;
const elapsedSec = (start) => start ? Math.floor((Date.now() - new Date(start).getTime()) / 1000) : 0;
const diffMin = (a, b) => a && b ? ((new Date(b) - new Date(a)) / 60000) : null;
const fmtTime = (d) => d ? new Date(d).toLocaleTimeString("en-GB",{hour:"2-digit",minute:"2-digit",second:"2-digit"}) : "—";
const fmtDT   = (d) => d ? new Date(d).toLocaleString("en-GB",{day:"2-digit",month:"short",year:"numeric",hour:"2-digit",minute:"2-digit"}) : "—";
const fmtDate = (d) => d ? new Date(d).toLocaleDateString("en-GB",{day:"2-digit",month:"short",year:"numeric"}) : "—";
const secsToMMSS = (s) => {
  const abs = Math.abs(Math.round(s));
  return `${s<0?"+":""}${String(Math.floor(abs/60)).padStart(2,"0")}:${String(abs%60).padStart(2,"0")}`;
};

const urgencyOf = (el, done) => {
  if (done) return "done";
  if (el > DEADLINE_SEC)    return "overdue";
  if (el >= 56*60)          return "red";
  if (el >= 46*60)          return "orange";
  if (el >= 31*60)          return "yellow";
  return "green";
};
const U_ORDER  = { overdue:0, red:1, orange:2, yellow:3, green:4, done:5 };
const U_COLOR  = { green:"#00e676", yellow:"#ffea00", orange:"#ff6d00", red:"#f50057", overdue:"#ff1744", done:"#546e7a" };
const U_LABEL  = { green:"GREEN", yellow:"CAUTION", orange:"WARNING", red:"CRITICAL", overdue:"OVERDUE", done:"COMPLETED" };
const U_BG     = { green:"rgba(0,230,118,.12)", yellow:"rgba(255,234,0,.09)", orange:"rgba(255,109,0,.12)", red:"rgba(245,0,87,.12)", overdue:"rgba(183,28,28,.22)", done:"rgba(84,110,122,.2)" };

const currentStepKey = (steps) => {
  if (!steps.ordered)      return "recognized";
  if (!steps.prepared)     return "ordered";
  if (!steps.administered) return "prepared";
  return "administered";
};

const buildCase = (form, id) => {
  const ts = nowISO();
  const recISO = form.recognition ? new Date(form.recognition).toISOString() : null;
  return {
    CaseID: id,
    HN: form.hn, PatientName: form.name, BedNumber: form.bed,
    ArrivalTime:          form.arrival ? new Date(form.arrival).toISOString() : null,
    SepsisRecognitionTime: recISO,
    ATBOrderTime: null, ATBPreparedTime: null, ATBAdministeredTime: null, CompletedTime: null,
    SuspectedSource: form.source, PlannedAntibiotic: form.antibiotic,
    OrderingPhysician: form.physician,
    Status: "active",
    TotalTimeToATBMinutes: null,
    WithinOneHour: null,
    DelayReason: "", OtherDelayReasonDetail: "",
    CreatedBy: "ER Nurse", CreatedAt: ts, UpdatedAt: ts,
    // runtime only
    _steps: { recognized: recISO, ordered: null, prepared: null, administered: null },
    _alerts: {},
  };
};

// ─── Persistent Storage — Google Sheets backend ──────────────────────────────
// load/loadAudit: fetch from Sheets on startup
// save: no-op (we use granular appendCase/updateCase instead)
// saveAudit: no-op (appendAudit is called directly per entry)
const persist = {
  async load() {
    try { return await sheetsGet("getCases"); }
    catch(e) { console.error("Sheets load failed, falling back to localStorage", e);
      try { const r = localStorage.getItem(STORAGE_KEY); return r ? JSON.parse(r) : []; }
      catch { return []; }
    }
  },
  async save(cases) {
    // Granular saves handled by appendCase/updateCase/softDelete
    // Full save only as fallback to localStorage for offline resilience
    try { localStorage.setItem(STORAGE_KEY, JSON.stringify(cases)); }
    catch(e) { /* ignore */ }
  },
  async loadAudit() {
    try { return await sheetsGet("getAudit"); }
    catch(e) { console.error("Sheets audit load failed", e);
      try { const r = localStorage.getItem(AUDIT_KEY); return r ? JSON.parse(r) : []; }
      catch { return []; }
    }
  },
  async saveAudit(log) {
    // Audit entries appended individually via appendAudit — no bulk save needed
    try { localStorage.setItem(AUDIT_KEY, JSON.stringify(log)); }
    catch(e) { /* ignore */ }
  },
};

// ─── CSV / Excel Export ───────────────────────────────────────────────────────
const EXPORT_COLS = [
  "CaseID","HN","PatientName","BedNumber","ArrivalTime","SepsisRecognitionTime",
  "ATBOrderTime","ATBPreparedTime","ATBAdministeredTime","CompletedTime",
  "SuspectedSource","PlannedAntibiotic","OrderingPhysician","Status",
  "TotalTimeToATBMinutes","WithinOneHour","DelayReason","OtherDelayReasonDetail",
  "CreatedBy","CreatedAt","UpdatedAt",
];

const casesToRows = (cases) => cases.map(c => EXPORT_COLS.map(col => {
  const v = c[col];
  if (v === null || v === undefined) return "";
  if (["ArrivalTime","SepsisRecognitionTime","ATBOrderTime","ATBPreparedTime",
       "ATBAdministeredTime","CompletedTime","CreatedAt","UpdatedAt"].includes(col) && v)
    return fmtDT(v);
  return String(v);
}));

// ── Export helpers ────────────────────────────────────────────────────────────
// exportXLSX — true .xlsx via SheetJS (no browser-appended extension)
// exportCSV  — UTF-8 BOM CSV, MIME text/csv
// Both fire "sepsis_exported" so App shows a confirmation toast.

const _triggerDownload = (blob, filename) => {
  const url = URL.createObjectURL(blob);
  const a   = document.createElement("a");
  a.href        = url;
  a.download    = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 2000);
  window.dispatchEvent(new CustomEvent("sepsis_exported", { detail: { filename } }));
};

// True .xlsx using SheetJS — filename must end in .xlsx, no extra extension added
const exportXLSX = (cases, filename) => {
  // Build array-of-arrays: header row + data rows
  const data = [EXPORT_COLS, ...casesToRows(cases)];
  const ws   = XLSX.utils.aoa_to_sheet(data);
  // Bold the header row
  const range = XLSX.utils.decode_range(ws["!ref"]);
  for (let col = range.s.c; col <= range.e.c; col++) {
    const cell = ws[XLSX.utils.encode_cell({ r: 0, c: col })];
    if (cell) { cell.s = { font: { bold: true } }; }
  }
  // Set column widths
  ws["!cols"] = EXPORT_COLS.map(h => ({ wch: Math.max(h.length + 2, 16) }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "SepsisCases");
  // writeFile with type "array" → Uint8Array → Blob with correct MIME
  const wbArray = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  const blob    = new Blob([wbArray], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  _triggerDownload(blob, filename);
};

// CSV with UTF-8 BOM so Excel opens it correctly
const exportCSV = (cases, filename) => {
  const rows    = [EXPORT_COLS, ...casesToRows(cases)];
  const csv     = rows.map(r => r.map(v => `"${String(v).replace(/"/g,'""')}"`).join(",")).join("\n");
  const blob    = new Blob(["\uFEFF" + csv], { type: "text/csv;charset=utf-8;" });
  _triggerDownload(blob, filename);
};

const median = (arr) => {
  if (!arr.length) return null;
  const s = [...arr].sort((a,b)=>a-b);
  const m = Math.floor(s.length/2);
  return s.length%2 ? s[m] : (s[m-1]+s[m])/2;
};

// ═══════════════════════════════════════════════════════════════════════════════
// CSS
// ═══════════════════════════════════════════════════════════════════════════════
const CSS = `
@import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;500;600;700&family=Syne:wght@400;600;700;800&display=swap');
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0;}
:root{
  --bg:#080d1a;--panel:#0d1425;--panel2:#111928;--border:#1a2740;--border2:#223058;
  --text:#ddeeff;--muted:#4a6688;--accent:#1e6fff;--accent2:#0d4bcc;
  --green:#00e676;--yellow:#ffe600;--orange:#ff6600;--red:#f50057;--overdue:#c62828;
  --done:#37474f;--mono:'JetBrains Mono',monospace;--head:'Syne',sans-serif;
}
body{background:var(--bg);color:var(--text);font-family:var(--mono);font-size:14px;}
input,select,textarea,button{font-family:var(--mono);}

/* ── Layout ── */
.app{min-height:100vh;display:flex;flex-direction:column;}
.header{
  background:linear-gradient(135deg,#0a1428 0%,#070c18 100%);
  border-bottom:2px solid var(--border2);padding:10px 18px;
  display:flex;align-items:center;justify-content:space-between;
  position:sticky;top:0;z-index:200;gap:10px;flex-wrap:wrap;
}
.logo{font-family:var(--head);font-size:1.2rem;font-weight:800;letter-spacing:-0.5px;white-space:nowrap;}
.logo em{color:var(--red);font-style:normal;}
.livebadge{background:var(--red);color:#fff;font-size:.55rem;font-weight:700;
  padding:2px 7px;border-radius:20px;letter-spacing:1px;animation:blink 1.8s infinite;}
@keyframes blink{0%,100%{opacity:1}50%{opacity:.5}}
.nav{display:flex;gap:3px;flex-wrap:wrap;}
.nb{background:transparent;border:1px solid var(--border2);color:var(--muted);
  font-size:.68rem;padding:5px 12px;border-radius:5px;cursor:pointer;
  transition:all .18s;letter-spacing:.5px;white-space:nowrap;}
.nb:hover{border-color:var(--accent);color:var(--text);}
.nb.on{background:var(--accent);border-color:var(--accent);color:#fff;font-weight:700;}
.hclock{font-size:.78rem;color:var(--muted);letter-spacing:1px;white-space:nowrap;}

/* ── Stats bar ── */
.statsbar{display:grid;grid-template-columns:repeat(6,1fr);gap:8px;
  padding:12px 18px;border-bottom:1px solid var(--border);background:var(--panel);}
.stat{background:var(--bg);border:1px solid var(--border);border-radius:7px;
  padding:10px 14px;position:relative;overflow:hidden;}
.stat::before{content:'';position:absolute;top:0;left:0;right:0;height:3px;background:var(--sc,var(--accent));}
.sv{font-family:var(--head);font-size:1.6rem;font-weight:800;line-height:1;}
.sl{font-size:.58rem;color:var(--muted);text-transform:uppercase;letter-spacing:1px;margin-top:3px;}

/* ── View ── */
.view{padding:14px 18px;flex:1;}
.stitle{font-family:var(--head);font-size:.7rem;font-weight:700;
  text-transform:uppercase;letter-spacing:2px;color:var(--muted);
  margin-bottom:12px;display:flex;align-items:center;gap:8px;}
.stitle::after{content:'';flex:1;height:1px;background:var(--border);}

/* ── Board ── */
.board-wrap{background:var(--panel);border-radius:8px;overflow:hidden;border:1px solid var(--border);}
.bh,.br{
  display:grid;
  grid-template-columns:70px 1fr 90px 145px 100px 100px 105px;
  gap:6px;padding:10px 14px;
}
.bh{background:var(--border);font-size:.58rem;text-transform:uppercase;
  letter-spacing:1.5px;color:var(--muted);}
.br{border-bottom:1px solid var(--border);border-left:4px solid var(--rc,var(--border));
  cursor:pointer;transition:background .15s;align-items:center;
  animation:slin .25s ease;}
@keyframes slin{from{opacity:0;transform:translateX(-8px)}to{opacity:1;transform:translateX(0)}}
.br:hover{background:rgba(255,255,255,.025);}
.br:last-child{border-bottom:none;}
.bedtag{background:rgba(30,111,255,.15);border:1px solid var(--accent);
  color:var(--accent);font-weight:700;font-size:.72rem;
  padding:3px 7px;border-radius:4px;text-align:center;display:inline-block;}
.pname{font-size:.8rem;font-weight:600;}
.phn{font-size:.6rem;color:var(--muted);}
.etimer{font-size:.95rem;font-weight:700;}
.ctimer{font-size:.95rem;font-weight:700;letter-spacing:.5px;}
.spill{font-size:.6rem;padding:3px 9px;border-radius:20px;font-weight:700;
  letter-spacing:.5px;white-space:nowrap;text-align:center;}
.ubadge{font-size:.6rem;font-weight:700;padding:3px 9px;border-radius:20px;
  text-transform:uppercase;letter-spacing:1px;text-align:center;white-space:nowrap;}

/* ── Patient cards ── */
.pgrid{display:grid;grid-template-columns:repeat(auto-fill,minmax(330px,1fr));gap:12px;}
.card{background:var(--panel);border:1px solid var(--border);
  border-top:4px solid var(--cc,var(--border));border-radius:9px;padding:14px;
  transition:transform .18s,box-shadow .18s;position:relative;}
.card:hover{transform:translateY(-2px);box-shadow:0 8px 28px rgba(0,0,0,.5);}
.ch{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:10px;}
.cbed{font-family:var(--head);font-size:1.3rem;font-weight:800;color:var(--accent);}
.cname{font-size:.82rem;font-weight:600;}
.chn{font-size:.62rem;color:var(--muted);margin-top:1px;}
.bigtimer{text-align:center;padding:10px 0;
  font-family:var(--head);font-size:2.4rem;font-weight:800;letter-spacing:-1px;line-height:1;}
.tlabel{font-size:.58rem;text-transform:uppercase;letter-spacing:2px;
  color:var(--muted);text-align:center;margin-top:2px;}
.barwrap{height:5px;background:var(--border);border-radius:3px;margin:8px 0;overflow:hidden;}
.bar{height:100%;border-radius:3px;transition:width 1s linear,background .5s;}
.igrid{display:grid;grid-template-columns:1fr 1fr;gap:5px;margin:10px 0;}
.ii label{font-size:.56rem;text-transform:uppercase;letter-spacing:1px;color:var(--muted);display:block;}
.ii span{font-size:.72rem;font-weight:600;}
.strack{display:flex;gap:3px;margin:10px 0;}
.si{flex:1;text-align:center;padding:5px 3px;border-radius:5px;
  border:1px solid var(--border);font-size:.56rem;cursor:default;background:var(--bg);}
.si.done{background:rgba(0,230,118,.08);border-color:var(--green);color:var(--green);}
.si.act{border-color:var(--accent);background:rgba(30,111,255,.12);}
.si-icon{font-size:.9rem;display:block;margin-bottom:1px;}
.cactions{display:flex;gap:6px;margin-top:10px;}
.btn{flex:1;padding:7px;border:none;border-radius:6px;
  font-family:var(--mono);font-size:.68rem;font-weight:700;
  cursor:pointer;transition:all .18s;letter-spacing:.3px;}
.btn-p{background:var(--accent);color:#fff;}
.btn-p:hover{background:#2979ff;}
.btn-s{background:var(--green);color:#000;}
.btn-s:hover{filter:brightness(.9);}
.btn-d{background:var(--red);color:#fff;}
.btn-g{background:var(--border);color:var(--muted);}
.btn-g:hover{color:var(--text);background:var(--border2);}
.btn-o{background:rgba(255,102,0,.15);border:1px solid var(--orange);color:var(--orange);}
.btn:disabled{opacity:.3;cursor:not-allowed;}
.xbtn{display:inline-flex;align-items:center;gap:5px;padding:7px 14px;
  border:1px solid var(--border2);border-radius:6px;background:var(--panel2);
  color:var(--text);font-size:.7rem;font-weight:700;cursor:pointer;
  transition:all .18s;letter-spacing:.3px;white-space:nowrap;}
.xbtn:hover{border-color:var(--accent);color:var(--accent);}
.xbtn.xl{background:rgba(30,111,255,.12);border-color:var(--accent);}
.xbtn.csv{background:rgba(0,230,118,.08);border-color:var(--green);color:var(--green);}

/* ── Alerts ── */
.atoasts{position:fixed;top:64px;right:14px;width:300px;z-index:400;
  display:flex;flex-direction:column;gap:7px;}
.atoast{background:var(--panel);border-left:4px solid var(--ac,var(--red));
  border-radius:7px;padding:11px 13px;box-shadow:0 8px 28px rgba(0,0,0,.7);
  animation:tin .25s ease;}
@keyframes tin{from{opacity:0;transform:translateX(18px)}to{opacity:1;transform:translateX(0)}}
.at-title{font-size:.7rem;font-weight:700;}
.at-body{font-size:.62rem;color:var(--muted);margin-top:2px;}
.at-close{float:right;background:none;border:none;color:var(--muted);cursor:pointer;
  font-size:1rem;line-height:1;margin-left:6px;}

/* ── Modal ── */
.moverlay{position:fixed;inset:0;background:rgba(0,0,0,.85);z-index:300;
  display:flex;align-items:center;justify-content:center;padding:16px;}
.modal{background:var(--panel);border:1px solid var(--border2);border-radius:11px;
  width:100%;max-width:660px;max-height:92vh;overflow-y:auto;padding:22px;}
.modal-lg{max-width:800px;}
.mtitle{font-family:var(--head);font-size:1.15rem;font-weight:800;margin-bottom:18px;}
.fgrid{display:grid;grid-template-columns:1fr 1fr;gap:12px;}
.fg{display:flex;flex-direction:column;gap:3px;}
.fg.full{grid-column:1/-1;}
.fg label{font-size:.62rem;text-transform:uppercase;letter-spacing:1px;color:var(--muted);}
.fg input,.fg select,.fg textarea{
  background:var(--bg);border:1px solid var(--border);border-radius:6px;
  color:var(--text);font-size:.78rem;padding:7px 9px;outline:none;transition:border-color .18s;}
.fg input:focus,.fg select:focus{border-color:var(--accent);}
.fg select option{background:var(--panel);}
.mactions{display:flex;gap:8px;margin-top:18px;justify-content:flex-end;}

/* ── History / Table ── */
.tfilters{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:12px;align-items:center;}
.finput{background:var(--panel);border:1px solid var(--border);border-radius:6px;
  color:var(--text);font-size:.72rem;padding:6px 10px;outline:none;min-width:140px;}
.finput:focus{border-color:var(--accent);}
.twrap{background:var(--panel);border:1px solid var(--border);border-radius:8px;overflow:hidden;}
.twrap table{width:100%;border-collapse:collapse;font-size:.72rem;}
.twrap th{background:var(--border);padding:9px 12px;text-align:left;
  font-size:.58rem;text-transform:uppercase;letter-spacing:1.5px;color:var(--muted);
  white-space:nowrap;cursor:pointer;user-select:none;}
.twrap th:hover{color:var(--text);}
.twrap td{padding:9px 12px;border-bottom:1px solid var(--border);white-space:nowrap;}
.twrap tr:last-child td{border-bottom:none;}
.twrap tr:hover td{background:rgba(255,255,255,.02);}
.twrap tr.comp td{opacity:.75;}
.badge-sm{font-size:.58rem;padding:2px 7px;border-radius:12px;font-weight:700;}
.sortarr{font-size:.6rem;color:var(--accent);margin-left:2px;}

/* ── Report / Monthly ── */
.rgrid{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:18px;}
.rcard{background:var(--panel);border:1px solid var(--border);border-radius:8px;padding:14px;}
.rv{font-family:var(--head);font-size:1.9rem;font-weight:800;}
.rl{font-size:.6rem;color:var(--muted);text-transform:uppercase;letter-spacing:1px;margin-top:3px;}
.mselect{display:flex;gap:8px;align-items:center;margin-bottom:14px;flex-wrap:wrap;}
.msel-label{font-size:.68rem;color:var(--muted);text-transform:uppercase;letter-spacing:1px;}

/* ── Delay ── */
.dreason{background:rgba(255,102,0,.07);border:1px solid rgba(255,102,0,.3);
  border-radius:5px;padding:5px 9px;font-size:.65rem;color:var(--orange);margin-top:6px;}

/* ── Empty ── */
.empty{text-align:center;padding:60px 20px;color:var(--muted);}
.eicon{font-size:2.8rem;margin-bottom:10px;}

/* ── Add btn ── */
.addbtn{background:var(--accent);color:#fff;border:none;border-radius:7px;
  font-family:var(--mono);font-size:.75rem;font-weight:700;
  padding:9px 18px;cursor:pointer;transition:all .18s;
  display:inline-flex;align-items:center;gap:5px;}
.addbtn:hover{background:#2979ff;transform:translateY(-1px);}

/* ── Row toolbar ── */
.rowtb{display:flex;gap:4px;justify-content:flex-end;}
.ricon{background:none;border:1px solid var(--border);color:var(--muted);
  border-radius:4px;padding:3px 7px;cursor:pointer;font-size:.65rem;transition:all .18s;}
.ricon:hover{border-color:var(--accent);color:var(--accent);}

/* ── Progress ── */
.progress-wrap{margin:6px 0;}
.progress-bar-outer{height:4px;background:var(--border);border-radius:2px;overflow:hidden;margin-top:2px;}
.progress-bar-inner{height:100%;border-radius:2px;}

::-webkit-scrollbar{width:5px;height:5px;}
::-webkit-scrollbar-track{background:var(--bg);}
::-webkit-scrollbar-thumb{background:var(--border);border-radius:3px;}

.del-btn{background:rgba(245,0,87,.08);border:1px solid rgba(245,0,87,.35);color:var(--red);
  border-radius:5px;padding:4px 10px;cursor:pointer;font-size:.65rem;font-family:var(--mono);
  transition:all .18s;display:inline-flex;align-items:center;gap:4px;}
.del-btn:hover{background:rgba(245,0,87,.18);border-color:var(--red);}
.admin-badge{background:rgba(255,102,0,.12);border:1px solid var(--orange);color:var(--orange);
  font-size:.6rem;font-weight:700;padding:3px 9px;border-radius:20px;letter-spacing:.5px;}
.pin-input{background:var(--bg);border:1px solid var(--border);border-radius:6px;
  color:var(--text);font-size:1rem;padding:8px 12px;text-align:center;letter-spacing:4px;
  width:140px;outline:none;font-family:var(--mono);}
.pin-input:focus{border-color:var(--accent);}
@media(max-width:768px){
  .statsbar{grid-template-columns:repeat(3,1fr);}
  .bh,.br{grid-template-columns:60px 1fr 80px 100px;}
  .br>*:nth-child(n+5){display:none;}
  .pgrid{grid-template-columns:1fr;}
  .fgrid{grid-template-columns:1fr;}
  .rgrid{grid-template-columns:1fr;}
}
`;

// ═══════════════════════════════════════════════════════════════════════════════
// App
// ═══════════════════════════════════════════════════════════════════════════════
// ── Admin context (simple PIN-based for ER use) ───────────────────────────────
const ADMIN_PIN = "1234"; // Change in production

export default function App() {
  const [cases, setCases]             = useState([]);
  const [auditLog, setAuditLog]       = useState([]);
  const [loaded, setLoaded]           = useState(false);
  const [tick, setTick]               = useState(0);
  const [view, setView]               = useState("board");
  const [showForm, setShowForm]       = useState(false);
  const [detailId, setDetailId]       = useState(null);
  const [alerts, setAlerts]           = useState([]);
  const [clock, setClock]             = useState("");
  const [exportToast, setExportToast] = useState(null);
  const [isAdmin, setIsAdmin]         = useState(false);
  const [deleteTarget, setDeleteTarget] = useState(null); // case to delete
  const alertedRef                    = useRef({});
  const caseCounter                   = useRef(1);

  // Listen for export completion events
  useEffect(() => {
    const handler = (e) => {
      setExportToast(e.detail);
      setTimeout(() => setExportToast(null), 6500);
    };
    window.addEventListener("sepsis_exported", handler);
    return () => window.removeEventListener("sepsis_exported", handler);
  }, []);

  const [syncStatus, setSyncStatus] = useState("loading"); // loading | online | offline

  // Load from Google Sheets on startup
  useEffect(() => {
    setSyncStatus("loading");
    Promise.all([persist.load(), persist.loadAudit()]).then(([saved, audit]) => {
      if (saved && saved.length) {
        const maxId = saved.reduce((mx, c) => {
          const n = parseInt((c.CaseID || "").replace(/[^0-9]/g,""), 10);
          return n > mx ? n : mx;
        }, 0);
        caseCounter.current = maxId + 1;
      }
      setCases(saved || []);
      setAuditLog(audit || []);
      setLoaded(true);
      setSyncStatus("online");
    }).catch(() => {
      setLoaded(true);
      setSyncStatus("offline");
    });
  }, []);

  // Auto-refresh from Sheets every 30 seconds so all devices stay in sync
  useEffect(() => {
    if (!loaded) return;
    const id = setInterval(async () => {
      try {
        const fresh = await sheetsGet("getCases");
        if (fresh) {
          setCases(fresh);
          setSyncStatus("online");
        }
      } catch(e) {
        setSyncStatus("offline");
      }
    }, 30000);
    return () => clearInterval(id);
  }, [loaded]);

  // Save whenever cases change
  useEffect(() => {
    if (loaded) persist.save(cases);
  }, [cases, loaded]);

  // Save audit log whenever it changes
  useEffect(() => {
    if (loaded) persist.saveAudit(auditLog);
  }, [auditLog, loaded]);

  // Tick
  useEffect(() => {
    const id = setInterval(() => {
      setTick(t => t + 1);
      setClock(new Date().toLocaleTimeString("en-GB",{hour:"2-digit",minute:"2-digit",second:"2-digit"}));
    }, 1000);
    return () => clearInterval(id);
  }, []);

  // Alerts
  useEffect(() => {
    cases.filter(c => c.Status === "active" && c.SepsisRecognitionTime).forEach(c => {
      const el = elapsedSec(c.SepsisRecognitionTime);
      const fire = (min, title, body, color) => {
        const k = `${c.CaseID}-${min}`;
        if (!alertedRef.current[k]) {
          alertedRef.current[k] = true;
          const id = Date.now() + Math.random();
          setAlerts(prev => [...prev.slice(-4), { id, title, body, color }]);
          setTimeout(() => setAlerts(prev => prev.filter(a => a.id !== id)), 9000);
        }
      };
      if (el >= 30*60) fire(30, `⚠ ${c.PatientName} — ${c.BedNumber}`, "30 minutes elapsed since sepsis recognition.", "#ffe600");
      if (el >= 45*60) fire(45, `🔶 ${c.PatientName} — ${c.BedNumber}`, "Antibiotic administration approaching deadline.", "#ff6600");
      if (el >= 55*60) fire(55, `🚨 URGENT — ${c.PatientName}`, "5 minutes remaining before 1-hour target!", "#f50057");
      if (el >= 60*60) fire(60, `❌ OVERDUE — ${c.PatientName}`, "Antibiotics NOT administered within 1 hour.", "#c62828");
    });
  }, [tick, cases]);

  // Derived
  const activeCases    = cases.filter(c => c.Status === "active" && !c.IsDeleted);
  const completedCases = cases.filter(c => c.Status === "completed" && !c.IsDeleted);
  const withinHour     = completedCases.filter(c => c.WithinOneHour === "Yes");
  const overdueCases   = activeCases.filter(c => elapsedSec(c.SepsisRecognitionTime) > DEADLINE_SEC);
  const compliance     = completedCases.length ? Math.round((withinHour.length / completedCases.length) * 100) : null;
  const avgMin         = completedCases.length
    ? (completedCases.reduce((a,c)=>a+(c.TotalTimeToATBMinutes||0),0)/completedCases.length).toFixed(0)
    : null;

  const sortedActive = [...activeCases].sort((a,b) => {
    const ua = urgencyOf(elapsedSec(a.SepsisRecognitionTime), false);
    const ub = urgencyOf(elapsedSec(b.SepsisRecognitionTime), false);
    return (U_ORDER[ua]??9) - (U_ORDER[ub]??9);
  });

  // Mutations
  const addCase = useCallback(async (form) => {
    const id  = `CASE-${String(caseCounter.current++).padStart(4,"0")}`;
    const newCase = buildCase(form, id);
    // Optimistic update — add to UI immediately
    setCases(prev => [...prev, newCase]);
    setShowForm(false);
    // Persist to Google Sheets
    try {
      await sheetsPost({ action: "appendCase", case: newCase });
    } catch(e) {
      console.error("Failed to save case to Sheets:", e);
    }
  }, []);

  const advanceStep = useCallback(async (caseId) => {
    let updatedCase = null;
    // Optimistic update — change UI immediately
    setCases(prev => prev.map(c => {
      if (c.CaseID !== caseId) return c;
      const ts = nowISO();
      const updated = { ...c, UpdatedAt: ts };
      if (!updated.ATBOrderTime)         { updated.ATBOrderTime = ts;     updated._steps = {...c._steps, ordered: ts}; }
      else if (!updated.ATBPreparedTime) { updated.ATBPreparedTime = ts;  updated._steps = {...c._steps, prepared: ts}; }
      else if (!updated.ATBAdministeredTime) {
        updated.ATBAdministeredTime = ts;
        updated.CompletedTime = ts;
        updated.Status = "completed";
        const total = diffMin(updated.SepsisRecognitionTime, ts);
        updated.TotalTimeToATBMinutes = total ? +total.toFixed(1) : null;
        updated.WithinOneHour = total !== null ? (total <= 60 ? "Yes" : "No") : null;
        updated._steps = {...c._steps, administered: ts};
      }
      updatedCase = updated;
      return updated;
    }));
    // Persist to Google Sheets
    if (updatedCase) {
      try {
        await sheetsPost({ action: "updateCase", case: updatedCase });
      } catch(e) {
        console.error("Failed to sync step to Sheets:", e);
      }
    }
  }, []);

  const updateDelay = useCallback(async (caseId, reason, detail) => {
    const ts = nowISO();
    let updatedCase = null;
    setCases(prev => prev.map(c => {
      if (c.CaseID !== caseId) return c;
      updatedCase = { ...c, DelayReason: reason, OtherDelayReasonDetail: detail||"", UpdatedAt: ts };
      return updatedCase;
    }));
    if (updatedCase) {
      try {
        await sheetsPost({ action: "updateCase", case: { CaseID: caseId, DelayReason: reason, OtherDelayReasonDetail: detail||"", UpdatedAt: ts } });
      } catch(e) {
        console.error("Failed to sync delay reason to Sheets:", e);
      }
    }
  }, []);

  // Soft-delete: marks IsDeleted=true, appends audit entry
  const softDelete = useCallback(async (caseId, reason, otherDetail, deletedBy) => {
    const ts = nowISO();
    const c  = cases.find(x => x.CaseID === caseId);
    // Optimistic UI update
    setCases(prev => prev.map(x => {
      if (x.CaseID !== caseId) return x;
      return { ...x, IsDeleted: true, DeletedAt: ts, DeletedBy: deletedBy,
               DeleteReason: reason, DeleteReasonDetail: otherDetail||"", UpdatedAt: ts };
    }));
    if (c) {
      const entry = {
        AuditID: `AUDIT-${Date.now()}`,
        CaseID: c.CaseID, HN: c.HN, BedNumber: c.BedNumber,
        PatientName: c.PatientName, DeletedBy: deletedBy,
        DeletedAt: ts, DeleteReason: reason,
        DeleteReasonDetail: otherDetail||"",
        ActionType: "SOFT_DELETE",
      };
      setAuditLog(prev => [...prev, entry]);
      // Persist to Google Sheets
      try {
        await sheetsPost({
          action: "softDelete",
          CaseID: caseId, HN: c.HN, BedNumber: c.BedNumber,
          PatientName: c.PatientName, PreviousStatus: c.Status,
          DeletedBy: deletedBy, DeleteReason: reason,
          DeleteReasonDetail: otherDetail||"",
        });
      } catch(e) {
        console.error("Failed to sync soft-delete to Sheets:", e);
      }
    }
    setDeleteTarget(null);
  }, [cases]);

  const detailCase = detailId ? cases.find(c => c.CaseID === detailId) : null;

  // Export helpers
  const thisMonth = new Date().toISOString().slice(0,7);
  const monthlyExportCases = completedCases.filter(c => (c.CompletedTime||"").slice(0,7) === thisMonth);

  // Show loading screen while fetching from Sheets
  if (!loaded) return (
    <>
      <style>{CSS}</style>
      <div style={{minHeight:"100vh",background:"var(--bg)",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",gap:16}}>
        <div style={{fontFamily:"var(--head)",fontSize:"1.5rem",fontWeight:800,color:"var(--text)"}}>⚕ SEPSIS<span style={{color:"var(--red)"}}>TRACK</span></div>
        <div style={{fontSize:".8rem",color:"var(--muted)"}}>Connecting to Google Sheets…</div>
        <div style={{width:200,height:4,background:"var(--border)",borderRadius:2,overflow:"hidden"}}>
          <div style={{width:"60%",height:"100%",background:"var(--accent)",borderRadius:2,animation:"blink 1s infinite"}}/>
        </div>
      </div>
    </>
  );

  return (
    <>
      <style>{CSS}</style>
      <div className="app">
        {/* Header */}
        <header className="header">
          <div style={{display:"flex",alignItems:"center",gap:10}}>
            <div className="logo">⚕ SEPSIS<em>TRACK</em></div>
            <span className="livebadge">LIVE</span>
            <span className="livebadge" style={{
              background: syncStatus==="online"?"#00695c": syncStatus==="offline"?"#c62828":"#555",
              animation: syncStatus==="loading"?"blink 1s infinite":"none",
            }}>
              {syncStatus==="online"?"☁ SYNCED": syncStatus==="offline"?"⚠ OFFLINE":"⟳ CONNECTING"}
            </span>
            {overdueCases.length > 0 && <span className="livebadge" style={{background:"#c62828"}}>⚠ {overdueCases.length} OVERDUE</span>}
          </div>
          <nav className="nav">
            {[["board","🖥 Alert Board"],["patients","🏥 Patients"],["history","📂 History"],["monthly","📅 Monthly"],["report","📊 Report"],["audit","🔐 Audit"]].map(([k,l]) => (
              <button key={k} className={`nb${view===k?" on":""}`} onClick={() => setView(k)}>{l}</button>
            ))}
          </nav>
          <div style={{display:"flex",alignItems:"center",gap:8}}>
            <div className="hclock">{clock}</div>
            <AdminToggle isAdmin={isAdmin} onToggle={setIsAdmin} />
          </div>
        </header>

        {/* Stats */}
        <div className="statsbar">
          {[
            {l:"Active",v:activeCases.length,c:"var(--accent)"},
            {l:"Completed",v:completedCases.length,c:"var(--done)"},
            {l:"Within 1hr",v:withinHour.length,c:"var(--green)"},
            {l:"Overdue",v:overdueCases.length,c:"var(--red)"},
            {l:"Compliance",v:compliance!==null?`${compliance}%`:"—",c:compliance===null?"var(--muted)":compliance>=80?"var(--green)":compliance>=60?"var(--yellow)":"var(--red)"},
            {l:"Avg Min",v:avgMin||"—",c:"var(--yellow)"},
          ].map(s => (
            <div key={s.l} className="stat" style={{"--sc":s.c}}>
              <div className="sv" style={{color:s.c}}>{s.v}</div>
              <div className="sl">{s.l}</div>
            </div>
          ))}
        </div>

        {/* Main */}
        <main className="view">
          {view==="board"    && <BoardView    cases={sortedActive} completed={completedCases} tick={tick} onAdd={() => setShowForm(true)} onDetail={setDetailId} isAdmin={isAdmin} onDelete={setDeleteTarget} />}
          {view==="patients" && <PatientsView cases={sortedActive} tick={tick} onAdd={() => setShowForm(true)} onAdvance={advanceStep} onUpdateDelay={updateDelay} onDetail={setDetailId} isAdmin={isAdmin} onDelete={setDeleteTarget} />}
          {view==="history"  && <HistoryView  cases={completedCases} onDetail={setDetailId} exportCSV={exportCSV} exportXLSX={exportXLSX} isAdmin={isAdmin} onDelete={setDeleteTarget} />}
          {view==="monthly"  && <MonthlyView  cases={completedCases} exportCSV={exportCSV} exportXLSX={exportXLSX} isAdmin={isAdmin} />}
          {view==="report"   && <ReportView   cases={[...activeCases,...completedCases]} completed={completedCases} withinHour={withinHour} avgMin={avgMin} compliance={compliance} exportCSV={exportCSV} exportXLSX={exportXLSX} isAdmin={isAdmin} />}
          {view==="audit"    && <AuditView    log={auditLog} deletedCases={cases.filter(c=>c.IsDeleted)} exportCSV={exportCSV} exportXLSX={exportXLSX} />}
        </main>

        {/* Alert toasts */}
        <div className="atoasts">
          {alerts.map(a => (
            <div key={a.id} className="atoast" style={{"--ac":a.color}}>
              <button className="at-close" onClick={() => setAlerts(p => p.filter(x => x.id !== a.id))}>×</button>
              <div className="at-title">{a.title}</div>
              <div className="at-body">{a.body}</div>
            </div>
          ))}
        </div>

        {/* Export this month - floating */}
        {isAdmin && (
          <div style={{position:"fixed",bottom:16,right:16,zIndex:150,display:"flex",gap:6}}>
            <button className="xbtn xl" onClick={() => exportXLSX(monthlyExportCases, `SepsisTrack_Monthly_Report_${thisMonth}.xlsx`)}>
              📥 Export This Month — Excel ({monthlyExportCases.length} cases)
            </button>
            <button className="xbtn csv" onClick={() => exportCSV(monthlyExportCases, `SepsisTrack_Monthly_Report_${thisMonth}.csv`)}>
              CSV
            </button>
          </div>
        )}

        {/* Forms / modals */}
        {showForm  && <PatientForm onSave={addCase} onClose={() => setShowForm(false)} />}
        {detailCase && (
          <DetailModal c={detailCase} tick={tick} onAdvance={advanceStep} onUpdateDelay={updateDelay} onClose={() => setDetailId(null)} />
        )}
        {exportToast && (
          <ExportToast filename={exportToast.filename} fallback={exportToast.fallback} onClose={() => setExportToast(null)} />
        )}
        {deleteTarget && (
          <DeleteModal
            c={deleteTarget}
            isAdmin={isAdmin}
            onConfirm={(reason, detail, by) => softDelete(deleteTarget.CaseID, reason, detail, by)}
            onClose={() => setDeleteTarget(null)}
          />
        )}
      </div>
    </>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// Board View
// ═══════════════════════════════════════════════════════════════════════════════
function BoardView({ cases, completed, tick, onAdd, onDetail, isAdmin, onDelete }) {
  return (
    <div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
        <div className="stitle" style={{margin:0,flex:1}}>🖥 SEPSIS ALERT BOARD — ACTIVE PATIENTS</div>
        <button className="addbtn" onClick={onAdd}>+ Register Patient</button>
      </div>

      {cases.length === 0 && completed.length === 0 ? (
        <div className="empty"><div className="eicon">✅</div><div>No active sepsis cases.</div></div>
      ) : (
        <>
          {cases.length > 0 && (
            <div className="board-wrap" style={{marginBottom:16}}>
              <div className="bh"><span>BED</span><span>PATIENT</span><span>ELAPSED</span><span>STATUS</span><span>COUNTDOWN</span><span>ANTIBIOTIC</span><span>URGENCY</span></div>
              {cases.map(c => <BoardRow key={c.CaseID} c={c} tick={tick} onClick={() => onDetail(c.CaseID)} />)}
            </div>
          )}
          {completed.length > 0 && (
            <>
              <div className="stitle">✅ RECENTLY COMPLETED (last 10)</div>
              <div className="board-wrap">
                <div className="bh"><span>BED</span><span>PATIENT</span><span>TOTAL TIME</span><span>STATUS</span><span>COMPLETED</span><span>ANTIBIOTIC</span><span>RESULT</span></div>
                {[...completed].sort((a,b) => new Date(b.CompletedTime)-new Date(a.CompletedTime)).slice(0,10).map(c => <BoardRowDone key={c.CaseID} c={c} onClick={() => onDetail(c.CaseID)} isAdmin={isAdmin} onDelete={onDelete} />)}
              </div>
            </>
          )}
        </>
      )}
    </div>
  );
}

function BoardRow({ c, tick, onClick }) {
  const el = elapsedSec(c.SepsisRecognitionTime);
  const u  = urgencyOf(el, false);
  const remaining = DEADLINE_SEC - el;
  const col = U_COLOR[u];
  const step = currentStepKey(c._steps || {});
  const stepInfo = STATUS_STEPS.find(s => s.key === step);
  return (
    <div className="br" style={{"--rc":col}} onClick={onClick}>
      <div><span className="bedtag">{c.BedNumber}</span></div>
      <div><div className="pname">{c.PatientName}</div><div className="phn">{c.HN}</div></div>
      <div className="etimer" style={{color:col}}>{secsToMMSS(el)}</div>
      <div><span className="spill" style={{background:U_BG[u],color:col,border:`1px solid ${col}`}}>{stepInfo?.icon} {stepInfo?.short}</span></div>
      <div className="ctimer" style={{color:remaining<=0?"#ff1744":col}}>{remaining<=0?`+${secsToMMSS(-remaining)}`:secsToMMSS(remaining)}</div>
      <div style={{fontSize:".68rem",color:"var(--muted)"}}>{c.PlannedAntibiotic?.split(" ")[0]}</div>
      <div><span className="ubadge" style={{background:U_BG[u],color:col,border:`1px solid ${col}`}}>{U_LABEL[u]}</span></div>
    </div>
  );
}

function BoardRowDone({ c, onClick, isAdmin, onDelete }) {
  const total = c.TotalTimeToATBMinutes;
  const within = c.WithinOneHour === "Yes";
  return (
    <div className="br" style={{"--rc":"var(--done)",opacity:.75}} onClick={onClick}>
      <div><span className="bedtag" style={{opacity:.5}}>{c.BedNumber}</span></div>
      <div><div className="pname">{c.PatientName}</div><div className="phn">{c.HN}</div></div>
      <div style={{color:within?"var(--green)":"var(--red)",fontWeight:700}}>{total?`${total} min`:"—"}</div>
      <div><span className="spill" style={{background:"rgba(0,230,118,.08)",color:"var(--green)",border:"1px solid var(--green)"}}>✅ Administered</span></div>
      <div style={{fontSize:".7rem",color:"var(--muted)"}}>{fmtTime(c.ATBAdministeredTime)}</div>
      <div style={{fontSize:".68rem",color:"var(--muted)"}}>{c.PlannedAntibiotic?.split(" ")[0]}</div>
      <div style={{display:"flex",gap:6,alignItems:"center"}}>
        <span className="ubadge" style={{background:within?"rgba(0,230,118,.1)":"rgba(245,0,87,.1)",color:within?"var(--green)":"var(--red)",border:`1px solid ${within?"var(--green)":"var(--red)"}`}}>{within?"✓ ON TIME":"✗ DELAYED"}</span>
        {isAdmin && <button className="ricon" style={{color:"var(--red)",borderColor:"var(--red)"}} onClick={e=>{e.stopPropagation();onDelete(c);}}>🗑</button>}
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// Patients View
// ═══════════════════════════════════════════════════════════════════════════════
function PatientsView({ cases, tick, onAdd, onAdvance, onUpdateDelay, onDetail, isAdmin, onDelete }) {
  return (
    <div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
        <div className="stitle" style={{margin:0,flex:1}}>🏥 ACTIVE PATIENT MANAGEMENT</div>
        <button className="addbtn" onClick={onAdd}>+ Register Patient</button>
      </div>
      {cases.length === 0 ? (
        <div className="empty"><div className="eicon">🏥</div><div>No active patients.</div></div>
      ) : (
        <div className="pgrid">
          {cases.map(c => <PatientCard key={c.CaseID} c={c} tick={tick} onAdvance={onAdvance} onUpdateDelay={onUpdateDelay} onDetail={onDetail} isAdmin={isAdmin} onDelete={onDelete} />)}
        </div>
      )}
    </div>
  );
}

function PatientCard({ c, tick, onAdvance, onUpdateDelay, onDetail, isAdmin, onDelete }) {
  const el   = elapsedSec(c.SepsisRecognitionTime);
  const done = c.Status === "completed";
  const u    = urgencyOf(el, done);
  const col  = U_COLOR[u];
  const remaining = DEADLINE_SEC - el;
  const pct  = Math.min(100, (el/DEADLINE_SEC)*100);
  const step = currentStepKey(c._steps||{});
  const nextStep = STATUS_STEPS[STATUS_STEPS.findIndex(s=>s.key===step)+1];

  return (
    <div className="card" style={{"--cc":col}}>
      <div className="ch">
        <div>
          <div className="cbed">{c.BedNumber}</div>
          <div className="cname">{c.PatientName}</div>
          <div className="chn">{c.HN} · {c.OrderingPhysician}</div>
        </div>
        <span className="ubadge" style={{background:U_BG[u],color:col,border:`1px solid ${col}`}}>{U_LABEL[u]}</span>
      </div>

      <div className="bigtimer" style={{color:col}}>{secsToMMSS(el)}</div>
      <div className="tlabel">TIME ELAPSED SINCE RECOGNITION</div>
      <div className="barwrap"><div className="bar" style={{width:`${pct}%`,background:col}}/></div>
      <div style={{display:"flex",justifyContent:"space-between",fontSize:".6rem",color:"var(--muted)"}}>
        <span>0:00</span>
        <span style={{color:remaining>0?col:"#ff1744",fontWeight:700}}>
          {remaining>0?`${secsToMMSS(remaining)} remaining`:`OVERDUE ${secsToMMSS(-remaining)}`}
        </span>
        <span>60:00</span>
      </div>

      <div className="igrid">
        <div className="ii"><label>Source</label><span>{c.SuspectedSource}</span></div>
        <div className="ii"><label>Antibiotic</label><span style={{fontSize:".66rem"}}>{c.PlannedAntibiotic}</span></div>
        <div className="ii"><label>Recognition</label><span>{fmtTime(c.SepsisRecognitionTime)}</span></div>
        <div className="ii"><label>Arrival</label><span>{fmtTime(c.ArrivalTime)}</span></div>
      </div>

      <div className="strack">
        {STATUS_STEPS.map(s => {
          const ts = c._steps?.[s.key];
          const isDone = !!ts;
          const isAct  = s.key === step && !isDone;
          return (
            <div key={s.key} className={`si${isDone?" done":isAct?" act":""}`}>
              <span className="si-icon">{isDone?"✅":s.icon}</span>
              <span>{s.short}</span>
              {isDone && <div style={{fontSize:".5rem",color:"var(--muted)",marginTop:1}}>{fmtTime(ts)}</div>}
            </div>
          );
        })}
      </div>

      {c.DelayReason && <div className="dreason">⚠ Delay: {c.DelayReason}{c.OtherDelayReasonDetail?" — "+c.OtherDelayReasonDetail:""}</div>}

      <div className="cactions">
        {nextStep && <button className="btn btn-p" onClick={() => onAdvance(c.CaseID)}>{nextStep.icon} {nextStep.label}</button>}
        <button className="btn btn-g" onClick={() => onDetail(c.CaseID)}>🔍</button>
        {isAdmin && <button className="btn btn-g" style={{color:"var(--red)",borderColor:"rgba(245,0,87,.4)",flex:"0 0 auto"}} onClick={() => onDelete(c)}>🗑</button>}
      </div>

      {u !== "green" && !c.DelayReason && (
        <DelaySelector caseId={c.CaseID} current={c.DelayReason} detail={c.OtherDelayReasonDetail} onSave={onUpdateDelay} />
      )}
    </div>
  );
}

function DelaySelector({ caseId, current, detail, onSave }) {
  const [reason, setReason] = useState(current||"");
  const [other, setOther]   = useState(detail||"");
  return (
    <div style={{marginTop:8}}>
      <select value={reason} onChange={e => { setReason(e.target.value); if(e.target.value !== "Other") onSave(caseId, e.target.value, ""); }}
        style={{width:"100%",background:"var(--bg)",border:"1px solid var(--border)",color:"var(--muted)",fontSize:".68rem",padding:"5px 8px",borderRadius:5}}>
        <option value="">Record delay reason…</option>
        {DELAY_REASONS.map(r => <option key={r}>{r}</option>)}
      </select>
      {reason === "Other" && (
        <div style={{marginTop:5,display:"flex",gap:5}}>
          <input value={other} onChange={e => setOther(e.target.value)} placeholder="Specify other reason…"
            style={{flex:1,background:"var(--bg)",border:"1px solid var(--border)",borderRadius:5,color:"var(--text)",fontSize:".68rem",padding:"5px 8px"}}/>
          <button className="btn btn-o" style={{flex:"0 0 auto",padding:"4px 10px"}} onClick={() => onSave(caseId, reason, other)}>Save</button>
        </div>
      )}
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// History View
// ═══════════════════════════════════════════════════════════════════════════════
function HistoryView({ cases, onDetail, exportCSV, exportXLSX, isAdmin, onDelete }) {
  const [search,   setSearch]   = useState("");
  const [dateFrom, setDateFrom] = useState("");
  const [dateTo,   setDateTo]   = useState("");
  const [filter,   setFilter]   = useState("all"); // all | within | delayed
  const [sortCol,  setSortCol]  = useState("SepsisRecognitionTime");
  const [sortDir,  setSortDir]  = useState(-1);

  const toggleSort = (col) => {
    if (sortCol === col) setSortDir(d => -d);
    else { setSortCol(col); setSortDir(-1); }
  };

  const filtered = cases.filter(c => {
    const q = search.toLowerCase();
    if (q && ![c.HN,c.PatientName,c.BedNumber,c.OrderingPhysician,c.PlannedAntibiotic,c.Status].join(" ").toLowerCase().includes(q)) return false;
    if (dateFrom && c.SepsisRecognitionTime < new Date(dateFrom).toISOString()) return false;
    if (dateTo   && c.SepsisRecognitionTime > new Date(dateTo+"T23:59:59").toISOString()) return false;
    if (filter === "within"  && c.WithinOneHour !== "Yes") return false;
    if (filter === "delayed" && c.WithinOneHour !== "No") return false;
    return true;
  }).sort((a,b) => {
    const av = a[sortCol]??""; const bv = b[sortCol]??"";
    return av < bv ? sortDir : av > bv ? -sortDir : 0;
  });

  const TH = ({col,label}) => (
    <th onClick={() => toggleSort(col)}>
      {label}{sortCol===col?<span className="sortarr">{sortDir===-1?"▼":"▲"}</span>:""}
    </th>
  );

  return (
    <div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10,flexWrap:"wrap",gap:8}}>
        <div className="stitle" style={{margin:0}}>📂 SEPSIS CASE HISTORY ({filtered.length} / {cases.length})</div>
        {isAdmin ? (
          <div style={{display:"flex",gap:6}}>
            <button className="xbtn xl" onClick={() => exportXLSX(filtered,`SepsisTrack_History_${new Date().toISOString().slice(0,10)}.xlsx`)}>📥 Export Excel</button>
            <button className="xbtn csv" onClick={() => exportCSV(filtered,`SepsisTrack_History_${new Date().toISOString().slice(0,10)}.csv`)}>CSV</button>
          </div>
        ) : (
          <span style={{fontSize:".68rem",color:"var(--muted)",display:"flex",alignItems:"center",gap:5}}>
            🔐 Admin login required to export
          </span>
        )}
      </div>

      <div className="tfilters">
        <input className="finput" placeholder="🔍 Search HN, name, bed, doctor, antibiotic…" value={search} onChange={e => setSearch(e.target.value)} style={{minWidth:260}}/>
        <input className="finput" type="date" value={dateFrom} onChange={e => setDateFrom(e.target.value)} title="From date"/>
        <span style={{color:"var(--muted)",fontSize:".72rem"}}>to</span>
        <input className="finput" type="date" value={dateTo} onChange={e => setDateTo(e.target.value)} title="To date"/>
        {[["all","All"],["within","✓ Within 1hr"],["delayed","✗ Delayed"]].map(([k,l]) => (
          <button key={k} className={`nb${filter===k?" on":""}`} onClick={() => setFilter(k)}>{l}</button>
        ))}
        <button className="nb" onClick={() => { setSearch(""); setDateFrom(""); setDateTo(""); setFilter("all"); }}>✕ Clear</button>
      </div>

      {filtered.length === 0 ? (
        <div className="empty"><div className="eicon">🔍</div><div>No cases match your filters.</div></div>
      ) : (
        <div className="twrap">
          <table>
            <thead>
              <tr>
                <TH col="CaseID" label="Case ID" />
                <TH col="HN" label="HN" />
                <TH col="PatientName" label="Patient" />
                <TH col="BedNumber" label="Bed" />
                <TH col="SepsisRecognitionTime" label="Recognition" />
                <TH col="ATBAdministeredTime" label="Administered" />
                <TH col="TotalTimeToATBMinutes" label="Time (min)" />
                <TH col="WithinOneHour" label="1hr Target" />
                <th>Delay Reason</th>
                <th>Physician</th>
                <th>Source</th>
                <th>Action</th>
              </tr>
            </thead>
            <tbody>
              {filtered.map(c => {
                const within = c.WithinOneHour === "Yes";
                return (
                  <tr key={c.CaseID} className="comp">
                    <td style={{color:"var(--muted)",fontSize:".65rem"}}>{c.CaseID}</td>
                    <td style={{color:"var(--muted)"}}>{c.HN}</td>
                    <td style={{fontWeight:600}}>{c.PatientName}</td>
                    <td><span className="bedtag" style={{fontSize:".62rem"}}>{c.BedNumber}</span></td>
                    <td style={{color:"var(--muted)",fontSize:".7rem"}}>{fmtDT(c.SepsisRecognitionTime)}</td>
                    <td style={{color:"var(--muted)",fontSize:".7rem"}}>{fmtDT(c.ATBAdministeredTime)}</td>
                    <td style={{color:within?"var(--green)":"var(--red)",fontWeight:700}}>{c.TotalTimeToATBMinutes??"-"}</td>
                    <td>
                      <span className="badge-sm" style={{background:within?"rgba(0,230,118,.12)":"rgba(245,0,87,.12)",color:within?"var(--green)":"var(--red)",border:`1px solid ${within?"var(--green)":"var(--red)"}`}}>
                        {within?"✓ Yes":"✗ No"}
                      </span>
                    </td>
                    <td style={{fontSize:".68rem",color:"var(--orange)",maxWidth:150,overflow:"hidden",textOverflow:"ellipsis"}}>
                      {c.DelayReason||"—"}{c.OtherDelayReasonDetail?" ("+c.OtherDelayReasonDetail+")":""}
                    </td>
                    <td style={{color:"var(--muted)",fontSize:".7rem"}}>{c.OrderingPhysician}</td>
                    <td style={{color:"var(--muted)",fontSize:".7rem"}}>{c.SuspectedSource}</td>
                    <td>
                      <div className="rowtb">
                        <button className="ricon" onClick={() => onDetail(c.CaseID)}>🔍 View</button>
                        {isAdmin && <button className="ricon" style={{color:"var(--red)",borderColor:"rgba(245,0,87,.4)"}} onClick={() => onDelete(c)}>🗑 Delete</button>}
                      </div>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// Monthly View
// ═══════════════════════════════════════════════════════════════════════════════
function MonthlyView({ cases, exportCSV, exportXLSX, isAdmin }) {
  const months = [...new Set(cases.map(c => (c.CompletedTime||c.SepsisRecognitionTime||"").slice(0,7)))].filter(Boolean).sort().reverse();
  const [selMonth, setSelMonth] = useState(months[0] || new Date().toISOString().slice(0,7));

  const mCases = cases.filter(c => (c.CompletedTime||c.SepsisRecognitionTime||"").slice(0,7) === selMonth);
  const mWithin = mCases.filter(c => c.WithinOneHour === "Yes");
  const mDelayed = mCases.filter(c => c.WithinOneHour === "No");
  const times = mCases.filter(c => c.TotalTimeToATBMinutes).map(c => c.TotalTimeToATBMinutes);
  const mAvg = times.length ? (times.reduce((a,b)=>a+b,0)/times.length).toFixed(1) : null;
  const mMedian = times.length ? median(times)?.toFixed(1) : null;
  const mComp = mCases.length ? Math.round((mWithin.length/mCases.length)*100) : null;

  const delayCounts = {};
  DELAY_REASONS.forEach(r => delayCounts[r] = 0);
  mCases.forEach(c => { if (c.DelayReason && delayCounts[c.DelayReason]!==undefined) delayCounts[c.DelayReason]++; });

  const bySrc = {};
  mCases.forEach(c => { if (c.SuspectedSource) bySrc[c.SuspectedSource]=(bySrc[c.SuspectedSource]||0)+1; });

  return (
    <div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12,flexWrap:"wrap",gap:8}}>
        <div className="stitle" style={{margin:0}}>📅 MONTHLY QUALITY IMPROVEMENT REPORT</div>
        {isAdmin ? (
          <div style={{display:"flex",gap:6}}>
            <button className="xbtn xl" onClick={() => exportXLSX(mCases, `SepsisTrack_Monthly_Report_${selMonth}.xlsx`)}>📥 Excel</button>
            <button className="xbtn csv" onClick={() => exportCSV(mCases, `SepsisTrack_Monthly_Report_${selMonth}.csv`)}>CSV</button>
          </div>
        ) : (
          <span style={{fontSize:".68rem",color:"var(--muted)",display:"flex",alignItems:"center",gap:5}}>
            🔐 Admin login required to export
          </span>
        )}
      </div>

      <div className="mselect">
        <span className="msel-label">Month:</span>
        <select className="finput" value={selMonth} onChange={e => setSelMonth(e.target.value)} style={{minWidth:160}}>
          {months.map(m => <option key={m} value={m}>{new Date(m+"-15").toLocaleDateString("en-GB",{month:"long",year:"numeric"})}</option>)}
        </select>
        <span style={{fontSize:".72rem",color:"var(--muted)"}}>{mCases.length} cases this month</span>
      </div>

      <div className="rgrid">
        {[
          {l:"Total Cases",v:mCases.length,c:"var(--accent)"},
          {l:"Within 1 Hour",v:mWithin.length,c:"var(--green)"},
          {l:"Delayed",v:mDelayed.length,c:"var(--red)"},
          {l:"Compliance %",v:mComp!==null?`${mComp}%`:"—",c:mComp===null?"var(--muted)":mComp>=80?"var(--green)":mComp>=60?"var(--yellow)":"var(--red)"},
          {l:"Median Time (min)",v:mMedian||"—",c:"var(--yellow)"},
          {l:"Average Time (min)",v:mAvg||"—",c:"var(--orange)"},
        ].map(s => (
          <div key={s.l} className="rcard">
            <div className="rv" style={{color:s.c}}>{s.v}</div>
            <div className="rl">{s.l}</div>
          </div>
        ))}
      </div>

      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14,marginBottom:14}}>
        <div>
          <div className="stitle">Delay Reasons</div>
          <div className="twrap">
            <table>
              <thead><tr><th>Reason</th><th>Count</th></tr></thead>
              <tbody>
                {Object.entries(delayCounts).filter(([,v])=>v>0).length===0
                  ? <tr><td colSpan={2} style={{color:"var(--muted)",textAlign:"center",padding:20}}>No delays recorded</td></tr>
                  : Object.entries(delayCounts).sort((a,b)=>b[1]-a[1]).map(([r,v]) => (
                    <tr key={r}>
                      <td>{r}</td>
                      <td><span style={{color:"var(--orange)",fontWeight:700}}>{v}</span></td>
                    </tr>
                  ))
                }
              </tbody>
            </table>
          </div>
        </div>
        <div>
          <div className="stitle">Cases by Infection Source</div>
          <div className="twrap">
            <table>
              <thead><tr><th>Source</th><th>Count</th><th>%</th></tr></thead>
              <tbody>
                {Object.entries(bySrc).length===0
                  ? <tr><td colSpan={3} style={{color:"var(--muted)",textAlign:"center",padding:20}}>No data</td></tr>
                  : Object.entries(bySrc).sort((a,b)=>b[1]-a[1]).map(([s,v]) => (
                    <tr key={s}>
                      <td>{s}</td>
                      <td style={{color:"var(--accent)",fontWeight:700}}>{v}</td>
                      <td style={{color:"var(--muted)"}}>{mCases.length?Math.round(v/mCases.length*100):0}%</td>
                    </tr>
                  ))
                }
              </tbody>
            </table>
          </div>
        </div>
      </div>

      {/* Compliance bar */}
      {mComp !== null && (
        <div style={{background:"var(--panel)",border:"1px solid var(--border)",borderRadius:8,padding:16,marginBottom:14}}>
          <div style={{display:"flex",justifyContent:"space-between",marginBottom:6}}>
            <span style={{fontSize:".7rem",color:"var(--muted)",textTransform:"uppercase",letterSpacing:1}}>1-Hour Compliance Rate</span>
            <span style={{fontFamily:"var(--head)",fontWeight:800,fontSize:"1.1rem",color:mComp>=80?"var(--green)":mComp>=60?"var(--yellow)":"var(--red)"}}>{mComp}%</span>
          </div>
          <div className="barwrap" style={{height:12}}>
            <div className="bar" style={{width:`${mComp}%`,background:mComp>=80?"var(--green)":mComp>=60?"var(--yellow)":"var(--red)",height:12}}/>
          </div>
          <div style={{fontSize:".65rem",color:"var(--muted)",marginTop:4}}>Target: ≥80% · {mWithin.length}/{mCases.length} cases administered within 1 hour</div>
        </div>
      )}

      <div className="stitle">Case Log — {selMonth}</div>
      <div className="twrap">
        <table>
          <thead>
            <tr><th>HN</th><th>Patient</th><th>Bed</th><th>Recognition</th><th>Administered</th><th>Time (min)</th><th>1hr Target</th><th>Delay Reason</th></tr>
          </thead>
          <tbody>
            {mCases.length===0
              ? <tr><td colSpan={8} style={{color:"var(--muted)",textAlign:"center",padding:20}}>No cases this month.</td></tr>
              : mCases.map(c => {
                  const within = c.WithinOneHour === "Yes";
                  return (
                    <tr key={c.CaseID}>
                      <td style={{color:"var(--muted)",fontSize:".68rem"}}>{c.HN}</td>
                      <td style={{fontWeight:600}}>{c.PatientName}</td>
                      <td><span className="bedtag" style={{fontSize:".62rem"}}>{c.BedNumber}</span></td>
                      <td style={{color:"var(--muted)",fontSize:".68rem"}}>{fmtDT(c.SepsisRecognitionTime)}</td>
                      <td style={{color:"var(--muted)",fontSize:".68rem"}}>{fmtDT(c.ATBAdministeredTime)}</td>
                      <td style={{color:within?"var(--green)":"var(--red)",fontWeight:700}}>{c.TotalTimeToATBMinutes??"-"}</td>
                      <td><span className="badge-sm" style={{background:within?"rgba(0,230,118,.12)":"rgba(245,0,87,.12)",color:within?"var(--green)":"var(--red)",border:`1px solid ${within?"var(--green)":"var(--red)"}`}}>{within?"✓ Yes":"✗ No"}</span></td>
                      <td style={{fontSize:".68rem",color:"var(--orange)"}}>{c.DelayReason||"—"}</td>
                    </tr>
                  );
              })
            }
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// Report View
// ═══════════════════════════════════════════════════════════════════════════════
function ReportView({ cases, completed, withinHour, avgMin, compliance, exportCSV, exportXLSX, isAdmin }) {
  const delayCounts = {};
  DELAY_REASONS.forEach(r => delayCounts[r]=0);
  cases.forEach(c => { if(c.DelayReason && delayCounts[c.DelayReason]!==undefined) delayCounts[c.DelayReason]++; });
  const bySrc = {};
  cases.forEach(c => { if(c.SuspectedSource) bySrc[c.SuspectedSource]=(bySrc[c.SuspectedSource]||0)+1; });
  const times = completed.filter(c=>c.TotalTimeToATBMinutes).map(c=>c.TotalTimeToATBMinutes);
  const med   = times.length ? median(times)?.toFixed(1) : null;

  return (
    <div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12,flexWrap:"wrap",gap:8}}>
        <div className="stitle" style={{margin:0}}>📊 OVERALL QI REPORT — ALL TIME</div>
        {isAdmin ? (
          <div style={{display:"flex",gap:6}}>
            <button className="xbtn xl" onClick={() => exportXLSX(completed,`SepsisTrack_All_Cases_${new Date().toISOString().slice(0,10)}.xlsx`)}>📥 All Data Excel</button>
            <button className="xbtn csv" onClick={() => exportCSV(completed,`SepsisTrack_All_Cases_${new Date().toISOString().slice(0,10)}.csv`)}>CSV</button>
          </div>
        ) : (
          <span style={{fontSize:".68rem",color:"var(--muted)",display:"flex",alignItems:"center",gap:5}}>
            🔐 Admin login required to export
          </span>
        )}
      </div>
      <div className="rgrid">
        {[
          {l:"Total Cases",v:cases.length,c:"var(--accent)"},
          {l:"Completed",v:completed.length,c:"var(--done)"},
          {l:"Within 1 Hour",v:withinHour.length,c:"var(--green)"},
          {l:"Compliance",v:compliance!==null?`${compliance}%`:"—",c:compliance===null?"var(--muted)":compliance>=80?"var(--green)":compliance>=60?"var(--yellow)":"var(--red)"},
          {l:"Median Time",v:med?`${med} min`:"—",c:"var(--yellow)"},
          {l:"Avg Time",v:avgMin?`${avgMin} min`:"—",c:"var(--orange)"},
        ].map(s => (
          <div key={s.l} className="rcard">
            <div className="rv" style={{color:s.c}}>{s.v}</div>
            <div className="rl">{s.l}</div>
          </div>
        ))}
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
        <div>
          <div className="stitle">Delay Reasons (All Time)</div>
          <div className="twrap">
            <table>
              <thead><tr><th>Reason</th><th>Count</th><th>%</th></tr></thead>
              <tbody>
                {Object.entries(delayCounts).filter(([,v])=>v>0).length===0
                  ? <tr><td colSpan={3} style={{color:"var(--muted)",textAlign:"center",padding:20}}>No delays recorded</td></tr>
                  : Object.entries(delayCounts).sort((a,b)=>b[1]-a[1]).map(([r,v]) => (
                    <tr key={r}>
                      <td>{r}</td>
                      <td style={{color:"var(--orange)",fontWeight:700}}>{v}</td>
                      <td style={{color:"var(--muted)"}}>{cases.length?Math.round(v/cases.length*100):0}%</td>
                    </tr>
                  ))
                }
              </tbody>
            </table>
          </div>
        </div>
        <div>
          <div className="stitle">Infection Sources (All Time)</div>
          <div className="twrap">
            <table>
              <thead><tr><th>Source</th><th>Count</th><th>%</th></tr></thead>
              <tbody>
                {Object.entries(bySrc).sort((a,b)=>b[1]-a[1]).map(([s,v]) => (
                  <tr key={s}>
                    <td>{s}</td>
                    <td style={{color:"var(--accent)",fontWeight:700}}>{v}</td>
                    <td style={{color:"var(--muted)"}}>{cases.length?Math.round(v/cases.length*100):0}%</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// Export Confirmation Toast
// ═══════════════════════════════════════════════════════════════════════════════
function ExportToast({ filename, fallback, onClose }) {
  useEffect(() => {
    const t = setTimeout(onClose, 6000);
    return () => clearTimeout(t);
  }, [onClose]);

  return (
    <div style={{
      position:"fixed", bottom:80, left:"50%", transform:"translateX(-50%)",
      zIndex:500, minWidth:340, maxWidth:480,
      background:"#0d2b1a", border:"2px solid var(--green)",
      borderRadius:10, padding:"14px 18px",
      boxShadow:"0 8px 32px rgba(0,0,0,.7)",
      animation:"tin .3s ease",
      display:"flex", alignItems:"flex-start", gap:12,
    }}>
      <span style={{fontSize:"1.4rem",lineHeight:1}}>✅</span>
      <div style={{flex:1}}>
        <div style={{fontSize:".8rem",fontWeight:700,color:"var(--green)",marginBottom:3}}>
          Export completed
        </div>
        <div style={{fontSize:".72rem",color:"var(--text)",marginBottom:2}}>
          <strong>{filename}</strong>
        </div>
        {fallback ? (
          <div style={{fontSize:".68rem",color:"var(--yellow)",marginTop:4,lineHeight:1.5}}>
            ⚠ Your browser restricted automatic download in this preview environment.
            The file was prepared but could not save automatically.
            Copy this app to a real browser tab for full download support.
          </div>
        ) : (
          <div style={{fontSize:".68rem",color:"var(--muted)",marginTop:2}}>
            Please check your <strong style={{color:"var(--text)"}}>Downloads folder</strong>.
          </div>
        )}
      </div>
      <button onClick={onClose} style={{background:"none",border:"none",color:"var(--muted)",cursor:"pointer",fontSize:"1.1rem",lineHeight:1,padding:0}}>×</button>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// Prefix Input Component  — keeps "HN-" or "ER-" always visible
// ═══════════════════════════════════════════════════════════════════════════════
function PrefixInput({ prefix, value, onChange, placeholder, digitOnly, padLength, style }) {
  // value stored without prefix; we display prefix+value
  const handleChange = (e) => {
    let v = e.target.value;
    if (digitOnly) v = v.replace(/\D/g, "");          // strip non-digits
    if (padLength)  v = v.slice(0, padLength + 2);    // limit length
    onChange(v);
  };
  const handleBlur = (e) => {
    if (padLength && value) {
      onChange(value.padStart(padLength, "0"));
    }
  };
  return (
    <div style={{display:"flex",alignItems:"stretch",borderRadius:6,overflow:"hidden",border:"1px solid var(--border)",transition:"border-color .18s",...style}}
      onFocusCapture={e => e.currentTarget.style.borderColor="var(--accent)"}
      onBlurCapture={e  => e.currentTarget.style.borderColor="var(--border)"}
    >
      <span style={{
        background:"var(--border2)",color:"var(--accent)",fontWeight:700,
        fontSize:".78rem",padding:"7px 9px",whiteSpace:"nowrap",
        borderRight:"1px solid var(--border)",userSelect:"none",lineHeight:1.4,
      }}>{prefix}</span>
      <input
        value={value}
        onChange={handleChange}
        onBlur={handleBlur}
        placeholder={placeholder}
        inputMode={digitOnly?"numeric":"text"}
        style={{
          flex:1,background:"var(--bg)",border:"none",outline:"none",
          color:"var(--text)",fontSize:".78rem",padding:"7px 9px",
          fontFamily:"var(--mono)",minWidth:0,
        }}
      />
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// Patient Form
// ═══════════════════════════════════════════════════════════════════════════════
function PatientForm({ onSave, onClose }) {
  const defDT = () => new Date().toISOString().slice(0,16);
  // hn and bed store only the numeric suffix; prefix is added on save
  const [f, setF] = useState({
    hnNum:"", name:"", bedNum:"", physician:"",
    arrival:defDT(), recognition:defDT(),
    source:SOURCES[0], antibiotic:ANTIBIOTICS[0],
  });
  const set = (k,v) => setF(p=>({...p,[k]:v}));

  const [errors, setErrors] = useState({});

  const submit = () => {
    const errs = {};
    if (!f.hnNum.trim())   errs.hn  = "HN number is required";
    if (!f.name.trim())    errs.name= "Patient name is required";
    if (!f.bedNum.trim())  errs.bed = "Bed number is required";
    if (Object.keys(errs).length) { setErrors(errs); return; }

    // Format: HN-XXXXXX (pad to 6 digits), ER-XX (pad to 2 digits)
    const hn  = "HN-" + f.hnNum.padStart(6,"0");
    const bed = "ER-" + f.bedNum.padStart(2,"0");
    onSave({ ...f, hn, bed });
  };

  const field = (k) => ({
    style: errors[k] ? {border:"1px solid var(--red)",borderRadius:6} : {},
  });

  return (
    <div className="moverlay" onClick={onClose}>
      <div className="modal" onClick={e=>e.stopPropagation()}>
        <div className="mtitle">🔴 Register New Sepsis Patient</div>
        <div className="fgrid">
          {/* HN with persistent prefix */}
          <div className="fg">
            <label>HN / Medical Record No. *</label>
            <PrefixInput
              prefix="HN-"
              value={f.hnNum}
              onChange={v => { set("hnNum",v); setErrors(p=>({...p,hn:""})); }}
              placeholder="005678"
              digitOnly={true}
              padLength={6}
              style={errors.hn?{borderColor:"var(--red)"}:{}}
            />
            {errors.hn && <span style={{fontSize:".62rem",color:"var(--red)",marginTop:2}}>{errors.hn}</span>}
            {f.hnNum && <span style={{fontSize:".6rem",color:"var(--muted)",marginTop:2}}>→ will save as <strong style={{color:"var(--green)"}}>HN-{f.hnNum.padStart(6,"0")}</strong></span>}
          </div>

          {/* Patient name */}
          <div className="fg">
            <label>Patient Name *</label>
            <input value={f.name} onChange={e=>{ set("name",e.target.value); setErrors(p=>({...p,name:""})); }}
              placeholder="Full name"
              style={errors.name?{border:"1px solid var(--red)"}:{}}/>
            {errors.name && <span style={{fontSize:".62rem",color:"var(--red)",marginTop:2}}>{errors.name}</span>}
          </div>

          {/* ER Bed with persistent prefix */}
          <div className="fg">
            <label>ER Bed Number *</label>
            <PrefixInput
              prefix="ER-"
              value={f.bedNum}
              onChange={v => { set("bedNum",v); setErrors(p=>({...p,bed:""})); }}
              placeholder="07"
              digitOnly={true}
              padLength={2}
              style={errors.bed?{borderColor:"var(--red)"}:{}}
            />
            {errors.bed && <span style={{fontSize:".62rem",color:"var(--red)",marginTop:2}}>{errors.bed}</span>}
            {f.bedNum && <span style={{fontSize:".6rem",color:"var(--muted)",marginTop:2}}>→ will save as <strong style={{color:"var(--green)"}}>ER-{f.bedNum.padStart(2,"0")}</strong></span>}
          </div>

          {/* Physician */}
          <div className="fg">
            <label>Ordering Physician</label>
            <input value={f.physician} onChange={e=>set("physician",e.target.value)} placeholder="Dr. Name"/>
          </div>

          <div className="fg">
            <label>Arrival Time</label>
            <input type="datetime-local" value={f.arrival} onChange={e=>set("arrival",e.target.value)}/>
          </div>
          <div className="fg">
            <label>⏱ Sepsis Recognition Time</label>
            <input type="datetime-local" value={f.recognition} onChange={e=>set("recognition",e.target.value)}/>
          </div>

          <div className="fg">
            <label>Suspected Infection Source</label>
            <select value={f.source} onChange={e=>set("source",e.target.value)}>
              {SOURCES.map(s=><option key={s}>{s}</option>)}
            </select>
          </div>
          <div className="fg">
            <label>Planned Antibiotic</label>
            <select value={f.antibiotic} onChange={e=>set("antibiotic",e.target.value)}>
              {ANTIBIOTICS.map(a=><option key={a}>{a}</option>)}
            </select>
          </div>
        </div>
        <div className="mactions">
          <button className="btn btn-g" onClick={onClose}>Cancel</button>
          <button className="btn btn-s" onClick={submit}>✅ Register &amp; Start Timer</button>
        </div>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// Detail Modal (read-only for completed, editable for active)
// ═══════════════════════════════════════════════════════════════════════════════
function DetailModal({ c, tick, onAdvance, onUpdateDelay, onClose }) {
  const done  = c.Status === "completed";
  const el    = elapsedSec(c.SepsisRecognitionTime);
  const u     = urgencyOf(el, done);
  const col   = U_COLOR[u];
  const step  = currentStepKey(c._steps||{});

  const [delay, setDelay]   = useState(c.DelayReason||"");
  const [other, setOther]   = useState(c.OtherDelayReasonDetail||"");
  const [saved, setSaved]   = useState(false);

  const saveDelay = () => {
    onUpdateDelay(c.CaseID, delay, other);
    setSaved(true); setTimeout(()=>setSaved(false),1500);
  };

  return (
    <div className="moverlay" onClick={onClose}>
      <div className="modal modal-lg" onClick={e=>e.stopPropagation()}>
        <div className="mtitle" style={{color:col}}>
          {done?"✅":"🔴"} {c.PatientName} — {c.BedNumber}
          <span style={{marginLeft:10,fontSize:".7rem",color:"var(--muted)",fontWeight:400}}>{c.CaseID}</span>
        </div>

        <div className="fgrid" style={{marginBottom:14}}>
          {[["HN / MRN",c.HN],["Bed",c.BedNumber],["Physician",c.OrderingPhysician],["Source",c.SuspectedSource],["Antibiotic",c.PlannedAntibiotic],["Status",c.Status.toUpperCase()]].map(([l,v])=>(
            <div key={l} className="fg">
              <label>{l}</label>
              <div style={{fontSize:".78rem",padding:"5px 0",color:"var(--text)"}}>{v||"—"}</div>
            </div>
          ))}
        </div>

        <div className="stitle">Workflow Timeline</div>
        <div style={{display:"flex",flexDirection:"column",gap:7,marginBottom:14}}>
          {STATUS_STEPS.map((s,i)=>{
            const ts    = c._steps?.[s.key];
            const isDone= !!ts;
            const isNext= !done && s.key===STATUS_STEPS[STATUS_STEPS.findIndex(x=>x.key===step)+1]?.key;
            return (
              <div key={s.key} style={{display:"flex",alignItems:"center",gap:10,padding:"9px 12px",background:isDone?"rgba(0,230,118,.06)":"var(--bg)",borderRadius:6,border:`1px solid ${isDone?"var(--green)":"var(--border)"}`}}>
                <span style={{fontSize:"1.2rem"}}>{isDone?"✅":s.icon}</span>
                <div style={{flex:1}}>
                  <div style={{fontSize:".78rem",fontWeight:600}}>{s.label}</div>
                  <div style={{fontSize:".65rem",color:"var(--muted)"}}>{isDone?fmtDT(ts):"Pending"}</div>
                </div>
                {isNext && <button className="btn btn-p" style={{flex:"0 0 auto",padding:"5px 12px"}} onClick={()=>onAdvance(c.CaseID)}>Mark Done ▶</button>}
              </div>
            );
          })}
        </div>

        {done ? (
          <div style={{padding:14,background:"rgba(0,230,118,.07)",border:"1px solid var(--green)",borderRadius:8,marginBottom:14}}>
            <div style={{fontSize:".62rem",color:"var(--green)",textTransform:"uppercase",letterSpacing:1,marginBottom:5}}>Case Summary</div>
            <div style={{fontFamily:"var(--head)",fontSize:"1.5rem",fontWeight:800,color:c.WithinOneHour==="Yes"?"var(--green)":"var(--red)"}}>{c.TotalTimeToATBMinutes} minutes</div>
            <div style={{fontSize:".72rem",color:"var(--muted)",marginTop:3}}>
              {c.WithinOneHour==="Yes"?"✅ Within 1-hour target":"❌ Target missed — delayed administration"}
            </div>
            {c.DelayReason && <div style={{fontSize:".72rem",color:"var(--orange)",marginTop:4}}>Delay reason: {c.DelayReason}{c.OtherDelayReasonDetail?" — "+c.OtherDelayReasonDetail:""}</div>}
          </div>
        ) : (
          <div>
            <div className="stitle">Delay Reason</div>
            <select value={delay} onChange={e=>{setDelay(e.target.value);setSaved(false);}}
              style={{width:"100%",marginBottom:6,background:"var(--bg)",border:"1px solid var(--border)",color:"var(--text)",fontSize:".78rem",padding:"7px 9px",borderRadius:6}}>
              <option value="">None</option>
              {DELAY_REASONS.map(r=><option key={r}>{r}</option>)}
            </select>
            {delay==="Other" && (
              <input value={other} onChange={e=>{setOther(e.target.value);setSaved(false);}} placeholder="Please specify other delay reason *" required
                style={{width:"100%",marginBottom:6,background:"var(--bg)",border:"1px solid var(--orange)",color:"var(--text)",fontSize:".78rem",padding:"7px 9px",borderRadius:6}}/>
            )}
            <button className="btn btn-o" style={{flex:"0 0 auto",padding:"6px 14px"}} onClick={saveDelay}>
              {saved?"✅ Saved!":"💾 Save Delay Reason"}
            </button>
          </div>
        )}

        <div className="mactions" style={{marginTop:10}}>
          <button className="btn btn-g" onClick={onClose}>Close</button>
        </div>
      </div>
    </div>
  );
}


// ═══════════════════════════════════════════════════════════════════════════════
// AdminToggle — PIN-protected admin mode
// ═══════════════════════════════════════════════════════════════════════════════
function AdminToggle({ isAdmin, onToggle }) {
  const [showPin, setShowPin] = useState(false);
  const [pin, setPin]         = useState("");
  const [err, setErr]         = useState(false);

  const tryLogin = () => {
    if (pin === ADMIN_PIN) { onToggle(true); setShowPin(false); setPin(""); setErr(false); }
    else { setErr(true); setPin(""); setTimeout(() => setErr(false), 2000); }
  };

  if (isAdmin) return (
    <div style={{display:"flex",alignItems:"center",gap:6}}>
      <span className="admin-badge">🔐 ADMIN</span>
      <button className="nb" style={{fontSize:".6rem",padding:"3px 8px"}} onClick={() => onToggle(false)}>Logout</button>
    </div>
  );

  return (
    <>
      <button className="nb" style={{fontSize:".65rem"}} onClick={() => setShowPin(true)}>🔐 Admin</button>
      {showPin && (
        <div className="moverlay" onClick={() => { setShowPin(false); setPin(""); }}>
          <div className="modal" style={{maxWidth:340,textAlign:"center"}} onClick={e=>e.stopPropagation()}>
            <div style={{fontSize:"2rem",marginBottom:8}}>🔐</div>
            <div className="mtitle" style={{textAlign:"center",marginBottom:6,fontSize:"1rem"}}>Admin Login</div>
            <div style={{fontSize:".72rem",color:"var(--muted)",marginBottom:14}}>Enter admin PIN to enable delete and audit features.</div>
            <input
              className="pin-input"
              type="password"
              inputMode="numeric"
              maxLength={6}
              placeholder="····"
              value={pin}
              onChange={e => { setPin(e.target.value.replace(/\D/g,"")); setErr(false); }}
              onKeyDown={e => e.key==="Enter" && tryLogin()}
              autoFocus
            />
            {err && <div style={{color:"var(--red)",fontSize:".7rem",marginTop:8}}>❌ Incorrect PIN</div>}
            <div style={{display:"flex",gap:8,justifyContent:"center",marginTop:16}}>
              <button className="btn btn-g" style={{flex:"0 0 auto",padding:"6px 16px"}} onClick={() => { setShowPin(false); setPin(""); }}>Cancel</button>
              <button className="btn btn-p" style={{flex:"0 0 auto",padding:"6px 20px"}} onClick={tryLogin}>Login</button>
            </div>
          </div>
        </div>
      )}
    </>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// DeleteModal — confirm + reason before soft delete
// ═══════════════════════════════════════════════════════════════════════════════
const DELETE_REASONS = [
  "Demo / sample case",
  "Wrong patient entry",
  "Duplicate case",
  "Test entry",
  "Other",
];

function DeleteModal({ c, isAdmin, onConfirm, onClose }) {
  const [reason, setReason] = useState("");
  const [other,  setOther]  = useState("");
  const [by,     setBy]     = useState("");
  const [step,   setStep]   = useState(1); // 1=confirm, 2=reason
  const [errors, setErrors] = useState({});

  if (!isAdmin) return (
    <div className="moverlay" onClick={onClose}>
      <div className="modal" style={{maxWidth:380,textAlign:"center"}} onClick={e=>e.stopPropagation()}>
        <div style={{fontSize:"2rem",marginBottom:10}}>🚫</div>
        <div className="mtitle" style={{textAlign:"center"}}>Admin Access Required</div>
        <div style={{fontSize:".78rem",color:"var(--muted)",margin:"10px 0 18px"}}>You must be logged in as Admin to delete cases.</div>
        <button className="btn btn-g" style={{padding:"7px 20px"}} onClick={onClose}>Close</button>
      </div>
    </div>
  );

  const proceed = () => {
    const errs = {};
    if (!reason) errs.reason = "Please select a reason";
    if (reason === "Other" && !other.trim()) errs.other = "Please specify the reason";
    if (!by.trim()) errs.by = "Please enter your name";
    if (Object.keys(errs).length) { setErrors(errs); return; }
    onConfirm(reason, reason === "Other" ? other : "", by);
  };

  return (
    <div className="moverlay" onClick={onClose}>
      <div className="modal" style={{maxWidth:480}} onClick={e=>e.stopPropagation()}>
        {step === 1 ? (
          <>
            <div style={{fontSize:"2rem",marginBottom:10,textAlign:"center"}}>⚠️</div>
            <div className="mtitle" style={{textAlign:"center",color:"var(--red)"}}>Delete Sepsis Case</div>
            <div style={{background:"rgba(245,0,87,.06)",border:"1px solid rgba(245,0,87,.25)",borderRadius:8,padding:14,margin:"12px 0",fontSize:".78rem",lineHeight:1.6}}>
              <div style={{fontWeight:700,marginBottom:6,color:"var(--text)"}}>Are you sure you want to delete this case?</div>
              <div style={{color:"var(--muted)"}}>This action <strong style={{color:"var(--red)"}}>cannot be undone</strong>. The case will be hidden from all views but retained in the audit log for compliance purposes.</div>
            </div>
            <div style={{background:"var(--bg)",border:"1px solid var(--border)",borderRadius:7,padding:12,fontSize:".75rem"}}>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:6}}>
                {[["Case ID",c.CaseID],["Patient",c.PatientName],["HN",c.HN],["Bed",c.BedNumber],["Status",c.Status?.toUpperCase()]].map(([l,v])=>(
                  <div key={l}><span style={{color:"var(--muted)"}}>{l}: </span><strong>{v}</strong></div>
                ))}
              </div>
            </div>
            <div className="mactions">
              <button className="btn btn-g" onClick={onClose}>Cancel</button>
              <button className="btn btn-d" onClick={() => setStep(2)}>Continue →</button>
            </div>
          </>
        ) : (
          <>
            <div className="mtitle" style={{color:"var(--red)"}}>🗑 Confirm Deletion</div>
            <div style={{display:"flex",flexDirection:"column",gap:12}}>
              <div className="fg">
                <label style={{fontSize:".65rem",textTransform:"uppercase",letterSpacing:1,color:"var(--muted)",display:"block",marginBottom:4}}>Delete Reason *</label>
                <select value={reason} onChange={e=>{setReason(e.target.value);setErrors(p=>({...p,reason:""}));}}
                  style={{width:"100%",background:"var(--bg)",border:`1px solid ${errors.reason?"var(--red)":"var(--border)"}`,color:"var(--text)",fontSize:".78rem",padding:"7px 9px",borderRadius:6}}>
                  <option value="">Select reason…</option>
                  {DELETE_REASONS.map(r=><option key={r}>{r}</option>)}
                </select>
                {errors.reason && <span style={{fontSize:".62rem",color:"var(--red)"}}>{errors.reason}</span>}
              </div>

              {reason === "Other" && (
                <div className="fg">
                  <label style={{fontSize:".65rem",textTransform:"uppercase",letterSpacing:1,color:"var(--muted)",display:"block",marginBottom:4}}>Please specify other reason *</label>
                  <input value={other} onChange={e=>{setOther(e.target.value);setErrors(p=>({...p,other:""}));}}
                    placeholder="Enter reason…"
                    style={{width:"100%",background:"var(--bg)",border:`1px solid ${errors.other?"var(--red)":"var(--border)"}`,color:"var(--text)",fontSize:".78rem",padding:"7px 9px",borderRadius:6}}/>
                  {errors.other && <span style={{fontSize:".62rem",color:"var(--red)"}}>{errors.other}</span>}
                </div>
              )}

              <div className="fg">
                <label style={{fontSize:".65rem",textTransform:"uppercase",letterSpacing:1,color:"var(--muted)",display:"block",marginBottom:4}}>Deleted By (your name) *</label>
                <input value={by} onChange={e=>{setBy(e.target.value);setErrors(p=>({...p,by:""}));}}
                  placeholder="e.g. Nurse Apinya"
                  style={{width:"100%",background:"var(--bg)",border:`1px solid ${errors.by?"var(--red)":"var(--border)"}`,color:"var(--text)",fontSize:".78rem",padding:"7px 9px",borderRadius:6}}/>
                {errors.by && <span style={{fontSize:".62rem",color:"var(--red)"}}>{errors.by}</span>}
              </div>

              <div style={{background:"rgba(245,0,87,.06)",border:"1px solid rgba(245,0,87,.2)",borderRadius:6,padding:"8px 12px",fontSize:".7rem",color:"var(--muted)"}}>
                🔒 This deletion will be logged in the Audit Trail with your name, timestamp, and reason.
              </div>
            </div>
            <div className="mactions">
              <button className="btn btn-g" onClick={() => setStep(1)}>← Back</button>
              <button className="btn btn-d" onClick={proceed}>🗑 Confirm Delete</button>
            </div>
          </>
        )}
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// Audit View
// ═══════════════════════════════════════════════════════════════════════════════
function AuditView({ log, deletedCases, exportCSV, exportXLSX }) {
  const today = new Date().toISOString().slice(0,10);

  const auditExportCols = ["AuditID","CaseID","HN","BedNumber","PatientName","DeletedBy","DeletedAt","DeleteReason","DeleteReasonDetail","ActionType"];
  const auditRows = log.map(e => auditExportCols.map(col => {
    const v = e[col];
    if (!v) return "";
    if (col === "DeletedAt") return fmtDT(v);
    return String(v);
  }));

  const exportAuditXLSX = () => {
    const data = [auditExportCols, ...auditRows];
    const ws = XLSX.utils.aoa_to_sheet(data);
    ws["!cols"] = auditExportCols.map(h => ({ wch: Math.max(h.length+2,16) }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "AuditLog");
    const arr = XLSX.write(wb, { bookType:"xlsx", type:"array" });
    const blob = new Blob([arr], { type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a"); a.href=url; a.download=`SepsisTrack_AuditLog_${today}.xlsx`;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 2000);
  };

  return (
    <div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12,flexWrap:"wrap",gap:8}}>
        <div className="stitle" style={{margin:0}}>🔐 AUDIT LOG — DELETED CASES</div>
        <div style={{display:"flex",gap:6}}>
          <button className="xbtn xl" onClick={exportAuditXLSX}>📥 Export Audit Log</button>
        </div>
      </div>

      {deletedCases.length > 0 && (
        <div style={{background:"rgba(245,0,87,.05)",border:"1px solid rgba(245,0,87,.2)",borderRadius:8,padding:14,marginBottom:14}}>
          <div style={{fontSize:".68rem",color:"var(--red)",fontWeight:700,marginBottom:8,textTransform:"uppercase",letterSpacing:1}}>
            🗑 Soft-Deleted Cases ({deletedCases.length}) — Hidden from views, retained for audit
          </div>
          <div className="twrap">
            <table>
              <thead><tr><th>Case ID</th><th>HN</th><th>Patient</th><th>Bed</th><th>Deleted By</th><th>Deleted At</th><th>Reason</th></tr></thead>
              <tbody>
                {deletedCases.map(c => (
                  <tr key={c.CaseID}>
                    <td style={{color:"var(--muted)",fontSize:".65rem"}}>{c.CaseID}</td>
                    <td style={{color:"var(--muted)"}}>{c.HN}</td>
                    <td style={{textDecoration:"line-through",opacity:.6}}>{c.PatientName}</td>
                    <td><span className="bedtag" style={{opacity:.5,fontSize:".62rem"}}>{c.BedNumber}</span></td>
                    <td style={{color:"var(--orange)",fontSize:".7rem"}}>{c.DeletedBy||"—"}</td>
                    <td style={{color:"var(--muted)",fontSize:".68rem"}}>{fmtDT(c.DeletedAt)}</td>
                    <td style={{color:"var(--red)",fontSize:".68rem"}}>{c.DeleteReason||"—"}{c.DeleteReasonDetail?" — "+c.DeleteReasonDetail:""}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      <div className="stitle">Deletion Actions Log ({log.length} entries)</div>
      {log.length === 0 ? (
        <div className="empty"><div className="eicon">📋</div><div>No audit entries yet. Delete actions will appear here.</div></div>
      ) : (
        <div className="twrap">
          <table>
            <thead>
              <tr><th>Audit ID</th><th>Case ID</th><th>HN</th><th>Patient</th><th>Bed</th><th>Action</th><th>Deleted By</th><th>Deleted At</th><th>Reason</th></tr>
            </thead>
            <tbody>
              {[...log].reverse().map(e => (
                <tr key={e.AuditID}>
                  <td style={{color:"var(--muted)",fontSize:".6rem"}}>{e.AuditID}</td>
                  <td style={{color:"var(--muted)",fontSize:".65rem"}}>{e.CaseID}</td>
                  <td style={{color:"var(--muted)"}}>{e.HN}</td>
                  <td>{e.PatientName}</td>
                  <td><span className="bedtag" style={{fontSize:".62rem"}}>{e.BedNumber}</span></td>
                  <td><span className="badge-sm" style={{background:"rgba(245,0,87,.1)",color:"var(--red)",border:"1px solid rgba(245,0,87,.3)"}}>{e.ActionType}</span></td>
                  <td style={{color:"var(--orange)"}}>{e.DeletedBy}</td>
                  <td style={{color:"var(--muted)",fontSize:".7rem"}}>{fmtDT(e.DeletedAt)}</td>
                  <td style={{color:"var(--red)",fontSize:".7rem"}}>{e.DeleteReason}{e.DeleteReasonDetail?" — "+e.DeleteReasonDetail:""}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}
