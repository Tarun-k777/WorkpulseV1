// ==UserScript==
// @name         WorkPulse — Associate
// @namespace    https://amazon.sharepoint.com/sites/teamdailytask/
// @version      3.1.0
// @description  WorkPulse — Associate productivity + attendance tracker
// @author       Your Team
// @match        https://amazon.sharepoint.com/sites/teamdailytask/*
// @match        https://amazon.sharepoint.com/*
// @match        https://www.grainger.com/*
// @match        https://*.grainger.com/*
// @grant        GM_xmlhttpRequest
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        GM_addStyle
// @connect      amazon.sharepoint.com
// @connect      *.sharepoint.com
// @connect      *
// @run-at       document-idle
// ==/UserScript==

(function () {
  'use strict';

  // ═══════════════════════════════════════════════════════════════════════
  // CONFIG — ONE SharePoint for everything
  // ═══════════════════════════════════════════════════════════════════════
  const SP = {
    SITE:        'https://amazon.sharepoint.com/sites/teamdailytask',
    TASK_LIST:   'DailyTaskReport',      // existing task list
    STATUS_LIST: 'AttendanceStatus',     // NEW: WFO/WFH/Leave tracking
    NPT_LIST:    'NPTLog',               // NEW: Missed NPT log
  };

  // ═══════════════════════════════════════════════════════════════════════
  // CONSTANTS
  // ═══════════════════════════════════════════════════════════════════════
  const WH = 8;
  const TASK_TYPES = [
    'CT Production','Pre-Prod Production','Simulator Production',
    'CT Audits','Pre-Prod Audits','Simulator Audits','Lack of Work',
    'Meeting','Ad-hoc Tasks','Other'
  ];
  const NPT_TASKS  = ['Lack of Work'];
  const STATUSES   = ['WFO','WFH','SL','CL','AL','Optional Off'];
  const PROCS      = ['Preprod Testing','Chat Transcripts','Adhoc','Quality Check','Training','Other'];
  const COLORS     = ['#6366f1','#8b5cf6','#06b6d4','#10b981','#f59e0b','#ef4444','#ec4899','#84cc16'];
  const STATUS_CFG = {
    WFO: { label:'Work From Office', cls:'sp-wfo', color:'#3fb950', bg:'rgba(63,185,80,.14)'  },
    WFH: { label:'Work From Home',   cls:'sp-wfh', color:'#22d3ee', bg:'rgba(6,182,212,.14)'  },
    SL:  { label:'Sick Leave',       cls:'sp-sl',  color:'#f85149', bg:'rgba(248,81,73,.14)'  },
    CL:  { label:'Casual Leave',     cls:'sp-cl',  color:'#d29922', bg:'rgba(210,153,34,.14)' },
    AL:  { label:'Annual Leave',     cls:'sp-al',  color:'#a371f7', bg:'rgba(163,113,247,.14)'},
    'Optional Off':{ label:'Optional Off', cls:'sp-oo', color:'#6b7280', bg:'rgba(107,114,128,.14)'},
  };

  // ═══════════════════════════════════════════════════════════════════════
  // STATE HELPERS
  // ═══════════════════════════════════════════════════════════════════════
  function safeLoad(k,fb){try{const v=GM_getValue(k,null);if(v===null)return fb;if(Array.isArray(v))return v;const p=JSON.parse(v);return Array.isArray(p)?p:fb;}catch(e){return fb;}}
  function safeSave(k,v){try{GM_setValue(k,Array.isArray(v)?v:[]);}catch(e){}}
  function safeLoadObj(k,fb){try{const v=GM_getValue(k,'');if(!v)return fb;const p=JSON.parse(v);return(p&&typeof p==='object'&&!Array.isArray(p))?p:fb;}catch(e){return fb;}}
  function safeSaveObj(k,v){try{GM_setValue(k,JSON.stringify(v));}catch(e){}}

  // Session
  function generateSID(){return Math.random().toString(36).slice(2)+Date.now().toString(36);}
  function loadSession(){try{const r=GM_getValue('dtr_sess2','');if(!r)return null;const s=JSON.parse(r);return(s&&s.name&&s.sid)?s:null;}catch(e){return null;}}
  function saveSession(name){const s={name,sid:generateSID(),at:new Date().toISOString()};GM_setValue('dtr_sess2',JSON.stringify(s));return s;}
  function clearSession(){GM_setValue('dtr_sess2','');}

  // App state
  let submissions  = safeLoad('dtr_subs4',[]);
  let excelRows    = safeLoad('dtr_excel4',[]);
  let spQueue      = safeLoad('dtr_spq2',[]);
  let statusCache  = safeLoadObj('dtr_status2',{});   // { "YYYY-MM-DD": { status, process, task } }
  let nptCache      = safeLoad('dtr_npt2',[]);
  let tkCustomCols  = safeLoad('dtr_tkcols4',[]);
  let teamStatusCache = safeLoadObj('dtr_teamcache',{});
  let theme        = GM_getValue('dtr_theme4','dark');

  let currentSession = loadSession();
  let authState = currentSession
    ? {loggedIn:true,name:currentSession.name,sid:currentSession.sid}
    : {loggedIn:false,name:'',sid:''};

  let currentView  = '';
  let taskCounter  = 1;
  let selectedDate = todayStr();
  let weekOffset   = 0;   // for My Week View
  let root;

  // ═══════════════════════════════════════════════════════════════════════
  // STYLES
  // ═══════════════════════════════════════════════════════════════════════
  GM_addStyle(`
    @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700&family=DM+Mono:wght@400;500&display=swap');
    *{box-sizing:border-box;margin:0;padding:0}
    :root{
      --bg:#0d1117;--bg2:#161b22;--bg3:#1c2333;--bg4:#21262d;
      --border:rgba(255,255,255,0.08);--border2:rgba(255,255,255,0.14);
      --text:#e6edf3;--text2:#8b949e;--text3:#484f58;
      --accent:#58a6ff;--accent2:#a371f7;
      --green:#3fb950;--amber:#d29922;--red:#f85149;--purple:#a371f7;
      --radius:10px;--font:'DM Sans',sans-serif;--mono:'DM Mono',monospace
    }
    .lt{
      --bg:#f6f8fa;--bg2:#ffffff;--bg3:#f0f3f6;--bg4:#e1e5ea;
      --border:rgba(0,0,0,0.09);--border2:rgba(0,0,0,0.15);
      --text:#1f2328;--text2:#656d76;--text3:#9198a1;
      --accent:#0969da;--accent2:#8250df;--green:#1a7f37;--amber:#9a6700;--red:#d1242f
    }

    /* ── Root shell ── */
    #dtr-root-outer{position:fixed;inset:0;z-index:2147483647;overflow:hidden;background:var(--bg)}#dtr-root{position:absolute;top:0;left:0;width:90.9%;height:90.9%;transform:scale(1.1);transform-origin:top left;font-family:var(--font);font-size:1rem;background:var(--bg);color:var(--text);display:flex;flex-direction:column;overflow:hidden}

    /* ── Top bar ── */
    #dtr-topbar{height:56px;flex-shrink:0;background:var(--bg2);border-bottom:1px solid var(--border);display:flex;align-items:center;padding:0 20px;gap:8px}
    .dtr-logo{display:flex;align-items:center;gap:8px;padding-right:14px;border-right:1px solid var(--border);margin-right:6px;flex-shrink:0}
    .dtr-logo-icon{width:32px;height:32px;border-radius:8px;background:linear-gradient(135deg,#58a6ff,#a371f7);display:flex;align-items:center;justify-content:center}
    .dtr-logo-icon svg{width:15px;height:15px}
    .dtr-logo-text{font-size:.965rem;font-weight:700;color:var(--text)}
    .dtr-tabs{display:flex;gap:2px;flex:1;overflow-x:auto;scrollbar-width:none}
    .dtr-tabs::-webkit-scrollbar{display:none}
    .dtr-tab{padding:5px 12px;border-radius:6px;border:none;background:transparent;color:var(--text2);font-family:var(--font);font-size:.86rem;font-weight:500;cursor:pointer;transition:all .15s;display:flex;align-items:center;gap:5px;white-space:nowrap;flex-shrink:0}
    .dtr-tab:hover{background:var(--bg3);color:var(--text)}
    .dtr-tab.active{background:var(--bg3);color:var(--accent);font-weight:600}
    .dtr-tab svg{width:13px;height:13px;flex-shrink:0}
    .dtr-topbar-r{margin-left:auto;display:flex;align-items:center;gap:7px;flex-shrink:0}
    .dtr-user-pill{display:flex;align-items:center;gap:6px;padding:4px 10px;border-radius:20px;background:var(--bg3);border:1px solid var(--border);font-size:.85rem;color:var(--text2)}
    .dtr-user-pill .av{width:18px;height:18px;border-radius:50%;background:linear-gradient(135deg,var(--accent),var(--accent2));font-size:9px;font-weight:700;color:#fff;display:flex;align-items:center;justify-content:center}
    .role-badge{padding:2px 7px;border-radius:4px;font-size:.77rem;font-weight:700;text-transform:uppercase;letter-spacing:.5px;background:rgba(88,166,255,.12);color:var(--accent)}
    .icon-btn{width:30px;height:30px;border-radius:6px;border:1px solid var(--border);background:var(--bg3);color:var(--text2);cursor:pointer;display:flex;align-items:center;justify-content:center;transition:all .15s}
    .icon-btn:hover{background:var(--bg4);color:var(--text)}
    .icon-btn.danger:hover{background:rgba(248,81,73,.1);color:var(--red);border-color:rgba(248,81,73,.3)}
    .icon-btn svg{width:14px;height:14px}

    /* ── Body ── */
    #dtr-body{flex:1;overflow:hidden;display:flex}
    #dtr-sidebar{width:218px;flex-shrink:0;background:var(--bg2);border-right:1px solid var(--border);display:flex;flex-direction:column;padding:12px 10px;overflow-y:auto}
    .sb-sec{font-size:.7rem;font-weight:700;text-transform:uppercase;letter-spacing:1px;color:var(--text3);padding:8px 8px 4px;margin-top:6px}
    .sb-item{display:flex;align-items:center;gap:9px;padding:8px 10px;border-radius:7px;border:none;background:transparent;color:var(--text2);font-family:var(--font);font-size:.875rem;font-weight:500;cursor:pointer;transition:all .15s;width:100%;text-align:left}
    .sb-item:hover{background:var(--bg3);color:var(--text)}
    .sb-item.active{background:rgba(88,166,255,.1);color:var(--accent);font-weight:600}
    .sb-item svg{width:14px;height:14px;flex-shrink:0}
    .sb-footer{margin-top:auto;padding:10px 8px;border-top:1px solid var(--border)}
    .sb-stat{margin-bottom:10px}
    .sb-stat .lbl{font-size:.75rem;color:var(--text3);margin-bottom:2px}
    .sb-stat .val{font-size:1.25rem;font-weight:700;color:var(--accent)}
    #dtr-main{flex:1;overflow-y:auto;padding:24px 28px;background:var(--bg)}
    #dtr-main::-webkit-scrollbar{width:4px}
    #dtr-main::-webkit-scrollbar-thumb{background:var(--bg4);border-radius:3px}

    /* ── Login card ── */
    #view-login{display:flex;align-items:center;justify-content:center;min-height:100%}
    .login-card{width:380px;background:var(--bg2);border:1px solid var(--border2);border-radius:14px;padding:32px}
    .login-logo{width:44px;height:44px;border-radius:12px;background:linear-gradient(135deg,#58a6ff,#a371f7);display:flex;align-items:center;justify-content:center;margin-bottom:16px}
    .login-logo svg{width:22px;height:22px}
    .login-title{font-size:1.2rem;font-weight:700;color:var(--text)}
    .login-sub{font-size:.9rem;color:var(--text2);margin-top:3px;margin-bottom:22px}
    .login-err{color:var(--red);font-size:.875rem;margin-bottom:10px;min-height:18px}

    /* ── Forms ── */
    .dtr-field{display:flex;flex-direction:column;gap:5px;margin-bottom:12px}
    .dtr-label{font-size:.82rem;font-weight:600;color:var(--text2);letter-spacing:.2px}
    .dtr-input,.dtr-select,.dtr-textarea{padding:8px 11px;border-radius:var(--radius);background:var(--bg3);border:1px solid var(--border);color:var(--text);font-family:var(--font);font-size:.96rem;transition:border-color .15s,box-shadow .15s;width:100%;appearance:none}
    .dtr-input:focus,.dtr-select:focus,.dtr-textarea:focus{outline:none;border-color:var(--accent);box-shadow:0 0 0 3px rgba(88,166,255,.15)}
    .dtr-input::placeholder{color:var(--text3)}
    .dtr-select option{background:var(--bg2)}
    .dtr-textarea{resize:vertical}

    /* ── Buttons ── */
    .btn{display:inline-flex;align-items:center;justify-content:center;gap:5px;padding:8px 16px;border-radius:var(--radius);font-family:var(--font);font-size:.94rem;font-weight:600;cursor:pointer;border:none;transition:all .15s;white-space:nowrap}
    .btn svg{width:13px;height:13px}
    .btn-primary{background:var(--accent);color:#fff}.btn-primary:hover{opacity:.88}
    .btn-ghost{background:var(--bg3);border:1px solid var(--border);color:var(--text2)}.btn-ghost:hover{background:var(--bg4);color:var(--text)}
    .btn-danger{background:rgba(248,81,73,.1);border:1px solid rgba(248,81,73,.25);color:var(--red)}.btn-danger:hover{background:rgba(248,81,73,.18)}
    .btn-full{width:100%}.btn-sm{padding:5px 11px;font-size:.875rem}.btn-xs{padding:3px 8px;font-size:.815rem}

    /* ── Page header ── */
    .ph{display:flex;align-items:flex-start;justify-content:space-between;margin-bottom:18px;gap:12px}
    .ph-left .ph-title{font-size:1.1rem;font-weight:700;color:var(--text);letter-spacing:-.3px}
    .ph-left .ph-sub{font-size:.86rem;color:var(--text3);margin-top:2px}
    .ph-actions{display:flex;gap:6px;align-items:center;flex-shrink:0}

    /* ── Cards ── */
    .card{background:var(--bg2);border:1px solid var(--border);border-radius:12px;padding:16px;margin-bottom:14px}
    .card-title{font-size:.78rem;font-weight:700;text-transform:uppercase;letter-spacing:.7px;color:var(--text3);margin-bottom:12px;display:flex;align-items:center;gap:6px}

    /* ── Stat grid ── */
    .stats-grid{display:grid;gap:12px;margin-bottom:16px}
    .sg4{grid-template-columns:repeat(4,1fr)}.sg3{grid-template-columns:repeat(3,1fr)}.sg2{grid-template-columns:1fr 1fr}
    .stat-card{background:var(--bg2);border:1px solid var(--border);border-radius:12px;padding:14px}
    .stat-card .lbl{font-size:.77rem;color:var(--text3);text-transform:uppercase;letter-spacing:.5px;margin-bottom:6px}
    .stat-card .val{font-size:1.8rem;font-weight:800;color:var(--text);letter-spacing:-1px;line-height:1}
    .stat-card .sub{font-size:.8rem;color:var(--text3);margin-top:4px}
    .ab{border-top:2px solid var(--accent)}.gb{border-top:2px solid var(--green)}.amb{border-top:2px solid var(--amber)}.pb{border-top:2px solid var(--purple)}.rb2{border-top:2px solid var(--red)}

    /* ── Grid helpers ── */
    .g2{display:grid;grid-template-columns:1fr 1fr;gap:12px}
    .g3{display:grid;grid-template-columns:repeat(3,1fr);gap:12px}
    .ga{display:grid;grid-template-columns:2fr 1fr 1fr 1fr;gap:10px}

    /* ── Tables ── */
    .tbl-wrap{border:1px solid var(--border);border-radius:10px;overflow:hidden}
    .dtr-table{width:100%;border-collapse:collapse;font-size:.91rem}
    .dtr-table th{background:var(--bg3);padding:11px 14px;text-align:left;color:var(--text3);font-weight:600;font-size:.78rem;text-transform:uppercase;letter-spacing:.6px;border-bottom:1px solid var(--border)}
    .dtr-table td{padding:11px 14px;color:var(--text2);border-bottom:1px solid var(--border)}
    .dtr-table tr:last-child td{border-bottom:none}
    .dtr-table tr:hover td{background:var(--bg3)}
    .dtr-table .bold{font-weight:600;color:var(--text)}
    .dtr-table .mono{font-family:var(--mono)}

    /* ── Badges ── */
    .badge{display:inline-flex;align-items:center;gap:3px;padding:2px 7px;border-radius:4px;font-size:.77rem;font-weight:700}
    .bg2{background:rgba(63,185,80,.12);color:var(--green)}
    .ba{background:rgba(210,153,34,.12);color:var(--amber)}
    .bb{background:rgba(88,166,255,.12);color:var(--accent)}
    .br2{background:rgba(248,81,73,.12);color:var(--red)}

    /* ── Status pills ── */
    .sp{display:inline-flex;align-items:center;padding:3px 9px;border-radius:6px;font-size:.75rem;font-weight:700;white-space:nowrap}
    .sp-wfo{background:rgba(63,185,80,.14);color:#3fb950}
    .sp-wfh{background:rgba(6,182,212,.14);color:#22d3ee}
    .sp-sl{background:rgba(248,81,73,.14);color:#f85149}
    .sp-cl{background:rgba(210,153,34,.14);color:#d29922}
    .sp-al{background:rgba(163,113,247,.14);color:#a371f7}
    .sp-oo{background:rgba(107,114,128,.14);color:#9ca3af}
    .sp-ns{background:var(--bg3);color:var(--text3);border:1px dashed var(--border2)}
    .sp-we{background:transparent;color:var(--text3);font-style:italic;font-size:.7rem}

    /* ── Bar chart ── */
    .bar-rows{display:flex;flex-direction:column;gap:8px}
    .bar-row{display:flex;align-items:center;gap:10px}
    .bar-label{font-size:.855rem;color:var(--text2);width:120px;flex-shrink:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
    .bar-track{flex:1;height:20px;background:var(--bg4);border-radius:5px;overflow:hidden;position:relative}
    .bar-fill{height:100%;border-radius:5px;display:flex;align-items:center;padding-left:7px;transition:width .7s cubic-bezier(.34,1.56,.64,1)}
    .bar-fill-text{font-size:.79rem;font-weight:700;color:#fff}
    .bar-aside{font-size:.79rem;color:var(--text3);font-family:var(--mono);width:38px;text-align:right;flex-shrink:0}
    .chart-title{font-size:.78rem;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--text3);margin-bottom:12px}

    /* ── Task cards (Submit tab) ── */
    .task-card{background:var(--bg3);border:1px solid var(--border);border-radius:10px;padding:14px;margin-bottom:10px}
    .task-card.leave{border-color:rgba(210,153,34,.3);background:rgba(210,153,34,.04)}
    .tc-header{display:flex;align-items:center;justify-content:space-between;margin-bottom:12px}
    .tc-num{font-size:.77rem;font-weight:700;text-transform:uppercase;letter-spacing:.8px;color:var(--accent)}
    .npt-box{padding:8px 11px;border-radius:var(--radius);background:rgba(210,153,34,.1);border:1px solid rgba(210,153,34,.2);color:var(--amber);font-size:1rem;font-weight:700;font-family:var(--mono);text-align:center}
    .npt-box.prod{background:rgba(63,185,80,.1);border-color:rgba(63,185,80,.2);color:var(--green)}
    .leave-tag{display:none;align-items:center;gap:8px;padding:9px 12px;border-radius:8px;margin-top:10px;background:rgba(210,153,34,.08);border:1px solid rgba(210,153,34,.2);color:var(--amber);font-size:.875rem;flex-wrap:wrap}
    .leave-tag.show{display:flex}
    .leave-type-btn{padding:4px 11px;border-radius:6px;border:1px solid rgba(210,153,34,.35);background:transparent;color:var(--amber);font-family:var(--font);font-size:.78rem;font-weight:600;cursor:pointer;transition:all .15s}
    .leave-type-btn.active{background:rgba(210,153,34,.2);border-color:var(--amber)}
    .comments-box{display:none;margin-top:10px}
    .opt-note-toggle{display:flex;align-items:center;gap:5px;margin-top:8px;color:var(--accent);font-size:.82rem;font-weight:600;cursor:pointer;padding:4px 0;user-select:none}
    .opt-note-toggle:hover{color:var(--accent2)}
    .opt-note-box{margin-top:2px}
    .req-star{color:var(--red)}
    .add-task-btn{width:100%;padding:9px;border-radius:8px;border:1px dashed var(--border2);background:transparent;color:var(--accent);cursor:pointer;font-family:var(--font);font-size:.92rem;font-weight:600;transition:all .15s;margin-bottom:12px}
    .add-task-btn:hover{background:rgba(88,166,255,.05);border-color:var(--accent)}
    .date-pill-row{display:flex;gap:8px;margin-top:4px;flex-wrap:wrap}
    .date-pill{padding:5px 14px;border-radius:20px;border:1px solid var(--border2);background:var(--bg3);color:var(--text2);font-family:var(--font);font-size:.85rem;font-weight:600;cursor:pointer;transition:all .15s}
    .date-pill:hover{border-color:var(--accent);color:var(--accent)}
    .date-pill.active{background:rgba(88,166,255,.12);border-color:var(--accent);color:var(--accent)}
    .mode-toggle{display:flex;background:var(--bg3);border-radius:7px;padding:2px;border:1px solid var(--border)}
    .mode-btn{padding:5px 13px;border-radius:5px;border:none;background:transparent;color:var(--text3);font-family:var(--font);font-size:.875rem;font-weight:600;cursor:pointer;transition:all .15s}
    .mode-btn.active{background:var(--bg2);color:var(--accent);box-shadow:0 1px 5px rgba(0,0,0,.15)}
    .viz-card{background:var(--bg2);border:1px solid var(--border);border-radius:12px;padding:16px;margin-bottom:14px}
    .viz-title{font-size:.78rem;font-weight:700;text-transform:uppercase;letter-spacing:.7px;color:var(--text3);margin-bottom:12px;display:flex;align-items:center;justify-content:space-between}
    .viz-bar-wrap{margin-bottom:10px}
    .viz-bar-label{display:flex;align-items:center;justify-content:space-between;margin-bottom:4px;font-size:.8rem}
    .viz-track{width:100%;height:20px;background:var(--bg4);border-radius:7px;overflow:hidden}
    .viz-fill{height:100%;border-radius:7px;transition:width .5s cubic-bezier(.34,1.56,.64,1);display:flex;align-items:center;justify-content:center}
    .viz-fill span{font-size:.72rem;font-weight:700;color:#fff}
    .viz-legend{display:flex;gap:14px;margin-top:10px;flex-wrap:wrap}
    .viz-dot{width:9px;height:9px;border-radius:3px;flex-shrink:0}
    .viz-leg-item{display:flex;align-items:center;gap:5px;font-size:.77rem;color:var(--text2)}
    .att-date-nav{display:flex;align-items:center;gap:10px;margin-bottom:20px;background:var(--bg2);border:1px solid var(--border);border-radius:12px;padding:14px 18px}
    .att-nav-btn{width:34px;height:34px;border-radius:9px;border:1px solid var(--border);background:var(--bg3);display:flex;align-items:center;justify-content:center;cursor:pointer;color:var(--text2);transition:.2s;flex-shrink:0}
    .att-nav-btn:hover{background:var(--bg4);color:var(--text)}
    .att-nav-btn svg{width:16px;height:16px}
    .att-date-info{flex:1}
    .att-date-big{font-size:1.1rem;font-weight:700;color:var(--text)}
    .att-date-sub{font-size:.8rem;color:var(--text3);margin-top:1px}
    .att-today-btn{padding:5px 12px;border-radius:7px;border:1px solid var(--border);background:var(--bg3);font-size:.8rem;font-weight:600;color:var(--text2);cursor:pointer;font-family:var(--font);transition:.2s}
    .att-today-btn:hover{background:rgba(88,166,255,.1);color:var(--accent);border-color:var(--accent)}
    .status-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:20px}
    .status-tile{background:var(--bg2);border:2px solid var(--border);border-radius:14px;padding:24px 16px 20px;cursor:pointer;transition:all .2s;text-align:center;position:relative;overflow:hidden}
    .status-tile:hover{transform:translateY(-2px);box-shadow:0 8px 24px rgba(0,0,0,.2)}
    .status-tile.selected{box-shadow:0 8px 24px rgba(0,0,0,.25)}
    .status-tile-icon{display:none}
    .status-tile-label{font-size:.95rem;font-weight:700;color:var(--text)}
    .status-tile-sub{font-size:.75rem;color:var(--text3);margin-top:3px}
    .status-tile .sel-check{position:absolute;top:10px;right:10px;width:20px;height:20px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:11px;opacity:0;transition:.2s}
    .status-tile.selected .sel-check{opacity:1}
    .proc-task-section{background:var(--bg2);border:1px solid var(--border);border-radius:12px;padding:16px;margin-bottom:14px}
    .proc-task-title{font-size:.78rem;font-weight:700;text-transform:uppercase;letter-spacing:.7px;color:var(--text3);margin-bottom:12px}
    .att-save-btn{width:100%;padding:14px;border-radius:10px;background:var(--accent);border:none;color:#fff;font-family:var(--font);font-size:1.05rem;font-weight:700;cursor:pointer;transition:all .2s;display:flex;align-items:center;justify-content:center;gap:8px}
    .att-save-btn:hover{opacity:.88}
    .att-save-btn:disabled{opacity:.5;cursor:not-allowed}
    .wv-controls{display:flex;align-items:center;gap:10px;margin-bottom:20px;background:var(--bg2);border:1px solid var(--border);border-radius:12px;padding:16px 20px;flex-wrap:wrap}
    .wv-nav-btn{width:34px;height:34px;border-radius:9px;border:1px solid var(--border);background:var(--bg3);display:flex;align-items:center;justify-content:center;cursor:pointer;color:var(--text2);transition:.2s;flex-shrink:0}
    .wv-nav-btn:hover{background:var(--bg4);color:var(--text)}
    .wv-nav-btn svg{width:16px;height:16px}
    .wv-range{flex:1;min-width:0}.wv-range-title{font-size:.975rem;font-weight:700;color:var(--text)}.wv-range-sub{font-size:.78rem;color:var(--text3)}
    .wv-today-btn{padding:5px 12px;border-radius:7px;border:1px solid var(--border);background:var(--bg3);font-size:.8rem;font-weight:600;color:var(--text2);cursor:pointer;font-family:var(--font);transition:.2s}
    .wv-today-btn:hover{background:rgba(88,166,255,.1);color:var(--accent);border-color:var(--accent)}
    .wv-grid-wrap{border:1px solid var(--border);border-radius:12px;overflow:hidden;background:var(--bg2)}
    .wv-header{display:grid;grid-template-columns:160px repeat(7,1fr);background:var(--bg3);border-bottom:2px solid var(--border)}
    .wv-hcell{padding:10px 8px;text-align:center;font-size:.72rem;font-weight:700;color:var(--text3);text-transform:uppercase;letter-spacing:.5px}
    .wv-hcell.today-col{color:var(--accent);border-bottom:2px solid var(--accent);margin-bottom:-2px}
    .wv-hcell .wv-hdate{font-size:.68rem;color:var(--text3);margin-top:2px;text-transform:none;letter-spacing:0;font-weight:500}
    .wv-hcell.today-col .wv-hdate{color:var(--accent)}
    .wv-row{display:grid;grid-template-columns:160px repeat(7,1fr);border-bottom:1px solid var(--border);transition:.15s}
    .wv-row:last-child{border-bottom:none}
    .wv-row:hover{background:rgba(88,166,255,.025)}
    .wv-name-cell{padding:12px;display:flex;align-items:center;gap:9px;border-right:1px solid var(--border)}
    .wv-av{width:28px;height:28px;border-radius:7px;background:linear-gradient(135deg,var(--accent),var(--accent2));display:flex;align-items:center;justify-content:center;font-size:10px;font-weight:700;color:#fff;flex-shrink:0}
    .wv-name{font-size:.83rem;font-weight:600;color:var(--text);overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
    .wv-cell{padding:8px 5px;text-align:center;border-right:1px solid rgba(255,255,255,.03);display:flex;flex-direction:column;align-items:center;justify-content:center;gap:3px;cursor:pointer}
    .wv-cell:last-child{border-right:none}
    .wv-cell.today-col{background:rgba(88,166,255,.03)}
    .wv-cell.weekend-col{opacity:.4}
    .wv-cell:hover{background:rgba(88,166,255,.07)}
    .wv-detail-panel{background:var(--bg2);border:1px solid var(--accent);border-radius:12px;padding:16px;margin-top:14px;animation:fi .2s ease}
    .wv-dp-header{display:flex;align-items:center;justify-content:space-between;margin-bottom:14px}
    .wv-dp-title{font-size:.95rem;font-weight:700;color:var(--text)}
    .wv-dp-close{background:none;border:none;color:var(--text2);cursor:pointer;padding:4px;border-radius:6px;display:flex;align-items:center}
    .wv-dp-close:hover{background:var(--bg3);color:var(--text)}
    .wv-dp-close svg{width:16px;height:16px}
    .wv-dp-users{display:flex;flex-direction:column;gap:8px}
    .wv-dp-user{display:flex;align-items:center;gap:10px;padding:10px 12px;border-radius:9px;background:var(--bg3);border:1px solid var(--border)}
    .wv-dp-uinfo{flex:1;min-width:0}
    .wv-dp-uname{font-size:.875rem;font-weight:600;color:var(--text)}
    .wv-dp-uproc{font-size:.75rem;color:var(--text3);margin-top:2px}
    .wv-dp-utask{font-size:.75rem;color:var(--text2);margin-top:2px}
    .npt-log-form{background:var(--bg2);border:1px solid var(--border);border-radius:12px;padding:16px;margin-bottom:16px}
    .npt-log-form-title{font-size:.78rem;font-weight:700;text-transform:uppercase;letter-spacing:.7px;color:var(--text3);margin-bottom:12px}
    .npt-types{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:12px}
    .npt-type-btn{padding:5px 14px;border-radius:20px;border:1px solid var(--border);background:var(--bg3);color:var(--text2);font-family:var(--font);font-size:.82rem;font-weight:600;cursor:pointer;transition:all .15s}
    .npt-type-btn:hover{border-color:var(--accent);color:var(--accent)}
    .npt-type-btn.active{background:rgba(88,166,255,.12);border-color:var(--accent);color:var(--accent)}
    .npt-table-wrap{border:1px solid var(--border);border-radius:10px;overflow:hidden}
    .npt-table{width:100%;border-collapse:collapse;font-size:.9rem}
    .npt-table th{background:var(--bg3);padding:9px 12px;text-align:left;color:var(--text3);font-weight:600;font-size:.73rem;text-transform:uppercase;letter-spacing:.6px;border-bottom:1px solid var(--border)}
    .npt-table td{padding:9px 12px;color:var(--text2);border-bottom:1px solid var(--border)}
    .npt-table tr:last-child td{border-bottom:none}
    .npt-table tr:hover td{background:var(--bg3)}
    .npt-summary-cards{display:flex;gap:10px;flex-wrap:wrap;margin-bottom:14px}
    .npt-sum-card{background:var(--bg2);border:1px solid var(--border);border-radius:10px;padding:12px 16px;flex:1;min-width:110px}
    .npt-sum-card .lbl{font-size:.73rem;color:var(--text3);text-transform:uppercase;letter-spacing:.5px;margin-bottom:4px}
    .npt-sum-card .val{font-size:1.4rem;font-weight:800;color:var(--text)}

    /* ── Mark Attendance weekly grid ── */
    .att-week-row{display:grid;grid-template-columns:repeat(7,1fr);gap:8px;margin-bottom:16px}
    .att-day-card{background:var(--bg2);border:2px solid var(--border);border-radius:12px;padding:14px 8px;text-align:center;cursor:pointer;transition:all .18s;position:relative;user-select:none}
    .att-day-card:hover{border-color:var(--accent);transform:translateY(-2px);box-shadow:0 6px 18px rgba(0,0,0,.2)}
    .att-day-card.att-day-selected{border-color:var(--accent)!important;box-shadow:0 0 0 3px rgba(88,166,255,.2)}
    .att-day-card.att-day-today .att-day-name{color:var(--accent);font-weight:800}
    .att-day-card.att-day-weekend{opacity:.5}
    .att-day-card.att-day-weekend:hover{opacity:.75}
    .att-day-card.att-day-multi{outline:3px solid var(--accent2);outline-offset:2px;background:rgba(163,113,247,.08)!important}
    .att-bulk-bar{background:var(--bg2);border:1px solid var(--accent2);border-radius:12px;padding:14px 18px;margin-bottom:14px;display:flex;align-items:center;gap:10px;flex-wrap:wrap;animation:fi .2s ease}
    .att-bulk-status{display:flex;gap:6px;flex-wrap:wrap;flex:1}
    .att-bulk-btn{padding:7px 16px;border-radius:8px;border:2px solid var(--border);background:var(--bg3);color:var(--text2);font-family:var(--font);font-size:.85rem;font-weight:700;cursor:pointer;transition:all .15s}
    .att-bulk-btn:hover{transform:translateY(-1px);box-shadow:0 4px 12px rgba(0,0,0,.2)}
    .att-bulk-btn.sel{border-color:var(--accent2);background:rgba(163,113,247,.12);color:var(--accent2)}
    .att-select-toggle{padding:6px 14px;border-radius:8px;border:1px solid var(--border);background:var(--bg3);color:var(--text2);font-family:var(--font);font-size:.82rem;font-weight:600;cursor:pointer;transition:all .15s}
    .att-select-toggle.active{background:rgba(163,113,247,.12);border-color:var(--accent2);color:var(--accent2)}
    .att-day-name{font-size:.7rem;font-weight:700;text-transform:uppercase;letter-spacing:.6px;color:var(--text3);margin-bottom:2px}
    .att-day-date{font-size:.75rem;color:var(--text2);margin-bottom:6px}
    .att-day-icon{font-size:22px;margin-bottom:6px;line-height:1}
    .att-day-dot{width:10px;height:10px;border-radius:50%;margin:6px auto 4px;transition:.2s}
    .att-day-status{font-size:.68rem;font-weight:700;text-transform:uppercase;letter-spacing:.4px}
    .att-day-active-dot{position:absolute;bottom:6px;left:50%;transform:translateX(-50%);width:6px;height:6px;border-radius:50%;background:var(--accent)}
    .att-edit-card{background:var(--bg2);border:1px solid var(--border2);border-radius:14px;padding:22px;margin-top:8px;animation:fi .2s ease}
    .att-edit-header{display:flex;align-items:center;justify-content:space-between;margin-bottom:16px;flex-wrap:wrap;gap:8px}
    .att-edit-date{font-size:1rem;font-weight:700;color:var(--text)}
    .att-edit-datesub{font-size:.78rem;color:var(--text3);margin-top:2px}
    .att-saved-pill{padding:3px 10px;border-radius:20px;font-size:.78rem;font-weight:700}

    /* ── My Calendar ── */
    .cal-grid-wrap{background:var(--bg2);border:1px solid var(--border);border-radius:12px;overflow:hidden}
    .cal-header{display:grid;grid-template-columns:repeat(7,1fr);background:var(--bg3);border-bottom:1px solid var(--border)}
    .cal-hcell{padding:10px 4px;text-align:center;font-size:.72rem;font-weight:700;color:var(--text3);text-transform:uppercase;letter-spacing:.5px}
    .cal-grid{display:grid;grid-template-columns:repeat(7,1fr)}
    .cal-cell{min-height:70px;border:1px solid var(--border);padding:6px;display:flex;flex-direction:column;align-items:center;cursor:pointer;transition:all .15s;position:relative}
    .cal-cell:hover{background:var(--bg3)}
    .cal-cell.cal-weekend{background:rgba(255,255,255,.015);cursor:default}
    .cal-cell.cal-weekend:hover{background:rgba(255,255,255,.015)}
    .cal-cell.cal-today{box-shadow:inset 0 0 0 2px var(--accent)}
    .cal-cell.cal-future{opacity:.5}
    .cal-cell.cal-empty{background:transparent;border-color:transparent;cursor:default}
    .cal-cell.cal-cell-active{box-shadow:inset 0 0 0 2px var(--accent2)!important}
    .cal-day-num{font-size:.8rem;font-weight:700;color:var(--text2);margin-bottom:4px;align-self:flex-start}
    .cal-today-num{color:var(--accent);font-weight:800}
    .cal-day-badge{padding:3px 7px;border-radius:5px;font-size:.68rem;font-weight:800;color:#fff;letter-spacing:.4px;text-transform:uppercase}
    .cal-day-wknd{font-size:.68rem;color:var(--text3)}
    .cal-day-empty{width:100%}
    #dtr-toast{position:fixed;bottom:20px;right:20px;z-index:2147483648;padding:10px 16px;border-radius:10px;background:var(--bg2);border:1px solid var(--border2);color:var(--text);font-family:var(--font);font-size:.92rem;font-weight:500;box-shadow:0 8px 32px rgba(0,0,0,.4);transform:translateY(16px);opacity:0;transition:all .25s cubic-bezier(.34,1.56,.64,1);display:flex;align-items:center;gap:8px;pointer-events:none;max-width:360px}
    #dtr-toast.show{transform:translateY(0);opacity:1}
    #dtr-toast.tok{border-color:rgba(63,185,80,.4)}
    #dtr-toast.terr{border-color:rgba(248,81,73,.4)}
    #dtr-toast.tinfo{border-color:rgba(88,166,255,.35)}
    hr.sep{border:none;border-top:1px solid var(--border);margin:12px 0}
    .anim{animation:fi .22s ease}
    @keyframes fi{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:translateY(0)}}
    .empty{text-align:center;padding:32px 20px;color:var(--text3)}
    .empty p{font-size:.94rem}
    .info-banner{padding:9px 13px;border-radius:8px;background:rgba(88,166,255,.07);border:1px solid rgba(88,166,255,.18);font-size:.81rem;color:var(--text2);margin-bottom:14px;display:flex;align-items:center;gap:8px;line-height:1.5}
    .warn-banner{padding:9px 13px;border-radius:8px;background:rgba(210,153,34,.07);border:1px solid rgba(210,153,34,.2);font-size:.81rem;color:var(--amber);margin-bottom:14px;display:flex;align-items:center;gap:8px;line-height:1.5}
  `);

  // ICONS
  const ic = {
    logo:`<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="white" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="3"/><path d="M7 8h10M7 12h6M7 16h8"/></svg>`,
    submit:`<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 5v14M5 12l7-7 7 7"/></svg>`,
    chart:`<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 3v18h18"/><path d="M7 14l4-4 4 4 4-4"/></svg>`,
    tracker:`<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M9 3v18"/></svg>`,
    mark:`<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M8 6h13M8 12h13M8 18h13M3 6h.01M3 12h.01M3 18h.01"/></svg>`,
    team:`<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87M16 3.13a4 4 0 0 1 0 7.75"/></svg>`,
    npt:`<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>`,
    settings:`<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="3"/><path d="M19.07 4.93l-1.41 1.41M4.93 4.93l1.41 1.41M12 2v2M12 20v2M20 12h2M2 12h2M19.07 19.07l-1.41-1.41M4.93 19.07l1.41-1.41"/></svg>`,
    sun:`<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="5"/><line x1="12" y1="1" x2="12" y2="3"/><line x1="12" y1="21" x2="12" y2="23"/><line x1="4.22" y1="4.22" x2="5.64" y2="5.64"/><line x1="18.36" y1="18.36" x2="19.78" y2="19.78"/><line x1="1" y1="12" x2="3" y2="12"/><line x1="21" y1="12" x2="23" y2="12"/><line x1="4.22" y1="19.78" x2="5.64" y2="18.36"/><line x1="18.36" y1="5.64" x2="19.78" y2="4.22"/></svg>`,
    moon:`<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"/></svg>`,
    close:`<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>`,
    logout:`<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4M16 17l5-5-5-5M21 12H9"/></svg>`,
    left:`<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="15 18 9 12 15 6"/></svg>`,
    right:`<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="9 18 15 12 9 6"/></svg>`,
    check:`<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="20 6 9 17 4 12"/></svg>`,
    add:`<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>`,
    trash:`<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14H6L5 6"/><path d="M10 11v6M14 11v6M9 6V4h6v2"/></svg>`,
    download:`<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>`,
    sync:`<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="23 4 23 10 17 10"/><polyline points="1 20 1 14 7 14"/><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/></svg>`,
    att:`<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>`,
  };

  const TABS = [
    { id:'submit',    label:'Submit Tasks',    icon:ic.submit,   group:'Task Reporter' },
    { id:'analytics', label:'My Analytics',    icon:ic.chart,    group:'Task Reporter' },
    { id:'tracker',   label:'My Tracker',      icon:ic.tracker,  group:'Task Reporter' },
    { id:'mark',      label:'Mark Attendance', icon:ic.mark,     group:'Attendance'    },
    { id:'calendar',  label:'My Calendar',     icon:ic.att,      group:'Attendance'    },
    { id:'missednpt', label:'Missed NPT',      icon:ic.npt,      group:'Attendance'    },
    { id:'settings',  label:'Settings',        icon:ic.settings, group:'Settings'      },
  ];

  function boot() {
    const ex = document.getElementById('dtr-root'); if (ex) ex.remove();
    const outer = document.createElement('div'); outer.id = 'dtr-root-outer';
    root = document.createElement('div'); root.id = 'dtr-root';
    if (theme === 'light') { root.classList.add('lt'); outer.classList.add('lt'); }
    outer.appendChild(root);
    document.body.appendChild(outer);
    root.addEventListener('click', handleClick);
    root.addEventListener('input', handleInput);
    root.addEventListener('change', handleChange);
    root.addEventListener('keydown', handleKeydown);
    if (window.location.hostname.includes('sharepoint.com')) setTimeout(flushQueue, 2000);
    teamStatusCache = safeLoadObj('dtr_teamcache', {});
    if (authState.loggedIn) {
      buildApp();
      // Silently sync own attendance from SP so calendar is populated
      setTimeout(() => fetchOwnAttendance(true), 1500);
    } else {
      renderLogin();
    }
  }

  const q  = sel => root.querySelector(sel);
  const qa = sel => root.querySelectorAll(sel);

  function renderLogin() {
    root.innerHTML = `<div id="view-login"><div class="login-card">
      <div class="login-logo">${ic.logo}</div>
      <div class="login-title">WorkPulse</div>
      <div class="login-sub">Enter your full name to start</div>
      <div class="login-err" id="login-err"></div>
      <div class="dtr-field"><label class="dtr-label">Full Name</label>
        <input type="text" id="login-name" class="dtr-input" placeholder="e.g. John Smith" autocomplete="off">
      </div>
      <button class="btn btn-primary btn-full" data-action="do-login" style="margin-top:4px">${ic.submit} Start Session</button>
    </div></div><div id="dtr-toast"></div>`;
    setTimeout(() => { const e = q('#login-name'); if (e) e.focus(); }, 100);
  }

  function doLogin() {
    const el = q('#login-name'), name = el ? el.value.trim() : '';
    const err = q('#login-err');
    if (!name || name.length < 2) { if (err) err.textContent = 'Please enter your full name (min 2 chars).'; return; }
    const sess = saveSession(name);
    authState = { loggedIn:true, name, sid:sess.sid };
    buildApp();
    toast('Welcome, ' + name + '!', 'ok');
    // Load own attendance from SP immediately after login
    setTimeout(() => fetchOwnAttendance(true), 1000);
  }

  function doLogout() {
    if (!confirm('Log out as "' + authState.name + '"?')) return;
    clearSession(); authState = { loggedIn:false, name:'', sid:'' };
    renderLogin(); toast('Logged out.', 'info');
  }

  function buildApp() {
    const ms = mySubmissions();
    const todaySt = statusCache[todayStr()];
    let sbHtml = '', grp = '';
    TABS.forEach(t => {
      if (t.group !== grp) { grp = t.group; sbHtml += '<div class="sb-sec">' + grp + '</div>'; }
      sbHtml += '<button class="sb-item" data-action="switch-tab" data-val="' + t.id + '">' + t.icon + ' ' + t.label + '</button>';
    });
    root.innerHTML =
      '<div id="dtr-topbar">' +
        '<div class="dtr-logo"><div class="dtr-logo-icon">' + ic.logo + '</div><span class="dtr-logo-text">WorkPulse</span></div>' +
        '<div class="dtr-tabs">' + TABS.map(t => '<button class="dtr-tab" data-action="switch-tab" data-val="' + t.id + '">' + t.icon + ' ' + t.label + '</button>').join('') + '</div>' +
        '<div class="dtr-topbar-r">' +
          (todaySt ? '<span class="sp sp-' + (todaySt.status||'').toLowerCase().replace(' ','-') + '" style="font-size:.75rem">' + todaySt.status + '</span>' : '') +
          '<div class="dtr-user-pill"><div class="av">' + authState.name[0].toUpperCase() + '</div><span>' + authState.name + '</span><span class="role-badge">Associate</span></div>' +
          '<button class="icon-btn" data-action="toggle-theme" id="theme-btn">' + (theme==='dark'?ic.sun:ic.moon) + '</button>' +
          '<button class="icon-btn danger" data-action="do-logout" title="Logout">' + ic.logout + '</button>' +
          '<button class="icon-btn danger" data-action="close-app">' + ic.close + '</button>' +
        '</div>' +
      '</div>' +
      '<div id="dtr-body">' +
        '<div id="dtr-sidebar">' + sbHtml +
          '<hr class="sep" style="margin:8px 0"><div class="sb-footer">' +
          '<div class="sb-stat"><div class="lbl">My Entries</div><div class="val" id="sb-total">' + ms.length + '</div></div>' +
          '<div class="sb-stat"><div class="lbl">Productivity</div><div class="val" style="color:var(--green)" id="sb-avg">' + calcAvgProd(ms).toFixed(0) + '%</div></div>' +
          '</div></div>' +
        '<div id="dtr-main">' + TABS.map(t => '<div class="dtr-view anim" id="view-' + t.id + '" style="display:none"></div>').join('') + '</div>' +
      '</div><div id="dtr-toast"></div>';
    switchTab(TABS[0].id);
  }

  function switchTab(tab) {
    currentView = tab;
    qa('.dtr-tab').forEach(b => b.classList.toggle('active', b.dataset.val === tab));
    qa('.sb-item').forEach(b => b.classList.toggle('active', b.dataset.val === tab));
    qa('.dtr-view').forEach(v => v.style.display = 'none');
    const el = q('#view-' + tab);
    if (el) { el.style.display = 'block'; el.classList.remove('anim'); void el.offsetWidth; el.classList.add('anim'); }
    renderView(tab);
  }

  function renderView(v) {
    if (v === 'submit')    renderSubmit();
    if (v === 'analytics') renderAnalytics();
    if (v === 'tracker')   renderTracker();
    if (v === 'mark')      renderMarkAttendance();
    if (v === 'calendar')  renderMyCalendar();
    if (v === 'missednpt') renderMissedNPT();
    if (v === 'settings')  renderSettings();
  }

  function toggleTheme() {
    theme = theme === 'dark' ? 'light' : 'dark';
    GM_setValue('dtr_theme4', theme);
    root.classList.toggle('lt', theme === 'light');
    const b = q('#theme-btn'); if (b) b.innerHTML = theme === 'dark' ? ic.sun : ic.moon;
  }

  function handleClick(e) {
    const btn = e.target.closest('[data-action]'); if (!btn) return;
    const a = btn.dataset.action, v = btn.dataset.val || '';
    switch (a) {
      case 'do-login':      doLogin(); break;
      case 'do-logout':     doLogout(); break;
      case 'switch-tab':    switchTab(v); break;
      case 'toggle-theme':  toggleTheme(); break;
      case 'close-app':     root.style.display = 'none'; break;
      case 'mode-single':   setTaskMode('single', btn); break;
      case 'mode-multi':    setTaskMode('multi', btn); break;
      case 'add-card':      addTaskCard(); break;
      case 'remove-card':   btn.closest('.task-card').remove(); renumberCards(); recalcNPT(); break;
      case 'leave-card':    makeLeave(btn.closest('.task-card')); break;
      case 'leave-type':    setLeaveType(btn.closest('.task-card'), v); break;
      case 'clear-submit':  if (confirm('Clear all tasks?')) renderSubmit(); break;
      case 'pick-date':     pickDate(v); break;
      case 'do-submit':     doSubmit(); break;
      case 'tracker-add':     toast('Tracker is auto-populated from submissions','info'); break;
      case 'tracker-add-col': trackerAddCol(); break;
      case 'tracker-del-col': trackerDelCol(btn.dataset.key); break;
      case 'tracker-del':     toast('Remove the task from Submit Tasks to remove it here','info'); break;
      case 'tracker-export':trackerExport(); break;
      case 'att-nav-prev':   attNavDate(-1); break;
      case 'att-week-prev': {
        attWeekOffset--;
        // Set attDate to Monday of new week
        const ws2 = getWeekStart(attWeekOffset);
        ws2.setDate(ws2.getDate()+1); // Monday
        attDate = ws2.toISOString().split('T')[0];
        attSelectedStatus = (statusCache[attDate]||{}).status||'';
        renderMarkAttendance();
        break;
      }
      case 'att-week-next': {
        attWeekOffset++;
        const ws3 = getWeekStart(attWeekOffset);
        ws3.setDate(ws3.getDate()+1); // Monday
        attDate = ws3.toISOString().split('T')[0];
        attSelectedStatus = (statusCache[attDate]||{}).status||'';
        renderMarkAttendance();
        break;
      }
      case 'att-week-today':   attWeekOffset=0; attDate=todayStr(); attSelectedStatus=(statusCache[attDate]||{}).status||''; renderMarkAttendance(); break;
      case 'att-toggle-bulk':  attToggleBulk(); break;
      case 'att-multi-pick':   attMultiPick(v); break;
      case 'att-bulk-apply':   attBulkApply(v); break;
      case 'att-clear-select': attMultiSelect.clear(); renderMarkAttendance(); break;
      case 'att-apply-week':   attApplyToWeek(); break;
      case 'week-picker-apply': {
        const status = v;
        const picker = document.getElementById('att-week-picker');
        const overwrite = picker?.querySelector('#wk-overwrite')?.checked || false;
        if (picker) picker.remove();
        (async () => {
          const ws = getWeekStart(attWeekOffset);
          const weekdays = Array.from({length:5}, (_,i) => { const d=new Date(ws);d.setDate(ws.getDate()+1+i);return d.toISOString().split('T')[0]; });
          let saved=0;
          for (const dk of weekdays) {
            if (!overwrite && statusCache[dk]?.status) continue;
            const entry = {status,process:'',task:'',date:dk,name:authState.name,updatedAt:new Date().toISOString()};
            statusCache[dk]={...entry};
            teamStatusCache[authState.name+'::'+dk]={status,process:'',task:''};
            const ok=await postAttendance(entry); if(ok)saved++;
          }
          safeSaveObj('dtr_status2',statusCache); safeSaveObj('dtr_teamcache',teamStatusCache);
          toast('✅ '+status+' applied to '+saved+' weekday(s)','ok');
          renderMarkAttendance();
        })();
        break;
      }
      case 'week-picker-cancel': { const p=document.getElementById('att-week-picker'); if(p)p.remove(); break; }
      case 'toggle-note': {
        const n = btn.dataset.n;
        const box = document.getElementById('opt-note-' + n);
        const lbl = btn.querySelector('.opt-note-lbl');
        if (box) { const shown = box.style.display === 'block'; box.style.display = shown ? 'none' : 'block'; if(lbl) lbl.textContent = shown ? 'Add optional note' : 'Hide note'; }
        break;
      }
      case 'att-pick-day': {
        attDate = v;
        attSelectedStatus = (statusCache[v]||{}).status||''; // set BEFORE buildAttEditPanel
        // Highlight selected day card
        qa('.att-day-card').forEach(c => c.classList.toggle('att-day-selected', c.dataset.val === v));
        // Re-render only the edit panel (attSelectedStatus already set correctly above)
        const p = q('#att-edit-panel');
        if (p) p.innerHTML = buildAttEditPanel(v);
        break;
      }
      case 'att-nav-next':  attNavDate(+1); break;
      case 'att-nav-today': attNavToday(); break;
      case 'att-status':    attSelectStatus(v); break;
      case 'att-save':      attSave(); break;
      case 'cal-sync':      toast('Loading your attendance...','info'); fetchOwnAttendance(false); break;
      case 'cal-prev':      calMonthOffset--; renderMyCalendar(); break;
      case 'cal-next':      calMonthOffset++; renderMyCalendar(); break;
      case 'cal-today':     calMonthOffset=0; renderMyCalendar(); break;
      case 'npt-type':      nptSelectType(v); break;
      case 'npt-log':       nptLogEntry(); break;
      case 'npt-del':       nptDel(+btn.dataset.idx); break;
      case 'test-sp':       testSP(); break;
      case 'flush-queue':   flushQueueManual(); break;
    }
  }

  function handleInput(e) {
    const el = e.target;
    if (el.classList.contains('dtr-hrs')) { recalcNPT(); updateDayViz(); }
  }

  function handleChange(e) {
    const el = e.target;
    if (el.classList.contains('dtr-tt'))    onTaskTypeChange(el);
    if (el.classList.contains('dtr-wtype')) recalcNPT();
    // exfield removed — tracker is now auto-populated from submissions
  }

  function handleKeydown(e) {
    if (e.key === 'Enter' && e.target.id === 'login-name') doLogin();
  }

  // SUBMIT TAB
  function renderSubmit() {
    taskCounter = 1;
    if (!isAllowedDate(selectedDate)) selectedDate = todayStr();
    const el = q('#view-submit'); if (!el) return;
    el.innerHTML =
      '<div class="ph"><div class="ph-left"><div class="ph-title">Submit Daily Tasks</div>' +
      '<div class="ph-sub">' + formatDay(selectedDate) + '</div></div>' +
      '<div class="ph-actions"><button class="btn btn-ghost btn-sm" data-action="clear-submit">Clear All</button></div></div>' +
      '<div class="card"><div class="card-title">Employee Info</div><div class="g2">' +
      '<div class="dtr-field"><label class="dtr-label">Full Name</label><input class="dtr-input" value="' + authState.name + '" readonly style="opacity:.7"></div>' +
      '<div class="dtr-field"><label class="dtr-label">Date <span style="font-size:.73rem;color:var(--text3)">(Today or Yesterday only)</span></label>' +
      '<div class="date-pill-row">' +
      '<button class="date-pill' + (selectedDate===todayStr()?' active':'') + '" data-action="pick-date" data-val="' + todayStr() + '">Today — ' + formatDate(todayStr()) + '</button>' +
      '<button class="date-pill' + (selectedDate===yesterdayStr()?' active':'') + '" data-action="pick-date" data-val="' + yesterdayStr() + '">Yesterday — ' + formatDate(yesterdayStr()) + '</button>' +
      '</div><input type="hidden" id="a-date" value="' + selectedDate + '"></div></div></div>' +
      '<div class="card"><div class="card-title" style="display:flex;justify-content:space-between;align-items:center">' +
      '<span>Tasks</span><div class="mode-toggle"><button class="mode-btn active" data-action="mode-single">Single</button>' +
      '<button class="mode-btn" data-action="mode-multi">Multi-Task</button></div></div>' +
      '<div id="a-cards">' + buildTaskCard(1, false) + '</div>' +
      '<button class="add-task-btn" id="a-add-btn" style="display:none" data-action="add-card">+ Add Another Task</button></div>' +
      '<div class="viz-card" id="day-viz"><div class="viz-title"><span>Daily Breakdown</span><span id="dviz-summary" style="font-weight:500;color:var(--text3)">0h / 8h</span></div>' +
      '<div class="viz-bar-wrap"><div class="viz-bar-label"><span style="color:var(--green)">Productive</span><span id="dviz-prod-val" style="color:var(--green)">0.0h</span></div>' +
      '<div class="viz-track"><div class="viz-fill" id="dviz-prod" style="width:0%;background:linear-gradient(90deg,#3fb950,#06b6d4)"><span></span></div></div></div>' +
      '<div class="viz-bar-wrap"><div class="viz-bar-label"><span style="color:var(--amber)">NPT</span><span id="dviz-npt-val" style="color:var(--amber)">0.0h</span></div>' +
      '<div class="viz-track"><div class="viz-fill" id="dviz-npt" style="width:0%;background:linear-gradient(90deg,#d29922,#ef4444)"><span></span></div></div></div></div>' +
      '<button class="btn btn-primary btn-full" style="padding:12px;font-size:1rem" data-action="do-submit">' + ic.submit + ' Submit Report</button>';
    recalcNPT(); updateDayViz();
  }

  function buildTaskCard(n, removable) {
    const opts = TASK_TYPES.map(t => '<option value="' + t + '">' + t + '</option>').join('');
    return '<div class="task-card" data-n="' + n + '">' +
      '<div class="tc-header"><span class="tc-num">Task ' + n + '</span><div style="display:flex;gap:5px">' +
      '<button class="btn btn-ghost btn-xs" data-action="leave-card">Leave</button>' +
      (removable ? '<button class="btn btn-danger btn-xs" data-action="remove-card">' + ic.close + '</button>' : '') +
      '</div></div>' +
      '<div class="ga">' +
      '<div class="dtr-field"><label class="dtr-label">Task Type</label><select class="dtr-select dtr-tt"><option value="">Select</option>' + opts + '<option value="Leave">Leave Day</option></select></div>' +
      '<div class="dtr-field"><label class="dtr-label">Work Type</label><select class="dtr-select dtr-wtype"><option value="Productive">Productive</option><option value="NPT">NPT</option></select></div>' +
      '<div class="dtr-field"><label class="dtr-label">Hours</label><input type="number" class="dtr-input dtr-hrs" min="0" max="8" step="0.5" placeholder="0.0"></div>' +
      '<div class="dtr-field"><label class="dtr-label">NPT Hours</label><div class="npt-box">—</div></div></div>' +
      '<div class="comments-box" id="req-comment-' + n + '"><div class="dtr-field" style="margin-top:4px"><label class="dtr-label dtr-comment-label">Comments <span class="req-star">*</span></label><textarea class="dtr-input dtr-adhoc" rows="2" placeholder="Add details..." style="resize:vertical;min-height:52px"></textarea></div></div>' +'<div class="opt-note-toggle" data-action="toggle-note" data-n="' + n + '">📝 <span class="opt-note-lbl">Add optional note</span></div><div class="opt-note-box" id="opt-note-' + n + '" style="display:none"><div class="dtr-field" style="margin-top:6px"><label class="dtr-label">Note <span style="color:var(--text3);font-weight:400">(optional)</span></label><textarea class="dtr-input dtr-note" rows="2" placeholder="Any additional context..." style="resize:vertical;min-height:48px"></textarea></div></div>' +
      '<div class="leave-tag"><span>Leave:</span>' +
      '<button class="leave-type-btn active" data-action="leave-type" data-val="full">Full Day (8h)</button>' +
      '<button class="leave-type-btn" data-action="leave-type" data-val="half-am">Half AM (4h)</button>' +
      '<button class="leave-type-btn" data-action="leave-type" data-val="half-pm">Half PM (4h)</button>' +
      '<span class="leave-hours-display" style="margin-left:auto;font-family:var(--mono);font-weight:700">8.0h</span></div></div>';
  }

  function pickDate(d) { if (!isAllowedDate(d)) { toast('Only Today/Yesterday allowed','err'); return; } selectedDate = d; const hi = q('#a-date'); if (hi) hi.value = d; qa('.date-pill').forEach(p => p.classList.toggle('active', p.dataset.val === d)); const ps = q('.ph-left .ph-sub'); if (ps) ps.textContent = formatDay(d); }
  function setTaskMode(mode, btn) { qa('.mode-btn').forEach(b => b.classList.remove('active')); btn.classList.add('active'); const ab = q('#a-add-btn'), cs = q('#a-cards'); if (mode === 'multi') { if (ab) ab.style.display = 'block'; if (cs && cs.querySelectorAll('.task-card').length === 1) addTaskCard(); } else { if (ab) ab.style.display = 'none'; if (cs) { const all = cs.querySelectorAll('.task-card'); for (let i = all.length - 1; i > 0; i--) all[i].remove(); taskCounter = 1; } } updateDayViz(); }
  function addTaskCard() { taskCounter++; const cs = q('#a-cards'); if (!cs) return; const d = document.createElement('div'); d.innerHTML = buildTaskCard(taskCounter, true); cs.appendChild(d.firstElementChild); updateDayViz(); }
  function renumberCards() { qa('.tc-num').forEach((el, i) => el.textContent = 'Task ' + (i+1)); taskCounter = qa('.task-card').length; }
  function onTaskTypeChange(sel) {
    const card = sel.closest('.task-card'); if (!card) return;
    const v = sel.value;
    const isLeave   = v === 'Leave';
    const isNPT     = NPT_TASKS.includes(v) || isLeave;
    const needsComment = v === 'Ad-hoc Tasks' || v === 'Other' || v === 'Meeting';
    // Show/hide comment box
    const box = card.querySelector('.comments-box');
    if (box) box.style.display = needsComment ? 'block' : 'none';
    // Req star visibility
    const star = card.querySelector('.req-star');
    if (star) star.style.display = needsComment ? '' : 'none';
    // Show/hide opt note toggle
    const optToggle = card.querySelector('.opt-note-toggle');
    if (optToggle) optToggle.style.display = needsComment ? 'none' : 'flex';
    // Update label text based on type
    const lbl = card.querySelector('.dtr-comment-label');
    const lblText = v==='Meeting' ? 'Meeting Details *' : v==='Other' ? 'Comments *' : 'Ad-hoc Description *';
    if (lbl) lbl.textContent = lblText;
    // Update placeholder
    const ta = card.querySelector('.dtr-adhoc');
    const phText = v==='Meeting' ? 'Meeting name, agenda, or purpose...' : v==='Other' ? 'Describe what you worked on...' : 'Describe the ad-hoc task...';
    if (ta) ta.placeholder = phText;
    card.querySelector('.leave-tag').classList.toggle('show', isLeave);
    if (isLeave) { makeLeave(card); return; }
    const ws = card.querySelector('.dtr-wtype'); if (ws) ws.value = isNPT ? 'NPT' : 'Productive';
    recalcNPT(); updateDayViz();
  }
  function makeLeave(card) { if (!card) return; card.classList.add('leave'); card.querySelector('.leave-tag').classList.add('show'); const ws = card.querySelector('.dtr-wtype'); if (ws) ws.value = 'NPT'; const hi = card.querySelector('.dtr-hrs'); if (hi) hi.value = ''; const nb = card.querySelector('.npt-box'); if (nb) { nb.textContent = '8.0h NPT'; nb.className = 'npt-box'; } const ts = card.querySelector('.dtr-tt'); if (ts) ts.value = 'Leave'; updateDayViz(); }
  function setLeaveType(card, type) { if (!card) return; card.querySelectorAll('.leave-type-btn').forEach(b => b.classList.toggle('active', b.dataset.val === type)); const h = type === 'full' ? 8 : 4; const ld = card.querySelector('.leave-hours-display'); if (ld) ld.textContent = h + '.0h'; const nb = card.querySelector('.npt-box'); if (nb) nb.textContent = h + '.0h NPT'; updateDayViz(); }
  function recalcNPT() { qa('.task-card').forEach(card => { if (card.classList.contains('leave')) return; const h = parseFloat(card.querySelector('.dtr-hrs')?.value)||0, wt = card.querySelector('.dtr-wtype')?.value||'Productive', nb = card.querySelector('.npt-box'); if (!nb) return; if (h===0){nb.textContent='—';nb.className='npt-box';return;} if (wt==='Productive'){nb.textContent=Math.max(0,WH-h).toFixed(1);nb.className='npt-box prod';}else{nb.textContent=h.toFixed(1)+' NPT';nb.className='npt-box';} }); }
  function updateDayViz() { let prod=0,npt=0; qa('.task-card').forEach(card=>{if(card.classList.contains('leave')){npt+=8;return;}const h=parseFloat(card.querySelector('.dtr-hrs')?.value)||0,wt=card.querySelector('.dtr-wtype')?.value||'Productive';if(wt==='NPT')npt+=h;else{prod+=h;npt+=Math.max(0,WH-h);}}); npt=Math.min(WH,Math.max(0,npt));prod=Math.min(WH,prod); const pEl=q('#dviz-prod'),nEl=q('#dviz-npt'),pv=q('#dviz-prod-val'),nv=q('#dviz-npt-val'),sv=q('#dviz-summary'); if(pEl)pEl.style.width=(prod/WH*100)+'%';if(nEl)nEl.style.width=(npt/WH*100)+'%';if(pv)pv.textContent=prod.toFixed(1)+'h';if(nv)nv.textContent=npt.toFixed(1)+'h';if(sv)sv.textContent=prod.toFixed(1)+'h prod + '+npt.toFixed(1)+'h NPT / 8h'; }

  async function doSubmit() {
    const dateVal = q('#a-date')?.value||selectedDate; const cards = qa('.task-card'); const tasks = []; let hasErr = false;
    cards.forEach(card => {
      const isLeave = card.classList.contains('leave');
      if (isLeave) { const lh = card.querySelector('.leave-hours-display')?.textContent.includes('4')?4:8; tasks.push({employeeName:authState.name,taskType:'Leave',hours:0,npt:lh,workType:'NPT',adhoc:'',date:dateVal,submittedAt:new Date().toISOString()}); return; }
      const tt = card.querySelector('.dtr-tt')?.value||'', wt = card.querySelector('.dtr-wtype')?.value||'Productive', h = parseFloat(card.querySelector('.dtr-hrs')?.value)||0;
      const reqComment = card.querySelector('.dtr-adhoc')?.value?.trim()||'';
      const optNote    = card.querySelector('.dtr-note')?.value?.trim()||'';
      const ad = reqComment + (reqComment && optNote ? ' | Note: ' + optNote : optNote ? 'Note: ' + optNote : '');
      if (!tt) { toast('Select task type','err'); hasErr=true; return; }
      if (h<=0&&wt!=='NPT') { toast('Enter hours','err'); hasErr=true; return; }
      if ((tt==='Ad-hoc Tasks'||tt==='Other'||tt==='Meeting')&&!reqComment) { toast('Please add ' + (tt==='Meeting'?'meeting details':'a comment') + ' for ' + tt,'err'); hasErr=true; return; }
      const npt = wt==='NPT'?h:Math.max(0,WH-h);
      tasks.push({employeeName:authState.name,taskType:tt,hours:h,npt,workType:wt,adhoc:ad,date:dateVal,submittedAt:new Date().toISOString()});
    });
    if (hasErr||!tasks.length) return;
    const sb = q('[data-action="do-submit"]'); if (sb) { sb.disabled=true; sb.textContent='Submitting...'; }
    tasks.forEach(t => submissions.push(t)); safeSave('dtr_subs4', submissions); updateSBStats();
    let posted = 0;
    for (const t of tasks) { const s = await postTask(t); if (s) posted++; else spQueue.push(t); }
    safeSave('dtr_spq2', spQueue);
    if (sb) { sb.disabled=false; sb.innerHTML=ic.submit+' Submit Report'; }
    toast(posted===tasks.length ? '✅ '+tasks.length+' task(s) submitted!' : 'Saved locally ('+( tasks.length-posted)+' queued)', 'ok');
    selectedDate = todayStr(); renderSubmit();
  }

  // ANALYTICS TAB
  function renderAnalytics() {
    const el = q('#view-analytics'); if (!el) return;
    const ms = mySubmissions(); const prod = calcAvgProd(ms), totalH = ms.reduce((a,s)=>a+(s.hours||0),0), streak = calcStreak(ms);
    const last7 = getLast7(); const dayData = last7.map(d => ({ d, h: Math.min(WH, ms.filter(s=>s.date===d).reduce((a,s)=>a+(s.hours||0),0)) }));
    const tb = {}; ms.filter(s=>!s.taskType?.startsWith('Leave')).forEach(s=>{tb[s.taskType]=(tb[s.taskType]||0)+(s.hours||0);});
    const maxT = Math.max(...Object.values(tb), 1);
    const rc = prod>=75?'var(--green)':prod>=50?'var(--amber)':'var(--red)';
    const circ = 2*Math.PI*38, off = circ-(prod/100)*circ;
    const attDays = Object.entries(statusCache);
    const wfoCnt = attDays.filter(([,v])=>v.status==='WFO').length;
    const wfhCnt = attDays.filter(([,v])=>v.status==='WFH').length;
    const lvCnt  = attDays.filter(([,v])=>['SL','CL','AL'].includes(v.status)).length;
    el.innerHTML =
      '<div class="ph"><div class="ph-left"><div class="ph-title">My Analytics</div><div class="ph-sub">' + authState.name + '</div></div></div>' +
      '<div class="stats-grid sg4"><div class="stat-card ab"><div class="lbl">Entries</div><div class="val">' + ms.length + '</div></div>' +
      '<div class="stat-card gb"><div class="lbl">Total Hours</div><div class="val">' + totalH.toFixed(0) + 'h</div></div>' +
      '<div class="stat-card amb"><div class="lbl">Avg Productivity</div><div class="val">' + prod.toFixed(0) + '%</div></div>' +
      '<div class="stat-card pb"><div class="lbl">Streak</div><div class="val">' + streak + '🔥</div></div></div>' +
      '<div class="stats-grid sg3"><div class="stat-card gb"><div class="lbl">WFO Days</div><div class="val" style="color:#3fb950">' + wfoCnt + '</div></div>' +
      '<div class="stat-card ab"><div class="lbl">WFH Days</div><div class="val" style="color:#22d3ee">' + wfhCnt + '</div></div>' +
      '<div class="stat-card amb"><div class="lbl">Leave Days</div><div class="val" style="color:var(--amber)">' + lvCnt + '</div></div></div>' +
      '<div class="g2"><div class="card"><div class="chart-title">Overall Score</div><div style="display:flex;align-items:center;gap:20px">' +
      '<svg width="90" height="90" viewBox="0 0 90 90"><circle cx="45" cy="45" r="38" fill="none" stroke="var(--bg4)" stroke-width="9"/>' +
      '<circle cx="45" cy="45" r="38" fill="none" stroke="' + rc + '" stroke-width="9" stroke-dasharray="' + circ.toFixed(2) + '" stroke-dashoffset="' + off.toFixed(2) + '" stroke-linecap="round" transform="rotate(-90 45 45)" style="transition:stroke-dashoffset 1s ease"/>' +
      '<text x="45" y="45" text-anchor="middle" dy=".35em" fill="' + rc + '" font-size="15" font-weight="800" font-family="DM Sans,sans-serif">' + prod.toFixed(0) + '%</text></svg>' +
      '<div><div style="font-size:1.4rem;font-weight:800;color:var(--text)">' + prod.toFixed(0) + '%</div><div style="font-size:.85rem;color:var(--text3);margin-top:3px">' + totalH.toFixed(1) + ' hrs · ' + ms.length + ' entries</div></div>' +
      '</div></div>' +
      '<div class="card"><div class="chart-title">Task Breakdown</div><div class="bar-rows">' +
      (Object.entries(tb).sort((a,b)=>b[1]-a[1]).map(([t,h],i) => '<div class="bar-row"><div class="bar-label">' + t + '</div><div class="bar-track"><div class="bar-fill" style="width:' + (h/maxT*100) + '%;background:' + COLORS[i%COLORS.length] + '"><span class="bar-fill-text">' + h.toFixed(1) + 'h</span></div></div></div>').join('') || '<div class="empty"><p>No data yet</p></div>') +
      '</div></div></div>' +
      '<div class="card"><div class="chart-title">Last 7 Days</div><div class="bar-rows">' +
      dayData.map(d => '<div class="bar-row"><div class="bar-label">' + formatDate(d.d) + '</div><div class="bar-track"><div class="bar-fill" style="width:' + (d.h/WH*100) + '%;background:linear-gradient(90deg,var(--accent),var(--accent2))">' + (d.h>0?'<span class="bar-fill-text">'+d.h.toFixed(1)+'h</span>':'') + '</div></div></div>').join('') +
      '</div></div>' +
      '<div class="card"><div class="card-title">Recent Submissions</div><div class="tbl-wrap"><table class="dtr-table"><thead><tr><th>Date</th><th>Task</th><th>Type</th><th>Hours</th><th>NPT</th></tr></thead><tbody>' +
      (ms.length ? ms.slice().reverse().slice(0,10).map(s => '<tr><td>' + s.date + '</td><td class="bold">' + s.taskType + '</td><td><span class="badge ' + (s.workType==='NPT'?'ba':'bg2') + '">' + (s.workType||'Productive') + '</span></td><td class="mono">' + (s.hours||0).toFixed(1) + '</td><td class="mono" style="color:var(--amber)">' + (s.npt||0).toFixed(1) + '</td></tr>').join('') : '<tr><td colspan="5"><div class="empty"><p>No entries yet</p></div></td></tr>') +
      '</tbody></table></div></div>';
  }

  // ── MY TRACKER — auto-populated from submissions, custom columns ──────

  function getTrackerRows() {
    // Build rows from submissions for this associate
    const ms = mySubmissions().slice().sort((a,b) => a.date < b.date ? 1 : -1);
    return ms.map((s,i) => ({
      _idx:     i,
      _source:  'submission',
      date:     s.date || '',
      employee: s.employeeName || authState.name,
      taskType: s.taskType || '',
      workType: s.workType || 'Productive',
      hours:    (s.hours||0).toFixed(1),
      nptHours: (s.npt||0).toFixed(1),
      leave:    s.taskType === 'Leave' ? '✓' : '',
      adhoc:    s.adhoc || '',
      submittedAt: s.submittedAt ? s.submittedAt.replace('T',' ').slice(0,16) : '',
    }));
  }

  function renderTracker() {
    const el = q('#view-tracker'); if (!el) return;
    const rows = getTrackerRows();
    // Fixed columns
    const fixedCols = [
      {key:'date',       label:'Date'},
      {key:'employee',   label:'Employee'},
      {key:'taskType',   label:'Task Type'},
      {key:'workType',   label:'Work Type'},
      {key:'hours',      label:'Hours'},
      {key:'nptHours',   label:'NPT Hours'},
      {key:'leave',      label:'Leave'},
      {key:'adhoc',      label:'Comments / Ad-hoc'},
      {key:'submittedAt',label:'Submitted At'},
    ];
    const allCols = [...fixedCols, ...tkCustomCols];

    el.innerHTML =
      '<div class="ph"><div class="ph-left"><div class="ph-title">My Tracker</div>' +
      '<div class="ph-sub">Auto-populated from your submitted tasks — ' + rows.length + ' entries</div></div>' +
      '<div class="ph-actions">' +
      '<button class="btn btn-ghost btn-sm" data-action="tracker-add-col">' + ic.add + ' Add Column</button>' +
      '<button class="btn btn-ghost btn-sm" data-action="tracker-export">' + ic.download + ' Export CSV</button>' +
      '</div></div>' +
      // Filter bar
      '<div class="filter-bar" style="display:flex;gap:8px;margin-bottom:12px;flex-wrap:wrap">' +
      '<input type="text" id="tk-filter" class="dtr-input" placeholder="Filter by task, date, type..." style="max-width:240px" oninput="renderTrackerFiltered()">' +
      '<select id="tk-type-filter" class="dtr-select" style="max-width:160px" onchange="renderTrackerFiltered()">' +
      '<option value="">All Work Types</option>' +
      '<option value="Productive">Productive</option><option value="NPT">NPT</option>' +
      '</select>' +
      '<select id="tk-task-filter" class="dtr-select" style="max-width:180px" onchange="renderTrackerFiltered()">' +
      '<option value="">All Task Types</option>' +
      TASK_TYPES.map(t => '<option value="' + t + '">' + t + '</option>').join('') +
      '<option value="Leave">Leave</option></select>' +
      '</div>' +
      // Summary cards
      '<div class="stats-grid sg4" style="margin-bottom:14px">' +
      '<div class="stat-card ab"><div class="lbl">Total Entries</div><div class="val">' + rows.length + '</div></div>' +
      '<div class="stat-card gb"><div class="lbl">Total Hours</div><div class="val">' + rows.reduce((a,r)=>a+parseFloat(r.hours||0),0).toFixed(1) + 'h</div></div>' +
      '<div class="stat-card amb"><div class="lbl">NPT Hours</div><div class="val">' + rows.reduce((a,r)=>a+parseFloat(r.nptHours||0),0).toFixed(1) + 'h</div></div>' +
      '<div class="stat-card rb2"><div class="lbl">Leave Days</div><div class="val">' + rows.filter(r=>r.leave).length + '</div></div>' +
      '</div>' +
      // Table
      '<div class="tbl-wrap" style="overflow-x:auto"><table class="dtr-table" id="tk-table">' +
      '<thead><tr>' +
      '<th style="color:var(--text3);font-size:.72rem">#</th>' +
      allCols.map(c =>
        '<th>' + c.label +
        (tkCustomCols.find(cc=>cc.key===c.key) ?
          ' <button style="background:none;border:none;color:var(--red);cursor:pointer;font-size:.7rem;padding:0 2px" data-action="tracker-del-col" data-key="' + c.key + '">✕</button>' : '') +
        '</th>'
      ).join('') +
      '</tr></thead>' +
      '<tbody id="tk-tbody">' + buildTrackerRows(rows, allCols) + '</tbody>' +
      '</table></div>';

    // Bind inline oninput/onchange via event delegation since CSP may block inline
    const filterEl = el.querySelector('#tk-filter');
    const typeEl   = el.querySelector('#tk-type-filter');
    const taskEl   = el.querySelector('#tk-task-filter');
    if (filterEl) filterEl.addEventListener('input',  () => refreshTkBody(allCols));
    if (typeEl)   typeEl.addEventListener('change',   () => refreshTkBody(allCols));
    if (taskEl)   taskEl.addEventListener('change',   () => refreshTkBody(allCols));

    // Bind custom column cells
    el.addEventListener('input', e => {
      const cell = e.target.closest('[data-tk-key]');
      if (!cell) return;
      const key = cell.dataset.tkKey, idx = +cell.closest('tr').dataset.idx;
      const ms = mySubmissions();
      if (!ms[idx]) return;
      if (!ms[idx]._custom) ms[idx]._custom = {};
      ms[idx]._custom[key] = cell.value;
      submissions = submissions.map(s => (s===ms[idx] ? ms[idx] : s));
      safeSave('dtr_subs4', submissions);
    });
  }

  function buildTrackerRows(rows, allCols) {
    if (!rows.length) return '<tr><td colspan="' + (allCols.length+1) + '"><div class="empty"><p>No submissions yet. Submit your first task to see it here.</p></div></td></tr>';
    const ms = mySubmissions();
    return rows.map((r, i) => {
      const sub = ms[r._idx] || {};
      const wt = r.workType;
      const rowColor = r.leave ? 'rgba(210,153,34,.04)' : wt==='NPT' ? 'rgba(248,81,73,.03)' : '';
      return '<tr data-idx="' + r._idx + '" style="' + (rowColor?'background:'+rowColor:'') + '">' +
        '<td style="color:var(--text3);font-size:.78rem;font-weight:600">' + (i+1) + '</td>' +
        allCols.map(c => {
          const isCustom = !!tkCustomCols.find(cc=>cc.key===c.key);
          const val = isCustom ? ((sub._custom||{})[c.key]||'') : (r[c.key]||'');
          if (c.key==='workType') return '<td><span class="badge ' + (wt==='NPT'?'ba':'bg2') + '">' + val + '</span></td>';
          if (c.key==='leave' && r.leave) return '<td><span class="badge ba">Leave</span></td>';
          if (c.key==='leave') return '<td></td>';
          if (c.key==='hours'||c.key==='nptHours') return '<td class="mono" style="' + (c.key==='nptHours'&&parseFloat(val)>0?'color:var(--amber)':'') + '">' + val + '</td>';
          if (isCustom) return '<td><input class="dtr-input" value="' + val.replace(/"/g,'&quot;') + '" data-tk-key="' + c.key + '" style="padding:3px 7px;font-size:.85rem;min-width:90px"></td>';
          return '<td style="' + (c.key==='taskType'?'font-weight:600;color:var(--text)':'') + '">' + (val||'—') + '</td>';
        }).join('') + '</tr>';
    }).join('');
  }

  function refreshTkBody(allCols) {
    const filter   = (q('#tk-filter')?.value||'').toLowerCase();
    const typeF    = q('#tk-type-filter')?.value||'';
    const taskF    = q('#tk-task-filter')?.value||'';
    let rows = getTrackerRows();
    if (filter)  rows = rows.filter(r => Object.values(r).join(' ').toLowerCase().includes(filter));
    if (typeF)   rows = rows.filter(r => r.workType === typeF);
    if (taskF)   rows = rows.filter(r => r.taskType === taskF);
    const tbody = q('#tk-tbody'); if (tbody) tbody.innerHTML = buildTrackerRows(rows, allCols);
  }

  function trackerExport() {
    const rows = getTrackerRows();
    if (!rows.length) { toast('No submissions to export','err'); return; }
    const fixedCols = ['date','employee','taskType','workType','hours','nptHours','leave','adhoc','submittedAt'];
    const fixedLabels = ['Date','Employee','Task Type','Work Type','Hours','NPT Hours','Leave','Comments / Ad-hoc','Submitted At'];
    const customLabels = tkCustomCols.map(c=>c.label);
    const header = [...fixedLabels,...customLabels].join(',');
    const ms = mySubmissions();
    const csvRows = rows.map((r,i) => {
      const sub = ms[r._idx]||{};
      const fixed = fixedCols.map(k => '"'+(r[k]||'').replace(/"/g,'""')+'"');
      const custom = tkCustomCols.map(c => '"'+(((sub._custom||{})[c.key])||'').replace(/"/g,'""')+'"');
      return [...fixed,...custom].join(',');
    });
    dlCSV([header,...csvRows].join('\n'), 'tracker_'+authState.name.replace(/\s+/g,'_')+'_'+todayStr()+'.csv');
    toast('Exported ' + rows.length + ' rows','ok');
  }

  function trackerAddCol() {
    const label = prompt('New column name:');
    if (!label || !label.trim()) return;
    const key = 'custom_' + Date.now();
    tkCustomCols.push({ key, label: label.trim() });
    safeSave('dtr_tkcols4', tkCustomCols);
    renderTracker();
    toast('Column "' + label.trim() + '" added','ok');
  }

  function trackerDelCol(key) {
    tkCustomCols = tkCustomCols.filter(c => c.key !== key);
    safeSave('dtr_tkcols4', tkCustomCols);
    renderTracker();
  }

  // MARK ATTENDANCE
  let attDate = todayStr();
  let attSelectedStatus = '';
  let attMultiSelect    = new Set(); // dates selected for bulk apply
  let attBulkMode       = false;     // whether bulk-select is active

  // ── MARK ATTENDANCE — weekly grid UI, any date editable ──────────────
  let attWeekOffset = 0;  // week offset for mark tab

  function renderMarkAttendance() {
    const el = q('#view-mark'); if (!el) return;
    const today = todayStr();
    const ws = getWeekStart(attWeekOffset);
    const days = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
    const dates = Array.from({length:7}, (_,i) => {
      const d = new Date(ws); d.setDate(ws.getDate()+i);
      const dk = d.toISOString().split('T')[0];
      return {dk, day:days[i], label:d.toLocaleDateString('en-US',{month:'short',day:'numeric'}),
              isToday:dk===today, isWeekend:i===0||i===6, isFuture:dk>today};
    });
    const wLabel = getWeekLabel(attWeekOffset);

    // Bulk apply status label
    const bulkStatusLabel = {WFO:'WFO',WFH:'WFH',SL:'Sick Leave',CL:'Casual Leave',AL:'Annual Leave','Optional Off':'Optional Off'};

    el.innerHTML =
      '<div class="ph"><div class="ph-left"><div class="ph-title">Mark Attendance</div>' +
      '<div class="ph-sub">Click a day to mark · Or use multi-select to apply one status to many days</div></div></div>' +

      // Week navigator
      '<div class="wv-controls" style="margin-bottom:14px">' +
      '<button class="wv-nav-btn" data-action="att-week-prev">' + ic.left + '</button>' +
      '<div class="wv-range"><div class="wv-range-title">' + wLabel + '</div>' +
      '<div class="wv-range-sub">' + (attWeekOffset===0?'Current Week':attWeekOffset<0?Math.abs(attWeekOffset)+' week(s) ago':attWeekOffset+' week(s) ahead') + '</div></div>' +
      (attWeekOffset!==0?'<button class="wv-today-btn" data-action="att-week-today">This Week</button>':'') +

      // Quick actions
      '<button class="att-select-toggle' + (attBulkMode?' active':'') + '" data-action="att-toggle-bulk">' +
      (attBulkMode ? '✕ Cancel Multi-select' : '☑ Multi-select') + '</button>' +
      (!attBulkMode ? '<button class="att-select-toggle" data-action="att-apply-week" title="Apply one status to all weekdays this week">⚡ Apply to Week</button>' : '') +
      '<button class="wv-nav-btn" data-action="att-week-next">' + ic.right + '</button></div>' +

      // Bulk bar — shown when multi-select is active AND days are selected
      (attBulkMode && attMultiSelect.size > 0 ?
        '<div class="att-bulk-bar">' +
        '<span style="font-size:.82rem;font-weight:700;color:var(--accent2)">' + attMultiSelect.size + ' day' + (attMultiSelect.size>1?'s':'') + ' selected — Apply:</span>' +
        '<div class="att-bulk-status">' +
        STATUSES.map(s => '<button class="att-bulk-btn" data-action="att-bulk-apply" data-val="' + s + '">' + s + '</button>').join('') +
        '</div>' +
        '<button class="btn btn-ghost btn-xs" data-action="att-clear-select">Clear</button>' +
        '</div>' : '') +

      // Week day row
      '<div class="att-week-row">' +
      dates.map(d => {
        const st = (statusCache[d.dk]||{}).status||'';
        const cfg = st ? STATUS_CFG[st] : null;
        const isSelected = !attBulkMode && attDate === d.dk;
        const isMulti    = attBulkMode && attMultiSelect.has(d.dk);
        return '<div class="att-day-card' +
          (isSelected ? ' att-day-selected' : '') +
          (isMulti    ? ' att-day-multi'    : '') +
          (d.isToday  ? ' att-day-today'    : '') +
          (d.isWeekend? ' att-day-weekend'  : '') +
          '" data-action="' + (attBulkMode ? 'att-multi-pick' : 'att-pick-day') + '" data-val="' + d.dk + '"' +
          (cfg && !isMulti ? ' style="border-color:' + cfg.color + ';background:' + cfg.bg + '"' : '') + '>' +
          '<div class="att-day-name">' + d.day + '</div>' +
          '<div class="att-day-date">' + d.label + '</div>' +
          (st
            ? '<div class="att-day-dot" style="background:' + (cfg?cfg.color:'var(--text3)') + '"></div>' +
              '<div class="att-day-status" style="color:' + (cfg?cfg.color:'var(--text3)') + ';font-size:.72rem;font-weight:800">' + st + '</div>'
            : '<div class="att-day-dot" style="background:var(--border2)"></div>' +
              '<div class="att-day-status" style="color:var(--text3)">' + (d.isWeekend?'Weekend':'Tap') + '</div>') +
          (isMulti ? '<div style="position:absolute;top:6px;right:6px;width:16px;height:16px;border-radius:50%;background:var(--accent2);display:flex;align-items:center;justify-content:center;font-size:9px;color:#fff;font-weight:800">✓</div>' : '') +
          '</div>';
      }).join('') + '</div>' +

      // Single-day edit panel (only in normal mode)
      (!attBulkMode ? '<div id="att-edit-panel">' + buildAttEditPanel(attDate) + '</div>' :
        (attMultiSelect.size === 0 ? '<div class="info-banner">👆 Tap multiple days above to select them, then choose a status to apply to all at once.</div>' : ''));
  }

  function attToggleBulk() {
    attBulkMode = !attBulkMode;
    attMultiSelect.clear();
    if (!attBulkMode) attSelectedStatus = (statusCache[attDate]||{}).status||'';
    renderMarkAttendance();
  }

  function attMultiPick(dk) {
    if (attMultiSelect.has(dk)) attMultiSelect.delete(dk);
    else attMultiSelect.add(dk);
    renderMarkAttendance();
  }

  async function attBulkApply(status) {
    if (!status || attMultiSelect.size === 0) return;
    const btn = q('[data-action="att-bulk-apply"][data-val="' + status + '"]');
    if (btn) { btn.disabled = true; btn.textContent = 'Saving...'; }
    const cfg = STATUS_CFG[status] || {};
    let saved = 0;
    for (const dk of attMultiSelect) {
      const entry = { status, process:'', task:'', date:dk, name:authState.name, updatedAt:new Date().toISOString() };
      statusCache[dk] = { ...entry, _fromSP: false };
      teamStatusCache[authState.name+'::'+dk] = { status, process:'', task:'' };
      const ok = await postAttendance(entry);
      if (ok) saved++;
    }
    safeSaveObj('dtr_status2', statusCache);
    safeSaveObj('dtr_teamcache', teamStatusCache);
    toast('✅ ' + status + ' applied to ' + saved + '/' + attMultiSelect.size + ' days', 'ok');
    attMultiSelect.clear();
    attBulkMode = false;
    renderMarkAttendance();
  }

  async function attApplyToWeek() {
    // Show a quick status picker overlay
    const existing = q('#att-week-picker');
    if (existing) { existing.remove(); return; }
    const el = q('#view-mark'); if (!el) return;
    const picker = document.createElement('div');
    picker.id = 'att-week-picker';
    picker.style.cssText = 'position:fixed;top:50%;left:50%;transform:translate(-50%,-50%);z-index:99999;background:var(--bg2);border:2px solid var(--accent2);border-radius:16px;padding:24px;min-width:360px;box-shadow:0 24px 64px rgba(0,0,0,.5);animation:fi .2s ease';
    picker.innerHTML =
      '<div style="font-size:.95rem;font-weight:700;color:var(--text);margin-bottom:4px">Apply to All Weekdays</div>' +
      '<div style="font-size:.8rem;color:var(--text3);margin-bottom:16px">Mon–Fri this week · Only unset days</div>' +
      '<div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px;margin-bottom:14px">' +
      STATUSES.map(s => {
        const cfg = STATUS_CFG[s]||{};
        return '<button data-action="week-picker-apply" data-val="' + s + '" style="padding:10px 8px;border-radius:10px;border:2px solid ' + cfg.color + ';background:' + cfg.bg + ';color:' + cfg.color + ';font-weight:700;font-size:.85rem;cursor:pointer;font-family:var(--font)">' + s + '</button>';
      }).join('') +
      '</div>' +
      '<div style="display:flex;justify-content:space-between;align-items:center">' +
      '<label style="font-size:.78rem;color:var(--text3);display:flex;align-items:center;gap:6px"><input type="checkbox" id="wk-overwrite" style="accent-color:var(--accent2)"> Overwrite already-marked days</label>' +
      '<button data-action="week-picker-cancel" style="padding:5px 12px;border-radius:7px;border:1px solid var(--border);background:var(--bg3);color:var(--text2);cursor:pointer;font-family:var(--font);font-size:.82rem">Cancel</button>' +
      '</div>';
    document.getElementById('dtr-root-outer').appendChild(picker);
    // Close on outside click
    setTimeout(() => {
      document.addEventListener('click', function dismiss(e) {
        if (!picker.contains(e.target)) { picker.remove(); document.removeEventListener('click', dismiss); }
      });
    }, 100);

    // week-picker-apply handled in handleClick
    // Build calendar grid cells
    const dayHeaders = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
    let cells = '';
    // Empty cells before first day
    for (let i = 0; i < firstDay; i++) {
      cells += '<div class="cal-cell cal-empty"></div>';
    }
    for (let d = 1; d <= totalDays; d++) {
      const dk  = year + '-' + String(month+1).padStart(2,'0') + '-' + String(d).padStart(2,'0');
      const dow = new Date(year, month, d).getDay();
      const isWeekend = dow === 0 || dow === 6;
      const isToday   = dk === today;
      const isFuture  = dk > today;
      const entry = statusCache[dk] || null;
      const st    = entry ? entry.status : '';
      const color = st ? STATUS_COLOR[st] : '';
      const short = st ? STATUS_SHORT[st] : '';
      cells += '<div class="cal-cell' +
        (isWeekend ? ' cal-weekend' : '') +
        (isToday   ? ' cal-today'   : '') +
        (isFuture  ? ' cal-future'  : '') +
        (st        ? ' cal-marked'  : '') +
        '"' +
        (st ? ' style="border-color:' + color + ';background:' + color + '18"' : '') +
        ' data-action="cal-day" data-dk="' + dk + '">' +
        '<div class="cal-day-num' + (isToday?' cal-today-num':'') + '">' + d + '</div>' +
        (st
          ? '<div class="cal-day-badge" style="background:' + color + '">' + short + '</div>'
          : (isWeekend ? '<div class="cal-day-wknd">—</div>' : '<div class="cal-day-empty"></div>')
        ) +
        '</div>';
    }

    el.innerHTML =
      '<div class="ph"><div class="ph-left"><div class="ph-title">My Attendance Calendar</div>' +
      '<div class="ph-sub">Your personal attendance history — click any day to mark or edit</div></div>' +
      '<div class="ph-actions"><button class="btn btn-ghost btn-sm" data-action="cal-sync">' + ic.sync + ' Sync</button></div></div>' +

      // Month navigator
      '<div class="wv-controls" style="margin-bottom:16px">' +
      '<button class="wv-nav-btn" data-action="cal-prev">' + ic.left + '</button>' +
      '<div class="wv-range"><div class="wv-range-title">' + monthName + '</div>' +
      '<div class="wv-range-sub">' + markedDays + ' of ' + workingDays + ' working days marked</div></div>' +
      (calMonthOffset !== 0 ? '<button class="wv-today-btn" data-action="cal-today">This Month</button>' : '') +
      '<button class="wv-nav-btn" data-action="cal-next">' + ic.right + '</button>' +
      '</div>' +

      // Stats row
      '<div class="stats-grid sg4" style="margin-bottom:16px">' +
      '<div class="stat-card gb"><div class="lbl">WFO Days</div><div class="val" style="color:#3fb950">' + wfoCnt + '</div></div>' +
      '<div class="stat-card ab"><div class="lbl">WFH Days</div><div class="val" style="color:#22d3ee">' + wfhCnt + '</div></div>' +
      '<div class="stat-card amb"><div class="lbl">Leaves</div><div class="val" style="color:var(--amber)">' + (slCnt+clCnt+alCnt) + '</div><div class="sub">SL ' + slCnt + ' · CL ' + clCnt + ' · AL ' + alCnt + '</div></div>' +
      '<div class="stat-card pb"><div class="lbl">Optional Off</div><div class="val" style="color:var(--text3)">' + ooCnt + '</div></div>' +
      '</div>' +

      // Calendar legend
      '<div style="display:flex;gap:12px;flex-wrap:wrap;margin-bottom:12px;align-items:center">' +
      Object.entries(STATUS_COLOR).map(([s,c]) =>
        '<div style="display:flex;align-items:center;gap:5px"><div style="width:10px;height:10px;border-radius:3px;background:' + c + '"></div>' +
        '<span style="font-size:.78rem;color:var(--text2);font-weight:600">' + s + '</span></div>'
      ).join('') + '</div>' +

      // Calendar grid
      '<div class="cal-grid-wrap">' +
      '<div class="cal-header">' + dayHeaders.map(d => '<div class="cal-hcell">' + d + '</div>').join('') + '</div>' +
      '<div class="cal-grid">' + cells + '</div>' +
      '</div>' +

      // Day edit panel (shown on click)
      '<div id="cal-edit-panel" style="margin-top:14px"></div>';

    // Bind day clicks
    const grid = el.querySelector('.cal-grid');
    if (grid) {
      grid.addEventListener('click', e => {
        const cell = e.target.closest('[data-action="cal-day"]');
        if (!cell) return;
        const dk = cell.dataset.dk;
        if (!dk) return;
        // Set attDate and switch to mark tab's edit panel inline
        attDate = dk;
        attSelectedStatus = (statusCache[dk]||{}).status||'';
        const panel = q('#cal-edit-panel');
        if (panel) panel.innerHTML = buildAttEditPanel(dk);
        // highlight selected cell
        el.querySelectorAll('.cal-cell').forEach(c => c.classList.toggle('cal-cell-active', c.dataset.dk === dk));
      });
    }
  }


  // MISSED NPT
  const NPT_TYPES = ['System Issue','Meeting Overrun','Training','Lack of Work','Power Outage','Network Issue','Admin Task','Other'];
  let nptActiveType = '';

  function renderMissedNPT() {
    const el = q('#view-missednpt'); if (!el) return;
    const myNPT = nptCache.filter(n => n.name === authState.name);
    const totalMins = myNPT.reduce((a,n)=>a+(n.minutes||0),0);
    const h = Math.floor(totalMins/60), m = totalMins%60;
    const thisMonth = new Date().toISOString().slice(0,7);
    const monthMins = myNPT.filter(n=>n.date&&n.date.startsWith(thisMonth)).reduce((a,n)=>a+(n.minutes||0),0);
    el.innerHTML =
      '<div class="ph"><div class="ph-left"><div class="ph-title">Missed NPT Log</div><div class="ph-sub">Record non-productive time not in daily tasks</div></div></div>' +
      '<div class="npt-summary-cards">' +
      '<div class="npt-sum-card"><div class="lbl">Entries</div><div class="val">' + myNPT.length + '</div></div>' +
      '<div class="npt-sum-card"><div class="lbl">Total Time</div><div class="val">' + (h>0?h+'h ':'')+m+'m</div></div>' +
      '<div class="npt-sum-card"><div class="lbl">This Month</div><div class="val">' + Math.floor(monthMins/60)+'h '+monthMins%60+'m</div></div></div>' +
      '<div class="npt-log-form"><div class="npt-log-form-title">Log Missed NPT</div>' +
      '<div class="npt-types">' + NPT_TYPES.map(t => '<button class="npt-type-btn' + (nptActiveType===t?' active':'') + '" data-action="npt-type" data-val="' + t + '">' + t + '</button>').join('') + '</div>' +
      '<div class="g3"><div class="dtr-field"><label class="dtr-label">Date</label><input type="date" class="dtr-input" id="npt-date" value="' + todayStr() + '"></div>' +
      '<div class="dtr-field"><label class="dtr-label">Duration (minutes)</label><input type="number" class="dtr-input" id="npt-mins" min="1" max="480" placeholder="e.g. 30"></div>' +
      '<div class="dtr-field" style="justify-content:flex-end"><button class="btn btn-primary" data-action="npt-log" style="align-self:flex-end">' + ic.add + ' Log Entry</button></div></div>' +
      '<div class="dtr-field"><label class="dtr-label">Description</label><input type="text" class="dtr-input" id="npt-desc" placeholder="Brief description..."></div></div>' +
      '<div class="card"><div class="card-title">My NPT Log</div>' +
      (myNPT.length ?
        '<div class="npt-table-wrap"><table class="npt-table"><thead><tr><th>Date</th><th>Type</th><th>Duration</th><th>Description</th><th></th></tr></thead><tbody>' +
        myNPT.slice().reverse().map(n => { const ri=nptCache.indexOf(n),mins=n.minutes||0,dh=Math.floor(mins/60),dm=mins%60; return '<tr><td>' + n.date + '</td><td><span class="badge bb">' + (n.type||'Other') + '</span></td><td class="mono">' + (dh>0?dh+'h ':'')+dm+'m</td><td style="max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="' + (n.desc||'').replace(/"/g,'&quot;') + '">' + (n.desc||'—') + '</td><td><button class="btn btn-danger btn-xs" data-action="npt-del" data-idx="' + ri + '">' + ic.trash + '</button></td></tr>'; }).join('') +
        '</tbody></table></div>' :
        '<div class="empty"><p>No NPT entries yet. Log your first one above.</p></div>') +
      '</div>';
  }

  function nptSelectType(type) { nptActiveType = type; qa('.npt-type-btn').forEach(b => b.classList.toggle('active', b.dataset.val === type)); }

  async function nptLogEntry() {
    const date=q('#npt-date')?.value||todayStr(), mins=parseInt(q('#npt-mins')?.value||'0'), desc=q('#npt-desc')?.value?.trim()||'', type=nptActiveType;
    if (!type) { toast('Select an NPT type','err'); return; }
    if (!mins||mins<1) { toast('Enter duration in minutes','err'); return; }
    const entry = { name:authState.name, date, type, minutes:mins, desc, loggedAt:new Date().toISOString() };
    nptCache.push(entry); safeSave('dtr_npt2', nptCache);
    const sent = await postNPT(entry);
    toast(sent ? '✅ NPT logged — '+type+' ('+mins+'m)' : 'Saved locally (SP sync pending)', sent?'ok':'info');
    nptActiveType = ''; renderMissedNPT();
  }

  function nptDel(idx) { if (!confirm('Delete this NPT entry?')) return; nptCache.splice(idx,1); safeSave('dtr_npt2',nptCache); renderMissedNPT(); }

  // SETTINGS
  function renderSettings() {
    const el = q('#view-settings'); if (!el) return;
    const qLen = (Array.isArray(spQueue)?spQueue:[]).length;
    el.innerHTML =
      '<div class="ph"><div class="ph-left"><div class="ph-title">Settings</div></div></div>' +
      '<div class="card"><div class="card-title">' + ic.sync + ' SharePoint — Single Site</div>' +
      '<div class="info-banner" style="margin-bottom:12px">All data goes to: <strong>' + SP.SITE + '</strong></div>' +
      '<div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:10px">' +
      '<button class="btn btn-ghost btn-sm" data-action="test-sp">Test Connection</button>' +
      (qLen>0?'<button class="btn btn-ghost btn-sm" data-action="flush-queue">🔄 Retry '+qLen+' Queued</button>':'') + '</div>' +
      '<p style="font-size:.8rem;color:var(--text3)">' + (qLen===0?'✅ All synced':qLen+' pending sync') + '</p>' +
      '<div style="margin-top:14px"><div style="font-size:.78rem;color:var(--text3)"><strong>SharePoint Lists Required:</strong></div>' +
      '<div style="margin-top:6px;font-size:.78rem;color:var(--text3);line-height:2">' +
      '📋 Task List: <code style="color:var(--accent)">' + SP.TASK_LIST + '</code><br>' +
      '📅 Attendance List: <code style="color:var(--accent)">' + SP.STATUS_LIST + '</code> (columns: EmployeeName, StatusDate, WorkStatus, Process, TaskNotes, UpdatedAt)<br>' +
      '⏱ NPT List: <code style="color:var(--accent)">' + SP.NPT_LIST + '</code> (columns: EmployeeName, NPTDate, NPTType, DurationMins, Description, LoggedAt)</div></div></div>' +
      '<div class="card"><div class="card-title">🗑 My Data</div>' +
      '<div style="font-size:.82rem;color:var(--text2);margin-bottom:10px">' +
      mySubmissions().length + ' task entries · ' + Object.keys(statusCache).length + ' attendance days · ' + nptCache.filter(n=>n.name===authState.name).length + ' NPT entries</div>' +
      '<button class="btn btn-danger btn-sm" onclick="if(confirm(\'Clear all YOUR local data?\')){' +
      'submissions=submissions.filter(s=>s.employeeName!==authState.name);safeSave(\'dtr_subs4\',submissions);' +
      'GM_setValue(\'dtr_status2\',\'{}\');GM_setValue(\'dtr_npt2\',\'[]\');toast(\'Cleared\',\'info\');}">Clear My Local Data</button></div>' +
      '<div class="card"><div class="card-title">👤 Session</div>' +
      '<div style="font-size:.85rem;color:var(--text2);margin-bottom:10px">Logged in as: <strong style="color:var(--text)">' + authState.name + '</strong></div>' +
      '<button class="btn btn-ghost btn-sm" data-action="do-logout">Log out</button></div>';
  }

  // SP FUNCTIONS
  // ═══════════════════════════════════════════════════════════════════
  // SHAREPOINT — robust posting with queue fallback
  // ═══════════════════════════════════════════════════════════════════

  async function getDigest() {
    return new Promise(resolve => {
      GM_xmlhttpRequest({
        method: 'POST',
        url: SP.SITE + '/_api/contextinfo',
        headers: {
          'Accept': 'application/json;odata=verbose',
          'Content-Type': 'application/json;odata=verbose'
        },
        withCredentials: true,
        onload: res => {
          try {
            if (res.status === 401 || res.status === 403) {
              toast('❌ SP auth error (' + res.status + ') — make sure you are logged into SharePoint in this browser', 'err');
              resolve(null); return;
            }
            const data = JSON.parse(res.responseText);
            const token = data.d.GetContextWebInformation.FormDigestValue;
            if (!token) { toast('❌ SP token empty — check SP site URL', 'err'); resolve(null); return; }
            console.log('[WorkPulse] SP auth token OK, length:', token.length);
            resolve(token);
          } catch(ex) {
            toast('❌ SP contextinfo failed (status ' + res.status + ') — ' + res.responseText.slice(0,80), 'err');
            resolve(null);
          }
        },
        onerror: (err) => {
          toast('❌ Cannot reach SharePoint — network error. Open SP site in a tab first.', 'err');
          resolve(null);
        }
      });
    });
  }

  async function spInsert(listName, metaType, body) {
    const token = await getDigest();
    if (!token) return false;
    return new Promise(resolve => {
      const fullBody = Object.assign({ '__metadata': { 'type': metaType } }, body);
      const payload = JSON.stringify(fullBody);
      GM_xmlhttpRequest({
        method: 'POST',
        url: SP.SITE + "/_api/web/lists/GetByTitle('" + listName + "')/items",
        headers: {
          'Accept': 'application/json;odata=verbose',
          'Content-Type': 'application/json;odata=verbose',
          'X-RequestDigest': token
        },
        data: payload,
        withCredentials: true,
        onload: res => {
          if (res.status >= 200 && res.status < 300) {
            resolve(true);
          } else {
            // Show full SP error in console AND toast
            console.error('[WorkPulse] SP insert failed:', res.status, res.responseText);
            let msg = '';
            try {
              const j = JSON.parse(res.responseText);
              msg = (j.error && j.error.message) ? (j.error.message.value || j.error.message) : res.responseText.slice(0,120);
            } catch { msg = res.responseText.slice(0, 120); }
            toast('❌ [' + listName + '] ' + res.status + ': ' + msg.slice(0, 100), 'err');
            resolve(false);
          }
        },
        onerror: (e) => {
          console.error('[WorkPulse] SP network error:', e);
          toast('❌ SP network error — check console (F12)', 'err');
          resolve(false);
        }
      });
    });
  }

  async function spUpdate(listName, itemId, metaType, body) {
    const token = await getDigest();
    if (!token) return false;
    return new Promise(resolve => {
      const fullBody = Object.assign({ '__metadata': { 'type': metaType } }, body);
      GM_xmlhttpRequest({
        method: 'POST',
        url: SP.SITE + "/_api/web/lists/GetByTitle('" + listName + "')/items(" + itemId + ")",
        headers: {
          'Accept': 'application/json;odata=verbose',
          'Content-Type': 'application/json;odata=verbose',
          'X-RequestDigest': token,
          'X-HTTP-Method': 'MERGE',
          'If-Match': '*'
        },
        data: JSON.stringify(fullBody),
        withCredentials: true,
        onload: res => {
          if (res.status >= 200 && res.status < 300) { resolve(true); }
          else {
            let msg = '';
            try { msg = JSON.parse(res.responseText).error.message.value || ''; } catch {}
            toast('❌ SP update failed (' + res.status + ')' + (msg ? ' — ' + msg.slice(0, 80) : ''), 'err');
            resolve(false);
          }
        },
        onerror: () => { toast('❌ SP network error during update', 'err'); resolve(false); }
      });
    });
  }

  // Keep spPost for postTask compatibility
  async function spPost(listName, body) {
    const token = await getDigest(); if (!token) return false;
    return new Promise(r => {
      GM_xmlhttpRequest({
        method: 'POST',
        url: SP.SITE + "/_api/web/lists/GetByTitle('" + listName + "')/items",
        headers: { 'Accept': 'application/json;odata=verbose', 'Content-Type': 'application/json;odata=verbose', 'X-RequestDigest': token },
        data: JSON.stringify(body), withCredentials: true,
        onload: res => r(res.status >= 200 && res.status < 300),
        onerror: () => r(false)
      });
    });
  }

  async function postTask(t) {
    return spPost(SP.TASK_LIST, {
      '__metadata': { 'type': 'SP.Data.' + SP.TASK_LIST + 'ListItem' },
      'EmployeeName': t.employeeName || '', 'TaskType': t.taskType || '',
      'HoursWorked': t.hours || 0, 'NPTHours': t.npt || 0,
      'WorkType': t.workType || 'Productive', 'AdHocDetails': t.adhoc || '',
      'TaskDate': t.date || '', 'SubmittedAt': t.submittedAt || new Date().toISOString()
    });
  }

  async function postNPT(e) {
    const token = await getDigest();
    if (!token) return false;
    const dateStr = e.date || todayStr();
    const title   = (e.type || 'NPT') + ' - ' + dateStr;
    // Build body WITHOUT __metadata — avoids type name mismatch issues entirely
    // SP REST API v1 accepts this format fine
    const body = {
      'Title':        title,
      'EmployeeName': e.name || '',
      'NPTDate':      dateStr + 'T00:00:00Z',
      'NPTType':      e.type || '',
      'DurationMins': parseInt(e.minutes) || 0,
      'Description':  e.desc || '',
      'LoggedAt':     e.loggedAt || new Date().toISOString()
    };
    console.log('[WorkPulse] Posting NPT:', JSON.stringify(body));
    return new Promise(resolve => {
      GM_xmlhttpRequest({
        method: 'POST',
        url: SP.SITE + "/_api/web/lists/GetByTitle('" + SP.NPT_LIST + "')/items",
        headers: {
          'Accept':           'application/json;odata=verbose',
          'Content-Type':     'application/json;odata=verbose',
          'X-RequestDigest':  token
        },
        data: JSON.stringify(body),
        withCredentials: true,
        onload: res => {
          console.log('[WorkPulse] NPT response:', res.status, res.responseText.slice(0, 300));
          if (res.status >= 200 && res.status < 300) {
            resolve(true);
          } else {
            // Try again with __metadata using the actual list type name from SP
            let metaType = '';
            try {
              // Extract real type from error or fetch it
              metaType = JSON.parse(res.responseText).error.message.value || '';
            } catch {}
            toast('❌ NPT SP error ' + res.status + ' — open F12 console for details', 'err');
            console.error('[WorkPulse] NPT full error:', res.responseText);
            resolve(false);
          }
        },
        onerror: err => {
          console.error('[WorkPulse] NPT network error:', err);
          toast('❌ NPT network error', 'err');
          resolve(false);
        }
      });
    });
  }

  async function postAttendance(e) {
    const spDate  = (e.date || todayStr()) + 'T00:00:00Z';
    const title   = (e.name || '') + ' - ' + (e.date || todayStr());
    const nameEsc = (e.name || '').replace(/'/g, "''");
    // Check if record exists for this person+date
    return new Promise(resolve => {
      GM_xmlhttpRequest({
        method: 'GET',
        url: SP.SITE + "/_api/web/lists/GetByTitle('" + SP.STATUS_LIST + "')/items" +
             "?$filter=EmployeeName eq '" + nameEsc + "' and StatusDate eq datetime'" + spDate + "'&$select=Id&$top=1",
        headers: { 'Accept': 'application/json;odata=verbose' },
        withCredentials: true,
        onload: async res => {
          try {
            const results = (JSON.parse(res.responseText).d || {}).results || [];
            const body = {
              'Title':        title,
              'EmployeeName': e.name || '',
              'StatusDate':   spDate,
              'WorkStatus':   e.status || '',
              'Process':      e.process || '',
              'TaskNotes':    e.task || '',
              'UpdatedAt':    e.updatedAt || new Date().toISOString()
            };
            const meta = 'SP.Data.AttendanceStatusListItem';
            if (results.length > 0) {
              resolve(await spUpdate(SP.STATUS_LIST, results[0].Id, meta, body));
            } else {
              resolve(await spInsert(SP.STATUS_LIST, meta, body));
            }
          } catch(ex) {
            toast('❌ Attendance lookup error: ' + ex.message, 'err');
            resolve(false);
          }
        },
        onerror: () => { toast('❌ Cannot reach SP for attendance check', 'err'); resolve(false); }
      });
    });
  }

  async function flushQueue() {
    const q2 = safeLoad('dtr_spq2', []); if (!q2.length) return;
    const rem = [];
    for (const t of q2) { const s = await postTask(t); if (!s) rem.push(t); }
    safeSave('dtr_spq2', rem);
  }

  async function flushQueueManual() {
    toast('Retrying queued items...', 'info');
    await flushQueue();
    const r = safeLoad('dtr_spq2', []).length;
    toast(r === 0 ? '✅ All synced' : '⚠️ ' + r + ' still pending', r === 0 ? 'ok' : 'err');
  }

  async function testSP() {
    toast('Testing SP auth...', 'info');
    const token = await getDigest();
    if (!token) return; // getDigest already showed error toast
    toast('✅ Auth OK — checking lists...', 'ok');
    setTimeout(() => {
      GM_xmlhttpRequest({
        method: 'GET',
        url: SP.SITE + "/_api/web/lists/GetByTitle('" + SP.NPT_LIST + "')?$select=Title,ListItemEntityTypeFullName,ItemCount",
        headers: { 'Accept': 'application/json;odata=verbose' }, withCredentials: true,
        onload: res => {
          try {
            const d = JSON.parse(res.responseText).d;
            toast('✅ NPTLog found: ' + d.ItemCount + ' rows | type: ' + d.ListItemEntityTypeFullName, 'ok');
            console.log('[WorkPulse] NPTLog ListItemEntityTypeFullName:', d.ListItemEntityTypeFullName);
          } catch { toast('❌ NPTLog not found — is list name exactly "NPTLog"?', 'err'); }
        }, onerror: () => toast('❌ Cannot reach NPTLog', 'err')
      });
    }, 1200);
    setTimeout(() => {
      GM_xmlhttpRequest({
        method: 'GET',
        url: SP.SITE + "/_api/web/lists/GetByTitle('" + SP.STATUS_LIST + "')/fields?$filter=Hidden eq false and ReadOnlyField eq false&$select=InternalName",
        headers: { 'Accept': 'application/json;odata=verbose' }, withCredentials: true,
        onload: res => {
          try {
            const cols = (JSON.parse(res.responseText).d.results || []).map(f => f.InternalName).join(', ');
            toast('✅ AttendanceStatus cols: ' + cols, 'ok');
          } catch { toast('❌ AttendanceStatus list not found', 'err'); }
        }, onerror: () => toast('❌ Cannot read AttendanceStatus list', 'err')
      });
    }, 2500);
  }

  function fetchOwnAttendance(silent) {
    // Fetch only THIS user's attendance records from SP → populate statusCache
    if (!authState.name) return;
    const nameEsc = authState.name.replace(/'/g,"''");
    const url = SP.SITE + "/_api/web/lists/GetByTitle('" + SP.STATUS_LIST + "')/items" +
      "?$filter=EmployeeName eq '" + nameEsc + "'" +
      "&$top=5000&$orderby=StatusDate desc" +
      "&$select=EmployeeName,StatusDate,WorkStatus,Process,TaskNotes";
    GM_xmlhttpRequest({
      method:'GET', url:url,
      headers:{'Accept':'application/json;odata=verbose'},
      withCredentials:true,
      onload: res => {
        try {
          const items = JSON.parse(res.responseText).d.results || [];
          items.forEach(it => {
            const rawDate = it.StatusDate || '';
            const dk = rawDate.includes('T') ? rawDate.split('T')[0] : rawDate;
            if (dk) {
              // Only update if SP has a value — don't overwrite newer local data
              if (!statusCache[dk] || statusCache[dk]._fromSP) {
                statusCache[dk] = {
                  status:  it.WorkStatus || '',
                  process: it.Process    || '',
                  task:    it.TaskNotes  || '',
                  _fromSP: true
                };
              }
            }
          });
          safeSaveObj('dtr_status2', statusCache);
          if (!silent) toast('✅ Attendance loaded from SharePoint','ok');
          // Re-render whichever view is active
          if (currentView === 'calendar')  renderMyCalendar();
          if (currentView === 'mark')      renderMarkAttendance();
          if (currentView === 'analytics') renderAnalytics();
          // Update topbar status badge if today is now known
          const todaySt = statusCache[todayStr()];
          const pill = root.querySelector('.dtr-topbar-r .sp');
          if (pill && todaySt) {
            pill.className = 'sp sp-' + todaySt.status.toLowerCase().replace(' ','-');
            pill.textContent = todaySt.status;
          }
        } catch(ex) {
          if (!silent) toast('❌ Attendance sync error: ' + ex.message,'err');
        }
      },
      onerror: () => { if (!silent) toast('❌ Cannot reach SharePoint','err'); }
    });
  }

  function fetchAllAttendance() {
    // Fetch all team attendance (for team cache) + own attendance (for calendar)
    fetchOwnAttendance(false);
    GM_xmlhttpRequest({method:'GET',url:SP.SITE+"/_api/web/lists/GetByTitle('"+SP.STATUS_LIST+"')/items?$top=5000&$orderby=StatusDate%20desc&$select=EmployeeName,StatusDate,WorkStatus,Process,TaskNotes",headers:{'Accept':'application/json;odata=verbose'},withCredentials:true,
    onload:res=>{try{
      const items=JSON.parse(res.responseText).d.results||[];
      teamStatusCache={};
      items.forEach(it=>{const name=it.EmployeeName||'';const rawDate=it.StatusDate||'';const dk=rawDate.includes('T')?rawDate.split('T')[0]:rawDate;if(name&&dk)teamStatusCache[name+'::'+dk]={status:it.WorkStatus||'',process:it.Process||'',task:it.TaskNotes||''};});
      Object.entries(statusCache).forEach(([dk,v])=>{teamStatusCache[authState.name+'::'+dk]={status:v.status,process:v.process||'',task:v.task||''};});
      safeSaveObj('dtr_teamcache',teamStatusCache);
    }catch{}},onerror:()=>{}});
  }
  // flushQueue defined above
  // flushQueueManual defined above


  // DATA HELPERS
  function getAllUsersFromCache() { const tc=safeLoadObj('dtr_teamcache',{});if(Object.keys(tc).length)teamStatusCache=tc;const names=new Set();Object.keys(teamStatusCache).forEach(k=>{const[name]=k.split('::');if(name)names.add(name);});names.add(authState.name);return[...names].sort(); }
  function getTeamStatus(user,dk)  { return(teamStatusCache[user+'::'+dk]||{}).status ||(user===authState.name?((statusCache[dk]||{}).status||''):''); }
  function getTeamProcess(user,dk) { return(teamStatusCache[user+'::'+dk]||{}).process||(user===authState.name?((statusCache[dk]||{}).process||''):''); }
  function getTeamTask(user,dk)    { return(teamStatusCache[user+'::'+dk]||{}).task   ||(user===authState.name?((statusCache[dk]||{}).task||''):''); }
  function mySubmissions() { return Array.isArray(submissions)?submissions.filter(s=>s&&s.employeeName===authState.name):[]; }
  function calcAvgProd(subs){ if(!Array.isArray(subs))return 0;const v=subs.filter(s=>s&&!s.taskType?.startsWith('Leave'));if(!v.length)return 0;return Math.min(100,v.reduce((a,s)=>a+((s.hours||0)/WH*100),0)/v.length); }
  function calcStreak(subs){ const days=[...new Set(subs.filter(s=>s&&!s.taskType?.startsWith('Leave')).map(s=>s.date))].sort().reverse();let streak=0,prev=new Date();for(const d of days){const dt=new Date(d+'T12:00:00'),diff=Math.round((prev-dt)/86400000);if(streak===0&&diff<=1){streak=1;prev=dt;}else if(diff===1){streak++;prev=dt;}else break;}return streak; }
  function getLast7(){ return Array.from({length:7},(_,i)=>{const d=new Date();d.setDate(d.getDate()-6+i);return d.toISOString().split('T')[0];}); }
  function getWeekStart(offset){ const d=new Date();d.setHours(0,0,0,0);d.setDate(d.getDate()-d.getDay()+(offset*7));return d; }
  function getWeekLabel(offset){ const ws=getWeekStart(offset),we=new Date(ws);we.setDate(ws.getDate()+6);const f=d=>d.toLocaleDateString('en-US',{month:'short',day:'numeric'});return f(ws)+' – '+f(we)+', '+we.getFullYear(); }
  function todayStr()     { return new Date().toISOString().split('T')[0]; }
  function yesterdayStr() { const d=new Date();d.setDate(d.getDate()-1);return d.toISOString().split('T')[0]; }
  function isAllowedDate(s){ return s===todayStr()||s===yesterdayStr(); }
  function formatDate(d)  { return new Date(d+'T12:00:00').toLocaleDateString('en-US',{month:'short',day:'numeric'}); }
  function formatDay(d)   { return new Date(d+'T12:00:00').toLocaleDateString('en-US',{weekday:'long',year:'numeric',month:'long',day:'numeric'}); }
  function dlCSV(csv,fn)  { const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([csv],{type:'text/csv'}));a.download=fn;a.click(); }
  function updateSBStats(){ const ms=mySubmissions();const s1=q('#sb-total'),s2=q('#sb-avg');if(s1)s1.textContent=ms.length;if(s2)s2.textContent=calcAvgProd(ms).toFixed(0)+'%'; }
  function toast(msg,type='info'){ let el=document.getElementById('dtr-toast');if(!el){el=document.createElement('div');el.id='dtr-toast';(document.getElementById('dtr-root-outer')||root).appendChild(el);}const icons={ok:'✅',err:'❌',info:'💡'};el.innerHTML='<span>'+(icons[type]||'💡')+'</span><span>'+msg+'</span>';el.className='show t'+type;clearTimeout(el._t);el._t=setTimeout(()=>{el.className='';},4000); }

  // INIT
  teamStatusCache = safeLoadObj('dtr_teamcache', {});
  if (document.readyState==='loading') { document.addEventListener('DOMContentLoaded', boot); } else { boot(); }

})();
