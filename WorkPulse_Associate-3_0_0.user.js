// ==UserScript==
// @name         WorkPulse — Universal v4
// @namespace    https://amazon.sharepoint.com/sites/teamdailytask/
// @version      4.0.0
// @description  WorkPulse — SSO, role-based access (Associate/Admin/SuperAdmin), SP direct sync
// @author       WorkPulse Team
// @match        https://share.amazon.com/Pages/default.aspx
// @grant        GM_xmlhttpRequest
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        GM_addStyle
// @connect      share.amazon.com
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
    TASK_LIST:   'TaskReport',            // new SP list
    STATUS_LIST: 'AttendanceLog',        // Attendance list
    NPT_LIST:    'NPT log',               // SharePoint list name (with space)
  };

  // ═══════════════════════════════════════════════════════════════════════
  // CONSTANTS
  // ═══════════════════════════════════════════════════════════════════════
  const WH     = 480; // total working minutes per day
  const TASK_TYPES = [
    'CT Production','Pre-Prod Production','Simulator Production',
    'CT Audits','Pre-Prod Audits','Simulator Audits','Lack of Work',
    'Meeting','Ad-hoc Tasks','Other'
  ];
  const NPT_TASKS  = ['Lack of Work'];
  const NPT_TYPES  = [
    'System Issue','Meeting Overrun','Training','Lack of Work',
    'Power Outage','Network Issue','Admin Task','Other'
  ];
  let nptActiveType = '';
  const STATUSES   = ['WFO','WFH','SL','CL','AL','Optional Off'];
  // ── Authorised users ──────────────────────────────────────────────────
  const ASSOCIATES = [
    'Ameralih','Bhurak','Pmred','Harikavr','Zshahnaz',
    'Meguvval','Vpulluri','Piyushts','Kuparima','Unairite',
    'Anshdeep','Gsridev','Arvindon','Psiranga','Varmana',
    'Haranbhe','Kalakuh','Alekhyya','Pankae','Madhureg',
    'Bsv','Heswitha','Mbahyal','Bhanupru','Harusn',
    'Remoch','Siqmadhu','Rajawbab','Inagajag',
    'Malsrira','Edharapa','Hshyaraj','Sofiykja','Rayyanms',
    'Kenumula','Tumkurs',
    'Mppunna','Joldapka'
  ]; // 42 associates

  const ADMINS      = ['Tkattula','Chaturay','Shamils','Nkandhur','Sherylv','Sheebyme','Awaispsh','Dvsanjay']; // 8 admins
  const SUPER_ADMIN = 'Tkattula'; // can access both admin + associate (Admin-tarun) views

  // Runtime role — set after SSO/login
  let currentRole = 'associate'; // 'associate' | 'admin'
  const PROCS      = ['Preprod Testing','Chat Transcripts','Adhoc','Quality Check','Training','Other'];
  const COLORS     = ['#6366f1','#8b5cf6','#06b6d4','#10b981','#f59e0b','#ef4444','#ec4899','#84cc16'];
  const STATUS_CFG = {
    WFO: { label:'Work From Office', cls:'sp-wfo', color:'#3fb950', bg:'rgba(63,185,80,.14)'  },
    WFH: { label:'Work From Home',   cls:'sp-wfh', color:'#22d3ee', bg:'rgba(6,182,212,.14)'  },
    SL:  { label:'Sick Leave',       cls:'sp-sl',  color:'#f85149', bg:'rgba(248,81,73,.14)'  },
    CL:  { label:'Casual Leave',     cls:'sp-cl',  color:'#d29922', bg:'rgba(210,153,34,.14)' },
    AL:  { label:'Annual Leave',     cls:'sp-al',  color:'#a371f7', bg:'rgba(163,113,247,.14)'},
    'Optional Off':   { label:'Optional Off',  cls:'sp-oo',  color:'#6b7280', bg:'rgba(107,114,128,.14)' },
    'Mandatory Off':  { label:'National Holiday', cls:'sp-mo', color:'#e8433a', bg:'rgba(232,67,58,.12)' },
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
  // Per-user shift preference — persists across refresh/logout
  function saveShift(name, shift) {
    if (!name || !shift) return;
    GM_setValue('dtr_shift_' + name.toLowerCase(), shift);
  }
  function loadShift(name) {
    if (!name) return '8-5';
    return GM_getValue('dtr_shift_' + name.toLowerCase(), '8-5') || '8-5';
  }
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

  let currentView    = '';
  let taskCounter    = 1;
  let selectedDate   = todayStr();
  let weekOffset     = 0;
  let attWeekOffset  = 0;
  let calMonthOffset = 0;
  let attDate        = todayStr();
  let attBulkMode    = false;
  let attMultiSelect = new Set();
  // Admin state
  let adminWeekOffset   = 0;
  let admSelectedAssocs  = new Set();
  let _deviceChosen = false; // tracks if device picker was shown this session
  let aanSelectedAssoc   = '';
  let aanMode            = 'single';
  let aanCompareSet      = new Set(); // multi-select for team tracker
  let teamStatusCache_adm = safeLoadObj('dtr_teamcache',{});
  let nptAllCache       = safeLoad('dtr_npt2',[]);
  let adminTkCustomCols = safeLoad('adm_tkcols',[]);
  let _listTypes        = { task:'', npt:'', att:'' };
  let root;

  // ═══════════════════════════════════════════════════════════════════════
  // STYLES
  // ═══════════════════════════════════════════════════════════════════════
  GM_addStyle(`
    @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700&family=DM+Mono:wght@400;500&display=swap');
    *{box-sizing:border-box;margin:0;padding:0}
    /* Dark theme (default) — vars scoped to #dtr-root so .lt override works */
    #dtr-root{
      --bg:rgba(12,14,20,1);
      --bg2:rgba(255,255,255,0.05);
      --bg3:rgba(255,255,255,0.08);
      --bg4:rgba(255,255,255,0.12);
      --bg-solid:#0c0e14;
      --bg3-solid:#1c2333;
      --border:rgba(255,255,255,0.1);
      --border2:rgba(255,255,255,0.18);
      --glass:rgba(255,255,255,0.06);
      --glass2:rgba(255,255,255,0.04);
      --glass-border:rgba(255,255,255,0.12);
      --text:#f0f4ff;--text2:rgba(240,244,255,0.6);--text3:rgba(240,244,255,0.35);
      --accent:#4f9eff;--accent2:#a371f7;
      --green:#30d158;--amber:#ffd60a;--red:#ff453a;--purple:#bf5af2;
      --shadow:0 8px 32px rgba(0,0,0,.6),0 2px 8px rgba(0,0,0,.4);
      --shadow-sm:0 2px 12px rgba(0,0,0,.4);
      --radius:12px;--font:'DM Sans',sans-serif;--mono:'DM Mono',monospace
    }
    #dtr-root.lt{
      --bg:rgba(242,242,247,1);
      --bg2:rgba(255,255,255,0.85);
      --bg3:rgba(255,255,255,0.65);
      --bg4:rgba(255,255,255,0.5);
      --bg-solid:#f2f2f7;
      --bg3-solid:#f0f3f6;
      --bg4:#e5e5ea;  /* Solid track bg in light mode */
      --border:rgba(0,0,0,0.08);
      --border2:rgba(0,0,0,0.14);
      --glass:rgba(255,255,255,0.7);
      --glass2:rgba(255,255,255,0.5);
      --glass-border:rgba(0,0,0,0.1);
      --text:#1c1c1e;--text2:rgba(28,28,30,0.55);--text3:rgba(28,28,30,0.35);
      --accent:#007aff;--accent2:#5856d6;
      --green:#34c759;--amber:#8250df;--red:#ff3b30;--purple:#af52de;
      --shadow:0 4px 20px rgba(0,0,0,.1),0 1px 4px rgba(0,0,0,.08);
      --shadow-sm:0 2px 8px rgba(0,0,0,.08);
      --radius:12px
    }

    /* ── Root shell ── */
    #dtr-root-outer{position:fixed;inset:0;z-index:2147483647;overflow:hidden;background:#0c0e14}#dtr-root{position:absolute;top:0;left:0;right:0;bottom:0;font-family:var(--font);font-size:clamp(13px,1.1vw,15px);background:var(--bg);color:var(--text);display:flex;flex-direction:column;overflow:hidden}

    /* ── Top bar ── */
    #dtr-topbar{height:54px;flex-shrink:0;background:var(--glass);backdrop-filter:blur(24px) saturate(180%);-webkit-backdrop-filter:blur(24px) saturate(180%);border-bottom:1px solid var(--glass-border);display:flex;align-items:center;padding:0 16px;gap:8px;position:relative;z-index:10}
    .dtr-logo{display:flex;align-items:center;gap:10px;padding-right:18px;border-right:1px solid var(--border);margin-right:8px;flex-shrink:0;cursor:default;user-select:none}
    .dtr-logo-icon{display:flex;align-items:center;justify-content:center;flex-shrink:0}
    .dtr-logo-icon svg{width:auto;height:auto;display:block}
    .dtr-logo-text{display:flex;flex-direction:column;gap:0px;line-height:1}
    .dtr-logo-team{font-size:.72rem;font-weight:800;letter-spacing:.45em;text-transform:uppercase;background:linear-gradient(90deg,#4f9eff,#a371f7,#4f9eff);background-size:200% auto;-webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text;animation:shimmerText 3s linear infinite;margin-bottom:2px}
    .dtr-logo-name{font-size:1.15rem;font-weight:900;letter-spacing:-.5px;background:linear-gradient(90deg,#4f9eff,#a371f7,#30d158,#4f9eff);background-size:300% auto;-webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text;animation:shimmerText 4s linear infinite}
    .dtr-tabs{display:flex;gap:2px;flex:1;overflow-x:auto;scrollbar-width:none}
    .dtr-tabs::-webkit-scrollbar{display:none}
    .dtr-tab{padding:6px 13px;border-radius:7px;border:none;background:transparent;color:var(--text2);font-family:var(--font);font-size:.86rem;font-weight:500;cursor:pointer;transition:all .12s;display:flex;align-items:center;gap:5px;white-space:nowrap;flex-shrink:0}
    .dtr-tab:hover{opacity:.85}
    .dtr-tab.active{background:linear-gradient(135deg,#4f9eff,#a371f7);color:#fff!important;font-weight:700;box-shadow:0 2px 12px rgba(79,158,255,.35)}
    .lt .dtr-tab.active{background:linear-gradient(135deg,#007aff,#5856d6);color:#fff!important;box-shadow:0 2px 10px rgba(0,122,255,.3)}
    .dtr-tab svg{width:13px;height:13px;flex-shrink:0}
    .dtr-topbar-r{margin-left:auto;display:flex;align-items:center;gap:7px;flex-shrink:0}
    .dtr-user-pill{display:flex;align-items:center;gap:6px;padding:5px 12px;border-radius:20px;background:var(--bg3);border:1px solid var(--border);font-size:.85rem}
    .dtr-user-pill .av{width:18px;height:18px;border-radius:50%;background:linear-gradient(135deg,var(--accent),var(--accent2));font-size:9px;font-weight:700;color:#fff;display:flex;align-items:center;justify-content:center}
    .role-badge{padding:2px 7px;border-radius:4px;font-size:.77rem;font-weight:700;text-transform:uppercase;letter-spacing:.5px;background:rgba(88,166,255,.12);color:var(--accent)}
    .icon-btn{width:30px;height:30px;border-radius:6px;border:1px solid var(--border);background:var(--bg3);color:var(--text2);cursor:pointer;display:flex;align-items:center;justify-content:center;transition:all .15s}
    .icon-btn:hover{background:var(--bg4);color:var(--text)}
    .icon-btn.danger:hover{background:rgba(248,81,73,.1);color:var(--red);border-color:rgba(248,81,73,.3)}
    .icon-btn svg{width:14px;height:14px}

    /* ── Body ── */
    #dtr-body{flex:1;overflow:hidden;display:flex}
    #dtr-sidebar{width:210px;flex-shrink:0;background:var(--glass);backdrop-filter:blur(20px) saturate(160%);-webkit-backdrop-filter:blur(20px) saturate(160%);border-right:1px solid var(--glass-border);display:flex;flex-direction:column;padding:10px 8px;overflow-y:auto;position:relative;z-index:2}
    .sb-sec{font-size:.68rem;font-weight:800;text-transform:uppercase;letter-spacing:1.3px;color:var(--text3);padding:12px 12px 5px;margin-top:4px}
    .sb-item{display:flex;align-items:center;gap:9px;padding:9px 12px;border-radius:8px;border:none;background:transparent;color:var(--text2);font-family:var(--font);font-size:.875rem;font-weight:500;cursor:pointer;transition:all .12s;width:100%;text-align:left;border-left:3px solid transparent}
    .sb-item:hover{background:var(--bg3);color:var(--text);border-left-color:var(--border2)}
    .sb-item.active{background:linear-gradient(135deg,#4f9eff,#a371f7);color:#fff!important;font-weight:700;border-left-color:transparent}
    .lt .sb-item.active{background:linear-gradient(135deg,#007aff,#5856d6);color:#fff!important}
    .sb-item svg{width:14px;height:14px;flex-shrink:0}
    .sb-footer{margin-top:auto;padding:10px 8px;border-top:1px solid var(--border)}
    .sb-stat{margin-bottom:10px}
    .sb-stat .lbl{font-size:.75rem;color:var(--text3);margin-bottom:2px}
    .sb-stat .val{font-size:1.25rem;font-weight:700;color:var(--accent)}
    #dtr-main{flex:1;overflow-y:auto;padding:20px 24px;background:transparent;position:relative;z-index:1}
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
    .dtr-input,.dtr-textarea{padding:9px 12px;border-radius:var(--radius);background:var(--bg3);border:1px solid var(--border);color:var(--text);font-family:var(--font);font-size:.96rem;transition:all .15s;width:100%;appearance:none}
    .dtr-select{padding:9px 12px;border-radius:var(--radius);background:var(--bg3);border:1px solid var(--border);color:var(--text);font-family:var(--font);font-size:.96rem;transition:border-color .12s;width:100%;appearance:none;color-scheme:dark}
    #dtr-root.lt .dtr-select{background:#ffffff;color:#1c1c1e;border-color:rgba(0,0,0,.12);color-scheme:light}
    .dtr-input:focus,.dtr-select:focus,.dtr-textarea:focus{outline:none;border-color:var(--accent);box-shadow:0 0 0 2px rgba(79,158,255,.15)}
    .dtr-input,.dtr-select,.dtr-textarea{color:var(--text)!important;-webkit-text-fill-color:var(--text)!important}
    .dtr-input[readonly]{opacity:1!important;color:var(--text)!important;-webkit-text-fill-color:var(--text)!important;background:var(--bg3)!important;cursor:default}
    .dtr-input::placeholder,.dtr-textarea::placeholder{color:var(--text3)!important;-webkit-text-fill-color:var(--text3)!important}
    input[type=date].dtr-input::-webkit-calendar-picker-indicator{filter:invert(.7) opacity(.7)}
    .dtr-input:disabled,.dtr-select:disabled{opacity:.5!important;cursor:not-allowed}
    .dtr-select option{background:#141826;color:#e8ecf4;font-family:-apple-system,sans-serif}
    .dtr-select option:hover,.dtr-select option:focus,.dtr-select option:checked{background:#1e2a45!important;color:#4f9eff!important}
    #dtr-root.lt .dtr-select option{background:#ffffff;color:#1c1c1e}
    #dtr-root.lt .dtr-select{color-scheme:light}
    .dtr-textarea{resize:vertical}

    /* ── Buttons ── */
    .btn{display:inline-flex;align-items:center;justify-content:center;gap:5px;padding:8px 16px;border-radius:var(--radius);font-family:var(--font);font-size:.94rem;font-weight:600;cursor:pointer;border:none;transition:all .15s;white-space:nowrap}
    .btn svg{width:13px;height:13px}
    .btn-primary{background:var(--accent);color:#fff;box-shadow:0 2px 12px rgba(79,158,255,.35)}.btn-primary:hover{opacity:.9;box-shadow:0 4px 16px rgba(79,158,255,.4)}
    .btn-ghost{background:var(--bg3);border:1px solid var(--border);color:var(--text2)}.btn-ghost:hover{background:var(--bg4);color:var(--text)}
    .btn-danger{background:rgba(248,81,73,.1);border:1px solid rgba(248,81,73,.25);color:var(--red)}.btn-danger:hover{background:rgba(248,81,73,.18)}
    .btn-full{width:100%}.btn-sm{padding:5px 11px;font-size:.875rem}.btn-xs{padding:3px 8px;font-size:.815rem}

    /* ── Page header ── */
    .ph{display:flex;align-items:flex-start;justify-content:space-between;margin-bottom:18px;gap:12px}
    .ph-left .ph-title{font-size:1.25rem;font-weight:800;letter-spacing:-.4px;background:linear-gradient(90deg,var(--text),var(--accent),var(--text));background-size:300% auto;-webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text;animation:shimmerText 5s linear infinite}
    .ph-left .ph-sub{font-size:.9rem;color:var(--text3);margin-top:3px}
    .ph-actions{display:flex;gap:6px;align-items:center;flex-shrink:0}

    /* ── Cards ── */
    .card{background:var(--glass);backdrop-filter:blur(16px);-webkit-backdrop-filter:blur(16px);border:1px solid var(--glass-border);border-radius:var(--radius);padding:18px;margin-bottom:14px;box-shadow:var(--shadow-sm)}
    .card-title{font-size:.78rem;font-weight:700;text-transform:uppercase;letter-spacing:.7px;color:var(--text3);margin-bottom:12px;display:flex;align-items:center;gap:6px}

    /* ── Stat grid ── */
    .stats-grid{display:grid;gap:12px;margin-bottom:16px}
    .sg4{grid-template-columns:repeat(4,1fr)}.sg3{grid-template-columns:repeat(3,1fr)}.sg2{grid-template-columns:1fr 1fr}
    .stat-card{background:var(--glass);backdrop-filter:blur(16px);-webkit-backdrop-filter:blur(16px);border:1px solid var(--glass-border);border-radius:var(--radius);padding:16px;box-shadow:var(--shadow-sm)}
    .stat-card .lbl{font-size:.8rem;color:var(--text3);text-transform:uppercase;letter-spacing:.5px;margin-bottom:6px}
    .stat-card .val{font-size:2rem;font-weight:800;letter-spacing:-1px;line-height:1;color:var(--text)}
    .stat-card .sub{font-size:.8rem;color:var(--text3);margin-top:4px}
    .ab{border-top:2px solid var(--accent)}.gb{border-top:2px solid var(--green)}.amb{border-top:2px solid var(--amber)}.pb{border-top:2px solid var(--purple)}.rb2{border-top:2px solid var(--red)}

    /* ── Grid helpers ── */
    .g2{display:grid;grid-template-columns:1fr 1fr;gap:12px}
    .g3{display:grid;grid-template-columns:repeat(3,1fr);gap:12px}
    .ga{display:grid;grid-template-columns:2fr 1fr 1fr 1fr;gap:10px}

    /* ── Tables ── */
    .tbl-wrap{border:1px solid var(--glass-border);border-radius:var(--radius);overflow:hidden;backdrop-filter:blur(8px);-webkit-backdrop-filter:blur(8px)}
    .dtr-table{width:100%;border-collapse:collapse;font-size:.91rem}
    .dtr-table th{background:var(--glass);padding:12px 14px;text-align:left;color:var(--text);font-weight:800;font-size:.8rem;text-transform:uppercase;letter-spacing:.7px;border-bottom:1px solid var(--glass-border)}
    .dtr-table td{padding:11px 14px;font-size:.92rem;font-weight:500;color:var(--text);border-bottom:1px solid var(--border)}
    .dtr-table tr:last-child td{border-bottom:none}
    .dtr-table tr:hover td{background:var(--glass)}
    .dtr-table .bold{font-weight:800;color:var(--text)}
    .dtr-table .mono{font-family:var(--mono);font-weight:400;font-size:.92rem}

    /* ── Badges ── */
    .badge{display:inline-flex;align-items:center;gap:3px;padding:2px 7px;border-radius:4px;font-size:.77rem;font-weight:700}
    .bg2{background:rgba(63,185,80,.12);color:var(--green)}
    .ba{background:rgba(130,80,223,.18);color:var(--amber);font-weight:800}
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

    /* ── Animations ─────────────────────────────────────────────── */
    @keyframes fadeSlideIn{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:translateY(0)}}
    @keyframes fadeSlideLeft{from{opacity:0;transform:translateX(-12px)}to{opacity:1;transform:translateX(0)}}
    @keyframes scaleIn{from{opacity:0;transform:scale(.97)}to{opacity:1;transform:scale(1)}}
    @keyframes pulse{0%,100%{transform:scale(1)}50%{transform:scale(1.04)}}
    @keyframes shimmer{0%{background-position:-400px 0}100%{background-position:400px 0}}
    @keyframes floatY{0%,100%{transform:translateY(0)}50%{transform:translateY(-4px)}}
    @keyframes logoFlow{0%{background-position:0% 50%}50%{background-position:100% 50%}100%{background-position:0% 50%}}
    @keyframes logoFade{0%,100%{opacity:.7}50%{opacity:1}}
    @keyframes meshFloat{0%{transform:translate(0,0) scale(1)}33%{transform:translate(40px,-30px) scale(1.08)}66%{transform:translate(-30px,40px) scale(.95)}100%{transform:translate(0,0) scale(1)}}
    @keyframes meshFloat2{0%{transform:translate(0,0) scale(1)}33%{transform:translate(-35px,25px) scale(1.06)}66%{transform:translate(25px,-20px) scale(.96)}100%{transform:translate(0,0) scale(1)}}
    @keyframes pulseRing{0%{opacity:.6}100%{opacity:0}}
    @keyframes shimmerText{0%{background-position:-200% center}100%{background-position:200% center}}
    @keyframes devicePopIn{from{opacity:0;transform:scale(.92)}to{opacity:1;transform:scale(1)}}
    .device-card{display:flex;flex-direction:column;align-items:center;justify-content:center;gap:12px;padding:28px 20px;border-radius:16px;border:2px solid var(--border);background:var(--bg3);cursor:pointer;transition:all .18s;flex:1}
    .device-card:hover{border-color:var(--accent);background:rgba(79,158,255,.06);transform:translateY(-2px)}
    .device-card.selected{border-color:var(--accent);background:rgba(79,158,255,.1)}
    .device-card svg{opacity:.8}
    .device-card .dc-label{font-size:.95rem;font-weight:700;color:var(--text)}
    .device-card .dc-sub{font-size:.75rem;color:var(--text3);text-align:center;line-height:1.4}
    @keyframes drift1{0%{transform:translate(0,0)}33%{transform:translate(5%,-8%)}66%{transform:translate(-6%,-5%)}100%{transform:translate(0,0)}}
    @keyframes drift2{0%{transform:translate(0,0)}33%{transform:translate(-7%,9%)}66%{transform:translate(4%,6%)}100%{transform:translate(0,0)}}
    #wp-bg-orb{position:absolute;inset:0;pointer-events:none;z-index:0;overflow:hidden}
    #wp-bg-orb .orb{position:absolute;border-radius:50%;filter:blur(60px);will-change:transform}
    #wp-bg-orb .o1{width:70%;height:70%;top:-15%;left:-20%;background:radial-gradient(circle,rgba(79,158,255,.07) 0%,transparent 70%);opacity:.8}
    #wp-bg-orb .o2{width:65%;height:65%;bottom:-15%;right:-15%;background:radial-gradient(circle,rgba(163,113,247,.06) 0%,transparent 70%);opacity:.7}
    #wp-bg-orb .o3{display:none}
    @keyframes screenBlink{0%,94%,100%{opacity:1}95%,97%{opacity:0}}
    @keyframes screenGlow{0%,100%{filter:drop-shadow(0 0 4px rgba(79,158,255,.5))}50%{filter:drop-shadow(0 0 10px rgba(79,158,255,.9))}}
    @keyframes typingCursor{0%,100%{opacity:1}50%{opacity:0}}
    @keyframes barGrow{0%{width:0}100%{width:100%}}
    @keyframes rotateBlade{from{transform:rotate(0deg)}to{transform:rotate(360deg)}}
    @keyframes glowPulse{0%,100%{box-shadow:0 0 0 0 rgba(79,158,255,0)}50%{box-shadow:0 0 0 6px rgba(79,158,255,.15)}}
    @keyframes typeIn{from{width:0;opacity:0}to{width:100%;opacity:1}}


    /* View transition */
    .dtr-view{animation:none}
    .dtr-view.anim{animation:fadeSlideIn .28s cubic-bezier(.22,1,.36,1) both}

    /* Sidebar items hover */
    .sb-item{transition:all .18s cubic-bezier(.22,1,.36,1)}
    .sb-item:hover{transform:translateX(3px)}
    .sb-item.active{animation:glowPulse 2.5s ease-in-out infinite}

    /* Tabs hover */
    .dtr-tab{transition:all .15s cubic-bezier(.22,1,.36,1)}
    .dtr-tab:hover{transform:translateY(-1px)}
    .dtr-tab.active{animation:none}

    /* Buttons */
    .btn{transition:all .18s cubic-bezier(.34,1.56,.64,1)}
    .btn:hover{transform:translateY(-2px);box-shadow:0 6px 20px rgba(0,0,0,.2)}
    .btn:active{transform:translateY(0) scale(.97)}
    .btn-primary:hover{opacity:.92}

    /* Icon buttons */
    .icon-btn{transition:all .18s cubic-bezier(.34,1.56,.64,1)}
    .icon-btn:hover{transform:scale(1.12);background:var(--bg4)}
    .icon-btn:active{transform:scale(.93)}

    /* Cards */
    .card{transition:all .22s cubic-bezier(.22,1,.36,1)}
    .card:hover{transform:translateY(-2px);box-shadow:0 12px 36px rgba(0,0,0,.25)}
    .stat-card{transition:all .22s cubic-bezier(.22,1,.36,1)}
    .stat-card:hover{transform:translateY(-2px) scale(1.01);box-shadow:0 10px 30px rgba(0,0,0,.2)}

    /* Date pills */
    .date-pill{transition:all .2s cubic-bezier(.34,1.56,.64,1)}
    .date-pill:hover{transform:scale(1.02)}


    /* Task cards */
    .task-card{transition:all .2s cubic-bezier(.22,1,.36,1)}
    .task-card:hover{transform:translateY(-1px);box-shadow:0 8px 28px rgba(0,0,0,.2)}

    /* Status tiles */
    .status-tile{transition:all .2s cubic-bezier(.34,1.56,.64,1)}
    .status-tile:hover{transform:scale(1.03) translateY(-2px)}


    /* Attendance day cards */
    .att-day-card{transition:all .18s cubic-bezier(.22,1,.36,1)}
    .att-day-card:hover{transform:translateY(-2px) scale(1.02)}

    /* Table rows */
    .dtr-table tr{transition:background .12s}
    .dtr-table tbody tr:hover td{background:var(--glass)!important;transform:none}

    /* Submit button pulse when form ready */
    [data-action="do-submit"]:not(:disabled){}
    [data-action="do-submit"]:hover{opacity:.92}

    /* Chip pills in admin */
    [data-action="adm-assoc-toggle"]{transition:opacity .1s}
    [data-action="adm-assoc-toggle"]:hover{opacity:.85}

    /* Laptop animation in topbar */
    #wp-laptop-anim{display:inline-flex;align-items:center;gap:4px;animation:floatY 3s ease-in-out infinite}
    #wp-laptop-anim .blade{animation:rotateBlade 2s linear infinite;transform-origin:center}
    /* Hover color flow on interactive boxes */


    /* ── Bar chart ── */
    .bar-rows{display:flex;flex-direction:column;gap:8px}
    .bar-row{display:flex;align-items:center;gap:10px}
    .bar-label{font-size:.855rem;color:var(--text2);width:120px;flex-shrink:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
    .bar-track{flex:1;height:22px;background:rgba(0,0,0,.1);border-radius:6px;overflow:hidden;position:relative;border:1px solid rgba(0,0,0,.06)}
    .bar-fill{height:100%;border-radius:5px;display:flex;align-items:center;padding-left:7px;transition:width .7s cubic-bezier(.34,1.56,.64,1)}
    .bar-fill-text{font-size:.79rem;font-weight:700;color:#fff}
    .bar-aside{font-size:.79rem;color:var(--text3);font-family:var(--mono);width:38px;text-align:right;flex-shrink:0}
    .chart-title{font-size:.78rem;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--text3);margin-bottom:12px}

    /* ── Task cards (Submit tab) ── */
    .task-card{background:var(--bg3);border:1px solid var(--border);border-radius:10px;padding:14px;margin-bottom:10px}
    .task-card.leave{border-color:rgba(255,214,10,.25);background:rgba(255,214,10,.06)}
    .tc-header{display:flex;align-items:center;justify-content:space-between;margin-bottom:12px}
    .tc-num{font-size:.77rem;font-weight:700;text-transform:uppercase;letter-spacing:.8px;color:var(--accent)}
    .npt-box{padding:8px 11px;border-radius:var(--radius);background:rgba(255,214,10,.1);border:1px solid rgba(255,214,10,.2);color:var(--amber);font-size:1rem;font-weight:700;font-family:var(--mono);text-align:center}
    .npt-box.prod{background:rgba(48,209,88,.1);border-color:rgba(48,209,88,.2);color:var(--green)}
    .leave-tag{display:none;align-items:center;gap:8px;padding:9px 12px;border-radius:8px;margin-top:10px;background:rgba(210,153,34,.08);border:1px solid rgba(210,153,34,.2);color:var(--amber);font-size:.875rem;flex-wrap:wrap}
    .leave-tag.show{display:flex}
    .leave-type-btn{padding:4px 11px;border-radius:6px;border:1px solid rgba(210,153,34,.35);background:transparent;color:var(--amber);font-family:var(--font);font-size:.78rem;font-weight:600;cursor:pointer;transition:all .15s}
    .leave-type-btn.active{background:rgba(210,153,34,.2);border-color:var(--amber)}
    .comments-box{display:none;margin-top:10px}
    .opt-note-toggle{display:flex;align-items:center;gap:5px;margin-top:8px;color:var(--accent);font-size:.82rem;font-weight:600;cursor:pointer;padding:4px 0;user-select:none}
    .opt-note-toggle:hover{color:var(--accent2)}
    .req-star{color:var(--red)}
    .opt-note-toggle{display:flex;align-items:center;gap:5px;margin-top:8px;color:var(--accent);font-size:.82rem;font-weight:600;cursor:pointer;padding:4px 0;user-select:none}
    .opt-note-toggle:hover{color:var(--accent2)}
    .opt-note-box{margin-top:2px}
    .req-star{color:var(--red)}
    .add-task-btn{width:100%;padding:9px;border-radius:8px;border:1px dashed var(--border2);background:transparent;color:var(--accent);cursor:pointer;font-family:var(--font);font-size:.92rem;font-weight:600;transition:all .15s;margin-bottom:12px}
    .add-task-btn:hover{background:rgba(88,166,255,.05);border-color:var(--accent)}
    .date-pill-row{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-top:8px}
    .date-pill{padding:14px 16px;border-radius:12px;border:2px solid var(--border);background:var(--bg3);color:var(--text2);font-family:var(--font);font-size:.88rem;font-weight:600;cursor:pointer;transition:all .2s;text-align:left;display:flex;flex-direction:column;gap:3px}
    .date-pill .dp-label{font-size:.7rem;text-transform:uppercase;letter-spacing:.8px;color:var(--text3);font-weight:700}
    .date-pill .dp-date{font-size:1.05rem;font-weight:800;color:var(--text)}
    .date-pill:hover{border-color:var(--accent);background:rgba(88,166,255,.05)}
    .date-pill.active{border-color:var(--accent);background:rgba(88,166,255,.08)}
    .date-pill.active .dp-date{color:var(--accent)}
    .date-pill.active .dp-label{color:var(--accent)}
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
    .wv-range{flex:1;min-width:0}.wv-range-title{font-size:1.05rem;font-weight:800;color:var(--text)}.wv-range-sub{font-size:.82rem;color:var(--text3)}
    .wv-today-btn{padding:5px 12px;border-radius:7px;border:1px solid var(--border);background:var(--bg3);font-size:.8rem;font-weight:600;color:var(--text2);cursor:pointer;font-family:var(--font);transition:.2s}
    .wv-today-btn:hover{background:rgba(88,166,255,.1);color:var(--accent);border-color:var(--accent)}
    .wv-grid-wrap{border:1px solid var(--border);border-radius:12px;overflow:hidden;background:var(--bg2)}
    .wv-header{display:grid;grid-template-columns:160px repeat(7,1fr);background:var(--bg3);border-bottom:2px solid var(--border)}
    .wv-hcell{padding:12px 8px;text-align:center;font-size:.82rem;font-weight:800;color:var(--text3);text-transform:uppercase;letter-spacing:.5px}
    .wv-hcell.today-col{color:var(--accent);border-bottom:2px solid var(--accent);margin-bottom:-2px}
    .wv-hcell .wv-hdate{font-size:.78rem;color:var(--text2);margin-top:2px;text-transform:none;letter-spacing:0;font-weight:600}
    .wv-hcell.today-col .wv-hdate{color:var(--accent)}
    .wv-row{display:grid;grid-template-columns:160px repeat(7,1fr);border-bottom:1px solid var(--border);transition:.15s}
    .wv-row:last-child{border-bottom:none}
    .wv-row:hover{background:rgba(88,166,255,.025)}
    .wv-name-cell{padding:12px;display:flex;align-items:center;gap:9px;border-right:1px solid var(--border)}
    .wv-av{width:28px;height:28px;border-radius:7px;background:linear-gradient(135deg,var(--accent),var(--accent2));display:flex;align-items:center;justify-content:center;font-size:10px;font-weight:700;color:#fff;flex-shrink:0}
    .wv-name{font-size:.83rem;font-weight:600;color:var(--text);overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
    .wv-cell{padding:10px 5px;text-align:center;border-right:1px solid rgba(255,255,255,.03);display:flex;flex-direction:column;align-items:center;justify-content:center;gap:4px;cursor:pointer}
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
    .att-day-card{background:var(--bg2);border:2px solid var(--border);border-radius:12px;padding:16px 8px;text-align:center;cursor:pointer;transition:all .18s;position:relative;user-select:none}
    .att-day-card:hover{border-color:var(--accent);transform:translateY(-2px);box-shadow:0 6px 18px rgba(0,0,0,.2)}
    .att-day-card.att-day-selected{border-color:var(--accent)!important;box-shadow:0 0 0 3px rgba(88,166,255,.2)}
    .att-day-card.att-day-today .att-day-name{color:var(--accent);font-weight:800}
    .att-day-card.att-day-weekend{opacity:.5}
    .att-day-card.att-day-weekend:hover{opacity:.75}
    .att-day-card.att-multi{outline:3px solid var(--accent2);outline-offset:2px;background:rgba(163,113,247,.08)!important}
    .att-bulk-bar{background:var(--bg2);border:1px solid var(--accent2);border-radius:12px;padding:14px 18px;margin-bottom:14px;display:flex;align-items:center;gap:10px;flex-wrap:wrap;animation:fi .2s ease}
    .att-bulk-btn{padding:7px 14px;border-radius:8px;border:2px solid var(--border);background:var(--bg3);color:var(--text2);font-family:var(--font);font-size:.85rem;font-weight:700;cursor:pointer;transition:all .15s}
    .att-bulk-btn:hover{transform:translateY(-1px);box-shadow:0 4px 12px rgba(0,0,0,.2)}
    .att-select-toggle{padding:6px 14px;border-radius:8px;border:1px solid var(--border);background:var(--bg3);color:var(--text2);font-family:var(--font);font-size:.82rem;font-weight:600;cursor:pointer;transition:all .15s}
    .att-select-toggle.on{background:rgba(163,113,247,.12);border-color:var(--accent2);color:var(--accent2)}
    .att-day-card.att-day-multi{outline:3px solid var(--accent2);outline-offset:2px;background:rgba(163,113,247,.08)!important}
    .att-bulk-bar{background:var(--bg2);border:1px solid var(--accent2);border-radius:12px;padding:14px 18px;margin-bottom:14px;display:flex;align-items:center;gap:10px;flex-wrap:wrap;animation:fi .2s ease}
    .att-bulk-status{display:flex;gap:6px;flex-wrap:wrap;flex:1}
    .att-bulk-btn{padding:7px 16px;border-radius:8px;border:2px solid var(--border);background:var(--bg3);color:var(--text2);font-family:var(--font);font-size:.85rem;font-weight:700;cursor:pointer;transition:all .15s}
    .att-bulk-btn:hover{transform:translateY(-1px);box-shadow:0 4px 12px rgba(0,0,0,.2)}
    .att-bulk-btn.sel{border-color:var(--accent2);background:rgba(163,113,247,.12);color:var(--accent2)}
    .att-select-toggle{padding:6px 14px;border-radius:8px;border:1px solid var(--border);background:var(--bg3);color:var(--text2);font-family:var(--font);font-size:.82rem;font-weight:600;cursor:pointer;transition:all .15s}
    .att-select-toggle.active{background:rgba(163,113,247,.12);border-color:var(--accent2);color:var(--accent2)}
    .att-day-name{font-size:.76rem;font-weight:800;text-transform:uppercase;letter-spacing:.7px;color:var(--text3);margin-bottom:2px}
    .att-day-date{font-size:.85rem;font-weight:700;color:var(--text);margin-bottom:6px}
    .att-day-icon{font-size:22px;margin-bottom:6px;line-height:1}
    .att-day-dot{width:12px;height:12px;border-radius:50%;margin:6px auto 4px;transition:.2s}
    .att-day-status{font-size:.78rem;font-weight:800;text-transform:uppercase;letter-spacing:.5px}
    .att-day-active-dot{position:absolute;bottom:6px;left:50%;transform:translateX(-50%);width:6px;height:6px;border-radius:50%;background:var(--accent)}
    .att-edit-card{background:var(--bg2);border:1px solid var(--border2);border-radius:14px;padding:22px;margin-top:8px;animation:fi .2s ease}
    .att-edit-header{display:flex;align-items:center;justify-content:space-between;margin-bottom:16px;flex-wrap:wrap;gap:8px}
    .att-edit-date{font-size:1rem;font-weight:700;color:var(--text)}
    .att-edit-datesub{font-size:.78rem;color:var(--text3);margin-top:2px}
    .att-saved-pill{padding:3px 10px;border-radius:20px;font-size:.78rem;font-weight:700}

    /* ── My Calendar ── */
    .cal-grid-wrap{background:var(--bg2);border:1px solid var(--border);border-radius:12px;overflow:hidden}
    .cal-header{display:grid;grid-template-columns:repeat(7,1fr);background:var(--bg3);border-bottom:1px solid var(--border)}
    .cal-hcell{padding:10px 4px;text-align:center;font-size:.8rem;font-weight:800;color:var(--text3);text-transform:uppercase;letter-spacing:.5px}
    .cal-grid{display:grid;grid-template-columns:repeat(7,1fr)}
    .cal-cell{min-height:78px;border:1px solid var(--border);padding:7px;display:flex;flex-direction:column;align-items:center;cursor:pointer;transition:all .15s;position:relative}
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
    trend:`<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="22 7 13.5 15.5 8.5 10.5 2 17"/><polyline points="16 7 22 7 22 13"/></svg>`,
  };

  // Associate tabs
  const ASSOC_TABS = [
    { id:'submit',    label:'Submit Tasks',     icon:ic.submit,   group:'Reporter'    },
    { id:'analytics', label:'My Analytics',     icon:ic.chart,    group:'Reporter'    },
    { id:'tracker',   label:'My Tracker',       icon:ic.tracker,  group:'Reporter'    },
    { id:'weekly',    label:'Weekly Calendar',  icon:ic.mark,     group:'Attendance'  },
    { id:'calendar',  label:'My Calendar',      icon:ic.att,      group:'Attendance'  },
    { id:'missednpt', label:'Missed NPT',       icon:ic.npt,      group:'Attendance'  },
    { id:'settings',  label:'Settings',         icon:ic.settings, group:'Settings'    },
  ];
  // Admin tabs
  const ADMIN_TABS = [
    { id:'overview',  label:'Overview',      icon:ic.chart,    group:'Dashboard'   },
    { id:'teamtrack', label:'Team Tracker',  icon:ic.tracker,  group:'Dashboard'   },
    { id:'admanalytics', label:'Analytics', icon:ic.trend, group:'Dashboard' },
    { id:'attweek',   label:'Week View',     icon:ic.att,      group:'Attendance'  },
    { id:'attnpt',    label:'NPT Log',       icon:ic.npt,      group:'Attendance'  },
    { id:'settings',  label:'Settings',      icon:ic.settings, group:'Settings'    },
  ];
  const TABS = ASSOC_TABS; // default — switched by buildApp per role

  function boot() {
    const ex = document.getElementById('dtr-root'); if (ex) ex.remove();
    const outer = document.createElement('div'); outer.id = 'dtr-root-outer';
    root = document.createElement('div'); root.id = 'dtr-root';
    if (theme === 'light') { root.classList.add('lt'); outer.classList.add('lt'); }
    outer.appendChild(root);
    document.body.appendChild(outer);
    root.addEventListener('click', handleClick);
    root.addEventListener('input', handleInput);
    root.addEventListener('scroll', ()=>{}, {passive:true});
    root.addEventListener('change', handleChange);
    root.addEventListener('keydown', handleKeydown);
    setTimeout(flushQueue, 3000); // flush queued items on every page
    teamStatusCache = safeLoadObj('dtr_teamcache', {});
    if (authState.loggedIn) {
      buildApp();
      // Silently sync own data from SP on startup
      setTimeout(() => fetchOwnAttendance(true), 1500);
      setTimeout(() => fetchOwnTasks(true),       3000);
      setTimeout(() => fetchOwnNPT(true),         4500);
    } else {
      renderLogin();
    }
  }

  const q  = sel => root.querySelector(sel);
  const qa = sel => root.querySelectorAll(sel);

  // ── SSO detection ────────────────────────────────────────────────────
  function detectSSOUser() {
    return new Promise((resolve, reject) => {
      GM_xmlhttpRequest({
        method: 'GET',
        url: SP.SITE + '/_api/web/currentuser',
        headers: { 'Accept': 'application/json;odata=verbose' },
        withCredentials: true,
        onload: res => {
          try {
            const d = JSON.parse(res.responseText).d;
            resolve({ displayName: d.Title||'', email: d.Email||'', loginName: d.LoginName||'' });
          } catch { reject(new Error('parse failed')); }
        },
        onerror: () => reject(new Error('network error'))
      });
    });
  }

  function resolveRole(displayName, email) {
    const dn  = (displayName || '').toLowerCase().trim();
    const pfx = (email || '').split('@')[0].toLowerCase().trim();
    // Check admins first — strict exact match
    const adminHit = ADMINS.find(a => {
      const al = a.toLowerCase();
      return al === dn || al === pfx || dn.startsWith(al) || pfx.startsWith(al);
    });
    if (adminHit) return { name: adminHit, role: 'admin' };
    // Check associates
    const assocHit = ASSOCIATES.find(a => {
      const al = a.toLowerCase();
      return al === dn || al === pfx || dn.startsWith(al) || pfx.startsWith(al);
    });
    if (assocHit) {
      if (!ADMINS.find(a => a.toLowerCase() === assocHit.toLowerCase()))
        return { name: assocHit, role: 'associate' };
    }
    return null;
  }

  function renderLogin() {
    root.innerHTML =
      '<div id="view-login" style="position:absolute;inset:0;display:flex;align-items:center;justify-content:center;background:linear-gradient(135deg,#080c14 0%,#12103a 50%,#080c14 100%);overflow:hidden">' +
      '<div style="position:absolute;width:500px;height:500px;border-radius:50%;background:radial-gradient(circle,rgba(99,102,241,.12),transparent 70%);top:-120px;right:-100px;pointer-events:none"></div>' +
      '<div style="position:absolute;width:400px;height:400px;border-radius:50%;background:radial-gradient(circle,rgba(163,113,247,.08),transparent 70%);bottom:-100px;left:-80px;pointer-events:none"></div>' +
      '<div style="position:relative;z-index:1;width:420px;max-width:92vw;background:rgba(255,255,255,.035);backdrop-filter:blur(32px);border:1px solid rgba(255,255,255,.07);border-radius:20px;padding:40px 36px 36px;box-shadow:0 32px 64px rgba(0,0,0,.5)">' +
        // Logo + title
        '<div style="display:flex;flex-direction:column;align-items:center;gap:10px;margin-bottom:28px">' +
          '<div style="width:54px;height:54px;border-radius:16px;background:linear-gradient(135deg,#58a6ff,#a371f7);display:flex;align-items:center;justify-content:center;box-shadow:0 8px 24px rgba(99,102,241,.35)">' + ic.logo + '</div>' +
          '<div style="text-align:center">' +
            '<div style="font-size:24px;font-weight:800;color:#fff;letter-spacing:-.5px">WorkPulse</div>' +
            '<div style="font-size:12px;color:rgba(255,255,255,.3);margin-top:3px">Associate Portal — Secure Access</div>' +
          '</div>' +
        '</div>' +
        // SSO spinner
        '<div id="sso-detecting" style="text-align:center;padding:18px 0">' +
          '<div style="display:inline-flex;align-items:center;gap:10px;color:rgba(255,255,255,.4);font-size:13px">' +
            '<svg id="sso-spin" style="width:18px;height:18px;flex-shrink:0" viewBox="0 0 24 24" fill="none">' +
              '<circle cx="12" cy="12" r="9" stroke="rgba(255,255,255,.1)" stroke-width="2"/>' +
              '<path d="M12 3a9 9 0 0 1 9 9" stroke="#a371f7" stroke-width="2" stroke-linecap="round"/>' +
            '</svg>' +
            'Detecting your identity via SharePoint SSO...' +
          '</div>' +
        '</div>' +
        '<div id="sso-result" style="display:none"></div>' +
        // Manual fallback — plain text input, no dropdown
        '<div id="sso-manual" style="display:none">' +
          '<div style="display:flex;align-items:center;gap:10px;margin:4px 0 18px;color:rgba(255,255,255,.18);font-size:10px;text-transform:uppercase;letter-spacing:1.5px;font-weight:600">' +
            '<div style="flex:1;height:1px;background:rgba(255,255,255,.06)"></div>Manual Login<div style="flex:1;height:1px;background:rgba(255,255,255,.06)"></div>' +
          '</div>' +
          '<input id="login-name" type="text" autocomplete="off" spellcheck="false" placeholder="Enter your login name (e.g. Bhurak)" style="width:100%;padding:12px 14px;border-radius:10px;background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.1);color:#fff;font-size:.9rem;font-family:inherit;outline:none;margin-bottom:10px;-webkit-text-fill-color:#fff">' +
          '<div id="login-err" style="color:#f85149;font-size:.82rem;margin-bottom:10px;min-height:16px;text-align:center"></div>' +
          '<button data-action="do-login" style="width:100%;padding:13px;border:none;border-radius:10px;background:linear-gradient(135deg,#58a6ff,#a371f7);color:#fff;font-size:.9rem;font-weight:700;cursor:pointer;font-family:inherit">→ Continue</button>' +
          '<div style="margin-top:10px;font-size:.72rem;color:rgba(255,255,255,.2);text-align:center">Only authorised logins accepted</div>' +
        '</div>' +
      '</div>' +
      '</div>' +
      '<div id="dtr-toast"></div>';

    // Spinner animation
    const spinEl = root.querySelector('#sso-spin');
    if (spinEl) {
      let angle = 0;
      const si = setInterval(() => {
        if (!root.querySelector('#sso-spin')) { clearInterval(si); return; }
        angle = (angle + 8) % 360;
        spinEl.style.transform = 'rotate(' + angle + 'deg)';
      }, 30);
    }
    startSSODetection();
  }

  function startSSODetection() {
    const detectEl = root.querySelector('#sso-detecting');
    const resultEl = root.querySelector('#sso-result');
    const manualEl = root.querySelector('#sso-manual');

    detectSSOUser().then(ssoUser => {
      const resolved = resolveRole(ssoUser.displayName, ssoUser.email);
      if (detectEl) detectEl.style.display = 'none';

      if (resolved) {
        // Check saved session — auto-login if same person
        const saved = loadSession();
        if (saved && saved.name === resolved.name) {
          doLoginWithRole(resolved.name, resolved.role);
          return;
        }
        if (resolved.name === SUPER_ADMIN && saved && saved.name === 'Admin-tarun') {
          doLoginWithRole('Admin-tarun', 'associate');
          return;
        }

        // Show identity card
        const isAdmin = resolved.role === 'admin';
        if (resultEl) {
          resultEl.style.display = 'block';
          resultEl.innerHTML =
            '<div style="background:rgba(99,102,241,.1);border:1px solid rgba(99,102,241,.2);border-radius:12px;padding:18px;text-align:center;margin-bottom:14px">' +
              '<div style="font-size:20px;font-weight:800;color:#a5b4fc;margin-bottom:4px">' + resolved.name + '</div>' +
              '<div style="font-size:12px;color:rgba(255,255,255,.3);margin-bottom:8px">' + (ssoUser.email || ssoUser.loginName || 'SharePoint SSO') + '</div>' +
              '<span style="padding:3px 12px;border-radius:20px;font-size:11px;font-weight:700;' +
              (isAdmin ? 'background:rgba(163,113,247,.2);color:#a371f7' : 'background:rgba(88,166,255,.2);color:#58a6ff') + '">' +
              (isAdmin ? '🛡 Admin' : '👤 Associate') + '</span>' +
            '</div>' +
            // For associate (or admin entering associate portal)
            '<button id="sso-go-btn" style="width:100%;padding:13px;border:none;border-radius:10px;background:linear-gradient(135deg,#58a6ff,#a371f7);color:#fff;font-size:.95rem;font-weight:700;cursor:pointer;font-family:inherit;margin-bottom:8px">' +
            (isAdmin ? '🛡 Enter Admin Portal' : '→ Enter WorkPulse') +
            '</button>' +
            // Super admin gets associate view option
            (resolved.name === SUPER_ADMIN
              ? '<button id="sso-assoc-btn" style="width:100%;padding:10px;border:1px solid rgba(88,166,255,.3);border-radius:10px;background:rgba(88,166,255,.07);color:#58a6ff;font-size:.85rem;font-weight:600;cursor:pointer;font-family:inherit;margin-bottom:8px">👤 Enter as Admin-tarun (Associate View)</button>'
              : '') +
            '<button id="sso-wrong-btn" style="width:100%;padding:8px;border:none;background:transparent;color:rgba(255,255,255,.25);font-size:.78rem;cursor:pointer;font-family:inherit">Not you? Enter manually</button>';

          resultEl.querySelector('#sso-go-btn').onclick    = () => doLoginWithRole(resolved.name, resolved.role);
          const assocBtn = resultEl.querySelector('#sso-assoc-btn');
          if (assocBtn) assocBtn.onclick = () => doLoginWithRole('Admin-tarun', 'associate');
          resultEl.querySelector('#sso-wrong-btn').onclick = () => {
            resultEl.style.display = 'none';
            if (manualEl) { manualEl.style.display = 'block'; const inp = root.querySelector('#login-name'); if (inp) inp.focus(); }
          };
        }
      } else {
        // Not authorised
        if (detectEl) detectEl.style.display = 'none';
        if (resultEl) {
          resultEl.style.display = 'block';
          resultEl.innerHTML =
            '<div style="background:rgba(248,81,73,.08);border:1px solid rgba(248,81,73,.2);border-radius:12px;padding:18px;text-align:center;margin-bottom:10px">' +
              '<div style="font-size:18px;font-weight:800;color:#f85149;margin-bottom:6px">⛔ Access Denied</div>' +
              '<div style="font-size:13px;color:rgba(255,255,255,.4);line-height:1.6">Your account <strong style="color:rgba(255,255,255,.6)">' + ssoUser.displayName + '</strong> is not registered. Contact your admin.</div>' +
            '</div>' +
            '<button id="sso-wrong-btn" style="width:100%;padding:9px;border:1px solid rgba(255,255,255,.1);border-radius:10px;background:transparent;color:rgba(255,255,255,.3);font-size:.82rem;cursor:pointer;font-family:inherit">Try a different login</button>';
          resultEl.querySelector('#sso-wrong-btn').onclick = () => {
            resultEl.style.display = 'none';
            if (manualEl) { manualEl.style.display = 'block'; }
          };
        }
      }
    }).catch(() => {
      // SSO failed — show manual input
      if (detectEl) detectEl.style.display = 'none';
      if (manualEl) { manualEl.style.display = 'block'; const inp = root.querySelector('#login-name'); if (inp) inp.focus(); }
    });
  }

  function doLoginWithRole(name, role) {
    // Strict role enforcement
    const isKnownAdmin = ADMINS.includes(name);
    const isKnownAssoc = ASSOCIATES.includes(name) || name === 'Admin-tarun';
    // Enforce correct role
    if (role === 'admin' && !isKnownAdmin) { toast('⛔ ' + name + ' is not an admin', 'err'); return; }
    if (role === 'associate' && isKnownAdmin && name !== 'Admin-tarun') { role = 'admin'; } // redirect admins
    if (!isKnownAdmin && !isKnownAssoc)    { toast('⛔ Not authorised: ' + name, 'err'); return; }
    currentRole = role;
    // For associates: load only own submissions
    if (role === 'associate') {
      const userKey = 'dtr_subs_' + name.toLowerCase().replace(/\s+/g,'_');
      // Normalise key — try both formats so old entries still load
      const normName = name.trim().toLowerCase().replace(/[^a-z0-9]+/g,'_');
      const userKeyOld = 'dtr_subs_' + name.toLowerCase().replace(/[^a-z0-9]+/g,'_');
      const userKeyNew = 'dtr_subs_' + normName;
      const loaded1 = safeLoad(userKeyNew, []);
      const loaded2 = userKeyNew !== userKeyOld ? safeLoad(userKeyOld, []) : [];
      const merged  = [...loaded1];
      loaded2.forEach(s => { if (!merged.find(m => m.submittedAt===s.submittedAt && m.taskType===s.taskType)) merged.push(s); });
      submissions = merged.filter(s => s && s.employeeName &&
        s.employeeName.toLowerCase().trim() === name.toLowerCase().trim());
    } else {
      submissions = [];
    }
    const sess = saveSession(name);
    authState  = { loggedIn:true, name, sid:sess.sid };
    // Apply mandatory national holidays for associates
    if (role !== 'admin') { applyMandatoryHolidays(name); }
    buildApp();
    toast('Welcome, ' + name + (role==='admin'?' 🛡':'') + '!', 'ok');
    // Load saved shift preference for this associate
    if (role !== 'admin') {
      authState.shift = loadShift(name);
    }
    // Pre-fetch SP list types in background
    setTimeout(async () => {
      await getListType(SP.TASK_LIST,   'task');
      await getListType(SP.STATUS_LIST, 'att');
      await getListType(SP.NPT_LIST,    'npt');
      console.log('[WorkPulse] SP list types cached:', JSON.stringify(_listTypes));
    }, 1500);
    // Sync data
    if (role === 'admin') {
      teamStatusCache_adm = safeLoadObj('dtr_teamcache',{});
      setTimeout(() => fetchAllAdminData(), 2000);
    } else {
      setTimeout(() => fetchOwnAttendance(true), 1000);
      setTimeout(() => fetchOwnTasks(true),       2500);
      setTimeout(() => fetchOwnNPT(true),         4000);
    }
  }

  function doLogin() {
    const el  = root.querySelector('#login-name');
    const raw = (el ? el.value : '').trim();
    const err = root.querySelector('#login-err');
    if (!raw) { if (err) err.textContent = 'Enter your login name.'; return; }
    const adminHit = ADMINS.find(a => a.toLowerCase() === raw.toLowerCase());
    const assocHit = ASSOCIATES.find(a => a.toLowerCase() === raw.toLowerCase());
    if (adminHit)      doLoginWithRole(adminHit, 'admin');
    else if (assocHit) doLoginWithRole(assocHit, 'associate');
    else { if (err) err.textContent = '⛔ "' + raw + '" is not an authorised login.'; }
  }

  function doLogout() {
    const name = authState.name;
    clearSession();
    clearDigestCache();
    submissions = [];
    authState   = { loggedIn:false, name:'', sid:'' };
    currentRole = 'associate';
    renderLogin();
    toast('Logged out — ' + name, 'info');
  }

  function showDevicePicker(onDone) {
    const isLt   = root && root.classList.contains('lt');
    const popBg  = isLt ? '#ffffff' : '#12151f';
    const bord   = isLt ? 'rgba(0,0,0,.09)' : 'rgba(255,255,255,.1)';
    const shadow = isLt ? '0 8px 40px rgba(0,0,0,.18)' : '0 24px 80px rgba(0,0,0,.9)';
    const ov = document.createElement('div');
    ov.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,.6);z-index:2147483646;display:flex;align-items:center;justify-content:center;backdrop-filter:blur(10px)';
    ov.innerHTML =
      '<div style="background:'+popBg+';border:1px solid '+bord+';border-radius:20px;width:min(460px,92vw);overflow:hidden;box-shadow:'+shadow+';font-family:-apple-system,sans-serif;animation:devicePopIn .25s cubic-bezier(.34,1.56,.64,1)">' +
        '<div style="height:3px;background:linear-gradient(90deg,#4f9eff,#a371f7)"></div>' +
        '<div style="padding:28px 24px">' +
          '<div style="text-align:center;margin-bottom:6px">' +
            '<div style="font-size:1.1rem;font-weight:800;background:linear-gradient(90deg,#4f9eff,#a371f7);-webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text;">Welcome to WorkPulse</div>' +
            '<div style="font-size:.82rem;color:var(--text3);margin-top:4px">Select your device for the best experience</div>' +
          '</div>' +
          '<div style="display:flex;gap:14px;margin:22px 0">' +
            '<div class="device-card" id="dev-desktop">' +
              '<svg width="48" height="48" viewBox="0 0 48 48" fill="none"><rect x="4" y="6" width="40" height="28" rx="3" stroke="#4f9eff" stroke-width="2" fill="rgba(79,158,255,.08)"/><path d="M16 40h16M24 34v6" stroke="#4f9eff" stroke-width="2" stroke-linecap="round"/><rect x="9" y="11" width="30" height="18" rx="1" fill="rgba(79,158,255,.12)"/></svg>' +
              '<div class="dc-label">Desktop</div>' +
              '<div class="dc-sub">Large monitor<br>+40% zoom (140%)</div>' +
            '</div>' +
            '<div class="device-card" id="dev-laptop">' +
              '<svg width="48" height="48" viewBox="0 0 48 48" fill="none"><rect x="6" y="8" width="36" height="24" rx="2.5" stroke="#a371f7" stroke-width="2" fill="rgba(163,113,247,.08)"/><path d="M2 34h44" stroke="#a371f7" stroke-width="2" stroke-linecap="round"/><path d="M18 34v2a2 2 0 002 2h8a2 2 0 002-2v-2" stroke="#a371f7" stroke-width="1.5" stroke-linecap="round"/><rect x="11" y="13" width="26" height="14" rx="1" fill="rgba(163,113,247,.12)"/></svg>' +
              '<div class="dc-label">Laptop</div>' +
              '<div class="dc-sub">Standard screen<br>Default zoom (100%)</div>' +
            '</div>' +
          '</div>' +
          '<button id="dev-confirm" disabled style="width:100%;padding:13px;border-radius:12px;border:none;background:linear-gradient(135deg,#4f9eff,#a371f7);color:#fff;font-size:.95rem;font-weight:700;cursor:not-allowed;opacity:.5;font-family:inherit;transition:all .2s">Confirm & Continue</button>' +
        '</div>' +
      '</div>';

    const rr = document.getElementById('dtr-root-outer') || document.body;
    rr.appendChild(ov);
    let chosen = null;
    const btn = ov.querySelector('#dev-confirm');

    ['desktop','laptop'].forEach(id => {
      ov.querySelector('#dev-'+id).addEventListener('click', () => {
        ov.querySelectorAll('.device-card').forEach(el => el.classList.remove('selected'));
        ov.querySelector('#dev-'+id).classList.add('selected');
        chosen = id;
        btn.disabled = false;
        btn.style.opacity = '1';
        btn.style.cursor  = 'pointer';
      });
    });
    btn.addEventListener('click', () => {
      ov.remove();
      const outer = document.getElementById('dtr-root-outer');
      if (outer) {
        if (chosen === 'desktop') {
          outer.style.zoom = '1.4';
        } else {
          outer.style.zoom = '1';
        }
      }
      GM_setValue('dtr_device', chosen);
      if (onDone) onDone();
    });
  }

  function buildApp() {
    const isAdmin     = currentRole === 'admin';
    const isSuperAdmin= authState.name === SUPER_ADMIN || authState.name === 'Admin-tarun';
    const activeTabs  = isAdmin ? ADMIN_TABS : ASSOC_TABS;
    const ms = mySubmissions();
    const todaySt = statusCache[todayStr()];

    let sbHtml = '', grp = '';
    activeTabs.forEach(t => {
      if (t.group !== grp) { grp = t.group; sbHtml += '<div class="sb-sec">' + grp + '</div>'; }
      sbHtml += '<button class="sb-item" data-action="switch-tab" data-val="' + t.id + '">' + t.icon + ' ' + t.label + '</button>';
    });

    // Role badge style
    const roleBadgeStyle = isAdmin
      ? 'background:rgba(163,113,247,.15);color:#a371f7'
      : 'background:rgba(88,166,255,.12);color:var(--accent)';
    const roleLabel = isAdmin ? '🛡 Admin' : 'Associate';

    // Switch view button (SuperAdmin only)
    const switchBtn = isSuperAdmin
      ? '<button class="icon-btn" data-action="switch-role" title="' + (isAdmin?'Switch to Associate View':'Switch to Admin Portal') + '" style="font-size:.7rem;width:auto;padding:0 8px;gap:3px">' +
        (isAdmin ? '👤' : '🛡') + '</button>'
      : '';

    // Sidebar footer — show team stats in admin, own stats for associate
    const sidebarFooter = isAdmin
      ? '<div class="sb-stat"><div class="lbl">Members</div><div class="val" id="sb-members">' + getAllMembers().length + '</div></div>' +
        '<div class="sb-stat"><div class="lbl">Total Entries</div><div class="val" id="sb-total">' + (Array.isArray(submissions)?submissions.length:0) + '</div></div>'
      : '<div class="sb-stat"><div class="lbl">My Entries</div><div class="val" id="sb-total">' + ms.length + '</div></div>' +
        '<div class="sb-stat"><div class="lbl">Productivity</div><div class="val" style="color:var(--green)" id="sb-avg">' + calcAvgProd(ms).toFixed(0) + '%</div></div>';

    root.innerHTML =
      '<div id="wp-bg-orb"><span class="orb o1"></span><span class="orb o2"></span><span class="orb o3"></span></div>' +
      '<div id="dtr-topbar">' +
        '<div class="dtr-logo">' +
        '<div class="dtr-logo-icon">' +
          '<svg width="40" height="38" viewBox="0 0 40 38" fill="none">' +
            '<defs>' +
              '<linearGradient id="g1" x1="0" y1="0" x2="1" y2="1"><stop stop-color="#4f9eff"/><stop offset="1" stop-color="#a371f7"/></linearGradient>' +
              '<linearGradient id="g2" x1="0" y1="1" x2="1" y2="0"><stop stop-color="#30d158"/><stop offset="1" stop-color="#4f9eff"/></linearGradient>' +
              '<linearGradient id="g3" x1="0" y1="0" x2="1" y2="0"><stop stop-color="#4f9eff"/><stop offset="1" stop-color="#a371f7"/></linearGradient>' +
            '</defs>' +
            '<rect x="1" y="4" width="30" height="33" rx="3" fill="rgba(79,158,255,.08)" stroke="url(#g1)" stroke-width="1.2"/>' +
            '<rect x="9" y="1" width="12" height="6" rx="2" fill="url(#g1)" opacity=".9"/>' +
            '<rect x="11" y="0" width="8" height="4" rx="1.5" fill="rgba(79,158,255,.25)" stroke="url(#g1)" stroke-width=".8"/>' +
            '<line x1="1" y1="14" x2="31" y2="14" stroke="rgba(79,158,255,.2)" stroke-width=".8"/>' +
            '<line x1="1" y1="20" x2="31" y2="20" stroke="rgba(79,158,255,.12)" stroke-width=".6"/>' +
            '<line x1="1" y1="26" x2="31" y2="26" stroke="rgba(79,158,255,.12)" stroke-width=".6"/>' +
            '<line x1="11" y1="14" x2="11" y2="37" stroke="rgba(79,158,255,.12)" stroke-width=".6"/>' +
            '<line x1="21" y1="14" x2="21" y2="37" stroke="rgba(79,158,255,.12)" stroke-width=".6"/>' +
            '<rect x="2" y="7" width="28" height="7" rx="1" fill="url(#g3)" opacity=".15"/>' +
            '<rect x="4" y="9.5" width="5" height="2" rx="1" fill="url(#g1)" opacity=".8"/>' +
            '<rect x="13" y="9.5" width="5" height="2" rx="1" fill="url(#g1)" opacity=".6"/>' +
            '<rect x="23" y="9.5" width="5" height="2" rx="1" fill="url(#g1)" opacity=".6"/>' +
            '<rect x="3" y="16" width="6" height="2.5" rx="1" fill="url(#g2)" opacity=".9"/>' +
            '<rect x="13" y="16" width="4" height="2.5" rx="1" fill="url(#g1)" opacity=".8"/>' +
            '<rect x="23" y="16" width="5" height="2.5" rx="1" fill="rgba(163,113,247,.8)"/>' +
            '<rect x="3" y="22" width="7" height="2.5" rx="1" fill="url(#g1)" opacity=".7"/>' +
            '<rect x="13" y="22" width="3" height="2.5" rx="1" fill="url(#g2)" opacity=".8"/>' +
            '<rect x="23" y="22" width="6" height="2.5" rx="1" fill="url(#g1)" opacity=".7"/>' +
            '<rect x="3" y="28" width="5" height="2.5" rx="1" fill="rgba(163,113,247,.7)"/>' +
            '<rect x="13" y="28" width="6" height="2.5" rx="1" fill="url(#g2)" opacity=".7"/>' +
            '<rect x="23" y="28" width="4" height="2.5" rx="1" fill="url(#g1)" opacity=".8"/>' +
            '<rect x="34" y="10" width="5" height="24" rx="1" fill="rgba(79,158,255,.08)" stroke="url(#g1)" stroke-width=".8"/>' +
            '<rect x="35" y="26" width="3" height="7" rx=".5" fill="url(#g2)"/>' +
            '<rect x="35" y="20" width="3" height="5" rx=".5" fill="url(#g1)" opacity=".6"/>' +
            '<rect x="35" y="15" width="3" height="4" rx=".5" fill="rgba(163,113,247,.5)"/>' +
            '<circle cx="28" cy="32" r="4.5" fill="url(#g2)" opacity=".95"/>' +
            '<path d="M25.8 32L27.3 33.6L30.2 30.4" stroke="#fff" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round"/>' +
          '</svg>' +
        '</div>' +
        '<div class="dtr-logo-text">' +
          '<div class="dtr-logo-team">OBX</div>' +
          '<div class="dtr-logo-name">WorkPulse</div>' +
        '</div>' +
      '</div>' +
        '<div class="dtr-tabs">' + activeTabs.map(t => '<button class="dtr-tab" data-action="switch-tab" data-val="' + t.id + '">' + t.icon + ' ' + t.label + '</button>').join('') + '</div>' +
        '<div class="dtr-topbar-r">' +
          (todaySt && !isAdmin ? '<span class="sp sp-' + (todaySt.status||'').toLowerCase().replace(' ','-') + '" style="font-size:.75rem">' + todaySt.status + '</span>' : '') +
          '<div class="dtr-user-pill">' +
          '<div class="av">' + authState.name[0].toUpperCase() + '</div>' +
          '<span style="background:linear-gradient(90deg,#4f9eff,#a371f7,#4f9eff);background-size:200% auto;-webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text;animation:shimmerText 3s linear infinite;font-weight:700">' + authState.name + '</span>' +
            '<span class="role-badge" style="' + roleBadgeStyle + '">' + roleLabel + '</span>' +
          '</div>' +
          switchBtn +
          '<button class="icon-btn" data-action="toggle-theme" id="theme-btn">' + (theme==='dark'?ic.sun:ic.moon) + '</button>' +
          '<button class="icon-btn danger" data-action="do-logout" title="Logout">' + ic.logout + '</button>' +
          '<button class="icon-btn danger" data-action="close-app">' + ic.close + '</button>' +
        '</div>' +
      '</div>' +
      '<div id="dtr-body">' +
        '<div id="dtr-sidebar">' + sbHtml +
          '<hr class="sep" style="margin:8px 0"><div class="sb-footer">' + sidebarFooter + '</div>' +
        '</div>' +
        '<div id="dtr-main">' + activeTabs.map(t => '<div class="dtr-view anim" id="view-' + t.id + '" style="display:none"></div>').join('') + '</div>' +
      '</div><div id="dtr-toast"></div>';

    switchTab(activeTabs[0].id);
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
    if (currentRole === 'admin') {
      if (v === 'overview')      renderAdminOverview();
      if (v === 'teamtrack')     renderAdminTeamTracker();
      if (v === 'admanalytics')  renderAdminAnalytics();
      if (v === 'attweek')   renderAdminWeekView();
      if (v === 'attnpt')    renderAdminNPTLog();
      if (v === 'settings')  renderSettings();
    } else {
      if (v === 'submit')    renderSubmit();
      if (v === 'analytics') renderAnalytics();
      if (v === 'tracker')   renderTracker();
      if (v === 'weekly')    renderWeeklyCalendar();   // attendance editing
      if (v === 'calendar')  renderMyCalendar();       // view-only
      if (v === 'missednpt') renderMissedNPT();
      if (v === 'settings')  renderSettings();
    }
  }

  function toggleTheme() {
    theme = theme === 'dark' ? 'light' : 'dark';
    GM_setValue('dtr_theme4', theme);
    root.classList.toggle('lt', theme === 'light');
    const b = q('#theme-btn'); if (b) b.innerHTML = theme === 'dark' ? ic.sun : ic.moon;
  }

  function handleClick(e) {
    const btn = e.target.closest('[data-action]'); if (!btn) return;
    const a = btn.dataset.action;
    const v = btn.dataset.val !== undefined ? btn.dataset.val : (btn.dataset.idx !== undefined ? btn.dataset.idx : '');
    switch (a) {
      case 'do-login':      doLogin(); break;
      case 'do-logout':     doLogout(); break;
      case 'switch-tab':    switchTab(v); break;
      case 'toggle-theme':  toggleTheme(); break;
      case 'close-app':     root.style.display = 'none'; break;
      case 'switch-role': {
        // SuperAdmin only — toggle between admin and associate view
        if (authState.name !== SUPER_ADMIN && authState.name !== 'Admin-tarun') break;
        if (currentRole === 'admin') {
          currentRole = 'associate';
          const oldName = authState.name;
          authState.name = 'Admin-tarun';
          saveSession('Admin-tarun');
          submissions = safeLoad('dtr_subs_admin-tarun', []).filter(s => s && s.employeeName === 'Admin-tarun');
          buildApp();
          toast('Switched to Associate view (Admin-tarun)', 'info');
          setTimeout(() => fetchOwnAttendance(true), 800);
        } else {
          currentRole = 'admin';
          authState.name = SUPER_ADMIN;
          saveSession(SUPER_ADMIN);
          submissions = [];
          teamStatusCache_adm = safeLoadObj('dtr_teamcache', {});
          buildApp();
          toast('Switched to Admin Portal', 'info');
          setTimeout(() => fetchAllAdminData(), 1000);
        }
        break;
      }
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
        attDate = ws2.getFullYear()+'-'+String(ws2.getMonth()+1).padStart(2,'0')+'-'+String(ws2.getDate()).padStart(2,'0');
        attSelectedStatus = (statusCache[attDate]||{}).status||'';
        renderWeeklyCalendar();
        break;
      }
      case 'att-week-next': {
        attWeekOffset++;
        const ws3 = getWeekStart(attWeekOffset);
        ws3.setDate(ws3.getDate()+1); // Monday
        attDate = ws3.getFullYear()+'-'+String(ws3.getMonth()+1).padStart(2,'0')+'-'+String(ws3.getDate()).padStart(2,'0');
        attSelectedStatus = (statusCache[attDate]||{}).status||'';
        renderWeeklyCalendar();
        break;
      }
      case 'att-week-today':   attWeekOffset=0; attDate=todayStr(); attSelectedStatus=(statusCache[attDate]||{}).status||''; renderWeeklyCalendar(); break;
      case 'att-toggle-bulk':  attToggleBulk(); break;
      case 'att-multi-pick':   attMultiPick(v); break;
      case 'att-bulk-apply':   attBulkApply(v); break;
      case 'att-clear-select':
        attMultiSelect.clear();
        qa('.att-day-card').forEach(c => { c.classList.remove('att-multi'); const chk=c.querySelector('.att-multi-chk'); if(chk) chk.remove(); });
        _updateBulkBar();
        break;

      case 'toggle-note': {
        const n = btn.dataset.n;
        const box = document.getElementById('opt-note-' + n);
        const lbl = btn.querySelector('.opt-note-lbl');
        if (box) { const shown = box.style.display === 'block'; box.style.display = shown ? 'none' : 'block'; if(lbl) lbl.textContent = shown ? 'Add optional note' : 'Hide note'; }
        break;
      }
      case 'att-pick-day': {
        // Block weekends — cannot mark Sat/Sun
        const _dow = new Date(v+'T12:00:00').getDay();
        if (_dow === 0 || _dow === 6) { toast('Saturday & Sunday are mandatory off days', 'err'); break; }
        // Block mandatory national holidays
        if (isMandatoryOff(v)) { toast('🇮🇳 ' + (statusCache[v]?.name||'National Holiday') + ' — mandatory off, cannot be changed', 'err'); break; }
        attDate = v;
        const _saved = statusCache[v] || {};
        // Strip sub-type suffix when restoring base status
        attSelectedStatus = (_saved.status||'').split(' - ')[0];
        attLeaveDay       = _saved.leaveDay  || '';
        attLeaveHalf      = _saved.leaveHalf || '';
        // Highlight selected day card
        qa('.att-day-card').forEach(c => c.classList.toggle('att-day-selected', c.dataset.val === v));
        // Re-render only the edit panel (attSelectedStatus already set correctly above)
        const p = q('#att-edit-panel');
        if (p) p.innerHTML = buildAttEditPanel(v);
        break;
      }
      case 'att-nav-next':  attNavDate(+1); break;
      case 'att-nav-today': attNavToday(); break;
      case 'toggle-opt-note': {
        const n = btn.dataset.n;
        const box = document.getElementById('opt-note-' + n);
        const lbl = btn.querySelector('.opt-note-lbl');
        if (box) { const shown = box.style.display==='block'; box.style.display=shown?'none':'block'; if(lbl) lbl.textContent=shown?'Add optional note':'Hide note'; }
        break;
      }
      case 'att-status':      attSelectStatus(v); break;
      case 'att-leave-day':   attSelectLeaveDay(v); break;
      case 'att-leave-half':  attSelectLeaveHalf(v); break;
      case 'att-save':      attSave(); break;
      case 'att-shift-change': {
        const _shiftEl = q('#att-shift');
        if (_shiftEl && _shiftEl.value && authState.name) {
          authState.shift = _shiftEl.value;
          saveShift(authState.name, _shiftEl.value);
        }
        break;
      }
      case 'cal-sync':      toast('Loading your attendance...','info'); fetchOwnAttendance(false); break;
      case 'cal-prev':      calMonthOffset--; renderMyCalendar(); break;
      case 'cal-next':      calMonthOffset++; renderMyCalendar(); break;
      case 'cal-today':     calMonthOffset=0; renderMyCalendar(); break;
      case 'npt-type':      nptSelectType(v); break;
      case 'npt-log':       nptLogEntry(); break;
      case 'npt-del':       nptDel(+btn.dataset.idx); break;
      case 'test-sp':        testSP(); break;
      case 'test-npt':       testNPTConnection(); break;
      case 'sync-my-data':   toast('Syncing your data from SP...','info'); fetchOwnAttendance(false); setTimeout(()=>fetchOwnTasks(false),1500); setTimeout(()=>fetchOwnNPT(false),3000); break;
      case 'flush-queue':   flushQueueManual(); break;
      case 'adm-tk-clear':    { admSelectedAssocs.clear(); renderAdminTeamTracker(); break; }
      case 'adm-assoc-clear': { admSelectedAssocs.clear(); renderAdminTeamTracker(); break; }
      case 'adm-assoc-toggle': {
        if (admSelectedAssocs.has(v)) admSelectedAssocs.delete(v);
        else admSelectedAssocs.add(v);
        renderAdminTeamTracker(); break;
      }
      case 'adm-assoc-remove': { admSelectedAssocs.delete(v); renderAdminTeamTracker(); break; }
      case 'adm-edit-entry': {
        const idx2 = parseInt(v);
        if (!isNaN(idx2) && idx2>=0 && idx2<submissions.length) {
          showAdminEditEntry(idx2);
        }
        break;
      }
      case 'adm-del-entry': {
        const idx3 = parseInt(v);
        if (!isNaN(idx3) && idx3>=0 && idx3<submissions.length) {
          showAdminDeleteEntry(idx3);
        }
        break;
      }
      case 'adm-add-leave': { showAdminAddLeave(); break; }
      case 'aan-select':  { aanSelectedAssoc=v; aanMode='single'; renderAdminAnalytics(); break; }
      case 'aan-compare': { if(aanCompareSet.has(v))aanCompareSet.delete(v);else if(aanCompareSet.size<4)aanCompareSet.add(v); renderAdminAnalytics(); break; }
      case 'aan-mode':    { if(['single','compare','weekly','monthly'].includes(v)) aanMode=v; renderAdminAnalytics(); break; }
      // Admin actions
      case 'adm-sync-all':   fetchAllAdminData(); break;
      case 'adm-sync-tasks': fetchAdminTasks(); break;
      case 'adm-sync-att':   fetchAdminAttendance(); break;
      case 'adm-sync-npt':   fetchAdminNPT(); break;
      case 'adm-wv-prev':    adminWeekOffset--; renderAdminWeekView(); break;
      case 'adm-wv-next':    adminWeekOffset++; renderAdminWeekView(); break;
      case 'adm-wv-today':   adminWeekOffset=0;  renderAdminWeekView(); break;
      case 'adm-att-edit':   { showAdminAttPicker(v); break; }
      case 'adm-att-set':    { /* handled via direct onclick in showAdminAttPicker */ break; }
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
    if (e.key === 'Enter' && e.target.tagName === 'SELECT') doLogin();
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
      '<div class="dtr-field"><label class="dtr-label">Full Name</label><input class="dtr-input" value="' + authState.name + '" readonly style="cursor:default"></div>' +
      '<div class="dtr-field"><label class="dtr-label">Shift</label><select id="a-shift" class="dtr-select">' + (function(sv){ return '<option value="8-5"'+(sv==='8-5'?' selected':'')+'>8 AM – 5 PM</option>'+'<option value="9-6"'+(sv==='9-6'?' selected':'')+'>9 AM – 6 PM</option>'+'<option value="10-7"'+(sv==='10-7'?' selected':'')+'>10 AM – 7 PM</option>'+'<option value="11-8"'+(sv==='11-8'?' selected':'')+'>11 AM – 8 PM</option>';})(authState.shift||loadShift(authState.name)) + '</select></div>' +
      '<div class="dtr-field"><label class="dtr-label">Date <span style="font-size:.73rem;color:var(--text3)">(Today or Yesterday only)</span></label>' +
      '<div class="date-pill-row">' +
      '<button class="date-pill' + (selectedDate===todayStr()?' active':'') + '" data-action="pick-date" data-val="' + todayStr() + '">Today — ' + formatDate(todayStr()) + '</button>' +
      '<button class="date-pill' + (selectedDate===prevWorkingDayStr()?' active':'') + '" data-action="pick-date" data-val="' + prevWorkingDayStr() + '">' + prevWorkingDayLabel() + ' — ' + formatDate(prevWorkingDayStr()) + '</button>' +
      '</div><input type="hidden" id="a-date" value="' + selectedDate + '"></div></div></div>' +
      '<div class="card"><div class="card-title" style="display:flex;justify-content:space-between;align-items:center">' +
      '<span>Tasks</span><div class="mode-toggle"><button class="mode-btn active" data-action="mode-single">Single</button>' +
      '<button class="mode-btn" data-action="mode-multi">Multi-Task</button></div></div>' +
      '<div id="day-limit-warn" style="display:none;color:var(--red);font-size:.82rem;font-weight:600;padding:6px 10px;background:rgba(248,81,73,.08);border-radius:8px;margin-bottom:8px"></div>' +
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
      '<div class="dtr-field"><label class="dtr-label">Minutes <span style="font-size:.71rem;color:var(--text3)">(max 480)</span></label><input type="number" class="dtr-input dtr-hrs" min="0" max="480" step="1" placeholder="e.g. 240"></div>' +
      '<div class="dtr-field"><label class="dtr-label">NPT Mins</label><div class="npt-box">—</div></div></div>' +
      '<div class="comments-box" id="req-comment-' + n + '"><div class="dtr-field" style="margin-top:4px"><label class="dtr-label dtr-comment-label">Comments <span class="req-star">*</span></label><textarea class="dtr-input dtr-adhoc" rows="2" placeholder="Add details..." style="resize:vertical;min-height:52px"></textarea></div></div>' +'<div class="opt-note-toggle" data-action="toggle-note" data-n="' + n + '">📝 <span class="opt-note-lbl">Add optional note</span></div><div class="opt-note-box" id="opt-note-' + n + '" style="display:none"><div class="dtr-field" style="margin-top:6px"><label class="dtr-label">Note <span style="color:var(--text3);font-weight:400">(optional)</span></label><textarea class="dtr-input dtr-note" rows="2" placeholder="Any additional context..." style="resize:vertical;min-height:48px"></textarea></div></div>' +
      '<div class="leave-tag"><span>Leave:</span>' +
      '<button class="leave-type-btn active" data-action="leave-type" data-val="full">Full Day (480 min)</button>' +
      '<button class="leave-type-btn" data-action="leave-type" data-val="half-am">Half Day (240 min)</button>' +
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
  function makeLeave(card) { if (!card) return; card.classList.add('leave'); card.querySelector('.leave-tag').classList.add('show'); const ws = card.querySelector('.dtr-wtype'); if (ws) ws.value = 'NPT'; const hi = card.querySelector('.dtr-hrs'); if (hi) hi.value = '480'; const nb = card.querySelector('.npt-box'); if (nb) { nb.textContent = '480 min NPT (full day)'; nb.className = 'npt-box'; } const ts = card.querySelector('.dtr-tt'); if (ts) ts.value = 'Leave'; updateDayViz(); }
  function setLeaveType(card, type) { if (!card) return; card.querySelectorAll('.leave-type-btn').forEach(b => b.classList.toggle('active', b.dataset.val === type)); const m = type === 'full' ? 480 : 240; const ld = card.querySelector('.leave-hours-display'); if (ld) { ld.textContent = m + ' min'; } const nb = card.querySelector('.npt-box'); if (nb) nb.textContent = m + ' min NPT'; updateDayViz(); }
  function recalcNPT() {
    // Validate total doesn't exceed 480 mins
    let totalMins = 0;
    qa('.task-card').forEach(card => {
      if (card.classList.contains('leave')) { totalMins += 480; return; }
      const mins = parseInt(card.querySelector('.dtr-hrs')?.value)||0;
      totalMins += mins;
    });
    const overLimit = totalMins > WH;
    // Update each card's NPT display
    qa('.task-card').forEach(card => {
      if (card.classList.contains('leave')) return;
      const mins = parseInt(card.querySelector('.dtr-hrs')?.value)||0;
      const wt   = card.querySelector('.dtr-wtype')?.value||'Productive';
      const nb   = card.querySelector('.npt-box');
      const inp  = card.querySelector('.dtr-hrs');
      if (inp) inp.style.borderColor = overLimit ? 'var(--red)' : '';
      if (!nb) return;
      if (mins === 0) { nb.textContent='—'; nb.className='npt-box'; return; }
      if (wt === 'Productive') {
        const npt = Math.max(0, WH - mins);
        nb.textContent = npt + ' min';
        nb.className = 'npt-box prod';
      } else {
        nb.textContent = mins + ' min';
        nb.className = 'npt-box';
      }
    });
    // Show total warning
    const warnEl = q('#day-limit-warn');
    if (warnEl) {
      warnEl.style.display = overLimit ? 'block' : 'none';
      warnEl.textContent = '⚠️ Total ' + totalMins + ' mins exceeds 480 mins/day limit';
    }
  }
  function updateDayViz() {
    let prod=0, npt=0;
    qa('.task-card').forEach(card => {
      if (card.classList.contains('leave')) { npt += 480; return; }
      const mins = parseInt(card.querySelector('.dtr-hrs')?.value)||0;
      const wt   = card.querySelector('.dtr-wtype')?.value||'Productive';
      if (wt === 'NPT') npt += mins;
      else { prod += mins; npt += Math.max(0, WH - mins); }
    });
    npt  = Math.min(WH, Math.max(0, npt));
    prod = Math.min(WH, prod);
    const pct   = m => Math.round(m/WH*100);
    const toHM  = m => { const h=Math.floor(m/60),mn=m%60; return h>0?(mn>0?h+'h '+mn+'m':h+'h'):mn+'m'; };
    const pEl=q('#dviz-prod'),nEl=q('#dviz-npt'),pv=q('#dviz-prod-val'),nv=q('#dviz-npt-val'),sv=q('#dviz-summary');
    if (pEl) pEl.style.width = pct(prod)+'%';
    if (nEl) nEl.style.width = pct(npt)+'%';
    if (pv)  pv.textContent  = toHM(prod);
    if (nv)  nv.textContent  = toHM(npt);
    if (sv)  sv.textContent  = toHM(prod)+' prod + '+toHM(npt)+' NPT / 8h (480 mins)';
  }

  function showSubmitConfirm(tasks, dateVal, shift) {
    return new Promise(resolve => {
      const isLt   = root && root.classList.contains('lt');
      const popBg  = isLt ? '#ffffff' : '#12151f';
      const rowBg  = isLt ? '#f5f5f7' : '#1c2030';
      const bord   = isLt ? 'rgba(0,0,0,.08)' : 'rgba(255,255,255,.1)';
      const textC  = isLt ? '#1c1c1e' : '#f0f4ff';
      const text2C = isLt ? 'rgba(28,28,30,.5)' : 'rgba(240,244,255,.5)';
      const shadow = isLt ? '0 8px 32px rgba(0,0,0,.15)' : '0 24px 70px rgba(0,0,0,.9)';
      const dateDisp = new Date(dateVal+'T12:00:00').toLocaleDateString('en-IN',{weekday:'long',day:'numeric',month:'long',year:'numeric'});

      const prodTasks = tasks.filter(t => t.workType !== 'NPT' && t.taskType !== 'Leave');
      const nptTasks  = tasks.filter(t => t.workType === 'NPT');
      const totalProdMins = prodTasks.reduce((a,t) => a+(t.minutes||0), 0);
      const totalNPTMins  = nptTasks.reduce((a,t)  => a+(t.minutes||0), 0);
      const remainingMins = Math.max(0, 480 - totalProdMins - totalNPTMins);
      const toHM = m => { const h=Math.floor(m/60),mn=m%60; return h>0?(mn>0?h+'h '+mn+'m':h+'h'):(mn?mn+'m':'0m'); };

      const taskRows = tasks.map((t,i) => {
        const isNPT  = t.workType === 'NPT';
        const dotClr = isNPT ? '#ff9f0a' : '#30d158';
        const mins   = t.minutes || 0;
        return '<div style="display:flex;align-items:center;gap:10px;padding:8px 12px;background:'+rowBg+';border-radius:8px;margin-bottom:6px">' +
          '<span style="width:8px;height:8px;border-radius:50%;background:'+dotClr+';flex-shrink:0;display:inline-block"></span>' +
          '<div style="flex:1">' +
            '<div style="font-size:.85rem;font-weight:700;color:'+textC+'">'+t.taskType+'</div>' +
            '<div style="font-size:.73rem;color:'+text2C+'">'+t.workType+' · '+toHM(mins)+'</div>' +
          '</div>' +
          (t.adhoc ? '<div style="font-size:.72rem;color:'+text2C+';max-width:120px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="'+t.adhoc.replace(/"/g,"&quot;")+'">'+t.adhoc+'</div>' : '') +
        '</div>';
      }).join('');

      const ov = document.createElement('div');
      ov.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,.55);z-index:2147483646;display:flex;align-items:center;justify-content:center;backdrop-filter:blur(8px)';
      ov.innerHTML =
        '<div style="background:'+popBg+';border:1px solid '+bord+';border-radius:18px;width:min(420px,94vw);max-height:90vh;overflow-y:auto;box-shadow:'+shadow+';font-family:-apple-system,sans-serif">' +
          '<div style="height:3px;background:linear-gradient(90deg,#4f9eff,#a371f7)"></div>' +
          '<div style="padding:24px">' +
            // Header
            '<div style="margin-bottom:16px">' +
              '<div style="font-size:1rem;font-weight:800;color:'+textC+';letter-spacing:-.2px">Confirm Submission</div>' +
              '<div style="font-size:.8rem;color:'+text2C+';margin-top:3px">'+dateDisp+' · Shift '+shift+'</div>' +
            '</div>' +
            // Task list
            '<div style="margin-bottom:14px">' +
              '<div style="font-size:.72rem;font-weight:800;text-transform:uppercase;letter-spacing:.7px;color:'+text2C+';margin-bottom:8px">'+tasks.length+' Task'+(tasks.length>1?'s':'')+' to Submit</div>' +
              taskRows +
            '</div>' +
            // Summary bar
            '<div style="background:'+rowBg+';border-radius:10px;padding:12px 14px;margin-bottom:16px;display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px;text-align:center">' +
              '<div><div style="font-size:1.1rem;font-weight:800;color:#30d158">'+toHM(totalProdMins)+'</div><div style="font-size:.68rem;color:'+text2C+'">Productive</div></div>' +
              '<div><div style="font-size:1.1rem;font-weight:800;color:#ff9f0a">'+toHM(totalNPTMins)+'</div><div style="font-size:.68rem;color:'+text2C+'">NPT</div></div>' +
              '<div><div style="font-size:1.1rem;font-weight:800;color:'+(remainingMins>0?'#ef4444':'#30d158')+'">'+toHM(remainingMins)+'</div><div style="font-size:.68rem;color:'+text2C+'">Remaining</div></div>' +
            '</div>' +
            (remainingMins > 0
              ? '<div style="background:rgba(239,68,68,.08);border:1px solid rgba(239,68,68,.2);border-radius:8px;padding:10px 12px;margin-bottom:14px;font-size:.78rem;color:#ef4444">⚠️ <strong>'+toHM(remainingMins)+' unaccounted</strong> — remaining time will be auto-logged as NPT</div>'
              : '<div style="background:rgba(34,197,94,.08);border:1px solid rgba(34,197,94,.2);border-radius:8px;padding:10px 12px;margin-bottom:14px;font-size:.78rem;color:#22c55e">✅ Full 8h accounted for</div>') +
            // Buttons
            '<div style="display:flex;gap:10px">' +
              '<button id="sc-cancel" style="flex:1;padding:11px;border-radius:10px;border:1px solid '+bord+';background:'+rowBg+';color:'+text2C+';font-size:.9rem;font-weight:600;cursor:pointer;font-family:inherit">Cancel</button>' +
              '<button id="sc-confirm" style="flex:2;padding:11px;border-radius:10px;border:none;background:linear-gradient(135deg,#4f9eff,#a371f7);color:#fff;font-size:.9rem;font-weight:700;cursor:pointer;font-family:inherit">Submit to SharePoint</button>' +
            '</div>' +
          '</div>' +
        '</div>';

      const rr = document.getElementById('dtr-root-outer') || document.body;
      rr.appendChild(ov);
      ov.querySelector('#sc-cancel').onclick  = () => { ov.remove(); resolve(false); };
      ov.querySelector('#sc-confirm').onclick = () => { ov.remove(); resolve(true); };
    });
  }

  async function doSubmit() {
    // ── Collect form values ──────────────────────────────────────────────
    const dateVal = q('#a-date')?.value || todayStr();
    const shift   = q('#a-shift')?.value || authState.shift || '8-5';
    // Persist shift choice for this user
    if (shift && authState.name) {
      authState.shift = shift;
      saveShift(authState.name, shift);
    }
    const cards   = Array.from(qa('.task-card'));
    const tasks   = [];
    let hasErr    = false;

    cards.forEach(card => {
      const isLeave = card.classList.contains('leave');
      if (isLeave) {
        // Detect leave type from button state
        const activeBtn = card.querySelector('.leave-type-btn.active');
        const leaveType = activeBtn ? activeBtn.dataset.val : 'full';
        const leaveMins = leaveType === 'full' ? 480 : 240;  // full=480, half=240
        const leaveHrs  = parseFloat((leaveMins/60).toFixed(2));
        tasks.push({
          employeeName: authState.name,
          taskType:     'Leave',
          hours:        0,
          npt:          leaveHrs,           // hrs for SP (8.0 or 4.0)
          nptMinutes:   leaveMins,          // mins for display (480 or 240)
          minutes:      0,
          workType:     'NPT',
          leaveType,                        // 'full' | 'half-am' | 'half-pm'
          adhoc:        card.querySelector('.dtr-adhoc')?.value || '',
          shift,
          date:         dateVal,
          submittedAt:  nowUTC()
        });
        return;
      }

      const tt  = card.querySelector('.dtr-tt')?.value    || '';
      const wt  = card.querySelector('.dtr-wtype')?.value || 'Productive';
      const mins = parseInt(card.querySelector('.dtr-hrs')?.value) || 0;
      const reqComment = (card.querySelector('.dtr-adhoc')?.value || '').trim();
      const optNote    = (card.querySelector('.dtr-note')?.value   || '').trim();
      const adhoc = reqComment + (reqComment && optNote ? ' | ' + optNote : optNote || '');

      // Validation
      if (!tt) {
        toast('Select a task type for each row', 'err'); hasErr = true; return;
      }
      if (mins <= 0 && wt !== 'NPT') {
        toast('Enter minutes worked (must be > 0)', 'err'); hasErr = true; return;
      }
      if (mins > 480) {
        toast('Max 480 minutes (8h) per task', 'err'); hasErr = true; return;
      }
      if ((tt === 'Ad-hoc Tasks' || tt === 'Other' || tt === 'Meeting') && !reqComment) {
        toast('Please add a comment for ' + tt, 'err'); hasErr = true; return;
      }

      const nptMins = wt === 'NPT' ? mins : Math.max(0, WH - mins);

      tasks.push({
        employeeName: authState.name,
        taskType:     tt,
        hours:        parseFloat((mins/60).toFixed(2)),
        npt:          parseFloat((nptMins/60).toFixed(2)),
        minutes:      mins,
        nptMinutes:   nptMins,
        workType:     wt,
        adhoc,
        shift,
        date:         dateVal,
        submittedAt:  localISOString()
      });
    });

    if (hasErr || !tasks.length) {
      if (!hasErr) toast('No tasks to submit', 'err');
      return;
    }
    // Enforce 480 min/day total
    const totalDayMins = tasks.reduce((a,t) => a + (t.minutes||Math.round(t.hours*60)||0), 0);
    if (totalDayMins > 480) {
      toast('❌ Total ' + totalDayMins + ' mins exceeds the 480 min/day (8h) limit. Please adjust.', 'err');
      return;
    }

    // ── Duplicate check: warn if same taskType+date already in queue/submissions ──
    const dupWarnings = [];
    tasks.forEach(t => {
      const sameDate  = submissions.filter(s => s.date === t.date && s.employeeName === t.employeeName);
      const inQueue   = spQueue.filter(s => s.date === t.date && s.employeeName === t.employeeName);
      const allExist  = [...sameDate, ...inQueue];
      // Check same taskType with DIFFERENT hours (allow) vs same taskType+same hours (warn)
      const exact = allExist.find(s =>
        s.taskType === t.taskType &&
        Math.abs((s.minutes||Math.round((s.hours||0)*60)) - (t.minutes||Math.round((t.hours||0)*60))) < 5 &&
        Math.abs((s.nptMinutes||Math.round((s.npt||0)*60)) - (t.nptMinutes||Math.round((t.npt||0)*60))) < 5
      );
      if (exact) dupWarnings.push(t.taskType + ' on ' + t.date + ' (already submitted with same hours)');
    });
    if (dupWarnings.length > 0) {
      const proceed = await new Promise(resolve => {
        const isLt = root && root.classList.contains('lt');
        const dv = document.createElement('div');
        dv.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,.55);z-index:2147483646;display:flex;align-items:center;justify-content:center;backdrop-filter:blur(8px)';
        const bg = isLt?'#fff':'#12151f', brd = isLt?'rgba(0,0,0,.08)':'rgba(255,255,255,.1)';
        const tc = isLt?'#1c1c1e':'#f0f4ff', t2 = isLt?'rgba(28,28,30,.5)':'rgba(240,244,255,.5)';
        dv.innerHTML = '<div style="background:'+bg+';border:1px solid '+brd+';border-radius:18px;width:min(380px,92vw);overflow:hidden;box-shadow:0 24px 70px rgba(0,0,0,.8);font-family:-apple-system,sans-serif">'
          +'<div style="height:3px;background:linear-gradient(90deg,#f59e0b,#ef4444)"></div>'
          +'<div style="padding:24px">'
          +'<div style="font-size:1rem;font-weight:800;color:'+tc+';margin-bottom:6px;">⚠️ Duplicate Entry Detected</div>'
          +'<div style="font-size:.82rem;color:'+t2+';margin-bottom:14px;">The following entries are already in your submission history with the same hours:</div>'
          +'<div style="background:rgba(245,158,11,.08);border:1px solid rgba(245,158,11,.25);border-radius:10px;padding:12px;margin-bottom:16px;font-size:.82rem;color:#f59e0b;">'
          + dupWarnings.map(w=>'• '+w).join('<br>') +
          '</div>'
          +'<div style="font-size:.8rem;color:'+t2+';margin-bottom:16px;">If the hours are different, you can still submit. Cancel to go back and check.</div>'
          +'<div style="display:flex;gap:10px">'
          +'<button id="dup-cancel" style="flex:1;padding:11px;border-radius:10px;border:1px solid '+brd+';background:rgba(239,68,68,.1);color:#ef4444;font-weight:700;cursor:pointer;font-family:inherit">Cancel</button>'
          +'<button id="dup-proceed" style="flex:1;padding:11px;border-radius:10px;border:none;background:var(--accent);color:#fff;font-weight:700;cursor:pointer;font-family:inherit">Submit Anyway</button>'
          +'</div></div></div>';
        const rr = document.getElementById('dtr-root-outer') || document.body;
        rr.appendChild(dv);
        dv.querySelector('#dup-cancel').onclick  = () => { dv.remove(); resolve(false); };
        dv.querySelector('#dup-proceed').onclick = () => { dv.remove(); resolve(true); };
      });
      if (!proceed) return;
    }

    // ── Show confirmation popup before submitting ───────────────────────
    const confirmed = await showSubmitConfirm(tasks, dateVal, shift);
    if (!confirmed) return; // user cancelled

    // ── Disable button & show progress ──────────────────────────────────
    const sb = q('[data-action="do-submit"]');
    if (sb) { sb.disabled = true; sb.textContent = 'Submitting...'; }

    // ── Save locally first ───────────────────────────────────────────────
    tasks.forEach(t => submissions.push(t));
    const userKey = 'dtr_subs_' + authState.name.trim().toLowerCase().replace(/[^a-z0-9]+/g,'_');
    safeSave(userKey, submissions);
    updateSBStats();

    // ── Post to SharePoint ───────────────────────────────────────────────
    let posted = 0;
    for (const t of tasks) {
      const ok = await postTask(t);
      if (ok) {
        posted++;
        console.log('[WorkPulse] ✅ SP saved:', t.employeeName, t.taskType, t.date);
      } else {
        spQueue.push(t);
        console.warn('[WorkPulse] ⚠ Queued:', t.taskType, t.date);
      }
    }
    safeSave('dtr_spq2', spQueue);

    // ── Restore button ───────────────────────────────────────────────────
    if (sb) { sb.disabled = false; sb.innerHTML = ic.submit + ' Submit Report'; }

    // ── Confirmation popup ──────────────────────────────────────────────
    showSubmitResult(posted, tasks.length, dateVal);

    // ── Reset form ───────────────────────────────────────────────────────
    selectedDate = todayStr();
    renderSubmit();
  }

  function showSubmitResult(posted, total, dateVal) {
    const isLt    = root && root.classList.contains('lt');
    const popBg   = isLt ? '#ffffff'            : '#12151f';
    const rowBg   = isLt ? '#f5f5f7'            : '#1c2030';
    const bord    = isLt ? 'rgba(0,0,0,.08)'    : 'rgba(255,255,255,.1)';
    const textC   = isLt ? '#1c1c1e'            : '#f0f4ff';
    const text2C  = isLt ? 'rgba(28,28,30,.5)'  : 'rgba(240,244,255,.5)';
    const shadow  = isLt ? '0 8px 32px rgba(0,0,0,.15)' : '0 24px 70px rgba(0,0,0,.9)';
    const dateDisp = new Date((dateVal||todayStr())+'T12:00:00').toLocaleDateString('en-IN',{weekday:'long',day:'numeric',month:'long',year:'numeric'});
    const allOk   = posted === total;
    const partial = posted > 0 && !allOk;
    const color   = allOk ? '#22c55e' : partial ? '#f59e0b' : '#ef4444';
    const bg      = allOk ? 'rgba(34,197,94,.08)'   : partial ? 'rgba(245,158,11,.08)'  : 'rgba(239,68,68,.08)';
    const brd     = allOk ? 'rgba(34,197,94,.25)'   : partial ? 'rgba(245,158,11,.25)'  : 'rgba(239,68,68,.25)';
    const icon    = allOk
      ? '<svg width="44" height="44" viewBox="0 0 44 44" fill="none"><circle cx="22" cy="22" r="21" fill="rgba(34,197,94,.12)" stroke="#22c55e" stroke-width="1.5"/><path d="M13 22.5l6.5 6.5 12-13" stroke="#22c55e" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"/></svg>'
      : partial
      ? '<svg width="44" height="44" viewBox="0 0 44 44" fill="none"><circle cx="22" cy="22" r="21" fill="rgba(245,158,11,.12)" stroke="#f59e0b" stroke-width="1.5"/><path d="M22 14v12M22 30v2" stroke="#f59e0b" stroke-width="2.2" stroke-linecap="round"/></svg>'
      : '<svg width="44" height="44" viewBox="0 0 44 44" fill="none"><circle cx="22" cy="22" r="21" fill="rgba(239,68,68,.12)" stroke="#ef4444" stroke-width="1.5"/><path d="M15 15l14 14M29 15L15 29" stroke="#ef4444" stroke-width="2.2" stroke-linecap="round"/></svg>';
    const title   = allOk ? 'Submission Complete' : partial ? 'Partially Submitted' : 'Saved Locally';

    const ov = document.createElement('div');
    ov.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,.55);z-index:2147483646;display:flex;align-items:center;justify-content:center;backdrop-filter:blur(8px)';
    ov.innerHTML =
      '<div style="background:'+popBg+';border:1px solid '+bord+';border-radius:18px;width:min(380px,92vw);overflow:hidden;box-shadow:'+shadow+';font-family:-apple-system,sans-serif">' +
        '<div style="height:3px;background:'+color+'"></div>' +
        '<div style="padding:28px 24px 20px">' +
          '<div style="display:flex;align-items:center;gap:16px;margin-bottom:20px">' +
            icon +
            '<div>' +
              '<div style="font-size:1rem;font-weight:700;color:'+textC+';letter-spacing:-.2px">'+title+'</div>' +
              '<div style="font-size:.78rem;color:'+text2C+';margin-top:3px">'+dateDisp+'</div>' +
            '</div>' +
          '</div>' +
          '<div style="background:'+bg+';border:1px solid '+brd+';border-radius:12px;padding:16px;margin-bottom:18px">' +
            '<div style="display:flex;align-items:baseline;gap:6px;margin-bottom:'+(allOk?'0':'10px')+'">' +
              '<span style="font-size:2rem;font-weight:800;color:'+color+';letter-spacing:-1px">'+posted+'</span>' +
              '<span style="font-size:.95rem;color:'+text2C+'">/ '+total+' tasks saved to SharePoint</span>' +
            '</div>' +
            (!allOk && (total-posted) > 0
              ? '<div style="font-size:.75rem;color:'+text2C+';padding-top:10px;border-top:1px solid '+brd+'">' +
                  (total-posted)+' task(s) queued locally.' +
                  '<br><b style="color:'+textC+'">To sync:</b> Open <a href="'+SP.SITE+'" target="_blank" style="color:#4f9eff;">SharePoint</a>, sign in, then go to <b style="color:'+textC+'">Settings → Retry Queue</b>' +
                '</div>'
              : '') +
          '</div>' +
          '<div style="display:flex;gap:10px">' +
            (!allOk ? '<button id="sr-retry" style="flex:1;padding:12px;border-radius:10px;border:none;background:var(--accent);color:#fff;font-size:.9rem;font-weight:700;cursor:pointer;font-family:inherit">Retry Queue</button>' : '') +
            '<button id="sr-close" style="flex:1;padding:12px;border-radius:10px;border:1px solid '+bord+';background:'+rowBg+';color:'+text2C+';font-size:.9rem;font-weight:600;cursor:pointer;font-family:inherit">Close</button>' +
          '</div>' +
        '</div>' +
      '</div>';
    const rr = document.getElementById('dtr-root-outer') || document.body;
    rr.appendChild(ov);
    ov.querySelector('#sr-close').onclick = () => ov.remove();
    const retryBtn = ov.querySelector('#sr-retry');
    if (retryBtn) retryBtn.onclick = () => { ov.remove(); switchTab('settings'); };
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
      nptHours: Math.round((s.npt||0)*60) + ' m',
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
      {key:'hours',      label:'Mins'},
      {key:'nptHours',   label:'NPT Mins'},
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
      '<input type="text" id="tk-filter" class="dtr-input" placeholder="Filter by task, date, type..." style="max-width:260px;background:var(--bg3);color:var(--text);border:1px solid var(--border)">' +
      '<select id="tk-type-filter" class="dtr-select" style="max-width:160px">' +
      '<option value="">All Work Types</option>' +
      '<option value="Productive">Productive</option><option value="NPT">NPT</option>' +
      '</select>' +
      '<select id="tk-task-filter" class="dtr-select" style="max-width:180px">' +
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
  let attSelectedStatus  = '';
  let attLeaveDay        = '';   // 'full' | 'half'
  let attLeaveHalf       = '';   // 'AM' | 'PM' (only when attLeaveDay==='half')
  const LEAVE_STATUSES   = ['SL','CL','AL']; // statuses that need full/half choice

  // ── MARK ATTENDANCE — weekly grid UI, any date editable ──────────────

  function renderWeeklyCalendar() {
    const el = q('#view-weekly'); if (!el) return;
    if (!attDate) attDate = todayStr();
    const today = todayStr();
    const ws = getWeekStart(attWeekOffset);
    const wLabel = getWeekLabel(attWeekOffset);
    const days = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
    const dates = Array.from({length:7}, (_,i) => {
      const d = new Date(ws); d.setDate(ws.getDate()+i);
      const dk = d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0');
      return {dk, day:days[i], label:d.toLocaleDateString('en-US',{month:'short',day:'numeric'}),
              isToday:dk===today, isWeekend:i===0||i===6};
    });

    el.innerHTML =
      '<div class="ph"><div class="ph-left"><div class="ph-title">Weekly Calendar</div>' +
      '<div class="ph-sub">Click a day · Multi-select for bulk · All attendance edits here</div></div></div>' +

      // Week navigator + controls
      '<div class="wv-controls" style="margin-bottom:14px">' +
      '<button class="wv-nav-btn" data-action="att-week-prev">' + ic.left + '</button>' +
      '<div class="wv-range"><div class="wv-range-title">' + wLabel + '</div>' +
      '<div class="wv-range-sub">' + (attWeekOffset===0?'Current Week':attWeekOffset<0?Math.abs(attWeekOffset)+' week(s) ago':attWeekOffset+' ahead') + '</div></div>' +
      (attWeekOffset!==0?'<button class="wv-today-btn" data-action="att-week-today">This Week</button>':'') +
      '<button class="att-select-toggle' + (attBulkMode?' on':'') + '" data-action="att-toggle-bulk">' + (attBulkMode?'✕ Cancel':'☑ Multi-select') + '</button>' +
      '' +
      '<button class="wv-nav-btn" data-action="att-week-next">' + ic.right + '</button></div>' +

      // Bulk bar rendered at bottom with stable ID

      // Week day cards
      '<div class="att-week-row">' +
      dates.map(d => {
        const st  = (statusCache[d.dk]||{}).status||'';
        const cfg = st ? STATUS_CFG[st] : null;
        const isSel   = !attBulkMode && attDate === d.dk;
        const isMulti = attBulkMode && attMultiSelect.has(d.dk);
        return '<div class="att-day-card' +
          (isSel   ? ' att-day-selected' : '') +
          (isMulti ? ' att-multi'        : '') +
          (d.isToday ? ' att-day-today'  : '') +
          (d.isWeekend ? ' att-day-weekend' : '') + '"' +
          ' data-action="' + (attBulkMode?'att-multi-pick':'att-pick-day') + '" data-val="' + d.dk + '"' +
          (cfg && !isMulti ? ' style="border-color:'+cfg.color+';background:'+cfg.bg+'"':'') + '>' +
          '<div class="att-day-name">' + d.day + '</div>' +
          '<div class="att-day-date">' + d.label + '</div>' +
          (st
            ? '<div class="att-day-dot" style="background:' + (cfg?cfg.color:'var(--text3)') + '"></div>' +
              '<div class="att-day-status" style="color:' + (cfg?cfg.color:'var(--text3)') + ';font-size:.72rem;font-weight:800">' + st + '</div>'
            : '<div class="att-day-dot" style="background:var(--border2)"></div>' +
              '<div class="att-day-status" style="color:var(--text3)">' + (d.isWeekend?'Weekend':'Tap') + '</div>') +
          (isMulti?'<div style="position:absolute;top:6px;right:6px;width:16px;height:16px;border-radius:50%;background:var(--accent2);display:flex;align-items:center;justify-content:center;font-size:9px;color:#fff;font-weight:800">✓</div>':'') +
          (isSel?'<div class="att-day-active-dot"></div>':'') +
          '</div>';
      }).join('') + '</div>' +

      // Bulk action bar (always rendered, shown/hidden via display)
      '<div id="att-bulk-bar" class="att-bulk-bar" style="display:' + (attBulkMode && attMultiSelect.size>0?'flex':'none') + '">' +
      '<span class="bulk-count" style="font-size:.82rem;font-weight:700;color:var(--accent2)">' + attMultiSelect.size + ' day(s) selected — Apply:</span>' +
      '<div style="display:flex;gap:6px;flex-wrap:wrap;flex:1">' +
      STATUSES.map(s => '<button class="att-bulk-btn" data-action="att-bulk-apply" data-val="' + s + '">' + s + '</button>').join('') +
      '</div><button class="btn btn-ghost btn-xs" data-action="att-clear-sel">Clear</button></div>' +

      // Bulk hint (shown in bulk mode with 0 selected)
      '<div id="att-bulk-hint" class="info-banner" style="margin-top:8px;display:' + (attBulkMode && attMultiSelect.size===0?'block':'none') + '">☑️ Tap days above to select, then choose a status to apply to all at once.</div>' +

      // Edit panel (single mode)
      '<div id="att-edit-panel" style="margin-top:8px;display:' + (attBulkMode?'none':'block') + '">' + buildAttEditPanel(attDate) + '</div>';
  }

  function attToggleBulk() {
    attBulkMode = !attBulkMode;
    attMultiSelect.clear();

    // DOM-only update — no full re-render
    const toggleBtn = q('[data-action="att-toggle-bulk"]');
    if (toggleBtn) {
      toggleBtn.textContent = attBulkMode ? '✕ Cancel' : '☑ Multi-select';
      toggleBtn.classList.toggle('on', attBulkMode);
    }

    // Swap all day card click actions
    qa('.att-day-card').forEach(card => {
      card.dataset.action = attBulkMode ? 'att-multi-pick' : 'att-pick-day';
      card.classList.remove('att-multi');
      const chk = card.querySelector('.att-multi-chk');
      if (chk) chk.remove();
    });

    // Show/hide edit panel vs bulk hint
    const editPanel = q('#att-edit-panel');
    const bulkHint  = q('#att-bulk-hint');
    const bulkBar   = q('#att-bulk-bar');
    if (editPanel) editPanel.style.display = attBulkMode ? 'none' : 'block';
    if (bulkHint)  bulkHint.style.display  = attBulkMode ? 'block' : 'none';
    if (bulkBar)   bulkBar.style.display   = 'none';
  }

  function attMultiPick(dk) {
    // Block weekends and mandatory holidays
    const dow = new Date(dk+'T12:00:00').getDay();
    if (dow === 0 || dow === 6) { toast('Saturday & Sunday are mandatory off days', 'err'); return; }
    if (isMandatoryOff(dk)) { toast('🇮🇳 ' + (statusCache[dk]?.name||'National Holiday') + ' — mandatory off', 'err'); return; }
    // Toggle selection in Set
    if (attMultiSelect.has(dk)) {
      attMultiSelect.delete(dk);
    } else {
      attMultiSelect.add(dk);
    }

    // DOM-only: update only the tapped card
    const card = q('.att-day-card[data-val="' + dk + '"]');
    if (card) {
      const selected = attMultiSelect.has(dk);
      card.classList.toggle('att-multi', selected);
      // Add/remove checkmark badge
      let chk = card.querySelector('.att-multi-chk');
      if (selected && !chk) {
        chk = document.createElement('div');
        chk.className = 'att-multi-chk';
        chk.style.cssText = 'position:absolute;top:5px;right:5px;width:16px;height:16px;border-radius:50%;background:var(--accent2);display:flex;align-items:center;justify-content:center;font-size:9px;color:#fff;font-weight:800;pointer-events:none';
        chk.textContent = '✓';
        card.appendChild(chk);
      } else if (!selected && chk) {
        chk.remove();
      }
    }

    // DOM-only: update bulk bar count and visibility
    _updateBulkBar();
  }

  function _updateBulkBar() {
    const count   = attMultiSelect.size;
    const bulkBar = q('#att-bulk-bar');
    const bulkHint= q('#att-bulk-hint');
    if (bulkBar) {
      if (count > 0) {
        bulkBar.style.display = 'flex';
        const countEl = bulkBar.querySelector('.bulk-count');
        if (countEl) countEl.textContent = count + ' day(s) selected — Apply:';
      } else {
        bulkBar.style.display = 'none';
      }
    }
    if (bulkHint) bulkHint.style.display = count > 0 ? 'none' : 'block';
  }

  async function attBulkApply(status) {
    if (!status || attMultiSelect.size === 0) return;
    // Filter out weekends — cannot mark Sat/Sun
    const dks = [...attMultiSelect].filter(dk => { const d = new Date(dk+'T12:00:00').getDay(); return d !== 0 && d !== 6; });
    if (dks.length === 0) { toast('No working days selected — Sat/Sun cannot be marked', 'err'); return; }
    const cfg = STATUS_CFG[status] || {};

    // 1. Instant DOM update — all cards immediately show new status
    dks.forEach(dk => {
      const card = q('.att-day-card[data-val="' + dk + '"]');
      if (card) {
        card.classList.remove('att-multi');
        card.style.borderColor = cfg.color || '';
        card.style.background  = cfg.bg    || '';
        const dot = card.querySelector('.att-day-dot');
        const lbl = card.querySelector('.att-day-status');
        const chk = card.querySelector('.att-multi-chk');
        if (dot) dot.style.background = cfg.color || 'var(--border2)';
        if (lbl) { lbl.textContent = status; lbl.style.color = cfg.color || 'var(--text3)'; }
        if (chk) chk.remove();
      }
      // Update local cache immediately
      statusCache[dk] = { status, shift: (statusCache[dk]||{}).shift||'', date:dk, name:authState.name, updatedAt:nowUTC() };
      teamStatusCache[authState.name+'::'+dk] = { status, shift: (statusCache[dk]||{}).shift||'' };
    });
    safeSaveObj('dtr_status2', statusCache);
    safeSaveObj('dtr_teamcache', teamStatusCache);

    // Reset bulk mode immediately (instant feel)
    attMultiSelect.clear();
    attBulkMode = false;
    attLeaveDay = ''; attLeaveHalf = '';
    const toggleBtn = q('[data-action="att-toggle-bulk"]');
    if (toggleBtn) { toggleBtn.textContent = '☑ Multi-select'; toggleBtn.classList.remove('on'); }
    qa('.att-day-card').forEach(c => { c.dataset.action = 'att-pick-day'; });
    const bulkBar  = q('#att-bulk-bar');
    const editPanel= q('#att-edit-panel');
    const bulkHint = q('#att-bulk-hint');
    if (bulkBar)   bulkBar.style.display   = 'none';
    if (bulkHint)  bulkHint.style.display  = 'none';
    if (editPanel) editPanel.style.display = 'block';

    toast('✅ Saving ' + dks.length + ' days as ' + status + '...', 'ok');

    // 2. Post to SP in parallel (not sequential)
    const results = await Promise.all(
      dks.map(dk => postAttendance({ status, shift:(statusCache[dk]||{}).shift||'', date:dk, name:authState.name, updatedAt:nowUTC() }))
    );
    const saved = results.filter(Boolean).length;
    if (saved === dks.length) {
      toast('✅ ' + status + ' applied to ' + saved + ' day(s) — saved to SharePoint', 'ok');
    } else {
      toast('⚠️ ' + saved + '/' + dks.length + ' saved to SP — rest queued locally', 'info');
    }
  }


  function buildAttEditPanel(dk) {
    const existing = statusCache[dk] || {};
    // Restore leave sub-type from saved entry
    if (existing.leaveDay)  attLeaveDay  = existing.leaveDay;
    if (existing.leaveHalf) attLeaveHalf = existing.leaveHalf;
    const selStatus = attSelectedStatus || existing.status || '';
    // Use local date parts to avoid UTC offset issues
    const now   = new Date();
    const today = now.getFullYear()+'-'+String(now.getMonth()+1).padStart(2,'0')+'-'+String(now.getDate()).padStart(2,'0');
    const yest  = new Date(now); yest.setDate(now.getDate()-1);
    const yesterdayLocal = yest.getFullYear()+'-'+String(yest.getMonth()+1).padStart(2,'0')+'-'+String(yest.getDate()).padStart(2,'0');
    const isFuture = dk > today;
    const isToday  = dk === today;
    const isYest   = dk === yesterdayLocal;
    // Always show the full date — no ambiguity
    const tag = isToday ? ' (Today)' : isYest ? ' (Yesterday)' : '';
    const dateLabel = formatDate(dk) + tag;
    return '<div class="att-edit-card">' +
      '<div class="att-edit-header">' +
      '<div><div class="att-edit-date">' + dateLabel + '</div>' +
      '<div class="att-edit-datesub">' + formatDay(dk) + '</div></div>' +
      (existing.status ? '<span class="att-saved-pill" style="background:' + (STATUS_CFG[existing.status]||{}).bg + ';color:' + (STATUS_CFG[existing.status]||{}).color + '">✓ Saved: ' + existing.status + '</span>' : '') +
      '</div>' +
      (isFuture ? '<div class="info-banner" style="margin-bottom:12px">📅 Future date — you can pre-mark it.</div>' : '') +
      '<div class="status-grid" id="att-status-grid">' +
      STATUSES.map(s => {
        const cfg = STATUS_CFG[s], sel = (selStatus === s);
        return '<div class="status-tile' + (sel ? ' selected' : '') + '" data-action="att-status" data-val="' + s + '"' +
          ' style="border-color:' + (sel ? cfg.color : 'var(--border)') + ';background:' + (sel ? cfg.bg : 'var(--bg2)') + '">' +
          '<div class="sel-check" style="background:' + cfg.color + ';color:#fff">' + ic.check + '</div>' +
          '<div class="status-tile-label">' + s + '</div>' +
          '<div class="status-tile-sub">' + cfg.label + '</div></div>';
      }).join('') + '</div>' +
      // Leave sub-type panel — shown inline for SL/CL/AL
      '<div id="att-leave-panel" style="display:' + (LEAVE_STATUSES.includes(selStatus)?'block':'none') + ';margin-bottom:2px">' +
      (LEAVE_STATUSES.includes(selStatus) ? _buildLeaveDayPanel(selStatus, STATUS_CFG[selStatus]||{}) : '') + '</div>' +
      '<div class="dtr-field" style="margin-bottom:14px">' +
      '<label class="dtr-label">Shift Timing</label>' +
      '<select class="dtr-select" id="att-shift" style="max-width:220px" data-action="att-shift-change">' +
      '<option value="">— Select shift —</option>' +
      (function(sv){ return '<option value="8-5"'+(sv==='8-5'?' selected':'')+'>8 AM – 5 PM</option>'+'<option value="9-6"'+(sv==='9-6'?' selected':'')+'>9 AM – 6 PM</option>'+'<option value="10-7"'+(sv==='10-7'?' selected':'')+'>10 AM – 7 PM</option>'+'<option value="11-8"'+(sv==='11-8'?' selected':'')+'>11 AM – 8 PM</option>';})(existing.shift || authState.shift || loadShift(authState.name)) + '</select></div>' +
      '<button class="att-save-btn" id="att-save-btn" data-action="att-save"' + (!selStatus ? ' disabled' : '') + '>' +
      ic.check + ' Save — ' + dateLabel + '</button>' +
      '</div>';
  }

  function attSelectStatus(status) {
    attSelectedStatus = status;
    attLeaveDay  = '';
    attLeaveHalf = '';
    const cfg = STATUS_CFG[status] || {};

    // Update tile highlights — pure DOM, no rebuild
    qa('#att-status-grid .status-tile').forEach(tile => {
      const s    = tile.dataset.val;
      const tcfg = STATUS_CFG[s] || {};
      const isSel = (s === status);
      tile.classList.toggle('selected', isSel);
      tile.style.borderColor = isSel ? tcfg.color : 'var(--border)';
      tile.style.background  = isSel ? tcfg.bg    : 'var(--bg2)';
    });

    // Show leave sub-type panel for SL/CL/AL — inline below status grid
    const leavePanel = q('#att-leave-panel');
    if (leavePanel) {
      if (LEAVE_STATUSES.includes(status)) {
        leavePanel.style.display = 'block';
        leavePanel.innerHTML = _buildLeaveDayPanel(status, cfg);
      } else {
        leavePanel.style.display = 'none';
        leavePanel.innerHTML = '';
      }
    }

    // For non-leave statuses: enable save immediately
    // For leave statuses: save button stays disabled until full/half chosen
    const saveBtn = q('#att-save-btn');
    if (saveBtn) {
      const needsChoice = LEAVE_STATUSES.includes(status);
      saveBtn.disabled = needsChoice;
      if (needsChoice) saveBtn.setAttribute('disabled','');
      else saveBtn.removeAttribute('disabled');
    }

    // Live update day card
    qa('.att-day-card').forEach(card => {
      if (card.dataset.val === attDate) {
        const dotEl = card.querySelector('.att-day-dot');
        const stEl  = card.querySelector('.att-day-status');
        if (dotEl) dotEl.style.background = cfg.color || 'var(--border2)';
        if (stEl)  { stEl.textContent = status; stEl.style.color = cfg.color || 'var(--text3)'; }
        if (cfg.color) { card.style.borderColor = cfg.color; card.style.background = cfg.bg; }
      }
    });
  }

  // Build the inline full/half day selection panel (no popup)
  function _buildLeaveDayPanel(status, cfg) {
    const color = cfg.color || 'var(--accent)';
    const bg    = cfg.bg    || 'var(--bg3)';
    return (
      '<div style="margin-top:10px;padding:12px 14px;border-radius:10px;border:1px solid ' + color + '30;background:' + bg + ';animation:fadeIn .15s">' +
      '<div style="font-size:.8rem;font-weight:700;color:' + color + ';margin-bottom:10px;letter-spacing:.3px">' + status + ' — Select Duration</div>' +

      // Full / Half row
      '<div style="display:flex;gap:8px;margin-bottom:8px">' +

      // Full Day button
      '<button id="ldBtn-full" data-action="att-leave-day" data-val="full" style="flex:1;padding:9px 6px;border-radius:8px;border:2px solid ' +
        (attLeaveDay==='full' ? color : 'var(--border)') + ';background:' +
        (attLeaveDay==='full' ? bg : 'var(--bg2)') + ';color:' +
        (attLeaveDay==='full' ? color : 'var(--text2)') + ';font-size:.82rem;font-weight:700;cursor:pointer;font-family:inherit;transition:all .1s">' +
        '📅 Full Day</button>' +

      // Half Day button
      '<button id="ldBtn-half" data-action="att-leave-day" data-val="half" style="flex:1;padding:9px 6px;border-radius:8px;border:2px solid ' +
        (attLeaveDay==='half' ? color : 'var(--border)') + ';background:' +
        (attLeaveDay==='half' ? bg : 'var(--bg2)') + ';color:' +
        (attLeaveDay==='half' ? color : 'var(--text2)') + ';font-size:.82rem;font-weight:700;cursor:pointer;font-family:inherit;transition:all .1s">' +
        '⏰ Half Day</button>' +

      '</div>' +

      // AM / PM row — only shown when half is selected
      (attLeaveDay === 'half' ?
        '<div style="display:flex;gap:8px;margin-top:4px">' +
        '<button id="ldBtn-AM" data-action="att-leave-half" data-val="AM" style="flex:1;padding:8px 6px;border-radius:8px;border:2px solid ' +
          (attLeaveHalf==='AM' ? color : 'var(--border)') + ';background:' +
          (attLeaveHalf==='AM' ? bg : 'var(--bg2)') + ';color:' +
          (attLeaveHalf==='AM' ? color : 'var(--text3)') + ';font-size:.79rem;font-weight:600;cursor:pointer;font-family:inherit">' +
          '🌅 First Half (Morning)</button>' +
        '<button id="ldBtn-PM" data-action="att-leave-half" data-val="PM" style="flex:1;padding:8px 6px;border-radius:8px;border:2px solid ' +
          (attLeaveHalf==='PM' ? color : 'var(--border)') + ';background:' +
          (attLeaveHalf==='PM' ? bg : 'var(--bg2)') + ';color:' +
          (attLeaveHalf==='PM' ? color : 'var(--text3)') + ';font-size:.79rem;font-weight:600;cursor:pointer;font-family:inherit">' +
          '🌆 Second Half (Afternoon)</button>' +
        '</div>' : '') +

      '</div>'
    );
  }

  // Handle Full/Half Day selection
  function attSelectLeaveDay(val) {
    attLeaveDay  = val;
    attLeaveHalf = '';
    const cfg = STATUS_CFG[attSelectedStatus] || {};
    // Rebuild just the leave panel
    const leavePanel = q('#att-leave-panel');
    if (leavePanel) leavePanel.innerHTML = _buildLeaveDayPanel(attSelectedStatus, cfg);
    // Enable/disable save: full day → can save; half → still need AM/PM
    const saveBtn = q('#att-save-btn');
    if (saveBtn) {
      const needsHalf = val === 'half';
      saveBtn.disabled = needsHalf;
      if (needsHalf) saveBtn.setAttribute('disabled','');
      else           saveBtn.removeAttribute('disabled');
    }
  }

  // Handle AM/PM selection
  function attSelectLeaveHalf(val) {
    attLeaveHalf = val;
    const cfg = STATUS_CFG[attSelectedStatus] || {};
    const leavePanel = q('#att-leave-panel');
    if (leavePanel) leavePanel.innerHTML = _buildLeaveDayPanel(attSelectedStatus, cfg);
    // Now can save
    const saveBtn = q('#att-save-btn');
    if (saveBtn) { saveBtn.disabled = false; saveBtn.removeAttribute('disabled'); }
  }

  // Compute final status label including sub-type
  // e.g. "SL (Full Day)" or "CL (Half - AM)"
  function getEffectiveStatus() {
    if (!LEAVE_STATUSES.includes(attSelectedStatus)) return attSelectedStatus;
    if (attLeaveDay === 'full')  return attSelectedStatus + ' - Full Day';
    if (attLeaveDay === 'half' && attLeaveHalf) return attSelectedStatus + ' - Half ' + attLeaveHalf;
    return attSelectedStatus;
  }

  async function attSave() {
    if (!attSelectedStatus) { toast('Select a status first', 'err'); return; }
    const btn = q('[data-action="att-save"]');
    if (btn) { btn.disabled = true; btn.textContent = 'Saving...'; }
    const shiftEl = q('#att-shift');
    const shift   = (shiftEl?.value) || authState.shift || loadShift(authState.name) || '';
    if (shift && authState.name) { authState.shift = shift; saveShift(authState.name, shift); }
    const effectiveStatus = getEffectiveStatus();
    const entry = { status: effectiveStatus, leaveDay: attLeaveDay, leaveHalf: attLeaveHalf, shift, date: attDate, name: authState.name, updatedAt: nowUTC() };
    statusCache[attDate] = entry;
    safeSaveObj('dtr_status2', statusCache);
    teamStatusCache[authState.name + '::' + attDate] = { status: effectiveStatus, shift };
    safeSaveObj('dtr_teamcache', teamStatusCache);
    const sent = await postAttendance(entry);
    if (!sent) toast('Saved locally (SP sync pending)', 'info');
    else toast('✅ Saved — ' + effectiveStatus, 'ok');
    attLeaveDay = ''; attLeaveHalf = '';
    if (btn) { btn.disabled = false; btn.innerHTML = ic.check + ' Save — ' + (attDate === todayStr() ? 'Today' : attDate === prevWorkingDayStr() ? 'Yesterday' : formatDay(attDate)); }
    // Update day card to show saved status
    qa('.att-day-card').forEach(card => {
      if (card.dataset.val === attDate) {
        const cfg = STATUS_CFG[attSelectedStatus] || {};
        const dotEl = card.querySelector('.att-day-dot');
        const stEl  = card.querySelector('.att-day-status');
        if (dotEl) dotEl.style.background = cfg.color || 'var(--border2)';
        if (stEl)  { stEl.textContent = attSelectedStatus; stEl.style.color = cfg.color || 'var(--text3)'; }
        card.style.borderColor = cfg.color || 'var(--border)';
        card.style.background  = cfg.bg    || 'var(--bg2)';
      }
    });
    // Re-render panel to show saved badge
    const p = q('#att-edit-panel');
    if (p) p.innerHTML = buildAttEditPanel(attDate);
    if (currentView === 'calendar') renderMyCalendar();
  }

  function attNavDate(delta) {
    const d = new Date(attDate + 'T12:00:00');
    d.setDate(d.getDate() + delta);
    attDate = d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0');
    attSelectedStatus = (statusCache[attDate] || {}).status || '';
    const todayD = new Date(todayStr() + 'T12:00:00');
    const selD   = new Date(attDate + 'T12:00:00');
    const todaySun = new Date(todayD); todaySun.setDate(todaySun.getDate() - todaySun.getDay());
    const selSun   = new Date(selD);   selSun.setDate(selSun.getDate() - selSun.getDay());
    attWeekOffset = Math.round((selSun - todaySun) / (7 * 24 * 60 * 60 * 1000));
    renderWeeklyCalendar();
  }

  function attNavToday() {
    attDate = todayStr();
    attWeekOffset = 0;
    attSelectedStatus = (statusCache[attDate] || {}).status || '';
    renderWeeklyCalendar();
  }

  // ── MY CALENDAR ────────────────────────────────────────────────────────

  // ═══════════════════════════════════════════════════════════════════════
  // MY ATTENDANCE CALENDAR — merged attendance + calendar
  // Sat/Sun = holidays, click any weekday to mark, multi-select supported
  // ═══════════════════════════════════════════════════════════════════════
  // MY CALENDAR — view-only attendance overview (no editing)
  function renderMyCalendar() {
    const el = q('#view-calendar'); if (!el) return;
    const now = new Date();
    const targetDate = new Date(now.getFullYear(), now.getMonth() + calMonthOffset, 1);
    const year  = targetDate.getFullYear();
    const month = targetDate.getMonth();
    const monthName = targetDate.toLocaleDateString('en-US', {month:'long', year:'numeric'});
    const today = todayStr();
    const firstDay  = new Date(year, month, 1).getDay();
    const totalDays = new Date(year, month + 1, 0).getDate();
    const monthKey  = year + '-' + String(month+1).padStart(2,'0');
    const monthEntries = Object.entries(statusCache).filter(([dk]) => dk.startsWith(monthKey));
    const countSt = s => monthEntries.filter(([,v]) => v.status===s).length;
    let workDays = 0;
    for (let d=1; d<=totalDays; d++) { const dow=new Date(year,month,d).getDay(); if(dow!==0&&dow!==6) workDays++; }
    const markedDays = monthEntries.filter(([dk])=>{const dow=new Date(dk+'T12:00:00').getDay();return dow!==0&&dow!==6;}).length;
    const STATUS_COLOR = {WFO:'#3fb950',WFH:'#22d3ee',SL:'#f85149',CL:'#d29922',AL:'#a371f7','Optional Off':'#6b7280'};
    const STATUS_BG    = {WFO:'rgba(63,185,80,.13)',WFH:'rgba(6,182,212,.13)',SL:'rgba(248,81,73,.13)',CL:'rgba(210,153,34,.13)',AL:'rgba(163,113,247,.13)','Optional Off':'rgba(107,114,128,.13)'};
    const DAY_HDRS = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];

    // Build cells — no click handlers (view-only)
    let cells = '';
    for (let i=0; i<firstDay; i++) cells += '<div class="cal-cell cal-empty"></div>';
    for (let d=1; d<=totalDays; d++) {
      const dk      = year+'-'+String(month+1).padStart(2,'0')+'-'+String(d).padStart(2,'0');
      const dow     = new Date(year,month,d).getDay();
      const isWknd  = dow===0||dow===6;
      const isToday = dk===today;
      const isFuture= dk>today;
      const st      = (statusCache[dk]||{}).status||'';
      const color   = st ? STATUS_COLOR[st] : '';
      const bg      = st ? STATUS_BG[st] : '';
      let cellStyle = 'cursor:default;';
      if (isWknd) cellStyle += 'background:rgba(255,255,255,.012);';
      else if (st) cellStyle += 'background:'+bg+';border-color:'+color+';';
      cells += '<div class="cal-cell'+(isWknd?' cal-weekend':'')+(isToday?' cal-today':'')+(isFuture&&!isWknd?' cal-future':'')+'" style="'+cellStyle+'">' +
        '<div class="cal-day-num'+(isToday?' cal-today-num':'')+'">'+d+'</div>' +
        (isWknd ? '<div style="font-size:.6rem;color:var(--text3);margin-top:2px">—</div>' :
          st ? '<div class="cal-day-badge" style="background:'+color+';color:#fff;padding:2px 6px;border-radius:4px;font-size:.68rem;font-weight:800;margin-top:4px">'+st+'</div>'
             : '') +
        '</div>';
    }

    // Summary pills
    const sumItems = [['WFO',countSt('WFO'),'#3fb950'],['WFH',countSt('WFH'),'#22d3ee'],['SL',countSt('SL'),'#f85149'],['CL',countSt('CL'),'#d29922'],['AL',countSt('AL'),'#a371f7'],['Off',countSt('Optional Off'),'#6b7280']].filter(([,n])=>n>0);

    el.innerHTML =
      '<div class="ph"><div class="ph-left"><div class="ph-title">My Calendar</div>' +
      '<div class="ph-sub">'+markedDays+' of '+workDays+' working days marked · View-only — use Weekly Calendar to edit</div></div>' +
      '<div class="ph-actions"><button class="btn btn-ghost btn-sm" data-action="cal-sync">'+ic.sync+' Sync</button></div></div>' +
      '<div class="info-banner" style="margin-bottom:14px">📋 This is a read-only view of your attendance. To mark or edit attendance, go to the <strong>Weekly Calendar</strong> tab.</div>' +
      '<div class="wv-controls" style="margin-bottom:14px">' +
      '<button class="wv-nav-btn" data-action="cal-prev">'+ic.left+'</button>' +
      '<div class="wv-range"><div class="wv-range-title">'+monthName+'</div>' +
      '<div class="wv-range-sub">'+(calMonthOffset===0?'Current Month':Math.abs(calMonthOffset)+' month(s) '+(calMonthOffset<0?'ago':'ahead'))+'</div></div>' +
      (calMonthOffset!==0?'<button class="wv-today-btn" data-action="cal-today">This Month</button>':'') +
      '<button class="wv-nav-btn" data-action="cal-next">'+ic.right+'</button></div>' +
      (sumItems.length ? '<div style="display:flex;gap:6px;flex-wrap:wrap;margin-bottom:12px">'+
        sumItems.map(([l,n,c])=>'<span style="padding:4px 10px;border-radius:20px;background:'+c+'18;border:1px solid '+c+'40;color:'+c+';font-size:.75rem;font-weight:700">'+l+' '+n+'</span>').join('')+'</div>' : '') +
      '<div class="cal-grid-wrap">' +
      '<div class="cal-header">'+DAY_HDRS.map(h=>'<div class="cal-hcell">'+h+'</div>').join('')+'</div>' +
      '<div class="cal-grid">'+cells+'</div></div>';
  }


  function renderMissedNPT() {
    const el = q('#view-missednpt'); if (!el) return;
    const myNPT     = nptCache.filter(n => n.name === authState.name);
    const totalMins = myNPT.reduce((a,n)=>a+(n.minutes||0),0);
    const h = Math.floor(totalMins/60), m = totalMins%60;
    const now = new Date();
    const thisMonth = now.getFullYear()+'-'+String(now.getMonth()+1).padStart(2,'0');
    const monthMins = myNPT.filter(n=>n.date&&n.date.startsWith(thisMonth)).reduce((a,n)=>a+(n.minutes||0),0);

    el.innerHTML =
      '<div class="ph"><div class="ph-left"><div class="ph-title">Missed NPT Log</div>' +
      '<div class="ph-sub">Log non-productive time entries · Maps to SharePoint \'NPT log\' list</div></div></div>' +

      // Summary cards
      '<div class="stats-grid sg3" style="margin-bottom:16px">' +
      '<div class="stat-card ab"><div class="lbl">Total Entries</div><div class="val">' + myNPT.length + '</div></div>' +
      '<div class="stat-card amb"><div class="lbl">Total Time</div><div class="val">' + (h>0?h+'h ':'')+m+'m</div></div>' +
      '<div class="stat-card pb"><div class="lbl">This Month</div><div class="val">' + Math.floor(monthMins/60)+'h '+monthMins%60+'m</div></div>' +
      '</div>' +

      // Log form
      '<div class="card" style="margin-bottom:16px">' +
      '<div class="card-title" style="margin-bottom:14px">Log NPT Entry</div>' +

      // NPT Type pills
      '<div class="dtr-label" style="margin-bottom:8px">NPT Type <span class="req-star">*</span></div>' +
      '<div class="npt-types" style="margin-bottom:14px">' +
      NPT_TYPES.map(t =>
        '<button class="npt-type-btn' + (nptActiveType===t?' active':'') + '" data-action="npt-type" data-val="' + t + '">' + t + '</button>'
      ).join('') +
      '</div>' +

      // Date + Duration row
      '<div class="g2" style="margin-bottom:12px">' +
      '<div class="dtr-field"><label class="dtr-label">NPT Date <span class="req-star">*</span></label>' +
      '<input type="date" class="dtr-input" id="npt-date" value="' + todayStr() + '"></div>' +
      '<div class="dtr-field"><label class="dtr-label">Duration (minutes) <span class="req-star">*</span></label>' +
      '<input type="number" class="dtr-input" id="npt-mins" min="1" max="480" placeholder="e.g. 30" style="font-family:var(--mono)"></div>' +
      '</div>' +

      // Employee name (read-only, auto-filled)
      '<div class="dtr-field" style="margin-bottom:12px">' +
      '<label class="dtr-label">Employee Name</label>' +
      '<input class="dtr-input" value="' + authState.name + '" readonly style="cursor:default;color:var(--text2)"></div>' +

      // Description
      '<div class="dtr-field" style="margin-bottom:14px">' +
      '<label class="dtr-label">Description <span class="req-star">*</span></label>' +
      '<textarea class="dtr-input" id="npt-desc" rows="3" placeholder="Describe the reason for non-productive time..." style="resize:vertical;min-height:64px"></textarea></div>' +

      // Submit button
      '<button class="btn btn-primary btn-full" style="padding:11px" data-action="npt-log">' +
      ic.add + ' Submit NPT Entry</button>' +
      '</div>' +

      // History table
      '<div class="card"><div class="card-title" style="display:flex;justify-content:space-between;align-items:center">' +
      '<span>My NPT History</span><span style="font-size:.78rem;color:var(--text3);font-weight:500">' + myNPT.length + ' entries</span></div>' +
      (myNPT.length ?
        '<div class="tbl-wrap"><table class="dtr-table">' +
        '<thead><tr>' +
        '<th>Date</th><th>Type</th><th>Duration</th><th>Description</th>' +
        '<th style="width:40px;text-align:center">SP</th><th style="width:40px"></th>' +
        '</tr></thead><tbody>' +
        myNPT.slice().reverse().map(n => {
          const ri   = nptCache.indexOf(n);
          const mins = n.minutes || 0;
          const dh   = Math.floor(mins/60), dm = mins%60;
          const spOk = n.savedToSP === true;
          return '<tr>' +
            '<td class="mono" style="font-size:.84rem">' + n.date + '</td>' +
            '<td><span class="badge bb" style="font-size:.75rem">' + (n.type||'Other') + '</span></td>' +
            '<td class="mono">' + (dh>0?dh+'h ':'')+dm+'m</td>' +
            '<td style="max-width:220px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:.85rem" title="' + (n.desc||'').replace(/"/g,'&quot;') + '">' + (n.desc||'—') + '</td>' +
            '<td style="text-align:center;font-size:.8rem">' + (spOk ? '✅' : '<span style="color:var(--text3)" title="Saved locally — will sync when SP is available">⏳</span>') + '</td>' +
            '<td><button class="btn btn-danger btn-xs" data-action="npt-del" data-idx="' + ri + '">' + ic.trash + '</button></td>' +
            '</tr>';
        }).join('') +
        '</tbody></table></div>' :
        '<div class="empty"><p>No NPT entries yet. Use the form above to log your first entry.</p></div>') +
      '</div>';
  }

  function nptSelectType(type) { nptActiveType = type; qa('.npt-type-btn').forEach(b => b.classList.toggle('active', b.dataset.val === type)); }

  async function nptLogEntry() {
    const date = q('#npt-date')?.value || todayStr();
    const mins  = parseInt(q('#npt-mins')?.value || '0');
    const desc  = (q('#npt-desc')?.value || '').trim();
    const type  = nptActiveType;

    // Validation
    if (!type)       { toast('Select an NPT type first', 'err'); return; }
    if (!mins||mins<1){ toast('Enter duration in minutes', 'err'); return; }
    if (!desc)        { toast('Add a description', 'err'); return; }

    // Disable button while submitting
    const btn = q('[data-action="npt-log"]');
    if (btn) { btn.disabled = true; btn.textContent = 'Submitting...'; }

    const entry = {
      name:      authState.name,
      date,
      type,
      minutes:   mins,
      desc,
      loggedAt:  nowUTC(),
      savedToSP: false
    };

    // Save locally first
    nptCache.push(entry);
    safeSave('dtr_npt2', nptCache);

    // Post to SharePoint
    const sent = await postNPT(entry);
    if (sent) {
      entry.savedToSP = true;
      safeSave('dtr_npt2', nptCache);
      toast('✅ NPT logged — ' + type + ' (' + mins + 'm) saved to SharePoint', 'ok');
    } else {
      toast('⚠️ NPT saved locally — queued for sync. Go to Settings → Retry Queue.', 'info');
    }

    if (btn) { btn.disabled = false; btn.innerHTML = ic.add + ' Submit NPT Entry'; }
    nptActiveType = '';
    renderMissedNPT();
  }

  function nptDel(idx) { if (!confirm('Delete this NPT entry?')) return; nptCache.splice(idx,1); safeSave('dtr_npt2',nptCache); renderMissedNPT(); }

  // SETTINGS
  function renderSettings() {
    const el = q('#view-settings'); if (!el) return;
    const qLen = (Array.isArray(spQueue)?spQueue:[]).length;
    el.innerHTML =
      '<div class="ph"><div class="ph-left"><div class="ph-title">Settings</div></div></div>' +
      '<div class="card"><div class="card-title">' + ic.sync + ' SharePoint — Single Site</div>' +
      '<div class="info-banner" style="margin-bottom:12px">' +
'<div><strong>Your data is private.</strong> You only see your own submissions, attendance and NPT. ' +
'Admins can see the full team. All data is stored in: <code>' + SP.SITE + '</code></div>' +
      '<div class="warn-banner" style="margin-top:10px">⚠️ <strong>Associates must have Contribute permission on the SharePoint site</strong> for their entries to sync directly. Ask your admin to grant Contribute access to all associates at: ' + SP.SITE + ' → Site Settings → Site Permissions → Grant Permissions.</div></div>' +
      '<div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:10px">' +
      '<button class="btn btn-ghost btn-sm" data-action="test-sp">Test Connection</button>' +
      '<button class="btn btn-ghost btn-sm" data-action="test-npt" style="border-color:rgba(163,113,247,.3);color:var(--accent2)">🧪 Test NPT Insert</button>' +
'<button class="btn btn-ghost btn-sm" data-action="sync-my-data">' + ic.sync + ' Sync My Data</button>' +
      (qLen>0?'<button class="btn btn-ghost btn-sm" data-action="flush-queue">🔄 Retry '+qLen+' Queued</button>':'') + '</div>' +
      '<p style="font-size:.8rem;color:var(--text3)">' + (qLen===0?'✅ All synced':qLen+' pending sync') + '</p>' +
      '<div style="margin-top:14px"><div style="font-size:.78rem;color:var(--text3)"><strong>SharePoint Lists Required:</strong></div>' +
      '<div style="margin-top:6px;font-size:.78rem;color:var(--text3);line-height:2">' +
      '📋 Task List: <code style="color:var(--accent)">' + SP.TASK_LIST + '</code> (TaskReport)<br>' +
      '📅 Attendance List: <code style="color:var(--accent)">' + SP.STATUS_LIST + '</code> (AttendanceLog — columns: EmployeeName, StatusDate, WorkStatus, Shift, UpdatedAt)<br>' +
      '⏱ NPT List: <code style="color:var(--accent)">' + SP.NPT_LIST + '</code> (columns: Title, EmployeeName, NPTDate, NPTType, DurationMins, Description, LoggedAt)</div></div></div>' +
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

  // Digest cache — valid for 25 minutes (SP digest expires at 30m)
  let _digestToken   = '';
  let _digestExpiry  = 0;

  async function getDigest() {
    // Return cached token if still valid
    if (_digestToken && Date.now() < _digestExpiry) {
      console.log('[WorkPulse] Using cached digest (expires in', Math.round((_digestExpiry-Date.now())/1000), 's)');
      return _digestToken;
    }
    return new Promise(resolve => {
      GM_xmlhttpRequest({
        method: 'POST',
        url: SP.SITE + '/_api/contextinfo',
        headers: {
          'Accept':       'application/json;odata=verbose',
          'Content-Type': 'application/json;odata=verbose'
        },
        withCredentials: true,
        onload: res => {
          console.log('[WorkPulse] getDigest response status:', res.status);
          if (res.status === 401 || res.status === 403) {
            toast('❌ SP login required — open ' + SP.SITE + ' in a tab, log in, then retry', 'err');
            console.error('[WorkPulse] getDigest auth error:', res.status, res.responseText.slice(0,200));
            resolve(null); return;
          }
          if (res.status === 0) {
            toast('❌ SP unreachable (status 0) — ensure ' + SP.SITE + ' is open in another tab', 'err');
            resolve(null); return;
          }
          try {
            const parsed = JSON.parse(res.responseText);
            const info   = parsed.d && parsed.d.GetContextWebInformation;
            if (!info) {
              console.error('[WorkPulse] getDigest: no GetContextWebInformation in response:', res.responseText.slice(0,300));
              toast('❌ SP digest response invalid (' + res.status + ') — see F12 console', 'err');
              resolve(null); return;
            }
            const token  = info.FormDigestValue;
            const webUrl = (info.WebFullUrl || '').toLowerCase();
            console.log('[WorkPulse] ✅ SP digest OK | scope:', webUrl, '| length:', token.length);
            // Cache for 25 minutes
            _digestToken  = token;
            _digestExpiry = Date.now() + 25 * 60 * 1000;
            resolve(token);
          } catch(ex) {
            console.error('[WorkPulse] getDigest parse error:', ex.message, '| status:', res.status, '| body:', res.responseText.slice(0,300));
            toast('❌ SP response parse error (' + res.status + '): ' + ex.message, 'err');
            resolve(null);
          }
        },
        onerror: (err) => {
          console.error('[WorkPulse] getDigest network error:', err);
          toast('❌ SP unreachable — open ' + SP.SITE + ' in a tab first, then retry', 'err');
          resolve(null);
        }
      });
    });
  }

  // Clear digest cache (called after scope changes or logout)
  function clearDigestCache() { _digestToken = ''; _digestExpiry = 0; }

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
            let msg = '';
            try {
              const j = JSON.parse(res.responseText);
              msg = (j.error && j.error.message) ? (j.error.message.value || JSON.stringify(j.error.message)) : res.responseText.slice(0,300);
            } catch { msg = res.responseText.slice(0, 300); }
            console.error('[WorkPulse] spInsert FAILED [' + listName + '] status:', res.status, '| error:', msg);

            if (res.status === 403) {
              // Permission denied — show clear one-time guidance
              const _403key = 'dtr_403_warned_' + (authState.name||'').toLowerCase();
              if (!GM_getValue(_403key, false)) {
                GM_setValue(_403key, true);
                toast('🔒 SharePoint access denied for ' + (authState.name||'this account') + '. Ask your admin to grant Contribute access at: ' + SP.SITE + ' → Site Permissions. Your data is saved locally — go to Settings → Retry Queue after access is granted.', 'err', 8000);
              } else {
                toast('🔒 SP access denied — saved locally. Settings → Retry Queue to sync when permissions are fixed.', 'info');
              }
            } else if (res.status === 401) {
              toast('🔐 SP session expired — open ' + SP.SITE + ' in a tab and sign in, then retry.', 'err');
            } else {
              toast('⚠️ SP save failed (' + res.status + ') — saved locally. Settings → Retry Queue.', 'info');
              console.error('[WorkPulse] spInsert detail:', msg.slice(0, 400));
            }
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

  // Dynamic list type cache — fetched once per session
  async function getListType(listName, key) {
    if (_listTypes[key]) return _listTypes[key];
    return new Promise(resolve => {
      GM_xmlhttpRequest({
        method: 'GET',
        url: SP.SITE + "/_api/web/lists/GetByTitle('" + listName + "')?$select=ListItemEntityTypeFullName",
        headers: {
          'Accept': 'application/json;odata=verbose'
        },
        withCredentials: true,
        onload: res => {
          if (res.status >= 200 && res.status < 300) {
            try {
              const fetched = JSON.parse(res.responseText).d.ListItemEntityTypeFullName;
              if (fetched) {
                _listTypes[key] = fetched;
                console.log('[WorkPulse] ✅ List type fetched —', listName, ':', fetched);
                resolve(fetched);
                return;
              }
            } catch(ex) {
              console.error('[WorkPulse] getListType parse error:', listName, ex.message, res.responseText.slice(0,200));
            }
          } else {
            console.error('[WorkPulse] getListType HTTP', res.status, 'for', listName, '— response:', res.responseText.slice(0,200));
          }
          // Fallback: use standard SP naming convention
          const fallback = 'SP.Data.' + listName.replace(/[^a-zA-Z0-9]/g,'') + 'ListItem';
          _listTypes[key] = fallback;
          console.warn('[WorkPulse] getListType fallback for', listName, ':', fallback);
          resolve(fallback);
        },
        onerror: (err) => {
          const fallback = 'SP.Data.' + listName.replace(/[^a-zA-Z0-9]/g,'') + 'ListItem';
          _listTypes[key] = fallback;
          console.error('[WorkPulse] getListType network error for', listName, err);
          resolve(fallback);
        }
      });
    });
  }
  async function getTaskListType() { return getListType(SP.TASK_LIST, 'task'); }
  async function getAttListType()  { return getListType(SP.STATUS_LIST, 'att'); }

  async function postTask(t) {
    // ── Exact v2.10.0 logic — uses spInsert() for full error surfacing ──
    const type = await getTaskListType();
    if (!type) { console.warn('[WorkPulse] postTask: could not get list type'); return false; }

    // Build body with only columns that exist in the DailyTaskReport SP list
    const body = {
      'Title':        (t.employeeName || authState.name || '') + '_' + (t.date || todayStr()),
      'EmployeeName': t.employeeName  || authState.name || '',
      'TaskType':     t.taskType      || '',
      'HoursWorked':  parseFloat(t.hours) || 0,
      'NPTHours':     parseFloat(t.npt)   || 0,
      'WorkType':     t.workType      || 'Productive',
      'AdHocDetails': t.adhoc         || '',
      'Shift':        t.shift         || '',
      'TaskDate':     localDateSP(t.date || todayStr()),
      'SubmittedAt':  t.submittedAt   || nowUTC()
    };

    console.log('[WorkPulse] postTask → spInsert:', SP.TASK_LIST, 'date:', body.TaskDate);
    const ok = await spInsert(SP.TASK_LIST, type, body);
    if (ok) {
      console.log('[WorkPulse] ✅ Task saved to SP:', t.employeeName, t.taskType, t.date);
    } else {
      console.warn('[WorkPulse] ⚠ Task failed SP save — queued for retry');
    }
    return ok;
  }

  // Cache the NPT list entity type so we only fetch it once
  // getNPTListType delegates to the authenticated getListType() cache
  async function getNPTListType() {
    return getListType(SP.NPT_LIST, 'npt');
  }

  async function postNPT(e) {
    const token = await getDigest();
    if (!token) { console.warn('[WorkPulse] postNPT: no digest token'); return false; }
    const listType = await getNPTListType();
    const dateStr  = e.date || todayStr();
    const bodyData = {
      'Title':        (e.type || 'NPT') + ' - ' + (e.name || authState.name || '') + ' - ' + dateStr,
      'EmployeeName': e.name      || authState.name || '',
      'NPTDate':      localDateSP(dateStr),
      'NPTType':      e.type      || '',
      'DurationMins': parseInt(e.minutes) || 0,
      'Description':  e.desc      || '',
      'LoggedAt':     e.loggedAt  || nowUTC()
    };
    // Add __metadata only if we have a real list type (not empty)
    const fullBody = listType
      ? Object.assign({ '__metadata': { 'type': listType } }, bodyData)
      : bodyData;
    console.log('[WorkPulse] postNPT → POST to', SP.NPT_LIST, '| type:', listType || 'none', '| body:', bodyData);
    return new Promise(resolve => {
      GM_xmlhttpRequest({
        method: 'POST',
        url: SP.SITE + "/_api/web/lists/GetByTitle('" + SP.NPT_LIST + "')/items",
        headers: {
          'Accept':          'application/json;odata=verbose',
          'Content-Type':    'application/json;odata=verbose',
          'X-RequestDigest': token
        },
        data: JSON.stringify(fullBody),
        withCredentials: true,
        onload: res => {
          console.log('[WorkPulse] postNPT response:', res.status, res.responseText.slice(0,400));
          if (res.status >= 200 && res.status < 300) {
            console.log('[WorkPulse] ✅ NPT saved to SP:', e.type, e.minutes+'m', dateStr);
            resolve(true);
          } else {
            let msg = '';
            try {
              const j = JSON.parse(res.responseText);
              msg = (j.error && j.error.message)
                ? (j.error.message.value || JSON.stringify(j.error.message))
                : res.responseText.slice(0, 300);
            } catch { msg = res.responseText.slice(0, 300); }
            console.error('[WorkPulse] postNPT FAILED:', res.status, msg);
            toast('❌ NPT save failed (' + res.status + '): ' + msg.slice(0, 150), 'err');
            resolve(false);
          }
        },
        onerror: err => {
          console.error('[WorkPulse] postNPT network error:', err);
          toast('❌ NPT network error — is SP open in a tab?', 'err');
          resolve(false);
        }
      });
    });
  }

  async function postAttendance(e) {
    const dateStr  = e.date || todayStr();
    const spDate   = localDateSP(dateStr);   // UTC ISO for SP storage
    const dateOnly = dateStr;                 // YYYY-MM-DD for filter
    const title    = (e.name || '') + ' - ' + dateStr;
    const nameEsc  = (e.name || '').replace(/'/g, "''");
    // Filter by date range (start of day to end of day in UTC)
    const dayStart = new Date(dateStr + 'T00:00:00').toISOString();
    const dayEnd   = new Date(dateStr + 'T23:59:59').toISOString();
    // Check if record exists for this person+date
    return new Promise(resolve => {
      GM_xmlhttpRequest({
        method: 'GET',
        url: SP.SITE + "/_api/web/lists/GetByTitle('" + SP.STATUS_LIST + "')/items" +
             "?$filter=EmployeeName eq '" + nameEsc + "' and StatusDate ge datetime'" + dayStart + "' and StatusDate le datetime'" + dayEnd + "'&$select=Id&$top=1",
        headers: { 'Accept': 'application/json;odata=verbose' },
        withCredentials: true,
        onload: async res => {
          try {
            const results = (JSON.parse(res.responseText).d || {}).results || [];
            // WorkStatus stores full status including sub-type e.g. "SL - Half AM"
            const body = {
              'Title':        title,
              'EmployeeName': e.name      || '',
              'StatusDate':   spDate,
              'WorkStatus':   e.status    || '',
              'Shift':        e.shift     || '',
              'UpdatedAt':    e.updatedAt || nowUTC()
            };
            // Use dynamic list type (fetched once, cached)
            const meta = await getAttListType();
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

  // ── MANDATORY NATIONAL HOLIDAYS ─────────────────────────────────────
  const MANDATORY_HOLIDAYS = [
    { month:1,  day:26, name:'Republic Day'        },
    { month:4,  day:15, name:'Ambedkar Jayanti'    },
    { month:5,  day:1,  name:'Labour Day'          },
    { month:6,  day:2,  name:'Telangana Day'       },
    { month:10, day:2,  name:'Gandhi Jayanti'      },
  ];

  function applyMandatoryHolidays(userName) {
    const years = [new Date().getFullYear(), new Date().getFullYear()+1];
    let changed = false;
    years.forEach(yr => {
      MANDATORY_HOLIDAYS.forEach(h => {
        const dk = yr+'-'+String(h.month).padStart(2,'0')+'-'+String(h.day).padStart(2,'0');
        if (!statusCache[dk] || statusCache[dk].mandatory !== true) {
          statusCache[dk] = { status:'Mandatory Off', name:h.name, date:dk, mandatory:true, updatedAt:nowUTC() };
          changed = true;
        }
      });
    });
    if (changed) safeSaveObj('dtr_status2', statusCache);
  }

  function isMandatoryOff(dk) {
    return !!(statusCache[dk] && statusCache[dk].mandatory === true);
  }


  async function flushQueue() {
    const q2 = safeLoad('dtr_spq2', []); if (!q2.length) return;
    const rem = [];
    for (const t of q2) { const s = await postTask(t); if (!s) rem.push(t); }
    safeSave('dtr_spq2', rem);
  }

  async function flushQueueManual() {
    toast('Retrying queued items...', 'info');
    clearDigestCache(); // force fresh token
    await flushQueue();
    const r = safeLoad('dtr_spq2', []).length;
    toast(r === 0 ? '✅ All synced' : '⚠️ ' + r + ' still pending', r === 0 ? 'ok' : 'err');
  }

  async function testNPTConnection() {
    toast('Step 1/4 — Getting SP auth token...', 'info');

    // ── Step 1: Get digest ───────────────────────────────────────────────
    const token = await getDigest();
    if (!token) {
      toast('❌ Step 1 FAILED — No SP token. Open ' + SP.SITE + ' in a tab first.', 'err');
      return;
    }
    toast('✅ Step 1 — Auth token OK. Step 2/4 — Fetching list type...', 'ok');

    // ── Step 2: Fetch real list type ─────────────────────────────────────
    _listTypes.npt = ''; // clear cache
    const listType = await getListType(SP.NPT_LIST, 'npt');
    console.log('[WorkPulse] NPT list type fetched:', listType, '| List:', SP.NPT_LIST);
    toast('✅ Step 2 — List type: ' + listType + '. Step 3/4 — Verifying list exists...', 'ok');

    // ── Step 3: Verify list exists by fetching its info ──────────────────
    const listOk = await new Promise(resolve => {
      GM_xmlhttpRequest({
        method: 'GET',
        url: SP.SITE + "/_api/web/lists/GetByTitle('" + SP.NPT_LIST + "')?$select=ItemCount",
        headers: { 'Accept': 'application/json;odata=verbose' },
        withCredentials: true,
        onload: res => {
          if (res.status >= 200 && res.status < 300) {
            try {
              const cnt = JSON.parse(res.responseText).d.ItemCount;
              toast('✅ Step 3 — List \''+SP.NPT_LIST+'\' found ('+cnt+' items). Step 4/4 — Inserting...', 'ok');
              resolve(true);
            } catch(ex) {
              toast('❌ Step 3 — List found but response parse failed: ' + ex.message, 'err');
              resolve(false);
            }
          } else {
            let msg='';
            try { msg=JSON.parse(res.responseText).error.message.value||''; } catch{}
            toast('❌ Step 3 — List NOT found ('+res.status+'): '+(msg||'List name may be wrong. Current: \''+SP.NPT_LIST+'\''), 'err');
            console.error('[WorkPulse] List check failed:', res.status, res.responseText.slice(0,300));
            resolve(false);
          }
        },
        onerror: () => { toast('❌ Step 3 — Network error checking list', 'err'); resolve(false); }
      });
    });
    if (!listOk) return;

    // ── Step 4: Direct POST insert (minimal body, no __metadata first) ───
    // Try without __metadata first — some SP configs accept this
    const bodyNoMeta = JSON.stringify({
      'Title':        'TEST-DELETE-ME ' + new Date().toISOString(),
      'EmployeeName': authState.name || 'Test',
      'NPTDate':      localDateSP(todayStr()),
      'NPTType':      'System Issue',
      'DurationMins': 1,
      'Description':  'WorkPulse connection test — safe to delete',
      'LoggedAt':     nowUTC()
    });

    const insertUrl = SP.SITE + "/_api/web/lists/GetByTitle('" + SP.NPT_LIST + "')/items";

    // First try: with __metadata (standard)
    const bodyWithMeta = JSON.stringify(Object.assign({ '__metadata': { 'type': listType } }, JSON.parse(bodyNoMeta)));

    await new Promise(resolve => {
      GM_xmlhttpRequest({
        method: 'POST',
        url: insertUrl,
        headers: {
          'Accept':          'application/json;odata=verbose',
          'Content-Type':    'application/json;odata=verbose',
          'X-RequestDigest': token
        },
        data: bodyWithMeta,
        withCredentials: true,
        onload: res => {
          console.log('[WorkPulse] NPT test insert response:', res.status, res.responseText.slice(0,500));
          if (res.status >= 200 && res.status < 300) {
            toast('✅ NPT INSERT SUCCESS! (\''+SP.NPT_LIST+'\' with type '+listType+'). Delete TEST item from SP.', 'ok');
          } else {
            // Extract full SP error
            let msg = res.responseText;
            try {
              const j = JSON.parse(res.responseText);
              msg = (j.error && j.error.message) ? (j.error.message.value || JSON.stringify(j.error.message)) : msg;
            } catch {}
            console.error('[WorkPulse] Insert FAILED:', res.status, msg);
            toast('❌ INSERT FAILED ('+res.status+'): '+msg.slice(0,150), 'err');
          }
          resolve();
        },
        onerror: err => {
          console.error('[WorkPulse] Insert network error:', err);
          toast('❌ Network error during insert', 'err');
          resolve();
        }
      });
    });
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
    if (!authState.name) return;
    const nameEsc = authState.name.replace(/'/g, "''");
    const url = SP.SITE + "/_api/web/lists/GetByTitle('" + SP.STATUS_LIST + "')/items" +
      "?$filter=EmployeeName eq '" + nameEsc + "'" +
      "&$top=5000&$orderby=StatusDate desc";
    GM_xmlhttpRequest({
      method: 'GET', url,
      headers: { 'Accept': 'application/json;odata=verbose' },
      withCredentials: true,
      onload: res => {
        if (res.status < 200 || res.status >= 300) {
          let msg=''; try{msg=JSON.parse(res.responseText).error.message.value||'';}catch{}
          console.error('[WorkPulse] fetchOwnAttendance:', res.status, msg||res.responseText.slice(0,200));
          if (!silent) toast('❌ Attendance sync failed (' + res.status + '): ' + (msg||'check F12'), 'err');
          return;
        }
        try {
          const parsed = JSON.parse(res.responseText);
          if (!parsed.d) {
            console.error('[WorkPulse] fetchOwnAttendance: no .d in response', res.responseText.slice(0,200));
            if (!silent) toast('❌ Attendance list error — check list name: ' + SP.STATUS_LIST, 'err');
            return;
          }
          const items = parsed.d.results || [];
          items.forEach(it => {
            const dk = (it.StatusDate||'').split('T')[0];
            if (dk) {
              statusCache[dk] = {
                status:  it.WorkStatus || '',
                shift:   it.Shift      || '',
                date:    dk,
                name:    authState.name,
                _fromSP: true
              };
            }
          });
          safeSaveObj('dtr_status2', statusCache);
          console.log('[WorkPulse] fetchOwnAttendance:', items.length, 'records from AttendanceLog');
          if (!silent) toast('✅ Attendance synced — ' + items.length + ' records', 'ok');
          if (currentView === 'weekly')   renderWeeklyCalendar();
          if (currentView === 'calendar') renderMyCalendar();
          if (currentView === 'analytics') renderAnalytics();
          const todaySt = statusCache[todayStr()];
          const pill = root.querySelector('.dtr-topbar-r .sp');
          if (pill && todaySt) {
            pill.className = 'sp sp-' + todaySt.status.toLowerCase().replace(' ','-');
            pill.textContent = todaySt.status;
          }
        } catch(ex) {
          console.error('[WorkPulse] fetchOwnAttendance parse:', ex.message, res.responseText.slice(0,200));
          if (!silent) toast('❌ Attendance sync error: ' + ex.message, 'err');
        }
      },
      onerror: () => {
        console.error('[WorkPulse] fetchOwnAttendance: network error');
        if (!silent) toast('❌ Cannot reach SP for attendance', 'err');
      }
    });
  }

  function fetchAllAttendance() {
    // Associates only fetch their OWN data — full team fetch is admin-only
    fetchOwnAttendance(false);
  }

  function fetchOwnTasks(silent) {
    if (!authState.name) return;
    const nameEsc = authState.name.replace(/'/g, "''");
    // No $select — fetch all columns to avoid missing column errors
    const url = SP.SITE + "/_api/web/lists/GetByTitle('" + SP.TASK_LIST + "')/items" +
      "?$filter=EmployeeName eq '" + nameEsc + "'" +
      "&$top=500&$orderby=TaskDate desc";
    GM_xmlhttpRequest({
      method: 'GET', url: url,
      headers: { 'Accept': 'application/json;odata=verbose' },
      withCredentials: true,
      onload: res => {
        if (res.status < 200 || res.status >= 300) {
          let msg=''; try{msg=JSON.parse(res.responseText).error.message.value||'';}catch{}
          console.error('[WorkPulse] fetchOwnTasks:', res.status, msg||res.responseText.slice(0,200));
          if (!silent) toast('❌ Task sync failed (' + res.status + '): ' + (msg||'check F12'), 'err');
          return;
        }
        try {
          const parsed = JSON.parse(res.responseText);
          if (!parsed.d) {
            console.error('[WorkPulse] fetchOwnTasks: no .d', res.responseText.slice(0,200));
            if (!silent) toast('❌ Task list error — check list name: ' + SP.TASK_LIST, 'err');
            return;
          }
          const items = parsed.d.results || [];
          // Keep only own entries then merge from SP
          submissions = submissions.filter(s => s && s.employeeName === authState.name);
          const existingKeys = new Set(submissions.map(s => (s.submittedAt||'')+'_'+(s.taskType||'')+'_'+(s.date||'')));
          let added = 0;
          items.forEach(it => {
            const spDate = (it.TaskDate||it.Created||'').split('T')[0];
            const key    = (it.SubmittedAt||it.Created||'')+'_'+(it.TaskType||'')+'_'+spDate;
            if (!existingKeys.has(key)) {
              submissions.push({
                employeeName: it.EmployeeName  || authState.name,
                taskType:     it.TaskType      || '',
                hours:        parseFloat(it.HoursWorked) || 0,
                npt:          parseFloat(it.NPTHours)    || 0,
                workType:     it.WorkType      || 'Productive',
                adhoc:        it.AdHocDetails  || '',
                shift:        it.Shift         || '',
                date:         spDate,
                submittedAt:  it.SubmittedAt   || it.Created || ''
              });
              added++;
            }
          });
          if (added > 0) {
            safeSave('dtr_subs_' + authState.name.toLowerCase().replace(/[^a-z0-9]+/g,'_'), submissions);
            updateSBStats();
            if (currentView === 'analytics') renderAnalytics();
            if (currentView === 'tracker')   renderTracker();
          }
          if (!silent) toast('✅ ' + items.length + ' task(s) synced from SharePoint', 'ok');
        } catch(ex) {
          if (!silent) toast('❌ Task sync error: ' + ex.message, 'err');
        }
      },
      onerror: () => { if (!silent) toast('❌ Cannot reach SharePoint', 'err'); }
    });
  }

  function fetchOwnNPT(silent) {
    if (!authState.name) return;
    const nameEsc = authState.name.replace(/'/g, "''");
    // No $select — fetch all columns so missing columns don't cause errors
    const url = SP.SITE + "/_api/web/lists/GetByTitle('" + SP.NPT_LIST + "')/items" +
      "?$filter=EmployeeName eq '" + nameEsc + "'" +
      "&$top=500&$orderby=Created desc";
    GM_xmlhttpRequest({
      method: 'GET', url,
      headers: { 'Accept': 'application/json;odata=verbose' },
      withCredentials: true,
      onload: res => {
        // Check HTTP status first
        if (res.status < 200 || res.status >= 300) {
          let spErr = '';
          try { spErr = JSON.parse(res.responseText).error.message.value || ''; } catch {}
          console.error('[WorkPulse] fetchOwnNPT HTTP', res.status, spErr || res.responseText.slice(0,200));
          if (!silent) toast('❌ NPT sync failed (' + res.status + '): ' + (spErr||'check F12 console'), 'err');
          return;
        }
        try {
          const parsed = JSON.parse(res.responseText);
          const items  = (parsed.d && parsed.d.results) ? parsed.d.results : [];
          const existingKeys = new Set(nptCache.map(n => (n.name||'')+'::'+((n.loggedAt||n.Created||''))));
          let added = 0;
          items.forEach(it => {
            const loggedAt = it.LoggedAt || it.Created || '';
            const key = (it.EmployeeName||'') + '::' + loggedAt;
            if (!existingKeys.has(key)) {
              nptCache.push({
                name:      it.EmployeeName || authState.name,
                date:      (it.NPTDate || it.Created || '').split('T')[0],
                type:      it.NPTType      || '',
                minutes:   parseInt(it.DurationMins) || 0,
                desc:      it.Description  || '',
                loggedAt,
                savedToSP: true
              });
              added++;
            }
          });
          if (added > 0) {
            safeSave('dtr_npt2', nptCache);
            if (currentView === 'missednpt') renderMissedNPT();
          }
          if (!silent) toast('✅ NPT synced — ' + items.length + ' entries from SP', 'ok');
          console.log('[WorkPulse] fetchOwnNPT:', items.length, 'items from SP');
        } catch(ex) {
          console.error('[WorkPulse] fetchOwnNPT parse error:', ex.message, res.responseText.slice(0,300));
          if (!silent) toast('❌ NPT parse error: ' + ex.message, 'err');
        }
      },
      onerror: () => {
        console.error('[WorkPulse] fetchOwnNPT: network error');
        if (!silent) toast('❌ Cannot reach SP for NPT sync', 'err');
      }
    });
  }
  // flushQueue defined above
  // flushQueueManual defined above


  // ═══════════════════════════════════════════════════════════════════════
  // ADMIN PORTAL FUNCTIONS
  // ═══════════════════════════════════════════════════════════════════════

  function getAllMembers() {
    const fromSubs = [...new Set((Array.isArray(submissions)?submissions:[]).map(s=>s.employeeName).filter(Boolean))];
    const fromAtt  = [...new Set(Object.keys(teamStatusCache_adm).map(k=>k.split('::')[0]).filter(Boolean))];
    const fromNPT  = [...new Set((Array.isArray(nptAllCache)?nptAllCache:[]).map(n=>n.name||n.employeeName).filter(Boolean))];
    return [...new Set([...fromSubs,...fromAtt,...fromNPT,...ASSOCIATES])].sort();
  }

  function calcTeamAvgProd(subs) {
    const members=[...new Set(subs.map(s=>s.employeeName).filter(Boolean))];
    if(!members.length) return 0;
    return members.reduce((a,m)=>a+calcAvgProd(subs.filter(s=>s.employeeName===m)),0)/members.length;
  }

  function getTeamShift(user,dk)   { return(teamStatusCache_adm[user+'::'+dk]||teamStatusCache[user+'::'+dk]||{}).shift||''; }

  function renderAdminAnalytics() {
    const el = q('#view-admanalytics'); if (!el) return;
    const toHM = m => { const h=Math.floor(m/60),mn=m%60; return h>0?(mn>0?h+'h '+mn+'m':h+'h'):(mn+'m'); };
    const toC  = v  => v>=80?'#30d158':v>=60?'#ff9f0a':'#ff453a';
    const allMembers = [...new Set(submissions.map(s=>s.employeeName).filter(Boolean))].sort();

    if (!allMembers.length) {
      el.innerHTML =
        '<div class="ph"><div class="ph-left"><div class="ph-title">Associate Analytics</div></div></div>' +
        '<div class="card" style="text-align:center;padding:60px;">' +
          '<div style="font-size:2.5rem;margin-bottom:12px;">📊</div>' +
          '<div style="font-size:.95rem;font-weight:700;color:var(--text);margin-bottom:6px;">No data loaded yet</div>' +
          '<div style="font-size:.85rem;color:var(--text3);">Go to <strong>Team Tracker → Sync</strong> to load data from SharePoint first.</div>' +
        '</div>';
      return;
    }

    if (!aanSelectedAssoc || !allMembers.includes(aanSelectedAssoc)) aanSelectedAssoc = allMembers[0];
    if (aanCompareSet.size === 0) {
      aanCompareSet.add(allMembers[0]);
      if (allMembers[1]) aanCompareSet.add(allMembers[1]);
    }

    const CHIP_COLORS = ['#4f9eff','#30d158','#ff9f0a','#a371f7'];

    /* ── Core stats calculator ────────────────────────────────────────── */
    const getStats = name => {
      const rows  = submissions.filter(x => x.employeeName === name);
      const prod  = rows.filter(x => x.workType !== 'NPT' && x.taskType !== 'Leave');
      const npt   = rows.filter(x => x.workType === 'NPT');
      const days  = [...new Set(rows.map(x => x.date).filter(Boolean))];
      const pMins = prod.reduce((a,x) => a + (x.minutes || Math.round((x.hours||0)*60)), 0);
      const nMins = npt.reduce((a,x)  => a + (x.minutes || Math.round((x.hours||0)*60)), 0);
      // Productivity EXCLUDING NPT = prodMins / (days × 480) × 100
      const pctExNPT = days.length ? Math.min(100, Math.round(pMins / (days.length*480) * 100)) : 0;
      // Productivity INCLUDING NPT = prodMins / (prod+npt) × 100
      const pctInNPT = (pMins+nMins) ? Math.min(100, Math.round(pMins / (pMins+nMins) * 100)) : 0;
      // Task breakdown
      const tasks = {};
      prod.forEach(x => { tasks[x.taskType] = (tasks[x.taskType]||0) + 1; });
      const topTask = Object.entries(tasks).sort((a,b)=>b[1]-a[1])[0]?.[0] || '—';
      // Attendance
      const _att = (typeof attCache !== 'undefined' && Array.isArray(attCache)) ? attCache : [];
      const att  = _att.filter(a => (a.name||a.employeeName||'').toLowerCase() === name.toLowerCase());
      // Weekly grouping
      const byWeek = {};
      days.forEach(d => {
        const dt = new Date(d+'T12:00:00'), mon = new Date(dt);
        mon.setDate(dt.getDate() - dt.getDay() + 1);
        const wk = mon.toISOString().slice(0,10);
        if (!byWeek[wk]) byWeek[wk] = {prod:0,npt:0,days:0};
        const ds = rows.filter(x => x.date === d);
        byWeek[wk].prod += ds.filter(x=>x.workType!=='NPT').reduce((a,x)=>a+(x.minutes||Math.round((x.hours||0)*60)),0);
        byWeek[wk].npt  += ds.filter(x=>x.workType==='NPT').reduce((a,x)=>a+(x.minutes||Math.round((x.hours||0)*60)),0);
        byWeek[wk].days++;
      });
      // Monthly grouping
      const byMonth = {};
      days.forEach(d => {
        const mo = d.slice(0,7);
        if (!byMonth[mo]) byMonth[mo] = {prod:0,npt:0,days:0};
        const ds = rows.filter(x => x.date === d);
        byMonth[mo].prod += ds.filter(x=>x.workType!=='NPT').reduce((a,x)=>a+(x.minutes||Math.round((x.hours||0)*60)),0);
        byMonth[mo].npt  += ds.filter(x=>x.workType==='NPT').reduce((a,x)=>a+(x.minutes||Math.round((x.hours||0)*60)),0);
        byMonth[mo].days++;
      });
      return {
        name, entries:rows.length, pMins, nMins, days:days.length,
        pctExNPT, pctInNPT, topTask, byWeek, byMonth,
        wfo : att.filter(a=>/^WFO/i.test(a.status||'')).length,
        wfh : att.filter(a=>/^WFH/i.test(a.status||'')).length,
        lv  : att.filter(a=>/^(SL|CL|AL|LEAVE)/i.test(a.status||'')).length
      };
    };

    /* ── Mode toggle bar ──────────────────────────────────────────────── */
    const MODES  = ['single','compare','weekly','monthly'];
    const MLABEL = {single:'Individual',compare:'Compare',weekly:'Weekly',monthly:'Monthly'};
    const modeBar =
      '<div style="display:inline-flex;border:1px solid var(--border);border-radius:10px;overflow:hidden;margin-left:auto;">' +
      MODES.map(m =>
        '<button data-action="aan-mode" data-val="'+m+'" style="padding:6px 14px;border:none;' +
        'background:'+(aanMode===m?'var(--accent)':'transparent')+';' +
        'color:'+(aanMode===m?'#fff':'var(--text2)')+';' +
        'font-size:.8rem;font-weight:700;cursor:pointer;font-family:inherit;">' +
        MLABEL[m]+'</button>'
      ).join('') +
      '</div>';

    let html =
      '<div class="ph" style="display:flex;align-items:center;flex-wrap:wrap;gap:8px;">' +
        '<div class="ph-left">' +
          '<div class="ph-title">Associate Analytics</div>' +
          '<div class="ph-sub">Individual · Compare · Weekly · Monthly</div>' +
        '</div>' + modeBar +
      '</div>';

    /* ── Formula card (always visible) ───────────────────────────────── */
    const formulaCard =
      '<div class="card" style="padding:14px 18px;margin-bottom:14px;border-left:3px solid var(--accent);">' +
        '<div style="font-size:.7rem;font-weight:800;text-transform:uppercase;letter-spacing:.8px;color:var(--accent);margin-bottom:10px;">📐 How Productivity % is Calculated</div>' +
        '<div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;">' +
          '<div style="background:var(--bg3);border-radius:10px;padding:12px;">' +
            '<div style="font-size:.82rem;font-weight:700;color:var(--text);margin-bottom:6px;">Excl. NPT — Task Productivity</div>' +
            '<div style="font-family:var(--mono);font-size:.78rem;color:#4f9eff;background:var(--bg4);padding:6px 8px;border-radius:6px;margin-bottom:6px;">Productive Mins ÷ (Days × 480) × 100</div>' +
            '<div style="font-size:.74rem;color:var(--text3);">How much of scheduled 8h/day was spent on real tasks. NPT is ignored in the denominator.</div>' +
          '</div>' +
          '<div style="background:var(--bg3);border-radius:10px;padding:12px;">' +
            '<div style="font-size:.82rem;font-weight:700;color:var(--text);margin-bottom:6px;">Incl. NPT — Time Utilisation</div>' +
            '<div style="font-family:var(--mono);font-size:.78rem;color:#30d158;background:var(--bg4);padding:6px 8px;border-radius:6px;margin-bottom:6px;">Productive Mins ÷ (Productive + NPT) × 100</div>' +
            '<div style="font-size:.74rem;color:var(--text3);">Ratio of productive to all logged time. High = less NPT overhead relative to work output.</div>' +
          '</div>' +
        '</div>' +
      '</div>';

    /* ── Pill selector helper ─────────────────────────────────────────── */
    const pillRow = (action, label) =>
      '<div class="card" style="padding:12px 16px;margin-bottom:14px;">' +
        '<div style="font-size:.68rem;font-weight:800;text-transform:uppercase;letter-spacing:.6px;color:var(--text3);margin-bottom:10px;">' + label + '</div>' +
        '<div style="display:flex;flex-wrap:wrap;gap:5px;">' +
        allMembers.map(m => {
          const sel = action==='aan-select' ? m===aanSelectedAssoc : aanCompareSet.has(m);
          return '<button data-action="'+action+'" data-val="'+m+'" style="padding:4px 12px;border-radius:20px;border:1.5px solid '+(sel?'var(--accent)':'var(--border)')+';background:'+(sel?'var(--accent)':'transparent')+';color:'+(sel?'#fff':'var(--text2)')+';font-size:.75rem;font-weight:600;cursor:pointer;font-family:inherit;">'+m+'</button>';
        }).join('') +
        '</div>' +
      '</div>';

    /* ════════════════════════════════════════════════════════════════════
       MODE: INDIVIDUAL
    ════════════════════════════════════════════════════════════════════ */
    if (aanMode === 'single') {
      const s = getStats(aanSelectedAssoc);
      html += pillRow('aan-select','Select Associate') + formulaCard;

      // 5 stat cards
      html +=
        '<div style="display:grid;grid-template-columns:repeat(5,1fr);gap:12px;margin-bottom:14px;">' +
          '<div class="stat-card ab"><div class="lbl">Entries</div><div class="val">'+s.entries+'</div></div>' +
          '<div class="stat-card gb"><div class="lbl">Productive</div><div class="val" style="font-size:1.2rem;">'+toHM(s.pMins)+'</div></div>' +
          '<div class="stat-card amb"><div class="lbl">NPT Time</div><div class="val" style="font-size:1.2rem;color:var(--amber);">'+toHM(s.nMins)+'</div></div>' +
          '<div class="stat-card" style="border-top:3px solid '+toC(s.pctExNPT)+';"><div class="lbl">Excl. NPT %</div><div class="val" style="color:'+toC(s.pctExNPT)+';">'+s.pctExNPT+'%</div><div style="font-size:.62rem;color:var(--text3);margin-top:2px;">Task Productivity</div></div>' +
          '<div class="stat-card" style="border-top:3px solid '+toC(s.pctInNPT)+';"><div class="lbl">Incl. NPT %</div><div class="val" style="color:'+toC(s.pctInNPT)+';">'+s.pctInNPT+'%</div><div style="font-size:.62rem;color:var(--text3);margin-top:2px;">Time Utilisation</div></div>' +
        '</div>';

      // Attendance + Task breakdown
      html +=
        '<div style="display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:14px;">' +
          '<div class="card"><div class="card-title">Attendance</div>' +
            '<div style="display:grid;grid-template-columns:repeat(3,1fr);gap:8px;">' +
              '<div style="text-align:center;padding:10px;background:rgba(48,209,88,.08);border:1px solid rgba(48,209,88,.18);border-radius:10px;"><div style="font-size:1.5rem;font-weight:800;color:#30d158;">'+s.wfo+'</div><div style="font-size:.7rem;color:var(--text3);margin-top:2px;">WFO</div></div>' +
              '<div style="text-align:center;padding:10px;background:rgba(10,132,255,.08);border:1px solid rgba(10,132,255,.18);border-radius:10px;"><div style="font-size:1.5rem;font-weight:800;color:#0a84ff;">'+s.wfh+'</div><div style="font-size:.7rem;color:var(--text3);margin-top:2px;">WFH</div></div>' +
              '<div style="text-align:center;padding:10px;background:rgba(163,113,247,.08);border:1px solid rgba(163,113,247,.18);border-radius:10px;"><div style="font-size:1.5rem;font-weight:800;color:#a371f7;">'+s.lv+'</div><div style="font-size:.7rem;color:var(--text3);margin-top:2px;">Leave</div></div>' +
            '</div>' +
          '</div>' +
          '<div class="card"><div class="card-title">Task Breakdown</div>' + (()=>{
            const rows2 = submissions.filter(x=>x.employeeName===aanSelectedAssoc&&x.workType!=='NPT');
            const tk = {};
            rows2.forEach(x=>{tk[x.taskType]=(tk[x.taskType]||0)+1;});
            const top = Object.entries(tk).sort((a,b)=>b[1]-a[1]).slice(0,6);
            if (!top.length) return '<div class="empty">No task data</div>';
            const mx = top[0][1];
            return top.map(([t,n])=>
              '<div style="margin-bottom:8px;"><div style="display:flex;justify-content:space-between;font-size:.8rem;margin-bottom:3px;"><span style="font-weight:600;color:var(--text);">'+t+'</span><span style="color:var(--accent);font-weight:700;">'+n+'</span></div>' +
              '<div style="height:5px;background:var(--bg4);border-radius:3px;"><div style="height:100%;width:'+Math.round(n/mx*100)+'%;background:linear-gradient(90deg,var(--accent),var(--accent2));border-radius:3px;"></div></div></div>'
            ).join('');
          })() + '</div>' +
        '</div>';

    /* ════════════════════════════════════════════════════════════════════
       MODE: COMPARE
    ════════════════════════════════════════════════════════════════════ */
    } else if (aanMode === 'compare') {
      const ca = [...aanCompareSet];
      html += pillRow('aan-compare','Select up to 4 associates to compare') + formulaCard;

      if (!ca.length) {
        html += '<div class="empty"><p>Select at least one associate above</p></div>';
      } else {
        const stats = ca.map(getStats);

        // Comparison table with both % metrics
        html +=
          '<div class="card" style="margin-bottom:14px;"><div class="card-title">Side-by-Side Comparison</div>' +
          '<div style="overflow-x:auto;"><table class="dtr-table" style="width:100%;"><thead><tr>' +
            '<th>Metric</th>' + ca.map((m,i)=>'<th style="color:'+CHIP_COLORS[i]+';">'+m+'</th>').join('') +
          '</tr></thead><tbody>' +
          [
            ['Entries',       s=>s.entries],
            ['Productive',    s=>toHM(s.pMins)],
            ['NPT Time',      s=>toHM(s.nMins)],
            ['% Excl. NPT',   s=>'<span style="font-weight:800;color:'+toC(s.pctExNPT)+';">'+s.pctExNPT+'%</span>'],
            ['% Incl. NPT',   s=>'<span style="font-weight:800;color:'+toC(s.pctInNPT)+';">'+s.pctInNPT+'%</span>'],
            ['Days Worked',   s=>s.days],
            ['WFO Days',      s=>s.wfo],
            ['WFH Days',      s=>s.wfh],
            ['Top Task',      s=>s.topTask],
          ].map(([label,fn])=>
            '<tr><td style="font-weight:700;color:var(--text2);">'+label+'</td>' +
            stats.map((s,i)=>'<td style="color:'+CHIP_COLORS[i]+';">'+fn(s)+'</td>').join('') +
            '</tr>'
          ).join('') +
          '</tbody></table></div></div>';

        // 4 visual bar charts
        html +=
          '<div style="display:grid;grid-template-columns:1fr 1fr;gap:14px;">' +
          [
            {label:'% Excl. NPT — Task Productivity', key:'pctExNPT', max:100,  fmt:v=>v+'%'},
            {label:'% Incl. NPT — Time Utilisation',  key:'pctInNPT', max:100,  fmt:v=>v+'%'},
            {label:'Productive Time',                  key:'pMins',   max:null, fmt:v=>toHM(v)},
            {label:'NPT Time',                         key:'nMins',   max:null, fmt:v=>toHM(v)},
          ].map(({label,key,max,fmt}) => {
            const mx = max !== null ? max : Math.max(...stats.map(s=>s[key]||0), 1);
            return '<div class="card"><div class="card-title">'+label+'</div>' +
              stats.map((s,i) => {
                const v = s[key]||0, pct = Math.round(v/mx*100);
                return '<div style="margin-bottom:10px;">' +
                  '<div style="display:flex;justify-content:space-between;font-size:.8rem;margin-bottom:4px;">' +
                    '<span style="font-weight:700;color:'+CHIP_COLORS[i]+';">'+s.name+'</span>' +
                    '<span style="font-weight:800;color:'+CHIP_COLORS[i]+';">'+fmt(v)+'</span>' +
                  '</div>' +
                  '<div style="height:8px;background:var(--bg4);border-radius:4px;overflow:hidden;">' +
                    '<div style="height:100%;width:'+pct+'%;background:'+CHIP_COLORS[i]+';border-radius:4px;transition:width 1s;"></div>' +
                  '</div>' +
                '</div>';
              }).join('') +
            '</div>';
          }).join('') +
          '</div>';
      }

    /* ════════════════════════════════════════════════════════════════════
       MODE: WEEKLY
    ════════════════════════════════════════════════════════════════════ */
    } else if (aanMode === 'weekly') {
      const s = getStats(aanSelectedAssoc);
      html += pillRow('aan-select','Select Associate — Weekly View') + formulaCard;

      const weeks = Object.entries(s.byWeek).sort((a,b)=>b[0].localeCompare(a[0])).slice(0,12);
      html += '<div class="card"><div class="card-title" style="display:flex;justify-content:space-between;"><span>Weekly Analysis — '+aanSelectedAssoc+'</span><span style="font-weight:500;color:var(--text3);font-size:.75rem;">Last '+weeks.length+' weeks</span></div>';

      if (!weeks.length) {
        html += '<div class="empty">No weekly data available. Sync from Team Tracker first.</div>';
      } else {
        html +=
          '<div style="overflow-x:auto;"><table class="dtr-table" style="width:100%;">' +
          '<thead><tr>' +
            '<th>Week Of</th><th>Days</th><th>Productive</th><th>NPT</th>' +
            '<th style="color:#4f9eff;">% Excl. NPT</th><th style="color:#30d158;">% Incl. NPT</th>' +
            '<th>Progress</th>' +
          '</tr></thead><tbody>' +
          weeks.map(([wk,{prod,npt,days}]) => {
            const ex  = days ? Math.min(100,Math.round(prod/(days*480)*100)) : 0;
            const inc = (prod+npt) ? Math.min(100,Math.round(prod/(prod+npt)*100)) : 0;
            const pp  = Math.min(Math.round(prod/(days*480)*100),100);
            const np  = Math.min(Math.round(npt/(days*480)*100),100);
            return '<tr>' +
              '<td style="font-weight:700;">'+wk+'</td>' +
              '<td>'+days+'</td>' +
              '<td>'+toHM(prod)+'</td>' +
              '<td style="color:var(--amber);">'+toHM(npt)+'</td>' +
              '<td><b style="color:'+toC(ex)+';">'+ex+'%</b></td>' +
              '<td><b style="color:'+toC(inc)+';">'+inc+'%</b></td>' +
              '<td style="min-width:120px;">' +
                '<div style="height:7px;background:var(--bg4);border-radius:4px;overflow:hidden;display:flex;">' +
                  '<div style="width:'+pp+'%;background:#30d158;border-radius:4px 0 0 4px;"></div>' +
                  '<div style="width:'+np+'%;background:var(--amber);"></div>' +
                '</div>' +
              '</td>' +
            '</tr>';
          }).join('') +
          '</tbody></table></div>';
      }
      html += '</div>';

    /* ════════════════════════════════════════════════════════════════════
       MODE: MONTHLY
    ════════════════════════════════════════════════════════════════════ */
    } else if (aanMode === 'monthly') {
      const s = getStats(aanSelectedAssoc);
      html += pillRow('aan-select','Select Associate — Monthly View') + formulaCard;

      const months = Object.entries(s.byMonth).sort((a,b)=>b[0].localeCompare(a[0]));
      html += '<div class="card"><div class="card-title">Monthly Analysis — '+aanSelectedAssoc+'</div>';

      if (!months.length) {
        html += '<div class="empty">No monthly data available.</div>';
      } else {
        html +=
          '<div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(200px,1fr));gap:12px;">' +
          months.map(([mo,{prod,npt,days}]) => {
            const ex  = days ? Math.min(100,Math.round(prod/(days*480)*100)) : 0;
            const inc = (prod+npt) ? Math.min(100,Math.round(prod/(prod+npt)*100)) : 0;
            const label = new Date(mo+'-01').toLocaleDateString('en-IN',{month:'long',year:'numeric'});
            return '<div class="stat-card" style="padding:14px 16px;">' +
              '<div style="font-size:.85rem;font-weight:800;color:var(--text);margin-bottom:10px;">'+label+'</div>' +
              '<div style="display:grid;grid-template-columns:auto 1fr;gap:3px 8px;align-items:center;margin-bottom:10px;">' +
                '<span style="font-size:.72rem;color:var(--text3);">Days</span><span style="font-size:.82rem;font-weight:700;text-align:right;">'+days+'</span>' +
                '<span style="font-size:.72rem;color:var(--text3);">Productive</span><span style="font-size:.82rem;font-weight:700;color:#30d158;text-align:right;">'+toHM(prod)+'</span>' +
                '<span style="font-size:.72rem;color:var(--text3);">NPT</span><span style="font-size:.82rem;font-weight:700;color:var(--amber);text-align:right;">'+toHM(npt)+'</span>' +
              '</div>' +
              '<div style="display:grid;grid-template-columns:1fr 1fr;gap:6px;">' +
                '<div style="padding:7px;background:rgba(79,158,255,.08);border-radius:8px;text-align:center;">' +
                  '<div style="font-size:1.05rem;font-weight:800;color:'+toC(ex)+';">'+ex+'%</div>' +
                  '<div style="font-size:.62rem;color:var(--text3);">Excl. NPT</div>' +
                '</div>' +
                '<div style="padding:7px;background:rgba(48,209,88,.08);border-radius:8px;text-align:center;">' +
                  '<div style="font-size:1.05rem;font-weight:800;color:'+toC(inc)+';">'+inc+'%</div>' +
                  '<div style="font-size:.62rem;color:var(--text3);">Incl. NPT</div>' +
                '</div>' +
              '</div>' +
            '</div>';
          }).join('') +
          '</div>';
      }
      html += '</div>';
    }

    el.innerHTML = html;
  }


  function renderAdminOverview() {
    const el = q('#view-overview'); if (!el) return;
    const subs    = Array.isArray(submissions) ? submissions : [];
    const members = getAllMembers();
    const today   = todayStr();
    const d       = new Date();
    const dayStr  = d.toLocaleDateString('en-US',{weekday:'long',month:'long',day:'numeric',year:'numeric'});
    const todaySubs = subs.filter(s=>s.date===today);
    const wfoCnt  = members.filter(m=>getTeamStatus(m,today)==='WFO').length;
    const wfhCnt  = members.filter(m=>getTeamStatus(m,today)==='WFH').length;
    const lvCnt   = members.filter(m=>['SL','CL','AL'].includes(getTeamStatus(m,today))).length;
    const markd   = members.filter(m=>getTeamStatus(m,today)).length;
    const pendCnt = members.length - markd;
    const totalH  = subs.reduce((a,s)=>a+(parseFloat(s.hours)||0),0);
    const avgProd = calcTeamAvgProd(subs);
    const pc      = avgProd>=75?'var(--green)':avgProd>=50?'var(--amber)':'var(--red)';
    const circ    = 2*Math.PI*52, off=circ-(avgProd/100)*circ;

    const memberStats = members.map(m => {
      const ms=subs.filter(s=>s.employeeName===m);
      const st=getTeamStatus(m,today), sh=getTeamShift(m,today);
      const ini=m.split(/[\s,]+/).filter(Boolean).map(x=>x[0].toUpperCase()).join('').slice(0,2);
      return {m,ini,count:ms.length,prod:calcAvgProd(ms),todaySt:st,shift:sh};
    }).sort((a,b)=>b.prod-a.prod);

    el.innerHTML =
      '<div class="ph"><div class="ph-left"><div class="ph-title">Team Overview</div>' +
      '<div class="ph-sub">'+dayStr+'</div></div>' +
      '<div class="ph-actions"><button class="btn btn-ghost btn-sm" data-action="adm-sync-all">'+ic.sync+' Sync All</button></div></div>' +
      (members.length===0 ? '<div class="warn-banner">⚠ No team data. Click Sync All to load from SharePoint.</div>' : '') +
      // Banner stats
      '<div style="display:grid;grid-template-columns:repeat(6,1fr);gap:1px;background:var(--border);border-radius:12px;overflow:hidden;margin-bottom:16px">' +
      [['Members',members.length,'var(--accent)'],['WFO Today',wfoCnt,'#3fb950'],['WFH Today',wfhCnt,'#22d3ee'],['On Leave',lvCnt,'var(--amber)'],['Not Marked',pendCnt,'var(--red)'],['Total Hours',totalH.toFixed(0)+'h','var(--text)']].map(([l,n,c])=>
        '<div style="background:var(--bg2);padding:18px 12px;text-align:center"><div style="font-size:1.8rem;font-weight:800;color:'+c+';letter-spacing:-1px">'+n+'</div><div style="font-size:.7rem;text-transform:uppercase;letter-spacing:.7px;color:var(--text3);margin-top:3px">'+l+'</div></div>'
      ).join('')+'</div>'+
      // ── WFO / WFH Today panel ────────────────────────────────────────
      (()=>{
        const wfoList2 = members.filter(m => getTeamStatus(m,today)==='WFO');
        const wfhList2 = members.filter(m => getTeamStatus(m,today)==='WFH');
        const pill2 = (m,bg,dot) =>
          '<span style="display:inline-flex;align-items:center;gap:5px;padding:3px 10px 3px 4px;border-radius:20px;background:'+bg+';margin:3px;">'
          +'<span style="width:20px;height:20px;border-radius:50%;background:'+dot+';color:#fff;font-size:.62rem;font-weight:800;display:inline-flex;align-items:center;justify-content:center;">'+m.slice(0,2).toUpperCase()+'</span>'
          +'<span style="font-size:.78rem;font-weight:600;color:var(--text);">'+m+'</span></span>';
        return '<div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:16px;">'
          +'<div class="card" style="padding:16px 18px;">'
          +'<div style="display:flex;align-items:center;gap:8px;margin-bottom:12px;">'
          +'<span style="width:9px;height:9px;border-radius:50%;background:#30d158;display:inline-block;"></span>'
          +'<span style="font-size:.7rem;font-weight:800;text-transform:uppercase;letter-spacing:.8px;color:var(--text3);">WFO Today</span>'
          +'<span style="margin-left:auto;font-size:.95rem;font-weight:800;color:#30d158;">'+wfoList2.length+'</span>'
          +'</div>'
          +(wfoList2.length
            ? '<div style="display:flex;flex-wrap:wrap;">'+wfoList2.map(m=>pill2(m,'rgba(48,209,88,.1)','#30d158')).join('')+'</div>'
            : '<p style="font-size:.82rem;color:var(--text3);font-style:italic;margin:0;">None marked yet</p>')
          +'</div>'
          +'<div class="card" style="padding:16px 18px;">'
          +'<div style="display:flex;align-items:center;gap:8px;margin-bottom:12px;">'
          +'<span style="width:9px;height:9px;border-radius:50%;background:#0a84ff;display:inline-block;"></span>'
          +'<span style="font-size:.7rem;font-weight:800;text-transform:uppercase;letter-spacing:.8px;color:var(--text3);">WFH Today</span>'
          +'<span style="margin-left:auto;font-size:.95rem;font-weight:800;color:#0a84ff;">'+wfhList2.length+'</span>'
          +'</div>'
          +(wfhList2.length
            ? '<div style="display:flex;flex-wrap:wrap;">'+wfhList2.map(m=>pill2(m,'rgba(10,132,255,.1)','#0a84ff')).join('')+'</div>'
            : '<p style="font-size:.82rem;color:var(--text3);font-style:italic;margin:0;">None marked yet</p>')
          +'</div>'
          +'</div>';
      })() +
      '<div class="g2" style="margin-bottom:16px"> ' +
      // Donut
      '<div class="card"><div class="chart-title" style="display:flex;justify-content:space-between"><span>Team Productivity</span><span style="color:'+pc+';font-weight:800">'+avgProd.toFixed(0)+'%</span></div>'+
      '<div style="display:flex;align-items:center;gap:20px;padding:8px 0">'+
      '<div style="position:relative;width:120px;height:120px;flex-shrink:0">'+
      '<svg width="120" height="120" viewBox="0 0 120 120" style="transform:rotate(-90deg)">'+
      '<circle cx="60" cy="60" r="52" fill="none" stroke="var(--bg4)" stroke-width="12"/>'+
      '<circle cx="60" cy="60" r="52" fill="none" stroke="'+pc+'" stroke-width="12" stroke-dasharray="'+circ.toFixed(1)+'" stroke-dashoffset="'+off.toFixed(1)+'" stroke-linecap="round" style="transition:stroke-dashoffset 1s ease"/></svg>'+
      '<div style="position:absolute;inset:0;display:flex;flex-direction:column;align-items:center;justify-content:center">'+
      '<div style="font-size:1.5rem;font-weight:800;color:'+pc+'">'+avgProd.toFixed(0)+'%</div>'+
      '<div style="font-size:.65rem;color:var(--text3)">avg</div></div></div>'+
      '<div style="display:flex;flex-direction:column;gap:7px;flex:1">' +
      [['WFO','#3fb950',wfoCnt],['WFH','#22d3ee',wfhCnt],['Leave','var(--amber)',lvCnt],['Pending','var(--text3)',pendCnt]].map(([l,c,n])=>
        '<div style="display:flex;align-items:center;gap:8px;font-size:.8rem;color:var(--text2)"><span style="width:10px;height:10px;border-radius:3px;background:'+c+';flex-shrink:0;display:inline-block"></span>'+l+'<span style="margin-left:auto;font-weight:700;color:var(--text)">'+n+'</span></div>'
      ).join('')+'</div></div></div>'+
      // Submissions today
      '<div class="card"><div class="chart-title">Today\'s Progress</div>'+
      '<div style="margin-bottom:10px"><div style="display:flex;justify-content:space-between;font-size:.8rem;margin-bottom:5px"><span style="color:var(--text2)">Submitted today</span><span style="font-weight:700">'+todaySubs.length+'/'+members.length+'</span></div>'+
      '<div style="height:8px;background:var(--bg4);border-radius:4px;overflow:hidden"><div style="height:100%;width:'+Math.round((todaySubs.length/Math.max(members.length,1))*100)+'%;background:linear-gradient(90deg,var(--accent),var(--accent2));border-radius:4px;transition:width .8s"></div></div></div>'+
      '<div style="margin-bottom:10px"><div style="display:flex;justify-content:space-between;font-size:.8rem;margin-bottom:5px"><span style="color:var(--text2)">Attendance marked</span><span style="font-weight:700">'+markd+'/'+members.length+'</span></div>'+
      '<div style="height:8px;background:var(--bg4);border-radius:4px;overflow:hidden"><div style="height:100%;width:'+Math.round((markd/Math.max(members.length,1))*100)+'%;background:linear-gradient(90deg,#3fb950,#22d3ee);border-radius:4px;transition:width .8s"></div></div></div>'+
      '<div style="display:flex;gap:4px;height:28px;border-radius:8px;overflow:hidden;margin-top:8px">'+
      [['#3fb950',wfoCnt,'WFO'],['#22d3ee',wfhCnt,'WFH'],['var(--amber)',lvCnt,'Leave'],['var(--bg4)',pendCnt,'Pending']].map(([c,n,l])=>
        n>0?'<div title="'+l+': '+n+'" style="background:'+c+';flex:'+n+';display:flex;align-items:center;justify-content:center;font-size:.65rem;font-weight:700;color:#fff">'+(n>1?n:'')+'</div>':''
      ).join('')+'</div></div></div>'+
      // Members grid
      '<div style="font-size:.7rem;font-weight:700;text-transform:uppercase;letter-spacing:1px;color:var(--text3);margin:4px 0 10px;padding-bottom:6px;border-bottom:1px solid var(--border)">All Members — '+members.length+'</div>'+
      '<div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(200px,1fr));gap:9px">' +
      (memberStats.length ? memberStats.map(s=>{
        const pc2=s.prod>=75?'var(--green)':s.prod>=50?'var(--amber)':'var(--red)';
        const stCfg=STATUS_CFG[s.todaySt]||{};
        return '<div style="background:var(--bg2);border:1px solid var(--border);border-radius:10px;padding:12px 14px;display:flex;align-items:center;gap:10px;transition:all .15s" onmouseover="this.style.borderColor=\'var(--accent)\'" onmouseout="this.style.borderColor=\'var(--border)\'">' +
          '<div style="width:34px;height:34px;border-radius:9px;background:var(--bg4);color:var(--text2);display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:700;flex-shrink:0">'+s.ini+'</div>' +
          '<div style="flex:1;min-width:0">' +
          '<div style="font-size:.83rem;font-weight:600;color:var(--text);overflow:hidden;text-overflow:ellipsis;white-space:nowrap">'+s.m+'</div>' +
          '<div style="font-size:.72rem;color:var(--text3)">'+s.count+' entries · <span style="color:'+pc2+';font-weight:700">'+s.prod.toFixed(0)+'%</span>' + (s.shift?' · '+s.shift:'') + '</div>' +
          '</div>' +
          (s.todaySt?'<span class="sp sp-'+s.todaySt.toLowerCase().replace(' ','-')+'" style="font-size:.65rem">'+s.todaySt+'</span>':'<span class="sp sp-ns" style="font-size:.65rem">—</span>') +
          '</div>';
      }).join('') : '<div class="empty"><p>No members yet. Sync from SharePoint.</p></div>') +
      '</div>';
  }

  function renderAdminTeamTracker() {
    const el = q('#view-teamtrack'); if (!el) return;
    const typeF = q('#adm-tk-type')?.value || '';
    const taskF = q('#adm-tk-task')?.value || '';
    const sortF = q('#adm-tk-sort')?.value || 'date';
    const subs  = Array.isArray(submissions) ? submissions : [];

    // Multi-select filter: if any selected, filter to those associates
    let rows = subs
      .filter(s => admSelectedAssocs.size === 0 || admSelectedAssocs.has((s.employeeName||'').toLowerCase()))
      .filter(s => !typeF || s.workType===typeF)
      .filter(s => !taskF || s.taskType===taskF);
    if (sortF==='hours')     rows.sort((a,b)=>parseFloat(b.hours||0)-parseFloat(a.hours||0));
    else if (sortF==='name') rows.sort((a,b)=>(a.employeeName||'').localeCompare(b.employeeName||''));
    else                     rows.sort((a,b)=>(b.date||'').localeCompare(a.date||''));
    const totalH = rows.reduce((a,r)=>a+parseFloat(r.hours||0),0);
    const nptH   = rows.reduce((a,r)=>a+parseFloat(r.npt||0),0);

    // Build selected chips
    const selectedChips = admSelectedAssocs.size > 0
      ? [...admSelectedAssocs].map(n =>
          '<span style="display:inline-flex;align-items:center;gap:5px;padding:3px 10px 3px 12px;border-radius:20px;background:var(--accent);color:#fff;font-size:.78rem;font-weight:700">' +
          n + '<button data-action="adm-assoc-remove" data-val="' + n + '" style="background:none;border:none;color:#fff;cursor:pointer;font-size:11px;padding:0;line-height:1;opacity:.8">✕</button></span>'
        ).join('')
      : '';

    // Associate picker chips — all associates as toggleable pills
    const assocChips = ASSOCIATES.map(m => {
      const ml  = m.toLowerCase();
      const sel = admSelectedAssocs.has(ml);
      return '<button data-action="adm-assoc-toggle" data-val="' + ml + '" style="padding:4px 11px;border-radius:20px;border:1px solid ' +
        (sel ? 'var(--accent)' : 'var(--border)') + ';background:' +
        (sel ? 'var(--accent)' : 'var(--bg3)') + ';color:' +
        (sel ? '#fff' : 'var(--text2)') + ';font-size:.76rem;font-weight:' +
        (sel ? '700' : '500') + ';cursor:pointer;font-family:inherit;white-space:nowrap">' + m + '</button>';
    }).join('');

    el.innerHTML =
      '<div class="ph"><div class="ph-left"><div class="ph-title">Team Tracker</div>' +
      '<div class="ph-sub">' + rows.length + ' entries · ' +
      (admSelectedAssocs.size > 0 ? admSelectedAssocs.size + ' associate(s) selected' : 'all ' + getAllMembers().length + ' members') +
      '</div></div>' +
      '<div class="ph-actions"><button class="btn btn-ghost btn-sm" data-action="adm-add-leave" style="margin-right:6px">+ Add Leave</button><button class="btn btn-ghost btn-sm" data-action="adm-sync-tasks">' + ic.sync + ' Sync</button></div></div>' +

      '<div class="stats-grid sg4" style="margin-bottom:12px">' +
      '<div class="stat-card ab"><div class="lbl">Entries</div><div class="val">' + rows.length + '</div></div>' +
      '<div class="stat-card gb"><div class="lbl">Total Hours</div><div class="val">' + totalH.toFixed(1) + 'h</div></div>' +
      '<div class="stat-card amb"><div class="lbl">NPT Hours</div><div class="val">' + nptH.toFixed(1) + 'h</div></div>' +
      '<div class="stat-card pb"><div class="lbl">Members</div><div class="val">' + getAllMembers().length + '</div></div></div>' +

      // Associate multi-select chip picker
      '<div style="background:var(--bg2);border:1px solid var(--border);border-radius:10px;padding:10px 12px;margin-bottom:10px">' +
      '<div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:8px">' +
      '<span style="font-size:.76rem;font-weight:700;text-transform:uppercase;letter-spacing:.8px;color:var(--text3)">Filter by Associate</span>' +
      (admSelectedAssocs.size > 0 ? '<button data-action="adm-assoc-clear" style="font-size:.72rem;color:var(--red);background:none;border:none;cursor:pointer;font-family:inherit;font-weight:600">✕ Clear All</button>' : '') +
      '</div>' +
      '<div style="display:flex;flex-wrap:wrap;gap:5px">' + assocChips + '</div>' +
      (admSelectedAssocs.size > 0 ? '<div style="margin-top:8px;display:flex;flex-wrap:wrap;gap:5px">' + selectedChips + '</div>' : '') +
      '</div>' +

      // Other filters row
      '<div style="display:flex;gap:8px;margin-bottom:12px;flex-wrap:wrap;align-items:center">' +
      '<select id="adm-tk-type" class="dtr-select" style="max-width:150px"><option value="">All Work Types</option><option value="Productive"' + (typeF==='Productive'?' selected':'') + '>Productive</option><option value="NPT"' + (typeF==='NPT'?' selected':'') + '>NPT</option></select>' +
      '<select id="adm-tk-task" class="dtr-select" style="max-width:170px"><option value="">All Task Types</option>' + TASK_TYPES.map(t=>'<option value="'+t+'"'+(taskF===t?' selected':'')+'>'+t+'</option>').join('') + '</select>' +
      '<select id="adm-tk-sort" class="dtr-select" style="max-width:130px"><option value="date"' + (sortF==='date'?' selected':'') + '>Latest</option><option value="name"' + (sortF==='name'?' selected':'') + '>By Name</option><option value="hours"' + (sortF==='hours'?' selected':'') + '>By Hours</option></select>' +
      '</div>' +
      '<div class="tbl-wrap"><table class="dtr-table"><thead><tr>' +
      '<th>#</th><th>Date</th><th>Employee</th><th>Task Type</th><th>Work Type</th><th>Hours</th><th>NPT Hrs</th><th>Shift</th><th>Notes</th><th>Submitted</th><th>Actions</th>' +
      '</tr></thead><tbody>' +
      (rows.length ? rows.map((r,i)=>{
        // Store real submissions index directly — indexOf is reliable here since r is object reference
        const realIdx = submissions.indexOf(r);
        const useIdx  = realIdx >= 0 ? realIdx : i;
        const editSVG = '<svg width="13" height="13" viewBox="0 0 24 24" fill="none"><path d="M11 4H4a2 2 0 00-2 2v14a2 2 0 002 2h14a2 2 0 002-2v-7" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><path d="M18.5 2.5a2.121 2.121 0 013 3L12 15l-4 1 1-4 9.5-9.5z" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>';
        const delSVG  = '<svg width="13" height="13" viewBox="0 0 24 24" fill="none"><path d="M3 6h18M8 6V4h8v2M19 6l-1 14H6L5 6" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>';
        return '<tr>'+
          '<td style="color:var(--text3);font-size:.8rem">'+(i+1)+'</td>'+
          '<td style="font-weight:600">'+r.date+'</td><td class="bold">'+r.employeeName+'</td>' +
          '<td style="font-weight:600">'+r.taskType+'</td>' +
          '<td><span class="badge '+(r.workType==='NPT'?'ba':'bg2')+'">'+r.workType+'</span></td>' +
          '<td class="mono">'+parseFloat(r.hours||0).toFixed(1)+'</td>' +
          '<td class="mono"'+(parseFloat(r.npt||0)>0?' style="color:var(--amber)"':'')+'>'+parseFloat(r.npt||0).toFixed(1)+'</td>' +
          '<td>'+(r.shift||'—')+'</td>' +
          '<td style="max-width:160px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="'+(r.adhoc||'').replace(/"/g,'&quot;')+'">'+(r.adhoc||'—')+'</td>' +
          '<td style="font-size:.78rem;color:var(--text3)">'+formatSPTime(r.submittedAt)+'</td>' +
          '<td><div style="display:flex;gap:5px">'+
            '<button data-action="adm-edit-entry" data-idx="'+useIdx+'" data-val="'+useIdx+'" title="Edit" style="width:28px;height:28px;border-radius:7px;border:1px solid var(--border);background:var(--bg3);color:var(--text);cursor:pointer;display:flex;align-items:center;justify-content:center">'+editSVG+'</button>'+
            '<button data-action="adm-del-entry" data-idx="'+useIdx+'" data-val="'+useIdx+'" title="Delete" style="width:28px;height:28px;border-radius:7px;border:1.5px solid rgba(255,69,58,.4);background:rgba(255,69,58,.1);color:#ff453a;cursor:pointer;display:flex;align-items:center;justify-content:center">'+delSVG+'</button>'+
          '</div></td>'+
        '</tr>';
      }).join('') :
        '<tr><td colspan="11"><div class="empty"><p>No entries. Sync from SP or check filters.</p></div></td></tr>') +
      '</tbody></table></div>';

    // Wire remaining filter dropdowns (type/task/sort)
    ['#adm-tk-type','#adm-tk-task','#adm-tk-sort'].forEach(sel => {
      const inp = el.querySelector(sel);
      if (inp) inp.addEventListener('change', () => renderAdminTeamTracker());
    });
  }

  // ── SP helpers for admin edit/delete ──────────────────────────────────
  async function spDeleteItem(listName, spId) {
    if (!spId) return false;
    const token = await getDigest();
    if (!token) return false;
    return new Promise(resolve => {
      GM_xmlhttpRequest({
        method:'POST', withCredentials:true,
        url: SP.SITE+"/_api/web/lists/GetByTitle('"+listName+"')/items("+spId+")" ,
        headers:{'Accept':'application/json;odata=verbose','Content-Type':'application/json;odata=verbose',
          'X-RequestDigest':token,'X-HTTP-Method':'DELETE','If-Match':'*'},
        onload:  r => resolve(r.status===200||r.status===204),
        onerror: ()=> resolve(false)
      });
    });
  }

  async function spPatchItem(listName, spId, metaType, body) {
    if (!spId) return false;
    const token = await getDigest();
    if (!token) return false;
    const data = JSON.stringify({__metadata:{type:metaType||'SP.Data.TaskReportListItem'},...body});
    return new Promise(resolve => {
      GM_xmlhttpRequest({
        method:'POST', withCredentials:true, data,
        url: SP.SITE+"/_api/web/lists/GetByTitle('"+listName+"')/items("+spId+")" ,
        headers:{'Accept':'application/json;odata=verbose','Content-Type':'application/json;odata=verbose',
          'X-RequestDigest':token,'X-HTTP-Method':'MERGE','If-Match':'*'},
        onload:  r => resolve(r.status===200||r.status===204),
        onerror: ()=> resolve(false)
      });
    });
  }

  function showAdminEditEntry(idx) {
    const r = submissions[idx]; if (!r) return;
    const isLt=root&&root.classList.contains('lt');
    const popBg=isLt?'#fff':'#12151f',rowBg=isLt?'#f5f5f7':'#1c2030';
    const bord=isLt?'rgba(0,0,0,.08)':'rgba(255,255,255,.1)';
    const textC=isLt?'#1c1c1e':'#f0f4ff',t2C=isLt?'rgba(28,28,30,.5)':'rgba(240,244,255,.5)';
    const inp = (id,type,val,extra='')=>'<input id="ae-'+id+'" type="'+type+'" value="'+val+'" '+extra+' style="width:100%;padding:8px 10px;border-radius:8px;border:1px solid '+bord+';background:'+rowBg+';color:'+textC+';font-family:inherit;font-size:.88rem;box-sizing:border-box">';
    const sel = (id,opts,val)=>'<select id="ae-'+id+'" style="width:100%;padding:8px 10px;border-radius:8px;border:1px solid '+bord+';background:'+rowBg+';color:'+textC+';font-family:inherit;font-size:.88rem">'+opts.map(o=>'<option value="'+o+'"'+(o===val?' selected':'')+'>'+o+'</option>').join('')+'</select>';
    const lbl = (l,el)=>'<div style="margin-bottom:10px"><label style="font-size:.68rem;font-weight:700;text-transform:uppercase;letter-spacing:.6px;color:'+t2C+';display:block;margin-bottom:4px">'+l+'</label>'+el+'</div>';
    const ov=document.createElement('div');
    ov.style.cssText='position:fixed;inset:0;background:rgba(0,0,0,.55);z-index:2147483646;display:flex;align-items:center;justify-content:center;backdrop-filter:blur(8px)';
    ov.innerHTML='<div style="background:'+popBg+';border:1px solid '+bord+';border-radius:18px;width:min(480px,94vw);max-height:90vh;overflow-y:auto;box-shadow:0 24px 70px rgba(0,0,0,.8);font-family:-apple-system,sans-serif">'
      +'<div style="height:3px;background:linear-gradient(90deg,#4f9eff,#a371f7)"></div>'
      +'<div style="padding:22px">'
      +'<div style="font-size:.95rem;font-weight:800;color:'+textC+';margin-bottom:16px">Edit Entry'+(r.spId?' <span style="font-size:.7rem;padding:2px 6px;border-radius:4px;background:rgba(79,158,255,.15);color:#4f9eff">SP #'+r.spId+'</span>':'')+'</div>'
      +'<div style="display:grid;grid-template-columns:1fr 1fr;gap:10px">'
      +lbl('Employee','<input value="'+(r.employeeName||'')+'" readonly style="width:100%;padding:8px 10px;border-radius:8px;border:1px solid '+bord+';background:'+rowBg+';color:'+t2C+';font-family:inherit;opacity:.7;box-sizing:border-box">')
      +lbl('Date', inp('date','date',r.date||''))
      +'</div>'
      +'<div style="display:grid;grid-template-columns:1fr 1fr;gap:10px">'
      +lbl('Task Type', sel('task', TASK_TYPES.concat(['Leave']), r.taskType||''))
      +lbl('Work Type', sel('work', ['Productive','NPT'], r.workType||'Productive'))
      +'</div>'
      +'<div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px">'
      +lbl('Hours', inp('hrs','number',(r.hours||0),'step="0.1" min="0" max="8"'))
      +lbl('NPT Hrs', inp('npt','number',(r.npt||0),'step="0.1" min="0" max="8"'))
      +lbl('Shift', sel('shift',['8-5','9-6','10-7','11-8'],r.shift||'8-5'))
      +'</div>'
      +lbl('Notes', '<input id="ae-notes" value="'+(r.adhoc||'').replace(/"/g,'&quot;')+'" style="width:100%;padding:8px 10px;border-radius:8px;border:1px solid '+bord+';background:'+rowBg+';color:'+textC+';font-family:inherit;font-size:.88rem;box-sizing:border-box">')
      +'<div style="display:flex;gap:10px;margin-top:6px">'
      +'<button id="ae-cancel" style="flex:1;padding:11px;border-radius:10px;border:1px solid '+bord+';background:'+rowBg+';color:'+t2C+';font-weight:600;cursor:pointer;font-family:inherit">Cancel</button>'
      +'<button id="ae-save" style="flex:2;padding:11px;border-radius:10px;border:none;background:var(--accent);color:#fff;font-weight:700;cursor:pointer;font-family:inherit">Save Changes</button>'
      +'</div></div></div>';
    const rr=document.getElementById('dtr-root-outer')||document.body;
    rr.appendChild(ov);
    ov.querySelector('#ae-cancel').onclick=()=>ov.remove();
    ov.querySelector('#ae-save').onclick=async()=>{
      const updated={...r,
        date:ov.querySelector('#ae-date').value||r.date,
        taskType:ov.querySelector('#ae-task').value||r.taskType,
        workType:ov.querySelector('#ae-work').value||r.workType,
        hours:parseFloat(ov.querySelector('#ae-hrs').value)||0,
        npt:parseFloat(ov.querySelector('#ae-npt').value)||0,
        shift:ov.querySelector('#ae-shift').value||r.shift,
        adhoc:ov.querySelector('#ae-notes').value,
      };
      submissions[idx]=updated;
      // Save locally
      const uk='dtr_subs_'+(updated.employeeName||authState.name).trim().toLowerCase().replace(/[^a-z0-9]+/g,'_');
      const allSubs=safeLoad(uk,[]);
      const si=allSubs.findIndex(s=>s.submittedAt===r.submittedAt&&s.taskType===r.taskType&&s.date===r.date);
      if(si>=0)allSubs[si]=updated; else allSubs.push(updated);
      safeSave(uk,allSubs);
      ov.remove();
      renderAdminTeamTracker();
      // Sync to SP if has spId
      if(updated.spId){
        toast('Updating SharePoint...','info');
        const listType=(typeof getListType==='function')?await getListType(SP.TASK_LIST):'SP.Data.TaskReportListItem';
        const ok=await spPatchItem(SP.TASK_LIST,updated.spId,listType,{
          TaskType:updated.taskType||'',HoursWorked:updated.hours||0,NPTHours:updated.npt||0,
          WorkType:updated.workType||'Productive',AdHocDetails:updated.adhoc||'',
          Shift:updated.shift||'',TaskDate:localDateSP(updated.date||todayStr())
        });
        toast(ok?'✅ Entry updated in SharePoint':'⚠️ Updated locally — SP sync failed',ok?'ok':'info');
      } else {
        toast('✅ Entry updated locally','ok');
      }
    };
  }

  function showAdminDeleteEntry(idx) {
    const r = submissions[idx]; if (!r) return;
    const isLt=root&&root.classList.contains('lt');
    const popBg=isLt?'#fff':'#12151f',bord=isLt?'rgba(0,0,0,.08)':'rgba(255,255,255,.1)';
    const textC=isLt?'#1c1c1e':'#f0f4ff',t2C=isLt?'rgba(28,28,30,.5)':'rgba(240,244,255,.5)';
    const ov=document.createElement('div');
    ov.style.cssText='position:fixed;inset:0;background:rgba(0,0,0,.55);z-index:2147483646;display:flex;align-items:center;justify-content:center;backdrop-filter:blur(8px)';
    ov.innerHTML='<div style="background:'+popBg+';border:1px solid '+bord+';border-radius:18px;width:min(380px,92vw);overflow:hidden;box-shadow:0 24px 70px rgba(0,0,0,.8);font-family:-apple-system,sans-serif">'
      +'<div style="height:3px;background:#ef4444"></div>'
      +'<div style="padding:24px">'
      +'<div style="font-size:.95rem;font-weight:800;color:'+textC+';margin-bottom:8px">Delete Entry?</div>'
      +'<div style="background:rgba(239,68,68,.08);border:1px solid rgba(239,68,68,.2);border-radius:10px;padding:12px;margin-bottom:16px;font-size:.85rem;color:'+textC+'">'
      +'<b>'+r.employeeName+'</b> — '+r.taskType+' on '+r.date
      +(r.spId?'<div style="font-size:.72rem;color:#ef4444;margin-top:4px">SP #'+r.spId+' — will be deleted from SharePoint</div>':'')
      +'</div>'
      +'<div style="display:flex;gap:10px">'
      +'<button id="ad-cancel" style="flex:1;padding:11px;border-radius:10px;border:1px solid '+bord+';background:transparent;color:'+t2C+';font-weight:600;cursor:pointer;font-family:inherit">Cancel</button>'
      +'<button id="ad-confirm" style="flex:1;padding:11px;border-radius:10px;border:none;background:#ef4444;color:#fff;font-weight:700;cursor:pointer;font-family:inherit">Delete</button>'
      +'</div></div></div>';
    const rr=document.getElementById('dtr-root-outer')||document.body;
    rr.appendChild(ov);
    ov.querySelector('#ad-cancel').onclick=()=>ov.remove();
    ov.querySelector('#ad-confirm').onclick=async()=>{
      const spId=r.spId;
      // Remove from local submissions
      submissions.splice(idx,1);
      // Remove from local storage too
      const uk='dtr_subs_'+(r.employeeName||authState.name).trim().toLowerCase().replace(/[^a-z0-9]+/g,'_');
      const allSubs=safeLoad(uk,[]).filter(s=>!(s.submittedAt===r.submittedAt&&s.taskType===r.taskType&&s.date===r.date));
      safeSave(uk,allSubs);
      ov.remove();
      renderAdminTeamTracker();
      if(spId){
        toast('Deleting from SharePoint...','info');
        const ok=await spDeleteItem(SP.TASK_LIST,spId);
        toast(ok?'✅ Deleted from SharePoint':'⚠️ Deleted locally — SP sync failed',ok?'ok':'info');
      } else {
        toast('✅ Entry removed','ok');
      }
    };
  }

  function showAdminAddLeave() {
    const isLt=root&&root.classList.contains('lt');
    const popBg=isLt?'#fff':'#12151f',rowBg=isLt?'#f5f5f7':'#1c2030';
    const bord=isLt?'rgba(0,0,0,.08)':'rgba(255,255,255,.1)';
    const textC=isLt?'#1c1c1e':'#f0f4ff',t2C=isLt?'rgba(28,28,30,.5)':'rgba(240,244,255,.5)';
    const allM=getAllMembers();
    const ov=document.createElement('div');
    ov.style.cssText='position:fixed;inset:0;background:rgba(0,0,0,.55);z-index:2147483646;display:flex;align-items:center;justify-content:center;backdrop-filter:blur(8px)';
    ov.innerHTML='<div style="background:'+popBg+';border:1px solid '+bord+';border-radius:18px;width:min(420px,94vw);overflow:hidden;box-shadow:0 24px 70px rgba(0,0,0,.8);font-family:-apple-system,sans-serif">'
      +'<div style="height:3px;background:linear-gradient(90deg,#a371f7,#4f9eff)"></div>'
      +'<div style="padding:22px">'
      +'<div style="font-size:.95rem;font-weight:800;color:'+textC+';margin-bottom:16px">Add Leave for Associate</div>'
      +'<div style="margin-bottom:10px"><label style="font-size:.7rem;font-weight:700;text-transform:uppercase;letter-spacing:.6px;color:'+t2C+';display:block;margin-bottom:4px">Associate</label>'
      +'<select id="al-name" style="width:100%;padding:9px 10px;border-radius:8px;border:1px solid '+bord+';background:'+rowBg+';color:'+textC+';font-family:inherit">'+allM.map(m=>'<option value="'+m+'">'+m+'</option>').join('')+'</select></div>'
      +'<div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:10px">'
      +'<div><label style="font-size:.7rem;font-weight:700;text-transform:uppercase;letter-spacing:.6px;color:'+t2C+';display:block;margin-bottom:4px">Date</label><input id="al-date" type="date" value="'+todayStr()+'" style="width:100%;padding:9px 10px;border-radius:8px;border:1px solid '+bord+';background:'+rowBg+';color:'+textC+';font-family:inherit;box-sizing:border-box"></div>'
      +'<div><label style="font-size:.7rem;font-weight:700;text-transform:uppercase;letter-spacing:.6px;color:'+t2C+';display:block;margin-bottom:4px">Leave Type</label>'
      +'<select id="al-type" style="width:100%;padding:9px 10px;border-radius:8px;border:1px solid '+bord+';background:'+rowBg+';color:'+textC+';font-family:inherit"><option value="SL">SL – Sick Leave</option><option value="CL">CL – Casual Leave</option><option value="AL">AL – Annual Leave</option></select></div>'
      +'</div>'
      +'<div style="margin-bottom:14px"><label style="font-size:.7rem;font-weight:700;text-transform:uppercase;letter-spacing:.6px;color:'+t2C+';display:block;margin-bottom:4px">Duration</label>'
      +'<div style="display:flex;gap:8px">'
      +'<button class="al-dur-btn active" data-dur="full" style="flex:1;padding:9px;border-radius:8px;border:1.5px solid var(--accent);background:var(--accent);color:#fff;font-weight:700;cursor:pointer;font-family:inherit;font-size:.85rem">Full Day (480 min)</button>'
      +'<button class="al-dur-btn" data-dur="half" style="flex:1;padding:9px;border-radius:8px;border:1.5px solid var(--border);background:transparent;color:'+textC+';font-weight:600;cursor:pointer;font-family:inherit;font-size:.85rem">Half Day (240 min)</button>'
      +'</div></div>'
      +'<div style="display:flex;gap:10px">'
      +'<button id="al-cancel" style="flex:1;padding:11px;border-radius:10px;border:1px solid '+bord+';background:'+rowBg+';color:'+t2C+';font-weight:600;cursor:pointer;font-family:inherit">Cancel</button>'
      +'<button id="al-save" style="flex:2;padding:11px;border-radius:10px;border:none;background:linear-gradient(135deg,#a371f7,#4f9eff);color:#fff;font-weight:700;cursor:pointer;font-family:inherit">Add Leave to SharePoint</button>'
      +'</div></div></div>';
    const rr=document.getElementById('dtr-root-outer')||document.body;
    rr.appendChild(ov);
    // Duration toggle
    let selDur='full';
    ov.querySelectorAll('.al-dur-btn').forEach(btn=>{
      btn.addEventListener('click',()=>{
        selDur=btn.dataset.dur;
        ov.querySelectorAll('.al-dur-btn').forEach(b=>{
          b.style.borderColor=b===btn?'var(--accent)':'var(--border)';
          b.style.background=b===btn?'var(--accent)':'transparent';
          b.style.color=b===btn?'#fff':textC;
          b.style.fontWeight=b===btn?'700':'600';
        });
      });
    });
    ov.querySelector('#al-cancel').onclick=()=>ov.remove();
    ov.querySelector('#al-save').onclick=async()=>{
      const name=ov.querySelector('#al-name').value;
      const date=ov.querySelector('#al-date').value;
      const leaveKind=ov.querySelector('#al-type').value;
      const leaveMins=selDur==='full'?480:240;
      const leaveHrs=parseFloat((leaveMins/60).toFixed(2));
      if(!name||!date){toast('Please fill all fields','info');return;}
      const task={employeeName:name,taskType:'Leave',hours:0,npt:leaveHrs,
        nptMinutes:leaveMins,minutes:0,workType:'NPT',leaveType:selDur,
        adhoc:leaveKind+' – Admin set',shift:'8-5',date,submittedAt:nowUTC()};
      // Save to associate's local storage
      const uk='dtr_subs_'+name.trim().toLowerCase().replace(/[^a-z0-9]+/g,'_');
      const assocSubs=safeLoad(uk,[]);
      assocSubs.push(task);
      safeSave(uk,assocSubs);
      // Update attendance status cache
      teamStatusCache_adm[name+'::'+date]={status:leaveKind,shift:'8-5'};
      teamStatusCache[name+'::'+date]={status:leaveKind,shift:'8-5'};
      safeSaveObj('dtr_teamcache',teamStatusCache_adm);
      ov.remove();
      toast('Saving leave for '+name+'...','info');
      // Post task to SP TaskReport list
      const ok1=await postTask(task);
      // Post attendance to SP AttendanceLog
      const ok2=await postAttendance({name,status:leaveKind,shift:'8-5',date,updatedAt:nowUTC()});
      toast((ok1||ok2)?'✅ Leave added for '+name+' on '+date+' in SharePoint':'⚠️ Leave saved locally — SP sync queued',(ok1||ok2)?'ok':'info');
      renderAdminWeekView();
      if(currentView==='overview')renderAdminOverview();
    };
  }


  function showAdminAttPicker(val) {
    const [name, dk] = val.split('|');
    if (!name || !dk) return;
    const existing   = getTeamStatus(name, dk);
    const existShift = getTeamShift(name, dk) || '';
    const isLt   = root && root.classList.contains('lt');
    const popBg  = isLt ? '#ffffff' : '#12151f';
    const rowBg  = isLt ? '#f5f5f7' : '#1c2030';
    const bord   = isLt ? 'rgba(0,0,0,.08)' : 'rgba(255,255,255,.1)';
    const textC  = isLt ? '#1c1c1e' : '#f0f4ff';
    const text2C = isLt ? 'rgba(28,28,30,.5)' : 'rgba(240,244,255,.5)';
    const shadow = isLt ? '0 8px 32px rgba(0,0,0,.15)' : '0 24px 70px rgba(0,0,0,.9)';
    const dateDisp = new Date(dk+'T12:00:00').toLocaleDateString('en-IN',{weekday:'long',day:'numeric',month:'long',year:'numeric'});

    const STATUSES_ADMIN = [
      {k:'WFO',  label:'Work From Office', clr:'#30d158', bg:'rgba(48,209,88,.12)'},
      {k:'WFH',  label:'Work From Home',   clr:'#0a84ff', bg:'rgba(10,132,255,.12)'},
      {k:'SL',   label:'Sick Leave',       clr:'#ff9f0a', bg:'rgba(255,159,10,.12)'},
      {k:'CL',   label:'Casual Leave',     clr:'#a371f7', bg:'rgba(163,113,247,.12)'},
      {k:'AL',   label:'Annual Leave',     clr:'#ff453a', bg:'rgba(255,69,58,.12)'},
      {k:'Optional Off', label:'Optional Off', clr:'#636366', bg:'rgba(99,99,102,.1)'},
      {k:'',     label:'Clear / Not Set',  clr:'#636366', bg:'rgba(99,99,102,.08)'},
    ];

    const ov = document.createElement('div');
    ov.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:2147483646;display:flex;align-items:center;justify-content:center;backdrop-filter:blur(8px)';
    ov.innerHTML =
      '<div style="background:'+popBg+';border:1px solid '+bord+';border-radius:18px;width:min(380px,92vw);box-shadow:'+shadow+';font-family:-apple-system,sans-serif;overflow:hidden">' +
        '<div style="height:3px;background:linear-gradient(90deg,#4f9eff,#a371f7)"></div>' +
        '<div style="padding:22px 22px 18px">' +
          '<div style="margin-bottom:16px">' +
            '<div style="font-size:.95rem;font-weight:800;color:'+textC+'">Edit Attendance</div>' +
            '<div style="font-size:.8rem;color:'+text2C+';margin-top:3px">'+name+' · '+dateDisp+'</div>' +
          '</div>' +
          '<div style="margin-bottom:14px">' +
            STATUSES_ADMIN.map((s,si) =>
              '<button type="button" class="adm-st-btn" data-si="'+si+'" style="width:100%;display:flex;align-items:center;gap:10px;padding:10px 12px;border-radius:10px;border:1.5px solid '+(existing===s.k?s.clr:'var(--border)')+';background:'+(existing===s.k?s.bg:'transparent')+';color:'+textC+';font-family:inherit;cursor:pointer;margin-bottom:6px;text-align:left;transition:all .12s">' +
                '<span style="width:10px;height:10px;border-radius:50%;background:'+s.clr+';flex-shrink:0;display:inline-block"></span>' +
                '<div style="flex:1"><div style="font-size:.85rem;font-weight:'+(existing===s.k?'800':'600')+'">'+(s.k||'Clear')+'</div><div style="font-size:.7rem;color:'+text2C+'">'+s.label+'</div></div>' +
                (existing===s.k?'<span style="margin-left:auto;font-size:.75rem;color:'+s.clr+'">✓ Current</span>':'') +
              '</button>'
            ).join('') +
          '</div>' +
          // Shift selector
          '<div style="margin-bottom:16px">' +
            '<div style="font-size:.72rem;font-weight:700;text-transform:uppercase;letter-spacing:.6px;color:'+text2C+';margin-bottom:8px">Shift</div>' +
            '<div style="display:flex;gap:6px;flex-wrap:wrap">' +
            ['8-5','9-6','10-7','11-8'].map(sh =>
              '<button id="adm-sh-'+sh+'" style="padding:6px 14px;border-radius:8px;border:1.5px solid '+(existShift===sh?'var(--accent)':'var(--border)')+';background:'+(existShift===sh?'var(--accent)':'transparent')+';color:'+(existShift===sh?'#fff':textC)+';font-family:inherit;cursor:pointer;font-size:.82rem;font-weight:600;transition:all .12s">'+sh+'</button>'
            ).join('') +
          '</div></div>' +
          '<button id="adm-att-cancel" style="width:100%;padding:11px;border-radius:10px;border:1px solid '+bord+';background:'+rowBg+';color:'+text2C+';font-size:.9rem;font-weight:600;cursor:pointer;font-family:inherit">Cancel</button>' +
        '</div>' +
      '</div>';

    const rr = document.getElementById('dtr-root-outer') || document.body;
    rr.appendChild(ov);
    ov.querySelector('#adm-att-cancel').addEventListener('click', e => { e.stopPropagation(); ov.remove(); });

    // Shift buttons
    let selShift = existShift;
    const shiftBtns = ['8-5','9-6','10-7','11-8'];
    shiftBtns.forEach(sh => {
      const btn = ov.querySelector('#adm-sh-'+sh);
      if (!btn) return;
      btn.addEventListener('click', e => {
        e.stopPropagation();
        selShift = sh;
        shiftBtns.forEach(s2 => {
          const b2 = ov.querySelector('#adm-sh-'+s2);
          if (b2) {
            b2.style.borderColor = s2===sh ? 'var(--accent)' : 'var(--border)';
            b2.style.background  = s2===sh ? 'var(--accent)' : 'transparent';
            b2.style.color       = s2===sh ? '#fff' : textC;
          }
        });
      });
    });

    // Status buttons — direct onclick via class
    ov.querySelectorAll('.adm-st-btn').forEach((btn, idx) => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const s = STATUSES_ADMIN[parseInt(btn.dataset.si)];
        if (!s) return;
        adminSetAttendance(name, dk, s.k, selShift);
        ov.remove();
      });
    });
  }

  async function adminSetAttendance(name, dk, status, shift) {
    if (!name || !dk) return;
    shift = shift || getTeamShift(name, dk) || '8-5';
    const cacheKey = name + '::' + dk;

    // ── 1. Update ALL caches immediately so UI reflects right away ────────
    if (status) {
      teamStatusCache_adm[cacheKey] = { status, shift };
      teamStatusCache[cacheKey]     = { status, shift };
    } else {
      delete teamStatusCache_adm[cacheKey];
      delete teamStatusCache[cacheKey];
    }
    // Persist to GM storage
    safeSaveObj('dtr_teamcache', teamStatusCache_adm);
    // If admin is editing their own or current associate's record, update statusCache too
    if (name === authState.name || name === (authState.originalName||'')) {
      if (status) statusCache[dk] = { status, shift, date:dk, name, updatedAt:nowUTC() };
      else delete statusCache[dk];
      safeSaveObj('dtr_status2', statusCache);
    }

    // ── 2. Update attAllCache (overview stats) ────────────────────────────
    if (!Array.isArray(attAllCache)) attAllCache = [];
    const existIdx = attAllCache.findIndex(e =>
      (e.name||e.employeeName||'').toLowerCase() === name.toLowerCase() &&
      ((e.date||e.StatusDate||'').split('T')[0]) === dk
    );
    if (status) {
      const entry = { name, employeeName:name, date:dk, StatusDate:dk,
        status, WorkStatus:status, shift, Shift:shift, UpdatedAt:nowUTC() };
      if (existIdx >= 0) attAllCache[existIdx] = entry;
      else attAllCache.push(entry);
    } else {
      if (existIdx >= 0) attAllCache.splice(existIdx, 1);
    }

    // ── 3. Re-render week view + overview immediately ─────────────────────
    renderAdminWeekView();
    if (currentView === 'overview') renderAdminOverview();

    if (!status) {
      toast('🗑 Cleared ' + name + ' on ' + dk, 'info');
      return;
    }

    // ── 4. Post to SP (non-blocking — UI already updated) ─────────────────
    const ok = await postAttendance({ name, status, shift, date: dk, updatedAt: nowUTC() });
    toast(ok
      ? '✅ ' + name + ': ' + status + ' on ' + dk + ' saved to SharePoint'
      : '⚠️ ' + name + ': ' + status + ' saved locally — SP sync queued', ok ? 'ok' : 'info');
  }

  function renderAdminWeekView() {
    const el = q('#view-attweek'); if (!el) return;
    const ws     = getWeekStart(adminWeekOffset);
    const wLabel = getWeekLabel(adminWeekOffset);
    const today  = todayStr();
    const days   = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
    const members= getAllMembers();

    const dates = Array.from({length:7},(_,i)=>{
      const d=new Date(ws); d.setDate(ws.getDate()+i);
      const dk=d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0');
      return {dk,day:days[i],label:d.toLocaleDateString('en-US',{month:'short',day:'numeric'}),isToday:dk===today,isWeekend:i===0||i===6};
    });

    el.innerHTML =
      '<div class="ph"><div class="ph-left"><div class="ph-title">Week View</div>' +
      '<div class="ph-sub">'+wLabel+' — '+members.length+' members</div></div>' +
      '<div class="ph-actions"><button class="btn btn-ghost btn-sm" data-action="adm-sync-att">'+ic.sync+' Sync</button></div></div>' +
      '<div class="wv-controls" style="margin-bottom:16px">' +
      '<button class="wv-nav-btn" data-action="adm-wv-prev">'+ic.left+'</button>' +
      '<div class="wv-range"><div class="wv-range-title">'+wLabel+'</div><div class="wv-range-sub">'+(adminWeekOffset===0?'Current Week':Math.abs(adminWeekOffset)+' week(s) '+(adminWeekOffset<0?'ago':'ahead'))+'</div></div>'+
      (adminWeekOffset!==0?'<button class="wv-today-btn" data-action="adm-wv-today">This Week</button>':'')+
      '<button class="wv-nav-btn" data-action="adm-wv-next">'+ic.right+'</button></div>'+
      '<div class="wv-grid-wrap"><div class="wv-header">' +
      '<div class="wv-hcell" style="text-align:left;padding-left:14px">Member</div>' +
      dates.map(d=>'<div class="wv-hcell'+(d.isToday?' today-col':'')+'" style="'+(d.isWeekend?'opacity:.4':'')+'">'+d.day+'<div class="wv-hdate">'+d.label+'</div></div>').join('') +
      '</div>' +
      (members.length ? members.map(m=>{
        const ini=m.split(/[\s,]+/).filter(Boolean).map(x=>x[0].toUpperCase()).join('').slice(0,2);
        return '<div class="wv-row">'+
          '<div class="wv-name-cell"><div class="wv-av" style="background:var(--bg4);color:var(--text2)">'+ini+'</div><div class="wv-name">'+m+'</div></div>'+
          dates.map(d=>{
            const st=getTeamStatus(m,d.dk), sh=getTeamShift(m,d.dk)||'';
            const cfg=STATUS_CFG[st]||{};
            const cellKey2 = m+'|'+d.dk;
            return '<div class="wv-cell'+(d.isToday?' today-col':'')+(d.isWeekend?' weekend-col':'')+'" style="'+(st&&!d.isWeekend?'background:'+cfg.bg:'')+';'+(d.isWeekend?'':'cursor:pointer;') + '"' +
              (d.isWeekend?'':' data-action="adm-att-edit" data-val="'+m+'|'+d.dk+'" title="Edit '+m+' on '+d.dk+'"') + '>' +
              (d.isWeekend?'<span style="font-size:.7rem;color:var(--text3)">—</span>':
                st?'<span class="sp sp-'+st.toLowerCase().replace(' ','-')+'" style="font-size:.7rem">'+st+'</span>'+
                   (sh?'<span style="font-size:.65rem;color:var(--text3)">'+sh+'</span>':''):
                '<span style="font-size:.75rem;color:var(--text3);opacity:.4">+</span>') +
            '</div>';
          }).join('')+
          '</div>';
      }).join('') : '<div style="padding:24px;text-align:center;color:var(--text3)">No members. Sync first.</div>') +
      '</div>';
  }

  function renderAdminNPTLog() {
    const el = q('#view-attnpt'); if (!el) return;
    const nf  = (q('#adm-npt-name')?.value||'').toLowerCase();
    const all = Array.isArray(nptAllCache) ? nptAllCache : [];
    const rows= nf ? all.filter(n=>(n.name||n.employeeName||'').toLowerCase().includes(nf)) : all;
    const total=rows.reduce((a,n)=>a+parseInt(n.minutes||n.DurationMins||0),0);
    el.innerHTML=
      '<div class="ph"><div class="ph-left"><div class="ph-title">NPT Log</div>' +
      '<div class="ph-sub">All team non-productive time entries</div></div>' +
      '<div class="ph-actions"><button class="btn btn-ghost btn-sm" data-action="adm-sync-npt">'+ic.sync+' Sync</button></div></div>'+
      '<div class="stats-grid sg3" style="margin-bottom:12px">' +
      '<div class="stat-card ab"><div class="lbl">Entries</div><div class="val">'+rows.length+'</div></div>' +
      '<div class="stat-card amb"><div class="lbl">Total Minutes</div><div class="val">'+total+'m</div></div>' +
      '<div class="stat-card pb"><div class="lbl">People</div><div class="val">'+[...new Set(rows.map(n=>n.name||n.employeeName).filter(Boolean))].length+'</div></div></div>'+
      '<div style="display:flex;gap:8px;margin-bottom:12px">' +
      '<input id="adm-npt-name" class="dtr-input" placeholder="Filter by name..." style="max-width:200px" value="'+(nf||'')+'"></div>'+
      '<div class="tbl-wrap"><table class="dtr-table"><thead><tr>' +
      '<th>#</th><th>Name</th><th>Date</th><th>Type</th><th>Minutes</th><th>Description</th><th>Logged</th>' +
      '</tr></thead><tbody>'+
      (rows.length?rows.slice().reverse().map((n,i)=>{
        const name=n.name||n.employeeName||'', date=n.date||(n.NPTDate||'').split('T')[0], type=n.type||n.NPTType||'', mins=n.minutes||n.DurationMins||0, desc=n.desc||n.Description||'', logged=n.loggedAt||n.LoggedAt||n.Created||'';
        return '<tr><td style="color:var(--text3)">'+(i+1)+'</td><td class="bold">'+name+'</td><td>'+date+'</td>' +
          '<td><span style="padding:2px 7px;border-radius:4px;background:rgba(88,166,255,.1);color:var(--accent);font-size:.75rem;font-weight:700">'+type+'</span></td>' +
          '<td class="mono">'+mins+'m</td>' +
          '<td style="max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="'+desc.replace(/"/g,"&quot;")+'">'+desc+'</td>' +
          '<td style="font-size:.8rem;color:var(--text3)">'+formatSPTime(logged)+'</td></tr>';
      }).join(''):'<tr><td colspan="7"><div class="empty"><p>No NPT entries. Sync from SP.</p></div></td></tr>')+
      '</tbody></table></div>';
    const inp=el.querySelector('#adm-npt-name');
    if(inp) inp.addEventListener('input',()=>renderAdminNPTLog());
  }

  // Admin fetch functions
  function fetchAdminTasks() {
    toast('Syncing tasks from TaskReport...','info');
    GM_xmlhttpRequest({
      method: 'GET',
      url: SP.SITE + "/_api/web/lists/GetByTitle('" + SP.TASK_LIST + "')/items?$top=5000&$orderby=TaskDate desc",
      headers: { 'Accept': 'application/json;odata=verbose' },
      withCredentials: true,
      onload: res => {
        if (res.status < 200 || res.status >= 300) {
          let msg=''; try { msg=JSON.parse(res.responseText).error.message.value||''; } catch{}
          toast('❌ Task sync failed (' + res.status + '): ' + (msg||'check F12'), 'err');
          console.error('[WorkPulse] fetchAdminTasks:', res.status, res.responseText.slice(0,300));
          return;
        }
        try {
          const parsed = JSON.parse(res.responseText);
          const items  = (parsed.d && parsed.d.results) ? parsed.d.results : [];
          // Map exact column names from TaskReport list
          submissions = items.map(it => ({
            employeeName: it.EmployeeName || '',
            taskType:     it.TaskType     || '',
            hours:        parseFloat(it.HoursWorked) || 0,
            npt:          parseFloat(it.NPTHours)    || 0,
            workType:     it.WorkType     || 'Productive',
            adhoc:        it.AdHocDetails || '',
            shift:        it.Shift        || '',
            date:         (it.TaskDate || it.Created || '').split('T')[0],
            submittedAt:  it.SubmittedAt  || it.Created || ''
          }));
          toast('✅ ' + items.length + ' tasks synced from TaskReport', 'ok');
          console.log('[WorkPulse] fetchAdminTasks:', items.length, 'items');
          if (currentView === 'overview')  renderAdminOverview();
          if (currentView === 'teamtrack') renderAdminTeamTracker();
        } catch(e) {
          toast('❌ Task sync parse error: ' + e.message, 'err');
          console.error('[WorkPulse] fetchAdminTasks parse:', e, res.responseText.slice(0,300));
        }
      },
      onerror: () => toast('❌ Cannot reach SP for tasks', 'err')
    });
  }

  function fetchAdminAttendance() {
    toast('Syncing attendance from AttendanceLog...','info');
    GM_xmlhttpRequest({
      method: 'GET',
      url: SP.SITE + "/_api/web/lists/GetByTitle('" + SP.STATUS_LIST + "')/items?$top=5000&$orderby=StatusDate desc",
      headers: { 'Accept': 'application/json;odata=verbose' },
      withCredentials: true,
      onload: res => {
        if (res.status < 200 || res.status >= 300) {
          let msg=''; try{msg=JSON.parse(res.responseText).error.message.value||'';}catch{}
          toast('❌ Attendance sync failed (' + res.status + '): ' + (msg||'check F12'), 'err');
          console.error('[WorkPulse] fetchAdminAttendance:', res.status, res.responseText.slice(0,300));
          return;
        }
        try {
          const parsed = JSON.parse(res.responseText);
          if (!parsed.d) {
            toast('❌ Attendance list error — check list name: ' + SP.STATUS_LIST, 'err');
            return;
          }
          // Merge SP data into existing cache (don't nuke local edits)
          const spData = {};
          (parsed.d.results || []).forEach(it => {
            const name = it.EmployeeName || '';
            const dk   = (it.StatusDate  || '').split('T')[0];
            if (name && dk) {
              spData[name + '::' + dk] = {
                status: it.WorkStatus || '',
                shift:  it.Shift      || ''
              };
            }
          });
          // SP data takes precedence (it's authoritative), then local edits
          teamStatusCache_adm = Object.assign({}, teamStatusCache_adm, spData);
          teamStatusCache     = Object.assign({}, teamStatusCache, spData);
          safeSaveObj('dtr_teamcache', teamStatusCache_adm);
          const total = Object.keys(teamStatusCache_adm).length;
          toast('✅ Attendance synced — ' + total + ' records', 'ok');
          console.log('[WorkPulse] fetchAdminAttendance:', total, 'records');
          if (currentView === 'attweek')  renderAdminWeekView();
          if (currentView === 'overview') renderAdminOverview();
        } catch(e) {
          toast('❌ Attendance parse error: ' + e.message, 'err');
          console.error('[WorkPulse] fetchAdminAttendance parse:', e, res.responseText.slice(0,300));
        }
      },
      onerror: () => toast('❌ Cannot reach SP for attendance', 'err')
    });
  }

  function fetchAdminNPT() {
    toast('Syncing NPT...','info');
    GM_xmlhttpRequest({
      method:'GET',
      url:SP.SITE+"/_api/web/lists/GetByTitle('"+SP.NPT_LIST+"')/items?$top=5000&$orderby=Created desc",
      headers:{'Accept':'application/json;odata=verbose'}, withCredentials:true,
      onload:res=>{
        try {
          const adminNPTJson = JSON.parse(res.responseText);
          if (!adminNPTJson.d) { toast('❌ NPT list error — check list name: ' + SP.NPT_LIST,'err'); return; }
          nptAllCache=(adminNPTJson.d.results||[]).map(it=>({
            name:     it.EmployeeName||'',
            date:     (it.NPTDate||it.Created||'').split('T')[0],
            type:     it.NPTType||'',
            minutes:  parseInt(it.DurationMins)||0,
            desc:     it.Description||'',
            loggedAt: it.LoggedAt||it.Created||''
          }));
          toast('✅ NPT synced','ok');
          if(currentView==='attnpt') renderAdminNPTLog();
        } catch(e){ toast('❌ NPT sync error: '+e.message,'err'); }
      },
      onerror:()=>toast('❌ Cannot reach SP','err')
    });
  }

  function fetchAllAdminData() { fetchAdminTasks(); fetchAdminAttendance(); fetchAdminNPT(); }


  // DATA HELPERS
  function getAllUsersFromCache() { const tc=safeLoadObj('dtr_teamcache',{});if(Object.keys(tc).length)teamStatusCache=tc;const names=new Set();Object.keys(teamStatusCache).forEach(k=>{const[name]=k.split('::');if(name)names.add(name);});names.add(authState.name);return[...names].sort(); }
  function getTeamStatus(user,dk)  {
    const admVal = (teamStatusCache_adm[user+'::'+dk]||{}).status;
    if (admVal) return admVal;
    const teamVal = (teamStatusCache[user+'::'+dk]||{}).status;
    if (teamVal) return teamVal;
    // For own user, also check statusCache (associate's local attendance)
    if (user === authState.name) return (statusCache[dk]||{}).status || '';
    return '';
  }
  function getTeamProcess(user,dk) { return(teamStatusCache[user+'::'+dk]||{}).process||(user===authState.name?((statusCache[dk]||{}).process||''):''); }
  function getTeamTask(user,dk)    { return(teamStatusCache[user+'::'+dk]||{}).task   ||(user===authState.name?((statusCache[dk]||{}).task||''):''); }
  function mySubmissions() { return Array.isArray(submissions)?submissions.filter(s=>s&&s.employeeName===authState.name):[]; }
  function calcAvgProd(subs){ if(!Array.isArray(subs))return 0;const v=subs.filter(s=>s&&!s.taskType?.startsWith('Leave'));if(!v.length)return 0;return Math.min(100,v.reduce((a,s)=>a+((s.hours||0)/WH*100),0)/v.length); }
  function calcStreak(subs){ const days=[...new Set(subs.filter(s=>s&&!s.taskType?.startsWith('Leave')).map(s=>s.date))].sort().reverse();let streak=0,prev=new Date();for(const d of days){const dt=new Date(d+'T12:00:00'),diff=Math.round((prev-dt)/86400000);if(streak===0&&diff<=1){streak=1;prev=dt;}else if(diff===1){streak++;prev=dt;}else break;}return streak; }
  function getLast7(){ return Array.from({length:7},(_,i)=>{const d=new Date();d.setDate(d.getDate()-6+i);return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0');}); }
  function getWeekStart(offset) {
    const d = new Date();
    // Use local midnight to avoid UTC day boundary issues
    d.setHours(12,0,0,0); // noon = safe from DST and UTC shifts
    d.setDate(d.getDate() - d.getDay() + (offset * 7));
    d.setHours(0,0,0,0);
    return d;
  }
  function getWeekLabel(offset){ const ws=getWeekStart(offset),we=new Date(ws);we.setDate(ws.getDate()+6);const f=d=>d.toLocaleDateString('en-US',{month:'short',day:'numeric'});return f(ws)+' – '+f(we)+', '+we.getFullYear(); }
  // ── All date/time helpers use LOCAL timezone ───────────────────────────
  function todayStr() {
    const d = new Date();
    return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0');
  }
  // Local ISO timestamp (e.g. 2026-05-05T14:30:00+05:30) — for SP datetime fields
  // ── Date/Time helpers — all IST-aware ─────────────────────────────────
  //
  // Rule: SP stores UTC internally, displays per regional settings.
  // For date-only SP fields (NPTDate, TaskDate, StatusDate):
  //   → Send local date at local NOON converted to UTC
  //   → Noon in IST (UTC+5:30) = 06:30 UTC — safe from day-boundary shift
  // For timestamp SP fields (LoggedAt, SubmittedAt, UpdatedAt):
  //   → Send real UTC (new Date().toISOString()) — SP displays in site TZ
  // For UI display of timestamps from SP:
  //   → Convert to local using formatSPTime()

  // Convert a YYYY-MM-DD local date to UTC ISO for SP date fields
  // Uses local noon to avoid day-boundary issues with IST (+5:30)
  function localDateSP(dateStr) {
    const d = new Date((dateStr || todayStr()) + 'T12:00:00'); // local noon
    return d.toISOString(); // converts local noon → UTC (e.g. 06:30Z for IST)
  }

  // Current time as UTC ISO — correct for SP timestamp fields
  function nowUTC() {
    return new Date().toISOString();
  }

  // Format a SP datetime string for LOCAL display (IST)
  // SP returns UTC strings like "2026-05-05T05:53:00Z"
  // This converts them to local time for display
  function formatSPTime(spDateStr) {
    if (!spDateStr) return '—';
    try {
      const d = new Date(spDateStr);
      if (isNaN(d)) return spDateStr.replace('T',' ').slice(0,16);
      return d.toLocaleDateString('en-IN', {
        year:'numeric', month:'2-digit', day:'2-digit',
        hour:'2-digit', minute:'2-digit', hour12:false
      }).replace(/\//g,'-');
    } catch { return spDateStr.slice(0,16); }
  }

  // Keep localISOString as alias for backward compat (now uses UTC)
  function localISOString() { return new Date().toISOString(); }
  function yesterdayStr() {
    const d = new Date(); d.setDate(d.getDate()-1);
    return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0');
  }
  // Previous WORKING day: Mon→Fri, Tue-Sat→prev calendar day, Sun→Fri
  function prevWorkingDayStr() {
    const d   = new Date();
    const dow = d.getDay(); // 0=Sun,1=Mon,...,6=Sat
    // Monday → go back 3 days to Friday
    // Sunday → go back 2 days to Friday
    const back = dow === 1 ? 3 : dow === 0 ? 2 : 1;
    d.setDate(d.getDate() - back);
    return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0');
  }
  function prevWorkingDayLabel() {
    const dow = new Date().getDay();
    return dow === 1 ? 'Friday' : dow === 0 ? 'Friday' : 'Yesterday';
  }
  function isAllowedDate(s) { return s===todayStr() || s===prevWorkingDayStr(); }
  function formatDate(d)  { return new Date(d+'T12:00:00').toLocaleDateString('en-US',{month:'short',day:'numeric'}); }
  function formatDay(d)   { return new Date(d+'T12:00:00').toLocaleDateString('en-US',{weekday:'long',year:'numeric',month:'long',day:'numeric'}); }
  function dlCSV(csv,fn)  { const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([csv],{type:'text/csv'}));a.download=fn;a.click(); }
  function updateSBStats(){ const ms=mySubmissions();const s1=q('#sb-total'),s2=q('#sb-avg');if(s1)s1.textContent=ms.length;if(s2)s2.textContent=calcAvgProd(ms).toFixed(0)+'%'; }
  function toast(msg,type='info'){ let el=document.getElementById('dtr-toast');if(!el){el=document.createElement('div');el.id='dtr-toast';(document.getElementById('dtr-root-outer')||root).appendChild(el);}const icons={ok:'✅',err:'❌',info:'💡'};el.innerHTML='<span>'+(icons[type]||'💡')+'</span><span>'+msg+'</span>';el.className='show t'+type;clearTimeout(el._t);el._t=setTimeout(()=>{el.className='';},4000); }

  // INIT
  teamStatusCache = safeLoadObj('dtr_teamcache', {});
  if (document.readyState==='loading') { document.addEventListener('DOMContentLoaded', boot); } else { boot(); }

})();
