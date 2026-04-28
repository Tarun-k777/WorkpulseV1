// ==UserScript==
// @name         WorkPulse — Admin
// @namespace    https://amazon.sharepoint.com/sites/teamdailytask/
// @version      2.1.0
// @description  Admin: Task Reporter + Team Attendance Week View + All NPT Overview. Single SharePoint.
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

  const ADMIN_NAME = 'Team Manager'; // << CHANGE THIS

  const SP = {
    SITE:        'https://amazon.sharepoint.com/sites/teamdailytask',
    TASK_LIST:   'DailyTaskReport',
    STATUS_LIST: 'AttendanceStatus',
    NPT_LIST:    'NPTLog',
  };

  const WH = 8;
  const TASK_TYPES = ['CT Production','Pre-Prod Production','Simulator Production','CT Audits','Pre-Prod Audits','Simulator Audits','Lack of Work','Ad-hoc Tasks'];
  const COLORS = ['#6366f1','#8b5cf6','#06b6d4','#10b981','#f59e0b','#ef4444','#ec4899','#84cc16'];
  const STATUSES = ['WFO','WFH','SL','CL','AL','Optional Off'];
  const PROCS = ['Preprod Testing','Chat Transcripts','Adhoc','Quality Check','Training','Other'];
  const STATUS_CFG = {
    WFO:{label:'Work From Office',color:'#3fb950',bg:'rgba(63,185,80,.14)',cls:'sp-wfo'},
    WFH:{label:'Work From Home',color:'#22d3ee',bg:'rgba(6,182,212,.14)',cls:'sp-wfh'},
    SL:{label:'Sick Leave',color:'#f85149',bg:'rgba(248,81,73,.14)',cls:'sp-sl'},
    CL:{label:'Casual Leave',color:'#d29922',bg:'rgba(210,153,34,.14)',cls:'sp-cl'},
    AL:{label:'Annual Leave',color:'#a371f7',bg:'rgba(163,113,247,.14)',cls:'sp-al'},
    'Optional Off':{label:'Optional Off',color:'#6b7280',bg:'rgba(107,114,128,.14)',cls:'sp-oo'},
  };
  const NPT_TYPES = ['System Issue','Meeting Overrun','Training','Lack of Work','Power Outage','Network Issue','Admin Task','Other'];

  function safeLoad(k,fb){try{const v=GM_getValue(k,null);if(v===null)return fb;if(Array.isArray(v))return v;const p=JSON.parse(v);return Array.isArray(p)?p:fb;}catch(e){return fb;}}
  function safeSave(k,v){try{GM_setValue(k,Array.isArray(v)?v:[]);}catch(e){}}
  function safeLoadObj(k,fb){try{const v=GM_getValue(k,'');if(!v)return fb;const p=JSON.parse(v);return(p&&typeof p==='object'&&!Array.isArray(p))?p:fb;}catch(e){return fb;}}
  function safeSaveObj(k,v){try{GM_setValue(k,JSON.stringify(v));}catch(e){}}

  let submissions     = safeLoad('dtr_subs4',[]);
  let customFields    = safeLoad('dtr_cf4',[]);
  let excelRows       = safeLoad('dtr_excel4',[]);
  let teamStatusCache = safeLoadObj('dtr_teamcache',{});
  let nptAllCache     = safeLoad('dtr_npt2',[]);
  let theme           = GM_getValue('dtr_theme4','dark');
  let currentView     = '';
  let rankPeriod      = 'month';
  let rankSelected    = [];
  let weekOffset      = 0;
  let nptFilterName   = '';
  let root;

  GM_addStyle(`
    @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700&family=DM+Mono:wght@400;500&display=swap');
    *{box-sizing:border-box;margin:0;padding:0}
    :root{--bg:#0d1117;--bg2:#161b22;--bg3:#1c2333;--bg4:#21262d;--border:rgba(255,255,255,0.08);--border2:rgba(255,255,255,0.14);--text:#e6edf3;--text2:#8b949e;--text3:#484f58;--accent:#58a6ff;--accent2:#a371f7;--green:#3fb950;--amber:#d29922;--red:#f85149;--purple:#a371f7;--radius:10px;--font:'DM Sans',sans-serif;--mono:'DM Mono',monospace}
    .lt{--bg:#f6f8fa;--bg2:#ffffff;--bg3:#f0f3f6;--bg4:#e1e5ea;--border:rgba(0,0,0,0.09);--border2:rgba(0,0,0,0.15);--text:#1f2328;--text2:#656d76;--text3:#9198a1;--accent:#0969da;--accent2:#8250df;--green:#1a7f37;--amber:#9a6700;--red:#d1242f}
    #dtr-root-outer{position:fixed;inset:0;z-index:2147483647;overflow:hidden;background:var(--bg)}#dtr-root{position:absolute;top:0;left:0;width:90.9%;height:90.9%;transform:scale(1.1);transform-origin:top left;font-family:var(--font);font-size:1rem;background:var(--bg);color:var(--text);display:flex;flex-direction:column;overflow:hidden}
    #dtr-topbar{height:56px;flex-shrink:0;background:var(--bg2);border-bottom:1px solid var(--border);display:flex;align-items:center;padding:0 20px;gap:8px}
    .dtr-logo{display:flex;align-items:center;gap:8px;padding-right:14px;border-right:1px solid var(--border);margin-right:6px;flex-shrink:0}
    .dtr-logo-icon{width:28px;height:28px;border-radius:7px;background:linear-gradient(135deg,#58a6ff,#a371f7);display:flex;align-items:center;justify-content:center}
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
    .role-badge{padding:2px 7px;border-radius:4px;font-size:.77rem;font-weight:700;text-transform:uppercase;letter-spacing:.5px;background:rgba(163,113,247,.15);color:var(--accent2)}
    .icon-btn{width:30px;height:30px;border-radius:6px;border:1px solid var(--border);background:var(--bg3);color:var(--text2);cursor:pointer;display:flex;align-items:center;justify-content:center;transition:all .15s}
    .icon-btn:hover{background:var(--bg4);color:var(--text)}
    .icon-btn.danger:hover{background:rgba(248,81,73,.1);color:var(--red);border-color:rgba(248,81,73,.3)}
    .icon-btn svg{width:14px;height:14px}
    #dtr-body{flex:1;overflow:hidden;display:flex}
    #dtr-sidebar{width:218px;flex-shrink:0;background:var(--bg2);border-right:1px solid var(--border);display:flex;flex-direction:column;padding:12px 10px;overflow-y:auto}
    .sb-sec{font-size:.7rem;font-weight:700;text-transform:uppercase;letter-spacing:1px;color:var(--text3);padding:8px 8px 4px;margin-top:6px}
    .sb-item{display:flex;align-items:center;gap:9px;padding:8px 10px;border-radius:7px;border:none;background:transparent;color:var(--text2);font-family:var(--font);font-size:.875rem;font-weight:500;cursor:pointer;transition:all .15s;width:100%;text-align:left}
    .sb-item:hover{background:var(--bg3);color:var(--text)}
    .sb-item.active{background:rgba(163,113,247,.1);color:var(--accent2);font-weight:600}
    .sb-item svg{width:14px;height:14px;flex-shrink:0}
    .sb-footer{margin-top:auto;padding:10px 8px;border-top:1px solid var(--border)}
    .sb-stat{margin-bottom:10px}
    .sb-stat .lbl{font-size:.75rem;color:var(--text3);margin-bottom:2px}
    .sb-stat .val{font-size:1.25rem;font-weight:700;color:var(--accent2)}
    #dtr-main{flex:1;overflow-y:auto;padding:24px 28px;background:var(--bg)}
    #dtr-main::-webkit-scrollbar{width:4px}
    #dtr-main::-webkit-scrollbar-thumb{background:var(--bg4);border-radius:3px}
    .dtr-field{display:flex;flex-direction:column;gap:5px;margin-bottom:12px}
    .dtr-label{font-size:.82rem;font-weight:600;color:var(--text2);letter-spacing:.2px}
    .dtr-input,.dtr-select,.dtr-textarea{padding:8px 11px;border-radius:var(--radius);background:var(--bg3);border:1px solid var(--border);color:var(--text);font-family:var(--font);font-size:.96rem;transition:border-color .15s,box-shadow .15s;width:100%;appearance:none}
    .dtr-input:focus,.dtr-select:focus,.dtr-textarea:focus{outline:none;border-color:var(--accent2);box-shadow:0 0 0 3px rgba(163,113,247,.15)}
    .dtr-input::placeholder{color:var(--text3)}
    .dtr-select option{background:var(--bg2)}
    .btn{display:inline-flex;align-items:center;justify-content:center;gap:5px;padding:8px 16px;border-radius:var(--radius);font-family:var(--font);font-size:.94rem;font-weight:600;cursor:pointer;border:none;transition:all .15s;white-space:nowrap}
    .btn svg{width:13px;height:13px}
    .btn-primary{background:var(--accent2);color:#fff}.btn-primary:hover{opacity:.88}
    .btn-ghost{background:var(--bg3);border:1px solid var(--border);color:var(--text2)}.btn-ghost:hover{background:var(--bg4);color:var(--text)}
    .btn-danger{background:rgba(248,81,73,.1);border:1px solid rgba(248,81,73,.25);color:var(--red)}.btn-danger:hover{background:rgba(248,81,73,.18)}
    .btn-full{width:100%}.btn-sm{padding:5px 11px;font-size:.875rem}.btn-xs{padding:3px 8px;font-size:.815rem}
    .ph{display:flex;align-items:flex-start;justify-content:space-between;margin-bottom:18px;gap:12px}
    .ph-left .ph-title{font-size:1.1rem;font-weight:700;color:var(--text);letter-spacing:-.3px}
    .ph-left .ph-sub{font-size:.86rem;color:var(--text3);margin-top:2px}
    .ph-actions{display:flex;gap:6px;align-items:center;flex-shrink:0}
    .card{background:var(--bg2);border:1px solid var(--border);border-radius:12px;padding:16px;margin-bottom:14px}
    .card-title{font-size:.78rem;font-weight:700;text-transform:uppercase;letter-spacing:.7px;color:var(--text3);margin-bottom:12px;display:flex;align-items:center;gap:6px}
    .stats-grid{display:grid;gap:12px;margin-bottom:16px}
    .sg4{grid-template-columns:repeat(4,1fr)}.sg3{grid-template-columns:repeat(3,1fr)}.sg2{grid-template-columns:1fr 1fr}
    .stat-card{background:var(--bg2);border:1px solid var(--border);border-radius:12px;padding:14px}
    .stat-card .lbl{font-size:.77rem;color:var(--text3);text-transform:uppercase;letter-spacing:.5px;margin-bottom:6px}
    .stat-card .val{font-size:1.8rem;font-weight:800;color:var(--text);letter-spacing:-1px;line-height:1}
    .stat-card .sub{font-size:.8rem;color:var(--text3);margin-top:4px}
    .ab{border-top:2px solid var(--accent)}.gb{border-top:2px solid var(--green)}.amb{border-top:2px solid var(--amber)}.pb{border-top:2px solid var(--purple)}.rb2{border-top:2px solid var(--red)}
    .g2{display:grid;grid-template-columns:1fr 1fr;gap:12px}
    .g3{display:grid;grid-template-columns:repeat(3,1fr);gap:12px}
    .tbl-wrap{border:1px solid var(--border);border-radius:10px;overflow:hidden}
    .dtr-table{width:100%;border-collapse:collapse;font-size:.91rem}
    .dtr-table th{background:var(--bg3);padding:11px 14px;text-align:left;color:var(--text3);font-weight:600;font-size:.78rem;text-transform:uppercase;letter-spacing:.6px;border-bottom:1px solid var(--border)}
    .dtr-table td{padding:11px 14px;color:var(--text2);border-bottom:1px solid var(--border)}
    .dtr-table tr:last-child td{border-bottom:none}
    .dtr-table tr:hover td{background:var(--bg3)}
    .dtr-table .bold{font-weight:600;color:var(--text)}
    .dtr-table .mono{font-family:var(--mono)}
    .badge{display:inline-flex;align-items:center;gap:3px;padding:2px 7px;border-radius:4px;font-size:.77rem;font-weight:700}
    .bg2{background:rgba(63,185,80,.12);color:var(--green)}.ba{background:rgba(210,153,34,.12);color:var(--amber)}.bb{background:rgba(88,166,255,.12);color:var(--accent)}.br2{background:rgba(248,81,73,.12);color:var(--red)}
    .chart-title{font-size:.78rem;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--text3);margin-bottom:12px}
    .bar-rows{display:flex;flex-direction:column;gap:8px}
    .bar-row{display:flex;align-items:center;gap:10px}
    .bar-label{font-size:.855rem;color:var(--text2);width:130px;flex-shrink:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
    .bar-track{flex:1;height:20px;background:var(--bg4);border-radius:5px;overflow:hidden}
    .bar-fill{height:100%;border-radius:5px;display:flex;align-items:center;padding-left:7px;transition:width .7s cubic-bezier(.34,1.56,.64,1)}
    .bar-fill-text{font-size:.79rem;font-weight:700;color:#fff}
    .filter-bar{display:flex;gap:8px;margin-bottom:14px;align-items:center;flex-wrap:wrap}
    .filter-bar .dtr-input,.filter-bar .dtr-select{max-width:180px;font-size:.9rem;padding:7px 10px}
    .member-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(140px,1fr));gap:10px;margin-bottom:14px}
    .member-card{background:var(--bg2);border:1px solid var(--border);border-radius:10px;padding:14px;cursor:pointer;transition:all .15s;text-align:center}
    .member-card:hover{border-color:var(--accent2);transform:translateY(-1px)}
    .member-card.selected{border-color:var(--accent2);background:rgba(163,113,247,.06)}
    .member-av{width:36px;height:36px;border-radius:50%;margin:0 auto 8px;display:flex;align-items:center;justify-content:center;font-size:14px;font-weight:800;color:#fff}
    .member-name{font-size:.9rem;font-weight:600;color:var(--text)}
    .member-sub{font-size:.77rem;color:var(--text3);margin-top:2px}
    .rank-chip{display:inline-flex;align-items:center;gap:5px;padding:5px 12px;border-radius:20px;border:1px solid var(--border);background:var(--bg3);color:var(--text2);font-size:.865rem;font-weight:600;cursor:pointer;transition:all .15s;user-select:none;margin:3px}
    .rank-chip:hover{border-color:var(--accent2);color:var(--text)}
    .rank-chip.on{background:rgba(163,113,247,.12);border-color:var(--accent2);color:var(--accent2)}
    .rpb{padding:5px 14px;border-radius:6px;border:1px solid var(--border);background:var(--bg3);color:var(--text2);font-family:var(--font);font-size:.875rem;font-weight:600;cursor:pointer;transition:all .15s;margin-right:5px}
    .rpb.active{background:rgba(163,113,247,.12);border-color:var(--accent2);color:var(--accent2)}
    .sp{display:inline-flex;align-items:center;padding:3px 9px;border-radius:6px;font-size:.75rem;font-weight:700;white-space:nowrap}
    .sp-wfo{background:rgba(63,185,80,.14);color:#3fb950}
    .sp-wfh{background:rgba(6,182,212,.14);color:#22d3ee}
    .sp-sl{background:rgba(248,81,73,.14);color:#f85149}
    .sp-cl{background:rgba(210,153,34,.14);color:#d29922}
    .sp-al{background:rgba(163,113,247,.14);color:#a371f7}
    .sp-oo{background:rgba(107,114,128,.14);color:#9ca3af}
    .sp-ns{background:var(--bg3);color:var(--text3);border:1px dashed var(--border2)}
    .sp-we{background:transparent;color:var(--text3);font-style:italic;font-size:.7rem}
    .wv-controls{display:flex;align-items:center;gap:10px;margin-bottom:20px;background:var(--bg2);border:1px solid var(--border);border-radius:12px;padding:16px 20px;flex-wrap:wrap}
    .wv-nav-btn{width:34px;height:34px;border-radius:9px;border:1px solid var(--border);background:var(--bg3);display:flex;align-items:center;justify-content:center;cursor:pointer;color:var(--text2);transition:.2s;flex-shrink:0}
    .wv-nav-btn:hover{background:var(--bg4);color:var(--text)}
    .wv-nav-btn svg{width:16px;height:16px}
    .wv-range{flex:1;min-width:0}.wv-range-title{font-size:.975rem;font-weight:700;color:var(--text)}.wv-range-sub{font-size:.78rem;color:var(--text3)}
    .wv-today-btn{padding:5px 12px;border-radius:7px;border:1px solid var(--border);background:var(--bg3);font-size:.8rem;font-weight:600;color:var(--text2);cursor:pointer;font-family:var(--font);transition:.2s}
    .wv-today-btn:hover{background:rgba(163,113,247,.1);color:var(--accent2);border-color:var(--accent2)}
    .wv-grid-wrap{border:1px solid var(--border);border-radius:12px;overflow:hidden;background:var(--bg2)}
    .wv-header{display:grid;grid-template-columns:170px repeat(7,1fr);background:var(--bg3);border-bottom:2px solid var(--border)}
    .wv-hcell{padding:10px 8px;text-align:center;font-size:.72rem;font-weight:700;color:var(--text3);text-transform:uppercase;letter-spacing:.5px}
    .wv-hcell.today-col{color:var(--accent2);border-bottom:2px solid var(--accent2);margin-bottom:-2px}
    .wv-hcell .wv-hdate{font-size:.68rem;color:var(--text3);margin-top:2px;text-transform:none;letter-spacing:0;font-weight:500}
    .wv-hcell.today-col .wv-hdate{color:var(--accent2)}
    .wv-row{display:grid;grid-template-columns:170px repeat(7,1fr);border-bottom:1px solid var(--border);transition:.15s}
    .wv-row:last-child{border-bottom:none}
    .wv-row:hover{background:rgba(163,113,247,.025)}
    .wv-name-cell{padding:12px;display:flex;align-items:center;gap:9px;border-right:1px solid var(--border)}
    .wv-av{width:28px;height:28px;border-radius:7px;display:flex;align-items:center;justify-content:center;font-size:10px;font-weight:700;color:#fff;flex-shrink:0}
    .wv-name{font-size:.83rem;font-weight:600;color:var(--text);overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
    .wv-cell{padding:8px 5px;text-align:center;border-right:1px solid rgba(255,255,255,.03);display:flex;flex-direction:column;align-items:center;justify-content:center;gap:3px;cursor:pointer}
    .wv-cell:last-child{border-right:none}
    .wv-cell.today-col{background:rgba(163,113,247,.03)}
    .wv-cell.weekend-col{opacity:.4}
    .wv-cell:hover{background:rgba(163,113,247,.07)}
    .wv-detail-panel{background:var(--bg2);border:1px solid var(--accent2);border-radius:12px;padding:16px;margin-top:14px;animation:fi .2s ease}
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
    .npt-type-chip{padding:4px 12px;border-radius:20px;border:1px solid var(--border);background:var(--bg3);color:var(--text2);font-family:var(--font);font-size:.8rem;font-weight:600;cursor:pointer;transition:all .15s}
    .npt-type-chip:hover{border-color:var(--accent2);color:var(--accent2)}
    .npt-type-chip.active{background:rgba(163,113,247,.12);border-color:var(--accent2);color:var(--accent2)}
    .info-banner{padding:9px 13px;border-radius:8px;background:rgba(163,113,247,.07);border:1px solid rgba(163,113,247,.2);font-size:.81rem;color:var(--text2);margin-bottom:14px;display:flex;align-items:center;gap:8px;line-height:1.5}
    #dtr-toast{position:fixed;bottom:20px;right:20px;z-index:2147483648;padding:10px 16px;border-radius:10px;background:var(--bg2);border:1px solid var(--border2);color:var(--text);font-family:var(--font);font-size:.92rem;font-weight:500;box-shadow:0 8px 32px rgba(0,0,0,.4);transform:translateY(16px);opacity:0;transition:all .25s cubic-bezier(.34,1.56,.64,1);display:flex;align-items:center;gap:8px;pointer-events:none;max-width:360px}
    #dtr-toast.show{transform:translateY(0);opacity:1}
    #dtr-toast.tok{border-color:rgba(63,185,80,.4)}
    #dtr-toast.terr{border-color:rgba(248,81,73,.4)}
    #dtr-toast.tinfo{border-color:rgba(163,113,247,.35)}
    hr.sep{border:none;border-top:1px solid var(--border);margin:12px 0}
    .anim{animation:fi .22s ease}
    @keyframes fi{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:translateY(0)}}
    .empty{text-align:center;padding:32px 20px;color:var(--text3)}
    .empty p{font-size:.94rem}
  `);

  // ICONS
  const ic = {
    logo:`<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="white" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="3"/><path d="M7 8h10M7 12h6M7 16h8"/></svg>`,
    overview:`<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="7" height="7"/><rect x="14" y="3" width="7" height="7"/><rect x="3" y="14" width="7" height="7"/><rect x="14" y="14" width="7" height="7"/></svg>`,
    team:`<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87M16 3.13a4 4 0 0 1 0 7.75"/></svg>`,
    rank:`<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 2l3.09 6.26L22 9.27l-5 4.87 1.18 6.88L12 17.77l-6.18 3.25L7 14.14 2 9.27l6.91-1.01L12 2z"/></svg>`,
    tracker:`<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M9 3v18"/></svg>`,
    att:`<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>`,
    npt:`<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>`,
    settings:`<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="3"/><path d="M19.07 4.93l-1.41 1.41M4.93 4.93l1.41 1.41M12 2v2M12 20v2M20 12h2M2 12h2M19.07 19.07l-1.41-1.41M4.93 19.07l1.41-1.41"/></svg>`,
    sun:`<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="5"/><line x1="12" y1="1" x2="12" y2="3"/><line x1="12" y1="21" x2="12" y2="23"/><line x1="4.22" y1="4.22" x2="5.64" y2="5.64"/><line x1="18.36" y1="18.36" x2="19.78" y2="19.78"/><line x1="1" y1="12" x2="3" y2="12"/><line x1="21" y1="12" x2="23" y2="12"/><line x1="4.22" y1="19.78" x2="5.64" y2="18.36"/><line x1="18.36" y1="5.64" x2="19.78" y2="4.22"/></svg>`,
    moon:`<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"/></svg>`,
    close:`<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>`,
    left:`<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="15 18 9 12 15 6"/></svg>`,
    right:`<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="9 18 15 12 9 6"/></svg>`,
    check:`<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="20 6 9 17 4 12"/></svg>`,
    add:`<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>`,
    trash:`<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14H6L5 6"/><path d="M10 11v6M14 11v6M9 6V4h6v2"/></svg>`,
    download:`<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>`,
    sync:`<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="23 4 23 10 17 10"/><polyline points="1 20 1 14 7 14"/><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/></svg>`,
    export:`<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>`,
  };

  const TABS = [
    {id:'overview',  label:'Overview',      icon:ic.overview, group:'Task Reporter'},
    {id:'teamrpt',   label:'Team Report',   icon:ic.team,     group:'Task Reporter'},
    {id:'rankings',  label:'Rankings',      icon:ic.rank,     group:'Task Reporter'},
    {id:'tracker',   label:'Team Tracker',  icon:ic.tracker,  group:'Task Reporter'},
    {id:'attweek',   label:'Week View',     icon:ic.att,      group:'Attendance'},
    {id:'attnpt',    label:'All NPT Log',   icon:ic.npt,      group:'Attendance'},
    {id:'settings',  label:'Settings',      icon:ic.settings, group:'Settings'},
  ];

  // BOOT
  function boot() {
    const exO = document.getElementById('dtr-root-outer'); if (exO) exO.remove();
    const outer = document.createElement('div'); outer.id = 'dtr-root-outer';
    root = document.createElement('div'); root.id = 'dtr-root';
    if (theme === 'light') { root.classList.add('lt'); outer.classList.add('lt'); }
    outer.appendChild(root);
    document.body.appendChild(outer);
    root.addEventListener('click',  handleClick);
    root.addEventListener('input',  handleInput);
    root.addEventListener('change', handleChange);
    if (window.location.hostname.includes('sharepoint.com')) setTimeout(fetchAllData, 2000);
    buildApp();
  }

  const q  = sel => root.querySelector(sel);
  const qa = sel => root.querySelectorAll(sel);

  function buildApp() {
    let sbHtml = '', grp = '';
    TABS.forEach(t => {
      if (t.group !== grp) { grp = t.group; sbHtml += '<div class="sb-sec">' + grp + '</div>'; }
      sbHtml += '<button class="sb-item" data-action="switch-tab" data-val="' + t.id + '">' + t.icon + ' ' + t.label + '</button>';
    });
    const members = getAllMembers();
    root.innerHTML =
      '<div id="dtr-topbar">' +
        '<div class="dtr-logo"><div class="dtr-logo-icon">' + ic.logo + '</div><span class="dtr-logo-text">WorkPulse Admin</span></div>' +
        '<div class="dtr-tabs">' + TABS.map(t => '<button class="dtr-tab" data-action="switch-tab" data-val="' + t.id + '">' + t.icon + ' ' + t.label + '</button>').join('') + '</div>' +
        '<div class="dtr-topbar-r">' +
          '<div class="dtr-user-pill"><div class="av">' + ADMIN_NAME[0].toUpperCase() + '</div><span>' + ADMIN_NAME + '</span><span class="role-badge">Admin</span></div>' +
          '<button class="icon-btn" data-action="toggle-theme" id="theme-btn">' + (theme==='dark'?ic.sun:ic.moon) + '</button>' +
          '<button class="icon-btn" data-action="sync-all" title="Sync All Data">' + ic.sync + '</button>' +
          '<button class="icon-btn danger" data-action="close-app">' + ic.close + '</button>' +
        '</div>' +
      '</div>' +
      '<div id="dtr-body">' +
        '<div id="dtr-sidebar">' + sbHtml +
          '<hr class="sep" style="margin:8px 0"><div class="sb-footer">' +
          '<div class="sb-stat"><div class="lbl">Team Members</div><div class="val" id="sb-members">' + members.length + '</div></div>' +
          '<div class="sb-stat"><div class="lbl">Total Entries</div><div class="val" id="sb-total">' + submissions.length + '</div></div>' +
          '</div></div>' +
        '<div id="dtr-main">' + TABS.map(t => '<div class="dtr-view anim" id="view-' + t.id + '" style="display:none"></div>').join('') + '</div>' +
      '</div><div id="dtr-toast"></div>';
    switchTab('overview');
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
    if (v === 'overview') renderOverview();
    if (v === 'teamrpt')  renderTeamReport();
    if (v === 'rankings') renderRankings();
    if (v === 'tracker')  renderTracker();
    if (v === 'attweek')  renderAttWeek();
    if (v === 'attnpt')   renderAllNPT();
    if (v === 'settings') renderSettings();
  }

  function toggleTheme() {
    theme = theme === 'dark' ? 'light' : 'dark';
    GM_setValue('dtr_theme4', theme);
    root.classList.toggle('lt', theme === 'light');
    const b = q('#theme-btn'); if (b) b.innerHTML = theme === 'dark' ? ic.sun : ic.moon;
  }

  // EVENTS
  function handleClick(e) {
    const btn = e.target.closest('[data-action]'); if (!btn) return;
    const a = btn.dataset.action, v = btn.dataset.val || '';
    switch (a) {
      case 'switch-tab':       switchTab(v); break;
      case 'toggle-theme':     toggleTheme(); break;
      case 'close-app':        { const o=document.getElementById('dtr-root-outer'); if(o)o.style.display='none'; } break;
      case 'sync-all':         fetchAllData(); break;
      case 'rank-period':      rankPeriod = v; qa('.rpb').forEach(b => b.classList.toggle('active', b.dataset.val===v)); renderRankings(); break;
      case 'rank-chip':        toggleRankMember(v); break;
      case 'rank-all':         rankSelected = []; qa('.rank-chip').forEach(c => c.classList.remove('on')); renderRankings(); break;
      case 'tracker-add':         trackerAdd(); break;
      case 'adm-tracker-add-col': admTrackerAddCol(); break;
      case 'adm-tracker-del-col': admTrackerDelCol(btn.dataset.key); break;
      case 'tracker-del':      trackerDel(+btn.dataset.idx, btn.dataset.owner); break;
      case 'tracker-export':   trackerExport(); break;
      case 'wv-prev':          weekOffset--; renderAttWeek(); break;
      case 'wv-next':          weekOffset++; renderAttWeek(); break;
      case 'wv-today':         weekOffset = 0; renderAttWeek(); break;
      case 'wv-close-detail':  { const d = q('#wv-detail'); if (d) d.innerHTML = ''; break; }
      case 'sync-att':         toast('Syncing attendance...','info'); fetchAttendance(); break;
      case 'sync-npt':         toast('Syncing NPT log...','info'); fetchNPT(); break;
      case 'npt-filter-clear': nptFilterName = ''; renderAllNPT(); break;
      case 'npt-export':       nptExport(); break;
      case 'test-sp':          testSP(); break;
    }
  }

  function handleInput(e) {
    const el = e.target;
    if (el.id === 'npt-name-filter') { nptFilterName = el.value.trim().toLowerCase(); renderAllNPT(); }
    if (el.id === 'rpt-filter-name' || el.id === 'rpt-filter-date') renderTeamReport();
  }

  function handleChange(e) {
    const el = e.target;
    if (el.dataset.exfield) {
      const idx = +el.closest('tr').dataset.idx;
      const owner = el.closest('tr').dataset.owner || '';
      const row = excelRows.find((r,i) => i === idx);
      if (row) { row[el.dataset.exfield] = el.value; safeSave('dtr_excel4', excelRows); }
    }
  }

  // ═══════════════════════════════════════════════════════════
  // OVERVIEW TAB
  // ═══════════════════════════════════════════════════════════
  function renderOverview() {
    const el = q('#view-overview'); if (!el) return;
    const subs = Array.isArray(submissions) ? submissions : [];
    const members = getAllMembers();
    const today = todayStr();
    const todayCount = subs.filter(s => s.date === today).length;
    const totalH = subs.reduce((a,s) => a+(s.hours||0), 0);
    const avgProd = calcTeamAvgProd(subs);
    // Attendance today counts
    const attToday = members.map(m => getTeamStatus(m, today)).filter(Boolean);
    const wfoCnt = attToday.filter(s => s==='WFO').length;
    const wfhCnt = attToday.filter(s => s==='WFH').length;
    const lvCnt  = attToday.filter(s => ['SL','CL','AL'].includes(s)).length;
    const notMark = members.length - attToday.length;
    // Per-member summary
    const memberStats = members.map(m => {
      const ms = subs.filter(s => s.employeeName === m);
      const prod = calcAvgProd(ms);
      const todaySt = getTeamStatus(m, today);
      return { m, count:ms.length, prod, todaySt };
    }).sort((a,b) => b.prod - a.prod);
    const rc = avgProd>=75?'var(--green)':avgProd>=50?'var(--amber)':'var(--red)';
    const circ = 2*Math.PI*38, off = circ-(avgProd/100)*circ;
    el.innerHTML =
      '<div class="ph"><div class="ph-left"><div class="ph-title">Team Overview</div><div class="ph-sub">Live dashboard — ' + today + '</div></div>' +
      '<div class="ph-actions"><button class="btn btn-ghost btn-sm" data-action="sync-all">' + ic.sync + ' Sync All</button></div></div>' +
      '<div class="stats-grid sg4">' +
      '<div class="stat-card ab"><div class="lbl">Team Members</div><div class="val">' + members.length + '</div><div class="sub">active</div></div>' +
      '<div class="stat-card gb"><div class="lbl">Today Submissions</div><div class="val">' + todayCount + '</div><div class="sub">entries today</div></div>' +
      '<div class="stat-card amb"><div class="lbl">Total Hours Logged</div><div class="val">' + totalH.toFixed(0) + 'h</div><div class="sub">all time</div></div>' +
      '<div class="stat-card pb"><div class="lbl">Avg Productivity</div><div class="val">' + avgProd.toFixed(0) + '%</div><div class="sub">team average</div></div></div>' +
      '<div class="stats-grid sg4" style="margin-bottom:18px">' +
      '<div class="stat-card gb"><div class="lbl">WFO Today</div><div class="val" style="color:#3fb950">' + wfoCnt + '</div></div>' +
      '<div class="stat-card ab"><div class="lbl">WFH Today</div><div class="val" style="color:#22d3ee">' + wfhCnt + '</div></div>' +
      '<div class="stat-card amb"><div class="lbl">On Leave</div><div class="val" style="color:var(--amber)">' + lvCnt + '</div></div>' +
      '<div class="stat-card"><div class="lbl">Not Marked</div><div class="val" style="color:var(--text3)">' + notMark + '</div></div></div>' +
      '<div class="g2">' +
      '<div class="card"><div class="chart-title">Team Productivity Score</div><div style="display:flex;align-items:center;gap:20px">' +
      '<svg width="90" height="90" viewBox="0 0 90 90"><circle cx="45" cy="45" r="38" fill="none" stroke="var(--bg4)" stroke-width="9"/>' +
      '<circle cx="45" cy="45" r="38" fill="none" stroke="' + rc + '" stroke-width="9" stroke-dasharray="' + circ.toFixed(2) + '" stroke-dashoffset="' + off.toFixed(2) + '" stroke-linecap="round" transform="rotate(-90 45 45)" style="transition:stroke-dashoffset 1s ease"/>' +
      '<text x="45" y="45" text-anchor="middle" dy=".35em" fill="' + rc + '" font-size="15" font-weight="800" font-family="DM Sans,sans-serif">' + avgProd.toFixed(0) + '%</text></svg>' +
      '<div><div style="font-size:1.4rem;font-weight:800;color:var(--text)">' + avgProd.toFixed(0) + '%</div><div style="font-size:.85rem;color:var(--text3);margin-top:3px">' + members.length + ' members · ' + subs.length + ' total entries</div></div></div></div>' +
      '<div class="card"><div class="chart-title">Today Attendance</div>' +
      (members.length === 0 ? '<div class="empty"><p>No team data. Sync to load.</p></div>' :
        '<div style="display:flex;flex-direction:column;gap:6px">' +
        members.map(m => {
          const st = getTeamStatus(m, today) || '';
          const ini = m.split(/[\s,]+/).filter(Boolean).map(x=>x[0].toUpperCase()).join('').slice(0,2);
          const idx = members.indexOf(m);
          return '<div style="display:flex;align-items:center;gap:8px;padding:6px 0;border-bottom:1px solid var(--border)">' +
            '<div class="wv-av" style="background:' + COLORS[idx%COLORS.length] + ';width:24px;height:24px;font-size:9px">' + ini + '</div>' +
            '<span style="flex:1;font-size:.85rem;color:var(--text)">' + m + '</span>' +
            (st ? '<span class="sp sp-' + st.toLowerCase().replace(' ','-') + '">' + st + '</span>' : '<span class="sp sp-ns">—</span>') +
            '</div>';
        }).join('') + '</div>') + '</div></div>' +
      '<div class="card"><div class="chart-title">Member Performance</div>' +
      '<div class="tbl-wrap"><table class="dtr-table"><thead><tr><th>#</th><th>Member</th><th>Entries</th><th>Avg Prod</th><th>Today Status</th></tr></thead><tbody>' +
      (memberStats.length ? memberStats.map((s,i) => {
        const rc2 = s.prod>=75?'var(--green)':s.prod>=50?'var(--amber)':'var(--red)';
        return '<tr><td style="color:var(--text3)">' + (i+1) + '</td><td class="bold">' + s.m + '</td><td>' + s.count + '</td>' +
          '<td><span style="color:' + rc2 + ';font-weight:700;font-family:var(--mono)">' + s.prod.toFixed(0) + '%</span></td>' +
          '<td>' + (s.todaySt ? '<span class="sp sp-' + s.todaySt.toLowerCase().replace(' ','-') + '">' + s.todaySt + '</span>' : '<span class="sp sp-ns">—</span>') + '</td></tr>';
      }).join('') : '<tr><td colspan="5"><div class="empty"><p>No data. Sync from SP or wait for associates to submit.</p></div></td></tr>') +
      '</tbody></table></div></div>';
  }

  // ═══════════════════════════════════════════════════════════
  // TEAM REPORT TAB
  // ═══════════════════════════════════════════════════════════
  function renderTeamReport() {
    const el = q('#view-teamrpt'); if (!el) return;
    const nameFilter = (q('#rpt-filter-name')?.value||'').toLowerCase();
    const dateFilter = q('#rpt-filter-date')?.value||'';
    const subs = (Array.isArray(submissions)?submissions:[])
      .filter(s => (!nameFilter || (s.employeeName||'').toLowerCase().includes(nameFilter)))
      .filter(s => (!dateFilter || s.date === dateFilter));
    const members = getAllMembers();
    el.innerHTML =
      '<div class="ph"><div class="ph-left"><div class="ph-title">Team Report</div><div class="ph-sub">All submissions across team</div></div>' +
      '<div class="ph-actions"><button class="btn btn-ghost btn-sm" data-action="tracker-export">' + ic.export + ' Export</button></div></div>' +
      '<div class="filter-bar">' +
      '<input type="text" class="dtr-input" id="rpt-filter-name" placeholder="Filter by name..." value="' + (nameFilter||'') + '" style="max-width:180px">' +
      '<input type="date" class="dtr-input" id="rpt-filter-date" value="' + dateFilter + '" style="max-width:155px">' +
      (dateFilter||nameFilter ? '<button class="btn btn-ghost btn-sm" onclick="document.getElementById(&quot;rpt-filter-name&quot;).value=&quot;&quot;;document.getElementById(&quot;rpt-filter-date&quot;).value=&quot;&quot;;renderTeamReport();">Clear</button>' : '') +
      '<span style="margin-left:auto;font-size:.82rem;color:var(--text3)">' + subs.length + ' entries</span></div>' +
      '<div class="tbl-wrap"><table class="dtr-table"><thead><tr><th>Date</th><th>Employee</th><th>Task Type</th><th>Work Type</th><th>Hours</th><th>NPT</th><th>Notes</th></tr></thead><tbody>' +
      (subs.length ? subs.slice().reverse().slice(0,100).map(s =>
        '<tr><td class="mono">' + s.date + '</td>' +
        '<td class="bold">' + (s.employeeName||'—') + '</td>' +
        '<td>' + (s.taskType||'—') + '</td>' +
        '<td><span class="badge ' + (s.workType==='NPT'?'ba':'bg2') + '">' + (s.workType||'Productive') + '</span></td>' +
        '<td class="mono">' + (s.hours||0).toFixed(1) + '</td>' +
        '<td class="mono" style="color:var(--amber)">' + (s.npt||0).toFixed(1) + '</td>' +
        '<td style="max-width:180px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:var(--text3);font-size:.8rem">' + (s.adhoc||'') + '</td></tr>'
      ).join('') : '<tr><td colspan="7"><div class="empty"><p>No submissions found.</p></div></td></tr>') +
      '</tbody></table></div>';
  }

  // ═══════════════════════════════════════════════════════════
  // RANKINGS TAB
  // ═══════════════════════════════════════════════════════════
  function renderRankings() {
    const el = q('#view-rankings'); if (!el) return;
    const members = getAllMembers();
    const subs = Array.isArray(submissions) ? submissions : [];
    const now = new Date(), thisMonth = now.toISOString().slice(0,7);
    const weekStart = getWeekStart(0).toISOString().split('T')[0];
    const filtered = subs.filter(s => {
      if (rankPeriod==='week')  return s.date >= weekStart;
      if (rankPeriod==='month') return (s.date||'').startsWith(thisMonth);
      return true;
    });
    const showMembers = rankSelected.length ? rankSelected : members;
    const ranks = showMembers.map(m => {
      const ms = filtered.filter(s => s.employeeName === m);
      const prod = calcAvgProd(ms), hours = ms.reduce((a,s)=>a+(s.hours||0),0), entries = ms.length;
      return { m, prod, hours, entries };
    }).sort((a,b) => b.prod - a.prod);
    el.innerHTML =
      '<div class="ph"><div class="ph-left"><div class="ph-title">Rankings</div><div class="ph-sub">Sorted by productivity score</div></div></div>' +
      '<div class="card"><div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin-bottom:12px">' +
      ['week','month','all'].map(p => '<button class="rpb' + (rankPeriod===p?' active':'') + '" data-action="rank-period" data-val="' + p + '">' + p.charAt(0).toUpperCase()+p.slice(1) + '</button>').join('') +
      '<button class="btn btn-ghost btn-xs" data-action="rank-all" style="margin-left:auto">All Members</button></div>' +
      '<div class="member-grid">' +
      members.map((m,i) => {
        const on = !rankSelected.length || rankSelected.includes(m);
        const ini = m.split(/[\s,]+/).filter(Boolean).map(x=>x[0].toUpperCase()).join('').slice(0,2);
        return '<div class="member-card' + (on?' selected':'') + '" data-action="rank-chip" data-val="' + m + '">' +
          '<div class="member-av" style="background:' + COLORS[i%COLORS.length] + '">' + ini + '</div>' +
          '<div class="member-name">' + m.split(/[\s,]+/)[0] + '</div></div>';
      }).join('') + '</div></div>' +
      '<div class="card"><div class="chart-title">Leaderboard — ' + rankPeriod + '</div>' +
      (ranks.length ? ranks.map((r,i) => {
        const medal = i===0?'🥇':i===1?'🥈':i===2?'🥉':'';
        const rc2 = r.prod>=75?'var(--green)':r.prod>=50?'var(--amber)':'var(--red)';
        return '<div style="display:flex;align-items:center;gap:12px;padding:10px 0;border-bottom:1px solid var(--border)">' +
          '<span style="font-size:1rem;width:24px;text-align:center">' + (medal||'#'+(i+1)) + '</span>' +
          '<div style="flex:1"><div style="font-weight:600;color:var(--text);font-size:.9rem">' + r.m + '</div>' +
          '<div style="font-size:.75rem;color:var(--text3);margin-top:2px">' + r.entries + ' entries · ' + r.hours.toFixed(1) + 'h logged</div></div>' +
          '<div style="text-align:right"><div style="font-size:1rem;font-weight:800;color:' + rc2 + '">' + r.prod.toFixed(0) + '%</div>' +
          '<div class="prog-wrap" style="width:80px;height:5px;margin-top:4px"><div class="prog-bar" style="width:' + r.prod + '%;background:' + rc2 + '"></div></div></div></div>';
      }).join('') : '<div class="empty"><p>No data for this period.</p></div>') + '</div>';
  }

  function toggleRankMember(m) {
    const idx = rankSelected.indexOf(m);
    if (idx > -1) rankSelected.splice(idx,1); else rankSelected.push(m);
    renderRankings();
  }

  // ═══════════════════════════════════════════════════════════
  // TEAM TRACKER — auto-populated from all submissions
  // ═══════════════════════════════════════════════════════════
  let adminTkCustomCols = safeLoad('adm_tkcols', []);

  function getAdminTrackerRows(nameFilter, typeFilter, taskFilter) {
    const subs = (Array.isArray(submissions) ? submissions : [])
      .filter(s => !nameFilter || (s.employeeName||'').toLowerCase().includes(nameFilter))
      .filter(s => !typeFilter || s.workType === typeFilter)
      .filter(s => !taskFilter || s.taskType === taskFilter)
      .slice().sort((a,b) => a.date < b.date ? 1 : -1);
    return subs.map((s,i) => ({
      _idx:     i,
      date:     s.date||'',
      employee: s.employeeName||'',
      taskType: s.taskType||'',
      workType: s.workType||'Productive',
      hours:    (s.hours||0).toFixed(1),
      nptHours: (s.npt||0).toFixed(1),
      leave:    s.taskType==='Leave' ? '✓' : '',
      adhoc:    s.adhoc||'',
      submittedAt: s.submittedAt ? s.submittedAt.replace('T',' ').slice(0,16) : '',
    }));
  }

  function renderTracker() {
    const el = q('#view-tracker'); if (!el) return;
    const fixedCols = [
      {key:'date',       label:'Date'},
      {key:'employee',   label:'Employee'},
      {key:'taskType',   label:'Task Type'},
      {key:'workType',   label:'Work Type'},
      {key:'hours',      label:'Hours'},
      {key:'nptHours',   label:'NPT Hours'},
      {key:'leave',      label:'Leave'},
      {key:'adhoc',      label:'Ad-hoc Notes'},
      {key:'submittedAt',label:'Submitted At'},
    ];
    const allCols = [...fixedCols, ...adminTkCustomCols];
    const rows = getAdminTrackerRows('','','');
    const members = getAllMembers();

    el.innerHTML =
      '<div class="ph"><div class="ph-left"><div class="ph-title">Team Tracker</div>' +
      '<div class="ph-sub">All submissions — ' + rows.length + ' entries across ' + members.length + ' members</div></div>' +
      '<div class="ph-actions">' +
      '<button class="btn btn-ghost btn-sm" data-action="adm-tracker-add-col">' + ic.add + ' Add Column</button>' +
      '<button class="btn btn-ghost btn-sm" data-action="tracker-export">' + ic.export + ' Export CSV</button>' +
      '</div></div>' +
      // Filters
      '<div class="filter-bar" style="display:flex;gap:8px;margin-bottom:12px;flex-wrap:wrap">' +
      '<input type="text" id="adm-tk-name" class="dtr-input" placeholder="Filter by name..." style="max-width:180px">' +
      '<select id="adm-tk-type" class="dtr-select" style="max-width:155px">' +
      '<option value="">All Work Types</option><option value="Productive">Productive</option><option value="NPT">NPT</option></select>' +
      '<select id="adm-tk-task" class="dtr-select" style="max-width:180px">' +
      '<option value="">All Task Types</option>' +
      TASK_TYPES.map(t=>'<option value="'+t+'">'+t+'</option>').join('') +
      '<option value="Leave">Leave</option></select>' +
      '<span style="margin-left:auto;font-size:.82rem;color:var(--text3)">' + rows.length + ' entries</span></div>' +
      // Summary cards
      '<div class="stats-grid sg4" style="margin-bottom:14px">' +
      '<div class="stat-card ab"><div class="lbl">Total Entries</div><div class="val">' + rows.length + '</div></div>' +
      '<div class="stat-card gb"><div class="lbl">Total Hours</div><div class="val">' + rows.reduce((a,r)=>a+parseFloat(r.hours||0),0).toFixed(1) + 'h</div></div>' +
      '<div class="stat-card amb"><div class="lbl">NPT Hours</div><div class="val">' + rows.reduce((a,r)=>a+parseFloat(r.nptHours||0),0).toFixed(1) + 'h</div></div>' +
      '<div class="stat-card rb2"><div class="lbl">Leave Entries</div><div class="val">' + rows.filter(r=>r.leave).length + '</div></div></div>' +
      // Table
      '<div class="tbl-wrap" style="overflow-x:auto"><table class="dtr-table" id="adm-tk-table">' +
      '<thead><tr><th style="color:var(--text3);font-size:.72rem">#</th>' +
      allCols.map(c =>
        '<th>' + c.label +
        (adminTkCustomCols.find(cc=>cc.key===c.key) ?
          ' <button style="background:none;border:none;color:var(--red);cursor:pointer;font-size:.7rem;padding:0 2px" data-action="adm-tracker-del-col" data-key="' + c.key + '">✕</button>' : '') +
        '</th>'
      ).join('') + '</tr></thead>' +
      '<tbody id="adm-tk-tbody">' + buildAdminTrackerRows(rows, allCols) + '</tbody>' +
      '</table></div>';

    // Filter bindings
    const bindFilter = () => {
      const nf = (q('#adm-tk-name')?.value||'').toLowerCase();
      const tf = q('#adm-tk-type')?.value||'';
      const kf = q('#adm-tk-task')?.value||'';
      const filtered = getAdminTrackerRows(nf, tf, kf);
      const tbody = q('#adm-tk-tbody');
      if (tbody) tbody.innerHTML = buildAdminTrackerRows(filtered, allCols);
    };
    ['#adm-tk-name','#adm-tk-type','#adm-tk-task'].forEach(sel => {
      const el2 = q(sel);
      if (el2) el2.addEventListener(sel==='#adm-tk-name'?'input':'change', bindFilter);
    });
  }

  function buildAdminTrackerRows(rows, allCols) {
    if (!rows.length) return '<tr><td colspan="' + (allCols.length+1) + '"><div class="empty"><p>No submissions found. Sync from SharePoint first.</p></div></td></tr>';
    return rows.map((r, i) => {
      const wt = r.workType;
      const rowColor = r.leave ? 'rgba(210,153,34,.04)' : wt==='NPT' ? 'rgba(248,81,73,.03)' : '';
      return '<tr style="' + (rowColor?'background:'+rowColor:'') + '">' +
        '<td style="color:var(--text3);font-size:.78rem;font-weight:600">' + (i+1) + '</td>' +
        allCols.map(c => {
          const val = r[c.key]||'';
          if (c.key==='workType') return '<td><span class="badge ' + (wt==='NPT'?'ba':'bg2') + '">' + val + '</span></td>';
          if (c.key==='leave' && r.leave) return '<td><span class="badge ba">Leave</span></td>';
          if (c.key==='leave') return '<td></td>';
          if (c.key==='hours'||c.key==='nptHours') return '<td class="mono" style="' + (c.key==='nptHours'&&parseFloat(val)>0?'color:var(--amber)':'') + '">' + val + '</td>';
          if (c.key==='employee') return '<td class="bold">' + val + '</td>';
          if (adminTkCustomCols.find(cc=>cc.key===c.key)) return '<td><input class="dtr-input" value="' + val + '" style="padding:3px 7px;font-size:.85rem;min-width:90px" readonly></td>';
          return '<td>' + (val||'—') + '</td>';
        }).join('') + '</tr>';
    }).join('');
  }

  function trackerExport() {
    const rows = getAdminTrackerRows('','','');
    if (!rows.length) { toast('No submissions to export','err'); return; }
    const fixedCols = ['date','employee','taskType','workType','hours','nptHours','leave','adhoc','submittedAt'];
    const fixedLabels = ['Date','Employee','Task Type','Work Type','Hours','NPT Hours','Leave','Ad-hoc Notes','Submitted At'];
    const header = [...fixedLabels,...adminTkCustomCols.map(c=>c.label)].join(',');
    const csvRows = rows.map(r => [...fixedCols.map(k=>'"'+(r[k]||'').replace(/"/g,'""')+'"')].join(','));
    dlCSV([header,...csvRows].join('\n'), 'team_tracker_'+todayStr()+'.csv');
    toast('Exported ' + rows.length + ' rows','ok');
  }

  function admTrackerAddCol() {
    const label = prompt('New column name:');
    if (!label || !label.trim()) return;
    adminTkCustomCols.push({ key:'custom_'+Date.now(), label:label.trim() });
    safeSave('adm_tkcols', adminTkCustomCols);
    renderTracker(); toast('Column added','ok');
  }

  function admTrackerDelCol(key) {
    adminTkCustomCols = adminTkCustomCols.filter(c=>c.key!==key);
    safeSave('adm_tkcols', adminTkCustomCols);
    renderTracker();
  }

  // ═══════════════════════════════════════════════════════════
  // ATTENDANCE WEEK VIEW TAB (Admin — read-only, all members)
  // ═══════════════════════════════════════════════════════════
  function renderAttWeek() {
    const el = q('#view-attweek'); if (!el) return;
    const ws = getWeekStart(weekOffset), wLabel = getWeekLabel(weekOffset), today = todayStr();
    const days = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
    const dates = Array.from({length:7}, (_,i) => {
      const d = new Date(ws); d.setDate(ws.getDate()+i);
      const dk = d.toISOString().split('T')[0];
      return {dk, day:days[i], label:d.toLocaleDateString('en-US',{month:'short',day:'numeric'}), isToday:dk===today, isWeekend:i===0||i===6};
    });
    const allUsers = getAllMembers();
    const workdays = dates.filter(d => !d.isWeekend);
    let wfo=0,wfh=0,sl=0,cl=0,al=0,ns=0;
    allUsers.forEach(u => { workdays.forEach(d => { const st=getTeamStatus(u,d.dk); if(st==='WFO')wfo++;else if(st==='WFH')wfh++;else if(st==='SL')sl++;else if(st==='CL')cl++;else if(st==='AL')al++;else ns++; }); });
    const total = (allUsers.length * workdays.length) || 1;

    el.innerHTML =
      '<div class="ph"><div class="ph-left"><div class="ph-title">Team Attendance Week View</div><div class="ph-sub">Click any day cell for detailed breakdown</div></div>' +
      '<div class="ph-actions"><button class="btn btn-ghost btn-sm" id="att-sync-btn">' + ic.sync + ' Sync</button></div></div>' +
      '<div class="wv-controls">' +
      '<button class="wv-nav-btn" data-action="wv-prev">' + ic.left + '</button>' +
      '<div class="wv-range"><div class="wv-range-title">' + wLabel + '</div><div class="wv-range-sub">' + (weekOffset===0?'Current Week':'') + '</div></div>' +
      (weekOffset!==0?'<button class="wv-today-btn" data-action="wv-today">This Week</button>':'') +
      '<button class="wv-nav-btn" data-action="wv-next">' + ic.right + '</button></div>' +
      '<div class="stats-grid sg4" style="margin-bottom:16px">' +
      '<div class="stat-card gb"><div class="lbl">WFO</div><div class="val" style="color:#3fb950">' + wfo + '</div><div class="sub">' + ((wfo/total)*100).toFixed(0) + '% of entries</div></div>' +
      '<div class="stat-card ab"><div class="lbl">WFH</div><div class="val" style="color:#22d3ee">' + wfh + '</div><div class="sub">' + ((wfh/total)*100).toFixed(0) + '% of entries</div></div>' +
      '<div class="stat-card amb"><div class="lbl">On Leave</div><div class="val" style="color:var(--amber)">' + (sl+cl+al) + '</div><div class="sub">SL ' + sl + ' · CL ' + cl + ' · AL ' + al + '</div></div>' +
      '<div class="stat-card"><div class="lbl">Not Marked</div><div class="val" style="color:var(--text3)">' + ns + '</div><div class="sub">pending</div></div></div>' +
      '<div class="wv-grid-wrap" id="att-grid">' +
      '<div class="wv-header"><div class="wv-hcell" style="text-align:left;padding-left:14px">Team Member</div>' +
      dates.map(d => '<div class="wv-hcell' + (d.isToday?' today-col':'') + '">' + d.day + '<div class="wv-hdate">' + d.label + '</div>' + (d.isToday?'<div style="font-size:9px;color:var(--accent2);margin-top:1px">TODAY</div>':'') + '</div>').join('') + '</div>' +
      (allUsers.length ? allUsers.map((user,idx) => {
        const ini = user.split(/[\s,]+/).filter(Boolean).map(x=>x[0].toUpperCase()).join('').slice(0,2);
        return '<div class="wv-row">' +
          '<div class="wv-name-cell"><div class="wv-av" style="background:' + COLORS[idx%COLORS.length] + '">' + ini + '</div><div class="wv-name" title="' + user + '">' + user + '</div></div>' +
          dates.map(d => {
            const st = getTeamStatus(user,d.dk);
            const pill = d.isWeekend&&!st ? '<span class="sp sp-we">wknd</span>' : st ? '<span class="sp sp-' + st.toLowerCase().replace(' ','-') + '">' + st + '</span>' : '<span class="sp sp-ns">—</span>';
            return '<div class="wv-cell' + (d.isToday?' today-col':'') + (d.isWeekend?' weekend-col':'') + '" data-dk="' + d.dk + '">' + pill + '</div>';
          }).join('') + '</div>';
      }).join('') : '<div style="padding:40px;text-align:center;color:var(--text3)"><p>No team data. Click Sync to load from SharePoint.</p></div>') +
      '</div><div id="wv-detail"></div>';

    const grid = el.querySelector('#att-grid');
    if (grid) grid.addEventListener('click', e => { const cell = e.target.closest('.wv-cell'); if (cell && cell.dataset.dk) wvShowDetail(cell.dataset.dk); });
    const syncBtn = el.querySelector('#att-sync-btn');
    if (syncBtn) syncBtn.addEventListener('click', () => { toast('Syncing...','info'); fetchAttendance(); });
  }

  function wvShowDetail(dk) {
    const detailEl = q('#wv-detail'); if (!detailEl) return;
    const allUsers = getAllMembers();
    const entries = allUsers.map(u => ({user:u, status:getTeamStatus(u,dk), process:getTeamProcess(u,dk), task:getTeamTask(u,dk)}));
    const marked   = entries.filter(e => e.status);
    const unmarked = entries.filter(e => !e.status).map(e => e.user);
    const grouped  = {};
    STATUSES.forEach(s => { grouped[s] = marked.filter(e => e.status===s); });

    detailEl.innerHTML =
      '<div class="wv-detail-panel"><div class="wv-dp-header">' +
      '<div class="wv-dp-title">📅 ' + formatDay(dk) + ' — Full Breakdown</div>' +
      '<button class="wv-dp-close" data-action="wv-close-detail">' + ic.close + '</button></div>' +
      (marked.length===0 ? '<div class="empty"><p>No attendance recorded for this date yet.</p></div>' :
        STATUSES.filter(s => grouped[s].length>0).map(s => {
          const cfg = STATUS_CFG[s];
          return '<div style="margin-bottom:16px">' +
            '<div style="font-size:.73rem;font-weight:700;text-transform:uppercase;letter-spacing:.6px;color:' + cfg.color + ';margin-bottom:8px;display:flex;align-items:center;gap:8px">' +
            '<span class="sp sp-' + s.toLowerCase().replace(' ','-') + '">' + s + '</span>' + cfg.label + ' — ' + grouped[s].length + ' person' + (grouped[s].length>1?'s':'') + '</div>' +
            '<div class="wv-dp-users">' +
            grouped[s].map(e => {
              const ini = e.user.split(/[\s,]+/).filter(Boolean).map(x=>x[0].toUpperCase()).join('').slice(0,2);
              return '<div class="wv-dp-user">' +
                '<div class="wv-av" style="background:' + cfg.color + '">' + ini + '</div>' +
                '<div class="wv-dp-uinfo"><div class="wv-dp-uname">' + e.user + '</div>' +
                (e.process?'<div class="wv-dp-uproc">📋 '+e.process+'</div>':'') +
                (e.task?'<div class="wv-dp-utask">📝 '+e.task+'</div>':'') + '</div>' +
                '<span class="sp sp-' + s.toLowerCase().replace(' ','-') + '">' + s + '</span></div>';
            }).join('') + '</div></div>';
        }).join('')) +
      (unmarked.length ? '<div style="margin-top:14px;padding-top:14px;border-top:1px solid var(--border)">' +
        '<div style="font-size:.73rem;font-weight:700;text-transform:uppercase;letter-spacing:.6px;color:var(--text3);margin-bottom:8px">⚠️ Not Marked — ' + unmarked.length + '</div>' +
        '<div style="display:flex;flex-wrap:wrap;gap:8px">' +
        unmarked.map(u => {
          const ini=u.split(/[\s,]+/).filter(Boolean).map(x=>x[0].toUpperCase()).join('').slice(0,2);
          return '<div style="display:flex;align-items:center;gap:6px;padding:5px 10px;border-radius:8px;background:var(--bg3);border:1px solid var(--border)">' +
            '<div class="wv-av" style="width:22px;height:22px;font-size:9px;background:var(--text3)">' + ini + '</div>' +
            '<span style="font-size:.8rem;color:var(--text2)">' + u + '</span></div>';
        }).join('') + '</div></div>' : '') +
      '</div>';
  }

  // ═══════════════════════════════════════════════════════════
  // ALL NPT LOG TAB (Admin — all members, filter + export)
  // ═══════════════════════════════════════════════════════════
  function renderAllNPT() {
    const el = q('#view-attnpt'); if (!el) return;
    const allNPT = Array.isArray(nptAllCache) ? nptAllCache : [];
    const filtered = nptFilterName ? allNPT.filter(n => (n.name||'').toLowerCase().includes(nptFilterName)) : allNPT;
    const totalMins = filtered.reduce((a,n) => a+(n.minutes||0), 0);
    const h = Math.floor(totalMins/60), m = totalMins%60;
    const thisMonth = new Date().toISOString().slice(0,7);
    const monthMins = filtered.filter(n => (n.date||'').startsWith(thisMonth)).reduce((a,n)=>a+(n.minutes||0),0);
    const members = getAllMembers();
    // Per-type breakdown
    const byType = {};
    NPT_TYPES.forEach(t => { byType[t] = filtered.filter(n=>n.type===t).reduce((a,n)=>a+(n.minutes||0),0); });
    const maxMins = Math.max(...Object.values(byType), 1);
    // Per-member breakdown
    const byMember = {};
    members.forEach(m2 => { byMember[m2] = filtered.filter(n=>n.name===m2).reduce((a,n)=>a+(n.minutes||0),0); });

    el.innerHTML =
      '<div class="ph"><div class="ph-left"><div class="ph-title">All NPT Log</div><div class="ph-sub">Non-productive time across team</div></div>' +
      '<div class="ph-actions">' +
      '<button class="btn btn-ghost btn-sm" data-action="sync-npt">' + ic.sync + ' Sync</button>' +
      '<button class="btn btn-ghost btn-sm" data-action="npt-export">' + ic.export + ' Export</button>' +
      '</div></div>' +
      '<div style="display:flex;gap:10px;flex-wrap:wrap;margin-bottom:14px">' +
      '<div style="flex:1;min-width:140px;background:var(--bg2);border:1px solid var(--border);border-radius:10px;padding:12px 16px"><div style="font-size:.73rem;color:var(--text3);text-transform:uppercase;letter-spacing:.5px;margin-bottom:4px">Total Entries</div><div style="font-size:1.5rem;font-weight:800;color:var(--text)">' + filtered.length + '</div></div>' +
      '<div style="flex:1;min-width:140px;background:var(--bg2);border:1px solid var(--border);border-radius:10px;padding:12px 16px"><div style="font-size:.73rem;color:var(--text3);text-transform:uppercase;letter-spacing:.5px;margin-bottom:4px">Total Time</div><div style="font-size:1.5rem;font-weight:800;color:var(--text)">' + (h>0?h+'h ':'')+m+'m</div></div>' +
      '<div style="flex:1;min-width:140px;background:var(--bg2);border:1px solid var(--border);border-radius:10px;padding:12px 16px"><div style="font-size:.73rem;color:var(--text3);text-transform:uppercase;letter-spacing:.5px;margin-bottom:4px">This Month</div><div style="font-size:1.5rem;font-weight:800;color:var(--text)">' + Math.floor(monthMins/60)+'h '+monthMins%60+'m</div></div>' +
      '<div style="flex:1;min-width:140px;background:var(--bg2);border:1px solid var(--border);border-radius:10px;padding:12px 16px"><div style="font-size:.73rem;color:var(--text3);text-transform:uppercase;letter-spacing:.5px;margin-bottom:4px">Members</div><div style="font-size:1.5rem;font-weight:800;color:var(--text)">' + new Set(filtered.map(n=>n.name)).size + '</div></div>' +
      '</div>' +
      '<div class="filter-bar">' +
      '<input type="text" class="dtr-input" id="npt-name-filter" placeholder="Filter by name..." value="' + (nptFilterName||'') + '" style="max-width:200px">' +
      (nptFilterName ? '<button class="btn btn-ghost btn-sm" data-action="npt-filter-clear">Clear</button>' : '') +
      '<span style="margin-left:auto;font-size:.82rem;color:var(--text3)">' + filtered.length + ' entries</span></div>' +
      '<div class="g2">' +
      '<div class="card"><div class="chart-title">By NPT Type (minutes)</div><div class="bar-rows">' +
      NPT_TYPES.filter(t=>byType[t]>0).map((t,i) => {
        const mins=byType[t], dh=Math.floor(mins/60), dm=mins%60;
        return '<div class="bar-row"><div class="bar-label">' + t + '</div><div class="bar-track"><div class="bar-fill" style="width:' + (mins/maxMins*100) + '%;background:' + COLORS[i%COLORS.length] + '"><span class="bar-fill-text">' + (dh>0?dh+'h ':'')+dm+'m</span></div></div></div>';
      }).join('') + (Object.values(byType).every(v=>v===0)?'<div class="empty"><p>No NPT data</p></div>':'') +
      '</div></div>' +
      '<div class="card"><div class="chart-title">By Member (minutes)</div><div class="bar-rows">' +
      members.filter(m2=>byMember[m2]>0).sort((a,b)=>byMember[b]-byMember[a]).map((m2,i) => {
        const mins=byMember[m2], dh=Math.floor(mins/60), dm=mins%60;
        return '<div class="bar-row"><div class="bar-label">' + m2.split(/[\s,]+/)[0] + '</div><div class="bar-track"><div class="bar-fill" style="width:' + (mins/Math.max(...Object.values(byMember),1)*100) + '%;background:' + COLORS[i%COLORS.length] + '"><span class="bar-fill-text">' + (dh>0?dh+'h ':'')+dm+'m</span></div></div></div>';
      }).join('') + (members.filter(m2=>byMember[m2]>0).length===0?'<div class="empty"><p>No NPT data</p></div>':'') +
      '</div></div></div>' +
      '<div class="card"><div class="card-title">Detailed NPT Entries</div>' +
      (filtered.length ? '<div class="tbl-wrap"><table class="dtr-table"><thead><tr><th>Date</th><th>Member</th><th>Type</th><th>Duration</th><th>Description</th></tr></thead><tbody>' +
        filtered.slice().reverse().slice(0,100).map(n => {
          const mins=n.minutes||0, dh=Math.floor(mins/60), dm=mins%60;
          return '<tr><td class="mono">' + n.date + '</td><td class="bold">' + (n.name||'—') + '</td>' +
            '<td><span class="badge bb">' + (n.type||'Other') + '</span></td>' +
            '<td class="mono">' + (dh>0?dh+'h ':'')+dm+'m</td>' +
            '<td style="max-width:220px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:var(--text3)" title="' + (n.desc||'').replace(/"/g,'&quot;') + '">' + (n.desc||'—') + '</td></tr>';
        }).join('') + '</tbody></table></div>' :
        '<div class="empty"><p>No NPT entries' + (nptFilterName?' for "'+nptFilterName+'"':'') + '.</p></div>') +
      '</div>';
  }

  function nptExport() {
    const allNPT = Array.isArray(nptAllCache)?nptAllCache:[];
    if (!allNPT.length) { toast('Nothing to export','err'); return; }
    dlCSV(['Member,Date,Type,Minutes,Description',...allNPT.map(n=>[n.name||'',n.date||'',n.type||'',n.minutes||0,n.desc||''].join(','))].join('\n'),'team_npt_log.csv');
    toast('Exported','ok');
  }

  // ═══════════════════════════════════════════════════════════
  // SETTINGS TAB
  // ═══════════════════════════════════════════════════════════
  function renderSettings() {
    const el = q('#view-settings'); if (!el) return;
    el.innerHTML =
      '<div class="ph"><div class="ph-left"><div class="ph-title">Settings</div></div></div>' +
      '<div class="card"><div class="card-title">' + ic.sync + ' SharePoint — Single Site</div>' +
      '<div class="info-banner">All data in: <strong>' + SP.SITE + '</strong></div>' +
      '<div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:10px">' +
      '<button class="btn btn-ghost btn-sm" data-action="test-sp">Test Connection</button>' +
      '<button class="btn btn-ghost btn-sm" data-action="sync-all">' + ic.sync + ' Sync All</button>' +
      '</div>' +
      '<div style="margin-top:14px;font-size:.78rem;color:var(--text3);line-height:2.2">' +
      '<strong>SharePoint Lists Required:</strong><br>' +
      '📋 Task List: <code style="color:var(--accent2)">' + SP.TASK_LIST + '</code><br>' +
      '📅 Attendance List: <code style="color:var(--accent2)">' + SP.STATUS_LIST + '</code>' +
      ' <span style="color:var(--text3)">(EmployeeName, StatusDate, WorkStatus, Process, TaskNotes, UpdatedAt)</span><br>' +
      '⏱ NPT List: <code style="color:var(--accent2)">' + SP.NPT_LIST + '</code>' +
      ' <span style="color:var(--text3)">(EmployeeName, NPTDate, NPTType, DurationMins, Description, LoggedAt)</span>' +
      '</div></div>' +
      '<div class="card"><div class="card-title">📊 Data Summary</div>' +
      '<div style="font-size:.85rem;color:var(--text2);line-height:2">' +
      'Team members detected: <strong>' + getAllMembers().length + '</strong><br>' +
      'Total task submissions: <strong>' + (Array.isArray(submissions)?submissions.length:0) + '</strong><br>' +
      'Attendance records: <strong>' + Object.keys(teamStatusCache).length + '</strong><br>' +
      'NPT entries: <strong>' + (Array.isArray(nptAllCache)?nptAllCache.length:0) + '</strong>' +
      '</div></div>' +
      '<div class="card"><div class="card-title">👤 Admin Info</div>' +
      '<div style="font-size:.85rem;color:var(--text2)">Logged in as: <strong style="color:var(--text)">' + ADMIN_NAME + '</strong></div>' +
      '<div style="font-size:.78rem;color:var(--text3);margin-top:6px">To change admin name, edit the ADMIN_NAME constant at the top of the script.</div></div>';
  }

  // ═══════════════════════════════════════════════════════════
  // SHAREPOINT
  // ═══════════════════════════════════════════════════════════
  async function getDigest() {
    return new Promise(r => { GM_xmlhttpRequest({method:'POST',url:SP.SITE+'/_api/contextinfo',headers:{'Accept':'application/json;odata=verbose','Content-Type':'application/json'},withCredentials:true,onload:res=>{try{r(JSON.parse(res.responseText).d.GetContextWebInformation.FormDigestValue);}catch{r(null);}},onerror:()=>r(null)}); });
  }

  function fetchAllData() { toast('Syncing all data...','info'); fetchTasks(); fetchAttendance(); fetchNPT(); }

  function fetchTasks() {
    GM_xmlhttpRequest({method:'GET',url:SP.SITE+'/_api/web/lists/GetByTitle(\''+SP.TASK_LIST+'\')/items?$top=5000&$orderby=TaskDate%20desc',headers:{'Accept':'application/json;odata=verbose'},withCredentials:true,
    onload:res=>{try{
      const items=JSON.parse(res.responseText).d.results||[];
      submissions=items.map(it=>({employeeName:it.EmployeeName||'',taskType:it.TaskType||'',hours:it.HoursWorked||0,npt:it.NPTHours||0,workType:it.WorkType||'Productive',adhoc:it.AdHocDetails||'',date:(it.TaskDate||'').split('T')[0],submittedAt:it.SubmittedAt||it.Created}));
      safeSave('dtr_subs4',submissions);
      updateSBStats();
      if(currentView==='overview')renderOverview();
      if(currentView==='teamrpt')renderTeamReport();
      if(currentView==='rankings')renderRankings();
      toast('Tasks synced ✓','ok');
    }catch(e){toast('Task sync error','err');}},
    onerror:()=>toast('SP unreachable','err')});
  }

  function fetchAttendance() {
    GM_xmlhttpRequest({method:'GET',url:SP.SITE+'/_api/web/lists/GetByTitle(\''+SP.STATUS_LIST+'\')/items?$top=5000&$orderby=StatusDate%20desc&$select=EmployeeName,StatusDate,WorkStatus,Process,TaskNotes',headers:{'Accept':'application/json;odata=verbose'},withCredentials:true,
    onload:res=>{try{
      const items=JSON.parse(res.responseText).d.results||[];
      teamStatusCache={};
      items.forEach(it=>{const name=it.EmployeeName||'',dk=(it.StatusDate||'').split('T')[0];if(name&&dk)teamStatusCache[name+'::'+dk]={status:it.WorkStatus||'',process:it.Process||'',task:it.TaskNotes||''};});
      safeSaveObj('dtr_teamcache',teamStatusCache);
      if(currentView==='attweek')renderAttWeek();
      if(currentView==='overview')renderOverview();
      toast('Attendance synced ✓','ok');
    }catch(e){toast('Attendance sync error','err');}},
    onerror:()=>toast('SP unreachable','err')});
  }

  function fetchNPT() {
    GM_xmlhttpRequest({method:'GET',url:SP.SITE+'/_api/web/lists/GetByTitle(\''+SP.NPT_LIST+'\')/items?$top=5000&$orderby=NPTDate%20desc&$select=EmployeeName,NPTDate,NPTType,DurationMins,Description',headers:{'Accept':'application/json;odata=verbose'},withCredentials:true,
    onload:res=>{try{
      const items=JSON.parse(res.responseText).d.results||[];
      nptAllCache=items.map(it=>({name:it.EmployeeName||'',date:(it.NPTDate||'').split('T')[0],type:it.NPTType||'',minutes:it.DurationMins||0,desc:it.Description||''}));
      safeSave('dtr_npt2',nptAllCache);
      if(currentView==='attnpt')renderAllNPT();
      toast('NPT log synced ✓','ok');
    }catch(e){toast('NPT sync error','err');}},
    onerror:()=>toast('SP unreachable','err')});
  }

  async function testSP() {
    toast('Testing...','info');
    try { const r=await fetch(SP.SITE+'/_api/contextinfo',{method:'POST',credentials:'include',headers:{'Accept':'application/json;odata=verbose','Content-Type':'application/json;odata=verbose'}}); toast(r.ok?'✅ SP active':'❌ SP returned '+r.status,r.ok?'ok':'err'); } catch { toast('❌ Cannot reach SP','err'); }
  }

  // ═══════════════════════════════════════════════════════════
  // DATA HELPERS
  // ═══════════════════════════════════════════════════════════
  function getAllMembers() {
    const fromSubs = [...new Set((Array.isArray(submissions)?submissions:[]).map(s=>s.employeeName).filter(Boolean))];
    const fromAtt  = [...new Set(Object.keys(teamStatusCache).map(k=>k.split('::')[0]).filter(Boolean))];
    const fromNPT  = [...new Set((Array.isArray(nptAllCache)?nptAllCache:[]).map(n=>n.name).filter(Boolean))];
    const all = new Set([...fromSubs,...fromAtt,...fromNPT]);
    return [...all].sort();
  }
  function getTeamStatus(user,dk)  { return(teamStatusCache[user+'::'+dk]||{}).status||''; }
  function getTeamProcess(user,dk) { return(teamStatusCache[user+'::'+dk]||{}).process||''; }
  function getTeamTask(user,dk)    { return(teamStatusCache[user+'::'+dk]||{}).task||''; }
  function calcAvgProd(subs){ if(!Array.isArray(subs))return 0;const v=subs.filter(s=>s&&!( s.taskType||'').startsWith('Leave'));if(!v.length)return 0;return Math.min(100,v.reduce((a,s)=>a+((s.hours||0)/WH*100),0)/v.length); }
  function calcTeamAvgProd(subs){ if(!subs.length)return 0;const members=[...new Set(subs.map(s=>s.employeeName).filter(Boolean))];if(!members.length)return 0;return members.reduce((a,m)=>a+calcAvgProd(subs.filter(s=>s.employeeName===m)),0)/members.length; }
  function getWeekStart(offset){ const d=new Date();d.setHours(0,0,0,0);d.setDate(d.getDate()-d.getDay()+(offset*7));return d; }
  function getWeekLabel(offset){ const ws=getWeekStart(offset),we=new Date(ws);we.setDate(ws.getDate()+6);const f=d=>d.toLocaleDateString('en-US',{month:'short',day:'numeric'});return f(ws)+' – '+f(we)+', '+we.getFullYear(); }
  function todayStr() { return new Date().toISOString().split('T')[0]; }
  function formatDay(d){ return new Date(d+'T12:00:00').toLocaleDateString('en-US',{weekday:'long',year:'numeric',month:'long',day:'numeric'}); }
  function dlCSV(csv,fn){ const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([csv],{type:'text/csv'}));a.download=fn;a.click(); }
  function updateSBStats(){ const m=getAllMembers(),s=q('#sb-members'),t=q('#sb-total');if(s)s.textContent=m.length;if(t)t.textContent=(Array.isArray(submissions)?submissions.length:0); }
  function toast(msg,type='info'){ let el=document.getElementById('dtr-toast');if(!el){el=document.createElement('div');el.id='dtr-toast';(document.getElementById('dtr-root-outer')||root).appendChild(el);}const icons={ok:'✅',err:'❌',info:'💡'};el.innerHTML='<span>'+(icons[type]||'💡')+'</span><span>'+msg+'</span>';el.className='show t'+type;clearTimeout(el._t);el._t=setTimeout(()=>{el.className='';},4000); }
  function prog_wrap_style(){ return ''; } // placeholder kept for CSS ref

  // INIT
  if (document.readyState==='loading') { document.addEventListener('DOMContentLoaded', boot); } else { boot(); }

})();
