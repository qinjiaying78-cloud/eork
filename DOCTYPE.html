<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>finance workbench · 全流程财务与穿透联查工作台</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/gsap/3.12.2/gsap.min.js"></script>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&family=Playfair+Display:ital,wght@0,400;0,600;1,400&display=swap" rel="stylesheet">

<style>
/* ================= 核心设计系统 ================= */
:root {
  --bg: #F5F4ED; --card: #FFFFFF; --border: #E8E6DC;
  --text-main: #141413; --text-muted: #6B6A65;
  --primary: #D97757; --primary-hover: #C46545;
  --accent-blue: #6B8CBA; --input-bg: #F9F9F6; --radius: 12px;
  --font-serif: "Playfair Display", "Georgia", "Tiempos Headline", serif;
  --font-sans: "Inter", -apple-system, BlinkMacSystemFont, "PingFang SC", "Microsoft YaHei", sans-serif;
  --match: #5D9C7A; --match-bg: #F0FDF4;
  --bank-only: #E0A25E; --bank-only-bg: #FFF7ED;
  --company-only: #6B8CBA; --company-only-bg: #EFF6FF;
  --diff: #D96262; --diff-bg: #FEF2F2;
  --error: #A3A3A3; --error-bg: #F8FAFC;
  --sidebar-width: 260px;
  --header-height: 64px;
}

* { box-sizing: border-box; margin: 0; padding: 0; }
body {
  font-family: var(--font-sans); background-color: var(--bg); color: var(--text-main);
  line-height: 1.6; min-height: 100vh; display: flex; flex-direction: column;
  -webkit-font-smoothing: antialiased;
}
h1, h2, h3, .serif-title { font-family: var(--font-serif); font-weight: 400; letter-spacing: -0.5px; }
p { color: var(--text-muted); }

/* ================= 顶部栏 ================= */
.app-header {
  height: var(--header-height);
  padding: 0 28px; display: flex; justify-content: space-between; align-items: center;
  border-bottom: 1px solid var(--border); background-color: var(--bg);
  position: sticky; top: 0; z-index: 200; gap: 16px;
}
.header-left { display: flex; align-items: center; gap: 12px; }
.sidebar-toggle-btn {
  display: none; align-items: center; justify-content: center;
  width: 36px; height: 36px; border-radius: var(--radius);
  border: 1px solid var(--border); background: var(--card); color: var(--text-main);
  cursor: pointer; font-size: 18px;
}
.app-logo {
  font-family: var(--font-serif); font-size: 22px; font-weight: 600;
  display: flex; gap: 10px; align-items: center; cursor: pointer; white-space: nowrap;
}
.app-logo span { color: var(--primary); font-size: 24px; }

/* ================= 左侧树形导航 + 右侧宽屏工作区 ================= */
.app-shell {
  display: flex; flex: 1; min-height: calc(100vh - var(--header-height));
  position: relative;
}
.app-sidebar {
  width: var(--sidebar-width); flex-shrink: 0;
  background: var(--card); border-right: 1px solid var(--border);
  display: flex; flex-direction: column;
  position: sticky; top: var(--header-height);
  height: calc(100vh - var(--header-height));
  overflow-y: auto; z-index: 150;
  transition: transform 0.25s ease;
}
.sidebar-header {
  padding: 16px 18px 12px;
  border-bottom: 1px solid var(--border);
  display: flex; align-items: center; justify-content: space-between;
}
.sidebar-header-title {
  font-size: 11px; font-weight: 600; letter-spacing: 1.5px;
  text-transform: uppercase; color: var(--text-muted);
}
.sidebar-home-btn {
  width: calc(100% - 24px); margin: 12px 12px 6px;
  padding: 10px 14px; border-radius: var(--radius);
  border: 1px solid transparent; background: transparent;
  display: flex; align-items: center; gap: 10px;
  font-size: 13px; font-weight: 500; color: var(--text-main);
  cursor: pointer; transition: all 0.2s ease; text-align: left;
}
.sidebar-home-btn:hover { background: var(--bg); }
.sidebar-home-btn.active {
  background: #FFF9F6; border-color: var(--primary); color: var(--primary); font-weight: 600;
}
.sidebar-tree { padding: 6px 12px 24px; flex: 1; }
.tree-group { margin-bottom: 4px; }
.tree-group-btn {
  width: 100%; padding: 10px 12px; border: none; background: transparent;
  border-radius: var(--radius); cursor: pointer;
  display: flex; align-items: center; justify-content: space-between; gap: 8px;
  font-size: 13px; font-weight: 600; color: var(--text-main);
  transition: background 0.2s ease; text-align: left;
}
.tree-group-btn:hover { background: var(--bg); }
.tree-group.open > .tree-group-btn { color: var(--primary); }
.tree-group-label { display: flex; align-items: center; gap: 8px; }
.tree-group-num {
  display: inline-flex; align-items: center; justify-content: center;
  width: 20px; height: 20px; border-radius: 6px;
  background: var(--bg); border: 1px solid var(--border);
  font-size: 11px; font-weight: 600; color: var(--text-muted); flex-shrink: 0;
}
.tree-group.open .tree-group-num {
  background: #FFF9F6; border-color: var(--primary); color: var(--primary);
}
.tree-chevron {
  font-size: 11px; color: var(--text-muted); transition: transform 0.2s ease;
}
.tree-group.open .tree-chevron { transform: rotate(90deg); color: var(--primary); }
.tree-children {
  display: none; padding: 2px 0 6px 28px;
  flex-direction: column; gap: 2px;
}
.tree-group.open .tree-children { display: flex; }
.tree-item-btn {
  width: 100%; padding: 8px 12px; border: none; background: transparent;
  border-radius: 8px; cursor: pointer; text-align: left;
  font-size: 13px; color: var(--text-muted);
  display: flex; align-items: center; justify-content: space-between; gap: 6px;
  transition: all 0.15s ease;
}
.tree-item-btn:hover { background: var(--bg); color: var(--text-main); }
.tree-item-btn.active {
  background: var(--bg); color: var(--primary); font-weight: 600;
  border-left: 2px solid var(--primary);
}
.tree-item-tag {
  font-size: 10px; padding: 1px 6px; border-radius: 8px;
  background: var(--bg); color: var(--text-muted); border: 1px solid var(--border);
}
.sidebar-backdrop {
  display: none; position: fixed; inset: 0; background: rgba(20, 20, 19, 0.35);
  z-index: 140;
}
.sidebar-backdrop.open { display: block; }

/* 右侧宽屏工作区 & 多标签页 */
.app-main {
  flex: 1; min-width: 0; display: flex; flex-direction: column;
  background: var(--bg);
}
.workspace-tabs-bar {
  display: flex; align-items: center; justify-content: space-between; gap: 12px;
  padding: 8px 24px 0; background: var(--bg);
  border-bottom: 1px solid var(--border);
  position: sticky; top: var(--header-height); z-index: 90;
}
.workspace-tabs-scroll {
  display: flex; align-items: center; gap: 6px; overflow-x: auto; flex: 1;
  scrollbar-width: thin;
}
.ws-tab {
  display: inline-flex; align-items: center; gap: 8px;
  padding: 9px 14px; border-radius: var(--radius) var(--radius) 0 0;
  border: 1px solid transparent; border-bottom: none;
  background: transparent; color: var(--text-muted);
  font-size: 13px; font-weight: 500; cursor: pointer; white-space: nowrap;
  transition: all 0.15s ease; user-select: none;
}
.ws-tab:hover { color: var(--text-main); background: rgba(255,255,255,0.6); }
.ws-tab.active {
  background: var(--card); color: var(--text-main); font-weight: 600;
  border-color: var(--border); margin-bottom: -1px; padding-bottom: 10px;
}
.ws-tab-dot {
  width: 6px; height: 6px; border-radius: 50%; background: var(--border);
}
.ws-tab.active .ws-tab-dot { background: var(--primary); }
.ws-tab-close {
  display: inline-flex; align-items: center; justify-content: center;
  width: 18px; height: 18px; border-radius: 50%; border: none;
  background: transparent; color: var(--text-muted); font-size: 13px;
  cursor: pointer; line-height: 1;
}
.ws-tab-close:hover { background: var(--diff-bg); color: var(--diff); }
.workspace-tabs-actions {
  padding-bottom: 6px; display: flex; align-items: center; gap: 6px; flex-shrink: 0;
}

.workspace-content { flex: 1; width: 100%; }
.module { display: none; padding: 32px 40px; max-width: 1440px; margin: 0 auto; width: 100%; opacity: 0; }
.module.active { display: block; opacity: 1; }

/* 占位提示页 */
.placeholder-card {
  background: var(--card); border: 1px solid var(--border); border-radius: var(--radius);
  padding: 56px 40px; text-align: center; max-width: 680px; margin: 32px auto;
  box-shadow: 0 4px 24px rgba(0,0,0,0.02);
}
.placeholder-badge {
  display: inline-block; padding: 4px 12px; border-radius: 12px;
  background: #FFF9F6; border: 1px solid var(--primary); color: var(--primary);
  font-size: 12px; font-weight: 500; margin-bottom: 16px;
}
.placeholder-card h2 { font-size: 30px; margin-bottom: 12px; color: var(--text-main); }
.placeholder-card p { font-size: 14px; line-height: 1.7; margin-bottom: 24px; }
.placeholder-actions { display: flex; justify-content: center; gap: 12px; flex-wrap: wrap; }

/* 退出/备份确认模态框 */
.modal-overlay {
  display: none; position: fixed; inset: 0; z-index: 1200;
  background: rgba(20, 20, 19, 0.45); backdrop-filter: blur(2px);
  align-items: center; justify-content: center; padding: 20px;
}
.modal-overlay.open { display: flex; }
.modal-card {
  background: var(--card); border: 1px solid var(--border); border-radius: var(--radius);
  width: min(460px, 100%); padding: 28px;
  box-shadow: 0 20px 48px rgba(0,0,0,0.14);
}
.modal-title { font-family: var(--font-serif); font-size: 22px; margin-bottom: 10px; color: var(--text-main); }
.modal-desc { font-size: 14px; color: var(--text-muted); line-height: 1.7; margin-bottom: 24px; }
.modal-actions { display: flex; justify-content: flex-end; gap: 10px; flex-wrap: wrap; }

/* 通用按钮与表单 */
.btn {
  padding: 10px 24px; border-radius: var(--radius); border: none; cursor: pointer;
  font-weight: 500; font-size: 14px; font-family: var(--font-sans);
  transition: all 0.2s ease; background: var(--card); color: var(--text-main); border: 1px solid var(--border);
}
.btn:hover { background: var(--bg); }
.btn-primary { background: var(--primary); color: #FFF; border-color: var(--primary); }
.btn-primary:hover { background: var(--primary-hover); border-color: var(--primary-hover); }
.btn:disabled { opacity: 0.5; cursor: not-allowed; }

input, select, textarea {
  width: 100%; padding: 12px 16px; background: var(--input-bg);
  border: 1px solid var(--border); border-radius: var(--radius);
  font-size: 14px; color: var(--text-main); font-family: var(--font-sans);
  outline: none; transition: border-color 0.2s ease;
}
input:focus, select:focus, textarea:focus { border-color: var(--primary); background: var(--card); }

#toast {
  visibility: hidden; min-width: 250px; background: var(--text-main); color: var(--bg);
  text-align: center; border-radius: var(--radius); padding: 16px;
  position: fixed; z-index: 4000; left: 50%; bottom: 30px;
  font-size: 14px; opacity: 0; box-shadow: 0 10px 30px rgba(0,0,0,0.1);
  transform: translateX(-50%);
}

button { font-family: var(--font-sans); }
button:focus-visible { outline: 2px solid var(--primary); outline-offset: 3px; }
button:disabled { opacity: .55; cursor: not-allowed; }
.btn-danger { color: var(--diff); }
.btn-danger:hover { background: var(--diff-bg); border-color: var(--diff); }
.btn-icon { min-width: 28px; min-height: 28px; border-radius: 6px; line-height: 1; }
#module-invoice .btn-del:hover, .attach-item .btn-icon:hover,
#module-voucher .voucher-entry-table .btn-remove-row:hover { background: var(--diff-bg); }
.workbench-actions { display:flex; align-items:center; gap:8px; flex-wrap:wrap; }
.workbench-actions .btn { padding:8px 14px; font-size:13px; white-space:nowrap; }
.workbench-actions .btn-icon-label { display:inline-flex; align-items:center; gap:7px; }
.workbench-actions svg { width:16px; height:16px; }
.draft-status { font-size:12px; color:var(--text-muted); }

/* 模块操作区 */
.section-actions { display:flex; align-items:center; flex-wrap:wrap; gap:10px; }
.section-actions .btn { white-space:nowrap; }
.recon-actions { justify-content:center; margin-top:28px; }
.recon-summary { text-align:center; margin:12px 0 0; }
.recon-summary:empty { display:none; }
.tax-task-actions {
  grid-column:1 / -1; justify-content:flex-end;
  padding-top:14px; border-top:1px solid var(--border);
}
#task-file { min-width:0; padding:6px; color:var(--text-muted); }
#task-file::file-selector-button {
  font:inherit; color:var(--text-main); background:var(--card);
  border:1px solid var(--border); border-radius:calc(var(--radius) - 4px);
  margin:0 10px 0 0; padding:7px 12px; cursor:pointer;
  transition:background .2s, border-color .2s;
}
#task-file::file-selector-button:hover { background:var(--bg); border-color:var(--primary); }
#task-file:focus-visible { outline:2px solid var(--primary); outline-offset:2px; }
.closing-actions { margin-top:20px; }

/* 首页 */
.home-container { display: flex; flex-direction: column; justify-content: center; min-height: 70vh; max-width: 1200px; margin: 0 auto; }
.home-header-text { margin-bottom: 40px; }
.home-subtitle-top {
  font-size: 12px; font-weight: 600; text-transform: uppercase; letter-spacing: 2px;
  color: var(--primary); margin-bottom: 14px;
}
.home-title { font-size: 64px; color: var(--text-main); line-height: 1.05; margin-bottom: 20px; font-weight: 400; letter-spacing: -1.5px; }
.home-title i { font-style: italic; font-weight: 400; }
.home-desc { font-size: 17px; color: var(--text-muted); max-width: 680px; line-height: 1.7; }

.home-cards { display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 22px; width: 100%; }
.home-card {
  background: var(--card); border: 1px solid var(--border); border-radius: 4px;
  padding: 28px; cursor: pointer; transition: all 0.3s ease;
  display: flex; flex-direction: column; position: relative; overflow: hidden;
  box-shadow: 0 4px 24px rgba(0,0,0,0.02); min-height: 190px;
}
.home-card:hover { border-color: var(--primary); transform: translateY(-4px); box-shadow: 0 12px 32px rgba(0,0,0,0.06); }
.home-card-content { flex: 1; z-index: 2; position: relative; }
.home-card h3 { font-size: 23px; color: var(--text-main); margin-bottom: 8px; }
.home-card p { font-size: 13px; margin-bottom: 22px; max-width: 90%; line-height: 1.6; }
.home-card .btn-enter {
  background: var(--text-main); color: #fff; padding: 8px 16px; border-radius: 4px;
  border: none; font-weight: 500; cursor: pointer; transition: background 0.2s; font-size: 13px;
}
.home-card:hover .btn-enter { background: var(--primary); }
.card-svg-bg {
  position: absolute; bottom: 10px; right: 10px; width: 104px; height: 104px;
  z-index: 1; opacity: 0.85; transition: transform 0.5s cubic-bezier(0.4, 0, 0.2, 1);
}
.home-card:hover .card-svg-bg { transform: scale(1.08) rotate(-3deg); }

/* 表格通用 */
.data-table { width: 100%; border-collapse: collapse; }
.data-table th, .data-table td { padding: 12px 14px; border-bottom: 1px solid var(--border); font-size: 13px; text-align: left; }
.data-table th { background: var(--bg); color: var(--text-muted); font-weight: 500; font-size: 12px; }
.data-table tbody tr:hover { background: #FAFAF7; }

/* 模块通用 */
.panel { background: var(--card); border: 1px solid var(--border); border-radius: var(--radius); padding: 28px; margin-bottom: 24px; }
.module-header { margin-bottom: 24px; display: flex; justify-content: space-between; align-items: flex-end; flex-wrap: wrap; gap: 16px; }
.module-header h1 { font-size: 36px; margin-bottom: 8px; }
.toolbar { display: flex; gap: 12px; align-items: center; flex-wrap: wrap; margin-bottom: 20px; background: var(--bg); padding: 16px; border-radius: var(--radius); border: 1px solid var(--border); }
.form-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 14px; margin-bottom: 14px; }
.form-group label { display: block; margin-bottom: 6px; font-size: 12px; color: var(--text-muted); }

/* 状态标签 */
.status-pill { display: inline-block; padding: 3px 10px; border-radius: 12px; font-size: 12px; font-weight: 500; white-space: nowrap; }
.status-pill.PENDING { background: #FFF7ED; color: #C2410C; }
.status-pill.DONE { background: #EFF6FF; color: #1D4ED8; }
.status-pill.POSTED { background: #F0FDF4; color: #15803D; }
.status-pill.REJECTED { background: #FEE2E2; color: #B91C1C; }
.status-pill.WARNING { background: #FEF3C7; color: #B45309; }

.btn-mini { padding: 5px 12px; border-radius: 6px; font-size: 12px; border: 1px solid var(--border); background: var(--card); cursor: pointer; color: var(--text-main); font-family: var(--font-sans); transition: 0.2s; }
.btn-mini:hover { background: var(--bg); }
.btn-mini-primary { background: var(--primary); color: #fff; border-color: var(--primary); }
.btn-mini-primary:hover { background: var(--primary-hover); }

/* ================= 融合增强：七列余额表、穿透联查与上下文菜单 ================= */
.hint-bar {
  display: flex; align-items: center; justify-content: space-between; gap: 12px; flex-wrap: wrap;
  background: #FFF9F6; border: 1px dashed var(--primary); color: var(--primary);
  border-radius: 10px; padding: 10px 16px; font-size: 12.5px; margin-bottom: 18px;
}
.fin-table { width: 100%; border-collapse: collapse; }
.fin-table th, .fin-table td { padding: 11px 14px; border-bottom: 1px solid var(--border); font-size: 13px; text-align: left; }
.fin-table th { background: var(--bg); color: var(--text-muted); font-weight: 500; font-size: 12px; white-space: nowrap; }
.fin-table td.num, .fin-table th.num {
  text-align: right;
  font-variant-numeric: tabular-nums; font-feature-settings: "tnum";
  font-family: ui-monospace, "SF Mono", "JetBrains Mono", Consolas, monospace;
  white-space: nowrap;
}
.fin-table tbody tr { transition: background .15s; }
.fin-table tbody tr:hover { background: #FFF9F6; }
.fin-table tbody tr.clickable { cursor: context-menu; }
.fin-table tbody tr.opening-row { background: #FAFAF5; color: var(--text-muted); font-style: italic; }
.fin-table tbody tr.total-row { background: #FAFAF5; font-weight: 600; }
.fin-table tfoot td { background: var(--bg); font-weight: 600; font-size: 13px; }
.fin-table tfoot td.num { font-variant-numeric: tabular-nums; font-family: ui-monospace, Consolas, monospace; }
.fin-table .code { font-family: ui-monospace, Consolas, monospace; color: var(--text-muted); }
.fin-table .strong { color: var(--text-main); font-weight: 600; }
.trace-tag {
  display: inline-block; margin-left: 8px; padding: 0 7px; border-radius: 8px;
  font-size: 10px; line-height: 17px; font-weight: 500;
  background: #FFF9F6; border: 1px solid var(--primary); color: var(--primary);
  vertical-align: middle;
}
.table-footnote { margin-top: 12px; font-size: 12px; color: var(--text-muted); line-height: 1.7; }
.dir-badge { display: inline-block; padding: 1px 8px; border-radius: 8px; font-size: 11px; font-weight: 600; }
.dir-badge.dr { background: #EFF6FF; color: #1D4ED8; }
.dir-badge.cr { background: #F0FDF4; color: #15803D; }

/* 右键上下文菜单 */
.ctx-menu {
  position: fixed; z-index: 3000; min-width: 240px; max-width: 320px;
  background: var(--card); border: 1px solid var(--border); border-radius: 10px;
  padding: 6px; display: none; box-shadow: 0 16px 40px rgba(0,0,0,.16);
  font-size: 13px; user-select: none;
}
.ctx-menu.open { display: block; animation: ctxIn .12s ease both; }
@keyframes ctxIn { from { opacity: 0; transform: scale(.97); } to { opacity: 1; transform: none; } }
.ctx-title {
  padding: 8px 12px 9px; font-size: 11.5px; font-weight: 700; letter-spacing: .4px;
  color: var(--text-muted); border-bottom: 1px solid var(--border); margin-bottom: 6px;
  white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
}
.ctx-item {
  display: flex; align-items: center; gap: 9px; padding: 8px 12px;
  border-radius: 8px; cursor: pointer; color: var(--text-main); transition: background .12s;
}
.ctx-item:hover { background: #FFF9F6; color: var(--primary); }
.ctx-icon { width: 16px; text-align: center; opacity: .85; flex-shrink: 0; }
.ctx-item.ctx-sub { padding-left: 30px; font-size: 12.5px; color: var(--text-muted); }
.ctx-item.ctx-sub:hover { color: var(--primary); }
.ctx-divider { height: 1px; background: var(--border); margin: 6px 6px; }

/* 穿透联查侧滑抽屉 */
.drawer-mask { position: fixed; inset: 0; background: rgba(20,20,19,.38); z-index: 2500; display: none; }
.drawer-mask.open { display: block; }
.drawer-panel {
  position: absolute; right: 0; top: 0; height: 100%; width: min(940px, 96vw);
  background: var(--card); display: flex; flex-direction: column;
  box-shadow: -14px 0 46px rgba(0,0,0,.16); animation: slideIn .26s cubic-bezier(.2,.8,.3,1) both;
}
@keyframes slideIn { from { transform: translateX(48px); opacity: .5; } to { transform: none; opacity: 1; } }
.drawer-head {
  padding: 22px 28px; border-bottom: 1px solid var(--border);
  display: flex; justify-content: space-between; align-items: flex-start; gap: 16px;
}
.drawer-head h2 { font-size: 24px; margin-bottom: 4px; }
.drawer-sub { font-size: 12.5px; color: var(--text-muted); margin: 0; }
.drawer-body { padding: 24px 28px 40px; overflow: auto; flex: 1; }

/* 银行账户卡片与汇总条 */
.bank-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(260px, 1fr)); gap: 16px; margin-bottom: 24px; }
.bank-card {
  border: 1px solid var(--border); border-radius: var(--radius); padding: 18px 20px;
  background: var(--input-bg); position: relative; overflow: hidden;
}
.bank-card::after {
  content: ''; position: absolute; right: -30px; top: -30px; width: 90px; height: 90px;
  border-radius: 50%; background: var(--primary); opacity: .07;
}
.bank-card h4 { font-size: 14px; margin-bottom: 4px; color: var(--text-main); font-weight: 600; font-family: var(--font-sans); }
.bank-card .acct { font-family: ui-monospace, Consolas, monospace; font-size: 12px; color: var(--text-muted); letter-spacing: .4px; margin-bottom: 14px; }
.bank-card .bal { font-size: 22px; font-family: var(--font-serif); color: var(--text-main); }
.bank-card .bal-label { font-size: 11px; color: var(--text-muted); letter-spacing: 1px; text-transform: uppercase; margin-bottom: 2px; }
.bank-meta { display: flex; gap: 18px; margin-top: 12px; font-size: 11.5px; color: var(--text-muted); }
.bank-meta b { color: var(--match); font-weight: 600; }
.bank-meta .out b { color: var(--diff); }
.sum-bar { display: flex; gap: 26px; flex-wrap: wrap; padding: 14px 18px; background: var(--bg); border-radius: var(--radius); margin-bottom: 18px; font-size: 12.5px; }
.sum-bar span { color: var(--text-muted); }
.sum-bar strong { color: var(--text-main); font-size: 15px; font-family: ui-monospace, Consolas, monospace; margin-left: 6px; }

/* 模式切换胶囊 */
.mode-switch { display: inline-flex; background: var(--bg); padding: 4px; border-radius: 10px; border: 1px solid var(--border); gap: 4px; }
.mode-switch-btn {
  padding: 7px 14px; border: none; background: transparent; border-radius: 8px;
  font-size: 13px; font-weight: 500; color: var(--text-muted); cursor: pointer; transition: .2s;
}
.mode-switch-btn.active { background: var(--card); color: var(--text-main); box-shadow: 0 1px 3px rgba(0,0,0,.06); font-weight: 600; }

/* 对账 */
#module-recon .upload-container { display: flex; gap: 32px; margin-bottom: 40px; }
#module-recon .drop-zone {
  flex: 1; border: 1.5px dashed var(--border); border-radius: var(--radius);
  padding: 48px; text-align: center; background: var(--card); cursor: pointer;
  transition: all 0.2s ease;
}
#module-recon .drop-zone:hover { border-color: var(--primary); }
#module-recon .mapping-area { display: none; background: var(--card); padding: 32px; border-radius: var(--radius); margin-bottom: 32px; border: 1px solid var(--border); }
#module-recon .mapping-row { display: flex; gap: 24px; margin-bottom: 16px; }
#module-recon .mapping-group { flex: 1; display: flex; align-items: center; gap: 12px; }
#module-recon .result-area { display: none; background: var(--card); border-radius: var(--radius); border: 1px solid var(--border); overflow: hidden; margin-top: 32px; }
#module-recon .tabs { display: flex; border-bottom: 1px solid var(--border); padding: 0 24px; gap: 8px; overflow-x: auto; background: var(--bg); }
#module-recon .tab { padding: 16px 12px; cursor: pointer; border-bottom: 2px solid transparent; font-size: 14px; font-weight: 500; color: var(--text-muted); }
#module-recon .tab.active { border-bottom-color: var(--primary); color: var(--text-main); }
#module-recon .tab .badge { display: inline-block; padding: 2px 8px; border-radius: 12px; font-size: 12px; color: #fff; margin-left: 6px; }
#recon-result-table { width: 100%; border-collapse: collapse; }
#recon-result-table th, #recon-result-table td { padding: 12px 14px; border-bottom: 1px solid var(--border); font-size: 13px; text-align: left; }
#recon-result-table th { background: var(--bg); color: var(--text-muted); font-weight: 500; }

/* 开票 */
#module-invoice .layout { display: flex; gap: 40px; align-items: flex-start; }
#module-invoice .form-panel { flex: 1.2; background: var(--card); padding: 40px; border-radius: var(--radius); border: 1px solid var(--border); }
#module-invoice .invoice-panel { flex: 1; background: var(--card); padding: 48px; border-radius: var(--radius); border: 1px solid var(--border); position: relative; box-shadow: 0 8px 32px rgba(0,0,0,0.04); }
#module-invoice h2 { font-size: 24px; margin-bottom: 24px; }
#module-invoice .form-group { margin-bottom: 20px; }
#module-invoice .form-group label { display: block; margin-bottom: 8px; font-size: 13px; color: var(--text-muted); }
.additional-info { background: var(--bg); padding: 24px; border-radius: var(--radius); margin-bottom: 24px; margin-top: 40px; border: 1px solid var(--border); }
.additional-info h3 { font-size: 18px; margin-bottom: 16px; }
.upload-zone-mini { border: 1px dashed var(--border); border-radius: var(--radius); padding: 16px; text-align: center; cursor: pointer; background: var(--card); margin-bottom: 16px; }
.upload-zone-mini:hover { border-color: var(--primary); }
.attachment-list { display: flex; flex-wrap: wrap; gap: 8px; margin-bottom: 16px; }
.attach-item { background: var(--card); border: 1px solid var(--border); padding: 6px 12px; border-radius: 20px; font-size: 12px; display: flex; align-items: center; gap: 8px; }
.attach-item button { background: transparent; border: none; color: var(--diff); cursor: pointer; font-size: 14px; }
.additional-info textarea { resize: vertical; min-height: 100px; }
#module-invoice .items-table { width: 100%; border-collapse: collapse; margin-bottom: 24px; table-layout: fixed; }
#module-invoice .items-table th { padding: 12px 8px; text-align: left; font-size: 13px; color: var(--text-muted); border-bottom: 1px solid var(--border); font-weight: 500; }
#module-invoice .items-table td { padding: 8px 4px; border-bottom: 1px solid var(--border); }
#module-invoice .items-table input { border: none; border-radius: 0; padding: 8px; background: transparent; border-bottom: 1px solid transparent; }
#module-invoice .items-table input:focus { border-bottom-color: var(--primary); }
#module-invoice .btn-add { background: var(--bg); color: var(--text-main); width: 100%; padding: 12px; margin-bottom: 24px; border-radius: var(--radius); border: 1px solid var(--border); cursor: pointer; font-weight: 500; }
#module-invoice .btn-submit { width: 100%; padding: 16px; font-size: 16px; background: var(--primary); color: #fff; border: none; border-radius: var(--radius); cursor: pointer; font-weight: 500; }
#module-invoice .btn-del { background: transparent; color: var(--diff); border: none; cursor: pointer; }
#module-invoice .invoice-header { text-align: center; border-bottom: 2px solid var(--primary); padding-bottom: 24px; margin-bottom: 32px; }
#module-invoice .invoice-header h1 { color: var(--primary); font-size: 32px; letter-spacing: 4px; }
#module-invoice .invoice-meta { display: flex; justify-content: space-between; font-size: 13px; color: var(--text-muted); margin-bottom: 24px; }
#module-invoice .invoice-info { background: var(--bg); padding: 20px; border-radius: var(--radius); font-size: 14px; margin-bottom: 32px; line-height: 1.8; }
#module-invoice .invoice-table { width: 100%; border-collapse: collapse; margin-bottom: 32px; text-align: left; }
#module-invoice .invoice-table th, #module-invoice .invoice-table td { padding: 14px 8px; border-bottom: 1px solid var(--border); font-size: 14px; }
#module-invoice .invoice-table th { color: var(--text-muted); font-weight: 500; }
#module-invoice .invoice-total { text-align: right; font-size: 14px; line-height: 2; margin-bottom: 24px; color: var(--text-muted); }
#module-invoice .total-amount { font-size: 28px; color: var(--text-main); margin-top: 8px; }
#invoice-stamp-area {
  position: absolute; bottom: 60px; right: 60px; width: 140px; height: 140px;
  display: flex; align-items: center; justify-content: center; opacity: 0.85; pointer-events: none;
}
#invoice-stamp-area img { max-width: 100%; max-height: 100%; object-fit: contain; mix-blend-mode: multiply; }
#module-invoice .ledger-panel { margin-top: 40px; background: var(--card); padding: 40px; border-radius: var(--radius); border: 1px solid var(--border); }
#module-invoice .ledger-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 32px; }
#module-invoice .filter-bar { display: flex; gap: 16px; margin-bottom: 24px; background: var(--bg); padding: 20px; border-radius: var(--radius); border: 1px solid var(--border); flex-wrap: wrap; }
#module-invoice .filter-bar input, #module-invoice .filter-bar select { padding: 10px; font-size: 13px; flex: 1; min-width: 140px; }
#module-invoice .ledger-table { width: 100%; border-collapse: collapse; }
#module-invoice .ledger-table th, #module-invoice .ledger-table td { padding: 16px 12px; border-bottom: 1px solid var(--border); text-align: left; font-size: 14px; }
#module-invoice .ledger-table th { color: var(--text-muted); font-weight: 500; background: var(--bg); }

/* 出纳 */
#module-cashier .quick-cards { display: grid; grid-template-columns: repeat(3, 1fr); gap: 20px; margin-bottom: 28px; }
#module-cashier .quick-card {
  background: var(--card); border: 1.5px solid var(--border); border-radius: var(--radius);
  padding: 20px 24px; cursor: pointer; transition: all .2s ease;
  display: flex; align-items: center; gap: 16px;
}
#module-cashier .quick-card:hover { border-color: var(--primary); transform: translateY(-2px); box-shadow: 0 8px 24px rgba(0,0,0,.06); }
#module-cashier .quick-card.active { border-color: var(--primary); background: #FFF9F6; }
#module-cashier .quick-icon {
  width: 44px; height: 44px; border-radius: 50%; display: flex; align-items: center; justify-content: center;
  font-size: 20px; font-weight: 600; flex-shrink: 0;
}
#module-cashier .quick-title { font-weight: 600; font-size: 15px; color: var(--text-main); }
#module-cashier .quick-desc { font-size: 12px; color: var(--text-muted); margin-top: 2px; }
#module-cashier .entry-panel { background: var(--card); border: 1px solid var(--border); border-radius: var(--radius); padding: 32px; margin-bottom: 28px; }
#module-cashier .entry-head { display: flex; justify-content: space-between; align-items: center; margin-bottom: 24px; flex-wrap: wrap; gap: 12px; }
#module-cashier .type-switch { display: flex; gap: 4px; background: var(--bg); padding: 4px; border-radius: 10px; }
#module-cashier .type-btn {
  padding: 8px 20px; border: none; background: transparent; border-radius: 8px;
  font-size: 13px; font-weight: 500; cursor: pointer; color: var(--text-muted);
  font-family: var(--font-sans); transition: all .2s;
}
#module-cashier .type-btn.active { background: var(--card); color: var(--text-main); box-shadow: 0 1px 3px rgba(0,0,0,.06); }
#module-cashier .cash-form-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 16px; }
#module-cashier .cf-group label { display: block; font-size: 12px; color: var(--text-muted); margin-bottom: 6px; }
#module-cashier .cf-span-2 { grid-column: span 2; }
#module-cashier .cash-receipt-upload { margin-top: 20px; padding: 16px; background: var(--bg); border-radius: var(--radius); border: 1px solid var(--border); }
#module-cashier .cash-receipt-upload .upload-zone-mini { margin-bottom: 0; background: var(--card); }
#module-cashier .cash-receipt-upload .attachment-list { margin-top: 10px; margin-bottom: 0; }
#module-cashier .cash-output { background: var(--card); border: 1px solid var(--border); border-radius: var(--radius); overflow: hidden; }
#module-cashier .cash-tabs { display: flex; gap: 4px; padding: 0 20px; background: var(--bg); border-bottom: 1px solid var(--border); overflow-x: auto; }
#module-cashier .cash-tab { padding: 14px 16px; font-size: 13px; font-weight: 500; color: var(--text-muted); cursor: pointer; border-bottom: 2px solid transparent; white-space: nowrap; }
#module-cashier .cash-tab.active { color: var(--text-main); border-bottom-color: var(--primary); }
#module-cashier .cash-tab-panel { display: none; padding: 24px; }
#module-cashier .cash-tab-panel.active { display: block; }
#module-cashier .cash-toolbar { display: flex; gap: 12px; margin-bottom: 20px; flex-wrap: wrap; align-items: center; }
#module-cashier .cash-toolbar input, #module-cashier .cash-toolbar select { width: auto; flex: 1; min-width: 120px; padding: 8px 12px; font-size: 13px; }
#module-cashier .cash-table { width: 100%; border-collapse: collapse; }
#module-cashier .cash-table th, #module-cashier .cash-table td { padding: 12px 14px; border-bottom: 1px solid var(--border); font-size: 13px; text-align: left; }
#module-cashier .cash-table th { background: var(--bg); color: var(--text-muted); font-weight: 500; font-size: 12px; }
#module-cashier .account-manager { background: var(--bg); border: 1px solid var(--border); border-radius: var(--radius); padding: 20px; margin-bottom: 24px; }
#module-cashier .acc-manage-table { width: 100%; border-collapse: collapse; }
#module-cashier .acc-manage-table th, #module-cashier .acc-manage-table td { padding: 8px 10px; border-bottom: 1px solid var(--border); font-size: 13px; text-align: left; }
#module-cashier .acc-manage-table th { color: var(--text-muted); font-weight: 500; font-size: 12px; }
#module-cashier .acc-manage-table input, #module-cashier .acc-manage-table select { padding: 6px 8px; font-size: 12px; }

/* 制证 */
#module-voucher .voucher-types { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 14px; margin-bottom: 28px; }
#module-voucher .voucher-type-card { background: var(--card); border: 1.5px solid var(--border); border-radius: var(--radius); padding: 16px 18px; cursor: pointer; transition: all .2s ease; display: flex; align-items: center; gap: 12px; }
#module-voucher .voucher-type-card:hover { border-color: var(--primary); transform: translateY(-2px); box-shadow: 0 8px 24px rgba(0,0,0,.06); }
#module-voucher .voucher-type-card.active { border-color: var(--primary); background: #FFF9F6; }
#module-voucher .voucher-type-icon { width: 38px; height: 38px; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-size: 16px; font-weight: 600; flex-shrink: 0; }
#module-voucher .voucher-type-title { font-weight: 600; font-size: 14px; color: var(--text-main); }
#module-voucher .voucher-layout { display: grid; grid-template-columns: 360px 1fr; gap: 28px; margin-bottom: 28px; }
#module-voucher .voucher-source-panel, #module-voucher .voucher-entry-panel { background: var(--card); border: 1px solid var(--border); border-radius: var(--radius); padding: 28px; }
#module-voucher .voucher-panel-title { font-size: 18px; margin-bottom: 20px; display: flex; align-items: center; gap: 10px; }
#module-voucher .voucher-panel-title .badge { font-size: 11px; font-weight: 500; background: var(--bg); color: var(--text-muted); padding: 2px 8px; border-radius: 10px; font-family: var(--font-sans); }
#module-voucher .vform-group { margin-bottom: 14px; }
#module-voucher .vform-group label { display: block; font-size: 12px; color: var(--text-muted); margin-bottom: 6px; }
#module-voucher .vform-group input, #module-voucher .vform-group select { padding: 9px 12px; font-size: 13px; }
#module-voucher .voucher-meta-row { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; margin-bottom: 16px; }
#module-voucher .voucher-meta-row label { font-size: 11px; color: var(--text-muted); display: block; margin-bottom: 4px; }
#module-voucher .voucher-meta-row input { padding: 8px 10px; font-size: 12px; }
#module-voucher .voucher-entry-table { width: 100%; border-collapse: collapse; margin-bottom: 12px; }
#module-voucher .voucher-entry-table th, #module-voucher .voucher-entry-table td { padding: 6px 4px; border-bottom: 1px solid var(--border); font-size: 12px; text-align: left; vertical-align: middle; }
#module-voucher .voucher-entry-table th { background: var(--bg); color: var(--text-muted); font-weight: 500; font-size: 11px; }
#module-voucher .voucher-entry-table input, #module-voucher .voucher-entry-table select { padding: 6px 8px; font-size: 12px; background: transparent; border: 1px solid transparent; border-radius: 4px; }
#module-voucher .voucher-entry-table input:focus, #module-voucher .voucher-entry-table select:focus { border-color: var(--primary); background: var(--card); }
#module-voucher .voucher-entry-table .num { text-align: right; font-variant-numeric: tabular-nums; }
#module-voucher .voucher-entry-table .btn-remove-row { background: transparent; border: none; color: var(--diff); cursor: pointer; font-size: 15px; }
#module-voucher .btn-add-entry { background: var(--bg); color: var(--text-main); padding: 8px; width: 100%; border-radius: 8px; border: 1px dashed var(--border); cursor: pointer; font-size: 12px; font-weight: 500; }
#module-voucher .btn-add-entry:hover { border-color: var(--primary); color: var(--primary); }
#module-voucher .voucher-totals { display: flex; justify-content: space-between; align-items: center; padding: 14px 18px; margin-top: 16px; background: var(--bg); border-radius: var(--radius); flex-wrap: wrap; gap: 12px; }
#module-voucher .voucher-total-item { display: flex; align-items: center; gap: 8px; font-size: 13px; }
#module-voucher .voucher-total-item strong { font-family: var(--font-serif); font-size: 20px; color: var(--text-main); }
#module-voucher .voucher-balance-status { font-size: 12px; padding: 4px 12px; border-radius: 12px; font-weight: 500; }
#module-voucher .voucher-balance-status.ok { background: var(--match-bg); color: var(--match); }
#module-voucher .voucher-balance-status.bad { background: var(--diff-bg); color: var(--diff); }
#module-voucher .voucher-submit-bar { background: var(--card); border: 1px solid var(--border); border-radius: var(--radius); padding: 18px 28px; display: flex; justify-content: space-between; align-items: center; margin-bottom: 28px; flex-wrap: wrap; gap: 12px; }
#module-voucher .voucher-submit-bar .hint { font-size: 12px; color: var(--text-muted); }
#module-voucher .voucher-list-panel { background: var(--card); border: 1px solid var(--border); border-radius: var(--radius); padding: 28px; }
#module-voucher .voucher-period-bar { display: flex; gap: 12px; align-items: center; flex-wrap: wrap; padding: 14px 18px; background: var(--bg); border-radius: var(--radius); margin-bottom: 20px; font-size: 12px; }
#module-voucher .voucher-period-bar input { padding: 6px 10px; font-size: 12px; width: auto; }
#module-voucher .voucher-table { width: 100%; border-collapse: collapse; }
#module-voucher .voucher-table th, #module-voucher .voucher-table td { padding: 12px 10px; border-bottom: 1px solid var(--border); font-size: 13px; text-align: left; vertical-align: top; }
#module-voucher .voucher-table th { background: var(--bg); color: var(--text-muted); font-weight: 500; font-size: 12px; }

/* 会计记账 */
#module-accounting .acc-period-bar { display: flex; gap: 12px; align-items: center; flex-wrap: wrap; background: var(--card); border: 1px solid var(--border); border-radius: var(--radius); padding: 16px 20px; margin-bottom: 20px; }
#module-accounting .acc-tabs { display: flex; gap: 4px; background: var(--bg); border-radius: var(--radius); padding: 4px; margin-bottom: 20px; overflow-x: auto; }
#module-accounting .acc-tab { padding: 10px 18px; font-size: 13px; font-weight: 500; color: var(--text-muted); cursor: pointer; border-radius: 8px; white-space: nowrap; transition: all .2s; }
#module-accounting .acc-tab.active { background: var(--card); color: var(--text-main); box-shadow: 0 1px 3px rgba(0,0,0,.06); }
#module-accounting .acc-badge { display: inline-block; padding: 1px 8px; border-radius: 10px; background: var(--primary); color: #fff; font-size: 11px; margin-left: 6px; }
#module-accounting .acc-tab-panel { display: none; background: var(--card); border: 1px solid var(--border); border-radius: var(--radius); padding: 24px; }
#module-accounting .acc-tab-panel.active { display: block; }
#module-accounting .acc-detail-bar { display: flex; gap: 12px; align-items: center; margin-bottom: 16px; flex-wrap: wrap; }
#module-accounting .acc-report-subtabs, #module-drilldown .acc-report-subtabs { display: flex; gap: 8px; border-bottom: 1px solid var(--border); margin-bottom: 12px; }
#module-accounting .acc-report-subtab, #module-drilldown .acc-report-subtab { padding: 10px 16px; font-size: 13px; font-weight: 500; color: var(--text-muted); cursor: pointer; border-bottom: 2px solid transparent; background: transparent; border-top: none; border-left: none; border-right: none; }
#module-accounting .acc-report-subtab.active, #module-drilldown .acc-report-subtab.active { color: var(--text-main); border-bottom-color: var(--primary); }
.acc-table { width: 100%; border-collapse: collapse; }
.acc-table th, .acc-table td { padding: 10px 12px; border-bottom: 1px solid var(--border); font-size: 13px; text-align: left; }
.acc-table th { background: var(--bg); color: var(--text-muted); font-weight: 500; font-size: 12px; }
.acc-table tfoot td { font-weight: 600; background: var(--bg); }

/* 应收、计税、结转 */
.module-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 24px; margin-bottom: 24px; }
.card-block { background: var(--card); border: 1px solid var(--border); border-radius: var(--radius); padding: 24px; margin-bottom: 24px; }
.card-block h3 { font-size: 18px; margin-bottom: 16px; border-bottom: 1px solid var(--border); padding-bottom: 12px; }
.list-item { display: flex; justify-content: space-between; align-items: center; padding: 10px 0; border-bottom: 1px dashed var(--border); font-size: 13px; }
.list-item:last-child { border-bottom: none; }
.flex-between { display: flex; justify-content: space-between; align-items: center; gap: 16px; flex-wrap: wrap; }
.note-text { font-size: 12px; color: var(--text-muted); margin-top: 8px; line-height: 1.5; }
.tax-badge { display: inline-block; padding: 2px 8px; border-radius: 4px; font-size: 11px; font-weight: 500; }
.tax-badge.vat { background: #EFF6FF; color: #1D4ED8; }
.tax-badge.cit { background: #F0FDF4; color: #15803D; }
.tax-badge.surtax { background: #FEF2F2; color: #B91C1C; }

@media print {
  body { background: #FFF; }
  .app-header, .app-sidebar, .workspace-tabs-bar, .no-print { display: none !important; }
  .app-shell { display: block; }
  .module { padding: 0; display: block !important; opacity: 1 !important; }
  .panel, .card-block { border: none !important; box-shadow: none !important; padding: 0 !important; }
}

@media (max-width: 1024px) {
  .sidebar-toggle-btn { display: inline-flex; }
  .app-sidebar {
    position: fixed; left: 0; top: var(--header-height);
    height: calc(100vh - var(--header-height));
    transform: translateX(-100%);
    box-shadow: 8px 0 30px rgba(0,0,0,0.08);
  }
  .app-sidebar.open { transform: translateX(0); }
  .app-header { padding: 12px 16px; height: auto; }
  .module { padding: 20px 16px; }
  .workspace-tabs-bar { padding: 8px 16px 0; top: 0; position: relative; }
  .home-title { font-size: 46px; }
  #module-invoice .layout { flex-direction: column; }
  #module-invoice .form-panel, #module-invoice .invoice-panel { width: 100%; min-width: 100%; }
  #module-recon .upload-container, #module-recon .mapping-row, #module-invoice .filter-bar { flex-direction: column; }
  #module-cashier .quick-cards { grid-template-columns: 1fr; }
  #module-cashier .cf-span-2 { grid-column: span 1; }
  #module-voucher .voucher-layout { grid-template-columns: 1fr; }
  .module-grid { grid-template-columns: 1fr; }
}
@media (max-width: 700px) {
  .workbench-actions { justify-content: flex-end; }
  .draft-status { display: none; }
}
</style>
</head>
<body>
<div id="toast">操作提示</div>

<!-- 顶部栏 -->
<header class="app-header no-print">
  <div class="header-left">
    <button type="button" class="sidebar-toggle-btn" id="sidebar-toggle" aria-label="展开或收起菜单">☰</button>
    <div class="app-logo" id="logo-home-btn">
      <span>✻</span> finance workbench
    </div>
  </div>
  <div class="workbench-actions">
    <span id="workbench-draft-status" class="draft-status" role="status" aria-live="polite"></span>
    <button type="button" id="workbench-save" class="btn btn-primary btn-icon-label" disabled title="保存智能开票、出纳收付和日常制证的草稿">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M5 3h12l4 4v12a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2Z"/><path d="M7 3v6h9V3M7 21v-8h10v8"/></svg>
      一键保存
    </button>
    <button type="button" id="workbench-export" class="btn btn-icon-label" disabled title="导出本机账套备份">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 3v12m-5-5 5 5 5-5M4 16v4a1 1 0 0 0 1 1h14a1 1 0 0 0 1-1v-4"/></svg>
      导出备份
    </button>
    <button type="button" id="workbench-safe-exit" class="btn btn-icon-label" title="安全退出并提示备份本地账套">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2a 2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"/><polyline points="16 17 21 12 16 7"/><line x1="21" y1="12" x2="9" y2="12"/></svg>
      安全退出 / 备份
    </button>
  </div>
</header>

<div class="sidebar-backdrop" id="sidebar-backdrop"></div>

<!-- 主体外壳：左侧树形导航 + 右侧宽屏工作区 -->
<div class="app-shell">
  <aside class="app-sidebar no-print" id="app-sidebar" aria-label="功能导航">
    <div class="sidebar-header">
      <span class="sidebar-header-title">业务模块导航</span>
    </div>
    <button type="button" class="sidebar-home-btn active" id="nav-home-btn" data-tab-key="home">
      <span>🏠</span> 工作台首页
    </button>
    <nav class="sidebar-tree" id="sidebar-tree"></nav>
  </aside>

  <main class="app-main">
    <!-- 多标签页栏 -->
    <div class="workspace-tabs-bar no-print" id="workspace-tabs-bar">
      <div class="workspace-tabs-scroll" id="workspace-tabs-list" role="tablist"></div>
      <div class="workspace-tabs-actions">
        <button type="button" class="btn-mini" id="tabs-close-others" title="关闭除当前外的其他标签页">关闭其他</button>
      </div>
    </div>

    <!-- 工作区内容 -->
    <div class="workspace-content">

<!-- ================= 首页 ================= -->
<div id="module-home" class="module active">
  <div class="home-container">
    <div class="home-header-text">
      <div class="home-subtitle-top">FINANCE WORKBENCH</div>
      <h1 class="home-title serif-title">finance<br><i>workbench</i></h1>
      <p class="home-desc">安全、精准的本地化财务工具。覆盖数据对账、智能开票、出纳收付、凭证制审、七列余额表与右键穿透联查、往来三栏账、税务及期末结转全流程。</p>
    </div>

    <div class="home-cards">
      <div class="home-card" data-open-tab="gl-balance">
        <div class="home-card-content">
          <h3 class="serif-title">科目余额与穿透</h3>
          <p>标准七列余额表，右键穿透银行专项与单位往来，支持 Excel 导入[cite: 1]。</p>
          <button type="button" class="btn-enter">进入余额表 →</button>
        </div>
        <svg class="card-svg-bg" viewBox="0 0 100 100" fill="none">
          <circle cx="45" cy="45" r="22" fill="#FFF" stroke="#141413" stroke-width="2"/>
          <line x1="61" y1="61" x2="80" y2="80" stroke="#D97757" stroke-width="3" stroke-linecap="round"/>
          <path d="M35 45 H55 M45 35 V55" stroke="#141413" stroke-width="1.8" stroke-linecap="round"/>
        </svg>
      </div>

      <div class="home-card" data-open-tab="ar-subledger">
        <div class="home-card-content">
          <h3 class="serif-title">单位往来三栏账</h3>
          <p>按客户/供应商切换三栏明细账，期初借贷发生与逐笔余额实时同步。</p>
          <button type="button" class="btn-enter">进入三栏账 →</button>
        </div>
        <svg class="card-svg-bg" viewBox="0 0 100 100" fill="none">
          <rect x="22" y="20" width="56" height="60" rx="4" fill="#FFF" stroke="#141413" stroke-width="2"/>
          <line x1="22" y1="36" x2="78" y2="36" stroke="#141413" stroke-width="1.5"/>
          <line x1="48" y1="36" x2="48" y2="80" stroke="#141413" stroke-width="1.5"/>
          <circle cx="66" cy="64" r="10" fill="#D97757" opacity="0.85"/>
        </svg>
      </div>

      <div class="home-card" data-open-tab="recon">
        <div class="home-card-content">
          <h3 class="serif-title">数据对账</h3>
          <p>Excel / CSV 流水导入，智能比对差异、漏记与余额调节表[cite: 1]。</p>
          <button type="button" class="btn-enter">进入对账 →</button>
        </div>
        <svg class="card-svg-bg" viewBox="0 0 100 100" fill="none">
          <path d="M20 70 C 40 70, 60 50, 80 30" stroke="#141413" stroke-width="2" stroke-linecap="round"/>
          <circle cx="20" cy="70" r="6" fill="#FFF" stroke="#141413" stroke-width="2"/>
          <circle cx="80" cy="30" r="6" fill="#FFF" stroke="#141413" stroke-width="2"/>
          <circle cx="50" cy="50" r="10" fill="#D97757" opacity="0.8"/>
        </svg>
      </div>

      <div class="home-card" data-open-tab="invoice">
        <div class="home-card-content">
          <h3 class="serif-title">智能开票</h3>
          <p>开具电子发票，上传印章与合同凭证，自动沉淀开票总账[cite: 1]。</p>
          <button type="button" class="btn-enter">进入开票 →</button>
        </div>
        <svg class="card-svg-bg" viewBox="0 0 100 100" fill="none">
          <rect x="30" y="20" width="40" height="60" rx="3" fill="#FFF" stroke="#141413" stroke-width="2"/>
          <line x1="40" y1="35" x2="60" y2="35" stroke="#141413" stroke-width="2" stroke-linecap="round"/>
          <line x1="40" y1="45" x2="60" y2="45" stroke="#141413" stroke-width="2" stroke-linecap="round"/>
          <circle cx="70" cy="70" r="15" fill="#D97757" opacity="0.8"/>
        </svg>
      </div>

      <div class="home-card" data-open-tab="fund-bank">
        <div class="home-card-content">
          <h3 class="serif-title">出纳收付</h3>
          <p>收付款登记，自定义账户，生成出纳日报、票据与日记账[cite: 1]。</p>
          <button type="button" class="btn-enter">进入出纳 →</button>
        </div>
        <svg class="card-svg-bg" viewBox="0 0 100 100" fill="none">
          <path d="M30 50 L70 50 M50 30 L50 70" stroke="#141413" stroke-width="2" stroke-linecap="round"/>
          <circle cx="50" cy="50" r="20" fill="#FFF" stroke="#141413" stroke-width="2"/>
        </svg>
      </div>

      <div class="home-card" data-open-tab="gl-voucher">
        <div class="home-card-content">
          <h3 class="serif-title">日常制证</h3>
          <p>业务单据制证，支持辅助核算、借贷平衡与期间锁定校验[cite: 1]。</p>
          <button type="button" class="btn-enter">进入制证 →</button>
        </div>
        <svg class="card-svg-bg" viewBox="0 0 100 100" fill="none">
          <rect x="25" y="15" width="50" height="70" rx="4" fill="#FFF" stroke="#141413" stroke-width="2"/>
          <path d="M35 30 L65 30 M35 45 L65 45 M35 60 L50 60" stroke="#141413" stroke-width="2" stroke-linecap="round"/>
          <circle cx="75" cy="75" r="12" fill="#D97757" opacity="0.8"/>
        </svg>
      </div>

      <div class="home-card" data-open-tab="gl-query">
        <div class="home-card-content">
          <h3 class="serif-title">会计记账</h3>
          <p>凭证审核记账，生成汇总、明细、总账、试算及财务报表[cite: 1]。</p>
          <button type="button" class="btn-enter">进入记账 →</button>
        </div>
        <svg class="card-svg-bg" viewBox="0 0 100 100" fill="none">
          <rect x="20" y="25" width="60" height="50" rx="2" fill="#FFF" stroke="#141413" stroke-width="2"/>
          <line x1="35" y1="25" x2="35" y2="75" stroke="#141413" stroke-width="2"/>
          <line x1="50" y1="25" x2="50" y2="75" stroke="#141413" stroke-width="2"/>
          <line x1="65" y1="25" x2="65" y2="75" stroke="#141413" stroke-width="2"/>
        </svg>
      </div>

      <div class="home-card" data-open-tab="ar-ledger">
        <div class="home-card-content">
          <h3 class="serif-title">应收函证与核销</h3>
          <p>客户月结核对、未分配回款冲销、年末询证函与回函差异[cite: 1]。</p>
          <button type="button" class="btn-enter">进入函证 →</button>
        </div>
        <svg class="card-svg-bg" viewBox="0 0 100 100" fill="none">
          <path d="M20 40 L50 20 L80 40 L80 80 L20 80 Z" fill="#FFF" stroke="#141413" stroke-width="2"/>
          <path d="M20 40 L50 60 L80 40" stroke="#141413" stroke-width="2" stroke-linecap="round"/>
        </svg>
      </div>

      <div class="home-card" data-open-tab="tax-calc">
        <div class="home-card-content">
          <h3 class="serif-title">税务管理</h3>
          <p>月末计税测算底稿、申报任务跟踪与完税证明归档[cite: 1]。</p>
          <button type="button" class="btn-enter">进入税务 →</button>
        </div>
        <svg class="card-svg-bg" viewBox="0 0 100 100" fill="none">
          <rect x="25" y="30" width="50" height="40" rx="4" fill="#FFF" stroke="#141413" stroke-width="2"/>
          <rect x="35" y="40" width="10" height="10" fill="#D97757" opacity="0.8"/>
          <rect x="55" y="40" width="10" height="10" fill="#141413" opacity="0.2"/>
        </svg>
      </div>

      <div class="home-card" data-open-tab="fin-cost">
        <div class="home-card-content">
          <h3 class="serif-title">期末结转</h3>
          <p>成本折旧结转、结账前检查与财务档案快照归档[cite: 1]。</p>
          <button type="button" class="btn-enter">进入结转 →</button>
        </div>
        <svg class="card-svg-bg" viewBox="0 0 100 100" fill="none">
          <path d="M30 70 C 20 70, 20 50, 40 50 C 40 30, 70 30, 70 50 C 90 50, 90 70, 80 70" stroke="#141413" stroke-width="2" stroke-linecap="round" fill="none"/>
          <path d="M50 50 L50 80 M40 70 L50 80 L60 70" stroke="#141413" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
        </svg>
      </div>
    </div>
  </div>
</div>

<!-- ================= 占位提示页 ================= -->
<div id="module-placeholder" class="module">
  <div class="placeholder-card">
    <div class="placeholder-badge" id="placeholder-module-name">模块功能建设中</div>
    <h2 class="serif-title" id="placeholder-title">功能模块规划中</h2>
    <p id="placeholder-desc">该子菜单对应的业务表单与账表逻辑将在后续迭代中逐步上线，当前可先使用左侧已挂载的核心账务功能[cite: 1]。</p>
    <div class="placeholder-actions">
      <button type="button" class="btn btn-primary" id="placeholder-back-home">返回工作台首页</button>
    </div>
  </div>
</div>

<!-- ================= 单位往来三栏账（代码1新增融合） ================= -->
<div id="module-ar-subledger" class="module">
  <div class="module-header">
    <div>
      <div class="home-subtitle-top">AR / AP · 往来明细</div>
      <h1 class="serif-title">单位往来三栏账</h1>
      <p style="max-width:820px;">按往来科目与往来单位切换，展示标准三栏账格式（日期、凭证号、摘要、借方、贷方、方向与逐笔余额），与已入账凭证实时同步。</p>
    </div>
    <div class="section-actions">
      <button type="button" class="btn btn-primary" id="btn-ar-sub-export">导出 Excel（CSV）</button>
    </div>
  </div>

  <div class="panel">
    <div class="toolbar">
      <div class="form-group" style="min-width:190px; margin:0;">
        <label>往来科目</label>
        <select id="ar-sub-subject">
          <option value="1122">1122 应收账款</option>
          <option value="2202">2202 应付账款</option>
        </select>
      </div>
      <div class="form-group" style="min-width:260px; margin:0;">
        <label>往来单位</label>
        <select id="ar-sub-unit"></select>
      </div>
      <div class="form-group" style="min-width:160px; margin:0;">
        <label>会计期间</label>
        <input type="month" id="ar-sub-period" value="2026-01">
      </div>
    </div>

    <div class="sum-bar" id="ar-sub-summary"></div>

    <div style="overflow:auto;">
      <table class="fin-table">
        <thead>
          <tr>
            <th width="110">日期</th>
            <th width="130">凭证号</th>
            <th>摘要</th>
            <th class="num" width="135">借方金额</th>
            <th class="num" width="135">贷方金额</th>
            <th width="70">方向</th>
            <th class="num" width="145">余额</th>
          </tr>
        </thead>
        <tbody id="ar-sub-body"></tbody>
      </table>
    </div>
  </div>
</div>

<!-- ================= 数据对账 ================= -->
<div id="module-recon" class="module">
  <div class="upload-container">
    <div class="drop-zone" id="recon-drop-bank">
      <h3 class="serif-title">上传【银行流水】</h3>
      <p>支持 .xlsx / .xls / .csv 格式[cite: 1]</p>
      <input type="file" id="recon-file-bank" accept=".xlsx,.xls,.csv" style="display:none">
    </div>
    <div class="drop-zone" id="recon-drop-company">
      <h3 class="serif-title">上传【内部记账】</h3>
      <p>支持 .xlsx / .xls / .csv 格式[cite: 1]</p>
      <input type="file" id="recon-file-company" accept=".xlsx,.xls,.csv" style="display:none">
    </div>
  </div>
  <div class="mapping-area" id="recon-mapping-area">
    <div class="mapping-row">
      <div class="mapping-group"><label>银行单号列</label><select id="recon-map-bank-id"></select></div>
      <div class="mapping-group"><label>银行金额列</label><select id="recon-map-bank-amount"></select></div>
    </div>
    <div class="mapping-row">
      <div class="mapping-group"><label>企业单号列</label><select id="recon-map-company-id"></select></div>
      <div class="mapping-group"><label>企业金额列</label><select id="recon-map-company-amount"></select></div>
    </div>
  </div>
  <div class="section-actions recon-actions" id="recon-actions">
    <button type="button" class="btn btn-primary" id="recon-btn-reconcile" disabled>开始对账</button>
  </div>
  <div class="result-area" id="recon-result-area">
    <div class="tabs" id="recon-tabs">
      <div class="tab active" data-type="ALL">全部记录 <span class="badge" id="badge-ALL" style="background:var(--text-muted)">0</span></div>
      <div class="tab" data-type="MATCH">完全匹配 <span class="badge" id="badge-MATCH" style="background:var(--match)">0</span></div>
      <div class="tab" data-type="BANK_ONLY">企业漏记 <span class="badge" id="badge-BANK_ONLY" style="background:var(--bank-only)">0</span></div>
      <div class="tab" data-type="COMPANY_ONLY">银行缺失 <span class="badge" id="badge-COMPANY_ONLY" style="background:var(--company-only)">0</span></div>
      <div class="tab" data-type="DIFF">金额不符 <span class="badge" id="badge-DIFF" style="background:var(--diff)">0</span></div>
      <div class="tab" data-type="ERROR">异常 <span class="badge" id="badge-ERROR" style="background:var(--error)">0</span></div>
    </div>
    <div style="padding:16px 24px; text-align:right; border-bottom:1px solid var(--border);">
      <button type="button" class="btn" id="recon-btn-export">导出当前分类</button>
    </div>
    <div style="overflow:auto">
      <table id="recon-result-table">
        <thead><tr><th width="16%">单号</th><th width="13%">状态</th><th width="12%">银行金额</th><th width="12%">账面金额</th><th width="47%">缺失/差异详情</th></tr></thead>
        <tbody id="recon-result-body"></tbody>
      </table>
    </div>
  </div>
</div>

<!-- ================= 智能开票 ================= -->
<div id="module-invoice" class="module">
  <div class="layout">
    <div class="form-panel no-print">
      <h2 class="serif-title">发票信息录入</h2>
      <div style="background:var(--bg); padding:20px; border-radius:var(--radius); margin-bottom:24px; border:1px solid var(--border);">
        <div class="form-group" style="margin-bottom:12px;"><label>销售方名称 (我方)</label><input type="text" id="inv-input-seller" class="inv-sync"></div>
        <div class="form-group" style="margin-bottom:0;"><label>销售方税号</label><input type="text" id="inv-input-seller-taxno" class="inv-sync"></div>
      </div>
      <div class="form-group"><label>购买方名称 (客户) <span style="color:var(--diff)">*</span></label><input type="text" id="inv-input-buyer" class="inv-sync" autocomplete="off"></div>
      <div class="form-group"><label>合同 / 开票通知编号</label><input id="inv-source"></div>
      <div class="form-group"><label>办理状态</label><select id="inv-status"><option>待开票</option><option>已开票</option><option>办理失败</option></select></div>
      <div class="form-group"><label>回执编号 / 失败原因</label><input id="inv-evidence"></div>
      <div class="form-group"><label>购买方税号</label><input type="text" id="inv-input-taxno" class="inv-sync" autocomplete="off"></div>
      <div class="form-group" style="display:flex; gap:16px;">
        <div style="flex:1"><label>实际发票号码</label><input type="text" id="inv-input-no" class="inv-sync" style="background:transparent; border:1px dashed var(--border);"></div>
        <div style="flex:1"><label>开票日期</label><input type="date" id="inv-input-date" class="inv-sync"></div>
      </div>
      <div class="form-group">
        <label>收款方式</label>
        <select id="inv-input-payment" class="inv-sync">
          <option value="银行转账">银行转账</option>
          <option value="支付宝">支付宝</option>
          <option value="微信支付">微信支付</option>
          <option value="现金">现金</option>
          <option value="支票">支票</option>
        </select>
      </div>

      <h2 class="serif-title" style="margin-top:40px;">商品明细</h2>
      <table class="items-table" id="inv-items-table">
        <thead><tr><th width="40%">商品名称</th><th width="15%">数量</th><th width="20%">单价</th><th width="20%">小计</th><th width="5%"></th></tr></thead>
        <tbody id="inv-items-body"></tbody>
      </table>
      <button type="button" class="btn-add" id="inv-btn-add">+ 添加明细</button>

      <div class="additional-info">
        <h3 class="serif-title">附加信息与凭证</h3>
        <label style="font-size:13px; color:var(--text-muted); display:block; margin-bottom:8px;">上传发货单、合同等证明文件：</label>
        <div class="upload-zone-mini" id="attach-upload-zone">
          <p style="margin:0;">点击选择文件 (PDF / JPG / PNG)</p>
          <input type="file" id="attach-file-input" multiple accept=".pdf,.jpg,.jpeg,.png" style="display:none">
        </div>
        <div class="attachment-list" id="attachment-list"></div>
        <label style="font-size:13px; color:var(--text-muted); display:block; margin-bottom:8px; margin-top:16px;">上传发票专用章 (PNG透明底最佳)：</label>
        <div class="upload-zone-mini" id="stamp-upload-zone">
          <p style="margin:0; color:var(--primary);">点击上传印章图片</p>
          <input type="file" id="stamp-file-input" accept=".png,.jpg,.jpeg" style="display:none">
        </div>
        <div id="stamp-preview-mini" style="display:none; margin-top:10px;">
            <img id="stamp-preview-img" style="max-height:60px; border:1px solid var(--border); border-radius:4px; padding:4px; background:#fff;" alt="印章预览">
            <button type="button" class="btn" style="padding:4px 8px; font-size:12px; margin-left:8px;" onclick="removeStamp()">清除印章</button>
        </div>
        <label style="font-size:13px; color:var(--text-muted); display:block; margin-bottom:8px; margin-top:16px;">补充日志 / 备注说明：</label>
        <textarea id="inv-input-logs" placeholder="记录特殊说明..."></textarea>
      </div>
      <button type="button" class="btn btn-submit" id="inv-btn-submit">生成发票并存入总账</button>
    </div>

    <div class="invoice-panel" id="inv-print-area">
      <div class="invoice-header"><h1 class="serif-title">增值税电子普通发票</h1></div>
      <div class="invoice-meta">
        <span>发票代码: 044001900111</span>
        <span>发票号码: <span id="inv-prev-no">-</span></span>
      </div>
      <div class="invoice-meta">
        <span>开票日期: <span id="inv-prev-date">-</span></span>
        <span>收款方式: <span id="inv-prev-payment">-</span></span>
      </div>
      <div class="invoice-info">
        <div style="font-weight:500; margin-bottom:8px; color:var(--text-main);">购买方信息</div>
        <div>名称：<span id="inv-prev-buyer">-</span></div>
        <div>纳税人识别号：<span id="inv-prev-taxno">-</span></div>
        <div style="font-weight:500; margin-top:16px; margin-bottom:8px; color:var(--text-main);">销售方信息</div>
        <div>名称：<span id="inv-prev-seller">-</span></div>
        <div>纳税人识别号：<span id="inv-prev-seller-taxno">-</span></div>
      </div>
      <table class="invoice-table">
        <thead><tr><th>服务名称</th><th>数量</th><th>单价</th><th>金额</th></tr></thead>
        <tbody id="inv-prev-items-body"></tbody>
      </table>
      <div class="invoice-total">
        <div>合计金额：￥<span id="inv-prev-total">0.00</span></div>
        <div>价税合计(大写)：<span id="inv-prev-upper" style="color:var(--text-main); font-weight:500;">零元整</span></div>
        <div class="total-amount serif-title">(小写) ￥<span id="inv-prev-total-small">0.00</span></div>
      </div>
      <div id="invoice-stamp-area">
        <img id="invoice-stamp-img" src="" alt="印章" style="display:none;">
      </div>
    </div>
  </div>

  <div class="ledger-panel no-print">
    <div class="ledger-header">
      <div>
        <h2 class="serif-title" style="display:inline-block; margin:0;">历史开票总账</h2>
        <button type="button" class="btn-mini btn-danger" id="inv-btn-clear" style="margin-left:12px;">清空</button>
      </div>
      <div style="display:flex; align-items:center; gap:16px;">
        <button type="button" class="btn" id="inv-btn-export">导出 Excel</button>
      </div>
    </div>
    <div class="filter-bar">
      <input type="text" id="f-invoice-no" placeholder="🔍 发票号码">
      <input type="text" id="f-buyer" placeholder="🔍 购买方名称">
      <select id="f-payment">
        <option value="">全部收款方式</option>
        <option value="银行转账">银行转账</option>
        <option value="支付宝">支付宝</option>
        <option value="微信支付">微信支付</option>
        <option value="现金">现金</option>
        <option value="支票">支票</option>
      </select>
      <div style="display:flex; align-items:center; gap:8px; flex:1;">
        <input type="date" id="f-date-start" title="开始日期" style="width:auto;">
        <span style="color:var(--text-muted);">至</span>
        <input type="date" id="f-date-end" title="结束日期" style="width:auto;">
      </div>
      <button type="button" class="btn" id="inv-filter-reset" style="padding:8px 16px; font-size:13px;">重置</button>
    </div>
    <table class="ledger-table">
      <thead><tr><th>操作</th><th>发票号码</th><th>开票日期</th><th>购买方</th><th>金额</th><th>收款方式</th><th>凭证与日志</th></tr></thead>
      <tbody id="inv-ledger-body"></tbody>
      <tfoot>
        <tr style="background:var(--bg); font-weight:bold;">
          <td colspan="4" style="text-align:right; padding-right:24px;">筛选结果总计：</td>
          <td style="color:var(--primary); font-size:16px;" id="filtered-total-amount">¥0.00</td>
          <td colspan="2" style="color:var(--text-muted);" id="filtered-total-count">共 0 张</td>
        </tr>
      </tfoot>
    </table>
  </div>
</div>

<!-- ================= 出纳收付 ================= -->
<div id="module-cashier" class="module">
  <div class="module-header">
    <div>
      <div class="home-subtitle-top">CASHIER · 出纳</div>
      <h1 class="serif-title">收付款 · 转账 · 日报</h1>
      <p style="max-width:700px;">账户字段支持下拉选择，也支持直接手输任意账户[cite: 1]。</p>
    </div>
  </div>

  <div class="quick-cards">
    <div class="quick-card" data-type="RECEIPT" onclick="setCashierType('RECEIPT')">
      <div class="quick-icon" style="background:var(--match-bg); color:var(--match);">↓</div>
      <div><div class="quick-title">收款登记</div><div class="quick-desc">客户回款 / 现金收据</div></div>
    </div>
    <div class="quick-card" data-type="PAYMENT" onclick="setCashierType('PAYMENT')">
      <div class="quick-icon" style="background:var(--diff-bg); color:var(--diff);">↑</div>
      <div><div class="quick-title">付款登记</div><div class="quick-desc">供应商付款 / 报销支出</div></div>
    </div>
    <div class="quick-card" data-type="TRANSFER" onclick="setCashierType('TRANSFER')">
      <div class="quick-icon" style="background:var(--company-only-bg); color:var(--company-only);">⇄</div>
      <div><div class="quick-title">内部转账</div><div class="quick-desc">账户间划转 · 双向记账</div></div>
    </div>
  </div>

  <div class="entry-panel">
    <div class="entry-head">
      <h2 class="serif-title" style="font-size:22px;">登记出纳业务</h2>
      <div style="display:flex; gap:12px; align-items:center; flex-wrap:wrap;">
        <div class="type-switch">
          <button type="button" class="type-btn active" data-type="RECEIPT" onclick="setCashierType('RECEIPT')">收款</button>
          <button type="button" class="type-btn" data-type="PAYMENT" onclick="setCashierType('PAYMENT')">付款</button>
          <button type="button" class="type-btn" data-type="TRANSFER" onclick="setCashierType('TRANSFER')">内部转账</button>
        </div>
        <button type="button" class="btn-mini" onclick="toggleAccountManager()">⚙️ 管理常用账户</button>
      </div>
    </div>

    <div id="account-manager" class="account-manager" style="display:none;">
      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
        <strong style="font-size:14px;">常用账户管理（仅作为下拉建议，也可直接在表单中手输任意账户）</strong>
        <button type="button" class="btn-mini" onclick="toggleAccountManager()">收起</button>
      </div>
      <table class="acc-manage-table">
        <thead><tr><th>账户名称</th><th width="140px">类型</th><th width="90px">操作</th></tr></thead>
        <tbody id="account-manage-body"></tbody>
      </table>
      <div style="display:flex; gap:8px; margin-top:12px; flex-wrap:wrap;">
        <input type="text" id="new-account-name" placeholder="新账户名称，如：华夏银行 · 6388 **** 0001" style="flex:1; min-width:180px; padding:8px 12px; font-size:13px;">
        <select id="new-account-kind" style="width:120px; padding:8px 12px; font-size:13px;">
          <option value="BANK">银行</option>
          <option value="CASH">现金</option>
          <option value="OTHER">其他</option>
        </select>
        <button type="button" class="btn btn-primary" style="padding:8px 16px; font-size:13px;" onclick="addNewAccount()">添加账户</button>
      </div>
    </div>

    <div class="cash-form-grid">
      <div class="cf-group"><label>业务日期</label><input type="date" id="cash-date"></div>
      <div class="cf-group"><label>金额 (元)</label><input type="number" id="cash-amount" min="0.01" step="0.01" placeholder="0.00"></div>
      <div class="cf-group">
        <label id="cash-label-account">收款账户</label>
        <input type="text" id="cash-account" list="cash-account-options" placeholder="选择或直接输入账户名称" autocomplete="off">
        <datalist id="cash-account-options"></datalist>
      </div>
      <div class="cf-group" id="cash-field-in-account" style="display:none;">
        <label>转入账户</label>
        <input type="text" id="cash-in-account" list="cash-account-options" placeholder="选择或直接输入账户名称" autocomplete="off">
      </div>
      <div class="cf-group" id="cash-field-counterparty"><label id="cash-label-counterparty">付款方名称</label><input type="text" id="cash-counterparty" placeholder="对方单位 / 个人"></div>
      <div class="cf-group" id="cash-field-counterparty-account"><label>对方账号</label><input type="text" id="cash-counterparty-account" placeholder="选填"></div>
      <div class="cf-group" id="cash-field-source">
        <label>关联来源单据号</label>
        <div style="display:flex; gap:8px;">
          <input type="text" id="cash-source-doc" placeholder="审批单 / 合同号" style="flex:1;">
          <select id="cash-source-type" style="width:110px; flex:none;">
            <option value="审批单">审批单</option>
            <option value="合同">合同</option>
            <option value="发票">发票</option>
            <option value="其他">其他</option>
          </select>
        </div>
      </div>
      <div class="cf-group">
        <label>凭证类型</label>
        <select id="cash-receipt-type">
          <option value="BANK_RECEIPT">银行回单</option>
          <option value="CASH_RECEIPT">现金收据</option>
          <option value="NOTE">票据</option>
        </select>
      </div>
      <div class="cf-group"><label>回单 / 收据编号</label><input type="text" id="cash-receipt-no" placeholder="用于重复校验"></div>
      <div class="cf-group cf-span-2"><label>票据信息</label><input type="text" id="cash-note-info" placeholder="票号 / 出票日 / 到期日等"></div>
      <div class="cf-group cf-span-2"><label>备注</label><input type="text" id="cash-memo" placeholder="选填"></div>
    </div>

    <div class="cash-receipt-upload">
      <label style="font-size:12px; color:var(--text-muted); display:block; margin-bottom:8px;">回单 / 收据图片（选填，支持 JPG / PNG）</label>
      <div class="upload-zone-mini" id="cash-receipt-upload-zone">
        <p style="margin:0;">📷 点击上传回单 / 收据照片</p>
        <input type="file" id="cash-receipt-file-input" accept="image/*" style="display:none">
      </div>
      <div class="attachment-list" id="cash-receipt-preview"></div>
    </div>

    <div style="text-align:right; margin-top:24px;">
      <button type="button" class="btn btn-primary" id="cash-btn-submit">登记收款</button>
    </div>
  </div>

  <div class="cash-output">
    <div class="cash-tabs">
      <div class="cash-tab active" data-tab="register">收付款与转账登记表</div>
      <div class="cash-tab" data-tab="daily">出纳日报</div>
      <div class="cash-tab" data-tab="notes">票据台账</div>
      <div class="cash-tab" data-tab="journal">现金及银行日记账</div>
    </div>

    <div class="cash-tab-panel active" id="cash-tab-register">
      <div class="cash-toolbar">
        <select id="cash-filter-type">
          <option value="">全部类型</option>
          <option value="RECEIPT">收款</option>
          <option value="PAYMENT">付款</option>
          <option value="TRANSFER">转账</option>
        </select>
        <select id="cash-filter-status">
          <option value="">全部状态</option>
          <option value="PENDING">待办理</option>
          <option value="DONE">已办理</option>
          <option value="POSTED">已入账</option>
        </select>
        <input type="date" id="cash-filter-start" title="开始日期">
        <input type="date" id="cash-filter-end" title="结束日期">
        <button type="button" class="btn" style="padding:8px 16px; font-size:13px; flex:none;" onclick="resetCashFilters()">重置</button>
      </div>
      <div style="overflow:auto;">
        <table class="cash-table">
          <thead>
            <tr><th>日期</th><th>类型</th><th>金额</th><th>账户</th><th>对方</th><th>来源单据</th><th>凭证 / 回单</th><th>状态</th><th>操作</th></tr>
          </thead>
          <tbody id="cash-register-body"></tbody>
        </table>
      </div>
    </div>

    <div class="cash-tab-panel" id="cash-tab-daily">
      <p style="font-size:12px; color:var(--text-muted); margin-bottom:16px;">按业务日期汇总，点击日期行可展开查看每笔交易明细。</p>
      <div id="cash-daily-list"></div>
    </div>

    <div class="cash-tab-panel" id="cash-tab-notes">
      <div style="overflow:auto;">
        <table class="cash-table">
          <thead>
            <tr><th>回单 / 收据编号</th><th>凭证类型</th><th>日期</th><th>业务类型</th><th style="text-align:right;">金额</th><th>票据信息</th><th>回单图片</th><th>状态</th></tr>
          </thead>
          <tbody id="cash-notes-body"></tbody>
        </table>
      </div>
    </div>

    <div class="cash-tab-panel" id="cash-tab-journal">
      <div class="cash-toolbar">
        <select id="cash-journal-account"><option value="">全部账户</option></select>
        <span style="font-size:12px; color:var(--text-muted);">（仅统计「已办理 / 已入账」的实际资金流动）</span>
      </div>
      <div id="cash-journal-body"></div>
    </div>
  </div>
</div>

<!-- ================= 日常制证 ================= -->
<div id="module-voucher" class="module">
  <div class="module-header">
    <div>
      <div class="home-subtitle-top">VOUCHER · 制证</div>
      <h1 class="serif-title">日常制证 · 凭证审核</h1>
      <p style="max-width:820px;">系统校验借贷平衡、辅助核算（银行账户 / 往来单位）与会计期间；提交审核并记账后，科目余额表与单位往来三栏账实时同步刷新[cite: 1]。</p>
    </div>
  </div>

  <div class="voucher-types" id="voucher-types"></div>

  <div class="voucher-layout">
    <div class="voucher-source-panel">
      <div class="voucher-panel-title serif-title">
        原始单据 <span class="badge" id="voucher-source-badge">关联信息</span>
      </div>
      <div class="voucher-meta-row">
        <div><label>业务日期</label><input type="date" id="voucher-date"></div>
        <div><label>会计期间</label><input type="text" id="voucher-period" readonly style="background:var(--bg);"></div>
      </div>
      <div class="vform-group">
        <label>来源类型</label>
        <select id="voucher-source-type"></select>
      </div>
      <div class="vform-group">
        <label id="voucher-source-no-label">来源单号</label>
        <input type="text" id="voucher-source-no" placeholder="如未取得可留空">
      </div>
      <div class="vform-group">
        <label>对方单位 / 经办人</label>
        <input type="text" id="voucher-counterparty" placeholder="选填">
      </div>
      <div class="vform-group">
        <label>原始单据附件（选填）</label>
        <div class="upload-zone-mini" id="voucher-attach-zone" style="margin-bottom:8px;">
          <p style="margin:0;">📎 点击上传原始单据</p>
          <input type="file" id="voucher-attach-input" accept="image/*,.pdf" style="display:none">
        </div>
        <div class="attachment-list" id="voucher-attach-list" style="margin-bottom:0;"></div>
      </div>
      <div class="vform-group" style="margin-top:16px;">
        <label>备注</label>
        <input type="text" id="voucher-memo" placeholder="选填">
      </div>
    </div>

    <div class="voucher-entry-panel">
      <div class="voucher-panel-title serif-title" style="justify-content: space-between;">
        <span>记账凭证 <span class="badge" id="voucher-no-badge">-</span></span>
        <span style="font-size:12px; color:var(--text-muted); font-weight:400;">第 <span id="voucher-attach-count">0</span> 张附件</span>
      </div>
      <div style="overflow:auto;">
        <table class="voucher-entry-table">
          <thead>
            <tr>
              <th width="22%">会计科目</th>
              <th width="22%">辅助核算 (银行/往来)</th>
              <th width="22%">摘要</th>
              <th width="14%" class="num">借方</th>
              <th width="14%" class="num">贷方</th>
              <th width="6%"></th>
            </tr>
          </thead>
          <tbody id="voucher-entries-body"></tbody>
        </table>
      </div>
      <button type="button" class="btn-add-entry" onclick="addVoucherEntry()">+ 添加分录行</button>

      <div class="voucher-totals">
        <div class="voucher-total-item">借方合计：<strong id="voucher-debit-total">0.00</strong></div>
        <div class="voucher-total-item">贷方合计：<strong id="voucher-credit-total">0.00</strong></div>
        <div class="voucher-balance-status bad" id="voucher-balance-status">借贷不平</div>
      </div>
    </div>
  </div>

  <div class="voucher-submit-bar">
    <div class="hint" id="voucher-submit-hint">说明：银行存款(1002)、应收账款(1122)、应付账款(2202) 选择辅助核算后支持右键穿透联查。</div>
    <div style="display:flex; gap:12px;">
      <button type="button" class="btn" onclick="resetVoucherForm()">重置</button>
      <button type="button" class="btn btn-primary" id="voucher-btn-submit">提交审核</button>
    </div>
  </div>

  <div class="voucher-list-panel">
    <div class="voucher-panel-title serif-title" style="justify-content: space-between;">
      <span>凭证列表与审核</span>
      <span style="font-size:12px; color:var(--text-muted); font-weight:400;">共 <span id="voucher-list-count">0</span> 张凭证</span>
    </div>

    <div class="voucher-period-bar">
      <span style="font-weight:500;">会计期间管理：</span>
      <input type="month" id="voucher-period-input" style="width:auto;">
      <button type="button" class="btn-mini" onclick="closePeriod()">关闭该期间</button>
      <span style="color:var(--text-muted);">已关闭：</span>
      <span id="voucher-closed-periods">无</span>
    </div>

    <div style="overflow:auto;">
      <table class="voucher-table">
        <thead>
          <tr><th>凭证号</th><th>日期</th><th>类型</th><th>摘要</th><th class="num" style="text-align:right;">借方</th><th class="num" style="text-align:right;">贷方</th><th>来源单</th><th>状态</th><th>操作</th></tr>
        </thead>
        <tbody id="voucher-list-body"></tbody>
      </table>
    </div>
  </div>
</div>

<!-- ================= 会计记账 ================= -->
<div id="module-accounting" class="module">
  <div class="module-header">
    <div>
      <div class="home-subtitle-top">ACCOUNTING · 记账与查询</div>
      <h1 class="serif-title">会计记账 · 账簿 · 报表</h1>
      <p style="max-width:760px;">从已审核凭证记账，生成科目汇总表、明细账、总账、试算平衡表和财务报表，支持按科目快速检索关联凭证[cite: 1]。</p>
    </div>
  </div>

  <div class="acc-period-bar">
    <span style="font-weight:500;">会计期间</span>
    <select id="acc-period-select" onchange="onAccPeriodChange(this.value)" style="width:auto; min-width:180px;"></select>
    <button type="button" class="btn-mini" onclick="refreshAccounting()">刷新</button>
    <span style="color:var(--text-muted); font-size:12px; margin-left:auto;" id="acc-period-summary"></span>
  </div>

  <div class="acc-tabs" id="acc-tabs">
    <div class="acc-tab active" data-tab="pending">待记账凭证 <span class="acc-badge" id="acc-badge-pending">0</span></div>
    <div class="acc-tab" data-tab="summary">科目汇总表</div>
    <div class="acc-tab" data-tab="detail">明细账</div>
    <div class="acc-tab" data-tab="ledger">总账</div>
    <div class="acc-tab" data-tab="trial">试算平衡表</div>
    <div class="acc-tab" data-tab="report">财务报表</div>
  </div>

  <div class="acc-tab-panel active" id="acc-panel-pending">
    <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:16px; flex-wrap:wrap; gap:12px;">
      <p style="font-size:13px; color:var(--text-muted); margin:0;">仅展示「审核通过」且未记账的凭证，同一凭证只允许入账一次[cite: 1]。</p>
      <button type="button" class="btn btn-primary" onclick="postAllPending()">一键记账全部</button>
    </div>
    <div style="overflow:auto;">
      <table class="acc-table">
        <thead><tr><th>凭证号</th><th>日期</th><th>业务类型</th><th>摘要</th><th style="text-align:right;">借方</th><th style="text-align:right;">贷方</th><th>状态</th><th>操作</th></tr></thead>
        <tbody id="acc-pending-body"></tbody>
      </table>
    </div>
  </div>

  <div class="acc-tab-panel" id="acc-panel-summary">
    <div style="overflow:auto;">
      <table class="acc-table">
        <thead><tr><th>科目编码</th><th>科目名称</th><th style="text-align:right;">本期借方发生</th><th style="text-align:right;">本期贷方发生</th></tr></thead>
        <tbody id="acc-summary-body"></tbody>
        <tfoot id="acc-summary-foot"></tfoot>
      </table>
    </div>
  </div>

  <div class="acc-tab-panel" id="acc-panel-detail">
    <div class="acc-detail-bar">
      <span style="font-size:13px; color:var(--text-muted);">科目：</span>
      <select id="acc-detail-account" style="width:auto; min-width:240px;"></select>
    </div>
    <div style="overflow:auto;">
      <table class="acc-table">
        <thead><tr><th>日期</th><th>凭证号</th><th>摘要</th><th style="text-align:right;">借方</th><th style="text-align:right;">贷方</th><th style="text-align:right;">余额</th></tr></thead>
        <tbody id="acc-detail-body"></tbody>
      </table>
    </div>
  </div>

  <div class="acc-tab-panel" id="acc-panel-ledger">
    <div style="overflow:auto;">
      <table class="acc-table">
        <thead><tr><th>科目编码</th><th>科目名称</th><th style="text-align:right;">期初余额</th><th style="text-align:right;">本期借方</th><th style="text-align:right;">本期贷方</th><th style="text-align:right;">期末余额</th></tr></thead>
        <tbody id="acc-ledger-body"></tbody>
      </table>
    </div>
  </div>

  <div class="acc-tab-panel" id="acc-panel-trial">
    <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:16px; flex-wrap:wrap; gap:12px;">
      <p style="font-size:13px; color:var(--text-muted); margin:0;">试算平衡校验：期初、本期、期末的借方合计应等于贷方合计[cite: 1]。</p>
      <button type="button" class="btn-mini" onclick="toggleOpeningEditor()">期初余额设置</button>
    </div>
    <div id="acc-opening-editor" style="display:none; background:var(--bg); padding:16px; border-radius:var(--radius); margin-bottom:16px; border:1px solid var(--border);">
      <p style="font-size:12px; color:var(--text-muted); margin-bottom:12px;">为当前期间设置期初余额。仅填借方或贷方一侧，未填视为 0。期初合计应借贷相等[cite: 1]。</p>
      <div style="overflow:auto; max-height:400px;">
        <table class="acc-table">
          <thead><tr><th>科目编码</th><th>科目名称</th><th>期初借方</th><th>期初贷方</th></tr></thead>
          <tbody id="acc-opening-body"></tbody>
        </table>
      </div>
      <div style="margin-top:12px; text-align:right;">
        <button type="button" class="btn btn-primary" onclick="saveOpeningFromEditor()">保存期初余额</button>
      </div>
    </div>
    <div style="overflow:auto;">
      <table class="acc-table">
        <thead><tr><th>科目编码</th><th>科目名称</th><th style="text-align:right;">期初借方</th><th style="text-align:right;">期初贷方</th><th style="text-align:right;">本期借方</th><th style="text-align:right;">本期贷方</th><th style="text-align:right;">期末借方</th><th style="text-align:right;">期末贷方</th></tr></thead>
        <tbody id="acc-trial-body"></tbody>
        <tfoot id="acc-trial-foot"></tfoot>
      </table>
    </div>
    <div id="acc-trial-check" style="margin-top:16px;"></div>
  </div>

  <div class="acc-tab-panel" id="acc-panel-report">
    <div class="acc-report-subtabs">
      <div class="acc-report-subtab active" data-rtab="bs" onclick="switchReportTab('bs')">资产负债表</div>
      <div class="acc-report-subtab" data-rtab="is" onclick="switchReportTab('is')">利润表</div>
    </div>
    <div id="acc-report-content" style="margin-top:20px;"></div>
    <div style="margin-top:24px;">
      <label style="font-size:12px; color:var(--text-muted); display:block; margin-bottom:6px;">编表说明</label>
      <textarea id="acc-report-note" placeholder="填写本报表的编制说明..." style="min-height:80px; font-size:13px;"></textarea>
      <div style="text-align:right; margin-top:8px;">
        <button type="button" class="btn-mini" onclick="saveReportNote()">保存说明</button>
      </div>
    </div>
  </div>
</div>

<!-- ================= 科目余额表与穿透联查（融合：内置演示账套七列余额表 + 右键穿透 + Excel导入模式） ================= -->
<div id="module-drilldown" class="module">
  <div class="module-header">
    <div>
      <div class="home-subtitle-top">SUBJECT BALANCE & DRILLDOWN · 科目余额表</div>
      <h1 class="serif-title">科目余额表 · 穿透联查</h1>
      <p style="max-width:850px;">支持「本地账套实时余额表（右键穿透银行专项与单位三栏账）」与「外部 Excel 余额表及序时账导入穿透」双模式。</p>
    </div>
    <div class="section-actions">
      <div class="mode-switch">
        <button type="button" class="mode-switch-btn active" id="bal-mode-live" onclick="switchBalanceMode('live')">账套七列余额表</button>
        <button type="button" class="mode-switch-btn" id="bal-mode-excel" onclick="switchBalanceMode('excel')">Excel 导入穿透</button>
      </div>
    </div>
  </div>

  <!-- 模式 A：七列科目余额表（默认展示，内置演示账套 + 与制证记账实时联动） -->
  <div id="bal-view-live">
    <div class="panel">
      <div class="hint-bar">
        <span>💡 提示：在任意科目行上点击<strong>鼠标右键</strong>（或双击），可穿透联查银行专项分类明细、单位往来余额表或三栏明细账</span>
        <div class="section-actions">
          <button type="button" class="btn-mini" id="btn-reset-demo-vouchers">重置演示凭证</button>
          <button type="button" class="btn-mini btn-mini-primary" id="btn-export-live-balance">导出余额表（CSV）</button>
        </div>
      </div>
      <div style="overflow:auto;">
        <table class="fin-table">
          <thead>
            <tr>
              <th width="90">科目编码</th>
              <th width="200">科目名称</th>
              <th width="80">借/贷方向</th>
              <th class="num" width="140">期初余额</th>
              <th class="num" width="150">本期借方发生额</th>
              <th class="num" width="150">本期贷方发生额</th>
              <th class="num" width="140">期末余额</th>
            </tr>
          </thead>
          <tbody id="live-balance-body"></tbody>
          <tfoot id="live-balance-foot"></tfoot>
        </table>
      </div>
      <div class="table-footnote" id="live-balance-note"></div>
    </div>
  </div>

  <!-- 模式 B：外部 Excel 导入穿透（原代码2完整保留） -->
  <div id="bal-view-excel" style="display:none;">
    <div class="panel">
      <h3 style="margin-bottom:12px;">Excel 数据导入</h3>
      <div class="toolbar">
        <label style="flex:1;min-width:220px;">科目余额表（必需）<input id="drill-balance-file" type="file" accept=".xlsx,.xls"></label>
        <label style="flex:1;min-width:220px;">序时账／凭证明细（可选）<input id="drill-journal-file" type="file" accept=".xlsx,.xls"></label>
        <button type="button" class="btn btn-primary" id="drill-import">识别并导入</button>
        <button type="button" class="btn" id="drill-clear">清空导入</button>
      </div>
      <div id="drill-source-status" class="note-text">尚未导入外部 Excel 数据。页面不会生成虚构余额[cite: 1]。</div>
    </div>
    <div class="panel">
      <div class="flex-between"><h3>科目余额层级</h3><div class="section-actions">
        <input id="drill-search" placeholder="按编码或名称定位" style="width:220px">
        <button type="button" class="btn-mini" id="drill-expand">全部展开</button><button type="button" class="btn-mini" id="drill-collapse">全部折叠</button>
      </div></div>
      <div style="overflow:auto;margin-top:14px"><table class="acc-table"><thead><tr><th>科目编码</th><th>科目名称</th><th>期初余额</th><th>本期借方</th><th>本期贷方</th><th>期末余额</th><th>方向</th><th>联查</th></tr></thead><tbody id="drill-tree-body"><tr><td colspan="8">请导入科目余额表[cite: 1]</td></tr></tbody></table></div>
    </div>
    <div class="panel">
      <div class="acc-report-subtabs">
        <button type="button" class="acc-report-subtab active" data-drill-report="profit">利润表</button>
        <button type="button" class="acc-report-subtab" data-drill-report="cashflow">现金流量表</button>
      </div>
      <div class="toolbar"><label>会计期间 <input type="month" id="drill-period" style="width:170px"></label>
        <button type="button" class="btn btn-primary" id="drill-build-report">生成报表</button>
        <button type="button" class="btn" id="drill-export-report">导出当前报表</button></div>
      <div id="drill-report-note" class="note-text">报表取数基于已导入的凭证明细；缺少映射或期间数据时将标记待补全[cite: 1]。</div>
      <div style="overflow:auto"><table class="acc-table"><thead><tr><th>报表项目</th><th>本期金额</th><th>取数状态／说明</th></tr></thead><tbody id="drill-report-body"><tr><td colspan="3">请选择报表并生成[cite: 1]</td></tr></tbody></table></div>
    </div>
  </div>
</div>

<!-- ================= 应收账款（月结核对与年末函证） ================= -->
<div id="module-ar" class="module">
  <div class="module-header">
    <div>
      <div class="home-subtitle-top">AR · 应收账款</div>
      <h1 class="serif-title">月结核对 · 年末函证</h1>
      <p style="max-width:800px;">汇总客户应收及回款，核实账面余额，编制并跟踪年末询证函。余额按账务记录计算，回款分配总额不得超过实际到账额[cite: 1]。</p>
    </div>
    <div style="display:flex; gap:12px;">
      <input type="month" id="ar-period" style="width:150px;">
      <button type="button" class="btn btn-primary" onclick="generateConfirmation()">生成/更新询证函</button>
    </div>
  </div>

  <div class="module-grid">
    <div class="card-block">
      <h3>客户应收明细与函证跟踪</h3>
      <div style="overflow:auto;">
        <table class="data-table">
          <thead>
            <tr>
              <th>客户名称</th>
              <th style="text-align:right;">期末账面余额</th>
              <th style="text-align:right;">待冲销回款</th>
              <th>函证状态</th>
              <th style="text-align:right;">回函差异</th>
              <th>操作</th>
            </tr>
          </thead>
          <tbody id="ar-customer-list"></tbody>
        </table>
      </div>
    </div>

    <div class="card-block">
      <h3>回款冲销与核实记录</h3>
      <p class="note-text" style="margin-bottom:12px;">为保证余额准确，需将未分配回款与应收明细进行冲销核销[cite: 1]。</p>
      <div style="margin-bottom:16px;">
        <select id="ar-unallocated" style="width:100%; margin-bottom:8px;"></select>
        <button type="button" class="btn" style="width:100%;" onclick="allocatePayment()">核销至选中客户</button>
      </div>
      <table class="data-table">
        <thead><tr><th>客户</th><th>冲销金额</th><th>冲销时间</th><th>核实状态</th></tr></thead>
        <tbody id="ar-allocations"></tbody>
      </table>
    </div>
  </div>

  <div class="card-block">
    <h3>函证差异表 (仅年末生成)</h3>
    <table class="data-table">
      <thead>
        <tr><th>客户名称</th><th style="text-align:right;">账面余额</th><th style="text-align:right;">回函金额</th><th style="text-align:right;">差异额</th><th>差异原因/核实依据</th><th>状态</th></tr>
      </thead>
      <tbody id="ar-confirmation-diffs"></tbody>
    </table>
  </div>
</div>

<!-- ================= 税务管理 ================= -->
<div id="module-tax" class="module">
  <div class="module-header">
    <div>
      <div class="home-subtitle-top">TAX · 税务</div>
      <h1 class="serif-title">计税测算 · 申报归档</h1>
      <p style="max-width:800px;">处理月末增值税及附加测算、季末企业所得税测算，跟踪申报任务与完税证明[cite: 1]。</p>
    </div>
    <div style="display:flex; gap:12px;">
      <input type="month" id="tax-period" style="width:150px;">
      <button type="button" class="btn btn-primary" onclick="calculateTax()">预览测算底稿</button>
    </div>
  </div>

  <div class="module-grid">
    <div class="card-block">
      <h3>1. 计税测算底稿</h3>
      <div style="margin-bottom:16px;">
        <div class="flex-between" style="margin-bottom:8px;">
          <span><span class="tax-badge vat">增值税</span> 销项-进项</span>
          <strong id="tax-vat-amount">待测算</strong>
        </div>
        <div class="note-text" id="tax-vat-detail">取数依据：开票汇总、进项抵扣明细[cite: 1]</div>
      </div>
      <div style="margin-bottom:16px;">
        <div class="flex-between" style="margin-bottom:8px;">
          <span><span class="tax-badge surtax">附加税费</span> 城建/教育/地方教育</span>
          <strong id="tax-surtax-amount">待测算</strong>
        </div>
      </div>
      <div style="margin-bottom:16px;">
        <div class="flex-between" style="margin-bottom:8px;">
          <span><span class="tax-badge cit">企业所得税</span> 利润总额×税率</span>
          <strong id="tax-cit-amount">待测算</strong>
        </div>
      </div>
      <div id="tax-missing-items" style="display:none; background:#FEF3C7; color:#B45309; padding:12px; border-radius:8px; font-size:12px; margin-top:12px;">
        <strong>⚠️ 缺项提示：</strong><span id="tax-missing-text"></span>
        <button type="button" class="btn-mini" style="margin-left:12px;" onclick="fixTaxMissing()">去处理</button>
      </div>
      <div style="text-align:right; margin-top:20px;">
        <button type="button" class="btn btn-primary" onclick="saveTaxResult()">复核并保存结果</button>
      </div>
    </div>

    <div class="card-block">
      <h3>2. 申报任务与完税归档</h3>
      <div class="filter-bar" style="margin-bottom:12px; padding:10px; background:var(--bg); border-radius:var(--radius); border:1px solid var(--border); display:flex; justify-content:space-between;">
        <span style="font-size:12px; font-weight:500;">所选税期</span>
        <span style="font-size:12px; color:var(--diff);">截止日以任务记录为准[cite: 1]</span>
      </div>
      <table class="data-table">
        <thead><tr><th>任务</th><th>状态</th><th>回执/证明</th><th>操作</th></tr></thead>
        <tbody id="tax-task-list"></tbody>
      </table>
    </div>
  </div>
</div>

<!-- ================= 期末结转 ================= -->
<div id="module-cost" class="module">
  <div class="module-header">
    <div>
      <div class="home-subtitle-top">CLOSING · 结转</div>
      <h1 class="serif-title">期末结转 · 财务档案</h1>
      <p style="max-width:800px;">完成月末损益结转、成本折旧、资产摊销，以及结账前检查与财务档案整理[cite: 1]。</p>
    </div>
    <div style="display:flex; gap:12px;">
      <input type="month" id="closing-period" style="width:150px;">
      <button type="button" class="btn btn-primary" onclick="executeSettlement()">执行期末结转</button>
    </div>
  </div>

  <div class="module-grid">
    <div class="card-block">
      <h3>成本与计提摊销</h3>
      <p class="note-text" style="margin-bottom:16px;">系统自动计算：销售成本结转、固定资产折旧、无形资产摊销、工资计提[cite: 1]。</p>
      <div class="list-item"><span>销售成本结转</span><span id="cost-cogs">待计算</span></div>
      <div class="list-item"><span>固定资产折旧</span><span id="cost-depr">待计算</span></div>
      <div class="list-item"><span>无形资产摊销</span><span id="cost-amort">待计算</span></div>
      <div class="list-item"><span>工资社保计提</span><span id="cost-payroll">待计算</span></div>
      <div style="margin-top:16px;">
        <button type="button" class="btn" style="width:100%;" onclick="calculateCost()">刷新作业汇总</button>
      </div>
      <div id="cost-missing" style="display:none; color:var(--diff); font-size:12px; margin-top:8px;">
        ⚠️ 缺少存货计价依据，暂无法计算结转成本[cite: 1]。
      </div>
    </div>

    <div class="card-block">
      <h3>结账检查与档案</h3>
      <div class="list-item">
        <span>1. 待记账凭证</span>
        <span id="check-vouchers" class="status-pill WARNING">检查中...</span>
      </div>
      <div class="list-item">
        <span>2. 银行对账结果</span>
        <span id="check-recon" class="status-pill WARNING">检查中...</span>
      </div>
      <div class="list-item">
        <span>3. 损益结转凭证</span>
        <span id="check-transfer" class="status-pill PENDING">未生成</span>
      </div>
      <div class="list-item">
        <span>4. 财务报表复核</span>
        <span id="check-report" class="status-pill PENDING">未生成</span>
      </div>
      <div class="section-actions closing-actions" id="closing-actions">
        <button type="button" class="btn" onclick="checkClosing()">重新检查</button>
        <button type="button" class="btn btn-primary" onclick="finalizeClosing()">正式结账并归档</button>
      </div>
    </div>
  </div>
</div>

    </div><!-- /.workspace-content -->
  </main>
</div><!-- /.app-shell -->

<!-- 右键上下文菜单（科目余额表穿透专用） -->
<div id="ctx-menu" class="ctx-menu" role="menu"></div>

<!-- 账套内穿透联查抽屉（银行专项 / 单位往来余额 / 单科目三栏账） -->
<div id="live-drawer-mask" class="drawer-mask">
  <div class="drawer-panel">
    <div class="drawer-head">
      <div>
        <h2 id="live-drawer-title" class="serif-title">明细</h2>
        <p class="drawer-sub" id="live-drawer-sub"></p>
      </div>
      <button type="button" class="btn" id="live-drawer-close" style="padding:7px 16px; font-size:13px;">关闭</button>
    </div>
    <div class="drawer-body" id="live-drawer-body"></div>
  </div>
</div>

<!-- Excel 导入模式科目联查抽屉 -->
<div id="drill-drawer" style="display:none;position:fixed;z-index:2400;inset:0;background:#0005;">
 <div style="position:absolute;right:0;top:0;height:100%;width:min(900px,96vw);overflow:auto;background:var(--card);padding:28px;box-shadow:-12px 0 40px #0002;">
  <div class="flex-between"><h2 id="drill-drawer-title">科目明细</h2><button type="button" class="btn" id="drill-drawer-close">关闭</button></div>
  <p id="drill-drawer-meta" class="note-text"></p>
  <div class="toolbar"><label>起始日期<input id="drill-date-from" type="date"></label><label>截止日期<input id="drill-date-to" type="date"></label>
   <input id="drill-keyword" placeholder="凭证号／摘要关键词"><button type="button" class="btn btn-primary" id="drill-apply-filter">筛选</button><button type="button" class="btn" id="drill-export-ledger">导出 Excel</button></div>
  <div style="overflow:auto"><table class="acc-table"><thead><tr><th>日期</th><th>凭证号</th><th>科目编码</th><th>摘要</th><th>借方</th><th>贷方</th><th>方向</th><th>余额</th></tr></thead><tbody id="drill-ledger-body"></tbody></table></div>
  <p id="drill-ledger-summary" class="note-text"></p>
 </div>
</div>

<!-- 通用确认与输入模态框 -->
<div id="ui-dialog-modal" class="modal-overlay" role="dialog" aria-modal="true" aria-labelledby="ui-dialog-title">
  <div class="modal-card">
    <h2 id="ui-dialog-title" class="modal-title">操作确认</h2>
    <p id="ui-dialog-msg" class="modal-desc"></p>
    <input id="ui-dialog-input" type="text" style="display:none; margin-bottom:20px;">
    <div class="modal-actions">
      <button type="button" class="btn" id="ui-dialog-cancel">取消</button>
      <button type="button" class="btn btn-primary" id="ui-dialog-ok">确认</button>
    </div>
  </div>
</div>

<!-- 安全退出 / 备份模态框 -->
<div id="safe-exit-modal" class="modal-overlay" role="dialog" aria-modal="true" aria-labelledby="safe-exit-title">
  <div class="modal-card">
    <h2 id="safe-exit-title" class="modal-title">安全退出与备份提醒</h2>
    <p class="modal-desc">是否立即执行本地账套数据备份？导出 JSON 备份文件可防止浏览器缓存清理导致账套记录丢失[cite: 1]。</p>
    <div class="modal-actions">
      <button type="button" class="btn" id="modal-exit-cancel">取消</button>
      <button type="button" class="btn btn-danger" id="modal-exit-direct">直接退出</button>
      <button type="button" class="btn btn-primary" id="modal-exit-backup">立即备份并下载</button>
    </div>
  </div>
</div>

<script>
/* ================= 1. 左侧树形导航 + 多标签页工作区 ================= */
const NAV_TREE = [
  {
    id: 'gl',
    num: '1',
    title: '总账模块',
    children: [
      { key: 'gl-voucher', label: '凭证制单', moduleId: 'module-voucher' },
      { key: 'gl-query', label: '凭证查询与记账', moduleId: 'module-accounting' },
      { key: 'gl-balance', label: '科目余额表', moduleId: 'module-drilldown' },
      { key: 'gl-detail', label: '三栏明细账', placeholder: '三栏明细账独立查询页正在规划中，可在「科目余额表」右键任意科目直接穿透查看三栏明细账。' },
      { key: 'gl-print', label: '账簿打印', placeholder: '总账、明细账及凭证套打排版模块正在规划中，当前可直接使用浏览器打印各账表页面[cite: 1]。' }
    ]
  },
  {
    id: 'ar',
    num: '2',
    title: '往来管理',
    children: [
      { key: 'ar-subledger', label: '单位往来三栏账', moduleId: 'module-ar-subledger' },
      { key: 'ar-ledger', label: '客户核对与函证', moduleId: 'module-ar' },
      { key: 'ar-balance', label: '单位往来余额表', placeholder: '按客户/供应商维度的往来余额汇总表正在规划中，可在「科目余额表」右键 1122 / 2202 打开单位往来余额表。' },
      { key: 'ar-baddebt', label: '往来坏账分析', placeholder: '账龄分析与坏账准备计提测算模块正在规划中[cite: 1]。' }
    ]
  },
  {
    id: 'fund',
    num: '3',
    title: '资金模块',
    children: [
      { key: 'fund-bank', label: '出纳与银行账户', moduleId: 'module-cashier' },
      { key: 'fund-loan', label: '银行流水对账', moduleId: 'module-recon' },
      { key: 'fund-internal', label: '内部户开立', placeholder: '集团内部资金账户开立与额度管理功能正在规划中[cite: 1]。' }
    ]
  },
  {
    id: 'fa',
    num: '4',
    title: '固定资产',
    children: [
      { key: 'fa-card', label: '资产卡片管理', placeholder: '固定资产卡片建档、变动及条码管理功能正在规划中[cite: 1]。' },
      { key: 'fa-depr', label: '月度折旧计提', placeholder: '月度折旧自动计提与折旧费用分配表功能正在规划中[cite: 1]。' }
    ]
  },
  {
    id: 'tax',
    num: '5',
    title: '税务管理与导航',
    children: [
      { key: 'tax-calc', label: '智能算税器', moduleId: 'module-tax' },
      { key: 'tax-wizard', label: '申报流程向导', placeholder: '分税种申报步骤指引与校验向导正在规划中[cite: 1]。' },
      { key: 'tax-calendar', label: '申报日历', placeholder: '月度/季度征期日历与逾期提醒功能正在规划中[cite: 1]。' },
      { key: 'tax-history', label: '历史与设置', placeholder: '历史完税归档查询与税率参数设置功能正在规划中[cite: 1]。' }
    ]
  },
  {
    id: 'fin',
    num: '6',
    title: '财务与成本分析',
    children: [
      { key: 'fin-cost', label: '期末结转与成本', moduleId: 'module-cost' },
      { key: 'fin-report', label: '简易报表生成', placeholder: '一键生成资产负债表、利润表与现金流量表模板正在规划中，可先在「科目余额表」或「凭证查询与记账」下查看报表[cite: 1]。' },
      { key: 'fin-kpi', label: '自定义指标分析', placeholder: '自定义毛利率、费用率及财务健康度指标看板正在规划中[cite: 1]。' }
    ]
  }
];

const EXTRA_TABS = {
  'home': { key: 'home', label: '工作台首页', moduleId: 'module-home', pinned: true },
  'recon': { key: 'recon', label: '数据对账', moduleId: 'module-recon', groupTitle: '资金模块' },
  'invoice': { key: 'invoice', label: '智能开票', moduleId: 'module-invoice', groupTitle: '开票业务' }
};

function findNavMeta(key) {
  if (EXTRA_TABS[key]) return EXTRA_TABS[key];
  for (const g of NAV_TREE) {
    const item = g.children.find(c => c.key === key);
    if (item) return { ...item, groupId: g.id, groupTitle: g.title };
  }
  return null;
}

const openTabs = [ { key: 'home', label: '工作台首页', pinned: true } ];
let activeTabKey = 'home';

function renderSidebarTree() {
  const container = document.getElementById('sidebar-tree');
  container.replaceChildren();
  NAV_TREE.forEach((group, idx) => {
    const groupEl = document.createElement('div');
    groupEl.className = 'tree-group' + (idx === 0 ? ' open' : '');
    groupEl.dataset.groupId = group.id;

    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'tree-group-btn';
    btn.innerHTML = `<span class="tree-group-label"><span class="tree-group-num">${group.num}</span><span>${escapeHtml(group.title)}</span></span><span class="tree-chevron">▸</span>`;
    btn.addEventListener('click', () => {
      const willOpen = !groupEl.classList.contains('open');
      container.querySelectorAll('.tree-group').forEach(g => g.classList.remove('open'));
      if (willOpen) groupEl.classList.add('open');
    });

    const childrenEl = document.createElement('div');
    childrenEl.className = 'tree-children';
    group.children.forEach(item => {
      const itemBtn = document.createElement('button');
      itemBtn.type = 'button';
      itemBtn.className = 'tree-item-btn';
      itemBtn.dataset.tabKey = item.key;
      itemBtn.innerHTML = `<span>${escapeHtml(item.label)}</span>` + (item.placeholder ? '<span class="tree-item-tag">待建</span>' : '');
      itemBtn.addEventListener('click', () => {
        openWorkspaceTab(item.key);
        closeMobileSidebar();
      });
      childrenEl.appendChild(itemBtn);
    });

    groupEl.append(btn, childrenEl);
    container.appendChild(groupEl);
  });
}

function renderWorkspaceTabs() {
  const list = document.getElementById('workspace-tabs-list');
  list.replaceChildren();
  openTabs.forEach(tab => {
    const el = document.createElement('div');
    el.className = 'ws-tab' + (tab.key === activeTabKey ? ' active' : '');
    el.setAttribute('role', 'tab');
    el.setAttribute('aria-selected', tab.key === activeTabKey ? 'true' : 'false');
    el.innerHTML = `<span class="ws-tab-dot"></span><span>${escapeHtml(tab.label)}</span>`;
    el.addEventListener('click', () => activateWorkspaceTab(tab.key));

    if (!tab.pinned) {
      const closeBtn = document.createElement('button');
      closeBtn.type = 'button';
      closeBtn.className = 'ws-tab-close';
      closeBtn.setAttribute('aria-label', '关闭标签页 ' + tab.label);
      closeBtn.textContent = '×';
      closeBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        closeWorkspaceTab(tab.key);
      });
      el.appendChild(closeBtn);
    }
    list.appendChild(el);
  });
}

function openWorkspaceTab(key) {
  const meta = findNavMeta(key);
  if (!meta) return;
  if (!openTabs.some(t => t.key === key)) {
    openTabs.push({ key: meta.key, label: meta.label, pinned: !!meta.pinned });
  }
  activateWorkspaceTab(key);
}

function closeWorkspaceTab(key) {
  const idx = openTabs.findIndex(t => t.key === key);
  if (idx === -1 || openTabs[idx].pinned) return;
  openTabs.splice(idx, 1);
  if (activeTabKey === key) {
    const nextTab = openTabs[Math.max(0, idx - 1)] || openTabs[0];
    activateWorkspaceTab(nextTab.key);
  } else {
    renderWorkspaceTabs();
  }
}

function activateWorkspaceTab(key) {
  const meta = findNavMeta(key);
  if (!meta) return;
  activeTabKey = key;

  const targetId = meta.moduleId || 'module-placeholder';
  if (!meta.moduleId && meta.placeholder) {
    document.getElementById('placeholder-module-name').textContent = meta.groupTitle || '功能模块';
    document.getElementById('placeholder-title').textContent = meta.label;
    document.getElementById('placeholder-desc').textContent = meta.placeholder;
  }

  switchModule(targetId);
  renderWorkspaceTabs();

  document.getElementById('nav-home-btn').classList.toggle('active', key === 'home');
  document.querySelectorAll('.tree-item-btn').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.tabKey === key);
  });
  if (meta.groupId) {
    document.querySelectorAll('.tree-group').forEach(g => {
      g.classList.toggle('open', g.dataset.groupId === meta.groupId);
    });
  }
}

function switchModule(targetId) {
  const target = document.getElementById(targetId);
  if (!target) return;
  const currentActive = document.querySelector('.module.active');
  if (currentActive && currentActive !== target) {
    animation.killTweensOf(currentActive);
    currentActive.classList.remove('active');
    currentActive.style.display = 'none';
  }
  target.classList.add('active');
  target.style.display = 'block';
  animation.killTweensOf(target);
  animation.fromTo(target, { opacity: 0, y: 12 }, { opacity: 1, y: 0, duration: 0.35, ease: "power2.out" });
  window.scrollTo({ top: 0, behavior: 'smooth' });
  if (targetId === 'module-accounting' && window.refreshAccounting) refreshAccounting();
  if (targetId === 'module-drilldown' && window.renderLiveBalance) renderLiveBalance();
  if (targetId === 'module-ar-subledger' && window.renderARSubledgerModule) renderARSubledgerModule();
  if (targetId === 'module-ar' && window.loadARData) loadARData();
  if (targetId === 'module-tax' && window.loadTaxData) loadTaxData();
  if (targetId === 'module-cost' && window.checkClosing) checkClosing();
}

function closeMobileSidebar() {
  document.getElementById('app-sidebar').classList.remove('open');
  document.getElementById('sidebar-backdrop').classList.remove('open');
}

const animation = window.gsap || {
  killTweensOf() {},
  set(target, values) { target.style.opacity = values.opacity; },
  fromTo(target, from, to) { if (target instanceof HTMLElement) target.style.opacity = to.opacity; }
};

/* ================= 2. 存储引擎 ================= */
const DB_NAME = 'FinanceWorkbenchCleanBooks';
const DB_VERSION = 5;
const STORE_WORK = 'Workpapers';
const STORE_NAME = 'Invoices';
const STORE_CASHIER = 'CashierRecords';
const STORE_VOUCHER = 'VoucherRecords';
const LS_KEY = 'clean_finance_ledger_fallback';
const LS_KEY_CASHIER = 'clean_finance_cashier_fallback';
const LS_KEY_VOUCHER = 'clean_finance_voucher_fallback';

let storageMode = 'MEMORY';
let db = null;
const memoryData = { 'Invoices': [], 'CashierRecords': [], 'VoucherRecords': [], 'Workpapers': [] };

function lsKeyFor(s) { return s === STORE_WORK ? 'clean_finance_workpapers' : s === STORE_CASHIER ? LS_KEY_CASHIER : (s === STORE_VOUCHER ? LS_KEY_VOUCHER : LS_KEY); }
function pkFor(s) { return s === STORE_NAME ? 'invoiceNo' : 'id'; }

function safeLSGet(key, fallback = '[]') {
  try {
    const val = localStorage.getItem(key);
    return val === null ? fallback : val;
  } catch (e) {
    return fallback;
  }
}
function safeLSSet(key, value) {
  try {
    localStorage.setItem(key, value);
    return true;
  } catch (e) {
    return false;
  }
}
function safeLSRemove(key) {
  try {
    localStorage.removeItem(key);
    return true;
  } catch (e) {
    return false;
  }
}

async function initStorage() {
  return new Promise((resolve) => {
    try {
      const request = indexedDB.open(DB_NAME, DB_VERSION);
      request.onblocked = () => showToast('请关闭其他工作台页面后重试');
      request.onerror = () => { tryLSFallback(resolve); };
      request.onsuccess = (e) => { db = e.target.result; storageMode = 'INDEXEDDB'; resolve(storageMode); };
      request.onupgradeneeded = (e) => {
        const dbInstance = e.target.result;
        if (!dbInstance.objectStoreNames.contains(STORE_WORK)) dbInstance.createObjectStore(STORE_WORK, { keyPath: 'id' });
        if (!dbInstance.objectStoreNames.contains(STORE_NAME)) dbInstance.createObjectStore(STORE_NAME, { keyPath: 'invoiceNo' });
        if (!dbInstance.objectStoreNames.contains(STORE_CASHIER)) dbInstance.createObjectStore(STORE_CASHIER, { keyPath: 'id' });
        if (!dbInstance.objectStoreNames.contains(STORE_VOUCHER)) dbInstance.createObjectStore(STORE_VOUCHER, { keyPath: 'id' });
      };
    } catch (err) { tryLSFallback(resolve); }
  });
}
function tryLSFallback(resolve) {
  try {
    localStorage.setItem('__t', '1');
    localStorage.removeItem('__t');
    storageMode = 'LOCALSTORAGE';
  } catch (e) {
    storageMode = 'MEMORY';
  }
  resolve(storageMode);
}
async function saveToStore(storeName, record) { return saveRecords(storeName, [record]); }
async function saveRecords(storeName, records, insert = false, completedDraft = null) {
  if (storageMode === 'MEMORY') {
    const pk = pkFor(storeName);
    for (const r of records) {
      const idx = memoryData[storeName].findIndex(x => x[pk] === r[pk]);
      if (insert && idx !== -1) throw new Error('记录编号已存在');
      if (idx !== -1) memoryData[storeName][idx] = r;
      else memoryData[storeName].push(r);
    }
    return;
  }
  if (storageMode === 'INDEXEDDB') {
    return new Promise((resolve, reject) => {
      const tx = db.transaction(completedDraft ? [storeName, STORE_WORK] : storeName, 'readwrite');
      tx.oncomplete = () => resolve();
      tx.onabort = () => reject(tx.error || new Error('保存失败'));
      const store = tx.objectStore(storeName);
      for (const record of records) store[insert ? 'add' : 'put'](record);
      if (completedDraft) tx.objectStore(STORE_WORK).delete('draft:' + completedDraft);
    });
  }
  const key = lsKeyFor(storeName), pk = pkFor(storeName);
  let parsed = [];
  try { parsed = JSON.parse(safeLSGet(key, '[]')); } catch (e) { parsed = []; }
  const data = new Map(parsed.map(row => [row[pk], row]));
  for (const record of records) {
    if (insert && data.has(record[pk])) throw new Error('记录编号已存在');
    data.set(record[pk], record);
  }
  if (!safeLSSet(key, JSON.stringify([...data.values()]))) throw new Error('本地存储空间不足');
  if (completedDraft) {
    const draftKey = lsKeyFor(STORE_WORK);
    let existing = [];
    try { existing = JSON.parse(safeLSGet(draftKey, '[]')); } catch (e) { existing = []; }
    const remaining = existing.filter(row => row.id !== 'draft:' + completedDraft);
    if (remaining.length !== existing.length) safeLSSet(draftKey, JSON.stringify(remaining));
  }
}
async function getAllFromStore(storeName) {
  if (storageMode === 'INDEXEDDB') {
    return new Promise((resolve, reject) => {
      const req = db.transaction([storeName], 'readonly').objectStore(storeName).getAll();
      req.onsuccess = () => resolve(req.result); req.onerror = (e) => reject(e.target.error);
    });
  } else if (storageMode === 'LOCALSTORAGE') {
    try { return JSON.parse(safeLSGet(lsKeyFor(storeName), '[]')); } catch (e) { return []; }
  } else {
    return [...memoryData[storeName]];
  }
}
async function clearStore(storeName) {
  if (storageMode === 'INDEXEDDB') {
    return new Promise((resolve, reject) => {
      const tx = db.transaction(storeName, 'readwrite');
      tx.oncomplete = () => resolve(); tx.onabort = () => reject(tx.error);
      tx.objectStore(storeName).clear();
    });
  } else if (storageMode === 'LOCALSTORAGE') {
    safeLSRemove(lsKeyFor(storeName));
  } else {
    memoryData[storeName] = [];
  }
}
async function saveRecordToDB(record) { return saveToStore(STORE_NAME, record); }
async function getAllRecordsFromDB() { return getAllFromStore(STORE_NAME); }
async function clearAllDB() { return clearStore(STORE_NAME); }

/* ================= 3. 全局工具与数字格式化 ================= */
function showToast(msg) {
  const t = document.getElementById('toast'); t.innerText = msg; t.style.visibility = 'visible';
  animation.killTweensOf(t); animation.fromTo(t, { opacity: 0, y: 20 }, { opacity: 1, y: 0, duration: 0.3, ease: "power2.out" });
  clearTimeout(t._timer);
  t._timer = setTimeout(() => {
    if (window.gsap) gsap.to(t, { opacity: 0, y: 20, duration: 0.3, onComplete: () => t.style.visibility = 'hidden' });
    else { t.style.opacity = 0; t.style.visibility = 'hidden'; }
  }, 3000);
}

function fmtNum(n) {
  const v = Number(n) || 0;
  const s = Math.abs(v).toLocaleString('zh-CN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  return v < 0 ? '-' + s : s;
}
function fmtOrBlank(n) {
  const v = Number(n) || 0;
  return v === 0 ? '' : fmtNum(v);
}

function uiConfirm(message, title = '操作确认') {
  return new Promise(resolve => {
    const modal = byId('ui-dialog-modal'), input = byId('ui-dialog-input');
    byId('ui-dialog-title').textContent = title;
    byId('ui-dialog-msg').textContent = message;
    input.style.display = 'none';
    modal.classList.add('open');
    const done = (res) => {
      modal.classList.remove('open');
      byId('ui-dialog-ok').onclick = null;
      byId('ui-dialog-cancel').onclick = null;
      resolve(res);
    };
    byId('ui-dialog-ok').onclick = () => done(true);
    byId('ui-dialog-cancel').onclick = () => done(false);
  });
}
function uiPrompt(message, defaultVal = '', title = '请输入') {
  return new Promise(resolve => {
    const modal = byId('ui-dialog-modal'), input = byId('ui-dialog-input');
    byId('ui-dialog-title').textContent = title;
    byId('ui-dialog-msg').textContent = message;
    input.style.display = 'block';
    input.value = defaultVal || '';
    modal.classList.add('open');
    setTimeout(() => input.focus(), 50);
    const done = (res) => {
      modal.classList.remove('open');
      byId('ui-dialog-ok').onclick = null;
      byId('ui-dialog-cancel').onclick = null;
      resolve(res);
    };
    byId('ui-dialog-ok').onclick = () => done(input.value);
    byId('ui-dialog-cancel').onclick = () => done(null);
  });
}

function escapeHtml(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }
function digitUppercase(n) {
  if (isNaN(n) || n == null) return '零元整';
  const fr = ['角','分'], d = ['零','壹','贰','叁','肆','伍','陆','柒','捌','玖'];
  const u = [['元','万','亿'],['','拾','佰','仟']];
  let head = n < 0 ? '欠' : ''; n = Math.abs(n); let s = '';
  for (let i = 0; i < fr.length; i++) s += (d[Math.floor(n * 10 * Math.pow(10, i)) % 10] + fr[i]).replace(/零./, '');
  s = s || '整'; n = Math.floor(n);
  for (let i = 0; i < u[0].length && n > 0; i++) {
    let p = '';
    for (let j = 0; j < u[1].length && n > 0; j++) { p = d[n % 10] + u[1][j] + p; n = Math.floor(n / 10); }
    s = p.replace(/(零.)*零$/, '').replace(/^$/, '零') + u[0][i] + s;
  }
  return head + s.replace(/(.)*零元/, '元').replace(/(零.)+/g, '零').replace(/^整$/, '零元整');
}

/* ================= 4. 统一科目体系、银行专项、往来单位与演示账套 ================= */
const VOUCHER_ACCOUNTS = [
  { code: '1001', name: '库存现金',     dir: '借', opening: 8650,     aux: 'none' },
  { code: '1002', name: '银行存款',     dir: '借', opening: 1286400,  aux: 'bank' },
  { code: '1122', name: '应收账款',     dir: '借', opening: 486200,   aux: 'customer' },
  { code: '1123', name: '预付账款',     dir: '借', opening: 0,        aux: 'none' },
  { code: '1221', name: '其他应收款',   dir: '借', opening: 0,        aux: 'none' },
  { code: '1401', name: '原材料',       dir: '借', opening: 320500,   aux: 'none' },
  { code: '1405', name: '库存商品',     dir: '借', opening: 0,        aux: 'none' },
  { code: '1601', name: '固定资产',     dir: '借', opening: 1500000,  aux: 'none' },
  { code: '1602', name: '累计折旧',     dir: '贷', opening: 250000,   aux: 'none' },
  { code: '2202', name: '应付账款',     dir: '贷', opening: 368900,   aux: 'supplier' },
  { code: '2203', name: '预收账款',     dir: '贷', opening: 0,        aux: 'none' },
  { code: '2211', name: '应付职工薪酬', dir: '贷', opening: 120000,   aux: 'none' },
  { code: '2221', name: '应交税费',     dir: '贷', opening: 45600,    aux: 'none' },
  { code: '2241', name: '其他应付款',   dir: '贷', opening: 0,        aux: 'none' },
  { code: '4001', name: '实收资本',     dir: '贷', opening: 2817250,  aux: 'none' },
  { code: '4103', name: '本年利润',     dir: '贷', opening: 0,        aux: 'none' },
  { code: '4104', name: '利润分配',     dir: '贷', opening: 0,        aux: 'none' },
  { code: '5001', name: '生产成本',     dir: '借', opening: 0,        aux: 'none' },
  { code: '6001', name: '主营业务收入', dir: '贷', opening: 0,        aux: 'none' },
  { code: '6401', name: '主营业务成本', dir: '借', opening: 0,        aux: 'none' },
  { code: '6601', name: '销售费用',     dir: '借', opening: 0,        aux: 'none' },
  { code: '6602', name: '管理费用',     dir: '借', opening: 0,        aux: 'none' },
  { code: '6603', name: '财务费用',     dir: '借', opening: 0,        aux: 'none' }
];
const SUBJ_MAP = Object.fromEntries(VOUCHER_ACCOUNTS.map(s => [s.code, s]));

const BANK_SPECIAL_ACCOUNTS = [
  { name: '工商银行基本户', bank: '中国工商银行', acct: '6222 0202 0000 1234 567', currency: 'CNY', opening: 800000 },
  { name: '招商银行一般户', bank: '招商银行',     acct: '7559 1234 5678 901',     currency: 'CNY', opening: 386400 },
  { name: '建设银行专用户', bank: '中国建设银行', acct: '3105 0000 8888 6666',    currency: 'CNY', opening: 100000 }
];
const AR_UNITS = ['华远科技有限公司', '蓝海贸易（上海）有限公司', '恒盛建设集团'];
const AP_UNITS = ['江南材料供应有限公司', '中电装备制造有限公司', '星辰物流有限公司'];
const AR_OPENING = { '华远科技有限公司': 210000, '蓝海贸易（上海）有限公司': 156200, '恒盛建设集团': 120000 };
const AP_OPENING = { '江南材料供应有限公司': 180000, '中电装备制造有限公司': 128900, '星辰物流有限公司': 60000 };

function getAuxOptionsForAccount(code) {
  if (code === '1002') return BANK_SPECIAL_ACCOUNTS.map(b => b.name);
  if (code === '1122') return AR_UNITS;
  if (code === '2202') return AP_UNITS;
  return [];
}

/* 演示凭证（首次初始化若 STORE_VOUCHER 为空则自动注入，与期初余额完全借贷平衡） */
const SEED_VOUCHERS = [
  { id: 'SEED-001', voucherNo: '记-202601-001', date: '2026-01-03', period: '2026-01', bizLabel: '银行往来', status: 'POSTED', source: '回单-0103', sourceType: '银行回单', totalDebit: 150000, totalCredit: 150000, entries: [
    { account: '1002', aux: '工商银行基本户', summary: '收到华远科技有限公司货款', debit: 150000, credit: 0 },
    { account: '1122', aux: '华远科技有限公司', summary: '收到华远科技有限公司货款', debit: 0, credit: 150000 }
  ]},
  { id: 'SEED-002', voucherNo: '记-202601-002', date: '2026-01-05', period: '2026-01', bizLabel: '采购成本', status: 'POSTED', source: '入库-0105', sourceType: '合同', totalDebit: 200000, totalCredit: 200000, entries: [
    { account: '1401', aux: '', summary: '采购原材料（江南材料）', debit: 200000, credit: 0 },
    { account: '2202', aux: '江南材料供应有限公司', summary: '采购原材料（江南材料）', debit: 0, credit: 200000 }
  ]},
  { id: 'SEED-003', voucherNo: '记-202601-003', date: '2026-01-06', period: '2026-01', bizLabel: '银行往来', status: 'POSTED', source: '付款-0106', sourceType: '银行回单', totalDebit: 120000, totalCredit: 120000, entries: [
    { account: '2202', aux: '江南材料供应有限公司', summary: '支付江南材料货款', debit: 120000, credit: 0 },
    { account: '1002', aux: '招商银行一般户', summary: '支付江南材料货款', debit: 0, credit: 120000 }
  ]},
  { id: 'SEED-004', voucherNo: '记-202601-004', date: '2026-01-08', period: '2026-01', bizLabel: '销售收入', status: 'POSTED', source: '发票-0108', sourceType: '发票', totalDebit: 113000, totalCredit: 113000, entries: [
    { account: '1122', aux: '蓝海贸易（上海）有限公司', summary: '销售商品确认收入（蓝海贸易）', debit: 113000, credit: 0 },
    { account: '6001', aux: '', summary: '销售商品确认收入（蓝海贸易）', debit: 0, credit: 100000 },
    { account: '2221', aux: '', summary: '销售商品销项税额', debit: 0, credit: 13000 }
  ]},
  { id: 'SEED-005', voucherNo: '记-202601-005', date: '2026-01-10', period: '2026-01', bizLabel: '采购成本', status: 'POSTED', source: '出库-0110', sourceType: '其他', totalDebit: 68000, totalCredit: 68000, entries: [
    { account: '6401', aux: '', summary: '结转销售成本', debit: 68000, credit: 0 },
    { account: '1401', aux: '', summary: '结转销售成本', debit: 0, credit: 68000 }
  ]},
  { id: 'SEED-006', voucherNo: '记-202601-006', date: '2026-01-12', period: '2026-01', bizLabel: '费用报销', status: 'POSTED', source: '报销-0112', sourceType: '报销单', totalDebit: 15600, totalCredit: 15600, entries: [
    { account: '6602', aux: '', summary: '支付办公及管理用费', debit: 15600, credit: 0 },
    { account: '1002', aux: '工商银行基本户', summary: '支付办公及管理用费', debit: 0, credit: 15600 }
  ]},
  { id: 'SEED-007', voucherNo: '记-202601-007', date: '2026-01-15', period: '2026-01', bizLabel: '银行往来', status: 'POSTED', source: '回单-0115', sourceType: '银行回单', totalDebit: 113000, totalCredit: 113000, entries: [
    { account: '1002', aux: '工商银行基本户', summary: '收回蓝海贸易货款', debit: 113000, credit: 0 },
    { account: '1122', aux: '蓝海贸易（上海）有限公司', summary: '收回蓝海贸易货款', debit: 0, credit: 113000 }
  ]},
  { id: 'SEED-008', voucherNo: '记-202601-008', date: '2026-01-18', period: '2026-01', bizLabel: '费用报销', status: 'POSTED', source: '水电-0118', sourceType: '报销单', totalDebit: 8400, totalCredit: 8400, entries: [
    { account: '6602', aux: '', summary: '支付水电费', debit: 8400, credit: 0 },
    { account: '1002', aux: '建设银行专用户', summary: '支付水电费', debit: 0, credit: 8400 }
  ]},
  { id: 'SEED-009', voucherNo: '记-202601-009', date: '2026-01-20', period: '2026-01', bizLabel: '银行往来', status: 'POSTED', source: '回单-0120', sourceType: '银行回单', totalDebit: 60000, totalCredit: 60000, entries: [
    { account: '1002', aux: '招商银行一般户', summary: '收到恒盛建设集团回款', debit: 60000, credit: 0 },
    { account: '1122', aux: '恒盛建设集团', summary: '收到恒盛建设集团回款', debit: 0, credit: 60000 }
  ]},
  { id: 'SEED-010', voucherNo: '记-202601-010', date: '2026-01-22', period: '2026-01', bizLabel: '工资计提', status: 'POSTED', source: '工资-0122', sourceType: '工资表', totalDebit: 120000, totalCredit: 120000, entries: [
    { account: '6602', aux: '', summary: '计提职工工资', debit: 120000, credit: 0 },
    { account: '2211', aux: '', summary: '计提职工工资', debit: 0, credit: 120000 }
  ]},
  { id: 'SEED-011', voucherNo: '记-202601-011', date: '2026-01-25', period: '2026-01', bizLabel: '工资计提', status: 'POSTED', source: '代发-0125', sourceType: '银行回单', totalDebit: 120000, totalCredit: 120000, entries: [
    { account: '2211', aux: '', summary: '发放职工工资', debit: 120000, credit: 0 },
    { account: '1002', aux: '工商银行基本户', summary: '发放职工工资', debit: 0, credit: 120000 }
  ]},
  { id: 'SEED-012', voucherNo: '记-202601-012', date: '2026-01-26', period: '2026-01', bizLabel: '银行往来', status: 'POSTED', source: '支票-0126', sourceType: '银行回单', totalDebit: 20000, totalCredit: 20000, entries: [
    { account: '1001', aux: '', summary: '提取备用金', debit: 20000, credit: 0 },
    { account: '1002', aux: '工商银行基本户', summary: '提取备用金', debit: 0, credit: 20000 }
  ]},
  { id: 'SEED-013', voucherNo: '记-202601-013', date: '2026-01-28', period: '2026-01', bizLabel: '采购成本', status: 'POSTED', source: '入库-0128', sourceType: '合同', totalDebit: 128900, totalCredit: 128900, entries: [
    { account: '1401', aux: '', summary: '采购原材料（中电装备）', debit: 128900, credit: 0 },
    { account: '2202', aux: '中电装备制造有限公司', summary: '采购原材料（中电装备）', debit: 0, credit: 128900 }
  ]},
  { id: 'SEED-014', voucherNo: '记-202601-014', date: '2026-01-30', period: '2026-01', bizLabel: '费用报销', status: 'POSTED', source: '运费-0130', sourceType: '发票', totalDebit: 12000, totalCredit: 12000, entries: [
    { account: '6602', aux: '', summary: '发生物流运费（星辰物流，未付）', debit: 12000, credit: 0 },
    { account: '2202', aux: '星辰物流有限公司', summary: '发生物流运费（星辰物流，未付）', debit: 0, credit: 12000 }
  ]}
];

async function ensureSeedVouchers() {
  const existing = await getAllFromStore(STORE_VOUCHER);
  const seededFlag = await work('meta:seeded_vouchers');
  if (!existing.length && !seededFlag) {
    await saveRecords(STORE_VOUCHER, SEED_VOUCHERS);
    await saveToStore(STORE_WORK, { id: 'meta:seeded_vouchers', done: true });
  }
}

const ACCOUNTS_INIT = [
  { id: 'CASH', name: '库存现金', kind: 'CASH' },
  { id: 'ICBC', name: '工商银行基本户', kind: 'BANK' },
  { id: 'CMB',  name: '招商银行一般户', kind: 'BANK' },
  { id: 'CCB',  name: '建设银行专用户', kind: 'BANK' }
];
let ACCOUNTS = (() => {
  try {
    const s = JSON.parse(safeLSGet('clean_cashier_accounts', 'null'));
    if (Array.isArray(s) && s.length) return s;
  } catch (e) {}
  return [...ACCOUNTS_INIT];
})();
function saveCashierAccounts() {
  safeLSSet('clean_cashier_accounts', JSON.stringify(ACCOUNTS));
}
function getAccountName(valStr) {
  if (!valStr) return '-';
  const byId = ACCOUNTS.find(x => x.id === valStr); if (byId) return byId.name;
  return valStr;
}

document.addEventListener('DOMContentLoaded', async () => {
  if ('scrollRestoration' in history) history.scrollRestoration = 'manual';
  window.scrollTo(0, 0);

  renderSidebarTree();
  renderWorkspaceTabs();

  document.getElementById('logo-home-btn').addEventListener('click', () => activateWorkspaceTab('home'));
  document.getElementById('nav-home-btn').addEventListener('click', () => { activateWorkspaceTab('home'); closeMobileSidebar(); });
  document.getElementById('placeholder-back-home').addEventListener('click', () => activateWorkspaceTab('home'));
  document.getElementById('sidebar-toggle').addEventListener('click', () => {
    document.getElementById('app-sidebar').classList.toggle('open');
    document.getElementById('sidebar-backdrop').classList.toggle('open');
  });
  document.getElementById('sidebar-backdrop').addEventListener('click', closeMobileSidebar);
  document.getElementById('tabs-close-others').addEventListener('click', () => {
    for (let i = openTabs.length - 1; i >= 0; i--) {
      const t = openTabs[i];
      if (!t.pinned && t.key !== activeTabKey) openTabs.splice(i, 1);
    }
    renderWorkspaceTabs();
  });
  document.querySelectorAll('.home-card[data-open-tab]').forEach(card => {
    card.addEventListener('click', () => openWorkspaceTab(card.dataset.openTab));
  });

  await initStorage();
  await ensureSeedVouchers();

  activateWorkspaceTab('home');
  animation.fromTo('.home-header-text', { opacity: 0, y: 20 }, { opacity: 1, y: 0, duration: 0.6, ease: "power2.out", delay: 0.05 });
  animation.fromTo('.home-card', { opacity: 0, y: 24 }, { opacity: 1, y: 0, duration: 0.6, stagger: 0.05, ease: "power2.out", delay: 0.1 });

  initInvoiceModule();
  initReconModule();
  initCashierModule();
  initVoucherModule();
  initAccountingModule();
  initDrilldownModule();
  initARSubledgerModule();
  initARModule();
  initTaxModule();
  initCostModule();
  await runAction(initBackup);
  if (storageMode === 'MEMORY') showToast('浏览器存储受限，当前以内存模式运行[cite: 1]');
});

/* ================= 5. 各模块逻辑 ================= */
// 5.1 开票
function initInvoiceModule() {
  let currentAttachments = [], currentStamp = null, allRecords = [];
  const itemsBody = document.getElementById('inv-items-body'), prevBody = document.getElementById('inv-prev-items-body');
  const today = new Date();
  let viewedKey = null;
  document.getElementById('inv-input-date').value = `${today.getFullYear()}-${String(today.getMonth()+1).padStart(2,'0')}-${String(today.getDate()).padStart(2,'0')}`;
  byId('inv-input-no').value = '';
  const fileInput = document.getElementById('attach-file-input'), attachList = document.getElementById('attachment-list');
  document.getElementById('attach-upload-zone').addEventListener('click', () => fileInput.click());
  fileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
      Array.from(e.target.files).forEach(file => {
        if (storageMode !== 'INDEXEDDB' && file.size > 800 * 1024) return showToast('当前存储模式不支持超大文件[cite: 1]');
        const reader = new FileReader();
        reader.onload = (ev) => { currentAttachments.push({ name: file.name, data: ev.target.result }); renderAttachments(); };
        readDraftAttachment(reader, file);
      });
      fileInput.value = '';
    }
  });
  function renderAttachments() {
    attachList.innerHTML = currentAttachments.map((file, index) => `<div class="attach-item">📄 ${escapeHtml(file.name)}<button type="button" class="btn-icon" aria-label="移除附件" onclick="removeAttachment(${index})">×</button></div>`).join('');
    currentAttachments.forEach((file, index) => {
      if (file.data || file.dataUrl) attachList.children[index].prepend(action('下载', () => downloadAttachment(file)));
    });
  }
  window.removeAttachment = function(index) { currentAttachments.splice(index, 1); renderAttachments(); };
  const stampInput = document.getElementById('stamp-file-input');
  document.getElementById('stamp-upload-zone').addEventListener('click', () => stampInput.click());
  stampInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
      const reader = new FileReader();
      reader.onload = (ev) => {
        currentStamp = ev.target.result;
        document.getElementById('stamp-preview-img').src = currentStamp;
        document.getElementById('stamp-preview-mini').style.display = 'block';
        document.getElementById('invoice-stamp-img').src = currentStamp;
        document.getElementById('invoice-stamp-img').style.display = 'block';
      };
      readDraftAttachment(reader, e.target.files[0]);
      stampInput.value = '';
    }
  });
  window.removeStamp = function() { currentStamp = null; document.getElementById('stamp-preview-mini').style.display = 'none'; document.getElementById('invoice-stamp-img').style.display = 'none'; };
  function addRow(item = null) {
    const tr = document.createElement('tr');
    tr.innerHTML = `<td><input type="text" placeholder="商品名称" class="item-name" value="${escapeHtml(item?.name || '')}"></td>
      <td><input type="number" min="1" value="${escapeHtml(item ? item.qty : 1)}" class="item-qty" style="text-align:right"></td>
      <td><input type="number" min="0" step="0.01" value="${escapeHtml(item ? item.price : '0.00')}" class="item-price" style="text-align:right"></td>
      <td class="item-subtotal" style="text-align:right; padding-right:12px; font-weight:500">0.00</td>
      <td style="text-align:center"><button type="button" class="btn-del btn-icon" aria-label="删除商品行">×</button></td>`;
    itemsBody.appendChild(tr); updatePreview();
  }
  function updatePreview() {
    document.getElementById('inv-prev-seller').textContent = document.getElementById('inv-input-seller').value || '未填写';
    document.getElementById('inv-prev-seller-taxno').textContent = document.getElementById('inv-input-seller-taxno').value || '未填写';
    document.getElementById('inv-prev-buyer').textContent = document.getElementById('inv-input-buyer').value || '未填写';
    document.getElementById('inv-prev-taxno').textContent = document.getElementById('inv-input-taxno').value || '未填写';
    document.getElementById('inv-prev-no').textContent = document.getElementById('inv-input-no').value;
    document.getElementById('inv-prev-date').textContent = document.getElementById('inv-input-date').value;
    document.getElementById('inv-prev-payment').textContent = document.getElementById('inv-input-payment').value;
    const rows = itemsBody.querySelectorAll('tr'); prevBody.innerHTML = ''; let total = 0;
    rows.forEach(row => {
      const name = row.querySelector('.item-name')?.value || '***';
      const qty = parseFloat(row.querySelector('.item-qty')?.value) || 0;
      const price = parseFloat(row.querySelector('.item-price')?.value) || 0;
      let sub = qty * price; if (isNaN(sub)) sub = 0; total += sub;
      row.querySelector('.item-subtotal').textContent = sub.toFixed(2);
      const tr = document.createElement('tr'); tr.innerHTML = `<td>${escapeHtml(name)}</td><td>${qty}</td><td>${price.toFixed(2)}</td><td>${sub.toFixed(2)}</td>`; prevBody.appendChild(tr);
    });
    const safe = isNaN(total) ? 0 : total;
    document.getElementById('inv-prev-total').textContent = safe.toFixed(2);
    document.getElementById('inv-prev-total-small').textContent = safe.toFixed(2);
    document.getElementById('inv-prev-upper').textContent = digitUppercase(safe);
  }
  document.querySelectorAll('.inv-sync').forEach(el => { el.addEventListener('input', updatePreview); el.addEventListener('change', updatePreview); });
  document.getElementById('inv-btn-add').addEventListener('click', () => addRow());
  itemsBody.addEventListener('click', e => { if (e.target.classList.contains('btn-del')) { e.target.closest('tr').remove(); updatePreview(); } });
  itemsBody.addEventListener('input', () => updatePreview());
  function resetInvoice() {
    viewedKey = null;
    for (const id of ['buyer', 'taxno', 'no']) byId('inv-input-' + id).value = '';
    byId('inv-source').value = ''; byId('inv-evidence').value = '';
    byId('inv-status').selectedIndex = 0;
    byId('inv-input-date').value = localDay();
    itemsBody.replaceChildren(); currentAttachments = []; renderAttachments(); removeStamp(); addRow();
  }
  byId('inv-btn-submit').after(action('新建记录', resetInvoice));
  onAction('inv-btn-submit', async () => {
    const invoiceNumber = val('inv-input-no'), status = val('inv-status');
    const seller = val('inv-input-seller'), buyer = val('inv-input-buyer'), date = val('inv-input-date');
    const source = val('inv-source'), evidence = val('inv-evidence');
    const sellerTax = val('inv-input-seller-taxno'), buyerTax = val('inv-input-taxno');
    if (!seller || !buyer || !date || !source) throw Error('请填写购销双方、日期及来源编号[cite: 1]');
    if (status === '已开票' && (!invoiceNumber || !sellerTax || !buyerTax || (!evidence && !currentAttachments.length))) {
      throw Error('已开票须填写实际发票号、双方税号，并附发票或回执依据[cite: 1]');
    }
    if (status === '办理失败' && !evidence) throw Error('请填写失败原因[cite: 1]');
    const items = [...itemsBody.rows].map(row => ({
      name: row.querySelector('.item-name').value.trim(),
      qty: Number(row.querySelector('.item-qty').value),
      price: Number(row.querySelector('.item-price').value)
    }));
    if (!items.length || items.some(i => !i.name || !Number.isFinite(i.qty) || i.qty <= 0 || !i.price || i.price < 0 || !Number.isFinite(i.price))) throw Error('请填写完整商品、数量和单价[cite: 1]');
    const total = items.reduce((sum, i) => sum + Math.round(i.qty * i.price * 100), 0) / 100;
    if (!Number.isSafeInteger(Math.round(total * 100)) || total <= 0) throw Error('金额无效[cite: 1]');
    const all = await getAllRecordsFromDB();
    if (invoiceNumber && all.some(i => i.invoiceNumber === invoiceNumber && i.invoiceNo !== viewedKey)) throw Error('该发票号码已登记[cite: 1]');
    const record = { invoiceNo: viewedKey || recordId(), invoiceNumber, date, seller, buyer, sellerTax, buyerTax,
      total, items, status, source, evidence, payment: val('inv-input-payment'), attachments: [...currentAttachments], stamp: currentStamp };
    await saveRecords(STORE_NAME, [record], false, 'invoice');
    await loadAndRenderLedger(); resetInvoice(); showToast('发票记录已保存[cite: 1]');
  });
  async function loadAndRenderLedger() { allRecords = await getAllRecordsFromDB(); filterLedger(); }
  function filtered() {
    return allRecords.filter(i => (i.invoiceNumber || '').includes(val('f-invoice-no')) &&
      (i.buyer.includes(val('f-buyer')) || (i.source || '').includes(val('f-buyer'))) &&
      (!val('f-payment') || i.payment === val('f-payment')) &&
      (!val('f-date-start') || i.date >= val('f-date-start')) && (!val('f-date-end') || i.date <= val('f-date-end')));
  }
  window.filterLedger = function() {
    const list = filtered(), tb = byId('inv-ledger-body'); tb.replaceChildren();
    for (const i of list) {
      const row = tableRow(tb, ['', i.invoiceNumber || '—', i.date, i.buyer, money(i.total), i.source, i.status]);
      row.cells[0].append(action('查看', () => viewInvoice(i.invoiceNo)));
    }
    byId('filtered-total-count').textContent = '共 ' + list.length + ' 张';
    byId('filtered-total-amount').textContent = money(list.reduce((sum, i) => sum + i.total, 0));
  };
  window.viewInvoice = function(no) {
    const r = allRecords.find(x => x.invoiceNo === no); if (!r) return;
    const fields = {'input-buyer':r.buyer, 'input-no':r.invoiceNumber, 'input-seller':r.seller, 'input-taxno':r.buyerTax,
      'input-seller-taxno':r.sellerTax, 'input-date':r.date, 'input-payment':r.payment, source:r.source, status:r.status, evidence:r.evidence};
    Object.entries(fields).forEach(([key, value]) => byId('inv-' + key).value = value || '');
    viewedKey = no;
    currentAttachments = [...(r.attachments || [])]; renderAttachments(); removeStamp();
    if (r.stamp) {
      currentStamp = r.stamp;
      for (const id of ['stamp-preview-img', 'invoice-stamp-img']) { byId(id).src = r.stamp; byId(id).style.display = ''; }
      byId('stamp-preview-mini').style.display = '';
    }
    itemsBody.replaceChildren(); r.items.forEach(addRow); updatePreview();
    byId('inv-input-no').focus();
  };
  for (const id of ['f-invoice-no','f-buyer','f-payment','f-date-start','f-date-end']) byId(id).addEventListener('input', filterLedger);
  byId('inv-filter-reset').addEventListener('click', () => {
    for (const id of ['f-invoice-no','f-buyer','f-payment','f-date-start','f-date-end']) byId(id).value = '';
    filterLedger();
  });
  byId('f-buyer').placeholder = '购买方 / 合同编号';
  onAction('inv-btn-clear', async () => {
    if (await uiConfirm('清空本账套全部发票记录？')) { await clearAllDB(); await loadAndRenderLedger(); resetInvoice(); }
  });
  onAction('inv-btn-export', () => exportRows('发票台账', ['发票号码','日期','购方','合同编号','金额','状态'],
    filtered().map(i => [i.invoiceNumber,i.date,i.buyer,i.source,i.total,i.status])));
  registerDraft('invoice', ['inv-input-seller', 'inv-input-seller-taxno', 'inv-input-buyer',
    'inv-input-taxno', 'inv-input-no', 'inv-input-date', 'inv-input-payment', 'inv-source', 'inv-status', 'inv-evidence'],
    () => ({items: fieldRows(itemsBody, '.item-', ['name','qty','price']), attachments: currentAttachments, stamp: currentStamp, viewedKey}),
    draft => {
      viewedKey = draft.viewedKey;
      currentAttachments = draft.attachments || []; renderAttachments();
      removeStamp();
      if (draft.stamp) {
        currentStamp = draft.stamp;
        for (const id of ['stamp-preview-img','invoice-stamp-img']) byId(id).src = currentStamp;
        for (const id of ['stamp-preview-mini','invoice-stamp-img']) byId(id).style.display = 'block';
      }
      itemsBody.replaceChildren(); draft.items.forEach(addRow); updatePreview();
    });
  resetInvoice(); loadAndRenderLedger();
}

// 5.3 出纳
function initCashierModule() {
  let cashierRecords = [], cashierCurrentType = 'RECEIPT';
  function refreshAccountDatalist() { document.getElementById('cash-account-options').innerHTML = ACCOUNTS.map(a => `<option value="${escapeHtml(a.name)}"></option>`).join(''); }
  function refreshJournalAccountOptions() {
    const jAcc = document.getElementById('cash-journal-account'); const prev = jAcc.value; const set = new Set();
    ACCOUNTS.forEach(a => set.add(a.name)); cashierRecords.forEach(r => { if (r.account) set.add(getAccountName(r.account)); });
    const names = Array.from(set).filter(Boolean).sort();
    jAcc.innerHTML = '<option value="">全部账户</option>' + names.map(n => `<option value="${escapeHtml(n)}">${escapeHtml(n)}</option>`).join('');
    if (prev && names.includes(prev)) jAcc.value = prev;
  }
  refreshAccountDatalist(); refreshJournalAccountOptions();
  if (ACCOUNTS[1]) document.getElementById('cash-account').value = ACCOUNTS[1].name;

  window.toggleAccountManager = function() {
    const el = document.getElementById('account-manager'); const open = el.style.display === 'none';
    el.style.display = open ? 'block' : 'none'; if (open) renderAccountManager();
  };
  function renderAccountManager() {
    const tbody = document.getElementById('account-manage-body');
    tbody.innerHTML = ACCOUNTS.map(a => `<tr><td><input type="text" value="${escapeHtml(a.name)}" onchange="updateAccountName('${a.id}', this.value)"></td><td><select onchange="updateAccountKind('${a.id}', this.value)"><option value="BANK" ${a.kind === 'BANK' ? 'selected' : ''}>银行</option><option value="CASH" ${a.kind === 'CASH' ? 'selected' : ''}>现金</option></select></td><td><button type="button" class="btn-mini btn-danger" onclick="removeAccount('${a.id}')">删除</button></td></tr>`).join('');
  }
  window.updateAccountName = function(id, name) { const a = ACCOUNTS.find(x => x.id === id); if (a) { a.name = name.trim() || a.name; saveCashierAccounts(); refreshAccountDatalist(); refreshJournalAccountOptions(); } };
  window.updateAccountKind = function(id, kind) { const a = ACCOUNTS.find(x => x.id === id); if (a) { a.kind = kind; saveCashierAccounts(); } };
  window.removeAccount = function(id) { ACCOUNTS = ACCOUNTS.filter(x => x.id !== id); saveCashierAccounts(); renderAccountManager(); refreshAccountDatalist(); refreshJournalAccountOptions(); };
  window.addNewAccount = function() {
    const name = document.getElementById('new-account-name').value.trim(); if (!name) return showToast('请填写账户名称[cite: 1]');
    ACCOUNTS.push({ id: 'ACC' + Date.now(), name, kind: document.getElementById('new-account-kind').value });
    saveCashierAccounts(); document.getElementById('new-account-name').value = ''; renderAccountManager(); refreshAccountDatalist(); refreshJournalAccountOptions(); showToast('已添加[cite: 1]');
  };

  window.setCashierType = function(type) {
    cashierCurrentType = type; const isT = type === 'TRANSFER';
    document.querySelectorAll('#module-cashier .quick-card').forEach(c => c.classList.toggle('active', c.dataset.type === type));
    document.querySelectorAll('#module-cashier .type-btn').forEach(b => b.classList.toggle('active', b.dataset.type === type));
    document.getElementById('cash-field-in-account').style.display = isT ? 'block' : 'none';
    document.getElementById('cash-field-counterparty').style.display = isT ? 'none' : 'block';
    document.getElementById('cash-label-account').textContent = type === 'RECEIPT' ? '收款账户' : (type === 'PAYMENT' ? '付款账户' : '转出账户');
    document.getElementById('cash-btn-submit').textContent = type === 'RECEIPT' ? '登记收款' : (type === 'PAYMENT' ? '登记付款' : '登记转账');
  };

  let attachment = null;
  const restoreCashAttachment = filePicker('cash-receipt-upload-zone', 'cash-receipt-file-input', 'cash-receipt-preview', file => attachment = file);
  const statusNames = {PENDING:'待办理', DONE:'已办理', POSTED:'已入账'};
  const types = {RECEIPT:'收款', PAYMENT:'付款', TRANSFER_OUT:'转出', TRANSFER_IN:'转入'};
  const signed = r => (['RECEIPT','TRANSFER_IN'].includes(r.type) ? 1 : -1) * Math.round(r.amount * 100);
  const cashFilterIds = ['cash-filter-type','cash-filter-status','cash-filter-start','cash-filter-end'];
  window.resetCashFilters = () => { for (const id of cashFilterIds) byId(id).value = ''; renderCash(); };
  for (const id of cashFilterIds) byId(id).addEventListener('change', renderCash);
  byId('cash-journal-account').addEventListener('change', renderCash);
  tabs('#module-cashier .cash-tab', 'cash-tab-', '#module-cashier .cash-tab-panel');
  function renderCash() {
    const [type, status, start, end] = cashFilterIds.map(val);
    const register = byId('cash-register-body'), notes = byId('cash-notes-body'),
      daily = byId('cash-daily-list'), journal = byId('cash-journal-body');
    for (const node of [register, notes, daily, journal]) node.replaceChildren();
    const groups = new Map(), selected = val('cash-journal-account');
    for (const r of [...cashierRecords].sort((a,b) => a.date.localeCompare(b.date))) {
      const amount = money(signed(r) / 100), name = getAccountName(r.account);
      if ((!type || r.type === type || (type === 'TRANSFER' && r.transferId)) &&
        (!status || r.status === status) && (!start || r.date >= start) && (!end || r.date <= end)) {
        const row = tableRow(register, [r.date, types[r.type], amount, name, r.counterparty, r.sourceDoc, r.receiptNo, statusNames[r.status], '']);
        if (r.status !== 'POSTED') row.lastChild.append(action(r.status === 'PENDING' ? '确认办理' : '确认入账', async () => {
          const current = await getAllFromStore(STORE_CASHIER);
          const next = current.filter(item => r.transferId ? item.transferId === r.transferId : item.id === r.id);
          if (next.some(item => item.status !== r.status)) throw Error('状态已变化，请刷新[cite: 1]');
          if (r.status === 'PENDING' && !r.receiptNo && !r.attachment && !r.note) throw Error('缺少回单或票据依据，请补交资料[cite: 1]');
          await saveRecords(STORE_CASHIER, next.map(item => ({...item, status:r.status === 'PENDING' ? 'DONE' : 'POSTED'})));
          await loadCashier();
        }));
        if (r.status === 'PENDING') row.lastChild.append(action('补录回单号', async () => {
          const number = (await uiPrompt('回单 / 收据编号', r.receiptNo || ''))?.trim(); if (!number) return;
          const current = await getAllFromStore(STORE_CASHIER);
          const same = item => r.transferId ? item.transferId === r.transferId : item.id === r.id;
          if (current.some(item => !same(item) && item.receiptNo === number)) throw Error('该回单已登记[cite: 1]');
          const records = current.filter(same);
          if (records.some(item => item.status !== 'PENDING')) throw Error('状态已变化，请刷新[cite: 1]');
          await saveRecords(STORE_CASHIER, records.map(item => ({...item, receiptNo:number}))); await loadCashier();
        }));
      }
      if ((r.receiptNo || r.attachment || r.note) && r.type !== 'TRANSFER_IN') {
        const row = tableRow(notes, [r.receiptNo, r.receiptType, r.date, types[r.type], money(r.amount), r.note, '', statusNames[r.status]]);
        if (r.attachment) row.cells[6].append(action('下载', () => downloadAttachment(r.attachment)));
      }
      if (r.status === 'PENDING') continue;
      const group = groups.get(r.date) || [];
      group.push([name, types[r.type], amount, r.receiptNo || '']); groups.set(r.date, group);
      if (!selected || name === selected) {
        const line = document.createElement('div');
        line.style.padding = '6px 0';
        line.style.borderBottom = '1px dashed var(--border)';
        line.textContent = [r.date, name, types[r.type], amount].join(' ');
        journal.append(line);
      }
    }
    for (const [date, rows] of groups) {
      const details = document.createElement('details'), title = document.createElement('summary');
      details.style.marginBottom = '10px';
      title.style.cursor = 'pointer';
      title.textContent = date + ' · ' + rows.length + ' 笔';
      details.append(title);
      rows.forEach(row => { const line = document.createElement('div'); line.style.padding = '4px 12px'; line.textContent = row.join(' '); details.append(line); });
      daily.append(details);
    }
  }
  async function loadCashier() {
    cashierRecords = await getAllFromStore(STORE_CASHIER);
    refreshJournalAccountOptions(); renderCash();
  }

  byId('cash-btn-submit').before(makeField('办理状态','cash-entry-status','select',[['PENDING','待办理'],['DONE','已办理']]));
  onAction('cash-btn-submit', async () => {
    const account = val('cash-account'), date = val('cash-date'), amt = cents(val('cash-amount')) / 100;
    const sourceDoc = val('cash-source-doc'), receiptNo = val('cash-receipt-no'), note = val('cash-note-info');
    const status = val('cash-entry-status'), type = cashierCurrentType, incoming = val('cash-in-account');
    if (!account || !date || !sourceDoc || amt <= 0) throw Error('请填写账户、日期、来源单号和有效金额[cite: 1]');
    if (type === 'TRANSFER' && (!incoming || incoming === account)) throw Error('请填写不同的转出和转入账户[cite: 1]');
    if (status === 'DONE' && !receiptNo && !attachment && !note) throw Error('已办理需要回单、附件或票据信息[cite: 1]');
    const all = await getAllFromStore(STORE_CASHIER);
    if (receiptNo && all.some(r => r.receiptNo === receiptNo)) throw Error('该回单已登记[cite: 1]');
    const details = { status,date,amount:amt,sourceDoc,receiptNo,note,attachment,sourceType:val('cash-source-type'),
      receiptType:val('cash-receipt-type'), memo:val('cash-memo'), counterpartyAccount:val('cash-counterparty-account') };
    const id = recordId();
    const records = type === 'TRANSFER'
      ? [{...details,id:id+'O',type:'TRANSFER_OUT',account,transferId:id},{...details,id:id+'I',type:'TRANSFER_IN',account:incoming,transferId:id}]
      : [{...details,id,type,account,counterparty:val('cash-counterparty')}];
    await saveRecords(STORE_CASHIER, records, true, 'cashier');
    for (const field of ['amount','source-doc','receipt-no','note-info','memo','counterparty','counterparty-account','in-account']) byId('cash-' + field).value = '';
    attachment = null; byId('cash-receipt-preview').replaceChildren();
    await loadCashier(); showToast('出纳登记已保存[cite: 1]');
  });

  const now = new Date(); document.getElementById('cash-date').value = `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,'0')}-${String(now.getDate()).padStart(2,'0')}`;
  registerDraft('cashier', ['date','amount','account','in-account','counterparty','counterparty-account',
    'source-doc','source-type','receipt-type','receipt-no','note-info','memo','entry-status'].map(id => 'cash-' + id),
    () => ({type:cashierCurrentType, attachment}),
    draft => { setCashierType(draft.type); restoreCashAttachment(draft.attachment); });
  setCashierType('RECEIPT'); loadCashier();
}

// 5.4 制证（支持辅助核算：银行账户 / 往来单位）
function initVoucherModule() {
  const V_BIZ = {
    BANK: { label: '银行往来', template: [{ account:'1002', aux:'工商银行基本户', summary:'银行往来', side:'D' },{ account:'1122', aux:'华远科技有限公司', summary:'银行往来', side:'C' }] },
    PAYROLL: { label: '工资计提', template: [{ account:'6602', aux:'', summary:'工资计提', side:'D' },{ account:'2211', aux:'', summary:'工资计提', side:'C' }] },
    PURCHASE: { label: '采购成本', template: [{ account:'1401', aux:'', summary:'采购成本', side:'D' },{ account:'2202', aux:'江南材料供应有限公司', summary:'采购成本', side:'C' }] },
    EXPENSE: { label: '费用报销', template: [{ account: '6602', aux:'', summary: '报销费用', side: 'D' }, { account: '1002', aux:'工商银行基本户', summary: '银行付款', side: 'C' }] },
    SALES: { label: '销售收入', template: [{ account: '1122', aux:'蓝海贸易（上海）有限公司', summary: '确认收入', side: 'D' }, { account: '6001', aux:'', summary: '确认收入', side: 'C' }] }
  };
  let vRecords = [], vType = 'EXPENSE';
  function renderTypes() {
    document.getElementById('voucher-types').innerHTML = Object.keys(V_BIZ).map(k => `<div class="voucher-type-card ${k === vType ? 'active' : ''}" onclick="setVType('${k}')"><div class="voucher-type-icon" style="background:#FFF9F6;color:var(--primary);">📄</div><div><div class="voucher-type-title">${V_BIZ[k].label}</div></div></div>`).join('');
  }
  window.setVType = function(k) { vType = k; renderTypes(); resetVEntries(); };
  function accOpts(sel) { return VOUCHER_ACCOUNTS.map(a => `<option value="${a.code}" ${sel === a.code ? 'selected' : ''}>${a.code} ${a.name}</option>`).join(''); }
  function renderAuxSelectEl(selEl, code, currentVal = '') {
    const opts = getAuxOptionsForAccount(code);
    if (!opts.length) {
      selEl.innerHTML = '<option value="">无辅助</option>';
      selEl.disabled = true;
    } else {
      selEl.innerHTML = opts.map(o => `<option value="${escapeHtml(o)}" ${o === currentVal ? 'selected' : ''}>${escapeHtml(o)}</option>`).join('');
      selEl.disabled = false;
    }
  }
  function addVEntry(data) {
    const tr = document.createElement('tr');
    const code = data?.account || '1002';
    tr.innerHTML = `<td><select class="entry-account">${accOpts(code)}</select></td>
      <td><select class="entry-aux"></select></td>
      <td><input type="text" class="entry-summary" value="${escapeHtml(data?.summary || '')}"></td>
      <td><input type="number" class="entry-debit num" value="${escapeHtml(data?.debit ?? '')}"></td>
      <td><input type="number" class="entry-credit num" value="${escapeHtml(data?.credit ?? '')}"></td>
      <td><button type="button" class="btn-remove-row btn-icon" aria-label="删除分录行" onclick="this.closest('tr').remove(); recalcV();">×</button></td>`;
    document.getElementById('voucher-entries-body').appendChild(tr);
    const accSel = tr.querySelector('.entry-account');
    const auxSel = tr.querySelector('.entry-aux');
    renderAuxSelectEl(auxSel, code, data?.aux || '');
    accSel.addEventListener('change', () => renderAuxSelectEl(auxSel, accSel.value, ''));
    tr.querySelectorAll('input').forEach(i => i.addEventListener('input', recalcV)); recalcV();
  }
  function resetVEntries() {
    document.getElementById('voucher-entries-body').innerHTML = '';
    V_BIZ[vType].template.forEach(t => addVEntry({ account: t.account, aux: t.aux || '', summary: t.summary, [t.side === 'D' ? 'debit' : 'credit']: '' }));
  }
  window.recalcV = function() {
    let d = 0, c = 0;
    document.querySelectorAll('#voucher-entries-body tr').forEach(tr => { d += parseFloat(tr.querySelector('.entry-debit').value) || 0; c += parseFloat(tr.querySelector('.entry-credit').value) || 0; });
    document.getElementById('voucher-debit-total').textContent = d.toFixed(2); document.getElementById('voucher-credit-total').textContent = c.toFixed(2);
    const status = document.getElementById('voucher-balance-status');
    if (Math.abs(d - c) < 0.005 && d > 0) { status.className = 'voucher-balance-status ok'; status.textContent = `借贷平衡 ¥${d.toFixed(2)}`; }
    else { status.className = 'voucher-balance-status bad'; status.textContent = `借贷不平 差额 ¥${(d - c).toFixed(2)}`; }
  };
  let attachment = null;
  const restoreVoucherAttachment = filePicker('voucher-attach-zone', 'voucher-attach-input', 'voucher-attach-list', file => {
    attachment = file; byId('voucher-attach-count').textContent = file ? '1' : '0';
  });
  byId('voucher-source-type').replaceChildren(...['合同','发票','银行回单','报销单','工资表','其他'].map(name => new Option(name,name)));
  function setVoucherPeriod() { byId('voucher-period').value = val('voucher-date').slice(0,7); }
  byId('voucher-date').addEventListener('change', setVoucherPeriod);
  window.resetVoucherForm = () => {
    for (const id of ['voucher-source-no','voucher-counterparty','voucher-memo']) byId(id).value = '';
    attachment = null; byId('voucher-attach-list').replaceChildren(); byId('voucher-attach-count').textContent = '0';
    resetVEntries();
  };
  window.addVoucherEntry = () => addVEntry();
  window.addEntryRow = () => addVEntry();
  async function refreshClosed() {
    const records = await getAllFromStore(STORE_WORK);
    byId('voucher-closed-periods').textContent = records.filter(r => r.id.startsWith('closed:')).map(r => r.period).sort().join('、') || '无';
  }
  window.closePeriod = () => runAction(async () => {
    const period = val('voucher-period-input'); if (!period) throw Error('请选择期间[cite: 1]');
    if (!(await uiConfirm('锁定该期间的制证、审核和记账？'))) return;
    await lockPeriod(period, '手动锁定'); await refreshClosed();
  });
  onAction('voucher-btn-submit', async () => {
    const date = val('voucher-date'), period = date.slice(0,7), source = val('voucher-source-no');
    if (!date || (!source && !attachment)) throw Error('请填写日期并关联来源单号或附件[cite: 1]');
    await ensureOpen(period);
    const all = await getAllFromStore(STORE_VOUCHER);
    const duplicates = source && all.some(v => v.source === source && v.sourceType === val('voucher-source-type'));
    if (duplicates && !(await uiConfirm('已存在相同来源的凭证，确认属于一单多凭证？'))) return;
    const entries = [...byId('voucher-entries-body').rows].map(tr => {
      const read = cls => tr.querySelector('.entry-' + cls).value;
      const auxEl = tr.querySelector('.entry-aux');
      const aux = auxEl && !auxEl.disabled ? auxEl.value : '';
      const debit = cents(read('debit') || '0'), credit = cents(read('credit') || '0');
      if (debit < 0 || credit < 0 || Boolean(debit) === Boolean(credit) || !read('summary').trim()) throw Error('每行须填写摘要，并只在借方或贷方录入正金额[cite: 1]');
      return {account:read('account'), aux, summary:read('summary'), debit:debit/100, credit:credit/100};
    });
    const [d,c] = ['debit','credit'].map(key => entries.reduce((sum,row) => sum + Math.round(row[key]*100),0));
    if (!Number.isSafeInteger(d) || d <= 0 || d !== c) throw Error('借贷金额不平或无有效分录[cite: 1]');
    const voucherNo = '记-'+date.replace(/-/g,'')+'-'+String(all.filter(v=>v.date===date).length+1).padStart(3,'0');
    await saveRecords(STORE_VOUCHER, [{id:recordId(), voucherNo,date,period,bizLabel:V_BIZ[vType].label,entries,
      totalDebit:d/100,totalCredit:c/100,status:'PENDING',source,sourceType:val('voucher-source-type'),
      counterparty:val('voucher-counterparty'),memo:val('voucher-memo'),attachment}], true, 'voucher');
    byId('voucher-no-badge').textContent = voucherNo;
    resetVoucherForm(); await loadVList(); showToast('凭证已提交审核（审核记账后自动更新科目余额表）');
  });
  async function loadVList() {
    vRecords = await getAllFromStore(STORE_VOUCHER);
    const tb = byId('voucher-list-body'); tb.replaceChildren();
    byId('voucher-list-count').textContent = vRecords.length;
    for (const v of [...vRecords].reverse()) {
      const auxHint = (v.entries || []).map(e => e.aux).filter(Boolean).join(' / ');
      const summaryText = (v.entries[0]?.summary || '') + (auxHint ? ` [${auxHint}]` : '');
      const row = tableRow(tb,[v.voucherNo,v.date,v.bizLabel,summaryText,money(v.totalDebit),money(v.totalCredit),
        v.source, {PENDING:'待审核',APPROVED:'已审核',POSTED:'已记账'}[v.status],'']);
      if (v.attachment) row.cells[6].append(action('附件', () => downloadAttachment(v.attachment)));
      if (v.status === 'PENDING') row.lastChild.append(action('审核通过', async () => {
        await changeVoucherStatus([v.id], 'PENDING', 'APPROVED'); await loadVList(); showToast('已审核，可前往「会计记账」入账[cite: 1]');
      }));
    }
  }
  window.refreshVouchers = loadVList;
  const entriesBody = byId('voucher-entries-body');
  registerDraft('voucher', ['voucher-date','voucher-source-type','voucher-source-no','voucher-counterparty','voucher-memo'],
    () => ({type:vType, entries:fieldRows(entriesBody,'.entry-',['account','aux','summary','debit','credit']), attachment}),
    draft => {
      vType = draft.type; renderTypes(); entriesBody.replaceChildren();
      draft.entries.forEach(addVEntry); recalcV(); setVoucherPeriod(); restoreVoucherAttachment(draft.attachment);
    });
  renderTypes(); resetVEntries(); loadVList(); refreshClosed();
  byId('voucher-date').value = localDay(); setVoucherPeriod();
}

// 5.5 会计记账
function initAccountingModule() {
  let accV = [], accPeriod = '';
  async function loadAcc() {
    accV = await getAllFromStore(STORE_VOUCHER);
    const set = new Set(['2026-01', localDay().slice(0, 7), ...(await getAllFromStore(STORE_WORK)).map(r => r.period).filter(Boolean)]);
    accV.forEach(v => set.add(v.period));
    const periods = Array.from(set).sort().reverse();
    if (!accPeriod || !set.has(accPeriod)) accPeriod = periods[0];
    document.getElementById('acc-period-select').innerHTML = periods.map(p => `<option value="${p}">${p}</option>`).join('');
    document.getElementById('acc-period-select').value = accPeriod;
    renderPending(); updateBadge(); await renderBooks(accV, accPeriod);
  }
  window.onAccPeriodChange = function(p) { accPeriod = p; renderPending(); updateBadge(); runAction(() => renderBooks(accV, p)); };
  function renderPending() {
    const list = accV.filter(v => v.status === 'APPROVED' && v.period === accPeriod);
    const tb = document.getElementById('acc-pending-body');
    if (!list.length) { tb.innerHTML = '<tr><td colspan="8" style="text-align:center;">无待记账凭证[cite: 1]</td></tr>'; return; }
    tb.replaceChildren();
    for (const v of list) {
      const row = tableRow(tb,[v.voucherNo,v.date,v.bizLabel,v.entries[0]?.summary,money(v.totalDebit),money(v.totalCredit),'已审核','']);
      row.lastChild.append(action('记账', () => postV(v.id)));
    }
  }
  window.postV = id => runAction(async () => {
    await changeVoucherStatus([id], 'APPROVED', 'POSTED');
    await loadAcc(); await refreshVouchers();
    if (window.renderLiveBalance) await renderLiveBalance();
    if (window.renderARSubledgerModule) await renderARSubledgerModule();
    showToast('已记账，科目余额表与三栏账已同步');
  });
  window.postAllPending = () => runAction(async () => {
    const ids = accV.filter(v => v.status === 'APPROVED' && v.period === accPeriod).map(v => v.id);
    if (!ids.length) return showToast('无待记账凭证[cite: 1]');
    await changeVoucherStatus(ids, 'APPROVED', 'POSTED');
    await loadAcc(); await refreshVouchers();
    if (window.renderLiveBalance) await renderLiveBalance();
    if (window.renderARSubledgerModule) await renderARSubledgerModule();
    showToast('全部记账完成');
  });
  tabs('#acc-tabs .acc-tab','acc-panel-','#module-accounting .acc-tab-panel');
  wireBookEditors(() => accPeriod, loadAcc);
  function updateBadge() { document.getElementById('acc-badge-pending').textContent = accV.filter(v => v.status === 'APPROVED' && v.period === accPeriod).length; }
  window.refreshAccounting = () => runAction(loadAcc);
  loadAcc();
}

// 5.6 公共工具与账簿
const byId = id => document.getElementById(id);
const val = id => byId(id).value.trim();
const recordId = () => crypto.randomUUID?.() || Array.from(crypto.getRandomValues(new Uint32Array(4)), n => n.toString(16)).join('-');
const money = amt => '¥' + Number(amt).toFixed(2);
function localDay() {
  const now = new Date();
  return [now.getFullYear(), now.getMonth()+1, now.getDate()].map(n => String(n).padStart(2,'0')).join('-');
}
function cents(value) {
  const text = String(value).trim();
  const number = Math.round(Number(text)*100);
  if (!/^-?\d+(?:\.\d{1,2})?$/.test(text) || !Number.isSafeInteger(number)) throw Error('请输入有效金额，最多保留两位小数[cite: 1]');
  return number;
}
async function runAction(handler, button) {
  if (button) { if (button.disabled) return; button.disabled = true; }
  try { await handler(); }
  catch (error) { showToast(error.name === 'QuotaExceededError' ? '存储空间不足，资料未保存' : (error.message || '操作失败')); }
  finally { if (button) button.disabled = false; }
}
function onAction(id, handler) { const button = byId(id); button.addEventListener('click', () => runAction(handler, button)); }
function action(label, handler, buttonClass = 'btn-mini') {
  const button = document.createElement('button');
  button.type = 'button'; button.className = buttonClass; button.textContent = label;
  button.addEventListener('click', () => runAction(handler, button));
  return button;
}
function tableRow(body, values) {
  const row = body.insertRow();
  values.forEach(value => row.insertCell().textContent = value ?? '—');
  return row;
}
function makeField(label, id, type = 'text', options = []) {
  const group = document.createElement('div'); group.className = 'form-group';
  const title = document.createElement('label'); title.htmlFor = id; title.textContent = label;
  const input = document.createElement(type === 'select' || type === 'textarea' ? type : 'input');
  input.id = id;
  if (type === 'select') input.append(...options.map(([value,name]) => new Option(name,value)));
  else if (type !== 'textarea') { input.type = type; if (type === 'number') input.step = '0.01'; }
  group.append(title,input); return group;
}
function filePicker(zoneId, inputId, listId, update) {
  const input = byId(inputId), list = byId(listId);
  function showFile(file) {
    update(file);
    list.replaceChildren();
    if (file) list.append(document.createTextNode(file.name + ' '), action('移除', () => showFile(null)));
  }
  byId(zoneId).addEventListener('click', event => { if (event.target !== input) input.click(); });
  input.addEventListener('change', async () => {
    const file = input.files[0]; input.value = '';
    if (!file) return;
    await runAction(async () => {
      const data = await new Promise((resolve,reject) => {
        const reader = new FileReader(); reader.onload = () => resolve(reader.result);
        reader.onerror = () => reject(Error('附件读取失败')); readDraftAttachment(reader, file);
      });
      showFile({name:file.name,data});
    });
  });
  return showFile;
}
function downloadAttachment(file) {
  const data = file.data || file.dataUrl;
  if (!/^data:[^,]*;base64,/i.test(data)) throw Error('附件格式无效[cite: 1]');
  const link = document.createElement('a'); link.href = data; link.download = file.name; link.click();
}
function downloadFile(name, content, mime) {
  const url = URL.createObjectURL(new Blob([content], {type:mime}));
  const link = document.createElement('a'); link.href = url; link.download = name; link.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
function exportRows(name, headers, rows) {
  if (window.XLSX) {
    const book = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(book,XLSX.utils.aoa_to_sheet([headers,...rows]),'数据');
    XLSX.writeFile(book,name+'.xlsx'); return;
  }
  const cell = value => '"' + String(typeof value === 'string' && /^[=+\-@\t\r]/.test(value) ? "'"+value : value ?? '').replaceAll('"','""') + '"';
  downloadFile(name+'.csv', '\ufeff' + [headers,...rows].map(row => row.map(cell).join(',')).join('\r\n'), 'text/csv;charset=utf-8');
}
function tabs(selector, prefix, panels) {
  const items = [...document.querySelectorAll(selector)];
  for (const item of items) item.addEventListener('click', () => {
    items.forEach(tab => tab.classList.toggle('active',tab === item));
    document.querySelectorAll(panels).forEach(panel => panel.classList.toggle('active',panel.id === prefix + item.dataset.tab));
  });
}
async function work(id) { return (await getAllFromStore(STORE_WORK)).find(item => item.id === id); }
async function ensureOpen(period) { if (!period || await work('closed:'+period)) throw Error('请选择有效的开放期间[cite: 1]'); }
async function lockPeriod(period, reason, archive = null) {
  await ensureOpen(period);
  const all = await getAllFromStore(STORE_VOUCHER);
  if (all.some(v => v.period === period && v.status !== 'POSTED')) throw Error('期间仍有未记账凭证[cite: 1]');
  await saveRecords(STORE_WORK, [{id:'closed:'+period,period,reason,archive,createdAt:new Date().toISOString()}],true);
}
async function changeVoucherStatus(ids, from, to) {
  const all = await getAllFromStore(STORE_VOUCHER), records = ids.map(id => all.find(v => v.id === id));
  if (records.some(v => !v || v.status !== from)) throw Error('凭证状态已变化，请刷新[cite: 1]');
  for (const period of new Set(records.map(v => v.period))) await ensureOpen(period);
  await saveRecords(STORE_VOUCHER, records.map(v => ({...v,status:to})));
}

let bookReport = 'bs', currentBook = [];
function bookRows(records, period, opening = {}) {
  return VOUCHER_ACCOUNTS.map(account => {
    const entries = records.filter(v => v.period === period && v.status === 'POSTED')
      .sort((a,b) => a.date.localeCompare(b.date) || a.voucherNo.localeCompare(b.voucherNo))
      .flatMap(v => v.entries.filter(e => e.account === account.code).map(e => ({...e,date:v.date,no:v.voucherNo,source:v.source,biz:v.bizLabel})));
    const debit = entries.reduce((sum,e) => sum+Math.round(e.debit*100),0);
    const credit = entries.reduce((sum,e) => sum+Math.round(e.credit*100),0);
    const defaultInitCents = (account.dir === '借' ? 1 : -1) * Math.round((account.opening || 0) * 100);
    const initial = opening[account.code] !== undefined ? opening[account.code] : defaultInitCents;
    return {...account,entries,debit,credit,initial,end:initial+debit-credit};
  });
}
async function openingFor(records, period) {
  const snapshots = (await getAllFromStore(STORE_WORK)).filter(r => r.id.startsWith('opening:') && r.period <= period)
    .sort((a,b) => b.period.localeCompare(a.period));
  const baseline = snapshots[0];
  const balances = baseline ? {...baseline.balances} : Object.fromEntries(VOUCHER_ACCOUNTS.map(a => [a.code, (a.dir === '借' ? 1 : -1) * Math.round((a.opening || 0) * 100)]));
  for (const v of records) {
    if (v.status !== 'POSTED' || v.period >= period || (baseline && v.period < baseline.period)) continue;
    for (const e of v.entries) balances[e.account] = (balances[e.account] || 0) + Math.round(e.debit*100)-Math.round(e.credit*100);
  }
  return balances;
}
function fromCents(value) { return money(value / 100); }
function renderBookDetail() {
  const tb = byId('acc-detail-body'); tb.replaceChildren();
  const account = currentBook.find(row => row.code === val('acc-detail-account'));
  if (!account) return;
  let running = account.initial;
  for (const e of account.entries) {
    running += Math.round(e.debit*100)-Math.round(e.credit*100);
    const row = tableRow(tb,[e.date,e.no,e.summary + (e.aux ? `（${e.aux}）` : ''),money(e.debit),money(e.credit),fromCents(running)]);
    if (e.source) row.cells[1].title = '来源：'+e.source;
  }
}
function renderBookReport() {
  const node = byId('acc-report-content'); node.replaceChildren();
  const table = document.createElement('table'); table.className='acc-table'; node.append(table);
  const body = table.createTBody();
  const rows = bookReport === 'bs'
    ? currentBook.filter(r => /^[1245]/.test(r.code)).map(r => [r,r.end])
    : currentBook.filter(r => r.code.startsWith('6')).map(r => [r,r.entries.filter(e => e.biz !== '期末结转').reduce((sum,e) => sum+Math.round(e.credit*100)-Math.round(e.debit*100),0)]);
  const total = prefix => rows.filter(([r]) => prefix.test(r.code)).reduce((sum,[,amt]) => sum+amt,0);
  tableRow(body,['科目','账套内部核对金额']);
  if (bookReport === 'bs') {
    const profit = currentBook.filter(r => r.code.startsWith('6')).reduce((sum,r) => sum-r.end,0);
    tableRow(body,['资产合计',fromCents(total(/^[15]/))]);
    tableRow(body,['负债合计',fromCents(-total(/^2/))]);
    tableRow(body,['权益及未结转损益合计',fromCents(-total(/^4/) + profit)]);
  } else tableRow(body,['本期损益合计',fromCents(total(/^6/))]);
  for (const [r,amt] of rows) {
    const row = tableRow(body,['',fromCents(bookReport === 'bs' && !/^[15]/.test(r.code) ? -amt : amt)]);
    row.firstChild.append(action(r.code+' '+r.name, () => {
      byId('acc-detail-account').value = r.code; renderBookDetail();
      document.querySelector('#acc-tabs [data-tab="detail"]').click();
    }));
  }
}
async function renderBooks(records, period) {
  const opening = await openingFor(records,period), note = await work('report-note:'+period);
  if (byId('acc-period-select').value !== period) return;
  currentBook = bookRows(records,period,opening);
  for (const id of ['summary-body','ledger-body','trial-body','summary-foot','trial-foot']) byId('acc-'+id).replaceChildren();
  const trial = currentBook.map(r => [Math.max(r.initial,0),Math.max(-r.initial,0),r.debit,r.credit,Math.max(r.end,0),Math.max(-r.end,0)]);
  currentBook.forEach((r,i) => {
    tableRow(byId('acc-summary-body'),[r.code,r.name,fromCents(r.debit),fromCents(r.credit)]);
    tableRow(byId('acc-ledger-body'),[r.code,r.name,...[r.initial,r.debit,r.credit,r.end].map(fromCents)]);
    tableRow(byId('acc-trial-body'),[r.code,r.name,...trial[i].map(fromCents)]);
  });
  const sums = trial[0].map((_,i) => trial.reduce((sum,row) => sum+row[i],0));
  tableRow(byId('acc-summary-foot'),['合计','',...sums.slice(2,4).map(fromCents)]);
  tableRow(byId('acc-trial-foot'),['合计','',...sums.map(fromCents)]);
  byId('acc-trial-check').textContent = [0,2,4].every(i => sums[i] === sums[i+1]) ? '期初、本期、期末借贷合计相等[cite: 1]' : '试算不平，请复核[cite: 1]';
  byId('acc-period-summary').textContent = '已记账 '+records.filter(v => v.period===period && v.status==='POSTED').length+' 张';
  byId('acc-report-note').value = note?.note || '';
  renderBookDetail(); renderBookReport();
}
function wireBookEditors(period, refresh) {
  const accounts = byId('acc-detail-account');
  accounts.replaceChildren(...VOUCHER_ACCOUNTS.map(a => new Option(a.code+' '+a.name,a.code)));
  accounts.addEventListener('change',renderBookDetail);
  window.switchReportTab = type => {
    bookReport = type;
    document.querySelectorAll('#module-accounting .acc-report-subtab').forEach(tab => tab.classList.toggle('active',tab.dataset.rtab === type));
    renderBookReport();
  };
  window.toggleOpeningEditor = () => {
    const box = byId('acc-opening-editor'), body = byId('acc-opening-body');
    if (box.style.display !== 'none') { box.style.display = 'none'; return; }
    const selectedPeriod = period(); body.dataset.period = selectedPeriod;
    body.replaceChildren();
    for (const a of currentBook) {
      const row = tableRow(body,[a.code,a.name,'','']);
      const toInput = (cell,value) => {
        const input = document.createElement('input'); input.type='number'; input.step='0.01'; input.min='0';
        input.value = value ? (value/100).toFixed(2) : ''; cell.append(input);
      };
      toInput(row.cells[2],Math.max(a.initial,0)); toInput(row.cells[3],Math.max(-a.initial,0));
    }
    box.style.display = 'block';
  };
  window.saveOpeningFromEditor = () => runAction(async () => {
    const p = period(), body = byId('acc-opening-body');
    if (p !== body.dataset.period) throw Error('期间已变化，请重新打开期初设置[cite: 1]');
    await ensureOpen(p);
    const history = await getAllFromStore(STORE_VOUCHER);
    if (history.some(v => v.status === 'POSTED' && v.period < p)) throw Error('该期由前期账簿结转，请在启用期间设置期初[cite: 1]');
    const balances = {};
    for (const row of body.rows) {
      const [d,c] = [...row.querySelectorAll('input')].map(input => cents(input.value || '0'));
      if (d < 0 || c < 0 || (d && c)) throw Error('期初每科目只能填写一侧正金额[cite: 1]');
      balances[row.cells[0].textContent] = d-c;
    }
    if (Object.values(balances).reduce((a,b)=>a+b,0) !== 0) throw Error('期初借贷不平[cite: 1]');
    await saveToStore(STORE_WORK,{id:'opening:'+p,period:p,balances});
    await refresh(); byId('acc-opening-editor').style.display='none'; showToast('期初余额已保存[cite: 1]');
  });
  window.saveReportNote = () => runAction(async () => {
    const p = period(); await saveToStore(STORE_WORK,{id:'report-note:'+p,period:p,note:val('acc-report-note')}); showToast('编制说明已保存[cite: 1]');
  });
}

/* ================= 5.7 融合核心：七列科目余额表 + 右键穿透 + 单位往来三栏账 ================= */
async function getPostedEntries(filter = {}) {
  const all = await getAllFromStore(STORE_VOUCHER);
  const rows = [];
  all.filter(v => v.status === 'POSTED').forEach(v => {
    (v.entries || []).forEach((e, idx) => {
      if (filter.subject && e.account !== filter.subject) return;
      if (filter.aux && (e.aux || '') !== filter.aux) return;
      rows.push({
        date: v.date, no: v.voucherNo, memo: e.summary || v.bizLabel || '',
        subject: e.account, aux: e.aux || '',
        debit: Number(e.debit) || 0, credit: Number(e.credit) || 0, idx
      });
    });
  });
  rows.sort((a, b) => a.date.localeCompare(b.date) || a.no.localeCompare(b.no) || a.idx - b.idx);
  return rows;
}

async function computeLiveBalances() {
  const all = await getAllFromStore(STORE_VOUCHER);
  const posted = all.filter(v => v.status === 'POSTED');
  const coreSubjects = VOUCHER_ACCOUNTS.filter(s => s.opening > 0 || posted.some(v => (v.entries || []).some(e => e.account === s.code)));
  const map = {};
  coreSubjects.forEach(s => {
    map[s.code] = { ...s, debit: 0, credit: 0, closing: s.opening, closeDir: s.dir };
  });
  posted.forEach(v => (v.entries || []).forEach(e => {
    const m = map[e.account];
    if (!m) return;
    m.debit += Number(e.debit) || 0;
    m.credit += Number(e.credit) || 0;
  }));
  coreSubjects.forEach(s => {
    const m = map[s.code];
    const net = m.debit - m.credit;
    let closing = s.dir === '借' ? (m.opening + net) : (m.opening - net);
    let closeDir = s.dir;
    if (closing < 0) { closing = -closing; closeDir = s.dir === '借' ? '贷' : '借'; }
    m.closing = closing;
    m.closeDir = closeDir;
  });
  return { list: coreSubjects.map(s => map[s.code]), postedCount: posted.length };
}

async function buildThreeColumnData(subjectCode, aux = '') {
  const subj = SUBJ_MAP[subjectCode];
  let opening = 0;
  if (!aux) opening = subj.opening || 0;
  else if (subjectCode === '1002') opening = BANK_SPECIAL_ACCOUNTS.find(x => x.name === aux)?.opening || 0;
  else if (subjectCode === '1122') opening = AR_OPENING[aux] || 0;
  else if (subjectCode === '2202') opening = AP_OPENING[aux] || 0;

  let bal = subj.dir === '借' ? opening : -opening;
  const entries = await getPostedEntries({ subject: subjectCode, aux });
  let debitTotal = 0, creditTotal = 0;
  const rows = entries.map(r => {
    debitTotal += r.debit; creditTotal += r.credit;
    bal += r.debit - r.credit;
    return { ...r, dir: bal >= 0 ? '借' : '贷', balance: Math.abs(bal) };
  });
  return {
    subj, opening, openingDir: subj.dir, rows, debitTotal, creditTotal,
    closing: Math.abs(bal), closeDir: bal >= 0 ? '借' : '贷'
  };
}

function openLiveDrawer(title, sub, html) {
  byId('live-drawer-title').textContent = title;
  byId('live-drawer-sub').textContent = sub || '';
  byId('live-drawer-body').innerHTML = html;
  byId('live-drawer-mask').classList.add('open');
}
function closeLiveDrawer() { byId('live-drawer-mask').classList.remove('open'); }

async function openBankSpecialDrawer() {
  const rows1002 = await getPostedEntries({ subject: '1002' });
  const list = BANK_SPECIAL_ACCOUNTS.map(b => {
    const subRows = rows1002.filter(r => r.aux === b.name);
    const inflow = subRows.reduce((s, r) => s + r.debit, 0);
    const outflow = subRows.reduce((s, r) => s + r.credit, 0);
    return { ...b, inflow, outflow, closing: b.opening + inflow - outflow };
  });
  const totalOpening = list.reduce((s, b) => s + b.opening, 0);
  const totalIn = list.reduce((s, b) => s + b.inflow, 0);
  const totalOut = list.reduce((s, b) => s + b.outflow, 0);
  const totalClosing = list.reduce((s, b) => s + b.closing, 0);

  let html = `<div class="sum-bar">
    <div><span>期初合计</span><strong>${fmtNum(totalOpening)}</strong></div>
    <div><span>本期收入</span><strong>${fmtNum(totalIn)}</strong></div>
    <div><span>本期支出</span><strong>${fmtNum(totalOut)}</strong></div>
    <div><span>期末可用合计</span><strong>${fmtNum(totalClosing)}</strong></div>
  </div><div class="bank-grid">`;

  list.forEach(b => {
    html += `<div class="bank-card">
      <h4>${escapeHtml(b.name)}</h4>
      <div class="acct">${escapeHtml(b.acct)} · ${b.currency}</div>
      <div class="bal-label">当前可用余额</div>
      <div class="bal">￥${fmtNum(b.closing)}</div>
      <div class="bank-meta"><div>期初 <b>${fmtNum(b.opening)}</b></div><div class="out">支出 <b>${fmtNum(b.outflow)}</b></div><div>收入 <b>${fmtNum(b.inflow)}</b></div></div>
    </div>`;
  });
  html += `</div><h3 class="serif-title" style="font-size:18px; margin:6px 0 12px;">账户资金变动明细</h3>
    <div style="overflow:auto;"><table class="fin-table"><thead><tr>
    <th width="110">日期</th><th width="130">凭证号</th><th>摘要</th><th width="150">银行账户</th>
    <th class="num" width="130">收入</th><th class="num" width="130">支出</th></tr></thead><tbody>`;
  if (!rows1002.length) {
    html += '<tr><td colspan="6" style="text-align:center; color:var(--text-muted); padding:24px;">本期无银行存款变动记录</td></tr>';
  } else {
    rows1002.forEach(r => {
      html += `<tr><td>${r.date}</td><td class="code">${escapeHtml(r.no)}</td><td>${escapeHtml(r.memo)}</td>
        <td>${escapeHtml(r.aux || '未指定账户')}</td><td class="num">${fmtOrBlank(r.debit)}</td><td class="num">${fmtOrBlank(r.credit)}</td></tr>`;
    });
  }
  html += '</tbody></table></div>';
  openLiveDrawer('1002 银行存款 · 银行专项分类明细', '按账户维度拆分的期初、本期收支与当前可用余额', html);
}

async function openUnitBalanceDrawer(code) {
  const subj = SUBJ_MAP[code];
  const units = code === '1122' ? AR_UNITS : AP_UNITS;
  const openingMap = code === '1122' ? AR_OPENING : AP_OPENING;
  const allEntries = await getPostedEntries({ subject: code });
  const list = units.map(u => {
    const opening = openingMap[u] || 0;
    const rows = allEntries.filter(r => r.aux === u);
    const dr = rows.reduce((s, r) => s + r.debit, 0);
    const cr = rows.reduce((s, r) => s + r.credit, 0);
    const closing = subj.dir === '借' ? (opening + dr - cr) : (opening + cr - dr);
    return { unit: u, opening, dr, cr, closing, dir: subj.dir };
  });

  let html = `<div class="sum-bar">
    <div><span>期初合计</span><strong>${fmtNum(list.reduce((s, u) => s + u.opening, 0))}</strong></div>
    <div><span>本期借方</span><strong>${fmtNum(list.reduce((s, u) => s + u.dr, 0))}</strong></div>
    <div><span>本期贷方</span><strong>${fmtNum(list.reduce((s, u) => s + u.cr, 0))}</strong></div>
    <div><span>期末合计</span><strong>${fmtNum(list.reduce((s, u) => s + u.closing, 0))}</strong></div>
  </div><div style="overflow:auto;"><table class="fin-table"><thead><tr>
    <th width="60">序号</th><th>往来单位</th><th width="70">方向</th>
    <th class="num" width="140">期初余额</th><th class="num" width="140">本期借方</th>
    <th class="num" width="140">本期贷方</th><th class="num" width="140">期末余额</th><th width="100">操作</th></tr></thead><tbody>`;

  list.forEach((u, i) => {
    html += `<tr><td class="code">${i + 1}</td><td>${escapeHtml(u.unit)}</td>
      <td><span class="dir-badge ${u.dir === '借' ? 'dr' : 'cr'}">${u.dir}</span></td>
      <td class="num">${fmtNum(u.opening)}</td><td class="num">${fmtNum(u.dr)}</td><td class="num">${fmtNum(u.cr)}</td>
      <td class="num strong">${fmtNum(u.closing)}</td>
      <td><button type="button" class="btn-mini" data-jump-unit="${escapeHtml(u.unit)}">三栏明细</button></td></tr>`;
  });
  html += '</tbody></table></div>';
  openLiveDrawer(`${code} ${subj.name} · 单位往来余额表`, '按往来单位汇总期初余额、本期发生额与期末余额', html);

  byId('live-drawer-body').querySelectorAll('button[data-jump-unit]').forEach(btn => {
    btn.addEventListener('click', () => {
      closeLiveDrawer();
      jumpToUnitThreeColumn(code, btn.dataset.jumpUnit);
    });
  });
}

async function openSubjectThreeColumnDrawer(code) {
  const data = await buildThreeColumnData(code, '');
  let html = `<div class="sum-bar">
    <div><span>期初余额</span><strong>${data.openingDir} ${fmtNum(data.opening)}</strong></div>
    <div><span>本期借方</span><strong>${fmtNum(data.debitTotal)}</strong></div>
    <div><span>本期贷方</span><strong>${fmtNum(data.creditTotal)}</strong></div>
    <div><span>期末余额</span><strong>${data.closeDir} ${fmtNum(data.closing)}</strong></div>
  </div><div style="overflow:auto;"><table class="fin-table"><thead><tr>
    <th width="110">日期</th><th width="130">凭证号</th><th>摘要</th>
    <th class="num" width="130">借方</th><th class="num" width="130">贷方</th><th width="70">方向</th><th class="num" width="140">余额</th></tr></thead><tbody>
    <tr class="opening-row"><td colspan="3">期初余额</td><td class="num"></td><td class="num"></td><td>${data.openingDir}</td><td class="num">${fmtNum(data.opening)}</td></tr>`;

  if (!data.rows.length) {
    html += '<tr><td colspan="7" style="text-align:center; color:var(--text-muted); padding:24px;">本期无已记账发生额</td></tr>';
  } else {
    data.rows.forEach(r => {
      html += `<tr><td>${r.date}</td><td class="code">${escapeHtml(r.no)}</td><td>${escapeHtml(r.memo)}${r.aux ? '（' + escapeHtml(r.aux) + '）' : ''}</td>
        <td class="num">${fmtOrBlank(r.debit)}</td><td class="num">${fmtOrBlank(r.credit)}</td><td>${r.dir}</td><td class="num strong">${fmtNum(r.balance)}</td></tr>`;
    });
  }
  html += `<tr class="total-row"><td colspan="3">本期合计</td><td class="num">${fmtNum(data.debitTotal)}</td><td class="num">${fmtNum(data.creditTotal)}</td><td>${data.closeDir}</td><td class="num">${fmtNum(data.closing)}</td></tr></tbody></table></div>`;
  openLiveDrawer(`${code} ${data.subj.name} · 三栏明细账`, '已记账凭证逐笔计算余额', html);
}

let ctxOpen = false;
function hideCtxMenu() { byId('ctx-menu').classList.remove('open'); ctxOpen = false; }
function showCtxMenu(x, y, title, items) {
  const menu = byId('ctx-menu');
  menu.innerHTML = '';
  const titleEl = document.createElement('div');
  titleEl.className = 'ctx-title';
  titleEl.textContent = title;
  menu.appendChild(titleEl);
  items.forEach(it => {
    if (it.divider) { const d = document.createElement('div'); d.className = 'ctx-divider'; menu.appendChild(d); return; }
    const el = document.createElement('div');
    el.className = 'ctx-item' + (it.cls ? ' ' + it.cls : '');
    el.innerHTML = `<span class="ctx-icon">${it.icon || ''}</span><span>${escapeHtml(it.label)}</span>`;
    el.addEventListener('click', e => { e.stopPropagation(); hideCtxMenu(); it.action(); });
    menu.appendChild(el);
  });
  menu.style.left = '0px'; menu.style.top = '0px'; menu.classList.add('open'); ctxOpen = true;
  const rect = menu.getBoundingClientRect();
  menu.style.left = (x + rect.width > window.innerWidth - 8 ? Math.max(8, window.innerWidth - rect.width - 8) : x) + 'px';
  menu.style.top = (y + rect.height > window.innerHeight - 8 ? Math.max(8, window.innerHeight - rect.height - 8) : y) + 'px';
}

function openSubjectContextMenu(x, y, code) {
  const subj = SUBJ_MAP[code];
  if (!subj) return;
  const items = [];
  if (code === '1002') {
    items.push({ icon: '🏦', label: '银行专项分类明细', action: () => runAction(openBankSpecialDrawer) });
    items.push({ icon: '📊', label: '查看该科目三栏明细账', action: () => runAction(() => openSubjectThreeColumnDrawer('1002')) });
    items.push({ divider: true });
    items.push({ icon: '🧾', label: '跳转到会计记账明细', action: () => jumpToAccountingDetail('1002') });
  } else if (code === '1122' || code === '2202') {
    const units = code === '1122' ? AR_UNITS : AP_UNITS;
    items.push({ icon: '📋', label: '单位往来余额表', action: () => runAction(() => openUnitBalanceDrawer(code)) });
    items.push({ icon: '📊', label: '单位往来三栏明细账', action: () => jumpToUnitThreeColumn(code, units[0]) });
    units.forEach(u => {
      items.push({ icon: ' ', label: '└ ' + u, cls: 'ctx-sub', action: () => jumpToUnitThreeColumn(code, u) });
    });
    items.push({ divider: true });
    items.push({ icon: '🧾', label: '跳转到会计记账明细', action: () => jumpToAccountingDetail(code) });
  } else {
    items.push({ icon: '📊', label: '查看该科目三栏明细账', action: () => runAction(() => openSubjectThreeColumnDrawer(code)) });
    items.push({ icon: '🧾', label: '跳转到会计记账明细', action: () => jumpToAccountingDetail(code) });
  }
  showCtxMenu(x, y, `${code}  ${subj.name}（${subj.dir}）`, items);
}

function jumpToAccountingDetail(code) {
  openWorkspaceTab('gl-query');
  setTimeout(() => {
    byId('acc-detail-account').value = code;
    renderBookDetail();
    document.querySelector('#acc-tabs [data-tab="detail"]')?.click();
  }, 60);
}

function jumpToUnitThreeColumn(code, unit) {
  openWorkspaceTab('ar-subledger');
  byId('ar-sub-subject').value = code;
  renderARSubledgerModule(unit);
}

window.switchBalanceMode = function(mode) {
  byId('bal-mode-live').classList.toggle('active', mode === 'live');
  byId('bal-mode-excel').classList.toggle('active', mode === 'excel');
  byId('bal-view-live').style.display = mode === 'live' ? 'block' : 'none';
  byId('bal-view-excel').style.display = mode === 'excel' ? 'block' : 'none';
  if (mode === 'live') renderLiveBalance();
};

window.renderLiveBalance = async function() {
  const { list, postedCount } = await computeLiveBalances();
  const tbody = byId('live-balance-body');
  tbody.innerHTML = list.map(m => {
    let tag = '';
    if (m.code === '1002') tag = '<span class="trace-tag">银行专项</span>';
    else if (m.code === '1122' || m.code === '2202') tag = '<span class="trace-tag">单位往来</span>';
    const dirCls = m.closeDir === '借' ? 'dr' : 'cr';
    return `<tr class="clickable" data-code="${m.code}">
      <td class="code">${m.code}</td>
      <td>${escapeHtml(m.name)}${tag}</td>
      <td><span class="dir-badge ${dirCls}">${m.closeDir}</span></td>
      <td class="num">${fmtNum(m.opening)}</td>
      <td class="num">${fmtNum(m.debit)}</td>
      <td class="num">${fmtNum(m.credit)}</td>
      <td class="num strong">${fmtNum(m.closing)}</td>
    </tr>`;
  }).join('');

  let opDr = 0, opCr = 0, sumDr = 0, sumCr = 0, clDr = 0, clCr = 0;
  list.forEach(m => {
    if (m.dir === '借') opDr += m.opening; else opCr += m.opening;
    sumDr += m.debit; sumCr += m.credit;
    if (m.closeDir === '借') clDr += m.closing; else clCr += m.closing;
  });
  byId('live-balance-foot').innerHTML = `<tr>
    <td colspan="3" style="text-align:right;">合 计</td>
    <td class="num">借 ${fmtNum(opDr)}<br>贷 ${fmtNum(opCr)}</td>
    <td class="num">${fmtNum(sumDr)}</td>
    <td class="num">${fmtNum(sumCr)}</td>
    <td class="num">借 ${fmtNum(clDr)}<br>贷 ${fmtNum(clCr)}</td>
  </tr>`;
  byId('live-balance-note').innerHTML = `期初借贷合计：借 ${fmtNum(opDr)} / 贷 ${fmtNum(opCr)} · 本期借贷发生额：借 ${fmtNum(sumDr)} / 贷 ${fmtNum(sumCr)} · 期末借贷合计：借 ${fmtNum(clDr)} / 贷 ${fmtNum(clCr)}（共 ${postedCount} 张已记账凭证已计入）。`;
};

function initARSubledgerModule() {
  byId('ar-sub-subject').addEventListener('change', () => renderARSubledgerModule());
  byId('ar-sub-unit').addEventListener('change', () => renderARSubledgerTable());
  byId('btn-ar-sub-export').addEventListener('click', () => runAction(async () => {
    const code = val('ar-sub-subject'), unit = val('ar-sub-unit'), data = await buildThreeColumnData(code, unit);
    const rows = [
      ['', '', '期初余额', '', '', data.openingDir, data.opening.toFixed(2)],
      ...data.rows.map(r => [r.date, r.no, r.memo, r.debit ? r.debit.toFixed(2) : '', r.credit ? r.credit.toFixed(2) : '', r.dir, r.balance.toFixed(2)]),
      ['', '', '本期合计', data.debitTotal.toFixed(2), data.creditTotal.toFixed(2), data.closeDir, data.closing.toFixed(2)]
    ];
    exportRows(`单位往来三栏账_${code}_${unit}`, ['日期', '凭证号', '摘要', '借方金额', '贷方金额', '方向', '余额'], rows);
    showToast('单位往来三栏账已导出');
  }));
}

window.renderARSubledgerModule = async function(preferUnit = '') {
  const code = byId('ar-sub-subject').value;
  const units = code === '1122' ? AR_UNITS : AP_UNITS;
  const sel = byId('ar-sub-unit');
  const prev = preferUnit || sel.value;
  sel.innerHTML = units.map(u => `<option value="${escapeHtml(u)}">${escapeHtml(u)}</option>`).join('');
  sel.value = units.includes(prev) ? prev : units[0];
  await renderARSubledgerTable();
};

async function renderARSubledgerTable() {
  const code = byId('ar-sub-subject').value, unit = byId('ar-sub-unit').value;
  if (!unit) return;
  const data = await buildThreeColumnData(code, unit);
  byId('ar-sub-summary').innerHTML = `
    <div><span>往来科目</span><strong>${code} ${escapeHtml(data.subj.name)}</strong></div>
    <div><span>往来单位</span><strong>${escapeHtml(unit)}</strong></div>
    <div><span>期初余额</span><strong>${data.openingDir} ${fmtNum(data.opening)}</strong></div>
    <div><span>本期借方</span><strong>${fmtNum(data.debitTotal)}</strong></div>
    <div><span>本期贷方</span><strong>${fmtNum(data.creditTotal)}</strong></div>
    <div><span>期末余额</span><strong>${data.closeDir} ${fmtNum(data.closing)}</strong></div>`;

  let html = `<tr class="opening-row"><td colspan="3">期初余额</td><td class="num"></td><td class="num"></td><td>${data.openingDir}</td><td class="num">${fmtNum(data.opening)}</td></tr>`;
  if (!data.rows.length) {
    html += '<tr><td colspan="7" style="text-align:center; color:var(--text-muted); padding:26px;">该往来单位本期无已记账发生额</td></tr>';
  } else {
    data.rows.forEach(r => {
      html += `<tr><td>${r.date}</td><td class="code">${escapeHtml(r.no)}</td><td>${escapeHtml(r.memo)}</td>
        <td class="num">${fmtOrBlank(r.debit)}</td><td class="num">${fmtOrBlank(r.credit)}</td><td>${r.dir}</td><td class="num strong">${fmtNum(r.balance)}</td></tr>`;
    });
  }
  html += `<tr class="total-row"><td colspan="3">本期合计 / 期末余额</td><td class="num">${fmtNum(data.debitTotal)}</td><td class="num">${fmtNum(data.creditTotal)}</td><td>${data.closeDir}</td><td class="num">${fmtNum(data.closing)}</td></tr>`;
  byId('ar-sub-body').innerHTML = html;
}

/* ================= 5.8 原有模块（对账 / 应收函证 / 税务 / 结转 / Excel导入穿透） ================= */
function readCsv(text) {
  const rows = []; let row = [], value = '', quoted = false;
  text = text.replace(/^\ufeff/,'');
  for (let i=0;i<text.length;i++) {
    const c = text[i];
    if (c === '"') {
      if (quoted && text[i+1] === '"') { value += '"'; i++; }
      else quoted = !quoted;
    } else if (!quoted && (c === ',' || c === '\n' || c === '\r')) {
      row.push(value); value = '';
      if (c !== ',') {
        if (row.some(Boolean)) rows.push(row);
        row = []; if (c === '\r' && text[i+1] === '\n') i++;
      }
    } else value += c;
  }
  if (quoted) throw Error('CSV 引号未闭合[cite: 1]');
  row.push(value); if (row.some(Boolean)) rows.push(row);
  return rows;
}
function compareFlows(bank, company) {
  const grouped = new Map();
  for (const [side,rows] of [['bank',bank],['company',company]]) {
    for (const row of rows) {
      const key = row.id || side + ':' + row.line;
      if (!grouped.has(key)) grouped.set(key,{id:row.id,bank:[],company:[]});
      grouped.get(key)[side].push(row);
    }
  }
  return [...grouped.values()].map(({id,bank,company}) => {
    const b = bank[0], c = company[0];
    const bad = !id || bank.length>1 || company.length>1 || [...bank,...company].some(r => r.amount === null);
    const status = bad ? 'ERROR' : !b ? 'COMPANY_ONLY' : !c ? 'BANK_ONLY' : b.amount === c.amount ? 'MATCH' : 'DIFF';
    const note = bad ? '单号重复、缺失或金额无效[cite: 1]' : status === 'MATCH' ? '单号与金额相同[cite: 1]' : '待核实[cite: 1]';
    return {id,status,bank,company,note};
  });
}
function initReconModule() {
  const data = {}, sides = ['bank','company'], reconTabs = [...document.querySelectorAll('#recon-tabs .tab')];
  const labels = {MATCH:'完全匹配', BANK_ONLY:'仅银行有',COMPANY_ONLY:'仅账面有',DIFF:'金额不符',ERROR:'异常'};
  let filter = 'ALL', results = [], verified = false, resultKey = null;
  const area = byId('recon-mapping-area');
  const control = document.createElement('div'); control.className = 'mapping-area'; control.style.display = 'block';
  const balances = [['银行期末余额','recon-bank-end'],['账面期末余额','recon-book-end'],['银行侧已核实调整净额','recon-bank-adjust'],['账面侧已核实调整净额','recon-book-adjust']];
  control.append(makeField('对账期间','recon-period','month'), makeField('银行账户','recon-account'));
  const inputs = document.createElement('div'); inputs.className='form-grid';
  for (const spec of balances) inputs.append(makeField(...spec,'number'));
  inputs.append(makeField('调整及差异核实依据','recon-evidence')); control.append(inputs); area.before(control);
  const key = () => val('recon-period')+':'+val('recon-account');
  let adjustment = null;
  function readAdjustment() {
    if (!val('recon-bank-end') || !val('recon-book-end')) return null;
    const [bank,book,bankAdj,bookAdj] = balances.map(([,id])=>cents(val(id)||'0'));
    return {bank,book,bankAdj,bookAdj,evidence:val('recon-evidence'),gap:bank+bankAdj-book-bookAdj};
  }
  byId('recon-period').value = localDay().slice(0,7);
  const btn = byId('recon-btn-reconcile');
  const controls = byId('recon-actions');
  const summary = document.createElement('p'); summary.className = 'note-text recon-summary';
  summary.setAttribute('role', 'status'); controls.after(summary);
  const save = action('保存对账结果',async () => {
    const period=val('recon-period'), account=val('recon-account');
    if (!period || !account || !results.length || resultKey !== key()) throw Error('请填写期间、账户并运行对账[cite: 1]');
    await ensureOpen(period);
    adjustment = readAdjustment();
    verified = Boolean(adjustment && adjustment.gap === 0 &&
      (adjustment.evidence || (!adjustment.bankAdj && !adjustment.bookAdj && results.every(r=>r.status==='MATCH'))) &&
      !results.some(r=>r.status==='ERROR'));
    await saveToStore(STORE_WORK,{id:'recon:'+period+':'+account,period,account,results,verified,adjustment,createdAt:new Date().toISOString()}); render();
    showToast('对账记录已保存[cite: 1]');
  }, 'btn');
  controls.append(save, action('读取已保存对账',async () => {
    const saved = await work('recon:'+val('recon-period')+':'+val('recon-account'));
    if (!saved) throw Error('该期间及账户暂无对账记录[cite: 1]');
    results = saved.results; resultKey=key(); verified=saved.verified; adjustment=saved.adjustment;
    ['bank','book','bankAdj','bookAdj'].forEach((field,i)=>byId(balances[i][1]).value=adjustment ? adjustment[field]/100 : '');
    byId('recon-evidence').value=adjustment?.evidence || ''; render();
  }, 'btn'));
  const visible = () => filter === 'ALL' ? results : results.filter(r => r.status === filter);
  const amounts = rows => rows.map(r => r.amount === null ? '无效' : fromCents(r.amount)).join(' / ');
  function render() {
    byId('recon-result-area').style.display = 'block';
    const tb = byId('recon-result-body'); tb.replaceChildren();
    for (const r of visible()) {
      const lines = [...r.bank.map(x=>'银行第'+x.line+'行'),...r.company.map(x=>'账面第'+x.line+'行')];
      tableRow(tb,[r.id,labels[r.status],amounts(r.bank),amounts(r.company),r.note+' · '+lines.join('、')]);
    }
    for (const tab of reconTabs) byId('badge-'+tab.dataset.type).textContent = tab.dataset.type === 'ALL' ? results.length : results.filter(r => r.status === tab.dataset.type).length;
    summary.textContent = !results.length ? '暂无对账结果' : adjustment
      ? '银行调节后 '+fromCents(adjustment.bank+adjustment.bankAdj)+' ／ 账面调节后 '+fromCents(adjustment.book+adjustment.bookAdj)+' ／ 差额 '+fromCents(adjustment.gap)+(verified?' · 已保存核对结果':' · 待核实')
      : '请补录双方期末余额，核对差异并保存[cite: 1]。';
  }
  for (const side of sides) {
    const input = byId('recon-file-'+side), zone = byId('recon-drop-'+side);
    zone.addEventListener('click',e => { if (e.target !== input) input.click(); });
    async function loadFile(file) {
      if (!file) return;
      delete data[side]; btn.disabled=true; results=[]; verified=false; render();
      let rows;
      if (/\.csv$/i.test(file.name)) rows = readCsv(await file.text());
      else {
        if (!window.XLSX) throw Error('Excel 组件未加载，请联网重试或另存为 UTF-8 CSV[cite: 1]');
        const book = XLSX.read(await file.arrayBuffer());
        rows = XLSX.utils.sheet_to_json(book.Sheets[book.SheetNames[0]],{header:1,raw:false,defval:''});
      }
      if (rows.length < 2 || (rows[0] || []).length < 2) throw Error('需要表头、单号列和金额列[cite: 1]');
      data[side]=rows.slice(1);
      for (const field of ['id','amount']) {
        const select = byId('recon-map-'+side+'-'+field);
        select.replaceChildren(...rows[0].map((name,i) => new Option(name || '第'+(i+1)+'列',i)));
        select.selectedIndex=field==='id' ? 0 : 1;
      }
      zone.querySelector('p').textContent='已载入 '+file.name;
      area.style.display='block'; btn.disabled=!sides.every(k=>data[k]);
    }
    input.addEventListener('change', () => { const file=input.files[0]; input.value=''; runAction(()=>loadFile(file)); });
    zone.addEventListener('dragover', e => e.preventDefault());
    zone.addEventListener('drop', e => { e.preventDefault(); runAction(()=>loadFile(e.dataTransfer.files[0])); });
  }
  onAction('recon-btn-reconcile', () => {
    const records = sides.map(side => {
      const idCol=val('recon-map-'+side+'-id'), amtCol=val('recon-map-'+side+'-amount');
      if (idCol===amtCol) throw Error('请选择不同的单号列和金额列[cite: 1]');
      return data[side].map((row,i) => {
        let amt=null; try { amt=cents(String(row[amtCol] ?? '').replaceAll(',','')); } catch {}
        return {id:String(row[idCol] ?? '').trim(),amount:amt,line:i+2};
      });
    });
    results=compareFlows(...records); resultKey=key(); verified=false; adjustment=readAdjustment(); render();
  });
  reconTabs.forEach(tab => tab.addEventListener('click', () => {
    filter=tab.dataset.type; reconTabs.forEach(t => t.classList.toggle('active',t===tab)); render();
  }));
  onAction('recon-btn-export', () => exportRows('对账结果',['单号','状态','银行金额','账面金额','说明'],
    visible().map(r=>[r.id,labels[r.status],amounts(r.bank),amounts(r.company),r.note])));
  controls.append(action('导出余额调节表', () => {
    if (!adjustment || resultKey !== key()) throw Error('请先录入双方期末余额并保存对账[cite: 1]');
    const a=adjustment; exportRows('余额调节表',['期间','账户','银行期末','银行调整','账面期末','账面调整','调节差额','依据'],[[val('recon-period'),val('recon-account'),a.bank/100,a.bankAdj/100,a.book/100,a.bookAdj/100,a.gap/100,a.evidence]]);
  }, 'btn'));
}

function initARModule() {
  const periodField=byId('ar-period'), customerBody=byId('ar-customer-list');
  periodField.value=localDay().slice(0,7);
  periodField.addEventListener('change',()=>loadARData());
  const fields = [['客户名称','ar-customer'],['单据编号','ar-source'],['日期','ar-date','date'],['金额','ar-amount','number']];
  const box = document.createElement('div'); box.className='card-block';
  const grid = document.createElement('div'); grid.className='form-grid';
  for (const spec of fields) grid.append(makeField(...spec));
  grid.append(makeField('记录类型','ar-type','select',[['due','已核对应收'],['receipt','待分配实收']]));
  box.append(grid);
  customerBody.closest('.module-grid').before(box);
  async function records() { return (await getAllFromStore(STORE_WORK)).filter(r=>r.id.startsWith('ar:')); }
  box.append(action('保存应收或回款',async () => {
    const customer=val('ar-customer'),source=val('ar-source'),date=val('ar-date'),type=val('ar-type'),amt=cents(val('ar-amount'));
    if (!customer || !source || !date || amt<=0) throw Error('请填写客户、日期、单据和正金额[cite: 1]');
    const period=date.slice(0,7); await ensureOpen(period);
    await saveRecords(STORE_WORK,[{id:'ar:'+type+':'+source,customer,source,date,period,type,amount:amt,allocations:[]}],true);
    for (const id of ['ar-source','ar-amount']) byId(id).value='';
    byId('ar-period').value=period; await loadARData(); showToast('已保存[cite: 1]');
  }, 'btn btn-primary'));
  byId('ar-date').value=localDay();
  function summarize(rows, period) {
    const totals=new Map();
    for (const r of rows.filter(r=>r.period<=period && ['due','receipt'].includes(r.type))) {
      const sum=totals.get(r.customer) || {balance:0,available:0};
      if (r.type==='due') sum.balance+=r.amount;
      else {
        const used=r.allocations.filter(a=>a.period<=period).reduce((n,a)=>n+a.amount,0);
        sum.available+=r.amount-used; sum.balance-=used;
      }
      totals.set(r.customer,sum);
    }
    return totals;
  }
  window.loadARData = () => runAction(async () => {
    const [rows,period]=[await records(),val('ar-period')], totals=summarize(rows,period);
    const diff=byId('ar-confirmation-diffs'), allocations=byId('ar-allocations'), options=byId('ar-unallocated');
    customerBody.replaceChildren(); diff.replaceChildren(); allocations.replaceChildren(); options.replaceChildren(new Option('选择待分配回款',''));
    for (const r of rows.filter(r=>r.type==='receipt' && r.period<=period)) {
      const left=r.amount-r.allocations.reduce((n,a)=>n+a.amount,0);
      if (left>0) options.add(new Option(r.customer+' · '+r.source+' · '+fromCents(left),r.id));
      for (const a of r.allocations.filter(a=>a.period===period)) tableRow(allocations,[r.customer,fromCents(a.amount),a.date,r.source]);
    }
    for (const [customer,sum] of totals) {
      const conf=rows.find(r=>r.type==='conf' && r.period===period && r.customer===customer);
      const delta=conf?.reply == null ? null : conf.reply-conf.balance;
      const status=!conf ? '未发函' : delta===null ? '未回函' : delta===0 ? '回函相符' : '回函有差异';
      const row=tableRow(customerBody,[customer,fromCents(sum.balance),fromCents(sum.available),status,delta===null?'—':fromCents(delta),'']);
      if (conf) row.lastChild.append(action('登记回函',async () => {
        const reply=await uiPrompt('回函确认金额'); if (reply===null) return;
        const evidence=(await uiPrompt('回函编号或核实依据'))?.trim(); if (!evidence) throw Error('请填写核实依据[cite: 1]');
        await saveToStore(STORE_WORK,{...conf,reply:cents(reply),evidence}); await loadARData();
      }));
      if (delta) tableRow(diff,[customer,fromCents(conf.balance),fromCents(conf.reply),fromCents(delta),conf.evidence,'待核实']);
    }
  });
  window.allocatePayment = () => runAction(async () => {
    const id=val('ar-unallocated'), period=val('ar-period'); await ensureOpen(period);
    const rows=await records(), receipt=rows.find(r=>r.type==='receipt' && r.id===id);
    if (!receipt || receipt.period>period) throw Error('请选择可用回款[cite: 1]');
    const value=await uiPrompt('本次分配金额'); if (value===null) return;
    const amt=cents(value), left=receipt.amount-receipt.allocations.reduce((n,a)=>n+a.amount,0);
    const balance=summarize(rows,period).get(receipt.customer).balance;
    if (amt<=0 || amt>left || amt>balance) throw Error('冲销金额超出可用回款或该客户应收余额[cite: 1]');
    await saveToStore(STORE_WORK,{...receipt,allocations:[...receipt.allocations,{amount:amt,period,date:localDay()}]});
    await loadARData(); showToast('冲销已保存[cite: 1]');
  });
  window.generateConfirmation = () => runAction(async () => {
    const period=val('ar-period'); if (!period.endsWith('-12')) throw Error('请选择年末期间[cite: 1]');
    const rows=await records(), list=[...summarize(rows,period)];
    if (!list.length) throw Error('没有应收记录[cite: 1]');
    exportRows(period+'应收询证函底稿',['客户','账面余额','截止期间','请填写回函金额'],list.map(([customer,sum])=>[customer,sum.balance/100,period,'']));
    if (!(await uiConfirm('如已实际发函，确认登记本批客户为未回函？'))) return;
    const newRecords=list.map(([customer,sum])=>({id:'ar:conf:'+period+':'+customer,type:'conf',customer,period,balance:sum.balance,reply:null}))
      .filter(r=>!rows.some(old=>old.id===r.id));
    if (newRecords.length) await saveRecords(STORE_WORK,newRecords,true);
    await loadARData();
  });
}

function initTaxModule() {
  const period=byId('tax-period'), amounts=['vat','surtax','cit'];
  period.value=localDay().slice(0,7);
  const fields=amounts.map((key,i)=>[['增值税确认金额','附加税费确认金额','企业所得税确认金额（季末）'][i],'tax-entry-'+key,'number']);
  const block=document.createElement('div'); block.className='form-grid';
  for (const spec of [...fields,['取数资料编号及计算口径版本','tax-basis']]) block.append(makeField(...spec));
  byId('tax-vat-amount').closest('.card-block').prepend(block);
  let draft=null;
  const quarterly=p=>['03','06','09','12'].includes(p.slice(-2));
  function show(record) {
    for (const key of amounts) byId('tax-'+key+'-amount').textContent=record?.values[key] != null ? fromCents(record.values[key]) : key==='cit' && !quarterly(period.value) ? '非季末' : '待录入';
    byId('tax-vat-detail').textContent=record ? '取数依据：'+record.basis : '录入财务已核对的测算金额及取数依据[cite: 1]';
  }
  const taskBody=byId('tax-task-list'), taskBlock=taskBody.closest('.card-block');
  const taskEditor=document.createElement('div'); taskEditor.className='form-grid'; taskEditor.style.marginTop='16px';
  const taskFields=[['税种 / 任务名称','task-name'],['截止日期','task-due','date'],['提交凭据','task-submit-ref'],
    ['申报反馈 / 回执编号','task-response'],['扣款结果 / 凭据','task-payment'],['完税证明编号','task-proof'],['交税凭证号','task-voucher']];
  taskFields.forEach(spec=>taskEditor.append(makeField(...spec)));
  const statusField=makeField('申报状态','task-status','select', [['未提交','未提交'],['已提交待反馈','已提交待反馈'],['申报成功','申报成功'],['申报失败','申报失败']]);
  taskEditor.append(statusField,makeField('回执或完税证明附件','task-file','file'));
  taskBlock.append(taskEditor);
  const taskActions = document.createElement('div');
  taskActions.className = 'section-actions tax-task-actions';
  taskEditor.append(taskActions);
  let editingTask=null;
  taskActions.append(action('新建申报任务', () => {
    editingTask=null; taskFields.forEach(([,id]) => byId(id).value=''); byId('task-status').selectedIndex=0; byId('task-file').value='';
  }, 'btn'));
  taskActions.append(action('保存申报任务',async () => {
    const [name,due,submit,response,payment,proof,voucher]=taskFields.map(([,id])=>val(id));
    const p=period.value,status=val('task-status');
    if (!name || !due || !p) throw Error('请填写税期、任务名称和截止日期[cite: 1]');
    if (status!=='未提交' && !submit) throw Error('请录入实际提交凭据[cite: 1]');
    if (['申报成功','申报失败'].includes(status) && !response) throw Error('请录入实际申报反馈[cite: 1]');
    const old=editingTask && await work(editingTask), file=byId('task-file').files[0];
    const attachment=file ? {name:file.name,data:await new Promise((resolve,reject)=>{
      const reader=new FileReader(); reader.onload=()=>resolve(reader.result); reader.onerror=reject; reader.readAsDataURL(file);
    })} : old?.attachment;
    await saveToStore(STORE_WORK,{id:old?.period===p ? old.id : 'tax-task:'+recordId(),period:p,name,due,submit,response,payment,proof,voucher,status,attachment});
    byId('task-file').value=''; await loadTaxData(); showToast('申报任务已保存[cite: 1]');
  }, 'btn btn-primary'));
  window.loadTaxData = () => runAction(async () => {
    const p=period.value, rows=await getAllFromStore(STORE_WORK); if (period.value!==p) return;
    draft=null; editingTask=null;
    for (const [,id] of taskFields) byId(id).value='';
    byId('task-status').selectedIndex=0; byId('task-file').value='';
    const latest=rows.filter(r=>r.id.startsWith('tax-result:') && r.period===p).sort((a,b)=>b.saved.localeCompare(a.saved))[0];
    for (const key of amounts) byId('tax-entry-'+key).value=latest?.values[key] != null ? latest.values[key]/100 : '';
    byId('tax-basis').value=latest?.basis || '';
    byId('tax-missing-items').style.display='none'; show(latest);
    taskBody.replaceChildren();
    for (const task of rows.filter(r=>r.id.startsWith('tax-task:') && r.period===p)) {
      const row=tableRow(taskBody,[task.name,task.status,task.due,'']);
      row.lastChild.append(action('编辑', () => {
        editingTask=task.id;
        const data=[task.name,task.due,task.submit,task.response,task.payment,task.proof,task.voucher];
        taskFields.forEach(([,id],i)=>byId(id).value=data[i]||'');
        byId('task-status').value=task.status;
      }));
      if (task.attachment) row.lastChild.append(action('附件',()=>downloadAttachment(task.attachment)));
    }
  });
  function readDraft() {
    const p=period.value, basis=val('tax-basis'), required=quarterly(p)?amounts:amounts.slice(0,2);
    const missing=required.filter(key=>!val('tax-entry-'+key));
    const labels={vat:'增值税',surtax:'附加税费',cit:'企业所得税'};
    if (!p || !basis || missing.length) throw Error('请补齐'+[!p&&'税期',!basis&&'取数及口径',...missing.map(key=>labels[key])].filter(Boolean).join('、'));
    return {period:p,basis,values:Object.fromEntries(required.map(key=>[key,cents(val('tax-entry-'+key))]))};
  }
  window.calculateTax = () => {
    const warning=byId('tax-missing-items');
    try { draft=readDraft(); show(draft); warning.style.display='none'; }
    catch (error) { draft=null; show(null); warning.style.display='block'; byId('tax-missing-text').textContent=error.message; }
  };
  window.fixTaxMissing = () => byId('tax-entry-vat').focus();
  window.saveTaxResult = () => runAction(async () => {
    if (!draft || JSON.stringify(readDraft())!==JSON.stringify(draft)) throw Error('请先核对并更新预览[cite: 1]');
    await saveRecords(STORE_WORK,[{id:'tax-result:'+recordId(),...draft,saved:new Date().toISOString()}],true);
    showToast('测算底稿版本已保存[cite: 1]');
  });
  period.addEventListener('change',loadTaxData); loadTaxData();
}

function initCostModule() {
  const period=byId('closing-period');
  period.value=localDay().slice(0,7);
  const types=[['cogs','销售成本结转'],['depr','固定资产折旧'],['amort','无形资产摊销'],['payroll','工资社保计提'],['split','费用分摊'],['estimate','暂估']];
  const block=byId('cost-cogs').closest('.card-block');
  for (const [key,label] of types.slice(4)) {
    const row=document.createElement('div'); row.className='list-item';
    const title=document.createElement('span'), value=document.createElement('span');
    title.textContent=label; value.id='cost-'+key; value.textContent='待录入';
    row.append(title,value); byId('cost-payroll').closest('.list-item').after(row);
  }
  const accountOptions=VOUCHER_ACCOUNTS.map(row=>[row.code,row.code+' '+row.name]);
  const form=document.createElement('div'); form.className='form-grid';
  const fields=[['作业类型','cost-type','select',types],['来源编号 / 计算与分摊依据','cost-basis'],['本次金额','cost-amount','number'],
    ['借方科目','cost-debit','select',[['','请选择'],...accountOptions]],['贷方科目','cost-credit','select',[['','请选择'],...accountOptions]]];
  fields.forEach(spec=>form.append(makeField(...spec)));
  block.prepend(form);
  const selectedPeriod=()=>period.value;
  form.append(action('生成待审核凭证',async () => {
    const p=selectedPeriod(), type=val('cost-type'), basis=val('cost-basis'), amt=cents(val('cost-amount')), debit=val('cost-debit'),credit=val('cost-credit');
    await ensureOpen(p);
    if (!basis || amt<=0 || !debit || !credit || debit===credit) throw Error('请填写依据、金额及不同的借贷科目[cite: 1]');
    const label=types.find(x=>x[0]===type)[1], amount=amt/100, id='cost:'+p+':'+type+':'+basis;
    await saveRecords(STORE_VOUCHER,[{id,voucherNo:'成本-'+p+'-'+recordId().slice(0,8),period:p,date:lastDay(p),bizLabel:label,
      source:basis,sourceType:'期末作业',status:'PENDING',costType:type,totalDebit:amount,totalCredit:amount,
      entries:[{account:debit,summary:label,debit:amount,credit:0},{account:credit,summary:label,debit:0,credit:amount}]}],true);
    await refreshVouchers(); await calculateCost(); await checkClosing(); showToast('作业凭证已生成，待审核[cite: 1]');
  }, 'btn btn-primary'));
  window.calculateCost = () => runAction(async () => {
    const p=selectedPeriod(), all=await getAllFromStore(STORE_VOUCHER);
    for (const [key] of types) {
      const rows=all.filter(v=>v.period===p && v.costType===key);
      byId('cost-'+key).textContent=rows.length ? fromCents(rows.reduce((n,v)=>n+Math.round(v.totalDebit*100),0)) : '待录入';
    }
  });
  async function closingData(p) {
    const all=await getAllFromStore(STORE_VOUCHER), periodRecords=all.filter(v=>v.period===p);
    const reconciliations=(await getAllFromStore(STORE_WORK)).filter(r=>r.id.startsWith('recon:') && r.period===p);
    const outstanding=periodRecords.filter(v=>v.status!=='POSTED').length;
    return {all,periodRecords,reconciliations,outstanding,ready:reconciliations.length>0 && reconciliations.every(r=>r.verified)};
  }
  window.checkClosing = () => runAction(async () => {
    const p=selectedPeriod(), data=await closingData(p), closed=await work('closed:'+p);
    const transfer=data.periodRecords.find(v=>v.id==='settlement:'+p);
    byId('check-vouchers').textContent=data.outstanding ? data.outstanding+' 张待审核或记账' : '无待审核或记账';
    byId('check-recon').textContent=data.ready ? '已保存流水匹配' : '缺对账记录或存在差异[cite: 1]';
    byId('check-transfer').textContent=transfer ? transfer.status==='POSTED' ? '已记账' : '待审核或记账' : '未生成';
    byId('check-report').textContent=closed?.archive ? '已保存期末快照' : '未归档';
    await calculateCost();
  });
  window.executeSettlement = () => runAction(async () => {
    const p=selectedPeriod(); await ensureOpen(p);
    const data=await closingData(p);
    if (!data.ready || data.outstanding) throw Error('请先保存无差异流水对账，并处理待审核或记账凭证[cite: 1]');
    if (data.periodRecords.some(v=>v.id==='settlement:'+p)) throw Error('该期已有结转凭证[cite: 1]');
    const rows=bookRows(data.all,p,await openingFor(data.all,p));
    const entries=[];
    function move(account,amt,summary) { if (amt) entries.push({account,summary,debit:Math.max(amt,0)/100,credit:Math.max(-amt,0)/100}); }
    for (const row of rows.filter(r=>r.code.startsWith('6'))) move(row.code,-row.end,'损益结转');
    const net=rows.filter(r=>r.code.startsWith('6')).reduce((n,r)=>n+r.end,0);
    move('4103',net,'损益结转');
    if (p.endsWith('-12')) {
      const profit=rows.find(r=>r.code==='4103').end+net;
      move('4103',-profit,'年末利润结转'); move('4104',profit,'年末利润结转');
    }
    if (!entries.length) throw Error('本期无待结转余额[cite: 1]');
    const totalDebit=entries.reduce((sum,e)=>sum+Math.round(e.debit*100),0)/100;
    await saveRecords(STORE_VOUCHER,[{id:'settlement:'+p,voucherNo:'结转-'+p,period:p,date:lastDay(p),bizLabel:'期末结转',
      source:'期末账簿 '+p,status:'PENDING',entries,totalDebit,totalCredit:totalDebit}],true);
    await refreshVouchers(); await checkClosing(); showToast('结转凭证已生成，待审核[cite: 1]');
  });
  window.finalizeClosing = () => runAction(async () => {
    const p=selectedPeriod(); await ensureOpen(p);
    const data=await closingData(p);
    if (!data.ready || data.outstanding || !data.periodRecords.length) throw Error('请处理待审核和待记账事项，并保存无差异对账[cite: 1]');
    const rows=bookRows(data.all,p,await openingFor(data.all,p));
    if (rows.some(r=>r.code.startsWith('6') && r.end) || (p.endsWith('-12') && rows.find(r=>r.code==='4103').end)) throw Error('仍有科目余额待结转[cite: 1]');
    if (rows.reduce((n,r)=>n+r.end,0)!==0) throw Error('期末余额试算不平[cite: 1]');
    if (!(await uiConfirm('确认账户期末余额已人工核对，并锁定本期间？'))) return;
    const records=(await getAllFromStore(STORE_WORK)).filter(r=>r.period===p);
    await lockPeriod(p,'正式结账',{rows,vouchers:data.periodRecords,workpapers:records}); await checkClosing();
    showToast('本期已结账并保存账簿及资料快照[cite: 1]');
  });
  byId('closing-actions').append(action('下载本期归档',async () => {
    const p=selectedPeriod(), record=await work('closed:'+p);
    if (!record?.archive) throw Error('本期间暂无结账归档[cite: 1]');
    downloadFile(p+'-财务归档.json', JSON.stringify(record,null,2),'application/json');
  }, 'btn'));
  period.addEventListener('change',checkClosing);
}
function lastDay(period) {
  const [year,month]=period.split('-').map(Number);
  return period+'-'+String(new Date(year,month,0).getDate()).padStart(2,'0');
}

const draftForms = new Map();
let pendingDraftAttachments = 0;
function readDraftAttachment(reader, file) {
  reader.readAsDataURL(file);
  pendingDraftAttachments++;
  reader.addEventListener('loadend', () => { pendingDraftAttachments--; }, {once:true});
}
function fieldRows(table, prefix, columns) {
  return [...table.rows].map(row => Object.fromEntries(columns.map(name => {
    const el = row.querySelector(prefix + name);
    return [name, el ? el.value : ''];
  })));
}
function registerDraft(module, ids, capture, restore) {
  const read = () => ({
    fields: Object.fromEntries(ids.map(id => [id, byId(id).value])),
    details: capture()
  });
  draftForms.set(module, {
    read,
    restore(draft) {
      for (const id of ids) if (typeof draft.fields[id] === 'string') byId(id).value = draft.fields[id];
      restore(draft.details);
    }
  });
}

async function performBackupExport() {
  const stores = [STORE_NAME, STORE_CASHIER, STORE_VOUCHER, STORE_WORK];
  const tables = Object.fromEntries(await Promise.all(stores.map(async key => [key, await getAllFromStore(key)])));
  downloadFile('finance-backup-'+localDay()+'.json', JSON.stringify({version:2,createdAt:new Date().toISOString(),tables,accounts:ACCOUNTS},null,2),'application/json');
  financeBackupDirty = false;
  const status = byId('workbench-draft-status');
  if (status) status.textContent = '备份文件已生成 · ' + new Date().toLocaleTimeString('zh-CN', {hour:'2-digit',minute:'2-digit'});
  showToast('已导出本机账套备份 (.json)[cite: 1]');
}

async function initBackup() {
  const save = byId('workbench-save'), exportButton = byId('workbench-export'), status = byId('workbench-draft-status');
  const safeExitBtn = byId('workbench-safe-exit'), modal = byId('safe-exit-modal');

  safeExitBtn.addEventListener('click', () => { modal.classList.add('open'); });
  byId('modal-exit-cancel').addEventListener('click', () => { modal.classList.remove('open'); });
  modal.addEventListener('click', (e) => { if (e.target === modal) modal.classList.remove('open'); });
  byId('modal-exit-backup').addEventListener('click', () => runAction(async () => {
    await performBackupExport();
    modal.classList.remove('open');
  }));
  byId('modal-exit-direct').addEventListener('click', () => {
    financeBackupDirty = false;
    modal.classList.remove('open');
    if (status) status.textContent = '已确认直接退出';
    activateWorkspaceTab('home');
    showToast('已安全退出当前编辑，未执行备份[cite: 1]');
  });

  if (storageMode === 'MEMORY') return;
  onAction(save.id, async () => {
    if (pendingDraftAttachments) throw Error('附件仍在读取，请稍后再保存草稿[cite: 1]');
    if (['inv-btn-submit','cash-btn-submit','voucher-btn-submit'].some(id => byId(id).disabled)) {
      throw Error('请等当前提交完成后再保存草稿[cite: 1]');
    }
    const savedAt = new Date().toISOString();
    const drafts = [...draftForms].map(([module, form]) => ({id:'draft:'+module, savedAt, ...form.read()}));
    await saveRecords(STORE_WORK, drafts);
    status.textContent = '草稿已保存 · ' + new Date(savedAt).toLocaleTimeString('zh-CN', {hour:'2-digit',minute:'2-digit'});
    showToast('开票、出纳和制证草稿已保存到本机[cite: 1]');
  });
  onAction(exportButton.id, performBackupExport);
  try {
    const saved = (await getAllFromStore(STORE_WORK)).filter(record => record.id.startsWith('draft:'));
    for (const draft of saved) draftForms.get(draft.id.slice(6))?.restore(draft);
    status.textContent = saved.length ? '已恢复上次保存的草稿' : '';
  } finally {
    save.disabled = false;
    exportButton.disabled = false;
  }
}

/* ================= 5.9 科目余额表（七列实时余额 + 右键穿透 + Excel导入穿透） ================= */
function initDrilldownModule() {
  // 绑定七列余额表的右键与双击事件
  byId('live-balance-body').addEventListener('contextmenu', e => {
    const tr = e.target.closest('tr[data-code]');
    if (!tr) return;
    e.preventDefault();
    openSubjectContextMenu(e.clientX, e.clientY, tr.dataset.code);
  });
  byId('live-balance-body').addEventListener('dblclick', e => {
    const tr = e.target.closest('tr[data-code]');
    if (!tr) return;
    const code = tr.dataset.code;
    if (code === '1002') runAction(openBankSpecialDrawer);
    else if (code === '1122' || code === '2202') runAction(() => openUnitBalanceDrawer(code));
    else runAction(() => openSubjectThreeColumnDrawer(code));
  });
  byId('live-drawer-close').addEventListener('click', closeLiveDrawer);
  byId('live-drawer-mask').addEventListener('click', e => { if (e.target === byId('live-drawer-mask')) closeLiveDrawer(); });

  document.addEventListener('contextmenu', e => { if (ctxOpen && !e.target.closest('#ctx-menu')) hideCtxMenu(); });
  document.addEventListener('click', e => { if (ctxOpen && !e.target.closest('#ctx-menu')) hideCtxMenu(); }, true);
  window.addEventListener('scroll', () => { if (ctxOpen) hideCtxMenu(); }, true);
  window.addEventListener('resize', () => { if (ctxOpen) hideCtxMenu(); });
  document.addEventListener('keydown', e => { if (e.key === 'Escape') { hideCtxMenu(); closeLiveDrawer(); } });

  byId('btn-export-live-balance').addEventListener('click', () => runAction(async () => {
    const { list } = await computeLiveBalances();
    const rows = list.map(m => [m.code, m.name, m.closeDir, m.opening.toFixed(2), m.debit.toFixed(2), m.credit.toFixed(2), m.closing.toFixed(2)]);
    exportRows('科目余额表_2026-01', ['科目编码', '科目名称', '借/贷方向', '期初余额', '本期借方发生额', '本期贷方发生额', '期末余额'], rows);
    showToast('科目余额表已导出');
  }));

  byId('btn-reset-demo-vouchers').addEventListener('click', () => runAction(async () => {
    if (!(await uiConfirm('确认将凭证库重置为初始演示账套？当前新录入的凭证将被清除。'))) return;
    await clearStore(STORE_VOUCHER);
    await saveRecords(STORE_VOUCHER, SEED_VOUCHERS);
    await renderLiveBalance();
    if (window.refreshVouchers) await refreshVouchers();
    if (window.refreshAccounting) await refreshAccounting();
    if (window.renderARSubledgerModule) await renderARSubledgerModule();
    showToast('已重置为演示账套数据');
  }));

  // Excel 导入穿透逻辑（保留原代码2）
  let balances=[], journal=[], selectedAccount=null, ledgerRows=[], reportType='profit', currentReport=[];
  const norm=s=>String(s??'').replace(/[\s\u3000\r\n]/g,'').replace(/[（）]/g,m=>m==='（'?'(' : ')').toLowerCase();
  const aliases={
    code:['科目编码','科目代码','科目编号','会计科目编码'], name:['科目名称','会计科目名称','会计科目','科目'],
    parent:['上级编码','父科目编码','上级科目代码'], level:['级次','科目级别','级别'],
    opening:['期初余额','年初余额'], opD:['期初借方余额','期初借方'], opC:['期初贷方余额','期初贷方'],
    debit:['本期借方','本期借方发生额','借方发生额'], credit:['本期贷方','本期贷方发生额','贷方发生额'],
    ending:['期末余额'], endD:['期末借方余额','期末借方'], endC:['期末贷方余额','期末贷方'],
    direction:['余额方向','期末方向','科目方向'],
    date:['凭证日期','日期','记账日期','业务日期'], voucher:['凭证号','凭证编号','字号'], summary:['摘要','分录摘要','业务说明'],
    jcode:['科目编码','科目代码','会计科目编码'], jname:['科目名称','会计科目名称'], jdebit:['借方金额','借方','借方发生额'], jcredit:['贷方金额','贷方','贷方发生额'],
    cashItem:['现金流量项目','流量项目编码'], bank:['银行账户','银行账号','账户编码']
  };
  const pick=(row,key)=>{const keys=Object.keys(row);const hit=keys.find(k=>aliases[key]?.some(a=>norm(k)===norm(a)));return hit?row[hit]:'';};
  const amount=v=>{if(v===null||v===undefined||String(v).trim()==='')return 0;let s=String(v).trim().replace(/[，,￥¥\s]/g,'');if(/^\(.*\)$/.test(s))s='-'+s.slice(1,-1);const n=Number(s);if(!Number.isFinite(n))throw Error('金额无法解析：'+v);return Math.round(n*100)/100;};
  async function readSheet(file,type) {
    if(!file) return [];
    if(!window.XLSX) throw Error('Excel 解析组件未加载，请联网打开页面后重试[cite: 1]');
    const data=XLSX.read(await file.arrayBuffer(),{type:'array',cellDates:true});
    let candidates=[];
    for(const sheetName of data.SheetNames){
      const rows=XLSX.utils.sheet_to_json(data.Sheets[sheetName],{defval:'',raw:false});
      if(rows.length) candidates.push({sheetName,rows});
    }
    if(!candidates.length) throw Error(file.name+' 中没有可识别数据[cite: 1]');
    const required=type==='balance'?['code','name','debit','credit']:['date','voucher','jcode','jdebit','jcredit'];
    let found=candidates.find(c=>required.every(k=>Object.keys(c.rows[0]).some(h=>aliases[k].some(a=>norm(h)===norm(a)))));
    if(!found) throw Error(file.name+' 未识别到必需字段：'+required.map(k=>aliases[k][0]).join('、')+'。请检查表头是否为第一行[cite: 1]。');
    return found.rows;
  }
  function parseBalance(rows){
    return rows.map((r,i)=>{
      const code=String(pick(r,'code')??'').trim(), name=String(pick(r,'name')??'').trim();
      if(!code||!name||/合计|总计/.test(name))return null;
      let opening,ending;
      if(pick(r,'opD')!==''||pick(r,'opC')!=='') opening=amount(pick(r,'opD'))-amount(pick(r,'opC'));
      else opening=amount(pick(r,'opening'));
      if(pick(r,'endD')!==''||pick(r,'endC')!=='') ending=amount(pick(r,'endD'))-amount(pick(r,'endC'));
      else {ending=amount(pick(r,'ending'));const dir=String(pick(r,'direction'));if(/贷/.test(dir))ending=-Math.abs(ending);}
      const direction=String(pick(r,'direction')||'');
      return {code,name,parent:String(pick(r,'parent')||''),level:Number(pick(r,'level'))||null,opening,debit:amount(pick(r,'debit')),credit:amount(pick(r,'credit')),ending,direction,sourceRow:i+2};
    }).filter(Boolean);
  }
  function parseJournal(rows){
    return rows.map((r,i)=>({date:String(pick(r,'date')||'').slice(0,10),voucher:String(pick(r,'voucher')||''),code:String(pick(r,'jcode')||'').trim(),
      name:String(pick(r,'jname')||''),summary:String(pick(r,'summary')||''),debit:amount(pick(r,'jdebit')),credit:amount(pick(r,'jcredit')),
      cashItem:String(pick(r,'cashItem')||''),bank:String(pick(r,'bank')||''),sourceRow:i+2})).filter(x=>x.code);
  }
  function treeRows(){
    const byCode=new Map(balances.map(x=>[x.code,{...x,children:[]}]));
    const roots=[];
    for(const node of byCode.values()){
      let parent=node.parent&&byCode.get(node.parent);
      if(!parent) parent=[...byCode.values()].filter(p=>p.code!==node.code&&node.code.startsWith(p.code)&&node.code.length>p.code.length).sort((a,b)=>b.code.length-a.code.length)[0];
      if(parent) parent.children.push(node); else roots.push(node);
    }
    const flat=[];const walk=(n,depth)=>{flat.push({...n,depth});n.children.sort((a,b)=>a.code.localeCompare(b.code,'zh'));n.children.forEach(c=>walk(c,depth+1));};
    roots.sort((a,b)=>a.code.localeCompare(b.code,'zh')).forEach(n=>walk(n,0));return flat;
  }
  const expanded=new Set();
  function renderTree(){
    const body=byId('drill-tree-body'), q=norm(byId('drill-search').value);body.replaceChildren();
    if(!balances.length){body.innerHTML='<tr><td colspan="8">请导入科目余额表[cite: 1]</td></tr>';return;}
    const flat=treeRows(), visible=flat.filter(n=>(!q||norm(n.code+n.name).includes(q))&&flat.slice(0,flat.indexOf(n)).every(p=>!(n.code.startsWith(p.code)&&p.code!==n.code&&p.depth<n.depth&&!expanded.has(p.code))));
    visible.forEach(n=>{
      const tr=body.insertRow();tr.dataset.code=n.code;
      const cells=[n.code,n.name,n.opening,n.debit,n.credit,n.ending,n.ending>0?'借':n.ending<0?'贷':'平',''];
      cells.forEach((v,i)=>{const td=tr.insertCell();if(i===0){td.style.paddingLeft=(12+n.depth*18)+'px';td.textContent=(n.children.length?(expanded.has(n.code)?'▾ ':'▸ '):'')+v;}else if(i>=2&&i<=5)td.textContent=Number(v).toLocaleString('zh-CN',{minimumFractionDigits:2,maximumFractionDigits:2});else td.textContent=v;});
      tr.cells[0].style.cursor='pointer';tr.cells[0].onclick=()=>{if(n.children.length){expanded.has(n.code)?expanded.delete(n.code):expanded.add(n.code);renderTree();}};
      const btn=document.createElement('button');btn.type='button';btn.className='btn-mini';btn.textContent='⋯ 联查';btn.onclick=()=>openLedger(n);tr.cells[7].append(btn);
      tr.oncontextmenu=e=>{e.preventDefault();openLedger(n);};
    });
  }
  function openLedger(account){
    selectedAccount=account;const scope=balances.filter(b=>b.code===account.code||b.code.startsWith(account.code));
    const codes=new Set(scope.map(x=>x.code));
    ledgerRows=journal.filter(j=>codes.has(j.code)).sort((a,b)=>a.date.localeCompare(b.date)||a.voucher.localeCompare(b.voucher,'zh',{numeric:true})||a.sourceRow-b.sourceRow);
    byId('drill-drawer-title').textContent=account.code+' '+account.name;
    byId('drill-drawer-meta').textContent=`范围：${account.code} 及已识别下级 · 期间：${byId('drill-period').value||'来源期间未指定'} · 余额来源：导入余额表；凭证来源：${journal.length?'已导入序时账':'未导入序时账'}[cite: 1]`;
    byId('drill-drawer').style.display='block';renderLedger();
  }
  function renderLedger(){
    const body=byId('drill-ledger-body');body.replaceChildren();
    if(!journal.length){body.innerHTML='<tr><td colspan="8">已保留科目联查入口；请补充序时账后查看凭证明细[cite: 1]。</td></tr>';return;}
    const from=byId('drill-date-from').value,to=byId('drill-date-to').value,q=norm(byId('drill-keyword').value);
    let bal=selectedAccount?.opening||0, rows=ledgerRows.map(r=>{bal+=r.debit-r.credit;return {...r,balance:bal};});
    const hit=rows.filter(r=>(!from||r.date>=from)&&(!to||r.date<=to)&&(!q||norm(r.voucher+r.summary+r.code+r.name+r.bank).includes(q)));
    hit.forEach(r=>{const tr=body.insertRow();[r.date,r.voucher,r.code,r.summary,r.debit.toFixed(2),r.credit.toFixed(2),r.balance>0?'借':r.balance<0?'贷':'平',Math.abs(r.balance).toFixed(2)].forEach(v=>tr.insertCell().textContent=v);});
    byId('drill-ledger-summary').textContent=`命中 ${hit.length} 条；借方合计 ${hit.reduce((s,r)=>s+r.debit,0).toFixed(2)}，贷方合计 ${hit.reduce((s,r)=>s+r.credit,0).toFixed(2)}[cite: 1]。`;
  }
  function buildReport(){
    const p=byId('drill-period').value, rows=journal.filter(r=>!p||r.date.startsWith(p));
    const report=reportType==='profit'?[
      ['营业收入',codes=>codes.startsWith('6001')||codes.startsWith('6051'),'贷方−借方'],
      ['营业成本',codes=>codes.startsWith('6401'),'借方−贷方'],
      ['税金及附加',codes=>codes.startsWith('6403'),'借方−贷方'],
      ['销售费用',codes=>codes.startsWith('6601'),'借方−贷方'],
      ['管理费用',codes=>codes.startsWith('6602'),'借方−贷方'],
      ['研发费用',codes=>codes.startsWith('6603'),'借方−贷方'],
      ['财务费用',codes=>codes.startsWith('6604'),'借方−贷方'],
      ['营业外收入',codes=>codes.startsWith('6301'),'贷方−借方'],
      ['营业外支出',codes=>codes.startsWith('6711'),'借方−贷方'],
      ['所得税费用',codes=>codes.startsWith('6801'),'借方−贷方']
    ]:[['经营活动现金流量（已标记项目）',(_,r)=>!!r.cashItem&&/经营/.test(r.cashItem),'按现金流项目字段'],['投资活动现金流量（已标记项目）',(_,r)=>!!r.cashItem&&/投资/.test(r.cashItem),'按现金流项目字段'],['筹资活动现金流量（已标记项目）',(_,r)=>!!r.cashItem&&/筹资/.test(r.cashItem),'按现金流项目字段'],['待分类／待核实',(_,r)=>!r.cashItem,'未提供现金流项目的现金分录暂无法归类']];
    const body=byId('drill-report-body');body.replaceChildren();currentReport=[];
    if(!journal.length){body.innerHTML='<tr><td colspan="3">尚未导入序时账，报表待补全[cite: 1]。</td></tr>';return;}
    report.forEach(([label,test,rule])=>{
      const selected=reportType==='profit'?rows.filter(r=>test(r.code)):rows.filter(r=>test(r.code,r));
      const value=selected.reduce((s,r)=>s+(reportType==='profit'?(label==='营业收入'||label==='营业外收入'?r.credit-r.debit:r.debit-r.credit):(r.debit-r.credit)),0);
      currentReport.push({项目:label,金额:Number(value.toFixed(2)),取数说明:rule});
      const tr=body.insertRow();[label,value.toFixed(2),rule+(selected.length?' · '+selected.length+'条分录':' · 无匹配分录')].forEach(v=>tr.insertCell().textContent=v);
    });
  }
  byId('drill-import').addEventListener('click',()=>runAction(async()=>{
    const bf=byId('drill-balance-file').files[0], jf=byId('drill-journal-file').files[0];if(!bf)throw Error('请先选择科目余额表[cite: 1]');
    const bRows=await readSheet(bf,'balance'), jRows=jf?await readSheet(jf,'journal'):[];
    const nextBalances=parseBalance(bRows), nextJournal=parseJournal(jRows);if(!nextBalances.length)throw Error('余额表未识别到科目记录[cite: 1]');
    balances=nextBalances;journal=nextJournal;expanded.clear();renderTree();
    byId('drill-source-status').textContent=`余额表：${bf.name}（${balances.length} 个科目）；序时账：${jf?jf.name+'（'+journal.length+' 条）':'未导入'}[cite: 1]。`;
    buildReport();
  }));
  byId('drill-clear').onclick=()=>{balances=[];journal=[];renderTree();byId('drill-source-status').textContent='已清空导入数据[cite: 1]';byId('drill-report-body').innerHTML='<tr><td colspan="3">请选择报表并生成[cite: 1]</td></tr>';};
  byId('drill-search').addEventListener('input',renderTree);
  byId('drill-expand').onclick=()=>{treeRows().forEach(n=>expanded.add(n.code));renderTree();};
  byId('drill-collapse').onclick=()=>{expanded.clear();renderTree();};
  byId('drill-drawer-close').onclick=()=>byId('drill-drawer').style.display='none';
  byId('drill-drawer').addEventListener('click',e=>{if(e.target.id==='drill-drawer')byId('drill-drawer').style.display='none';});
  byId('drill-apply-filter').onclick=renderLedger;
  byId('drill-build-report').onclick=buildReport;
  document.querySelectorAll('[data-drill-report]').forEach(b=>b.onclick=()=>{reportType=b.dataset.drillReport;document.querySelectorAll('[data-drill-report]').forEach(x=>x.classList.toggle('active',x===b));buildReport();});
  byId('drill-export-ledger').onclick=()=>{if(!window.XLSX)return showToast('Excel 组件未加载');const from=byId('drill-date-from').value,to=byId('drill-date-to').value,q=norm(byId('drill-keyword').value);const data=ledgerRows.filter(r=>(!from||r.date>=from)&&(!to||r.date<=to)&&(!q||norm(r.voucher+r.summary+r.code).includes(q)));const ws=XLSX.utils.json_to_sheet(data);const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'明细账');XLSX.writeFile(wb,'科目穿透明细.xlsx');};
  byId('drill-export-report').onclick=()=>{if(!currentReport.length)buildReport();const ws=XLSX.utils.json_to_sheet(currentReport);const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'报表');XLSX.writeFile(wb,(reportType==='profit'?'利润表':'现金流量表')+'.xlsx');};
  byId('drill-period').value=localDay().slice(0,7);

  renderLiveBalance();
}

/* ================= 关页前自动备份提醒 ================= */
let financeBackupDirty = false;
function markFinanceDirty() {
  financeBackupDirty = true;
  const s = byId('workbench-draft-status');
  if (s) s.textContent = '有变更尚未备份 · 请点击「安全退出 / 备份」[cite: 1]';
}
window.addEventListener('beforeunload', (e) => {
  if (financeBackupDirty) { e.preventDefault(); e.returnValue = ''; }
});
document.addEventListener('input', (e) => {
  if (e.target.closest('.module') && !e.target.matches('input[type="file"]')) markFinanceDirty();
}, true);
document.addEventListener('change', (e) => {
  if (e.target.closest('.module')) markFinanceDirty();
}, true);
document.addEventListener('click', (e) => {
  const b = e.target.closest('button');
  if (b && b.closest('.module') && !b.id.startsWith('drill-export') && !b.id.startsWith('btn-export') && !b.classList.contains('btn-enter')) markFinanceDirty();
}, true);
</script>
</body>
</html>
