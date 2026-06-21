"use strict";
/* Liquidity Tracker — append-only event-log sync.
   Every edit is appended as a timestamped event to the AppLog table in your OneDrive workbook.
   Current state is rebuilt by replaying events in TIMESTAMP order, so edits from your Mac and
   phone never conflict and arrival order does not matter. The AppData sheet is a derived snapshot
   the app refreshes for the human-readable Excel summary.
   Talks only to your own OneDrive via Microsoft Graph. No server, no client secret. */

const MONTHS = ["Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec","Jan","Feb","Mar"];
const SCOPES = ["User.Read", "Files.ReadWrite"];
const fmt = v => "₹" + Math.round(v || 0).toLocaleString("en-IN");
const $ = id => document.getElementById(id);

let pca = null, account = null;
let state = [];                 // [{id,item,cat,amount[12],paid[12],exists}]
let byId = new Map();
let curMonth = ((new Date().getMonth()) - 3 + 12) % 12;  // FY index, Apr=0
let editing = null, freq = "month";
let maxRowsSeen = 0, evCounter = 0, loading = false;
let chain = Promise.resolve();   // serialises writes so they never overlap

let deviceId = localStorage.getItem("liq_device");
if(!deviceId){
  deviceId = (window.crypto && crypto.randomUUID) ? crypto.randomUUID() : ("dev-" + Date.now() + "-" + Math.random().toString(36).slice(2));
  localStorage.setItem("liq_device", deviceId);
}

function cfgOk(){ return window.APP_CONFIG && APP_CONFIG.clientId && APP_CONFIG.clientId.indexOf("PASTE") < 0; }

function initMsal(){
  pca = new msal.PublicClientApplication({
    auth: { clientId: APP_CONFIG.clientId, authority: "https://login.microsoftonline.com/common",
            redirectUri: (APP_CONFIG.redirectUri || window.location.href.split("#")[0].split("?")[0]) },
    cache: { cacheLocation: "localStorage", storeAuthStateInCookie: false },
    system: { loggerOptions: { piiLoggingEnabled: false } }
  });
}
async function ensureToken(){
  try { const r = await pca.acquireTokenSilent({ scopes: SCOPES, account }); return r.accessToken; }
  catch(e){ const r = await pca.acquireTokenPopup({ scopes: SCOPES }); account = r.account; return r.accessToken; }
}
function getFilePath(){ return (localStorage.getItem("liq_filepath") || "").replace(/^\/+/, "").trim(); }
function gUrl(suffix){
  const path = getFilePath().split("/").map(encodeURIComponent).join("/");
  return `https://graph.microsoft.com/v1.0/me/drive/root:/${path}:/workbook/${suffix}`;
}
async function graph(suffix, method = "GET", body){
  const tok = await ensureToken();
  const res = await fetch(gUrl(suffix), { method,
    headers: { Authorization: "Bearer " + tok, "Content-Type": "application/json" },
    body: body ? JSON.stringify(body) : undefined });
  if(!res.ok){ let d=""; try{ const j=await res.json(); d=(j.error&&j.error.message)||""; }catch(_){} throw new Error(res.status + (d?" — "+d:"")); }
  return res.status === 204 ? null : res.json();
}

/* ---------- event log ---------- */
function mkEvent(itemId, field, mi, value){
  return ["ev-" + deviceId.slice(0,8) + "-" + Date.now() + "-" + (evCounter++), new Date().toISOString(), deviceId, itemId, field, mi, value];
}
async function readEvents(){
  let r;
  try { r = await graph("tables('tblLog')/dataBodyRange"); }
  catch(e){ if(String(e.message).indexOf("EmptyTable") >= 0 || String(e.message).indexOf("404") >= 0) return []; throw e; }
  const vals = (r && r.values) || [];
  return vals.filter(row => row && row[0] !== null && row[0] !== "").map(row => ({
    eid: String(row[0]), ts: String(row[1]), itemId: String(row[3]),
    field: String(row[4]), mi: Number(row[5]), value: row[6]
  }));
}
async function appendEvents(rows){ if(rows.length) await graph("tables('tblLog')/rows", "POST", { values: rows }); }

function replay(events){
  events.sort((a,b) => a.ts < b.ts ? -1 : a.ts > b.ts ? 1 : (a.eid < b.eid ? -1 : a.eid > b.eid ? 1 : 0));
  const map = new Map();
  const TRUE = v => String(v).toUpperCase() === "TRUE" || v === true;
  for(const e of events){
    let r = map.get(e.itemId);
    if(!r){ r = { id: e.itemId, item: "", cat: "", amount: Array(12).fill(0), paid: Array(12).fill(false), exists: true }; map.set(e.itemId, r); }
    if(e.field === "exists") r.exists = TRUE(e.value);
    else if(e.field === "item") r.item = String(e.value);
    else if(e.field === "cat") r.cat = String(e.value);
    else if(e.field === "amount"){ if(e.mi >= 0 && e.mi < 12) r.amount[e.mi] = Number(e.value) || 0; }
    else if(e.field === "paid"){ if(e.mi >= 0 && e.mi < 12) r.paid[e.mi] = TRUE(e.value); }
  }
  return [...map.values()].filter(r => r.exists);
}
function reindex(){ byId = new Map(state.map(r => [r.id, r])); }

/* ---------- derived snapshot for Excel (best-effort) ---------- */
async function materialize(){
  try{
    const rows = [];
    state.forEach(r => { for(let mi=0; mi<12; mi++){ if(r.amount[mi] > 0) rows.push([r.id, r.item, r.cat, mi, MONTHS[mi], r.amount[mi], r.paid[mi]]); } });
    const n = rows.length, end = n + 1;
    if(n > 0) await graph(`worksheets('AppData')/range(address='A2:G${end}')`, "PATCH", { values: rows });
    if(maxRowsSeen > n) await graph(`worksheets('AppData')/range(address='A${end+1}:G${maxRowsSeen+1}')/clear`, "POST", { applyTo: "Contents" });
    maxRowsSeen = Math.max(n, 0);
  }catch(e){ /* derived view only; log remains source of truth */ }
}

function enqueue(fn){ chain = chain.then(fn).catch(e => setStatus("Sync error: " + e.message, false, true)); return chain; }

async function load(){
  if(loading) return;
  if(!getFilePath()){ openSettings(true); return; }
  loading = true;
  try{
    setStatus("Syncing from OneDrive…", true);
    const events = await readEvents();
    state = replay(events); reindex(); render();
    setStatus("Synced ✓ " + new Date().toLocaleTimeString());
    enqueue(materialize);
  }catch(e){ setStatus("Couldn't sync: " + e.message, false, true); }
  finally{ loading = false; }
}

/* ---------- rendering ---------- */
const monthItems = m => state.filter(r => r.amount[m] > 0).map(r => ({ rec: r, mi: m, amount: r.amount[m], paid: r.paid[m] }));
const monthTotal = m => state.reduce((s, r) => s + (r.amount[m] || 0), 0);

function render(){
  const m = curMonth; $("msel").value = m;
  const due = monthItems(m).sort((a,b) => b.amount - a.amount);
  const el = $("rows"); el.innerHTML = "";
  let total = 0, paid = 0;
  due.forEach(d => {
    total += d.amount; if(d.paid) paid += d.amount;
    const row = document.createElement("div"); row.className = "row" + (d.paid ? " paid" : "");
    const tick = document.createElement("div"); tick.className = "tick"; tick.textContent = d.paid ? "✓" : "";
    tick.onclick = () => togglePaid(d.rec, m);
    const name = document.createElement("div"); name.className = "name";
    const n = document.createElement("div"); n.className = "n"; n.textContent = d.rec.item;
    const c = document.createElement("div"); c.className = "c"; c.textContent = d.rec.cat;
    name.appendChild(n); name.appendChild(c); name.onclick = () => togglePaid(d.rec, m);
    const amt = document.createElement("div"); amt.className = "amt"; amt.textContent = fmt(d.amount);
    const edit = document.createElement("button"); edit.className = "edit"; edit.textContent = "✎"; edit.setAttribute("aria-label", "Edit");
    edit.onclick = e => { e.stopPropagation(); openEdit(d.rec, m); };
    row.appendChild(tick); row.appendChild(name); row.appendChild(amt); row.appendChild(edit);
    el.appendChild(row);
  });
  if(!due.length){ const e = document.createElement("div"); e.className = "status"; e.textContent = "No payments due this month."; el.appendChild(e); }
  $("mNeed").textContent = fmt(total); $("mPaid").textContent = fmt(paid); $("mLeft").textContent = fmt(total - paid);
  drawBars();
}
function drawBars(){
  const totals = MONTHS.map((_, i) => monthTotal(i)), max = Math.max(1, ...totals);
  const el = $("bars"); el.innerHTML = "";
  totals.forEach((t, i) => {
    const b = document.createElement("div"); b.className = "b" + (i === curMonth ? " cur" : "");
    const bar = document.createElement("div"); bar.className = "bar"; bar.style.height = Math.round(t / max * 70) + "px";
    const ml = document.createElement("div"); ml.className = "ml"; ml.textContent = MONTHS[i];
    b.appendChild(bar); b.appendChild(ml); b.onclick = () => { curMonth = i; render(); }; el.appendChild(b);
  });
}

/* ---------- edits → events ---------- */
function togglePaid(rec, mi){
  rec.paid[mi] = !rec.paid[mi]; render();
  const val = rec.paid[mi] ? "TRUE" : "FALSE";
  enqueue(async () => { await appendEvents([mkEvent(rec.id, "paid", mi, val)]); await materialize(); setStatus("Saved ✓ " + new Date().toLocaleTimeString()); });
}
function openEdit(rec, mi){
  editing = { rec, mi }; $("sheetTitle").textContent = "Edit expense";
  $("fName").value = rec.item; $("fCat").value = rec.cat || "Other"; $("fAmt").value = rec.amount[mi];
  $("freqField").style.display = "none"; $("delBtn").classList.remove("hide"); $("sheet").classList.add("open");
}
function openAdd(){
  editing = null; $("sheetTitle").textContent = "Add expense";
  $("fName").value = ""; $("fCat").value = "Other"; $("fAmt").value = "";
  setFreq("month"); $("freqField").style.display = ""; $("delBtn").classList.add("hide"); $("sheet").classList.add("open");
}
function setFreq(f){ freq = f; document.querySelectorAll("#fFreq button").forEach(b => b.classList.toggle("on", b.dataset.f === f)); }

function commitSheet(){
  const name = $("fName").value.trim(), cat = $("fCat").value, amt = Math.max(0, Math.round(Number($("fAmt").value) || 0));
  if(!name){ alert("Please enter a name."); return; }
  $("sheet").classList.remove("open");
  if(editing){
    const rec = editing.rec, mi = editing.mi, evs = [];
    if(name !== rec.item){ rec.item = name; evs.push(mkEvent(rec.id, "item", -1, name)); }
    if(cat !== rec.cat){ rec.cat = cat; evs.push(mkEvent(rec.id, "cat", -1, cat)); }
    if(amt !== rec.amount[mi]){ rec.amount[mi] = amt; evs.push(mkEvent(rec.id, "amount", mi, amt)); }
    render();
    if(evs.length) enqueue(async () => { await appendEvents(evs); await materialize(); setStatus("Saved ✓ " + new Date().toLocaleTimeString()); });
  } else {
    const id = "usr-" + Date.now().toString(36) + "-" + Math.random().toString(36).slice(2,6);
    const rec = { id, item: name, cat, amount: Array(12).fill(0), paid: Array(12).fill(false), exists: true };
    const evs = [mkEvent(id, "exists", -1, "TRUE"), mkEvent(id, "item", -1, name), mkEvent(id, "cat", -1, cat)];
    for(let mi=0; mi<12; mi++){ const a = (freq === "all") ? amt : (mi === curMonth ? amt : 0); rec.amount[mi] = a; if(a > 0) evs.push(mkEvent(id, "amount", mi, a)); }
    state.push(rec); byId.set(id, rec); render();
    enqueue(async () => { await appendEvents(evs); await materialize(); setStatus("Saved ✓ " + new Date().toLocaleTimeString()); });
  }
}
function deleteItem(){
  if(!editing) return; const rec = editing.rec;
  if(!confirm("Delete “" + rec.item + "” from all months?")) return;
  $("sheet").classList.remove("open");
  state = state.filter(r => r.id !== rec.id); byId.delete(rec.id); render();
  enqueue(async () => { await appendEvents([mkEvent(rec.id, "exists", -1, "FALSE")]); await materialize(); setStatus("Saved ✓ " + new Date().toLocaleTimeString()); });
}

function setStatus(t, isBusy, isErr){
  const s = $("status"); s.textContent = "";
  if(isBusy){ const sp = document.createElement("span"); sp.className = "spin"; s.appendChild(sp); s.appendChild(document.createTextNode(" ")); }
  s.appendChild(document.createTextNode(t)); s.className = "status" + (isErr ? " err" : "");
}
function showApp(on){ $("gate").classList.toggle("hide", on); $("app").classList.toggle("hide", !on); $("fab").classList.toggle("hide", !on); $("signBtn").textContent = on ? "Sign out" : "Sign in"; }

async function signIn(){
  try{ const r = await pca.loginPopup({ scopes: SCOPES, prompt: "select_account" }); account = r.account; $("hsub").textContent = account.username; showApp(true); await load(); }
  catch(e){ setStatus("Sign-in failed: " + e.message, false, true); }
}
function signOut(){ try{ pca.logoutPopup({ account }); }catch(_){} account = null; showApp(false); }

/* ---------- OneDrive folder picker ---------- */
let pickerPath = "";
const encPath = p => p.split("/").filter(Boolean).map(encodeURIComponent).join("/");

async function graphApi(pathAndQuery){
  const tok = await ensureToken();
  const res = await fetch("https://graph.microsoft.com/v1.0" + pathAndQuery, { headers: { Authorization: "Bearer " + tok } });
  if(!res.ok){ let d=""; try{ const j=await res.json(); d=(j.error&&j.error.message)||""; }catch(_){} throw new Error(res.status + (d?" — "+d:"")); }
  return res.json();
}

function openSettings(firstRun){
  $("set2").classList.add("open");
  if(!account){ $("pickerMsg").textContent = "Please sign in first."; $("pickerList").innerHTML = ""; $("pickerPath").textContent = "/"; return; }
  if(firstRun) setStatus("Choose your workbook to continue.", false);
  const existing = getFilePath();
  const startDir = existing.includes("/") ? existing.split("/").slice(0, -1).join("/") : "";
  browse(startDir);
}

async function browse(path){
  pickerPath = path || "";
  $("pickerPath").textContent = "/" + pickerPath;
  $("pickerUp").disabled = !pickerPath;
  $("pickerMsg").textContent = "Loading…";
  const list = $("pickerList"); list.innerHTML = "";
  try{
    const data = await listChildren(pickerPath);
    const items = data.value || [];
    const folders = items.filter(i => i.folder).sort((a,b) => a.name.localeCompare(b.name));
    const wbCount = items.filter(i => i.file && /\.xls[xm]$/i.test(i.name)).length;
    folders.forEach(f => list.appendChild(folderRow(f.name, () => browse((pickerPath ? pickerPath + "/" : "") + f.name))));
    if(!folders.length){ const e = document.createElement("div"); e.className = "status"; e.textContent = "No subfolders here."; list.appendChild(e); }
    $("pickerMsg").textContent = wbCount
      ? (wbCount + " Excel workbook" + (wbCount > 1 ? "s" : "") + " in this folder — tap “Use this folder”.")
      : "No Excel workbook here — open the folder that contains it.";
  }catch(e){ $("pickerMsg").textContent = "Couldn't open folder: " + e.message; }
}

function listChildren(path){
  const base = path ? `/me/drive/root:/${encPath(path)}:/children` : "/me/drive/root/children";
  return graphApi(base + "?$select=name,id,folder,file&$top=200");
}

function folderRow(name, onClick){
  const row = document.createElement("div");
  row.style.cssText = "display:flex; align-items:center; gap:10px; padding:11px 12px; border-bottom:1px solid var(--line); cursor:pointer;";
  row.onclick = onClick;
  const ic = document.createElement("span"); ic.textContent = "📁"; ic.style.fontSize = "16px"; ic.setAttribute("aria-hidden", "true");
  const nm = document.createElement("span"); nm.textContent = name;
  nm.style.cssText = "flex:1; font-size:14px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;";
  const ch = document.createElement("span"); ch.textContent = "›"; ch.style.cssText = "font-size:16px; color:var(--muted);";
  row.appendChild(ic); row.appendChild(nm); row.appendChild(ch);
  return row;
}

async function useFolder(){
  $("pickerMsg").textContent = "Looking for your workbook…";
  try{
    const data = await listChildren(pickerPath);
    const xlsx = (data.value || []).filter(i => i.file && /\.xls[xm]$/i.test(i.name)).sort((a,b) => a.name.localeCompare(b.name));
    if(xlsx.length === 0){ $("pickerMsg").textContent = "No Excel workbook in this folder. Open the folder that contains it."; return; }
    if(xlsx.length === 1){ saveWorkbook(pickerPath, xlsx[0].name); return; }
    chooseWorkbook(xlsx);  // rare: more than one workbook → confirm which once
  }catch(e){ $("pickerMsg").textContent = "Couldn't read folder: " + e.message; }
}

function chooseWorkbook(files){
  const list = $("pickerList"); list.innerHTML = "";
  $("pickerMsg").textContent = "This folder has several workbooks — tap the one the tracker uses:";
  const cur = getFilePath();
  files.forEach(f => {
    const full = (pickerPath ? pickerPath + "/" : "") + f.name;
    const row = document.createElement("div");
    row.style.cssText = "display:flex; align-items:center; gap:10px; padding:11px 12px; border-bottom:1px solid var(--line); cursor:pointer;";
    row.onclick = () => saveWorkbook(pickerPath, f.name);
    const ic = document.createElement("span"); ic.textContent = "📄"; ic.style.fontSize = "16px"; ic.setAttribute("aria-hidden", "true");
    const nm = document.createElement("span"); nm.textContent = f.name;
    nm.style.cssText = "flex:1; font-size:14px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;";
    const tag = document.createElement("span"); tag.textContent = full === cur ? "current ✓" : "use";
    tag.style.cssText = "font-size:11px; color:" + (full === cur ? "var(--green)" : "var(--blue)") + ";";
    row.appendChild(ic); row.appendChild(nm); row.appendChild(tag);
    list.appendChild(row);
  });
}

function saveWorkbook(folder, file){
  const full = ((folder ? folder + "/" : "") + file).replace(/^\/+/, "");
  localStorage.setItem("liq_filepath", full);
  $("set2").classList.remove("open");
  if(account) load();
}
function pickerUp(){ if(pickerPath) browse(pickerPath.split("/").slice(0, -1).join("/")); }

function buildMonthSel(){
  const s = $("msel"); s.innerHTML = "";
  MONTHS.forEach((m, i) => { const o = document.createElement("option"); o.value = i; o.textContent = m + (i >= 9 ? " 2027" : " 2026"); s.appendChild(o); });
  s.onchange = () => { curMonth = +s.value; render(); };
}

window.addEventListener("DOMContentLoaded", () => {
  buildMonthSel();
  if(!cfgOk()){ $("cfgwarn").textContent = "Setup needed: open config.js and paste your Microsoft Client ID (see SETUP_GUIDE)."; return; }
  initMsal();
  const accts = pca.getAllAccounts();
  if(accts.length){ account = accts[0]; $("hsub").textContent = account.username; showApp(true); load(); }
  $("signBtn").onclick = () => account ? signOut() : signIn();
  $("signBtn2").onclick = signIn;
  $("refreshBtn").onclick = () => load();
  $("addBtn").onclick = openAdd;
  $("cancelBtn").onclick = () => $("sheet").classList.remove("open");
  $("saveBtn").onclick = commitSheet;
  $("delBtn").onclick = deleteItem;
  $("setBtn").onclick = () => openSettings(false);
  $("setCancel").onclick = () => $("set2").classList.remove("open");
  $("setSignOut").onclick = () => { $("set2").classList.remove("open"); signOut(); };
  $("pickerUp").onclick = pickerUp;
  $("pickerReload").onclick = () => browse(pickerPath);
  $("useFolderBtn").onclick = useFolder;
  document.querySelectorAll("#fFreq button").forEach(b => b.onclick = () => setFreq(b.dataset.f));
});

if("serviceWorker" in navigator){ navigator.serviceWorker.register("sw.js").catch(() => {}); }
