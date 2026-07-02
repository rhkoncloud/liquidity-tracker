"use strict";
/* Dashboard tab — spend analysis charts built from the SAME OneDrive workbook the tracker
   uses ("Monthly Spend by Category" + "Liquidity Calendar" sheets), fetched live via
   Microsoft Graph. Nothing is stored in this repo or cached to disk. */
(function(){
let charts = {}, dashLoaded = false, dash = null;
const css = v => getComputedStyle(document.documentElement).getPropertyValue(v).trim();
const fmtL = v => "₹" + (Math.abs(v) >= 100000 ? (v/100000).toFixed(1) + "L" : Math.round(v/1000) + "k");

async function sheetVals(name){
  const r = await graph(`worksheets('${encodeURIComponent(name)}')/usedRange?$select=values`);
  return (r && r.values) || [];
}

function parseData(msc, lc){
  const hi = msc.findIndex(r => r && r[0] === "Category");
  if(hi < 0) throw new Error("'Monthly Spend by Category' sheet layout not recognised");
  const months = msc[hi].slice(1, 13).map(String);
  const cats = {}; let ti = -1;
  for(let i = hi + 1; i < msc.length; i++){
    const n = msc[i] && msc[i][0];
    if(n == null || n === "") continue;
    if(String(n).indexOf("TOTAL") === 0){ ti = i; break; }
    cats[String(n)] = msc[i].slice(1, 13).map(v => Number(v) || 0);
  }
  const total = ti >= 0 ? msc[ti].slice(1, 13).map(v => Number(v) || 0) : null;
  const baseRow = msc.find(r => r && String(r[0]).indexOf("Base take-home") === 0);
  const base = baseRow ? (Number(baseRow[1]) || 260000) : 260000;
  const find = p => lc.find(r => r && String(r[0]).indexOf(p) === 0);
  const ess = find("ESSENTIAL RECURRING"), tot = find("TOTAL incl. SIP");
  return { months, cats, total, base,
    essential: ess ? ess.slice(1, 13).map(v => Number(v) || 0) : null,
    totalSip: tot ? tot.slice(1, 13).map(v => Number(v) || 0) : null };
}

function outliers(d){
  const out = [];
  Object.keys(d.cats).forEach(c => {
    const vals = d.cats[c], nz = vals.filter(v => v > 0);
    if(nz.length < 3) return;
    const m = nz.reduce((a,b) => a+b, 0) / nz.length;
    const sd = Math.sqrt(nz.reduce((a,b) => a + (b-m)*(b-m), 0) / nz.length);
    vals.forEach((v, i) => { if(v > m + 1.5*sd && v - m > 8000) out.push({ cat: c, mi: i, value: Math.round(v), avg: Math.round(m), x: +(v/m).toFixed(1) }); });
  });
  return out.sort((a,b) => b.value - a.value);
}

function mk(id, cfg){ if(charts[id]) charts[id].destroy(); charts[id] = new Chart($(id), cfg); }

function render(d){
  const blue = css("--blue") || "#2E5496", green = css("--green"), red = css("--red"),
        muted = css("--muted"), line = css("--line");
  Chart.defaults.color = muted; Chart.defaults.borderColor = line; Chart.defaults.font.size = 10;
  const yr = d.total.reduce((a,b) => a+b, 0), over = d.total.filter(v => v > d.base).length;
  $("dTot").textContent = fmtL(yr); $("dAvg").textContent = fmtL(yr/12); $("dOver").textContent = over + " / 12";
  const tick = { callback: v => fmtL(v) };
  const tip = { callbacks: { label: c => c.dataset.label + ": " + fmt(c.raw) } };

  mk("dcTrend", { data: { labels: d.months, datasets: [
    { type: "bar", label: "Total spend", data: d.total, borderRadius: 4,
      backgroundColor: d.total.map(v => v > d.base ? red : blue) },
    { type: "line", label: "Base " + fmtL(d.base), data: d.months.map(() => d.base),
      borderColor: green, pointRadius: 0, borderWidth: 2 }
  ]}, options: { maintainAspectRatio: false, scales: { y: { ticks: tick } }, plugins: { tooltip: tip } } });

  const names = Object.keys(d.cats), tots = names.map(n => d.cats[n].reduce((a,b) => a+b, 0));
  const ord = tots.map((t,i) => i).sort((a,b) => tots[b] - tots[a]);
  const oN = ord.map(i => names[i]), oT = ord.map(i => Math.round(tots[i]));
  mk("dcCat", { type: "bar", data: { labels: oN, datasets: [{ data: oT, backgroundColor: blue, borderRadius: 4 }] },
    options: { indexAxis: "y", maintainAspectRatio: false,
      plugins: { legend: { display: false }, tooltip: { callbacks: { label: c => fmt(c.raw) + " / yr" } } },
      scales: { x: { ticks: tick } },
      onClick: (e, els) => { if(els.length){ $("dsel").value = oN[els[0].index]; drill(d, oN[els[0].index]); } } } });

  const sel = $("dsel"); sel.innerHTML = "";
  oN.forEach(n => { const o = document.createElement("option"); o.textContent = n; sel.appendChild(o); });
  sel.onchange = () => drill(d, sel.value);
  drill(d, oN[0]);

  const ol = $("dOut"); ol.innerHTML = "";
  const outs = outliers(d);
  outs.forEach(o => {
    const row = document.createElement("div"); row.className = "row";
    const name = document.createElement("div"); name.className = "name";
    const n = document.createElement("div"); n.className = "n"; n.textContent = o.cat;
    const c = document.createElement("div"); c.className = "c"; c.textContent = d.months[o.mi] + " · category avg " + fmt(o.avg);
    name.appendChild(n); name.appendChild(c);
    const amt = document.createElement("div"); amt.className = "amt"; amt.textContent = fmt(o.value);
    const x = document.createElement("div");
    x.style.cssText = "font-size:12px; font-weight:700; min-width:38px; text-align:right; color:" + (o.x >= 3 ? red : "#b8860b");
    x.textContent = o.x + "×";
    row.appendChild(name); row.appendChild(amt); row.appendChild(x); ol.appendChild(row);
  });
  if(!outs.length){ const e = document.createElement("div"); e.className = "status"; e.textContent = "No outliers detected."; ol.appendChild(e); }

  if(d.essential && d.totalSip){
    mk("dcLiq", { data: { labels: d.months, datasets: [
      { type: "bar", label: "Essential recurring", data: d.essential, backgroundColor: blue, borderRadius: 4 },
      { type: "bar", label: "+ planned SIP", data: d.totalSip.map((v,i) => Math.max(0, v - d.essential[i])), backgroundColor: muted, borderRadius: 4 },
      { type: "line", label: "Base " + fmtL(d.base), data: d.months.map(() => d.base), borderColor: green, pointRadius: 0, borderWidth: 2 }
    ]}, options: { maintainAspectRatio: false, plugins: { tooltip: tip },
      scales: { x: { stacked: true }, y: { stacked: true, ticks: tick } } } });
  }
}

function drill(d, name){
  const vals = d.cats[name], nz = vals.filter(v => v > 0);
  const avg = nz.length ? nz.reduce((a,b) => a+b, 0) / nz.length : 0;
  const sd = nz.length ? Math.sqrt(nz.reduce((a,b) => a + (b-avg)*(b-avg), 0) / nz.length) : 0;
  mk("dcDrill", { data: { labels: d.months, datasets: [
    { type: "bar", label: name, data: vals, borderRadius: 4,
      backgroundColor: vals.map(v => v > avg + 1.5*sd ? css("--red") : css("--green")) },
    { type: "line", label: "avg (active months)", data: d.months.map(() => avg),
      borderColor: "#b8860b", borderDash: [5,4], pointRadius: 0 }
  ]}, options: { maintainAspectRatio: false, scales: { y: { ticks: { callback: v => fmtL(v) } } },
    plugins: { tooltip: { callbacks: { label: c => c.dataset.label + ": " + fmt(c.raw) } } } } });
}

async function loadDash(force){
  if(dashLoaded && !force) return;
  if(!account){ $("dMsg").textContent = "Sign in to load the dashboard."; return; }
  if(typeof Chart === "undefined"){ $("dMsg").textContent = "Charts library didn't load — check your connection and tap ↻."; return; }
  $("dMsg").textContent = "Loading from OneDrive…";
  try{
    const both = await Promise.all([sheetVals("Monthly Spend by Category"), sheetVals("Liquidity Calendar")]);
    dash = parseData(both[0], both[1]);
    render(dash); dashLoaded = true;
    $("dMsg").textContent = "FY 2025-26 actuals · synced " + new Date().toLocaleTimeString();
  }catch(e){ $("dMsg").textContent = "Couldn't load dashboard: " + e.message; }
}

function showTab(t){
  $("tabTrack").classList.toggle("on", t === "track");
  $("tabDash").classList.toggle("on", t === "dash");
  $("app").classList.toggle("hide", t !== "track" || !account);
  $("fab").classList.toggle("hide", t !== "track" || !account);
  $("dash").classList.toggle("hide", t !== "dash");
  if(t === "dash") loadDash(false);
}

window.addEventListener("DOMContentLoaded", () => {
  $("tabTrack").onclick = () => showTab("track");
  $("tabDash").onclick = () => showTab("dash");
  $("dRefresh").onclick = () => loadDash(true);
});
})();
