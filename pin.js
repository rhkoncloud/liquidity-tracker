"use strict";
/* App PIN lock — deters casual access on a shared/unlocked device.
   Stores only a salted SHA-256 hash in localStorage (never the PIN itself, never synced).
   Note: this is a convenience lock, not encryption — real protection is your Microsoft sign-in. */
(function(){
const K = "liq_pin";
const getPin = () => { try{ return JSON.parse(localStorage.getItem(K) || "null"); }catch(_){ return null; } };
async function hash(salt, pin){
  const d = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(salt + ":" + pin));
  return Array.prototype.map.call(new Uint8Array(d), b => b.toString(16).padStart(2, "0")).join("");
}
let tries = 0;

async function tryUnlock(){
  const rec = getPin(); if(!rec) return;
  const v = $("lockPin").value.trim();
  if(!/^\d{4,6}$/.test(v)){ $("lockMsg").textContent = "Enter your 4–6 digit PIN."; return; }
  if(await hash(rec.salt, v) === rec.hash){
    sessionStorage.setItem("liq_unlocked", "1");
    $("lock").classList.add("hide"); $("lockPin").value = ""; $("lockMsg").textContent = ""; tries = 0;
  } else {
    tries++; $("lockPin").value = "";
    $("lockMsg").textContent = "Wrong PIN." + (tries >= 3 ? " Try again in a moment…" : "");
    if(tries >= 3){ $("lockBtn").disabled = true; setTimeout(() => { $("lockBtn").disabled = false; }, 3000 * (tries - 2)); }
  }
}
function showLock(){
  if(!getPin()) return;
  sessionStorage.removeItem("liq_unlocked");
  $("lock").classList.remove("hide");
  setTimeout(() => $("lockPin").focus(), 60);
}

/* Re-lock after 5 minutes in the background */
let hiddenAt = 0;
document.addEventListener("visibilitychange", () => {
  if(document.hidden) hiddenAt = Date.now();
  else if(getPin() && hiddenAt && Date.now() - hiddenAt > 5*60*1000) showLock();
});

function openPinSheet(){
  const rec = getPin();
  $("pinCurF").style.display = rec ? "" : "none";
  $("pinCur").value = ""; $("pinNew").value = ""; $("pinNew2").value = "";
  $("pinMsg").textContent = rec ? "Leave the new PIN empty to remove the lock." : "";
  $("pinSheet").classList.add("open");
}
async function savePin(){
  const rec = getPin();
  if(rec && await hash(rec.salt, $("pinCur").value.trim()) !== rec.hash){ $("pinMsg").textContent = "Current PIN is wrong."; return; }
  const n = $("pinNew").value.trim(), n2 = $("pinNew2").value.trim();
  if(n === "" && rec){ localStorage.removeItem(K); $("pinBtn").textContent = "Set app PIN"; $("pinSheet").classList.remove("open"); return; }
  if(!/^\d{4,6}$/.test(n)){ $("pinMsg").textContent = "PIN must be 4–6 digits."; return; }
  if(n !== n2){ $("pinMsg").textContent = "PINs don't match."; return; }
  const salt = (window.crypto && crypto.randomUUID) ? crypto.randomUUID() : String(Date.now()) + Math.random();
  localStorage.setItem(K, JSON.stringify({ salt, hash: await hash(salt, n) }));
  sessionStorage.setItem("liq_unlocked", "1");
  $("pinBtn").textContent = "Change / remove PIN";
  $("pinSheet").classList.remove("open");
}

window.addEventListener("DOMContentLoaded", () => {
  $("lockBtn").onclick = tryUnlock;
  $("lockPin").addEventListener("keydown", e => { if(e.key === "Enter") tryUnlock(); });
  $("pinBtn").onclick = () => { $("set2").classList.remove("open"); openPinSheet(); };
  $("pinCancel").onclick = () => $("pinSheet").classList.remove("open");
  $("pinSave").onclick = savePin;
  if(getPin()){
    $("pinBtn").textContent = "Change / remove PIN";
    if(sessionStorage.getItem("liq_unlocked") !== "1") showLock();
  }
});
})();
