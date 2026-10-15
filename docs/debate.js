// Debate-night vote. The page sets window.PHASE to "open" or "final" before loading this.
// Backend: apps-script/Debate.gs on its own sheet (URL in debate-config.js). One vote per phase per device.
const API_URL = window.DEBATE_API;
const PHASE = window.PHASE;
const KEY = "staxi-debate-" + PHASE;
const LABEL = { for: "For", against: "Against", undecided: "Undecided" };

const $ = s => document.querySelector(s);
const $$ = s => [...document.querySelectorAll(s)];
const store = {
  get(k) { try { return JSON.parse(localStorage.getItem(k)); } catch { return null; } },
  set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch {} }
};
function show(id) { $$(".screen").forEach(s => s.classList.toggle("on", s.id === id)); }

async function post(payload) {
  const r = await fetch(API_URL, { method: "POST", body: JSON.stringify(payload) });
  return r.json();
}
async function status() {
  try {
    const r = await fetch(API_URL + "?debate=1");
    const d = await r.json();
    return d && d[PHASE] ? d[PHASE] : null;
  } catch { return null; }
}

// Testing: add ?reset to the address to clear this device's vote.
if (new URLSearchParams(location.search).has("reset")) { try { localStorage.removeItem(KEY); localStorage.removeItem(KEY + "-follow"); } catch {} }

let polling = null;
function closed() {
  show("s-closed");
  if (polling) return;
  polling = setInterval(async () => {
    const s = await status();
    if (s && s.accepting) { clearInterval(polling); polling = null; show("s-vote"); }
  }, 8000);
}

function done(vote) {
  $("#you").textContent = LABEL[vote] || "";
  show("s-done");
  if (store.get(KEY + "-follow")) { const f = $("#follow"); if (f) f.hidden = true; }
}

$$("[data-vote]").forEach(b => b.addEventListener("click", async () => {
  const vote = b.dataset.vote;
  $$("[data-vote]").forEach(x => { x.disabled = true; x.classList.toggle("picked", x === b); });
  $("#vote-err").hidden = true;
  try {
    const d = await post({ type: "debate", phase: PHASE, vote });
    if (d && d.ok) { store.set(KEY, { vote }); done(vote); return; }
    if (d && d.error === "closed") { closed(); return; }
    throw new Error((d && d.error) || "failed");
  } catch {
    $("#vote-err").hidden = false;
  } finally {
    $$("[data-vote]").forEach(x => { x.disabled = false; });
  }
}));

const form = $("#follow-form");
if (form) form.addEventListener("submit", async e => {
  e.preventDefault();
  const contact = $("#contact").value.trim(), suggestion = $("#suggestion").value.trim();
  const err = $("#follow-err");
  if (!contact && !suggestion) { err.textContent = "Add contact details or a suggestion first."; err.hidden = false; return; }
  err.hidden = true;
  const btn = $("#follow-send");
  btn.disabled = true; btn.textContent = "Sending…";
  try {
    const d = await post({ type: "followup", contact, suggestion });
    if (!d || !d.ok) throw new Error();
    store.set(KEY + "-follow", true);
    form.hidden = true;
    $("#follow-thanks").hidden = false;
  } catch {
    err.textContent = "That didn't send. Check your signal and try again.";
    err.hidden = false;
  } finally {
    btn.disabled = false; btn.textContent = "Send";
  }
});

(async () => {
  const prior = store.get(KEY);
  if (prior && prior.vote) { done(prior.vote); return; }
  // The buttons show straight away; switch to "not open yet" if the sheet says so.
  const s = await status();
  if (s && !s.accepting && !store.get(KEY)) closed();
})();
